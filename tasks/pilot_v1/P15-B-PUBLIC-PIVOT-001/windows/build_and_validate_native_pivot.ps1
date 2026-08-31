[CmdletBinding()]
param(
    [string]$InputWorkbook = (Join-Path (Split-Path -Parent $PSScriptRoot) "data\input_files\starting_workbook.xlsx"),
    [Parameter(Mandatory = $true)]
    [string]$OutputWorkbook,
    [string]$ReceiptPath = "",
    [double]$ExpectedQ2Participants = 132,
    [double]$ExpectedQ2Spend = 396000,
    [string]$FocusRegion = "North",
    [string]$FocusProgram = "Outreach",
    [double]$ExpectedFocusParticipants = 50,
    [int]$RefreshTimeoutSeconds = 120
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$xlSrcRange = 1
$xlYes = 1
$xlDatabase = 1
$xlRowField = 1
$xlColumnField = 2
$xlPageField = 3
$xlSum = -4157
$xlColumnClustered = 51
$xlLocationAsObject = 2
$xlOpenXMLWorkbook = 51
$xlPivotTableVersion15 = 5
$xlMissingItemsNone = 0
$xlDone = 0
$msoAutomationSecurityForceDisable = 3

function Assert-Equal {
    param($Actual, $Expected, [string]$Message)
    if ($Actual -ne $Expected) {
        throw "$Message; expected=[$Expected] actual=[$Actual]"
    }
}

function Assert-Near {
    param([double]$Actual, [double]$Expected, [double]$Tolerance, [string]$Message)
    if ([Math]::Abs($Actual - $Expected) -gt $Tolerance) {
        throw "$Message; expected=[$Expected] actual=[$Actual] tolerance=[$Tolerance]"
    }
}

function Release-ComObject {
    param($Object)
    if ($null -ne $Object) {
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object) } catch { }
    }
}

function Get-WorksheetObject {
    param($Workbook, [string]$Name)
    $collection = $null
    try {
        $collection = $Workbook.Worksheets
        return $collection.Item($Name)
    } finally {
        Release-ComObject $collection
    }
}

function Get-RangeValue2 {
    param($Sheet, [string]$Address)
    $range = $null
    try {
        $range = $Sheet.Range($Address)
        return $range.Value2
    } finally {
        Release-ComObject $range
    }
}

function Get-RangeFormula {
    param($Sheet, [string]$Address)
    $range = $null
    try {
        $range = $Sheet.Range($Address)
        return $range.Formula
    } finally {
        Release-ComObject $range
    }
}

function Set-RangeFormula {
    param($Sheet, [string]$Address, [string]$Formula)
    $range = $null
    try {
        $range = $Sheet.Range($Address)
        $range.Formula = $Formula
    } finally {
        Release-ComObject $range
    }
}

function Write-Utf8Json {
    param([string]$Path, $Value, [int]$Depth = 12)
    [IO.File]::WriteAllText(
        $Path,
        ($Value | ConvertTo-Json -Depth $Depth) + [Environment]::NewLine,
        [Text.UTF8Encoding]::new($false)
    )
}

function Invoke-RefreshCalculationBarrier {
    param($Excel, $Workbook, [string]$Phase, [int]$TimeoutSeconds)

    try {
        $Workbook.RefreshAll()
    } catch {
        throw "$Phase RefreshAll failed: $($_.Exception.Message)"
    }
    try {
        $Excel.CalculateUntilAsyncQueriesDone()
    } catch {
        throw "$Phase async-query barrier failed: $($_.Exception.Message)"
    }
    try {
        $Excel.CalculateFullRebuild()
    } catch {
        throw "$Phase CalculateFullRebuild failed: $($_.Exception.Message)"
    }

    $deadline = [DateTime]::UtcNow.AddSeconds($TimeoutSeconds)
    while ([DateTime]::UtcNow -lt $deadline) {
        if ([int]$Excel.CalculationState -eq $xlDone) {
            try {
                $Excel.CalculateUntilAsyncQueriesDone()
            } catch {
                throw "$Phase final async-query barrier failed: $($_.Exception.Message)"
            }
            if ([int]$Excel.CalculationState -eq $xlDone) {
                return
            }
        }
        Start-Sleep -Milliseconds 250
    }
    throw "$Phase timed out after $TimeoutSeconds seconds waiting for Excel calculation state DONE."
}

function Promote-WorkbookAtomically {
    param(
        [string]$CandidatePath,
        [string]$DestinationPath,
        [bool]$DestinationExisted,
        [string]$BackupPath
    )
    if ($DestinationExisted) {
        [IO.File]::Replace($CandidatePath, $DestinationPath, $BackupPath, $true)
    } else {
        [IO.File]::Move($CandidatePath, $DestinationPath)
    }
}

function Promote-ReceiptAtomically {
    param(
        [string]$CandidatePath,
        [string]$DestinationPath,
        [bool]$DestinationExisted,
        [string]$BackupPath
    )
    if ($DestinationExisted) {
        [IO.File]::Replace($CandidatePath, $DestinationPath, $BackupPath, $true)
    } else {
        [IO.File]::Move($CandidatePath, $DestinationPath)
    }
}

$inputFull = [IO.Path]::GetFullPath($InputWorkbook)
$outputFull = [IO.Path]::GetFullPath($OutputWorkbook)
if (-not (Test-Path -LiteralPath $inputFull -PathType Leaf)) {
    throw "Input workbook not found: $inputFull"
}
if ($inputFull -eq $outputFull) {
    throw "OutputWorkbook must be a new candidate path; the source starter cannot be overwritten."
}
if ([IO.Path]::GetExtension($outputFull) -ne ".xlsx") {
    throw "OutputWorkbook must use the .xlsx extension: $outputFull"
}

$outputDirectory = Split-Path -Parent $outputFull
if (-not (Test-Path -LiteralPath $outputDirectory)) {
    [void](New-Item -ItemType Directory -Path $outputDirectory)
}
if ([string]::IsNullOrWhiteSpace($ReceiptPath)) {
    $ReceiptPath = "$outputFull.windows-com-validation.json"
}
$receiptFull = [IO.Path]::GetFullPath($ReceiptPath)
if ($receiptFull -eq $outputFull) {
    throw "ReceiptPath must not be the OutputWorkbook path."
}
$receiptDirectory = Split-Path -Parent $receiptFull
if (-not (Test-Path -LiteralPath $receiptDirectory)) {
    [void](New-Item -ItemType Directory -Path $receiptDirectory)
}

$token = [Guid]::NewGuid().ToString("N")
$outputStem = [IO.Path]::GetFileNameWithoutExtension($outputFull)
$stagingFull = Join-Path $outputDirectory (".{0}.native-build.{1}.staging.xlsx" -f $outputStem, $token)
$receiptStageFull = Join-Path $receiptDirectory (".{0}.{1}.staging.json" -f [IO.Path]::GetFileName($receiptFull), $token)
if ((Test-Path -LiteralPath $outputFull) -and -not (Test-Path -LiteralPath $outputFull -PathType Leaf)) {
    throw "OutputWorkbook exists but is not a file: $outputFull"
}
if ((Test-Path -LiteralPath $receiptFull) -and -not (Test-Path -LiteralPath $receiptFull -PathType Leaf)) {
    throw "ReceiptPath exists but is not a file: $receiptFull"
}
$outputExisted = Test-Path -LiteralPath $outputFull -PathType Leaf
$receiptExisted = Test-Path -LiteralPath $receiptFull -PathType Leaf
$timestamp = (Get-Date).ToUniversalTime().ToString("yyyyMMddTHHmmssfffZ")
$outputBackupFull = $null
$receiptBackupFull = $null
if ($outputExisted) {
    $outputBackupFull = Join-Path $outputDirectory ("{0}.pre-native-promotion.{1}.{2}.backup.xlsx" -f $outputStem, $timestamp, $token)
}
if ($receiptExisted) {
    $receiptBackupFull = Join-Path $receiptDirectory ("{0}.pre-native-promotion.{1}.{2}.backup.json" -f [IO.Path]::GetFileNameWithoutExtension($receiptFull), $timestamp, $token)
}

$excel = $null
$workbooks = $null
$workbook = $null
$worksheets = $null
$dataSheet = $null
$pivotSheet = $null
$kpiSheet = $null
$listObjects = $null
$sourceRange = $null
$sourceTable = $null
$pivotCaches = $null
$cache = $null
$pivotTable = $null
$pivotFields = $null
$pivotTables = $null
$dataFields = $null
$eventField = $null
$regionField = $null
$programField = $null
$quarterField = $null
$participantsSourceField = $null
$spendSourceField = $null
$participantsField = $null
$spendField = $null
$charts = $null
$chartSheet = $null
$chartTitle = $null
$locatedChart = $null
$chartObjects = $null
$chartObject = $null
$embeddedChart = $null
$pivotLayout = $null
$layoutPivotTable = $null
$excelVersion = $null
$receipt = $null

try {
    # Build only on a unique same-directory candidate. The caller's output is not
    # touched until the COM build, reopen/readback, assertions, and staged receipt succeed.
    Copy-Item -LiteralPath $inputFull -Destination $stagingFull

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.AskToUpdateLinks = $false
        $excel.AutomationSecurity = $msoAutomationSecurityForceDisable
        Assert-Equal ([int]$excel.AutomationSecurity) $msoAutomationSecurityForceDisable "Excel AutomationSecurity"
        $excelVersion = [string]$excel.Version

        $workbooks = $excel.Workbooks
        $workbook = $workbooks.Open($stagingFull, 0, $false)
        Release-ComObject $workbooks
        $workbooks = $null

        $dataSheet = Get-WorksheetObject $workbook "Program_Data"
        Assert-Equal ([string](Get-RangeValue2 $dataSheet "A3")) "Event ID" "Program_Data header 1"
        Assert-Equal ([string](Get-RangeValue2 $dataSheet "B3")) "Region" "Program_Data header 2"
        Assert-Equal ([string](Get-RangeValue2 $dataSheet "C3")) "Program" "Program_Data header 3"
        Assert-Equal ([string](Get-RangeValue2 $dataSheet "D3")) "Quarter" "Program_Data header 4"
        Assert-Equal ([string](Get-RangeValue2 $dataSheet "E3")) "Participants" "Program_Data header 5"
        Assert-Equal ([string](Get-RangeValue2 $dataSheet "F3")) "Spend" "Program_Data header 6"

        $listObjects = $dataSheet.ListObjects
        while ($listObjects.Count -gt 0) {
            $oldTable = $listObjects.Item(1)
            $oldTable.Unlist()
            Release-ComObject $oldTable
        }
        $sourceRange = $dataSheet.Range("A3:F11")
        $sourceTable = $listObjects.Add($xlSrcRange, $sourceRange, $null, $xlYes)
        Release-ComObject $sourceRange
        $sourceRange = $null
        $sourceTable.Name = "ProgramEventsTable"
        $sourceTable.TableStyle = "TableStyleMedium2"
        $sourceTableRange = $sourceTable.Range
        Assert-Equal ([string]$sourceTableRange.Address($false, $false)) "A3:F11" "Excel Table source range"
        Release-ComObject $sourceTableRange

        $worksheets = $workbook.Worksheets
        for ($index = $worksheets.Count; $index -ge 1; $index--) {
            $existingSheet = $worksheets.Item($index)
            if ([string]$existingSheet.Name -eq "Pivot_Report") {
                $existingSheet.Delete()
            }
            Release-ComObject $existingSheet
        }
        $pivotSheet = $worksheets.Add()
        Release-ComObject $worksheets
        $worksheets = $null
        $pivotSheet.Name = "Pivot_Report"
        $titleRange = $pivotSheet.Range("A1:H1")
        $titleRange.Merge()
        Release-ComObject $titleRange
        $titleCell = $pivotSheet.Range("A1")
        $titleCell.Value2 = "2024Q2 Program Delivery Pivot"
        Release-ComObject $titleCell

        $pivotCaches = $workbook.PivotCaches()
        $cache = $pivotCaches.Create($xlDatabase, "ProgramEventsTable", $xlPivotTableVersion15)
        Release-ComObject $pivotCaches
        $pivotCaches = $null
        $cache.EnableRefresh = $true
        $cache.RefreshOnFileOpen = $true
        $cache.MissingItemsLimit = $xlMissingItemsNone
        try {
            $cache.BackgroundQuery = $false
        } catch {
            throw "Unable to disable PivotCache.BackgroundQuery: $($_.Exception.Message)"
        }
        Assert-Equal ([bool]$cache.BackgroundQuery) $false "PivotCache BackgroundQuery"

        $pivotDestination = $pivotSheet.Range("A3")
        $pivotTable = $cache.CreatePivotTable($pivotDestination, "ProgramDeliveryPivot")
        Release-ComObject $pivotDestination
        $pivotTable.ManualUpdate = $true
        $pivotFields = $pivotTable.PivotFields()
        $regionField = $pivotFields.Item("Region")
        $regionField.Orientation = $xlRowField
        $regionField.Position = 1
        $programField = $pivotFields.Item("Program")
        $programField.Orientation = $xlColumnField
        $programField.Position = 1
        $quarterField = $pivotFields.Item("Quarter")
        $quarterField.Orientation = $xlPageField
        $quarterField.Position = 1
        $quarterField.ClearAllFilters()
        $quarterField.CurrentPage = "2024Q2"
        $participantsSourceField = $pivotFields.Item("Participants")
        $participantsField = $pivotTable.AddDataField($participantsSourceField, "Sum of Participants", $xlSum)
        $spendSourceField = $pivotFields.Item("Spend")
        $spendField = $pivotTable.AddDataField($spendSourceField, "Sum of Spend", $xlSum)
        $participantsField.NumberFormat = "0"
        $spendField.NumberFormat = '$#,##0'
        $pivotTable.ManualUpdate = $false
        [void]$cache.Refresh()
        [void]$pivotTable.RefreshTable()

        $kpiSheet = Get-WorksheetObject $workbook "KPI_Summary"
        Set-RangeFormula $kpiSheet "B4" '=GETPIVOTDATA("Sum of Participants",Pivot_Report!$A$3)'
        Set-RangeFormula $kpiSheet "B5" '=GETPIVOTDATA("Sum of Spend",Pivot_Report!$A$3)'
        $focusFormula = '=GETPIVOTDATA("Sum of Participants",Pivot_Report!$A$3,"Region","{0}","Program","{1}")' -f $FocusRegion, $FocusProgram
        Set-RangeFormula $kpiSheet "B6" $focusFormula

        $pivotSheet.Activate()
        $pivotRange = $pivotTable.TableRange2
        $pivotRange.Select()
        Release-ComObject $pivotRange
        $charts = $workbook.Charts
        $chartSheet = $charts.Add()
        Release-ComObject $charts
        $charts = $null
        $chartSheet.ChartType = $xlColumnClustered
        $chartSheet.HasTitle = $true
        $chartTitle = $chartSheet.ChartTitle
        $chartTitle.Text = "2024Q2 Program Delivery by Region and Program"
        Release-ComObject $chartTitle
        $chartTitle = $null
        $locatedChart = $chartSheet.Location($xlLocationAsObject, "Pivot_Report")
        Release-ComObject $chartSheet
        $chartSheet = $null
        Release-ComObject $locatedChart
        $locatedChart = $null

        $chartObjects = $pivotSheet.ChartObjects()
        $chartObject = $chartObjects.Item($chartObjects.Count)
        $embeddedChart = $chartObject.Chart
        $pivotLayout = $embeddedChart.PivotLayout
        if ($null -eq $pivotLayout) {
            throw "Excel created a regular chart; PivotLayout is missing."
        }
        $layoutPivotTable = $pivotLayout.PivotTable
        Assert-Equal ([string]$layoutPivotTable.Name) "ProgramDeliveryPivot" "PivotChart PivotLayout binding"
        Assert-Equal ([int]$embeddedChart.ChartType) $xlColumnClustered "PivotChart clustered-column type"

        Invoke-RefreshCalculationBarrier $excel $workbook "native-build" $RefreshTimeoutSeconds
        [void]$pivotTable.RefreshTable()
        Invoke-RefreshCalculationBarrier $excel $workbook "native-build-post-pivot-refresh" $RefreshTimeoutSeconds
        $workbook.SaveAs($stagingFull, $xlOpenXMLWorkbook)

        Release-ComObject $layoutPivotTable; $layoutPivotTable = $null
        Release-ComObject $pivotLayout; $pivotLayout = $null
        Release-ComObject $embeddedChart; $embeddedChart = $null
        Release-ComObject $chartObject; $chartObject = $null
        Release-ComObject $chartObjects; $chartObjects = $null
        Release-ComObject $participantsField; $participantsField = $null
        Release-ComObject $spendField; $spendField = $null
        Release-ComObject $participantsSourceField; $participantsSourceField = $null
        Release-ComObject $spendSourceField; $spendSourceField = $null
        Release-ComObject $regionField; $regionField = $null
        Release-ComObject $programField; $programField = $null
        Release-ComObject $quarterField; $quarterField = $null
        Release-ComObject $pivotFields; $pivotFields = $null
        Release-ComObject $pivotTable; $pivotTable = $null
        Release-ComObject $cache; $cache = $null
        Release-ComObject $sourceTable; $sourceTable = $null
        Release-ComObject $listObjects; $listObjects = $null
        Release-ComObject $kpiSheet; $kpiSheet = $null
        Release-ComObject $pivotSheet; $pivotSheet = $null
        Release-ComObject $dataSheet; $dataSheet = $null
        $workbook.Close($true)
        Release-ComObject $workbook
        $workbook = $null

        # Reopen the staged candidate and read native objects back from Excel before promotion.
        $workbooks = $excel.Workbooks
        $workbook = $workbooks.Open($stagingFull, 0, $false)
        Release-ComObject $workbooks
        $workbooks = $null

        $dataSheet = Get-WorksheetObject $workbook "Program_Data"
        $listObjects = $dataSheet.ListObjects
        $sourceTable = $listObjects.Item("ProgramEventsTable")
        $sourceTableRange = $sourceTable.Range
        Assert-Equal ([string]$sourceTableRange.Address($false, $false)) "A3:F11" "Reopened table range"
        Release-ComObject $sourceTableRange

        $pivotSheet = Get-WorksheetObject $workbook "Pivot_Report"
        $pivotTables = $pivotSheet.PivotTables()
        $pivotTable = $pivotTables.Item("ProgramDeliveryPivot")
        Release-ComObject $pivotTables
        $pivotTables = $null
        $cache = $pivotTable.PivotCache()
        Assert-Equal ([string]$cache.SourceData) "ProgramEventsTable" "Reopened PivotCache source"
        $pivotFields = $pivotTable.PivotFields()
        $eventField = $pivotFields.Item("Event ID")
        $regionField = $pivotFields.Item("Region")
        $programField = $pivotFields.Item("Program")
        $quarterField = $pivotFields.Item("Quarter")
        $participantsSourceField = $pivotFields.Item("Participants")
        $spendSourceField = $pivotFields.Item("Spend")
        Assert-Equal ([string]$eventField.SourceName) "Event ID" "Event ID source-field identity"
        Assert-Equal ([string]$regionField.SourceName) "Region" "Region source-field identity"
        Assert-Equal ([string]$programField.SourceName) "Program" "Program source-field identity"
        Assert-Equal ([string]$quarterField.SourceName) "Quarter" "Quarter source-field identity"
        Assert-Equal ([string]$participantsSourceField.SourceName) "Participants" "Participants source-field identity"
        Assert-Equal ([string]$spendSourceField.SourceName) "Spend" "Spend source-field identity"
        Assert-Equal ([int]$eventField.Index) 1 "Event ID one-based field index"
        Assert-Equal ([int]$regionField.Index) 2 "Region one-based field index"
        Assert-Equal ([int]$programField.Index) 3 "Program one-based field index"
        Assert-Equal ([int]$quarterField.Index) 4 "Quarter one-based field index"
        Assert-Equal ([int]$participantsSourceField.Index) 5 "Participants one-based field index"
        Assert-Equal ([int]$spendSourceField.Index) 6 "Spend one-based field index"
        Assert-Equal ([string]$regionField.Orientation) ([string]$xlRowField) "Region orientation"
        Assert-Equal ([string]$programField.Orientation) ([string]$xlColumnField) "Program orientation"
        Assert-Equal ([string]$quarterField.Orientation) ([string]$xlPageField) "Quarter orientation"
        Assert-Equal ([string]$quarterField.CurrentPage) "2024Q2" "Quarter selected item"
        $dataFields = $pivotTable.DataFields
        Assert-Equal ([int]$dataFields.Count) 2 "Data-field count"
        $participantsField = $dataFields.Item("Sum of Participants")
        $spendField = $dataFields.Item("Sum of Spend")
        Assert-Equal ([int]$participantsField.Function) $xlSum "Participants aggregation"
        Assert-Equal ([int]$spendField.Function) $xlSum "Spend aggregation"
        Assert-Equal ([bool]$cache.RefreshOnFileOpen) $true "RefreshOnFileOpen"
        Assert-Equal ([bool]$cache.EnableRefresh) $true "EnableRefresh"
        Assert-Equal ([bool]$cache.BackgroundQuery) $false "Reopened PivotCache BackgroundQuery"

        [void]$cache.Refresh()
        [void]$pivotTable.RefreshTable()
        Invoke-RefreshCalculationBarrier $excel $workbook "native-reopen-readback" $RefreshTimeoutSeconds

        $kpiSheet = Get-WorksheetObject $workbook "KPI_Summary"
        Assert-Near ([double](Get-RangeValue2 $kpiSheet "B4")) $ExpectedQ2Participants 0.001 "Q2 participant oracle"
        Assert-Near ([double](Get-RangeValue2 $kpiSheet "B5")) $ExpectedQ2Spend 0.001 "Q2 spend oracle"
        Assert-Near ([double](Get-RangeValue2 $kpiSheet "B6")) $ExpectedFocusParticipants 0.001 "$FocusRegion $FocusProgram participant oracle"
        Assert-Equal ([bool]([string](Get-RangeFormula $kpiSheet "B4") -match "GETPIVOTDATA")) $true "KPI B4 native link"
        Assert-Equal ([bool]([string](Get-RangeFormula $kpiSheet "B5") -match "GETPIVOTDATA")) $true "KPI B5 native link"
        Assert-Equal ([bool]([string](Get-RangeFormula $kpiSheet "B6") -match "GETPIVOTDATA")) $true "KPI B6 native link"

        $pivotChartFound = $false
        $worksheets = $workbook.Worksheets
        for ($sheetIndex = 1; $sheetIndex -le $worksheets.Count; $sheetIndex++) {
            $worksheet = $worksheets.Item($sheetIndex)
            $candidateChartObjects = $worksheet.ChartObjects()
            for ($chartIndex = 1; $chartIndex -le $candidateChartObjects.Count; $chartIndex++) {
                $candidateChartObject = $candidateChartObjects.Item($chartIndex)
                $candidateChart = $candidateChartObject.Chart
                $candidateLayout = $candidateChart.PivotLayout
                if ($null -ne $candidateLayout) {
                    $candidateLayoutPivot = $candidateLayout.PivotTable
                    if ([string]$candidateLayoutPivot.Name -eq "ProgramDeliveryPivot") {
                        Assert-Equal ([int]$candidateChart.ChartType) $xlColumnClustered "Reopened PivotChart clustered-column type"
                        $pivotChartFound = $true
                    }
                    Release-ComObject $candidateLayoutPivot
                }
                Release-ComObject $candidateLayout
                Release-ComObject $candidateChart
                Release-ComObject $candidateChartObject
            }
            Release-ComObject $candidateChartObjects
            Release-ComObject $worksheet
        }
        Release-ComObject $worksheets
        $worksheets = $null
        Assert-Equal $pivotChartFound $true "PivotChart relationship"

        $workbook.Save()
        Release-ComObject $participantsField; $participantsField = $null
        Release-ComObject $spendField; $spendField = $null
        Release-ComObject $dataFields; $dataFields = $null
        Release-ComObject $eventField; $eventField = $null
        Release-ComObject $participantsSourceField; $participantsSourceField = $null
        Release-ComObject $spendSourceField; $spendSourceField = $null
        Release-ComObject $regionField; $regionField = $null
        Release-ComObject $programField; $programField = $null
        Release-ComObject $quarterField; $quarterField = $null
        Release-ComObject $pivotFields; $pivotFields = $null
        Release-ComObject $pivotTable; $pivotTable = $null
        Release-ComObject $cache; $cache = $null
        Release-ComObject $sourceTable; $sourceTable = $null
        Release-ComObject $listObjects; $listObjects = $null
        Release-ComObject $kpiSheet; $kpiSheet = $null
        Release-ComObject $pivotSheet; $pivotSheet = $null
        Release-ComObject $dataSheet; $dataSheet = $null
        $workbook.Close($true)
        Release-ComObject $workbook
        $workbook = $null
    }
    finally {
        if ($null -ne $workbook) {
            try { $workbook.Close($false) } catch { }
        }
        if ($null -ne $excel) {
            try { $excel.Quit() } catch { }
        }
        Release-ComObject $layoutPivotTable
        Release-ComObject $pivotLayout
        Release-ComObject $embeddedChart
        Release-ComObject $chartObject
        Release-ComObject $chartObjects
        Release-ComObject $chartSheet
        Release-ComObject $chartTitle
        Release-ComObject $locatedChart
        Release-ComObject $charts
        Release-ComObject $participantsField
        Release-ComObject $spendField
        Release-ComObject $dataFields
        Release-ComObject $eventField
        Release-ComObject $participantsSourceField
        Release-ComObject $spendSourceField
        Release-ComObject $regionField
        Release-ComObject $programField
        Release-ComObject $quarterField
        Release-ComObject $pivotFields
        Release-ComObject $pivotTables
        Release-ComObject $pivotTable
        Release-ComObject $cache
        Release-ComObject $pivotCaches
        Release-ComObject $sourceTable
        Release-ComObject $sourceRange
        Release-ComObject $listObjects
        Release-ComObject $kpiSheet
        Release-ComObject $pivotSheet
        Release-ComObject $dataSheet
        Release-ComObject $worksheets
        Release-ComObject $workbook
        Release-ComObject $workbooks
        Release-ComObject $excel
        $workbook = $null
        $excel = $null
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }

    if (-not (Test-Path -LiteralPath $stagingFull -PathType Leaf)) {
        throw "Validated staging workbook disappeared before promotion: $stagingFull"
    }
    $stagingHash = (Get-FileHash -LiteralPath $stagingFull -Algorithm SHA256).Hash.ToLowerInvariant()
    $receipt = [ordered]@{
        task_id = "P15-B-PUBLIC-PIVOT-001"
        status = "WINDOWS_EXCEL_COM_NATIVE_OBJECTS_VALIDATED"
        excel_version = $excelVersion
        automation_security = $msoAutomationSecurityForceDisable
        output_workbook = $outputFull
        output_sha256 = $stagingHash
        promotion = [ordered]@{
            strategy = $(if ($outputExisted) { "same-directory File.Replace" } else { "same-directory File.Move" })
            replaced_existing_output = [bool]$outputExisted
            previous_output_backup = $outputBackupFull
            receipt_published_before_workbook_promotion = $true
            receipt_rollback_on_workbook_promotion_failure = $true
            hash_basis = "validated staging bytes preserved by same-directory atomic promotion"
        }
        table = [ordered]@{ name = "ProgramEventsTable"; range = "Program_Data!A3:F11" }
        pivot_cache = [ordered]@{ source = "ProgramEventsTable"; refresh_on_open = $true; enable_refresh = $true; background_query = $false }
        pivot_table = [ordered]@{ name = "ProgramDeliveryPivot"; row_field = "Region"; column_field = "Program"; filter = "Quarter=2024Q2"; measures = @("Sum of Participants", "Sum of Spend") }
        oracle = [ordered]@{
            q2_participants = $ExpectedQ2Participants
            q2_spend = $ExpectedQ2Spend
            focus_region = $FocusRegion
            focus_program = $FocusProgram
            focus_participants = $ExpectedFocusParticipants
        }
        pivot_chart = [ordered]@{ pivot_layout_table = "ProgramDeliveryPivot"; chart_type = "clustered column"; chart_type_id = $xlColumnClustered; reopened_readback = $true }
    }

    # Serialize and atomically publish the complete receipt before workbook promotion.
    # If workbook promotion then fails, restore the previous receipt (or absence).
    # File.Replace/Move preserves the validated bytes represented by stagingHash.
    Write-Utf8Json $receiptStageFull $receipt 12
    $receiptPromoted = $false
    $workbookPromoted = $false
    try {
        Promote-ReceiptAtomically $receiptStageFull $receiptFull ([bool]$receiptExisted) $receiptBackupFull
        $receiptPromoted = $true
        Promote-WorkbookAtomically $stagingFull $outputFull ([bool]$outputExisted) $outputBackupFull
        $workbookPromoted = $true
    } catch {
        $promotionError = $_
        if ($receiptPromoted -and -not $workbookPromoted) {
            try {
                if ($receiptExisted) {
                    $failedReceiptFull = Join-Path $receiptDirectory ("{0}.failed-promotion.{1}.{2}.json" -f [IO.Path]::GetFileNameWithoutExtension($receiptFull), $timestamp, $token)
                    [IO.File]::Replace($receiptBackupFull, $receiptFull, $failedReceiptFull, $true)
                } else {
                    [IO.File]::Delete($receiptFull)
                }
            } catch {
                throw "Workbook promotion failed, and receipt rollback also failed. Promotion error: $($promotionError.Exception.Message); receipt rollback error: $($_.Exception.Message)"
            }
        }
        throw $promotionError
    }
    Write-Output ($receipt | ConvertTo-Json -Depth 12)
}
finally {
    # These are unique staging files only; caller-owned output, its backup, and
    # the published receipt are never deleted by failure cleanup.
    if (Test-Path -LiteralPath $stagingFull -PathType Leaf) {
        [IO.File]::Delete($stagingFull)
    }
    if (Test-Path -LiteralPath $receiptStageFull -PathType Leaf) {
        [IO.File]::Delete($receiptStageFull)
    }
}
