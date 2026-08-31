[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$OutputWorkbook,
    [string]$ReceiptPath = "",
    [string]$PythonExe = "python"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Assert-Equal {
    param($Actual, $Expected, [string]$Message)
    if ($Actual -ne $Expected) {
        throw "$Message; expected=[$Expected] actual=[$Actual]"
    }
}

function Write-Utf8Json {
    param([string]$Path, $Value, [int]$Depth = 14)
    [IO.File]::WriteAllText(
        $Path,
        ($Value | ConvertTo-Json -Depth $Depth) + [Environment]::NewLine,
        [Text.UTF8Encoding]::new($false)
    )
}

function Promote-FileAtomically {
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

function Restore-FileAfterFailedPromotion {
    param(
        [string]$DestinationPath,
        [bool]$DestinationExisted,
        [string]$OriginalSha256,
        [string]$BackupPath,
        [string]$FailedCandidatePath
    )
    if ($DestinationExisted) {
        if (-not (Test-Path -LiteralPath $DestinationPath -PathType Leaf)) {
            if (-not (Test-Path -LiteralPath $BackupPath -PathType Leaf)) {
                throw "Destination and rollback backup are both missing: $DestinationPath"
            }
            [IO.File]::Move($BackupPath, $DestinationPath)
        } elseif ((Get-FileHash -LiteralPath $DestinationPath -Algorithm SHA256).Hash.ToLowerInvariant() -cne $OriginalSha256) {
            if (-not (Test-Path -LiteralPath $BackupPath -PathType Leaf)) {
                throw "Destination changed but rollback backup is missing: $DestinationPath"
            }
            [IO.File]::Replace($BackupPath, $DestinationPath, $FailedCandidatePath, $true)
        }
        $restoredHash = (Get-FileHash -LiteralPath $DestinationPath -Algorithm SHA256).Hash.ToLowerInvariant()
        if ($restoredHash -cne $OriginalSha256) {
            throw "Rollback did not restore the original destination hash: $DestinationPath"
        }
        if (Test-Path -LiteralPath $BackupPath -PathType Leaf) { [IO.File]::Delete($BackupPath) }
        if (Test-Path -LiteralPath $FailedCandidatePath -PathType Leaf) { [IO.File]::Delete($FailedCandidatePath) }
        return "RESTORED_ORIGINAL"
    }
    if (Test-Path -LiteralPath $DestinationPath -PathType Leaf) { [IO.File]::Delete($DestinationPath) }
    return "RESTORED_ABSENCE"
}

$taskRoot = Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))
$contractPath = Join-Path $taskRoot "tests\confirm\contract.json"
$contract = Get-Content -LiteralPath $contractPath -Raw | ConvertFrom-Json
if ($contract.status -ne "TASK_INVALID" -or $contract.blocker -ne "PENDING_EXTERNAL_WINDOWS_EXCEL") {
    throw "Confirm split must remain TASK_INVALID/PENDING_EXTERNAL_WINDOWS_EXCEL until this native build succeeds."
}
if ($null -ne $contract.reference_workbook) {
    throw "Confirm contract must not claim a reference workbook before native Excel creates it."
}

$expected = $contract.expected_native_objects
if ($expected.table.name -ne "ProgramEventsTable" -or $expected.table.range -ne "Program_Data!A3:F11") {
    throw "Confirm contract has an unexpected Table identity or range."
}
if (($expected.table.headers -join "|") -ne "Event ID|Region|Program|Quarter|Participants|Spend") {
    throw "Confirm contract has unexpected exact source-field identities."
}
if ($expected.pivot_cache.source -ne "ProgramEventsTable") {
    throw "Confirm contract has an unexpected PivotCache source."
}
if (
    $expected.pivot_cache.field_indexes.Region -ne 1 -or
    $expected.pivot_cache.field_indexes.Program -ne 2 -or
    $expected.pivot_cache.field_indexes.Quarter -ne 3 -or
    $expected.pivot_cache.field_indexes.Participants -ne 4 -or
    $expected.pivot_cache.field_indexes.Spend -ne 5
) {
    throw "Confirm contract has unexpected zero-based PivotCache field indexes."
}
if ($expected.pivot_table.row_field -ne "Region" -or $expected.pivot_table.column_field -ne "Program" -or $expected.pivot_table.page_filter -ne "Quarter=2024Q2") {
    throw "Confirm contract has unexpected PivotTable axes/filter."
}
if (($expected.pivot_table.measures -join "|") -ne "Sum of Participants|Sum of Spend") {
    throw "Confirm contract has unexpected native SUM measures."
}
if (-not $expected.pivot_cache.refresh_on_open -or -not $expected.pivot_cache.enable_refresh) {
    throw "Confirm contract must require both refresh-on-open and enabled refresh."
}
if ($expected.pivot_chart.pivot_layout_table -ne "ProgramDeliveryPivot" -or $expected.pivot_chart.chart_type -ne "clustered column") {
    throw "Confirm contract has an unexpected PivotChart relationship or chart type."
}

$inputWorkbook = [IO.Path]::GetFullPath((Join-Path $taskRoot "tests\confirm\input_files\starting_workbook.xlsx"))
$nativeBuilder = Join-Path $taskRoot "windows\build_and_validate_native_pivot.ps1"
$judgePath = Join-Path $taskRoot "tests\evaluate.py"
$outputFull = [IO.Path]::GetFullPath($OutputWorkbook)
if ($outputFull -eq $inputWorkbook) {
    throw "OutputWorkbook must not overwrite the verifier-private confirm starter."
}
if ([IO.Path]::GetExtension($outputFull) -ne ".xlsx") {
    throw "OutputWorkbook must use the .xlsx extension: $outputFull"
}
$outputDirectory = Split-Path -Parent $outputFull
if (-not (Test-Path -LiteralPath $outputDirectory)) {
    [void](New-Item -ItemType Directory -Path $outputDirectory)
}
if ([string]::IsNullOrWhiteSpace($ReceiptPath)) {
    $ReceiptPath = "$outputFull.confirm-windows-validation.json"
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
$timestamp = (Get-Date).ToUniversalTime().ToString("yyyyMMddTHHmmssfffZ")
$outputStem = [IO.Path]::GetFileNameWithoutExtension($outputFull)
$confirmStageFull = Join-Path $outputDirectory (".{0}.confirm-judge.{1}.staging.xlsx" -f $outputStem, $token)
$nativeReceiptStageFull = Join-Path $receiptDirectory (".{0}.{1}.native-builder.json" -f [IO.Path]::GetFileName($receiptFull), $token)
$receiptStageFull = Join-Path $receiptDirectory (".{0}.{1}.confirm.staging.json" -f [IO.Path]::GetFileName($receiptFull), $token)
$judgeStderrFull = Join-Path $receiptDirectory (".{0}.{1}.confirm-judge.stderr.txt" -f [IO.Path]::GetFileName($receiptFull), $token)
if ((Test-Path -LiteralPath $outputFull) -and -not (Test-Path -LiteralPath $outputFull -PathType Leaf)) {
    throw "OutputWorkbook exists but is not a file: $outputFull"
}
if ((Test-Path -LiteralPath $receiptFull) -and -not (Test-Path -LiteralPath $receiptFull -PathType Leaf)) {
    throw "ReceiptPath exists but is not a file: $receiptFull"
}
$outputExisted = Test-Path -LiteralPath $outputFull -PathType Leaf
$receiptExisted = Test-Path -LiteralPath $receiptFull -PathType Leaf
$outputOriginalHash = if ($outputExisted) { (Get-FileHash -LiteralPath $outputFull -Algorithm SHA256).Hash.ToLowerInvariant() } else { $null }
$receiptOriginalHash = if ($receiptExisted) { (Get-FileHash -LiteralPath $receiptFull -Algorithm SHA256).Hash.ToLowerInvariant() } else { $null }
$outputBackupFull = $null
$receiptBackupFull = $null
$failedOutputFull = Join-Path $outputDirectory (".{0}.failed-confirm-promotion.{1}.xlsx" -f $outputStem, $token)
$failedReceiptFull = Join-Path $receiptDirectory (".{0}.failed-confirm-promotion.{1}.json" -f [IO.Path]::GetFileNameWithoutExtension($receiptFull), $token)
if ($outputExisted) {
    $outputBackupFull = Join-Path $outputDirectory ("{0}.pre-confirm-promotion.{1}.{2}.backup.xlsx" -f $outputStem, $timestamp, $token)
}
if ($receiptExisted) {
    $receiptBackupFull = Join-Path $receiptDirectory ("{0}.pre-confirm-promotion.{1}.{2}.backup.json" -f [IO.Path]::GetFileNameWithoutExtension($receiptFull), $timestamp, $token)
}

try {
    # The native builder writes only to this wrapper-owned candidate. The caller's
    # OutputWorkbook remains untouched until the explicit confirm Judge passes.
    $nativeBuilderOutput = & $nativeBuilder `
        -InputWorkbook $inputWorkbook `
        -OutputWorkbook $confirmStageFull `
        -ReceiptPath $nativeReceiptStageFull `
        -ExpectedQ2Participants ([double]$expected.oracle.q2_participants) `
        -ExpectedQ2Spend ([double]$expected.oracle.q2_spend) `
        -FocusRegion ([string]$expected.oracle.focus_region) `
        -FocusProgram ([string]$expected.oracle.focus_program) `
        -ExpectedFocusParticipants ([double]$expected.oracle.focus_participants)
    if (-not (Test-Path -LiteralPath $confirmStageFull -PathType Leaf)) {
        throw "Native builder did not produce the wrapper staging workbook."
    }
    if (-not (Test-Path -LiteralPath $nativeReceiptStageFull -PathType Leaf)) {
        throw "Native builder did not produce its staged receipt."
    }

    $candidateHash = (Get-FileHash -LiteralPath $confirmStageFull -Algorithm SHA256).Hash.ToLowerInvariant()
    $nativeReceipt = Get-Content -LiteralPath $nativeReceiptStageFull -Raw | ConvertFrom-Json
    Assert-Equal ([string]$nativeReceipt.status) "WINDOWS_EXCEL_COM_NATIVE_OBJECTS_VALIDATED" "Native builder receipt status"
    Assert-Equal ([string]$nativeReceipt.output_sha256) $candidateHash "Native builder candidate hash"
    Assert-Equal ([string]$nativeReceipt.pivot_chart.chart_type) "clustered column" "Native builder PivotChart type readback"
    Assert-Equal ([bool]$nativeReceipt.pivot_chart.reopened_readback) $true "Native builder reopened PivotChart readback"

    if (Test-Path -LiteralPath $judgeStderrFull -PathType Leaf) { [IO.File]::Delete($judgeStderrFull) }
    $judgeJson = @(& $PythonExe $judgePath $confirmStageFull --split confirm 2> $judgeStderrFull)
    $judgeExitCode = $LASTEXITCODE
    $judgeStderr = if (Test-Path -LiteralPath $judgeStderrFull) { @(Get-Content -LiteralPath $judgeStderrFull) } else { @() }
    if ($judgeExitCode -ne 0) {
        throw "Confirm Judge invocation failed with exit code $judgeExitCode. stderr: $($judgeStderr -join ' | ')"
    }
    $judge = (($judgeJson -join [Environment]::NewLine) | ConvertFrom-Json)
    $judgeChecks = [ordered]@{
        status = (($judge.PSObject.Properties.Name -contains "status") -and ([string]$judge.status -ceq "NATIVE_OBJECT_CHECKED"))
        task_id = (($judge.PSObject.Properties.Name -contains "task_id") -and ([string]$judge.task_id -ceq "P15-B-PUBLIC-PIVOT-001"))
        split = (($judge.PSObject.Properties.Name -contains "split") -and ([string]$judge.split -ceq "confirm"))
        pass = (($judge.PSObject.Properties.Name -contains "pass") -and ($judge.pass -is [bool]) -and ($judge.pass -eq $true))
        score = (($judge.PSObject.Properties.Name -contains "normalized_score") -and ($judge.normalized_score -is [ValueType]) -and ([double]$judge.normalized_score -eq 1.0))
        failure_codes_empty = (($judge.PSObject.Properties.Name -contains "failure_codes") -and ($judge.failure_codes -is [Array]) -and (@($judge.failure_codes).Count -eq 0))
        stderr_empty = ($judgeStderr.Count -eq 0)
    }
    $failedJudgeChecks = @($judgeChecks.GetEnumerator() | Where-Object { -not [bool]$_.Value } | ForEach-Object { [string]$_.Key })
    if ($failedJudgeChecks.Count -ne 0) {
        throw "Native confirm candidate failed Judge contract checks: $($failedJudgeChecks -join ', '). stderr: $($judgeStderr -join ' | ')"
    }

    # Rebind the receipt from the wrapper-owned candidate to the final caller path.
    # The content hash is stable across the same-directory atomic rename/replace.
    $nativePromotion = $nativeReceipt.promotion
    $nativeReceipt.output_workbook = $outputFull
    $nativeReceipt.output_sha256 = $candidateHash
    $nativeReceipt | Add-Member -NotePropertyName split -NotePropertyValue "confirm" -Force
    $nativeReceipt | Add-Member -NotePropertyName judge -NotePropertyValue $judge -Force
    $nativeReceipt | Add-Member -NotePropertyName judge_contract -NotePropertyValue ([ordered]@{
        checks = $judgeChecks
        stderr = @($judgeStderr)
    }) -Force
    $nativeReceipt | Add-Member -NotePropertyName native_builder_promotion -NotePropertyValue $nativePromotion -Force
    $nativeReceipt | Add-Member -NotePropertyName promotion -NotePropertyValue ([ordered]@{
        strategy = $(if ($outputExisted) { "same-directory File.Replace after confirm Judge" } else { "same-directory File.Move after confirm Judge" })
        replaced_existing_output = [bool]$outputExisted
        previous_output_backup = $outputBackupFull
        previous_receipt_backup = $receiptBackupFull
        receipt_published_before_workbook_promotion = $true
        receipt_rollback_on_workbook_promotion_failure = $true
        hash_basis = "Judge-passed staging bytes preserved by same-directory atomic promotion"
    }) -Force

    # Complete receipt serialization and atomic publication precede caller output
    # promotion. If workbook promotion fails, restore the prior receipt (or absence).
    Write-Utf8Json $receiptStageFull $nativeReceipt 14
    $receiptCandidateHash = (Get-FileHash -LiteralPath $receiptStageFull -Algorithm SHA256).Hash.ToLowerInvariant()
    try {
        Promote-FileAtomically $receiptStageFull $receiptFull ([bool]$receiptExisted) $receiptBackupFull
        $publishedReceiptHash = (Get-FileHash -LiteralPath $receiptFull -Algorithm SHA256).Hash.ToLowerInvariant()
        if ($publishedReceiptHash -cne $receiptCandidateHash) {
            throw "Published confirm receipt hash differs from its fully validated staging bytes."
        }
        Promote-FileAtomically $confirmStageFull $outputFull ([bool]$outputExisted) $outputBackupFull
        $publishedWorkbookHash = (Get-FileHash -LiteralPath $outputFull -Algorithm SHA256).Hash.ToLowerInvariant()
        if ($publishedWorkbookHash -cne $candidateHash) {
            throw "Published confirm workbook hash differs from its Judge-passed staging bytes."
        }
    } catch {
        $promotionError = $_
        $rollbackErrors = @()
        try {
            [void](Restore-FileAfterFailedPromotion $outputFull ([bool]$outputExisted) $outputOriginalHash $outputBackupFull $failedOutputFull)
        } catch {
            $rollbackErrors += "output: $($_.Exception.Message)"
        }
        try {
            [void](Restore-FileAfterFailedPromotion $receiptFull ([bool]$receiptExisted) $receiptOriginalHash $receiptBackupFull $failedReceiptFull)
        } catch {
            $rollbackErrors += "receipt: $($_.Exception.Message)"
        }
        if ($rollbackErrors.Count -ne 0) {
            throw "Confirm promotion failed and rollback could not prove original state. Promotion error: $($promotionError.Exception.Message); rollback errors: $($rollbackErrors -join ' | ')"
        }
        throw $promotionError
    }
    Write-Output ($nativeReceipt | ConvertTo-Json -Depth 14)
}
finally {
    # Only wrapper-owned staging files are cleaned. Caller output and any backup
    # created by a successful atomic promotion are retained.
    foreach ($path in @($confirmStageFull, $nativeReceiptStageFull, $receiptStageFull, $judgeStderrFull, $failedOutputFull, $failedReceiptFull)) {
        if (Test-Path -LiteralPath $path -PathType Leaf) {
            [IO.File]::Delete($path)
        }
    }
}
