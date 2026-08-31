on run argv
	set targetWorkbookPath to item 1 of argv
	set targetWorkbookName to item 2 of argv
	set focusRegion to item 3 of argv
	set focusProgram to item 4 of argv

	with timeout of 180 seconds
		tell application "Microsoft Excel"
			set display alerts to false
			set targetWb to missing value
			try
				open POSIX file targetWorkbookPath
				delay 1
				set targetWb to workbook targetWorkbookName

				set reportSheet to worksheet "Pivot_Report" of targetWb
				set reportPivot to pivot table "ProgramDeliveryPivot" of reportSheet
				set reportCache to pivot cache of reportPivot
				set enable refresh of reportCache to true
				set refresh on file open of reportCache to true
				refresh reportCache
				refresh table reportPivot
				set current page of pivot field "Quarter" of reportPivot to "2024Q2"

				set formula of range "B4" of worksheet "KPI_Summary" of targetWb to "=GETPIVOTDATA(\"Sum of Participants\",Pivot_Report!$A$3)"
				set formula of range "B5" of worksheet "KPI_Summary" of targetWb to "=GETPIVOTDATA(\"Sum of Spend\",Pivot_Report!$A$3)"
				set formula of range "B6" of worksheet "KPI_Summary" of targetWb to "=GETPIVOTDATA(\"Sum of Participants\",Pivot_Report!$A$3,\"Region\",\"" & focusRegion & "\",\"Program\",\"" & focusProgram & "\")"

				refresh reportCache
				refresh table reportPivot
				calculate full rebuild
				save targetWb

				set reportSheet to worksheet "Pivot_Report" of targetWb
				set reportPivot to pivot table "ProgramDeliveryPivot" of reportSheet
				set reportCache to pivot cache of reportPivot
				set chartCount to count of chart objects of reportSheet
				set participantValue to value of range "B4" of worksheet "KPI_Summary" of targetWb
				set spendValue to value of range "B5" of worksheet "KPI_Summary" of targetWb
				set focusValue to value of range "B6" of worksheet "KPI_Summary" of targetWb
				set readbackText to (version as text) & "|" & (count of pivot caches of targetWb) & "|" & (count of pivot tables of reportSheet) & "|" & (count of data fields of reportPivot) & "|" & chartCount & "|" & (current page of pivot field "Quarter" of reportPivot as text) & "|" & participantValue & "|" & spendValue & "|" & focusValue
				close targetWb saving no
				return readbackText
			on error errorMessage number errorNumber
				if targetWb is not missing value then
					try
						close targetWb saving no
					end try
				end if
				error errorMessage number errorNumber
			end try
		end tell
	end timeout
end run
