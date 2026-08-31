on run argv
	set workbookName to item 1 of argv
	set cellAddress to item 2 of argv
	set newValue to (item 3 of argv) as real
	with timeout of 90 seconds
		tell application "Microsoft Excel"
			set display alerts to false
			set wb to workbook workbookName
			set value of range cellAddress of worksheet "Program_Data" of wb to newValue
			set reportPivot to pivot table "ProgramDeliveryPivot" of worksheet "Pivot_Report" of wb
			refresh pivot cache of reportPivot
			refresh table reportPivot
			calculate full rebuild
			set participants to value of range "B4" of worksheet "KPI_Summary" of wb
			set spend to value of range "B5" of worksheet "KPI_Summary" of wb
			set focusValue to value of range "B6" of worksheet "KPI_Summary" of wb
			return (participants as text) & "|" & (spend as text) & "|" & (focusValue as text)
		end tell
	end timeout
end run
