on run argv
	set workbookPath to item 1 of argv
	set workbookName to item 2 of argv
	with timeout of 120 seconds
		tell application "Microsoft Excel"
			set display alerts to false
			open POSIX file workbookPath
			set wb to workbook workbookName
			set reportSheet to worksheet "Pivot_Report" of wb
			set reportPivot to pivot table "ProgramDeliveryPivot" of reportSheet
			set position of data field "Sum of Spend" of reportPivot to 1
			set position of data field "Sum of Participants" of reportPivot to 2
			set embeddedChart to chart object 1 of reportSheet
			set left position of embeddedChart to 520
			set top of embeddedChart to 45
			set reportChart to chart of embeddedChart
			set has title of reportChart to true
			set formula of chart title of reportChart to "Q2 Program Delivery — Alternate Native Layout"
			refresh pivot cache of reportPivot
			refresh table reportPivot
			calculate full rebuild
			save workbook as wb filename workbookPath file format Excel XML file format
			return (name of data field 1 of reportPivot) & "|" & (name of data field 2 of reportPivot) & "|" & (count of chart objects of reportSheet)
		end tell
	end timeout
end run
