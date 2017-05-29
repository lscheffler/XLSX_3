PUBLIC goExcel   && to keep it from being destroyed and closing the cursors
LOCAL lcFile, lnWb, loSheets, lnSh, lnRow, lnCol
lcFile = GETFILE("xlsx", "Workbook", "Load", 0, "Select Workbook to load into Class")
IF !EMPTY(lcFile)
	goExcel = NEWOBJECT("VFPxWorkbookXLSX", "VFPxWorkbookXLSX.vcx")
	lnWb = goExcel.OpenXlsxWorkbook(lcFile)
	loSheets = goExcel.GetWorkbookSheets(lnWb)
	FOR lnSh=1 TO loSheets.Count
		? loSheets.List[lnSh, 1]      && Displays sheet index (which may not be the same as lnSh)
		? loSheets.List[lnSh, 2]      && Displays sheet name
		FOR lnRow=1 TO goExcel.GetLastRowNumber(lnWb, loSheets.List[lnSh, 1])
			FOR lnCol=1 TO goExcel.GetMaxColumnNumber(lnWb, loSheets.List[lnSh, 1])
				IF goExcel.IsCellFormula(lnWb, loSheets.List[lnSh, 1], lnRow, lnCol)
					? goExcel.GetCellFormula(lnWb, loSheets.List[lnSh, 1], lnRow, lnCol)
				ELSE
					? goExcel.GetCellValue(lnWb, loSheets.List[lnSh, 1], lnRow, lnCol)
				ENDIF
			ENDFOR
		ENDFOR
	ENDFOR
ENDIF