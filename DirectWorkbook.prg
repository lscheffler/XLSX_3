lcTable = GETFILE("dbf", "table", "Export", 0, "Select Table to Export")
IF !EMPTY(lcTable)
	lcAlias = CHRTRAN(JUSTSTEM(lcTable), " ", "")

	IF !USED(lcAlias)
		USE (lcTable) IN 0 EXCLUSIVE ALIAS &lcAlias
	ENDIF
	SELECT &lcAlias

	loExcel = NEWOBJECT("VFPxWorkbookXLSX", "VFPxWorkbookXLSX.vcx")

	loExcel.Savetabletoworkbook(lcAlias, lcAlias + "_test1.xlsx", .T., .T., lcAlias)

	loExcel.Savetabletoworkbookex(lcAlias, lcAlias + "_test2.xlsx", .NULL., .T., lcAlias)

	USE IN SELECT(lcAlias)
ENDIF