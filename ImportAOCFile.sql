USE Planeamiento
--- IMPORT FROM EXCEL FILE TO DATABASE USING xlsx EXTENSION
--- NAME'S TAB "Hoja1" 
DROP TABLE AOC_TMP
DROP TABLE PROP_TMP
DROP TABLE #Tabla_AOC
DROP TABLE #AOC_ADD
GO

SELECT * INTO #Tabla_AOC
FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0',
		'Excel 12.0;Database=D:\Documentos Personales\Documents\Planeamiento\Scripts_Planeamiento\Reporte_AOC_RECEP\Analisis_OrdenCompra\Doc_Origen\AOC_DOC\Anlisis de orden de compra_ 01082019_11032020.xlsx;HDR=YES', 
		'SELECT * FROM [Anlisis de orden de compra_ 010$]')

SELECT * INTO #AOC_ADD
FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0',
		'Excel 12.0;Database=D:\Documentos Personales\Documents\Planeamiento\Scripts_Planeamiento\Reporte_AOC_RECEP\Analisis_OrdenCompra\Doc_Origen\AOC_DOC\Añadir_AOC.xlsx;HDR=YES', 
		'SELECT * FROM [Hoja1$]')

SELECT * INTO PROP_TMP
FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0',
		'Excel 12.0;Database=D:\Documentos Personales\Documents\Planeamiento\Scripts_Planeamiento\Reporte_AOC_RECEP\Analisis_OrdenCompra\Doc_Origen\PROP_DOC\Lista PAC 2020 & Precios HistoricosV1.xlsx;HDR=YES', 
		'SELECT * FROM [lista$]')
GO

SELECT * INTO AOC_TMP FROM #Tabla_AOC UNION SELECT * FROM #AOC_ADD
