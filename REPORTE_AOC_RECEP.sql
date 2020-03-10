-- Base de dato utilizado y actualizar tabla creada
USE Planeamiento
DROP TABLE REPORTE_AOC_RECEP_TABLA
DROP TABLE #lista_TIP_CAT
DROP TABLE #lista_TIP_CAT_FINAL
DROP TABLE #tmp
GO
-- IMPORTAR DATA DE RECECPCION
SELECT * INTO #lista_TIP_CAT_FINAL
FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0',
		'Excel 12.0;Database=D:\Documentos Personales\Documents\Planeamiento\Scripts_Planeamiento\Reporte_AOC_RECEP\Analisis_OrdenCompra\Doc_Origen\PROP_DOC\Lista PAC 2020 & Precios HistoricosV1.xlsx;HDR=YES', 
		'SELECT * FROM [precios$]')
GO
SELECT * INTO #lista_TIP_CAT
FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0',
		'Excel 12.0;Database=D:\Documentos Personales\Documents\Planeamiento\Scripts_Planeamiento\Reporte_AOC_RECEP\Analisis_OrdenCompra\Doc_Origen\PROP_DOC\Lista PAC 2020 & Precios HistoricosV1.xlsx;HDR=YES', 
		'SELECT * FROM [lista$]')
GO

SELECT aoc.FONDO, aoc.ITEM, aoc.ORDEN, aoc.FECHA_ENTREGA, 
	   CASE  
			WHEN aoc.COMM = '26010248' THEN '26010060'
			WHEN aoc.COMM = '26010251' THEN '26010090' 
			WHEN aoc.COMM = '26010245' THEN '26010026'
			WHEN aoc.COMM = '26010258' THEN '27010076'
			WHEN aoc.COMM = '26010246' THEN '26010027'
			WHEN aoc.COMM = '26010253' THEN '26010125'
			WHEN aoc.COMM = '26010254' THEN '26010160' 
			WHEN aoc.COMM = '26010252' THEN '26010101'
			WHEN aoc.COMM = '26010255' THEN '26010161' 
			WHEN aoc.COMM = '26010247' THEN '26010028' 
			WHEN aoc.COMM = '26010249' THEN '26010085'
			WHEN aoc.COMM = '26010257' THEN '26010229' 
			WHEN aoc.COMM = '26010259' THEN '26010059' 
			WHEN aoc.COMM = '26010260' THEN '26010159'
			WHEN aoc.COMM = '26010261' THEN '26010098' 
			WHEN aoc.COMM = '26010262' THEN '26010099' 
			WHEN aoc.COMM = '26010263' THEN '26010126'
			WHEN aoc.COMM = '26010256' THEN '26010197' 
			ELSE aoc.COMM 
	   END AS COMM, 
       IIF(aoc.BIEN LIKE '=%',RIGHT(LEFT(aoc.BIEN,LEN(aoc.BIEN)-1),LEN(LEFT(aoc.BIEN,LEN(aoc.BIEN)-1))-2),aoc.BIEN) AS BIEN, 
	   aoc.UND, aoc.PROVEEDOR, 
	   RIGHT(aoc.COMPRADOR,LEN(aoc.COMPRADOR)-7) AS COMPRADOR, 
	   --SUM(aoc.CANT_ORDEN2) AS CANT_ORDEN2, 
	   aoc.IMPORTE_BIEN2,
	   CAST (ISNULL (SUM(aoc.CANT_ORDEN2),0) AS float) AS CANT_ORDEN, 
	   CAST (ISNULL(recep.CANT_ENTREGADA,0) AS float) AS CANT_ENTRE, 
	   CAST (ISNULL(recep.CANT_RECIBIDA,0) AS float) AS CANT_RECEP,
	   --recep.CANT_ENTREGADA, recep.CANT_RECIBIDA,
	   glog.ORIGEN

INTO #tmp
FROM AOC_TABLA AS aoc
LEFT JOIN RECEP_TABLA AS recep ON recep.FONDO = aoc.FONDO 
								  AND recep.ORDEN_COMPRA  = aoc.ORDEN
								  AND recep.CODIGO_BIEN = aoc.COMM 
								  AND aoc.ITEM = recep.ELEMENTO
LEFT JOIN COMPRA_GLOG AS glog ON glog.COMPRADOR = aoc.COMPRADOR
WHERE aoc.CANT_ORDEN2 > 0
GROUP BY aoc.ORDEN, aoc.FONDO, aoc.ITEM, aoc.COMM, aoc.BIEN, aoc.UND, aoc.PROVEEDOR, aoc.COMPRADOR,
		 aoc.FECHA_ENTREGA, aoc.IMPORTE_BIEN2, recep.CANT_ENTREGADA, recep.CANT_RECIBIDA, 
		 recep.ORDEN_COMPRA, recep.CODIGO_BIEN,glog.ORIGEN
ORDER BY CODIGO_BIEN
GO

DELETE FROM #tmp WHERE CANT_ORDEN=0

SELECT fond.FUND + ' - ' + fond.FONDO AS FONDO, 
	   tmp.ITEM,tmp.ORDEN, tmp.FECHA_ENTREGA,
	   IIF(lista.BIEN_PREDESEDOR2 IS NULL,'SIN CATEGORIA',lista.BIEN_PREDESEDOR2) AS CATEGORIA,
	   tmp.COMM AS CODIGO, tmp.BIEN AS DESCRIPCION_BIEN, tmp.UND, tmp.PROVEEDOR, tmp.COMPRADOR,
	   tmp.CANT_ORDEN, tmp.CANT_ENTRE, tmp.CANT_RECEP,
	   --CAST (ISNULL (tmp.CANT_ORDEN2,0) AS float) AS CANT_ORDEN, 
	   --CAST (ISNULL(tmp.CANT_ENTREGADA,0) AS float) AS CANT_ENTRE, 
	   --CAST (ISNULL(tmp.CANT_RECIBIDA,0) AS float) AS CANT_RECEP,
	   ROUND(CAST(tmp.IMPORTE_BIEN2 AS float),3) AS IMPORTE_BIEN,ISNULL(tmp.ORIGEN,'DZ') AS ORIGEN,
	   lista.TIPO_CATEGORIA, --precios.TIPO_CAT_FINAL
	   CASE WHEN precios.TIPO_CAT_FINAL = 'CORP_NOAD' THEN 'NO_CORP'
			WHEN precios.TIPO_CAT_FINAL = 'NO_CORP' THEN 'NO_CORP'
			ELSE 'CORP' END AS TIPO_CAT_FINAL,
      IIF(tmp.CANT_ENTRE/tmp.CANT_ORDEN >= 1, 'COMPLETA','PENDIENTE') AS CONDICION_LINEA
   
INTO REPORTE_AOC_RECEP_TABLA
FROM #tmp AS tmp
LEFT JOIN FONDOS AS fond ON fond.FUND = tmp.FONDO
LEFT JOIN #lista_TIP_CAT AS lista ON lista.COD_BIEN = tmp.COMM
LEFT JOIN #lista_TIP_CAT_FINAL AS precios ON precios.COD_BIEN = tmp.COMM
ORDER BY fond.FUND
GO

/*SELECT * FROM REPORTE_AOC_RECEP_TABLA WHERE ORDEN = 'P0213468' AND CODIGO = '91020156' --AND ORDEN = 'P0213468'
SELECT * FROM AOC_TABLA WHERE ORDEN = 'P0213468' AND COMM = '91020156'
SELECT * FROM RECEP_TABLA WHERE ORDEN_COMPRA = 'P0213468' AND CODIGO_BIEN = '91020156'*/