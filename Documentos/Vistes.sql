--> Vista [SEI_EstocLotsArticles]
SELECT     TOP (100) PERCENT a.ItemCode AS CodiArticle, a.ItemName AS NomArticle, m.WhsCode AS Magatzem, m.OnHand AS EstocArticle, b.DistNumber AS Lot, l.Quantity AS EstocLot
FROM        dbo.OITM AS a LEFT OUTER JOIN
            dbo.OITW AS m ON m.ItemCode = a.ItemCode LEFT OUTER JOIN
            dbo.OBTQ AS l ON l.ItemCode = a.ItemCode AND l.WhsCode = m.WhsCode LEFT OUTER JOIN
            dbo.OBTN AS b ON l.ItemCode = b.ItemCode AND l.SysNumber = b.SysNumber


--> Vista [SEI_EstocMatriculasBobinas]
SELECT distinct a.ItemCode AS CodiArticle, a.ItemName AS NomArticle, m.WhsCode AS Magatzem, 
	m.OnHand AS EstocArticle, b.DistNumber AS Matricula, l.Quantity AS EstocMatricula 
FROM        dbo.OITM AS a 
LEFT OUTER JOIN   dbo.OITW AS m ON m.ItemCode = a.ItemCode 
LEFT OUTER JOIN   dbo.OSRQ AS l ON l.ItemCode = a.ItemCode AND l.WhsCode = m.WhsCode 
LEFT OUTER JOIN   dbo.OSRN AS b ON l.ItemCode = b.ItemCode AND l.SysNumber = b.SysNumber
where a.ItmsGrpCod=129 and a.ManSerNum ='Y' and IsNull(b.DistNumber,'')<>''


--> Vista [SEI_Families]
SELECT     ItmsGrpCod AS Codi, ItmsGrpNam AS NomFamilia
FROM        dbo.OITB