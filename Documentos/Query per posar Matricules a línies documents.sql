--> Números de Séria per línia de comanda
Select (SELECT  LTrim(RTrim(Replace(          
					( select ' ' + OSRN.DistNumber + ' '
			--select p.LineNum, OSRN.DistNumber, Sum(ITL1.Quantity) --, p.LineNum, *
					From RDR1 p  
					left JOIN OITL ON p.DocEntry = OITL.ApplyEntry AND OITL.ApplyType='17' and OITL.ApplyLine=p.LineNum 
					left JOIN ITL1 ON OITL.LogEntry = ITL1.LogEntry  
					left JOIN OSRN ON ITL1.ItemCode = OSRN.ItemCode AND ITL1.SysNumber = OSRN.SysNumber
					where p.DocEntry=pp.DocEntry and p.LineNum=pp.LineNum
					group by p.LineNum, OSRN.DistNumber
					having Sum(ITL1.AllocQty)>0
					for xml path('')  
					),'  ', ', ')))
		), U_SEIMatric, *
From RDR1 pp  
Where pp.DocEntry=1 --and p.LineNum=2
 

 --> Números de Séria per línia d'entrega
Select (SELECT  LTrim(RTrim(Replace(          
					( select ' ' + OSRN.DistNumber + ' '
			--select p.LineNum, OSRN.DistNumber, Sum(ITL1.Quantity) --, p.LineNum, *
					From DLN1 p  
					left JOIN OITL ON p.DocEntry = OITL.ApplyEntry AND OITL.ApplyType='15' and OITL.ApplyLine=p.LineNum 
					left JOIN ITL1 ON OITL.LogEntry = ITL1.LogEntry  
					left JOIN OSRN ON ITL1.ItemCode = OSRN.ItemCode AND ITL1.SysNumber = OSRN.SysNumber
					where p.DocEntry=pp.DocEntry and p.LineNum=pp.LineNum
					group by p.LineNum, OSRN.DistNumber
					having Sum(ITL1.Quantity)<0
					for xml path('')  
					),'  ', ', ')))
		), U_SEIMatric, *
From DLN1 pp  
Where pp.DocEntry=6 --and p.LineNum=2


 --> Números de Séria per línia d'abonament
Select (SELECT  LTrim(RTrim(Replace(          
					( select ' ' + OSRN.DistNumber + ' '
			--select p.LineNum, OSRN.DistNumber, Sum(ITL1.Quantity) --, p.LineNum, *
					From RIN1 p  
					left JOIN OITL ON p.DocEntry = OITL.ApplyEntry AND OITL.ApplyType='14' and OITL.ApplyLine=p.LineNum 
					left JOIN ITL1 ON OITL.LogEntry = ITL1.LogEntry  
					left JOIN OSRN ON ITL1.ItemCode = OSRN.ItemCode AND ITL1.SysNumber = OSRN.SysNumber
					where p.DocEntry=pp.DocEntry and p.LineNum=pp.LineNum
					group by p.LineNum, OSRN.DistNumber
					having Sum(ITL1.Quantity)>0
					for xml path('')  
					),'  ', ', ')))
		), U_SEIMatric, *
From RIN1 pp  
Where pp.DocEntry=935 --and p.LineNum=2
--select * from oRIN



