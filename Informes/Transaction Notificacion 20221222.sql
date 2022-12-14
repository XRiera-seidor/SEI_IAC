USE [IAC_PRU]
GO
/****** Object:  StoredProcedure [dbo].[SBO_SP_TransactionNotification]    Script Date: 22/12/2022 17:42:43 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
ALTER proc [dbo].[SBO_SP_TransactionNotification] 

@object_type nvarchar(30), 				-- SBO Object Type
@transaction_type nchar(1),			-- [A]dd, [U]pdate, [D]elete, [C]ancel, C[L]ose
@num_of_cols_in_key int,
@list_of_key_cols_tab_del nvarchar(255),
@list_of_cols_val_tab_del nvarchar(255)

AS

begin

-- Return values
declare @error  int				-- Result (0 for no error)
declare @error_message nvarchar (200) 		-- Error string to be displayed
select @error = 0
select @error_message = N'Ok'

--------------------------------------------------------------------------------------------------------------------------------

--	ADD	YOUR	CODE	HERE


If @object_type='17' and @transaction_type in ('A','U') --> Comandes de Venta
BEGIN
	--> Si el camp d'usuari d'Ordre que ve del GESTION està ple assumim aquest ordre visual
	Update RDR1 Set VisOrder=U_SEIOrV_G
	from RDR1
	Where IsNull(U_SEIOrV_G,-1)<>-1
		and DocEntry=@list_of_cols_val_tab_del

	--> Arrossegar les matrícules (números de sèrie) a les línies
	Update pp
	Set U_SEIMatric= (SELECT  LTrim(RTrim(Replace(          
												( select ' ' + OSRN.DistNumber + ' '
												From RDR1 p  
												left JOIN OITL ON p.DocEntry = OITL.ApplyEntry AND OITL.ApplyType='17' and OITL.ApplyLine=p.LineNum 
												left JOIN ITL1 ON OITL.LogEntry = ITL1.LogEntry  
												left JOIN OSRN ON ITL1.ItemCode = OSRN.ItemCode AND ITL1.SysNumber = OSRN.SysNumber
												where p.DocEntry=pp.DocEntry and p.LineNum=pp.LineNum
												group by p.LineNum, OSRN.DistNumber
												having Sum(ITL1.AllocQty)>0
												for xml path('')  
												),'  ', ', ')))
					)
	From RDR1 pp  
	Where pp.DocEntry=@list_of_cols_val_tab_del

END 

If @object_type='15' and @transaction_type in ('A','U') --> Entregues de Venta
BEGIN
	
	--> Arrossegar les matrícules (números de sèrie) a les línies
	Update pp
	Set U_SEIMatric= (SELECT  LTrim(RTrim(Replace(          
												( select ' ' + OSRN.DistNumber + ' '
												From DLN1 p  
												left JOIN OITL ON p.DocEntry = OITL.ApplyEntry AND OITL.ApplyType='15' and OITL.ApplyLine=p.LineNum 
												left JOIN ITL1 ON OITL.LogEntry = ITL1.LogEntry  
												left JOIN OSRN ON ITL1.ItemCode = OSRN.ItemCode AND ITL1.SysNumber = OSRN.SysNumber
												where p.DocEntry=pp.DocEntry and p.LineNum=pp.LineNum
												group by p.LineNum, OSRN.DistNumber
												having Sum(ITL1.Quantity)<0
												for xml path('')  
												),'  ', ', ')))
					) 
	From DLN1 pp  
	Where pp.DocEntry=@list_of_cols_val_tab_del

END 


If @object_type='14' and @transaction_type in ('A','U') --> Abonaments de Venta
BEGIN
	
	--> Arrossegar les matrícules (números de sèrie) a les línies
	Update pp
	Set U_SEIMatric= (SELECT  LTrim(RTrim(Replace(          
												( select ' ' + OSRN.DistNumber + ' '
												From RIN1 p  
												left JOIN OITL ON p.DocEntry = OITL.ApplyEntry AND OITL.ApplyType='14' and OITL.ApplyLine=p.LineNum 
												left JOIN ITL1 ON OITL.LogEntry = ITL1.LogEntry  
												left JOIN OSRN ON ITL1.ItemCode = OSRN.ItemCode AND ITL1.SysNumber = OSRN.SysNumber
												where p.DocEntry=pp.DocEntry and p.LineNum=pp.LineNum
												group by p.LineNum, OSRN.DistNumber
												having Sum(ITL1.Quantity)>0
												for xml path('')  
												),'  ', ', ')))
					) 
	From RIN1 pp  
	Where pp.DocEntry=@list_of_cols_val_tab_del

END 




--------------------------------------------------------------------------------------------------------------------------------
-- VALIDACIONES SII
IF @error = 0
BEGIN
EXEC [SEI_VALIDACIONES_SII] @object_type, @transaction_type,@num_of_cols_in_key,@list_of_key_cols_tab_del, @list_of_cols_val_tab_del, @error output, @error_message output
END
-- VALIDACIONES SII PERSONAL
IF @error = 0
BEGIN
EXEC [SEI_VALIDACIONES_SII_PERSONAL] @object_type, @transaction_type,@num_of_cols_in_key,@list_of_key_cols_tab_del, @list_of_cols_val_tab_del, @error output, @error_message output
END

-- Select the return values
select @error, @error_message

end