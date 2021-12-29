/****** Object:  UserDefinedFunction "EXO_CobrosEfectivoTPV"    Script Date: 01/12/2021 17:54:32 ******/
SET SCHEMA "RODAMIENTOS_HURYZA";
CREATE FUNCTION "EXO_CobrosEfectivoTPV"
(	
	IN DesdeFecha TIMESTAMP,
	IN HastaFecha TIMESTAMP
)
	RETURNS TABLE 
	(
		"DocEntry" INTEGER, 
		"DocNum" INTEGER, 
		"CardCode" NVARCHAR(50), 
		"CardName" NVARCHAR(100), 
		"Efectivo" DECIMAL(21,6), 
		"DocDate" TIMESTAMP
	) LANGUAGE SQLSCRIPT 
		SQL SECURITY INVOKER 
		AS

	BEGIN
		RETURN 
		SELECT T0."DocEntry" , T0."DocNum", T0."CardCode" , T0."CardName" , T0."CashSumSy"  AS "Efectivo", T0."DocDate" 
 		FROM "ORCT" T0 
		WHERE T0."Canceled" = 'N' AND coalesce(T0."CashSumSy", 0) <> 0 
	  		and T0."DocDate" >= :DesdeFecha AND T0."DocDate" <= :HastaFecha;

END;

