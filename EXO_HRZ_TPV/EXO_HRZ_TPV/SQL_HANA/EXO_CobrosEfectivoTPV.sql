/****** Object:  UserDefinedFunction "EXO_CobrosEfectivoTPV"    Script Date: 01/02/2022 17:54:32 ******/
CREATE FUNCTION "EXO_CobrosEfectivoTPV"
(     
       IN DesdeFecha TIMESTAMP,
       IN HastaFecha TIMESTAMP,
	   IN Almacen nvarchar(100)
)

       RETURNS TABLE
       (
 			 "SeriesName" NVARCHAR(100),
             "DocEntry" INTEGER,
             "DocNum" INTEGER,
             "CardCode" NVARCHAR(50),
             "CardName" NVARCHAR(100),         
             "Tipo" NVARCHAR(50),
             "Total" DECIMAL(21,6),   
             "DocDate" TIMESTAMP
       ) LANGUAGE SQLSCRIPT

             SQL SECURITY INVOKER
             AS

 

       BEGIN
             RETURN

             SELECT T0."U_EXO_SERIE" as "SeriesName", T0."DocEntry" , T0."DocNum", T0."CardCode" , T0."CardName" , 
             T0."CounterRef" as "Tipo", T0."DocTotal"  AS "Total",T0."DocDate"
             FROM "ORCT" T0
             WHERE T0."Canceled" = 'N' and T0."CounterRef"='CAJA' 
                    and (CASE WHEN LEFT(T0."U_EXO_SERIE",4)='TQ-Q' THEN '02' 
       					  WHEN LEFT(T0."U_EXO_SERIE",4)='TQ-Z' THEN '01' 
       						ELSE '03' 
					END)= :Almacen
                     and T0."DocDate" >= :DesdeFecha AND T0."DocDate" <= :HastaFecha;

END;