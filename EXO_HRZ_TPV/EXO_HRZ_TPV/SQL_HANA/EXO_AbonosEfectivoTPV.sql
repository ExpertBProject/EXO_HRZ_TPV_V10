/****** Object:  UserDefinedFunction "EXO_AbonosEfectivoTPV"    Script Date: 01/02/2022 17:54:32 ******/
CREATE FUNCTION "EXO_AbonosEfectivoTPV"
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
             "U_EXO_CTIPO" NVARCHAR(50),
             "Total" DECIMAL(21,6),                  
             "U_EXO_CDOCNUM" INTEGER,
             "DocDate" TIMESTAMP,
			 "Almacen" NVARCHAR(50)
       ) LANGUAGE SQLSCRIPT

             SQL SECURITY INVOKER
             AS
 

       BEGIN
             RETURN
             (SELECT NNM1."SeriesName",T0."DocEntry" , T0."DocNum", T0."CardCode" , T0."CardName" , T0."U_EXO_CTIPO", T0."DocTotal" as "Total", 
					T0."U_EXO_CDOCNUM", T0."DocDate",
					(CASE WHEN LEFT(NNM1."SeriesName",5)='FR-TQ' THEN '02' 
       					  WHEN LEFT(NNM1."SeriesName",5)='FR-TZ' THEN '01' 
       						ELSE '03' 
					END) as "Almacen"
             FROM "ORIN" T0
             INNER JOIN NNM1 ON NNM1."ObjectCode"=T0."ObjType" and NNM1."Series"=T0."Series"
             WHERE T0."CANCELED" = 'N' and (T0."U_EXO_CTIPO"='C' )
					 and (CASE WHEN LEFT(NNM1."SeriesName",5)='FR-TQ' THEN '02' 
       					  WHEN LEFT(NNM1."SeriesName",5)='FR-TZ' THEN '01' 
       						ELSE '03' 
					END)= :Almacen
                    and T0."DocDate" >= :DesdeFecha AND T0."DocDate" <= :HastaFecha);

END;
