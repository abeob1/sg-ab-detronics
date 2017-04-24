DROP PROCEDURE "AB_VENDOR";
CREATE PROCEDURE "AB_VENDOR"
(
 IN DateFrom DATE,
 IN DateTo DATE,
 IN nobalance  VARCHAR(10)
)
AS
FromDate VARCHAR(20);
ToDate VARCHAR(20);
BEGIN
	
FromDate:=TO_CHAR(:DateFrom ,'YYYY-MM-DD');
ToDate:=TO_CHAR(:DateTo ,'YYYY-MM-DD');

CREATE COLUMN TABLE "TBFINAL"(CardCode VARCHAR(20),CardName varchar(100), NumAtCard varchar(100),
								DocDate DATETIME,DocDueDate DATETIME, PoNumber VARCHAR(50),
								DocTotal numeric(18,3),Balance numeric(18,3), AppliedAmount Numeric(18,3),
								DocCur varchar(10),DocType varchar(10),TransNo INTEGER);
INSERT INTO "TBFINAL"
SELECT DISTINCT P."CardCode", P."CardName", p."VENDor_ref",p."DocDate",P."DocDueDate",P."PoNumber",P."DocTotal",
(P."DocTotal"-P."AppliedAmount") AS "Balance",P."AppliedAmount",P."DocCur",P."Document_Type" 
,P."TransNo"
from 
(select  A."CardCode", A."CardName",
CASE 
	WHEN B."ObjType"='18' THEN B."NumAtCard" 
	WHEN C."ObjType"='19' THEN C."NumAtCard" 
	WHEN D."ObjType"='204' THEN D."NumAtCard" 
 END as "VENDor_ref",
CASE 
	WHEN B."ObjType"='18' THEN B."DocDate" 
	WHEN C."ObjType"='19' THEN C."DocDate" 
	WHEN D."ObjType"='204' THEN D."DocDate"
 END AS "DocDate",
CASE WHEN B."ObjType"='18' THEN B."DocDueDate" WHEN C."ObjType"='19' THEN C."DocDueDate" WHEN D."ObjType"='204' THEN D."DocDueDate"
 END AS "DocDueDate",
CASE WHEN TO_CHAR(B."ObjType")='18' THEN 'IN' WHEN TO_CHAR(C."ObjType")='19' THEN 'CN' WHEN TO_CHAR(D."ObjType")='204' THEN 'PI'
 END AS "Document_Type",
CASE WHEN B."ObjType"='18' THEN E."BaseRef" WHEN C."ObjType"='19' THEN F."BaseRef" WHEN D."ObjType"='204' THEN G."BaseRef"
 END AS "PoNumber",
CASE WHEN B."ObjType"='18' AND B."CurSource"='L' AND A."BASE_REF"=B."DocNum" THEN B."DocTotal"
	WHEN B."ObjType"='18' AND B."CurSource"='C' AND A."BASE_REF"=B."DocNum" THEN B."DocTotalFC"
	WHEN C."ObjType"='19' AND C."CurSource"='L' AND A."BASE_REF"=C."DocNum" THEN B."DocTotal" 
	WHEN C."ObjType"='19' AND C."CurSource"='C' AND A."BASE_REF"=C."DocNum" THEN B."DocTotalFC"
	WHEN D."ObjType"='204' AND D."CurSource"='L' AND A."BASE_REF"=D."DocNum" THEN B."DocTotal"
	WHEN D."ObjType"='204' AND D."CurSource"='C' AND A."BASE_REF"=D."DocNum" THEN B."DocTotalFC"
	ELSE 0 END AS "DocTotal",
CASE WHEN B."ObjType"='18' AND A."BASE_REF"=B."DocNum" THEN B."PaidToDate" 
	WHEN C."ObjType"='19' AND A."BASE_REF"=C."DocNum" THEN C."PaidToDate"
	WHEN D."ObjType"='204' AND A."BASE_REF"=D."DocNum" THEN D."PaidToDate"
	ELSE 0 END AS "AppliedAmount",
CASE WHEN B."ObjType"='18' AND A."BASE_REF"=B."DocNum" THEN B."DocCur" 
	WHEN C."ObjType"='19' AND A."BASE_REF"=C."DocNum" THEN C."DocCur"
	WHEN D."ObjType"='204' AND A."BASE_REF"=D."DocNum" THEN D."DocCur"
	ELSE '' END as "DocCur",

CASE WHEN B."ObjType"='18' AND A."BASE_REF"=B."DocNum" THEN B."DocNum" 
	WHEN C."ObjType"='19' AND A."BASE_REF"=C."DocNum" THEN C."DocNum"
	WHEN D."ObjType"='204' AND A."BASE_REF"=D."DocNum" THEN D."DocNum"
	ELSE 0 END as "TransNo"
		
from "OINM" A
left join "OPCH" B on A."TransType"=B."ObjType" AND A."BASE_REF"=B."DocNum"
left join "PCH1" E on B."DocEntry"=E."DocEntry"
left join "ORPC" C on A."TransType"=C."ObjType" AND A."BASE_REF"=C."DocNum"
left join "RPC1" F on C."DocEntry"=F."DocEntry"
left join "ODPO" D on A."TransType"=D."ObjType" AND A."BASE_REF"=D."DocNum"
left join "DPO1" G on D."DocEntry"=G."DocEntry"
where A."TransType" in ('18','19','204')
--AND B."DocStatus"<> 'C' OR C."DocStatus"<> 'C' OR D."DocStatus"<> 'C'
AND B."CANCELED"<> 'Y' OR C."CANCELED"<> 'Y' OR D."CANCELED"<> 'Y'

union all
select  B."CardCode", B."CardName",
CASE 
	WHEN B."ObjType"='18' THEN B."NumAtCard" 

 END as "VENDor_ref",
CASE 
	WHEN B."ObjType"='18' THEN B."DocDate" 
 END AS "DocDate",
CASE WHEN B."ObjType"='18' THEN B."DocDueDate" 
 END AS "DocDueDate",
CASE WHEN TO_CHAR(B."ObjType")='18' THEN 'IN' 
 END AS "Document_Type",
CASE WHEN B."ObjType"='18' THEN E."BaseRef" 
 END AS "PoNumber",
CASE WHEN B."ObjType"='18' AND B."CurSource"='L' THEN B."DocTotal"
	WHEN B."ObjType"='18' AND B."CurSource"='C' THEN B."DocTotalFC"
	
	ELSE 0 END AS "DocTotal",
CASE WHEN B."ObjType"='18' THEN B."PaidToDate" 
	
	ELSE 0 END AS "AppliedAmount",
CASE WHEN B."ObjType"='18' THEN B."DocCur" 
	
	ELSE '' END as "DocCur",

CASE WHEN B."ObjType"='18'  THEN B."DocNum" 
	
	ELSE 0 END as "TransNo"
		
from "OPCH" B

inner join "PCH1" E on B."DocEntry"=E."DocEntry"

--AND B."DocStatus"<> 'C' 
AND B."CANCELED"<> 'Y' 
AND B."DocType" = 'S'

)P 
WHERE P."DocDate" between :FromDate AND :ToDate;

if :nobalance = 'YES' THEN
	SELECT (select Top 1 "LogoImage" from "OADP") as "LogoImage", 
	(SELECT "CompnyName" FROM OADM ) AS "CompnyName",
	(SELECT A."AliasName" FROM OADM A ) AS "AliasName",
	(SELECT B."StreetNo" FROM ADM1 B) AS "StreetNo",
	(SELECT A."Phone1" FROM OADM A) AS "Phone1", 
	(SELECT  A."Fax" FROM OADM A) AS "Fax",
	(SELECT  B."IntrntAdrs" FROM ADM1 B)AS "IntrntAdrs",* FROM "TBFINAL";
ELSE 
	SELECT (select Top 1 "LogoImage" from "OADP") as "LogoImage", 
	(SELECT "CompnyName" FROM OADM ) AS "CompnyName",
	(SELECT A."AliasName" FROM OADM A ) AS "AliasName",
	(SELECT B."StreetNo" FROM ADM1 B) AS "StreetNo",
	(SELECT A."Phone1" FROM OADM A) AS "Phone1", 
	(SELECT  A."Fax" FROM OADM A) AS "Fax",
	(SELECT  B."IntrntAdrs" FROM ADM1 B)AS "IntrntAdrs",* FROM "TBFINAL" WHERE Balance > 0;
END IF;

drop table "TBFINAL";

END