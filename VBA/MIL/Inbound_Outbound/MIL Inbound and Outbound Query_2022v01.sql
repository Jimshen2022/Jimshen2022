	
			
-- MIL Inbound Trx of Item class not like 'Z%'
SELECT t1.HOUSE,t1.TCODE,t1.ORDNO,t1.ITNBR,t2.ITCLS, t1.UPDDT,t1.UPDTM,t1.TRQTY,t1.ENTUM,t1.VNDNR,t1.REFNO,t1.LLOCN,t1.NLOCN,t1.BATCH,t1.TRMID,
CHAR(t1.UPDDT||' '||right('000000'||ltrim(t1.UPDTM),6)) AS "TrxTime"

FROM AMFLIBL.IMHIST  t1, AMFLIBL.ITMRVA t2, AMFLIBL.WHSMST t3
WHERE t1.ITNBR=t2.ITNBR  AND t2.STID = t3.STID AND t1.HOUSE = t3.WHID AND t1.TRQTY > 0 AND t1.TCODE IN ('RP','RM','PQ') AND 
CHAR(t1.UPDDT||' '||right('000000'||ltrim(t1.UPDTM),6)) BETWEEN CHAR('1'||VARCHAR_FORMAT(current date - 2 days,'yymmdd hh24:mi:ss'))  AND CHAR('1'||VARCHAR_FORMAT(current timestamp, 'yymmdd hh24:mi:ss'))
AND t2.ITCLS NOT LIKE 'Z%'

                
	
			
-- TRX TW
SELECT t1.HOUSE,t1.TCODE,t1.ORDNO,t1.ITNBR,t2.ITCLS, t1.UPDDT,t1.UPDTM,t1.TRQTY,t1.ENTUM,t1.VNDNR,t1.REFNO,t1.LLOCN,t1.NLLOC,t1.BATCH,t1.TRMID,
CHAR(t1.UPDDT||' '||right('000000'||ltrim(t1.UPDTM),6)) AS "TrxTime"

FROM AMFLIBL.IMHIST  t1, AMFLIBL.ITMRVA t2, AMFLIBL.WHSMST t3
WHERE t1.ITNBR=t2.ITNBR  AND t2.STID = t3.STID AND t1.HOUSE = t3.WHID AND t1.TRQTY > 0 AND t1.TCODE IN ('TW') AND  t1.NLLOC IN ('S01ST1','1A011') 
AND CHAR(t1.UPDDT||' '||right('000000'||ltrim(t1.UPDTM),6))  BETWEEN ? AND ?  





/*
CHAR(t1.UPDDT||' '||right('000000'||ltrim(t1.UPDTM),6)) BETWEEN CHAR('1'||VARCHAR_FORMAT(current date - 2 days,'yymmdd hh24:mi:ss'))  AND CHAR('1'||VARCHAR_FORMAT(current timestamp, 'yymmdd hh24:mi:ss'))
*/