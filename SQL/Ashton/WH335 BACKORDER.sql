-- 存档的
(SELECT 
concat(T1.NTRIP,'') TRIP#,concat(t2.BHWHS#,'') WH,t1.NDROP,t1.NORD#,t1.NCUSNO,t1.NSHPTO,CONCAT(t1.NITEM,'') MODEL,t1.NQTYUN as BO_Qty,
t1.NCODE,CONCAT(t1.NDATE,'') Date,CONCAT(t1.NTIME,'') Time,t1.NUSER as Supervisor,t1.NCUSNM,t2.BHTITM as Trip_Qty,t2.BHTCUB as trip_cubes,
t6.REDESC,t6.RESTAT,t1.NQTYUN/t2.BHTITM as COMPLETION,t3.Cubes*t1.NQTYUN as BO_Cubes,t4.MOHTQ as OnHand, 
DATE(SUBSTR(t1.NDATE,1,4)||'-'||SUBSTR(NDATE,5,2)||'-'||SUBSTR(NDATE,7,2)) as BKDate, 
week(DATE(SUBSTR(t1.NDATE,1,4)||'-'||SUBSTR(NDATE,5,2)||'-'||SUBSTR(NDATE,7,2))) as WK,
Month(DATE(SUBSTR(t1.NDATE,1,4)||'-'||SUBSTR(NDATE,5,2)||'-'||SUBSTR(NDATE,7,2))) as BKMonth,
Year(DATE(SUBSTR(t1.NDATE,1,4)||'-'||SUBSTR(NDATE,5,2)||'-'||SUBSTR(NDATE,7,2))) as BKYEAR

From ASHLEYARCH.BTRSNCDEA t1,ASHLEYARCH.BTTRIPH t2,DISTLIB.DWBOLRC t6,AFILELIB.ITMEXT t3,AMFLIBA.ITEMBL t4

Where t1.NCODE=t6.RECODE and t1.NTRIP=t2.BHTRP# and t1.NDATE 
BETWEEN  20210101 
AND int(substr(trim(char(CURRENT DATE)),1,4)||substr(trim(char(CURRENT DATE)),6,2)||substr(trim(char(CURRENT DATE)),9,2)) 
And T2.BHWHS#='335' and t1.NITEM=t3.ITNBR and t1.NITEM = t4.itnbr and t2.bhwhs# = t4.house

order by Date,trip#)

union

-- 当前的
(SELECT 
concat(T1.NTRIP,'') TRIP#,concat(t2.BHWHS#,'') WH,t1.NDROP,t1.NORD#,t1.NCUSNO,t1.NSHPTO,CONCAT(t1.NITEM,'') MODEL,t1.NQTYUN as BO_Qty,
t1.NCODE,CONCAT(t1.NDATE,'') Date,CONCAT(t1.NTIME,'') Time,t1.NUSER as Supervisor,t1.NCUSNM,t2.BHTITM as Trip_Qty,t2.BHTCUB as trip_cubes,
t6.REDESC,t6.RESTAT,t1.NQTYUN/t2.BHTITM as COMPLETION,t3.Cubes*t1.NQTYUN as BO_Cubes,t4.MOHTQ as OnHand, 
DATE(SUBSTR(t1.NDATE,1,4)||'-'||SUBSTR(NDATE,5,2)||'-'||SUBSTR(NDATE,7,2)) as BKDate, 
week(DATE(SUBSTR(t1.NDATE,1,4)||'-'||SUBSTR(NDATE,5,2)||'-'||SUBSTR(NDATE,7,2))) as WK,
Month(DATE(SUBSTR(t1.NDATE,1,4)||'-'||SUBSTR(NDATE,5,2)||'-'||SUBSTR(NDATE,7,2))) as BKMonth,
Year(DATE(SUBSTR(t1.NDATE,1,4)||'-'||SUBSTR(NDATE,5,2)||'-'||SUBSTR(NDATE,7,2))) as BKYEAR

From DISTLIB.BTRSNCDE t1,DISTLIB.BTTRIPH t2,DISTLIB.DWBOLRC t6,AFILELIB.ITMEXT t3,AMFLIBA.ITEMBL t4

Where t1.NCODE=t6.RECODE and t1.NTRIP=t2.BHTRP# and t1.NDATE 
BETWEEN  20210101 
AND int(substr(trim(char(CURRENT DATE)),1,4)||substr(trim(char(CURRENT DATE)),6,2)||substr(trim(char(CURRENT DATE)),9,2)) 
And T2.BHWHS#='335' and t1.NITEM=t3.ITNBR and t1.NITEM = t4.itnbr and t2.bhwhs# = t4.house

order by Date,trip#)