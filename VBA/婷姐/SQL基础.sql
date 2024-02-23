Select  t1.BDTRP#,t1.BDCUS#,t1.BDORD#,t1.BDITM#,t1.BDITMD,t1.BDINVN,t1.BDCTL#,t1.BDICLS,t1.BDCCLS,t1.BDITQT,t1.BDITCT,t1.BDITWT,t1.BDCTIM,
t1.BDSHPNO,t2.BHTRP#,t2.BHWHS#,t2.BHTRPS,t2.BHPRVS,t2.BHCNTI,t2.BHCNTN,t2.BHSEL1,t2.BHLUSR,t2.BHLDAT,t2.BHLTIM,t2.BHLTYP,T2.BHRDAT, 
t2.BHRTIM,t2.BHZDAT,T2.BHZTIM,T2.BHTSNS

from DISTLIBQ.BTTRIPD t1 full join DISTLIBQ.BTTRIPH t2 on t1.BDTRP# = t2.BHTRP# 

where t2.BHWHS# IN ('232') and t2.BHRDAT between int(substr(trim(char(CURRENT DATE - 15 days)),1,4)||substr(trim(char(CURRENT DATE- 15 days)),6,2)||substr(trim(char(CURRENT DATE- 15 days)),9,2)) 
AND int(substr(trim(char(CURRENT DATE)),1,4)||substr(trim(char(CURRENT DATE)),6,2)||substr(trim(char(CURRENT DATE)),9,2))

order by t2.BHLDAT,t1.BDTRP#,t1.BDITM#