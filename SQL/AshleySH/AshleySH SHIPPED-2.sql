SELECT t4.INIVDT, t6.ITITNO, t4.INWHSE, t6.ITITCL, t6.ITSHQT, t2.UUCCIM, t7.XCUS#, t6.ITWHSE, t6.ITCSNO, 
t6.ITSPNO, t5.XNTRPN,t5.XNINVR, t5.XNORNO, t1.STACD, t2.CUBES, t4.INPONO, t1.CUSNM, t7.XFRGHT, t7.XTCONF, t7.XDSCNT, t6.ITPRIC,t1.CUSA3,
t1.cmacustomerclasscode,t1.cctyn,t1.cmacurrencycode

FROM AFILELIBQ.ACUSMASJ t1, AFILELIBQ.ITMEXT t2, AMFLIBQ.MBBZRES1 t3, AFILELIBQ.TSININA1 t4, AFILELIBQ.TSINXN t5, AFILELIBQ.TSITIN t6, AFILELIBQ.TSITXN t7

WHERE t4.INORNO = t7.XTORNO AND t7.XTORNO = t6.ITORNO AND t6.ITORNO = t5.XNORNO AND t4.ININVR = t7.XTINVR AND t7.XTINVR = t6.ITINVR AND t7.XTITNO = t6.ITITNO 
AND t6.ITCSNO = t1.CUSNO AND t6.ITITNO = t3.BZAITX AND t3.BZAITX = t2.ITNBR AND t4.INWHSE = t6.ITWHSE AND t6.ITINVR = t5.XNINVR AND t6.ITCSNO = t7.XCUS# 
AND t7.XCUS# = t4.INCSNO AND t7.XTITSQ = t6.ITITSQ AND t4.INRIDT = t7.XRIDAT AND t7.XRIDAT = t6.ITRIDT AND ((t4.INWHSE='232') 
AND (t4.INIVDT>= 20210101 And t4.INIVDT<= int(substr(trim(char(CURRENT DATE - 1 days)),1,4)||substr(trim(char(CURRENT DATE - 1 days)),6,2)||
substr(trim(char(CURRENT DATE - 1 days)),9,2))) AND (t6.ITSHQT>0))

ORDER BY t4.INIVDT,t5.XNTRPN