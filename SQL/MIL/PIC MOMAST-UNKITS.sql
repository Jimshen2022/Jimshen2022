SELECT A.FITWH, A.REFNO, A.ORDNO, A.FITEM, A.FDESC, 
        A.ORQTY + A.QTDEV -A.QTYRC as MQTY,A.QTYRC, 
        Date(Substr(Char(A.ODUDT+ 19000000), 1, 4) || '-'||  Substr(Char(A.ODUDT + 19000000), 5, 2)|| '-' ||substr(Char(A.ODUDT + 19000000), 7, 2)) AS FG_DUE,
        A.OSTAT, A.JOBNO, A.ITCL 
        FROM AMFLIBL.MOMAST A 
        WHERE (A.FITWH='51') AND (substr(A.ORDNO,1,2)='MA') 
        AND (SUBSTR(A.JOBNO, 12, 1) NOT IN ('O','S','P')) 
        AND (A.OSTAT Not In ('99','45','55')) AND (A.ORQTY + A.QTDEV -A.QTYRC <>0)