
#参数  有很多SQL语句用单行来写并不是很方便，所以你也可以使用三引号的字符串来写：
import pyodbc
cnxn = pyodbc.connect('DSN=AFIPROD; PWD=MJ2062')
cursor =cnxn.cursor()
cursor.execute("""
    SELECT T1.ITNBR, T1.HOUSE, T1.ITCLS, T1.MOHTQ, T1.WHSLC, T1.QTSYR, T2.ITDSC 
    FROM AMFLIBA.ITEMBL T1, AMFLIBA.ITMRVA T2, AMFLIBA.WHSMST T3 
    WHERE T2.ITCLS = T1.ITCLS AND T2.ITNBR = T1.ITNBR AND T2.STID = T3.STID AND T1.HOUSE = T3.WHID AND T1.HOUSE='335' AND T1.MOHTQ<>0
    """)
row = cursor.fetchone()
print('name:', row[1])  # access by column index
print('name:', row.ITNBR)  # or access by name





import pyodbc
cnxn = pyodbc.connect('DSN=AFIPROD; PWD=MJ2062')
cursor =cnxn.cursor()
cursor.execute("""
    SELECT T1.ITNBR, T1.HOUSE, T1.ITCLS, T1.MOHTQ, T1.WHSLC, T1.QTSYR, T2.ITDSC 
    FROM AMFLIBA.ITEMBL T1, AMFLIBA.ITMRVA T2, AMFLIBA.WHSMST T3 
    WHERE T2.ITCLS = T1.ITCLS AND T2.ITNBR = T1.ITNBR AND T2.STID = T3.STID AND T1.HOUSE = T3.WHID AND T1.HOUSE='335' AND T1.MOHTQ<>0
     AND T1.ITNBR = 'D372-124'
    """)
row = cursor.fetchone()
print('name:', row[1])  # access by column index
print('name:', row.ITNBR)  # or access by name




import pyodbc
cnxn = pyodbc.connect('DSN=AFIPROD; PWD=MJ2062')
cursor =cnxn.cursor()
cursor.execute("""
    SELECT T1.ITNBR, T1.HOUSE, T1.ITCLS, T1.MOHTQ, T1.WHSLC, T1.QTSYR, T2.ITDSC 
    FROM AMFLIBA.ITEMBL T1, AMFLIBA.ITMRVA T2, AMFLIBA.WHSMST T3 
    WHERE T2.ITCLS = T1.ITCLS AND T2.ITNBR = T1.ITNBR AND T2.STID = T3.STID AND T1.HOUSE = T3.WHID AND T1.HOUSE='335' AND T1.MOHTQ<>0
     AND T1.ITNBR = ? """,'D372-124')
row = cursor.fetchone()
print('name:', row[1])  # access by column index
print('name:', row.ITNBR)  # or access by name



import pyodbc
cnxn = pyodbc.connect('DSN=AFIPROD; PWD=MJ2062')
cursor =cnxn.cursor()
cursor.execute("""
    SELECT T1.ITNBR, T1.HOUSE, T1.ITCLS, T1.MOHTQ, T1.WHSLC, T1.QTSYR, T2.ITDSC 
    FROM AMFLIBA.ITEMBL T1, AMFLIBA.ITMRVA T2, AMFLIBA.WHSMST T3 
    WHERE T2.ITCLS = T1.ITCLS AND T2.ITNBR = T1.ITNBR AND T2.STID = T3.STID AND T1.HOUSE = T3.WHID AND T1.HOUSE='335' AND T1.MOHTQ<>0
     AND T1.ITNBR = ? and T1.MOHTQ = ? """,['1040225',1])
row = cursor.fetchone()
print('name:', row[1])  # access by column index
print('name:', row.ITNBR)  # or access by name
print('name:', row.MOHTQ)  # or access by name





