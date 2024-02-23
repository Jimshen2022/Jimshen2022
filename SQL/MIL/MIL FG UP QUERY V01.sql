select
t1.RPZ0D7 as "Site identifier",t1.RPAITX as "Site identifier",t1.RPZ0D9 as "Item number",t1.RPAENB as "Item revision",
t1.RPBRCD as "Company number",t1.RPBLDT as "Currency ID",t1.RPAMVA as "Price effective date",t1.RPAAJ7 as "Foreign currency price",
t1.RPELDT as "Pricing U/M",t1.RPALDT as "Price effective end date",t1.RPABTM as "Create date",t1.RPAFVN as "Create time",
t1.RPAMDT as "Created by user",t1.RPACTM as "Change date",t1.RPAHVN as "Change time",t1.RPAHVN as "Changed by user"
FROM AMFLIBL.ITMFPR t1
WHERE t1.RPZ0D7 IN ('51')



select *
FROM AMFLIBL.ITMFPR t1
Limit 100


