
-- MIL PRODUCTS SEPERATED -- updated on Jun.03.2022


(CASE 
        WHEN c.ITCLS LIKE 'TAF%' THEN 'RP'
        WHEN c.ITCLS IN ('PACS') THEN 'UnKits'
        WHEN c.ITCLS LIKE 'Z%' AND c.ITCLS LIKE '%K' THEN 'UnKits'
        WHEN c.ITCLS IN ('ZACM','ZASU','ZMLH','ZMLR','ZUSR','ZUSU','ZVUC','ZXUC','ZUSU','ZUMU','ZAMU','ZASM','ZASR','ZDMA','ZMUC','ZSUS','ZUMS','ZUSM','ZVMA','ZVUS','ZXLH','ZXLM','ZXLR','ZXMS','ZXMU') THEN 'UPH'
        WHEN c.ITCLS IN ('ZDAA','ZDAY','ZVAA','ZDAB','ZDAW','ZDYB','ZDBC','ZABC','ZECD') THEN 'CG'
        WHEN c.ITCLS IN ('ZKIS') THEN 'Bedding'		
        WHEN c.ITCLS IN ('WPLS') THEN 'Plastics'
        WHEN c.ITCLS IN ('WVBC','WVCS') THEN 'Foundation'		
        WHEN c.ITCLS IN ('PANL') THEN 'Panel'
        WHEN c.ITCLS IN ('ZKIZ') THEN 'ZipperCover'
        WHEN c.ITCLS IN ('BBFR','WVHC') THEN 'Verona'		
        WHEN c.ITCLS NOT LIKE 'Z%' THEN 'RawMaterial'
        ELSE 'Check' END) AS Product


(CASE 
        WHEN t3.ITCLS LIKE 'TAF%' THEN 'RP'
        WHEN t3.ITCLS IN ('PACS') THEN 'UnKits'
        WHEN t3.ITCLS LIKE 'Z%' AND t3.ITCLS LIKE '%K' THEN 'UnKits'
        WHEN t3.ITCLS IN ('ZACM','ZASU','ZMLH','ZMLR','ZUSR','ZUSU','ZVUC','ZXUC','ZUSU','ZUMU','ZAMU','ZASM','ZASR','ZDMA','ZMUC','ZSUS','ZUMS','ZUSM','ZVMA','ZVUS','ZXLH','ZXLM','ZXLR','ZXMS','ZXMU') THEN 'UPH'
        WHEN t3.ITCLS IN ('ZDAA','ZDAY','ZVAA','ZDAB','ZDAW','ZDYB','ZDBC','ZABC','ZECD') THEN 'CG'
        WHEN t3.ITCLS IN ('ZKIS') THEN 'Bedding'		
        WHEN t3.ITCLS IN ('WPLS') THEN 'Plastics'
        WHEN t3.ITCLS IN ('WVBC','WVCS') THEN 'Foundation'		
        WHEN t3.ITCLS IN ('PANL') THEN 'Panel'
        WHEN t3.ITCLS IN ('ZKIZ') THEN 'ZipperCover'
        WHEN t3.ITCLS IN ('BBFR','WVHC') THEN 'Verona'		
        WHEN t3.ITCLS NOT LIKE 'Z%' THEN 'Raw'
        ELSE 'Check' END) AS Product



----------------------------------------------------------------------------
SELECT t1.ITCLS,
(CASE 
        WHEN t1.ITCLS LIKE 'TAF%' THEN 'RP'
        WHEN t1.ITCLS IN ('PACS') THEN 'UnKits'
        WHEN t1.ITCLS LIKE 'Z%' AND t1.ITCLS LIKE '%K' THEN 'UnKits'
        WHEN t1.ITCLS IN ('ZACM','ZASU','ZMLH','ZMLR','ZUSR','ZUSU','ZVUC','ZXUC','ZUSU','ZUMU','ZAMU','ZASM','ZASR','ZDMA','ZMUC','ZSUS','ZUMS','ZUSM','ZVMA','ZVUS','ZXLH','ZXLM','ZXLR','ZXMS','ZXMU') THEN 'UPH'
        WHEN t1.ITCLS IN ('ZDAA','ZDAY','ZVAA','ZDAB','ZDAW','ZDYB','ZDBC','ZABC','ZECD') THEN 'CG'
        WHEN t1.ITCLS IN ('ZKIS') THEN 'Bedding'		
        WHEN t1.ITCLS IN ('WPLS') THEN 'Plastics'
        WHEN t1.ITCLS IN ('WVBC','WVCS') THEN 'Foundation'		
        WHEN t1.ITCLS IN ('PANL') THEN 'Panel'
        WHEN t1.ITCLS IN ('ZKIZ') THEN 'ZipperCover'
        WHEN t1.ITCLS IN ('BBFR','WVHC') THEN 'Verona'		
        WHEN t1.ITCLS NOT LIKE 'Z%' THEN 'Raw'
        ELSE 'Check' END) AS Product	
FROM AMFLIBL.ITMRVA t1
GROUP BY t1.ITCLS, (CASE 
        WHEN t1.ITCLS LIKE 'TAF%' THEN 'RP'
        WHEN t1.ITCLS IN ('PACS') THEN 'UnKits'
        WHEN t1.ITCLS LIKE 'Z%' AND t1.ITCLS LIKE '%K' THEN 'UnKits'
        WHEN t1.ITCLS IN ('ZACM','ZASU','ZMLH','ZMLR','ZUSR','ZUSU','ZVUC','ZXUC','ZUSU','ZUMU','ZAMU','ZASM','ZASR','ZDMA','ZMUC','ZSUS','ZUMS','ZUSM','ZVMA','ZVUS','ZXLH','ZXLM','ZXLR','ZXMS','ZXMU') THEN 'UPH'
        WHEN t1.ITCLS IN ('ZDAA','ZDAY','ZVAA','ZDAB','ZDAW','ZDYB','ZDBC','ZABC','ZECD') THEN 'CG'
        WHEN t1.ITCLS IN ('ZKIS') THEN 'Bedding'		
        WHEN t1.ITCLS IN ('WPLS') THEN 'Plastics'
        WHEN t1.ITCLS IN ('WVBC','WVCS') THEN 'Foundation'		
        WHEN t1.ITCLS IN ('PANL') THEN 'Panel'
        WHEN t1.ITCLS IN ('ZKIZ') THEN 'ZipperCover'
        WHEN t1.ITCLS IN ('BBFR','WVHC') THEN 'Verona'		
        WHEN t1.ITCLS NOT LIKE 'Z%' THEN 'Raw'
        ELSE 'Check' END)

-------------------------------------------------------------------------------------------------------------


(CASE 
        WHEN a1.ITCLS LIKE 'TAF%' THEN 'RP'
		WHEN a1.ITCLS IN ('PACS') THEN 'UnKits'
        WHEN a1.ITCLS LIKE 'Z%' AND a1.ITCLS LIKE '%K' THEN 'UnKits'
        WHEN a1.ITCLS IN ('ZACM','ZASU','ZMLH','ZMLR','ZUSR','ZUSU','ZVUC','ZXUC','ZUSU','ZUMU','ZAMU','ZASM','ZASR','ZDMA','ZMUC','ZSUS','ZUMS','ZUSM','ZVMA','ZVUS','ZXLH','ZXLM','ZXLR','ZXMS','ZXMU') THEN 'UPH'
        WHEN a1.ITCLS IN ('ZDAA','ZDAY','ZVAA','ZDAB','ZDAW','ZDYB','ZDBC','ZABC','ZECD') THEN 'CG'
        WHEN a1.ITCLS IN ('ZKIS') THEN 'Bedding'		
        WHEN a1.ITCLS IN ('WPLS') THEN 'Plastics'
        WHEN a1.ITCLS IN ('WVBC','WVCS') THEN 'Foundation'		
		WHEN a1.ITCLS IN ('PANL') THEN 'Panel'
		WHEN a1.ITCLS IN ('ZKIZ') THEN 'ZipperCover'
        WHEN a1.ITCLS IN ('BBFR','WVHC') THEN 'Verona'		
		WHEN a1.ITCLS NOT LIKE 'Z%' THEN 'Raw'
        ELSE 'Check' END) AS Product


(CASE 
        WHEN t1.ITCLS LIKE 'TAF%' THEN 'RP'
		WHEN t1.ITCLS IN ('PACS') THEN 'UnKits'
        WHEN t1.ITCLS LIKE 'Z%' AND t1.ITCLS LIKE '%K' THEN 'UnKits'
        WHEN t1.ITCLS IN ('ZACM','ZASU','ZMLH','ZMLR','ZUSR','ZUSU','ZVUC','ZXUC','ZUSU','ZUMU','ZAMU','ZASM','ZASR','ZDMA','ZMUC','ZSUS','ZUMS','ZUSM','ZVMA','ZVUS','ZXLH','ZXLM','ZXLR','ZXMS','ZXMU') THEN 'UPH'
        WHEN t1.ITCLS IN ('ZDAA','ZDAY','ZVAA','ZDAB','ZDAW','ZDYB','ZDBC','ZABC','ZECD') THEN 'CG'
        WHEN t1.ITCLS IN ('ZKIS') THEN 'Bedding'		
        WHEN t1.ITCLS IN ('WPLS') THEN 'Plastics'
        WHEN t1.ITCLS IN ('WVBC','WVCS') THEN 'Foundation'		
		WHEN t1.ITCLS IN ('PANL') THEN 'Panel'
		WHEN t1.ITCLS IN ('ZKIZ') THEN 'ZipperCover'
        WHEN t1.ITCLS IN ('BBFR','WVHC') THEN 'Verona'		
		WHEN t1.ITCLS NOT LIKE 'Z%' THEN 'Raw'
        ELSE 'Check' END) AS Product	


--PQ
= Table.AddColumn(#"Expanded ITEMS", "Product", 
each if Text.StartsWith([ITCLS], "TAF") then "RP" 
else if Text.StartsWith([ITCLS], "PACS") then "UnKits" 
else if Text.StartsWith([ITCLS], "Z") and Text.EndsWith([ITCLS], "K")  then "UnKits" 

else if [ITCLS]="ZACM" or [ITCLS]="ZASU" or [ITCLS]="ZMLH" or [ITCLS]="ZMLR" or [ITCLS]="ZUSR" 
or [ITCLS]="ZUSU" or [ITCLS]="ZVUC" or [ITCLS]="ZXUC" or [ITCLS]="ZUSU" or [ITCLS]="ZUMU" then "UPH"

else if [ITCLS]="ZDAA" or [ITCLS]="ZDAY" or [ITCLS]="ZVAA" or [ITCLS]="ZDAB" or [ITCLS]="ZDAW" 
or [ITCLS]="ZDYB" or [ITCLS]="ZDBC" or [ITCLS]="ZABC" or [ITCLS]="ZECD" then "CG"

else if [ITCLS] = "ZKIS" then "Bedding" 
else if [ITCLS] = "WPLS" then "Plastics" 
else if [ITCLS] = "WVBC" or [ITCLS] ="WVCS" then "Foundation" 
else if [ITCLS] = "PANL" then "Panel" 
else if [ITCLS] = "ZKIZ" then "ZipperCover" 
else if [ITCLS] = "BBFR" then "Verona" 
else if [ITCLS] = "WVHC" then "Verona" 		
else if Text.StartsWith([ITCLS],"Z") then "Raw"
else "CHECK")
		


(CASE 
        WHEN t3.ITCLS LIKE 'TAF%' THEN 'RP'
		WHEN t3.ITCLS IN ('PACS') THEN 'UnKits'
        WHEN t3.ITCLS LIKE 'Z%' AND t3.ITCLS LIKE '%K' THEN 'UnKits'
        WHEN t3.ITCLS IN ('ZACM','ZASU','ZMLH','ZMLR','ZUSR','ZUSU','ZVUC','ZXUC','ZUSU','ZUMU') THEN 'UPH'
        WHEN t3.ITCLS IN ('ZDAA','ZDAY','ZVAA','ZDAB','ZDAW','ZDYB') THEN 'CG'
        WHEN t3.ITCLS IN ('ZKIS') THEN 'Bedding'		
        WHEN t3.ITCLS IN ('WPLS') THEN 'Plastics'
        WHEN t3.ITCLS IN ('WVBC','WVCS') THEN 'Foundation'		
		WHEN t3.ITCLS IN ('PANL') THEN 'Panel'
		WHEN t3.ITCLS IN ('ZKIZ') THEN 'ZipperCover'
        WHEN t3.ITCLS IN ('BBFR','WVHC') THEN 'Verona'		
		WHEN t3.ITCLS NOT LIKE 'Z%' THEN 'RAW'
        ELSE 'Check' END) AS Product



















(CASE 
        WHEN t2.ITCLS IN ('SLDK') THEN 'RP'
        WHEN t2.ITCLS LIKE 'T%' THEN 'RP'
		WHEN t2.ITCLS LIKE 'R%' THEN 'RP'
		WHEN t2.ITCLS IN ('PACS') THEN 'UnKits'
        WHEN t2.ITCLS LIKE 'Z%' AND t2.ITCLS LIKE '%K' THEN 'UnKits'
        WHEN t2.ITCLS IN ('ZACM','ZASU','ZMLH','ZMLR','ZUSR','ZUSU','ZVUC','ZXUC','ZUSU','ZUMU') THEN 'UPH'
        WHEN t2.ITCLS IN ('ZDAA','ZDAY','ZVAA','ZDAB','ZDAW','ZDYB') THEN 'CG'
        WHEN t2.ITCLS IN ('ZKIS') THEN 'Bedding'		
        WHEN t2.ITCLS IN ('WPLS') THEN 'Plastics'
        WHEN t2.ITCLS IN ('WVBC','WVCS') THEN 'Foundation'		
		WHEN t2.ITCLS IN ('PANL') THEN 'Panel'
		WHEN t2.ITCLS IN ('ZKIZ') THEN 'ZipperCover'
        WHEN t2.ITCLS IN ('BBFR','WVHC') THEN 'Verona'		
		WHEN t2.ITCLS NOT LIKE 'Z%' THEN 'RawMaterial'
        ELSE 'Check' END) AS Product


(CASE 
        WHEN t3.ITCLS IN ('WPLS') THEN 'Plastics'
		WHEN t3.ITCLS IN ('PANL') THEN 'Panel'
        WHEN t3.ITCLS IN ('WVBC','WVCS') THEN 'Foundation'
        WHEN t3.ITCLS IN ('SLDK') THEN 'RP'
        WHEN t3.ITCLS LIKE 'T%' THEN 'RP'
		WHEN t3.ITCLS LIKE 'R%' THEN 'RP'
        WHEN t3.ITCLS IN ('ZKIS') THEN 'Bedding'
		WHEN t3.ITCLS IN ('ZKIZ') THEN 'ZipperCover'
        WHEN t3.ITCLS LIKE 'Z%' AND t3.ITCLS LIKE '%K' THEN 'UnKits'
		WHEN t3.ITCLS IN ('PACS') THEN 'UnKits'
        WHEN t3.ITCLS IN ('BBFR','WVHC') THEN 'Verona'
        WHEN t3.ITCLS IN ('ZDAA','ZDAY','ZVAA','ZDAB','ZDAW','ZDYB') THEN 'CG'
        WHEN t3.ITCLS IN ('ZACM','ZASU','ZMLH','ZMLR','ZUSR','ZUSU','ZVUC','ZXUC') THEN 'UPH'
        ELSE 'RawMaterial' END))








(CASE 
        WHEN y2.ITCLS IN ('SLDK') THEN 'RP'
        WHEN y2.ITCLS LIKE 'T%' THEN 'RP'
		WHEN y2.ITCLS LIKE 'R%' THEN 'RP'
		WHEN y2.ITCLS IN ('PACS') THEN 'UnKits'
        WHEN y2.ITCLS LIKE 'Z%' AND y2.ITCLS LIKE '%K' THEN 'UnKits'
        WHEN y2.ITCLS IN ('ZACM','ZASU','ZMLH','ZMLR','ZUSR','ZUSU','ZVUC','ZXUC','ZUSU','ZUMU') THEN 'UPH'
        WHEN y2.ITCLS IN ('ZDAA','ZDAY','ZVAA','ZDAB','ZDAW','ZDYB') THEN 'CG'
        WHEN y2.ITCLS IN ('ZKIS') THEN 'Bedding'		
        WHEN y2.ITCLS IN ('WPLS') THEN 'Plastics'
        WHEN y2.ITCLS IN ('WVBC','WVCS') THEN 'Foundation'		
		WHEN y2.ITCLS IN ('PANL') THEN 'Panel'
		WHEN y2.ITCLS IN ('ZKIZ') THEN 'ZipperCover'
        WHEN y2.ITCLS IN ('BBFR','WVHC') THEN 'Verona'		
		WHEN y2.ITCLS NOT LIKE 'Z%' THEN 'RawMaterial'
        ELSE 'Check' END) AS Product,


(CASE 
        WHEN c.ITCLS IN ('SLDK') THEN 'RP'
        WHEN c.ITCLS LIKE 'T%' THEN 'RP'
		WHEN c.ITCLS LIKE 'R%' THEN 'RP'
		WHEN c.ITCLS IN ('PACS') THEN 'UnKits'
        WHEN c.ITCLS LIKE 'Z%' AND c.ITCLS LIKE '%K' THEN 'UnKits'
        WHEN c.ITCLS IN ('ZACM','ZASU','ZMLH','ZMLR','ZUSR','ZUSU','ZVUC','ZXUC','ZUSU') THEN 'UPH'
        WHEN c.ITCLS IN ('ZDAA','ZDAY','ZVAA','ZDAB','ZDAW','ZDYB') THEN 'CG'
        WHEN c.ITCLS IN ('ZKIS') THEN 'Bedding'		
        WHEN c.ITCLS IN ('WPLS') THEN 'Plastics'
        WHEN c.ITCLS IN ('WVBC','WVCS') THEN 'Foundation'		
		WHEN c.ITCLS IN ('PANL') THEN 'Panel'
		WHEN c.ITCLS IN ('ZKIZ') THEN 'ZipperCover'
        WHEN c.ITCLS IN ('BBFR','WVHC') THEN 'Verona'		
		WHEN c.ITCLS NOT LIKE 'Z%' THEN 'RawMaterial'
        ELSE 'Check' END) AS Product
		
		
SELECT 
t1.RPLSERIALNUMBER,t1.RPLORDERNUMBER,t1.RPLORDERCREATEDATE,t1.RPLITEMNUMBER,t1.RPLORDERQUANTITY
,t1.RPLRMQUANTITY,t1.RPLPRINTQUANTITY,t1.RPPRINTTOTALNUMBER,t1.RPPRINTSEQUENCE,t1.RPLCONTAINERNUMBER
,t1.RPLLOADSTATUS,t1.RPLPRINTDATE,t1.RPLPRINTUSER,t1.RPLLOADDATE,t1.RPLLOADUSERNAME,
t1.RPLUNLOADDATE,t1.RPLUNLOADUSER,t2.ITCLS, 
(CASE 
	WHEN t2.ITCLS IN ('WPLS') THEN 'PLASTICS' 
	WHEN t2.ITCLS IN ('WVBC','WVHC') THEN 'FOUNDATION' 
	WHEN t2.ITCLS IN ('SLDK','QA','QB') THEN 'RP' 
	WHEN t2.ITCLS LIKE 'T%' THEN 'RP' 
	WHEN t2.ITCLS IN ('ZKIS') THEN 'BEDDING' 
	WHEN t2.ITCLS IN ('ZKIZ') THEN 'ZIPPERCOVER' 
	WHEN t2.ITCLS LIKE 'Z%' AND t2.ITCLS LIKE '%K' THEN 'UNKITS' 
	WHEN t2.ITCLS IN ('PACS') THEN 'UNKITS' 
	WHEN t2.ITCLS IN ('BBFR') THEN 'FR SOCK' 
	WHEN t2.ITCLS IN ('ZDAA','ZDAY','ZVAA','ZDAB','ZDAW','ZDYB') THEN 'CG' 
	WHEN t2.ITCLS LIKE 'Z%' THEN 'UPH' ELSE 'CHECK' END) AS PRODUCT 

FROM RGNFILL.PC228RPF t1, AMFLIBL.ITMRVA t2
WHERE t1.RPLITEMNUMBER = t2.ITNBR AND  t1.RPLPRINTDATE BETWEEN ? AND ?


(CASE 
        WHEN c.ITCLS IN ('SLDK') THEN 'RP'
        WHEN c.ITCLS LIKE 'T%' THEN 'RP'
		WHEN c.ITCLS LIKE 'R%' THEN 'RP'
		WHEN c.ITCLS IN ('PACS') THEN 'UnKits'
        WHEN c.ITCLS LIKE 'Z%' AND c.ITCLS LIKE '%K' THEN 'UnKits'
        WHEN c.ITCLS IN ('ZACM','ZASU','ZMLH','ZMLR','ZUSR','ZUSU','ZVUC','ZXUC') THEN 'UPH'
        WHEN c.ITCLS IN ('ZDAA','ZDAY','ZVAA','ZDAB','ZDAW','ZDYB') THEN 'CG'
        WHEN c.ITCLS IN ('ZKIS') THEN 'Bedding'		
        WHEN c.ITCLS IN ('WPLS') THEN 'Plastics'
        WHEN c.ITCLS IN ('WVBC','WVCS') THEN 'Foundation'		
		WHEN c.ITCLS IN ('PANL') THEN 'Panel'
		WHEN c.ITCLS IN ('ZKIZ') THEN 'ZipperCover'
        WHEN c.ITCLS IN ('BBFR','WVHC') THEN 'Verona'		
		WHEN c.ITCLS IN ('MIS','UHS','UEUS','BMIS','MZIP','PLST','USTP','UFRP','ELIT','MISN','BLBL','LBL','UEPM','BAGE','UFHW','UEFB','UBBP','BELI','CHF','BFR','CPM','HDWR','UFRH','AA4','BBGE','CHS','CQD','LMF','BFMR','UECD','WNWD','UUPS','WNPS','BRTH','CHM','UHTG','CUS','CFWB','CRVW','BNFR','CNEL','CMD','MCMF','WNWB','UFRK','CME','URRM','CKX','UMEC','UPDM','CRE','UESH','FRM','CSF','PKLN','UAAF','UFBW','UBBS','UTRA','BDWR','UFRC','PILW','CBDK','CNJ','UFMA','UFRM','CLL','CRFM','BGLU','CNF','UFMC','CNO','CNG','UEPP','CNP','BDES','PLW','CGLU','CKC','GLUE','SAFB','UEFM','CMW','UFRT','UF83','CNB','CHY','CBF','CSCB','CNK','CCU','CHZP','CHD','CNA','CNRL','FFRN','CND','CNH','CBSR','UERK','WVHD','CNL','UFTK','CPSK','CNI','SCSK','CRF','USPR','BEMP','CSDA','WNCS','BRCT','WNPU','BCLP','UEMS','BBBC','PRW','OA','UEPA','BPLN','BCLB','FFR','CFA','CR','UETR','FPUM','DECK','QA','PVN') THEN 'RawMaterial'
        ELSE 'Check' END) AS Product


-- MIL PRODUCTS SEPERATED -- updated on Aug.28.2021
(CASE 
        WHEN c.ITCLS IN ('WPLS') THEN 'Plastics'
		WHEN c.ITCLS IN ('PANL') THEN 'Panel'
		WHEN c.ITCLS IN ('DECK','QA') THEN 'RawMaterial'
        WHEN c.ITCLS IN ('WVBC','WVHC','WVCS') THEN 'Foundation'
        WHEN c.ITCLS IN ('SLDK') THEN 'RP'
        WHEN c.ITCLS LIKE 'T%' THEN 'RP'
        WHEN c.ITCLS IN ('ZKIS') THEN 'Bedding'
		WHEN c.ITCLS IN ('ZKIZ') THEN 'ZipperCover'
        WHEN c.ITCLS LIKE 'Z%' AND c.ITCLS LIKE '%K' THEN 'UnKits'
		WHEN c.ITCLS IN ('PACS') THEN 'UnKits'
        WHEN c.ITCLS IN ('BBFR') THEN 'Verona'
        WHEN c.ITCLS IN ('ZDAA','ZDAY','ZVAA','ZDAB','ZDAW','ZDYB') THEN 'CG'
        WHEN c.ITCLS LIKE 'Z%' THEN 'UPH'
        ELSE 'RP' END) AS Product

-- MIL stage report by building  -- updated on Aug.28.2021
(CASE 
        WHEN t1.ITCL IN ('WPLS') THEN 'Plastics'
		WHEN t1.ITCL IN ('PANL') THEN 'Panel'
		WHEN t1.ITCL IN ('DECK','QA') THEN 'RawMaterial'
        WHEN t1.ITCL IN ('WVBC','WVHC','WVCS') THEN 'Foundation'
        WHEN t1.ITCL IN ('SLDK') THEN 'RP'
        WHEN t1.ITCL LIKE 'T%' THEN 'RP'
        WHEN t1.ITCL IN ('ZKIS') THEN 'Bedding'
		WHEN t1.ITCL IN ('ZKIZ') THEN 'ZipperCover'
        WHEN t1.ITCL LIKE 'Z%' AND t1.ITCL LIKE '%K' THEN 'UnKits'
		WHEN t1.ITCL IN ('PACS') THEN 'UnKits'
        WHEN t1.ITCL IN ('BBFR') THEN 'Verona'
        WHEN t1.ITCL IN ('ZDAA','ZDAY','ZVAA','ZDAB','ZDAW','ZDYB') THEN 'CG'
        WHEN t1.ITCL LIKE 'Z%' THEN 'UPH'
        ELSE 'RP' END) AS Product


-- MIL PRODUCTS SEPERATED -- updated on Aug.28.2021
(CASE 
        WHEN t3.ITCLS IN ('WPLS') THEN 'Plastics'
		WHEN t3.ITCLS IN ('PANL') THEN 'Panel'
		WHEN t3.ITCLS IN ('DECK','QA') THEN 'RawMaterial'
        WHEN t3.ITCLS IN ('WVBC','WVHC','WVCS') THEN 'Foundation'
        WHEN t3.ITCLS IN ('SLDK') THEN 'RP'
        WHEN t3.ITCLS LIKE 'T%' THEN 'RP'
        WHEN t3.ITCLS IN ('ZKIS') THEN 'Bedding'
		WHEN t3.ITCLS IN ('ZKIZ') THEN 'ZipperCover'
        WHEN t3.ITCLS LIKE 'Z%' AND t3.ITCLS LIKE '%K' THEN 'UnKits'
		WHEN t3.ITCLS IN ('PACS') THEN 'UnKits'
        WHEN t3.ITCLS IN ('BBFR') THEN 'Verona'
        WHEN t3.ITCLS IN ('ZDAA','ZDAY','ZVAA','ZDAB','ZDAW','ZDYB') THEN 'CG'
        WHEN t3.ITCLS LIKE 'Z%' THEN 'UPH'
        ELSE 'RP' END) AS Product


-- MIL PRODUCTS SEPERATED -- updated on Aug.28.2021
(CASE 
        WHEN t2.ITCLS IN ('WPLS') THEN 'Plastics'
		WHEN t2.ITCLS IN ('PANL') THEN 'Panel'
		WHEN t2.ITCLS IN ('DECK','QA') THEN 'RawMaterial'
        WHEN t2.ITCLS IN ('WVBC','WVHC','WVCS') THEN 'Foundation'
        WHEN t2.ITCLS IN ('SLDK') THEN 'RP'
        WHEN t2.ITCLS LIKE 'T%' THEN 'RP'
        WHEN t2.ITCLS IN ('ZKIS') THEN 'Bedding'
		WHEN t2.ITCLS IN ('ZKIZ') THEN 'ZipperCover'
        WHEN t2.ITCLS LIKE 'Z%' AND t2.ITCLS LIKE '%K' THEN 'UnKits'
		WHEN t2.ITCLS IN ('PACS') THEN 'UnKits'
        WHEN t2.ITCLS IN ('BBFR') THEN 'Verona'
        WHEN t2.ITCLS IN ('ZDAA','ZDAY','ZVAA','ZDAB','ZDAW','ZDYB') THEN 'CG'
        WHEN t2.ITCLS LIKE 'Z%' THEN 'UPH'
        ELSE 'RP' END) AS Product
