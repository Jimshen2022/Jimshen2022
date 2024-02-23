-- BI MIL NEW ITEM
-- putaway calss is null and pickput id is null
	(Select t1.ITNBR, t1.HOUSE, t1.MOHTQ, t1.WHSLC, t1.ITCLS, t1.QTSYR, t1.B2Z95S, 
	t1.ITDSC, t1.TIHIUNLD, t1.PICKPUT, t1.ITMCLSID, t1.UNITSWIDE, t1.UNITLAYERS, t1.UNITSDEEP, 
	t1.SCOOPQTY, t1.SKIDSIZE, t1.QTYCR, t1.NBSEAT, t1.CRTWIN, t1.CRTLIN, t1.CRTHIN, t1.PRDWIN, 
	t1.PRDHIN, t1.PRDLIN, t1.ITMWEGHT,t2.Open_CO_Qty,
	(CASE 
        WHEN t1.ITCLS IN ('WPLS') THEN 'Plastics'
        WHEN t1.ITCLS IN ('WVBC','WVHC') THEN 'Foundation'
        WHEN t1.ITCLS IN ('SLDK') THEN 'RP'
        WHEN t1.ITCLS LIKE 'T%' THEN 'RP'
        WHEN t1.ITCLS IN ('ZKIS') THEN 'Bedding'
		WHEN t1.ITCLS IN ('ZKIZ') THEN 'ZipperCover'
        WHEN t1.ITCLS LIKE 'Z%' AND t1.ITCLS LIKE '%K' THEN 'UnKits'
		WHEN t1.ITCLS IN ('PACS') THEN 'UnKits'
        WHEN t1.ITCLS IN ('BBFR') THEN 'Verona'
        WHEN t1.ITCLS IN ('ZDAA','ZDAY','ZVAA','ZDAB','ZDAW','ZDYB') THEN 'CG'
        WHEN t1.ITCLS LIKE 'Z%' THEN 'UPH'
        ELSE 'Others' END) AS Product

	from 
	(SELECT ITMEXT.ITNBR as Item#,ITEMBL.ITNBR, ITEMBL.HOUSE, ITEMBL.MOHTQ, ITEMBL.WHSLC, ITEMBL.ITCLS, 
	ITEMBL.QTSYR, ITMRVA.B2Z95S, ITMRVA.ITDSC, ITBEXT.TIHIUNLD, ITBEXT.PICKPUT, ITBEXT.ITMCLSID, ITBEXT.UNITSWIDE, 
	ITBEXT.UNITLAYERS, ITBEXT.UNITSDEEP, ITBEXT.SCOOPQTY, ITBEXT.SKIDSIZE, ITMEXT.QTYCR, ITMEXT.NBSEAT, ITMEXT.CRTWIN, 
	ITMEXT.CRTLIN, ITMEXT.CRTHIN, ITMEXT.PRDWIN, ITMEXT.PRDHIN, ITMEXT.PRDLIN, ITMEXT.ITMWEGHT
	FROM AFILELIBL.ITBEXT ITBEXT, AMFLIBL.ITEMBL ITEMBL, AFILELIBL.ITMEXT ITMEXT, AMFLIBL.ITMRVA ITMRVA
	WHERE ITBEXT.ITNBR = ITMRVA.ITNBR AND ITEMBL.HOUSE = ITBEXT.HOUSE AND ITEMBL.ITNBR = ITBEXT.ITNBR 
	AND ITEMBL.ITNBR = ITMRVA.ITNBR AND ITEMBL.ITCLS = ITMRVA.ITCLS AND ITMEXT.ITNBR = ITBEXT.ITNBR 
	AND ITMEXT.ITNBR = ITEMBL.ITNBR AND ITMEXT.ITNBR = ITMRVA.ITNBR AND ITMRVA.STID = ITBEXT.HOUSE 
	AND (ITBEXT.HOUSE IN ('51')) AND (ITEMBL.ITCLS like 'Z%' and ITEMBL.ITCLS not like '%K')  
	AND ITMRVA.ITNBR NOT LIKE 'A%') as T1, 

	(SELECT CDA3CD as Wanek,CDAITX as item,CDGLCD as ItemClass, sum(CDAGNV) as Open_CO_Qty  
	FROM AMFLIBL.MBCDRESM MBCDRESM  
	WHERE CDAGNV >0 and CDGLCD like 'Z%' and CDGLCD not like '%K' 
	Group by  CDA3CD,CDAITX,CDGLCD) AS T2 

	where T1.Item# = T2.item and T1.HOUSE=T2.Wanek and T1.ITCLS = T2.ItemClass and (t1.ITMCLSID= '' or t1.PICKPUT = '')
	order by t1.itnbr)

union

-- CG Tihi is 0
	(Select t1.ITNBR, t1.HOUSE, t1.MOHTQ, t1.WHSLC, t1.ITCLS, t1.QTSYR, t1.B2Z95S, 
	t1.ITDSC, t1.TIHIUNLD, t1.PICKPUT, t1.ITMCLSID, t1.UNITSWIDE, t1.UNITLAYERS, t1.UNITSDEEP, 
	t1.SCOOPQTY, t1.SKIDSIZE, t1.QTYCR, t1.NBSEAT, t1.CRTWIN, t1.CRTLIN, t1.CRTHIN, t1.PRDWIN, 
	t1.PRDHIN, t1.PRDLIN, t1.ITMWEGHT,t2.Open_CO_Qty,
	(CASE 
        WHEN t1.ITCLS IN ('WPLS') THEN 'Plastics'
        WHEN t1.ITCLS IN ('WVBC','WVHC') THEN 'Foundation'
        WHEN t1.ITCLS IN ('SLDK') THEN 'RP'
        WHEN t1.ITCLS LIKE 'T%' THEN 'RP'
        WHEN t1.ITCLS IN ('ZKIS') THEN 'Bedding'
		WHEN t1.ITCLS IN ('ZKIZ') THEN 'ZipperCover'
        WHEN t1.ITCLS LIKE 'Z%' AND t1.ITCLS LIKE '%K' THEN 'UnKits'
		WHEN t1.ITCLS IN ('PACS') THEN 'UnKits'
        WHEN t1.ITCLS IN ('BBFR') THEN 'Verona'
        WHEN t1.ITCLS IN ('ZDAA','ZDAY','ZVAA','ZDAB','ZDAW','ZDYB') THEN 'CG'
        WHEN t1.ITCLS LIKE 'Z%' THEN 'UPH'
        ELSE 'Others' END) AS Product
	from 
	(SELECT ITMEXT.ITNBR as Item#,ITEMBL.ITNBR, ITEMBL.HOUSE, ITEMBL.MOHTQ, ITEMBL.WHSLC, ITEMBL.ITCLS, 
	ITEMBL.QTSYR, ITMRVA.B2Z95S, ITMRVA.ITDSC, ITBEXT.TIHIUNLD, ITBEXT.PICKPUT, ITBEXT.ITMCLSID, ITBEXT.UNITSWIDE, 
	ITBEXT.UNITLAYERS, ITBEXT.UNITSDEEP, ITBEXT.SCOOPQTY, ITBEXT.SKIDSIZE, ITMEXT.QTYCR, ITMEXT.NBSEAT, ITMEXT.CRTWIN, 
	ITMEXT.CRTLIN, ITMEXT.CRTHIN, ITMEXT.PRDWIN, ITMEXT.PRDHIN, ITMEXT.PRDLIN, ITMEXT.ITMWEGHT
	FROM AFILELIBL.ITBEXT ITBEXT, AMFLIBL.ITEMBL ITEMBL, AFILELIBL.ITMEXT ITMEXT, AMFLIBL.ITMRVA ITMRVA
	WHERE ITBEXT.ITNBR = ITMRVA.ITNBR AND ITEMBL.HOUSE = ITBEXT.HOUSE AND ITEMBL.ITNBR = ITBEXT.ITNBR 
	AND ITEMBL.ITNBR = ITMRVA.ITNBR AND ITEMBL.ITCLS = ITMRVA.ITCLS AND ITMEXT.ITNBR = ITBEXT.ITNBR 
	AND ITMEXT.ITNBR = ITEMBL.ITNBR AND ITMEXT.ITNBR = ITMRVA.ITNBR AND ITMRVA.STID = ITBEXT.HOUSE 
	AND (ITBEXT.HOUSE IN ('51')) AND (ITEMBL.ITCLS like 'Z%' and ITEMBL.ITCLS not like '%K')  
	AND ITMRVA.ITNBR NOT LIKE 'A%') as T1, 

	(SELECT CDA3CD as Wanek,CDAITX as item,CDGLCD as ItemClass, sum(CDAGNV) as Open_CO_Qty  
	FROM AMFLIBL.MBCDRESM MBCDRESM  
	WHERE CDAGNV >0 and CDGLCD like 'Z%' and CDGLCD not like '%K' 
	Group by  CDA3CD,CDAITX,CDGLCD) AS T2 

	where T1.Item# = T2.item and T1.HOUSE=T2.Wanek and T1.ITCLS = T2.ItemClass and 
	 ((t1.itnbr like 'E%' and t1.UNITSWIDE =0) or (t1.itnbr like 'E%' and t1.UNITLAYERS =0) or (t1.itnbr like 'E%' and t1.SCOOPQTY =0)
	 or (t1.itnbr like 'E%' and t1.SKIDSIZE =0))	
	order by t1.itnbr)
	