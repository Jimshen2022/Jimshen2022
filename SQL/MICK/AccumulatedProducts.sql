-- by product to accumulate the product as string, made by JimShen on Jul.18.2022
select a8.WCIORIGIN,a8.CONTAINER#,a8.CONTAINERTYPE,replace(replace(xml2clob(xmlagg(xmlelement(NAME A, a8.PRODUCT||','))),'<A>',''),'</A>','') as PRODUCT 
FROM
(
SELECT a9.WCIORIGIN,a9.CONTAINER#,a9.PRODUCT,a9.CONTAINERTYPE
FROM 
(...........)
) a8
GROUP BY a8.WCIORIGIN,a8.CONTAINER#,a8.CONTAINERTYPE