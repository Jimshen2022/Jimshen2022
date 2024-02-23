SELECT p.PITEM, MAX(p.PAMNT) as PAMNT 
FROM AFILELIBQ.PRICE p 
group by p.pitem 
order by p.pitem