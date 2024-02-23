SELECT *
FROM afi_atp
LIMIT 10`wh3
-- DB2 
SELECT t1.APITNB, t1.APHOUS, t2.iWeek, t2.QTY, t3.WD
FROM  AFILELIB.ATPSUM t1,
TABLE(
VALUES
('WK1',t1.APAT01),('WK2',t1.APAT02),('WK3',t1.APAT03),('WK4',t1.APAT04),('WK5',t1.APAT05),('WK6',t1.APAT06),('WK7',t1.APAT07),
('WK8',t1.APAT08),('WK9',t1.APAT09),('WK10',t1.APAT10),('WK11',t1.APAT11),('WK12',t1.APAT12),('WK13',t1.APAT13),('WK14',t1.APAT14),
('WK15',t1.APAT15),('WK16',t1.APAT16),('WK17',t1.APAT17),('WK18',t1.APAT18),('WK19',t1.APAT19),('WK20',t1.APAT20),('WK21',t1.APAT21),
('WK22',t1.APAT22),('WK23',t1.APAT23),('WK24',t1.APAT24),('WK25',t1.APAT25),('WK26',t1.APAT26),('WK27',t1.APAT27),('WK28',t1.APAT28),
('WK29',t1.APAT29),('WK30',t1.APAT30),('WK31',t1.APAT31),('WK32',t1.APAT32),('WK33',t1.APAT33),('WK34',t1.APAT34),('WK35',t1.APAT35),
('WK36',t1.APAT36),('WK37',t1.APAT37),('WK38',t1.APAT38),('WK39',t1.APAT39),('WK40',t1.APAT40),('WK41',t1.APAT41),('WK42',t1.APAT42),
('WK43',t1.APAT43)) AS t2(iWeek, QTY),
TABLE(
VALUES
(t1.APWK01),(t1.APWK02),(t1.APWK03),(t1.APWK04),(t1.APWK05),(t1.APWK06),(t1.APWK07),
(t1.APWK08),(t1.APWK09),(t1.APWK10),(t1.APWK11),(t1.APWK12),(t1.APWK13),(t1.APWK14),
(t1.APWK15),(t1.APWK16),(t1.APWK17),(t1.APWK18),(t1.APWK19),(t1.APWK20),(t1.APWK21),
(t1.APWK22),(t1.APWK23),(t1.APWK24),(t1.APWK25),(t1.APWK26),(t1.APWK27),(t1.APWK28),
(t1.APWK29),(t1.APWK30),(t1.APWK31),(t1.APWK32),(t1.APWK33),(t1.APWK34),(t1.APWK35),
(t1.APWK36),(t1.APWK37),(t1.APWK38),(t1.APWK39),(t1.APWK40),(t1.APWK41),(t1.APWK42),
(t1.APWK43)) AS t3(WD)

WHERE t1.APHOUS in ('335')




-- Mysql Union all 列转行

select t1.APITNB, t1.APHOUS, 't1.APAT01' Week, t1.APAT01 QTY
from afi_atp t1
union all
SELECT t1.APITNB, t1.APHOUS, 't1.APAT02' WEEK, t1.APAT02 QTY
FROM afi_atp t1
union all
SELECT t1.APITNB, t1.APHOUS, 't1.APAT03' WEEK, t1.APAT03 QTY
FROM afi_atp t1


-- Mysql crossjoin
select t1.APITNB, t1.APHOUS,
  c.col,
  case c.col
    when 't1.APAT01' then 'a'
    when 't1.APAT02' then 'b'
    when 't1.APAT03' then 'c'
  end as data
from afi_atp t1

cross join
(
  select 't1.APAT01' as col
  union all select 't1.APAT02'
  union all select 't1.APAT03'
) c





SELECT *
FROM 335_shipped

















