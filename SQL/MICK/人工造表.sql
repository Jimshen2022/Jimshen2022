-- 人工造时间表

with temp (date) as (
select date('01.01.2021') as date from sysibm.sysdummy1
union all
select date + 1 day from temp
where date < date('31.12.2021'))
select * from temp
	
	
/*
以上代码显示结果如下：
DATE
2/23/2016
2/24/2016
2/25/2016
2/26/2016
2/27/2016
2/28/2016
2/29/2016
3/1/2016
3/2/2016


	*/
	
	
-- 取周别	
with temp (date) as 
( 
select date('01.01.2021') as date 
from sysibm.sysdummy1  
union all
select date + 1 day 
from temp
where date <CURRENT DATE
)

select distinct Week(Date) DATE
from temp
group by date