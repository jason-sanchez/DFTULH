select panum, COUNT(panum) from [075ChargeDt]

group by panum

select distinct panum from [075ChargeDt]
where hl7Generated is null
and paNum in ('999999013', '999999003')