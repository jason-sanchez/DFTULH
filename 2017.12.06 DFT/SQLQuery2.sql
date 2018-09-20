select panum, COUNT(panum) from [075ChargeDt]
where paNum = '999999013'
group by panum

select distinct panum from [075ChargeDt]
where hl7Generated is null
and paNum = '999999013'