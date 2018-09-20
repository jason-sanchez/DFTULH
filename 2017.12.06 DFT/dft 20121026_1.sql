select panum, COUNT(panum) from [075ChargeDt]
where paNum in ('999999013', '999999003')
group by panum

select
 c.batchID, 
c.recordID,
c.paNum,
c.serviceCode, 
c.serviceDate, 
c.serviceQty, 
c.procMod, 
c.procMod2, 
c.paNum ,
c.HL7Generated,
 s.[description], 
 e.epnum,
 e.Fname,
 e.Lname,
 e.Mname,
 e.DOB,
 e.AdminDate,
 e.SocSec 
from [075ChargeDt] c, [130Service] s, [001Episode] e
where c.paNum = '999999013'
and c.serviceCode = s.servcode

