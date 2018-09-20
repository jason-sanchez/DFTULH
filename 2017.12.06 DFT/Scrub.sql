select e.epnum, d.recordID as exRecordID, d.panum as exPANum, s.cpt as exCPT, 
        convert(nvarchar,d.serviceDate,101) as exServiceDate, d2.panum as PANum, 
        convert(nvarchar,d2.serviceDate,101) as serviceDate, s2.cpt as CPT, d2.recordID as recordID, 
        ex.excludeAction, ex.mutuallyExclusive, d.procMod, d.procMod2, s.[id] as snPlanID 
        from [075chargeDt] d 
        left join [075chargeDT] d2 on d.panum = d2.panum 
        and convert(nvarchar,d.serviceDate,101) = convert(nvarchar,d2.serviceDate,101) 
        and d.recordid <> d2.recordid 
        and (d2.serviceQty + (-1 * ISNULL(d2.credits, 0))) > 0 
        and d2.serviceCode <> '20001095' 
        inner join [007snPlan] s on d.snhdid = s.snhdid 
        and d.serviceCode = s.serviceCode 
        left join [007snPlan] s2 on d2.snhdid = s2.snhdid 
        and d2.serviceCode = s2.serviceCode 
        inner join [001episode] e on d.panum = e.panum 
        inner join [113payor] p on e.primaryIns = p.planCode 
        inner join [130cptExclusion] ex on left(s.cpt, 5) = ex.cptExclude 
        and left(s2.cpt, 5) = ex.cpt 
        and ex.inactive = 0 
        
        where e.status in ('OA') 
        
        and (p.medicare = 1 
        
        or p.medicaid = 1) 
        and e.intakeFacility in (200,300)
        and e.testCase = 0
        and (d.serviceQty + (-1 * ISNULL(d.credits, 0))) > 0 
        and (d.batchID = 0 or d.batchID is null) 
        and d.panum > '' 
        
        and d.serviceCode <> '20001095' 
        and (ex.excludeAction = 0 
        or (ex.excludeAction = 1 and d.procMod is null) 
        or (ex.excludeAction = 1 and d.procMod2 is null and d.procMod <> '59'))
        order by d.panum, s.cpt, exServiceDate 