/****** Script for SelectTopNRows command from SSMS  ******/
SELECT TOP 1000 [recordID]
      ,[batchID]
      ,[paNum]
      ,[serviceDate]
      ,[serviceCode]
      ,[serviceQty]
      ,[procMod]
      ,[snHdId]
      ,[credits]
      ,[procMod2]
  FROM [ITW].[dbo].[075ChargeDt]
  where panum = '12397964'
  and servicecode = 26401034