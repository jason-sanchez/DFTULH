/****** Script for SelectTopNRows command from SSMS  ******/
SELECT TOP 1000 [batchID]
      ,[transmissionBatch]
      ,[patientStatus]
      ,[terminalID]
      ,[batchDate]
      ,[numRecords]
      ,[numTransactions]
      ,[TransCreated]
      ,[batchFac]
      ,[departmentId]
      ,[billingStatus]
  FROM [ITW].[dbo].[075ChargeHd]
  where batchID = 46819