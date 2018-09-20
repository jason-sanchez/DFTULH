SELECT     [075ChargeDt].recordID, [075ChargeDt].batchID, [075ChargeDt].paNum, [075ChargeDt].serviceDate, [075ChargeDt].serviceCode, [075ChargeDt].serviceQty, 
                      [075ChargeDt].procMod, [075ChargeDt].procMod2, [075ChargeDt].HL7Generated, [130Service].description
FROM         [075ChargeDt] INNER JOIN
                      [130Service] ON [075ChargeDt].serviceCode = [130Service].servCode
                      
 where paNum IN ('999999013', '999999003')                     