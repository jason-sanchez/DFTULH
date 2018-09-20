SELECT     [075ChargeDt].batchID, [075ChargeDt].recordID, [075ChargeDt].serviceDate, [075ChargeDt].serviceQty, [075ChargeDt].procMod, [075ChargeDt].procMod2, 
                      [130Service].description, [075ChargeDt].serviceCode, [075ChargeDt].paNum
FROM         [001Episode] INNER JOIN
                      [075ChargeDt] ON [001Episode].PANum = [075ChargeDt].paNum INNER JOIN
                      [130Service] ON [075ChargeDt].serviceCode = [130Service].servCode
WHERE     ([001Episode].PANum = '999999013')
and [075ChargeDt].HL7Generated is null
order by serviceDate
