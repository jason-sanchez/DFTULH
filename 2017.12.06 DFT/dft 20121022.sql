SELECT     [075ChargeDt].serviceCode, [075ChargeDt].paNum, [075ChargeDt].serviceDate, [075ChargeDt].serviceQty, [130Service].description, [130Service].discipline, 
                      [130Service].department, [130Service].cpt
FROM         [075ChargeDt] INNER JOIN
                      [130Service] ON [075ChargeDt].serviceCode = [130Service].servCode
                      
                      where [075ChargeDt].serviceDate > '10/22/2012'