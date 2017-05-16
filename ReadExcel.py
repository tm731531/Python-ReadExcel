import xlrd
import pymssql
data = xlrd.open_workbook('EXCEL檔案位置')
conn = pymssql.connect("DB位置", "帳號", "密碼", "DB")
cursor = conn.cursor()
table = data.sheets()[0]   
nrows = table.nrows 
for rownum in range(1,nrows):
    colnames =  table.row_values(rownum) 
    cursor.execute("  Insert into [dbo].[TB名稱] values (N'%s',N'%s',N'%s',N'%s',N'%s',N'%s',N'%s',N'%s')"
                  %(colnames[0],colnames[1],colnames[2],colnames[3],colnames[4],colnames[5],colnames[6],colnames[7]))
    print(colnames[0])
conn.commit()
