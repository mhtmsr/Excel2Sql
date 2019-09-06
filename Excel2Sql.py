''' Code By Mohit Misra<mhtmsr>  '''
import xlrd
#Path of the File is Given Here
loc=("sample.xls")
#Opening the WorkBook
wb = xlrd.open_workbook(loc) 
sheet = wb.sheet_by_index(0) 
#Extracting Column and Rows Counts
columnno=sheet.ncols
rowno=sheet.nrows
 #To Store the List containing all the Columns
lstcol=[]
#To Store the List containing value of one row 
lst=[]
 #To Store the Table Name
tablename="tablename"
#Extracting Column Names and Storing in Column list
for i in range(sheet.ncols): 
    lstcol.append(sheet.cell_value(0, i)) 
for i in range(1,rowno):
    lst[:]=[]
    for j in range(columnno):
        
        try:
            value=int(sheet.cell_value(i,j))
            lst.append(str(value))
        except ValueError:
                str1= f"'{sheet.cell_value(i,j)}'"
                lst.append(str1)
    sqlquery="insert into table "+tablename+"("+','.join(lstcol)+") values("+','.join(lst)+");"
    
    print(sqlquery)