import tkinter
from tkinter import *
m=tkinter.Tk()
src=StringVar()
dest=StringVar()
tname=StringVar()
def browsefunc():
    from tkinter import filedialog
    filename = filedialog.askopenfilename()
    global src
    return src.set(str(filename))
def file_save():
    from tkinter import filedialog
    filename = filedialog.asksaveasfilename(defaultextension=".txt")
    global dest
    return dest.set(str(filename))
def setTname():
    global tname
    return tname.set(str(tname.get()))

def generatefunc():
    ''' Code By Mohit Misra<mhtmsr>  '''
    import xlrd
    import os
    global src
    global dest
    #Path of the File is Given Here
    from pathlib import Path
    loc=(Path(str(src.get())))
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
    tablename=pathlabel2.get()

    f= open(Path(str(dest.get())),"w+")

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
        f.write(sqlquery+"\n")
              
    
    f.close()
    os.startfile(Path(str(dest.get())))


m.geometry("600x500")
m.configure(background="Blue")
m.resizable(width=False, height=False)
m.title("Excel2SQL Generator")
menubar=Menu(m)
m.config(menu=menubar)
filemenu=Menu(menubar, tearoff=0)
filemenu.add_separator()
filemenu.add_command(label="Exit", command=m.destroy)
menubar.add_cascade(label="File", menu=filemenu)
mainbox = Label(m, text="Excel2Sql Generator v1.0 by Mohit Misra<mhtmsr>\n<Generate Insert Statements Sql Queries from Excel Files>\nFeautres:Automatic Quotes Based on Type in Insert Statements",fg="White",bg="Black", font="30", bd="5", height=6, width="50", relief="raised", cursor="arrow")
mainbox.place(relx=0.52, rely=0.2, anchor=CENTER)
################################################
tname = Label(m, text="Table Name::>", font="30", bd="5",bg="Black",fg="White")
tname.place(relx=0.19, rely=0.39, anchor=CENTER)
pathlabel2 = Entry(m,fg="Black",relief="raised",width="50")
pathlabel2.place(relx=0.6, rely=0.39,anchor=CENTER)

#################################################
addfile = Label(m, text="ADD EXCEL FILE::->", font="30", bd="5",bg="Black",fg="White")
addfile.place(relx=0.19, rely=0.5, anchor=CENTER)
browsebutton = Button(m, text="Browse",fg="White",bg="Black", command=browsefunc)
browsebutton.place(relx=0.9, rely=0.5, anchor=CENTER)
pathlabel = Entry(m,textvar=src,fg="Black",relief="raised",width="50")
pathlabel.place(relx=0.6, rely=0.5,anchor=CENTER)
generatebutton = Button(m, text="GENERATE SQL FILE",fg="White",bg="Black",height="5",width="30", command=lambda:generatefunc())
generatebutton.place(relx=0.5, rely=0.85, anchor=CENTER)
##################################################
destfile = Label(m, text="DESTINATION::->", font="30", bd="5",bg="Black",fg="White")
destfile.place(relx=0.19, rely=0.65, anchor=CENTER)
browsebutton1 = Button(m, text="Browse",fg="White",bg="Black", command=lambda:file_save())
browsebutton1.place(relx=0.9, rely=0.65, anchor=CENTER)
pathlabel1 = Entry(m,textvar=dest,fg="Black",relief="raised",width="50")
pathlabel1.place(relx=0.6, rely=0.65,anchor=CENTER)

################################################################################3
m.mainloop()
