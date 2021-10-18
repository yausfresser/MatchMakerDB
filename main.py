import pymysql
import pyodbc
import os
from tkinter import *
from tkinter import messagebox
import tkinter as tk
import tabulate
from tkinter import simpledialog


def donothing():
    #filewin = Toplevel(root)
   # button = Button(filewin, text="Do nothing button")
  #  connecttodb()
  #  button.pack()
    tmpfilename=os.getcwd() + '\\'+ 'userentry.py'
    os.system(tmpfilename)
def DisplayMessagebox(message):
    messagebox.showinfo(message, message)
def DisplaySearchDialogue():
    tmpfilename = os.getcwd() + '\\' + 'Searchbox.py'
    os.system(tmpfilename)
def helloCallBack():
   msg = messagebox.showinfo( "Hello Python", "Hello World")

#connects to microsoft access database, returns all records in contact table
def connecttodb():
    # ServerName = r'pathtodb\\database.mdb'
    # connStr = 'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=%s;' % ServerName
    path = os.getcwd() + '\\' + "Shidduch.accdb"
    connectstr = 'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=%s;' % path
   #have user enter age
    agetosearch = simpledialog.askstring("Input", "age are you looking for",
                                    parent=root)
   # connectstr="Driver={Microsoft.Jet.OLEDB.4.0}; Data Source="+ path
    #print(connectstr)
    #conn = pyodbc.connect('Driver=Microsoft Access Driver (*.mdb, *.accdb)}; DBQ='""+ path)
    conn=pyodbc.connect(connectstr) #('Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Temp\Shidduch.accdb;')
    cursor = conn.cursor()
    sqlstatement='(select * from contact where age >= ?)'
    print(sqlstatement)
    cursor.execute(sqlstatement,agetosearch)

    for row in cursor.fetchall():
        messagebox.showinfo("test",row)

#this function connects to microsoft DB and adds in new record based on data user entered on data entry form
def insertrecord():
    conn = pyodbc.connect(
        r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\temp\Shidduch.accdb;')
    cursor = conn.cursor()
    cursor.execute('''
                INSERT INTO contact (FirstName, LastName, Age, Dob, height)
                VALUES('Mike', 'Jordan',55,'01311988',51)
              ''')
    conn.commit()
# Function to create menu
def createmenu():

    menubar = Menu(root)

    filemenu = Menu(menubar, tearoff=0)
    filemenu.add_command(label="Exit", command=root.quit)

    menubar.add_cascade(label="File", menu=filemenu)

    editmenu = Menu(menubar, tearoff=0)
    menubar.add_cascade(label="Record", menu=editmenu)
    editmenu.add_command(label="New", command=donothing)
    # editmenu.add_command(label="Edit", command=donothing)
    # editmenu.add_command(label="Delete", command=donothing)

    Searchmenu = Menu(menubar, tearoff=0)
    menubar.add_cascade(label="Search", menu=Searchmenu,command=DisplaySearchDialogue)
    Searchmenu.add_command(label="Age", command=DisplaySearchDialogue)

    root.config(menu=menubar)
    name_var = tk.StringVar()


try:
    # TODO: this is a test
    root = Tk()
    root.geometry('800x400')
    root.configure(background="grey");
    root.title("new shidduch entry")

    createmenu()
    root.mainloop()

except TypeError:
    print("Error invalid type")
except ValueError:
    print ("error invalid error")
except Exception as e:
    print("Error Occured")
    print("this is the error", e)


