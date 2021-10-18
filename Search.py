import tkinter as tk
from tkinter import ttk
from tkinter import *

# import filedialog module
from tkinter import filedialog
import re
from tkinter import messagebox
import pymysql
import pyodbc
import os


global rows_count
global tree
#create searchbox to
try:
    root = tk.Tk()
    root.geometry('800x600')
    root.configure(background="grey");
    root.title("Search dialogue")

    var_age = tk.StringVar()
    var_gender = tk.StringVar()
    var_JobStatus = tk.StringVar()
    var_radioOption=tk.StringVar()
    var_Treeview=tk.StringVar()
    # setup variables for search result

    def searchrecords():


        age=var_age.get()
        jobstatus=var_JobStatus.get()
        gender=var_gender.get()
        option=var_radioOption.get()
       # messagebox.showinfo("message", "option selected=" + option)
       # messagebox.showinfo("message", age+jobstatus+gender)
        #BUILD sql query
        strWhereStatement="select firstname,lastname,age,gender,height,employment_status,contact,email,fathersname,mothername,martial_status from contact"
        if age!="":
            if option=="0": # use the greater than or equal sign
                strWhereStatement+=" where age>=?"
            elif option=="1": #use less than or equal to
                strWhereStatement += " where age<=?"
            else:
                strWhereStatement += " where age=?"
        if jobstatus !="":
            if age!="":
                strWhereStatement+= " and employment_status=?"
            else:
                strWhereStatement+= " WHERE employment_status=?"
        if gender!="":
            if age!="" or jobstatus!="":
                strWhereStatement+=" and gender=?"
            else:
                strWhereStatement += "WHERE gender=?"

       # messagebox.showinfo("message", strWhereStatement)
        path = "C:\\Users\\13474\\PycharmProjects\\FinalPythonProject\\Shidduch.accdb"
        connectstr = 'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=%s;' % path
        conn = pyodbc.connect(connectstr)  # ('Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Temp\Shidduch.accdb;')
        cursor = conn.cursor()
        # sqlstatement = '(select * from contact where age >= ?)'
        # print(sqlstatement)
        # parameters to pass in
        if age!="" and gender!="" and jobstatus!="":
            rows_count = cursor.execute(strWhereStatement, age,jobstatus,gender)
        elif age!="" and gender!="":
            rows_count = cursor.execute(strWhereStatement, age, gender)
        elif age!="" and jobstatus!="":
            rows_count = cursor.execute(strWhereStatement, age, jobstatus)
        elif jobstatus!="" and gender!="":
            rows_count = cursor.execute(strWhereStatement, jobstatus,gender)
        elif jobstatus!="" and gender!="":
            rows_count = cursor.execute(strWhereStatement, jobstatus,gender)
        elif age!="":
            rows_count = cursor.execute(strWhereStatement, age)

        elif jobstatus!="":
            rows_count= cursor.execute(strWhereStatement, jobstatus)
        elif gender!="":
            rows_count= cursor.execute(strWhereStatement, gender)
        elif gender=="" and age=="" and jobstatus=="":
            rows_count= cursor.execute(strWhereStatement)

        rows=cursor.fetchall()
        if len(rows)==0:
            messagebox.showinfo("No records", "No records found")
        else:
            for row in rows:
                messagebox.showinfo("test", row)
                # messagebox.showinfo("test", row[0])
                # messagebox.showinfo("test", row[1])
                #tree.insert("", tk.END, values=row)#this puts data into one line but messesses up the formatting
                tree.insert("", tk.END, values=(row[0], row[1], row[2],row[3], row[4], row[5],row[6], row[7], row[8],row[9], row[10])) #this puts data into treeview but handles data formatting
                print(row)
        # if rows_count > 0:
        #     for row in cursor.fetchall():
        #         messagebox.showinfo("test", row)
        # else:
        #



    age_label = tk.Label(root, text='age',
                           font=('calibre',
                                 10, 'bold'))
    age_entry = tk.Entry(root,
                           textvariable=var_age, font=('calibre', 10, 'normal'))
    radiob1 = tk.Radiobutton(root, text="Greater or Equal to", variable=var_radioOption, value=0)
    radiob1.deselect()
    #radiob1.pack()


    radio2 = tk.Radiobutton(root, text="Less than or Equal to", variable=var_radioOption, value=1)
    #radio2.pack()
    radio2.deselect()
    gender_label = tk.Label(root, text='gender',
                           font=('calibre',
                                 10, 'bold'))

    dropParent = ttk.Combobox(root, width=27,
                              textvariable=var_gender)
    # Adding combobox drop down list
    dropParent['values'] = (' Male',
                            ' Female',
                            )
    type_label = tk.Label(root, text='type',
                           font=('calibre',
                                 10, 'bold'))
    drop = ttk.Combobox(root, width=27,
                                    textvariable=var_JobStatus)

        # Adding combobox drop down list

    drop['values'] = (' Learning FT',
                                  ' Learning PT',
                                  ' Working/Koveah Itim',
                                  ' Learning PT/Working PT',
                                  ' Learning/College',
                                  ' Working/College',
                                  ' College/Working PT'
                        )
    sub_btn = tk.Button(root, text='Search',
                            command=searchrecords)
    button_exit = Button(root,
                         text="Exit",
                         command=exit)
    #loading tree view

    tree = ttk.Treeview(root, columns=('firstName', 'lastName','Age','Gender','Height','employment_status','contact','email','fathersname','mothername','Martial_Status'))

    # Set the heading (Attribute Names)
    tree.heading('#1', text='FirstName')
    tree.heading('#2', text='LastName')
    tree.heading('#3', text='Age')
    tree.heading('#4', text='Gender')
    tree.heading('#5', text='Height')
    tree.heading('#6', text='Employment Status')
    tree.heading('#7', text='Contact')
    tree.heading('#8', text='Email')
    tree.heading('#9', text='Fathers Name')
    tree.heading('#10', text='Mothers Name')
    tree.heading('#11', text='Martial Status')
    # Specify attributes of the columns (We want to stretch it!)
    tree.column('#1', stretch=tk.YES,minwidth=0,width=100)
    tree.column('#2', stretch=tk.YES,minwidth=0,width=100)
    tree.column('#3', stretch=tk.YES,minwidth=0,width=100)
    tree.column('#4', stretch=tk.YES,minwidth=0,width=100)
    tree.column('#5', stretch=tk.YES,minwidth=0,width=100)
    tree.column('#6', stretch=tk.YES,minwidth=0,width=100)
    tree.column('#7', stretch=tk.YES,minwidth=0,width=100)
    tree.column('#8', stretch=tk.YES,minwidth=0,width=100)
    tree.column('#9', stretch=tk.YES,minwidth=0,width=100)
    tree.column('#10', stretch=tk.YES,minwidth=0,width=100)
    tree.column('#11', stretch=tk.YES,minwidth=0,width=100)


    tree.grid(row=12, columnspan=15)
    age_label.grid(row=0, column=0)
    age_entry.grid(row=0, column=1)
    radiob1.grid(row=3,column=0)
    radio2.grid(row=3,column=1)
    gender_label.grid(row=2, column=0)
    dropParent.grid(row=2, column=1)
    type_label.grid(row=1,column=0)
    drop.grid(row=1,column=1)
    sub_btn.grid(row=4,column=0)
    button_exit.grid(row=4,column=1)

    root.mainloop()

except TypeError:
    print("Error invalid type")
except ValueError:
    print("error invalid error")
except Exception as e:
    print("Error Occured")
    print("this is the error", e)