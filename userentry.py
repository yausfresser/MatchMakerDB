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
import inspect
import sys

try:
    root = tk.Tk()
    # make regular expression to validate email address
    regex = '^[a-z0-9]+[\._]?[a-z0-9]+[@]\w+[.]\w{2,3}$'
    # setting the windows size
    root.geometry("800x400")
    root.title("Shidduch Data entry form")
    # todo: this is a test
    # declaring string variable
    # for storing name and password
    var_Firstname = tk.StringVar()
    var_Lastname=tk.StringVar()
    var_Age=tk.StringVar()
    var_Dob=tk.StringVar()
    var_Height=tk.StringVar()
    var_Address=tk.StringVar()
    var_Dropdown = tk.StringVar()
    var_Contactphone=tk.StringVar()
    var_Fathername=tk.StringVar()
    var_Mothername=tk.StringVar()
    var_ParentMartialStatus=tk.StringVar()
    var_Education=tk.StringVar()
    var_Volunteer=tk.StringVar()
    var_Summercamps=tk.StringVar()
    var_Siblings=tk.StringVar()
    var_ShulAttend=tk.StringVar()
    var_Rovofshul=tk.StringVar()
    var_FamilyReferences=tk.StringVar()
    var_Note=tk.StringVar()
    var_Photo=tk.StringVar()
    var_CurrentStatus=tk.StringVar()
    var_Email=tk.StringVar()
    var_FriendReferences=tk.StringVar()
    var_Resume=tk.StringVar()
    var_DropdownLearning=tk.StringVar()
    var_Gender=tk.StringVar()
    var_FileName=tk.StringVar()
    var_PhotoName=tk.StringVar()
    var_DocName=tk.StringVar()

    # defining a function that will
    # get the name and password and
    # print them on the screen

    # Function for opening the
    # file explorer window
    filename=""
    #adding browse button
    def browseFiles(value):
       if value=='photo':
            var_PhotoName.set(filedialog.askopenfilename(initialdir="/",
                                              title="Select a File",
                                              filetypes=(("Photo",
                                                          "*.jp*"),
                                                         ("all files",
                                                          "*.*"))))
            print(var_Photo)
       else:
            var_DocName.set(filedialog.askopenfilename(initialdir="/",
                                              title="Select a File",
                                              filetypes=(("Word Document",
                                                          "*.doc*"),
                                                         ("PDF",
                                                          "*.pdf*"),
                                                         ("all files",
                                                          "*.*"))))
            print(var_DocName)


    #function call to display message box
    def msgbox(msgtodisplay):
        messagebox.showinfo("Message Box", msgtodisplay)
    #returns file extension
    def getfileextension(path):
        split_tup = os.path.splitext(path)
        return split_tup[1]
    # function that submits data to save to db
    def submit():
        fname = var_Firstname.get()
        lname = var_Lastname.get()
        age=var_Age.get()
        gender=var_Gender.get()
        height = var_Height.get()
        parentstatus=var_ParentMartialStatus.get()
        fathername=var_Fathername.get()
        mothername=var_Mothername.get()
        learningstatus=var_DropdownLearning.get()
        contact=var_Contactphone.get()
        docname=var_DocName.get()
        photoname=var_PhotoName.get()
        # validate main fields contain data
        if fname=="" or lname=="" or age=="" or gender=="" or parentstatus=="" or fathername=="" or mothername=="" or learningstatus=="" or contact=="":
            msgbox("Enter in required fields before continuing. required fields are marked with *")
            # need to quit program
        else:
            # assure email address is valie format
            if ValidateEmail(var_Email.get()) == "false":
                msgbox("Invalid email format. Please reenter a valid email address need to end gracefully")
                email = "invalid email"

            else:
                #saving files into program directories
                if docname!="":
                    #TODO save file into program photo or document directory. use shiduch first and lastname as filename+ phonenumber. save stored path into local directory. if directory doesnt exist create it
                    #check if path exists
                    if os.path.exists(os.getcwd() + "\\documents")==False:
                        #create directory
                        os.mkdir(os.getcwd() + "\\documents")
                    print(docname)
                    docnewfilename =os.getcwd() + "\\documents\\"+ fname + lname + contact + getfileextension(docname)
                    print(docnewfilename)
                    os.rename(docname,docnewfilename)
                #saving files into program directories
                if photoname!="":
                    #TODO save file into program photo or document directory. use shiduch first and lastname as filename+ phonenumber. save stored path into local directory. if directory doesnt exist create it
                    #check if path exists
                    if os.path.exists(os.getcwd() + "\\Photos")==False:
                        #create directory
                        os.mkdir(os.getcwd() + "\\Photos")
                    print(photoname)
                    photonnewfilename =os.getcwd() + "\\Photos\\"+ fname + lname + contact + getfileextension(photoname)
                    print(photonnewfilename)
                    os.rename(photoname,photonnewfilename)

                email = var_Email.get()
                #building insert statement
                insertstatement= (
                     "Insert into contact (firstname,lastname,age,gender,martial_status,fathersname,mothername,contact,employment_status,email,resume_path,photo_path) "
                     "VALUES (?,?,?,?,?,?,?,?,?,?,?,?)"
                 )
                path = os.getcwd() + '\\' + "Shidduch.accdb"
                connectstr = 'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=%s;' % path
                conn = pyodbc.connect(connectstr)
                    #r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Temp\Shidduch.accdb;')
                cursor = conn.cursor()

                cursor.execute(insertstatement, (fname , lname ,age,gender,parentstatus,fathername ,mothername ,contact ,learningstatus ,email,docnewfilename,photonnewfilename ))
                conn.commit()





                #clear fields after submitting statement
                var_Firstname.set("")
                var_Lastname.set("")
                var_Age.set("")
                var_Height.set("")
    #this function connects to microsoft DB and adds in new record based on data user entered on data entry form
    def insertrecord(sqlstatement,value):

        conn = pyodbc.connect(
            r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\13474\Documents\Shidduch.accdb;')
        cursor = conn.cursor()
        # cursor.execute('''
        #             INSERT INTO contact (FirstName, LastName, Age, Dob, height)
        #             VALUES('Mike', 'Jordan',55,'01311988',51)
        #           ''')
        cursor.execute(sqlstatement,value)
        conn.commit()

    # validates text contains valie email address
    def ValidateEmail(email):
        # pass the regular expression
        # and the string in search() method
        if (re.search(regex, email)):
            return("true")

        else:
            return("false")
          #  exit()
        # creating a label for
    # name using widget Label
    Fname_label = tk.Label(root, text='FirstName*',
                          font=('calibre',
                               10, 'bold'))
    # creating a entry for input
    # name using widget Entry
    Fname_entry = tk.Entry(root,
                          textvariable=var_Firstname, font = ('calibre', 10, 'normal'))

    # creating a label for password
    Lname_label = tk.Label(root,
                           text='Lastname*',
                           font=('calibre', 10, 'bold'))

    # creating a entry for password
    Lname_entry = tk.Entry(root,
                           textvariable=var_Lastname,
                           font=('calibre', 10, 'normal'),
                           )
    Age_label = tk.Label(root,
                           text='Age*',
                           font=('calibre', 10, 'bold'))

    # creating a entry for password
    Age_entry = tk.Entry(root,
                           textvariable=var_Age,
                           font=('calibre', 10, 'normal'),
                           )
    Height_label = tk.Label(root,
                           text='Height',
                           font=('calibre', 10, 'bold'))

    # creating a entry for password
    Height_entry = tk.Entry(root,
                           textvariable=var_Height,
                           font=('calibre', 10, 'normal'),
                           )
    # creating a button using the widget
    # Button that will call the submit function
    sub_btn = tk.Button(root, text='Save',
                        command=submit)
    Gender_label = tk.Label(root,
                           text='Gender*',
                           font=('calibre', 10, 'bold'))
    #drop = tk.Combobox(root,var_Dropdown,'Male','Female')
    dropGender=ttk.Combobox(root, width=27,
                                textvariable=var_Gender)
    # Adding combobox drop down list

    dropGender['values'] = (' Male',
                              ' Female'
                          )
    Dob_label = tk.Label(root,
                           text='Dob',
                           font=('calibre', 10, 'bold'))
    Dob_entry = tk.Entry(root,
                           textvariable=var_Dob,
                           font=('calibre', 10, 'normal'),
                           )
    Address_label = tk.Label(root,
                           text='Address',
                           font=('calibre', 10, 'bold'))
    Address_entry = tk.Entry(root,
                           textvariable=var_Address,
                           font=('calibre', 10, 'normal'),
                           )
    Email_label = tk.Label(root,
                           text='Email*',
                           font=('calibre', 10, 'bold'))
    Email_entry = tk.Entry(root,
                           textvariable=var_Email,
                           font=('calibre', 10, 'normal'),
                           )
    Father_label = tk.Label(root,
                           text='Father*',
                           font=('calibre', 10, 'bold'))
    Father_entry = tk.Entry(root,
                           textvariable=var_Fathername,
                           font=('calibre', 10, 'normal'),
                           )
    Mother_label = tk.Label(root,
                           text='Mother*',
                           font=('calibre', 10, 'bold'))
    Mother_entry = tk.Entry(root,
                           textvariable=var_Mothername,
                           font=('calibre', 10, 'normal'),
                           )
    ParentMartialStatus_label = tk.Label(root,
                           text='Parent Martial Status*',
                           font=('calibre', 10, 'bold'))

    dropParent = ttk.Combobox(root, width=27,
                                textvariable=var_ParentMartialStatus)

    # Adding combobox drop down list

    dropParent['values'] = (' Married',
                              ' Seperated',
                              ' Father widowed',
                              ' Mother is Widow',
                              'Divorced',
                            'Both parents deceased'
                          )
    Note_label = tk.Label(root,
                           text='Note',
                           font=('calibre', 10, 'bold'))
    Note_entry = tk.Entry(root,
                           textvariable=var_Note,
                           font=('calibre', 10, 'normal'))
    Photo_label = tk.Label(root,
                           text='Attach Photo',
                           font=('calibre', 10, 'bold'))

    Photo_explore = Button(root,
                            text="Browse Photo",
                            command=lambda: browseFiles('photo'))

    Resume_label = tk.Label(root,
                           text='Attach Resume',
                           font=('calibre', 10, 'bold'))

    button_explore = Button(root,
                            text="Browse Document",
                            command=lambda: browseFiles('document'))

    CurrentStatus_label = tk.Label(root,
                           text='Current Working/Learning Status*',
                           font=('calibre', 10, 'bold'))
    drop = ttk.Combobox(root, width=27,
                                textvariable=var_DropdownLearning)

    # Adding combobox drop down list

    drop['values'] = (' Learning FT',
                              ' Learning PT',
                              ' Working/Koveah Itim',
                              ' Learning PT/Working PT',
                              ' Learning/College',
                              ' Working/College',
                              ' College/Working PT'
                    )
    ContactPhone_label = tk.Label(root,
                           text='Contact Phone number*',
                           font=('calibre', 10, 'bold'))
    ContactPhone_entry = tk.Entry(root,
                           textvariable=var_Contactphone,
                           font=('calibre', 10, 'normal'))

    button_exit = Button(root,
                         text="Exit",
                         command=exit)
    # placing the label and entry in
    # the required position using grid
    # method
    Fname_label.grid(row=0, column=0)
    Fname_entry.grid(row=0, column=1)
    Lname_label.grid(row=1, column=0)
    Lname_entry.grid(row=1, column=1)
    Age_label.grid(row=2, column=0)
    Age_entry.grid(row=2, column=1)
    Height_label.grid(row=3, column=0)
    Height_entry.grid(row=3, column=1)
    Gender_label.grid(row=4,column=0)
    #creates dropdown listbox
    drop.grid(row=4,column=1)
    dropGender.grid(row=4,column=1)
    Dob_label.grid(row=0, column=2)
    Dob_entry.grid(row=0, column=3)
    Address_label.grid(row=1, column=2)
    Address_entry.grid(row=1, column=3)
    Email_label.grid(row=2, column=2)
    Email_entry.grid(row=2, column=3)
    Father_label.grid(row=3, column=2)
    Father_entry.grid(row=3, column=3)
    Mother_label.grid(row=4, column=2)
    Mother_entry.grid(row=4, column=3)
    ParentMartialStatus_label.grid(row=5, column=0)
    dropParent.grid(row=5,column=1)
    Note_label.grid(row=9, column=0)
    Note_entry.grid(row=9, column=1)
    Resume_label.grid(row=9, column=2)
    button_explore.grid(row=9,column=3)
    #Resume_entry.grid(row=9, column=3)
    Photo_label.grid(row=10, column=0)
    Photo_explore.grid(row=10,column=1)
    #Photo_entry.grid(row=10, column=1)
    CurrentStatus_label.grid(row=11, column=0)
    drop.grid(row=11, column=1)
    ContactPhone_label.grid(row=12, column=0)
    ContactPhone_entry.grid(row=12, column=1)
    #button_explore.grid(column=15, row=0)
    button_exit.grid(row=15, column=0)
    drop.grid()
    sub_btn.grid(row=15, column=1)


    # performing an infinite loop
    # for the window to display
    root.mainloop()
except TypeError:
    print("Error invalid type")
except ValueError:
    print ("error invalid error")
except Exception as e:
    print("Error Occured")
    print("this is the error", e)