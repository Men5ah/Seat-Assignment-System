import tkinter as tk
from tkinter import *
from tkinter import ttk
from tkinter import messagebox
import pyodbc
import math

#================================================HELPFUL FUNCTIONS======================================================
#Clearing students sata section.
def clearData():
    identityEntry.delete(0, 'end')
    First_Name_entry.delete(0, 'end')
    Last_Name_entry.delete(0, 'end')
    age_entry.delete(0, 'end')

#Updating the treeview table
def Update(rows):
    tree.delete(*tree.get_children())
    for row in rows:
        display = []
        for item in row:
            display.append(str(item))
        tree.insert("", "end", values= display)

#Treeview Clearing
def Clear():
    searchEntry.delete(0, 'end')
    query = "SELECT * from Students"
    cursor.execute(query)
    rows = cursor.fetchall()
    Update(rows)

def getRows(event):
    rowid = tree.identify_row(event.y)
    item = tree.item(tree.focus())
    identityEntry.insert(0, item['values'][0])
    First_Name_entry.insert(0, item['values'][1])
    Last_Name_entry.insert(0, item['values'][2])
    age_entry.insert(0, item['values'][5])
    grade.set(item['values'][3])
    hostel.set(item['values'][6])
    gender.set(item['values'][4])

#=========================================================ASSIGNING STUDENTS TO TABLES==================================
#The Assign function runs the whole program, assigning students to tables in the MPH
def Assign():
    conn = pyodbc.connect(
        r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};'
        r'DBQ=C:/Users/HP/source/repos/Criterion C/Student Database.accdb;')
    cursor = conn.cursor()

    # Calculating the number of tables, round up using ceiling function
    maxTables = 0
    cursor.execute("select count(*) from Students")
    results = cursor.fetchall()
    for row in results:
        maximum = row[0] / 8
        maxTables = int(math.ceil(maximum))

    # Delete statement allows for running the program without encountering any integrity errors due to duplication of records
    cursor.execute("DELETE  * FROM TableGroups")
    conn.commit()

    # First loop in order to assign group numbers to all existing records on the database
    genderIndex = ["F", "M"]
    tableNumber = 1
    for grade in range(7, 13):
        for gender in genderIndex:
            sql = "SELECT ID from Students WHERE Grade = " + str(grade) + " AND Gender = '" + gender + "'"
            cursor.execute(sql)
            results = cursor.fetchall()
            for row in results:
                tableNumber = tableNumber + 1
                if tableNumber > maxTables:
                    tableNumber = 1
                cursor.execute(
                    "INSERT INTO TableGroups (ID, GroupNumber, Grade) VALUES(" + str(row[0]) + "," + str(tableNumber) + "," + str(grade) + ")")
    conn.commit()

    done_lbl = tk.Label(root, text="")
    done_lbl.grid(row=9, column=2, padx=2, pady=2)

    file_write = open('groups.txt', 'w')
    for GNumber in range(1, maxTables + 1):
        file_write.write('\n')
        file_write.write("TABLE NUMBER" + " " + str(GNumber) + " " + '\n')
        sqlStmt = "SELECT S.FirstName, S.LastName, S.Grade FROM Students S INNER JOIN TableGroups T ON S.ID = T.ID "
        sqlStmt = sqlStmt + " WHERE T.GroupNumber = " + str(GNumber)
        cursor.execute(sqlStmt)
        results = cursor.fetchall()

        count = 1
        for row in results:
            file_write.write(" " + str(count) + ". ")
            file_write.write(row[0] + " " + row[1] + " " + "(Grade " + str(row[2]) + ")" + '\n')
            done_lbl.config(text="Tables Assigned Successfully!", fg="green")
            count = count+1
    file_write.close()
#==============================================================SEARCHING FOR STUDENTS===================================
#Searching the treeview table/database
#ID, FirstName, LastName, Grade, Gender, Age, HostelCode
def Search():
    user = user_input.get()
    query = "SELECT * FROM Students WHERE FirstName LIKE '%" + user + "%' OR LastName LIKE '%" + user + "%' OR Grade LIKE '%" + user + "%'" \
        "OR Gender LIKE '%" + user + "%' OR HostelCode LIKE '%" + user + "%'"
    cursor.execute(query)
    rows = cursor.fetchall()
    Update(rows)
#Database Connection
conn = pyodbc.connect(
    r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=C:/Users/HP/source/repos/Criterion C/Student Database.accdb;')
cursor = conn.cursor()

root = tk.Tk()
user_input = StringVar()

#=================================================LABELS FOR EACH SECTION OF THE PROGRAM================================

#Making the labels for each section of the window
label1 = LabelFrame(root, text="Student List")
label1.pack(fill="both", expand="yes", padx=20, pady=10)
label2 = LabelFrame(root, text="Search")
label2.pack(fill="both", expand="yes", padx=20, pady=10)
label3 = LabelFrame(root, text="Student Data")
label3.pack(fill="both", expand="yes", padx=20, pady=10)

#================================================DISPLAY TABLE==========================================================

#Treeview table
tree = ttk.Treeview(label1, columns=(1, 2, 3, 4, 5, 6, 7), show='headings')
tree.column(1, width=100)
tree.column(2, width=100)
tree.column(3, width=100)
tree.column(4, width=100)
tree.column(5, width=100)
tree.column(6, width=100)
tree.column(7, width=100)

#Scrollbar
tree_scroll = Scrollbar(label1)
tree_scroll.pack(side = RIGHT, fill = Y)
#configure scrollbar
tree_scroll.config(command = tree.yview)

# Making headings
#ID, FirstName, LastName, Grade, Gender, Age, HostelCode
tree.heading(1, text="Student ID", anchor=CENTER)
tree.heading(2, text="First Name")
tree.heading(3, text="Last Name")
tree.heading(4, text="Grade", anchor=CENTER)
tree.heading(5, text="Gender", anchor=CENTER)
tree.heading(6, text="Age", anchor=CENTER)
tree.heading(7, text="Hostel Code", anchor=CENTER)
tree.pack()
tree.bind('<Double 1>', getRows)

tree_fetch = "SELECT * FROM Students"
cursor.execute(tree_fetch)
rows = cursor.fetchall()
conn.commit()
Update(rows)

#======================================DELETING STUDENT RECORDS=========================================================

#PARAMETERS ARE ID, FirstName, LastName, Grade, Gender, Age, HostelCode
def deleteStudent():
    studentID = identityEntry.get()
    if messagebox.askyesno("Confirm Delete?", "Are you sure you want to delete this record?"):
        #delete = "DELETE FROM Students WHERE ID = " + studentID
        delete = "DELETE FROM Students WHERE ID LIKE '%" + str(studentID) + "%'"
        cursor.execute(delete)
        conn.commit()
        Clear()
    else:
        return True

#===================================================ADDING NEW STUDENTS=================================================

def addStudent(FN, LN, Age, Grade, Gender, HC):
    cursor = conn.cursor()
    cursor.execute('SELECT * from Students')
    sql = "INSERT INTO Students (FirstName, LastName, Grade, Gender, Age, HostelCode) "
    sql = sql + "VALUES('" + FN + "','" + LN + "','" + str(Grade) + "','" + str(Gender) + "','" + str(
        Age) + "','" + str(HC) + "')"
    cursor.execute(sql)
    Clear()
    conn.commit()

def addNew():
    try:
        int(age_entry.get())
    except ValueError:
        warning = messagebox.showerror("Error", "Age Should be a Number.")
        if warning == 0:
            root.destroy()
        else:
            return True
    else:
        confirmation = messagebox.askyesno("Are you sure?", "Are you sure you want to add this student's record?")
        if confirmation == 1:
            FName = First_Name_entry.get()
            LName = Last_Name_entry.get()
            Age = age_entry.get()
            Grade = grade.get()
            Gender = gender.get()
            HCode = hostel.get()
            addStudent(FName, LName, Age, Grade, Gender, HCode)
            success_lbl.config(text="Student Successfully Added!", fg='black')
        else:
            return True
#============================================UPDATING STUDENT DATA======================================================

#Updating Students Details
def Update_Student():
    studentID = identityEntry.get()
    FirstName = First_Name_entry.get()
    LastName = Last_Name_entry.get()
    Grade = int(grade.get())
    Gender = str(gender.get())
    Age = int(age_entry.get())
    Hostel = str(hostel.get())
    print(studentID, FirstName, LastName, Grade, Gender, Age, Hostel)

    if messagebox.askyesno("Confirm Please", "Are you sure you want to update this student's record?"):
        query = "UPDATE Students SET FirstName = '" + FirstName + "'," + " LastName = '" + LastName + "'," + \
                " Grade = " + str(Grade) + "," + " Gender = '" + Gender + "'," + " Age = " + str(Age) + "," + \
                " HostelCode = '" + Hostel + "' WHERE ID = " + studentID
        print(query)
        cursor.execute(query)
        conn.commit()
        Clear()
    else:
        return True

#-----------------------------------------------------------------------------------------------------------------------
GradeArray = [7, 8, 9, 10, 11, 12]
GenderArray = ["F", "M"]
HostelArray = ['AN', 'FR', 'CA', 'CE']
#-----------------------------------------------------------------------------------------------------------------------
grade = StringVar()
grade.set(GradeArray[0])
gender = StringVar()
gender.set(GenderArray[0])
hostel = StringVar()
hostel.set(HostelArray[0])
#================================================ENTER======DATA=======SECTION==========================================

# Student Data Entries
identity = tk.Label(label3, text="Student ID")
identity.grid(row=1, column=1)
identityEntry = tk.Entry(label3)
identityEntry.grid(row=1, column=2)

fn = tk.Label(label3, text="First Name:")
fn.grid(row=2, column=1, padx=2, pady=2)
First_Name_entry = tk.Entry(label3)
First_Name_entry.grid(row=2, column=2, padx=2, pady=2)

ln = tk.Label(label3, text="Last Name:")
ln.grid(row=3, column=1, padx=2, pady=2)
Last_Name_entry = tk.Entry(label3)
Last_Name_entry.grid(row=3, column=2, padx=2, pady=2)

age = tk.Label(label3, text="Age:")
age.grid(row=4, column=1, padx=2, pady=2)
age_entry = tk.Entry(label3)
age_entry.grid(row=4, column=2, padx=2, pady=2)

gender_lbl = tk.Label(label3, text="Gender:")
gender_lbl.grid(row=5, column=1, padx=2, pady=2)
gender_drop = tk.OptionMenu(label3, gender, *GenderArray)
gender_drop.config(width=10)
gender_drop.grid(row=5, column=2, padx=2, pady=2)

gl = tk.Label(label3, text="Grade Level:")
gl.grid(row=6, column=1, padx=2, pady=2)
grade_level_drop = tk.OptionMenu(label3, grade, *GradeArray)
grade_level_drop.config(width=10)
grade_level_drop.grid(row=6, column=2, padx=2, pady=2)

hostelcode_lbl = tk.Label(label3, text="Hostel:")
hostelcode_lbl.grid(row=7, column=1, padx=2, pady=2)
hostel_drop = tk.OptionMenu(label3, hostel, *HostelArray)
hostel_drop.config(width=10)
hostel_drop.grid(row=7, column=2, padx=2, pady=2)

searchEntry = tk.Entry(label2, textvariable=user_input)
searchEntry.grid(row=1, column=1, padx=20)
Search_button = tk.Button(label2, text="Search", command=Search)
Search_button.grid(row=1, column=2, padx=2)
clearSearch = tk.Button(label2, text="Clear Search", command=Clear)
clearSearch.grid(row=1, column=3, padx=10)

#Successful Addition Notice
success_lbl = tk.Label(label3, text="")
success_lbl.grid(row=10, column=2, padx=2, pady=2)

#Various Buttons
#Add adds a new record to the database
Add_button = tk.Button(label3, text="Add", command=addNew)
Add_button.grid(row=8, column=1, padx=2, pady=20)

#A button that updates a record
Update_Button = tk.Button(label3, text="Update", command=Update_Student)
Update_Button.grid(row=8, column=2, pady=20)

# Add an remove button that will run the delete function of the database
remove_button = tk.Button(label3, text="Remove", command=deleteStudent)
remove_button.grid(row=8, column=3, pady=20)

#This is the assign button, it does the assigns the tables
AssignButton = tk.Button(label3, text = "Assign", command = Assign)
AssignButton.grid(row=9, column=2)

#This clears the student data area
clearData = tk.Button(label3, text="Clear Data", command = clearData)
clearData.grid(row=9, column=1)

#Adding an exit button
exit_button = tk.Button(label3, text='Exit', command=root.quit)
exit_button.grid(row=9, column=3, padx=2, pady=2)

#============================================MAIN========WINDOW=========================================================
root.title("Seat Assigner Application")
root.iconbitmap('C:/Users/HP/Desktop/IBDP/CS/IA/tis_black.ico')
root.geometry("800x700")
conn = pyodbc.connect(
    r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=C:/Users/HP/source/repos/Criterion C/Student Database.accdb;')
cursor = conn.cursor()

root.mainloop()
