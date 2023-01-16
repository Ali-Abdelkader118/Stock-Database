import sqlite3
import openpyxl
from tkinter import *
import tkinter.ttk as ttk
from barcode import Code39
from barcode.writer import ImageWriter

#Creating The Database File And opening It
db = sqlite3.connect("Stock.db")
c = db.cursor()

files = ['A', 'B', 'C', 'D', 'E']

def Search_data():
    code = input('Input Your Code: ').upper()
    print(code)
    if len(code) != 6 :
        print("Code Incorrect")
    else:
        if code[0] in files  :
            l = code[0]
            d = c.execute(f"SELECT * FROM {l} WHERE Barcode=?",(code,),).fetchall()
            print(d)
        else:
            print("code Not Found")



def Search_Name():
    name = input("Enter Product_Name : ")
    for i in files:
        c.execute(f"SELECT * FROM {i} WHERE ")

def Create_Barcode():
    files = ['A', 'B', 'C', 'D', 'E']
    #Looping All File Names
    for i in files:
        #Getting Data From The Database
        c.execute(f'''select Barcode , Product_Name from {i}''')
        data = c.fetchall()
        #Looping For Creating The Barcode
        for i in data:
            print(r"ID : ", i[0])
            print(r"Name :", i[1])
            print("\n")
            #Creating The Barcode
            my_code = Code39(i[0], writer=ImageWriter())
            #Saving The Barcode As Image
            my_code.save(f"Barcodes/{i[0]}")



def Excel_Data():
    #Getting The User Input For The File Name And Number Of Rows 
    file = input("Enter File Name: ").upper()
    n_rows = int(input("Enter Number Of Rows: "))
    #Opening the Excel file
    wb = openpyxl.load_workbook(f"Data/{file}.xlsx")
    ws = wb.active
    ID = 1
    #Creating The Table If Not Exist
    c.execute(f"DROP TABLE IF EXISTS {file}")
    c.execute(f"CREATE TABLE {file}(Barcode TEXT NOT NULL UNIQUE, Product_Name TEXT NOT NULL , QTY TEXT NOT NULL)")
    while ID <= n_rows:
        # Getting The Data From Excel
        Barcode = ws[f"A{ID}"].value
        Product_Name = ws[f"B{ID}"].value
        n = 1
        #Inserting The Data Into Sqlite Database
        c.execute(f"INSERT INTO {file}(Barcode , Product_Name, QTY) values(?, ?, ?)",(Barcode , Product_Name, n))
        #Printing The Final Result
        print("Data Inserted in the table: ")
        data=c.execute(f'''SELECT * FROM {file}''')
        for row in data:
            print(row)
        #increasing the ID Count to Switch to next Row
        ID = ID + 1
    db.commit()



def Add_item():
    Item_ID = input("Enter Item ID: ").upper()
    n = int(input("Enter Added Number: "))

    if len(Item_ID) > 6 or len(Item_ID) < 6 and type(Item_ID) != str() and Item_ID[0] in files:
        print("Code Incorrect")
    else:
            l = Item_ID[0]
            d = c.execute(f"SELECT QTY FROM {l} WHERE Barcode=?",(Item_ID,),).fetchone()
            for i in d :
                print(type(i))
                i = int(i)
                print(f"Old Number : {i}")
                new_n = i + n
                print(f"New Number : {new_n}")
                c.execute(f"UPDATE {Item_ID[0]} SET QTY={new_n} WHERE Barcode=?",(Item_ID,),)
            d = c.execute(f"SELECT * FROM {l} WHERE Barcode=?",(Item_ID,),).fetchall()
            print(d)
            db.commit()



def Remove_item():
    Item_ID = input("Enter Item ID: ").upper()
    n = int(input("Enter Removed Number: "))

    if len(Item_ID) > 6 or len(Item_ID) < 6 and type(Item_ID) != str() and Item_ID[0] in files:
        print("Code Incorrect")
    else:
            l = Item_ID[0]
            d = c.execute(f"SELECT QTY FROM {l} WHERE Barcode=?",(Item_ID,),).fetchone()
            for i in d :
                print(type(i))
                i = int(i)
                print(f"Old Number : {i}")
                new_n = i - n
                print(f"New Number : {new_n}")
                c.execute(f"UPDATE {Item_ID[0]} SET QTY={new_n} WHERE Barcode=?",(Item_ID,),)
            d = c.execute(f"SELECT * FROM {l} WHERE Barcode=?",(Item_ID,),).fetchall()
            print(d)
            db.commit()



def DisplayForm():
    #creating window
    display_screen = Tk()
    #setting width and height for window
    display_screen.geometry("800x800")
    #setting title for window
    display_screen.title("BMA Stock Database")
    global tree
    global SEARCH
    SEARCH = StringVar()
    #creating frame
    TopViewForm = Frame(display_screen, width=600, bd=1, relief=SOLID)
    TopViewForm.pack(side=TOP, fill=X)
    LeftViewForm = Frame(display_screen, width=600)
    LeftViewForm.pack(side=LEFT, fill=Y)
    MidViewForm = Frame(display_screen, width=600)
    MidViewForm.pack(side=RIGHT)
    lbl_text = Label(TopViewForm, text="BMA Stock Database", font=('verdana', 18), width=600,bg="#1C2833",fg="white")
    lbl_text.pack(fill=X)
    lbl_txtsearch = Label(LeftViewForm, text="Search", font=('verdana', 15))
    lbl_txtsearch.pack(side=TOP, anchor=W)

    search = Entry(LeftViewForm, textvariable=SEARCH, font=('verdana', 15), width=10)
    search.pack(side=TOP, padx=10, fill=X)
    btn_search = Button(LeftViewForm, text="Search", command=SearchRecord)
    btn_search.pack(side=TOP, padx=10, pady=10, fill=X)
    btn_search = Button(LeftViewForm, text="View All", command=DisplayData)
    btn_search.pack(side=TOP, padx=10, pady=10, fill=X)
        #setting scrollbar
    scrollbarx = Scrollbar(MidViewForm, orient=HORIZONTAL)
    scrollbary = Scrollbar(MidViewForm, orient=VERTICAL)
    tree = ttk.Treeview(MidViewForm,columns=("Barcode", "Product_Name", "QTY"),
                        selectmode="extended", height=100, yscrollcommand=scrollbary.set, xscrollcommand=scrollbarx.set)
    scrollbary.config(command=tree.yview)
    scrollbary.pack(side=RIGHT, fill=Y)
    scrollbarx.config(command=tree.xview)
    scrollbarx.pack(side=BOTTOM, fill=X)
    #setting headings for the columns
    tree.heading('Barcode', text="Barcode", anchor=W)
    tree.heading('Product_Name', text="Product_Name", anchor=W)
    tree.heading('QTY', text="QTY", anchor=W)
    #setting width of the columns
    tree.column('#0', stretch=NO, minwidth=0, width=0)
    tree.column('#1', stretch=NO, minwidth=0, width=100)
    tree.column('#2', stretch=NO, minwidth=0, width=150)

    tree.pack()
    DisplayData()
#function to search data
def SearchRecord():
    #checking search text is empty or not
    if SEARCH.get() != "":
        #clearing current display data
        tree.delete(*tree.get_children())
        #open database
        conn = sqlite3.connect('Stock.db')
        #select query with where clause
        cursor=conn.execute("SELECT * FROM A WHERE Product_Name LIKE ?", ('%' + str(SEARCH.get()) + '%',))
        #fetch all matching records
        fetch = cursor.fetchall()
        #loop for displaying all records into GUI
        for data in fetch:
            tree.insert('', 'end', values=(data))
        cursor.close()
        conn.close()
#defining function to access data from SQLite database
def DisplayData():
    #clear current data
    tree.delete(*tree.get_children())
    # open databse
    conn = sqlite3.connect('Stock.db')
    #select query
    cursor=conn.execute("SELECT * FROM A")
    #fetch all data from database
    fetch = cursor.fetchall()
    #loop for displaying all data in GUI
    for data in fetch:
        tree.insert('', 'end', values=(data))
    cursor.close()
    conn.close()

DisplayForm()
if __name__=='__main__':
#Running Application
    mainloop()


f = input("Hi , What Would You Like To Do ? ( Search - ADD - Remove - Create Barcodes) : ").upper()
if f != "SEARCH" and f != 'ADD' and f != 'REMOVE' and f != r"CREATE BARCODES":
    print("Invalid Command")
elif f == 'SEARCH' :
    Search_data()
elif f == 'ADD' :
    Add_item()
elif f == 'REMOVE' :
    Remove_item()
elif f ==  r"CREATE BARCODES" :
    Create_Barcode()