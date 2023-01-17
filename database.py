import sqlite3
import openpyxl
from barcode import Code39
from barcode.writer import ImageWriter

#Creating The Database File And opening It
db = sqlite3.connect("Stock.db")
c = db.cursor()


def Search_data():
    code = input('Input Your Code: ').upper()
    print(code)
    if len(code) != 6 :
        print("Code Incorrect")
    else:
        d = c.execute(f"SELECT * FROM Main WHERE Barcode=?",(code,),).fetchall()
        print(d)




def Create_Barcode():
    #Getting Data From The Database
    c.execute(f'''select Barcode , Product_Name from Main''')
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
    n_rows = int(input("Enter Number Of Rows: "))
    #Opening the Excel 
    wb = openpyxl.load_workbook(f"Main.xlsx")
    ws = wb.active
    ID = 1
    #Creating The Table If Not Exist
    c.execute(f"DROP TABLE IF EXISTS Main")
    c.execute(f"CREATE TABLE Main(Barcode TEXT NOT NULL UNIQUE, Product_Name TEXT NOT NULL , QTY TEXT NOT NULL)")
    while ID <= n_rows:
        # Getting The Data From Excel
        Barcode = ws[f"A{ID}"].value
        Product_Name = ws[f"B{ID}"].value
        n = 1
        #Inserting The Data Into Sqlite Database
        c.execute(f"INSERT INTO Main(Barcode , Product_Name, QTY) values(?, ?, ?)",(Barcode , Product_Name, n))
        #Printing The Final Result
        print("Data Inserted in the table: ")
        data=c.execute(f'''SELECT * FROM Main''')
        for row in data:
            print(row)
        #increasing the ID Count to Switch to next Row
        ID = ID + 1
    db.commit()



def Add_item():
    Item_ID = input("Enter Item ID: ").upper()
    n = int(input("Enter Added Number: "))

    if len(Item_ID) > 6 or len(Item_ID) < 6 and type(Item_ID) != str():
        print("Code Incorrect")
    else:
            d = c.execute(f"SELECT QTY FROM Main WHERE Barcode=?",(Item_ID,),).fetchone()
            for i in d :
                print(type(i))
                i = int(i)
                print(f"Old Number : {i}")
                new_n = i + n
                print(f"New Number : {new_n}")
                c.execute(f"UPDATE {Item_ID[0]} SET QTY={new_n} WHERE Barcode=?",(Item_ID,),)
            d = c.execute(f"SELECT * FROM Main WHERE Barcode=?",(Item_ID,),).fetchall()
            print(d)
            db.commit()



def Remove_item():
    Item_ID = input("Enter Item ID: ").upper()
    n = int(input("Enter Removed Number: "))

    if len(Item_ID) > 6 or len(Item_ID) < 6 and type(Item_ID) != str():
        print("Code Incorrect")
    else:
            d = c.execute(f"SELECT QTY FROM Main WHERE Barcode=?",(Item_ID,),).fetchone()
            for i in d :
                print(type(i))
                i = int(i)
                print(f"Old Number : {i}")
                new_n = i - n
                print(f"New Number : {new_n}")
                c.execute(f"UPDATE {Item_ID[0]} SET QTY={new_n} WHERE Barcode=?",(Item_ID,),)
            d = c.execute(f"SELECT * FROM Main WHERE Barcode=?",(Item_ID,),).fetchall()
            print(d)
            db.commit()


f = input("Hi , What Would You Like To Do ? ( Search - ADD - Remove - Create Barcodes) : ").lower()
if f != "search" and f != 'add' and f != 'remove' and f != r"create barcodes" and f != "excel":
    print("Invalid Command")
elif f == 'search' :
    Search_data()
elif f == 'add' :
    Add_item()
elif f == 'remove' :
    Remove_item()
elif f ==  r"create barcodes" :
    Create_Barcode()
elif f == r"excel" :
    Excel_Data()