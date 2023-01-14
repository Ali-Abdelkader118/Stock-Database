import sqlite3
import openpyxl
import barcode

#Opening the Excel file
wb = openpyxl.load_workbook("Movies-Database.xlsx")
ws = wb.active
ID = 2

#Creating The Database File And opening It
db = sqlite3.connect("Movies.db")
c = db.cursor()

#Creating The Table If Not Exist
c.execute("DROP TABLE IF EXISTS Movies")
c.execute("CREATE TABLE Movies (MovieID TEXT NOT NULL UNIQUE, Movie_Name TEXT NOT NULL ,Movie_Year TEXT,Movie_Rate TEXT , Genere_1 TEXT ,Genere_2 TEXT ,Genere_3 TEXT ,Director TEXT, Runtime TEXT, Poster TEXT , Plot Text)")

while ID <= 1030:
    # Getting The Data From Excel  
    imdb = ws[f"H{ID}"].value
    M_name =ws[f"A{ID}"].value 
    M_year = ws[f"B{ID}"].value
    M_rate = ws[f"C{ID}"].value
    gn1 = ws[f"D{ID}"].value
    gn2 = ws[f"E{ID}"].value
    gn3 = ws[f"F{ID}"].value
    dr = ws[f"G{ID}"].value
    rt = ws[f"I{ID}"].value
    ps = ws[f"J{ID}"].value
    pl = ws[f"K{ID}"].value
    #Inserting The Data Into Sqlite Database
    c.execute("INSERT INTO Movies(MovieID , Movie_Name , Movie_Year , Movie_Rate ,Genere_1 , Genere_2 , Genere_3 , Director , Runtime , Poster , Plot ) values(?, ?, ?, ?, ?, ?, ?, ?, ?, ?,?)",(imdb , M_name , M_year , M_rate , gn1 , gn2, gn3, dr ,rt , ps , pl))
    #Printing The Fianl Result
    print("Data Inserted in the table: ")
    data=c.execute('''SELECT * FROM Movies''')
    for row in data:
        print(row)
    #increasing the ID Count to Switch to next Row
    ID = ID + 1

#Saving And Closing
db.commit()
db.close()