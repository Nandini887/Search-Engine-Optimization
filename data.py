from pyexcel_xlsx import read_data
from urllib.request import urlopen
from bs4 import BeautifulSoup
import re
import sqlite3
import xlsxwriter
data=read_data("data.xlsx")
url=''
keyword=''
for x in data:
    url=data[x][0][0]
    keyword=data[x][0][1]


mydata=urlopen(url)
s=BeautifulSoup(mydata,'html5lib')
pagedata=s.body.get_text
pattern=re.compile(keyword)
nlist=pattern.findall(str(pagedata))
totsize=len(str(pagedata).split())
density=len(nlist)/totsize
con=sqlite3.connect("mydb.db")
try:
    con.execute("create table data_table(url text,keyword text,density float)")
    con.commit()
except:
    print("cannot execute")

con.execute("insert into data_table values(?,?,?)",(url,keyword,density))
con.commit()
nurl=0
nkeyword=0
ndensity=0
curr=con.execute("select * from data_table")
for x in curr:
    nurl=x[0]
    nkeyword=x[1]
    ndensity=x[2]

workbook=xlsxwriter.Workbook("output.xlsx")
worksheet=workbook.add_worksheet()
worksheet.write('A1',nurl)
worksheet.write('A2',nkeyword)
worksheet.write('A3',ndensity)
worksheet.write('A4',100)
chart=workbook.add_chart({'type':'pie'})
chart.add_series({'values':'=Sheet1!$A$3:$A$4'})
worksheet.insert_chart('D5',chart)
workbook.close()
















    





















