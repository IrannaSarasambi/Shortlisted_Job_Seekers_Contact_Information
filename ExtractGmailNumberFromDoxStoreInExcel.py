'''python Project to extract the resume file name, Gmail id and mobile number
from all shortlisted resume files given in  .docx format in a folder and
store all extracted information in excel sheet in a given path '''

import os
import re
import docx2txt
import xlsxwriter as xl

allgmail = []
allnumber = []
allresume = []

path = r'C:\Users\Desktop\allresme2'
files = os.listdir(path)

for file in files:
    allresume.append(file)
    filename = os.path.join(path,file)
    text = docx2txt.process(filename)

    email = re.compile(r"[a-z0-9\.\-+_]+@[a-z0-9\.\-+_]+\.[a-z]+")
    match1 = email.finditer(text)

    mobnum = re.compile(r"\d\d\d\d\d\d\d\d\d\d")
    match2 = mobnum.finditer(text)
    for x in match1:
        g = x.group(0)
        allgmail.append(g)

    for x in match2:
        p = x.group(0)
        allnumber.append(p)



w = xl.Workbook("C:\\Users\\che53832\Desktop\\all_xl_file\\gamilAndnumber.xlsx")
w1 = w.add_worksheet('gmailpnumber')
w1.write("A1","Resume File Name")
w1.write("B1","Gmail Id")
w1.write("C1", "Mobile Number")

for i in range(1,len(allgmail)+1):
    w1.write(i, 0, allresume[i-1])
    w1.write(i, 1, allgmail[i-1])
    w1.write(i, 2, allnumber[i-1])
w.close()

print("File name, Gmail ID and Mobile number successfully stored in excel sheet in given path")

