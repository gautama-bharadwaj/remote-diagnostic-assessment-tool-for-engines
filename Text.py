import xlrd
import tkinter
from tkinter import filedialog
import os
root = tkinter.Tk()
root.withdraw() #use to hide tkinter window
currdir = os.getcwd()
filename = filedialog.askopenfilename(parent=root,initialdir=currdir, title='Please select the Excel file')
if len(filename) > 0:
    loc = (filename)
# To open Workbook
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
root = tkinter.Tk()
root.withdraw() #use to hide tkinter window
currdir = os.getcwd()
tempdir = filedialog.askdirectory(parent=root, initialdir=currdir, title='Please select a folder for the text file')
f = open(tempdir+'/value.txt','w')
val=[]
val1=[]
for i in range(sheet.ncols):
	for j in range(sheet.nrows):
		if "Fuel Cell Current:" in sheet.cell_value(0,i):
			val.append(sheet.cell_value(j,i))
		elif "Fuel Cell Voltage:" in sheet.cell_value(0,i):
			val1.append(sheet.cell_value(j,i))	
mapped = zip(val,val1)
mapped = list(mapped)
k=1
while k < sheet.nrows:
	f.write(str(mapped[k]))
	f.write("\n")
	k+=1
f.close()
