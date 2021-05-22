from openpyxl import *
import sysconfig
import os
def findlinkbykitab (kitab):
	kitabdbpath="~/Downloads/Kitab_Database.xlsx"
	kitabdbpath= os.path.expanduser(kitabdbpath) 
	wb1=load_workbook(kitabdbpath)
	sheet_obj1 = wb1.active

	if kitab == "Aalim Ghulam":
		cell_obj1 = sheet_obj1.cell(row = 1, column = 2)
	elif kitab == "Academic Writing":
		cell_obj1 = sheet_obj1.cell(row = 2, column = 2)
	elif kitab == "Adab Arabi":
		cell_obj1 = sheet_obj1.cell(row = 3, column = 2)
	elif kitab == "Adab Fatemi":
		cell_obj1 = sheet_obj1.cell(row = 4, column = 2)
	elif kitab == "Barnamaj":
		cell_obj1 = sheet_obj1.cell(row = 5, column = 2)
	elif kitab == "Emotional":
		cell_obj1 = sheet_obj1.cell(row = 6, column = 2)
	elif kitab == "Free Period":
		cell_obj1 = sheet_obj1.cell(row = 7, column = 2)
	elif kitab == "HCIW":
		cell_obj1 = sheet_obj1.cell(row = 8, column = 2)
	elif kitab == "Ikhwan":
		cell_obj1 = sheet_obj1.cell(row = 9, column = 2)
	elif kitab == "Language":
		cell_obj1 = sheet_obj1.cell(row = 10, column = 2)
	elif kitab == "Literature":
		cell_obj1 = sheet_obj1.cell(row = 11, column = 2)
	elif kitab == "Majalis":
		cell_obj1 = sheet_obj1.cell(row = 12, column = 2)
	elif kitab == "Management":
		cell_obj1 = sheet_obj1.cell(row = 13, column = 2)
	elif kitab == "Maqamat":
		cell_obj1 = sheet_obj1.cell(row = 14, column = 2)
	elif kitab == "Maqraat":
		cell_obj1 = sheet_obj1.cell(row = 15, column = 2)
	elif kitab == "Masool":
		cell_obj1 = sheet_obj1.cell(row = 16, column = 2)
	elif kitab == "Mukhtasar":
		cell_obj1 = sheet_obj1.cell(row = 17, column = 2)
	elif kitab == "Muntakhaba":
		cell_obj1 = sheet_obj1.cell(row = 18, column = 2)
	elif kitab == "Nehj":
		cell_obj1 = sheet_obj1.cell(row = 19, column = 2)
	elif kitab == "Risala Alif":
		cell_obj1 = sheet_obj1.cell(row = 20, column = 2)
	elif kitab == "Risala B":
		cell_obj1 = sheet_obj1.cell(row = 21, column = 2)
	elif kitab == "Takhassus":
		cell_obj1 = sheet_obj1.cell(row = 22, column = 2)
	elif kitab == "Uloom Quran":
		cell_obj1 = sheet_obj1.cell(row = 23, column = 2)
	elif kitab == "Uyun":
		cell_obj1 = sheet_obj1.cell(row = 24, column = 2)
	else:
		print ("Something went wrong")
	link=(cell_obj1.value)

	webbrowser.open_new (link)

def savepagenumber (subject, newnumber):
	listofkitabs={"Aalim Ghulam":"B1", "Academic Writing":"B2", "Adab Arabi":"B3", "Adab Fatemi":"B4", "Barnamaj": "B5", "Emotional":"B6", "Free Period":"B7", 
	"HCIW":"B8", "Ikhwan": "B9", "Language":"B10", "Literature":"B11", "Majalis":"B12", "Management":"B13", "Maqamat":"B14","Maqraat":"B15", "Masool":"B16","Mukhtasar":"B17", 
	"Muntakhaba":"B18", "Nehj":"B19", "Risala Alif": "B20", "Risala B":"B21", "Takhassus":"B22", "Uloom Quran":"B23","Uyun":"B24"
	}
	from openpyxl import load_workbook
	nishanipath="~/Downloads/Nishani.xlsx"
	nishanipath=os.path.expanduser(nishanipath)
	workbook = load_workbook(nishanipath)
	sheet_nishani = workbook.active
	sheet_nishani [listofkitabs.get(subject)] = newnumber
	workbook.save(nishanipath)

def findpreviouspages (kitab):
	

	nishanipath="~/Downloads/Nishani.xlsx"
	nishanipath=os.path.expanduser(nishanipath)
	wb2=load_workbook(nishanipath)
	sheet_obj2 = wb2.active
	if kitab == "Aalim Ghulam":
		cell_obj2 = sheet_obj2.cell(row = 1, column = 2)
	elif kitab == "Academic Writing":
		cell_obj2 = sheet_obj2.cell(row = 2, column = 2)
	elif kitab == "Adab Arabi":
		cell_obj2 = sheet_obj2.cell(row = 3, column = 2)
	elif kitab == "Adab Fatemi":
		cell_obj2 = sheet_obj2.cell(row = 4, column = 2)
	elif kitab == "Barnamaj":
		cell_obj2 = sheet_obj2.cell(row = 5, column = 2)
	elif kitab == "Emotional":
		cell_obj2 = sheet_obj2.cell(row = 6, column = 2)
	elif kitab == "Free Period":
		cell_obj2 = sheet_obj2.cell(row = 7, column = 2)
	elif kitab == "HCIW":
		cell_obj2 = sheet_obj2.cell(row = 8, column = 2)
	elif kitab == "Ikhwan":
		cell_obj2 = sheet_obj2.cell(row = 9, column = 2)
	elif kitab == "Language":
		cell_obj2 = sheet_obj2.cell(row = 10, column = 2)
	elif kitab == "Literature":
		cell_obj2 = sheet_obj2.cell(row = 11, column = 2)
	elif kitab == "Majalis":
		cell_obj2 = sheet_obj2.cell(row = 12, column = 2)
	elif kitab == "Management":
		cell_obj2 = sheet_obj2.cell(row = 13, column = 2)
	elif kitab == "Maqamat":
		cell_obj2 = sheet_obj2.cell(row = 14, column = 2)
	elif kitab == "Maqraat":
		cell_obj2 = sheet_obj2.cell(row = 15, column = 2)
	elif kitab == "Masool":
		cell_obj2 = sheet_obj2.cell(row = 16, column = 2)
	elif kitab == "Mukhtasar":
		cell_obj2 = sheet_obj2.cell(row = 17, column = 2)
	elif kitab == "Muntakhaba":
		cell_obj2 = sheet_obj2.cell(row = 18, column = 2)
	elif kitab == "Nehj":
		cell_obj2 = sheet_obj2.cell(row = 19, column = 2)
	elif kitab == "Risala Alif":
		cell_obj2 = sheet_obj2.cell(row = 20, column = 2)
	elif kitab == "Risala B":
		cell_obj2 = sheet_obj2.cell(row = 21, column = 2)
	elif kitab == "Takhassus":
		cell_obj2 = sheet_obj2.cell(row = 22, column = 2)
	elif kitab == "Uloom Quran":
		cell_obj2 = sheet_obj2.cell(row = 23, column = 2)
	elif kitab == "Uyun":
		cell_obj2 = sheet_obj2.cell(row = 24, column = 2)
	else:
		print ("Something went wrong")

	prevpage=(cell_obj2.value)
	print ("In {}, you last reached {}".format(kitab, prevpage))

import webbrowser
webbrowser.open ('https://meet.google.com/xpz-etja-vee')
from datetime import datetime
import openpyxl
schedpath = "~/Downloads/Schedule.xlsx"
schedpath=os.path.expanduser(schedpath)
wb_obj = openpyxl.load_workbook(schedpath)
sheet_obj = wb_obj.active
daynum=datetime.today().weekday()
daynum=daynum+2


periodindex = [2,3,4,5,6,7,8,9,10]
for index in periodindex:
	cell_obj = sheet_obj.cell(row = index, column = daynum)
	subject = cell_obj.value
	findpreviouspages(subject)
	findlinkbykitab (subject)
	pagenumber=input ("Enter the page number you reached in this period. ")
	savepagenumber(subject, pagenumber)
	




