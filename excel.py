
# Take information from a cell of differnt files , then make folders and arrange the files into the folders made, Romil Bahukhandi
import os
import shutil
import xlrd 
import glob 
from shutil import move
#change this location to the loaction of your files
os.chdir('C:\Users\Triotree\Desktop\New Folder')

modules=[]
with_slash=[]
files=[]
exceptions=[]

files=glob.glob('*.xlsx')
# list of  all files with extension .xlsx in new folder stored in list files 

for i in files :
	book=xlrd.open_workbook(i)
	sheet=book.sheet_by_index(0)
	#each workbok is opened and data from cell no b3 is extracted and printed
	cell=sheet.cell(2,1).value
	cell=str(cell)
	cell=str.lower(cell)
	if  cell== None or cell == "" or  cell == " " :
			exceptions.append(i) #the file name with missing or blank data is added to the exceptions tab 
	else:
		if cell not in modules:#checking for duplicates 
			modules.append(cell)
		else:
			pass

count=0
for a in files :
	book=xlrd.open_workbook(a)
	sheet=book.sheet_by_index(0)
	cell=sheet.cell(2,1).value
	cell=str(cell)
	cell=str.lower(cell)
	for j in  modules:
		if j==cell:
			print a
			shutil.move(os.path.join(source,a),os.path.join(destination,cell))
			count=count+1
			print count
	print "Moved"		
#checking the data in colum B3 and moving the file into the folder with the same name 			