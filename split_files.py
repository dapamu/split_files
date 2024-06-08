#split_files.py
#utility to split one file into separate files
#DPM 
import os
from tkinter import filedialog
import xlwings as xw
#import pandas as pd
import pdfquery as pq
from PyPDF2 import PdfWriter, PdfReader

#split pages of pdf file into separate files
def pdf(file, extension):
	fh = PdfReader(file)

	stpos = file.rfind('/')
	endpos = file.rfind('.')
	file_name = file[stpos + 1:endpos]
	
	#lines = fh.pages[0]
	#print(lines.extract_text())
	for i in range(len(fh.pages)):
		output = PdfWriter()
		output.add_page(fh.pages[i])
		#with open("document-page%s.pdf" % i, "wb") as outputStream:
		#	output.write(outputStream)
		with open(file_name + "_" + str(i) + "." + extension, "wb") as outputStream:
			output.write(outputStream)

#extract information from pdf file
def pdf_info(file):
	fh = pq.PDFQuery(file)
	fh.load()
	fh.tree.write("file.xml",pretty_print = True)
	fh

#do something with an excel file
def excel(file, extension):
	#fh = pd.ExcelFile(file)
	#print(fh.sheet_names)
	excel_app = xw.App(visible=False)
	wb = excel_app.books.open(file)
	for sheet in wb.sheets:
		sheet.api.Copy()
		wb_new = xw.books.active
		wb_new.save(f"{sheet.name}.{extension}")
		wb_new.close()
	excel_app.quit()


#do something with text files
def text(file):
	with open(file, 'r') as fh:
		lines = fh.readlines()
		print(lines)

#ask what file to open
flag = input("Do you want to split a file? Y/N: ")
#flag = "Y"

if flag.upper() == "Y":
	file_path = filedialog.askopenfilename()

	if file_path == "":
		print("No file selected. Exiting...")
	else:
		atpos = file_path.rfind('.')

		#if there is no extension, keep blank otherwise get extension
		if atpos == -1:
			extension = ""
		else:
			extension = file_path[atpos + 1:]
		if extension == 'pdf': #do something based on extension
			pdf(file_path, extension)
			#pdf_info(file_path)
		elif extension == 'xls':
			excel(file_path, extension)
		elif extension =='txt':
			text(file_path)
		else: #try to open as text file. if unable, just exit out
			try:
				text(file_path)
			except:
				print("File type not supported")
else:
	print("No file selected. Exiting...")

