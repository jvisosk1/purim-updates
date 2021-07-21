import datetime
import xlrd
# https://www.geeksforgeeks.org/reading-excel-file-using-python/

from docx.document import Document as _Document
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.table import _Cell, Table, _Row
from docx.text.paragraph import Paragraph
# https://github.com/python-openxml/python-docx/issues/650
import docx
# https://stackabuse.com/reading-and-writing-ms-word-files-in-python-via-python-docx-module/
# https://pbpython.com/python-word-template.html
# https://python-docx.readthedocs.io/en/latest/user/quickstart.html#adding-a-table
# CHANGE FONT AND SIZE
# https://python-docx.readthedocs.io/en/latest/user/text.html#paragraph-properties
import os
# OS https://stackabuse.com/creating-and-deleting-directories-with-python/
import shutil
from docx.shared import Pt


GREETING_COST = 5
FREE_RECIP_GREETINGS = 10
RECIPROCITY_COST = 25
UNLIMITED_COST = 250
COLUMNS = 3
SHOW_ID = False


class Member:
  	def __init__(self, id, firstName, lastName):
	    self.id = id
	    self.firstName = firstName
	    self.lastName = lastName
	    self.cellText = str(id) + " " + firstName + " " + lastName
	    self.greeters = []


def iter_block_items(parent):
    if isinstance(parent, _Document):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    elif isinstance(parent, _Row):
        parent_elm = parent._tr
    else:
        raise ValueError("something's not right")
    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)


# READ MEMBER LIST, CREATE MEMBER OBJECTS, APPEND TO LIST  
def readMemberList(members):
	# OPEN EXCEL WORKBOOK, SELECT SHEET, CREATE DOCX OBJECT
	members_file = xlrd.open_workbook("./members.xlsx") 
	sheet = members_file.sheet_by_index(0) 

	numRows = sheet.nrows
	numCols = sheet.ncols

	for i in range(1, numRows):
		id = int(sheet.cell_value(i, 0))
		firstName = sheet.cell_value(i, 2).strip()
		lastName = sheet.cell_value(i, 1).strip()
		members.append(Member(id, firstName, lastName))


# WRITE MEMBER NAMES/ID TO GREETING REQUEST WORD DOC
def printMemberNames(members):

	# doc = docx.Document()
	doc = docx.Document('order_form_template.docx')

	for block in iter_block_items(doc):

		if isinstance(block, Table):
			rowsWordDoc = int(len(members) / COLUMNS + 1)
			table = block

			j = 0
			row = table.rows[0]
			row.height = Pt(24)

			for k in members:
				row.cells[j].text = k.cellText
				if j == 2:
					j = 0
					row = table.add_row()
					row.height = Pt(24)
				else: j+=1 

	doc.save("./output/" +'member_names.docx')


# CREATE DIRECTORY FOR ALL OUTPUT FILES
def createFolder():
	path = os.getcwd() + "/output"
	
	try:
		os.mkdir(path)
	except OSError:
		print("Output folder already exists.")


members = []
orders = []
greetings = []


createFolder()
readMemberList(members)
printMemberNames(members)



