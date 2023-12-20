import sys,openpyxl,subprocess, os
from pathlib import Path
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox

SUCCESS_CODE = 0
ERROR_CODE = 1
EMPTY_NAME_ERROR = 2
EMPTY_ORIGINAL_FILE = 3
ORIGINAL_FILE_NOT_FOUND = 4

def execute_multiple(file_path, names):
	if not file_path:
		return EMPTY_ORIGINAL_FILE
	if not names:
		return EMPTY_NAME_ERROR

	name_array = [x.strip() for x in names.split(',')]
	results = []
	for name in name_array:
		results.append(execute_single(file_path, name))
		pass
	if all(code == SUCCESS_CODE for code in results):
		return SUCCESS_CODE
	else:
		return ERROR_CODE

def execute_single(file_path, name):
	if not file_path:
		return EMPTY_ORIGINAL_FILE
	if not name:
		return EMPTY_NAME_ERROR

	original_file_path = Path(file_path)
	generated_file_path = Path(f"{original_file_path.parent.absolute()}/{name}.xlsx")
	try:
	    os.remove(generated_file_path.absolute())
	except OSError:
	    pass
	
	try:
		subprocess.run(f"cp {original_file_path.absolute()} {generated_file_path.absolute()}", shell = True, executable="/bin/bash")
	except FileNotFoundError:
		return ORIGINAL_FILE_NOT_FOUND
		pass

	new_wb = openpyxl.load_workbook(generated_file_path)
	 
	grace_sheet = new_wb.active

	deleted_item_count = 0

	for row in reversed(list(grace_sheet.iter_rows())):
		if name not in str(row[5].value) and row[0].row not in range(6):
			deleted_item_count += 1
			grace_sheet.delete_rows(row[0].row)
			pass
	if deleted_item_count > 0:
		new_wb.save(generated_file_path)
		return SUCCESS_CODE
	else:
		subprocess.run(f"rm {generated_file_path}", shell = True, executable="/bin/bash")
		return ERROR_CODE
		pass

def error_popup(message):
	msg = QMessageBox()
	msg.setIcon(QMessageBox.Critical)
	msg.setText("Failed")
	msg.setInformativeText(message)
	msg.setWindowTitle("Error")
	msg.exec_()

def window():

	app = QApplication(sys.argv)
	win = QMainWindow()
	win.setWindowTitle("Testing")
	win.setGeometry(1200, 300, 700, 700)

	label_name = QtWidgets.QLabel(win)
	label_name.setText("Enter a name:")
	label_name.move(50,50)
	txt_name = QtWidgets.QLineEdit(win)
	txt_name.setFixedWidth(200)
	txt_name.move(200, 50)


	label_file = QtWidgets.QLabel(win)
	label_file.setText("Select the original file:")
	label_file.move(50,100)
	label_file.adjustSize()
	txt_file = QtWidgets.QLineEdit(win)
	txt_file.move(200, 100)
	txt_file.setFixedWidth(200)


	def dialog(self):
		fname = QFileDialog.getOpenFileName(win, "Select the entry excel file", "", "Excel Files (*.xlsx)")
		if fname:
			txt_file.setText(fname[0])
			pass

	button_browse_file = QtWidgets.QPushButton(win)
	button_browse_file.setText("Browse files")
	button_browse_file.move(400, 100)
	button_browse_file.clicked.connect(dialog)

	def click(self):
		result = execute_multiple(txt_file.text(),txt_name.text())
		if result == SUCCESS_CODE:
			msg = QMessageBox()
			msg.setIcon(QMessageBox.Information)
			msg.setText("Ok")
			msg.setInformativeText('File generated successfully')
			msg.setWindowTitle("Ok")
			msg.exec_()
			msg.buttonClicked.connect(sys.exit())
			pass
		elif result == EMPTY_NAME_ERROR:
			error_popup("Names are missing")
			pass
		elif result == EMPTY_ORIGINAL_FILE:
			error_popup("Original file is missing.")
			pass
		elif result == ORIGINAL_FILE_NOT_FOUND:
			error_popup("The original file is not found.")
			pass
		else:
			error_popup("Failed to generate data.")
			pass
		
	bt_generate = QtWidgets.QPushButton(win)
	bt_generate.setText("Filter")
	bt_generate.clicked.connect(click)
	bt_generate.move(200, 150)

	win.show()
	sys.exit(app.exec_())

window()




