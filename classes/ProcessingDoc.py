import sys, os
root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

import win32com.client as win32
from PyQt5.QtCore import QVariant

class WordDocument():
	'''	Объект докумета Word, а также все действия с ним
	'''
	def __init__(self):
		self._word = win32.gencache.EnsureDispatch('Word.Application')
		self._document = self._word.Documents.Add(root + r"/documentTemplates/Шаблон.docx")
	
	
	def _showDoc(self):
		self._word.Visible = True

	
	def closeDoc(self):
		self._document.Close(False)
		self._word.Application.Quit()

	
	def addTableOnBookmark(self, bookmarkName, tableData, enumerated=False):
		'''Вставка таблицы после закладки и наполнение её семантикой
		'''
		table_columns = 3 if enumerated else 2
		insertion_indent = 2 if enumerated else 1
		# Получить Range закладки
		r = self._word.ActiveDocument.Bookmarks(bookmarkName).Range
		# Чтобы при вставке закладка сохранялась, добавляем после неё текст, и в его Range заменяем текст таблицей
		r.InsertParagraphAfter()
		r.InsertAfter(bookmarkName)
		newR = self._word.ActiveDocument.Content
		newR.Find.Execute(FindText=bookmarkName)
		# Создать и сформировать таблицу на месте текста
		tTableData = tuple(tableData)
		table = self._word.ActiveDocument.Tables.Add(newR, len(tTableData), table_columns)
		table.Borders.InsideLineStyle = 1
		table.Borders.OutsideLineStyle = 1
		for index, row in enumerate(tTableData):
			if enumerated:
				table.Cell(index + 1, 1).Range.Text = index + 1
			table.Cell(index + 1, insertion_indent).Range.Text = row[0]
			table.Cell(index + 1, insertion_indent + 1).Range.Text = '' if type(row[1]) is QVariant  else row[1]
		self._showDoc()
	
	def addSelectionPictureOnBookmark(self, bookmarkName):
		'''Вставка изображения после закладки
		'''
		picturePath = root + r"/picTmp/123.png"
		# Получить Range закладки
		r = self._word.ActiveDocument.Bookmarks(bookmarkName).Range
		# Чтобы при вставке закладка сохранялась, добавляем после неё текст, и в его Range заменяем текст таблицей
		r.InsertParagraphAfter()
		r.InsertAfter(bookmarkName)
		newR = self._word.ActiveDocument.Content
		newR.Find.Execute(FindText=bookmarkName)
		# Вставить изображение на место текста
		newR.Text = ""
		newR.InlineShapes.AddPicture(picturePath, Range=newR)
		
