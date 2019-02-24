import sys, os

root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


from odf.opendocument import load
from odf import teletype, draw
from odf.table import Table, TableColumn, TableRow, TableCell
from odf.text import P, Bookmark, BookmarkRef, BookmarkStart, BookmarkEnd, Span

from PyQt5.QtCore import QVariant
class OdtDocument():
	
	
	def __init__(self):
		self.doc = load(root + r"/documentTemplates/Шаблон.odt")
	
	
	def addTableOnBookmark(self, bookmarkName, tableData, enumerated=False):
		'''Вставка таблицы перед закладкой
		'''
		table_columns = 3 if enumerated else 2
		#Создание и заполнение таблицы
		table = Table()
		table.addElement(TableColumn(numbercolumnsrepeated=table_columns))
		for index, row in enumerate(tableData):
			tr = TableRow()
			table.addElement(tr)
			if enumerated:
				tc = TableCell()
				tr.addElement(tc)
				tc.addElement(P(text=str(index + 1)))
			for item in row:
				tc = TableCell()
				tr.addElement(tc)
				tc.addElement(P(text=str(item) if type(item)!=QVariant else ''))
		bookmarks = self.doc.getElementsByType(BookmarkStart)
		#Вставка таблицы в content.xml
		for bookmark in bookmarks:
			if bookmark.getAttribute("name") == bookmarkName:
				bookmark.parentNode.parentNode.insertBefore(table, bookmark.parentNode)
				bookmark.parentNode.parentNode.insertBefore(P(text=""), bookmark.parentNode)
		self.doc.save(root + r"/releasedDocs/Документ", True)
		
		
	def addSelectionPictureOnBookmark(self, bookmarkName):
		'''Вставка изображения перед закладкой
		'''
		picturePath = root + r"/picTmp/123.png"
		#Копирование изображения в документ
		pic = self.doc.addPicture(filename=picturePath)
		bookmarks = self.doc.getElementsByType(BookmarkStart)
		#Вставка изображения в content.xml
		for bookmark in bookmarks:
			if bookmark.getAttribute("name") == bookmarkName:
				df = draw.Frame(width="15cm", height="15cm")
				df.addElement(draw.Image(href=pic))
				p = P()
				p.appendChild(df)
				bookmark.parentNode.parentNode.insertBefore(p, bookmark.parentNode)
		bookmark.parentNode.parentNode.insertBefore(P(text=""), bookmark.parentNode)
		self.doc.save(root + r"/releasedDocs/Документ", True)