from .classes.GetSelectionSemantic import GetSelectionSemantic as Semantic
from .classes.ProcessingDoc import WordDocument
from .classes.ProcessingOdt import OdtDocument
from qgis.core import QgsVectorLayer, QgsFeature, QgsGeometry, QgsPointXY, QgsProject, QgsPalLayerSettings, QgsField, QgsRenderContext
from PyQt5.QtCore import QVariant

def createButtonSlot():
	semanticData = Semantic.getSelectionSemanticData()
	if not semanticData:
		return
	
	Semantic.makeSelectionScreenPicture()
	
	# Создание docx документа
	wd = WordDocument()
	wd.addTableOnBookmark("Место1", Semantic.getSelectionCoordinates(reverse=True), enumerated=True)
	wd.addTableOnBookmark("Место1", Semantic.getSelectionSemanticData(), enumerated=False)
	wd.addSelectionPictureOnBookmark("Место1")
	
	# Создание odt документа
	odt = OdtDocument()
	odt.addTableOnBookmark("Место1", Semantic.getSelectionSemanticData())
	odt.addTableOnBookmark("Место1", Semantic.getSelectionCoordinates(reverse=True), enumerated=True)
	odt.addSelectionPictureOnBookmark("Место1")
	
	# Создание нового слоя, для отображения точек геометрии
	memory_layer = QgsVectorLayer("Point", "temp", "memory")
	provider = memory_layer.dataProvider()
	provider.addAttributes([QgsField("Номер", QVariant.Int)])
	memory_layer.updateFields()
	for num, point_coords in enumerate(Semantic.getSelectionCoordinates(reverse=False)):
		feature = QgsFeature()
		feature.setGeometry(QgsGeometry.fromPointXY(QgsPointXY(point_coords[0], point_coords[1])))
		feature.setFields(memory_layer.fields())
		feature["Номер"] = num + 1
		provider.addFeature(feature)
	
		
	QgsProject.instance().addMapLayer(memory_layer)
	
	#TODO Labels для созданного слоя
	'''
	memory_layer.setCustomProperty("labeling", "pal")
	memory_layer.setCustomProperty("labeling/drawLabels", True)
	memory_layer.setCustomProperty("labeling/enabled", True)
	memory_layer.setCustomProperty("labeling/fieldName", "Номер")
	memory_layer.setCustomProperty("labeling/fontSize", 10)
	memory_layer.setCustomProperty("labeling/isExpression", True)
	memory_layer.setLabelsEnabled(True)
	memory_layer.triggerRepaint()
	'''
	
	
	