from qgis.utils import iface
import os
root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

class GetSelectionSemantic():

	@staticmethod
	def _tryGetSelection():
		activeLayer = iface.activeLayer()
		if activeLayer.selectedFeatureCount() != 1:
			return None
		else:
			return activeLayer
		
	@staticmethod
	def getSelectionSemanticData():
		activeLayer = GetSelectionSemantic._tryGetSelection()
		if not activeLayer:
			return None
		else:
			return ((field.name(), feature[field.name()]) for feature in activeLayer.getSelectedFeatures() for field in feature.fields())
        
		
	@staticmethod
	def getSelectionCoordinates(reverse=False):
		import re
		activeLayer = GetSelectionSemantic._tryGetSelection()
		if not activeLayer:
			return None
		geometryStr = "".join((feature.geometry().asWkt() for feature in activeLayer.getSelectedFeatures()))
		r = re.findall(r"[.\d]+", geometryStr)	
		rounded_coords = [round(float(coord), 2) for coord in r]
		start_pos, end_pos  = (1, 0) if reverse else (0, 1)
		return zip(rounded_coords[start_pos::2],rounded_coords[end_pos::2])
		
		
			
	@staticmethod
	def makeSelectionScreenPicture():
		iface.mapCanvas().saveAsImage(root + r"/picTmp/123.png")