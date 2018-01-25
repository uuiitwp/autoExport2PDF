#coding:utf-8



'''
author:yzj
time:2017‎年‎9‎月‎21‎日19:23
'''




import arcpy
import pythonaddins

#####################出图前务必确认以下参数##########################


param = []
MxdPath = r"CURRENT"#########模板地址
PDFOutPath = r""
TitleName1 = "海宁市海昌街道迎丰村".decode('utf-8')
TitleName2 = "农村土地承包经营权地块分布图".decode('utf-8')
MinScale = 1000
MaxScale = 500

#####################出图前务必确认以上参数##########################


class ButtonClass1(object):
	"""Implementation for addin_addin.button (Button)"""
	def __init__(self):
		self.enabled = True
		self.checked = False
	def onClick(self):
		mxd = arcpy.mapping.MapDocument(MxdPath)
		layers = arcpy.mapping.ListLayers(mxd)#获取所有要素层
		if check(mxd,layers) == False:
			return
		print "开始出图".decode('utf-8')
		layer_target = layers[1]
		layer_other = layers[2]
		path = layer_other.dataSource
		path = path[0:path.rindex("\\")]
		path = path[0:path.rindex("\\")]
		#path为矢量库的根目录
		global PDFOutPath 
		PDFOutPath = path
		list_ZB = getList_ZB(layer_other)
		for ZB in list_ZB:
			exportPDFBy_ZB(ZB,layer_target,layer_other,mxd)
		print "出图完成".decode('utf-8')
		pythonaddins.MessageBox(("出图完成,PDF与mxd文件保存在矢量库目录:"+path.encode('utf-8')).decode('utf-8'),"完成".decode('utf-8'))
		
		

class ComboBoxClass7(object):
	"""Implementation for addin_addin.combobox (ComboBox)"""
	def __init__(self):
		self.items = ["海宁市海昌街道XX村".decode("utf-8"), "农村土地承包经营权地块分布图".decode("utf-8")]
		self.editable = True
		self.enabled = True
		self.dropdownWidth = 'WWWWWW'
		self.width = 'WWWWWWWWWWWWWWWWWWWW'
		self.index = -1
		self.text = ''
	def onSelChange(self, selection):
		self.index = self.items.index(selection)
		#print self.index
	def onEditChange(self, text):
		self.text = text
	def onFocus(self, focused):
		pass
	def onEnter(self):
		self.items[self.index] = self.text
		global param 
		param = self.items
	def refresh(self):
		pass


class ComboBoxClass11(object):
	"""Implementation for addin_addin.combobox_1 (ComboBox)"""
	def __init__(self):
		self.items = ["800".decode('utf-8'), "1000".decode('utf-8')]
		self.editable = True
		self.enabled = True
		self.dropdownWidth = 'WWWWWW'
		self.width = 'WWWWWW'
		self.index = -1
		self.text = ''
	def onSelChange(self, selection):
		self.index = self.items.index(selection)
	def onEditChange(self, text):
		self.text = text
	def onFocus(self, focused):
		pass
	def onEnter(self):
		self.items[self.index] = self.text
		global MinScale
		global MaxScale
		print type(self.items[1])
		print self.items[1]
		MinScale = int(self.items[1].encode('utf-8'))
		MaxScale = int(self.items[0].encode('utf-8'))
	def refresh(self):
		pass




#以下为代码

def check(mxd,layers):

		global param
		
		if len(param) <> 2:
			pythonaddins.MessageBox("请设置标题".decode('utf-8'),"错误".decode('utf-8'))
			return False
			
		if layers[1].isBroken == True or layers[2].isBroken:
			pythonaddins.MessageBox("请设置地块数据源".decode('utf-8'),"错误".decode('utf-8'))
			return False
			
		if layers[5].isBroken == True:
			pythonaddins.MessageBox("请设置图幅要素数据源".decode('utf-8'),"错误".decode('utf-8'))
			return False
			
		if layers[19].isBroken == True:
			pythonaddins.MessageBox("请设置影像要素数据源".decode('utf-8'),"错误".decode('utf-8'))
			return False
			
		return True
		
#返回最小的大于该比例尺的整百比例尺
def getScale(OldScale):
	return OldScale // 100 * 100 + 100


#获取组别唯一值
def getList_ZB(layer_other):

	layer_other.definitionQuery = ''
	rows = arcpy.SearchCursor(layer_other)
	list_ZB = []
	
	for row in rows:
		ZB = row.getValue('ZB')
		if ZB not in list_ZB:
			list_ZB.append(ZB)
			
	del rows
	del row
	return list_ZB

	
#获取标题
def getTitle(mxd):

	elements = arcpy.mapping.ListLayoutElements(mxd)
	for e in elements:
		if e.name == 'title':
			return e

			
			
def setTitleText(mxd,text):

	getTitle(mxd).text = text

def getWidthHeight(mxd):

	layers = arcpy.mapping.ListLayers(mxd)
	layer = layers[1]
	ex = layer.getExtent()
	return (ex.width,ex.height)
	
#根据组别选取		
def exportPDFBy_ZB(ZB,layer_target,layer_other,mxd):

	layer_target.definitionQuery = "ZB = '"+ ZB +"'"
	layer_other.definitionQuery = "ZB <> '"+ ZB +"'"	
	arcpy.SelectLayerByAttribute_management(r"数据\地块-当前组","NEW_SELECTION","[ZB] = '" + ZB + "'")
	df = arcpy.mapping.ListDataFrames(mxd)[0]
	df.zoomToSelectedFeatures()
	if df.scale < MaxScale:
		df.scale = MaxScale
	else:
		df.scale = getScale(df.scale)
	arcpy.SelectLayerByAttribute_management(r"数据\地块-当前组","CLEAR_SELECTION")
	mxdPageSize = mxd.pageSize
	global MinScale
	WH = getWidthHeight(mxd)
	
	global param
	setTitleText(mxd,param[0]+ZB+param[1])
	arcpy.RefreshTOC()
	arcpy.RefreshActiveView()
	global PDFOutPath

	WidthHeight = WH[0]/WH[1]
	if mxdPageSize[0]/mxdPageSize[1] >= 1 and WidthHeight >= 1:
		#横版
		if df.scale > MinScale:
		# 进入分幅模式
			tempGird = arcpy.CreateScratchName()
			arcpy.Delete_management(tempGird)
			arcpy.Delete_management(tempGird)

			arcpy.GridIndexFeatures_cartography(tempGird, r"数据\地块-当前组", "", "", "",(mxdPageSize[0]-5)/100* (MinScale - 400) ,(mxdPageSize[1]-5)/100* (MinScale-400),)
			layers = arcpy.mapping.ListLayers(mxd)
			grid_layer = None
			for layer in layers:
				tgname = tempGird[tempGird.rindex("\\")+1:]
				if layer.name == tgname:
					grid_layer = layer
					break
			cursors = arcpy.SearchCursor(grid_layer.dataSource)
			list_PageNumber = []
			for cursor in cursors:
				list_PageNumber.append(cursor.getValue("PageNumber"))
			for pn in list_PageNumber:
				grid_layer.definitionQuery = "PageNumber = " + str(pn)
				arcpy.SelectLayerByAttribute_management(tgname,"NEW_SELECTION","PageNumber = " + str(pn))
				df.zoomToSelectedFeatures()
				arcpy.SelectLayerByAttribute_management(tgname,"CLEAR_SELECTION")
				grid_layer.visible = False
				df.scale = MinScale 
				strpn = str(pn)
				while(len(strpn))<3:
					strpn = "0" + strpn
				setTitleText(mxd,param[0]+ZB+param[1]+strpn)
				arcpy.RefreshTOC()
				arcpy.RefreshActiveView()
				arcpy.mapping.ExportToPDF(mxd,PDFOutPath + "\\" + param[0]+ZB+param[1]+ strpn + '.pdf',"PAGE_LAYOUT",0,0,300,"NORMAL")
				mxd.saveACopy(PDFOutPath + "\\" + param[0]+ZB+param[1]+ strpn + '.mxd')
				tempdf = arcpy.mapping.ListDataFrames(mxd)[0]
				grid_layer.visible = True
			arcpy.Delete_management(tempGird)
			arcpy.Delete_management(tempGird)
			arcpy.mapping.RemoveLayer(tempdf,grid_layer)
				
			
		else:
			#直接出图
			arcpy.mapping.ExportToPDF(mxd,PDFOutPath + "\\" + param[0]+ZB+param[1] + '.pdf',"PAGE_LAYOUT",0,0,300,"NORMAL")
			mxd.saveACopy(PDFOutPath + "\\" + param[0]+ZB+param[1] + '.mxd')
			
	if mxdPageSize[0]/mxdPageSize[1] < 1 and WidthHeight < 1:
		if df.scale > MinScale:
		# 进入分幅模式
			tempGird = arcpy.CreateScratchName()
			arcpy.Delete_management(tempGird)
			arcpy.Delete_management(tempGird)

			arcpy.GridIndexFeatures_cartography(tempGird, r"数据\地块-当前组", "", "", "",(mxdPageSize[0]-5)/100* (MinScale - 400) ,(mxdPageSize[1]-5)/100* (MinScale-400),)
			layers = arcpy.mapping.ListLayers(mxd)
			grid_layer = None
			for layer in layers:
				tgname = tempGird[tempGird.rindex("\\")+1:]
				if layer.name == tgname:
					grid_layer = layer
					break
			cursors = arcpy.SearchCursor(grid_layer.dataSource)
			list_PageNumber = []
			for cursor in cursors:
				list_PageNumber.append(cursor.getValue("PageNumber"))
			for pn in list_PageNumber:
				grid_layer.definitionQuery = "PageNumber = " + str(pn)
				arcpy.SelectLayerByAttribute_management(tgname,"NEW_SELECTION","PageNumber = " + str(pn))
				df.zoomToSelectedFeatures()
				arcpy.SelectLayerByAttribute_management(tgname,"CLEAR_SELECTION")
				grid_layer.visible = False
				df.scale = MinScale 
				strpn = str(pn)
				while(len(strpn))<3:
					strpn = "0" + strpn
				setTitleText(mxd,param[0]+ZB+param[1]+strpn)
				arcpy.RefreshTOC()
				arcpy.RefreshActiveView()
				arcpy.mapping.ExportToPDF(mxd,PDFOutPath + "\\" + param[0]+ZB+param[1]+ strpn + '.pdf',"PAGE_LAYOUT",0,0,300,"NORMAL")
				mxd.saveACopy(PDFOutPath + "\\" + param[0]+ZB+param[1]+ strpn + '.mxd')
				tempdf = arcpy.mapping.ListDataFrames(mxd)[0]
				grid_layer.visible = True
			arcpy.Delete_management(tempGird)
			arcpy.Delete_management(tempGird)
			arcpy.mapping.RemoveLayer(tempdf,grid_layer)
		else:
			arcpy.mapping.ExportToPDF(mxd,PDFOutPath + "\\" + param[0]+ZB+param[1] + '.pdf',"PAGE_LAYOUT",0,0,300,"NORMAL")
			mxd.saveACopy(PDFOutPath + "\\" + param[0]+ZB+param[1] + '.mxd')
	
	
if __name__ == '__main__':

	print "开始出图"
	mxd = arcpy.mapping.MapDocument(MxdPath)
	layers = arcpy.mapping.ListLayers(mxd)#获取所有要素层
	layer_target = layers[1]
	layer_other = layers[2]
	list_ZB = getList_ZB(layer_other)
	
	for ZB in list_ZB:
		exportPDFBy_ZB(ZB,layer_target,layer_other,mxd)
	print "出图完成"
