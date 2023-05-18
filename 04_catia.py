import win32com.client.dynamic # Module for COM­Client
import sys, os   # Module for File­HandlingActiveDocumentSet
import win32gui # Module for MessageBox
import numpy as np 



################## saca la vista frontal de un part #######################


CATIA = win32com.client.Dispatch("catia.Application")

documents1 = CATIA.Documents

drawingDocument1 = documents1.Add("Drawing")

drawingSheets1 = drawingDocument1.Sheets

drawingSheet1 = drawingSheets1.Item("Sheet.1")

drawingViews1 = drawingSheet1.Views

drawingView1 = drawingViews1.Add("AutomaticNaming")

drawingViewGenerativeLinks1 = drawingView1.GenerativeLinks

drawingViewGenerativeBehavior1 = drawingView1.GenerativeBehavior

partDocument = CATIA.Documents

partDocument1 = partDocument.Item("Part2.CATPart")

product1 = partDocument1.GetItem("Part2")

drawingViewGenerativeBehavior1.Document = product1

drawingViewGenerativeBehavior1.DefineFrontView(1.000000, 0.000000, 0.000000, 0.000000, 1.000000, 0.000000)

drawingView1.x = 148.500000

drawingView1.y = 105.000000

drawingView1.Scale = 1.000000

drawingViewGenerativeBehavior1 = drawingView1.GenerativeBehavior

drawingViewGenerativeBehavior1.Update()

drawingView1.Activate()
