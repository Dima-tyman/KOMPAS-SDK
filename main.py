import time
import pythoncom
from win32com.client import Dispatch, gencache


PATH = "D:\\KOMPAS SDK\\Parts\\test1.m3d"
DXFPATH = "D:\\KOMPAS SDK\\Parts\\test1.dxf"


# Application init
module7 = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
api7 = module7.IApplication(Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(module7.IApplication.CLSID, pythoncom.IID_IDispatch))
const7 = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants
api7.Visible = True
api7.HideMessage = const7.ksHideMessageYes

# Documents init and create
doc = api7.Documents.Add(1)
IKompasDocument2D = doc._oleobj_.QueryInterface(module7.NamesToIIDMap['IKompasDocument2D'], pythoncom.IID_IDispatch)
doc2D = module7.IKompasDocument2D(IKompasDocument2D)
IKompasDocument2D1 = doc._oleobj_.QueryInterface(module7.NamesToIIDMap['IKompasDocument2D1'], pythoncom.IID_IDispatch)
doc2D1 = module7.IKompasDocument2D1(IKompasDocument2D1)

# Change document sheet format
sheet = doc.LayoutSheets(0)
sheet.LayoutLibraryFileName = "C:\\Program Files\\ASCON\\KOMPAS-3D v22\\Sys\\graphic.lyt"
sheet.LayoutStyleNumber = 15 # 13 - not into frame, 15 - empty sheet
sheet.Update()

# Views init and create
views = doc2D.ViewsAndLayersManager.Views
view = views.Add(3)
IAssociationView = view._oleobj_.QueryInterface(module7.NamesToIIDMap['IAssociationView'], pythoncom.IID_IDispatch)
viewAssoc = module7.IAssociationView(IAssociationView)
IAssociationViewElements = view._oleobj_.QueryInterface(module7.NamesToIIDMap['IAssociationViewElements'], pythoncom.IID_IDispatch)
viewAssocElement = module7.IAssociationViewElements(IAssociationViewElements)
IViewDesignation = view._oleobj_.QueryInterface(module7.NamesToIIDMap['IViewDesignation'], pythoncom.IID_IDispatch)
viewDesign = module7.IViewDesignation(IViewDesignation)

# Add view parameter
viewAssoc.SourceFileName = PATH
viewAssoc.ProjectionName = "#Развертка"
view.Update()
viewAssoc.Unfold = True
viewAssocElement.CreateCentresMarkers = False
viewAssocElement.CreateCircularCentres = False
viewAssocElement.CreateLinearCentres = False
viewDesign.ShowUnfold = False
viewDesign.ShowScale = False
view.Update()
doc2D1.RebuildDocument()

# Save to DXF
doc.SaveAs(DXFPATH)

# Clear document / for test
time.sleep(1)
view.Delete()
doc.Close(0)
