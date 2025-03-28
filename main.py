# import pythoncom
# from win32com.client import gencache, Dispatch
#
# kompas6_constants = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants
# kompas6_constants_3d = gencache.EnsureModule("{2CAF168C-7961-4B90-9DA2-701419BEEFE3}", 0, 1, 0).constants
#
# kompas_api7_module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
# kompas_api7_module.IApplication(Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(kompas_api7_module.IApplication.CLSID, pythoncom.IID_IDispatch))

# print(type(kompas_api7_module))

import pythoncom
from win32com.client import Dispatch, gencache

module7 = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
api7 = module7.IApplication(Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(module7.IApplication.CLSID, pythoncom.IID_IDispatch))
const7 = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants
api7.Visible = True
print(type(api7))
doc2D = api7.ActiveDocument
print(type(doc2D))

KompasDocument2D = doc2D._oleobj_.QueryInterface(module7.NamesToIIDMap['IKompasDocument2D'], pythoncom.IID_IDispatch)
KompasDocument2D1 = doc2D._oleobj_.QueryInterface(module7.NamesToIIDMap['IKompasDocument2D1'], pythoncom.IID_IDispatch)
doc2D = module7.IKompasDocument2D(KompasDocument2D)
views = doc2D.ViewsAndLayersManager.Views
AssociationView = views.Add(3)._oleobj_.QueryInterface(module7.NamesToIIDMap['IAssociationView'], pythoncom.IID_IDispatch)
view = module7.IAssociationView(AssociationView)
view.SourceFileName = "D:\\KOMPAS SDK\\Parts\\test1.m3d"
view.Unfold = True
view.Update()
print(type(doc2D))
doc2D = module7.IKompasDocument2D1(KompasDocument2D1)
print(type(doc2D))
doc2D.RebuildDocument()