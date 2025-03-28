import time
import pythoncom
from win32com.client import Dispatch, gencache
from win32com.server.policy import regPolicy


def get_kompas_api7():
    module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
    api = module.IKompasAPIObject(
        Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(module.IKompasAPIObject.CLSID,
                                                                 pythoncom.IID_IDispatch))
    const = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants
    return module, api, const

module7, api7, const7 = get_kompas_api7()  # Подключаемся к API7
app7 = api7.Application  # Получаем основной интерфейс
docs = app7.Documents
app7.Visible = True  # Показываем окно пользователю (если скрыто)
# app7.HideMessage = const7.ksHideMessageNo  # Отвечаем НЕТ на любые вопросы программы
print(app7.ApplicationName(FullName=True))  # Печатаем название программы

# while True:
print("Hello")
activeDoc = app7.ActiveDocument

KompasDocument2D = activeDoc._oleobj_.QueryInterface(module7.NamesToIIDMap['IKompasDocument2D'], pythoncom.IID_IDispatch)
doc2D = module7.IKompasDocument2D(KompasDocument2D)
views = doc2D.ViewsAndLayersManager.Views
if views(1) is not None:
    views(1).Delete()
    print("DELETE")

View = views.Add(1)._oleobj_.QueryInterface(module7.NamesToIIDMap['IView'], pythoncom.IID_IDispatch)
view = module7.IView(View)

view.Scale = 2
view.Name = "Test2"
view.Number = 5
view.Update()