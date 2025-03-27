import time
import pythoncom
from pywin.mfc.object import Object
from win32com.client import Dispatch, gencache

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
app7.HideMessage = const7.ksHideMessageNo  # Отвечаем НЕТ на любые вопросы программы
print(app7.ApplicationName(FullName=True))  # Печатаем название программы


def amount_sheet(doc7):
    sheets = {"A0": 0, "A1": 0, "A2": 0, "A3": 0, "A4": 0, "A5": 0}
    for sheet in range(doc7.LayoutSheets.Count):
        format = doc7.LayoutSheets.Item(sheet).Format  # sheet - номер листа, отсчёт начинается от 0
        sheets["A" + str(format.Format)] += 1 * format.FormatMultiplicity
    return sheets

while True:
    print("Hello")
    activeDoc = app7.ActiveDocument
    KompasDocument2D = activeDoc._oleobj_.QueryInterface(module7.NamesToIIDMap['IKompasDocument2D'],
                                                        pythoncom.IID_IDispatch)
    doc2D = module7.IKompasDocument2D(KompasDocument2D)

    View = doc2D.ViewsAndLayersManager.Views(1)._oleobj_.QueryInterface(module7.NamesToIIDMap['IView'],
                                                         pythoncom.IID_IDispatch)
    view = module7.IView(View)

    # print(view.Name)
    print(view.Scale)
    # view.Name = "Gigigi"
    # view.Scale = 1
    # view.Update()
    # print(view.Name)
    # print(view.Scale)

    time.sleep(2)

    # ViewsAndLayersManager
