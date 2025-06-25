import pythoncom

from win32com.client import Dispatch, gencache

def __appinit():
    # Application init
    module7 = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
    api7 = module7.IApplication(
        Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(module7.IApplication.CLSID, pythoncom.IID_IDispatch))
    const7 = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants
    api7.Visible = True
    api7.HideMessage = const7.ksHideMessageYes

    # Documents init and create
    doc_source = api7.ActiveDocument
    doc = api7.Documents.Add(1)
    IKompasDocument2D = doc._oleobj_.QueryInterface(module7.NamesToIIDMap['IKompasDocument2D'], pythoncom.IID_IDispatch)
    doc2D = module7.IKompasDocument2D(IKompasDocument2D)
    IKompasDocument2D1 = doc._oleobj_.QueryInterface(module7.NamesToIIDMap['IKompasDocument2D1'],
                                                     pythoncom.IID_IDispatch)
    doc2D1 = module7.IKompasDocument2D1(IKompasDocument2D1)

    # Change document sheet format
    sheet = doc.LayoutSheets(0)
    sheet.LayoutLibraryFileName = "C:\\Program Files\\ASCON\\KOMPAS-3D v22\\Sys\\graphic.lyt"
    sheet.LayoutStyleNumber = 15  # 13 - not into frame, 15 - empty sheet
    sheet.Update()

    # Views init and create
    views = doc2D.ViewsAndLayersManager.Views
    view = views.Add(3)
    IAssociationView = view._oleobj_.QueryInterface(module7.NamesToIIDMap['IAssociationView'], pythoncom.IID_IDispatch)
    viewAssoc = module7.IAssociationView(IAssociationView)
    IAssociationViewElements = view._oleobj_.QueryInterface(module7.NamesToIIDMap['IAssociationViewElements'],
                                                            pythoncom.IID_IDispatch)
    viewAssocElement = module7.IAssociationViewElements(IAssociationViewElements)
    IViewDesignation = view._oleobj_.QueryInterface(module7.NamesToIIDMap['IViewDesignation'], pythoncom.IID_IDispatch)
    viewDesign = module7.IViewDesignation(IViewDesignation)

    # Embodiments init
    IEmbodimentsManager = viewAssoc._oleobj_.QueryInterface(module7.NamesToIIDMap['IEmbodimentsManager'],
                                                            pythoncom.IID_IDispatch)
    embodiments = module7.IEmbodimentsManager(IEmbodimentsManager)
    return doc, doc2D1, doc_source, sheet, view, viewAssoc, viewAssocElement, viewDesign, embodiments

def __rebuild(view, doc2D1):
    view.Update()
    doc2D1.RebuildDocument()

def __setparamDXF(view, doc2D1, viewAssoc, viewAssocElement, viewDesign, sheet):
    viewAssoc.ProjectionName = "#Развертка"
    __rebuild(view, doc2D1)
    viewAssoc.Unfold = True
    viewAssoc.BendLinesVisible = False
    viewAssocElement.CreateCentresMarkers = False
    viewAssocElement.CreateCircularCentres = False
    viewAssocElement.CreateLinearCentres = False
    viewDesign.ShowUnfold = False
    viewDesign.ShowScale = False
    __rebuild(view, doc2D1)
    sheet.LayoutStyleNumber = 15  # 13 - not into frame, 15 - empty sheet
    sheet.Update()

def __setparamDXF_foractive(view, doc2D1, viewAssoc, viewAssocElement, viewDesign, doc_source, sheet):
    viewAssoc.SourceFileName = doc_source.PathName
    __setparamDXF(view, doc2D1, viewAssoc, viewAssocElement, viewDesign, sheet)

def __setparamDXF_topath(view, doc2D1, viewAssoc, viewAssocElement, viewDesign, filePath, sheet):
    viewAssoc.SourceFileName = filePath
    __setparamDXF(view, doc2D1, viewAssoc, viewAssocElement, viewDesign, sheet)

def __setparamCDW(view, doc2D1, viewAssoc, viewDesign, sheet):
    viewAssoc.BendLinesVisible = True
    __rebuild(view, doc2D1)
    viewDesign.ShowUnfold = False
    viewDesign.ShowScale = False
    __rebuild(view, doc2D1)
    sheet.LayoutStyleNumber = 13  # 13 - not into frame, 15 - empty sheet
    sheet.Update()

def __input_count(fileName):
    count = input(fileName + "\nInput count: ")
    if not count.isdigit() or int(count) <= 0:
        return 0
    else:
        return count

def __set_enbodiment_name(oldName, enbodimentNum):
    if "." in oldName[:-4]:
        # Name with description
        newName = oldName[:oldName.find(" ")] + enbodimentNum + oldName[oldName.find(" "):]
    else:
        # Name without description
        newName = enbodimentNum + " " + oldName
    return newName


def importActive_to_dxf(path_to_save):
    doc, doc2D1, doc_source, sheet, view, viewAssoc, viewAssocElement, viewDesign, embodiments = __appinit()

    j = 0
    # Save to DXF base
    __setparamDXF_foractive(view, doc2D1, viewAssoc, viewAssocElement, viewDesign, doc_source, sheet)
    fileName = doc_source.Name
    count_part = __input_count(fileName)
    if count_part == 0:
        doc.SaveAs(path_to_save + fileName.replace(".m3d", ".dxf"))
    else:
        doc.SaveAs(path_to_save + fileName.replace(".m3d", " - %s шт..dxf" % count_part))

    # Save to CDW base
    __setparamCDW(view, doc2D1, viewAssoc, viewDesign, sheet)
    doc.SaveAs(path_to_save + fileName.replace(".m3d", ".cdw"))

    # Work with embodiment
    countEmbodiment = embodiments.EmbodimentCount

    if countEmbodiment > 1:
        j = 1
        for i in range(1, countEmbodiment):
            # Set next enbodiment
            embodiments.SetCurrentEmbodiment(i)
            __rebuild(view, doc2D1)
            embodimentNum = embodiments.GetCurrentEmbodimentMarking(-1, False)
            newName = __set_enbodiment_name(doc_source.Name, embodimentNum)

            # Save DXF
            __setparamDXF(view, doc2D1, viewAssoc, viewAssocElement, viewDesign, sheet)
            count_part = __input_count(newName)
            if count_part == 0:
                doc.SaveAs(path_to_save + newName.replace(".m3d", ".dxf"))
            else:
                doc.SaveAs(path_to_save + newName.replace(".m3d", " - %s шт..dxf" % count_part))

            # Save CDW
            __setparamCDW(view, doc2D1, viewAssoc, viewDesign, sheet)
            doc.SaveAs(path_to_save + newName.replace(".m3d", ".cdw"))

            j = j + 1

    doc.Close(0)
    return j

def importPath_to_dxf(file_path, file_name, path_to_saveDXF, path_to_saveCDW):
    doc, doc2D1, doc_source, sheet, view, viewAssoc, viewAssocElement, viewDesign, embodiments = __appinit()

    j = 1
    # Save to DXF
    __setparamDXF_topath(view, doc2D1, viewAssoc, viewAssocElement, viewDesign, file_path + file_name, sheet)
    count_part = __input_count(file_name)
    if count_part == 0:
        doc.SaveAs(path_to_saveDXF + file_name.replace(".m3d", ".dxf"))
    else:
        doc.SaveAs(path_to_saveDXF + file_name.replace(".m3d", " - %s шт..dxf" % count_part))

    # Save to CDW
    __setparamCDW(view, doc2D1, viewAssoc, viewDesign, sheet)
    doc.SaveAs(path_to_saveCDW + file_name.replace(".m3d", ".cdw"))

    # Work with embodiment
    countEmbodiment = embodiments.EmbodimentCount

    if countEmbodiment > 1:
        for i in range(1, countEmbodiment):
            # Set next enbodiment
            embodiments.SetCurrentEmbodiment(i)
            __rebuild(view, doc2D1)
            embodimentNum = embodiments.GetCurrentEmbodimentMarking(-1, False)[-3:]
            newName = __set_enbodiment_name(file_name, embodimentNum)

            # Save DXF
            __setparamDXF(view, doc2D1, viewAssoc, viewAssocElement, viewDesign, sheet)
            count_part = __input_count(newName)
            if count_part == 0:
                doc.SaveAs(path_to_saveDXF + newName.replace(".m3d", ".dxf"))
            else:
                doc.SaveAs(path_to_saveDXF + newName.replace(".m3d", " - %s шт..dxf" % count_part))

            # Save CDW
            __setparamCDW(view, doc2D1, viewAssoc, viewDesign, sheet)
            doc.SaveAs(path_to_saveCDW + newName.replace(".m3d", ".cdw"))

            j = j + 1

    doc.Close(0)
    return j
