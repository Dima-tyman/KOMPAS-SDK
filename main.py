import os
import time
from itertools import count

from newDXFimport import importActive_to_dxf, importPath_to_dxf


def __convert_complete_mes(count):
    if count == 0:
        print("Complete")
    else:
        print("Converted %s files" % count)
    time.sleep(1)


def from_active_doc(path_to_save):
    count = importActive_to_dxf(path_to_save)
    __convert_complete_mes(count)

def from_path(path_to_save, path_to_saveDXF = None, path_to_saveCDW= None):

    if path_to_saveDXF is None:
        path_to_saveDXF = path_to_save
    if path_to_saveCDW is None:
        path_to_saveCDW = path_to_save

    i = 0

    for doc_name in os.listdir(path_to_save):
        if ".m3d" in doc_name:
            count2 = importPath_to_dxf(path_to_save, doc_name, path_to_saveDXF, path_to_saveCDW)
            i = i + count2
    __convert_complete_mes(i)


# from_path("X:\\1. Основное производство\\АИП мобильный аккумуляторный\\ПАПКА ДЛЯ ТЕСТА\\")