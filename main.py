import os
import time
from dxfImport import *


i = 0


def from_path(path_to_save):
    global i
    i = 0
    for doc_name in os.listdir(path_to_save):
        if doc_name.find(".m3d") < 0: continue

        print(doc_name)
        count_part = input("Input count: ")
        if count_part == "": count_part = 1

        res_name_dxf = doc_name.replace(".m3d", " - %s шт..dxf" %count_part)
        res_name_cdw = doc_name.replace(".m3d", ".cdw")
        path_source = path_to_save + doc_name
        path_result_dxf = path_to_save + res_name_dxf
        path_result_cdw = path_to_save + res_name_cdw

        import_to_dxf(path_source, path_result_dxf, path_result_cdw)
        i = i + 1
    convert_complete_mes()

def from_active_doc(path_to_save):
    count_part = input("Input count: ")
    importActive_to_dxf(path_to_save, count_part)
    convert_complete_mes(True)

def convert_complete_mes(simple = False):
    if simple : print("Complete")
    else: print("Converted %s files" % i)
    time.sleep(1)