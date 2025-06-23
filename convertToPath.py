import os
import time

from main import from_path


PATH = "D:\\KOMPAS SDK\\Parts\\"
startConvert = False

newPath = os.path.normpath(input("Input path: "))

if newPath[1:3] == ":\\":
    if os.path.exists(newPath):
        startConvert = True
    else:
        print("This path exist.")
else:
    print("Not input disk. Trying to search..")

    splitList = [x.strip() for x in newPath.split("\\") if x != ""]
    newPath = "\\".join(splitList)

    for disk in os.listdrives():
        if os.path.exists(disk + newPath):
            print("New path: " + disk + newPath)
            if input("If is not need disk, enter n: ").lower() != "n":
                newPath = disk + newPath
                startConvert = True
                break

if startConvert:
    print("Start convert..")
    if not os.path.isdir(newPath + "\\DXF"):
        os.mkdir(newPath + "\\DXF")
    if not os.path.isdir(newPath + "\\CDW"):
        os.mkdir(newPath + "\\CDW")
    from_path(newPath, newPath + "\\DXF", newPath + "\\CDW")
else:
    print("Path not found or exist")
    time.sleep(2)