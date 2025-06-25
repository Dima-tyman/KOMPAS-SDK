import os

os.chdir("Parts")

for file in os.listdir():
    if file.find(".lnk") < 0 and file.find(".m3d") < 0:
        os.remove(os.getcwd() + "\\" + file)
