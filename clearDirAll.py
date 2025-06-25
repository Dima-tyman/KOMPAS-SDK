import os

os.chdir("Parts")

if input("FULL CLEAR!\nInput any symbol for cancel: ") == "":
    for file in os.listdir():
        if file.find(".lnk") < 0:
            os.remove(os.getcwd() + "\\" + file)
