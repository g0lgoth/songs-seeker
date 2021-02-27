import os
import time
from PyQt5.QtWidgets import QApplication, QLabel

def listCreation(path):
    os.chdir(path)
    list_folder = os.listdir("./")
    list_folder.insert(0, len(list_folder))
    print(list_folder)
    list_folder=listCleaner(list_folder)
    return list1

def listCleaner(list_to_clean):
    i=1
    for i in range list_to_clean[0]:
        if list_to_clean[i](-4) == ".":
            list_to_clean.remove(list_to_clean[i])
        else:
            continue
    print(list_to_clean)
    return list_to_clean
            
# def listIterator(path, list):

def tableCreation(path):
    location = os.chdir(path)
    list=os.listdir(location)
    listNumber=len(list)
    return list, listNumber

def guiCreation():
    app = QApplication([])
    label = QLabel('Hello World!')
    label.show()
    app.exec()  

def main():
    path = "E:/UltraStar Deluxe/songs"
    folder_list=()
    song_list=()
    folder_list=listCreation(path)

if __name__ == "__main__":
    main()