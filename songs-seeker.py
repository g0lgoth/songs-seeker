import os
import time
from PyQt5.QtWidgets import QApplication, QLabel

def listCreation(path):
    os.chdir(path)
    list1 = os.listdir("./")
    list1.insert(0, len(list1))
    print(list1)
    return list1

def listCleaner(list):
    for i in list[0]:
        

def listIterator(path, list):


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