import os
import time
from PyQt5.QtWidgets import QApplication, QLabel
import xlsxwriter

def listCreation(path):
    new_list = []
    os.chdir(path)
    new_list = os.listdir("./")
    new_list.insert(0, len(new_list))
    new_list=listCleaner(new_list)
    return new_list

def listCleaner(list_to_clean):
    list_ok = []
    max_value = list_to_clean[0]+1
    for i in range (1, max_value):
        if len(list_to_clean[i])<4:
            continue
        elif list_to_clean[i][-5:] == ".xlsx":
            list_to_clean[i] = ""
        elif list_to_clean[i][-4] == ".":
            list_to_clean[i] = ""
        else:
            continue
    for i in range (1, max_value):
        if list_to_clean[i] == "":
            continue
        else:
            list_ok.append(list_to_clean[i])
    return list_ok
            
def listExploration(path, entry_list):
    for i in range(1, entry_list[0]):
        os.chdir("{}\\{}".format(path, entry_list[i]))
        list_songs = os.listdir("./")
        list_songs= listCleaner(list_songs)


def tableCreation(path):
    location = os.chdir(path)
    list=os.listdir(location)
    listNumber=len(list)
    return list, listNumber

def excelManagement(path, file_name, entry_list):
    for folder in entry_list:
        os.chdir(path)
        wb = xlsxwriter.Workbook(file_name)
        worksheet = wb.add_worksheet(folder)
        songs_list = listCreation("{}\\{}".format(path, folder))
        # wrap_format = wb.add_format({'text_wrap': True}) # permet de sauter des lignes dans une cellule
        # cell_format_title = wb.add_format({'bold': True})
        # cell_format_title.set_font_size(13)
        # cell_format_title.set_align('center')
        # cell_format_title.set_align('vcenter')
        # cell_format_normal = wb.add_format({'bold': True})
        # cell_format_normal.set_font_size(13)
        # cell_format_normal.set_align('center')
        # cell_format_normal.set_align('vcenter')
        # worksheet.set_column('A:A', 40)
        # worksheet.write(0,0,"Songs", cell_format_title)
        # worksheet.set_column('B:B', 40)
        # worksheet.write(0,1,"Title", cell_format_title)
        # worksheet.set_column('C:C', 40)
        # worksheet.write(0,2,"Video size", cell_format_title)
        # worksheet.set_column('D:D', 40)
        # worksheet.write(0,3,"mp3 size", cell_format_title)
        iterator = 1
        for songs in songs_list:
            print(folder, songs)
            worksheet.write(iterator, 0, songs)
            iterator += 1
    wb.close()
    return

def guiCreation():
    app = QApplication([])
    label = QLabel('Hello World!')
    label.show()
    app.exec()  

def main():
    # path = "E:/UltraStar Deluxe/songs"
    main_path = "C:\\Programmes Gautier\\Gastoul-perso\\Download perso\\Vocaluxe_0.4.1_Windows_x64\\Songs"
    excel_name = 'songs_list.xlsx'
    folder_list = ()
    song_list = ()
    os.chdir(main_path)
    if os.path.exists(excel_name):
        os.remove(excel_name) 
    folder_list = listCreation(main_path)
    excelManagement(main_path, excel_name, folder_list)

if __name__ == "__main__":
    main()