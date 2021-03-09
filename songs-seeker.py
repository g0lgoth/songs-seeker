import os
import time
from PyQt5.QtWidgets import QApplication, QLabel
import xlsxwriter

def listCreation(path):
    new_list = []
    os.chdir(path)
    new_list = os.listdir("./")
    # new_list.insert(0, len(new_list))
    new_list=listCleaner(new_list)
    return new_list

def listCleaner(list_to_clean):
    """
    Prend une liste et enlève les éléments qui ne sont pas un dossier
    fichiers zip, excel etc...
    """
    list_ok = []
    for i, item in zip(range (len(list_to_clean)), list_to_clean):
        if len(item)<4:
            continue
        elif item[-5:] == ".xlsx":
            list_to_clean[i] = ""
        elif item[-4] == ".":
            list_to_clean[i] = ""
        else:
            continue
    for item in list_to_clean:
        if item == "":
            continue
        else:
            list_ok.append(item)
    # print("input", list_to_clean)
    # print("output", list_ok)
    return list_ok
            
def listExploration(path, entry_list):
    for i in range(1, entry_list[0]):
        os.chdir("{}\\{}".format(path, entry_list[i]))
        list_songs = os.listdir("./")
        list_songs= listCleaner(list_songs)


def tableCreation(path):
    location = os.chdir(path)
    local_list = os.listdir(location)
    listNumber = len(local_list)
    return local_list, listNumber

def picturePresence(path):
    nok = "no"
    ok = "yes"
    try:
        for file in os.listdir(path): 
            if file.endswith(".jpg"):
                return ok
            else:
                return nok
        return nok
    except ValueError:
        return nok

def clipPresence(path):
    new_list = []
    temporary_list = []
    nok = "no"
    ok = "yes"
    try :
        for file in os.listdir(path):
            if file.endswith(".txt"):
                with open("{}{}".format(path, file), 'r') as f:
                    lines = f.readlines()
                for line in lines:
                    try:
                        if '#VIDEO:' in line:
                            new_line_name = line.replace('#VIDEO:', '')
                            new_line_name = new_line_name.replace('\n', '')
                            if "Ãª" in new_line_name:
                                new_line_name = new_line_name.replace("Ãª", 'ê')
                            if "Ã¨" in new_line_name:
                                new_line_name = new_line_name.replace("Ã¨", 'è')
                            if "Ã©" in new_line_name:
                                new_line_name = new_line_name.replace("Ã©", 'é')
                            if "Ã" in new_line_name:
                                new_line_name = new_line_name.replace("Ã", 'à')
                            # print(new_line_name)
                            try:
                                temporary_variable = os.stat("{}{}".format(path, new_line_name))
                                video_size = temporary_variable.st_size / (1024 * 1024)
                                return ok, video_size
                            except FileNotFoundError:
                                return "no file", 0    
                        else:
                            continue
                    except ValueError:
                        return "no video line in file", 0
            else :
                continue
        return nok, 0
    except ValueError:
        return "no file in directory", 0
    except IndexError:
        return "index error", 0

def sorter(path, folder_list):
    final_list = []
    for folder in folder_list:
        songs_list = listCreation("{}\\{}".format(path, folder))
        songs_list = listCleaner(songs_list)
        for song in songs_list:
            temporary_list = []
            temporary_list.append(folder)
            temporary_list.append(song)
            current_location = "{}\\{}\\{}\\".format(path, folder, song)
            video_status, video_size = clipPresence(current_location)
            temporary_list.append(video_status)
            temporary_list.append(video_size)
            image_status = picturePresence(current_location)
            temporary_list.append(image_status)
            final_list.append(temporary_list)
    # print("liste finale", final_list)
    return final_list

def excelManagement(path, entry_list, name_file):
    init = True
    os.chdir(path)
    wb = xlsxwriter.Workbook(name_file)
    for item in entry_list:
        # print(item)
        if init == True:
            worksheet = wb.add_worksheet(item[0])
            temporary_variable = item[0]
            init = False
            excel_new_worksheet = True
            line_iterator = 1
        else:
            if item[0] != temporary_variable:
                worksheet = wb.add_worksheet(item[0])
                temporary_variable = item[0]
                excel_new_worksheet = True
                line_iterator = 1
            else:
                excel_new_worksheet = False
                line_iterator += 1
        if excel_new_worksheet == True:
            wrap_format = wb.add_format({'text_wrap': True}) # permet de sauter des lignes dans une cellule
            cell_format_title = wb.add_format({'bold': True})
            cell_format_title.set_font_size(13)
            cell_format_title.set_align('center')
            cell_format_title.set_align('vcenter')
            cell_format_normal = wb.add_format({'italic': True})
            cell_format_normal.set_font_size(11)
            cell_format_normal.set_align('center')
            cell_format_normal.set_align('vcenter')
            worksheet.set_column('A:A', 40)
            worksheet.write(0,0,"Songs", cell_format_title)
            worksheet.set_column('B:B', 40)
            worksheet.write(0,1,"Video status", cell_format_title)
            worksheet.set_column('C:C', 40)
            worksheet.write(0,2,"Video size", cell_format_title)
            worksheet.set_column('D:D', 40)
            worksheet.write(0,3,"jpg status", cell_format_title)
        worksheet.write(line_iterator, 0, item[1], cell_format_normal)
        worksheet.write(line_iterator, 1, item[2], cell_format_normal)
        worksheet.write(line_iterator, 2, item[3], cell_format_normal)
        worksheet.write(line_iterator, 3, item[4], cell_format_normal)
        # while entry_list[item][0]
        # songs_list = listCreation("{}\\{}".format(path, folder))

        # iterator = 1
        # for songs in songs_list:
        #     # print(folder, songs)
        #     worksheet.write(iterator, 0, songs)
        #     iterator += 1
    wb.close()
    return

def guiCreation():
    app = QApplication([])
    label = QLabel('Hello World!')
    label.show()
    app.exec()  

def main():
    # main_path = "C:\\Users\\gastoul\\Vocaluxe_0.4.1_Windows_x64\\Songs"
    main_path = "\\\\GOGO-MS451JT\\songs1"
    excel_name = 'songs_list.xlsx'
    folder_list = ()
    song_list = ()
    os.chdir(main_path)
    if os.path.exists(excel_name):
        os.remove(excel_name) 
    folder_list = listCreation(main_path)
    all_songs_list = sorter(main_path, folder_list)
    excelManagement(main_path, all_songs_list, excel_name)

if __name__ == "__main__":
    main()