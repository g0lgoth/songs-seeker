import os
import time
import xlsxwriter

def listCreation(path):
    """
    Créer une liste et la nettoie pour qu'il ne reste que des dossiers
    path = chemin pour création de liste
    """
    new_list = []
    os.chdir(path)
    new_list = os.listdir("./")
    new_list=listCleaner(new_list)
    return new_list

def listCleaner(list_to_clean):
    """
    Prend une liste et enlève les éléments qui ne sont pas un dossier
    fichiers zip, excel etc...
    list_to_clean = liste à nettoyer
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
    return list_ok

def picturePresence(path):
    """
    A un chemin donné vérifie s'il y a un fichier image
    path = chemin à contrôler
    """
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
    """
    A un chemin donné vérifie s'il y a un fichier vidéo
    Si oui retourne sa taille
    path = chemin à contrôler
    """
    new_list = []
    temporary_list = []
    txt_list = []
    nok = "no"
    ok = "yes"
    no_video = "0"
    no_error = 0
    error_value = 1
    try :
        for file in os.listdir(path):
            if file.endswith(".txt"):
                with open("{}{}".format(path, file), 'r') as f:
                    lines = f.readlines()
                for line in lines:
                    # try:
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
                        try:
                            temporary_variable = os.stat("{}{}".format(path, new_line_name))
                            video_size = temporary_variable.st_size / (1024 * 1024)
                            return ok, video_size, no_error
                        except FileNotFoundError:
                            print("error between video value in file and video file")
                            return "no file", no_video, error_value
                    else:
                        continue
                        # else:
                        #     return "no video line in file", no_video, no_error
                    # except TypeError:
                    #     print("erreur dans la recherche de balise #VIDEO")
                    #     return "no video line in file", no_video, error_value                   
            # else :
            #     return "no text file", no_video, no_error 
        return nok, no_video, error_value
    except ValueError:
        print("erreur sur la création de la liste")
        return nok, no_video, error_value     

def duoPresence(path):
    """
    A un chemin donné vérifie dans un fichier texte si certaines lignes sont présentes
    path = chemin à contrôler
    """
    txt_number = 0
    duo_variable = 0
    for file in os.listdir(path):
        if file.endswith(".txt"):
            txt_number += 1
            with open("{}{}".format(path, file), 'r') as f:
                lines = f.readlines()
            for line in lines:
                if "#DUETSINGERP1" in line:
                    duo_variable += 1
                if "#DUETSINGERP2" in line:
                    duo_variable += 1
    if duo_variable == 2 :
        return txt_number, "yes"
    else:
        return txt_number, "no"

def sorter(path, folder_list):
    """
    A un chemin donné créé une liste de liste
    path = chemin ou créer la liste
    folder_list = liste de dossier parents
    """
    final_list = []
    error_counter = 0
    for folder in folder_list:
        songs_list = listCreation("{}\\{}".format(path, folder))
        songs_list = listCleaner(songs_list)
        for song in songs_list:
            temporary_list = []
            temporary_list.append(folder)
            temporary_list.append(song)
            current_location = "{}\\{}\\{}\\".format(path, folder, song)
            video_status, video_size, error_value = clipPresence(current_location)
            error_counter += error_value
            temporary_list.append(video_status)
            temporary_list.append(video_size)
            image_status = picturePresence(current_location)
            temporary_list.append(image_status)
            txt_file_number, duo_status = duoPresence(current_location)
            temporary_list.append(txt_file_number)
            temporary_list.append(duo_status)
            final_list.append(temporary_list)
            print("folder:", folder, "| song name:", song, "\nPath:", current_location, "\n")
    return final_list

def excelInit(path, name_file):
    """
    Vérifie si le fichier existe, si oui le supprime
    path = chemin à contrôler
    name_file = nom du fichier
    """
    os.chdir(path)
    if os.path.exists(name_file):
        os.remove(name_file)
    return

def excelManagement(path, entry_list, name_file):
    """
    A un chemin donné créé un excel et copie les informations de la liste
    path = chemin ou créer le fichier
    entry_list = liste à copier dans le fichier excel
    name_file = nom du fichier excel
    """
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
            worksheet.set_column('A:A', 50)
            worksheet.write(0,0,"Songs", cell_format_title)
            worksheet.set_column('B:B', 30)
            worksheet.write(0,1,"Video status", cell_format_title)
            worksheet.set_column('C:C', 30)
            worksheet.write(0,2,"Video size", cell_format_title)
            worksheet.set_column('D:D', 20)
            worksheet.write(0,3,"jpg status", cell_format_title)
            worksheet.set_column('E:E', 25)
            worksheet.write(0,4,"txt file number", cell_format_title)
            worksheet.set_column('F:F', 20)
            worksheet.write(0,5,"duo status", cell_format_title)
        worksheet.write(line_iterator, 0, item[1], cell_format_normal)
        worksheet.write(line_iterator, 1, item[2], cell_format_normal)
        worksheet.write(line_iterator, 2, item[3], cell_format_normal)
        worksheet.write(line_iterator, 3, item[4], cell_format_normal)
        worksheet.write(line_iterator, 4, item[5], cell_format_normal)
        worksheet.write(line_iterator, 5, item[6], cell_format_normal)
        worksheet.autofilter('A01:F999')
    wb.close()
    return

def main():
    # main_path = "C:\\Users\\gastoul\\Vocaluxe_0.4.1_Windows_x64\\Songs"
    main_path = "\\\\GOGO-MS451JT\\songs1"
    excel_name = 'songs_list.xlsx'
    folder_list = ()
    song_list = ()
    folder_list = listCreation(main_path)
    all_songs_list = sorter(main_path, folder_list)
    excelInit(main_path, excel_name)
    excelManagement(main_path, all_songs_list, excel_name)

if __name__ == "__main__":
    main()