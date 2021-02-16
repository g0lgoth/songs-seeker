import os
import svn.remote
import xlsxwriter
import time

# global variable
GENERAL_PATH = "E:/UltraStar Deluxe/songs"
# on obtient le chemin d'où le code est lancé
# il faut lancer le programme du fichier songs dans Ultrastar
# GENERAL_PATH = os.path.abspath(os.getcwd())

# création de l'excel avec la mise en forme spécifique
def excel_creation():
    # je me place dans le dossier ou je veux lister les éléments
    excel_directory_move = os.chdir(GENERAL_PATH)
    # je vérifie que le fichier que je veux créer n'existe pas
    # s'il existe je le supprime, sinon je le créé
    if os.path.exists("songs_summary.xlsx"):
        os.remove("songs_summary.xlsx")
    wb = xlsxwriter.Workbook('songs_summary.xlsx')
    # je fais une liste des éléments dans le dossier
    karaoke_folder_list = os.listdir("./")
    for element in karaoke_folder_list:
        worksheet = wb.add_worksheet(element)
        wrap_format = wb.add_format({'text_wrap': True}) # permet de sauter des lignes dans une cellule
        cell_format_title = wb.add_format({'bold': True})
        cell_format_title.set_font_size(13)
        cell_format_title.set_align('center')
        cell_format_title.set_align('vcenter')
        worksheet.set_column('A:A', 50)
        worksheet.write(0,0,"Song title", cell_format_title)
        worksheet.set_column('B:B', 10)
        worksheet.write(0,1,"mp3", cell_format_title)
        worksheet.set_column('C:C', 10)
        worksheet.write(0,2,"avi", cell_format_title)
        worksheet.set_column('D:D', 10)
        worksheet.write(0,3,"jpg", cell_format_title)
        worksheet.set_column('E:E', 10)
        worksheet.write(0,4,"duo", cell_format_title)
        worksheet.set_column('F:F', 10)
        worksheet.write(0,5,"size", cell_format_title)
        worksheet.set_column('G:G', 50)
        worksheet.write(0,6,"list of files", cell_format_title)
        cell_format_song = wb.add_format({})
        cell_format_song.set_align('vcenter')
        cell_format_OK = wb.add_format({})
        cell_format_NOK = wb.add_format({})
        cell_format_OK.set_bg_color('green')
        cell_format_OK.set_align('center')
        cell_format_OK.set_align('vcenter')
        cell_format_NOK.set_align('center')
        cell_format_NOK.set_align('vcenter')
        cell_format_NOK.set_bg_color('red')
        index_song = 1
        for folder in karaoke_folder_list:
            karaoke_song_list = []
            os.chdir("{}/{}/".format(GENERAL_PATH, folder))
            karaoke_song_list = os.listdir("./")
            for song in karaoke_song_list:
                worksheet.write(index_song, 0, song)
                for root, dirnames, filenames in os.walk("{}/{}/".format(GENERAL_PATH, folder)):
                    for file in filenames:
                        os.chdir("{}/".format(root))
                    try:
                        # mp3 = [True for file in filenames if ".mp3" in filenames]
                        # if mp3:
                        if path.exist("{}.mp3".format(song)):
                            worksheet.write(index_song, 1, "OK", cell_format_OK)
                        else:

                            worksheet.write(index_song, 1, "NOK", cell_format_NOK)
                    except:
                        worksheet.write(index_song, 1, "NOK", cell_format_NOK)
                index_song += 1
    os.chdir("{}/".format(GENERAL_PATH))
    wb.close()
    return wb, worksheet, wrap_format, karaoke_folder_list

excel_creation()