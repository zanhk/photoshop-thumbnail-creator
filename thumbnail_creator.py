# 1. Before first execution run following commands from cmd:
# pip install psd-tools
# pip install pypiwin32
# python -m pip install --upgrade pip
# 2. Run script as ADMINISTRATOR
# 3. Follow Instructions

import win32com.client
psApp = win32com.client.Dispatch("Photoshop.Application")

# ----------------------------------------------------------------------------------------
# Edit here the path for the psd and the png (use all path bcuz relative path cause error)
# Remember to include / at the end of PathToPng
# IF YOU COPY AND PASTE BE SURE THAT YOU ARE USING '/' (Forward Slash) AND NOT '\' (Backslash)

PathToPsd = "E:/OneDrive/Progetti/fcamuso/BL3/borderlands_video_cover_720.psd"
PathToThumbnailsFolder = "E:/OneDrive/Progetti/fcamuso/BL3/Thumbnails/"

# ----------------------------------------------------------------------------------------

# Opens a PSD file
psApp.Open(PathToPsd)
doc = psApp.Application.ActiveDocument  # Get active document object

continue_cicle = 1

while (continue_cicle == 1):

    while True:
        try:
            title = input("Insert title between \"\": ")
            if type(title) is str:
                break
            else:
                print("\nInsert a string!\n")
                continue
        except:
            print("\nInsert string between brackets. Example -> \"Video Title\"!\n")
            continue

    while True:
        try:
            number = input("Insert video number: ")
            if type(number) is int:
                break
            else:
                print("\nInsert a number (integer)!\n")
                continue
        except:
            print("\nInsert video number as an Int!\n")
            continue

    title = str(title)
    number = str(number)

    layer_facts = doc.ArtLayers["TITLE"]
    text_of_layer = layer_facts.TextItem
    text_of_layer.contents = title

    layer_facts = doc.ArtLayers["NUMBER"]
    text_of_layer = layer_facts.TextItem
    text_of_layer.contents = "#" + number

    options = win32com.client.Dispatch('Photoshop.ExportOptionsSaveForWeb')
    options.Format = 13   # PNG Format
    options.PNG8 = False  # Sets it to PNG-24 bit

    pngfilePath = PathToThumbnailsFolder + \
        "borderlands3_thumbnail" + "_" + number + ".png"

    doc.Export(ExportIn=pngfilePath, ExportAs=2, Options=options)

    print("\nPNG file saved to " + pngfilePath + "\n")

    while True:
        try:
            continue_cicle = input(
                "Do you want to create a new one? (1=yes / 0=no): ")
            if type(continue_cicle) is int:
                print("\n")
                break
            else:
                print("\nInsert a Integer (0-1)!\n")
                continue
        except:
            print("\nInsert 0 (stop) or 1 (continue)!\n")
            continue


print("\nClosing Photoshop App")
doc.Close(2)
psApp.Quit()
