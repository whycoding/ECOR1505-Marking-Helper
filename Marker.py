import PySimpleGUI as sg 
import xlrd
import os


layoutMain = [[sg.Text("1. select the allocation excel"), sg.Input(key="-Excel-"), sg.FileBrowse(file_types=("*.xlsx"))],
            [sg.Text("2. select the result folder"), sg.Input(key="-Folder-"), sg.FolderBrowse()],
            [sg.Button("Submit")],
            [sg.Text("select the marker"), sg.Drop(values=("None"), key="-Marker-", size=(15, 1)), sg.Text("navigate students", key="-Student-"), sg.Button("Previous"), sg.Button("Next")], 
            [sg.Multiline(size=(50, 10), key="-Contend-")]
]

def ReadExcel(excelName):
    wb = xlrd.open_workbook(excelName)
    sheet = wb.sheet_by_index(0)

    tasks = sheet.col_values(0)
    markers = sheet.col_values(1)

    return tasks, markers
    
def DisplayTask(folderName):
    contend = ""
    if not os.path.exists(folderName):
        contend =  "Error!!! Folder not exist " + folderName
    else:
        for parent, dirnames, filenames in os.walk(folderName):
            for filename in filenames:
                shortname, extension = os.path.splitext(filename)
                if extension != '.py':
                    contend =  "Error!!! Filename not correct " + filename
                with open(os.path.join(parent, filename)) as f:
                    contend = contend + "\n" + "*" * 20
                    contend = contend + "\n" + filename
                    contend = contend + "\n" + "*" * 20
                    contend = contend + "\n" + f.read()

    windowMain["-Student-"].Update(folderName)
    windowMain["-Contend-"].Update(contend)



def ExploreFolder(folder):
    pass

def EventButton(windowMain):
    while True:
        button, value = windowMain.Read()

        if button == None:
            break

        if button == "Submit":
            tasks, markers = ReadExcel(value["-Excel-"])
            folderResult = value["-Folder-"]

            markerList = list(set(markers))
            markerList.sort()

            #windowMain.Element("-Marker-").Update(values=markerList)
            windowMain["-Marker-"].Update(values=markerList)

            #tasksAssigned = [tasks[i] for i in range(len(tasks)) if markers[i] == value["-Marker-"]]
            tasksAssigned = [tasks[i] for i in range(len(tasks)) if markers[i] == "Haoyuan Wei"]

            taskIndex = 0
            folderTask = folderResult + "/" + tasksAssigned[taskIndex]
            DisplayTask(folderTask)

        if button == "Next":
            taskIndex = taskIndex + 1
            folderTask = folderResult + "/" + tasksAssigned[taskIndex]
            DisplayTask(folderTask)

        if button == "Previous":
            taskIndex = taskIndex - 1
            folderTask = folderResult + "/" + tasksAssigned[taskIndex]
            DisplayTask(folderTask)


        



if __name__ == "__main__":
    windowMain = sg.Window("ECOR1505 Marking Helper", layoutMain)

    EventButton(windowMain)

    windowMain.close()