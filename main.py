import xlsxwriter
from bs4 import BeautifulSoup
import os


def loadMasterFile():
    links = []
    emptyName = []
    print("Enter Location of Master File")
    loc = input()
    if not loc == "":
        f = open(loc, "r")
        lines = f.readlines()
        counter = 0
        for line in lines:
            if line[0] == "#":  # allows for commenting
                pass
            else:
                links.append(line)

                emptyName.append("masterFile{num}?id={id}".format(id=str(line.split("?id=")[1]).rstrip(), num=counter))
                counter += 1
        return links, emptyName
    else:
        return "",""

def cutDown(htmlIn, master):
    labels = []
    dataPure = []
    if master[0][0] != "" and master[0][1] != "":
        labels = master[0][1]
        dataPure = master[0][0]
    dataCut = []
    dataCutRefined=[]
    steamLinks = []
    steamLinksRefined = []
    formattedID = []
    deleteLocationsChars = []
    html = open(htmlIn, "r")
    soup = BeautifulSoup(html, "html.parser")
    table = soup.find("div", {"class": "mod-list"})
    rows = table.find_all("tr")

    for row in rows:
        dataPure.append(str(row.find_all("td")[2].text))
        labels.append(str(row.find_all("td")[0].text))
    i=0
    while i < len(dataPure): #first while loop will replace all dupes with blanks
        if not (str(dataPure[i].split("?id=")[1]).rstrip() in dataCutRefined):
            dataCutRefined.append(str(dataPure[i].split("?id=")[1]).rstrip())
            steamLinksRefined.append(str(dataPure[i]).rstrip().strip("\n"))
        else:

            deleteLocationsChars.append(labels[i])
        i += 1
    i = 0


    for i in deleteLocationsChars:
        if i in labels:
            labels.remove(i)

    for i in range(len(steamLinksRefined)):
        formattedID.append(str("@{ID};".format(ID=dataCutRefined[i])))

    return [labels, dataCutRefined, steamLinksRefined, formattedID]


def appendToExcel(modFiles, htmlFile):
    totalString = ""
    path = str(os.path.dirname(__file__)) + "/modDataSheets/{fHTML}".format(
        fHTML=os.path.basename(htmlFile).split(".")[0] + ".xlsx")
    print("Parsing {HTML}, end file location: {PATH}".format(HTML=htmlFile, PATH=path))
    if not os.path.isfile(path):
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "w") as f:
            f.close()
    workbook = xlsxwriter.Workbook(path)
    worksheet = workbook.add_worksheet()

    worksheet.write('A1', 'Mod Name')
    worksheet.write('B1', 'Mod ID')
    worksheet.write('C1', 'Steam Link')
    worksheet.write('D1', 'Formatted ID')
    worksheet.write('G1', 'Formatted ID One Line')
    row = 1

    for col, data in enumerate(modFiles):
        worksheet.write_column(row, col, data)


    for i in range(len(modFiles[3])):
        totalString = totalString + str(modFiles[3][i])
    worksheet.write('G2', totalString)
    workbook.close()
    print("Process Completed. Press Enter To Continue")
    cont = input()


def menu():
    clearScreen()
    master = []
    master.append(loadMasterFile())
    htmlFile = str(input("Enter full File Name and Path of HTML which you want to extract Mod Name and ID\n"))
    modNameModFile = cutDown(htmlFile, master)
    appendToExcel(modNameModFile, htmlFile)
    menu()


def clearScreen():
    if os.name == "nt":
        os.system('cls')  # windows
    else:
        os.system('clear')  # other (mac/linux)


if __name__ == "__main__":
    menu()
