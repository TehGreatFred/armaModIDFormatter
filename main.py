import xlsxwriter
from bs4 import BeautifulSoup
import os


def cutDown(htmlIn):
    labels = []
    dataPure = []
    dataCut = []
    steamLinks = []
    formattedID = []
    html = open(htmlIn, "r")
    soup = BeautifulSoup(html, "html.parser")
    table = soup.find("div", {"class": "mod-list"})
    rows = table.find_all("tr")

    for row in rows:
        labels.append(str(row.find_all("td")[0].text))
        dataPure.append(str(row.find_all("td")[2].text))

    for i in range(len(dataPure)):
        dataCut.append(str(dataPure[i].split("?id=")[1]).rstrip())
        steamLinks.append(str(dataPure[i]))

    for i in range(len(dataCut)):

        formattedID.append(str("@{ID};".format(ID=dataCut[i])))

    return [labels, dataCut, steamLinks,formattedID]


def appendToExcel(modFiles, htmlFile):
    totalString = ""
    path = str(os.path.dirname(__file__)) + "/modDataSheets/{fHTML}".format(fHTML=os.path.basename(htmlFile).split(".")[0] + ".xlsx")
    print("Parsing {HTML}, end file location: {PATH}".format(HTML=htmlFile,PATH=path))
    if not os.path.isfile(path):
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "w") as f:
            f.close()
    workbook = xlsxwriter.Workbook(path)
    worksheet = workbook.add_worksheet()

    worksheet.write('A1', 'Mod Name')
    worksheet.write('B1', 'Mod ID')
    worksheet.write('C1', 'Steam Link')
    worksheet.write('D1','Formatted ID')
    worksheet.write('E1', 'Formatted ID One Line')
    row = 1
    for col, data in enumerate(modFiles):
        worksheet.write_column(row, col, data)

    for i in range(len(modFiles[3])):
        totalString = totalString + str(modFiles[3][i])
    worksheet.write('E2', totalString)
    workbook.close()
    print("Process Completed. Press Enter To Continue")
    cont = input()
def menu():
    clearScreen()
    htmlFile = str(input("Enter File Name of HTML which you want to extract Mod Name and ID\n"))
    modNameModFile = cutDown(htmlFile)
    appendToExcel(modNameModFile, htmlFile)
    menu()

def clearScreen():
    if os.name == "nt":
        os.system('cls') #windows
    else:
        os.system('clear') #other (mac/linux)

if __name__ == "__main__":
    menu()
