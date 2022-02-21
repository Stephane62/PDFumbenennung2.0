from PyPDF2 import PdfFileReader
import os
import pandas as pd
import openpyxl as op


wrkbk = op.load_workbook('Excel/mappe.xlsx')
sh = wrkbk.active


personalIds = []
Vertragsnummern = []
Names = []



for i in range(2, sh.max_row + 1):
    for j in range(1, sh.max_column + 1):
        cell_obj = sh.cell(row=i, column=j)
        #print(cell_obj.value , " ",i," ", j ,end=' ')
        #list2 = [cell_obj.value]
        #print(list2)
        if j == 1:
            personalIds.append(cell_obj.value)
        elif j == 2:
            Vertragsnummern.append(cell_obj.value)
        elif j == 3:
            Names.append(cell_obj.value)
        elif j == 4:
            Names[i-2] += " " + cell_obj.value


Documents = os.path.join("test")

def main():
    for root,directory,files in os.walk(Documents):
        for my_file in files:
            old_name = Documents + "\\" + my_file
            temp = open(old_name, 'rb')
            _pdfRead = PdfFileReader(temp)
            page = _pdfRead.getPage(0)
            text = page.extractText()
            temp.close()
            _requestedInfo = getInfos(text)
            _organisedData = setDataForList(_requestedInfo)
            info_list = getInfoList(_organisedData)
            client_id = setClientId(info_list[0])
            real_name = reverseName(info_list[1])
            _index_in_Names = Names.index(real_name)
            personalnummer = str(personalIds[_index_in_Names])
            vertragsnummer = str(Vertragsnummern[_index_in_Names])
            new_name = Documents + "\\[" + personalnummer + "-" + vertragsnummer + "-" + real_name + "].pdf"
            if os.path.isfile(new_name):
                print("File with the same name already exists !!!")
            else:
                os.rename(old_name, new_name)
                print("Successfully changed the name of the file with the name : " + my_file)



def getInfos(text):
    wantedInfos = ""
    for i in range(len(text)):
        if text[i] == 'N':
            if checkIfKey(text[i] + text[i+1] + text[i+2] + text[i+3]):
                _curIndex = i+4
                while text[_curIndex] != '\n':
                    wantedInfos += text[_curIndex]
                    _curIndex += 1
                return wantedInfos

def checkIfKey(info):
    return info == "Nr. "

def setDataForList(info):
    Info = ""
    indexOfTrimValue = info.find('-')
    for i in range(len(info)):
        if i == indexOfTrimValue:
            Info += '_'
        elif i < indexOfTrimValue - 1:
            Info += info[i]
        elif i > indexOfTrimValue + 1:
            Info += info[i]
    return Info

def getInfoList(organisedInfo):
    info_list = list()
    curInfo = ""
    for i in range(len(organisedInfo)):
        if organisedInfo[i] != '_':
            curInfo += organisedInfo[i]
            if i == len(organisedInfo) - 1:
                info_list.append(curInfo)
        else:
            info_list.append(curInfo)
            curInfo = ""
    return info_list

def setClientId(nonProcessedId):
    id = ""
    for i in nonProcessedId:
        if i != '/':
            id += i
    return id

def reverseName(name):
    firstname = ""
    familyName = ""
    _spaceIndex = name.find(' ')
    for i in range(len(name)):
        if i < _spaceIndex:
            familyName += name[i]
        elif i > _spaceIndex:
            firstname += name[i]
    return firstname + " " + familyName

if __name__ == '__main__':
    main()
