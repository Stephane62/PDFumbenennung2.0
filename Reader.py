from PyPDF2 import PdfFileReader
import os
import pandas as pd


xl = pd.read_excel('test/Excel/mappe.xlsx')
print(xl)
Documents = os.path.join("test")

def main():
    for root,directory,files in os.walk(Documents):
        for my_file in files:
            old_name = Documents + "\\" + my_file
            print(old_name)
            temp = open(old_name, 'rb')
            _pdfRead = PdfFileReader(temp)
            page = _pdfRead.getPage(0)
            text = page.extractText()
            temp.close()
            _requestedInfo = getInfos(text)
            _organisedData = setDataForList(_requestedInfo)
            info_list = getInfoList(_organisedData)
            client_id = setClientId(info_list[0])
            new_name = Documents + "\\[" +client_id + " " + info_list[1] + "].pdf"
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
if __name__ == '__main__':
    main()
