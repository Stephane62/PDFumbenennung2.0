from PyPDF2 import PdfFileReader
import os

Documents = os.path.join("test")

def main():
    temp = open(Documents+"\\sd,nf.pdf", 'rb')
    _pdfRead = PdfFileReader(temp)
    page = _pdfRead.getPage(0)
    text = page.extractText()
    getInfos(text)


def getInfos(text):
    wantedInfos = ""
    for i in range(len(text)):
        if text[i] == 'N':
            if checkIfKey(text[i] + text[i+1] + text[i+2] + text[i+3]):
                _curIndex = i+4
                while text[_curIndex] != '\n':
                    wantedInfos += text[_curIndex]
                    _curIndex += 1
                print(wantedInfos)
                return wantedInfos





def checkIfKey(info):
    return info == "Nr. "


if __name__ == '__main__':
    main()
