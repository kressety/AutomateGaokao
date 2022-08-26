from os import walk
from os.path import abspath
from re import match, search

import pdfplumber
from win32com.client import DispatchEx

Subject = {
    0: '文科-本科统招',
    1: '理科-本科统招',
    2: '文科-地方专项',
    3: '理科-地方专项'
}

Client = DispatchEx('Excel.Application')
Client.Visible = False
Client.DisplayAlerts = False


def StringCombine(List):
    return '|{:<6}|{:<50}|{:<4}|{:<2}|{:<6}|'.format(List[0], List[1], List[2], List[3], List[4])


def ManualPage(P, S, PDF, PL, R=1):
    if R == 1:
        NewBook = Client.Workbooks.Add()
        NewBook.Worksheets.Add().Name = '{}-{}'.format(S, Subject[S])
        NewBook.SaveAs(abspath('data/' + Subject[S] + '.xlsx'))
    ExcelBooks = Client.Workbooks.Open('data/' + abspath(Subject[S] + '.xlsx'), ReadOnly=False)
    ExcelSheet = ExcelBooks.Worksheets('{}-{}'.format(S, Subject[S]))
    ManualContent = PDF.pages[P].extract_tables()
    for ManualIndex in range(len(ManualContent)):
        print()
        print('\033[34m{}: {}\033[0m'.format(ManualIndex, StringCombine(ManualContent[ManualIndex][1])))
        print('......')
    print()
    IndexSelected = input('请选择属于{}的子表格（用空格隔开，对应页数{}）：'.format(Subject[S], P + 1)).split(' ')
    while len(set([str(i) for i in list(range(len(ManualContent)))] + IndexSelected)) > len(ManualContent):
        IndexSelected = input('输入有误，请重新输入：'.format(Subject[S], P)).split(' ')
    for ContentIndex in IndexSelected:
        for ManualTable in ManualContent[int(ContentIndex)][1:]:
            ExcelSheet.Range('A{}:E{}'.format(R, R)).Select()
            Client.Selection.NumberFormatLocal = "@"
            ExcelSheet.Range('A{}:E{}'.format(R, R)).Value = ManualTable
            R += 1
    print('[{}] 页{}已写入({:.2%})'.format(Subject[S], P + 1, (
            P - PL[2 * S] + 1) / (
            PL[2 * S + 1] - PL[2 * S] + 1
    )))
    ExcelBooks.Save()
    ExcelBooks.Close()
    return R


if __name__ == '__main__':
    RegularFiles = []
    for _, _, Files in walk('data'):
        for File in Files:
            try:
                if File[File.rindex('.'):] == '.pdf':
                    RegularFiles.append(File)
            except ValueError:
                continue
        break
    for RegularIndex in range(len(RegularFiles)):
        print('{}: {}'.format(RegularIndex, RegularFiles[RegularIndex]))
    FileSelected = input('请选择PDF：')
    while FileSelected not in [str(i) for i in range(len(RegularFiles))]:
        FileSelected = input('输入有误，请重新选择：')
    PDFFile = pdfplumber.open('data/' + RegularFiles[int(FileSelected)])
    Catalogue = PDFFile.pages[int(input('请输入目录页码：')) - 1].extract_text().split('\n')
    Target = '二.*本.*科.*院.*校'
    PageList = []
    for ItemIndex in range(len(Catalogue)):
        if match(Target, Catalogue[ItemIndex]) is not None:
            PageList.append(int(search('[0-9]+', Catalogue[ItemIndex]).group()) + 3)
            PageList.append(int(search('[0-9]+', Catalogue[ItemIndex + 1]).group()) + 3)
    for LocalIndex in [2, 3]:
        PageRange = input('请输入{}的页码范围（例如12-15, 14-14）：'.format(Subject[LocalIndex])).split('-')
        while True:
            try:
                for i in range(2):
                    if (int(PageRange[i]) < 1) or (int(PageRange[i]) > len(PDFFile.pages)):
                        PageRange = input('输入有误，请重新输入：').split('-')
                        continue
                else:
                    break
            except IndexError and ValueError:
                PageRange = input('输入有误，请重新输入：').split('-')
        PageRange = [int(i) - 1 for i in PageRange]
        PageRange.sort()
        PageList.append(PageRange[0])
        PageList.append(PageRange[1])

    for SubjectIndex in Subject:
        ThisRow = ManualPage(PageList[2 * SubjectIndex], SubjectIndex, PDFFile, PageList)
        if PageList[2 * SubjectIndex + 1] - PageList[2 * SubjectIndex] >= 2:
            ThisBooks = Client.Workbooks.Open(abspath('data/' + Subject[SubjectIndex] + '.xlsx'), ReadOnly=False)
            ThisSheet = ThisBooks.Worksheets('{}-{}'.format(SubjectIndex, Subject[SubjectIndex]))
            for PageIndex in range(PageList[2 * SubjectIndex] + 1, PageList[2 * SubjectIndex + 1]):
                PageContent = PDFFile.pages[PageIndex].extract_tables()
                for Content in PageContent:
                    for Table in Content[1:]:
                        ThisSheet.Range('A{}:E{}'.format(ThisRow, ThisRow)).Select()
                        Client.Selection.NumberFormatLocal = "@"
                        ThisSheet.Range('A{}:E{}'.format(ThisRow, ThisRow)).Value = Table
                        ThisRow += 1
                print('[{}] 页{}已写入({:.2%})'.format(Subject[SubjectIndex], PageIndex + 1, (
                        PageIndex - PageList[2 * SubjectIndex] + 1) / (
                        PageList[2 * SubjectIndex + 1] - PageList[2 * SubjectIndex] + 1
                )))
            ThisBooks.Save()
            ThisBooks.Close()
        if PageList[2 * SubjectIndex + 1] - PageList[2 * SubjectIndex] >= 1:
            ManualPage(PageList[2 * SubjectIndex + 1], SubjectIndex, PDFFile, PageList, ThisRow)
    PDFFile.close()

    print('开始整理表格...')
    for SubjectIndex in Subject:
        ThisBooks = Client.Workbooks.Open('data/' + abspath(Subject[SubjectIndex] + '.xlsx'), ReadOnly=False)
        ThisSheet = ThisBooks.Worksheets('{}-{}'.format(SubjectIndex, Subject[SubjectIndex]))
        ThisBooks.Worksheets('Sheet1').Select()
        Client.ActiveWindow.SelectedSheets.Delete()
        print('[{}] 已删除空白工作表'.format(Subject[SubjectIndex]))
        Row = 1
        while Row <= ThisSheet.UsedRange.Rows.Count:
            if (ThisSheet.Range('A' + str(Row)).Value is None) or (
                    ThisSheet.Range('A' + str(Row)).Value == '') or (
                    ThisSheet.Range('A' + str(Row)).Value == 0):
                ThisSheet.Range('B' + str(Row - 1)).Value += ThisSheet.Range('B' + str(Row)).Value
                ThisSheet.Rows(Row).Delete()
                continue
            print('[{}] 正在合并空白行...{:.2%}'.format(Subject[SubjectIndex], Row / ThisSheet.UsedRange.Rows.Count))
            Row += 1
        print('[{}] 空白行合并完毕'.format(Subject[SubjectIndex]))
        ThisSheet.Cells.Select()
        ThisSheet.Cells.EntireColumn.AutoFit()
        Client.Selection.RowHeight = 13.9
        print('[{}] 行高列宽调整完毕'.format(Subject[SubjectIndex]))
        for Row in range(1, ThisSheet.UsedRange.Rows.Count + 1):
            NewString = ''
            for Piece in ThisSheet.Range('B' + str(Row)).Text.split('\n'):
                NewString += Piece
            ThisSheet.Range('B' + str(Row)).Value = NewString
            print('[{}] 正在删除换行符...{:.2%}'.format(Subject[SubjectIndex], Row / ThisSheet.UsedRange.Rows.Count))
        print('[{}] 换行符删除完毕'.format(Subject[SubjectIndex]))
        ReStandard = [4, 6, 2]
        for Row in range(1, ThisSheet.UsedRange.Rows.Count + 1):
            for Level in range(len(ReStandard)):
                if len(ThisSheet.Range('A' + str(Row)).Text) == ReStandard[Level]:
                    for i in range(Level):
                        ThisSheet.Rows(Row).Select()
                        Client.Selection.Rows.Group()
                    break
            print('[{}] 正在三级分类...{:.2%}'.format(Subject[SubjectIndex], Row / ThisSheet.UsedRange.Rows.Count))
        print('[{}] 三级分类完毕'.format(Subject[SubjectIndex]))
        ThisBooks.Save()
        ThisBooks.Close()

    Client.Quit()
    input('已完成, 点按回车以继续...')
