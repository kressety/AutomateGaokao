from os import walk
from os.path import abspath
from re import match

from win32com.client import DispatchEx

Subject = {
    0: ['文科-本科统招', '本科', '历史'],
    1: ['理科-本科统招', '本科', '物理'],
    2: ['文科-地方专项', '提前', '历史'],
    3: ['理科-地方专项', '提前', '物理']
}

Client = DispatchEx('Excel.Application')
Client.Visible = False
Client.DisplayAlerts = False

LineFiles = []

for _, _, Files in walk('data'):
    for File in Files:
        try:
            if File[File.rindex('.'):] not in ['.xls', 'xlsx']:
                continue
            elif File[: File.rindex('.')] in [Subject[SubjectIndex][0] for SubjectIndex in Subject]:
                continue
            else:
                LineFiles.append(File)
        except ValueError:
            continue

if len(LineFiles) != 4:
    print('分数线文件数目与招生专刊文件数目不等！')
else:
    MatchList = ['.*{}批次.*{}.*'.format(Subject[i].pop(-2), Subject[i].pop(-1)) for i in Subject]
    for File in LineFiles:
        BookLine = Client.Workbooks.Open(abspath('data/' + File))
        SheetLine = BookLine.Worksheets[0]
        Row = 1
        RowHead = SheetLine.Range('A' + str(Row)).Text.replace('\n', '').strip()
        ThisSubjectIndex = None
        while not (RowHead.isdigit() and len(RowHead) == 4):
            for SubjectIndex in Subject:
                if len(Subject[SubjectIndex]) == 1:
                    if match(MatchList[SubjectIndex], RowHead):
                        Subject[SubjectIndex].append(File)
                        ThisSubjectIndex = SubjectIndex
                        print('已将{}与{}匹配'.format(File, Subject[ThisSubjectIndex][0]))
            Row += 1
            RowHead = SheetLine.Range('A' + str(Row)).Text.replace('\n', '').strip()
        if not ThisSubjectIndex:
            raise RuntimeError('没有可与{}匹配的专刊！'.format(File))
        Subject[ThisSubjectIndex].append(str(Row))
        BookLine.Close()

    for SubjectIndex in Subject:
        BookMain = Client.Workbooks.Open(abspath('data/' + Subject[SubjectIndex][0] + '.xlsx'))
        SheetMain = BookMain.Worksheets[0]
        BookLine = Client.Workbooks.Open(abspath('data/' + Subject[SubjectIndex][1]))
        SheetLine = BookLine.Worksheets[0]

        DictLine = {}
        Row = int(Subject[SubjectIndex][2])
        RowHead = SheetLine.Range('A' + str(Row)).Text.replace('\n', '').strip()
        while len(RowHead) == 4:
            Name = SheetLine.Range('B' + str(Row)).Text.replace('\n', '').strip()
            Name = Name[Name.index('专') - 2: Name.index('专')]
            DictLine[RowHead + Name] = SheetLine.Range('C' + str(Row)).Value
            Row += 1
            RowHead = SheetLine.Range('A' + str(Row)).Text.replace('\n', '').strip()
        print('{}查询表建立完成！'.format(Subject[SubjectIndex][0]))

        Row = 1
        RowHead = SheetMain.Range('A' + str(Row)).Text.replace('\n', '').strip()
        while len(RowHead) > 1:
            if RowHead in DictLine:
                SheetMain.Range('F' + str(Row)).Value = DictLine[RowHead]
            Row += 1
            RowHead = SheetMain.Range('A' + str(Row)).Text.replace('\n', '').strip()
        print('{}分数线已全部填入！'.format(Subject[SubjectIndex][0]))

        BookMain.Save()
        BookMain.Close()
        BookLine.Close()

Client.Quit()
