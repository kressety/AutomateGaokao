from os.path import abspath
from re import search
from sqlite3 import connect, OperationalError
from sys import argv

from win32com.client import DispatchEx

Subject = {
    0: ['文科-本科统招', 'wenke_global'],
    1: ['理科-本科统招', 'like_global'],
    2: ['文科-地方专项', 'wenke_local'],
    3: ['理科-地方专项', 'like_local']
}

Client = DispatchEx('Excel.Application')
Client.Visible = False
Client.DisplayAlerts = False

GaokaoDB = connect('gaokaoDB')
GaokaoDBBuffer = {ID: [] for ID in Subject}
for ThisID in Subject:
    for Back in ['', '_spGroups']:
        GaokaoDBBuffer[ThisID].append(
            [DataHead[0] for DataHead in GaokaoDB.execute('select 代号 from {}'.format(Subject[ThisID][1] + Back))]
        )
    GaokaoDBBuffer[ThisID].append(
        [(DataHead[0], DataHead[1]) for DataHead in GaokaoDB.execute(
            'select 代号, 所属专业组 from {}_sps'.format(Subject[ThisID][1])
        )]
    )
print('缓存读取完毕')

try:
    Year = int(argv[1])
    assert Year >= 2021
    print('已设定数据年份：{}'.format(Year))
except Exception:
    Year = input('命令行未输入参数或参数有误，请手动输入数据年份：')

for ThisID in Subject:

    Book = Client.Workbooks.Open(abspath('data/' + Subject[ThisID][0] + '.xlsx'), ReadOnly=False)
    Sheet = Book.Worksheets('{}-{}'.format(ThisID, Subject[ThisID][0]))
    print('正在打开{}工作表'.format(Subject[ThisID][0]))

    try:
        GaokaoDB.execute("alter table {} add '{}计划数' integer".format(Subject[ThisID][1], Year))
        GaokaoDB.execute("alter table {}_spGroups add '{}计划数' integer".format(Subject[ThisID][1], Year))
        GaokaoDB.execute("alter table {}_spGroups add '{}分数线' integer".format(Subject[ThisID][1], Year))
        GaokaoDB.execute("alter table {}_sps add '{}计划数' integer".format(Subject[ThisID][1], Year))
        GaokaoDB.commit()
        print('已在{}数据库中建立{}年的记录'.format(Subject[ThisID][0], Year))
    except OperationalError:
        GaokaoDB.commit()
        print('数据条目已建立，跳过')

    Row = 1
    CollegeID = None
    GroupID = None
    RowHead = Sheet.Range('A' + str(Row)).Text
    while len(RowHead) > 1:
        Name = Sheet.Range('B' + str(Row)).Text.strip('★').strip('▲')
        if len(RowHead) == 2:
            Name = search('^[^\(（\[【\{]*', Name).group()
            try:
                Number = int(Sheet.Range('C' + str(Row)).Value)
            except ValueError:
                Number = 0
            Time = Sheet.Range('D' + str(Row)).Text
            Fee = Sheet.Range('E' + str(Row)).Text
            if (RowHead, GroupID) not in GaokaoDBBuffer[ThisID][2]:
                GaokaoDB.execute(
                    "insert into {}_sps (代号, 所属专业组, 所属院校, 专业名称, 学制, 学费, '{}计划数') values ('{}', '{}', '{}', '{}', '{}', '{}', {})"
                    .format(Subject[ThisID][1], Year, RowHead, GroupID, CollegeID, Name, Time, Fee, Number)
                )
                print('[{}] 数据库存入行...{}'.format(Subject[ThisID][0], Row))
            else:
                GaokaoDB.execute(
                    "update {}_sps set 学制='{}', 学费='{}', '{}计划数'={} where 代号='{}' and 所属专业组='{}'"
                    .format(Subject[ThisID][1], Time, Fee, Year, Number, RowHead, GroupID)
                )
                print('[{}] 数据库更新行...{}'.format(Subject[ThisID][0], Row))
        elif len(RowHead) == 4:
            CollegeID = RowHead
            try:
                Number = int(Sheet.Range('C' + str(Row)).Value)
            except ValueError:
                Number = 0
            if CollegeID not in GaokaoDBBuffer[ThisID][0]:
                GaokaoDB.execute(
                    "insert into {} (代号, 院校名称, '{}计划数') values ('{}', '{}', {})"
                    .format(Subject[ThisID][1], Year, CollegeID, Name, Number)
                )
                print('[{}] 数据库存入行...{}'.format(Subject[ThisID][0], Row))
            else:
                GaokaoDB.execute(
                    "update {} set '{}计划数'={} where 代号={}"
                    .format(Subject[ThisID][1], Year, Number, CollegeID)
                )
                print('[{}] 数据库更新行...{}'.format(Subject[ThisID][0], Row))
        elif len(RowHead) == 6:
            try:
                Limitation = search('[\(（][(不限)(化学)(生物)(地理)(历史)(思想政治)].*?[\)）]', Name).group()
                Name = search('.*?专业组([\(（].*?[\)）]){,2}', Name).group().replace(Limitation, '')
                Limitation = Limitation[1: -1]
            except AttributeError:
                Limitation = search('[\(（][(不限)(化学)(生物)(地理)(历史)(思想政治)].*', Name).group()
                Name = search('.*?专业组([\(（].*?[\)）]){,2}', Name).group().replace(Limitation, '')
                Limitation = Limitation[1:]
            GroupID = RowHead
            try:
                Number = int(Sheet.Range('C' + str(Row)).Value)
            except ValueError:
                Number = 0
            try:
                Line = int(Sheet.Range('F' + str(Row)).Value)
            except ValueError:
                Line = 0
            if GroupID not in GaokaoDBBuffer[ThisID][1]:
                GaokaoDB.execute(
                    "insert into {}_spGroups (代号, 所属院校, 专业组名称, 限制, '{}计划数', '{}分数线') values ('{}', '{}', '{}', '{}', {}, {})"
                    .format(Subject[ThisID][1], Year, Year, GroupID, CollegeID, Name, Limitation, Number, Line)
                )
                print('[{}] 数据库存入行...{}'.format(Subject[ThisID][0], Row))
            else:
                GaokaoDB.execute(
                    "update {}_spGroups set 限制='{}', '{}计划数'={}, '{}分数线'={} where 代号='{}'"
                    .format(Subject[ThisID][1], Limitation, Year, Number, Year, Line, GroupID)
                )
                print('[{}] 数据库更新行...{}'.format(Subject[ThisID][0], Row))
        Row += 1
        RowHead = Sheet.Range('A' + str(Row)).Text
        GaokaoDB.commit()

    Book.Close()

Client.Quit()
GaokaoDB.close()
