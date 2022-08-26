from sqlite3 import connect

from pyecharts import options as opts
from pyecharts.charts import Pie, Timeline, Page

StringMap = {
    'wenke': '文科',
    'like': '理科'
}

TimeLine = Timeline()

GaokaoDB = connect('gaokaoDB')

TotalDetail = {}
SpGroupsDetail = {}
SpsDetail = {}
for Subject in StringMap:
    CollegeData = GaokaoDB.execute('select "2021计划数", "2022计划数" from {}_global where 代号="1101"'.format(Subject))
    for Data in CollegeData:
        TotalDetail[Subject] = [Data[0], Data[1]]
    SpGroupsData = GaokaoDB.execute(
        'select 代号, 专业组名称, 限制, "2021计划数", "2021分数线", "2022计划数", "2022分数线" from {}_global_spGroups where 所属院校="1101"'
        .format(Subject)
    )
    SpGroupsDetail[Subject] = {}
    SpsDetail[Subject] = {}
    for Data in SpGroupsData:
        Data = list(Data)
        SpsDetail[Subject][Data[0]] = {}
        SpGroupsDetail[Subject][Data.pop(0)] = Data
    SpsData = GaokaoDB.execute(
        'select 所属专业组, 专业名称, 学制, 学费, "2021计划数", "2022计划数" from {}_global_sps where 所属院校="1101"'
        .format(Subject)
    )
    for Data in SpsData:
        Data = list(Data)
        SpsDetail[Subject][Data.pop(0)] = Data

for Subject in StringMap:
    DataSet = [(SpGroupsDetail[Subject][i][0], SpGroupsDetail[Subject][i][-2]) for i in SpGroupsDetail[Subject]]
    for Data in DataSet:
        if not Data[1]:
            DataSet.remove(Data)
    Pie_SpGroupsCount = (
        Pie()
        .add(
            '{}专业组2022计划数'.format(StringMap[Subject]),
            DataSet,
            rosetype='radius',
            radius=['30%', '55%'],
        )
        .set_global_opts(title_opts=opts.TitleOpts('{}专业组计划数'.format(StringMap[Subject])))
    )
    TimeLine.add(Pie_SpGroupsCount, StringMap[Subject])

CombinedPage = Page(
    page_title='AutomateGaokao Demo NJU',
    layout=Page.DraggablePageLayout
)
CombinedPage.add(
    TimeLine
)

CombinedPage.render()
