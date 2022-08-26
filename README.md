# AutomateGaokao
新高考招生信息全流程自动化
# 各个文件的作用
`./main.py`

从招生专刊.pdf文件中提取信息，写入excel并进行优化

`./lineMix.py`

将分数线混入excel

`./updateDB.py`

使用已经混入分数线的excel更新数据库

`./gaokaoDB`

主数据库，使用sqlite3

`./dataTree.py`

DataTree类，用于数据查询和可视化信息展示

`./visualized.py`

数据可视化，使用pyecharts
# 工作
|项目|进度|
|-|-|
|dataTree|基本完成|
|pdf->excel|部分测试|
|updateDB|部分测试|
|lineMix|部分测试|
|visualize|未完成|
|功能耦合|待开发|
|web部署|待开发|
