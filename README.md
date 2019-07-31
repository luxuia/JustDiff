# ExcelMerge

[![Build status](https://ci.appveyor.com/api/projects/status/ven9wb4k5wrrajub?svg=true)](https://ci.appveyor.com/project/luxuia/excelmerge)

Features:

    自带svn日志，方便看各个版本的差异
    预先对excel的行列按内容排序，忽略行位置变化导致的修改
    可以查看已打开的文档
    默认找到第一个有差异的sheet
    背景色表示差异类型，黄色[修改]，灰色[删除]，绿色[新增]
    单格内红色文本表示差异，删除线删除，下划线新增

![demo](demo.jpg)


使用方式:

打开应用后，拖拽一个对比文件A到左边。
 
    1.如果文件A在SVN管理下，可以在上边的<版本>下拉框中选择近两个月的修改记录。选中后直接显示这次修改的差异
    2.拖拽另一个文件B到右侧，点击上边 <对比>按钮。显示两个文件的差异行

TortoiseSVN软件可以设置默认的Diff工具:

Settings->Diff Viewer -> Advanced.. -> 选择xls，修改-> 执行程序替换成 "[软件安装路径]\bin\Release\ExcelMerge.exe" %base %mine
