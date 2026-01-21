1、WALLA.mve/厦门市轨道...mve修改Typical导致闪退，报错如下：

File "C:\Users\CNXIGAO13\Desktop\ABB_Production_Tools_\FAST_V1.9\need\MVE_Pre_Configure.py", line 1072, in saveedit
    for i in range(0, sum(int(num) for num in switchgearmountlist)-Panel_count[switchgear_index]):
  File "C:\Users\CNXIGAO13\Desktop\ABB_Production_Tools_\FAST_V1.9\need\MVE_Pre_Configure.py", line 1072, in <genexpr>
    for i in range(0, sum(int(num) for num in switchgearmountlist)-Panel_count[switchgear_index]):
ValueError: invalid literal for int() with base 10: '无此节点'

初步研究发现是站下面的柜体数量信息为[无此节点]，将具体数量配置后，重新点击读取，即不会闪退。
修复思路应该是自动读取柜体数量。暂时的解决方案是提示。

2、上述bug需要解决，同时，为一些功能增加错误追踪，以减少报错闪退问题。
暂时取消，无法过于细致try,except。

3、PR和EBOM做对比，PR中的物料应该全出现在EBOM里面。
已取消该功能。

4、EBOM两个时期差异对比。
已取消该功能。

5、MVE新增Typical。

6、林德.mve删除非柜体站号报错。
已解决，1683行for循环改为for i in range(len(busbarbridge_count)-1, -1, -1):
当存在多个非柜体站号时，应从后面先删除，以避免前面先删除后面索引超出。
range(a, b, -1)表示序列从a到b+1，倒序，这就要求a>b。

7、2023.11.6林锐邮件，503959562-EBOM(EPLAN和SAP对比功能BUG)
已解决，利用span节点来确定同一行的子节点，然后将同一个span节点下子节点.text进行拼接，这样就可以应对同一行内容出现多个子节点
同时解决了站号可能出现三行的问题。

8、管理员权限，PE_DE文件均放置公共盘和本地，当找不到公共盘时，按照本地文件，这样可以在公共盘添加新成员。
已解决，目前按照放在J盘配置。

9、程序放置到Sharepoint站点，用户使用本地程序，当程序更新时，自动同步到本地程序，无需去J盘拷贝。
已实现。

10、document/PE_DE.xlsx中蔡志雄邮箱漏掉姓氏。
已修改。

11、读写数据库时，修改完成，执行一步将数据库复制到sharepoint站点，以方便Irina做数据读取和统计。


12、设计传递表功能优化：表格是否可优化，保证能够拷贝出来。


13、将设计传递表中客户线号，客户特殊需求加入到数据库。

14、设计传递表模板列超出页面问题处理。

15、设计传递表典型柜配置不能超过17行问题需要解决。

16、CT/PT/避雷器等厂家信息加入到设计传递表中。
CT/PT：ABB、DYH、其他
避雷器：神电、ABB

17、Typical 数量写入的时候改成只在第一个item 写入，其余为零。

18、传递表的界面在接受时间还没填写的情况可以允许重复写入，然后能够读取已经写入的数据，方便他们更改。一旦我这边写入接受时间则不允许再通过传递表的界面写入。

19、MVE属性更新手动修改功能中取消界面自动刷新，以避免属性多个属性修改时，因界面刷新问题导致的其他属性写入失败。
已解决。

20、因PMP平台网址变更，设计传递表和项目信息管理功能中涉及到PMP平台网址的地方进行修改。
已解决。

21、500的微动开关判断逻辑。
已解决。

22、500的容性分压装置判断逻辑。
已解决。

23、AIS的低压室照明判断逻辑。
已解决。

24、图纸意见的传递路径触发条件目前是接收时间触发还是启动时间触发？改成接收时间吧。


25、传递表的填写后刷新数据就没有掉，考虑保留最近一次数据，增加清除按钮。（实现难度大可考虑下期升版。）

26、设计检查结果可以输出一份PDF 报告打印审核。





27、Z5报表重新复刻DeviceLabel报表，实现保护LED灯描述非标提醒。

28、Z7报表复刻P-BOM报表，实现微动等自检逻辑。

29、线号文件检查中，增加物料节点线号定位与BOM清单中物料定位对比，差异时，给出提示。

30、Irina与Lily沟通，能否做出一份包含所有物料的接线报表（节点接线情况，未接线情况），用于检查物料接线情况。
31、导EBOM时，跳过“空柜”、含“DUMMY”字样的Typical。














