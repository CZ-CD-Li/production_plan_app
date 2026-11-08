Sub 导出控制计划()
Application.ScreenUpdating = False   '禁止屏幕刷新
'Application.EnableEvents = False  '禁止,禁止交互模式.当我们运行程序时,如删除有数据的工作表而不需要提示,
Application.DisplayAlerts = False  '禁止提示
Application.Calculation = xlCalculationManual   '计算模式为手动

    Dim lastRow As Long
    Dim lastCol As Long
    Dim cell As Range


    Dim flCol As Long, gctxCol As Long, xmCol As Long, yqCol As Long, pcCol As Long, ffCol As Long, jlbgCol As Long, fyjhCol As Long, pd As Long
    Dim xms As Byte, ll As Byte
    Dim gxmc As String, gxbh As String, sbmc As String, sbxh As String
    
    Dim j As Long, jj As Long 'j全局数组数据序号
    Dim rng As Range ' 查找区域*
    Dim 区域范围总列数 As Byte '查询区域范围总列数*
    Dim ws As Worksheet ' *
    Dim 制造条件起始行号 As Long ' 制造条件确认起始行号*
    Dim 制造条件起始列号 As Long ' 制造条件确认起始列号*
    Dim 制造条件所占列数 As Long ' 制造条件确认所占列数*
    Dim 制造条件所在单元格A As Range ' 制造条件所在单元格*
    Dim 制造条件项目单元格区域 As Range ' 查询制造条件单元格项目单元格区域*
    Dim 品质特性所在单元格A As Range ' 品质特性所在单元格*
    Dim 品质特性起始行号 As Long ' 品质特性确认起始行号*
    Dim 品质特性起始列号 As Long ' 品质特性确认起始列号*
    Dim 品质特性所占列数 As Long ' 品质特性确认所占列数*
    Dim 安全要求A As Range ' 安全要求所在单元格*
    
    Dim gzb As Range '工作表单元格范围*
    Dim gx As Variant '存储工序要查找的信息名称如工序名称等*
    Dim gxm(0 To 3) As Variant '存储工序查找后的信息名称如工序名称等*
    Dim PC() As String ', PC() As String '临时存放频次数据
    Dim i As Long, L As Long, hhss As Long, mm As Long, gys As Long, m1 As Long, gyss As Long, gzs As Long, ljs As Long, fs As Long, fcs As Long, fcls As Long
    
    Dim btm As Variant '表头项目名存储分类，过程特性，项目，频次等名称*
    Dim btl(0 To 7) As Long '表头项目存储分类，过程特性，项目，频次等所在列*
    Dim hh As Byte, mh As Byte 'hh为循环调用值，mh为作业指导数最大行*
    Dim hpd(1 To 70) As Byte, lzs As Byte '定义判定空行存储数组*
    Dim kzjhbt As Variant '标准控制计划表头项目名存储分类，过程特性，项目，频次等名称*
    Dim kzjhz(0 To 400, 0 To 17) As Variant  '存储控制计划值*
    Dim kzjhl(0 To 400, 0 To 17) As Long  '存储控制合并单元格行数*
    Dim xm(1 To 70, 1 To 8) As Byte '存储单作业指导数内容
    Dim fcl(0 To 70) As Long '存储防错列
    Dim xssy(0 To 5) As Long, xssys As Long '存储检验增加的型试实验3项目
    Dim bdm As Variant  '存储控制计划值*
    Dim bdh As Variant  '存储控制计划值*
    j = 0
    fcls = 0 '辅助记列
    gx = Array("工序名称", "工序编号", "设备名称", "设备型号") '给工序查询项目数组付值*
    btm = Array("分类", "过程特性", "项目", "要求", "频次", "方法", "记录表格", "反应计划") '给工序查询项目数组付值*
    '标准控制计划表头为13个项目+4个分类项目共17个*
    kzjhbt = Array("*工序编号", "*工序名称", "*机器工装量具辅材", "*序号", "*产品特性", "*过程特性", "*特殊特性分类", "*产品过程要求", "*评价测量技术", "*数量", "*频率", "*控制方法", "*反应计划", "*分类", "*担当", "*维护保养", "*防错管理", "*记录表单编号")
    bdm1 = Array("过程质量记录表", "制造条件变更记录表", "检验记录表", "生产线工位照明检查记录表", "理化拉压力试验报告单", "防错") '给工序查询项目数组付值*
    bdm = Array("过程质量", "制造条件", "检验", "照明", "拉压力", "防错") '给工序查询项目数组付值*
    bdh = Array("8731101", "8722010", "8731102", "8731102", "871022102", "8722201") '给工序查询项目数组付值*


Set wsDest = ThisWorkbook.Sheets("控制计划") ' 设置目标输出工作表名称方便后续调用
        ' 遍历工作簿中的每个工作表
Set tstx = wsDest.Cells(28, 1).Resize(1, 8) '输出模板文件中设定特殊特性主机对应符号
       
For Each ws In ThisWorkbook.Sheets '偏历工作表

ws.UsedRange.Value = ws.UsedRange.Value    '******将公式转换为文本
        ' 跳过目标工作表
If ws.Name <> wsDest.Name Then '只要不是目标输出工作表就执行

    Set gzb = ws.UsedRange '将工作表单元格引用到gzb中方便后续引用
    Set ffff = gzb.Find(What:="品质特性确认", LookIn:=xlValues, LookAt:=xlWhole) '查询工作表是否包含品质特性确认
    
If Not ffff Is Nothing Then ' 判断是否是作业指导书或者是附图

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row '工作表最大行数
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column '工作表最大列数

    '存储工序查找后的信息名称如工序名称等*
    For hh = 0 To UBound(gx)
        gxm(hh) = gzb.Find(What:=gx(hh), LookIn:=xlValues, LookAt:=xlWhole).Offset(0, 1).Value
    Next
    '存储工序查找后的信息名称如工序名称等*
    
    Set 品质特性所在单元格A = gzb.Find(What:="品质特性确认", LookIn:=xlValues, LookAt:=xlWhole) ' 查找制造条件确认单元格*
    品质特性起始行号 = 品质特性所在单元格A.Row ' 查找品质特性确认单元格起始行号*
    品质特性起始列号 = 品质特性所在单元格A.Column ' 查找品质特性确认单元格起始列号*
    品质特性所占列数 = 品质特性所在单元格A.MergeArea.Columns.Count ' 品质特性确认所占列数*
    
    Set 制造条件所在单元格A = gzb.Find(What:="制造条件确认", LookIn:=xlValues, LookAt:=xlWhole) ' 查找制造条件确认单元格*
    If 制造条件所在单元格A Is Nothing Then
        MsgBox "工序：" & gxm(0) & "必须要有制造条件确认及下一行共两行，可以不要制造条件具体项目"
        Exit Sub
    End If
    制造条件起始行号 = 制造条件所在单元格A.Row ' 制造条件确认起始行号*
    制造条件起始列号 = 制造条件所在单元格A.Column ' 制造条件确认起始列号*
    制造条件所占列数 = 制造条件所在单元格A.MergeArea.Columns.Count ' 制造条件确认所占列数*
    
    '判定制造条件要在品质特性之前并且在同一列
    If 品质特性起始行号 < 制造条件起始行号 Then
        MsgBox "工序：" & gxm(0) & "品质特性不能到制造条件前面，请将品质特性调整到制造条件后面"
        Exit Sub
    End If
    If 品质特性起始列号 <> 制造条件起始列号 Then
        MsgBox "工序：" & gxm(0) & "品质特性要和制造条件在同一列，请将品质特性和制造条件调整到同一列"
        Exit Sub
    End If
    '判定制造条件要在品质特性之前并且在同一列
    
    '存储查找项目所在列*
    Set 制造条件项目单元格区域 = ws.Cells(制造条件起始行号 + 1, 品质特性起始列号).Resize(制造条件起始行号 + 1, 品质特性起始列号 + 品质特性所占列数 - 1) ' 查询制造条件单元格项目单元格区域*
    For hh = 0 To UBound(btm)
        btl(hh) = 制造条件项目单元格区域.Find(What:=btm(hh), LookIn:=xlValues, LookAt:=xlWhole).Column
    Next
    '存储查找项目所在列*

    Set 安全要求A = gzb.Find(What:="安全要求", LookIn:=xlValues, LookAt:=xlWhole) ' 查找制造条件确认单元格*
    If 安全要求A Is Nothing Then
        MsgBox "工序：" & gxm(0) & "作业指导数中必须要有要求要求"
        Exit Sub
    End If
    
    mh = 安全要求A.Row ' 查找作业指导数最大行数*如果要设置的话可以设置mh=70为最大行
    If mh < 品质特性起始行号 Then
        MsgBox "工序：" & gxm(0) & "安全要求必须在作业指导书左边最下面"
        Exit Sub
    End If
    
    区域范围总列数 = btl(7) - btl(0) + 1 '查询区域范围总列数*
    Set rng = ws.Cells(1, 1).Resize(mh, btl(7)) '设置查询区域范围*
    
    '统计从项目到反应计划每行单元格空值的数量*
    
    lzs = btl(7) - btl(2) + 1 ' 统计从项目到反应计划区域列总数*

    For hh = 制造条件起始行号 + 2 To mh
        hpd(hh) = Application.WorksheetFunction.CountIf(ws.Range(ws.Cells(hh, btl(2)), ws.Cells(hh, btl(7))), "")  '使用CountIf函数统计各行所需要列空单元格的数量
    Next
    '统计从项目到反应计划每行单元格空值的数量*
    'MsgBox UBound(hpd)

    hhss = 1 '设定工作表非空列数初始值
    '制造条件确认项目填写
    gys = 0
    xu = 1
    m = j '设备参数行数
    gzs = 0 '工装行数总数
    ljs = 0 '量具行数总数
    fs = 0 '辅材行数总数
    gys = 0 '设备参数行数总数
    fcs = 0 '防错项目行总数
    jj = 0
        
    For hh = 制造条件起始行号 + 2 To 品质特性起始行号 - 1 '工艺参数，工装，量具，辅材过程参数
        
    
        If rng.Cells(hh, btl(0)) <> "" Then xmm = rng.Cells(hh, btl(0)) '预先提取项目名称，如参数，工装，量具等

        If hpd(hh) <> lzs Then  'lzs统计从项目到反应计划区域列总数,判断是否该行所有需要列是否都为空
            hhss = hhss + 1 '设定工作表非空列数
            'kzjhz(j, 0) = gxm(1)  '输入工序编号
            'kzjhz(j, 1) = gxm(0) '输入工序名称
            kzjhz(j, 13) = xmm '输入项目名称，如工艺参数，工装，量具等
            kzjhl(j, 13) = 1
            If rng.Cells(hh, btl(1)) <> "" Then
                If InStr(rng.Cells(hh, btl(1)), "/") > 0 Then
                    kzjhz(j, 6) = rng.Cells(hh, btl(1)) '输入过程特性
                ElseIf InStr(rng.Cells(hh, btl(1)), "Z") > 0 Or InStr(rng.Cells(hh, btl(1)), "z") > 0 Then
                    If tstx(1, 1) <> "" Then
                        kzjhz(j, 6) = tstx(1, 1) & "/" & rng.Cells(hh, btl(1)) '输入过程特性
                    Else
                        kzjhz(j, 6) = rng.Cells(hh, btl(1))  '输入过程特性
                    End If
                ElseIf InStr(rng.Cells(hh, btl(1)), "G") > 0 Or InStr(rng.Cells(hh, btl(1)), "g") > 0 Then
                    If tstx(1, 2) <> "" Then
                        kzjhz(j, 6) = tstx(1, 2) & "/" & rng.Cells(hh, btl(1)) '输入过程特性
                    Else
                        kzjhz(j, 6) = rng.Cells(hh, btl(1))  '输入过程特性
                    End If
                End If
            End If
            
            'kzjhl(j, 3) = 1 '记录表单合并单元格合并行数
            'kzjhl(j, 4) = 1 '记录表单合并单元格合并行数
            'kzjhl(j, 5) = 1 '记录表单合并单元格合并行数
            'kzjhl(j, 6) = 1 '记录表单合并单元格合并行数
            
            If rng.Cells(hh, btl(2)) <> "" Then '输入工序序号
                kzjhz(j, 3) = xu '输入工序序号
                kzjhl(j, 3) = 1 '
                kzjhl(j, 4) = 1 '
                kzjhl(j, 5) = 1 '
                kzjhl(j, 6) = 1 '
                xu = xu + 1
            End If
            
            
            
            If xmm = "工艺参数" Then '主要针对作业指导书分类，过程特性，项目
                If gxm(3) = "——" Or gxm(3) = "--" Then gxm(3) = ""
                If gys = 0 Then kzjhz(j, 2) = gxm(2) & gxm(3) '输入设备名称
                kzjhz(j, 5) = rng.Cells(hh, btl(2)) '输入参数项目
                If rng.Cells(hh, btl(2)) <> "" Then
                End If
                
                gys = gys + 1 '工艺参数行数
                
                If rng.Cells(hh, btl(2)).MergeCells And rng.Cells(hh, btl(2)).MergeArea.Cells(1, 1).Row = hh Then ' 如果是合并单元格并且为合并单元格首行
                    ls = rng.Cells(hh, btl(2)).MergeArea.Rows.Count ' 获取合并单元格的行数
                    For ll = 2 To ls
                        If hpd(hh + ll - 1) = lzs Then ls = ls - 1
                    Next
                    kzjhl(j, 5) = ls '工艺参数中项目合并单元格合并行数
                    kzjhl(j, 6) = ls '工艺参数中项目过程特性合并单元格合并行数
                    kzjhl(j, 4) = ls '工艺参数中产品特性合并单元格合并行数
                    kzjhl(j, 3) = ls '工艺参数中序号合并单元格合并行数
                End If
            Else
                If j - 1 >= 0 Then
                If kzjhz(j - 1, 13) = "工艺参数" Then
                    If InStr(gxm(2), ".") > 0 Or InStr(gxm(2), "数片机") > 0 Or InStr(gxm(2), "拧紧") > 0 Or InStr(gxm(2), "工作缸") > 0 Or InStr(gxm(2), "示功") > 0 Or InStr(gxm(0), "阻尼力") > 0 Or InStr(gxm(2), "端盖") > 0 Or InStr(gxm(2), "装簧") > 0 Then
                    kzjhz(j, 5) = "防错" '工艺参数中项目
                    kzjhz(j, 7) = "按《防错项目管理清单》执行" '工艺参数中项目
                    kzjhl(j, 7) = 1 '工艺参数中项目
                    kzjhz(m, 16) = "《防错项目管理清单》" '防错项目文件名称
                    kzjhz(j, 12) = "反应计划(5)" '工艺参数中项目
                    kzjhl(j, 12) = 1 '工艺参数中项目
                    kzjhz(j + 1, 3) = xu '工艺参数中序号
                    kzjhz(j, 2) = gxm(2) '工艺参数中项目
                    kzjhl(j, 2) = 1 '工艺参数中项目
                    kzjhz(j, 13) = "防错" '工艺参数中项目
                    kzjhl(j, 13) = 1
                    kzjhz(j, 14) = "操作者"
                    kzjhz(j, 15) = "——" '设备维护
                    kzjhz(j, 16) = "——" '防错项目文件名称
                    kzjhz(j, 17) = "8722201" '录表单编号
                    
                    kzjhl(j, 14) = 1
                    kzjhl(j, 15) = 1 '设备维护
                    kzjhl(j, 16) = 1 '防错项目文件名称
                    kzjhl(j, 17) = 1 '录表单编号
                    fcl(fcls) = j
                    fcs = 1
                    j = j + 1
                    jj = jj + 1
                    xu = xu + 1
                    fcls = fcls + 1
                    Else
                        kzjhz(m, 16) = "——" '防错项目文件名称
                    End If
                    End If
                End If
                
                If rng.Cells(hh, btl(2)) <> "" And xmm <> "工艺参数" Then kzjhl(j, 2) = 1
                
                If xmm = "工装" Then
                    kzjhz(j, 5) = "外观"
                    kzjhz(j, 13) = "工装"
                    kzjhl(j, 13) = 1
                    gzs = gzs + 1
                End If
                If xmm = "量具" Then
                    kzjhz(j, 5) = "有效性"
                    ljs = ljs + 1
                End If
                If xmm = "辅材" Then
                    kzjhz(j, 5) = "型号规格"
                    fs = fs + 1
                End If
                kzjhz(j, 2) = rng.Cells(hh, btl(2)) '输入工装量具辅材项目名称
                
                If rng.Cells(hh, btl(2)).MergeCells And rng.Cells(hh, btl(2)).MergeArea.Cells(1, 1).Row = hh Then ' 如果是合并单元格并且为合并单元格首行
                    ls = rng.Cells(hh, btl(2)).MergeArea.Rows.Count ' 获取合并单元格的行数
                    For ll = 2 To ls
                        If hpd(hh + ll - 1) = lzs Then ls = ls - 1
                    Next
                    kzjhl(j, 2) = ls '工艺参数中项目合并单元格合并行数
                    kzjhl(j, 6) = ls '工艺参数中项目过程特性合并单元格合并行数
                    kzjhl(j, 4) = ls '工艺参数中产品特性合并单元格合并行数
                    kzjhl(j, 3) = ls '工艺参数中序号合并单元格合并行数
                    kzjhl(j, 5) = ls '工艺参数中项目合并单元格合并行数
                End If
                
            End If
            
            
            If rng.Cells(hh, btl(3)) <> "" Then kzjhl(j, 7) = 1
            kzjhz(j, 7) = rng.Cells(hh, btl(3)) '输入要求
                If rng.Cells(hh, btl(3)).MergeCells And rng.Cells(hh, btl(3)).MergeArea.Cells(1, 1).Row = hh Then ' 如果是合并单元格并且为合并单元格首行
                    ls = rng.Cells(hh, btl(3)).MergeArea.Rows.Count ' 获取合并单元格的行数
                    For ll = 2 To ls
                        If hpd(hh + ll - 1) = lzs Then ls = ls - 1
                    Next
                    kzjhl(j, 7) = ls '要求中项目合并单元格合并行数
                End If
            If rng.Cells(hh, btl(5)) <> "" Then kzjhl(j, 8) = 1
            kzjhz(j, 8) = rng.Cells(hh, btl(5)) '输入方法
                If rng.Cells(hh, btl(5)).MergeCells And rng.Cells(hh, btl(5)).MergeArea.Cells(1, 1).Row = hh Then ' 如果是合并单元格并且为合并单元格首行
                    ls = rng.Cells(hh, btl(5)).MergeArea.Rows.Count ' 获取合并单元格的行数
                    For ll = 2 To ls
                        If hpd(hh + ll - 1) = lzs Then ls = ls - 1
                    Next
                    kzjhl(j, 8) = ls '方法中项目合并单元格合并行数
                End If
            If rng.Cells(hh, btl(6)) <> "" Then
                kzjhz(j, 14) = "操作者"
                kzjhz(j, 11) = rng.Cells(hh, btl(6)) '输入记录表格
                kzjhz(j, 17) = "——"
                For xx = 0 To UBound(bdh)
                    If InStr(kzjhz(j, 11), bdm(xx)) > 0 Then
                        kzjhz(j, 17) = bdh(xx) '输入录表单编号和记录表格对应，输入表格名称方便后续替换
                    End If
                Next

                kzjhl(j, 14) = 1
                kzjhl(j, 11) = 1
                kzjhl(j, 17) = 1
                If InStr(1, gxm(0), "检验", vbTextCompare) > 0 Then '输入担当操作者或检验
                    kzjhz(j, 14) = "检验员"
                End If
            End If
                If rng.Cells(hh, btl(6)).MergeCells And rng.Cells(hh, btl(6)).MergeArea.Cells(1, 1).Row = hh Then ' 如果是合并单元格并且为合并单元格首行
                    ls = rng.Cells(hh, btl(6)).MergeArea.Rows.Count ' 获取合并单元格的行数
                    For ll = 2 To ls
                        If hpd(hh + ll - 1) = lzs Then ls = ls - 1
                    Next
                    kzjhl(j, 11) = ls '记录表格中项目合并单元格合并行数
                    kzjhl(j, 17) = ls '录表单编号中项目合并单元格合并行数
                    kzjhl(j, 14) = ls '担当操作者合并单元格合并行数
                    
                    If InStr(1, gxm(0), "检验", vbTextCompare) > 0 Then '输入担当操作者或检验
                        kzjhz(j, 14) = "检验员"

                    End If
                    
                End If
                
            If rng.Cells(hh, btl(7)) <> "" Then kzjhl(j, 12) = 1
            kzjhz(j, 12) = rng.Cells(hh, btl(7)) '输入反应计划
                If rng.Cells(hh, btl(7)).MergeCells And rng.Cells(hh, btl(7)).MergeArea.Cells(1, 1).Row = hh Then ' 如果是合并单元格并且为合并单元格首行
                    ls = rng.Cells(hh, btl(7)).MergeArea.Rows.Count ' 获取合并单元格的行数
                    For ll = 2 To ls
                        If hpd(hh + ll - 1) = lzs Then ls = ls - 1
                    Next
                    kzjhl(j, 12) = ls '记录表格中项目合并单元格合并行数
                End If
                
            If rng.Cells(hh, btl(4)) <> "" Then
                kzjhl(j, 9) = 1
                kzjhl(j, 10) = 1
            End If
            If InStr(1, rng.Cells(hh, btl(4)), "/", vbTextCompare) > 0 Then 'rng.Cells(hh, btl(4)) <> "" Then
                PC = Split(rng.Cells(hh, btl(4)), "/") '将频次内容按/拆分后存储在PC中
                kzjhz(j, 9) = PC(0) '输入数量/频次中数量
                kzjhz(j, 10) = PC(1) '输入频率/频次中频率
            ElseIf rng.Cells(hh, btl(4)) <> "" Then
                kzjhz(j, 9) = rng.Cells(hh, btl(4))
            
            End If
                If rng.Cells(hh, btl(4)).MergeCells And rng.Cells(hh, btl(4)).MergeArea.Cells(1, 1).Row = hh Then ' 如果是合并单元格并且为合并单元格首行
                    ls = rng.Cells(hh, btl(4)).MergeArea.Rows.Count ' 获取合并单元格的行数
                    For ll = 2 To ls
                        If hpd(hh + ll - 1) = lzs Then ls = ls - 1
                    Next
                    kzjhl(j, 9) = ls '数量/频次中数量表格中项目合并单元格合并行数
                    kzjhl(j, 10) = ls '频率/频次表格中项目合并单元格合并行数
                End If
                
                
            j = j + 1
            jj = jj + 1
        End If


    Next
    If gys <> 0 Then
        kzjhl(m, 2) = gys '设备名称也是参数名称合并单元格合并行数
    End If
    If gzs <> 0 Then
        kzjhl(m + gys + fcs, 15) = gzs  '工装数在设备维护保养列总数
        kzjhz(m + gys + fcs, 15) = "——"  '工装数在设备维护保养列值
        kzjhl(m + gys + fcs, 16) = gzs  '工装数在防错列总数
        kzjhz(m + gys + fcs, 16) = "——"  '工装数在防错列值
    End If
    If ljs <> 0 Then
        kzjhl(m + gys + fcs + gzs, 15) = ljs '量具数在设备维护保养列总数
        kzjhz(m + gys + fcs + gzs, 15) = "——" '量具数在设备维护保养列值
        kzjhl(m + gys + fcs + gzs, 16) = ljs '量具数在防错列总数
        kzjhz(m + gys + fcs + gzs, 16) = "——" '量具数在防错列值
    End If
    If fs <> 0 Then
        kzjhl(m + gys + fcs + gzs + ljs, 15) = fs '辅材数在设备维护保养列总数
        kzjhz(m + gys + fcs + gzs + ljs, 15) = "——" '辅材数在设备维护保养列值
        kzjhl(m + gys + fcs + gzs + ljs, 16) = fs '辅材数在防错列总数
        kzjhz(m + gys + fcs + gzs + ljs, 16) = "——" '辅材数在防错列值
    End If
    
    kzjhl(m, 16) = gys '日产防错合并单元格合并行数
    
    If InStr(gxm(0), "活塞") > 0 Or InStr(gxm(0), "底阀") > 0 Or InStr(gxm(0), "固定环") > 0 Or InStr(gxm(0), "拧紧") > 0 Or InStr(gxm(0), "工作缸") > 0 Or InStr(gxm(0), "清洗") > 0 Or InStr(gxm(0), "注油") > 0 Or InStr(gxm(0), "力值") > 0 Or InStr(gxm(0), "充气") > 0 Or InStr(gxm(0), "口") > 0 Or InStr(gxm(0), "端盖") > 0 Or InStr(gxm(2), "装簧") > 0 Or InStr(gxm(0), "装后") > 0 Or InStr(gxm(0), "定位垫") > 0 Or InStr(gxm(0), "套管") > 0 Or InStr(gxm(0), "打码") > 0 Then
        kzjhl(m, 15) = gys '设备维护保养行数；gys代表设备参数行数
        kzjhz(m, 15) = "《设备点检、润滑基准书》" '设备维护保养表单名称

    End If
    
    
    
    gyss = 0 '产品品质总行数
    
    m1 = j
    
    For hh = 品质特性起始行号 + 2 To mh  '品质确认
        
    
        If rng.Cells(hh, btl(0)) <> "" Then xmm = "品质确认" '预先提取项目名称，如参数，工装，量具等

        If hpd(hh) <> lzs Then  'And rng(.Cells(hh, btl(0))) <> ""
            hhss = hhss + 1 '设定工作表非空列数
            'kzjhz(j, 0) = gxm(1)  '输入工序编号
            'kzjhz(j, 1) = gxm(0) '输入工序名称
            kzjhz(j, 13) = xmm '输入项目名称，如工艺参数，工装，量具等
            kzjhl(j, 13) = 1

            If rng.Cells(hh, btl(1)) <> "" Then
                If InStr(rng.Cells(hh, btl(1)), "/") > 0 Then
                    kzjhz(j, 6) = rng.Cells(hh, btl(1)) '输入过程特性

                ElseIf InStr(rng.Cells(hh, btl(1)), "Z") > 0 Or InStr(rng.Cells(hh, btl(1)), "z") > 0 Then
                    If tstx(1, 3) <> "" Then
                        kzjhz(j, 6) = tstx(1, 3) & "/" & rng.Cells(hh, btl(1)) '输入过程特性
                    Else
                        kzjhz(j, 6) = rng.Cells(hh, btl(1))  '输入过程特性
                    End If
                ElseIf InStr(rng.Cells(hh, btl(1)), "G") > 0 Or InStr(rng.Cells(hh, btl(1)), "g") > 0 Then
                    If tstx(1, 4) <> "" Then
                        kzjhz(j, 6) = tstx(1, 4) & "/" & rng.Cells(hh, btl(1)) '输入过程特性
                    Else
                        kzjhz(j, 6) = rng.Cells(hh, btl(1))  '输入过程特性
                    End If
                End If
            End If
            
            'kzjhl(j, 3) = 1 '记录表单合并单元格合并行数
            'kzjhl(j, 4) = 1 '记录表单合并单元格合并行数
            'kzjhl(j, 5) = 1 '记录表单合并单元格合并行数
            'kzjhl(j, 6) = 1 '记录表单合并单元格合并行数


            If rng.Cells(hh, btl(2)) <> "" Then '输入工序序号
                kzjhz(j, 3) = xu '输入工序序号
                kzjhl(j, 3) = 1 '输入工序序号
                kzjhl(j, 4) = 1 '输入工序序号
                kzjhl(j, 5) = 1 '
                kzjhl(j, 6) = 1
                xu = xu + 1
            End If
            
            gyss = gyss + 1 '工艺参数行数
            kzjhz(j, 4) = rng.Cells(hh, btl(2)) '输入产品品质项目

                
            If rng.Cells(hh, btl(2)).MergeCells And rng.Cells(hh, btl(2)).MergeArea.Cells(1, 1).Row = hh Then ' 如果是合并单元格并且为合并单元格首行
                ls = rng.Cells(hh, btl(2)).MergeArea.Rows.Count ' 获取合并单元格的行数
                For ll = 2 To ls
                    If hpd(hh + ll - 1) = lzs Then ls = ls - 1
                Next
                kzjhl(j, 5) = ls '工艺参数中项目合并单元格合并行数
                kzjhl(j, 6) = ls '工艺参数中项目过程特性合并单元格合并行数
                kzjhl(j, 4) = ls '工艺参数中产品特性合并单元格合并行数
                kzjhl(j, 3) = ls '工艺参数中序号合并单元格合并行数
            End If

            If rng.Cells(hh, btl(3)) <> "" Then kzjhl(j, 7) = 1
            kzjhz(j, 7) = rng.Cells(hh, btl(3)) '输入要求
                If rng.Cells(hh, btl(3)).MergeCells And rng.Cells(hh, btl(3)).MergeArea.Cells(1, 1).Row = hh Then ' 如果是合并单元格并且为合并单元格首行
                    ls = rng.Cells(hh, btl(3)).MergeArea.Rows.Count ' 获取合并单元格的行数
                    For ll = 2 To ls
                        If hpd(hh + ll - 1) = lzs Then ls = ls - 1
                    Next
                    kzjhl(j, 7) = ls '要求中项目合并单元格合并行数
                End If
            If rng.Cells(hh, btl(5)) <> "" Then kzjhl(j, 8) = 1
            kzjhz(j, 8) = rng.Cells(hh, btl(5)) '输入方法
                If rng.Cells(hh, btl(5)).MergeCells And rng.Cells(hh, btl(5)).MergeArea.Cells(1, 1).Row = hh Then ' 如果是合并单元格并且为合并单元格首行
                    ls = rng.Cells(hh, btl(5)).MergeArea.Rows.Count ' 获取合并单元格的行数
                    For ll = 2 To ls
                        If hpd(hh + ll - 1) = lzs Then ls = ls - 1
                    Next
                    kzjhl(j, 8) = ls '方法中项目合并单元格合并行数
                End If
            If rng.Cells(hh, btl(6)) <> "" Then
                kzjhz(j, 14) = "操作者"
                kzjhz(j, 11) = rng.Cells(hh, btl(6)) '输入记录表格
                kzjhz(j, 17) = "——"
                For xx = 0 To UBound(bdh)
                    If InStr(kzjhz(j, 11), bdm(xx)) > 0 Then
                        kzjhz(j, 17) = bdh(xx) '输入录表单编号和记录表格对应，输入表格名称方便后续替换
                    End If
                Next
                kzjhl(j, 14) = 1
                kzjhl(j, 11) = 1
                kzjhl(j, 17) = 1
                If InStr(1, gxm(0), "检验", vbTextCompare) > 0 Then '输入担当操作者或检验
                    kzjhz(j, 14) = "检验员"
                End If
            End If
                If rng.Cells(hh, btl(6)).MergeCells And rng.Cells(hh, btl(6)).MergeArea.Cells(1, 1).Row = hh Then ' 如果是合并单元格并且为合并单元格首行
                    ls = rng.Cells(hh, btl(6)).MergeArea.Rows.Count ' 获取合并单元格的行数
                    For ll = 2 To ls
                        If hpd(hh + ll - 1) = lzs Then ls = ls - 1
                    Next
                    kzjhl(j, 11) = ls '记录表格中项目合并单元格合并行数
                    kzjhl(j, 17) = ls '录表单编号中项目合并单元格合并行数
                    kzjhl(j, 14) = ls '担当操作者合并单元格合并行数
                    If InStr(1, gxm(0), "检验", vbTextCompare) > 0 Then '输入担当操作者或检验
                        kzjhz(j, 14) = "检验员"

                    End If
                End If
            If rng.Cells(hh, btl(7)) <> "" Then kzjhl(j, 12) = 1
            kzjhz(j, 12) = rng.Cells(hh, btl(7)) '输入反应计划
                If rng.Cells(hh, btl(7)).MergeCells And rng.Cells(hh, btl(7)).MergeArea.Cells(1, 1).Row = hh Then ' 如果是合并单元格并且为合并单元格首行
                    ls = rng.Cells(hh, btl(7)).MergeArea.Rows.Count ' 获取合并单元格的行数
                    For ll = 2 To ls
                        If hpd(hh + ll - 1) = lzs Then ls = ls - 1
                    Next
                    kzjhl(j, 12) = ls '记录表格中项目合并单元格合并行数
                End If
                
            If rng.Cells(hh, btl(4)) <> "" Then
                kzjhl(j, 9) = 1
                kzjhl(j, 10) = 1
            End If
            If InStr(1, rng.Cells(hh, btl(4)), "/", vbTextCompare) > 0 Then 'rng.Cells(hh, btl(4)) <> "" Then
                PC = Split(rng.Cells(hh, btl(4)), "/") '将频次内容按/拆分后存储在PC中
                kzjhz(j, 9) = PC(0) '输入数量/频次中数量
                kzjhz(j, 10) = PC(1) '输入频率/频次中频率
            ElseIf rng.Cells(hh, btl(4)) <> "" Then
                kzjhz(j, 9) = rng.Cells(hh, btl(4))
            
            End If
                If rng.Cells(hh, btl(4)).MergeCells And rng.Cells(hh, btl(4)).MergeArea.Cells(1, 1).Row = hh Then ' 如果是合并单元格并且为合并单元格首行
                    ls = rng.Cells(hh, btl(4)).MergeArea.Rows.Count ' 获取合并单元格的行数
                    For ll = 2 To ls
                        If hpd(hh + ll - 1) = lzs Then ls = ls - 1
                    Next
                    kzjhl(j, 9) = ls '数量/频次中数量表格中项目合并单元格合并行数
                    kzjhl(j, 10) = ls '频率/频次表格中项目合并单元格合并行数
                End If
                
                
                If hh = mh Then
                    If InStr(gxm(0), "检验") > 0 And InStr(gxm(0), "总成") > 0 Then
                    j = j + 1
                    jj = jj + 1
                    gyss = gyss + 1
                    'kzjhz(j, 0) = gxm(1)  '输入工序编号
                    'kzjhz(j, 1) = gxm(0) '输入工序名称
                    kzjhz(j, 4) = "型式试验" '工艺参数中项目
                    kzjhz(j, 7) = "型式试验计划及报告" '工艺参数中项目
                    kzjhl(j, 7) = 1
                    kzjhz(j, 12) = "参照《型式试验、测试失败或检查不通过时的应对》" '工艺参数中项目
                    kzjhl(j, 12) = 4 '工艺参数中项目
                    kzjhz(j, 3) = xu '工艺参数中序号
                    kzjhz(j, 13) = xmm '工艺参数中项目
                    kzjhl(j, 13) = 1
                    kzjhz(j, 14) = "检验员"
                    kzjhz(j, 17) = "——"
                    xssys = 0
                    xssy(xssys) = j
                    xssys = xssys + 1
            kzjhl(j, 3) = 1 '记录表单合并单元格合并行数
            kzjhl(j, 4) = 1 '记录表单合并单元格合并行数
            kzjhl(j, 5) = 1 '记录表单合并单元格合并行数
            kzjhl(j, 6) = 1 '记录表单合并单元格合并行数
                    j = j + 1
                    jj = jj + 1
                    xu = xu + 1
                    gyss = gyss + 1
                    'kzjhz(j, 0) = gxm(1)  '输入工序编号
                    'kzjhz(j, 1) = gxm(0) '输入工序名称                    kzjhz(j, 5) = "产品审核" '工艺参数中项目
                    kzjhz(j, 4) = "产品审核" '工艺参数中项目
                    kzjhz(j, 7) = "产品审核计划" '工艺参数中项目
                    kzjhl(j, 7) = 1
                    kzjhz(j, 3) = xu '工艺参数中序号
                    kzjhz(j, 13) = xmm '工艺参数中项目
                    kzjhl(j, 13) = 1
                    kzjhz(j, 14) = "检验员"
                    kzjhz(j, 17) = "——"
                    xssy(xssys) = j
                    xssys = xssys + 1
            kzjhl(j, 3) = 1 '记录表单合并单元格合并行数
            kzjhl(j, 4) = 1 '记录表单合并单元格合并行数
            kzjhl(j, 5) = 1 '记录表单合并单元格合并行数
            kzjhl(j, 6) = 1 '记录表单合并单元格合并行数
                    j = j + 1
                    jj = jj + 1
                    xu = xu + 1
                    gyss = gyss + 1
                    'kzjhz(j, 0) = gxm(1)  '输入工序编号
                    'kzjhz(j, 1) = gxm(0) '输入工序名称                    kzjhz(j, 5) = "全尺寸检测" '工艺参数中项目
                    kzjhz(j, 4) = "全尺寸检测" '工艺参数中项目
                    kzjhz(j, 7) = "全尺寸检测计划及报告" '工艺参数中项目
                    kzjhl(j, 7) = 1
                    kzjhz(j, 3) = xu '工艺参数中序号
                    kzjhz(j, 13) = xmm '工艺参数中项目
                    kzjhl(j, 13) = 1
                    kzjhz(j, 14) = "检验员"
                    kzjhz(j, 17) = "——"
                    xssy(xssys) = j
                    xssys = xssys + 1
            kzjhl(j, 3) = 1 '记录表单合并单元格合并行数
            kzjhl(j, 4) = 1 '记录表单合并单元格合并行数
            kzjhl(j, 5) = 1 '记录表单合并单元格合并行数
            kzjhl(j, 6) = 1 '记录表单合并单元格合并行数
                    j = j + 1
                    jj = jj + 1
                    xu = xu + 1
                    gyss = gyss + 1
                    'kzjhz(j, 0) = gxm(1)  '输入工序编号
                    'kzjhz(j, 1) = gxm(0) '输入工序名称                    kzjhz(j, 5) = "全尺寸检测" '工艺参数中项目
                    kzjhz(j, 4) = "MSA分析" '工艺参数中项目
                    kzjhz(j, 7) = "MSA分析计划及报告" '工艺参数中项目
                    kzjhl(j, 7) = 1
                    kzjhz(j, 3) = xu '工艺参数中序号
                    kzjhz(j, 13) = xmm '工艺参数中项目
                    kzjhl(j, 13) = 1
                    kzjhz(j, 14) = "检验员"
                    kzjhz(j, 17) = "——"
                    xssy(xssys) = j
                    xssys = xssys + 1
            kzjhl(j, 3) = 1 '记录表单合并单元格合并行数
            kzjhl(j, 4) = 1 '记录表单合并单元格合并行数
            kzjhl(j, 5) = 1 '记录表单合并单元格合并行数
            kzjhl(j, 6) = 1 '记录表单合并单元格合并行数

                    
                    Else
                        kzjhz(m, 16) = "——" '防错项目文件名称
                    End If
                End If
            'If hh = mh + 1 Then Exit For
                
            j = j + 1
            jj = jj + 1
        End If


    Next

    kzjhl(m1, 2) = gyss '产品品质总行数
    kzjhl(m1, 16) = gyss '防错合并单元列项目格合并行数
    kzjhl(m1, 15) = gyss '设备维护列项目保养行数；
    kzjhz(m1, 16) = "——" '防错合并单元列项目内容
    kzjhz(m1, 15) = "——" '设备维护列项目内容；
    
    kzjhz(m, 0) = gxm(1)  '输入工序编号
    kzjhz(m, 1) = gxm(0) '输入工序名
    kzjhl(m, 0) = jj '工序编号合并单元格合并行数
    kzjhl(m, 1) = jj '工序名称合并单元格合并行数


    ' 判断是否所有单元格都为空
    'If emptyCount = rng.Cells.Count Then
    
    

    'End If
        'hh = hh + 1

      
    ' 输出结果
    
End If
End If
Next ws

'j = j - 1


Dim zdls As Long, cxwz As Long, cx As Range ' 查找区域*
Dim sckzjhz() As Variant  '存储控制计划值*
Dim sckzjhl() As Long  '存储控制合并单元格行数*
Dim hgkzjhl() As Long  '存储控制合并单元格行数*
Dim hgz() As Long '存储行高值
Dim kzjhs() As Long   '存储字符串长度*
Dim lss() As Long  '存储列合并数*
Dim lk() As Long  '存储项目列宽*
Dim rn As Range, zdz As Single, hb As Single  'zdz存储非合并单元格对比最大值,hb合并单元格临时存储

cxwz = 30 '设置查询行所在行数
zdls = wsDest.Cells(cxwz, wsDest.Columns.Count).End(xlToLeft).Column '查询列数
zhls = 1

ReDim sckzjhz(0 To j, 0 To zdls - 1) ' 使用ReDim关键字来分配数组的大小
ReDim sckzjhl(0 To j, 0 To zdls - 1) ' 使用ReDim关键字来分配数组的大小
ReDim hgkzjhl(0 To j, 0 To zdls - 1) ' 使用ReDim关键字来分配数组的大小
ReDim hgz(0 To j)  '存储行高值
ReDim kzjhs(0 To j, 0 To zdls - 1) As Long   '存储字符串长度*
ReDim lss(0 To zdls - 1) ' 合并单元格列数
ReDim lk(0 To zdls - 1) '存储项目列宽*

Set cx = wsDest.Cells(cxwz, 1).Resize(1, zdls)

For i = 0 To zdls - 1
    Set rng = cx.Cells(1, i + 1)
    Set rng1 = cx.Cells(1, i + 1).MergeArea
    If rng <> "" Then
        For ii = 0 To 17 '查询项目与kzjhbt数组中存储的项目对应位置
            If rng = "*产品过程要求" Then fcqsl = rng.Column '识别按《防错项目管理清单》执行合并起始列
            If rng = "*反应计划" Then fcjsl = rng.Column - 1 '识别按《防错项目管理清单》执行合并结束列
            If rng = kzjhbt(ii) Then
                pd = 1
                For jj = 0 To j
                    If kzjhz(jj, ii) <> "" Then kzjhs(jj, i) = Len(kzjhz(jj, ii))  '存储字符串长度*
                    sckzjhz(jj, i) = kzjhz(jj, ii) 'sckzjhz(jj, i) = kzjhz(jj, ii) & "*" & kzjhl(jj, ii)
                    sckzjhl(jj, i) = kzjhl(jj, ii)
                    If sckzjhl(jj, i) = 1 Then '将合并单元格中合并行数放到合并单元格最后的单元格中
                        hgkzjhl(jj, i) = kzjhl(jj, ii)
                    ElseIf sckzjhl(jj, i) > 1 Then
                        hgkzjhl(jj + kzjhl(jj, ii) - 1, i) = kzjhl(jj, ii)
                        kzjhs(jj + kzjhl(jj, ii) - 1, i) = kzjhs(jj, i) '存储字符串长度*
                    End If
                Next
                Exit For
            End If
        Next
            'lss(i) = IsMergeFirstColumn(rng) '查询合并单元格列数
        If rng.MergeCells And rng.Address = rng.MergeArea.Cells(1, 1).Address Then
            lss(i) = rng1.Columns.Count
            k = 0
            For kk = 0 To lss(i) - 1
                k = k + cx.Cells(1, i + 1 + kk).ColumnWidth
            Next
            lk(i) = k
            'MsgBox i + 1 & "列数" & lss(i) & "列宽" & lk(i)
        Else
            lss(i) = 1
            lk(i) = rng.ColumnWidth
            'MsgBox i + 1 & "列数" & lss(i) & "列宽" & lk(i)
        End If
    
    End If
Next
      'wsDest.Range("A26:AK26") = lk'输出列宽
    'MsgBox "hpd: " & kzjhz(0, 5)
    'wsDest.Range("a31:o431") = kzjhz 'wsDest.Range("a31:q431") = kzjhz
    wsDest.Range(Cells(cxwz + 1, 1), Cells(cxwz + 1 + j, zdls)) = sckzjhz '输出控制计划，wsDest.Range("a31:q431") = kzjhz
    'wsDest.Range("a30:q430") = kzjhl
    For i = 0 To j '合并控制计划单元格
        For L = 0 To zdls - 1 '        For L = 0 To 14
            If sckzjhl(i, L) > 0 Then
                Range(Cells(i + cxwz + 1, L + 1), Cells(i + cxwz + sckzjhl(i, L), L + lss(L))).Merge
            End If
        Next
    Next
    For i = 0 To fcls - 1
        Range(Cells(fcl(i) + cxwz + 1, fcqsl), Cells(fcl(i) + cxwz + 1, fcjsl)).Merge '合并防错
    Next
    For i = 0 To xssys - 1
        Range(Cells(xssy(i) + cxwz + 1, fcqsl), Cells(xssy(i) + cxwz + 1, fcjsl)).Merge '合并检验形式实验项目
    Next

    Dim g As Single, gy As Single, bs As Single, kh As Single, bss As Single, qsbss As Single, zzss As Single '存储项目列宽*
    g = 11 '以字体为10时单行字正好11为1行
    gy = 4 '以字体为10时单行字边缘余高
    kh = 1.75 '以字体为10时10个中字为一行时列宽为17.5，则单个字为1.75
    bs = 1 ' 合并单元格预留倍数余量
    bss = 1.4 ' 超过限定行数合并单元格预留倍数余量
    qsbss = 5 ' 限定行数
    For i = 0 To j '调整行高以值体高度10作业依据,存储非合并单元格行数最大值
        zdz = 0
        For L = 0 To zdls - 1 '        For L = 0 To 14
            If sckzjhz(i, L) <> "" And sckzjhl(i, L) <= 1 Then
                    If zdz < kzjhs(i, L) * kh / lk(L) Then
                        If kzjhs(i, L) * kh / lk(L) >= qsbss Then
                            zdz = kzjhs(i, L) * kh / lk(L) * bss
                        Else
                            zdz = kzjhs(i, L) * kh / lk(L) * bs
                        End If

                    End If
            End If
        Next
        If zdz > (400 - gy) / g Then zdz = (400 - gy) / g
        hgz(i) = Application.WorksheetFunction.Ceiling(zdz, 1)
        'MsgBox i + 31 & "行数：" & hgz(i)
    Next
    For i = 0 To j '调整行高以值体高度10作业依据，合并单元格参与对比
        For L = 0 To zdls - 1 '        For L = 0 To 14
            If hgkzjhl(i, L) > 1 Then
                dddd = sckzjhz(i, L)
                dddd22 = sckzjhz(i - 1, L)
                If kzjhs(i, L) * kh / lk(L) * bs > hgkzjhl(i, L) Then
                    zdz = 0
                    For jj = 0 To hgkzjhl(i, L) - 1
                        zdz = zdz + hgz(i - jj)
                    Next
                    'MsgBox i + 31 & "内容：" & fffff & "行数：" & zdz
                    hb = 0
                    If zdz < kzjhs(i, L) * kh / lk(L) Then
                        If kzjhs(i, L) * kh / lk(L) >= qsbss Then
                            zzss = bss
                        Else
                            zzss = bs
                        End If
                        hb = Application.WorksheetFunction.Ceiling(kzjhs(i, L) * kh / lk(L), 1) * zzss / hgkzjhl(i, L)
                        If hb > (400 - gy) / g Then hb = (400 - gy) / g
                        hb = Application.WorksheetFunction.Ceiling(hb, 1)
                        For jj = 0 To hgkzjhl(i, L) - 1
                            hgz(i - jj) = hb
                        Next
                    End If
                End If
            
            End If
        Next
    Next
    
    For i = 0 To j '调整行高以值体高度10作业依据，合并单元格参与对比
        Range("A" & i + cxwz + 1).EntireRow.RowHeight = hgz(i) * g + gy
    Next
    
    For i = j + cxwz + 1 To 430 '调整行高以值体高度10作业依据，合并单元格参与对比
            Range("A" & i).EntireRow.RowHeight = g + gy
    Next

    

        Cells.Replace What:="本工艺卡片", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        Cells.Replace What:="工艺卡片", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    'Range(Cells(j + cxwz + 1 + 5, 1), Cells(420, 200)).Select '删除值后多余孔表行
    'Selection.Delete Shift:=xlUp '删除值后多余孔表行
        
    '删除前部分多余行
    Set gzb = wsDest.UsedRange  '将工作表单元格引用到gzb中方便后续引用
    Set qshh = gzb.Find(What:="**起始行", LookIn:=xlValues, LookAt:=xlWhole) ' 查找起始单元格*
    If qshh Is Nothing Then
        MsgBox "无" & """" & "  **起始行  " & """" & "字符标定删除起始行，前多余行请手动删除"
        Exit Sub
    Else
        jj = qshh.Row ' 制造条件确认起始行号*
        'Range(Cells(jj, 1), Cells(30, 200)).Delete  '删除值后多余孔表行
        Rows(jj & ":" & 30).Delete
        'Selection.Delete Shift:=xlUp
    End If
    '删除前部分多余
    Rows(j + cxwz + 1 - 31 + jj + 2 & ":" & 430 - 31 + jj - 2).Delete '删除值后多余孔表行
    'Range(Cells(j + cxwz + 1 + 5 - 31 + jj, 1), Cells(420 - 31 + jj, 200)).Select '删除值后多余孔表行
    'Selection.Delete Shift:=xlUp '删除值后多余孔表行
    
Application.Calculation = xlCalculationAutomatic
Application.DisplayAlerts = True  '禁止提示
Application.ScreenUpdating = True


'Set dd = ws.Range("l77:t77")
'dd.Delete
End Sub