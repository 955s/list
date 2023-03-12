Attribute VB_Name = "公用模块1"
Sub 取消所有工作表保护()
        Application.ScreenUpdating = False   '禁刷新
        With ThisWorkbook
            For i = 1 To .Worksheets.Count    '遍历文件的工作表数
                表名 = .Worksheets(i).Name
                If 表名 <> "out" Then '判断不为out的表格 即所以表
                    .Sheets(表名).Unprotect
                End If
            Next i
        End With
End Sub

Sub 保护所有工作表()
        Application.ScreenUpdating = False   '禁刷新
        With ThisWorkbook
            For i = 1 To .Worksheets.Count    '遍历文件的工作表数
                表名 = .Worksheets(i).Name
                If 表名 <> "透视表" Then '判断不为out的表格 即所以表
                    .Sheets(表名).Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
                End If
            Next i
        End With
End Sub

Sub 清空合表()
    Sheets("合").Range("A2:P2000").ClearContents  '清空原始数据，写死out表格，防止清错
    Worksheets("合").Activate
    mbx = 2000
    Rows(2 & ":" & mbx).Delete
End Sub

Sub 清除填充批注()
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.ClearComments
End Sub
Sub 去除线()
    Cells.Select
    Selection.Borders.LineStyle = 0 '去除框线
End Sub

Sub 纯白填充()
    'x = ActiveSheet.[A65535].End(xlUp).Row
    x = 1000
    Range("A2:AZ" & x).Select
    With Selection.Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub

Sub 获取表名粘贴到A列()
    For x = 1 To Sheets.Count
        Cells(x, 1) = Sheets(x).Name
    Next x
End Sub
    
Sub 筛选合表() '筛选合表
    If ActiveSheet.AutoFilterMode = True Then ActiveSheet.AutoFilterMode = False
    i = ActiveSheet.Range("B" & Rows.Count).End(xlUp).Row
    Range("A1:P" & i).Select
    Selection.AutoFilter

End Sub

Sub 筛选批次() '筛选合表
    'If ActiveSheet.AutoFilterMode = True Then ActiveSheet.AutoFilterMode = False
    i = ActiveSheet.Range("B" & Rows.Count).End(xlUp).Row
    Range("A2:P" & i).Select
    Selection.AutoFilter

End Sub

Sub 筛选供货单() '筛选合表
    If ActiveSheet.AutoFilterMode = True Then ActiveSheet.AutoFilterMode = False
    i = ActiveSheet.Range("B" & Rows.Count).End(xlUp).Row
    Range("A1:Q" & i).Select
    Selection.AutoFilter

End Sub

Sub 筛选方案() '筛选方案
    'If ActiveSheet.AutoFilterMode = True Then ActiveSheet.AutoFilterMode = False
    i = ActiveSheet.Range("J" & Rows.Count).End(xlUp).Row
    Range("I1:S" & i).Select
    Selection.AutoFilter

End Sub

Sub 筛选合同() '筛选方案
    'If ActiveSheet.AutoFilterMode = True Then ActiveSheet.AutoFilterMode = False
    i = ActiveSheet.Range("T" & Rows.Count).End(xlUp).Row
    Range("T1:Y" & i).Select
    Selection.AutoFilter

End Sub

Sub 筛选合同汇总表() '筛选方案
    Dim StartTime
    StartTime = Timer
    
    'If ActiveSheet.AutoFilterMode = True Then ActiveSheet.AutoFilterMode = False
    i = ActiveSheet.Range("AW" & Rows.Count).End(xlUp).Row
    Range("AV1:BN" & i).Select
    Selection.AutoFilter
    
    'MsgBox Timer - StartTime
End Sub

Sub 筛选发票() '筛选方案
    'If ActiveSheet.AutoFilterMode = True Then ActiveSheet.AutoFilterMode = False
    i = ActiveSheet.Range("AM" & Rows.Count).End(xlUp).Row
    Range("AL1:AR" & i).Select
    Selection.AutoFilter

End Sub

Sub 筛选付款() '筛选方案
    'If ActiveSheet.AutoFilterMode = True Then ActiveSheet.AutoFilterMode = False
    i = ActiveSheet.Range("AM" & Rows.Count).End(xlUp).Row
    Range("Z1:AJ" & i).Select
    Selection.AutoFilter

End Sub

Sub 更改透视表数据源()
    Dim oPT  As PivotTable
    Dim oPC As PivotCache
    Dim oWK As Worksheet
    Set oWK = Worksheets("透视表")
    iRow = Worksheets("合").Range("B65536").End(xlUp).Row
    With oWK
        Set oPT = .PivotTables(1)
        With oPT
            '获取原来的数据透视表的数据源
           sOrign = .SourceData
           '直接将数据源更改为其它单元格区域
          .SourceData = Worksheets("合").Range("A1:P" & iRow).Address(True, True, xlR1C1, True)
          '获取最新的数据透视表的数据源
          sNew = .SourceData
          '刷新透视表
          .RefreshTable
          '刷新数据源
          .Update
        End With
    End With
End Sub
