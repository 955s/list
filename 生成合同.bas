Attribute VB_Name = "生成合同"
Sub 取消工作表保护()

    ActiveSheet.Unprotect
    
End Sub

Sub 取值()
    Application.ScreenUpdating = False   '禁刷新
    Dim arr1, arr, cx, i, c
    Dim 批次, 厂家, 表行
    批次 = Sheets("M").Range("P1").Value
    厂家 = Sheets("M").Range("O1").Value
    表行 = Sheets(批次).Range("D" & Rows.Count).End(xlUp).Row '查询当前表占用的行
    Worksheets(批次).Activate
    arr1 = Range("D3:D" & 表行)
    '循环查找第一次出现的位置
    For i = 1 To UBound(arr1)
        If InStrRev(arr1(i, 1), 厂家) Then
            n = i + 2
        Exit For
            End If
    Next
    'Sheets("第三批").Range("D" & n).Select
    '循环查找最后一次出现的位置
    For i = 1 To UBound(arr1)
        If InStrRev(arr1(i, 1), 厂家) Then
            k = i + 2
        End If
    Next
    
    arr = Sheets(批次).Range("E" & n & ":H" & k) '指定区域
    arr_dj = Sheets(批次).Range("j" & n & ":j" & k) '指定区域
    Worksheets("M").Activate
    Sheets("M").Unprotect
    Sheets("M").Unprotect
    mbx = Sheets("M").Range("B" & Rows.Count).End(xlUp).Row
    If mbx > 3 Then
        Rows(3 & ":" & mbx).Delete
        Else
    End If
    
    Sheets("M").Range("B3").Resize(UBound(arr, 1), UBound(arr, 2)) = arr '将汇总结果输出
    Sheets("M").Range("J3").Resize(UBound(arr_dj, 1), UBound(arr_dj, 2)) = arr_dj '将汇总结果输出
    Call 格式
End Sub

Sub 格式()
    mbx = Sheets("M").Range("B" & Rows.Count).End(xlUp).Row '查询M表占用的行
    Range("F3:F" & mbx).Formula = "=J3/1.13" '到站单价
    Range("G3:G" & mbx).Formula = "=E3*F3" '到站合价
    Range("H3:H" & mbx).Value = "13%" '税率
    Range("I3:I" & mbx).Formula = "=K3-G3" '税额
    Range("K3:K" & mbx).Formula = "=E3*J3" '合价（含税）
    Range("M3:M" & mbx).Value = "长沙市" '税率
    mbx_1 = Sheets("M").Range("B" & Rows.Count).End(xlUp).Row + 1 '查询M表占用的行
    Range("G" & mbx_1).Formula = "=Sum(G" & mbx_1 - 1 & ":G3)" '到站价合价（不含税）
    Range("I" & mbx_1).Formula = "=Sum(I" & mbx_1 - 1 & ":I3)" '税额合计
    Range("K" & mbx_1).Formula = "=Sum(K" & mbx_1 - 1 & ":K3)" '合价（含税）
    Range("F3:G" & mbx_1).NumberFormatLocal = "0.00" '指定区域保留二位小数
    Range("I3:K" & mbx_1).NumberFormatLocal = "0.00" '指定区域保留二位小数
    Range("B" & mbx_1).Value = "合计：" '写入合计：
    '填充序号
    Range("A3").Value = "1"
    Range("A3").Select
    Selection.AutoFill Destination:=Range("A3:A" & mbx), Type:=xlFillSeries
    Range("A2:M" & mbx + 1).Borders.LineStyle = 1 '画框线
    mbx = Sheets("M").Range("B" & Rows.Count).End(xlUp).Row '查询M表占用的行
    Range("C" & mbx_1).Formula = "=K" & mbx '到站合价
    Range("C" & mbx_1).Select
    Selection.NumberFormatLocal = "[DBNum2][$-zh-CN]G/通用格式"
    
    [M3].Select
    '核对公式
    '=SUMIF(第三批!D:D,O1,第三批!J:J)-SUMIF(B:B,"合计：",K:K)
End Sub

Sub 清空()

    Application.ScreenUpdating = False   '禁刷新
    Worksheets("M").Activate
    Sheets("M").Unprotect
    Sheets("M").Unprotect
    mbx = Sheets("M").Range("B" & Rows.Count).End(xlUp).Row '查询M表占用的行
    
    If Sheets("M").Range("B" & Rows.Count).End(xlUp).Row > 3 Then
        Rows(3 & ":" & mbx).Delete
        Else
    End If
        
    [M3].Select
End Sub

Sub 保护所有工作表()
        Application.ScreenUpdating = False   '禁刷新
        With ThisWorkbook
            For i = 1 To .Worksheets.Count    '遍历文件的工作表数
                表名 = .Worksheets(i).Name
                If 表名 <> "out" Then '判断不为out的表格 即所以表
                    .Sheets(表名).Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
                End If
            Next i
        End With
End Sub

Sub qx()
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


