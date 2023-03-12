Attribute VB_Name = "透视表"
Sub 合并各批次()
    Application.ScreenUpdating = False   '禁刷新
    'StartTime = Timer
    Dim n&, i&, j&, gs
    Dim 输出()
    Worksheets("合").Activate
    n = 1
    If Sheets("合").AutoFilterMode = True Then Selection.AutoFilter '如果有筛选就先取消筛选
    With ThisWorkbook  '使用 GetObject 函数可以访问文件
        For i = 1 To .Worksheets.Count    '遍历文件的工作表数
            表名 = .Worksheets(i).Name
            If InStr(表名, "批") Then
                ddq = Sheets(表名).Range("E" & Rows.Count).End(xlUp).Row
                arr = .Sheets(表名).Range("B3:P" & ddq)   '整个sheet载入内存
                For k = 1 To UBound(arr)    '轮询整个表
                    ReDim Preserve 输出(1 To UBound(arr, 2), 1 To n)    '指定一个输出表
                    输出(1, n) = n   '写入每一行的基础数据
                    For l = 1 To UBound(arr, 2)  '循环写入标准数据
                        输出(l, n) = arr(k, l)
                    Next l
                    n = n + 1   '计数器增加
                Next k
            End If
            'n = n + 1'空行需求
        Next i
    End With

    输出 = Application.WorksheetFunction.Transpose(输出)    '改造输出结构
    
    If Sheets("合").Range("B" & Rows.Count).End(xlUp).Row > 1 Then
        Sheets("合").Rows(2 & ":" & Range("B" & Rows.Count).End(xlUp).Row).Delete
    Else
    End If
    Sheets("合").Unprotect
    Sheets("合").Range("B2").Resize(UBound(输出, 1), UBound(输出, 2)) = 输出 '将汇总结果输出
    Call 合表格式
    'MsgBox Timer - StartTime
End Sub

Sub 合表格式()
    'Application.ScreenUpdating = False
    'On Error Resume Next
    If ActiveSheet.AutoFilterMode = True Then Selection.AutoFilter '如果有筛选就先取消筛选
    Dim n, 序号, dq, x, i
    dq = Sheets("合").Name
    x = Sheets("合").[B65535].End(xlUp).Row
    arr = Sheets("合").Range("B2:B" & x)
    ReDim 序号(1 To UBound(arr), 1 To 1)
    For i = 1 To UBound(arr)
        If arr(i, 1) <> "" Then
            序号(i, 1) = n + 1: n = n + 1
        End If
    Next i
    n = 0
    
    Sheets("合").Activate
    Sheets("合").Unprotect
    Range("A1").Select
    Range("A2").Resize(UBound(序号), 1) = 序号
    Range("A1:P" & x).Borders.LineStyle = 1 '指定区域框线
    Rows("2:" & x).RowHeight = 20 '指定区域行高
    Range("A1:P" & x).Font.Size = 10 '指定区域字号
    Range("A1:P" & x).HorizontalAlignment = xlCenter '居中
End Sub

Sub 修正透视表()

    Application.ScreenUpdating = False   '禁刷新
    Dim arr1, arr, cx, i, c
    Dim 批次, 厂家, 表行
    表行 = Sheets("透视表").Range("D" & Rows.Count).End(xlUp).Row '查询当前表占用的行
    Worksheets("透视表").Activate
    Range("E2:K200").Interior.Pattern = xlNone '清理填充
    Range("E2:K200").ClearComments '清理批注
    arr1 = Range("A2:A" & 表行)
    Call 纯白填充
    '循环查找第一次出现的位置，修正BD-13
    For i = 1 To UBound(arr1)
        If InStrRev(arr1(i, 1), "BD-13") Then
            n = i + 1
        Exit For
            End If
        Next
    Range("G" & n).ClearContents
    Range("G" & n).Value = "=(912176.65+113519.50)*80%"
    Range("G" & n).Interior.Color = 15773696
    Range("G" & n).AddCommentThreaded ("与DL-13/BD-13签一个合同，手动修正状态")

End Sub

Sub 刷新透视表()
    
    With Worksheets("透视表")
        .Activate
        Call 取消所有工作表保护
        Call 合并各批次
        Call 更改透视表数据源
        .Activate
    
        r = .[B65535].End(xlUp).Row
        [G2:N2].AutoFill Destination:=Range("G2:N" & r) '公式填充
        Rows(1).RowHeight = 17 '指定区域行高


        '忽略错误求和
        n = Sheets("透视表").Range("A" & Rows.Count).End(xlUp).Row '查询当前表占用的行
        .Range("G" & n).Formula2 = "=SUM(IF(ISNUMBER(G2:G" & r & "),G2:G" & r & ",0))"
        .Range("H" & n).Formula2 = "=SUM(IF(ISNUMBER(H2:H" & r & "),H2:H" & r & ",0))"
        .Range("I" & n).Formula2 = "=SUM(IF(ISNUMBER(I2:I" & r & "),I2:I" & r & ",0))"
        .Range("M" & n).Formula2 = "=SUM(IF(ISNUMBER(M2:M" & r & "),M2:M" & r & ",0))"
        .Range("N" & n).Formula2 = "=SUM(IF(ISNUMBER(N2:N" & r & "),N2:N" & r & ",0))"
        
        .Range("A1:F1").Interior.Color = RGB(221, 235, 247) '修正表头颜色
        .Range("A2:" & "B" & n).Font.Bold = True
        .Range("D" & n & ":" & "N" & n).Font.Bold = True
        
        .Range("G" & n - 1 & ":" & "N" & n - 1).ClearContents
        Call 去除线

        Call 修正透视表
        .Range("A1").Select
    End With
End Sub


