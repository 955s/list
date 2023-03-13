Attribute VB_Name = "合同台账_SQL加载"
Sub 获取合同台账()
    'Kill "E:\Download\物资付款单*.xls" '删除文件
    Application.ScreenUpdating = False   '禁刷新
    Dim cnn As Object, rs As Object
    Dim sql As String
    Dim arr1
    Dim dq,i
    Set cnn = CreateObject("Adodb.Connection")
    cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties='Excel 12.0;HDR=NO;IMEX=1';Data Source=" & "E:\Download\合同台账.xls"
    Set rs = CreateObject("Adodb.RecordSet")
    sql = "select * from [合同台账$A:K] Where [F1] like '%长沙西%'" _
    & " and [F2] <> 'CRDH0004-CSXZ-LWPQ-2022-0017'" _
    & " and [F2] <> 'CRDH0004-CSXZ-WZtxsb-2022-0025'" _
    & " and [F2] <> 'CRDH0004-CSXZ--ZL-2022-0017'" _
    & " and [F2] <> 'CRDH0004-CSXZ--ZL-2022-0018'" _
    & " and [F2] <> 'CRDH0004-CSXZ-ZL-2022-0023'" _
    & " and [F2] <> 'CRDH0004-CSXZ--WZxh-2022-0018（1）'" _
    & " and [F2] <> 'CRDH0004-CSXZ-WZxh-2022-0017'" _
    & " and [F2] <> 'CRDH0004-CSXZ-WZtxsb-2022-0021'" _
    & " and [F2] <> 'CRDH0004-CSXZ-WZxh-2022-0002'" _
    & " and [F2] <> 'CRDH0004-CSXZ-ZL-2022-0001'" _
    & " and [F2] <> 'CRDH0004-CSXZ-ZL-2022-0003'" _
    & " and [F2] <> 'CRDH0004-CSXZ-ZL-2022-0008'" _
    & " and [F2] <> 'CRDH0004-CSXZ-ZL-2022-0020'" _
    & " and [F2] <> 'CRDH0004-CSXZ-ZL-2022-0021'" _
    & " and [F2] <> 'CRDH0004-CSXZ-ZL-2022-0022'" _
    & " and [F2] <> 'CRDH0004-CSXZ-ZL-2023-0001'" _
    & " and [F2] <> 'CRDH0004-CSXZ-ZL-2023-0002'" _
    & " and [F2] <> 'CRDH0004-CSXZ-WZqybd-2022-0032'"
    
    Set rs = cnn.Execute(sql)
    arr1 = rs.GetRows '放入数组内
    arr1 = Application.WorksheetFunction.Transpose(arr1)    '改造输出结构
    With Worksheets("上报告")
    行数 = UBound(arr1, 1)
    dq = Worksheets("上报告").Range("V" & Rows.Count).End(xlUp).Row
    If dq < 2 Then dq = dq + 1 '防止清表头
    .Range("T2:X" & dq).ClearContents '清空原始数据
    .Range("T2:X" & dq).Interior.Pattern = xlNone '清理填充
    .Range("T2:X" & dq).ClearComments '清理批注
    .Range("U2:X200").Interior.Pattern = xlNone '清理填充
    .Range("U2:X200").ClearComments '清理批注
    '包件名称    合同号  金额
    For i = 1 To UBound(arr1, 1)

        .Range("V" & i + 1) = arr1(i, 1)
        .Range("W" & i + 1) = arr1(i, 2)
        .Range("X" & i + 1) = arr1(i, 7)

    Next
    'Sheets("out").Range("A2").CopyFromRecordset rs'整体查询输出
    rs.Close '关闭记录集
    cnn.Close '关闭与数据库的链接
    Set rs = Nothing '释放对象
    Set cnn = Nothing '释放对象
    End With
End Sub

Sub 替换合同台账包件名称()
    Worksheets("上报告").Activate
    x = Worksheets("上报告").Range("V" & Rows.Count).End(xlUp).Row
    With Worksheets("上报告").Range("V2:V" & x)
        .Replace What:="*（", Replacement:="", SearchOrder:=xlByColumns
        .Replace What:="）", Replacement:="", SearchOrder:=xlByColumns
        .Replace What:=" ", Replacement:="", SearchOrder:=xlByColumns
    End With
End Sub

Sub 修正合同台账()
    'Application.ScreenUpdating = False
    Dim StartTime,x
    StartTime = Timer
    Worksheets("上报告").Activate
    x = Worksheets("上报告").Range("V" & Rows.Count).End(xlUp).Row
    If x < 2 Then x = x + 1 '防止清表头
    Call 替换合同台账包件名称
    With Worksheets("上报告")
        If Worksheets("上报告").AutoFilterMode = True Then Selection.AutoFilter '如果有筛选就先取消筛选

        .Range("T1:Y" & x).Font.Size = 10 '指定区域字号
        .Range("T1:Y" & x).HorizontalAlignment = xlCenter '居中
        .Range("X2:Y" & x).NumberFormatLocal = "0!.0,!0" '设置万元单位
        .Range("T2").Formula = "=VLOOKUP(U2,透视表!A:E,4,0)-X2" '核对
        .Range("U2").Formula = "=VLOOKUP(V2,IF({1,0},K:K,J:J),2,FALSE)" '包件号
                
        [T2:U2].AutoFill Destination:=Range("T2:U" & x) '公式填充
        [Y2:Y2].AutoFill Destination:=Range("Y2:Y" & x) '公式填充
    End With

    Call 修正合同台账包件

    MsgBox Timer - StartTime
End Sub

Sub 判断合同台账是否存在()
    Dim MyFile As Object
    Dim Str As String
    Dim StrMsg As String
    Str = "E:\Download\合同台账.xls"
    Set MyFile = CreateObject("Scripting.FileSystemObject")
    If MyFile.FileExists(Str) Then
        MsgBox "合同台账存在！"
    Else
        MsgBox "合同台账不存在！"
    End If
End Sub

Sub 合同台账排序()
    Dim rng As Range
    Worksheets("上报告").Activate
    With Worksheets("上报告")
    x = .Range("V" & Rows.Count).End(xlUp).Row
    Set rng = .Range("U1:X" & x)
    rng.Sort Key1:="包件号", Order1:=xlAscending, Header:=xlYes
    
    End With
End Sub

