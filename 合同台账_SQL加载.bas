Attribute VB_Name = "合同台账_SQL加载"
Sub 获取合同台账()
    'Kill "E:\Download\物资付款单*.xls" '删除文件
    Application.ScreenUpdating = False   '禁刷新
    Dim cnn As Object, rs As Object
    Dim sql As String
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
    Worksheets("上报告").Range("T2:X" & dq).ClearContents '清空原始数据
    Worksheets("上报告").Range("T2:V200").Interior.Pattern = xlNone '清理填充
    Worksheets("上报告").Range("T2:V200").ClearComments '清理批注
    
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

    .Range("U2:X200").Interior.Pattern = xlNone '清理填充
    .Range("U2:X200").ClearComments '清理批注
    End With
End Sub

