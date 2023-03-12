Attribute VB_Name = "SQL加载付款台账"
Sub 获取付款台账()
    'Kill "E:\Download\物资付款单*.xls" '删除文件
    Application.ScreenUpdating = False   '禁刷新
    Dim cnn As Object, rs As Object
    Dim sql As String
    Set cnn = CreateObject("Adodb.Connection")
    cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties='Excel 12.0;HDR=NO;IMEX=1';Data Source=" & "E:\Download\物资付款单.xlsx"
    Set rs = CreateObject("Adodb.RecordSet")
    sql = "select * from [物资付款单$A:R] Where [F3] like '%长沙西%' and [F18] like '%审批%'"
    Set rs = cnn.Execute(sql)
    arr1 = rs.GetRows '放入数组内
    arr1 = Application.WorksheetFunction.Transpose(arr1)    '改造输出结构
    行数 = UBound(arr1, 1)
    dq = Worksheets("上报告").Range("Z" & Rows.Count).End(xlUp).Row
    Worksheets("上报告").Range("Z2:AI" & dq).ClearContents '清空原始数据
    '单据编号    单据日期    合同名称    录入合同编号 客商名称 合同金额 开累应付金额 开累已付金额 应付未付金额 本次付款金额
    For i = 1 To UBound(arr1, 1)
        With Worksheets("上报告")
            .Range("Z" & i + 1) = arr1(i, 1)
            .Range("AA" & i + 1) = arr1(i, 2)
            .Range("AB" & i + 1) = arr1(i, 8)
            .Range("AC" & i + 1) = arr1(i, 9)
            .Range("AD" & i + 1) = arr1(i, 10)
            .Range("AE" & i + 1) = arr1(i, 11)
            .Range("AF" & i + 1) = arr1(i, 13)
            .Range("AG" & i + 1) = arr1(i, 14)
            .Range("AH" & i + 1) = arr1(i, 15)
            .Range("AI" & i + 1) = arr1(i, 16)
            .Range("AJ" & i + 1) = arr1(i, 18)
        End With
    Next
    'Sheets("out").Range("A2").CopyFromRecordset rs'整体查询输出
    rs.Close '关闭记录集
    cnn.Close '关闭与数据库的链接
    Set rs = Nothing '释放对象
    Set cnn = Nothing '释放对象
    Call 替换付款台账合同名称
    Call 替换付款台账客商名称
    Call 付款台账格式
    
End Sub

Sub 替换付款台账合同名称()
    Worksheets("上报告").Activate
    x = Worksheets("上报告").Range("AB" & Rows.Count).End(xlUp).Row
    With Worksheets("上报告").Range("AB2:AB" & x)
        .Select
        .Replace What:="长沙西站项目自购物资采购合同（", Replacement:="", SearchOrder:=xlByColumns
        .Replace What:="）", Replacement:="", SearchOrder:=xlByColumns
        .Replace What:=" ", Replacement:="", SearchOrder:=xlByColumns
    End With
End Sub

Sub 替换付款台账客商名称()
    Worksheets("上报告").Activate
    x = Worksheets("上报告").Range("AD" & Rows.Count).End(xlUp).Row
    With Worksheets("上报告").Range("AD2:AD" & x)
        .Select
        .Replace What:="本级", Replacement:="", SearchOrder:=xlByColumns
    End With
End Sub

Sub 付款台账格式()
    Application.ScreenUpdating = False
    With Worksheets("上报告")
        If Worksheets("上报告").AutoFilterMode = True Then Selection.AutoFilter '如果有筛选就先取消筛选
        Dim n, 序号
        dq = Worksheets("上报告").Name
        x = Worksheets("上报告").Range("Z" & Rows.Count).End(xlUp).Row
        .Range("Z1:AJ" & x).Font.Size = 10 '指定区域字号
        .Range("Z1:AJ" & x).HorizontalAlignment = xlCenter '居中
        .Range("AE2:AI" & x).Select
        Selection.NumberFormatLocal = "0.00_ ;[红色]-0.00 "
        .Range("Z2").Select
    End With
End Sub

Sub 判断付款台账是否存在()
    Dim MyFile As Object
    Dim Str As String
    Dim StrMsg As String
    Str = "E:\Download\物资付款单.xls"
    Set MyFile = CreateObject("Scripting.FileSystemObject")
    If MyFile.FileExists(Str) Then
        MsgBox "付款台账存在！"
    Else
        MsgBox "付款台账不存在！"
    End If
End Sub

