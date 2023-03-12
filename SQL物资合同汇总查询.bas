Attribute VB_Name = "SQL物资合同汇总查询"
Sub 获取物资合同汇总查询()
    'Kill "E:\Download\物资付款单*.xls" '删除文件
    Application.ScreenUpdating = False   '禁刷新
    Dim cnn As Object, rs As Object
    Dim sql As String
    Set cnn = CreateObject("Adodb.Connection")
    cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties='Excel 12.0;HDR=NO;IMEX=1';Data Source=" & "E:\Download\物资设备采购合同汇总查询.xlsx"
    Set rs = CreateObject("Adodb.RecordSet")
    sql = "select * from [sheet1$A5:BJ1000]"
    Set rs = cnn.Execute(sql)
    Worksheets("上报告").Range("AW2:DF1000").ClearContents '清空原始数据
    Sheets("上报告").Range("AW2").CopyFromRecordset rs '整体查询输出
    dq = Worksheets("上报告").Range("AW" & Rows.Count).End(xlUp).Row
    arr1 = Sheets("上报告").Range("AW2:DF" & dq)
    Worksheets("上报告").Range("AW2:DF" & dq).ClearContents '清空原始数据
    
    'arr1 = Application.WorksheetFunction.Transpose(arr1)    '改造输出结构
    行数 = UBound(arr1, 1)
    
    For i = 1 To UBound(arr1, 1)
        With Worksheets("上报告")
            .Range("AW" & i + 1) = arr1(i, 5) '合同名称
            .Range("AX" & i + 1) = arr1(i, 6) '签定时间
            .Range("AY" & i + 1) = arr1(i, 7) '合同编号
            .Range("AZ" & i + 1) = arr1(i, 9) '客商名称
            .Range("BA" & i + 1) = arr1(i, 15) '税率
            .Range("BB" & i + 1) = arr1(i, 17) '调整后合同额
            .Range("BC" & i + 1) = arr1(i, 21) '结算后可支付比例
            .Range("BD" & i + 1) = arr1(i, 22) '开通后可支付比例
            .Range("BE" & i + 1) = arr1(i, 28) '质保金
            .Range("BF" & i + 1) = arr1(i, 33) '结算含税
            .Range("BG" & i + 1) = arr1(i, 42) '已开票金额（）
            .Range("BH" & i + 1) = arr1(i, 46) '应付购货款
            .Range("BI" & i + 1) = arr1(i, 49) '已付购货款
            .Range("BJ" & i + 1) = arr1(i, 55) '可支付金额
        End With
    Next
    rs.Close '关闭记录集
    cnn.Close '关闭与数据库的链接
    Set rs = Nothing '释放对象
    Set cnn = Nothing '释放对象
    
    'Call 替换物资合同汇总名称
    'Call 付款台账格式
    
End Sub

Sub 替换物资合同汇总名称()
    
    n = Worksheets("上报告").Range("AW" & Rows.Count).End(xlUp).Row
    With Worksheets("上报告").Range("AW2:AW" & n)
        .Activate
        .Range("AW2:AW" & n).Select
        .Replace What:="长沙西站项目自购物资采购合同（", Replacement:="", SearchOrder:=xlByColumns
        .Replace What:="）", Replacement:="", SearchOrder:=xlByColumns
                
        .Range("BK2").Value = "=(BG2*BC2)-BI2"  '可付公式
        .Range("BL2").Value = "=SUMIF(AL:AL,AV2,AR:AR)-BG2" '票差公式
        .Range("BM2").Value = "=VLOOKUP(AV2,透视表!$A$2:$E$200,5,0)-BF2" '点差公式
        [BK2:BM2].AutoFill Destination:=Range("BK2:BM" & n) '公式填充
        .Range("BK" & n + 1).Formula2 = "=SUM(IF(ISNUMBER(BK2:BK" & n & "),BK2:BK" & n & ",0))"
        .Range("BL" & n + 1).Formula2 = "=SUM(IF(ISNUMBER(BL2:BL" & n & "),BL2:BL" & n & ",0))"
        .Range("BM" & n + 1).Formula2 = "=SUM(IF(ISNUMBER(BM2:BM" & n & "),BM2:BM" & n & ",0))"
        .Range("BF" & n + 1 & ":" & "BN" & n + 1).Font.Bold = True '加粗
        '.Range("BF" & n + 1 & ":" & "BN" & n + 1).ClearContents '清理
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

Sub 判断物资合同台账是否存在()
    Dim MyFile As Object
    Dim Str As String
    Dim StrMsg As String
    Str = "E:\Download\物资设备采购合同汇总查询.xlsx"
    Set MyFile = CreateObject("Scripting.FileSystemObject")
    If MyFile.FileExists(Str) Then
        MsgBox "付款台账存在！"
    Else
        MsgBox "付款台账不存在！"
    End If
End Sub



