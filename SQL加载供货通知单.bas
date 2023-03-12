Attribute VB_Name = "SQL加载供货通知单"
Sub 获取接触网供货通知单汇总表()
    'Kill "E:\Download\物资付款单*.xls" '删除文件
    Application.ScreenUpdating = False   '禁刷新
    Dim cnn As Object, rs As Object
    Dim sql As String
    Set cnn = CreateObject("Adodb.Connection")
    cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties='Excel 12.0;HDR=NO;IMEX=1';Data Source=" & "E:\cooo\OneDrive - Collect\gs\Chaos\Take Over\长沙西站\供货通知单\长沙西站供货通知单-接触网.xlsm"
    Set rs = CreateObject("Adodb.RecordSet")
    sql = "select * from [汇总表$A:Q] where F14 <> '电话'"
    Set rs = cnn.Execute(sql)
    x = Worksheets("供货单").Range("A" & Rows.Count).End(xlUp).Row
    Sheets("供货单").Range("A2:" & "Q" & x).ClearContents

    Sheets("供货单").Range("A2").CopyFromRecordset rs '整体查询输出
    rs.Close '关闭记录集
    cnn.Close '关闭与数据库的链接
    Set rs = Nothing '释放对象
    Set cnn = Nothing '释放对象
    'xdr = Format(Sheets("供货单").Range("O2:O1000"), "Long Date")
    Sheets("供货单").Range("O:O").TextToColumns other:=True, otherchar:=""
    Sheets("供货单").Range("P:P").TextToColumns other:=True, otherchar:=""
    
End Sub

Sub 获取信号供货通知单汇总表()
    'Kill "E:\Download\物资付款单*.xls" '删除文件
    Application.ScreenUpdating = False   '禁刷新
    Dim cnn As Object, rs As Object
    Dim sql As String
    Set cnn = CreateObject("Adodb.Connection")
    cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties='Excel 12.0;HDR=NO;IMEX=1';Data Source=" & "E:\cooo\OneDrive - Collect\gs\Chaos\Take Over\长沙西站\供货通知单\长沙西站供货通知单-信号.xlsm"
    Set rs = CreateObject("Adodb.RecordSet")
    sql = "select * from [汇总表$A:Q] where F14 <> '电话'"
    Set rs = cnn.Execute(sql)
    x = Worksheets("供货单").Range("B" & Rows.Count).End(xlUp).Row
    'Sheets("供货单").UsedRange.Offset(1).ClearContents
    Sheets("供货单").Range("A" & x).CopyFromRecordset rs '整体查询输出
    rs.Close '关闭记录集
    cnn.Close '关闭与数据库的链接
    Set rs = Nothing '释放对象
    Set cnn = Nothing '释放对象
End Sub

Sub 获取通信供货通知单汇总表()
    'Kill "E:\Download\物资付款单*.xls" '删除文件
    Application.ScreenUpdating = False   '禁刷新
    Dim cnn As Object, rs As Object
    Dim sql As String
    Set cnn = CreateObject("Adodb.Connection")
    cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties='Excel 12.0;HDR=NO;IMEX=1';Data Source=" & "E:\cooo\OneDrive - Collect\gs\Chaos\Take Over\长沙西站\供货通知单\长沙西站供货通知单-通信.xlsm"
    Set rs = CreateObject("Adodb.RecordSet")
    sql = "select * from [汇总表$A:Q] where F14 <> '电话'"
    Set rs = cnn.Execute(sql)
    x = Worksheets("供货单").Range("B" & Rows.Count).End(xlUp).Row
    'Sheets("供货单").UsedRange.Offset(1).ClearContents
    Sheets("供货单").Range("A" & x).CopyFromRecordset rs '整体查询输出
    rs.Close '关闭记录集
    cnn.Close '关闭与数据库的链接
    Set rs = Nothing '释放对象
    Set cnn = Nothing '释放对象
End Sub

Sub 获取电力供货通知单汇总表()
    'Kill "E:\Download\物资付款单*.xls" '删除文件
    Application.ScreenUpdating = False   '禁刷新
    Dim cnn As Object, rs As Object
    Dim sql As String
    Set cnn = CreateObject("Adodb.Connection")
    cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties='Excel 12.0;HDR=NO;IMEX=1';Data Source=" & "E:\cooo\OneDrive - Collect\gs\Chaos\Take Over\长沙西站\供货通知单\长沙西站供货通知单-电力.xlsm"
    Set rs = CreateObject("Adodb.RecordSet")
    sql = "select * from [汇总表$A:Q] where F14 <> '电话'"
    Set rs = cnn.Execute(sql)
    x = Worksheets("供货单").Range("B" & Rows.Count).End(xlUp).Row
    'Sheets("供货单").UsedRange.Offset(1).ClearContents
    Sheets("供货单").Range("A" & x).CopyFromRecordset rs '整体查询输出
    rs.Close '关闭记录集
    cnn.Close '关闭与数据库的链接
    Set rs = Nothing '释放对象
    Set cnn = Nothing '释放对象
End Sub

Sub 获取变电供货通知单汇总表()
    'Kill "E:\Download\物资付款单*.xls" '删除文件
    Application.ScreenUpdating = False   '禁刷新
    Dim cnn As Object, rs As Object
    Dim sql As String
    Set cnn = CreateObject("Adodb.Connection")
    cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties='Excel 12.0;HDR=NO;IMEX=1';Data Source=" & "E:\cooo\OneDrive - Collect\gs\Chaos\Take Over\长沙西站\供货通知单\长沙西站供货通知单-变电.xlsm"
    Set rs = CreateObject("Adodb.RecordSet")
    sql = "select * from [汇总表$A:Q] where F14 <> '电话'"
    Set rs = cnn.Execute(sql)
    x = Worksheets("供货单").Range("B" & Rows.Count).End(xlUp).Row
    'Sheets("供货单").UsedRange.Offset(1).ClearContents
    Sheets("供货单").Range("A" & x).CopyFromRecordset rs '整体查询输出
    rs.Close '关闭记录集
    cnn.Close '关闭与数据库的链接
    Set rs = Nothing '释放对象
    Set cnn = Nothing '释放对象
End Sub

Sub 供货通知单汇总表()
    'Kill "E:\Download\物资付款单*.xls" '删除文件
    Call 取消隐藏隐藏供货单列
    Call 获取接触网供货通知单汇总表
    Call 获取信号供货通知单汇总表
    Call 获取通信供货通知单汇总表
    Call 获取电力供货通知单汇总表
    Call 获取变电供货通知单汇总表
    Call 供货通知单格式
End Sub

Sub 供货通知单格式()
    Sheets("供货单").Range("G:G").TextToColumns other:=True, otherchar:=""
    Sheets("供货单").Range("I:I").TextToColumns other:=True, otherchar:=""
    Sheets("供货单").Range("N:N").TextToColumns other:=True, otherchar:=""
    Sheets("供货单").Range("O:O").TextToColumns other:=True, otherchar:=""
    Sheets("供货单").Range("P:P").TextToColumns other:=True, otherchar:=""
    x = Sheets("供货单").[B65535].End(xlUp).Row
    'Range("A1:P" & x).Borders.LineStyle = 1 '指定区域框线
    Sheets("供货单").Rows("2:" & x).RowHeight = 20 '指定区域行高
    Sheets("供货单").Range("A1:Q" & x).Font.Size = 10 '指定区域字号
    Sheets("供货单").Range("A1:Q" & x).HorizontalAlignment = xlCenter '居中
    
    'xdr = Format(Sheets("供货单").Range("O2:O1000"), "Long Date")
    
End Sub

