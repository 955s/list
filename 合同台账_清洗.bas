Attribute VB_Name = "合同台账_清洗"
Sub 替换合同台账包件名称()
    Worksheets("上报告").Activate
    x = Worksheets("上报告").Range("V" & Rows.Count).End(xlUp).Row
    With Worksheets("上报告").Range("V2:V" & x)
        .Replace What:="*（", Replacement:="", SearchOrder:=xlByColumns
        .Replace What:="）", Replacement:="", SearchOrder:=xlByColumns
        .Replace What:=" ", Replacement:="", SearchOrder:=xlByColumns
    End With
End Sub

Sub 合同台账格式()
    'Application.ScreenUpdating = False
    Worksheets("上报告").Activate
    Dim n, 序号
    dq = Worksheets("上报告").Name
    x = Worksheets("上报告").Range("V" & Rows.Count).End(xlUp).Row
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




