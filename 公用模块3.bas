Attribute VB_Name = "公用模块3"
Sub 数据源()
'
' 宏1 宏
'

'
    ActiveSheet.PivotTables("数据透视表1").ChangePivotCache ActiveWorkbook.PivotCaches. _
        Create(SourceType:=xlDatabase, SourceData:= _
        "https://cooo-my.sharepoint.com/personal/o_cooo_onmicrosoft_com/Documents/gs/Chaos/Take Over/长沙西站/[清单.xlsm]合!R1C2:R2000C16" _
        , Version:=7)
End Sub
