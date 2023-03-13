Attribute VB_Name = "old"

Sub 替换合同台账包件名称_old()
    Application.ScreenUpdating = False   '禁刷新
    Worksheets("上报告").Activate
    Columns("V:V").Select
    Selection.Replace What:="*（", Replacement:="", LookAt:= _
    xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        
    Selection.Replace What:="）", Replacement:="", LookAt:= _
    xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    
    Selection.Replace What:=" ", Replacement:="", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2

End Sub