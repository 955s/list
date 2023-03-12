Attribute VB_Name = "公用模块2"
Option Explicit
Option Compare Text
Public endrow, h, i, r, j, x, ck, cb, dq, arr, t
Public Declare PtrSafe Function MessageBoxTimeout Lib "user32" Alias "MessageBoxTimeoutA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long, ByVal wlange As Long, ByVal dwTimeout As Long) As Long
Public Sub MsgTimeout(Optional msg = "Done", Optional t = 1)
    '重新包装一下,为了让自己在vba中更方便的调用 直接使用MsgTimeout(x)
    '    msg就是要显示的提示文字
    '    t 是自动关闭需要等待的时间 默认3秒
    MessageBoxTimeout 0, msg, "Microsoft Excel", 0, 0, t * 600
End Sub

Sub test() '测试程序 每隔5秒就弹出对话框提示进度
    MsgTimeout "已经完成任务"
End Sub

Sub 查询()
Attribute 查询.VB_ProcData.VB_Invoke_Func = "q\n14"
    UserForm1.Show
End Sub

Sub 批次格式()
    'Application.ScreenUpdating = False
    'On Error Resume Next
    Dim n, 序号, dq, x, i, bm
    bm = ActiveSheet.Name
    x = Sheets(bm).[B65535].End(xlUp).Row

    Rows("2:" & x).RowHeight = 20 '指定区域行高

End Sub

Sub 查厂家()
    查询窗口.Show (0)
End Sub

Sub 测试代码执行耗时示例() '筛选方案
    Dim StartTime
    StartTime = Timer
    Call 筛选发票
    MsgBox Timer - StartTime
End Sub
