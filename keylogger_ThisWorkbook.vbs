Option Explicit

'PERSONAL.xls
'[not module] ThisWorkbook
'
'[manual change (output log path) Line:68 or 74]
    '  datFile = "C:\<somePath>\ExcelInputLog.txt"
    '  datFile = ActiveWorkbook.Path & "\ExcelInputLog.txt"

'<how to use>
'step1 line 20~105 copy to paste PERSONAL.xls ThisWorkbook
'step2 line 112 ~ 126 copy to paste PERSONAL.xls some module
'
'if you need monitoring
'PowerShell
'  Get-Content  C:\<somePath>\ExcelInputLog.txt  -wait  -tail  0
'================================================================

Dim WithEvents xlApp As Application
Private Sub Workbook_Open()
   Set xlApp = Application
End Sub
Private Sub xlApp_WorkbookAfterSave(ByVal Wb As Workbook, ByVal Success As Boolean)
    Application.StatusBar = "PERSONALxls Reloading..."
    Call ab001_reloadPERSONALxls
End Sub

Private Sub xlApp_WorkbookOpen(ByVal Wb As Workbook)
    On Error Resume Next
    Application.StatusBar = False
    If Not Wb Is ThisWorkbook Then
        Debug.Print "<<" & Wb.Name & ">> is Open"
    End If
End Sub
Private Sub xlApp_SheetChange(ByVal sh As Object, ByVal Target As Range)

    Debug.Print "--------------------------------"
    Debug.Print "TimeStamp:" & Format(Now, "yyyy/mm/dd HH:MM:SS")
    Debug.Print "Book:" & ActiveWorkbook.Name
    Debug.Print "Sheet:" & sh.Name
    Debug.Print "Cell:" & Target.Address(RowAbsolute:=False, ColumnAbsolute:=False)
    
    Dim Debug_c As Variant
    Dim Debug_loopBreakCount As Integer
    For Each Debug_c In Range(Target.Address)
        Debug_loopBreakCount = Debug_loopBreakCount + 1
        If Debug_loopBreakCount > 2 Then
            Exit Sub
            'MsgBox ("If Excel Frieze:Ctrl + Pause/Break")
        End If
        If Debug_c.Value <> "" Then
            On Error Resume Next
            Debug.Print "Value:" & Range(Target.Address).Value
        Else
            Debug.Print "Value:Empty"
        End If
    Next
    
    'ExcelInputLog
    On Error Resume Next
    
    Dim datFile As String
    
    'Path Log.txt====================================================================================
    'datFile���w�肷��ꍇ�i�R�����g�A�E�g��������ꍇ�j�A�S�Ă�Excel�����text���O�t�@�C������ɂ܂Ƃ܂�܂��B
    'If use Log.txt only 1 file log.
    '  datFile = "C:\<somePath>\ExcelInputLog.txt"
    
    'datFile���w�肵�Ȃ��ꍇ�i�R�����g�A�E�g�̂܂܂̏ꍇ�j�A
    'text���O�t�@�C���͊eExcel�t�@�C��������t�H���_�̉��֑��삷�邽�тɍ쐬����܂��B
    '�i���Excel�ɂ���̃��O�t�@�C���j�B
    'Else use Log.txt "Any" Excel changes.
    '  datFile = ActiveWorkbook.Path & "\ExcelInputLog.txt"
    
    'Path Log.txt====================================================================================
    
    Open datFile For Append As #1
    
    Print #1, "--------------------------------"
    Print #1, "TimeStamp:" & Format(Now, "yyyy/mm/dd HH:MM:SS")
    Print #1, "Book:" & ActiveWorkbook.Name
    Print #1, "Sheet:" & sh.Name
    Print #1, "Cell:" & Target.Address(RowAbsolute:=False, ColumnAbsolute:=False)
    
    Dim c As Variant
    Dim loopBreakCount As Integer
    For Each c In Range(Target.Address)
        loopBreakCount = loopBreakCount + 1
        If loopBreakCount > 2 Then
            Exit Sub
            'MsgBox ("If Excel Frieze:Ctrl + Pause/Break")
        End If
        If c.Value <> "" Then
            On Error Resume Next
            Print #1, "Value:" & Range(Target.Address).Value
        Else
            Print #1, "Value:Empty"
        End If
    Next
    
    Close #1
    
End Sub