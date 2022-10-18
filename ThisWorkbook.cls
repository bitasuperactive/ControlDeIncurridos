VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()

    weekNumCell = "B7"
    targetRange = "C9:I50"
    timerCell = "B2"
    timerOn = False
    taskCell = "B4"
    task = ""
    Application.StatusBar = "No estas realizando ninguna tarea."
    
    Dim weekNum As Integer
    weekNum = Application.WorksheetFunction.weekNum(Now)
    
    If ("Semana " + CStr(weekNum) + " del año." <> Worksheets(1).Range(weekNumCell).Value) Then
        Worksheets(1).Range(targetRange).Value = ""
        Worksheets(1).Range(weekNumCell).Value = "Semana " + CStr(weekNum) + " del año."
    End If
    
    Application.OnTime Now + TimeValue("00:10:00"), "Reminder"

End Sub

' Controlador del evento generado al cambiar la tarea incurrida.
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)

    If (Target.Count > 1) Then
        Exit Sub
    End If
    
    If (Target = Worksheets(1).Range(taskCell) And timerOn = True) Then
        Call SetCell
        Call ResetTimer
        task = Target.Value
    End If

End Sub

' Controlador del evento generado al cerrar el workbook.
' Registra el tiempo incurrido si se estuviera incurriendo alguna tarea.
Private Sub Workbook_BeforeClose(Cancel As Boolean)

    If (timerOn = True) Then
        Call SetCell
        timerOn = False
        Call ResetTimer
    End If

End Sub
