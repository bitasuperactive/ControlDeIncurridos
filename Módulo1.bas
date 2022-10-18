Attribute VB_Name = "Módulo1"
Public weekNumCell As String
Public targetRange As String
Public timerCell As String
Public timerOn As Boolean
Public taskCell As String
Public task As String

Public Sub SetTimer()

    Application.OnTime Now + TimeValue("00:00:01"), "MoveTimer"

End Sub

Public Sub MoveTimer()
    
    If (timerOn = True) Then
        Worksheets(1).Range(timerCell).Value = Worksheets(1).Range(timerCell).Value + TimeValue("00:00:01")
        Application.StatusBar = task + ": " + Format(Range(timerCell).Value, "hh:mm:ss")
        Call SetTimer
    End If

End Sub

Public Sub ResetTimer()
    
    If (timerOn = False) Then
        Worksheets(1).Range(timerCell).Value = TimeValue("00:00:00")
        Application.StatusBar = "No estas realizando ninguna tarea."
    Else
        Worksheets(1).Range(timerCell).Value = TimeValue("00:00:01")
    End If

End Sub

' Actualiza la celda correspondiente a la tarea y al día de la semana con el tiempo incurrido.
Public Sub SetCell()
    
    Dim row As Integer
    Dim column As Integer
    
    For row = 9 To 50
        If (Cells(row, 2).Value = task Or Cells(row, 2).Value = "") Then
            Exit For
        End If
    Next
    
    column = Weekday(Date, 2) + 2
    
    Cells(row, column).Value = Cells(row, column).Value + Worksheets(1).Range(timerCell).Value * 24
    
End Sub

' Redondea el tiempo incurrido en la última tarea si faltan 10 minutos o menos para completar la jornada laboral.
' * [Método no implemetado]
' Falta por implementar una forma de que el usuario introduzca su jornada.
Public Sub RoundResult()

    Dim row As Integer
    Dim column As Integer
    
    row = 7
    column = Weekday(Date, 2) + 2
    
    If (Cells(row, column).Value < 7.5 And Cells(row, column).Value >= 7.33) Then
        Worksheets(1).Range(timerCell).Value = (7.5 - Cells(row, column).Value) / 24
        Call SetCell
    End If

End Sub

' Lanza un mensaje de recordatorio cada 10 minutos si no se está incurriendo ninguna tarea.
Public Sub Reminder()
    
    If (task = "") Then
        MsgBox "No estas incurriendo ninguna tarea.", vbExclamation, "Incurridos Excel"
        Application.OnTime Now + TimeValue("00:10:00"), "Reminder"
    End If
    
End Sub
