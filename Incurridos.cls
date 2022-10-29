VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub startButton_Click()
    
    ' Si no se está incurriendo ninguna tarea, iniciar el timer,
    ' de lo contrario, guardar el tiempo incurrido y resetear el timer.
    If (incurredTask = "") Then
        timerOn = True
        Call SetTimer
    Else
        Call SetCell
        Call ResetTimer
    End If
    
    ' Almacena la tarea a incurrir.
    incurredTask = Worksheets(1).range(incurredTaskRange).Value
    
    ' Selecciona una celda para evitar el parpadeo del cuadro de nombres.
    ActiveSheet.Cells(4, 2).Select
    
End Sub

Private Sub stopButton_Click()
    
    ' Si no se está incurriendo ninguna tarea, indicarselo al usuario,
    ' de lo contrario, guardar el tiempo incurrido, redondear el resultado si corresponde
    ' y resetear el timer y la tarea a incurrir.
    If (incurredTask = "") Then
        MsgBox "No estas realizando ninguna tarea.", vbInformation, "Incurridos Excel"
    Else
        Call RoundResult(SetCell)
        incurredTask = ""
        timerOn = False
        Call ResetTimer
    End If
    
    ' Establece el recordatorio de incurrimiento.
    Application.OnTime Now + TimeValue("00:10:00"), "Reminder"
        
End Sub
