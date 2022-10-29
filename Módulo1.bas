Attribute VB_Name = "Módulo1"
' Celda que almacena el tiempo incurrido (cronómetro).
Public timerRange As String
' Celda que almacena la tarea a incurrir.
Public incurredTaskRange As String
' Celda que almacena el número de la semana actual del año.
Public weekNumRange As String
' Celda que almacena la jornada laboral del usuario.
Public dailyShiftRange As String
' Rango de celdas que almacenan los tiempos incurridos durante la semana.
Public incurredTimesRange As String
' Tarea que se está incurriendo.
Public incurredTask As String
' Cronómetro encendido.
Public timerOn As Boolean

' Establece la actualización del cronómetro.
Public Sub SetTimer()

    Application.OnTime Now + TimeValue("00:00:01"), "MoveTimer"

End Sub

' Actualiza la celda del cronómetro y el texto de la barra de estado,
' además de volver a llamar a la función "SetTimer()", generando un bucle.
Public Sub MoveTimer()
    
    If (timerOn = True) Then
        Worksheets(1).range(timerRange).Value = Worksheets(1).range(timerRange).Value + TimeValue("00:00:01")
        Application.StatusBar = incurredTask + ": " + Format(range(timerRange).Value, "hh:mm:ss")
        Call SetTimer
    End If

End Sub

' Restablece el cronómetro y, si corresponde, la barra de estado.
Public Sub ResetTimer()
    
    If (timerOn = False) Then
        Worksheets(1).range(timerRange).Value = TimeValue("00:00:00")
        Application.StatusBar = "No estas realizando ninguna tarea."
    Else
        Worksheets(1).range(timerRange).Value = TimeValue("00:00:01")
    End If

End Sub

' Suma el tiempo incurrido a la celda correspondiente a la tarea y al día de la semana.
' Devuelve el rango de la celda mencionada.
Public Function SetCell() As String
    
    Dim row As Integer
    Dim column As Integer
    Dim incurredTime As Double
    
    For row = 9 To 50
        If (Cells(row, 2).Value = incurredTask Or Cells(row, 2).Value = "") Then
            Exit For
        End If
    Next
    
    column = Weekday(Date, 2) + 2
    
    incurredTime = Worksheets(1).range(timerRange).Value * 24
    
    Cells(row, column).Value = Cells(row, column).Value + incurredTime
    
    Dim incurredTimesRange As String
    incurredTimesRange = Cells(row, column).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    
    SetCell = incurredTimesRange
    
End Function

' Redondea el tiempo incurrido de la última tarea si faltan 15 o menos minutos para completar la jornada laboral.
' incurredTimesRange: Rango de la celda a redondear.
Public Sub RoundResult(incurredTimesRange As String)
    
    Dim dailyShift As Double
    dailyShift = Worksheets(1).range(dailyShiftRange).Value
    
    If (range(incurredTimesRange).Value < dailyShift And range(incurredTimesRange).Value >= (dailyShift - 0.25)) Then
        range(incurredTimesRange).Value = range(incurredTimesRange).Value + (dailyShift - range(incurredTimesRange).Value)
    End If

End Sub

' Lanza un mensaje de aviso si no se está incurriendo ninguna tarea.
Public Sub Reminder()
    
    If (incurredTask = "") Then
        MsgBox "No estas incurriendo ninguna tarea.", vbExclamation, "Sistema de incurridos"
    End If
    
End Sub

' Pregunta la jornada laboral del usuario para hacer uso de la función "RoundResult()".
Public Sub AskDailyShift()

    Worksheets(1).range(dailyShiftRange).Value = Application.InputBox("Por favor, introduce tu jornada laboral en horas." + _
    vbCrLf + "Por ejemplo: 7,5", "Sistema de incurridos", Type:=1)
    
    If (Worksheets(1).range(dailyShiftRange).Value = FALSO) Then
        Worksheets(1).range(dailyShiftRange).Value = 0
    End If

End Sub
