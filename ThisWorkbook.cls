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

    Worksheets(1).Protect "", UserInterfaceOnly:=True

    ' Inicializa las variables del Módulo1.
    timerRange = "B2"
    incurredTaskRange = "B4"
    weekNumRange = "B1"
    dailyShiftRange = "B7"
    incurredTimesRange = "C9:I50"
    incurredTask = ""
    timerOn = False
    Application.StatusBar = "No estas realizando ninguna tarea."
    
    ' Obtiene y almacena el número correspondiente a la semana actual del año.
    Dim weekNum As Integer
    weekNum = Application.WorksheetFunction.weekNum(Now, 2)
    
    ' Si el número de la semana obtenido es diferente al almacenado en el excel,
    ' elimina los tiempos incurridos durante la semana y sobrescribe el valor del excel.
    If (CStr(weekNum) <> Worksheets(1).range(weekNumRange).Value) Then
        Worksheets(1).range(incurredTimesRange).Value = ""
        Worksheets(1).range(weekNumRange).Value = CStr(weekNum)
    End If
    
    ' Si la celda que almacena la jornada laboral está vacía,
    ' la solicita al usuario.
    If (Worksheets(1).range(dailyShiftRange).Value = "") Then
        Call AskDailyShift
    End If
    
    ' Establece el recordatorio de incurrimiento.
    Application.OnTime Now + TimeValue("00:10:00"), "Reminder"

End Sub

' Controlador del evento generado al cambiar la tarea incurrida.
' Almacena el tiempo incurrido, cambia la tarea a incurrir y resetea el timer.
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As range)

    If (Target.Count > 1) Or (Target.Address <> Worksheets(1).range(incurredTaskRange).Address) Then
        Exit Sub
    End If
    
    ' Impide que el usuario deje vacía la tarea a incurrir.
    If (Target = "") Then
        If (incurredTask <> "") Then
            Target = incurredTask
        Else
            If (Worksheets(2).Cells(1, 1).Value <> "") Then
                Target = Worksheets(2).Cells(1, 1).Value
            End If
        End If
    End If
    
    ' Facilita el incurrimiento de otra tarea.
    If (Target <> "") And (timerOn = True) Then
        Call SetCell
        Call ResetTimer
        incurredTask = Target.Value
    End If

End Sub

' Controlador del evento generado al cerrar el workbook.
' Almacena el tiempo incurrido si se estuviera incurriendo alguna tarea.
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    
    If (timerOn = True) Then
        Call SetCell
        timerOn = False
        Call ResetTimer
    End If

End Sub
