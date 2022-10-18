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

    If (task = "") Then
        timerOn = True
        Call ResetTimer
        Call SetTimer
    Else
        Call SetCell
        Call ResetTimer
    End If
    
    task = Worksheets(1).Range(taskCell).Value
    
End Sub

Private Sub stopButton_Click()

    If (task = "") Then
        MsgBox "No estas realizando ninguna tarea.", vbInformation, "Incurridos Excel"
    Else
        Call SetCell
        ' Call RoundResult
        task = ""
        timerOn = False
        Call ResetTimer
    End If
    
    Application.OnTime Now + TimeValue("00:10:00"), "Reminder"
        
End Sub
