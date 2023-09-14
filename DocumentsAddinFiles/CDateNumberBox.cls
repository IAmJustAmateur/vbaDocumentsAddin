VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDateNumberBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Implements number boxes for Dates
' We assume that parent form contains 3 textboxes: tbDay, tbMonth, tbYear
Private iMax As Integer
Private iMin As Integer

Private parentForm As OrderForm

Public WithEvents NumberGroup As MSForms.TextBox
Attribute NumberGroup.VB_VarHelpID = -1

Public Property Set form(f As OrderForm)
    Set parentForm = f
End Property

Public Property Get form() As OrderForm
    Set form = parentForm
End Property

Public Property Set control(tb As MSForms.TextBox)
    Set NumberGroup = tb
End Property

Private Property Get min() As Integer
    min = iMin
End Property

Private Property Get max() As Integer
    max = iMax
End Property

Public Sub set_borders(ilow As Integer, ihigh As Integer)
    iMin = ilow
    iMax = ihigh
End Sub

Private Sub NumberGroup_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' only digits
    ' only numbers between iMin and iMax values
    Dim ivalue As Integer
    
    On Error Resume Next

        ivalue = CInt(NumberGroup.Text + Chr(KeyAscii))
        If (ivalue > iMax) Or (ivalue < iMin) Then
            KeyAscii = 0
        End If
        If Err.Number > 0 Then
            KeyAscii = 0
        End If
            
    On Error GoTo 0

End Sub

Private Sub NumberGroup_change()
    ' process the situation like 31.09 or 30.02 or some other nonexisting dates
    Dim d As Date
    
    If form.tbDay.value = "" Then
        Exit Sub
    End If
    If form.tbMonth.value = "" Then
        Exit Sub
    End If
    If form.tbYear.value = "" Then
        Exit Sub
    End If
    
    On Error Resume Next
        d = DateSerial(CInt(form.tbYear.value), CInt(form.tbMonth.value), CInt(form.tbDay.value))
    On Error GoTo 0
        
    form.tbYear.value = Year(d)
    form.sbYear.value = Year(d)
    
    form.tbMonth.value = Month(d)
    form.sbMonth.value = Month(d)
    
    form.tbDay.value = Day(d)
    form.sbDay.value = Day(d)

End Sub

