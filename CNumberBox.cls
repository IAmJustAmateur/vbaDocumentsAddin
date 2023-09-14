VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNumberBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Implements TextBox for whole numbers, min and max values
Private iMax As Integer
Private iMin As Integer
Private irow As Integer

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
