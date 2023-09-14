VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CPriceTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private irow As Integer
Private parentForm As OrderForm
Public WithEvents PriceGroup As MSForms.TextBox
Attribute PriceGroup.VB_VarHelpID = -1

Public Property Set control(tb As MSForms.TextBox)
    Set PriceGroup = tb
End Property

Public Property Set form(f As OrderForm)
    Set parentForm = f
End Property

Public Property Get form() As OrderForm
    Set form = parentForm
End Property


Private Sub PriceGroup_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' only digits and comma (or dot)
    Dim dprice As Double
    Dim sPrice As String
    Dim C As String
    Dim iCommaPos As Integer
    
    On Error Resume Next
        C = Chr(KeyAscii)
        If C = "," Then C = ",0"
        sPrice = PriceGroup.Text + C
        dprice = CDbl(sPrice)
        If Err.Number > 0 Then
            KeyAscii = 0
        End If
        iCommaPos = InStr(PriceGroup.Text, ",")
        If iCommaPos <> 0 Then
            If Len(PriceGroup.Text) - iCommaPos >= 2 Then
                KeyAscii = 0
            End If
        End If
    On Error GoTo 0

End Sub

Private Sub PriceGroup_Change()
    Dim frame As MSForms.frame
    Dim rf As CRowFrame
    Dim sFrameName As String
    Set frame = PriceGroup.Parent
    sFrameName = Mid(frame.name, Len("RowFrame") + 1)
    Set rf = form.RowFrames(sFrameName)
    rf.setAmount
    form.setTotalAmount
End Sub

