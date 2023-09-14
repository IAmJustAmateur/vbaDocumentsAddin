VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CQtySpinButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private connectedNumberBox As CNumberBox
Private parentForm As OrderForm

Public WithEvents NumberSpinGroup As MSForms.SpinButton
Attribute NumberSpinGroup.VB_VarHelpID = -1

Public Property Set form(f As OrderForm)
    Set parentForm = f
End Property

Public Property Get form() As OrderForm
    Set form = parentForm
End Property

Public Property Set control(sb As MSForms.SpinButton)
    Set NumberSpinGroup = sb
End Property

Public Property Set NumberBox(cnb As CNumberBox)
    Set connectedNumberBox = cnb
End Property

Public Property Get NumberBox()
    Set NumberBox = connectedNumberBox
End Property

Private Sub NumberSpinGroup_Change()
    Dim frame As MSForms.frame
    Dim rf As CRowFrame
    Dim sFrameName As String
    connectedNumberBox.NumberGroup.Text = CStr(Me.NumberSpinGroup.value)
    Set frame = NumberSpinGroup.Parent
    sFrameName = Mid(frame.name, Len("RowFrame") + 1)
    Set rf = form.RowFrames(sFrameName)
    rf.setAmount
    form.setTotalAmount
End Sub

