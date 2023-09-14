VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDateSpinButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Implements SpinButtons for Dates
' We assume that parent form contains 3 SpinButtons sbDay, sbMonth, sbYear

Private connectedNumberBox As CDateNumberBox
Private parentForm As OrderForm

Public WithEvents DateSpinGroup As MSForms.SpinButton
Attribute DateSpinGroup.VB_VarHelpID = -1

Public Property Set form(f As OrderForm)
    Set parentForm = f
End Property

Public Property Get form() As OrderForm
    Set form = parentForm
End Property

Public Property Set control(sb As MSForms.SpinButton)
    Set DateSpinGroup = sb
End Property

Public Property Set NumberBox(cnb As CDateNumberBox)
    Set connectedNumberBox = cnb
End Property

Public Property Get NumberBox()
    Set NumberBox = connectedNumberBox
End Property

Private Sub DateSpinGroup_Change()
    Dim d As Date
    
    On Error Resume Next
        d = DateSerial(form.sbYear.value, form.sbMonth.value, form.sbDay.value)
        ' TODO implement form method to set date value
        If Err.Number = 0 Then
            form.tbDay = Day(d)
            form.tbMonth = Month(d)
            form.tbYear = Year(d)
            
            form.sbDay.value = Day(d)
            form.sbMonth.value = Month(d)
            form.sbYear.value = Year(d)
        End If
    On Error GoTo 0
    
End Sub
