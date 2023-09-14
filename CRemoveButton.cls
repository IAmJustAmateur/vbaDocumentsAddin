VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CRemoveButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents RemoveGroup As MSForms.CommandButton
Attribute RemoveGroup.VB_VarHelpID = -1

Private parentForm As OrderForm
Private irow As Integer

Public Property Set control(button As MSForms.CommandButton)
    Set RemoveGroup = button
End Property

Public Property Get control() As MSForms.CommandButton
    Set control = RemoveGroup
End Property


Public Property Set form(f As OrderForm)
    Set parentForm = f
End Property

Public Property Get form() As OrderForm
    Set form = parentForm
End Property

Public Sub RemoveGroup_click()
    Dim fr As MSForms.frame
    Dim ctl As control
    Dim sRowFrameName As String
    Set fr = RemoveGroup.Parent
    sRowFrameName = Mid(fr.name, Len("RowFrame") + 1)
    If sRowFrameName = "1" Then
        Exit Sub
    End If
    
    form.RowFrames.Remove sRowFrameName
    For Each ctl In fr.Controls
        form.Controls.Remove ctl.name
    Next ctl
    form.Controls.Remove fr.name
    form.enumerateFrames
    form.redraw
    form.setTotalAmount

End Sub

