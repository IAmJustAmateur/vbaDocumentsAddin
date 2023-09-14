VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CRowFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public fr As MSForms.frame

Public cmbWork As CWorkComboBox
Public tbWorkPrice As CPriceTextBox
Public tbPartPrice As CPriceTextBox
Public cnbQty As CNumberBox
Public sbQty As CQtySpinButton
Public lblTotal As MSForms.Label
Public rmButton As CRemoveButton
Public index As Integer

Public lblWorkNumber As MSForms.Label
Public WithEvents work As MSForms.ComboBox
Attribute work.VB_VarHelpID = -1
Public WithEvents workPrice As MSForms.TextBox
Attribute workPrice.VB_VarHelpID = -1
Public WithEvents partPrice As MSForms.TextBox
Attribute partPrice.VB_VarHelpID = -1
Public WithEvents qty As MSForms.TextBox
Attribute qty.VB_VarHelpID = -1
Public WithEvents qtySpin As MSForms.SpinButton
Attribute qtySpin.VB_VarHelpID = -1
Public WithEvents btn As MSForms.CommandButton
Attribute btn.VB_VarHelpID = -1

Public sUnique As String

Private parentForm As OrderForm


Public Property Get form() As OrderForm
    Set form = parentForm
End Property
Public Property Set form(of As OrderForm)
    Set parentForm = of
End Property
Private Sub Class_Initialize()
    Dim dt As Date
    dt = Now
    sUnique = CStr(Hour(dt)) + CStr(Minute(dt)) + CStr(Second(dt)) + Mid(CStr(Rnd()), 3, 2)
End Sub


Public Sub create(parentForm As OrderForm)
    Dim lbl As MSForms.Label
   
    Set form = parentForm
    
    If form.RowFrames.Count = 1 Then
        Set fr = form.RowFrame1
        
        Set cmbWork = New CWorkComboBox
        Set cmbWork.control = form.cmbWorks1
        Set cmbWork.form = form
        
        Set tbWorkPrice = New CPriceTextBox
        Set tbWorkPrice.control = form.tbWorkPrice1
        Set tbWorkPrice.form = form
        Set cmbWork.workPriceBox = tbWorkPrice
        
        Set tbPartPrice = New CPriceTextBox
        Set tbPartPrice.control = form.tbPartPrice1
        Set tbPartPrice.form = form
        Set cmbWork.partPriceBox = tbPartPrice
        
        Set cnbQty = New CNumberBox
        Set cnbQty.form = form
        Set cnbQty.control = form.tbWorksQty1
        cnbQty.set_borders 1, 100
        Set sbQty = New CQtySpinButton
        Set sbQty.control = form.sbWorksQty1
        Set sbQty.NumberBox = cnbQty
        Set sbQty.form = form
        Set sbQty.NumberBox = cnbQty
        Set rmButton = New CRemoveButton
        Set rmButton.control = form.cbRemove1
        Set rmButton.form = form
        
        Set lblWorkNumber = form.lblbWorkNumber
        Set lblTotal = form.lblWorkAmount
        Set qty = form.tbWorksQty1
        Set workPrice = form.tbWorkPrice1
        Set partPrice = form.tbPartPrice1
        Set work = form.cmbWorks1
        Set qtySpin = form.sbWorksQty1
    Else
        Set fr = form.Controls.Add("forms.frame.1", 1, True)
        fr.name = "RowFrame" + sUnique
        
        fr.Top = form.RowFrames(form.RowFrames.Count - 1).fr.Top + form.RowFrames(form.RowFrames.Count - 1).fr.Height + 10
        fr.Left = form.RowFrame1.Left
        fr.Width = form.RowFrame1.Width
        fr.Height = form.RowFrame1.Height
        
        Set lblWorkNumber = createLblWorkNumber(sUnique)
        Set cmbWork = createWorkComboBox(sUnique)
        Set cnbQty = createQtyBox(sUnique)
        Set sbQty = createQtySpinButton(cnbQty, sUnique)
        Set tbPartPrice = createPartPriceTextBox(sUnique)
        Set tbWorkPrice = createWorkPriceTextBox(sUnique)
        Set lblTotal = createTotalLabel(sUnique)
        Set rmButton = createRemoveButton(sUnique)
        
    End If
    
End Sub

Public Function addWorkComboBox() As MSForms.ComboBox
    Set addWorkComboBox = form.Controls.Add("forms.combobox.1", "cmbWork" + CStr(index), True)
End Function

Public Function addPriceBox() As MSForms.TextBox
    Set addPriceBox = form.Controls.Add("forms.textbox.1", "tbWorkPrice" + CStr(index), True)
End Function

Public Function addQtyTextBox() As MSForms.TextBox
    Set addQtyTextBox = form.Controls.Add("forms.textbox.1", "tbWorksQty" + CStr(index), True)
End Function

Public Function addQtySpinButton() As MSForms.SpinButton
    Set addQtySpinButton = form.Controls.Add("forms.spinbutton.1", "sbWorksQty" + CStr(index), True)
End Function

Public Function addTotalLabel() As MSForms.Label
    Set addTotalLabel = form.Controls.Add("forms.label.1", "lblWorksAmount" + CStr(index), True)
End Function

Public Function addRemoveButton() As MSForms.CommandButton
    Set addRemoveButton = form.Controls.Add("forms.commandbutton.1", "btnWorkRemove" + CStr(index), True)
End Function

Public Sub loadWorks(works As Collection)
    Dim work As iwork
    For Each work In works
        cmbWork.WorkGroup.AddItem work.workName
    Next work
End Sub
Public Function createLblWorkNumber(sUnique As String) As MSForms.Label
    Dim lbl As MSForms.Label
    Set lbl = fr.Controls.Add("forms.label.1", 1, True)
    lbl.name = "lblWorkAmount" + sUnique
    With lbl
        .Top = 5
        .Left = 5
        .Width = 5
        .Caption = CStr(getFrameNumber)
    End With
    Set createLblWorkNumber = lbl
    
End Function

Public Function createWorkComboBox(sUnique As String) As CWorkComboBox
    Dim cmbWork As New CWorkComboBox
    Set work = fr.Controls.Add("forms.combobox.1", 1, True)
    work.name = "cmbWorks" + sUnique
    With work
        .Top = 5
        .Left = form.cmbWorks1.Left
        .Width = form.cmbWorks1.Width
    End With
    Set cmbWork.control = work
    Set cmbWork.form = form
    cmbWork.load_works form.works
    Set createWorkComboBox = cmbWork
    
End Function


Public Function createPartPriceTextBox(sUnique As String) As CPriceTextBox
    Dim tbPartPrice As New CPriceTextBox
    Set partPrice = fr.Controls.Add("forms.textbox.1", 1, True)
    partPrice.name = "tbPartPrice" + sUnique
    With partPrice
        .Top = 5
        .Left = form.tbPartPrice1.Left
        .Width = form.tbPartPrice1.Width
    End With
    Set tbPartPrice.control = partPrice
    Set tbPartPrice.form = form
    Set cmbWork.partPriceBox = tbPartPrice
    Set createPartPriceTextBox = tbPartPrice
End Function


Public Function createWorkPriceTextBox(sUnique As String) As CPriceTextBox
    Dim tbWorkPrice As New CPriceTextBox
    Set workPrice = fr.Controls.Add("forms.textbox.1", 1, True)
    workPrice.name = "tbWorkPrice" + sUnique
    With workPrice
        .Top = 5
        .Left = form.tbWorkPrice1.Left
        .Width = form.tbWorkPrice1.Width
    End With
    Set tbWorkPrice.control = workPrice
    Set tbWorkPrice.form = form
    Set cmbWork.workPriceBox = tbWorkPrice
    Set createWorkPriceTextBox = tbWorkPrice
End Function

Public Function createQtyBox(sUnique As String) As CNumberBox
    Dim nbQty As New CNumberBox
    Set qty = fr.Controls.Add("forms.textbox.1", 1, True)
    qty.name = "tbWorksQty" + sUnique
    With qty
        .Top = 5
        .Left = form.tbWorksQty1.Left
        .Width = form.tbWorksQty1.Width
    End With
    Set nbQty.control = qty
    Set nbQty.form = form
    nbQty.set_borders 1, 100
    Set createQtyBox = nbQty
End Function

Public Function createQtySpinButton(cnb As CNumberBox, sUnique As String) As CQtySpinButton
    Dim sqb As New CQtySpinButton
    Set qtySpin = fr.Controls.Add("forms.spinbutton.1", 1, True)
    qtySpin.name = "sbQty" + sUnique
    With qtySpin
        .Top = 5
        .Left = form.sbWorksQty1.Left
    End With
    Set sqb.control = qtySpin
    Set sqb.form = form
    Set sqb.NumberBox = cnb
    Set createQtySpinButton = sqb

End Function

Public Function createTotalLabel(sUnique As String) As MSForms.Label
    Dim lbl As MSForms.Label
    Set lbl = fr.Controls.Add("forms.label.1", 1, True)
    lbl.name = "lblWorkAmount" + sUnique
    With lbl
        .Top = 5
        .Left = form.lblWorkAmount.Left
        .Width = form.lblWorkAmount.Width
        .Caption = "0.00"
    End With
    Set createTotalLabel = lbl
End Function

Public Function createRemoveButton(sUnique As String) As CRemoveButton
    Dim crb As New CRemoveButton
    Set btn = fr.Controls.Add("forms.commandbutton.1", 1, True)
    btn.name = "cbRemove" + sUnique
    With btn
        .Top = 5
        .Left = form.cbRemove1.Left
        .Width = form.cbRemove1.Width
        .Caption = "-"
        .Font.Size = 14
    End With
    Set crb.control = btn
    Set crb.form = form
    Set createRemoveButton = crb
End Function

Public Sub setAmount()
    Dim iqty As Integer
    Dim dWorkPrice As Double
    Dim dPartPrice As Double
    On Error Resume Next
        iqty = CInt(cnbQty.NumberGroup.Text)
        dWorkPrice = CDbl(tbWorkPrice.PriceGroup.Text)
        dPartPrice = CDbl(tbPartPrice.PriceGroup.Text)
    On Error GoTo 0
    lblTotal.Caption = CStr(iqty * (dWorkPrice + dPartPrice))
End Sub

Public Function getFrameNumber() As Integer
    Dim inumber As Integer
    Dim f As CRowFrame
    inumber = 1
    For Each f In form.RowFrames
        If f Is Me Then
            getFrameNumber = inumber
            Exit Function
        Else
            inumber = inumber + 1
        End If
    Next f
End Function
Public Property Get sWorkName() As String
    sWorkName = Me.work.Text
End Property
Public Property Get sPartName() As String
    sPartName = Me.form.getWork(sWorkName).partName
End Property


Public Property Get dWorkPrice() As Double
    dWorkPrice = CDbl(Me.workPrice.Text)
End Property
Public Property Get dPartPrice() As Double
    dPartPrice = CDbl(Me.partPrice.Text)
End Property
Public Property Get iWorksQty() As Integer
    On Error Resume Next
    iWorksQty = CInt(Me.qty.Text)
    On Error GoTo 0
End Property

