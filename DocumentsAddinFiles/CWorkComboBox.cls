VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CWorkComboBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private irow As Integer

Public WithEvents WorkGroup As MSForms.ComboBox
Attribute WorkGroup.VB_VarHelpID = -1

Private tbWorkPrice As CPriceTextBox
Private tbPartPrice As CPriceTextBox
Private parentForm As OrderForm

Public Property Set form(f As OrderForm)
    Set parentForm = f
End Property

Public Property Get form() As OrderForm
    Set form = parentForm
End Property

Public Property Set control(cmb As MSForms.ComboBox)
    Set WorkGroup = cmb
End Property

Public Property Set workPriceBox(newPriceBox As CPriceTextBox)
    Set tbWorkPrice = newPriceBox
End Property

Public Property Get workPriceBox() As CPriceTextBox
    Set workPriceBox = tbWorkPrice
End Property

Public Property Set partPriceBox(newPriceBox As CPriceTextBox)
    Set tbPartPrice = newPriceBox
End Property

Public Property Get partPriceBox() As CPriceTextBox
    Set partPriceBox = tbPartPrice
End Property

Private Sub WorkGroup_Change()
    Dim sWork As String
    Dim iwork As Integer
        
    Dim work As iwork
    Dim dprice As Double
    
    sWork = WorkGroup.value
    iwork = WorkGroup.TopIndex
    Set work = form.getWork(sWork)
    
    If Not work Is Nothing Then
        tbWorkPrice.PriceGroup.Text = CStr(work.workPrice)
        tbPartPrice.PriceGroup.Text = CStr(work.partPrice)
    End If
    
End Sub

Public Sub load_works(works As Collection)
    Dim work As iwork
    If Not works Is Nothing Then
        For Each work In works
            WorkGroup.AddItem work.workName
        Next work
    End If

End Sub
