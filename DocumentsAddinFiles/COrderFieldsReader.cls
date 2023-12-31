VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "COrderFieldsReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IFieldsReader
Private form As OrderForm
Public Sub ifieldsReader_setParams(Optional ByVal sPath As String = "", Optional ByVal sCardWBName As String = "", Optional sCardShName = "")

End Sub
Public Property Set oForm(of As OrderForm)
    Set form = of
End Property
Public Property Get oForm() As OrderForm
    Set oForm = form
End Property

Public Function IFieldsReader_getFields() As Scripting.Dictionary
    Dim fields As New Scripting.Dictionary
    Dim iWorksCount As Integer
    Dim i As Integer
    Dim rf As CRowFrame
    
    Dim rowNumbers() As String
    Dim works() As String
    Dim parts() As String
    Dim worksQty() As String
    
    Dim worksPrices() As String
    Dim partPrices() As String
    Dim WorksAmounts() As String
    Dim PartsAmount() As String
    
    Dim dworksAmount As Double
    Dim dPartsAmount As Double
    Dim dVatRate As Double
    Dim dVatAmount As Double
    Dim dTotalAmount As Double
    Dim cTotalAmount As Currency
    Dim damount As Double
    
    iWorksCount = form.worksQty
    
    ReDim rowNumbers(1 To iWorksCount)
    ReDim works(1 To iWorksCount)
    ReDim parts(1 To iWorksCount)
    ReDim worksQty(1 To iWorksCount)
    
    ReDim workPrices(1 To iWorksCount)
    ReDim partPrices(1 To iWorksCount)
    ReDim WorksAmounts(1 To iWorksCount)
    ReDim partsAmounts(1 To iWorksCount)
    
    For i = 1 To iWorksCount
        Set rf = form.RowFrames(i)
        rowNumbers(i) = rf.getFrameNumber
        works(i) = rf.sWorkName
        parts(i) = rf.sPartName
        worksQty(i) = CStr(rf.iWorksQty)
        workPrices(i) = Format(rf.dWorkPrice, "#,##0.00")
        dworksAmount = dworksAmount + rf.dWorkPrice * rf.iWorksQty
        partPrices(i) = Format(rf.dPartPrice, "#,##0.00")
        dPartsAmount = dPartsAmount + rf.dPartPrice * rf.iWorksQty
        WorksAmounts(i) = Format(rf.iWorksQty * rf.dWorkPrice, "#,##0.00")
        partsAmounts(i) = Format(rf.iWorksQty * rf.dPartPrice, "#,##0.00")
    Next i
    
    fields.Add "����� ������ � �����-������", rowNumbers
    fields.Add "������������ ������", works
    fields.Add "����������", worksQty
    fields.Add "��������� ������", workPrices
    fields.Add "���� ��������", partPrices
    fields.Add "������������ ��������", parts
    
    fields.Add "��������� �����", WorksAmounts
    fields.Add "��������� ���������", partsAmounts
    
    fields.Add "�����", form.lblOrderNumber.Caption
    fields.Add "����", form.sOrderDate
    fields.Add "��������.������������ �������", form.cmbCustomers.Text
    fields.Add "����� ������", form.cmbCarNumber.Text
    fields.Add "��������� ����� ������", Format(dworksAmount, "#,##0.00")
    fields.Add "��������� ��������� ������", Format(dPartsAmount, "#,##0.00")
        
    damount = dPartsAmount + dworksAmount
    fields.Add "�����", Format(damount, "#,##0.00")
    
    dVatRate = 0.2
    dVatAmount = dVatRate * damount
    fields.Add "����� ���", dVatAmount
    
    dTotalAmount = damount + dVatAmount
    fields.Add "����� c ���", Format(dTotalAmount, "#,##0.00")
    
    fields.Add "��������", form.tbCustomerRepr.Text
    
    fields.Add "����� �������", "������ �.�."
    
    cTotalAmount = CCur(dTotalAmount)
    fields.Add "����� ��������", �������������.�������������������(cTotalAmount, False)
    
    Set IFieldsReader_getFields = fields
End Function

Public Function IFieldsReader_emptyFields() As Boolean

End Function



