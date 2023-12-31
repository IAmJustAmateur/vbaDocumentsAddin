VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTestTableFieldsReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IFieldsReader
Public Sub ifieldsReader_setParams(Optional ByVal sPath As String = "", Optional ByVal sCardWBName As String = "", Optional sCardShName = "")

End Sub

Public Function IFieldsReader_getFields() As Scripting.Dictionary
    Dim fields As New Scripting.Dictionary
    Dim dworksAmount As Double
    Dim dPartsAmount As Double
    Dim dVatRate As Double
    Dim dVatAmount As Double
    Dim dTotalAmount As Double
    Dim cTotalAmount As Currency
    Dim damount As Double
    
    fields.Add "����� ������ � �����-������", Array(1, 2)
    fields.Add "������������ ������", Array("������ 1", "������ 2")
    fields.Add "����������", Array(3, 4)
    fields.Add "��������� ������", Array(Format(30, "#,##0.00"), Format(40, "#,##0.00"))
    fields.Add "���� ��������", Array(Format(50, "#,##0.00"), Format(60, "#,##0.00"))
    fields.Add "������������ ��������", Array("�������� 1", "�������� 2")
    
    dworksAmount = 3 * 30 + 40
    dPartsAmount = 3 * 50 + 4 * 60
    
    fields.Add "��������� �����", Array(Format(3 * 30, "#,##0.00"), Format(4 * 40, "#,##0.00"))
    fields.Add "��������� ���������", Array(Format(3 * 50, "#,##0.00"), Format(4 * 60, "#,##0.00"))
    
    fields.Add "�����", "4-121"
    fields.Add "����", "30.08.2023"
    fields.Add "��������.������������ �������", "��� ���������"
    fields.Add "����� ������", "AX1111-7"
    fields.Add "��������� ����� ������", Format(dworksAmount, "#,##0.00")
    fields.Add "��������� ��������� ������", Format(dPartsAmount, "#,##0.00")
        
    damount = dPartsAmount + dworksAmount
    fields.Add "�����", Format(damount, "#,##0.00")
    
    dVatRate = 0.2
    dVatAmount = dVatRate * damount
    fields.Add "����� ���", Format(dVatAmount, "#,##0.00")
    
    dTotalAmount = damount + dVatAmount
    fields.Add "����� c ���", dTotalAmount
    fields.Add "��������", "������ �.�."
    
    fields.Add "����� �������", "������ �.�."
    
    cTotalAmount = CCur(dTotalAmount)
    fields.Add "����� ��������", �������������.�������������������(cTotalAmount, False)
    
    Set IFieldsReader_getFields = fields
End Function

Public Function IFieldsReader_emptyFields() As Boolean

End Function


