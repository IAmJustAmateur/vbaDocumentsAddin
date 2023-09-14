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
    
    fields.Add "Номер строки в заказ-наряде", Array(1, 2)
    fields.Add "Наименование работы", Array("Работа 1", "Работа 2")
    fields.Add "Количество", Array(3, 4)
    fields.Add "Стоимость работы", Array(Format(30, "#,##0.00"), Format(40, "#,##0.00"))
    fields.Add "Цена запчасти", Array(Format(50, "#,##0.00"), Format(60, "#,##0.00"))
    fields.Add "Наименование запчасти", Array("Запчасть 1", "Запчасть 2")
    
    dworksAmount = 3 * 30 + 40
    dPartsAmount = 3 * 50 + 4 * 60
    
    fields.Add "Стоимость работ", Array(Format(3 * 30, "#,##0.00"), Format(4 * 40, "#,##0.00"))
    fields.Add "Стоимость запчастей", Array(Format(3 * 50, "#,##0.00"), Format(4 * 60, "#,##0.00"))
    
    fields.Add "Номер", "4-121"
    fields.Add "Дата", "30.08.2023"
    fields.Add "Заказчик.Наименование краткое", "ООО «Заказчик»"
    fields.Add "Номер машины", "AX1111-7"
    fields.Add "Стоимость работ заказа", Format(dworksAmount, "#,##0.00")
    fields.Add "Стоимость запчастей заказа", Format(dPartsAmount, "#,##0.00")
        
    damount = dPartsAmount + dworksAmount
    fields.Add "Сумма", Format(damount, "#,##0.00")
    
    dVatRate = 0.2
    dVatAmount = dVatRate * damount
    fields.Add "Сумма НДС", Format(dVatAmount, "#,##0.00")
    
    dTotalAmount = damount + dVatAmount
    fields.Add "Сумма c НДС", dTotalAmount
    fields.Add "Водитель", "Иванов И.И."
    
    fields.Add "Заказ оформил", "Петров П.П."
    
    cTotalAmount = CCur(dTotalAmount)
    fields.Add "Сумма прописью", СуммаПрописью.СуммаРублейПрописью(cTotalAmount, False)
    
    Set IFieldsReader_getFields = fields
End Function

Public Function IFieldsReader_emptyFields() As Boolean

End Function


