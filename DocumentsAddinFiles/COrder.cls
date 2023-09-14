VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "COrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Implements works order
' fieldsReader provide order with fields from OrderForm

Public lOrderNumber As Long

Public orderDate As Date
Public sCustomerName As String
Public sCarNumber As String

Private dOrderAmount As Double
Private orderGenerator As CDefaultDocument
Private docFields As Scripting.Dictionary
Private orderFieldsReader As COrderFieldsReader
Private orderTables As ITables

Private sOrderTemplateName As String
Private sOrderTemplatePath As String
Private sOrderPath As String
Private bCloseOrder As Boolean
Private leftSeparator As String
Private rightSeparator As String

Public Property Get DocGenerator() As CDefaultDocument
    Set DocGenerator = orderGenerator
End Property

Public Property Set DocGenerator(DocGenerator As CDefaultDocument)
    Set orderGenerator = DocGenerator
End Property

Public Property Get fieldsReader() As COrderFieldsReader
    Set fieldsReader = orderFieldsReader
End Property

Public Property Set fieldsReader(fr As COrderFieldsReader)
    Set orderFieldsReader = fr
End Property

Private Sub Class_Initialize()

    Set orderGenerator = New CDefaultDocument
        
End Sub

    
Public Sub setParams(sDocTemplateName As String, sDocTemplatePath As String, sDocPath As String, _
                     bClose As Boolean, lFieldSeparator As String, rFieldSeparator As String, _
                     tables As ITables)
    sOrderTemplateName = sDocTemplateName
    sOrderTemplatePath = sDocTemplatePath
    sOrderPath = sDocPath
    bCloseOrder = bClose
    leftSeparator = lFieldSeparator
    rightSeparator = rFieldSeparator
    Set orderTables = tables

End Sub

Public Function createDocument()
    Dim sOrderName As String
    sOrderName = orderGenerator.IDocumentTemplate_createDocument(sOrderTemplateName, _
                    orderFieldsReader, sOrderTemplatePath, _
                    sOrderPath, False, "<<", ">>", orderTables)
    createDocument = sOrderName
End Function

