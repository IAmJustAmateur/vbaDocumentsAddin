VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDefaultDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' implements document filling by default

Public doc As Document

Public sDocPath As String
Public sDocTemplatePath As String
Private sTemplateName As String
Private sDocumentName As String

Public lFieldSeparator As String
Public rFieldSeparator As String

Private docFields As Scripting.Dictionary
Private tableFields As Scripting.Dictionary
Private nameGenerator As INameGenerator
Private fieldsReader As IFieldsReader

Private docTablesSelector As ITables ' should be removed in future, one needs to parse all the tables
Private docTables As Variant

Implements IDocumentTemplate

Private Sub Class_Initialize()
    Set nameGenerator = New CDefaultNameGenerator
End Sub
Public Property Get documentName() As String
    documentName = sDocumentName
End Property

Public Property Get IDocumentTemplate_tables() As Variant
    IDocumentTemplate_tables = docTables
End Property

Public Property Get IDocumentTemplate_docTableFields() As Scripting.Dictionary
    Set IDocumentTemplate_docTableFields = tableFields
End Property

Public Property Get IDocumentTemplate_documentName() As String
    IDocumentTemplate_documentName = doc.name
End Property
Public Property Set IDocumentTemplate_fields(docFields As Scripting.Dictionary)
    Set docFields = docFields
End Property

Public Property Get IDocumentTemplate_fields() As Scripting.Dictionary
    Set IDocumentTemplate_fields = docFields
End Property

Public Property Set setNameGenerator(newNameGenerator As INameGenerator)
    Set nameGenerator = newNameGenerator
End Property

Public Property Get documentFieldsReader() As IFieldsReader
    Set documentFieldsReader = fieldsReader
End Property

Public Property Set documentFieldsReader(newFieldsReader As IFieldsReader)
    Set fieldsReader = newFieldsReader
End Property

Public Sub setPathParams(sDocumentTemplate As String, Optional sTemplatePath = "", Optional sDocumentPath = "")
    sTemplateName = sDocumentTemplate
    sDocPath = sDocumentPath
    sDocTemplatePath = sTemplatePath
End Sub

Public Sub setFieldValue(sFieldName As String, sFieldValue As String)
    
    replaceFieldWithValue doc, sFieldName, sFieldValue, lFieldSeparator, rFieldSeparator
    
End Sub

Public Sub fillTables()
' set values to all tableFields
    Dim sFieldName As String
    Dim i As Integer
             
    For i = 0 To UBound(tableFields.Keys)
        sFieldName = tableFields.Keys(i)
        setTableFieldValues lFieldSeparator + sFieldName + rFieldSeparator, tableFields.Item(sFieldName)
    Next i

End Sub

Public Sub setTableFieldValues(sFieldName As String, values As Variant)
    Dim t As Table
    Dim itable As Integer
    
    For itable = LBound(docTables) To UBound(docTables)
        Set t = docTables(itable)
        If isFieldInTable(sFieldName, t) Then
            Do While t.Rows.Count < UBound(values) + 1
                duplicate2ndRow t
            Loop
            replaceInTableBody t, sFieldName, values
        End If
    Next itable

End Sub

Public Sub fillDocumentFields()
    Dim i As Integer
    Dim sKey As String
    Dim sValue As String
    
    If Not docTablesSelector Is Nothing Then
        docTables = docTablesSelector.getTables(doc)
        If Not isEmptyArray(docTables) Then
            getTableFields
            fillTables
        End If
    End If
    
    For i = 0 To UBound(docFields.Keys)
        sKey = docFields.Keys(i)
        sValue = docFields.Item(sKey)
        setFieldValue sKey, sValue
    Next i
    
End Sub

Public Sub GenerateDocumentName()
   
   sDocumentName = nameGenerator.sNameGenerator(docFields) 'docFields(1).sFieldValue + ".docx" ' we assume that field 1 is date
End Sub


Public Sub fillDocument()
    Dim oWord As Word.Application
    
    Set oWord = getWordApplication
    
    On Error Resume Next
        oWord.Visible = True
        Set doc = oWord.Documents.Add(sDocTemplatePath + sTemplateName, newTemplate:=False, DocumentType:=0)
        If doc Is Nothing Then
            Err.Raise errors.err_document_template_does_not_exist, errors.msg_document_template_does_not_exist
            Exit Sub
        End If
    On Error GoTo 0
        
    fillDocumentFields
    
    GenerateDocumentName
    
    On Error Resume Next
        doc.SaveAs2 (sDocPath + sDocumentName)
        If Err.Number > 0 Then
            ' On Error Resume Next
            Err.Raise errors.err_can_not_save_document, errors.msg_can_not_save_document + " " + sDocPath + IDocumentTemplate_documentName()
            Exit Sub
        End If
    On Error GoTo 0
    
End Sub


Public Function IDocumentTemplate_createDocument(sDocumentTemplateName As String _
                            , fieldsReader As IFieldsReader _
                            , Optional sTemplatePath = "" _
                            , Optional sDocumentPath = "" _
                            , Optional bCloseAfterFilling = True _
                            , Optional leftFieldSeparator = "{{" _
                            , Optional rightFieldSeparator = "}}" _
                            , Optional tables As ITables = Nothing) As String
    
    setPathParams sDocumentTemplateName, sTemplatePath, sDocumentPath
    'Application.StatusBar = "Открываем карточку заказчика..."
    'fieldsReader.setParams sFieldsCardPath, sCustomerCardName
    Set docFields = fieldsReader.getFields()
    Set docTablesSelector = tables
    lFieldSeparator = leftFieldSeparator
    rFieldSeparator = rightFieldSeparator
    
    If fieldsReader.emptyFields Then
        Application.DisplayAlerts = False
        #If DEBUG_MODE = 1 Then
            On Error Resume Next
        #End If
        
            Err.Raise errors.err_card_does_not_exist, Description:=errors.msg_card_does_not_exist
            Exit Function
        
    End If
    
    Application.StatusBar = "Заполняем документ..."
    
    fillDocument
    
    If doc Is Nothing Then
        Application.DisplayAlerts = False
        Exit Function
    End If
    If bCloseAfterFilling Then
        doc.Close (False)
    End If
    Application.StatusBar = False
    
    IDocumentTemplate_createDocument = sDocumentName
    
End Function

Public Sub getTableFields()
    Dim field As Variant
    
    For Each field In docFields
        If IsArray(docFields(field)) Then
            If tableFields Is Nothing Then
                Set tableFields = New Scripting.Dictionary
            End If
            tableFields.Add field, docFields(field)
            docFields.Remove field
        End If
    Next field
    
End Sub
