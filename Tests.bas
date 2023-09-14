Attribute VB_Name = "Tests"
Option Explicit
Public Function getDefaultPathForTests() As String
' необходимо переопределить и указать свой путь для тестов
    Dim sPath As String
    sPath = Environ("Onedrive") + Application.PathSeparator + "Documents" + Application.PathSeparator + "vbaDocumentsAddin" + Application.PathSeparator
    getDefaultPathForTests = sPath
End Function

Public Sub test_isStrInDoc()
    Dim sPath As String
    Dim oWord As Word.Application
    Dim sDocName As String
    Dim doc As Document
    Dim sInDoc  As String
    Dim sNotInDoc As String
    
    sInDoc = "ЗАКАЗ"
    sNotInDoc = "Hello, my dear!"
    sDocName = "order_template.docx"
    
    Set oWord = getWordApplication()
    sPath = getDefaultPathForTests() + sDocName
    Set doc = oWord.Documents.Open(sPath)
    
    Debug.Assert isStrInDoc(sInDoc, doc) = True
    Debug.Assert isStrInDoc(sNotInDoc, doc) = False
    doc.Close (False)
   
End Sub

Public Sub test_create_contract_successfull()
    Dim contract As New CDefaultDocument
    Dim fieldsReader As New CFieldsReaderFromXL
    Dim sPath As String
    Dim sCustomerCardPath As String
    Dim sContractTemplatePath As String
    Dim sContractPath As String
    Dim sfield As Variant
    
    Dim fields As Scripting.Dictionary
    Dim icounter As Integer
        
    sPath = getDefaultPathForTests()
    sCustomerCardPath = sPath + Application.PathSeparator
    sContractTemplatePath = sPath + Application.PathSeparator
    sContractPath = sPath + Application.PathSeparator
    
    fieldsReader.setParams sPath, "test_customer_card.xlsx"
    
    contract.IDocumentTemplate_createDocument "test_contract.docx", _
                                               fieldsReader, _
                                               sTemplatePath:=sContractTemplatePath, _
                                               sDocumentPath:=sContractPath, _
                                               bCloseAfterFilling:=False
                                               
    Debug.Assert Not contract.doc Is Nothing
    
    Set fields = contract.IDocumentTemplate_fields()
    
    Debug.Assert contract.doc.name = fields.Item("номер договора") + ".docx"
        
    For Each sfield In fields.Keys()
        icounter = 0
        With contract.doc.Content.Find
            .Text = contract.lFieldSeparator + sfield + contract.rFieldSeparator
            .Forward = True
            Do While .Execute
                icounter = icounter + 1
            Loop
        End With
        Debug.Assert icounter = 0
        
    Next sfield
    
    contract.doc.Close (False)

End Sub

Public Sub test_create_contract_unsuccessfull_card_does_not_exist()
    Dim contract As New CDefaultDocument
    Dim fieldsReader As New CFieldsReaderFromXL
    Dim sPath As String
    Dim sCustomerCardPath As String
    Dim sContractTemplatePath As String
    Dim sContractPath As String
        
    sPath = getDefaultPathForTests
    sCustomerCardPath = sPath + Application.PathSeparator
    sContractTemplatePath = sPath + Application.PathSeparator
    sContractPath = sPath + Application.PathSeparator
    fieldsReader.setParams sPath, "test_customer_card_0.xlsx"
    
    contract.IDocumentTemplate_createDocument "test_contract.docx", _
                                               fieldsReader, _
                                               sTemplatePath:=sContractTemplatePath, _
                                               sDocumentPath:=sContractPath, _
                                               bCloseAfterFilling:=True
    Debug.Assert Err.Number = errors.err_card_does_not_exist
   
End Sub

Public Sub test_create_contract_unsuccessfull_template_does_not_exist()
    Dim contract As New CDefaultDocument
    Dim fieldsReader As New CFieldsReaderFromXL
    Dim sPath As String
    Dim sCustomerCardPath As String
    Dim sContractTemplatePath As String
    Dim sContractPath As String
    
    sPath = getDefaultPathForTests
    sCustomerCardPath = sPath + Application.PathSeparator
    sContractTemplatePath = sPath + Application.PathSeparator
    sContractPath = sPath + Application.PathSeparator
    
    fieldsReader.setParams sPath, "test_customer_card.xlsx"
    
    contract.IDocumentTemplate_createDocument "test_contract_0.docx", _
                                               fieldsReader, _
                                               sTemplatePath:=sContractTemplatePath, _
                                               sDocumentPath:=sContractPath, _
                                               bCloseAfterFilling:=True
    Debug.Assert Err.Number = errors.err_document_template_does_not_exist
   
End Sub

Public Sub test_create_contract_unsuccessfull_path_to_document_does_not_exist()
    Dim contract As New CDefaultDocument
    Dim fieldsReader As New CFieldsReaderFromXL
    Dim sPath As String
    Dim sCustomerCardPath As String
    Dim sContractTemplatePath As String
    Dim sContractPath As String
        
    sPath = getDefaultPathForTests
    sCustomerCardPath = sPath + Application.PathSeparator
    sContractTemplatePath = sPath + Application.PathSeparator
    sContractPath = "non_existing_folder" + Application.PathSeparator
    
    fieldsReader.setParams sPath, "test_customer_card.xlsx"
    contract.IDocumentTemplate_createDocument "test_contract.docx", _
                                               fieldsReader, _
                                               sTemplatePath:=sContractTemplatePath, _
                                               sDocumentPath:=sContractPath, _
                                               bCloseAfterFilling:=True
    Debug.Assert Err.Number = errors.err_can_not_save_document
   
End Sub
Public Sub test_name_in_genitive()
    Dim sFullName As String
    Dim sNameInGenitive As String
    
    sFullName = "Иванов Иван Иванович"
    sNameInGenitive = FullNameInGenitive(sFullName)
    Debug.Assert sNameInGenitive = "Иванова Ивана Ивановича"
    
    sFullName = "Иванова Мария Ивановна"
    sNameInGenitive = FullNameInGenitive(sFullName)
    Debug.Assert sNameInGenitive = "Ивановой Марии Ивановны"
    
    sFullName = "Иванов Игорь Иванович"
    sNameInGenitive = FullNameInGenitive(sFullName)
    Debug.Assert sNameInGenitive = "Иванова Игоря Ивановича"
    
    sFullName = "Иванова Любовь Ивановна"
    sNameInGenitive = FullNameInGenitive(sFullName)
    Debug.Assert sNameInGenitive = "Ивановой Любови Ивановны"
    
    sFullName = "Петрович Любовь Ивановна"
    sNameInGenitive = FullNameInGenitive(sFullName)
    Debug.Assert sNameInGenitive = "Петрович Любови Ивановны"
    
    sFullName = "Петрович Иван Иванович"
    sNameInGenitive = FullNameInGenitive(sFullName)
    Debug.Assert sNameInGenitive = "Петровича Ивана Ивановича"
    
    sFullName = "Ивановский Иван Иванович"
    sNameInGenitive = FullNameInGenitive(sFullName)
    Debug.Assert sNameInGenitive = "Ивановского Ивана Ивановича"
    
    sFullName = "Ивановская Любовь Ивановна"
    sNameInGenitive = FullNameInGenitive(sFullName)
    Debug.Assert sNameInGenitive = "Ивановской Любови Ивановны"
    
    sFullName = "Черняк Иван Иванович"
    sNameInGenitive = FullNameInGenitive(sFullName)
    Debug.Assert sNameInGenitive = "Черняка Ивана Ивановича"
    
    sFullName = "Черняк Лариса Иванович"
    sNameInGenitive = FullNameInGenitive(sFullName)
    Debug.Assert sNameInGenitive = "Черняк Ларисы Ивановича"
    
    sFullName = "Наливайко Иван Иванович"
    sNameInGenitive = FullNameInGenitive(sFullName)
    Debug.Assert sNameInGenitive = "Наливайко Ивана Ивановича"
    
    sFullName = "Черняк Л.И."
    sNameInGenitive = FullNameInGenitive(sFullName)
    Debug.Assert sNameInGenitive = "Черняка Л.И."
    
    sFullName = "Ивановская Л.И."
    sNameInGenitive = FullNameInGenitive(sFullName)
    Debug.Assert sNameInGenitive = "Ивановской Л.И."
    
    sFullName = "Наливайко Л.И"
    sNameInGenitive = FullNameInGenitive(sFullName)
    Debug.Assert sNameInGenitive = "Наливайко Л.И"
     
End Sub

Public Sub test_replaceInTableBody()
    Dim sPath As String
    Dim orderFileName As String
    Dim oWord As New Word.Application
    Dim orderDoc As Document
    Dim documentTables As tables
    Dim t As Table
    Dim irow As Integer
    Dim newStrs(1 To 2) As String
    Dim sToReplace As String
    Dim sText As String
    Dim C As cell
    
    orderFileName = "order_template.docx"
    oWord.Visible = True
    
    sPath = getDefaultPathForTests
    Set orderDoc = oWord.Documents.Add(sPath + Application.PathSeparator + orderFileName)
    Set t = orderDoc.tables(3)
    duplicate2ndRow t
    newStrs(1) = "1."
    newStrs(2) = "2."
    sToReplace = "<<Номер строки в заказ-наряде>>"
    
    replaceInTableBody t, sToReplace, newStrs
    For irow = 2 To t.Rows.Count
        Set C = t.Rows(irow).Cells(1)
        sText = C.Range.Text
        ' removing vbCr, which word adds to the end of cell.range.text
        Debug.Assert Mid(C.Range.Text, 1, Len(C.Range.Text) - 2) = Mid(Replace(sText, sToReplace, newStrs(irow - 1)), 1, Len(sText) - 2)
        
    Next irow
    orderDoc.Close (False)
    
End Sub

Sub test_default_table_doc_generator()
    Dim tableDocGenerator As New CDefaultDocument
    Dim tableSelector As New CDefaultTables
    Dim nameGenerator As New CDefaultNameGenerator
    Dim fieldsReader As New CTestTableFieldsReader
    Dim t As Table
    
    Dim fields As Scripting.Dictionary
        
    Dim sPath As String
    Dim orderFileName As String
    Dim oWord As New Word.Application
    Dim orderDoc As Document
    Dim templateDoc As Document
    
    Dim sfield As String
    Dim i As Integer
    Dim tableFields As Scripting.Dictionary
        
    sPath = getDefaultPathForTests()
    orderFileName = "order_template.docx"
    
    tableDocGenerator.IDocumentTemplate_createDocument orderFileName, _
                                               fieldsReader, _
                                               sTemplatePath:=sPath, _
                                               sDocumentPath:=sPath, _
                                               bCloseAfterFilling:=False, _
                                               tables:=tableSelector, _
                                               leftFieldSeparator:="<<", rightFieldSeparator:=">>"
    
    Set tableFields = tableDocGenerator.IDocumentTemplate_docTableFields
    For i = LBound(tableDocGenerator.IDocumentTemplate_tables) To UBound(tableDocGenerator.IDocumentTemplate_tables)
        Set t = tableDocGenerator.IDocumentTemplate_tables(i)
        Debug.Assert t.Rows.Count = UBound(tableFields.Items(0)) - LBound(tableFields.Items(0)) + 1
    Next i
    
    Set templateDoc = oWord.Documents.Open(sPath + orderFileName)
    
    Set fields = tableDocGenerator.IDocumentTemplate_fields
    For i = LBound(fields.Keys) To UBound(fields.Keys)
        sfield = "<<" + fields.Keys(i) + ">>"
        Debug.Assert isStrInDoc(sfield, tableDocGenerator.doc) = False
    Next i
    
    tableDocGenerator.doc.Close (False)
    
End Sub

Public Sub all_tests()
    ' установить DEBUG_MODE = 1 в tools/ContractGenerator properties
    Dim dStart As Double
    Dim sMinutesElapsed As String
    
    Debug.Print "test execution starts: ", Now
    dStart = Timer
    
    test_name_in_genitive
    test_isStrInDoc
    
    test_create_contract_successfull
    test_create_contract_unsuccessfull_card_does_not_exist
    test_create_contract_unsuccessfull_path_to_document_does_not_exist
    test_create_contract_unsuccessfull_template_does_not_exist
    
    test_replaceInTableBody
    test_default_table_doc_generator
    
    sMinutesElapsed = Format((Timer - dStart) / 86400, "hh:mm:ss")

    Debug.Print "tests complete, execution time: ", sMinutesElapsed
    
End Sub


Public Sub test_duplicate2ndRow()
    Dim sPath As String
    Dim orderFileName As String
    Dim oWord As New Word.Application
    Dim orderDoc As Document
    Dim documentTables As tables
    Dim t As Table
    Dim i As Integer
    
    orderFileName = "order_template.docx"
    oWord.Visible = True
    
    sPath = getDefaultPathForTests
    Set orderDoc = oWord.Documents.Add(sPath + Application.PathSeparator + orderFileName)
    
    Set t = orderDoc.tables(3)
    duplicate2ndRow t
    Debug.Assert t.Rows.Count = 3
    orderDoc.Close (False)

End Sub


Public Sub test_form()
    Dim test_form As New COrderForm
    Dim loaders As New CTestLoaders
    
    Dim tableDocGenerator As New CDefaultDocument
    Dim tableSelector As New CDefaultTables
    Dim nameGenerator As New CDefaultNameGenerator
    Dim orderTemplateName As String
    Dim sPath As String
    Dim tables As New CDefaultTables
    
    Dim fields As Scripting.Dictionary
    Dim i As Integer
    
    sPath = getDefaultPathForTests()
    orderTemplateName = "order_template.docx"
        
    Set test_form.form_loaders = loaders
    test_form.setDocParams _
        tableDocGenerator, _
        tableSelector, _
        nameGenerator, _
        orderTemplateName, _
        fields, _
        sTemplatePath:=sPath, sDocumentPath:=sPath, bCloseAferFilling:=False, _
        leftFieldSeparator:="<<", rightFieldSeparator:=">>", tables:=tables
    
    test_form.showForm
    
End Sub
