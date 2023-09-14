VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "COrderForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private form As OrderForm
Private loaders As Iloaders
Private customers As Collection
Private works As Collection

' Document params
Private orderDocGenerator As IDocumentTemplate

Private docTablesSelector As ITables
Private docNameGenerator As INameGenerator
Private sDocTemplateName As String
Private docFields As Scripting.Dictionary

Private sDocTemplatePath As String
Private sDocPath As String
Private bCloseDocument As Boolean
Private lFieldSeparator As String
Private rFieldSeparator As String
Private docTables As ITables

Public Property Set oForm(f As OrderForm)
    Set form = f
End Property

Public Property Get oForm()
    Set oForm = form
End Property

Public Property Get form_loaders() As Iloaders
    Set form_loaders = loaders
End Property

Public Sub setDocParams(DocGenerator As IDocumentTemplate _
                        , tablesSelector As ITables _
                        , nameGenerator As INameGenerator _
                        , sTemplateName As String _
                        , fields As Scripting.Dictionary _
                        , Optional sTemplatePath = "" _
                        , Optional sDocumentPath = "" _
                        , Optional bCloseAferFilling = True _
                        , Optional leftFieldSeparator = "{{" _
                        , Optional rightFieldSeparator = "}}" _
                        , Optional tables As ITables = Nothing)
                        
    Set orderDocGenerator = DocGenerator
    Set docTablesSelector = tablesSelector
    Set docNameGenerator = nameGenerator
    sDocTemplateName = sTemplateName
    Set docFields = fields
    
    sDocTemplatePath = sTemplatePath
    sDocPath = sDocumentPath
    bCloseDocument = bCloseAferFilling
    lFieldSeparator = leftFieldSeparator
    rFieldSeparator = rightFieldSeparator
    Set docTables = tables

End Sub


Public Property Set form_loaders(new_loaders As Iloaders)
    Set loaders = new_loaders
End Property
Public Property Get orderWorks() As Collection
    Set orderWorks = works
End Property

Public Property Set orderWorks(cworks As Collection)
    Set works = cworks
End Property

Public Sub showForm()
    
    form.load_initial_values loaders
    form.setDocumentParams orderDocGenerator, _
                            sDocTemplateName, _
                            docTablesSelector, docNameGenerator, _
                            docFields, _
                            sDocTemplatePath, sDocPath, bCloseDocument, _
                            lFieldSeparator, rFieldSeparator
    form.Show False
    
End Sub

Private Sub Class_Initialize()
    Set form = New OrderForm

End Sub
