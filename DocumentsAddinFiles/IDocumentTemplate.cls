VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDocumentTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Function createDocument(sDocumentTemplateName As String _
                            , fieldsReader As IFieldsReader _
                            , Optional sTemplatePath = "" _
                            , Optional sDocumentPath = "" _
                            , Optional bCloseAferFilling = True _
                            , Optional leftFieldSeparator = "{{" _
                            , Optional rightFieldSeparator = "}}" _
                            , Optional tables As ITables = Nothing _
                            ) As String

End Function


Public Property Set fields(docFields As Scripting.Dictionary)
    
End Property
Public Property Get fields() As Scripting.Dictionary

End Property

Public Property Get documentName() As String

End Property

Public Property Get tables() As Variant

End Property

Public Property Get docTableFields() As Scripting.Dictionary

End Property




