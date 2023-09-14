VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDefaultNameGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text

Implements INameGenerator

Public Function INameGenerator_sNameGenerator(fields As Scripting.Dictionary) As String
' looking for 1st field name containing word "номер" and use it as document number and document name
    Dim sfield As String
    Dim i As Integer
    
    For i = 0 To UBound(fields.Keys)
        sfield = fields.Keys(i)
        If Not IsArray(fields(sfield)) Then
            If sfield Like "*номер*" Then
                INameGenerator_sNameGenerator = fields.Item(sfield) + ".docx"
                Exit Function
            End If
        End If
    Next i
End Function

