VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFieldsReaderFromXL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private bEmpty As Boolean

Private sCardPath As String
Private sWBName As String
Private sShName As String
Implements IFieldsReader
Public Sub setParams(Optional ByVal sPath As String = "", Optional ByVal sCardWBName As String = "", Optional sCardShName = "")
    sCardPath = sPath
    sWBName = sCardWBName
    sShName = sCardShName
End Sub

Public Function IFieldsReader_getFields() As Scripting.Dictionary
' читает поля документа из карточки
' sCardWBName - имя excel workbook
    Dim rFieldsRange As Range
    Dim iLastRow As Integer
    Dim irow, icol As Integer
    Dim fields As New Scripting.Dictionary
    Dim customerCardWB As Workbook
    Dim shDocumentFields As Worksheet
    
    On Error Resume Next
        
        Application.DisplayAlerts = False
        Set customerCardWB = Workbooks.Open(sCardPath + sWBName, UpdateLinks:=False, ReadOnly:=True)
        Application.DisplayAlerts = True
        If customerCardWB Is Nothing Then
            bEmpty = True
            Exit Function
        End If
        If sShName = "" Then
            Set shDocumentFields = customerCardWB.Worksheets(1)
        Else
            Set shDocumentFields = customerCardWB.Worksheets(sShName)
        End If
        
        iLastRow = shDocumentFields.UsedRange.Rows(shDocumentFields.UsedRange.Rows.Count).row
        Set rFieldsRange = shDocumentFields.UsedRange.Range(Cells(2, 1), Cells(iLastRow, 2))
        ReDim documentFields(rFieldsRange.Rows.Count)
        
        For irow = 1 To rFieldsRange.Rows.Count
            fields.Add rFieldsRange.Cells(irow, 1).value, CStr(rFieldsRange.Cells(irow, 2).value)
            
        Next irow
        Application.DisplayAlerts = False
        customerCardWB.Close
        Application.DisplayAlerts = True
    On Error GoTo 0
    
    bEmpty = False
    Set IFieldsReader_getFields = fields
    
End Function

Public Function IFieldsReader_emptyFields() As Boolean
    IFieldsReader_emptyFields = bEmpty
End Function
