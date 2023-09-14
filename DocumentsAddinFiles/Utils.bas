Attribute VB_Name = "Utils"
Option Explicit

Public Function isEmptyArray(arr As Variant) As Boolean
' return True if arr is Empty
    Dim bInitialized As Boolean
    On Error Resume Next
        bInitialized = IsNumeric(UBound(arr))
        isEmptyArray = Not bInitialized
    On Error GoTo 0
End Function

Public Sub replaceFieldWithValue(doc As Document, sFieldName As String, sFieldValue As String, leftFieldSeparator As String, rightFieldSeparator As String)
    With doc.Content.Find
        .Text = leftFieldSeparator + sFieldName + rightFieldSeparator
        .Forward = True
        .Replacement.Text = sFieldValue
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With
End Sub

Public Function duplicate2ndRow(t As Table)
    t.Rows(2).Range.Copy
    t.Rows(2).Range.Paste
End Function

Public Sub replaceInTableBody(t As Table, oldStr As String, newStrs As Variant)
' replace in table t all occurencies of oldStr with newStrs values
' we consider 1st table row as table headers,
' t.rows.count <= UBound(newStrs)

    Dim irow As Integer
    Dim icell As Integer
    Dim row As row
    Dim C As cell
    Dim r As Object
    Dim sCellText As String
    For irow = 2 To t.Rows.Count
        Set row = t.Rows(irow)
            For icell = 1 To row.Cells.Count
                Set C = row.Cells(icell)
                sCellText = C.Range.Text
                If InStr(sCellText, oldStr) > 0 Then
                    Set r = C.Range
                    sCellText = Replace(sCellText, oldStr, newStrs(irow - 1))
                    C.Range.Text = sCellText
                    r.Text = Replace(r.Text, vbCr, "") ' , Len(r.Text) - 3  or Len(r.Text) - 2, 1, vbTextCompare: when i try to replace only last occurence, it does not work correctly
                    
                End If
            Next icell
    Next irow

End Sub

Public Function isFieldInTable(sFieldName As String, t As Table) As Boolean
' return True if str sFieldName is in table t
    Dim irow As Integer
    Dim icolumn As Integer
    
    Dim C As cell
    
    For irow = 2 To t.Rows.Count
        For icolumn = 1 To t.Columns.Count
            Set C = t.cell(irow, icolumn)
            If InStr(C.Range.Text, sFieldName) > 0 Then
                isFieldInTable = True
                Exit Function
            End If
        Next icolumn
    Next irow
    isFieldInTable = False
    
End Function

Public Function isStrInDoc(s As String, doc As Document) As Boolean
' return True is str is in document
    Dim rDoc As Range
    Dim rfound As Range
    
    'Set rDoc = doc.Content
    'Set rfound = doc.Content.Find(s)
    'isStrInDoc = Not rfound Is Nothing
    With doc.Content.Find
        .Text = s
        .Execute
        If .Found Then
            isStrInDoc = True
        Else
            isStrInDoc = False
        End If
        
    End With
    
End Function

Public Function getWordApplication() As Word.Application
' return Word.Application, launch a new one if necessary
    Dim oWord As Word.Application
    
    On Error Resume Next
        Set oWord = GetObject(, "Word.Application")
        If oWord Is Nothing Then
            Set oWord = New Word.Application
        End If
    On Error GoTo 0
    
    Set getWordApplication = oWord

End Function
