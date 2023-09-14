Attribute VB_Name = "Factories"
Option Explicit
' rowFrame factory

Public Function createRowFrame(form As OrderForm) As CRowFrame
    Dim crf As CRowFrame
    Dim sFrName As String
    
    Set crf = New CRowFrame
    If form.RowFrames.Count = 0 Then
        sFrName = "1"
    Else
        sFrName = crf.sUnique
    End If
    form.RowFrames.Add crf, sFrName
    crf.create form
    Set createRowFrame = crf
End Function





