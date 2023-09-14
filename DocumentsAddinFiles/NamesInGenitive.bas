Attribute VB_Name = "NamesInGenitive"
Option Explicit

Public Sub NameInGenitive(sFullName As String, ByRef sSurName As String, ByRef sName As String, ByRef sMidName As String)
' ‘»ќ в родительном падеже
' sFullName - фамили€ им€ отчество

    Dim sNames() As String
    Dim bWoman As Boolean
    Dim sLast1 As String
    Dim sLast2 As String
        
    sNames = Split(sFullName)
    sSurName = sNames(0)
    If UBound(sNames) > 1 Then
        sName = sNames(1)
        If UBound(sNames) > 1 Then
            sMidName = sNames(2)
        Else
            sMidName = ""
        End If
    Else
        sName = ""
    End If
    
    ' определение пола
    If sName <> "" Then
        ' определение пола по имени
        sLast1 = Mid(sName, Len(sName), 1)
        Select Case LCase(sLast1)
            Case "а", "€": bWoman = True
            Case Else:
                If (StrComp(LCase(sName), "любовь", vbTextCompare) = 0) Or (StrComp(LCase(sName), "нинель", vbTextCompare) = 0) Then
                    bWoman = True
                Else
                    bWoman = False
                End If
        End Select
    Else
        ' определение пола по фамилии
        sLast1 = Mid(sSurName, Len(sSurName), 1)
        Select Case (sLast1)
            Case "а", "€": bWoman = True
            Case Else: bWoman = False
        End Select
    End If
    
    ' ќбработка фамилии
    sLast2 = LCase(Mid(sSurName, Len(sSurName) - 1, 2))
    sLast1 = LCase(Mid(sSurName, Len(sSurName), 1))
    If (StrComp(sLast2, "ий", vbTextCompare) = 0) Or (StrComp(sLast2, "ый", vbTextCompare) = 0) Then
        sSurName = Mid(sSurName, 1, Len(sSurName) - 2)
        sSurName = sSurName + "ого"
    ElseIf StrComp(sLast2, "а€", vbTextCompare) = 0 Then
        sSurName = Mid(sSurName, 1, Len(sSurName) - 2)
        sSurName = sSurName + "ой"
    Else
        Select Case sLast1
            Case "б", "в", "г", "д", "ж", "з", "к", "л", "м", "н", "п", "р", "с", "т", "ф", "х", "ц", "ч", "ш", "щ":
                If Not bWoman Then
                    sSurName = sSurName + "а"
                End If
            Case "й", "ь":
                sSurName = Mid(sSurName, 1, Len(sSurName) - 1)
                sSurName = sSurName + "€"
            Case "а":
                sSurName = Mid(sSurName, 1, Len(sSurName) - 1)
                sSurName = sSurName + "ой"
        End Select
    End If
    
    ' ќбработка имени
    If sName = "" Then
        Exit Sub
    End If
    sLast1 = LCase(Mid(sName, Len(sName), 1))
    Select Case sLast1
        Case "б", "в", "г", "д", "ж", "з", "к", "л", "м", "н", "п", "р", "с", "т", "ф", "х", "ц", "ч", "ш", "щ":
                sName = sName + "а"
        Case "й":
                sName = Mid(sName, 1, Len(sName) - 1)
                sName = sName + "€"
        Case "ь":
            If StrComp(sName, "Ћюбовь", vbTextCompare) = 0 Then
                sName = "Ћюбови"
            Else
                sName = Mid(sName, 1, Len(sName) - 1)
                sName = sName + "€"
            End If
        Case "а":
                sName = Mid(sName, 1, Len(sName) - 1)
                sName = sName + "ы"
        Case "€":
                sName = Mid(sName, 1, Len(sName) - 1)
                sName = sName + "и"
    End Select
    
    ' ќбработка отчества
    If sMidName <> "" Then
        sLast2 = LCase(Mid(sMidName, Len(sMidName) - 1, 2))
    Else
        sLast2 = ""
    End If
    Select Case sLast2
        Case "ич":
                sMidName = sMidName + "а"
        Case "на":
                sMidName = Mid(sMidName, 1, Len(sMidName) - 1)
                sMidName = sMidName + "ы"
    End Select
         
End Sub
Public Function FullNameInGenitive(sFullName As String) As String
    Dim s1 As String, s2 As String, s3 As String
    Dim sTmp As String
    Dim sSurName As String
    Dim sSurNameInGenitive As String
    Dim sNameRest As String
    If InStr(sFullName, ".") = 0 Then
        NameInGenitive sFullName, s1, s2, s3
        sTmp = s1
        If s2 <> "" Then
            sTmp = sTmp + " " + s2
        End If
        If s3 <> "" Then
            sTmp = sTmp + " " + s3
        End If
        FullNameInGenitive = sTmp
    Else
        sSurName = Split(sFullName)(0)
        sNameRest = Mid(sFullName, Len(sSurName) + 1)
        NameInGenitive sSurName, s1, s2, s3
        FullNameInGenitive = s1 + sNameRest
    End If
End Function

Public Function FullNameToShort(sFullName As String) As String
    Dim sArr() As String
    Dim sShortName As String
    If sFullName <> "" Then
       sArr = Split(sFullName, , 3, vbBinaryCompare)
       sShortName = sArr(0)
       If UBound(sArr) >= 1 Then
                sShortName = sShortName + " " + Mid(sArr(1), 1, 1) + "."
        End If
        If UBound(sArr) >= 2 Then
                sShortName = sShortName + " " + Mid(sArr(2), 1, 1) + "."
         End If
    End If
    FullNameToShort = sShortName
End Function
