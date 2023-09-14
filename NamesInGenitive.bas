Attribute VB_Name = "NamesInGenitive"
Option Explicit

Public Sub NameInGenitive(sFullName As String, ByRef sSurName As String, ByRef sName As String, ByRef sMidName As String)
' ��� � ����������� ������
' sFullName - ������� ��� ��������

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
    
    ' ����������� ����
    If sName <> "" Then
        ' ����������� ���� �� �����
        sLast1 = Mid(sName, Len(sName), 1)
        Select Case LCase(sLast1)
            Case "�", "�": bWoman = True
            Case Else:
                If (StrComp(LCase(sName), "������", vbTextCompare) = 0) Or (StrComp(LCase(sName), "������", vbTextCompare) = 0) Then
                    bWoman = True
                Else
                    bWoman = False
                End If
        End Select
    Else
        ' ����������� ���� �� �������
        sLast1 = Mid(sSurName, Len(sSurName), 1)
        Select Case (sLast1)
            Case "�", "�": bWoman = True
            Case Else: bWoman = False
        End Select
    End If
    
    ' ��������� �������
    sLast2 = LCase(Mid(sSurName, Len(sSurName) - 1, 2))
    sLast1 = LCase(Mid(sSurName, Len(sSurName), 1))
    If (StrComp(sLast2, "��", vbTextCompare) = 0) Or (StrComp(sLast2, "��", vbTextCompare) = 0) Then
        sSurName = Mid(sSurName, 1, Len(sSurName) - 2)
        sSurName = sSurName + "���"
    ElseIf StrComp(sLast2, "��", vbTextCompare) = 0 Then
        sSurName = Mid(sSurName, 1, Len(sSurName) - 2)
        sSurName = sSurName + "��"
    Else
        Select Case sLast1
            Case "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�":
                If Not bWoman Then
                    sSurName = sSurName + "�"
                End If
            Case "�", "�":
                sSurName = Mid(sSurName, 1, Len(sSurName) - 1)
                sSurName = sSurName + "�"
            Case "�":
                sSurName = Mid(sSurName, 1, Len(sSurName) - 1)
                sSurName = sSurName + "��"
        End Select
    End If
    
    ' ��������� �����
    If sName = "" Then
        Exit Sub
    End If
    sLast1 = LCase(Mid(sName, Len(sName), 1))
    Select Case sLast1
        Case "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�":
                sName = sName + "�"
        Case "�":
                sName = Mid(sName, 1, Len(sName) - 1)
                sName = sName + "�"
        Case "�":
            If StrComp(sName, "������", vbTextCompare) = 0 Then
                sName = "������"
            Else
                sName = Mid(sName, 1, Len(sName) - 1)
                sName = sName + "�"
            End If
        Case "�":
                sName = Mid(sName, 1, Len(sName) - 1)
                sName = sName + "�"
        Case "�":
                sName = Mid(sName, 1, Len(sName) - 1)
                sName = sName + "�"
    End Select
    
    ' ��������� ��������
    If sMidName <> "" Then
        sLast2 = LCase(Mid(sMidName, Len(sMidName) - 1, 2))
    Else
        sLast2 = ""
    End If
    Select Case sLast2
        Case "��":
                sMidName = sMidName + "�"
        Case "��":
                sMidName = Mid(sMidName, 1, Len(sMidName) - 1)
                sMidName = sMidName + "�"
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
