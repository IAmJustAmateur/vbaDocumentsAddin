Attribute VB_Name = "�������������"
#If Win64 Then
Public Function �������������(src_num As LongLong) As String
#Else
Public Function �������������(src_num As Long) As String
#End If
    ' ����� ��������, ��� ��������������� �����, <= 999 999 999 999

    Dim Digits(15) As Long
    Dim wrd(15) As String
    Dim rez As String
    Dim inp_str As String
    Dim i, j, NumberOfDigits As Integer
            
    i = 0
    Do While (i <= 14) And (src_num > 0)
        Digits(i) = src_num Mod 10
        src_num = src_num \ 10
        i = i + 1
    Loop
    
    NumberOfDigits = i
    i = 0
    Do While i < NumberOfDigits
        If (i Mod 3 = 0) Then
            Select Case Digits(i)
                Case 0:
                    If i = 3 Then
                        If (Digits(4) <> 0) Or (Digits(5) <> 0) Then
                            wrd(i) = "�����"
                        Else
                            wrd(i) = ""
                        End If
                     ElseIf i = 6 Then
                        If (Digits(7) <> 0) Or (Digits(8) <> 0) Then
                            wrd(i) = "���������"
                        Else
                            wrd(i) = ""
                        End If
                    ElseIf i = 9 Then
                        If (Digits(10) <> 0) Or (Digits(11) <> 0) Then
                            wrd(i) = "����������"
                        End If
                    ElseIf i = 12 Then
                        If (Digits(13) <> 0) Or (Digits(14) <> 0) Then
                            wrd(i) = "����������"
                        End If
                    Else
                         wrd(i) = ""
                    End If
                Case 1:
                    If (Digits(i + 1) <> 1) Then
                        If i = 3 Then
                            wrd(i) = "���� ������"
                        ElseIf i = 6 Then
                            wrd(i) = "���� �������"
                        ElseIf i = 9 Then
                            wrd(i) = "���� ��������"
                        ElseIf i = 12 Then
                            wrd(i) = "���� ��������"
                        Else
                            wrd(i) = "����"
                        End If
                    Else
                        Select Case i
                            Case 0: wrd(i) = ""
                            Case 3: wrd(i) = "�����"
                            Case 6: wrd(i) = "���������"
                            Case 9: wrd(i) = "����������"
                            Case 12: wrd(i) = "����������"
                        End Select
                    End If
                    
                Case 2:
                    If (Digits(i + 1) <> 1) Then
                        If i = 3 Then
                            wrd(i) = "��� ������"
                        ElseIf i = 6 Then
                            wrd(i) = "��� ��������"
                        ElseIf i = 9 Then
                            wrd(i) = "��� ���������"
                        ElseIf i = 9 Then
                            wrd(i) = "��� ���������"
                        Else
                            wrd(i) = "���"
                        End If
                    Else
                    Select Case i
                            Case 0: wrd(i) = ""
                            Case 3: wrd(i) = "�����"
                            Case 6: wrd(i) = "���������"
                            Case 9: wrd(i) = "����������"
                            Case 12: wrd(i) = "����������"
                        End Select
                    End If
                Case 3:
                    If (Digits(i + 1) <> 1) Then
                        If i = 3 Then
                            wrd(i) = "��� ������"
                        ElseIf i = 6 Then
                            wrd(i) = "��� ��������"
                        ElseIf i = 9 Then
                            wrd(i) = "��� ���������"
                        ElseIf i = 12 Then
                            wrd(i) = "��� ���������"
                        Else
                            wrd(i) = "���"
                        End If
                    Else
                        Select Case i
                            Case 0: wrd(i) = ""
                            Case 3: wrd(i) = "�����"
                            Case 6: wrd(i) = "���������"
                            Case 9: wrd(i) = "����������"
                            Case 12: wrd(i) = "����������"
                        End Select
                    End If
                Case 4:
                    If (Digits(i + 1) <> 1) Then
                        If i = 3 Then
                            wrd(i) = "������ ������"
                        ElseIf i = 6 Then
                            wrd(i) = "������ ��������"
                        ElseIf i = 9 Then
                            wrd(i) = "������ ���������"
                        ElseIf i = 9 Then
                            wrd(i) = "������ ���������"
                        Else
                            wrd(i) = "������"
                        End If
                    Else
                        Select Case i
                            Case 0: wrd(i) = ""
                            Case 3: wrd(i) = "�����"
                            Case 6: wrd(i) = "���������"
                            Case 9: wrd(i) = "����������"
                            Case 12: wrd(i) = "����������"
                        End Select
                    End If
                Case 5:
                    If (Digits(i + 1) <> 1) Then
                        If i = 3 Then
                            wrd(i) = "���� �����"
                        ElseIf i = 6 Then
                            wrd(i) = "���� ���������"
                        ElseIf i = 9 Then
                            wrd(i) = "���� ����������"
                        ElseIf i = 9 Then
                            wrd(i) = "���� ����������"
                        Else
                            wrd(i) = "����"
                        End If
                    Else
                        Select Case i
                            Case 0: wrd(i) = ""
                            Case 3: wrd(i) = "�����"
                            Case 6: wrd(i) = "���������"
                            Case 9: wrd(i) = "����������"
                            Case 12: wrd(i) = "����������"
                        End Select
                    End If
                Case 6:
                    If (Digits(i + 1) <> 1) Then
                        If i = 3 Then
                            wrd(i) = "����� �����"
                        ElseIf i = 6 Then
                            wrd(i) = "����� ���������"
                        ElseIf i = 9 Then
                            wrd(i) = "����� ����������"
                        ElseIf i = 12 Then
                            wrd(i) = "����� ����������"
                        Else
                            wrd(i) = "�����"
                        End If
                     Else
                        Select Case i
                            Case 0: wrd(i) = ""
                            Case 3: wrd(i) = "�����"
                            Case 6: wrd(i) = "���������"
                            Case 9: wrd(i) = "����������"
                            Case 12: wrd(i) = "����������"
                        End Select
                    End If
                Case 7:
                    If (Digits(i + 1) <> 1) Then
                        If i = 3 Then
                            wrd(i) = "���� �����"
                        ElseIf i = 6 Then
                            wrd(i) = "���� ���������"
                        ElseIf i = 9 Then
                            wrd(i) = "���� ����������"
                         ElseIf i = 12 Then
                            wrd(i) = "���� ����������"
                        Else
                            wrd(i) = "����"
                        End If
                    Else
                        Select Case i
                            Case 0: wrd(i) = ""
                            Case 3: wrd(i) = "�����"
                            Case 6: wrd(i) = "���������"
                            Case 9: wrd(i) = "����������"
                            Case 9: wrd(i) = "����������"
                        End Select
                    End If
                Case 8:
                    If (Digits(i + 1) <> 1) Then
                        If i = 3 Then
                            wrd(i) = "������ �����"
                        ElseIf i = 6 Then
                            wrd(i) = "������ ���������"
                        ElseIf i = 9 Then
                            wrd(i) = "������ ����������"
                        ElseIf i = 9 Then
                            wrd(i) = "������ ����������"
                        Else
                            wrd(i) = "������"
                        End If
                     Else
                        Select Case i
                            Case 0: wrd(i) = ""
                            Case 3: wrd(i) = "�����"
                            Case 6: wrd(i) = "���������"
                            Case 9: wrd(i) = "����������"
                            Case 12: wrd(i) = "����������"
                        End Select
                    End If
                Case 9:
                    If (Digits(i + 1) <> 1) Then
                        If i = 3 Then
                            wrd(i) = "������ �����"
                            ElseIf i = 6 Then
                            wrd(i) = "������ ���������"
                        ElseIf i = 9 Then
                            wrd(i) = "������ ����������"
                        ElseIf i = 12 Then
                            wrd(i) = "������ ����������"
                        Else
                            wrd(i) = "������"
                        End If
                    Else
                        Select Case i
                            Case 0: wrd(i) = ""
                            Case 3: wrd(i) = "�����"
                            Case 6: wrd(i) = "���������"
                            Case 9: wrd(i) = "����������"
                            Case 12: wrd(i) = "����������"
                        End Select
                    End If
            End Select
        ElseIf (i Mod 3 = 1) Then
               Select Case Digits(i)
                Case 0: wrd(i) = ""
                Case 1:
                    If Digits(i - 1) = 0 Then
                        wrd(i) = "������"
                    Else
                        Select Case Digits(i - 1)
                            Case 1: wrd(i) = "�����������"
                            Case 2: wrd(i) = "����������"
                            Case 3: wrd(i) = "����������"
                            Case 4: wrd(i) = "������������"
                            Case 5: wrd(i) = "����������"
                            Case 6: wrd(i) = "�����������"
                            Case 7: wrd(i) = "����������"
                            Case 8: wrd(i) = "������������"
                            Case 9: wrd(i) = "������������"
                        End Select
                    End If
                Case 2: wrd(i) = "��������"
                Case 3: wrd(i) = "��������"
                            
                Case 4: wrd(i) = "�����"
                            
                Case 5: wrd(i) = "���������"
                         
                Case 6: wrd(i) = "����������"
                            
                Case 7: wrd(i) = "���������"
                           
                Case 8: wrd(i) = "�����������"
                Case 9: wrd(i) = "���������"
                End Select
         ElseIf (i Mod 3 = 2) Then
               Select Case Digits(i)
                Case 0: wrd(i) = ""
                Case 1: wrd(i) = "���"
                Case 2: wrd(i) = "������"
                Case 3: wrd(i) = "������"
                            
                Case 4: wrd(i) = "���������"
                            
                Case 5: wrd(i) = "�������"
                         
                Case 6: wrd(i) = "��������"
                            
                Case 7: wrd(i) = "�������"
                           
                Case 8: wrd(i) = "���������"
                Case 9: wrd(i) = "���������"
                End Select
            End If
        i = i + 1
    Loop
    rez = ""
    For i = NumberOfDigits To 1 Step -1
        If (wrd(i - 1) <> "") Then
            rez = rez & " " & wrd(i - 1)
        End If
    Next i
    ������������� = rez
End Function
Public Function �������������������(dsource As Currency, Optional bRub = False) As String
' bRub, ���� ������ (True), ���������� �����
' �� ��������� ����(False), ����������� �����
' ���� ����� ������� ����� 2� ������, ������������ ���������� �� ������
' ��� ��������������� �����, <= 999'999'999'999'999
#If Win64 Then
    Dim lWhole As LongLong
#Else
    Dim lWhole As Long
#End If
    Dim lFrac As Long
    Dim s As String
    Dim sRouble As String
    Dim sKop As String
    lWhole = Val(dsource)
    lFrac = CInt((dsource - lWhole) * 100)
    If ((lWhole Mod 100) > 10) And ((lWhole Mod 100) < 20) Then
        sRouble = " ����������� ������"
    Else
        Select Case (lWhole Mod 10)
            Case 1: sRouble = " ����������� �����"
            Case 2, 3, 4: sRouble = " ����������� �����"
            Case 5, 6, 7, 8, 9, 0: sRouble = " ����������� ������"
        End Select
    End If
    
    If ((lFrac > 10) And (lFrac < 20)) Then
        sKop = " ������"
    Else
        Select Case (lFrac Mod 10)
            Case 1: sKop = " �������"
            Case 2, 3, 4: sKop = " �������"
            Case 5, 6, 7, 8, 9, 0: sKop = " ������"
        End Select
    End If
    If Len(CStr(lFrac)) = 1 Then
        sKop = "0" + CStr(lFrac) + sKop
    Else
        sKop = CStr(lFrac) + sKop
    End If
           
    s = �������������(lWhole) + sRouble + " " + sKop
    If bRub Then
        s = Replace(s, "����������", "���������", Start:=1, Compare:=vbTextCompare)
    End If
    ������������������� = s

End Function
