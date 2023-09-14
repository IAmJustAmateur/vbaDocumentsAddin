Attribute VB_Name = "СуммаПрописью"
#If Win64 Then
Public Function СуммаПрописью(src_num As LongLong) As String
#Else
Public Function СуммаПрописью(src_num As Long) As String
#End If
    ' число прописью, дла неотрицательных чисел, <= 999 999 999 999

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
                            wrd(i) = "тысяч"
                        Else
                            wrd(i) = ""
                        End If
                     ElseIf i = 6 Then
                        If (Digits(7) <> 0) Or (Digits(8) <> 0) Then
                            wrd(i) = "миллионов"
                        Else
                            wrd(i) = ""
                        End If
                    ElseIf i = 9 Then
                        If (Digits(10) <> 0) Or (Digits(11) <> 0) Then
                            wrd(i) = "миллиардов"
                        End If
                    ElseIf i = 12 Then
                        If (Digits(13) <> 0) Or (Digits(14) <> 0) Then
                            wrd(i) = "триллионов"
                        End If
                    Else
                         wrd(i) = ""
                    End If
                Case 1:
                    If (Digits(i + 1) <> 1) Then
                        If i = 3 Then
                            wrd(i) = "одна тысяча"
                        ElseIf i = 6 Then
                            wrd(i) = "один миллион"
                        ElseIf i = 9 Then
                            wrd(i) = "один миллиард"
                        ElseIf i = 12 Then
                            wrd(i) = "один триллион"
                        Else
                            wrd(i) = "один"
                        End If
                    Else
                        Select Case i
                            Case 0: wrd(i) = ""
                            Case 3: wrd(i) = "тысяч"
                            Case 6: wrd(i) = "миллионов"
                            Case 9: wrd(i) = "миллиардов"
                            Case 12: wrd(i) = "триллионов"
                        End Select
                    End If
                    
                Case 2:
                    If (Digits(i + 1) <> 1) Then
                        If i = 3 Then
                            wrd(i) = "две тысячи"
                        ElseIf i = 6 Then
                            wrd(i) = "два миллиона"
                        ElseIf i = 9 Then
                            wrd(i) = "два миллиарда"
                        ElseIf i = 9 Then
                            wrd(i) = "два триллиона"
                        Else
                            wrd(i) = "два"
                        End If
                    Else
                    Select Case i
                            Case 0: wrd(i) = ""
                            Case 3: wrd(i) = "тысяч"
                            Case 6: wrd(i) = "миллионов"
                            Case 9: wrd(i) = "миллиардов"
                            Case 12: wrd(i) = "триллионов"
                        End Select
                    End If
                Case 3:
                    If (Digits(i + 1) <> 1) Then
                        If i = 3 Then
                            wrd(i) = "три тысячи"
                        ElseIf i = 6 Then
                            wrd(i) = "три миллиона"
                        ElseIf i = 9 Then
                            wrd(i) = "три миллиарда"
                        ElseIf i = 12 Then
                            wrd(i) = "три триллиона"
                        Else
                            wrd(i) = "три"
                        End If
                    Else
                        Select Case i
                            Case 0: wrd(i) = ""
                            Case 3: wrd(i) = "тысяч"
                            Case 6: wrd(i) = "миллионов"
                            Case 9: wrd(i) = "миллиардов"
                            Case 12: wrd(i) = "триллионов"
                        End Select
                    End If
                Case 4:
                    If (Digits(i + 1) <> 1) Then
                        If i = 3 Then
                            wrd(i) = "четыре тысячи"
                        ElseIf i = 6 Then
                            wrd(i) = "четыре миллиона"
                        ElseIf i = 9 Then
                            wrd(i) = "четыре миллиарда"
                        ElseIf i = 9 Then
                            wrd(i) = "четыре триллиона"
                        Else
                            wrd(i) = "четыре"
                        End If
                    Else
                        Select Case i
                            Case 0: wrd(i) = ""
                            Case 3: wrd(i) = "тысяч"
                            Case 6: wrd(i) = "миллионов"
                            Case 9: wrd(i) = "миллиардов"
                            Case 12: wrd(i) = "триллионов"
                        End Select
                    End If
                Case 5:
                    If (Digits(i + 1) <> 1) Then
                        If i = 3 Then
                            wrd(i) = "пять тысяч"
                        ElseIf i = 6 Then
                            wrd(i) = "пять миллионов"
                        ElseIf i = 9 Then
                            wrd(i) = "пять миллиардов"
                        ElseIf i = 9 Then
                            wrd(i) = "пять триллионов"
                        Else
                            wrd(i) = "пять"
                        End If
                    Else
                        Select Case i
                            Case 0: wrd(i) = ""
                            Case 3: wrd(i) = "тысяч"
                            Case 6: wrd(i) = "миллионов"
                            Case 9: wrd(i) = "миллиардов"
                            Case 12: wrd(i) = "триллионов"
                        End Select
                    End If
                Case 6:
                    If (Digits(i + 1) <> 1) Then
                        If i = 3 Then
                            wrd(i) = "шесть тысяч"
                        ElseIf i = 6 Then
                            wrd(i) = "шесть миллионов"
                        ElseIf i = 9 Then
                            wrd(i) = "шесть миллиардов"
                        ElseIf i = 12 Then
                            wrd(i) = "шесть триллионов"
                        Else
                            wrd(i) = "шесть"
                        End If
                     Else
                        Select Case i
                            Case 0: wrd(i) = ""
                            Case 3: wrd(i) = "тысяч"
                            Case 6: wrd(i) = "миллионов"
                            Case 9: wrd(i) = "миллиардов"
                            Case 12: wrd(i) = "триллионов"
                        End Select
                    End If
                Case 7:
                    If (Digits(i + 1) <> 1) Then
                        If i = 3 Then
                            wrd(i) = "семь тысяч"
                        ElseIf i = 6 Then
                            wrd(i) = "семь миллионов"
                        ElseIf i = 9 Then
                            wrd(i) = "семь миллиардов"
                         ElseIf i = 12 Then
                            wrd(i) = "семь триллионов"
                        Else
                            wrd(i) = "семь"
                        End If
                    Else
                        Select Case i
                            Case 0: wrd(i) = ""
                            Case 3: wrd(i) = "тысяч"
                            Case 6: wrd(i) = "миллионов"
                            Case 9: wrd(i) = "миллиардов"
                            Case 9: wrd(i) = "триллионов"
                        End Select
                    End If
                Case 8:
                    If (Digits(i + 1) <> 1) Then
                        If i = 3 Then
                            wrd(i) = "восемь тысяч"
                        ElseIf i = 6 Then
                            wrd(i) = "восемь миллионов"
                        ElseIf i = 9 Then
                            wrd(i) = "восемь миллиардов"
                        ElseIf i = 9 Then
                            wrd(i) = "восемь триллионов"
                        Else
                            wrd(i) = "восемь"
                        End If
                     Else
                        Select Case i
                            Case 0: wrd(i) = ""
                            Case 3: wrd(i) = "тысяч"
                            Case 6: wrd(i) = "миллионов"
                            Case 9: wrd(i) = "миллиардов"
                            Case 12: wrd(i) = "триллионов"
                        End Select
                    End If
                Case 9:
                    If (Digits(i + 1) <> 1) Then
                        If i = 3 Then
                            wrd(i) = "девять тысяч"
                            ElseIf i = 6 Then
                            wrd(i) = "девять миллионов"
                        ElseIf i = 9 Then
                            wrd(i) = "девять миллиардов"
                        ElseIf i = 12 Then
                            wrd(i) = "девять триллионов"
                        Else
                            wrd(i) = "девять"
                        End If
                    Else
                        Select Case i
                            Case 0: wrd(i) = ""
                            Case 3: wrd(i) = "тысяч"
                            Case 6: wrd(i) = "миллионов"
                            Case 9: wrd(i) = "миллиардов"
                            Case 12: wrd(i) = "триллионов"
                        End Select
                    End If
            End Select
        ElseIf (i Mod 3 = 1) Then
               Select Case Digits(i)
                Case 0: wrd(i) = ""
                Case 1:
                    If Digits(i - 1) = 0 Then
                        wrd(i) = "десять"
                    Else
                        Select Case Digits(i - 1)
                            Case 1: wrd(i) = "одиннадцать"
                            Case 2: wrd(i) = "двенадцать"
                            Case 3: wrd(i) = "тринадцать"
                            Case 4: wrd(i) = "четырнадцать"
                            Case 5: wrd(i) = "пятнадцать"
                            Case 6: wrd(i) = "шестнадцать"
                            Case 7: wrd(i) = "семнадцать"
                            Case 8: wrd(i) = "восемнадцать"
                            Case 9: wrd(i) = "девятнадцать"
                        End Select
                    End If
                Case 2: wrd(i) = "двадцать"
                Case 3: wrd(i) = "тридцать"
                            
                Case 4: wrd(i) = "сорок"
                            
                Case 5: wrd(i) = "пятьдесят"
                         
                Case 6: wrd(i) = "шестьдесят"
                            
                Case 7: wrd(i) = "семьдесят"
                           
                Case 8: wrd(i) = "восемьдесят"
                Case 9: wrd(i) = "девяносто"
                End Select
         ElseIf (i Mod 3 = 2) Then
               Select Case Digits(i)
                Case 0: wrd(i) = ""
                Case 1: wrd(i) = "сто"
                Case 2: wrd(i) = "двести"
                Case 3: wrd(i) = "триста"
                            
                Case 4: wrd(i) = "четыреста"
                            
                Case 5: wrd(i) = "пятьсот"
                         
                Case 6: wrd(i) = "шестьсот"
                            
                Case 7: wrd(i) = "семьсот"
                           
                Case 8: wrd(i) = "восемьсот"
                Case 9: wrd(i) = "девятьсот"
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
    СуммаПрописью = rez
End Function
Public Function СуммаРублейПрописью(dsource As Currency, Optional bRub = False) As String
' bRub, если Истина (True), российские рубли
' По умолчанию Ложь(False), белорусские рубли
' если после запятой более 2х знаков, производится округление до копеек
' для неотрицательных чисел, <= 999'999'999'999'999
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
        sRouble = " белорусских рублей"
    Else
        Select Case (lWhole Mod 10)
            Case 1: sRouble = " белорусский рубль"
            Case 2, 3, 4: sRouble = " белорусских рубля"
            Case 5, 6, 7, 8, 9, 0: sRouble = " белорусских рублей"
        End Select
    End If
    
    If ((lFrac > 10) And (lFrac < 20)) Then
        sKop = " копеек"
    Else
        Select Case (lFrac Mod 10)
            Case 1: sKop = " копейка"
            Case 2, 3, 4: sKop = " копейки"
            Case 5, 6, 7, 8, 9, 0: sKop = " копеек"
        End Select
    End If
    If Len(CStr(lFrac)) = 1 Then
        sKop = "0" + CStr(lFrac) + sKop
    Else
        sKop = CStr(lFrac) + sKop
    End If
           
    s = СуммаПрописью(lWhole) + sRouble + " " + sKop
    If bRub Then
        s = Replace(s, "белорусски", "российски", Start:=1, Compare:=vbTextCompare)
    End If
    СуммаРублейПрописью = s

End Function
