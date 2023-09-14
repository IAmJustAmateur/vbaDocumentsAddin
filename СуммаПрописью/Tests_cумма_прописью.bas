Attribute VB_Name = "Tests"
Option Explicit

Public Sub Test_—уммаѕрописью()
    Dim lValue As Long
    Dim sValue As String
    Dim sResult As String
    
    lValue = 121456780012012#
    
    sValue = —уммаѕрописью.—уммаѕрописью(lValue)
    sResult = " сто двадцать один триллион четыреста п€тьдес€т шесть миллиардов семьсот восемьдес€т миллионов двенадцать тыс€ч двенадцать"
    Debug.Assert sValue = sResult
    
End Sub

Public Sub Test_—умма–ублейѕрописью_DefaultCurrency()
    Dim cValue As Currency
    Dim sValue As String
    Dim sResult As String
    
    cValue = 121456780012012#
    
    sValue = —уммаѕрописью.—умма–ублейѕрописью(cValue)
    sResult = " сто двадцать один триллион четыреста п€тьдес€т шесть миллиардов семьсот восемьдес€т миллионов двенадцать тыс€ч двенадцать белорусских рублей 12 копеек"
    Debug.Assert sValue = sResult
    
End Sub

Public Sub Test_—умма–ублейѕрописью_BYN()
    Dim cValue As Currency
    Dim sValue As String
    Dim sResult As String
    
    cValue = 121456780012012#
    
    sValue = —уммаѕрописью.—умма–ублейѕрописью(cValue, False)
    sResult = " сто двадцать один триллион четыреста п€тьдес€т шесть миллиардов семьсот восемьдес€т миллионов двенадцать тыс€ч двенадцать белорусских рублей 12 копеек"
    Debug.Assert sValue = sResult
    
End Sub

Public Sub Test_—умма–ублейѕрописью_RUB()
    Dim cValue As Currency
    Dim sValue As String
    Dim sResult As String
    
    cValue = 121456780012012#
    
    sValue = —уммаѕрописью.—умма–ублейѕрописью(cValue, True)
    sResult = " сто двадцать один триллион четыреста п€тьдес€т шесть миллиардов семьсот восемьдес€т миллионов двенадцать тыс€ч двенадцать российских рублей 12 копеек"
    Debug.Assert sValue = sResult
    
End Sub

Public Sub AllTests()
    Test_—уммаѕрописью
    Test_—умма–ублейѕрописью_DefaultCurrency
    Test_—умма–ублейѕрописью_BYN
    Test_—умма–ублейѕрописью_RUB
End Sub
