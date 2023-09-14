Attribute VB_Name = "Tests"
Option Explicit

Public Sub Test_�������������()
    Dim lValue As Long
    Dim sValue As String
    Dim sResult As String
    
    lValue = 121456780012012#
    
    sValue = �������������.�������������(lValue)
    sResult = " ��� �������� ���� �������� ��������� ��������� ����� ���������� ������� ����������� ��������� ���������� ����� ����������"
    Debug.Assert sValue = sResult
    
End Sub

Public Sub Test_�������������������_DefaultCurrency()
    Dim cValue As Currency
    Dim sValue As String
    Dim sResult As String
    
    cValue = 121456780012012#
    
    sValue = �������������.�������������������(cValue)
    sResult = " ��� �������� ���� �������� ��������� ��������� ����� ���������� ������� ����������� ��������� ���������� ����� ���������� ����������� ������ 12 ������"
    Debug.Assert sValue = sResult
    
End Sub

Public Sub Test_�������������������_BYN()
    Dim cValue As Currency
    Dim sValue As String
    Dim sResult As String
    
    cValue = 121456780012012#
    
    sValue = �������������.�������������������(cValue, False)
    sResult = " ��� �������� ���� �������� ��������� ��������� ����� ���������� ������� ����������� ��������� ���������� ����� ���������� ����������� ������ 12 ������"
    Debug.Assert sValue = sResult
    
End Sub

Public Sub Test_�������������������_RUB()
    Dim cValue As Currency
    Dim sValue As String
    Dim sResult As String
    
    cValue = 121456780012012#
    
    sValue = �������������.�������������������(cValue, True)
    sResult = " ��� �������� ���� �������� ��������� ��������� ����� ���������� ������� ����������� ��������� ���������� ����� ���������� ���������� ������ 12 ������"
    Debug.Assert sValue = sResult
    
End Sub

Public Sub AllTests()
    Test_�������������
    Test_�������������������_DefaultCurrency
    Test_�������������������_BYN
    Test_�������������������_RUB
End Sub
