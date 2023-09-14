VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTestLoaders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements Iloaders

Private Sub Class_Initialize()
    Iloaders_loadCustomers
    Iloaders_loadWorks
End Sub

Public Function Iloaders_loadCustomers() As Collection
    Dim C As New Collection
    Dim customer As CTestCustomer
    
    Dim i As Integer
    Dim sCustomer As String
    
    For i = 1 To 10
        sCustomer = "Заказчик " + CStr(i)
        Set customer = New CTestCustomer
        customer.customerName = sCustomer
        C.Add customer
    Next i
    
    Set Iloaders_loadCustomers = C
    
End Function

Public Function Iloaders_loadWorks() As Collection
    Dim C As New Collection
    Dim work As CTestWork
    Dim i As Integer
    Dim sWork As String
    
    For i = 1 To 10
        sWork = "Работа " + CStr(i)
        Set work = New CTestWork
        work.IWork_workName = sWork
        
        work.IWork_partPrice = i + 0.25
        work.IWork_workPrice = i - 0.5
        C.Add work
    Next i
    Set Iloaders_loadWorks = C

End Function

Public Function Iloaders_getNewNumber() As Integer
    Randomize
    Iloaders_getNewNumber = Round(Rnd * 1000, 0)
End Function

