VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTestCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ICustomer
Public customerName As String

Public Property Get ICustomer_name() As String
    ICustomer_name = customerName
End Property
Public Property Let ICustomer_name(sValue As String)
    customerName = sValue
End Property
