VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTestWork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Implements iwork

Private workName As String
Private partName As String
Private dWorkPrice As Double
Private dPartPrice As Double

Public Property Get IWork_workName() As String
    IWork_workName = workName
End Property

Public Property Let IWork_workName(sValue As String)
    workName = sValue
End Property

Public Property Get IWork_partName() As String
    IWork_partName = workName
End Property

Public Property Let IWork_partName(sValue As String)
    partName = sValue
End Property

Public Property Get IWork_workPrice() As Double
    IWork_workPrice = dWorkPrice
End Property
Public Property Let IWork_workPrice(value As Double)
    dWorkPrice = value
End Property

Public Property Get IWork_partPrice() As Double
    IWork_partPrice = dPartPrice
End Property
Public Property Let IWork_partPrice(value As Double)
    dPartPrice = value
End Property
