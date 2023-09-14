VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDefaultTables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ITables
 
Public Function ITables_getTables(doc As Document) As Table()
    
    Dim ts(1 To 2) As Table
    Set ts(1) = doc.tables(3)
    Set ts(2) = doc.tables(4)
    
    ITables_getTables = ts

End Function
