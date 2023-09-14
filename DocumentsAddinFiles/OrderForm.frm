VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} OrderForm 
   Caption         =   "Заказ-наряд на техническое обслуживание автоцистены"
   ClientHeight    =   6195
   ClientLeft      =   10020
   ClientTop       =   465
   ClientWidth     =   13605
   OleObjectBlob   =   "OrderForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "OrderForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public tbDatePartCollection As Collection ' contains CDateNumberBox objects
Public sbDateCollection As Collection ' contains CDateSpinButton objects

Public floatControlsCollection As Collection ' contains objects with floating borders that change location when the order form changes
Public RowFrames As Collection ' contains frames/order rows

Private customers As Collection ' customer names for cmbCustomers
Private workList As Collection

' Document params

Private sTemplateName As String  ' template filename
Private docTablesSelector As ITables ' tables from template to fill with values
Private docNameGenerator As INameGenerator   ' order file name generator
Private docFields As Scripting.Dictionary

Private sDocTemplatePath As String  ' path to order template
Private sDocPath As String          ' path to the folder with finished documents
Private bCloseDocument As Boolean   ' close document or not
Private lFieldSeparator As String
Private rFieldSeparator As String

Public Sub load_initial_values(loaders As Iloaders)
   
    Dim obCustomer As ICustomer
    Dim obWork As iwork
    
    Set customers = loaders.loadCustomers()
    Set workList = loaders.loadWorks()
    
    For Each obCustomer In customers
        cmbCustomers.AddItem obCustomer.name
    Next obCustomer
    
    For Each obWork In works
        cmbWorks1.AddItem obWork.workName
    Next obWork
    
    Me.lblOrderNumber.Caption = loaders.getNewNumber()

End Sub
Public Sub setDocumentParams(DocGenerator As CDefaultDocument _
                        , sDocumentTemplateName As String _
                        , tablesSelector As ITables _
                        , nameGenerator As INameGenerator _
                        , fields As Dictionary _
                        , Optional sTemplatePath = "" _
                        , Optional sDocumentPath = "" _
                        , Optional bCloseAferFilling = True _
                        , Optional leftFieldSeparator = "{{" _
                        , Optional rightFieldSeparator = "}}")
                        
    
    sTemplateName = sDocumentTemplateName
    Set docTablesSelector = tablesSelector
    Set docNameGenerator = nameGenerator
    Set docFields = fields
    
    sDocTemplatePath = sTemplatePath
    sDocPath = sDocumentPath
    bCloseDocument = bCloseAferFilling
    lFieldSeparator = leftFieldSeparator
    rFieldSeparator = rightFieldSeparator

End Sub

Public Property Get works() As Collection
    Set works = workList
End Property

Public Function getWork(sWork As String) As iwork
    Dim work As iwork
    If works Is Nothing Then
        Set getWork = Nothing
        Exit Function
    End If
    
    For Each work In works
        If LCase(work.workName) = LCase(sWork) Then
            Set getWork = work
            Exit Function
        End If
    Next work
End Function

Public Sub setTotalAmount()
' calculate order Total amount
    
    Dim damount As Double
    Dim i As Integer
    damount = 0
    For i = 1 To Me.RowFrames.Count
        On Error Resume Next
        damount = damount + CDbl(Me.RowFrames(i).lblTotal)
        On Error GoTo 0
    Next i
    
    Me.lblTotalAmountValue.Caption = Format(damount, "#,##0.00")
    
End Sub

Private Sub btnAddWorkFrame_Click()
' add new order row

    Dim rowFrame As CRowFrame
    Dim ctl As control
    Set rowFrame = createRowFrame(Me)
    
    For Each ctl In floatControlsCollection
        moveUp ctl, -(Me.RowFrames(Me.RowFrames.Count).fr.Height + 10)
    Next ctl
    
End Sub

Private Sub cbOk_Click()
    Dim order As COrder
    Set order = createOrder
    MsgBox "Готово:"
End Sub

Private Sub cbQuit_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    
    Dim dateBox As CDateNumberBox
    Dim dsb As CDateSpinButton
    Dim ctrl As control
    Dim RowFrame1 As CRowFrame
             
    Set floatControlsCollection = New Collection
    Set RowFrames = New Collection
    
    Set RowFrame1 = createRowFrame(Me)
    
    floatControlsCollection.Add Me.cbOk
    floatControlsCollection.Add Me.cbQuit
    floatControlsCollection.Add Me.lblTotalAmount
    floatControlsCollection.Add Me.lblTotalAmountValue
    
    ' date
    Set tbDatePartCollection = New Collection
    Set sbDateCollection = New Collection
    
    ' day
    Set dateBox = New CDateNumberBox
    Set dateBox.control = Me.tbDay
    Set dateBox.form = Me
    dateBox.set_borders 1, 31
    Me.tbDay.Text = CStr(Day(Now))
    tbDatePartCollection.Add dateBox
    
    Set dsb = New CDateSpinButton
    Set dsb.control = Me.sbDay
    Set dsb.NumberBox = dateBox
    Set dsb.form = Me
    tbDatePartCollection.Add dsb
    
    ' month
    Set dateBox = New CDateNumberBox
    Set dateBox.control = Me.tbMonth
    dateBox.set_borders 1, 12
    Set dateBox.form = Me
    Me.tbMonth.Text = CStr(Month(Now))
    tbDatePartCollection.Add dateBox
    
    Set dsb = New CDateSpinButton
    Set dsb.control = Me.sbMonth
    Set dsb.NumberBox = dateBox
    Set dsb.form = Me
    tbDatePartCollection.Add dsb
    
    ' year
    Set dateBox = New CDateNumberBox
    Set dateBox.control = Me.tbYear
    Set dateBox.form = Me
    dateBox.set_borders 2015, 2030
    Me.tbYear.Text = CStr(Year(Now))
    tbDatePartCollection.Add dateBox
    
    Set dsb = New CDateSpinButton
    Set dsb.control = Me.sbYear
    Set dsb.NumberBox = dateBox
    Set dsb.form = Me
    tbDatePartCollection.Add dsb
    
    sbDay.value = Day(Now)
    sbMonth.value = Month(Now)
    sbYear.value = Year(Now)

End Sub

Public Sub moveUp(ctrl As control, value As Integer)
    ctrl.Top = ctrl.Top - value
End Sub

Public Sub enumerateFrames()
    Dim inumber As Integer
    
    For inumber = 2 To Me.RowFrames.Count
        Me.RowFrames(inumber).lblWorkNumber.Caption = CStr(inumber)
    Next inumber
    
End Sub

Public Sub redraw()
    Dim inumber As Integer
    Dim ctl As control
    For inumber = 2 To Me.RowFrames.Count
         Me.RowFrames(inumber).fr.Top = Me.RowFrames(inumber - 1).fr.Top + Me.RowFrames(inumber - 1).fr.Height + 10
    Next inumber
    For Each ctl In Me.floatControlsCollection
        moveUp ctl, Me.RowFrame1.Height + 10
    Next ctl
End Sub

Public Function createOrder() As COrder
    Dim order As New COrder
    Dim fieldsReader As New COrderFieldsReader

    Set fieldsReader.oForm = Me
    Set order.fieldsReader = fieldsReader
      
    order.setParams sTemplateName, sDocTemplatePath, sDocPath, bCloseDocument, lFieldSeparator, rFieldSeparator, docTablesSelector
    order.createDocument

End Function

Public Property Get worksQty() As Integer
    worksQty = Me.RowFrames.Count
End Property

Public Property Get sOrderDate() As String
    sOrderDate = tbDay.Text + "." + tbMonth.Text + "." + tbYear.Text
End Property
