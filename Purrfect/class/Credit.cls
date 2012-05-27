VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Credit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarcustomer_info As New Customers 'local copy
Private mvarlimit As Double 'local copy
Public Sub load_info(customers_id As Integer)
End Sub

Public Sub update()
Dim update As String

update = "UPDATE customers SET credit_limit = " & Me.limit & " WHERE customers_id = " & Me.customer_info.customers_id
db.execute update

End Sub

Public Property Let limit(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.limit = 5
    mvarlimit = vData
End Property


Public Property Get limit() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.limit
    limit = mvarlimit
End Property



Public Property Set customer_info(ByVal vData As Customers)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.customer_info = Form1
    Set mvarcustomer_info = vData
End Property


Public Property Get customer_info() As Customers
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.customer_info
    Set customer_info = mvarcustomer_info
End Property



