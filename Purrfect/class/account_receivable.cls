VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "account_receivable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarsales_order_no As String 'local copy
Private mvarremarks As String 'local copy
Public Sub load_account_receivable(sales_order_no As String)
    Dim rs As New ADODB.Recordset
    Dim sql As String
    sql = "SELECT * from account_receivable WHERE sales_order_no = '" & sales_order_no & "' LIMIT 1"
    Set rs = db.execute(sql)
    If rs.RecordCount > 0 Then
        Me.sales_order_no = rs.Fields("sales_order_no").Value
        Me.remarks = rs.Fields("remarks").Value
    End If
End Sub

Public Property Let remarks(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.remarks = 5
    mvarremarks = vData
End Property


Public Property Get remarks() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.remarks
    remarks = mvarremarks
End Property



Public Property Let sales_order_no(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sales_order_no = 5
    mvarsales_order_no = vData
End Property


Public Property Get sales_order_no() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sales_order_no
    sales_order_no = mvarsales_order_no
End Property



