VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Payment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvaramount As Double 'local copy
Private mvarremarks As Variant 'local copy
Private mvarcustomer As Customers 'local copy
Private mvarsales_order_no As String 'local copy
Private mvarpayment_date As String 'local copy
Private mvarbalance As Double 'local copy
Private mvarpayment_id As Integer 'local copy
'local variable(s) to hold property value(s)
Private mvaractive_sales As Sales 'local copy
'local variable(s) to hold property value(s)
Private mvarreceived_by As String 'local copy
Public Property Let received_by(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.received_by = 5
    mvarreceived_by = vData
End Property


Public Property Get received_by() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.received_by
    received_by = mvarreceived_by
End Property



Public Sub isFullyPaid()
End Sub

Public Property Set active_sales(ByVal vData As Sales)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.active_sales = Form1
    Set mvaractive_sales = vData
End Property


Public Property Get active_sales() As Sales
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.active_sales
    Set active_sales = mvaractive_sales
End Property
Public Function countPaymentOfThisSale() As Integer
    Dim rs As New ADODB.Recordset
    Dim sql As String
    sql = "SELECT COUNT(*) FROM payment_records WHERE sales_order_no = '" & Me.sales_order_no & "'"
    Set rs = db.execute(sql)
    If rs.RecordCount > 0 Then
        countPaymentOfThisSale = rs.Fields(0).Value
    Else
        countPaymentOfThisSale = 0
    End If
End Function


Public Sub printPaymentInfoAndNewBalance()
    Dim cntl As Object
    Dim rs As New ADODB.Recordset
    Dim remarks As String
    Dim ctr As Integer
    
    Set rs = db.execute("SELECT * from payment_records")
    
    Set dtaSalesOrder.DataSource = rs
    For Each cntl In dtaSalesOrder.Sections(1).Controls
      cntl.Visible = False
      'Debug.Print cntl.Visible
    Next
    
    For ctr = 2 To Me.countPaymentOfThisSale
        remarks = remarks & " " & vbCrLf & vbCrLf
    Next ctr
    remarks = remarks & "AMT PAID: P" & FormatNumber(Me.amount, 2) & " BAL: P" & (Me.getActualBalance) & vbCrLf & FormatDateTime(Date, vbShortDate)
    dtaSalesOrder.Sections(1).Controls("lblRemarks").Caption = remarks
    dtaSalesOrder.Sections(1).Controls("lblRemarks").Visible = True
    dtaSalesOrder.Show 1
End Sub
Public Sub loadRemarksToListview(lsv As ListView)
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim list As ListItem
    
    sql = "SELECT * from payment_records WHERE sales_order_no = '" & Me.sales_order_no & "'"
    Set rs = db.execute(sql)
    If rs.RecordCount > 0 Then
        lsv.ListItems.Clear
        ctr = 1
        Do Until rs.EOF
            Set list = lsv.ListItems.Add(, , ctr)
            list.SubItems(1) = FormatNumber(rs.Fields("amount").Value, 2)
            list.SubItems(2) = FormatNumber(rs.Fields("balance").Value, 2)
            list.SubItems(3) = rs.Fields("payment_date").Value
        ctr = ctr + 1
        rs.MoveNext
        Loop
    End If
End Sub

Public Function getActualBalance() As Double
    getActualBalance = Me.active_sales.info_grand_total - Me.getTotalPaidOfThisSalesOrder
End Function
Function getTotalPaidOfThisSalesOrder() As Double
    Dim rs As New ADODB.Recordset
    Dim sql As String
    
    sql = "SELECT SUM(amount) as total_paid FROM `payment_records` " & _
            "WHERE sales_order_no = '" & Me.sales_order_no & "'" & _
            "GROUP BY sales_order_no"
    Set rs = db.execute(sql)
    If rs.RecordCount > 0 Then
        getTotalPaidOfThisSalesOrder = rs.Fields(0).Value
    End If
End Function
Public Sub savePayment()
    Dim insert As String
    'id, sales_order_no, amount, balance, payment_date, remarks
    insert = "INSERT INTO payment_records VALUES(null,'" & Me.sales_order_no & "'," & Me.amount & "," & Me.balance & ",CURDATE(),'" & Me.remarks & "','" & Me.received_by & "')"
    db.execute insert
End Sub

Public Sub loadPayment(id As Integer)
    Dim rs As New ADODB.Recordset
    Dim sql As String
    'id, sales_order_no, amount, balance, payment_date, remarks
    sql = "SELECT * from payment_records WHERE id = " & id
    Set rs = db.execute(sql)
    If rs.RecordCount > 0 Then
        Me.sales_order_no = rs.Fields("sales_order_no").Value
        Me.payment_id = rs.Fields("id").Value
        Me.amount = rs.Fields("amount").Value
        Me.balance = rs.Fields("balance").Value
        Me.remarks = rs.Fields("remarks").Value
        Me.received_by = rs.Fields("received_by").Value
        Me.active_sales.loadSalesOrder (rs.Fields("sales_order_no").Value)
    End If
End Sub

Public Property Let payment_id(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.payment_id = 5
    mvarpayment_id = vData
End Property


Public Property Get payment_id() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.payment_id
     payment_id = mvarpayment_id
End Property



Public Property Let balance(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.balance = 5
    mvarbalance = vData
End Property


Public Property Get balance() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.balance
     balance = mvarbalance
End Property



Public Property Let payment_date(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.payment_date = 5
    mvarpayment_date = vData
End Property


Public Property Get payment_date() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.payment_date
     payment_date = mvarpayment_date
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



Public Property Set customer(ByVal vData As Customers)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.customer = Form1
    Set mvarcustomer = vData
End Property


Public Property Get customer() As Customers
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.customer
    Set customer = mvarcustomer
End Property



Public Sub getRemarksInfo()
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



Public Property Let amount(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.amount = 5
    mvaramount = vData
End Property


Public Property Get amount() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.amount
    amount = mvaramount
End Property



