VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Sales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvartransaction_id As String 'local copy
Private mvarsold_to As New Customers 'local copy
Private mvarpayment_type As Integer 'local copy
Private mvartransaction_type As String 'local copy
Private mvardate_transact As String 'local copy

Private mvarhasDiscount As Boolean 'local copy
'local variable(s) to hold property value(s)
Private mvaritems_sold As New cart 'local copy
'local variable(s) to hold property value(s)
Private mvartendered_amount As Double 'local copy
'local variable(s) to hold property value(s)
Private mvarchange As Double 'local copy
'local variable(s) to hold property value(s)
Private mvarisSoldToGuest As Boolean 'local copy
Public acr As New account_receivable

Public info_grand_total As Double
Public info_net_total As Double
Public info_discount As Double
Public hasRecord As Boolean
'local variable(s) to hold property value(s)
Private mvarprepared_by As String 'local copy
Public Property Let prepared_by(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.prepared_by = 5
    mvarprepared_by = vData
End Property


Public Property Get prepared_by() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.prepared_by
    prepared_by = mvarprepared_by
End Property




Public Sub delete()
    Dim delete As String
    delete = "DELETE FROM stock_out_transaction WHERE sales_order_no = '" & Me.transaction_id & "'"
    db.execute delete
End Sub


Public Property Let isSoldToWalkIn(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.isSoldToGuest = 5
    mvarisSoldToGuest = vData
End Property


Public Property Get isSoldToWalkIn() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.isSoldToGuest
    isSoldToWalkIn = mvarisSoldToGuest
End Property



Public Property Let change(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.change = 5
    mvarchange = vData
End Property


Public Property Get change() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.change
    change = mvarchange
End Property



Public Property Let tendered_amount(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.tendered_amount = 5
    mvartendered_amount = vData
End Property


Public Property Get tendered_amount() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.tendered_amount
    tendered_amount = mvartendered_amount
End Property
Sub addThisItemToCurrentSalesOrder(stockout_id)
    Dim insert As String
    insert = "INSERT INTO stock_out_transaction_stock_out_items VALUES ('" & Me.transaction_id & "'," & stockout_id & ")"
    db.execute insert
End Sub

Sub printDeliveryReceipt()
    Dim cart As New cart
    Dim items As New cart_items
    Dim sales_order_responsible As New SalesOrder_Responsible
    
    Dim rs As New ADODB.Recordset
    Set rs = db.execute("SELECT * FROM `stock_out_transaction`")
    Set dtaSalesOrder.DataSource = rs
    Dim qty As String
    Dim unit As String
    Dim desc As String
    Dim unitPrice As String
    Dim amount As String
    
    Set cart = Me.items_sold
    'QTY, UNIT, DESC, UNITPRICE, AMOUNT

    For Each items In cart
        qty = qty & items.qty_purchased & vbCrLf
        unit = unit & items.Item.unit_of_measure & vbCrLf
        desc = desc & items.Item.item_description & vbCrLf
        If Me.sold_to.dealers_type = DEALER Then
            unitPrice = unitPrice & FormatNumber((items.Item.dealers_price) + items.tracking_price, 2) & vbCrLf
            amount = amount & FormatNumber(((items.Item.dealers_price) + items.tracking_price) * items.qty_purchased, 2) & vbCrLf
        Else
            unitPrice = unitPrice & FormatNumber((items.Item.item_price) + items.tracking_price, 2) & vbCrLf
            amount = amount & FormatNumber(((items.Item.item_price) + items.tracking_price) * items.qty_purchased, 2) & vbCrLf
        End If
        
    Next
    
    With dtaSalesOrder.Sections(1)
            If activeSales.payment_type = PAYMENT_COD Then
               .Controls("lblType").Caption = "COD"
               .Controls("lblStatus").Visible = True
               .Controls("paidBorder").Visible = True
               .Controls("lblDatePaid").Visible = True
               .Controls("lblDatePaid").Caption = FormatDateTime(Date, vbShortDate)
            Else
               '.Controls("lblType").Caption = "ACCOUNT RECEIVABLE"
               
               'modified by aris, march 10,2012
               .Controls("lblType").Caption = "DELIVERY RECEIPT"
               .Controls("lblStatus").Visible = False
               .Controls("paidBorder").Visible = False
               .Controls("lblDatePaid").Visible = False
            End If
            .Controls("lblSalesOrderNo").Caption = Me.transaction_id
            .Controls("lblDate").Caption = FormatDateTime(Date, vbLongDate)
            .Controls("lblAgentName").Caption = Me.sold_to.mvaragent.agent_name & "   Contact No:(" & Me.sold_to.mvaragent.agent_contact_number & ")"
            .Controls("lblCustomerName").Caption = Me.sold_to.customers_name
            .Controls("lblAddress").Caption = Me.sold_to.customers_add
            .Controls("lblTotalItems").Caption = "( " & Me.items_sold.Count & " ) Total item/s"
            
            .Controls("lblQty").Caption = qty
            .Controls("lblUnit").Caption = unit
            .Controls("lblDescription").Caption = desc
            .Controls("lblUnitPrice").Caption = unitPrice
            .Controls("lblAmount").Caption = amount
            .Controls("lblDiscount").Caption = FormatNumber(activeSales.get_discount_total, 2)
            .Controls("lblTotalAmount").Caption = FormatNumber(activeSales.get_total_amount + activeSales.get_discount_total, 2)
            .Controls("lblNetTotal").Caption = FormatNumber(activeSales.get_total_amount, 2)
            
            'for sales order responsible data
            
            sales_order_responsible.loadToSalesOrder_Responsible
            .Controls("lblprepared_by").Caption = activeUser.username  'sales_order_responsible.prepared_by
            .Controls("lblchecked_by").Caption = sales_order_responsible.checked_by
            .Controls("lblposted_by").Caption = sales_order_responsible.posted_by
            .Controls("lbldelivered_by").Caption = sales_order_responsible.delivered_by
            
    End With
    
    dtaSalesOrder.Show 1
End Sub


Function save_sales() As Integer
   
        Dim insert As String
        'sales_order_no, responsible_customer, responsible_agent, discount, grand_total, net_total, tendered_amount, delivery_date
        insert = "INSERT INTO stock_out_transaction VALUES ('" & _
                 Me.transaction_id & "'," & _
                 Me.sold_to.customers_id & "," & _
                 Me.sold_to.mvaragent.agent_id & ",0," & _
                 Me.get_total_amount & "," & _
                 Me.get_total_amount & "," & _
                 Me.tendered_amount & "," & _
                 Me.change & ",'" & _
                 Me.date_transact & "','" & _
                 Me.prepared_by & "')"
       
        db.execute insert
        Call save_in_payment_type
        Call save_stock_out_records

End Function

Sub save_stock_out_records()
    Dim cart As New cart
    Dim items As New cart_items
    Dim stockout_id As Integer
    
    Set cart = Me.items_sold
    'QTY, UNIT, DESC, UNITPRICE, AMOUNT

    For Each items In cart
        stockout_id = items.saveStockout
        Call Me.addThisItemToCurrentSalesOrder(stockout_id)
    Next
End Sub
Sub save_in_payment_type()
    If Me.payment_type = PAYMENT_COD Then
        Dim cod_insert As String
        cod_insert = "INSERT INTO cod VALUES('" & Me.transaction_id & "','',NOW())"
        db.execute cod_insert
    ElseIf Me.payment_type = PAYMENT_ACCOUNT_RECEIVABLE Then
        Dim ar_insert As String
        ar_insert = "INSERT INTO account_receivable VALUES('" & Me.transaction_id & "','unsettled',NOW())"
        db.execute ar_insert
    End If
End Sub
Function get_grand_total_amount() As Double
    Dim items As New cart_items
    Dim sum As Double
    Dim cart As New cart
    Set cart = Me.items_sold
    
    If Me.sold_to.dealers_type = "dealer" Then
        For Each items In cart
            sum = sum + items.get_dealers_total_price_without_discount
        Next
    Else
        For Each items In cart
            sum = sum + items.get_total_price
        Next
    End If
    get_grand_total_amount = sum
End Function
Function get_total_amount() As Double
    Dim items As New cart_items
    Dim sum As Double
    Dim cart As New cart
    Set cart = Me.items_sold
    
    If Me.sold_to.dealers_type = "dealer" Then
        For Each items In cart
            sum = sum + items.get_dealers_total_price
        Next
    Else
        For Each items In cart
            sum = sum + items.get_total_price
        Next
    End If
    get_total_amount = sum
End Function
Function get_discount_total() As Double
    Dim items As New cart_items
    Dim sum As Double
    Dim cart As New cart
    Set cart = Me.items_sold
    
    For Each items In cart
        sum = sum + (items.discount * items.qty_purchased)
    Next
    get_discount_total = sum
End Function
Function get_tracking_total() As Double
    Dim items As New cart_items
    Dim sum As Double
    Dim cart As New cart
    Set cart = Me.items_sold
    
    For Each items In cart
        sum = sum + items.tracking_price
    Next
    get_tracking_total = sum
End Function
Public Property Set items_sold(ByVal vData As cart)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.items_sold = Form1
    Set mvaritems_sold = vData
End Property


Public Property Get items_sold() As cart
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.items_sold
    Set items_sold = mvaritems_sold
End Property



Public Property Let hasDiscount(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.hasDiscount = 5
    mvarhasDiscount = vData
End Property


Public Property Get hasDiscount() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.hasDiscount
    hasDiscount = mvarhasDiscount
End Property

Public Property Let date_transact(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.date_transact = 5
    mvardate_transact = vData
End Property


Public Property Get date_transact() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.date_transact
    date_transact = mvardate_transact
End Property


Public Property Let transaction_type(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.transaction_type = 5
    mvartransaction_type = vData
End Property


Public Property Get transaction_type() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.transaction_type
    transaction_type = mvartransaction_type
End Property



Public Property Let payment_type(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.payment_type = 5
    mvarpayment_type = vData
End Property


Public Property Get payment_type() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.payment_type
     payment_type = mvarpayment_type
End Property



Public Property Set sold_to(ByVal vData As Customers)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.sold_to = Form1
    Set mvarsold_to = vData
End Property


Public Property Get sold_to() As Customers
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sold_to
    Set sold_to = mvarsold_to
End Property



Public Property Let transaction_id(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.transaction_id = 5
    mvartransaction_id = vData
End Property


Public Property Get transaction_id() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.transaction_id
    transaction_id = mvartransaction_id
End Property

Sub loadSalesOrder(sales_order_no As String)
    Dim sql As String
    Dim rs As New ADODB.Recordset
    'sales_order_no, responsible_customer, responsible_agent, discount, grand_total, net_total, tendered_amount, change, delivery_date
    
    sql = "SELECT * from stock_out_transaction WHERE sales_order_no = '" & sales_order_no & "'"
    Set rs = db.execute(sql)
    If rs.RecordCount > 0 Then
            Me.transaction_id = rs.Fields("sales_order_no").Value
            Me.items_sold.load_cart (rs.Fields("sales_order_no").Value)
            
        On Error Resume Next
            If Val(rs.Fields("responsible_customer").Value) = 0 Then
                Me.isSoldToWalkIn = True
            Else
                Me.sold_to.load_customers (rs.Fields("responsible_customer").Value)
            End If
            
            Me.tendered_amount = rs.Fields("tendered_amount").Value
            Me.change = rs.Fields("change").Value
            Me.date_transact = rs.Fields("delivery_date").Value
            Me.info_grand_total = rs.Fields("grand_total").Value
            Me.info_net_total = rs.Fields("net_total").Value
            Me.info_discount = rs.Fields("discount").Value
            Me.prepared_by = rs.Fields("prepared_by").Value
        Me.hasDiscount = True
    Else
        Me.hasRecord = False
    End If
End Sub


Private Sub Class_Initialize()
    Dim x As Integer
    x = getLastRerefenceNo
    Me.transaction_id = "SO-" & String((7 - Len(x)), "0") & x
End Sub
Function updateReferenceNo()
    db.execute ("UPDATE stock_out_reference SET reference_no = reference_no + 1")
End Function
Function getLastRerefenceNo() As Integer
    Dim sql As String
    sql = "SELECT * FROM stock_out_reference"
    On Error Resume Next
    getLastRerefenceNo = Val(db.execute(sql).Fields(0).Value)
End Function

Sub printSalesReport(strpayment_type As String)

    If strpayment_type = "COD" Then
    
        Set dtaSales.DataSource = getAllCODTodaySales
            dtaSales.Sections(1).Controls("lblpaymenttype").Caption = "CASH ON DELIVERY"
            dtaSales.Sections(1).Controls("lbldate").Caption = FormatDateTime(Date, vbLongDate)
    ElseIf strpayment_type = "ACCOUNT RECEIVABLE" Then
        Set dtaSales.DataSource = getAllAccountReceivableTodaySales
        dtaSales.Sections(1).Controls("lblpaymenttype").Caption = "ACCOUNT RECEIVABLE"
        dtaSales.Sections(1).Controls("lbldate").Caption = FormatDateTime(Date, vbLongDate)
    Else
        Set dtaSales.DataSource = getTodayAllSales
        dtaSales.Sections(1).Controls("lblpaymenttype").Caption = "ALL SALES"
        dtaSales.Sections(1).Controls("lbldate").Caption = FormatDateTime(Date, vbLongDate)
    End If
   dtaSales.Show 1
End Sub

Sub updateRemarksToFullyPaid()
    'If Me.payment_type = PAYMENT_ACCOUNT_RECEIVABLE Then
        Dim update As String
        
        update = "UPDATE account_receivable SET remarks = 'fully paid' WHERE sales_order_no = '" & Me.transaction_id & "'"
        db.execute update
    'End If
End Sub
