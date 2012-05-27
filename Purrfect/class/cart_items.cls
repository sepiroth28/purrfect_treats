VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cart_items"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvaritem As New items 'local copy
Private mvarqty_purchased As Double 'local copy
Private mvaritem_price As Double 'local copy
Private mvartotal_price As Double 'local copy
'local variable(s) to hold property value(s)
Private mvardiscount As Double 'local copy
'local variable(s) to hold property value(s)
Private mvartracking_price As Double 'local copy
Public Property Let tracking_price(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.tracking_price = 5
    mvartracking_price = vData
End Property


Public Property Get tracking_price() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.tracking_price
     tracking_price = mvartracking_price
End Property



Public Property Let discount(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.discount = 5
    mvardiscount = vData
End Property


Public Property Get discount() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.discount
    discount = mvardiscount
End Property




Sub load_cart_item(stockout_id)
    Dim sql As String
    Dim rs As New ADODB.Recordset
    'stockout_id, item_id, qty_out, amount, discount, tracking_price
    
    sql = "SELECT * from stock_out WHERE stockout_id = " & stockout_id
    Set rs = db.execute(sql)
    On Error Resume Next
    If rs.RecordCount > 0 Then
        Me.Item.load_item (rs.Fields("item_id").Value)
        Me.item_price = rs.Fields("price").Value
        Me.discount = rs.Fields("discount").Value
        Me.tracking_price = rs.Fields("tracking_price").Value
        Me.qty_purchased = rs.Fields("qty_out").Value
        Me.total_price = rs.Fields("amount").Value
    End If
End Sub

Function saveStockout() As Integer
'stockout_id, item_id, qty_out, amount, discount, tracking_price,price
    Dim sql As String
    sql = "INSERT INTO stock_out VALUES (null," & Me.Item.item_id & "," & Me.qty_purchased & "," & Me.get_total_price & "," & Me.discount & "," & Me.tracking_price & "," & Me.item_price & ")"
    db.execute sql
    saveStockout = db.execute("select last_insert_id()").Fields(0).Value
End Function

Public Function get_total_price() As Double
    get_total_price = ((Me.item_price - Me.discount) + Me.tracking_price) * Me.qty_purchased
End Function
Public Function get_dealers_total_price() As Double
    get_dealers_total_price = ((Me.Item.dealers_price - Me.discount) + Me.tracking_price) * Me.qty_purchased
End Function
Public Function get_dealers_total_price_without_discount() As Double
    get_dealers_total_price_without_discount = ((Me.Item.dealers_price) + Me.tracking_price) * Me.qty_purchased
End Function
'local variable(s) to hold property value(s)
Public Property Let qty_purchased(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.qty_purchased = 5
    mvarqty_purchased = vData
End Property


Public Property Get qty_purchased() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.qty_purchased
    qty_purchased = mvarqty_purchased
End Property



Public Property Let total_price(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.total_price = 5
    mvartotal_price = vData
End Property


Public Property Get total_price() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.total_price
    total_price = mvartotal_price
End Property



Public Property Let item_price(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.item_price = 5
    mvaritem_price = vData
End Property


Public Property Get item_price() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.item_price
    item_price = mvaritem_price
End Property






Public Property Set Item(ByVal vData As items)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.item = Form1
    Set mvaritem = vData
End Property


Public Property Get Item() As items
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.item
    Set Item = mvaritem
End Property



Private Sub Class_Initialize()
Me.discount = 0
End Sub
