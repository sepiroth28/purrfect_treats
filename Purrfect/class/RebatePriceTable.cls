VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RebatePriceTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarid As Integer 'local copy
Private mvarqtyFrom As Double 'local copy
Private mvarqtyTo As Double 'local copy
Private mvarpriceApplied As Double 'local copy
Public Sub delete_rebate_price_table(id As Integer)
db.execute "DELETE FROM rebate_price_table WHERE id = " & id
End Sub

Public Sub load_rebate_price_table(id As Integer)
Dim sql As String
Dim rs As New ADODB.Recordset

sql = "SELECT * FROM rebate_price_table WHERE id = " & id & " LIMIT 1"
Set rs = db.execute(sql)

If rs.RecordCount > 0 Then
    Me.id = rs.Fields(0).Value
    Me.qtyFrom = Val(rs.Fields("qty_from").Value)
    Me.qtyTo = Val(rs.Fields("qty_to").Value)
    Me.priceApplied = Val(rs.Fields("applied_price").Value)
End If

Set rs = Nothing
End Sub

Public Sub save_rebate_price_table()
    Dim save As String
    'id, qty_from, qty_to, applied_price
    If id = 0 Then
        save = "INSERT INTO rebate_price_table VALUES(null," & _
                Me.qtyFrom & "," & Me.qtyTo & "," & Me.priceApplied & _
                ")"
    Else
        save = "UPDATE rebate_price_table SET qty_from=" & Me.qtyFrom & ", qty_to=" & Me.qtyTo & ", applied_price=" & Me.priceApplied & " WHERE id = " & Me.id
    End If
    db.execute save
End Sub

Public Property Let priceApplied(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.priceApplied = 5
    mvarpriceApplied = vData
End Property


Public Property Get priceApplied() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.priceApplied
    priceApplied = mvarpriceApplied
End Property



Public Property Let qtyTo(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.qtyTo = 5
    mvarqtyTo = vData
End Property


Public Property Get qtyTo() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.qtyTo
    qtyTo = mvarqtyTo
End Property



Public Property Let qtyFrom(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.qtyFrom = 5
    mvarqtyFrom = vData
End Property


Public Property Get qtyFrom() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.qtyFrom
     qtyFrom = mvarqtyFrom
End Property



Public Property Let id(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.id = 5
    mvarid = vData
End Property


Public Property Get id() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.id
    id = mvarid
End Property



Private Sub Class_Initialize()
Me.id = 0
End Sub
