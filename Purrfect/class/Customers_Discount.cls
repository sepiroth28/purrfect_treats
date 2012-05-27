VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Customers_Discount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvardiscount_id As String 'local copy
Private mvardiscount_code As String 'local copy
Private mvardiscount_name As String 'local copy
Private mvaramount As Double 'local copy

Public Sub Update_Discount()
    Dim tbl_update As String
    
    'discount_id, discount_code, discount_name, amount
    tbl_update = " UPDATE discount SET discount_code ='" & Me.discount_code & "'," & _
                 "discount_name = '" & Me.discount_name & "'," & _
                 " amount = " & Me.discount_amount & "" & _
                 "  WHERE discount_id = " & Me.discount_id
    db.execute (tbl_update)
    MsgBox "Record has been successfully save.", vbInformation, "Nutrimart"
    
End Sub

Public Sub Save_Discount()

    Dim tbl_save As String
    
    tbl_save = "INSERT INTO discount" & _
              " VALUES(NULL " & _
              ", '" & Me.discount_code & "'," & _
              " '" & Me.discount_name & "'," & _
              " " & Me.discount_amount & ")"
    db.execute (tbl_save)
    MsgBox "Record has been successfully save.", vbInformation, "Nutrimart"
End Sub

Public Function load_discount(discount_id As Integer) As Customers_Discount
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim temp As New Customers_Discount
    
    sql = "SELECT * FROM discount WHERE discount_id =" & discount_id
    
    Set rs = db.execute(sql)
    
    Me.discount_id = rs.Fields(0).Value
    Me.discount_code = rs.Fields(1).Value
    Me.discount_name = rs.Fields(2).Value
    Me.discount_amount = rs.Fields(3).Value
    
    Set rs = Nothing
    
    
End Function

Public Property Let discount_amount(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.amount = 5
    mvaramount = vData
End Property


Public Property Get discount_amount() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.amount
    discount_amount = mvaramount
End Property



Public Property Let discount_name(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.discount_name = 5
    mvardiscount_name = vData
End Property


Public Property Get discount_name() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.discount_name
    discount_name = mvardiscount_name
End Property



Public Property Let discount_code(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.discount_code = 5
    mvardiscount_code = vData
End Property


Public Property Get discount_code() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.discount_code
    discount_code = mvardiscount_code
End Property



Public Property Let discount_id(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.discount_id = 5
    mvardiscount_id = vData
End Property


Public Property Get discount_id() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.discount_id
    discount_id = mvardiscount_id
End Property


Private Sub Class_Initialize()
    Me.discount_id = "NULL"
End Sub
