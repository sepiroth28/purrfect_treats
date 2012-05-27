VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Customers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarcustomers_id As String
Private mvarcustomers_name As String
Private mvarcustomers_add As String
Private mvarcustomers_number As String
Public mvaragent As New agent
'local variable(s) to hold property value(s)
Private mvardealers_type As String 'local copy
Public credit_limit As Double

Public Property Let dealers_type(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dealers_type = 5
    mvardealers_type = vData
End Property


Public Property Get dealers_type() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dealers_type
     dealers_type = mvardealers_type
End Property


Public Property Let customers_id(ByVal vData As String)
    mvarcustomers_id = vData
End Property

Public Property Get customers_id() As String
    customers_id = mvarcustomers_id
End Property

Public Property Let customers_name(ByVal vData As String)
    mvarcustomers_name = vData
End Property

Public Property Get customers_name() As String
    customers_name = mvarcustomers_name
End Property

Public Property Let customers_add(ByVal vData As String)
    mvarcustomers_add = vData
End Property

Public Property Get customers_add() As String
    customers_add = mvarcustomers_add
End Property

Public Property Let customers_number(ByVal vData As String)
    mvarcustomers_number = vData
End Property
Public Property Get customers_number() As String
    customers_number = mvarcustomers_number
End Property

Function insert() As Integer
    Dim tbl_insert_customer As String
    'customers_id, customers_name, customers_add, customers_number, dealers_type, credit_limit, visible
    tbl_insert_customer = "INSERT INTO customers VALUES(" & Me.customers_id & ",'" & Me.customers_name & "', " & _
                            "'" & Me.customers_add & "', '" & Me.customers_number & "','" & Me.dealers_type & "',0,1)"
    db.execute tbl_insert_customer
    insert = db.execute("select last_insert_id()").Fields(0).Value
End Function

Public Sub updateData()
    Dim tbl_update_customers As String
 '   tbl_update_customers = "UPDATE FROM customers SET customers_name = '" & Me.customers_name & "',customers_add  = '" & Me.customers_add & "',customers_num = " & Me.customers_number & "   WHERE customers_id = " & Me.customers_id


    tbl_update_customers = "UPDATE  customers SET customers_name = '" & Me.customers_name & "', " & _
                           "customers_add = '" & Me.customers_add & "'," & _
                           "customers_number = '" & Me.customers_number & "',dealers_type='" & Me.dealers_type & "' WHERE customers_id = " & Me.customers_id & ""
    db.execute tbl_update_customers
End Sub

Public Function isExist() As Boolean
    Dim rs As New ADODB.Recordset
    Dim tbl_isExist_customer As String
    Dim tbl_result
        tbl_isExist_customer = "SELECT * FROM customers " & _
                               " WHERE customers_name = " & _
                               "'" & Me.customers_name & "'"
    
End Function

Public Function load_customers(customers_id As Integer) As Customers

    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim temp As New Customers
    
    On Error Resume Next
'    customers_id, customers_name, customers_add, customers_number
    sql = "SELECT  c.customers_id, c.customers_name, c.customers_add, c.customers_number,ac.agent_id,c.dealers_type,c.credit_limit FROM customers c LEFT JOIN agent_customers ac ON c.customers_id = ac.customers_id WHERE c.customers_id =" & customers_id
    'setCustomersID (rs.Fields("customers_id").Value)
    Set rs = db.execute(sql)
    'On Error Resume Next
    Me.customers_id = rs.Fields(0).Value
    Me.customers_name = rs.Fields(1).Value
    Me.customers_add = rs.Fields(2).Value
    Me.customers_number = rs.Fields(3).Value
    Me.dealers_type = rs.Fields(5).Value
    If rs.Fields(4).Value Then
        mvaragent.load_agent (Val(rs.Fields(4).Value))
    Else
        Set mvaragent = Nothing
        Set mvaragent = New agent
    End If
    Me.credit_limit = rs.Fields("credit_limit").Value
    
End Function

Private Sub Class_Initialize()
    Me.customers_id = "Null"
End Sub
