VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StockInCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"StockIn"
Attribute VB_Ext_KEY = "Member0" ,"StockIn"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable to hold collection
Private mCol As Collection
'local variable(s) to hold property value(s)
Private mvarreference_no As String 'local copy
Private mvarstock_in_to As String 'local copy
Private mvarremarks As String 'local copy
Private mvarprepared_by As String 'local copy
Private mvarapproved_by As String 'local copy
Private mvarreceived_by As String 'local copy
'local variable(s) to hold property value(s)
Private mvarsupplier_id As Integer 'local copy
Private mvarsupplier As manufacturers 'local copy
Private last_id As Integer
Public Sub insertStockInRecordToThisTransaction(last_insert_id As Integer)

    Dim insert As String
    'stock_in_transaction_id, stock_id
    insert = "INSERT INTO stock_in_transaction_to_stock_in_items VALUES(" & last_id & "," & last_insert_id & ")"
    db.execute insert
End Sub


Function get_last_id()
    get_last_id = last_id
End Function

Public Property Let supplier_id(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.supplier_id = 5
    mvarsupplier_id = vData
End Property


Public Property Get supplier_id() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.supplier_id
    supplier_id = mvarsupplier_id
End Property

Function get_total_items()
    Dim stock As New StockIn
    Dim total As Double
    For Each stock In mCol
        total = total + stock.QtyToBeAdd
    Next
    Set stock = Nothing
    get_total_items = total
End Function

Public Sub insert()
'stock_in_transaction_id, reference_no, stocked_in_to, from_supplier, remarks, stock_in_date, total_number_of_items, total_qty, prepared_by, approved_by, received_by
Dim insert As String


insert = "INSERT INTO stock_in_transaction " & _
            "VALUES (null," & _
            "'" & Me.reference_no & "'," & _
            "'" & Me.stock_in_to & "'," & _
            "" & Me.supplier_id & "," & _
            "'" & Me.remarks & "'," & _
            "CURDATE()," & _
            "" & Me.Count & "," & _
            "" & Me.get_total_items & "," & _
            "'" & Me.prepared_by & "'," & _
            "'" & Me.approved_by & "'," & _
            "'" & Me.received_by & "')"

db.execute insert
last_id = db.execute("select last_insert_id()").Fields(0).Value

End Sub

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



Public Property Let approved_by(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.approved_by = 5
    mvarapproved_by = vData
End Property


Public Property Get approved_by() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.approved_by
    approved_by = mvarapproved_by
End Property



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

Public Property Let stock_in_to(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.stock_in_to = 5
    mvarstock_in_to = vData
End Property


Public Property Get stock_in_to() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.stock_in_to
    stock_in_to = mvarstock_in_to
End Property



Public Property Let reference_no(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.reference_no = 5
    mvarreference_no = vData
End Property


Public Property Get reference_no() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.reference_no
    reference_no = mvarreference_no
End Property




Public Function Add(objNewMember As StockIn, Optional sKey As String) As StockIn
    'create a new object
'    Dim objNewMember As StockIn
'    Set objNewMember = New StockIn


    'set the properties passed into the method
'    Set objNewMember.items = items
'    objNewMember.QtyToBeAdd = QtyToBeAdd
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As StockIn
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
    Dim x As Integer
    x = getLastRerefenceNo
    Me.reference_no = "SI-" & String((7 - Len(x)), "0") & x
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub
Function updateReferenceNo()
    db.execute ("UPDATE stock_in_reference SET reference_no = reference_no + 1")
End Function
Function getLastRerefenceNo() As Integer
    Dim sql As String
    Dim rs As New ADODB.Recordset
    sql = "SELECT * FROM stock_in_reference"
    Set rs = db.execute(sql)
    If rs.RecordCount > 0 Then
        getLastRerefenceNo = Val(rs.Fields(0).Value)
    End If
End Function
