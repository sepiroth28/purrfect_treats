VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StockIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"items"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvaritems As New items 'local copy
Private mvarQtyToBeAdd As Double 'local copy
Private last_id As Integer
Function get_last_id()
    get_last_id = last_id
End Function

Public Sub insert()
    'stockin_id, item_id, qty_in
Dim insert As String

insert = "INSERT INTO stock_in " & _
            "VALUES (null," & _
            "" & Me.items.item_id & "," & _
            "" & Me.QtyToBeAdd & ")"
db.execute insert
last_id = db.execute("select last_insert_id()").Fields(0).Value
End Sub

Public Property Let QtyToBeAdd(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.QtyToBeAdd = 5
    mvarQtyToBeAdd = vData
End Property


Public Property Get QtyToBeAdd() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.QtyToBeAdd
     QtyToBeAdd = mvarQtyToBeAdd
End Property



Public Property Set items(ByVal vData As items)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.items = Form1
    Set mvaritems = vData
End Property


Public Property Get items() As items
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.items
    Set items = mvaritems
End Property
Sub print_stock_in(ByVal iPrintStockIn As Integer)

    If iPrintStockIn = 2 Then
         Set dtaAll_Stockin.DataSource = getStockInRecordsByDate(Format(activeDate, "yyyy-mm-dd"))
             dtaAll_Stockin.Sections(1).Controls("lbldate").Caption = FormatDateTime(Date, vbLongDate)
    Else
         Set dtaAll_Stockin.DataSource = getAllStockInRecords
             dtaAll_Stockin.Sections(1).Controls("lbldate").Caption = FormatDateTime(Date, vbLongDate)
    End If
        dtaAll_Stockin.Show 1
End Sub


