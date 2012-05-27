VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ItemDescription"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvaritem_code As String 'local copy
Private mvaritem_name As String 'local copy
Private mvaritem_description As String 'local copy
Private mvarimage As String 'local copy
Private mvarstatus As Byte 'local copy
Private mvarunit_of_measure As String 'local copy

Public Property Let unit_of_measure(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.unit_of_measure = 5
    mvarunit_of_measure = vData
End Property


Public Property Get unit_of_measure() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.unit_of_measure
    unit_of_measure = mvarunit_of_measure
End Property



Public Property Let status(ByVal vData As Byte)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.status = 5
    mvarstatus = vData
End Property


Public Property Get status() As Byte
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.status
     status = mvarstatus
End Property



Public Property Let image(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.image = 5
    mvarimage = vData
End Property


Public Property Get image() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.image
    image = mvarimage
End Property



Public Property Let item_description(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.item_description = 5
    mvaritem_description = vData
End Property


Public Property Get item_description() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.item_description
    item_description = mvaritem_description
End Property



Public Property Let item_name(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.item_name = 5
    mvaritem_name = vData
End Property


Public Property Get item_name() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.item_name
    item_name = mvaritem_name
End Property



Public Property Let item_code(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.item_code = 5
    mvaritem_code = vData
End Property


Public Property Get item_code() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.item_code
    item_code = mvaritem_code
End Property



