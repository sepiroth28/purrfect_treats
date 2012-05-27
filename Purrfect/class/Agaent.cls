VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Agaent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvaragent_name As String 'local copy
Private mvaragent_address As String 'local copy
Private mvaragent_contact_number As String 'local copy
Public Sub UpdateAgent()

End Sub

Public Sub SaveAgent()

End Sub

Public Property Let agent_contact_number(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.agent_contact_number = 5
    mvaragent_contact_number = vData
End Property


Public Property Get agent_contact_number() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.agent_contact_number
    agent_contact_number = mvaragent_contact_number
End Property



Public Property Let agent_address(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.agent_address = 5
    mvaragent_address = vData
End Property


Public Property Get agent_address() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.agent_address
    agent_address = mvaragent_address
End Property



Public Property Let agent_name(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.agent_name = 5
    mvaragent_name = vData
End Property


Public Property Get agent_name() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.agent_name
    agent_name = mvaragent_name
End Property



