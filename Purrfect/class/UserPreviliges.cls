VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UserPreviliges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"User_Previleges"
'local variable(s) to hold property value(s)
Private mvarcanManageItem As Boolean 'local copy
Private mvarcanManageCustomer As Boolean 'local copy
Private mvarcanManageAgent As Boolean 'local copy
Private mvarcanStockIn As Boolean 'local copy
Private mvarcanView As Boolean 'local copy
Private mvarcanProcessPayment As Boolean 'local copy
Private mvarcanProcessInventory As Boolean 'local copy
Private mvarcanProcessSales As Boolean 'local copy
Private mvarusername As String 'local copy
Private mvarusertype As String 'local copy
'local variable(s) to hold property value(s)
Private mvarcanManageUserAccount As Boolean 'local copy
Private mvarcanManageSalesRepresentative As Boolean 'local copy
'local variable(s) to hold property value(s)
Private mvarcanViewSales As Boolean 'local copy
'local variable(s) to hold property value(s)
Private mvarcanDeleteCustomer As Boolean 'local copy
Private mvarUser_Previleges As User_Previleges
'local variable(s) to hold property value(s)
Private mvarcanManageCustomerVisibility As Boolean 'local copy
Public Property Let canManageCustomerVisibility(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.canManageCustomerVisibility = 5
    mvarcanManageCustomerVisibility = vData
End Property


Public Property Get canManageCustomerVisibility() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.canManageCustomerVisibility
    canManageCustomerVisibility = mvarcanManageCustomerVisibility
End Property





Private Sub Class_Initialize()
    'create the mUser_Previleges object when the UserPreviliges class is created
    Set mvarUser_Previleges = New User_Previleges
End Sub


Public Property Get User_Previleges() As User_Previleges
    Set User_Previleges = mvarUser_Previleges
End Property


Public Property Set User_Previleges(vData As User_Previleges)
    Set mvarUser_Previleges = vData
End Property
Private Sub Class_Terminate()
    Set mvarUser_Previleges = Nothing
End Sub



Public Property Let canDeleteCustomer(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.canDeleteCustomer = 5
    mvarcanDeleteCustomer = vData
End Property


Public Property Get canDeleteCustomer() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.canDeleteCustomer
    canDeleteCustomer = mvarcanDeleteCustomer
End Property



Public Property Let canViewSales(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.canViewSales = 5
    mvarcanViewSales = vData
End Property


Public Property Get canViewSales() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.canViewSales
    canViewSales = mvarcanViewSales
End Property



Public Property Let canManageSalesRepresentative(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.canManageSalesRepresentative = 5
    mvarcanManageSalesRepresentative = vData
End Property


Public Property Get canManageSalesRepresentative() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.canManageSalesRepresentative
    canManageSalesRepresentative = mvarcanManageSalesRepresentative
End Property



Public Property Let canManageUserAccount(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.canManageUserAccount = 5
    mvarcanManageUserAccount = vData
End Property


Public Property Get canManageUserAccount() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.canManageUserAccount
    canManageUserAccount = mvarcanManageUserAccount
End Property



Public Sub grantAll()
    Me.canManageAgent = True
    Me.canManageCustomer = True
    Me.canManageItem = True
    Me.canProcessInventory = True
    Me.canProcessPayment = True
    Me.canProcessSales = True
    Me.canStockIn = True
    Me.canView = True
    Me.canManageSalesRepresentative = True
    Me.canManageUserAccount = True
    Me.canViewSales = True
End Sub
Public Sub clearPreviliges()
    Me.canManageAgent = False
    Me.canManageCustomer = False
    Me.canManageItem = False
    Me.canProcessInventory = False
    Me.canProcessPayment = False
    Me.canProcessSales = False
    Me.canStockIn = False
    Me.canView = False
End Sub
Public Property Let usertype(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.usertype = 5
    mvarusertype = vData
End Property


Public Property Get usertype() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.usertype
     usertype = mvarusertype
End Property



Public Property Let username(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.username = 5
    mvarusername = vData
End Property


Public Property Get username() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.username
    username = mvarusername
End Property



Public Property Let canProcessSales(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.canProcessSales = 5
    mvarcanProcessSales = vData
End Property


Public Property Get canProcessSales() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.canProcessSales
    canProcessSales = mvarcanProcessSales
End Property



Public Property Let canProcessInventory(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.canProcessInventory = 5
    mvarcanProcessInventory = vData
End Property


Public Property Get canProcessInventory() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.canProcessInventory
    canProcessInventory = mvarcanProcessInventory
End Property



Public Property Let canProcessPayment(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.canProcessPayment = 5
    mvarcanProcessPayment = vData
End Property


Public Property Get canProcessPayment() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.canProcessPayment
    canProcessPayment = mvarcanProcessPayment
End Property



Public Property Let canView(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.canView = 5
    mvarcanView = vData
End Property


Public Property Get canView() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.canView
    canView = mvarcanView
End Property



Public Property Let canStockIn(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.canStockIn = 5
    mvarcanStockIn = vData
End Property


Public Property Get canStockIn() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.canStockIn
    canStockIn = mvarcanStockIn
End Property



Public Property Let canManageAgent(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.canManageAgent = 5
    mvarcanManageAgent = vData
End Property


Public Property Get canManageAgent() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.canManageAgent
    canManageAgent = mvarcanManageAgent
End Property



Public Property Let canManageCustomer(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.canManageCustomer = 5
    mvarcanManageCustomer = vData
End Property


Public Property Get canManageCustomer() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.canManageCustomer
    canManageCustomer = mvarcanManageCustomer
End Property



Public Property Let canManageItem(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.canManageItem = 5
    mvarcanManageItem = vData
End Property


Public Property Get canManageItem() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.canManageItem
    canManageItem = mvarcanManageItem
End Property



