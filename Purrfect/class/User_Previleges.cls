VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "User_Previleges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
'local variable(s) to hold property value(s)
Private mvarcanProcessPayment As Boolean 'local copy
Private mvarcanAddCustomer As Boolean 'local copy
Private mvarcanStockIn As Boolean 'local copy
Private mvarcanInventory As Boolean 'local copy
Private mvarcanStockOut As Boolean 'local copy
Private mvarcanManagetechnician As Boolean 'local copy
Private mvarcanManageManufacturer As Boolean 'local copy
Private mvarcanManageSORep As Boolean 'local copy
Private mvarcanManageUserAccount As Boolean 'local copy
Private mvarcanViewSales As Boolean 'local copy
Private mvarcanPrintSalesDetails As Boolean 'local copy
Private mvarcanManageCreditLimit As Boolean 'local copy
Private mvarcanViewStockIn As Boolean 'local copy
Private mvarcanSalesAdjustment As Boolean 'local copy
Private mvarcanPrintReceipt As Boolean 'local copy
Private mvarcanDeleteCustomer As Boolean 'local copy
'local variable(s) to hold property value(s)
Private mvarcanManageItem As Boolean 'local copy
Private mvarcanDeleteItem As Boolean 'local copy
Private mvarcanManageCustomerVisibility As Boolean 'local copy
'local variable(s) to hold property value(s)
Private mvarcan_accept_remit_payments As Boolean 'local copy
'local variable(s) to hold property value(s)
Private mvarcan_issue_rebate As Boolean 'local copy
Public Property Let can_issue_rebate(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.can_issue_rebate = 5
    mvarcan_issue_rebate = vData
End Property


Public Property Get can_issue_rebate() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.can_issue_rebate
    can_issue_rebate = mvarcan_issue_rebate
End Property



Public Property Let can_accept_remit_payments(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.can_accept_remit_payments = 5
    mvarcan_accept_remit_payments = vData
End Property


Public Property Get can_accept_remit_payments() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.can_accept_remit_payments
    can_accept_remit_payments = mvarcan_accept_remit_payments
End Property



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

Public Property Let canDeleteItem(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.canDeleteItem = 5
    mvarcanDeleteItem = vData
End Property


Public Property Get canDeleteItem() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.canDeleteItem
    canDeleteItem = mvarcanDeleteItem
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



Public Property Let canPrintReceipt(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.canPrintReceipt = 5
    mvarcanPrintReceipt = vData
End Property


Public Property Get canPrintReceipt() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.canPrintReceipt
     canPrintReceipt = mvarcanPrintReceipt
End Property



Public Property Let canSalesAdjustment(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.canSalesAdjustment = 5
    mvarcanSalesAdjustment = vData
End Property


Public Property Get canSalesAdjustment() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.canSalesAdjustment
     canSalesAdjustment = mvarcanSalesAdjustment
End Property



Public Property Let canViewStockIn(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.canViewStockIn = 5
    mvarcanViewStockIn = vData
End Property


Public Property Get canViewStockIn() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.canViewStockIn
     canViewStockIn = mvarcanViewStockIn
End Property



Public Property Let canManageCreditLimit(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.canManageCreditLimit = 5
    mvarcanManageCreditLimit = vData
End Property


Public Property Get canManageCreditLimit() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.canManageCreditLimit
     canManageCreditLimit = mvarcanManageCreditLimit
End Property



Public Property Let canPrintSalesDetails(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.canPrintSalesDetails = 5
    mvarcanPrintSalesDetails = vData
End Property


Public Property Get canPrintSalesDetails() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.canPrintSalesDetails
     canPrintSalesDetails = mvarcanPrintSalesDetails
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



Public Property Let canManageSORep(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.canManageSORep = 5
    mvarcanManageSORep = vData
End Property


Public Property Get canManageSORep() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.canManageSORep
     canManageSORep = mvarcanManageSORep
End Property



Public Property Let canManageManufacturer(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.canManageManufacturer = 5
    mvarcanManageManufacturer = vData
End Property


Public Property Get canManageManufacturer() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.canManageManufacturer
     canManageManufacturer = mvarcanManageManufacturer
End Property



Public Property Let canManagetechnician(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.canManagetechnician = 5
    mvarcanManagetechnician = vData
End Property


Public Property Get canManagetechnician() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.canManagetechnician
     canManagetechnician = mvarcanManagetechnician
End Property



Public Property Let canStockOut(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.canStockOut = 5
    mvarcanStockOut = vData
End Property


Public Property Get canStockOut() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.canStockOut
     canStockOut = mvarcanStockOut
End Property



Public Property Let canInventory(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.canInventory = 5
    mvarcanInventory = vData
End Property


Public Property Get canInventory() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.canInventory
     canInventory = mvarcanInventory
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



Public Property Let canAddCustomer(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.canAddCustomer = 5
    mvarcanAddCustomer = vData
End Property


Public Property Get canAddCustomer() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.canAddCustomer
     canAddCustomer = mvarcanAddCustomer
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



