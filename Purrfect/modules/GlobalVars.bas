Attribute VB_Name = "GlobalVars"
Public db As New db
Public editmode As Boolean
Public activeItemId As Integer
Public activeStockInList As StockInCollection
Public municipal_list As String
Public activemunicipalID As Integer
Public activecustomer As Integer
Public activeDiscout_id As Integer
Public activeSales As Sales
Public activeaseraccount_name As String
Public activeUser As New User_Account

Public activeDateTextbox As TextBox
Public activeDate As Date
Public activestockId As Integer
Public activeReprintStockIN As Integer
Public activeSalesOrderForViewSales As String
Public activeSalesOrderForViewSalesDetails As String
Public activeSalesOrderForPaymentHistory As String
Public activeCustomerIdForRebate As Integer


'rebate variable
Public rebate_grand_total As Double
Public rebate_grand_total_qty As Double

Public selectedSOForHistory As String

Public customer_id_for_list_of_account_receivable As Integer


Public editManufacturer As Boolean
Public edit_manufacturer_id As Integer

Public amount_to_be_debt As Double

Sub resetAllGlobalVars()
Set activeSales = New Sales
Set activeUser = New User_Account
Set activeStockInList = New StockInCollection
End Sub
