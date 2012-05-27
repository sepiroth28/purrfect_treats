Attribute VB_Name = "Helper_TrackingPrice"
Function getTrackingPriceOfCurrentCustomer(customer_id As Integer)
Dim address As String
Dim cus As New Customers
Dim temp() As String
cus.load_customers (customer_id)
temp = Split(cus.customers_add, ",")

If cus.customers_add <> "" Then
    If UBound(temp) Then
        address = temp(0)
        getTrackingPriceOfCurrentCustomer = getTrackingPrice(address)
    ElseIf UBound(temp) = 0 And cus.customers_add <> "" Then
        address = cus.customers_add
        getTrackingPriceOfCurrentCustomer = getTrackingPrice(address)
    Else
        getTrackingPriceOfCurrentCustomer = 0
    End If
End If

End Function

Function getTrackingPrice(municipal As String)
Dim sql As String
Dim rs As New ADODB.Recordset
Dim tracking_price As Double
'municipal_id, municipal_name, tracking_price
sql = "SELECT * FROM `municipalities` WHERE municipal_name='" & municipal & "'"
Set rs = db.execute(sql)
On Error Resume Next
If rs.RecordCount > 0 Then
    tracking_price = Val(rs.Fields("tracking_price").Value)
    getTrackingPrice = tracking_price
Else
    getTrackingPrice = 0
End If

Set rs = Nothing
End Function
