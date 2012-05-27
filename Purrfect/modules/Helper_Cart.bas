Attribute VB_Name = "Helper_Cart"
Public Sub updateCartInfo()
    
End Sub
Function getPriceToBeUsed(items As cart_items) As Double
    If activeSales.sold_to.dealers_type = DEALER Then
        getPriceToBeUsed = items.Item.dealers_price
    ElseIf activeSales.sold_to.dealers_type = CONSUMER Then
        getPriceToBeUsed = items.Item.item_price
    End If
End Function
