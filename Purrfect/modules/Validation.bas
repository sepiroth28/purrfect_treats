Attribute VB_Name = "Validation"
'Validation
Public validate_msg As New Collection

Function validateRequiredValueInForm(frm As Form) As Boolean
Dim cntl As Control
validateRequiredValueInForm = True

For Each cntl In frm.Controls
    If TypeOf cntl Is TextBox Then
        If Mid(cntl.Tag, 1, 1) = "*" Then
            If cntl.Text = "" Then
                cntl.BackColor = &H80C0FF
                validate_msg.Add Mid(cntl.Tag, 2)
                validateRequiredValueInForm = False
                
            End If
        End If
    End If
Next

End Function
Function ClearFields(frm As Form) As Boolean
    Dim cntl As Control
    
    For Each cntl In frm.Controls
        If TypeOf cntl Is TextBox Then
            cntl = " "
        End If
    Next
    
End Function
