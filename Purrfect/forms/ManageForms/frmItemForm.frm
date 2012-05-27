VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmItemForm 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item form"
   ClientHeight    =   8790
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7305
   Icon            =   "frmItemForm.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   7305
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H80000008&
      Height          =   8655
      Left            =   60
      ScaleHeight     =   8625
      ScaleWidth      =   7125
      TabIndex        =   0
      Top             =   60
      Width           =   7155
      Begin VB.ComboBox cboCategory 
         Height          =   288
         Left            =   4920
         TabIndex        =   26
         Top             =   540
         Width           =   1872
      End
      Begin VB.TextBox txtDealersPrice 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3240
         TabIndex        =   24
         Tag             =   "*Item price"
         Top             =   4140
         Width           =   2955
      End
      Begin VB.TextBox txtQty 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3240
         TabIndex        =   5
         Tag             =   "*Item price"
         Top             =   5940
         Width           =   2955
      End
      Begin VB.TextBox txtNewUnitOfMeasure 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3240
         TabIndex        =   7
         Top             =   5100
         Width           =   2955
      End
      Begin VB.ComboBox cboUnitOfMeasure 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   6
         Text            =   "- Unit of measure -"
         Top             =   5100
         Width           =   2895
      End
      Begin MSComctlLib.ListView lsvManufacturers 
         Height          =   975
         Left            =   240
         TabIndex        =   21
         Top             =   7260
         Visible         =   0   'False
         Width           =   5955
         _ExtentX        =   10504
         _ExtentY        =   1720
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.CommandButton cmdAddNewItem 
         Caption         =   "SAVE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   4320
         TabIndex        =   11
         Top             =   7800
         Width           =   2655
      End
      Begin VB.CheckBox chkStatus 
         BackColor       =   &H80000018&
         Caption         =   "Enable"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   435
         Left            =   240
         TabIndex        =   19
         Top             =   7500
         Value           =   1  'Checked
         Width           =   2715
      End
      Begin VB.CommandButton cmdBrowseManufacturer 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "..."
         Height          =   435
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   6840
         Width           =   675
      End
      Begin VB.TextBox txtItemPrice 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   240
         TabIndex        =   4
         Tag             =   "*Item price"
         Top             =   4140
         Width           =   2895
      End
      Begin VB.TextBox txtManufacturers 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   240
         TabIndex        =   9
         Top             =   6840
         Width           =   5955
      End
      Begin VB.TextBox txtReorderPoint 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   240
         TabIndex        =   8
         Text            =   "0"
         Top             =   5940
         Width           =   2895
      End
      Begin VB.TextBox txtItemDescription 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   2760
         Width           =   6555
      End
      Begin VB.TextBox txtItemName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   240
         TabIndex        =   2
         Tag             =   "*Item name"
         Top             =   1860
         Width           =   6555
      End
      Begin VB.TextBox txtItemCode 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   240
         TabIndex        =   1
         Tag             =   "*Item code"
         Top             =   1020
         Width           =   6555
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Dealear's Price"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   432
         Left            =   3240
         TabIndex        =   25
         Top             =   3840
         Width           =   2112
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Qty"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   432
         Left            =   3240
         TabIndex        =   23
         Top             =   5640
         Width           =   2112
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Add new Unit of measure here..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   435
         Left            =   3300
         TabIndex        =   22
         Top             =   4800
         Width           =   3555
      End
      Begin VB.Label lblRequiredMsg 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "  Please fill up requireed fields..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   312
         Left            =   240
         TabIndex        =   20
         Top             =   120
         Visible         =   0   'False
         Width           =   6552
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Item Price"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   240
         TabIndex        =   18
         Top             =   3840
         Width           =   2115
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Manufacturer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   240
         TabIndex        =   17
         Top             =   6540
         Width           =   2115
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Reorder point"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   240
         TabIndex        =   16
         Top             =   5640
         Width           =   2115
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Unit of measure"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   240
         TabIndex        =   15
         Top             =   4740
         Width           =   2115
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Item Description"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   240
         TabIndex        =   14
         Top             =   2460
         Width           =   2115
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Item name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   240
         TabIndex        =   13
         Top             =   1560
         Width           =   2115
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Item Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   240
         TabIndex        =   12
         Top             =   720
         Width           =   2115
      End
   End
End
Attribute VB_Name = "frmItemForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim new_item As New items
Dim edit_item As New items

Private Sub cmdAddNewItem_Click()
Dim validated As Boolean
validated = validateRequiredValueInForm(frmItemForm)
If validated Then
     lblRequiredMsg.Visible = False
   With new_item
        .item_code = txtItemCode.Text
        .item_name = txtItemName.Text
        .item_description = txtItemDescription.Text
        .item_price = Val(txtItemPrice.Text)
        .dealers_price = Val(txtDealersPrice.Text)
        .item_qty = Val(txtQty.Text)
    
        If cboUnitOfMeasure.Text = "- Unit of measure -" Then
            If txtNewUnitOfMeasure.Text = "" Then
                .unit_of_measure = ""
            Else
                .unit_of_measure = txtNewUnitOfMeasure.Text
            End If
        Else
            .unit_of_measure = cboUnitOfMeasure.Text
        End If
        .reorder_point = txtReorderPoint.Text
       
'        .manufacturers_id = lsvManufacturers.SelectedItem.Text
        .item_status = chkStatus.Value
        
        '#####checkSql is for debugging use, this displays the sql insert statement on items and items_description table
        
         If editmode = True Then
                With edit_item
                    .item_code = txtItemCode.Text
                    .item_name = txtItemName.Text
                    .item_description = txtItemDescription.Text
                    .item_price = Val(txtItemPrice.Text)
                    .item_qty = Val(txtQty.Text)
                    .dealers_price = Val(txtDealersPrice.Text)
                    .unit_of_measure = cboUnitOfMeasure.Text
                    '.manufacturers_id = editManufacturersId
                    .update
                    .item_category.item_code = .item_code
                    .item_category.category = cboCategory.Text
                    .item_category.update
               ' .checkSql
               End With
                editmode = False
                activeItemId = 0
                editManufacturersId = 0
                
        Else
            .manufacturers_id = lsvManufacturers.SelectedItem.Text
            .insert
            .item_category.item_code = .item_code
            .item_category.category = cboCategory.Text
            .item_category.save
            editmode = False
             MsgBox "Save successfully!"
        End If
        Call loadAllItemsToListview(frmManageItem.lsvItemList, "item_code")
        Set edit_item = Nothing
        Unload Me
   End With
Else
     Dim msg As String
    
       For Each col In validate_msg
         msg = msg & "Please input " & col & vbCrLf
       Next
    Set validate_msg = Nothing
    Set validate_msg = New Collection
    MsgBox msg
    lblRequiredMsg.Visible = True
End If

End Sub

Private Sub cmdBrowseManufacturer_Click()
Call toogleListView(lsvManufacturers)
End Sub

Private Sub Form_Load()

Call setManufacturersColumns(lsvManufacturers)
Call loadAllmanufacturersToListview(lsvManufacturers)
        lsvManufacturers.ColumnHeaders(1).width = 0
        lsvManufacturers.ColumnHeaders(2).width = 5000
        lsvManufacturers.ColumnHeaders(3).width = 0
        lsvManufacturers.ColumnHeaders(4).width = 0
'        lsvManufacturers.ColumnHeaders(5).width = 0
Call loadUnitOfMeasureToCombo(cboUnitOfMeasure)
cboCategory.AddItem "feeds"
cboCategory.AddItem "rice"
cboCategory.AddItem "fish"
cboCategory.AddItem "medicine"
cboCategory.AddItem "Accessories"

    If editmode = True Then

        edit_item.load_item activeItemId
        edit_item.item_category.load_item_category (edit_item.item_code)
        txtItemCode.Text = edit_item.item_code
        
        txtItemName.Text = edit_item.item_name
        txtItemDescription.Text = edit_item.item_description
        txtItemPrice.Text = edit_item.item_price
        txtDealersPrice.Text = edit_item.dealers_price
        txtQty.Text = edit_item.item_qty
        cboUnitOfMeasure.Text = edit_item.unit_of_measure
        txtReorderPoint.Text = edit_item.reorder_point
        txtManufacturers.Text = edit_item.manufacturer.manufacturers_name
        cboCategory.Text = edit_item.item_category.category
    End If



End Sub


Private Sub Form_Unload(Cancel As Integer)
editmode = False
activeItemId = 0
End Sub

Private Sub lsvManufacturers_Click()
On Error Resume Next
If editmode = True Then
    edit_item.manufacturers_id = Val(lsvManufacturers.SelectedItem.Text)
    txtManufacturers.Text = lsvManufacturers.SelectedItem.SubItems(1)
    Call toogleListView(lsvManufacturers)
Else
    txtManufacturers.Text = lsvManufacturers.SelectedItem.SubItems(1)
    Call toogleListView(lsvManufacturers)
End If
End Sub

Private Sub txtItemPrice_KeyPress(KeyAscii As Integer)
 Select Case KeyAscii
   Case 8, 48 To 57  ' BS, 0 - 9
   Case 46        ' . (decimal sep)
     If InStr(txtItemPrice.Text, ".") > 0 Then
       KeyAscii = 0
     End If
   Case Else
     KeyAscii = 0
 End Select

End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
 Select Case KeyAscii
   Case 8, 48 To 57  ' BS, 0 - 9
    Case 46        ' . (decimal sep)
     If InStr(txtItemPrice.Text, ".") > 0 Then
       KeyAscii = 0
     End If
   Case Else
     KeyAscii = 0
 End Select

End Sub

Private Sub txtReorderPoint_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 8, 48 To 57  ' BS, 0 - 9
    Case Else
         KeyAscii = 0
     End Select
End Sub
