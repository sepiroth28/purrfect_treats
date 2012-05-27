VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmManageItem 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manage Item"
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   14280
   Icon            =   "frmManageItem.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   14280
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   8115
      Left            =   60
      ScaleHeight     =   8085
      ScaleWidth      =   14085
      TabIndex        =   0
      Top             =   60
      Width           =   14115
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Refresh"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   7140
         TabIndex        =   7
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
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
         Left            =   5400
         TabIndex        =   6
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CommandButton cmdAddNewItem 
         Caption         =   "ADD NEW ITEM"
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
         Left            =   11280
         TabIndex        =   5
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox txtSearchItemCode 
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
         Left            =   180
         TabIndex        =   4
         Top             =   1200
         Width           =   5055
      End
      Begin MSComctlLib.ListView lsvItemList 
         Height          =   6255
         Left            =   180
         TabIndex        =   2
         Top             =   1680
         Width           =   13755
         _ExtentX        =   24262
         _ExtentY        =   11033
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Search Itemcode"
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
         Left            =   180
         TabIndex        =   3
         Top             =   900
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   180
         X2              =   13860
         Y1              =   540
         Y2              =   540
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Manage Item"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   180
         TabIndex        =   1
         Top             =   120
         Width           =   2115
      End
   End
   Begin VB.Menu mnu_item_menu 
      Caption         =   "Item Menu"
      Begin VB.Menu mnu_delete_item 
         Caption         =   "delete"
      End
   End
End
Attribute VB_Name = "frmManageItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAddNewItem_Click()
frmItemForm.Show 1
End Sub

Private Sub cmdRefresh_Click()
txtSearchItemCode.Text = ""
Call loadAllItemsToListview(lsvItemList, "item_code")
End Sub

Private Sub cmdSearch_Click()
Call loadSearchItemsToListview(lsvItemList, txtSearchItemCode.Text)
End Sub

Private Sub Form_Load()
Call setItemsDescriptionColumns(lsvItemList)
lsvItemList.ColumnHeaders(1).width = 0
lsvItemList.ColumnHeaders(2).width = 1500
lsvItemList.ColumnHeaders(3).width = 2500
lsvItemList.ColumnHeaders(4).width = 2500
lsvItemList.ColumnHeaders(5).Alignment = lvwColumnRight
lsvItemList.ColumnHeaders(6).Alignment = lvwColumnRight
lsvItemList.ColumnHeaders(7).width = 1900
lsvItemList.ColumnHeaders(8).width = 1900
Call loadAllItemsToListview(lsvItemList, "item_code")


If activeUser.previliges.canDeleteItem = True Then
    mnu_delete_item.Enabled = True
Else
    mnu_delete_item.Enabled = False
End If

End Sub

Private Sub lsvItemList_DblClick()
    editmode = True
If lsvItemList.ListItems.Count > 0 Then
    activeItemId = Val(lsvItemList.SelectedItem.Text)
    frmItemForm.Show 1
End If
End Sub

Private Sub lsvItemList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    PopupMenu mnu_item_menu
End If
End Sub

Private Sub lsvItemList_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
   ' Me.PopupMenu mnu_item_menu
End If
End Sub

Private Sub mnu_delete_item_Click()
Dim ans As Byte
If MsgBox("Are you sure you want to delete? Item will be loss and may affect other important information... Please ask admin for assisstance", vbExclamation + vbYesNoCancel, "Delete Item") = vbYes Then
        Call deleteItem(lsvItemList.SelectedItem.SubItems(1))
        Call loadAllItemsToListview(lsvItemList, "item_code")
End If
End Sub

