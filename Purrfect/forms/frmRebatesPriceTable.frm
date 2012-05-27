VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmRebatesPriceTable 
   BackColor       =   &H00000080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rebate Price Table"
   ClientHeight    =   7065
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C7FEF3&
      ForeColor       =   &H80000008&
      Height          =   6915
      Left            =   60
      ScaleHeight     =   6885
      ScaleWidth      =   5565
      TabIndex        =   0
      Top             =   60
      Width           =   5595
      Begin VB.CommandButton cmdAdd 
         Caption         =   "ADD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4140
         TabIndex        =   9
         Top             =   6120
         Width           =   1275
      End
      Begin VB.TextBox txtAppliedPrice 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2640
         TabIndex        =   7
         Top             =   6120
         Width           =   1335
      End
      Begin VB.TextBox txtTo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1380
         TabIndex        =   5
         Top             =   6120
         Width           =   1095
      End
      Begin VB.TextBox txtFrom 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   180
         TabIndex        =   3
         Top             =   6120
         Width           =   1095
      End
      Begin MSComctlLib.ListView lsvRebatesPriceTable 
         Height          =   4995
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   8811
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "id"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Qty From"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Qty To"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Applied price"
            Object.Width           =   2999
         EndProperty
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Applied price"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2640
         TabIndex        =   8
         Top             =   5760
         Width           =   1275
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "qty to"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1380
         TabIndex        =   6
         Top             =   5760
         Width           =   525
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "qty from"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   4
         Top             =   5760
         Width           =   780
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Rebates Price Table"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   3735
      End
   End
   Begin VB.Menu mnu_rebate_table_file 
      Caption         =   "File"
      Begin VB.Menu mnu_rebate_table_rate_file_delete 
         Caption         =   "Delete"
      End
   End
End
Attribute VB_Name = "frmRebatesPriceTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim edit_rebate As RebatePriceTable
Private Sub cmdAdd_Click()
Dim rebate As New RebatePriceTable

With rebate
    .qtyFrom = Val(txtFrom.Text)
    .qtyTo = Val(txtTo.Text)
    .priceApplied = Val(txtAppliedPrice.Text)
    .save_rebate_price_table
End With

Call loadRebatePriceTable(lsvRebatesPriceTable)
txtFrom.Text = ""
txtTo.Text = ""
txtAppliedPrice.Text = ""

End Sub

Private Sub Form_Load()
Call loadRebatePriceTable(lsvRebatesPriceTable)
End Sub

Private Sub lsvRebatesPriceTable_DblClick()

If lsvRebatesPriceTable.ListItems.Count > 0 Then
    edit_rebate.load_rebate_price_table (Val(lsvRebatesPriceTable.SelectedItem.Text))
    txtFrom.Text = edit_rebate.qtyFrom
    txtTo.Text = edit_rebate.qtyTo
    txtAppliedPrice.Text = edit_rebate.priceApplied
End If

End Sub

Private Sub lsvRebatesPriceTable_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu mnu_rebate_table_file
    End If
End Sub

Private Sub mnu_rebate_table_rate_file_delete_Click()
Dim delete_rebate As New RebatePriceTable

If lsvRebatesPriceTable.ListItems.Count > 0 Then
    delete_rebate.delete_rebate_price_table Val(lsvRebatesPriceTable.SelectedItem.Text)
End If
 
Call loadRebatePriceTable(lsvRebatesPriceTable)
End Sub
