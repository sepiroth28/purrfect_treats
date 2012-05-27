VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmManageCustomer 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manage Customer"
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   11805
   Icon            =   "frmManageCustomer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   11805
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   8115
      Left            =   60
      ScaleHeight     =   8085
      ScaleWidth      =   11625
      TabIndex        =   0
      Top             =   60
      Width           =   11655
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
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1200
         Visible         =   0   'False
         Width           =   5055
      End
      Begin VB.CommandButton cmdAddNewCustomer 
         Caption         =   "ADD NEW CUSTOMER"
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
         Left            =   8820
         TabIndex        =   1
         Top             =   960
         Width           =   2655
      End
      Begin MSComctlLib.ListView lsvCustomer 
         Height          =   6255
         Left            =   180
         TabIndex        =   3
         Top             =   1680
         Width           =   11295
         _ExtentX        =   19923
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
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Manage Customer"
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
         TabIndex        =   5
         Top             =   120
         Width           =   4995
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   180
         X2              =   13860
         Y1              =   540
         Y2              =   540
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Search Customer"
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
         TabIndex        =   4
         Top             =   900
         Visible         =   0   'False
         Width           =   2115
      End
   End
   Begin VB.Menu mnu_customer_menu 
      Caption         =   "Customer Menu"
      Begin VB.Menu mnu_so_history 
         Caption         =   "SO History"
      End
      Begin VB.Menu mnu_customer_rebates 
         Caption         =   "View Rebates"
      End
      Begin VB.Menu mnu_sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_change_agent 
         Caption         =   "Change Technician"
      End
      Begin VB.Menu mnu_sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_delete 
         Caption         =   "Delete"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmManageCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAddNewCustomer_Click()
editmode = False
frmCustomer.Show 1

End Sub

Private Sub Form_Load()

If activeUser.previliges.canDeleteCustomer = True Then
    mnu_customer_menu.Enabled = True
    mnu_delete.Enabled = True
Else
    mnu_customer_menu.Enabled = False
    mnu_delete.Enabled = False
End If

Call setCustomersColumns(lsvCustomer)
lsvCustomer.ColumnHeaders(1).width = 0
lsvCustomer.ColumnHeaders(2).width = 3000
lsvCustomer.ColumnHeaders(3).width = 5000
lsvCustomer.ColumnHeaders(4).width = 2000

Call loadAllCustomersToListview(lsvCustomer)

End Sub

Private Sub lsvCustomer_DblClick()
On Error Resume Next
    editmode = True
    activecustomer = Val(lsvCustomer.SelectedItem.Text)
    frmCustomer.Show 1
End Sub

Private Sub lsvCustomer_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    Me.PopupMenu mnu_customer_menu
End If
End Sub

Private Sub mnu_change_agent_Click()
If lsvCustomer.ListItems.Count > 0 Then
    activecustomer = Val(lsvCustomer.SelectedItem.Text)
    frmChangeAgent.Show 1
End If
End Sub

Private Sub mnu_customer_rebates_Click()
If lsvCustomer.ListItems.Count > 0 Then
    activeCustomerIdForRebate = Val(lsvCustomer.SelectedItem.Text)
    frmViewRebates.Show 1
End If
End Sub

Private Sub mnu_delete_Click()
    If MsgBox("Are you sure you want to delete?", vbYesNo, "Delete Customer") = vbYes Then
        deleteCustomer (Val(lsvCustomer.SelectedItem.Text))
        Call loadAllCustomersToListview(lsvCustomer)
    End If
End Sub

Private Sub mnu_so_history_Click()
If lsvCustomer.ListItems.Count > 0 Then
    customer_id_for_list_of_account_receivable = Val(lsvCustomer.SelectedItem.Text)
    frmCustomerAccountReceivable.Show 1
End If
End Sub
