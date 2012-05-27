VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAgingAccounts 
   Appearance      =   0  'Flat
   BackColor       =   &H00000080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Aging Accounts"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8070
   Icon            =   "frmAgingAccounts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   8070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   8055
      Left            =   60
      ScaleHeight     =   8025
      ScaleWidth      =   7905
      TabIndex        =   0
      Top             =   60
      Width           =   7935
      Begin VB.CommandButton cmdSearch 
         Caption         =   "SEARCH"
         Height          =   555
         Left            =   5940
         TabIndex        =   7
         Top             =   480
         Width           =   1815
      End
      Begin VB.ComboBox cboMonths 
         Height          =   315
         Left            =   1620
         TabIndex        =   5
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   5220
         TabIndex        =   1
         Top             =   7260
         Width           =   2532
      End
      Begin MSComctlLib.ListView lsvCustomers 
         Height          =   6075
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   7635
         _ExtentX        =   13467
         _ExtentY        =   10716
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "id"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Customer Name"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "# of unsettled SO"
            Object.Width           =   2646
         EndProperty
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "months"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3060
         TabIndex        =   6
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "PAST"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   900
         TabIndex        =   4
         Top             =   600
         Width           =   795
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   12600
         Y1              =   420
         Y2              =   420
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "AGING ACCOUNTS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   612
         Left            =   120
         TabIndex        =   3
         Top             =   60
         Width           =   3732
      End
   End
End
Attribute VB_Name = "frmAgingAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdSearch_Click()
Call loadAgingAccounts(lsvCustomers, Val(cboMonths.Text))
End Sub

Private Sub Form_Load()
cboMonths.AddItem "2"
cboMonths.AddItem "4"
cboMonths.AddItem "6"
cboMonths.AddItem "12"

End Sub

Private Sub lsvCustomers_DblClick()

If lsvCustomers.ListItems.Count > 0 Then
    customer_id_for_list_of_account_receivable = Val(lsvCustomers.SelectedItem.Text)
    frmCustomerAccountReceivable.Show 1
End If
End Sub
