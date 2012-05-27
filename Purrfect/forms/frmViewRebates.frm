VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmViewRebates 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "View Rebates"
   ClientHeight    =   8760
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13155
   Icon            =   "frmViewRebates.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   13155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   8655
      Left            =   60
      ScaleHeight     =   8625
      ScaleWidth      =   13005
      TabIndex        =   0
      Top             =   60
      Width           =   13035
      Begin VB.PictureBox picIssueRebate 
         Appearance      =   0  'Flat
         BackColor       =   &H00C7FEF3&
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   180
         ScaleHeight     =   825
         ScaleWidth      =   6705
         TabIndex        =   8
         Top             =   7680
         Width           =   6735
         Begin VB.CommandButton cmdIssueRebate 
            Caption         =   "DONE ISSUE REBATE"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   675
            Left            =   3960
            TabIndex        =   10
            Top             =   60
            Width           =   2655
         End
         Begin VB.TextBox txtIssueBy 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1260
            TabIndex        =   9
            Top             =   120
            Width           =   2475
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Issue by:"
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
            TabIndex        =   11
            Top             =   120
            Width           =   855
         End
      End
      Begin VB.ComboBox cboMonth 
         Height          =   315
         Left            =   3900
         TabIndex        =   5
         Top             =   780
         Width           =   2595
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
         Left            =   10320
         TabIndex        =   1
         Top             =   7800
         Width           =   2532
      End
      Begin MSComctlLib.ListView lsvItemList 
         Height          =   5775
         Left            =   180
         TabIndex        =   2
         Top             =   1260
         Width           =   12675
         _ExtentX        =   22357
         _ExtentY        =   10186
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "stock_out_id"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "item_id"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Item Name"
            Object.Width           =   7937
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Total bought"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Unit"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "rebate price"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "total rebate amount"
            Object.Width           =   3881
         EndProperty
      End
      Begin MSComctlLib.ListView lsvTotal 
         Height          =   375
         Left            =   180
         TabIndex        =   12
         Top             =   7200
         Width           =   12675
         _ExtentX        =   22357
         _ExtentY        =   661
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "stock_out_id"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "item_id"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Item Name"
            Object.Width           =   7937
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Total bought"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Unit"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "rebate price"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "total rebate amount"
            Object.Width           =   3881
         EndProperty
      End
      Begin VB.Label lblCustomerName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "View Customer Rebates"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   10050
         TabIndex        =   14
         Top             =   120
         Width           =   2745
      End
      Begin VB.Label lblDoneRemit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DONE ISSUE REBATE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   240
         Left            =   360
         TabIndex        =   13
         Top             =   7920
         Width           =   1860
      End
      Begin VB.Label lblYear 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "year "
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
         Left            =   6600
         TabIndex        =   7
         Top             =   780
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "List all items purchased on Month of "
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
         Left            =   240
         TabIndex        =   6
         Top             =   780
         Width           =   3555
      End
      Begin VB.Label lblPreparedBy 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   1800
         TabIndex        =   4
         Top             =   7320
         Width           =   45
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   13020
         Y1              =   540
         Y2              =   540
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "View Customer Rebates"
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
Attribute VB_Name = "frmViewRebates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m As String

Private Sub cboMonth_Click()
Call loadItemsQualifiedForRebatesByCustomer(activeCustomerIdForRebate, cboMonth.Text, lsvItemList)
Call renderRebateTableRates(lsvItemList)

lsvTotal.ListItems.Clear
Set list = lsvTotal.ListItems.Add(, , "")
    list.SubItems(6) = rebate_grand_total
    list.SubItems(3) = rebate_grand_total_qty

prepareRebateButton cboMonth.Text
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdIssueRebate_Click()
If txtIssueBy.Text <> "" Then
    Call issueRebate(activeCustomerIdForRebate, cboMonth.Text, rebate_grand_total, rebate_grand_total_qty, txtIssueBy.Text)
MsgBox "Issue rebate successfully...", vbOKOnly, "Rebate"
prepareRebateButton cboMonth.Text
End If


End Sub

Private Sub Form_Load()
Dim list As ListItem
Dim x As Integer
Dim cus As New Customers
cus.load_customers (activeCustomerIdForRebate)

lblCustomerName.Caption = cus.customers_name

For x = 1 To 12
    cboMonth.AddItem MonthName(x)
Next x
lblYear.Caption = "year " & Year(Now)



m = MonthName(Format(Date, "m") - 1)

Call loadItemsQualifiedForRebatesByCustomer(activeCustomerIdForRebate, m, lsvItemList)
Call renderRebateTableRates(lsvItemList)
cboMonth.Text = m

lsvTotal.ListItems.Clear
Set list = lsvTotal.ListItems.Add(, , "")
list.SubItems(6) = rebate_grand_total
list.SubItems(3) = rebate_grand_total_qty

prepareRebateButton m


Set cus = Nothing
End Sub

Sub prepareRebateButton(mo As String)

If isDoneIssueRebate(activeCustomerIdForRebate, mo) Then
    picIssueRebate.Visible = False
Else
    picIssueRebate.Visible = True
End If

picIssueRebate.Visible = activeUser.previliges.can_issue_rebate
lblDoneRemit.Visible = activeUser.previliges.can_issue_rebate
End Sub

