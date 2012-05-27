VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCustomerVisibility 
   BackColor       =   &H00000040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer visibility"
   ClientHeight    =   8100
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10545
   Icon            =   "frmCustomerVisibility.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   10545
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   7155
      Left            =   120
      TabIndex        =   2
      Top             =   60
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   12621
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Visible Customer"
      TabPicture(0)   =   "frmCustomerVisibility.frx":058A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lsvVisibleCustomer"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Hidden Customer"
      TabPicture(1)   =   "frmCustomerVisibility.frx":05A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lsvHiddenCustomer"
      Tab(1).ControlCount=   1
      Begin MSComctlLib.ListView lsvVisibleCustomer 
         Height          =   6735
         Left            =   60
         TabIndex        =   3
         Top             =   360
         Width           =   10155
         _ExtentX        =   17912
         _ExtentY        =   11880
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Width           =   5292
         EndProperty
      End
      Begin MSComctlLib.ListView lsvHiddenCustomer 
         Height          =   6735
         Left            =   -74940
         TabIndex        =   4
         Top             =   360
         Width           =   10155
         _ExtentX        =   17912
         _ExtentY        =   11880
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Width           =   5292
         EndProperty
      End
   End
   Begin VB.TextBox txtSearchCustomer 
      Appearance      =   0  'Flat
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
      TabIndex        =   1
      Top             =   7380
      Width           =   5955
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "HIDE SELECTED CUSTOMER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   6360
      TabIndex        =   0
      Top             =   7320
      Width           =   4095
   End
End
Attribute VB_Name = "frmCustomerVisibility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim activeCustomerLsv As ListView
Private Sub cmdSelect_Click()

End Sub

Private Sub cmdApply_Click()
If activeCustomerLsv.ListItems.Count < 1 Then
    Exit Sub
End If

Dim list As ListItem
Dim action As Integer

If cmdApply.Tag = "hide" Then
    action = 0
Else
    action = 1
End If

For Each list In activeCustomerLsv.ListItems
    If list.Checked = True Then
        'to do
        db.execute "UPDATE customers SET visible = " & action & " WHERE customers_id = " & Val(list.Text)

    End If
Next
Call loadAllCustomersToListview(lsvVisibleCustomer)
Call loadAllCustomersToListviewHidden(lsvHiddenCustomer)
End Sub

Private Sub Form_Load()
Call loadAllCustomersToListview(lsvVisibleCustomer)
Call loadAllCustomersToListviewHidden(lsvHiddenCustomer)
Set activeCustomerLsv = lsvVisibleCustomer
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

If SSTab1.TabCaption(PreviousTab) = "Visible Customer" Then
    cmdApply.Caption = "SHOW SELECTED CUSTOMER"
    Set activeCustomerLsv = lsvHiddenCustomer
    cmdApply.Tag = "show"
ElseIf SSTab1.TabCaption(PreviousTab) = "Hidden Customer" Then
    cmdApply.Caption = "HIDE SELECTED CUSTOMER"
    Set activeCustomerLsv = lsvVisibleCustomer
    cmdApply.Tag = "hide"
End If
End Sub

Private Sub txtSearchCustomer_Change()
Dim rs As New ADODB.Recordset
Set rs = searchCustomersByName(txtSearchCustomer.Text)
Call loadCustomerRSToListView(activeCustomerLsv, rs)
End Sub
