VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmManageItemRebates 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "View Rebates"
   ClientHeight    =   8715
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9660
   Icon            =   "frmManageItemRebates.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   9660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   8595
      Left            =   60
      ScaleHeight     =   8565
      ScaleWidth      =   9525
      TabIndex        =   0
      Top             =   60
      Width           =   9555
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
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
         Left            =   6840
         TabIndex        =   7
         Top             =   7800
         Width           =   2532
      End
      Begin VB.ComboBox cboCategory 
         Height          =   315
         Left            =   6840
         TabIndex        =   4
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
         Left            =   4200
         TabIndex        =   1
         Top             =   7800
         Width           =   2532
      End
      Begin MSComctlLib.ListView lsvItemList 
         Height          =   6435
         Left            =   180
         TabIndex        =   6
         Top             =   1260
         Width           =   9195
         _ExtentX        =   16219
         _ExtentY        =   11351
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
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
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "select all"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   180
         TabIndex        =   8
         Tag             =   "0"
         Top             =   960
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item Catgory"
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
         Left            =   5340
         TabIndex        =   5
         Top             =   780
         Width           =   1275
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "List of Item's has rebate"
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
      Begin VB.Line Line1 
         X1              =   120
         X2              =   10200
         Y1              =   420
         Y2              =   420
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
         TabIndex        =   2
         Top             =   7320
         Width           =   45
      End
   End
End
Attribute VB_Name = "frmManageItemRebates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboCategory_Click()
Dim icat As String
    icat = cboCategory.Text
    loadItemsByCategory icat, lsvItemList
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdUpdate_Click()
    Dim list As ListItem
    
    For Each list In lsvItemList.ListItems
        Call updateItemsRebate(Val(list.Text), list.Checked)
    Next
    MsgBox "Successfully updated", vbInformation, "updated"
End Sub

Private Sub Form_Load()

    Call setItemsDescriptionColumns(lsvItemList)
    lsvItemList.ColumnHeaders(1).width = 250
    lsvItemList.ColumnHeaders(1).Text = ""
    lsvItemList.ColumnHeaders(2).width = 3500
    lsvItemList.ColumnHeaders(3).width = 4500
    lsvItemList.ColumnHeaders(4).width = 0
    lsvItemList.ColumnHeaders(5).Alignment = lvwColumnRight
    lsvItemList.ColumnHeaders(6).Alignment = lvwColumnRight
    lsvItemList.ColumnHeaders(5).width = 0
    lsvItemList.ColumnHeaders(6).width = 0
    lsvItemList.ColumnHeaders(7).width = 0
    lsvItemList.ColumnHeaders(8).width = 0
    lsvItemList.ColumnHeaders(9).width = 0
    
    Call loadAllItemsToListviewForRebates(lsvItemList, "item_code")
    
    Call load_to_category_combo(cboCategory)
   
End Sub

Private Sub lbl_Click()
Dim list As ListItem
Dim action As Boolean

If lbl.Tag = 0 Then
    lbl.Caption = "Unselect All"
    lbl.Tag = 1
    action = True
Else
    lbl.Caption = "Select All"
    lbl.Tag = 0
    action = False
End If

For Each list In lsvItemList.ListItems
    list.Checked = action
Next
End Sub
