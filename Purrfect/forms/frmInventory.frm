VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInventory 
   BackColor       =   &H00C7FEF3&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventory"
   ClientHeight    =   9270
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11955
   Icon            =   "frmInventory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9270
   ScaleWidth      =   11955
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H80000008&
      Height          =   9135
      Left            =   60
      ScaleHeight     =   9105
      ScaleWidth      =   11805
      TabIndex        =   0
      Top             =   60
      Width           =   11835
      Begin VB.ComboBox cboCategory 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "frmInventory.frx":058A
         Left            =   9540
         List            =   "frmInventory.frx":058C
         TabIndex        =   8
         Text            =   "cboCategory"
         Top             =   180
         Width           =   2112
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   5160
         TabIndex        =   7
         Top             =   8220
         Width           =   1992
      End
      Begin VB.ComboBox cboInventoryDate 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         ItemData        =   "frmInventory.frx":058E
         Left            =   2280
         List            =   "frmInventory.frx":0590
         TabIndex        =   6
         Text            =   "Today"
         Top             =   8520
         Visible         =   0   'False
         Width           =   2532
      End
      Begin VB.CommandButton cmdCreateEndingBalance 
         Caption         =   "CREATE ENDING BALANCE and PRINT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   7260
         TabIndex        =   4
         Top             =   8220
         Width           =   4392
      End
      Begin MSComctlLib.ListView lsvInventory 
         Height          =   7455
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   13150
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   128
         BackColor       =   16777215
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
            Text            =   "Item Code"
            Object.Width           =   10583
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Beginning Balance"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Ending Balance"
            Object.Width           =   4410
         EndProperty
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Select Category :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   432
         Left            =   7980
         TabIndex        =   9
         Top             =   180
         Width           =   1572
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Inventory Date: "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   228
         Left            =   120
         TabIndex        =   5
         Top             =   8520
         Visible         =   0   'False
         Width           =   2028
      End
      Begin VB.Label lblDate 
         BackStyle       =   0  'Transparent
         Caption         =   "Ending Balance as of"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   435
         Left            =   3180
         TabIndex        =   3
         Top             =   120
         Width           =   8475
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ending Balance as of"
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
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   3075
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   120
         X2              =   13800
         Y1              =   600
         Y2              =   600
      End
   End
End
Attribute VB_Name = "frmInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboCategory_Click()
 Dim icat As String
    
    icat = cboCategory.Text

   Call loadCategoryToListview(icat, lsvInventory)
End Sub

Private Sub cmdCreateEndingBalance_Click()
Dim i As New Inventory
Call i.setNewEndingBalance
MsgBox "Successfully save Ending Balance...", vbInformation
'i.as_of = cboInventoryDate.Text
'Set dtaInventory.DataSource = i.getTodaysInventoryForPrinting()
'dtaInventory.Show 1
End Sub

Private Sub cmdPrint_Click()
Dim i As New Inventory

If cboInventoryDate.Text = "Today" Then
    Set dtaInventory.DataSource = i.getTodaysInventoryForPrinting(cboCategory.Text)
    dtaInventory.Sections(1).Controls("lblNutrimart").Caption = "Nutrimart Enterprises" & vbCrLf & "Calape, Bohol" & vbCrLf & "255-7304"
    dtaInventory.Sections(1).Controls("lblDate").Caption = FormatDateTime(Date, vbLongDate)
    dtaInventory.Sections(1).Controls("lblItemCategory").Caption = cboCategory.Text
    dtaInventory.Show 1
Else
    i.as_of = Format(cboInventoryDate.Text, "yyyy-mm-dd H:M:S")
    Set dtaInventory.DataSource = i.getTodaysInventoryForPrinting(cboCategory)
    dtaInventory.Sections(1).Controls("lblNutrimart").Caption = "Nutrimart Enterprises" & vbCrLf & "Calape, Bohol" & vbCrLf & "255-7304"
    dtaInventory.Sections(1).Controls("lblDate").Caption = FormatDateTime(Date, vbLongDate)
    dtaInventory.Sections(1).Controls("lblItemCategory").Caption = cboCategory.Text
    dtaInventory.Show 1
End If
End Sub

Private Sub Form_Load()
    lblDate.Caption = FormatDateTime(Date, vbLongDate)
    Call loadTodaysInventoryToListView(lsvInventory)
    Call loadInventoryDateSelection(cboInventoryDate)
    Call load_to_category_combo(cboCategory)
    
End Sub
