VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStockInPreview 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stock In Preview"
   ClientHeight    =   10125
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12405
   Icon            =   "frmStockInPreview.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10125
   ScaleWidth      =   12405
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9975
      Left            =   60
      ScaleHeight     =   9945
      ScaleWidth      =   12225
      TabIndex        =   0
      Top             =   60
      Width           =   12255
      Begin MSComctlLib.ListView lsvManufacturers 
         Height          =   975
         Left            =   1980
         TabIndex        =   27
         Top             =   2940
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
      Begin VB.CommandButton cmdShowSupplier 
         Caption         =   "..."
         Height          =   435
         Left            =   7260
         TabIndex        =   26
         Top             =   2460
         Width           =   675
      End
      Begin VB.TextBox txtRemarks 
         Appearance      =   0  'Flat
         BackColor       =   &H00C7FEF3&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   1980
         TabIndex        =   20
         Top             =   3120
         Width           =   9975
      End
      Begin VB.TextBox txtSupplier 
         Appearance      =   0  'Flat
         BackColor       =   &H00C7FEF3&
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
         Left            =   1980
         TabIndex        =   19
         Top             =   2460
         Width           =   5235
      End
      Begin VB.TextBox txtStockedInTo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C7FEF3&
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
         Left            =   1980
         TabIndex        =   18
         Text            =   "WH-02 STOCKROOM(BODEGA)"
         Top             =   1920
         Width           =   5235
      End
      Begin VB.TextBox txtReferenceNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C7FEF3&
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
         Left            =   1980
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   1380
         Width           =   2055
      End
      Begin VB.CommandButton cmdProcessStockedIn 
         Caption         =   "PROCESS STOCK IN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   8820
         TabIndex        =   12
         Top             =   9060
         Width           =   3195
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   915
         Left            =   300
         ScaleHeight     =   885
         ScaleWidth      =   11625
         TabIndex        =   9
         Top             =   7800
         Width           =   11655
         Begin VB.TextBox txtReceivedBy 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   8460
            TabIndex        =   16
            Top             =   420
            Width           =   2835
         End
         Begin VB.TextBox txtApproveBy 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4500
            TabIndex        =   15
            Top             =   420
            Width           =   2835
         End
         Begin VB.TextBox txtPreparedBy 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   360
            Locked          =   -1  'True
            TabIndex        =   13
            Top             =   420
            Width           =   2835
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Prepared By:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   360
            TabIndex        =   14
            Top             =   60
            Width           =   1110
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Received By:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   8460
            TabIndex        =   11
            Top             =   60
            Width           =   1155
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Approved By:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4500
            TabIndex        =   10
            Top             =   60
            Width           =   1155
         End
      End
      Begin MSComctlLib.ListView lsvStockInItems 
         Height          =   3375
         Left            =   300
         TabIndex        =   6
         Top             =   3780
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   5953
         View            =   3
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
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "50.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   10455
         TabIndex        =   25
         Top             =   7200
         Width           =   690
      End
      Begin VB.Label lblTotalItems 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(2)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3780
         TabIndex        =   24
         Top             =   7260
         Width           =   285
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. of Items :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2040
         TabIndex        =   23
         Top             =   7260
         Width           =   1365
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   300
         TabIndex        =   22
         Top             =   7260
         Width           =   750
      End
      Begin VB.Label lblDate 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DateTime"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   10920
         TabIndex        =   21
         Top             =   1920
         Width           =   1035
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   300
         X2              =   12000
         Y1              =   8940
         Y2              =   8940
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   780
         TabIndex        =   8
         Top             =   3180
         Width           =   1065
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stock In Date and Time : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   9300
         TabIndex        =   7
         Top             =   1500
         Width           =   2610
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From Supplier: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   5
         Top             =   2520
         Width           =   1590
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stocked In To: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   4
         Top             =   1980
         Width           =   1575
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reference No.  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   3
         Top             =   1440
         Width           =   1635
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "STOCK-IN / RETURN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8880
         TabIndex        =   2
         Top             =   180
         Width           =   2985
      End
      Begin VB.Label lblNutrimart 
         BackStyle       =   0  'Transparent
         Caption         =   "Nutrimart Enterprises"
         Height          =   1335
         Left            =   180
         TabIndex        =   1
         Top             =   120
         Width           =   2715
      End
   End
End
Attribute VB_Name = "frmStockInPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdProcessStockedIn_Click()
Dim itemToBeStock As New StockIn
Dim stock_in As New StockIn
Dim rs As New ADODB.Recordset

With activeStockInList
    .stock_in_to = txtStockedInTo.Text
    .supplier_id = Val(lsvManufacturers.SelectedItem.Text)
    .remarks = txtRemarks.Text
    .prepared_by = txtPreparedBy.Text
    .approved_by = txtApproveBy.Text
    .received_by = txtReceivedBy.Text
    .insert
End With

'this performs adding of qty to the items
    For Each stock_in In activeStockInList
        stock_in.insert
        activeStockInList.insertStockInRecordToThisTransaction stock_in.get_last_id()
        stock_in.items.addStock stock_in.QtyToBeAdd
    Next


Set rs = db.execute("SELECT * FROM stock_in")
  
activeStockInList.updateReferenceNo
Set dtaStockIn.DataSource = rs
dtaStockIn.Sections(1).Controls("lblNutrimart").Caption = lblNutrimart.Caption
dtaStockIn.Sections(1).Controls("lblreferenceNo").Caption = txtReferenceNo.Text
dtaStockIn.Sections(1).Controls("lblStocked_In_To").Caption = txtStockedInTo.Text
dtaStockIn.Sections(1).Controls("lblFrom_Supplier").Caption = txtSupplier.Text
dtaStockIn.Sections(1).Controls("lblremarks").Caption = txtRemarks.Text
dtaStockIn.Sections(1).Controls("lblStock_In_DateTime").Caption = lblDate.Caption
 
 Dim item_code, item_description, UM, qty As String
    For Each itemToBeStock In activeStockInList
       item_code = item_code & itemToBeStock.items.item_code & vbCrLf
       item_description = item_description & itemToBeStock.items.item_description & vbCrLf
       UM = UM & itemToBeStock.items.unit_of_measure & vbCrLf
       qty = qty & itemToBeStock.QtyToBeAdd & vbCrLf
       
    Next
    dtaStockIn.Sections(1).Controls("lblItem_Code").Caption = item_code
    dtaStockIn.Sections(1).Controls("lbldescription").Caption = item_description
    dtaStockIn.Sections(1).Controls("lblunit_of_measure").Caption = UM
    dtaStockIn.Sections(1).Controls("lblqty").Caption = qty
    dtaStockIn.Sections(1).Controls("lblNoOfItems").Caption = "(" & activeStockInList.Count & ")"
    dtaStockIn.Sections(1).Controls("lblTotalNoOfItems").Caption = activeStockInList.get_total_items

dtaStockIn.Sections(1).Controls("lblpreparedby").Caption = txtPreparedBy.Text
dtaStockIn.Sections(1).Controls("lblapprovedby").Caption = txtApproveBy.Text
dtaStockIn.Sections(1).Controls("lblreceivedby").Caption = txtReceivedBy.Text

Unload Me
dtaStockIn.Show 1
End Sub

Private Sub cmdShowSupplier_Click()
Call toogleListView(lsvManufacturers)
End Sub

Private Sub Form_Load()
Dim itemToBeStock As New StockIn

Call setStockInPreviewListview(lsvStockInItems)
lsvStockInItems.ColumnHeaders(1).width = 0
lsvStockInItems.ColumnHeaders(2).width = 3000
lsvStockInItems.ColumnHeaders(3).width = 5000
lsvStockInItems.ColumnHeaders(4).width = 1500
lsvStockInItems.ColumnHeaders(5).Alignment = lvwColumnRight

Call setManufacturersColumns(lsvManufacturers)
Call loadAllmanufacturersToListview(lsvManufacturers)
        lsvManufacturers.ColumnHeaders(1).width = 0
        lsvManufacturers.ColumnHeaders(2).width = 5000
        lsvManufacturers.ColumnHeaders(3).width = 0
        lsvManufacturers.ColumnHeaders(4).width = 0
        
lblNutrimart.Caption = "Nutrimart Enterprises" & vbCrLf & "Calape, Bohol" & vbCrLf & "255-7304"
lblDate.Caption = Now
txtReferenceNo.Text = activeStockInList.reference_no


    For Each itemToBeStock In activeStockInList
        Dim lst As ListItem
        Set lst = lsvStockInItems.ListItems.Add(, , itemToBeStock.items.item_id)
        lst.SubItems(1) = itemToBeStock.items.item_code
        lst.SubItems(2) = itemToBeStock.items.item_description
        lst.SubItems(3) = itemToBeStock.items.unit_of_measure
        lst.SubItems(4) = itemToBeStock.QtyToBeAdd
    Next
lblTotal.Caption = activeStockInList.get_total_items
lblTotalItems.Caption = "(" & activeStockInList.Count & ")"
txtPreparedBy.Text = activeUser.username
End Sub

Private Sub lsvManufacturers_Click()
txtSupplier.Text = lsvManufacturers.SelectedItem.SubItems(1)
Call toogleListView(lsvManufacturers)
End Sub
