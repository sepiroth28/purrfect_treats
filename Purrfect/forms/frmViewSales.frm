VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmViewSales 
   Appearance      =   0  'Flat
   BackColor       =   &H00C8761C&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Sales"
   ClientHeight    =   8790
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   14970
   Icon            =   "frmViewSales.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   14970
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   8655
      Left            =   60
      ScaleHeight     =   8625
      ScaleWidth      =   14805
      TabIndex        =   0
      Top             =   60
      Width           =   14835
      Begin VB.CommandButton cmdCODRemit 
         Caption         =   "Remit COD"
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
         Left            =   12060
         TabIndex        =   16
         Top             =   7200
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.CommandButton cmdLoadRecords 
         Caption         =   "Load Records"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   12600
         TabIndex        =   15
         Top             =   60
         Width           =   2115
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "PRINT"
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
         Left            =   12060
         TabIndex        =   7
         Top             =   7860
         Width           =   2655
      End
      Begin VB.CommandButton cmdShowDate 
         BackColor       =   &H00C7FEF3&
         Caption         =   "..."
         Height          =   375
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   555
      End
      Begin VB.TextBox txtSalesDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   5
         Top             =   120
         Width           =   3555
      End
      Begin VB.ComboBox cboPaymentType 
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
         ItemData        =   "frmViewSales.frx":058A
         Left            =   1680
         List            =   "frmViewSales.frx":0597
         TabIndex        =   4
         Text            =   "ALL"
         Top             =   7980
         Width           =   3135
      End
      Begin MSComctlLib.ListView lsvSales 
         Height          =   5475
         Left            =   180
         TabIndex        =   2
         Top             =   720
         Width           =   14535
         _ExtentX        =   25638
         _ExtentY        =   9657
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label lblRemitted 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "COD Amount is already remitted"
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
         Height          =   375
         Left            =   10740
         TabIndex        =   17
         Top             =   7380
         Visible         =   0   'False
         Width           =   3795
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "click for details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00D0B328&
         Height          =   240
         Left            =   9720
         TabIndex        =   14
         Top             =   6720
         Width           =   1590
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Payments Received :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   6720
         Width           =   4455
      End
      Begin VB.Label lblPaymentReceivedtotal 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   8040
         TabIndex        =   12
         Top             =   6720
         Width           =   1455
      End
      Begin VB.Label lblNetTotal 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   8280
         TabIndex        =   11
         Top             =   6360
         Width           =   1215
      End
      Begin VB.Label lblGrandTotal 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   6480
         TabIndex        =   10
         Top             =   6360
         Width           =   1455
      End
      Begin VB.Label lblTotalDiscount 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   4860
         TabIndex        =   9
         Top             =   6360
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Totals:  "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   8
         Top             =   6360
         Width           =   1455
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         X1              =   240
         X2              =   14700
         Y1              =   7740
         Y2              =   7740
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Payment type: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   180
         TabIndex        =   3
         Top             =   8040
         Width           =   1410
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   180
         X2              =   14700
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "SALES as of "
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
         Left            =   180
         TabIndex        =   1
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuSalesOrderHistory 
         Caption         =   "Sales Order History"
      End
   End
End
Attribute VB_Name = "frmViewSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboPaymentType_Change()

If cboPaymentType.Text = "COD" Then
    Call loadAllSalesToListview(lsvSales, True, PAYMENT_COD)
    Call updateTotals(activeDate)
   
ElseIf cboPaymentType.Text = "ACCOUNT RECEIVABLE" Then
    Call loadAllSalesToListview(lsvSales, True, PAYMENT_ACCOUNT_RECEIVABLE)
    Call updateTotals(activeDate)
Else
    Call loadAllSalesToListview(lsvSales, True, 3)
    Call updateTotals(activeDate)
End If
End Sub

Private Sub cboPaymentType_Click()
cmdCODRemit.Visible = False
If cboPaymentType.Text = "COD" Then
    Call loadAllSalesToListview(lsvSales, True, PAYMENT_COD)
    Call updateTotals(activeDate)
    If activeUser.previliges.can_accept_remit_payments Then
        cmdCODRemit.Visible = True
    End If
ElseIf cboPaymentType.Text = "ACCOUNT RECEIVABLE" Then
    Call loadAllSalesToListview(lsvSales, True, PAYMENT_ACCOUNT_RECEIVABLE)
    Call updateTotals(activeDate)
Else
    Call loadAllSalesToListview(lsvSales, True, 3)
    Call updateTotals(activeDate)
End If
End Sub

Private Sub cmdCODRemit_Click()
Dim insert As String
Dim remit_by As String
remit_by = InputBox("Please input name who remitted the cod sales", "Remit by")
'id, sales_date, remit_by, received_by, date_accepted, amount
insert = "INSERT INTO cod_remitted VALUES(null,'" & Format(activeDate, "yyyy-mm-dd") & "','" & remit_by & "','" & activeUser.username & "',CURDATE()," & getNetTotalAsOfTodaySales(PAYMENT_COD) & ")"
If MsgBox("Are you sure you want to save this data?", vbInformation + vbYesNo, "Remit COD") = vbYes Then
    db.execute insert
End If

Dim is_cod_remitted As Boolean
is_cod_remitted = checkCODIfRemitted(activeDate)
lblRemitted.Visible = is_cod_remitted
cmdCODRemit.Visible = Not is_cod_remitted
End Sub

Private Sub cmdLoadRecords_Click()


'If cboPaymentType.Text = "COD" Then
    Call loadAllSalesToListview(lsvSales, False, 3, Format(activeDate, "yyyy-mm-dd"))
'End If
updateTotals (activeDate)


    Dim is_cod_remitted As Boolean
    is_cod_remitted = checkCODIfRemitted(activeDate)
    lblRemitted.Visible = is_cod_remitted
    If cboPaymentType.Text = "COD" Then
        If activeUser.previliges.can_accept_remit_payments Then
            cmdCODRemit.Visible = Not is_cod_remitted
        End If
    End If
End Sub

Private Sub cmdPrint_Click()
    Dim s As New Sales
        s.printSalesReport (cboPaymentType.Text)
End Sub

Private Sub cmdShowDate_Click()
Set activeDateTextbox = txtSalesDate
frmCalendar.Show 1
End Sub

Private Sub Form_Load()

Call setSalesListview(lsvSales)
lsvSales.ColumnHeaders(2).width = 1800
lsvSales.ColumnHeaders(6).width = 0
lsvSales.ColumnHeaders(7).width = 1800
lsvSales.ColumnHeaders(8).width = 1000
lsvSales.ColumnHeaders(9).width = 2200

lsvSales.ColumnHeaders(4).Alignment = lvwColumnRight
lsvSales.ColumnHeaders(5).Alignment = lvwColumnRight
lsvSales.ColumnHeaders(6).Alignment = lvwColumnRight
lsvSales.ColumnHeaders(7).Alignment = lvwColumnRight
lsvSales.ColumnHeaders(8).Alignment = lvwColumnRight

Call loadAllSalesToListview(lsvSales, True, 3)
txtSalesDate.Text = FormatDateTime(Date, vbLongDate)
updateTotals (Date)



End Sub

Sub updateTotals(date_to_used As Date)
    If cboPaymentType.Text = "ALL" Then
        lblTotalDiscount.Caption = FormatNumber(getTotalDiscountAsOfTodaySales(3), 3)
        lblGrandTotal.Caption = FormatNumber(getGrandTotalAsOfTodaySales(3), 2)
        lblNetTotal.Caption = FormatNumber(getNetTotalAsOfTodaySales(3), 2)
        lblPaymentReceivedtotal.Caption = FormatNumber(getTotalPaymentReceiveToday(date_to_used))
    ElseIf cboPaymentType.Text = "COD" Then
        lblTotalDiscount.Caption = FormatNumber(getTotalDiscountAsOfTodaySales(PAYMENT_COD), 3)
        lblGrandTotal.Caption = FormatNumber(getGrandTotalAsOfTodaySales(PAYMENT_COD), 2)
        lblNetTotal.Caption = FormatNumber(getNetTotalAsOfTodaySales(PAYMENT_COD), 2)
        lblPaymentReceivedtotal.Caption = FormatNumber(getTotalPaymentReceiveToday(date_to_used))
    ElseIf cboPaymentType.Text = "ACCOUNT RECEIVABLE" Then
        lblTotalDiscount.Caption = FormatNumber(getTotalDiscountAsOfTodaySales(PAYMENT_ACCOUNT_RECEIVABLE), 3)
        lblGrandTotal.Caption = FormatNumber(getGrandTotalAsOfTodaySales(PAYMENT_ACCOUNT_RECEIVABLE), 2)
        lblNetTotal.Caption = FormatNumber(getNetTotalAsOfTodaySales(PAYMENT_ACCOUNT_RECEIVABLE), 2)
        lblPaymentReceivedtotal.Caption = FormatNumber(getTotalPaymentReceiveToday(date_to_used))
    End If
End Sub

Private Sub Label6_Click()
frmPaymentReceivedDetails.Show 1
End Sub

Private Sub lsvSales_DblClick()
    On Error Resume Next
    activeSalesOrderForViewSalesDetails = lsvSales.SelectedItem.Text
    frmViewSalesInDetails.Show 1
End Sub

Private Sub lsvSales_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    PopupMenu mnuFile
End If
End Sub

Private Sub mnuSalesOrderHistory_Click()
If lsvSales.ListItems.Count > 0 Then
    selectedSOForHistory = lsvSales.SelectedItem.Text
    frmSOHistory.Show 1
End If
End Sub

Private Sub txtSalesDate_Click()
Set activeDateTextbox = txtSalesDate
frmCalendar.Show 1
End Sub
