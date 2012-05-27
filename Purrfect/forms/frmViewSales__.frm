VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmViewSales 
   Appearance      =   0  'Flat
   BackColor       =   &H00C8761C&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Sales"
   ClientHeight    =   8796
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   14964
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8796
   ScaleWidth      =   14964
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   8655
      Left            =   60
      ScaleHeight     =   8628
      ScaleWidth      =   14808
      TabIndex        =   0
      Top             =   60
      Width           =   14835
      Begin VB.CommandButton cmdPrint 
         Caption         =   "PRINT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   10.8
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
         Left            =   5220
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox txtSalesDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1620
         TabIndex        =   5
         Top             =   120
         Width           =   3555
      End
      Begin VB.ComboBox cboPaymentType 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmViewSales.frx":0000
         Left            =   1680
         List            =   "frmViewSales.frx":000D
         TabIndex        =   4
         Text            =   "ALL"
         Top             =   7980
         Width           =   3135
      End
      Begin MSComctlLib.ListView lsvSales 
         Height          =   6435
         Left            =   180
         TabIndex        =   2
         Top             =   720
         Width           =   14535
         _ExtentX        =   25633
         _ExtentY        =   11345
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
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label lblNetTotal 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.6
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
         Top             =   7200
         Width           =   1215
      End
      Begin VB.Label lblGrandTotal 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.6
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
         Top             =   7200
         Width           =   1455
      End
      Begin VB.Label lblTotalDiscount 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.6
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
         Top             =   7200
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Totals:  "
         BeginProperty Font 
            Name            =   "Century Gothic"
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
         Top             =   7200
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
         Y1              =   540
         Y2              =   540
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "SALES as of "
         BeginProperty Font 
            Name            =   "Century Gothic"
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
End
Attribute VB_Name = "frmViewSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboPaymentType_Change()
If cboPaymentType.Text = "COD" Then
    Call loadAllSalesToListview(lsvSales, True, PAYMENT_COD)
    Call updateTotals
ElseIf cboPaymentType.Text = "ACCOUNT RECEIVABLE" Then
    Call loadAllSalesToListview(lsvSales, True, PAYMENT_ACCOUNT_RECEIVABLE)
    Call updateTotals
Else
    Call loadAllSalesToListview(lsvSales, True)
    Call updateTotals
End If
End Sub

Private Sub cboPaymentType_Click()
If cboPaymentType.Text = "COD" Then
    Call loadAllSalesToListview(lsvSales, True, PAYMENT_COD)
    Call updateTotals
ElseIf cboPaymentType.Text = "ACCOUNT RECEIVABLE" Then
    Call loadAllSalesToListview(lsvSales, True, PAYMENT_ACCOUNT_RECEIVABLE)
    Call updateTotals
Else
    Call loadAllSalesToListview(lsvSales, True, 3)
    Call updateTotals
End If
End Sub

Private Sub Form_Load()
Call setSalesListview(lsvSales)
lsvSales.ColumnHeaders(2).width = 1800
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
updateTotals

End Sub

Sub updateTotals()
    If cboPaymentType.Text = "ALL" Then
        lblTotalDiscount.Caption = FormatNumber(getTotalDiscountAsOfTodaySales(3), 3)
        lblGrandTotal.Caption = FormatNumber(getGrandTotalAsOfTodaySales(3), 2)
        lblNetTotal.Caption = FormatNumber(getNetTotalAsOfTodaySales(3), 2)
    ElseIf cboPaymentType.Text = "COD" Then
        lblTotalDiscount.Caption = FormatNumber(getTotalDiscountAsOfTodaySales(PAYMENT_COD), 3)
        lblGrandTotal.Caption = FormatNumber(getGrandTotalAsOfTodaySales(PAYMENT_COD), 2)
        lblNetTotal.Caption = FormatNumber(getNetTotalAsOfTodaySales(PAYMENT_COD), 2)
    ElseIf cboPaymentType.Text = "ACCOUNT RECEIVABLE" Then
        lblTotalDiscount.Caption = FormatNumber(getTotalDiscountAsOfTodaySales(PAYMENT_ACCOUNT_RECEIVABLE), 3)
        lblGrandTotal.Caption = FormatNumber(getGrandTotalAsOfTodaySales(PAYMENT_ACCOUNT_RECEIVABLE), 2)
        lblNetTotal.Caption = FormatNumber(getNetTotalAsOfTodaySales(PAYMENT_ACCOUNT_RECEIVABLE), 2)
    End If
End Sub
