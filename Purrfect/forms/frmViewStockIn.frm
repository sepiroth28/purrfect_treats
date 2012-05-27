VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmViewStockIn 
   Appearance      =   0  'Flat
   BackColor       =   &H00C8761C&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stock In Records"
   ClientHeight    =   9360
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15135
   Icon            =   "frmViewStockIn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9360
   ScaleWidth      =   15135
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H80000008&
      Height          =   9255
      Left            =   60
      ScaleHeight     =   9225
      ScaleWidth      =   14985
      TabIndex        =   0
      Top             =   60
      Width           =   15015
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Refresh"
         Height          =   435
         Left            =   5700
         TabIndex        =   8
         Top             =   780
         Width           =   1335
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
         Height          =   675
         Left            =   10500
         TabIndex        =   7
         Top             =   660
         Width           =   1935
      End
      Begin MSComctlLib.ListView lsvStockIn 
         Height          =   7635
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   14835
         _ExtentX        =   26167
         _ExtentY        =   13467
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
      Begin VB.CommandButton cmdLoadRecords 
         Caption         =   "LOAD RECORDS"
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
         Left            =   12540
         TabIndex        =   5
         Top             =   660
         Width           =   2355
      End
      Begin VB.CommandButton cmdCalendar 
         Caption         =   "..."
         Height          =   435
         Left            =   4920
         TabIndex        =   4
         Top             =   780
         Width           =   495
      End
      Begin VB.TextBox txtDate 
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
         Left            =   960
         TabIndex        =   3
         Top             =   780
         Width           =   3915
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   300
         TabIndex        =   2
         Top             =   840
         Width           =   555
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "VIEW STOCK IN RECORDS"
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
         Width           =   2835
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   180
         X2              =   14880
         Y1              =   540
         Y2              =   540
      End
   End
End
Attribute VB_Name = "frmViewStockIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCalendar_Click()
Set activeDateTextbox = txtDate
frmCalendar.Show 1
End Sub

Private Sub cmdLoadRecords_Click()
    activeReprintStockIN = 2
    Call loadStockInListByDate(Format(activeDate, "yyyy-mm-dd"), lsvStockIn)
End Sub

Private Sub cmdPrint_Click()
    Dim stk As New StockIn
    Dim iReprintStockIN As Integer
      iReprintStockIN = activeReprintStockIN
      stk.print_stock_in (iReprintStockIN)
End Sub

Private Sub cmdRefresh_Click()
    
    Call ClearFields(Me)
    activeReprintStockIN = 1
    Call loadAlStockInList(lsvStockIn)
End Sub

Private Sub Form_Load()

    Call setViewStockInListview(lsvStockIn)
    With lsvStockIn
        .ColumnHeaders(1).width = 0
        .ColumnHeaders(2).width = 2000
        .ColumnHeaders(3).width = 4000
        .ColumnHeaders(4).width = 3000
        .ColumnHeaders(5).width = 3000
    '    .ColumnHeaders(6).width = 2500
    End With
    
    Call loadStockInListByDate(Format(activeDate, "yyyy-mm-dd"), lsvStockIn)

End Sub

Private Sub lsvStockIn_DblClick()
activestockId = Val(lsvStockIn.SelectedItem.Text)
frmStockInItem.Show 1
End Sub

Private Sub txtDate_Click()
Set activeDateTextbox = txtDate
frmCalendar.Show 1
End Sub
