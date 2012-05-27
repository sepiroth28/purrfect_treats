VERSION 5.00
Begin VB.Form frmSalesOrder_Responsible 
   BackColor       =   &H00FF0000&
   Caption         =   "Sales Order Responsible "
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5775
   Icon            =   "frmSalesOrder_Responsible.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   5775
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H80000008&
      Height          =   5655
      Left            =   120
      ScaleHeight     =   5625
      ScaleWidth      =   5505
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      Begin VB.TextBox txtChecked_by 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   2160
         Width           =   5055
      End
      Begin VB.CommandButton cmdAddNewItem 
         Caption         =   "SAVE"
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
         Left            =   2640
         TabIndex        =   5
         Top             =   4800
         Width           =   2655
      End
      Begin VB.TextBox txtPrepared_by 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   1
         Top             =   1200
         Width           =   5055
      End
      Begin VB.TextBox txtPosted_by 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   3120
         Width           =   5055
      End
      Begin VB.TextBox txtDelivered_by 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   4080
         Width           =   5055
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Posted By"
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
         TabIndex        =   10
         Top             =   2820
         Width           =   930
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Prepared By"
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
         TabIndex        =   9
         Top             =   840
         Width           =   1200
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Checked By"
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
         TabIndex        =   8
         Top             =   1860
         Width           =   1185
      End
      Begin VB.Label lblRequiredMsg 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "  Please fill up requireed fields..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   240
         TabIndex        =   7
         Top             =   120
         Width           =   5175
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Delivered By"
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
         Top             =   3720
         Width           =   1230
      End
   End
End
Attribute VB_Name = "frmSalesOrder_Responsible"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAddNewItem_Click()
    saveData
End Sub


Sub saveData()
    Dim sales_order_responsible As New SalesOrder_Responsible
    Dim mvarprepared_by, mvarchecked_by, mvarposted_by, mvardelivered_by As String
    
    mvarposted_by = txtPosted_by.Text
    mvarchecked_by = txtChecked_by.Text
    mvarprepared_by = txtPrepared_by.Text
    mvardelivered_by = txtDelivered_by.Text
    
    With sales_order_responsible
        .prepared_by = mvarprepared_by
        .checked_by = mvarchecked_by
        .posted_by = mvarposted_by
        .delivered_by = mvardelivered_by
        
        If editmode = True Then
            .Update_SalesOrder_Responsible
        Else
            
            
            .Save_SalesOrder_Responsible
            MsgBox "Record has been successfully saved."
        End If
        
    End With
    
End Sub

Private Sub Form_Load()
    Dim sales_rep As New SalesOrder_Responsible
    sales_rep.loadToSalesOrder_Responsible
    
    txtChecked_by.Text = sales_rep.checked_by
    txtDelivered_by.Text = sales_rep.delivered_by
    txtPosted_by.Text = sales_rep.posted_by
    txtPrepared_by = sales_rep.prepared_by
    
    
End Sub
