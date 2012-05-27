VERSION 5.00
Begin VB.Form frmDiscount 
   Appearance      =   0  'Flat
   BackColor       =   &H00C8761C&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Discount"
   ClientHeight    =   4995
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   5775
   Icon            =   "frmDiscount.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   5775
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H80000008&
      Height          =   4875
      Left            =   60
      ScaleHeight     =   4845
      ScaleWidth      =   5625
      TabIndex        =   0
      Top             =   60
      Width           =   5655
      Begin VB.TextBox txtDiscount_Amount 
         Alignment       =   1  'Right Justify
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
      Begin VB.TextBox txtDiscount_Code 
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
         Left            =   2820
         TabIndex        =   4
         Top             =   4020
         Width           =   2655
      End
      Begin VB.TextBox txtDiscount_Name 
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
         TabIndex        =   8
         Top             =   120
         Width           =   5175
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Discount Name"
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
         TabIndex        =   7
         Top             =   1860
         Width           =   1470
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Discount Code"
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
         Top             =   840
         Width           =   1410
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Discount Amount"
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
         TabIndex        =   5
         Top             =   2820
         Width           =   1635
      End
   End
End
Attribute VB_Name = "frmDiscount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim discounted As New Customers_Discount
Private Sub cmdAddNewItem_Click()
    saveData
    Call loadAllDiscountToListview(frmManageDiscount.lsvDiscount)
    Unload Me
End Sub

Sub saveData()
    Dim mvardiscounted_id, mvardiscount_code, mvardiscount_name As String
    Dim mvardiscount_amount As Double
    
    With discounted
    
        .discount_code = txtDiscount_Code.Text
        .discount_name = txtDiscount_Name.Text
        .discount_amount = Val(txtDiscount_Amount.Text)
    
    
        If editmode = True Then
            
          .discount_id = activeDiscout_id
          txtDiscount_Code.Text = .discount_code
          txtDiscount_Name.Text = .discount_name
          txtDiscount_Amount.Text = .discount_amount
          .Update_Discount
        Else
        
            .Save_Discount
        End If
        
    End With
   
    

End Sub

Private Sub Form_Load()
        With discounted
            If editmode = True Then
               Call .load_discount(activeDiscout_id)
              txtDiscount_Code.Text = .discount_code
              txtDiscount_Name.Text = .discount_name
              txtDiscount_Amount.Text = .discount_amount
            End If
        End With
        
End Sub
