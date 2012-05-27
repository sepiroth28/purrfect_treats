VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   6795
   ClientLeft      =   3930
   ClientTop       =   2040
   ClientWidth     =   8640
   LinkTopic       =   "Form2"
   ScaleHeight     =   6795
   ScaleWidth      =   8640
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Tag             =   "*"
      Text            =   "Text3"
      Top             =   3360
      Width           =   4455
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   2400
      Width           =   4815
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Tag             =   "*"
      Text            =   "Text1"
      Top             =   1560
      Width           =   4095
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Call validateRequiredValueInForm(Me)
frmlstCustomer.Show
End Sub
