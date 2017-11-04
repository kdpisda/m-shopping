VERSION 5.00
Begin VB.Form UserPanelForm 
   Caption         =   "Select Mobiles"
   ClientHeight    =   3780
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9555
   LinkTopic       =   "Form1"
   ScaleHeight     =   3780
   ScaleWidth      =   9555
   StartUpPosition =   3  'Windows Default
   Begin VB.Label WelcomeLabel 
      Caption         =   "Welcome, "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label UserNameLabel 
      Caption         =   "UserName"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "UserPanelForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim username As String

Private Sub Form_Load()
    username = WelcomeForm.UsernameTextBox
    UserNameLabel.Caption = username
End Sub
