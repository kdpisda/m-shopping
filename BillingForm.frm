VERSION 5.00
Begin VB.Form BillingForm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Billing"
   ClientHeight    =   8910
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12165
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form2"
   Picture         =   "BillingForm.frx":0000
   ScaleHeight     =   8910
   ScaleWidth      =   12165
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      BackColor       =   &H00404040&
      Caption         =   "ORDER"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8160
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00404040&
      Caption         =   "PREVIOUS"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8160
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "YOUR ORDER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1560
      TabIndex        =   11
      Top             =   3600
      Width           =   2775
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "DELIEVER TO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1560
      TabIndex        =   10
      Top             =   5520
      Width           =   2775
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "RS"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6840
      TabIndex        =   9
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "CONTACT "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2400
      TabIndex        =   8
      Top             =   6480
      Width           =   3015
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "NAME"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      Top             =   6000
      Width           =   3255
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "ADDRESS"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   6960
      Width           =   10335
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   4800
      Width           =   3255
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "NAME"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6840
      TabIndex        =   4
      Top             =   3960
      Width           =   3255
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "MOBILE"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   4200
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "PRICE"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7320
      TabIndex        =   2
      Top             =   4800
      Width           =   2295
   End
End
Attribute VB_Name = "BillingForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
BILLING.Hide
DETAILS.Show
End Sub

Private Sub Command2_Click()
Dim P As Integer
P = MsgBox("THANKYOU FOR ORDERING, HAVE A NICE DAY", vbOKOnly)
If P = vbOK Then
    Unload Me
End If


End Sub

Private Sub Command3_Click()
Dim P As Integer
P = MsgBox("THANKYOU FOR ORDERING, HAVE A NICE DAY", vbOKOnly)
If P = vbOK Then
    UserPanelForm.Show
    Unload Me
End If

End Sub

