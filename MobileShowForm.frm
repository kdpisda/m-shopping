VERSION 5.00
Begin VB.Form MobileShowForm 
   BackColor       =   &H8000000E&
   Caption         =   "Form1"
   ClientHeight    =   5595
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10470
   LinkTopic       =   "Form1"
   ScaleHeight     =   5595
   ScaleWidth      =   10470
   StartUpPosition =   3  'Windows Default
   Begin VB.Label PriceCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      Top             =   1200
      Width           =   3855
   End
   Begin VB.Label MobileCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile Caption"
      Height          =   255
      Left            =   5520
      TabIndex        =   1
      Top             =   960
      Width           =   4695
   End
   Begin VB.Label MobileName 
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   0
      Top             =   480
      Width           =   3975
   End
   Begin VB.Image MobileImage 
      Height          =   5655
      Left            =   0
      Top             =   0
      Width           =   5295
   End
End
Attribute VB_Name = "MobileShowForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
