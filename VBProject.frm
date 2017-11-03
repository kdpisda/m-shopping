VERSION 5.00
Begin VB.Form WelcomeForm 
   Caption         =   "Login to m-Shopping"
   ClientHeight    =   6120
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   10770
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   10770
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   6375
      Left            =   0
      Picture         =   "VBProject.frx":0000
      ScaleHeight     =   6315
      ScaleWidth      =   4275
      TabIndex        =   9
      Top             =   0
      Width           =   4335
      Begin VB.CommandButton Admin 
         Caption         =   "Admin Login"
         Height          =   495
         Left            =   840
         TabIndex        =   10
         Top             =   5400
         Width           =   2415
      End
   End
   Begin VB.CommandButton RegisterButton 
      Caption         =   "Register"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4560
      TabIndex        =   6
      Top             =   4440
      Width           =   5895
   End
   Begin VB.CommandButton LoginButton 
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4560
      TabIndex        =   5
      Top             =   3000
      Width           =   5895
   End
   Begin VB.TextBox PasswordTextBox 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      IMEMode         =   3  'DISABLE
      Left            =   6360
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2160
      Width           =   4215
   End
   Begin VB.TextBox UsernameTextBox 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   6360
      TabIndex        =   2
      Top             =   1320
      Width           =   4215
   End
   Begin VB.Label Heading 
      Alignment       =   2  'Center
      Caption         =   "m-Shopping"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4560
      TabIndex        =   0
      Top             =   240
      Width           =   5895
   End
   Begin VB.Label FooterLabel 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile Shopping by Group No 11"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   8
      Top             =   5400
      Width           =   5775
   End
   Begin VB.Label OrLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Or"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   7
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Label PasswordLabel 
      Alignment       =   2  'Center
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   3
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label UsernameLaabel 
      Alignment       =   2  'Center
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   1440
      Width           =   1695
   End
End
Attribute VB_Name = "WelcomeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As ADODB.Connection
Dim records As ADODB.Recordset
Dim rec_ary As Variant
Dim username As String
Dim password As String

Private Sub Admin_Click()
    frmLogin.Show
End Sub

Private Sub LoginButton_Click()
    LoadingSplash.Show
    If LoadingSplash.Visible Then
        Set db = New ADODB.Connection
        db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path + "\mshopping.mdb"
        Set records = New ADODB.Recordset
        records.Open "Select count(*) from [users] where username = '" + username + "' and password = '" + password & "';", db, adOpenStatic, adLockOptimistic
        rec_ary = records.GetRows(1)
        If rec_ary = 1 Then
            SampleForm.Show
            Unload Me
        Else
            LoadingSplash.Hide
        End If
    End If
    LoadingSplash.Hide
    Unload Me
End Sub

Private Sub PasswordTextBox_Change()
    password = PasswordTextBox.Text
End Sub

Private Sub UsernameTextBox_Change()
    username = UsernameTextBox.Text
End Sub

