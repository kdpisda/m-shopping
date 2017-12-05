VERSION 5.00
Begin VB.Form WelcomeForm 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFC0&
   Caption         =   "Login to m-Shopping"
   ClientHeight    =   7425
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   12255
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   12255
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   7455
      Left            =   -960
      Picture         =   "VBProject.frx":0000
      ScaleHeight     =   7395
      ScaleWidth      =   6435
      TabIndex        =   9
      Top             =   0
      Width           =   6495
   End
   Begin VB.CommandButton RegisterButton 
      BackColor       =   &H8000000B&
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
      Left            =   6120
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5280
      Width           =   5895
   End
   Begin VB.CommandButton LoginButton 
      BackColor       =   &H8000000B&
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
      Left            =   6120
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3360
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
      Height          =   525
      IMEMode         =   3  'DISABLE
      Left            =   7800
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2280
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
      Height          =   525
      Left            =   7800
      TabIndex        =   0
      Top             =   1440
      Width           =   4215
   End
   Begin VB.Label Heading 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H80000007&
      Height          =   735
      Left            =   6000
      TabIndex        =   2
      Top             =   360
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
      Left            =   6120
      TabIndex        =   8
      Top             =   6360
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
      Left            =   8160
      TabIndex        =   7
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label PasswordLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Left            =   6000
      TabIndex        =   6
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label UsernameLaabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Left            =   5880
      TabIndex        =   5
      Top             =   1440
      Width           =   1815
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

Private Sub Form_Load()
    LoginButton.Enabled = False
End Sub

Private Sub Admin_Click()
    frmLogin.Show
End Sub
'LoginButton.ForeColor = vbWhite

Private Sub LoginButton_Click()
    LoadingSplash.Show
    If LoadingSplash.Visible Then
        Set db = New ADODB.Connection
        db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path + "\mshopping.mdb"
        Set records = New ADODB.Recordset
        records.Open "Select count(*) from [users] where username = '" + username + "' and password = '" + password & "';", db, adOpenStatic, adLockOptimistic
        rec_ary = records.GetRows(1)
        If rec_ary(0, 0) = 1 Then
            Set records = New ADODB.Recordset
            records.Open "Select user_type, id from [users] where username = '" + username + "' and password = '" + password & "';", db, adOpenStatic, adLockOptimistic
            rec_ary = records.GetRows(1)
            If rec_ary(0, 0) = "admin" Then
                AdminPanelForm.Show
            Else
                UserPanelForm.UserIdHidden.Text = rec_ary(1, 0)
                UserPanelForm.UserNameHidden.Text = username
                UserPanelForm.Show
            End If
            Unload Me
        Else
            LoadingSplash.Hide
            MsgBox ("Wrong username and password combination")
        End If
    End If
    LoadingSplash.Hide
    'Unload Me
End Sub

Private Sub PasswordTextBox_Change()
    password = PasswordTextBox.Text
    If username = "" Then
        LoginButton.Enabled = False
    Else
        If password = "" Then
            LoginButton.Enabled = False
        Else
            LoginButton.Enabled = True
        End If
    End If
End Sub

Private Sub RegisterButton_Click()
    RegistrationForm.Show
    Unload Me
End Sub

Private Sub UsernameTextBox_Change()
    username = UserNameTextBox.Text
    If username = "" Then
        LoginButton.Enabled = False
    Else
        If password = "" Then
            LoginButton.Enabled = False
        Else
            LoginButton.Enabled = True
        End If
    End If
End Sub

