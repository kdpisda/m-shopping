VERSION 5.00
Begin VB.Form RegistrationForm 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Create your account"
   ClientHeight    =   6450
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   7080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton SubmitButton 
      BackColor       =   &H8000000E&
      Height          =   1095
      Left            =   4080
      MaskColor       =   &H00FFFFFF&
      Picture         =   "RegistrationForm.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton CancelButton 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
      Height          =   1095
      Left            =   1920
      Picture         =   "RegistrationForm.frx":1402
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5160
      Width           =   1215
   End
   Begin VB.TextBox ContactNoBox 
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
      Left            =   3120
      TabIndex        =   8
      Top             =   4320
      Width           =   3255
   End
   Begin VB.TextBox PasswordBox 
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
      IMEMode         =   3  'DISABLE
      Left            =   3120
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   3600
      Width           =   3255
   End
   Begin VB.TextBox UserNameTextBox 
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
      Left            =   3120
      TabIndex        =   4
      Top             =   2880
      Width           =   3255
   End
   Begin VB.TextBox NameTextBox 
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
      Left            =   3120
      TabIndex        =   2
      Top             =   2160
      Width           =   3255
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Contact No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Top             =   4320
      Width           =   2175
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Registration"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2280
      TabIndex        =   0
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000E&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   5055
      Left            =   0
      Top             =   960
      Width           =   7095
   End
   Begin VB.Image Image1 
      Height          =   7500
      Left            =   -1200
      Picture         =   "RegistrationForm.frx":2804
      Top             =   -3240
      Width           =   7500
   End
End
Attribute VB_Name = "RegistrationForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim id As Integer
Dim UserName As String
Dim InputName As String
Dim InputPassword As String
Dim ContactNo As String
Dim UserType As String
Dim rec_ary As Variant
Dim rs As ADODB.Recordset
Dim cn As ADODB.Connection

Private Sub Form_Load()
    Set db = New ADODB.Connection
    db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path + "\mshopping.mdb"
    Set records = New ADODB.Recordset
    records.Open "Select count(*) from [users];", db, adOpenStatic, adLockOptimistic
    rec_ary = records.GetRows(1)
    id = rec_ary(0, 0) + 1
    UserType = "Guest"
End Sub

Private Sub CancelButton_Click()
    WelcomeForm.Show
    Unload Me
End Sub

Private Sub ContactNoBox_Change()
    ContactNo = ContactNoBox.Text
End Sub

Private Sub NameTextBox_Change()
    InputName = NameTextBox.Text
End Sub

Private Sub PasswordBox_Change()
    InputPassword = PasswordBox.Text
End Sub

Private Sub SubmitButton_Click()
    
    db_path = App.Path + "\mshopping.mdb"
    Set cn = New ADODB.Connection
        cn.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0; Data Source=" & db_path
    Set rs = New ADODB.Recordset
        rs.Open "Select * from users", cn, adOpenStatic, adLockOptimistic
        
    With rs
        .AddNew
        .Fields("id").Value = id
        .Fields("username").Value = UserName
        .Fields("name").Value = InputName
        .Fields("password").Value = InputPassword
        .Fields("contact_no").Value = ContactNo
        .Fields("user_type").Value = UserType
        .Update
    End With
    'On Error GoTo ErrHandle
'ErrHandle:
    'MsgBox "Happily chugging away in the error handler"
    'Set records = New ADODB.Recordset
    'records.AddNew
    'records.Fields("id").Value = id
    'records.Fields("username").Value = UserName
    'records.Fields("name").Value = InputName
    'records.Fields("password").Value = Passsword
    'records.Fields("contact_no").Value = ContactNo
    'records.Fields("user_type").Value = UserType
    'records.Update
    'records.Open "INSERT INTO [users]" & "(id, username, name, password, contact_no, user_type)" & " VALUES (" & id & ", " & UserName & ", " & Name & ", " & Password & ", " & ContactNo & ", guest )", db, adOpenStatic, adLockOptimistic
    'rec_ary = records.GetRows(1)
    'If rec_ary(0, 0) = "admin" Then
    '    AdminPanelForm.Show
    'Else
    '    UserPanelForm.UserIdHidden.Text = rec_ary(1, 0)
    '    UserPanelForm.UserNameHidden.Text = UserName
    '    UserPanelForm.Show
    'End If
    'MsgBox (records)
    WelcomeForm.Show
    Unload Me
End Sub

Private Sub UserNameTextBox_Change()
    UserName = UserNameTextBox.Text
End Sub

