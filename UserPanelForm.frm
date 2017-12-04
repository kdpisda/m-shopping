VERSION 5.00
Begin VB.Form UserPanelForm 
   Caption         =   "Select Mobiles"
   ClientHeight    =   7455
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13635
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   13635
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox UserNameHidden 
      Height          =   285
      Left            =   11040
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   6240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox UserIdHidden 
      Height          =   285
      Left            =   9960
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   6240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox SelectMobileId 
      Height          =   285
      Left            =   8280
      TabIndex        =   5
      Text            =   "Mobile Id"
      Top             =   6240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ListBox MobileListBox 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2310
      Left            =   8280
      TabIndex        =   4
      Top             =   3720
      Width           =   4335
   End
   Begin VB.CommandButton HistoryButton 
      Caption         =   "History"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10560
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton LogoutButton 
      Caption         =   "Logout"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12120
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   7500
      Left            =   -240
      Picture         =   "UserPanelForm.frx":0000
      Top             =   0
      Width           =   7500
   End
   Begin VB.Label FindMobileLabel 
      Alignment       =   2  'Center
      Caption         =   "Please select a mobile to continue"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8400
      TabIndex        =   3
      Top             =   2520
      Width           =   4095
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00800000&
      Height          =   735
      Left            =   0
      Top             =   0
      Width           =   13695
   End
   Begin VB.Label UserNameLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   7320
      TabIndex        =   0
      Top             =   840
      Width           =   6255
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00008080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C00000&
      Height          =   735
      Left            =   6600
      Top             =   600
      Width           =   7695
   End
End
Attribute VB_Name = "UserPanelForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim username As String
Dim db As ADODB.Connection
Dim records As ADODB.Recordset
Dim MobileNumbers As Integer
Dim rec_ary As Variant
Dim MobileDisplayName As String
Dim mobilename As String
Dim MobileDescription As String
Dim MobileRam As String
Dim MobileColor As String
Dim MobilePrice As String
Dim MobileImage As String


Private Sub Form_Load()
    username = WelcomeForm.UsernameTextBox
    UserNameLabel.Caption = "Welcome, " + username
    LoadingSplash.Show
    If LoadingSplash.Visible Then
        Set db = New ADODB.Connection
        db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path + "\mshopping.mdb"
        Set records = New ADODB.Recordset
        records.Open "Select count(*) from [mobiles] ;", db, adOpenStatic, adLockOptimistic
        rec_ary = records.GetRows(1)
        MobileNumbers = rec_ary(0, 0)
        Set records = New ADODB.Recordset
        records.Open "Select * from [mobiles] ;", db, adOpenStatic, adLockOptimistic
        rec_ary = records.GetRows(MobileNumbers)
        For i = 0 To MobileNumbers - 1
            MobileDisplayName = rec_ary(1, i) + " " + rec_ary(5, i) + " " + rec_ary(6, i) + " " + rec_ary(3, i) + " Rs"
            MobileListBox.AddItem (MobileDisplayName)
        Next i
    End If
    LoadingSplash.Hide
End Sub

Private Sub LogoutButton_Click()
    WelcomeForm.UsernameTextBox = ""
    username = ""
    WelcomeForm.Show
    Unload Me
End Sub

Private Sub MobileListBox_Click()
    SelectMobileId.Text = rec_ary(0, MobileListBox.ListIndex)
    'MsgBox (SelectMobileId.Text)
    mobilename = rec_ary(1, MobileListBox.ListIndex)
    'MsgBox (MobileName)
    MobilePrice = rec_ary(3, MobileListBox.ListIndex)
    MobileDescription = rec_ary(4, MobileListBox.ListIndex)
    MobileColor = rec_ary(6, MobileListBox.ListIndex)
    MobileRam = rec_ary(14, MobileListBox.ListIndex)
    MobileImage = rec_ary(17, MobileListBox.ListIndex)
    OrderForm.UserIdHidden.Text = UserIdHidden.Text
    OrderForm.UserNameHidden.Text = UserNameHidden.Text
    OrderForm.Show
    Unload Me
    'MsgBox (MobileListBox.ListIndex)
End Sub
