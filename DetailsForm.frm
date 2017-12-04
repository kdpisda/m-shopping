VERSION 5.00
Begin VB.Form DetailsForm 
   BackColor       =   &H00808000&
   Caption         =   "Please fill the Details"
   ClientHeight    =   7605
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9780
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   18
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   Picture         =   "DetailsForm.frx":0000
   ScaleHeight     =   7605
   ScaleWidth      =   9780
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "BACK"
      Height          =   615
      Left            =   480
      TabIndex        =   13
      Top             =   6840
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "NEXT"
      Height          =   645
      Left            =   6600
      TabIndex        =   12
      Top             =   6840
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CLEAR"
      Height          =   615
      Left            =   3600
      TabIndex        =   11
      Top             =   6840
      Width           =   2415
   End
   Begin VB.TextBox PHNO 
      Height          =   645
      Left            =   4080
      TabIndex        =   9
      Top             =   5760
      Width           =   3975
   End
   Begin VB.TextBox PIN 
      Height          =   645
      Left            =   4080
      TabIndex        =   7
      Top             =   4680
      Width           =   3975
   End
   Begin VB.TextBox ADD2 
      Height          =   645
      Left            =   4080
      TabIndex        =   5
      Top             =   3720
      Width           =   3975
   End
   Begin VB.TextBox ADD1 
      Height          =   645
      Left            =   4080
      TabIndex        =   3
      Top             =   2760
      Width           =   3975
   End
   Begin VB.TextBox NAME11 
      Height          =   645
      Left            =   4080
      TabIndex        =   1
      Top             =   1800
      Width           =   3975
   End
   Begin VB.Label MobileIdCaption 
      Caption         =   "id"
      Height          =   615
      Left            =   8280
      TabIndex        =   15
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "m-Shopping"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1080
      TabIndex        =   14
      Top             =   240
      Width           =   3975
   End
   Begin VB.Image LogoutImageButton 
      Height          =   945
      Left            =   8760
      MouseIcon       =   "DetailsForm.frx":1B75AE
      MousePointer    =   99  'Custom
      Picture         =   "DetailsForm.frx":1B7700
      ToolTipText     =   "Logout"
      Top             =   120
      Width           =   945
   End
   Begin VB.Image HomeImageButton 
      Height          =   1020
      Left            =   7560
      MouseIcon       =   "DetailsForm.frx":1B8B02
      MousePointer    =   99  'Custom
      Picture         =   "DetailsForm.frx":1B8C54
      ToolTipText     =   "Home"
      Top             =   120
      Width           =   1005
   End
   Begin VB.Image BackImageButton 
      Height          =   945
      Left            =   120
      MouseIcon       =   "DetailsForm.frx":1BA2A6
      MousePointer    =   99  'Custom
      Picture         =   "DetailsForm.frx":1BA3F8
      ToolTipText     =   "Back"
      Top             =   120
      Width           =   945
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   10095
   End
   Begin VB.Label Label6 
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "CONTACT NO."
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   5760
      Width           =   2895
   End
   Begin VB.Label Label5 
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "PINCODE"
      Height          =   495
      Left            =   1200
      TabIndex        =   8
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "CITY"
      Height          =   495
      Left            =   1920
      TabIndex        =   6
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "ADRESS"
      Height          =   495
      Left            =   1320
      TabIndex        =   4
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "NAME"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "PERSONAL DETAILS"
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   2400
      TabIndex        =   0
      Top             =   1200
      Width           =   4695
   End
End
Attribute VB_Name = "DetailsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NAME1 As String
Dim ADD11 As String
Dim ADD21 As String
Dim PIN1 As String
Dim PHONE As String
Dim ADDRESS As String
Dim ContactNo As String


Private Sub BackImageButton_Click()
    UserPanelForm.Show
    Unload Me
End Sub

Private Sub Command1_Click()
    NAME11.Text = ""
    ADD1.Text = ""
    ADD2.Text = ""
    PIN.Text = ""
    PHNO.Text = ""
End Sub

Private Sub Command2_Click()
    Dim MSG As Integer
    
    NAME1 = NAME11.Text
    ADD11 = ADD1.Text
    ADD21 = ADD2.Text
    PIN1 = PIN.Text
    PHONE = PHNO.Text
    
    
    ADDRESS = ADD11 + "," + ADD21 + "," + PIN1 + "."
    
    If NAME11.Text = "" Then
        MSG = MsgBox("OOPS!! Name Field Cant be Empty", vbOKOnly)
    ElseIf ADD1.Text = "" Then
        MSG = MsgBox("OOPS!! Address Field Cant be Empty", vbOKOnly)
    ElseIf ADD2.Text = "" Then
        MSG = MsgBox("OOPS!! City Field Cant be Empty", vbOKOnly)
    ElseIf PIN.Text = "" Then
        MSG = MsgBox("OOPS!! Pincode Field Cant be Empty", vbOKOnly)
    End If
    If Len(ContactNo) = 10 Then
        DetailsForm.Hide
        BillingForm.Show
        BillingForm.Label14.Caption = PHONE
        BillingForm.Label12.Caption = NAME1
        BillingForm.Label13.Caption = ADDRESS
    Else
        MsgBox "contact number has to be 10 digits"
    End If
    OrderForm.MobileName = BillingForm.BillingMobileName
    OrderForm.MobilePrice = BillingForm.BillingMobilePrice
End Sub

Private Sub Command3_Click()
DETAILS.Hide
Order.Show

End Sub

Private Sub HomeImageButton_Click()
    UserPanelForm.Show
    Unload Me
End Sub

Private Sub LogoutImageButton_Click()
    UserPanelForm.Show
    Unload Me
End Sub

Private Sub Form_Load()
    MobileIdCaption.Caption = OrderForm.MobileIdCaption
End Sub

Private Sub PHNO_Change()
    ContactNo = PHNO.Text
End Sub
