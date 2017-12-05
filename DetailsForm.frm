VERSION 5.00
Begin VB.Form DetailsForm 
   BackColor       =   &H00808000&
   Caption         =   "Please fill the Details"
   ClientHeight    =   8670
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
   ScaleHeight     =   8670
   ScaleWidth      =   9780
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox SelectedMobileIdHidden 
      Height          =   645
      Left            =   8760
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   2880
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox UserNameHidden 
      Height          =   645
      Left            =   8760
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   5400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox UserIdHidden 
      Height          =   645
      Left            =   8760
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   4440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox QuantityHidden 
      Height          =   645
      Left            =   8760
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   3600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   12
      Top             =   7920
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "NEXT"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   6600
      TabIndex        =   11
      Top             =   7920
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      TabIndex        =   10
      Top             =   7920
      Width           =   2415
   End
   Begin VB.TextBox PHNO 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   4080
      TabIndex        =   8
      Top             =   6840
      Width           =   3975
   End
   Begin VB.TextBox PIN 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   4080
      TabIndex        =   6
      Top             =   5760
      Width           =   3975
   End
   Begin VB.TextBox ADD2 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   4080
      TabIndex        =   4
      Top             =   4800
      Width           =   3975
   End
   Begin VB.TextBox ADD1 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   4080
      TabIndex        =   2
      Top             =   3840
      Width           =   3975
   End
   Begin VB.TextBox NAME11 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   4080
      TabIndex        =   0
      Top             =   2880
      Width           =   3975
   End
   Begin VB.Label QuantityLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1 Pc"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   6240
      TabIndex        =   17
      Top             =   1920
      Width           =   930
   End
   Begin VB.Label SelectedMobilePriceLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1000 $"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   3480
      TabIndex        =   16
      Top             =   2160
      Width           =   2655
   End
   Begin VB.Label SelectedMobileNameLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile Name"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   3480
      TabIndex        =   15
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Image SelectedMobileImage 
      Enabled         =   0   'False
      Height          =   855
      Left            =   2520
      Picture         =   "DetailsForm.frx":1B75AE
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   855
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H8000000E&
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   -120
      Top             =   1800
      Width           =   10335
   End
   Begin VB.Label GreetTextLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Please add your details for shippping"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   1320
      Width           =   9255
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   -120
      Top             =   1200
      Width           =   10215
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
      TabIndex        =   13
      Top             =   240
      Width           =   3975
   End
   Begin VB.Image LogoutImageButton 
      Height          =   945
      Left            =   8760
      MouseIcon       =   "DetailsForm.frx":1D4A00
      MousePointer    =   99  'Custom
      Picture         =   "DetailsForm.frx":1D4B52
      ToolTipText     =   "Logout"
      Top             =   120
      Width           =   945
   End
   Begin VB.Image HomeImageButton 
      Height          =   1020
      Left            =   7560
      MouseIcon       =   "DetailsForm.frx":1D5F54
      MousePointer    =   99  'Custom
      Picture         =   "DetailsForm.frx":1D60A6
      ToolTipText     =   "Home"
      Top             =   120
      Width           =   1005
   End
   Begin VB.Image BackImageButton 
      Height          =   945
      Left            =   120
      MouseIcon       =   "DetailsForm.frx":1D76F8
      MousePointer    =   99  'Custom
      Picture         =   "DetailsForm.frx":1D784A
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
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Contact No"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   9
      Top             =   6840
      Width           =   2895
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "PIN Code"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   7
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "City"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   5
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   3
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
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
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   3000
      Width           =   1575
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
    UserPanelForm.UserIdHidden.Text = UserIdHidden.Text
    UserPanelForm.UserNameHidden.Text = UserNameHidden.Text
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
        BillingForm.SelectedMobileIdHidden.Text = SelectedMobileIdHidden.Text
        BillingForm.UserIdHidden.Text = UserIdHidden.Text
        BillingForm.Label14.Caption = PHONE
        BillingForm.Label12.Caption = NAME1
        BillingForm.Label13.Caption = ADDRESS
        BillingForm.MobileNameLabel.Caption = SelectedMobileNameLabel.Caption
        BillingForm.QuantityLabel.Caption = QuantityLabel.Caption
        BillingForm.PriceLabel.Caption = SelectedMobilePriceLabel.Caption
    Else
        MsgBox "contact number has to be 10 digits"
    End If
    'OrderForm.mobilename = BillingForm.BillingMobileName
    'OrderForm.MobilePrice = BillingForm.BillingMobilePrice
End Sub

Private Sub Command3_Click()
    UserPanelForm.Show
    UserPanelForm.UserIdHidden.Text = UserIdHidden.Text
    UserPanelForm.UserNameHidden.Text = UserNameHidden.Text
    Unload Me
End Sub

Private Sub HomeImageButton_Click()
    UserPanelForm.Show
    UserPanelForm.UserIdHidden.Text = UserIdHidden.Text
    UserPanelForm.UserNameHidden.Text = UserNameHidden.Text
    Unload Me
End Sub

Private Sub LogoutImageButton_Click()
    UserPanelForm.Show
    UserPanelForm.UserIdHidden.Text = UserIdHidden.Text
    UserPanelForm.UserNameHidden.Text = UserNameHidden.Text
    Unload Me
End Sub

Private Sub Form_Load()
    SelectedMobileIdHidden.Text = OrderForm.SelectedMobileIdHidden.Text
    SelectedMobileImage.Height = 855
    SelectedMobileImage.Width = 855
End Sub

Private Sub PHNO_Change()
    ContactNo = PHNO.Text
End Sub

