VERSION 5.00
Begin VB.Form OrderForm 
   BackColor       =   &H00008080&
   Caption         =   "Order"
   ClientHeight    =   8670
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8895
   FillColor       =   &H0000C0C0&
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "OrderForm.frx":0000
   ScaleHeight     =   8670
   ScaleWidth      =   8895
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox MobileQuantity 
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
      Left            =   6960
      TabIndex        =   14
      Text            =   "00"
      Top             =   6000
      Width           =   1095
   End
   Begin VB.TextBox SelectedMobileIdHidden 
      Height          =   285
      Left            =   7560
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   1920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox UserNameHidden 
      Height          =   285
      Left            =   5280
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   1920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox UserIdHidden 
      Height          =   285
      Left            =   6600
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   1920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton BackButtonOrderSelect 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5160
      MaskColor       =   &H000080FF&
      MouseIcon       =   "OrderForm.frx":D947
      MousePointer    =   99  'Custom
      Picture         =   "OrderForm.frx":DA99
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Cancel"
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton SubmitImageButton 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6840
      MaskColor       =   &H000080FF&
      MouseIcon       =   "OrderForm.frx":EE9B
      MousePointer    =   99  'Custom
      Picture         =   "OrderForm.frx":EFED
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Confirm Order"
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Image MobileImageSelect 
      Height          =   5775
      Left            =   840
      Picture         =   "OrderForm.frx":103EF
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   3495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Quantity:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   4920
      TabIndex        =   13
      Top             =   6000
      Width           =   2055
   End
   Begin VB.Label GreetLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1320
      Width           =   6375
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   -120
      Top             =   1200
      Width           =   9135
   End
   Begin VB.Label MobileModelName 
      BackStyle       =   0  'Transparent
      Caption         =   "Model Name"
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
      Left            =   4920
      TabIndex        =   7
      Top             =   3000
      Width           =   4695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "m-Shopping"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   26.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1080
      TabIndex        =   6
      Top             =   240
      Width           =   3735
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   7920
      MouseIcon       =   "OrderForm.frx":2D841
      MousePointer    =   99  'Custom
      Picture         =   "OrderForm.frx":2D993
      ToolTipText     =   "Logout"
      Top             =   120
      Width           =   945
   End
   Begin VB.Image HomeImageButton 
      Height          =   1020
      Left            =   6840
      MouseIcon       =   "OrderForm.frx":2ED95
      MousePointer    =   99  'Custom
      Picture         =   "OrderForm.frx":2EEE7
      ToolTipText     =   "Home"
      Top             =   120
      Width           =   1005
   End
   Begin VB.Image BackImageButton 
      Height          =   945
      Left            =   120
      MouseIcon       =   "OrderForm.frx":30539
      MousePointer    =   99  'Custom
      Picture         =   "OrderForm.frx":3068B
      ToolTipText     =   "Back"
      Top             =   120
      Width           =   945
   End
   Begin VB.Label MobilePrice 
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "price"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   4920
      TabIndex        =   4
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label MobileColor 
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "color"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   6600
      TabIndex        =   3
      Top             =   4680
      Width           =   1935
   End
   Begin VB.Label MobileRam 
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "ram"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   4920
      TabIndex        =   2
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Label MobileDescription 
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1095
      Left            =   4920
      TabIndex        =   1
      Top             =   3480
      Width           =   3735
   End
   Begin VB.Label mobilename 
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
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
      Height          =   615
      Left            =   4920
      TabIndex        =   0
      Top             =   2400
      Width           =   4575
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H8000000E&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      FillColor       =   &H00800000&
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   14055
   End
End
Attribute VB_Name = "OrderForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SelectedMobileName As String
Dim SelectedMobileDescription As String
Dim SelectedMobileRam As String
Dim SelectedMobileColor As String
Dim SelectedMobilePrice As String
Dim SelectedMobileImage As String
Dim db As ADODB.Connection
Dim records As ADODB.Recordset

Private Sub BackButtonOrderSelect_Click()
    UserPanelForm.SelectMobileId = ""
    UserPanelForm.Show
    Unload Me
End Sub

Private Sub BackImageButton_Click()
    UserPanelForm.SelectMobileId = ""
    UserPanelForm.Show
    Unload Me
End Sub

Private Sub Form_Load()
    Dim MobileId As Integer
    MobileId = UserPanelForm.SelectMobileId
    SelectedMobileIdHidden.Text = MobileId
    GreetLabel.Caption = "Welcome " & UserNameHidden.Text
    Set db = New ADODB.Connection
    db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path + "\mshopping.mdb"
    Set records = New ADODB.Recordset
    records.Open "Select * from [mobiles] Where id = " & MobileId & ";", db, adOpenStatic, adLockOptimistic
    rec_ary = records.GetRows(1)
    SelectedMobileName = rec_ary(1, 0)
    SelectedMobilePrice = rec_ary(3, 0)
    SelectedMobileDescription = rec_ary(4, 0)
    SelectedMobileColor = rec_ary(6, 0)
    SelectedMobileRam = rec_ary(14, 0)
    SelectedMobileImage = rec_ary(17, 0)
    mobilename.Caption = SelectedMobileName
    MobilePrice.Caption = SelectedMobilePrice
    MobileDescription.Caption = SelectedMobileDescription
    MobileColor.Caption = SelectedMobileColor
    MobileRam.Caption = SelectedMobileRam
    MobileImageSelect.Picture = LoadPicture(App.Path + "\uploads\" + SelectedMobileImage)
End Sub

Private Sub HomeImageButton_Click()
    UserPanelForm.UserIdHidden.Text = UserIdHidden.Text
    UserPanelForm.UserNameHidden.Text = UserNameHidden.Text
    UserPanelForm.Show
    Unload Me
End Sub


Private Sub SubmitImageButton_Click()
    If MobileQuantity.Text > 0 And MobileQuantity.Text < 21 Then
        DetailsForm.Show
        DetailsForm.SelectedMobileNameLabel.Caption = mobilename.Caption
        DetailsForm.SelectedMobilePriceLabel.Caption = (MobilePrice.Caption * MobileQuantity.Text) & " Rs"
        If MobileQuantity.Text = 1 Then
            DetailsForm.QuantityLabel.Caption = MobileQuantity.Text & " Pc"
        Else
            DetailsForm.QuantityLabel.Caption = MobileQuantity.Text & " Pcs"
        End If
        DetailsForm.SelectedMobileImage.Picture = MobileImageSelect.Picture
        DetailsForm.SelectedMobileImage.Width = 855
        DetailsForm.SelectedMobileImage.Height = 855
        DetailsForm.QuantityHidden.Text = MobileQuantity.Text
        DetailsForm.SelectedMobileIdHidden.Text = SelectedMobileIdHidden.Text
        DetailsForm.UserIdHidden.Text = UserIdHidden.Text
        DetailsForm.UserNameHidden.Text = UserNameHidden.Text
        Unload Me
    Else
        MsgBox ("You can enter quantity in range of 1 - 20 both inclusive")
    End If
End Sub
