VERSION 5.00
Begin VB.Form BillingForm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Billing"
   ClientHeight    =   8910
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   11280
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form2"
   Picture         =   "BillingForm.frx":0000
   ScaleHeight     =   8910
   ScaleWidth      =   11280
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox UserIdHidden 
      Height          =   285
      Left            =   8520
      TabIndex        =   11
      Text            =   "Text2"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox SelectedMobileIdHidden 
      Height          =   285
      Left            =   8520
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00404040&
      Caption         =   "Confirm"
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
      Left            =   6960
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8160
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00404040&
      Caption         =   "Previous"
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
      Left            =   2160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8160
      Width           =   2175
   End
   Begin VB.Label QuantityLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   6240
      TabIndex        =   9
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
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
      Left            =   1920
      TabIndex        =   8
      Top             =   2880
      Width           =   7455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Shipping Address"
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
      Height          =   495
      Left            =   2400
      TabIndex        =   7
      Top             =   5520
      Width           =   3135
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "CONTACT "
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   2400
      TabIndex        =   6
      Top             =   6480
      Width           =   3015
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "NAME"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   6000
      Width           =   3255
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "ADDRESS"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   6960
      Width           =   10335
   End
   Begin VB.Label PriceLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label MobileNameLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "MOBILE"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   3720
      Width           =   6255
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   5535
      Left            =   1560
      Shape           =   4  'Rounded Rectangle
      Top             =   2280
      Width           =   8055
   End
End
Attribute VB_Name = "BillingForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OrderId As Integer
Dim MobileId As Integer
Dim UserId As Integer
Dim Quantity As Integer
Dim Price As Integer
Dim Attendee As String
Dim Address As String
Dim ContactNo As String
Dim rs As ADODB.Recordset
Dim cn As ADODB.Connection

Private Sub Command1_Click()
BillingForm.Hide
DetailsForm.Show
End Sub

Private Sub Command2_Click()
    Dim P As Integer
    db_path = App.Path + "\mshopping.mdb"
    Set cn = New ADODB.Connection
        cn.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0; Data Source=" & db_path
    Set rs = New ADODB.Recordset
        rs.Open "Select * from orders", cn, adOpenStatic, adLockOptimistic
        
    With rs
        .AddNew
        .Fields("id").Value = OrderId
        .Fields("mobile_id").Value = MobileId
        .Fields("user_id").Value = UserId
        .Fields("attendee").Value = Attendee
        .Fields("address").Value = Address
        .Fields("contact_no").Value = ContactNo
        .Fields("quantity").Value = Quantity
        .Fields("price").Value = Price
        .Update
    End With
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

Private Sub Form_Load()
    Set db = New ADODB.Connection
    db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path + "\mshopping.mdb"
    Set records = New ADODB.Recordset
    records.Open "Select count(*) from [orders];", db, adOpenStatic, adLockOptimistic
    rec_ary = records.GetRows(1)
    OrderId = rec_ary(0, 0) + 1
    MobileId = Val(SelectedMobileIdHidden.Text)
    UserId = Val(UserIdHidden.Text)
    Quantity = Val(QuantityLabel.Caption)
    Price = Val(PriceLabel.Caption)
    ContactNo = Label14.Caption
    Address = Label13.Caption
    Attendee = Label12.Caption
End Sub
