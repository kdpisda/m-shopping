VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form WelcomeSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5580
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   12540
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   12540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   5355
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12225
      Begin VB.Timer ProgressBarTimer 
         Interval        =   100
         Left            =   11400
         Top             =   2400
      End
      Begin ComctlLib.ProgressBar LoadingProgressBar 
         Height          =   495
         Left            =   6960
         TabIndex        =   7
         Top             =   4560
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   873
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Frame Frame2 
         Caption         =   "Group 11"
         Height          =   1335
         Left            =   6960
         TabIndex        =   4
         Top             =   3000
         Width           =   4815
         Begin VB.Label Label1 
            Caption         =   "Kuldeep Pisda, Bhavna Sahu, Prince Jain, Shanmukha"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   120
            TabIndex        =   5
            Top             =   360
            Width           =   4575
         End
      End
      Begin VB.PictureBox Picture1 
         Height          =   5175
         Left            =   0
         Picture         =   "frmSplash.frx":000C
         ScaleHeight     =   5115
         ScaleWidth      =   6675
         TabIndex        =   1
         Top             =   120
         Width           =   6735
      End
      Begin VB.Label Label2 
         Caption         =   "Mobile Shopping"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   13.5
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6960
         TabIndex        =   6
         Top             =   2280
         Width           =   4695
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         Caption         =   "m-Shopping"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   6960
         TabIndex        =   3
         Top             =   1320
         Width           =   3915
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         Caption         =   "VB Project Group 11"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   11895
      End
   End
End
Attribute VB_Name = "WelcomeSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    'lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    'lblProductName.Caption = App.Title
    
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub ProgressBarTimer_Timer()
    ProgressBarTimer.Interval = Rnd * 300 + 10
    LoadingProgressBar.Value = LoadingProgressBar.Value + 2
    If LoadingProgressBar.Value = 100 Then
        WelcomeForm.Show
    Unload Me
    End If
End Sub
