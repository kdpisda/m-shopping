VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form WelcomeSplash 
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5730
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   12540
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000E&
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   12540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Height          =   5475
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12225
      Begin ComctlLib.ProgressBar LoadingProgressBar 
         Height          =   495
         Left            =   7320
         TabIndex        =   6
         ToolTipText     =   "Loading Please Wait..."
         Top             =   4560
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   873
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Timer ProgressBarTimer 
         Interval        =   100
         Left            =   11400
         Top             =   2400
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000E&
         Caption         =   "Group 11"
         Height          =   1335
         Left            =   7320
         TabIndex        =   3
         Top             =   3000
         Width           =   4695
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            Caption         =   "Kuldeep Pisda, Bhawana Sahu, Prince Jain, Shanmukha"
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
            TabIndex        =   4
            Top             =   360
            Width           =   4335
         End
      End
      Begin VB.Image Image1 
         Height          =   7500
         Left            =   -360
         Picture         =   "frmSplash.frx":000C
         Top             =   -960
         Width           =   7500
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   7320
         TabIndex        =   5
         Top             =   2280
         Width           =   4455
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00FF0000&
         Height          =   915
         Left            =   7320
         TabIndex        =   2
         Top             =   1320
         Width           =   3915
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         TabIndex        =   1
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

Private Sub Form_Load()
    'lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    'lblProductName.Caption = App.Title
    ProgressBarTimer.Enabled = True
    LoadingProgressBar.Value = 0
End Sub

Private Sub ProgressBarTimer_Timer()
    ProgressBarTimer.Interval = Rnd * 300 + 10
    LoadingProgressBar.Value = LoadingProgressBar.Value + 2
    If LoadingProgressBar.Value = 100 Then
        ProgressBarTimer.Enabled = False
        WelcomeForm.Show
        Unload Me
    End If
    'WelcomeForm.Show
    'Unload Me
End Sub
