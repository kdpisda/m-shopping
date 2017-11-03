VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form LoadingSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2430
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   5850
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "loadingSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   5850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2235
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   5505
      Begin VB.Timer ProgressBarTimer 
         Interval        =   100
         Left            =   4680
         Top             =   480
      End
      Begin ComctlLib.ProgressBar LoadingProgressBar 
         Height          =   735
         Left            =   360
         TabIndex        =   2
         Top             =   1080
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   1296
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         Caption         =   "m-Shopping"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   1440
         TabIndex        =   1
         Top             =   240
         Width           =   2640
      End
   End
End
Attribute VB_Name = "LoadingSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    LoadingProgressBar.Value = 0
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
        Unload Me
    End If
End Sub
