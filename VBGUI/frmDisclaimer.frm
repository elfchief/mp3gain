VERSION 5.00
Begin VB.Form frmDisclaimer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DISCLAIMER"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5640
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "DISCLAIMER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   5175
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   $"frmDisclaimer.frx":0000
         Height          =   780
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Width           =   4815
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   $"frmDisclaimer.frx":0114
         Height          =   780
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   4695
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   2400
      Width           =   1215
   End
End
Attribute VB_Name = "frmDisclaimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim start7 As Single
    Dim start8 As Single
    Dim diff7 As Single
    Dim diff8 As Single
        
    Me.Icon = LoadResPicture("APPICON", vbResIcon)
    
    start7 = Label7.Height
    start8 = Label8.Height
    
    fillCaptions Me
    diff7 = Label7.Height - start7
    diff8 = Label8.Height - start8
    
    Label8.Top = Label8.Top + diff7
    Frame1.Height = Frame1.Height + diff7 + diff8
    cmdOK.Top = cmdOK.Top + diff7 + diff8
    Me.Height = Me.Height + diff7 + diff8
End Sub
