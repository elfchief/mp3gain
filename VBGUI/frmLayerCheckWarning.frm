VERSION 5.00
Begin VB.Form frmLayerCheckWarning 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "WARNING!"
   ClientHeight    =   2685
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkIGetIt 
      Caption         =   "Don't show this warning again"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   4335
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   $"frmLayerCheckWarning.frx":0000
      Height          =   780
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4245
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   $"frmLayerCheckWarning.frx":00F0
      Height          =   585
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   4215
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmLayerCheckWarning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim diff2 As Single
    Dim diff1 As Single
    
    Me.Icon = LoadResPicture("APPICON", vbResIcon)
    
    diff2 = Label2.Height
    diff1 = Label1.Height
    
    fillCaptions Me
    
    diff2 = Label2.Height - diff2
    diff1 = Label1.Height - diff1
    
    Label1.Top = Label1.Top + diff2
    chkIGetIt.Top = chkIGetIt.Top + diff2 + diff1
    OKButton.Top = OKButton.Top + diff2 + diff1
    Me.Height = Me.Height + diff2 + diff1
    
End Sub

Private Sub OKButton_Click()
    If Me.chkIGetIt.Value Then
        frmMain.blnRecklessWarning = False
    Else
        frmMain.blnRecklessWarning = True
    End If
    Unload Me
End Sub
