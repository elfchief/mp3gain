VERSION 5.00
Begin VB.Form frmSkipTagsWarn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "WARNING!"
   ClientHeight    =   1950
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4005
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   4005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton OKButton 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CheckBox chkIGetIt 
      Caption         =   "Don't show this warning again"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   3735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   $"frmSkipTagsWarn.frx":0000
      Height          =   705
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3765
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmSkipTagsWarn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim diff1 As Single
    
    diff1 = Label1.Height
    
    fillCaptions Me
    
    diff1 = Label1.Height - diff1
    
    chkIGetIt.Top = chkIGetIt.Top + diff1
    OKButton.Top = OKButton.Top + diff1
    Me.Height = Me.Height + diff1
End Sub

Private Sub OKButton_Click()
    If Me.chkIGetIt.Value Then
        frmMain.blnSkipTagsWarn = False
    Else
        frmMain.blnSkipTagsWarn = True
    End If
    Unload Me
End Sub

