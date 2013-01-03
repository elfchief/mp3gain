VERSION 5.00
Begin VB.Form frmResetWarn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Clear Analysis?"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3510
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   3510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkShowResetWarn 
      Caption         =   "Don't ask me again"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   3255
   End
   Begin VB.CommandButton cmdNo 
      Cancel          =   -1  'True
      Caption         =   "&No"
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "&Yes"
      Default         =   -1  'True
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "This will clear all current analysis results. Are you sure?"
      Height          =   390
      Left            =   225
      TabIndex        =   0
      Top             =   120
      Width           =   3045
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmResetWarn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdNo_Click()
    frmMain.blnResetWarn = True
    frmMain.blnResetWarnResult = False
    Unload Me
End Sub

Private Sub cmdYes_Click()
    If Me.chkShowResetWarn.Value = vbChecked Then
        frmMain.blnResetWarn = False
    Else
        frmMain.blnResetWarn = True
    End If
    frmMain.blnResetWarnResult = True
    Unload Me
End Sub

Private Sub Form_Load()
    Dim diff1 As Single
    
    diff1 = Label1.Height
    
    fillCaptions Me
    
    diff1 = Label1.Height - diff1
    
    chkShowResetWarn.Top = chkShowResetWarn.Top + diff1
    cmdYes.Top = cmdYes.Top + diff1
    cmdNo.Top = cmdNo.Top + diff1
    Me.Height = Me.Height + diff1
End Sub
