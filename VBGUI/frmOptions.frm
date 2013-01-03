VERSION 5.00
Begin VB.Form frmOptions 
   Caption         =   "Advanced Options"
   ClientHeight    =   3600
   ClientLeft      =   1740
   ClientTop       =   2265
   ClientWidth     =   6660
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   3600
   ScaleWidth      =   6660
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkMaximizing 
      Caption         =   "Enable ""Maximizing"" features"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   3000
      Width           =   4935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Performance"
      Height          =   2655
      Left            =   3300
      TabIndex        =   10
      Top             =   180
      Width           =   3135
      Begin VB.CheckBox chkNoShowFileStatus 
         Caption         =   "Do not show file progress"
         Height          =   495
         Left            =   300
         TabIndex        =   6
         Top             =   1920
         Width           =   2535
      End
      Begin VB.CheckBox chkNoTempFiles 
         Caption         =   "Do not use Temp files"
         Height          =   435
         Left            =   300
         TabIndex        =   5
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Label Label3 
         Caption         =   $"frmOptions.frx":0000
         Height          =   1095
         Left            =   180
         TabIndex        =   11
         Top             =   240
         Width           =   2835
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   5280
      TabIndex        =   8
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Frame fraPriority 
      Caption         =   "Thread Priority"
      Height          =   2655
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   2775
      Begin VB.OptionButton optIdle 
         Caption         =   "Idle"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1440
         Width           =   2055
      End
      Begin VB.OptionButton optNorm 
         Caption         =   "Normal"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   2055
      End
      Begin VB.OptionButton optHigh 
         Caption         =   "High"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   2055
      End
      Begin VB.OptionButton optReal 
         Caption         =   "Realtime"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Note: High and Realtime are NOT Recommended"
         Height          =   735
         Left            =   240
         TabIndex        =   9
         Top             =   1800
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim blnActivating As Boolean

Private Sub chkMaximizing_Click()
    Dim sngBase As Single
    
    If blnActivating Then
        Exit Sub
    End If
    
    If chkMaximizing.Value = vbChecked Then
        frmMaximizing.Show vbModal, Me
        
        frmMain.mnuMaxAmp.Visible = True
        frmMain.mnuSep2.Visible = True
        frmMain.mnuMaxNoClipGain.Visible = True
        frmMain.mnuGroupNoClip.Visible = True
        frmMain.mnuSep11.Visible = True
        
        frmMain.ResetColumnWidths
    Else
        frmMain.mnuMaxAmp.Visible = False
        frmMain.mnuSep2.Visible = False
        frmMain.mnuMaxNoClipGain.Visible = False
        frmMain.mnuGroupNoClip.Visible = False
        frmMain.mnuSep11.Visible = False
        
        frmMain.ResetColumnWidths
    End If
End Sub

Private Sub chkNoShowFileStatus_Click()
    If chkNoShowFileStatus.Value = vbChecked Then
        frmMain.blnShowFileStatus = False
        frmMain.prgFile.Visible = False
        frmMain.lblFileProg.Visible = False
    Else
        frmMain.blnShowFileStatus = True
        frmMain.prgFile.Visible = True
        frmMain.lblFileProg.Visible = True
    End If
End Sub

Private Sub chkNoTempFiles_Click()
    If chkNoTempFiles.Value = vbChecked Then
        frmMain.blnUseTempFiles = False
    Else
        frmMain.blnUseTempFiles = True
    End If
End Sub

Private Sub cmdOK_Click()
    Me.Hide
End Sub

Private Sub Form_Activate()
    blnActivating = True
    
    Select Case lngThreadPriority
        Case IDLE_PRIORITY_CLASS
            optIdle.Value = True
            optIdle.SetFocus
        Case HIGH_PRIORITY_CLASS
            optHigh.Value = True
            optHigh.SetFocus
        Case NORMAL_PRIORITY_CLASS
            optNorm.Value = True
            optNorm.SetFocus
        Case REALTIME_PRIORITY_CLASS
            optReal.Value = True
            optReal.SetFocus
    End Select
    
    If frmMain.mnuMaxAmp.Visible Then
        chkMaximizing.Value = vbChecked
    End If
    
    If frmMain.blnUseTempFiles Then
        chkNoTempFiles.Value = vbUnchecked
    Else
        chkNoTempFiles.Value = vbChecked
    End If
    
    If frmMain.blnShowFileStatus Then
        Me.chkNoShowFileStatus.Value = vbUnchecked
    Else
        chkNoShowFileStatus.Value = vbChecked
    End If
    
    blnActivating = False
End Sub

Private Sub Form_Load()
    fillCaptions Me
End Sub

Private Sub optHigh_Click()
    lngThreadPriority = HIGH_PRIORITY_CLASS
End Sub

Private Sub optIdle_Click()
    lngThreadPriority = IDLE_PRIORITY_CLASS
End Sub

Private Sub optNorm_Click()
    lngThreadPriority = NORMAL_PRIORITY_CLASS
End Sub

Private Sub optReal_Click()
    lngThreadPriority = REALTIME_PRIORITY_CLASS
End Sub
