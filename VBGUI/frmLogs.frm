VERSION 5.00
Begin VB.Form frmLogs 
   Caption         =   "Log options"
   ClientHeight    =   2910
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6930
   LinkTopic       =   "Form2"
   ScaleHeight     =   2910
   ScaleWidth      =   6930
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5520
      TabIndex        =   8
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   4080
      TabIndex        =   7
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdBrowseChange 
      Caption         =   "..."
      Height          =   375
      Left            =   6360
      TabIndex        =   6
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton cmdBrowseAnalysis 
      Caption         =   "..."
      Height          =   375
      Left            =   6360
      TabIndex        =   4
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox txtChangeLog 
      Height          =   285
      Left            =   1680
      TabIndex        =   5
      ToolTipText     =   "Gain changes are logged to this file"
      Top             =   1560
      Width           =   4575
   End
   Begin VB.CheckBox chkChangeLog 
      Caption         =   "Change Log"
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   1320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtAnalysisLog 
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      ToolTipText     =   "Analysis results are logged to this file"
      Top             =   960
      Width           =   4575
   End
   Begin VB.CheckBox chkAnalysisLog 
      Caption         =   "Analysis Log"
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   720
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdBrowseErr 
      Caption         =   "..."
      Height          =   375
      Left            =   6360
      TabIndex        =   2
      Top             =   360
      Width           =   375
   End
   Begin VB.TextBox txtErrLog 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      ToolTipText     =   "Errors are logged to this file"
      Top             =   360
      Width           =   4575
   End
   Begin VB.CheckBox chkErrLog 
      Caption         =   "Error Log"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Change Log"
      Height          =   195
      Left            =   0
      TabIndex        =   13
      ToolTipText     =   "Gain changes are logged to this file"
      Top             =   1560
      Width           =   1575
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Analysis Log"
      Height          =   195
      Left            =   0
      TabIndex        =   12
      ToolTipText     =   "Analysis results are logged to this file"
      Top             =   960
      Width           =   1575
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Error Log"
      Height          =   195
      Left            =   0
      TabIndex        =   11
      ToolTipText     =   "Errors are logged to this file"
      Top             =   360
      Width           =   1575
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmLogs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBrowseClick(txtCur As TextBox, ctlNext As Control)
    On Error Resume Next
    Dim strFileName As String
    Dim strFilter As String
    Dim lngFlags As Long
    Dim strCurPath As String
    Dim strCurFile As String
    
    If Len(txtCur.Text) > 0 Then
        strCurPath = Left$(txtCur.Text, InStrRev(txtCur.Text, "\") - 1)
        strCurFile = Mid$(txtCur.Text, InStrRev(txtCur.Text, "\") + 1)
    Else
        strCurPath = strAppPath
        strCurFile = ""
    End If
    
    lngFlags = ahtOFN_HIDEREADONLY
    strFilter = GetLocalString("frmLogs.LCL_LOG_FILES", "Log files") & _
                    " (*.log)" & vbNullChar & "*.log" & vbNullChar & _
                GetLocalString("frmLogs.LCL_TEXT_FILES", "Text files") & _
                    " (*.txt)" & vbNullChar & "*.txt" & vbNullChar & _
                GetLocalString("frmMain.LCL_OPEN_FILE_FILTER2", "All files") & _
                    " (*.*)" & vbNullChar & "*.*" & vbNullChar
    
    strFileName = ""
    strFileName = ahtCommonFileOpenSave(lngFlags, strCurPath, strFilter, 0, , strCurFile, , Me.hWnd, True)
    If Len(strFileName) > 0 Then
        txtCur.Text = strFileName
        ctlNext.SetFocus
    End If
End Sub

Private Sub cmdBrowseAnalysis_Click()
    Call cmdBrowseClick(txtAnalysisLog, txtChangeLog)
End Sub

Private Sub cmdBrowseChange_Click()
    Call cmdBrowseClick(txtChangeLog, cmdOK)
End Sub

Private Sub cmdBrowseErr_Click()
    Call cmdBrowseClick(txtErrLog, txtAnalysisLog)
End Sub

Private Sub cmdCancel_Click()
    Me.txtAnalysisLog.Text = frmMain.strAnalysisLog
    Me.txtChangeLog.Text = frmMain.strChangeLog
    Me.txtErrLog.Text = frmMain.strErrLog
    Me.Hide
End Sub

Private Function CheckFile(strFile As String) As Boolean
    On Error Resume Next
    Open strFile For Append As #1
    If Err.Number <> 0 Then
        CheckFile = False
    Else
        CheckFile = True
        Close #1
    End If
End Function

Private Sub cmdOK_Click()
On Error GoTo cmdOK_Click_Error

    If Me.chkAnalysisLog.Value = vbChecked Then
        If Not CheckFile(Me.txtAnalysisLog.Text) Then
            MsgBox Replace(GetLocalString("frmLogs.LCL_ANALYSIS_LOG", "Can't open or modify Analysis log file %%filename%%"), "%%filename%%", Me.txtAnalysisLog.Text)
            Exit Sub
        End If
    End If
    If Me.chkChangeLog.Value = vbChecked Then
        If Not CheckFile(Me.txtChangeLog.Text) Then
            MsgBox Replace(GetLocalString("frmLogs.LCL_CHANGE_LOG", "Can't open or modify Change log file %%filename%%"), "%%filename%%", Me.txtChangeLog.Text)
            Exit Sub
        End If
    End If
    If Me.chkErrLog.Value = vbChecked Then
        If Not CheckFile(Me.txtErrLog.Text) Then
            MsgBox Replace(GetLocalString("frmLogs.LCL_ERROR_LOG", "Can't open or modify Error log file %%filename%%"), "%%filename%%", Me.txtErrLog.Text)
            Exit Sub
        End If
    End If
    
    frmMain.strAnalysisLog = Me.txtAnalysisLog.Text
    frmMain.strChangeLog = Me.txtChangeLog.Text
    frmMain.strErrLog = Me.txtErrLog.Text
    
    frmMain.blnAnalysisLog = Me.chkAnalysisLog.Value
    frmMain.blnChangeLog = Me.chkChangeLog.Value
    frmMain.blnErrLog = Me.chkErrLog.Value
    
    Me.Hide
    
    Exit Sub
cmdOK_Click_Error:
    HandleError "cmdOK_Click_Error"
End Sub

Private Sub Form_Activate()
On Error GoTo Form_Activate_Error
    Me.txtAnalysisLog.Text = frmMain.strAnalysisLog
    Me.txtChangeLog.Text = frmMain.strChangeLog
    Me.txtErrLog.Text = frmMain.strErrLog
    
    If frmMain.blnAnalysisLog Then
        Me.chkAnalysisLog.Value = vbChecked
    Else
        Me.chkAnalysisLog.Value = vbUnchecked
    End If
    
    If frmMain.blnChangeLog Then
        Me.chkChangeLog.Value = vbChecked
    Else
        Me.chkChangeLog.Value = vbUnchecked
    End If
    
    If frmMain.blnErrLog Then
        Me.chkErrLog.Value = vbChecked
    Else
        Me.chkErrLog.Value = vbUnchecked
    End If

    Exit Sub
    
Form_Activate_Error:
    HandleError "Form_Activate_Error"
End Sub

Private Sub Form_Load()
    
    Me.Icon = LoadResPicture("APPICON", vbResIcon)
    
    fillCaptions Me
    txtChangeLog.ToolTipText = GetLocalString("frmLogs.txtChangeLog.ToolTipText", txtChangeLog.ToolTipText)
    txtErrLog.ToolTipText = GetLocalString("frmLogs.txtErrorLog.ToolTipText", txtErrLog.ToolTipText)
    txtAnalysisLog.ToolTipText = GetLocalString("frmLogs.txtAnalysisLog.ToolTipText", txtAnalysisLog.ToolTipText)
    Label1.ToolTipText = GetLocalString("frmLogs.Label1.ToolTipText", Label1.ToolTipText)
    Label2.ToolTipText = GetLocalString("frmLogs.Label2.ToolTipText", Label2.ToolTipText)
    Label3.ToolTipText = GetLocalString("frmLogs.Label3.ToolTipText", Label3.ToolTipText)
End Sub

Private Sub txtAnalysisLog_Change()
    If Me.txtAnalysisLog.Text = "" Then
        Me.chkAnalysisLog.Value = vbUnchecked
    Else
        Me.chkAnalysisLog.Value = vbChecked
    End If
End Sub

Private Sub txtAnalysisLog_GotFocus()
    Me.txtAnalysisLog.SelStart = 0
    Me.txtAnalysisLog.SelLength = Len(Me.txtAnalysisLog.Text)
End Sub

Private Sub txtAnalysisLog_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Me.txtChangeLog.SetFocus
End Sub

Private Sub txtChangeLog_Change()
    If Me.txtChangeLog.Text = "" Then
        Me.chkChangeLog.Value = vbUnchecked
    Else
        Me.chkChangeLog.Value = vbChecked
    End If
End Sub

Private Sub txtChangeLog_GotFocus()
    Me.txtChangeLog.SelStart = 0
    Me.txtChangeLog.SelLength = Len(Me.txtChangeLog.Text)
End Sub

Private Sub txtChangeLog_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Me.cmdOK.SetFocus
End Sub

Private Sub txtErrLog_Change()
    If Me.txtErrLog.Text = "" Then
        Me.chkErrLog.Value = vbUnchecked
    Else
        Me.chkErrLog.Value = vbChecked
    End If
End Sub

Private Sub txtErrLog_GotFocus()
    Me.txtErrLog.SelStart = 0
    Me.txtErrLog.SelLength = Len(Me.txtErrLog.Text)
End Sub

Private Sub txtErrLog_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Me.txtAnalysisLog.SetFocus
End Sub
