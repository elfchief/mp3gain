VERSION 5.00
Begin VB.Form frmSysTray 
   Caption         =   "frmSysTray"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmSysTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'this procedure receives the callbacks from the System Tray icon.
    Dim Result As Long
    Dim msg As Long
    'the value of X will vary depending upon the scalemode setting
    msg = X / Screen.TwipsPerPixelX
    Select Case msg
        Case WM_LBUTTONUP        '514 restore form window
            frmMain.WindowState = vbNormal
            Result = SetForegroundWindow(frmMain.hWnd)
            frmMain.Show
        Case WM_LBUTTONDBLCLK    '515 restore form window
            frmMain.WindowState = vbNormal
            Result = SetForegroundWindow(frmMain.hWnd)
            frmMain.Show
        Case WM_RBUTTONUP        '517 display popup menu
            Result = SetForegroundWindow(frmMain.hWnd)
            frmMain.PopupMenu frmMain.mPopupSys
    End Select
End Sub


