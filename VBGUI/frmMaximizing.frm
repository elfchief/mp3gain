VERSION 5.00
Begin VB.Form frmMaximizing 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Note on Maximizing"
   ClientHeight    =   2445
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5070
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton OKButton 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "http://replaygain.hydrogenaudio.org/faq_norm.html"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   3615
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "http://mp3gain.sourceforge.net/faq.php#peak"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   1500
      Width           =   3300
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Or go here to download some sound samples that demonstrate that maximizing is not the same as volume normalizing:"
      Height          =   390
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   4695
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   $"frmMaximizing.frx":0000
      Height          =   585
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4785
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmMaximizing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub HyperJump(ByVal URL As String)
   Call ShellExecute(0&, vbNullString, URL, _
      vbNullString, vbNullString, vbNormalFocus)
End Sub

Private Sub Form_Load()
    Dim diff1 As Single
    Dim diff5 As Single
    Dim diff2 As Single
    
    Me.Label2.DragIcon = LoadResPicture("LINK", vbResCursor)
    Me.Label2.MouseIcon = LoadResPicture("LINK", vbResCursor)
    Me.Label4.DragIcon = LoadResPicture("LINK", vbResCursor)
    Me.Label4.MouseIcon = LoadResPicture("LINK", vbResCursor)
    
    diff1 = Label1.Height
    diff5 = Label5.Height
    diff2 = Label2.Width
    
    fillCaptions Me
    
    diff1 = Label1.Height - diff1
    diff5 = Label5.Height - diff5
    diff2 = Label2.Width - diff2
    
    Label2.Top = Label2.Top + diff1
    Label5.Top = Label5.Top + diff1
    Label4.Top = Label4.Top + diff1 + diff5
    OKButton.Top = OKButton.Top + diff1 + diff5
    Me.Height = Me.Height + diff1 + diff5
    
    Me.Width = Me.Width + diff2
End Sub

Private Sub OKButton_Click()
  Unload Me
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   ' If the mouse is over the label, the control
   ' must be in drag mode.  In this case, the
   ' DragDrop event occurs when the mouse is
   ' clicked by the user.  Fire up the URL!
   '
   ' Thanks to Mike Bolser for this observation!
   
   If Source Is Label2 Then
      With Label2
         Call HyperJump(.Caption)
         .Font.Underline = False
         '.ForeColor = vbBlack
      End With
   End If
End Sub

Private Sub Label2_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
   ' If the control is in dragmode, you can detect
   ' MouseLeave easily by observing the State parameter.
   
   If State = vbLeave Then
      With Label2
         .Drag vbEndDrag
         .Font.Underline = False
         '.ForeColor = vbBlack
      End With
   End If
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ' Entering dragmode on the first MouseMove allows
   ' easy detection of MouseLeave.
   
      With Label2
         '.ForeColor = vbBlue
         .Font.Underline = True
         .Drag vbBeginDrag
      End With
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   ' If the mouse is over the label, the control
   ' must be in drag mode.  In this case, the
   ' DragDrop event occurs when the mouse is
   ' clicked by the user.  Fire up the URL!
   '
   ' Thanks to Mike Bolser for this observation!
   
   If Source Is Label4 Then
      With Label4
         Call HyperJump(.Caption)
         .Font.Underline = False
         '.ForeColor = vbBlack
      End With
   End If
End Sub

Private Sub Label4_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
   ' If the control is in dragmode, you can detect
   ' MouseLeave easily by observing the State parameter.
   
   If State = vbLeave Then
      With Label4
         .Drag vbEndDrag
         .Font.Underline = False
         '.ForeColor = vbBlack
      End With
   End If
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ' Entering dragmode on the first MouseMove allows
   ' easy detection of MouseLeave.
   
      With Label4
         '.ForeColor = vbBlue
         .Font.Underline = True
         .Drag vbBeginDrag
      End With
End Sub

