VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   5985
   ClientLeft      =   1680
   ClientTop       =   615
   ClientWidth     =   5940
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdDonate 
      Caption         =   "I love this program so much, I'd like to know how to send a donation to the author!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      TabIndex        =   12
      Top             =   4800
      Width           =   2655
   End
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   240
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   240
      Width           =   510
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4200
      TabIndex        =   0
      Top             =   5160
      Width           =   1260
   End
   Begin VB.Label lblTranslation 
      AutoSize        =   -1  'True
      Height          =   15
      Left            =   1080
      TabIndex        =   17
      Top             =   4560
      Visible         =   0   'False
      Width           =   4455
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "http://replaygain.hydrogenaudio.org"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   1320
      Width           =   4575
   End
   Begin VB.Label lblMpglib 
      Caption         =   "http://www.mpg123.de"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   1080
      TabIndex        =   16
      Top             =   3360
      Width           =   4635
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "The back end makes use of a modified version of mpglib. The original version of mpglib can be found at"
      Height          =   390
      Left            =   1080
      TabIndex        =   15
      Top             =   2940
      Width           =   4575
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblBackEndVersion 
      Caption         =   "Version 1.00"
      Height          =   255
      Left            =   3240
      TabIndex        =   14
      Top             =   480
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lblBackEndTitle 
      Caption         =   "Back end (mp3gain.exe)"
      Height          =   255
      Left            =   3240
      TabIndex        =   13
      Top             =   240
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   $"frmAbout.frx":08CA
      Height          =   585
      Left            =   1080
      TabIndex        =   11
      Top             =   3720
      Width           =   4575
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label6 
      Caption         =   "mp3gain@hotmail.com"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   1080
      TabIndex        =   10
      Top             =   2580
      Width           =   4575
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Contact the author, Glen Sawyer, at"
      Height          =   195
      Left            =   1080
      TabIndex        =   9
      Top             =   2340
      Width           =   4575
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label4 
      Caption         =   "http://www.geocities.com/mp3gain"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   1080
      TabIndex        =   8
      Top             =   1980
      Width           =   4575
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Latest version of MP3Gain at"
      Height          =   195
      Left            =   1080
      TabIndex        =   7
      Top             =   1740
      Width           =   4575
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "For more information about ReplayGain, go to"
      Height          =   195
      Left            =   1080
      TabIndex        =   6
      Top             =   1080
      Width           =   4575
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblDescription 
      AutoSize        =   -1  'True
      Caption         =   "Performs ReplayGain analysis of MP3 files."
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1080
      TabIndex        =   2
      Top             =   840
      Width           =   4605
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblTitle 
      Caption         =   "Application Title"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   240
      Width           =   3885
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version"
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   480
      Width           =   3885
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'  frmAbout.frm - MP3Gain "About" window
'
'  Copyright (C) 2003 Glen Sawyer
'
'  This library is free software; you can redistribute it and/or
'  modify it under the terms of the GNU Lesser General Public
'  License as published by the Free Software Foundation; either
'  version 2.1 of the License, or (at your option) any later version.
'
'  This library is distributed in the hope that it will be useful,
'  but WITHOUT ANY WARRANTY; without even the implied warranty of
'  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
'  Lesser General Public License for more details.
'
'  You should have received a copy of the GNU Lesser General Public
'  License along with this library; if not, write to the Free Software
'  Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'  coding by Glen Sawyer (mp3gain@hotmail.com) 735 W 255 N, Orem, UT 84057-4505 USA
'
'  DISCLAIMER: The MP3GainGUI code is ugly. REALLY ugly. It was A) my first major
'     VB program, and B) supposed to be just a small proof-of-concept thing, but
'     it kept growing and growing...
'     So there's a LOT in here that I would do completely differently if I were
'     to start over from scratch.
'

Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub HyperJump(ByVal URL As String)
   Call ShellExecute(0&, vbNullString, URL, _
      vbNullString, vbNullString, vbNormalFocus)
End Sub

Private Sub cmdDonate_Click()
    frmDonate.Show vbModal, Me
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    Dim lngRetVal As Long
    Dim sBlah As String
    Dim intVer As Integer
    Dim intVerLen As Integer
    Dim startDescription As Single
    Dim start2 As Single
    Dim start3 As Single
    Dim start5 As Single
    Dim start7 As Single
    Dim start10 As Single
    Dim startTranslation As Single
    Dim diffDescription As Single
    Dim diff2 As Single
    Dim diff3 As Single
    Dim diff5 As Single
    Dim diff7 As Single
    Dim diff10 As Single
    Dim diffTranslation As Single
    
    Me.Icon = LoadResPicture("APPICON", vbResIcon)
    Me.Label1.DragIcon = LoadResPicture("LINK", vbResCursor)
    Me.Label1.MouseIcon = LoadResPicture("LINK", vbResCursor)
    Me.lblMpglib.DragIcon = LoadResPicture("LINK", vbResCursor)
    Me.lblMpglib.MouseIcon = LoadResPicture("LINK", vbResCursor)
    Me.Label4.DragIcon = LoadResPicture("LINK", vbResCursor)
    Me.Label4.MouseIcon = LoadResPicture("LINK", vbResCursor)
    Me.Label6.DragIcon = LoadResPicture("LINK", vbResCursor)
    Me.Label6.MouseIcon = LoadResPicture("LINK", vbResCursor)
    
    startDescription = lblDescription.Height
    start2 = Label2.Height
    start3 = Label3.Height
    start5 = Label5.Height
    start7 = Label7.Height
    start10 = Label10.Height
    startTranslation = lblTranslation.Height
    
    fillCaptions Me
    
    diffDescription = lblDescription.Height - startDescription
    diff2 = Label2.Height - start2
    diff3 = Label3.Height - start3
    diff5 = Label5.Height - start5
    diff7 = Label7.Height - start7
    diff10 = Label10.Height - start10
    diffTranslation = lblTranslation.Height - startTranslation
    
    Label2.Top = Label2.Top + diffDescription
    Label1.Top = Label1.Top + diffDescription + diff2
    Label3.Top = Label3.Top + diffDescription + diff2
    Label4.Top = Label4.Top + diffDescription + diff2 + diff3
    Label5.Top = Label5.Top + diffDescription + diff2 + diff3
    Label6.Top = Label6.Top + diffDescription + diff2 + diff3 + diff5
    Label7.Top = Label7.Top + diffDescription + diff2 + diff3 + diff5
    lblMpglib.Top = lblMpglib.Top + diffDescription + diff2 + diff3 + diff5 + diff7
    Label10.Top = Label10.Top + diffDescription + diff2 + diff3 + diff5 + diff7
    lblTranslation.Top = lblTranslation.Top + diffDescription + diff2 + diff3 + diff5 + diff7 + diff10
    cmdDonate.Top = cmdDonate.Top + diffDescription + diff2 + diff3 + diff5 + diff7 + diff10 + diffTranslation
    cmdOK.Top = cmdOK.Top + diffDescription + diff2 + diff3 + diff5 + diff7 + diff10 + diffTranslation
    Me.Height = Me.Height + diffDescription + diff2 + diff3 + diff5 + diff7 + diff10 + diffTranslation
    
    If Len(lblTranslation.Caption) > 0 Then lblTranslation.Visible = True
    
    Me.Caption = Replace(GetLocalString("frmAbout.LCL_ABOUT_PROGRAM", "About %%programName%%"), "%%programName%%", App.Title)
    lblVersion.Caption = Replace(GetLocalString("frmAbout.LCL_VERSION_NUMBER", "Version %%versionNumber%%"), "%%versionNumber%%", App.Major & "." & App.Minor & "." & App.Revision)
    lblTitle.Caption = App.Title
    lngRetVal = GetCommandOutput(sBlah, strAppPath & "mp3gain /v", strAppPath, False, True)
    intVer = InStr(LCase$(sBlah), "version")
    If intVer > 0 Then
        intVerLen = Len(Mid$(sBlah, intVer + 8)) - 2
        If intVerLen > 0 Then
            lblBackEndTitle.Visible = True
            lblBackEndVersion.Caption = Replace(GetLocalString("frmAbout.LCL_VERSION_NUMBER", "Version %%versionNumber%%"), "%%versionNumber%%", Mid$(sBlah, intVer + 8, Len(Mid$(sBlah, intVer + 8)) - 2))
            lblBackEndVersion.Visible = True
        End If
    End If
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   ' If the mouse is over the label, the control
   ' must be in drag mode.  In this case, the
   ' DragDrop event occurs when the mouse is
   ' clicked by the user.  Fire up the URL!
   '
   ' Thanks to Mike Bolser for this observation!
   
   If Source Is Label1 Then
      With Label1
         Call HyperJump(.Caption)
         .Font.Underline = False
         '.ForeColor = vbBlack
      End With
   End If
End Sub

Private Sub Label1_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
   ' If the control is in dragmode, you can detect
   ' MouseLeave easily by observing the State parameter.
   
   If State = vbLeave Then
      With Label1
         .Drag vbEndDrag
         .Font.Underline = False
         '.ForeColor = vbBlack
      End With
   End If
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ' Entering dragmode on the first MouseMove allows
   ' easy detection of MouseLeave.
   
      With Label1
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

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   ' If the mouse is over the label, the control
   ' must be in drag mode.  In this case, the
   ' DragDrop event occurs when the mouse is
   ' clicked by the user.  Fire up the URL!
   '
   ' Thanks to Mike Bolser for this observation!
   
   If Source Is Label6 Then
      With Label6
         Call HyperJump("mailto:" & .Caption)
         .Font.Underline = False
         '.ForeColor = vbBlack
      End With
   End If
End Sub

Private Sub Label6_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
   ' If the control is in dragmode, you can detect
   ' MouseLeave easily by observing the State parameter.
   
   If State = vbLeave Then
      With Label6
         .Drag vbEndDrag
         .Font.Underline = False
         '.ForeColor = vbBlack
      End With
   End If
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ' Entering dragmode on the first MouseMove allows
   ' easy detection of MouseLeave.
   
      With Label6
         '.ForeColor = vbBlue
         .Font.Underline = True
         .Drag vbBeginDrag
      End With
End Sub

Private Sub lblMpglib_DragDrop(Source As Control, X As Single, Y As Single)
    If Source Is lblMpglib Then
        Call HyperJump(lblMpglib.Caption)
        lblMpglib.Font.Underline = False
    End If
End Sub

Private Sub lblMpglib_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    If State = vbLeave Then
        lblMpglib.Drag vbEndDrag
        lblMpglib.Font.Underline = False
    End If
End Sub

Private Sub lblMpglib_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblMpglib.Font.Underline = True
    lblMpglib.Drag vbBeginDrag
End Sub

