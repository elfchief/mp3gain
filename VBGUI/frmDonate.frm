VERSION 5.00
Begin VB.Form frmDonate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Donations"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4950
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label lblDonate 
      AutoSize        =   -1  'True
      Caption         =   $"frmDonate.frx":0000
      Height          =   780
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4575
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgDonate 
      Height          =   465
      Left            =   2040
      Picture         =   "frmDonate.frx":00E1
      Tag             =   $"frmDonate.frx":0446
      Top             =   2040
      Width           =   930
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   $"frmDonate.frx":04E3
      Height          =   585
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   4695
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblEMail 
      Caption         =   "mp3gain@hotmail.com"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1740
      Width           =   4575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   $"frmDonate.frx":0575
      Height          =   780
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   4575
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmDonate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'  frmDonate.frm - MP3Gain "donation" window
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

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim startDonate
    Dim start1
    Dim start3
    Dim diffDonate
    Dim diff1
    Dim diff3
    
    Me.Icon = LoadResPicture("APPICON", vbResIcon)
    Me.lblEMail.DragIcon = LoadResPicture("LINK", vbResCursor)
    Me.lblEMail.MouseIcon = LoadResPicture("LINK", vbResCursor)
    Me.imgDonate.DragIcon = LoadResPicture("LINK", vbResCursor)
    Me.imgDonate.MouseIcon = LoadResPicture("LINK", vbResCursor)
    
    startDonate = lblDonate.Height
    start1 = Label1.Height
    start3 = Label3.Height
    
    fillCaptions Me
    
    diffDonate = lblDonate.Height - startDonate
    diff1 = Label1.Height - start1
    diff3 = Label3.Height - start3
    
    Label1.Top = Label1.Top + diffDonate
    lblEMail.Top = lblEMail.Top + diffDonate + diff1
    imgDonate.Top = imgDonate.Top + diffDonate + diff1
    Label3.Top = Label3.Top + diffDonate + diff1
    cmdOK.Top = cmdOK.Top + diffDonate + diff1 + diff3
    Me.Height = Me.Height + diffDonate + diff1 + diff3
    
End Sub

Private Sub imgDonate_DragDrop(Source As Control, X As Single, Y As Single)
    If Source Is imgDonate Then
        Call HyperJump(imgDonate.Tag)
    End If
End Sub

Private Sub imgDonate_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    If State = vbLeave Then
        imgDonate.Drag vbEndDrag
    End If
End Sub

Private Sub imgDonate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgDonate.Drag vbBeginDrag
End Sub

Private Sub lblEMail_DragDrop(Source As Control, X As Single, Y As Single)
   ' If the mouse is over the label, the control
   ' must be in drag mode.  In this case, the
   ' DragDrop event occurs when the mouse is
   ' clicked by the user.  Fire up the URL!
   '
   ' Thanks to Mike Bolser for this observation!
   
   If Source Is lblEMail Then
      With lblEMail
         Call HyperJump("mailto:" & .Caption)
         .Font.Underline = False
         '.ForeColor = vbBlack
      End With
   End If
End Sub

Private Sub lblEMail_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
   ' If the control is in dragmode, you can detect
   ' MouseLeave easily by observing the State parameter.
   
   If State = vbLeave Then
      With lblEMail
         .Drag vbEndDrag
         .Font.Underline = False
         '.ForeColor = vbBlack
      End With
   End If
End Sub

Private Sub lblEMail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ' Entering dragmode on the first MouseMove allows
   ' easy detection of MouseLeave.
   
      With lblEMail
         '.ForeColor = vbBlue
         .Font.Underline = True
         .Drag vbBeginDrag
      End With
End Sub


