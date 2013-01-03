VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGetGain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Constant Gain Change"
   ClientHeight    =   3735
   ClientLeft      =   3690
   ClientTop       =   3135
   ClientWidth     =   3735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2206.761
   ScaleMode       =   0  'User
   ScaleWidth      =   3506.963
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.Slider Slider1 
      CausesValidation=   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   873
      _Version        =   393216
      LargeChange     =   4
      Min             =   -8
      Max             =   8
      TextPosition    =   1
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   240
      TabIndex        =   1
      Top             =   3120
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   390
      Left            =   2400
      TabIndex        =   2
      Top             =   3120
      Width           =   1140
   End
   Begin VB.CheckBox chkConstOneChannel 
      Caption         =   "Apply to only one channel"
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   1800
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   240
      TabIndex        =   8
      Top             =   1800
      Width           =   3255
      Begin VB.OptionButton optRight 
         Caption         =   "Channel 2 (Right)"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   2175
      End
      Begin VB.OptionButton optLeft 
         Caption         =   "Channel 1 (Left)"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Value           =   -1  'True
         Width           =   2175
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Select gain change to apply to all files"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label Label4 
      Caption         =   "dB"
      Height          =   255
      Left            =   2160
      TabIndex        =   6
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label lbldBGain 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.0"
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "mp3 gain"
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   1080
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblMp3Gain 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "frmGetGain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'  frmGetGain.frm - MP3GainGUI window for getting "Constant Gain" value
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

Public blnGotGain As Boolean
Public intGainChange As Integer

Private Sub chkConstOneChannel_Click()
    If chkConstOneChannel.Value Then
        If frmMain.blnStereoWarning Then
            frmStereoWarning.Show vbModal, Me
        End If
        Me.optLeft.Enabled = True
        Me.optRight.Enabled = True
    Else
        Me.optLeft.Enabled = False
        Me.optRight.Enabled = False
    End If
End Sub

Private Sub cmdCancel_Click()
    blnGotGain = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    If IsNumeric(Me.lblMp3Gain.Caption) Then
        intGainChange = CInt(Me.lblMp3Gain.Caption)
        blnGotGain = True
    Else
        blnGotGain = False
    End If
    Me.Hide
End Sub

Private Sub Form_Activate()
    Me.Slider1.Value = intGainChange
    Me.lblMp3Gain.Caption = intGainChange
    Me.lbldBGain.Caption = Format$(CDbl(intGainChange) * 5# * Log(2#) / Log(10#), "0.0")
End Sub

Private Sub Form_Load()
    Me.Icon = LoadResPicture("APPICON", vbResIcon)
    fillCaptions Me
    intGainChange = 0
End Sub

Private Sub Slider1_Scroll()
    Me.lblMp3Gain.Caption = Me.Slider1.Value
    Me.lbldBGain.Caption = Format$(CDbl(Me.Slider1.Value) * 5# * Log(2#) / Log(10#), "0.0")
    Me.Slider1.Text = Me.lbldBGain.Caption
End Sub
