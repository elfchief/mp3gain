VERSION 5.00
Begin VB.Form frmReadOnly 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Read-only file"
   ClientHeight    =   1905
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7260
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdNoAll 
      Caption         =   "N&o to All"
      Height          =   495
      Left            =   4440
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdNo 
      Caption         =   "&No"
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdYesAll 
      Caption         =   "Yes to &All"
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   5880
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "&Yes"
      Default         =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Add file to list anyway?"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   720
      Width           =   6855
   End
   Begin VB.Label lblFile 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   480
      Width           =   6855
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Warning- Read-only file will not be able to be modified:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   6855
   End
End
Attribute VB_Name = "frmReadOnly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public intResponse As Integer

Private Sub CancelButton_Click()
    intResponse = 4
    Me.Hide
End Sub

Private Sub cmdNo_Click()
    intResponse = 2
    Me.Hide
End Sub

Private Sub cmdNoAll_Click()
    intResponse = 3
    Me.Hide
End Sub

Private Sub cmdYes_Click()
    intResponse = 0
    Me.Hide
End Sub

Private Sub cmdYesAll_Click()
    intResponse = 1
    Me.Hide
End Sub

Private Sub Form_Load()
    fillCaptions Me
End Sub
