VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "MP3 Gain"
   ClientHeight    =   6405
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9540
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6405
   ScaleWidth      =   9540
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   615
      HelpContextID   =   1008
      Left            =   180
      TabIndex        =   11
      Top             =   1260
      Width           =   4155
      Begin VB.TextBox txtTargetDec 
         Height          =   285
         Left            =   2400
         TabIndex        =   15
         Text            =   "0"
         Top             =   180
         Width           =   255
      End
      Begin VB.TextBox txtTargetInt 
         Alignment       =   1  'Right Justify
         Height          =   285
         HelpContextID   =   1008
         Left            =   1920
         TabIndex        =   14
         Text            =   "89"
         Top             =   180
         Width           =   375
      End
      Begin VB.Label lblDecimal 
         Caption         =   "."
         Height          =   255
         Left            =   2340
         TabIndex        =   16
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "dB  (default 89.0)"
         Height          =   195
         Left            =   2700
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Target ""Normal"" Volume:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1770
      End
   End
   Begin MSComctlLib.ImageList smallHotImageList 
      Left            =   8520
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0DCD
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1849
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1DB7
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":22CD
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":27E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2D1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3216
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList bigHotImageList 
      Left            =   7320
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3711
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3FE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":498B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":53E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5EAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":67A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7016
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7980
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8196
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtProgWatch 
      Height          =   375
      Left            =   6720
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   3270
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar prgFile 
      Height          =   255
      Left            =   1320
      TabIndex        =   7
      Top             =   5500
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ProgressBar prgTot 
      Height          =   255
      Left            =   1320
      TabIndex        =   6
      Top             =   5760
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   7800
      TabIndex        =   2
      Top             =   5070
      Width           =   1455
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   1050
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   9540
      _ExtentX        =   16828
      _ExtentY        =   1852
      ButtonWidth     =   1984
      ButtonHeight    =   1799
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      HelpContextID   =   1009
      Style           =   1
      ImageList       =   "bigHotImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add File(s)"
            Key             =   "addfiles"
            Description     =   "Add mp3 file(s) to the list"
            Object.ToolTipText     =   "Add mp3 file(s) to the list"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add Folder"
            Key             =   "addfolder"
            Description     =   "Add all mp3 files in a folder"
            Object.ToolTipText     =   "Add all mp3 files in a folder"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Track Analysis"
            Key             =   "analysis"
            Description     =   "Do ReplayGain analysis on mp3 files"
            Object.ToolTipText     =   "Do Replay Gain analysis on files"
            ImageIndex      =   1
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   4
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Key             =   "radioanalysis"
                  Text            =   "Track Analysis"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "albumanalysis"
                  Text            =   "Album Analysis"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "clearanalysis"
                  Text            =   "Clear Analysis"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Track Gain"
            Key             =   "gain"
            Description     =   "Apply gain to mp3 files"
            Object.ToolTipText     =   "Save volume changes to files"
            ImageIndex      =   3
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Key             =   "radiogain"
                  Text            =   "Track Gain"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "albumgain"
                  Text            =   "Album Gain"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "constantgain"
                  Text            =   "Constant Gain"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Clear File(s)"
            Key             =   "clearfiles"
            Description     =   "Remove selected file(s) from list"
            Object.ToolTipText     =   "Remove selected file(s) from list"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Clear All"
            Key             =   "clearall"
            Description     =   "Remove all files from list"
            Object.ToolTipText     =   "Remove all files from list"
            ImageIndex      =   9
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtAlbumMonitor 
      Height          =   285
      Left            =   6960
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   3990
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   5070
      Width           =   1455
   End
   Begin MSComctlLib.StatusBar stbStat 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   6150
      Width           =   9540
      _ExtentX        =   16828
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Text            =   "0"
            TextSave        =   "0"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstvMain 
      Height          =   3015
      HelpContextID   =   1007
      Left            =   120
      TabIndex        =   0
      Top             =   1950
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   5318
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDropMode     =   1
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDropMode     =   1
      NumItems        =   0
   End
   Begin VB.Label lblTotProg 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Total progress"
      Height          =   195
      Left            =   210
      TabIndex        =   9
      Top             =   5790
      Width           =   1005
   End
   Begin VB.Label lblFileProg 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "File progress"
      Height          =   195
      Left            =   330
      TabIndex        =   8
      Top             =   5500
      Width           =   885
   End
   Begin VB.Label lblNoUndo 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " NO UNDO "
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   8310
      TabIndex        =   17
      Top             =   1560
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuLoadAnalysis 
         Caption         =   "L&oad Analysis results"
      End
      Begin VB.Menu mnuSaveAnalysis 
         Caption         =   "&Save Analysis results"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSep14 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddFile 
         Caption         =   "Add &Files"
         HelpContextID   =   1003
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuAddFolder 
         Caption         =   "Add Fol&der"
         HelpContextID   =   1004
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "Select &All Files"
         Enabled         =   0   'False
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuSelectNone 
         Caption         =   "Select &No Files"
         Enabled         =   0   'False
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuSelectReverse 
         Caption         =   "In&vert selection"
         Enabled         =   0   'False
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuSep10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClearFiles 
         Caption         =   "&Clear selected file(s)"
         Enabled         =   0   'False
         HelpContextID   =   1005
      End
      Begin VB.Menu mnuClearAll 
         Caption         =   "C&lear All files"
         Enabled         =   0   'False
         HelpContextID   =   1006
      End
      Begin VB.Menu mnuSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuAnalysis 
      Caption         =   "&Analysis"
      HelpContextID   =   1001
      Begin VB.Menu mnuRadio 
         Caption         =   "&Track Analysis"
         Enabled         =   0   'False
         HelpContextID   =   1001
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuAlbum 
         Caption         =   "&Album Analysis"
         Enabled         =   0   'False
         HelpContextID   =   1001
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMaxAmp 
         Caption         =   "&Max No-clip analysis"
         Enabled         =   0   'False
         Shortcut        =   ^M
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClearAnalysis 
         Caption         =   "&Clear Analysis"
         Enabled         =   0   'False
         HelpContextID   =   1001
      End
   End
   Begin VB.Menu mnuGain 
      Caption         =   "&Modify Gain"
      HelpContextID   =   1002
      Begin VB.Menu mnuRadioGain 
         Caption         =   "Apply &Track Gain"
         Enabled         =   0   'False
         HelpContextID   =   1002
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuAlbumGain 
         Caption         =   "Apply &Album Gain"
         Enabled         =   0   'False
         HelpContextID   =   1002
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuConstantGain 
         Caption         =   "Apply &Constant Gain..."
         Enabled         =   0   'False
         HelpContextID   =   1002
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSep11 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMaxNoClipGain 
         Caption         =   "Apply Ma&x No-clip Gain for Each file"
         Enabled         =   0   'False
         Shortcut        =   ^X
         Visible         =   0   'False
      End
      Begin VB.Menu mnuGroupNoClip 
         Caption         =   "Apply Max &No-clip Gain for Album"
         Enabled         =   0   'False
         Shortcut        =   ^L
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep20 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUndoGain 
         Caption         =   "&Undo Gain changes"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      HelpContextID   =   1010
      Begin VB.Menu mnuAlwaysTop 
         Caption         =   "Always on &Top"
         HelpContextID   =   1011
      End
      Begin VB.Menu mnuSelectedFiles 
         Caption         =   "Work on &Selected files only"
      End
      Begin VB.Menu mnuEachAlbum 
         Caption         =   "&Each folder is album"
         Checked         =   -1  'True
         HelpContextID   =   1012
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAddSubs 
         Caption         =   "Add S&ubfolders"
         Checked         =   -1  'True
         HelpContextID   =   1013
      End
      Begin VB.Menu mnuKeepTime 
         Caption         =   "&Preserve file date/time"
      End
      Begin VB.Menu mnuReckless 
         Caption         =   "&No check for Layer I or II"
      End
      Begin VB.Menu mnuDontAddClipping 
         Caption         =   "Don't clip when doing Track Gain"
      End
      Begin VB.Menu mnuTagHead 
         Caption         =   "Ta&gs"
         Begin VB.Menu mnuSkipTags 
            Caption         =   "&Ignore (do not read or write tags)"
         End
         Begin VB.Menu mnuReCalcTags 
            Caption         =   "&Re-calculate (do not read tags)"
         End
         Begin VB.Menu mnuSkipTagsWhileAdding 
            Caption         =   "Don't check while adding files"
         End
         Begin VB.Menu mnuSep22 
            Caption         =   "-"
         End
         Begin VB.Menu mnuDeleteTags 
            Caption         =   "Remove Tags from files"
         End
      End
      Begin VB.Menu mnuLogs 
         Caption         =   "&Logs..."
         HelpContextID   =   1015
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolbar 
         Caption         =   "Tool&bar"
         HelpContextID   =   1014
         Begin VB.Menu mnuToolBarBig 
            Caption         =   "&Big"
            Checked         =   -1  'True
            HelpContextID   =   1014
         End
         Begin VB.Menu mnuToolbarSmall 
            Caption         =   "&Small"
            HelpContextID   =   1014
         End
         Begin VB.Menu mnuToolbarText 
            Caption         =   "&Text only"
            HelpContextID   =   1014
         End
         Begin VB.Menu mnuToolbarNone 
            Caption         =   "&None"
            HelpContextID   =   1014
         End
      End
      Begin VB.Menu mnuFileDisplayOptions 
         Caption         =   "&Filename Display"
         Begin VB.Menu mnuPathWithFile 
            Caption         =   "Show Path\File"
            Checked         =   -1  'True
            HelpContextID   =   1016
         End
         Begin VB.Menu mnuFileOnly 
            Caption         =   "Show File only"
            HelpContextID   =   1016
         End
         Begin VB.Menu mnuPathSepFile 
            Caption         =   "Show Path && File"
            HelpContextID   =   1016
         End
      End
      Begin VB.Menu mnuSysTray 
         Caption         =   "Minimi&ze to Tray"
      End
      Begin VB.Menu mnuBeep 
         Caption         =   "&Beep when finished"
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuResetColumns 
         Caption         =   "Reset Default Column &Widths"
      End
      Begin VB.Menu mnuResetWarnings 
         Caption         =   "Reset ""Warning"" &messages"
      End
      Begin VB.Menu mnuSep7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAdvancedOptions 
         Caption         =   "&Advanced..."
      End
   End
   Begin VB.Menu mnuLanguageList 
      Caption         =   "&Language"
      Visible         =   0   'False
      Begin VB.Menu mnuLanguage 
         Caption         =   "Original (U.S. English)"
         Checked         =   -1  'True
         Index           =   0
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuContents 
         Caption         =   "&Contents"
      End
      Begin VB.Menu mnuSep8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDisclaimer 
         Caption         =   "D&isclaimer..."
      End
      Begin VB.Menu mnuSep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About..."
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "PopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuPopupRadio 
         Caption         =   "Track Analysis"
      End
      Begin VB.Menu mnuPopupAlbum 
         Caption         =   "Album Analysis"
      End
      Begin VB.Menu mnuPopupMaxAmp 
         Caption         =   "Max No-clip analysis"
      End
      Begin VB.Menu mnuSep12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupClearAnalysis 
         Caption         =   "Clear Analysis"
      End
      Begin VB.Menu mnuSep13 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupRadioGain 
         Caption         =   "Apply Track Gain"
      End
      Begin VB.Menu mnuPopupAlbumGain 
         Caption         =   "Apply Album Gain"
      End
      Begin VB.Menu mnuPopupConstantGain 
         Caption         =   "Apply Constant Gain..."
      End
      Begin VB.Menu mnuPopupMaxNoclipGain 
         Caption         =   "Apply Max No-clip Gain for Each file"
      End
      Begin VB.Menu mnuPopupGroupNoclip 
         Caption         =   "Apply Max &No-clip Gain for Album"
      End
      Begin VB.Menu mnuSep21 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupUndoGain 
         Caption         =   "Undo Gain changes"
      End
      Begin VB.Menu mnuSep23 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupRemoveTags 
         Caption         =   "Remove Tags from files"
      End
   End
   Begin VB.Menu mPopupSys 
      Caption         =   "&SysTray"
      Visible         =   0   'False
      Begin VB.Menu mnuPopRestore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu mPopExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mcollOriginalCaptions As Collection
Dim strOrgButtonTip(1 To 8) As String
Dim strOrgButtonMenu(1 To 6) As String
Public strCurLanguageFileName As String
Public sOrder As Boolean
Dim sortColumn As Integer
Dim glErrCount As Long

Dim strAddFolderPath As String
Dim strAddFilesPath As String
Dim blnDoSubDirs As Boolean
Dim sngPrevWidth As Single
Dim sngPrevHeight As Single
Dim blnDoingAlbum As Boolean
Dim blnStartAlbum As Boolean
Dim blnListChangeable As Boolean
Dim flsMaster As Collection
Dim blnCurrentlyProcessing As Boolean
Dim blnExitWhenDone As Boolean
Dim blnAddAllReadOnly As Boolean
Dim blnIgnoreAllReadOnly As Boolean
Dim blnTxtProgChanging As Boolean
Dim intPathFileStat As Integer
Dim intToolBarSize As Integer
Dim blnFirstLoad As Boolean
Dim blnMaxAmpOnly As Boolean
Dim intAnalysisMode As Integer
Dim intGainMode As Integer
Dim sngFormHeight
Dim sngFormWidth
Dim dblGainAdjust As Double
Dim blnTargetIsChanging As Boolean
Public blnRecklessWarning As Boolean
Dim strClipYes As String
Dim strClipMaybe As String
Dim strSaveLogsPath As String
Dim strSaveLogsFile As String
Public intCurLanguage As Integer
Dim sngProgLabelsLeft As Single
Dim sngShowAddedFilesTimer As Single
Dim sngETCStart As Single
Dim FiveMM As Single
Dim blnAddingUndoSpace As Boolean

Enum TagAction
    taUndoGain = &H1
    taDeleteTags = &H2
End Enum

' ListSubItem indices
Dim lsVolume As Integer
Dim lsClip As Integer
Dim lsRadioGain As Integer
Dim lsRadioClip As Integer
Dim lsMaxNoClip As Integer
Dim lsAlbumVolume As Integer
Dim lsAlbumGain As Integer
Dim lsAlbumClip As Integer
Public glPath As Integer
Public glFile As Integer
Dim lsMaxAmp As Integer

Public blnErrLog As Boolean
Public strErrLog As String
Public blnAnalysisLog As Boolean
Public strAnalysisLog As String
Public blnChangeLog As Boolean
Public strChangeLog As String
Public blnStereoWarning As Boolean
Public blnResetWarn As Boolean
Public blnUseTempFiles As Boolean
Public blnShowFileStatus As Boolean
Public blnResetWarnResult As Boolean
Public blnSkipTagsWarn As Boolean


Private Const DEFAULTTARGET = 89

Private Const FIVELOG10TWO = 1.50514997831991

Private Const PATHFILE = 0
Private Const FILEONLY = 1
Private Const PATHSEPFILE = 2

Private Const MINHEIGHTlstvMain = 800
Private Const MINWIDTHlstvMain = 200

Private Const MINFORMWIDTH = 5310
Private Const MINFORMHEIGHT = 4740

Private Const LVM_GETCOLUMNORDERARRAY = &H1000 + 59
Private Const LVM_SETCOLUMNORDERARRAY = &H1000 + 58

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hWnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long


Private Declare Function SetWindowPos Lib "user32" ( _
    ByVal hWnd As Long, _
    ByVal hWndInsertAfter As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal cx As Long, _
    ByVal cy As Long, _
    ByVal wFlags As Long) As Long

Const HH_DISPLAY_TOPIC = &H0

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
  (ByVal hWnd As Long, ByVal lpOperation As String, _
   ByVal lpFile As String, ByVal lpParameters As String, _
   ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" _
  (ByVal hwndCaller As Long, ByVal pszFile As String, _
   ByVal uCommand As Long, ByVal dwData As Long) As Long
   
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" _
  (ByVal lpFileName As String) As Long

Private Sub modTrayToolTip(strNewTip As String)
    nid.szTip = strNewTip & vbNullChar
    If mnuSysTray.Checked And Me.WindowState = vbMinimized Then
        Shell_NotifyIcon NIM_MODIFY, nid
    End If
End Sub

Private Function FormatTimeShow(lngTime As Long) As String
    Dim strOut As String
    
    strOut = Format(Int(lngTime / 3600), "0") & ":"
    strOut = strOut & Format(Int((lngTime Mod 3600) / 60), "00") & ":"
    strOut = strOut & Format(lngTime Mod 60, "00")
    
    FormatTimeShow = strOut
End Function

Private Sub UpdateCaptionPercentage(strPercent As String)
    On Error Resume Next
    Dim sngETC As Single 'estimated time to complete
    Dim sngEstimateTotal As Single
    Dim sngCurrentElapsed As Single
    
    If strPercent = "" Then
        Me.Caption = App.Title
        If (mnuBeep.Checked) Then
            Beep
        End If
    Else
        If Me.Caption = App.Title Then
            sngETCStart = Timer
        Else
            sngCurrentElapsed = Timer - sngETCStart
            If (sngCurrentElapsed < 0) Then 'we passed midnight since we started
                sngCurrentElapsed = sngCurrentElapsed + 86400 'seconds in a 24-hour period
            End If
            sngEstimateTotal = sngCurrentElapsed * ((CDbl(prgTot.Max - prgTot.Min)) / CDbl(prgTot.Value))
            sngETC = sngEstimateTotal - sngCurrentElapsed
            stbStat.Panels(2).Text = FormatTimeShow(CLng(sngETC))
        End If
        
        Me.Caption = App.Title & " [" & strPercent & "%]"
    End If
    modTrayToolTip Me.Caption
End Sub

Private Sub LogErr(strErrMsg As String)
    On Error Resume Next
    Dim intFileNum As Integer
    Dim mbrYesNo As VbMsgBoxResult
    
    intFileNum = FreeFile
    If InStr(strErrMsg, "not a layer III file") > 0 Then
        strErrMsg = strErrMsg & vbCrLf & vbCrLf & _
            Replace(GetLocalString("frmMain.LCL_NO_CHECK", "If you think this is incorrect, you can try enabling the %%noLayerCheckOption%% option"), _
            "%%noLayerCheckOption%%", _
            DeMenufy(GetLocalString("frmMain.mnuReckless.Caption", _
            "No check for Layer I or II")))
    End If
    
    If Not blnErrLog Then
        mbrYesNo = MsgBox(strErrMsg & vbCrLf & vbCrLf & _
            GetLocalString("frmMain.LCL_ENTER_LOG", _
            "Would you like to write these errors to a log file instead of seeing these pop-up messages?") _
            , vbDefaultButton2 Or vbYesNo Or vbExclamation)
        If mbrYesNo = vbYes Then
            If mnuAlwaysTop.Checked Then
                SetWindowPos hWnd, conHwndNoTopmost, 0, 0, 0, 0, &H10 Or &H1 Or &H2
            End If
            
            frmLogs.Show vbModal, Me
            
            If mnuAlwaysTop.Checked Then
                SetWindowPos hWnd, conHwndTopmost, 0, 0, 0, 0, &H10 Or &H1 Or &H2
            End If
        End If
    End If
    
    If blnErrLog Then 'might have been modified since the last time we checked
        glErrCount = glErrCount + 1
        Err.Clear
        Open strErrLog For Append As #intFileNum
        If Err.Number Then
            MsgBox Replace(GetLocalString("frmLogs.LCL_ERROR_LOG", "Can't open or modify Error log file %%filename%%"), "%%filename%%", Me.strErrLog)
            Err.Clear
            Call mnuLogs_Click
        Else
            Print #intFileNum, CStr(Now) & vbTab & strErrMsg
            Close #intFileNum
        End If
    Else
    End If
    
    Exit Sub
LogErr_Error:
    HandleError "LogErr"
End Sub

Private Sub LogChange(strFile As String, intChange As Integer)
    On Error GoTo LogChange_Error
    Dim intFileNum As Integer
    
    intFileNum = FreeFile
    
    If blnChangeLog Then
        On Error Resume Next
        Open strChangeLog For Append As #intFileNum
        If Err.Number Then
            MsgBox Replace(GetLocalString("frmLogs.LCL_CHANGE_LOG", "Can't open or modify Change log file %%filename%%"), "%%filename%%", Me.strChangeLog)
            Err.Clear
            Call mnuLogs_Click
        Else
            Print #intFileNum, CStr(Now) & vbTab & strFile & vbTab & Format$(CDbl(intChange) * 1.505, "0.0")
            Close #intFileNum
        End If
    End If
    
    Exit Sub
    
LogChange_Error:
    HandleError "LogChange"
End Sub

Private Sub LogRadioAnalysis(strFile As String, lngCurrMaxAmp As Double, dblRadio As Double)
    On Error GoTo LogRadioAnalysis_Error
    Dim intFileNum As Integer
    
    intFileNum = FreeFile
    
    If blnAnalysisLog Then
        On Error Resume Next
        Open strAnalysisLog For Append As #intFileNum
        If Err.Number Then
            MsgBox Replace(GetLocalString("frmLogs.LCL_ANALYSIS_LOG", "Can't open or modify Analysis log file %%filename%%"), "%%filename%%", Me.strAnalysisLog)
            Err.Clear
            Call mnuLogs_Click
        Else
        
            Print #intFileNum, CStr(Now) & vbTab & strFile & vbTab & "TrackdB: " & _
                Str(dblRadio) & vbTab & "MaxAmp: " & CStr(lngCurrMaxAmp)
            Close #intFileNum
        End If
    End If
    
    Exit Sub
    
LogRadioAnalysis_Error:
    HandleError "LogRadioAnalysis"
End Sub

Private Sub LogAlbumAnalysis(strFile As String, dblAlbum As Double)
    On Error GoTo LogAlbumAnalysis_Error
    Dim intFileNum As Integer
    
    intFileNum = FreeFile
    
    If blnAnalysisLog Then
        On Error Resume Next
        Open strAnalysisLog For Append As #intFileNum
        If Err.Number Then
            MsgBox Replace(GetLocalString("frmLogs.LCL_ANALYSIS_LOG", "Can't open or modify Analysis log file %%filename%%"), "%%filename%%", Me.strAnalysisLog)
            Err.Clear
            Call mnuLogs_Click
        Else
            Print #intFileNum, CStr(Now) & vbTab & strFile & vbTab & "AlbumdB: " & _
                Str(dblAlbum)
            Close #intFileNum
        End If
    End If
    
    Exit Sub
    
LogAlbumAnalysis_Error:
    HandleError "LogAlbumAnalysis"
End Sub

Public Sub ResetColumnWidths()
    Dim sngV As Single
    Dim sngC As Single
    Dim sngRC As Single
    Dim sngRG As Single
    Dim sngMC As Single
    Dim sngAV As Single
    Dim sngAG As Single
    Dim sngAC As Single
    Dim sngMA As Single
    Dim sngP As Single
    Dim sngF As Single
    Dim sngPF As Single
    Dim sngTot As Single
    Dim sngWidth As Single
    Dim sngSqueeze As Single
    Dim lngCount As Long
    Dim itmX As ListItem
    
    Me.ScaleMode = vbPixels
    
    With lstvMain.ColumnHeaders
        Select Case intPathFileStat
        Case PATHFILE
            sngPF = TextWidth(.Item(1).Text) + FiveMM
            sngP = 0
            sngF = 0
        Case PATHSEPFILE
            sngPF = 0
            sngP = TextWidth(.Item("chPath").Text) + FiveMM
            sngF = TextWidth(.Item("chFile").Text) + FiveMM
        Case FILEONLY
            sngPF = 0
            sngP = 0
            sngF = TextWidth(.Item("chFile").Text) + FiveMM
        End Select
        
        sngV = TextWidth(.Item(lsVolume + 1).Text) + FiveMM
        sngC = TextWidth(.Item(lsClip + 1).Text) + FiveMM
        sngRG = TextWidth(.Item(lsRadioGain + 1).Text) + FiveMM
        sngRC = TextWidth(.Item(lsRadioClip + 1).Text) + FiveMM
        If Me.mnuMaxAmp.Visible Then
            sngMC = TextWidth(.Item(lsMaxNoClip + 1).Text) + FiveMM
        Else
            sngMC = 0
        End If
        sngAV = TextWidth(.Item(lsAlbumVolume + 1).Text) + FiveMM
        sngAG = TextWidth(.Item(lsAlbumGain + 1).Text) + FiveMM
        sngAC = TextWidth(.Item(lsAlbumClip + 1).Text) + FiveMM
        
        sngTot = sngV + sngC + sngRG + sngRC + sngMC + sngAV + sngAG + sngAC
        sngWidth = lstvMain.Width - 22#
        If sngWidth - sngTot < (sngPF + sngP + sngF) Then 'Need full "Path/File" showing...
            sngSqueeze = (sngWidth - (sngPF + sngP + sngF)) / sngTot 'This is how much we need to squeeze each column
            sngV = sngV * sngSqueeze
            sngC = sngC * sngSqueeze
            sngRG = sngRG * sngSqueeze
            sngRC = sngRC * sngSqueeze
            sngMC = sngMC * sngSqueeze
            sngAV = sngAV * sngSqueeze
            sngAG = sngAG * sngSqueeze
            sngAC = sngAC * sngSqueeze
        Else 'expand Path/File to take up available space
            If sngPF > 0 Then
                sngPF = sngWidth - sngTot
            ElseIf sngP > 0 Then
                sngSqueeze = sngWidth - sngTot - sngP - sngF 'this is the extra space
                If sngP > sngF Then
                    If sngSqueeze > sngP - sngF Then
                        sngSqueeze = sngSqueeze - (sngP - sngF)
                        sngF = sngP + (sngSqueeze / 2#)
                        sngP = sngP + (sngSqueeze / 2#)
                    Else
                        sngF = sngF + sngSqueeze
                    End If
                Else
                    If sngSqueeze > sngF - sngP Then
                        sngSqueeze = sngSqueeze - (sngF - sngP)
                        sngP = sngF + (sngSqueeze / 2#)
                        sngF = sngF + (sngSqueeze / 2#)
                    Else
                        sngP = sngP + sngSqueeze
                    End If
                End If
            Else
                sngF = sngWidth - sngTot
            End If
        End If
        
        
        .Item(1).Width = sngPF
        .Item("chPath").Width = sngP
        .Item("chFile").Width = sngF
        
        .Item(lsVolume + 1).Width = sngV
        .Item(lsClip + 1).Width = sngC
        .Item(lsRadioGain + 1).Width = sngRG
        .Item(lsRadioClip + 1).Width = sngRC
        .Item(lsMaxNoClip + 1).Width = sngMC
        .Item(lsAlbumVolume + 1).Width = sngAV
        .Item(lsAlbumGain + 1).Width = sngAG
        .Item(lsAlbumClip + 1).Width = sngAC
        
        .Item(lsMaxAmp + 1).Width = 0
    End With
    
    Me.ScaleMode = vbTwips

End Sub

Private Function PathFileDisplay(intNewOrder As Integer) As Boolean
    On Error GoTo PathFileDisplay_Error
    Dim success As Long
    Dim i As Long
    Dim colCount As Long
    Dim colArray() As Long
    Dim lngPath As Long
    Dim lngFile As Long
    Dim lngPathFile As Long
    Dim tmpWidth
    Dim lngTmpSwap As Long
    Dim intPathIdx As Integer
    Dim intFileIdx As Integer
     
    intPathIdx = lstvMain.ColumnHeaders("chPath").Index
    intFileIdx = lstvMain.ColumnHeaders("chFile").Index
     
    PathFileDisplay = False
     
    If intNewOrder = intPathFileStat Then Exit Function
     
    'create an array of columns totaling
    'the number of columns in the control
    colCount = lstvMain.ColumnHeaders.Count
    ReDim colArray(0 To colCount - 1) As Long
    
    success = SendMessage(lstvMain.hWnd, _
        LVM_GETCOLUMNORDERARRAY, _
        colCount, colArray(0))
     
  
    'if success is non-zero, then the call above succeeded,
    'and the array contains the current order of columns
    'from left to right.  This order was determined by the
    'order the columns were added to the control
    '(typically 0, 1, 2, 3 etc)
  
    If success <> 0 Then
        'find the current positions of Path, FileName, and PathWithFileName columns
        For i = 0 To colCount - 1
            Select Case colArray(i)
            Case 0: lngPathFile = i
            Case intPathIdx - 1: lngPath = i
            Case intFileIdx - 1: lngFile = i
            End Select
        Next
        Select Case intPathFileStat
        Case PATHFILE
            Select Case intNewOrder
            Case PATHSEPFILE
                colArray(lngPathFile) = intPathIdx - 1
                colArray(lngPath) = 0
                If lngFile > lngPathFile Then
                    For i = lngFile To lngPathFile + 1 Step -1
                        colArray(i) = colArray(i - 1)
                    Next i
                Else
                    For i = lngFile To lngPathFile
                        colArray(i) = colArray(i + 1)
                    Next i
                End If
                colArray(lngPathFile + 1) = intFileIdx - 1
                tmpWidth = lstvMain.ColumnHeaders(1).Width / 2
                lstvMain.ColumnHeaders(1).Width = 0
                lstvMain.ColumnHeaders(intPathIdx).Width = tmpWidth
                lstvMain.ColumnHeaders(intFileIdx).Width = tmpWidth
            Case FILEONLY
                colArray(lngPathFile) = intFileIdx - 1
                colArray(lngFile) = 0
                tmpWidth = lstvMain.ColumnHeaders(1).Width
                lstvMain.ColumnHeaders(1).Width = 0
                lstvMain.ColumnHeaders(intPathIdx).Width = 0
                lstvMain.ColumnHeaders(intFileIdx).Width = tmpWidth
            End Select
        Case PATHSEPFILE
            Select Case intNewOrder
            Case PATHFILE
                colArray(lngPathFile) = intPathIdx - 1
                colArray(lngPath) = 0
                tmpWidth = lstvMain.ColumnHeaders(intPathIdx).Width + lstvMain.ColumnHeaders(intFileIdx).Width
                lstvMain.ColumnHeaders(intPathIdx).Width = 0
                lstvMain.ColumnHeaders(intFileIdx).Width = 0
                lstvMain.ColumnHeaders(1).Width = tmpWidth
            Case FILEONLY
                tmpWidth = lstvMain.ColumnHeaders(intPathIdx).Width
                lstvMain.ColumnHeaders(1).Width = 0
                lstvMain.ColumnHeaders(intPathIdx).Width = 0
                lstvMain.ColumnHeaders(intFileIdx).Width = lstvMain.ColumnHeaders(intFileIdx).Width + tmpWidth
            End Select
        Case FILEONLY
            Select Case intNewOrder
            Case PATHFILE
                colArray(lngPathFile) = intFileIdx - 1
                colArray(lngFile) = 0
                tmpWidth = lstvMain.ColumnHeaders(intFileIdx).Width
                lstvMain.ColumnHeaders(intFileIdx).Width = 0
                lstvMain.ColumnHeaders(intPathIdx).Width = 0
                lstvMain.ColumnHeaders(1).Width = tmpWidth
            Case PATHSEPFILE
                colArray(lngFile) = intPathIdx - 1
                If lngPath > lngFile Then
                    For i = lngPath To lngFile + 1 Step -1
                        colArray(i) = colArray(i - 1)
                    Next i
                    colArray(lngFile + 1) = intFileIdx - 1
                Else
                    For i = lngPath To lngFile
                        colArray(i) = colArray(i + 1)
                    Next i
                    colArray(lngFile) = intFileIdx - 1
                End If
                lstvMain.ColumnHeaders(1).Width = 0
                lstvMain.ColumnHeaders(intFileIdx).Width = lstvMain.ColumnHeaders(intFileIdx).Width / 2
                lstvMain.ColumnHeaders(intPathIdx).Width = lstvMain.ColumnHeaders(intFileIdx).Width
            End Select
        End Select
        'change the column order by setting the
        'array to new values
         
        'and call the API to make the change
        success = SendMessage(lstvMain.hWnd, _
            LVM_SETCOLUMNORDERARRAY, _
            colCount, colArray(0))
        If success <> 0 Then
            PathFileDisplay = True
            intPathFileStat = intNewOrder
        End If
        'the columns have changed, but the control needs
        'to redisplay its contents in the new order
        lstvMain.Refresh
         
    End If
   
    Exit Function
   
PathFileDisplay_Error:
    HandleError "PathFileDisplay"
End Function

Private Sub ClearAll()
    On Error GoTo ClearAll_Error
    Dim i As Integer
    
    lstvMain.ListItems.Clear
    stbStat.Panels(3).Text = "0"
    
    For i = flsMaster.Count To 1 Step -1
        flsMaster.Remove i
    Next
    EnableJunk (False)
    Me.txtTargetInt.Enabled = True
    Me.txtTargetDec.Enabled = True
    Me.mnuAddFile.Enabled = True
    Me.mnuAddFolder.Enabled = True
    lstvMain.OLEDragMode = ccOLEDragManual
    lstvMain.OLEDropMode = ccOLEDropManual
    Me.mnuEachAlbum.Enabled = True
    Me.mnuLogs.Enabled = True
    Me.mnuSelectedFiles.Enabled = True
    Me.Toolbar1.Buttons("addfiles").Enabled = True
    Me.Toolbar1.Buttons("addfolder").Enabled = True
    Me.cmdExit.Enabled = True
    Me.mnuLoadAnalysis.Enabled = True
    
    Exit Sub
    
ClearAll_Error:
    HandleError "ClearAll"
End Sub

Private Sub ClearFiles()
    On Error GoTo ClearFiles_Error
    Dim i As Integer
    
    For i = lstvMain.ListItems.Count To 1 Step -1
        If lstvMain.ListItems(i).Selected Then
            flsMaster.Remove lstvMain.ListItems(i).Key
            lstvMain.ListItems.Remove i
        End If
    Next
    
    If lstvMain.ListItems.Count = 0 Then
        EnableJunk (False)
        Me.txtTargetInt.Enabled = True
        Me.txtTargetDec.Enabled = True
        Me.mnuAddFile.Enabled = True
        Me.mnuAddFolder.Enabled = True
        lstvMain.OLEDragMode = ccOLEDragManual
        lstvMain.OLEDropMode = ccOLEDropManual
        Me.mnuEachAlbum.Enabled = True
        Me.mnuLogs.Enabled = True
        Me.mnuSelectedFiles.Enabled = True
        Me.Toolbar1.Buttons("addfiles").Enabled = True
        Me.Toolbar1.Buttons("addfolder").Enabled = True
        Me.cmdExit.Enabled = True
        Me.mnuLoadAnalysis.Enabled = True
    End If
    
    stbStat.Panels(3).Text = lstvMain.ListItems.Count
    Exit Sub
    
ClearFiles_Error:
    HandleError "ClearFiles"
End Sub

Public Function YesNoAllFile(strCaption As String, strError As String, strFile As String, strQuestion As String) As Integer
    On Error GoTo YesNoAllFile_Error
    Dim strSaveCaption As String
    Dim strSaveError As String
    Dim strSaveQuestion As String
    
    strSaveCaption = frmReadOnly.Caption
    strSaveError = frmReadOnly.lblTitle.Caption
    strSaveQuestion = frmReadOnly.Label1.Caption
    
    frmReadOnly.Caption = strCaption
    frmReadOnly.lblTitle.Caption = strError
    frmReadOnly.lblFile.Caption = Replace(strFile, "&", "&&")
    frmReadOnly.Label1.Caption = strQuestion
    
    If mnuAlwaysTop.Checked Then
        SetWindowPos hWnd, conHwndNoTopmost, 0, 0, 0, 0, &H10 Or &H1 Or &H2
    End If
    
    frmReadOnly.Show vbModal, Me
    
    If mnuAlwaysTop.Checked Then
        SetWindowPos hWnd, conHwndTopmost, 0, 0, 0, 0, &H10 Or &H1 Or &H2
    End If
    
    frmReadOnly.Caption = strSaveCaption
    frmReadOnly.lblTitle.Caption = strSaveError
    frmReadOnly.lblFile.Caption = ""
    frmReadOnly.Label1.Caption = strSaveQuestion
    
    YesNoAllFile = frmReadOnly.intResponse
    
    Exit Function
    
YesNoAllFile_Error:
    HandleError ("YesNoAllFile")
End Function

Public Function AddSingleFile(strName As String) As String
    On Error GoTo AddSingleFile_Error

    Dim gitmX As ListItem
    Dim fileInfo As Mp3Info
    Dim lngPathSplit As Long
    Dim blnOkayAddIt As Boolean
    Dim strNewKeyVal As String
    Dim i As Long
    Dim intStrCmp As Integer
    Dim faReadOnlyCheck As VbFileAttribute
    Dim strCmdCheckTag As String
    Dim lngRetVal As Long
    Dim strTagInfo As String
    Dim intLF As Integer
    Dim arrVals() As String
    Dim blnHaveTag As Boolean
    
    AddSingleFile = ""
    
    If blnCancel Then Exit Function
    
    blnOkayAddIt = False
    
    On Error Resume Next
    faReadOnlyCheck = (GetAttr(strName) And vbReadOnly)
    If Err.Number Then Exit Function 'Can't check file
    On Error GoTo AddSingleFile_Error
    
    If faReadOnlyCheck = vbReadOnly Then
        If Not blnIgnoreAllReadOnly And Not blnAddAllReadOnly Then
            frmReadOnly.lblFile.Caption = Replace(strName, "&", "&&")

            If mnuAlwaysTop.Checked Then
                SetWindowPos hWnd, conHwndNoTopmost, 0, 0, 0, 0, &H10 Or &H1 Or &H2
            End If
            
            frmReadOnly.Show vbModal, Me

            If mnuAlwaysTop.Checked Then
                SetWindowPos hWnd, conHwndTopmost, 0, 0, 0, 0, &H10 Or &H1 Or &H2
            End If
            
            Select Case frmReadOnly.intResponse
            Case 0 'Yes
                blnOkayAddIt = True
            Case 1 'Yes to all
                blnOkayAddIt = True
                blnAddAllReadOnly = True
            Case 2 'No
            Case 3 'No to all
                blnIgnoreAllReadOnly = True
            Case 4
                blnCancel = True
                Exit Function
            End Select
        ElseIf blnAddAllReadOnly Then
            blnOkayAddIt = True
        End If
    Else
        blnOkayAddIt = True
    End If
    
    If blnOkayAddIt Then
        If (Timer - sngShowAddedFilesTimer > 1#) Then
            sngShowAddedFilesTimer = Timer
            DoEvents
        End If
        
        strNewKeyVal = LCase$(strName)
        
        On Error Resume Next
        Set gitmX = lstvMain.ListItems.Add(, strNewKeyVal, strName)
        If Err.Number = 0 Then
            On Error GoTo AddSingleFile_Error
            
            stbStat.Panels(3).Text = lstvMain.ListItems.Count
            
            lngPathSplit = InStrRev(strName, "\")
            
            gitmX.ListSubItems.Add lsVolume, "Volume"
            gitmX.ListSubItems.Add lsClip, "Clip"
            gitmX.ListSubItems.Add lsRadioGain, "RadioGain"
            gitmX.ListSubItems.Add lsRadioClip, "RadioClip"
            gitmX.ListSubItems.Add lsMaxNoClip, "MaxNoClip"
            gitmX.ListSubItems.Add lsAlbumVolume, "AlbumVolume"
            gitmX.ListSubItems.Add lsAlbumGain, "AlbumGain"
            gitmX.ListSubItems.Add lsAlbumClip, "AlbumClip"
            gitmX.ListSubItems.Add glPath, "Path", Left$(strName, lngPathSplit)
            gitmX.ListSubItems.Add glFile, "File", Mid$(strName, lngPathSplit + 1)
            gitmX.ListSubItems.Add lsMaxAmp, "MaxAmp"
            
            Set fileInfo = New Mp3Info
            
            blnHaveTag = False
            
            strCmdCheckTag = """" & strAppPath & "mp3Gain"" /o /s c """ & strName & """"
            
            If mnuSkipTags.Checked Or mnuReCalcTags.Checked Or mnuSkipTagsWhileAdding.Checked Then
                lngRetVal = 0
            Else
                If blnShowFileStatus Then
                    lngRetVal = GetCommandOutput(strTagInfo, strCmdCheckTag, strAppPath, True, False, False, 100, , Me.txtProgWatch, False)
                Else
                    lngRetVal = GetCommandOutput(strTagInfo, strCmdCheckTag, strAppPath, True, False, False, 100, , , False)
                End If
            End If
            
            If (lngRetVal = 1) Then
                
                intLF = InStr(strTagInfo, vbCrLf)
                arrVals = Split(Mid$(strTagInfo, intLF + 2), vbTab)
                
                If (UBound(arrVals) = 10) Then
                    If (IsNumeric(arrVals(2))) Then
                        fileInfo.RadiodBGain = CDbl(Val(arrVals(2)))
                        blnHaveTag = True
                    End If
                    If (IsNumeric(arrVals(3))) Then
                        fileInfo.CurrMaxAmp = CDbl(Val(arrVals(3)))
                        blnHaveTag = True
                    End If
                    If (IsNumeric(arrVals(4))) Then
                        fileInfo.CurrMaxGain = CInt(Val(arrVals(4)))
                        blnHaveTag = True
                    End If
                    If (IsNumeric(arrVals(5))) Then
                        fileInfo.CurrMinGain = CInt(Val(arrVals(5)))
                        blnHaveTag = True
                    End If
                    If (IsNumeric(arrVals(7))) Then
                        fileInfo.AlbumdBGain = CDbl(Val(arrVals(7)))
                        blnHaveTag = True
                    End If
                End If
            End If
            
            fileInfo.ModifydBGain = dblGainAdjust
            flsMaster.Add fileInfo, strNewKeyVal
            
            If blnHaveTag Then DispJunk gitmX, fileInfo
            
            Set fileInfo = Nothing
            AddSingleFile = strNewKeyVal
        ElseIf Err.Number = 35602 Then 'No problem-- we already have this file in the list
            AddSingleFile = strNewKeyVal
        Else
            HandleError "AddSingleFile"
            Exit Function
        End If
        If mnuSelectedFiles.Checked Then
            gitmX.Selected = True
        Else
            gitmX.Selected = False
        End If
        
        Set gitmX = Nothing
        On Error GoTo AddSingleFile_Error
        
        If Not mnuAlbum.Enabled Then EnableJunk (True)
        
    End If
    Exit Function
    
AddSingleFile_Error:
    HandleError "AddSingleFile"
End Function

    Static Function Log10(X)
    Log10 = Log(X) / 2.30258509299405  'Log(10#)
End Function

Public Sub DispJunk(itmX As ListItem, mp3Inf As Mp3Info)
    On Error GoTo DispJunk_Error
    itmX.ForeColor = vbBlack
    itmX.ListSubItems(glPath).ForeColor = vbBlack
    itmX.ListSubItems(glPath) = itmX.ListSubItems(glPath)
    itmX.ListSubItems(glFile).ForeColor = vbBlack
    itmX.ListSubItems(glFile) = itmX.ListSubItems(glFile)
    If mp3Inf.RadiodBGain <> NOREALNUM Then
        itmX.ListSubItems(lsVolume).Text = Format$(DEFAULTTARGET + mp3Inf.ModifydBGain - mp3Inf.RadiodBGain, "0.0")
        itmX.ListSubItems(lsRadioGain).Text = Format$(CDbl(mp3Inf.RadioMp3Gain) * 1.505, "0.0")
    Else
        itmX.ListSubItems(lsVolume).Text = ""
        itmX.ListSubItems(lsRadioGain).Text = ""
    End If

    If mp3Inf.CurrMaxAmp <> NOREALNUM Then
        itmX.ListSubItems(lsMaxAmp) = Round(mp3Inf.CurrMaxAmp)
        If (mp3Inf.CurrMaxAmp < 1000000) Then
            itmX.ListSubItems(lsMaxNoClip) = Format$(CDbl(mp3Inf.MaxNoclipMp3Gain) * 1.505, "0.0")
            If mp3Inf.MaxNoclipMp3Gain < 0 Then
                itmX.ListSubItems(lsClip) = strClipYes
                itmX.ListSubItems(lsClip).ForeColor = vbRed
                itmX.ListSubItems(lsMaxNoClip).ForeColor = vbRed
                itmX.ForeColor = vbRed
                itmX.ListSubItems(glPath).ForeColor = vbRed
                itmX.ListSubItems(glPath) = itmX.ListSubItems(glPath)
                itmX.ListSubItems(glFile).ForeColor = vbRed
                itmX.ListSubItems(glFile) = itmX.ListSubItems(glFile)
            Else
                itmX.ListSubItems(lsMaxNoClip).ForeColor = vbBlack
                itmX.ListSubItems(lsClip) = ""
                itmX.ListSubItems(lsClip).ForeColor = vbBlack
            End If
        Else
            itmX.ListSubItems(lsMaxNoClip) = strClipMaybe
            itmX.ListSubItems(lsMaxNoClip).ForeColor = vbBlue
            itmX.ListSubItems(lsClip) = strClipMaybe
            itmX.ListSubItems(lsClip).ForeColor = vbBlue
            itmX.ForeColor = vbBlue
            itmX.ListSubItems(glPath).ForeColor = vbBlue
            itmX.ListSubItems(glPath) = itmX.ListSubItems(glPath)
            itmX.ListSubItems(glFile).ForeColor = vbBlue
            itmX.ListSubItems(glFile) = itmX.ListSubItems(glFile)
        End If
        
        If mp3Inf.RadiodBGain <> NOREALNUM Then
            itmX.ListSubItems(lsRadioClip).Text = mp3Inf.CurrMaxAmp * 2# ^ (CDbl(mp3Inf.RadioMp3Gain) / 4#)
            If mp3Inf.CurrMaxAmp * 2# ^ (CDbl(mp3Inf.RadioMp3Gain) / 4#) > 32767# Then
                If mp3Inf.CurrMaxAmp * 2# ^ ((83# - (DEFAULTTARGET + mp3Inf.ModifydBGain - mp3Inf.RadiodBGain)) / 6.0206) > 32767# Then
                    itmX.ListSubItems(lsRadioClip).Text = strClipMaybe
                    itmX.ListSubItems(lsRadioClip).ForeColor = vbBlue
                    itmX.ForeColor = vbBlue
                    itmX.ListSubItems(glPath).ForeColor = vbBlue
                    itmX.ListSubItems(glPath) = itmX.ListSubItems(glPath)
                    itmX.ListSubItems(glFile).ForeColor = vbBlue
                    itmX.ListSubItems(glFile) = itmX.ListSubItems(glFile)
                Else
                    itmX.ListSubItems(lsRadioClip).Text = strClipYes
                    itmX.ListSubItems(lsRadioClip).ForeColor = vbRed
                    itmX.ForeColor = vbRed
                    itmX.ListSubItems(glPath).ForeColor = vbRed
                    itmX.ListSubItems(glPath) = itmX.ListSubItems(glPath)
                    itmX.ListSubItems(glFile).ForeColor = vbRed
                    itmX.ListSubItems(glFile) = itmX.ListSubItems(glFile)
                End If
            Else
                itmX.ListSubItems(lsRadioClip).Text = ""
                itmX.ListSubItems(lsRadioClip).ForeColor = vbBlack
            End If
        End If
        
        If mp3Inf.AlbumdBGain <> NOREALNUM Then
            If mp3Inf.CurrMaxAmp * 2# ^ (CDbl(mp3Inf.AlbumMp3Gain) / 4#) > 32767# Then
                If mp3Inf.CurrMaxAmp * 2# ^ ((83# - (DEFAULTTARGET + mp3Inf.ModifydBGain - mp3Inf.AlbumdBGain)) / 6.0206) > 32767# Then
                    itmX.ListSubItems(lsAlbumClip).Text = strClipMaybe
                    itmX.ListSubItems(lsAlbumClip).ForeColor = vbBlue
                    If itmX.ForeColor = vbBlack Then 'Don't change to blue if already red
                        itmX.ForeColor = vbBlue
                        itmX.ListSubItems(glPath).ForeColor = vbBlue
                        itmX.ListSubItems(glPath) = itmX.ListSubItems(glPath)
                        itmX.ListSubItems(glFile).ForeColor = vbBlue
                        itmX.ListSubItems(glFile) = itmX.ListSubItems(glFile)
                    End If
                Else
                    itmX.ListSubItems(lsAlbumClip).Text = strClipYes
                    itmX.ListSubItems(lsAlbumClip).ForeColor = vbRed
                    itmX.ForeColor = vbRed
                    itmX.ListSubItems(glPath).ForeColor = vbRed
                    itmX.ListSubItems(glPath) = itmX.ListSubItems(glPath)
                    itmX.ListSubItems(glFile).ForeColor = vbRed
                    itmX.ListSubItems(glFile) = itmX.ListSubItems(glFile)
                End If
            Else
                itmX.ListSubItems(lsAlbumClip).Text = ""
                itmX.ListSubItems(lsAlbumClip).ForeColor = vbBlack
            End If
        End If
    Else
        itmX.ListSubItems(lsRadioClip).Text = ""
        itmX.ListSubItems(lsRadioClip).ForeColor = vbBlack
        itmX.ListSubItems(lsMaxNoClip).Text = ""
        itmX.ListSubItems(lsMaxNoClip).ForeColor = vbBlack
        itmX.ListSubItems(lsClip) = ""
        itmX.ListSubItems(lsClip).ForeColor = vbBlack
        itmX.ListSubItems(lsMaxAmp).Text = ""
        itmX.ListSubItems(lsAlbumClip).Text = ""
        itmX.ListSubItems(lsAlbumClip).ForeColor = vbBlack
    End If

    If mp3Inf.AlbumdBGain <> NOREALNUM Then
        itmX.ListSubItems(lsAlbumVolume).Text = Format$(DEFAULTTARGET + mp3Inf.ModifydBGain - mp3Inf.AlbumdBGain, "0.0")
        itmX.ListSubItems(lsAlbumGain).Text = Format$(CDbl(mp3Inf.AlbumMp3Gain) * 1.505, "0.0")
    Else
        itmX.ListSubItems(lsAlbumVolume).Text = ""
        itmX.ListSubItems(lsAlbumGain).Text = ""
    End If

    Exit Sub

DispJunk_Error:
    HandleError "DispJunk"
End Sub

Private Sub EnableJunk(blnEnableVal As Boolean)
    On Error GoTo EnableJunk_Error
    
    Me.mnuLoadAnalysis.Enabled = blnEnableVal
    Me.mnuSaveAnalysis.Enabled = blnEnableVal
    Me.txtTargetInt.Enabled = blnEnableVal
    Me.txtTargetDec.Enabled = blnEnableVal
    Me.mnuLogs.Enabled = blnEnableVal
    Me.Toolbar1.Buttons("analysis").Enabled = blnEnableVal
    Me.Toolbar1.Buttons("gain").Enabled = blnEnableVal
    Me.Toolbar1.Buttons("addfiles").Enabled = blnEnableVal
    Me.Toolbar1.Buttons("addfolder").Enabled = blnEnableVal
    Me.Toolbar1.Buttons("clearfiles").Enabled = blnEnableVal
    Me.Toolbar1.Buttons("clearall").Enabled = blnEnableVal
    Me.mnuEachAlbum.Enabled = blnEnableVal
    Me.mnuAddFile.Enabled = blnEnableVal
    Me.mnuAddFolder.Enabled = blnEnableVal
    If blnEnableVal Then
        lstvMain.OLEDragMode = ccOLEDragManual
        lstvMain.OLEDropMode = ccOLEDropManual
    Else
        lstvMain.OLEDragMode = ccOLEDragAutomatic
        lstvMain.OLEDropMode = ccOLEDropNone
    End If
    Me.mnuClearAnalysis.Enabled = blnEnableVal
    Me.cmdExit.Enabled = blnEnableVal
    Me.mnuAlbum.Enabled = blnEnableVal
    Me.mnuRadio.Enabled = blnEnableVal
    Me.mnuMaxAmp.Enabled = blnEnableVal
    Me.mnuMaxNoClipGain.Enabled = blnEnableVal
    Me.mnuGroupNoClip.Enabled = blnEnableVal
    Me.mnuAlbumGain.Enabled = blnEnableVal
    Me.mnuRadioGain.Enabled = blnEnableVal
    Me.mnuConstantGain.Enabled = blnEnableVal
    Me.mnuUndoGain.Enabled = blnEnableVal
    Me.mnuSelectAll.Enabled = blnEnableVal
    Me.mnuSelectNone.Enabled = blnEnableVal
    Me.mnuSelectReverse.Enabled = blnEnableVal
    Me.mnuSelectedFiles.Enabled = blnEnableVal
    Me.mnuClearAll.Enabled = blnEnableVal
    Me.mnuClearFiles.Enabled = blnEnableVal
    Me.mnuDeleteTags.Enabled = blnEnableVal

    blnListChangeable = blnEnableVal
    
    Exit Sub
    
EnableJunk_Error:
    HandleError "EnableJunk"
End Sub

Private Sub GainAdjust(dbldB As Double)
    On Error GoTo GainAdjust_Error

    Dim itmX As ListItem
    Dim mp3Inf As Mp3Info
    
    For Each itmX In lstvMain.ListItems
        Set mp3Inf = flsMaster.Item(itmX.Key)
        mp3Inf.ModifydBGain = dbldB
        DispJunk itmX, mp3Inf
        Set mp3Inf = Nothing
    Next
    lstvMain.Refresh
    Exit Sub
    
GainAdjust_Error:
    HandleError "GainAdjust"
End Sub

Private Sub AddFiles()
    On Error GoTo AddFiles_Error

    Dim intStart As Integer
    Dim intNext As Integer
    Dim strPath As String
    Dim strCurFile As String
    Dim strFileName As String
    Dim strFilter As String
    Dim lngFlags As Long
    
    blnIgnoreAllReadOnly = False
    blnAddAllReadOnly = False
    blnCancel = False
    
    strFilter = GetLocalString("frmMain.LCL_OPEN_FILE_FILTER1", "MP3 files/lists") & _
        " (*.mp3;*.m3u)" & vbNullChar & "*.mp3;*.m3u" & vbNullChar & _
        GetLocalString("frmMain.LCL_OPEN_FILE_FILTER2", "All files") & _
        " (*.*)" & vbNullChar & "*.*" & vbNullChar
    
        lngFlags = ahtOFN_ALLOWMULTISELECT Or _
        ahtOFN_EXPLORER Or _
        ahtOFN_LONGNAMES Or _
        ahtOFN_HIDEREADONLY Or _
        ahtOFN_FILEMUSTEXIST
    
    strFileName = ""
    strFileName = ahtCommonFileOpenSave(lngFlags, strAddFilesPath, strFilter, 0, , "", , Me.hWnd, True)
    If Len(strFileName) > 0 Then
        intStart = InStr(1, strFileName, vbNullChar, vbBinaryCompare)
        If intStart = 0 Then
            strAddFilesPath = Left$(strFileName, InStrRev(strFileName, "\"))
            If LCase$(Right$(strFileName, 4)) = ".m3u" Then
                AddM3U (strFileName)
            Else
                AddSingleFile (strFileName)
            End If
        Else
            strPath = Left$(strFileName, intStart - 1)
            If Right$(strPath, 1) <> "\" Then strPath = strPath & "\"
            strAddFilesPath = strPath
            intStart = intStart + 1
            intNext = InStr(intStart, strFileName, vbNullChar, vbBinaryCompare)
            While (intNext > 0) And Not blnCancel
                strCurFile = strPath & Mid$(strFileName, intStart, intNext - intStart)
                If LCase$(Right$(strCurFile, 4)) = ".m3u" Then
                    AddM3U (strCurFile)
                Else
                    AddSingleFile (strCurFile)
                End If
                intStart = intNext + 1
                intNext = InStr(intStart, strFileName, vbNullChar, vbBinaryCompare)
            Wend
            If Not blnCancel Then AddSingleFile (strPath & Mid$(strFileName, intStart))
        End If
    End If

    '    ClearListDupes
    
    Exit Sub
    
AddFiles_Error:
    HandleError "AddFiles"
    On Error Resume Next
    doSortColumn
End Sub

Private Sub AddFolderFiles(strPath As String, colFolderList As Collection)
    On Error GoTo AddFolderFiles_Error
    Dim strFile As String
    Dim strCheck As String
    Dim intYNC As Integer
    Dim faCur As Long
    
    If blnCancel Then Exit Sub
    
    If mnuAddSubs.Checked Then
        On Error Resume Next 'in case the path doesn't exist
            strFile = Dir(strPath, vbNormal Or vbHidden Or _
                vbReadOnly Or vbArchive Or vbSystem Or vbDirectory)
            If Err.Number <> 0 Then strFile = ""
        On Error GoTo AddFolderFiles_Error
        While (strFile <> "") And Not blnCancel
            If (strFile <> ".") And (strFile <> "..") Then
                On Error Resume Next 'in case the folder/file doesn't exist at all
                    faCur = (GetFileAttributes(strPath & strFile) And vbDirectory)
                    If Err.Number <> 0 Then faCur = -1
                On Error GoTo AddFolderFiles_Error
                If faCur = vbDirectory Then
                    colFolderList.Add strPath & strFile & "\"
                Else
                    If (faCur <> -1) And (LCase$(Right$(strFile, 4)) = ".mp3") Then
                        AddSingleFile strPath & strFile
                    End If
                End If
            End If
            strFile = Dir
        Wend
    Else
        strFile = Dir(strPath & "*.mp3", vbNormal Or vbHidden Or _
            vbReadOnly Or vbArchive Or vbSystem)
        While (strFile <> "") And Not blnCancel
            AddSingleFile strPath & strFile
            strFile = Dir
        Wend
    End If
    
    Exit Sub
    
AddFolderFiles_Error:
    HandleError "AddFolderFiles"
    On Error Resume Next
    doSortColumn
End Sub

Private Sub AddFolder()
    On Error GoTo AddFolder_Error
    Dim strPath As String
    Dim colFolderList As Collection
    
    blnCancel = False
    blnCurrentlyProcessing = True
    EnableJunk (False)
    Me.MousePointer = vbArrowHourglass
    Me.cmdCancel.Enabled = True
    Me.cmdCancel.Default = True
    
    strPath = BrowseForFolder(Me, _
        GetLocalString("frmMain.LCL_CHOOSE_FOLDER", "Choose Folder"), _
        strAddFolderPath)
    If Len(strPath) > 0 Then
        strAddFolderPath = strPath
        blnIgnoreAllReadOnly = False
        blnAddAllReadOnly = False
        
        Set colFolderList = New Collection
        
        If Right$(strPath, 1) <> "\" Then
            colFolderList.Add strPath & "\"
        Else
            colFolderList.Add strPath
        End If
    
        While colFolderList.Count > 0
            strPath = colFolderList(1)
            colFolderList.Remove (1)
            AddFolderFiles strPath, colFolderList
        Wend
        
        Set colFolderList = Nothing
    End If
    
    stbStat.Panels(1).Text = ""
    Me.cmdCancel.Default = False
    Me.cmdCancel.Enabled = False
    Me.MousePointer = vbDefault
    EnableJunk (True)
    blnCurrentlyProcessing = False
    If blnExitWhenDone Then Unload Me
    
    Exit Sub
    
AddFolder_Error:
    HandleError "AddFolder"
    Me.MousePointer = vbDefault
End Sub

Private Sub AddM3U(strM3UFile As String)
    On Error GoTo AddM3U_Error
    
    Dim intFileNum As Integer
    Dim strLine As String
    Dim strFile As String
    Dim strM3UPath As String
    Dim strM3UDrive As String
    
    intFileNum = FreeFile
    strM3UPath = Left$(strM3UFile, InStrRev(strM3UFile, "\"))
    strM3UDrive = GetDrivePartThing(strM3UPath)
    
    Open strM3UFile For Input As #intFileNum
    Do While Not EOF(intFileNum)
        Line Input #intFileNum, strLine
        If Left$(strLine, 1) <> "#" Then
            If LCase$(Right$(strLine, 4)) = ".mp3" Then
                strLine = CleanOutRelativePathInfo(strLine)
                If (Mid$(strLine, 2, 1) = ":") Or (Left$(strLine, 2) = "\\") Then
                    AddSingleFile (strLine)
                ElseIf Left$(strLine, 1) = "\" Then
                    AddSingleFile (strM3UDrive & strLine)
                Else
                    AddSingleFile (strM3UPath & strLine)
                End If
            End If
        End If
    Loop
    Close #intFileNum
    
    Exit Sub
    
AddM3U_Error:
    HandleError "AddM3U"
    On Error Resume Next
    doSortColumn
End Sub

Private Function GetLeftPath(strBlah As String) As String
    On Error GoTo GetLeftPath_Error

    GetLeftPath = Left$(strBlah, InStrRev(strBlah, "\"))

    Exit Function
GetLeftPath_Error:
    HandleError "GetLeftPath"
End Function

Private Sub SingleAlbum(strAlbum As String, Optional blnFromGain As Boolean = False)
    On Error GoTo SingleAlbum_Error
    Dim strCmd As String
    Dim subItmX As ListItem
    Dim strBlah As String
    Dim lngRetVal As Long
    
    strCmd = """" & strAppPath & "mp3Gain"" /o "
    
    If mnuKeepTime.Checked Then
        strCmd = strCmd & "/p "
    End If
    
    If Not blnShowFileStatus Then
        strCmd = strCmd & "/q "
    End If
    
    If mnuSkipTags.Checked Then
        strCmd = strCmd & "/s s "
    End If
    
    If mnuReCalcTags.Checked Then
        strCmd = strCmd & "/s r "
    End If
    
    If mnuReckless.Checked Then
        strCmd = strCmd & "/f "
    End If
    
    For Each subItmX In lstvMain.ListItems
        If (Not mnuSelectedFiles.Checked) Or (subItmX.Checked) Then
            If subItmX.ListSubItems(glPath) = strAlbum Then
                If subItmX.Tag <> "X" Then subItmX.Tag = "Y"
                strCmd = strCmd & " """ & subItmX.Text & """"
            Else
                If subItmX.Tag <> "X" Then subItmX.Tag = "N"
            End If
        End If
    Next
    
    strBlah = ""
    Me.txtAlbumMonitor.Text = ""
    
    blnStartAlbum = True
    blnDoingAlbum = True
    stbStat.Panels(1).Text = GetLocalString("frmMain.LCL_DOING_ALBUM", "Doing album analysis...")
    If blnShowFileStatus Then
        lngRetVal = GetCommandOutput(strBlah, strCmd, strAppPath, True, False, False, 100, Me.txtAlbumMonitor, Me.txtProgWatch)
    Else
        lngRetVal = GetCommandOutput(strBlah, strCmd, strAppPath, True, False, False, 100, Me.txtAlbumMonitor)
    End If
    Me.prgFile.Value = 0
    
    If (lngRetVal <> 1) And (Not blnCancel) Then
        If strBlah <> "" Then
            LogErr GetLocalString("frmMain.LCL_BACKEND_ERROR", "Error running mp3gain.exe") & ":" & vbCrLf & strBlah
        Else
            LogErr GetLocalString("frmMain.LCL_BACKEND_ERROR", "Error running mp3gain.exe")
        End If
    End If
    blnDoingAlbum = False
    
    Exit Sub
    
SingleAlbum_Error:
    HandleError "SingleAlbum"
End Sub

Private Sub ClearAlbumResults(strAlbum As String)
    Dim subItmX As ListItem
    Dim subMp3Inf As Mp3Info
    
    For Each subItmX In lstvMain.ListItems
        If (Not mnuSelectedFiles.Checked) Or (subItmX.Checked) Then
            If subItmX.ListSubItems(glPath) = strAlbum Then
                subItmX.Tag = "C" 'Checked
                If subItmX.ListSubItems(lsAlbumVolume) <> "" Then
                    subItmX.ForeColor = vbBlack
                    subItmX.ListSubItems(glPath).ForeColor = vbBlack
                    subItmX.ListSubItems(glPath) = subItmX.ListSubItems(glPath)
                    subItmX.ListSubItems(glFile).ForeColor = vbBlack
                    subItmX.ListSubItems(glFile) = subItmX.ListSubItems(glFile)
                    Set subMp3Inf = flsMaster.Item(subItmX.Key)
                    subMp3Inf.ResetVals
                    Set subMp3Inf = Nothing
                    subItmX.ListSubItems(lsVolume).Text = ""
                    subItmX.ListSubItems(lsRadioGain).Text = ""
                    subItmX.ListSubItems(lsRadioClip).Text = ""
                    subItmX.ListSubItems(lsMaxNoClip).Text = ""
                    subItmX.ListSubItems(lsAlbumVolume).Text = ""
                    subItmX.ListSubItems(lsAlbumGain).Text = ""
                    subItmX.ListSubItems(lsAlbumClip).Text = ""
                    subItmX.ListSubItems(lsMaxAmp).Text = ""
                End If
            End If
        End If
    Next
End Sub

Private Sub Album(Optional blnFromGain As Boolean = False)
    On Error GoTo Album_Error

    Dim strBlah As String
    Dim lngRetVal As Long
    Dim itmX As ListItem
    Dim subItmX As ListItem
    Dim strCmd As String
    Dim strTempPath As String
    Dim mp3Inf As Mp3Info
    Dim subMp3Inf As Mp3Info
    Dim i As Long
    Dim firstSel As Long
    
    blnCancel = False
    blnCurrentlyProcessing = True
    blnAllowProcessCancel = True
    
    EnableJunk (False)
    Me.MousePointer = vbArrowHourglass
    Me.cmdCancel.Enabled = True
    Me.cmdCancel.Default = True
    
    For Each itmX In lstvMain.ListItems
        itmX.Tag = "U" 'Unchecked
        itmX.Checked = itmX.Selected
    Next
    
    If Not blnFromGain Then
        glErrCount = 0
        prgTot.Value = 0
        UpdateCaptionPercentage "0"
        If Not mnuEachAlbum.Checked Then
            If mnuSelectedFiles.Checked Then
                prgTot.Max = 1
                For Each itmX In lstvMain.ListItems
                    If itmX.Selected Then
                        prgTot.Max = prgTot.Max + 1
                        itmX.Checked = True
                    Else
                        itmX.Checked = False
                    End If
                Next
                If prgTot.Max > 1 Then prgTot.Max = prgTot.Max - 1
            Else
                prgTot.Max = lstvMain.ListItems.Count
            End If
        End If
    End If
    
    If mnuEachAlbum.Checked Then
        'Start loop thingie
        For Each itmX In lstvMain.ListItems
            If (Not mnuSelectedFiles.Checked) Or (itmX.Checked) Then
                If itmX.Tag = "U" Then
                    Set mp3Inf = flsMaster.Item(itmX.Key)
                    If mp3Inf.AlbumdBGain = NOREALNUM Then
                        ClearAlbumResults itmX.ListSubItems(glPath)
                    End If
                    Set mp3Inf = Nothing
                End If
            End If
        Next
        
        If Not blnFromGain Then
            prgTot.Max = 1
            For Each itmX In lstvMain.ListItems
                If (Not mnuSelectedFiles.Checked) Or (itmX.Checked) Then
                    If itmX.ListSubItems(lsAlbumVolume).Text = "" Then
                        prgTot.Max = prgTot.Max + 1
                    End If
                End If
            Next
            If prgTot.Max > 1 Then prgTot.Max = prgTot.Max - 1
        End If
        
        strTempPath = ""
        For Each itmX In lstvMain.ListItems
            If (Not mnuSelectedFiles.Checked) Or (itmX.Checked) Then
                If Not blnCancel Then
                    If itmX.ListSubItems(lsAlbumVolume).Text = "" And itmX.Tag <> "X" Then
                        If itmX.ListSubItems(glPath) <> strTempPath Then
                            strTempPath = itmX.ListSubItems(glPath)
                            SingleAlbum strTempPath, blnFromGain
                        End If
                    End If
                End If
            End If
        Next
        'End loop thingie
    Else
        For Each itmX In lstvMain.ListItems
            If (Not mnuSelectedFiles.Checked) Or (itmX.Checked) Then
                Set mp3Inf = flsMaster.Item(itmX.Key)
                If mp3Inf.AlbumdBGain = NOREALNUM Then
                    For Each subItmX In lstvMain.ListItems
                        If (Not mnuSelectedFiles.Checked) Or (subItmX.Checked) Then
                            subItmX.ForeColor = vbBlack
                            subItmX.ListSubItems(glPath).ForeColor = vbBlack
                            subItmX.ListSubItems(glPath) = subItmX.ListSubItems(glPath)
                            subItmX.ListSubItems(glFile).ForeColor = vbBlack
                            subItmX.ListSubItems(glFile) = subItmX.ListSubItems(glFile)
                            Set subMp3Inf = flsMaster.Item(subItmX.Key)
                            subMp3Inf.ResetVals
                            Set subMp3Inf = Nothing
                            subItmX.ListSubItems(lsVolume).Text = ""
                            subItmX.ListSubItems(lsRadioGain).Text = ""
                            subItmX.ListSubItems(lsRadioClip).Text = ""
                            subItmX.ListSubItems(lsMaxNoClip).Text = ""
                            subItmX.ListSubItems(lsAlbumVolume).Text = ""
                            subItmX.ListSubItems(lsAlbumGain).Text = ""
                            subItmX.ListSubItems(lsAlbumClip).Text = ""
                            subItmX.ListSubItems(lsMaxAmp).Text = ""
                        End If
                    Next
                    Set mp3Inf = Nothing
                    Exit For
                End If
                Set mp3Inf = Nothing
            End If
        Next
        
        firstSel = 1
        
        If mnuSelectedFiles.Checked Then
            firstSel = 0
            For i = 1 To lstvMain.ListItems.Count
                If lstvMain.ListItems(i).Checked Then
                    firstSel = i
                    Exit For
                End If
            Next i
        End If
        
        If firstSel > 0 Then
            If lstvMain.ListItems(firstSel).ListSubItems(lsAlbumVolume).Text = "" Then
                If blnFromGain Then
                    If mnuSelectedFiles.Checked Then
                        For Each itmX In lstvMain.ListItems
                            If itmX.Checked Then prgTot.Max = prgTot.Max + 1
                        Next
                    Else
                        prgTot.Max = prgTot.Max + lstvMain.ListItems.Count
                    End If
                End If
                
                strCmd = """" & strAppPath & "mp3Gain"" /o "
                
                If mnuKeepTime.Checked Then
                    strCmd = strCmd & "/p "
                End If
                
                If Not blnShowFileStatus Then
                    strCmd = strCmd & "/q "
                End If
                
                If mnuReckless.Checked Then
                    strCmd = strCmd & "/f "
                End If
                
                If mnuSkipTags.Checked Then
                    strCmd = strCmd & "/s s "
                End If
                
                If mnuReCalcTags.Checked Then
                    strCmd = strCmd & "/s r "
                End If
                
                For Each itmX In lstvMain.ListItems
                    If (Not mnuSelectedFiles.Checked) Or (itmX.Checked) Then
                        If itmX.Tag <> "X" Then itmX.Tag = "Y"
                        strCmd = strCmd & " """ & itmX.Text & """"
                    Else
                        If itmX.Tag <> "X" Then itmX.Tag = "N"
                    End If
                Next
                
                strBlah = ""
                Me.txtAlbumMonitor.Text = ""
                
                blnStartAlbum = True
                blnDoingAlbum = True
                stbStat.Panels(1).Text = GetLocalString("frmMain.LCL_DOING_ALBUM", "Doing album analysis...")
                If blnShowFileStatus Then
                    lngRetVal = GetCommandOutput(strBlah, strCmd, strAppPath, True, False, False, 100, Me.txtAlbumMonitor, Me.txtProgWatch)
                Else
                    lngRetVal = GetCommandOutput(strBlah, strCmd, strAppPath, True, False, False, 100, Me.txtAlbumMonitor)
                End If
                Me.prgFile.Value = 0
                
                If (lngRetVal <> 1) And (Not blnCancel) Then
                    If strBlah <> "" Then
                        LogErr GetLocalString("frmMain.LCL_BACKEND_ERROR", "Error running mp3gain.exe") & ":" & vbCrLf & strBlah
                    Else
                        LogErr GetLocalString("frmMain.LCL_BACKEND_ERROR", "Error running mp3gain.exe")
                    End If
                End If
                blnDoingAlbum = False
            End If
        End If
    End If
    
    stbStat.Panels(1).Text = ""
    Me.cmdCancel.Default = False
    Me.cmdCancel.Enabled = False
    Me.MousePointer = vbDefault
    EnableJunk (True)
    
    If Not blnFromGain Then
        prgTot.Value = 0
        UpdateCaptionPercentage ""
        ShowBulkErrors
    End If
    
    blnCurrentlyProcessing = False
    If blnExitWhenDone Then Unload Me
    
    Exit Sub
    
Album_Error:
    HandleError "Album"
    On Error Resume Next
    blnCurrentlyProcessing = False
    EnableJunk (True)
End Sub

Private Sub AlbumGain()
    On Error GoTo AlbumGain_Error

    Dim itmX As ListItem
    Dim subItmX As ListItem
    Dim dbldB As Double
    Dim blnNeedAnalysis As Boolean
    Dim strBlah As String
    Dim lngRetVal As Long
    Dim mp3Inf As Mp3Info
    Dim strCurrAlbum As String
    Dim strCmd As String
    
    glErrCount = 0
    
    If mnuSelectedFiles.Checked Then
        prgTot.Value = 0
        UpdateCaptionPercentage "0"
        prgTot.Max = 1
        For Each itmX In lstvMain.ListItems
            itmX.Tag = "U"
            If itmX.Selected Then
                prgTot.Max = prgTot.Max + 1
                itmX.Checked = True
            Else
                itmX.Checked = False
            End If
        Next
        If prgTot.Max > 1 Then prgTot.Max = prgTot.Max - 1
    Else
        For Each itmX In lstvMain.ListItems
            itmX.Tag = "U"
        Next
        prgTot.Max = lstvMain.ListItems.Count
    End If
    
    blnCancel = False
    
    blnNeedAnalysis = False
    
    For Each itmX In lstvMain.ListItems
        If itmX.Tag = "U" Then
            If (Not mnuSelectedFiles.Checked) Or (itmX.Checked) Then
                Set mp3Inf = flsMaster.Item(itmX.Key)
                If mp3Inf.AlbumdBGain = NOREALNUM Then
                    If mnuEachAlbum.Checked Then
                        ClearAlbumResults itmX.ListSubItems(glPath)
                    End If
                    blnNeedAnalysis = True
                End If
                Set mp3Inf = Nothing
            End If
        End If
    Next
    
    If mnuEachAlbum.Checked Then
        prgTot.Max = 1
        For Each itmX In lstvMain.ListItems
            If (Not mnuSelectedFiles.Checked) Or (itmX.Checked) Then
                If itmX.ListSubItems(lsAlbumVolume) = "" Then
                    prgTot.Max = prgTot.Max + 2
                Else
                    Set mp3Inf = flsMaster.Item(itmX.Key)
                    If mp3Inf.AlbumMp3Gain <> 0 Then
                        prgTot.Max = prgTot.Max + 1
                    End If
                End If
            End If
        Next
        If prgTot.Max > 1 Then prgTot.Max = prgTot.Max - 1
    End If
    
    If blnNeedAnalysis And Not mnuEachAlbum.Checked Then
        Album True
    End If
        
    blnCurrentlyProcessing = True
    blnAllowProcessCancel = False
    
    EnableJunk (False)
    Me.MousePointer = vbArrowHourglass
    Me.cmdCancel.Enabled = True
    Me.cmdCancel.Default = True
    
    strCurrAlbum = ""
    
    For Each itmX In lstvMain.ListItems
        If (Not mnuSelectedFiles.Checked) Or (itmX.Checked) Then
            If Not blnCancel Then
                Set mp3Inf = flsMaster.Item(itmX.Key)
                If mp3Inf.AlbumdBGain = NOREALNUM And mnuEachAlbum.Checked Then
                    If itmX.ListSubItems(glPath) <> strCurrAlbum Then
                        strCurrAlbum = itmX.ListSubItems(glPath)
                        blnAllowProcessCancel = True
                        SingleAlbum strCurrAlbum, True
                        blnAllowProcessCancel = False
                    End If
                    If mp3Inf.AlbumMp3Gain = 0 Then prgTot.Value = prgTot.Value + 1
                End If
                
                If (mp3Inf.AlbumMp3Gain <> 0) And Not blnCancel Then
                                
                    stbStat.Panels(1).Text = Replace(Replace(GetLocalString("frmMain.LCL_APPLY_GAIN", "Applying gain of %%dbGain%% dB to %%filename%%"), _
                        "%%dbGain%%", Format$(CDbl(mp3Inf.AlbumMp3Gain) * 1.505, "0.0")), _
                        "%%filename%%", itmX.Text)

                    Refresh
                    strBlah = ""
                    strCmd = """" & strAppPath & "mp3gain"" /g " & mp3Inf.AlbumMp3Gain & " "
                    If blnUseTempFiles Then
                        strCmd = strCmd & "/t "
                    End If
                    If Not blnShowFileStatus Then
                        strCmd = strCmd & "/q "
                    End If
                    
                    If mnuKeepTime.Checked Then
                        strCmd = strCmd & "/p "
                    End If
                    
                    If mnuReckless.Checked Then
                        strCmd = strCmd & "/f "
                    End If
                    
                    If mnuSkipTags.Checked Then
                        strCmd = strCmd & "/s s "
                    End If
                    
                    strCmd = strCmd & """" & itmX.Text & """"
                    
                    If blnShowFileStatus Then
                        lngRetVal = GetCommandOutput(strBlah, strCmd, strAppPath, True, True, False, 100, , Me.txtProgWatch)
                    Else
                        lngRetVal = GetCommandOutput(strBlah, strCmd, strAppPath, True, True, False, 100)
                    End If
                    
                    Me.prgFile.Value = 0
                    If InStr(LCase$(strBlah), "not enough temp space on disk") Then
                        LogErr Replace(GetLocalString("frmMain.LCL_NO_TEMP_SPACE_1", _
                            "Not enough temp space on disk to modify %%filename%%"), _
                            "%%filename%%", itmX.Text) & vbCrLf & GetLocalString("frmMain.LCL_NO_TEMP_SPACE_2", _
                            "Either clear space on disk, or go to ""Options->Advanced..."" and check the ""Do not use Temp files"" box.")
                    ElseIf InStr(LCase$(strBlah), "can't open") Then
                        LogErr Replace(GetLocalString("frmMain.LCL_MODIFY_ERROR", _
                            "Can't modify file %%filename%%"), "%%filename%%", itmX.Text)
                    ElseIf lngRetVal <> 1 Then
                        If Not blnCancel Then
                            If strBlah <> "" Then
                                LogErr GetLocalString("frmMain.LCL_BACKEND_ERROR", "Error running mp3gain.exe") & ":" & vbCrLf & strBlah
                            Else
                                LogErr GetLocalString("frmMain.LCL_BACKEND_ERROR", "Error running mp3gain.exe")
                            End If
                        End If
                    Else
                        LogChange itmX.Text, mp3Inf.AlbumMp3Gain
                        mp3Inf.AlterDb -CDbl(mp3Inf.AlbumMp3Gain) * FIVELOG10TWO
                        DispJunk itmX, mp3Inf
                    End If
                    prgTot.Value = prgTot.Value + 1
                End If
                Set mp3Inf = Nothing
            End If
            If prgTot.Max > 100 Then
                UpdateCaptionPercentage Format$(CSng(prgTot.Value) * 100! / CSng(prgTot.Max), "0.0")
            Else
                UpdateCaptionPercentage CLng((prgTot.Value * 100) / prgTot.Max)
            End If
        End If
    Next
    stbStat.Panels(1).Text = ""
    Me.cmdCancel.Default = False
    Me.cmdCancel.Enabled = False
    Me.MousePointer = vbDefault
    EnableJunk (True)
    blnCurrentlyProcessing = False
    
    prgTot.Value = 0
    UpdateCaptionPercentage ""
    
    ShowBulkErrors
    
    If blnExitWhenDone Then Unload Me
    
    Exit Sub
    
AlbumGain_Error:
    HandleError "AlbumGain"
    On Error Resume Next
    blnCurrentlyProcessing = False
    EnableJunk (True)
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo cmdCancel_Click_Error

    blnCancel = True
    stbStat.Panels(1).Text = GetLocalString("frmMain.LCL_CANCELLING", "Cancelling...")
    
    Exit Sub
    
cmdCancel_Click_Error:
    HandleError "cmdCancel_Click"
End Sub

Private Sub ConstGain()
    Dim itmX As ListItem
    
    glErrCount = 0
    
    If mnuAlwaysTop.Checked Then
        SetWindowPos hWnd, conHwndNoTopmost, 0, 0, 0, 0, &H10 Or &H1 Or &H2
    End If
    
    frmGetGain.Show vbModal, Me
    
    If mnuAlwaysTop.Checked Then
        SetWindowPos hWnd, conHwndTopmost, 0, 0, 0, 0, &H10 Or &H1 Or &H2
    End If

    If Not frmGetGain.blnGotGain Then
        Exit Sub
    End If
    
    If frmGetGain.intGainChange = 0 Then
        Exit Sub
    End If
    
    blnCancel = False
    blnCurrentlyProcessing = True
    blnAllowProcessCancel = False
    EnableJunk (False)
    Me.MousePointer = vbArrowHourglass
    Me.cmdCancel.Enabled = True
    Me.cmdCancel.Default = True
    
    prgTot.Value = 0
    prgTot.Max = 1
    For Each itmX In lstvMain.ListItems
        itmX.Tag = "Y"
        If mnuSelectedFiles.Checked Then
            If itmX.Selected Then
                prgTot.Max = prgTot.Max + 1
                itmX.Checked = True
            Else
                itmX.Checked = False
            End If
        Else
            prgTot.Max = prgTot.Max + 1
        End If
    Next
    If prgTot.Max > 1 Then prgTot.Max = prgTot.Max - 1
    
    UpdateCaptionPercentage "0"
    
    Call ApplyConstGain(frmGetGain.intGainChange, frmGetGain.chkConstOneChannel, frmGetGain.optLeft)
    
    prgTot.Value = 0
    UpdateCaptionPercentage ""
    
    For Each itmX In lstvMain.ListItems
        itmX.Tag = ""
    Next
    
    Me.cmdCancel.Default = False
    Me.cmdCancel.Enabled = False
    Me.MousePointer = vbDefault
    EnableJunk (True)
    blnCurrentlyProcessing = False
    
    ShowBulkErrors
    
End Sub


Private Sub ApplyConstGain(intGainChange As Integer, _
        Optional blnSingleChannel As Boolean = False, _
        Optional blnLeftChannel As Boolean = True)
    On Error GoTo ApplyConstGain_Error

    Dim strBlah As String
    Dim strCmd As String
    Dim lngRetVal As Long
    Dim itmX As ListItem
    Dim dbldB As Double
    Dim mp3Inf As Mp3Info
    Dim strDBChange As String
    
    If intGainChange = 0 Then Exit Sub
    
    For Each itmX In lstvMain.ListItems
        If ((Not mnuSelectedFiles.Checked) Or (itmX.Checked)) And (itmX.Tag = "Y") Then
            If Not blnCancel Then
                strCmd = """" & strAppPath & "mp3Gain"" "
                
                strDBChange = Format$(CDbl(intGainChange) * 1.505, "0.0")
                If blnSingleChannel Then
                    If blnLeftChannel Then
                        stbStat.Panels(1).Text = Replace(Replace(GetLocalString("frmMain.LCL_CHANGING_LEFT", _
                            "Changing gain of channel 1 (Left) by %%dbGain%%dB : %%filename%%"), _
                            "%%dbGain%%", strDBChange), "%%filename%%", itmX.Text)
                        strCmd = strCmd & "/l0 " & intGainChange & " "
                    Else
                        stbStat.Panels(1).Text = Replace(Replace(GetLocalString("frmMain.LCL_CHANGING_RIGHT", _
                            "Changing gain of channel 2 (Right) by %%dbGain%%dB : %%filename%%"), _
                            "%%dbGain%%", strDBChange), "%%filename%%", itmX.Text)
                        strCmd = strCmd & "/l1 " & intGainChange & " "
                    End If
                Else
                    stbStat.Panels(1).Text = Replace(Replace(GetLocalString("frmMain.LCL_CHANGING_GAIN", _
                        "Changing gain by %%dbGain%%dB : %%filename%%"), _
                        "%%dbGain%%", strDBChange), "%%filename%%", itmX.Text)
                    strCmd = strCmd & "/g " & intGainChange & " "
                End If
                            
                If blnUseTempFiles Then
                    strCmd = strCmd & "/t "
                End If
                
                If Not blnShowFileStatus Then
                    strCmd = strCmd & "/q "
                End If
                
                If mnuKeepTime.Checked Then
                    strCmd = strCmd & "/p "
                End If
                
                If mnuReckless.Checked Then
                    strCmd = strCmd & "/f "
                End If
                
                If mnuSkipTags.Checked Then
                    strCmd = strCmd & "/s s "
                End If
                
                strCmd = strCmd & """" & itmX.Text & """"
                
                Refresh
                strBlah = ""
                
                If blnShowFileStatus Then
                    lngRetVal = GetCommandOutput(strBlah, strCmd, strAppPath, True, True, False, 100, , Me.txtProgWatch)
                Else
                    lngRetVal = GetCommandOutput(strBlah, strCmd, strAppPath, True, True, False, 100)
                End If
                
                Me.prgFile.Value = 0
                
                If InStr(LCase$(strBlah), "not enough temp space on disk") Then
                    LogErr Replace(GetLocalString("frmMain.LCL_NO_TEMP_SPACE_1", _
                        "Not enough temp space on disk to modify %%filename%%"), _
                        "%%filename%%", itmX.Text) & vbCrLf & GetLocalString("frmMain.LCL_NO_TEMP_SPACE_2", _
                        "Either clear space on disk, or go to ""Options->Advanced..."" and check the ""Do not use Temp files"" box.")
                ElseIf InStr(LCase$(strBlah), "can't open") Then
                    LogErr Replace(GetLocalString("frmMain.LCL_MODIFY_ERROR", _
                        "Can't modify file %%filename%%"), "%%filename%%", itmX.Text)
                ElseIf InStr(LCase$(strBlah), "can't adjust single channel") Then
                    LogErr Replace(GetLocalString("frmMain.LCL_NOT_STEREO", _
                        "%%filename%% is not a stereo or dual-channel mp3"), _
                        "%%filename%%", itmX.Text & vbCrLf)
                ElseIf lngRetVal <> 1 Then
                    If Not blnCancel Then
                        If strBlah <> "" Then
                            LogErr GetLocalString("frmMain.LCL_BACKEND_ERROR", "Error running mp3gain.exe") & ":" & vbCrLf & strBlah
                        Else
                            LogErr GetLocalString("frmMain.LCL_BACKEND_ERROR", "Error running mp3gain.exe")
                        End If
                    End If
                Else
                    LogChange itmX.Text, intGainChange
                    Set mp3Inf = flsMaster.Item(itmX.Key)
                    If mp3Inf.CurrMaxAmp <> NOREALNUM Then
                        mp3Inf.AlterDb -CDbl(intGainChange) * FIVELOG10TWO
                        DispJunk itmX, mp3Inf
                    End If
                    Set mp3Inf = Nothing
                End If
                prgTot.Value = prgTot.Value + 1
                If prgTot.Max > 100 Then
                    UpdateCaptionPercentage Format$(CSng(prgTot.Value) * 100! / CSng(prgTot.Max), "0.0")
                Else
                    UpdateCaptionPercentage CLng((prgTot.Value * 100) / prgTot.Max)
                End If
            End If
        End If
    Next
        
    stbStat.Panels(1).Text = ""
    
    If blnExitWhenDone Then Unload Me
    
    Exit Sub
    
ApplyConstGain_Error:
    HandleError "ApplyConstGain"
    On Error Resume Next
    blnCurrentlyProcessing = False
    EnableJunk (True)
End Sub

Private Sub cmdExit_Click()
    On Error GoTo cmdExit_Click_Error

    If blnCurrentlyProcessing Then
        blnCancel = True
        blnExitWhenDone = True
        stbStat.Panels(1).Text = GetLocalString("frmMain.LCL_EXITING", "Exiting...")
    Else
        Unload Me
    End If
    
    Exit Sub
    
cmdExit_Click_Error:
    HandleError "cmdExit_Click"
End Sub

Private Sub RadioSingleFile(itmX As ListItem, mp3Inf As Mp3Info)
    On Error GoTo RadioSingleFile_Error
    Dim strBlah As String
    Dim strCmd As String
    Dim lngRetVal As Long
    Dim arrTokens() As String
    
    stbStat.Panels(1).Text = Replace(GetLocalString("frmMain.LCL_ANALYZING", _
        "Analyzing %%filename%%"), "%%filename%%", itmX.Text)
    Refresh
    strBlah = ""
    strCmd = """" & strAppPath & "mp3gain"" /o "
    
    If mnuKeepTime.Checked Then
        strCmd = strCmd & "/p "
    End If
    
    If blnMaxAmpOnly Then
        strCmd = strCmd & "/x "
    End If
    
    If mnuReckless.Checked Then
        strCmd = strCmd & "/f "
    End If
    
    If mnuSkipTags.Checked Then
        strCmd = strCmd & "/s s "
    End If
    
    If mnuReCalcTags.Checked Then
        strCmd = strCmd & "/s r "
    End If
    
    
    If blnShowFileStatus Then
        lngRetVal = GetCommandOutput(strBlah, strCmd & """" & itmX.Text & """", strAppPath, True, False, False, 100, , Me.txtProgWatch)
    Else
        lngRetVal = GetCommandOutput(strBlah, strCmd & "/q """ & itmX.Text & """", strAppPath, True, False, False, 100)
    End If
    
    Me.prgFile.Value = 0
    
    strBlah = Replace(strBlah, vbCrLf, vbTab, , , vbBinaryCompare)
    arrTokens = Split(strBlah, vbTab, , vbBinaryCompare)
    If UBound(arrTokens) < 18 Then
        If Not blnCancel Then
            If UBound(arrTokens) >= 4 Then
                If InStr(arrTokens(4), itmX.Text) Then
                    LogErr GetLocalString("frmMain.LCL_ERROR_ANALYZING", "Error while analyzing") & ": " & arrTokens(4)
                Else
                    LogErr Replace(GetLocalString("frmMain.LCL_FILE_ERROR_ANALYZING", _
                        "Error while analyzing %%filename%%"), _
                        "%%filename%%", itmX.Text)
                End If
            Else
                LogErr Replace(GetLocalString("frmMain.LCL_FILE_ERROR_ANALYZING", _
                    "Error while analyzing %%filename%%"), _
                    "%%filename%%", itmX.Text)
            End If
        End If
    Else
        mp3Inf.ResetVals
        mp3Inf.CurrMaxAmp = CDbl(Val(arrTokens(9)))
        If Not blnMaxAmpOnly Then
            mp3Inf.RadiodBGain = CDbl(Val(arrTokens(8)))
            LogRadioAnalysis itmX.Text, Round(CDbl(Val(arrTokens(9)))), _
                CDbl(Val(arrTokens(8)))
        End If
        mp3Inf.CurrMaxGain = CInt(Val(arrTokens(10)))
        mp3Inf.CurrMinGain = CInt(Val(arrTokens(11)))
        mp3Inf.ModifydBGain = dblGainAdjust
        DispJunk itmX, mp3Inf
        
    End If
    
    Exit Sub
RadioSingleFile_Error:
    HandleError "RadioSingleFile"
End Sub

Private Sub Radio()
    On Error GoTo Radio_Error

    Dim itmX As ListItem
    Dim mp3Inf As Mp3Info
    
    blnCancel = False
    blnCurrentlyProcessing = True
    blnAllowProcessCancel = True
    
    EnableJunk (False)
    Me.MousePointer = vbArrowHourglass
    Me.cmdCancel.Enabled = True
    Me.cmdCancel.Default = True
        
    glErrCount = 0
    prgTot.Value = 0
    prgTot.Max = 1
    
    For Each itmX In lstvMain.ListItems
        If mnuSelectedFiles.Checked Then
            If itmX.Selected Then
                itmX.Checked = True
                Set mp3Inf = flsMaster.Item(itmX.Key)
                If blnMaxAmpOnly Then
                    If mp3Inf.CurrMaxAmp = NOREALNUM Then
                        prgTot.Max = prgTot.Max + 1
                    End If
                Else
                    If mp3Inf.RadiodBGain = NOREALNUM Then
                        prgTot.Max = prgTot.Max + 1
                    End If
                End If
                Set mp3Inf = Nothing
            Else
                itmX.Checked = False
            End If
        Else
            Set mp3Inf = flsMaster.Item(itmX.Key)
            If blnMaxAmpOnly Then
                If mp3Inf.CurrMaxAmp = NOREALNUM Then
                    prgTot.Max = prgTot.Max + 1
                End If
            Else
                If mp3Inf.RadiodBGain = NOREALNUM Then
                    prgTot.Max = prgTot.Max + 1
                End If
            End If
            Set mp3Inf = Nothing
        End If
    Next
        
    If prgTot.Max > 1 Then prgTot.Max = prgTot.Max - 1
    
    UpdateCaptionPercentage "0"
    
    For Each itmX In lstvMain.ListItems
        If (Not mnuSelectedFiles.Checked) Or (itmX.Checked) Then
            If Not blnCancel Then
                Set mp3Inf = flsMaster.Item(itmX.Key)
                If blnMaxAmpOnly Then
                    If mp3Inf.CurrMaxAmp = NOREALNUM Then
                        RadioSingleFile itmX, mp3Inf
                        prgTot.Value = prgTot.Value + 1
                    End If
                Else
                    If mp3Inf.RadiodBGain = NOREALNUM Then
                        RadioSingleFile itmX, mp3Inf
                        prgTot.Value = prgTot.Value + 1
                    End If
                End If
                Set mp3Inf = Nothing
            End If
            
            If prgTot.Max > 100 Then
                UpdateCaptionPercentage Format$(CSng(prgTot.Value) * 100! / CSng(prgTot.Max), "0.0")
            Else
                UpdateCaptionPercentage CLng((prgTot.Value * 100) / prgTot.Max)
            End If
        End If
    Next
    stbStat.Panels(1).Text = ""
    Me.cmdCancel.Default = False
    Me.cmdCancel.Enabled = False
    Me.MousePointer = vbDefault
    EnableJunk (True)
    blnCurrentlyProcessing = False
    
    prgTot.Value = 0
    UpdateCaptionPercentage ""
    
    ShowBulkErrors
    
    If blnExitWhenDone Then Unload Me
    
    Exit Sub
    
Radio_Error:
    HandleError "Radio"
    On Error Resume Next
    blnCurrentlyProcessing = False
    EnableJunk (True)
End Sub

Private Sub RadioGain()
    On Error GoTo RadioGain_Error

    Dim itmX As ListItem
    Dim dbldB As Double
    Dim blnNeedAnalysis As Boolean
    Dim strBlah As String
    Dim strCmd As String
    Dim lngRetVal As Long
    Dim mp3Inf As Mp3Info
    Dim intGainChange As Integer
    
    glErrCount = 0
    
    blnCancel = False
    
    prgTot.Value = 0
    UpdateCaptionPercentage "0"
    
    prgTot.Max = 1
    For Each itmX In lstvMain.ListItems
        If mnuSelectedFiles.Checked Then
            If itmX.Selected Then
                itmX.Checked = True
                Set mp3Inf = flsMaster.Item(itmX.Key)
                If mp3Inf.RadiodBGain = NOREALNUM Then
                    prgTot.Max = prgTot.Max + 2
                ElseIf mp3Inf.RadioMp3Gain <> 0 Then
                    prgTot.Max = prgTot.Max + 1
                ElseIf mnuDontAddClipping.Checked And _
                        mp3Inf.MaxNoclipMp3Gain < 0 Then
                    prgTot.Max = prgTot.Max + 1
                End If
                Set mp3Inf = Nothing
            Else
                itmX.Checked = False
            End If
        Else
            Set mp3Inf = flsMaster.Item(itmX.Key)
            If mp3Inf.RadiodBGain = NOREALNUM Then
                prgTot.Max = prgTot.Max + 2
            ElseIf mp3Inf.RadioMp3Gain <> 0 Then
                prgTot.Max = prgTot.Max + 1
            ElseIf mnuDontAddClipping.Checked And _
                    mp3Inf.MaxNoclipMp3Gain < 0 Then
                prgTot.Max = prgTot.Max + 1
            End If
            Set mp3Inf = Nothing
        End If
    Next
    If prgTot.Max > 1 Then prgTot.Max = prgTot.Max - 1
        
    blnCurrentlyProcessing = True
    blnAllowProcessCancel = False
    
    EnableJunk (False)
    Me.MousePointer = vbArrowHourglass
    Me.cmdCancel.Enabled = True
    Me.cmdCancel.Default = True
    
    
    For Each itmX In lstvMain.ListItems
        If (Not mnuSelectedFiles.Checked) Or (itmX.Checked) Then
            If Not blnCancel Then
                Set mp3Inf = flsMaster.Item(itmX.Key)
                If mp3Inf.RadiodBGain = NOREALNUM Then
                    blnAllowProcessCancel = True
                    RadioSingleFile itmX, mp3Inf
                    blnAllowProcessCancel = False
                    If (mp3Inf.RadioMp3Gain = 0) And _
                       ((Not mnuDontAddClipping.Checked) Or mp3Inf.MaxNoclipMp3Gain >= 0) Then
                        prgTot.Value = prgTot.Value + 2
                    Else
                        prgTot.Value = prgTot.Value + 1
                    End If
                End If
                If (mp3Inf.RadioMp3Gain <> 0 Or _
                        (mnuDontAddClipping.Checked And _
                         mp3Inf.MaxNoclipMp3Gain < 0)) _
                    And Not blnCancel Then
                    
                    strBlah = ""
                    intGainChange = mp3Inf.RadioMp3Gain
                    
                    If (mnuDontAddClipping.Checked) Then
                        If (mp3Inf.CurrMaxAmp <> NOREALNUM) Then
                            If (mp3Inf.MaxNoclipMp3Gain < mp3Inf.RadioMp3Gain) Then
                                intGainChange = mp3Inf.MaxNoclipMp3Gain
                            End If
                        End If
                    End If
                    
                    stbStat.Panels(1).Text = Replace(Replace(GetLocalString("frmMain.LCL_APPLY_GAIN", "Applying gain of %%dbGain%% dB to %%filename%%"), _
                        "%%dbGain%%", Format$(CDbl(intGainChange) * 1.505, "0.0")), _
                        "%%filename%%", itmX.Text)

                    Refresh
                    
                    strCmd = """" & strAppPath & "mp3gain"" /g " & intGainChange & " "
                    
                    If blnUseTempFiles Then
                        strCmd = strCmd & "/t "
                    End If
                    If Not blnShowFileStatus Then
                        strCmd = strCmd & "/q "
                    End If
                    
                    If mnuKeepTime.Checked Then
                        strCmd = strCmd & "/p "
                    End If
                    
                    If mnuReckless.Checked Then
                        strCmd = strCmd & "/f "
                    End If
                    
                    If mnuSkipTags.Checked Then
                        strCmd = strCmd & "/s s "
                    End If
                    
                    strCmd = strCmd & """" & itmX.Text & """"
                    
                    If blnShowFileStatus Then
                        lngRetVal = GetCommandOutput(strBlah, strCmd, strAppPath, True, True, False, 100, , Me.txtProgWatch)
                    Else
                        lngRetVal = GetCommandOutput(strBlah, strCmd, strAppPath, True, True, False, 100)
                    End If
                    
                    Me.prgFile.Value = 0
                    If InStr(LCase$(strBlah), "not enough temp space on disk") Then
                        LogErr Replace(GetLocalString("frmMain.LCL_NO_TEMP_SPACE_1", _
                            "Not enough temp space on disk to modify %%filename%%"), _
                            "%%filename%%", itmX.Text) & vbCrLf & GetLocalString("frmMain.LCL_NO_TEMP_SPACE_2", _
                            "Either clear space on disk, or go to ""Options->Advanced..."" and check the ""Do not use Temp files"" box.")
                    ElseIf InStr(LCase$(strBlah), "can't open") Then
                        LogErr Replace(GetLocalString("frmMain.LCL_MODIFY_ERROR", _
                            "Can't modify file %%filename%%"), "%%filename%%", itmX.Text)
                    ElseIf lngRetVal <> 1 Then
                        If Not blnCancel Then
                            If strBlah <> "" Then
                                LogErr GetLocalString("frmMain.LCL_BACKEND_ERROR", "Error running mp3gain.exe") & ":" & vbCrLf & strBlah
                            Else
                                LogErr GetLocalString("frmMain.LCL_BACKEND_ERROR", "Error running mp3gain.exe")
                            End If
                        End If
                    Else
                        LogChange itmX.Text, intGainChange
                        mp3Inf.AlterDb -CDbl(intGainChange) * FIVELOG10TWO
                        DispJunk itmX, mp3Inf
                    End If
                    prgTot.Value = prgTot.Value + 1
                End If
                Set mp3Inf = Nothing
            End If
            
            If prgTot.Max > 100 Then
                UpdateCaptionPercentage Format$(CSng(prgTot.Value) * 100! / CSng(prgTot.Max), "0.0")
            Else
                UpdateCaptionPercentage CLng((prgTot.Value * 100) / prgTot.Max)
            End If
        End If
    Next
    stbStat.Panels(1).Text = ""
    Me.cmdCancel.Default = False
    Me.cmdCancel.Enabled = False
    Me.MousePointer = vbDefault
    EnableJunk (True)
    blnCurrentlyProcessing = False
    
    prgTot.Value = 0
    UpdateCaptionPercentage ""
    
    ShowBulkErrors
    
    If blnExitWhenDone Then Unload Me
    
    Exit Sub
    
RadioGain_Error:
    HandleError "RadioGain"
    On Error Resume Next
    blnCurrentlyProcessing = False
    EnableJunk (True)
End Sub

Private Function SetNextAlbumGroupTags() As Boolean
    Dim blnHaveFolder As Boolean
    Dim itmX As ListItem
    Dim strFolder As String
    
    blnHaveFolder = False
    
    For Each itmX In lstvMain.ListItems
        If itmX.Tag = "Y" Then itmX.Tag = "F"
        If itmX.Tag = "U" Then
            If blnHaveFolder Then
                If itmX.ListSubItems(glPath) = strFolder Then
                    itmX.Tag = "Y"
                End If
            Else
                strFolder = itmX.ListSubItems(glPath)
                blnHaveFolder = True
                itmX.Tag = "Y"
            End If
        End If
    Next
    
    SetNextAlbumGroupTags = blnHaveFolder
    
End Function

Private Sub GroupMaxNoClipGain()
    On Error GoTo GroupMaxNoClipGain_Error
    Dim intMaxNoClip As Integer
    Dim itmX As ListItem
    Dim dbldB As Double
    Dim blnNeedAnalysis As Boolean
    Dim strBlah As String
    Dim strCmd As String
    Dim lngRetVal As Long
    Dim mp3Inf As Mp3Info
    Dim blnDoneLooping As Boolean
    
    blnCancel = False
    blnCurrentlyProcessing = True
    blnAllowProcessCancel = False
    EnableJunk (False)
    Me.MousePointer = vbArrowHourglass
    Me.cmdCancel.Enabled = True
    Me.cmdCancel.Default = True
    
    glErrCount = 0
    
    blnMaxAmpOnly = True
    
    prgTot.Value = 0
    prgTot.Max = 1
    For Each itmX In lstvMain.ListItems
        
        If mnuEachAlbum.Checked Then
            itmX.Tag = "U"  '"U"nprocessed
        Else
            itmX.Tag = "Y"
        End If
        
        If mnuSelectedFiles.Checked Then
            If itmX.Selected Then
                Set mp3Inf = flsMaster.Item(itmX.Key)
                If mp3Inf.CurrMaxAmp = NOREALNUM Then
                    prgTot.Max = prgTot.Max + 1
                End If
                Set mp3Inf = Nothing
                prgTot.Max = prgTot.Max + 1
                itmX.Checked = True
            Else
                itmX.Checked = False
            End If
        Else
            Set mp3Inf = flsMaster.Item(itmX.Key)
            If mp3Inf.CurrMaxAmp = NOREALNUM Then
                prgTot.Max = prgTot.Max + 1
            End If
            Set mp3Inf = Nothing
            prgTot.Max = prgTot.Max + 1
        End If
    Next
    If prgTot.Max > 1 Then prgTot.Max = prgTot.Max - 1
    
    blnDoneLooping = False
    
    UpdateCaptionPercentage "0"
    

    If mnuEachAlbum.Checked Then
        blnDoneLooping = Not SetNextAlbumGroupTags
    End If
    
    While (Not blnDoneLooping) And (Not blnCancel)
        'Scan through list and find maximum noclip gain
    
        'Run Radio Analysis if necessary, just getting Max Amp
        For Each itmX In lstvMain.ListItems
            If ((Not mnuSelectedFiles.Checked) Or (itmX.Checked)) _
                And (itmX.Tag = "Y") And (Not blnCancel) Then
                Set mp3Inf = flsMaster.Item(itmX.Key)
                If mp3Inf.CurrMaxAmp = NOREALNUM Then
                    RadioSingleFile itmX, mp3Inf
                    prgTot.Value = prgTot.Value + 1
                End If
                Set mp3Inf = Nothing
            End If
        Next
        
        'Find Minimum Max Amp for all files
        
        intMaxNoClip = 1000
        
        For Each itmX In lstvMain.ListItems
            If ((Not mnuSelectedFiles.Checked) Or (itmX.Checked)) And (itmX.Tag = "Y") Then
                Set mp3Inf = flsMaster.Item(itmX.Key)
                If mp3Inf.CurrMaxAmp <> NOREALNUM Then
                    If mp3Inf.MaxNoclipMp3Gain < intMaxNoClip Then
                        intMaxNoClip = mp3Inf.MaxNoclipMp3Gain
                    End If
                End If
                Set mp3Inf = Nothing
            End If
        Next
        
        If (intMaxNoClip = 0) Then
            For Each itmX In lstvMain.ListItems
                If ((Not mnuSelectedFiles.Checked) Or (itmX.Checked)) _
                    And (itmX.Tag = "Y") Then
                    prgTot.Value = prgTot.Value + 1
                End If
            Next
        End If
        'Now that we have it, use it
        If (intMaxNoClip <> 1000) And (Not blnCancel) Then
            Call ApplyConstGain(intMaxNoClip)
        End If
        
        If mnuEachAlbum.Checked Then
            blnDoneLooping = Not SetNextAlbumGroupTags
        Else
            blnDoneLooping = True
        End If
        
    Wend
    
    prgTot.Value = 0
    UpdateCaptionPercentage ""
    
    blnMaxAmpOnly = False
    
    ShowBulkErrors
    
    Me.cmdCancel.Default = False
    Me.cmdCancel.Enabled = False
    Me.MousePointer = vbDefault
    EnableJunk (True)
    blnCurrentlyProcessing = False
    
    If blnExitWhenDone Then Unload Me
    
    Exit Sub
    
GroupMaxNoClipGain_Error:
    HandleError "GroupMaxNoClipGain"
    On Error Resume Next
    blnCurrentlyProcessing = False
    EnableJunk (True)
End Sub

Private Sub MaxNoClipGain()
    On Error GoTo MaxNoClipGain_Error

    Dim itmX As ListItem
    Dim dbldB As Double
    Dim blnNeedAnalysis As Boolean
    Dim strBlah As String
    Dim strCmd As String
    Dim lngRetVal As Long
    Dim mp3Inf As Mp3Info
    
    glErrCount = 0
    
    blnCancel = False
    
    prgTot.Value = 0
    
    UpdateCaptionPercentage "0"
    
    prgTot.Max = 1
    For Each itmX In lstvMain.ListItems
        If mnuSelectedFiles.Checked Then
            If itmX.Selected Then
                Set mp3Inf = flsMaster.Item(itmX.Key)
                If mp3Inf.CurrMaxAmp = NOREALNUM Then
                    prgTot.Max = prgTot.Max + 2
                ElseIf mp3Inf.MaxNoclipMp3Gain <> 0 Then
                    prgTot.Max = prgTot.Max + 1
                End If
                Set mp3Inf = Nothing
                itmX.Checked = True
            Else
                itmX.Checked = False
            End If
        Else
            Set mp3Inf = flsMaster.Item(itmX.Key)
            If mp3Inf.CurrMaxAmp = NOREALNUM Then
                prgTot.Max = prgTot.Max + 2
            ElseIf mp3Inf.MaxNoclipMp3Gain <> 0 Then
                prgTot.Max = prgTot.Max + 1
            End If
            Set mp3Inf = Nothing
        End If
    Next
    If prgTot.Max > 1 Then prgTot.Max = prgTot.Max - 1
        
    blnCurrentlyProcessing = True
    blnAllowProcessCancel = False
    
    EnableJunk (False)
    Me.MousePointer = vbArrowHourglass
    Me.cmdCancel.Enabled = True
    Me.cmdCancel.Default = True
    
    
    For Each itmX In lstvMain.ListItems
        If (Not mnuSelectedFiles.Checked) Or (itmX.Checked) Then
            If Not blnCancel Then
                Set mp3Inf = flsMaster.Item(itmX.Key)
                If mp3Inf.CurrMaxAmp = NOREALNUM Then
                    blnAllowProcessCancel = True
                    blnMaxAmpOnly = True
                    RadioSingleFile itmX, mp3Inf
                    blnMaxAmpOnly = False
                    blnAllowProcessCancel = False
                    If mp3Inf.MaxNoclipMp3Gain = 0 Then
                        prgTot.Value = prgTot.Value + 2
                    Else
                        prgTot.Value = prgTot.Value + 1
                    End If
                End If
                If mp3Inf.MaxNoclipMp3Gain <> 0 And Not blnCancel Then
                    stbStat.Panels(1).Text = Replace(Replace(GetLocalString("frmMain.LCL_APPLY_GAIN", "Applying gain of %%dbGain%% dB to %%filename%%"), _
                        "%%dbGain%%", Format$(CDbl(mp3Inf.MaxNoclipMp3Gain) * 1.505, "0.0")), _
                        "%%filename%%", itmX.Text)
                    
                    Refresh
                    strBlah = ""
                    strCmd = """" & strAppPath & "mp3gain"" /g " & mp3Inf.MaxNoclipMp3Gain & " "
                    If blnUseTempFiles Then
                        strCmd = strCmd & "/t "
                    End If
                    If Not blnShowFileStatus Then
                        strCmd = strCmd & "/q "
                    End If
                    
                    If mnuKeepTime.Checked Then
                        strCmd = strCmd & "/p "
                    End If
                    
                    If mnuReckless.Checked Then
                        strCmd = strCmd & "/f "
                    End If
                    
                    If mnuSkipTags.Checked Then
                        strCmd = strCmd & "/s s "
                    End If
                    
                    strCmd = strCmd & """" & itmX.Text & """"
                    
                    If blnShowFileStatus Then
                        lngRetVal = GetCommandOutput(strBlah, strCmd, strAppPath, True, True, False, 100, , Me.txtProgWatch)
                    Else
                        lngRetVal = GetCommandOutput(strBlah, strCmd, strAppPath, True, True, False, 100)
                    End If
                    
                    Me.prgFile.Value = 0
                    If InStr(LCase$(strBlah), "not enough temp space on disk") Then
                        LogErr Replace(GetLocalString("frmMain.LCL_NO_TEMP_SPACE_1", _
                            "Not enough temp space on disk to modify %%filename%%"), _
                            "%%filename%%", itmX.Text) & vbCrLf & GetLocalString("frmMain.LCL_NO_TEMP_SPACE_2", _
                            "Either clear space on disk, or go to ""Options->Advanced..."" and check the ""Do not use Temp files"" box.")
                    ElseIf InStr(LCase$(strBlah), "can't open") Then
                        LogErr Replace(GetLocalString("frmMain.LCL_MODIFY_ERROR", _
                            "Can't modify file %%filename%%"), "%%filename%%", itmX.Text)
                    ElseIf lngRetVal <> 1 Then
                        If Not blnCancel Then
                            If strBlah <> "" Then
                                LogErr GetLocalString("frmMain.LCL_BACKEND_ERROR", "Error running mp3gain.exe") & ":" & vbCrLf & strBlah
                            Else
                                LogErr GetLocalString("frmMain.LCL_BACKEND_ERROR", "Error running mp3gain.exe")
                            End If
                        End If
                    Else
                        LogChange itmX.Text, mp3Inf.MaxNoclipMp3Gain
                        mp3Inf.AlterDb -CDbl(mp3Inf.MaxNoclipMp3Gain) * FIVELOG10TWO
                        DispJunk itmX, mp3Inf
                    End If
                    prgTot.Value = prgTot.Value + 1
                End If
                Set mp3Inf = Nothing
            End If
            If prgTot.Max > 100 Then
                UpdateCaptionPercentage Format$(CSng(prgTot.Value) * 100! / CSng(prgTot.Max), "0.0")
            Else
                UpdateCaptionPercentage CLng((prgTot.Value * 100) / prgTot.Max)
            End If
        End If
    Next
    stbStat.Panels(1).Text = ""
    Me.cmdCancel.Default = False
    Me.cmdCancel.Enabled = False
    Me.MousePointer = vbDefault
    EnableJunk (True)
    blnCurrentlyProcessing = False
    
    prgTot.Value = 0
    UpdateCaptionPercentage ""
    
    ShowBulkErrors
    
    If blnExitWhenDone Then Unload Me
    
    Exit Sub
    
MaxNoClipGain_Error:
    HandleError "MaxNoClipGain"
    On Error Resume Next
    blnCurrentlyProcessing = False
    EnableJunk (True)
End Sub

Private Sub ResetAnalysis()
    On Error GoTo ResetAnalysis_Error

    Dim itmX As ListItem
    Dim mp3Inf As Mp3Info
    Dim lngRetVal As Long
    
    
    blnResetWarnResult = True
    
    If blnResetWarn Then
        If mnuAlwaysTop.Checked Then
            SetWindowPos hWnd, conHwndNoTopmost, 0, 0, 0, 0, &H10 Or &H1 Or &H2
        End If
            
        frmResetWarn.Show vbModal, Me

        If mnuAlwaysTop.Checked Then
            SetWindowPos hWnd, conHwndTopmost, 0, 0, 0, 0, &H10 Or &H1 Or &H2
        End If
    End If
        
    If blnResetWarnResult Then
        For Each itmX In lstvMain.ListItems
            If (Not mnuSelectedFiles.Checked) Or (itmX.Selected) Then
                Set mp3Inf = flsMaster.Item(itmX.Key)
                mp3Inf.ResetVals
                DispJunk itmX, mp3Inf
                Set mp3Inf = Nothing
            End If
        Next
    End If
    
    Exit Sub
    
ResetAnalysis_Error:
    HandleError "ResetAnalysis"
End Sub

Private Sub ToolBarResize()
    On Error GoTo ToolBarResize_Error

    Dim tmpHeight, tmpDelta
    
    mnuToolbarText.Checked = False
    mnuToolbarSmall.Checked = False
    mnuToolBarBig.Checked = False
    mnuToolbarNone.Checked = False
    
    If Toolbar1.Visible Then
        tmpHeight = Toolbar1.Height
    Else
        tmpHeight = 0
    End If
    
    Select Case intToolBarSize
    Case 0:
        mnuToolbarNone.Checked = True
        Toolbar1.Visible = False
            
    Case 1:
        mnuToolbarText.Checked = True
        Toolbar1.Visible = True
        Toolbar1.ImageList = Nothing
        Toolbar1.TextAlignment = tbrTextAlignBottom
        Toolbar1.Buttons("addfiles").Caption = GetLocalString("frmMain.Button4.Caption", "Add File(s)")
        Toolbar1.Buttons("addfolder").Caption = GetLocalString("frmMain.Button5.Caption", "Add Folder")
        Toolbar1.Buttons("clearfiles").Caption = GetLocalString("frmMain.Button7.Caption", "Clear File(s)")
        Toolbar1.Buttons("clearall").Caption = GetLocalString("frmMain.Button8.Caption", "Clear All")
        Select Case intAnalysisMode
        Case 1
            Toolbar1.Buttons("analysis").Caption = GetLocalString("frmMain.Button1Menu1.Text", "Track Analysis")
        Case 2
            Toolbar1.Buttons("analysis").Caption = GetLocalString("frmMain.Button1Menu2.Text", "Album Analysis")
        End Select
        Select Case intGainMode
        Case 1
            Toolbar1.Buttons("gain").Caption = GetLocalString("frmMain.Button2Menu1.Text", "Track Gain")
        Case 2
            Toolbar1.Buttons("gain").Caption = GetLocalString("frmMain.Button2Menu2.Text", "Album Gain")
        Case 3
            Toolbar1.Buttons("gain").Caption = GetLocalString("frmMain.Button2Menu3.Text", "Constant Gain")
        End Select
            
    Case 2:
        mnuToolbarSmall.Checked = True
        Toolbar1.Visible = True
        Toolbar1.ImageList = Nothing
        Toolbar1.ImageList = smallHotImageList
        Toolbar1.TextAlignment = tbrTextAlignBottom
        Toolbar1.Buttons("addfiles").Caption = GetLocalString("frmMain.LCL_SMALL_FILES", "Files")
        Toolbar1.Buttons("addfiles").Image = 6
        Toolbar1.Buttons("addfolder").Caption = GetLocalString("frmMain.LCL_SMALL_FOLDER", "Folder")
        Toolbar1.Buttons("addfolder").Image = 7
        Toolbar1.Buttons("clearfiles").Caption = GetLocalString("frmMain.LCL_SMALL_FIlES", "Files")
        Toolbar1.Buttons("clearfiles").Image = 8
        Toolbar1.Buttons("clearall").Caption = GetLocalString("frmMain.LCL_SMALL_ALL", "All")
        Toolbar1.Buttons("clearall").Image = 9
        Select Case intAnalysisMode
        Case 1
            Toolbar1.Buttons("analysis").Caption = GetLocalString("frmMain.LCL_SMALL_RADIO", "Track")
            Toolbar1.Buttons("analysis").Image = 1
        Case 2
            Toolbar1.Buttons("analysis").Caption = GetLocalString("frmMain.LCL_SMALL_ALBUM", "Album")
            Toolbar1.Buttons("analysis").Image = 2
        End Select
        Select Case intGainMode
        Case 1
            Toolbar1.Buttons("gain").Caption = GetLocalString("frmMain.LCL_SMALL_RADIO", "Track")
            Toolbar1.Buttons("gain").Image = 3
        Case 2
            Toolbar1.Buttons("gain").Caption = GetLocalString("frmMain.LCL_SMALL_ALBUM", "Album")
            Toolbar1.Buttons("gain").Image = 4
        Case 3
            Toolbar1.Buttons("gain").Caption = GetLocalString("frmMain.LCL_SMALL_CONSTANT", "Const")
            Toolbar1.Buttons("gain").Image = 5
        End Select
        
    Case 3:
        mnuToolBarBig.Checked = True
        Toolbar1.Visible = True
        Toolbar1.ImageList = Nothing
        Toolbar1.ImageList = bigHotImageList
        Toolbar1.TextAlignment = tbrTextAlignBottom
        Toolbar1.Buttons("addfiles").Caption = GetLocalString("frmMain.Button4.Caption", "Add File(s)")
        Toolbar1.Buttons("addfiles").Image = 6
        Toolbar1.Buttons("addfolder").Caption = GetLocalString("frmMain.Button5.Caption", "Add Folder")
        Toolbar1.Buttons("addfolder").Image = 7
        Toolbar1.Buttons("clearfiles").Caption = GetLocalString("frmMain.Button7.Caption", "Clear File(s)")
        Toolbar1.Buttons("clearfiles").Image = 8
        Toolbar1.Buttons("clearall").Caption = GetLocalString("frmMain.Button8.Caption", "Clear All")
        Toolbar1.Buttons("clearall").Image = 9
        Select Case intAnalysisMode
        Case 1
            Toolbar1.Buttons("analysis").Caption = GetLocalString("frmMain.Button1Menu1.Text", "Track Analysis")
            Toolbar1.Buttons("analysis").Image = 1
        Case 2
            Toolbar1.Buttons("analysis").Caption = GetLocalString("frmMain.Button1Menu2.Text", "Album Analysis")
            Toolbar1.Buttons("analysis").Image = 2
        End Select
        Select Case intGainMode
        Case 1
            Toolbar1.Buttons("gain").Caption = GetLocalString("frmMain.Button2Menu1.Text", "Track Gain")
            Toolbar1.Buttons("gain").Image = 3
        Case 2
            Toolbar1.Buttons("gain").Caption = GetLocalString("frmMain.Button2Menu2.Text", "Album Gain")
            Toolbar1.Buttons("gain").Image = 4
        Case 3
            Toolbar1.Buttons("gain").Caption = GetLocalString("frmMain.Button2Menu3.Text", "Constant Gain")
            Toolbar1.Buttons("gain").Image = 5
        End Select
            
    End Select
    
    Toolbar1.Refresh
    If intToolBarSize <> 0 Then
        tmpDelta = tmpHeight - Toolbar1.Height
    Else
        tmpDelta = tmpHeight
    End If
    
    Me.lstvMain.Top = Me.lstvMain.Top - tmpDelta
    Me.lstvMain.Height = Me.lstvMain.Height + tmpDelta
    Me.Frame2.Top = Me.Frame2.Top - tmpDelta
    Me.lblNoUndo.Top = Me.lblNoUndo.Top - tmpDelta
    
    Exit Sub
    
ToolBarResize_Error:
    HandleError "ToolBarResize"
End Sub

Private Sub Form_Activate()
    Dim sngLeft As Single
    Dim sngTop As Single
    
    If blnFirstLoad Then
        blnFirstLoad = False
        
        ToolBarResize
        On Error Resume Next
        If sngFormHeight <> -1 Then
            Me.Height = sngFormHeight
        End If
        If sngFormWidth <> -1 Then
            Me.Width = sngFormWidth
        End If
        
        With nid
            .cbSize = Len(nid)
            .hWnd = frmSysTray.hWnd
            .uId = vbNull
            .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
            .uCallBackMessage = WM_MOUSEMOVE
            .hIcon = frmMain.Icon
            .szTip = frmMain.Caption & vbNullChar
        End With
            
        If (Me.Left + Me.Width) > Screen.Width Then
            sngLeft = Screen.Width - Me.Width
            If sngLeft < 0 Then sngLeft = 0
            Me.Left = sngLeft
        End If
        If (Me.Top + Me.Height) > Screen.Height Then
            sngTop = Screen.Height - Me.Height
            If sngTop < 0 Then sngTop = 0
            Me.Top = sngTop
        End If
        
        If Len(Command) > 0 Then
            StartupParseCommand
        End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo Form_KeyDown_Error
    Exit Sub
    
Form_KeyDown_Error:
    HandleError "Form_KeyDown"
End Sub

Private Sub Form_Load()
    On Error GoTo Form_Load_Error

    Dim ctrlBlah As Control
    Dim intDefOrder As Integer
    Dim sngOld As Single
    Dim sngDelta As Single
    Dim sngBase As Single
    Dim i As Integer
    Dim strKey As String
    Dim strVal As String
    Dim sngColWidth As Single
    Dim dblTarget As Double
    Dim diff2 As Single
    Dim diff3 As Single
    Dim diffProg As Single
    
    Dim ctrlItem As Control
    
    Dim lngRetVal As Long
    Dim intVer As Integer
    Dim intVerLen As Integer
    Dim sBlah As String
    Dim blnBackEndOK As Boolean
    
    blnAddingUndoSpace = False
    
    FiveMM = ScaleX(4, vbMillimeters, vbPixels)
    
    strAppPath = App.Path
    If Right$(strAppPath, 1) <> "\" Then strAppPath = strAppPath & "\"
    strAppDrive = GetDrivePartThing(strAppPath)
    
    App.HelpFile = strAppPath & "MP3Gain.chm"
    
    Me.Icon = LoadResPicture("APPICON", vbResIcon)
    
    SetDefaultControlFontName
    
    Set mcollOriginalCaptions = New Collection
    
    On Error Resume Next
    For Each ctrlItem In Me.Controls
        mcollOriginalCaptions.Add ctrlItem.Caption, ctrlItem.Name
    Next
    
    On Error GoTo Form_Load_Error
    
    sngProgLabelsLeft = lblFileProg.Left
    If (lblTotProg.Left < sngProgLabelsLeft) Then
        sngProgLabelsLeft = lblTotProg.Left
    End If
    
    diff2 = Label2.Width
    diff3 = Label3.Width
    
    mnuLanguage(0).Tag = "ORIGINAL"
    
    strCurLanguageFileName = GetSetting("MP3GainAnalysis", "StartUp", "LanguageFile", "")
    LoadLocalization strCurLanguageFileName
        
    strCurLanguageFileName = UCase$(mnuLanguage(intCurLanguage).Tag)
    
    blnBackEndOK = False
    lngRetVal = GetCommandOutput(sBlah, strAppPath & "mp3gain /v", strAppPath, False, True)
    intVer = InStr(LCase$(sBlah), "version")
    If intVer > 0 Then
        intVerLen = Len(Mid$(sBlah, intVer + 8)) - 2
        If intVerLen > 0 Then
            blnBackEndOK = True
        End If
    End If
    If Not blnBackEndOK Then
        MsgBox Replace(GetLocalString("frmMain.LCL_NO_BACK_END_1", "%%BACKENDFILE%% NOT FOUND. You will not be able to analyze or change your mp3 files."), "%%BACKENDFILE%%", strAppPath & "mp3gain.exe") & _
               vbCrLf & vbCrLf & _
               GetLocalString("frmMain.LCL_NO_BACK_END_2", "If you copied or moved MP3GainGUI.exe to this new folder, then either re-copy MP3GainGUI.exe into this new folder as a shortcut to the old folder, or move mp3gain.exe into this new folder.")
    End If
    
    fillCaptions Me
    mPopExit.Caption = mnuExit.Caption
    
    mnuPopupRadio.Caption = mnuRadio.Caption
    mnuPopupAlbum.Caption = mnuAlbum.Caption
    mnuPopupMaxAmp.Caption = mnuMaxAmp.Caption
    mnuPopupClearAnalysis.Caption = mnuClearAnalysis.Caption
    mnuPopupRadioGain.Caption = mnuRadioGain.Caption
    mnuPopupAlbumGain.Caption = mnuAlbumGain.Caption
    mnuPopupConstantGain.Caption = mnuConstantGain.Caption
    mnuPopupMaxNoclipGain.Caption = mnuMaxNoClipGain.Caption
    mnuPopupGroupNoclip.Caption = mnuGroupNoClip.Caption
    mnuPopupRemoveTags.Caption = mnuDeleteTags.Caption
    mnuPopupUndoGain.Caption = mnuUndoGain.Caption
    
    diffProg = lblFileProg.Left
    If (lblTotProg.Left < diffProg) Then diffProg = lblTotProg.Left
    
    diffProg = sngProgLabelsLeft - diffProg
    
    lblFileProg.Left = lblFileProg.Left + diffProg
    lblTotProg.Left = lblTotProg.Left + diffProg
    
    prgFile.Width = prgFile.Width - diffProg
    prgTot.Width = prgTot.Width - diffProg
    prgFile.Left = prgFile.Left + diffProg
    prgTot.Left = prgTot.Left + diffProg
    
    diff2 = Label2.Width - diff2
    
    Label2.Left = Label2.Left + diff2
    txtTargetInt.Left = txtTargetInt.Left + diff2
    lblDecimal.Left = lblDecimal.Left + diff2
    txtTargetDec.Left = txtTargetDec.Left + diff2
    Label3.Left = Label3.Left + diff2
    
    strClipYes = GetLocalString("frmMain.LCL_CLIP_YES", "Y")
    strClipMaybe = GetLocalString("frmMain.LCL_CLIP_MAYBE", "???")
    
    strOrgButtonTip(1) = Toolbar1.Buttons("analysis").ToolTipText
    strOrgButtonTip(2) = Toolbar1.Buttons("gain").ToolTipText
    strOrgButtonTip(4) = Toolbar1.Buttons("addfiles").ToolTipText
    strOrgButtonTip(5) = Toolbar1.Buttons("addfolder").ToolTipText
    strOrgButtonTip(7) = Toolbar1.Buttons("clearfiles").ToolTipText
    strOrgButtonTip(8) = Toolbar1.Buttons("clearall").ToolTipText
    strOrgButtonMenu(1) = Toolbar1.Buttons("analysis").ButtonMenus("radioanalysis").Text
    strOrgButtonMenu(2) = Toolbar1.Buttons("analysis").ButtonMenus("albumanalysis").Text
    strOrgButtonMenu(3) = Toolbar1.Buttons("analysis").ButtonMenus("clearanalysis").Text
    strOrgButtonMenu(4) = Toolbar1.Buttons("gain").ButtonMenus("radiogain").Text
    strOrgButtonMenu(5) = Toolbar1.Buttons("gain").ButtonMenus("albumgain").Text
    strOrgButtonMenu(6) = Toolbar1.Buttons("gain").ButtonMenus("constantgain").Text
    
    Toolbar1.Buttons("analysis").ToolTipText = GetLocalString("frmMain.Button1.ToolTipText", Toolbar1.Buttons("analysis").ToolTipText)
    Toolbar1.Buttons("analysis").ButtonMenus("radioanalysis").Text = GetLocalString("frmMain.Button1Menu1.Text", Toolbar1.Buttons("analysis").ButtonMenus("radioanalysis").Text)
    Toolbar1.Buttons("analysis").ButtonMenus("albumanalysis").Text = GetLocalString("frmMain.Button1Menu2.Text", Toolbar1.Buttons("analysis").ButtonMenus("albumanalysis").Text)
    Toolbar1.Buttons("analysis").ButtonMenus("clearanalysis").Text = GetLocalString("frmMain.Button1Menu4.Text", Toolbar1.Buttons("analysis").ButtonMenus("clearanalysis").Text)
    
    Toolbar1.Buttons("gain").ToolTipText = GetLocalString("frmMain.Button2.ToolTipText", Toolbar1.Buttons("gain").ToolTipText)
    Toolbar1.Buttons("gain").ButtonMenus("radiogain").Text = GetLocalString("frmMain.Button2Menu1.Text", Toolbar1.Buttons("gain").ButtonMenus("radiogain").Text)
    Toolbar1.Buttons("gain").ButtonMenus("albumgain").Text = GetLocalString("frmMain.Button2Menu2.Text", Toolbar1.Buttons("gain").ButtonMenus("albumgain").Text)
    Toolbar1.Buttons("gain").ButtonMenus("constantgain").Text = GetLocalString("frmMain.Button2Menu3.Text", Toolbar1.Buttons("gain").ButtonMenus("constantgain").Text)
    
    Toolbar1.Buttons("addfiles").ToolTipText = GetLocalString("frmMain.Button4.ToolTipText", Toolbar1.Buttons("addfiles").ToolTipText)
    Toolbar1.Buttons("addfolder").ToolTipText = GetLocalString("frmMain.Button5.ToolTipText", Toolbar1.Buttons("addfolder").ToolTipText)
    Toolbar1.Buttons("clearfiles").ToolTipText = GetLocalString("frmMain.Button7.ToolTipText", Toolbar1.Buttons("clearfiles").ToolTipText)
    Toolbar1.Buttons("clearall").ToolTipText = GetLocalString("frmMain.Button8.ToolTipText", Toolbar1.Buttons("clearall").ToolTipText)
    
    blnRecklessWarning = True
    
    blnMaxAmpOnly = False

    blnTargetIsChanging = False
    
    lsVolume = 1
    lsClip = 2
    lsRadioGain = 3
    lsRadioClip = 4
    lsMaxNoClip = 5
    lsAlbumVolume = 6
    lsAlbumGain = 7
    lsAlbumClip = 8
    glPath = 9
    glFile = 10
    lsMaxAmp = 11
    
    txtTargetInt.Text = DEFAULTTARGET
    lblDecimal.Caption = Mid$(Format$(1.2, "0.0"), 2, 1)
    txtTargetDec.Text = 0
    Me.Label3.Caption = "dB  " & Replace(GetLocalString("frmMain.LCL_TARGET_DB", "(default %%defaultTarget%%)"), "%%defaultTarget%%", Format$(DEFAULTTARGET, "0.0"))
    diff3 = Label3.Width - diff3
    Frame2.Width = Frame2.Width + diff2 + diff3
    Me.WindowState = GetSetting("MP3GainAnalysis", "StartUp", "WindowState", vbNormal)
    blnResetWarn = GetSetting("MP3GainAnalysis", "StartUp", "ResetWarn", True)
    blnUseTempFiles = GetSetting("MP3GainAnalysis", "StartUp", "UseTempFiles", True)
    mnuReckless.Checked = GetSetting("MP3GainAnalysis", "StartUp", "NoLayerCheck", False)
    blnRecklessWarning = GetSetting("MP3GainAnalysis", "StartUp", "NoLayerCheckWarning", True)
    strSaveLogsPath = GetSetting("MP3GainAnalysis", "StartUp", "SaveLogsPath", strAppPath)
    strSaveLogsFile = GetSetting("MP3GainAnalysis", "StartUp", "SaveLogsFile", "")
    strAddFolderPath = GetSetting("MP3GainAnalysis", "StartUp", "AddFolderPath")
    strAddFilesPath = GetSetting("MP3GainAnalysis", "StartUp", "AddFilesPath")
    lngThreadPriority = GetSetting("MP3GainAnalysis", "StartUp", "ThreadPriority", IDLE_PRIORITY_CLASS)
    dblTarget = Val(GetSetting("MP3GainAnalysis", "StartUp", "NormalTarget", DEFAULTTARGET))
    txtTargetInt.Text = Fix(dblTarget)
    txtTargetDec.Text = Round(Abs(dblTarget - Fix(dblTarget)) * 10#)
    If (Not IsNumeric(txtTargetInt.Text)) Or (Not IsNumeric(txtTargetDec.Text)) Then
        txtTargetInt.Text = DEFAULTTARGET
        txtTargetDec.Text = 0
    End If
    mnuAlwaysTop.Checked = GetSetting("MP3GainAnalysis", "StartUp", "AlwaysOnTop", 0)
    If mnuAlwaysTop.Checked Then SetWindowPos hWnd, conHwndTopmost, 0, 0, 0, 0, &H10 Or &H1 Or &H2

    mnuMaxAmp.Visible = GetSetting("MP3GainAnalysis", "StartUp", "MaximizingFeatures", False)
    mnuSep2.Visible = mnuMaxAmp.Visible
    mnuMaxNoClipGain.Visible = mnuMaxAmp.Visible
    mnuGroupNoClip.Visible = mnuMaxAmp.Visible
    mnuSep11.Visible = mnuMaxAmp.Visible
    
    mnuAddSubs.Checked = GetSetting("MP3GainAnalysis", "StartUp", "AddSubFolders", 1)
    mnuKeepTime.Checked = GetSetting("MP3GainAnalysis", "StartUp", "PreserveFileDate", False)
    mnuSelectedFiles.Checked = GetSetting("MP3GainAnalysis", "StartUp", "WorkOnSelectedFiles", False)
    mnuSysTray.Checked = GetSetting("MP3GainAnalysis", "StartUp", "TrayMinimize", False)
    mnuBeep.Checked = GetSetting("MP3GainAnalysis", "StartUp", "BeepWhenFinished", False)
    
    mnuSkipTags.Checked = GetSetting("MP3GainAnalysis", "StartUp", "IgnoreTags", False)
    WarnSkipTags mnuSkipTags.Checked
    
    mnuReCalcTags.Checked = GetSetting("MP3GainAnalysis", "StartUp", "ReCalculateTags", False)
    mnuSkipTagsWhileAdding.Checked = GetSetting("MP3GainAnalysis", "StartUp", "SkipTagsWhileAdding", False)
    
    mnuDontAddClipping.Checked = GetSetting("MP3GainAnalysis", "StartUp", "TrackNoClip", False)
    
    intAnalysisMode = GetSetting("MP3GainAnalysis", "StartUp", "AnalysisMode", 1)
    Select Case intAnalysisMode
    Case 1
        Toolbar1.Buttons("analysis").ButtonMenus("albumanalysis").Visible = True
        Toolbar1.Buttons("analysis").ButtonMenus("radioanalysis").Visible = False
    Case 2
        Toolbar1.Buttons("analysis").ButtonMenus("albumanalysis").Visible = False
        Toolbar1.Buttons("analysis").ButtonMenus("radioanalysis").Visible = True
    End Select
    
    intGainMode = GetSetting("MP3GainAnalysis", "StartUp", "GainMode", 1)
    Select Case intGainMode
    Case 1
        Toolbar1.Buttons("gain").ButtonMenus("radiogain").Visible = False
        Toolbar1.Buttons("gain").ButtonMenus("albumgain").Visible = True
        Toolbar1.Buttons("gain").ButtonMenus("constantgain").Visible = True
    Case 2
        Toolbar1.Buttons("gain").ButtonMenus("radiogain").Visible = True
        Toolbar1.Buttons("gain").ButtonMenus("albumgain").Visible = False
        Toolbar1.Buttons("gain").ButtonMenus("constantgain").Visible = True
    Case 3
        Toolbar1.Buttons("gain").ButtonMenus("radiogain").Visible = True
        Toolbar1.Buttons("gain").ButtonMenus("albumgain").Visible = True
        Toolbar1.Buttons("gain").ButtonMenus("constantgain").Visible = False
    End Select
    
    strErrLog = GetSetting("MP3GainAnalysis", "StartUp", "ErrLog", strAppPath & App.EXEName & "Errors.log")
    If strErrLog <> "" Then
        blnErrLog = True
    Else
        blnErrLog = False
    End If
    strChangeLog = GetSetting("MP3GainAnalysis", "StartUp", "ChangeLog", strAppPath & App.EXEName & "Changes.log")
    If strChangeLog <> "" Then
        blnChangeLog = True
    Else
        blnChangeLog = False
    End If
    strAnalysisLog = GetSetting("MP3GainAnalysis", "StartUp", "AnalysisLog", "")
    If strAnalysisLog <> "" Then
        blnAnalysisLog = True
    Else
        blnAnalysisLog = False
    End If
    
    intToolBarSize = GetSetting("MP3GainAnalysis", "StartUp", "ToolbarSize", 3)
    intAnalysisMode = GetSetting("MP3GainAnalysis", "StartUp", "AnalysisMode", 1)
    intGainMode = GetSetting("MP3GainAnalysis", "StartUp", "GainMode", 1)
    On Error Resume Next
    sngFormWidth = CSng(GetSetting("MP3GainAnalysis", "StartUp", "FormWidth", -1))
    If Err.Number <> 0 Then
        sngFormWidth = -1
    End If
    On Error Resume Next
    sngFormHeight = CSng(GetSetting("MP3GainAnalysis", "StartUp", "FormHeight", -1))
    If Err.Number <> 0 Then
        sngFormHeight = -1
    End If
    On Error GoTo Form_Load_Error
    
    blnStereoWarning = GetSetting("MP3GainAnalysis", "StartUp", "StereoWarning", True)
    blnSkipTagsWarn = GetSetting("MP3GainAnalysis", "StartUp", "SkipTagsWarning", True)
    
    blnShowFileStatus = GetSetting("MP3GainAnalysis", "StartUp", "ShowFileProgress", True)
    If blnShowFileStatus = False Then
        Me.prgFile.Visible = False
        Me.lblFileProg.Visible = False
    End If
    
    
    intPathFileStat = PATHFILE
    
    Set flsMaster = New Collection
    
    blnExitWhenDone = False
    
    blnTxtProgChanging = False
    
    Load frmGetGain
    Load frmOptions
    Load frmReadOnly
    Load frmLogs
    
    
    blnDoingAlbum = False
    lstvMain.ColumnHeaders.Add 1, "chPathFile", _
        GetLocalString("frmMain.LCL_COLUMN_PATH_FILE", "Path\File")
        lstvMain.ColumnHeaders.Add 2, "chVolume", _
        GetLocalString("frmMain.LCL_COLUMN_VOLUME", "Volume")
        lstvMain.ColumnHeaders.Add 3, "chClip", _
        GetLocalString("frmMain.LCL_COLUMN_CLIPPING", "clipping")
        lstvMain.ColumnHeaders.Add 4, "chRadioGain", _
        GetLocalString("frmMain.LCL_COLUMN_RADIO_GAIN", "Track Gain")
        lstvMain.ColumnHeaders.Add 5, "chRadioClip", _
        GetLocalString("frmMain.LCL_COLUMN_RADIO_CLIP", "clip(Track)")
        lstvMain.ColumnHeaders.Add 6, "chMaxNoClip", _
        GetLocalString("frmMain.LCL_COLUMN_MAXIMUM_NOCLIP", "Max Noclip Gain")
        lstvMain.ColumnHeaders.Add 7, "chAlbumVolume", _
        GetLocalString("frmMain.LCL_COLUMN_ALBUM_VOLUME", "Album Volume")
        lstvMain.ColumnHeaders.Add 8, "chAlbumGain", _
        GetLocalString("frmMain.LCL_COLUMN_ALBUM_GAIN", "Album Gain")
        lstvMain.ColumnHeaders.Add 9, "chAlbumClip", _
        GetLocalString("frmMain.LCL_COLUMN_ALBUM_CLIP", "clip(Album)")
        lstvMain.ColumnHeaders.Add 10, "chPath", _
        GetLocalString("frmMain.LCL_COLUMN_PATH", "Path")
        lstvMain.ColumnHeaders.Add 11, "chFile", _
        GetLocalString("frmMain.LCL_COLUMN_FILE", "File")
        lstvMain.ColumnHeaders.Add 12, "chMaxAmp", _
        GetLocalString("frmMain.LCL_COLUMN_MAXIMUM_AMPLITUDE", "Curr Max Amp")
    
    intDefOrder = GetSetting("MP3GainAnalysis", "StartUp", "PathFileDisplay", PATHFILE)
    Select Case intDefOrder
    Case PATHSEPFILE
        If PathFileDisplay(PATHSEPFILE) Then
            mnuPathWithFile.Checked = False
            mnuFileOnly.Checked = False
            mnuPathSepFile.Checked = True
        End If
    Case FILEONLY
        If PathFileDisplay(FILEONLY) Then
            mnuPathWithFile.Checked = False
            mnuPathSepFile.Checked = False
            mnuFileOnly.Checked = True
        End If
    End Select
    
    ResetColumnWidths 'For default
    
    For i = 1 To lstvMain.ColumnHeaders.Count
        strKey = "Column" & Format$(i, "00")
        On Error Resume Next
        sngColWidth = CSng(GetSetting("MP3GainAnalysis", "StartUp", strKey, ""))
        If Err.Number = 0 Then
            lstvMain.ColumnHeaders(i).Width = sngColWidth
        End If
        On Error GoTo Form_Load_Error
    Next i

    sngPrevWidth = Me.Width
    sngPrevHeight = Me.Height

    prgFile.Value = 0
    
    Me.ScaleMode = vbPixels
    stbStat.Panels(3).Width = TextWidth("99999") + FiveMM
    stbStat.Panels(2).Width = TextWidth("00:00:00") + FiveMM
    stbStat.Panels(2).Visible = False
    Me.ScaleMode = vbTwips
    'stbStat.Panels(1).Width = stbStat.Width - (300 + stbStat.Panels(2).Width + stbStat.Panels(3).Width)
    stbStat.Panels(1).Width = stbStat.Width - (300 + stbStat.Panels(3).Width)
    
    blnFirstLoad = True
    
    sngShowAddedFilesTimer = Timer
    
    Exit Sub
    
Form_Load_Error:
    HandleError "Form_Load"
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo Form_QueryUnload_Error

    If UnloadMode = vbFormControlMenu Then
        If blnCurrentlyProcessing Then
            blnCancel = True
            blnExitWhenDone = True
            Cancel = True
            stbStat.Panels(1).Text = GetLocalString("frmMain.LCL_EXITING", "Exiting...")
        End If
    End If
    
    Exit Sub
    
Form_QueryUnload_Error:
    HandleError "Form_QueryUnload"
End Sub

Private Sub Form_Resize()
    On Error GoTo Form_Resize_Error

    Dim sngDHeight As Single
    Dim sngDWidth As Single
    
    If Me.WindowState <> vbMinimized Then
        If mnuSysTray.Checked Then
            Shell_NotifyIcon NIM_DELETE, nid
        End If
        Me.Caption = App.Title
        modTrayToolTip Me.Caption
        If Me.Height < MINFORMHEIGHT Then
            Me.Height = MINFORMHEIGHT
        End If
        If Me.Width < MINFORMWIDTH Then
            Me.Width = MINFORMWIDTH
        End If
        
        sngDHeight = sngPrevHeight - Me.Height
        sngDWidth = sngPrevWidth - Me.Width
        sngPrevWidth = Me.Width
        sngPrevHeight = Me.Height
        
        lstvMain.Height = lstvMain.Height - sngDHeight
        
        lstvMain.Width = lstvMain.Width - sngDWidth
        
        cmdCancel.Top = cmdCancel.Top - sngDHeight
        cmdExit.Top = cmdExit.Top - sngDHeight
        lblFileProg.Top = lblFileProg.Top - sngDHeight
        lblTotProg.Top = lblTotProg.Top - sngDHeight
        prgFile.Top = prgFile.Top - sngDHeight
        prgTot.Top = prgTot.Top - sngDHeight
        prgFile.Width = prgFile.Width - sngDWidth
        prgTot.Width = prgTot.Width - sngDWidth
        lblNoUndo.Left = lblNoUndo.Left - sngDWidth
        
        cmdCancel.Left = cmdCancel.Left - sngDWidth / 2
        cmdExit.Left = cmdExit.Left - sngDWidth
        stbStat.Panels(1).Width = stbStat.Panels(1).Width - sngDWidth
    Else
        If mnuSysTray.Checked Then
            Me.Hide
            Shell_NotifyIcon NIM_ADD, nid
        End If
        If (prgTot.Value > 0) Or (blnCurrentlyProcessing) Then
            If prgTot.Max > 100 Then
                Me.Caption = App.Title & " [" & Format$(CSng(prgTot.Value) * 100! / CSng(prgTot.Max), "0.0") & "%]"
                modTrayToolTip Me.Caption
            Else
                Me.Caption = App.Title & " [" & CLng((prgTot.Value * 100) / prgTot.Max) & "%]"
                modTrayToolTip Me.Caption
            End If
        End If
    End If
    
    Exit Sub
    
Form_Resize_Error:
    HandleError "Form_Resize"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo Form_Unload_Error

    Dim i As Integer
    Dim strKey As String
    
    Unload frmLogs
    Unload frmReadOnly
    Unload frmOptions
    Unload frmGetGain
    
    Shell_NotifyIcon NIM_DELETE, nid
    Unload frmSysTray
    
    SaveSetting "MP3GainAnalysis", "StartUp", "LanguageFile", strCurLanguageFileName

    If Me.WindowState <> vbMinimized Then
        SaveSetting "MP3GainAnalysis", "StartUp", "FormHeight", Me.Height
        SaveSetting "MP3GainAnalysis", "StartUp", "FormWidth", Me.Width
        SaveSetting "MP3GainAnalysis", "StartUp", "WindowState", Me.WindowState
    End If
    SaveSetting "MP3GainAnalysis", "StartUp", "SaveLogsPath", strSaveLogsPath
    SaveSetting "MP3GainAnalysis", "StartUp", "SaveLogsFile", strSaveLogsFile
    SaveSetting "MP3GainAnalysis", "StartUp", "AddFolderPath", strAddFolderPath
    SaveSetting "MP3GainAnalysis", "StartUp", "AddFilesPath", strAddFilesPath
    SaveSetting "MP3GainAnalysis", "StartUp", "ThreadPriority", lngThreadPriority
    SaveSetting "MP3GainAnalysis", "StartUp", "NormalTarget", txtTargetInt.Text & "." & txtTargetDec.Text
    SaveSetting "MP3GainAnalysis", "StartUp", "AlwaysOnTop", mnuAlwaysTop.Checked
    SaveSetting "MP3GainAnalysis", "StartUp", "EachFolderIsAlbum", mnuEachAlbum.Checked
    SaveSetting "MP3GainAnalysis", "StartUp", "AddSubFolders", mnuAddSubs.Checked
    SaveSetting "MP3GainAnalysis", "StartUp", "PreserveFileDate", mnuKeepTime.Checked
    SaveSetting "MP3GainAnalysis", "StartUp", "ToolbarSize", intToolBarSize
    SaveSetting "MP3GainAnalysis", "StartUp", "AnalysisMode", intAnalysisMode
    SaveSetting "MP3GainAnalysis", "StartUp", "GainMode", intGainMode
    SaveSetting "MP3GainAnalysis", "StartUp", "WorkOnSelectedFiles", mnuSelectedFiles.Checked
    SaveSetting "MP3GainAnalysis", "StartUp", "NoLayerCheck", mnuReckless.Checked
    SaveSetting "MP3GainAnalysis", "StartUp", "NoLayerCheckWarning", blnRecklessWarning
    SaveSetting "MP3GainAnalysis", "StartUp", "TrayMinimize", mnuSysTray.Checked
    SaveSetting "MP3GainAnalysis", "StartUp", "BeepWhenFinished", mnuBeep.Checked

    SaveSetting "MP3GainAnalysis", "StartUp", "MaximizingFeatures", mnuMaxAmp.Visible

    SaveSetting "MP3GainAnalysis", "StartUp", "ErrLog", strErrLog
    SaveSetting "MP3GainAnalysis", "StartUp", "ChangeLog", strChangeLog
    SaveSetting "MP3GainAnalysis", "StartUp", "AnalysisLog", strAnalysisLog
    SaveSetting "MP3GainAnalysis", "StartUp", "PathFileDisplay", intPathFileStat
    
    SaveSetting "MP3GainAnalysis", "StartUp", "StereoWarning", blnStereoWarning
    SaveSetting "MP3GainAnalysis", "StartUp", "UseTempFiles", blnUseTempFiles
    SaveSetting "MP3GainAnalysis", "StartUp", "ShowFileProgress", blnShowFileStatus
    SaveSetting "MP3GainAnalysis", "StartUp", "ResetWarn", blnResetWarn
    
    SaveSetting "MP3GainAnalysis", "StartUp", "IgnoreTags", mnuSkipTags.Checked
    SaveSetting "MP3GainAnalysis", "StartUp", "ReCalculateTags", mnuReCalcTags.Checked
    SaveSetting "MP3GainAnalysis", "StartUp", "SkipTagsWhileAdding", mnuSkipTagsWhileAdding.Checked
    
    SaveSetting "MP3GainAnalysis", "StartUp", "TrackNoClip", mnuDontAddClipping.Checked
    
    SaveSetting "MP3GainAnalysis", "StartUp", "SkipTagsWarning", blnSkipTagsWarn
    
    For i = 1 To lstvMain.ColumnHeaders.Count
        strKey = "Column" & Format$(i, "00")
        SaveSetting "MP3GainAnalysis", "StartUp", strKey, lstvMain.ColumnHeaders(i).Width
    Next i

          
    Exit Sub
    
Form_Unload_Error:
    HandleError "Form_Unload"
End Sub

Private Sub DeModSortLongCol(sCol As Integer)
    Dim i As Integer
    Dim sbiCur As ListSubItem
    
    For i = 1 To lstvMain.ListItems.Count
        Set sbiCur = lstvMain.ListItems(i).ListSubItems(sCol)
        If IsNumeric(sbiCur.Text) Then
            sbiCur.Text = Round(CDbl(sbiCur.Text))
        End If
    Next i
    Set sbiCur = Nothing
End Sub

Private Sub DeModSortDblCol(sCol As Integer, dblModVal As Double)
    Dim i As Integer
    Dim sbiCur As ListSubItem
    
    For i = 1 To lstvMain.ListItems.Count
        Set sbiCur = lstvMain.ListItems(i).ListSubItems(sCol)
        If IsNumeric(sbiCur.Text) Then
            sbiCur.Text = Format$((CDbl(sbiCur.Text) / 10#) - dblModVal, "0.0")
        End If
    Next i
    Set sbiCur = Nothing
End Sub

Private Sub ModSortLongCol(sCol As Integer)
    Dim i As Integer
    Dim sbiCur As ListSubItem
    
    For i = 1 To lstvMain.ListItems.Count
        Set sbiCur = lstvMain.ListItems(i).ListSubItems(sCol)
        If IsNumeric(sbiCur.Text) Then sbiCur.Text = Format$(sbiCur.Text, "000000000000000")
    Next i
    Set sbiCur = Nothing
End Sub

Private Function ModSortDblCol(sCol As Integer) As Double
    Dim i As Integer
    Dim dblModVal As Double
    Dim sbiCur As ListSubItem
    
    dblModVal = 100000#
    For i = 1 To lstvMain.ListItems.Count
        Set sbiCur = lstvMain.ListItems(i).ListSubItems(sCol)
        If IsNumeric(sbiCur.Text) Then
            If CDbl(sbiCur.Text) < dblModVal Then
                dblModVal = CDbl(sbiCur.Text)
            End If
        End If
    Next i
    
    If dblModVal < 1# Then
        dblModVal = 1# - dblModVal
    Else
        dblModVal = 0#
    End If
    
    For i = 1 To lstvMain.ListItems.Count
        Set sbiCur = lstvMain.ListItems(i).ListSubItems(sCol)
        If IsNumeric(sbiCur.Text) Then
            sbiCur.Text = Format$(CLng((CDbl(sbiCur.Text) + dblModVal) * 10#), "00000")
        End If
    Next i
    
    Set sbiCur = Nothing
    ModSortDblCol = dblModVal
End Function

Public Sub doSortColumn()
    Dim dblModVal As Double
    
    lstvMain.SortKey = sortColumn

    Select Case sortColumn
    Case lsMaxAmp:
        'Use sort routine to sort by long
        
        Call ModSortLongCol(sortColumn)
        lstvMain.SortOrder = Abs(sOrder)
        lstvMain.Sorted = True
        lstvMain.Sorted = False
        
        Call DeModSortLongCol(sortColumn)
        
    Case lsVolume, lsRadioGain, lsMaxNoClip, lsAlbumVolume, lsAlbumGain:
        'Use sort routine to sort by double
        
        dblModVal = ModSortDblCol(sortColumn)
        lstvMain.SortOrder = Abs(sOrder)
        lstvMain.Sorted = True
        lstvMain.Sorted = False
           
        Call DeModSortDblCol(sortColumn, dblModVal)
        
    Case Else:
        'Use default sorting to sort the items in the list
        lstvMain.SortOrder = Abs(sOrder)
        lstvMain.Sorted = True
    End Select
End Sub

Private Sub lblNoUndo_Change()
    If Not blnAddingUndoSpace Then
        blnAddingUndoSpace = True
        If Left$(lblNoUndo.Caption, 1) <> " " Then
            lblNoUndo.Caption = " " & lblNoUndo.Caption
        End If
        If Right$(lblNoUndo.Caption, 1) <> " " Then
            lblNoUndo.Caption = lblNoUndo.Caption & " "
        End If
        blnAddingUndoSpace = False
    End If
End Sub

Private Sub lstvMain_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    On Error GoTo lstvMain_ColumnClick_Error
    
    If blnListChangeable Then
        
        sOrder = Not sOrder
        
        sortColumn = ColumnHeader.Index - 1
        
        doSortColumn
        

    End If
    lstvMain.Refresh
    Exit Sub
    
lstvMain_ColumnClick_Error:
    HandleError "lstvMain_ColumnClick"
End Sub

Private Sub lstvMain_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo lstvMain_KeyDown_Error
    Dim itmX As ListItem
    
    If blnListChangeable Then
        Select Case KeyCode
        Case 46
            ClearFiles
        Case 65
            If Shift = 2 Then
                For Each itmX In lstvMain.ListItems
                    itmX.Selected = True
                Next itmX
            End If
        End Select
    End If
    
    Exit Sub
    
lstvMain_KeyDown_Error:
    HandleError "lstvMain_KeyDown"
End Sub

Private Sub lstvMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo lstvMain_MouseUp_Error
    Dim tmpSelected As Boolean
    
    If Button = 2 Then
        mnuPopupRadio.Enabled = mnuRadio.Enabled
        mnuPopupAlbum.Enabled = mnuAlbum.Enabled
        mnuPopupMaxAmp.Enabled = mnuMaxAmp.Enabled
        mnuPopupClearAnalysis.Enabled = mnuClearAnalysis.Enabled
        mnuPopupRadioGain.Enabled = mnuRadioGain.Enabled
        mnuPopupAlbumGain.Enabled = mnuAlbumGain.Enabled
        mnuPopupConstantGain.Enabled = mnuConstantGain.Enabled
        mnuPopupMaxNoclipGain.Enabled = mnuMaxNoClipGain.Enabled
        mnuPopupGroupNoclip.Enabled = mnuGroupNoClip.Enabled
        
        mnuPopupMaxAmp.Visible = mnuMaxAmp.Visible
        mnuPopupMaxNoclipGain.Visible = mnuMaxNoClipGain.Visible
        mnuPopupGroupNoclip.Visible = mnuGroupNoClip.Visible
        
        mnuPopupUndoGain.Enabled = mnuUndoGain.Enabled
        
        mnuPopupRemoveTags.Enabled = mnuDeleteTags.Enabled
        
        tmpSelected = mnuSelectedFiles.Checked
        
        mnuSelectedFiles.Checked = True
        Me.PopupMenu mnuPopup
        mnuSelectedFiles.Checked = tmpSelected
    End If
    
    Exit Sub

lstvMain_MouseUp_Error:
    HandleError ("lstvMain_MouseUp")
End Sub

Private Sub lstvMain_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, _
        Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo lstvMain_OLEDragDrop_Error
    Dim i As Long
    Dim arrFiles() As String
    
    blnCancel = False
    blnCurrentlyProcessing = True
    EnableJunk (False)
    Me.MousePointer = vbArrowHourglass
    Me.cmdCancel.Enabled = True
    Me.cmdCancel.Default = True
        
    If Data.GetFormat(vbCFFiles) Then
        ReDim arrFiles(0 To Data.Files.Count - 1)
        For i = 1 To Data.Files.Count
            arrFiles(i - 1) = Data.Files(i)
        Next i
        AddListOfFiles arrFiles, strAppDrive, strAppPath
    End If
    
    stbStat.Panels(1).Text = ""
    Me.cmdCancel.Default = False
    Me.cmdCancel.Enabled = False
    Me.MousePointer = vbDefault
    EnableJunk (True)
    blnCurrentlyProcessing = False
    If blnExitWhenDone Then Unload Me
    
    Exit Sub
    
lstvMain_OLEDragDrop_Error:
    HandleError "lstvMain_OLEDragDrop"
End Sub

Private Sub mnuAbout_Click()
    On Error GoTo mnuAbout_Click_Error

    If mnuAlwaysTop.Checked Then
        SetWindowPos hWnd, conHwndNoTopmost, 0, 0, 0, 0, &H10 Or &H1 Or &H2
    End If
    
    frmAbout.Show vbModal, Me
    
    If mnuAlwaysTop.Checked Then
        SetWindowPos hWnd, conHwndTopmost, 0, 0, 0, 0, &H10 Or &H1 Or &H2
    End If
    
    Exit Sub
    
mnuAbout_Click_Error:
    HandleError "mnuAbout_Click"
End Sub

Private Sub mnuAddFile_Click()
    On Error GoTo mnuAddFile_Click_Error

    AddFiles
    
    Exit Sub
    
mnuAddFile_Click_Error:
    HandleError "mnuAddFile_Click"
End Sub

Private Sub mnuAddFolder_Click()
    On Error GoTo mnuAddFolder_Click_Error

    AddFolder
    
    Exit Sub
    
mnuAddFolder_Click_Error:
    HandleError "mnuAddFolder_Click"
End Sub

Private Sub mnuAddSubs_Click()
    On Error GoTo mnuAddSubs_Click_Error

    mnuAddSubs.Checked = Not mnuAddSubs.Checked
    
    Exit Sub
    
mnuAddSubs_Click_Error:
    HandleError "mnuAddSubs_Click"
End Sub

Private Sub mnuAdvancedOptions_Click()
    On Error GoTo mnuAdvancedOptions_Click_Error


    If mnuAlwaysTop.Checked Then
        SetWindowPos hWnd, conHwndNoTopmost, 0, 0, 0, 0, &H10 Or &H1 Or &H2
    End If
    
    frmOptions.Show vbModal, Me
    
    If mnuAlwaysTop.Checked Then
        SetWindowPos hWnd, conHwndTopmost, 0, 0, 0, 0, &H10 Or &H1 Or &H2
    End If
    
    Exit Sub
    
mnuAdvancedOptions_Click_Error:
    HandleError "mnuAdvancedOptions_Click"
End Sub

Private Sub mnuAlbum_Click()
    On Error GoTo mnuAlbum_Click_Error

    Album
    
    Exit Sub
    
mnuAlbum_Click_Error:
    HandleError "mnuAlbum_Click"
End Sub

Private Sub mnuAlbumGain_Click()
    On Error GoTo mnuAlbumGain_Click_Error

    AlbumGain
    
    Exit Sub
    
mnuAlbumGain_Click_Error:
    HandleError "mnuAlbumGain_Click"
End Sub

Private Sub mnuAlwaysTop_Click()
    On Error GoTo mnuAlwaysTop_Click_Error

    mnuAlwaysTop.Checked = Not mnuAlwaysTop.Checked
    If mnuAlwaysTop.Checked Then
        SetWindowPos hWnd, conHwndTopmost, 0, 0, 0, 0, &H10 Or &H1 Or &H2
    Else
        SetWindowPos hWnd, conHwndNoTopmost, 0, 0, 0, 0, &H10 Or &H1 Or &H2
    End If

    Exit Sub
mnuAlwaysTop_Click_Error:
    HandleError "mnuAlwaysTop_Click"
End Sub

Private Sub mnuBeep_Click()
    mnuBeep.Checked = Not mnuBeep.Checked
End Sub

Private Sub mnuClearAll_Click()
    ClearAll
End Sub

Private Sub mnuClearAnalysis_Click()
    On Error GoTo mnuClearAnalysis_Click_Error

    ResetAnalysis
    
    Exit Sub
    
mnuClearAnalysis_Click_Error:
    HandleError "mnuClearAnalysis_Click"
End Sub

Private Sub mnuClearFiles_Click()
    ClearFiles
End Sub

Private Sub mnuConstantGain_Click()
    On Error GoTo mnuConstantGain_Click_Error

    ConstGain
    
    Exit Sub
    
mnuConstantGain_Click_Error:
    HandleError "mnuConstantGain_Click"
End Sub

Private Sub mnuContents_Click()
    On Error Resume Next
    Dim hwndHelp As Long
    Dim strCheckLocalized As String
    
    'App.HelpFile is set in LoadLocalization or ReloadLocalization
    If Len(Dir(App.HelpFile)) > 0 Then
        hwndHelp = HtmlHelp(GetDesktopWindow(), App.HelpFile, HH_DISPLAY_TOPIC, 0)
    Else
        MsgBox Replace(GetLocalString("frmMain.LCL_NO_HELP_FOUND", "%%HELPFILE%% NOT FOUND. If you copied or moved MP3GainGUI.exe to this folder after installation, then please move the .chm file to this folder also."), "%%HELPFILE%%", App.HelpFile)
    End If
End Sub

Private Sub mnuDeleteTags_Click()
    Call DoFileAction(taDeleteTags)
End Sub

Private Sub mnuDisclaimer_Click()
    If mnuAlwaysTop.Checked Then
        SetWindowPos hWnd, conHwndNoTopmost, 0, 0, 0, 0, &H10 Or &H1 Or &H2
    End If
    
    frmDisclaimer.Show vbModal, Me
    
    If mnuAlwaysTop.Checked Then
        SetWindowPos hWnd, conHwndTopmost, 0, 0, 0, 0, &H10 Or &H1 Or &H2
    End If
End Sub

Private Sub mnuDonate_Click()
    If mnuAlwaysTop.Checked Then
        SetWindowPos hWnd, conHwndNoTopmost, 0, 0, 0, 0, &H10 Or &H1 Or &H2
    End If
    
    frmDonate.Show vbModal, Me
    
    If mnuAlwaysTop.Checked Then
        SetWindowPos hWnd, conHwndTopmost, 0, 0, 0, 0, &H10 Or &H1 Or &H2
    End If
End Sub

Private Sub mnuDontAddClipping_Click()
    mnuDontAddClipping.Checked = Not mnuDontAddClipping.Checked
End Sub

Private Sub mnuEachAlbum_Click()
    On Error GoTo mnuEachAlbum_Click_Error

    mnuEachAlbum.Checked = Not mnuEachAlbum.Checked
    
    Exit Sub
    
mnuEachAlbum_Click_Error:
    HandleError "mnuEachAlbum_Click"
End Sub

Private Sub mnuExit_Click()
    On Error GoTo mnuExit_Click_Error

    If blnCurrentlyProcessing Then
        blnCancel = True
        blnExitWhenDone = True
        stbStat.Panels(1).Text = GetLocalString("frmMain.LCL_EXITING", "Exiting...")
    Else
        Unload Me
    End If
    
    Exit Sub
    
mnuExit_Click_Error:
    HandleError "mnuExit_Click"
End Sub

Private Sub mnuFileOnly_Click()
    On Error GoTo mnuFileOnly_Click_Error
    If intPathFileStat <> FILEONLY Then
        If PathFileDisplay(FILEONLY) Then
            mnuPathWithFile.Checked = False
            mnuPathSepFile.Checked = False
            mnuFileOnly.Checked = True
        End If
    End If
    
    Exit Sub
    
mnuFileOnly_Click_Error:
    HandleError "mnuFileOnly_Click"
End Sub

Private Sub mnuGroupNoClip_Click()
    Call GroupMaxNoClipGain
End Sub

Private Sub mnuKeepTime_Click()
    mnuKeepTime.Checked = Not mnuKeepTime.Checked
End Sub

Private Sub mnuLanguage_Click(Index As Integer)
    Dim ctrlItem As Control
    Dim diff2 As Single
    Dim diff3 As Single
    Dim diffProg As Single
    
    If Index = intCurLanguage Then
        Exit Sub
    End If
    
    If Not ReloadLocalization(Index, strAppPath & mnuLanguage(Index).Tag) Then Index = 0
        
    mnuLanguage(intCurLanguage).Checked = False
    mnuLanguage(Index).Checked = True
    intCurLanguage = Index
    
    strCurLanguageFileName = UCase$(mnuLanguage(Index).Tag)
    
    Unload frmLogs
    Unload frmReadOnly
    Unload frmOptions
    Unload frmGetGain
    
    Load frmGetGain
    Load frmOptions
    Load frmReadOnly
    Load frmLogs
    
    diff2 = Label2.Width
    diff3 = Label3.Width

    'First, reset to English in case there are captions which were translated
    'in previous language, but not in this one
    On Error Resume Next
    For Each ctrlItem In Me.Controls
        If ctrlItem.Name <> "mnuLanguage" Then
            ctrlItem.Caption = mcollOriginalCaptions(ctrlItem.Name)
        End If
    Next
    On Error GoTo mnuLanguage_Click_Error
    
    fillCaptions Me
    mPopExit.Caption = mnuExit.Caption
    
    mnuPopupRadio.Caption = mnuRadio.Caption
    mnuPopupAlbum.Caption = mnuAlbum.Caption
    mnuPopupMaxAmp.Caption = mnuMaxAmp.Caption
    mnuPopupClearAnalysis.Caption = mnuClearAnalysis.Caption
    mnuPopupRadioGain.Caption = mnuRadioGain.Caption
    mnuPopupAlbumGain.Caption = mnuAlbumGain.Caption
    mnuPopupConstantGain.Caption = mnuConstantGain.Caption
    mnuPopupMaxNoclipGain.Caption = mnuMaxNoClipGain.Caption
    mnuPopupGroupNoclip.Caption = mnuGroupNoClip.Caption
    
    diffProg = lblFileProg.Left
    If (lblTotProg.Left < diffProg) Then diffProg = lblTotProg.Left
    
    diffProg = sngProgLabelsLeft - diffProg

    lblFileProg.Left = lblFileProg.Left + diffProg
    lblTotProg.Left = lblTotProg.Left + diffProg
    
    prgFile.Width = prgFile.Width - diffProg
    prgTot.Width = prgTot.Width - diffProg
    prgFile.Left = prgFile.Left + diffProg
    prgTot.Left = prgTot.Left + diffProg
    
    diff2 = Label2.Width - diff2
    
    Label2.Left = Label2.Left + diff2
    txtTargetInt.Left = txtTargetInt.Left + diff2
    lblDecimal.Left = lblDecimal.Left + diff2
    txtTargetDec.Left = txtTargetDec.Left + diff2
    Label3.Left = Label3.Left + diff2
    
    strClipYes = GetLocalString("frmMain.LCL_CLIP_YES", "Y")
    strClipMaybe = GetLocalString("frmMain.LCL_CLIP_MAYBE", "???")
    
    Toolbar1.Buttons("analysis").ToolTipText = GetLocalString("frmMain.Button1.ToolTipText", strOrgButtonTip(1))
    Toolbar1.Buttons("analysis").ButtonMenus("radioanalysis").Text = GetLocalString("frmMain.Button1Menu1.Text", strOrgButtonMenu(1))
    Toolbar1.Buttons("analysis").ButtonMenus("albumanalysis").Text = GetLocalString("frmMain.Button1Menu2.Text", strOrgButtonMenu(2))
    Toolbar1.Buttons("analysis").ButtonMenus("clearanalysis").Text = GetLocalString("frmMain.Button1Menu4.Text", strOrgButtonMenu(3))
    
    Toolbar1.Buttons("gain").ToolTipText = GetLocalString("frmMain.Button2.ToolTipText", strOrgButtonTip(2))
    Toolbar1.Buttons("gain").ButtonMenus("radiogain").Text = GetLocalString("frmMain.Button2Menu1.Text", strOrgButtonMenu(4))
    Toolbar1.Buttons("gain").ButtonMenus("albumgain").Text = GetLocalString("frmMain.Button2Menu2.Text", strOrgButtonMenu(5))
    Toolbar1.Buttons("gain").ButtonMenus("constantgain").Text = GetLocalString("frmMain.Button2Menu3.Text", strOrgButtonMenu(6))
    
    Toolbar1.Buttons("addfiles").ToolTipText = GetLocalString("frmMain.Button4.ToolTipText", strOrgButtonTip(4))
    Toolbar1.Buttons("addfolder").ToolTipText = GetLocalString("frmMain.Button5.ToolTipText", strOrgButtonTip(5))
    Toolbar1.Buttons("clearfiles").ToolTipText = GetLocalString("frmMain.Button7.ToolTipText", strOrgButtonTip(7))
    Toolbar1.Buttons("clearall").ToolTipText = GetLocalString("frmMain.Button8.ToolTipText", strOrgButtonTip(8))
    
    ToolBarResize
    
    Me.Label3.Caption = "dB  " & Replace(GetLocalString("frmMain.LCL_TARGET_DB", "(default %%defaultTarget%%)"), "%%defaultTarget%%", Format$(DEFAULTTARGET, "0.0"))
    diff3 = Label3.Width - diff3
    Frame2.Width = Frame2.Width + diff2 + diff3
    
    lstvMain.ColumnHeaders(1).Text = GetLocalString("frmMain.LCL_COLUMN_PATH_FILE", "Path\File")
    lstvMain.ColumnHeaders(2).Text = GetLocalString("frmMain.LCL_COLUMN_VOLUME", "Volume")
    lstvMain.ColumnHeaders(3).Text = GetLocalString("frmMain.LCL_COLUMN_CLIPPING", "clipping")
    lstvMain.ColumnHeaders(4).Text = GetLocalString("frmMain.LCL_COLUMN_RADIO_GAIN", "Track Gain")
    lstvMain.ColumnHeaders(5).Text = GetLocalString("frmMain.LCL_COLUMN_RADIO_CLIP", "clip(Track)")
    lstvMain.ColumnHeaders(6).Text = GetLocalString("frmMain.LCL_COLUMN_MAXIMUM_NOCLIP", "Max Noclip Gain")
    lstvMain.ColumnHeaders(7).Text = GetLocalString("frmMain.LCL_COLUMN_ALBUM_VOLUME", "Album Volume")
    lstvMain.ColumnHeaders(8).Text = GetLocalString("frmMain.LCL_COLUMN_ALBUM_GAIN", "Album Gain")
    lstvMain.ColumnHeaders(9).Text = GetLocalString("frmMain.LCL_COLUMN_ALBUM_CLIP", "clip(Album)")
    lstvMain.ColumnHeaders(10).Text = GetLocalString("frmMain.LCL_COLUMN_PATH", "Path")
    lstvMain.ColumnHeaders(11).Text = GetLocalString("frmMain.LCL_COLUMN_FILE", "File")
    lstvMain.ColumnHeaders(12).Text = GetLocalString("frmMain.LCL_COLUMN_MAXIMUM_AMPLITUDE", "Curr Max Amp")
    
    ResetColumnWidths
    
    targetChange 'The simplest way to refresh the analysis results in the file list
    
    Exit Sub
mnuLanguage_Click_Error:
    HandleError ("mnuLanguage_Click")
End Sub

Private Sub mnuLoadAnalysis_Click()
    On Error GoTo mnuLoadAnalysis_Click_Error
    Dim arrFiles() As String
    Dim strFileName As String
    Dim lngFlags As Long
    Dim strFilter As String
    
    strFilter = GetLocalString("frmMain.LCL_COMMA_SEPARATED", "Comma-separated files") & _
        " (*.m3g;*.csv)" & vbNullChar & "*.m3g;*.csv" & vbNullChar & _
        GetLocalString("frmMain.LCL_OPEN_FILE_FILTER2", "All files") & _
        " (*.*)" & vbNullChar & "*.*" & vbNullChar
        
        lngFlags = ahtOFN_ALLOWMULTISELECT Or _
        ahtOFN_EXPLORER Or _
        ahtOFN_LONGNAMES Or _
        ahtOFN_FILEMUSTEXIST Or _
        ahtOFN_HIDEREADONLY
    
    strFileName = """"
    strFileName = ahtCommonFileOpenSave(lngFlags, strSaveLogsPath, strFilter, 0, , strSaveLogsFile, , Me.hWnd, True)
    If Len(strFileName) = 0 Then
        Exit Sub
    End If
    
    
    If InStr(strFileName, vbNullChar) > 0 Then
        arrFiles = Split(strFileName, vbNullChar)
        strSaveLogsFile = ""
    Else
        ReDim arrFiles(0 To 1) As String
        arrFiles(0) = Left$(strFileName, InStrRev(strFileName, "\") - 1)
        arrFiles(1) = Mid$(strFileName, InStrRev(strFileName, "\") + 1)
        strSaveLogsFile = arrFiles(1)
    End If
    
    strSaveLogsPath = arrFiles(0)
    
    LoadGainAnalysis arrFiles, lstvMain, flsMaster
    Exit Sub

mnuLoadAnalysis_Click_Error:
    HandleError ("mnuLoadAnalysis_Click")
End Sub

Private Sub mnuLogs_Click()
    On Error GoTo mnuLogs_Click_Error
    If mnuAlwaysTop.Checked Then
        SetWindowPos hWnd, conHwndNoTopmost, 0, 0, 0, 0, &H10 Or &H1 Or &H2
    End If
    
    frmLogs.Show vbModal, Me
    
    If mnuAlwaysTop.Checked Then
        SetWindowPos hWnd, conHwndTopmost, 0, 0, 0, 0, &H10 Or &H1 Or &H2
    End If
    
    Exit Sub
    
mnuLogs_Click_Error:
    HandleError "mnuLogs_Click"
End Sub

Private Sub mnuMaxAmp_Click()
    blnMaxAmpOnly = True
    Call Radio
    blnMaxAmpOnly = False
End Sub

Private Sub mnuMaxNoClipGain_Click()
    Call MaxNoClipGain
End Sub

Private Sub mnuPathSepFile_Click()
    On Error GoTo mnuPathSepFile_Click_Error
    If intPathFileStat <> PATHSEPFILE Then
        If PathFileDisplay(PATHSEPFILE) Then
            mnuPathWithFile.Checked = False
            mnuFileOnly.Checked = False
            mnuPathSepFile.Checked = True
        End If
    End If
    
    Exit Sub
    
mnuPathSepFile_Click_Error:
    HandleError "mnuPathSepFile_Click"
End Sub

Private Sub mnuPathWithFile_Click()
    On Error GoTo mnuPathWithFile_Click_Error
    If intPathFileStat <> PATHFILE Then
        If PathFileDisplay(PATHFILE) Then
            mnuPathSepFile.Checked = False
            mnuFileOnly.Checked = False
            mnuPathWithFile.Checked = True
        End If
    End If
    
    Exit Sub
    
mnuPathWithFile_Click_Error:
    HandleError "mnuPathWithFile_Click"
End Sub

Private Sub mnuPopupAlbum_Click()
    mnuAlbum_Click
End Sub

Private Sub mnuPopupAlbumGain_Click()
    mnuAlbumGain_Click
End Sub

Private Sub mnuPopupClearAnalysis_Click()
    mnuClearAnalysis_Click
End Sub

Private Sub mnuPopupConstantGain_Click()
    mnuConstantGain_Click
End Sub

Private Sub mnuPopupGroupNoclip_Click()
    mnuGroupNoClip_Click
End Sub

Private Sub mnuPopupMaxAmp_Click()
    mnuMaxAmp_Click
End Sub

Private Sub mnuPopupMaxNoclipGain_Click()
    mnuMaxNoClipGain_Click
End Sub

Private Sub mnuPopupRadio_Click()
    mnuRadio_Click
End Sub

Private Sub mnuPopupRadioGain_Click()
    mnuRadioGain_Click
End Sub

Private Sub mnuPopupRemoveTags_Click()
    Call DoFileAction(taDeleteTags)
End Sub

Private Sub mnuPopupUndoGain_Click()
    Call DoFileAction(taUndoGain)
End Sub

Private Sub mnuRadio_Click()
    On Error GoTo mnuRadio_Click_Error

    Radio
    
    Exit Sub
    
mnuRadio_Click_Error:
    HandleError "mnuRadio_Click"
End Sub

Private Sub mnuRadioGain_Click()
    On Error GoTo mnuRadioGain_Click_Error

    RadioGain
    
    Exit Sub
    
mnuRadioGain_Click_Error:
    HandleError "mnuRadioGain_Click"
End Sub

Private Sub mnuReCalcTags_Click()
    mnuReCalcTags.Checked = Not mnuReCalcTags.Checked
    If mnuReCalcTags.Checked Then
        mnuSkipTags.Checked = False
        WarnSkipTags False
    End If
End Sub

Private Sub mnuReckless_Click()
    mnuReckless.Checked = Not mnuReckless.Checked
    
    If mnuReckless.Checked And blnRecklessWarning Then
        If mnuAlwaysTop.Checked Then
            SetWindowPos hWnd, conHwndNoTopmost, 0, 0, 0, 0, &H10 Or &H1 Or &H2
        End If
        
        frmLayerCheckWarning.Show vbModal, Me

        If mnuAlwaysTop.Checked Then
            SetWindowPos hWnd, conHwndTopmost, 0, 0, 0, 0, &H10 Or &H1 Or &H2
        End If
    End If
End Sub

Private Sub mnuResetColumns_Click()
    ResetColumnWidths
End Sub

Private Sub mnuResetWarnings_Click()
    Me.blnResetWarn = True
    Me.blnStereoWarning = True
    Me.blnRecklessWarning = True
    Me.blnSkipTagsWarn = True
End Sub

Private Sub mnuSaveAnalysis_Click()
    On Error GoTo mnuSaveAnalysis_Click_Error
    Dim strFileName As String
    Dim lngFlags As Long
    Dim strFilter As String
    
    strFilter = GetLocalString("frmMain.LCL_COMMA_SEPARATED", "Comma-separated files") & _
        " (*.m3g;*.csv)" & vbNullChar & "*.m3g;*.csv" & vbNullChar & _
        GetLocalString("frmMain.LCL_OPEN_FILE_FILTER2", "All files") & _
        " (*.*)" & vbNullChar & "*.*" & vbNullChar
    
        lngFlags = ahtOFN_EXPLORER Or _
        ahtOFN_LONGNAMES Or _
        ahtOFN_PATHMUSTEXIST Or _
        ahtOFN_HIDEREADONLY
    
    strFileName = ""
    strFileName = ahtCommonFileOpenSave(lngFlags, strSaveLogsPath, strFilter, 0, , strSaveLogsFile, , Me.hWnd, False)
    If Len(strFileName) = 0 Then
        Exit Sub
    End If
    
    strSaveLogsPath = Left$(strFileName, InStrRev(strFileName, "\") - 1)
    strSaveLogsFile = Mid$(strFileName, InStrRev(strFileName, "\") + 1)
    
    
    SaveGainAnalysis strFileName, lstvMain, flsMaster
    
    Exit Sub
    
mnuSaveAnalysis_Click_Error:
    HandleError ("mnuSaveAnalysis_Click")
End Sub

Private Sub mnuSelectAll_Click()
    Dim itmX As ListItem
    
    For Each itmX In lstvMain.ListItems
        itmX.Selected = True
    Next
End Sub

Private Sub mnuSelectedFiles_Click()
    mnuSelectedFiles.Checked = Not mnuSelectedFiles.Checked
End Sub

Private Sub mnuSelectNone_Click()
    Dim itmX As ListItem
    
    For Each itmX In lstvMain.ListItems
        itmX.Selected = False
    Next
End Sub

Private Sub mnuSelectReverse_Click()
    Dim itmX As ListItem
    
    For Each itmX In lstvMain.ListItems
        itmX.Selected = Not itmX.Selected
    Next
End Sub

Private Sub mnuSkipTags_Click()
    mnuSkipTags.Checked = Not mnuSkipTags.Checked
    If mnuSkipTags.Checked Then
        mnuReCalcTags.Checked = False
    End If
    WarnSkipTags mnuSkipTags.Checked
End Sub

Private Sub mnuSkipTagsWhileAdding_Click()
    mnuSkipTagsWhileAdding.Checked = Not mnuSkipTagsWhileAdding.Checked
End Sub

Private Sub mnuSysTray_Click()
    mnuSysTray.Checked = Not mnuSysTray.Checked
End Sub

Private Sub mnuToolBarBig_Click()
    If intToolBarSize <> 3 Then
        intToolBarSize = 3
        ToolBarResize
    End If
End Sub

Private Sub mnuToolbarNone_Click()
    If intToolBarSize <> 0 Then
        intToolBarSize = 0
        ToolBarResize
    End If
End Sub

Private Sub mnuToolbarSmall_Click()
    If intToolBarSize <> 2 Then
        intToolBarSize = 2
        ToolBarResize
    End If
End Sub

Private Sub mnuToolbarText_Click()
    If intToolBarSize <> 1 Then
        intToolBarSize = 1
        ToolBarResize
    End If
End Sub

Private Sub mnuUndoGain_Click()
    Call DoFileAction(taUndoGain)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error GoTo Toolbar1_ButtonClick_Error
    
    Select Case Button.Key
    Case "analysis"
        Select Case intAnalysisMode
        Case 1
            Radio
        Case 2
            Album
        End Select
    Case "gain"
        Select Case intGainMode
        Case 1
            RadioGain
        Case 2
            AlbumGain
        Case 3
            ConstGain
        End Select
    Case "addfiles"
        AddFiles
    Case "addfolder"
        AddFolder
    Case "clearfiles"
        ClearFiles
    Case "clearall"
        ClearAll
    End Select
    
    Exit Sub
    
Toolbar1_ButtonClick_Error:
    HandleError "Toolbar1_ButtonClick"
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    On Error GoTo Toolbar1_ButtonMenuClick_Error

    Select Case ButtonMenu.Key
    Case "radioanalysis"
        intAnalysisMode = 1
        Toolbar1.Buttons("analysis").ButtonMenus("albumanalysis").Visible = True
        Toolbar1.Buttons("analysis").ButtonMenus("radioanalysis").Visible = False
        If intToolBarSize = 2 Then
            Toolbar1.Buttons("analysis").Caption = GetLocalString("frmMain.LCL_SMALL_RADIO", "Track")
        Else
            Toolbar1.Buttons("analysis").Caption = GetLocalString("frmMain.Button1Menu1.Text", "Track Analysis")
        End If
        If intToolBarSize <> 1 Then Toolbar1.Buttons("analysis").Image = 1
        
    Case "albumanalysis"
        intAnalysisMode = 2
        Toolbar1.Buttons("analysis").ButtonMenus("albumanalysis").Visible = False
        Toolbar1.Buttons("analysis").ButtonMenus("radioanalysis").Visible = True
        If intToolBarSize = 2 Then
            Toolbar1.Buttons("analysis").Caption = GetLocalString("frmMain.LCL_SMALL_ALBUM", "Album")
        Else
            Toolbar1.Buttons("analysis").Caption = GetLocalString("frmMain.Button1Menu2.Text", "Album Analysis")
        End If
        If intToolBarSize <> 1 Then Toolbar1.Buttons("analysis").Image = 2
        
    Case "clearanalysis"
        ResetAnalysis
            
    Case "radiogain"
        intGainMode = 1
        Toolbar1.Buttons("gain").ButtonMenus("radiogain").Visible = False
        Toolbar1.Buttons("gain").ButtonMenus("albumgain").Visible = True
        Toolbar1.Buttons("gain").ButtonMenus("constantgain").Visible = True
        If intToolBarSize = 2 Then
            Toolbar1.Buttons("gain").Caption = GetLocalString("frmMain.LCL_SMALL_RADIO", "Track")
        Else
            Toolbar1.Buttons("gain").Caption = GetLocalString("frmMain.Button2Menu1.Text", "Track Gain")
        End If
        If intToolBarSize <> 1 Then Toolbar1.Buttons("gain").Image = 3
        
    Case "albumgain"
        intGainMode = 2
        Toolbar1.Buttons("gain").ButtonMenus("radiogain").Visible = True
        Toolbar1.Buttons("gain").ButtonMenus("albumgain").Visible = False
        Toolbar1.Buttons("gain").ButtonMenus("constantgain").Visible = True
        If intToolBarSize = 2 Then
            Toolbar1.Buttons("gain").Caption = GetLocalString("frmMain.LCL_SMALL_ALBUM", "Album")
        Else
            Toolbar1.Buttons("gain").Caption = GetLocalString("frmMain.Button2Menu2.Text", "Album Gain")
        End If
        If intToolBarSize <> 1 Then Toolbar1.Buttons("gain").Image = 4
        
    Case "constantgain"
        intGainMode = 3
        Toolbar1.Buttons("gain").ButtonMenus("radiogain").Visible = True
        Toolbar1.Buttons("gain").ButtonMenus("albumgain").Visible = True
        Toolbar1.Buttons("gain").ButtonMenus("constantgain").Visible = False
        If intToolBarSize = 2 Then
            Toolbar1.Buttons("gain").Caption = GetLocalString("frmMain.LCL_SMALL_CONSTANT", "Const")
        Else
            Toolbar1.Buttons("gain").Caption = GetLocalString("frmMain.Button2Menu3.Text", "Constant Gain")
        End If
        If intToolBarSize <> 1 Then Toolbar1.Buttons("gain").Image = 5
        
    End Select
    
    Exit Sub
    
Toolbar1_ButtonMenuClick_Error:
    HandleError "Toolbar1_ButtonMenuClick"
End Sub

Private Sub txtAlbumMonitor_Change()
    On Error GoTo txtAlbumMonitor_Change_Error

    Dim strLine As String
    Dim intLF As Integer
    Dim arrStuff() As String
    Dim itmX As ListItem
    Dim mp3Inf As Mp3Info
    Dim i As Integer
    
    If blnDoingAlbum Then
        blnDoingAlbum = False
        intLF = InStr(txtAlbumMonitor.Text, vbCrLf)
        While intLF > 0
            strLine = Left$(txtAlbumMonitor.Text, intLF - 1)
            
            If blnStartAlbum Then
                blnStartAlbum = False
            Else
                arrStuff = Split(strLine, vbTab, , vbBinaryCompare)
                If UBound(arrStuff) < 3 Then
                    If Not blnCancel And (UBound(arrStuff) > -1) Then
                        LogErr GetLocalString("frmMain.LCL_ERROR_ANALYZING", "Error while analyzing") & ": " & arrStuff(0)
                        For i = 1 To lstvMain.ListItems.Count
                            If lstvMain.ListItems(i).Text = arrStuff(0) Then
                                lstvMain.ListItems(i).Tag = "X"
                                Exit For
                            End If
                        Next i
                    End If
                Else
                    If arrStuff(0) = """Album""" Then
                        For Each itmX In lstvMain.ListItems
                            If (Not mnuSelectedFiles.Checked) Or (itmX.Checked) Then
                                If itmX.Tag = "Y" Then
                                    Set mp3Inf = flsMaster.Item(itmX.Key)
                                    If mp3Inf.RadiodBGain <> NOREALNUM Then
                                        LogAlbumAnalysis itmX.Text, CDbl(Val(arrStuff(2)))
                                        mp3Inf.AlbumdBGain = CDbl(Val(arrStuff(2)))
                                    End If
                                    DispJunk itmX, mp3Inf
                                    Set mp3Inf = Nothing
                                End If
                            End If
                        Next
                    Else
                        prgTot.Value = prgTot.Value + 1
                        If prgTot.Max > 100 Then
                            UpdateCaptionPercentage Format$(CSng(prgTot.Value) * 100! / CSng(prgTot.Max), "0.0")
                        Else
                            UpdateCaptionPercentage CLng((prgTot.Value * 100) / prgTot.Max)
                        End If
                        For i = 1 To lstvMain.ListItems.Count
                            If lstvMain.ListItems(i).Text = arrStuff(0) Then
                                Set itmX = lstvMain.ListItems(i)
                                i = lstvMain.ListItems.Count
                            End If
                        Next i
                        If itmX Is Nothing Then
                            'MsgBox "You have found the one bug I'm still working on before the official 1.0 release. If you e-mail me at mp3gain@hotmail.com and tell me exactly what happened, I might be able to solve the problem more quickly."
#If GLENDEBUG Then
                            Dim intFileNum As Integer
                            Dim lngDebugFileLen As Long
                            Dim blstItem() As Byte
                            intFileNum = FreeFile
                            Open "C:\MP3GainDebugJunk.txt" For Output As #intFileNum
                            Print #intFileNum, "----------------------------------"
                            Print #intFileNum, "ListItem(1).Text UNICODE:"
                            Close #intFileNum
                            lngDebugFileLen = FileLen("C:\MP3GainDebugJunk.txt")
                            lngDebugFileLen = lngDebugFileLen + 1
                            
                            Open "C:\MP3GainDebugJunk.txt" For Binary Access Read Write As #intFileNum
                            BSTRtoUniBytes lstvMain.ListItems(1).Text, blstItem
                            Put #intFileNum, lngDebugFileLen, blstItem
                            Close #intFileNum
                            
                            Open "C:\MP3GainDebugJunk.txt" For Append As #intFileNum
                            Print #intFileNum, vbCrLf & "-----" & vbCrLf & "arrStuff(0):"
                            Print #intFileNum, arrStuff(0)
                            Print #intFileNum, "-----" & vbCrLf & "strLine:"
                            Print #intFileNum, strLine
                            Print #intFileNum, "-----" & vbCrLf & "strDebugCmdLineCopy:"
                            Print #intFileNum, strDebugCmdLineCopy
                            Print #intFileNum, "-----" & vbCrLf & "strDebugOutputCopy:"
                            Print #intFileNum, strDebugOutputCopy
                            Print #intFileNum, "-----" & vbCrLf & "bytDebugOutput:"
                            Close #intFileNum
                            lngDebugFileLen = FileLen("C:\MP3GainDebugJunk.txt")
                            lngDebugFileLen = lngDebugFileLen + 1
                            Open "C:\MP3GainDebugJunk.txt" For Binary Access Read Write As #intFileNum
                            Put #intFileNum, lngDebugFileLen, bytDebugOutput
                            Close #intFileNum
                            MsgBox "Send Glen the file ""C:\MP3GainDebugJunk.txt"""
#End If
                        Else
                            Set mp3Inf = flsMaster.Item(itmX.Key)
                            mp3Inf.ResetVals
                            mp3Inf.RadiodBGain = CDbl(Val(arrStuff(2)))
                            mp3Inf.CurrMaxAmp = CDbl(Val(arrStuff(3)))
                            LogRadioAnalysis itmX.Text, Round(CDbl(Val(arrStuff(3)))), _
                                CDbl(Val(arrStuff(2)))
                            mp3Inf.CurrMaxGain = CInt(Val(arrStuff(4)))
                            mp3Inf.CurrMinGain = CInt(Val(arrStuff(5)))
                            mp3Inf.ModifydBGain = dblGainAdjust
                            DispJunk itmX, mp3Inf
                            Set mp3Inf = Nothing
                            Set itmX = Nothing
                        End If
                    End If
                End If
            End If
            
            txtAlbumMonitor.Text = Mid$(txtAlbumMonitor.Text, intLF + 2)
            intLF = InStr(txtAlbumMonitor.Text, vbCrLf)
        Wend
        blnDoingAlbum = True
        Me.Refresh
    End If
    
    Exit Sub
    
txtAlbumMonitor_Change_Error:
    HandleError "txtAlbumMonitor_Change"
End Sub

Private Sub txtProgWatch_Change()
    On Error GoTo txtProgWatch_Change_Error
    Dim intPos As Integer
    Dim strChunk As String
    Dim strSplit() As String
    
    If blnTxtProgChanging Then Exit Sub
        
    intPos = InStr(Me.txtProgWatch.Text, Chr$(13))
    While intPos > 0
        strChunk = Left$(Me.txtProgWatch.Text, intPos - 1)
        strSplit = Split(strChunk, " ")
        If UBound(strSplit) = 6 Then
            If strSplit(3) = "of" And strSplit(5) = "bytes" Then
                On Error Resume Next
                Me.prgFile.Value = Left$(strSplit(2), 1)
                On Error GoTo txtProgWatch_Change_Error
            End If
        ElseIf UBound(strSplit) = 5 Then
            If strSplit(2) = "of" And strSplit(4) = "bytes" Then
                On Error Resume Next
                Me.prgFile.Value = Left$(strSplit(1), 2)
                On Error GoTo txtProgWatch_Change_Error
            End If
        End If
        blnTxtProgChanging = True
        Me.txtProgWatch.Text = Mid$(Me.txtProgWatch.Text, intPos + 1)
        blnTxtProgChanging = False
        intPos = InStr(Me.txtProgWatch.Text, Chr$(13))
    Wend
    
    Exit Sub

txtProgWatch_Change_Error:
    HandleError "txtProgWatch_Change"
End Sub

Private Sub targetChange()
    If IsNumeric(txtTargetInt.Text) And IsNumeric(txtTargetDec.Text) Then
        txtTargetInt.Text = Int(txtTargetInt.Text)
        txtTargetDec.Text = Int(txtTargetDec.Text)
        If txtTargetInt.Text > 500 Then
            txtTargetInt.Text = 500
        End If
        If Len(txtTargetDec.Text) > 1 Then
            txtTargetDec.Text = Mid$(txtTargetDec.Text, 1, 1)
            txtTargetDec.SelStart = 0
            txtTargetDec.SelLength = Len(txtTargetDec.Text)
        End If
        dblGainAdjust = Round(CDbl(txtTargetInt.Text) + (CDbl(txtTargetDec.Text) / 10#), 1) - DEFAULTTARGET
        Call GainAdjust(dblGainAdjust)
    End If
End Sub

Private Sub txtTargetDec_KeyPress(KeyAscii As Integer)
    If Len(txtTargetDec.Text) = 1 Then
        blnTargetIsChanging = True
        txtTargetDec.Text = ""
        blnTargetIsChanging = False
    End If
End Sub

Private Sub txtTargetInt_Change()
    If Not blnTargetIsChanging Then
        blnTargetIsChanging = True
        targetChange
        blnTargetIsChanging = False
    End If
End Sub

Private Sub txtTargetDec_Change()
    If Not blnTargetIsChanging Then
        blnTargetIsChanging = True
        targetChange
        blnTargetIsChanging = False
    End If
End Sub

Private Sub txtTargetInt_GotFocus()
    txtTargetInt.SelStart = 0
    txtTargetInt.SelLength = Len(txtTargetInt.Text)
End Sub

Private Sub txtTargetDec_GotFocus()
    txtTargetDec.SelStart = 0
    txtTargetDec.SelLength = Len(txtTargetDec.Text)
End Sub

Private Sub txtTargetInt_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc(lblDecimal.Caption) Then
        txtTargetDec.SetFocus
    End If
End Sub

Private Sub txtTargetInt_LostFocus()
    If Not IsNumeric(txtTargetInt.Text) Then
        txtTargetInt.Text = DEFAULTTARGET
    Else
        txtTargetInt.Text = Int(txtTargetInt.Text)
    End If
End Sub

Private Sub txtTargetDec_LostFocus()
    If Not IsNumeric(txtTargetDec.Text) Then
        txtTargetDec.Text = 0
    Else
        txtTargetDec.Text = Int(txtTargetDec.Text)
    End If
End Sub


Private Sub mPopExit_Click()
    'called when user clicks the popup menu Exit command
    If blnCurrentlyProcessing Then
        blnCancel = True
        blnExitWhenDone = True
        stbStat.Panels(1).Text = GetLocalString("frmMain.LCL_EXITING", "Exiting...")
    Else
        Unload Me
    End If
End Sub

Private Sub mnuPopRestore_Click()
    'called when the user clicks the popup menu Restore command
    Dim Result As Long
    Me.WindowState = vbNormal
    Result = SetForegroundWindow(Me.hWnd)
    Me.Show
End Sub

Public Sub AddListOfFiles(arrInputFiles() As String, strListDrive As String, strListPath As String)
    On Error GoTo AddListOfFiles_Error
    Dim i As Integer
    Dim strPathFile As String
    Dim strPath As String
    Dim strFile As String
    Dim arrFiles() As String
    Dim fl
    Dim flFormat As String
    Dim intDivider As Long
    Dim lngIsDirectory As Long
    
    i = 1
    For Each fl In arrInputFiles
        flFormat = fl
        If Mid$(flFormat, 2, 1) <> ":" And Left$(flFormat, 2) <> "\\" Then
            If Left$(flFormat, 1) = "\" Then
                flFormat = strListDrive & flFormat
            Else
                flFormat = strListPath & flFormat
            End If
        End If

        Select Case LCase$(Right$(flFormat, 4))
        Case ".mp3"
            flFormat = CleanOutRelativePathInfo(flFormat)
            AddSingleFile (flFormat)
        Case ".m3u"
            flFormat = CleanOutRelativePathInfo(flFormat)
            AddM3U (flFormat)
        Case ".m3g", ".csv"
            flFormat = CleanOutRelativePathInfo(flFormat)
            intDivider = InStrRev(flFormat, "\")
            strPath = Left$(flFormat, intDivider - 1)
            strFile = Mid$(flFormat, intDivider + 1)
            ReDim Preserve arrFiles(0 To i) As String
            If i > 1 And strPath <> arrFiles(0) Then
                i = 1
                LoadGainAnalysis arrFiles, lstvMain, flsMaster
                ReDim arrFiles(0 To 1) As String
            End If
            arrFiles(0) = strPath
            arrFiles(i) = strFile
            i = i + 1
        Case Else
            If Right$(flFormat, 1) <> "\" Then
                flFormat = flFormat & "\"
            End If
            flFormat = CleanOutRelativePathInfo(flFormat)
            On Error Resume Next 'in case file doesn't exist at all
            lngIsDirectory = GetFileAttributes(flFormat) And vbDirectory
            On Error GoTo AddListOfFiles_Error
            If lngIsDirectory = vbDirectory Then
                Dim colFolderList As Collection
                Set colFolderList = New Collection
                colFolderList.Add flFormat
                While colFolderList.Count > 0
                    Dim strFolder As String
                    strFolder = colFolderList(1)
                    colFolderList.Remove (1)
                    AddFolderFiles strFolder, colFolderList
                Wend
                Set colFolderList = Nothing
            End If
        End Select
    Next fl
    If i > 1 Then LoadGainAnalysis arrFiles, lstvMain, flsMaster
    
    Exit Sub
AddListOfFiles_Error:
    HandleError "AddListOfFiles"
    On Error Resume Next
    doSortColumn
End Sub

Public Sub StartupParseCommand()
    On Error GoTo StartupParseCommand_Error
    Dim strCmdLine As String
    Dim arrFiles() As String
    Dim i As Long
    Dim intDivider As Long
    Dim strCurDrive As String
    Dim strCurPath As String
    
    strCurPath = CurDir
    If Right$(strCurPath, 1) <> "\" Then strCurPath = strCurPath & "\"
    strCurDrive = GetDrivePartThing(strCurPath)

    strCmdLine = Command
    i = 0
    While Len(strCmdLine) > 0
        ReDim Preserve arrFiles(0 To i) As String
        If Mid$(strCmdLine, 1, 1) = """" Then
            arrFiles(i) = Mid$(strCmdLine, 2, InStr(2, strCmdLine, """") - 2)
            strCmdLine = Mid$(strCmdLine, Len(arrFiles(i)) + 4)
        Else
            intDivider = InStr(strCmdLine, " ")
            If intDivider = 0 Then intDivider = InStr(strCmdLine, vbTab)
            If intDivider = 0 Then intDivider = Len(strCmdLine) + 1
            arrFiles(i) = Mid$(strCmdLine, 1, intDivider - 1)
            strCmdLine = Mid$(strCmdLine, Len(arrFiles(i)) + 2)
        End If
        i = i + 1
    Wend
    
    If i > 0 Then
        AddListOfFiles arrFiles, strCurDrive, strCurPath
    End If
    
    Exit Sub
StartupParseCommand_Error:
    
End Sub

Sub ShowBulkErrors()
    Dim strMsg As String
    Dim mbrResult As VbMsgBoxResult
    
    If glErrCount = 1 Then
        strMsg = Replace( _
            GetLocalString("frmMain.LCL_SHOW_ONE_ERROR_COUNT" _
            , "%%COUNT%% error during processing.") _
            , "%%COUNT%%", glErrCount)
    ElseIf glErrCount > 1 Then
        strMsg = Replace( _
            GetLocalString("frmMain.LCL_SHOW_MANY_ERROR_COUNT" _
            , "%%COUNT%% errors during processing.") _
            , "%%COUNT%%", glErrCount)
    End If
    
    If glErrCount > 0 Then
        mbrResult = MsgBox(strMsg & vbCrLf & GetLocalString("frmMain.LCL_VIEW_LOG", "View Error Log?"), vbExclamation Or vbYesNo Or vbDefaultButton2)
        If mbrResult = vbYes Then
            If ShellExecute(0&, vbNullString, strErrLog, _
                vbNullString, vbNullString, vbNormalFocus) < 33 Then
                    MsgBox Replace( _
                    GetLocalString("frmMain.LCL_CANT_VIEW_LOG", "Error trying to auto-open the log file %%filename%%. You will need to open the file from Windows Explorer instead.") _
                    , "%%filename%%", strErrLog), vbExclamation
            End If
                        
        End If
    End If
    
End Sub

Private Sub DeleteFileTags()
On Error GoTo DeleteFileTags_Error
    Dim itmX As ListItem
    Dim strCmd As String
    Dim strBlah As String
    Dim lngRetVal As Long
    
    For Each itmX In lstvMain.ListItems
        If ((Not mnuSelectedFiles.Checked) Or (itmX.Checked)) And (itmX.Tag = "Y") Then
            If Not blnCancel Then
                strCmd = """" & strAppPath & "mp3Gain"" /o /s d "
                
                stbStat.Panels(1).Text = Replace(GetLocalString("frmMain.LCL_REMOVING_TAGS", _
                    "Removing MP3Gain tags from %%filename%%"), _
                    "%%filename%%", itmX.Text)
                            
                If Not blnShowFileStatus Then
                    strCmd = strCmd & "/q "
                End If
                
                If mnuKeepTime.Checked Then
                    strCmd = strCmd & "/p "
                End If
                
                strCmd = strCmd & """" & itmX.Text & """"
                
                Refresh
                strBlah = ""
                
                If blnShowFileStatus Then
                    lngRetVal = GetCommandOutput(strBlah, strCmd, strAppPath, True, True, False, 100, , Me.txtProgWatch)
                Else
                    lngRetVal = GetCommandOutput(strBlah, strCmd, strAppPath, True, True, False, 100)
                End If
                
                Me.prgFile.Value = 0
                
                If InStr(LCase$(strBlah), "not enough temp space on disk") Then
                    LogErr Replace(GetLocalString("frmMain.LCL_NO_TEMP_SPACE_1", _
                        "Not enough temp space on disk to modify %%filename%%"), _
                        "%%filename%%", itmX.Text)
                ElseIf InStr(LCase$(strBlah), "can't open") Then
                    LogErr Replace(GetLocalString("frmMain.LCL_MODIFY_ERROR", _
                        "Can't modify file %%filename%%"), "%%filename%%", itmX.Text)
                ElseIf lngRetVal <> 1 Then
                    If Not blnCancel Then
                        If strBlah <> "" Then
                            LogErr GetLocalString("frmMain.LCL_BACKEND_ERROR", "Error running mp3gain.exe") & ":" & vbCrLf & strBlah
                        Else
                            LogErr GetLocalString("frmMain.LCL_BACKEND_ERROR", "Error running mp3gain.exe")
                        End If
                    End If
                Else
                    'nada
                End If
                prgTot.Value = prgTot.Value + 1
                If prgTot.Max > 100 Then
                    UpdateCaptionPercentage Format$(CSng(prgTot.Value) * 100! / CSng(prgTot.Max), "0.0")
                Else
                    UpdateCaptionPercentage CLng((prgTot.Value * 100) / prgTot.Max)
                End If
            End If
        End If
    Next
        
    stbStat.Panels(1).Text = ""
    
    If blnExitWhenDone Then Unload Me
    
    Exit Sub
DeleteFileTags_Error:
    HandleError "DeleteFileTags"
End Sub

Private Sub UndoFileGain()
On Error GoTo UndoFileGain_Error

    Dim strBlah As String
    Dim strCmd As String
    Dim lngRetVal As Long
    Dim itmX As ListItem
    Dim dbldB As Double
    Dim mp3Inf As Mp3Info
    Dim strDBChange As String
    Dim intGainChange As Integer
    
    For Each itmX In lstvMain.ListItems
        If ((Not mnuSelectedFiles.Checked) Or (itmX.Checked)) And (itmX.Tag = "Y") Then
            If Not blnCancel Then
                strCmd = """" & strAppPath & "mp3Gain"" /o /u "
                
                stbStat.Panels(1).Text = Replace(GetLocalString("frmMain.LCL_UNDOING_CHANGES", _
                    "Un-doing MP3Gain changes to %%filename%%"), _
                    "%%filename%%", itmX.Text)
                            
                If blnUseTempFiles Then
                    strCmd = strCmd & "/t "
                End If
                
                If Not blnShowFileStatus Then
                    strCmd = strCmd & "/q "
                End If
                
                If mnuKeepTime.Checked Then
                    strCmd = strCmd & "/p "
                End If
                
                If mnuReckless.Checked Then
                    strCmd = strCmd & "/f "
                End If
                                
                strCmd = strCmd & """" & itmX.Text & """"
                
                Refresh
                strBlah = ""
                
                If blnShowFileStatus Then
                    lngRetVal = GetCommandOutput(strBlah, strCmd, strAppPath, True, True, False, 100, , Me.txtProgWatch)
                Else
                    lngRetVal = GetCommandOutput(strBlah, strCmd, strAppPath, True, True, False, 100)
                End If
                
                Me.prgFile.Value = 0
                
                If InStr(LCase$(strBlah), "not enough temp space on disk") Then
                    LogErr Replace(GetLocalString("frmMain.LCL_NO_TEMP_SPACE_1", _
                        "Not enough temp space on disk to modify %%filename%%"), _
                        "%%filename%%", itmX.Text) & vbCrLf & GetLocalString("frmMain.LCL_NO_TEMP_SPACE_2", _
                        "Either clear space on disk, or go to ""Options->Advanced..."" and check the ""Do not use Temp files"" box.")
                ElseIf InStr(LCase$(strBlah), "can't open") Then
                    LogErr Replace(GetLocalString("frmMain.LCL_MODIFY_ERROR", _
                        "Can't modify file %%filename%%"), "%%filename%%", itmX.Text)
                ElseIf InStr(LCase$(strBlah), "can't adjust single channel") Then
                    LogErr Replace(GetLocalString("frmMain.LCL_NOT_STEREO", _
                        "%%filename%% is not a stereo or dual-channel mp3"), _
                        "%%filename%%", itmX.Text & vbCrLf)
                ElseIf lngRetVal <> 1 Then
                    If Not blnCancel Then
                        If strBlah <> "" Then
                            LogErr GetLocalString("frmMain.LCL_BACKEND_ERROR", "Error running mp3gain.exe") & ":" & vbCrLf & strBlah
                        Else
                            LogErr GetLocalString("frmMain.LCL_BACKEND_ERROR", "Error running mp3gain.exe")
                        End If
                    End If
                Else
                    Dim arrRows() As String
                    Dim arrValues() As String
                    
                    arrRows = Split(strBlah, vbCrLf)
                    arrValues = Split(arrRows(1), vbTab)
                    intGainChange = arrValues(1)
                    LogChange itmX.Text, intGainChange
                    Set mp3Inf = flsMaster.Item(itmX.Key)
                    mp3Inf.AlterDb -CDbl(intGainChange) * FIVELOG10TWO
                    DispJunk itmX, mp3Inf
                    Set mp3Inf = Nothing
                End If
                prgTot.Value = prgTot.Value + 1
                If prgTot.Max > 100 Then
                    UpdateCaptionPercentage Format$(CSng(prgTot.Value) * 100! / CSng(prgTot.Max), "0.0")
                Else
                    UpdateCaptionPercentage CLng((prgTot.Value * 100) / prgTot.Max)
                End If
            End If
        End If
    Next
        
    stbStat.Panels(1).Text = ""
    
    If blnExitWhenDone Then Unload Me
    
    Exit Sub

UndoFileGain_Error:
    HandleError "UndoFileGain"
    On Error Resume Next
    Me.cmdCancel.Default = False
    Me.cmdCancel.Enabled = False
    Me.MousePointer = vbDefault
    blnCurrentlyProcessing = False
    EnableJunk (True)
End Sub

Sub DoFileAction(doWhat As TagAction)
    On Error GoTo UndoGain_Error
    
    Dim itmX As ListItem
    
    glErrCount = 0
    
    blnCancel = False
    blnCurrentlyProcessing = True
    blnAllowProcessCancel = False
    EnableJunk (False)
    Me.MousePointer = vbArrowHourglass
    Me.cmdCancel.Enabled = True
    Me.cmdCancel.Default = True
    
    prgTot.Value = 0
    prgTot.Max = 1
    For Each itmX In lstvMain.ListItems
        itmX.Tag = "Y"
        If mnuSelectedFiles.Checked Then
            If itmX.Selected Then
                prgTot.Max = prgTot.Max + 1
                itmX.Checked = True
            Else
                itmX.Checked = False
            End If
        Else
            prgTot.Max = prgTot.Max + 1
        End If
    Next
    If prgTot.Max > 1 Then prgTot.Max = prgTot.Max - 1
    
    UpdateCaptionPercentage "0"
    
    Select Case doWhat
        Case taUndoGain
            Call UndoFileGain
        Case taDeleteTags
            Call DeleteFileTags
    End Select
    
    prgTot.Value = 0
    UpdateCaptionPercentage ""
    
    For Each itmX In lstvMain.ListItems
        itmX.Tag = ""
    Next
    
    Me.cmdCancel.Default = False
    Me.cmdCancel.Enabled = False
    Me.MousePointer = vbDefault
    EnableJunk (True)
    blnCurrentlyProcessing = False
    
    ShowBulkErrors
    
    
    Exit Sub
UndoGain_Error:
    HandleError "UndoGain"
    On Error Resume Next
    Me.cmdCancel.Default = False
    Me.cmdCancel.Enabled = False
    Me.MousePointer = vbDefault
    blnCurrentlyProcessing = False
    EnableJunk (True)
End Sub

Private Sub WarnSkipTags(blnChecked As Boolean)
    If blnChecked Then
        If blnSkipTagsWarn Then
            If mnuAlwaysTop.Checked Then
                SetWindowPos hWnd, conHwndNoTopmost, 0, 0, 0, 0, &H10 Or &H1 Or &H2
            End If
                
            frmSkipTagsWarn.Show vbModal, Me
    
            If mnuAlwaysTop.Checked Then
                SetWindowPos hWnd, conHwndTopmost, 0, 0, 0, 0, &H10 Or &H1 Or &H2
            End If
        End If
        lblNoUndo.Visible = True
    Else
        lblNoUndo.Visible = False
    End If
End Sub
