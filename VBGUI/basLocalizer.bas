Attribute VB_Name = "basLocalizer"
Option Explicit

Private mcollStrings As Collection

Private mbLoaded As Boolean

Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" _
    (ByVal lpAppName As String, _
     ByVal lpReturnedString As String, _
     ByVal nSize As Long, _
     ByVal lpFileName As String) As Long

Private Declare Function GetPrivateProfileSectionNames Lib "kernel32" Alias "GetPrivateProfileSectionNamesA" _
    (ByVal lpReturnedString As String, _
     ByVal nSize As Long, _
     ByVal lpFileName As String) As Long

Private Const ANSI_CHARSET = 0
Private Const DEFAULT_CHARSET = 1
Private Const SYMBOL_CHARSET = 2
Private Const MONO_FONT = 8
Private Const MAC_CHARSET = 77
Private Const SHIFTJIS_CHARSET = 128
Private Const HANGEUL_CHARSET = 129
Private Const JOHAB_CHARSET = 130
Private Const CHINESESIMPLIFIED_CHARSET = 134
Private Const CHINESEBIG5_CHARSET = 136
Private Const GREEK_CHARSET = 161
Private Const TURKISH_CHARSET = 162
Private Const HEBREW_CHARSET = 177
Private Const ARABIC_CHARSET = 178
Private Const BALTIC_CHARSET = 186
Private Const RUSSIAN_CHARSET = 204
Private Const THAI_CHARSET = 222
Private Const EASTEUROPE_CHARSET = 238
Private Const OEM_CHARSET = 255
Private Declare Function GetUserDefaultLCID Lib "kernel32" () As Long
Private userDefaultLCID As Long

Public strDefaultFont As String

Public Sub SetDefaultControlFontName()
    Dim i As Integer
    
    strDefaultFont = ""
    For i = 1 To Screen.FontCount
        If UCase$(Screen.Fonts(i)) = "TAHOMA" Then
            strDefaultFont = Screen.Fonts(i)
            Exit For
        End If
    Next
    If strDefaultFont = "" Then strDefaultFont = "MS Sans Serif"
End Sub

Public Sub SetProperFont(obj As Object)
    On Error GoTo ErrorSetProperFont
    Select Case userDefaultLCID
    Case &H404 ' Traditional Chinese
        obj.Charset = CHINESEBIG5_CHARSET
        obj.Name = ChrW(&H65B0) + ChrW(&H7D30) + ChrW(&H660E) _
            + ChrW(&H9AD4)   'New Ming-Li
        obj.Size = 9
    Case &H411 ' Japan
        obj.Charset = SHIFTJIS_CHARSET
        obj.Name = ChrW(&HFF2D) + ChrW(&HFF33) + ChrW(&H20) + _
            ChrW(&HFF30) + ChrW(&H30B4) + ChrW(&H30B7) + ChrW(&H30C3) + _
            ChrW(&H30AF)
        obj.Size = 9
    Case &H412 'Korea UserLCID
        obj.Charset = HANGEUL_CHARSET
        obj.Name = ChrW(&HAD74) + ChrW(&HB9BC)
        obj.Size = 9
    Case &H804 ' Simplified Chinese
        obj.Charset = CHINESESIMPLIFIED_CHARSET
        obj.Name = ChrW(&H5B8B) + ChrW(&H4F53)
        obj.Size = 9
    Case Else   ' The other countries
        obj.Charset = DEFAULT_CHARSET
        obj.Name = strDefaultFont   ' Get the default UI font.
        'obj.Size = 12 'HUGE, for testing purposes
    End Select
    Exit Sub
ErrorSetProperFont:
    Err.Number = Err
End Sub


Public Function DeMenufy(asMenuItem As String)
    Dim strTemp As String
    strTemp = Replace(asMenuItem, "&&", "^^TEMPAMP^^")
    strTemp = Replace(strTemp, "&", "")
    DeMenufy = Replace(strTemp, "^^TEMPAMP^^", "&")
End Function

Private Function LoadLocalStrings(asFileName As String) As Boolean
    On Error GoTo LoadLocalStrings_Error
    Dim lsTemp As String
    Dim laSections() As String
    Dim llSize As Long
    Dim llRet As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim laItems() As String
    Dim laItemSplit() As String
    Dim lsKey As String
    Dim lsValue As String
    
    LoadLocalStrings = True
    For i = mcollStrings.Count To 1 Step -1
        mcollStrings.Remove i
    Next i
    llSize = 200
    Do
        lsTemp = Space$(llSize)
        
        llRet = GetPrivateProfileSectionNames(lsTemp, llSize, asFileName)
        
        If llRet = 0 Then
            Err.Raise 20000, , "No section names"
        ElseIf llRet = llSize - 2 Then
            llSize = llSize * 2
        Else
            laSections = Split(Left$(lsTemp, InStrRev(lsTemp, vbNullChar) - 2), vbNullChar)
            Exit Do
        End If
    Loop

    For i = LBound(laSections) To UBound(laSections)
        'Load section from INI file
        llSize = 20000
        Do
            lsTemp = Space$(llSize)
            llRet = GetPrivateProfileSection(laSections(i), lsTemp, llSize, asFileName)
                
            If llRet = 0 Then
                Exit Do
            ElseIf llRet = llSize - 2 Then
                llSize = llSize * 2
            Else
                laItems = Split(Left$(lsTemp, InStrRev(lsTemp, vbNullChar) - 2), vbNullChar)
                Exit Do
            End If
        Loop
        
        For j = LBound(laItems) To UBound(laItems)
            laItemSplit = Split(laItems(j), "=")
            If UBound(laItemSplit) - LBound(laItemSplit) > 0 Then 'We have a key=value pair
                lsKey = laItemSplit(LBound(laItemSplit))
                lsValue = laItemSplit(LBound(laItemSplit) + 1)
                'just in case we have any text strings that have a "=" in them
                For k = LBound(laItemSplit) + 2 To UBound(laItemSplit)
                    lsValue = lsValue & "=" & laItemSplit(k)
                Next k
                mcollStrings.Add lsValue, laSections(i) & "." & lsKey
                '                    If UCase$(Left$(lsKey, 4)) <> "LCL_" Then
                '                        mcollOriginals.Add "(empty)", laSections(i) & "." & lsKey
                '                    End If
            End If
        Next j
    Next i
    
    Exit Function
LoadLocalStrings_Error:
    LoadLocalStrings = False
    MsgBox "Error " & Err.Number & " loading local language information from " & asFileName & vbCrLf & _
        vbCrLf & _
        "Line " & Erl & " - " & Err.Description & vbCrLf & _
        vbCrLf & _
        App.Title & " will continue, but in English", vbExclamation
    Err.Clear
End Function

Public Function ReloadLocalization(idx As Integer, asFile As String) As Boolean
    On Error GoTo ReloadLocalization_Error
    Dim strLocalHelp As String
    
    ReloadLocalization = True
    If idx = 0 Then
        mbLoaded = False
        ReloadLocalization = True
        App.HelpFile = strAppPath & "MP3Gain.chm"
        Exit Function
    End If
    
    mbLoaded = LoadLocalStrings(asFile)
    
    If mbLoaded Then
        strLocalHelp = strAppPath & "MP3Gain" & frmMain.mnuLanguage(idx).Caption & ".chm"
        If Len(Dir(strLocalHelp)) > 0 Then
            App.HelpFile = strLocalHelp
        Else
            App.HelpFile = strAppPath & "MP3Gain.chm"
        End If
    End If
    
    ReloadLocalization = mbLoaded
    
    Exit Function
    
ReloadLocalization_Error:
    ReloadLocalization = False
    HandleError ("ReloadLocalization")
End Function

Public Sub LoadLocalization(asFileName As String)
    On Error GoTo LoadLocalization_Error
    Dim i As Long
    Dim lsFile As String
    Dim j As Long
    Dim idx As Long
    Dim strLocalHelp As String
    
    Set mcollStrings = New Collection
    
    userDefaultLCID = GetUserDefaultLCID()
    
    mbLoaded = False
        
    frmMain.intCurLanguage = 0
    
    lsFile = Dir(strAppPath & "*.mp3gain.ini")
    If Len(lsFile) = 0 Then
        Exit Sub
    End If

    frmMain.intCurLanguage = 1
    frmMain.mnuLanguageList.Visible = True
    i = 1
    While Len(lsFile) > 0
        idx = i
        For j = 1 To i - 1
            If lsFile < frmMain.mnuLanguage(j).Tag Then
                idx = j
                j = i
            End If
        Next j
        Load frmMain.mnuLanguage(i)
        frmMain.mnuLanguage(i).Checked = False
        If idx < i Then
            For j = i To idx + 1 Step -1
                frmMain.mnuLanguage(j).Caption = frmMain.mnuLanguage(j - 1).Caption
                frmMain.mnuLanguage(j).Tag = frmMain.mnuLanguage(j - 1).Tag
            Next
        End If
        frmMain.mnuLanguage(idx).Caption = Mid$(lsFile, 1, Len(lsFile) - 12)
        frmMain.mnuLanguage(idx).Tag = lsFile
        lsFile = Dir
        i = i + 1
    Wend
        
        
    If (asFileName = "ORIGINAL") Then
        frmMain.intCurLanguage = 0
        Exit Sub
    End If
        
    For i = 1 To frmMain.mnuLanguage.Count - 1
        If UCase$(frmMain.mnuLanguage(i).Tag) = asFileName Then
            frmMain.intCurLanguage = i
            i = frmMain.mnuLanguage.Count
        End If
    Next i
    frmMain.mnuLanguage(0).Checked = False
    frmMain.mnuLanguage(frmMain.intCurLanguage).Checked = True
        
    lsFile = strAppPath & frmMain.mnuLanguage(frmMain.intCurLanguage).Tag
    
    mbLoaded = LoadLocalStrings(lsFile)
    
    If Not mbLoaded Then
        frmMain.mnuLanguage(frmMain.intCurLanguage).Checked = False
        frmMain.mnuLanguage(0).Checked = True
        frmMain.intCurLanguage = 0
        App.HelpFile = strAppPath & "MP3Gain.chm"
    Else
        strLocalHelp = strAppPath & "MP3Gain" & frmMain.mnuLanguage(frmMain.intCurLanguage).Caption & ".chm"
        If Len(Dir(strLocalHelp)) > 0 Then
            App.HelpFile = strLocalHelp
        Else
            App.HelpFile = strAppPath & "MP3Gain.chm"
        End If
    End If
        
    Exit Sub
    
LoadLocalization_Error:
    mbLoaded = False
    MsgBox "Error " & Err.Number & " loading local language information" & vbCrLf & _
        vbCrLf & _
        "Line " & Erl & " - " & Err.Description & vbCrLf & _
        vbCrLf & _
        App.Title & " will continue, but in English", vbExclamation
    Err.Clear
End Sub

Public Function GetLocalString(asKey As String, asDefault As String) As String
    GetLocalString = asDefault
    
    If Not mbLoaded Then Exit Function
    
    On Error Resume Next
    GetLocalString = mcollStrings(asKey)
End Function

Public Sub fillCaptions(afrmForm As Form)
    Dim lsFormName As String
    Dim ctrlItem
    
    If mbLoaded Then
        lsFormName = afrmForm.Name
        On Error Resume Next
        For Each ctrlItem In afrmForm.Controls
            SetProperFont ctrlItem.Font
            ctrlItem.Caption = GetLocalString(lsFormName & "." & ctrlItem.Name & ".Caption", ctrlItem.Caption)
        Next
        SetProperFont afrmForm.Font
        afrmForm.Caption = GetLocalString(lsFormName & "." & lsFormName & ".Caption", afrmForm.Caption)
    Else
        On Error Resume Next
        For Each ctrlItem In afrmForm.Controls
            SetProperFont ctrlItem.Font
        Next
        SetProperFont afrmForm.Font
    End If
End Sub
