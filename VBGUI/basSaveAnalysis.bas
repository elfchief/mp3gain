Attribute VB_Name = "basSaveAnalysis"
Option Explicit

Public Const NOREALNUM = -666.24601
Public Const conHwndTopmost = -1
Public Const conHwndNoTopmost = -2

Public Sub SaveGainAnalysis(strSaveFileName As String, lstInfo As ListView, colMp3Info As Collection)
On Error GoTo SaveGainAnalysis_Error
    Dim intFileNum As Integer
    Dim liBlah As ListItem
    Dim lngFileLen As Long
    Dim dteFileDate
    Dim strFile As String
    Dim strPath As String
    Dim strFileName As String
    Dim varRadio As Variant
    Dim dblMaxAmp As Double
    Dim varAlbum As Variant
    Dim mp3Inf As Mp3Info
   
    If Len(Dir(strSaveFileName)) > 0 Then
        If MsgBox( _
                GetLocalString("basSaveAnalysis.LCL_OVERWRITE_FILE", _
                "Overwrite existing file?"), vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
    
    intFileNum = FreeFile
    Open strSaveFileName For Output As #intFileNum
    
    For Each liBlah In lstInfo.ListItems
        strFile = liBlah.Text
        strFileName = Dir(strFile)
        If strFileName = liBlah.ListSubItems(frmMain.glFile) Then
            strPath = liBlah.ListSubItems(frmMain.glPath)
            lngFileLen = FileLen(strFile)
            dteFileDate = FileDateTime(strFile)
            
            Set mp3Inf = colMp3Info.Item(liBlah.Key)
            
            dblMaxAmp = mp3Inf.CurrMaxAmp
            If dblMaxAmp <> NOREALNUM Then
                dblMaxAmp = Round(dblMaxAmp, 3)
                varRadio = mp3Inf.RawRadiodBGain
                If varRadio = NOREALNUM Then
                    varRadio = "?"
                Else
                    varRadio = Round(varRadio, 3)
                End If
                
                varAlbum = mp3Inf.RawAlbumdBGain
                If varAlbum = NOREALNUM Then
                    varAlbum = "?"
                Else
                    varAlbum = Round(varAlbum, 3)
                End If
                
                Write #intFileNum, strPath, strFileName, dteFileDate, lngFileLen, dblMaxAmp, varRadio, varAlbum
            
            End If
            
            Set mp3Inf = Nothing
            
        End If
    Next
    
    Close #intFileNum
    
    Exit Sub

SaveGainAnalysis_Error:
    HandleError ("SaveGainAnalysis")
    On Error Resume Next
    Close #intFileNum
End Sub

Public Function GetDrivePartThing(strPath As String) As String
    Dim intDriveEnd As Integer
    
    If (Mid$(strPath, 2, 1) = ":") Then
        GetDrivePartThing = Left$(strPath, 2)
    Else
        If Left$(strPath, 2) = "\\" Then
            intDriveEnd = InStr(3, strPath, "\")
            If intDriveEnd > 0 Then intDriveEnd = InStr(intDriveEnd + 1, strPath, "\")
            If intDriveEnd > 0 Then
                GetDrivePartThing = Left$(strPath, intDriveEnd - 1)
            Else
                GetDrivePartThing = strPath
            End If
        Else
            GetDrivePartThing = strPath
        End If
    End If
End Function

Public Sub LoadGainAnalysis(strLoadPathFileName() As String, lstInfo As ListView, colMp3Info As Collection)
On Error GoTo LoadGainAnalysis_Error
    Dim intFileNum As Integer
    Dim liBlah As ListItem
    Dim lngFileLen As Long
    Dim dteFileDate
    Dim strFile As String
    Dim strPath As String
    Dim strFileName As String
    Dim varRadio As Variant
    Dim dblMaxAmp As Double
    Dim varAlbum As Variant
    Dim mp3Inf As Mp3Info
    Dim blnDateLenOK As Boolean
    Dim blnCheckFileDate As Boolean
    Dim blnNoAllDate As Boolean
    Dim blnCheckFileLen As Boolean
    Dim blnNoAllLen As Boolean
    Dim intResponse As Integer
    Dim i As Long
    Dim j As Long
    Dim strMatch As String
    Dim strRelativeBasePath As String
    Dim strRelativeBaseDrive As String
    
    frmMain.MousePointer = vbHourglass
    
    intFileNum = FreeFile
    
    blnCheckFileDate = True
    blnCheckFileLen = True
    blnNoAllLen = False
    blnNoAllDate = False
    
    strRelativeBasePath = strLoadPathFileName(LBound(strLoadPathFileName))
    strRelativeBaseDrive = GetDrivePartThing(strRelativeBasePath)
    
    For i = LBound(strLoadPathFileName) + 1 To UBound(strLoadPathFileName)
        
        On Error Resume Next
        Open strRelativeBasePath & "\" & strLoadPathFileName(i) For Input As #intFileNum
        If Err.Number > 0 Then
            GoTo SkipLoadLoop
        End If
        
        On Error Resume Next
        Input #intFileNum, strPath, strFileName, dteFileDate, lngFileLen, dblMaxAmp, varRadio, varAlbum
        While (Err.Number = 0)
            If (Mid$(strPath, 2, 1) <> ":") And (Left$(strPath, 2) <> "\\") Then
                If Left$(strPath, 1) = "\" Then
                    strPath = strRelativeBaseDrive & strPath
                Else
                    strPath = strRelativeBasePath & "\" & strPath
                End If
            End If
            If Right$(strPath, 1) <> "\" Then strPath = strPath & "\"
            strPath = CleanOutRelativePathInfo(strPath)
            On Error GoTo LoadGainAnalysis_Error
                strFile = strPath & strFileName
                On Error Resume Next
                    strFileName = Dir(strFile)
                    If Err.Number Then strFileName = ""
                On Error GoTo LoadGainAnalysis_Error
                If Len(strFileName) > 0 Then
                    strMatch = frmMain.AddSingleFile(strFile)
                    If Len(strMatch) = 0 Then
                        blnDateLenOK = False
                    Else
                        blnDateLenOK = True
                    End If
                    
                    If blnCheckFileDate Then
                        If dteFileDate <> FileDateTime(strFile) Then
                            If blnNoAllDate Then
                                blnDateLenOK = False
                            Else
                                intResponse = frmMain.YesNoAllFile( _
                                    GetLocalString("basSaveAnalysis.LCL_MODIFICATION_CAPTION", "Modification Warning"), _
                                    GetLocalString("basSaveAnalysis.LCL_MODIFICATION_WARNING", "Warning - File may have been modified after analysis was saved") & ":", _
                                    strFile, _
                                    GetLocalString("basSaveAnalysis.LCL_IGNORE_WARNING", "Load saved analysis results anyhow?"))
                                Select Case intResponse
                                    Case 0: 'Yes
                                        
                                    Case 1: 'Yes to All
                                        blnCheckFileDate = False
                                    Case 2: 'No
                                        blnDateLenOK = False
                                    Case 3: 'No to All
                                        blnDateLenOK = False
                                        blnNoAllDate = True
                                    Case 4: 'Cancel
                                        Close #intFileNum
                                        Exit Sub
                                End Select
                            End If
                        End If
                    End If
                    
                    If blnDateLenOK And blnCheckFileLen Then
                        If lngFileLen <> FileLen(strFile) Then
                            If blnNoAllLen Then
                                blnDateLenOK = False
                            Else
                                intResponse = frmMain.YesNoAllFile( _
                                    GetLocalString("basSaveAnalysis.LCL_SIZE_CAPTION", "Size Change Warning"), _
                                    GetLocalString("basSaveAnalysis.LCL_SIZE_WARNING", "Warning - File size changed after analysis was saved") & ":", _
                                    strFile, _
                                    GetLocalString("basSaveAnalysis.LCL_IGNORE_WARNING", "Load saved analysis results anyhow?"))
                                Select Case intResponse
                                    Case 0: 'Yes
                                        
                                    Case 1: 'Yes to All
                                        blnCheckFileLen = False
                                    Case 2: 'No
                                        blnDateLenOK = False
                                    Case 3: 'No to All
                                        blnDateLenOK = False
                                        blnNoAllLen = True
                                    Case 4: 'Cancel
                                        Close #intFileNum
                                        Exit Sub
                                End Select
                            End If
                        End If
                    End If
                    
                    If blnDateLenOK Then
                        Set mp3Inf = colMp3Info(lstInfo.ListItems(strMatch).Key)
                        mp3Inf.ResetVals
                        mp3Inf.CurrMaxAmp = dblMaxAmp
                        If IsNumeric(varRadio) Then mp3Inf.RadiodBGain = varRadio
                        If IsNumeric(varAlbum) Then mp3Inf.AlbumdBGain = varAlbum
                        frmMain.DispJunk lstInfo.ListItems(strMatch), mp3Inf
                        Set mp3Inf = Nothing
                    End If
                End If
            On Error Resume Next
            Input #intFileNum, strPath, strFileName, dteFileDate, lngFileLen, dblMaxAmp, varRadio, varAlbum
        Wend
        
        Close #intFileNum
SkipLoadLoop:
    Next i
    
    frmMain.MousePointer = vbDefault
    
    Exit Sub

LoadGainAnalysis_Error:
    HandleError ("LoadGainAnalysis")
    On Error Resume Next
    Close #intFileNum
    frmMain.MousePointer = vbDefault
    frmMain.doSortColumn
End Sub

Public Function CleanOutRelativePathInfo(strIn As String) As String
    Dim strOut As String
    Dim lngRelPos As Long
    Dim lngRelBackOne As Long
    
    strOut = strIn
    lngRelPos = InStr(strOut, "\.\")
    While lngRelPos > 0
        strOut = Left$(strOut, lngRelPos) & Mid$(strOut, lngRelPos + 3)
        lngRelPos = InStr(strOut, "\.\")
    Wend
    
    lngRelPos = InStr(strOut, "\..\")
    While lngRelPos > 0
        If lngRelPos = 1 Then
            lngRelBackOne = 1
        Else
            lngRelBackOne = InStrRev(strOut, "\", lngRelPos - 1)
        End If
        
        If lngRelBackOne = 0 Then lngRelBackOne = lngRelPos
        strOut = Left$(strOut, lngRelBackOne) & Mid$(strOut, lngRelPos + 4)
        lngRelPos = InStr(lngRelBackOne, strOut, "\..\")
    Wend
    CleanOutRelativePathInfo = strOut
End Function
