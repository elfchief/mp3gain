Attribute VB_Name = "basCommandOutput"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Name:         MGetCmdOutput
'' Filename:     MGetCmdOutput.bas
'' Author:       Mattias Sjögren (mattias@mvps.org)
''               http://www.msjogren.net/dotnet/
''
'' Description:  Generic module for launching a console app and
''               capture its output.
''
''
'' Copyright ©2000-2001, Mattias Sjögren
'' Extensively modified (probably made worse!) by Glen Sawyer
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

'''''''''''''''''''''
'''   Constants   '''
'''''''''''''''''''''

' STARTUPINFO flags
Public Const HIGH_PRIORITY_CLASS = &H80
Public Const IDLE_PRIORITY_CLASS = &H40
Public Const NORMAL_PRIORITY_CLASS = &H20
Public Const REALTIME_PRIORITY_CLASS = &H100

Public blnCancel As Boolean
Public blnAllowProcessCancel As Boolean
Public lngThreadPriority As Long
Public strAppPath As String
Public strAppDrive As String

#If GLENDEBUG Then
Public strDebugCheck As String
Public lngDebugCount As String
Public strDebugOutputCopy As String
Public strDebugCmdLineCopy As String
Public bytDebugOutput() As Byte
#End If

Private Const STARTF_USESHOWWINDOW = &H1
Private Const STARTF_USESTDHANDLES = &H100

' ShowWindow flags
Private Const SW_HIDE = 0
Public Const SW_SHOWNORMAL = 1

' DuplicateHandle flags
Private Const DUPLICATE_CLOSE_SOURCE = &H1
Private Const DUPLICATE_SAME_ACCESS = &H2

' Error codes
Private Const ERROR_BROKEN_PIPE = 109


'''''''''''''''''
'''   Types   '''
'''''''''''''''''

Private Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Long
End Type

Private Type STARTUPINFO
  cb As Long
  lpReserved As String
  lpDesktop As String
  lpTitle As String
  dwX As Long
  dwY As Long
  dwXSize As Long
  dwYSize As Long
  dwXCountChars As Long
  dwYCountChars As Long
  dwFillAttribute As Long
  dwFlags As Long
  wShowWindow As Integer
  cbReserved2 As Integer
  lpReserved2 As Long
  hStdInput As Long
  hStdOutput As Long
  hStdError As Long
End Type

Private Type PROCESS_INFORMATION
  hProcess As Long
  hThread As Long
  dwProcessId As Long
  dwThreadId As Long
End Type

Private Const BUFSIZE = 1024      ' pipe buffer size

''''''''''''''''''''
'''   Declares   '''
''''''''''''''''''''

Private Declare Function CreatePipe Lib "kernel32" ( _
  phReadPipe As Long, _
  phWritePipe As Long, _
  lpPipeAttributes As Any, _
  ByVal nSize As Long) As Long

Private Declare Function ReadFile Lib "kernel32" ( _
  ByVal hFile As Long, _
  lpBuffer As Any, _
  ByVal nNumberOfBytesToRead As Long, _
  lpNumberOfBytesRead As Long, _
  lpOverlapped As Any) As Long


Public Declare Function WriteFile _
   Lib "kernel32" _
   (ByVal hFile As Long, _
    lpBuffer As Any, _
    ByVal nNumberOfBytesToWrite As Long, _
    lpNumberOfBytesWritten As Long, _
    lpOverlapped As Any) As Long


Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" ( _
  ByVal lpApplicationName As String, _
  ByVal lpCommandLine As String, _
  lpProcessAttributes As Any, _
  lpThreadAttributes As Any, _
  ByVal bInheritHandles As Long, _
  ByVal dwCreationFlags As Long, _
  lpEnvironment As Any, _
  ByVal lpCurrentDriectory As String, _
  lpStartupInfo As STARTUPINFO, _
  lpProcessInformation As PROCESS_INFORMATION) As Long

Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, _
    ByVal uExitCode As Long) As Long

Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

Private Declare Function GetExitCodeProcess Lib "kernel32" _
    (ByVal hProcess As Long, lpExitCode As Long) As Long

Private Declare Function PeekNamedPipe Lib "kernel32" (ByVal hNamedPipe As Long, _
    lpBuffer As Any, ByVal nBufferSize As Long, lpBytesRead As Long, _
    lpTotalBytesAvail As Long, lpBytesLeftThisMessage As Long) As Long

Private Declare Function DuplicateHandle Lib "kernel32" ( _
  ByVal hSourceProcessHandle As Long, _
  ByVal hSourceHandle As Long, _
  ByVal hTargetProcessHandle As Long, _
  lpTargetHandle As Long, _
  ByVal dwDesiredAccess As Long, _
  ByVal bInheritHandle As Long, _
  ByVal dwOptions As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" ( _
  ByVal hObject As Long) As Long

Private Declare Function OemToCharBuff Lib "user32" Alias "OemToCharBuffA" ( _
  lpszSrc As Any, _
  ByVal lpszDst As String, _
  ByVal cchDstLength As Long) As Long
  
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, _
    ByVal dwMilliseconds As Long) As Long
    
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (Destination As Any, _
     Source As Any, _
     ByVal Length As Long)

Private Const WAIT_TIMEOUT = &H102&

#If GLENDEBUG Then
Public Function BSTRtoUniBytes(strIn As String, bytOut() As Byte) As Long
    Dim lngBytes As Long
    
    lngBytes = LenB(strIn)
    
    ReDim bytOut(0 To lngBytes - 1) As Byte
    
    CopyMemory bytOut(0), ByVal StrPtr(strIn), lngBytes
    
    BSTRtoUniBytes = lngBytes
End Function
#End If

Public Function BytesToBSTR(bytIn() As Byte, lngBytes As Long) As String
    Dim bytOut() As Byte
    
    ReDim bytOut(0 To lngBytes - 1) As Byte
    
    CopyMemory bytOut(0), bytIn(0), lngBytes
    
    BytesToBSTR = StrConv(bytOut(), vbUnicode)
End Function


''''''''''''''''''''''''''
'''   Public methods   '''
''''''''''''''''''''''''''

'
' Function GetCommandOutput
'
' sCommandLine:  [in] Command line to launch
' fStdOut        [in,opt] True (defualt) to capture output to STDOUT
' fStdErr        [in,opt] True to capture output to STDERR. False is default.
' fOEMConvert:   [in,opt] True (default) to convert DOS characters to Windows, False to skip conversion
'
' Returns:       String with STDOUT and/or STDERR output
'
Public Function GetCommandOutput(sOutput As String, sCommandLine As String, Optional sStartingDirectory As String, _
        Optional fStdOut As Boolean = True, _
        Optional fStdErr As Boolean = False, _
        Optional fOEMConvert As Boolean = True, _
        Optional MillisecondRefreshRate As Long = 1000, _
        Optional cStatusWatch As TextBox, _
        Optional cErrWatch As TextBox, _
        Optional blnAllowEvents As Boolean = True) As Long

    Dim hPipeRead As Long, hPipeWrite1 As Long, hPipeWrite2 As Long
    Dim hPipeReadSep As Long, hPipeWriteSep As Long
  
    Dim hCurProcess As Long
    Dim sa As SECURITY_ATTRIBUTES
    Dim si As STARTUPINFO
    Dim pi As PROCESS_INFORMATION
    Dim baOutput(0 To BUFSIZE) As Byte '0 is base, but we'll leave the extra byte on the end just for fun
    Dim sNewOutput As String
    Dim lBytesRead As Long
    Dim lTotalBytesAvail As Long
    Dim lBlah As Long
    Dim lngOuch As Long
    Dim blnSepErr As Boolean
    Dim fTwoHandles As Boolean
  
    Dim lRet As Long
#If GLENDEBUG Then
    Dim lngDebugBytesOut As Long
    
    strDebugOutputCopy = ""
    lngDebugCount = 0
    strDebugCheck = "Start: 0"
    strDebugCmdLineCopy = sCommandLine
    lngDebugBytesOut = 0
    ReDim bytDebugOutput(0 To 0) As Byte
#End If
    lRet = 0
  
  
    ' At least one of them should be True, otherwise there's no point in calling the function
    If (Not fStdOut) And (Not fStdErr) Then Err.Raise 5         ' Invalid Procedure call or Argument
  
    ' If both are true, we need two write handles. If not, one is enough.
    fTwoHandles = fStdOut And fStdErr
    
    With sa
        .nLength = Len(sa)
        .bInheritHandle = 1    ' get inheritable pipe handles
    End With
  
    If CreatePipe(hPipeRead, hPipeWrite1, sa, BUFSIZE) = 0 Then Exit Function
    hCurProcess = GetCurrentProcess()
  
    ' Replace our inheritable read handle with an non-inheritable. Not that it
    ' seems to be necessary in this case, but the docs say we should.
    Call DuplicateHandle(hCurProcess, hPipeRead, hCurProcess, hPipeRead, 0&, _
        0&, DUPLICATE_SAME_ACCESS Or DUPLICATE_CLOSE_SOURCE)
    
    ' If both STDOUT and STDERR should be redirected, get an extra handle.
    If fTwoHandles Then
        Call DuplicateHandle(hCurProcess, hPipeWrite1, hCurProcess, hPipeWrite2, 0&, _
            1&, DUPLICATE_SAME_ACCESS)
    End If
  
    blnSepErr = False
    If Not (cErrWatch Is Nothing) Then
        If (Not fStdErr) Then
            blnSepErr = True
      
            If CreatePipe(hPipeReadSep, hPipeWriteSep, sa, BUFSIZE) = 0 Then
                Call CloseHandle(hPipeRead)
                Call CloseHandle(hPipeWrite1)
                If hPipeWrite2 Then Call CloseHandle(hPipeWrite2)
          
                Exit Function
            End If
            Call DuplicateHandle(hCurProcess, hPipeReadSep, hCurProcess, hPipeReadSep, 0&, _
                0&, DUPLICATE_SAME_ACCESS Or DUPLICATE_CLOSE_SOURCE)
        End If
    End If
  
    With si
        .cb = Len(si)
        .dwFlags = STARTF_USESHOWWINDOW Or STARTF_USESTDHANDLES
        .wShowWindow = SW_HIDE 'SHOWNORMAL          ' hide the window
    
        If fTwoHandles Then
            .hStdOutput = hPipeWrite1
            .hStdError = hPipeWrite2
        ElseIf fStdOut Then
            .hStdOutput = hPipeWrite1
            If blnSepErr Then
                .hStdError = hPipeWriteSep
            End If
        Else
            .hStdError = hPipeWrite1
        End If
    End With

    lngOuch = CreateProcess(vbNullString, sCommandLine, ByVal 0&, ByVal 0&, 1, lngThreadPriority, _
        ByVal 0&, sStartingDirectory, si, pi)
    If lngOuch Then
    
        ' Close thread handle - we don't need it
        Call CloseHandle(pi.hThread)
    
        ' Also close our handle(s) to the write end of the pipe. This is important, since
        ' ReadFile will *not* return until all write handles are closed or the buffer is full.
        Call CloseHandle(hPipeWrite1)
        hPipeWrite1 = 0
        If hPipeWrite2 Then
            Call CloseHandle(hPipeWrite2)
            hPipeWrite2 = 0
        End If
        If hPipeWriteSep Then
            Call CloseHandle(hPipeWriteSep)
            hPipeWriteSep = 0
        End If

        Do
            If blnCancel And blnAllowProcessCancel Then
                TerminateProcess pi.hProcess, -1
            End If
            ' Add a DoEvents to allow more data to be written to the buffer for each call.
            ' This results in fewer, larger chunks to be read.
      
#If GLENDEBUG Then
            lngDebugCount = lngDebugCount + 1
            strDebugCheck = "Line 50: " & lngDebugCount
#End If
            If blnAllowEvents Then DoEvents
      
            ' Check the buffer to see if data is available
            ' We do this because PeekNamedPipe always returns immediately, while ReadFile
            ' sometimes blocks while waiting for input
            If PeekNamedPipe(hPipeRead, 0&, 0&, 0&, lTotalBytesAvail, 0&) <> 0& Then
                While lTotalBytesAvail > 0
                    ' Get the available data from the buffer
                    If lTotalBytesAvail > BUFSIZE Then
                        If ReadFile(hPipeRead, baOutput(0), BUFSIZE, lBytesRead, ByVal 0&) = 0& Then
                            lTotalBytesAvail = 0
                            lBytesRead = 0
                        End If
                    Else
                        If ReadFile(hPipeRead, baOutput(0), lTotalBytesAvail, lBytesRead, ByVal 0&) = 0& Then
                            lTotalBytesAvail = 0
                            lBytesRead = 0
                        End If
                    End If
                    If lBytesRead > 0 Then
                        If fOEMConvert Then
                            ' convert from "DOS" to "Windows" characters
                            sNewOutput = String$(lBytesRead, 0)
                            Call OemToCharBuff(baOutput(0), sNewOutput, lBytesRead)
                        Else
                            ' perform no conversion (except to Unicode)
                            'Doing this manually because Chinese computers screw it up
                            sNewOutput = BytesToBSTR(baOutput, lBytesRead)
                        End If
                        
#If GLENDEBUG Then
                        lngDebugBytesOut = lngDebugBytesOut + lBytesRead
                        ReDim Preserve bytDebugOutput(0 To lngDebugBytesOut - 1)
                        CopyMemory bytDebugOutput(lngDebugBytesOut - lBytesRead), baOutput(0), lBytesRead
                        strDebugOutputCopy = sOutput
#End If
                        sOutput = sOutput & sNewOutput
                        
                        If Not blnSepErr Then
                            If Not (cErrWatch Is Nothing) Then
                                cErrWatch.Text = cErrWatch.Text & sNewOutput
                            End If
                        End If
            
                        If Not (cStatusWatch Is Nothing) Then
                            cStatusWatch.Text = cStatusWatch.Text & sNewOutput
                        End If
                    End If
                    lTotalBytesAvail = lTotalBytesAvail - lBytesRead
#If GLENDEBUG Then
                    lngDebugCount = lngDebugCount + 1
                    strDebugCheck = "Line 75: " & lngDebugCount
#End If
                    If blnAllowEvents Then DoEvents
                Wend
            End If
      
            If blnSepErr Then
                If PeekNamedPipe(hPipeReadSep, 0&, 0&, 0&, lTotalBytesAvail, 0&) <> 0 Then
                    While lTotalBytesAvail > 0
                        ' Get the available data from the buffer
                        If lTotalBytesAvail > BUFSIZE Then
                            If ReadFile(hPipeReadSep, baOutput(0), BUFSIZE, lBytesRead, ByVal 0&) = 0& Then
                                lTotalBytesAvail = 0
                                lBytesRead = 0
                            End If
                        Else
                            If ReadFile(hPipeReadSep, baOutput(0), lTotalBytesAvail, lBytesRead, ByVal 0&) = 0& Then
                                lTotalBytesAvail = 0
                                lBytesRead = 0
                            End If
                        End If
                        If lBytesRead > 0 Then
                            If fOEMConvert Then
                                ' convert from "DOS" to "Windows" characters
                                sNewOutput = String$(lBytesRead, 0)
                                Call OemToCharBuff(baOutput(0), sNewOutput, lBytesRead)
                            Else
                                ' perform no conversion (except to Unicode)
                                sNewOutput = BytesToBSTR(baOutput, lBytesRead)
                            End If
                
                            cErrWatch.Text = cErrWatch.Text & sNewOutput
                        End If
                        lTotalBytesAvail = lTotalBytesAvail - lBytesRead
#If GLENDEBUG Then
                        lngDebugCount = lngDebugCount + 1
                        strDebugCheck = "Line 97: " & lngDebugCount
#End If
                        If blnAllowEvents Then DoEvents
                    Wend
                End If
            End If
      
            ' If you are executing an application that outputs data during a long time,
            ' and don't want to lock up your application, it might be a better idea to
            ' wrap this code in a class module in an ActiveX EXE and execute it asynchronously.
            ' Then you can raise an event here each time more data is available.
            'RaiseEvent OutputAvailabele(sNewOutput)
      
            ' Loop if the process is still running
        Loop While WaitForSingleObject(pi.hProcess, MillisecondRefreshRate) = WAIT_TIMEOUT
    
        ' Get any final data from the pipe buffer
        If PeekNamedPipe(hPipeRead, 0&, 0&, 0&, lTotalBytesAvail, 0&) <> 0 Then
            While lTotalBytesAvail > 0
                If lTotalBytesAvail > BUFSIZE Then
                    If ReadFile(hPipeRead, baOutput(0), BUFSIZE, lBytesRead, ByVal 0&) = 0& Then
                        lTotalBytesAvail = 0
                        lBytesRead = 0
                    End If
                Else
                    If ReadFile(hPipeRead, baOutput(0), lTotalBytesAvail, lBytesRead, ByVal 0&) = 0& Then
                        lTotalBytesAvail = 0
                        lBytesRead = 0
                    End If
                End If
            
                If lBytesRead > 0 Then
                    If fOEMConvert Then
                        ' convert from "DOS" to "Windows" characters
                        sNewOutput = String$(lBytesRead, 0)
                        Call OemToCharBuff(baOutput(0), sNewOutput, lBytesRead)
                    Else
                        ' perform no conversion (except to Unicode)
                        sNewOutput = BytesToBSTR(baOutput, lBytesRead)
                    End If
                
#If GLENDEBUG Then
                    lngDebugBytesOut = lngDebugBytesOut + lBytesRead
                    ReDim Preserve bytDebugOutput(0 To lngDebugBytesOut - 1)
                    CopyMemory bytDebugOutput(lngDebugBytesOut - lBytesRead), baOutput(0), lBytesRead
                    strDebugOutputCopy = sOutput
#End If
                    sOutput = sOutput & sNewOutput
                    
                    If Not blnSepErr Then
                        If Not (cErrWatch Is Nothing) Then
                            cErrWatch.Text = cErrWatch.TabIndex & sNewOutput
                        End If
                    End If
                
                    If Not (cStatusWatch Is Nothing) Then
                        cStatusWatch.Text = cStatusWatch.Text & sNewOutput
                    End If
                End If
                lTotalBytesAvail = lTotalBytesAvail - lBytesRead
#If GLENDEBUG Then
                lngDebugCount = lngDebugCount + 1
                strDebugCheck = "Line 123: " & lngDebugCount
#End If
                If blnAllowEvents Then DoEvents
            Wend
#If GLENDEBUG Then
            lngDebugCount = lngDebugCount + 1
            strDebugCheck = "Line 127: " & lngDebugCount
#End If
        End If
    
        If blnSepErr Then
            If PeekNamedPipe(hPipeReadSep, 0&, 0&, 0&, lTotalBytesAvail, 0&) <> 0 Then
                While lTotalBytesAvail > 0
                    If lTotalBytesAvail > BUFSIZE Then
                        If ReadFile(hPipeReadSep, baOutput(0), BUFSIZE, lBytesRead, ByVal 0&) = 0& Then
                            lTotalBytesAvail = 0
                            lBytesRead = 0
                        End If
                    Else
                        If ReadFile(hPipeReadSep, baOutput(0), lTotalBytesAvail, lBytesRead, ByVal 0&) = 0& Then
                            lTotalBytesAvail = 0
                            lBytesRead = 0
                        End If
                    End If
                
                    If lBytesRead Then
                        If fOEMConvert Then
                            ' convert from "DOS" to "Windows" characters
                            sNewOutput = String$(lBytesRead, 0)
                            Call OemToCharBuff(baOutput(0), sNewOutput, lBytesRead)
                        Else
                            ' perform no conversion (except to Unicode)
                            sNewOutput = BytesToBSTR(baOutput, lBytesRead)
                        End If
                    
                        cErrWatch.Text = cErrWatch.Text & sNewOutput
                    End If
                    lTotalBytesAvail = lTotalBytesAvail - lBytesRead
#If GLENDEBUG Then
                    lngDebugCount = lngDebugCount + 1
                    strDebugCheck = "Line 148: " & lngDebugCount
#End If
                    If blnAllowEvents Then DoEvents
                Wend
#If GLENDEBUG Then
                lngDebugCount = lngDebugCount + 1
                strDebugCheck = "Line 152: " & lngDebugCount
#End If
            End If
        End If
        ' When the process terminates successfully, Err.LastDllError will be
        ' ERROR_BROKEN_PIPE (109). Other values indicates an error.
        Call GetExitCodeProcess(pi.hProcess, lRet&)
        Call CloseHandle(pi.hProcess)
    
    End If
  
    ' clean up
    Call CloseHandle(hPipeRead)
    If hPipeReadSep Then Call CloseHandle(hPipeReadSep)
    If hPipeWrite1 Then Call CloseHandle(hPipeWrite1)
    If hPipeWrite2 Then Call CloseHandle(hPipeWrite2)
    If hPipeWriteSep Then Call CloseHandle(hPipeWriteSep)
    GetCommandOutput = lRet
End Function

