Attribute VB_Name = "basHandleError"
Option Explicit

Public Sub HandleError(errSrc As String)
    Dim strErrLog As String
    Dim strErrMess As String
    Dim errNum As Long
    Dim errDesc As String
    Dim errSource As String
    Dim errLine As Long
    Dim intFileNum As Integer
    
    errNum = Err.Number
    errDesc = Err.Description
    errSource = Err.Source
    errLine = Erl
    
    On Error Resume Next
    
    strErrMess = "ERROR " & errNum & " in " & errSrc
    If (errSource <> "") Then
        strErrMess = strErrMess & " (" & errSource & ")"
    End If
    If (errLine <> 0) Then
        strErrMess = strErrMess & " line " & errLine
    End If
    strErrMess = strErrMess & " : " & errDesc
    
    MsgBox strErrMess
    
    strErrLog = strAppPath & App.EXEName & ".log"
    
    intFileNum = FreeFile
    Open strErrLog For Append As #intFileNum
    Print #intFileNum, Now & vbTab & strErrMess
    Close #intFileNum
        

End Sub
