Attribute VB_Name = "basHandleError"
'
'  HandleError.bas - Simple error log writer for MP3GainGUI
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
