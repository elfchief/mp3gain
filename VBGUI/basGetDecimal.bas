Attribute VB_Name = "basGetDecimal"
'**************************************
'Windows API/Global Declarations for :GetInfo
'**************************************
'get format of currency with API call GetLocalInfo

Public Const LOCALE_USER_DEFAULT = &H400
Public Const LOCALE_SCURRENCY = &H14 ' local monetary symbol
Public Const LOCALE_SINTLSYMBOL = &H15 ' intl monetary symbol
Public Const LOCALE_SMONDECIMALSEP = &H16 ' monetary decimal separator
Public Const LOCALE_SMONTHOUSANDSEP = &H17 ' monetary thousand separator
Public Const LOCALE_SMONGROUPING = &H18 ' monetary grouping
Public Const LOCALE_ICURRDIGITS = &H19 ' # local monetary digits
Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
'**************************************
'Name: GetInfo
'Description:Get info in windows 95: Local curency symbol,International currency symbol, Decimaal separator,Thousand separator, Number of digits in group, Number of digits behind the decimal separator. http://137.56.41.168:2080/VisualBasicSource/vb4GetLocalInfo.txt
'By: Found on the World Wide Web
'
'
'Inputs: None
'
'Returns: None
'
'Assumes: None
'
'Side Effects: None
'**************************************

'
' Locale specific information
'
Public Function getDecimalSeparator() As String
Dim buffer As String * 100
Dim dl&
'compare this with
'Start/Settings/Control Panel/Regional Settings/Currency
#If Win32 Then
'    dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SCURRENCY, buffer, 99)
'    Form1.list1.AddItem " Local curency symbol: " & LPSTRToVBString(buffer)
'    dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SINTLSYMBOL, buffer, 99)
'    Form1.list1.AddItem " International currency symbol: " & LPSTRToVBString(buffer)
    dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SMONDECIMALSEP, buffer, 99)
    getDecimalSeparator = LPSTRToVBString(buffer)
'    dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SMONTHOUSANDSEP, buffer, 99)
'    Form1.list1.AddItem " Thousand separator: " & LPSTRToVBString(buffer)
'    dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SMONGROUPING, buffer, 99)
'    Form1.list1.AddItem " Number of digits in group: " & LPSTRToVBString(buffer)
'    dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_ICURRDIGITS, buffer, 99)
'    Form1.list1.AddItem " Number of digits behind the decimal separator: " & LPSTRToVBString(buffer)
#Else
    getDecimalSeparator = Mid(CStr(1.2), 2, 1)
#End If
End Function
'
' Extracts a VB string from a buffer containing a null terminated
' string
Public Function LPSTRToVBString$(ByVal s$)
    Dim nullpos&

    nullpos& = InStr(s$, Chr$(0))
    If nullpos > 0 Then
        LPSTRToVBString = Left$(s$, nullpos - 1)
    Else
        LPSTRToVBString = ""
    End If
End Function
        
