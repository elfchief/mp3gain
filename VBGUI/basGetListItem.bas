Attribute VB_Name = "basGetListItem"
Option Explicit

'Structures

Public Type LV_ITEM
    mask As Long
    iItem As Long
    iSubItem As Long
    State As Long
    stateMask As Long
    pszText As Long
    cchTextMax As Long
    iImage As Long
    lParam As Long
    iIndent As Long
End Type

'Constants
Public Const LVIF_TEXT = &H1

Public Const LVM_FIRST = &H1000
Public Const LVM_GETITEMTEXT = LVM_FIRST + 45

Public Const LVM_GETCOLUMNORDERARRAY = LVM_FIRST + 59
Public Const LVM_SETCOLUMNORDERARRAY = LVM_FIRST + 58
'API declarations

Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hWnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long

Public Sub ListView_GetListSubItem(lngIndex As Long, _
                                      hWnd As Long, _
                                      strName As String, _
                                      lngSubItemNum As Long)
    Dim objItem As LV_ITEM
    Dim baBuffer(2000) As Byte
    Dim lngLength As Long
    
    '
    ' Obtain the name of the specified list view item
    '
    objItem.mask = LVIF_TEXT
    objItem.iSubItem = lngSubItemNum
    objItem.pszText = VarPtr(baBuffer(0))
    objItem.cchTextMax = UBound(baBuffer)
    lngLength = SendMessage(hWnd, LVM_GETITEMTEXT, lngIndex, _
                            VarPtr(objItem))
    strName = Left$(StrConv(baBuffer, vbUnicode), lngLength)


End Sub


