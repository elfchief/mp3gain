VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Mp3Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarRadiodBGain As Double 'local copy
Private mvarAlbumdBGain As Double 'local copy
Private mvarModdB As Double
Private mvarCurrMaxAmp As Double 'local copy
Private mvarCurrMaxGain As Integer
Private mvarCurrMinGain As Integer

Private Const FIVELOG10TWO = 1.50514997831991

Public Property Get RawAlbumdBGain() As Double
    RawAlbumdBGain = mvarAlbumdBGain
End Property

Public Property Get RawRadiodBGain() As Double
    RawRadiodBGain = mvarRadiodBGain
End Property
    
    Private Static Function Log10(X)
1       Log10 = Log(X) / 2.30258509299405 'Log(10#)
    End Function

        Private Static Function Log2(X)
2           Log2 = Log(X) / 0.693147180559945 'Log(2#)
    End Function

            Public Property Let CurrMaxAmp(ByVal vData As Double)
                'used when assigning a value to the property, on the left side of an assignment.
                'Syntax: X.CurrMaxAmp = 5
3               mvarCurrMaxAmp = vData
    End Property

            Public Property Let CurrMaxGain(ByVal vData As Integer)
                'used when assigning a value to the property, on the left side of an assignment.
                'Syntax: X.CurrMaxAmp = 5
3               mvarCurrMaxGain = vData
    End Property

            Public Property Let CurrMinGain(ByVal vData As Integer)
                'used when assigning a value to the property, on the left side of an assignment.
                'Syntax: X.CurrMaxAmp = 5
3               mvarCurrMinGain = vData
    End Property


    Public Property Get CurrMaxAmp() As Double
        'used when retrieving value of a property, on the right side of an assignment.
        'Syntax: Debug.Print X.CurrMaxAmp
1       CurrMaxAmp = mvarCurrMaxAmp
    End Property


    Public Property Get CurrMaxGain() As Integer
        'used when retrieving value of a property, on the right side of an assignment.
        'Syntax: Debug.Print X.CurrMaxAmp
1       CurrMaxGain = mvarCurrMaxGain
    End Property


    Public Property Get CurrMinGain() As Integer
        'used when retrieving value of a property, on the right side of an assignment.
        'Syntax: Debug.Print X.CurrMaxAmp
1       CurrMinGain = mvarCurrMinGain
    End Property



    Public Property Let AlbumdBGain(ByVal vData As Double)
        'used when assigning a value to the property, on the left side of an assignment.
        'Syntax: X.AlbumdBGain = 5
1       mvarAlbumdBGain = vData
    End Property


    Public Property Get AlbumdBGain() As Double
        'used when retrieving value of a property, on the right side of an assignment.
        'Syntax: Debug.Print X.AlbumdBGain
1       If mvarAlbumdBGain <> NOREALNUM Then
2           AlbumdBGain = mvarAlbumdBGain + mvarModdB
3       Else
4           AlbumdBGain = mvarAlbumdBGain
        End If
    End Property


    Public Property Get AlbumMp3Gain() As Integer
        'used when retrieving value of a property, on the right side of an assignment.
        'Syntax: Debug.Print X.AlbumdBGain
1       If mvarAlbumdBGain <> NOREALNUM Then
2           AlbumMp3Gain = CStr(CInt(Round((mvarAlbumdBGain + mvarModdB) / (FIVELOG10TWO))))
3       Else
4           AlbumMp3Gain = 0
        End If
    End Property



    Public Property Let RadiodBGain(ByVal vData As Double)
        'used when assigning a value to the property, on the left side of an assignment.
        'Syntax: X.RadiodBGain = 5
1       mvarRadiodBGain = vData
    End Property


    Public Property Get RadiodBGain() As Double
        'used when retrieving value of a property, on the right side of an assignment.
        'Syntax: Debug.Print X.RadiodBGain
1       If mvarRadiodBGain <> NOREALNUM Then
2           RadiodBGain = mvarRadiodBGain + mvarModdB
3       Else
4           RadiodBGain = mvarRadiodBGain
        End If
    End Property

    Public Property Get RadioMp3Gain() As Integer
        On Error GoTo GetRadioMp3Gain_Error
        'used when retrieving value of a property, on the right side of an assignment.
        'Syntax: Debug.Print X.AlbumdBGain
1       If mvarRadiodBGain <> NOREALNUM Then
2           RadioMp3Gain = CInt(Round((mvarRadiodBGain + mvarModdB) / (FIVELOG10TWO)))
3       Else
4           RadioMp3Gain = 0
        End If
        Exit Property
GetRadioMp3Gain_Error:
5       HandleError "Get_RadioMp3Gain"
    End Property

    Public Property Get MaxNoclipMp3Gain() As Long
        On Error GoTo GetMaxNoclipMp3Gain_Error
        Dim dblAdjust As Double
    
1       If (mvarCurrMaxAmp <> NOREALNUM) And _
            (mvarCurrMaxAmp < 1000000#) And _
            (mvarCurrMaxAmp > 0) Then
2           dblAdjust = 4# * Log2(32767# / mvarCurrMaxAmp)
3           If CDbl(Int(dblAdjust)) > dblAdjust Then
4               MaxNoclipMp3Gain = Int(dblAdjust) - 1
5           Else
6               MaxNoclipMp3Gain = Int(dblAdjust)
            End If
7       Else
8           MaxNoclipMp3Gain = 0
        End If
        Exit Property
GetMaxNoclipMp3Gain_Error:
9       HandleError "Get_MaxNoclipMp3Gain"
    End Property

    Public Sub AlterDb(ByVal vData As Double)
        On Error GoTo AlterDb_Error
        Dim intGainChange As Integer
        
1       If (mvarRadiodBGain <> NOREALNUM) Then
2           mvarRadiodBGain = mvarRadiodBGain + vData
        End If
3       If (mvarAlbumdBGain <> NOREALNUM) Then
4           mvarAlbumdBGain = mvarAlbumdBGain + vData
        End If
        intGainChange = CInt(-vData / (FIVELOG10TWO))
5       If (mvarCurrMaxAmp <> NOREALNUM) Then
6           mvarCurrMaxAmp = mvarCurrMaxAmp * (2# ^ (intGainChange / 4#))
        End If
        If (mvarCurrMaxGain <> -1) Then
            mvarCurrMaxGain = mvarCurrMaxGain + intGainChange
        End If
        If (mvarCurrMinGain <> -1) Then
            mvarCurrMinGain = mvarCurrMinGain + intGainChange
        End If
        Exit Sub
AlterDb_Error:
7       HandleError "AlterDb"
    End Sub

    Public Property Let ModifydBGain(ByVal vData As Double)
1       mvarModdB = vData
    End Property

    Public Property Get ModifydBGain() As Double
1       ModifydBGain = mvarModdB
    End Property

    Public Sub ResetVals()
        On Error GoTo ResetVals_Error
1       mvarRadiodBGain = NOREALNUM
2       mvarAlbumdBGain = NOREALNUM
3       mvarCurrMaxAmp = NOREALNUM
4       mvarCurrMaxGain = -1
5       mvarCurrMinGain = -1
        Exit Sub
ResetVals_Error:
6       HandleError "ResetVals"
    End Sub

    Private Sub Class_Initialize()
        On Error GoTo Class_Initialize_Error
1       ResetVals
2       mvarModdB = 0
        Exit Sub
Class_Initialize_Error:
3       HandleError "Class_Initialize"
    End Sub
