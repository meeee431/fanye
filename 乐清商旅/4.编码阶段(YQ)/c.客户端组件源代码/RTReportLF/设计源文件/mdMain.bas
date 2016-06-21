Attribute VB_Name = "mdMain"

Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const HWND_TOPMOST = -1
Private Const HWND_TOP = 0
Private Const HWND_NOTOPMOST = -2

Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'���ö���


'�õ�һ�������ָ��ά�ĳ��ȣ����д��򷵻�0
Public Function ArrayLength(paszIn As Variant, Optional pnIndex As Integer = 1) As Long
    Dim lLow As Long, lHigh As Long
    On Error GoTo Here
    
    lLow = LBound(paszIn, pnIndex)
    lHigh = UBound(paszIn, pnIndex)
    ArrayLength = lHigh - lLow + 1
    Exit Function
Here:
    ArrayLength = 0
End Function

Public Function EncodeString(ByVal pszString As String) As String
'���ַ�����[]����
    EncodeString = "[" & pszString & "]"
End Function
Public Function UnEncodeString(ByVal pszString As String) As String
'�õ�[]�е��ַ���
    pszString = LeftAndRight(pszString, False, "[")
    UnEncodeString = LeftAndRight(pszString, True, "]")
End Function
Public Function EncodeKeyValue(ByVal pszKey As String, ByVal pszValue As String, Optional ByVal pbHaveRange As Boolean = False) As String
'�γ�[key=value]��ʽ
    EncodeKeyValue = pszKey & "=" & pszValue
    If pbHaveRange Then
        EncodeKeyValue = EncodeString(EncodeKeyValue)
    End If
End Function
Public Sub UnEncodeKeyValue(ByVal pszString As String, ByRef szReturnKey As String, ByRef szReturnValue As String)
'����KEY��VALUE
    'pszString = UnEncodeString(pszString)
    szReturnKey = ""
    szReturnValue = ""
    If InStr(pszString, "[") <> 0 And InStr(pszString, "]") <> 0 Then
        pszString = UnEncodeString(pszString)
        szReturnKey = LeftAndRight(pszString, True, "=")
        szReturnValue = LeftAndRight(pszString, False, "=")
    End If
    
End Sub
Public Function SplitEncodeStringArray(ByVal pszString As String) As String()
'��[string1][string2]...[stringn]��ʽ���ַ����Ԫ�����鷵��
    Dim atReturn() As String
    Dim nItemCount As Integer
    Dim szitem As String
    
    Do
        szitem = LeftAndRight(pszString, True, "]")
        pszString = LeftAndRight(pszString, False, "]")
        szitem = LeftAndRight(szitem, False, "[")
'        If szitem = "" Then Exit Do
        nItemCount = nItemCount + 1
        ReDim Preserve atReturn(1 To nItemCount)
        atReturn(nItemCount) = EncodeString(szitem)
    Loop Until pszString = ""
    SplitEncodeStringArray = atReturn
End Function



'�õ���Ӣ�Ļ���ַ������ַ�����
Public Function LenA(ByVal pszString As String) As Long
    LenA = LenB(StrConv(pszString, vbFromUnicode))
End Function

'����unicode���ַ�����ȡ����
Public Function MidA(ByVal pszString, plStart, plLen) As String
    Dim abyReturn() As Byte
    abyReturn = StrConv(pszString, vbFromUnicode)
    Dim aReturn() As Byte
    ReDim aReturn(plLen - 1)
    Dim i As Long
    For i = 0 To plLen - 1
        aReturn(i) = abyReturn(i + plStart - 1)
    Next i
    MidA = StrConv(aReturn, vbUnicode)
End Function

'ȡ�ַ���
Public Function LeftAndRight(InString As String, IsLeft As Boolean, FCHAR As String) As String
    Dim n As Integer
    n = InStr(1, InString, FCHAR)
    If n = 0 Then
        If IsLeft Then
            LeftAndRight = InString
        Else
            LeftAndRight = ""
        End If
    Else
        If IsLeft Then
            LeftAndRight = Left(InString, InStr(1, InString, FCHAR) - 1)
        Else
            LeftAndRight = Right(InString, Len(InString) - InStr(1, InString, FCHAR))
        End If
    End If
End Function

Public Function ShowProgess(ByVal lCurrProgess As Long, ByVal lTotalProgess As Long, Optional ByVal ifFirst As Boolean, Optional pszCaption As String)
'    If ifFirst Then
'        frmProgess.ProgessCaption = pszCaption
'        frmProgess.Show
'        SetWindowPos frmProgess.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
'        DoEvents
'    End If
'    frmProgess.ProgessBar.Value = Int(lCurrProgess / lTotalProgess * 100)
'    frmProgess.ProgessBar.Refresh
'
''    MsgBox "k"
''    frmProgess.Refresh
''    DoEvents
''    SetWindowPos frmProgess.hwnd, WND_TOPMOST, 0, 0, 0, 0, HWND_NOTOPMOST
End Function

Public Function CloseProgess()
'    Unload frmProgess
End Function


