Attribute VB_Name = "mdCommonLib"
'=========================================================================================
'Author:
'Detail:�м�㹲��ģ��
'=========================================================================================


Option Explicit
'------------------------------------------------------------------------------------------
'���³�������
'---------------------------------------------------------
Public Const cszRegKeyCompany = "Software\RTsoft"       '��˾ע���
Public Const cszRegKeyProduct = "RTStation"             '��Ʒע���
Public Const cszUnrepeatString = "@#111#&afsdjf*@#&!("  '�����ظ����ַ��������ڱ��Ψһ���ַ���

'-----------------------------------------------------------------------------------------
Public Const cszEmptyDateStr = "1900-01-01"
Public Const cszForeverDateStr = "2050-01-01"
Public Const cszEmptyTimeStr = "00:00:00"

Public Const cszDateStr = "YYYY-MM-DD"
Public Const cszTimeStr = "HH:mm:SS"
Public Const cszDateTimeStr = cszDateStr & " " & cszTimeStr

Public Const cdtEmptyDate = #1/1/1900#
Public Const cdtEmptyTime = #12:00:00 AM#
Public Const cdtEmptyDateTime = #1/1/1900#


Const szTen = "ʰ"
Const szTmillion = "��"
Const szMillion = "��"
Const szThousand = "Ǫ"
Const szHundred = "��"






Public Function GetLString(InString As String) As String
    GetLString = Trim(LeftAndRight(InString, True, "["))
End Function
'��������ά��
Public Function GetArrayDimension(Arr As Variant) As Integer
    Dim i As Integer
    Dim temp As Integer
    On Error GoTo err
    For i = 1 To 100
        temp = UBound(Arr, i)
    Next i
err:
    GetArrayDimension = i - 1
End Function

'ȡ�ַ���
Public Function LeftAndRight(InString As String, IsLeft As Boolean, FCHAR As String) As String
'InString:Դ�ַ���
'IsLeft:�Ƿ���
'FCHAR:�ָ��ַ�
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
            LeftAndRight = Right(InString, Len(InString) - InStr(1, InString, FCHAR) - Len(FCHAR) + 1)
        End If
    End If
End Function


''��ʽ���ַ������ж��Ƿ��зǷ��ַ���ȱʡ�ķǷ��ַ���',Ҳ���Լ�ָ������Ȼ���ַ����ĺ󵼿ո�ȥ��
'Public Function FormatStr(pszInStr As String, Optional pszInValidChars As String = "';") As String
'    Dim nStrLen As Integer
'    Dim i As Integer
'    nStrLen = Len(pszInValidChars)
'    If nStrLen > 0 Then
'        For i = 1 To nStrLen
'            If InStr(1, pszInStr, Mid(pszInValidChars, i, 1), vbTextCompare) > 0 Then RaiseError 501  'ERR_StrIllegal
'        Next i
'    End If
'    FormatStr = Trim(pszInStr)
'End Function

'�õ�һ�������ָ��ά�ĳ��ȣ����д��򷵻�0
Public Function ArrayLength(paszIn As Variant, Optional pnIndex As Integer = 1) As Long
    Dim lLow As Long, lHigh As Long
    On Error GoTo ErrorHandle
    
    lLow = LBound(paszIn, pnIndex)
    lHigh = UBound(paszIn, pnIndex)
    ArrayLength = lHigh - lLow + 1
    Exit Function
ErrorHandle:
    ArrayLength = 0
End Function


'�ж�ָ����VB�����Ƿ��ǿյ�
Public Function VBDateIsEmpty(pdtIn As Date) As Boolean
    Dim dtTemp As Date
    VBDateIsEmpty = IIf(Format(pdtIn, cszDateStr) = Format(dtTemp, cszDateStr), True, False)
End Function

'�ж�ָ����VBʱ���Ƿ��ǿյ�
Public Function VBTimeIsEmpty(pdtIn As Date) As Boolean
    Dim dtTemp As Date
    VBTimeIsEmpty = IIf(Format(pdtIn, cszTimeStr) = Format(dtTemp, cszTimeStr), True, False)
End Function

'�ж�ָ����VB����ʱ���Ƿ��ǿյ�
Public Function VBDateTimeIsEmpty(pdtIn As Date) As Boolean
    Dim dtTemp As Date
    VBDateTimeIsEmpty = IIf(Format(pdtIn, cszDateTimeStr) = Format(dtTemp, cszDateTimeStr), True, False)
End Function


'����һ�����������ں͵ڶ�������ʱ��ϲ�Ϊ����ʱ��
Public Function SelfGetFullDateTime(ByVal dtDate As Date, ByVal dtTime As Date) As Date
    SelfGetFullDateTime = CDate(Format(dtDate, cszDateStr) & " " & Format(dtTime, cszTimeStr))
End Function
 
Public Function GetUnEncodeKey(ByVal pszOrgCode As String) As String
    GetUnEncodeKey = Mid(pszOrgCode, 2)
End Function

'�����ַ������μ�MakeDisplayString
Public Function ResolveDisplayEx(ByVal pszDisplayString As String, Optional ByRef pszName As String = "*****") As String

    Dim i  As Integer, nStrLen As Integer
    i = InStr(1, pszDisplayString, "[")
    nStrLen = Len(pszDisplayString)
    If i > 0 Then
        pszName = Left(pszDisplayString, i - 1)
        If pszDisplayString <> "*****" Then
            ResolveDisplayEx = Mid(pszDisplayString, i + 1, nStrLen - i - 1)
        End If
    Else
        ResolveDisplayEx = pszDisplayString
    End If
End Function


'�����ַ������μ�MakeDisplayString
Public Function ResolveDisplay(ByVal pszDisplayString As String, Optional ByRef pszName As String = "*****") As String
    Dim i  As Integer, nStrLen As Integer
    i = InStr(1, pszDisplayString, "[")
    nStrLen = Len(pszDisplayString)
    If i > 0 Then
        ResolveDisplay = Left(pszDisplayString, i - 1)
        If pszDisplayString <> "*****" Then
            pszName = Mid(pszDisplayString, i + 1, nStrLen - i - 1)
        End If
    Else
        ResolveDisplay = pszDisplayString
        pszName = ""
    End If
End Function

Public Function NumAdd(ByVal pszSource As String, ByVal lNum As Long) As String
    Dim i As Integer
    Dim nLength As Integer
    Dim szNumPart As String
    nLength = Len(pszSource)
    For i = nLength To 1 Step -1
        If Mid(pszSource, i, 1) < "0" Or Mid(pszSource, i, 1) > "9" Then
            Exit For
        End If
    Next i
    szNumPart = Right(pszSource, nLength - i)
    Dim nNumPartLen As Integer
    nNumPartLen = Len(szNumPart)
    If nNumPartLen > 0 Then
        szNumPart = Format(Val(szNumPart) + lNum, String(Len(szNumPart), "0"))
        szNumPart = Right(szNumPart, nNumPartLen)
    End If
    NumAdd = Left(pszSource, i) & szNumPart
End Function
Public Function NumSub(ByVal pszSource As String, ByVal lNum As Long)
    Dim i As Integer
    Dim nLength As Integer
    Dim szNumPart As String
    Dim lNumPart As Long, nNumPartLen As Integer
    nLength = Len(pszSource)
    For i = nLength To 1 Step -1
        If Mid(pszSource, i, 1) < "0" Or Mid(pszSource, i, 1) > "9" Then
            Exit For
        End If
    Next i
    szNumPart = Right(pszSource, nLength - i)
    nNumPartLen = Len(szNumPart)
    If nNumPartLen > 0 Then
        lNumPart = Val(szNumPart)
        If lNumPart - lNum < 0 Then
            lNumPart = 10 ^ nNumPartLen + lNumPart - lNum
        Else
            lNumPart = lNumPart - lNum
        End If
        szNumPart = Format(lNumPart, String(nNumPartLen, "0"))
    End If
    NumSub = Left(pszSource, i) & szNumPart
End Function

Public Function GetTextToNumeric(pszText As String, Optional pbCanBeNegative As Boolean = True, Optional pbCanBeDecimal As Boolean = True) As String
    '���ַ���תΪ��Ч����
On Error GoTo ErrHandle
    Dim dlValue As Double
    pszText = Trim(pszText)
    dlValue = Val(pszText)
    Dim lTmp As Long
    Dim i As Integer
    
    If dlValue <> 0 Or pszText = "0" Or pszText = "0." Then
        
        GetTextToNumeric = Trim(Str(dlValue))
'        If Abs(dlValue) < 1 And Abs(dlValue) > 0 Then GetTextToNumeric = "0" & GetTextToNumeric
        If Right(pszText, 1) = "." And pbCanBeDecimal And dlValue = Int(dlValue) Then
            GetTextToNumeric = GetTextToNumeric & "."
        Else
            '��β����0
            If dlValue = Int(dlValue) And Len(Trim(dlValue)) <> Len(pszText) Then
                lTmp = InStr(1, pszText, ".", vbBinaryCompare)
                If lTmp > 0 Then
                    For i = lTmp + 1 To Len(pszText)
                        If Mid(pszText, i, 1) <> "0" Then
                            Exit For
                        End If
                    Next i
                    If i - lTmp - 1 > 0 Then
                        GetTextToNumeric = GetTextToNumeric & "." & String(i - lTmp - 1, "0")
                    End If
                End If
            
            End If
        End If
        Exit Function
    End If
    GetTextToNumeric = ""
    Select Case pszText
        Case "-"
            If pbCanBeNegative Then
                GetTextToNumeric = "-"
            End If
        Case "."
            If pbCanBeDecimal Then
                GetTextToNumeric = "."
            End If
    End Select
    Exit Function
ErrHandle:
    GetTextToNumeric = "0"
End Function


Public Function FormatMoney(pvStr As Variant) As String
FormatMoney = Format(pvStr, "0.00")
End Function

Public Function FormatSeparateMoney(pvStr As Variant) As String
    '�����ö��ŷָ��Ľ��
    FormatSeparateMoney = Format(pvStr, "##,##0.00")
End Function



Public Function GetUnicodeBySize(pszString As String, pnSize As Integer) As String
    Dim szTmp As String
    szTmp = StrConv(pszString, vbFromUnicode)
    If LenB(szTmp) > pnSize Then
        GetUnicodeBySize = StrConv(LeftB(szTmp, pnSize), vbUnicode)
    Else
        GetUnicodeBySize = pszString
    End If
End Function



  Public Function IfNumber(nAsc As Integer) As Integer
  If nAsc >= 48 And nAsc <= 57 Or nAsc = 13 Or nAsc = 8 Then
     IfNumber = nAsc
  Else
    IfNumber = 0
     
  End If
  End Function
Public Function LenString(pszString As String) As Integer
    '�õ��ַ�����ʵ�ʳ���(һ�������ַ�������λ)
    Dim szTmp As String
    szTmp = StrConv(pszString, vbFromUnicode)
    LenString = LenB(szTmp)
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
    If InStr(pszString, "[") <> 0 And InStr(pszString, "]") <> 0 Then
        pszString = UnEncodeString(pszString)
    End If
    szReturnKey = LeftAndRight(pszString, True, "=")
    szReturnValue = LeftAndRight(pszString, False, "=")
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
Public Function MidA(pszString As String, ByVal plStart As Long, Optional ByVal plLen As Long) As String
On Error GoTo ErrHandle
    If plStart + plLen > LenA(pszString) Or plLen = 0 Then
        plLen = LenA(pszString) - plStart + 1
    End If
    Dim abyReturn() As Byte
    abyReturn = StrConv(pszString, vbFromUnicode)
    Dim aReturn() As Byte
    ReDim aReturn(plLen - 1)
    Dim i As Long
    For i = 0 To plLen - 1
        aReturn(i) = abyReturn(i + plStart - 1)
    Next i
    MidA = StrConv(aReturn, vbUnicode)
    Exit Function
ErrHandle:
    MidA = ""
End Function


' *******************************************************************
' *   Brief Description: ���ܿ���                                   *
' *   Engineer: ½����                                              *
' *   Date Generated: 2002/06/21                                    *
' *   Last Revision Date:                                           *
' *******************************************************************
Public Function EncryptPassword(ByVal pszPassword As String) As String
'pszPassword ����
'ѡ��һ�ּ����㷨�Կ�����м���
    
    Dim nLen As Integer
    Dim nPwdLen As Integer
    Dim i As Integer
    Dim szResult As String
    Dim nIndex As Integer

    nPwdLen = Len(pszPassword)
    If nPwdLen = 0 Then Exit Function
    If nPwdLen > 99 Then
        nPwdLen = 99
        pszPassword = Left(pszPassword, nPwdLen)
    End If
    
    Dim szTmp As String
    Dim szTmp2 As String
    szResult = ""
    
    For i = 1 To Len(pszPassword)
        szTmp = Hex(Asc(Mid(pszPassword, i, 1)))
        If Len(szTmp) = 1 Then szTmp = "0" & szTmp
        szResult = szResult & szTmp
    Next i
    Dim nTmp As Integer
    szResult = XOREncrypt(szResult)
    nLen = Len(szResult)
    nTmp = nLen / 3
    szResult = Right(szResult, nLen - nTmp) & Left(szResult, nTmp) '���һ���
    szResult = XOREncrypt(szResult)
    szResult = Right(szResult, nLen - nTmp) & Left(szResult, nTmp) '���һ���
    
    szResult = Right(Format(nPwdLen, "00"), 1) & szResult & Left(Format(nPwdLen, "00"), 1)
    EncryptPassword = szResult
End Function
Private Function XOREncrypt(ByVal pszSource As String) As String
    Const cnPerNum = 3
    Const cnXorValue = 987
    Dim szTmp As String
    Do
        Dim nTmp As Integer
        szTmp = Left(pszSource, cnPerNum)
        If Len(szTmp) < cnPerNum Then   'С�����Χ�ں���
'            szTmp = Hex(HexToDec(szTmp) Xor (cnXorValue And 10 ^ (Len(szTmp)) - 1))
        Else
            szTmp = Hex(HexToDec(szTmp) Xor cnXorValue)
            If Len(szTmp) < cnPerNum Then szTmp = String(cnPerNum - Len(szTmp), "0") & szTmp   '��0
        End If
        XOREncrypt = XOREncrypt & szTmp
        If Len(pszSource) < cnPerNum Then
            Exit Do
        Else
            pszSource = Right(pszSource, Len(pszSource) - cnPerNum)
        End If
    Loop
End Function
Private Function HexToDec(ByVal pszHex As String) As Long
    Dim i As Integer
    For i = 1 To Len(pszHex)
        HexToDec = HexToDec + 16 ^ (Len(pszHex) - i) * HexCharToDec(Mid(pszHex, i, 1))
    Next i
End Function

' *******************************************************************
' *   Brief Description: ���ܿ���                                   *
' *   Engineer: ½����                                              *
' *   Date Generated: 2001/02/16                                    *
' *   Last Revision Date:                                           *
' *******************************************************************
Public Function UnEncryptPassword(ByVal pszEncryptedPassword As String) As String
'pszEncodedPassword ����ܿ���
    
    Dim nLen As Integer
    Dim i As Integer
    Dim szResult As String
    Dim nIndex As Integer

    If Len(pszEncryptedPassword) = 0 Then Exit Function
    
    szResult = pszEncryptedPassword
    nLen = Val(Right(szResult, 1) & Left(szResult, 1))
    szResult = Mid(szResult, 2, Len(szResult) - 2)
    
    Dim nTmp As Integer
    nTmp = Len(szResult) / 3
    szResult = Right(szResult, nTmp) & Left(szResult, Len(szResult) - nTmp)
    szResult = XOREncrypt(szResult)
    szResult = Right(szResult, nTmp) & Left(szResult, Len(szResult) - nTmp)
    szResult = XOREncrypt(szResult)
    
    For i = 1 To nLen
       UnEncryptPassword = UnEncryptPassword & Chr(HexToDec(Mid(szResult, (i - 1) * 2 + 1, 2)))
    Next i
    
End Function
Private Function HexCharToDec(pszChar As String) As Integer
    Select Case UCase(pszChar)
        Case "0" To "9"
            HexCharToDec = Val(pszChar)
        Case "A" To "F"
            HexCharToDec = 10 + Asc(pszChar) - Asc("A")
        Case Else
            HexCharToDec = 0
    End Select
End Function



Public Function MakeDisplayStringEx(ByVal pszID As String, ByVal pszName As String) As String
    '����ǰ,�����ں�
    MakeDisplayStringEx = pszName & EncodeString(pszID)
    
End Function

Public Function MakeDisplayString(ByVal pszID As String, ByVal pszName As String) As String
    '������ǰ,���ں�
    MakeDisplayString = pszID & EncodeString(pszName)
End Function




Public Function GetNumber(Optional dbValue As Double = 0, Optional nBit As Integer = 0) As String
  Dim szResult As Double
  Dim szChar As String
  Dim szmChar As String
  Dim sztChar As String
  Dim szdChar As String
  Dim bMillon As Boolean
  Dim j As Integer
  Dim i As Integer
  Dim szNum(9) As String
  Dim nValue As Double
  Dim nCount As Integer
  Dim dValue As Integer
  Dim nType As Integer
  On Error Resume Next
  
  dbValue = dbValue * 10 ^ nBit
  nValue = LeftAndRight(Trim(Str(Format(Abs(dbValue), "0.00"))), True, ".")
  If nValue < 0 Then
      nCount = Len(Trim(Str(nValue))) - 1
  Else
      nCount = Len(Trim(Str(nValue)))
  End If
  szNum(0) = "��"
  szNum(1) = "Ҽ"
  szNum(2) = "��"
  szNum(3) = "��"
  szNum(4) = "��"
  szNum(5) = "��"
  szNum(6) = "½"
  szNum(7) = "��"
  szNum(8) = "��"
  szNum(9) = "��"
  If nCount > 12 Then Exit Function
 For i = 1 To nCount
Next1:
                szResult = nValue - Int(nValue / 10) * 10
                j = j + 1
                If j = 1 Then
                      If szResult = 0 Then
                        szChar = ""
                      Else
                        szChar = szNum(szResult)
                      End If
                ElseIf j = 2 Then
                      If szResult = 0 Then
                         If szChar <> "" Then
                           szChar = szNum(szResult) & szChar
                         Else
                           szChar = ""
                         End If
                      Else
                         szChar = szNum(szResult) & szTen & szChar
                      End If
                ElseIf j = 3 Then
                      If szResult = 0 Then
                          If Left(szChar, 1) <> "��" And szChar <> "" Then
                              szChar = szNum(szResult) & szChar
                          End If
                      Else
                         szChar = szNum(szResult) & szHundred & szChar
                      End If
                ElseIf j = 4 Then
                      If szResult = 0 Then
                          If Left(szChar, 1) <> "��" And szChar <> "" Then
                             szChar = szNum(szResult) & szChar
                          End If
                      Else
                          szChar = szNum(szResult) & szThousand & szChar
                      End If
                ElseIf j >= 5 And j < 9 And bMillon = False Then
                      szmChar = szChar
                      j = j - 5
                      bMillon = True
                      nType = 1
                      GoTo Next1
                ElseIf i >= 9 Then
                      sztChar = szChar & szMillion & szmChar
                      j = i - 9
                      nType = 2
                      GoTo Next1
               End If
               nValue = Int(nValue / 10)
    
 Next i
 dValue = (Abs(dbValue) - Int(Abs(dbValue))) * 100
 If dValue >= 10 Then
       If Val(Mid(dValue, 2, 1)) <> 0 Then
           szdChar = szNum(Val(Mid(dValue, 1, 1))) & "��" & szNum(Val(Mid(dValue, 2, 1))) & "��"
       Else
           szdChar = szNum(Val(Mid(dValue, 1, 1))) & "��"
       End If
  Else
       If dValue <> 0 Then
         szdChar = szNum(Val(Mid(dValue, 1, 1))) & "��"
       End If
  End If
 
  If nType = 2 Then
     GetNumber = szChar & szTmillion & sztChar & "Ԫ" & szdChar
  ElseIf nType = 1 Then
     GetNumber = szChar & szMillion & szmChar & "Ԫ" & szdChar
  Else
     GetNumber = szChar & "Ԫ" & szdChar
  End If
  If dbValue < 0 Then
      GetNumber = "-" & GetNumber
  End If
End Function


Public Function ApartBaseFig(ByVal nNumber As String, Optional ByVal bolRead As Boolean = False) As String() 'nNumber<9999
    Dim i As Integer, j As Integer, nLength As Integer
    Dim szResult() As String, szReadResult As String
    Dim bolZero As Boolean, bolLastZero As Boolean
    Dim nBit As Integer
    Dim szNum(9) As String
    Dim szBit(1 To 5) As String
    Dim szTemp As String
    
    szBit(1) = ""
    szBit(2) = "ʰ"
    szBit(3) = "��"
    szBit(4) = "Ǫ"
    szBit(5) = "��"
    
    szNum(0) = "��"
    szNum(1) = "Ҽ"
    szNum(2) = "��"
    szNum(3) = "��"
    szNum(4) = "��"
    szNum(5) = "��"
    szNum(6) = "½"
    szNum(7) = "��"
    szNum(8) = "��"
    szNum(9) = "��"
    If Trim(nNumber) = "" Then Exit Function
    bolZero = False
    bolLastZero = True
    szTemp = Trim(nNumber)
    nLength = Len(szTemp)
    ReDim szResult(1 To nLength) As String
    For i = nLength To 1 Step -1
        j = j + 1
        nBit = CInt(Mid(szTemp, i, 1))
        szResult(j) = szNum(nBit)
        If nBit = 0 Then
            If bolLastZero = False And bolZero = False Then
                szReadResult = szResult(j) & szReadResult
            End If
            bolZero = True
        Else
            szReadResult = szResult(j) & szBit(j) & szReadResult
            bolLastZero = False
            bolZero = False
        End If
    Next i
    If bolRead = True Then
        ReDim szResult(1 To 1) As String
        szResult(1) = szReadResult
    End If
    ApartBaseFig = szResult
End Function

Public Function ApartFig(ByVal pnNumber As Double) As String()
    Dim i As Integer, j As Integer
    Dim sFirst As String, sSecond As String
    Dim sThird As String, sLast As String
    Dim szFirst() As String, szSecond() As String
    Dim szThird() As String, szLast() As String
    Dim nPoint As Integer, nLength As Integer
    Dim szResult() As String, sTemp() As String
    Dim nCount As Integer, nTemp As Integer
    Dim szTemp As String
    
    '1,2345,6789.00
    szTemp = Trim(CStr(Format(pnNumber, "0.00")))
    nPoint = InStr(1, szTemp, ".")
    sLast = Trim(Mid(szTemp, nPoint + 1, 2))
    sFirst = Trim(Mid(szTemp, 1, nPoint - 1))
    nLength = Len(sFirst)
    nCount = 4
    If nLength >= 5 Then
        sSecond = Trim(Mid(sFirst, 1, nLength - 4))
        sFirst = Trim(Right(sFirst, 4))
        nCount = nCount + 1
        If nLength >= 9 Then
            sThird = Trim(Mid(sSecond, 1, Len(sSecond) - 4))
            sSecond = Trim(Right(sSecond, 4))
            nCount = nCount + 1
        End If
    End If
    ReDim szResult(1 To nCount + 2) As String
    szLast = ApartBaseFig(sLast)
    szFirst = ApartBaseFig(sFirst)
    nTemp = UBound(szFirst) + 2
    For i = 1 To 2
        szResult(i) = szLast(i)
    Next i
    For j = i To nTemp
        szResult(j) = szFirst(j - 2)
    Next j
    If nLength >= 5 Then
        sTemp = ApartBaseFig(sSecond, True)
        szResult(nTemp + 1) = sTemp(1)
        If nLength >= 9 Then
            sTemp = ApartBaseFig(sThird, True)
            szResult(nTemp + 2) = sTemp(1)
        End If
    End If
    ApartFig = szResult
End Function

'/////////////////////////////////////
'�����ַ���λ�ַ�����
Public Function StringToTeam(pszString As String) As String()
    Dim i As Integer
    Dim aszTemp() As String
    Dim nLen As Integer
    Dim szLeft As String
    Dim szRight As String
    nLen = 0
    For i = 1 To Len(pszString)
        If Mid(pszString, i, 1) = "," Then nLen = nLen + 1
    Next
    If nLen = 0 Then
        ReDim aszTemp(1 To 1)
        aszTemp(1) = Trim(pszString)
    Else
        ReDim aszTemp(1 To nLen + 1)
        szLeft = LeftAndRight(pszString, True, ",")
        szRight = LeftAndRight(pszString, False, ",")
        For i = 1 To nLen
           aszTemp(i) = Trim(szLeft)
           szLeft = LeftAndRight(szRight, True, ",")
           szRight = LeftAndRight(szRight, False, ",")
        Next
        aszTemp(nLen + 1) = Trim(szLeft)
    End If
    StringToTeam = aszTemp
End Function

'������ת�����ַ���,ֻת����ά����
Public Function TeamToString(paszString() As String, Optional pNo As Integer = 1) As String
'pNo���ڼ�ά

    Dim i As Integer
    Dim szTemp As String
    Dim nCount As Integer
    nCount = ArrayLength(paszString)
    For i = 1 To nCount - 1
        szTemp = szTemp & "'" & Trim(paszString(i, pNo)) & "',"
    Next i
    If nCount > 0 Then
        szTemp = szTemp & "'" & Trim(paszString(i, pNo)) & "'"
    End If
    
    TeamToString = szTemp
    
    
End Function

'��VB������ʱ����ת��Ϊ���ݿ�������ͣ��˴�������ʱ��ӦΪ���ڣ�
Public Function ToDBDate(pdtIn As Date) As String
    If VBDateIsEmpty(pdtIn) Then
        ToDBDate = cszEmptyDateStr
    Else
        ToDBDate = Format(pdtIn, cszDateStr)
    End If
End Function

'��VB������ʱ����ת��Ϊ���ݿ��ʱ���ͣ��˴�������ʱ��ӦΪʱ�䣩
Public Function ToDBTime(pdtIn As Date) As String
    'ToDBTime = cszEmptyDateStr & " " & Format(pdtIn, cszTimeStr)
    ToDBTime = Format(pdtIn, cszTimeStr)
End Function

'��VB������ʱ����ת��Ϊ���ݿ������ʱ���ͣ��˴�������ʱ��ӦΪ����ʱ�䣩
Public Function ToDBDateTime(pdtIn As Date) As String
    ToDBDateTime = Format(pdtIn, cszDateTimeStr)
End Function

'��ʽ�������ݿ��з��ص�ֵ����Ҫ�Ը��գ�NULL��ֵ
Public Function FormatDbValue(pvaIn As Field) As Variant
    If Not IsNull(pvaIn.Value) Then
        If VarType(pvaIn.Value) = vbString Then
            FormatDbValue = Trim(pvaIn.Value)
        Else
            FormatDbValue = pvaIn.Value
        End If
    End If
End Function


'ȡ�ַ������ַ�����
Public Function NumStr(InString As String, FCHAR As String) As Integer
   Dim n As Integer, i As Integer
   Dim szTemp As String
   n = 1
   szTemp = InString
   Do While n <> 0
      n = InStr(1, szTemp, FCHAR)
      If n = 0 Then Exit Do
      i = i + 1
      szTemp = Right(szTemp, Len(szTemp) - n)
   Loop
   NumStr = i + 1
End Function


Public Function ValidationMoney(ByRef pvValue As Variant, ByVal pszName As String, Optional pbAboveZero As Boolean = False) As Boolean
'��֤����Ƿ���Ч��Ҳ������֤����Ƿ�����㣩
'����˵��
'pvValue Ҫ��֤��ֵ, pszName ֵ������(������Ϣ����), pbAboveZero�Ƿ�ֻ�Ǵ�����,ȱʡ������
    If Not IsNumeric(pvValue) Then
        MsgBox EncodeString(pszName) & "����Ϊ���֣�", vbInformation, "�������"
        ValidationMoney = False
        Exit Function
    Else
        If pbAboveZero Then
            If pvValue < 0 Then
                MsgBox EncodeString(pszName) & "������ڵ���0��", vbInformation, "�������"
                ValidationMoney = False
                Exit Function
            End If
        End If
    End If
'    pvValue = FormatMoney(pvValue)
    ValidationMoney = True
End Function


Public Sub SetBusy()
    Screen.MousePointer = vbHourglass
End Sub

Public Sub SetNormal()
    Screen.MousePointer = vbArrow
End Sub

Public Function GetFirstMonthDay(pdyDate As Date) As Date
    '�õ��������ڵĸ����µĵ�һ��
    GetFirstMonthDay = Format(pdyDate, "YYYY-MM") & "-01"
End Function

Public Function GetLastMonthDay(pdyDate As Date) As Date
    '�õ��������ڵĸ����µ����һ��
    GetLastMonthDay = DateAdd("d", -1, DateAdd("m", 1, Format(pdyDate, "YYYY-MM") & "-01"))
End Function


Public Function FindQuotationMarks(szStr As String, Optional szName As String = 0) As Boolean
    
    FindQuotationMarks = True
    If InStr(1, szStr, """") <> 0 Or InStr(1, szStr, "'") <> 0 Then
        
        MsgBox szName & "�в����е����Ż�˫����", vbCritical, "����"
        FindQuotationMarks = False
        
    End If
    
End Function


'��TransferStopTime��������
'ת��ͣ��ʱ��
'��3.4 ת��Ϊ3:40
Public Function GetStopTime(psgStopTime As String, Optional pbIsRegular As Boolean = True) As Variant
    Dim szPrefix As String
    Dim szSuffix As String
    Dim szStopTime As String
    Dim vTemp As Variant
    
    szStopTime = Format(psgStopTime, ".00")
    szPrefix = LeftAndRight(szStopTime, True, ".")
    szSuffix = LeftAndRight(szStopTime, False, ".")
    
    If pbIsRegular Then
         If Left(szSuffix, 1) <> "0" Then
          vTemp = CInt(szPrefix) * 60 + CInt(Val((Left(szSuffix, 1)))) * 10 + CInt(Val((Right(szSuffix, 1))))
        Else
          vTemp = CInt(szPrefix) * 60 + CInt(Val(szSuffix))
        End If
    Else
        vTemp = Format(szPrefix & ":" & szSuffix, "hh:mm")
    End If
    GetStopTime = vTemp
End Function



Public Function GetEncodedKey(ByVal pszOrgCode As String) As String
    GetEncodedKey = "A" & pszOrgCode
End Function



'����������ȡΨһ������
Public Function RemoveSameString(szTemp() As String) As String()
    Dim i As Integer, nCount As Integer
    Dim cCount As Integer, j As Integer
    Dim aszTemp() As String
    Dim bIsSame As Boolean
    Dim n As Integer
    n = 1
    ReDim aszTemp(1 To 1) As String
    
    nCount = ArrayLength(szTemp)
    aszTemp(1) = szTemp(1)
    For i = 1 To nCount
        cCount = ArrayLength(aszTemp)
        For j = 1 To cCount
            If aszTemp(j) = szTemp(i) Then
                bIsSame = True
                Exit For
            Else
                bIsSame = False
            End If
        Next j
        If bIsSame = False Then
            n = n + 1
            ReDim Preserve aszTemp(1 To n) As String
            aszTemp(n) = szTemp(i)
        End If
    Next i
    RemoveSameString = aszTemp
End Function


'�ж�ĳһ�ַ����Ƿ������ĳһ������
Public Function IsExistParent(szParent As String, szSon As String) As Boolean
    Dim aszTemp() As String
    Dim i As Integer
    Dim nCount As Integer
    aszTemp = StringToTeam(szParent)
    nCount = ArrayLength(aszTemp)
    For i = 1 To nCount
        If aszTemp(i) = Trim(szSon) Then
            IsExistParent = True
            Exit For
        Else
            IsExistParent = False
        End If
    Next
End Function

Public Function Max(ByVal plFirst As Long, ByVal plSecond) As Long
    Max = IIf(plFirst > plSecond, plFirst, plSecond)
End Function

Public Function Min(ByVal plFirst As Long, ByVal plSecond) As Long
    Min = IIf(plFirst > plSecond, plSecond, plFirst)
End Function


