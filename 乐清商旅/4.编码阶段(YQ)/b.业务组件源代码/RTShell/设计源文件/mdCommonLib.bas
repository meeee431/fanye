Attribute VB_Name = "mdCommonLib"
'=========================================================================================
'Author:
'Detail:通用函数模块
'=========================================================================================


Option Explicit
'------------------------------------------------------------------------------------------
'以下常量声明
'---------------------------------------------------------
Public Const cszRegKeyCompany = "Software\RTsoft"       '公司注册键
Public Const cszRegKeyProduct = "RTFood"                '产品注册键

'-----------------------------------------------------------------------------------------
Public Const cszEmptyDateStr = "1900-01-01"
Public Const cszForeverDateStr = "2050-01-01"
Public Const cszEmptyTimeStr = "00:00:00"

Public Const cszDateStr = "YYYY-MM-DD"
Public Const cszTimeStr = "hh:mm:ss"
Public Const cszDateTimeStr = cszDateStr & " " & cszTimeStr

Public Const cdtEmptyDate = #1/1/1900#
Public Const cdtEmptyTime = #12:00:00 AM#
Public Const cdtEmptyDateTime = #1/1/1900#


Public Function GetLString(InString As String) As String
    GetLString = Trim(LeftAndRight(InString, True, "["))
End Function
'获得数组的维数
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

'得到某个字符左边或右边的字符串
'例:LeftAndRight("0000[示例]",True,"[")="0000"
Public Function LeftAndRight(InString As String, IsLeft As Boolean, FCHAR As String) As String
'InString:源字符串
'IsLeft:是否左部
'FCHAR:分隔字符
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


''格式化字符串，判断是否有非法字符（缺省的非法字符是',也可自己指定），然后将字符串的后导空格去掉
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

'得到一个数组的指定维的长度，如有错，则返回0
Public Function ArrayLength(paszIn As Variant, Optional pnIndex As Integer = 1) As Long
'paszIn:数组
'pnIndex:第几维
    Dim lLow As Long, lHigh As Long
    On Error GoTo Here
    
    lLow = LBound(paszIn, pnIndex)
    lHigh = UBound(paszIn, pnIndex)
    ArrayLength = lHigh - lLow + 1
    Exit Function
Here:
    ArrayLength = 0
End Function

'判断指定的VB日期是否是空的
Public Function VBDateIsEmpty(pdtIn As Date) As Boolean
    Dim dtTemp As Date
    VBDateIsEmpty = IIf(Format(pdtIn, cszDateStr) = Format(dtTemp, cszDateStr), True, False)
End Function

'判断指定的VB时间是否是空的
Public Function VBTimeIsEmpty(pdtIn As Date) As Boolean
    Dim dtTemp As Date
    VBTimeIsEmpty = IIf(Format(pdtIn, cszTimeStr) = Format(dtTemp, cszTimeStr), True, False)
End Function

'判断指定的VB日期时间是否是空的
Public Function VBDateTimeIsEmpty(pdtIn As Date) As Boolean
    Dim dtTemp As Date
    VBDateTimeIsEmpty = IIf(Format(pdtIn, cszDateTimeStr) = Format(dtTemp, cszDateTimeStr), True, False)
End Function


'将第一个参数的日期和第二个参数时间合并为日期时间
Public Function SelfGetFullDateTime(ByVal dtDate As Date, ByVal dtTime As Date) As Date
    SelfGetFullDateTime = CDate(Format(dtDate, cszDateStr) & " " & Format(dtTime, cszTimeStr))
End Function
 

'将ID及对应的名称拼成一个显示字符串，多用于码表
'例:MakeDisplayString("00","代码一")="00[代码一]"
Public Function MakeDisplayString(ByVal pszID As String, ByVal pszName As String) As String
    MakeDisplayString = pszName & EncodeString(pszID)
End Function

'解析得到显示字符串中的ID或名称，对应MakeDisplayString函数
'例:ResolveDisplay("00[代码一]",szName)="00"  szName="代码一"
Public Function ResolveDisplay(ByVal pszDisplayString As String, Optional ByRef pszName As String = "*****") As String
'返回ID
'pszDisplayString:源字符串
'pszName:返回的名称
    Dim i  As Integer, nStrLen As Integer
    i = InStr(1, pszDisplayString, "[")
    nStrLen = Len(pszDisplayString)
    If i > 0 Then
        ResolveDisplay = Left(pszDisplayString, i - 1)
        If pszDisplayString <> "*****" Then
            pszName = Mid(pszDisplayString, i + 1, nStrLen - i - 1)
        End If
    End If
End Function

'对代码编号进行数量的相加形成新的代码编号
'例: NumAdd("00001", 5) = "00006"
Public Function NumAdd(ByVal pszSource As String, ByVal lNum As Long) As String
'pszSource:源字符串
'lNum:相加值
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

'对代码编号进行数量的相减形成新的代码编号
'例: NumSub("00006", 5) = "00001"
Public Function NumSub(ByVal pszSource As String, ByVal lNum As Long) As String
'pszSource:源字符串
'lNum:相减值
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
    
'将字符串转为有效数字
'例:GetTextToNumeric("12.5",False,False)="12"
Public Function GetTextToNumeric(pszText As String, Optional pbCanBeNegative As Boolean = True, Optional pbCanBeDecimal As Boolean = True) As String
'返回数值字符串
'pszText:源字符串
'pbCanBeNegative:是否允许负数
'pbCanBeDecimal:是否允许小数
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
            '后尾加载0
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

'返回金额显示方式
Public Function FormatMoney(pvStr As Variant) As String
    FormatMoney = Format(pvStr, "0.00")
End Function

'返回用逗号分隔的金额
Public Function FormatSeparateMoney(pvStr As Variant) As String
    FormatSeparateMoney = Format(pvStr, "##,##0.00")
End Function

'得到指定长度的Unicode字符串
'例:GetUnicodeBySize("我是VB",5)="我是V"
Public Function GetUnicodeBySize(pszString As String, pnSize As Integer) As String
    Dim szTmp As String
    szTmp = StrConv(pszString, vbFromUnicode)
    If LenB(szTmp) > pnSize Then
        GetUnicodeBySize = StrConv(LeftB(szTmp, pnSize), vbUnicode)
    Else
        GetUnicodeBySize = pszString
    End If
End Function

'ASCII是否是数值
Public Function IfNumber(nAsc As Integer) As Integer
    If nAsc >= 48 And nAsc <= 57 Or nAsc = 13 Or nAsc = 8 Then
       IfNumber = nAsc
    Else
      IfNumber = 0
       
    End If
End Function
    

'将字符串用[]括起
Public Function EncodeString(ByVal pszString As String) As String
    EncodeString = "[" & pszString & "]"
End Function

'得到[]中的字符串,对应EncodeString
Public Function UnEncodeString(ByVal pszString As String) As String
    pszString = LeftAndRight(pszString, False, "[")
    UnEncodeString = LeftAndRight(pszString, True, "]")
End Function

'形成[key=value]样式
'例:EncodeKeyValue("Key","Value",True)="[Key=Value]"
Public Function EncodeKeyValue(ByVal pszKey As String, ByVal pszValue As String, Optional ByVal pbHaveRange As Boolean = False) As String
'pszKey:键
'pszValue:值
'pbHaveRange:是否加[]
    EncodeKeyValue = pszKey & "=" & pszValue
    If pbHaveRange Then
        EncodeKeyValue = EncodeString(EncodeKeyValue)
    End If
End Function

'返回KEY和VALUE，对应EncodeKeyValue
'例:UnEncodeKeyValue("[Key=Value]",szReturnKey,szReturnValue)   szReturnKey="Key"  szReturnValue="Value"
Public Sub UnEncodeKeyValue(ByVal pszString As String, ByRef szReturnKey As String, ByRef szReturnValue As String)
'pszString:源字符串
'szReturnKey:返回的键
'szReturnValue:返回的值
    'pszString = UnEncodeString(pszString)
    If InStr(pszString, "[") <> 0 And InStr(pszString, "]") <> 0 Then
        pszString = UnEncodeString(pszString)
    End If
    szReturnKey = LeftAndRight(pszString, True, "=")
    szReturnValue = LeftAndRight(pszString, False, "=")
End Sub

'将[string1][string2]...[stringn]样式的字符组成元素数组返回
Public Function SplitEncodeStringArray(ByVal pszString As String) As String()
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


'得到中英文混合字符串的字符长度(一个中文字符代表两位)
Public Function LenA(ByVal pszString As String) As Long
    LenA = LenB(StrConv(pszString, vbFromUnicode))
End Function

'忽略unicode的字符串截取函数
'例:MidA("我是VB",3,3)="是V"
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

' *******************************************************************
' *   Brief Description: 加密口令                                   *
' *   Engineer: 陆勇庆                                              *
' *   Date Generated: 2002/06/21                                    *
' *   Last Revision Date:                                           *
' *******************************************************************
Public Function EncryptPassword(ByVal pszPassword As String) As String
'pszPassword 口令
'选择一种加密算法对口令进行加密
    
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
    szResult = Right(szResult, nLen - nTmp) & Left(szResult, nTmp) '左右互换
    szResult = XOREncrypt(szResult)
    szResult = Right(szResult, nLen - nTmp) & Left(szResult, nTmp) '左右互换
    
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
        If Len(szTmp) < cnPerNum Then   '小于异或范围内忽略
'            szTmp = Hex(HexToDec(szTmp) Xor (cnXorValue And 10 ^ (Len(szTmp)) - 1))
        Else
            szTmp = Hex(HexToDec(szTmp) Xor cnXorValue)
            If Len(szTmp) < cnPerNum Then szTmp = String(cnPerNum - Len(szTmp), "0") & szTmp   '补0
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
' *   Brief Description: 解密口令                                   *
' *   Engineer: 陆勇庆                                              *
' *   Date Generated: 2001/02/16                                    *
' *   Last Revision Date:                                           *
' *******************************************************************
Public Function UnEncryptPassword(ByVal pszEncryptedPassword As String) As String
'pszEncodedPassword 需解密口令
    
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


'格式化从数据库中返回的值，主要对付空（NULL）值
Public Function FormatDbValue(pvaIn As Field) As Variant
    If Not IsNull(pvaIn.Value) Then
        If VarType(pvaIn.Value) = vbString Then
            FormatDbValue = Trim(pvaIn.Value)
        Else
            FormatDbValue = pvaIn.Value
        End If
    End If
End Function

'得到数字的大写字符串
'例:GetNumber(1235.2,2)="壹仟贰佰参拾伍元贰角"
Public Function GetNumber(Optional dbValue As Double = 0, Optional nBit As Integer = 0) As String
'dbValue:要转换的数值
'nBit:小数后几位
  

Const szTen = "拾"
Const szTmillion = "亿"
Const szMillion = "万"
Const szThousand = "仟"
Const szHundred = "佰"
  
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
  szNum(0) = "零"
  szNum(1) = "壹"
  szNum(2) = "贰"
  szNum(3) = "叁"
  szNum(4) = "肆"
  szNum(5) = "伍"
  szNum(6) = "陆"
  szNum(7) = "柒"
  szNum(8) = "捌"
  szNum(9) = "玖"
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
                          If Left(szChar, 1) <> "零" And szChar <> "" Then
                              szChar = szNum(szResult) & szChar
                          End If
                      Else
                         szChar = szNum(szResult) & szHundred & szChar
                      End If
                ElseIf j = 4 Then
                      If szResult = 0 Then
                          If Left(szChar, 1) <> "零" And szChar <> "" Then
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
           szdChar = szNum(Val(Mid(dValue, 1, 1))) & "角" & szNum(Val(Mid(dValue, 2, 1))) & "分"
       Else
           szdChar = szNum(Val(Mid(dValue, 1, 1))) & "角"
       End If
  Else
       If dValue <> 0 Then
         szdChar = szNum(Val(Mid(dValue, 1, 1))) & "分"
       End If
  End If
 
  If nType = 2 Then
     GetNumber = szChar & szTmillion & sztChar & "元" & szdChar
  ElseIf nType = 1 Then
     GetNumber = szChar & szMillion & szmChar & "元" & szdChar
  Else
     GetNumber = szChar & "元" & szdChar
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
    szBit(2) = "拾"
    szBit(3) = "佰"
    szBit(4) = "仟"
    szBit(5) = "万"
    
    szNum(0) = "零"
    szNum(1) = "壹"
    szNum(2) = "贰"
    szNum(3) = "叁"
    szNum(4) = "肆"
    szNum(5) = "伍"
    szNum(6) = "陆"
    szNum(7) = "柒"
    szNum(8) = "捌"
    szNum(9) = "玖"
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

'解析字符串位字符数组
'例:StringToTeam("1,2,3,4")    a(1)="1":a(2)="2":a(3)="3":a(4)="4"
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

'将数组转换成字符串,只转换二维数组，对应StringToTeam
'例:a(1)="1":a(2)="2":a(3)="3":a(4)="4"  TeamToString(a)="1,2,3,4"
Public Function TeamToString(paszString() As String, Optional pNo As Integer = 1) As String
'pNo即第几维

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

'取字符串中的某字符的个数
'例:NumStr("fffjjffjj","j")=4
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
'验证金额是否有效（也可以验证金额是否大于零）
'参数说明
'pvValue 要验证的值, pszName 值的名称(错误信息中用), pbAboveZero是否只是大于零,缺省都允许
    If Not IsNumeric(pvValue) Then
        MsgBox EncodeString(pszName) & "必须为数字！", vbInformation, "输入错误"
        ValidationMoney = False
        Exit Function
    Else
        If pbAboveZero Then
            If pvValue < 0 Then
                MsgBox EncodeString(pszName) & "必须大于等于0！", vbInformation, "输入错误"
                ValidationMoney = False
                Exit Function
            End If
        End If
    End If
'    pvValue = FormatMoney(pvValue)
    ValidationMoney = True
End Function

