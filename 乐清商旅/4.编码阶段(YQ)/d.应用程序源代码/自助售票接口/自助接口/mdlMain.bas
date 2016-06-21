Attribute VB_Name = "mdlMain"
Option Explicit


Public Type TFileInfomation '文档信息
    TempFilePath As String  '模板文件名
    FileName As String      '文件说明
    FileNote As String      '文件说明
    SplitLine As Boolean  '是否需要分行
End Type

Public Const g_cszDBSetSection = "DataBaseSet"
'Public Const g_cszDocumentDir = "RegisterCode"
Public Const g_cszRegisterCode = "RegisterCode"
Public Const g_cszDefDocumentDir = "C:"

Public Const g_cszDefaultFontName = "宋体"
Public Const g_cszDefaultFontSize = "18"


Public g_atFileInfo() As TFileInfomation
Public g_nPageCount As Integer


'读取序列号
Public Declare Function GetVolumeInformation& Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal pVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long)
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Const MAX_FILENAME_LEN = 256



'将一个文本按每个字体宽度和页面宽度自动拆分成多行， 以数组返回
Public Function SplitTextByPaperSize(pszText As String, ByVal pdCharSize As Double, ByVal pdPerLineWidth As Double) As String()
    'pszText源文本
    'pdCharSize字符的宽度,以磅为单位
    'pdPerLineWidth 页面宽度,以磅为单位
    Const cszTabString = "    " 'Tab键对应的字符串
    Dim aszSourceLine() As String
    '将回车换行分成不同的字符串行
    Dim szTmp As String
    Dim szLeft As String
    Dim i As Integer
    szTmp = Replace(pszText, vbTab, cszTabString)       '将Tab键转成空格
    Dim nLen As Integer
    Do
        nLen = nLen + 1
        szLeft = LeftAndRight(szTmp, True, vbCrLf)
        szTmp = LeftAndRight(szTmp, False, vbCrLf)
        ReDim Preserve aszSourceLine(1 To nLen)
        aszSourceLine(nLen) = szLeft ' IIf(Left(szLeft, 1) = vbCrLf, Right(szLeft, Len(szLeft) - 1), szLeft)
    Loop Until szTmp = ""
        
    
    Dim nLineCharNum As Integer     '每行允许的字符字数
    nLineCharNum = Int(pdPerLineWidth / pdCharSize)
    
    Dim aszReturnLine() As String
    Dim szTmpChar As String
    Dim nActLineLen As Integer      '每行分割时的实际字数
    nLen = 0
    For i = 1 To ArrayLength(aszSourceLine)
        Do      '分隔每一行
            '处理行末尾有半个汉字，则忽略该半个的处理
            szTmp = MidA(aszSourceLine(i), nLineCharNum + 1) '右半部分
            nActLineLen = Len(aszSourceLine(i)) - Len(szTmp)
            szTmp = Left(aszSourceLine(i), nActLineLen) '当前行内容
            szTmpChar = Mid(aszSourceLine(i), nActLineLen + 1, 1)   '行末尾是否是标点符号，如果是则补上
            If IsPunctuationChar(szTmpChar) Then
                szTmp = szTmp & szTmpChar
                nActLineLen = nActLineLen + 1
            End If
            aszSourceLine(i) = Right(aszSourceLine(i), Len(aszSourceLine(i)) - nActLineLen)
            nLen = nLen + 1
            ReDim Preserve aszReturnLine(1 To nLen)
            aszReturnLine(nLen) = szTmp
        Loop Until aszSourceLine(i) = ""
    Next i
    SplitTextByPaperSize = aszReturnLine
End Function

'是否要缩进的标点符号
Private Function IsPunctuationChar(pszChar As String) As Boolean
    Select Case pszChar
        Case ",", "，", ".", "。", "!", "！", "?", "？", ")", "）", ":", "：", ";", "；", """", "”", "'", "’", ">", "》"
            IsPunctuationChar = True
        Case Else
            IsPunctuationChar = False
    End Select
End Function


'得到指定的驱动器的序列号
Public Function DriveSerial() As Long
    Dim RetVal As Long
    Dim str As String * MAX_FILENAME_LEN
    Dim str2 As String * MAX_FILENAME_LEN
    Dim a As Long
    Dim b As Long

    GetVolumeInformation GetDrive & ":\", str, MAX_FILENAME_LEN, RetVal, a, b, str2, MAX_FILENAME_LEN
    DriveSerial = RetVal
End Function

'得到系统目录所在的驱动器
Public Function GetDrive() As String
    Dim szSystemDir As String * MAX_FILENAME_LEN
    
    GetSystemDirectory szSystemDir, MAX_FILENAME_LEN
    GetDrive = Left(Trim(szSystemDir), 1)
End Function



Public Function MakeCode(Some As String) As String
    Dim i As Long
    Dim j As Long
    Dim myKey As Long
    Dim tempStr As String
    
    myKey = 738302124
    j = 0
    For i = 1 To Len(Some)
       j = j + Asc(Mid(Some, i, 1))
    Next i
    j = j + myKey
    tempStr = CStr(j)
    If Mid(tempStr, 1, 1) = "-" Then
        tempStr = Mid(tempStr, 2, Len(tempStr) - 1)
    End If
    For i = 1 To Len(tempStr)
        MakeCode = MakeCode + Chr(Asc("a") + Int(CLng(Mid(tempStr, i, 1)) * 2.5))
    Next i
End Function



Public Function CompareRegisterCode() As Boolean
    Dim oReg As New CFreeReg
    Dim szRegisterCode As String
    On Error GoTo Error_Handle
    oReg.Init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany
    szRegisterCode = oReg.GetSetting(g_cszDBSetSection, g_cszRegisterCode, "")
    If UnEncryptPassword(szRegisterCode) = MakeCode(CStr(DriveSerial)) Then
        CompareRegisterCode = True
    End If
    
    Exit Function
Error_Handle:
    

End Function

'
'Public Sub Main()
'
'    If CompareRegisterCode Then
'        MdiMain.Show
'    Else
'        RegisterCode
'    End If
'
'End Sub
'
'
'
'
'Private Sub RegisterCode()
'    Dim szTemp As String
'    Dim oReg As New CFreeReg
'    Dim i As Integer
'    i = 3
'here:
'    szTemp = InputBox("请将以下串发给浙江方苑公司'" & CStr(DriveSerial) & "'请输入注册码:", "FOYOND软件-行包系统")
'
'    On Error Resume Next
'    If MakeCode(CStr(DriveSerial)) = szTemp Then
'        oReg.Init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany
'        oReg.SaveSetting g_cszDBSetSection, g_cszRegisterCode, EncryptPassword(szTemp)
'        MdiMain.Show
'    Else
'        If i = 1 Then
'            MsgBox "错误输入注册码次数太多", , "注意"
'            End
'        Else
'            i = i - 1
'            MsgBox "注册码有误!请重新输入!您还有" & i & "次机会", , "注意"
'
'            GoTo here
'        End If
'    End If
'End Sub



