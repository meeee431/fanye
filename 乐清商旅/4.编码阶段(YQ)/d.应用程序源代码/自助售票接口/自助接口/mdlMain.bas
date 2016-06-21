Attribute VB_Name = "mdlMain"
Option Explicit


Public Type TFileInfomation '�ĵ���Ϣ
    TempFilePath As String  'ģ���ļ���
    FileName As String      '�ļ�˵��
    FileNote As String      '�ļ�˵��
    SplitLine As Boolean  '�Ƿ���Ҫ����
End Type

Public Const g_cszDBSetSection = "DataBaseSet"
'Public Const g_cszDocumentDir = "RegisterCode"
Public Const g_cszRegisterCode = "RegisterCode"
Public Const g_cszDefDocumentDir = "C:"

Public Const g_cszDefaultFontName = "����"
Public Const g_cszDefaultFontSize = "18"


Public g_atFileInfo() As TFileInfomation
Public g_nPageCount As Integer


'��ȡ���к�
Public Declare Function GetVolumeInformation& Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal pVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long)
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Const MAX_FILENAME_LEN = 256



'��һ���ı���ÿ�������Ⱥ�ҳ�����Զ���ֳɶ��У� �����鷵��
Public Function SplitTextByPaperSize(pszText As String, ByVal pdCharSize As Double, ByVal pdPerLineWidth As Double) As String()
    'pszTextԴ�ı�
    'pdCharSize�ַ��Ŀ��,�԰�Ϊ��λ
    'pdPerLineWidth ҳ����,�԰�Ϊ��λ
    Const cszTabString = "    " 'Tab����Ӧ���ַ���
    Dim aszSourceLine() As String
    '���س����зֳɲ�ͬ���ַ�����
    Dim szTmp As String
    Dim szLeft As String
    Dim i As Integer
    szTmp = Replace(pszText, vbTab, cszTabString)       '��Tab��ת�ɿո�
    Dim nLen As Integer
    Do
        nLen = nLen + 1
        szLeft = LeftAndRight(szTmp, True, vbCrLf)
        szTmp = LeftAndRight(szTmp, False, vbCrLf)
        ReDim Preserve aszSourceLine(1 To nLen)
        aszSourceLine(nLen) = szLeft ' IIf(Left(szLeft, 1) = vbCrLf, Right(szLeft, Len(szLeft) - 1), szLeft)
    Loop Until szTmp = ""
        
    
    Dim nLineCharNum As Integer     'ÿ��������ַ�����
    nLineCharNum = Int(pdPerLineWidth / pdCharSize)
    
    Dim aszReturnLine() As String
    Dim szTmpChar As String
    Dim nActLineLen As Integer      'ÿ�зָ�ʱ��ʵ������
    nLen = 0
    For i = 1 To ArrayLength(aszSourceLine)
        Do      '�ָ�ÿһ��
            '������ĩβ�а�����֣�����Ըð���Ĵ���
            szTmp = MidA(aszSourceLine(i), nLineCharNum + 1) '�Ұ벿��
            nActLineLen = Len(aszSourceLine(i)) - Len(szTmp)
            szTmp = Left(aszSourceLine(i), nActLineLen) '��ǰ������
            szTmpChar = Mid(aszSourceLine(i), nActLineLen + 1, 1)   '��ĩβ�Ƿ��Ǳ����ţ����������
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

'�Ƿ�Ҫ�����ı�����
Private Function IsPunctuationChar(pszChar As String) As Boolean
    Select Case pszChar
        Case ",", "��", ".", "��", "!", "��", "?", "��", ")", "��", ":", "��", ";", "��", """", "��", "'", "��", ">", "��"
            IsPunctuationChar = True
        Case Else
            IsPunctuationChar = False
    End Select
End Function


'�õ�ָ���������������к�
Public Function DriveSerial() As Long
    Dim RetVal As Long
    Dim str As String * MAX_FILENAME_LEN
    Dim str2 As String * MAX_FILENAME_LEN
    Dim a As Long
    Dim b As Long

    GetVolumeInformation GetDrive & ":\", str, MAX_FILENAME_LEN, RetVal, a, b, str2, MAX_FILENAME_LEN
    DriveSerial = RetVal
End Function

'�õ�ϵͳĿ¼���ڵ�������
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
'    szTemp = InputBox("�뽫���´������㽭��Է��˾'" & CStr(DriveSerial) & "'������ע����:", "FOYOND���-�а�ϵͳ")
'
'    On Error Resume Next
'    If MakeCode(CStr(DriveSerial)) = szTemp Then
'        oReg.Init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany
'        oReg.SaveSetting g_cszDBSetSection, g_cszRegisterCode, EncryptPassword(szTemp)
'        MdiMain.Show
'    Else
'        If i = 1 Then
'            MsgBox "��������ע�������̫��", , "ע��"
'            End
'        Else
'            i = i - 1
'            MsgBox "ע��������!����������!������" & i & "�λ���", , "ע��"
'
'            GoTo here
'        End If
'    End If
'End Sub



