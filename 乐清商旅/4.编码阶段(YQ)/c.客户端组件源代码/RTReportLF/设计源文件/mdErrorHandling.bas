Attribute VB_Name = "mdErrorLib"
'=========================================================================================
'Author:
'Detail:�м�㹲��ģ��
'=========================================================================================


'=========================================================================================
'���³�������
'=========================================================================================
'ϵͳ��������
'---------------------------------------------------------
Public Const cszDefErrMsg = "��������Դ�ļ����Ҳ�����Ӧ����Դ�����Դ˴��������Ϊ�ա�"
Public Const cszOrgErrMsgNote = "Դ��������:"
Public Const cszCustomErrSource = "�Զ������"

'������󹫹�����
'---------------------------------------------------------
Public Const ERR_CoreErrorStart = 10000      '���Ĵ�������
Public Const ERR_DBObjErrorStart = 10500     '���ݶ����������
Public Const ERR_BusObjErrorStart = 11000    'ҵ�������������
Public Const ERR_ClientToolStart = 12000    '�ͻ��˹��ߴ�������
Public Const ERR_DLLStep = 50               'ÿ������ɷ���Ĵ���ŷ�Χ
Public Const ERR_DLLIndex = 1               '���������

'---------------------------------------------------------
Public Const ERR_InvalidProgramer = ERR_CoreErrorStart + 1 'δ����Ȩ�Ŀ�����!
Public Const ERR_AssertCreditCard = ERR_CoreErrorStart + 2 '�û�����޷�ʶ��!
Public Const ERR_INIFileNotValid = ERR_ClientToolStart + 3 '�Ƿ�ϵͳ�����ļ�!

'---------------------------------------------------------
Public Const ERR_FileNotExist = ERR_ClientToolStart + ERR_DLLIndex * ERR_DLLStep + 1    '�ļ�������
Public Const ERR_TemplateFileNotFound = ERR_ClientToolStart + ERR_DLLIndex * ERR_DLLStep + 2    'ģ���ļ��Ҳ���!
Public Const ERR_FileExportError = ERR_ClientToolStart + ERR_DLLIndex * ERR_DLLStep + 3 '�ļ���������
Public Const ERR_FileSaveError = ERR_ClientToolStart + ERR_DLLIndex * ERR_DLLStep + 4 '�ļ��������
Public Const ERR_FileTypeNotSupport = ERR_ClientToolStart + ERR_DLLIndex * ERR_DLLStep + 5 '�ļ����ʹ���!
Public Const ERR_BodyDataInvalid = ERR_ClientToolStart + ERR_DLLIndex * ERR_DLLStep + 6 '�������ݲ�������ȷ!


' *******************************************************************
' *   Brief Description: ͨ�ô��󴥷�����                           *
' *   Engineer: ½����                                              *
' *   Date Generated: 2002/01/20                                    *
' *   Last Revision Date:                                           *
' *******************************************************************
Public Sub RaiseError(ByVal plErrNum As Long, Optional pszSource As String = "", Optional pszOrgMsg As String = "")
    Dim szErrMsg As String, szSource As String
'    On Error GoTo Here
On Error Resume Next
    If pszSource <> "" Then
        szSource = pszSource
    Else
        szSource = Err.Source
    End If
    
    '�õ�Ԥ����Ĵ����
    szErrMsg = LoadResString(plErrNum)
    On Error GoTo 0
    If szErrMsg = "" Then szErrMsg = cszDefErrMsg
     
    '�õ�ϵͳ�����Ƿ���ʾԭʼ������Ϣ
    Dim bShowOrgError As Boolean
    bShowOrgError = True 'ֵ���ȡע���
    If bShowOrgError Then
        szErrMsg = szErrMsg & vbCrLf & vbCrLf & cszOrgErrMsgNote & vbCrLf & pszOrgMsg
    End If
    
    If szErrMsg = cszOrgErrMsgNote Then
        szErrMsg = pszOrgMsg
    End If
    
    Err.Raise plErrNum, szSource, szErrMsg
End Sub
