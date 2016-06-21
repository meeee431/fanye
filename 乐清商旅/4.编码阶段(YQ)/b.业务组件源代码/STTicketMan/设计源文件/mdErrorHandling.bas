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
Public Const cszOrgErrMsgNote = "[Դ��������]:"
Public Const cszCustomErrSource = "�Զ������"

'������󹫹�����
'---------------------------------------------------------
Public Const ERR_CoreErrorStart = 10000      '���Ĵ�������
Public Const ERR_DBObjErrorStart = 10500     '���ݶ����������
Public Const ERR_BusObjErrorStart = 11000    'ҵ�������������
Public Const ERR_ClientToolStart = 12000    '�ͻ��˹��ߴ�������
Public Const ERR_DLLStep = 50               'ÿ������ɷ���Ĵ���ŷ�Χ

'��ʵ������������޸ġ�ɾ���Ĵ����λ�ƣ��������Ĵ����Ҫ�������Ĵ�����ʼ�ţ�
Public Enum ENormalOffset
    ERR_AddDuplicate = 1  '��������ʱָ���Ķ����Ѿ�����
    ERR_AddNoParent = 2  '��������ʱ������Ҫ��ĸ�������
    
    ERR_EditNoParent = 4  '�޸Ķ���ʱ������Ҫ��ĸ�������
    ERR_EditChildExist = 5  '�޸Ķ���ʱ�����������Ӵ���
    
    ERR_DeleteChildExist = 7  'ɾ������ʱ�����Ӵ���
    ERR_DeleteNotExist = 8  'ɾ������ʱ��ָ���Ķ���
    
    '---------------------------------------------
    ERR_NoActiveUser = 11     '����δ���û�û�����
    ERR_NotAvailable = 12     '�������ڴ��ڲ�����״̬
    ERR_EditObj = 13    '�������ڴ����޸�״̬
    ERR_AddObj = 14     '�������ڴ�������״̬
    ERR_NormalObj = 15     '�������ڴ�������״̬
    ERR_NotAddObj = 16    '�������ڲ���������״̬
End Enum

Public Enum EErrDatabase_Trigger    '����������
    ERR_DBAddNoParent = 30002 '����ʱ��������
    
    ERR_DBEditNoParent = 30003 '��������,�ӱ����޸�
    ERR_DBEditChildExist = 30005 '�ӱ����,�������޸�
    
    ERR_DBDeleteChildExist = 30006 '�ӱ����,����ɾ������
End Enum

'---------------------------------------------------------
Public Const ERR_TransContext_MTSNotSupport = ERR_CoreErrorStart + 0 'MTS��֧��!
Public Const ERR_InvalidProgramer = ERR_CoreErrorStart + 1 'δ����Ȩ�Ŀ�����!
Public Const ERR_AssertCreditCard = ERR_CoreErrorStart + 2 '�û�����޷�ʶ��!
Public Const ERR_INIFileNotValid = ERR_CoreErrorStart + 3 '�Ƿ�ϵͳ�����ļ�!
Public Const ERR_DatabaseConnectErr = ERR_CoreErrorStart + 4   'ͨ�ô����ݿ����!
Public Const ERR_DuplicatePrimaryKeyErr = ERR_CoreErrorStart + 5 '�ü�¼�Ѿ����ڣ��޷����!
Public Const ERR_ForeignKeyConflict = ERR_CoreErrorStart + 6    '�ü�¼�������ط��Ѿ������ã��޷��޸�!
Public Const ERR_ColumnName = ERR_CoreErrorStart + 7  '��¼���ֶ��������ݿⲻ��!
Public Const ERR_IDName = ERR_CoreErrorStart + 8  '����ʱ����¼���ֶ��в��ܸ�������ֵ!
Public Const ERR_ColumnNum = ERR_CoreErrorStart + 9 '��¼���ֶ��������ݿⲻ��!
Public Const ERR_SameRecordExist = ERR_CoreErrorStart + 10  'ͬ����¼�Ѵ���!
Public Const ERR_BuiltinRecordNotDelete = ERR_CoreErrorStart + 11  'ϵͳ���ü�¼������ɾ��!
Public Const ERR_InterfaceNotUsable = ERR_CoreErrorStart + 12 '�ӿ���ʱ������

    


' *******************************************************************
' *   Brief Description: ͨ�ô��󴥷�����                           *
' *   Engineer: ½����                                              *
' *   Date Generated: 2002/01/20                                    *
' *   Last Revision Date:                                           *
' *******************************************************************
Public Sub RaiseError(ByVal plErrNum As Long, Optional ByVal pszSource As String = "", Optional ByVal pszOrgMsg As String = "", Optional ByVal pbWantOrgMsg As Boolean = False)
    Dim szErrMsg As String, szSource As String
'    On Error GoTo Here
On Error Resume Next
    If pszSource <> "" Then
        szSource = pszSource
    Else
        szSource = err.Source
    End If
    
    '�õ�Ԥ����Ĵ����
    szErrMsg = LoadResString(plErrNum)
    On Error GoTo 0
    If szErrMsg = "" Then
        If pszOrgMsg <> "" Then
            If Left(pszOrgMsg, Len(cszOrgErrMsgNote)) <> cszOrgErrMsgNote Then
                szErrMsg = cszOrgErrMsgNote & pszOrgMsg
            Else
                szErrMsg = pszOrgMsg
            End If
        Else
            szErrMsg = cszDefErrMsg
            If pbWantOrgMsg Then szErrMsg = szErrMsg & vbCrLf & cszOrgErrMsgNote & pszOrgMsg
        End If
    End If
     
'    '�õ�ϵͳ�����Ƿ���ʾԭʼ������Ϣ������NOTES���޷��õ���������������
'    If Left(szErrMsg, 2) <> "����" Then
'        szErrMsg = "���� " & plErrNum & vbCrLf & szErrMsg
'    End If
    err.Raise plErrNum, szSource, szErrMsg
End Sub

' *******************************************************************
' *   Brief Description: ͨ�ô��󴥷�����(���ݾɰ�ϵͳ)             *
' *   Engineer: ½����                                              *
' *   Date Generated: 2002/01/20                                    *
' *   Last Revision Date:                                           *
' *******************************************************************
Public Sub ShowError(ByVal plErrNum As Long, Optional ByVal pszSource As String = "", Optional ByVal pszOrgMsg As String = "", Optional ByVal pbWantOrgMsg As Boolean = False)
    Dim szErrMsg As String, szSource As String
'    On Error GoTo Here
On Error Resume Next
    If pszSource <> "" Then
        szSource = pszSource
    Else
        szSource = err.Source
    End If
    
    '�õ�Ԥ����Ĵ����
    szErrMsg = LoadResString(plErrNum)
    On Error GoTo 0
    If szErrMsg = "" Then
        If pszOrgMsg <> "" Then
            If Left(pszOrgMsg, Len(cszOrgErrMsgNote)) <> cszOrgErrMsgNote Then
                szErrMsg = cszOrgErrMsgNote & pszOrgMsg
            Else
                szErrMsg = pszOrgMsg
            End If
        Else
            szErrMsg = cszDefErrMsg
            If pbWantOrgMsg Then szErrMsg = szErrMsg & vbCrLf & cszOrgErrMsgNote & pszOrgMsg
        End If
    End If
     
'    '�õ�ϵͳ�����Ƿ���ʾԭʼ������Ϣ������NOTES���޷��õ���������������
'    If Left(szErrMsg, 2) <> "����" Then
'        szErrMsg = "���� " & plErrNum & vbCrLf & szErrMsg
'    End If
    err.Raise plErrNum, szSource, szErrMsg
End Sub
'��֤�����������
Public Sub AssertAddObjectError(ByVal plErrBegin As Long, poDb As RTConnection)
    If poDb.HaveThisNativeErr(ERR_DBAddDuplicate) Then   '����������ش�
        RaiseError plErrBegin + ERR_AddDuplicate
    ElseIf poDb.HaveThisNativeErr(ERR_DBAddNoParent) Then '����������޸�����
        RaiseError plErrBegin + ERR_AddNoParent
    Else
        RaiseError err.Number, , err.Description
    End If
End Sub
'��֤���Ķ������
Public Sub AssertUpdateObjectError(plErrBegin As Long, poDb As RTConnection)
    If poDb.HaveThisNativeErr(ERR_DBEditChildExist) Then '����Ǹ����Ӵ��ڴ�
        RaiseError plErrBegin + ERR_EditChildExist
    ElseIf poDb.HaveThisNativeErr(ERR_DBEditNoParent) Then '����Ǹ��¸������ڴ�
        RaiseError plErrBegin + ERR_EditNoParent
    Else
        RaiseError err.Number, , err.Description
    End If
End Sub

'��֤ɾ���������
Public Sub AssertDeleteObjectError(plErrBegin As Long, poDb As RTConnection)
    If poDb.HaveThisNativeErr(ERR_DBDeleteChildExist) Then '�����ɾ���Ӵ��ڴ�
        RaiseError plErrBegin + ERR_DeleteChildExist
    Else
        RaiseError err.Number, , err.Description
    End If
End Sub
'�ָ�ɾ���������
Public Sub AssertReCoverObjectError(plErrBegin As Long, poDb As RTConnection)
    If poDb.HaveThisNativeErr(ERR_DBDeleteChildExist) Then '����ǻָ��Ӵ��ڴ�
        RaiseError plErrBegin + ERR_DeleteChildExist
    Else
        RaiseError err.Number, , err.Description
    End If
End Sub
