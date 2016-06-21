Attribute VB_Name = "mdSMC"
'Declare Public Variant Module

Option Explicit
Public Const cszMsg = "ϵͳ����"

Public Const cszConRoot = "Root"
Public Const cszUserGroupMan = "UserGroupMan"
Public Const cszLogMan = "LogMan"
Public Const cszLoginLogMan = "LoginLogMan"
Public Const cszOperateLogMan = "OperateLogMan"
Public Const cszUnitMan = "UnitMan"
Public Const cszFunctionMan = "FunctionMan"
Public Const cszActiveUserMan = "ActiveUserMan"
Public Const cszFun_GroupMan = "Fun_GroupMan"
Public Const cszComponent = "Component"
Public Const cszStationMan = "StationMan"
Public Const cszSystemMan = "SystemMan"

Public Const cGreyColor = &HC0C0C0


Public Enum EAEUserGroup
    MProperty = 0
    mPropertyGroup = 1
    MDelectUse = 2
    mDelectGroup = 3
    MAddUser = 4
    MAddGroup = 5
    
    
End Enum

Public Enum EAEActUser
    MActUserRefresh = 1
    MFroceLogout = 2
End Enum

Public Enum EAEUnit
    MAddUnit = 1
    MDeleteUnit = 2
    MRecoverUnit = 3
End Enum

Public Enum EAEFunction
    MAddCOM = 1
    MDeleteCOM = 2
    MFunGroup = 3
End Enum

Public Enum EEOperateLog
    MSelect = 1
    MDeleteSel = 2
    MSelfDelete = 3
End Enum


'***************************************
'ϵͳ�����Զ�������(STSysMan\SystemMan����Ҳ�в��ֶ���)
Public Type TRemoteInfo
    szRemoteID As String
    szLocalIDa As String
    g_szPassword As String
    szAnnotation As String
End Type

Public Type TCOMFunctionInfoShort
    FunID As String
    FunName As String
    FunGroup As String
End Type

Public Type TUserGroupShort
    GroupID As String
    GroupName As String
End Type

Public Type TUserShort
    UserID As String
    UserName As String
End Type


Public g_szCurrentTask As String

Public m_frmActiveSMC As frmStoreMenu


'**********�ڴ�ά������
'(lvDetail\lvDetail2��ѡ�е�Item.Text����)
Public g_alvItemText() As String
Public g_alvItemText2() As String
 
Public g_atAllUnit() As TUnit '���е�λ��Ϣ
Public g_atAllSellStation() As TDepartmentInfo  '������Ʊ��վ��Ϣ

Public g_atAllUnitDelTag() As TUnit '�����Ѵ�ɾ����ǵĵ�λ
Public g_atAllUserInfo() As TUserInfo '�����û���Ϣ
Public g_atUserInfo() As TUserInfo '����δ��ɾ����ǵ��û�
Public g_atUserInfoDelTag() As TUserInfo '�����Ѵ�ɾ����ǵ��û�(������λ��ɾ��(��־)�ĳ���)
Public g_atUserGroupInfo() As TUserGroupInfo '��������Ϣ
Public g_atAllFun() As TCOMFunctionInfo1  '���й�������,���ڳ�ʼ��tvAuthor,��cmdSelect(0)����

'�û��趨��ʱʹ��
Public g_atAllGroup() As TUserGroupShort
Public g_atBelongGroup() As TUserGroupShort
Public g_atBelongGroupOld() As TUserGroupShort
Public g_atUnBelongGroup() As TUserGroupShort
Public g_bRightNull As Boolean '��עadGroup�ұ�Ϊ��
Public g_bLeftNull As Boolean '��עadGroup���Ϊ��

'��Ȩʱʹ��
Public g_atAuthored() As TCOMFunctionInfoShort '��ʼ��ʱlvbrowse������,����û�����ȡ����,Ϊ���������
Public g_atAddBrowse() As TCOMFunctionInfoShort 'tvAuthor������ѡ�е�����
Public g_atInBrowse() As TCOMFunctionInfoShort 'lvBrowse�����е�����,Ҳ�����������
Public g_aszFunOld() As String 'ԭ�е�Ȩ��(���û���˵Ϊֱ��Ȩ��)
Public g_bBrowseNull As Boolean '��ע��ȨΪ��

'�û�����Ӻ�ɾ���û�ʱʹ��
Public g_atIncludUser() As TUserShort
Public g_atIncludUserOld() As TUserShort
Public g_atExcludUser() As TUserShort

'�޸�Զ���û�ʱʹ��
Public g_atRemoteUser() As TRemoteUserInfo 'frmUnitBeUser
Public g_aszUsedLocUser() As String 'frmUnitBeUser��frmaddremoteUser(��ע�ѷ�����˵�λ�ı����û�����)

'����λ
Public g_szLocalUnit As String

Public g_oSysMan As New SystemMan
Public g_oSysParam  As New SystemParam
'��ʱ(��¼)
Public g_oActUser As ActiveUser
Public g_oLogin As New CommShell

'��������Ȩ
Public g_aszUser() As String
Public g_aszUserGroup() As String
Public g_aszUserAdd() As String
Public g_aszUserGroupAdd() As String

'************�ڴ�ά������
Public g_bShowUserInfo As Boolean '����ı䵱ǰ�û�����ʱCommonDG���治���е�Bug
Public g_szPassword As String

Public Sub Main()
    Dim szLoginCommandLine As String
    szLoginCommandLine = TransferLoginParam(Trim(Command()))
    
    If App.PrevInstance Then
        End
    End If
    If szLoginCommandLine = "" Then
    
        g_szCurrentTask = cszConRoot
        Set g_oActUser = g_oLogin.ShowLogin()
    Else
        Set g_oActUser = New ActiveUser
        g_oActUser.Login GetLoginParam(szLoginCommandLine, cszUserName), GetLoginParam(szLoginCommandLine, cszUserPassword), GetComputerName()
    End If
    If g_oActUser Is Nothing Then
        Exit Sub
    Else
        App.HelpFile = SetHTMLHelpStrings("PSTSysMan.chm")
        'g_oLogin.ShowSplash  "rtStationϵͳ����", "rtStation System Manager Console", , App.Major, App.Minor,  App.Revision
        frmSMCMain.Show
    End If
    SetHTMLHelpStrings "STSysMan.chm"
End Sub
