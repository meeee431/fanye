Attribute VB_Name = "mdSMC"
'Declare Public Variant Module

Option Explicit
Public Const cszMsg = "系统管理"

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
'系统管理自定义类型(STSysMan\SystemMan类中也有部分定义)
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


'**********内存维护数据
'(lvDetail\lvDetail2的选中的Item.Text数组)
Public g_alvItemText() As String
Public g_alvItemText2() As String
 
Public g_atAllUnit() As TUnit '所有单位信息
Public g_atAllSellStation() As TDepartmentInfo  '所有售票车站信息

Public g_atAllUnitDelTag() As TUnit '所有已打删除标记的单位
Public g_atAllUserInfo() As TUserInfo '所有用户信息
Public g_atUserInfo() As TUserInfo '所有未打删除标记的用户
Public g_atUserInfoDelTag() As TUserInfo '所有已打删除标记的用户(所属单位已删除(标志)的除外)
Public g_atUserGroupInfo() As TUserGroupInfo '所有组信息
Public g_atAllFun() As TCOMFunctionInfo1  '所有功能数据,用于初始化tvAuthor,及cmdSelect(0)单击

'用户设定组时使用
Public g_atAllGroup() As TUserGroupShort
Public g_atBelongGroup() As TUserGroupShort
Public g_atBelongGroupOld() As TUserGroupShort
Public g_atUnBelongGroup() As TUserGroupShort
Public g_bRightNull As Boolean '标注adGroup右边为空
Public g_bLeftNull As Boolean '标注adGroup左边为空

'授权时使用
Public g_atAuthored() As TCOMFunctionInfoShort '初始化时lvbrowse的数据,如果用户单击取消键,为最后结果数据
Public g_atAddBrowse() As TCOMFunctionInfoShort 'tvAuthor单击后选中的数据
Public g_atInBrowse() As TCOMFunctionInfoShort 'lvBrowse中已有的数据,也是最后结果数据
Public g_aszFunOld() As String '原有的权利(对用户来说为直接权限)
Public g_bBrowseNull As Boolean '标注授权为空

'用户组添加和删除用户时使用
Public g_atIncludUser() As TUserShort
Public g_atIncludUserOld() As TUserShort
Public g_atExcludUser() As TUserShort

'修改远程用户时使用
Public g_atRemoteUser() As TRemoteUserInfo 'frmUnitBeUser
Public g_aszUsedLocUser() As String 'frmUnitBeUser和frmaddremoteUser(标注已分配给此单位的本地用户代码)

'本单位
Public g_szLocalUnit As String

Public g_oSysMan As New SystemMan
Public g_oSysParam  As New SystemParam
'临时(登录)
Public g_oActUser As ActiveUser
Public g_oLogin As New CommShell

'按功能授权
Public g_aszUser() As String
Public g_aszUserGroup() As String
Public g_aszUserAdd() As String
Public g_aszUserGroupAdd() As String

'************内存维护数据
Public g_bShowUserInfo As Boolean '解决改变当前用户密码时CommonDG窗替不居中的Bug
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
        'g_oLogin.ShowSplash  "rtStation系统管理", "rtStation System Manager Console", , App.Major, App.Minor,  App.Revision
        frmSMCMain.Show
    End If
    SetHTMLHelpStrings "STSysMan.chm"
End Sub
