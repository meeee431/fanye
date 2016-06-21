Attribute VB_Name = "mdSystem"
Option Explicit
Public Const cszBaseInfo = ""

Public Const cszLocalStationID = "LocalStationID" '本站站代码

Public Const cszLocalUnitID = "LocalUnitID" '本单位代码
Public Const cszCLPrefix = "A"
Public Const cszSystemMan = ""


'Public Const cszRunEnv  = ""
'Public Const cszPriceMan = ""
'Public Const cszCheckTicket = ""

Private Const cszActiveUserGroup = "ActiveUserGroup"
Private Const cszActiveUser = "ActiveUser"
Private Const cszCheckActiveUserTimer = "CheckActiveUserTimer"

Public Const ERR_CoreErrorStart = 10000      '核心错误号起点
'系统常量定义
'---------------------------------------------------------
 

Public Const cszSettle = ""

Public Const MyUnhandledError = ""

   
Public Const cszUserFunction = " user_function_lst "
Public Const cszUsergroupFunction = " usergroup_function_lst "
Public Const cszGroupUser = " Group_user_info "
Public Const cszTableSystemParam = " System_param_info"
   
   
'#If ACTIVEUSER_USE_MEMORY Then
'    Public m_clActiveUserData As Collection
'
'    #If IN_MTS Then
'        Public m_spgmActiveUserData As SharedPropertyGroupManager
'        Public m_spgActiveUserDate As SharedPropertyGroup
'        Public m_spActiveUserData As SharedProperty
'
'        Public m_spCheckActiveUserTimer As SharedProperty
'    #End If
'#Else
'
'#End If


Public Function GetNextEventID(plTemp As Long) As Long
'#If ACTIVEUSER_USE_MEMORY Then
'    Dim oTemp As ActiveUserData
'    On Error GoTo Here
'    Randomize Timer
'retry:
'    GetNextEventID = (2 ^ 31) * Rnd() + 1
'    Set oTemp = m_clActiveUserData(cszCLPrefix & GetNextEventID)
'    GoTo retry
'Here:
'#Else
    Dim oDB As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    Randomize Timer
retry:
    GetNextEventID = (2 ^ 31) * Rnd() + 1
    oDB.ConnectionString = GetConnectionStr
    szSql = "SELECT * FROM active_user_info WHERE login_id=" & GetNextEventID
    Set rsTemp = oDB.Execute(szSql)
    If rsTemp.RecordCount = 1 Then GoTo retry
'#End If

End Function

Public Sub Main()
    InitActiveUser
End Sub
'
'


Public Sub CheckActiveUser(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
'    Dim nTimeOut As Integer
'    Dim oActiveUserData As ActiveUserData
'    nTimeOut = 1
'
'    If Not m_clActiveUserData Is Nothing Then
'        For Each oActiveUserData In m_clActiveUserData
'            If DateDiff("n", oActiveUserData.m_dtLastTime, Now) > nTimeOut Then
'                SelfLogoutActiveUser oActiveUserData.m_lLoginEventID, oActiveUserData.m_lLoginEventID2, TP_TimeOutLogout
'            End If
'        Next
'    End If
End Sub

'Public Sub SetCheckActiveUserTimer(pnInterval As Integer)
'    If m_spCheckActiveUserTimer.Value <> 0 Then KillTimer 0, m_spCheckActiveUserTimer.Value
'    m_spCheckActiveUserTimer.Value = SetTimer(0, 0, pnInterval * 60 * 100, AddressOf CheckActiveUser)
'End Sub



Public Sub SelfLogoutActiveUser(ByVal plLoginEventID As Long, ByVal plLoginEvnetID2 As Long, Optional pnLogoutType As ELogoutType = TP_NormalLogout)
    Dim oDB As New RTConnection
    Dim szSql As String

    oDB.ConnectionString = GetConnectionStr
'    #If ACTIVEUSER_USE_MEMORY Then
'        m_clActiveUserData.Remove cszCLPrefix & plLoginEventID
'    #Else
        szSql = "DELETE FROM active_user_info WHERE login_id=" & plLoginEventID
        oDB.Execute szSql
'    #End If

    szSql = "UPDATE login_log_lst SET " _
    & " login_off_type =" & pnLogoutType _
    & ",login_off_time='" & ToDBDateTime(SelfNowDateTime()) & "' " _
    & " WHERE login_event_id=" & plLoginEvnetID2
    oDB.Execute szSql

End Sub




Public Sub InitActiveUser()
'    #If ACTIVEUSER_USE_MEMORY Then
'
'        '#If IN_MTS Then
'            Dim bExist As Boolean
'            Dim bExistTimer As Boolean
'            Set m_spgmActiveUserData = New SharedPropertyGroupManager
'            Set m_spgActiveUserDate = m_spgmActiveUserData.CreatePropertyGroup(cszActiveUserGroup, 0, 0, bExist)
'
'            Set m_spCheckActiveUserTimer = m_spgActiveUserDate.CreateProperty(cszCheckActiveUserTimer, bExistTimer)
'            Set m_spActiveUserData = m_spgActiveUserDate.CreateProperty(cszActiveUser, bExist)
'            'OutDebugInfo CStr(bExist)
'            If Not bExist Then
'                m_spActiveUserData.Value = New Collection
'            End If
'            If Not bExistTimer Then
'                m_spCheckActiveUserTimer.Value = 0
'                SetCheckActiveUserTimer 1
'            End If
'            Set m_clActiveUserData = m_spActiveUserData.Value
'
'    '    #Else
'    '        Set m_clActiveUserData = New Collection
'    '    #End If
'    #Else
'
'    #End If

End Sub




Public Sub AssertActiveUserValid(poActiveUser As ActiveUser, ByVal plErrBegin As Long) '测试对象的活动用户对象是否有效（不为Nothing且IsValid为真）
    If poActiveUser Is Nothing Then RaiseError ERR_NoActiveUser + plErrBegin
    poActiveUser.AssertActiveUserValid
End Sub

'判断指定的操作员是否有相应的权限
Public Sub AssertHaveRight(poActiveUser As ActiveUser, ByVal plProgramRightID As Long)
#If IN_DEBUG = 0 Then
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim oDB As New RTConnection
    Dim szRight As String
    
    szRight = GetRightID(plProgramRightID)
    oDB.ConnectionString = GetConnectionStr
    '查询用户方法表中用户是否有该权限
    szSql = "SELECT function_id FROM " & cszUserFunction & " WHERE " _
    & " user_id='" & poActiveUser.UserID & "' AND " _
    & " function_id='" & szRight & "'"
    
    Set rsTemp = oDB.Execute(szSql)
    If rsTemp.RecordCount = 0 Then
        '如果用户没有权限则查询用户组是否有权限
        szSql = "SELECT function_id FROM " & cszUsergroupFunction & "  tbu, " & cszGroupUser & " tbg WHERE tbu.usergroup_id=" _
        & " tbg.usergroup_id AND user_id='" & poActiveUser.UserID & "' AND function_id='" & szRight & "'"
        Set rsTemp = oDB.Execute(szSql)
        If rsTemp.RecordCount = 0 Then
            Dim szRightName As String
            szRightName = GetRightName(plProgramRightID)
            err.Raise plProgramRightID, , "用户" & poActiveUser.UserID & "无" & szRightName & "的权限!!!"
        End If
    End If
#End If

End Sub

'得到组件功能的代码(真正的权限代码字符串的,存在于数据库中)
Public Function GetRightID(ByVal plProgramRightID As Long) As String
    GetRightID = LoadResString(plProgramRightID + RD_RightID)
End Function

''得到组件功能的功能号
'Public Function GetRightLongID(pnCOMID As ECOMID, ByVal pnInnerRightID As Integer) As Long
'    GetRightLongID = cnMidErrBegin + pnCOMID * cnMidErrStep + pnInnerRightID * cnMidRightStep + RD_RightID
'End Function


'得到组件功能的名字
Public Function GetRightName(ByVal plProgramRightID As Long) As String
    GetRightName = LoadResString(plProgramRightID + RD_RightName)
End Function

'得到组件功能的功能组
Public Function GetRightGroup(ByVal plProgramRightID As Long) As String
    GetRightGroup = LoadResString(plProgramRightID + RD_RightGroup)
End Function

'得到组件功能的是否写日志
Public Function GetRightWriteLog(ByVal plProgramRightID As Long) As Boolean
    Dim szTemp As String
    szTemp = LoadResString(plProgramRightID + RD_RightID)
    GetRightWriteLog = IIf(Trim(szTemp) = "1", True, False)
End Function


'定操作日志
Public Function WriteOperateLog(poAcitveUser As ActiveUser, ByVal plProgramRightID As Long, Optional pszAddInfo As String = "") As Boolean
#If IN_DEBUG = 0 Then
    Dim szRight As String, szRightGroup As String
    Dim szSql As String, oDB As New RTConnection
    
    szRight = GetRightID(plProgramRightID)
    szRightGroup = GetRightGroup(plProgramRightID)
    
    szSql = "INSERT operation_log_lst(" _
    & "user_id," _
    & "function_group_id," _
    & "function_id," _
    & "operation_time," _
    & "annotation) " _
    & " VALUES( '" _
    & poAcitveUser.UserID & "','" _
    & szRightGroup & "','" _
    & szRight & "','" _
    & ToDBDateTime(SelfNowDateTime()) & "','" _
    & GetUnicodeBySize(pszAddInfo, 255) & "')"
    
    oDB.ConnectionString = GetConnectionStr
    On Error GoTo here
    oDB.Execute szSql
    WriteOperateLog = True
    Exit Function
here:
    WriteOperateLog = False
#End If
End Function


Public Function BusProjectExecutePrice(tdDate As Date, ByRef plErrCode As Long) As String
    Dim oDB As New RTConnection
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim nTemp As Long
    oDB.ConnectionString = GetConnectionStr
    On Error GoTo here
    tdDate = ToDBDate(tdDate)
    If VBDateIsEmpty(tdDate) = True Then GoTo here
    '按时间查取执行票价表
'    If pszProjectID = "" Then
'    szSql = "SELECT price_table_id FROM price_table_info WHERE " _
'                            & "start_run_time=(SELECT MAX(start_run_time) FROM price_table_info WHERE " _
'                            & "  convert(datetime,convert(char(10),start_run_time,101))<='" & ToDBDate(tdDate) & "' )"
'    Else
'        szSql = "SELECT price_table_id FROM price_table_info WHERE project_id='" & pszProjectID & "' AND " _
'                            & "start_run_time=(SELECT MAX(start_run_time) FROM price_table_info WHERE " _
'                            & " project_id='" & pszProjectID & "' AND convert(datetime,convert(char(10),start_run_time,101))<='" & ToDBDate(tdDate) & "' )"
'    End If
    szSql = "SELECT price_table_id FROM price_table_info WHERE " _
        & " start_run_time=(SELECT MAX(start_run_time) FROM price_table_info WHERE " _
        & " convert(datetime,convert(char(10),start_run_time,101))<='" & ToDBDate(tdDate) & "' )"
'

    Set rsTemp = oDB.Execute(szSql)
    If rsTemp.RecordCount = 0 Then 'ShowError ERR_NoRoutePriceTable
here:
       plErrCode = 1
    Else
       plErrCode = 0
       BusProjectExecutePrice = FormatDbValue(rsTemp!price_table_id)
    End If
    
End Function



'内部用得到本单位的代码
Public Function GetUnitID() As String
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim oDB As New RTConnection
    
    oDB.ConnectionString = GetConnectionStr
'    '=========================================================================
'    'RTStation 数据库
'    '-------------------------------------------------------------------------
'    szSql = "SELECT * FROM System_param_info WHERE parameter_name='" & cszLocalUnitID & "'"
'    '=========================================================================
    '=========================================================================
    '嘉兴数据库
    '-------------------------------------------------------------------------
    szSql = "SELECT * FROM " & cszTableSystemParam & " WHERE parameter_name='" & cszLocalUnitID & "'"
    '=========================================================================
    
    Set rsTemp = oDB.Execute(szSql)
    If rsTemp.RecordCount = 1 Then
        GetUnitID = FormatDbValue(rsTemp!parameter_value)
    End If
End Function

