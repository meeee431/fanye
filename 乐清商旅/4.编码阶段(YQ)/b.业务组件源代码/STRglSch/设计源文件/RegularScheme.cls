VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RegularScheme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'**********************************************************
'* Source File Name:RegularScheme.cls
'* Project Name:STRglSch.vbp
'* Engineer:
'* Data Generated:
'* Last Revision Date:
'* Brief Description:常规调度
'* Relational Document:
'**********************************************************

Option Explicit

'出错
Public Enum EErrRegularScheme
    ERR_RegularSchemeNotExist = ERR_RegularScheme + 21 '18421常规调度不存在
    ERR_DateNotProject = ERR_RegularScheme + 22 '18422该日期无执行计划
    ERR_ProjectNotEndDate = ERR_RegularScheme + 23 '18423计划无终止日期
    ERR_NotProject = ERR_RegularScheme + 24 '18424无计划
    ERR_ProjectNotExist = ERR_RegularScheme + 25 '18425计划不存在
    ERR_NotProjectExecute = ERR_RegularScheme + 26 '18426没有计划在执行
    ERR_RegularSchemeProjectExistInSchedule = ERR_RegularScheme + 27 '18427调度中该序号已有指定的计划，不能再新增序号的计划
    ERR_BusProjectNotExistInSchemSystem = ERR_RegularScheme + 28 '18428车次计划不在执行表中
    ERR_ModifyProjectNotNull = ERR_RegularScheme + 29  '18429编辑的车次计划不能为空
    ERR_ModifyProjectBeginDateNotNull = ERR_RegularScheme + 30 '18430编辑车次计划的开始日期
    ERR_ModefyProjectNotExecute = ERR_RegularScheme + 31 '18431车次计划不执行
    ERR_ProjectBeginDate = ERR_RegularScheme + 32 '18432车次计划开始日期不能小于当日
    ERR_ExecuteProjectSerialNoDelete = ERR_RegularScheme + 33 '18433正在执行的计划不能删除
    ERR_ExecuteProjectNoModify = ERR_RegularScheme + 34 '18434正在执行的计划不能修改开始执行日期
    ERR_ExecutePriceTable = ERR_RegularScheme + 35 '
End Enum

'权限
Public Enum ERightRegularScheme
     RIGHT_RegularSchemeManagement = ERR_RegularScheme + cnMidRightBegin + cnMidRightStep * 1 '18507常规调度
End Enum

Public Type TSchemeArrangementEx
        szProjectID  As String
        szProjectName As String
        szPriceTable As String
        eProjectStatus As EBusProjectStatus
        dtBeginDate  As Date
        szAnnotation As String
        nSerialNo As Integer
End Type

Private m_oActiveUser As ActiveUser 'P1所有的类都有的私有类变量

'**************************************************
'Member Code:F1
'Brief Description:通过日期获得该日期执行的计划
'Engineer:
'Date Generated:
'Last Revision Date:
'**************************************************
Public Function GetExecuteBusProject(ExeDate As Date) As TSchemeArrangement
'**************************************************
'ExeDate(通过该日期获得运行计划)
'**************************************************
'Dim ErrorNo As Long
'Dim tScheme As TSchemeArrangement
'tScheme = SelfGetExecuteBusProject(ExeDate, ErrorNo)
'If ErrorNo = 0 Then
'GetExecuteBusProject = tScheme
'End If
End Function

'**************************************************
'Member Code:F2
'Brief Description:通过计划代码获得该计划是否执行及执行的信息
'Engineer:
'Date Generated:
'Last Revision Date:
'**************************************************
'Public Function GetBusProject(ProjectID As String) As TSchemeArrangement()
''**************************************************
''ProjectID(计划代码获得执行计划的信息)
''**************************************************
'Dim oDb As New RTConnection
'Dim szSql As String
'Dim tScheme() As TSchemeArrangement
'Dim i As Integer
'Dim rsTemp As Recordset
'AssertObjIsValid
'    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
'    szSql = " SELECT * FROM project_execute_lst WHERE project_id='" & ProjectID & "'"
'    Set rsTemp = oDb.Execute(szSql)
'    If rsTemp.RecordCount <> 0 Then
'    ReDim tScheme(1 To rsTemp.RecordCount) As TSchemeArrangement
'    For i = 1 To rsTemp.RecordCount
'        tScheme(i).szProjectID = FormatDbValue(rsTemp!project_id)
'        tScheme(i).dtBeginDate = FormatDbValue(rsTemp!start_time)
'        tScheme(i).szAnnotation = FormatDbValue(rsTemp!Annotation)
'        tScheme(i).nSerialNo = FormatDbValue(rsTemp!serial_no)
'        rsTemp.MoveNext
'    Next
'    GetBusProject = tScheme
'    End If
'End Function

'**************************************************
'Member Code:F3
'Brief Description:获得系统的所有执行的车次计划的执行信息
'Engineer:
'Date Generated:
'Last Revision Date:
'**************************************************
'Public Function GetAllArrangement(Optional bSerial As Boolean = True) As TSchemeArrangement()
'Dim nProjectCount As Integer
'Dim rsTemp As Recordset
'Dim oDb As New RTConnection
'Dim i As Integer
'Dim tScheme() As TSchemeArrangement
'AssertObjIsValid
'    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
'    If bSerial Then
'        Set rsTemp = oDb.Execute("SELECT * FROM project_execute_lst ORDER BY serial_no")
'    Else
'        Set rsTemp = oDb.Execute("SELECT * FROM project_execute_lst ORDER BY start_time")
'    End If
'    If rsTemp.RecordCount <> 0 Then
'    nProjectCount = rsTemp.RecordCount
'    ReDim tScheme(1 To nProjectCount) As TSchemeArrangement
'    For i = 1 To nProjectCount
'        tScheme(i).szProjectID = FormatDbValue(rsTemp!project_id)
'        tScheme(i).dtBeginDate = FormatDbValue(rsTemp!start_time)
'        tScheme(i).szAnnotation = FormatDbValue(rsTemp!Annotation)
'        tScheme(i).nSerialNo = FormatDbValue(rsTemp!serial_no)
'        rsTemp.MoveNext
'    Next
'    GetAllArrangement = tScheme
'    End If
'End Function

'Public Function GetAllArrangementEx(Optional bSerial As Boolean = True) As TSchemeArrangementEx()
'Dim nProjectCount As Integer
'Dim rsTemp As Recordset
'Dim oDb As New RTConnection
'Dim i As Integer
'Dim tScheme() As TSchemeArrangementEx
'AssertObjIsValid
'    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
'    If bSerial Then
'        Set rsTemp = oDb.Execute("SELECT tbe.*,tbb.project_name,tbb.status FROM project_execute_lst tbe,bus_project_info tbb WHERE tbe.project_id=tbb.project_id ORDER BY serial_no")
'    Else
'        Set rsTemp = oDb.Execute("SELECT tbe.*,tbb.project_name,tbb.status FROM project_execute_lst tbe,bus_project_info tbb WHERE tbe.project_id=tbb.project_id ORDER BY start_time")
'    End If
'    If rsTemp.RecordCount <> 0 Then
'    nProjectCount = rsTemp.RecordCount
'    ReDim tScheme(1 To nProjectCount) As TSchemeArrangementEx
'    For i = 1 To nProjectCount
'        tScheme(i).szProjectID = rsTemp!project_id
'        tScheme(i).dtBeginDate = rsTemp!start_time
'        tScheme(i).szAnnotation = FormatDbValue(rsTemp!Annotation)
'        tScheme(i).nSerialNo = rsTemp!serial_no
'        tScheme(i).eProjectStatus = rsTemp!Status
'        tScheme(i).szProjectName = rsTemp!project_name
'        rsTemp.MoveNext
'    Next
'    GetAllArrangementEx = tScheme
'    End If
'End Function
'**************************************************
'Member Code:F4
'Brief Description:获得系统的所有车次计划
'Engineer:
'Date Generated:
'Last Revision Date:
''**************************************************
'Public Function GetAllBusProject() As String()
'    Dim szaTemp() As String
'    Dim oDb As New RTConnection
'    Dim i As Integer
'    Dim rsTemp As Recordset
'AssertObjIsValid
'    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
'    Set rsTemp = oDb.Execute("SELECT project_id FROM bus_project_info ORDER BY project_id")
'    If rsTemp.RecordCount <> 0 Then
'    ReDim szaTemp(1 To rsTemp.RecordCount) As String
'    For i = 1 To rsTemp.RecordCount
'        szaTemp(i) = FormatDbValue(rsTemp!project_id)
'        rsTemp.MoveNext
'    Next
'    GetAllBusProject = szaTemp
'    End If
'End Function

'**************************************************
'Member Code:F4
'Brief Description:获得系统的所有车次计划
'Engineer:
'Date Generated:
'Last Revision Date:
'**************************************************
'Public Function GetAllBusProjectEx() As String()
'    Dim szaTemp() As String
'    Dim oDb As New RTConnection
'    Dim i As Integer
'    Dim OPara As New SystemParam
'    Dim szSql As String
'    Dim lErrCode As Long
'    Dim m_szExecutePriceTable As String
'    Dim rsTemp As Recordset
'AssertObjIsValid
'    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
'    OPara.Init m_oActiveUser
'    m_szExecutePriceTable = BusProjectExecutePrice("", OPara.NowDate, lErrCode)
'    If lErrCode = 1 Then ShowError ERR_ExecutePriceTableNotExist
'
'    szSql = "SELECT tbe.*,tbp.price_table_name FROM bus_project_info tbe, " _
'            & " price_table_info tbp WHERE tbe.project_id*=tbp.project_id " _
'            & " AND tbp.price_table_id ='" & m_szExecutePriceTable & "'"
'    Set rsTemp = oDb.Execute(szSql)
'
'    If rsTemp.RecordCount <> 0 Then
'    ReDim szaTemp(1 To rsTemp.RecordCount, 1 To 5) As String
'    For i = 1 To rsTemp.RecordCount
'        szaTemp(i, 1) = rsTemp!project_id
'        szaTemp(i, 2) = FormatDbValue(rsTemp!project_name)
'        szaTemp(i, 3) = rsTemp!Status
'        szaTemp(i, 4) = FormatDbValue(rsTemp!price_table_name)
'        szaTemp(i, 5) = FormatDbValue(rsTemp!Annotation)
'        rsTemp.MoveNext
'    Next
'    GetAllBusProjectEx = szaTemp
'    End If
'End Function
'**************************************************
'Member Code:F5
'Brief Description:获得某执行计划的结束日期
'Engineer:
'Date Generated:
'Last Revision Date:
'**************************************************
'Public Function GetBusProjectEndDate(ProjectID As String, ProjectSerial As Integer) As Date
''**************************************************
''ProjectID(计划代码)
''**************************************************
'    Dim oDb As New RTConnection
'    Dim szSql As String
'    Dim tScheme As TSchemeArrangement
'    Dim rsTemp As Recordset
'    AssertObjIsValid
'    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
'    szSql = " SELECT * FROM project_execute_lst WHERE start_time>(SELECT start_time FROM project_execute_lst WHERE project_id='" & ProjectID & "' AND serial_no =" & ProjectSerial & ") ORDER BY start_time"
'    Set rsTemp = oDb.Execute(szSql)
'    If rsTemp.RecordCount <> 0 Then
'        GetBusProjectEndDate = DateAdd("d", -1, FormatDbValue(rsTemp!start_time))
'    End If
'End Function

'**************************************************
'Member Code:F6
'Brief Description:获得最大的序号
'Engineer:
'Date Generated:
'Last Revision Date:
'**************************************************
'Public Function GetMaxSerial() As Integer
'    Dim oDb As New RTConnection
'    Dim szSql As String
'    Dim rsTemp As Recordset
'AssertObjIsValid
'    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
'    szSql = " SELECT MAX(serial_no) AS maxserial FROM project_execute_lst"
'    Set rsTemp = oDb.Execute(szSql)
'    If rsTemp.RecordCount = 0 Then
'        GetMaxSerial = 0
'    Else
'        GetMaxSerial = rsTemp!maxserial
'    End If
'End Function


'**************************************************
'Member Code:F8
'Brief Description:获得某执行计划的开始日期和结束日期
'Engineer:
'Date Generated:
'Last Revision Date:
'**************************************************
'Public Function GetBusProjectDate(ProjectID As String, ProjectSerial As Integer) As Date()
''**************************************************
''ProjectID(计划代码)
''**************************************************
'    Dim oDb As New RTConnection
'    Dim szSql As String
'    Dim tScheme As TSchemeArrangement
'    Dim dtProjectDate(1 To 2) As Date
'    Dim rsTemp As Recordset
'
'    AssertObjIsValid
'    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
'
'    szSql = " SELECT * FROM project_execute_lst WHERE project_id='" & ProjectID & "' AND serial_no =" & ProjectSerial
'    Set rsTemp = oDb.Execute(szSql)
'    If rsTemp.RecordCount <> 0 Then
'        dtProjectDate(1) = rsTemp!start_time
'    Else
'        dtProjectDate(1) = CDate(cszEmptyDateStr)
'    End If
'
'    szSql = " SELECT * FROM project_execute_lst WHERE start_time>(SELECT start_time FROM project_execute_lst WHERE project_id='" & ProjectID & "' AND serial_no =" & ProjectSerial & ") ORDER BY start_time"
'    Set rsTemp = oDb.Execute(szSql)
'    If rsTemp.RecordCount <> 0 Then
'        dtProjectDate(2) = DateAdd("d", -1, FormatDbValue(rsTemp!start_time))
'        Else
'        dtProjectDate(2) = CDate(cszEmptyDateStr)
'    End If
'    GetBusProjectDate = dtProjectDate
'End Function

'**************************************************
'Member Code:P1
'Brief Description:活动用户
'**************************************************
Public Property Get SelfUser() As ActiveUser
    Set SelfUser = m_oActiveUser
End Property
Public Property Set SelfUser(vNewValue As ActiveUser)
    Set m_oActiveUser = vNewValue
End Property

'**************************************************
'Member Code:S1
'Brief Description:初始化对象
'Engineer:
'Date Generated:
'Last Revision Date:
'**************************************************
Public Sub Init(poAUser As ActiveUser)
'**************************************************
'poAuser(活动的用户)
'**************************************************
    Set m_oActiveUser = poAUser
    'AssertHaveRight m_oActiveUser, RIGHT_RegularSchemeManagement
End Sub

'**************************************************
'Member Code:S2
'Brief Description:测试对象是是否有效（活动用户对象有效且对象的状态的状态有效）
'Engineer:
'Date Generated:
'Last Revision Date:
'**************************************************
Private Sub AssertObjIsValid()
    AssertActiveUserValid m_oActiveUser, ERR_RegularScheme
End Sub

'**************************************************
'Member Code:S3
'Brief Description:安排车次计划去执行
'Engineer:
'Date Generated:
'Last Revision Date:
'**************************************************
'Public Sub AddBusProjectToSchedule(SchemeMan As TSchemeArrangement)
''**************************************************
''SchemeMan(计划管理结构)
''**************************************************
'    Dim oDb As New RTConnection
'    Dim rsTemp As Recordset
'    Dim szTempSql As String
'    Dim szSqlField As String
'    Dim szSqlContext As String
'AssertObjIsValid
'    AssertHaveRight m_oActiveUser, RIGHT_RegularSchemeManagement
'    If VBDateIsEmpty(SchemeMan.dtBeginDate) Then ShowError ERR_ModifyProjectBeginDateNotNull
'    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
'    If DateDiff("d", SchemeMan.dtBeginDate, Now) >= 0 Then ShowError ERR_ProjectBeginDate
'    '1查询该计划是否存在
'    szTempSql = "SELECT * FROM bus_project_info  WHERE project_id='" & SchemeMan.szProjectID & "' and status='" & ST_BusProjectNormal & "'"
'    Set rsTemp = oDb.Execute(szTempSql)
'    If rsTemp.RecordCount = 0 Then ShowError ERR_BusProjectNotExistInSchemSystem '调度系统中无该计划
'
'    '2查询计划中是否有该序号
'    szTempSql = "SELECT * FROM project_execute_lst WHERE serial_no=" & Str(SchemeMan.nSerialNo) & " OR start_time='" & ToDBDate(SchemeMan.dtBeginDate) & "'"
'    Set rsTemp = oDb.Execute(szTempSql)
'    If rsTemp.RecordCount <> 0 Then ShowError ERR_RegularSchemeProjectExistInSchedule '调度中该序号已有指定的计划，不能再新增序号的计划
'
'    '在计划中新增车次
'    szTempSql = "INSERT project_execute_lst ("
'    szSqlField = szSqlField & "serial_no,"
'    szSqlField = szSqlField & "project_id,"
'    szSqlField = szSqlField & "start_time,"
'    szSqlField = szSqlField & "annotation)"
'
'    szSqlContext = " VALUES (" & Str(SchemeMan.nSerialNo) & ","
'    szSqlContext = szSqlContext & " '" & SchemeMan.szProjectID & "',"
'    szSqlContext = szSqlContext & " '" & ToDBDate(SchemeMan.dtBeginDate) & "',"
'    szSqlContext = szSqlContext & " '" & SchemeMan.szAnnotation & "')"
'
'    szTempSql = szTempSql & szSqlField & szSqlContext
'    oDb.Execute (szTempSql)
'    WriteOperateLog m_oActiveUser, RIGHT_RegularSchemeManagement, "安排车次计划[" & SchemeMan.szProjectID & "]去执行"
'End Sub

'**************************************************
'Member Code:S4
'Brief Description:修改执行的车次计划
'Engineer:
'Date Generated:
'Last Revision Date:
'**************************************************
'Public Sub ModifyBusProjectArrangement(SchemeMan As TSchemeArrangement)
''**************************************************
''SchemeMan(计划管理结构)
''**************************************************
'    Dim oDb As New RTConnection
'    Dim szTempSql As String
'    Dim rsTemp As Recordset
'    Dim szSqlField As String
'    Dim szSqlWhere As String
'    '修改计划基本信息
'AssertObjIsValid
'     AssertHaveRight m_oActiveUser, RIGHT_BusProjectManagement
'    '-----------------------------
'    '车次计划不能为空
'    If SchemeMan.szProjectID = "" Then ShowError ERR_ModifyProjectNotNull
'    '-----------------------------
'    '正在执行的计划不能修改开始日期
'    If Me.GetExecuteBusProject(Now).nSerialNo = SchemeMan.nSerialNo Then ShowError ERR_ExecuteProjectNoModify
'    '车次无开始日期，不行
'    If VBDateIsEmpty(SchemeMan.dtBeginDate) Then ShowError ERR_ModifyProjectBeginDateNotNull
'    '修改的开始日期小于当日不能修改
'    If DateDiff("d", SchemeMan.dtBeginDate, Now) > 0 Then ShowError ERR_ProjectBeginDate
'    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
'
'    szTempSql = "SELECT * FROM project_execute_lst WHERE serial_no=" & SchemeMan.nSerialNo
'    Set rsTemp = oDb.Execute(szTempSql)
'    '-----------------------------
'    '该序号没有执行
'    If rsTemp.RecordCount = 0 Then ShowError ERR_ModefyProjectNotExecute
'    szTempSql = "UPDATE project_execute_lst SET "
'    szSqlField = "project_id='" & SchemeMan.szProjectID & "',"
'    szSqlField = szSqlField & "start_time='" & ToDBDate(SchemeMan.dtBeginDate) & "',"
'    szSqlField = szSqlField & "annotation='" & SchemeMan.szAnnotation & "'"
'    szSqlWhere = " WHERE serial_no=" & SchemeMan.nSerialNo
'    szTempSql = szTempSql & szSqlField & szSqlWhere
'    oDb.Execute szTempSql
'    WriteOperateLog m_oActiveUser, RIGHT_BusProjectManagement, "修改执行的车次计划" & SchemeMan.szProjectID
'End Sub


'**************************************************
'Member Code:S5
'Brief Description:删除执行的车次计划中计划表
'Engineer:
'Date Generated:
'Last Revision Date:
'**************************************************
'Public Sub DeleteBusProjectFromSchedule(ProjectSerial As Integer)
''**************************************************
''ProjectSerial(删除的执行计划序号)
''**************************************************
'    Dim oDb As New RTConnection
'    Dim rsTemp As Recordset
'    Dim szTempSql As String
'AssertObjIsValid
'     AssertHaveRight m_oActiveUser, RIGHT_RegularSchemeManagement
'    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
'    If Me.GetExecuteBusProject(Now).nSerialNo = ProjectSerial Then ShowError ERR_ExecuteProjectSerialNoDelete
'    szTempSql = "SELECT * FROM project_execute_lst WHERE serial_no=" & ProjectSerial
'    Set rsTemp = oDb.Execute(szTempSql)
'    '-----------------------------
'    '该序号没有执行
'    If rsTemp.RecordCount = 0 Then ShowError ERR_ModefyProjectNotExecute
'
'    szTempSql = "DELETE project_execute_lst WHERE serial_no=" & Str(ProjectSerial)
'    oDb.Execute szTempSql
'    WriteOperateLog m_oActiveUser, RIGHT_RegularSchemeManagement, "删除执行的车次计划中计划表,序号为" & ProjectSerial
'End Sub

'获得车次票价项
Public Function GetBusPrice(Item As String) As String()
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim szPrice(1 To 2) As String
    Dim szTempSql As String
AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    szTempSql = "SELECT * FROM price_item_info WHERE price_item='" & Item & "'"
    Set rsTemp = oDb.Execute(szTempSql)
    '-----------------------------
    '该序号没有执行
    If rsTemp.RecordCount <> 0 Then
        szPrice(1) = rsTemp!chinese_name
        szPrice(2) = rsTemp!use_mark
    Else
        szPrice(1) = ""
        szPrice(2) = "0"
    End If
    GetBusPrice = szPrice
End Function

'Public Function ProjectExistExeTable(ProjectID As String) As Integer
'    Dim oDb As New RTConnection
'    Dim szTempSql As String
'    AssertObjIsValid
'    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
'    szTempSql = "SELECT COUNT(*) AS execount FROM project_execute_lst WHERE project_id='" & Trim(ProjectID) & "'"
'    ProjectExistExeTable = oDb.Execute(szTempSql)!execount
'End Function

Public Function ProjectExistTable() As String()
    Dim oDb As New RTConnection
    Dim rsPrice As Recordset
    Dim szTempSql As String
    Dim szTemp() As String
    Dim i As Integer

    AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    szTempSql = "SELECT *  FROM price_table_info " 'WHERE "project_id='" & Trim(ProjectID) & "'"
    Set rsPrice = oDb.Execute(szTempSql)
    ReDim szTemp(1 To rsPrice.RecordCount, 6)
    Do While Not rsPrice.EOF
       i = i + 1
'       szTemp(i, 1) = rsPrice!project_id
       szTemp(i, 2) = rsPrice!price_table_id
       szTemp(i, 3) = rsPrice!price_table_name
       szTemp(i, 4) = rsPrice!create_time
       szTemp(i, 5) = rsPrice!last_modify_time
       szTemp(i, 6) = rsPrice!start_run_time
       rsPrice.MoveNext
    Loop
    ProjectExistTable = szTemp
    Set rsPrice = Nothing
    Set oDb = Nothing
End Function

'Public Function GetRunPriceTable() As String()
'Dim oDb As New RTConnection
'Dim rsTemp As Recordset
'Dim szExePriceTable As String
'Dim szSql As String
'Dim lErrCode As Long
'Dim OPara As New SystemParam
'Dim szTemp(1 To 1, 1 To 6) As String
'Dim TTemp As TSchemeArrangement
'
'    AssertObjIsValid
'    OPara.Init m_oActiveUser
'    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
'    'szExePriceTable = BusProjectExecutePrice(m_szProjectID, OPara.NowDate, lErrCode)
'    TTemp = GetExecuteBusProject(OPara.NowDate)
'
'    szExePriceTable = BusProjectExecutePrice(TTemp.szProjectID, OPara.NowDate, lErrCode)
'    If lErrCode = 1 Then ShowError ERR_ExecutePriceTable
'    szSql = "SELECT pfo.* FROM project_execute_lst pst,price_table_info pfo " _
'           & "  where pst.project_id=pfo.project_id " _
'           & "  and  pfo.price_table_id ='" & szExePriceTable & "' "  ' status=0"
'    Set rsTemp = oDb.Execute(szSql)
'    If rsTemp.RecordCount = 1 Then
'        szTemp(1, 1) = rsTemp!project_id
'        szTemp(1, 2) = rsTemp!price_table_id
'        szTemp(1, 3) = rsTemp!price_table_name
'        szTemp(1, 4) = rsTemp!create_time
'        szTemp(1, 5) = rsTemp!last_modify_time
'        '表示运行 --原status
'        szTemp(1, 6) = rsTemp!start_run_time
'    End If
'    GetRunPriceTable = szTemp
'    Set rsTemp = Nothing
'    Set oDb = Nothing
'End Function
Public Function GetRunPriceTableEx(dtRundate As Date) As String
Dim szExePriceTable As String
Dim szSql As String
Dim lErrCode As Long
    
    szExePriceTable = BusProjectExecutePrice(dtRundate, lErrCode)
    If lErrCode = 1 Then ShowError ERR_ExecutePriceTable
    GetRunPriceTableEx = szExePriceTable
End Function
Public Function GetPriceTableInfo(dtRundate As Date) As String()
Dim oDb As New RTConnection
Dim rsTemp As Recordset
Dim szExePriceTable As String
Dim szSql As String
Dim lErrCode As Long
Dim szTemp(1 To 1, 1 To 6) As String
'Dim TTemp As TSchemeArrangement

    AssertObjIsValid
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    'szExePriceTable = BusProjectExecutePrice(m_szProjectID, OPara.NowDate, lErrCode)
   ' TTemp = GetExecuteBusProject(dtRundate)
    
    szExePriceTable = BusProjectExecutePrice(dtRundate, lErrCode)
    If lErrCode = 1 Then ShowError ERR_ExecutePriceTable
    szSql = "SELECT pfo.* FROM price_table_info pfo " _
           & "  WHERE pfo.price_table_id ='" & szExePriceTable & "' "  ' status=0"
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 1 Then
        szTemp(1, 2) = rsTemp!price_table_id
        szTemp(1, 3) = rsTemp!price_table_name
        szTemp(1, 4) = rsTemp!create_time
        szTemp(1, 5) = rsTemp!last_modify_time
        '表示运行 --原status
        szTemp(1, 6) = rsTemp!start_run_time
    End If
    GetPriceTableInfo = szTemp
    Set rsTemp = Nothing
    Set oDb = Nothing
End Function


'得到在指定车次计划中运行在此线路上的所有车型.。
Public Function GetAllVehicleModel(ByVal pszRouteID As String) As String()
    Dim oDb As New RTConnection
    Dim aszVehicleModel() As String
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim i As Long
    szSql = "SELECT DISTINCT vehicle_type_code.vehicle_type_code,vehicle_type_short_name " _
    & "FROM bus_info, bus_vehicle_code, Vehicle_info, vehicle_type_code " _
    & "WHERE route_id='" & pszRouteID & "' AND " _
    & "bus_vehicle_code.vehicle_id=Vehicle_info.vehicle_id AND " _
    & "Vehicle_info.vehicle_type_code = vehicle_type_code.vehicle_type_code "
    
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount > 0 Then
        ReDim aszVehicleModel(1 To rsTemp.RecordCount, 1 To 2)
        For i = 1 To rsTemp.RecordCount
            aszVehicleModel(i, 1) = FormatDbValue(rsTemp!vehicle_type_code)
            aszVehicleModel(i, 2) = FormatDbValue(rsTemp!vehicle_type_short_name)
            rsTemp.MoveNext
        Next
    End If
    GetAllVehicleModel = aszVehicleModel
    Set rsTemp = Nothing
    Set oDb = Nothing
End Function

