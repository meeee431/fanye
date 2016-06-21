VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CheckTicket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'检票

Option Explicit

'错误声明
Public Enum EErrChkTk
    ERR_ChkTkAddDuplicateObj = cnErroeStartNo + ERR_AddDuplicate
    ERR_ChkTkAddNoParentObj = cnErroeStartNo + ERR_AddNoParent
    ERR_ChkTkEditChildExistObj = cnErroeStartNo + ERR_EditChildExist
    ERR_ChkTkEditNoParentObj = cnErroeStartNo + ERR_EditNoParent
    ERR_ChkTkDeleteChildExistObj = cnErroeStartNo + ERR_DeleteChildExist
    ERR_ChkTkDeleteNotExistObj = cnErroeStartNo + ERR_DeleteNotExist
    ERR_ChkTkNormalObj = cnErroeStartNo + ERR_DeleteNotExist + 2
    
    ERR_ChkTkCheckGateIDNotAppointed = cnErroeStartNo + 20                   '未指定检票口
    ERR_ChkTkTicketStatusError = cnErroeStartNo + 21                         '票的状态错误
    ERR_ChkTkTicketStartTimeIsOver = cnErroeStartNo + 22                     '此票的车次的发车时间已过（此票已自动作废）
    
    ERR_ChkTkCheckTimeIsNotStarting = cnErroeStartNo + 23                    '未到开检时间
    ERR_ChkTkCheckTimeIsOver = cnErroeStartNo + 24                           '已过正常检票时间
    ERR_ChkTkBusAlreadyExist = cnErroeStartNo + 25                           '指定车次指定序号的流水车次已存在
    ERR_ChkTkExtraCheckTimeIsNotStarting = cnErroeStartNo + 26               '未到补检时间
    ERR_ChkTkBusNotExist = cnErroeStartNo + 27                               '指定的车次不存在
    ERR_ChkTkBusIsNotFlowBus = cnErroeStartNo + 28                           '指定的车次不是流水车次
    ERR_ChkTkFlowBusNotExist = cnErroeStartNo + 29                           '指定的流水车次不存在
    
    ERR_ChkTkBusIsNotChecked = cnErroeStartNo + 30                           '指定的车次还未检票
    ERR_ChkTkCheckSheetIDAlreadyExist = cnErroeStartNo + 31                  '指定的路单存在
    ERR_ChkTkCheckSheetIDNotExist = cnErroeStartNo + 32                      '指定的路单不存在
    ERR_ChkTkCheckSheetIDIsNotValid = cnErroeStartNo + 33                    '指定的路单已经是作废了的
    ERR_ChkTkBusNotPassStation = cnErroeStartNo + 34                         '此车票的到站当指定车次不过（实际是不过或不可售）
    ERR_ChkTkBusNoSelect = cnErroeStartNo + 35                               '没有车次被选择
    ERR_ChkTkCheckGateIDNotValid = cnErroeStartNo + 36                       '无此检票口
    ERR_ChkTkResultIsNull = cnErroeStartNo + 37                              '无记录
    ERR_ChkTkTicketNotChecked = cnErroeStartNo + 38                          '不能检该票
    ERR_ChkTkCheckGateIsNotRight = cnErroeStartNo + 39                       '指定的车次不在当前检票口检票
    ERR_ChkTkExtraCheckTimeAlreadyEnd = cnErroeStartNo + 40                  '已超过补检时间
    ERR_ChkTkCheckSheetNoIsNull = cnErroeStartNo + 41                        '路单号为空
    ERR_ChkTkVehicleNotExist = cnErroeStartNo + 42                           '指定的车辆不存在
    
    ERR_TicketNoExist = cnErroeStartNo + 43                                  '车票不存在
    ERR_ChkTkUserNotCheckGate = cnErroeStartNo + 44                          '检票用户不是该检票口
    ERR_CheckSheetMakeOk = cnErroeStartNo + 45                              '路单已形成不能注销检票
End Enum

'权限声明
Public Enum ERightChkTk
    RIGHT_ChkTkCheckTicket = cnRightStartNo + cnMidRightBegin + cnMidRightStep * 0                    '检票
    RIGHT_ChkTkCheckBus = cnRightStartNo + cnMidRightBegin + cnMidRightStep * 1                       '开检车次
    RIGHT_ChkTkWriteOffCheckTicket = cnRightStartNo + cnMidRightBegin + cnMidRightStep * 2            '注销检票
    RIGHT_ChkTkMakeCheckSheet = cnRightStartNo + cnMidRightBegin + cnMidRightStep * 3                 '生成路单或更改路单编号
    RIGHT_ChkTkExtraCheckBus = cnRightStartNo + cnMidRightBegin + cnMidRightStep * 4                  '补检固定车次
    RIGHT_ChkTkCancelCheckSheet = cnRightStartNo + cnMidRightBegin + cnMidRightStep * 5               '作废路单
    RIGHT_ChkTkStopCheckBus = cnRightStartNo + cnMidRightBegin + cnMidRightStep * 6                   '停检车次
    RIGHT_ChkTkChangeCheckBus = cnRightStartNo + cnMidRightBegin + cnMidRightStep * 7                 '改乘车次
    RIGHT_ChkTkChangeBusBeforeCheet = cnRightStartNo + cnMidRightBegin + cnMidRightStep * 8           '路单前换车
    RIGHT_ChkTkChangeBusAfterCheet = cnRightStartNo + cnMidRightBegin + cnMidRightStep * 9            '路单后换车
    RIGHT_ChkTkExPrintCheckSheet = cnRightStartNo + cnMidRightBegin + cnMidRightStep * 10             '补打路单
End Enum

'类型声明
Public Type M_TRunVehicle      '运行车次信息
    VehicleID As String
    Vehicle As String
    Owner As String
    VehicleType As String
    Company As String
    Status As Integer
End Type

Public Type tCheckBusLstInfo              '检票列表信息结构
    BusID As String     '车次号
    BusSerial As Integer '车次序号
    BusMode As Integer '车次类型
    StartUpTime As Date '发车时间
    StartChkTime As Date '开检时间
    StopChkTime As Date '停检时间
    EndStationName As String '到站
    Company As String   '参营公司
    SplitCompany As String   '参营公司
    Vehicle As String   '车辆
    Owner As String     '车主
    CheckSheet As String   '路单号码
    Status As Integer   '状态
    CheckGate As String '检票口
    CheckGateType As Integer '检票口类型
End Type



Type TGateInfo      '检票口信息结构
    GateId As String
    GateName As String
    GateDetail As String
End Type

Type TOperatorInfo   '操作员信息结构
    OperatorId As String
    OperatorName As String
End Type


'枚取声明
Public Enum EIsWritelog
    WriteLog = 0
    NotWirteLog = 1
End Enum

Public Type TInterfaceCheckTicket
' SellStationID As String
    BusID As String
    BusDate As Date
    TicketStatus As ETicketStatus
    TicketType As ETicketType
    SeatNo As String
    StationName As String
    Price As Double
 SellStationID As String
End Type

Public Type TInterfaceCheckTicketEx
    BusID As String
    BusDate As Date
    TicketStatus As ETicketStatus
    TicketType As ETicketType
    SeatNo As String
    StationName As String
    Price As Double
    
    TicketID As String
    BusStartTime As Date
    ChildrenTicket As Boolean
    SellStationID As String
    SellStationName As String
    
    CardType As String      '证件类型
    IDCardNo As String      '证件号
    PersonName As String    '姓名
    Sex As String           '性别
    National As String      '民族(本国)'国籍(国外)
    Address As String       '地址
    PersonPicture As String '证件照片
End Type

'----------------------
'各类特有的变量
Private m_szFunctionID As String
Private m_szComponentID As String
Private m_szFunctionGroup As String
Private m_szFunctionName As String
Private m_eIfWriteLog As EIsWritelog
Private m_szAnnotation As String
'---------------------
'所有的实体类都有的私有类变量<<
Private m_nObjectStatus As EObjectStatus

'所有的类都有的私有类变量
Private m_oActiveUser As ActiveUser
Private m_szUser As String

'本类具有的私有的类变量
Private m_oMsg As RTSndMsg.CSendMsg
Private m_sCheckGateNo As String
Private m_sCheckGateName As String

Private mp_szBusID As String '车次代码
Private mp_dtBusDate As Date '车次日期
Private mp_szMergeBusID As String '并班车次
Private mp_bChangeBus As Boolean '车次是否改乘
Private mp_bMergeBusStatus As Boolean   '车次是否有其他车次并入
Private m_szSellStation As String       '当前上车站

Private Const sczChangeSeatReserved = "改乘预留"


Public Sub Init(poAUser As ActiveUser)
    Set SelfUser = poAUser
    m_szUser = Trim(SelfUser.UserID)
End Sub

Public Property Get SelfUser() As ActiveUser
    Set SelfUser = m_oActiveUser
    m_oActiveUser.IAmLive
End Property

Public Property Set SelfUser(NewValue As ActiveUser)
    If m_nObjectStatus = ST_AddObj Then
        ShowError ERR_AddObj
    ElseIf m_nObjectStatus = ST_EditObj Then
        ShowError ERR_EditObj
    ElseIf m_nObjectStatus = ST_NormalObj Then
        ShowError ERR_NormalObj
    Else
        Set m_oActiveUser = NewValue
    End If
    m_oActiveUser.IAmLive
    m_szUser = Trim(SelfUser.UserID)
End Property

Public Property Let CheckGateNo(ByVal NewCheckGateNo As String)
    If NewCheckGateNo <> m_sCheckGateNo Then
        Dim oDb As New RTConnection
        Dim rsTmp As Recordset
        Dim szSql As String
        oDb.ConnectionString = GetConnectionStr(cszSystemMan)
        szSql = "SELECT c.check_gate_name , s.sell_station_id " _
            & " FROM Checkgate_info c , sell_station_info s " _
            & " WHERE c.sell_station_id *= s.sell_station_id AND  check_gate_id='" & NewCheckGateNo & "'"
        Set rsTmp = oDb.Execute(szSql)
        If rsTmp.RecordCount = 0 Then
            ShowError ERR_ChkTkCheckGateIDNotValid
        Else
            m_sCheckGateNo = NewCheckGateNo
            m_sCheckGateName = FormatDbValue(rsTmp!check_gate_name)
            m_szSellStation = FormatDbValue(rsTmp!sell_station_id)
        End If
    End If
End Property

Public Property Get CheckGateNo() As String
    CheckGateNo = m_sCheckGateNo
End Property

'得到下一个未检车次
Public Function GetNextCheckBus() As REBus
    Dim oReg As CFreeReg
    Dim szCheckGateType As String

    Dim oDb As New RTConnection
    Dim rsTemp As New Recordset
    Dim szSql As String
    Dim dtMinDate As Date
    Dim oREBus As New REBus
    Dim dCheckTime As Double
    Dim rsRecord As Recordset       '返回全部记录集
    Dim szChar As String
    Dim rsFind As Recordset
    Dim nPostion As Integer
    Dim szMinBusID As String
    
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    AssertActiveUserValid m_oActiveUser, cnErroeStartNo
    If Len(m_sCheckGateNo) = 0 Then '未指定检票口
        ShowError ERR_ChkTkCheckGateIDNotAppointed
        Exit Function
    End If
    
    '当前发车时间大于当前时间和最晚补检时间之和的第一班车次
    '    dCheckTime = GetParameterValue(cszExtraStartCheckTime)    '取最晚补检时间
    '    If dCheckTime < 0 Then dCheckTime = 0
    
    If Len(szCheckGateType) = 0 Then
'        szSql = "SELECT bus_id,bus_date FROM Work_env_bus_info " & _
'        "WHERE " & DBDateCompare(Date, "bus_date") & " AND check_gate_id='" & Trim(m_sCheckGateNo) & "' AND (status=" & Str(EREBusStatus.ST_BusNormal) & " or status=32 OR status = 16) AND bus_start_time=(SELECT MIN(bus_start_time) FROM Work_env_bus_info " & _
'        "WHERE " & DBTimeCompareEX(Now, "bus_start_time", "<") & " AND check_gate_id='" & Trim(m_sCheckGateNo) & "' AND " & DBDateCompare(Date, "bus_date") & " AND (status=" & Str(EREBusStatus.ST_BusNormal) & " or status=32 OR status = 16)) order by bus_start_time,bus_id"
        szSql = "SELECT a.bus_id,a.bus_date,b.check_gate_id FROM Work_env_bus_info a,Work_env_bus_allot_lst b " & _
        "WHERE b.can_sale_quantity <>0 AND a.bus_type<>1 AND " & DBDateCompare(Date, "a.bus_date") & " AND b.check_gate_id='" & Trim(m_sCheckGateNo) & "' AND (b.status=" & Str(EREBusStatus.ST_BusNormal) & " or b.status=32 OR b.status = 16) AND b.bus_start_time=(SELECT MIN(bus_start_time) FROM Work_env_bus_allot_lst " & _
        "WHERE  can_sale_quantity <>0 AND  " & DBTimeCompareEX(Now, "bus_start_time", "<") & " AND check_gate_id='" & Trim(m_sCheckGateNo) & "' AND " & DBDateCompare(Date, "bus_date") & " AND (status=" & Str(EREBusStatus.ST_BusNormal) & " or status=32 OR status = 16))" & _
        " AND a.bus_date=b.bus_date AND a.bus_id=b.bus_id "
        ' &         " order by bus_start_time,bus_id"
    Else
'        szSql = "SELECT * FROM Work_env_bus_info WHERE  check_gate_id='" & Trim(m_sCheckGateNo) & "' AND " & DBDateCompare(Date, "bus_date") & " order by bus_start_time,bus_id" '该检票口的车次
'        Set rsRecord = oDb.Execute(szSql)
'        Set rsFind = rsRecord
'        '得到下一班车次的发车时间
'        szSql = "SELECT MIN(bus_start_time) as bus_start_time FROM Work_env_bus_info WHERE " & DBTimeCompareEX(Now, "bus_start_time", "<") & " AND check_gate_id='" & Trim(m_sCheckGateNo) & "' AND " & DBDateCompare(Date, "bus_date") & " AND ((status=" & Str(EREBusStatus.ST_BusNormal) & ") or (status=32) OR (status = 16)) order by bus_start_time"
'        Set rsTemp = oDb.Execute(szSql)
'here1:  If rsTemp.RecordCount = 0 Then Exit Function
'        dtMinDate = FormatDbValue(rsTemp!bus_start_time)
'        '查找下一班车次在记录集中的位置
'        rsFind.Filter = "bus_start_time='" & ToDBDateTime(dtMinDate) & "' AND status=0 or status=16 "
'        If rsFind.RecordCount = 0 Then Exit Function
'        nPostion = rsFind.Bookmark
'        If rsFind.RecordCount > 1 Then '同一时间有多班车次
'            szSql = "SELECT  bus_start_time, bus_id FROM Work_env_bus_info WHERE " & DBTimeCompareEX(dtMinDate, "bus_start_time", "=") & " AND check_gate_id='" & Trim(m_sCheckGateNo) & "' AND " & DBDateCompare(Date, "bus_date") & " AND (status=" & Str(EREBusStatus.ST_BusNormal) & " or status=32 OR status = 16)  order by bus_start_time,bus_id"
'            Set rsTemp = oDb.Execute(szSql)
'            If rsTemp.RecordCount = 0 Then Exit Function '无该时间发车车次
'
'Here2:      If rsTemp.EOF Then '当前时间当前检票口的单（或双）号无车次要检票（都已检、在检或停班），必须转换到下一检票时间
'                szSql = "SELECT MIN(bus_start_time) as bus_start_time FROM Work_env_bus_info WHERE " & DBTimeCompareEX(dtMinDate, "bus_start_time", "<") & " AND check_gate_id='" & Trim(m_sCheckGateNo) & "' AND " & DBDateCompare(Date, "bus_date") & " AND (status=" & Str(EREBusStatus.ST_BusNormal) & " or status=32 OR status = 16) order by bus_start_time"
'                Set rsTemp = oDb.Execute(szSql)
'                GoTo here1
'            Else  '查找当前时间当前检票口的未检车次
'                Set rsFind = rsRecord
'                szMinBusID = FormatDbValue(rsTemp!bus_id)
'                rsFind.Filter = "bus_start_time='" & ToDBDateTime(dtMinDate) & "' AND bus_id='" & szMinBusID & "'"
'                If rsFind.RecordCount = 0 Then Exit Function
'                nPostion = rsFind.Bookmark
'                '单号检票台＼双号位置车次 或　双号检票台＼单号位置车次 继续查找
'                '否则对该车次检票，　单号位检票台＼单号位车次在单号台检，双号位检票台＼双号位车次在双号台检
'
'                If (szCheckGateType = 1 AND nPostion Mod 2 = 0) Or (szCheckGateType = 2 AND nPostion Mod 2 = 1) Then
'                    If Not rsTemp.EOF Then rsTemp.MoveNext
'                    GoTo Here2
'                End If
'            End If
'        Else
'            szMinBusID = FormatDbValue(rsFind!bus_id)
'            '单号位检票台＼双号位车次 或 双号位检票台＼单号位车次 则更换下一不同时间，
'            '否则对该车次检票，　单号位检票台＼单号位车次在单号台检，双号位检票台＼双号位车次在双号台检
'            If (szCheckGateType = 1 AND nPostion Mod 2 = 0) Or (szCheckGateType = 2 AND nPostion Mod 2 = 1) Then
'                szSql = "SELECT MIN(bus_start_time) as bus_start_time FROM Work_env_bus_info WHERE " & DBTimeCompareEX(dtMinDate, "bus_start_time", "<") & " AND check_gate_id='" & Trim(m_sCheckGateNo) & "' AND " & DBDateCompare(Date, "bus_date") & " AND (status=" & Str(EREBusStatus.ST_BusNormal) & " or status=32 OR status = 16) order by bus_start_time"
'                Set rsTemp = oDb.Execute(szSql)
'                GoTo here1
'            End If
'        End If
'        '能找到则对以下车次检票，不能找到则不会执行以下语句
'
'        szSql = "SELECT bus_id,bus_date FROM Work_env_bus_info " & _
'        "WHERE " & DBDateCompare(Date, "bus_date") & " AND bus_id='" & Trim(szMinBusID) & "'"
'
    End If
    
    If Trim(szSql) = "" Then Exit Function
    Set rsTemp = oDb.Execute(szSql, , -1)
    If rsTemp.RecordCount = 0 Then '没有车次被选择
        Exit Function
    End If
    
    oREBus.Init m_oActiveUser
   ' oREBus.Init m_szUser
    oREBus.Identify rsTemp.Fields("bus_id"), rsTemp.Fields("bus_date"), rsTemp.Fields("check_gate_id")
    

    Set GetNextCheckBus = oREBus

End Function


'检票
'Public Sub CheckTicket(BusID As String, BusSerialNo As Integer, TicketID As String, dtBusDate As Date)
'    AssertActiveUserValid m_oActiveUser, cnErroeStartNo
'    'AssertHaveRight m_oActiveUser, RIGHT_ChkTkCheckTicket
'    Dim szChangeBusCheckGate As String '改乘车票的车次检票口
'    Dim odb  As New RTConnection
'    Dim rsTemp As New Recordset
'    Dim szSql As String
'    odb.ConnectionString = GetConnectionStr(cszSystemMan)
'
'    Dim nCheckStatus As Integer
'    Dim szDesStationID As String
'    Dim szTicketBusID As String
'    Dim dbTicketPrice As Double
'    Dim szSeatNo As String
'    Dim eTPType As ETicketType
''    Dim oSale As New SellTicketService
'    BusID = Trim(BusID)
'    TicketID = Trim(TicketID)
'    nCheckStatus = ECheckStatus.NormalTicket
'    If Len(m_sCheckGateNo) = 0 Then '未指定检票口
'        ShowError ERR_ChkTkCheckGateIDNotAppointed
'        Exit Sub
'    End If
'
'    szSql = " SELECT seat_no,ticket_type,ticket_price,bus_id,status,des_station_id FROM Ticket_sell_lst " & _
'            " WHERE ticket_id='" & TicketID & "' "
'    Set rsTemp = odb.Execute(szSql, , -1)
'    szSeatNo = rsTemp!seat_no
'    szDesStationID = rsTemp!des_station_id
'    szTicketBusID = Trim(rsTemp!bus_id)
'    dbTicketPrice = rsTemp!ticket_price
'    eTPType = rsTemp!ticket_type
'
'    Select Case rsTemp!Status
'           Case ST_TicketNormal
'           Case ST_TicketSellChange
'           Case Else '判断此票是否有效
'                ShowError ERR_ChkTkTicketStatusError
'    End Select
'    If BusID <> UCase(szTicketBusID) Then '检票车次不用原因(1)车次是并班车次(2)车票是改乘车票
'        If mp_bMergeBusStatus AND IsMergeBus(szTicketBusID, mp_szMergeBusID) = True Then '判断此车次是否有并班,如果有并班
'            nCheckStatus = ECheckStatus.MergeTicket
'            GoTo CanCheckTicket
'        End If
'        If mp_bChangeBus Then '判断系统是否让车辆改乘(1)能改乘(2)报错不能改乘
'            '(1)能改乘则判断检票车次是否通过车票所要经过站点和车次是否通一检票口
'            'szSql = "SELECT check_gate_id FROM Work_env_bus_info WHERE " & DBDateCompare(Date, "bus_date") & " AND bus_id='" & szTicketBusID & "'"
'            'Set rsTemp = oDb.Execute(szSql, , -1)
'            'szChangeBusCheckGate = Trim(rsTemp!check_gate_id)
'            'If szChangeBusCheckGate <> Trim(m_sCheckGateNo) Then  '判断指定的车次是否在本检票口检票
'            '    ShowError ERR_ChkTkCheckGateIsNotRight
'            'End If
'            '(2)车次是否通过该站点
'            szSql = " SELECT station_id FROM Work_env_bus_station_lst " & _
'                    " WHERE " & DBDateCompare(Date, "bus_date") & " AND bus_id = '" & BusID & "'" & _
'                    " AND station_id='" & szDesStationID & "'"
'            Set rsTemp = odb.Execute(szSql, , -1)
'            If rsTemp.RecordCount = 0 Then '指定的车次不过此车票的到站
'                ShowError ERR_ChkTkBusNotPassStation
'            Else
'                nCheckStatus = ECheckStatus.ChangeTicket
'                GoTo CanCheckTicket
'            End If
'        End If
'        '不是改乘也不是并班则
'        ShowError ERR_ChkTkTicketNotChecked
'    End If
'CanCheckTicket:
'        szSql = "INSERT check_ticket_lst(ticket_id,bus_date,bus_serial_no,check_time,bus_id,ticket_type,ticket_price,status,end_station_id) VALUES("
'        szSql = szSql & "'" & TicketID & "',"
'        szSql = szSql & "'" & ToDBDateTime(Date) & "',"
'        szSql = szSql & BusSerialNo & ","
'        szSql = szSql & "'" & ToDBDateTime(Now) & "',"
'        szSql = szSql & "'" & BusID & "',"
'        szSql = szSql & eTPType & ","
'        szSql = szSql & dbTicketPrice & ","
'        szSql = szSql & nCheckStatus & ",'" & szDesStationID & "')"
'    odb.BeginTrans
'        odb.Execute szSql
'        szSql = "UPDATE Ticket_sell_lst SET status=status|" & Str(ETicketStatus.ST_TicketChecked) & " WHERE ticket_id='" & TicketID & "'"
'        odb.Execute szSql
'    If nCheckStatus = ECheckStatus.ChangeTicket Then '如果车次是改乘车票,修改原车次座位状态和可售座位数
'        szSql = "UPDATE Env_bus_seat_lst SET status=" & Str(ST_SeatCanSell) & " WHERE bus_id='" & szTicketBusID & "' AND  bus_date='" & ToDBDate(dtBusDate) & "' AND seat_no='" & szSeatNo & "'"
'        odb.Execute szSql
'        szSql = "UPDATE Env_bus_info SET sale_seat_quantity=sale_seat_quantity+1 WHERE bus_id='" & szTicketBusID & "' AND  bus_date='" & ToDBDate(dtBusDate) & "'"
'        odb.Execute szSql
'    End If
'    odb.CommitTrans
'End Sub

Public Sub CheckTicket(BusID As String, BusSerialNo As Integer, TicketID As String, dtBusDate As Date, Optional nTicketStatus As Integer = 0)
    AssertActiveUserValid m_oActiveUser, cnErroeStartNo
    'AssertHaveRight m_oActiveUser, RIGHT_ChkTkCheckTicket
    Dim szChangeBusCheckGate As String '改乘车票的车次检票口
    Dim oDb  As New RTConnection
    Dim rsTemp As New Recordset
    Dim szSql As String
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    
    Dim nCheckStatus As Integer
    Dim szDesStationID As String
    Dim szTicketBusID As String
    Dim dbTicketPrice As Double
    Dim szSeatNo As String
    Dim eTPType As ETicketType
'    Dim oSale As New SellTicketService
    BusID = Trim(BusID)
    TicketID = Trim(TicketID)
    nCheckStatus = ECheckStatus.NormalTicket
    If Len(m_sCheckGateNo) = 0 Then '未指定检票口
        ShowError ERR_ChkTkCheckGateIDNotAppointed
        Exit Sub
    End If
    szSql = " SELECT seat_no,ticket_type,ticket_price,bus_id,status,des_station_id FROM Ticket_sell_lst " & _
            " WHERE ticket_id='" & TicketID & "' "
    Set rsTemp = oDb.Execute(szSql, , -1)
    szSeatNo = rsTemp!seat_no
    szDesStationID = rsTemp!des_station_id
    szTicketBusID = Trim(rsTemp!bus_id)
    dbTicketPrice = rsTemp!ticket_price
    eTPType = rsTemp!ticket_type
    
    Select Case rsTemp!Status
           Case ST_TicketNormal
           Case ST_TicketSellChange
           Case Else '判断此票是否有效
                ShowError ERR_ChkTkTicketStatusError
    End Select
    If BusID <> UCase(szTicketBusID) Then '检票车次不用原因(1)车次是并班车次(2)车票是改乘车票
        If mp_bMergeBusStatus And IsMergeBus(szTicketBusID, mp_szMergeBusID) = True Then '判断此车次是否有并班,如果有并班
            nCheckStatus = ECheckStatus.MergeTicket
            GoTo CanCheckTicket
        End If
        If mp_bChangeBus Then '判断系统是否让车辆改乘(1)能改乘(2)报错不能改乘
            '(1)能改乘则判断检票车次是否通过车票所要经过站点和车次是否通一检票口
            'szSql = "SELECT check_gate_id FROM Work_env_bus_info WHERE " & DBDateCompare(Date, "bus_date") & " AND bus_id='" & szTicketBusID & "'"
            'Set rsTemp = oDb.Execute(szSql, , -1)
            'szChangeBusCheckGate = Trim(rsTemp!check_gate_id)
            'If szChangeBusCheckGate <> Trim(m_sCheckGateNo) Then  '判断指定的车次是否在本检票口检票
            '    ShowError ERR_ChkTkCheckGateIsNotRight
            'End If
            '(2)车次是否通过该站点
            szSql = " SELECT station_id FROM Work_env_bus_station_lst " & _
                    " WHERE " & DBDateCompare(dtBusDate, "bus_date") & " AND bus_id = '" & BusID & "'" & _
                    " AND station_id='" & szDesStationID & "'"
            Set rsTemp = oDb.Execute(szSql, , -1)
            If rsTemp.RecordCount = 0 Then '指定的车次不过此车票的到站
                ShowError ERR_ChkTkBusNotPassStation
            Else
                nCheckStatus = ECheckStatus.ChangeTicket
                GoTo CanCheckTicket
            End If
        End If
        '不是改乘也不是并班则
        ShowError ERR_ChkTkTicketNotChecked
    End If
CanCheckTicket:
        szSql = "INSERT check_ticket_lst(check_gate_id,ticket_id,bus_date,bus_serial_no,check_time,bus_id,ticket_type,ticket_price,status,end_station_id) VALUES("
        szSql = szSql & "'" & m_sCheckGateNo & "',"
        szSql = szSql & "'" & TicketID & "',"
        szSql = szSql & "'" & ToDBDateTime(dtBusDate) & "',"
        szSql = szSql & BusSerialNo & ","
        szSql = szSql & "'" & ToDBDateTime(Now) & "',"
        szSql = szSql & "'" & BusID & "',"
        szSql = szSql & eTPType & ","
        szSql = szSql & dbTicketPrice & ","
        szSql = szSql & nCheckStatus & ",'" & szDesStationID & "')"
    oDb.BeginTrans
        oDb.Execute szSql
        szSql = "UPDATE Ticket_sell_lst SET status=status|" & Str(ETicketStatus.ST_TicketChecked) & " WHERE ticket_id='" & TicketID & "'"
        oDb.Execute szSql
        
    If nCheckStatus = ECheckStatus.ChangeTicket Then '如果车次是改乘车票,修改原车次座位状态和可售座位数
'        If nTicketStatus = 0 Then
'             RecycleSeat TicketID, ST_SeatCanSell, oDb, BusID
'        ElseIf nTicketStatus = 1 Then
'             szSql = "UPDATE Env_bus_seat_lst SET status=" & Str(ST_SeatReserved) & " WHERE bus_id='" & szTicketBusID & "' AND  bus_date='" & ToDBDate(dtBusDate) & "' AND seat_no='" & szSeatNo & "'"
'             oDb.Execute szSql
'        End If
        szSql = "UPDATE Env_bus_seat_lst SET status=" & Str(ST_SeatReserved) & ",remark='" & sczChangeSeatReserved & "' WHERE bus_id='" & szTicketBusID & "' AND  bus_date='" & ToDBDate(dtBusDate) & "' AND seat_no='" & szSeatNo & "'"
        oDb.Execute szSql
        '把原上车站点已售票减1
        szSql = "UPDATE env_bus_allot_lst SET sell_quantity=sell_quantity-1 WHERE bus_id='" & szTicketBusID & "' AND bus_date='" & ToDBDate(dtBusDate) & "' AND sell_station_id =(select sell_station_id from ticket_sell_lst where ticket_id='" & TicketID & "')"
        oDb.Execute szSql
    End If
    
    oDb.CommitTrans
    If nCheckStatus = ECheckStatus.ChangeTicket Then
        '如果为改乘,则写日志
        WriteOperateLog m_oActiveUser, RIGHT_ChkTkChangeCheckBus, "车票" & TicketID & "改乘到" & BusID & "车次"
        
    End If
    Exit Sub
ErrorHandle:
    err.Raise err.Number, err.Source, err.Description
    
End Sub
'开检流水车次
Public Sub StartCheckScrollBus(BusID As String, BusSerialNo As Integer, VehicleID As String, Optional BusDate As Date)
    'AssertActiveUserValid m_oActiveUser, cnErroeStartNo
    'AssertHaveRight m_oActiveUser, RIGHT_ChkTkCheckBus
    
    Dim oDb As New RTConnection
    Dim rsTemp  As Recordset
    Dim szSql As String
    Dim dtDate As Date
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    
    Dim szCheckGateID As String
    
    If BusDate = "0:00:00" Then
       dtDate = Date
    Else
       dtDate = BusDate
    End If
    
    If Len(m_sCheckGateNo) = 0 Then '未指定检票口
        ShowError ERR_ChkTkCheckGateIDNotAppointed
        Exit Sub
    End If
    If CheckUserIsCheckGate(oDb, m_oActiveUser.UserID) = False Then
        ShowError ERR_ChkTkUserNotCheckGate
        Exit Sub
    End If
'    szSql = "SELECT vehicle_id,owner_id,transport_company_id,vehicle_type_code,split_company_id FROM Vehicle_info WHERE vehicle_id='" & Trim(VehicleID) & "'"
    szSql = "SELECT v.*,t.vehicle_type_short_name" & _
            " FROM Vehicle_info v,vehicle_type_code t" & _
            " WHERE v.vehicle_id='" & VehicleID & "' AND v.vehicle_type_code=t.vehicle_type_code"
    Set rsTemp = oDb.Execute(szSql, , -1)
    If rsTemp.RecordCount = 0 Then '指定的车辆不存在
        ShowError ERR_ChkTkVehicleNotExist
        Exit Sub
    End If
    '车辆信息
    Dim szvehicle_id As String
    Dim szvehicle_type_code As String
    Dim szowner_id As String
    Dim sztransport_company_id As String
    Dim szvehicle_type_name As String
    Dim szsplitcompany_id As String
    
    szvehicle_type_code = FormatDbValue(rsTemp!vehicle_type_code)
    szowner_id = FormatDbValue(rsTemp!owner_id)
    sztransport_company_id = FormatDbValue(rsTemp!transport_company_id)
    szvehicle_id = FormatDbValue(rsTemp!vehicle_id)
    szvehicle_type_name = FormatDbValue(rsTemp!vehicle_type_short_name)
    szsplitcompany_id = FormatDbValue(rsTemp!split_company_id)
    '查询检票车次信息是否已存在
    
    szSql = "SELECT bus_id FROM Bus_check_lst WHERE " & DBDateCompare(dtDate, "bus_date") & "AND bus_id='" & BusID & "'" & "AND bus_serial_no=" & Str(BusSerialNo) _
         & " AND check_gate_id='" & m_sCheckGateNo & "'"
    Set rsTemp = oDb.Execute(szSql, , -1)
    If rsTemp.RecordCount > 0 Then '指定车次指定序号的流水车次已存在
        ShowError ERR_ChkTkBusAlreadyExist
        Exit Sub
    End If
    szSql = "SELECT route_id,bus_date,bus_id,check_gate_id,bus_start_time,register_status,status,owner_id,transport_company_id,split_company_id,vehicle_id " _
        & " FROM Work_env_bus_info WHERE bus_id='" & BusID & "' AND " & DBDateCompare(dtDate, "bus_date", "=")
        
    Set rsTemp = oDb.Execute(szSql, , -1)
    szCheckGateID = rsTemp.Fields("check_gate_id")
    If rsTemp.RecordCount = 0 Then '指定的车次不存在
        ShowError ERR_ChkTkBusNotExist
        Exit Sub
    End If
    If Trim(szCheckGateID) <> Trim(m_sCheckGateNo) Then '指定的车次不在当前检票口检票
        ShowError ERR_ChkTkCheckGateIsNotRight
        Exit Sub
    End If
    oDb.BeginTrans
        '写入检票车次明细表
        szSql = "INSERT Bus_check_lst(check_gate_id,bus_date,bus_id,bus_serial_no,vehicle_id,bus_start_time,check_start_time,check_end_time,checker,transport_company_id,split_company_id,owner_id,route_id,vehicle_type_code) VALUES("
        szSql = szSql & "'" & m_sCheckGateNo & "',"
        szSql = szSql & "'" & ToDBDateTime(dtDate) & "',"
        szSql = szSql & "'" & BusID & "',"
        szSql = szSql & Str(BusSerialNo) & ","
        szSql = szSql & "'" & szvehicle_id & "',"
        szSql = szSql & "'" & ToDBDateTime(rsTemp.Fields("bus_start_time")) & "',"
        szSql = szSql & "'" & ToDBDateTime(Now) & "',"
        szSql = szSql & "'" & ToDBDateTime(cszEmptyDateStr & " " & cszEmptyTimeStr) & "',"
        szSql = szSql & "'" & m_szUser & "',"
        szSql = szSql & "'" & sztransport_company_id & "',"
        szSql = szSql & "'" & szsplitcompany_id & "',"
        szSql = szSql & "'" & szowner_id & "','" & rsTemp!route_id & "','" & szvehicle_type_code & "')"
        oDb.Execute szSql
    '           " SET status=" & EREBusStatus.ST_BusChecking &
        szSql = "UPDATE Env_bus_info" & _
                 " SET vehicle_id='" & szvehicle_id & "'" & _
                 "    ,transport_company_id='" & sztransport_company_id & "'" & _
                 "    ,owner_id='" & szowner_id & "'" & _
                 "    ,split_company_id='" & szsplitcompany_id & "'" & _
                 "    ,vehicle_type_code='" & szvehicle_type_code & "'" & _
                 "    ,vehicle_type_name='" & szvehicle_type_name & "'" & _
                 " WHERE bus_id='" & BusID & "' AND " & DBDateCompare(dtDate, "bus_date")
        '滚动车次不写车次状态
        ''szSql = szSql & " , status= & TransFieldValueToString( EREBusStatus.ST_BusChecking) & "
                 
                 
        
        oDb.Execute szSql '流水车次回写环境车次表状态
    oDb.CommitTrans
    WriteOperateLog m_oActiveUser, RIGHT_ChkTkCheckBus, "在开检时，把车次[" & BusID & "]日期为[" & dtDate & "]序号为[" & Str(BusSerialNo) & "]，当前车次车辆改为" & VehicleID

'以下新加消息发送代码
#If SUPPORT_MSG = 1 Then
    Dim szBusDate As String
    szBusDate = ToDBDateTime(dtDate)
    Dim aszValue(1 To 3) As String
    aszValue(1) = BusID
    aszValue(2) = szBusDate
    aszValue(3) = Str(BusSerialNo)
    m_oMsg.SellStation = m_szSellStation
    m_oMsg.SendMsg Now, cszStartCheckBus, aszValue
#End If
End Sub

'开检固定车次
Public Sub StartCheckRegularBus(BusID As String, Optional VehicleID As String, Optional BusDate As Date)
    AssertActiveUserValid m_oActiveUser, cnErroeStartNo
    'AssertHaveRight m_oActiveUser, RIGHT_ChkTkCheckBus
    
    Dim oDb As New RTConnection
    Dim rsTemp  As New Recordset
    Dim dtDate As Date
    Dim szSql As String
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    
    Dim szCheckGateID As String
    
    If BusDate <> "0:00:00" Then
       dtDate = BusDate
    Else
       dtDate = Date
    End If
    
    If Len(m_sCheckGateNo) = 0 Then '未指定检票口
        ShowError ERR_ChkTkCheckGateIDNotAppointed
        Exit Sub
    End If
    If CheckUserIsCheckGate(oDb, m_oActiveUser.UserID) = False Then
        ShowError ERR_ChkTkUserNotCheckGate
        Exit Sub
    End If
    szSql = "SELECT route_id,vehicle_type_code,a.bus_date,a.bus_id,b.check_gate_id,b.bus_start_time, " & _
            " register_status,b.status,owner_id,transport_company_id,split_company_id,vehicle_id, " & _
            " vehicle_type_name vehicle_type_short_name" & _
            " FROM Work_env_bus_info a,Work_env_bus_allot_lst b" & _
            " WHERE a.bus_id='" & BusID & "' AND " & DBDateCompare(dtDate, "a.bus_date", "=") & _
            " AND a.bus_id=b.bus_id AND a.bus_date=b.bus_date" & _
            " AND b.check_gate_id=" & TransFieldValueToString(m_sCheckGateNo)
    Set rsTemp = oDb.Execute(szSql, , -1)
    
    
    If rsTemp.RecordCount = 0 Then '指定的车次不存在
        ShowError ERR_ChkTkBusNotExist
        Exit Sub
    End If
    If FormatDbValue(rsTemp!Status) = EREBusStatus.ST_BusStopped Then
        ShowError ERR_REBusIsStopped, , "车次已停班"
    End If
    
    szCheckGateID = rsTemp!check_gate_id
    
    '更改车辆信息
    Dim szvehicle_id As String
    Dim szvehicle_type_code As String
    Dim szowner_id As String
    Dim sztransport_company_id As String
    Dim szvehicle_type_name As String
    Dim szsplitcompany_id As String
    
    Dim rsVehicle As Recordset
    If Trim(rsTemp.Fields("vehicle_id")) <> Trim(VehicleID) And VehicleID <> "" Then
        szSql = "SELECT v.*,t.vehicle_type_short_name" & _
                " FROM Vehicle_info v,vehicle_type_code t" & _
                " WHERE v.vehicle_id='" & VehicleID & "' AND v.vehicle_type_code=t.vehicle_type_code"
        Set rsVehicle = oDb.Execute(szSql)
        If rsVehicle.RecordCount = 0 Then ShowError ERR_ChkTkVehicleNotExist
'        If rsVehicle!Status = 0 Then ShowError
        szvehicle_type_code = FormatDbValue(rsVehicle!vehicle_type_code)
        szowner_id = FormatDbValue(rsVehicle!owner_id)
        sztransport_company_id = FormatDbValue(rsVehicle!transport_company_id)
        szvehicle_id = FormatDbValue(rsVehicle!vehicle_id)
        szvehicle_type_name = FormatDbValue(rsVehicle!vehicle_type_short_name)
        szsplitcompany_id = FormatDbValue(rsVehicle!split_company_id)
    Else
        szvehicle_type_code = FormatDbValue(rsTemp!vehicle_type_code)
        szowner_id = FormatDbValue(rsTemp!owner_id)
        sztransport_company_id = FormatDbValue(rsTemp!transport_company_id)
        szvehicle_id = FormatDbValue(rsTemp!vehicle_id)
        szvehicle_type_name = FormatDbValue(rsTemp!vehicle_type_short_name)
        szsplitcompany_id = FormatDbValue(rsTemp!split_company_id)
    End If
    
    
'    If Trim(szCheckGateID) <> Trim(m_sCheckGateNo) Then
'        ShowError ERR_ChkTkCheckGateIsNotRight '指定的车次不在当前检票口检票
'        Exit Sub
'    End If

    '********以下注释是因为富阳站存在某些车次未到发车时间,提前检票
    
'    If DateAdd("n", -GetParameterValue(cszStartCheckTime), rsTemp.Fields("bus_start_time")) > Now Then   '未到开检时间，出错
'        ShowError ERR_ChkTkCheckTimeIsNotStarting
'        Exit Sub
'    End If
    '********
    '写入检票车次明细表
    szSql = "INSERT Bus_check_lst(check_gate_id,bus_date,bus_id,bus_serial_no,vehicle_id,bus_start_time,check_start_time,check_end_time,checker,transport_company_id,split_company_id,owner_id,route_id,vehicle_type_code) VALUES("
    szSql = szSql & "'" & m_sCheckGateNo & "',"
    szSql = szSql & "'" & ToDBDateTime(rsTemp.Fields("bus_date")) & "',"
    szSql = szSql & "'" & BusID & "',"
    szSql = szSql & 0 & ","
    szSql = szSql & "'" & szvehicle_id & "',"
    szSql = szSql & "'" & ToDBDateTime(rsTemp.Fields("bus_start_time")) & "',"
    szSql = szSql & "'" & ToDBDateTime(Now) & "',"
    szSql = szSql & "'" & ToDBDateTime(cszEmptyDateStr & " " & cszEmptyTimeStr) & "',"
    szSql = szSql & "'" & SelfUser.UserID & "',"
    szSql = szSql & "'" & sztransport_company_id & "',"
    szSql = szSql & "'" & szsplitcompany_id & "',"
    szSql = szSql & "'" & szowner_id & "','" & rsTemp!route_id & "','" & szvehicle_type_code & "')"
    oDb.BeginTrans
    oDb.Execute szSql
    szSql = "UPDATE Env_bus_info" & _
             " SET status=" & EREBusStatus.ST_BusChecking & _
             "    ,vehicle_id='" & szvehicle_id & "'" & _
             "    ,transport_company_id='" & sztransport_company_id & "'" & _
             "    ,owner_id='" & szowner_id & "'" & _
             "    ,split_company_id='" & szsplitcompany_id & "'" & _
             "    ,vehicle_type_code='" & szvehicle_type_code & "'" & _
             "    ,vehicle_type_name='" & szvehicle_type_name & "'" & _
             " WHERE bus_id='" & BusID & "' AND " & DBDateCompare(dtDate, "bus_date")
    oDb.Execute szSql '固定车次回写环境车次表状态
    
    szSql = "UPDATE Env_bus_allot_lst" & _
             " SET status=" & EREBusStatus.ST_BusChecking & _
             " WHERE bus_id='" & BusID & "' AND " & DBDateCompare(dtDate, "bus_date") & " AND check_gate_id=" & TransFieldValueToString(m_sCheckGateNo)
    oDb.Execute szSql '固定车次回写环境车次表状态
    oDb.CommitTrans
    
    WriteOperateLog m_oActiveUser, RIGHT_ChkTkCheckBus, "在开检时，把车次[" & BusID & "]日期为[" & dtDate & "]，当前车次车辆改为" & VehicleID
 
    
'以下新加消息发送代码
#If SUPPORT_MSG = 1 Then
    Dim szBusDate As String
    szBusDate = ToDBDateTime(rsTemp.Fields("bus_date"))
    Dim aszValue(1 To 3) As String
    aszValue(1) = BusID
    aszValue(2) = szBusDate
    aszValue(3) = "0"
    m_oMsg.SellStation = m_szSellStation
    m_oMsg.SendMsg Now, cszStartCheckBus, aszValue
#End If

End Sub

'补检
Public Sub ExtraStartCheckBus(BusID As String, Optional pnSerialNo As Integer = 0, Optional pdtBusDate As Date, Optional pbTertian As Boolean)
    AssertActiveUserValid m_oActiveUser, cnErroeStartNo
    'AssertHaveRight m_oActiveUser, RIGHT_ChkTkExtraCheckBus
    
    Dim oDb As New RTConnection
    Dim rsTemp  As New Recordset
    Dim szSql As String
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    
    Dim nBusSerial As Integer
    Dim nBusType As Integer
    Dim szCheckGateID As String
    Dim dBusStartTime As Date
    Dim gDelayTime As Single
    Dim dtDate As Date
    
    If pdtBusDate <> "0:00:00" Then
       dtDate = pdtBusDate
    Else
       dtDate = Date
    End If
    
    If Len(m_sCheckGateNo) = 0 Then '未指定检票口
        ShowError ERR_ChkTkCheckGateIDNotAppointed
        Exit Sub
    End If
   
    
    szSql = "SELECT a.bus_date,a.bus_id,bus_type,b.check_gate_id,b.bus_start_time,register_status,b.status,owner_id,transport_company_id,vehicle_id " & _
            " FROM Work_env_bus_info a,Work_env_bus_allot_lst b" & _
            " WHERE a.bus_id='" & BusID & "' AND " & DBDateCompare(dtDate, "a.bus_date", "=") & _
            " AND a.bus_id=b.bus_id AND a.bus_date=b.bus_date" & _
            " AND b.check_gate_id=" & TransFieldValueToString(m_sCheckGateNo)
        
    Set rsTemp = oDb.Execute(szSql, , -1)
    If rsTemp.RecordCount = 0 Then
        ShowError ERR_ChkTkBusNotExist '指定的车次不存在
        Exit Sub
    End If
    szCheckGateID = rsTemp.Fields("check_gate_id")
'    If Trim(szCheckGateID) <> Trim(m_sCheckGateNo) Then '指定的车次不在当前检票口检票
'        ShowError ERR_ChkTkCheckGateIsNotRight
'        Exit Sub
'    End If
    
    nBusType = FormatDbValue(rsTemp.Fields("bus_type"))
    dBusStartTime = rsTemp.Fields("bus_start_time")
    gDelayTime = 0.05
    
    Dim bExistCheckRecord As Boolean
    If nBusType <> TP_ScrollBus Then
        nBusSerial = 0
'        If Format(rsTemp.Fields("bus_start_time") - GetParameterValue(cszStartCheckTime), "hh:mm:ss") > Format(Now, "hh:mm:ss") Then '开检时间大于当前时间，则不能补检
'           ShowError ERR_ChkTkExtraCheckTimeIsNotStarting
'           Exit Sub
'        End If
'        If Format(rsTemp.Fields("bus_start_time") - GetParameterValue(cszExtraStartCheckTime), "hh:mm:ss") < Format(Now, "hh:mm:ss") Then '补检时间大于当前时间，则不能补检
'           ShowError ERR_ChkTkExtraCheckTimeAlreadyEnd
'           Exit Sub
'        End If
'原来不能隔天检的判断
'        If DateDiff("n", FormatDateTime(SelfNowDateTime), DateAdd("n", -GetParameterValue(cszStartCheckTime), rsTemp.Fields("bus_start_time"))) > 0 Then   '开检时间大于当前时间，则不能补检
'           ShowError ERR_ChkTkExtraCheckTimeIsNotStarting
'           Exit Sub
'        End If
'        If DateDiff("n", FormatDateTime(SelfNowDateTime), DateAdd("n", GetParameterValue(cszExtraStartCheckTime), rsTemp.Fields("bus_start_time"))) > 0 Then        '补检时间大于当前时间，则不能补检
'           ShowError ERR_ChkTkExtraCheckTimeAlreadyEnd
'           Exit Sub
'        End If
'原来不能隔天检的判断
        
        If pbTertian = False Then '是当天补检，需要检查时间，否则是隔日或几日，不检查时间
            If DateDiff("n", FormatDateTime(SelfNowDateTime), DateAdd("n", -GetParameterValue(cszStartCheckTime), rsTemp.Fields("bus_start_time"))) > 0 Then '开检时间大于当前时间，则不能补检
               ShowError ERR_ChkTkExtraCheckTimeIsNotStarting
               Exit Sub
            End If
            If DateDiff("n", FormatDateTime(SelfNowDateTime), DateAdd("n", GetParameterValue(cszExtraStartCheckTime), rsTemp.Fields("bus_start_time"))) > 0 Then          '补检时间大于当前时间，则不能补检
               ShowError ERR_ChkTkExtraCheckTimeAlreadyEnd
               Exit Sub
            End If
        End If
        
        '查询检票车次信息是否已存在
        szSql = "SELECT bus_id FROM Bus_check_lst WHERE bus_id='" & BusID & "' AND " & DBDateCompare(dtDate, "bus_date") _
            & " AND check_gate_id='" & m_sCheckGateNo & "'"
        Set rsTemp = oDb.Execute(szSql)
        If rsTemp.RecordCount = 0 Then bExistCheckRecord = False Else bExistCheckRecord = True
    Else
        If pnSerialNo <> 0 Then
            nBusSerial = pnSerialNo
        Else
            '得到最大的车次序号
            szSql = "SELECT MAX(bus_serial_no) bus_serial_no FROM Bus_check_lst WHERE " & DBDateCompare(dtDate, "bus_date") & " AND bus_id='" & BusID & "'" _
                & " AND check_gate_id='" & m_sCheckGateNo & "'"
    
    
            Set rsTemp = oDb.Execute(szSql)
            nBusSerial = FormatDbValue(rsTemp.Fields("bus_serial_no"))
        End If
        If nBusSerial = 0 Then bExistCheckRecord = False Else bExistCheckRecord = True
    End If
    If Not bExistCheckRecord Then ShowError ERR_ChkTkBusIsNotChecked
    '得到对应的路单
    szSql = "SELECT check_sheet_id FROM sheet_lst WHERE " & DBDateCompare(dtDate, "bus_date") & _
            " AND bus_id='" & BusID & "'" & " AND bus_serial_no=" & nBusSerial & " AND valid_mark=1" _
         & " AND check_gate_id='" & m_sCheckGateNo & "'"
    Set rsTemp = oDb.Execute(szSql, , -1)
    
    oDb.BeginTrans
'        If rsTemp.RecordCount > 0 Then
'            CancelCheckSheet oDb, FormatDbValue(rsTemp.Fields("check_sheet_id"))
'        End If
        '修改检票车次明细表
        szSql = "UPDATE Bus_check_lst SET check_end_time=" & "'" & ToDBDateTime(dBusStartTime - GetParameterValue(cszStartCheckTime) - gDelayTime) & "'" & ",checker='" & SelfUser.UserID & "'" _
                 & " WHERE bus_id='" & BusID & "' AND bus_serial_no = " & nBusSerial & " AND " & DBDateCompare(dtDate, "bus_date") _
                 & " AND check_gate_id=" & TransFieldValueToString(m_sCheckGateNo)
        oDb.Execute szSql
        '如果为滚动车次,则不写车次的状态
        If nBusType <> TP_ScrollBus Then
            szSql = "UPDATE Env_bus_info SET status=" & EREBusStatus.ST_BusExtraChecking & " WHERE bus_id='" & BusID & "' AND " & DBDateCompare(dtDate, "bus_date")
            oDb.Execute szSql
            szSql = "UPDATE Env_bus_allot_lst SET status=" & EREBusStatus.ST_BusExtraChecking & _
                    " WHERE bus_id='" & BusID & "' AND " & DBDateCompare(dtDate, "bus_date") & " AND check_gate_id=" & TransFieldValueToString(m_sCheckGateNo)
            oDb.Execute szSql
        End If
    oDb.CommitTrans
    
'以下消息发送代码
#If SUPPORT_MSG = 1 Then
    Dim szBusDate  As String
    szBusDate = ToDBDateTime(dtDate)
    Dim aszValue(1 To 2) As String
    aszValue(1) = BusID
    aszValue(2) = szBusDate
    m_oMsg.SellStation = m_szSellStation
    m_oMsg.SendMsg Now, cszExStartCheckBus, aszValue
#End If
    
    'WriteOperateLog m_oActiveUser, RIGHT_ChkTkExtraCheckBus
End Sub

'停检
Public Sub StopCheckTicket(BusID As String, Optional pnSerialNo As Integer = 0, Optional pszCheckMode As Boolean, Optional pdtBusDate As Date)
    AssertActiveUserValid m_oActiveUser, cnErroeStartNo
    'AssertHaveRight m_oActiveUser, RIGHT_ChkTkStopCheckBus
    Dim nBusType As Integer
    Dim oDb As New RTConnection
    Dim rsTemp  As New Recordset
    Dim rsCheckSheetIDTemp As New Recordset
    Dim szSql As String
    Dim dtDate As Date
    Dim nBusSerialNo As Integer
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    
    If pdtBusDate <> "0:00:00" Then
       dtDate = pdtBusDate
    Else
       dtDate = Date
    End If
    
    If Len(m_sCheckGateNo) = 0 Then '未指定检票口
        ShowError ERR_ChkTkCheckGateIDNotAppointed
        Exit Sub
    End If
    szSql = "SELECT bus_type FROM Work_env_bus_info WHERE bus_id='" & BusID & "' AND " & DBDateCompare(dtDate, "bus_date")
    Set rsTemp = oDb.Execute(szSql, , -1)
    nBusSerialNo = 0
    nBusType = rsTemp!bus_type
    If nBusType = EBusType.TP_ScrollBus Then
        
        If pnSerialNo <> 0 Then
            nBusSerialNo = pnSerialNo
        Else
            '得到最大的车次序号
            szSql = "SELECT MAX(bus_serial_no) AS bus_serial_no FROM Bus_check_lst WHERE bus_id='" & BusID & "' AND " & DBDateCompare(dtDate, "bus_date") _
               & " AND check_gate_id='" & m_sCheckGateNo & "'"
            Set rsTemp = oDb.Execute(szSql, , -1)
            If Not IsNull(rsTemp.Fields("bus_serial_no")) Then
                nBusSerialNo = rsTemp.Fields("bus_serial_no")
            Else
                nBusSerialNo = 1
            End If
            
        End If
    End If
    
    If pszCheckMode = True Then
    '得到对应的路单
    szSql = "SELECT check_sheet_id FROM sheet_lst WHERE " & DBDateCompare(dtDate, "bus_date") & _
            " AND bus_id='" & BusID & "'" & " AND bus_serial_no=" & nBusSerialNo & " AND valid_mark=1" _
         & " AND check_gate_id='" & m_sCheckGateNo & "'"
    Set rsCheckSheetIDTemp = oDb.Execute(szSql, , -1)
    '作废路单
    If rsCheckSheetIDTemp.RecordCount > 0 Then
        CancelCheckSheet oDb, FormatDbValue(rsCheckSheetIDTemp.Fields("check_sheet_id"))
    End If
    End If
    
    oDb.BeginTrans
        '修改检票车次明细表
        szSql = "UPDATE Bus_check_lst SET check_end_time='" & ToDBDateTime(Now) & "'" _
            & " WHERE bus_id='" & BusID & "' AND " & DBDateCompare(dtDate, "bus_date") & " AND bus_serial_no=" & Str(nBusSerialNo) _
            & " AND check_gate_id='" & m_sCheckGateNo & "'"
        oDb.Execute szSql
        If nBusType <> EBusType.TP_ScrollBus Then

            szSql = "UPDATE Env_bus_allot_lst SET status=" & Str(EREBusStatus.ST_BusStopCheck) & _
                    " WHERE bus_id='" & BusID & "' AND " & DBDateCompare(dtDate, "bus_date") & " AND check_gate_id=" & TransFieldValueToString(m_sCheckGateNo)
            oDb.Execute szSql
            '如果车次的发车时间是最后的班次则,更新该车次的检票状态为已检
            
            szSql = "UPDATE Env_bus_info SET status=" & Str(EREBusStatus.ST_BusStopCheck) & _
                    " WHERE bus_id='" & BusID & "' AND " & DBDateCompare(dtDate, "bus_date")
            oDb.Execute szSql
        Else
            '滚动车次不写车次检票状态
            
''            szSql = "UPDATE Env_bus_allot_lst SET status=" & Str(EREBusStatus.ST_BusNormal) & _
''                    " WHERE bus_id='" & BusID & "' AND " & DBDateCompare(Date, "bus_date") & " AND check_gate_id=" & TransFieldValueToString(m_sCheckGateNo)
''            oDb.Execute szSql
''            szSql = "UPDATE Env_bus_info SET status=" & Str(EREBusStatus.ST_BusNormal) & _
''                    " WHERE bus_id='" & BusID & "' AND " & DBDateCompare(Date, "bus_date")
''            oDb.Execute szSql
        End If
    
    oDb.CommitTrans
    'WriteOperateLog m_oActiveUser, RIGHT_ChkTkStopCheckBus
    
'以下新加消息发送代码
#If SUPPORT_MSG = 1 Then
    Dim szBusDate As String
    szBusDate = Format(dtDate, cszDateTimeStr)
    Dim aszValue(1 To 3) As String
    aszValue(1) = BusID
    aszValue(2) = szBusDate
    aszValue(3) = Str(nBusSerialNo)
    m_oMsg.SellStation = m_szSellStation
    m_oMsg.SendMsg Now, cszStopCheckBus, aszValue
#End If
    
End Sub

'注销检票
Public Sub WriteOffCheckTicket(TicketID As String)
    AssertActiveUserValid m_oActiveUser, cnErroeStartNo
    'AssertHaveRight m_oActiveUser, RIGHT_ChkTkWriteOffCheckTicket
    
    Dim oDb As New RTConnection
    Dim rsTemp  As New Recordset
    Dim szSql As String
    Dim szBusID As String
    Dim dtBusDate As Date
    Dim nBusSerial As Integer
    On Error GoTo ErrorHandle
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    
    szSql = "SELECT bus_id,bus_date,bus_serial_no , end_station_id, ticket_type,ticket_price FROM check_ticket_lst WHERE ticket_id='" & TicketID & "'"
    Set rsTemp = oDb.Execute(szSql, , -1)
    If rsTemp.RecordCount = 0 Then '没有该已检的车票
        ShowError ERR_ChkTkResultIsNull
    End If
    '======判断车次是否
    oDb.BeginTrans
'        szSql = "UPDATE sheet_station_lst SET  quantity=quantity-1 FROM Ticket_sell_lst a,sheet_lst c " _
'            & " WHERE a.ticket_id='" & TicketID & "' AND a.des_station_id=sheet_station_lst.station_id " _
'            & " AND a.ticket_type=sheet_station_lst.ticket_type"
'        szSql = szSql & " AND sheet_station_lst.check_sheet_id=c.check_sheet_id AND valid_mark=1"
'        oDb.Execute szSql
'
        szSql = "UPDATE sheet_station_lst SET  quantity=quantity-1 , ticket_price= ticket_price - " & TransFieldValueToString(rsTemp!ticket_price) _
            & " FROM sheet_station_lst s , sheet_lst c" _
            & " WHERE bus_id = " & TransFieldValueToString(rsTemp!bus_id) _
            & " AND bus_date =" & TransFieldValueToString(rsTemp!bus_date) & " AND  station_id = " & TransFieldValueToString(rsTemp!end_station_id) _
            & " AND bus_serial_no = " & TransFieldValueToString(rsTemp!bus_serial_no) _
            & " AND c.check_sheet_id = s.check_sheet_id " _
            & " AND ticket_type = " & TransFieldValueToString(rsTemp!ticket_type)
        
        oDb.Execute szSql
        
        '========删除车票=======
        szSql = "DELETE check_ticket_lst WHERE ticket_id='" & TicketID & "'"
        oDb.Execute szSql
        '========将车票值回正常售出===========
        szSql = "UPDATE Ticket_sell_lst SET status=1 WHERE ticket_id='" & TicketID & "'"
        oDb.Execute szSql
    oDb.CommitTrans
    WriteOperateLog m_oActiveUser, RIGHT_ChkTkWriteOffCheckTicket, "注销检票" & TicketID
    Exit Sub
ErrorHandle:
    oDb.RollbackTrans
    err.Raise err.Number, err.Source, err.Description
    
End Sub

'得到车次信息
Public Function GetBusInfoRs(BusDate As Date, CheckGateID As String) As Recordset
    AssertActiveUserValid m_oActiveUser, cnErroeStartNo
    
    Dim oDb As New RTConnection
    Dim rsTemp  As New Recordset
    Dim szSql As String
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
       
    If Len(CheckGateID) = 0 Then '未指定检票口
        ShowError ERR_ChkTkCheckGateIDNotAppointed
        Exit Function
    End If
    If DateDiff("d", BusDate, Now) <= 0 Then '如果查询的是环境中的车次则使用工作表
        szSql = "SELECT a.check_gate_id,e.bus_id,a.bus_start_time,e.bus_type,v.license_tag_no,c.transport_company_short_name" _
            & " ,o.owner_name,a.status,0 AS bus_serial_no,v.vehicle_id,c.transport_company_id,o.owner_id,e.end_station_name " _
            & " FROM Work_env_bus_info e,Work_env_bus_allot_lst a,Vehicle_info v,Company_info c,owner_info o"
    Else
        szSql = "SELECT a.check_gate_id,e.bus_id,a.bus_start_time,e.bus_type,v.license_tag_no,c.transport_company_short_name" _
            & " ,o.owner_name,a.status,0 AS bus_serial_no,v.vehicle_id,c.transport_company_id,o.owner_id ,e.end_station_name" _
            & " FROM Env_bus_info e,Env_bus_allot_lst a,Vehicle_info v,Company_info c,owner_info o"
    End If
    szSql = szSql & " WHERE e.vehicle_id = v.vehicle_id  AND a.can_sale_quantity <>0 AND e.transport_company_id=c.transport_company_id  AND e.owner_id=o.owner_id " _
                & " AND a.bus_date=e.bus_date  AND a.bus_id=e.bus_id" _
                & " AND " & DBDateCompare(BusDate, "e.bus_date") & " AND a.check_gate_id='" & CheckGateID & "'order by e.bus_start_time,e.bus_id"
    
    Set rsTemp = oDb.Execute(szSql, , -1)
    If rsTemp.RecordCount = 0 Then '无记录
        'ShowError ERR_ChkTkResultIsNull
        Set GetBusInfoRs = rsTemp
        Exit Function
    End If
    Set GetBusInfoRs = rsTemp
    
End Function
'得到车次信息
Public Function GetBusCheckSheetRs(BusDate As Date, CheckGateID As String) As Recordset
    AssertActiveUserValid m_oActiveUser, cnErroeStartNo
    
    Dim oDb As New RTConnection
    Dim szSql As String
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
       
    If Len(CheckGateID) = 0 Then '未指定检票口
        ShowError ERR_ChkTkCheckGateIDNotAppointed
        Exit Function
    End If
    If DateDiff("d", BusDate, Now) <= 0 Then '如果查询的是环境中的车次则使用工作表
        szSql = "SELECT e.*,b.bus_type,b.end_station_name,s.check_sheet_id,v.license_tag_no,c.transport_company_short_name,o.owner_name" & _
            " FROM Bus_check_lst e,Work_env_bus_info b,sheet_lst s,Vehicle_info v,Company_info c,owner_info o" & _
            " WHERE e.vehicle_id = v.vehicle_id  AND e.transport_company_id=c.transport_company_id  AND e.owner_id=o.owner_id " & _
            " AND e.bus_date=b.bus_date AND e.bus_id=b.bus_id AND e.bus_date=s.bus_date AND e.check_gate_id = s.check_gate_id AND e.bus_id=s.bus_id AND e.bus_serial_no=s.bus_serial_no AND s.valid_mark=1" & _
            " AND " & DBDateCompare(BusDate, "b.bus_date") & " AND e.check_gate_id='" & CheckGateID & "' ORDER BY e.bus_start_time,e.bus_id"
    Else
        szSql = "SELECT e.*,b.bus_type,b.end_station_name,s.check_sheet_id,v.license_tag_no,c.transport_company_short_name,o.owner_name " & _
            " FROM Bus_check_lst e,Env_bus_info b,Vehicle_info v,Company_info c,owner_info o" & _
            " WHERE e.vehicle_id = v.vehicle_id  AND e.transport_company_id=c.transport_company_id  AND e.owner_id=o.owner_id " & _
            " AND e.bus_date=b.bus_date AND e.bus_id=b.bus_id AND e.bus_date=s.bus_date AND e.check_gate_id = s.check_gate_id AND e.bus_id=s.bus_id AND e.bus_serial_no=s.bus_serial_no AND s.valid_mark=1" & _
            " AND " & DBDateCompare(BusDate, "b.bus_date") & " AND e.check_gate_id='" & CheckGateID & "' ORDER BY e.bus_start_time,e.bus_id"
    End If
    Set GetBusCheckSheetRs = oDb.Execute(szSql, , -1)
End Function
'得到车次类型名（因兼容性加此接口函数）
Public Function GetBusTypeName(BusType As Integer) As String
    Dim oDb As New RTConnection
    Dim rsTemp  As New Recordset
    Dim szSql As String
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    szSql = "SELECT bus_type_name FROM Bus_type_code WHERE bus_type=" & BusType & ""
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 0 Then
        GetBusTypeName = ""
    Else
        GetBusTypeName = rsTemp("bus_type_name")
    End If
    Set rsTemp = Nothing
    Set oDb = Nothing
End Function

'得到所有车次
Public Function GetAllBus(BusDate As Date, CheckGateID As String) As String()
    AssertActiveUserValid m_oActiveUser, cnErroeStartNo
    
    Dim oDb As New RTConnection
    Dim rsTemp  As New Recordset
    Dim szSql As String
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    
    Dim aszBusID() As String
    If DateDiff("d", BusDate, Now) <= 0 Then '如果查询的是环境中的车次则使用工作表
        szSql = "SELECT e.bus_id" & _
            " FROM Work_env_bus_info e,Bus_check_lst b " & _
            " WHERE e.bus_id*=b.bus_id AND e.bus_date*=b.bus_date AND " & DBDateCompare(BusDate, "e.bus_date") & _
            " AND b.check_gate_id='" & CheckGateID & "'"
'        szSql = "SELECT e.bus_id " & _
'            "FROM Work_env_bus_info e,Vehicle_info v,Company_info c,owner_info o,Bus_check_lst b " & _
'            "WHERE e.vehicle_id = v.vehicle_id  AND e.transport_company_id=c.transport_company_id  AND e.owner_id=o.owner_id " & _
'            "AND e.bus_id*=b.bus_id AND e.bus_date*=b.bus_date AND " & DBDateCompare(BusDate, "e.bus_date") & " " & _
'            "AND check_gate_id='" & CheckGateID & "'" & " AND b.check_gate_id='" & CheckGateID & "'"
    Else
        szSql = "SELECT e.bus_id" & _
            " FROM Env_bus_info e,Bus_check_lst b " & _
            " WHERE e.bus_id*=b.bus_id AND e.bus_date*=b.bus_date AND " & DBDateCompare(BusDate, "e.bus_date") & _
            " AND b.check_gate_id='" & CheckGateID & "'"
'        szSql = "SELECT e.bus_id" & _
'            "FROM Env_bus_info e,Vehicle_info v,Company_info c,owner_info o,Bus_check_lst b " & _
'            "WHERE e.vehicle_id = v.vehicle_id  AND e.transport_company_id=c.transport_company_id  AND e.owner_id=o.owner_id " & _
'            "AND e.bus_id*=b.bus_id AND e.bus_date*=b.bus_date AND " & DBDateCompare(BusDate, "e.bus_date") & " " & _
'            "AND check_gate_id='" & CheckGateID & "'" & " AND b.check_gate_id='" & CheckGateID & "'"
    End If
    
    Set rsTemp = oDb.Execute(szSql, , -1)
    If rsTemp.RecordCount = 0 Then '无记录
        'ShowError ERR_ChkTkResultIsNull
        Exit Function
    End If
    Dim nLoop As Integer
    ReDim aszBusID(1 To rsTemp.RecordCount)
    For nLoop = 1 To rsTemp.RecordCount
        aszBusID(nLoop) = rsTemp.Fields("bus_id")
        rsTemp.MoveNext
    Next nLoop
    GetAllBus = aszBusID
End Function

'得到符合指定条件的检票车次
Public Function GetCheckBus(Optional BusID As String = "", Optional BusDate As Date = cszEmptyDateStr, Optional CheckGateID As String = "") As String()
    AssertActiveUserValid m_oActiveUser, cnErroeStartNo
    
    Dim oDb As New RTConnection
    Dim rsTemp  As New Recordset
    Dim aszBus() As String
    Dim szSql As String
    Dim szCondition As String
    
    BusID = Trim(BusID)
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    
    szSql = "SELECT bus_id,bus_serial_no,bus_date,check_gate_id " & _
            " FROM Bus_check_lst" & _
            " WHERE 1=1"
    
    If BusID <> "" Then szCondition = szCondition & " AND bus_id LIKE '" & Trim(BusID) & "%'"
    If Not DBDateIsEmpty(BusDate) Then szCondition = szCondition & " AND bus_date='" & Format(BusDate, "YYYY-MM-DD") & "'"
    If CheckGateID <> "" Then szCondition = szCondition & " AND check_gate_id='" & CheckGateID & "'"
    szSql = szSql & szCondition
    Set rsTemp = oDb.Execute(szSql, , -1)
    If rsTemp.RecordCount > 0 Then '无记录
        ReDim aszBus(1 To rsTemp.RecordCount, 1 To 4)
        Dim i As Long
        For i = 1 To rsTemp.RecordCount
            aszBus(i, 1) = FormatDbValue(rsTemp!bus_id)
            aszBus(i, 2) = Trim(Str(FormatDbValue(rsTemp!bus_serial_no)))
            aszBus(i, 3) = Format(rsTemp!bus_date, "YYYY-MM-DD")
            aszBus(i, 4) = FormatDbValue(rsTemp!check_gate_id)
            rsTemp.MoveNext
        Next i
    End If
    GetCheckBus = aszBus
End Function

'得到检票口状态
Public Function GetCheckGateStatus() As ECheckGateStatus
    AssertActiveUserValid m_oActiveUser, cnErroeStartNo
    
    Dim oDb As New RTConnection
    Dim rsTemp  As New Recordset
    Dim szSql As String
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    
    If Len(m_sCheckGateNo) = 0 Then '未指定检票口
        ShowError ERR_ChkTkCheckGateIDNotAppointed
        Exit Function
    End If
    szSql = "SELECT c.bus_id FROM Bus_check_lst c " & _
            " WHERE  (c.check_start_time>c.check_end_time OR " & DBDateCompare(cszEmptyDateStr, "c.check_end_time", "=") & ")" & _
            " AND c.check_gate_id='" & m_sCheckGateNo & "' " & _
            " AND " & DBDateCompare(Date, "c.bus_date")
    Set rsTemp = oDb.Execute(szSql, , -1)
    If rsTemp.RecordCount = 0 Then
        GetCheckGateStatus = ECheckGateStatus.ST_CheckGateWait
    Else
        GetCheckGateStatus = ECheckGateStatus.ST_CheckGateChecking
    End If
    
End Function

'得到正在检票的车次
Public Function GetCheckingBus() As TBusCheckInfo()
    AssertActiveUserValid m_oActiveUser, cnErroeStartNo
    
    Dim oDb As New RTConnection
    Dim rsTemp  As New Recordset
    Dim szSql As String
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    
    Dim aszBus() As TBusCheckInfo
    
    If Len(m_sCheckGateNo) = 0 Then '未指定检票口
        ShowError ERR_ChkTkCheckGateIDNotAppointed
        Exit Function
    End If
        
    szSql = "SELECT c.bus_date,c.bus_id,c.bus_serial_no,c.vehicle_id,c.bus_start_time,c.check_start_time,c.check_end_time, " _
        & " c.checker,c.transport_company_id,c.split_company_id,c.owner_id" _
        & " FROM Bus_check_lst c" _
        & " WHERE (c.check_start_time>c.check_end_time OR " & DBDateCompare(cszEmptyDateStr, "c.check_end_time", "=") & ") " _
        & " AND c.check_gate_id='" & m_sCheckGateNo & "' AND " & DBDateCompare(Date, "c.bus_date")
        
    Set rsTemp = oDb.Execute(szSql, , -1)
    If rsTemp.RecordCount = 0 Then
        'ShowError ERR_ChkTkResultIsNull
        Exit Function
    End If
    ReDim aszBus(1 To rsTemp.RecordCount)
    Dim nLoop As Integer
    For nLoop = 1 To rsTemp.RecordCount
        aszBus(nLoop).dtDate = FormatDbValue(rsTemp.Fields("bus_date"))
        aszBus(nLoop).szBusID = FormatDbValue(rsTemp.Fields("bus_id"))
        aszBus(nLoop).nBusSerialNo = rsTemp.Fields("bus_serial_no")
        aszBus(nLoop).szVehicleID = FormatDbValue(rsTemp.Fields("vehicle_id"))
        aszBus(nLoop).dtStartupTime = rsTemp.Fields("bus_start_time")
        aszBus(nLoop).dtBeginCheckTime = rsTemp.Fields("check_start_time")
        aszBus(nLoop).dtEndCheckTime = rsTemp.Fields("check_end_time")
        aszBus(nLoop).szChecker = FormatDbValue(rsTemp.Fields("checker"))
        aszBus(nLoop).szCompanyID = FormatDbValue(rsTemp.Fields("transport_company_id"))
        aszBus(nLoop).szSplitCompanyID = FormatDbValue(rsTemp.Fields("split_company_id"))
        aszBus(nLoop).szOwnerID = FormatDbValue(rsTemp.Fields("owner_id"))
'        aszBus(nLoop).szCheckSheet = m_sCheckGateNo
        aszBus(nLoop).szCheckGateID = m_sCheckGateNo
        aszBus(nLoop).szCheckGateName = m_sCheckGateName
        rsTemp.MoveNext
    Next nLoop
    GetCheckingBus = aszBus

End Function

'得到下一个将要检票的车次
Public Function GetNextScrollNo(BusID As String, Optional BusDate As Date = cszEmptyDateStr) As Integer
    'AssertActiveUserValid m_oActiveUser, cnErroeStartNo
    
    Dim oDb As New RTConnection
    Dim rsTemp  As New Recordset
    Dim szSql As String
    If Len(m_sCheckGateNo) = 0 Then '未指定检票口
        ShowError ERR_ChkTkCheckGateIDNotAppointed
        Exit Function
    End If
    
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    
    If BusDate = cszEmptyDateStr Then
        BusDate = Date
    End If
    If DateDiff("d", Date, BusDate) >= 0 Then
        szSql = "SELECT bus_type FROM Work_env_bus_info " & _
                "WHERE bus_id='" & BusID & "' AND " & DBDateCompare(BusDate, "bus_date")
    Else
        szSql = "SELECT bus_type FROM Env_bus_info " & _
                "WHERE bus_id='" & BusID & "' AND " & DBDateCompare(BusDate, "bus_date")
    End If
    Set rsTemp = oDb.Execute(szSql, , -1)
    
    If rsTemp.RecordCount > 0 Then
        If rsTemp.Fields("bus_type") <> TP_ScrollBus Then
            ShowError ERR_ChkTkBusIsNotFlowBus '指定的车次不是流水车次
            Exit Function
        End If
    Else
        ShowError ERR_ChkTkFlowBusNotExist '指定的流水车次不存在
        Exit Function
    End If
    '得到检票车次明细表中该车次的最大序号
    szSql = "SELECT MAX(bus_serial_no)+1 FROM Bus_check_lst " & _
            "WHERE bus_id='" & BusID & "' AND " & DBDateCompare(BusDate, "bus_date") _
            & " AND check_gate_id='" & m_sCheckGateNo & "'"
    Set rsTemp = oDb.Execute(szSql, , -1)
    If IsNull(rsTemp.Fields(0)) Then
        GetNextScrollNo = 1
    Else
        GetNextScrollNo = rsTemp.Fields(0)
    End If

End Function

'得到车次检票信息
Public Function GetBusCheckInfo(BusDate As Date, BusID As String, BusSerialNo As Integer) As TBusCheckInfo()
    AssertActiveUserValid m_oActiveUser, cnErroeStartNo
    
'    If Len(m_sCheckGateNo) = 0 Then '未指定检票口
'        ShowError ERR_ChkTkCheckGateIDNotAppointed
'        Exit Function
'    End If
    
    Dim oDb As New RTConnection
    Dim rsTemp  As New Recordset
    Dim szSql As String
    Dim i, nCount As Integer
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    
    Dim uBus() As TBusCheckInfo
    
    szSql = "SELECT  c.bus_date,c.bus_id,c.bus_serial_no,c.vehicle_id,c.bus_start_time,c.check_start_time,c.check_end_time " _
        & " ,c.checker,c.transport_company_id,c.split_company_id,c.owner_id " _
        & " FROM Bus_check_lst c " _
        & " WHERE " & DBDateCompare(BusDate, "c.bus_date") & " AND c.bus_id='" & BusID & "' AND c.bus_serial_no=" & BusSerialNo _
'        & " AND c.check_gate_id='" & m_sCheckGateNo & "'"

    
    If m_sCheckGateNo <> "" Then szSql = szSql & " AND check_gate_id='" & m_sCheckGateNo & "'"
    
    szSql = szSql & " ORDER BY c.check_start_time "
    
    Set rsTemp = oDb.Execute(szSql, , -1)
    If rsTemp.RecordCount = 0 Then
        ShowError ERR_ChkTkBusNotExist '指定的车次不存在
        Exit Function
    End If
    nCount = rsTemp.RecordCount
    ReDim uBus(1 To nCount)
    
    For i = 1 To rsTemp.RecordCount
        uBus(i).dtDate = FormatDbValue(rsTemp.Fields("bus_date"))
        uBus(i).szBusID = FormatDbValue(rsTemp.Fields("bus_id"))
        uBus(i).nBusSerialNo = FormatDbValue(rsTemp.Fields("bus_serial_no"))
        uBus(i).szVehicleID = FormatDbValue(rsTemp.Fields("vehicle_id"))
        uBus(i).dtStartupTime = FormatDbValue(rsTemp.Fields("bus_start_time"))
        uBus(i).dtBeginCheckTime = FormatDbValue(rsTemp.Fields("check_start_time"))
        uBus(i).dtEndCheckTime = FormatDbValue(rsTemp.Fields("check_end_time"))
        uBus(i).szChecker = FormatDbValue(rsTemp.Fields("checker"))
        uBus(i).szCompanyID = FormatDbValue(rsTemp.Fields("transport_company_id"))
        uBus(i).szSplitCompanyID = FormatDbValue(rsTemp.Fields("split_company_id"))
        uBus(i).szOwnerID = FormatDbValue(rsTemp.Fields("owner_id"))
        uBus(i).szCheckSheet = ""
        rsTemp.MoveNext
    Next i
    
    szSql = "SELECT  s.check_sheet_id, c.check_gate_id, t.check_gate_name FROM Bus_check_lst c,sheet_lst s, checkgate_info t " & _
            "WHERE c.bus_date=s.bus_date AND c.bus_start_time=s.bus_start_time AND c.bus_id=s.bus_id " & _
            " AND c.check_gate_id = s.check_gate_id AND c.bus_serial_no=s.bus_serial_no AND " & DBDateCompare(BusDate, "c.bus_date") & _
            " AND c.bus_id='" & BusID & "' AND c.bus_serial_no=" & Str(BusSerialNo) & " AND s.valid_mark=1 AND t.check_gate_id = c.check_gate_id "
    
    szSql = szSql & " ORDER BY c.check_start_time "
    
    
    
    If m_sCheckGateNo <> "" Then szSql = szSql & " AND c.check_gate_id='" & m_sCheckGateNo & "'"
    
    Set rsTemp = oDb.Execute(szSql, , -1)
    For i = 1 To rsTemp.RecordCount
        If i <= nCount Then
            uBus(i).szCheckSheet = FormatDbValue(rsTemp.Fields("check_sheet_id"))
            uBus(i).szCheckGateID = FormatDbValue(rsTemp.Fields("check_gate_id"))
            uBus(i).szCheckGateName = FormatDbValue(rsTemp.Fields("check_gate_name"))
        End If
       rsTemp.MoveNext
    Next i
       
    GetBusCheckInfo = uBus
    
End Function


Public Function GetBusCheckInfoRS(BusDate As Date, BusID As String, Optional BusSerialNo As Integer = -1) As Recordset
    AssertActiveUserValid m_oActiveUser, cnErroeStartNo
    
    
    Dim oDb As New RTConnection
    Dim rsTemp  As New Recordset
    Dim szSql As String
    Dim i, nCount As Integer
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    
    
    szSql = "SELECT  c.bus_date,c.bus_id,c.bus_serial_no,c.vehicle_id,c.bus_start_time,c.check_start_time,c.check_end_time " _
        & " ,c.checker,c.transport_company_id,c.split_company_id,c.owner_id , p.transport_company_short_name , v.license_tag_no , o.owner_name" _
        & " FROM Bus_check_lst c , company_info p , vehicle_info v , owner_info o " _
        & " WHERE c.transport_company_id = p.transport_company_id and c.vehicle_id = v.vehicle_id AND c.owner_id = o.owner_id " _
        & " AND  " & DBDateCompare(BusDate, "c.bus_date") & " AND c.bus_id='" & BusID & "' "
    
    
    If m_sCheckGateNo <> "" Then szSql = szSql & " AND check_gate_id='" & m_sCheckGateNo & "'"
    
    If BusSerialNo <> -1 Then
        szSql = szSql & " and c.bus_serial_no = " & TransFieldValueToString(BusSerialNo)
    End If
    
    Set rsTemp = oDb.Execute(szSql, , -1)
    
    Set GetBusCheckInfoRS = rsTemp
    
    
    
End Function



'生成路单
Public Sub MakeCheckSheet(BusDate As Date, BusID As String, BusSerialNo As Integer, CheckSheetNo As String)
    AssertActiveUserValid m_oActiveUser, cnErroeStartNo
    'AssertHaveRight m_oActiveUser, RIGHT_ChkTkMakeCheckSheet
    If Len(m_sCheckGateNo) = 0 Then '未指定检票口
        ShowError ERR_ChkTkCheckGateIDNotAppointed
        Exit Sub
    End If
    
    Dim oDb As New RTConnection
    Dim rsTemp  As New Recordset
    Dim rsTemp2 As New Recordset
    Dim nLoop As Integer
    Dim szSql As String
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    
    Dim uCheckSheetInfo As TCheckSheetInfo
    Dim uCheckSheetStationInfo() As TCheckSheetStationInfoEx
    
    If Len(CheckSheetNo) = 0 Then
        ShowError ERR_ChkTkCheckSheetNoIsNull '路单号为空
        Exit Sub
    End If
    '
    szSql = "SELECT valid_mark FROM sheet_lst WHERE check_sheet_id='" & CheckSheetNo & "'"
    Set rsTemp = oDb.Execute(szSql, , -1)
    If rsTemp.RecordCount > 0 Then '已有相同路单号
        ShowError ERR_ChkTkCheckSheetIDAlreadyExist
        Exit Sub
    End If
    '查询检票车次信息是否已存在
    szSql = "SELECT * FROM Bus_check_lst " & _
            "WHERE bus_id='" & Trim(BusID) & "' AND bus_serial_no=" & BusSerialNo & " AND " & DBDateCompare(BusDate, "bus_date") _
            & " AND check_gate_id='" & m_sCheckGateNo & "'"
    Set rsTemp = oDb.Execute(szSql, , -1)
    If rsTemp.RecordCount = 0 Then '指定的车次还未检票
        ShowError ERR_ChkTkBusIsNotChecked
        Exit Sub
    End If
    uCheckSheetInfo.szCheckSheet = Trim(CheckSheetNo)
    uCheckSheetInfo.szBusID = Trim(BusID)
    uCheckSheetInfo.nBusSerialNo = BusSerialNo
    uCheckSheetInfo.dtDate = BusDate
    uCheckSheetInfo.szVehicleID = rsTemp!vehicle_id
    uCheckSheetInfo.szCompanyID = rsTemp!transport_company_id
    uCheckSheetInfo.szSplitCompanyID = rsTemp!split_company_id
    uCheckSheetInfo.szRouteID = rsTemp!route_id
    uCheckSheetInfo.szVehicleModelID = rsTemp!vehicle_type_code
    uCheckSheetInfo.szOwnerID = rsTemp!owner_id
    uCheckSheetInfo.dtStartupTime = rsTemp!bus_start_time
    uCheckSheetInfo.szMakeSheetUser = rsTemp!Checker
    uCheckSheetInfo.dtMakeSheetDateTime = Now
    uCheckSheetInfo.szCheckGateID = m_sCheckGateNo
    uCheckSheetInfo.bValid = True
    '汇总检票站点
    szSql = "SELECT a.end_station_id,a.ticket_type,SUM(1) AS totalquantity,SUM(a.ticket_price) AS total_price,MAX(a.status) AS check_status" _
            & " FROM check_ticket_lst a" _
            & " WHERE a.bus_date='" & ToDBDate(BusDate) & "' AND a.bus_id='" & Trim(BusID) & "' AND a.bus_serial_no=" & BusSerialNo _
            & " AND a.check_gate_id='" & m_sCheckGateNo & "'" _
            & " GROUP BY end_station_id,ticket_type order by a.end_station_id "
    Set rsTemp = oDb.Execute(szSql, , -1)
    If rsTemp.RecordCount > 0 Then
        ReDim uCheckSheetStationInfo(1 To rsTemp.RecordCount)
        
        If rsTemp.RecordCount > 0 Then
            For nLoop = 1 To rsTemp.RecordCount
                uCheckSheetStationInfo(nLoop).nManCount = rsTemp!totalquantity
                uCheckSheetStationInfo(nLoop).nTicketType = rsTemp!ticket_type
                uCheckSheetStationInfo(nLoop).nCheckStatus = rsTemp!check_status
'                uCheckSheetStationInfo(nLoop).sgMileage= rstemp!mileage
                uCheckSheetStationInfo(nLoop).sgTicketPrice = rsTemp!total_price
                uCheckSheetStationInfo(nLoop).szCheckSheet = CheckSheetNo
                uCheckSheetStationInfo(nLoop).szStationID = rsTemp!end_station_id
                rsTemp.MoveNext
            Next nLoop
        End If
    End If
    MakeCheckSheetFromMemo uCheckSheetInfo, uCheckSheetStationInfo
    
End Sub



'写入路单数据
Public Sub MakeCheckSheetFromMemo(CheckSheetInfo As TCheckSheetInfo, CheckSheetStationInfo() As TCheckSheetStationInfoEx)
    AssertActiveUserValid m_oActiveUser, cnErroeStartNo

    Dim oDb As New RTConnection
    Dim rsTemp  As New Recordset
    Dim szSql As String
    Dim nLoop As Integer
    Dim BusID As String
    Dim BusDate As Date
    
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    
    BusID = CheckSheetInfo.szBusID
    BusDate = CheckSheetInfo.dtDate
    
    'szSql = "SELECT * FROM Bus_check_lst WHERE bus_id='" & BusID & "' AND " & DBDateCompare(BusDate, "bus_date") & " " & _
    '        "AND check_start_time<check_end_time" ' AND " & DBTimeCompareEX(Now, "check_end_time", ">=")
    'Set rsTemp = oDb.Execute(szSql, , -1)
    'If rsTemp.RecordCount = 0 Then '指定的车次还未停检
    '    ShowError ERR_ChkTkBusIsNotChecked
    '    Exit Sub
    'End If
    'szSql = "SELECT check_sheet_id FROM sheet_lst WHERE check_sheet_id='" & CheckSheetInfo.szCheckSheet & "'"
    'Set rsTemp = oDb.Execute(szSql, , -1)
    'If rsTemp.RecordCount > 0 Then '指定的路单号已存在
    '    ShowError ERR_ChkTkCheckSheetIDAlreadyExist
    '    Exit Sub
    'End If
    On Error GoTo HereErr
    oDb.BeginTrans
    '写入路单信息
    szSql = "INSERT sheet_lst(check_gate_id,check_sheet_id,bus_id,bus_serial_no,bus_date,vehicle_id,transport_company_id,split_company_id,route_id,vehicle_type_code,owner_id,bus_start_time,valid_mark,checker,sheet_make_date) VALUES( "
    szSql = szSql & "'" & CheckSheetInfo.szCheckGateID & "',"
    szSql = szSql & "'" & CheckSheetInfo.szCheckSheet & "',"
    szSql = szSql & "'" & CheckSheetInfo.szBusID & "',"
    szSql = szSql & Str(CheckSheetInfo.nBusSerialNo) & ","
    szSql = szSql & "'" & Format(CheckSheetInfo.dtDate, "YYYY-MM-DD hh:mm:ss") & "',"
    szSql = szSql & "'" & CheckSheetInfo.szVehicleID & "',"
    szSql = szSql & "'" & CheckSheetInfo.szCompanyID & "',"
    szSql = szSql & "'" & CheckSheetInfo.szSplitCompanyID & "',"
    szSql = szSql & "'" & CheckSheetInfo.szRouteID & "',"
    szSql = szSql & "'" & CheckSheetInfo.szVehicleModelID & "',"
    szSql = szSql & "'" & CheckSheetInfo.szOwnerID & "',"
    szSql = szSql & "'" & ToDBDateTime(CheckSheetInfo.dtStartupTime) & "',"
    szSql = szSql & Str(IIf(CheckSheetInfo.bValid, 1, 0)) & ","
    szSql = szSql & "'" & CheckSheetInfo.szMakeSheetUser & "',"
    szSql = szSql & "'" & ToDBDateTime(CheckSheetInfo.dtMakeSheetDateTime) & "')"
    oDb.Execute szSql
    
    For nLoop = 1 To ArrayLength(CheckSheetStationInfo)
        szSql = "INSERT sheet_station_lst(check_sheet_id,station_id,ticket_type,quantity,ticket_price,check_type) VALUES( "
        szSql = szSql & "'" & CheckSheetStationInfo(nLoop).szCheckSheet & "',"
        szSql = szSql & "'" & CheckSheetStationInfo(nLoop).szStationID & "',"
        szSql = szSql & CheckSheetStationInfo(nLoop).nTicketType & ","
        szSql = szSql & CheckSheetStationInfo(nLoop).nManCount & ","
        szSql = szSql & CheckSheetStationInfo(nLoop).sgTicketPrice & ","
        szSql = szSql & CheckSheetStationInfo(nLoop).nCheckStatus & ")"
        oDb.Execute szSql
    Next
    oDb.CommitTrans
    Exit Sub
HereErr: oDb.RollbackTrans
    err.Raise err.Number, err.Source, err.Description
    Set oDb = Nothing
    Set rsTemp = Nothing
    'WriteOperateLog m_oActiveUser, RIGHT_ChkTkMakeCheckSheet

End Sub

'更改路单号
Public Sub ChangeCheckSheetNo(OldCheckSheetNo As String, NewCheckSheetNo As String)
    AssertActiveUserValid m_oActiveUser, cnErroeStartNo
    'AssertHaveRight m_oActiveUser, RIGHT_ChkTkMakeCheckSheet
    Dim oDb As New RTConnection
    Dim rsTemp  As New Recordset
    Dim szSql As String
    Dim ExeSQL As String
    Dim szSet As String
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    
    If Len(NewCheckSheetNo) = 0 Then
        ShowError ERR_ChkTkCheckSheetNoIsNull '路单号为空
        Exit Sub
    End If
    
    szSql = "SELECT check_sheet_id FROM sheet_lst WHERE check_sheet_id='" & NewCheckSheetNo & "'"
    Set rsTemp = oDb.Execute(szSql, , -1)
    If rsTemp.RecordCount > 0 Then '指定的新路单已存在
        ShowError ERR_ChkTkCheckSheetIDAlreadyExist
        Exit Sub
    End If
    
    szSql = "SELECT valid_mark FROM sheet_lst WHERE check_sheet_id='" & OldCheckSheetNo & "'"
    Set rsTemp = oDb.Execute(szSql, , -1)
    If rsTemp.RecordCount = 0 Then '指定的原路单不存在
        ShowError ERR_ChkTkCheckSheetIDNotExist
        Exit Sub
    End If
    If rsTemp!valid_mark <> 1 Then
        ShowError ERR_ChkTkCheckSheetIDIsNotValid
        Exit Sub
    End If
On Error GoTo here
    oDb.BeginTrans
    szSql = "UPDATE sheet_lst SET valid_mark=0 WHERE check_sheet_id='" & OldCheckSheetNo & "'"
    oDb.Execute szSql
    szSql = ""
    szSet = ""
    szSql = "INSERT INTO sheet_lst ("
    szSql = szSql & "check_gate_id,"
    szSql = szSql & "check_sheet_id,"
    szSet = szSet & "check_gate_id,"
    szSet = szSet & "'" & Trim(NewCheckSheetNo) & "' AS check_sheet_id,"
    szSql = szSql & " bus_date,bus_id,bus_serial_no,vehicle_id,transport_company_id,split_company_id,route_id,vehicle_type_code,owner_id,bus_start_time,valid_mark,checker,sheet_make_date,settlement_sheet_id,protocol_id,settlement_time,settlement_status "
    szSet = szSet & " bus_date,bus_id,bus_serial_no,vehicle_id,transport_company_id,split_company_id,route_id,vehicle_type_code,owner_id,bus_start_time,1 AS valid_mark," & TransFieldValueToString(m_oActiveUser.UserID) & " ,sheet_make_date,settlement_sheet_id,protocol_id,settlement_time,settlement_status "
    ExeSQL = szSql & ") SELECT " & szSet _
    & " FROM sheet_lst WHERE check_sheet_id='" & OldCheckSheetNo & "'"
    oDb.Execute ExeSQL
    
    '2路单站点信息
    szSql = ""
    szSet = ""
    szSql = "INSERT INTO sheet_station_lst ("
    szSql = szSql & "check_sheet_id,"
    szSet = szSet & "'" & Trim(NewCheckSheetNo) & "' AS check_sheet_id,"
    szSql = szSql & "station_id,ticket_type,quantity,ticket_price,check_type"
    szSet = szSet & "station_id,ticket_type,quantity,ticket_price,check_type"
    ExeSQL = szSql & ") SELECT " & szSet _
    & " FROM sheet_station_lst WHERE check_sheet_id='" & OldCheckSheetNo & "'"
    oDb.Execute ExeSQL
    oDb.CommitTrans
Exit Sub
here:
    oDb.RollbackTrans
    ShowError err.Number
End Sub

'作废路单
Public Sub CancelCheckSheet(poDb As RTConnection, CheckSheetID As String)
    AssertActiveUserValid m_oActiveUser, cnErroeStartNo
    'AssertHaveRight m_oActiveUser, RIGHT_ChkTkCancelCheckSheet
    
'    Dim poDb As New RTConnection
    Dim rsTemp  As New Recordset
    Dim szSql As String
    
    poDb.ConnectionString = GetConnectionStr(cszSystemMan)
    
    szSql = "SELECT valid_mark FROM sheet_lst WHERE check_sheet_id='" & CheckSheetID & "'"
    Set rsTemp = poDb.Execute(szSql, , -1)
    If rsTemp.RecordCount = 0 Then '路单不存在
        ShowError ERR_ChkTkCheckSheetIDAlreadyExist
        Exit Sub
    End If
    If rsTemp.Fields("valid_mark") = 0 Then '指定的路单已经是作废了的
        ShowError ERR_ChkTkCheckSheetIDIsNotValid
        Exit Sub
    End If
    
    szSql = "UPDATE sheet_lst SET valid_mark=0 WHERE check_sheet_id='" & CheckSheetID & "'"
    poDb.Execute szSql
    'WriteOperateLog m_oActiveUser, RIGHT_ChkTkCancelCheckSheet
End Sub

'得到路单信息
Public Function GetCheckSheetInfo(CheckSheetID As String) As TCheckSheetInfo
    AssertActiveUserValid m_oActiveUser, cnErroeStartNo
    
    Dim oDb As New RTConnection
    Dim rsTemp  As New Recordset
    Dim szSql As String
    
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    
    Dim uCheckSheetInfo As TCheckSheetInfo
    
    szSql = "SELECT a.check_sheet_id,a.bus_id,a.bus_serial_no,a.bus_date,a.vehicle_id,a.transport_company_id,a.split_company_id" _
        & " ,a.split_company_id,a.route_id,a.vehicle_type_code,a.owner_id,a.bus_start_time,a.valid_mark,a.settlement_time" _
        & " ,a.settlement_sheet_id,a.checker,a.sheet_make_date,a.check_gate_id,c.check_gate_name FROM sheet_lst a , checkgate_info c " _
        & " WHERE a.check_gate_id = c.check_gate_id AND a.check_sheet_id='" & CheckSheetID & "'"
    Set rsTemp = oDb.Execute(szSql, , -1)
    If rsTemp.RecordCount = 0 Then '路单不存在
'        ShowError ERR_ChkTkCheckSheetIDNotExist
        Exit Function
    End If
    uCheckSheetInfo.bValid = rsTemp.Fields("valid_mark")
    uCheckSheetInfo.dtDate = rsTemp.Fields("bus_date")
    uCheckSheetInfo.dtMakeSheetDateTime = IIf(IsNull(rsTemp.Fields("sheet_make_date")), cdtEmptyDate, rsTemp.Fields("sheet_make_date"))
    uCheckSheetInfo.dtSettleDateTime = IIf(IsNull(rsTemp.Fields("settlement_time")), cdtEmptyDate, rsTemp.Fields("settlement_time"))
    uCheckSheetInfo.dtStartupTime = rsTemp.Fields("bus_start_time")
    uCheckSheetInfo.nBusSerialNo = rsTemp.Fields("bus_serial_no")
    uCheckSheetInfo.nSettleSheetSerialNo = 0
    uCheckSheetInfo.szBusID = rsTemp.Fields("bus_id")
    uCheckSheetInfo.szCheckSheet = rsTemp.Fields("check_sheet_id")
    uCheckSheetInfo.szCompanyID = rsTemp.Fields("transport_company_id")
    uCheckSheetInfo.szSplitCompanyID = rsTemp.Fields("split_company_id")
    uCheckSheetInfo.szMakeSheetUser = rsTemp.Fields("checker")
    uCheckSheetInfo.szOwnerID = rsTemp.Fields("owner_id")
    uCheckSheetInfo.szRouteID = rsTemp.Fields("route_id")
    uCheckSheetInfo.szSettleSheetID = IIf(IsNull(rsTemp.Fields("settlement_sheet_id")), "", rsTemp.Fields("settlement_sheet_id"))
    uCheckSheetInfo.szVehicleID = rsTemp.Fields("vehicle_id")
    uCheckSheetInfo.szVehicleModelID = rsTemp.Fields("vehicle_type_code")
    uCheckSheetInfo.szCheckGateID = rsTemp.Fields("check_gate_id")
    uCheckSheetInfo.szCheckGateName = FormatDbValue(rsTemp.Fields("check_gate_name"))
    GetCheckSheetInfo = uCheckSheetInfo

End Function

'得到车次路单信息
'(为保持兼容，此接口暂时保留)
'Public Function GetBusCheckSheet(BusDate As Date, BusID As String, BusSerialNo As Integer) As TCheckSheetInfo
'    AssertActiveUserValid m_oActiveUser, cnErroeStartNo
'
'    Dim odb As New RTConnection
'    Dim rsTemp  As New Recordset
'    Dim szSql As String
'
'    odb.ConnectionString = GetConnectionStr(cszSystemMan)
'
'    Dim uCheckSheetInfo As TCheckSheetInfo
'
'    szSql = ""
'    szSql = "SELECT check_sheet_id,bus_id,bus_serial_no,bus_date,vehicle_id,transport_company_id,split_company_id,route_id,vehicle_type_code,owner_id,bus_start_time,valid_mark,settlement_time,settlement_sheet_id,checker,sheet_make_date FROM sheet_lst " & _
'            "WHERE " & DBDateCompare(BusDate, "bus_date") & "AND bus_id='" & BusID & "' AND bus_serial_no=" & Str(BusSerialNo) & " AND valid_mark=1"
'    Set rsTemp = odb.Execute(szSql, , -1)
'    If rsTemp.RecordCount = 0 Then '指定的车次不存在
'        'ShowError ERR_ChkTkBusNotExist
'        Exit Function
'    End If
'    uCheckSheetInfo.bValid = rsTemp.Fields("valid_mark")
'    uCheckSheetInfo.dtDate = rsTemp.Fields("bus_date")
'    uCheckSheetInfo.dtMakeSheetDateTime = IIf(IsNull(rsTemp.Fields("sheet_make_date")), cdtEmptyDate, rsTemp.Fields("sheet_make_date"))
'    uCheckSheetInfo.dtSettleDateTime = IIf(IsNull(rsTemp.Fields("settlement_time")), cdtEmptyDate, rsTemp.Fields("settlement_time"))
'    uCheckSheetInfo.dtStartupTime = rsTemp.Fields("bus_start_time")
'    uCheckSheetInfo.nBusSerialNo = rsTemp.Fields("bus_serial_no")
'    uCheckSheetInfo.nSettleSheetSerialNo = 0
'    uCheckSheetInfo.szBusID = rsTemp.Fields("bus_id")
'    uCheckSheetInfo.szCheckSheet = rsTemp.Fields("check_sheet_id")
'    uCheckSheetInfo.szCompanyID = rsTemp.Fields("transport_company_id")
'    uCheckSheetInfo.szSplitCompanyID = rsTemp.Fields("split_company_id")
'    uCheckSheetInfo.szMakeSheetUser = rsTemp.Fields("checker")
'    uCheckSheetInfo.szOwnerID = rsTemp.Fields("owner_id")
'    uCheckSheetInfo.szRouteID = rsTemp.Fields("route_id")
'    uCheckSheetInfo.szSettleSheetID = IIf(IsNull(rsTemp.Fields("settlement_sheet_id")), "", rsTemp.Fields("settlement_sheet_id"))
'    uCheckSheetInfo.szVehicleID = rsTemp.Fields("vehicle_id")
'    uCheckSheetInfo.szVehicleModelID = rsTemp.Fields("vehicle_type_code")
'    'Debug.Print rsTemp!Checker
'
'    GetBusCheckSheet = uCheckSheetInfo
'
'End Function

Public Function GetBusCheckSheet(BusDate As Date, CheckGateID As String) As TCheckSheetInfo()
    AssertActiveUserValid m_oActiveUser, cnErroeStartNo
    
    Dim oDb As New RTConnection
    Dim rsTemp  As New Recordset
    Dim szSql As String
    Dim j, nCount As Integer
    
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    
    Dim uCheckSheetInfo() As TCheckSheetInfo

    szSql = "SELECT check_sheet_id,bus_id,bus_serial_no,bus_date,vehicle_id,transport_company_id,split_company_id,route_id,vehicle_type_code" _
        & " ,owner_id,bus_start_time,valid_mark,settlement_time,settlement_sheet_id,checker,sheet_make_date " _
        & " FROM sheet_lst " _
        & " WHERE " & DBDateCompare(BusDate, "bus_date") & "AND check_gate_id=" & TransFieldValueToString(CheckGateID) _
        & " ORDER BY bus_start_time,bus_id"
    Set rsTemp = oDb.Execute(szSql, , -1)
    nCount = rsTemp.RecordCount
    If nCount = 0 Then '指定的车次不存在
        'ShowError ERR_ChkTkBusNotExist
        Exit Function
    End If
    ReDim uCheckSheetInfo(1 To nCount)
    
    For j = 1 To nCount
        uCheckSheetInfo(j).bValid = rsTemp.Fields("valid_mark")
        uCheckSheetInfo(j).dtDate = rsTemp.Fields("bus_date")
        uCheckSheetInfo(j).dtMakeSheetDateTime = IIf(IsNull(rsTemp.Fields("sheet_make_date")), cdtEmptyDate, rsTemp.Fields("sheet_make_date"))
        uCheckSheetInfo(j).dtSettleDateTime = IIf(IsNull(rsTemp.Fields("settlement_time")), cdtEmptyDate, rsTemp.Fields("settlement_time"))
        uCheckSheetInfo(j).dtStartupTime = rsTemp.Fields("bus_start_time")
        uCheckSheetInfo(j).nBusSerialNo = rsTemp.Fields("bus_serial_no")
        uCheckSheetInfo(j).nSettleSheetSerialNo = 0
        uCheckSheetInfo(j).szBusID = rsTemp.Fields("bus_id")
        uCheckSheetInfo(j).szCheckSheet = rsTemp.Fields("check_sheet_id")
        uCheckSheetInfo(j).szCompanyID = rsTemp.Fields("transport_company_id")
        uCheckSheetInfo(j).szSplitCompanyID = rsTemp.Fields("split_company_id")
        uCheckSheetInfo(j).szMakeSheetUser = rsTemp.Fields("checker")
        uCheckSheetInfo(j).szOwnerID = rsTemp.Fields("owner_id")
        uCheckSheetInfo(j).szRouteID = rsTemp.Fields("route_id")
        uCheckSheetInfo(j).szSettleSheetID = IIf(IsNull(rsTemp.Fields("settlement_sheet_id")), "", rsTemp.Fields("settlement_sheet_id"))
        uCheckSheetInfo(j).szVehicleID = rsTemp.Fields("vehicle_id")
        uCheckSheetInfo(j).szVehicleModelID = rsTemp.Fields("vehicle_type_code")
        rsTemp.MoveNext
    Next
    GetBusCheckSheet = uCheckSheetInfo
    
End Function

'得到车次检票信息
Public Function GetBusCheckTicket(BusDate As Date, BusID As String, BusSerialNo As Integer, Optional CheckGateNo As String) As TCheckedTicketInfo()
    AssertActiveUserValid m_oActiveUser, cnErroeStartNo
'    If Len(m_sCheckGateNo) = 0 Then '未指定检票口
'        ShowError ERR_ChkTkCheckGateIDNotAppointed
'        Exit Function
'    End If
    
    Dim oDb As New RTConnection
    Dim rsTemp  As New Recordset
    Dim szSql As String
    
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    
    Dim uCheckedTicketInfo() As TCheckedTicketInfo
    
    szSql = "SELECT ticket_id,bus_date,bus_serial_no,check_time,bus_id,ticket_price,status,ticket_type FROM check_ticket_lst " & _
            "WHERE " & DBDateCompare(BusDate, "bus_date") & "AND bus_id='" & BusID & "'AND bus_serial_no=" & Str(BusSerialNo)


    If CheckGateNo <> "" Then szSql = szSql & " AND check_gate_id='" & CheckGateNo & "'"
    
    Set rsTemp = oDb.Execute(szSql, , -1)
    If rsTemp.RecordCount = 0 Then '指定的车次不存在
        'ShowError ERR_ChkTkBusNotExist
        Exit Function
    End If
    ReDim uCheckedTicketInfo(1 To rsTemp.RecordCount)
    Dim nLoop As Integer
    For nLoop = 1 To rsTemp.RecordCount
        uCheckedTicketInfo(nLoop).nCheckTicketType = rsTemp.Fields("status")
        uCheckedTicketInfo(nLoop).dtBusDate = rsTemp.Fields("bus_date")
        uCheckedTicketInfo(nLoop).dtCheckTime = rsTemp.Fields("check_time")
        uCheckedTicketInfo(nLoop).nBusSerialNo = rsTemp.Fields("bus_serial_no")
        uCheckedTicketInfo(nLoop).szBusID = rsTemp.Fields("bus_id")
        uCheckedTicketInfo(nLoop).szTicketId = rsTemp.Fields("ticket_id")
        rsTemp.MoveNext
    Next nLoop
    
    GetBusCheckTicket = uCheckedTicketInfo

End Function

'得到路单站点信息
Public Function GetCheckSheetStationInfo(CheckSheetID As String) As TCheckSheetStationInfoEx()
'    AssertActiveUserValid m_oActiveUser, cnErroeStartNo
'
'    Dim oDb As New RTConnection
'    Dim rsTemp  As New Recordset
'    Dim szSql As String
'    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
'
'    Dim uCheckSheetStationInfo() As TCheckSheetStationInfoEx
'
'    szSql = "SELECT tbc.station_id,ticket_type,quantity,ticket_price,check_type,station_name " _
'        & " FROM sheet_station_lst tbc,station_info tbs,Env_bus_station_lst b,sheet_lst c " _
'        & " WHERE tbc.check_sheet_id='" & CheckSheetID & "' AND tbc.station_id=tbs.station_id " _
'        & " AND tbc.station_id=b.station_id AND tbc.check_sheet_id =c.check_sheet_id AND b.bus_id=c.bus_id AND b.bus_date=c.bus_date  " _
'        & " Order by b.mileage"
'    Set rsTemp = oDb.Execute(szSql, , -1)
'    Dim nLoop As Integer
'    If rsTemp.RecordCount = 0 Then '指定的路单不存在
''        ShowError ERR_ChkTkCheckSheetIDNotExist
'        Exit Function
'    End If
'    ReDim uCheckSheetStationInfo(1 To rsTemp.RecordCount)
'    For nLoop = 1 To rsTemp.RecordCount
'        uCheckSheetStationInfo(nLoop).nManCount = rsTemp!quantity
'        uCheckSheetStationInfo(nLoop).nTicketType = rsTemp!ticket_type
'        uCheckSheetStationInfo(nLoop).sgMileage = rsTemp!check_type
'        uCheckSheetStationInfo(nLoop).sgTicketPrice = rsTemp!ticket_price
'        uCheckSheetStationInfo(nLoop).szCheckSheet = CheckSheetID & "[" & rsTemp!station_name & "]"
'        '需要调整，调整地方有 检票管理里路单窗体-->GetCheckSheetInfo
'        uCheckSheetStationInfo(nLoop).szStationID = rsTemp!station_id
'        rsTemp.MoveNext
'    Next nLoop
'
'    GetCheckSheetStationInfo = uCheckSheetStationInfo
'
    AssertActiveUserValid m_oActiveUser, cnErroeStartNo
    
    Dim oDb As New RTConnection
    Dim rsTemp  As New Recordset
    Dim szSql As String
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    
    Dim uCheckSheetStationInfo() As TCheckSheetStationInfoEx
    
    szSql = "SELECT tbc.station_id,ticket_type,quantity,ticket_price,check_type,station_name , b.mileage " _
        & " FROM sheet_station_lst tbc,station_info tbs,Env_bus_station_lst b,sheet_lst c , checkgate_info g " _
        & " WHERE tbc.check_sheet_id='" & CheckSheetID & "' AND tbc.station_id=tbs.station_id " _
        & " AND tbc.station_id=b.station_id AND tbc.check_sheet_id =c.check_sheet_id AND b.bus_id=c.bus_id AND b.bus_date=c.bus_date  " _
        & " AND g.sell_station_id = b.sell_station_id AND g.check_gate_id = c.check_gate_id " _
        & " Order by b.mileage,tbc.ticket_type "
    Set rsTemp = oDb.Execute(szSql, , -1)
    Dim nLoop As Integer
    If rsTemp.RecordCount = 0 Then '指定的路单不存在
'        ShowError ERR_ChkTkCheckSheetIDNotExist
        Exit Function
    End If
    ReDim uCheckSheetStationInfo(1 To rsTemp.RecordCount)
    For nLoop = 1 To rsTemp.RecordCount
        uCheckSheetStationInfo(nLoop).nManCount = rsTemp!quantity
        uCheckSheetStationInfo(nLoop).nTicketType = rsTemp!ticket_type
        uCheckSheetStationInfo(nLoop).nCheckStatus = rsTemp!check_type
        uCheckSheetStationInfo(nLoop).sgMileage = rsTemp!mileage
        uCheckSheetStationInfo(nLoop).sgTicketPrice = rsTemp!ticket_price
        uCheckSheetStationInfo(nLoop).szCheckSheet = CheckSheetID & "[" & rsTemp!station_name & "]"
        '需要调整，调整地方有 检票管理里路单窗体-->GetCheckSheetInfo
        uCheckSheetStationInfo(nLoop).szStationID = rsTemp!station_id
        rsTemp.MoveNext
    Next nLoop
    
    GetCheckSheetStationInfo = uCheckSheetStationInfo
    

End Function

'得到路单站点信息
Public Function GetCheckSheetStationInfoEx(Optional CheckSheetID As String, Optional BusID As String, Optional BusDate As Date, Optional BusSerialNo As String) As TCheckSheetStationInfoEx2()

    AssertActiveUserValid m_oActiveUser, cnErroeStartNo
    
    Dim oDb As New RTConnection
    Dim rsTemp  As New Recordset
    Dim szSql As String
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    
    Dim uCheckSheetStationInfo() As TCheckSheetStationInfoEx2
    
    szSql = "select ctl.end_station_id,station_name,ctl.ticket_type,count(*) as quantity,ctl.ticket_price as single_price,sum(ctl.ticket_price) as ticket_price,mileage " _
        & " FROM check_ticket_lst ctl,station_info si,Env_bus_station_lst e,checkgate_info g ,ticket_sell_lst t,sheet_lst s " _
        & " WHERE s.bus_date=ctl.bus_date AND s.bus_id=ctl.bus_id  "
        
    If BusID <> "" Then szSql = szSql & "  AND ctl.bus_id='" & BusID & "'"
    If BusDate <> "0:00:00" Then szSql = szSql & " AND ctl.bus_date='" & BusDate & "'"
    
    If BusSerialNo <> "" Then szSql = szSql & " AND ctl.bus_serial_no='" & BusSerialNo & "'"
    If CheckSheetID <> "" Then szSql = szSql & "  AND s.check_sheet_id='" & CheckSheetID & "'"
    
        szSql = szSql _
        & " AND si.station_id=ctl.end_station_id AND e.station_id=ctl.end_station_id " _
        & " AND e.bus_id=ctl.bus_id AND e.bus_date=ctl.bus_date AND s.check_gate_id=ctl.check_gate_id  " _
        & " AND g.sell_station_id = e.sell_station_id AND g.check_gate_id = ctl.check_gate_id and t.ticket_id=ctl.ticket_id" _
        & " Group by ctl.ticket_type,ctl.ticket_price,station_name,ctl.end_station_id,mileage " _
        & " Order by mileage "
    Set rsTemp = oDb.Execute(szSql, , -1)
    Dim nLoop As Integer
    If rsTemp.RecordCount = 0 Then '指定的路单不存在
        ReDim uCheckSheetStationInfo(1 To 1)
        GetCheckSheetStationInfoEx = uCheckSheetStationInfo
'        ShowError ERR_ChkTkCheckSheetIDNotExist
        Exit Function
    End If
    ReDim uCheckSheetStationInfo(1 To rsTemp.RecordCount)
    For nLoop = 1 To rsTemp.RecordCount
        uCheckSheetStationInfo(nLoop).nManCount = rsTemp!quantity
        uCheckSheetStationInfo(nLoop).szSinglePrice = rsTemp!single_price
        uCheckSheetStationInfo(nLoop).nTicketType = rsTemp!ticket_type
        'uCheckSheetStationInfo(nLoop).nCheckStatus = rsTemp!check_type
        uCheckSheetStationInfo(nLoop).sgMileage = rsTemp!mileage
        uCheckSheetStationInfo(nLoop).sgTicketPrice = rsTemp!ticket_price
        uCheckSheetStationInfo(nLoop).szCheckSheet = CheckSheetID & "[" & rsTemp!station_name & "]"
        '需要调整，调整地方有 检票管理里路单窗体-->GetCheckSheetInfo
        uCheckSheetStationInfo(nLoop).szStationID = rsTemp!end_station_id
        rsTemp.MoveNext
    Next nLoop
    
    GetCheckSheetStationInfoEx = uCheckSheetStationInfo
    
End Function


'得到车次检票信息
Public Function GetTicketCheckInfo(TicketID As String) As TCheckedTicketInfo
    AssertActiveUserValid m_oActiveUser, cnErroeStartNo
    Dim oDb As New RTConnection
    Dim rsTemp  As New Recordset
    Dim szSql As String
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    Dim uTicketCheckInfo As TCheckedTicketInfo
    szSql = " SELECT t.ticket_id,t.bus_date,t.bus_serial_no,t.check_time,t.bus_id,t.status,t.check_gate_id,b.checker " & _
            " FROM check_ticket_lst t,Bus_check_lst b" & _
            " WHERE t.bus_date = b.bus_date AND t.bus_id = b.bus_id " & _
            " AND t.bus_serial_no=b.bus_serial_no " & _
            " AND t.check_gate_id=b.check_gate_id " & _
            " AND t.ticket_id='" & Trim(TicketID) & "'"
    If m_sCheckGateNo <> "" Then
       szSql = szSql & " AND b.check_gate_id='" & m_sCheckGateNo & "'"
    End If
    Set rsTemp = oDb.Execute(szSql, , -1)
    If rsTemp.RecordCount > 0 Then
        uTicketCheckInfo.dtBusDate = rsTemp.Fields("bus_date")
        uTicketCheckInfo.dtCheckTime = rsTemp.Fields("check_time")
        uTicketCheckInfo.nBusSerialNo = rsTemp.Fields("bus_serial_no")
        uTicketCheckInfo.nCheckTicketType = rsTemp.Fields("status")
        uTicketCheckInfo.szBusID = rsTemp.Fields("bus_id")
        uTicketCheckInfo.szCheckerID = rsTemp.Fields("checker")
        uTicketCheckInfo.szCheckGateID = rsTemp.Fields("check_gate_id")
        uTicketCheckInfo.szTicketId = rsTemp.Fields("ticket_id")
        GetTicketCheckInfo = uTicketCheckInfo
    End If
End Function

'得到车次状态
Public Function GetBusStatus(BusDate As Date, BusID As String, BusSerialNo As Integer, CheckGateID As String) As Integer
    'AssertActiveUserValid m_oActiveUser, cnErroeStartNo
    If Len(m_sCheckGateNo) = 0 Then '未指定检票口
        ShowError ERR_ChkTkCheckGateIDNotAppointed
        Exit Function
    End If
    
    Dim oDb As New RTConnection
    Dim rsTemp  As New Recordset
    Dim szSql As String
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    If DateDiff("d", Date, BusDate) >= 0 Then
        szSql = "SELECT status FROM Work_env_bus_allot_lst"
    Else
        szSql = "SELECT status FROM Env_bus_allot_lst"
    End If
    szSql = szSql & " WHERE bus_id=" & TransFieldValueToString(BusID) & _
                    " AND " & DBDateCompare(BusDate, "bus_date") & _
                    " AND check_gate_id=" & TransFieldValueToString(CheckGateID)
    
    Set rsTemp = oDb.Execute(szSql, , -1)
    If rsTemp.RecordCount = 0 Then
         ShowError ERR_ChkTkBusNotExist
    End If
    GetBusStatus = FormatDbValue(rsTemp.Fields("status"))
    '原有代码
'    If rsTemp.Fields("bus_type") <> EBusType.TP_ScrollBus Then
'        GetBusStatus = rsTemp.Fields("status")
'        Exit Function
'    End If
'    If DateDiff("d", Date, BusDate) >= 0 Then
'    szSql = "SELECT c.check_start_time,c.check_end_time FROM Work_env_bus_info e,Bus_check_lst c WHERE e.bus_date=c.bus_date AND e.bus_id=c.bus_id " & _
'            "AND " & DBDateCompare(BusDate, "e.bus_date") & " AND e.bus_id='" & Trim(BusID) & "' AND c.bus_serial_no = " & Str(BusSerialNo)
'    Else
'    szSql = "SELECT c.check_start_time,c.check_end_time FROM Env_bus_info e,Bus_check_lst c WHERE e.bus_date=c.bus_date AND e.bus_id=c.bus_id " & _
'            "AND " & DBDateCompare(BusDate, "e.bus_date") & " AND e.bus_id='" & Trim(BusID) & "' AND c.bus_serial_no = " & Str(BusSerialNo)
'    End If
'    Set rsTemp = oDb.Execute(szSql, , -1)
'    If rsTemp.RecordCount = 0 Then
'        GetBusStatus = EREBusStatus.ST_BusNormal
'        Exit Function
'    End If
'    If Format(rsTemp.Fields("check_end_time"), "YYYY-MM-DD") = Format(cszEmptyDateStr, "YYYY-MM-DD") Then
'        GetBusStatus = EREBusStatus.ST_BusChecking
'        Exit Function
'    End If
'    If Format(rsTemp.Fields("check_start_time"), "hh:mm:ss") > Format(rsTemp.Fields("check_end_time"), "hh:mm:ss") Then
'        GetBusStatus = EREBusStatus.ST_BusExtraChecking
'        Exit Function
'    End If
'    If Format(rsTemp.Fields("check_start_time"), "hh:mm:ss") < Format(rsTemp.Fields("check_end_time"), "hh:mm:ss") Then
'        GetBusStatus = EREBusStatus.ST_BusStopCheck
'        Exit Function
'    End If
'    If Format(rsTemp.Fields("check_start_time"), "hh:mm:ss") = Format(rsTemp.Fields("check_end_time"), "hh:mm:ss") Then
'        GetBusStatus = EREBusStatus.ST_BusStopCheck
'        Exit Function
'    End If
    
End Function

'得到新的检票员
'Public Function GetNewChecker() As String()
'    AssertActiveUserValid m_oActiveUser, cnErroeStartNo
'
'    Dim oDb As New RTConnection
'    Dim rsTemp  As New Recordset
'    Dim szSql As String
'    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
'
'    Dim aszUserID() As String
'
'    szSql = "SELECT DISTINCT user_id,user_name FROM user_info WHERE user_id not in( SELECT user_id FROM Checker_gate_info ) "
'
'    Set rsTemp = oDb.Execute(szSql, , -1)
'    If rsTemp.RecordCount > 0 Then
'        ReDim aszUserID(1 To rsTemp.RecordCount, 1 To 2)
'        Dim nLoop As Integer
'        rsTemp.MoveFirst
'        For nLoop = 1 To rsTemp.RecordCount
'            aszUserID(nLoop, 1) = rsTemp.Fields(0)
'            rsTemp.MoveNext
'        Next nLoop
'    End If
'    GetNewChecker = aszUserID
'
'End Function
'得到检票人员
Public Function GetAllUser(Optional pszCheckGate As String) As String()
    AssertActiveUserValid m_oActiveUser, cnErroeStartNo
    
    Dim oDb As New RTConnection
    Dim rsTemp  As New Recordset
    Dim szSql As String
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    
    Dim aszUsersInfo() As String
    
    szSql = "SELECT a.user_id,user_name,b.check_gate_id FROM user_info a ,Checker_gate_info b" & _
            " WHERE a.user_id=b.user_id"
    If pszCheckGate <> "" Then szSql = szSql & " AND b.check_gate_id=" & TransFieldValueToString(pszCheckGate)
            
    Set rsTemp = oDb.Execute(szSql, , -1)
    If rsTemp.RecordCount > 0 Then
        ReDim aszUsersInfo(1 To rsTemp.RecordCount, 1 To 3)
        Dim nLoop As Integer
        rsTemp.MoveFirst
        For nLoop = 1 To rsTemp.RecordCount
            aszUsersInfo(nLoop, 1) = FormatDbValue(rsTemp.Fields(0))
            aszUsersInfo(nLoop, 2) = FormatDbValue(rsTemp.Fields(1))
            aszUsersInfo(nLoop, 3) = FormatDbValue(rsTemp.Fields(2))
            rsTemp.MoveNext
        Next nLoop
    End If
    GetAllUser = aszUsersInfo
End Function

Public Function GetRunVehicle(szBusID As String) As M_TRunVehicle()
'得到某一车次的运行车辆
    Dim aTmpVehicle() As M_TRunVehicle
    'Dim aszVehicle() As String
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTmp As Recordset
    Dim lErr As Long
'    Dim szProjectID As String
    Dim i As Integer
    
    'AssertActiveUserValid m_oActiveUser, cnErroeStartNo
    If m_nObjectStatus = ST_AddObj Then
        ShowError ERR_AddObj
    Else
        oDb.ConnectionString = GetConnectionStr(cszSystemMan)
'        szProjectID = Trim(SelfGetExecuteBusProject(Date, lErr).szProjectID)
        szSql = "SELECT B.vehicle_id vehicleId,I.license_tag_no license,O.owner_name owner,V.vehicle_type_short_name vehicleType" _
            & " ,T.transport_company_name company,I.status status FROM " _
            & "owner_info O,Company_info T,vehicle_type_code V,Vehicle_info I,bus_vehicle_code B " _
            & " WHERE B.bus_id='" & szBusID & "' AND B.vehicle_id=I.vehicle_id AND I.owner_id=O.owner_id " _
            & " AND I.transport_company_id=T.transport_company_id AND I.vehicle_type_code=V.vehicle_type_code "
        Set rsTmp = oDb.Execute(szSql)
        If rsTmp.RecordCount <> 0 Then
            ReDim aTmpVehicle(1 To rsTmp.RecordCount)
            For i = 1 To rsTmp.RecordCount
                aTmpVehicle(i).VehicleID = FormatDbValue(rsTmp!VehicleID)
                aTmpVehicle(i).Vehicle = FormatDbValue(rsTmp!license)
                aTmpVehicle(i).Owner = FormatDbValue(rsTmp!Owner)
                aTmpVehicle(i).VehicleType = FormatDbValue(rsTmp!VehicleType)
                aTmpVehicle(i).Company = FormatDbValue(rsTmp!Company)
                aTmpVehicle(i).Status = FormatDbValue(rsTmp!Status)
                rsTmp.MoveNext
            Next i
        End If
        GetRunVehicle = aTmpVehicle
    End If
End Function

Public Property Get CheckGateName() As String
    CheckGateName = m_sCheckGateName
End Property

Public Function GetTicketLst(Optional szBgnTicketId As String = "", Optional szEndTicketId As String = "", Optional nQueryMode = 0, Optional dtBeginDate As Date = cszEmptyDateStr, Optional dtEndDate As Date = cszEmptyDateStr) As String()
'************************************************************************
'参数注释
'-------------
'nQueryMode 查询方式   0----车票模糊匹配方式   1----车票范围查询方式
'szBgnTicketId 车票ID匹配串（方式 0），起始车票ID（方式 1）
'szEndTicketId 终止车票ID（方式 1）
'dtBeginDate  起始日期
'dtEndDate 终止日期
'************************************************************************
    AssertActiveUserValid m_oActiveUser, cnErroeStartNo
    
    Dim aRtnArray() As String
    Dim oDb As New RTConnection
    Dim rsTemp  As New Recordset
    Dim szSql As String
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    
    Dim szCondition As String
    szCondition = ""
    If nQueryMode = 0 Then
        If szBgnTicketId <> "" Then szCondition = szCondition & " AND ticket_id LIKE '" & szBgnTicketId & "%'"
    Else
        If szBgnTicketId <> "" Then
            szCondition = " AND ticket_id>='" & szBgnTicketId & "'"
        End If
        If szEndTicketId <> "" Then
            szCondition = szCondition & " AND ticket_id<='" & szEndTicketId & "'"
        End If
    End If
    If dtBeginDate <> cszEmptyDateStr Then
        szCondition = szCondition & " AND bus_date>='" & Format(dtBeginDate, "YYYY-MM-DD") & "'"
    End If
    If dtEndDate <> cszEmptyDateStr Then
        szCondition = szCondition & " AND bus_date<='" & Format(dtEndDate, "YYYY-MM-DD") & "'"
    End If
'    szCondition = szCondition & " AND check_gate_id=" & TransFieldValueToString(m_sCheckGateNo)
    
    szSql = "SELECT ticket_id " & _
            "FROM check_ticket_lst  WHERE 1=1" & szCondition & " ORDER BY ticket_id"
    Set rsTemp = oDb.Execute(szSql, , -1)
    If rsTemp.RecordCount > 0 Then
        ReDim aRtnArray(1 To rsTemp.RecordCount)
        Dim i As Long
        For i = 1 To rsTemp.RecordCount
            aRtnArray(i) = FormatDbValue(rsTemp!ticket_id)
            rsTemp.MoveNext
        Next i
    End If
    GetTicketLst = aRtnArray
End Function

Private Sub Class_Initialize()
#If SUPPORT_MSG = 1 Then
    

    Set m_oMsg = New CSendMsg
    m_oMsg.Unit = GetUnitID
    m_oMsg.MsgSource = cszSNChkTK
#End If

End Sub


Public Function GetCheckSheetLstByCode(Optional szBgnSheetNo As String = "", Optional szEndSheetNo As String = "", Optional dtBusDate As Date = cszEmptyDateStr) As String()
     AssertActiveUserValid m_oActiveUser, cnErroeStartNo
    
    Dim oDb As New RTConnection
    Dim rsTemp  As New Recordset
    Dim szSql As String
    
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    
    Dim aszSheetNo() As String
    
    szSql = "SELECT check_sheet_id FROM sheet_lst WHERE 1=1 "
    If dtBusDate <> cszEmptyDateStr Then
        szSql = szSql & " AND bus_date='" & Format(dtBusDate, "YYYY-MM-DD") & "'"
    End If
    If szBgnSheetNo <> "" Then
        szSql = szSql & " AND check_sheet_id>='" & szBgnSheetNo & "' "
    End If
    If szEndSheetNo <> "" Then
        szSql = szSql & " AND check_sheet_id<='" & szEndSheetNo & "' "
    End If
    szSql = szSql & " ORDER BY check_sheet_id"
    
    Set rsTemp = oDb.Execute(szSql, , -1)
    If rsTemp.RecordCount > 0 Then '路单不存在
        ReDim aszSheetNo(1 To rsTemp.RecordCount)
        Dim i As Long
        For i = 1 To rsTemp.RecordCount
            aszSheetNo(i) = rsTemp("check_sheet_id")
            rsTemp.MoveNext
        Next i
    End If
    GetCheckSheetLstByCode = aszSheetNo

End Function
Public Function GetCheckSheetLstByBus(Optional dtBusDate As Date = cszEmptyDateStr, Optional szBusID As String = "", Optional nSerialNo As Integer = 0) As String()
     AssertActiveUserValid m_oActiveUser, cnErroeStartNo
    
    Dim oDb As New RTConnection
    Dim rsTemp  As New Recordset
    Dim szSql As String
    
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    
    Dim aszSheetNo() As String
    
    szSql = "SELECT check_sheet_id FROM sheet_lst WHERE 1=1 "
    If dtBusDate <> cszEmptyDateStr Then
        szSql = szSql & " AND bus_date='" & Format(dtBusDate, "YYYY-MM-DD") & "' "
    End If
    If szBusID <> "" Then
        szSql = szSql & " AND bus_id LIKE '" & szBusID & "%' "
    End If
    If nSerialNo <> 0 Then
        szSql = szSql & " AND bus_serial_no= " & Str(nSerialNo) & " "
    End If
 '   szSql = szSql & " AND check_gate_id=" & TransFieldValueToString(m_sCheckGateNo)
    
    
    Set rsTemp = oDb.Execute(szSql, , -1)
    If rsTemp.RecordCount > 0 Then '路单不存在
        ReDim aszSheetNo(1 To rsTemp.RecordCount)
        Dim i As Long
        For i = 1 To rsTemp.RecordCount
            aszSheetNo(i) = rsTemp("check_sheet_id")
            rsTemp.MoveNext
        Next i
    End If
    GetCheckSheetLstByBus = aszSheetNo

End Function

'得到固定车次的座位的检票状态
'返回TSeatStatus数组
Public Function GetBusSeatCheckInfo(ByVal BusID As String, Optional ByVal BusDate As Date, Optional ByVal CheckGateNo As String) As TSeatStatus()
    Dim dtBusDate As Date
    Dim oParam As New SystemParam
    
    Dim szSql As String, oDb As New RTConnection
    Dim rsSeatCheckInfo As Recordset
    Dim szCheckGateNo As String
    Dim nInfoCount As Integer, i As Integer
    Dim asiSeatCheckStatus() As TSeatStatus
    
    AssertActiveUserValid m_oActiveUser, ERR_CheckTicket
    
    If CheckGateNo <> "" Then szCheckGateNo = " AND check_ticket_lst.check_gate_id = " & TransFieldValueToString(CheckGateNo)
    
    oParam.Init m_oActiveUser
    If VBDateIsEmpty(BusDate) Then
        dtBusDate = oParam.NowDate
    Else
        dtBusDate = BusDate
    End If
    On Error GoTo Error_HANDle
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    szSql = "SELECT seat_no,ticket_id FROM Env_bus_seat_lst,check_ticket_lst WHERE " _
    & " Env_bus_seat_lst.bus_id='" & BusID & "' AND " _
    & " Env_bus_seat_lst.bus_date='" & ToDBDate(dtBusDate) & "' AND " _
    & " Env_bus_seat_lst.ticket_no*=check_ticket_lst.ticket_id " & szCheckGateNo _
    & " ORDER BY seat_no "
    
    Set rsSeatCheckInfo = oDb.Execute(szSql)
    If Not rsSeatCheckInfo.EOF Then
        rsSeatCheckInfo.MoveLast
    End If
    nInfoCount = rsSeatCheckInfo.RecordCount
    If nInfoCount > 0 Then
        ReDim asiSeatCheckStatus(1 To nInfoCount)
        rsSeatCheckInfo.MoveFirst
        For i = 1 To nInfoCount
            asiSeatCheckStatus(i).szSeatNo = FormatDbValue(rsSeatCheckInfo!seat_no)
            asiSeatCheckStatus(i).nSeatStatus = IIf(FormatDbValue(rsSeatCheckInfo!ticket_id) = "", 0, 1)
            asiSeatCheckStatus(i).szTicketId = FormatDbValue(rsSeatCheckInfo!ticket_id)
            rsSeatCheckInfo.MoveNext
        Next
    End If
    GetBusSeatCheckInfo = asiSeatCheckStatus
    Exit Function
Error_HANDle:
    err.Raise err.Number
End Function

'得到固定车次的座位的检票状态
'返回TSeatStatus数组
'Public Function GetBusSeatCheckInfoEx(ByVal BusID As String, Optional ByVal BusDate As Date) As TSeatStatusEx()
'    Dim dtBusDate As Date
'    Dim oParam As New SystemParam
'
'    Dim szSql As String, oDb As New RTConnection
'    Dim rsSeatCheckInfo As Recordset
'    Dim nInfoCount As Integer, i As Integer
'    Dim asiSeatCheckStatus() As TSeatStatusEx
'
'    AssertActiveUserValid m_oActiveUser, ERR_CheckTicket
'
'    oParam.Init m_oActiveUser
'    If VBDateIsEmpty(BusDate) Then
'        dtBusDate = oParam.NowDate
'    Else
'        dtBusDate = BusDate
'    End If
'    On Error GoTo Error_HANDle
'    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
'    szSql = "SELECT seat_no,ticket_id,ticket_type FROM Env_bus_seat_lst,check_ticket_lst WHERE " _
'    & " Env_bus_seat_lst.bus_id='" & BusID & "' AND " _
'    & " Env_bus_seat_lst.bus_date='" & ToDBDate(dtBusDate) & "' AND " _
'    & " Env_bus_seat_lst.ticket_no*=check_ticket_lst.ticket_id " _
'    & " ORDER BY seat_no "
'
'    Set rsSeatCheckInfo = oDb.Execute(szSql)
'    If Not rsSeatCheckInfo.EOF Then
'        rsSeatCheckInfo.MoveLast
'    End If
'    nInfoCount = rsSeatCheckInfo.RecordCount
'    If nInfoCount > 0 Then
'        ReDim asiSeatCheckStatus(1 To nInfoCount)
'        rsSeatCheckInfo.MoveFirst
'        For i = 1 To nInfoCount
'            asiSeatCheckStatus(i).szSeatNo = FormatDbValue(rsSeatCheckInfo!seat_no)
'            asiSeatCheckStatus(i).nSeatStatus = IIf(FormatDbValue(rsSeatCheckInfo!ticket_id) = "", 0, 1)
'            asiSeatCheckStatus(i).szTicketType = FormatDbValue(rsSeatCheckInfo!ticket_type)
'            rsSeatCheckInfo.MoveNext
'        Next
'    End If
'    GetBusSeatCheckInfoEx = asiSeatCheckStatus
'    Exit Function
'Error_HANDle:
'    err.Raise err.Number
'End Function


Public Function GetOneTicketInfo(TicketID As String) As TInterfaceCheckTicket

' Function GetOneTicketInfo(TicketID As String) As STChkTk.TInterfaceCheckTicket
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim tData As TInterfaceCheckTicket
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    szSql = "SELECT s.*,sn.station_name " & _
            "FROM Ticket_sell_lst s,station_info sn " & _
            "WHERE s.ticket_id='" & TicketID & "' AND s.des_station_id=sn.station_id"
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 0 Then ShowError ERR_TicketNoExist
    
    tData.BusDate = rsTemp!bus_date
    tData.BusID = rsTemp!bus_id
    tData.SeatNo = rsTemp!seat_no
    tData.TicketStatus = rsTemp!Status
    tData.TicketType = rsTemp!ticket_type
    tData.StationName = rsTemp!station_name
    tData.Price = rsTemp!ticket_price
    tData.SellStationID = rsTemp!sell_station_id
    GetOneTicketInfo = tData
End Function

Public Function GetOneTicketInfoEx(TicketID As String) As TInterfaceCheckTicketEx

    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim tData As TInterfaceCheckTicketEx
    
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    szSql = "SELECT s.*,sn.station_name " & _
            "FROM Ticket_sell_lst s,station_info sn " & _
            "WHERE s.ticket_id='" & TicketID & "' AND s.des_station_id=sn.station_id"
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 0 Then ShowError ERR_TicketNoExist
    
    tData.BusDate = rsTemp!bus_date
    tData.BusID = rsTemp!bus_id
    tData.SeatNo = rsTemp!seat_no
    tData.TicketStatus = rsTemp!Status
    tData.TicketType = rsTemp!ticket_type
    tData.StationName = rsTemp!station_name
    tData.Price = rsTemp!ticket_price
    tData.SellStationID = rsTemp!sell_station_id
    
    tData.CardType = rsTemp!card_type
    tData.IDCardNo = rsTemp!id_card_no
    tData.PersonName = rsTemp!person_name
    tData.National = ""
    tData.Sex = rsTemp!Sex
    tData.Address = ""
    tData.PersonPicture = rsTemp!person_picture
    
    GetOneTicketInfoEx = tData
End Function

'得到车次座位情况
Public Function GetBusSeatInfo(dtBusDate As Date, szBusID As String, szSeatNo As String) As TSeatInfo
Dim tTemp As TSeatInfo
Dim oDb As New RTConnection
Dim rsTemp As New Recordset
On Error GoTo ErrorHandle
    
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    AssertActiveUserValid m_oActiveUser, cnErroeStartNo
    Set rsTemp = oDb.Execute("Select t.*,c.ticket_type_name,s.station_name,i.sell_station_name,y.seat_type_name " _
    & " from ticket_sell_lst t,station_info s,ticket_type_code c, sell_station_info i,seat_type_code y " _
    & " where t.des_station_id=s.station_id and t.ticket_type=c.ticket_type_id and t.sell_station_id=i.sell_station_id and t.seat_type_id=y.seat_type_id and bus_date='" & dtBusDate & "' and bus_id='" & szBusID & "' and seat_no='" & szSeatNo & "' and (t.status in (" & ST_TicketNormal & "," & ST_TicketSellChange & ")" _
    & " or (t.status=" & (ST_TicketNormal & "|" & ST_TicketChecked) & " and ticket_id  in (select ticket_id from check_ticket_lst where bus_date='" & dtBusDate & "' and bus_id='" & szBusID & "' )))")
    If rsTemp.RecordCount > 0 Then
       tTemp.szDestName = Trim(rsTemp!station_name)
'       tTemp.szDestStationID = Trim(rsTemp!des_station_id)
'       tTemp.szSellStationName = Trim(rsTemp!sell_station_name)
'       tTemp.szSellStationID = Trim(rsTemp!sell_station_id)
       tTemp.szSeatNo = Trim(rsTemp!seat_no)
'       tTemp.szTicketType = Trim(rsTemp!ticket_type)
'       tTemp.szTicketTypeName = Trim(rsTemp!ticket_type_name)
       tTemp.szTicketNo = Trim(rsTemp!ticket_id)
       tTemp.szTicketPrice = (rsTemp!ticket_price)
       tTemp.szSeatType = Trim(rsTemp!seat_type_id)
       tTemp.szSeatTypeName = Trim(rsTemp!seat_type_name)
       
    End If
    Set rsTemp = Nothing
    Set oDb = Nothing
    GetBusSeatInfo = tTemp
    Exit Function
    
ErrorHandle:
    err.Raise err.Number, err.Source, err.Description

End Function



Public Sub CheckTicketBeforInitBus(BusID As String, BusDate As Date, REBus As REBus)
    '获得车次状态和车次检票口
    mp_szBusID = REBus.BusID
    mp_dtBusDate = REBus.RunDate
     mp_szMergeBusID = ""
    If REBus.BusType <> TP_ScrollBus Then
        mp_szMergeBusID = Trim(REBus.BeMergedBus.szBusID)
        If mp_szMergeBusID = "" Then
            mp_bMergeBusStatus = False
        Else
            mp_bMergeBusStatus = True
        End If
    Else
        mp_bMergeBusStatus = False
    End If
End Sub

Public Sub InitSystemParam(oAUser As ActiveUser, Optional First As Boolean = True, Optional bAllowChangeRide As Boolean = False)
    If First Then
        Dim oSystem As New SystemParam
        '获得是否改乘
        oSystem.Init oAUser
        mp_bChangeBus = oSystem.AllowChangeRide
    Else
        mp_bChangeBus = bAllowChangeRide
    End If
End Sub

'校对用户是否能在该检票口检票
Private Function CheckUserIsCheckGate(oDb As RTConnection, UserID As String) As Boolean
    Dim rsTemp As Recordset
    Set rsTemp = oDb.Execute("SELECT * FROM Checker_gate_info WHERE check_gate_id='" & m_sCheckGateNo & "' AND user_id='" & UserID & "'")
    If rsTemp.RecordCount <> 0 Then
        CheckUserIsCheckGate = True
    Else
        CheckUserIsCheckGate = False
    End If
End Function

'得到签发单总金额，总件数
Public Function GetLsInfo(szBusID As String, nSerialNo As Integer, dtBusDate As Date) As Recordset
Dim oDb As New RTConnection
Dim rsTemp As Recordset
Dim szSql As String
'
'       oDb.ConnectionString = GetConnectionStr("")
'       szSql = "SELECT  baggage_number,total_price FROM  check_luggage_info " _
'            & " WHERE bus_id='" & szBusID & "' AND bus_serial_no=" & nSerialNo _
'            & "     AND bus_date='" & dtBusDate & "'"
'       Set rsTemp = oDb.Execute(szSql)
'       Set GetLsInfo = rsTemp
End Function

'根据车次代码得被并车次并班、顶班的座位检票信息
Public Function GetMergeSeatCheckInfo(szBusID As String, dtDate As Date) As String()
Dim oDb As New RTConnection
Dim szSql As String
Dim i As Integer
Dim szaTicket() As String
Dim rsTemp As Recordset
Dim nCount As Integer
Dim nTicketCount As Integer
    oDb.ConnectionString = GetConnectionStr("")
    szSql = "SELECT  tbe.bus_id,tbe.ticket_no,tbe.seat_no,tbe.bus_date FROM  Env_bus_seat_lst tbe ,Ticket_sell_lst tbs  " _
            & "WHERE tbe.bus_id='" & szBusID & "' AND tbe.bus_date='" & ToDBDate(dtDate) & "'" _
            & " AND tbe.status>=4 AND tbe.status<=5 AND tbe.bus_date=tbs.bus_date AND tbe.ticket_no=tbs.ticket_id AND tbs.status <> " & ST_TicketNormal
    Set rsTemp = oDb.Execute(szSql)
    '---------------------
    If rsTemp.RecordCount = 0 Then Exit Function
    ReDim szaTicket(1 To rsTemp.RecordCount, 1 To 5) As String
    For i = 1 To rsTemp.RecordCount
        szaTicket(i, 1) = rsTemp!bus_id
        szaTicket(i, 2) = rsTemp!bus_date
        szaTicket(i, 3) = rsTemp!ticket_no
        szaTicket(i, 4) = rsTemp!seat_no
        szaTicket(i, 5) = nCount
        rsTemp.MoveNext
    Next
   GetMergeSeatCheckInfo = szaTicket
    Set oDb = Nothing
End Function

'根据车次代码得被并车次并班、顶班的座位
Public Function GetMergeSeatInfo(szBusID As String, dtDate As Date) As String()
Dim oDb As New RTConnection
Dim szSql As String
Dim i As Integer
Dim szaTicket() As String
Dim rsTemp As Recordset
Dim nCount As Integer
Dim nTicketCount As Integer
    oDb.ConnectionString = GetConnectionStr("")
    szSql = "SELECT  tbe.bus_id,tbe.ticket_no,tbe.seat_no,tbe.bus_date FROM  Env_bus_seat_lst tbe    " _
            & "WHERE tbe.bus_id='" & szBusID & "' AND tbe.bus_date='" & ToDBDate(dtDate) & "'" _
            & " AND tbe.status>=4 AND tbe.status<=5 "
    Set rsTemp = oDb.Execute(szSql)
    '---------------------
    If rsTemp.RecordCount = 0 Then Exit Function
    ReDim szaTicket(1 To rsTemp.RecordCount, 1 To 5) As String
    For i = 1 To rsTemp.RecordCount
        szaTicket(i, 1) = rsTemp!bus_id
        szaTicket(i, 2) = rsTemp!bus_date
        szaTicket(i, 3) = rsTemp!ticket_no
        szaTicket(i, 4) = rsTemp!seat_no
        szaTicket(i, 5) = nCount
        rsTemp.MoveNext
    Next
   GetMergeSeatInfo = szaTicket
    Set oDb = Nothing
End Function
'得到并班车次
Public Function GetMergeBus(szBusID As String, dtDate As Date) As String
Dim oDb As New RTConnection
Dim szSql As String
Dim i As Integer
Dim rsTemp As Recordset
Dim szMergeBus As String

    oDb.ConnectionString = GetConnectionStr("")
    szSql = "SELECT  tbe.merge_bus_id FROM  Mergebus_split_lst tbe    " _
            & "WHERE tbe.bus_id='" & szBusID & "' AND tbe.bus_date='" & ToDBDate(dtDate) & "'"
    Set rsTemp = oDb.Execute(szSql)
    '---------------------
    If rsTemp.RecordCount = 0 Then Exit Function
    For i = 1 To rsTemp.RecordCount
        If szMergeBus = "" Then
           szMergeBus = rsTemp!merge_bus_id
        Else
           szMergeBus = szMergeBus & "," & rsTemp!merge_bus_id
        End If
        rsTemp.MoveNext
    Next
   GetMergeBus = szMergeBus
   Set oDb = Nothing
End Function
'判断是否是并班车次
Private Function IsMergeBus(szBusID As String, szChar As String) As Boolean
Dim i As Integer, nCount As Integer
Dim szTemp() As String
Dim bStatus As Boolean
szTemp = StringToTeam(szChar)
nCount = ArrayLength(szTemp)
For i = 1 To nCount
  If szTemp(i) = szBusID Then
     IsMergeBus = True
    Exit For
  Else
    IsMergeBus = False
  End If
Next

End Function
Private Sub RecycleSeat(pszTicketID As String, nTicketStatus As Integer, oDb As RTConnection, Optional pszBusID As String)
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim szBusID As String
    Dim dtDate As Date
    Dim szDesStationID As String
    Dim szSeatNo As String
    Dim szSeatTypeID As String
    Dim szSellStationID As String
    
    szSql = "SELECT * FROM Ticket_sell_lst WHERE ticket_id='" & pszTicketID & "'"
    Set rsTemp = oDb.Execute(szSql)
    szBusID = FormatDbValue(rsTemp!bus_id)
    dtDate = FormatDbValue(rsTemp!bus_date)
    szDesStationID = FormatDbValue(rsTemp!des_station_id)
    szSeatNo = FormatDbValue(rsTemp!seat_no)
    szSellStationID = FormatDbValue(rsTemp!sell_station_id)
    
    If szSeatNo = "ST" Then '是站票
        '将该环境车次的可售站票数加一
        szSql = "UPDATE Env_bus_info SET " _
        & " sale_stAND_seat_quantity=sale_stAND_seat_quantity+1 WHERE " _
        & " bus_id='" & szBusID & "' AND " _
        & " bus_date='" & ToDBDate(dtDate) & "'"
        oDb.Execute szSql
    Else '不是站票
        szSeatTypeID = FormatDbValue(rsTemp!seat_type_id)
        
        szSeatTypeID = FormatDbValue(rsTemp!seat_type_id)
        
        '将该环境车次的可售座位数减一和可售座位类型余留减一
        szSql = "UPDATE Env_bus_info SET sale_seat_quantity=sale_seat_quantity+1 , "
        Select Case szSeatTypeID
            Case cszSeatType
                szSql = szSql & " seat_remain=seat_remain + 1  "
            Case cszBedType
                szSql = szSql & " bed_remain=bed_remain + 1"
            Case cszAdditionalType
                szSql = szSql & " additional_remain=additional_remain + 1 "
            Case cszOtherType1
                szSql = szSql & " other_remain_1 =other_remain_1 + 1 "
            Case cszOtherType2
                szSql = szSql & " other_remain_2 =other_remain_2 + 1 "
        End Select
        szSql = szSql & " WHERE bus_id='" & szBusID & "' AND bus_date='" & ToDBDate(dtDate) & "'"
        oDb.Execute szSql
                
        '改变相应座位号的状态
        szSql = "UPDATE Env_bus_seat_lst SET " _
        & " status=" & nTicketStatus & "," _
        & " ticket_no='' WHERE " _
        & " bus_id='" & szBusID & "' AND " _
        & " bus_date='" & ToDBDate(dtDate) & "' AND " _
        & " seat_no='" & szSeatNo & "'"
        
        oDb.Execute szSql
    End If
    
    '改变该车次相应到站的可售票数
    szSql = "UPDATE Env_bus_station_lst SET " _
    & " sale_ticket_quantity=sale_ticket_quantity+1 WHERE " _
    & " bus_id='" & szBusID & "' AND " _
    & " bus_date='" & ToDBDate(dtDate) & "' AND " _
    & " station_id='" & szDesStationID & "' AND " _
    & " max_sale_quantity>0 "
    
    oDb.Execute szSql
     
    '配载，售票点已售张数各自都减掉
    szSql = "UPDATE env_bus_sell_station_lst SET " _
    & " sell_quantity=sell_quantity-1 WHERE " _
    & " bus_id='" & szBusID & "' AND " _
    & " bus_date='" & ToDBDate(dtDate) & "' AND " _
    & " sell_station_id='" & szSellStationID & "'"
    
    oDb.Execute szSql
    
    szSql = "UPDATE env_bus_allot_lst SET " _
    & " sell_quantity=sell_quantity-1 WHERE " _
    & " bus_id='" & szBusID & "' AND " _
    & " bus_date='" & ToDBDate(dtDate) & "' AND " _
    & " sell_station_id='" & szSellStationID & "'"
    
    oDb.Execute szSql
End Sub

'返回当前改乘票数
Public Function GetChangeTicket(szBusID As String, dtBusDate As Date) As Integer
Dim oDb As New RTConnection
Dim szSql As String
Dim rsTemp As Recordset
    oDb.ConnectionString = GetConnectionStr("")
    szSql = "SELECT count(*) as n FROM   check_ticket_lst     " _
            & "WHERE bus_id='" & szBusID & "' AND bus_date='" & ToDBDate(dtBusDate) & "' AND status=2"
    Set rsTemp = oDb.Execute(szSql)
    GetChangeTicket = FormatDbValue(rsTemp!n)
   Set oDb = Nothing

End Function

'检票时,更改车次车辆
Public Sub ChangeVehicle(pszVehicle As String, pnBusSerialNo As Integer, pszBusID As String, pdyBusDate As Date)
    '更改车辆信息
    Dim oDb As New RTConnection
    Dim szSql As String
    
    
    Dim szVehicleTypeCode As String
    Dim szOwnerID As String
    Dim szTransportCompanyID As String
    Dim szVehicleTypeName As String
    Dim szSplitCompanyID As String
    
    Dim rsVehicle As Recordset

    On Error GoTo ErrorHandle
    oDb.ConnectionString = GetConnectionStr
    szSql = "SELECT v.*,t.vehicle_type_short_name" & _
            " FROM Vehicle_info v,vehicle_type_code t" & _
            " WHERE v.vehicle_id='" & pszVehicle & "' AND v.vehicle_type_code=t.vehicle_type_code"
    Set rsVehicle = oDb.Execute(szSql)
    If rsVehicle.RecordCount = 0 Then ShowError ERR_ChkTkVehicleNotExist
    
    
    szVehicleTypeCode = FormatDbValue(rsVehicle!vehicle_type_code)
    szOwnerID = FormatDbValue(rsVehicle!owner_id)
    szTransportCompanyID = FormatDbValue(rsVehicle!transport_company_id)
    szVehicleTypeName = FormatDbValue(rsVehicle!vehicle_type_short_name)
    szSplitCompanyID = FormatDbValue(rsVehicle!split_company_id)
    
    '开始修改
    oDb.BeginTrans
        'Bus_check_lst
        '修改检票车次表
        szSql = " UPDATE bus_check_lst " & _
            " SET vehicle_id='" & pszVehicle & "'" & _
            "    ,transport_company_id='" & szTransportCompanyID & "'" & _
            "    ,owner_id='" & szOwnerID & "'" & _
            "    ,split_company_id='" & szSplitCompanyID & "'" & _
            "    ,vehicle_type_code='" & szVehicleTypeCode & "'" & _
            " WHERE bus_id='" & pszBusID & "' AND " & DBDateCompare(pdyBusDate, "bus_date") _
            & " AND bus_serial_no = '" & pnBusSerialNo & "'" & " AND check_gate_id = '" & m_sCheckGateNo & "'"
             'SET vehicle_id = , transport_company_id = , vehicle_type_code = , owner_id = , split_company_id =  "
        oDb.Execute szSql
        
        
        '环境车次表
        szSql = "UPDATE Env_bus_info" & _
             " SET vehicle_id='" & pszVehicle & "'" & _
             "    ,transport_company_id='" & szTransportCompanyID & "'" & _
             "    ,owner_id='" & szOwnerID & "'" & _
             "    ,split_company_id='" & szSplitCompanyID & "'" & _
             "    ,vehicle_type_code='" & szVehicleTypeCode & "'" & _
             "    ,vehicle_type_name='" & szVehicleTypeName & "'" & _
             " WHERE bus_id='" & pszBusID & "' AND " & DBDateCompare(pdyBusDate, "bus_date")
    oDb.Execute szSql '固定车次回写环境车次表状
    
    oDb.CommitTrans
    WriteOperateLog m_oActiveUser, RIGHT_ChkTkCheckBus, "在检票时，把车次[" & pszBusID & "]日期为[" & pdyBusDate & "]序号为[" & pnBusSerialNo & "]，车辆改为" & pszVehicle
    

    Exit Sub
ErrorHandle:
    oDb.RollbackTrans
    err.Raise err.Number, err.Source, err.Description
    
End Sub







'得到最后的路单号
Public Function GetLastCheckSheetID(Checker As String) As String
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim oParam As New SystemParam
    oDb.ConnectionString = GetConnectionStr()
    
    szSql = "SELECT MAX(check_sheet_id) AS check_sheet_id FROM sheet_lst WHERE " _
    & " sheet_make_date=(" _
    & " SELECT MAX(sheet_make_date) FROM sheet_lst WHERE " _
    & " checker='" & Checker & "') "
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 1 Then
        GetLastCheckSheetID = FormatDbValue(rsTemp!check_sheet_id)
    Else
        oParam.Init m_oActiveUser
        GetLastCheckSheetID = String(oParam.CheckSheetLen, "0")
        Set oParam = Nothing
    End If
End Function

Public Sub RightChangeBus()
    AssertHaveRight m_oActiveUser, RIGHT_ChkTkChangeCheckBus
End Sub


'设置LED显示状态
Public Sub SetLED(BusID As String, dtBusDate As Date, pnLED As Long)
    AssertActiveUserValid m_oActiveUser, cnErroeStartNo
    'AssertHaveRight m_oActiveUser, RIGHT_ChkTkCheckTicket
    Dim oDb  As New RTConnection
    Dim szSql As String
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    
    BusID = Trim(BusID)
    If Len(m_sCheckGateNo) = 0 Then '未指定检票口
        ShowError ERR_ChkTkCheckGateIDNotAppointed
        Exit Sub
    End If
    szSql = "UPDATE env_bus_allot_lst SET led=" & pnLED & _
            " WHERE bus_id='" & BusID & "' AND " & DBDateCompare(dtBusDate, "bus_date") & _
            " AND check_gate_id='" & m_sCheckGateNo & "'"
    oDb.Execute szSql
    
            
End Sub

'--判断是否可以改车，开检前和开检时
'--06-01-24 qlh
Public Function SelectChangeBusBeforeCheetIsValid() As Boolean
    AssertActiveUserValid m_oActiveUser, cnErroeStartNo
'    #If IN_DEBUG = 0 Then
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim oDb As New RTConnection
    Dim szRight As String
    
    szRight = GetRightID(RIGHT_ChkTkChangeBusBeforeCheet)
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    '查询用户方法表中用户是否有该权限
    szSql = "SELECT function_id FROM user_function_lst WHERE " _
    & " user_id='" & m_oActiveUser.UserID & "' AND " _
    & " function_id='" & szRight & "'"
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 0 Then
        '如果用户没有权限则查询用户组是否有权限
        szSql = "SELECT function_id FROM usergroup_function_lst tbu,Group_user_info tbg WHERE tbu.usergroup_id=" _
        & " tbg.usergroup_id AND user_id='" & m_oActiveUser.UserID & "' AND function_id='" & szRight & "'"
        Set rsTemp = oDb.Execute(szSql)
        If rsTemp.RecordCount = 0 Then
            SelectChangeBusBeforeCheetIsValid = False
        Else
            SelectChangeBusBeforeCheetIsValid = True
        End If
    Else
        SelectChangeBusBeforeCheetIsValid = True
    End If
'#End If
End Function

'--判断是否可以改车，路单打出后
'--06-01-24 qlh
Public Function SelectChangeBusAfterCheetIsValid() As Boolean
    AssertActiveUserValid m_oActiveUser, cnErroeStartNo
'    #If IN_DEBUG = 0 Then
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim oDb As New RTConnection
    Dim szRight As String
    
    szRight = GetRightID(RIGHT_ChkTkChangeBusAfterCheet)
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    '查询用户方法表中用户是否有该权限
    szSql = "SELECT function_id FROM user_function_lst WHERE " _
    & " user_id='" & m_oActiveUser.UserID & "' AND " _
    & " function_id='" & szRight & "'"
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 0 Then
        '如果用户没有权限则查询用户组是否有权限
        szSql = "SELECT function_id FROM usergroup_function_lst tbu,Group_user_info tbg WHERE tbu.usergroup_id=" _
        & " tbg.usergroup_id AND user_id='" & m_oActiveUser.UserID & "' AND function_id='" & szRight & "'"
        Set rsTemp = oDb.Execute(szSql)
        If rsTemp.RecordCount = 0 Then
            SelectChangeBusAfterCheetIsValid = False
        Else
            SelectChangeBusAfterCheetIsValid = True
        End If
    Else
        SelectChangeBusAfterCheetIsValid = True
    End If
'#End If
End Function

'判断是否有补打路单的权限
Public Function SelectExPrintSheetIsValid() As Boolean
    AssertActiveUserValid m_oActiveUser, cnErroeStartNo
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim oDb As New RTConnection
    Dim szRight As String
    
    szRight = GetRightID(RIGHT_ChkTkExPrintCheckSheet)
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    '查询用户方法表中用户是否有该权限
    szSql = "SELECT function_id FROM user_function_lst WHERE " _
    & " user_id='" & m_oActiveUser.UserID & "' AND " _
    & " function_id='" & szRight & "'"
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 0 Then
        '如果用户没有权限则查询用户组是否有权限
        szSql = "SELECT function_id FROM usergroup_function_lst tbu,Group_user_info tbg WHERE tbu.usergroup_id=" _
        & " tbg.usergroup_id AND user_id='" & m_oActiveUser.UserID & "' AND function_id='" & szRight & "'"
        Set rsTemp = oDb.Execute(szSql)
        If rsTemp.RecordCount = 0 Then
            SelectExPrintSheetIsValid = False
        Else
            SelectExPrintSheetIsValid = True
        End If
    Else
        SelectExPrintSheetIsValid = True
    End If
End Function

'得到下一个未检车次
'************************lyq 2006-2-13 新增接口*********************
'用于为检票屏取车次用
Public Function GetNextLEDCheckBus() As REBus
    Dim oReg As CFreeReg
    Dim szCheckGateType As String

    Dim oDb As New RTConnection
    Dim rsTemp As New Recordset
    Dim szSql As String
    Dim dtMinDate As Date
    Dim oREBus As New REBus
    Dim dCheckTime As Double
    Dim rsRecord As Recordset       '返回全部记录集
    Dim szChar As String
    Dim rsFind As Recordset
    Dim nPostion As Integer
    Dim szMinBusID As String
    
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    AssertActiveUserValid m_oActiveUser, cnErroeStartNo
    If Len(m_sCheckGateNo) = 0 Then '未指定检票口
        ShowError ERR_ChkTkCheckGateIDNotAppointed
        Exit Function
    End If
    
    '当前发车时间大于当前时间和最晚补检时间之和的第一班车次
    '    dCheckTime = GetParameterValue(cszExtraStartCheckTime)    '取最晚补检时间
    '    If dCheckTime < 0 Then dCheckTime = 0
    
    If Len(szCheckGateType) = 0 Then
'        szSql = "SELECT bus_id,bus_date FROM Work_env_bus_info " & _
'        "WHERE " & DBDateCompare(Date, "bus_date") & " AND check_gate_id='" & Trim(m_sCheckGateNo) & "' AND (status=" & Str(EREBusStatus.ST_BusNormal) & " or status=32 OR status = 16) AND bus_start_time=(SELECT MIN(bus_start_time) FROM Work_env_bus_info " & _
'        "WHERE " & DBTimeCompareEX(Now, "bus_start_time", "<") & " AND check_gate_id='" & Trim(m_sCheckGateNo) & "' AND " & DBDateCompare(Date, "bus_date") & " AND (status=" & Str(EREBusStatus.ST_BusNormal) & " or status=32 OR status = 16)) order by bus_start_time,bus_id"
        szSql = "SELECT a.bus_id,a.bus_date,b.check_gate_id FROM Work_env_bus_info a,Work_env_bus_allot_lst b " & _
        "WHERE b.can_sale_quantity <>0 AND a.bus_type=" & TP_RegularBus & " AND " & DBDateCompare(Date, "a.bus_date") & " AND b.check_gate_id='" & Trim(m_sCheckGateNo) & "' AND (b.status in (" & EREBusStatus.ST_BusNormal & ",32,16," & EREBusStatus.ST_BusChecking & ")) AND b.bus_start_time=(SELECT MIN(bus_start_time) FROM Work_env_bus_allot_lst " & _
        "WHERE  can_sale_quantity <>0 AND  " & DBTimeCompareEX(Now, "bus_start_time", "<") & " AND check_gate_id='" & Trim(m_sCheckGateNo) & "' AND " & DBDateCompare(Date, "bus_date") & " AND (status  in (" & EREBusStatus.ST_BusNormal & ",32,16," & EREBusStatus.ST_BusChecking & ")))" & _
        " AND a.bus_date=b.bus_date AND a.bus_id=b.bus_id "
        ' &         " order by bus_start_time,bus_id"
    Else
'        szSql = "SELECT * FROM Work_env_bus_info WHERE  check_gate_id='" & Trim(m_sCheckGateNo) & "' AND " & DBDateCompare(Date, "bus_date") & " order by bus_start_time,bus_id" '该检票口的车次
'        Set rsRecord = oDb.Execute(szSql)
'        Set rsFind = rsRecord
'        '得到下一班车次的发车时间
'        szSql = "SELECT MIN(bus_start_time) as bus_start_time FROM Work_env_bus_info WHERE " & DBTimeCompareEX(Now, "bus_start_time", "<") & " AND check_gate_id='" & Trim(m_sCheckGateNo) & "' AND " & DBDateCompare(Date, "bus_date") & " AND ((status=" & Str(EREBusStatus.ST_BusNormal) & ") or (status=32) OR (status = 16)) order by bus_start_time"
'        Set rsTemp = oDb.Execute(szSql)
'here1:  If rsTemp.RecordCount = 0 Then Exit Function
'        dtMinDate = FormatDbValue(rsTemp!bus_start_time)
'        '查找下一班车次在记录集中的位置
'        rsFind.Filter = "bus_start_time='" & ToDBDateTime(dtMinDate) & "' AND status=0 or status=16 "
'        If rsFind.RecordCount = 0 Then Exit Function
'        nPostion = rsFind.Bookmark
'        If rsFind.RecordCount > 1 Then '同一时间有多班车次
'            szSql = "SELECT  bus_start_time, bus_id FROM Work_env_bus_info WHERE " & DBTimeCompareEX(dtMinDate, "bus_start_time", "=") & " AND check_gate_id='" & Trim(m_sCheckGateNo) & "' AND " & DBDateCompare(Date, "bus_date") & " AND (status=" & Str(EREBusStatus.ST_BusNormal) & " or status=32 OR status = 16)  order by bus_start_time,bus_id"
'            Set rsTemp = oDb.Execute(szSql)
'            If rsTemp.RecordCount = 0 Then Exit Function '无该时间发车车次
'
'Here2:      If rsTemp.EOF Then '当前时间当前检票口的单（或双）号无车次要检票（都已检、在检或停班），必须转换到下一检票时间
'                szSql = "SELECT MIN(bus_start_time) as bus_start_time FROM Work_env_bus_info WHERE " & DBTimeCompareEX(dtMinDate, "bus_start_time", "<") & " AND check_gate_id='" & Trim(m_sCheckGateNo) & "' AND " & DBDateCompare(Date, "bus_date") & " AND (status=" & Str(EREBusStatus.ST_BusNormal) & " or status=32 OR status = 16) order by bus_start_time"
'                Set rsTemp = oDb.Execute(szSql)
'                GoTo here1
'            Else  '查找当前时间当前检票口的未检车次
'                Set rsFind = rsRecord
'                szMinBusID = FormatDbValue(rsTemp!bus_id)
'                rsFind.Filter = "bus_start_time='" & ToDBDateTime(dtMinDate) & "' AND bus_id='" & szMinBusID & "'"
'                If rsFind.RecordCount = 0 Then Exit Function
'                nPostion = rsFind.Bookmark
'                '单号检票台＼双号位置车次 或　双号检票台＼单号位置车次 继续查找
'                '否则对该车次检票，　单号位检票台＼单号位车次在单号台检，双号位检票台＼双号位车次在双号台检
'
'                If (szCheckGateType = 1 AND nPostion Mod 2 = 0) Or (szCheckGateType = 2 AND nPostion Mod 2 = 1) Then
'                    If Not rsTemp.EOF Then rsTemp.MoveNext
'                    GoTo Here2
'                End If
'            End If
'        Else
'            szMinBusID = FormatDbValue(rsFind!bus_id)
'            '单号位检票台＼双号位车次 或 双号位检票台＼单号位车次 则更换下一不同时间，
'            '否则对该车次检票，　单号位检票台＼单号位车次在单号台检，双号位检票台＼双号位车次在双号台检
'            If (szCheckGateType = 1 AND nPostion Mod 2 = 0) Or (szCheckGateType = 2 AND nPostion Mod 2 = 1) Then
'                szSql = "SELECT MIN(bus_start_time) as bus_start_time FROM Work_env_bus_info WHERE " & DBTimeCompareEX(dtMinDate, "bus_start_time", "<") & " AND check_gate_id='" & Trim(m_sCheckGateNo) & "' AND " & DBDateCompare(Date, "bus_date") & " AND (status=" & Str(EREBusStatus.ST_BusNormal) & " or status=32 OR status = 16) order by bus_start_time"
'                Set rsTemp = oDb.Execute(szSql)
'                GoTo here1
'            End If
'        End If
'        '能找到则对以下车次检票，不能找到则不会执行以下语句
'
'        szSql = "SELECT bus_id,bus_date FROM Work_env_bus_info " & _
'        "WHERE " & DBDateCompare(Date, "bus_date") & " AND bus_id='" & Trim(szMinBusID) & "'"
'
    End If
    
    If Trim(szSql) = "" Then Exit Function
    Set rsTemp = oDb.Execute(szSql, , -1)
    If rsTemp.RecordCount = 0 Then '没有车次被选择
        Exit Function
    End If
    
    oREBus.Init m_oActiveUser
   ' oREBus.Init m_szUser
    oREBus.Identify rsTemp.Fields("bus_id"), rsTemp.Fields("bus_date"), rsTemp.Fields("check_gate_id")
    

    Set GetNextLEDCheckBus = oREBus

End Function

Public Function GetSelledSpecialTicketCount(pszBusID As String, pszBusDate As Date, pszSpecialTicketTypePosition As String) As Integer
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset
    Dim nSelledSpecialTicketTypeCount As Integer
    oDb.ConnectionString = GetConnectionStr("")
On Error GoTo ErrorHandle
    szSql = "SELECT count(*) as selled_special_ticket_count FROM agent_ticket_sale_lst " _
      & " WHERE bus_id='" & pszBusID & "' AND  bus_date='" & ToDBDate(pszBusDate) & "'" _
      & " AND  status='1' AND ticket_type='" & pszSpecialTicketTypePosition & "' "
    Set rsTemp = oDb.Execute(szSql)

        nSelledSpecialTicketTypeCount = FormatDbValue(rsTemp!selled_special_ticket_count)

    GetSelledSpecialTicketCount = nSelledSpecialTicketTypeCount
    Exit Function
ErrorHandle:
    err.Raise err.Number
End Function


Public Function GetAllotStationTicketsInfo(pszBusID As String, pdtBusDate As Date) As Recordset
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset

    oDb.ConnectionString = GetConnectionStr("")
On Error GoTo ErrorHandle
    
    szSql = " SELECT weal.sell_station_id, MAX(ssi.sell_station_name) sell_station_name, COUNT(tsl.ticket_id) ticket_num " _
            & " FROM work_env_bus_allot_lst weal, ticket_sell_lst tsl, sell_station_info ssi " _
            & " WHERE weal.bus_id *= tsl.bus_id AND weal.bus_date *= tsl.bus_date " _
            & " AND weal.sell_station_id *= tsl.sell_station_id AND ssi.sell_station_id = weal.sell_station_id AND tsl.status IN(1,2,33,34) " _
            & " AND weal.bus_id = " & TransFieldValueToString(pszBusID) & " AND weal.bus_date = " & TransFieldValueToString(pdtBusDate) _
            & " GROUP BY weal.sell_station_id "
    
    Set rsTemp = oDb.Execute(szSql)

    Set GetAllotStationTicketsInfo = rsTemp
    Exit Function
ErrorHandle:
    err.Raise err.Number
End Function


Public Function GetSellStationTicketsInfo(pszBusID As String, pdtBusDate As Date) As Recordset
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset

    oDb.ConnectionString = GetConnectionStr("")
On Error GoTo ErrorHandle
    
    szSql = " SELECT wbsl.sell_station_id, wbsl.bus_id, wbsl.bus_date,wbsl.sell_station_name,ISNULL(tsl.ticket_num, 0) ticket_num FROM " _
            & " (SELECT si.sell_station_id,bus_id, tsl.bus_date, MAX(ssi.sell_station_name) sell_station_name, COUNT(tsl.ticket_id) ticket_num " _
            & " FROM  ticket_sell_lst tsl, sell_station_info ssi, user_info si " _
            & " WHERE ssi.sell_station_id = si.sell_station_id AND tsl.status IN(1,2,33,34) " _
            & " AND tsl.user_id = si.user_id AND tsl.bus_id = " & TransFieldValueToString(pszBusID) & " AND tsl.bus_date = " & TransFieldValueToString(pdtBusDate) _
            & " GROUP BY si.sell_station_id,tsl.bus_id, tsl.bus_date) tsl, " _
            & " (SELECT ssi.sell_station_id,sell_station_name,bus_id,bus_date FROM work_env_bus_sell_station_lst wbsl,sell_station_info ssi " _
            & " WHERE wbsl.sell_station_id = ssi.sell_station_id AND bus_id = " & TransFieldValueToString(pszBusID) & " AND bus_date = " & TransFieldValueToString(pdtBusDate) & ") wbsl " _
            & " WHERE wbsl.sell_station_id *= tsl.sell_station_id AND wbsl.bus_id *= tsl.bus_id AND wbsl.bus_date *= tsl.bus_date "
            
    Set rsTemp = oDb.Execute(szSql)

    Set GetSellStationTicketsInfo = rsTemp
    Exit Function
ErrorHandle:
    err.Raise err.Number
End Function

Public Function GetAllotTicketsCheckInfo(pszBusID As String, pdtBusDate As Date) As Recordset
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim rsTemp As Recordset

    oDb.ConnectionString = GetConnectionStr("")
On Error GoTo ErrorHandle
    
    szSql = " SELECT weal.sell_station_id, weal.bus_id, weal.bus_date,si.sell_station_name, ISNULL(cti.ticket_num,0) ticket_num FROM " _
            & " (SELECT ci.sell_station_id,ctl.bus_id, ctl.bus_date, MAX(si.sell_station_name) sell_station_name,COUNT(ctl.ticket_id) ticket_num " _
            & " FROM check_ticket_lst ctl, checkgate_info ci, sell_station_info si " _
            & " Where ctl.check_gate_id = ci.check_gate_id And ci.sell_station_id = si.sell_station_id " _
            & " AND ctl.bus_id = " & TransFieldValueToString(pszBusID) & " AND ctl.bus_date = " & TransFieldValueToString(pdtBusDate) _
            & " GROUP BY ci.sell_station_id, ctl.bus_id, ctl.bus_date) cti, " _
            & " (SELECT bus_id, bus_date, sell_station_id FROM work_env_bus_allot_lst " _
            & " WHERE bus_id = " & TransFieldValueToString(pszBusID) _
            & " AND bus_date = " & TransFieldValueToString(pdtBusDate) & " ) weal, sell_station_info si " _
            & " WHERE weal.bus_id *= cti.bus_id AND weal.bus_date *= cti.bus_date " _
            & " AND weal.sell_station_id  = si.sell_station_id AND weal.sell_station_id *= cti.sell_station_id"
    
    Set rsTemp = oDb.Execute(szSql)

    Set GetAllotTicketsCheckInfo = rsTemp
    Exit Function
ErrorHandle:
    err.Raise err.Number
End Function

Public Function GetUnCheckBusInfo(StartDate As Date, EndDate As Date, Optional VehicleID As String, Optional BusID As String) As String()
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim aszTemp() As String
    Dim nCount As Integer
    Dim i As Integer
    
    On Error GoTo ErrorHandle
    '3,4,5对应 ST_SeatSold, ST_SeatReplace, ST_SeatSlitp
    oDb.ConnectionString = GetConnectionStr()
    szSql = "select e.*,a.bus_start_time as bus_start_time1,c.transport_company_short_name ," _
        & " c1.transport_company_short_name as split_company_short_name,license_tag_no,route_name,bus_type_name " _
        & " from work_env_bus_info e,company_info c,vehicle_info v ,company_info c1 ,route_info r,bus_type_code b ,env_bus_allot_lst a, " _
        & " ( select bus_date,bus_id,count(ticket_no) as ticket_number from env_bus_seat_lst " _
        & " where status in(3,4,5) and ticket_no<>'' and ticket_no not in (select ticket_id from check_ticket_lst ) group by bus_date,bus_id) s " _
        & " Where e.split_company_id = c1.transport_company_id And c.transport_company_id = e.transport_company_id " _
        & " and r.route_id=e.route_id and b.bus_type=e.bus_type and s.bus_id=e.bus_id and s.bus_date=e.bus_date " _
        & " and v.vehicle_id=e.vehicle_id and e.bus_date>='" & StartDate & "' and e.bus_date<='" & EndDate & "'" _
        & " and a.bus_id=e.bus_id and a.bus_date=e.bus_date AND a.bus_start_time<'" & Now & "'" _
        & " and a.sell_station_id='" & m_oActiveUser.SellStationID & "' and a.check_gate_id='" & m_sCheckGateNo & "'"
        
        If VehicleID <> "" Then szSql = szSql & " and e.vehicle_id like '%" & VehicleID & "%'"
        If BusID <> "" Then szSql = szSql & " and e.bus_id like '%" & BusID & "%'"
        
    Set rsTemp = oDb.Execute(szSql)
    nCount = rsTemp.RecordCount
    If rsTemp.RecordCount > 0 Then
        ReDim aszTemp(1 To rsTemp.RecordCount, 1 To 16)
        For i = 1 To rsTemp.RecordCount
            aszTemp(i, 1) = FormatDbValue(rsTemp!bus_id)
            aszTemp(i, 2) = FormatDbValue(rsTemp!bus_start_time1)
            
            aszTemp(i, 3) = FormatDbValue(rsTemp!route_name)
            aszTemp(i, 4) = FormatDbValue(rsTemp!check_gate_id)
            aszTemp(i, 5) = FormatDbValue(rsTemp!license_tag_no)
            aszTemp(i, 6) = FormatDbValue(rsTemp!vehicle_type_name)
            aszTemp(i, 7) = FormatDbValue(rsTemp!Status)
            aszTemp(i, 8) = FormatDbValue(rsTemp!bus_type)
            aszTemp(i, 9) = FormatDbValue(rsTemp!is_all_refundment)
            aszTemp(i, 10) = FormatDbValue(rsTemp!end_station_name)
            aszTemp(i, 11) = FormatDbValue(rsTemp!total_seat)
            aszTemp(i, 12) = FormatDbValue(rsTemp!transport_company_short_name)
            aszTemp(i, 13) = FormatDbValue(rsTemp!split_company_short_name)
            aszTemp(i, 14) = FormatDbValue(rsTemp!bus_type_name)
            aszTemp(i, 15) = FormatDbValue(rsTemp!sale_seat_quantity)
            
'            aszTemp(i, 16) = FormatDbValue(rsTemp!bus_kind) '车次种类
'            aszTemp(i, 17) = FormatDbValue(rsTemp!add_status) '是否加班车
            aszTemp(i, 16) = FormatDbValue(rsTemp!vehicle_id) '车辆代码
            
            rsTemp.MoveNext
        Next
    End If
    Set oDb = Nothing
    Set rsTemp = Nothing
    GetUnCheckBusInfo = aszTemp
    Exit Function
    
ErrorHandle:
    
    Set oDb = Nothing
    Set rsTemp = Nothing
    err.Raise err.Number, err.Source, err.Description
End Function

