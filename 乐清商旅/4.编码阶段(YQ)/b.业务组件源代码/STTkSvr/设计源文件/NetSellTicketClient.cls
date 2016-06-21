VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NetSellTicketClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
 
Const cszStandTicketSeatNo = "ST"
Const cszAutoSeatNo = ""
Const cszServiceFinish = "对方车站已写入,请通知管理员在对方车站将此票删除！"

'此处定义错误号:
Public Enum EErrNetSellTicketClient
    ERR_SellTicketClientNoActiveUser = ERR_SellTicketClient + ERR_NoActiveUser
    ERR_UserMustBeLocalUser = ERR_SellTicketClient + 21  '使用售票客户的用户必须是本单位用户
End Enum

Public Enum EErrBook

    ERR_BookNoActiveUser = ERR_Book + ERR_NoActiveUser '30011
    ERR_BookSeatSaled = ERR_Book + 21 '指定的座位已经售出  30021
    ERR_BookSeatNotExist = ERR_Book + 22 '指定的座位不存在 30022
    ERR_BookSeatReserved = ERR_Book + 23 '指定的座位已经被预留30023
    ERR_BOOKSeatBooked = ERR_Book + 24 '指定的座位已经被预定30024
    ERR_BookTimeOut = ERR_Book + 25 '预定时间已过30025
    
    ERR_SeatNotBooked = ERR_Book + 26 '指定的座位没有被预定30026
    ERR_BusNotExist = ERR_Book + 27 '指定的车次不存在30027
    ERR_BusNotRegularBus = ERR_Book + 28 '指定的车次不是固定车次30028
    ERR_BusStationMatchLimitedCount = ERR_Book + 29 '已到指定车次到站的限售张数30029
    ERR_SeatBookedCanBeCancel = ERR_Book + 30 '座位只有处于预定状态才能被取消预定30030
  
    ERR_NoPushMan = ERR_Book + 31 '没有指定刷卡售票用户30031
End Enum

Public Enum EBookStatus
    ST_BOOKED = 0 '已被预定
    ST_BOOKCANCELED = 1 '取消了预定
    ST_BOOKSOLD = 2 '已被预定人买走
    ST_BOOKFREESOLD = 3 '超时被自由购买
    ST_CHANGE = 4 '转预留
End Enum

Private Enum ETicketClientType
    TP_CancelTicket = 1 '废票
    TP_ReturnTicket = 2 '退票
    TP_ChangeTicket = 3 '改签
End Enum
    


Public Type TMultiSeatTypeBus2
    adBusDate() As Date
    aszBusID() As String
End Type

Private Type TUserLoginUnitInfo
    szLoginUserID As String
    szPassword As String
    szIPAddress As String
End Type

Dim m_clSellService As New Collection
Dim m_oCurrentSellService As NetSellTicketService
Dim m_oServerProxy As New ServerProxy

Dim m_oCurrentServiceTicket As ServiceTicket
Private m_oActiveUser As ActiveUser
Private m_oParam As Object
Private m_bSellStationCanSellEachOther As Boolean

Private oSystemAgent As SystemAgent
Private oAgentActiveUser As ActiveUser

Private m_bAllowScrollBusSaleForever As Boolean
Private m_bFirst As Boolean '是否为远程单位
Private m_szCurrentUnitID As String '本地单位

Private bIsUsingRelationBus As Boolean
Const SeatTypeID = "01"
Const cszEmptyDate = "0:00:00"
Const cszNOPrint = 0
Public Type TTicketTypeInfo '所选择票种信息
    szTicketName As String
    szTicketNum As Integer
    szTicketTypeName As String
End Type
Public Function GetServiceUnitInfo()
    'SNProxy.ServerProxy
    'm_oServerProxy.GetServiceUnitInfo
    GetServiceUnitInfo = m_oServerProxy.GetServiceUnitInfo()
End Function

'///////////////////////////////
Public Function GetTicketClient(TicketID As String) As ClientTicket
    Dim oDb As New RTConnection
    Dim szsql As String
    Dim dtReturnDateTime As Date
    Dim oParam As New SystemParam
    Dim rsTemp As Recordset
    Dim rsTemp1 As Recordset
    Dim szTemp As String
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    
    szsql = "SELECT service_unit_id FROM Agent_ticket_sale_lst WHERE " _
    & "ticket_id='" & TicketID & "'"
    Set rsTemp = oDb.Execute(szsql)
    If rsTemp.RecordCount <> 1 Then
        szsql = "SELECT u.unit_id,service_type FROM Ticket_sell_lst s,user_info u,connect_unit_info c WHERE s.ticket_id='" & TicketID & "' AND " _
                  & " s.user_id=u.user_id  AND u.unit_id=c.unit_id "
                  
        Set rsTemp = oDb.Execute(szsql)
        
        If rsTemp.RecordCount <> 1 Then ShowError ERR_TicketNotExist
        If Trim(rsTemp!service_type) <> TP_UnitClient Then
            
'            szTemp = m_oServerProxy.CurrentServiceUnit.szUnitID
'            m_oServerProxy.ChangeServiceUnitUserBuffer (rsTemp!unit_id)
'            Set GetTicketClient = m_oServerProxy.CreateServiceObject("STTkSvr.ClientTicket")
'            m_oServerProxy.ChangeServiceUnitUserBuffer szTemp
'            GetTicketClient.Identify TicketID
        Else
            Set GetTicketClient = Nothing
        End If
    Else
        
        szTemp = m_oServerProxy.CurrentServiceUnit.szUnitID
        m_oServerProxy.ChangeServiceUnitUserBuffer (szTemp)
        Set GetTicketClient = m_oServerProxy.CreateServiceObject("ST6TkSvr.ClientTicket")

        m_oServerProxy.ChangeServiceUnitUserBuffer szTemp


        GetTicketClient.Identify TicketID
    End If
End Function

Public Function GetUnitIP(pszUnitID As String) As String
    Dim oDb As New RTConnection
    Dim szsql As String
    Dim rsTemp As Recordset
    Dim szUnitIP As String
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
On Error GoTo ErrorHandle
    szsql = "SELECT ip_address FROM connect_unit_info WHERE unit_id='" & pszUnitID & "'"
    Set rsTemp = oDb.Execute(szsql)
    If rsTemp.RecordCount = 1 Then
        szUnitIP = FormatDbValue(rsTemp!ip_address)
    End If
    GetUnitIP = szUnitIP
    Exit Function
ErrorHandle:
    err.Raise err.Number
End Function

Public Function GetBusExRs(BusDate As Date, BusID As String, SellStationID As String, EndStationID As String) As Recordset
     If Trim(m_oCurrentSellService.SelfUser.ServerUnit) <> Trim(m_oCurrentSellService.SelfUser.UserUnitID) Then
        '如果用户单位和服务单位不一样则是售远程票否则售的是本地票
        Set GetBusExRs = m_oCurrentSellService.GetBusExRs(BusDate, BusID, SellStationID, EndStationID)
    Else
        Set GetBusExRs = GetLocalBusEXRs(BusDate, BusID, SellStationID, EndStationID)
    End If
        
End Function

'得到对方车站的线路名称和发车时间

Public Function GetRouteAndTime(pdtBusDate As Date, pszBusID As String) As String()
On Error GoTo Here
    GetRouteAndTime = m_oCurrentSellService.GetRouteAndTime(pdtBusDate, pszBusID)
Exit Function
Here:
    err.Raise err.Number
End Function
'///////////////////////////
'得到对方车站的所有的检票口
Public Function GetAllCheckGate(Optional CheckGate As String = "") As String()
 
    GetAllCheckGate = m_oCurrentSellService.GetAllCheckGate(CheckGate)
    
End Function
'得到对方车站座位类型代码和名称
Public Function GetAllSeatType() As String()

GetAllSeatType = m_oCurrentSellService.GetAllSeatType

End Function
'当前提供售票服务的单位代码
Public Property Let SellUnitCode(ByVal vData As String)
    Dim nErrCount As Integer: nErrCount = 0
    Dim szOldSellUnitCode As String
    Dim OldSellSellService As NetSellTicketService
    
    szOldSellUnitCode = SellUnitCode
    Set OldSellSellService = m_oCurrentSellService
    
    vData = Trim(vData)
    If Trim(vData) = Trim(SellUnitCode) And Not m_oCurrentSellService Is Nothing Then Exit Property
    
    On Error GoTo Error_Handle
GoOn:
    Set m_oCurrentSellService = m_clSellService(GetEncodedKey(vData))
    m_oServerProxy.ChangeServiceUnitUserBuffer Trim(vData)
    Set m_oParam = m_oServerProxy.CreateServiceObject("STSystemBase.SystemParam")
    
    Exit Property
CreateIt:
    '先改变当前的代理对象的服务站点,以后并通过该代理创建远程对象
    nErrCount = nErrCount + 1 '如果系统创建对象超过5次则认为系统错误
    m_oServerProxy.ChangeServiceUnitUserBuffer vData
    Set m_oCurrentSellService = m_oServerProxy.CreateServiceObject("STTkSvr.NetSellTicketService")
    m_clSellService.Add m_oCurrentSellService, GetEncodedKey(vData)
    GoTo GoOn
Error_Handle:
    If nErrCount >= 4 Then
        err.Raise err.Number
        Exit Property
    End If
    If err.Number = 5 Then
        Resume CreateIt
    End If
    SellUnitCode = szOldSellUnitCode
    Set m_oCurrentSellService = OldSellSellService
    
    err.Raise err.Number
End Property
'P2
Public Property Get SellUnitCode() As String
    SellUnitCode = m_oServerProxy.CurrentServiceUnit.szUnitID
End Property

'当前服务单位的简称
Public Property Get SellUnitShortName() As String
    SellUnitShortName = m_oServerProxy.CurrentServiceUnit.szUnitShortName
End Property

'当前服务单位的名称
Public Property Get SellUnitName() As String
    SellUnitName = m_oServerProxy.CurrentServiceUnit.szUnitFullName
End Property


Public Property Get SelfUser() As ActiveUser
    Set SelfUser = m_oActiveUser
    m_oActiveUser.IAmLive
End Property

Public Property Set SelfUser(vNewValue As ActiveUser)
    If vNewValue.UserUnitID <> vNewValue.ServerUnit Then ShowError ERR_UserMustBeLocalUser
    Set m_oActiveUser = vNewValue
    m_oServerProxy.Init m_oActiveUser
    SellUnitCode = m_oActiveUser.ServerUnit
End Property


'售票
Public Function SellTicket(BusDate() As Date, BusID() As String, SellStationID() As String, DesStationID() As String, DesStationName() As String, aspSellTicketParam() As TSellTicketParam, anInsurance() As Integer, bScroll As Boolean) As TSellTicketResult()
    'anInsurance 是 保险费
    
    Dim srTemp() As TSellTicketResult
    Dim iLen As Integer
    Dim iLen1 As Integer
    Dim nCount As Integer
    Dim aszDesStationName() As String
    
    iLen = 0
    iLen1 = 0
    iLen = ArrayLength(BusID)
    ReDim srTemp(1 To iLen)
    If Trim(m_oCurrentSellService.SelfUser.ServerUnit) <> Trim(m_oCurrentSellService.SelfUser.UserUnitID) Then
        '如果用户单位和服务单位不一样则是售远程票否则售的是本地票
        On Error GoTo ErrorHandle
        
        srTemp = m_oCurrentSellService.SelfSellTK(BusDate, BusID, SellStationID, DesStationID, DesStationName, aspSellTicketParam, anInsurance, bScroll)
        
        iLen1 = ArrayLength(srTemp)
        ReDim aszDesStationName(1 To iLen1)
        For nCount = 1 To iLen1
            aszDesStationName(nCount) = srTemp(nCount).szDesStationName
        Next nCount
        On Error GoTo 0
        On Error GoTo ErrorHandle2
        SelfSellTicket BusDate, BusID, SellStationID, DesStationID, DesStationName, aspSellTicketParam, srTemp, anInsurance
        On Error GoTo 0
    Else
        srTemp = SelfSellTK(BusDate, BusID, SellStationID, DesStationID, DesStationName, aspSellTicketParam, anInsurance, bScroll)
    End If
    SellTicket = srTemp
    Exit Function
ErrorHandle:
    err.Raise err.Number
    Exit Function
ErrorHandle2:
    '加入提示"对方车站已写入,请通知管理员在对方车站将此票删除"
    err.Raise err.Number, err.Source, cszServiceFinish & err.Description
    Exit Function
End Function
'内部售票函数
Private Sub SelfSellTicket(BusDate() As Date, BusID() As String, SellStationID() As String, DesStationID() As String, DesStationName() As String, SellTicket() As TSellTicketParam, srTemp() As TSellTicketResult, panInsurance() As Integer, Optional pnTicketStatus As ETicketStatus = ST_TicketNormal)
    Dim m_oSellTicketService As Object
    Dim nBuyTicketCount As Integer
    Dim nIndex As Integer
    Dim szsql As String
    Dim oDb As New RTConnection
    Dim oParam As New SystemParam
    Dim dtTime As Date
    
    Dim szStartStationID As String, szStartStationName As String
    Dim szUnitID As String
    Dim nCount, iLen As Integer
    
    Dim nSellTicketType As Integer
    Dim nSellTicketStatus As ETicketStatus
    
    nSellTicketStatus = pnTicketStatus
    
    Select Case pnTicketStatus
        Case ST_TicketReturned   '补票售出
            nSellTicketType = TP_ExtraSold
            nSellTicketStatus = ST_TicketNormal
            
        Case ST_TicketSellChange    '改签售出
            nSellTicketType = TP_ChangeSold
        Case Else           '其它
            nSellTicketType = 0
    End Select
    
    iLen = 0
    iLen = ArrayLength(BusID)
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    oParam.Init m_oActiveUser
    dtTime = oParam.NowDateTime
    
    'AssertActiveUserValid m_oActiveUser
    'AssertHaveRight m_oActiveUser, RIGHT_SellTicketClient
    
    
    On Error GoTo Error_Handle
    szStartStationID = m_oCurrentSellService.StationID
    szStartStationName = m_oCurrentSellService.StationName
    szUnitID = m_oCurrentSellService.SelfUser.ServerUnit
    
    oDb.BeginTrans
    For nCount = 1 To iLen
        nBuyTicketCount = ArrayLength(SellTicket(nCount).BuyTicketInfo)
        For nIndex = 1 To nBuyTicketCount
            '写售出的票的信息
            szsql = "INSERT INTO Agent_ticket_sale_lst(" _
            & " sell_station_id," _
            & " ticket_id," _
            & " bus_date," _
            & " bus_id," _
            & " service_unit_id," _
            & " start_station_id," _
            & " start_station_name," _
            & " des_station_id," _
            & " des_station_name," _
            & " ticket_type," _
            & " user_id," _
            & " seat_no," _
            & " ticket_price," _
            & " status, " _
            & " sale_ticket_type, " _
            & " operation_time , " _
            & " insurance ) "
            
            
            szsql = szsql & " VALUES('" _
            & SellStationID(nCount) & "','" _
            & SellTicket(nCount).BuyTicketInfo(nIndex).szTicketNo & "','" _
            & ToDBDate(BusDate(nCount)) & "','" _
            & BusID(nCount) & "','" _
            & szUnitID & "','" _
            & szStartStationID & "','" _
            & szStartStationName & "','" _
            & DesStationID(nCount) & "','" _
            & srTemp(nCount).szDesStationName & "'," _
            & SellTicket(nCount).BuyTicketInfo(nIndex).nTicketType & ",'" _
            & m_oActiveUser.UserID & "','" _
            & srTemp(nCount).aszSeatNO(nIndex) & "'," _
            & srTemp(nCount).asgTicketPrice(nIndex) & "," _
            & nSellTicketStatus & "," & nSellTicketType & ",'" _
            & ToDBDateTime(dtTime) & "', " _
            & TransFieldValueToString(panInsurance(nCount)) & " ) "
            
            oDb.Execute szsql
        Next nIndex
    Next nCount
    oDb.CommitTrans
    Exit Sub
Error_Handle:
    oDb.RollbackTrans
    err.Raise err.Number
End Sub





'得到指定日期得到指定车站的所有车次及其相关信息。
'日期比较
Public Function GetBusRs(BusDate As Date, DesStationID As String, Optional IgnoreOffTime As Boolean = False, Optional BusID As String = "", Optional IsDiscount As Boolean = False) As Recordset
     If Trim(m_oCurrentSellService.SelfUser.ServerUnit) <> Trim(m_oCurrentSellService.SelfUser.UserUnitID) Then
        '如果用户单位和服务单位不一样则是售远程票否则售的是本地票
        Set GetBusRs = m_oCurrentSellService.GetBusRs(BusDate, DesStationID, IgnoreOffTime, BusID, IsDiscount)
    Else
        Set GetBusRs = GetLocalBusRs(BusDate, DesStationID, IgnoreOffTime, BusID, IsDiscount)
    End If
        
End Function
Public Function GetBusRsEx(BusDate As Date, DesStationID As String, Optional pszCheckGateID As String = "", Optional IgnoreOffTime As Boolean = False, Optional BusID As String = "") As Recordset
'    If Trim(m_oCurrentSellService.SelfUser.ServerUnit) <> Trim(m_oCurrentSellService.SelfUser.UserUnitID) Then
'        '如果用户单位和服务单位不一样则是售远程票否则售的是本地票
'        Set GetBusRsEx = m_oCurrentSellService.GetBusRs(BusDate, DesStationID, IgnoreOffTime, BusID)
'    Else
'        Set GetBusRsEx = GetLocalBusRsEx88(BusDate, DesStationID, pszCheckGateID, IgnoreOffTime, BusID)
'    End If
    
End Function
'得到指定日期能到指定车站的所有车次及其相关信息
'日期比较
Public Function GetLocalBusRs(BusDate As Date, DesStationID As String, Optional IgnoreOffTime As Boolean = False, Optional BusID As String = "", Optional IsDiscount As Boolean) As Recordset
    'IsDiscount 是否折扣票查询
    Dim oDb As New RTConnection
    Dim oParam As New SystemParam
    Dim szsql As String
    Dim rsTemp As Recordset
    Dim nStopSellTime As Integer
    Dim dtNowDateTime As Date

    
    Dim dbExtraBeforeTime As Double
    Dim dbExtraAfterTime As Double
    
    Dim szSelect As String
    Dim szTable As String
    Dim szWhere As String
    
    oParam.Init m_oActiveUser
    nStopSellTime = oParam.StopSellTime
    dtNowDateTime = oParam.NowDateTime
    
    
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    
    szSelect = "SELECT Work_env_bus_info.* ,e.check_gate_id sell_check_gate_id ,seat_type_id,route_name ,Work_env_bus_station_lst.mileage end_station_mileage" _
        & " ,Work_env_bus_station_lst.sell_station_id,sell_station_name,e.bus_start_time as busstarttime, " _
        & "         full_price, half_price,preferential_ticket1,preferential_ticket2,preferential_ticket3," _
        & "         sale_ticket_quantity,stop_sale_time , Work_env_bus_station_lst.book_quantity book_quantity "
        
    szTable = "  FROM Work_env_bus_info,Work_env_bus_station_lst ,route_info,sell_station_info s,Work_env_bus_allot_lst e  "
    
    
    szWhere = "  WHERE s.sell_station_id=Work_env_bus_station_lst.sell_station_id AND e.sell_station_id=s.sell_station_id AND " _
        & "  e.bus_date=Work_env_bus_info.bus_date AND e.bus_id=Work_env_bus_info.bus_id AND "
    
    If Not IgnoreOffTime Then
        If Not IsDiscount Then
            szWhere = szWhere & "((e.bus_start_time>='" & ToDBDateTime(DateAdd("s", nStopSellTime * 60, dtNowDateTime)) & "'"
        Else
            '如果为折扣票,则过了发车时间10分钟,仍显示
            szWhere = szWhere & "((e.bus_start_time>='" & ToDBDateTime(DateAdd("s", -600, dtNowDateTime)) & "' AND e.bus_start_time < '" & ToDBDateTime(DateAdd("n", 30, dtNowDateTime)) & "'"
        End If
        
        If Not IsDiscount Then
            szWhere = szWhere & IIf(Not m_bAllowScrollBusSaleForever, "", " AND bus_kind<>" & TP_ScrollBus) _
                & " )  OR ( " _
                & IIf(Not m_bAllowScrollBusSaleForever, "", " bus_kind=" & TP_ScrollBus & " OR ") _
                & " e.bus_start_time>='" & ToDBDateTime(DateAdd("s", nStopSellTime * 60, dtNowDateTime)) & "'" _
                & " )) AND  "
        Else
            szWhere = szWhere & " )) and "
        End If
        szWhere = szWhere _
            & " Work_env_bus_info.bus_date='" & ToDBDate(BusDate) & "' AND " _
            & " Work_env_bus_info.bus_id=Work_env_bus_station_lst.bus_id AND " _
            & " Work_env_bus_info.bus_date=Work_env_bus_station_lst.bus_date AND " _
            & " Work_env_bus_station_lst.station_id='" & DesStationID & "' AND " _
            & " ( sale_seat_quantity>=0 OR sale_stand_seat_quantity>0 " _
            & IIf(Not m_bAllowScrollBusSaleForever, "", " OR bus_kind=" & TP_ScrollBus) _
            & ")  "
        If Not IsDiscount Then
            szWhere = szWhere & " AND ( ( max_sale_quantity>0 AND  sale_ticket_quantity>=0) OR  ( max_sale_quantity<0) ) "
        Else
            '如果为折扣票,则距发车时间不到半小时内的所有站点,如果该站点不可售,也可以售折扣票
            szWhere = szWhere & " AND ( ( max_sale_quantity>0 AND  sale_ticket_quantity>=0) OR  ( max_sale_quantity<0) or (datediff(n,getdate(),e.bus_start_time)<30 ))"
        End If
        szWhere = szWhere & " AND Work_env_bus_info.route_id=route_info.route_id AND (e.can_sale_quantity <>0 )" _
    '        & " AND  rs.sell_station_id=e.sell_station_id AND Work_env_bus_info.route_id=rs.route_id AND rs.end_station_id=Work_env_bus_station_lst.station_id "
            
            'ORDER BY Work_env_bus_info.bus_start_time,Work_env_bus_info.bus_id  "
    
    Else
        dbExtraBeforeTime = oParam.ExtraBusShowBefore
        dbExtraAfterTime = oParam.ExtraBusShowAfter
        
        szWhere = szWhere & "'" & ToDBDateTime(DateAdd("n", dbExtraAfterTime, dtNowDateTime)) & "'>=e.bus_start_time AND " _
            & "'" & ToDBDateTime(DateAdd("n", -dbExtraBeforeTime, dtNowDateTime)) & "'<=e.bus_start_time " _
            & IIf(Not m_bAllowScrollBusSaleForever, "", " AND bus_kind<>" & TP_ScrollBus) _
            & " AND Work_env_bus_info.bus_date='" & ToDBDate(BusDate) & "' AND " _
            & " Work_env_bus_info.bus_id=Work_env_bus_station_lst.bus_id AND " _
            & " Work_env_bus_info.bus_date=Work_env_bus_station_lst.bus_date AND " _
            & " Work_env_bus_station_lst.station_id='" & DesStationID & "' AND " _
            & " ( sale_seat_quantity>=0 OR sale_stand_seat_quantity>0" _
            & IIf(Not m_bAllowScrollBusSaleForever, "", " OR bus_kind=" & TP_ScrollBus) _
            & ")  " _
            & " AND ( ( max_sale_quantity>0 AND  sale_ticket_quantity>=0) OR  ( max_sale_quantity<0) ) " _
            & " AND Work_env_bus_info.route_id=route_info.route_id "
            'ORDER BY Work_env_bus_info.bus_start_time,Work_env_bus_info.bus_id "
        
    End If
    
    If BusID <> "" Then
           szWhere = szWhere & " AND Work_env_bus_info.bus_id='" & Trim(BusID) & "'"
    End If
    
    If Not m_bSellStationCanSellEachOther Then
        If m_oActiveUser.SellStationID <> "" Then
            szWhere = szWhere & " AND Work_env_bus_station_lst.sell_station_id='" & Trim(m_oActiveUser.SellStationID) & "'"
        End If
    Else
        szTable = szTable & " , work_env_bus_sell_station_lst t "
        szWhere = szWhere & " AND t.bus_id = Work_env_bus_info.bus_id " _
            & " AND t.bus_date = Work_env_bus_info.bus_date " _
            & " AND t.sell_station_id = '" & Trim(m_oActiveUser.SellStationID) & "'   and (t.can_sale_quantity >=t.sell_quantity or t.can_sale_quantity <0) "
            
    End If
  '  szWhere = szWhere & "and left(e.bus_id,1)<>'t'"'为温州大学城限死用的，那里T开的车次不能在自助机里卖
    szWhere = szWhere & " ORDER BY e.bus_start_time,Work_env_bus_info.bus_id "
    szsql = szSelect & szTable & szWhere
    
    Set rsTemp = oDb.Execute(szsql)
    Set GetLocalBusRs = rsTemp
    Set rsTemp = Nothing
    Set oParam = Nothing
    Set oDb = Nothing
    
End Function

'得到指定日期的所有可能到站及其相关信息。
'日期比较
Public Function GetStationRs(BusDate As Date, Optional IgnoreOffTime As Boolean = False) As Recordset
    Set GetStationRs = m_oCurrentSellService.GetStationRs(BusDate, IgnoreOffTime)
End Function
Public Sub Init(AUser As ActiveUser)
    Set SelfUser = AUser
    m_oParam.Init SelfUser
    m_bSellStationCanSellEachOther = False  '自助售票中设置,嘉兴为了防止旅客不会选，所以设死
    
    m_bAllowScrollBusSaleForever = m_oParam.AllowScrollBusSaleForever
    
'    bIsUsingRelationBus = m_oParam.IsUsingRelationBus

End Sub
'得到对方车站的所有票种
Public Function GetAllTicketType(Optional nValid As Integer = TP_TicketTypeAll) As TTicketType()
On Error GoTo Here
    GetAllTicketType = m_oCurrentSellService.GetAllTicketType(nValid)
    Exit Function
Here:
    err.Raise err.Number
End Function
'得到指定售票员最后的一张票号
'fpd修改,因为当最后一张票为退票时，票号为退票号而非售票号
Public Function GetLastTicketNo(Seller As String) As String
    Dim szsql As String
    Dim szSql2 As String
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim rsTemp2 As Recordset
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    
'    szSql = " SELECT MAX(operation_time) as operation_time FROM Agent_ticket_sale_lst WHERE " _
'    & " user_id='" & Seller & "' "
'    Set rstemp = oDb.Execute(szSql)
'
'    szSql2 = " SELECT MAX(operation_time) as operation_time FROM agent_ticket_return_lst WHERE " _
'    & " user_id='" & Seller & "' "
'    Set rsTemp2 = oDb.Execute(szSql2)
    
'    If rstemp.RecordCount = 1 Or rsTemp2.RecordCount = 1 Then
'        If FormatDbValue(rstemp!operation_time) > FormatDbValue(rsTemp2!operation_time) Then
            szsql = "SELECT MAX(ticket_id) AS ticket_no FROM ticket_sell_lst WHERE " _
            & " user_id='" & Seller & "' "
            Set rsTemp = oDb.Execute(szsql)
            GetLastTicketNo = FormatDbValue(rsTemp!ticket_no)
'        Else
'            szSql = "SELECT MAX(credence_sheet_id) AS ticket_no FROM agent_ticket_return_lst WHERE " _
'            & " operation_time='" & rsTemp2!operation_time & "' AND " _
'            & " user_id='" & Seller & "' "
'            Set rstemp = oDb.Execute(szSql)
'            GetLastTicketNo = FormatDbValue(rstemp!ticket_no)
'        End If
'    Else
'        GetLastTicketNo = String(m_oParam.TicketNumberLen, "0")
'    End If
End Function


'得到当前系统中的所有站点信息
Public Function GetAllStationRs() As Recordset
    Set GetAllStationRs = m_oCurrentSellService.GetAllStationRs()
End Function

'内部本地售票
Public Function SelfSellLocalTicket(BusStartupDate() As Date, BusID() As String, SellStationID() As String, DesStationID() As String _
             , aspSellTicketParam() As TSellTicketParam, panInsurance() As Integer, Optional oDb As RTConnection) As TSellTicketResult()
    'panInsurance 保险费
    
'    Dim oDb As New RTConnection
    Dim rsTemp() As TSellTicketResult
    Dim nCount As Integer
    Dim iLen As Integer
    iLen = ArrayLength(BusID)

    
    AssertActiveUserValid m_oActiveUser, ERR_SellTicketService
    AssertHaveRight m_oActiveUser, RIGHT_SellTicket
    ReDim rsTemp(1 To iLen)

'    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    
    oDb.BeginTrans
    On Error GoTo Here
    For nCount = 1 To iLen
        rsTemp(nCount) = SelfSellTicketLocalTrans(BusStartupDate(nCount), BusID(nCount), SellStationID(nCount), DesStationID(nCount), aspSellTicketParam(nCount).BuyTicketInfo, aspSellTicketParam(nCount).pasgSellTicketPrice, panInsurance(nCount), False, ST_TicketNormal, oDb)
    Next nCount
    SelfSellLocalTicket = rsTemp
    oDb.CommitTrans
    Exit Function
Here:
    oDb.RollbackTrans

    
        err.Raise err.Number
End Function


'内部售票事务函数
Private Function SelfSellTicketLocalTrans(BusStartupDate As Date, BusID As String, SellStationID As String, DesStationID As String, SellTicketInfo() As TBuyTicketInfo _
    , pasgSellTicketPrice() As Single, pnInsurance As Integer, Optional pbExtraSell As Boolean = False _
    , Optional pnTicketStatus As ETicketStatus = ST_TicketNormal, Optional poDb As RTConnection) As TSellTicketResult
    
    Dim oParam As New SystemParam
    Dim szsql As String
    Dim rsTemp As Recordset
    Dim dtBusFullDateTime As Date
    Dim nTemp As Double
    Dim dtNow As Date
    Dim i As Integer
    Dim nBusStatus As EREBusStatus
    Dim nBuyTicketCount As Integer
    Dim sgFullPrice As Double, sgHalfPrice As Double
    Dim lRowAffect As Long, nStationLimitCount As Integer
    Dim szSeatNo As String, sgTicketPrice As Double
    Dim nBusType As EBusType
    Dim atiTicketPrice() As TRETicketPrice
    Dim siSellInfo As TSellTicketResult
    
    Dim lErrorCode As Long
    Dim rsRange As Recordset
    
    Dim nStationRemainCount As Integer
    
    Dim sgPreferentialTicket1 As Single '折扣票1
    Dim sgPreferentialTicket2 As Single '折扣票2
    Dim sgPreferentialTicket3 As Single '折扣票3
    
    
    Dim bNeedBook As Boolean '是否上了预定模块
    Dim bBookValid As Boolean '是否还需判断预定
    Dim bSeatBooked As Boolean
    Dim szBookStationID As String '预定的到站
    Dim rsTemp2 As Recordset
    
    Dim szSeatType As String '座位类型
    Dim szSqlSeat As String  '座位票价
    Dim nSeatStatus As Integer '座位状态
    
    Dim szStopTime As String
    
    Dim nSellTicketType As Integer
    Dim nSellTicketStatus As ETicketStatus
    
    
    Dim bSellStationCanSellEachOther As Boolean
    
    
    ReDim atiTicketPrice(1 To TP_TicketTypeCount)
    
    nSellTicketStatus = pnTicketStatus
    
    Select Case pnTicketStatus
        Case ST_TicketReturned
            nSellTicketType = TP_ExtraSold
            nSellTicketStatus = ST_TicketNormal
            
        Case ST_TicketSellChange
            nSellTicketType = TP_ChangeSold
        Case ST_TicketNormal
            nSellTicketType = TP_NormalSold
        Case Else
            nSellTicketType = 0
        
    End Select
    
    oParam.Init m_oActiveUser
    '判断车次的日期是否超过预售期限
    
    dtNow = oParam.NowDateTime
    
    nTemp = DateDiff("d", dtNow, BusStartupDate)
    
    '如果要购票的车次不在当天和预售天之内则出错
    If nTemp > oParam.PreSellDate Or nTemp < 0 Then ShowError ERR_BusNotInIlegalDate
    If nTemp > oParam.UserCanSellDate Then ShowError ERR_UserCannotSell
    
    bSellStationCanSellEachOther = m_oParam.SellStationCanSellEachOther
    'poDb.ConnectionString = GetConnectionStr(cszSellTicket)
    '得到环境车次及车次站点信息
'    szSql = "SELECT * FROM work_env_bus_info,work_env_bus_station_lst " _
'    & " WHERE work_env_bus_info.bus_id='" & BusID & "' AND " _
'    & " work_env_bus_info.bus_date='" & ToDBDate(BusStartupDate) & "' AND " _
'    & " work_env_bus_info.bus_id=work_env_bus_station_lst.bus_id AND " _
'    & " work_env_bus_info.bus_date=work_env_bus_station_lst.bus_date AND " _
'    & " station_id='" & DesStationID & "'" _
'    & " AND work_env_bus_station_lst.sell_station_id='" & SellStationID & "'"
'    Set rsTemp = poDb.Execute(szSql)


    szsql = "SELECT e.bus_date , e.bus_id , e.vehicle_type_name , e.vehicle_type_code , e.total_seat " _
        & " , e.sale_seat_quantity , e.total_stand_quantity , e.sale_stand_seat_quantity , e.bus_start_time , e.bus_kind" _
        & " , e.is_all_refundment , e.route_id , e.end_station_id , e.end_station_name" _
        & " , e.owner_id , e.transport_company_id , e.vehicle_id , e.split_company_id , e.baggage_number , e.fact_weight" _
        & " , e.calculate_weight , e.over_weight_number , e.luggage_total_price  , e.scrollbus_check_time" _
        & " , e.seat_remain , e.bed_remain  , e.additional_remain , e.other_remain_1 , e.other_remain_2" _
        & " , s.sell_station_id , s.station_id , s.seat_type_id , s.station_serial_no , s.max_sale_quantity , s.sale_ticket_quantity" _
        & " , s.book_quantity , s.stop_sale_time , s.mileage  , s.full_price , s.half_price , s.preferential_ticket1 , s.preferential_ticket2" _
        & " , s.preferential_ticket3" _
        & " , a.check_gate_id , a.bus_start_time" _
        & " , a.status , a.can_sale_quantity , a.sell_quantity" _
        & " FROM work_env_bus_info e ,work_env_bus_station_lst s , work_env_bus_allot_lst a" _
        & " WHERE e.bus_id='" & BusID & "' AND  e.bus_date='" & ToDBDate(BusStartupDate) & "'" _
        & " AND  e.bus_id=s.bus_id AND  e.bus_date=s.bus_date" _
        & " AND  s.station_id='" & DesStationID & "' AND s.sell_station_id='" & SellStationID & "'" _
        & " AND a.bus_id = s.bus_id AND a.bus_date = s.bus_date AND a.sell_station_id = '" & SellStationID & "'" _
        
        
    Set rsTemp = poDb.Execute(szsql)
    
    
    
    
    If rsTemp.RecordCount = 0 Then ShowError ERR_NoThisBus '没有指定的车次或车次不过指定的到站
    nBusType = FormatDbValue(rsTemp!bus_kind)
    '流水车次不能补票
    If nBusType = TP_ScrollBus And pbExtraSell Then ShowError ERR_ScrollBusCanNotExtraSell
    nBusStatus = FormatDbValue(rsTemp!Status)
    '并班车次不能售票
    If nBusStatus = ST_BusMergeStopped Then ShowError ERR_BusMerged
    '拆分停班车次不能售票
    If nBusStatus = ST_BusSlitpStop Then ShowError ERR_BusSlitpStop
    '车次正在拆分不能售票
    If nBusStatus >= 32 Then ShowError ERR_BusIsSplit
    '停检之后不能售票（实际不太可能出现，除非停售时间设得太小了）如果是补票则不限制
'    If pbExtraSell And nBusStatus = ST_BusStopCheck Then ShowError ERR_BusStopCheck
    If nBusStatus = ST_BusStopCheck Then ShowError ERR_BusStopCheck
    If nBusStatus = ST_BusStopped Then ShowError ERR_BusStopped
    dtBusFullDateTime = CDate(Format(FormatDbValue(rsTemp!bus_date), cszDateStr) & " " & Format(FormatDbValue(rsTemp!bus_start_time), cszTimeStr))
    '判断是否到了停售时间(如果是补票则要过了正常售票时间才能补)
    If nBusType <> TP_ScrollBus Then '如果是固定车次才限制
        If pbExtraSell Then
            If DateDiff("s", dtNow, dtBusFullDateTime) >= oParam.StopSellTime * 60 Then ShowError ERR_BusNotStopSellTicket
        Else
'            If DateDiff("s", dtNow, dtBusFullDateTime) < oParam.StopSellTime * 60 Then ShowError ERR_BusStopSellTicket
        End If
    End If
    szStopTime = FormatDbValue(rsTemp!stop_sale_time)
    If nBusType <> TP_ScrollBus Then '如果是固定车次才限制
        If szStopTime <> "-1" Then '如果该车次到指定的站限制售票时间则要进行判断
            If DateDiff("n", dtNow, dtBusFullDateTime) > GetStopTime(szStopTime) Then ShowError ERR_BusStaionStopSellTicket, "SellTicketClient.SelfSellTicketLocalTrans"
        End If
        bNeedBook = oParam.IsBookValid
        If bNeedBook Then
            bBookValid = IIf(DateAdd("n", oParam.BookTime, dtNow) >= rsTemp!bus_start_time, False, True)
        End If
    Else
        If szStopTime <> "-1" Then
           dtBusFullDateTime = CDate(Format(FormatDbValue(rsTemp!bus_date), cszDateStr) & " " & GetStopTime(szStopTime, False))
           If DateDiff("n", dtNow, dtBusFullDateTime) < 0 Then ShowError ERR_BusStaionStopSellTicket, "SellTicketService.SelfSellTicket"
        End If
    End If
    nBuyTicketCount = ArrayLength(SellTicketInfo)
    If nBuyTicketCount > 0 Then '要售票的张数大于0
        nStationLimitCount = FormatDbValue(rsTemp!max_sale_quantity)
        nStationRemainCount = FormatDbValue(rsTemp!sale_ticket_quantity)
        If nStationRemainCount > 0 Then  '此处为了折扣票由>= 改为了>
            If nBuyTicketCount > nStationRemainCount Then ShowError ERR_BusReachMaxStationTicketCount
        End If
        
        sgFullPrice = FormatDbValue(rsTemp!full_price)
        sgHalfPrice = FormatDbValue(rsTemp!half_price)
        sgPreferentialTicket1 = FormatDbValue(rsTemp!preferential_ticket1)
        sgPreferentialTicket2 = FormatDbValue(rsTemp!preferential_ticket2)
        sgPreferentialTicket3 = FormatDbValue(rsTemp!preferential_ticket3)

        ReDim siSellInfo.asgTicketPrice(1 To nBuyTicketCount)
        ReDim siSellInfo.aszSeatNO(1 To nBuyTicketCount)
        ReDim siSellInfo.aszTicketType(1 To nBuyTicketCount)
        ReDim siSellInfo.aszSeatType(1 To nBuyTicketCount)
        For i = 1 To nBuyTicketCount
            szSeatNo = ""
            Select Case SellTicketInfo(i).nTicketType
            Case TP_FullPrice, TP_FreeTicket
                sgTicketPrice = sgFullPrice
            Case TP_HalfPrice
                sgTicketPrice = sgHalfPrice
            Case TP_PreferentialTicket1
                sgTicketPrice = sgPreferentialTicket1
            Case TP_PreferentialTicket2
                sgTicketPrice = sgPreferentialTicket2
            Case TP_PreferentialTicket3
                sgTicketPrice = sgPreferentialTicket3
            End Select
            
            
            
            
            If nBusType <> TP_ScrollBus Then '如果是固定车次才限制
                '判断到站的限售张数够用否
                If nStationLimitCount >= 0 Then
                    If nStationLimitCount = 0 Then
'                        ShowError ERR_BusReachMaxStationTicketCount
                    End If
                End If
            End If
            If nBusType <> TP_ScrollBus Then '如果是固定车次才进行座位处理
                If SellTicketInfo(i).szSeatNo <> cszStandTicketSeatNo Then
                    If SellTicketInfo(i).szSeatNo = cszAutoSeatNo Then  '系统自动给出座位号
RetrySeatNo:
                        '取座位该票价和类型
                        If bNeedBook And Not bBookValid Then
                            '如果可以预定
                            szsql = "SELECT p.ticket_price_total AS TicketPrice, s.seat_no AS SeatNo,s.status,p.seat_type_id AS SeatType FROM Env_bus_seat_lst s,work_env_bus_price_lst p" _
                            & " WHERE " _
                            & " s.bus_id='" & BusID & "' AND " _
                            & " s.bus_date='" & ToDBDate(BusStartupDate) & "' AND " _
                            & " p.sell_station_id='" & SellStationID & "' AND" _
                            & " s.seat_no=" _
                            & "(SELECT MIN(seat_no) FROM Env_bus_seat_lst WHERE " _
                            & " bus_id='" & BusID & "' AND " _
                            & " bus_date='" & ToDBDate(BusStartupDate) & "' AND (" _
                            & " status=" & ST_SeatCanSell & " OR status=" & ST_SeatBooked & " or status=" & ST_SeatProjectBooked & ") AND seat_type_id ='" & SellTicketInfo(i).szSeatTypeID & "') AND " _
                            & "s.bus_id=p.bus_id AND " _
                            & "s.bus_date=p.bus_date AND "
                            szsql = szsql & "s.seat_type_id=p.seat_type_id AND " _
                                        & "p.station_id='" & DesStationID & "' AND " _
                                        & "p.ticket_type='" & IIf(SellTicketInfo(i).nTicketType = TP_FreeTicket, TP_FullPrice, SellTicketInfo(i).nTicketType) & "'"
                        Else
                            szsql = "SELECT p.ticket_price_total AS TicketPrice,s.seat_no AS SeatNo,s.seat_type_id AS SeatType,p.ticket_price_total AS TicketPrice,s.status  FROM Env_bus_seat_lst s, work_env_bus_price_lst p WHERE " _
                            & " s.bus_id='" & BusID & "' AND " _
                            & " p.sell_station_id='" & SellStationID & "' AND " _
                            & " s.bus_date='" & ToDBDate(BusStartupDate) & "' AND " _
                            & "s.seat_no=(SELECT MIN(seat_no) FROM Env_bus_seat_lst WHERE bus_id='" & BusID & "' AND " _
                            & "bus_date='" & ToDBDate(BusStartupDate) & "' AND status=" & ST_SeatCanSell & " AND seat_type_id ='" & SellTicketInfo(i).szSeatTypeID & "') AND " _
                            & "s.bus_id=p.bus_id AND " _
                            & "s.bus_date=p.bus_date AND "
                            szsql = szsql & "s.seat_type_id=p.seat_type_id AND " _
                                        & "p.station_id='" & DesStationID & "' AND " _
                                        & "p.ticket_type='" & IIf(SellTicketInfo(i).nTicketType = TP_FreeTicket, TP_FullPrice, SellTicketInfo(i).nTicketType) & "'"
                        End If
                                              
                        Set rsTemp = poDb.Execute(szsql)
                        If rsTemp.RecordCount = 0 Then ShowError ERR_BusReachMaxTicketCount, "SellTicketClient.SelfSellTicketLocalTrans", "或" & SellTicketInfo(i).szSeatTypeName & "座位已售完。"
                        szSeatNo = FormatDbValue(rsTemp!SeatNo) '得到座位号
                        szSeatType = FormatDbValue(rsTemp!SeatType) '得到座位类型
                        sgTicketPrice = rsTemp!TicketPrice   '得到票价
                        nSeatStatus = rsTemp!Status
                        
                        If bNeedBook And Not bBookValid Then
                            bSeatBooked = IIf(rsTemp!Status = ST_SeatBooked, True, False)
'                            If rsTemp!Status = ST_SeatProjectBooked Then
'                                AssertHaveRight m_oActiveUser, RIGHT_RightOther1 '判断是否有计划预定权限
'                            End If
                        Else
                            bSeatBooked = False
                        End If
                        If bSeatBooked Then
                            szsql = UpdateSeatStatusStr2(BusID, BusStartupDate, szSeatNo, SellTicketInfo(i).szTicketNo, bBookValid)
                        Else
                            szsql = UpdateSeatStatusStr(BusID, BusStartupDate, szSeatNo, SellTicketInfo(i).szTicketNo)
                        End If
                        If rsTemp!Status = ST_SeatProjectBooked Then szsql = UpdateSeatStatusStr3(BusID, BusStartupDate, szSeatNo, SellTicketInfo(i).szTicketNo)
                        poDb.Execute szsql, lRowAffect
                        
                        If lRowAffect = 0 Then GoTo RetrySeatNo
                    Else
                        szSeatNo = SellTicketInfo(i).szSeatNo
                        
                         '取座位该票价和类型
                        szSqlSeat = "SELECT p.ticket_price_total AS TicketPrice, p.seat_type_id AS SeatType,s.status AS SeatStatus " _
                                        & " FROM work_env_bus_price_lst p,Env_bus_seat_lst s WHERE " _
                                        & " p.bus_id='" & BusID & "' AND p.station_id='" & DesStationID & "' AND " _
                                        & " p.bus_date='" & ToDBDate(BusStartupDate) & "' AND " _
                                        & " p.ticket_type=" & SellTicketInfo(i).nTicketType & " AND " _
                                        & " p.seat_type_id =(SELECT seat_type_id FROM Env_bus_seat_lst WHERE " _
                                        & " bus_id='" & BusID & "' AND bus_date='" & ToDBDate(BusStartupDate) & "' AND seat_no='" & szSeatNo & "') AND " _
                                        & " s.status=(SELECT status FROM Env_bus_seat_lst WHERE " _
                                        & " bus_id='" & BusID & "' AND bus_date='" & ToDBDate(BusStartupDate) & "' AND seat_no='" & szSeatNo & "')" _
                                        & " AND p.sell_station_id ='" & SellStationID & "'" _
                                        & " AND p.bus_id = s.bus_id AND p.bus_date = s.bus_date AND s.seat_no = '" & szSeatNo & "'" '
                                        
                        Set rsTemp = poDb.Execute(szSqlSeat)
                        
                        If rsTemp.RecordCount = 0 Then ShowError ERR_BusReachMaxTicketCount
                        sgTicketPrice = rsTemp!TicketPrice
                        szSeatType = FormatDbValue(rsTemp!SeatType)
                        nSeatStatus = rsTemp!seatstatus
                        
                        szsql = UpdateSeatStatusStr(BusID, BusStartupDate, SellTicketInfo(i).szSeatNo, SellTicketInfo(i).szTicketNo)
                        poDb.Execute szsql, lRowAffect
                        If lRowAffect = 0 Then
                            If bNeedBook Then
                                If nSeatStatus = ST_SeatProjectBooked Then
                                
                                    AssertHaveRight m_oActiveUser, RIGHT_RightOther1
                                    szsql = UpdateSeatStatusStr3(BusID, BusStartupDate, SellTicketInfo(i).szSeatNo, SellTicketInfo(i).szTicketNo)
                                    poDb.Execute szsql, lRowAffect
                                    If lRowAffect = 0 Then ShowError ERR_SeatNoCanNotSell, "SellTicketService.SelfSellTicket"
                                    bSeatBooked = False
                                Else
                                    szsql = UpdateSeatStatusStr2(BusID, BusStartupDate, SellTicketInfo(i).szSeatNo, SellTicketInfo(i).szTicketNo, bBookValid, SellTicketInfo(i).szReserved)
                                    poDb.Execute szsql, lRowAffect
                                    If lRowAffect = 0 Then ShowError ERR_SeatNoCanNotSell, "SellTicketClient.SelfSellTicketLocalTrans"
                                
                                    bSeatBooked = True
                                End If
                            Else
                                ShowError ERR_SeatNoCanNotSell, "SellTicketClient.SelfSellTicketLocalTrans"
                            End If
                        Else
                            bSeatBooked = False
                        End If
                        
                    End If
                    '如果座位预定
                    If bSeatBooked Then
                        szsql = "SELECT * FROM ticket_book_lst WHERE " _
                        & " bus_id='" & BusID & "' AND " _
                        & " bus_date='" & ToDBDate(BusStartupDate) & "' AND " _
                        & " seat_no='" & szSeatNo & "' AND " _
                        & " status=0"
                        Set rsTemp2 = poDb.Execute(szsql)
                        szBookStationID = Trim(rsTemp2!station_id)
                        Set rsTemp2 = Nothing

                        
                        'If nStationLimitCount > 0 Then
                            '更新站点的预定张数
                        szsql = "UPDATE Env_bus_station_lst SET " _
                        & " book_quantity=book_quantity-1 WHERE " _
                        & " bus_id='" & BusID & "' AND " _
                        & " bus_date='" & ToDBDate(BusStartupDate) & "' AND " _
                        & " station_id='" & szBookStationID & "' AND " _
                        & " sell_station_id='" & SellStationID & "' AND " _
                        & " book_quantity>0"
                        
                        poDb.Execute szsql
                        'End If
                        
                        If bBookValid Then
                            szsql = "UPDATE ticket_book_lst SET status=2 WHERE " _
                            & " bus_id='" & BusID & "' AND " _
                            & " bus_date='" & ToDBDate(BusStartupDate) & "' AND " _
                            & " seat_no='" & szSeatNo & "' AND " _
                            & " status=0 " '0状态为预定
                        Else
                            szsql = "UPDATE ticket_book_lst SET status=3 WHERE " _
                            & " bus_id='" & BusID & "' AND " _
                            & " bus_date='" & ToDBDate(BusStartupDate) & "' AND " _
                            & " seat_no='" & szSeatNo & "' AND " _
                            & " status=0 " '0状态为预定
                        End If
                        poDb.Execute szsql
                    End If
                    
                    '得到环境车次配载信息
                    Set rsTemp2 = poDb.Execute("SELECT * from work_env_bus_allot_lst WHERE bus_date='" _
                    & ToDBDate(BusStartupDate) & "' AND " _
                    & " bus_id='" & BusID & "' AND " _
                    & " sell_station_id='" & SellStationID & "'")
                    If rsTemp2.RecordCount = 0 Then ShowError ERR_BusAllotInfoNotExist
                    
                    '总共允许售票张数是否大于"已售张数"+1
                    
                    If rsTemp2!can_sale_quantity >= rsTemp2!sell_quantity + 1 Or rsTemp2!can_sale_quantity = -1 Then
                        poDb.Execute ("UPDATE Env_bus_allot_lst SET sell_quantity=sell_quantity + 1 WHERE bus_date='" _
                            & ToDBDate(BusStartupDate) & "' AND " _
                            & " bus_id='" & BusID & "' AND " _
                            & " sell_station_id='" & SellStationID & "'")
                    Else
                        ShowError ERR_SellQuantityISFull '超过允许售票张数
                    End If
                    
                    If bSellStationCanSellEachOther Then
                        '得到环境车次售票点信息
                        Set rsTemp2 = poDb.Execute("SELECT * from work_env_bus_sell_station_lst WHERE bus_date='" _
                        & ToDBDate(BusStartupDate) & "' AND " _
                        & " bus_id='" & BusID & "' AND " _
                        & " sell_station_id='" & m_oActiveUser.SellStationID & "'")
                        If rsTemp2.RecordCount = 0 Then ShowError ERR_BusSellStationInfoNotExist
                        
                        '总共允许售票张数是否大于"已售张数"+1
                        
                        If rsTemp2!can_sale_quantity >= rsTemp2!sell_quantity + 1 Or rsTemp2!can_sale_quantity = -1 Then
                            poDb.Execute ("UPDATE env_bus_sell_station_lst SET sell_quantity=sell_quantity + 1 WHERE bus_date='" _
                                & ToDBDate(BusStartupDate) & "' AND " _
                                & " bus_id='" & BusID & "' AND " _
                                & " sell_station_id='" & m_oActiveUser.SellStationID & "'")
                        Else
                            ShowError ERR_SellQuantityISFull '超过允许售票张数
                        End If
                    End If

                    Set rsTemp2 = Nothing
                    If nSeatStatus = ST_SeatProjectBooked Or nSeatStatus = ST_SeatBooked Then
                        '如果为售预定票,则不减剩余座位数
                    Else
                        '判断到站的限售张数够用否
                        If nStationRemainCount >= 0 Then
                            If nStationRemainCount = 0 Then
'                                ShowError ERR_BusReachMaxStationTicketCount, "SellTicketClient.SelfSellTicketLocalTrans"
                            Else
                                '更新可售张数
                                szsql = "UPDATE Env_bus_station_lst SET " _
                                & " sale_ticket_quantity=sale_ticket_quantity-1 WHERE " _
                                & " bus_id='" & BusID & "' AND " _
                                & " bus_date='" & ToDBDate(BusStartupDate) & "' AND " _
                                & " sale_ticket_quantity>book_quantity AND " _
                                & " station_id='" & DesStationID & "'"
'                                & " sell_station_id='" & SellStationID & "' AND "
                                poDb.Execute szsql, lRowAffect
                                
                                If lRowAffect = 0 Then ShowError ERR_BusReachMaxStationTicketCount, "SellTicketClient.SelfSellTicketLocalTrans"
                            End If
                            
                        End If
                        szsql = "UPDATE Env_bus_info SET sale_seat_quantity=sale_seat_quantity-1 , "
                        Select Case szSeatType
                            Case cszSeatType
                                szsql = szsql & "seat_remain=seat_remain - 1 "
                            Case cszBedType
                                szsql = szsql & "bed_remain=bed_remain - 1 "
                            Case cszAdditionalType
                                szsql = szsql & "additional_remain=additional_remain - 1 "
                            Case cszOtherType1
                                szsql = szsql & "other_remain_1 =other_remain_1 - 1"
                            Case cszOtherType2
                                szsql = szsql & "other_remain_2=other_remain_2 - 1"
                        End Select
                        szsql = szsql & " WHERE bus_id='" & BusID & "' AND bus_date='" & ToDBDate(BusStartupDate) & "'"
                        
                        poDb.Execute szsql
                    End If
                    
                Else
                    szsql = "UPDATE Env_bus_info SET sale_stand_seat_quantity=sale_stand_seat_quantity-1 " _
                    & " WHERE bus_id='" & BusID & "' AND " _
                    & "bus_date='" & ToDBDate(BusStartupDate) & "' AND " _
                    & " sale_stand_seat_quantity>0 "
                    
                    poDb.Execute szsql, lRowAffect
                    If lRowAffect = 0 Then ShowError ERR_NoStandTicket, "SellTicketClient.SelfSellTicketLocalTrans"
                    szSeatNo = cszStandTicketSeatNo
                End If
            Else
                szSeatType = cszSeatType
            End If
            '如果票价为0且不为免票,则不允许出售
            If sgTicketPrice = 0 And SellTicketInfo(i).nTicketType <> TP_FreeTicket Then
                RaiseError ERR_CannotSellZeroPrice
            End If
            
            
            '写售票明细表
            szsql = "INSERT INTO Ticket_sell_lst(" _
            & " sell_station_id," _
            & " ticket_id," _
            & " bus_date," _
            & " bus_id," _
            & " des_station_id," _
            & " status," _
            & " ticket_type," _
            & " user_id," _
            & " seat_no," _
            & " ticket_price," _
            & " sale_ticket_type," _
            & " seat_type_id ," _
            & " operation_time, " _
            & " insurance ) "
            
            szsql = szsql & " VALUES('" _
            & SellStationID & "','" _
            & SellTicketInfo(i).szTicketNo & "','" _
            & ToDBDate(BusStartupDate) & "','" _
            & BusID & "','" _
            & DesStationID & "'," _
            & nSellTicketStatus & "," _
            & SellTicketInfo(i).nTicketType & ",'" _
            & m_oActiveUser.UserID & "','" _
            & szSeatNo & "'," _
            & sgTicketPrice & "," _
            & nSellTicketType & ",'" _
            & szSeatType & "','" _
            & ToDBDateTime(dtNow) & "', " _
            & TransFieldValueToString(pnInsurance) & ")"
            
            On Error GoTo MayBeTicketIDDup
            poDb.Execute szsql
            
            '写详细票价信息
'            If atiTicketPrice(SellTicketInfo(i).nTicketType).sgBase = 0 Then
                szsql = GetTicketPriceDetailStr(BusID, BusStartupDate, SellStationID, DesStationID, szSeatType, IIf(SellTicketInfo(i).nTicketType = TP_FreeTicket, TP_FullPrice, SellTicketInfo(i).nTicketType))
                Set rsTemp = poDb.Execute(szsql)
                LetPriceDetailFromRsToType rsTemp, atiTicketPrice(SellTicketInfo(i).nTicketType), 0
'            End If
            szsql = GetInsertPriceDetailStr(atiTicketPrice(SellTicketInfo(i).nTicketType), SellTicketInfo(i).szTicketNo)
            poDb.Execute szsql
            
            siSellInfo.aszSeatType(i) = szSeatType
            siSellInfo.aszTicketType(i) = SellTicketInfo(i).nTicketType
            siSellInfo.asgTicketPrice(i) = sgTicketPrice
            siSellInfo.aszSeatNO(i) = szSeatNo
        Next
        szsql = "SELECT station_name FROM station_info WHERE station_id='" & DesStationID & "'"
        Set rsTemp = poDb.Execute(szsql)
        
        siSellInfo.szDesStationName = FormatDbValue(rsTemp!station_name)
        
    Else
        ShowError ERR_NoTicketInfoSpecified
    End If
    
    SelfSellTicketLocalTrans = siSellInfo
    
    
    Dim nIndex As Integer
    Dim szStartStationID As String, szStartStationName As String
    Dim szUnitID As String
    Dim szDesStationName As String
    szDesStationName = siSellInfo.szDesStationName
    
    nBuyTicketCount = ArrayLength(SellTicketInfo)
    '此处需改为从参数上车站中得到的站点及站点名称
    szsql = " SELECT s.* , t.station_name FROM sell_station_info s, station_info t WHERE s.station_id = t.station_id "
    Set rsTemp = poDb.Execute(szsql)
    szStartStationID = FormatDbValue(rsTemp!station_id)
    szStartStationName = FormatDbValue(rsTemp!station_name)
    '不需要从系统参数中得到,改为从传入的参数中得到
    
'    szStartStationID = m_oCurrentSellService.StationID
'    szStartStationName = m_oCurrentSellService.StationName
    
    szUnitID = m_oCurrentSellService.SelfUser.ServerUnit
    
    For nIndex = 1 To nBuyTicketCount
        '写代理售票明细表
        szsql = "INSERT INTO Agent_ticket_sale_lst(" _
        & " sell_station_id," _
        & " ticket_id," _
        & " bus_date," _
        & " bus_id," _
        & "service_unit_id," _
        & "start_station_id," _
        & "start_station_name," _
        & " des_station_id," _
        & " des_station_name," _
        & " ticket_type," _
        & " user_id," _
        & "seat_no," _
        & " ticket_price," _
        & "status," _
        & " sale_ticket_type, " _
        & " seat_type_id," _
        & " operation_time, " _
        & " insurance )  " _
        
        
        szsql = szsql & " VALUES('" _
        & SellStationID & "','" _
        & SellTicketInfo(nIndex).szTicketNo & "','" _
        & ToDBDate(BusStartupDate) & "','" _
        & BusID & "','" _
        & szUnitID & "','" _
        & szStartStationID & "','" _
        & szStartStationName & "','" _
        & DesStationID & "','" _
        & szDesStationName & "'," _
        & SellTicketInfo(nIndex).nTicketType & ",'" _
        & m_oActiveUser.UserID & "','" _
        & siSellInfo.aszSeatNO(nIndex) & "'," _
        & siSellInfo.asgTicketPrice(nIndex) & "," _
        & nSellTicketStatus & "," _
        & nSellTicketType & ",'" _
        & siSellInfo.aszSeatType(nIndex) & "','" _
        & ToDBDateTime(dtNow) & "'," _
        & TransFieldValueToString(pnInsurance) & " ) "
        
        poDb.Execute szsql
    Next nIndex
    Set rsTemp = Nothing
    Set rsRange = Nothing
    Set rsTemp2 = Nothing
    Set oParam = Nothing
Exit Function
MayBeTicketIDDup:
    
    If poDb.HaveThisNativeErr(ERR_DBAddDuplicate) Then

        ShowError ERR_TicketNoDup
    Else
        
        err.Raise err.Number
    End If
     Set rsTemp = Nothing
     Set rsRange = Nothing
     Set rsTemp2 = Nothing
     Set oParam = Nothing
End Function
Private Function UpdateSeatStatusStr(pszBusID As String, pszDate As Date, pszSeatNo As String, pszTicketID As String) As String
    Dim szsql As String
    
    szsql = "UPDATE Env_bus_seat_lst SET status=" & ST_SeatSold & "," _
    & " ticket_no='" & pszTicketID & "' " _
    & " WHERE bus_id='" & pszBusID & "' AND " _
    & " bus_date='" & ToDBDate(pszDate) & "' AND " _
    & " seat_no='" & pszSeatNo & "' AND " _
    & " status=" & ST_SeatCanSell
    UpdateSeatStatusStr = szsql
End Function

Private Function UpdateSeatStatusStr3(pszBusID As String, pszBusDate As Date, pszSeatNo As String, pszTicketID As String) As String
    Dim szsql As String
    szsql = "UPDATE Env_bus_seat_lst SET status=" & ST_SeatSold & "," _
            & " ticket_no='" & pszTicketID & "' " _
            & " WHERE bus_id='" & pszBusID & "' AND " _
            & " bus_date='" & ToDBDate(pszBusDate) & "'AND" _
            & " seat_no='" & pszSeatNo & "'AND " _
            & " status=" & ST_SeatProjectBooked & " "
    UpdateSeatStatusStr3 = szsql
End Function

Private Function GetTicketPriceDetailStr(pszBusID As String, pdtDate As Date, pszSellStationID As String, pszStaionID As String, pszSeatType As String, pnTicketType As ETicketType) As String
    Dim szsql As String
    Dim szTemp As String
    Dim oParam As New SystemParam
    oParam.Init m_oActiveUser
    
    szsql = "SELECT * FROM Work_env_bus_price_lst WHERE " _
    & " bus_id='" & pszBusID & "' AND " _
    & " bus_date='" & ToDBDate(pdtDate) & "' AND " _
    & " station_id='" & pszStaionID & "' AND " _
    & "seat_type_id='" & pszSeatType & "' AND" _
    & " sell_station_id='" & pszSellStationID & "' AND " _
    & " ticket_type=" & pnTicketType

    GetTicketPriceDetailStr = szsql
End Function


Private Function GetInsertPriceDetailStr(ptiPriceDetailInfo As TRETicketPrice, pszTicketID As String) As String
    Dim szsql As String
    Dim i As Integer
    
    szsql = " INSERT INTO ticket_price_lst(" _
    & "ticket_id," _
    & " ticket_price_total," _
    & " base_price," _
    & " price_item_1," _
    & " price_item_2," _
    & " price_item_3," _
    & " price_item_4," _
    & " price_item_5," _
    & " price_item_6," _
    & " price_item_7," _
    & " price_item_8," _
    & " price_item_9," _
    & " price_item_10," _
    & " price_item_11," _
    & " price_item_12," _
    & " price_item_13," _
    & " price_item_14," _
    & " price_item_15) VALUES('" _
    & pszTicketID & "'," _
    & ptiPriceDetailInfo.sgTotal & "," _
    & ptiPriceDetailInfo.sgBase
    For i = 1 To 15
        szsql = szsql & "," & ptiPriceDetailInfo.asgPrice(i)
    Next
    szsql = szsql & ")"
    
    GetInsertPriceDetailStr = szsql
End Function


Private Sub LetPriceDetailFromRsToType(prsIn As Recordset, ptiInOut As TRETicketPrice, sgDiscountPrice As Double)
    Dim i As Integer
    ptiInOut.sgBase = FormatDbValue(prsIn!base_carriage) + sgDiscountPrice
    ptiInOut.sgTotal = FormatDbValue(prsIn!ticket_price_total) + sgDiscountPrice
    For i = 1 To 15
        ptiInOut.asgPrice(i) = prsIn("price_item_" & i)
    Next
    
End Sub

Private Function UpdateSeatStatusStr2(pszBusID As String, pszDate As Date, pszSeatNo As String, pszTicketID As String, pbBookValid As Boolean, Optional pszBookNumber As String = "") As String
    Dim szsql As String
    szsql = "UPDATE Env_bus_seat_lst SET status=" & ST_SeatSold & "," _
    & " ticket_no='" & pszTicketID & "' FROM ticket_book_lst " _
    & " WHERE Env_bus_seat_lst.bus_id='" & pszBusID & "' AND " _
    & " Env_bus_seat_lst.bus_date='" & ToDBDate(pszDate) & "' AND " _
    & " Env_bus_seat_lst.seat_no='" & pszSeatNo & "' AND " _
    & " ticket_book_lst.status=0 AND " _
    & " Env_bus_seat_lst.bus_id=ticket_book_lst.bus_id AND " _
    & " Env_bus_seat_lst.bus_date=ticket_book_lst.bus_date AND " _
    & " Env_bus_seat_lst.seat_no=ticket_book_lst.seat_no AND " _
    & " Env_bus_seat_lst.status=" & ST_SeatBooked
    If pbBookValid Then
'        szSql = szSql & " AND book_number='" & pszBookNumber & "' "
    End If
    UpdateSeatStatusStr2 = szsql
End Function
'///////////////////////////////////
Public Function DiscountIsValid() As Boolean
    AssertActiveUserValid m_oActiveUser, ERR_SellTicketService
    #If IN_DEBUG = 0 Then
    Dim szsql As String
    Dim rsTemp As Recordset
    Dim oDb As New RTConnection
    Dim szRight As String
    
    szRight = GetRightID(RIGHT_SellDiscountTicket)
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    '查询用户方法表中用户是否有该权限
    szsql = "SELECT function_id FROM user_function_lst WHERE " _
    & " user_id='" & m_oActiveUser.UserID & "' AND " _
    & " function_id='" & szRight & "'"
    Set rsTemp = oDb.Execute(szsql)
    If rsTemp.RecordCount = 0 Then
        '如果用户没有权限则查询用户组是否有权限
        szsql = "SELECT function_id FROM usergroup_function_lst tbu,Group_user_info tbg WHERE tbu.usergroup_id=" _
        & " tbg.usergroup_id AND user_id='" & m_oActiveUser.UserID & "' AND function_id='" & szRight & "'"
        Set rsTemp = oDb.Execute(szsql)
        If rsTemp.RecordCount = 0 Then
            DiscountIsValid = True
        Else
            DiscountIsValid = False
        End If
    Else
        DiscountIsValid = False
    End If
#End If
End Function
'////////////////////////////////////
Public Function OrderSeatIsValid() As Boolean
    AssertActiveUserValid m_oActiveUser, ERR_SellTicketService
    #If IN_DEBUG = 0 Then
    Dim szsql As String
    Dim rsTemp As Recordset
    Dim oDb As New RTConnection
    Dim szRight As String
    
    szRight = GetRightID(RIGHT_SellOrderSeat)
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    '查询用户方法表中用户是否有该权限
    szsql = "SELECT function_id FROM user_function_lst WHERE " _
    & " user_id='" & m_oActiveUser.UserID & "' AND " _
    & " function_id='" & szRight & "'"
    Set rsTemp = oDb.Execute(szsql)
    If rsTemp.RecordCount = 0 Then
        '如果用户没有权限则查询用户组是否有权限
        szsql = "SELECT function_id FROM usergroup_function_lst tbu,Group_user_info tbg WHERE tbu.usergroup_id=" _
        & " tbg.usergroup_id AND user_id='" & m_oActiveUser.UserID & "' AND function_id='" & szRight & "'"
        Set rsTemp = oDb.Execute(szsql)
        If rsTemp.RecordCount = 0 Then
            OrderSeatIsValid = True
        Else
            OrderSeatIsValid = False
        End If
    Else
        OrderSeatIsValid = False
    End If
#End If
End Function
'////////////////////////////////////
Public Function SelectReturnIsValid() As Boolean
    AssertActiveUserValid m_oActiveUser, ERR_SellTicketService
    #If IN_DEBUG = 0 Then
    Dim szsql As String
    Dim rsTemp As Recordset
    Dim oDb As New RTConnection
    Dim szRight As String
    
    szRight = GetRightID(RIGHT_SelectReturnRatio)
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
    '查询用户方法表中用户是否有该权限
    szsql = "SELECT function_id FROM user_function_lst WHERE " _
    & " user_id='" & m_oActiveUser.UserID & "' AND " _
    & " function_id='" & szRight & "'"
    Set rsTemp = oDb.Execute(szsql)
    If rsTemp.RecordCount = 0 Then
        '如果用户没有权限则查询用户组是否有权限
        szsql = "SELECT function_id FROM usergroup_function_lst tbu,Group_user_info tbg WHERE tbu.usergroup_id=" _
        & " tbg.usergroup_id AND user_id='" & m_oActiveUser.UserID & "' AND function_id='" & szRight & "'"
        Set rsTemp = oDb.Execute(szsql)
        If rsTemp.RecordCount = 0 Then
            SelectReturnIsValid = True
        Else
            SelectReturnIsValid = False
        End If
    Else
        SelectReturnIsValid = False
    End If
#End If
End Function

'////////////////////////////////////
'得到滚动车次发车时间
Public Function GetScrollBusStartTime() As Date
    Dim dDate As Date
    Dim dbLatestTime As Double
    Dim oParam As New SystemParam
    oParam.Init m_oActiveUser
    dDate = ToDBTime(SelfNowDateTime())
    dbLatestTime = oParam.ScrollBusReturnTime
    dDate = DateAdd("n", dbLatestTime * 60, dDate)
    GetScrollBusStartTime = Format(dDate, "HH:MM")
End Function



'回收座位或站票的资源信息
Private Sub RecycleSeat(pszTicketID As String, poDb As RTConnection)
    Dim szsql As String
    Dim rsTemp As Recordset
    Dim szBusID As String
    Dim dtDate As Date
    Dim szDesStationID As String
    Dim szSeatNo As String
    Dim szSeatTypeID As String
    Dim szSellStationID As String
    
    szsql = "SELECT * FROM Ticket_sell_lst WHERE ticket_id='" & pszTicketID & "'"
    Set rsTemp = poDb.Execute(szsql)
    
    szBusID = FormatDbValue(rsTemp!bus_id)
    dtDate = FormatDbValue(rsTemp!bus_date)
    szDesStationID = FormatDbValue(rsTemp!des_station_id)
    szSeatNo = FormatDbValue(rsTemp!seat_no)
    szSellStationID = FormatDbValue(rsTemp!sell_station_id)
    If szSeatNo = cszStandTicketSeatNo Then '是站票
        '将该环境车次的可售站票数加一
        szsql = "UPDATE Env_bus_info SET " _
        & " sale_stand_seat_quantity=sale_stand_seat_quantity+1 WHERE " _
        & " bus_id='" & szBusID & "' AND " _
        & " bus_date='" & ToDBDate(dtDate) & "'"
        poDb.Execute szsql
    Else '不是站票
        szSeatTypeID = FormatDbValue(rsTemp!seat_type_id)
        
        szSeatTypeID = FormatDbValue(rsTemp!seat_type_id)
        
        '将该环境车次的可售座位数减一和可售座位类型余留减一
        szsql = "UPDATE Env_bus_info SET sale_seat_quantity=sale_seat_quantity+1 , "
        Select Case szSeatTypeID
            Case cszSeatType
                szsql = szsql & " seat_remain=seat_remain + 1  "
            Case cszBedType
                szsql = szsql & " bed_remain=bed_remain + 1 "
            Case cszAdditionalType
                szsql = szsql & " additional_remain=additional_remain + 1 "
            Case cszOtherType1
                szsql = szsql & " other_remain_1 =other_remain_1 + 1 "
            Case cszOtherType2
                szsql = szsql & " other_remain_2 =other_remain_2 + 1 "
        End Select
        szsql = szsql & " WHERE bus_id='" & szBusID & "' AND bus_date='" & ToDBDate(dtDate) & "'"
        poDb.Execute szsql
        
        '改变相应座位号的状态
        szsql = "UPDATE Env_bus_seat_lst SET " _
        & " status=" & ST_SeatCanSell & "," _
        & " ticket_no='' WHERE " _
        & " bus_id='" & szBusID & "' AND " _
        & " bus_date='" & ToDBDate(dtDate) & "' AND " _
        & " seat_no='" & szSeatNo & "'"
        
        poDb.Execute szsql
    End If
    
    '更改环境车次站点表的可售票数
    szsql = "UPDATE Env_bus_station_lst SET " _
    & " sale_ticket_quantity=sale_ticket_quantity+1 WHERE " _
    & " bus_id='" & szBusID & "' AND " _
    & " bus_date='" & ToDBDate(dtDate) & "' AND " _
    & " station_id='" & szDesStationID & "' AND " _
    & " sell_station_id='" & szSellStationID & "' AND " _
    & " max_sale_quantity>0 "
    
    poDb.Execute szsql
    
    '如果是携童票，减一
'    If rsTemp!if_take_children Then
'        szSql = "UPDATE env_bus_info SET take_children_quantity=take_children_quantity-1 WHERE bus_id='" & rsTemp!bus_id & "'" _
'               & " AND bus_date='" & rsTemp!bus_date & "'"
'        poDb.Execute szSql
'    End If
    
End Sub

'释放座位
Private Sub ResumeSeat(pszTicket As String, poDb As RTConnection)
    Dim szsql As String
    Dim rsTemp As Recordset
    Dim rsTemp1 As Recordset
    Dim szBusID As String
    Dim dtDate As Date
    Dim szDesStationID As String
    Dim szSeatNo As String
    Dim szSeatTypeID As String
    Dim szSellStationID As String
    
    On Error GoTo Here
    szsql = "SELECT * FROM Ticket_sell_lst WHERE ticket_id='" & pszTicket & "'"
    Set rsTemp = poDb.Execute(szsql)
    
    szBusID = FormatDbValue(rsTemp!bus_id)
    dtDate = FormatDbValue(rsTemp!bus_date)
    szDesStationID = FormatDbValue(rsTemp!des_station_id)
    szSeatNo = FormatDbValue(rsTemp!seat_no)
    szSellStationID = FormatDbValue(rsTemp!sell_station_id)
    
     '////////////////////////////
    '判断是否还有可售票数
    szsql = "SELECT * FROM Env_bus_info WHERE " _
    & " bus_id='" & szBusID & "'AND " _
    & " bus_date='" & ToDBDate(dtDate) & "'"
    Set rsTemp1 = poDb.Execute(szsql)
   
    
    If szSeatNo = cszStandTicketSeatNo Then '是站票
       If rsTemp1!sale_stand_seat_quantity >= 1 Then
            '将该环境车次的可售站票数减一
            szsql = "UPDATE Env_bus_info SET " _
            & " sale_stand_seat_quantity=sale_stand_seat_quantity-1 WHERE " _
            & " bus_id='" & szBusID & "' AND " _
            & " bus_date='" & ToDBDate(dtDate) & "'"
            poDb.Execute szsql
       Else
            GoTo Here
       End If
    Else '不是站票
        
        
        If rsTemp1!sale_seat_quantity >= 1 Then
        
            szSeatTypeID = FormatDbValue(rsTemp!seat_type_id)
        
            '将该环境车次的可售座位数减一和可售座位类型余留减一
            szsql = "UPDATE Env_bus_info SET sale_seat_quantity=sale_seat_quantity-1 , "
            Select Case szSeatTypeID
                Case cszSeatType
                    szsql = szsql & " seat_remain=seat_remain - 1  "
                Case cszBedType
                    szsql = szsql & " bed_remain=bed_remain - 1 "
                Case cszAdditionalType
                    szsql = szsql & " additional_remain=additional_remain - 1 "
                Case cszOtherType1
                    szsql = szsql & " other_remain_1 =other_remain_1 - 1 "
                Case cszOtherType2
                    szsql = szsql & " other_remain_2 =other_remain_2 - 1 "
            End Select
            szsql = szsql & " WHERE bus_id='" & szBusID & "' AND bus_date='" & ToDBDate(dtDate) & "'"

            poDb.Execute szsql
            
            '改变相应座位号的状态
            szsql = "UPDATE Env_bus_seat_lst SET " _
            & " status=" & ST_SeatSold & "," _
            & " ticket_no='" & pszTicket & "' WHERE " _
            & " bus_id='" & szBusID & "' AND " _
            & " bus_date='" & ToDBDate(dtDate) & "' AND " _
            & " seat_no='" & szSeatNo & "'"
            poDb.Execute szsql
        Else
            GoTo Here
        End If
    End If
    
    '更改环境车次站点表的可售票数
    szsql = "UPDATE Env_bus_station_lst SET " _
    & " sale_ticket_quantity=sale_ticket_quantity-1 WHERE " _
    & " bus_id='" & szBusID & "' AND " _
    & " bus_date='" & ToDBDate(dtDate) & "' AND " _
    & " station_id='" & szDesStationID & "' AND " _
    & " sell_station_id='" & szSellStationID & "' AND " _
    & " max_sale_quantity>0 "
    
    poDb.Execute szsql
    Set rsTemp = Nothing
    Set rsTemp1 = Nothing
    On Error GoTo 0
    Exit Sub
Here:
    Set rsTemp = Nothing
    Set rsTemp1 = Nothing
'    MsgBox "已无可售票数" & vbCrLf & "取消退票或取消废票可能产生超载", vbInformation, "售票台:"
End Sub







'取热门站点
Public Function HotStation() As Recordset
    Set HotStation = m_oCurrentSellService.HotStation()
End Function




Public Sub LocalSelfSellTrans(pszTicket() As TSellTicketParam, pszBankResult As String, pszCardID As String) '写明细

    Dim szsql As String
    Dim oDb As New RTConnection
    Dim lRowAffect As Long
    Dim SeatCount As Integer, i As Integer
    Dim szTemp As String
    Dim rsTemp As Recordset
    
    Dim chapmanID As String * 15
    Dim port As String * 8
    Dim BankID As String * 2
    Dim CenterScrollID As String * 12
    Dim BusyDate As String * 4
    Dim BusyTime As String * 6
    Dim CardAbleDate As String * 4
    Dim BatID As String * 6
    Dim ChapmanScrollID As String * 6
    Dim AuthorID As String * 6
    Dim pay As String * 12
    Dim cardID As String * 19
    
    chapmanID = Mid(pszBankResult, 7, 15)
    port = Mid(pszBankResult, 22, 8)
    BankID = Mid(pszBankResult, 30, 2)
    cardID = Mid(pszBankResult, 32, 19)
    CenterScrollID = Mid(pszBankResult, 51, 12)
    BusyDate = Mid(pszBankResult, 63, 4)
    BusyTime = Mid(pszBankResult, 67, 6)
    CardAbleDate = Mid(pszBankResult, 73, 4)
    BatID = Mid(pszBankResult, 77, 6)
    ChapmanScrollID = Mid(pszBankResult, 83, 6)
    AuthorID = Mid(pszBankResult, 89, 6)
    pay = Mid(pszBankResult, 95, 12)
    

    
    oDb.ConnectionString = GetConnectionStr("")
    oDb.BeginTrans
    On Error GoTo ErrorHandle '写自助售票表
       szsql = "INSERT ticket_self_sell_info(bank_card_id," _
        & "chapman_id," _
        & "port," _
        & "bank_id," _
        & "center_scroll_id," _
        & "busy_date," _
        & "busy_time," _
        & "card_able_date," _
        & "bat_id, " _
        & "chapman_scroll_id," _
        & "author_id," _
        & "pay_count) " _
        & " VALUES('" & cardID _
        & "','" & chapmanID _
        & "','" & port _
        & "','" & BankID _
        & "','" & CenterScrollID
        
        szsql = szsql _
        & "','" & BusyDate _
        & "','" & BusyTime _
        & "','" & CardAbleDate _
        & "','" & BatID _
        & "','" & ChapmanScrollID _
        & "','" & AuthorID _
        & "','" & pay & "')"
        oDb.Execute szsql
        

    For i = 1 To ArrayLength(pszTicket(1).BuyTicketInfo)
       szsql = "INSERT ticket_self_sell_lst(ticket_id," _
        & "bank_card_id," _
        & "busy_date," _
        & "center_scroll_id," _
        & "busy_time) " _
        & " VALUES('" & pszTicket(1).BuyTicketInfo(i).szTicketNo _
        & "','" & cardID _
        & "','" & BusyDate _
        & "','" & CenterScrollID _
        & "','" & BusyTime & "')"
        oDb.Execute szsql
    Next i
    oDb.CommitTrans



    Exit Sub
ErrorHandle:
    oDb.RollbackTrans
    err.Raise err.Number
    Exit Sub

End Sub



'自助售票统计
Public Function StatBySelfSellCount(pdtStrDate As Date, pdtEndDate As Date) As Recordset
On Error GoTo ErrHandle
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim szsql As String
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    
    szsql = "SELECT bank_card_id,cast(pay_count as money )/100 pay_count,center_scroll_id,chapman_scroll_id  from ticket_self_sell_info " _
            & "WHERE  busy_date>='" & Format(pdtStrDate, "mmdd") & "'" _
            & " AND  busy_date<'" & Format(pdtEndDate, "mmdd") & "'"
            
    
    Set rsTemp = oDb.Execute(szsql)
    Set StatBySelfSellCount = rsTemp
    Set oDb = Nothing
    Exit Function
ErrHandle:
    err.Raise err.Number
End Function


'自助售票明细
Public Function StatBySelfSellLst(pdtStrDate As Date, pdtEndDate As Date) As Recordset
On Error GoTo ErrHandle
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim szsql As String
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    
    szsql = "SELECT a.bank_card_id,b.ticket_id,a.center_scroll_id,a.chapman_scroll_id,t.ticket_price  from ticket_self_sell_info a,ticket_self_sell_lst b,ticket_sell_lst t" _
            & " WHERE  a.busy_date>='" & Format(pdtStrDate, "mmdd") & "'" _
            & " AND  a.busy_date<'" & Format(pdtEndDate, "mmdd") & "'" _
            & " AND t.ticket_id=b.ticket_id and a.bank_card_id=b.bank_card_id and a.center_scroll_id=b.center_scroll_id and a.busy_date=b.busy_date and a.busy_time=b.busy_time order by a.bank_card_id,b.ticket_id "
    
    Set rsTemp = oDb.Execute(szsql)
    Set StatBySelfSellLst = rsTemp
    Set oDb = Nothing
    Exit Function
ErrHandle:
    err.Raise err.Number
End Function

Public Function BankInfoPrint(pszTicket As String) As Recordset '银行信息打印
On Error GoTo ErrHandle
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim szsql As String
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    
    szsql = "SELECT a.bank_card_id,b.ticket_id,a.center_scroll_id,a.chapman_id,a.pay_count,a.port  from ticket_self_sell_info a,ticket_self_sell_lst b" _
            & " WHERE  b.ticket_id='" & pszTicket & "'" _
            & " AND a.busy_date=b.busy_date and a.busy_time=b.busy_time  and a.bank_card_id=b.bank_card_id and a.center_scroll_id=b.center_scroll_id order by a.bank_card_id,b.ticket_id "
    
    Set rsTemp = oDb.Execute(szsql)
    Set BankInfoPrint = rsTemp
    Set oDb = Nothing
    Exit Function
ErrHandle:
    err.Raise err.Number
End Function

' *******************************************************************
' *   Member Name: InternetSellTicket                               *
' *   Brief Description: 取网上售票                                   *
' *   Engineer: xxd                                                 *
' *   Date Generated: 2006-11-20          *
' *   Last Revision Date: 最后一次修改日期(日期格式：YYYY/MM/DD)    *
' *******************************************************************


Public Function InterNetValiDate(ByVal pszGetTicketID As String, ByVal pszValiDate As String, Optional pszCardID As String, Optional pdtBusDate As Date) As Recordset
     If Trim(m_oCurrentSellService.SelfUser.ServerUnit) <> Trim(m_oCurrentSellService.SelfUser.UserUnitID) Then
        '如果用户单位和服务单位不一样则是售远程票否则售的是本地票
       Set InterNetValiDate = m_oCurrentSellService.InterNetValiDate(pszGetTicketID, pszValiDate, pszCardID, pdtBusDate)
    Else
       Set InterNetValiDate = InterNetValiDateTrans(pszGetTicketID, pszValiDate, pszCardID, pdtBusDate)
    End If
    
    
    
End Function
'验证网上售票
Public Function InterNetValiDateTrans(Optional pszGetTicketID As String, Optional pszValiDate As String, Optional pszCardID As String, Optional pdtBusDate As Date) As Recordset

    Dim szsql As String
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim i As Integer
On Error GoTo Error_Handle
    oDb.ConnectionString = GetConnectionStr("")

    szsql = "SELECT * FROM internet_sell_ticket_info WHERE if_print=0"
            
    If Trim(pszGetTicketID) <> "" Then
        szsql = szsql & " AND getticket_id='" & pszGetTicketID & "' "
    End If
    
    If Trim(pszValiDate) <> "" Then
        szsql = szsql & " AND validate_id='" & pszValiDate & "'"
    End If
    
    If Trim(pszCardID) <> "" Then
        szsql = szsql & " AND card_id='" & pszCardID & "'"
    End If

    If pdtBusDate <> cszEmptyDate Then
        szsql = szsql & " AND bus_date='" & ToDBDate(pdtBusDate) & "'"
    End If

    Set rsTemp = oDb.Execute(szsql)
    Set InterNetValiDateTrans = rsTemp
    Exit Function
Error_Handle:
    err.Raise err.Number
End Function

'网上售票
Public Function GetInterNetTicket(ByVal rsTemp As Recordset, ByVal TicketID As String, ByRef aspSellTicketParam() As TSellTicketParam) As TSellTicketResult()
'
    Dim srTemp() As TSellTicketResult
'    Dim iLen As Integer
'    Dim iLen1 As Integer
'    Dim nCount As Integer
'    Dim aszDesStationName() As String
'
'    iLen = 0
'    iLen1 = 0
'    iLen = ArrayLength(BusID)
'    ReDim srTemp(1 To iLen)
    If Trim(m_oCurrentSellService.SelfUser.ServerUnit) <> Trim(m_oCurrentSellService.SelfUser.UserUnitID) Then
        '如果用户单位和服务单位不一样则是售远程取票否则售的是本地取票
        On Error GoTo ErrorHandle
        
        srTemp = m_oCurrentSellService.SellInetNetTicket(rsTemp, TicketID)
        
'        iLen1 = ArrayLength(srTemp)
'        ReDim aszDesStationName(1 To iLen1)
'        For nCount = 1 To iLen1
'            aszDesStationName(nCount) = srTemp(nCount).szDesStationName
'        Next nCount
'        On Error GoTo 0
'        On Error GoTo ErrorHandle2
'        SelfSellLocalInterNetTicket BusDate, BusID, SellStationID, DesStationID, aspSellTicketParam, anInsurance, oDb
'        On Error GoTo 0
    Else
        srTemp = SellInetNetTicket(rsTemp, TicketID, aspSellTicketParam)
    End If
    GetInterNetTicket = srTemp
    Exit Function
ErrorHandle:
    err.Raise err.Number

End Function
'取网上售票
Public Function SellInetNetTicket(ByVal rsTemp As Recordset, ByVal TicketID As String, ByRef atSellTKParam() As TSellTicketParam) As TSellTicketResult()
    
    Dim szsql As String
    Dim oDb As New RTConnection
    Dim rsTemp2 As Recordset
    Dim rsTemp3 As Recordset
    Dim i As Integer
    Dim tTSeatInfo(1 To 5) As TSeatInfoCount
    Dim nSaleSeatQuantity As Integer
    Dim szSqlSet As String
    Dim szSeatNoTemplog As String
    Dim j As Integer
    
    Dim adyDate_new() As Date
    Dim aszBusID_new()  As String
    Dim aszEndStationID_new() As String
'    Dim atSellTKParam() As TSellTicketParam
    Dim aszSeatNO() As String
    Dim atBuyTicketInfo() As TBuyTicketInfo
    Dim atSellTicketResultInfo() As TSellTicketResult

    Dim oParam As New SystemParam
    oParam.Init m_oActiveUser
    
    Dim szEndTicket As String
    Dim rs As Recordset
    Dim bChgTkNum As Boolean '判断是否为自助机的网上取票。
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)

    On Error GoTo Error_Handle
    bChgTkNum = False
    oDb.BeginTrans
    If Trim(TicketID) = "" Then
        bChgTkNum = True
        szsql = "select parameter_value from system_param_info  where parameter_name='StrTicketID'"
        Set rs = oDb.Execute(szsql)
        TicketID = Trim(rs!parameter_value)
    End If
    m_lTicketNo = ResolveTicketNo(TicketID, m_szTicketPrefix)


    For i = 1 To rsTemp.RecordCount
        szsql = "SELECT seat_no FROM Env_bus_seat_lst WHERE bus_id='" & Trim(rsTemp!bus_id) & "'" _
               & " AND  bus_date='" & ToDBDate(Trim(rsTemp!bus_date)) & "'" _
               & "AND seat_no ='" & Trim(rsTemp!seat_no) & "'" _
               & " AND (Status <> " & ST_SeatReserved & " AND Status <> " & ST_SeatProjectBooked & " AND Status <> " & ST_SeatReserved & " )"
        Set rsTemp2 = oDb.Execute(szsql)
        
        If rsTemp2.RecordCount > 0 Then
         szSeatNoTemplog = ""
         
          For j = 0 To rsTemp2.RecordCount - 1
           szSeatNoTemplog = szSeatNoTemplog & "[" & FormatDbValue(rsTemp2!seat_no) & "]"
           rsTemp2.MoveNext
          Next
    
        ShowError ERR_SeatNotBooked, "座位:" & szSeatNoTemplog
        
        End If
        szsql = "UPDATE Env_bus_seat_lst SET status=" & Str(ST_SeatCanSell) & " ," _
                & " remark='' " _
                & "    WHERE bus_id='" & Trim(rsTemp!bus_id) & "' AND  bus_date='" & ToDBDate(Trim(rsTemp!bus_date)) & "' " _
                & " AND seat_no  ='" & Trim(rsTemp!seat_no) & "'" _
                & " AND (Status = " & ST_SeatReserved & " or Status = " & ST_SeatProjectBooked & " or Status = " & ST_SeatReserved & " )"
        oDb.Execute szsql
    

        
        szsql = " select seat_type_id ,count(*)as CountSeat from Env_bus_seat_lst " _
                & " where  bus_id='" & Trim(rsTemp!bus_id) & "'and  status='" & ST_SeatCanSell & "' and bus_date='" & ToDBDate(Trim(rsTemp!bus_date)) & "' group by seat_type_id "
        Set rsTemp2 = oDb.Execute(szsql)
        
        For j = 1 To rsTemp2.RecordCount
            tTSeatInfo(CInt(rsTemp2!seat_type_id)).SeatCount = FormatDbValue(rsTemp2!CountSeat)
            rsTemp2.MoveNext
        Next
        szSqlSet = ""
        nSaleSeatQuantity = 0
        szSqlSet = szSqlSet & "seat_remain ='" & tTSeatInfo(1).SeatCount & "',"
        szSqlSet = szSqlSet & "bed_remain ='" & tTSeatInfo(2).SeatCount & "',"
        szSqlSet = szSqlSet & "additional_remain ='" & tTSeatInfo(3).SeatCount & "',"
        szSqlSet = szSqlSet & "other_remain_1 ='" & tTSeatInfo(4).SeatCount & "',"
        szSqlSet = szSqlSet & "other_remain_2 ='" & tTSeatInfo(5).SeatCount & "',"
        
        For j = 1 To 5
          nSaleSeatQuantity = nSaleSeatQuantity + tTSeatInfo(j).SeatCount
        Next
        
        szSqlSet = szSqlSet & "sale_seat_quantity='" & nSaleSeatQuantity & "'"
        szsql = "UPDATE Env_bus_info SET  " & szSqlSet _
                & "  WHERE bus_id='" & Trim(rsTemp!bus_id) & "' AND  bus_date='" & ToDBDate(Trim(rsTemp!bus_date)) & "'"
        oDb.Execute (szsql)
        ReDim Preserve aszSeatNO(1 To i + 1)
        aszSeatNO(i) = Trim(rsTemp!seat_no)
        rsTemp.MoveNext
    Next i
        Dim SellStationID(1 To 1) As String
        ReDim atBuyTicketInfo(1 To rsTemp.RecordCount)
        ReDim aszBusID_new(1 To 1)
        ReDim adyDate_new(1 To 1)
        ReDim aszEndStationID_new(1 To 1)
        ReDim atSellTKParam(1 To 1)
        Dim anInsurrance(1 To 1) As Integer
        Dim aszSellStationID() As String
        Dim aszUserID(1 To 1) As String
        Dim aszBankTime(1 To 1) As Date
 
        rsTemp.MoveFirst
        
        SellStationID(1) = Trim(rsTemp!sell_station_id)

        aszBusID_new(1) = Trim(rsTemp!bus_id)
        adyDate_new(1) = Trim(rsTemp!bus_date)
        aszEndStationID_new(1) = Trim(rsTemp!des_station_id)
        aszUserID(1) = Trim(rsTemp!bank_id)
        aszBankTime(1) = FormatDbValue(rsTemp!busy_date)
        ReDim aszResultSeatID(1 To rsTemp.RecordCount)
        Dim asgTicketPrices() As Single
        ReDim asgTicketPrices(1 To rsTemp.RecordCount)
    
        For i = 1 To rsTemp.RecordCount
        
            atBuyTicketInfo(i).nTicketType = Trim(rsTemp!ticket_type)
            atBuyTicketInfo(i).szReserved = ""
            If i > ArrayLength(aszSeatNO) Then
                '如果传过来的座位号清单的数目小于当前的数字的长数,则出错
                ShowError ERR_BookSeatNotExist
            Else
                atBuyTicketInfo(i).szSeatNo = aszSeatNO(i)
            End If
'            Dim szTmp As String
'            szTmp = TicketID + i
'            atBuyTicketInfo(i).szTicketNo = Left(TicketID, Len(TicketID) - Len(szTmp)) & TicketID + i - 1
''            atBuyTicketInfo(i).szTicketNo = GetTicketNo(i - 1 + nTemp + nTicketOffset)
            
            atBuyTicketInfo(i).szTicketNo = GetTicketNo(1)
            szEndTicket = atBuyTicketInfo(i).szTicketNo
            m_lTicketNo = m_lTicketNo + 1
            
            atBuyTicketInfo(i).szSeatTypeID = "01" '座位类型固定为01
            asgTicketPrices(i) = 0
'            atBuyTicketInfo(i).szInvoiceID = FormatDbValue(rsTemp!has_child)
            rsTemp.MoveNext
        Next i
        atSellTKParam(1).BuyTicketInfo = atBuyTicketInfo
        atSellTKParam(1).pasgSellTicketPrice = asgTicketPrices
        anInsurrance(1) = 0
        atSellTicketResultInfo = SelfSellLocalInterNetTicket(adyDate_new, aszBusID_new, SellStationID, aszEndStationID_new, atSellTKParam, anInsurrance, aszUserID, aszBankTime, oDb)
        SellInetNetTicket = atSellTicketResultInfo
    Dim ssql As String
    rsTemp.MoveFirst
    For i = 1 To rsTemp.RecordCount '把票号返写回网上售票表.
        ssql = "UPDATE internet_sell_ticket_info SET if_print=1 , print_user_id ='" & m_oActiveUser.UserID & "',ticket_id='" & atBuyTicketInfo(i).szTicketNo & "',print_time='" & oParam.NowDateTime & "' WHERE bus_id='" & Trim(rsTemp!bus_id) & "' AND bus_date='" & ToDBDate(Trim(rsTemp!bus_date)) & "'" _
            & " AND bank_card_id='" & Trim(rsTemp!bank_card_id) & "' AND busy_date='" & Trim(rsTemp!busy_date) & "'  AND seat_no='" & Trim(rsTemp!seat_no) & "'  AND getticket_id='" & Trim(rsTemp!getticket_id) & "'  AND validate_id='" & Trim(rsTemp!validate_id) & "' "
        oDb.Execute ssql
        rsTemp.MoveNext
    Next i
        If bChgTkNum = True Then
        szsql = "update system_param_info set parameter_value='" & szEndTicket & "' where parameter_name='StrTicketID'"
        oDb.Execute szsql
        End If
    oDb.CommitTrans
    Exit Function
Error_Handle:
    oDb.RollbackTrans
    err.Raise err.Number
End Function
'网上售票
Public Function SellInterNetTicket(BusDate() As Date, BusID() As String, SellStationID() As String, DesStationID() As String, aspSellTicketParam() As TSellTicketParam, anInsurance() As Integer, aszUserID() As String, Optional oDb As RTConnection) As TSellTicketResult()
    'anInsurance 是 保险费
    

End Function

'网上内部本地售票
Public Function SelfSellLocalInterNetTicket(BusStartupDate() As Date, BusID() As String, SellStationID() As String, DesStationID() As String _
             , aspSellTicketParam() As TSellTicketParam, panInsurance() As Integer, paszUserID() As String, paszBankTime() As Date, Optional oDb As RTConnection) As TSellTicketResult()
    'panInsurance 保险费
    
'    Dim oDb As New RTConnection
    Dim rsTemp() As TSellTicketResult
    Dim nCount As Integer
    Dim iLen As Integer
    iLen = ArrayLength(BusID)

    
    AssertActiveUserValid m_oActiveUser, ERR_SellTicketService
    AssertHaveRight m_oActiveUser, RIGHT_SellTicket
    ReDim rsTemp(1 To iLen)

'    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    
    oDb.BeginTrans
    On Error GoTo Here
    For nCount = 1 To iLen
        rsTemp(nCount) = SelfSellInterNetTicketLocalTrans(BusStartupDate(nCount), BusID(nCount), SellStationID(nCount), DesStationID(nCount), aspSellTicketParam(nCount).BuyTicketInfo, aspSellTicketParam(nCount).pasgSellTicketPrice, panInsurance(nCount), paszUserID(nCount), paszBankTime(nCount), False, ST_TicketNormal, oDb)
    Next nCount
    SelfSellLocalInterNetTicket = rsTemp
    oDb.CommitTrans
    Exit Function
Here:
    oDb.RollbackTrans
    Set oDb = Nothing
    
        err.Raise err.Number
End Function


'网上内部售票事务函数
Private Function SelfSellInterNetTicketLocalTrans(BusStartupDate As Date, BusID As String, SellStationID As String, DesStationID As String, SellTicketInfo() As TBuyTicketInfo _
    , pasgSellTicketPrice() As Single, pnInsurance As Integer, pszUserID As String, paszBankTime As Date, Optional pbExtraSell As Boolean = False _
    , Optional pnTicketStatus As ETicketStatus = ST_TicketNormal, Optional poDb As RTConnection) As TSellTicketResult

    Dim oParam As New SystemParam
    Dim szsql As String
    Dim rsTemp As Recordset
    Dim dtBusFullDateTime As Date
    Dim nTemp As Double
    Dim dtNow As Date
    Dim i As Integer
    Dim nBusStatus As EREBusStatus
    Dim nBuyTicketCount As Integer
    Dim sgFullPrice As Double, sgHalfPrice As Double
    Dim lRowAffect As Long, nStationLimitCount As Integer
    Dim szSeatNo As String, sgTicketPrice As Double
    Dim nBusType As EBusType
    
    Dim atiTicketPrice() As TRETicketPrice
    Dim siSellInfo As TSellTicketResult
    
    Dim lErrorCode As Long
    Dim rsRange As Recordset
    
    Dim nStationRemainCount As Integer
    
    Dim sgPreferentialTicket1 As Single '折扣票1
    Dim sgPreferentialTicket2 As Single '折扣票2
    Dim sgPreferentialTicket3 As Single '折扣票3
    
    
    Dim bNeedBook As Boolean '是否上了预定模块
    Dim bBookValid As Boolean '是否还需判断预定
    Dim bSeatBooked As Boolean
    Dim szBookStationID As String '预定的到站
    Dim rsTemp2 As Recordset
    
    Dim szSeatType As String '座位类型
    Dim szSqlSeat As String  '座位票价
    Dim nSeatStatus As Integer '座位状态
    
    Dim szStopTime As String
    
    Dim nSellTicketType As Integer
    Dim nSellTicketStatus As ETicketStatus
    
    
    Dim bSellStationCanSellEachOther As Boolean
    
    
    ReDim atiTicketPrice(1 To TP_TicketTypeCount)
    
    nSellTicketStatus = pnTicketStatus
    
    Select Case pnTicketStatus
        Case ST_TicketReturned
            nSellTicketType = TP_ExtraSold
            nSellTicketStatus = ST_TicketNormal
            
        Case ST_TicketSellChange
            nSellTicketType = TP_ChangeSold
        Case ST_TicketNormal
            nSellTicketType = TP_NormalSold
        Case Else
            nSellTicketType = 0
        
    End Select
    
    oParam.Init m_oActiveUser
    '判断车次的日期是否超过预售期限
    
    dtNow = oParam.NowDateTime
    
    nTemp = DateDiff("d", dtNow, BusStartupDate)
    
    '如果要购票的车次不在当天和预售天之内则出错
'    If nTemp > oParam.PreSellDate Or nTemp < 0 Then ShowError ERR_BusNotInIlegalDate
'    If nTemp > oParam.UserCanSellDate Then ShowError ERR_UserCannotSell
    
    bSellStationCanSellEachOther = m_oParam.SellStationCanSellEachOther
    'poDb.ConnectionString = GetConnectionStr(cszSellTicket)
    '得到环境车次及车次站点信息
'    szSql = "SELECT * FROM work_env_bus_info,work_env_bus_station_lst " _
'    & " WHERE work_env_bus_info.bus_id='" & BusID & "' AND " _
'    & " work_env_bus_info.bus_date='" & ToDBDate(BusStartupDate) & "' AND " _
'    & " work_env_bus_info.bus_id=work_env_bus_station_lst.bus_id AND " _
'    & " work_env_bus_info.bus_date=work_env_bus_station_lst.bus_date AND " _
'    & " station_id='" & DesStationID & "'" _
'    & " AND work_env_bus_station_lst.sell_station_id='" & SellStationID & "'"
'    Set rsTemp = poDb.Execute(szSql)


    szsql = "SELECT e.bus_date , e.bus_id , e.vehicle_type_name , e.vehicle_type_code , e.total_seat " _
        & " , e.sale_seat_quantity , e.total_stand_quantity , e.sale_stand_seat_quantity , e.bus_start_time , e.bus_type" _
        & " , e.is_all_refundment , e.route_id , e.end_station_id , e.end_station_name" _
        & " , e.owner_id , e.transport_company_id , e.vehicle_id , e.split_company_id , e.baggage_number , e.fact_weight" _
        & " , e.calculate_weight , e.over_weight_number , e.luggage_total_price ,  e.scrollbus_check_time" _
        & " , e.seat_remain , e.bed_remain  , e.additional_remain , e.other_remain_1 , e.other_remain_2" _
        & " , s.sell_station_id , s.station_id , s.seat_type_id , s.station_serial_no , s.max_sale_quantity , s.sale_ticket_quantity" _
        & " ,  s.stop_sale_time , s.mileage  , s.full_price , s.half_price , s.preferential_ticket1 , s.preferential_ticket2" _
        & " , s.preferential_ticket3" _
        & " , a.check_gate_id , a.bus_start_time" _
        & " , a.status , a.can_sale_quantity , a.sell_quantity" _
        & " FROM work_env_bus_info e ,work_env_bus_station_lst s , work_env_bus_allot_lst a" _
        & " WHERE e.bus_id='" & BusID & "' AND  e.bus_date='" & ToDBDate(BusStartupDate) & "'" _
        & " AND  e.bus_id=s.bus_id AND  e.bus_date=s.bus_date" _
        & " AND  s.station_id='" & DesStationID & "' AND s.sell_station_id='" & SellStationID & "'" _
        & " AND a.bus_id = s.bus_id AND a.bus_date = s.bus_date AND a.sell_station_id = '" & SellStationID & "'" _
        
        
    Set rsTemp = poDb.Execute(szsql)
    
    
    
    
    If rsTemp.RecordCount = 0 Then ShowError ERR_NoThisBus '没有指定的车次或车次不过指定的到站
    nBusType = FormatDbValue(rsTemp!bus_type)
    '流水车次不能补票
'    If nBusType = TP_ScrollBus And pbExtraSell Then ShowError ERR_ScrollBusCanNotExtraSell
    nBusStatus = FormatDbValue(rsTemp!Status)
    '并班车次不能售票
'    If nBusStatus = ST_BusMergeStopped Then ShowError ERR_BusMerged
'    '拆分停班车次不能售票
'    If nBusStatus = ST_BusSlitpStop Then ShowError ERR_BusSlitpStop
'    '车次正在拆分不能售票
'    If nBusStatus >= 32 Then ShowError ERR_BusIsSplit
'    '停检之后不能售票（实际不太可能出现，除非停售时间设得太小了）如果是补票则不限制
''    If pbExtraSell And nBusStatus = ST_BusStopCheck Then ShowError ERR_BusStopCheck
'    If nBusStatus = ST_BusStopCheck Then ShowError ERR_BusStopCheck
'    If nBusStatus = ST_BusStopped Then ShowError ERR_BusStopped
    dtBusFullDateTime = CDate(Format(FormatDbValue(rsTemp!bus_date), cszDateStr) & " " & Format(FormatDbValue(rsTemp!bus_start_time), cszTimeStr))
    '判断是否到了停售时间(如果是补票则要过了正常售票时间才能补)
    If nBusType <> TP_ScrollBus Then '如果是固定车次才限制
        If pbExtraSell Then
'            If DateDiff("s", dtNow, dtBusFullDateTime) >= oParam.StopSellTime * 60 Then ShowError ERR_BusNotStopSellTicket
        Else
'            If DateDiff("s", dtNow, dtBusFullDateTime) < oParam.StopSellTime * 60 Then ShowError ERR_BusStopSellTicket
        End If
    End If
    szStopTime = FormatDbValue(rsTemp!stop_sale_time)
    If nBusType <> TP_ScrollBus Then '如果是固定车次才限制
        If szStopTime <> "-1" Then '如果该车次到指定的站限制售票时间则要进行判断
'            If DateDiff("n", dtNow, dtBusFullDateTime) > GetStopTime(szStopTime) Then ShowError ERR_BusStaionStopSellTicket, "SellTicketClient.SelfSellTicketLocalTrans"
        End If
        bNeedBook = oParam.IsBookValid
        If bNeedBook Then
            bBookValid = IIf(DateAdd("n", oParam.BookTime, dtNow) >= rsTemp!bus_start_time, False, True)
        End If
    Else
        If szStopTime <> "-1" Then
           dtBusFullDateTime = CDate(Format(FormatDbValue(rsTemp!bus_date), cszDateStr) & " " & GetStopTime(szStopTime, False))
'           If DateDiff("n", dtNow, dtBusFullDateTime) < 0 Then ShowError ERR_BusStaionStopSellTicket, "SellTicketService.SelfSellTicket"
        End If
    End If
    nBuyTicketCount = ArrayLength(SellTicketInfo)
    If nBuyTicketCount > 0 Then '要售票的张数大于0
        nStationLimitCount = FormatDbValue(rsTemp!max_sale_quantity)
        nStationRemainCount = FormatDbValue(rsTemp!sale_ticket_quantity)
        If nStationRemainCount > 0 Then  '此处为了折扣票由>= 改为了>
            If nBuyTicketCount > nStationRemainCount Then ShowError ERR_BusReachMaxStationTicketCount
        End If
        
        sgFullPrice = FormatDbValue(rsTemp!full_price)
        sgHalfPrice = FormatDbValue(rsTemp!half_price)
        sgPreferentialTicket1 = FormatDbValue(rsTemp!preferential_ticket1)
        sgPreferentialTicket2 = FormatDbValue(rsTemp!preferential_ticket2)
        sgPreferentialTicket3 = FormatDbValue(rsTemp!preferential_ticket3)

        ReDim siSellInfo.asgTicketPrice(1 To nBuyTicketCount)
        ReDim siSellInfo.aszSeatNO(1 To nBuyTicketCount)
        ReDim siSellInfo.aszTicketType(1 To nBuyTicketCount)
        ReDim siSellInfo.aszSeatType(1 To nBuyTicketCount)
        For i = 1 To nBuyTicketCount
            szSeatNo = ""
            Select Case SellTicketInfo(i).nTicketType
            Case TP_FullPrice, TP_FreeTicket
                sgTicketPrice = sgFullPrice
            Case TP_HalfPrice
                sgTicketPrice = sgHalfPrice
            Case TP_PreferentialTicket1
                sgTicketPrice = sgPreferentialTicket1
            Case TP_PreferentialTicket2
                sgTicketPrice = sgPreferentialTicket2
            Case TP_PreferentialTicket3
                sgTicketPrice = sgPreferentialTicket3
            End Select
            
            
            
            
            If nBusType <> TP_ScrollBus Then '如果是固定车次才限制
                '判断到站的限售张数够用否
                If nStationLimitCount >= 0 Then
                    If nStationLimitCount = 0 Then
'                        ShowError ERR_BusReachMaxStationTicketCount
                    End If
                End If
            End If
            If nBusType <> TP_ScrollBus Then '如果是固定车次才进行座位处理
                If SellTicketInfo(i).szSeatNo <> cszStandTicketSeatNo Then
                    If SellTicketInfo(i).szSeatNo = cszAutoSeatNo Then  '系统自动给出座位号
RetrySeatNo:
                        '取座位该票价和类型
                        If bNeedBook And Not bBookValid Then
                            '如果可以预定
                            szsql = "SELECT p.ticket_price_total AS TicketPrice, s.seat_no AS SeatNo,s.status,p.seat_type_id AS SeatType FROM Env_bus_seat_lst s,work_env_bus_price_lst p" _
                            & " WHERE " _
                            & " s.bus_id='" & BusID & "' AND " _
                            & " s.bus_date='" & ToDBDate(BusStartupDate) & "' AND " _
                            & " p.sell_station_id='" & SellStationID & "' AND" _
                            & " s.seat_no=" _
                            & "(SELECT MIN(seat_no) FROM Env_bus_seat_lst WHERE " _
                            & " bus_id='" & BusID & "' AND " _
                            & " bus_date='" & ToDBDate(BusStartupDate) & "' AND (" _
                            & " status=" & ST_SeatCanSell & " OR status=" & ST_SeatBooked & " or status=" & ST_SeatProjectBooked & ") AND seat_type_id ='" & SellTicketInfo(i).szSeatTypeID & "') AND " _
                            & "s.bus_id=p.bus_id AND " _
                            & "s.bus_date=p.bus_date AND "
                            szsql = szsql & "s.seat_type_id=p.seat_type_id AND " _
                                        & "p.station_id='" & DesStationID & "' AND " _
                                        & "p.ticket_type='" & IIf(SellTicketInfo(i).nTicketType = TP_FreeTicket, TP_FullPrice, SellTicketInfo(i).nTicketType) & "'"
                        Else
                            szsql = "SELECT p.ticket_price_total AS TicketPrice,s.seat_no AS SeatNo,s.seat_type_id AS SeatType,p.ticket_price_total AS TicketPrice,s.status  FROM Env_bus_seat_lst s, work_env_bus_price_lst p WHERE " _
                            & " s.bus_id='" & BusID & "' AND " _
                            & " p.sell_station_id='" & SellStationID & "' AND " _
                            & " s.bus_date='" & ToDBDate(BusStartupDate) & "' AND " _
                            & "s.seat_no=(SELECT MIN(seat_no) FROM Env_bus_seat_lst WHERE bus_id='" & BusID & "' AND " _
                            & "bus_date='" & ToDBDate(BusStartupDate) & "' AND status=" & ST_SeatCanSell & " AND seat_type_id ='" & SellTicketInfo(i).szSeatTypeID & "') AND " _
                            & "s.bus_id=p.bus_id AND " _
                            & "s.bus_date=p.bus_date AND "
                            szsql = szsql & "s.seat_type_id=p.seat_type_id AND " _
                                        & "p.station_id='" & DesStationID & "' AND " _
                                        & "p.ticket_type='" & IIf(SellTicketInfo(i).nTicketType = TP_FreeTicket, TP_FullPrice, SellTicketInfo(i).nTicketType) & "'"
                        End If
                                              
                        Set rsTemp = poDb.Execute(szsql)
                        If rsTemp.RecordCount = 0 Then ShowError ERR_BusReachMaxTicketCount, "SellTicketClient.SelfSellTicketLocalTrans", "或" & SellTicketInfo(i).szSeatTypeName & "座位已售完。"
                        szSeatNo = FormatDbValue(rsTemp!SeatNo) '得到座位号
                        szSeatType = FormatDbValue(rsTemp!SeatType) '得到座位类型
                        sgTicketPrice = rsTemp!TicketPrice   '得到票价
                        nSeatStatus = rsTemp!Status
                        
                        If bNeedBook And Not bBookValid Then
                            bSeatBooked = IIf(rsTemp!Status = ST_SeatBooked, True, False)
'                            If rsTemp!Status = ST_SeatProjectBooked Then
'                                AssertHaveRight m_oActiveUser, RIGHT_RightOther1 '判断是否有计划预定权限
'                            End If
                        Else
                            bSeatBooked = False
                        End If
                        If bSeatBooked Then
                            szsql = UpdateSeatStatusStr2(BusID, BusStartupDate, szSeatNo, SellTicketInfo(i).szTicketNo, bBookValid)
                        Else
                            szsql = UpdateSeatStatusStr(BusID, BusStartupDate, szSeatNo, SellTicketInfo(i).szTicketNo)
                        End If
                        If rsTemp!Status = ST_SeatProjectBooked Then szsql = UpdateSeatStatusStr3(BusID, BusStartupDate, szSeatNo, SellTicketInfo(i).szTicketNo)
                        poDb.Execute szsql, lRowAffect
                        
                        If lRowAffect = 0 Then GoTo RetrySeatNo
                    Else
                        szSeatNo = SellTicketInfo(i).szSeatNo
                        
                         '取座位该票价和类型
                        szSqlSeat = "SELECT p.ticket_price_total AS TicketPrice, p.seat_type_id AS SeatType,s.status AS SeatStatus " _
                                        & " FROM work_env_bus_price_lst p,Env_bus_seat_lst s WHERE " _
                                        & " p.bus_id='" & BusID & "' AND p.station_id='" & DesStationID & "' AND " _
                                        & " p.bus_date='" & ToDBDate(BusStartupDate) & "' AND " _
                                        & " p.ticket_type=" & SellTicketInfo(i).nTicketType & " AND " _
                                        & " p.seat_type_id =(SELECT seat_type_id FROM Env_bus_seat_lst WHERE " _
                                        & " bus_id='" & BusID & "' AND bus_date='" & ToDBDate(BusStartupDate) & "' AND seat_no='" & szSeatNo & "') AND " _
                                        & " s.status=(SELECT status FROM Env_bus_seat_lst WHERE " _
                                        & " bus_id='" & BusID & "' AND bus_date='" & ToDBDate(BusStartupDate) & "' AND seat_no='" & szSeatNo & "')" _
                                        & " AND p.sell_station_id ='" & SellStationID & "'" _
                                        & " AND p.bus_id = s.bus_id AND p.bus_date = s.bus_date AND s.seat_no = '" & szSeatNo & "'" '
                                        
                        Set rsTemp = poDb.Execute(szSqlSeat)
                        
                        If rsTemp.RecordCount = 0 Then ShowError ERR_BusReachMaxTicketCount
                        sgTicketPrice = rsTemp!TicketPrice
                        szSeatType = FormatDbValue(rsTemp!SeatType)
                        nSeatStatus = rsTemp!seatstatus
                        
                        szsql = UpdateSeatStatusStr(BusID, BusStartupDate, SellTicketInfo(i).szSeatNo, SellTicketInfo(i).szTicketNo)
                        poDb.Execute szsql, lRowAffect
                        If lRowAffect = 0 Then
                            If bNeedBook Then
                                If nSeatStatus = ST_SeatProjectBooked Then
                                
                                    AssertHaveRight m_oActiveUser, RIGHT_RightOther1
                                    szsql = UpdateSeatStatusStr3(BusID, BusStartupDate, SellTicketInfo(i).szSeatNo, SellTicketInfo(i).szTicketNo)
                                    poDb.Execute szsql, lRowAffect
                                    If lRowAffect = 0 Then ShowError ERR_SeatNoCanNotSell, "SellTicketService.SelfSellTicket"
                                    bSeatBooked = False
                                Else
                                    szsql = UpdateSeatStatusStr2(BusID, BusStartupDate, SellTicketInfo(i).szSeatNo, SellTicketInfo(i).szTicketNo, bBookValid, SellTicketInfo(i).szReserved)
                                    poDb.Execute szsql, lRowAffect
                                    If lRowAffect = 0 Then ShowError ERR_SeatNoCanNotSell, "SellTicketClient.SelfSellTicketLocalTrans"
                                
                                    bSeatBooked = True
                                End If
                            Else
                                ShowError ERR_SeatNoCanNotSell, "SellTicketClient.SelfSellTicketLocalTrans"
                            End If
                        Else
                            bSeatBooked = False
                        End If
                        
                    End If
                    '如果座位预定
                    If bSeatBooked Then
                        szsql = "SELECT * FROM ticket_book_lst WHERE " _
                        & " bus_id='" & BusID & "' AND " _
                        & " bus_date='" & ToDBDate(BusStartupDate) & "' AND " _
                        & " seat_no='" & szSeatNo & "' AND " _
                        & " status=0"
                        Set rsTemp2 = poDb.Execute(szsql)
                        szBookStationID = Trim(rsTemp2!station_id)
                        Set rsTemp2 = Nothing

                        
                        'If nStationLimitCount > 0 Then
                            '更新站点的预定张数
                        szsql = "UPDATE Env_bus_station_lst SET " _
                        & " book_quantity=book_quantity-1 WHERE " _
                        & " bus_id='" & BusID & "' AND " _
                        & " bus_date='" & ToDBDate(BusStartupDate) & "' AND " _
                        & " station_id='" & szBookStationID & "' AND " _
                        & " sell_station_id='" & SellStationID & "' AND " _
                        & " book_quantity>0"
                        
                        poDb.Execute szsql
                        'End If
                        
                        If bBookValid Then
                            szsql = "UPDATE ticket_book_lst SET status=2 WHERE " _
                            & " bus_id='" & BusID & "' AND " _
                            & " bus_date='" & ToDBDate(BusStartupDate) & "' AND " _
                            & " seat_no='" & szSeatNo & "' AND " _
                            & " status=0 " '0状态为预定
                        Else
                            szsql = "UPDATE ticket_book_lst SET status=3 WHERE " _
                            & " bus_id='" & BusID & "' AND " _
                            & " bus_date='" & ToDBDate(BusStartupDate) & "' AND " _
                            & " seat_no='" & szSeatNo & "' AND " _
                            & " status=0 " '0状态为预定
                        End If
                        poDb.Execute szsql
                    End If
                    
                    '得到环境车次配载信息
                    Set rsTemp2 = poDb.Execute("SELECT * from work_env_bus_allot_lst WHERE bus_date='" _
                    & ToDBDate(BusStartupDate) & "' AND " _
                    & " bus_id='" & BusID & "' AND " _
                    & " sell_station_id='" & SellStationID & "'")
                    If rsTemp2.RecordCount = 0 Then ShowError ERR_BusAllotInfoNotExist
                    
                    '总共允许售票张数是否大于"已售张数"+1
                    
                    If rsTemp2!can_sale_quantity >= rsTemp2!sell_quantity + 1 Or rsTemp2!can_sale_quantity = -1 Then
                        poDb.Execute ("UPDATE Env_bus_allot_lst SET sell_quantity=sell_quantity + 1 WHERE bus_date='" _
                            & ToDBDate(BusStartupDate) & "' AND " _
                            & " bus_id='" & BusID & "' AND " _
                            & " sell_station_id='" & SellStationID & "'")
                    Else
                        ShowError ERR_SellQuantityISFull '超过允许售票张数
                    End If
                    
                    If bSellStationCanSellEachOther Then
                        '得到环境车次售票点信息
                        Set rsTemp2 = poDb.Execute("SELECT * from work_env_bus_sell_station_lst WHERE bus_date='" _
                        & ToDBDate(BusStartupDate) & "' AND " _
                        & " bus_id='" & BusID & "' AND " _
                        & " sell_station_id='" & m_oActiveUser.SellStationID & "'")
                        If rsTemp2.RecordCount = 0 Then ShowError ERR_BusSellStationInfoNotExist
                        
                        '总共允许售票张数是否大于"已售张数"+1
                        
                        If rsTemp2!can_sale_quantity >= rsTemp2!sell_quantity + 1 Or rsTemp2!can_sale_quantity = -1 Then
                            poDb.Execute ("UPDATE env_bus_sell_station_lst SET sell_quantity=sell_quantity + 1 WHERE bus_date='" _
                                & ToDBDate(BusStartupDate) & "' AND " _
                                & " bus_id='" & BusID & "' AND " _
                                & " sell_station_id='" & m_oActiveUser.SellStationID & "'")
                        Else
                            ShowError ERR_SellQuantityISFull '超过允许售票张数
                        End If
                    End If

                    Set rsTemp2 = Nothing
                    If nSeatStatus = ST_SeatProjectBooked Or nSeatStatus = ST_SeatBooked Then
                        '如果为售预定票,则不减剩余座位数
                    Else
                        '判断到站的限售张数够用否
                        If nStationRemainCount >= 0 Then
                            If nStationRemainCount = 0 Then
'                                ShowError ERR_BusReachMaxStationTicketCount, "SellTicketClient.SelfSellTicketLocalTrans"
                            Else
                                '更新可售张数
                                szsql = "UPDATE Env_bus_station_lst SET " _
                                & " sale_ticket_quantity=sale_ticket_quantity-1 WHERE " _
                                & " bus_id='" & BusID & "' AND " _
                                & " bus_date='" & ToDBDate(BusStartupDate) & "' AND " _
                                & " sale_ticket_quantity>book_quantity AND " _
                                & " station_id='" & DesStationID & "'"
'                                & " sell_station_id='" & SellStationID & "' AND "
                                poDb.Execute szsql, lRowAffect
                                
                                If lRowAffect = 0 Then ShowError ERR_BusReachMaxStationTicketCount, "SellTicketClient.SelfSellTicketLocalTrans"
                            End If
                            
                        End If
                        szsql = "UPDATE Env_bus_info SET sale_seat_quantity=sale_seat_quantity-1 , "
                        Select Case szSeatType
                            Case cszSeatType
                                szsql = szsql & "seat_remain=seat_remain - 1 "
                            Case cszBedType
                                szsql = szsql & "bed_remain=bed_remain - 1 "
                            Case cszAdditionalType
                                szsql = szsql & "additional_remain=additional_remain - 1 "
                            Case cszOtherType1
                                szsql = szsql & "other_remain_1 =other_remain_1 - 1"
                            Case cszOtherType2
                                szsql = szsql & "other_remain_2=other_remain_2 - 1"
                        End Select
                        szsql = szsql & " WHERE bus_id='" & BusID & "' AND bus_date='" & ToDBDate(BusStartupDate) & "'"
                        
                        poDb.Execute szsql
                    End If
                    
                Else
                    szsql = "UPDATE Env_bus_info SET sale_stand_seat_quantity=sale_stand_seat_quantity-1 " _
                    & " WHERE bus_id='" & BusID & "' AND " _
                    & "bus_date='" & ToDBDate(BusStartupDate) & "' AND " _
                    & " sale_stand_seat_quantity>0 "
                    
                    poDb.Execute szsql, lRowAffect
                    If lRowAffect = 0 Then ShowError ERR_NoStandTicket, "SellTicketClient.SelfSellTicketLocalTrans"
                    szSeatNo = cszStandTicketSeatNo
                End If
            Else
                szSeatType = cszSeatType
            End If
            '如果票价为0且不为免票,则不允许出售
            If sgTicketPrice = 0 And SellTicketInfo(i).nTicketType <> TP_FreeTicket Then
                RaiseError ERR_CannotSellZeroPrice
            End If
            
            
            '写售票明细表
            szsql = "INSERT INTO Ticket_sell_lst(" _
            & " sell_station_id," _
            & " ticket_id," _
            & " bus_date," _
            & " bus_id," _
            & " des_station_id," _
            & " status," _
            & " ticket_type," _
            & " user_id," _
            & " seat_no," _
            & " ticket_price," _
            & " sale_ticket_type," _
            & " seat_type_id ," _
            & " operation_time, " _
            & " insurance ) "
            
            szsql = szsql & " VALUES('" _
            & SellStationID & "','" _
            & SellTicketInfo(i).szTicketNo & "','" _
            & ToDBDate(BusStartupDate) & "','" _
            & BusID & "','" _
            & DesStationID & "'," _
            & nSellTicketStatus & "," _
            & SellTicketInfo(i).nTicketType & ",'" _
            & pszUserID & "','" _
            & szSeatNo & "'," _
            & sgTicketPrice & "," _
            & nSellTicketType & ",'" _
            & szSeatType & "','" _
            & ToDBDateTime(paszBankTime) & "', " _
            & TransFieldValueToString(pnInsurance) & ")"
            On Error GoTo MayBeTicketIDDup
            poDb.Execute szsql
            
            '写详细票价信息
'            If atiTicketPrice(SellTicketInfo(i).nTicketType).sgBase = 0 Then
                szsql = GetTicketPriceDetailStr(BusID, BusStartupDate, SellStationID, DesStationID, szSeatType, IIf(SellTicketInfo(i).nTicketType = TP_FreeTicket, TP_FullPrice, SellTicketInfo(i).nTicketType))
                Set rsTemp = poDb.Execute(szsql)
                LetPriceDetailFromRsToType rsTemp, atiTicketPrice(SellTicketInfo(i).nTicketType), 0
'            End If
            szsql = GetInsertPriceDetailStr(atiTicketPrice(SellTicketInfo(i).nTicketType), SellTicketInfo(i).szTicketNo)
            poDb.Execute szsql
            
            siSellInfo.aszSeatType(i) = szSeatType
            siSellInfo.aszTicketType(i) = SellTicketInfo(i).nTicketType
            siSellInfo.asgTicketPrice(i) = sgTicketPrice
            siSellInfo.aszSeatNO(i) = szSeatNo
        Next
        szsql = "SELECT station_name FROM station_info WHERE station_id='" & DesStationID & "'"
        Set rsTemp = poDb.Execute(szsql)
        
        siSellInfo.szDesStationName = FormatDbValue(rsTemp!station_name)
        
    Else
        ShowError ERR_NoTicketInfoSpecified
    End If
    
    SelfSellInterNetTicketLocalTrans = siSellInfo
    
    
'    Dim nIndex As Integer
'    Dim szStartStationID As String, szStartStationName As String
'    Dim szUnitID As String
'    Dim szDesStationName As String
'    szDesStationName = siSellInfo.szDesStationName
'
'    nBuyTicketCount = ArrayLength(SellTicketInfo)
'    '此处需改为从参数上车站中得到的站点及站点名称
'    szSql = " SELECT s.* , t.station_name FROM sell_station_info s, station_info t WHERE s.station_id = t.station_id "
'    Set rstemp = poDb.Execute(szSql)
'    szStartStationID = FormatDbValue(rstemp!station_id)
'    szStartStationName = FormatDbValue(rstemp!station_name)
'    '不需要从系统参数中得到,改为从传入的参数中得到
'
''    szStartStationID = m_oCurrentSellService.StationID
''    szStartStationName = m_oCurrentSellService.StationName
'
'    szUnitID = m_oCurrentSellService.SelfUser.ServerUnit
'
'    For nIndex = 1 To nBuyTicketCount
'        '写代理售票明细表
'        szSql = "INSERT INTO Agent_ticket_sale_lst(" _
'        & " sell_station_id," _
'        & " ticket_id," _
'        & " bus_date," _
'        & " bus_id," _
'        & "service_unit_id," _
'        & "start_station_id," _
'        & "start_station_name," _
'        & " des_station_id," _
'        & " des_station_name," _
'        & " ticket_type," _
'        & " user_id," _
'        & "seat_no," _
'        & " ticket_price," _
'        & "status," _
'        & " sale_ticket_type, " _
'        & " seat_type_id," _
'        & " operation_time, " _
'        & " insurance )  " _
'
'
'        szSql = szSql & " VALUES('" _
'        & SellStationID & "','" _
'        & SellTicketInfo(nIndex).szTicketNo & "','" _
'        & ToDBDate(BusStartupDate) & "','" _
'        & BusID & "','" _
'        & szUnitID & "','" _
'        & szStartStationID & "','" _
'        & szStartStationName & "','" _
'        & DesStationID & "','" _
'        & szDesStationName & "'," _
'        & SellTicketInfo(nIndex).nTicketType & ",'" _
'        & m_oActiveUser.UserID & "','" _
'        & siSellInfo.aszSeatNO(nIndex) & "'," _
'        & siSellInfo.asgTicketPrice(nIndex) & "," _
'        & nSellTicketStatus & "," _
'        & nSellTicketType & ",'" _
'        & siSellInfo.aszSeatType(nIndex) & "','" _
'        & ToDBDateTime(dtNow) & "'," _
'        & TransFieldValueToString(pnInsurance) & " ) "
'
'        poDb.Execute szSql
'    Next nIndex
    Set rsTemp = Nothing
    Set rsRange = Nothing
    Set rsTemp2 = Nothing
    Set oParam = Nothing
Exit Function
MayBeTicketIDDup:
    
    If poDb.HaveThisNativeErr(ERR_DBAddDuplicate) Then

        ShowError ERR_TicketNoDup
    Else
        
        err.Raise err.Number
    End If
     Set rsTemp = Nothing
     Set rsRange = Nothing
     Set rsTemp2 = Nothing
     Set oParam = Nothing
End Function




Public Function GetLocalBusEXRs(BusDate As Date, BusID As String, SellStationID As String, EndStationID As String) As Recordset

    Dim oDb As New RTConnection
    Dim oParam As New SystemParam
    Dim szsql As String
    Dim rsTemp As Recordset

    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    
    szsql = "select ebi.* ,e.check_gate_id sell_check_gate_id ,e.bus_start_time sell_bus_start_time ,sell_station_name,st.station_name  " _
        & " from Work_env_bus_info ebi ,Work_env_bus_allot_lst e ,sell_station_info s,station_info st" _
        & " WHERE  e.sell_station_id=s.sell_station_id AND  e.bus_date=ebi.bus_date AND e.bus_id=ebi.bus_id "
        
    szsql = szsql & " AND s.sell_station_id='" & SellStationID & "'and  ebi.bus_date='" & ToDBDate(BusDate) & "' and ebi.bus_id='" & BusID & "'" _
        & " and  st.station_id='" & EndStationID & "'"
    Set rsTemp = oDb.Execute(szsql)
    Set GetLocalBusEXRs = rsTemp
    Set rsTemp = Nothing
    Set oDb = Nothing
    
End Function


Public Function BookTK(BusDate() As Date, BusID() As String, SellStationID() As String, DesStationID() As String, DesStationName() As String, aspSellTicketParam() As TSellTicketParam, anInsurance() As Integer) As TSellTicketResult()
    
    Dim srTemp() As TSellTicketResult
    Dim iLen As Integer
    Dim iLen1 As Integer
    Dim nCount As Integer
    Dim aszDesStationName() As String
    
    iLen = 0
    iLen1 = 0
    iLen = ArrayLength(BusID)
    ReDim srTemp(1 To iLen)
    If Trim(m_oCurrentSellService.SelfUser.ServerUnit) <> Trim(m_oCurrentSellService.SelfUser.UserUnitID) Then
        '如果用户单位和服务单位不一样则是售远程票否则售的是本地票
        On Error GoTo ErrorHandle
        
        srTemp = m_oCurrentSellService.BookTK(BusDate, BusID, SellStationID, DesStationID, aspSellTicketParam, anInsurance)
    
    Else
        srTemp = BookLocalTicket(BusDate, BusID, SellStationID, DesStationID, aspSellTicketParam, anInsurance)
    End If
    BookTK = srTemp
    Exit Function
ErrorHandle:
    err.Raise err.Number
    Exit Function
ErrorHandle2:
    '加入提示"对方车站已写入,请通知管理员在对方车站将此票删除"
    err.Raise err.Number, err.Source, cszServiceFinish & err.Description
    Exit Function
End Function




Public Function BookLocalTicket(BusStartupDate() As Date, BusID() As String, SellStationID() As String, DesStationID() As String _
             , aspSellTicketParam() As TSellTicketParam, panInsurance() As Integer) As TSellTicketResult()
    'panInsurance 保险费

    Dim oDb As New RTConnection
    Dim rsTemp() As TSellTicketResult
    Dim nCount As Integer
    Dim iLen As Integer
    iLen = ArrayLength(BusID)


    AssertActiveUserValid m_oActiveUser, ERR_SellTicketService
    AssertHaveRight m_oActiveUser, RIGHT_SellTicket
    ReDim rsTemp(1 To iLen)

    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    
    oDb.BeginTrans
    On Error GoTo Here
    For nCount = 1 To iLen
        rsTemp(nCount) = BookLocalTrans(BusStartupDate(nCount), BusID(nCount), SellStationID(nCount), DesStationID(nCount), aspSellTicketParam(nCount).BuyTicketInfo, aspSellTicketParam(nCount).pasgSellTicketPrice, panInsurance(nCount), False, ST_TicketNormal, oDb)
    Next nCount
    BookLocalTicket = rsTemp
    oDb.CommitTrans
    Set oDb = Nothing
    Exit Function
Here:
    oDb.RollbackTrans
    Set oDb = Nothing
    
        err.Raise err.Number
End Function



Private Function BookLocalTrans(BusStartupDate As Date, BusID As String, SellStationID As String, DesStationID As String, SellTicketInfo() As TBuyTicketInfo _
    , pasgSellTicketPrice() As Single, pnInsurance As Integer, Optional pbExtraSell As Boolean = False _
    , Optional pnTicketStatus As ETicketStatus = ST_TicketNormal, Optional poDb As RTConnection) As TSellTicketResult
    
 
    Dim nBookSeatCount As Integer, i As Integer '预留座位数
    Dim oDb As New RTConnection
    Dim szsql As String
    Dim szBookID As String
    Dim oParam As New SystemParam
    Dim rsTemp As Recordset, rsTemp2 As Recordset
    Dim lRowAffect As Long
    Dim szTemp As String
    Dim nStationLimitedCount As Integer
    
    Dim BookID() As TBuyTicketInfo '预留票号
    Dim BookSeatID() As String '预留座位号
    
    Dim szSellStationID As String
    Dim szSeatID As String '座位号
    
    
    
    
    Dim dtBusFullDateTime As Date
    Dim nTemp As Double
    Dim dtNow As Date
    Dim nBusStatus As EREBusStatus
    Dim nBuyTicketCount As Integer
    Dim sgFullPrice As Double, sgHalfPrice As Double
    Dim nStationLimitCount As Integer
    Dim szSeatNo As String, sgTicketPrice As Double
    Dim nBusType As EBusType
    
    Dim atiTicketPrice() As TRETicketPrice
    Dim siSellInfo As TSellTicketResult
      
    Dim lErrorCode As Long
    Dim rsRange As Recordset
    
    Dim nStationRemainCount As Integer
    Dim nCanSellTicket As Integer '可售座位数
    Dim bNeedBook As Boolean '是否上了预定模块
    Dim bBookValid As Boolean '是否还需判断预定
    Dim bSeatBooked As Boolean
    Dim szBookStationID As String '预定的到站
    
    Dim szSeatType As String '座位类型
    Dim szSqlSeat As String  '座位票价
    Dim nSeatStatus As Integer '座位状态
    
    Dim szStopTime As String
    
    Dim nSellTicketType As Integer
    Dim nSellTicketStatus As ETicketStatus
    Dim tTSeatInfo() As TSeatInfoCount
    Dim j As Integer
    Dim szSqlSet As String
    Dim nSaleSeatQuantity As Integer
    
    oParam.Init m_oActiveUser

    dtNow = oParam.NowDateTime
    
    nTemp = DateDiff("d", dtNow, BusStartupDate)
    
    '如果要购票的车次不在当天和预售天之内则出错
    If nTemp > oParam.PreSellDate Or nTemp < 0 Then ShowError ERR_BusNotInIlegalDate
    If nTemp > oParam.UserCanSellDate Then ShowError ERR_UserCannotSell


   
'    '如果票价为0且不为免票,则不允许出售
'    For i = 1 To ArrayLength(pszTicketID)
'    If pszTPrice(1).pasgSellTicketPrice(i) = 0 And pszTicketID(i).nTicketType <> TP_FreeTicket Then
'        RaiseError ERR_CannotSellZeroPrice
'    End If
'    Next

    
    

    oDb.ConnectionString = GetConnectionStr("")
    
    szsql = "SELECT * FROM work_env_bus_info e,work_env_bus_station_lst s " _
    & " WHERE e.bus_id='" & BusID & "' AND " _
    & " e.bus_date='" & ToDBDate(BusStartupDate) & "' AND " _
    & " e.bus_id=s.bus_id AND " _
    & " e.bus_date=s.bus_date AND " _
    & " station_id='" & DesStationID & "' and s.sell_station_id='" & Trim(SellStationID) & "'"
    
    Set rsTemp = oDb.Execute(szsql)
    If rsTemp.RecordCount < 1 Then
        ShowError ERR_BusNotExist, "CBook.Book", "BusID=" & BusID & ",BusDate=" & ToDBDate(BusStartupDate)
    Else
        szSellStationID = FormatDbValue(rsTemp!sell_station_id)
        nStationLimitedCount = rsTemp!max_sale_quantity
        If nStationLimitedCount = 0 Then ShowError ERR_BusStationMatchLimitedCount, "CBook.Book"
    End If
    
'*********************************************************
    nBusType = FormatDbValue(rsTemp!bus_kind)
    '流水车次不能补票
'    If nBusType = TP_ScrollBus And pbExtraSell Then ShowError ERR_ScrollBusCanNotExtraSell
    nBusStatus = FormatDbValue(rsTemp!Status)
    '并班车次不能售票
    If nBusStatus = ST_BusMergeStopped Then ShowError ERR_BusMerged
    '拆分停班车次不能售票
    If nBusStatus = ST_BusSlitpStop Then ShowError ERR_BusSlitpStop
    '车次正在拆分不能售票
    If nBusStatus >= 32 Then ShowError ERR_BusIsSplit
    '停检之后不能售票（实际不太可能出现，除非停售时间设得太小了）如果是补票则不限制
'    If pbExtraSell And nBusStatus = ST_BusStopCheck Then ShowError ERR_BusStopCheck
    If nBusStatus = ST_BusStopCheck Then ShowError ERR_BusStopCheck
    If nBusStatus = ST_BusStopped Then ShowError ERR_BusStopped
    dtBusFullDateTime = CDate(Format(FormatDbValue(rsTemp!bus_date), cszDateStr) & " " & Format(FormatDbValue(rsTemp!bus_start_time), cszTimeStr))
    '判断是否到了停售时间(如果是补票则要过了正常售票时间才能补)
'    If nBusType <> TP_ScrollBus Then '如果是固定车次才限制
'        If pbExtraSell Then
'            If DateDiff("s", dtNow, dtBusFullDateTime) >= oParam.StopSellTime * 60 Then ShowError ERR_BusNotStopSellTicket
'        Else
''            If DateDiff("s", dtNow, dtBusFullDateTime) < oParam.StopSellTime * 60 Then ShowError ERR_BusStopSellTicket
'        End If
'    End If



    szStopTime = FormatDbValue(rsTemp!stop_sale_time)
    If nBusType <> TP_ScrollBus Then '如果是固定车次才限制
        If szStopTime <> "-1" Then '如果该车次到指定的站限制售票时间则要进行判断
            If DateDiff("n", dtNow, dtBusFullDateTime) > GetStopTime(szStopTime) Then ShowError ERR_BusStaionStopSellTicket, "SellTicketClient.SelfSellTicketLocalTrans"
        End If
        bNeedBook = oParam.IsBookValid
        If bNeedBook Then
            bBookValid = IIf(DateAdd("n", oParam.BookTime, dtNow) >= rsTemp!bus_start_time, False, True)
        End If
    Else
        If szStopTime <> "-1" Then
           dtBusFullDateTime = CDate(Format(FormatDbValue(rsTemp!bus_date), cszDateStr) & " " & GetStopTime(szStopTime, False))
           If DateDiff("n", dtNow, dtBusFullDateTime) < 0 Then ShowError ERR_BusStaionStopSellTicket, "SellTicketService.SelfSellTicket"
        End If
    End If
    nBuyTicketCount = ArrayLength(SellTicketInfo)
    nStationLimitCount = FormatDbValue(rsTemp!max_sale_quantity)
    nStationRemainCount = FormatDbValue(rsTemp!sale_ticket_quantity)
    nCanSellTicket = FormatDbValue(rsTemp!sale_seat_quantity)
    If nStationRemainCount > 0 Then  '此处为了折扣票由>= 改为了>
            If nBuyTicketCount > nStationRemainCount Then ShowError ERR_BusReachMaxStationTicketCount
    End If
    
    '判断可售座位数
    nBookSeatCount = ArrayLength(SellTicketInfo)
    
    If nCanSellTicket < nBookSeatCount Then
        ShowError ERR_BusReachMaxTicketCount
    End If
    
    
    '得到环境车次配载信息
    Set rsTemp2 = oDb.Execute("SELECT * from work_env_bus_allot_lst WHERE bus_date='" _
                    & ToDBDate(BusStartupDate) & "' AND " _
                    & " bus_id='" & BusID & "' AND " _
                    & " sell_station_id='" & SellStationID & "'")
    If rsTemp2.RecordCount = 0 Then ShowError ERR_BusAllotInfoNotExist
                    
    '总共允许售票张数是否大于"已售张数"+1
                    
    If rsTemp2!can_sale_quantity >= rsTemp2!sell_quantity + 1 Or rsTemp2!can_sale_quantity = -1 Then

    Else
        ShowError ERR_SellQuantityISFull '超过允许售票张数
    End If
    ReDim siSellInfo.aszSeatNO(1 To nBookSeatCount)
    ReDim siSellInfo.asgTicketPrice(1 To nBookSeatCount)
    ReDim siSellInfo.aszSeatType(1 To nBookSeatCount)
    ReDim siSellInfo.aszTicketType(1 To nBookSeatCount)
'*********************************************************
    
    oDb.BeginTrans
    On Error GoTo Error_Handle
    
    For i = 1 To nBookSeatCount
        ReDim tTSeatInfo(1 To 5)
        szSqlSet = ""
        szsql = "SELECT MIN(seat_no) seat_no FROM Env_bus_seat_lst WHERE " _
        & " bus_id='" & BusID & "' AND " _
        & " is_special_price= 0 AND " _
        & " bus_date='" & ToDBDate(BusStartupDate) & "' AND " _
        & " status=" & ST_SeatCanSell & " AND seat_type_id ='" & SellTicketInfo(i).szSeatTypeID & "'"
        Set rsTemp = oDb.Execute(szsql)
        If rsTemp.RecordCount > 0 Then
            szSeatID = rsTemp!seat_no
        Else
            ShowError ERR_BOOKSeatBooked, "CBook.Book", "BusID=" & BusID & ",BusDate=" & ToDBDate(BusStartupDate) & ",SeatNo=" & szSeatID
        End If
        szsql = "SELECT is_special_price FROM Env_bus_seat_lst WHERE " _
            & " bus_id='" & BusID & "' AND " _
            & " seat_no='" & szSeatID & "' AND " _
            & " bus_date='" & ToDBDate(BusStartupDate) & "' AND " _
            & " status=" & ST_SeatCanSell & " AND seat_type_id ='" & SellTicketInfo(i).szSeatTypeID & "'"
        Set rsTemp = poDb.Execute(szsql)
        '特价座位不能预订
        If rsTemp.RecordCount > 0 Then
'            If FormatDbValue(rsTemp!is_special_price) = 1 Then ShowError ERR_SpecialPriceSeatCanNotSell, "CBook.Book"
        End If
        
            szsql = "UPDATE Env_bus_seat_lst SET status=" & Str(ST_SeatReserved) & "   ," _
                  & " remark='self于" & Date & "留' " _
                  & "  WHERE bus_id='" & BusID & "' " _
                  & " AND  bus_date='" & ToDBDate(BusStartupDate) & "' " _
                  & " AND seat_no ='" & szSeatID & "' " _
                  & " AND  Status=" & ST_SeatCanSell & ""
            oDb.Execute szsql, lRowAffect
            If lRowAffect <> 1 Then ShowError ERR_BusStationMatchLimitedCount, "CBook.Book"
        
        
            szsql = " select seat_type_id ,count(*)as CountSeat from Env_bus_seat_lst " _
                    & " where  bus_id='" & BusID & "'and  status='" & ST_SeatCanSell & "' and bus_date='" & ToDBDate(BusStartupDate) & "' group by seat_type_id "
            Set rsTemp = oDb.Execute(szsql)
            
            For j = 1 To rsTemp.RecordCount
                tTSeatInfo(CInt(rsTemp!seat_type_id)).SeatCount = FormatDbValue(rsTemp!CountSeat)
                rsTemp.MoveNext
            Next
            
            szSqlSet = szSqlSet & "seat_remain ='" & tTSeatInfo(1).SeatCount & "',"
            szSqlSet = szSqlSet & "bed_remain ='" & tTSeatInfo(2).SeatCount & "',"
            szSqlSet = szSqlSet & "additional_remain ='" & tTSeatInfo(3).SeatCount & "',"
            szSqlSet = szSqlSet & "other_remain_1 ='" & tTSeatInfo(4).SeatCount & "',"
            szSqlSet = szSqlSet & "other_remain_2 ='" & tTSeatInfo(5).SeatCount & "',"
            nSaleSeatQuantity = 0
            For j = 1 To 5
              nSaleSeatQuantity = nSaleSeatQuantity + tTSeatInfo(j).SeatCount
            Next
            
            szSqlSet = szSqlSet & "sale_seat_quantity='" & nSaleSeatQuantity & "'"
            szsql = "UPDATE Env_bus_info SET  " & szSqlSet _
                    & "  WHERE bus_id='" & BusID & "' AND  bus_date='" & ToDBDate(BusStartupDate) & "'"
            oDb.Execute (szsql)
            
            
            
            szSqlSeat = "SELECT p.ticket_price_total AS TicketPrice, p.seat_type_id AS SeatType,s.status AS SeatStatus " _
                            & " FROM work_env_bus_price_lst p,Env_bus_seat_lst s WHERE " _
                            & " p.bus_id='" & BusID & "' AND p.station_id='" & DesStationID & "' AND " _
                            & " p.bus_date='" & ToDBDate(BusStartupDate) & "' AND " _
                            & " p.ticket_type=" & SellTicketInfo(i).nTicketType & " AND " _
                            & " p.seat_type_id =(SELECT seat_type_id FROM Env_bus_seat_lst WHERE " _
                            & " bus_id='" & BusID & "' AND bus_date='" & ToDBDate(BusStartupDate) & "' AND seat_no='" & szSeatID & "') AND " _
                            & " s.status=(SELECT status FROM Env_bus_seat_lst WHERE " _
                            & " bus_id='" & BusID & "' AND bus_date='" & ToDBDate(BusStartupDate) & "' AND seat_no='" & szSeatID & "')" _
                            & " AND p.sell_station_id ='" & SellStationID & "'" _
                            & " AND p.bus_id = s.bus_id AND p.bus_date = s.bus_date AND s.seat_no = '" & szSeatID & "'" '
                                        
            Set rsTemp = oDb.Execute(szSqlSeat)
            sgTicketPrice = rsTemp!TicketPrice
            szSeatType = FormatDbValue(rsTemp!SeatType)
            
            siSellInfo.aszSeatNO(i) = szSeatID
            siSellInfo.asgTicketPrice(i) = sgTicketPrice
            siSellInfo.aszSeatType(i) = szSeatType
            siSellInfo.aszTicketType(i) = SellTicketInfo(i).nTicketType

    Next i
    oDb.CommitTrans
    BookLocalTrans = siSellInfo
    Exit Function
Error_Handle:
    oDb.RollbackTrans
    err.Raise err.Number
End Function
Public Sub UnBookEx(paszBusID() As String, paszBusDate() As Date, paszSellTicket() As TSellTicketParam)
    
    Dim srTemp() As String
    Dim iLen As Integer
    Dim iLen1 As Integer
    Dim nCount As Integer
    Dim aszDesStationName() As String
    
    iLen = 0
    iLen1 = 0
'    iLen = ArrayLength(pszBusID)
'    ReDim srTemp(1 To iLen)
    If Trim(m_oCurrentSellService.SelfUser.ServerUnit) <> Trim(m_oCurrentSellService.SelfUser.UserUnitID) Then
        '如果用户单位和服务单位不一样则是售远程票否则售的是本地票
        On Error GoTo ErrorHandle
        
        m_oCurrentSellService.UnBookTK paszBusID, paszBusDate, paszSellTicket
        

    Else
        UnBookTK paszBusID, paszBusDate, paszSellTicket
    End If
'    BookTicket = srTemp
    Exit Sub
ErrorHandle:
    err.Raise err.Number
    Exit Sub
ErrorHandle2:
    '加入提示"对方车站已写入,请通知管理员在对方车站将此票删除"
    err.Raise err.Number, err.Source, cszServiceFinish & err.Description
    Exit Sub
End Sub
Public Function UnBookTK(paszBusID() As String, paszBusDate() As Date, paszSellTicket() As TSellTicketParam) As String '取消预留
    
    Dim szsql As String
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim rsTemp2 As Recordset
    Dim rsTemp3 As Recordset
    Dim i As Integer
    Dim tTSeatInfo(1 To 5) As TSeatInfoCount
    Dim nSaleSeatQuantity As Integer
    Dim szSqlSet As String
    Dim szSeatNoTemplog As String
    Dim j As Integer
    Dim rs As Recordset
    Dim adyDate_new() As Date
    Dim aszBusID_new()  As String
    Dim aszEndStationID_new() As String
    Dim atSellTKParam() As TSellTicketParam
    Dim aszSeatNO() As String
    Dim atBuyTicketInfo() As TBuyTicketInfo
    Dim atSellTicketResultInfo() As TSellTicketResult
    Dim szRs As String
    oDb.ConnectionString = GetConnectionStr("")

    On Error GoTo Error_Handle
    oDb.BeginTrans

    For i = 1 To ArrayLength(paszSellTicket(1).BuyTicketInfo)
 
        szsql = "SELECT seat_no FROM Env_bus_seat_lst WHERE bus_id='" & Trim(paszBusID(1)) & "'" _
               & " AND  bus_date='" & ToDBDate(paszBusDate(1)) & "'" _
               & "AND seat_no ='" & Trim(paszSellTicket(1).BuyTicketInfo(i).szSeatNo) & "'" _
               & " AND (Status <> " & ST_SeatReserved & " AND Status <> " & ST_SeatProjectBooked & " )"
        Set rsTemp2 = oDb.Execute(szsql)
        
        If rsTemp2.RecordCount > 0 Then
         szSeatNoTemplog = ""
         
          For j = 0 To rsTemp2.RecordCount - 1
           szSeatNoTemplog = szSeatNoTemplog & "[" & FormatDbValue(rsTemp2!seat_no) & "]"
           rsTemp2.MoveNext
          Next
    
        ShowError ERR_SeatNotBooked, "座位:" & szSeatNoTemplog
        
        End If
        szsql = "UPDATE Env_bus_seat_lst SET status=" & Str(ST_SeatCanSell) & " ," _
                & " remark='' " _
                & "    WHERE bus_id='" & Trim(paszBusID(1)) & "' AND  bus_date='" & ToDBDate(Trim(paszBusDate(1))) & "' " _
                & " AND seat_no  ='" & Trim(paszSellTicket(1).BuyTicketInfo(i).szSeatNo) & "'" _
                & " AND (Status = " & ST_SeatReserved & " or Status = " & ST_SeatProjectBooked & " )"
        oDb.Execute szsql
    

        
        szsql = " select seat_type_id ,count(*)as CountSeat from Env_bus_seat_lst " _
                & " where  bus_id='" & Trim(paszBusID(1)) & "'and  status='" & ST_SeatCanSell & "' and bus_date='" & ToDBDate(Trim(paszBusDate(1))) & "' group by seat_type_id "
        Set rsTemp2 = oDb.Execute(szsql)
        
        For j = 1 To rsTemp2.RecordCount
            tTSeatInfo(CInt(rsTemp2!seat_type_id)).SeatCount = FormatDbValue(rsTemp2!CountSeat)
            rsTemp2.MoveNext
        Next
        szSqlSet = ""
        nSaleSeatQuantity = 0
        szSqlSet = szSqlSet & "seat_remain ='" & tTSeatInfo(1).SeatCount & "',"
        szSqlSet = szSqlSet & "bed_remain ='" & tTSeatInfo(2).SeatCount & "',"
        szSqlSet = szSqlSet & "additional_remain ='" & tTSeatInfo(3).SeatCount & "',"
        szSqlSet = szSqlSet & "other_remain_1 ='" & tTSeatInfo(4).SeatCount & "',"
        szSqlSet = szSqlSet & "other_remain_2 ='" & tTSeatInfo(5).SeatCount & "',"
        
        For j = 1 To 5
          nSaleSeatQuantity = nSaleSeatQuantity + tTSeatInfo(j).SeatCount
        Next
        
        szSqlSet = szSqlSet & "sale_seat_quantity='" & nSaleSeatQuantity & "'"
        szsql = "UPDATE Env_bus_info SET  " & szSqlSet _
                & "  WHERE bus_id='" & Trim(paszBusID(1)) & "' AND  bus_date='" & ToDBDate(Trim(paszBusDate(1))) & "'"
        oDb.Execute (szsql)



    Next i
    oDb.CommitTrans

    Exit Function
Error_Handle:
    oDb.RollbackTrans
    err.Raise err.Number
End Function


Public Function SelfSellTK(BusStartupDate() As Date, BusID() As String, SellStationID() As String, DesStationID() As String _
             , DesStationName() As String, aspSellTicketParam() As TSellTicketParam, panInsurance() As Integer, bScroll As Boolean) As TSellTicketResult()
    '自助售票
    Dim szsql As String
    Dim oDb As New RTConnection
    Dim rsTemp2 As Recordset
    Dim rsTemp3 As Recordset
    Dim i As Integer
    Dim tTSeatInfo(1 To 5) As TSeatInfoCount
    Dim nSaleSeatQuantity As Integer
    Dim szSqlSet As String
    Dim szSeatNoTemplog As String
    Dim aszSeatNO() As String
    Dim j As Integer
    Dim atSellTicketResultInfo() As TSellTicketResult

    Dim TicketID As String
    Dim szEndTicket As String
    Dim rs As Recordset

    oDb.ConnectionString = GetConnectionStr(cszSellTicket)

    On Error GoTo Error_Handle
    oDb.BeginTrans
    
    szsql = "select parameter_value from system_param_info  where parameter_name='StrTicketID'"
    Set rs = oDb.Execute(szsql)
    TicketID = Trim(rs!parameter_value)
    m_lTicketNo = ResolveTicketNo(TicketID, m_szTicketPrefix)


    If bScroll = False Then
    For i = 1 To ArrayLength(aspSellTicketParam(1).BuyTicketInfo)
        szsql = "SELECT seat_no FROM Env_bus_seat_lst WHERE bus_id='" & Trim(BusID(1)) & "'" _
               & " AND  bus_date='" & ToDBDate(Trim(BusStartupDate(1))) & "'" _
               & "AND seat_no ='" & Trim(aspSellTicketParam(1).BuyTicketInfo(i).szSeatNo) & "'" _
               & " AND (Status <> " & ST_SeatReserved & " AND Status <> " & ST_SeatProjectBooked & " )"
        Set rsTemp2 = oDb.Execute(szsql)
        
        If rsTemp2.RecordCount > 0 Then
         szSeatNoTemplog = ""
         
          For j = 0 To rsTemp2.RecordCount - 1
           szSeatNoTemplog = szSeatNoTemplog & "[" & FormatDbValue(rsTemp2!seat_no) & "]"
           rsTemp2.MoveNext
          Next
    
        ShowError ERR_SeatNotBooked, "座位:" & szSeatNoTemplog
        
        End If
        szsql = "UPDATE Env_bus_seat_lst SET status=" & Str(ST_SeatCanSell) & " ," _
                & " remark='' " _
                & "    WHERE bus_id='" & Trim(BusID(1)) & "' AND  bus_date='" & ToDBDate(Trim(BusStartupDate(1))) & "' " _
                & " AND seat_no  ='" & Trim(aspSellTicketParam(1).BuyTicketInfo(i).szSeatNo) & "'" _
                & " AND (Status = " & ST_SeatReserved & " or Status = " & ST_SeatProjectBooked & " )"
        oDb.Execute szsql
    

        
        szsql = " select seat_type_id ,count(*)as CountSeat from Env_bus_seat_lst " _
                & " where  bus_id='" & Trim(BusID(1)) & "'and  status='" & ST_SeatCanSell & "' and bus_date='" & ToDBDate(Trim(BusStartupDate(1))) & "' group by seat_type_id "
        Set rsTemp2 = oDb.Execute(szsql)
        
        For j = 1 To rsTemp2.RecordCount
            tTSeatInfo(CInt(rsTemp2!seat_type_id)).SeatCount = FormatDbValue(rsTemp2!CountSeat)
            rsTemp2.MoveNext
        Next
        szSqlSet = ""
        nSaleSeatQuantity = 0
        szSqlSet = szSqlSet & "seat_remain ='" & tTSeatInfo(1).SeatCount & "',"
        szSqlSet = szSqlSet & "bed_remain ='" & tTSeatInfo(2).SeatCount & "',"
        szSqlSet = szSqlSet & "additional_remain ='" & tTSeatInfo(3).SeatCount & "',"
        szSqlSet = szSqlSet & "other_remain_1 ='" & tTSeatInfo(4).SeatCount & "',"
        szSqlSet = szSqlSet & "other_remain_2 ='" & tTSeatInfo(5).SeatCount & "',"
        
        For j = 1 To 5
          nSaleSeatQuantity = nSaleSeatQuantity + tTSeatInfo(j).SeatCount
        Next
        
        szSqlSet = szSqlSet & "sale_seat_quantity='" & nSaleSeatQuantity & "'"
        szsql = "UPDATE Env_bus_info SET  " & szSqlSet _
                & "  WHERE bus_id='" & Trim(BusID(1)) & "' AND  bus_date='" & ToDBDate(Trim(BusStartupDate(1))) & "'"
        oDb.Execute (szsql)
        ReDim Preserve aszSeatNO(1 To i + 1)
        aszSeatNO(i) = Trim(aspSellTicketParam(1).BuyTicketInfo(i).szSeatNo)

    Next i
    End If
    For i = 1 To ArrayLength(aspSellTicketParam(1).BuyTicketInfo)
        aspSellTicketParam(1).BuyTicketInfo(i).szTicketNo = GetTicketNo(1)
        szEndTicket = Trim(aspSellTicketParam(1).BuyTicketInfo(i).szTicketNo)
        m_lTicketNo = m_lTicketNo + 1
    Next i
        atSellTicketResultInfo = SelfSellLocalTicket(BusStartupDate, BusID, SellStationID, DesStationID, aspSellTicketParam, panInsurance, oDb)
        SelfSellTK = atSellTicketResultInfo
        szsql = "update system_param_info set parameter_value='" & szEndTicket & "' where parameter_name='StrTicketID'"
        oDb.Execute szsql
    oDb.CommitTrans
    Exit Function
Error_Handle:
    oDb.RollbackTrans
    err.Raise err.Number
End Function








Public Function RePrintNetTK(paszTicketID() As String, pszGetTicketID As String, pszValiDate As String, TicketID As String) As TSellTicketResult()
    '重打网上票
    Dim srTemp() As String
    Dim iLen As Integer
    Dim iLen1 As Integer
    Dim nCount As Integer
    Dim aszDesStationName() As String
    Dim aszResult() As TSellTicketResult
    iLen = 0
    iLen1 = 0

    If Trim(m_oCurrentSellService.SelfUser.ServerUnit) <> Trim(m_oCurrentSellService.SelfUser.UserUnitID) Then
        '如果用户单位和服务单位不一样则是售远程票否则售的是本地票
        On Error GoTo ErrorHandle
        
        RePrintNetTK = m_oCurrentSellService.RePrintNetTKTrans(paszTicketID, pszGetTicketID, pszValiDate, TicketID)
        

    Else
        RePrintNetTK = RePrintNetTKTrans(paszTicketID, pszGetTicketID, pszValiDate, TicketID)
    End If
    Exit Function
ErrorHandle:
    err.Raise err.Number
    Exit Function
ErrorHandle2:
    '加入提示"对方车站已写入,请通知管理员在对方车站将此票删除"
    err.Raise err.Number, err.Source, cszServiceFinish & err.Description
    Exit Function
End Function

Public Function RePrintNetTKTrans(paszTicketID() As String, pszGetTicketID As String, pszValiDate As String, TicketID As String) As TSellTicketResult()
    '重打网上售票
    Dim oDb As New RTConnection
    Dim szsql As String
    Dim i As Integer
    Dim j As Integer
    Dim nLen As Integer
    Dim szRs As String
    Dim rs As Recordset
    Dim atResult() As TSellTicketResult
    Dim rsTemp As Recordset
    Dim nCount As Integer
    Dim iLen As Integer
    Dim rs2 As Recordset
    Dim szTemp As String
    Dim oParam As New SystemParam
    oParam.Init m_oActiveUser
    nLen = 0
    nLen = ArrayLength(paszTicketID)
    AssertActiveUserValid m_oActiveUser, ERR_SellTicketService
    AssertHaveRight m_oActiveUser, RIGHT_CancelTicket

    szTemp = ""
    ReDim atResult(1 To nLen)
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    oDb.BeginTrans
'****************************************改网上售票表状态,3为废网上票
        szRs = "SELECT * FROM internet_sell_ticket_info WHERE getticket_id='" & pszGetTicketID & "' AND validate_id='" & pszValiDate & "'"
        Set rs = oDb.Execute(szRs)

        szRs = "SELECT * FROM internet_sell_ticket_info WHERE getticket_id='" & pszGetTicketID & "' AND validate_id='" & pszValiDate & "'" _
                & " AND if_print=2 "
        Set rs2 = oDb.Execute(szRs)
        
        If rs.RecordCount = rs2.RecordCount Then ShowError ERR_NetBookCancel  '座位已退
        For i = 1 To ArrayLength(paszTicketID)
            szTemp = szTemp & "," & TransFieldValueToString(paszTicketID(i))
        Next i
        szTemp = " ticket_id in (" & Mid(szTemp, 2, Len(szTemp) - 1) & ")"
        szsql = "SELECT * FROM internet_sell_ticket_info WHERE getticket_id='" & pszGetTicketID & "' AND validate_id='" & pszValiDate & "'" _
                & " AND if_print=1  AND " & szTemp
        Set rsTemp = oDb.Execute(szsql)
        If rsTemp.RecordCount = 0 Then ShowError ERR_NoGetNet
        For j = 1 To rsTemp.RecordCount
            szsql = " UPDATE internet_sell_ticket_info set if_print=3 WHERE bus_id='" & Trim(rsTemp!bus_id) & "' AND bus_date='" & ToDBDate(Trim(rsTemp!bus_date)) & "'" _
                & " AND if_print=1  AND getticket_id='" & pszGetTicketID & "' AND validate_id='" & pszValiDate & "'  AND seat_no='" & Trim(rsTemp!seat_no) & "'"
            oDb.Execute szsql
            rsTemp.MoveNext
            
        Next j
'*********************************************
        On Error GoTo Here
        SelfCancelTicket paszTicketID, oDb, TP_NormalCancel '真正的废票
        ReInsertNetTK rsTemp, oDb '插入新的一条记录
        
    Dim nGetNum As Integer
    Dim szSqlSet As String
    Dim szSeatNoTemplog As String
    Dim adyDate_new() As Date
    Dim aszBusID_new()  As String
    Dim aszEndStationID_new() As String
    Dim atSellTKParam() As TSellTicketParam
    Dim aszSeatNO() As String
    Dim atBuyTicketInfo() As TBuyTicketInfo
    Dim atSellTicketResultInfo() As TSellTicketResult
    Dim SellStationIDRX(1 To 1) As String
    ReDim atBuyTicketInfo(1 To nLen)
    ReDim aszBusID_new(1 To 1)
    ReDim adyDate_new(1 To 1)
    ReDim aszEndStationID_new(1 To 1)
    ReDim atSellTKParam(1 To 1)
    Dim anInsurrance(1 To 1) As Integer
    Dim aszSellStationID() As String
    Dim aszUserID(1 To 1) As String
    Dim aszBankTime(1 To 1) As Date
                nGetNum = rsTemp.RecordCount


                
                rsTemp.MoveFirst
                
                SellStationIDRX(1) = Trim(rsTemp!sell_station_id)
        
                aszBusID_new(1) = Trim(rsTemp!bus_id)
                adyDate_new(1) = Trim(rsTemp!bus_date)
                aszEndStationID_new(1) = Trim(rsTemp!des_station_id)
                aszUserID(1) = Trim(rsTemp!bank_id)
                aszBankTime(1) = FormatDbValue(rsTemp!busy_date)
                ReDim aszResultSeatID(1 To nGetNum)
                Dim asgTicketPrices() As Single
                ReDim asgTicketPrices(1 To nGetNum)
            
                For i = 1 To nGetNum
                
                    atBuyTicketInfo(i).nTicketType = Trim(rsTemp!ticket_type)
                    atBuyTicketInfo(i).szReserved = ""
                    atBuyTicketInfo(i).szSeatNo = Trim(rsTemp!seat_no)

                    Dim szTmp As String
                    szTmp = TicketID + i
                    atBuyTicketInfo(i).szTicketNo = Left(TicketID, Len(TicketID) - Len(szTmp)) & TicketID + i - 1
                    atBuyTicketInfo(i).szSeatTypeID = "01" '座位类型固定为01
                    asgTicketPrices(i) = 0
'                    atBuyTicketInfo(i).szInvoiceID = FormatDbValue(rsTemp!has_child)
                    rsTemp.MoveNext
                Next i
                atSellTKParam(1).BuyTicketInfo = atBuyTicketInfo
                atSellTKParam(1).pasgSellTicketPrice = asgTicketPrices
                anInsurrance(1) = 0
                atSellTicketResultInfo = SelfSellLocalInterNetTicket(adyDate_new, aszBusID_new, SellStationIDRX, aszEndStationID_new, atSellTKParam, anInsurrance, aszUserID, aszBankTime, oDb)
                RePrintNetTKTrans = atSellTicketResultInfo
            Dim ssql As String
            rsTemp.MoveFirst
            For i = 1 To nGetNum '把票号返写回网上售票表.
                ssql = "UPDATE internet_sell_ticket_info SET if_print=1  , print_user_id ='" & m_oActiveUser.UserID & "', ticket_id='" & atBuyTicketInfo(i).szTicketNo & "',print_time='" & oParam.NowDateTime & "' WHERE bus_id='" & Trim(rsTemp!bus_id) & "'" _
                    & " AND bank_card_id='" & Trim(rsTemp!bank_card_id) & "' AND bus_date='" & Trim(rsTemp!bus_date) & "'  AND seat_no='" & Trim(rsTemp!seat_no) & "' AND getticket_id='" & Trim(rsTemp!getticket_id) & "'" _
                    & " AND validate_id='" & Trim(rsTemp!validate_id) & "' AND if_print=" & cszNOPrint
                oDb.Execute ssql
                rsTemp.MoveNext
            Next i
    oDb.CommitTrans
    
    On Error GoTo 0
    For i = 1 To nLen
    WriteOperateLog m_oActiveUser, RIGHT_CancelTicket, "废票票号：" & paszTicketID(i)
    Next i


    Exit Function
Here:
    oDb.RollbackTrans
    err.Raise err.Number
End Function

Private Sub ReInsertNetTK(rs As Recordset, oDb As RTConnection)
    Dim szsql As String
    Dim i As Integer
    rs.MoveFirst
    For i = 1 To rs.RecordCount
        szsql = "INSERT INTO internet_sell_ticket_info(" _
        & " bank_card_id," _
        & " busy_date," _
        & " bus_date," _
        & " bus_id," _
        & " seat_no," _
        & " sell_station_id," _
        & " des_station_id," _
        & " ticket_type," _
        & " bank_id," _
        & " center_scroll_id," _
        & " getticket_id," _
        & " validate_id," _
        & " card_id," _
        & " pay_count, " _
        & " passenger, " _
        & " id_type, " _
        & " phone, " _
        & " sell_mode, " _
        & " has_child, " _
        & " handle_charge, " _
        & " ticket_owner, " _
        & " if_print )  " _

        
        szsql = szsql & " VALUES('" _
        & FormatDbValue(rs!bank_card_id) & "','" _
        & FormatDbValue(rs!busy_date) & "','" _
        & FormatDbValue(rs!bus_date) & "','" _
        & FormatDbValue(rs!bus_id) & "','" _
        & FormatDbValue(rs!seat_no) & "','" _
        & FormatDbValue(rs!sell_station_id) & "','" _
        & FormatDbValue(rs!des_station_id) & "','" _
        & FormatDbValue(rs!ticket_type) & "','" _
        & FormatDbValue(rs!bank_id) & "','" _
        & FormatDbValue(rs!center_scroll_id) & "','" _
        & FormatDbValue(rs!getticket_id) & "','" _
        & FormatDbValue(rs!validate_id) & "','" _
        & FormatDbValue(rs!card_id) & "','" _
        & FormatDbValue(rs!pay_count) & "','" _
        & FormatDbValue(rs!passenger) & "','" _
        & FormatDbValue(rs!id_type) & "','" _
        & FormatDbValue(rs!phone) & "','" _
        & FormatDbValue(rs!sell_mode) & "','" _
        & IIf(FormatDbValue(rs!has_child) = True, 1, 0) & "','" _
        & Val(FormatDbValue(rs!handle_charge)) & "','" _
        & FormatDbValue(rs!ticket_owner) & "','" _
        & cszNOPrint & "' ) "
    oDb.Execute (szsql)
        rs.MoveNext
    Next i
End Sub


Public Function NetTK(ByVal pszGetTicketID As String, ByVal pszValiDate As String, paszCanTKID() As String, Optional pszCardID As String, Optional pdtBusDate As Date) As Recordset
     If Trim(m_oCurrentSellService.SelfUser.ServerUnit) <> Trim(m_oCurrentSellService.SelfUser.UserUnitID) Then
        '如果用户单位和服务单位不一样则是售远程票否则售的是本地票
       Set NetTK = m_oCurrentSellService.NetTKTrans(pszGetTicketID, pszValiDate, paszCanTKID)
    Else
       Set NetTK = NetTKTrans(pszGetTicketID, pszValiDate, paszCanTKID, pszCardID, pdtBusDate)
    End If
    
    
    
End Function
'验证网上售票
Public Function NetTKTrans(pszGetTicketID As String, pszValiDate As String, paszTicketID() As String, Optional pszCardID As String, Optional pdtBusDate As Date) As Recordset

    Dim szsql As String
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim i As Integer
    Dim szTemp As String
On Error GoTo Error_Handle
    oDb.ConnectionString = GetConnectionStr("")
    For i = 1 To ArrayLength(paszTicketID)
        szTemp = szTemp & "," & TransFieldValueToString(paszTicketID(i))
    Next i
    szTemp = " ticket_id in (" & Mid(szTemp, 2, Len(szTemp) - 1) & ")"
    szsql = "SELECT * FROM internet_sell_ticket_info WHERE (if_print=1)"
            
    If Trim(pszGetTicketID) <> "" Then
        szsql = szsql & " AND getticket_id='" & pszGetTicketID & "' "
    End If
    
    If Trim(pszValiDate) <> "" Then
        szsql = szsql & " AND validate_id='" & pszValiDate & "'"
    End If

    If Trim(pszCardID) <> "" Then
        szsql = szsql & " AND card_id='" & pszCardID & "'"
    End If

    If pdtBusDate <> cszEmptyDate Then
        szsql = szsql & " AND bus_date='" & ToDBDate(pdtBusDate) & "'"
    End If
    
    szsql = szsql & " AND " & szTemp

    Set rsTemp = oDb.Execute(szsql)
    Set NetTKTrans = rsTemp
    Exit Function
Error_Handle:
    err.Raise err.Number
End Function
'网上售票统计
Public Function StatByInterNetSellCount(pdtStrDate As Date, pdtEndDate As Date) As Recordset
On Error GoTo ErrHandle
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim szsql As String
    Dim i As Integer
    Dim j As Integer
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)

    szsql = "select busy_date,isnull(sum(case when (if_print=1 ) then cast(pay_count as money )end),0) as SellMoney," _
            & "isnull(count(case when (if_print=1 ) then id end),0) as SellCount," _
            & "isnull(sum(case when (if_print=0) then cast(pay_count as money ) end),0) as BookMoney," _
            & "isnull(count(case when (if_print=0) then id end),0) as BookCount," _
            & "isnull(sum(case when (if_print=2) then cast(pay_count as money )end),0) as ReturnMoney," _
            & "isnull(count(case when (if_print=2) then id end),0) as ReturnCount," _
            & "isnull(sum(case when (if_print=3) then cast(pay_count as money )end),0) as CancelMoney," _
            & "isnull(count(case when (if_print=3) then id end),0) as CancelCount," _
            & "isnull(sum(cast(pay_count as money )) ,0) as TotalMoney," _
            & "isnull(count(*),0) TotalCount"
    szsql = szsql + " from internet_sell_ticket_info" _
            & " WHERE  busy_date>='" & ToDBDate(pdtStrDate) & "'" _
            & " AND  busy_date<'" & ToDBDate(pdtEndDate + 1) & "' group by busy_date order by busy_date"
    Set rsTemp = oDb.Execute(szsql)
    Set StatByInterNetSellCount = rsTemp
    Set oDb = Nothing
    Exit Function
ErrHandle:
    err.Raise err.Number
End Function


'网上售票明细
Public Function StatByInterNetSellLst(pdtStrDate As Date, pdtEndDate As Date) As Recordset
On Error GoTo ErrHandle
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim szsql As String
    oDb.ConnectionString = GetConnectionStr(cszBaseInfo)
    
    szsql = "select busy_date,bank_card_id,bus_id,bus_date,seat_no,cast(pay_count as money) as ticket_price,ticket_id,getticket_id,validate_id,ticket_type_name as ticket_type,RTRIM(print_user_id) as print_user_id,print_time," _
        & "(case when if_print=0 then '订票' when if_print=1 then '已取票' when if_print=2 then '退票' when if_print=3 then '作废' end ) as status" _
        & " from ticket_type_code,internet_sell_ticket_info where ticket_type_code.ticket_type_id=internet_sell_ticket_info.ticket_type"
    
    szsql = szsql + " AND busy_date>='" & ToDBDate(pdtStrDate) & "'" _
            & " AND  busy_date<'" & ToDBDate(pdtEndDate + 1) & "' order by busy_date"
    
    Set rsTemp = oDb.Execute(szsql)
    Set StatByInterNetSellLst = rsTemp
    Set oDb = Nothing
    Exit Function
ErrHandle:
    err.Raise err.Number
End Function

Public Function ResolveTicketNo(pszFullTicketNo, ByRef pszTicketPrefix As String) As Long
    Dim i As Integer, j As Integer
    Dim nCount As Integer, nTemp As Integer, nTicketPrefixLen As Integer
    'On Error Resume Next
    pszFullTicketNo = Trim(pszFullTicketNo)
    nCount = Len(pszFullTicketNo)
    m_lTicketNoNumLen = m_oParam.TicketNumberLen
    For i = 1 To nCount
        nTemp = Asc(Mid(pszFullTicketNo, nCount - i + 1, 1))
        If nTemp < vbKey0 Or nTemp > vbKey9 Then
            Exit For
        End If
    Next
    i = i - 1
    If i > 0 Then
        nTemp = m_lTicketNoNumLen
        nTemp = IIf(nTemp > i, i, nTemp)
        ResolveTicketNo = CLng(Right(pszFullTicketNo, nTemp))
        
        nTicketPrefixLen = m_oParam.TicketPrefixLen
        If nTicketPrefixLen <= Len(pszFullTicketNo) Then
            pszTicketPrefix = Left(pszFullTicketNo, nTicketPrefixLen)
        Else
            pszTicketPrefix = pszFullTicketNo
        End If
        
    Else
        pszTicketPrefix = ""
        ResolveTicketNo = 0
    End If
    
    
    
End Function

Public Function InvaildTk(Seller As String) As Boolean '验证票号是否相等
    Dim szsql As String
    Dim oDb As New RTConnection
    Dim TicketID As String
    Dim rs As Recordset
    Dim GetLastTicketNo As String
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)


    oDb.BeginTrans
    
    szsql = "select parameter_value from system_param_info  where parameter_name='StrTicketID'"
    Set rs = oDb.Execute(szsql)
    TicketID = FormatDbValue(rs!parameter_value)
    
    
    szsql = "SELECT MAX(ticket_id) AS ticket_no FROM ticket_sell_lst WHERE " _
            & " user_id='" & Seller & "' "
            Set rs = oDb.Execute(szsql)
    GetLastTicketNo = FormatDbValue(rs!ticket_no)
    

    If TicketID = GetLastTicketNo Then
        InvaildTk = True
    Else
        InvaildTk = False
    End If
    oDb.CommitTrans
    
End Function

'内部用废票
Private Sub SelfCancelTicket(paszTicketID() As String, poDb As RTConnection, pnCancelType As ECancelTicketType)
    Dim szsql As String
    Dim oParam As New SystemParam
    Dim szUserID As String
    Dim nLen As Integer
    Dim i As Integer
    Dim rsTemp1 As Recordset
    nLen = 0
    nLen = ArrayLength(paszTicketID)
    On Error GoTo Here
    For i = 1 To nLen
        szsql = "SELECT s.*, u.sell_station_id seller_sell_station_id FROM Ticket_sell_lst s , user_info u " _
            & " WHERE s.user_id = u.user_id AND ticket_id='" & paszTicketID(i) & "'"
            
        Set rsTemp1 = poDb.Execute(szsql)
        If rsTemp1.RecordCount <> 1 Then ShowError ERR_TicketNotExist, , "票号为:" & paszTicketID(i)
        
        If rsTemp1!Status = ST_TicketSellChange Then ShowError ERR_TicketIsChangeTicket, , "票号为:" & paszTicketID(i)
        
        szsql = "UPDATE Ticket_sell_lst SET " _
        & " status=status | " & ST_TicketCanceled & " WHERE " _
        & " ticket_id='" & paszTicketID(i) & "'"
        poDb.Execute szsql
        
        szsql = "UPDATE Env_bus_allot_lst SET sell_quantity=sell_quantity-1 WHERE bus_id='" & rsTemp1!bus_id & "'" _
              & " AND sell_station_id='" & rsTemp1!sell_station_id & "' AND bus_date='" & rsTemp1!bus_date & "'"
        poDb.Execute szsql
        
        If m_bSellStationCanSellEachOther Then
            szsql = "UPDATE Env_bus_sell_station_lst SET sell_quantity=sell_quantity-1 WHERE bus_id='" & rsTemp1!bus_id & "'" _
                  & " AND sell_station_id='" & rsTemp1!seller_sell_station_id & "' AND bus_date='" & rsTemp1!bus_date & "'"
            poDb.Execute szsql
        End If
        
        oParam.Init m_oActiveUser
        If pnCancelType = TP_ForceCancel Then
            Dim rsTemp As Recordset
            szsql = "SELECT user_id FROM Ticket_sell_lst WHERE ticket_id='" & paszTicketID(i) & "'"
            Set rsTemp = poDb.Execute(szsql)
           ' szUserID = FormatDbValue(rsTemp!user_id)
        Else
           
        End If
        
        szUserID = m_oActiveUser.UserID
        
        szsql = "INSERT INTO ticket_cancel_lst( " _
        & " ticket_id," _
        & " cancel_ticket_time," _
        & " user_id," _
        & " cancel_mode ) VALUES('" _
        & paszTicketID(i) & "','" _
        & ToDBDateTime(oParam.NowDateTime) & "','" _
        & szUserID & "'," _
        & pnCancelType & ")"
        
        poDb.Execute szsql
        
        
'         szSql = "SELECT service_unit_id,status FROM Agent_ticket_sale_lst WHERE ticket_id='" & paszTicketID(i) & "'"
'         Set rsTemp = poDb.Execute(szSql)
'         If rsTemp.RecordCount = 1 Then
'            If rsTemp!Status = ST_TicketSellChange Then ShowError ERR_TicketIsChangeTicket, , "票号为:" & paszTicketID(i)
'
'
'
'            szSql = "UPDATE Agent_ticket_sale_lst SET status=status| " & ST_TicketCanceled _
'                  & " WHERE ticket_id= '" & paszTicketID(i) & "'"
'            poDb.Execute szSql
'            '写入代理废票明细表
'            szSql = "INSERT INTO Agent_ticket_cancel_lst (" _
'            & "ticket_id," _
'            & "cancel_time," _
'            & "user_id," _
'            & "cancel_mode) VALUES ('" _
'            & paszTicketID(i) & "','" _
'            & ToDBDateTime(oParam.NowDateTime) & "','" _
'            & m_oActiveUser.UserID & "'," _
'            & TP_NormalCancel & ")"
'            poDb.Execute (szSql)
'        Else
'            ShowError ERR_TicketNotExist, , "票号为:" & paszTicketID(i)
'        End If
        '回收座位
        RecycleSeat paszTicketID(i), poDb
    Next i
    On Error GoTo 0
    Exit Sub
Here:
    err.Raise err.Number
End Sub



