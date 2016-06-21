VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SellTicketClient"
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
Public Enum EErrSellTicketClient
    ERR_SellTicketClientNoActiveUser = ERR_SellTicketClient + ERR_NoActiveUser
    ERR_UserMustBeLocalUser = ERR_SellTicketClient + 21  '使用售票客户的用户必须是本单位用户
End Enum

Private Enum ETicketClientType
    TP_CancelTicket = 1 '废票
    TP_ReturnTicket = 2 '退票
    TP_ChangeTicket = 3 '改签
End Enum
    
'此定义售票传递参数
Public Type TSellTicketParam
    BuyTicketInfo() As TBuyTicketInfo
    pasgSellTicketPrice() As Single
    aszOrgTicket() As String    '改签用
    aszChangeSheetID() As String   '改签用
End Type

Public Type TMultiSeatTypeBus
    adBusDate() As Date
    aszBusID() As String
End Type

Private Type TUserLoginUnitInfo
    szLoginUserID As String
    szPassword As String
    szIPAddress As String
End Type

Dim m_clSellService As New Collection
Dim m_oCurrentSellService As SellTicketService
Dim m_oServerProxy As New ServerProxy

Dim m_oCurrentServiceTicket As ServiceTicket
Private m_oActiveUser As ActiveUser
Private m_oParam As New SystemParam
Private m_bSellStationCanSellEachOther As Boolean

Private oSystemAgent As SystemAgent
Private oAgentActiveUser As ActiveUser

Private m_bAllowScrollBusSaleForever As Boolean


'当前提供售票服务的单位代码
Public Property Let SellUnitCode(ByVal vData As String)
    Dim nErrCount As Integer: nErrCount = 0
    Dim szOldSellUnitCode As String
    Dim OldSellSellService As SellTicketService
    
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
    Set m_oCurrentSellService = m_oServerProxy.CreateServiceObject("STTkSvr.SellTicketService")
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
Public Function SellTicket(BusDate() As Date, BusID() As String, SellStationID() As String, DesStationID() As String, DesStationName() As String, aspSellTicketParam() As TSellTicketParam, anInsurance() As Integer, Optional aszRealNameInfo As Variant) As TSellTicketResult()
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
        
        srTemp = m_oCurrentSellService.SellTicket(BusDate, BusID, SellStationID, DesStationID, aspSellTicketParam, anInsurance, aszRealNameInfo)
        
        iLen1 = ArrayLength(srTemp)
        ReDim aszDesStationName(1 To iLen1)
        For nCount = 1 To iLen1
            aszDesStationName(nCount) = srTemp(nCount).szDesStationName
        Next nCount
        On Error GoTo 0
        On Error GoTo ErrorHandle2
        SelfSellTicket BusDate, BusID, SellStationID, DesStationID, aszDesStationName, aspSellTicketParam, srTemp, anInsurance, ST_TicketNormal, aszRealNameInfo
        On Error GoTo 0
    Else
        srTemp = SelfSellLocalTicket(BusDate, BusID, SellStationID, DesStationID, aspSellTicketParam, anInsurance, aszRealNameInfo)
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

'补票
Public Function ExtraSellTicket(BusStartupDate() As Date, BusID() As String, SellStationID() As String, DesStationID() As String, DesStationName() As String, aspSellTicketParam() As TSellTicketParam, panInsurance() As Integer, Optional aszRealNameInfo As Variant) As TSellTicketResult()
    Dim srTemp() As TSellTicketResult
    On Error GoTo ErrorHandle
    
    srTemp = LocalExtraSellTicket(BusStartupDate, BusID, SellStationID, DesStationID, aspSellTicketParam, panInsurance, aszRealNameInfo)
    ExtraSellTicket = srTemp
    Exit Function
ErrorHandle:
    err.Raise err.Number
    Exit Function
End Function


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
    '玉环限制每个售票窗口的停售时间，fpd
    Dim oReg As New CFreeReg
    oReg.Init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany
    nStopSellTime = oReg.GetSetting("SellTk", "StopSellTime", oParam.StopSellTime)
'    nStopSellTime = oParam.StopSellTime
    dtNowDateTime = oParam.NowDateTime
    
    
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    
    szSelect = "SELECT Work_env_bus_info.* ,e.check_gate_id sell_check_gate_id ,seat_type_id,route_name ,Work_env_bus_station_lst.mileage end_station_mileage" _
        & " ,Work_env_bus_station_lst.sell_station_id,sell_station_name,e.bus_start_time as busstarttime, " _
        & "         full_price, half_price,preferential_ticket1,preferential_ticket2,preferential_ticket3," _
        & "         sale_ticket_quantity,stop_sale_time , Work_env_bus_station_lst.book_count ,route_info.id_card "
        
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
            szWhere = szWhere & IIf(Not m_bAllowScrollBusSaleForever, "", " AND bus_type<>" & TP_ScrollBus) _
                & " )  OR ( " _
                & IIf(Not m_bAllowScrollBusSaleForever, "", " bus_type=" & TP_ScrollBus & " OR ") _
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
            & IIf(Not m_bAllowScrollBusSaleForever, "", " OR bus_type=" & TP_ScrollBus) _
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
            & IIf(Not m_bAllowScrollBusSaleForever, "", " AND bus_type<>" & TP_ScrollBus) _
            & " AND Work_env_bus_info.bus_date='" & ToDBDate(BusDate) & "' AND " _
            & " Work_env_bus_info.bus_id=Work_env_bus_station_lst.bus_id AND " _
            & " Work_env_bus_info.bus_date=Work_env_bus_station_lst.bus_date AND " _
            & " Work_env_bus_station_lst.station_id='" & DesStationID & "' AND " _
            & " ( sale_seat_quantity>=0 OR sale_stand_seat_quantity>0" _
            & IIf(Not m_bAllowScrollBusSaleForever, "", " OR bus_type=" & TP_ScrollBus) _
            & ")  " _
            & " AND ( ( max_sale_quantity>0 AND  sale_ticket_quantity>=0) OR  ( max_sale_quantity<0) ) " _
            & " AND Work_env_bus_info.route_id=route_info.route_id "
            'ORDER BY Work_env_bus_info.bus_start_time,Work_env_bus_info.bus_id "
        
    End If
    
    If BusID <> "" Then
           szWhere = szWhere & " AND Work_env_bus_info.bus_id='" & Trim(BusID) & "'"
    End If
    
    Dim szSellStationID As String
'    If Trim(m_oActiveUser.SellStationID) = "dmy" Then
'        szSellStationID = "yh"
'    Else
'        szSellStationID = Trim(m_oActiveUser.SellStationID)
'    End If
    szSellStationID = Trim(m_oActiveUser.SellStationID)
    
    
    If Not m_bSellStationCanSellEachOther Then
        If m_oActiveUser.SellStationID <> "" Then
            szWhere = szWhere & " AND Work_env_bus_station_lst.sell_station_id='" & szSellStationID & "'"
        End If
    Else
        szTable = szTable & " , work_env_bus_sell_station_lst t "
        szWhere = szWhere & " AND t.bus_id = Work_env_bus_info.bus_id " _
            & " AND t.bus_date = Work_env_bus_info.bus_date " _
            & " AND t.sell_station_id = '" & szSellStationID & "'   and (t.can_sale_quantity >=t.sell_quantity or t.can_sale_quantity <0) "
            
    End If
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
'改签
Public Function ChangeTicket(BusStartupDate() As Date, BusID() As String, SellStationID() As String, DesStationID() As String, SellTicketInfo() As TSellTicketParam, panInsurance() As Integer, Optional aszRealNameInfo As Variant) As TSellTicketResult()
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim szsql As String
    Dim rsTemp2 As Recordset
    Dim srTemp() As TSellTicketResult
    Dim szStationName() As String
    Dim iLen As Integer
    Dim nCount As Integer
    
    Dim i As Integer
    Dim nLen As Integer
    Dim aszOldTicketNo() As Single
    
    Dim asiTemp() As TSellTicketParam
    Dim oParam As New SystemParam
    oParam.Init m_oActiveUser
    iLen = ArrayLength(BusID)
    ReDim srTemp(1 To iLen)
    ReDim szStationName(1 To iLen)
    On Error GoTo ErrorHandle
    srTemp = m_oCurrentSellService.ChangeTicket(BusStartupDate, BusID, SellStationID, DesStationID, SellTicketInfo, panInsurance, aszRealNameInfo)
    On Error GoTo 0
    On Error GoTo ErrorHandle2
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    oDb.BeginTrans
    For nCount = 1 To iLen
         
         szsql = "SELECT station_name FROM station_info WHERE station_id='" & DesStationID(nCount) & "'"
         Set rsTemp2 = oDb.Execute(szsql)
         szStationName(nCount) = rsTemp2!station_name
    
    Next nCount
    
    SelfSellTicket BusStartupDate, BusID, SellStationID, DesStationID, szStationName, SellTicketInfo, srTemp, panInsurance, ST_TicketSellChange, aszRealNameInfo
    
    For nCount = 1 To iLen
        nLen = ArrayLength(SellTicketInfo(nCount).aszOrgTicket)
        For i = 1 To nLen
            '查询代理售票明细表中是否存在此票
            szsql = "SELECT * FROM  Agent_ticket_sale_lst WHERE ticket_id='" & SellTicketInfo(nCount).aszOrgTicket(i) & "'"
            Set rsTemp = oDb.Execute(szsql)
            
            'On Error GoTo here
                If rsTemp.RecordCount = 1 Then
                    '如果此票为本地售出的
                    
                    
                    ' 找到对应的票号
                    ' 将相应的新票信息写入代理售票明细表（agent_Ticket_sell_lst）中，
                    '票状态应为改签售出。
                    szsql = "UPDATE Agent_ticket_sale_lst " _
                    & "SET status=status | " & ST_TicketChanged & "WHERE " _
                    & "ticket_id='" & SellTicketInfo(nCount).aszOrgTicket(i) & "'"
                    oDb.Execute (szsql)
                    szsql = "INSERT INTO Agent_ticket_change_lst(" _
                    & "former_sell_station_id," _
                    & "new_ticket_id," _
                    & "change_time," _
                    & "former_ticket_id," _
                    & "former_bus_date," _
                    & "former_bus_id," _
                    & "former_start_station_id," _
                    & "former_start_station_name," _
                    & "former_des_station_id," _
                    & "former_des_station_name," _
                    & "former_ticket_type," _
                    & "former_seat_no," _
                    & "former_ticket_price," _
                    & "credence_sheet_id," _
                    & "change_mode," _
                    & "change_charge,"
                    
                    szsql = szsql & "user_id) VALUES ('" _
                    & rsTemp!sell_station_id & "','" _
                    & SellTicketInfo(nCount).BuyTicketInfo(i).szTicketNo & "','" _
                    & ToDBDateTime(oParam.NowDateTime) & "','" _
                    & SellTicketInfo(nCount).aszOrgTicket(i) & "','" _
                    & ToDBDate(rsTemp!bus_date) & "','" _
                    & rsTemp!bus_id & "','" _
                    & m_oCurrentSellService.StationID & "','" _
                    & m_oCurrentSellService.StationName & "','" _
                    & rsTemp!des_station_id & "','" _
                    & rsTemp!des_station_name & "'," _
                    & rsTemp!ticket_type & ",'" _
                    & rsTemp!seat_no & "'," _
                    & rsTemp!ticket_price & ",'" _
                    & SellTicketInfo(nCount).aszChangeSheetID(i) & "'," _
                    & 0 & "," _
                    & CSng(srTemp(nCount).szReserved) & ",'" _
                    & m_oActiveUser.UserID & "')"
                    oDb.Execute (szsql)
               
                Else
                    '如果此票不为本地售出的
                    Set m_oCurrentServiceTicket = m_oServerProxy.CreateServiceObject("STTkSvr.ServiceTicket")
                    
                    m_oCurrentServiceTicket.Identify (SellTicketInfo(nCount).aszOrgTicket(i))
                    
                    szsql = "INSERT INTO Agent_ticket_change_lst(" _
                    & "former_sell_station_id," _
                    & "new_ticket_id," _
                    & "change_time," _
                    & "former_ticket_id," _
                    & "former_bus_date," _
                    & "former_bus_id," _
                    & "former_start_station_id," _
                    & "former_start_station_name," _
                    & "former_des_station_id," _
                    & "former_des_station_name," _
                    & "former_ticket_type," _
                    & "former_seat_no," _
                    & "former_ticket_price," _
                    & "credence_sheet_id," _
                    & "change_mode," _
                    & "change_charge,"
                    
                    szsql = szsql & "user_id) VALUES ('" _
                    & SellStationID(i) & "','" _
                    & SellTicketInfo(nCount).BuyTicketInfo(i).szTicketNo & "','" _
                    & ToDBDateTime(oParam.NowDateTime) & "','" _
                    & SellTicketInfo(nCount).aszOrgTicket(i) & "','" _
                    & ToDBDate(m_oCurrentServiceTicket.REBusDate) & "','" _
                    & m_oCurrentServiceTicket.REBusID & "','" _
                    & m_oCurrentSellService.StationID & "','" _
                    & m_oCurrentSellService.StationName & "','" _
                    & m_oCurrentServiceTicket.ToStation & "','" _
                    & m_oCurrentServiceTicket.ToStationName & "'," _
                    & m_oCurrentServiceTicket.TicketType & ",'" _
                    & m_oCurrentServiceTicket.SeatNo & "'," _
                    & m_oCurrentServiceTicket.TicketPrice & ",'" _
                    & SellTicketInfo(nCount).aszChangeSheetID(i) & "'," _
                    & 0 & "," _
                    & CSng(srTemp(nCount).szReserved) & ",'" _
                    & m_oActiveUser.UserID & "')"
                    oDb.Execute (szsql)
                    'error
                End If
            Next i
        Next nCount
    oDb.CommitTrans
    ChangeTicket = srTemp
    
    Exit Function
ErrorHandle:
    err.Raise err.Number
    Exit Function
ErrorHandle2:
    '加入提示"对方车站已写入,请通知管理员在对方车站将此票删除"
    oDb.RollbackTrans
    err.Raise err.Number, err.Source, cszServiceFinish & err.Description
    Exit Function
End Function
'强行退票
Public Sub ForceReturnTicket(TicketID() As String, SheetID As String, OrgReturnCharge() As Double, pasgReturnRatio() As Double)
        
    Dim szsql As String
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim rsTemp1 As Recordset
    
    Dim szTemp As String
    Dim i As Integer
    Dim oParam As New SystemParam
    
    Dim oSellClient As SellTicketClient
    
    Dim oSellService As SellTicketService
    
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    
    szsql = "SELECT service_unit_id FROM Agent_ticket_sale_lst WHERE ticket_id='" & TicketID(1) & "'"
    Set rsTemp = oDb.Execute(szsql)
    If rsTemp.RecordCount <> 1 Then
        
        '如果代码不存在,此票不是本单位售的
        szTemp = m_oServerProxy.CurrentServiceUnit.szUnitID
        
        szsql = "SELECT unit_id FROM Ticket_sell_lst,user_info WHERE ticket_id='" & TicketID(1) & "' AND " _
                 & " Ticket_sell_lst.user_id=user_info.user_id "
        
        Set rsTemp = oDb.Execute(szsql)
        '则如果服务表不存在,说明车票不存在,则出错
        If rsTemp.RecordCount <> 1 Then ShowError ERR_TicketNotExist
        
        On Error GoTo ErrorHandle
        m_oCurrentSellService.ForceReturnTicket TicketID, SheetID, OrgReturnCharge, pasgReturnRatio
        On Error GoTo 0
        On Error GoTo ErrorHandle2
        oDb.BeginTrans
         
            For i = 1 To ArrayLength(TicketID)
            
                szsql = "SELECT * FROM Ticket_sell_lst s,station_info si WHERE ticket_id='" & TicketID(i) & "' AND " _
                          & "s.des_station_id=si.station_id "
                    Set rsTemp = oDb.Execute(szsql)
                    
'                    szSql = "UPDATE Agent_ticket_sale_lst SET status=status|" _
'                        & ST_TicketReturned & " WHERE " _
'                        & " ticket_id='" & TicketID(i) & "'"
'                    oDb.Execute (szSql)
                    '写入代理退票明细表
                    szsql = "INSERT INTO Agent_ticket_return_lst(" _
                        & " sell_station_id," _
                        & " ticket_id," & "bus_date," & "bus_id," _
                        & " service_unit_id," _
                        & "start_station_id," _
                        & "start_station_name," _
                        & " des_station_id," _
                        & " des_station_name," _
                        & " ticket_type," _
                        & " seat_number," _
                        & " ticket_price," _
                        & " operation_time," _
                        & " credence_sheet_id," _
                        & " return_charge," _
                        & " user_id,"
                        szsql = szsql & "return_mode,return_ratio)  VALUES ('" _
                        & rsTemp!sell_station_id & "','" _
                        & TicketID(i) & "','" & ToDBDate(rsTemp!bus_date) & "','" & rsTemp!bus_id & "','" _
                        & m_oCurrentSellService.StationID & "','" _
                        & m_oCurrentSellService.StationID & "','" _
                        & m_oCurrentSellService.StationName & "','" _
                        & rsTemp!des_station_id & "','" _
                        & rsTemp!station_name & "'," _
                        & rsTemp!ticket_type & ",'" _
                        & rsTemp!seat_no & "'," _
                        & rsTemp!ticket_price & ",'" _
                        & ToDBDateTime(oParam.NowDateTime) & "','" _
                        & SheetID & "'," _
                        & OrgReturnCharge(i) & ",'" _
                        & m_oActiveUser.UserID & "'," _
                        & TP_ForceReturn & "," _
                        & pasgReturnRatio(i) & ")"
                        
                    oDb.Execute (szsql)
            Next i
            
        oDb.CommitTrans
        On Error GoTo 0

        
    Else
        '如果代理表存在.
        szsql = "SELECT * FROM Ticket_sell_lst WHERE ticket_id='" & TicketID(1) & "'"
        Set rsTemp1 = oDb.Execute(szsql)
        If rsTemp1.RecordCount <> 1 Then

            On Error GoTo ErrorHandle
            Set oSellService = CreateRemoteObject(rsTemp!service_unit_id, "STTkSvr.SellTicketService")
            
            oSellService.ForceReturnTicket TicketID, SheetID, OrgReturnCharge, pasgReturnRatio
            On Error GoTo 0
            On Error GoTo ErrorHandle2
            SelfLocalForceReturn TicketID, SheetID, OrgReturnCharge, pasgReturnRatio
            On Error GoTo 0
            
        Else
        
            LocalForceReturnTicket TicketID, SheetID, OrgReturnCharge, pasgReturnRatio
            
        End If
        
    End If
    Set oDb = Nothing
    Set oSellService = Nothing
    Set oParam = Nothing
    Exit Sub
ErrorHandle:
    Set oDb = Nothing
    Set oSellService = Nothing
    Set oParam = Nothing
    err.Raise err.Number
    Exit Sub
ErrorHandle2:
    '加入提示"对方车站已写入,请通知管理员在对方车站将此票删除"
    oDb.RollbackTrans
    Set oDb = Nothing
    Set oSellService = Nothing
    Set oParam = Nothing
    err.Raise err.Number, err.Source, cszServiceFinish & err.Description
    Exit Sub
End Sub

'内部强行退票
Public Sub SelfLocalForceReturn(TicketID() As String, SheetID As String, OrgReturnCharge() As Double, pasgReturnRatio() As Double)
  Dim szsql As String
  Dim poDb As New RTConnection
  Dim rsTemp As Recordset
  Dim oParam As New SystemParam
  Dim i As Integer
  
  poDb.ConnectionString = GetConnectionStr(cszSellTicket)
  oParam.Init m_oActiveUser
  
On Error GoTo Here
  poDb.BeginTrans
    For i = 1 To ArrayLength(TicketID)
    
        szsql = "SELECT * FROM Agent_ticket_sale_lst WHERE ticket_id='" & TicketID(i) & "'"
            Set rsTemp = poDb.Execute(szsql)
            
            szsql = "UPDATE Agent_ticket_sale_lst SET status=status|" _
                & ST_TicketReturned & " WHERE " _
                & " ticket_id='" & TicketID(i) & "'"
            poDb.Execute (szsql)
                szsql = "INSERT INTO Agent_ticket_return_lst(" _
                & " sell_station_id," _
                & " ticket_id," & "bus_date," & "bus_id," _
                & " service_unit_id," _
                & "start_station_id," _
                & "start_station_name," _
                & " des_station_id," _
                & " des_station_name," _
                & " ticket_type," _
                & " seat_number," _
                & " ticket_price," _
                & " operation_time," _
                & " credence_sheet_id," _
                & " return_charge," _
                & " user_id,"
                szsql = szsql & "return_mode,return_ratio)  VALUES ('" _
                & rsTemp!sell_station_id & "','" _
                & TicketID(i) & "','" & ToDBDate(rsTemp!bus_date) & "','" & rsTemp!bus_id & "','" _
                & rsTemp!service_unit_id & "','" _
                & m_oCurrentSellService.StationID & "','" _
                & m_oCurrentSellService.StationName & "','" _
                & rsTemp!des_station_id & "','" _
                & rsTemp!des_station_name & "'," _
                & rsTemp!ticket_type & ",'" _
                & rsTemp!seat_no & "'," _
                & rsTemp!ticket_price & ",'" _
                & ToDBDateTime(oParam.NowDateTime) & "','" _
                & SheetID & "'," _
                & OrgReturnCharge(i) & ",'" _
                & m_oActiveUser.UserID & "'," _
                & TP_NormalReturn & "," _
                & pasgReturnRatio(i) & ")"
                
                poDb.Execute (szsql)
    Next i
    
    poDb.CommitTrans
    Set rsTemp = Nothing
    Set poDb = Nothing
    Set oParam = Nothing
    Exit Sub
Here:
    poDb.RollbackTrans
    
    Set rsTemp = Nothing
    Set oParam = Nothing
    Set poDb = Nothing
    
    
End Sub
'强行退票
Public Sub LocalForceReturnTicket(paszTicketID() As String, pszReturnSheet As String, pOrgReturnCharge() As Double, pasgReturnRatio() As Double)
    Dim oDb As New RTConnection
    Dim i As Integer
    Dim nLen As Integer
    Dim asgReturnCharge() As Double
    Dim asgReturnRatio() As Double
    nLen = 0
    nLen = ArrayLength(paszTicketID)
    ReDim asgReturnCharge(1 To nLen)
    ReDim asgReturnRatio(1 To nLen)
    AssertActiveUserValid m_oActiveUser, ERR_SellTicketService
    AssertHaveRight m_oActiveUser, RIGHT_ForceReturnTicket
    
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    oDb.BeginTrans
    
    For i = 1 To nLen
        AssertTicketCanCancelReturn paszTicketID(i), oDb, TP_ReturnTicket, True
        asgReturnCharge(i) = pOrgReturnCharge(i)
        asgReturnRatio(i) = pasgReturnRatio(i)
    Next i
        
    On Error GoTo Here
        SelfReturnlTicket paszTicketID, pszReturnSheet, asgReturnCharge, oDb, TP_ForceReturn, asgReturnRatio

    oDb.CommitTrans
    
    
'    WriteOperateLog m_oActiveUser, RIGHT_ForceReturnTicket, "强行退票票号:" & TicketID
    On Error GoTo 0
    Set oDb = Nothing
    Exit Sub
Here:
    oDb.RollbackTrans
    Set oDb = Nothing
    err.Raise err.Number
End Sub
'正常退票
Public Sub ReturnTicket(TicketID() As String, SheetID As String, OrgReturnCharge() As Double, ReturnRatio() As Double)
'    If Trim(m_oCurrentSellService.SelfUser.ServerUnit) <> Trim(m_oCurrentSellService.SelfUser.UserUnitID) Then
'        m_oCurrentSellService.ReturnTicket TicketID, SheetID, OrgReturnCharge, ReturnRatio
'
'    Else
'
'        LocalReturnTicket TicketID, SheetID, OrgReturnCharge, ReturnRatio
'
'    End If
        
    Dim szsql As String
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim rsTemp1 As Recordset
    
    Dim szTemp As String
    Dim i As Integer
    Dim oParam As New SystemParam
    
    Dim oSellClient As SellTicketClient
    
    Dim oSellService As SellTicketService
    
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    
    szsql = "SELECT service_unit_id FROM Agent_ticket_sale_lst WHERE ticket_id='" & TicketID(1) & "'"
    Set rsTemp = oDb.Execute(szsql)
    If rsTemp.RecordCount <> 1 Then
        
        '如果代码不存在,此票不是本单位售的
        szTemp = m_oServerProxy.CurrentServiceUnit.szUnitID
        
        szsql = "SELECT unit_id FROM Ticket_sell_lst,user_info WHERE ticket_id='" & TicketID(1) & "' AND " _
                 & " Ticket_sell_lst.user_id=user_info.user_id "
        
        Set rsTemp = oDb.Execute(szsql)
        '则如果服务表不存在,说明车票不存在,则出错
        If rsTemp.RecordCount <> 1 Then ShowError ERR_TicketNotExist
        
        On Error GoTo ErrorHandle
        m_oCurrentSellService.ReturnTicket TicketID, SheetID, OrgReturnCharge, ReturnRatio
        On Error GoTo 0
        On Error GoTo ErrorHandle2
        oDb.BeginTrans
         
            For i = 1 To ArrayLength(TicketID)
            
                szsql = "SELECT * FROM Ticket_sell_lst s,station_info si WHERE ticket_id='" & TicketID(i) & "' AND " _
                          & "s.des_station_id=si.station_id "
                    Set rsTemp = oDb.Execute(szsql)
                    
'                    szSql = "UPDATE Agent_ticket_sale_lst SET status=status|" _
'                        & ST_TicketReturned & " WHERE " _
'                        & " ticket_id='" & TicketID(i) & "'"
'                    oDb.Execute (szSql)
                    '写入代理退票明细表
                    szsql = "INSERT INTO Agent_ticket_return_lst(" _
                        & " sell_station_id," _
                        & " ticket_id," & "bus_date," & "bus_id," _
                        & " service_unit_id," _
                        & "start_station_id," _
                        & "start_station_name," _
                        & " des_station_id," _
                        & " des_station_name," _
                        & " ticket_type," _
                        & " seat_number," _
                        & " ticket_price," _
                        & " operation_time," _
                        & " credence_sheet_id," _
                        & " return_charge," _
                        & " user_id,"
                        szsql = szsql & "return_mode,return_ratio)  VALUES ('" _
                        & rsTemp!sell_station_id & "','" _
                        & TicketID(i) & "','" & ToDBDate(rsTemp!bus_date) & "','" & rsTemp!bus_id & "','" _
                        & m_oCurrentSellService.StationID & "','" _
                        & m_oCurrentSellService.StationID & "','" _
                        & m_oCurrentSellService.StationName & "','" _
                        & rsTemp!des_station_id & "','" _
                        & rsTemp!station_name & "'," _
                        & rsTemp!ticket_type & ",'" _
                        & rsTemp!seat_no & "'," _
                        & rsTemp!ticket_price & ",'" _
                        & ToDBDateTime(oParam.NowDateTime) & "','" _
                        & SheetID & "'," _
                        & OrgReturnCharge(i) & ",'" _
                        & m_oActiveUser.UserID & "'," _
                        & TP_ForceReturn & "," _
                        & ReturnRatio(i) & ")"
                        
                    oDb.Execute (szsql)
            Next i
            
        oDb.CommitTrans
        On Error GoTo 0

        
    Else
        '如果代理表存在.
        szsql = "SELECT * FROM Ticket_sell_lst WHERE ticket_id='" & TicketID(1) & "'"
        Set rsTemp1 = oDb.Execute(szsql)
        If rsTemp1.RecordCount <> 1 Then

            On Error GoTo ErrorHandle
            Set oSellService = CreateRemoteObject(rsTemp!service_unit_id, "STTkSvr.SellTicketService")
            
            oSellService.ReturnTicket TicketID, SheetID, OrgReturnCharge, ReturnRatio
            On Error GoTo 0
            On Error GoTo ErrorHandle2
            SelfLocalForceReturn TicketID, SheetID, OrgReturnCharge, ReturnRatio
            On Error GoTo 0
            
        Else
        
            LocalReturnTicket TicketID, SheetID, OrgReturnCharge, ReturnRatio
            
        End If
        
    End If
    Set oDb = Nothing
    Set oSellService = Nothing
    Set oParam = Nothing
    Exit Sub
ErrorHandle:
    Set oDb = Nothing
    Set oSellService = Nothing
    Set oParam = Nothing
    err.Raise err.Number
    Exit Sub
ErrorHandle2:
    '加入提示"对方车站已写入,请通知管理员在对方车站将此票删除"
    oDb.RollbackTrans
    Set oDb = Nothing
    Set oSellService = Nothing
    Set oParam = Nothing
    err.Raise err.Number, err.Source, cszServiceFinish & err.Description
    Exit Sub
End Sub

Public Sub LocalReturnTicket(paszTicketID() As String, pszSheetID As String, pOrgReturnCharge() As Double, pReturnRatio() As Double)
    Dim oDb As New RTConnection
    Dim nLen As Integer
    Dim i As Integer
    nLen = 0
    nLen = ArrayLength(paszTicketID)
    AssertActiveUserValid m_oActiveUser, ERR_SellTicketService
    AssertHaveRight m_oActiveUser, RIGHT_ReturnTicket
    
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    On Error GoTo Here
    oDb.BeginTrans
        For i = 1 To nLen
            AssertTicketCanCancelReturn paszTicketID(i), oDb, TP_ReturnTicket, False
        Next i
        
        SelfReturnlTicket paszTicketID, pszSheetID, pOrgReturnCharge, oDb, TP_NormalReturn, pReturnRatio
        'RecycleSeat TicketID, oDb
   oDb.CommitTrans
        
    '    On Error GoTo 0
    For i = 1 To nLen
        WriteOperateLog m_oActiveUser, RIGHT_ReturnTicket, "退票票号:" & paszTicketID(i)
    Next i
    Set oDb = Nothing
    Exit Sub
Here:
    oDb.RollbackTrans
    Set oDb = Nothing
    err.Raise err.Number
End Sub
'内部用退票方法
Private Sub SelfReturnlTicket(paszTicketID() As String, pszSheetID As String, pasgOrgReturnCharge() As Double, poDb As RTConnection, pnReturnType As EReturnTicketType, pReturnRatio() As Double)
    Dim szsql As String
    Dim oParam As New SystemParam
    Dim sgReturnCharge As Double, sgTicketPrice As Double
    Dim dtBusStartupDateTime As Date
    Dim rsTemp As Recordset
    Dim lBusType As EBusType

    Dim i As Integer
    Dim nLen As Integer
    nLen = 0
    nLen = ArrayLength(paszTicketID)
    oParam.Init m_oActiveUser
    For i = 1 To nLen
        
            
            szsql = "SELECT Work_env_bus_info.* , ticket_price,Ticket_sell_lst.sell_station_id , u.sell_station_id seller_sell_station_id FROM Ticket_sell_lst,Work_env_bus_info, user_info u WHERE " _
            & " ticket_id='" & paszTicketID(i) & "' AND " _
            & " Ticket_sell_lst.bus_id= Work_env_bus_info.bus_id AND " _
            & " Ticket_sell_lst.bus_date=Work_env_bus_info.bus_date AND Ticket_sell_lst.user_id = u.user_id  "
            
            Set rsTemp = poDb.Execute(szsql)
            sgTicketPrice = FormatDbValue(rsTemp!ticket_price)
            dtBusStartupDateTime = SelfGetFullDateTime(FormatDbValue(rsTemp!bus_date), FormatDbValue(rsTemp!bus_start_time))
            lBusType = FormatDbValue(rsTemp!bus_type)
            'If FormatDbValue(rsTemp!is_all_refundment) = 1 Then '是全额退票
               ' sgReturnCharge = 0
           ' Else
             sgReturnCharge = pasgOrgReturnCharge(i)
    '            If lBusType = TP_RegularBus Then
    '                'sgReturnCharge = Round(sgTicketPrice * oParam.GetReturnChargeRatio(dtBusStartupDateTime), 2)
    '                sgReturnCharge = oParam.GetReturnCharge(dtBusStartupDateTime, sgTicketPrice)
    '            Else
    '                sgReturnCharge = Round(oParam.ScrollBusReturnRatio * sgTicketPrice + 0.01, 1)
    '            End If
           ' End If
        If sgTicketPrice = sgReturnCharge And sgTicketPrice <> 0 Then
            err.Raise ERR_SellTicketClient, , "手续费为100%,不允许退票"
        End If
    '    If sgReturnCharge <> psgOrgReturnCharge Then ShowError ERR_OrgReturnChargeOutDated

        
        '退票费为0也应该写退票表
        'If sgReturnCharge > 0 Then
        '写入退票明细表
            szsql = "INSERT INTO Ticket_return_lst( " _
            & " ticket_id," _
            & " operation_time," _
            & " credence_sheet_id," _
            & " return_charge," _
            & " user_id," _
            & " return_ticket_mode,return_ratio ) VALUES('" _
            & paszTicketID(i) & "','" _
            & ToDBDateTime(oParam.NowDateTime) & "','" _
            & pszSheetID & "'," _
            & sgReturnCharge & ",'" _
            & m_oActiveUser.UserID & "'," _
            & pnReturnType & "," _
            & pReturnRatio(i) & ")"
            poDb.Execute szsql
            
            
            Dim rsATemp As Recordset
            szsql = "SELECT * FROM Agent_ticket_sale_lst WHERE ticket_id='" & paszTicketID(i) & "'"
            Set rsATemp = poDb.Execute(szsql)
            
        Dim rsRTemp As Recordset
        szsql = "SELECT * FROM Ticket_sell_lst WHERE bus_id='" & rsTemp!bus_id & "' AND bus_date='" & rsTemp!bus_date & "' AND seat_no='" & rsATemp!seat_no & "' AND (status='" & ST_TicketNormal & "' or status='" & ST_TicketSellChange & "')"
        Set rsRTemp = poDb.Execute(szsql)
         
        
        '更新售票明细表状态
        szsql = "UPDATE Ticket_sell_lst SET " _
        & " status=status | " & ST_TicketReturned & " WHERE " _
        & " ticket_id='" & paszTicketID(i) & "'"
        poDb.Execute szsql
        
        If rsRTemp.RecordCount = 1 Then
               
            szsql = "UPDATE Env_bus_allot_lst SET sell_quantity=sell_quantity-1 WHERE bus_id='" & rsTemp!bus_id & "'" _
                   & " AND sell_station_id='" & rsTemp!sell_station_id & "' AND bus_date='" & rsTemp!bus_date & "'"
            poDb.Execute szsql
            
            szsql = "UPDATE env_bus_sell_station_lst SET sell_quantity=sell_quantity-1 WHERE bus_id='" & rsTemp!bus_id & "'" _
                   & " AND sell_station_id='" & rsTemp!seller_sell_station_id & "' AND bus_date='" & rsTemp!bus_date & "'"
            poDb.Execute szsql
            
            
            '处理并班车次的座位
            szsql = " UPDATE env_bus_seat_lst SET status = " & ST_SeatCanSell & " , ticket_no = ''" _
                    & " FROM env_bus_seat_lst a, " _
                    & " (SELECT ebsl.*, msl.merge_bus_id FROM mergebus_split_lst msl, env_bus_seat_lst ebsl " _
                    & " WHERE msl.bus_id = ebsl.bus_id And msl.bus_date = ebsl.bus_date" _
                    & " AND msl.merge_bus_id = " & TransFieldValueToString(rsTemp!bus_id) _
                    & " AND msl.bus_date = " & TransFieldValueToString(rsTemp!bus_date) & " AND ebsl.status = " & TransFieldValueToString(ST_SeatSlitp) _
                    & " AND ebsl.ticket_no = " & TransFieldValueToString(paszTicketID(i)) & " ) b " _
                    & " WHERE a.bus_id = b.bus_id And a.bus_date = b.bus_date And a.seat_no = b.seat_no"
            poDb.Execute szsql
            
        End If
        
            '更新代理售票明细表
            szsql = "UPDATE Agent_ticket_sale_lst SET status=status|" _
                & ST_TicketReturned & " WHERE " _
                & " ticket_id='" & paszTicketID(i) & "'"
            poDb.Execute (szsql)
            '写入代理退票明细表
            szsql = "INSERT INTO Agent_ticket_return_lst(" _
                & "sell_station_id," _
                & " ticket_id," & "bus_date," & "bus_id," _
                & " service_unit_id," _
                & "start_station_id," _
                & "start_station_name," _
                & " des_station_id," _
                & " des_station_name," _
                & " ticket_type," _
                & " seat_number," _
                & " ticket_price," _
                & " operation_time," _
                & " credence_sheet_id," _
                & " return_charge," _
                & " user_id,"
                szsql = szsql & "return_mode,return_ratio)  VALUES ('" _
                & rsATemp!sell_station_id & "','" _
                & paszTicketID(i) & "','" & ToDBDate(rsATemp!bus_date) & "','" & rsATemp!bus_id & "','" _
                & rsATemp!service_unit_id & "','" _
                & m_oCurrentSellService.StationID & "','" _
                & m_oCurrentSellService.StationName & "','" _
                & rsATemp!des_station_id & "','" _
                & rsATemp!des_station_name & "'," _
                & rsATemp!ticket_type & ",'" _
                & rsATemp!seat_no & "'," _
                & rsATemp!ticket_price & ",'" _
                & ToDBDateTime(oParam.NowDateTime) & "','" _
                & pszSheetID & "'," _
                & sgReturnCharge & ",'" _
                & m_oActiveUser.UserID & "'," _
                & TP_NormalReturn & "," _
                & pReturnRatio(i) & ")"
                poDb.Execute (szsql)
           
        'End If
        
        '回收座位
        RecycleSeat paszTicketID(i), poDb
    Next i
    Set rsTemp = Nothing
    
End Sub

'强行废票
Public Sub ForceCancelTicket(ForceCancelTicket() As String)

    Dim szsql As String
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim rsTemp1 As Recordset
    
    Dim szTemp As String
    
    Dim i As Integer
    Dim oParam As New SystemParam
    
    Dim oSellClient As New SellTicketClient
    
    Dim oSellService As New SellTicketService
    
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    szsql = "SELECT * FROM Agent_ticket_sale_lst WHERE ticket_id='" & ForceCancelTicket(1) & "'"
    Set rsTemp = oDb.Execute(szsql)
    If rsTemp.RecordCount <> 1 Then
        szsql = "SELECT * FROM Ticket_sell_lst WHERE ticket_id='" & ForceCancelTicket(1) & "'"
        Set rsTemp = oDb.Execute(szsql)
        If rsTemp.RecordCount <> 1 Then ShowError ERR_TicketNotExist
        
        szTemp = m_oServerProxy.CurrentServiceUnit.szUnitID
        
        szsql = "SELECT unit_id FROM Ticket_sell_lst,user_info WHERE ticket_id='" & ForceCancelTicket(1) & "' AND " _
                 & " Ticket_sell_lst.user_id=user_info.user_id "
        
        Set rsTemp = oDb.Execute(szsql)
        If rsTemp.RecordCount <> 1 Then ShowError ERR_TicketNotExist
        
        On Error GoTo ErrorHandle
        m_oCurrentSellService.ForceCancelTicket ForceCancelTicket
        On Error GoTo 0
'        Set oSellService = CreateRemoteObject(rsTemp!unit_id, "STTkSvr.SellTicketService")
'
'        oSellService.SelfLocalCancel ForceCancelTicket, True
        oDb.BeginTrans
        On Error GoTo ErrorHandle2
             For i = 1 To ArrayLength(ForceCancelTicket)
                  szsql = "SELECT * FROM Ticket_sell_lst WHERE ticket_id='" & ForceCancelTicket(i) & "'"
                  
                  Set rsTemp = oDb.Execute(szsql)
                  If rsTemp.RecordCount = 1 Then
                     If rsTemp!Status = ST_TicketSellChange Then ShowError ERR_TicketIsChangeTicket, , "票号为:" & ForceCancelTicket(i)
                     '写入代理废票明细表
                     szsql = "INSERT INTO Agent_ticket_cancel_lst (" _
                     & "sell_station_id," _
                     & "ticket_id," _
                     & "cancel_time," _
                     & "user_id," _
                     & "cancel_mode) VALUES ('" _
                     & rsTemp!sell_station_id & "','" _
                     & ForceCancelTicket(i) & "','" _
                     & ToDBDateTime(oParam.NowDateTime) & "','" _
                     & m_oActiveUser.UserID & "',"
                     
                     szsql = szsql & "" & TP_ForceCancel & ")"
                     
                     
                     oDb.Execute (szsql)
                 Else
                     ShowError ERR_TicketNotExist, , "票号为:" & ForceCancelTicket(i)
                         
                 End If
            Next i
        On Error GoTo 0
        oDb.CommitTrans
        
        Set oDb = Nothing
        Set oSellClient = Nothing
    Else
        szsql = "SELECT * FROM Ticket_sell_lst WHERE ticket_id='" & ForceCancelTicket(1) & "'"
        Set rsTemp1 = oDb.Execute(szsql)
        If rsTemp1.RecordCount <> 1 Then
'            szTemp = m_oServerProxy.CurrentServiceUnit.szUnitID
'            m_oServerProxy.ChangeServiceUnitUserBuffer (rsTemp!service_unit_id)
'            Set oSellService = m_oServerProxy.CreateServiceObject("STTkSvr.SellTicketService")
'            m_oServerProxy.ChangeServiceUnitUserBuffer szTemp
            On Error GoTo ErrorHandle
            Set oSellService = CreateRemoteObject(rsTemp!service_unit_id, "STTkSvr.SellTicketService")
            oSellService.ForceCancelTicket ForceCancelTicket
            On Error GoTo 0
            On Error GoTo ErrorHandle2
            SelfLocalCancel ForceCancelTicket, True
            On Error GoTo 0

            
        Else
            LocalForceCancelTicket ForceCancelTicket
            
        End If
        
    End If
    Set rsTemp = Nothing
    Set rsTemp1 = Nothing
    Set oDb = Nothing
    Set oSellService = Nothing
    Set oSellClient = Nothing
    Set oParam = Nothing
    Exit Sub
ErrorHandle:
   
   Set rsTemp = Nothing
   Set rsTemp1 = Nothing
    Set oDb = Nothing
    Set oSellService = Nothing
    Set oSellClient = Nothing
    Set oParam = Nothing
    err.Raise err.Number
    Exit Sub
ErrorHandle2:
    '加入提示"对方车站已写入,请通知管理员在对方车站将此票删除"
    oDb.RollbackTrans
    Set rsTemp = Nothing
    Set rsTemp1 = Nothing
    Set oDb = Nothing
    Set oSellService = Nothing
    Set oSellClient = Nothing
    Set oParam = Nothing
    err.Raise err.Number, err.Source, cszServiceFinish & err.Description
    Exit Sub
End Sub
'强行废票
Public Sub LocalForceCancelTicket(paszTicketID() As String)
    Dim oDb As New RTConnection
    Dim szsql As String
    Dim i As Integer
    Dim nLen As Integer
    Dim iLen As Integer
    AssertActiveUserValid m_oActiveUser, ERR_SellTicketService
    AssertHaveRight m_oActiveUser, RIGHT_ForceCancelTicket
    
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    For i = 1 To nLen
        AssertTicketCanCancelReturn paszTicketID(i), oDb, TP_CancelTicket, True
    Next i
    oDb.BeginTrans
    On Error GoTo Here
    SelfCancelTicket paszTicketID, oDb, TP_ForceCancel
    oDb.CommitTrans
    On Error GoTo 0
    For i = 1 To iLen
     WriteOperateLog m_oActiveUser, RIGHT_ForceCancelTicket, "强行废票票号:" & paszTicketID(i)
    Next
    Exit Sub
Here:
    oDb.RollbackTrans
    err.Raise err.Number
End Sub
''正常废票
Public Sub CancelTicket(TicketID() As String)

       
    Dim szsql As String
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim rsTemp1 As Recordset
    
    Dim szTemp As String
    
    Dim i As Integer
    Dim oParam As New SystemParam
    
    Dim oSellClient As SellTicketClient
    
    Dim oSellService As SellTicketService
    
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    szsql = "SELECT service_unit_id FROM Agent_ticket_sale_lst WHERE ticket_id='" & TicketID(1) & "'"
    Set rsTemp = oDb.Execute(szsql)
    If rsTemp.RecordCount <> 1 Then
        '当单位不存在时,也就是本票不是本单位买的，是其他单位买的
        
        szsql = "SELECT * FROM Ticket_sell_lst WHERE ticket_id='" & TicketID(1) & "'"
        Set rsTemp = oDb.Execute(szsql)
        If rsTemp.RecordCount <> 1 Then ShowError ERR_TicketNotExist
        
        szTemp = m_oServerProxy.CurrentServiceUnit.szUnitID
        
        szsql = "SELECT unit_id FROM Ticket_sell_lst,user_info WHERE ticket_id='" & TicketID(1) & "' AND " _
        & " Ticket_sell_lst.user_id=user_info.user_id "
        
        Set rsTemp = oDb.Execute(szsql)
        If rsTemp.RecordCount <> 1 Then ShowError ERR_TicketNotExist
        On Error GoTo ErrorHandle
        
        m_oCurrentSellService.CancelTicket TicketID
        On Error GoTo 0
        '        Set oSellService = CreateRemoteObject(rsTemp!unit_id, "STTkSvr.SellTicketService")
        '
        '        oSellService.SelfLocalCancel TicketID
        oDb.BeginTrans
            On Error GoTo ErrorHandle2
            For i = 1 To ArrayLength(TicketID)
                szsql = "SELECT * FROM Ticket_sell_lst WHERE ticket_id='" & TicketID(i) & "'"
                Set rsTemp = oDb.Execute(szsql)
                If rsTemp.RecordCount = 1 Then
                    If rsTemp!Status = ST_TicketSellChange Then ShowError ERR_TicketIsChangeTicket, , "票号为:" & TicketID(i)
                    
                    
                    '修改代理售票明细表
                    szsql = "UPDATE Agent_ticket_sale_lst SET status=status| " & ST_TicketCanceled _
                          & " WHERE ticket_id= '" & TicketID(i) & "'"
                    oDb.Execute szsql
                    
                    
                    
                    
                    '写入代理废票明细表
                    szsql = "INSERT INTO Agent_ticket_cancel_lst (" _
                        & "ticket_id," _
                        & "cancel_time," _
                        & "user_id," _
                        & "cancel_mode) VALUES ('" _
                        & TicketID(i) & "','" _
                        & ToDBDateTime(oParam.NowDateTime) & "','" _
                        & m_oActiveUser.UserID & "'," _
                        & TP_NormalCancel & ")"
                    oDb.Execute (szsql)
                    
                Else
                    ShowError ERR_TicketNotExist, , "票号为:" & TicketID(i)
                
                End If
            Next i
            On Error GoTo 0
        oDb.CommitTrans
        
    Else
        szsql = "SELECT * FROM Ticket_sell_lst WHERE ticket_id='" & TicketID(1) & "'"
        Set rsTemp1 = oDb.Execute(szsql)
        If rsTemp1.RecordCount <> 1 Then
'            szTemp = m_oServerProxy.CurrentServiceUnit.szUnitID
'            m_oServerProxy.ChangeServiceUnitUserBuffer (rsTemp!service_unit_id)
'            Set oSellService = m_oServerProxy.CreateServiceObject("STTkSvr.SellTicketService")
'            m_oServerProxy.ChangeServiceUnitUserBuffer szTemp
            On Error GoTo ErrorHandle
            Set oSellService = CreateRemoteObject(rsTemp!service_unit_id, "STTkSvr.SellTicketService")
            
            oSellService.CancelTicket TicketID
            On Error GoTo 0
            On Error GoTo ErrorHandle2
            SelfLocalCancel TicketID
            On Error GoTo 0
            
        Else
            
            LocalCancelTicket TicketID
            
        End If
        
    End If
    Set oSellClient = Nothing
    Set oSellService = Nothing
    Set rsTemp = Nothing
    Set oDb = Nothing
    Set oParam = Nothing
    Exit Sub
ErrorHandle:
    
    Set oSellClient = Nothing
    Set oSellService = Nothing
    Set rsTemp = Nothing
    Set oDb = Nothing
    Set oParam = Nothing
    err.Raise err.Number
    Exit Sub
ErrorHandle2:
    '加入提示"对方车站已写入,请通知管理员在对方车站将此票删除"
    oDb.RollbackTrans
    Set oSellClient = Nothing
    Set oSellService = Nothing
    Set rsTemp = Nothing
    Set oDb = Nothing
    Set oParam = Nothing
    
    err.Raise err.Number, err.Source, cszServiceFinish & err.Description
    Exit Sub
End Sub
'内部废票
Public Sub SelfLocalCancel(TicketID() As String, Optional bFlag As Boolean = False)
    Dim szsql As String
    Dim oDb As New RTConnection
    Dim oParam As New SystemParam
    Dim i As Integer
    Dim rsTemp As Recordset
    oParam.Init m_oActiveUser
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
On Error GoTo Here
    oDb.BeginTrans
    For i = 1 To ArrayLength(TicketID)
         szsql = "SELECT service_unit_id,status FROM Agent_ticket_sale_lst WHERE ticket_id='" & TicketID(i) & "'"
         
         Set rsTemp = oDb.Execute(szsql)
         If rsTemp.RecordCount = 1 Then
            If rsTemp!Status = ST_TicketSellChange Then ShowError ERR_TicketIsChangeTicket, , "票号为:" & TicketID(i)
            
            szsql = "UPDATE Agent_ticket_sale_lst SET status=status| " & ST_TicketCanceled _
                  & " WHERE ticket_id= '" & TicketID(i) & "'"
            oDb.Execute szsql
            '写入代理废票明细表
            szsql = "INSERT INTO Agent_ticket_cancel_lst (" _
            & "ticket_id," _
            & "cancel_time," _
            & "user_id," _
            & "cancel_mode) VALUES ('" _
            & TicketID(i) & "','" _
            & ToDBDateTime(oParam.NowDateTime) & "','" _
            & m_oActiveUser.UserID & "',"
            If bFlag Then
                szsql = szsql & "" & TP_ForceCancel & ")"
            Else
                szsql = szsql & "" & TP_NormalCancel & ")"
            End If
            oDb.Execute (szsql)
        Else
            ShowError ERR_TicketNotExist, , "票号为:" & TicketID(i)
        End If
   Next i
   oDb.CommitTrans
   Set rsTemp = Nothing
   Set oParam = Nothing
   Set oDb = Nothing
   Exit Sub
Here:
  oDb.RollbackTrans
  err.Raise err.Number
  Set rsTemp = Nothing
  Set oParam = Nothing
  Set oDb = Nothing
End Sub

'正常废票
Public Sub LocalCancelTicket(paszTicketID() As String)
    Dim oDb As New RTConnection
    Dim szsql As String
    Dim i As Integer
    Dim nLen As Integer
    nLen = 0
    nLen = ArrayLength(paszTicketID)
    AssertActiveUserValid m_oActiveUser, ERR_SellTicketService
    AssertHaveRight m_oActiveUser, RIGHT_CancelTicket
    
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    oDb.BeginTrans
        For i = 1 To nLen
            AssertTicketCanCancelReturn paszTicketID(i), oDb, TP_CancelTicket, False
        Next i
    
        On Error GoTo Here
        SelfCancelTicket paszTicketID, oDb, TP_NormalCancel
    oDb.CommitTrans
    
    On Error GoTo 0
    For i = 1 To nLen
    WriteOperateLog m_oActiveUser, RIGHT_CancelTicket, "废票票号:" & paszTicketID(i)
    Next i
    Exit Sub
Here:
    oDb.RollbackTrans
    err.Raise err.Number
End Sub

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
        
        Dim rsRTemp As Recordset
        szsql = "SELECT * FROM Ticket_sell_lst WHERE bus_id='" & rsTemp1!bus_id & "' AND bus_date='" & rsTemp1!bus_date & "' AND seat_no='" & rsTemp1!seat_no & "' AND (status='" & ST_TicketNormal & "' or status='" & ST_TicketSellChange & "')"
        Set rsRTemp = poDb.Execute(szsql)
        If rsRTemp.RecordCount = 1 Then
            szsql = "UPDATE Env_bus_allot_lst SET sell_quantity=sell_quantity-1 WHERE bus_id='" & rsTemp1!bus_id & "'" _
                  & " AND sell_station_id='" & rsTemp1!sell_station_id & "' AND bus_date='" & rsTemp1!bus_date & "'"
            poDb.Execute szsql
            
            
            szsql = "UPDATE Env_bus_sell_station_lst SET sell_quantity=sell_quantity-1 WHERE bus_id='" & rsTemp1!bus_id & "'" _
                  & " AND sell_station_id='" & rsTemp1!seller_sell_station_id & "' AND bus_date='" & rsTemp1!bus_date & "'"
            poDb.Execute szsql
            
            '处理并班车次的座位
            szsql = " UPDATE env_bus_seat_lst SET status = " & ST_SeatCanSell & " , ticket_no = ''" _
                    & " FROM env_bus_seat_lst a, " _
                    & " (SELECT ebsl.*, msl.merge_bus_id FROM mergebus_split_lst msl, env_bus_seat_lst ebsl " _
                    & " WHERE msl.bus_id = ebsl.bus_id And msl.bus_date = ebsl.bus_date" _
                    & " AND msl.merge_bus_id = " & TransFieldValueToString(rsTemp1!bus_id) _
                    & " AND msl.bus_date = " & TransFieldValueToString(rsTemp1!bus_date) & " AND ebsl.status = " & TransFieldValueToString(ST_SeatSlitp) _
                    & " AND ebsl.ticket_no = " & TransFieldValueToString(paszTicketID(i)) & " ) b " _
                    & " WHERE a.bus_id = b.bus_id And a.bus_date = b.bus_date And a.seat_no = b.seat_no"
            poDb.Execute szsql
        End If
        
        szsql = "UPDATE Ticket_sell_lst SET " _
        & " status=status | " & ST_TicketCanceled & " WHERE " _
        & " ticket_id='" & paszTicketID(i) & "'"
        poDb.Execute szsql
        
        
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
        
        
         szsql = "SELECT service_unit_id,status FROM Agent_ticket_sale_lst WHERE ticket_id='" & paszTicketID(i) & "'"
         Set rsTemp = poDb.Execute(szsql)
         If rsTemp.RecordCount = 1 Then
            If rsTemp!Status = ST_TicketSellChange Then ShowError ERR_TicketIsChangeTicket, , "票号为:" & paszTicketID(i)
            

            
            szsql = "UPDATE Agent_ticket_sale_lst SET status=status| " & ST_TicketCanceled _
                  & " WHERE ticket_id= '" & paszTicketID(i) & "'"
            poDb.Execute szsql
            '写入代理废票明细表
            szsql = "INSERT INTO Agent_ticket_cancel_lst (" _
            & "ticket_id," _
            & "cancel_time," _
            & "user_id," _
            & "cancel_mode) VALUES ('" _
            & paszTicketID(i) & "','" _
            & ToDBDateTime(oParam.NowDateTime) & "','" _
            & m_oActiveUser.UserID & "'," _
            & TP_NormalCancel & ")"
            poDb.Execute (szsql)
        Else
            ShowError ERR_TicketNotExist, , "票号为:" & paszTicketID(i)
        End If
        '回收座位
        RecycleSeat paszTicketID(i), poDb
    Next i
    On Error GoTo 0
    Exit Sub
Here:
    err.Raise err.Number
End Sub
Public Sub Init(AUser As ActiveUser)

    Set SelfUser = AUser
    m_oParam.Init SelfUser
    m_bSellStationCanSellEachOther = m_oParam.SellStationCanSellEachOther
'    If m_oActiveUser.SellStationID = "km" Then
'        m_bSellStationCanSellEachOther = True
'    Else
'        m_bSellStationCanSellEachOther = False
'    End If
    
    m_bAllowScrollBusSaleForever = m_oParam.AllowScrollBusSaleForever
    
End Sub
 
'内部售票函数
Private Sub SelfSellTicket(BusDate() As Date, BusID() As String, SellStationID() As String, DesStationID() As String, DesStationName() As String, SellTicket() As TSellTicketParam, srTemp() As TSellTicketResult, panInsurance() As Integer, Optional pnTicketStatus As ETicketStatus = ST_TicketNormal, Optional aszRealNameInfo As Variant)
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
    
    Dim nStart As Integer
    For nCount = 1 To iLen
        nBuyTicketCount = ArrayLength(SellTicket(nCount).BuyTicketInfo)
        For nIndex = 1 To nBuyTicketCount
            '写售出的票的信息
            szsql = "INSERT INTO Agent_ticket_sale_lst(" _
                & " sell_station_id," _
                & " ticket_id," _
                & " bus_date," _
                & " bus_id," _
                & " card_type," _
                & " id_card_no," _
                & " person_name," _
                & " sex," _
                & " person_picture," _
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
                & aszRealNameInfo(nStart + nIndex).szCardType & "','" _
                & aszRealNameInfo(nStart + nIndex).szIDCardNo & "','" _
                & aszRealNameInfo(nStart + nIndex).szPersonName & "','" _
                & aszRealNameInfo(nStart + nIndex).szSex & "','" _
                & aszRealNameInfo(nStart + nIndex).szPersonPicture & "','" _
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
        nStart = nStart + nBuyTicketCount
    Next nCount
    oDb.CommitTrans
    Exit Sub
Error_Handle:
    oDb.RollbackTrans
    err.Raise err.Number
End Sub

'
Public Function GetServiceUnitInfo()
    'SNProxy.ServerProxy
    'm_oServerProxy.GetServiceUnitInfo
    GetServiceUnitInfo = m_oServerProxy.GetServiceUnitInfo()
End Function
'得到指定车次日期的座位详细情况
Public Function GetSeatRs(BusDate As Date, BusID As String) As Recordset
    Set GetSeatRs = m_oCurrentSellService.GetSeatRs(BusDate, BusID)
End Function


'得到指定车次日期的预定详细情况
Public Function GetBookRs(BusDate As Date, BusID As String) As Recordset
    Set GetBookRs = m_oCurrentSellService.GetBookRs(BusDate, BusID)
End Function

'得到指定售票员，指定时间内的售票情况明细,不包括改签售出的。
Public Function GetSellTicketRs(QueryDate As Variant, Optional UserID As String = "") As Recordset
    Dim oDb As New RTConnection
    Dim szsql As String
    Dim dtSellDateTime As Date
    Dim nQueryTime As Integer
    Dim oParam As New SystemParam
    'oParam.Init m_oActiveUser

   If UserID = "" Then UserID = m_oActiveUser.UserID
    AssertActiveUserValid m_oActiveUser, ERR_SellTicketService
    If UserID = m_oActiveUser.UserID Then
        AssertHaveRight m_oActiveUser, RIGHT_QuerySelfThings
    Else
        AssertHaveRight m_oActiveUser, RIGHT_QueryOtherThings
    End If
    oParam.Init m_oActiveUser
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    '12月2日修改,使用用户仅能查询某段时间的售票
    nQueryTime = oParam.QueryTime ' FormatDbValue(oDb.Execute("SELECT query_limit_time FROM user_info WHERE user_id='" & UserID & "'")!query_limit_time)
    
    If IsDate(QueryDate) Then
        '12月2日时间查询,则查询看该时间范围以内的售票,-1则为不限制
        If nQueryTime <> -1 Then
            dtSellDateTime = DateAdd("n", -nQueryTime, oParam.NowDateTime)
            dtSellDateTime = IIf(DateDiff("n", QueryDate, dtSellDateTime) >= 0, dtSellDateTime, QueryDate)
        Else
            dtSellDateTime = QueryDate
        End If
    ElseIf IsNumeric(QueryDate) Then
            If nQueryTime <> -1 Then
                QueryDate = IIf(QueryDate > nQueryTime, nQueryTime, QueryDate)
                dtSellDateTime = DateAdd("n", -QueryDate, oParam.NowDateTime)
            Else
                dtSellDateTime = DateAdd("n", -QueryDate, oParam.NowDateTime)
            End If
    Else
        ShowError ERR_InputParamTypeErr
    End If
    '
    szsql = "SELECT Agent_ticket_sale_lst.* FROM Agent_ticket_sale_lst WHERE " _
    & " user_id='" & UserID & "' AND " _
    & " operation_time>'" & ToDBDateTime(dtSellDateTime) & "' AND " _
    & " status & " & ST_TicketNormal & "<>0"
    
    Set GetSellTicketRs = oDb.Execute(szsql)
    '在代理售票明细表(agent_Ticket_sell_lst)查找是指定售票员在指定时间（或时间段）售出的，
    '且车票状态有正常售出的车票记录返回。
End Function
'得到指定售票员，指定时间内的废票情况明细
Public Function GetCancelTicketRs(QueryDate As Variant, Optional UserID As String = "") As Recordset
    Dim oDb As New RTConnection
    Dim szsql As String
    Dim dtCancelDateTime As Date
    Dim oParam As New SystemParam
    If UserID = "" Then UserID = m_oActiveUser.UserID
    AssertActiveUserValid m_oActiveUser, ERR_SellTicketService
    If UserID = m_oActiveUser.UserID Then
        AssertHaveRight m_oActiveUser, RIGHT_QuerySelfThings
    Else
        AssertHaveRight m_oActiveUser, RIGHT_QueryOtherThings
    End If
    oParam.Init m_oActiveUser
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    If IsDate(QueryDate) Then
        dtCancelDateTime = QueryDate
    ElseIf IsNumeric(QueryDate) Then
'        If QueryDate = 0 Then
'            szSql = "SELECT Agent_ticket_sale_lst.* ,station_name AS des_station_name FROM Agent_ticket_sale_lst,station_info WHERE " _
'            & " user_id='" & UserID & "' AND " _
'            & DBDateCompare(oParam.NowDate, "operation_time") & " AND " _
'            & " status & " & ST_TicketNormal & "<>0 AND " _
'            & " des_station_id=station_id"
'            'Set GetSellTicketRs = oDb.Execute(szSql)
'            Set GetCancelTicketRs = oDb.Execute(szSql)
'        Else
        
            dtCancelDateTime = DateAdd("n", -QueryDate, oParam.NowDateTime)
'        End If
    Else
        ShowError ERR_InputParamTypeErr
    End If
    
    szsql = "SELECT Agent_ticket_sale_lst.*, cancel_time ,cancel_mode,sell_station_info.sell_station_name FROM " _
    & " Agent_ticket_cancel_lst,Agent_ticket_sale_lst,sell_station_info WHERE  " _
    & " Agent_ticket_cancel_lst.user_id='" & UserID & "' AND " _
    & " cancel_time>'" & ToDBDateTime(dtCancelDateTime) & "' AND " _
    & " Agent_ticket_cancel_lst.ticket_id=Agent_ticket_sale_lst.ticket_id AND sell_station_info.sell_station_id =Agent_ticket_sale_lst.sell_station_id"

    
    Set GetCancelTicketRs = oDb.Execute(szsql)

End Function
'得到指定售票员，指定时间内的改签售票情况明细
Public Function GetChangeTicketRs(QueryDate As Variant, Optional UserID As String = "") As Recordset
    Dim oDb As New RTConnection
    Dim szsql As String
    Dim dtChangeDateTime As Date
    Dim oParam As New SystemParam
    If UserID = "" Then UserID = m_oActiveUser.UserID
    AssertActiveUserValid m_oActiveUser, ERR_SellTicketService
    If UserID = m_oActiveUser.UserID Then
        AssertHaveRight m_oActiveUser, RIGHT_QuerySelfThings
    Else
        AssertHaveRight m_oActiveUser, RIGHT_QueryOtherThings
    End If
    
    oParam.Init m_oActiveUser
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    
    If IsDate(QueryDate) Then
        dtChangeDateTime = QueryDate
    ElseIf IsNumeric(QueryDate) Then
'        If QueryDate = 0 Then
'            szSql = "SELECT Agent_ticket_sale_lst.* ,station_name AS des_station_name FROM Agent_ticket_sale_lst,station_info WHERE " _
'            & " user_id='" & UserID & "' AND " _
'            & DBDateCompare(oParam.NowDate, "operation_time") & " AND " _
'            & " status & " & ST_TicketNormal & "<>0 AND " _
'            & " des_station_id=station_id"
'            Set GetChangeTicketRs = oDb.Execute(szSql)
'
'        Else
'
            dtChangeDateTime = DateAdd("n", -QueryDate, oParam.NowDateTime)
'        End If
    Else
        ShowError ERR_InputParamTypeErr
    End If
    
    szsql = "SELECT Agent_ticket_sale_lst.* ,new_ticket_id, change_time,former_ticket_id," _
    & " former_ticket_price,change_charge , credence_sheet_id,sell_station_info.sell_station_name" _
    & " FROM Agent_ticket_change_lst,Agent_ticket_sale_lst,sell_station_info WHERE " _
    & " Agent_ticket_change_lst.user_id='" & UserID & "' AND " _
    & " change_time>'" & ToDBDateTime(dtChangeDateTime) & "' AND " _
    & " Agent_ticket_change_lst.new_ticket_id=Agent_ticket_sale_lst.ticket_id " _
    & " AND sell_station_info.sell_station_id=Agent_ticket_sale_lst.sell_station_id"

    
    Set GetChangeTicketRs = oDb.Execute(szsql)

End Function
'得到指定售票员，指定时间内的退票情况明细
Public Function GetReturnTicketRs(QueryDate As Variant, Optional UserID As String = "") As Recordset
    Dim oDb As New RTConnection
    Dim szsql As String
    Dim dtReturnDateTime As Date
    Dim oParam As New SystemParam
    If UserID = "" Then UserID = m_oActiveUser.UserID
    AssertActiveUserValid m_oActiveUser, ERR_SellTicketService
    If UserID = m_oActiveUser.UserID Then
        AssertHaveRight m_oActiveUser, RIGHT_QuerySelfThings
    Else
        AssertHaveRight m_oActiveUser, RIGHT_QueryOtherThings
    End If
    
    oParam.Init m_oActiveUser
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    
    If IsDate(QueryDate) Then
        dtReturnDateTime = QueryDate
    ElseIf IsNumeric(QueryDate) Then
        dtReturnDateTime = DateAdd("n", -QueryDate, oParam.NowDateTime)
    Else
        ShowError ERR_InputParamTypeErr
    End If
    szsql = "SELECT Agent_ticket_sale_lst.*, Agent_ticket_return_lst.operation_time AS return_time,credence_sheet_id,return_charge,return_mode," _
    & "sell_station_info.sell_station_name " _
    & "  FROM Agent_ticket_return_lst,Agent_ticket_sale_lst,sell_station_info WHERE  " _
    & " Agent_ticket_return_lst.user_id='" & UserID & "' AND " _
    & " Agent_ticket_return_lst.operation_time>'" & ToDBDateTime(dtReturnDateTime) & "' AND " _
    & " Agent_ticket_return_lst.ticket_id=Agent_ticket_sale_lst.ticket_id  " _
    & " AND sell_station_info.sell_station_id=Agent_ticket_sale_lst.sell_station_id"
    
    Set GetReturnTicketRs = oDb.Execute(szsql)
End Function

'返回票对象
Public Function GetTicket(TicketID As String) As ServiceTicket
    Dim oDb As New RTConnection
    Dim szsql As String
    Dim dtReturnDateTime As Date
    Dim oParam As New SystemParam
    Dim rsTemp As Recordset
    Dim szTemp As String
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    
    szsql = "SELECT unit_id FROM Ticket_sell_lst s,user_info u WHERE " _
    & "s.ticket_id='" & TicketID & "' AND  s.user_id=u.user_id "
    
    Set rsTemp = oDb.Execute(szsql)
    If rsTemp.RecordCount <> 1 Then
        szsql = "SELECT * FROM Agent_ticket_sale_lst WHERE ticket_id='" & TicketID & "' "
        Set rsTemp = oDb.Execute(szsql)
        
        If rsTemp.RecordCount <> 1 Then ShowError ERR_TicketNotExist
            
        szTemp = m_oServerProxy.CurrentServiceUnit.szUnitID
        m_oServerProxy.ChangeServiceUnitUserBuffer (rsTemp!service_unit_id)
        Set GetTicket = m_oServerProxy.CreateServiceObject("STTkSvr.ServiceTicket")
        m_oServerProxy.ChangeServiceUnitUserBuffer szTemp
        GetTicket.Identify TicketID
        
    Else
        szTemp = m_oServerProxy.CurrentServiceUnit.szUnitID
        m_oServerProxy.ChangeServiceUnitUserBuffer (szTemp)
        Set GetTicket = m_oServerProxy.CreateServiceObject("STTkSvr.ServiceTicket")
        
        m_oServerProxy.ChangeServiceUnitUserBuffer szTemp
        GetTicket.Identify TicketID
    End If

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
            
            szTemp = m_oServerProxy.CurrentServiceUnit.szUnitID
            m_oServerProxy.ChangeServiceUnitUserBuffer (rsTemp!unit_id)
            Set GetTicketClient = m_oServerProxy.CreateServiceObject("STTkSvr.ClientTicket")
            m_oServerProxy.ChangeServiceUnitUserBuffer szTemp
            GetTicketClient.Identify TicketID
        Else
            Set GetTicketClient = Nothing
        End If
    Else
        
        szTemp = m_oServerProxy.CurrentServiceUnit.szUnitID
        m_oServerProxy.ChangeServiceUnitUserBuffer (szTemp)
        Set GetTicketClient = m_oServerProxy.CreateServiceObject("STTkSvr.ClientTicket")

        m_oServerProxy.ChangeServiceUnitUserBuffer szTemp


        GetTicketClient.Identify TicketID
    End If
End Function


'得到当前服务单位上的服务票对象
Public Function GetServerTicketAtCurrentUnit(TicketNo As String) As ServiceTicket
    Dim szsql As String
    Dim oTicket As ServiceTicket
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim szTemp As String
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    szsql = "SELECT * FROM Ticket_sell_lst WHERE ticket_id='" & TicketNo & "' "
    Set rsTemp = oDb.Execute(szsql)
    If rsTemp.RecordCount <> 1 Then
        szsql = "SELECT service_unit_id FROM Agent_ticket_sale_lst WHERE ticket_id='" & TicketNo & "'"
        Set rsTemp = oDb.Execute(szsql)
        If rsTemp.RecordCount <> 1 Then ShowError ERR_TicketNotExist
        szTemp = m_oServerProxy.CurrentServiceUnit.szUnitID
        m_oServerProxy.ChangeServiceUnitUserBuffer (rsTemp!service_unit_id)
        Set GetServerTicketAtCurrentUnit = m_oServerProxy.CreateServiceObject("STTkSvr.ServiceTicket")
        m_oServerProxy.ChangeServiceUnitUserBuffer szTemp
        GetServerTicketAtCurrentUnit.Identify TicketNo
    Else
        Set oTicket = m_oServerProxy.CreateServiceObject("STTkSvr.ServiceTicket")
        oTicket.Identify TicketNo
        Set GetServerTicketAtCurrentUnit = oTicket
    End If
    Set oTicket = Nothing
    Set oDb = Nothing
    Set rsTemp = Nothing
End Function

'得到当前服务单位上创建指定的对象
Public Function CreateServiceObject(ByVal ObjectClassName As String) As Object
        Set CreateServiceObject = m_oServerProxy.CreateServiceObject(ObjectClassName)
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
    
    szsql = " SELECT MAX(operation_time) as operation_time FROM Agent_ticket_sale_lst WHERE " _
    & " user_id='" & Seller & "' "
    Set rsTemp = oDb.Execute(szsql)
    
    szSql2 = " SELECT MAX(operation_time) as operation_time FROM agent_ticket_return_lst WHERE " _
    & " user_id='" & Seller & "' "
    If m_oParam.IsPrintZeroValueReturnSheet = False Then
        szSql2 = szSql2 & " and return_ratio<>0 "
    End If
    Set rsTemp2 = oDb.Execute(szSql2)
    
    If rsTemp.RecordCount = 1 Or rsTemp2.RecordCount = 1 Then
        If FormatDbValue(rsTemp!operation_time) > FormatDbValue(rsTemp2!operation_time) Then
            szsql = "SELECT MAX(ticket_id) AS ticket_no FROM Agent_ticket_sale_lst WHERE " _
            & " operation_time='" & rsTemp!operation_time & "'  AND " _
            & " user_id='" & Seller & "' "
            Set rsTemp = oDb.Execute(szsql)
            GetLastTicketNo = FormatDbValue(rsTemp!ticket_no)
        Else
            szsql = "SELECT MAX(credence_sheet_id) AS ticket_no FROM agent_ticket_return_lst WHERE " _
            & " operation_time='" & rsTemp2!operation_time & "' AND " _
            & " user_id='" & Seller & "' "
            Set rsTemp = oDb.Execute(szsql)
            GetLastTicketNo = FormatDbValue(rsTemp!ticket_no)
        End If
    Else
        GetLastTicketNo = String(m_oParam.TicketNumberLen, "0")
    End If
End Function

'得到指定售票员的结束票号fpd
Public Function GetEndLastTicketNo(Seller As String) As String
    Dim szsql As String
    Dim szSql2 As String 'xxd加,防止售票员票子用完后拿来的票比原先的小,以致于不能结束票号不能更新.
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim rsTemp2 As Recordset
    Dim i As Integer
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    
    szsql = "SELECT lastticket_no AS end_ticket_no FROM getticket WHERE " _
    & " user_id='" & Seller & "'" _
    & " AND getticket_date=(" _
    & " SELECT MAX(getticket_date) FROM getticket WHERE " _
    & " isuseing='1' AND " _
    & " user_id='" & Seller & "') "
    Set rsTemp = oDb.Execute(szsql)
'*********************************************************************************
    szSql2 = "SELECT MAX(ticket_id) AS ticket_no FROM Agent_ticket_sale_lst WHERE " _
    & " user_id='" & Seller & "' "
    Set rsTemp2 = oDb.Execute(szSql2)

' ********************************************************************************
    If rsTemp.RecordCount > 0 Then
        For i = 1 To rsTemp.RecordCount
'            If UCase(FormatDbValue(rsTemp!end_ticket_no)) > UCase(FormatDbValue(rsTemp2!ticket_no)) Then
                GetEndLastTicketNo = FormatDbValue(rsTemp!end_ticket_no)
'            Else
'                rsTemp.MoveNext
'            End If
        Next i
    Else
        Dim oParam As New SystemParam
        oParam.Init m_oActiveUser
        GetEndLastTicketNo = String(oParam.TicketNumberLen, "0")
    End If
End Function

'给定车次,得到站点
Public Function GetStationFromBusRs(BusDate As Date, BusID As String) As Recordset
    Set GetStationFromBusRs = m_oCurrentSellService.GetStationFromBusRs(BusDate, BusID)
End Function

'得到当前服务单位的系统参数对象
Public Property Get ServerParamAtCurrentUnit() As SystemParam
    Set ServerParamAtCurrentUnit = m_oParam
End Property


'得到当前系统中的所有站点信息
Public Function GetAllStationRs() As Recordset
    Set GetAllStationRs = m_oCurrentSellService.GetAllStationRs()
End Function

'注销车票
Public Sub WriteOffTicket(TicketID As String, Optional bOnlyDelLocal As Boolean = False)
    Dim szsql As String
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim szNewUnitID As String
    Dim szOldUnitID As String
    
    AssertActiveUserValid m_oActiveUser, ERR_SellTicketService
    AssertHaveRight m_oActiveUser, RIGHT_WriteOffTiecket

    szsql = "SELECT * FROM Agent_ticket_sale_lst WHERE ticket_id='" & TicketID & "'"
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    
    On Error GoTo Here
    Set rsTemp = oDb.Execute(szsql)
    On Error GoTo 0
    If rsTemp.RecordCount < 1 Then ShowError ERR_NoExistWriteOffTiecket
    szNewUnitID = FormatDbValue(rsTemp!service_unit_id)
    szOldUnitID = SellUnitCode
    
'    If bOnlyDelLocal = False Then
'        On Error GoTo herehere
'        SellUnitCode = szNewUnitID
'        On Error GoTo 0
'    End If
'
'    If (szNewUnitID = szOldUnitID) Then
'    On Error GoTo here
'    m_oCurrentSellService.WriteOffTiecket (TicketID)
'    On Error GoTo 0
    If bOnlyDelLocal = True Then
        If szOldUnitID = szNewUnitID Then
            On Error GoTo Here
            m_oCurrentSellService.WriteOffTicket (TicketID)
            On Error GoTo 0
        End If
    Else
        On Error GoTo herehere
        SellUnitCode = szNewUnitID
        On Error GoTo Here
        m_oCurrentSellService.WriteOffTicket (TicketID)
        On Error GoTo 0
    End If
    
    On Error GoTo there
    oDb.BeginTrans
    DeleteAgentTicket TicketID, oDb
    oDb.CommitTrans
    On Error GoTo 0
    
    SellUnitCode = szOldUnitID
    WriteOperateLog m_oActiveUser, RIGHT_WriteOffTiecket, "注销客户端车票[:" & TicketID & "]"
Exit Sub
herehere:
    err.Raise err.Number, , err.Description
Exit Sub
Here:
    ShowError ERR_WriteOffTiecketFail
    SellUnitCode = szOldUnitID
Exit Sub
there:
    oDb.RollbackTrans
    SellUnitCode = szOldUnitID
    ShowError ERR_WriteOffTiecketFail
End Sub

'删除代理票,包括表:agent_Ticket_sell_lst,Agent_ticket_cancel_lst,Agent_ticket_change_lst,Agent_ticket_return_lst
Private Sub DeleteAgentTicket(TicketID As String, poDb As RTConnection)
    Dim szsql As String
    szsql = "DELETE Agent_ticket_cancel_lst WHERE ticket_id='" & TicketID & "'"
    poDb.Execute szsql
    szsql = "DELETE Agent_ticket_change_lst WHERE new_ticket_id='" & TicketID & "'"
    poDb.Execute szsql
    szsql = "DELETE Agent_ticket_return_lst WHERE ticket_id='" & TicketID & "'"
    poDb.Execute szsql
    szsql = "DELETE Agent_ticket_sale_lst WHERE ticket_id='" & TicketID & "'"
    poDb.Execute szsql
End Sub

'内部本地售票
Public Function SelfSellLocalTicket(BusStartupDate() As Date, BusID() As String, SellStationID() As String, DesStationID() As String _
             , aspSellTicketParam() As TSellTicketParam, panInsurance() As Integer, Optional aszRealNameInfo As Variant) As TSellTicketResult()
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
    
    Dim nStart As Integer
    For nCount = 1 To iLen
        rsTemp(nCount) = SelfSellTicketLocalTrans(BusStartupDate(nCount), BusID(nCount), SellStationID(nCount), DesStationID(nCount), aspSellTicketParam(nCount).BuyTicketInfo, aspSellTicketParam(nCount).pasgSellTicketPrice, panInsurance(nCount), False, ST_TicketNormal, oDb, nStart, aszRealNameInfo)
        nStart = nStart + ArrayLength(aspSellTicketParam(nCount).BuyTicketInfo)
    Next nCount
    SelfSellLocalTicket = rsTemp
    oDb.CommitTrans
    Set oDb = Nothing
    Exit Function
Here:
    oDb.RollbackTrans
    Set oDb = Nothing
    
        err.Raise err.Number
End Function


'内部售票事务函数
Private Function SelfSellTicketLocalTrans(BusStartupDate As Date, BusID As String, SellStationID As String, DesStationID As String, SellTicketInfo() As TBuyTicketInfo _
    , pasgSellTicketPrice() As Single, pnInsurance As Integer, Optional pbExtraSell As Boolean = False _
    , Optional pnTicketStatus As ETicketStatus = ST_TicketNormal, Optional poDb As RTConnection, Optional nRealNameStart As Integer, Optional aszRealNameInfo As Variant) As TSellTicketResult
    
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
    Dim szBookSellStationID As String '预定的上车站
    Dim rsTemp2 As Recordset
    
    Dim szSeatType As String '座位类型
    Dim szSqlSeat As String  '座位票价
    Dim nSeatStatus As Integer '座位状态
    
    Dim bIsRealName As Boolean  '是否需要实名制
    Dim rsSellCardInfo As Recordset
    Dim nCardBuyTicketInterval As Integer '同一证件购票时间间隔（单位分钟）
    Dim dtBusStartTime As Date
    Dim szWhereCard As String
    Dim szCardTemp As String
    
    Dim szStopTime As String
    
    Dim nSellTicketType As Integer
    Dim nSellTicketStatus As ETicketStatus
    
    
    Dim bSellStationCanSellEachOther As Boolean
    
    ReDim atiTicketPrice(1 To TP_TicketTypeCount)
    
    Dim m_bSpecialTicketTypePosition As String
    Dim m_bSpecialTicketTypeRatio As Double
    Dim m_bHalfTicketTypeRatio As Double
    m_bSpecialTicketTypePosition = m_oParam.SpecialTicketTypePosition
    m_bSpecialTicketTypeRatio = m_oParam.SpecialTicketTypeRatio
    m_bHalfTicketTypeRatio = m_oParam.HalfTicketTypeRatio
    
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
    
    nCardBuyTicketInterval = Val(oParam.CardBuyTicketInterval)
    
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
        & " , e.register_status , e.is_all_refundment , e.route_id , e.end_station_id , e.end_station_name" _
        & " , e.owner_id , e.transport_company_id , e.vehicle_id , e.split_company_id , e.baggage_number , e.fact_weight" _
        & " , e.calculate_weight , e.over_weight_number , e.luggage_total_price , e.internet_status , e.scrollbus_check_time" _
        & " , e.seat_remain , e.bed_remain  , e.additional_remain , e.other_remain_1 , e.other_remain_2" _
        & " , s.sell_station_id , s.station_id , s.seat_type_id , s.station_serial_no , s.max_sale_quantity , s.sale_ticket_quantity" _
        & " , s.book_count , s.stop_sale_time , s.mileage  , s.full_price , s.half_price , s.preferential_ticket1 , s.preferential_ticket2" _
        & " , s.preferential_ticket3" _
        & " , a.check_gate_id , a.bus_start_time" _
        & " , a.status , a.can_sale_quantity , a.sell_quantity ,r.id_card" _
        & " FROM work_env_bus_info e ,work_env_bus_station_lst s , work_env_bus_allot_lst a ,route_info r" _
        & " WHERE e.bus_id='" & BusID & "' AND  e.bus_date='" & ToDBDate(BusStartupDate) & "'" _
        & " AND  e.bus_id=s.bus_id AND  e.bus_date=s.bus_date AND e.route_id =r.route_id" _
        & " AND  s.station_id='" & DesStationID & "' AND s.sell_station_id='" & SellStationID & "'" _
        & " AND a.bus_id = s.bus_id AND a.bus_date = s.bus_date AND a.sell_station_id = '" & SellStationID & "'" _
        
        
    Set rsTemp = poDb.Execute(szsql)
    
    
    
    
    If rsTemp.RecordCount = 0 Then ShowError ERR_NoThisBus '没有指定的车次或车次不过指定的到站
    nBusType = FormatDbValue(rsTemp!bus_type)
    
    bIsRealName = FormatDbValue(rsTemp!id_card)
    dtBusStartTime = FormatDbValue(rsTemp!bus_start_time)
    
'******************************************************************************
'判断特种票是否售完fpd
    If nBusType <> TP_ScrollBus Then
    
        Dim nSpecialTicketTypeCount As Integer
        Dim rsSelledSpecialTicketCount As Recordset
        nSpecialTicketTypeCount = 0
        
        For i = 1 To ArrayLength(SellTicketInfo)
            If SellTicketInfo(i).nTicketType = m_bSpecialTicketTypePosition Then
                nSpecialTicketTypeCount = nSpecialTicketTypeCount + 1
            End If
        Next i
        
        If nSpecialTicketTypeCount > 0 Then
            szsql = "SELECT count(*) as selled_special_ticket_count From ticket_sell_lst " _
              & " WHERE bus_id='" & BusID & "' AND  bus_date='" & ToDBDate(BusStartupDate) & "'" _
              & " AND  status in ('" & ST_TicketNormal & "','" & ST_TicketNormal + ST_TicketChecked & "','" & ST_TicketSellChange & "','" & ST_TicketSellChange + ST_TicketChecked & "' ) AND ticket_type='" & m_bSpecialTicketTypePosition & "' "
            Set rsSelledSpecialTicketCount = poDb.Execute(szsql)
            
            If (FormatDbValue(rsSelledSpecialTicketCount!selled_special_ticket_count) + nSpecialTicketTypeCount) > Int(m_bSpecialTicketTypeRatio * FormatDbValue(rsTemp!total_seat) / 100) Then
                ShowError ERR_CanNotSellSpecialTicket
            End If
        End If
        
    End If
'******************************************************************************

'******************************************************************************
'判断半票是否售完fpd
    If nBusType <> TP_ScrollBus Then
    
        Dim nHalfTicketTypeCount As Integer
        Dim rsSelledHalfTicketCount As Recordset
        nHalfTicketTypeCount = 0
        
        For i = 1 To ArrayLength(SellTicketInfo)
            If SellTicketInfo(i).nTicketType = TP_HalfPrice Then
                nHalfTicketTypeCount = nHalfTicketTypeCount + 1
            End If
        Next i
        
        If nHalfTicketTypeCount > 0 Then
            szsql = "SELECT count(*) as selled_half_ticket_count From ticket_sell_lst " _
              & " WHERE bus_id='" & BusID & "' AND  bus_date='" & ToDBDate(BusStartupDate) & "'" _
              & " AND  status in ('" & ST_TicketNormal & "','" & ST_TicketNormal + ST_TicketChecked & "','" & ST_TicketSellChange & "','" & ST_TicketSellChange + ST_TicketChecked & "' ) AND ticket_type='" & TP_HalfPrice & "' "
            Set rsSelledHalfTicketCount = poDb.Execute(szsql)
            
            If (FormatDbValue(rsSelledHalfTicketCount!selled_half_ticket_count) + nHalfTicketTypeCount) > Int(m_bHalfTicketTypeRatio * FormatDbValue(rsTemp!total_seat) / 100) Then
                ShowError ERR_CanNotSellHalfTicket
            End If
        End If
        
    End If
'******************************************************************************
    
    
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
    If nBusStatus = ST_BusStopCheck Then ShowError ERR_BusStopCheck  'And m_oActiveUser.SellStationID <> "dmy"
    If nBusStatus = ST_BusStopped Then ShowError ERR_BusStopped
    dtBusFullDateTime = CDate(Format(FormatDbValue(rsTemp!bus_date), cszDateStr) & " " & Format(FormatDbValue(rsTemp!bus_start_time), cszTimeStr))
    '判断是否到了停售时间(如果是补票则要过了正常售票时间才能补)
    
    '玉环限制每个售票窗口的停售时间，fpd
    Dim oReg As New CFreeReg
    Dim nStopSellTime As Integer
    oReg.Init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany
    nStopSellTime = oReg.GetSetting("SellTk", "StopSellTime", oParam.StopSellTime)
    
    If nBusType <> TP_ScrollBus Then '如果是固定车次才限制
        If pbExtraSell Then
            If DateDiff("s", dtNow, dtBusFullDateTime) >= nStopSellTime * 60 Then ShowError ERR_BusNotStopSellTicket
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
            bBookValid = IIf(DateAdd("n", oParam.BookTime, dtNow) >= IIf(Trim(m_oActiveUser.SellStationID) = "km", DateAdd("n", 30, rsTemp!bus_start_time), rsTemp!bus_start_time), False, True)
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
        If nStationRemainCount >= 0 Then
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
        
''            If oParam.TicketIfSupervise Then
'                '判断车票领用
'                Set rsTemp2 = poDb.Execute("SELECT * from getticket " _
'                            & " WHERE user_id = '" & m_oActiveUser.UserID & "' " _
'                            & " AND firstticket_no <='" & SellTicketInfo(i).szTicketNo & "' " _
'                            & " AND lastticket_no >= '" & SellTicketInfo(i).szTicketNo & "' " _
'                            & " AND isuseing = 1")
'
'                If rsTemp2.RecordCount = 0 Then ShowError ERR_TicketNoDraw
''            End If
        
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
                            szsql = "SELECT s.ticket_no,p.ticket_price_total AS TicketPrice, s.seat_no AS SeatNo,s.status,p.seat_type_id AS SeatType FROM Env_bus_seat_lst s,work_env_bus_price_lst p" _
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
                            szsql = "SELECT s.ticket_no,p.ticket_price_total AS TicketPrice,s.seat_no AS SeatNo,s.seat_type_id AS SeatType,p.ticket_price_total AS TicketPrice,s.status  FROM Env_bus_seat_lst s, work_env_bus_price_lst p WHERE " _
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
                        If FormatDbValue(rsTemp!ticket_no) <> "" Then  'fpd
                            If IsForceSell = True Then
                                If MsgBox("座位号重，是否要强行出售？", vbInformation + vbYesNo + vbDefaultButton2, "座位号重") = vbNo Then
                                    ShowError ERR_SeatRepeat
                                End If
                            Else
                                ShowError ERR_SeatRepeat
                            End If
                        End If
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
                    '系统不是自动给出座位号
                        szSeatNo = SellTicketInfo(i).szSeatNo
                        
                         '取座位该票价和类型
                        szSqlSeat = "SELECT s.ticket_no,p.ticket_price_total AS TicketPrice, p.seat_type_id AS SeatType,s.status AS SeatStatus " _
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
                        If FormatDbValue(rsTemp!ticket_no) <> "" Then  'fpd
                            If IsForceSell = True Then
                                If MsgBox("座位号重，是否要强行出售？", vbInformation + vbYesNo + vbDefaultButton2, "座位号重") = vbNo Then
                                    ShowError ERR_SeatRepeat
                                End If
                            Else
                                ShowError ERR_SeatRepeat
                            End If
                        End If
                        sgTicketPrice = rsTemp!TicketPrice
                        szSeatType = FormatDbValue(rsTemp!SeatType)
                        nSeatStatus = rsTemp!seatstatus
                        
                        'fpd，堵塞漏洞：当选座位时进行预定操作
                        If nSeatStatus = ST_SeatBooked Or nSeatStatus = ST_SeatProjectBooked Then
                            If SellTicketInfo(i).szReserved <> nSeatStatus Then ShowError ERR_SeatHasBooked
                        End If
                        
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
                                ElseIf nSeatStatus = ST_SeatSold Then
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
                        szBookSellStationID = Trim(rsTemp2!sell_station_id)
                        Set rsTemp2 = Nothing

                        
                        'If nStationLimitCount > 0 Then
                            '更新站点的预定张数
                        szsql = "UPDATE Env_bus_station_lst SET " _
                        & " book_count=book_count-1 WHERE " _
                        & " bus_id='" & BusID & "' AND " _
                        & " bus_date='" & ToDBDate(BusStartupDate) & "' AND " _
                        & " station_id='" & szBookStationID & "' AND " _
                        & " sell_station_id='" & szBookSellStationID & "' AND " _
                        & " book_count>0"
'                        & " sell_station_id='" & SellStationID & "' AND " _原来是这个条件 zyw

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
                    
'                    If bSellStationCanSellEachOther Then
                        '得到环境车次售票点信息
'                        If m_oActiveUser.SellStationID <> "dmy" Then
                            Set rsTemp2 = poDb.Execute("SELECT * from work_env_bus_sell_station_lst WHERE bus_date='" _
                            & ToDBDate(BusStartupDate) & "' AND " _
                            & " bus_id='" & BusID & "' AND " _
                            & " sell_station_id='" & m_oActiveUser.SellStationID & "'")
                            If rsTemp2.RecordCount = 0 Then ShowError ERR_BusSellStationInfoNotExist
'                        End If
                        
                        '总共允许售票张数是否大于"已售张数"+1
                        
                        If rsTemp2!can_sale_quantity >= rsTemp2!sell_quantity + 1 Or rsTemp2!can_sale_quantity = -1 Then
                            poDb.Execute ("UPDATE env_bus_sell_station_lst SET sell_quantity=sell_quantity + 1 WHERE bus_date='" _
                                & ToDBDate(BusStartupDate) & "' AND " _
                                & " bus_id='" & BusID & "' AND " _
                                & " sell_station_id='" & m_oActiveUser.SellStationID & "'")
'                            poDb.Execute ("UPDATE work_env_bus_sell_station_lst SET sell_quantity=sell_quantity + 1 WHERE bus_date='" _
'                                & ToDBDate(BusStartupDate) & "' AND " _
'                                & " bus_id='" & BusID & "' AND " _
'                                & " sell_station_id='" & m_oActiveUser.SellStationID & "'")
                        Else
                            ShowError ERR_SellQuantityISFull '超过允许售票张数
                        End If
'                    End If

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
                                & " sale_ticket_quantity>book_count AND " _
                                & " station_id='" & DesStationID & "'"
'                                & " sell_station_id='" & SellStationID & "' AND "
                                poDb.Execute szsql, lRowAffect
                                
                                If lRowAffect = 0 Then ShowError ERR_BusReachMaxStationTicketCount, "SellTicketClient.SelfSellTicketLocalTrans"
                            End If
                            
                        End If
                        If nSeatStatus <> ST_SeatSold Then
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
            
            If bIsRealName = True Then
                '如果需要实名制
                
                If aszRealNameInfo(nRealNameStart + i).szCardType = "身份证" Then
                '如果证件为身份证
                    If SellTicketInfo(i).nTicketType = TP_HalfPrice Then
                    '如果是半票
                        '证件号、姓名不允许同时为空
                        If aszRealNameInfo(nRealNameStart + i).szIDCardNo = "" And aszRealNameInfo(nRealNameStart + i).szPersonName = "" Then ShowError ERR_IDCardNoAndPersonNameCanNotNullOne
                    Else
                    '不是半票
                        '证件号、姓名两者都不能为空
                        If aszRealNameInfo(nRealNameStart + i).szIDCardNo = "" Or aszRealNameInfo(nRealNameStart + i).szPersonName = "" Then ShowError ERR_IDCardNoAndPersonNameCanNotNullAll
                    End If
                Else
                '如果证件不是身份证
                    '证件号不允许为空
                    If aszRealNameInfo(nRealNameStart + i).szIDCardNo = "" Then ShowError ERR_IDCardNoCanNotNull
                End If
                
                '实名制，判断同一证件购票间隔是否超过规定时间间隔
                If aszRealNameInfo(nRealNameStart + i).szIDCardNo <> "" Then
                    szCardTemp = aszRealNameInfo(nRealNameStart + i).szIDCardNo
                    szWhereCard = " AND s.id_card_no =" & TransFieldValueToString(szCardTemp)
                Else
                    szCardTemp = aszRealNameInfo(nRealNameStart + i).szPersonName
                    szWhereCard = " AND s.person_name =" & TransFieldValueToString(szCardTemp)
                End If
                szsql = "SELECT s.bus_id ,a.bus_start_time FROM ticket_sell_lst s ,env_bus_allot_lst a WHERE" _
                    & " s.bus_date = a.bus_date And s.bus_id = a.bus_id And s.sell_station_id = a.sell_station_id" _
                    & " AND s.status IN (" & ST_TicketNormal & "," & ST_TicketNormal & "|" & ST_TicketChecked & "," & ST_TicketSellChange & "," & ST_TicketSellChange & "|" & ST_TicketChecked & ")" _
                    & " AND s.id_card_no =" & TransFieldValueToString(aszRealNameInfo(nRealNameStart + i).szIDCardNo) _
                    & " AND s.bus_date >=" & TransFieldValueToString(DateAdd("D", -1, dtNow)) _
                    & " AND ABS(DATEDIFF(s ,'" & dtBusStartTime & "',a.bus_start_time)) <=" & nCardBuyTicketInterval * 60
                Set rsSellCardInfo = poDb.Execute(szsql)
                If rsSellCardInfo.RecordCount > 0 Then
                    err.Raise 7777, , "证件[" & szCardTemp & "]" & "未超过系统规定的时间间隔[" & nCardBuyTicketInterval & "]分钟" & vbCrLf & "已购买车次[" & Trim(FormatDbValue(rsSellCardInfo!bus_id)) & "]" & ",发车时间[" & ToDBDateTime(FormatDbValue(rsSellCardInfo!bus_start_time)) & "]"
                End If
            End If
            
            '写售票明细表
            szsql = "INSERT INTO Ticket_sell_lst(" _
                & " sell_station_id," _
                & " ticket_id," _
                & " bus_date," _
                & " bus_id," _
                & " card_type," _
                & " id_card_no," _
                & " person_name," _
                & " sex," _
                & " person_picture," _
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
                & aszRealNameInfo(nRealNameStart + i).szCardType & "','" _
                & aszRealNameInfo(nRealNameStart + i).szIDCardNo & "','" _
                & aszRealNameInfo(nRealNameStart + i).szPersonName & "','" _
                & aszRealNameInfo(nRealNameStart + i).szSex & "','" _
                & aszRealNameInfo(nRealNameStart + i).szPersonPicture & "','" _
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
    
    'fpd
    If rsTemp.RecordCount <> 0 Then
        rsTemp.MoveFirst
        For i = 1 To rsTemp.RecordCount
            If FormatDbValue(rsTemp!sell_station_id) = SellStationID Then
                szStartStationID = FormatDbValue(rsTemp!station_id)
                szStartStationName = FormatDbValue(rsTemp!station_name)
            End If
            rsTemp.MoveNext
        Next i
    End If
    'fpd
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
            & " card_type," _
            & " id_card_no," _
            & " person_name," _
            & " sex," _
            & " person_picture," _
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
            & aszRealNameInfo(nRealNameStart + nIndex).szCardType & "','" _
            & aszRealNameInfo(nRealNameStart + nIndex).szIDCardNo & "','" _
            & aszRealNameInfo(nRealNameStart + nIndex).szPersonName & "','" _
            & aszRealNameInfo(nRealNameStart + nIndex).szSex & "','" _
            & aszRealNameInfo(nRealNameStart + nIndex).szPersonPicture & "','" _
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

'///////////////////////////////
'取消退票
Public Function ResumeReturnTicket(TicketID() As String) As String()
    Dim szsql As String
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim rsTemp1 As Recordset
    Dim oSellTicketService As SellTicketService
    Dim aszReturnSeatNo() As String
    
    Dim oParam As New SystemParam
    Dim i As Integer
    
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
        
    szsql = "SELECT service_unit_id FROM Agent_ticket_sale_lst WHERE ticket_id='" & TicketID(1) & "'"
    Set rsTemp = oDb.Execute(szsql)
    If rsTemp.RecordCount <> 1 Then
        szsql = "SELECT unit_id FROM Ticket_sell_lst s,user_info u WHERE " _
                & " s.ticket_id='" & TicketID(1) & "' AND s.user_id=u.user_id "
        Set rsTemp = oDb.Execute(szsql)
        If rsTemp.RecordCount = 0 Then ShowError ERR_TicketNotExist
        On Error GoTo ErrorHandle
        aszReturnSeatNo = m_oCurrentSellService.ResumeReturnTicket(TicketID)
        On Error GoTo 0
'        Set oSellTicketService = CreateRemoteObject(rsTemp!unit_id, "STTkSvr.SellTicketService")
'
'        oSellTicketService.ResumeReturnClient TicketID, aszReturnSeatNo
       oDb.BeginTrans
       On Error GoTo ErrorHandle2
              For i = 1 To ArrayLength(TicketID)
                szsql = "SELECT * FROM Agent_ticket_return_lst WHERE ticket_id='" & TicketID(i) & "'"
                    Set rsTemp = oDb.Execute(szsql)
                    
                    
                    If rsTemp.RecordCount = 1 Then
                        '写入代理取消退票明细表
'                        szSql = "INSERT INTO agent_resumereturn_lst(sell_station_id,ticket_id,operation_time," _
'                            & "user_id," _
'                            & "bus_date," _
'                            & "bus_id," _
'                            & "service_unit_id," _
'                            & "start_station_id," _
'                            & "start_station_name," _
'                            & "des_station_id," _
'                            & "des_station_name," _
'                            & "ticket_type," _
'                            & "seat_number," _
'                            & "ticket_price," _
'                            & "return_time," _
'                            & "credence_sheet_id," _
'                            & "return_charge," _
'                            & "return_mode ," _
'                            & "return_ratio)"
'                        szSql = szSql & "  VALUES('" & rsTemp!sell_station_id & "','" & TicketID(i) & "','" _
'                            & ToDBDateTime(oParam.NowDateTime) & "','" _
'                            & m_oActiveUser.UserID & "','" _
'                            & rsTemp!bus_date & "','" _
'                            & rsTemp!bus_id & "','" _
'                            & rsTemp!service_unit_id & "','" _
'                            & rsTemp!start_station_id & "','" _
'                            & rsTemp!start_station_name & "','" _
'                            & rsTemp!des_station_id & "','" _
'                            & rsTemp!des_station_name & "'," _
'                            & rsTemp!ticket_type & ",'" _
'                            & aszReturnSeatNo(i) & "'," _
'                            & rsTemp!ticket_price & ",'" _
'                            & rsTemp!operation_time & "','" _
'                            & rsTemp!credence_sheet_id & "'," _
'                            & rsTemp!return_charge & "," _
'                            & rsTemp!return_mode & "," _
'                            & rsTemp!return_ratio & ")"
'                        oDb.Execute (szSql)
        
                        szsql = "DELETE  FROM  Agent_ticket_return_lst WHERE ticket_id='" & TicketID(i) & "'"
                        oDb.Execute (szsql)
                  End If
            Next i
        On Error GoTo 0
        oDb.CommitTrans
        
        
    Else
        szsql = "SELECT unit_id FROM Ticket_sell_lst s,user_info u WHERE " _
                & " s.ticket_id='" & TicketID(1) & "' AND s.user_id=u.user_id "
        Set rsTemp1 = oDb.Execute(szsql)
        If rsTemp1.RecordCount = 0 Then
            On Error GoTo ErrorHandle
            Set oSellTicketService = CreateRemoteObject(rsTemp!service_unit_id, "STTkSvr.SellTicketService")
            aszReturnSeatNo = oSellTicketService.ResumeReturnTicket(TicketID)
            On Error GoTo 0
            On Error GoTo ErrorHandle2
            ResumeReturnClient TicketID, aszReturnSeatNo
            On Error GoTo 0
        Else
            LocalResumeReturnTicket TicketID
        End If
    End If
    Set rsTemp = Nothing
    Set rsTemp1 = Nothing
    Set oSellTicketService = Nothing
    Set oDb = Nothing
    Set oParam = Nothing
Exit Function
ErrorHandle:
   Set rsTemp = Nothing
    Set rsTemp1 = Nothing
    Set oSellTicketService = Nothing
    Set oDb = Nothing
    Set oParam = Nothing
    err.Raise err.Number
    Exit Function
ErrorHandle2:
    '加入提示"对方车站已写入,请通知管理员在对方车站将此票删除"
    oDb.RollbackTrans
    Set rsTemp = Nothing
    Set rsTemp1 = Nothing
    Set oSellTicketService = Nothing
    Set oDb = Nothing
    Set oParam = Nothing
    err.Raise err.Number, err.Source, cszServiceFinish & err.Description
    Exit Function
End Function
'取消退票
Public Function LocalResumeReturnTicket(paszTicketID() As String) As String()
    Dim oDb As New RTConnection
    Dim i As Integer
    Dim nLen As Integer
    
    AssertActiveUserValid m_oActiveUser, ERR_SellTicketService
    AssertHaveRight m_oActiveUser, RIGHT_ResumeReturnTicket
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    oDb.BeginTrans
        On Error GoTo Here
        SelfResumeReturnlTicket paszTicketID, oDb
    oDb.CommitTrans
    
    
    nLen = ArrayLength(paszTicketID)
    For i = 1 To nLen
        WriteOperateLog m_oActiveUser, RIGHT_ResumeReturnTicket, "取消退票票号:" & paszTicketID(i)
    Next i
    
    Exit Function
Here:
    oDb.RollbackTrans
    err.Raise err.Number
End Function

'内部用取消退票方法
Private Function SelfResumeReturnlTicket(paszTicketID() As String, poDb As RTConnection) As String()
    Dim szsql As String
    Dim oParam As New SystemParam
    Dim sgReturnCharge As Double, sgTicketPrice As Double
    Dim dtBusStartupDateTime As Date
    Dim rsTemp As Recordset
    Dim rsTemp1 As Recordset
    Dim rsTemp2 As Recordset
    Dim rsTemp3 As Recordset
    Dim lBusType As EBusType
    
    Dim i As Integer
    Dim nLen As Integer
    
    Dim szSeatNo As String
    Dim aszReturnSeatNo() As String
    
    nLen = 0
    nLen = ArrayLength(paszTicketID)
    
    oParam.Init m_oActiveUser
    On Error GoTo ErrorHandle
    For i = 1 To nLen
        szsql = "SELECT * FROM Ticket_return_lst WHERE ticket_id ='" & paszTicketID(i) & "'"
        Set rsTemp = poDb.Execute(szsql)
        
        szsql = "SELECT status,bus_id,bus_date,seat_no,sell_station_id  FROM Ticket_sell_lst WHERE ticket_id='" & paszTicketID(i) & "'"
        Set rsTemp2 = poDb.Execute(szsql)
        If rsTemp2.RecordCount = 1 Then
            If rsTemp2!Status <> ST_TicketReturned + ST_TicketNormal Then
                ShowError ERR_ServiceTicketNotReturn, , "票号为:" & paszTicketID(i)
            End If
        Else
            ShowError ERR_ServiceTicketNotExist, , "票号为:" & paszTicketID(i)
        End If
        If rsTemp.RecordCount = 1 Then
            '如果服务表不存在
            szSeatNo = FormatDbValue(rsTemp2!seat_no)
            If szSeatNo <> "" Then '不是滚动车次
                
                '取得该票的座位类型和状态
                szsql = "SELECT seat_type_id ,status FROM Env_bus_seat_lst WHERE " _
                    & "bus_id ='" & rsTemp2!bus_id & "' AND " _
                    & "bus_date='" & ToDBDate(rsTemp2!bus_date) & "' AND " _
                    & "seat_no='" & rsTemp2!seat_no & "' "
                Set rsTemp3 = poDb.Execute(szsql)
                If rsTemp2!Status <> ST_SeatCanSell Then
                    szsql = "SELECT TOP 1 * FROM Env_bus_seat_lst WHERE " _
                        & "bus_id='" & rsTemp2!bus_id & "' AND " _
                        & "bus_date='" & ToDBDate(rsTemp2!bus_date) & "' AND " _
                        & "seat_type_id='" & rsTemp3!seat_type_id & "' AND " _
                        & "status=" & ST_SeatCanSell & ""
                
                    Set rsTemp3 = poDb.Execute(szsql)
                    If rsTemp3.RecordCount = 0 Then ShowError ERR_ReturnTicketNotResume
                    szSeatNo = rsTemp3!seat_no
                End If
            End If
    
            szsql = "UPDATE  Ticket_sell_lst SET status='" & ST_TicketNormal & "',seat_no='" & szSeatNo & "' WHERE ticket_id='" & paszTicketID(i) & "'"
            poDb.Execute szsql
            '写入代理取消退票明细表
'            szSql = "INSERT INTO resume_returnticket_lst(" _
'                & "ticket_id," _
'                & "operation_time," _
'                & "user_id," _
'                & "return_time," _
'                & "credence_sheet_id," _
'                & "return_charge," _
'                & "return_user_id," _
'                & "return_mode,"
'            szSql = szSql & "return_ratio)  VALUES('" _
'                & paszTicketID(i) & "','" _
'                & ToDBDateTime(oParam.NowDateTime) & "','" _
'                & m_oActiveUser.UserID & "','" _
'                & ToDBDateTime(rsTemp!operation_time) & "','" _
'                & rsTemp!credence_sheet_id & "'," _
'                & rsTemp!return_charge & ",'" _
'                & rsTemp!user_id & "'," _
'                & rsTemp!return_ticket_mode & "," _
'                & rsTemp!return_ratio & ")"
'            poDb.Execute szSql
            
            szsql = "DELETE Ticket_return_lst WHERE ticket_id='" & paszTicketID(i) & "'"
            poDb.Execute szsql
        End If
            
        szsql = "SELECT * FROM Agent_ticket_return_lst WHERE ticket_id='" & paszTicketID(i) & "'"
        Set rsTemp = poDb.Execute(szsql)
'        szSql = "SELECT seat_no FROM Ticket_sell_lst WHERE ticket_id='" & paszTicketID(i) & "'"
'        Set rsTemp1 = poDb.Execute(szSql)
        szsql = "SELECT bus_id, bus_date, seat_no, ticket_price, ticket_sell_lst.sell_station_id, user_info.sell_station_id seller_sell_station_id  FROM Ticket_sell_lst, user_info WHERE " _
            & " ticket_id='" & paszTicketID(i) & "' AND " _
            & " Ticket_sell_lst.user_id = user_info.user_id "
        Set rsTemp1 = poDb.Execute(szsql)
        
        szSeatNo = FormatDbValue(rsTemp1!seat_no)
        
        
        If rsTemp1.RecordCount = 1 Then
            szsql = "UPDATE Env_bus_allot_lst SET sell_quantity=sell_quantity+1 WHERE bus_id='" & rsTemp1!bus_id & "'" _
                   & " AND sell_station_id='" & rsTemp1!sell_station_id & "' AND bus_date='" & rsTemp1!bus_date & "'"
            poDb.Execute szsql
            
            szsql = "UPDATE env_bus_sell_station_lst SET sell_quantity=sell_quantity+1 WHERE bus_id='" & rsTemp1!bus_id & "'" _
                   & " AND sell_station_id='" & rsTemp1!seller_sell_station_id & "' AND bus_date='" & rsTemp1!bus_date & "'"
            poDb.Execute szsql
            
        End If
        
        If rsTemp.RecordCount = 1 Then
        
            szsql = "UPDATE Agent_ticket_sale_lst SET status=" _
                & ST_TicketNormal & ",seat_no='" & szSeatNo & "' WHERE " _
                & " ticket_id='" & paszTicketID(i) & "'"
            poDb.Execute (szsql)
            '写入代理取消退票明细表
'            szSql = "INSERT INTO agent_resumereturn_lst(sell_station_id,ticket_id,operation_time," _
'                & "user_id," _
'                & "bus_date," _
'                & "bus_id," _
'                & "service_unit_id," _
'                & "start_station_id," _
'                & "start_station_name," _
'                & "des_station_id," _
'                & "des_station_name," _
'                & "ticket_type," _
'                & "seat_number," _
'                & "ticket_price," _
'                & "return_time," _
'                & "credence_sheet_id," _
'                & "return_charge," _
'                & "return_mode ," _
'                & "return_ratio)"
'            szSql = szSql & "  VALUES('" & rsTemp!sell_station_id & "','" & paszTicketID(i) & "','" _
'                & ToDBDateTime(oParam.NowDateTime) & "','" _
'                & m_oActiveUser.UserID & "','" _
'                & rsTemp!bus_date & "','" _
'                & rsTemp!bus_id & "','" _
'                & rsTemp!service_unit_id & "','" _
'                & rsTemp!start_station_id & "','" _
'                & rsTemp!start_station_name & "','" _
'                & rsTemp!des_station_id & "','" _
'                & rsTemp!des_station_name & "'," _
'                & rsTemp!ticket_type & ",'" _
'                & szSeatNo & "'," _
'                & rsTemp!ticket_price & ",'" _
'                & rsTemp!operation_time & "','" _
'                & rsTemp!credence_sheet_id & "'," _
'                & rsTemp!return_charge & "," _
'                & rsTemp!return_mode & "," _
'                & rsTemp!return_ratio & ")"
'            poDb.Execute (szSql)
            szsql = "DELETE  FROM  Agent_ticket_return_lst WHERE ticket_id='" & paszTicketID(i) & "'"
            poDb.Execute (szsql)
    
            '释放座位
            ResumeSeat paszTicketID(i), poDb
        End If
    Next i
    Set rsTemp = Nothing
    Set rsTemp1 = Nothing
    Set rsTemp2 = Nothing
    Set rsTemp3 = Nothing
    On Error GoTo 0
    Exit Function
ErrorHandle:
    Set rsTemp = Nothing
    Set rsTemp1 = Nothing
    Set rsTemp2 = Nothing
    Set rsTemp3 = Nothing
    err.Raise err.Number
End Function
'////////////////////////////////
'取消废票
Public Function ResumeCancelTicket(TicketID() As String) As String()
    
    Dim oDb As New RTConnection
    Dim szsql As String
    Dim rsTemp As Recordset
    Dim rsTemp1 As Recordset
    Dim aszCancelSeatNo() As String
    Dim oSellTicketService As SellTicketService
    
    Dim i As Integer
    Dim oParam As New SystemParam
    
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    szsql = "SELECT service_unit_id FROM Agent_ticket_sale_lst WHERE ticket_id='" & TicketID(1) & "' "
    Set rsTemp = oDb.Execute(szsql)
    If rsTemp.RecordCount <> 1 Then
        szsql = "SELECT unit_id FROM Ticket_sell_lst s,user_info u WHERE " _
            & " s.ticket_id='" & TicketID(1) & "' AND s.user_id=u.user_id "
        Set rsTemp = oDb.Execute(szsql)
        If rsTemp.RecordCount = 0 Then ShowError ERR_TicketNotExist
        On Error GoTo ErrorHandle
        aszCancelSeatNo = m_oCurrentSellService.ResumeCancelTicket(TicketID)
        On Error GoTo 0
        '        Set oSellTicketService = CreateRemoteObject(rsTemp!unit_id, "STTkSvr.SellTicketService")
        '        oSellTicketService.ResumeCancelClient TicketID, aszCancelSeatNo
        oDb.BeginTrans
        On Error GoTo ErrorHandle2
        For i = 1 To ArrayLength(TicketID)
            szsql = "SELECT * FROM Agent_ticket_cancel_lst WHERE ticket_id='" & TicketID(i) & "'"
            Set rsTemp = oDb.Execute(szsql)
        
            If rsTemp.RecordCount = 1 Then
                '写入代理取消废票明细表
'                szSql = "INSERT INTO agent_resumecancel_lst (" _
'                & "sell_station_id," _
'                & "ticket_id," _
'                & "operation_time," _
'                & "user_id," _
'                & "cancel_time," _
'                & "cancel_mode," _
'                & "cancel_user_id) VALUES ('" _
'                & rsTemp!sell_station_id & "','" _
'                & TicketID(i) & "','" _
'                & ToDBDateTime(oParam.NowDateTime) & "','" _
'                & m_oActiveUser.UserID & "','" _
'                & rsTemp!cancel_time & "'," _
'                & rsTemp!cancel_mode & ",'" _
'                & rsTemp!user_id & "')"
'                oDb.Execute szSql
                
                szsql = "DELETE Agent_ticket_cancel_lst WHERE ticket_id='" & TicketID(i) & "'"
                oDb.Execute szsql
            End If
    
        Next i
        On Error GoTo 0
        oDb.CommitTrans
    Else
        szsql = "SELECT unit_id FROM Ticket_sell_lst s,user_info u WHERE " _
            & " s.ticket_id='" & TicketID(1) & "' AND s.user_id=u.user_id "
        Set rsTemp1 = oDb.Execute(szsql)
        If rsTemp1.RecordCount = 0 Then
            On Error GoTo ErrorHandle
            Set oSellTicketService = CreateRemoteObject(rsTemp!service_unit_id, "STTkSvr.SellTicketService")
            aszCancelSeatNo = oSellTicketService.ResumeCancelTicket(TicketID)
            On Error GoTo 0
            On Error GoTo ErrorHandle2
            ResumeCancelClient TicketID, aszCancelSeatNo
            On Error GoTo 0
        Else
            LocalResumeCancelTicket TicketID
        End If
    End If
    
    Set oSellTicketService = Nothing
    Set rsTemp = Nothing
    Set rsTemp1 = Nothing
    Set oDb = Nothing
    Set oParam = Nothing
    Exit Function
ErrorHandle:
    Set oSellTicketService = Nothing
    Set rsTemp = Nothing
    Set rsTemp1 = Nothing
    Set oDb = Nothing
    Set oParam = Nothing
    err.Raise err.Number
    Exit Function
ErrorHandle2:
    '加入提示"对方车站已写入,请通知管理员在对方车站将此票删除"
    oDb.RollbackTrans
    Set oSellTicketService = Nothing
    Set rsTemp = Nothing
    Set rsTemp1 = Nothing
    Set oDb = Nothing
    Set oParam = Nothing
    err.Raise err.Number, err.Source, cszServiceFinish & err.Description
    Exit Function
End Function

'取消废票
Public Function LocalResumeCancelTicket(paszTicketID() As String) As String()
    Dim oDb As New RTConnection
    Dim szsql As String
    Dim i As Integer
    Dim nLen As Integer
    nLen = 0
    nLen = ArrayLength(paszTicketID)
    AssertActiveUserValid m_oActiveUser, ERR_SellTicketService
    AssertHaveRight m_oActiveUser, RIGHT_ResumeCancelTicket
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    oDb.BeginTrans
        On Error GoTo Here
        SelfResumeCancelTicket paszTicketID, oDb
    oDb.CommitTrans
    For i = 1 To nLen
        WriteOperateLog m_oActiveUser, RIGHT_ResumeCancelTicket, "取消废票票号:" & paszTicketID(i)
    Next i
    Exit Function
Here:
    oDb.RollbackTrans
    err.Raise err.Number
End Function

'内部用取消废票
Private Function SelfResumeCancelTicket(paszTicketID() As String, poDb As RTConnection) As String()
    Dim szsql As String
    Dim oParam As New SystemParam
    Dim rsTemp As Recordset
    Dim rsTemp1 As Recordset
    Dim rsTemp2 As Recordset
    Dim i As Integer
    Dim nLen As Integer
    Dim szSeatNo As String
    Dim aszReturnSeatNo() As String
    nLen = 0
    nLen = ArrayLength(paszTicketID)
    
    oParam.Init m_oActiveUser
    On Error GoTo ErrorHandle
    For i = 1 To nLen
        szsql = "SELECT * FROM ticket_cancel_lst WHERE ticket_id='" & paszTicketID(i) & "'"
        Set rsTemp = poDb.Execute(szsql)
'        szSql = "SELECT status,bus_date,bus_id,seat_no FROM Ticket_sell_lst WHERE ticket_id='" & paszTicketID(i) & "'"
'        Set rsTemp1 = poDb.Execute(szSql)
        
        szsql = "SELECT ticket_sell_lst.status, bus_date, bus_id, seat_no, ticket_price, ticket_sell_lst.sell_station_id, user_info.sell_station_id seller_sell_station_id  FROM Ticket_sell_lst, user_info WHERE " _
            & " ticket_id='" & paszTicketID(i) & "' AND " _
            & " Ticket_sell_lst.user_id = user_info.user_id "
        Set rsTemp1 = poDb.Execute(szsql)
        
        If rsTemp1.RecordCount = 1 Then
            If rsTemp1!Status <> ST_TicketCanceled + ST_TicketNormal Then
                ShowError ERR_ServiceTicketNotCancel, , "票号为:" & paszTicketID(i)
            End If
        Else
            If rsTemp1.RecordCount = 0 Then
                ShowError ERR_ServiceTicketNotExist, , "票号为:" & paszTicketID(i)
            End If
        End If
        If rsTemp.RecordCount = 1 Then
            szSeatNo = FormatDbValue(rsTemp1!seat_no)
            
            szsql = "UPDATE Env_bus_allot_lst SET sell_quantity=sell_quantity+1 WHERE bus_id='" & rsTemp1!bus_id & "'" _
                   & " AND sell_station_id='" & rsTemp1!sell_station_id & "' AND bus_date='" & rsTemp1!bus_date & "'"
            poDb.Execute szsql
            
            szsql = "UPDATE env_bus_sell_station_lst SET sell_quantity=sell_quantity+1 WHERE bus_id='" & rsTemp1!bus_id & "'" _
                   & " AND sell_station_id='" & rsTemp1!seller_sell_station_id & "' AND bus_date='" & rsTemp1!bus_date & "'"
            poDb.Execute szsql
            
            
            If szSeatNo <> "" Then
                
                szsql = "SELECT seat_type_id ,status FROM Env_bus_seat_lst WHERE " _
                    & "bus_id='" & rsTemp1!bus_id & "' AND " _
                    & "bus_date='" & ToDBDate(rsTemp1!bus_date) & "' AND " _
                    & "seat_no='" & rsTemp1!seat_no & "'"
                Set rsTemp2 = poDb.Execute(szsql)
                
                If rsTemp2!Status <> ST_SeatCanSell Then
                szsql = "SELECT TOP 1 seat_no FROM Env_bus_seat_lst WHERE  " _
                    & "bus_id='" & rsTemp1!bus_id & "' AND " _
                    & "bus_date='" & ToDBDate(rsTemp1!bus_date) & "' AND " _
                    & "seat_type_id='" & rsTemp2!seat_type_id & "' AND " _
                    & "status=" & ST_SeatCanSell & ""
                Set rsTemp2 = poDb.Execute(szsql)
                If rsTemp2.RecordCount = 0 Then ShowError ERR_CancelTicketNotResume
                szSeatNo = rsTemp2!seat_no
            End If
        End If
        szsql = "UPDATE Ticket_sell_lst SET " _
            & " status= " & ST_TicketNormal & ",seat_no='" & szSeatNo & "' WHERE " _
            & " ticket_id='" & paszTicketID(i) & "'"
        poDb.Execute szsql
        '写入取消废票明细表
'        szSql = "INSERT INTO resume_cancelticket_lst( " _
'            & " ticket_id," _
'            & " operation_time," _
'            & " user_id," _
'            & " cancel_time," _
'            & " cancel_user_id," _
'            & " cancel_mode ) VALUES('" _
'            & paszTicketID(i) & "','" _
'            & ToDBDateTime(oParam.NowDateTime) & "','" _
'            & m_oActiveUser.UserID & "','" _
'            & rsTemp!cancel_ticket_time & "','" _
'            & rsTemp!user_id & "'," _
'            & rsTemp!cancel_mode & ")"
'        poDb.Execute szSql
        
        szsql = "DELETE ticket_cancel_lst WHERE ticket_id='" & paszTicketID(i) & "'"
        poDb.Execute szsql
        
    End If
    szsql = "SELECT * FROM Agent_ticket_cancel_lst WHERE ticket_id='" & paszTicketID(i) & "'"
    Set rsTemp = poDb.Execute(szsql)
    szsql = "SELECT seat_no FROM Ticket_sell_lst WHERE ticket_id='" & paszTicketID(i) & "'"
    Set rsTemp1 = poDb.Execute(szsql)
    
    If rsTemp.RecordCount = 1 Then
        
        
        szsql = "UPDATE Agent_ticket_sale_lst SET status=" & ST_TicketNormal & ",seat_no='" & rsTemp1!seat_no & "'" _
            & " WHERE ticket_id= '" & paszTicketID(i) & "'"
        poDb.Execute szsql
        '写入代理取消废票明细表
'        szSql = "INSERT INTO agent_resumecancel_lst (" _
'            & "ticket_id," _
'            & "operation_time," _
'            & "user_id," _
'            & "cancel_time," _
'            & "cancel_mode," _
'            & "cancel_user_id) VALUES ('" _
'            & paszTicketID(i) & "','" _
'            & ToDBDateTime(oParam.NowDateTime) & "','" _
'            & m_oActiveUser.UserID & "','" _
'            & rsTemp!cancel_time & "'," _
'            & rsTemp!cancel_mode & ",'" _
'            & rsTemp!user_id & "')"
'        poDb.Execute szSql
        
        szsql = "DELETE Agent_ticket_cancel_lst WHERE ticket_id='" & paszTicketID(i) & "'"
        poDb.Execute szsql
    End If
    '恢复座位
    ResumeSeat paszTicketID(i), poDb
    Next i
    Set rsTemp = Nothing
    Set rsTemp1 = Nothing
    Set rsTemp2 = Nothing
    On Error GoTo 0
    Exit Function
ErrorHandle:
    Set rsTemp = Nothing
    Set rsTemp1 = Nothing
    Set rsTemp2 = Nothing
    err.Raise err.Number
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
'//////////////////////////////////////
'得到退票废率
Public Function GetReturnRatioValue() As RETURNFEES()
    
    GetReturnRatioValue = m_oCurrentSellService.GetReturnRatioValue
End Function
'////////////////////////////////////////////////////////
'得到对方车站的所有票种
Public Function GetAllTicketType(Optional nValid As Integer = TP_TicketTypeAll) As TTicketType()
On Error GoTo Here
    GetAllTicketType = m_oCurrentSellService.GetAllTicketType(nValid)
    Exit Function
Here:
    err.Raise err.Number
End Function
' ////////////////////////////////
'得到对方车站有不同座位类型的车次
Public Function GetMultiSeatTypeBus() As TMultiSeatTypeBus
On Error GoTo Here
    GetMultiSeatTypeBus = m_oCurrentSellService.GetMultiSeatTypeBus()
    Exit Function
Here:
    err.Raise err.Number
End Function
'//////////////////////////////////////
'得到对方车站售票明细票价和相应的座位号
Public Function GetSaleTicketPrice(BusID As String, BusStartDate As Date, EndStationID As String, TicketType() As ETicketType, SeatNo() As String, Optional nOrderSeatNo As Integer = 0) As Single()
On Error GoTo Here
    GetSaleTicketPrice = m_oCurrentSellService.GetSaleTicketPrice(BusID, BusStartDate, EndStationID, TicketType, SeatNo, nOrderSeatNo)
 Exit Function
Here:
    err.Raise err.Number
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
'////////////////////////
'得到对方车站座位类型代码和名称
Public Function GetAllSeatType() As String()

GetAllSeatType = m_oCurrentSellService.GetAllSeatType

End Function



Private Function GetLocalBusRsEx88(BusDate As Date, DesStationID As String, szCheckGateID As String, Optional IgnoreOffTime As Boolean = False, Optional BusID As String = "") As Recordset
'Dim oDb As New RTConnection
'    Dim oParam As New SystemParam
'    Dim szSql As String
'    Dim rsTemp As Recordset
'    Dim nStopSellTime As Integer
'    Dim dtNowDateTime As Date
'
'    Dim dbExtraBeforeTime As Double
'    Dim dbExtraAfterTime As Double
'
'    Dim aszCheckGateID() As String
'    Dim i As Integer
'    Dim nLen As Integer
'
'
'    aszCheckGateID = StringToTeam(szCheckGateID)
'    nLen = ArrayLength(aszCheckGateID)
'
'    oParam.Init m_oActiveUser
'    nStopSellTime = oParam.StopSellTime
'    dtNowDateTime = oParam.NowDateTime
'
'
'    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
'
'    szSql = "SELECT Work_env_bus_info.* ,seat_type_id,route_name ,rs.end_station_mileage,Work_env_bus_station_lst.sell_station_id,sell_station_name,e.bus_start_time as busstarttime," _
'        & "         full_price, half_price,preferential_ticket1,preferential_ticket2,preferential_ticket3," _
'        & "         sale_ticket_quantity,stop_sale_time " _
'        & "  FROM Work_env_bus_info,Work_env_bus_station_lst ,route_info,sell_station_info s,Work_env_bus_allot_lst e,route_section_lst rs" _
'        & "  WHERE s.sell_station_id=Work_env_bus_station_lst.sell_station_id AND " _
'        & "  e.bus_date=Work_env_bus_info.bus_date AND e.bus_id=Work_env_bus_info.bus_id AND "
'    If nLen = 1 Then
'        szSql = szSql & " Work_env_bus_info.check_gate_id='" & aszCheckGateID(1) & "'"
'    Else
'        szSql = szSql & " ( Work_env_bus_info.check_gate_id='" & aszCheckGateID(1) & "' OR "
'        For i = 2 To nLen
'            If i = nLen Then
'                szSql = szSql & " Work_env_bus_info.check_gate_id='" & aszCheckGateID(i) & "' ) And "
'            Else
'                szSql = szSql & "  Work_env_bus_info.check_gate_id='" & aszCheckGateID(i) & "' OR "
'            End If
'        Next i
'    End If
'
'    If Not IgnoreOffTime Then
'        szSql = szSql & "(('" & ToDBDateTime(DateAdd("s", nStopSellTime * 60, dtNowDateTime)) & "'<=e.bus_start_time AND bus_type<>" & TP_ScrollBus & ") OR " _
'        & " (bus_type=" & TP_ScrollBus & " OR e.bus_start_time>='" & ToDBDateTime(DateAdd("s", nStopSellTime * 60, dtNowDateTime)) & "')) AND  "
'        szSql = szSql _
'        & " Work_env_bus_info.bus_date='" & ToDBDate(BusDate) & "' AND " _
'        & "(Work_env_bus_info.bus_type=2 OR Work_env_bus_info.bus_type=3) AND" _
'        & " Work_env_bus_info.bus_id=Work_env_bus_station_lst.bus_id AND " _
'        & " Work_env_bus_info.bus_date=Work_env_bus_station_lst.bus_date AND " _
'        & " Work_env_bus_station_lst.station_id='" & DesStationID & "' AND " _
'        & " ( sale_seat_quantity>=0 OR sale_stand_seat_quantity>0 OR bus_type=" & TP_ScrollBus & ") AND " _
'        & " ( ( max_sale_quantity>0 AND  sale_ticket_quantity>=0) OR  " _
'        & " ( max_sale_quantity<0) ) AND Work_env_bus_info.route_id=route_info.route_id "
'        'ORDER BY Work_env_bus_info.bus_start_time ,Work_env_bus_info.bus_id "
'
'    Else
'        dbExtraBeforeTime = oParam.ExtraBusShowBefore
'        dbExtraBeforeTime = oParam.ExtraBusShowAfter
'
'        szSql = szSql & "'" & ToDBDateTime(DateAdd("n", dbExtraAfterTime, dtNowDateTime)) & "'>=e.bus_start_time AND " _
'        & "'" & ToDBDateTime(DateAdd("n", -dbExtraBeforeTime, dtNowDateTime)) & "'<=e.bus_start_time AND" _
'        & " bus_type<>" & TP_ScrollBus & " AND  " _
'        & " Work_env_bus_info.bus_date='" & ToDBDate(BusDate) & "' AND " _
'        & "(Work_env_bus_info.bus_type=2 OR Work_env_bus_info.bus_type=3) AND" _
'        & " Work_env_bus_info.bus_id=Work_env_bus_station_lst.bus_id AND " _
'        & " Work_env_bus_info.bus_date=Work_env_bus_station_lst.bus_date AND " _
'        & " Work_env_bus_station_lst.station_id='" & DesStationID & "' AND " _
'        & " ( sale_seat_quantity>=0 OR sale_stand_seat_quantity>0 OR bus_type=" & TP_ScrollBus & ") AND " _
'        & " ( ( max_sale_quantity>0 AND  sale_ticket_quantity>0) OR  " _
'        & " ( max_sale_quantity<0) ) AND Work_env_bus_info.route_id=route_info.route_id " _
'        & " AND  rs.sell_station_id=e.sell_station_id AND Work_env_bus_info.route_id=rs.route_id AND rs.end_station_id=Work_env_bus_station_lst.station_id "
'        'ORDER BY Work_env_bus_info.bus_start_time,Work_env_bus_info.bus_id  "
'
'
'    End If
'
'    '判断是否可以相互互联售票
'       If BusID <> "" Then
'           szSql = szSql & " AND Work_env_bus_info.bus_id='" & Trim(BusID) & "'"
'       End If
'    If Not m_bSellStationCanSellEachOther Then
'       If m_oActiveUser.SellStationID <> "" Then
'           szSql = szSql & " AND Work_env_bus_station_lst.sell_station_id='" & Trim(m_oActiveUser.SellStationID) & "'"
'       End If
'    End If
'
'    szSql = szSql & " ORDER BY e.bus_start_time,Work_env_bus_info.bus_id "
'
'    Set rsTemp = oDb.Execute(szSql)
'    Set GetLocalBusRsEx88 = rsTemp
'    Set rsTemp = Nothing
'    Set oParam = Nothing
'    Set oDb = Nothing
End Function



'检查了票状态，是否是同一人，是否已过正常退（废）时间,免票不能改签、退
Private Sub AssertTicketCanCancelReturn(pszTicketID As String, poDb As RTConnection, pnDealType As ETicketClientType, pbForce As Boolean)
    Dim szsql As String
    Dim rsTemp As Recordset
    Dim nTicketStatus As ETicketStatus
    Dim nTicketType As ETicketType
    Dim szSeller As String
    Dim dtSellDateTime As Date
    Dim oParam As New SystemParam
    Dim dtNowDateTime As Date
    Dim dtStartupDateTime As Date
    Dim lBusType As EBusType
    Dim dtBusDate As Date
    
'    szSql = "SELECT Ticket_sell_lst.*,bus_start_time,bus_type FROM Ticket_sell_lst,Work_env_bus_info WHERE " _
'    & "Ticket_sell_lst.bus_id=Work_env_bus_info.bus_id AND " _
'    & "Ticket_sell_lst.bus_date=Work_env_bus_info.bus_date AND " _
'    & "ticket_id='" & pszTicketID & "'"
'    Set rsTemp = poDb.Execute(szSql)
'    If rsTemp.RecordCount <> 1 Then ShowError ERR_TicketNotExist, , "票号为:" & pszTicketID
'
    szsql = "SELECT Ticket_sell_lst.*,a.bus_start_time,bus_type FROM Ticket_sell_lst,Work_env_bus_info , Work_env_bus_allot_lst a WHERE " _
    & "Ticket_sell_lst.bus_id=Work_env_bus_info.bus_id AND " _
    & "Ticket_sell_lst.bus_date=Work_env_bus_info.bus_date AND " _
    & "Ticket_sell_lst.bus_id = a.bus_id AND Ticket_sell_lst.bus_date = a.bus_date AND Ticket_sell_lst.sell_station_id = a.sell_station_id " _
    & " AND ticket_id='" & pszTicketID & "'"
    
    Set rsTemp = poDb.Execute(szsql)
    If rsTemp.RecordCount <> 1 Then ShowError ERR_TicketNotExist, , "票号为：" & pszTicketID
    



    nTicketStatus = FormatDbValue(rsTemp!Status)
    nTicketType = FormatDbValue(rsTemp!ticket_type)
    lBusType = FormatDbValue(rsTemp!bus_type)
    dtBusDate = FormatDbValue(rsTemp!bus_date)
    
    If (nTicketStatus And ST_TicketCanceled) <> 0 Then ShowError ERR_TicketIsCanceled, , "票号为:" & pszTicketID
    If (nTicketStatus And ST_TicketChanged) <> 0 Then ShowError ERR_TicketIsChanged, , "票号为:" & pszTicketID
    If (nTicketStatus And ST_TicketChecked) <> 0 Then ShowError ERR_TicketIsChecked, , "票号为:" & pszTicketID
    If (nTicketStatus And ST_TicketReturned) <> 0 Then ShowError ERR_TicketIsReturned, , "票号为:" & pszTicketID
    
    oParam.Init m_oActiveUser
    
    If pnDealType = TP_ReturnTicket Then '进行退票(免票不能退)
        If nTicketType = TP_FreeTicket Then ShowError ERR_FreeTicketCanNotReturnChange, , "票号为:" & pszTicketID
    Else '进行废票或改签
        If (nTicketStatus And ST_TicketSellChange) <> 0 Then ShowError ERR_TicketIsChangeTicket, , "票号为:" & pszTicketID
        
        If pnDealType = TP_CancelTicket Then '废票
            If Not pbForce Then
                szSeller = FormatDbValue(rsTemp!user_id)
                If Trim(szSeller) <> Trim(m_oActiveUser.UserID) Then ShowError ERR_UserIsNotTicketSeller, , "票号为:" & pszTicketID
            Else ' 强行废票
                If m_oActiveUser.UserUnitID <> oParam.UnitID Then ShowError ERR_ForceCancelTicketRemoteNotAllowed, , "票号为:" & pszTicketID
            End If
        Else '改签
            If nTicketType = TP_FreeTicket Then ShowError ERR_FreeTicketCanNotReturnChange, , "票号为:" & pszTicketID
        End If
    End If

    
    If Not pbForce Then
        dtSellDateTime = FormatDbValue(rsTemp!operation_time)
        dtNowDateTime = oParam.NowDateTime
        
        If pnDealType = TP_ReturnTicket Then '进行退票
            If lBusType <> TP_ScrollBus Then
                dtStartupDateTime = SelfGetFullDateTime(FormatDbValue(rsTemp!bus_date), FormatDbValue(rsTemp!bus_start_time))
                If DateDiff("n", dtNowDateTime, dtStartupDateTime) < oParam.ReturnTicketTime3 Then ShowError ERR_ReturnTicketTimeOut, , "票号为:" & pszTicketID
            Else
                If oParam.NowDate > dtBusDate Then ShowError ERR_ScroolBusReturnTicketDateOut, , "票号为:" & pszTicketID
                If Not (oParam.ScrollBusCanReturnTicket Or pbForce) Then ShowError ERR_ScrollBusTicketCanNotReturn, , "票号为:" & pszTicketID
            End If
            
        ElseIf pnDealType = TP_ChangeTicket Then '进行改签
            If lBusType <> TP_ScrollBus Then
                dtStartupDateTime = SelfGetFullDateTime(FormatDbValue(rsTemp!bus_date), FormatDbValue(rsTemp!bus_start_time))
                If DateDiff("n", dtNowDateTime, dtStartupDateTime) < oParam.ReturnTicketTime3 Then ShowError ERR_ReturnTicketTimeOut, , "票号为:" & pszTicketID
            Else
                If oParam.NowDate > dtBusDate Then ShowError ERR_ScroolBusReturnTicketDateOut, , "票号为:" & pszTicketID
            End If
        Else '进行废票
            If DateDiff("n", dtSellDateTime, dtNowDateTime) > oParam.CancelTicketTime Then ShowError ERR_CancelTicketTimeOut, , "票号为:" & pszTicketID
        End If
    End If

End Sub

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
    
    Dim rsSTemp As Recordset
    szsql = "SELECT * FROM Ticket_sell_lst WHERE bus_id='" & szBusID & "' AND bus_date='" & dtDate & "' AND seat_no='" & szSeatNo & "' AND (status='" & ST_TicketNormal & "' or status='" & ST_TicketSellChange & "')"
    Set rsSTemp = poDb.Execute(szsql)

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
        If rsSTemp.RecordCount = 0 Then
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
    End If
    
    '更改环境车次站点表的可售票数
    If rsSTemp.RecordCount = 0 Then
    szsql = "UPDATE Env_bus_station_lst SET " _
    & " sale_ticket_quantity=sale_ticket_quantity+1 WHERE " _
    & " bus_id='" & szBusID & "' AND " _
    & " bus_date='" & ToDBDate(dtDate) & "' AND " _
    & " station_id='" & szDesStationID & "' AND " _
    & " sell_station_id='" & szSellStationID & "' AND " _
    & " max_sale_quantity>0 "
    
    poDb.Execute szsql
    End If
    
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

Private Function GetUserLoginUnitInfo(pszUnitID As String) As TUserLoginUnitInfo
    Dim oDb As New RTConnection
    Dim szsql As String
    Dim rsTemp As Recordset
    Dim uliTemp As TUserLoginUnitInfo
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    
    szsql = "SELECT ip_address,user_password,user_login_unit_info.remote_user_id FROM connect_unit_info,user_login_unit_info,remote_user_info WHERE " _
    & " connect_unit_info.unit_id='" & pszUnitID & "' AND " _
    & " user_login_unit_info.unit_id=connect_unit_info.unit_id AND " _
    & " user_login_unit_info.user_id='" & m_oActiveUser.UserID & "' AND " _
    & " user_login_unit_info.unit_id=remote_user_info.unit_id AND " _
    & " user_login_unit_info.remote_user_id=remote_user_info.remote_user_id "
    Set rsTemp = oDb.Execute(szsql)
    If rsTemp.RecordCount <> 1 Then ShowError ERR_UserCanNotLoginSpecialUnit
    
    uliTemp.szLoginUserID = FormatDbValue(rsTemp!remote_user_id)
    uliTemp.szPassword = FormatDbValue(rsTemp!user_password)
    uliTemp.szIPAddress = FormatDbValue(rsTemp!ip_address)
    
    GetUserLoginUnitInfo = uliTemp
End Function

'///////////////////////////////////////////
Private Function CreateRemoteObject(pszUnitID As String, pszObj As String) As Object

        Dim uliTemp As TUserLoginUnitInfo
        Dim oTemp As Object
        
        uliTemp = GetUserLoginUnitInfo(Trim(pszUnitID))
        
        
        Set oSystemAgent = CreateObject("STSystemBase.SystemAgent", uliTemp.szIPAddress)
        Set oAgentActiveUser = CreateObject("STSystemBase.ActiveUser", uliTemp.szIPAddress)
        oAgentActiveUser.Login uliTemp.szLoginUserID, uliTemp.szPassword, m_oActiveUser.LoginHost
        
        Set oTemp = oSystemAgent.CreateServiceObject(pszObj)

        If oSystemAgent.LocalUnit.szUnitID = Trim(m_oActiveUser.ServerUnit) Then
           oTemp.Init m_oActiveUser
        Else
           oTemp.Init oAgentActiveUser
        End If
        Set CreateRemoteObject = oTemp
        
        Set oTemp = Nothing
       
        
End Function

'客户端取消退票
Private Function ResumeReturnClient(TicketID() As String, SeatNo() As String) As String()
    Dim oDb As New RTConnection
    Dim szsql As String
    Dim rsTemp As Recordset
    Dim oParam As New SystemParam
    Dim i As Integer
    Dim szSeatNo As String
    
    oParam.Init m_oActiveUser
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    On Error GoTo ErrHandle
    oDb.BeginTrans
      For i = 1 To ArrayLength(TicketID)
        szsql = "SELECT * FROM Agent_ticket_return_lst WHERE ticket_id='" & TicketID(i) & "'"
            Set rsTemp = oDb.Execute(szsql)
            
            szSeatNo = SeatNo(i)
            If rsTemp.RecordCount = 1 Then

                szsql = "UPDATE Agent_ticket_sale_lst SET status=" _
                & ST_TicketNormal & ",seat_no='" & szSeatNo & "' WHERE " _
                & " ticket_id='" & TicketID(i) & "'"
                oDb.Execute (szsql)
                '写入代理取消退票明细表
'                szSql = "INSERT INTO agent_resumereturn_lst(sell_station_id,ticket_id,operation_time," _
'                & "user_id," _
'                & "bus_date," _
'                & "bus_id," _
'                & "service_unit_id," _
'                & "start_station_id," _
'                & "start_station_name," _
'                & "des_station_id," _
'                & "des_station_name," _
'                & "ticket_type," _
'                & "seat_number," _
'                & "ticket_price," _
'                & "return_time," _
'                & "credence_sheet_id," _
'                & "return_charge," _
'                & "return_mode ," _
'                & "return_ratio)"
'                szSql = szSql & "  VALUES('" & rsTemp!sell_station_id & "','" & TicketID(i) & "','" _
'                & ToDBDateTime(oParam.NowDateTime) & "','" _
'                & m_oActiveUser.UserID & "','" _
'                & rsTemp!bus_date & "','" _
'                & rsTemp!bus_id & "','" _
'                & rsTemp!service_unit_id & "','" _
'                & rsTemp!start_station_id & "','" _
'                & rsTemp!start_station_name & "','" _
'                & rsTemp!des_station_id & "','" _
'                & rsTemp!des_station_name & "'," _
'                & rsTemp!ticket_type & ",'" _
'                & szSeatNo & "'," _
'                & rsTemp!ticket_price & ",'" _
'                & rsTemp!operation_time & "','" _
'                & rsTemp!credence_sheet_id & "'," _
'                & rsTemp!return_charge & "," _
'                & rsTemp!return_mode & "," _
'                & rsTemp!return_ratio & ")"
'                oDb.Execute (szSql)

                szsql = "DELETE  FROM  Agent_ticket_return_lst WHERE ticket_id='" & TicketID(i) & "'"
                oDb.Execute (szsql)
            
                
            End If
    Next i
    oDb.CommitTrans
    Set oParam = Nothing
    Set rsTemp = Nothing
    Set oDb = Nothing
Exit Function
ErrHandle:
   oDb.RollbackTrans
   Set oParam = Nothing
   Set oDb = Nothing
   Set rsTemp = Nothing
   err.Raise err.Number
End Function

Private Function ResumeCancelClient(TicketID() As String, SeatNo() As String) As String
Dim oDb As New RTConnection
Dim oParam As New SystemParam
Dim rsTemp As Recordset
Dim szsql As String
Dim szSeatNo As String
Dim i As Integer

oParam.Init m_oActiveUser
oDb.ConnectionString = GetConnectionStr(cszSellTicket)
On Error GoTo ErrHandle
    oDb.BeginTrans
    For i = 1 To ArrayLength(TicketID)
         szsql = "SELECT * FROM Agent_ticket_cancel_lst WHERE ticket_id='" & TicketID(i) & "'"
         Set rsTemp = oDb.Execute(szsql)
         
         If rsTemp.RecordCount = 1 Then
                szSeatNo = SeatNo(i)
                szsql = "UPDATE Agent_ticket_sale_lst SET status=" & ST_TicketNormal & ",seat_no='" & szSeatNo & "'" _
                      & " WHERE ticket_id= '" & TicketID(i) & "'"
                oDb.Execute szsql
                
'                szSql = "INSERT INTO agent_resumecancel_lst (" _
'                & "sell_station_id," _
'                & "ticket_id," _
'                & "operation_time," _
'                & "user_id," _
'                & "cancel_time," _
'                & "cancel_mode," _
'                & "cancel_user_id) VALUES ('" _
'                & rsTemp!sell_station_id & "','" _
'                & TicketID(i) & "','" _
'                & ToDBDateTime(oParam.NowDateTime) & "','" _
'                & m_oActiveUser.UserID & "','" _
'                & rsTemp!cancel_time & "'," _
'                & rsTemp!cancel_mode & ",'" _
'                & rsTemp!user_id & "')"
'                oDb.Execute szSql
                
                szsql = "DELETE Agent_ticket_cancel_lst WHERE ticket_id='" & TicketID(i) & "'"
                oDb.Execute szsql
        End If

    Next i
    oDb.CommitTrans
   
    Set oParam = Nothing
    Set rsTemp = Nothing
    Set oDb = Nothing
Exit Function
ErrHandle:
    oDb.RollbackTrans
    Set oParam = Nothing
    Set rsTemp = Nothing
    Set oDb = Nothing
    err.Raise err.Number
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

Public Function LocalExtraSellTicket(BusStartupDate() As Date, BusID() As String, SellStationID() As String, DesStationID() As String _
             , aspSellTicketParam() As TSellTicketParam, panInsurance() As Integer, Optional aszRealNameInfo As Variant) As TSellTicketResult()
    
    Dim oDb As New RTConnection
    Dim rsTemp() As TSellTicketResult
    Dim nCount As Integer
    Dim iLen As Integer
    iLen = ArrayLength(BusID)

    
    AssertActiveUserValid m_oActiveUser, ERR_SellTicketService
    AssertHaveRight m_oActiveUser, RIGHT_ExtraSellTicket
    ReDim rsTemp(1 To iLen)

    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    
    oDb.BeginTrans
    On Error GoTo Here
    
    Dim nStart As Integer
    For nCount = 1 To iLen
        rsTemp(nCount) = SelfSellTicketLocalTrans(BusStartupDate(nCount), BusID(nCount), SellStationID(nCount), DesStationID(nCount), aspSellTicketParam(nCount).BuyTicketInfo, aspSellTicketParam(nCount).pasgSellTicketPrice, panInsurance(nCount), True, ST_TicketReturned, oDb, nStart, aszRealNameInfo)
        nStart = nStart + ArrayLength(aspSellTicketParam(nCount).BuyTicketInfo)
    Next nCount
    LocalExtraSellTicket = rsTemp
    oDb.CommitTrans
    Set oDb = Nothing
    Exit Function
Here:
    oDb.RollbackTrans
    Set oDb = Nothing
    
    err.Raise err.Number
End Function



'取消保险费
Public Sub CancelInsurance(paszTicketID() As String)
    Dim oDb As New RTConnection
    Dim szsql As String
    Dim i As Integer
    Dim nLen As Integer
    Dim szWhere As String

    AssertActiveUserValid m_oActiveUser, RIGHT_ForceReturnTicket
    
    nLen = ArrayLength(paszTicketID)
    If nLen = 0 Then Exit Sub

    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    
    szWhere = " "
    For i = 1 To nLen - 1
        szWhere = szWhere & TransFieldValueToString(paszTicketID(i)) & ","
    Next i
    szWhere = szWhere & TransFieldValueToString(paszTicketID(i))
    

    
    If Trim(m_oCurrentSellService.SelfUser.ServerUnit) <> Trim(m_oCurrentSellService.SelfUser.UserUnitID) Then
    '如果用户单位和服务单位不一样则是售远程票否则售的是本地票
        
        m_oCurrentSellService.CancelInsurance paszTicketID
        
        
        szsql = " UPDATE ticket_sell_lst SET insurance = 0 WHERE  ticket_id IN ( " & szWhere & " ) "
        oDb.Execute szsql
        
    Else
        oDb.BeginTrans
            szsql = " UPDATE ticket_sell_lst SET insurance = 0 WHERE  ticket_id IN ( " & szWhere & " ) "
            oDb.Execute szsql
            
            szsql = " UPDATE agent_ticket_sale_lst SET insurance = 0 WHERE  ticket_id IN ( " & szWhere & " ) "
            oDb.Execute szsql
            
            
        oDb.CommitTrans
        
    End If
    On Error GoTo 0
    
    Exit Sub
Here:
    oDb.RollbackTrans
    err.Raise err.Number
End Sub

Public Function IsAllReturn() As Boolean
    Dim szStatus As Boolean
    AssertActiveUserValid m_oActiveUser, ERR_SellTicketService
    Dim szsql As String
    Dim rsTemp As Recordset
    Dim oDb As New RTConnection
    Dim szRight As String
    
    szRight = GetRightID(RIGHT_AllMoneyReturn)
    oDb.ConnectionString = GetConnectionStr
    '查询用户方法表中用户是否有该权限
    szsql = "SELECT function_id FROM " & cszUserFunction & " WHERE " _
    & " user_id='" & m_oActiveUser.UserID & "' AND " _
    & " function_id='" & szRight & "'"
    
    Set rsTemp = oDb.Execute(szsql)
    If rsTemp.RecordCount = 0 Then
        '如果用户没有权限则查询用户组是否有权限
        szsql = "SELECT function_id FROM " & cszUsergroupFunction & "  tbu, " & cszGroupUser & " tbg WHERE tbu.usergroup_id=" _
        & " tbg.usergroup_id AND user_id='" & m_oActiveUser.UserID & "' AND function_id='" & szRight & "'"
        Set rsTemp = oDb.Execute(szsql)
        If rsTemp.RecordCount <> 0 Then
            szStatus = True
        Else
            szStatus = False
        End If
    Else
        szStatus = True
    End If
    IsAllReturn = szStatus
End Function

Public Function IsForceSell() As Boolean
    Dim szStatus As Boolean
    AssertActiveUserValid m_oActiveUser, ERR_SellTicketService
    Dim szsql As String
    Dim rsTemp As Recordset
    Dim oDb As New RTConnection
    Dim szRight As String
    
    szRight = GetRightID(RIGHT_ForceSell)
    oDb.ConnectionString = GetConnectionStr
    '查询用户方法表中用户是否有该权限
    szsql = "SELECT function_id FROM " & cszUserFunction & " WHERE " _
    & " user_id='" & m_oActiveUser.UserID & "' AND " _
    & " function_id='" & szRight & "'"
    
    Set rsTemp = oDb.Execute(szsql)
    If rsTemp.RecordCount = 0 Then
        '如果用户没有权限则查询用户组是否有权限
        szsql = "SELECT function_id FROM " & cszUsergroupFunction & "  tbu, " & cszGroupUser & " tbg WHERE tbu.usergroup_id=" _
        & " tbg.usergroup_id AND user_id='" & m_oActiveUser.UserID & "' AND function_id='" & szRight & "'"
        Set rsTemp = oDb.Execute(szsql)
        If rsTemp.RecordCount <> 0 Then
            szStatus = True
        Else
            szStatus = False
        End If
    Else
        szStatus = True
    End If
    IsForceSell = szStatus
End Function

Public Function ChangeTicketType(pszTicketID As String, pszTicketTypeID As String) As Boolean
    AssertActiveUserValid m_oActiveUser, ERR_SellTicketService
    Dim szsql As String
    Dim rsTemp As Recordset
    Dim rsTemp1 As Recordset
    Dim oDb As New RTConnection
    Dim nSpecialTicketTypeCount As Integer
    Dim rsSelledSpecialTicketCount As Recordset
    If pszTicketTypeID = 1 Then
        nSpecialTicketTypeCount = 0
    Else
        nSpecialTicketTypeCount = 1
    End If
    
    oDb.ConnectionString = GetConnectionStr
    
    Dim m_bSpecialTicketTypePosition As String
    Dim m_bSpecialTicketTypeRatio As Double
    m_bSpecialTicketTypePosition = m_oParam.SpecialTicketTypePosition
    m_bSpecialTicketTypeRatio = m_oParam.SpecialTicketTypeRatio
    
         szsql = "SELECT * FROM Agent_ticket_sale_lst WHERE ticket_id='" & pszTicketID & "'"
         Set rsTemp = oDb.Execute(szsql)
         
         If rsTemp.RecordCount = 1 Then
         
            szsql = "SELECT e.bus_date , e.bus_id , e.vehicle_type_name , e.vehicle_type_code , e.total_seat " _
                & " , e.sale_seat_quantity , e.total_stand_quantity , e.sale_stand_seat_quantity , e.bus_start_time , e.bus_type" _
                & " , e.register_status , e.is_all_refundment , e.route_id , e.end_station_id , e.end_station_name" _
                & " , e.owner_id , e.transport_company_id , e.vehicle_id , e.split_company_id , e.baggage_number , e.fact_weight" _
                & " , e.calculate_weight , e.over_weight_number , e.luggage_total_price , e.internet_status , e.scrollbus_check_time" _
                & " , e.seat_remain , e.bed_remain  , e.additional_remain , e.other_remain_1 , e.other_remain_2" _
                & " , s.sell_station_id , s.station_id , s.seat_type_id , s.station_serial_no , s.max_sale_quantity , s.sale_ticket_quantity" _
                & " , s.book_count , s.stop_sale_time , s.mileage  , s.full_price , s.half_price , s.preferential_ticket1 , s.preferential_ticket2" _
                & " , s.preferential_ticket3" _
                & " , a.check_gate_id , a.bus_start_time" _
                & " , a.status , a.can_sale_quantity , a.sell_quantity" _
                & " FROM work_env_bus_info e ,work_env_bus_station_lst s , work_env_bus_allot_lst a" _
                & " WHERE e.bus_id='" & FormatDbValue(rsTemp!bus_id) & "' AND  e.bus_date='" & ToDBDate(rsTemp!bus_date) & "'" _
                & " AND  e.bus_id=s.bus_id AND  e.bus_date=s.bus_date" _
                & " AND  s.station_id='" & FormatDbValue(rsTemp!des_station_id) & "' AND s.sell_station_id='" & FormatDbValue(rsTemp!sell_station_id) & "'" _
                & " AND a.bus_id = s.bus_id AND a.bus_date = s.bus_date AND a.sell_station_id = '" & FormatDbValue(rsTemp!sell_station_id) & "'"
            Set rsTemp1 = oDb.Execute(szsql)
            If rsTemp1.RecordCount = 0 Then ShowError ERR_NoThisBus '没有指定的车次或车次不过指定的到站
         
             szsql = "SELECT count(*) as selled_special_ticket_count From Agent_ticket_sale_lst " _
               & " WHERE bus_id='" & FormatDbValue(rsTemp!bus_id) & "' AND  bus_date='" & ToDBDate(rsTemp!bus_date) & "'" _
               & " AND  status in ('" & ST_TicketNormal & "','" & ST_TicketNormal + ST_TicketChecked & "','" & ST_TicketSellChange & "', '" & ST_TicketSellChange + ST_TicketChecked & "' ) AND ticket_type='" & m_bSpecialTicketTypePosition & "' "
             Set rsSelledSpecialTicketCount = oDb.Execute(szsql)
            
             If (FormatDbValue(rsSelledSpecialTicketCount!selled_special_ticket_count) + nSpecialTicketTypeCount) > Int(m_bSpecialTicketTypeRatio * FormatDbValue(rsTemp1!total_seat) / 100) And nSpecialTicketTypeCount > 0 Then
                ChangeTicketType = False
'                ShowError ERR_CanNotSellSpecialTicket
                Exit Function
             Else
                szsql = "UPDATE Agent_ticket_sale_lst SET ticket_type='" & pszTicketTypeID & "'" _
                      & " WHERE ticket_id= '" & pszTicketID & "'"
                oDb.Execute szsql
                ChangeTicketType = True
             End If
         End If
                
         szsql = "SELECT * FROM ticket_sell_lst WHERE ticket_id='" & pszTicketID & "'"
         Set rsTemp = oDb.Execute(szsql)
         
         If rsTemp.RecordCount = 1 Then
         
            szsql = "SELECT e.bus_date , e.bus_id , e.vehicle_type_name , e.vehicle_type_code , e.total_seat " _
                & " , e.sale_seat_quantity , e.total_stand_quantity , e.sale_stand_seat_quantity , e.bus_start_time , e.bus_type" _
                & " , e.register_status , e.is_all_refundment , e.route_id , e.end_station_id , e.end_station_name" _
                & " , e.owner_id , e.transport_company_id , e.vehicle_id , e.split_company_id , e.baggage_number , e.fact_weight" _
                & " , e.calculate_weight , e.over_weight_number , e.luggage_total_price , e.internet_status , e.scrollbus_check_time" _
                & " , e.seat_remain , e.bed_remain  , e.additional_remain , e.other_remain_1 , e.other_remain_2" _
                & " , s.sell_station_id , s.station_id , s.seat_type_id , s.station_serial_no , s.max_sale_quantity , s.sale_ticket_quantity" _
                & " , s.book_count , s.stop_sale_time , s.mileage  , s.full_price , s.half_price , s.preferential_ticket1 , s.preferential_ticket2" _
                & " , s.preferential_ticket3" _
                & " , a.check_gate_id , a.bus_start_time" _
                & " , a.status , a.can_sale_quantity , a.sell_quantity" _
                & " FROM work_env_bus_info e ,work_env_bus_station_lst s , work_env_bus_allot_lst a" _
                & " WHERE e.bus_id='" & FormatDbValue(rsTemp!bus_id) & "' AND  e.bus_date='" & ToDBDate(rsTemp!bus_date) & "'" _
                & " AND  e.bus_id=s.bus_id AND  e.bus_date=s.bus_date" _
                & " AND  s.station_id='" & FormatDbValue(rsTemp!des_station_id) & "' AND s.sell_station_id='" & FormatDbValue(rsTemp!sell_station_id) & "'" _
                & " AND a.bus_id = s.bus_id AND a.bus_date = s.bus_date AND a.sell_station_id = '" & FormatDbValue(rsTemp!sell_station_id) & "'"
            Set rsTemp1 = oDb.Execute(szsql)
            If rsTemp1.RecordCount = 0 Then ShowError ERR_NoThisBus '没有指定的车次或车次不过指定的到站
         
             szsql = "SELECT count(*) as selled_special_ticket_count From ticket_sell_lst " _
               & " WHERE bus_id='" & FormatDbValue(rsTemp!bus_id) & "' AND  bus_date='" & ToDBDate(rsTemp!bus_date) & "'" _
               & " AND  status in ('" & ST_TicketNormal & "','" & ST_TicketNormal + ST_TicketChecked & "', '" & ST_TicketSellChange & "', '" & ST_TicketSellChange + ST_TicketChecked & "' ) AND ticket_type='" & m_bSpecialTicketTypePosition & "' "
             Set rsSelledSpecialTicketCount = oDb.Execute(szsql)
            
             If (FormatDbValue(rsSelledSpecialTicketCount!selled_special_ticket_count) + nSpecialTicketTypeCount) > Int(m_bSpecialTicketTypeRatio * FormatDbValue(rsTemp1!total_seat) / 100) And nSpecialTicketTypeCount > 0 Then
                ChangeTicketType = False
'                ShowError ERR_CanNotSellSpecialTicket
                Exit Function
             Else
                szsql = "UPDATE ticket_sell_lst SET ticket_type='" & pszTicketTypeID & "'" _
                      & " WHERE ticket_id= '" & pszTicketID & "'"
                oDb.Execute szsql
                ChangeTicketType = True
             End If
         End If
         
         '如已检票，修改检票记录
         szsql = "SELECT * FROM check_ticket_lst WHERE ticket_id='" & pszTicketID & "'"
         Set rsTemp = oDb.Execute(szsql)
         
         If rsTemp.RecordCount = 1 Then
            szsql = "UPDATE check_ticket_lst SET ticket_type='" & pszTicketTypeID & "'" _
                  & " WHERE ticket_id= '" & pszTicketID & "'"
            oDb.Execute szsql
            ChangeTicketType = True
         End If
         
End Function



