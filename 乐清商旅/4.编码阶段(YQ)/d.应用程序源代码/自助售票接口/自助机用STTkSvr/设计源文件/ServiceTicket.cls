VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ServiceTicket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Enum EErrServiceTicket
    ERR_ServiceTicketNotExist = ERR_ServiceTicket + 10 '无此车票
    ERR_ServiceTicketNotReturn = ERR_ServiceTicket + 11 '此票还未被退
'    ERR_ServiceTicketPriceNotChanged = ERR_ServiceTicket + 12 '此票不是改签售出的
    ERR_ServiceTicketNotCancel = ERR_ServiceTicket + 13 '此票还未被废
    ERR_ServiceTicketNotCheck = ERR_ServiceTicket + 14 '此票还未被检
'    ERR_ServiceTicketPriceNotExist = ERR_ServiceTicket + 15 '无此票的票价
End Enum

Private m_nObjectStatus As EObjectStatus

Private m_oActiveUser As ActiveUser
Private m_szTicketID As String '票号
Private m_szREBusID As String '车次ID
Private m_dtREBusDate As Date '车次日期
Private m_szToStation As String '到站ID
Private m_nTicketStatus As ETicketStatus '票的状态
Private m_nTicketType As ETicketType '票的类型
Private m_szOperator As String '售票员
Private m_szSeatNo As String '票的座号
Private m_sgTicketPrice As Double '该票的总票价
Private m_dtSellTime As Date '该票的售票时间
Private m_sgReturnRatio As Double '要退该票所需的退票费率
Private m_tReturnedInfo As TReturnedTicketInfo '该票的退票相关信息
Private m_tChangedInfo As TChangedTicketInfo '该票的改签费（如此票已是改签售出票）
Private m_tCanceledInfo As TCanceledTicketInfo '该票的废票住处（如此票已废）
Private m_tCheckedInfo As TCheckedTicketInfo '该票的检票信息（如此票已检）
Private m_tTicketPriceInfo As TRETicketPrice '该票的票价明细
Private m_szBeChangedToTicket As String '返回该票被改签后的票号
Private m_szToStationName As String '到站的站名
Private m_dtBusStartUpTime As Date '车次发车时间
Private m_dblReturnTime As Double  '流水车次退票时间段
Private m_nBusType As EBusType
Private m_szSellStationID As String '售票站ID
Private m_szSellStationName As String '售票站名称

'Private m_szRouteID As String '线路代码
'Private m_szRouteName As String  '线路名称
'Private m_szVehicleTypeID As String '车型代码
'Private m_szVehicleTypeName As String '车型名称


Private m_nInsurance As Integer '保险费


Public Property Let ToStationName(ByVal vData As String)
    m_szToStationName = vData
End Property


Public Property Get ToStationName() As String
    ToStationName = m_szToStationName
End Property



Public Property Let BeChanedToTicket(ByVal vData As String)
    m_szBeChangedToTicket = vData
End Property


Public Property Get BeChanedToTicket() As String
    BeChanedToTicket = m_szBeChangedToTicket
End Property



Public Property Let TicketPriceInfo(vData As TRETicketPrice)
    m_tTicketPriceInfo = vData
End Property


Public Property Get TicketPriceInfo() As TRETicketPrice
     TicketPriceInfo = m_tTicketPriceInfo
End Property



Public Property Let CheckedInfo(vData As TCheckedTicketInfo)
    m_tCheckedInfo = vData
End Property


Public Property Get CheckedInfo() As TCheckedTicketInfo
    CheckedInfo = m_tCheckedInfo
End Property



Public Property Let CanceledInfo(vData As TCanceledTicketInfo)
     m_tCanceledInfo = vData
End Property


Public Property Get CanceledInfo() As TCanceledTicketInfo
     CanceledInfo = m_tCanceledInfo
End Property



Public Property Let ChangedInfo(vData As TChangedTicketInfo)
     m_tChangedInfo = vData
End Property


Public Property Get ChangedInfo() As TChangedTicketInfo
     ChangedInfo = m_tChangedInfo
End Property

Public Property Let ReturnedInfo(vData As TReturnedTicketInfo)
     m_tReturnedInfo = vData
End Property


Public Property Get ReturnedInfo() As TReturnedTicketInfo
     ReturnedInfo = m_tReturnedInfo
End Property

Public Property Get ReturnRatio() As Double
    ReturnRatio = m_sgReturnRatio
End Property

Public Property Let SellTime(ByVal vData As Date)
    m_dtSellTime = vData
End Property


Public Property Get SellTime() As Date
    SellTime = m_dtSellTime
End Property



Public Property Let TicketPrice(ByVal vData As Double)
    m_sgTicketPrice = vData
End Property


Public Property Get TicketPrice() As Double
    TicketPrice = m_sgTicketPrice
End Property



Public Property Let SeatNo(ByVal vData As String)
    m_szSeatNo = vData
End Property


Public Property Get SeatNo() As String
    SeatNo = m_szSeatNo
End Property



Public Property Let Operator(ByVal vData As String)
    m_szOperator = vData
End Property


Public Property Get Operator() As String
    Operator = m_szOperator
End Property



Public Property Let TicketType(ByVal vData As ETicketType)
     m_nTicketType = vData
End Property


Public Property Get TicketType() As ETicketType
     TicketType = m_nTicketType
End Property



Public Property Let TicketStatus(ByVal vData As ETicketStatus)
     m_nTicketStatus = vData
End Property


Public Property Get TicketStatus() As ETicketStatus
     TicketStatus = m_nTicketStatus
End Property



Public Property Let ToStation(ByVal vData As String)
    m_szToStation = vData
End Property


Public Property Get ToStation() As String
    ToStation = m_szToStation
End Property



Public Property Let REBusDate(ByVal vData As Date)
    m_dtREBusDate = vData
End Property


Public Property Get REBusDate() As Date
     REBusDate = m_dtREBusDate
End Property



Public Property Let REBusID(ByVal vData As String)
    m_szREBusID = vData
End Property


Public Property Get REBusID() As String
    REBusID = m_szREBusID
End Property

Public Property Let ReturnTime(ByVal vData As Double)
    ReturnTime = vData
End Property

Public Property Get ReturnTime() As Double
    ReturnTime = m_dblReturnTime
End Property



'Public Property Let RouteID(ByVal vData As String)
'    m_szRouteID = vData
'End Property
'
'
'Public Property Get RouteID() As String
'    RouteID = m_szRouteID
'End Property
'
'Public Property Let RouteName(ByVal vData As String)
'    m_szRouteName = vData
'End Property
'
'Public Property Get RouteName() As String
'    RouteName = m_szRouteName
'End Property
'
'Public Property Let VehicleTypeCode(ByVal vData As String)
'    m_szVehicleTypeID = vData
'End Property
'
'Public Property Get VehicleTypeCode() As String
'    VehicleTypeCode = m_szVehicleTypeID
'End Property
'
'Public Property Let VehicleTypeName(ByVal vData As String)
'    m_szVehicleTypeName = vData
'End Property
'
'Public Property Get VehicleTypeName() As String
'    VehicleTypeName = m_szVehicleTypeName
'End Property
'

Public Sub Init(poAUser As ActiveUser)
    Set SelfUser = poAUser
End Sub

Public Property Get SelfUser() As ActiveUser
    Set SelfUser = m_oActiveUser
    m_oActiveUser.IAmLive
End Property

Public Property Set SelfUser(vNewValue As ActiveUser)
    If m_nObjectStatus = ST_AddObj Then
        ShowError ERR_AddObj
    ElseIf m_nObjectStatus = ST_EditObj Then
        ShowError ERR_EditObj
    ElseIf m_nObjectStatus = ERR_NormalObj Then
        ShowError ERR_NormalObj
    Else
        Set m_oActiveUser = vNewValue
    End If
    m_oActiveUser.IAmLive
End Property


Private Sub AssertStatusAvailable()
    If m_nObjectStatus = ST_NotAvailable Then ShowError ERR_NotAvailable
End Sub
Public Sub AssertObjIsValid()
    AssertActiveUserValid m_oActiveUser, ERR_ServiceTicket
    AssertStatusAvailable
End Sub


Public Sub Identify(TicketID As String)
    AssertActiveUserValid m_oActiveUser, ERR_ServiceTicket
    
    m_szTicketID = TicketID
    RefreshMemoryInfo
    
    m_nObjectStatus = ST_NormalObj
End Sub


Public Property Get ObjStatus() As EObjectStatus
    ObjStatus = m_nObjectStatus
End Property
Public Property Get dtBusStartUpTime() As Date
    dtBusStartUpTime = m_dtBusStartUpTime
End Property

Private Sub RefreshMemoryInfo() '按照对象现在的ID将数据库中相应的信息读至对象内

    Dim rsTemp As Recordset
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim m_oParam As New SystemParam
    
    AssertActiveUserValid m_oActiveUser, ERR_ServiceTicket

    oDb.ConnectionString = GetConnectionStr(cszSellTicket)

    '得到该票的明细信息
    szSql = "SELECT s.*,a.bus_start_time,e.is_all_refundment,e.bus_type,sn.station_name,ssi.sell_station_name " _
        & " FROM Ticket_sell_lst s,Work_env_bus_info e,station_info sn,sell_station_info ssi , work_env_bus_allot_lst a " _
        & " WHERE s.ticket_id='" & m_szTicketID & "' AND s.bus_id=e.bus_id AND s.bus_date=e.bus_date " _
        & " AND s.des_station_id=sn.station_id  AND s.sell_station_id=ssi.sell_station_id " _
        & " AND a.bus_date = s.bus_date AND a.bus_id = s.bus_id AND a.sell_station_id= s.sell_station_id "
        
        ', e.route_id , e.vehicle_type_code, r.route_name, v.vehicle_type_short_name
        ' ,route_info r , vehicle_type_code v
        ' AND r.route_id = e.route_id AND v.vehicle_type_code = e.vehicle_type_code
        
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount <> 1 Then
        ShowError ERR_ServiceTicketNotExist
        Exit Sub
    End If
    m_szREBusID = FormatDbValue(rsTemp!bus_id)
    m_dtREBusDate = FormatDbValue(rsTemp!bus_date)
    m_dtBusStartUpTime = FormatDbValue(rsTemp!bus_start_time)
    m_szToStation = FormatDbValue(rsTemp!des_station_id)
    m_szToStationName = FormatDbValue(rsTemp!station_name)
    m_nTicketStatus = FormatDbValue(rsTemp!Status)
    m_nTicketType = FormatDbValue(rsTemp!ticket_type)
    m_szOperator = FormatDbValue(rsTemp!user_id)
    m_szSeatNo = FormatDbValue(rsTemp!seat_no)
    m_sgTicketPrice = FormatDbValue(rsTemp!ticket_price)
    m_dtSellTime = FormatDbValue(rsTemp!operation_time)
    m_nBusType = FormatDbValue(rsTemp!bus_type)
    m_szSellStationID = FormatDbValue(rsTemp!sell_station_id) '增加售票站代码
    m_szSellStationName = FormatDbValue(rsTemp!sell_station_name) '增加售票站名称

    
    m_nInsurance = FormatDbValue(rsTemp!Insurance) '保险费
    
    m_oParam.Init m_oActiveUser
    m_dblReturnTime = m_oParam.ScrollBusReturnTime
    
    If rsTemp.RecordCount = 1 And rsTemp!is_all_refundment <> "0" Then
        m_sgReturnRatio = 0
    Else
'        If rsTemp!bus_type = TP_ScrollBus Then
'            If m_oParam.ScrollBusCanReturnTicket Then
'                m_sgReturnRatio = FormatMoney(m_oParam.ScrollBusReturnRatio * 100)
'            Else
'                m_sgReturnRatio = 100
'            End If
'        Else
            m_sgReturnRatio = m_oParam.GetReturnChargeRatio(rsTemp!bus_start_time)
'        End If
    End If
    
    If m_nTicketStatus And ST_TicketSellChange Then  '改签售出票
        '如果车票为改签售出票,则得到该票的明细改签信息
        szSql = "SELECT * FROM ticket_change_lst c,Ticket_sell_lst s " & _
                "WHERE  c.former_ticket_id=s.ticket_id " & _
                "AND c.new_ticket_id='" & m_szTicketID & "'"
        Set rsTemp = oDb.Execute(szSql)
        If rsTemp.RecordCount = 1 Then
            m_tChangedInfo.dtBusID = FormatDbValue(rsTemp!bus_id)
            m_tChangedInfo.dtDate = FormatDbValue(rsTemp!bus_date)
            m_tChangedInfo.nTicketType = FormatDbValue(rsTemp!ticket_type)
            m_tChangedInfo.sgChangeCharge = FormatDbValue(rsTemp!change_charge)
            m_tChangedInfo.sgTicketPrice = FormatDbValue(rsTemp!former_ticket_price)
            m_tChangedInfo.szCredenceID = FormatDbValue(rsTemp!credence_sheet_id)
            m_tChangedInfo.szSeatNo = FormatDbValue(rsTemp!seat_no)
            m_tChangedInfo.szTicketId = FormatDbValue(rsTemp!ticket_id)
            m_tChangedInfo.szBeginStationID = Trim(m_oParam.StationID)
            m_tChangedInfo.szEndStationID = Trim(FormatDbValue(rsTemp!des_station_id))
        
            szSql = "SELECT * FROM station_info " & _
                    "WHERE station_id='" & m_tChangedInfo.szEndStationID & "' OR station_id='" & m_tChangedInfo.szBeginStationID & "'"
            Set rsTemp = oDb.Execute(szSql)
            While Not rsTemp.EOF
                If Trim(FormatDbValue(rsTemp!station_id)) = m_tChangedInfo.szBeginStationID Then
                    m_tChangedInfo.szBeginStationName = FormatDbValue(rsTemp!station_name)
                End If
                If Trim(FormatDbValue(rsTemp!station_id)) = m_tChangedInfo.szEndStationID Then
                    m_tChangedInfo.szEndStationName = FormatDbValue(rsTemp!station_name)
                End If
                rsTemp.MoveNext
            Wend
        End If
    End If
    If m_nTicketStatus And ST_TicketChanged Then  '被改签票
        '如果车票被改签
        szSql = "SELECT * FROM ticket_change_lst WHERE former_ticket_id='" & m_szTicketID & "'"
        Set rsTemp = oDb.Execute(szSql)
        If rsTemp.RecordCount > 0 Then
            m_szBeChangedToTicket = FormatDbValue(rsTemp!new_ticket_id)
        End If
    End If
    If m_nTicketStatus And ST_TicketCanceled Then   '废票
        szSql = "SELECT * FROM ticket_cancel_lst WHERE ticket_id='" & m_szTicketID & "'"
        Set rsTemp = oDb.Execute(szSql)
        If rsTemp.RecordCount = 1 Then
            m_tCanceledInfo.dtCancelTime = FormatDbValue(rsTemp!cancel_ticket_time)
            m_tCanceledInfo.nCancelType = FormatDbValue(rsTemp!cancel_mode)
            m_tCanceledInfo.szOperatorID = FormatDbValue(rsTemp!user_id)
        End If
    End If
    If m_nTicketStatus And ST_TicketReturned Then    '已退票
        szSql = "SELECT * FROM Ticket_return_lst WHERE ticket_id='" & m_szTicketID & "'"
        Set rsTemp = oDb.Execute(szSql)
        If rsTemp.RecordCount = 1 Then
            m_tReturnedInfo.dtReturnTime = FormatDbValue(rsTemp!operation_time)
            m_tReturnedInfo.szCredenceID = FormatDbValue(rsTemp!credence_sheet_id)
            m_tReturnedInfo.sgReturnCharge = FormatDbValue(rsTemp!return_charge)
            m_tReturnedInfo.szOperatorID = FormatDbValue(rsTemp!user_id)
            m_tReturnedInfo.nReturnType = FormatDbValue(rsTemp!return_ticket_mode)
        End If
    End If
    If m_nTicketStatus And ST_TicketChecked Then     '已检票
        szSql = "SELECT * FROM check_ticket_lst WHERE ticket_id='" & m_szTicketID & "'"
        Set rsTemp = oDb.Execute(szSql)
        If rsTemp.RecordCount = 1 Then
            m_tCheckedInfo.dtBusDate = FormatDbValue(rsTemp!bus_date)
            m_tCheckedInfo.dtCheckTime = FormatDbValue(rsTemp!check_time)
            m_tCheckedInfo.szTicketId = FormatDbValue(rsTemp!ticket_id)
            m_tCheckedInfo.szBusID = FormatDbValue(rsTemp!bus_id)
            m_tCheckedInfo.nBusSerialNo = FormatDbValue(rsTemp!bus_serial_no)
            m_tCheckedInfo.nCheckTicketType = FormatDbValue(rsTemp!Status)
        End If
    End If
    
    szSql = "SELECT * FROM ticket_price_lst WHERE ticket_price_lst.ticket_id='" & m_szTicketID & "'"
    Set rsTemp = oDb.Execute(szSql)
    
    Dim k As Integer
    If rsTemp.RecordCount = 1 Then
        m_tTicketPriceInfo.sgBase = FormatDbValue(rsTemp!base_price)
        For k = 1 To 15
            m_tTicketPriceInfo.asgPrice(k) = rsTemp("price_item_" & k)
        Next
        
        m_tTicketPriceInfo.nTicketType = m_nTicketType
        m_tTicketPriceInfo.szStationID = m_szToStation
    End If

End Sub

Public Property Get ReturnCharge() As Double
    Dim oParam As New SystemParam
    
    AssertObjIsValid
    oParam.Init m_oActiveUser
    If m_nBusType <> TP_ScrollBus Then
        If ReturnRatio > 0 Then
            ReturnCharge = oParam.GetReturnCharge(m_dtBusStartUpTime, m_sgTicketPrice)
        Else
            ReturnCharge = 0
        End If
    Else
        If oParam.ScrollBusCanReturnTicket Then
            ReturnCharge = oParam.GetReturnCharge(DateAdd("n", m_dblReturnTime * 60, m_dtSellTime), m_sgTicketPrice)
        Else
            ReturnCharge = m_sgTicketPrice
        End If
    End If
End Property
'//////////////////////////////
'得到时间间隔时间


'返回上车站代码
Public Property Get SellStationID() As String
       SellStationID = m_szSellStationID
End Property

Public Property Let SellStationID(ByVal vNewValue As String)
       m_szSellStationID = vNewValue
End Property
'返回上车站名称
Public Property Get SellStationName() As String
       SellStationName = m_szSellStationName
End Property

'保险费
Public Property Get Insurance() As Integer
    Insurance = m_nInsurance
    
End Property


