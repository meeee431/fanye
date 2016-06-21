VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClientTicket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public Enum EErrClientTicket
     ERR_ClientTicketNotExist = ERR_ClientTicket + 10 '此车票不存在
'    ERR_TicketNotCancel = ERR_ClientTicket + 10 '此车票还未作废
'    ERR_TicketNotChange = ERR_ClientTicket + 11 '此车票不是改签售出的
End Enum
Private m_nObjectStatus As EObjectStatus
Private m_oActiveUser As ActiveUser
Private m_szTicketID As String '票号
Private m_dtREBusDate As Date '运行车次日期
Private m_szREBusID As String '运行车次代码
Private m_szUnitID As String '所属的单位代码
Private m_szStartStationID As String '始发站ID
Private m_szStartStaionName As String '始发站名称
Private m_szToStationID As String '到站ID
Private m_szToStationName As String '到站名称
Private m_szOperator As String '售票员
Private m_szSeatNo As String '座号
Private m_sgTicketPrice As Double '票价
Private m_dtSellTime As Date '售票时间
Private m_tCanceledInfo As TCanceledTicketInfo '废票信息
Private m_tChangedInfo As TChangedTicketInfo '改签信息
Private m_nTicketStatus As ETicketStatus '票种
Private m_nTicketType As ETicketType '车票的状态
Private m_szSellStationID As String '售票站ID
Private m_szSellStationName As String '售票站名称

Private m_nInsurance As Integer '保险费

Public Property Let TicketType(vData As ETicketType)
    m_nTicketType = vData
End Property


Public Property Get TicketType() As ETicketType
    TicketType = m_nTicketType
End Property



Public Property Let TicketStatus(vData As ETicketStatus)
    m_nTicketStatus = vData
End Property


Public Property Get TicketStatus() As ETicketStatus
    TicketStatus = m_nTicketStatus
End Property





Public Property Let ChangedInfo(vData As TChangedTicketInfo)
    m_tChangedInfo = vData
End Property


Public Property Get ChangedInfo() As TChangedTicketInfo
    ChangedInfo = m_tChangedInfo
End Property


Public Property Let CanceledInfo(vData As TCanceledTicketInfo)
    m_tCanceledInfo = vData
End Property


Public Property Get CanceledInfo() As TCanceledTicketInfo
    CanceledInfo = m_tCanceledInfo
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





Public Property Let ToStationName(ByVal vData As String)
    m_szToStationName = vData
End Property


Public Property Get ToStationName() As String
    ToStationName = m_szToStationName
End Property



Public Property Let ToStationID(ByVal vData As String)
    m_szToStationID = vData
End Property


Public Property Get ToStationID() As String
    ToStationID = m_szToStationID
End Property



Public Property Let StartStaionName(ByVal vData As String)
    m_szStartStaionName = vData
End Property


Public Property Get StartStaionName() As String
    StartStaionName = m_szStartStaionName
End Property



Public Property Let StartStationID(ByVal vData As String)
    m_szStartStationID = vData
End Property


Public Property Get StartStationID() As String
    StartStationID = m_szStartStationID
End Property


Public Property Let UnitID(ByVal vData As String)
    m_szUnitID = vData
End Property


Public Property Get UnitID() As String
    UnitID = m_szUnitID
End Property



Public Property Let REBusID(ByVal vData As String)
    m_szREBusID = vData
End Property


Public Property Get REBusID() As String
    REBusID = m_szREBusID
End Property



Public Property Let REBusDate(ByVal vData As Date)
    m_dtREBusDate = vData
End Property


Public Property Get REBusDate() As Date
    AssertObjIsValid
    REBusDate = m_dtREBusDate
End Property



'--------------------------->>
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


Private Sub Class_Initialize()
    m_nObjectStatus = ST_NotAvailable
End Sub
Private Sub AssertStatusAvailable()
If m_nObjectStatus = ST_NotAvailable Then ShowError ERR_NotAvailable
End Sub
Public Sub AssertObjIsValid()
    AssertActiveUserValid m_oActiveUser, ERR_ClientTicket
    AssertStatusAvailable
End Sub


Public Sub Identify(TicketID As String)
    AssertActiveUserValid m_oActiveUser, ERR_ClientTicket
    
    m_szTicketID = TicketID
    RefreshMemoryInfo
    
    m_nObjectStatus = ST_NormalObj
End Sub






Private Sub RefreshMemoryInfo() '按照对象现在的ID将数据库中相应的信息读至对象内
    Dim rsTemp As Recordset
    Dim oDb As New RTConnection
    Dim szSql As String
    AssertActiveUserValid m_oActiveUser, ERR_ClientTicket
    oDb.ConnectionString = GetConnectionStr(cszSellTicket)
    '得到票信息
    szSql = "SELECT a.*,s.sell_station_name FROM Agent_ticket_sale_lst a,sell_station_info s WHERE a.sell_station_id=s.sell_station_id AND ticket_id='" & m_szTicketID & "'"
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 1 Then
        m_dtREBusDate = FormatDbValue(rsTemp!bus_date)
        m_szREBusID = FormatDbValue(rsTemp!bus_id)
        m_szUnitID = FormatDbValue(rsTemp!service_unit_id)
        m_szStartStationID = FormatDbValue(rsTemp!start_station_id)
        m_szStartStaionName = FormatDbValue(rsTemp!start_station_name)
        m_szToStationID = FormatDbValue(rsTemp!des_station_id)
        m_szToStationName = FormatDbValue(rsTemp!des_station_name)
        m_nTicketStatus = FormatDbValue(rsTemp!Status)
        m_nTicketType = FormatDbValue(rsTemp!ticket_type)
        m_szOperator = FormatDbValue(rsTemp!user_id)
        m_szSeatNo = FormatDbValue(rsTemp!seat_no)
        m_sgTicketPrice = FormatDbValue(rsTemp!ticket_price)
        m_dtSellTime = FormatDbValue(rsTemp!operation_time)
        m_szSellStationID = FormatDbValue(rsTemp!sell_station_id) '增加售票站代码
        m_szSellStationName = FormatDbValue(rsTemp!sell_station_name) '增加售票站名称
        m_nInsurance = FormatDbValue(rsTemp!Insurance) '保险费
    Else
        ShowError ERR_ClientTicketNotExist
    End If
    '如果票已废,则得到废票信息
    If m_nTicketStatus And ST_TicketCanceled Then
        szSql = "SELECT * FROM Agent_ticket_cancel_lst WHERE ticket_id='" & m_szTicketID & "'"
        Set rsTemp = oDb.Execute(szSql)
        If rsTemp.RecordCount = 1 Then
            m_tCanceledInfo.dtCancelTime = FormatDbValue(rsTemp!cancel_time)
            m_tCanceledInfo.nCancelType = FormatDbValue(rsTemp!cancel_mode)
            m_tCanceledInfo.szOperatorID = FormatDbValue(rsTemp!user_id)
        End If
    End If
    '如果票已改签,则得到改签票信息
    If m_nTicketStatus And ST_TicketChanged Then
        szSql = "SELECT * FROM Agent_ticket_change_lst WHERE new_ticket_id='" & m_szTicketID & "'"
        Set rsTemp = oDb.Execute(szSql)
        If rsTemp.RecordCount = 1 Then
            m_tChangedInfo.dtBusID = FormatDbValue(rsTemp!former_bus_id)
            m_tChangedInfo.dtDate = FormatDbValue(rsTemp!change_time)
            'm_tChangedInfo.nChangeType = FormatDbValue(rsTemp!change_mode)
            m_tChangedInfo.nTicketType = FormatDbValue(rsTemp!former_ticket_type)
            m_tChangedInfo.sgChangeCharge = FormatDbValue(rsTemp!change_charge)
            m_tChangedInfo.sgTicketPrice = FormatDbValue(rsTemp!former_ticket_price)
            m_tChangedInfo.szBeginStationID = FormatDbValue(rsTemp!former_start_station_id)
            m_tChangedInfo.szBeginStationName = FormatDbValue(rsTemp!former_start_station_name)
            m_tChangedInfo.szCredenceID = FormatDbValue(rsTemp!credence_sheet_id)
            m_tChangedInfo.szEndStationID = FormatDbValue(rsTemp!former_des_station_id)
            m_tChangedInfo.szEndStationName = FormatDbValue(rsTemp!former_des_station_name)
            m_tChangedInfo.szSeatNo = FormatDbValue(rsTemp!former_seat_no)
            'm_tChangedInfo.szTicketId = FormatDbValue(rsTemp!former_ticket_id)
        End If
    End If
    
End Sub
'返回上车站代码 2002/11/13 wjb
Public Property Get SellStationID() As String
       SellStationID = m_szSellStationID
End Property

Public Property Let SellStationID(ByVal vNewValue As String)
       m_szSellStationID = vNewValue
End Property
'返回上车站名称 2002/11/13 wjb
Public Property Get SellStationName() As String
       SellStationName = m_szSellStationName
End Property


'保险费
Public Property Get Insurance() As Integer
    Insurance = m_nInsurance
    
End Property



