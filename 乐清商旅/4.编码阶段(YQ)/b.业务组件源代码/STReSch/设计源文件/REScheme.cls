VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "REScheme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
  
Public Enum EErrREScheme
    ERR_BusHaveTicketSold = ERR_REScheme + 10 '19210指定的车次已有售票
    ERR_BusHaveBeStopped = ERR_REScheme + 11 '19211车次已经停班
    ERR_RESchemeNotThisBus = ERR_REScheme + 12 '19212无指定的车次
    ERR_BusNotRunVehicle = ERR_REScheme + 13 '19213车次在指定的日期无执行车辆错
    ERR_VehicleHaveBeStopped = ERR_REScheme + 14 '19214车辆已经停班
    ERR_BusNoStation = ERR_REScheme + 15 '19215车次没有一个过站
    ERR_RoutePriceInfoErr = ERR_REScheme + 16 '19216线路票价表不全(无
    ERR_NoRoutePriceTable = ERR_REScheme + 17 '19217车次无票价表
    ERR_NotStationTicketTypePrice = ERR_REScheme + 18 '19218没有相应的站点，票种的票价
    ERR_NoExecuteProject = ERR_REScheme + 19 '19219没有执行车次计划
    ERR_DeleteEBusFailed = ERR_REScheme + 20 '19220删除环境车次失败(应该是车次已经并有其它车次)
    ERR_NoOnePlanNoCountStopBus = ERR_REScheme + 21 '19221不是同一计划的车次，不能统计停班车次量
    ERR_SeatCountISError = ERR_REScheme + 22
    ERR_NoPriceSeatTypeItem = ERR_REScheme + 23 ' 票价表中是无该座位类型的票价.
    ERR_NotStationPrice = ERR_REScheme + 25 '票价表中站点有信息,但常调站点表中没有相应的站点信息
    ERR_Haveluggage = ERR_REScheme + 26 '车次已受理行包
    ERR_NotBusSeatCount = ERR_REScheme + 27 '已售票车次重新生成环境中，不能完全分配已售座位，请增加车辆的可售座位数。
    ERR_SplitBus = ERR_REScheme + 28 '指定车次被拆分或有其她车次并入该车次，不能生成环境．
    Err_NOPriceOrStionID = ERR_REScheme + 29 '车次无票价或线路更改而计划车次没同步
    ERR_BusHaveTicketInternetSold = ERR_REScheme + 30 '19230指定的车次已有网上售票
End Enum

'******下面的变量声明应该删除
Public szSellStationID As String
Public szSellStationName As String
'******
Public Enum ERightREScheme
    RIGHT_MakeRunEvironment = ERR_REScheme + cnMidRightBegin + cnMidRightStep * 1 '生成运行环境
End Enum
Private m_oParam As SystemParam
'所有的类都有的私有类变量
Private m_oActiveUser As ActiveUser
'prviate m_tSeatTypeInfo
#If SUPPORT_MSG = 1 Then
    Private m_oSndMsg As New CSendMsg
    Private m_bOffTimeChanged As Boolean
    Private m_bCheckGateChanged As Boolean
#End If
'
Public Sub Init(poAUser As ActiveUser)
    Set SelfUser = poAUser
End Sub

Public Property Get SelfUser() As ActiveUser
    Set SelfUser = m_oActiveUser
    m_oActiveUser.IAmLive
End Property

Public Property Set SelfUser(vNewValue As ActiveUser)
    Set m_oActiveUser = vNewValue
    If m_oParam Is Nothing Then
        Set m_oParam = New SystemParam
    End If
    m_oParam.Init m_oActiveUser
    m_oActiveUser.IAmLive
End Property


''得到指定日期、车次的执行车辆的序号
Public Function GetExecuteVehicleSerialNo(ByVal RunCycle As Integer, ByVal CycleStartSerialNo As Integer, Optional ByVal RunDate As Date = cdtEmptyDate) As Integer
    Dim saTemp As TSchemeArrangement
    Dim dtDate As Date, LerrCode As Long
    If RunDate = cdtEmptyDate Then
        dtDate = m_oParam.NowDate
    Else
        dtDate = RunDate
    End If
'    saTemp = SelfGetExecuteBusProject(dtDate, LerrCode)
    If LerrCode <> 0 Then ShowError ERR_NoExecuteProject
'    GetExecuteVehicleSerialNo = (DateDiff("d", Date, dtDate) + CycleStartSerialNo - 1) Mod RunCycle + 1
    
    
    GetExecuteVehicleSerialNo = (DateDiff("d", "2003-01-02", dtDate) Mod RunCycle + CycleStartSerialNo - 1) Mod RunCycle + 1
    
'    GetExecuteVehicleSerialNo = (DateDiff("d", saTemp.dtBeginDate, dtDate) Mod RunCycle + CycleStartSerialNo - 1) Mod RunCycle + 1
End Function


Public Function GetBusEx(RunDate As Date, EndDate As Date) As String()
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim szWhere As String
    Dim szWHERE2 As String
    Dim rsTemp As Recordset
    Dim i As Integer
    Dim szTemp() As String
    Dim nCountInfo As Integer
    
     On Error GoTo here
'    szSql = " SELECT bus_id,bus_start_time,end_station_name,vehicle_type_name " _
'    & " FROM Work_env_bus_info " _
'    & " WHERE bus_date >= '" & RunDate & "' AND " _
'    & " bus_date <= '" & EndDate & "'"
    ''bus_id,bus_start_time,end_station_name,vehicle_type_name
    If DateDiff("d", RunDate, Now) <= 0 Then '如果查询的是环境中的车次则使用工作表
        szSql = " SELECT * FROM Work_env_bus_info " _
        & "WHERE bus_date >= '" & RunDate & "' AND " _
        & " bus_date <= '" & EndDate & "'"
    Else
        szSql = " SELECT * FROM Env_bus_info " _
        & "WHERE bus_date >= '" & RunDate & "' AND " _
        & " bus_date <= '" & EndDate & "'"
    End If
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    Set rsTemp = oDb.Execute(szSql)
    nCountInfo = rsTemp.RecordCount
    If nCountInfo <> 0 Then
    ReDim szTemp(0 To nCountInfo - 1, 0 To 3)
        For i = 1 To rsTemp.RecordCount
    
            szTemp(i, 0) = FormatDbValue(rsTemp!bus_id)
            If FormatDbValue(rsTemp!bus_type) <> TP_ScrollBus Then
                szTemp(i, 1) = FormatDbValue(rsTemp!bus_start_time)
            Else
                szTemp(i, 1) = "滚动车次"
            End If
            szTemp(i, 2) = FormatDbValue(rsTemp!end_station_name)
            szTemp(i, 3) = FormatDbValue(rsTemp!vehicle_type_name)
        Next
    End If
    GetBusEx = szTemp
    If Not rsTemp Is Nothing Then Set rsTemp = Nothing
   Set oDb = Nothing
here:
    err.Raise err.Number
    Set oDb = Nothing
    If Not rsTemp Is Nothing Then Set rsTemp = Nothing
End Function
Public Function GetBus(RunDate As Date, Optional LikeBusID As String, Optional RouteID As String, Optional DescStation As String, Optional bflg As Boolean, Optional SellStation As String, Optional EqualBusID As String) As String()
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim szWhere As String
    Dim szWHERE2 As String
    Dim rsTemp As Recordset
    Dim i As Integer
    Dim aszTemp() As String
    
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    If LikeBusID <> "" Then
        szWhere = " AND e.bus_id LIKE '" & LikeBusID & "%' "
    End If
       
    If RouteID <> "" Then
        szWhere = szWhere & " AND r.route_id LIKE '" & RouteID & "%' "
    End If
    
    If EqualBusID <> "" Then
        szWhere = szWhere & " AND e.bus_id = '" & EqualBusID & "'"
    End If
    
    If DescStation = "" Then
        If SellStation = "" Then
            szSql = " SELECT v.license_tag_no,e.*,r.route_name,c.transport_company_short_name ,bt.bus_type_name ,c2.transport_company_short_name split_company_short_name ,dbo.GetSaledSeatCount(e.bus_date,e.bus_id) saled_seat_count ,dbo.BookNums(e.bus_date,e.bus_id) book_count " _
                & " FROM " & IIf(DateDiff("d", RunDate, Now) <= 0, "Work_env_bus_info", "Env_bus_info")
                
            szSql = szSql & " e,route_info r,Vehicle_info v,Company_info c ,bus_type_code bt,Company_info c2 " _
                & " WHERE bus_date='" & ToDBDate(RunDate) & "' AND e.route_id=r.route_id AND v.vehicle_id=e.vehicle_id " _
                & " AND  c.transport_company_id = e.transport_company_id " _
                & " AND  c2.transport_company_id  = e.split_company_id " _
                & " and bt.bus_type=e.bus_type"
                
            szSql = szSql & szWhere & " order by e.bus_date ,e.bus_start_time ,e.Bus_id ,e.bus_type"
        Else
            szSql = " SELECT v.license_tag_no,e.*,r.route_name,c.transport_company_short_name ,bt.bus_type_name ,c2.transport_company_short_name split_company_short_name ,dbo.GetSaledSeatCount(e.bus_date,e.bus_id) saled_seat_count ,dbo.BookNums(e.bus_date,e.bus_id) book_count " _
                & " FROM " & IIf(DateDiff("d", RunDate, Now) <= 0, "Work_env_bus_info", "Env_bus_info")
                
            szSql = szSql & " e,Work_env_bus_allot_lst l,route_info r,Vehicle_info v,Company_info c ,bus_type_code bt,Company_info c2 " _
                & " WHERE e.bus_date='" & ToDBDate(RunDate) & "' AND e.route_id=r.route_id AND v.vehicle_id=e.vehicle_id " _
                & " AND c.transport_company_id = e.transport_company_id " _
                & " AND c2.transport_company_id  = e.split_company_id " _
                & " AND bt.bus_type=e.bus_type" _
                & " AND e.bus_date=l.bus_date" _
                & " AND e.bus_id=l.bus_id" _
                & " AND l.sell_station_id = " & TransFieldValueToString(SellStation)
                
            szSql = szSql & szWhere & " order by e.bus_date ,e.bus_start_time ,e.Bus_id ,e.bus_type"
        End If
    Else
    
        If bflg = True Then
           '可售站点
           szWHERE2 = " and  bs.max_sale_quantity<>0"
        End If
        
        szSql = "SELECT DISTINCT v.license_tag_no,e.*,r.route_name,c.transport_company_short_name ," _
            & " c2.transport_company_short_name split_company_short_name ,bt.bus_type_name ,dbo.GetSaledSeatCount(e.bus_date,e.bus_id) saled_seat_count ,dbo.BookNums(e.bus_date,e.bus_id) book_count" _
            & " FROM " & IIf(DateDiff("d", RunDate, Now) <= 0, "Work_env_bus_station_lst bs,Work_env_bus_info e", "env_bus_station_lst bs,Env_bus_info e")
        szSql = szSql & " ,route_info r, " _
            & " Vehicle_info v,station_info ,Company_info c ,Company_info c2 ,bus_type_code bt" _
            & " WHERE  e.bus_date='" & ToDBDate(RunDate) & "' AND e.route_id=r.route_id AND v.vehicle_id=e.vehicle_id AND bs.station_id=station_info.station_id " _
            & " AND  e.bus_id=bs.bus_id AND  e.bus_date=bs.bus_date AND  c.transport_company_id = e.transport_company_id "
        szSql = szSql & szWHERE2 _
            & " AND  c2.transport_company_id  = e.split_company_id and bt.bus_type=e.bus_type" _
            & " AND  (station_info.station_name='" & Trim(DescStation) & "' " _
            & " OR station_info.station_input_code LIKE '" & DescStation & "%' OR station_info.station_id='" & DescStation & "') " & szWhere & " order by e.bus_date ,e.bus_start_time ,e.Bus_id ,e.bus_type"
            '& " OR station_info.station_input_code='" & DescStation & "' OR station_info.station_id='" & DescStation & "') " & szWhere & " order by e.bus_date ,e.bus_start_time ,e.Bus_id ,e.bus_type"
    End If
    Set rsTemp = oDb.Execute(szSql)
    
    If rsTemp.RecordCount > 0 Then
        ReDim aszTemp(1 To rsTemp.RecordCount, 1 To 17) As String
        For i = 1 To rsTemp.RecordCount
            aszTemp(i, 1) = FormatDbValue(rsTemp!bus_id)
            If FormatDbValue(rsTemp!bus_type) <> TP_ScrollBus Then
                aszTemp(i, 2) = FormatDbValue(rsTemp!bus_start_time)
            Else
                aszTemp(i, 2) = "滚动车次"
            End If
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
            aszTemp(i, 16) = FormatDbValue(rsTemp!saled_seat_count)
            aszTemp(i, 17) = FormatDbValue(rsTemp!book_count)
            rsTemp.MoveNext
        Next
    GetBus = aszTemp
    End If
End Function

Public Function GetReBusBookAndReserveSeatInfo(RunDate As Date, Optional LikeBusID As String, Optional RouteID As String, Optional DescStation As String, Optional bflg As Boolean) As String()
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim szWhere As String
    Dim szWHERE2 As String
    Dim rsTemp As Recordset
    Dim i As Integer
    Dim szFROM As String
    Dim szFROM2 As String
    Dim aszTemp() As String
    Dim szWHEREBus As String
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    
    If LikeBusID <> "" Then
        szWhere = "  AND e.bus_id LIKE '" & LikeBusID & "%' "
        szWHEREBus = szWhere
    End If
       
    If RouteID <> "" Then
        szWhere = szWhere & "  AND r.route_id LIKE '" & RouteID & "%' "
    End If
    szSql = "SELECT "
    szFROM = IIf(DateDiff("d", RunDate, Now) <= 0, "Work_env_bus_info", "Env_bus_info") & " e"

   Dim rsSeatInfo As Recordset
   Dim j As Integer
   Dim nSeatCountInfo As Integer
  '取得基本信息
   szSql = " SELECT e.bus_id,e.bus_start_time,e.total_seat,e.sale_seat_quantity," _
              & " r.route_name,st.station_name,e.vehicle_type_name ," _
              & " e.status,ISNULL(p.seat_book_Count,0)as Book_Count" _
              & " FROM " & szFROM & ",route_info r, station_info st ,project_book_info  p WHERE  " _
              & "  e.bus_date='" & ToDBDate(RunDate) & "' " _
              & szWhere _
              & "  AND e.route_id=r.route_id " _
              & " and  e.end_station_id=st.station_id " _
              & " and p.start_bus_date<='" & ToDBDate(RunDate) & "' and p.end_bus_date>='" & ToDBDate(RunDate) & "' " _
              & "  and e.bus_id*=p.Bus_id" _
              & " order by e.bus_start_time ,e.Bus_id "
   szSql = szSql
  Set rsTemp = oDb.Execute(szSql)
  '取得座位信息
  szSql = "SELECT e.bus_id,e.bus_start_time,s.status,seat_no FROM Env_bus_seat_lst s," & szFROM _
               & " WHERE  s.bus_id= e.bus_id  and e.bus_date='" & ToDBDate(RunDate) & "' " _
               & " and  s.bus_date=e.bus_date " & szWHEREBus & " order by e.bus_start_time ,e.bus_id,s.status"
  Set rsSeatInfo = oDb.Execute(szSql)
  '记录集处理
  If rsTemp.RecordCount > 0 Then
    ReDim aszTemp(0 To rsTemp.RecordCount - 1, 0 To 12) As String
    For i = 0 To rsTemp.RecordCount - 1
            aszTemp(i, 0) = rsTemp!bus_id
            aszTemp(i, 1) = rsTemp!bus_start_time
            aszTemp(i, 2) = rsTemp!total_seat
            aszTemp(i, 3) = rsTemp!sale_seat_quantity
            aszTemp(i, 8) = rsTemp!route_name
            aszTemp(i, 9) = rsTemp!station_name
            aszTemp(i, 10) = rsTemp!vehicle_type_name
            aszTemp(i, 11) = rsTemp!Status
            
            aszTemp(i, 5) = 0
            aszTemp(i, 6) = 0
            aszTemp(i, 7) = 0
            aszTemp(i, 12) = 0
            
            rsSeatInfo.Find "bus_id='" & aszTemp(i, 0) & "' "
            
            If rsSeatInfo.EOF = False Then
                nSeatCountInfo = 0
                Do While Trim(rsSeatInfo!bus_id) = Trim(aszTemp(i, 0))
                   nSeatCountInfo = nSeatCountInfo + 1
                   Select Case CInt(rsSeatInfo!Status)
                          Case 64  '- --计划预订信息
                             aszTemp(i, 5) = aszTemp(i, 5) + 1
                          Case 1  ' 环境预留信息
                             aszTemp(i, 6) = aszTemp(i, 6) + 1
                          Case 2  '环境预订信息
                             aszTemp(i, 7) = aszTemp(i, 7) + 1
                          Case 3, 4, 5 ' 已售票
                             aszTemp(i, 12) = aszTemp(i, 12) + 1
                          Case 0
                  End Select
                  
                  If nSeatCountInfo < CInt(aszTemp(i, 2)) Then
                      rsSeatInfo.MoveNext
                  Else
                      Exit Do
                  End If
                Loop
                rsSeatInfo.MovePrevious
           End If
           
           aszTemp(i, 4) = CInt(rsTemp!book_count)
           
           rsTemp.MoveNext
           
     Next
  End If

   GetReBusBookAndReserveSeatInfo = aszTemp
  '  End If
End Function





'断言指定的车次无售票
Private Sub AssertBusHaveNotSellTicket(poDb As RTConnection, pdtBusDate As Date, pszBusID As String)
    Dim szSql As String
    Dim rsTemp As Recordset
    szSql = "SELECT COUNT(*) AS countx FROM Ticket_sell_lst WHERE " _
    & " bus_id='" & pszBusID & "' AND " _
    & " bus_date='" & ToDBDate(pdtBusDate) & "'"
    Set rsTemp = poDb.Execute(szSql)
    If FormatDbValue(rsTemp!countx) > 0 Then
        ShowError ERR_BusHaveTicketSold
    End If
End Sub


Private Function SelfGetBusStatus(pdtBeginStopDate As Date, pdtEndStopDate As Date, pdtRunDate As Date, pbStopContinueMake As Boolean) As EREBusStatus
    SelfGetBusStatus = ST_BusNormal
    '如果用户设定开始日期为1900年01月01日
    'If Not DBDateIsEmpty(pdtBeginStopDate) Then     '如果开始停班日期为空就不停班
        'If Not DBDateIsEmpty(pdtBeginStopDate) Then
            If pdtRunDate >= pdtBeginStopDate And pdtRunDate <= pdtEndStopDate Then
                SelfGetBusStatus = ST_BusStopped
                If Not pbStopContinueMake Then ShowError ERR_BusHaveBeStopped
            End If
        'Else
        '    If pdtRunDate >= pdtBeginStopDate Then
        '        SelfGetBusStatus = ST_BusStopped
        '        If Not pbStopContinueMake Then ShowError ERR_BusHaveBeStopped
        '    End If
        'End If
    'End If
End Function

'内部用将指定的环境车次的信息删除
Private Function DeleteREBus(busdate As Date, BusID As String, Optional poDb As Object = Nothing) As Integer
    Dim szSql As String
    Dim oDb As RTConnection
    If poDb Is Nothing Then
        Set oDb = New RTConnection
        oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    Else
        Set oDb = poDb
    End If
    
    szSql = "DELETE Env_bus_info WHERE bus_id='" & BusID & "' AND bus_date='" & ToDBDate(busdate) & "'"
    
    On Error GoTo here
    
    oDb.Execute szSql
    
    Exit Function
here:
    ShowError ERR_DeleteEBusFailed
End Function


'内部用将线路票价信息的记录集移合适的位置
Private Function FindRightPriceItem(prsPriceInfo As Recordset, ByVal pszStationID As String, ByVal pnTicketType As ETicketType, Optional pbFROMFirst As Boolean = False) As Integer
    Dim szMsg As String
    Dim m_lTemp As Long
    
    On Error GoTo Error_Handle
    If pbFROMFirst Then
        prsPriceInfo.MoveFirst
    Else
        If prsPriceInfo.Bookmark <> 1 Then prsPriceInfo.MovePrevious
    End If
    
    Do While Not prsPriceInfo.EOF
        If FormatDbValue(prsPriceInfo!station_id) = pszStationID And FormatDbValue(prsPriceInfo!ticket_type) = pnTicketType Then
            Exit Function
        End If
        prsPriceInfo.MoveNext
    Loop
Error_Handle:
    szMsg = "站点" & pszStationID & "票种信息"
'    If pnTicketType = TP_FullPrice Then
'        szMsg = szMsg & "全票"
'    Else
'        szMsg = szMsg & "半票"
'    End If
    ShowError ERR_NotStationTicketTypePrice, , szMsg
End Function

Public Function GetREBusReport(StartDate As Date, EndDate As Date) As Recordset
    Dim oDb As New RTConnection
    Dim szSql As String
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    szSql = "SELECT tbe.*,tbv.license_tag_no,tbc.transport_company_short_name,tbo.owner_name,tbr.route_name " _
            & " ,tbc2.transport_company_short_name split_company_short_name " _
            & " FROM Env_bus_info tbe,route_info tbr,Vehicle_info " _
            & " tbv,Company_info tbc ,owner_info tbo ,Company_info tbc2 " _
            & " WHERE bus_date>='" & ToDBDate(StartDate) & "' AND bus_date<='" & ToDBDate(EndDate) & "' " _
            & " AND tbe.route_id=tbr.route_id AND tbv.vehicle_id=tbe.vehicle_id " _
            & " AND tbc.transport_company_id=tbe.transport_company_id AND tbo.owner_id=tbe.owner_id" _
            & " AND tbc2.transport_company_id=tbe.split_company_id "
            
    Set GetREBusReport = oDb.Execute(szSql)
End Function

Public Function GetREBusStopReport(StartDate As Date, EndDate As Date) As Recordset
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    Dim rsStopBus As New Recordset
    Dim szSql As String
    Dim rsVehicle As Recordset
    Dim i As Integer, nCount As Integer
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    rsStopBus.Fields.Append "bus_id", adChar, 5
    rsStopBus.Fields.Append "route_id", adChar, 4
    rsStopBus.Fields.Append "route_name", adChar, 16
    rsStopBus.Fields.Append "vehicle_id", adChar, 5
    rsStopBus.Fields.Append "company_name", adChar, 10
    rsStopBus.Fields.Append "owner_name", adChar, 8
    rsStopBus.Fields.Append "license_tag_no", adChar, 10
    rsStopBus.Fields.Append "stop_count", adInteger
    rsStopBus.Fields.Append "bus_start_time", adDate
    rsStopBus.Fields.Append "split_company_name", adChar, 10
    rsStopBus.Open
    szSql = "SELECT bus_id,tbe.route_id,vehicle_id,route_name,COUNT(*) AS stopcount,MAX(bus_start_time) AS bus_start_time" _
            & " FROM Env_bus_info tbe,route_info tbr WHERE tbe.route_id = tbr.route_id" _
            & " AND tbe.bus_date>='" & ToDBDate(StartDate) & "' AND tbe.bus_date<='" & ToDBDate(EndDate) & "'" _
            & " AND tbe.status IN (" & ST_BusStopped & "," & ST_BusMergeStopped & ") GROUP BY bus_id,tbe.route_id,vehicle_id,route_name "
    Set rsTemp = oDb.Execute(szSql)
    nCount = rsTemp.RecordCount
    If rsTemp.RecordCount = 0 Then Set GetREBusStopReport = rsStopBus: Exit Function
    For i = 1 To nCount
        rsStopBus.AddNew
        rsStopBus!bus_id = Trim(rsTemp!bus_id)
        rsStopBus!route_id = Trim(rsTemp!route_id)
        rsStopBus!route_name = Trim(rsTemp!route_name)
        rsStopBus!vehicle_id = Trim(rsTemp!vehicle_id)
        rsStopBus!stop_count = Val(rsTemp!stopcount)
        szSql = "SELECT license_tag_no,tbc.transport_company_short_name ,owner_name ,tbc2.transport_company_short_name split_company_short_name " _
              & " FROM Vehicle_info tbt,Company_info tbc,owner_info tbo ,Company_info tbc2" _
              & " WHERE tbt.transport_company_id=tbc.transport_company_id " _
              & " AND tbt.owner_id=tbo.owner_id AND vehicle_id='" & Trim(rsTemp!vehicle_id) & "'" _
              & " AND tbt.split_company_id=tbc2.transport_company_id"
         Set rsVehicle = oDb.Execute(szSql)
        rsStopBus!license_tag_no = Trim(rsVehicle!license_tag_no)
        rsStopBus!company_name = Trim(rsVehicle!transport_company_short_name)
        rsStopBus!owner_name = Trim(rsVehicle!owner_name)
        rsStopBus!bus_start_time = Trim(rsTemp!bus_start_time)
        rsStopBus!split_company_name = Trim(rsVehicle!split_company_short_name)
        rsStopBus.Update
        rsTemp.MoveNext
    Next
    rsStopBus.MoveFirst
    Set GetREBusStopReport = rsStopBus
End Function

'环境车次加班统计
Public Function GetREBusAddReport(StartDate As Date, EndDate As Date) As Recordset
    Dim oDb As New RTConnection
    Dim sysPar As New SystemParam
    Dim szAddPre As String
    Dim rsTemp As Recordset
    Dim rsAddBus As New Recordset
    Dim szSql As String
    Dim rsVehicle As Recordset
    Dim i As Integer, nCount As Integer
    sysPar.Init m_oActiveUser
    szAddPre = Left(sysPar.AdditionBusPreFix, 1)
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    rsAddBus.Fields.Append "bus_id", adChar, 5
    rsAddBus.Fields.Append "route_id", adChar, 4
    rsAddBus.Fields.Append "route_name", adChar, 16
    rsAddBus.Fields.Append "vehicle_id", adChar, 5
    rsAddBus.Fields.Append "company_name", adChar, 10
    rsAddBus.Fields.Append "owner_name", adChar, 8
    rsAddBus.Fields.Append "license_tag_no", adChar, 10
    rsAddBus.Fields.Append "add_count", adInteger
    rsAddBus.Fields.Append "bus_start_time", adDate
    rsAddBus.Fields.Append "split_company_name", adChar, 10
    rsAddBus.Open
    szSql = "SELECT bus_id,tbe.route_id,vehicle_id,route_name,COUNT(*) AS addcount,MAX(bus_start_time) AS bus_start_time" _
            & " FROM Env_bus_info tbe,route_info tbr WHERE tbe.route_id = tbr.route_id" _
            & " AND tbe.bus_date>='" & ToDBDate(StartDate) & "' AND tbe.bus_date<='" & ToDBDate(EndDate) & "'" _
            & " AND tbe.bus_id like '%" & szAddPre & "%' AND tbe.status Not IN (" & ST_BusStopped & "," & ST_BusMergeStopped & ") GROUP BY bus_id,tbe.route_id,vehicle_id,route_name "
    Set rsTemp = oDb.Execute(szSql)
    nCount = rsTemp.RecordCount
    If rsTemp.RecordCount = 0 Then Set GetREBusAddReport = rsAddBus: Exit Function
    For i = 1 To nCount
        rsAddBus.AddNew
        rsAddBus!bus_id = Trim(rsTemp!bus_id)
        rsAddBus!route_id = Trim(rsTemp!route_id)
        rsAddBus!route_name = Trim(rsTemp!route_name)
        rsAddBus!vehicle_id = Trim(rsTemp!vehicle_id)
        rsAddBus!add_count = Val(rsTemp!addcount)
        szSql = "SELECT license_tag_no,tbc.transport_company_short_name ,owner_name ,tbc2.transport_company_short_name split_company_short_name " _
              & " FROM Vehicle_info tbt,Company_info tbc,owner_info tbo ,Company_info tbc2" _
              & " WHERE tbt.transport_company_id=tbc.transport_company_id " _
              & " AND tbt.owner_id=tbo.owner_id AND vehicle_id='" & Trim(rsTemp!vehicle_id) & "'" _
              & " AND tbt.split_company_id=tbc2.transport_company_id"
         Set rsVehicle = oDb.Execute(szSql)
        rsAddBus!license_tag_no = Trim(rsVehicle!license_tag_no)
        rsAddBus!company_name = Trim(rsVehicle!transport_company_short_name)
        rsAddBus!owner_name = Trim(rsVehicle!owner_name)
        rsAddBus!bus_start_time = Trim(rsTemp!bus_start_time)
        rsAddBus!split_company_name = Trim(rsVehicle!split_company_short_name)
        rsAddBus.Update
        rsTemp.MoveNext
    Next
    rsAddBus.MoveFirst
    Set GetREBusAddReport = rsAddBus
End Function

Public Function GetPlanStopReport(StartDate As Date, EndDate As Date) As Recordset
    Dim oDb As New RTConnection
    Dim i As Integer, nCount As Integer
    Dim dtRunDate As Date
    Dim szSql As String
'    Dim tPlan As TSchemeArrangement, tTemp As TSchemeArrangement
    Dim rsTempP As Recordset, rsTempE As Recordset
    Dim vBookMark As Variant
    Dim plErrLong As Long
    Dim rsStopReport As New Recordset
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
'    tPlan = SelfGetExecuteBusProject(StartDate, plErrLong)
    nCount = DateDiff("d", StartDate, EndDate)
    For i = 1 To nCount
        dtRunDate = DateAdd("d", i, StartDate)
'        tTemp = SelfGetExecuteBusProject(dtRunDate, plErrLong)
'        If tPlan.szProjectID <> tTemp.szProjectID Then
'            ShowError ERR_NoOnePlanNoCountStopBus
'        End If
    Next
    rsStopReport.Fields.Append "bus_id", adChar, 6
    rsStopReport.Fields.Append "route_id", adChar, 4
    rsStopReport.Fields.Append "route_name", adChar, 16
    rsStopReport.Fields.Append "bus_start_time", adDate
    rsStopReport.Fields.Append "bus_run_cycle", adInteger
    rsStopReport.Fields.Append "stop_count", adInteger
    rsStopReport.Fields.Append "count", adInteger
    rsStopReport.Open
'    szSql = "SELECT bus_id,tbb.route_id,tbr.route_name,bus_start_time,bus_run_cycle FROM bus_info tbb,route_info tbr WHERE project_id='" & Trim(tPlan.szProjectID) & "' AND tbb.route_id=tbr.route_id"
    szSql = "SELECT bus_id,tbb.route_id,tbr.route_name,bus_start_time,bus_run_cycle " _
            & " FROM bus_info tbb,route_info tbr " _
            & " WHERE  tbb.route_id=tbr.route_id"
    
    Set rsTempP = oDb.Execute(szSql)
    szSql = "SELECT bus_id,COUNT(*) AS run_count FROM Env_bus_info WHERE bus_date>='" & ToDBDate(StartDate) & "' AND bus_date<='" & ToDBDate(EndDate) & "' GROUP BY bus_id"
    Set rsTempE = oDb.Execute(szSql)
    If rsTempE.RecordCount = 0 Then
         Set GetPlanStopReport = rsStopReport: Exit Function
    Else
        vBookMark = rsTempE.Bookmark
    End If
    nCount = nCount + 1
    For i = 1 To rsTempP.RecordCount
             rsStopReport.AddNew
             rsStopReport!bus_id = Trim(rsTempP!bus_id)
             rsStopReport!route_id = Trim(rsTempP!route_id)
             rsStopReport!route_name = Trim(rsTempP!route_name)
             rsStopReport!bus_start_time = FormatDbValue(rsTempP!bus_start_time)
             rsStopReport!bus_run_cycle = rsTempP!bus_run_cycle
             rsTempE.Bookmark = vBookMark
             rsTempE.Find ("bus_id='" & Trim(rsTempP!bus_id) & "'")
             If rsTempE.EOF Then
                rsStopReport!stop_count = nCount
             Else
                rsStopReport!stop_count = nCount - rsTempE!run_count
             End If
             rsStopReport!Count = nCount
             rsTempP.MoveNext
        Next
    rsStopReport.MoveFirst
    Set GetPlanStopReport = rsStopReport
End Function

'生成环境根据车次和日期
Public Sub MakeRunEvironment(ByVal pdtRunDate As Date, ByVal pszBusID As String, Optional ByVal pbStopContinueMake As Boolean = False)
    Dim oDb As New RTConnection
    Dim rsTemp As Recordset
    
    Dim LerrCode As Long
    Dim szSql As String
    Dim i As Integer
    
    Dim bEndStationCanSale As Boolean

    Dim szSellStationID As String
    
    
    Dim saTemp As TSchemeArrangement
    Dim nBusStatus As EREBusStatus
    Dim nCurrentBusvehicleNo As Integer '当前车次执行车辆序号
    Dim nBusRunCycle As Integer '车次运行周期
    Dim nBusRunFirstNo As Integer '车次运行周期的起始序号
    
    Dim nBeginSeatNo As Integer '车次车辆起始座位号
    Dim nBeginReserveSeatNo As Integer
    Dim nReserveSeatCount As Integer
    nReserveSeatCount = 0
    
    '用于插入环境表的变量声明
    Dim nInternetStatus As Integer
    Dim sac As Integer
    Dim nScrollBusCheckTime As Integer
    Dim nTotalSeat As Integer
    Dim nSaleSeatQuantity As Integer
    Dim nTotalStandQuantity As Integer
    Dim nSaleStandSeatQuantity As Integer
    Dim szCheckGateID As String
    Dim szVehicleTypeName As String
    Dim szVehicleTypeCode As String
    Dim dtOffTime As Date
    Dim nBusType As Integer
    Dim nRegisterStatus As Integer
    Dim nIsAllRefundment As Integer
    Dim szRouteID As String
    Dim szEndStationId As String
    Dim szEndstationName As String
    Dim szOwnerID As String
    Dim szTransportCompanyID As String
    Dim szVehicleID As String
    Dim szSplitCompanyID As String
    Dim tvTemp() As TVehcileSeatType
    Dim sysPar As New SystemParam
    Dim szAddBusFix  As String '加班车前缀
    Dim bFlgSaleBusCanMake As Boolean
    Dim nBaggageNumber As Integer
    Dim sgFactWeight As Single
    Dim sgCalculateWeight As Single
    Dim nOverWeightNumber As Integer
    Dim szEnvBookSeatNo As String  '环境预定的座位号
    
    Dim sgLuggageTotalPrice As Single
    
    Dim szaStationSale() As String
    Dim szaStationBook() As String '环境站点预定信息
    Dim szOldVehicleID As String
    Dim szOldVehicleTypeCode As String
    Dim bflgSaleBus As Boolean
    Dim szTicketPriceTable As String
    Dim i1 As Integer
    Debug.Print "生成开始:" & Time
    ''
    AssertActiveUserValid m_oActiveUser, ERR_REScheme
    AssertHaveRight m_oActiveUser, RIGHT_MakeRunEvironment
    
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    SetSysPar oDb, True
    
    sysPar.Init m_oActiveUser
    '加班车前缀
    szAddBusFix = sysPar.AdditionBusPreFix
    bEndStationCanSale = sysPar.EndStationCanSale
    bFlgSaleBusCanMake = sysPar.MakeEnviromentSaleBus
    pbStopContinueMake = sysPar.MekeStopEnviroment
    
    Set sysPar = Nothing
    
    '得到相应执行计划信息
'    saTemp = SelfGetExecuteBusProject(pdtRunDate, LerrCode)
    If LerrCode <> 0 Then ShowError ERR_NoExecuteProject
    
    
    '得到票价表
    szTicketPriceTable = BusProjectExecutePrice(ToDBDate(pdtRunDate), LerrCode)
    If LerrCode <> 0 Then ShowError ERR_NoRoutePriceTable
    
    
    '从售票明细表（sale_ticket_lst）中判断该车次是否有售票
    szSql = "SELECT * FROM Ticket_sell_lst WHERE bus_id='" & pszBusID & "' AND " _
        & "bus_date='" & ToDBDate(pdtRunDate) & "'"
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount > 0 Then
        If bFlgSaleBusCanMake = False Then
            ShowError ERR_BusHaveTicketSold
        Else
            bflgSaleBus = True  '车次有售票
        End If
    End If
    '判断该车次是否有预定票 在后面处理，改为默认自动填充 2008-03-10 zyw
'    szSql = "SELECT bus_id FROM ticket_book_lst WHERE bus_id='" & pszBusID & "' AND " _
'        & "bus_date='" & ToDBDate(pdtRunDate) & "' AND " _
'        & "status=" & SE_SeatBooked & ""
'    Set rsTemp = oDb.Execute(szSql)
'    If rsTemp.RecordCount <> 0 And bFlgSaleBusCanMake = False Then ShowError ERR_BusHaveTicketSold

    '从环境车次座位表（env_bus_seat_lst）中判断该车次是否有网上售票 fpd 2015-08-13
    szSql = "SELECT * FROM env_bus_seat_lst WHERE bus_id ='" & pszBusID & "' AND " _
        & "bus_date ='" & ToDBDate(pdtRunDate) & "'" _
        & " AND Status IN (" & ST_SeatReserved & ") AND remark LIKE '%网%'"
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount > 0 Then
        ShowError ERR_BusHaveTicketInternetSold
    End If

    '是否受理行包
    szSql = "SELECT * FROM Work_env_bus_info WHERE bus_id='" & pszBusID & "' AND " _
        & "bus_date='" & ToDBDate(pdtRunDate) & "'and baggage_number<>0 and fact_weight<> 0 and calculate_weight<>0 and over_weight_number<>0 and luggage_total_price<>0"
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount <> 0 And bFlgSaleBusCanMake = False Then ShowError ERR_Haveluggage
    
    If rsTemp.RecordCount <> 0 And bFlgSaleBusCanMake = True Then
        nBaggageNumber = rsTemp!baggage_number
        sgFactWeight = rsTemp!fact_weight
        sgCalculateWeight = rsTemp!calculate_weight
        nOverWeightNumber = rsTemp!over_weight_number
        sgLuggageTotalPrice = rsTemp!luggage_total_price
    End If
    
    '得到车次信息
'    szSql = "SELECT bi.* ,ri.* ,si.station_name AS EndStationName FROM bus_info bi,route_info ri ,station_info si WHERE bus_id='" & pszBusID & "' AND " _
'        & "project_id='" & saTemp.szProjectID & "' AND " _
'        & "bi.route_id=ri.route_id AND " _
'        & "ri.end_station_id=si.station_id"
        
     szSql = "SELECT bi.* ,ri.* ,si.station_name AS EndStationName " _
        & " FROM bus_info bi,route_info ri ,station_info si " _
        & " WHERE bus_id='" & pszBusID & "' AND " _
        & "bi.route_id=ri.route_id AND " _
        & "ri.end_station_id=si.station_id"
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount <> 1 Then ShowError ERR_RESchemeNotThisBus
    nBusRunCycle = rsTemp!bus_run_cycle
    nBusRunFirstNo = rsTemp!run_start_serial
    
    
    '得到车次状态
    nBusStatus = SelfGetBusStatus(rsTemp!stop_start_date, rsTemp!stop_end_date, pdtRunDate, pbStopContinueMake)
    szCheckGateID = rsTemp!check_gate_id
    nBusType = rsTemp!bus_type
    szRouteID = rsTemp!route_id
    dtOffTime = rsTemp!bus_start_time
    szEndStationId = rsTemp!end_station_id
    szEndstationName = rsTemp!EndStationName
    nScrollBusCheckTime = rsTemp!ScrollBus_Check_Time
    nInternetStatus = rsTemp!internet_status
    dtOffTime = Format(dtOffTime, cszTimeStr)
    dtOffTime = Format(pdtRunDate, cszDateStr) & " " & dtOffTime
    
    If nBusStatus = ST_BusStopped And pbStopContinueMake = False Then ShowError ERR_BusHaveBeStopped
    
'    GetExecuteVehicleSerialNo(nBusRunCycle, nBusRunFirstNo, pdtRunDate)   '
    
    '得到车次执行车辆序号
    nCurrentBusvehicleNo = GetExecuteVehicleSerialNo(nBusRunCycle, nBusRunFirstNo, pdtRunDate) ' (DateDiff("d", saTemp.dtBeginDate, pdtRunDate) Mod nBusRunCycle + nBusRunFirstNo - 1) Mod nBusRunCycle + 1
    '    nCurrentBusvehicleNo = DateDiff("d", saTemp.dtBeginDate, pdtRunDate) Mod nBusRunCycle + nBusRunFirstNo
    
    '得到车次执行车辆信息
    szSql = "SELECT bva.*,tvi.*,vtc.vehicle_type_short_name FROM bus_vehicle_code bva,Vehicle_info tvi,vehicle_type_code vtc WHERE bva.bus_id='" & pszBusID & "' AND " _
        & "bva.vehicle_serial=" & nCurrentBusvehicleNo & " AND " _
        & "bva.vehicle_id=tvi.vehicle_id AND " _
        & "tvi.vehicle_type_code=vtc.vehicle_type_code"
    Set rsTemp = oDb.Execute(szSql)
    
    If rsTemp.RecordCount <> 1 Then ShowError ERR_BusNotRunVehicle
    
    '判断车次车辆停班否
    If nBusStatus = ST_BusNormal Then
        nBusStatus = SelfGetBusStatus(FormatDbValue(rsTemp!stop_start_date), FormatDbValue(rsTemp!stop_end_date), pdtRunDate, pbStopContinueMake)
    End If
    
    '判断车辆停班否
    If nBusStatus = ST_BusNormal Then
        If FormatDbValue(rsTemp!Status) = ST_VehicleStop Then
            nBusStatus = ST_BusStopped
            If Not pbStopContinueMake Then ShowError ERR_VehicleHaveBeStopped
        End If
    End If
    
    nBeginSeatNo = FormatDbValue(rsTemp!start_seat_no)
    nTotalSeat = rsTemp!seat_quantity  '得到总座位数
    szVehicleID = rsTemp!vehicle_id
    nSaleStandSeatQuantity = rsTemp!sale_stand_ticket_quantity
    szOwnerID = rsTemp!owner_id
    szTransportCompanyID = rsTemp!transport_company_id
    szVehicleTypeCode = rsTemp!vehicle_type_code
    szVehicleTypeName = rsTemp!vehicle_type_short_name
    szSplitCompanyID = rsTemp!split_company_id
    
    
    
    '得到车辆座位类型信息
    Dim nCountBusVehicleSeat As Integer
    Dim btvTempNoSeatTypeIsNormal As Boolean
    Dim bHaveMormalSeatType As Boolean
    
    tvTemp = GetBusVehicleSeatType(szVehicleID)
    
    For i = 1 To ArrayLength(tvTemp)
        If Trim(tvTemp(i).szSeatTypeID) = cszSeatTypeIsNormal Then bHaveMormalSeatType = True
        
        If CInt(tvTemp(i).szEndSeatNo) > CInt(tvTemp(i).szStartSeatNo) Then
            nCountBusVehicleSeat = nCountBusVehicleSeat + CInt(tvTemp(i).szEndSeatNo) - CInt(tvTemp(i).szStartSeatNo) + 1
        Else
            nCountBusVehicleSeat = nCountBusVehicleSeat + CInt(tvTemp(i).szStartSeatNo) - CInt(tvTemp(i).szEndSeatNo) + 1
        End If
    
    Next
    
    If nCountBusVehicleSeat < nTotalSeat And bHaveMormalSeatType = False Then
        btvTempNoSeatTypeIsNormal = True
    Else
        btvTempNoSeatTypeIsNormal = False
    End If
    
    
    '检查是否有票价
    
    Dim rstempPrice As New Recordset
    
    
    szSql = " SELECT DISTINCT seat_type_id  FROM vehicle_seat_type_info WHERE vehicle_id='" & szVehicleID & "'"
    Set rsTemp = oDb.Execute(szSql)
    Dim rsSellStation As Recordset

    If rsTemp.RecordCount <> 0 Then

            For i = 1 To rsTemp.RecordCount
                szSql = "SELECT * FROM bus_price_lst WHERE bus_id='" & pszBusID & "' and price_table_id='" & szTicketPriceTable & " ' " _
                    & " AND vehicle_type_code='" & szVehicleTypeCode & "' and seat_type_id ='" & Trim(rsTemp!seat_type_id) & "'"
                Set rstempPrice = oDb.Execute(szSql)
                If rstempPrice.RecordCount = 0 Then ShowError ERR_NoRoutePriceTable
                rsTemp.MoveNext
                   
            Next
    End If
    If btvTempNoSeatTypeIsNormal = True Then
        szSql = "SELECT * FROM bus_price_lst WHERE bus_id='" & pszBusID & "' and price_table_id='" & szTicketPriceTable & " ' " _
            & "AND vehicle_type_code='" & szVehicleTypeCode & "' and seat_type_id ='" & cszSeatTypeIsNormal & "'"
        Set rstempPrice = oDb.Execute(szSql)
        If rstempPrice.RecordCount = 0 Then ShowError ERR_NoRoutePriceTable
    End If
    
    Set rstempPrice = Nothing
    
    '得到车次车辆座位预留信息
    szSql = "SELECT * FROM Bus_reserve_seat_lst WHERE bus_id='" & pszBusID & "' AND " _
          & "vehicle_serial=" & nCurrentBusvehicleNo & ""
    Set rsTemp = oDb.Execute(szSql)
    nBeginReserveSeatNo = -1
    
    If rsTemp.RecordCount <> 0 Then
        If SelfGetBusStatus(FormatDbValue(rsTemp!start_time), FormatDbValue(rsTemp!end_time), pdtRunDate, True) = ST_BusStopped Then
            nReserveSeatCount = FormatDbValue(rsTemp!seat_quantity)
            If nReserveSeatCount > 0 Then
                nBeginReserveSeatNo = FormatDbValue(rsTemp!start_seat_no)
            End If
        End If
    End If
    
    Debug.Print "得到生成必需信息:" & Time
    nSaleSeatQuantity = nTotalSeat - nReserveSeatCount
    
    '拆分车次不能重新生成环境
    szSql = "SELECT * FROM Mergebus_split_lst WHERE (bus_id='" & pszBusID & "' or merge_bus_id='" & pszBusID & "') and bus_date='" & ToDBDate(pdtRunDate) & "'"
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount <> 0 Then ShowError ERR_SplitBus
    
    If bFlgSaleBusCanMake = True Then
        szSql = "SELECT * FROM Env_bus_station_lst WHERE bus_id ='" & pszBusID & "' " _
            & " AND bus_date='" & ToDBDate(pdtRunDate) & " '" _
            & " and ( max_sale_quantity <> -1   " _
            & " OR sale_ticket_quantity<> -1 )"
        Set rsTemp = oDb.Execute(szSql)
        
        If rsTemp.RecordCount > 0 Then
            ReDim szaStationSale(1 To rsTemp.RecordCount, 1 To 5)
            For i = 1 To rsTemp.RecordCount
                szaStationSale(i, 1) = rsTemp!station_id
                szaStationSale(i, 2) = rsTemp!max_sale_quantity
                szaStationSale(i, 3) = rsTemp!sale_ticket_quantity
                szaStationSale(i, 4) = rsTemp!seat_type_id
                szaStationSale(i, 5) = rsTemp!stop_sale_time
                
                rsTemp.MoveNext
            Next
        End If
    End If
    
    '取得环境预定的站点信息
    szSql = "SELECT * FROM Env_bus_station_lst WHERE bus_id ='" & pszBusID & "' " _
        & " AND bus_date='" & ToDBDate(pdtRunDate) & " '" _
        & " AND book_count <> 0  "
    Set rsTemp = oDb.Execute(szSql)
    
    If rsTemp.RecordCount > 0 Then
        ReDim szaStationBook(1 To rsTemp.RecordCount, 1 To 4)
        For i = 1 To rsTemp.RecordCount
            szaStationBook(i, 1) = rsTemp!sell_station_id
            szaStationBook(i, 2) = rsTemp!station_id
            szaStationBook(i, 3) = rsTemp!seat_type_id
            szaStationBook(i, 4) = rsTemp!book_count
            
            rsTemp.MoveNext
        Next i
    End If
    
    
    '开始事务
    On Error GoTo here
    oDb.BeginTrans
    '删除指定环境车次
    If bFlgSaleBusCanMake = True And bflgSaleBus = True Then
        SetSysPar oDb
    End If

    DeleteREBus pdtRunDate, pszBusID, oDb
    
    '插入环境车次表信息（Env_bus_info）
    szSql = "INSERT Env_bus_info( " _
        & "bus_date," _
        & "bus_id," _
        & "check_gate_id," _
        & "vehicle_type_name," _
        & "vehicle_type_code," _
        & "total_seat," _
        & "sale_seat_quantity," _
        & "total_stand_quantity," _
        & "sale_stand_seat_quantity,"
    szSql = szSql & "bus_start_time," _
        & "bus_type," _
        & "status," _
        & "route_id," _
        & "end_station_id," _
        & "end_station_name," _
        & "owner_id," _
        & "transport_company_id," _
        & "split_company_id, " _
        & "vehicle_id ," _
        & "baggage_number ," _
        & "fact_weight," _
        & "calculate_weight ," _
        & "over_weight_number ," _
        & "luggage_total_price," _
        & "ScrollBus_Check_Time ," _
        & "internet_status) "
    
    szSql = szSql & " VALUES( '" _
        & ToDBDate(pdtRunDate) & "','" _
        & pszBusID & "','" _
        & szCheckGateID & "','" _
        & szVehicleTypeName & "','" _
        & szVehicleTypeCode & "'," _
        & nTotalSeat & "," _
        & nSaleSeatQuantity & "," _
        & nTotalStandQuantity & "," _
        & nTotalStandQuantity & ","
    szSql = szSql & "'" & ToDBDateTime(dtOffTime) & "'," _
        & nBusType & "," _
        & nBusStatus & ",'" _
        & szRouteID & "','" _
        & szEndStationId & "','" _
        & szEndstationName & "','" _
        & szOwnerID & "','" _
        & szTransportCompanyID & "','" _
        & szSplitCompanyID & "','" _
        & szVehicleID & "'," _
        & nBaggageNumber & "," _
        & sgFactWeight & "," _
        & sgCalculateWeight & "," _
        & nOverWeightNumber & "," _
        & sgLuggageTotalPrice & "," _
        & nScrollBusCheckTime & "," _
        & nInternetStatus & " )"
    oDb.Execute szSql
    
    Debug.Print "插入环境车次表信息:" & Time
    '插入环境座位表信息（Env_bus_seat_lst）
    If nBusType <> TP_ScrollBus Then
        Dim nTemp1 As Integer
        Dim nTemp2 As Integer
        Dim szSeatTypeID As String
        Dim nToalseatTemp As Integer
        Dim nSeatTypeSirel As Integer
        Dim tSeatInfoCountTemp(1 To 5) As TSeatInfoCount
        
        
        '''''''''计划预定'''''''''''
        Dim nProjectBookCount As Integer '计划预定数
        Dim bBookYes As Boolean '要处理计划预定
        Dim nProjectBookCountTemp As Integer
        Dim nBeginBookSeatNo As Integer
        Dim szProjectBusBookInfo() As String '计划预定信息
        szProjectBusBookInfo = GetProjectBookInfo(pdtRunDate, pszBusID)
    
        If ArrayLength(szProjectBusBookInfo) > 0 Then
            bBookYes = True
            nProjectBookCount = CInt(szProjectBusBookInfo(3))
        End If
        
        
        '环境预定的座位信息取得 2008-03-10 by zyw
        szSql = " SELECT seat_no FROM ticket_book_lst " _
              & " WHERE bus_date = '" & ToDBDate(pdtRunDate) & "'" _
              & " AND bus_id = '" & pszBusID & "'" _
              & " AND status = 0 "
        Set rsTemp = oDb.Execute(szSql)
        
        For i = 1 To rsTemp.RecordCount
            szEnvBookSeatNo = szEnvBookSeatNo & FormatDbValue(rsTemp!seat_no) & ","
            rsTemp.MoveNext
        Next i
        If szEnvBookSeatNo <> "" Then szEnvBookSeatNo = Left(szEnvBookSeatNo, Len(szEnvBookSeatNo) - 1)
        
        ''''''''''''''''''''
        '删除环境座位表相应的信息
        szSql = "DELETE FROM Env_bus_seat_lst WHERE bus_date='" & ToDBDate(pdtRunDate) & "' AND bus_id='" & pszBusID & "'"
        oDb.Execute szSql
        nToalseatTemp = 0
        For i = 1 To nTotalSeat
            nTemp1 = ST_SeatCanSell
            nTemp2 = i - 1 + nBeginSeatNo
            nToalseatTemp = nToalseatTemp + 1
            If nToalseatTemp > nTotalSeat Then Exit For
            '取座位类型
            If ArrayLength(tvTemp) = 0 Then
                szSeatTypeID = cszSeatTypeIsNormal
            Else
                szSeatTypeID = GetSeatTypeID(nTemp2, tvTemp)
                If szSeatTypeID = "" Then szSeatTypeID = cszSeatTypeIsNormal
            End If
            nSeatTypeSirel = Val(szSeatTypeID)
    
            tSeatInfoCountTemp(nSeatTypeSirel).seatCount = tSeatInfoCountTemp(nSeatTypeSirel).seatCount + 1
    
            '处理预留座位
            If nBeginReserveSeatNo > 0 Then
                If nTemp2 >= nBeginReserveSeatNo And nTemp2 <= nBeginReserveSeatNo + nReserveSeatCount - 1 Then
                    nTemp1 = ST_SeatReserved
                    nSeatTypeSirel = Val(szSeatTypeID)
                    tSeatInfoCountTemp(nSeatTypeSirel).seatCount = tSeatInfoCountTemp(nSeatTypeSirel).seatCount - 1
                End If
            End If
            '处理计划预定
            If bBookYes = True And nProjectBookCountTemp < nProjectBookCount Then
                If nTemp1 = ST_SeatCanSell Then
            
                    If nBeginSeatNo > CInt(szProjectBusBookInfo(3)) Then
            
                        '计划预定号小于车次起始座号,那么当前座号就被预定
                        If nBeginBookSeatNo = 0 Then nBeginBookSeatNo = nTemp2
                        nTemp1 = ST_SeatProjectBooked
                        nProjectBookCountTemp = nProjectBookCountTemp + 1
                    Else
                        '计划预定起始座号大于等于车次起始座号,那么当前座号大于等于就被预定
                        If nTemp2 >= CInt(szProjectBusBookInfo(2)) Then
                            If nBeginBookSeatNo = 0 Then nBeginBookSeatNo = nTemp2
                            nTemp1 = ST_SeatProjectBooked
                            nProjectBookCountTemp = nProjectBookCountTemp + 1
                        End If
                    End If
            
                End If
            End If
            
            '处理环境预留
            If szEnvBookSeatNo <> "" And InStr(1, szEnvBookSeatNo, BuildSeatNo(nTemp2)) > 0 Then
                nTemp1 = ST_SeatBooked
                tSeatInfoCountTemp(nSeatTypeSirel).seatCount = tSeatInfoCountTemp(nSeatTypeSirel).seatCount - 1
            End If
    
            szSql = "INSERT Env_bus_seat_lst( " _
                & "bus_date," _
                & "bus_id," _
                & "seat_no," _
                & "status," _
                & "ticket_no," _
                & "seat_type_id) "
            szSql = szSql & " VALUES('" _
                & ToDBDate(pdtRunDate) & "','" _
                & pszBusID & "','" _
                & BuildSeatNo(nTemp2) & "'," _
                & nTemp1 & ",'','" _
                & szSeatTypeID & "')"
            oDb.Execute szSql
        Next i
    
        '修正计划预定,使 计划预定的起始座号 和 计划预定数 同实际计划预定数 相符合
        If bBookYes = True Then
            If nProjectBookCountTemp < nProjectBookCount Or nBeginBookSeatNo <> CInt(szProjectBusBookInfo(2)) Then
            
                szSql = "Update project_book_info set seat_start_no=" & nBeginBookSeatNo & " ,  " _
                    & " seat_book_Count=" & nProjectBookCountTemp & "  " _
                    & "   WHERE " _
                    & " start_bus_date<='" & ToDBDate(pdtRunDate) & " '" _
                    & " and end_bus_date>='" & ToDBDate(pdtRunDate) & " '" _
                    & " and bus_id='" & pszBusID & "'"
                oDb.Execute szSql
            End If
        End If
        '更新环境车次表座位总数
        UpdateEnviromentSeatCountEx pszBusID, pdtRunDate, oDb
    
    End If
    Debug.Print "环境座位表:" & Time
    '插入环境车次站点表信息（enviroment_bus_station_info）和环境票价表信息（Env_bus_price_lst）
    Dim nStationCount As Integer
    ' Dim szTicketPriceTable As String
    Dim j As Integer
    Dim k As Integer
    Dim rsTemp1 As Recordset
    Dim rsBusPrice As Recordset
    
    '各票种总票价
    Dim sgFullPrice As Single
    Dim sgHalfPrice As Single
    Dim sgPreferentialPrice1 As Single
    Dim sgPreferentialPrice2 As Single
    Dim sgPreferentialPrice3 As Single
    
    Dim sgTotalPrice As Single
    Dim nCount As Integer
    
    szSql = "SELECT * FROM Bus_station_info WHERE " _
        & "bus_id='" & pszBusID & "' ORDER BY station_serial_no"
    Set rsTemp = oDb.Execute(szSql)
    
    If rsTemp.RecordCount = 0 Then SetSysPar oDb, True: ShowError ERR_BusNoStation
    nStationCount = rsTemp.RecordCount
            
    szSql = "DELETE FROM Env_bus_station_lst WHERE bus_id ='" & pszBusID & "' AND bus_date='" & ToDBDate(pdtRunDate) & "'"
    oDb.Execute szSql
    '工作表不需要删除
    '        zSql = "DELETE FROM Work_env_bus_station_lst WHERE bus_id ='" & pszBusID & "' AND bus_date='" & ToDBDate(pdtRunDate) & "'"
    '        oDb.Execute szSql
    szSql = "DELETE FROM Env_bus_price_lst WHERE bus_id ='" & pszBusID & "' AND bus_date='" & ToDBDate(pdtRunDate) & "'"
    oDb.Execute szSql
    '        szSql = "DELETE FROM Work_env_bus_price_lst WHERE bus_id ='" & pszBusID & "' AND bus_date='" & ToDBDate(pdtRunDate) & "'"
    '        oDb.Execute szSql
    '工作表不需要删除
    
    Dim bflgWatchtvTemp   As Boolean
    Dim lgRow As Long
    If btvTempNoSeatTypeIsNormal = True Then
        bflgWatchtvTemp = True
    End If
    If bflgWatchtvTemp = True Then
        btvTempNoSeatTypeIsNormal = True
    End If
    
    If ArrayLength(tvTemp) = 0 Then
        '如果车辆座位类型无设置
        szSql = " INSERT Env_bus_station_lst " _
            & " ( " _
            & " bus_date , bus_id, sell_station_id, station_id, seat_type_id, station_serial_no, max_sale_quantity, sale_ticket_quantity, book_count " _
            & " , stop_sale_time , mileage , full_price , half_price , preferential_ticket1 , preferential_ticket2 , preferential_ticket3 " _
            & " ) " _
            & " SELECT " _
            & " '" & ToDBDate(pdtRunDate) & "' , g.bus_id ,s.sell_station_id, g.station_id  , '" & cszSeatTypeIsNormal & "' ,station_serial_no , max_sale_quantity , max_sale_quantity , 0 " _
            & " , stop_sale_time , end_station_mileage , ISNULL(f.ticket_price_total,0) , ISNULL(h.ticket_price_total,0) , ISNULL(p1.ticket_price_total,0) , ISNULL(p2.ticket_price_total,0) , ISNULL(p3.ticket_price_total,0) " _
            & " FROM Bus_station_info g ,  Route_section_lst s , " _
            & " ( SELECT bus_id , station_id , sell_station_id , base_carriage + price_item_1 + price_item_2 + price_item_3 + price_item_4 + price_item_5 + price_item_6 " _
            & "     + price_item_7 + price_item_8 + price_item_9 + price_item_10 + price_item_11 + price_item_12 + price_item_13 " _
            & "     + price_item_14 + price_item_15 ticket_price_total FROM bus_price_lst " _
            & " WHERE ticket_type = " & TP_FullPrice & " AND price_table_id = '" & Trim(szTicketPriceTable) & "' AND vehicle_type_code = '" & szVehicleTypeCode & "' AND bus_id = '" & pszBusID & "' AND seat_type_id  =  '" & cszSeatTypeIsNormal & "' " _
            & "     ) f , " _
            & " ( SELECT bus_id , station_id , sell_station_id, base_carriage + price_item_1 + price_item_2 + price_item_3 + price_item_4 + price_item_5 + price_item_6 " _
            & " + price_item_7 + price_item_8 + price_item_9 + price_item_10 + price_item_11 + price_item_12 + price_item_13 " _
            & " + price_item_14 + price_item_15 ticket_price_total FROM bus_price_lst " _
            & " WHERE ticket_type = " & TP_HalfPrice & "  AND price_table_id = '" & Trim(szTicketPriceTable) & "' AND vehicle_type_code = '" & szVehicleTypeCode & "' AND bus_id = '" & pszBusID & "' AND seat_type_id  =  '" & cszSeatTypeIsNormal & "' " _
            & " ) h , " _
            & " ( SELECT bus_id , station_id , sell_station_id, base_carriage + price_item_1 + price_item_2 + price_item_3 + price_item_4 + price_item_5 + price_item_6 " _
            & " + price_item_7 + price_item_8 + price_item_9 + price_item_10 + price_item_11 + price_item_12 + price_item_13 " _
            & " + price_item_14 + price_item_15 ticket_price_total FROM bus_price_lst " _
            & " WHERE ticket_type = " & TP_PreferentialTicket1 & "  AND price_table_id = '" & Trim(szTicketPriceTable) & "' AND vehicle_type_code = '" & szVehicleTypeCode & "' AND bus_id = '" & pszBusID & "' AND seat_type_id  =  '" & cszSeatTypeIsNormal & "' " _
            & " ) p1 , " _
            & " ( SELECT bus_id , station_id , sell_station_id, base_carriage + price_item_1 + price_item_2 + price_item_3 + price_item_4 + price_item_5 + price_item_6 "
        
        szSql = szSql _
            & " + price_item_7 + price_item_8 + price_item_9 + price_item_10 + price_item_11 + price_item_12 + price_item_13 " _
            & " + price_item_14 + price_item_15 ticket_price_total FROM bus_price_lst " _
            & " WHERE ticket_type = " & TP_PreferentialTicket2 & "  AND price_table_id = '" & Trim(szTicketPriceTable) & "' AND vehicle_type_code = '" & szVehicleTypeCode & "' AND bus_id = '" & pszBusID & "' AND seat_type_id  =  '" & cszSeatTypeIsNormal & "' " _
            & " ) p2 , " _
            & " ( SELECT bus_id , station_id , sell_station_id, base_carriage + price_item_1 + price_item_2 + price_item_3 + price_item_4 + price_item_5 + price_item_6 " _
            & " + price_item_7 + price_item_8 + price_item_9 + price_item_10 + price_item_11 + price_item_12 + price_item_13 " _
            & " + price_item_14 + price_item_15 ticket_price_total FROM bus_price_lst " _
            & " WHERE ticket_type = " & TP_PreferentialTicket3 & "  AND price_table_id = '" & Trim(szTicketPriceTable) & "' AND vehicle_type_code = '" & szVehicleTypeCode & "' AND bus_id = '" & pszBusID & "' AND seat_type_id  =  '" & cszSeatTypeIsNormal & "' " _
            & " ) p3 " _
            & " WHERE  g.bus_id = '" & pszBusID & "' " _
            & " AND g.bus_id *= f.bus_id AND g.bus_id *= h.bus_id AND g.bus_id *= p1.bus_id AND g.bus_id *= p2.bus_id AND g.bus_id *=p3.bus_id " _
            & " AND g.station_id *= f.station_id AND g.station_id *= h.station_id AND g.station_id *= p1.station_id AND g.station_id *= p2.station_id AND g.station_id *=p3.station_id "
            
        szSql = szSql & " AND g.station_id = s.end_station_id AND s.route_id = '" & szRouteID & "' AND s.sell_station_id *= f.sell_station_id AND s.sell_station_id *= h.sell_station_id " & " AND s.sell_station_id *= p1.sell_station_id AND s.sell_station_id *= p2.sell_station_id AND s.sell_station_id *= p3.sell_station_id"
            
            
        oDb.Execute szSql, lgRow
    
        If lgRow <= 0 Then SetSysPar oDb, True: ShowError Err_NOPriceOrStionID
        
        '环境车次票价
        szSql = "INSERT Env_bus_price_lst " _
            & " (  " _
            & " bus_date , bus_id,sell_station_id, station_id, ticket_type, seat_type_id, station_serial_no, ticket_price_total, base_carriage " _
            & " , price_item_1 , price_item_2 , price_item_3 , price_item_4 , price_item_5 , price_item_6 , price_item_7 , price_item_8 " _
            & " , price_item_9 , price_item_10 , price_item_11 , price_item_12 , price_item_13 , price_item_14 , price_item_15 " _
            & " ) " _
            & " SELECT '" & ToDBDate(pdtRunDate) & "' , " _
            & " bus_id ,sell_station_id, station_id, ticket_type, seat_type_id, station_serial_no " _
            & " , base_carriage + price_item_1 + price_item_2 + price_item_3 + price_item_4 + price_item_5 + price_item_6 " _
            & " + price_item_7 + price_item_8 + price_item_9 + price_item_10 + price_item_11 + price_item_12 + price_item_13" _
            & " + price_item_14 + price_item_15 ticket_price_total " _
            & " , base_carriage , price_item_1 , price_item_2 , price_item_3 , price_item_4 , price_item_5 , price_item_6 " _
            & " , price_item_7 , price_item_8 , price_item_9 , price_item_10 , price_item_11 , price_item_12 , price_item_13 " _
            & " , price_item_14 , price_item_15 " _
            & " FROM bus_price_lst " _
            & " WHERE price_table_id = '" & Trim(szTicketPriceTable) & "' AND vehicle_type_code = '" & szVehicleTypeCode & "' AND bus_id = '" & pszBusID & "' AND seat_type_id  =  '" & cszSeatTypeIsNormal & "' "
        
        oDb.Execute szSql
        '环境车次
    Else
    '如果车辆座位类型有设置
        Dim szSeatTypeTemp As String
        Dim nCountSeatType As Integer
        szSql = "SELECT DISTINCT seat_type_id  FROM vehicle_seat_type_info WHERE vehicle_id='" & szVehicleID & "' ORDER BY seat_type_id"
        Set rsTemp1 = oDb.Execute(szSql)
        nCountSeatType = rsTemp1.RecordCount
        '默认座位 btvTempNoSeatTypeIsNormal = True
        'vehicle_seat_type_info 有些座位没设置，将作为默认座位处理
        If btvTempNoSeatTypeIsNormal = True Then
            nCountSeatType = nCountSeatType + 1
        End If
    
        For j = 1 To nCountSeatType
            szSeatTypeTemp = FormatDbValue(rsTemp1!seat_type_id)
            If btvTempNoSeatTypeIsNormal = True Then
                szSeatTypeTemp = cszSeatTypeIsNormal
            End If
            szSql = " INSERT Env_bus_station_lst " _
                & " ( " _
                & " bus_date , bus_id, station_id, sell_station_id,  seat_type_id, station_serial_no, max_sale_quantity, sale_ticket_quantity, book_count " _
                & " , stop_sale_time , mileage , full_price , half_price , preferential_ticket1 , preferential_ticket2 , preferential_ticket3 " _
                & " ) " _
                & " SELECT " _
                & " '" & ToDBDate(pdtRunDate) & "' , g.bus_id , g.station_id , s.sell_station_id  , '" & szSeatTypeTemp & "' ,station_serial_no , max_sale_quantity , max_sale_quantity , 0  " _
                & " , stop_sale_time , end_station_mileage , ISNULL(f.ticket_price_total,0) , ISNULL(h.ticket_price_total,0) , ISNULL(p1.ticket_price_total,0) , ISNULL(p2.ticket_price_total,0) , ISNULL(p3.ticket_price_total,0) " _
                & " FROM Bus_station_info g ,  Route_section_lst s ,  " _
                & " ( SELECT bus_id , station_id   , sell_station_id, base_carriage + price_item_1 + price_item_2 + price_item_3 + price_item_4 + price_item_5 + price_item_6 " _
                & "     + price_item_7 + price_item_8 + price_item_9 + price_item_10 + price_item_11 + price_item_12 + price_item_13 " _
                & "     + price_item_14 + price_item_15 ticket_price_total FROM bus_price_lst " _
                & " WHERE ticket_type = " & TP_FullPrice & " AND price_table_id = '" & Trim(szTicketPriceTable) & "' AND vehicle_type_code = '" & szVehicleTypeCode & "' AND bus_id = '" & pszBusID & "' AND seat_type_id  =  '" & szSeatTypeTemp & "' " _
                & "     ) f , " _
                & " ( SELECT bus_id , station_id  , sell_station_id, base_carriage + price_item_1 + price_item_2 + price_item_3 + price_item_4 + price_item_5 + price_item_6 " _
                & " + price_item_7 + price_item_8 + price_item_9 + price_item_10 + price_item_11 + price_item_12 + price_item_13 " _
                & " + price_item_14 + price_item_15 ticket_price_total FROM bus_price_lst " _
                & " WHERE ticket_type = " & TP_HalfPrice & "  AND price_table_id = '" & Trim(szTicketPriceTable) & "' AND vehicle_type_code = '" & szVehicleTypeCode & "' AND bus_id = '" & pszBusID & "' AND seat_type_id  =  '" & szSeatTypeTemp & "' " _
                & " ) h , " _
                & " ( SELECT bus_id , station_id  , sell_station_id, base_carriage + price_item_1 + price_item_2 + price_item_3 + price_item_4 + price_item_5 + price_item_6 " _
                & " + price_item_7 + price_item_8 + price_item_9 + price_item_10 + price_item_11 + price_item_12 + price_item_13 " _
                & " + price_item_14 + price_item_15 ticket_price_total FROM bus_price_lst " _
                & " WHERE ticket_type = " & TP_PreferentialTicket1 & "  AND price_table_id = '" & Trim(szTicketPriceTable) & "' AND vehicle_type_code = '" & szVehicleTypeCode & "' AND bus_id = '" & pszBusID & "' AND seat_type_id  =  '" & szSeatTypeTemp & "' " _
                & " ) p1 , " _
                & " ( SELECT bus_id , station_id  , sell_station_id, base_carriage + price_item_1 + price_item_2 + price_item_3 + price_item_4 + price_item_5 + price_item_6 "
    
            szSql = szSql _
                & " + price_item_7 + price_item_8 + price_item_9 + price_item_10 + price_item_11 + price_item_12 + price_item_13 " _
                & " + price_item_14 + price_item_15 ticket_price_total FROM bus_price_lst " _
                & " WHERE ticket_type = " & TP_PreferentialTicket2 & "  AND price_table_id = '" & Trim(szTicketPriceTable) & "' AND vehicle_type_code = '" & szVehicleTypeCode & "' AND bus_id = '" & pszBusID & "' AND seat_type_id  =  '" & szSeatTypeTemp & "' " _
                & " ) p2 , " _
                & " ( SELECT bus_id , station_id  , sell_station_id, base_carriage + price_item_1 + price_item_2 + price_item_3 + price_item_4 + price_item_5 + price_item_6 " _
                & " + price_item_7 + price_item_8 + price_item_9 + price_item_10 + price_item_11 + price_item_12 + price_item_13 " _
                & " + price_item_14 + price_item_15 ticket_price_total FROM bus_price_lst " _
                & " WHERE ticket_type = " & TP_PreferentialTicket3 & "  AND price_table_id = '" & Trim(szTicketPriceTable) & "' AND vehicle_type_code = '" & szVehicleTypeCode & "' AND bus_id = '" & pszBusID & "' AND seat_type_id  =  '" & szSeatTypeTemp & "' " _
                & " ) p3 " _
                & " WHERE  g.bus_id = '" & pszBusID & "' " _
                & " AND g.bus_id *= f.bus_id AND g.bus_id *= h.bus_id AND g.bus_id *= p1.bus_id AND g.bus_id *= p2.bus_id AND g.bus_id *=p3.bus_id " _
                & " AND g.station_id *= f.station_id AND g.station_id *= h.station_id AND g.station_id *= p1.station_id AND g.station_id *= p2.station_id AND g.station_id *=p3.station_id "
            
            szSql = szSql _
                & " AND g.station_id = s.end_station_id AND s.route_id = '" & szRouteID & "' AND s.sell_station_id *= f.sell_station_id AND s.sell_station_id *= h.sell_station_id " _
                & " AND s.sell_station_id *= p1.sell_station_id AND s.sell_station_id *= p2.sell_station_id AND s.sell_station_id *= p3.sell_station_id"
            
            oDb.Execute szSql, lgRow
            If lgRow <= 0 Then SetSysPar oDb, True: ShowError Err_NOPriceOrStionID
            '环境车次票价
            szSql = "INSERT Env_bus_price_lst " _
                & " (  " _
                & " bus_date , bus_id,sell_station_id, station_id, ticket_type, seat_type_id, station_serial_no, ticket_price_total, base_carriage " _
                & " , price_item_1 , price_item_2 , price_item_3 , price_item_4 , price_item_5 , price_item_6 , price_item_7 , price_item_8 " _
                & " , price_item_9 , price_item_10 , price_item_11 , price_item_12 , price_item_13 , price_item_14 , price_item_15 " _
                & " ) " _
                & " SELECT '" & ToDBDate(pdtRunDate) & "' , " _
                & " bus_id , sell_station_id,station_id, ticket_type, seat_type_id, station_serial_no " _
                & " , base_carriage + price_item_1 + price_item_2 + price_item_3 + price_item_4 + price_item_5 + price_item_6 " _
                & " + price_item_7 + price_item_8 + price_item_9 + price_item_10 + price_item_11 + price_item_12 + price_item_13" _
                & " + price_item_14 + price_item_15 ticket_price_total " _
                & " , base_carriage , price_item_1 , price_item_2 , price_item_3 , price_item_4 , price_item_5 , price_item_6 " _
                & " , price_item_7 , price_item_8 , price_item_9 , price_item_10 , price_item_11 , price_item_12 , price_item_13 " _
                & " , price_item_14 , price_item_15 " _
                & " FROM bus_price_lst " _
                & " WHERE price_table_id = '" & Trim(szTicketPriceTable) & "' AND vehicle_type_code = '" & szVehicleTypeCode & "' AND bus_id = '" & pszBusID & "' AND seat_type_id  =  '" & szSeatTypeTemp & "' "
            oDb.Execute szSql
            '                    '环境车次
            If btvTempNoSeatTypeIsNormal = True Then
                btvTempNoSeatTypeIsNormal = False
            Else
                rsTemp1.MoveNext
            End If
        Next j
    End If
    
    '更新环境预定信息   2008-03-10 by zyw
    For j = 1 To ArrayLength(szaStationBook)
        szSql = " UPDATE env_bus_station_lst SET book_count = " & szaStationBook(j, 4) _
              & " WHERE bus_id = '" & pszBusID & "'" _
              & " AND bus_date = '" & ToDBDate(pdtRunDate) & "'" _
              & " AND sell_station_id = '" & szaStationBook(j, 1) & "'" _
              & " AND station_id = '" & szaStationBook(j, 2) & "'" _
              & " AND seat_type_id = '" & szaStationBook(j, 3) & "'"
        oDb.Execute szSql
    Next j
    
    
'    Debug.Print "环境车次站点表和环境票价表2:" & Time
    '插入环境车次配载表
    szSql = " INSERT Env_bus_allot_lst ( bus_date , bus_id , sell_station_id , check_gate_id , bus_start_time , status , can_sale_quantity , sell_quantity ) " _
        & " SELECT '" & ToDBDate(pdtRunDate) & "' , bus_id , sell_station_id , check_gate_id ,'" & ToDBDate(pdtRunDate) & "' + ' ' + convert(char(8),bus_start_time ,108) , " & nBusStatus & " , can_sale_quantity , 0 " _
        & " FROM bus_allot_info " _
        & " WHERE  bus_id = '" & pszBusID & "' "
    oDb.Execute szSql
    
    '插入环境车次售票站表
    szSql = " INSERT env_bus_sell_station_lst ( bus_date , bus_id , sell_station_id , can_sale_quantity , sell_quantity ) " _
        & " SELECT '" & ToDBDate(pdtRunDate) & "' , bus_id , sell_station_id  , can_sale_quantity , 0 " _
        & " FROM bus_sell_station_lst " _
        & " WHERE  bus_id = '" & pszBusID & "' "
    oDb.Execute szSql
    
    '加班车处理
    If bEndStationCanSale = True Then
        Dim nString As Integer
        nString = InStr(1, szAddBusFix, "%")
        If nString <> 0 Then
            szAddBusFix = Left(szAddBusFix, nString - 1)
        End If
        
        nString = Len(szAddBusFix)
    
        If szAddBusFix <> "" And nString <> 0 Then
    
            If Trim(Left(Trim(pszBusID), nString)) = Trim(szAddBusFix) Then
                '0----站点不可售
                szSql = "Update Env_bus_station_lst set max_sale_quantity=0 WHERE station_serial_no< ( SELECT Max(station_serial_no)  as station_serial_no FROM Env_bus_station_lst  " _
                & "  WHERE bus_id='" & pszBusID & "' AND bus_date='" & ToDBDate(pdtRunDate) & "'and max_sale_quantity=-1)  " _
                & "  and bus_id='" & pszBusID & "' AND bus_date='" & ToDBDate(pdtRunDate) & "'"
                oDb.Execute szSql
            End If
        End If
    End If
    '已售票车次处理
    If bFlgSaleBusCanMake = True And bflgSaleBus = True Then
        MakeSaleBusEn pszBusID, pdtRunDate, szaStationSale, tSeatInfoCountTemp, oDb
        SetSysPar oDb, True
    End If
    
    '结束事务
    'SetSysPar odb, True
    
    oDb.CommitTrans
    
    Set rsTemp = Nothing
    Set rsTemp1 = Nothing
    Set rsBusPrice = Nothing
    
    Exit Sub
here:
    
    oDb.RollbackTrans
    SetSysPar oDb, True
    Set oDb = Nothing
    Set rsTemp = Nothing
    Set rsTemp1 = Nothing
    Set rsBusPrice = Nothing
    err.Raise err.Number
End Sub
'车次车辆停班,复班
Public Function StopOrResumBusVehile(szBusID As String, szVehicleID As String, bflg As Boolean)
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim nTmpCount As Integer
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    On Error GoTo ErrorHandle
    
    oDb.BeginTrans
        If bflg Then
            '把该车次车辆设定为复班
            szSql = "Update Env_bus_info set status=" & ST_BusNormal _
                & " WHERE bus_id='" & szBusID & "' and vehicle_id='" & szVehicleID & "' and bus_date>='" & ToDBDate(Now) & "' "
            oDb.Execute szSql
            '与计划车次车辆表进行关联,把表中设定的日期中的数据进行停班
            szSql = "Update Env_bus_info set status=" & ST_BusNormal _
                & " from Env_bus_info b , bus_vehicle_code v  " _
                & " Where b.bus_id = v.bus_id And b.vehicle_id = v.vehicle_id And b.bus_date >= stop_start_date " _
                & " and b.bus_date <=dateadd(day,1 ,stop_end_date) " _
                & " and b.bus_id='" & szBusID & "' and b.vehicle_id='" & szVehicleID & "'"
            oDb.Execute szSql
        Else
            '与计划车次车辆表进行关联,把表中设定的日期中的数据进行停班
            szSql = "Update Env_bus_info set status=" & ST_BusStopped _
                & " from Env_bus_info b , bus_vehicle_code v  " _
                & " Where b.bus_id = v.bus_id And b.vehicle_id = v.vehicle_id And b.bus_date >= stop_start_date " _
                & " and b.bus_date <=dateadd(day,1 ,stop_end_date) " _
                & " and b.bus_id='" & szBusID & "' and b.vehicle_id='" & szVehicleID & "'"
            oDb.Execute szSql
        End If
    oDb.CommitTrans
'    oDb.Execute szSql
    Exit Function
ErrorHandle:
    oDb.RollbackTrans
    err.Raise err.Number, err.Source, err.Description
End Function

'''''''''''''''
'功能：已售车次重新生成环境
'  2001-08-09
'可能丢失站点限售张数
'szNewSeatTypeInfo---  当前车次的座位分配情况
'szStationInfo-----    原车次环境站点售票情况
'szatSeatTypeCount---- 原车次已售座位情况
'''''''''''''''
Private Function MakeSaleBusEn(szBusID As String, dtDate As Date, szStationInfo() As String, szNewSeatTypeInfo() As TSeatInfoCount, odbTemp As RTConnection)
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim i As Integer
    Dim rsTemp As New Recordset
    Dim nCount As Integer
    Dim nSeatType As Integer
    Dim rs As Recordset
    Dim szatSeatTypeCOUNT(1 To 5) As TSeatInfoCount
    Dim j As Integer
    Dim lRow As Long
    If Not odbTemp Is Nothing Then
        Set oDb = odbTemp
    Else
        oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    End If
    
    '环境座位数据恢复
    szSql = "SELECT * FROM Ticket_sell_lst WHERE bus_id ='" & szBusID & "' and bus_date='" & ToDBDate(dtDate) & "'" _
        & " AND (status & 28) =0 "
        
    Set rsTemp = oDb.Execute(szSql)
    nCount = rsTemp.RecordCount
    If nCount > 0 Then
    
        For i = 1 To nCount
            szSql = "update Env_bus_seat_lst set status='" & ST_SeatSold & "' ,ticket_no='" & Trim(rsTemp!ticket_id) & "'" _
                & " WHERE bus_id='" & szBusID & "' and bus_date='" & ToDBDate(dtDate) & "' and seat_no ='" & Trim(rsTemp!seat_no) & "'"
            oDb.Execute szSql, lRow
        
            If lRow < 0 Then
                szSql = "update Env_bus_seat_lst set status='" & ST_SeatSold & "' ,ticket_no='" & Trim(rsTemp!ticket_id) & "'" _
                & " WHERE bus_id='" & szBusID & "' and bus_date='" & ToDBDate(dtDate) & "' and seat_no =( " _
                & "SELECT min(seat_no) FROM Env_bus_seat_lst  WHERE bus_id='" & szBusID & "' and bus_date='" & ToDBDate(dtDate) & "' " _
                & "and status='" & ST_SeatCanSell & "' and ticket_no<>'" & Trim(rsTemp!ticket_id) & "')"
                oDb.Execute szSql
            
            End If
            
            nSeatType = Val(rsTemp!seat_type_id)
            szatSeatTypeCOUNT(nSeatType).seatCount = szatSeatTypeCOUNT(nSeatType).seatCount - 1
            rsTemp.MoveNext
        Next
    End If
    
    '环境站点数据恢复
    nCount = ArrayLength(szStationInfo)
    If nCount <> 0 Then
        For i = 1 To nCount
            szSql = "update Env_bus_station_lst set  max_sale_quantity=" & CInt(szStationInfo(i, 2)) & " ," _
            & "  sale_ticket_quantity='" & szStationInfo(i, 3) & "', " _
            & "  stop_sale_time=" & Val(szStationInfo(i, 5)) _
            & "  WHERE bus_id='" & szBusID & "'" _
            & "  and bus_date='" & ToDBDate(dtDate) & "'" _
            & "  and station_id='" & szStationInfo(i, 1) & "'" _
            & "  and seat_type_id='" & szStationInfo(i, 4) & "'"
            oDb.Execute szSql
        Next
    End If
    
    szSql = "SELECT bus_type FROM env_bus_info WHERE bus_id ='" & szBusID & "' AND bus_date='" & ToDBDate(dtDate) & "'"
    Set rs = oDb.Execute(szSql)
    If FormatDbValue(rs!bus_type) <> TP_ScrollBus Then
        UpdateEnviromentSeatCountEx szBusID, dtDate, oDb
    End If
    
    '车次配载信息恢复
    szSql = "SELECT sell_station_id , COUNT(*) sell_quantity FROM Ticket_sell_lst WHERE bus_date = '" & ToDBDate(dtDate) & "' AND bus_id = '" & szBusID & "'" _
        & " AND (status & 28) =0 " _
        & " GROUP BY sell_station_id "
    Set rsTemp = oDb.Execute(szSql)
    nCount = rsTemp.RecordCount
    For i = 1 To nCount
        szSql = "UPDATE env_bus_allot_lst SET sell_quantity = " & FormatDbValue(rsTemp!sell_quantity) _
            & " WHERE sell_station_id = '" & FormatDbValue(rsTemp!sell_station_id) & "'" _
            & " AND bus_date = '" & ToDBDate(dtDate) & "'" & " AND bus_id ='" & szBusID & "'"
        rsTemp.MoveNext
        oDb.Execute szSql
    Next i
    
    
    '车次售票点信息恢复
    szSql = " SELECT u.sell_station_id , COUNT(*) sell_quantity " _
        & " FROM Ticket_sell_lst s , user_info u " _
        & " WHERE s.user_id = u.user_id AND bus_date = '" & ToDBDate(dtDate) & "' AND bus_id = '" & szBusID & "'" _
        & " AND (s.status & 28) =0 " _
        & " GROUP BY u.sell_station_id "
    Set rsTemp = oDb.Execute(szSql)
    nCount = rsTemp.RecordCount
    For i = 1 To nCount
        szSql = "UPDATE env_bus_sell_station_lst SET sell_quantity = " & FormatDbValue(rsTemp!sell_quantity) _
            & " WHERE sell_station_id = '" & FormatDbValue(rsTemp!sell_station_id) & "'" _
            & " AND bus_date = '" & ToDBDate(dtDate) & "'" & " AND bus_id ='" & szBusID & "'"
        rsTemp.MoveNext
        oDb.Execute szSql
    Next i
    
    
    Set rsTemp = Nothing
End Function
'''''''临时改变一下
Private Function SetSysPar(odbTemp As RTConnection, Optional bflg As Boolean = False)
Dim szSql As String
Dim oDb As New RTConnection
  If Not odbTemp Is Nothing Then
        Set oDb = odbTemp
   Else
       oDb.ConnectionString = GetConnectionStr(cszRunEnv)
   End If
 If bflg = False Then
     szSql = "update System_param_info  set parameter_value=parameter_value+32 WHERE parameter_name ='AllowMakeEnviromentSaleBus' and parameter_value=1"
 Else
     szSql = "update System_param_info  set parameter_value=1 WHERE parameter_name ='AllowMakeEnviromentSaleBus' and parameter_value<>0"
 End If
 oDb.Execute szSql
End Function

Public Function EnBusRollbackTrans() As Long
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim lRow As Long
    Dim nPreSaleDate As Integer
    Dim sysPar As New SystemParam
    Dim dtStartdate As Date
    Dim dtEnddate As Date
    
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    
    sysPar.Init m_oActiveUser
    
    nPreSaleDate = sysPar.PreSellDate
    dtStartdate = DateAdd("d", -nPreSaleDate, Now)
    
    dtEnddate = DateAdd("d", nPreSaleDate, Now)
    
    On Error GoTo here
    
    szSql = "UPDATE Env_bus_info SET status=status - " & CnSlitpStatus & " WHERE " _
           & " status>=" & CnSlitpStatus & " and ( bus_date  BETWEEN  '" & ToDBDate(dtStartdate) & "'and '" & ToDBDate(dtEnddate) & "') "
    
    oDb.Execute szSql, lRow
    
    szSql = "update System_param_info  set parameter_value=1 WHERE parameter_name ='AllowMakeEnviromentSaleBus' and parameter_value>1"
    oDb.Execute szSql
    
    
    EnBusRollbackTrans = lRow
    Set sysPar = Nothing
    Set oDb = Nothing
    Exit Function
here:
    Set sysPar = Nothing
    Set oDb = Nothing
    err.Raise err.Number
End Function

'用来删除新增车次失败的垃圾车次
Public Function DeletAllBusInfo(szBusID As String, busdate As Date, Optional szString As String = "")
    Dim szSql As String
    Dim oDb As New RTConnection
    AssertHaveRight m_oActiveUser, RIGHT_MakeRunEvironment
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    szSql = "delete Env_bus_info WHERE bus_id='" & szBusID & "' and bus_date='" & ToDBDate(busdate) & "'"
    oDb.Execute szSql
    szSql = "delete bus_info WHERE bus_id='" & szBusID & "'"
    oDb.Execute szSql
    Set oDb = Nothing
    WriteOperateLog m_oActiveUser, RIGHT_MakeRunEvironment, "环境新增车次[" & szBusID & "]失败." & szString
End Function


Public Function IdentifyNotBusId(szBusID As String, busdate As Date, szproject As String) As Boolean
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim rsTemp As New Recordset
    On Error GoTo here
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    szSql = "SELECT * FROM  Work_env_bus_info WHERE bus_id='" & szBusID & "' and bus_date='" & ToDBDate(busdate) & "'"
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount <> 0 Then
        IdentifyNotBusId = False
        Set oDb = Nothing
        Set rsTemp = Nothing
        Exit Function
    End If
    szSql = "SELECT * FROM bus_info WHERE bus_id='" & szBusID & "'"
    oDb.Execute szSql
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount <> 0 Then
        IdentifyNotBusId = False
        Set oDb = Nothing
        Set rsTemp = Nothing
        
        Exit Function
    End If
    IdentifyNotBusId = True
    Set oDb = Nothing
    Set rsTemp = Nothing
    Exit Function
here:
    err.Raise err.Number
End Function

'计划复班影响环境的长复班()
Public Sub AllResumeBus(szBusID As String, StartDate As Date, EndDate As Date, Optional AllRefundment As Boolean = False)
    Dim szSql As String
    Dim szTemp As String
    Dim oDb As New RTConnection
    Dim lRow As Long
    Dim szlog As String
    Dim bflg As Boolean

      AssertHaveRight m_oActiveUser, RIGHT_REBusResumeStopBus
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    If AllRefundment Then
       szTemp = cbAllRefundment
    Else
        szTemp = cbNotAllRefundment
    End If
    If StartDate = CDate(cszEmptyDateStr) And StartDate = EndDate Then
       bflg = True
        '复班环境所有车次
        szSql = "UPDATE Env_bus_info SET status=" & Str(ST_BusNormal) & ",is_all_refundment=" & Str(szTemp) & " WHERE bus_id='" & szBusID & "' AND " _
        & " bus_date>='" & ToDBDate(Now) & "'"
       
        szlog = "发车日期大于等于" & ToDBDate(Now) & "的环境车次[" & szBusID & "]全被复班"
    Else
      '5 修改车次的状态和是否全额退票
      szSql = "UPDATE Env_bus_info SET status=" & Str(ST_BusNormal) & ",is_all_refundment=" & Str(szTemp) & " WHERE bus_id='" & szBusID & "' AND " _
      & " bus_date>='" & ToDBDate(StartDate) & "' AND bus_date<='" & ToDBDate(EndDate) & "'"
      szlog = "车次" & szBusID & "在" & ToDBDate(StartDate) & "--" & ToDBDate(EndDate) & "复班"
    End If
    oDb.Execute szSql, lRow
    
      szSql = "UPDATE work_env_bus_allot_lst SET status=" & Str(ST_BusNormal) & " WHERE bus_id='" & szBusID & "' AND " _
      & " bus_date>='" & ToDBDate(Now) & "' "
      oDb.Execute szSql, lRow
      
      szSql = "UPDATE Env_bus_allot_lst SET status=" & Str(ST_BusNormal) & " WHERE bus_id='" & szBusID & "' AND " _
      & " bus_date>='" & ToDBDate(Now) & "' "
      oDb.Execute szSql, lRow
      
    #If SUPPORT_MSG = 1 Then
    Dim oParam As New SystemParam
    Dim aszValues(1 To 2) As String
    aszValues(1) = szBusID
    If bflg = True Then
      aszValues(2) = ToDBDate(Now)
    Else
     aszValues(2) = ToDBDate(StartDate)
    End If
      m_oSndMsg.SendMsg oParam.NowDateTime, cszResumeBus, aszValues
    
    #End If
    
    If lRow <> 0 Then
        WriteOperateLog m_oActiveUser, RIGHT_REBusResumeStopBus, szlog
    End If
    
End Sub
'计划长停班影响环境的长停班
Public Function AllStopBus(szBusID As String, StartDate As Date, EndDate As Date, Optional AllRefundment As Boolean = True) As String
    Dim szSql As String
    Dim szTemp As String
    Dim strlog As String
    Dim oDb As New RTConnection
    Dim szlog As String
    Dim lRow As Long
    Dim bflg As Boolean
    Dim szsqltemp As String
    Dim szSaleCount As String
    Dim rsTem As New Recordset
    Dim i As Integer
    AssertHaveRight m_oActiveUser, RIGHT_REBusStopBus
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    If AllRefundment Then
        szTemp = cbAllRefundment
    Else
        szTemp = cbNotAllRefundment
    End If
    
    If StartDate = CDate(cszEmptyDateStr) And StartDate = EndDate Then
        bflg = True
        szSql = "UPDATE Env_bus_info SET status=" & Str(ST_BusStopped) & ",is_all_refundment=" & Str(szTemp) & " WHERE bus_id='" & szBusID & "' AND " _
            & " bus_date>='" & ToDBDate(Now) & "'"
        oDb.Execute szSql
        szSql = " UPDATE Env_bus_allot_lst SET status=" & ST_BusStopped & _
                " WHERE bus_id=" & TransFieldValueToString(szBusID) & " AND bus_date>='" & ToDBDate(Now) & "'"
        oDb.Execute szSql
        szlog = "发车日期大于等于" & ToDBDate(Now) & "的环境车次[" & szBusID & "]全被停班"
        
        
'        szSql = "UPDATE Env_bus_info SET status=" & Str(ST_BusStopped) & ",is_all_refundment=" & Str(szTemp) & " WHERE bus_id='" & szBusID & "' AND " _
'        & " bus_date>='" & ToDBDate(Now) & "'"
        szlog = "发车日期大于等于" & ToDBDate(Now) & "的环境车次[" & szBusID & "]全被停班"
        szsqltemp = "SELECT COUNT(*)  AS seatcount , bus_date FROM  Env_bus_seat_lst WHERE (status=" & ST_SeatSold & " or status=" & ST_SeatReplace & "  or status=" & ST_SeatSlitp & " )" _
        & " AND bus_id='" & szBusID & "' AND" _
        & " bus_date>='" & ToDBDate(Now) & " ' GROUP by bus_date"
    
    Else
        '复班环境所有车次
        '5修改被车次的状态和是否全额退票
        
        szSql = "UPDATE Env_bus_info SET status=" & Str(ST_BusStopped) & ",is_all_refundment=" & Str(szTemp) & " WHERE bus_id='" & szBusID & "' AND " _
                & " bus_date>='" & ToDBDate(StartDate) & "' AND bus_date<='" & ToDBDate(EndDate) & "'"
        oDb.Execute szSql, lRow
        szSql = " UPDATE Env_bus_allot_lst SET status=" & ST_BusStopped & _
                " WHERE bus_id=" & TransFieldValueToString(szBusID) & " AND bus_date>='" & ToDBDate(StartDate) & "' AND bus_date<='" & ToDBDate(EndDate) & "'"
        oDb.Execute szSql
        szlog = "车次[" & szBusID & "] 车次状态" & strlog & ":在" & ToDBDate(StartDate) & "--" & ToDBDate(EndDate) & "停班"

    
'
'        szSql = "UPDATE Env_bus_info SET status=" & Str(ST_BusStopped) & ",is_all_refundment=" & Str(szTemp) & " WHERE bus_id='" & szBusID & "' AND " _
'        & " bus_date>='" & ToDBDate(StartDate) & "' AND bus_date<='" & ToDBDate(EndDate) & "'"
'        szlog = "车次" & szBusID & "车次状态：" & strlog & ",在" & ToDBDate(StartDate) & "--" & ToDBDate(EndDate) & "停班"
        
        szsqltemp = "SELECT COUNT(*)  AS seatcount , bus_date FROM  Env_bus_seat_lst WHERE (status=" & ST_SeatSold & " or status=" & ST_SeatReplace & "  or status=" & ST_SeatSlitp & " )" _
        & " AND bus_id='" & szBusID & "' AND" _
        & "  bus_date>='" & ToDBDate(StartDate) & "' AND bus_date<='" & ToDBDate(EndDate) & "'" _
        & "  GROUP by bus_date"
    End If
'    oDb.Execute szSql, lRow
    Set rsTem = oDb.Execute(szsqltemp)
    If rsTem.RecordCount <> 0 Then
        For i = 1 To rsTem.RecordCount
            If rsTem!seatCount <> 0 Then
            szSaleCount = szSaleCount & "车次[" & szBusID & "]在" & Format(rsTem!bus_date, "YYYY-MM-DD") & "日期已售" & rsTem!seatCount & "张" & Chr(10)
            End If
            rsTem.MoveNext
        Next
    End If
    AllStopBus = szSaleCount
    #If SUPPORT_MSG = 1 Then
        Dim oParam As New SystemParam
        Dim aszValues(1 To 2) As String
        aszValues(1) = szBusID
        If bflg = True Then
            aszValues(2) = ToDBDate(Now)
        Else
            aszValues(2) = ToDBDate(StartDate)
        End If
        m_oSndMsg.SendMsg oParam.NowDateTime, cszStopBus, aszValues
    #End If
    ' strlog = "车次状态：" & strlog & ",在"
    If lRow <> 0 Then
        WriteOperateLog m_oActiveUser, RIGHT_REBusStopBus, szlog
    End If
End Function
Private Sub Class_Initialize()
    #If SUPPORT_MSG = 1 Then
        m_oSndMsg.Unit = GetUnitID()
        m_oSndMsg.MsgSource = cszSNRunEnv
    #End If
End Sub
'得到车次站点已售票信息--
'szBusId---车次代码 dtRunBusDate－－－车次发车日期
'Public Function GetBusStatinSellInfo(szBusID As String, dtRunBusdate As Date) As String()
'
'   Dim oDb As New RTConnection
'   Dim szSql As String
'   Dim rsTemp As New Recordset
'   Dim nCountInfo As Integer
'   Dim SaleStaionInfo() As String
'   Dim i As Integer
'
'     On Error GoTo here
'
'
'   AssertHaveRight m_oActiveUser, RIGHT_REBusStopBus
'   oDb.ConnectionString = GetConnectionStr(cszRunEnv)
'
''    ST_TicketNormal = 1 '车票正常售出
''    ST_TicketSellChange = 2 ' 车票改签售出
''    ST_TicketCanceled = 4 '车票已废
''    ST_TicketChanged = 8 '被改签
''    ST_TicketReturned = 16 '车票已退
''    ST_TicketChecked = 32 '车票已检
''
'     szSql = " SELECT  DISTINCT ew.station_id ,staInfo.station_name,ISNULL(a.SellCount,0)as SellCount ," _
'        & " ISNULL(a.ticket_type,'100') as ticket_type "
'
'    If DateDiff("d", dtRunBusdate, Now) <= 0 Then
'          szSql = szSql & "  FROM  Work_env_bus_station_lst ew, station_info stainfo, "
'    Else
'          szSql = szSql & "  FROM  Env_bus_station_lst ew, station_info stainfo, "
'    End If
'
'    szSql = szSql & " ( " _
'          & " SELECT COUNT(*) as sellCount ,sa.ticket_type ," _
'          & "   sa.des_station_id " _
'          & " FROM  Ticket_sell_lst sa  " _
'          & " WHERE  bus_date='" & ToDBDate(dtRunBusdate) & " ' and bus_id='" & szBusID & "' and (sa.status<>17 and sa.status<>5  and sa.status<>9 ) " _
'          & " GROUP by sa.ticket_type, sa.des_station_id " _
'          & " ) a " _
'          & " WHERE ew.station_id=a.des_station_id and ew.bus_id='" & szBusID & "' " _
'          & " and ew.bus_date='" & ToDBDate(dtRunBusdate) & " '  and ew.station_id =staInfo.station_id " _
'          & " order by ew.station_id"
'
'
'   Set rsTemp = oDb.Execute(szSql)
'   nCountInfo = rsTemp.RecordCount
'   If nCountInfo <> 0 Then
'       ReDim SaleStaionInfo(0 To nCountInfo - 1, 0 To 3)
'
'       For i = 0 To nCountInfo - 1
'           SaleStaionInfo(i, 0) = rsTemp!station_id
'           SaleStaionInfo(i, 1) = rsTemp!station_name
'           SaleStaionInfo(i, 2) = rsTemp!SellCount
'           SaleStaionInfo(i, 3) = rsTemp!ticket_type
'          ' SaleStaionInfo(i, 4) = rsTemp!Seat_type
'          ' SaleStaionInfo(i, 5) = rsTemp!seat_type_name
'          rsTemp.MoveNext
'      Next
'
'   End If
'   GetBusStatinSellInfo = SaleStaionInfo
'   If Not rsTemp Is Nothing Then Set rsTemp = Nothing
'   Set oDb = Nothing
'   Exit Function
'here:
'    err.Raise err.Number
'    Set oDb = Nothing
'    If Not rsTemp Is Nothing Then Set rsTemp = Nothing
'End Function
'生成环境时，得到计划预信息
Private Function GetProjectBookInfo(dtRunBusdate As Date, szBusID As String) As String()

    Dim oDb As New RTConnection
    Dim i As String
    Dim nCount As Integer
    Dim szSql As String
    Dim rsTemp As New Recordset
    Dim szBookBusInfo(0 To 4) As String
    
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    szSql = " SELECT * FROM " _
            & " project_book_info WHERE " _
            & " start_bus_date<='" & ToDBDate(dtRunBusdate) & " '" _
            & " and end_bus_date>='" & ToDBDate(dtRunBusdate) & " '" _
            & " and bus_id='" & szBusID & "'"
    Set rsTemp = oDb.Execute(szSql)
    nCount = rsTemp.RecordCount
    If nCount = 0 Then Exit Function
    
    szBookBusInfo(0) = rsTemp!start_bus_date
    szBookBusInfo(1) = rsTemp!end_bus_date
    szBookBusInfo(2) = rsTemp!seat_start_no
    szBookBusInfo(3) = rsTemp!seat_book_Count
    szBookBusInfo(4) = rsTemp!remark_info
    
    GetProjectBookInfo = szBookBusInfo

End Function

'新增环境车次
Public Function AddEnviromentBus(ByVal pdtRunDate As Date, ByVal pdtRunTime As Date, ByVal pszBusID As String, szVehicleID As String, _
    szCheckGateID As String, szBusType As String, nInternetStatus As Integer, ByVal pbDelayStatus As Boolean, nScrollCheckTime As Integer, tTBusPriceInfo() As TBusPriceDetailInfo)
    Dim szSql As String
    Dim szRouteID As String
    Dim szMaxStationID As String
    Dim szVehicleTypeName As String
    Dim szVehicleTypeCode As String
    Dim nNewVehicleSeatCount As Integer
    Dim szOwnerID As String
    Dim szTransportCompanyID As String
    Dim szSplitCompanyID As String
    Dim rsTemp As New Recordset
    Dim tTVehicleSeatInfo(1 To 5) As TSeatInfoCount
    Dim nStation As Integer
    Dim szStationName As String
    
    Dim datatime As String
    
    Dim atBusPrice() As TREBusStationInfo
    Dim nBeginSeatNo As Integer
    
    
    Dim oDb As New RTConnection
    Dim LErr As Integer
    Dim bIsOnlyNormalSeatType As Boolean
    Dim nCount As Integer
    
    #If SUPPORT_MSG = 1 Then
    Dim oParam As New SystemParam
    Dim aszValues() As String
    #End If
    
    AssertActiveUserValid m_oActiveUser, ERR_REScheme
    AssertHaveRight m_oActiveUser, RIGHT_MakeRunEvironment
    
    nCount = ArrayLength(tTBusPriceInfo)
    If nCount = 0 Then Exit Function
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    '得到座位信息
    'tTVehicleSeatInfo = GetVehicleSeatInfo(odb, szVehicleID, nNewVehicleSeatCount, LErr)
    
    If LErr = 0 Then bIsOnlyNormalSeatType = True
    '取得站点票价信息
    atBusPrice = FormatTem(tTBusPriceInfo, pszBusID)
    '取得基本信息
    szSql = "SELECT start_seat_no,owner_id,seat_quantity,transport_company_id,split_company_id FROM Vehicle_info WHERE  vehicle_id='" & szVehicleID & "'"
    
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount = 0 Then Exit Function
    
    szRouteID = tTBusPriceInfo(nCount).szRouteID
    szMaxStationID = tTBusPriceInfo(nCount).szStationID
    szVehicleTypeName = tTBusPriceInfo(nCount).szVehicleModelName
    szVehicleTypeCode = tTBusPriceInfo(nCount).szVehicleModel
    szStationName = tTBusPriceInfo(nCount).szStationName
    nNewVehicleSeatCount = rsTemp!seat_quantity
    szOwnerID = rsTemp!owner_id
    szTransportCompanyID = rsTemp!transport_company_id
    szSplitCompanyID = rsTemp!split_company_id
    nBeginSeatNo = CInt(rsTemp!start_seat_no)
    Set rsTemp = Nothing
    
    '得到座位信息
    Dim tvTemp() As TVehcileSeatType
    tvTemp = GetBusVehicleSeatType(szVehicleID)
    
    'Public Const cszDateStr = "YYYY-MM-DD"
    'Public Const cszTimeStr = "HH:MM:SS"
    datatime = Format((Format(pdtRunDate, cszDateStr) & " " & Format(pdtRunTime, cszTimeStr)), "YYYY-MM-DD HH:MM:SS")
    '开始事务
'    oDb.BeginTrans
        'Inset Enveriment_bus_info
        szSql = "INSERT Env_bus_info( " _
            & "bus_date," _
            & "bus_id," _
            & "check_gate_id," _
            & "vehicle_type_name," _
            & "vehicle_type_code," _
            & "total_seat," _
            & "sale_seat_quantity," _
            & "total_stand_quantity," _
            & "sale_stand_seat_quantity,"
        szSql = szSql & "bus_start_time," _
            & "bus_type," _
            & "status," _
            & "route_id," _
            & "end_station_id," _
            & "end_station_name," _
            & "owner_id," _
            & "transport_company_id," _
            & "split_company_id, " _
            & "vehicle_id ," _
            & "baggage_number ," _
            & "fact_weight," _
            & "calculate_weight ," _
            & "over_weight_number ," _
            & "luggage_total_price," _
            & "ScrollBus_Check_Time ," _
            & "seat_remain ," _
            & "bed_remain ," _
            & "additional_remain ," _
            & "other_remain_1 ," _
            & "other_remain_2 ," _
            & "delay_status ," _
            & "internet_status) "
        
        szSql = szSql & " VALUES( '" _
            & ToDBDate(pdtRunDate) & "','" _
            & pszBusID & "','" _
            & szCheckGateID & "','" _
            & szVehicleTypeName & "','" _
            & szVehicleTypeCode & "'," _
            & nNewVehicleSeatCount & "," _
            & nNewVehicleSeatCount & "," _
            & 0 & "," _
            & 0 & ","
            szSql = szSql & "'" & datatime & "','" _
            & szBusType & "'," _
            & ST_BusStopped & ",'" _
            & szRouteID & "','" _
            & szMaxStationID & "','" _
            & szStationName & "','" _
            & szOwnerID & "','" _
            & szTransportCompanyID & "','" _
            & szSplitCompanyID & "','" _
            & szVehicleID & "'," _
            & 0 & "," _
            & 0 & "," _
            & 0 & "," _
            & 0 & "," _
            & 0 & "," _
            & nScrollCheckTime & "," _
            & 0 & "," _
            & 0 & "," _
            & 0 & "," _
            & 0 & "," _
            & 0 & "," _
            & TransFieldValueToString(pbDelayStatus) & "," _
            & nInternetStatus & " )"
        oDb.Execute szSql
        
        '插入环境站点表
        
        
        
        Dim i As Integer
        nStation = ArrayLength(atBusPrice)
        For i = 0 To nStation - 1
            szSql = "INSERT INTO Env_bus_station_lst( " _
                & " bus_date," _
                & " bus_id," _
                & " sell_station_id," _
                & " station_id," _
                & " station_serial_no," _
                & " seat_Type_id ," _
                & " max_sale_quantity," _
                & " sale_ticket_quantity," _
                & " book_count ," _
                & " stop_sale_time," _
                & " mileage," _
                & " full_price," _
                & " half_price," _
                & " preferential_ticket1," _
                & " preferential_ticket2," _
                & " preferential_ticket3 )"
            szSql = szSql & " VALUES('" _
                & ToDBDate(pdtRunDate) & "','" _
                & pszBusID & "','" _
                & atBusPrice(i).szSellStationID & "','" _
                & atBusPrice(i).szStationID & "'," _
                & CInt(atBusPrice(i).szStationSerialNo) & ",'" _
                & Trim(atBusPrice(i).szSeatTypeID) & "' ,"
            If Trim(atBusPrice(i).szStationID) <> szMaxStationID Then
                szSql = szSql & 0 & "," _
                & 0 & ","
            Else
                szSql = szSql & -1 & "," _
                & -1 & ","
            End If
            szSql = szSql & 0 & ","
            szSql = szSql & CInt(atBusPrice(i).sgLimitedSellTime) & "," _
                & atBusPrice(i).nMileage & "," _
                & atBusPrice(i).sgFullPrice & "," _
                & atBusPrice(i).sgHalfPrice & "," _
                & atBusPrice(i).sgPreferentialPrice1 & "," _
                & atBusPrice(i).sgPreferentialPrice2 & "," _
                & atBusPrice(i).sgPreferentialPrice3 & ")"
            oDb.Execute szSql
        Next i
        '插入环境票价表
        For i = 1 To nCount
            szSql = "INSERT INTO Env_bus_price_lst( " _
                & " bus_date," _
                & " bus_id," _
                & " sell_station_id," _
                & " station_id," _
                & " station_serial_no," _
                & " seat_type_id ," _
                & " ticket_type," _
                & " ticket_price_total," _
                & " base_carriage," _
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
                & " price_item_15) "
            szSql = szSql & " VALUES('" _
                & ToDBDate(pdtRunDate) & "','" _
                & pszBusID & "','" _
                & tTBusPriceInfo(i).szSellStationID & "','" _
                & tTBusPriceInfo(i).szStationID & "'," _
                & tTBusPriceInfo(i).nSerialNo & ",'" _
                & tTBusPriceInfo(i).szSeatTypeID & "'," _
                & tTBusPriceInfo(i).nTicketType & "," _
                & tTBusPriceInfo(i).sgTotalPrice & "," _
                & tTBusPriceInfo(i).sgBaseCarriage & "," _
                & tTBusPriceInfo(i).asgItem(1) & "," _
                & tTBusPriceInfo(i).asgItem(2) & "," _
                & tTBusPriceInfo(i).asgItem(3) & "," _
                & tTBusPriceInfo(i).asgItem(4) & "," _
                & tTBusPriceInfo(i).asgItem(5) & "," _
                & tTBusPriceInfo(i).asgItem(6) & "," _
                & tTBusPriceInfo(i).asgItem(7) & "," _
                & tTBusPriceInfo(i).asgItem(8) & "," _
                & tTBusPriceInfo(i).asgItem(9) & "," _
                & tTBusPriceInfo(i).asgItem(10) & "," _
                & tTBusPriceInfo(i).asgItem(11) & "," _
                & tTBusPriceInfo(i).asgItem(12) & "," _
                & tTBusPriceInfo(i).asgItem(13) & "," _
                & tTBusPriceInfo(i).asgItem(14) & " ," _
                & tTBusPriceInfo(i).asgItem(15) & ") "
            oDb.Execute szSql
        Next
        '插入环境车次配载表
        szSql = "INSERT Env_bus_allot_lst(bus_date, bus_id, sell_station_id, check_gate_id, bus_start_time, Status, can_sale_quantity, sell_quantity) " _
            & " SELECT DISTINCT " & TransFieldValueToString(pdtRunDate) _
            & " , " & TransFieldValueToString(pszBusID) _
            & " , sell_station_id " _
            & " , " & TransFieldValueToString(szCheckGateID) _
            & " , " & TransFieldValueToString(datatime) _
            & " , " & TransFieldValueToString(ST_BusStopped) _
            & " , -1 " _
            & " , 0 " _
            & " FROM Route_section_lst where route_id = " & TransFieldValueToString(szRouteID)
        oDb.Execute (szSql)
        
        
        '插入环境车次售票点表
        szSql = "INSERT env_bus_sell_station_lst(bus_date, bus_id, sell_station_id , can_sale_quantity, sell_quantity) " _
            & " SELECT  " & TransFieldValueToString(pdtRunDate) _
            & " , " & TransFieldValueToString(pszBusID) _
            & " , sell_station_id " _
            & " , -1 " _
            & " , 0 " _
            & " FROM sell_station_info "
        oDb.Execute (szSql)
        
        
        
        '环境座位表
        Dim nSeatTemp As Integer
        Dim nTotalSeat As Integer
        Dim szSeatTypeID As String
        nTotalSeat = nNewVehicleSeatCount
        Dim nTemp1 As Integer, nTemp2 As Integer, nToalseatTemp As Integer
        For i = 1 To nNewVehicleSeatCount
            nTemp1 = ST_SeatCanSell
            nTemp2 = i - 1 + nBeginSeatNo
            nToalseatTemp = nToalseatTemp + 1
            If nToalseatTemp > nTotalSeat Then Exit For
            '取座位类型
            If ArrayLength(tvTemp) = 0 Then
                szSeatTypeID = cszSeatTypeIsNormal
            Else
                szSeatTypeID = GetSeatTypeID(nTemp2, tvTemp)
                If szSeatTypeID = "" Then szSeatTypeID = cszSeatTypeIsNormal
            End If
            szSql = "INSERT Env_bus_seat_lst( " _
                & "bus_date," _
                & "bus_id," _
                & "seat_no," _
                & "status," _
                & "ticket_no," _
                & "seat_type_id) "
            szSql = szSql & " VALUES('" _
                & ToDBDate(pdtRunDate) & "','" _
                & pszBusID & "','" _
                & BuildSeatNo(nTemp2) & "'," _
                & nTemp1 & ",'','" _
                & szSeatTypeID & "')"
            oDb.Execute szSql
        Next i
        
        '更新环境车次表座位总数
        UpdateEnviromentSeatCountEx pszBusID, pdtRunDate, oDb
    
'    oDb.CommitTrans
    Dim szlog As String
    szlog = "环境中新增车次[" & pszBusID & "]" & "发车时间为[" & pdtRunDate & "]" & "终点站为[" & szStationName & "]"
    WriteOperateLog m_oActiveUser, RIGHT_MakeRunEvironment, szlog
    
    #If SUPPORT_MSG = 1 Then
    ReDim aszValues(1 To 3)
    aszValues(1) = pszBusID
    aszValues(2) = ToDBDate(pdtRunDate)
    aszValues(3) = Trim(szCheckGateID)
    oParam.Init m_oActiveUser
    m_oSndMsg.SendMsg oParam.NowDateTime, cszAddBus, aszValues
    'm_bCheckGateChanged = False
    'm_bOffTimeChanged = False
    #End If
    Exit Function
    oDb.RollbackTrans
    ShowError err.Number, , err.Description

End Function

Private Function FormatTem(tTBusPriceInfo() As TBusPriceDetailInfo, szBusID As String) As TREBusStationInfo()
Dim szSql As String
Dim szRouteID As String
Dim szMaxStationID As String
Dim tTVehicleSeatInfo() As TSeatInfoCount
Dim i As Integer, j As Integer, k As Integer
Dim tTBusStation() As TREBusStationInfo

'''''
Dim nCount As Integer, nCountStation As Integer
Dim szStationIDTemp As String
Dim szSeatTypeID As String
Dim szSellStationID As String
'''''
'ReDim m_tSeatTypeInfo(1 To 5) As TSeatInfoCount
nCount = ArrayLength(tTBusPriceInfo)

szStationIDTemp = tTBusPriceInfo(1).szStationID
szSeatTypeID = tTBusPriceInfo(1).szSeatTypeID

szSellStationID = tTBusPriceInfo(1).szSellStationID
ReDim tTBusStation(0 To 0)
nCountStation = 1
    For i = 1 To nCount
        If Trim(szSellStationID) <> Trim(tTBusPriceInfo(i).szSellStationID) Or Trim(szStationIDTemp) <> Trim(tTBusPriceInfo(i).szStationID) Or Trim(szSeatTypeID) <> Trim(tTBusPriceInfo(i).szSeatTypeID) Then
           ReDim Preserve tTBusStation(0 To nCountStation)
           nCountStation = nCountStation + 1
           szSeatTypeID = tTBusPriceInfo(i).szSeatTypeID
           szStationIDTemp = tTBusPriceInfo(i).szStationID
           szSellStationID = tTBusPriceInfo(i).szSellStationID
        End If
        
        tTBusStation(nCountStation - 1).szSellStationID = tTBusPriceInfo(i).szSellStationID
        tTBusStation(nCountStation - 1).szSellStationName = tTBusPriceInfo(i).szSellStationName
        
        'tTBusStation(nCountStation - 1).dtBusDate = Date
        Select Case tTBusPriceInfo(i).nTicketType
        Case TP_FullPrice
            tTBusStation(nCountStation - 1).sgFullPrice = tTBusPriceInfo(i).sgTotalPrice
            tTBusStation(nCountStation - 1).sgBasePrice = tTBusPriceInfo(i).sgTotalPrice
            tTBusStation(nCountStation - 1).sgHalfPrice = tTBusPriceInfo(i).sgTotalPrice
            
            
        End Select
'        tTBusStation(nCountStation - 1).dTicket(tTBusPriceInfo(i).nTicketType) = tTBusPriceInfo(i).sgTotalPrice
        'tTBusStation(nCountStation - 1).nBookCount = 0
'        tTBusStation(nCountStation - 1).nMaxSaleQuantity = -1
        tTBusStation(nCountStation - 1).nLimitedSellCount = -1
        tTBusStation(nCountStation - 1).nMileage = tTBusPriceInfo(i).sgMileage
        tTBusStation(nCountStation - 1).sgLimitedSellTime = -1
'        tTBusStation(nCountStation - 1).szBusID = szBusID
        tTBusStation(nCountStation - 1).szSeatTypeID = tTBusPriceInfo(i).szSeatTypeID
        'm_tSeatTypeInfo(CInt(tTBusPriceInfo(i).szSeatTypeID)) = m_tSeatTypeInfo(CInt(tTBusPriceInfo(i).szSeatTypeID)) + 1
        tTBusStation(nCountStation - 1).szStationID = tTBusPriceInfo(i).szStationID
        tTBusStation(nCountStation - 1).szStationSerialNo = tTBusPriceInfo(i).nSerialNo
   Next
 FormatTem = tTBusStation
End Function

'得到车次票价信息
Public Function GetBusTicketInfoRS(pdyBusDate As Date, paszBusID() As String) As Recordset
    Dim szSql As String
    'Dim atTicketPrice() As TRETicketPriceEx
    Dim rsTemp As New Recordset
    Dim oDb As New RTConnection
    Dim i As Integer
    Dim project As TSchemeArrangement
    Dim tBusPriceDetail() As TBusPriceDetailInfo
    Dim plErr As Long
    Dim nCount As Integer
    Dim szIn As String
    
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    
    '删除了DISTINCT ,加入tbs.sell_station_id,tbs.sell_station_name,和FROM处加入了sell_station_info tbs
    
    szSql = "SELECT p.* , tbs.sell_station_name, area_code, e.route_id,t.seat_type_name,vehicle_type_name,vehicle_type_code,route_name, " _
          & " station_name , w.mileage ,e.bus_type,e.bus_start_time , y.ticket_type_name " _
          & " FROM Seat_type_code t,route_info r, Station_info s, Ticket_type_code y ,sell_station_info tbs " _
          
    If DateDiff("d", pdyBusDate, Now) <= 0 Then  '如果查询的是环境中的车次则使用工作表
        szSql = szSql & ",Work_env_bus_station_lst w,Work_env_bus_price_lst p ,Work_env_bus_info e "
    Else
        szSql = szSql & ",Env_bus_price_lst p, Env_bus_info e  , Env_bus_station_lst w  "
    End If
'    project = SelfGetExecuteBusProject(ToDBDate(pdyBusDate), plErr)
    nCount = ArrayLength(paszBusID)
    If nCount > 0 Then
        szIn = " AND p.bus_id IN ( "
        For i = 1 To nCount - 1
            szIn = szIn & "'" & paszBusID(i, 1) & "',"
        Next i
        szIn = szIn & "'" & paszBusID(i, 1) & "')"
    End If
    szSql = szSql & " WHERE tbs.sell_station_id=w.sell_station_id AND p.seat_type_id = t.seat_type_id " _
        & " And e.bus_date = p.bus_date And e.bus_id = p.bus_id " _
        & " AND r.route_id=e.route_id " _
        & " AND s.station_id=p.station_id " _
        & " AND w.station_id=p.station_id " _
        & " AND w.bus_id=p.bus_id " _
        & " AND w.bus_date=p.bus_date AND w.sell_station_id = p.sell_station_id " _
        & szIn & " AND p.bus_date='" & ToDBDate(pdyBusDate) & "' " _
        & " AND y.ticket_type_id = p.ticket_type AND w.seat_type_id = p.seat_type_id " _
        & " ORDER BY  p.bus_date,p.bus_id, p.sell_station_id,  p.station_serial_no,p.seat_type_id,ticket_type "
    Set rsTemp = oDb.Execute(szSql)


    Set GetBusTicketInfoRS = rsTemp
    Set rsTemp = Nothing
    Set oDb = Nothing
End Function




'得到车次信息
Public Function GetBusInfoRsReport(busdate As Date, CheckGateID As String, szSellStationID As String) As Recordset
'    AssertActiveUserValid m_oActiveUser, cnErroeStartNo
    
    Dim oDb As New RTConnection
    Dim rsTemp  As New Recordset
    Dim szSql As String
    oDb.ConnectionString = GetConnectionStr(cszSystemMan)
       
    If Len(CheckGateID) = 0 Then '未指定检票口
'        ShowError ERR_ChkTkCheckGateIDNotAppointed
        Exit Function
    End If
    If DateDiff("d", busdate, Now) <= 0 Then '如果查询的是环境中的车次则使用工作表
        szSql = "SELECT a.check_gate_id,convert(char(5), a.bus_start_time,108)  bus_start_time,e.bus_type,v.license_tag_no,c.transport_company_short_name," _
            & " e.bus_id bus_id " _
            & " ,o.owner_name,a.status,0 AS bus_serial_no,v.vehicle_id,c.transport_company_id,o.owner_id,e.end_station_name " _
            & " FROM Work_env_bus_info e,Work_env_bus_allot_lst a,Vehicle_info v,Company_info c,owner_info o"
    Else
        szSql = "SELECT a.check_gate_id,convert(char(5), a.bus_start_time,108)  bus_start_time,e.bus_type,v.license_tag_no,c.transport_company_short_name," _
           & " e.bus_id bus_id " _
           & " ,o.owner_name,a.status,0 AS bus_serial_no,v.vehicle_id,c.transport_company_id,o.owner_id ,e.end_station_name" _
            & " FROM Env_bus_info e,Env_bus_allot_lst a,Vehicle_info v,Company_info c,owner_info o"
    End If
    szSql = szSql & " WHERE e.vehicle_id = v.vehicle_id  AND a.can_sale_quantity <>0 AND e.transport_company_id=c.transport_company_id  AND e.owner_id=o.owner_id " _
                & " AND  a.sell_station_id = " & TransFieldValueToString(szSellStationID) _
                & " AND a.bus_date=e.bus_date  AND a.bus_id=e.bus_id" _
                & " AND a.check_gate_id= " & TransFieldValueToString(CheckGateID) _
                & " AND e.bus_date= " & TransFieldValueToString(busdate) _
                & " AND e.status=0 " _
                & " ORDER BY  bus_start_time,e.bus_id"
    
    Set rsTemp = oDb.Execute(szSql, , -1)
    If rsTemp.RecordCount = 0 Then '无记录
        'ShowError ERR_ChkTkResultIsNull
        Set GetBusInfoRsReport = rsTemp
        Exit Function
    End If
    Set GetBusInfoRsReport = rsTemp
    
End Function

'得到上车站环境车次
Public Function GetRESellStationBusReport(szSellStationID As String, dtpbusdate As Date, Optional szFindString As String = "") As Recordset
    '***************************************************************************************************************
    'szSellStationID 上车站
    'dtpbusdate 发车日期
    'szFindString 查询条件，根据不同的条件进行查询，如szFindString为空，则查询所有的车次，否则查询所有正常运行的车次
    '***************************************************************************************************************
    
    Dim oDb As New RTConnection
    Dim szSql As String
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)

    szSql = " SELECT  convert(char(5),ebal.bus_start_time ,108) bus_start_time ,ebi.vehicle_type_code,ebi.route_id,ebi.sale_seat_quantity," _
           & " ebi.bus_id bus_id," _
           & " ebi.end_station_name,cg.check_gate_name,ebi.vehicle_type_name,ci.transport_company_short_name,bi.bus_run_cycle," _
           & " ebi.total_seat ,case ebi.status when 0 then '' when 1 then '停班' end status,ri.route_name,ri.mileage,  'A'  dj,ebp.ticket_price_total," _
           & " ebi.split_company_id fx " _
           & " FROM work_env_bus_info ebi ,work_env_bus_allot_lst ebal,checkgate_info cg,company_info ci,route_info ri,env_bus_price_lst ebp,bus_info bi " _
           & " WHERE ebi.bus_id=ebal.bus_id AND ebi.bus_date=ebal.bus_date AND ebi.bus_date=" & TransFieldValueToString(dtpbusdate) _
           & " AND ebal.sell_station_id= " & TransFieldValueToString(szSellStationID) _
           & " AND  ebi.transport_company_id=ci.transport_company_id " _
           & " AND ebal.check_gate_id=cg.check_gate_id AND ebi.route_id=ri.route_id " _
           & " AND ebi.bus_date=ebp.bus_date AND ebi.bus_id=ebp.bus_id AND ebp.station_id=ebi.end_station_id " _
           & " AND ebp.sell_station_id=ebal.sell_station_id AND ebp.ticket_type=1 AND bi.bus_id=ebi.bus_id "
    If szFindString <> "" Then
        szSql = szSql & szFindString
    End If
        szSql = szSql & " ORDER  BY ebal.bus_start_time"
    Set GetRESellStationBusReport = oDb.Execute(szSql)
End Function


'得到指定日期的计划车次

Public Function GetPlanSellStationBusReport(szSellStationID As String, Optional szFindString As String = "") As Recordset
    Dim oDb As New RTConnection
    Dim szSql As String
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
       
 szSql = " SELECT bi.bus_id bus_id ,bi.bus_run_cycle , " _
        & " Convert(char(5),bai.bus_start_time ,108) bus_start_time,si.station_name,cg.check_gate_name,vtc.vehicle_type_short_name,ci.transport_company_short_name,bpl.mileage,vi.license_tag_no ," _
        & " base_carriage +price_item_1 +price_item_2+ price_item_3 +price_item_4+ price_item_5+ price_item_6+ price_item_7+ price_item_8+ price_item_9+ price_item_10+ price_item_11+ price_item_12+ price_item_13+ price_item_14 + price_item_15 total_price ,vi.seat_quantity- isnull(brsl.seat_quantity,0) can_sell_seat,vi.seat_quantity total_seat," _
        & " case  when getdate() >= bi.stop_start_date   AND getdate()<= bi.stop_end_date then '停班' else  '' end  status " _
        & " FROM bus_info bi,bus_vehicle_code bv,vehicle_info vi,bus_allot_info bai,route_info ri,station_info si,bus_reserve_seat_lst brsl," _
        & " checkgate_info cg,vehicle_type_code vtc,company_info ci,bus_price_lst bpl " _
        & " WHERE bi.bus_id=bv.bus_id AND bv.vehicle_id=vi.vehicle_id AND " _
        & " bi.bus_id = bai.bus_id And bi.route_id = ri.route_id " _
        & " AND bai.sell_station_id= " & TransFieldValueToString(szSellStationID) _
        & " AND si.station_id=ri.end_station_id " _
        & " AND  bai.check_gate_id =cg.check_gate_id " _
        & " AND vtc.vehicle_type_code=vi.vehicle_type_code " _
        & " AND ci.transport_company_id=vi.transport_company_id " _
        & " AND bv.bus_id *= brsl.bus_id AND bv.vehicle_serial *= brsl.vehicle_serial " _
        & " AND bi.bus_id=bpl.bus_id AND bpl.sell_station_id=bai.sell_station_id AND bpl.station_id=ri.end_station_id AND bpl.ticket_type=1  AND bpl.vehicle_type_code=vtc.vehicle_type_code "
  If szFindString <> "" Then
    szSql = szSql & szFindString
  End If
  szSql = szSql & " ORDER BY bai.bus_start_time "
       
Set GetPlanSellStationBusReport = oDb.Execute(szSql)
End Function

'得到班次资料
Public Function GetBusInfo() As Recordset
    Dim oDb As New RTConnection
    Dim szSql As String
    oDb.ConnectionString = GetConnectionStr(cszRunEnv)
    
    szSql = "select a.*,b.station_name from (SELECT bsi.bus_id,r.route_name,bsi.station_id,bsi.max_sale_quantity,bsi.stop_sale_time,bsi.mileage," _
            & " bai.sell_station_id ,s.station_name sell_station, rtrim(bai.check_gate_id)+'['+rtrim(check_gate_name)+']' AS check_gate, convert(char(5),bai.bus_start_time,108) bus_start_time, bai.can_sale_quantity " _
            & " FROM bus_station_info bsi,bus_allot_info bai,route_info r,bus_info bi,station_info s,checkgate_info c " _
            & " WHERE  bsi.bus_id = bai.bus_id and bi.bus_id=bsi.bus_id and bi.route_id=r.route_id " _
            & " and bai.sell_station_id=s.station_input_code and c.check_gate_id=bai.check_gate_id) a,station_info b" _
            & " where a.station_id=b.station_id " _
            & " order by bus_id,sell_station_id,mileage"
    Set GetBusInfo = oDb.Execute(szSql)
End Function
