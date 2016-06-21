VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BusStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Base 0

Const TRADE_OK = "OK"
Const SERVER_CONNECTED = "接入"
Const SERVER_RELEASE = "断开"
Const ERR_CONNECT = "3001"
Const ERR_STATIONS = "3002"
Const ERR_SCHEDULES = "3003"
Const ERR_NOTICKETID = "3004"
Const ERR_NOSCHEDULEID = "2107"
Const ERR_DAILYREC = "3006"
Const ERR_NOSTARTID = "3007"
Const ERR_OPERATOR = "3008"
Const ERR_ERRDATE = "3010"
Const ERR_ValiDate = "3011" '错误的验证码(网上售票用)
Const ERR_ERRFIND = "3012"

Const cnIniPosStartStationSerial = 1
Const cnIniPosStartStationID = 2
Const cnIniPosStartStationName = 3
Const cnIniPosSellStationID = 4



'Private moGate(0 To cnSellStationCount - 1) As Object
'Private moUser(0 To cnSellStationCount - 1) As Object
'Private m_oSellTicketService(0 To 3) As Object
'Private LoginId1 As Long
Public LocalCn As ADODB.Connection
'车站个数,此处临时用一下,到时改为从配置文件中读取
'Const cnSellStationCount = 5

Private aszStationID() As String * cnStationLenID
Private aszStationInputID() As String * cnStationLenInputID
Private aszStationName() As String * cnStationLenName
Private m_oActiveUser() As Object
Private m_oSellTicketService() As Object
Private m_oSystemParam() As Object
Private m_oServiceTicket() As Object

Private m_oInternetSellTicketService() As Object
Private m_oInternetServiceTicket() As Object
Private m_oInternetActiveUser() As Object
Private m_oInternetSystemParam() As Object


Private g_aszAllStartStation() As String '起点站
Private m_nSellStationCount As Integer '站点个数



Private m_szDatabaseType As String
Private m_szServer As String
Private m_szUser As String
Private m_szPassword As String
Private m_szDatabase As String
Private m_szTimeout As String

Private odb As ADODB.Connection

Private Sub Class_Initialize()
    Dim i As Integer
    Dim FileNo As Integer

    '====================================================================
    '读取自定义数据
    '====================================================================
    FileNo = FreeFile
    Open App.Path + "\StartStation.ini" For Input As #FileNo
        Input #FileNo, m_nSellStationCount '车站个数
        If m_nSellStationCount > 0 Then
            ReDim g_aszAllStartStation(1 To m_nSellStationCount, 1 To 4)
            For i = 1 To m_nSellStationCount
                Input #FileNo, g_aszAllStartStation(i, cnIniPosStartStationSerial) '起点站序号
                Input #FileNo, g_aszAllStartStation(i, cnIniPosStartStationID) '起点站代码
                Input #FileNo, g_aszAllStartStation(i, cnIniPosStartStationName) '起点站名
                Input #FileNo, g_aszAllStartStation(i, cnIniPosSellStationID) '起点站对应的上车站代码
            Next i
        End If
        If err.Number <> 0 Then
            err.Raise err.Number, "配置文件StartStation.ini错误"
            
        End If
    Close #FileNo
    
    
    FileNo = FreeFile
    Open App.Path + "\Param.ini" For Input As #FileNo
        On Error Resume Next
        '自定义数据
        '去掉无用的数据
        Input #FileNo, m_szDatabaseType '远程机器
        Input #FileNo, m_szDatabaseType '远程端口
        Input #FileNo, m_szDatabaseType  '是否使用快速打印
        
        '以下是有用的数据
        Input #FileNo, m_szDatabaseType '
        Input #FileNo, m_szServer '
        Input #FileNo, m_szUser '
        Input #FileNo, m_szPassword '
        Input #FileNo, m_szDatabase '
        Input #FileNo, m_szTimeout '
        If err.Number <> 0 Then
            
            err.Raise err.Number, "配置文件Param.ini错误"
        End If
    Close #FileNo
    
    
    
    Set LocalCn = New ADODB.Connection
    LocalCn.ConnectionString = GetAdodcConnectionStr
    LocalCn.Open
    
    Set odb = New ADODB.Connection
    odb.ConnectionString = GetConnectionStr
    odb.Open
    
    If m_nSellStationCount > 0 Then
        ReDim aszStationID(0 To m_nSellStationCount - 1)
        ReDim aszStationInputID(0 To m_nSellStationCount - 1)
        ReDim aszStationName(0 To m_nSellStationCount - 1)
        ReDim m_oActiveUser(0 To m_nSellStationCount - 1)
        ReDim m_oSellTicketService(0 To m_nSellStationCount - 1)
        ReDim m_oSystemParam(0 To m_nSellStationCount - 1)
        ReDim m_oServiceTicket(0 To m_nSellStationCount - 1)
        ReDim m_oInternetSellTicketService(0 To m_nSellStationCount - 1)
        ReDim m_oInternetActiveUser(0 To m_nSellStationCount - 1)
        ReDim m_oInternetSystemParam(0 To m_nSellStationCount - 1)
        ReDim m_oInternetServiceTicket(0 To m_nSellStationCount - 1)
    End If
    '====================================================================
End Sub

'联接车站
Public Function ConnectStation(Index As Integer, ServerName As String, UserName As String, _
Password As String, CardNo As String, Company As String) As String
    Dim szErrStr As String * 4
'    Dim i As Integer
 
    '创建对象
    Set m_oActiveUser(Index) = CreateObject("STSystemBase.ActiveUser")
    Set m_oSellTicketService(Index) = CreateObject("STTkSvr.SellTicketService")
    Set m_oSystemParam(Index) = CreateObject("STSystemBase.SystemParam")
    Set m_oServiceTicket(Index) = CreateObject("STTkSvr.ServiceTicket")
    
'    Set m_oInternetActiveUser(Index) = CreateObject("FS6SystemBase.ActiveUser", ServerName)
'    Set m_oInternetSellTicketService(Index) = CreateObject("ST6TkSvr.NetSellTicketService", ServerName)
'    Set m_oInternetSystemParam(Index) = CreateObject("FS6SystemBase.SystemParam", ServerName)
'    Set m_oInternetServiceTicket(Index) = CreateObject("ST6TkSvr.ServiceTicket", ServerName)
    '用户登陆
    m_oActiveUser(Index).Login UserName, Password, ServerName
    m_oSellTicketService(Index).Init m_oActiveUser(Index)
    m_oServiceTicket(Index).Init m_oActiveUser(Index)
    aszStationID(Index) = Trim(m_oSellTicketService(Index).StationID)
    aszStationName(Index) = Trim(m_oSellTicketService(Index).StationName)
    
'    m_oInternetActiveUser(Index).Login UserName, Password, ServerName

'    m_oInternetSellTicketService(Index).Init m_oActiveUser(Index)
'    m_oInternetServiceTicket(Index).Init m_oInternetActiveUser(Index)
'    GetAllStartStationID
    
    ConnectStation = cszCorrectRetCode
    Exit Function
ConnectErrhandle:
    Set m_oActiveUser(Index) = Nothing
    Set m_oSellTicketService(Index) = Nothing
    szErrStr = Trim(Str(err.Number))
    ConnectStation = szErrStr
End Function


Public Function GetStationsStr(i As Integer, Optional ByRef a As Integer, Optional ByRef Bcode As String) As String
    Dim szStr As String
    Dim szName As String
'    Dim i As Integer
    Dim rsTemp As Recordset
On Error GoTo here
    szStr = ""
    
    '再把当前车站的所有站点加上去
    If Not m_oSellTicketService(i) Is Nothing Then
        '得到当前车站的所有站点
        Set rsTemp = m_oSellTicketService(i).GetAllStationRs
        szStr = ConvertStationRSToString(rsTemp)
        a = rsTemp.RecordCount
        'Debug.Print szStr
    Else
        Bcode = ERR_CONNECT & FormatLen("车站未接入", cnLenReserved) & "@"
    End If
    GetStationsStr = szStr
    Exit Function
here:
    If Trim(Str(err.Number)) = 0 Then
        Bcode = "1233" & FormatLen("连接错误", cnLenReserved) & "@"
    Else
        Bcode = Trim(Str(Old2New(err.Number))) & FormatLen(MidA(err.Description, 1, cnLenReserved), cnLenReserved) & "@"
    End If
End Function

Public Function GetCheckGateStr(i As Integer) As String
    Dim szStr As String
    Dim szName As String
'    Dim i As Integer
    Dim rsCheckgate() As String
    szStr = ""
    
    If Not m_oSellTicketService(i) Is Nothing Then
        '得到当前车站的所有检票口
         rsCheckgate = m_oSellTicketService(i).GetAllCheckGate
        szStr = ConvertCheckSTRToString(rsCheckgate)
        
        'Debug.Print szStr
    End If
    GetCheckGateStr = szStr
End Function


Public Function GetSchedulesStr(i As Integer, pdyBusDate As Date, pszEndStationID As String, RxStr As String, Optional ByRef b As Integer, Optional ByRef pszCode As String) As String
    Dim szStr As String
    Dim rsTemp As Recordset
    Dim nCount As Integer
    Dim szSellStationID As String
On Error GoTo here
    szStr = ""
    
    szSellStationID = Trim(MidA(RxStr, 36, 8))
    If Not m_oSellTicketService(i) Is Nothing Then
        '得到经过该站点的车次
        nCount = 0
        Set rsTemp = m_oSellTicketService(i).GetBusRs(pdyBusDate, pszEndStationID)
        szStr = ConvertBusRSToString(rsTemp, nCount, szSellStationID)
        b = rsTemp.RecordCount - nCount
    Else
        pszCode = ERR_CONNECT & FormatLen("车站未接入", cnLenReserved) & "@"
    End If
    
    GetSchedulesStr = szStr
    Exit Function
here:
    If Trim(Str(err.Number)) = 0 Then
        pszCode = "1233" & FormatLen("连接错误", cnLenReserved) & "@"
    Else
        pszCode = Trim(Str(Old2New(err.Number))) & FormatLen(MidA(err.Description, 1, cnLenReserved), cnLenReserved) & "@"
    End If
End Function


'得到站点
Public Function GetAllStations(i As Integer, Optional IsToday As Boolean, Optional RxStr As String, Optional TxStr As String) As String
    Dim szStr As String
    Dim dyDate As Date
    
    Dim rsRec As ADODB.Recordset
    Dim szTemp As String '转换字符串
    Dim rsStsRec As ADODB.Recordset
    Dim szErrStr As String * 4
    i = CInt(Mid(RxStr, 58, 9))
    Set rsStsRec = New ADODB.Recordset
    rsStsRec.CursorType = adOpenKeyset
    rsStsRec.LockType = adLockOptimistic
    szStr = ERR_STATIONS
    On Error GoTo GetSationsErrHandle
    If m_oSellTicketService(i) Is Nothing Then
        TxStr = Left(RxStr, 243 - 1) & ERR_CONNECT & Mid(RxStr, 247, 80) & "@"
        GetAllStations = ERR_CONNECT
        Exit Function
    End If
    If Not m_oSellTicketService(i) Is Nothing Then
        rsStsRec.Open "delete from Stations where StartNo=" & Str(i), LocalCn, , , adCmdText
        rsStsRec.Open "Select  * from Stations where StartNo=" & Str(i), LocalCn, , , adCmdText
        dyDate = m_oSystemParam(i).NowDate
        If Not IsToday Then
            dyDate = DateAdd("d", 1, dyDate)
        End If
        szStr = Trim(Str(m_oSellTicketService(i).StationID))
        Set rsRec = m_oSellTicketService(i).GetStationRs(dyDate)
         
        If Not rsRec.BOF Then
            'For i = 1 To rsRec.RecordCount
            'Debug.Print rsRec!station_name
            'rsRec.MoveNext
            'Next
            rsRec.MoveFirst
        End If
        With rsRec
        Do While Not .EOF
            If rsStsRec.EOF Then
                rsStsRec.AddNew
            End If
            rsStsRec!StandardID = Trim(.Fields("station_id"))
            szTemp = .Fields("station_name")
            rsStsRec!Name = szTemp
            rsStsRec!InputID = Trim(.Fields("station_input_code"))
            rsStsRec!restime = LIMITTIME
            rsStsRec!ResNum = LIMITTICKETS
            rsStsRec!StartID = szStr
            rsStsRec!StartNo = i
            .MoveNext
            rsStsRec.Update
            rsStsRec.MoveNext
        Loop
        End With
        Set rsRec = Nothing
        rsStsRec.Close
    End If
    TxStr = Left(RxStr, 243 - 1) & cszCorrectRetCode & FormatLen(MidA(err.Description, 1, 80), 80) & "@"
    GetAllStations = cszCorrectRetCode
    Exit Function
GetSationsErrHandle:
    
    szErrStr = Trim(Str(Old2New(err.Number)))
    Debug.Print err.Number & "---" & err.Description
    TxStr = Left(RxStr, 243 - 1) & szErrStr & FormatLen(MidA(err.Description, 1, 80), 80) & "@"
    GetAllStations = szErrStr
End Function

'得到车次
Public Function GetAllSchedules(StartNo As Integer, Optional IsToday As Boolean, Optional RxStr As String, Optional TxStr As String) As String
    Dim rsStsRec As ADODB.Recordset
    Dim rsSchsRec As ADODB.Recordset
    Dim rsSchRec As ADODB.Recordset
    Dim dyDate As Date
    Dim tmpNo As Integer
    Dim tmpCount As Integer
    Dim szStr As String
    Dim nGetDay As String
    Dim i  As Integer
    Dim szErrStr As String * 4
    On Error GoTo GetSchedulesErrHandle
    
    nGetDay = CInt(MidA(RxStr, 247, 2))
    StartNo = CInt(Mid(RxStr, 58, 9))
    Set rsStsRec = New ADODB.Recordset
    Set rsSchsRec = New ADODB.Recordset
    rsStsRec.CursorType = adOpenDynamic
    rsStsRec.LockType = adLockOptimistic
    rsStsRec.Open "Select  * from Stations  where startno=" & Str(StartNo) & " order by Standardid", LocalCn, , , adCmdText
    rsSchsRec.CursorType = adOpenKeyset
    rsSchsRec.LockType = adLockOptimistic
    rsSchsRec.Open "delete from Schedules ", LocalCn, , , adCmdText
    rsSchsRec.Open "select * from Schedules where startno=" & Str(StartNo) & " order by StandardID ", LocalCn, , , adCmdText
    tmpCount = 1
    If (Not rsStsRec.BOF) And (rsStsRec.RecordCount <> 0) Then
        rsStsRec.MoveFirst
    End If
    szStr = ERR_SCHEDULES
    dyDate = m_oSystemParam(StartNo).NowDate
    If Not IsToday Then
'        dyDate = DateAdd("d", 1, dyDate)
    End If
    Dim aaa
    If m_oSellTicketService(StartNo) Is Nothing Then
        TxStr = Left(RxStr, 243 - 1) & ERR_CONNECT & Mid(RxStr, 247, 80) & "@"
        GetAllSchedules = ERR_CONNECT
        Exit Function
    End If

    Do While Not rsStsRec.EOF
        aaa = rsStsRec.RecordCount
        If Not m_oActiveUser(StartNo) Is Nothing Then
            For i = 0 To nGetDay - 1
            Set rsSchRec = m_oSellTicketService(StartNo).GetBusRs(DateAdd("d", i, dyDate), rsStsRec.Fields("StandardID").Value)
            If (Not rsSchRec.BOF) And rsSchRec.RecordCount <> 0 Then
                rsSchRec.MoveFirst
            End If
            With rsSchRec
                Do While Not .EOF
                    If IsNumeric(.Fields("Bus_id")) And (.Fields("bus_kind")) = TP_RegularBus Then
                        If rsSchsRec.EOF Then
                            rsSchsRec.AddNew
                        End If
                        rsSchsRec!Number = StartNo
                        '       rsSchsRec!InputID = rsStsRec!InputID
                        rsSchsRec!StandardID = rsStsRec!StandardID
                        
                        rsSchsRec!scheduleid = Trim(.Fields("Bus_id"))
                        rsSchsRec!BusName = .Fields("Vehicle_type_name")
                        rsSchsRec!CheckDoor = .Fields("check_gate_id")
                        
                        rsSchsRec!BusID = cszCorrectRetCode
                        rsSchsRec!BusType = Left(Trim(.Fields("Vehicle_type_code")), 2)
                        rsSchsRec!StartTime = CDate(.Fields("busstarttime"))
                        rsSchsRec!FPrice = .Fields("full_price")
                        rsSchsRec!HPrice = .Fields("half_price")
                        rsSchsRec!StartID = rsStsRec!StartID
                        rsSchsRec!destname = rsStsRec!Name
                        rsSchsRec!StartNo = rsStsRec!StartNo
                        rsSchsRec!startname = Trim(.Fields("sell_station_id"))
                        'tmpCount = tmpCount + 1
                        rsSchsRec.MoveNext
                    End If
                    
                    .MoveNext
                Loop
            End With
            Next i
        End If
        
        rsStsRec.MoveNext
    Loop
    
    rsSchsRec.Close
    rsStsRec.Close

    TxStr = Left(RxStr, 243 - 1) & cszCorrectRetCode & FormatLen(MidA(err.Description, 1, 80), 80) & "@"
    GetAllSchedules = cszCorrectRetCode
    Exit Function
GetSchedulesErrHandle:
    szErrStr = Trim(Str(Old2New(err.Number)))
    Debug.Print err.Number & "---" & err.Description
    TxStr = Left(RxStr, 243 - 1) & szErrStr & FormatLen(MidA(err.Description, 1, 80), 80) & "@"
    GetAllSchedules = szErrStr
End Function

'售票
Public Function SellTickets(RxStr As String, TxStr As String) As String
    Dim dyDate As Date
    Dim adyDate_new() As Date
    
    Dim rsRec As ADODB.Recordset
    Dim nConnectNo As Integer
    Dim szBusID  As String
    Dim aszBusID_new()  As String
    
    Dim atSellTKParam() As TSellTicketParam
    
    Dim szEndStationID As String
    Dim aszEndStationID_new() As String
    
    Dim szStartStationID As String
    Dim szBusType As String
    Dim nTkNums As Integer
    Dim szTicketType As String
    Dim szStartTime As String * 5
    Dim aszResultSeatID() As String * 2 '售票售出后,传回的座位信息,数组形式
    Dim szResultSeatID As String '售票售出后,传回的座位信息
    Dim aszSourceSeatID() As String  '传输过来的座位信息,数组形式
    Dim szSourceSeatID As String '传输过来的座位信息
    Dim szTkPrice As String * 6
    Dim szExamNo As String * 5
    Dim szTkNo As String
    Dim lTkNol As Long '票号转换为数字
    'Dim atBuyTicketInfo() As String
    Dim atBuyTicketInfo() As TBuyTicketInfo
    Dim i As Integer
'    Dim tmpTkO As Object
    'Dim atSellTicketResultInfo As Object
'    Dim rsBusRec As ADODB.Recordset
    Dim atSellTicketResultInfo() As TSellTicketResult
    Dim anInsurrance() As Integer
    Dim anChildren() As Integer
    Dim szErrStr As String * 4
    
    Dim szDBOperatorID As String
    
    Dim sgTicketPrice As Single
    Dim szPix As String
    On Error GoTo SellTicketsErrHandle
    
    szDBOperatorID = GetOperatorID(RxStr)
    szBusID = GetPackageBusID(RxStr)   ' Mid(RxStr, 50 + 1, 4)
'    i = CInt(szBusID)
'    szBusID = Trim(Str(i))
    szStartStationID = Trim(MidA(RxStr, 35, 1)) '嘉兴需求。GetStartStationID(RxStr) 'Trim(Mid(RxStr, 32 + 1, 9))
    nConnectNo = szStartStationID
    szEndStationID = GetDestStationID(RxStr) ' Trim(Mid(RxStr, 41 + 1, 9))
    szTicketType = GetTicketType(RxStr)  ' Mid(RxStr, 31 + 1, 1)
    nTkNums = GetTicketNum(RxStr)  ' CInt(Mid(RxStr, 56 + 1, 2))
    dyDate = PackageToDate(GetBusOffDate(RxStr))    ' CDate(Mid(RxStr, 64 + 1, 4) & "/" &     Mid(RxStr, 68 + 1, 2) & "/" & Mid(RxStr, 70 + 1, 2))
    szSourceSeatID = GetSeatID(RxStr)
    aszSourceSeatID = Split(szSourceSeatID, ",")
    szTkNo = GetTicketID(RxStr)
    szPix = Left(szTkNo, 1)
    If IsNumeric(szPix) = False Then
        szTkNo = Mid(szTkNo, 2, Len(szTkNo) - 1)
    End If
    
    lTkNol = CLng(szTkNo)
    szStartTime = GetBusOffTime(RxStr)
    sgTicketPrice = GetTicketPrice(RxStr) / 100

    '未连上,则出错
    If m_oSellTicketService(szStartStationID) Is Nothing Then
        TxStr = Left(RxStr, cnPosRetCode - 1) & ERR_CONNECT & FormatLen("车站未接入", cnLenReserved) & "@"
        SellTickets = ERR_CONNECT
        Exit Function
    End If
    
'    Set rsRec = New ADODB.Recordset
'    rsRec.CursorType = adOpenKeyset
'    rsRec.LockType = adLockOptimistic
'    rsRec.Open "SELECT * FROM Tickets", LocalCn, , , adCmdText
'    If (rsRec.RecordCount <> 0) And (Not rsRec.EOF) Then
'        rsRec.MoveLast
'    End If
    
    Dim SellStationID(1 To 1) As String
    ReDim atBuyTicketInfo(1 To nTkNums)
    ReDim aszBusID_new(1 To 1)
    ReDim adyDate_new(1 To nTkNums)
    ReDim anChildren(1 To nTkNums)
    ReDim aszEndStationID_new(1 To 1)
    ReDim atSellTKParam(1 To 1)
    ReDim anInsurrance(1 To 1)
    
    Dim aszSellStationID() As String
    
'    aszSellStationID = GetAllSellStationID()
    
   SellStationID(1) = Trim(MidA(RxStr, 36, 8)) '嘉兴需求。
    'SellStationID(1) = g_aszAllStartStation(nConnectNo + 1, cnIniPosSellStationID)
    

    aszBusID_new(1) = szBusID
    adyDate_new(1) = dyDate
    aszEndStationID_new(1) = szEndStationID
    
    anChildren(1) = 0
    
    
    ReDim aszResultSeatID(1 To nTkNums)
    
    
    
    '-----------lyq 2lines modified
    Dim asgTicketPrices() As Single
    ReDim asgTicketPrices(1 To nTkNums)
    
    For i = 1 To nTkNums
    
        atBuyTicketInfo(i).nTicketType = Val(szTicketType)
        atBuyTicketInfo(i).szReserved = ""
        If i > ArrayLength(aszSourceSeatID) Then
            '如果传过来的座位号清单的数目小于当前的数字的长数,则其他座位号为自动获取
            atBuyTicketInfo(i).szSeatNo = ""
        Else
            atBuyTicketInfo(i).szSeatNo = aszSourceSeatID(i - 1)
        End If
        If IsNumeric(szPix) = False Then
            atBuyTicketInfo(i).szTicketNo = Trim(Str(lTkNol + i - 1))
            atBuyTicketInfo(i).szTicketNo = String(8 - Len(atBuyTicketInfo(i).szTicketNo), "0") & atBuyTicketInfo(i).szTicketNo
            atBuyTicketInfo(i).szTicketNo = szPix & atBuyTicketInfo(i).szTicketNo
        Else
            atBuyTicketInfo(i).szTicketNo = Trim(Str(lTkNol + i - 1))
            atBuyTicketInfo(i).szTicketNo = String(9 - Len(atBuyTicketInfo(i).szTicketNo), "0") & atBuyTicketInfo(i).szTicketNo  '此处票号长度固定为8位
        End If
        
        
        atBuyTicketInfo(i).szSeatTypeID = "01" '座位类型固定为01
        asgTicketPrices(i) = sgTicketPrice
        'szTkNo = Trim(Str(lTkNol + i - 1))
        'atBuyTicketInfo(i, 1) = String(7 - Len(szTkNo), &H30) & szTkNo
        'atBuyTicketInfo(i, 2) = ""
        'atBuyTicketInfo(i, 3) = szTicketType
    Next i
    atSellTKParam(1).BuyTicketInfo = atBuyTicketInfo
    atSellTKParam(1).pasgSellTicketPrice = asgTicketPrices
    anInsurrance(1) = 0
    odb.BeginTrans
    
    atSellTicketResultInfo = m_oSellTicketService(nConnectNo).SellTicket(adyDate_new, aszBusID_new, SellStationID, aszEndStationID_new, atSellTKParam, anInsurrance)
    
    szResultSeatID = ""
    
'    Set rsBusRec = New ADODB.Recordset
'    rsBusRec.CursorType = adOpenKeyset
'    rsBusRec.LockType = adLockOptimistic
'    rsBusRec.Open "select * from schedules where ScheduleID ='" & szBusID & "' and number = " & nConnectNo, LocalCn, , , adCmdText
'    On Error GoTo 0
'    On Error Resume Next
'    rsBusRec.MoveFirst
'    If rsBusRec.BOF And rsBusRec.EOF Then
'        TxStr = Left(RxStr, cnPosRetCode - 1) & ERR_NOSTARTID & Right(RxStr, cnLenReserved) & "@"
'        SellTickets = ERR_NOSTARTID
'        Exit Function
'    End If
'    rsBusRec.MoveFirst
    With atSellTicketResultInfo(1)
        For i = 1 To nTkNums
            aszResultSeatID(i) = .aszSeatNo(i)
            szResultSeatID = szResultSeatID & aszResultSeatID(i)
'            rsRec.AddNew
'            rsRec!id = atBuyTicketInfo(i).szTicketNo
'            If atBuyTicketInfo(i).nTicketType = "0" Then
'            rsRec!TicketType = atBuyTicketInfo(i).nTicketType
'            Else
'                rsRec!TicketType = "0"
'            End If
            'rsRec!TicketType = atBuyTicketInfo(i).nTicketType
'            rsRec!scheduleid = szBusID
'            rsRec!StartTime = Format(rsBusRec.Fields("starttime"), "hh:mm")
'            rsRec!StartID = aszStationID(nConnectNo)
'            rsRec!DestID = szEndStationID
'            rsRec!SeatID = .aszSeatNo(i)
'            rsRec!StartName = Trim(aszStationName(nConnectNo))
'            rsRec!DestName = .szDesStationName
'            rsRec!SellDate = Now
'            rsRec!StartNo = nConnectNo
'            rsRec!StartTicket = atBuyTicketInfo(1).szTicketNo

'            rsRec!price = .asgTicketPrice(i)
            'rsRec!price = .asgTicketPrice
'            rsRec!OperatorID = Mid(RxStr, 14, 5)
'            rsRec.MoveNext
            
            

            '得到票的信息
            m_oServiceTicket(nConnectNo).Identify atBuyTicketInfo(i).szTicketNo
            
            Dim szDBBusID As String
            Dim szDBEndStationID As String
            Dim szDBStartStationID As String
            Dim szDBEndStationName As String
            Dim szDBStartStationName As String
            Dim dyDBBusDate As Date
            Dim dyDBStartTime As Date
            Dim szDBSeatID As String
            Dim dbDBTkPrice As Double
            Dim szDBTicketStatus As String
            Dim nDBTicketType As Integer
            
            szDBBusID = m_oServiceTicket(nConnectNo).REBusID
            szDBEndStationID = m_oServiceTicket(nConnectNo).ToStation
            szDBStartStationID = m_oServiceTicket(nConnectNo).SellStationID
            szDBEndStationName = m_oServiceTicket(nConnectNo).ToStationName
            szDBStartStationName = m_oServiceTicket(nConnectNo).SellStationName
            dyDBBusDate = m_oServiceTicket(nConnectNo).REBusDate
            dyDBStartTime = m_oServiceTicket(nConnectNo).dtBusStartUpTime
            
            szDBSeatID = m_oServiceTicket(nConnectNo).seatNo
            dbDBTkPrice = m_oServiceTicket(nConnectNo).TicketPrice
            szDBTicketStatus = "1" '把票的状态固定写为1
            nDBTicketType = m_oServiceTicket(nConnectNo).TicketType
            
            
            '将售票的明细信息写入数据库中
            
            Dim szSql As String
            szSql = " INSERT tickets ( " _
                & " id , " _
                & " scheduleid , " _
                & " starttime , " _
                & " startid , " _
                & " destid , " _
                & " seatid , " _
                & " startname , " _
                & " destname , " _
                & " selldate , " _
                & " startno , " _
                & " price , " _
                & " operatorid , " _
                & " tickettype , " _
                & " status ) "
                
            szSql = szSql & " VALUES ( " _
                & TransFieldValueToString(atBuyTicketInfo(i).szTicketNo) & " , " _
                & TransFieldValueToString(szDBBusID) & " , " _
                & TransFieldValueToString(ToDBDate(dyDBBusDate) & " " & ToDBTime(dyDBStartTime)) & " , " _
                & TransFieldValueToString(szStartStationID) & " , " _
                & TransFieldValueToString(szEndStationID) & " , " _
                & TransFieldValueToString(szDBSeatID) & " , " _
                & TransFieldValueToString(szDBStartStationName) & " , " _
                & TransFieldValueToString(szDBEndStationName) & " , " _
                & TransFieldValueToString(ToDBDateTime(Now)) & " , " _
                & TransFieldValueToString(nConnectNo) & " , " _
                & TransFieldValueToString(dbDBTkPrice) & " , " _
                & TransFieldValueToString(szDBOperatorID) & " , " _
                & TransFieldValueToString(nDBTicketType) & " , " _
                & TransFieldValueToString(szDBTicketStatus) & " )   "
                
            odb.Execute (szSql)
            
        Next i
    End With
    odb.CommitTrans
    

    szTkPrice = atSellTicketResultInfo(1).asgTicketPrice(1)
    szTkPrice = Trim(Str(CCur(szTkPrice) * 100))
    

    
    TxStr = Left(RxStr, cnPosTicketPrice - 1) & FormatLen(szTkPrice, cnLenTicketPrice) & MidA(RxStr, cnPosBusOffDate, cnLenBusOffDate) & FormatLen(Format(szStartTime, "hhmm"), cnLenBusOffTime) & _
    FormatLen(szResultSeatID, cnLenSeatID) & cszCorrectRetCode
'    TxStr = TxStr & szExamNo & szBusType
    TxStr = TxStr & FormatLen("成功", cnLenReserved) & "@"
'    TxStr = TxStr & Right(RxStr, FIXPKGLEN - Len(TxStr))
    SellTickets = cszCorrectRetCode
    Exit Function
SellTicketsErrHandle:
    odb.RollbackTrans '错误回滚
    szErrStr = Trim(Str(Old2New(err.Number)))
    Debug.Print err.Number & "---" & err.Description
    TxStr = Left(RxStr, cnPosRetCode - 1) & szErrStr & FormatLen(MidA(err.Description, 1, cnLenReserved), cnLenReserved) & "@"
    SellTickets = Trim(Str(Old2New(err.Number)))
End Function

'废票
Public Function CancelTicket(RxStr As String, TxStr As String) As String
'    Dim szTicketID As String
'    Dim aszTicketID() As String
'    Dim nSumTicket As Integer
'
'    Dim rsRec As ADODB.Recordset
'    Dim szPrice As String * 6
'    Dim szTicketType As String * 1
'    Dim szErrStr As String * 4
'    Dim i As Integer
'    szTicketID = GetTicketID(RxStr)
'    ReDim aszTicketID(1 To 1)
'    aszTicketID(1) = szTicketID
'
'    On Error GoTo CancelTicketErrHandle
'
'
'    m_oSellTicketService(GetStartStationID(RxStr)).CancelTicket aszTicketID
'    '返回码设置为正确
'    TxStr = Left(RxStr, cnPosRetCode - 1) & cszCorrectRetCode & GetReserved(RxStr) & "@"
'    CancelTicket = cszCorrectRetCode
'    Exit Function
'CancelTicketErrHandle:
'    szErrStr = Trim(Str(Old2New(err.Number)))
'    '预留处返回为错误描述
'
'    TxStr = Left(RxStr, cnPosRetCode - 1) & szErrStr & FormatLen(err.Description, cnLenReserved) & "@"
'
'    CancelTicket = szErrStr
    Dim szTicketID As String
    Dim aszTicketID() As String
    Dim nSumTicket As Integer
    
    Dim rsRec As ADODB.Recordset
    Dim szPrice As String * 6
    Dim szTicketType As String * 1
    Dim szErrStr As String * 4
    Dim i As Integer
    szTicketID = GetTicketID(RxStr)
    ReDim aszTicketID(1 To 1)
    aszTicketID(1) = szTicketID
    
    On Error GoTo CancelTicketErrHandle
    
            Dim nConnectNo As Integer
            Dim szSql As String
            nConnectNo = Trim(MidA(RxStr, 35, 1)) 'GetStartStationID(RxStr)
            m_oServiceTicket(nConnectNo).Identify szTicketID
            
            Dim szDBBusID As String
            Dim szDBEndStationID As String
            Dim szDBStartStationID As String
            Dim szDBEndStationName As String
            Dim szDBStartStationName As String
            Dim dyDBBusDate As Date
            Dim dyDBStartTime As Date
            Dim szDBSeatID As String
            Dim dbDBTkPrice As Double
            Dim szDBTicketStatus As String
            Dim nDBTicketType As Integer
            
            szDBBusID = m_oServiceTicket(nConnectNo).REBusID
            szDBEndStationID = m_oServiceTicket(nConnectNo).ToStation
            szDBStartStationID = m_oServiceTicket(nConnectNo).SellStationID
            szDBEndStationName = m_oServiceTicket(nConnectNo).ToStationName
            szDBStartStationName = m_oServiceTicket(nConnectNo).SellStationName
            dyDBBusDate = m_oServiceTicket(nConnectNo).REBusDate
            dyDBStartTime = m_oServiceTicket(nConnectNo).dtBusStartUpTime
            szDBSeatID = m_oServiceTicket(nConnectNo).seatNo
            dbDBTkPrice = m_oServiceTicket(nConnectNo).TicketPrice
            szDBTicketStatus = "2" '把票的状态固定写为2
            nDBTicketType = m_oServiceTicket(nConnectNo).TicketType
            
            Dim szTradeID As String
            szTradeID = GetTradeID(RxStr)

    
    
    
    m_oSellTicketService(nConnectNo).CancelTicket aszTicketID
    '返回码设置为正确
    TxStr = Left(RxStr, cnPosRetCode - 1) & cszCorrectRetCode & Mid(RxStr, 164, 31)
    CancelTicket = cszCorrectRetCode
            If m_oServiceTicket(nConnectNo).TicketStatus = "1" And szTradeID <> AUTOCANCELTKID Then
                '废票的明细信息写入数据库中
                szSql = " INSERT tickets ( " _
                    & " id , " _
                    & " scheduleid , " _
                    & " starttime , " _
                    & " startid , " _
                    & " destid , " _
                    & " seatid , " _
                    & " startname , " _
                    & " destname , " _
                    & " selldate , " _
                    & " startno , " _
                    & " price , " _
                    & " operatorid , " _
                    & " tickettype , " _
                    & " status ) "
                    
                szSql = szSql & " VALUES ( " _
                    & TransFieldValueToString(szTicketID) & " , " _
                    & TransFieldValueToString(szDBBusID) & " , " _
                    & TransFieldValueToString(ToDBDate(dyDBBusDate) & " " & ToDBTime(dyDBStartTime)) & " , " _
                    & TransFieldValueToString(GetStartStationID(RxStr)) & " , " _
                    & TransFieldValueToString(GetDestStationID(RxStr)) & " , " _
                    & TransFieldValueToString(szDBSeatID) & " , " _
                    & TransFieldValueToString(szDBStartStationName) & " , " _
                    & TransFieldValueToString(szDBEndStationName) & " , " _
                    & TransFieldValueToString(ToDBDateTime(Now)) & " , " _
                    & TransFieldValueToString(nConnectNo) & " , " _
                    & TransFieldValueToString(dbDBTkPrice) & " , " _
                    & TransFieldValueToString(GetOperatorID(RxStr)) & " , " _
                    & TransFieldValueToString(nDBTicketType) & " , " _
                    & TransFieldValueToString(szDBTicketStatus) & " )   "
    
                odb.Execute (szSql)
            ElseIf m_oServiceTicket(nConnectNo).TicketStatus = "1" And szTradeID = AUTOCANCELTKID Then
                '废票的明细信息写入数据库中
                szSql = " INSERT tickets ( " _
                    & " id , " _
                    & " scheduleid , " _
                    & " starttime , " _
                    & " startid , " _
                    & " destid , " _
                    & " seatid , " _
                    & " startname , " _
                    & " destname , " _
                    & " selldate , " _
                    & " startno , " _
                    & " price , " _
                    & " operatorid , " _
                    & " tickettype , " _
                    & " status ) "
                    
                szSql = szSql & " VALUES ( " _
                    & TransFieldValueToString(szTicketID) & " , " _
                    & TransFieldValueToString(szDBBusID) & " , " _
                    & TransFieldValueToString(ToDBDate(dyDBBusDate) & " " & ToDBTime(dyDBStartTime)) & " , " _
                    & TransFieldValueToString(GetStartStationID(RxStr)) & " , " _
                    & TransFieldValueToString(GetDestStationID(RxStr)) & " , " _
                    & TransFieldValueToString(szDBSeatID) & " , " _
                    & TransFieldValueToString(szDBStartStationName) & " , " _
                    & TransFieldValueToString(szDBEndStationName) & " , " _
                    & TransFieldValueToString(ToDBDateTime(Now)) & " , " _
                    & TransFieldValueToString(nConnectNo) & " , " _
                    & TransFieldValueToString(dbDBTkPrice) & " , " _
                    & TransFieldValueToString(GetOperatorID(RxStr)) & " , " _
                    & TransFieldValueToString(nDBTicketType) & " , " _
                    & TransFieldValueToString(3) & " )   " '自动废票
    
                odb.Execute (szSql)
            End If
    Exit Function
CancelTicketErrHandle:
    szErrStr = Trim(Str(Old2New(err.Number)))
    '预留处返回为错误描述
    
'    TxStr = Left(RxStr, cnPosRetCode - 1) & szErrStr & FormatLen(err.Description, cnLenReserved) & "@"
    TxStr = Left(RxStr, cnPosRetCode - 1) & szErrStr & FormatLen(MidA(err.Description, 1, cnLenReserved), cnLenReserved) & "@"
    CancelTicket = szErrStr
End Function

'老程序前台得到座位
Public Function GetSeats(RxStr As String, TxStr As String) As String
    Dim dyDate As Date
    Dim nConnectNo As Integer
    Dim szBusID As String
    Dim szEndStationID As String
    Dim szStartStationID As String
    Dim szSeats As String * 2
    Dim rsRec As ADODB.Recordset
    Dim i As Integer
    Dim szErrStr As String * 4
    On Error GoTo GetSeatsErrHandle
    szBusID = GetPackageBusID(RxStr)   ' Mid(RxStr, 50 + 1, 4)
    szStartStationID = GetStartStationID(RxStr)
    szEndStationID = GetDestStationID(RxStr)
    
    dyDate = PackageToDate(GetBusOffDate(RxStr))    ' CDate(Mid(RxStr, 64 + 1, 4) & "/" & Mid(RxStr, 68 + 1, 2) & "/" & Mid(RxStr, 70 + 1, 2))
    nConnectNo = szStartStationID
    
    If m_oSellTicketService(nConnectNo) Is Nothing Then
        TxStr = Left(RxStr, cnPosRetCode - 1) & ERR_CONNECT & FormatLen("车站未接入", cnLenReserved) & "@"
        GetSeats = ERR_CONNECT
        Exit Function
    End If
    Set rsRec = m_oSellTicketService(nConnectNo).GetSeatRs(dyDate, szBusID)
    
    TxStr = ConvertSeatRSToString(rsRec)
    
    GetSeats = cszCorrectRetCode
    Exit Function
GetSeatsErrHandle:
    szErrStr = Trim(Str(Old2New(err.Number)))
    TxStr = Left(RxStr, 155 + 1) & szErrStr & Right(RxStr, 31)
    GetSeats = szErrStr
End Function

'得到本地数据库的通讯记录 流水明细
Public Function DailyRec(IP As String, TxStr As String) As String
    Dim szStr As String
    Dim TradeID As String
    Dim rsRec As ADODB.Recordset
    TradeID = GetTradeID(TxStr)
    Set rsRec = New ADODB.Recordset
    rsRec.CursorType = adOpenKeyset
    rsRec.LockType = adLockOptimistic
    rsRec.Open "select * from DailyBook", LocalCn, , , adCmdText
    If rsRec.RecordCount > 0 Then
        rsRec.MoveLast
    End If
    szStr = ERR_DAILYREC
    On Error GoTo DailyRecErrHandle
    If TradeID <> INTERNETSELL Or TradeID <> INTERNETCANCEL Or TradeID <> AUTOCANCELINTTK Then
        With rsRec
            .AddNew
            !opDate = Now
            !MachineIP = IP
            '暂写起点站
            !bankId = FormatLen(GetOperatorBankID(TxStr), cnLenOperatorBankID) ' GetOperatorBankID(TxStr)   'Mid(TxStr, 19, 5)
            !OperatorId = GetOperatorID(TxStr) 'Mid(TxStr, 14, 5)
            !TradeID = GetTradeID(TxStr) 'Mid(TxStr, 10, 4)
            If GetRetCode(TxStr) = cszCorrectRetCode Then
                !TradeOk = TRADE_OK
            Else
                !TradeOk = GetRetCode(TxStr) ' Mid(TxStr, 156 + 1, 4)
            End If
            If !TradeID = BUYTICKETSID Or !TradeID = CANCELTICKETSID Or !TradeID = AUTOCANCELTKID _
                    Or !TradeID = GETTKINFO Then
                !TicketID = GetTicketID(TxStr) ' Mid(TxStr, 24, 7 + 1)
                !Amount = 1  ' CInt(Mid(TxStr, 56 + 1, 2))
                
                !SumMoney = PackageToMoney(GetTicketPrice(TxStr) * GetTicketNum(TxStr))   'CCur(Mid(TxStr, 58 + 1, 6)) * !Amount / 100
            Else
                !TicketID = ""
    '            !Amount = 0
    '            !SumMoney = 0
            
            End If
            .Update
        End With
    Else
'        With rsRec
'            .AddNew
'            !opDate = Now
'            !MachineIP = IP
'            '暂写起点站
'            !bankId = FormatLen(GetOperatorBankID(TxStr), cnLenOperatorBankID) ' GetOperatorBankID(TxStr)   'Mid(TxStr, 19, 5)
'            !OperatorID = GetOperatorID(TxStr) 'Mid(TxStr, 14, 5)
'            !TradeID = GetTradeID(TxStr) 'Mid(TxStr, 10, 4)
'            If MidA(TxStr, 243, 4) = cszCorrectRetCode Then
'                !TradeOk = TRADE_OK
'            Else
'                !TradeOk = MidA(TxStr, 243, 4) ' Mid(TxStr, 156 + 1, 4)
'            End If
'            If !TradeID = INTERNETSELL Then
'                !seatNo = Trim(MidA(TxStr, 103, 80))
'                !Amount = CInt(Mid(TxStr, 83, 2))
'                !SumMoney = PackageToMoney(MidA(TxStr, 85, 6) * CInt(Mid(TxStr, 83, 2))) 'CCur(Mid(TxStr, 58 + 1, 6)) * !Amount / 100
'            ElseIf !TradeID = INTERNETCANCEL Or !TradeID = AUTOCANCELINTTK Then
'                !Amount = CInt(Mid(TxStr, 83, 2))
'                !SumMoney = PackageToMoney(MidA(TxStr, 85, 6) * CInt(Mid(TxStr, 83, 2)))
'            Else
'                !ticketID = ""
'    '            !Amount = 0
'    '            !SumMoney = 0
'
'            End If
'            .Update
'        End With
    
    End If
    DailyRec = cszCorrectRetCode
    Exit Function
DailyRecErrHandle:
    DailyRec = szStr
End Function

'清空本地数据库的通讯记录 流水明细
Public Sub DailyClear()
    Dim rsRec As ADODB.Recordset
    Set rsRec = New ADODB.Recordset
    rsRec.CursorType = adOpenDynamic
    rsRec.LockType = adLockOptimistic
    rsRec.Open "delete from DailyBook", LocalCn, , , adCmdText
    '  If (Not rsRec.BOF) And rsRec.RecordCount <> 0 Then
    '    rsRec.MoveFirst
    '  End If
    '  With rsRec
    '    Do While Not .EOF
    '    .Delete
    '    .MoveNext
    '    Loop
    '  End With
    '  rsRec.Close
    rsRec.Open "delete from Tickets", LocalCn, , , adCmdText
    '  If (Not rsRec.BOF) And rsRec.RecordCount > 0 Then
    '   rsRec.MoveFirst
    '  End If
    '  With rsRec
    '    Do While Not .EOF
    '    .Delete
    '    .MoveNext
    '    Loop
    '  End With
End Sub

Public Function GetCnStatus(Index As Integer)
    If m_oSellTicketService(Index) Is Nothing Then
        GetCnStatus = SERVER_RELEASE
    Else
        GetCnStatus = SERVER_CONNECTED
    End If
End Function

Public Function GetAccountList(RxStr As String, TxStr As String) As String
    Dim rsDailyRec As ADODB.Recordset
    Dim rsTkRec As ADODB.Recordset
    Dim szStr As String
    Dim szSelStr As String
    Dim dyDate As Date
    Dim dyDate_1 As Date
    Dim szBankID As String
    Dim szPrice As String * 6
    Dim szAmountStr As String * 2
    Dim szSumStr As String * 7
    Dim szBusID As String * 4
    Dim szErrStr As String * 4
    On Error GoTo ErrProc
    szBankID = Mid(RxStr, 19, 5)
    dyDate = CDate(Mid(RxStr, 64 + 1, 4) & "/" & Mid(RxStr, 68 + 1, 2) & "/" & Mid(RxStr, 70 + 1, 2))
    dyDate_1 = DateAdd("d", 1, dyDate)
    szSelStr = "select opDate,operatorID,TradeID,TicketID,Amount,SumMoney from DailyBook " & _
    "where convert(char(8),opdate,112)='" & Format(dyDate, "yyyymmdd") & _
    "' and  bankId='" & szBankID & "' and TradeOK='OK' and (TradeID='8001' or TradeID='8011' or tradeid='9999')"
    Set rsDailyRec = New ADODB.Recordset
    Set rsTkRec = New ADODB.Recordset
    rsDailyRec.CursorType = adOpenKeyset
    rsDailyRec.LockType = adLockOptimistic
    rsTkRec.CursorType = adOpenKeyset
    rsTkRec.LockType = adLockOptimistic
    rsDailyRec.Open szSelStr, LocalCn, , , adCmdText
    TxStr = Left(RxStr, 1) & "00000191" & Right(RxStr, 181 + 1)
    With rsDailyRec
        Do While Not .EOF
            rsTkRec.Open "Select * from tickets where ID='" & !TicketID & _
            "'", LocalCn, , , adCmdText
            szAmountStr = Trim(Str(!Amount))
            szPrice = Trim(Str(rsTkRec!Price * 100))
            szSumStr = Trim(Str(!SumMoney * 100))
            szBusID = String(4 - Len(rsTkRec!scheduleid), &H30) & rsTkRec!scheduleid
            TxStr = TxStr & !TradeID & Format(!opDate, "hhmm") & !OperatorId & _
            !TicketID & rsTkRec!startname & rsTkRec!destname & szBusID & _
            Trim(rsTkRec!TicketType) & szPrice & szAmountStr & szSumStr & "|"
            .MoveNext
            rsTkRec.Close
        Loop
    End With
    TxStr = TxStr & "@"
    Exit Function
ErrProc:
    szErrStr = ERR_ERRDATE
    TxStr = Left(RxStr, 155 + 1) & szErrStr & Right(RxStr, 31)
    GetAccountList = szErrStr
End Function
Private Sub Class_Terminate()
    Dim i As Integer
    For i = 0 To 3
        If Not m_oSellTicketService(i) Is Nothing Then
            Set m_oSellTicketService(i) = Nothing
        End If
'        If Not moUser(i) Is Nothing Then
'            Set moUser(i) = Nothing
'        End If
'        If Not moGate(i) Is Nothing Then
'            Set moGate(i) = Nothing
'        End If
    Next i
    LocalCn.Close
End Sub

 

Public Function GetStartName(i As Integer) As String
    GetStartName = Trim(aszStationName(i))

End Function

'得到车次票价
Public Function GetTkPrice(RxStr As String, TxStr As String) As String
    Dim rsSchsRec As ADODB.Recordset
    Dim szBusID As String
    Dim szEndStationID As String
    Dim szStartStationID As String
    Dim szTkPrice As String * 6
    Dim i As Integer
    
    ' Schedules InputID <----> bzm
    On Error GoTo GetTkPriceHandle
    Set rsSchsRec = New ADODB.Recordset
    szBusID = Mid(RxStr, 50 + 1, 4)
    i = CInt(szBusID)
    szBusID = Trim(Str(i))
    szStartStationID = Trim(Mid(RxStr, 32 + 1, 9))
    szEndStationID = Trim(Mid(RxStr, 41 + 1, 9))
    For i = 0 To 3
        '      If szStartStationID = UCase(Trim(aszStationInputID(i))) Then
        If szStartStationID = UCase(Trim(aszStationID(i))) Then
            
            szStartStationID = aszStationID(i)
            Exit For
        End If
    Next i
    
    rsSchsRec.CursorType = adOpenKeyset
    rsSchsRec.LockType = adLockOptimistic
    rsSchsRec.Open "select FPrice,HPrice from Schedules where StartID='" & szStartStationID & _
    "' and StandardID='" & szEndStationID & "' and ScheduleID='" & szBusID & "'", LocalCn, , , adCmdText
    If rsSchsRec.RecordCount = 0 Then
        TxStr = Left(RxStr, 155 + 1) & ERR_NOSCHEDULEID & Right(RxStr, 31)
        GetTkPrice = ERR_NOSCHEDULEID
        Exit Function
    End If
    szTkPrice = Trim(Str(rsSchsRec!FPrice * 100))
    TxStr = Left(RxStr, 57 + 1) & szTkPrice & Right(RxStr, 127)
    GetTkPrice = cszCorrectRetCode
GetTkPriceHandle:
 
End Function


'得到车票的相关信息
Public Function GetTicketInfo(RxStr As String, TxStr As String) As String

    Dim szBusID As String
    Dim szEndStationID As String
    Dim szStartStationID As String
    Dim szEndStationName As String
    Dim szStartStationName As String
    Dim dyBusDate As Date
    Dim dyBusStartTime As Date
    Dim szUserID As String
    Dim szSeatID As String
    Dim sgTicketPrice As Double
    Dim nTicketStatus As Integer
    Dim nTicketType As Integer
    
    Dim szTicketID As String '票号
    
    Dim nConnectNo As Integer
    Dim szErrStr As String
    
    On Error GoTo ErrorHandle
    
    szTicketID = GetTicketID(RxStr)
    nConnectNo = Val(GetStartStationID(RxStr))
    
    If m_oServiceTicket(nConnectNo) Is Nothing Then
        '未接入,则出错
        TxStr = Left(RxStr, cnPosRetCode - 1) & ERR_CONNECT & FormatLen("车站未接入", cnLenReserved) & "@"
        GetTicketInfo = ERR_CONNECT
        Exit Function
    End If
    
    '得到票的信息
    m_oServiceTicket(nConnectNo).Identify szTicketID
    
    szBusID = m_oServiceTicket(nConnectNo).REBusID
    szEndStationID = m_oServiceTicket(nConnectNo).ToStation
    szStartStationID = m_oServiceTicket(nConnectNo).SellStationID
    szEndStationName = m_oServiceTicket(nConnectNo).ToStationName
    szStartStationName = m_oServiceTicket(nConnectNo).SellStationName
    dyBusDate = m_oServiceTicket(nConnectNo).REBusDate
    dyBusStartTime = m_oServiceTicket(nConnectNo).dtBusStartUpTime
    szUserID = m_oServiceTicket(nConnectNo).Operator
    szSeatID = m_oServiceTicket(nConnectNo).seatNo
    sgTicketPrice = m_oServiceTicket(nConnectNo).TicketPrice
    nTicketStatus = m_oServiceTicket(nConnectNo).TicketStatus
    nTicketType = m_oServiceTicket(nConnectNo).TicketType
    
    TxStr = GetQueryTicketInfoRequestStr(szTicketID, nTicketType, szStartStationID, szEndStationID, szBusID, dyBusDate, szStartStationName _
        , szEndStationName, dyBusStartTime, szUserID, szSeatID, sgTicketPrice, nTicketStatus)
    GetTicketInfo = cszCorrectRetCode
    
    Exit Function
ErrorHandle:
 
    szErrStr = Trim(Str(Old2New(err.Number)))
    '预留处返回为错误描述
'    TxStr = Left(RxStr, cnPosRetCode - 1) & szErrStr & FormatLen(err.Description, cnLenReserved) & "@"
    TxStr = Left(RxStr, cnPosRetCode - 1) & szErrStr & FormatLen(MidA(err.Description, 1, cnLenReserved), cnLenReserved) & "@"
    GetTicketInfo = szErrStr
End Function

Public Function Old2New(OldRet As Long) As Long
Dim lRet As Long
    If OldRet > 9999 Then
        lRet = OldRet - 10000
    ElseIf OldRet < 999 Then
        lRet = OldRet + 1000
    Else
        lRet = OldRet
    End If
    
    
    
    Old2New = lRet
    If (Old2New <> 0) Then
    
      'TxStr = Left(TxStr, 160) & mid(Err.Description,1,30) &
    End If
End Function




Private Function GetConnectionStr(Optional ByVal pszWhich As String) As String
'    Dim oReg As New CFreeReg
    Dim szDriverType As String
    Dim szIntegrated As String '是否集成帐户
    '1先将默认值读出

        
    Dim szDBSetSection As String
    Select Case m_szDatabaseType
        Case "SQLOLEDB.1"   'SQL Server
            GetConnectionStr = "Provider=" & m_szDatabaseType _
            & ";Persist Security Info=False" _
            & IIf(szIntegrated <> "", ";Integrated Security=" & szIntegrated, ";User ID=" & m_szUser & ";Password=" & m_szPassword) _
            & ";Initial Catalog=" & m_szDatabase _
            & ";Data Source=" & m_szServer _
            & IIf(m_szTimeout = "", "", ";Timeout=" & Val(m_szTimeout))
        Case Else
            GetConnectionStr = ""
    End Select
End Function

'网上售票xxd
Public Function InterNetSellTickets(RxStr As String, TxStr As String) As String
    Dim dyDate As Date
    Dim adyDate_new() As Date
    
    Dim rsRec As ADODB.Recordset
    Dim nConnectNo As Integer
    Dim szBusID  As String
    Dim aszBusID_new()  As String
    
    Dim atSellTKParam() As TSellTicketParam
    
    Dim szEndStationID As String
    Dim aszEndStationID_new() As String
    
    Dim szStartStationID As String
    Dim szBusType As String
    Dim nTkNums As Integer
    Dim szTicketType As String
    Dim szStartTime As String * 5
    Dim aszResultSeatID() As String * 2 '售票售出后,传回的座位信息,数组形式
    Dim szResultSeatID As String '售票售出后,传回的座位信息
    Dim aszSourceSeatID() As String  '传输过来的座位信息,数组形式
    Dim szSourceSeatID As String '传输过来的座位信息
    Dim szTkPrice As String * 6
    Dim szExamNo As String * 5
    Dim szTkNo As String
    Dim lTkNol As Long '票号转换为数字
    'Dim atBuyTicketInfo() As String
    Dim atBuyTicketInfo() As TBuyTicketInfo
    Dim i As Integer
'    Dim tmpTkO As Object
    'Dim atSellTicketResultInfo As Object
'    Dim rsBusRec As ADODB.Recordset
    Dim atSellTicketResultInfo() As TSellTicketResult
    Dim anInsurrance() As Integer
    Dim szErrStr As String * 4
    Dim szBankCardID As String * 19
    Dim szDBOperatorID As String
    Dim szBusyDate As String
    Dim szBusyTime As String * 6
    Dim szCenterScrollID As String * 12
    Dim szGetTicketID As String * 13
    Dim szValiDateID As String * 5
    Dim szCardID As String * 18
    Dim szPayCount As String * 12
    Dim szBankID As String * 5
    Dim j As Integer
    Dim seatNo As String
    On Error GoTo SellTicketsErrHandle
    
    szBankID = Mid(RxStr, 19, 5)
    szDBOperatorID = GetOperatorID(RxStr)
    szBusID = Mid(RxStr, 76, 5) ' Mid(RxStr, 50 + 1, 4)
    szBankCardID = Mid(RxStr, 24, 19)
    szBusyDate = PackageToDate(Mid(RxStr, 44, 8))
    szBusyTime = Mid(RxStr, 52, 6)
    szStartStationID = Mid(RxStr, 58, 9) 'Trim(Mid(RxStr, 32 + 1, 9))
    nConnectNo = szStartStationID
    szEndStationID = Mid(RxStr, 67, 9) ' Trim(Mid(RxStr, 41 + 1, 9))
    szTicketType = Mid(RxStr, 43, 1) ' Mid(RxStr, 31 + 1, 1)
    nTkNums = Mid(RxStr, 83, 2) ' CInt(Mid(RxStr, 56 + 1, 2))
    dyDate = PackageToDate(Mid(RxStr, 91, 8))  ' CDate(Mid(RxStr, 64 + 1, 4) & "/" &     Mid(RxStr, 68 + 1, 2) & "/" & Mid(RxStr, 70 + 1, 2))
'    szSourceSeatID = GetSeatID(RxStr)
'    aszSourceSeatID = Split(szSourceSeatID, ",")
'    szTkNo = GetTicketID(RxStr)
'    lTkNol = CLng(szTkNo)
    szStartTime = Mid(RxStr, 99, 4)
'    szBusType = Mid(RxStr, 81, 2)
    szCenterScrollID = Mid(RxStr, 183, 12)
    szGetTicketID = Mid(RxStr, 195, 13)
    szValiDateID = Mid(RxStr, 208, 5)
    szCardID = Mid(RxStr, 213, 18)
    szPayCount = Mid(RxStr, 231, 12)
    
    '未连上,则出错
    If m_oSellTicketService(szStartStationID) Is Nothing Then
        TxStr = Left(RxStr, 243 - 1) & ERR_CONNECT & Mid(RxStr, 247, 80) & "@"
        InterNetSellTickets = ERR_CONNECT
        Exit Function
    End If
    
'    Set rsRec = New ADODB.Recordset
'    rsRec.CursorType = adOpenKeyset
'    rsRec.LockType = adLockOptimistic
'    rsRec.Open "SELECT * FROM Tickets", LocalCn, , , adCmdText
'    If (rsRec.RecordCount <> 0) And (Not rsRec.EOF) Then
'        rsRec.MoveLast
'    End If
    
    Dim SellStationID(1 To 1) As String
    ReDim atBuyTicketInfo(1 To nTkNums)
    ReDim aszBusID_new(1 To 1)
    ReDim adyDate_new(1 To nTkNums)
    ReDim aszEndStationID_new(1 To 1)
    ReDim atSellTKParam(1 To 1)
    ReDim anInsurrance(1 To 1)
    
    Dim aszSellStationID() As String
    
'    aszSellStationID = GetAllSellStationID()

    SellStationID(1) = g_aszAllStartStation(nConnectNo + 1, cnIniPosSellStationID)
'    SellStationID(1) = Trim(Mid(RxStr, 59, 2)) '嘉兴需求。

    aszBusID_new(1) = szBusID
    adyDate_new(1) = dyDate
    aszEndStationID_new(1) = szEndStationID
    
    
    ReDim aszResultSeatID(1 To nTkNums)
    
    
    
    '-----------lyq 2lines modified
    Dim asgTicketPrices() As Single
    Dim aszBankParam() As String
    ReDim asgTicketPrices(1 To nTkNums)
    ReDim aszBankParam(1 To nTkNums, 1 To 8)
    For i = 1 To nTkNums
    
        atBuyTicketInfo(i).nTicketType = Val(szTicketType)
        atBuyTicketInfo(i).szReserved = ""
        If i > ArrayLength(aszSourceSeatID) Then
            '如果传过来的座位号清单的数目小于当前的数字的长数,则其他座位号为自动获取
            atBuyTicketInfo(i).szSeatNo = ""
        Else
            atBuyTicketInfo(i).szSeatNo = aszSourceSeatID(i - 1)
        End If
        atBuyTicketInfo(i).szTicketNo = Trim(Str(lTkNol + i - 1))
        atBuyTicketInfo(i).szTicketNo = String(8 - Len(atBuyTicketInfo(i).szTicketNo), "0") & atBuyTicketInfo(i).szTicketNo  '此处票号长度固定为8位
        atBuyTicketInfo(i).szSeatTypeID = "01" '座位类型固定为01
        asgTicketPrices(i) = 0
        'szTkNo = Trim(Str(lTkNol + i - 1))
        'atBuyTicketInfo(i, 1) = String(7 - Len(szTkNo), &H30) & szTkNo
        'atBuyTicketInfo(i, 2) = ""
        'atBuyTicketInfo(i, 3) = szTicketType
        aszBankParam(i, 1) = szBusyDate & " " & Left(szBusyTime, 2) & ":" & Mid(szBusyTime, 3, 2) & ":" & Right(szBusyTime, 2) '交易日期
        aszBankParam(i, 2) = szBankCardID '银行卡号
        aszBankParam(i, 3) = szCenterScrollID '中心流水号
        aszBankParam(i, 4) = szGetTicketID '取票号
        aszBankParam(i, 5) = szValiDateID ''验证码
        aszBankParam(i, 6) = szCardID '卡号
        aszBankParam(i, 7) = Val(szPayCount) / 100 / CInt(nTkNums) '交易金额
        aszBankParam(i, 8) = szBankID '银行号
    Next i
    
    atSellTKParam(1).BuyTicketInfo = atBuyTicketInfo
    atSellTKParam(1).pasgSellTicketPrice = asgTicketPrices
    anInsurrance(1) = 0
    odb.BeginTrans
    
    atSellTicketResultInfo = m_oInternetSellTicketService(nConnectNo).InterNetSellLocalTicket(adyDate_new, aszBusID_new, SellStationID, aszEndStationID_new, atSellTKParam, anInsurrance, aszBankParam)

            Dim szSql As String
    With atSellTicketResultInfo(1)
        For i = 1 To nTkNums
            szSql = " INSERT InterTickets ( " _
                & " Bus_id , " _
                & " StartTime , " _
                & " StartID , " _
                & " DestID , " _
                & " SellDate , " _
                & " SeatNo , " _
                & " StartNo , " _
                & " price , " _
                & " OperatorID , " _
                & " TicketType , " _
                & " GetTicketID , " _
                & " ValiDateID , " _
                & " status ) "
                
            szSql = szSql & " VALUES ( " _
                & TransFieldValueToString(szBusID) & " , " _
                & TransFieldValueToString(ToDBDate(dyDate) & " " & Format(szStartTime, cszTimeStr)) & " , " _
                & TransFieldValueToString(SellStationID(1)) & " , " _
                & TransFieldValueToString(szEndStationID) & " , " _
                & TransFieldValueToString(ToDBDateTime(Now)) & " , " _
                & TransFieldValueToString(.aszSeatNo(i)) & " , " _
                & TransFieldValueToString(szStartStationID) & " , " _
                & TransFieldValueToString(.asgTicketPrice(i)) & " , " _
                & TransFieldValueToString(szDBOperatorID) & " , " _
                & TransFieldValueToString(szTicketType) & " , " _
                & TransFieldValueToString(Trim(szGetTicketID)) & " , " _
                & TransFieldValueToString(Trim(szValiDateID)) & " , " _
                & TransFieldValueToString(1) & " )   " '票的状态固定为１
                
            odb.Execute (szSql)
        Next i
    End With
    odb.CommitTrans
    For j = 1 To ArrayLength(atSellTicketResultInfo(1).aszSeatNo)
        seatNo = seatNo & atSellTicketResultInfo(1).aszSeatNo(j)
    Next j


'
    TxStr = Left(RxStr, 102) & FormatLen(seatNo, 80)
    TxStr = TxStr & MidA(RxStr, 183, 60) & cszCorrectRetCode
    TxStr = TxStr & MidA(RxStr, 247, 81)
    InterNetSellTickets = cszCorrectRetCode
    Exit Function
SellTicketsErrHandle:
    odb.RollbackTrans '错误回滚
    szErrStr = Trim(Str(Old2New(err.Number)))
    Debug.Print err.Number & "---" & err.Description
'    TxStr = Left(RxStr, 243 - 1) & szErrStr & FormatLen(err.Description, cnLenReserved) & "@"
    TxStr = Left(RxStr, 243 - 1) & szErrStr & FormatLen(MidA(err.Description, 1, 80), 80) & "@"
    InterNetSellTickets = Trim(Str(Old2New(err.Number)))
End Function



Public Function UnInterNetTicket(RxStr As String, TxStr As String) As String '网上退票
    Dim szTicketID As String
    Dim nSumTicket As Integer
    
    Dim rsRec As ADODB.Recordset
    Dim szGetTicketID As String
    Dim szValiDateID As String
    Dim i As Integer
    Dim szResult As String
    Dim szErrStr As String * 4
    Dim szStartStationID As String
    Dim nConnectNo As Integer
    Dim rs As Recordset
    Dim szTicketType As String
    Dim szTradeID As String
    Dim szSql As String
    Dim nGetNum As Integer
    Dim aszSeatID() As String
    Dim szSeatNo As String
    szGetTicketID = Mid(RxStr, 195, 13)
    szValiDateID = Mid(RxStr, 208, 5)
     szStartStationID = Mid(RxStr, 58, 9) 'Trim(Mid(RxStr, 32 + 1, 9))
    nConnectNo = szStartStationID
    szTicketType = Mid(RxStr, 43, 1)
    szTradeID = GetTradeID(RxStr)
    nGetNum = CInt(MidA(RxStr, 83, 2))
    szSeatNo = MidA(RxStr, 103, 80)
    aszSeatID = StringToTeamRX(szSeatNo, 2)
    On Error GoTo CancelTicketErrHandle
    
       Set rs = m_oInternetSellTicketService(nConnectNo).InterNetValiDate(szGetTicketID, szValiDateID)
       If rs.RecordCount <= 0 Then
            TxStr = Left(RxStr, 243 - 1) & ERR_ValiDate & FormatLen(MidA("无效订票号", 1, 80), 80) & "@"
            UnInterNetTicket = ERR_ValiDate
            Exit Function
       End If
       odb.BeginTrans
       szResult = m_oInternetSellTicketService(nConnectNo).UnInetNetTicket(szGetTicketID, szValiDateID, nGetNum, aszSeatID)
       If szResult = "0000" Then
            
       End If
            If rs.RecordCount > 0 And szTradeID <> AUTOCANCELINTTK Then
                '废票的明细信息写入数据库中
                For i = 1 To nGetNum
                szSql = " UPDATE InterTickets SET status=2 where  bus_id='" & Trim(rs!bus_id) & "' AND starttime='" & CDate(rs!bus_date) _
                    & "' AND seatno='" & Trim(aszSeatID(i)) & "' AND startno='" & nConnectNo _
                    & "'  AND Getticketid='" & Trim(szGetTicketID) & "' AND ValiDateid='" & Trim(szValiDateID) & "'" '退票
                rs.MoveNext
                odb.Execute (szSql)
                Next i
            ElseIf rs.RecordCount > 0 And szTradeID = AUTOCANCELINTTK Then
                For i = 1 To nGetNum
                szSql = " UPDATE InterTickets SET status=3 where  bus_id='" & Trim(rs!bus_id) & "' AND starttime='" & CDate(rs!bus_date) _
                    & "' AND seatno='" & Trim(aszSeatID(i)) & "' AND startno='" & nConnectNo _
                    & "'  AND Getticketid='" & Trim(szGetTicketID) & "' AND ValiDateid='" & Trim(szValiDateID) & "'" ''自动网上退票
                rs.MoveNext
                odb.Execute (szSql)

                Next i
            End If
        odb.CommitTrans
    TxStr = Left(RxStr, 242) & cszCorrectRetCode
    TxStr = TxStr & MidA(RxStr, 247, 81)
    UnInterNetTicket = cszCorrectRetCode
    Exit Function
CancelTicketErrHandle:
    odb.RollbackTrans
    szErrStr = Trim(Str(Old2New(err.Number)))
    '预留处返回为错误描述
    TxStr = Left(RxStr, 243 - 1) & szErrStr & FormatLen(MidA(err.Description, 1, 80), 80) & "@"
    UnInterNetTicket = szErrStr
End Function

Public Function GetInterNetTicket(RxStr As String, TxStr As String) As String '取网上售票
    Dim rs As Recordset
    Dim aszResult() As TSellTicketResult
    Dim szTicketID As String
    Dim nSumTicket As Integer
    
    Dim rsRec As ADODB.Recordset
    Dim szGetTicketID As String
    Dim szValiDateID As String
    Dim i As Integer
    Dim szResult As String
    Dim szErrStr As String * 4
    Dim szStartStationID As String
    Dim szCardID As String
    Dim szTradeID As String
    Dim szSql As String
    Dim szBusOffDate As Date
    Dim nConnectNo As Integer
    Dim szStartTKID As String
    Dim rsTmp As Recordset
    Dim szSeatID As String
    Dim nGetNum As Integer  '取票数量
    Dim szSql2 As String
    szStartStationID = Mid(RxStr, 58, 9)
    szGetTicketID = Mid(RxStr, 195, 13)
    szValiDateID = Mid(RxStr, 208, 5)
    szBusOffDate = PackageToDate(MidA(RxStr, 91, 8))
    nConnectNo = szStartStationID
    szCardID = MidA(RxStr, 213, 18)
    szStartTKID = MidA(RxStr, 247, 8)
    nGetNum = CInt(MidA(RxStr, 83, 2))
On Error GoTo GetInterNetTicketError
    If Trim(szCardID) <> "" And Trim(szBusOffDate) <> "" Then
        Set rs = m_oInternetSellTicketService(nConnectNo).InterNetValiDate(szGetTicketID, szValiDateID, szCardID, szBusOffDate) '取售票记录
    Else
        Set rs = m_oInternetSellTicketService(nConnectNo).InterNetValiDate(szGetTicketID, szValiDateID)
    End If
    If rs.RecordCount <= 0 Then
        TxStr = Left(RxStr, 243 - 1) & ERR_ValiDate & FormatLen(MidA("无效订票号", 1, 80), 80) & "@"
        GetInterNetTicket = ERR_ValiDate
        Exit Function
    Else
        odb.BeginTrans
        aszResult = m_oInternetSellTicketService(nConnectNo).SellInetNetTicket(rs, szStartTKID, nGetNum) '售网上订票
        szSeatID = ""

'        Set rsTmp = m_oInternetSellTicketService(nConnectNo).GetBusExRs(Trim(rs!bus_date), Trim(rs!bus_id), Trim(rs!sell_station_id), Trim(rs!des_station_id))
        
        For i = 1 To nGetNum
'            m_oInternetSellTicketService(nConnectNo).FinPrint CDate(Trim(rs!bus_date)), Trim(rs!bus_id), aszResult(1).aszSeatNo(i), Trim(rs!bank_card_id)
            szSeatID = szSeatID & aszResult(1).aszSeatNo(i)
            szSql2 = "UPDATE InterTickets SET ticketid='" & szStartTKID + i - 1 & "',status=4 where bus_id='" & Trim(rs!bus_id) & "' AND starttime='" & CDate(rs!bus_date) _
                    & "' AND seatno='" & Trim(aszResult(1).aszSeatNo(i)) & "' AND startno='" & nConnectNo & "' AND status=1  AND Getticketid='" & Trim(szGetTicketID) & "' AND ValiDateid='" & Trim(szValiDateID) & "'"
            odb.Execute szSql2
        Next i
        
        odb.CommitTrans
        TxStr = Left(RxStr, 57) & FormatLen(Trim(rs!sell_station_id), 9) & FormatLen(Trim(rs!des_station_id), 9) & FormatLen(Trim(rs!bus_id), 5) & MidA(RxStr, 81, 2) & FormatLen(rs.RecordCount, 2) & FormatLen(aszResult(1).asgTicketPrice(1) * 100, 6)
        TxStr = TxStr & FormatLen(Format(Trim(rs!bus_date), "YYYYMMDD"), 8) & FormatLen(" ", 4) & FormatLen(szSeatID, 80)
        TxStr = TxStr & MidA(RxStr, 183, 60) & cszCorrectRetCode & MidA(RxStr, 247, 81)
        GetInterNetTicket = cszCorrectRetCode
    End If
    Exit Function
GetInterNetTicketError:
    odb.RollbackTrans
    szErrStr = Trim(Str(Old2New(err.Number)))
    '预留处返回为错误描述
    TxStr = Left(RxStr, 243 - 1) & szErrStr & FormatLen(MidA(err.Description, 1, 80), 80) & "@"
    GetInterNetTicket = szErrStr
End Function
Public Function GetAccountNetTicket(RxStr As String, rsOut As Recordset, TxStr As String) As String '查询网上售票情况
    Dim rs As ADODB.Recordset
    Dim szOperatorID As String
    Dim dtStrDate As Date
    Dim dtEndDate As Date
    Dim szSql As String
    Dim szErrStr As String
    szOperatorID = GetOperatorID(RxStr)
    dtStrDate = PackageToDate(MidA(RxStr, 19, 8))
    dtEndDate = PackageToDate(MidA(RxStr, 27, 8))
    Set rs = New ADODB.Recordset
On Error GoTo GetAccountNetTicketError
    szSql = "select GetTicketID,ValiDateID,Bus_id,StartTime,StartID,DestID,SellDate,SeatNo,StartNo,price,OperatorID,TicketType,case Status when '1' then '售出' when '2' then '已退' when '3' then '自动退票' when '4' then '已取票'  when '6' then '废票'  when '5' then '重打票' end as Status,ticketid from InterTickets  where " _
        & "  selldate >= " & TransFieldValueToString(dtStrDate) _
        & " AND selldate < " & TransFieldValueToString(DateAdd("d", 1, dtEndDate)) _
        & " order by selldate"
    rs.CursorType = adOpenKeyset
    rs.LockType = adLockOptimistic
    rs.Open szSql, LocalCn, , , adCmdText
    If rs.RecordCount <= 0 Then
        TxStr = Left(RxStr, 34) & ERR_ERRFIND & FormatLen(MidA("查询条件有误", 1, 80), 80) & "@"
        GetAccountNetTicket = ERR_ERRFIND
        Exit Function
    Else

        Set rsOut = rs
        TxStr = Left(RxStr, 34) & cszCorrectRetCode
        TxStr = TxStr & MidA(RxStr, 35, 81)
        GetAccountNetTicket = cszCorrectRetCode
    End If
    Exit Function
GetAccountNetTicketError:
    szErrStr = Trim(Str(Old2New(err.Number)))
    '预留处返回为错误描述
    TxStr = Left(RxStr, 34) & szErrStr & FormatLen(MidA(err.Description, 1, 80), 80) & "@"
    GetAccountNetTicket = szErrStr
End Function
Public Function GetAccountTicket(RxStr As String, rsOut As Recordset, TxStr As String) As String '查询售票情况
    Dim rs As ADODB.Recordset
    Dim szOperatorID As String
    Dim dtStrDate As Date
    Dim dtEndDate As Date
    Dim szSql As String
    Dim szErrStr As String
    szOperatorID = GetOperatorID(RxStr)
    dtStrDate = PackageToDate(MidA(RxStr, 19, 8))
    dtEndDate = PackageToDate(MidA(RxStr, 27, 8))
    Set rs = New ADODB.Recordset
On Error GoTo GetAccountNetTicketError
   szSql = "select ID,scheduleid,StartTime,StartID,DestID,SeatID,StartName,DestName,SellDate,StartNo,price,t.OperatorID,TicketType,case t.Status when '2' then '作废' when '3' then '自动作废' else '售出' end as Status from tickets t where " _
        & "  t.selldate >= " & TransFieldValueToString(dtStrDate) _
        & " AND t.selldate < " & TransFieldValueToString(DateAdd("d", 1, dtEndDate)) _
        & " order by selldate"
    rs.CursorType = adOpenKeyset
    rs.LockType = adLockOptimistic
    rs.Open szSql, LocalCn, , , adCmdText
    
    If rs.RecordCount <= 0 Then
        TxStr = Left(RxStr, 34) & ERR_ERRFIND & FormatLen(MidA("查询条件有误", 1, 80), 80) & "@"
        GetAccountTicket = ERR_ERRFIND
        Exit Function
    Else

        Set rsOut = rs
        TxStr = Left(RxStr, 34) & cszCorrectRetCode
        TxStr = TxStr & MidA(RxStr, 35, 81)
        GetAccountTicket = cszCorrectRetCode
    End If
    Exit Function
GetAccountNetTicketError:
    szErrStr = Trim(Str(Old2New(err.Number)))
    '预留处返回为错误描述
    TxStr = Left(RxStr, 34) & szErrStr & FormatLen(MidA(err.Description, 1, 80), 80) & "@"
    GetAccountTicket = szErrStr
End Function










'重打网上售票接口xxd
Public Function CancelPreTK(RxStr As String, TxStr As String) As String
    Dim dyDate As Date
    Dim adyDate_new() As Date
    
    Dim rsRec As ADODB.Recordset
    Dim nConnectNo As Integer
    Dim szBusID  As String
    Dim aszBusID_new()  As String
    
    Dim atSellTKParam() As TSellTicketParam
    
    Dim szEndStationID As String
    Dim aszEndStationID_new() As String
    
    Dim szStartStationID As String
    Dim szBusType As String
    Dim nTkNums As Integer
    Dim szTicketType As String
    Dim szStartTime As String * 5
    Dim aszResultSeatID() As String * 2 '售票售出后,传回的座位信息,数组形式
    Dim szResultSeatID As String '售票售出后,传回的座位信息
    Dim aszSourceSeatID() As String  '传输过来的座位信息,数组形式
    Dim szSourceSeatID As String '传输过来的座位信息
    Dim szTkPrice As String * 6
    Dim szExamNo As String * 5
    Dim szTkNo As String
    Dim lTkNol As Long '票号转换为数字
    Dim aszTicketID() As String
    'Dim atBuyTicketInfo() As String
    Dim atBuyTicketInfo() As TBuyTicketInfo
    Dim i As Integer
'    Dim tmpTkO As Object
    'Dim atSellTicketResultInfo As Object
'    Dim rsBusRec As ADODB.Recordset
    Dim atSellTicketResultInfo() As TSellTicketResult
    Dim anInsurrance() As Integer
    Dim szErrStr As String * 4
    Dim szBankCardID As String * 19
    Dim szDBOperatorID As String
    Dim szBusyDate As String
    Dim szBusyTime As String * 6
    Dim szCenterScrollID As String * 12
    Dim szGetTicketID As String * 13 '新的取票号
    Dim szValiDateID As String * 5 '新的验证码
    Dim szCardID As String * 18
    Dim szPayCount As String * 12
    Dim szBankID As String * 5
    Dim szOldGetTicketID As String * 13
    Dim szOldValiDateID As String * 5
    Dim j As Integer
    Dim seatNo As String
    Dim szStartTKID As String '旧票号
    Dim szTKID As String '新的票号
    Dim szSql2 As String
    Dim rs As Recordset
    Dim szTemp As String
    On Error GoTo SellTicketsErrHandle
    
    szBankID = Mid(RxStr, 19, 5)
    szDBOperatorID = GetOperatorID(RxStr)
    szBusID = Mid(RxStr, 76, 5)
    szBankCardID = Mid(RxStr, 24, 19)
    szBusyDate = PackageToDate(Mid(RxStr, 44, 8))
    szBusyTime = Mid(RxStr, 52, 6)
    szStartStationID = Mid(RxStr, 58, 9)
    nConnectNo = szStartStationID
    szEndStationID = Mid(RxStr, 67, 9)
    szTicketType = Mid(RxStr, 43, 1)
    nTkNums = Mid(RxStr, 83, 2)
    dyDate = PackageToDate(Mid(RxStr, 91, 8))
    szStartTime = Mid(RxStr, 99, 4)
    szCenterScrollID = Mid(RxStr, 183, 12)
    szGetTicketID = Mid(RxStr, 195, 13)
    szValiDateID = Mid(RxStr, 208, 5)
    szCardID = Mid(RxStr, 213, 18)
    szPayCount = Mid(RxStr, 231, 12)
    szStartTKID = MidA(RxStr, 247, 8)
    szOldGetTicketID = MidA(RxStr, 255, 13)
    szOldValiDateID = MidA(RxStr, 268, 5)
    szTKID = MidA(RxStr, 273, 8)
    '未连上,则出错
    If m_oSellTicketService(szStartStationID) Is Nothing Then
        TxStr = Left(RxStr, 243 - 1) & ERR_CONNECT & Mid(RxStr, 247, 80) & "@"
        CancelPreTK = ERR_CONNECT
        Exit Function
    End If
    
    Dim SellStationID(1 To 1) As String
    ReDim atBuyTicketInfo(1 To nTkNums)
    ReDim aszBusID_new(1 To 1)
    ReDim adyDate_new(1 To nTkNums)
    ReDim aszEndStationID_new(1 To 1)
    ReDim atSellTKParam(1 To 1)
    ReDim anInsurrance(1 To 1)
    ReDim aszTicketID(1 To nTkNums)
    Dim aszSellStationID() As String
    SellStationID(1) = g_aszAllStartStation(nConnectNo + 1, cnIniPosSellStationID)
    Set rs = m_oInternetSellTicketService(nConnectNo).UnRecordset(szOldGetTicketID, szOldValiDateID)
    If rs.RecordCount <= 0 Then
        TxStr = Left(RxStr, 243 - 1) & ERR_ValiDate & FormatLen(MidA("无效订票号", 1, 80), 80) & "@"
        CancelPreTK = ERR_ValiDate
        Exit Function
    End If
    For i = 1 To nTkNums
        szTemp = szStartTKID + i - 1
        aszTicketID(i) = Trim(Left(szStartTKID, Len(szStartTKID) - Len(szTemp)) & szTemp)
        
    Next i
    aszBusID_new(1) = Trim(rs!bus_id)
    adyDate_new(1) = Trim(rs!bus_date)
    aszEndStationID_new(1) = Trim(rs!des_station_id)
    szTicketType = Trim(rs!ticket_type)
    dyDate = Trim(rs!bus_date)
    ReDim aszResultSeatID(1 To nTkNums)

    Dim asgTicketPrices() As Single
    Dim aszBankParam() As String
    ReDim asgTicketPrices(1 To nTkNums)
    ReDim aszBankParam(1 To nTkNums, 1 To 8)
    For i = 1 To nTkNums
    
        atBuyTicketInfo(i).nTicketType = Val(Trim(rs!ticket_type))
        atBuyTicketInfo(i).szReserved = ""
        If i > ArrayLength(aszSourceSeatID) Then
            '如果传过来的座位号清单的数目小于当前的数字的长数,则其他座位号为自动获取
            atBuyTicketInfo(i).szSeatNo = ""
        Else
            atBuyTicketInfo(i).szSeatNo = aszSourceSeatID(i - 1)
        End If
        atBuyTicketInfo(i).szTicketNo = Trim(Str(lTkNol + i - 1))
        atBuyTicketInfo(i).szTicketNo = String(8 - Len(atBuyTicketInfo(i).szTicketNo), "0") & atBuyTicketInfo(i).szTicketNo  '此处票号长度固定为8位
        atBuyTicketInfo(i).szSeatTypeID = "01" '座位类型固定为01
        asgTicketPrices(i) = 0
        aszBankParam(i, 1) = Trim(rs!busy_date) '交易日期
        aszBankParam(i, 2) = Trim(rs!bank_card_id) '银行卡号
        aszBankParam(i, 3) = Trim(rs!center_scroll_id) '中心流水号
        aszBankParam(i, 4) = szGetTicketID '取票号
        aszBankParam(i, 5) = szValiDateID ''验证码
        aszBankParam(i, 6) = Trim(rs!card_id) '卡号
        aszBankParam(i, 7) = Trim(rs!pay_count) '交易金额
        aszBankParam(i, 8) = Trim(rs!bank_id) '银行号
        rs.MoveNext
    Next i
    
    atSellTKParam(1).BuyTicketInfo = atBuyTicketInfo
    atSellTKParam(1).pasgSellTicketPrice = asgTicketPrices
    anInsurrance(1) = 0
    odb.BeginTrans
    
    atSellTicketResultInfo = m_oInternetSellTicketService(nConnectNo).CancelPreNetTK(adyDate_new, aszBusID_new, SellStationID, aszEndStationID_new, atSellTKParam, anInsurrance, aszBankParam, aszTicketID, szOldGetTicketID, szOldValiDateID, szGetTicketID, szValiDateID, szTKID)
    
            Dim szSql As String
    With atSellTicketResultInfo(1)
        For i = 1 To nTkNums
'            m_oInternetSellTicketService(nConnectNo).FinPrint CDate(dyDate), Trim(szBusID), .aszSeatNo(i), Trim(szBankCardID)
            szSql2 = "UPDATE InterTickets SET status=6 where bus_id='" & Trim(aszBusID_new(1)) & "' AND starttime='" & CDate(adyDate_new(1)) _
                    & "' AND seatno='" & Trim(.aszSeatNo(i)) & "' AND startno='" & nConnectNo & "' AND (status=4 or status=5)" _
                    & "  AND Getticketid='" & Trim(szOldGetTicketID) & "' AND ValiDateid='" & Trim(szOldValiDateID) & "'" '废票
            odb.Execute szSql2
            szSql = " INSERT InterTickets ( " _
                & " Bus_id , " _
                & " StartTime , " _
                & " StartID , " _
                & " DestID , " _
                & " SellDate , " _
                & " SeatNo , " _
                & " StartNo , " _
                & " price , " _
                & " OperatorID , " _
                & " TicketType , " _
                & " GetTicketID , " _
                & " ValiDateID , " _
                & " ticketid , " _
                & " status ) "
                
            szSql = szSql & " VALUES ( " _
                & TransFieldValueToString(aszBusID_new(1)) & " , " _
                & TransFieldValueToString(dyDate) & " , " _
                & TransFieldValueToString(SellStationID(1)) & " , " _
                & TransFieldValueToString(aszEndStationID_new(1)) & " , " _
                & TransFieldValueToString(ToDBDateTime(Now)) & " , " _
                & TransFieldValueToString(.aszSeatNo(i)) & " , " _
                & TransFieldValueToString(szStartStationID) & " , " _
                & TransFieldValueToString(.asgTicketPrice(i)) & " , " _
                & TransFieldValueToString(szDBOperatorID) & " , " _
                & TransFieldValueToString(szTicketType) & " , " _
                & TransFieldValueToString(Trim(szGetTicketID)) & " , " _
                & TransFieldValueToString(Trim(szValiDateID)) & " , " _
                & TransFieldValueToString(szTKID + i - 1) & " , " _
                & TransFieldValueToString(5) & " )   " '票的状态固定为5,网上售票重打
                
            odb.Execute (szSql)
        Next i
    End With

    odb.CommitTrans
    For j = 1 To ArrayLength(atSellTicketResultInfo(1).aszSeatNo)
        seatNo = seatNo & atSellTicketResultInfo(1).aszSeatNo(j)
    Next j


'
    TxStr = Left(RxStr, 102) & FormatLen(seatNo, 80)
    TxStr = TxStr & MidA(RxStr, 183, 60) & cszCorrectRetCode
    TxStr = TxStr & MidA(RxStr, 247, 81)
    CancelPreTK = cszCorrectRetCode
    Exit Function
SellTicketsErrHandle:
    odb.RollbackTrans '错误回滚
    szErrStr = Trim(Str(Old2New(err.Number)))
    Debug.Print err.Number & "---" & err.Description
'    TxStr = Left(RxStr, 243 - 1) & szErrStr & FormatLen(err.Description, cnLenReserved) & "@"
    TxStr = Left(RxStr, 243 - 1) & szErrStr & FormatLen(MidA(err.Description, 1, 80), 80) & "@"
    CancelPreTK = Trim(Str(Old2New(err.Number)))
End Function
Public Function GetSeatsEX(RxStr As String, TxStr As String) As String '取剩余座位数
    Dim dyDate As Date
    Dim nConnectNo As Integer
    Dim szBusID As String
    Dim szEndStationID As String
    Dim szStartStationID As String
    Dim szSeats As String * 2
    Dim rsRec As ADODB.Recordset
    Dim i As Integer
    Dim szErrStr As String * 4
    Dim szSeatNum As String
    On Error GoTo GetSeatsErrHandle
    szBusID = Mid(RxStr, 76, 5)
    szStartStationID = Mid(RxStr, 58, 9)
'    szEndStationID = GetDestStationID(RxStr)
    
    dyDate = PackageToDate(Mid(RxStr, 91, 8))     ' CDate(Mid(RxStr, 64 + 1, 4) & "/" & Mid(RxStr, 68 + 1, 2) & "/" & Mid(RxStr, 70 + 1, 2))
    nConnectNo = szStartStationID
    
    If m_oSellTicketService(nConnectNo) Is Nothing Then
        TxStr = Left(RxStr, 243 - 1) & ERR_CONNECT & Mid(RxStr, 247, 80) & "@"
        GetSeatsEX = ERR_CONNECT
        Exit Function
    End If
    Set rsRec = m_oSellTicketService(nConnectNo).GetSeatEx(dyDate, szBusID)
    If rsRec.RecordCount > 0 Then
        szSeatNum = Format(Trim(rsRec!sale_seat_quantity), "00")
    Else
        szSeatNum = -1
    End If
    TxStr = Left(RxStr, 243 - 1) & cszCorrectRetCode & FormatLen(MidA(szSeatNum, 1, 2), 2) & FormatLen(MidA(err.Description, 1, 78), 78) & "@"
    GetSeatsEX = cszCorrectRetCode
    Exit Function
GetSeatsErrHandle:
    szErrStr = Trim(Str(Old2New(err.Number)))
    TxStr = Left(RxStr, 243 - 1) & szErrStr & FormatLen(MidA(err.Description, 1, 80), 80) & "@"
    GetSeatsEX = szErrStr
End Function

Public Function UpdBankCard(RxStr As String, TxStr As String) As String '更新银行卡号
    Dim szTicketID As String
    Dim nSumTicket As Integer
    
    Dim rsRec As ADODB.Recordset
    Dim szGetTicketID As String
    Dim szValiDateID As String
    Dim i As Integer
    Dim szResult As String
    Dim szErrStr As String * 4
    Dim szStartStationID As String
    Dim nConnectNo As Integer
    Dim rs As Recordset
    Dim szTicketType As String
    Dim szTradeID As String
    Dim szSql As String
    Dim szCardID As String
    szGetTicketID = Mid(RxStr, 195, 13)
    szValiDateID = Mid(RxStr, 208, 5)
     szStartStationID = Mid(RxStr, 58, 9)
    nConnectNo = szStartStationID
    szCardID = Mid(RxStr, 24, 19)

    On Error GoTo CancelTicketErrHandle
    
       

       szResult = m_oInternetSellTicketService(nConnectNo).UpdBankCardID(szGetTicketID, szValiDateID, szCardID)
       If szResult = "1111" Then
            TxStr = Left(RxStr, 243 - 1) & szErrStr & FormatLen(MidA("无效的订票号", 1, 80), 80) & "@"
       End If
    TxStr = Left(RxStr, 242) & cszCorrectRetCode
    TxStr = TxStr & MidA(RxStr, 247, 81)
    UpdBankCard = cszCorrectRetCode
    Exit Function
CancelTicketErrHandle:
    odb.RollbackTrans
    szErrStr = Trim(Str(Old2New(err.Number)))
    '预留处返回为错误描述
    TxStr = Left(RxStr, 243 - 1) & szErrStr & FormatLen(MidA(err.Description, 1, 80), 80) & "@"
    UpdBankCard = szErrStr
End Function

Public Function VaildateUser(RxStr As String, TxStr As String) As String
    Dim szUser As String
    Dim szPwd As String
    Dim szSql As String
    Dim rs As Recordset
    Dim szBankID As String
    Dim szErrStr As String
    
    szUser = Trim(GetOperatorID(RxStr))
    szPwd = Trim(MidA(RxStr, 164, 30))
    szErrStr = ""
    Set rs = New ADODB.Recordset
    rs.CursorType = adOpenKeyset
    rs.LockType = adLockOptimistic
    rs.Open "SELECT * FROM user_info where operatorid = '" & szUser & "' and user_password= '" & szPwd & "'", LocalCn, , , adCmdText

    If rs.RecordCount <= 0 Then
        szErrStr = "1111"
        TxStr = Left(RxStr, cnPosRetCode - 1) & szErrStr & FormatLen(MidA(err.Description, 1, cnLenReserved), cnLenReserved) & "@"
        
    Else
        szErrStr = "0000"
        TxStr = Left(RxStr, cnPosOperatorBankID - 1) & FormatLen(FormatDbValue(rs!bank_id), 5) & MidA(RxStr, cnPosTicketID, cnPosRetCode - cnPosTicketID) & szErrStr & FormatLen(MidA(FormatDbValue(rs!Annotation), 1, cnLenReserved), cnLenReserved) & "@"
    
    End If
    rs.Close
    
End Function



Public Function SellDetail(RxStr As String, TxStr As String) As String
    '汇总得到银行网点的代售信息
On Error GoTo here
    Dim rs As Recordset
    Dim szSql As String
    Dim pszUser As String
    Dim pdyStartDate As Date
    Dim pdyEndDate As Date
    Set rs = New ADODB.Recordset
    rs.CursorType = adOpenKeyset
    rs.LockType = adLockOptimistic
    pszUser = GetOperatorID(RxStr)
    pdyStartDate = PackageToDate(Mid(RxStr, 19, 8))
    pdyEndDate = PackageToDate(Mid(RxStr, 27, 8))
'    szSql = " SELECT u.bank_id + '[' + max(u.bank_name) + ']' bank ,  FROM tickets t , user_info u  " _
        & " WHERE t.operatorid = u.operatorid " _
        & " AND t.selldate >= " & TransFieldValueToString(pdyStartDate) _
        & " AND t.selldate < " & TransFieldValueToString(DateAdd("d", 1, pdyEndDate)) _
        & " GROUP BY u.bank_id"
    Dim szWhere As String
    Dim szOperatorID As String
    If Trim(pszUser) <> "" Then
        szOperatorID = ResolveDisplay(pszUser)
        szWhere = " AND t.operatorid='" & szOperatorID & "' "
    End If
        
    szSql = "select ID,scheduleid,StartTime,StartID,DestID,SeatID,StartName,DestName,SellDate,StartNo,price,t.OperatorID,TicketType,case t.Status when '2' then '作废' else '售出' end as Status from tickets t where " _
        & "  t.selldate >= " & TransFieldValueToString(pdyStartDate) _
        & " AND t.selldate < " & TransFieldValueToString(DateAdd("d", 1, pdyEndDate)) _
        & szWhere
    rs.Open szSql, LocalCn, , , adCmdText
    TxStr = Left(RxStr, 18) & "|" & ConvertSellDetailRSToString(rs) & "@"
    If rs.RecordCount = 0 Then GoTo here
    
    SellDetail = ""
'    Set rsTemp = odb.Execute(szSql)
'    Set SellDetail = rsTemp
    Exit Function
here:
    If Trim(Str(err.Number)) = 0 Then
        TxStr = "1233" & FormatLen("连接错误", cnLenReserved) & "@"
    Else
        TxStr = Trim(Str(Old2New(err.Number))) & FormatLen(MidA(err.Description, 1, cnLenReserved), cnLenReserved) & "@"
    End If
End Function



Public Function SellCount(RxStr As String, TxStr As String) As String
    '汇总得到银行网点的代售信息
On Error GoTo here
    Dim rs As Recordset
    Dim szSql As String
    Dim pszUser As String
    Dim pdyStartDate As Date
    Dim pdyEndDate As Date
    Set rs = New ADODB.Recordset
    rs.CursorType = adOpenKeyset
    rs.LockType = adLockOptimistic
    pszUser = GetOperatorID(RxStr)
    pdyStartDate = PackageToDate(Mid(RxStr, 19, 8))
    pdyEndDate = PackageToDate(Mid(RxStr, 27, 8))

    Dim szWhere As String
    Dim szOperatorID As String
    If Trim(pszUser) <> "" Then
        szOperatorID = ResolveDisplay(pszUser)
        szWhere = " AND t.operatorid='" & szOperatorID & "' "
    End If
        
    szSql = "select a.operatorid , isnull(a.ticket_num,0) ticket_num , isnull(a.total_price,0) total_price , isnull(b.cancel_ticket_num,0) cancel_ticket_num , isnull(b.cancel_total_price,0) cancel_total_price , a.max_ticket_id , a.min_ticket_id" _
        & " From " _
        & " ( " _
        & " SELECT u.operatorid  ,  isnull(count(id),0) ticket_num , isnull(sum(price),0) total_price , max(id) max_ticket_id , min(id) min_ticket_id " _
        & " FROM tickets t , user_info u " _
        & " Where t.operatorid = u.operatorid and t.status = 1" _
        & " AND t.selldate >= " & TransFieldValueToString(pdyStartDate) _
        & " AND t.selldate <" & TransFieldValueToString(DateAdd("d", 1, pdyEndDate)) _
        & szWhere _
        & " GROUP BY u.operatorid " _
        & " ) a , ( " _
        & " select u.operatorid  ,  isnull(count(id),0) cancel_ticket_num , isnull(sum(price),0) cancel_total_price " _
        & " from tickets t , user_info u " _
        & " Where t.operatorid = u.operatorid and t.status = 2" _
        & " AND t.selldate >= " & TransFieldValueToString(pdyStartDate) _
        & " AND t.selldate <" & TransFieldValueToString(DateAdd("d", 1, pdyEndDate)) _
        & szWhere _
        & " group by u.operatorid " _
        & " ) b " _
        & " where a.operatorid *= b.operatorid "
    rs.Open szSql, LocalCn, , , adCmdText
    If rs.RecordCount = 0 Then GoTo here
    TxStr = Left(RxStr, 18) & "|" & ConvertSellCountRSToString(rs) & "@"
    
    
    SellCount = ""
'    Set rsTemp = odb.Execute(szSql)
'    Set SellDetail = rsTemp
    Exit Function
here:
    If Trim(Str(err.Number)) = 0 Then
        TxStr = "1233" & FormatLen("连接错误", cnLenReserved) & "@"
    Else
        TxStr = Trim(Str(Old2New(err.Number))) & FormatLen(MidA(err.Description, 1, cnLenReserved), cnLenReserved) & "@"
    End If
End Function



