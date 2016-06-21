VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''**********************************************************
''* Source File Name:Book.cls
''* Project Name:STBook.vbp
''* Engineer:陈峰
''* Data Generated:2003/09/30
''* Last Revision Date:2003/09/30
''* Brief Description:预定类
''* Relational Document:
''**********************************************************

Option Explicit


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
  
End Enum

Public Enum ERightBook
    RIGHT_BOOKSEAT = ERR_Book + cnMidRightBegin + cnMidRightStep * 1   '预定座位 30106
    RIGHT_UNBOOKSEAT = ERR_Book + cnMidRightBegin + cnMidRightStep * 2   '取消预定座位 30112
    RIGHT_QUERYBOOKINFO = ERR_Book + cnMidRightBegin + cnMidRightStep * 3 '查询预定信息 30118
    RIGHT_DELETEBOOKREC = ERR_Book + cnMidRightBegin + cnMidRightStep * 4 '删除预定信息 30124
End Enum

Public Enum EBookStatus
    ST_BOOKED = 0 '已被预定
    ST_BOOKCANCELED = 1 '取消了预定
    ST_BOOKSOLD = 2 '已被预定人买走
    ST_BOOKFREESOLD = 3 '超时被自由购买
End Enum

'所有的类都有的私有类变量
Private m_oActiveUser As ActiveUser
'--------------------------->>
Public Sub Init(poAUser As ActiveUser)
    Set SelfUser = poAUser
End Sub

Public Property Get SelfUser() As ActiveUser
    Set SelfUser = m_oActiveUser
    m_oActiveUser.IAmLive
End Property

Public Property Set SelfUser(vNewValue As ActiveUser)
    Set m_oActiveUser = vNewValue
    m_oActiveUser.IAmLive
End Property


'预定指定车次、日期、到站的指定座位
Public Function Book(ByVal pszBusID As String, ByVal pdtBusDate As Date, pszSellStationID As String, pszStationID As String, paszSeatInfo() As String, ByVal pszBookMan As String, Optional ByVal pszTelephone As String = "", Optional ByVal pszEmail As String = "", Optional ByVal pszAnnotation As String = "") As String
    AssertActiveUserValid m_oActiveUser, ERR_Book
    AssertHaveRight m_oActiveUser, RIGHT_BOOKSEAT
    
    Dim nBookSeatCount As Integer, i As Integer
    Dim oDb As New RTConnection
    Dim szSql As String
    Dim szBookID As String
    Dim oParam As New SystemParam
    Dim rsTemp As Recordset, rsTemp2 As Recordset
    Dim lRowAffect As Long
    Dim szTemp As String
    Dim nStationLimitedCount As Integer
    
    
    Dim szSellStationID As String
    
    
    oParam.Init m_oActiveUser
    Randomize Timer
    szBookID = Int(Rnd() * (10 ^ oParam.BookNumberLen))
    
    oDb.ConnectionString = GetConnectionStr("")
    
    '考虑有配载，预定只有玉环用到，所以预定得都为玉环的，写死了
'    pszSellStationID = "yh"
    
    szSql = "SELECT * FROM work_env_bus_info e,work_env_bus_station_lst s " _
    & " WHERE e.bus_id='" & pszBusID & "' AND " _
    & " e.bus_date='" & ToDBDate(pdtBusDate) & "' AND " _
    & " e.bus_id=s.bus_id AND " _
    & " e.bus_date=s.bus_date AND " _
    & " station_id='" & pszStationID & "'"
    
    If pszSellStationID <> "" Then
        szSql = szSql & "  AND sell_station_id='" & pszSellStationID & "'   "
    End If
    
    Set rsTemp = oDb.Execute(szSql)
    If rsTemp.RecordCount < 1 Then
        ShowError ERR_BusNotExist, "CBook.Book", "BusID=" & pszBusID & ",BusDate=" & ToDBDate(pdtBusDate)
    Else
        If Trim(m_oActiveUser.SellStationID) = "" Then
            szSellStationID = FormatDbValue(rsTemp!sell_station_Id)
        Else
            szSellStationID = Trim(m_oActiveUser.SellStationID)
        End If
        If rsTemp!bus_type = TP_ScrollBus Then ShowError ERR_BusNotRegularBus, "CBook.Book", "BusID=" & pszBusID & ",BusDate=" & ToDBDate(pdtBusDate)
        
        If DateAdd("n", oParam.BookTime, oParam.NowDateTime) >= rsTemp!bus_start_time Then
            ShowError ERR_BookTimeOut, "CBook.Book", "BusID=" & pszBusID & ",BusDate=" & ToDBDate(pdtBusDate)
        End If
        nStationLimitedCount = rsTemp!max_sale_quantity
        If nStationLimitedCount = 0 Then ShowError ERR_BusStationMatchLimitedCount, "CBook.Book"
    End If
    
    nBookSeatCount = ArrayLength(paszSeatInfo)
    oDb.BeginTrans
    On Error GoTo Error_Handle
    
    For i = 1 To nBookSeatCount
        If nStationLimitedCount > 0 Then
            szSql = "UPDATE env_bus_station_lst SET " _
            & " book_count=book_count+1 , sale_ticket_quantity=sale_ticket_quantity-1  WHERE " _
            & " bus_id='" & pszBusID & "' AND " _
            & " bus_date='" & ToDBDate(pdtBusDate) & "' AND " _
            & " sale_ticket_quantity>book_count AND " _
            & " station_id='" & pszStationID & "' AND sell_station_id = '" & szSellStationID & "' AND " _
            & " seat_type_id=" _
            & " (Select seat_type_id from env_bus_seat_lst WHERE bus_id='" & pszBusID & "' " _
            & " AND bus_date='" & ToDBDate(pdtBusDate) & "' AND seat_no='" & paszSeatInfo(i) & " ')"
            oDb.Execute szSql, lRowAffect
            If lRowAffect <> 1 Then ShowError ERR_BusStationMatchLimitedCount, "CBook.Book"
        Else
            '更改预订张数
            szSql = "UPDATE env_bus_station_lst SET " _
            & " book_count=book_count+1 WHERE " _
            & " bus_id='" & pszBusID & "' AND " _
            & " bus_date='" & ToDBDate(pdtBusDate) & "' AND " _
            & " station_id='" & pszStationID & "' AND sell_station_id = '" & szSellStationID & "' AND " _
            & " seat_type_id=" _
            & " (Select seat_type_id from env_bus_seat_lst WHERE bus_id='" & pszBusID & "' " _
            & " AND bus_date='" & ToDBDate(pdtBusDate) & "' AND seat_no='" & paszSeatInfo(i) & " ')"
            oDb.Execute szSql, lRowAffect
        
        End If
        
        '修改环境车次的可售张数
        szSql = "UPDATE Env_bus_info SET sale_seat_quantity=sale_seat_quantity-1 , "
        '这里默认座位类型为"坐位"
        szSql = szSql & "seat_remain=seat_remain - 1 "
        szSql = szSql & " WHERE bus_id='" & pszBusID & "' AND bus_date='" & ToDBDate(pdtBusDate) & "'"
        oDb.Execute szSql
        
        
        
        szSql = "UPDATE env_bus_seat_lst SET " _
        & " status=" & ST_SeatBooked & " WHERE " _
        & "bus_id='" & pszBusID _
        & "' AND bus_date='" & ToDBDate(pdtBusDate) _
        & "' AND seat_no='" & paszSeatInfo(i) _
        & "' AND status=" & ST_SeatCanSell
        oDb.Execute szSql, lRowAffect
        
        If lRowAffect = 0 Then
            szSql = "SELECT * FROM env_bus_seat_lst WHERE " _
            & " bus_id='" & pszBusID _
            & "' AND bus_date='" & ToDBDate(pdtBusDate) _
            & "' AND seat_no='" & paszSeatInfo(i) & "'"
            Set rsTemp = oDb.Execute(szSql)
            If rsTemp.RecordCount = 1 Then
                Select Case rsTemp!Status
                    Case ST_SeatBooked
                    ShowError ERR_BOOKSeatBooked, "CBook.Book", "BusID=" & pszBusID & ",BusDate=" & ToDBDate(pdtBusDate) & ",SeatNo=" & paszSeatInfo(i)
                    Case ST_SeatReserved
                    ShowError ERR_BookSeatReserved, "CBook.Book", "BusID=" & pszBusID & ",BusDate=" & ToDBDate(pdtBusDate) & ",SeatNo=" & paszSeatInfo(i)
                    Case ST_SeatSold
                    ShowError ERR_BookSeatSaled, "CBook.Book", "BusID=" & pszBusID & ",BusDate=" & ToDBDate(pdtBusDate) & ",SeatNo=" & paszSeatInfo(i)
                    Case ST_SeatSlitp
                      ShowError ERR_BookSeatSaled, "CBook.Book", "BusID=" & pszBusID & ",BusDate=" & ToDBDate(pdtBusDate) & ",SeatNo=" & paszSeatInfo(i)
                    Case ST_SeatReplace
                      ShowError ERR_BookSeatSaled, "CBook.Book", "BusID=" & pszBusID & ",BusDate=" & ToDBDate(pdtBusDate) & ",SeatNo=" & paszSeatInfo(i)
                End Select
            Else
                ShowError ERR_BookSeatNotExist, "CBook.Book", "BusID=" & pszBusID & ",BusDate=" & ToDBDate(pdtBusDate) & ",SeatNo=" & paszSeatInfo(i)
            End If
        End If
        
        szSql = "INSERT ticket_book_lst(bus_date," _
        & "bus_id," _
        & "seat_no," _
        & "station_id," _
        & "sell_station_id," _
        & "user_id," _
        & "cancel_user_id," _
        & "operate_time," _
        & "book_man, " _
        & "book_number," _
        & "telephone," _
        & "email," _
        & "status," _
        & "annotation) " _
        & " VALUES('" & ToDBDate(pdtBusDate) _
        & "','" & pszBusID _
        & "','" & paszSeatInfo(i) _
        & "','" & pszStationID _
        & "','" & szSellStationID
        
        szSql = szSql _
        & "','" & m_oActiveUser.UserID _
        & "','" _
        & "','" & ToDBDateTime(oParam.NowDateTime) _
        & "','" & pszBookMan _
        & "','" & szBookID _
        & "','" & pszTelephone _
        & "','" & pszEmail _
        & "'," & ST_BOOKED _
        & ",'" & pszAnnotation & "')"
        oDb.Execute szSql
    Next
    
    oDb.CommitTrans
    Book = szBookID
    szTemp = paszSeatInfo(1)
    For i = 2 To nBookSeatCount
        szTemp = szTemp & "," & paszSeatInfo(i)
    Next
    WriteOperateLog m_oActiveUser, RIGHT_BOOKSEAT, "BusID=" & pszBusID & ",BusDate=" & ToDBDate(pdtBusDate) & ",SeatNo=" & szTemp
    Exit Function
Error_Handle:
    oDb.RollbackTrans
    err.Raise err.Number
End Function

'取消预定指定车次、日期的指定座位
Public Function UnBook(ByVal pszBusID As String, ByVal pdtBusDate As Date, paszSeatInfo() As String) As Long
    AssertActiveUserValid m_oActiveUser, ERR_Book
    AssertHaveRight m_oActiveUser, RIGHT_UNBOOKSEAT
    
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim lRowAffect As Long
    Dim nSeatCount As Integer, i As Integer
    Dim szTemp As String
    Dim rsTemp As Recordset
    Dim nStationLimitedCount As Integer
    Dim szStationID As String
    Dim szSellStationID As String
    
    nSeatCount = ArrayLength(paszSeatInfo)
    
    oDb.ConnectionString = GetConnectionStr("")
    oDb.BeginTrans
    On Error GoTo Error_Handle
    
    For i = 1 To nSeatCount
        szSql = "SELECT * FROM ticket_book_lst WHERE  " _
        & " bus_id='" & pszBusID _
        & "' AND bus_date='" & ToDBDate(pdtBusDate) _
        & "' AND seat_no='" & paszSeatInfo(i) _
        & "' AND status=" & ST_BOOKED

        Set rsTemp = oDb.Execute(szSql)
        If rsTemp.RecordCount <> 1 Then ShowError ERR_SeatBookedCanBeCancel, "CBook.UnBook", "BusID=" & pszBusID & ",BusDate=" & ToDBDate(pdtBusDate) & ",SeatNo=" & paszSeatInfo(i)
        szStationID = Trim(rsTemp!station_id)
        szSellStationID = FormatDbValue(rsTemp!sell_station_Id)
        
        szSql = "SELECT * FROM work_env_bus_info e ,work_env_bus_station_lst s " _
        & " WHERE e.bus_id='" & pszBusID & "' AND " _
        & " e.bus_date='" & ToDBDate(pdtBusDate) & "' AND " _
        & " e.bus_id=s.bus_id AND " _
        & " e.bus_date=s.bus_date AND " _
        & " station_id='" & szStationID & "'" _
        & " AND sell_station_id = '" & szSellStationID & "'"
        
        Set rsTemp = oDb.Execute(szSql)
        nStationLimitedCount = rsTemp!max_sale_quantity
        If nStationLimitedCount > 0 Then
            szSql = "UPDATE env_bus_station_lst SET " _
                & " book_count=book_count-1, sale_ticket_quantity=sale_ticket_quantity+1   WHERE " _
                & " bus_id='" & pszBusID & "' AND " _
                & " bus_date='" & ToDBDate(pdtBusDate) & "' AND " _
                & " station_id='" & szStationID & "' AND sell_station_id = '" & szSellStationID & "'" _
                & " AND seat_type_id =(select seat_type_id from env_bus_seat_lst where " _
                & " bus_id='" & pszBusID & "' and bus_date='" & ToDBDate(pdtBusDate) & "' and seat_no='" & paszSeatInfo(i) & "')"
            oDb.Execute szSql
        Else
            szSql = "UPDATE env_bus_station_lst SET " _
                & " book_count=book_count-1   WHERE " _
                & " bus_id='" & pszBusID & "' AND " _
                & " bus_date='" & ToDBDate(pdtBusDate) & "' AND " _
                & " station_id='" & szStationID & "' AND sell_station_id = '" & szSellStationID & "'" _
                & " AND seat_type_id =(select seat_type_id from env_bus_seat_lst where " _
                & " bus_id='" & pszBusID & "' and bus_date='" & ToDBDate(pdtBusDate) & "' and seat_no='" & paszSeatInfo(i) & "')"
            oDb.Execute szSql
        End If
        
        '修改环境车次的可售张数
        szSql = "UPDATE Env_bus_info SET sale_seat_quantity=sale_seat_quantity+1 , "
        '这里默认座位类型为"坐位"
        szSql = szSql & "seat_remain=seat_remain + 1 "
        szSql = szSql & " WHERE bus_id='" & pszBusID & "' AND bus_date='" & ToDBDate(pdtBusDate) & "'"
        oDb.Execute szSql
        
        szSql = "UPDATE env_bus_seat_lst " _
        & " SET status=" & ST_SeatCanSell & " WHERE " _
        & " bus_id='" & pszBusID _
        & "' AND bus_date='" & ToDBDate(pdtBusDate) _
        & "' AND seat_no='" & paszSeatInfo(i) _
        & "' AND status=" & ST_SeatBooked
        oDb.Execute szSql, lRowAffect
        If lRowAffect <> 1 Then ShowError ERR_SeatNotBooked, "CBook.UnBook", "BusID=" & pszBusID & ",BusDate=" & ToDBDate(pdtBusDate) & ",SeatNo=" & paszSeatInfo(i)
        
        
        
        
        szSql = "UPDATE ticket_book_lst SET " _
        & "status=" & ST_BOOKCANCELED _
        & ",cancel_user_id='" & m_oActiveUser.UserID _
        & "' WHERE " _
        & " bus_id='" & pszBusID _
        & "' AND bus_date='" & ToDBDate(pdtBusDate) _
        & "' AND seat_no='" & paszSeatInfo(i) _
        & "' AND status=" & ST_BOOKED
        oDb.Execute szSql, lRowAffect
        If lRowAffect <> 1 Then ShowError ERR_SeatNotBooked, "CBook.UnBook", "BusID=" & pszBusID & ",BusDate=" & ToDBDate(pdtBusDate) & ",SeatNo=" & paszSeatInfo(i)
    Next
    oDb.CommitTrans
    szTemp = paszSeatInfo(1)
    For i = 2 To nSeatCount
        szTemp = szTemp & "," & paszSeatInfo(i)
    Next
    WriteOperateLog m_oActiveUser, RIGHT_UNBOOKSEAT, "BusID=" & pszBusID & ",BusDate=" & ToDBDate(pdtBusDate) & ",SeatNo=" & szTemp
    Exit Function
Error_Handle:
    oDb.RollbackTrans
    err.Raise err.Number
End Function

'得到指定日期车次的预定座位情况
Public Function GetBookedSeat(ByVal pdtBusDate As Date, paszBusID() As String, Optional szStation As String) As Recordset
    AssertActiveUserValid m_oActiveUser, ERR_Book
    AssertHaveRight m_oActiveUser, RIGHT_QUERYBOOKINFO
    
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim nBusCount As Integer, i As Integer
    Dim oDb As New RTConnection
    oDb.ConnectionString = GetConnectionStr("")
    szSql = "SELECT b.*,u1.user_name , " _
        & " u2.user_name AS cancel_user_name,station_name,bus.bus_start_time " _
        & " FROM ticket_book_lst b ,user_info u1 ,user_info u2 ,station_info si ,work_env_bus_info bus WHERE " _
        & " bus.bus_id=b.bus_id AND " _
        & " bus.bus_date=b.bus_date AND " _
        & " b.station_id=si.station_id AND " _
        & " b.user_id=u1.user_id AND " _
        & " b.cancel_user_id*=u2.user_id AND " _
        & " b.bus_date='" & ToDBDate(pdtBusDate) & "' "
        
    If szStation <> "" Then
        szSql = szSql & " AND ( b.station_id LIKE '%" & szStation & "%' OR dbo.fGetPy(station_name) LIKE '%" & szStation & "%' OR station_name LIKE '%" & szStation & "%' )"
    End If
    
    nBusCount = ArrayLength(paszBusID)
    If nBusCount > 0 Then
        szSql = szSql & " AND b.bus_id IN("
        szSql = szSql & "'" & paszBusID(1) & "'"
        For i = 2 To nBusCount
            szSql = szSql & ",'" & paszBusID(i) & "'"
        Next
        szSql = szSql & ")"
    End If
    Set rsTemp = oDb.Execute(szSql)
    
    Set GetBookedSeat = rsTemp
End Function
                 
'删除预定信息
Public Function DeleteBookRec(paszBookEventID() As Long) As Long
    AssertActiveUserValid m_oActiveUser, ERR_Book
    AssertHaveRight m_oActiveUser, RIGHT_DELETEBOOKREC
    
    Dim szTemp As String
    Dim nEventCount As Integer, i As Integer
    Dim szSql As String
    Dim oDb As New RTConnection
    Dim lRowAffect As Long
    nEventCount = ArrayLength(paszBookEventID)
    If nEventCount > 0 Then
        oDb.ConnectionString = GetConnectionStr("")
        szSql = "DELETE ticket_book_lst WHERE (book_event_id=" & paszBookEventID(1)
        For i = 2 To nEventCount
            szSql = szSql & " OR book_event_id=" & paszBookEventID(i)
        Next
        szSql = szSql & ") AND status<>0"
        oDb.Execute szSql, lRowAffect
        
        DeleteBookRec = lRowAffect
        WriteOperateLog m_oActiveUser, RIGHT_DELETEBOOKREC
    End If
End Function

Public Function GetEnvBusAllotInfo(ByVal pdtBusDate As Date, pszBusID As String) As Recordset
    
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim oDb As New RTConnection
    oDb.ConnectionString = GetConnectionStr("")
    
    szSql = "SELECT a.*,si.sell_station_name ,ci.check_gate_name " _
        & " FROM work_env_bus_allot_lst a ,sell_station_info si ,checkgate_info ci WHERE " _
        & " a.can_sale_quantity <>0 AND a.sell_station_id =si.sell_station_id AND a.check_gate_id =ci.check_gate_id AND " _
        & " a.bus_date='" & ToDBDate(pdtBusDate) & "' AND a.bus_id ='" & pszBusID & "' ORDER BY a.bus_start_time"
    
    Set rsTemp = oDb.Execute(szSql)
    
    Set GetEnvBusAllotInfo = rsTemp
End Function
