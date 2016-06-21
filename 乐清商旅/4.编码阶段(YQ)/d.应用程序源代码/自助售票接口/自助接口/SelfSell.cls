VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SelfSell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private m_oActiveUser As ActiveUser '用户
Private m_oSell As NetSellTicketClient '售票
Private m_oBook As CBook
Private m_oSysParm As SystemParam
Private Const cnRecordSplit As String = "%"
Private Const cnSplit As String = "|"
Private Const cszScrollBus = "滚动"
Private Const cszOK = "OK"
Private Const cszError = "Error"
'************系统参数
Private g_nBookTime As Long '预定释放时间(单位:分钟)
Private g_bIsBookValid As Boolean '是否使用预定系统
Private m_bScroll As Boolean '售滚动班次
'************
Private Function CompareRegisterCode() As Boolean
    Dim oReg As New CFreeReg
    Dim szRegisterCode As String
    On Error GoTo Error_Handle
    oReg.Init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany
    szRegisterCode = oReg.GetSetting(g_cszDBSetSection, g_cszRegisterCode, "")
    If UnEncryptPassword(szRegisterCode) = MakeCode(CStr(DriveSerial)) Then
        CompareRegisterCode = True
    End If
    
    Exit Function
Error_Handle:
    

End Function
Public Function RegisterCode(pszCode As String) As String

    Dim oReg As New CFreeReg
On Error GoTo Error_Handle
    If MakeCode(CStr(DriveSerial)) = pszCode Then
        oReg.Init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany
        oReg.SaveSetting g_cszDBSetSection, g_cszRegisterCode, EncryptPassword(pszCode)
        RegisterCode = cszOK & cnSplit
    Else
        RegisterCode = cszError & cnSplit & "无效的序列号" & cnSplit
    End If
    Exit Function
Error_Handle:
    RegisterCode = cszError & cnSplit & err.Description & cnSplit
End Function
Public Function Login(pszUserID As String, pszUserPwd As String) As String
On Error GoTo Error_Handle
    If CompareRegisterCode Then

    Else
        Login = cszError & cnSplit & "未注册,请将以下串发给浙江方苑公司" & CStr(DriveSerial)
        Exit Function
        
    End If

    Set m_oActiveUser = New ActiveUser
    Set m_oSell = New NetSellTicketClient
    Set m_oSysParm = New SystemParam
    Set m_oBook = New CBook
    
    m_oActiveUser.Login pszUserID, pszUserPwd, "外自助"
    m_oSell.Init m_oActiveUser
    m_oSysParm.Init m_oActiveUser
    m_oBook.Init m_oActiveUser
    
    g_nBookTime = m_oSysParm.BookTime '预定释放时间(单位:分钟)
    g_bIsBookValid = m_oSysParm.IsBookValid
'    m_bScroll = IIf(m_oSysParm.GetParam("SelfSellScroll").szParamValue = 1, True, False)
    Login = cszOK & cnSplit
    Exit Function
Error_Handle:
    
    Login = cszError & cnSplit & err.Description & cnSplit
End Function
'取检票口
Public Function GetCheckGate(ByRef pszReturn As String) As String
On Error GoTo Error_Handle
    Dim szReturn As String
    Dim m_aszCheckGateInfo() As String
    Dim i As Integer
    m_aszCheckGateInfo = m_oSell.GetAllCheckGate()
    For i = 1 To ArrayLength(m_aszCheckGateInfo)
        szReturn = szReturn & m_aszCheckGateInfo(i, 1) & cnRecordSplit
        szReturn = szReturn & m_aszCheckGateInfo(i, 2) & cnRecordSplit
        szReturn = szReturn & cnSplit
    Next
    If ArrayLength(m_aszCheckGateInfo) > 0 Then
        pszReturn = ArrayLength(m_aszCheckGateInfo) & cnSplit & szReturn
    Else
        pszReturn = "0" & cnSplit
    End If
    
    GetCheckGate = cszOK & cnSplit
    Exit Function
Error_Handle:
    GetCheckGate = cszError & cnSplit & err.Description & cnSplit
End Function
'取站点
Public Function GetStation(ByRef pszReturn As String) As String

On Error GoTo Error_Handle
    Dim rsTemp As Recordset
    Dim szReturn As String
    Dim i As Integer
    Set rsTemp = m_oSell.GetAllStationRs()
    For i = 1 To rsTemp.RecordCount
        szReturn = szReturn & FormatDbValue(rsTemp!station_id) & cnRecordSplit
        szReturn = szReturn & FormatDbValue(rsTemp!station_name) & cnRecordSplit
        szReturn = szReturn & FormatDbValue(rsTemp!station_input_code) & cnRecordSplit
        szReturn = szReturn & cnSplit
        rsTemp.MoveNext
    Next
    If rsTemp.RecordCount > 0 Then
        pszReturn = rsTemp.RecordCount & cnSplit & szReturn
    Else
        pszReturn = "0" & cnSplit
    End If
    GetStation = cszOK & cnSplit
    Exit Function
Error_Handle:
    GetStation = cszError & cnSplit & err.Description & cnSplit

End Function
'取车次
Public Function GetBus(pszDate As String, pszStationID As String, ByRef pszReturn As String) As String

On Error GoTo Error_Handle
    Dim rsCountTemp As Recordset
    Dim szReturn As String
    Dim i As Integer
    Dim nCount As Integer
    Set rsCountTemp = m_oSell.GetBusRs(CDate(pszDate), pszStationID)
    For i = 1 To rsCountTemp.RecordCount

        If rsCountTemp!bus_type = TP_ScrollBus Then GoTo NextFor
        
        If rsCountTemp!Status = ST_BusStopped Or rsCountTemp!Status = ST_BusMergeStopped Or rsCountTemp!Status = ST_BusSlitpStop Or rsCountTemp!sale_seat_quantity = 0 Then
    
            GoTo NextFor
    
    
        End If
        
        szReturn = szReturn & FormatDbValue(rsCountTemp!bus_id) & cnRecordSplit
        szReturn = szReturn & FormatDbValue(rsCountTemp!sell_station_id) & cnRecordSplit
        szReturn = szReturn & FormatDbValue(rsCountTemp!sell_station_name) & cnRecordSplit
        szReturn = szReturn & FormatDbValue(rsCountTemp!end_station_id) & cnRecordSplit
        szReturn = szReturn & FormatDbValue(rsCountTemp!end_station_name) & cnRecordSplit
        
        szReturn = szReturn & FormatDbValue(rsCountTemp!route_name) & cnRecordSplit

        szReturn = szReturn & FormatDbValue(rsCountTemp!vehicle_type_name) & cnRecordSplit
        szReturn = szReturn & FormatDbValue(rsCountTemp!full_price) & cnRecordSplit
        szReturn = szReturn & FormatDbValue(rsCountTemp!sell_check_gate_id) & cnRecordSplit
    
        If rsCountTemp!bus_type <> TP_ScrollBus Then
            szReturn = szReturn & Format(rsCountTemp!BusStartTime, "hh:mm") & cnRecordSplit

        Else
            szReturn = szReturn & cszScrollBus & cnRecordSplit
        End If

        
        
        If rsCountTemp!bus_type <> TP_ScrollBus Then
            If g_bIsBookValid And DateAdd("n", -g_nBookTime, Format(rsCountTemp!BusStartTime, "hh:mm")) < Time And ToDBDate(CDate(pszDate)) = ToDBDate(Date) Then
                '如果车次日期为当天,且已过预定时限,则将预定人数加到可售张数上面.
                szReturn = szReturn & rsCountTemp!sale_seat_quantity + rsCountTemp!book_count & cnRecordSplit
               
            Else
                 szReturn = szReturn & rsCountTemp!sale_seat_quantity & cnRecordSplit

            End If
        Else
            szReturn = szReturn & "不限" & cnRecordSplit
        End If
        

        nCount = nCount + 1
        szReturn = szReturn & cnSplit
NextFor:
        rsCountTemp.MoveNext
    Next
    If nCount > 0 Then
        pszReturn = nCount & cnSplit & szReturn
    Else
        pszReturn = "0" & cnSplit
    End If
    GetBus = cszOK & cnSplit
    
    Exit Function
Error_Handle:

    GetBus = cszError & cnSplit & err.Description & cnSplit
End Function
Public Function BookTicket(pszBusID As String, pszDate As String, pszSellStationID As String, pszEndStationID As String, pnNums As Integer, ByRef pszReturn As String) As String
On Error GoTo Error_Handle
        Dim i As Integer
        Dim szReturn As String
        
        Dim dyBusDate() As Date
        Dim szBusID() As String
        Dim szDesStationID() As String
        Dim szDesStationName() As String
        Dim psgDiscount() As Single
        Dim srSellResult() As TSellTicketResult
        Dim szBookNumber() As String
        Dim aSellTicket() As TSellTicketParam
        Dim szSellStationID() As String
        Dim szSellStationName() As String
        Dim anInsurance() As Integer '售票用
        
        ReDim dyBusDate(1 To 1)
        ReDim szBusID(1 To 1)
        ReDim szDesStationID(1 To 1)
        ReDim szDesStationName(1 To 1)
        ReDim psgDiscount(1 To 1)
        ReDim srSellResult(1 To 1)
        ReDim szBookNumber(1 To 1)
        ReDim aSellTicket(1 To 1)
        ReDim szSellStationID(1 To 1)
        ReDim szSellStationName(1 To 1)
        ReDim anInsurance(1 To 1)
        
        ReDim aSellTicket(1).BuyTicketInfo(1 To pnNums)
        ReDim aSellTicket(1).pasgSellTicketPrice(1 To pnNums)
        
        For i = 1 To pnNums
            aSellTicket(1).BuyTicketInfo(i).nTicketType = TP_FullPrice
            aSellTicket(1).BuyTicketInfo(i).szTicketNo = ""
            aSellTicket(1).BuyTicketInfo(i).szSeatNo = ""
            aSellTicket(1).pasgSellTicketPrice(i) = 0
            aSellTicket(1).BuyTicketInfo(i).szReserved = ""
            aSellTicket(1).BuyTicketInfo(i).szSeatTypeID = "01"
            aSellTicket(1).BuyTicketInfo(i).szSeatTypeName = ""
        Next i
        dyBusDate(1) = CDate(pszDate)
        szBusID(1) = pszBusID
        szDesStationID(1) = pszEndStationID
        szDesStationName(1) = ""
        psgDiscount(1) = CSng(0)
        szSellStationID(1) = pszSellStationID
        szSellStationName(1) = ""
        srSellResult = m_oSell.BookTK(dyBusDate, szBusID, szSellStationID, szDesStationID, szDesStationName, aSellTicket, anInsurance)
        
        For i = 1 To ArrayLength(srSellResult(1).aszSeatNo)
            szReturn = szReturn & srSellResult(1).aszSeatNo(i) & cnSplit
            
        Next
        'szReturn = szReturn & cnSplit
        pszReturn = ArrayLength(srSellResult(1).aszSeatNo) & cnSplit & szReturn
    BookTicket = cszOK & cnSplit
    Exit Function
Error_Handle:
    BookTicket = cszError & cnSplit & err.Description & cnSplit
End Function
'取消预留
Public Function UnBookTicket(pszBusID As String, pszDate As String, pszSeats As String) As String
On Error GoTo Error_Handle
        Dim dyBusDate() As Date
        Dim szBusID() As String
        Dim aSellTicket() As TSellTicketParam
        
        ReDim dyBusDate(1 To 1)
        ReDim szBusID(1 To 1)
        ReDim aSellTicket(1 To 1)
        Dim aszSeats() As String
        Dim i As Integer
        aszSeats = Split(pszSeats, cnSplit)

        ReDim aSellTicket(1).BuyTicketInfo(1 To ArrayLength(aszSeats))
        For i = 0 To ArrayLength(aszSeats) - 1
            aSellTicket(1).BuyTicketInfo(i + 1).szSeatNo = aszSeats(i)
        Next
        dyBusDate(1) = CDate(pszDate)
        szBusID(1) = pszBusID
        m_oSell.UnBookEx szBusID, dyBusDate, aSellTicket
    UnBookTicket = cszOK & cnSplit
    Exit Function
Error_Handle:
    UnBookTicket = cszError & cnSplit & err.Description & cnSplit
End Function
Public Function CheckTicketStatus(pszTicketNo As String) As String
On Error GoTo Error_Handle
    Dim aszTicketNo() As String
    Dim bResult As Boolean
    Dim szReturn As String
    Dim i As Integer
    aszTicketNo = Split(pszTicketNo, cnSplit)
        
        For i = 1 To ArrayLength(aszTicketNo)
            bResult = m_oSell.CheckTicketNoStatus(aszTicketNo(i - 1))
            If bResult = True Then
                szReturn = szReturn & aszTicketNo(i - 1) & "已用" & cnSplit
            End If
            
        Next i
        If Trim(szReturn) <> "" Then
            CheckTicketStatus = cszError & cnSplit & szReturn
        Else
            CheckTicketStatus = cszOK & cnSplit
        End If
        
    Exit Function
Error_Handle:
    CheckTicketStatus = cszError & cnSplit & err.Description & cnSplit
End Function
'售出
Public Function SetSell(pszBusID As String, pszDate As String, pszSellStationID As String, pszEndStationID As String, pnNums As Integer, pszSeats As String, pszTicketNo As String, ByRef pszReturn As String) As String
On Error GoTo Error_Handle
        Dim i As Integer
        Dim szReturn As String
        
        Dim dyBusDate() As Date
        Dim szBusID() As String
        Dim szDesStationID() As String
        Dim szDesStationName() As String

        Dim srSellResult() As TSellTicketResult
        Dim szBookNumber() As String
        Dim aSellTicket() As TSellTicketParam
        Dim szSellStationID() As String
        Dim szSellStationName() As String
        Dim anInsurance() As Integer '售票用
        
        ReDim dyBusDate(1 To 1)
        ReDim szBusID(1 To 1)
        ReDim szDesStationID(1 To 1)
        ReDim szDesStationName(1 To 1)

        ReDim srSellResult(1 To 1)
        ReDim szBookNumber(1 To 1)
        ReDim aSellTicket(1 To 1)
        ReDim szSellStationID(1 To 1)
        ReDim szSellStationName(1 To 1)
        ReDim anInsurance(1 To 1)
        
        ReDim aSellTicket(1).BuyTicketInfo(1 To pnNums)
        ReDim aSellTicket(1).pasgSellTicketPrice(1 To pnNums)
        
        
        Dim aszSeats() As String
        Dim aszTicketNo() As String
        
        aszSeats = Split(pszSeats, cnSplit)
        aszTicketNo = Split(pszTicketNo, cnSplit)
        
        For i = 1 To pnNums
            aSellTicket(1).BuyTicketInfo(i).nTicketType = TP_FullPrice
            aSellTicket(1).BuyTicketInfo(i).szTicketNo = aszTicketNo(i - 1)
            aSellTicket(1).BuyTicketInfo(i).szSeatNo = aszSeats(i - 1)
            aSellTicket(1).pasgSellTicketPrice(i) = 0
            aSellTicket(1).BuyTicketInfo(i).szReserved = ""
            aSellTicket(1).BuyTicketInfo(i).szSeatTypeID = "01"
            aSellTicket(1).BuyTicketInfo(i).szSeatTypeName = ""
        Next i
        dyBusDate(1) = CDate(pszDate)
        szBusID(1) = pszBusID
        szDesStationID(1) = pszEndStationID
        szDesStationName(1) = ""

        szSellStationID(1) = pszSellStationID
        szSellStationName(1) = ""
        srSellResult = m_oSell.SellTicket(dyBusDate, szBusID, szSellStationID, szDesStationID, szDesStationName, aSellTicket, anInsurance, False)
        
        For i = 1 To pnNums
        
            szReturn = szReturn & aSellTicket(1).BuyTicketInfo(i).szTicketNo & cnSplit
            
        Next
        'szReturn = szReturn & cnSplit
        pszReturn = pnNums & cnSplit & szReturn
    SetSell = cszOK & cnSplit
    Exit Function
Error_Handle:
    SetSell = cszError & cnSplit & err.Description & cnSplit
End Function
'取网上售票
Public Function ListNetTicket(pszGetTicetID As String, pszVailData As String, pszCardID As String, ByRef pszReturn As String) As String
On Error GoTo Error_Handle
    Dim rs As Recordset
    Dim pszRsTemp As Recordset
    Dim i As Integer
    Dim szReturn As String
    Set pszRsTemp = m_oSell.InterNetValiDate(pszGetTicetID, pszVailData, pszCardID)
    If pszRsTemp.RecordCount > 0 Then
        pszRsTemp.MoveFirst
    End If
    
    For i = 1 To pszRsTemp.RecordCount
        Set rs = m_oSell.GetBusExRs(Trim(pszRsTemp!bus_date), Trim(pszRsTemp!bus_id), Trim(pszRsTemp!sell_station_id), Trim(pszRsTemp!des_station_id))
            szReturn = szReturn & Trim(rs!bus_id) & cnRecordSplit
            szReturn = szReturn & Trim(pszRsTemp!sell_station_id) & cnRecordSplit
            szReturn = szReturn & Trim(rs!sell_station_name) & cnRecordSplit
            szReturn = szReturn & Trim(pszRsTemp!des_station_id) & cnRecordSplit
            szReturn = szReturn & Trim(rs!station_name) & cnRecordSplit
            szReturn = szReturn & Trim(rs!vehicle_type_name) & cnRecordSplit
            szReturn = szReturn & Trim(pszRsTemp!pay_count) & cnRecordSplit
            szReturn = szReturn & Trim(rs!sell_check_gate_id) & cnRecordSplit
            szReturn = szReturn & Format(Trim(rs!bus_start_time2), "YYYY-MM-DD ") & cnRecordSplit
            szReturn = szReturn & Format(Trim(rs!bus_start_time2), "hh:mm") & cnRecordSplit
            szReturn = szReturn & Trim(pszRsTemp!seat_no) & cnRecordSplit
            szReturn = szReturn & cnSplit
            pszRsTemp.MoveNext
    Next i
    
    pszReturn = pszRsTemp.RecordCount & cnSplit & szReturn
    ListNetTicket = cszOK & cnSplit
    Exit Function
Error_Handle:
    ListNetTicket = cszError & cnSplit & err.Description & cnSplit
End Function

Public Function GetNetPrint(pszGetTicetID As String, pszVailData As String, pszCardID As String, pszTicketNo As String, ByRef pszReturn As String) As String
On Error GoTo Error_Handle
        Dim rs As Recordset
        Dim aszResult() As TSellTicketResult
        Dim atSellParam() As TSellTicketParam
        Dim i As Integer
        Dim szReturn As String
        Dim aszTicketNo() As String
        Dim j As Integer
        Dim rsTemp As Recordset
        Dim aTicketID(0 To 0) As String
        aszTicketNo = Split(pszTicketNo, cnSplit)
        Set rs = m_oSell.InterNetValiDate(pszGetTicetID, pszVailData, pszCardID)
        Set rsTemp = rs.Clone
        For j = 1 To rs.RecordCount
            rsTemp.Filter = "id='" & Trim(rs!id) & "'"
            aTicketID(0) = aszTicketNo(j - 1)
            aszResult = m_oSell.GetInterNetTicket(rsTemp, aTicketID, atSellParam)
            For i = 1 To rsTemp.RecordCount
            
                szReturn = szReturn & atSellParam(1).BuyTicketInfo(i).szTicketNo & cnSplit
                
            Next
            rs.MoveNext
        Next
        pszReturn = rs.RecordCount & cnSplit & szReturn
        GetNetPrint = cszOK & cnSplit
        Exit Function
Error_Handle:
    GetNetPrint = cszError & cnSplit & err.Description & cnSplit
End Function
'取预定票
Public Function ListBook(pszPhone As String, pszBookID As String, ByRef pszReturn As String) As String
On Error GoTo Error_Handle
    Dim rsTemp As Recordset
    Dim aszTemp() As String
    Dim i As Integer
    Dim szReturn As String
    Set rsTemp = m_oBook.GetBookedSeatEX(cszEmptyDateStr, aszTemp, pszPhone, pszBookID, , , ST_BOOKED)
        For i = 1 To rsTemp.RecordCount
            szReturn = szReturn & Trim(rsTemp!bus_id) & cnRecordSplit
            szReturn = szReturn & Trim(rsTemp!sell_station_id) & cnRecordSplit
            szReturn = szReturn & Trim(rsTemp!sell_station_name) & cnRecordSplit
            szReturn = szReturn & Trim(rsTemp!station_id) & cnRecordSplit
            szReturn = szReturn & Trim(rsTemp!station_name) & cnRecordSplit
            szReturn = szReturn & Trim(rsTemp!vehicle_type_name) & cnRecordSplit
            szReturn = szReturn & Trim(rsTemp!full_price) & cnRecordSplit
            szReturn = szReturn & Trim(rsTemp!check_gate_id) & cnRecordSplit
            szReturn = szReturn & Format(Trim(rsTemp!bus_start_time), "YYYY-MM-DD ") & cnRecordSplit
            szReturn = szReturn & Format(Trim(rsTemp!bus_start_time), "hh:mm") & cnRecordSplit
            szReturn = szReturn & Trim(rsTemp!seat_no) & cnRecordSplit

            szReturn = szReturn & Trim(rsTemp!telephone) & cnRecordSplit
            szReturn = szReturn & Trim(rsTemp!book_number) & cnRecordSplit
            If Trim(rsTemp!Status) = 0 Then
                szReturn = szReturn & "预定" & cnRecordSplit
            ElseIf Trim(rsTemp!Status) = 1 Then
                szReturn = szReturn & "已取消" & cnRecordSplit
            
            ElseIf Trim(rsTemp!Status) = 2 Then
                szReturn = szReturn & "已取票" & cnRecordSplit
            ElseIf Trim(rsTemp!Status) = 3 Then
                szReturn = szReturn & "已被其他人买走" & cnRecordSplit
            
            
            End If

            szReturn = szReturn & cnSplit
            rsTemp.MoveNext
        Next
        
        pszReturn = rsTemp.RecordCount & cnSplit & szReturn
        ListBook = cszOK & cnSplit
    Exit Function
Error_Handle:
    ListBook = cszError & cnSplit & err.Description & cnSplit
End Function
'把预定车票改为预留
Public Function SetBookChange(pszPhone As String, pszBookID As String) As String
On Error GoTo Error_Handle
        Dim i As Integer
        Dim rsTemp As Recordset
        Dim aszTemp() As String
        Dim szReturn As String
        
        Dim dyBusDate() As Date
        Dim szBusID() As String
        Dim szDesStationID() As String
        Dim szDesStationName() As String
        Dim psgDiscount() As Single
        Dim srSellResult() As TSellTicketResult
        Dim szBookNumber() As String
        Dim aSellTicket() As TSellTicketParam
        Dim szSellStationID() As String
        Dim szSellStationName() As String
        Dim anInsurance() As Integer '售票用
        
        ReDim dyBusDate(1 To 1)
        ReDim szBusID(1 To 1)
        ReDim szDesStationID(1 To 1)
        ReDim szDesStationName(1 To 1)
        ReDim psgDiscount(1 To 1)
        ReDim srSellResult(1 To 1)
        ReDim szBookNumber(1 To 1)
        ReDim aSellTicket(1 To 1)
        ReDim szSellStationID(1 To 1)
        ReDim szSellStationName(1 To 1)
        ReDim anInsurance(1 To 1)
        

        
        Dim szTmpBusID As String
        Set rsTemp = m_oBook.GetBookedSeatEX(cszEmptyDateStr, aszTemp, pszPhone, pszBookID, , , ST_BOOKED)
        If rsTemp.RecordCount > 0 Then
            For i = 1 To rsTemp.RecordCount
                If i = 1 Then
                    szTmpBusID = FormatDbValue(rsTemp!bus_id)
                    
                Else
                    If szTmpBusID <> FormatDbValue(rsTemp!bus_id) Then
                        err.Raise 9999, , "自助机不支持取多个车次车票"
                    End If
                    
                End If
                rsTemp.MoveNext
            Next
            ReDim aSellTicket(1).BuyTicketInfo(1 To rsTemp.RecordCount)
            ReDim aSellTicket(1).pasgSellTicketPrice(1 To rsTemp.RecordCount)
            rsTemp.MoveFirst
            
            For i = 1 To rsTemp.RecordCount
                If i = 1 Then
                
                    dyBusDate(1) = CDate(FormatDbValue(rsTemp!bus_date))
                    szBusID(1) = FormatDbValue(rsTemp!bus_id)
                    szDesStationID(1) = FormatDbValue(rsTemp!station_id)
                    szDesStationName(1) = ""
                    psgDiscount(1) = CSng(0)
                    szSellStationID(1) = FormatDbValue(rsTemp!sell_station_id)
                    szSellStationName(1) = ""
                End If
                aSellTicket(1).BuyTicketInfo(i).nTicketType = TP_FullPrice
                aSellTicket(1).BuyTicketInfo(i).szTicketNo = ""
                aSellTicket(1).BuyTicketInfo(i).szSeatNo = FormatDbValue(rsTemp!seat_no)
                aSellTicket(1).pasgSellTicketPrice(i) = 0
                aSellTicket(1).BuyTicketInfo(i).szReserved = ""
                aSellTicket(1).BuyTicketInfo(i).szSeatTypeID = "01"
                aSellTicket(1).BuyTicketInfo(i).szSeatTypeName = ""
                rsTemp.MoveNext
            Next i
            

             m_oSell.SetBookChange dyBusDate, szBusID, szSellStationID, szDesStationID, aSellTicket, anInsurance

            
        Else
            err.Raise 999, , "预定转预留失败"
        End If
        SetBookChange = cszOK & cnSplit
    Exit Function
Error_Handle:
    SetBookChange = cszError & cnSplit & err.Description & cnSplit
End Function

Public Function SellScroll(pszBusID As String, pszDate As String, pszSellStationID As String, pszEndStationID As String, pnNums As Integer, pszTicketNo As String, ByRef pszReturn As String) As String
On Error GoTo Error_Handle
        Dim i As Integer
        Dim szReturn As String
        
        Dim dyBusDate() As Date
        Dim szBusID() As String
        Dim szDesStationID() As String
        Dim szDesStationName() As String

        Dim srSellResult() As TSellTicketResult
        Dim szBookNumber() As String
        Dim aSellTicket() As TSellTicketParam
        Dim szSellStationID() As String
        Dim szSellStationName() As String
        Dim anInsurance() As Integer '售票用
        
        ReDim dyBusDate(1 To 1)
        ReDim szBusID(1 To 1)
        ReDim szDesStationID(1 To 1)
        ReDim szDesStationName(1 To 1)

        ReDim srSellResult(1 To 1)
        ReDim szBookNumber(1 To 1)
        ReDim aSellTicket(1 To 1)
        ReDim szSellStationID(1 To 1)
        ReDim szSellStationName(1 To 1)
        ReDim anInsurance(1 To 1)
        
        ReDim aSellTicket(1).BuyTicketInfo(1 To pnNums)
        ReDim aSellTicket(1).pasgSellTicketPrice(1 To pnNums)
        
        
'        Dim aszSeats() As String
        Dim aszTicketNo() As String
'
'        aszSeats = Split(pszSeats, cnSplit)
        aszTicketNo = Split(pszTicketNo, cnSplit)
'
        For i = 1 To pnNums
            aSellTicket(1).BuyTicketInfo(i).nTicketType = TP_FullPrice
            aSellTicket(1).BuyTicketInfo(i).szTicketNo = aszTicketNo(i - 1)
            aSellTicket(1).BuyTicketInfo(i).szSeatNo = ""
            aSellTicket(1).pasgSellTicketPrice(i) = 0
            aSellTicket(1).BuyTicketInfo(i).szReserved = ""
            aSellTicket(1).BuyTicketInfo(i).szSeatTypeID = "01"
            aSellTicket(1).BuyTicketInfo(i).szSeatTypeName = ""
        Next i
        dyBusDate(1) = CDate(pszDate)
        szBusID(1) = pszBusID
        szDesStationID(1) = pszEndStationID
        szDesStationName(1) = ""

        szSellStationID(1) = pszSellStationID
        szSellStationName(1) = ""
        srSellResult = m_oSell.SellTicket(dyBusDate, szBusID, szSellStationID, szDesStationID, szDesStationName, aSellTicket, anInsurance, True)
        
        For i = 1 To pnNums
        
            szReturn = szReturn & aSellTicket(1).BuyTicketInfo(i).szTicketNo & cnSplit
            
        Next
        'szReturn = szReturn & cnSplit
        pszReturn = pnNums & cnSplit & szReturn
    SellScroll = cszOK & cnSplit
    Exit Function
Error_Handle:
    SellScroll = cszError & cnSplit & err.Description & cnSplit
End Function

