Attribute VB_Name = "mdlMain"
Option Explicit

Public Const cszPrimaryKey = "SellTk"
Public Const cszSubKey_ExtraSellType = "ExtraSellType"

Public m_aszCheckGateInfo() As String '检票口

Public m_nPrintBusIDLen As Integer
Public m_bPrintScrollBusMode As Boolean

Public m_szRegValue As String

Public Type TPrintTicketInfo
    szTicketNo As String
    nTicketType As ETicketType
    sgTicketPrice As Double
    szSeatNo As String
End Type

Public Type TPrintTicketParam
    aptPrintTicketInfo() As TPrintTicketInfo
End Type

Public Enum ETaskType
    RT_SellTicket = 1 '售票
    RT_ExtraSellTicket = 2 '补票
    RT_ChangeTicket = 3 '改签
    RT_ReturnTicket = 4 '退票
    RT_CancelTicket = 5 '废票
End Enum

'主界面状态条字符串区域
Public Enum EStatusBarArea
    ESB_WorkingInfo = 1
    ESB_ResultCountInfo = 2
    ESB_UserInfo = 3
    ESB_LoginTime = 4
End Enum
Public Enum EBusInfoIndex
    ID_BusType = 1
    ID_OffTime = 2
    ID_RouteName = 3
    ID_EndStation = 4
    ID_TotalSeat = 5
    ID_BookCount = 6
    ID_SeatCount = 7
    ID_SeatTypeCount = 8
    ID_BedTypeCount = 9
    ID_AdditionalCount = 10
    ID_VehicleModel = 11
    ID_FullPrice = 12
    ID_HalfPrice = 13
    ID_FreePrice = 14
    ID_PreferentialPrice1 = 15
    ID_PreferentialPrice2 = 16
    ID_PreferentialPrice3 = 17
    ID_BedFullPrice = 18
    ID_BedHalfPrice = 19
    ID_BedFreePrice = 20
    ID_BedPreferentialPrice1 = 21
    ID_BedPreferentialPrice2 = 22
    ID_BedPreferentialPrice3 = 23
    ID_AdditionalFullPrice = 24
    ID_AdditionalHalfPrice = 25
    ID_AdditionalFreePrice = 26
    ID_AdditionalPreferential1 = 27
    ID_AdditionalPreferential2 = 28
    ID_AdditionalPreferential3 = 29
    ID_LimitedCount = 30
    ID_LimitedTime = 31
    ID_BusType1 = 32
    ID_CheckGate = 33
    ID_StandCount = 34
'    ID_SellStationID = 35
'    ID_SellStationName = 36
End Enum
'暂放的票的枚举
Public Enum InstantTicketInfo
    IT_BusType = 1
    IT_OffTime = 2
    IT_BusDate = 3
    IT_StartStation = 4
    IT_EndStation = 5
    IT_VehicleModel = 6
    IT_SumTicketNum = 7
    IT_SumPrice = 8
    IT_OrderSeat = 9
    IT_FullPrice = 10
    IT_FullNum = 11
    IT_HalfPrice = 12
    IT_HalfNum = 13
    IT_FreeType = 14
    IT_FreeNum = 15
    IT_PreferentialType1 = 16
    IT_PreferentialNum1 = 17
    IT_PreferentialType2 = 18
    IT_PreferentialNum2 = 19
    IT_PreferentialType3 = 20
    IT_PreferentialNum3 = 21
    IT_DiscountPrice = 22
    IT_Discount = 23
    IT_StandCount = 24
    IT_CheckGate = 25
    IT_LimitedCount = 26
    IT_BoundText = 27
    IT_SetSeatEnable = 28
    IT_SetSeatValue = 29
    IT_SeatNo = 30
    IT_TicketPrice = 31
    IT_TicketType = 32
    IT_SeatType = 33
    IT_TerminateName = 34
  
End Enum


Public Enum ESortKeyChange
    SK_VehicleModel = 1
    SK_OffTime = 2
    SK_TicketPrice = 3
    SK_SeatCount = 4
End Enum


Public Const cszSellTicket = "Sell"
Public Const cszExtraSellTicket = "Extra"
Public Const cszChangeTicket = "Change"
Public Const cszReturnTicket = "Return"
Public Const cszCancelTicket = "Cancel"

Public Const cszHelp = "Help"
Public Const cszAbout = "About"
Public Const cszExit = "Exit"
Public Const clActiveColor = &HFF0000
Public Const cszScrollBus = "滚动"
Public Const cszScrollBusTime = "之前"

Public Const cszMiddleTime = "11:30" '中午的时间


''定义快捷键
'
'Public Const cnKeySetSeat = vbKeyF8
'Public Const cnKeyChangeSeatType = vbKeyF9


'*****************************************************
Public m_clSell As New Collection '售票窗口集合
Public m_clChange As New Collection '改签窗口集合
'Public m_clExtra As New Collection '补票窗口集合
Public m_clReturn As New Collection '退票窗口集合
Public m_clCancel As New Collection '废票窗口集合




'*****************************************************
'Public m_oAUser As ActiveUser
'Public m_oSell As New SellTicketClient
'Public m_oSellService As New SellTicketService
'Public m_oParam As New SystemParam
Public m_bSellStationCanSellEachOther As Boolean
'Public m_oCmdDlg As New STShell.CommDialog
'Public m_oShell As New STShell.CommShell

Public m_lTicketNo As Long
Public m_lEndTicketNo As Long '结束票号(fpd添加）
Public m_szTicketPrefix As String

Private m_szTicketNoFromatStr As String

Public m_szCurrentUnitID As String '当前提供票务服务的单位
Public m_nCurrentTask As ETaskType  '当前的票务类型

Public m_bSelfChangeUnitOrFun As Boolean
Public m_lStopBusColor As OLE_COLOR
Public m_lNormalBusColor As OLE_COLOR


Public m_ISellScreenShow As Integer      '是否分屏显示  2006-01-20 qlh
Public g_nBookTime As Long '预定释放时间(单位:分钟)
Public g_bIsBookValid As Boolean '是否使用预定系统
Public g_nDiscountTicketInTicketTypePosition As Integer '折扣票在票种项的位置
Public m_bListNoSeatBus As Boolean      '是否列出已售完车次,2005-12-6 lyq追加

Public m_aszFirstBus() As String
Public m_aszFirstStation() As String
'设定某个站点？车次显示在第一行
Private m_szLastStatus As String '保存的状态栏内的状态



Public m_nCanSellDay  As Integer
Public m_bUseFastPrint As Boolean       '是否使用快速打印

Public g_aszAllStartStation() As String '所有的起点站及名称

Private m_szDatabaseType As String
Private m_szServer As String
Private m_szUser As String
Private m_szPassword As String
Private m_szDatabase As String
Private m_szTimeout As String
Public m_Sell() As String
Public Sub Main()
    Dim i As Integer
'    Dim oSysMan As New User
    Dim szLoginCommandLine As String
    Dim FileNo As Integer
    Dim nUseFastPrint As Integer
    
    On Error GoTo Error_Handle
    
    If App.PrevInstance Then
        End
    End If
    If Not IsPrinterValid Then
        MsgBox "打印机未配置！", vbInformation, "打印机出错:"
        End
        Exit Sub
    End If
    
    '====================================================================
    '读取自定义数据
    '====================================================================
    FileNo = FreeFile
    Open App.Path + "\Param.ini" For Input As #FileNo
        On Error Resume Next
        '自定义数据
        Input #FileNo, m_szRemoteHost '远程机器
        Input #FileNo, m_szRemotePort '远程端口
        Input #FileNo, nUseFastPrint '是否使用快速打印
        
        Input #FileNo, m_szDatabaseType '
        Input #FileNo, m_szServer '
        Input #FileNo, m_szUser '
        Input #FileNo, m_szPassword '
        Input #FileNo, m_szDatabase '
        Input #FileNo, m_szTimeout '
        If err.Number <> 0 Then
            MsgBox "配置文件Param.ini错误"
            End
        End If
        m_bUseFastPrint = IIf(nUseFastPrint = 0, False, True)
    Close #FileNo
    
    
    If Not ShowLogin Then
        '如果登陆错误,则退出
        End
    End If
    

    FileNo = FreeFile
    Open App.Path + "\StartStation.ini" For Input As #FileNo
        Input #FileNo, m_nSellStationCount '车站个数
        If m_nSellStationCount > 0 Then
            ReDim g_aszAllStartStation(1 To m_nSellStationCount, 1 To 4)
            For i = 1 To m_nSellStationCount
                Input #FileNo, g_aszAllStartStation(i, 1) '起点站序号
                Input #FileNo, g_aszAllStartStation(i, 2) '起点站代码
                Input #FileNo, g_aszAllStartStation(i, 3) '起点站名
                Input #FileNo, g_aszAllStartStation(i, 4) '起点站对应的上车站代码
                
            Next i
        End If
        If err.Number <> 0 Then
            MsgBox "配置文件StartStation.ini错误"
            End
        End If
    Close #FileNo
    '====================================================================
    
    
    
    m_bSelfChangeUnitOrFun = False
    m_lStopBusColor = RGB(255, 0, 0)
    m_lNormalBusColor = RGB(0, 0, 0)
    m_nCurrentTask = RT_SellTicket
    m_nCanSellDay = 15
    m_bSellStationCanSellEachOther = False
    m_nPrintBusIDLen = 5
    m_bPrintScrollBusMode = False
    m_bUseFastPrint = True
    m_bListNoSeatBus = False
    m_ISellScreenShow = False
    g_nBookTime = 30
    g_bIsBookValid = True
    
    GetIniFile
    m_szRegValue = GetRegInfo
    '*****************
    '语音显示器
    g_lComPort = IIf(Val(ReadReg(cszComPort)) = 2, 2, 1)

    SetInit
    '*****************



    '初始化检票口
    GetInitCheckGate  '此语句临时注释
    Load MDISellTicket
    MDISellTicket.SetPrintEnabled False
    MDISellTicket.Show
    
    Exit Sub
Error_Handle:
    ShowErrorMsg
End Sub

Private Sub ReadSetFistData()
Dim szStationID As String
Dim szBusID As String

Dim aszStationID() As String
Dim aszBusID() As String

szStationID = ReadReg("FirstStation")
szBusID = ReadReg("FirstBus")
m_aszFirstBus = ConvertStrToArray(szBusID)
m_aszFirstStation = ConvertStrToArray(szStationID)



End Sub

Private Function ConvertStrToArray(szStr) As String()
Dim aszArray() As String
Dim iCount As Integer
Dim iStart As Integer


iStart = 1
iCount = 0
Do While InStr(iStart, szStr, ",") > 0
    iCount = iCount + 1
    ReDim Preserve aszArray(1 To iCount)
    aszArray(iCount) = Mid(szStr, iStart, InStr(iStart, szStr, ",") - iStart)
    iStart = InStr(iStart, szStr, ",") + 1
Loop
If Trim(Mid(szStr, iStart, Len(szStr) - iStart + 1)) <> "" Then
   iCount = iCount + 1
   ReDim Preserve aszArray(1 To iCount)
   aszArray(iCount) = Mid(szStr, iStart, Len(szStr) - iStart + 1)
End If
ConvertStrToArray = aszArray
End Function





Public Sub AdjustLocation(aForm As Form)
    aForm.Left = (Screen.Width - 640 * Screen.TwipsPerPixelX) / 2
    aForm.Top = (Screen.Height - 480 * Screen.TwipsPerPixelY) / 2
End Sub

Public Sub AdjustFraLoc(AFra As Frame)
    AFra.Left = 10
    AFra.Top = 74
End Sub

Public Function ToStandardDateStr(pdtDate As Date) As String
    ToStandardDateStr = Format(pdtDate, "YYYY-MM-DD")
End Function

Public Function ToStandardTimeStr(pdtTime As Date) As String
    ToStandardTimeStr = Format(pdtTime, "hh:mm:ss")
End Function

Public Function ToStandardDateTimeStr(pdtDateTime As Date) As String
    ToStandardDateTimeStr = Format(pdtDateTime, "YYYY-MM-DD hh:mm:ss")
End Function

Public Function GetTicketNo(Optional pnOffset As Integer = 0) As String
    GetTicketNo = MakeTicketNo(m_lTicketNo + pnOffset, m_szTicketPrefix)
End Function

Public Function GetEndTicketNo(Optional pnOffset As Integer = 0) As String
    GetEndTicketNo = MakeEndTicketNo(m_lEndTicketNo + pnOffset, m_szTicketPrefix)
End Function

Public Sub IncTicketNo(Optional pnOffset As Integer = 1, Optional pbNoShow As Boolean = False)
    m_lTicketNo = m_lTicketNo + pnOffset
    If Not pbNoShow Then
        MDISellTicket.lblTicketNo.Caption = GetTicketNo()
        MDISellTicket.SetLeaveNum
    End If
End Sub

Public Function MakeTicketNo(plTicketNo As Long, Optional pszPrefix As String = "") As String
    MakeTicketNo = pszPrefix & Format(plTicketNo, TicketNoFormatStr())
End Function

Public Function MakeEndTicketNo(plEndTicketNo As Long, Optional pszPrefix As String = "") As String
    MakeEndTicketNo = pszPrefix & Format(plEndTicketNo, TicketNoFormatStr())
End Function

Public Function TicketNoNumLen() As Integer
'    If m_lTicketNoNumLen = 0 Then
'        m_lTicketNoNumLen = 8 'm_oParam.TicketNumberLen
'    End If
    TicketNoNumLen = m_lTicketNoNumLen
End Function

Private Function TicketNoFormatStr() As String
    Dim i As Integer
    If m_szTicketNoFromatStr = "" Then
        m_szTicketNoFromatStr = String(TicketNoNumLen(), "0")
    End If
    TicketNoFormatStr = m_szTicketNoFromatStr
End Function

Public Function ResolveTicketNo(pszFullTicketNo, ByRef pszTicketPrefix As String) As Long
'    Dim i As Integer, j As Integer
'    Dim nCount As Integer, nTemp As Integer, nTicketPrefixLen As Integer
'    'On Error Resume Next
'    pszFullTicketNo = Trim(pszFullTicketNo)
'    nCount = Len(pszFullTicketNo)
'
'    For i = 1 To nCount
'        nTemp = Asc(Mid(pszFullTicketNo, nCount - i + 1, 1))
'        If nTemp < vbKey0 Or nTemp > vbKey9 Then
'            Exit For
'        End If
'    Next
'    i = i - 1
'    If i > 0 Then
'        nTemp = TicketNoNumLen()
'        nTemp = IIf(nTemp > i, i, nTemp)
'        ResolveTicketNo = CLng(Right(pszFullTicketNo, nTemp))
'
'        nTicketPrefixLen = m_oParam.TicketPrefixLen
'        If nTicketPrefixLen <= Len(pszFullTicketNo) Then
'            pszTicketPrefix = Left(pszFullTicketNo, nTicketPrefixLen)
'        Else
'            pszTicketPrefix = pszFullTicketNo
'        End If
'
'    Else
'        pszTicketPrefix = ""
'        ResolveTicketNo = 0
'    End If
    
    
    
End Function


Public Sub GetAppSetting()
    Dim szLastTicketNo As String
    szLastTicketNo = GetLastTicketNo(GetActiveUserID)
    m_lTicketNo = szLastTicketNo
    IncTicketNo , True
End Sub

Private Function GetLastTicketNo(Seller As String) As String
    '得到当前售票员的最后票号
    Dim odb As New ADODB.Connection
    Dim rsTemp As Recordset
    Dim szSql As String
    Dim szWhere As String
    Dim szOperatorID As String
    odb.ConnectionString = GetConnectionStr
    odb.CursorLocation = adUseClient
    odb.Open
        
    szSql = "SELECT MAX(id) AS ticket_no FROM tickets WHERE " _
    & " SellDate=(" _
    & " SELECT MAX(SellDate) FROM tickets s,user_info u WHERE " _
    & " bank_id='" & m_cszOperatorBankID & "' and s.status=1) "
    Set rsTemp = odb.Execute(szSql)
    If rsTemp.RecordCount = 1 Then
        GetLastTicketNo = FormatDbValue(rsTemp!ticket_no)
    Else
        GetLastTicketNo = 0
    End If
End Function


Public Function GetObjecInCollection(pvaIndex As Variant, pclCollection As Collection) As Object
    On Error GoTo here
    Set GetObjecInCollection = pclCollection(pvaIndex)
    Exit Function
here:
End Function

Public Sub ShowStatusInMDI(pszMsg As String)
    m_szLastStatus = MDISellTicket.abMenuTool.Bands("statusBar").Tools("pnWorkingInfo").Caption
'    m_szLastStatus = MDISellTicket.sbMain.Panels(1)
'    MDISellTicket.sbMain.Panels(1) = pszMsg
End Sub

Public Sub RestoreStatusInMDI()
    ShowSBInfo m_szLastStatus
'    MDISellTicket.sbMain.Panels(1) = m_szLastStatus
End Sub

'得到一以逗号(,)分隔的字符串中的项数
Public Function GetTotalSeat(pszSeatStr As String) As Integer
    Dim i As Integer, j As Integer
        
    i = 0
    If pszSeatStr <> "" Then
        For j = 1 To Len(pszSeatStr)
            If Mid(pszSeatStr, j, 1) = "," Then i = i + 1
        Next
        i = i + 1
    End If
    
    GetTotalSeat = i
End Function

'得到一以逗号(,)分隔的字符串中的指定序号的项
Public Function GetSeatNo(pszSeatStr As String, pnIndex As Integer) As String
    Dim i As Integer, j As Integer
    Dim szTemp As String
    Dim nCount As Integer
    Dim nTemp As Integer
    If pszSeatStr <> "" Then
        i = 1
        If pnIndex > 1 Then
            nCount = pnIndex - 1
            For i = 1 To Len(pszSeatStr)
                If Mid(pszSeatStr, i, 1) = "," Then
                    nCount = nCount - 1
                    If nCount = 0 Then
                        i = i + 1
                        Exit For
                    End If
                End If
            Next
        End If
        nTemp = InStr(i, pszSeatStr, ",", vbTextCompare)
        If nTemp > 0 Then
            GetSeatNo = Trim(Mid(pszSeatStr, i, nTemp - i))
        Else
            GetSeatNo = Trim(Mid(pszSeatStr, i))
        End If
    Else
        GetSeatNo = ""
    End If
End Function

''从售票信息和售票结果信息中得到相应票价的信息
'Public Function SelfGetTicketPriceFromSellResult(ByVal pnTicketType As ETicketType, pabiTemp() As TBuyTicketInfo, psrTemp As TSellTicketResult) As Double
'    Dim nCount  As Integer, i As Integer
'    Dim nTicketType As ETicketType
'
'    nCount = ArrayLength(pabiTemp)
'    For i = 1 To nCount
'        If pabiTemp(i).nTicketType <> TP_HalfPrice Then
'            nTicketType = TP_FullPrice
'        Else
'            nTicketType = TP_HalfPrice
'        End If
'        If nTicketType = pnTicketType Then
'            SelfGetTicketPriceFromSellResult = psrTemp.asgTicketPrice(i)
'            Exit For
'        End If
'    Next
'End Function


'计算指定LISTVIEW的指定列的所选项(如没有任何一行选取则视为全部)数值总和
Public Function CaculateListView(plvListView As ListView, panIndex() As Integer) As Double()
    Dim asgTemp() As Double
    Dim nTemp As Integer
    Dim i As Integer, j As Integer, k As Integer
    Dim liTemp As ListItem
    Dim bAll As Boolean
    
    If plvListView.SelectedItem Is Nothing Then
        bAll = True
    Else
        bAll = False
    End If
    
    nTemp = ArrayLength(panIndex)
    If nTemp > 0 Then
        ReDim asgTemp(1 To nTemp)
        For i = 1 To plvListView.ListItems.count
            Set liTemp = plvListView.ListItems(i)
            
            If (bAll Or liTemp.Selected) Then
                k = 1
                For j = 1 To nTemp
                    If panIndex(j) = 1 Then
                        asgTemp(k) = asgTemp(k) + CDbl(plvListView.ListItems(i))
                        k = k + 1
                    Else
                        asgTemp(k) = asgTemp(k) + CDbl(liTemp.SubItems(panIndex(j) - 1))
                        k = k + 1
                    End If
                Next
            End If
        Next
    End If
    CaculateListView = asgTemp
    If Not liTemp Is Nothing Then Set liTemp = Nothing
End Function

Public Sub DecBusListViewSeatInfo(plvListView As ListView, pnCount As Integer, pbSeatCount As Boolean)
    Dim liTemp As ListItem
    Set liTemp = plvListView.SelectedItem
    If Not liTemp Is Nothing And liTemp.SubItems(ID_BusType1) <> TP_ScrollBus And pnCount > 0 Then
        If pbSeatCount Then
            liTemp.SubItems(ID_SeatCount) = CInt(liTemp.SubItems(ID_SeatCount)) - pnCount
        Else
            liTemp.SubItems(ID_StandCount) = CInt(liTemp.SubItems(ID_StandCount)) - pnCount
        End If
        
    End If
    Set liTemp = Nothing
End Sub



Public Function DealWithChildKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer) As Long
    Dim nIndex As Integer
    If Shift And 2 <> 0 Then '如果Alt键按下
        nIndex = KeyCode - vbKey1 + 1
        If nIndex >= 1 And nIndex <= MDISellTicket.tsUnit.Tabs.count Then
            MDISellTicket.tsUnit.Tabs(nIndex).Selected = True
        End If
    End If
End Function

'从站点组合框中的字符串得到站点名称
Public Function GetStationNameInCbo(pszText As String) As String
    Dim szTemp As String
    Dim nTemp As Integer
    nTemp = InStr(1, Trim(pszText), " ", vbTextCompare)
    If nTemp > 0 Then
        szTemp = Trim(Mid(Trim(pszText), nTemp))
        nTemp = InStr(1, szTemp, " ", vbTextCompare)
        If nTemp > 0 Then
            GetStationNameInCbo = Left(szTemp, nTemp - 1)
        Else
            GetStationNameInCbo = szTemp
        End If
    End If
End Function

Public Function ReadReg(szSubKey As String) As String
    Dim oReg As New CFreeReg
    Dim szTmp As String
    oReg.Init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany
    szTmp = oReg.GetSetting(cszPrimaryKey, szSubKey)
    ReadReg = szTmp
End Function
Public Function WriteReg(szSubKey As String, szValue As String) As Boolean
    Dim oReg As New CFreeReg
    oReg.Init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany
    oReg.SaveSetting cszPrimaryKey, szSubKey, szValue
End Function


Public Function GetMenuUnitName(pszUnitName1 As String) As String
'    GetMenuUnitName = Left(pszUnitName1, InStr(1, pszUnitName1, "(", vbTextCompare) - 1)
End Function


Public Function GetTicketTypeStr2(ByVal pnTicketType As Integer) As String
Dim j As Integer
Dim TicketType() As TTicketType
Dim intEnableTicketNo As Integer

   TicketType = GetAllTicketType(1)
   intEnableTicketNo = UBound(TicketType) - LBound(TicketType) + 1
    For j = 1 To intEnableTicketNo
        If TicketType(j).nTicketTypeID = pnTicketType And TicketType(j).nTicketTypeValid = TP_TicketTypeValid Then
           GetTicketTypeStr2 = TicketType(j).szTicketTypeName
           Exit For
        End If
    Next j
End Function

'得到打印车次代码
Public Function GetBusID(pszBusID As String) As String
    Dim nBusIDLen As Integer
    Dim szBusID As String
    nBusIDLen = m_nPrintBusIDLen
    If nBusIDLen = 0 Then
        szBusID = Trim(pszBusID)
    Else
        szBusID = Right(Trim(pszBusID), nBusIDLen)
    End If
    GetBusID = szBusID
End Function


'得到滚动车次发车时间打印方式
Public Function GetPrintScrollMode() As Boolean
    GetPrintScrollMode = m_bPrintScrollBusMode
End Function


''判断车次是否属于有座类型的车次
'Public Function IsSeatTypeBus(pdBusDate As Date, pdBusID As String, pSeatTypeBus As TMultiSeatTypeBus) As Boolean
'Dim nLen As Integer
'Dim i As Integer
'nLen = 0
'nLen = ArrayLength(pSeatTypeBus.adBusDate)
'For i = 1 To nLen
'    If pdBusDate = pSeatTypeBus.adBusDate(i) And pdBusID = pSeatTypeBus.aszBusID(i) Then
'        IsSeatTypeBus = True
'        Exit Function
'    End If
'Next i
'IsSeatTypeBus = False
'End Function

'给数组赋初始值
Public Sub SetArrayInit(aInitArray() As Variant, InitValue As Variant)
Dim nLen As Integer
Dim i As Integer
nLen = 0
nLen = ArrayLength(aInitArray)
For i = 1 To nLen
    aInitArray(i) = InitValue
Next i
End Sub

'得到座位数
Public Function GetSeatCount(szSeatNo As String, nTotalSeat As Integer) As Integer
Dim i As Integer
Dim nCount As Integer
Dim nLen As Integer
nLen = 0
nCount = 0
nLen = Len(szSeatNo)
For i = 1 To nLen
    If Mid(szSeatNo, i, 1) = "," Then nCount = nCount + 1
    
Next i
nCount = nCount + 1
If nCount <= nTotalSeat Then
    GetSeatCount = nCount
Else
    GetSeatCount = nTotalSeat
End If

End Function

'得到座位号
Public Function GetSeatNumber(szSeatNo As String, nSeatNo As Integer) As String()
    Dim aszSeatNo() As String
    Dim i As Integer
    Dim nCount As Integer
    Dim nLast As Integer
    nCount = 0
    ReDim aszSeatNo(1 To nSeatNo)
    nLast = 1
    For i = 1 To Len(szSeatNo)
        If Mid(szSeatNo, i, 1) = "," Then
            If nCount < nSeatNo Then
                nCount = nCount + 1
                aszSeatNo(nCount) = Mid(szSeatNo, nLast, i - nLast)
                nLast = i + 1
            Else
                GetSeatNumber = aszSeatNo
                Exit Function
            End If
        End If
    Next i
    If nCount = nSeatNo - 1 Then
        aszSeatNo(nCount + 1) = Mid(szSeatNo, nLast, Len(szSeatNo) - nLast)
    End If
    GetSeatNumber = aszSeatNo
End Function
'ListView出错处理
Public Function ReturnValue(pszListView As String) As Integer
    On Error GoTo here
    ReturnValue = Val(pszListView)
Exit Function
here:
    ReturnValue = 0
End Function


Public Function GetRegInfo() As String
    Dim oReg As New CFreeReg
    Dim szValue As String
    oReg.Init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany
    szValue = oReg.GetSetting("SNSellTK", "CheckGateID", "")
    GetRegInfo = szValue
End Function



' *******************************************************************
' *   Member Name: ShowSBInfo                                      *
' *   Brief Description: 写系统状态条信息                           *
' *   Engineer: 陆勇庆                                              *
' *******************************************************************
Public Sub ShowSBInfo(Optional pszInfo As String = "", Optional peArea As EStatusBarArea = ESB_WorkingInfo)
'参数注释
'*************************************
'pnArea(状态条区域,默认为应用程序状态区)
'pszInfo(信息内容)
'*************************************
    With MDISellTicket
        Select Case peArea
            Case EStatusBarArea.ESB_WorkingInfo
                .abMenuTool.Bands("statusBar").Tools("pnWorkingInfo").Caption = pszInfo
            Case EStatusBarArea.ESB_ResultCountInfo
                .abMenuTool.Bands("statusBar").Tools("pnResultCountInfo").Caption = pszInfo
            Case EStatusBarArea.ESB_UserInfo
                .abMenuTool.Bands("statusBar").Tools("pnUserInfo").Caption = pszInfo
            Case EStatusBarArea.ESB_LoginTime
                If pszInfo <> "" Then pszInfo = "登录时间: " & pszInfo
                .abMenuTool.Bands("statusBar").Tools("pnLoginTime").Caption = pszInfo
        End Select
    End With
End Sub
' *******************************************************************
' *   Member Name: WriteProcessBar                                  *
' *   Brief Description: 写系统进程条状态                           *
' *   Engineer: 陆勇庆                                              *
' *******************************************************************
Public Sub WriteProcessBar(Optional pbVisual As Boolean = True, Optional ByVal plCurrValue As Variant = 0, Optional ByVal plMaxValue As Variant = 0, Optional pszShowInfo As String = cszUnrepeatString)
'参数注释
'*************************************
'plCurrValue(当前进度值)
'plMaxValue(最大进度值)
'*************************************
    If plMaxValue = 0 And pbVisual = True Then Exit Sub
    Dim nCurrProcess As Integer
    With MDISellTicket.abMenuTool.Bands("statusBar").Tools("progressBar")
        If pbVisual Then
            If Not .Visible Then
                .Visible = True
                MDISellTicket.pbLoad.Max = 100
            End If
            nCurrProcess = Int(plCurrValue / plMaxValue * 100)
            MDISellTicket.pbLoad.Value = nCurrProcess
        Else
            .Visible = False
        End If
    End With
    If pszShowInfo <> cszUnrepeatString Then ShowSBInfo pszShowInfo, ESB_WorkingInfo
End Sub


Public Sub GetInitCheckGate()
'    Dim szTemp As String
'
'    szTemp = m_oSell.SellUnitCode
'    m_oSell.SellUnitCode = m_szCurrentUnitID
'
'    m_aszCheckGateInfo = m_oSell.GetAllCheckGate()
'
'    m_oSell.SellUnitCode = szTemp
    
End Sub

'得到检票口名称和代码
Public Function GetCheckName(pszCheckGateID As String) As String
    Dim i As Integer
    Dim szResult As String
    Dim nLen As Integer
    nLen = 0
    nLen = ArrayLength(m_aszCheckGateInfo)
    szResult = ""
    For i = 1 To nLen
        If Trim(m_aszCheckGateInfo(i, 1)) = Trim(pszCheckGateID) Then
            szResult = Trim(m_aszCheckGateInfo(i, 2))
            Exit For
        End If
    Next i
    GetCheckName = szResult

End Function



Public Function ShowLogin() As Boolean
    '此处读取数据库
    Dim frmTemp As New frmLogin
    Dim bSuccess As Boolean
    Dim odb As New ADODB.Connection
    Dim rsTemp As Recordset
    Dim szSql As String
    bSuccess = False
    frmTemp.Show vbModal
    If frmTemp.m_bLoginOk Then
        '验证用户及口令是否正确
        odb.ConnectionString = GetConnectionStr
        szSql = "SELECT * FROM user_info where operatorid = '" & frmTemp.m_szUserID & "' "
        odb.CursorLocation = adUseClient
        odb.Open
        Set rsTemp = odb.Execute(szSql)
        If rsTemp.RecordCount > 0 Then
            '如果查询出来有记录
            If FormatDbValue(rsTemp!user_password) = Trim(frmTemp.m_szPasword) Then
                '用户及口令都相同 ,验证通过
                bSuccess = True
                '赋银行号及用户号
                m_cszOperatorBankID = FormatLen(FormatDbValue(rsTemp!bank_id), cnLenOperatorBankID)
                m_cszOperatorID = FormatLen(FormatDbValue(rsTemp!operatorid), cnLenOperatorID)
                
            Else
                MsgBox "用户的密码不正确", vbOKOnly, "错误"
            End If
        Else
            MsgBox "该用户不存在", vbOKOnly, "错误"
        End If
        
    End If
    ShowLogin = bSuccess
End Function




Public Function GetConnectionStr(Optional ByVal pszWhich As String) As String
    Dim oReg As New CFreeReg
    Dim szDriverType As String
    Dim szIntegrated As String '是否集成帐户
    oReg.Init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany       'HKEY_LOCAL_MACHINE
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

Public Function GetUniqueTeam(prsTemp As Recordset, paszTemp() As String) As String()
'得到唯一的数组
Dim nCount As Integer
Dim i As Integer
Dim j As Integer
Dim nCount2 As Integer
    nCount = ArrayLength(paszTemp)
    For i = 1 To prsTemp.RecordCount
        nCount2 = nCount
        For j = 1 To nCount2
            If UCase(Trim(prsTemp!bank_id)) = UCase(Trim(paszTemp(j))) Then
                Exit For
            End If
        Next j
        If j > nCount2 Then
        '当此用户不存在时，则添加到数组中去
            nCount = nCount + 1
            ReDim Preserve paszTemp(1 To nCount)
            paszTemp(nCount) = Trim(prsTemp!bank_id)
        End If
        prsTemp.MoveNext
    Next i
    GetUniqueTeam = paszTemp
End Function
