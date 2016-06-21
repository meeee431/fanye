Attribute VB_Name = "mdlMain"
Option Explicit

Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Const MAX_LEN = 200 '�ַ�����󳤶�
Public g_oCommDialog As Object

Public Const cszPrimaryKey = "SellTk"
Public Const cszSubKey_ExtraSellType = "ExtraSellType"

Public m_aszCheckGateInfo() As String '��Ʊ��

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
    RT_SellTicket = 1 '��Ʊ
    RT_ExtraSellTicket = 2 '��Ʊ
    RT_ChangeTicket = 3 '��ǩ
    RT_ReturnTicket = 4 '��Ʊ
    RT_CancelTicket = 5 '��Ʊ
End Enum

'������״̬���ַ�������
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
'    ID_TotalSeat = 5
'    ID_BookCount = 6
'    ID_SeatCount = 7
'    ID_SeatTypeCount = 8
'    ID_BedTypeCount = 9
'    ID_AdditionalCount = 10
'    ID_VehicleModel = 11
    ID_TotalSeat = 5
    ID_SeatCount = 6
    ID_SeatTypeCount = 7
    ID_BedTypeCount = 8
    ID_AdditionalCount = 9
    ID_VehicleModel = 10
    ID_BookCount = 11
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
    ID_RealName = 35        '�Ƿ�ʵ����
End Enum
'�ݷŵ�Ʊ��ö��
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
    IT_RealName = 35 '�Ƿ�ʵ����
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
Public Const cszScrollBus = "����"
Public Const cszScrollBusTime = "֮ǰ"

Public Const cszMiddleTime = "11:30" '�����ʱ��


''�����ݼ�
'
'Public Const cnKeySetSeat = vbKeyF8
'Public Const cnKeyChangeSeatType = vbKeyF9


'*****************************************************
Public m_clSell As New Collection '��Ʊ���ڼ���
Public m_clChange As New Collection '��Ʊ���ڼ���
'Public m_clExtra As New Collection '��Ʊ���ڼ���
Public m_clReturn As New Collection '��Ʊ���ڼ���
Public m_clCancel As New Collection '��Ʊ���ڼ���




'*****************************************************

Public m_oAUser As ActiveUser
Public m_oSell As New SellTicketClient
Public m_oSellService As New SellTicketService
Public m_oParam As New SystemParam
Public m_bSellStationCanSellEachOther As Boolean
 
Public m_oCmdDlg As New STShell.CommDialog
Public m_oShell As New STShell.CommShell

Public m_lTicketNo As Long
Public m_lEndTicketNo As Long '����Ʊ��(fpd��ӣ�
Public m_lEndTicketNoOld As Long '����Ʊ��(fpd��ӣ�
Public m_szTicketPrefix As String

Private m_lTicketNoNumLen As Long
Private m_szTicketNoFromatStr As String

Public m_szCurrentUnitID As String '��ǰ�ṩƱ�����ĵ�λ
Public m_nCurrentTask As ETaskType  '��ǰ��Ʊ������

Public m_bSelfChangeUnitOrFun As Boolean
Public m_lStopBusColor As OLE_COLOR
Public m_lNormalBusColor As OLE_COLOR



Public m_aszFirstBus() As String
Public m_aszFirstStation() As String
'�趨ĳ��վ�㣿������ʾ�ڵ�һ��

Private m_szLastStatus As String '�����״̬���ڵ�״̬

'-----------------------
Public m_nCanSellDay  As Integer

Public g_nDiscountTicketInTicketTypePosition As Integer '�ۿ�Ʊ��Ʊ�����λ��
'-----------------------
Public m_bListNoSeatBus As Boolean      '�Ƿ��г������공��,2005-12-6 lyq׷��
Public m_bUseFastPrint As Boolean       '�Ƿ�ʹ�ÿ��ٴ�ӡ
Public m_ISellScreenShow As Integer      '�Ƿ������ʾ  2006-01-20 qlh
Public g_nBookTime As Long 'Ԥ���ͷ�ʱ��(��λ:����)
Public g_bIsBookValid As Boolean '�Ƿ�ʹ��Ԥ��ϵͳ

Public m_szSpecialTicketTypePosition As String '����Ʊ����Ʊ�ִ���
Public g_bIsUseInsurance As Boolean '�Ƿ�ʹ�ó�����ϵͳ

Public m_oNetSell As New NetSellTicketClient
Public Sub Main()
    Dim i As Integer
    Dim oSysMan As New User
    Dim szLoginCommandLine As String
    
    On Error GoTo Error_Handle
    
    If App.PrevInstance Then
        End
    End If
    If Not IsPrinterValid Then
        MsgBox "��ӡ��δ���ã�", vbInformation, "��ӡ������:"
        End
        Exit Sub
    End If
    szLoginCommandLine = TransferLoginParam(Trim(Command()))
    If szLoginCommandLine = "" Then
        Set m_oAUser = m_oShell.ShowLogin()
        
    Else
        Set m_oAUser = New ActiveUser
        m_oAUser.Login GetLoginParam(szLoginCommandLine, cszUserName), GetLoginParam(szLoginCommandLine, cszUserPassword), GetComputerName()
        m_oCmdDlg.Init m_oAUser
    End If
    If Not m_oAUser Is Nothing Then
        App.HelpFile = SetHTMLHelpStrings("PSTSellTK.chm") '�趨App.HelpFile
        m_oParam.Init m_oAUser
        m_oNetSell.Init m_oAUser
        Date = m_oParam.NowDate
        Time = m_oParam.NowDateTime
        
        '�趨ĳ��վ�㣿������ʾ�ڵ�һ��
        ReadSetFistData
        
        m_oSell.Init m_oAUser

        m_bSelfChangeUnitOrFun = False
        m_lStopBusColor = RGB(255, 0, 0)
        m_lNormalBusColor = RGB(0, 0, 0)
        'frmSplash.Show
        
        m_szCurrentUnitID = m_oParam.UnitID
        m_nCurrentTask = RT_SellTicket
        GetAppSetting
        
        
        
        
        SetHTMLHelpStrings "PSTSellTK.chm"
        
        oSysMan.Init m_oAUser
        oSysMan.Identify m_oAUser.UserID
        m_nCanSellDay = oSysMan.CanSellDay
        
       '����λ��Ʊվ֮���Ƿ����
        m_bSellStationCanSellEachOther = m_oParam.SellStationCanSellEachOther
'        If m_oAUser.SellStationID = "km" Then
'            m_bSellStationCanSellEachOther = True
'        Else
'            m_bSellStationCanSellEachOther = False
'        End If
        
        m_nPrintBusIDLen = m_oParam.PrintBusIDLen
        m_bPrintScrollBusMode = m_oParam.PrintScrollBusMode
        g_nDiscountTicketInTicketTypePosition = m_oParam.DiscountTicketInTicketTypePosition '�ۿ�Ʊ������λ��
        
        '2005-12-6 lyq ��������
        m_bUseFastPrint = IIf(Val(m_oParam.GetParam("WantDirectSheetPrint").szParamValue) = 1, True, False)
        m_bListNoSeatBus = IIf(Val(m_oParam.GetParam("WantListNoSeatBus").szParamValue) = 1, True, False)
        '2006-01-20 qlh �Ƿ������ʾ
        m_ISellScreenShow = Val(m_oParam.GetParam("AllowSellScreenShow").szParamValue)
        g_nBookTime = m_oParam.BookTime 'Ԥ���ͷ�ʱ��(��λ:����)
        g_bIsBookValid = m_oParam.IsBookValid
        
        m_szSpecialTicketTypePosition = m_oParam.SpecialTicketTypePosition '����Ʊ����Ʊ�ִ���
        
        GetIniFile
        
        m_szRegValue = GetRegInfo
        
        '*****************
        '������ʾ��
        g_lComPort = IIf(Val(ReadReg(cszComPort)) = 0, 1, Val(ReadReg(cszComPort))) 'IIf(Val(ReadReg(cszComPort)) = 2, 2, 1)
        
        SetInit
        '*****************
        
        '��ʼ����Ʊ��
        GetInitCheckGate
        
        frmChgStartTktNumber.Show vbModal
        
        If Not frmChgStartTktNumber.m_bOk Then Exit Sub
        
'        Dim szSystemPath As String '����ϵͳ·��
'        Dim sTmp As String * MAX_LEN '��Ž���Ĺ̶����ȵ��ַ���
'        Dim nLength As Long '�ַ�����ʵ�ʳ���
'        nLength = GetSystemDirectory(sTmp, MAX_LEN)
'        szSystemPath = Left(sTmp, nLength)
'        g_bIsUseInsurance = FileIsExist(szSystemPath & "\ST6InsuranceOperation.exe")
'        If g_bIsUseInsurance = False Then
'            If MsgBox("û�а�װ[��Է���˳����մ�ϵͳ]���Ƿ����������Ʊ̨��", vbYesNo + vbDefaultButton2, "��ʾ") = vbNo Then
'                Exit Sub
'            End If
'        Else
'            Set g_oCommDialog = Nothing
'            Set g_oCommDialog = CreateObject("ST6InsuranceOperation.CommDialog")
'        End If
        
        MDISellTicket.Show
        
'        Set oSysMan = Nothing
'        Set m_oSell = Nothing
'        Set m_oSellService = Nothing
'        Set m_oAUser = Nothing
'        Set m_oCmdDlg = Nothing
'        Set m_oParam = Nothing
    End If
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
    'm_lEndTicketNo = m_lEndTicketNo + pnOffset
    If Not pbNoShow Then
        MDISellTicket.lblTicketNo.Caption = GetTicketNo()
        MDISellTicket.lblEndTicketNo.Caption = GetEndTicketNo()
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
    If m_lTicketNoNumLen = 0 Then
        m_lTicketNoNumLen = m_oParam.TicketNumberLen
    End If
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
    Dim i As Integer, j As Integer
    Dim nCount As Integer, nTemp As Integer, nTicketPrefixLen As Integer
    'On Error Resume Next
    pszFullTicketNo = Trim(pszFullTicketNo)
    nCount = Len(pszFullTicketNo)
    
    For i = 1 To nCount
        nTemp = Asc(Mid(pszFullTicketNo, nCount - i + 1, 1))
        If nTemp < vbKey0 Or nTemp > vbKey9 Then
            Exit For
        End If
    Next
    i = i - 1
    If i > 0 Then
        nTemp = TicketNoNumLen()
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


Public Sub GetAppSetting()
'    Dim oReg As New CFreeReg
    Dim szLastTicketNo As String
    Dim szEndLastTicketNo As String
    szLastTicketNo = m_oSell.GetLastTicketNo(m_oAUser.UserID)
    szEndLastTicketNo = m_oSell.GetEndLastTicketNo(m_oAUser.UserID)
    m_lTicketNo = ResolveTicketNo(szLastTicketNo, m_szTicketPrefix)
    m_lEndTicketNo = ResolveTicketNo(szEndLastTicketNo, m_szTicketPrefix)
    IncTicketNo , True
'    Set oReg = Nothing
End Sub


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

'�õ�һ�Զ���(,)�ָ����ַ����е�����
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

'�õ�һ�Զ���(,)�ָ����ַ����е�ָ����ŵ���
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

'����Ʊ��Ϣ����Ʊ�����Ϣ�еõ���ӦƱ�۵���Ϣ
Public Function SelfGetTicketPriceFromSellResult(ByVal pnTicketType As ETicketType, pabiTemp() As TBuyTicketInfo, psrTemp As TSellTicketResult) As Double
    Dim nCount  As Integer, i As Integer
    Dim nTicketType As ETicketType
    
    nCount = ArrayLength(pabiTemp)
    For i = 1 To nCount
        If pabiTemp(i).nTicketType <> TP_HalfPrice Then
            nTicketType = TP_FullPrice
        Else
            nTicketType = TP_HalfPrice
        End If
        If nTicketType = pnTicketType Then
            SelfGetTicketPriceFromSellResult = psrTemp.asgTicketPrice(i)
            Exit For
        End If
    Next
End Function


'����ָ��LISTVIEW��ָ���е���ѡ��(��û���κ�һ��ѡȡ����Ϊȫ��)��ֵ�ܺ�
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
    If Shift And 2 <> 0 Then '���Alt������
        nIndex = KeyCode - vbKey1 + 1
        If nIndex >= 1 And nIndex <= MDISellTicket.tsUnit.Tabs.count Then
            MDISellTicket.tsUnit.Tabs(nIndex).Selected = True
        End If
    End If
End Function

'��վ����Ͽ��е��ַ����õ�վ������
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
    GetMenuUnitName = Left(pszUnitName1, InStr(1, pszUnitName1, "(", vbTextCompare) - 1)
End Function


Public Function GetTicketTypeStr2(ByVal pnTicketType As Integer) As String
Dim j As Integer
Dim TicketType() As TTicketType
Dim intEnableTicketNo As Integer

   TicketType = m_oSell.GetAllTicketType(1)
   intEnableTicketNo = UBound(TicketType) - LBound(TicketType) + 1
    For j = 1 To intEnableTicketNo
        If TicketType(j).nTicketTypeID = pnTicketType And TicketType(j).nTicketTypeValid = TP_TicketTypeValid Then
           GetTicketTypeStr2 = TicketType(j).szTicketTypeName
           Exit For
        End If
    Next j
End Function

'�õ���ӡ���δ���
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


'�õ��������η���ʱ���ӡ��ʽ
Public Function GetPrintScrollMode() As Boolean
    GetPrintScrollMode = m_bPrintScrollBusMode
End Function


'�жϳ����Ƿ������������͵ĳ���
Public Function IsSeatTypeBus(pdBusDate As Date, pdBusID As String, pSeatTypeBus As TMultiSeatTypeBus) As Boolean
Dim nlen As Integer
Dim i As Integer
nlen = 0
nlen = ArrayLength(pSeatTypeBus.adBusDate)
For i = 1 To nlen
    If pdBusDate = pSeatTypeBus.adBusDate(i) And pdBusID = pSeatTypeBus.aszBusID(i) Then
        IsSeatTypeBus = True
        Exit Function
    End If
Next i
IsSeatTypeBus = False
End Function

'�����鸳��ʼֵ
Public Sub SetArrayInit(aInitArray() As Variant, InitValue As Variant)
Dim nlen As Integer
Dim i As Integer
nlen = 0
nlen = ArrayLength(aInitArray)
For i = 1 To nlen
    aInitArray(i) = InitValue
Next i
End Sub

'�õ���λ��
Public Function GetSeatCount(szSeatNo As String, nTotalSeat As Integer) As Integer
Dim i As Integer
Dim nCount As Integer
Dim nlen As Integer
nlen = 0
nCount = 0
nlen = Len(szSeatNo)
For i = 1 To nlen
    If Mid(szSeatNo, i, 1) = "," Then nCount = nCount + 1
    
Next i
nCount = nCount + 1
If nCount <= nTotalSeat Then
    GetSeatCount = nCount
Else
    GetSeatCount = nTotalSeat
End If

End Function

'�õ���λ��
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
'ListView������
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
' *   Brief Description: дϵͳ״̬����Ϣ                           *
' *   Engineer: ½����                                              *
' *******************************************************************
Public Sub ShowSBInfo(Optional pszInfo As String = "", Optional peArea As EStatusBarArea = ESB_WorkingInfo)
'����ע��
'*************************************
'pnArea(״̬������,Ĭ��ΪӦ�ó���״̬��)
'pszInfo(��Ϣ����)
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
                If pszInfo <> "" Then pszInfo = "��¼ʱ��: " & pszInfo
                .abMenuTool.Bands("statusBar").Tools("pnLoginTime").Caption = pszInfo
        End Select
    End With
End Sub
' *******************************************************************
' *   Member Name: WriteProcessBar                                  *
' *   Brief Description: дϵͳ������״̬                           *
' *   Engineer: ½����                                              *
' *******************************************************************
Public Sub WriteProcessBar(Optional pbVisual As Boolean = True, Optional ByVal plCurrValue As Variant = 0, Optional ByVal plMaxValue As Variant = 0, Optional pszShowInfo As String = cszUnrepeatString)
'����ע��
'*************************************
'plCurrValue(��ǰ����ֵ)
'plMaxValue(������ֵ)
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
    Dim szTemp As String

    szTemp = m_oSell.SellUnitCode
    m_oSell.SellUnitCode = m_szCurrentUnitID
    
    m_aszCheckGateInfo = m_oSell.GetAllCheckGate()
    
    m_oSell.SellUnitCode = szTemp
    
End Sub

'�õ���Ʊ�����ƺʹ���
Public Function GetCheckName(pszCheckGateID As String) As String
    Dim i As Integer
    Dim szResult As String
    Dim nlen As Integer
    nlen = 0
    nlen = ArrayLength(m_aszCheckGateInfo)
    szResult = ""
    For i = 1 To nlen
        If Trim(m_aszCheckGateInfo(i, 1)) = Trim(pszCheckGateID) Then
            szResult = Trim(m_aszCheckGateInfo(i, 2))
            Exit For
        End If
    Next i
    GetCheckName = szResult

End Function

'������ת����ũ������������
Public Function GetChinaDate(pszDate As String) As String

    Dim WeekName(7), MonthAdd(11), NongliData(99), TianGan(9), DiZhi(11), ShuXiang(11), DayName(30), MonName(12)
    Dim curYear, curMonth, curDay, curWeekday
    Dim GongliStr, WeekdayStr, NongliStr, NongliDayStr
    Dim i, m, n, k, isEnd, bit, TheDate

    ''������
    WeekName(0) = " * "
    WeekName(1) = "������"
    WeekName(2) = "����һ"
    WeekName(3) = "���ڶ�"
    WeekName(4) = "������"
    WeekName(5) = "������"
    WeekName(6) = "������"
    WeekName(7) = "������"
    
    ''�������
    TianGan(0) = "��"
    TianGan(1) = "��"
    TianGan(2) = "��"
    TianGan(3) = "��"
    TianGan(4) = "��"
    TianGan(5) = "��"
    TianGan(6) = "��"
    TianGan(7) = "��"
    TianGan(8) = "��"
    TianGan(9) = "��"
    
    ''��֧����
    DiZhi(0) = "��"
    DiZhi(1) = "��"
    DiZhi(2) = "��"
    DiZhi(3) = "î"
    DiZhi(4) = "��"
    DiZhi(5) = "��"
    DiZhi(6) = "��"
    DiZhi(7) = "δ"
    DiZhi(8) = "��"
    DiZhi(9) = "��"
    DiZhi(10) = "��"
    DiZhi(11) = "��"
    
    ''��������
    ShuXiang(0) = "��"
    ShuXiang(1) = "ţ"
    ShuXiang(2) = "��"
    ShuXiang(3) = "��"
    ShuXiang(4) = "��"
    ShuXiang(5) = "��"
    ShuXiang(6) = "��"
    ShuXiang(7) = "��"
    ShuXiang(8) = "��"
    ShuXiang(9) = "��"
    ShuXiang(10) = "��"
    ShuXiang(11) = "��"
    
    ''ũ��������
    DayName(0) = "*"
    DayName(1) = "��һ"
    DayName(2) = "����"
    DayName(3) = "����"
    DayName(4) = "����"
    DayName(5) = "����"
    DayName(6) = "����"
    DayName(7) = "����"
    DayName(8) = "����"
    DayName(9) = "����"
    DayName(10) = "��ʮ"
    DayName(11) = "ʮһ"
    DayName(12) = "ʮ��"
    DayName(13) = "ʮ��"
    DayName(14) = "ʮ��"
    DayName(15) = "ʮ��"
    DayName(16) = "ʮ��"
    DayName(17) = "ʮ��"
    DayName(18) = "ʮ��"
    DayName(19) = "ʮ��"
    DayName(20) = "��ʮ"
    DayName(21) = "إһ"
    DayName(22) = "إ��"
    DayName(23) = "إ��"
    DayName(24) = "إ��"
    DayName(25) = "إ��"
    DayName(26) = "إ��"
    DayName(27) = "إ��"
    DayName(28) = "إ��"
    DayName(29) = "إ��"
    DayName(30) = "��ʮ"
    
    ''ũ���·���
    MonName(0) = "*"
    MonName(1) = "��"
    MonName(2) = "��"
    MonName(3) = "��"
    MonName(4) = "��"
    MonName(5) = "��"
    MonName(6) = "��"
    MonName(7) = "��"
    MonName(8) = "��"
    MonName(9) = "��"
    MonName(10) = "ʮ"
    MonName(11) = "ʮһ"
    MonName(12) = "��"
    
    ''����ÿ��ǰ�������
    MonthAdd(0) = 0
    MonthAdd(1) = 31
    MonthAdd(2) = 59
    MonthAdd(3) = 90
    MonthAdd(4) = 120
    MonthAdd(5) = 151
    MonthAdd(6) = 181
    MonthAdd(7) = 212
    MonthAdd(8) = 243
    MonthAdd(9) = 273
    MonthAdd(10) = 304
    MonthAdd(11) = 334
    
    ''ũ������
    NongliData(0) = 2635
    NongliData(1) = 333387
    NongliData(2) = 1701
    NongliData(3) = 1748
    NongliData(4) = 267701
    NongliData(5) = 694
    NongliData(6) = 2391
    NongliData(7) = 133423
    NongliData(8) = 1175
    NongliData(9) = 396438
    NongliData(10) = 3402
    NongliData(11) = 3749
    NongliData(12) = 331177
    NongliData(13) = 1453
    NongliData(14) = 694
    NongliData(15) = 201326
    NongliData(16) = 2350
    NongliData(17) = 465197
    NongliData(18) = 3221
    NongliData(19) = 3402
    NongliData(20) = 400202
    NongliData(21) = 2901
    NongliData(22) = 1386
    NongliData(23) = 267611
    NongliData(24) = 605
    NongliData(25) = 2349
    NongliData(26) = 137515
    NongliData(27) = 2709
    NongliData(28) = 464533
    NongliData(29) = 1738
    NongliData(30) = 2901
    NongliData(31) = 330421
    NongliData(32) = 1242
    NongliData(33) = 2651
    NongliData(34) = 199255
    NongliData(35) = 1323
    NongliData(36) = 529706
    NongliData(37) = 3733
    NongliData(38) = 1706
    NongliData(39) = 398762
    NongliData(40) = 2741
    NongliData(41) = 1206
    NongliData(42) = 267438
    NongliData(43) = 2647
    NongliData(44) = 1318
    NongliData(45) = 204070
    NongliData(46) = 3477
    NongliData(47) = 461653
    NongliData(48) = 1386
    NongliData(49) = 2413
    NongliData(50) = 330077
    NongliData(51) = 1197
    NongliData(52) = 2637
    NongliData(53) = 268877
    NongliData(54) = 3365
    NongliData(55) = 531109
    NongliData(56) = 2900
    NongliData(57) = 2922
    NongliData(58) = 398042
    NongliData(59) = 2395
    NongliData(60) = 1179
    NongliData(61) = 267415
    NongliData(62) = 2635
    NongliData(63) = 661067
    NongliData(64) = 1701
    NongliData(65) = 1748
    NongliData(66) = 398772
    NongliData(67) = 2742
    NongliData(68) = 2391
    NongliData(69) = 330031
    NongliData(70) = 1175
    NongliData(71) = 1611
    NongliData(72) = 200010
    NongliData(73) = 3749
    NongliData(74) = 527717
    NongliData(75) = 1452
    NongliData(76) = 2742
    NongliData(77) = 332397
    NongliData(78) = 2350
    NongliData(79) = 3222
    NongliData(80) = 268949
    NongliData(81) = 3402
    NongliData(82) = 3493
    NongliData(83) = 133973
    NongliData(84) = 1386
    NongliData(85) = 464219
    NongliData(86) = 605
    NongliData(87) = 2349
    NongliData(88) = 334123
    NongliData(89) = 2709
    NongliData(90) = 2890
    NongliData(91) = 267946
    NongliData(92) = 2773
    NongliData(93) = 592565
    NongliData(94) = 1210
    NongliData(95) = 2651
    NongliData(96) = 395863
    NongliData(97) = 1323
    NongliData(98) = 2707
    NongliData(99) = 265877
    
    
    ''���ɵ�ǰ�����ꡢ�¡��� ==> GongliStr
    curYear = Year(pszDate)
    curMonth = Month(pszDate)
    curDay = Day(pszDate)

    GongliStr = curYear & "��"
    If (curMonth < 10) Then
        GongliStr = GongliStr & "0" & curMonth & "��"
    Else
        GongliStr = GongliStr & curMonth & "��"
    End If
    If (curDay < 10) Then
        GongliStr = GongliStr & "0" & curDay & "��"
    Else
        GongliStr = GongliStr & curDay & "��"
    End If

    ''���ɵ�ǰ�������� ==> WeekdayStr
    curWeekday = Weekday(pszDate)
    WeekdayStr = WeekName(curWeekday)

    ''���㵽��ʼʱ��1921��2��8�յ�������1921-2-8(���³�һ)
    TheDate = (curYear - 1921) * 365 + Int((curYear - 1921) / 4) + curDay + MonthAdd(curMonth - 1) - 38
    If ((curYear Mod 4) = 0 And curMonth > 2) Then
        TheDate = TheDate + 1
    End If
    
    ''����ũ����ɡ���֧���¡���
    isEnd = 0
    m = 0
    
    Do
    If (NongliData(m) < 4095) Then
        k = 11
    Else
        k = 12
    End If
    
    n = k
    Do
    If (n < 0) Then
    Exit Do
    End If

    ''��ȡNongliData(m)�ĵ�n��������λ��ֵ
    bit = NongliData(m)
    For i = 1 To n Step 1
        bit = Int(bit / 2)
    Next
    bit = bit Mod 2
    
    If (TheDate <= 29 + bit) Then
        isEnd = 1
        Exit Do
    End If
    
    TheDate = TheDate - 29 - bit
    
    n = n - 1
    Loop
    
    If (isEnd = 1) Then
        Exit Do
    End If
    
    m = m + 1
    Loop

    curYear = 1921 + m
    curMonth = k - n + 1
    curDay = TheDate
    
    If (k = 12) Then
        If (curMonth = (Int(NongliData(m) / 65536) + 1)) Then
            curMonth = 1 - curMonth
        ElseIf (curMonth > (Int(NongliData(m) / 65536) + 1)) Then
            curMonth = curMonth - 1
        End If
    End If
    
    ''����ũ����ɡ���֧������ ==> NongliStr
    NongliStr = TianGan(((curYear - 4) Mod 60) Mod 10) & DiZhi(((curYear - 4) Mod 60) Mod 12) & "��"
    NongliStr = NongliStr & "(" & ShuXiang(((curYear - 4) Mod 60) Mod 12) & ")"
    
    ''����ũ���¡��� ==> NongliDayStr
    If (curMonth < 1) Then
        NongliDayStr = "��" & MonName(-1 * curMonth)
    Else
        NongliDayStr = MonName(curMonth)
    End If
    NongliDayStr = NongliDayStr & "��"
    
    NongliDayStr = NongliDayStr & DayName(curDay)
    
    '����ũ��������
    GetChinaDate = MakeDisplayString(NongliStr & NongliDayStr, WeekdayStr)

End Function
