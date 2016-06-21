Attribute VB_Name = "mdlMain"
Option Explicit

Public Const m_cRegParamKey = "DataBaseSet"

Public Const cszLuggageAccount = "LugAcc"
'====================================================================
'以下定义枚举


'主界面状态条字符串区域
Public Enum EStatusBarArea
    ESB_WorkingInfo = 1
    ESB_ResultCountInfo = 2
    ESB_UserInfo = 3
    ESB_LoginTime = 4
End Enum
'对话框调用状态
Public Enum EFormStatus
    EFS_AddNew = 0
    EFS_Modify = 1
    EFS_Show = 2
    EFS_Delete = 3
End Enum
Private Declare Function GetComputerNameAPI Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'====================================================================
'以下全局变量定义
'--------------------------------------------------------------------
Public m_szLuggageNo As String '受理单号
Public m_szLuggagePrefix As String '受理单号前缀
Public m_szTicketNoFromatStr As String
Public m_nLuggageNoNumLen As Integer '受理单号数字长度

Public g_szAcceptSheetID   As String      '当前受理单号
Public g_szCarrySheetID  As String        '当前签发单号

Public moAcceptSheet As New AcceptSheet
Public moCarrySheet As New CarrySheet
Public moSysParam As New STLuggage.LuggageParam
Public moLugSvr As New LuggageSvr
Public m_oParam As New SystemParam
Public m_oBase As New BaseInfo
Public m_bIsRelationWithVehicleType As Boolean '行包的公式是否与车型有关系

Public m_bIsDispSettlePriceInAccept As Boolean '是否在受理时显示应结费用
Public m_bIsDispSettlePriceInCheck As Boolean '是否在签发时显示应结费用
Public m_bIsSettlePriceFromAcceptInCheck As Boolean '结算运费是不是从受理的结算运费中汇总得到
Public m_bIsPrintCheckSheet As Boolean '是否打印签发单


Public m_szCustom As String '行包打印的自定义信息


Public Const clActiveColor = &HFF0000
Public m_nCanSellDay  As Integer
Public m_oShell As New STShell.CommShell
Public m_oCmdDlg As New STShell.CommDialog
Public m_oAUser As ActiveUser

Public Const szAcceptTypeGeneral = "快件"
Public Const szAcceptTypeMan = "普通"
Public Const szPickTypeGeneral = "自提"
Public Const szPickTypeEms = "送货"
Public Const szLuggageQucke = "快件"

Public Const szAcceptStatus = 0 '"正常"

Public m_oPrintTicket As FastPrint ' BPrint

Public m_oPrintCarrySheet As BPrint

Public g_rsPriceItem As Recordset   '运费项记录


Public g_szOurCompany As String
'以下定义套打枚举
'正常受理单套打

Public Enum PrintAcceptObjectIndexEnum
    PAI_LabelNo = 1 '标签号
    PAI_Shipper = 3 '托运人
    PAI_CalWeight = 4 '计重
    PAI_Picker = 5 '收件人
    PAI_Pack = 6 '包装
    
    PAI_StartStation = 7 '始发站
    PAI_TransType = 8  '托运方式
    
    PAI_LongLuggageID = 10 '长票号
    
    PAI_UserName = 12 '开票人
    PAI_LuggageName = 13 '货名
    
    PAI_EndStation = 15 '到站名
    PAI_LuggageNumber = 16 '件数
    PAI_TotalPriceBig = 17 '合计大写
    PAI_TotalPrice = 18 '合计（小写）
    PAI_AcceptDate = 19 ' 受理时间
    PAI_Vehicle = 20 '车牌号
    PAI_OperationDate = 21 '操作日期
    PAI_ShipperPhone = 22 '托运人电话
    PAI_PickerPhone = 23 '收件人电话
    PAI_PickerAddress = 24 '收件人地址
    PAI_ActWeight = 25 '实重
    PAI_OverNumber = 26    '超重件数
    PAI_BusID = 27  '车次
    PAI_StartTime = 28 '发车时间
    PAI_LicenseTagNo = 29 '车牌号
    PAI_BusDate = 30  '车次日期
    PAI_TotalPrice2 = 31 '合计小写2
    PAI_TotalPriceName = 32 '合计小写名称
    
    PAI_Year = 33  '年
    PAI_Month = 34 '月
    PAI_Day = 35 '日
    
    PAI_BusID2 = 36  '车次2
    PAI_Vehicle2 = 37 '车牌号2
    PAI_StartTime2 = 38  '发车时间2
    
    '扩展支持部分
    PAI_PriceItem1 = 40  '票价1（运费）
    PAI_PriceItem2 = 41  '票价2（服务费）
    PAI_PriceItem3 = 42  '票价3（上门接送费）
    PAI_PriceItem4 = 43  '
    PAI_PriceItem5 = 44
    PAI_PriceItem6 = 45
    PAI_PriceItem7 = 46
    PAI_PriceItem8 = 47
    PAI_PriceItem9 = 48
    PAI_PriceItem10 = 49
    
    
    '大写金额位数
    PAI_Cent = 51
    PAI_Jiao = 52
    PAI_Yuan = 53
    PAI_Ten = 54
    PAI_Hundred = 55
    PAI_Thousand = 56
    
    
    
    PAI_StartStation2 = 60 '始发站2
    PAI_EndStation2 = 61 '到站名2
    PAI_StartStation3 = 63 '始发站3
    PAI_EndStation3 = 64 '到站名3
    PAI_StartStation4 = 65 '始发站4
    PAI_EndStation4 = 66 '到站名4
    
    PAI_TransTicketID1 = 71    '运输单号1
    PAI_TransTicketID2 = 72    '运输单号2
    PAI_TransTicketID3 = 73    '运输单号3
    PAI_TransTicketID4 = 74    '运输单号4
    
    
    PAI_InsuranceID1 = 75    '保险单号1
    PAI_Annotation1 = 76    '备注1
    PAI_Annotation2 = 77    '备注2
    
    PAI_LuggageNumber2 = 80 '件数2
    PAI_LuggageNumber3 = 81 '件数3
    PAI_LuggageNumber4 = 82 '件数4
    PAI_LuggageName2 = 83 '货名2
    PAI_LuggageName3 = 84 '货名3
    PAI_LuggageName4 = 85 '货名4
    
    
    PAI_Year2 = 90  '年
    PAI_Month2 = 91 '月
    PAI_Day2 = 92 '日
    PAI_Year3 = 93  '年
    PAI_Month3 = 94 '月
    PAI_Day3 = 95 '日
    PAI_Year4 = 96  '年
    PAI_Month4 = 97 '月
    PAI_Day4 = 98 '日
    
    
    PAI_SettlePrice = 99 '应结运费
    
    PAI_BasePrice = 100 '行包运费2
    PAI_BasePriceName = 101 '行包运费名称
    PAI_SettlePriceName = 102 '行包应结运费名称
    
    
    PAI_Custom1 = 110    '自定义信息1
    PAI_Custom2 = 111    '自定义信息2
    
    
    PAI_LongLuggageID2 = 112 '长票号
    PAI_LongLuggageID3 = 113 '长票号
    PAI_LongLuggageID4 = 114 '长票号
    
    PAI_ActWeight2 = 115 '实重2
    PAI_CalWeight2 = 116 '计重2
    PAI_OperationDate2 = 117 '开票日期2
    PAI_TotalPriceBig2 = 118 '合计大写2
    PAI_UserName2 = 119 '开票人2
    PAI_LabelNo2 = 120 '标签2
    PAI_TransType2 = 121 '托运方式
    
    PAI_Mark = 122 '受理标记
    PAI_Mark2 = 123 '受理标记2
    
    
    
End Enum
'受理单退运套打
Public Enum PrintReturnAcceptObjectIndexEnum
    PRI_LongLuggageID = 10 '长票号
    PRI_CredenceID = 9 '退票凭证号
    PRI_TransType = 8  '托运方式
    PRI_StartStation = 7 '始发站
    PRI_EndStation = 15 '到站名
    PRI_LuggageNumber = 16 '件数
    PRI_Shipper = 3 '托运人
    PRI_CalWeight = 4 '计重
    PRI_Picker = 5 '收件人
    PRI_LuggageName = 13 '货名
    PRI_LabelNo = 1 '标签号
    PRI_TotalPrice = 17 '合计（小写）
    PRI_TotalPriceBig = 18 '合计大写
    PRI_ReturnCharge = 19 '退运费
    PRI_ReturnChargeBig = 20 '退运费大写
    PRI_UserName = 12 '开票人
    
    PRI_OperationDate = 21 '操作日期
    PRI_OverNumber = 26    '超重件数
    
    PRI_ReturnDate = 27 '退运时间
    PRI_ReturnChargeName = 28 '退运手续费名称
End Enum
'签发单套打


Const cnDetailItemCount = 4    '详细项目数量

Public Enum PrintCarrySheetObjectIndexEnum
    '行包受理清单联
    PCI_SheetID = 1         '签发单号
    PCI_TransType = 2  '托运方式
    PCI_StartStation = 3 '始发站
    PCI_Year = 4
    PCI_Month = 5
    PCI_Day = 6
    PCI_AcceptDetailItem = 7     '受理清单
    PCI_UserID = 8
    
    
    '行包结算联
    PCI_Year2 = 9
    PCI_Month2 = 10
    PCI_Day2 = 11
    PCI_SheetID2 = 12 '签发单号
    PCI_EndStation = 13 '终点站
    PCI_StartTime = 14 '发车时间
    PCI_LicenseTagNo = 15 '车牌号
    PCI_TotalPrice = 16 '总运费 合计（小写）
    PCI_TotalPriceBig = 17 '合计大写
    PCI_UserID2 = 18
    PCI_Number = 19 '件数
    
    
    '装车清单联
    PCI_Year3 = 20
    PCI_Month3 = 21
    PCI_Day3 = 22
    PCI_SheetID3 = 23         '签发单号
    PCI_UserID3 = 24
    PCI_LicenseTagNo2 = 25 '车牌号
    PCI_StartTime2 = 26 '发车时间
    PCI_CarryDetailItem = 27     '装车清单
    
    
    PCI_ShipperPhone = 28     '发件人电话清单
    
    PCI_CarryTime = 29       '签发时间
    PCI_CarryTime2 = 30     '签发时间2
    PCI_CarryTime3 = 31     '签发时间3
    PCI_TotalPrice2 = 32    '总运费2 合计（小写)
    
    PCI_MoveWorker = 33 '新加打印“装卸工”,fpd
End Enum

Public Function ToStandardDateStr(pdtDate As Date) As String
    ToStandardDateStr = Format(pdtDate, "YYYY-MM-DD")
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
    With mdiMain
        Select Case peArea
        Case EStatusBarArea.ESB_WorkingInfo
            .abMenu.Bands("statusBar").Tools("pnWorkingInfo").Caption = pszInfo
        Case EStatusBarArea.ESB_ResultCountInfo
            .abMenu.Bands("statusBar").Tools("pnResultCountInfo").Caption = pszInfo
        Case EStatusBarArea.ESB_UserInfo
            .abMenu.Bands("statusBar").Tools("progressBar").Visible = False
            .abMenu.Bands("statusBar").Tools("pnUserInfo").Caption = pszInfo
        Case EStatusBarArea.ESB_LoginTime
            If pszInfo <> "" Then pszInfo = "登录时间: " & pszInfo
            .abMenu.Bands("statusBar").Tools("pnLoginTime").Caption = pszInfo
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
    If pszShowInfo <> cszUnrepeatString Then ShowSBInfo pszShowInfo, ESB_WorkingInfo
    If plMaxValue = 0 And pbVisual = True Then Exit Sub
    Dim nCurrProcess As Integer
    With mdiMain.abMenu.Bands("statusBar")
        If pbVisual Then
            If Not .Tools("progressBar").Visible Then
                .Tools("progressBar").Visible = True
                .Tools("pnResultCountInfo").Caption = ""
                .Tools("pnResultCountInfo").Visible = False
                mdiMain.pbLoad.Max = 100
                mdiMain.abMenu.RecalcLayout
            End If
            nCurrProcess = Int(plCurrValue / plMaxValue * 100)
            mdiMain.pbLoad.Value = nCurrProcess
        Else
            .Tools("progressBar").Visible = False
            .Tools("pnResultCountInfo").Visible = True
        End If
    End With
End Sub

Public Function GetComputerName() As String
    ' Set or retrieve the name of the computer.
    Dim strBuffer As String
    Dim lngLen As Long
        
    strBuffer = Space(255 + 1)
    lngLen = Len(strBuffer)
    If CBool(GetComputerNameAPI(strBuffer, lngLen)) Then
        GetComputerName = MidA(strBuffer, 1, lngLen)
    Else
        GetComputerName = ""
    End If
End Function

Sub Main()
 On Error GoTo Error_Handle
    Dim szLoginCommandLine As String
    Dim oLugSysParam As New STLuggage.LuggageParam
'     Dim oSysMan As New User
'    判断打印是否存在，不存在则出错
'    If App.PrevInstance Then
'        End
'    End If
    If Not IsPrinterValid Then
        MsgBox "打印机未配置！", vbInformation, "打印机出错:"
        End
        Exit Sub
    End If
    '登录
    szLoginCommandLine = TransferLoginParam(Trim(Command()))
    If szLoginCommandLine = "" Then
        Set m_oAUser = m_oShell.ShowLogin()
        
    Else
        Set m_oAUser = New ActiveUser
        m_oAUser.Login GetLoginParam(szLoginCommandLine, cszUserName), GetLoginParam(szLoginCommandLine, cszUserPassword), GetComputerName()
        m_oCmdDlg.Init m_oAUser
    End If
    If Not m_oAUser Is Nothing Then
'        App.HelpFile = SetHTMLHelpStrings("SNSellTK.chm") '设定App.HelpFile
        
        m_oParam.Init m_oAUser
        Date = m_oParam.NowDate
        Time = m_oParam.NowDateTime
        m_bIsRelationWithVehicleType = False 'm_oParam.IsRelationWithVehicleType

        m_bIsDispSettlePriceInAccept = False 'm_oParam.IsDispSettlePriceInAccept
        m_bIsDispSettlePriceInCheck = False 'm_oParam.IsDispSettlePriceInCheck
        m_bIsSettlePriceFromAcceptInCheck = False 'm_oParam.IsSettlePriceFromAcceptInCheck
        m_bIsPrintCheckSheet = False 'm_oParam.IsPrintCheckSheet
        
        SetHTMLHelpStrings "pstLugDesk.chm"
        
        
        
        '读出注册表里的注意事项
        '行包注意事项
        Dim oFreeReg As CFreeReg
        Set oFreeReg = New CFreeReg
        oFreeReg.Init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany
        m_szCustom = oFreeReg.GetSetting(cszLuggageAccount, "CareContent") '自定义信息
        
        
        
        
        oLugSysParam.Init m_oAUser
        Set g_rsPriceItem = oLugSysParam.GetPriceItemRS(0)
        

        moAcceptSheet.Init m_oAUser
        moCarrySheet.Init m_oAUser
        moSysParam.Init m_oAUser
        moLugSvr.Init m_oAUser
        'frmSplash.Show

        Set m_oPrintTicket = New FastPrint
        Set m_oPrintCarrySheet = New BPrint
        '初始化打印签发单
    
        m_oPrintTicket.ReadFormatFile App.Path & "\AcceptSheet.ini"
'        m_oPrintTicket.ClearAll
        m_oPrintCarrySheet.ReadFormatFileA App.Path & "\CarrySheet.bpf"
        
        g_szOurCompany = oLugSysParam.OurCompany

    '打开frmChgSheetNo窗体，设置起始的行包单号和签发单号， 放入g_szAcceptSheetID和g_szCarrySheetID
    
    '得到起始的行包单号和签发单号
        GetAppSetting2
        
        frmChgSheetNo.Show vbModal
        If Not frmChgSheetNo.m_bOk Then
            End
        End If
        Load mdiMain
        DoEvents
    '    显示splash窗体
        
        mdiMain.Show
        '关闭splash窗体
        

    End If
  Exit Sub
Error_Handle:
ShowErrorMsg
End Sub

'设置主界面上的标签号
Public Sub SetSheetNoLabel(pbIsAccept As Boolean, pszSheetNo As String)
    'pbIsAccept 是否是受理单号
    If pbIsAccept Then
        mdiMain.lblSheetNoName = "当前受理单号:"
    Else
        mdiMain.lblSheetNoName = "当前签发单号:"
    End If
    mdiMain.lblSheetNoName.Visible = True
    mdiMain.lblSheetNo.Visible = True
    mdiMain.lblSheetNo.Caption = FormatSheetID(pszSheetNo)
End Sub
'隐藏主界面上的标签号
Public Sub HideSheetNoLabel()
    mdiMain.lblSheetNoName.Visible = False
    mdiMain.lblSheetNo.Visible = False
End Sub







'起始受理单和签发单的处理

Public Function TicketNoNumLen() As Integer
    If m_nLuggageNoNumLen = 0 Then
        m_nLuggageNoNumLen = moSysParam.LuggageIDNumberLen
    End If
    TicketNoNumLen = m_nLuggageNoNumLen
End Function


Public Function GetTicketNo(Optional pnOffset As Integer = 0) As String
    GetTicketNo = MakeTicketNo(m_szLuggageNo + pnOffset, m_szLuggagePrefix)
End Function

'签发单号自增
Public Sub IncSheetID(Optional pnOffset As Integer = 1, Optional pbNoShow As Boolean = False)
    g_szCarrySheetID = g_szCarrySheetID + pnOffset
    If Not pbNoShow Then
        mdiMain.lblSheetNoName.Caption = "当前签发单号:"
        mdiMain.lblSheetNo.Caption = FormatSheetID(g_szCarrySheetID)
    End If
End Sub


'受理单号自增
Public Sub IncTicketNo(Optional pnOffset As Integer = 1, Optional pbNoShow As Boolean = False)
    Dim oReg As New CFreeReg
    Dim szConnectName As String
    
    szConnectName = "Luggage"
    
    On Error GoTo ErrHandle
    
    oReg.Init "RTStation", HKEY_LOCAL_MACHINE, "Software\RTSoft"
    m_szLuggageNo = m_szLuggageNo + pnOffset
    If Not pbNoShow Then
        mdiMain.lblSheetNoName.Caption = "当前受理单号:"
        mdiMain.lblSheetNo.Caption = GetTicketNo()
'        g_szAcceptSheetID = GetTicketNo
    End If
    
    g_szAcceptSheetID = GetTicketNo
    oReg.SaveSetting szConnectName, "CurrentSheetID", g_szAcceptSheetID
    
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

'生成票号
Public Function MakeTicketNo(plTicketNo As Long, Optional pszPrefix As String = "") As String
   
    MakeTicketNo = pszPrefix & Format(plTicketNo, TicketNoFormatStr())
End Function

'票号字符串
Private Function TicketNoFormatStr() As String
    Dim i As Integer
    If m_szTicketNoFromatStr = "" Then
        m_szTicketNoFromatStr = String(TicketNoNumLen(), "0")
    End If
    TicketNoFormatStr = m_szTicketNoFromatStr
End Function

'将票号分解成前缀部分与数字部分
Public Function ResolveTicketNo(pszFullTicketNo, ByRef pszTicketPrefix As String) As Long
    Dim i As Integer, j As Integer
    Dim nCount As Integer, nTemp As Integer, nTicketPrefixLen As Integer
    
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
        
        nTicketPrefixLen = m_oParam.LuggageIDPrefixLen
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

Public Function FormatSheetID(pszCheckID As String)
    FormatSheetID = Format(IIf(pszCheckID <> "", pszCheckID, 0), String(moSysParam.CarrySheetIDNumberLen, "0"))
End Function

Public Sub GetAppSetting()
    Dim szLastTicketNo As String
    
    Dim szLastSheetID As String
    On Error GoTo here
    szLastSheetID = moLugSvr.GetLastSheetID(m_oAUser.UserID)
    szLastTicketNo = moLugSvr.GetLastLuggageID(m_oAUser.UserID)
    
    g_szCarrySheetID = FormatSheetID(szLastSheetID)
    
    m_szLuggageNo = ResolveTicketNo(szLastTicketNo, m_szLuggagePrefix)
    
    IncTicketNo , True
    IncSheetID , True
    
    Exit Sub
here:
    ShowErrorMsg
    
    
End Sub

'楚门改的读注册表里的单号
Public Sub GetAppSetting2()
    Dim szLastTicketNo As String
    Dim szLastSheetID As String
    Dim oReg As New CFreeReg
    Dim szConnectName As String

    
    oReg.Init "RTStation", HKEY_LOCAL_MACHINE, "Software\RTSoft"
    
    szConnectName = "Luggage"
    
    On Error GoTo here
    szLastSheetID = moLugSvr.GetLastSheetID(m_oAUser.UserID)
    szLastTicketNo = oReg.GetSetting(szConnectName, "CurrentSheetID", "")
    
    g_szCarrySheetID = FormatSheetID(szLastSheetID)
    
    m_szLuggageNo = ResolveTicketNo(IIf(Val(szLastTicketNo) - 1 >= 0, Val(szLastTicketNo) - 1, 0), m_szLuggagePrefix)
    m_szLuggageNo = GetTicketNo
    
    IncTicketNo , True
    IncSheetID , True
    
    Exit Sub
here:
    ShowErrorMsg
End Sub


'刷新界面上的单据号
Public Sub RefreshCurrentSheetID()
    Dim szLastTicketNo As String
    Dim oReg As New CFreeReg
    Dim szConnectName As String
    
    
    oReg.Init "RTStation", HKEY_LOCAL_MACHINE, "Software\RTSoft"
    
    szConnectName = "Luggage"
    
    On Error GoTo here
    szLastTicketNo = oReg.GetSetting(szConnectName, "CurrentSheetID", "")
    
    m_szLuggageNo = szLastTicketNo
    
    mdiMain.lblSheetNoName.Caption = "当前受理单号:"
    mdiMain.lblSheetNo.Caption = GetTicketNo()
    g_szAcceptSheetID = GetTicketNo
    
    Exit Sub
here:
    ShowErrorMsg
End Sub

'打印行包受理单
Public Sub PrintAcceptSheet(poAcceptSheet As AcceptSheet, pdtBusStartTime As Date)
           
#If PRINT_SHEET <> 0 Then

''    m_oPrintTicket.ClearAll
''    m_oPrintTicket.ReadFormatFileA App.Path & "\AcceptSheet.bpf"
'
'    '受理时间
'    m_oPrintTicket.SetCurrentObject PAI_AcceptDate
'    m_oPrintTicket.LabelSetCaption Format(Date, "yyyy年MM月dd日")
'
'    '车牌号
'    m_oPrintTicket.SetCurrentObject PAI_Vehicle
'    m_oPrintTicket.LabelSetCaption Trim(frmAccept.cboVehicle.Text)
'
'    '托运方式
'    m_oPrintTicket.SetCurrentObject PAI_TransType
'    m_oPrintTicket.LabelSetCaption poAcceptSheet.AcceptType
'
'    '长票号
'    m_oPrintTicket.SetCurrentObject PAI_LongLuggageID
'    m_oPrintTicket.LabelSetCaption poAcceptSheet.SheetID
'
'    '始发站
'    m_oPrintTicket.SetCurrentObject PAI_StartStation
'    m_oPrintTicket.LabelSetCaption poAcceptSheet.StartStationName
'
'    '到站名
'    m_oPrintTicket.SetCurrentObject PAI_EndStation
'    m_oPrintTicket.LabelSetCaption poAcceptSheet.DesStationName
'
'    '托运人
'    m_oPrintTicket.SetCurrentObject PAI_Shipper
'    m_oPrintTicket.LabelSetCaption poAcceptSheet.Shipper
'
'    '收件人
'    m_oPrintTicket.SetCurrentObject PAI_Picker
'    m_oPrintTicket.LabelSetCaption poAcceptSheet.Picker
'
'    '包装
'    m_oPrintTicket.SetCurrentObject PAI_Pack
'    m_oPrintTicket.LabelSetCaption poAcceptSheet.Pack
'
'
'    '货名
'    m_oPrintTicket.SetCurrentObject PAI_LuggageName
'    m_oPrintTicket.LabelSetCaption Trim(poAcceptSheet.LuggageName)
'
'    '标签号
'    m_oPrintTicket.SetCurrentObject PAI_LabelNo
'    m_oPrintTicket.LabelSetCaption Trim(IIf(poAcceptSheet.StartLabelID = poAcceptSheet.EndLabelID, poAcceptSheet.StartLabelID, poAcceptSheet.StartLabelID & "-" & poAcceptSheet.EndLabelID))
'
'    '件数
'    m_oPrintTicket.SetCurrentObject PAI_LuggageNumber
'    m_oPrintTicket.LabelSetCaption poAcceptSheet.Number
'
'    '计重
'    m_oPrintTicket.SetCurrentObject PAI_CalWeight
'    m_oPrintTicket.LabelSetCaption poAcceptSheet.CalWeight
'
'
'    '收费项
'    Dim i As Integer
'    Dim atTmpPriceItem() As TLuggagePriceItem
'    atTmpPriceItem = poAcceptSheet.PriceItems
'    For i = 1 To ArrayLength(atTmpPriceItem)
'        m_oPrintTicket.SetCurrentObject PAI_PriceItem1 + Val(atTmpPriceItem(i).PriceID)        '输出有效项
'        m_oPrintTicket.LabelSetCaption FormatMoney(atTmpPriceItem(i).PriceValue)
'    Next i
'
'    '金额大写
'    m_oPrintTicket.SetCurrentObject PAI_TotalPriceBig
'    m_oPrintTicket.LabelSetCaption GetNumber(poAcceptSheet.TotalPrice)
'
'    '金额小写
'    m_oPrintTicket.SetCurrentObject PAI_TotalPrice
'    m_oPrintTicket.LabelSetCaption FormatMoney(poAcceptSheet.TotalPrice)
'
'    '开票人
'    m_oPrintTicket.SetCurrentObject PAI_UserName
'    m_oPrintTicket.LabelSetCaption m_oAUser.UserID
'
'
'    '金额小写
'    m_oPrintTicket.SetCurrentObject PAI_TotalPrice2
'    m_oPrintTicket.LabelSetCaption poAcceptSheet.TotalPrice
'
'    '开票人
'    m_oPrintTicket.SetCurrentObject PAI_TotalPriceName
'    m_oPrintTicket.LabelSetCaption "运费+服务费"
'
'
''
'    '收件人电话
'    m_oPrintTicket.SetCurrentObject PAI_PickerPhone
'    m_oPrintTicket.LabelSetCaption poAcceptSheet.PickerPhone
'
'    '托运人电话
'    m_oPrintTicket.SetCurrentObject PAI_ShipperPhone
'    m_oPrintTicket.LabelSetCaption poAcceptSheet.LuggageShipperPhone
'
'    '收件人地址
'    m_oPrintTicket.SetCurrentObject PAI_PickerAddress
'    m_oPrintTicket.LabelSetCaption poAcceptSheet.PickerAddress
'
'    '发车时间
'    m_oPrintTicket.SetCurrentObject PAI_StartTime
'    m_oPrintTicket.LabelSetCaption Format(pdtBusStartTime, "MM-dd hh:mm")
'
'    '年
'    m_oPrintTicket.SetCurrentObject PAI_Year
'    m_oPrintTicket.LabelSetCaption Format(Date, "yyyy")
'
'    '月
'    m_oPrintTicket.SetCurrentObject PAI_Month
'    m_oPrintTicket.LabelSetCaption Format(Date, "MM")
'
'    '日
'    m_oPrintTicket.SetCurrentObject PAI_Day
'    m_oPrintTicket.LabelSetCaption Format(Date, "dd")
'
'
''    '操作日期
''    m_oPrintTicket.SetCurrentObject PAI_OperationDate
''    m_oPrintTicket.LabelSetCaption Format(poAcceptSheet.OperateTime, cszDateStr)
''
''    '实重
''    m_oPrintTicket.SetCurrentObject PAI_ActWeight
''    m_oPrintTicket.LabelSetCaption poAcceptSheet.ActWeight
''
''    '超重件数
''    m_oPrintTicket.SetCurrentObject PAI_OverNumber
''    m_oPrintTicket.LabelSetCaption poAcceptSheet.OverNumber
''
'
'    '承运车次
'    m_oPrintTicket.SetCurrentObject PAI_BusID
'    m_oPrintTicket.LabelSetCaption poAcceptSheet.BusID
''
'    '车次日期
'    m_oPrintTicket.SetCurrentObject PAI_BusDate
'    m_oPrintTicket.LabelSetCaption Format(poAcceptSheet.BusDate, cszDateStr)
'
'    m_oPrintTicket.PrintA 100

''    m_oPrintTicket.ReadFormatFile App.Path & "\AcceptSheet.bpf"
    
'
    m_oPrintTicket.ClosePort
    m_oPrintTicket.OpenPort
    
    '受理时间
    m_oPrintTicket.SetObject PAI_AcceptDate
    m_oPrintTicket.SetCaption Format(Date, "yyyy年MM月dd日")
    
    '车牌号
    m_oPrintTicket.SetObject PAI_Vehicle
    m_oPrintTicket.SetCaption Trim(frmAccept.cboVehicle.Text)
    
    '车牌号2
    m_oPrintTicket.SetObject PAI_Vehicle2
    m_oPrintTicket.SetCaption Trim(frmAccept.cboVehicle.Text)
     
    '托运方式
    m_oPrintTicket.SetObject PAI_TransType
    m_oPrintTicket.SetCaption poAcceptSheet.AcceptType
    
    '长票号
    m_oPrintTicket.SetObject PAI_LongLuggageID
    m_oPrintTicket.SetCaption poAcceptSheet.SheetID
    
    '始发站
    m_oPrintTicket.SetObject PAI_StartStation
    m_oPrintTicket.SetCaption poAcceptSheet.StartStationName
    
    '到站名
    m_oPrintTicket.SetObject PAI_EndStation
    m_oPrintTicket.SetCaption poAcceptSheet.DesStationName
    
    '托运人
    m_oPrintTicket.SetObject PAI_Shipper
    m_oPrintTicket.SetCaption poAcceptSheet.Shipper
    
    '收件人
    m_oPrintTicket.SetObject PAI_Picker
    m_oPrintTicket.SetCaption poAcceptSheet.Picker
    
    '包装
    m_oPrintTicket.SetObject PAI_Pack
    m_oPrintTicket.SetCaption poAcceptSheet.Pack
    
    
    '货名
    m_oPrintTicket.SetObject PAI_LuggageName
    m_oPrintTicket.SetCaption Trim(poAcceptSheet.LuggageName)
    
    '标签号
    m_oPrintTicket.SetObject PAI_LabelNo
    m_oPrintTicket.SetCaption Trim(IIf(poAcceptSheet.StartLabelID = poAcceptSheet.EndLabelID, poAcceptSheet.StartLabelID, poAcceptSheet.StartLabelID & "-" & poAcceptSheet.EndLabelID))
    
    '件数
    m_oPrintTicket.SetObject PAI_LuggageNumber
    m_oPrintTicket.SetCaption poAcceptSheet.Number
    
    '计重
    m_oPrintTicket.SetObject PAI_CalWeight
    m_oPrintTicket.SetCaption poAcceptSheet.CalWeight
    
    
    '收费项
    Dim i As Integer
    Dim atTmpPriceItem() As TLuggagePriceItem
    atTmpPriceItem = poAcceptSheet.PriceItems
    For i = 1 To ArrayLength(atTmpPriceItem)
        m_oPrintTicket.SetObject PAI_PriceItem1 + Val(atTmpPriceItem(i).PriceID)        '输出有效项
        m_oPrintTicket.SetCaption FormatMoney(atTmpPriceItem(i).PriceValue)
    Next i
    
    '金额大写
    m_oPrintTicket.SetObject PAI_TotalPriceBig
    m_oPrintTicket.SetCaption GetNumber(poAcceptSheet.TotalPrice)
    
    
    Dim aszFig() As String
    aszFig = ApartFig(poAcceptSheet.TotalPrice)
    
    
    '大写金额位数
    
    m_oPrintTicket.SetObject PAI_Cent
    m_oPrintTicket.SetCaption GetSplitNum(aszFig, 1)
    
    m_oPrintTicket.SetObject PAI_Jiao
    m_oPrintTicket.SetCaption GetSplitNum(aszFig, 2)
    
    m_oPrintTicket.SetObject PAI_Yuan
    m_oPrintTicket.SetCaption GetSplitNum(aszFig, 3)
    
    m_oPrintTicket.SetObject PAI_Ten
    m_oPrintTicket.SetCaption GetSplitNum(aszFig, 4)
    
    m_oPrintTicket.SetObject PAI_Hundred
    m_oPrintTicket.SetCaption GetSplitNum(aszFig, 5)
    
    m_oPrintTicket.SetObject PAI_Thousand
    m_oPrintTicket.SetCaption GetSplitNum(aszFig, 6)
    
    
    '金额小写
    m_oPrintTicket.SetObject PAI_TotalPrice
    m_oPrintTicket.SetCaption FormatMoney(poAcceptSheet.TotalPrice)
    
    '开票人
    m_oPrintTicket.SetObject PAI_UserName
    m_oPrintTicket.SetCaption m_oAUser.UserID
    
    
    '金额小写
    m_oPrintTicket.SetObject PAI_TotalPrice2
    m_oPrintTicket.SetCaption poAcceptSheet.TotalPrice
    
    '开票人
    m_oPrintTicket.SetObject PAI_TotalPriceName
    m_oPrintTicket.SetCaption "运费+服务费"
    
    

    '收件人电话
    m_oPrintTicket.SetObject PAI_PickerPhone
    m_oPrintTicket.SetCaption poAcceptSheet.PickerPhone
    
    '托运人电话
    m_oPrintTicket.SetObject PAI_ShipperPhone
    m_oPrintTicket.SetCaption poAcceptSheet.LuggageShipperPhone
    
    '收件人地址
    m_oPrintTicket.SetObject PAI_PickerAddress
    m_oPrintTicket.SetCaption poAcceptSheet.PickerAddress
    
    '发车时间
    m_oPrintTicket.SetObject PAI_StartTime
    m_oPrintTicket.SetCaption Format(pdtBusStartTime, "MM-dd hh:mm")

    '年
    m_oPrintTicket.SetObject PAI_Year
    m_oPrintTicket.SetCaption Format(Date, "yyyy")
    
    '月
    m_oPrintTicket.SetObject PAI_Month
    m_oPrintTicket.SetCaption Format(Date, "MM")
    
    '日
    m_oPrintTicket.SetObject PAI_Day
    m_oPrintTicket.SetCaption Format(Date, "dd")
    

    '操作日期
    m_oPrintTicket.SetObject PAI_OperationDate
    m_oPrintTicket.SetCaption Format(poAcceptSheet.OperateTime, cszDateStr)
'
    '实重
    m_oPrintTicket.SetObject PAI_ActWeight
    m_oPrintTicket.SetCaption poAcceptSheet.ActWeight
'
'    '超重件数
'    m_oPrintTicket.SetObject PAI_OverNumber
'    m_oPrintTicket.SetCaption poAcceptSheet.OverNumber
'

    '承运车次
    m_oPrintTicket.SetObject PAI_BusID
    m_oPrintTicket.SetCaption poAcceptSheet.BusID
    
    '承运车次
    m_oPrintTicket.SetObject PAI_BusID2
    m_oPrintTicket.SetCaption poAcceptSheet.BusID
'
    '车次日期
    m_oPrintTicket.SetObject PAI_BusDate
    m_oPrintTicket.SetCaption Format(poAcceptSheet.BusDate, cszDateStr)
    
    
    '始发站
    m_oPrintTicket.SetObject PAI_StartStation2
    m_oPrintTicket.SetCaption poAcceptSheet.StartStationName
    
    '到站名
    m_oPrintTicket.SetObject PAI_EndStation2
    m_oPrintTicket.SetCaption poAcceptSheet.DesStationName
    
    '始发站
    m_oPrintTicket.SetObject PAI_StartStation3
    m_oPrintTicket.SetCaption poAcceptSheet.StartStationName
    
    '到站名
    m_oPrintTicket.SetObject PAI_EndStation3
    m_oPrintTicket.SetCaption poAcceptSheet.DesStationName
    
    '始发站
    m_oPrintTicket.SetObject PAI_StartStation4
    m_oPrintTicket.SetCaption poAcceptSheet.StartStationName
    
    '到站名
    m_oPrintTicket.SetObject PAI_EndStation4
    m_oPrintTicket.SetCaption poAcceptSheet.DesStationName
    
    
    '运输单号1
    m_oPrintTicket.SetObject PAI_TransTicketID1
    m_oPrintTicket.SetCaption poAcceptSheet.TransTicketID
    
    '运输单号2
    m_oPrintTicket.SetObject PAI_TransTicketID2
    m_oPrintTicket.SetCaption poAcceptSheet.TransTicketID
    '运输单号3
    m_oPrintTicket.SetObject PAI_TransTicketID3
    m_oPrintTicket.SetCaption poAcceptSheet.TransTicketID
    '运输单号4
    m_oPrintTicket.SetObject PAI_TransTicketID4
    m_oPrintTicket.SetCaption poAcceptSheet.TransTicketID

    '保险单号1
    m_oPrintTicket.SetObject PAI_InsuranceID1
    m_oPrintTicket.SetCaption poAcceptSheet.InsuranceID

    '备注1
    m_oPrintTicket.SetObject PAI_Annotation1
    m_oPrintTicket.SetCaption poAcceptSheet.Annotation1

    '备注2
    m_oPrintTicket.SetObject PAI_Annotation2
    m_oPrintTicket.SetCaption poAcceptSheet.Annotation2

    
    '件数
    m_oPrintTicket.SetObject PAI_LuggageNumber2
    m_oPrintTicket.SetCaption poAcceptSheet.Number

    '件数
    m_oPrintTicket.SetObject PAI_LuggageNumber3
    m_oPrintTicket.SetCaption poAcceptSheet.Number

    '件数
    m_oPrintTicket.SetObject PAI_LuggageNumber4
    m_oPrintTicket.SetCaption poAcceptSheet.Number

    


    '货名
    m_oPrintTicket.SetObject PAI_LuggageName2
    m_oPrintTicket.SetCaption Trim(poAcceptSheet.LuggageName)
    '货名
    m_oPrintTicket.SetObject PAI_LuggageName3
    m_oPrintTicket.SetCaption Trim(poAcceptSheet.LuggageName)
    '货名
    m_oPrintTicket.SetObject PAI_LuggageName4
    m_oPrintTicket.SetCaption Trim(poAcceptSheet.LuggageName)
    
    
    '年
    m_oPrintTicket.SetObject PAI_Year2
    m_oPrintTicket.SetCaption Format(Date, "yyyy")
    '月
    m_oPrintTicket.SetObject PAI_Month2
    m_oPrintTicket.SetCaption Format(Date, "MM")
    '日
    m_oPrintTicket.SetObject PAI_Day2
    m_oPrintTicket.SetCaption Format(Date, "dd")
    '年
    m_oPrintTicket.SetObject PAI_Year3
    m_oPrintTicket.SetCaption Format(Date, "yyyy")
    '月
    m_oPrintTicket.SetObject PAI_Month3
    m_oPrintTicket.SetCaption Format(Date, "MM")
    '日
    m_oPrintTicket.SetObject PAI_Day3
    m_oPrintTicket.SetCaption Format(Date, "dd")
    '年
    m_oPrintTicket.SetObject PAI_Year4
    m_oPrintTicket.SetCaption Format(Date, "yyyy")
    '月
    m_oPrintTicket.SetObject PAI_Month4
    m_oPrintTicket.SetCaption Format(Date, "MM")
    '日
    m_oPrintTicket.SetObject PAI_Day4
    m_oPrintTicket.SetCaption Format(Date, "dd")
    

    '应结运费
    m_oPrintTicket.SetObject PAI_SettlePrice
    m_oPrintTicket.SetCaption FormatMoney(poAcceptSheet.SettlePrice)
    
    '行包运费
    m_oPrintTicket.SetObject PAI_BasePrice
    m_oPrintTicket.SetCaption FormatMoney(atTmpPriceItem(1).PriceValue)

    '行包运费名称
    m_oPrintTicket.SetObject PAI_BasePriceName
    m_oPrintTicket.SetCaption "运费："
    '行包应结运费名称
    m_oPrintTicket.SetObject PAI_SettlePriceName
    m_oPrintTicket.SetCaption "应结运费："
    
    
    '自定义信息1
    m_oPrintTicket.SetObject PAI_Custom1
    m_oPrintTicket.SetCaption m_szCustom
    
    
'    '自定义信息2
'    m_oPrintTicket.SetObject PAI_Custom2
'    m_oPrintTicket.SetCaption ""
'
'    PAI_Custom1 = 110    '自定义信息1
'    PAI_Custom2 = 111    '自定义信息2
'
'
'
'
'
'
'
'
    
    '长票号
    m_oPrintTicket.SetObject PAI_LongLuggageID2
    m_oPrintTicket.SetCaption poAcceptSheet.SheetID
    
    '长票号
    m_oPrintTicket.SetObject PAI_LongLuggageID3
    m_oPrintTicket.SetCaption poAcceptSheet.SheetID
    
    '长票号
    m_oPrintTicket.SetObject PAI_LongLuggageID4
    m_oPrintTicket.SetCaption poAcceptSheet.SheetID
    
    '发车时间2
    m_oPrintTicket.SetObject PAI_StartTime2
    m_oPrintTicket.SetCaption Format(pdtBusStartTime, "MM-dd hh:mm")
    
    '标签号2
    m_oPrintTicket.SetObject PAI_LabelNo2
    m_oPrintTicket.SetCaption Trim(IIf(poAcceptSheet.StartLabelID = poAcceptSheet.EndLabelID, poAcceptSheet.StartLabelID, poAcceptSheet.StartLabelID & "-" & poAcceptSheet.EndLabelID))
    
    '实重2
    m_oPrintTicket.SetObject PAI_ActWeight2
    m_oPrintTicket.SetCaption poAcceptSheet.ActWeight
    
    '计重2
    m_oPrintTicket.SetObject PAI_CalWeight2
    m_oPrintTicket.SetCaption poAcceptSheet.CalWeight
    
    '开票日期2
    m_oPrintTicket.SetObject PAI_OperationDate2
    m_oPrintTicket.SetCaption Format(poAcceptSheet.OperateTime, cszDateStr)
    
    '开票人2
    m_oPrintTicket.SetObject PAI_UserName2
    m_oPrintTicket.SetCaption poAcceptSheet.OperatorID
    
    '金额大写2
    m_oPrintTicket.SetObject PAI_TotalPriceBig2
    m_oPrintTicket.SetCaption GetNumber(poAcceptSheet.TotalPrice)
    
    '托运方式2
    m_oPrintTicket.SetObject PAI_TransType2
    m_oPrintTicket.SetCaption poAcceptSheet.AcceptType
    
    '受理标记
    m_oPrintTicket.SetObject PAI_Mark
    m_oPrintTicket.SetCaption "→"
    
    '受理标记2
    m_oPrintTicket.SetObject PAI_Mark2
    m_oPrintTicket.SetCaption "→"
    
    
    
    m_oPrintTicket.PrintFile

    m_oPrintTicket.ClosePort
    
#End If
End Sub

'签发单打印
Public Sub PrintCarrySheet(poCarrySheet As CarrySheet)
#If PRINT_SHEET <> 0 Then

    Dim i As Integer
    Dim j As Integer
    Dim oLugSvr As New LuggageSvr
    Dim rsAccept As Recordset
    Dim szDetail() As String
    Dim nNumber As Integer '总件数
    Dim dbPrice As Double
    Dim szAcceptType As String
    Dim szStartStation As String
    Dim szEndStation As String '到站
    Dim szLuggageName As String '品名
    Dim nBaggageNumber As Integer '件数
    Dim szPack As String '包装3
    Dim dbCalWeight As Double '计重
    Dim szLuggageID As String '受理单号
    Dim szPicker As String '收货人
    Dim szPickerPhone As String '收货人电话

    Dim szBusEndStation As String

'    m_oPrintCarrySheet.ClearAll
'    m_oPrintCarrySheet.ReadFormatFileA App.Path & "\CarrySheet.bpf"


    On Error GoTo ErrorHandle
    Dim szYear As String
    Dim szMonth As String
    Dim szDay As String
    Dim szSenderTel As String '发件人电话

    szYear = Year(poCarrySheet.OperateTime)
    szMonth = Month(poCarrySheet.OperateTime)
    szDay = Day(poCarrySheet.OperateTime)

    oLugSvr.Init m_oAUser
    Set rsAccept = oLugSvr.GetAcceptSheetRS(cszEmptyDateStr, cszForeverDateStr, , , , poCarrySheet.SheetID)

    '=====================================
    '受理清单栏
    '=====================================
    
    
    
    
    '签发单号
    m_oPrintCarrySheet.SetCurrentObject PCI_SheetID
    m_oPrintCarrySheet.LabelSetCaption poCarrySheet.SheetID

'    m_oPrintCarrySheet.SetCurrentObject PCI_TransType
'    m_oPrintCarrySheet.LabelSetCaption ""


'    m_oPrintCarrySheet.SetCurrentObject PCI_Year
'    m_oPrintCarrySheet.LabelSetCaption szYear
'
'
'    m_oPrintCarrySheet.SetCurrentObject PCI_Month
'    m_oPrintCarrySheet.LabelSetCaption szMonth
'
'    m_oPrintCarrySheet.SetCurrentObject PCI_Day
'    m_oPrintCarrySheet.LabelSetCaption szDay

'    m_oPrintCarrySheet.SetCurrentObject PCI_UserID
'    m_oPrintCarrySheet.LabelSetCaption poCarrySheet.OperatorID

'    '始发站
'    m_oPrintCarrySheet.SetCurrentObject PCI_StartStation
'    m_oPrintCarrySheet.LabelSetCaption szStartStation

    rsAccept.MoveFirst
    m_oPrintCarrySheet.SetCurrentObject PCI_AcceptDetailItem '受理清单表格
    szSenderTel = ""
    For i = 1 To cnDetailItemCount
        If Not rsAccept.EOF Then
            szSenderTel = szSenderTel & FormatDbValue(rsAccept!shipper_phone) & " "

            szAcceptType = GetLuggageTypeString(FormatDbValue(rsAccept!accept_type))    '得到第一条记录的受理类型
'            szStartStation = FormatDbValue(rsAccept!start_station_name) '起点站
            szLuggageName = FormatDbValue(rsAccept!luggage_name)
            szEndStation = FormatDbValue(rsAccept!des_station_name)
            nBaggageNumber = FormatDbValue(rsAccept!baggage_number)
            szLuggageID = FormatDbValue(rsAccept!luggage_id)
            szPack = FormatDbValue(rsAccept!Pack)
            szPicker = FormatDbValue(rsAccept!Picker)
            szPickerPhone = FormatDbValue(rsAccept!picker_phone)
            dbCalWeight = FormatDbValue(rsAccept!cal_weight)
            ReDim szDetail(1 To 9)
            szDetail(1) = szEndStation
            szDetail(2) = szLuggageID
            szDetail(3) = szLuggageName
            szDetail(4) = nBaggageNumber
            szDetail(5) = szPack
            szDetail(6) = ""
            szDetail(7) = dbCalWeight
            szDetail(8) = szPicker
            szDetail(9) = szPickerPhone




            For j = 1 To 9

                m_oPrintCarrySheet.GridGetCertainCell i, j
                m_oPrintCarrySheet.GridLabelSetCaption szDetail(j)
            Next j


            nNumber = nNumber + FormatDbValue(rsAccept!baggage_number)
            dbPrice = dbPrice + FormatDbValue(rsAccept!price_item_1)        '运费合计

            rsAccept.MoveNext

        Else

            '填充剩余的,用空串

            For j = 1 To 9
                m_oPrintCarrySheet.GridGetCertainCell i, j
                m_oPrintCarrySheet.GridLabelSetCaption ""
            Next j
        End If
    Next i
    
    m_oPrintCarrySheet.SetCurrentObject PCI_CarryTime
    m_oPrintCarrySheet.LabelSetCaption poCarrySheet.OperateTime


'    '行包结算联
    '=====================================
    '行包结算栏
    '=====================================

    '签发单号
    m_oPrintCarrySheet.SetCurrentObject PCI_SheetID2
    m_oPrintCarrySheet.LabelSetCaption poCarrySheet.SheetID

'    m_oPrintCarrySheet.SetCurrentObject PCI_Year2
'    m_oPrintCarrySheet.LabelSetCaption szYear
'
'
'    m_oPrintCarrySheet.SetCurrentObject PCI_Month2
'    m_oPrintCarrySheet.LabelSetCaption szMonth
'
'    m_oPrintCarrySheet.SetCurrentObject PCI_Day2
'    m_oPrintCarrySheet.LabelSetCaption szDay

'    m_oPrintCarrySheet.SetCurrentObject PCI_UserID2
'    m_oPrintCarrySheet.LabelSetCaption poCarrySheet.OperatorID

    szBusEndStation = poCarrySheet.EndStation

    m_oPrintCarrySheet.SetCurrentObject PCI_EndStation
    m_oPrintCarrySheet.LabelSetCaption szBusEndStation


    '车辆号(楚门改为车次)
    m_oPrintCarrySheet.SetCurrentObject PCI_LicenseTagNo
    m_oPrintCarrySheet.LabelSetCaption poCarrySheet.BusID

'    '发车时间
'    m_oPrintCarrySheet.SetCurrentObject PCI_StartTime
'    m_oPrintCarrySheet.LabelSetCaption Format(poCarrySheet.BusStartOffTime, "HH:mm")

    '件数
    m_oPrintCarrySheet.SetCurrentObject PCI_Number
    m_oPrintCarrySheet.LabelSetCaption nNumber

    '金额大写
    m_oPrintCarrySheet.SetCurrentObject PCI_TotalPriceBig
    m_oPrintCarrySheet.LabelSetCaption GetNumber(poCarrySheet.PrintSettlePrice)   'GetNumber(dbPrice)

    '金额小写
    m_oPrintCarrySheet.SetCurrentObject PCI_TotalPrice
    m_oPrintCarrySheet.LabelSetCaption FormatMoney(poCarrySheet.PrintSettlePrice) 'FormatMoney(dbPrice)
    
    m_oPrintCarrySheet.SetCurrentObject PCI_CarryTime2
    m_oPrintCarrySheet.LabelSetCaption poCarrySheet.OperateTime

'
'
'End Enum

    '=====================================
    '装车清单栏
    '=====================================

    '签发单号
    m_oPrintCarrySheet.SetCurrentObject PCI_SheetID3
    m_oPrintCarrySheet.LabelSetCaption poCarrySheet.SheetID

'    m_oPrintCarrySheet.SetCurrentObject PCI_Year3
'    m_oPrintCarrySheet.LabelSetCaption szYear
'
'
'    m_oPrintCarrySheet.SetCurrentObject PCI_Month3
'    m_oPrintCarrySheet.LabelSetCaption szMonth
'
'    m_oPrintCarrySheet.SetCurrentObject PCI_Day3
'    m_oPrintCarrySheet.LabelSetCaption szDay

'    m_oPrintCarrySheet.SetCurrentObject PCI_UserID3
'    m_oPrintCarrySheet.LabelSetCaption poCarrySheet.OperatorID

    '车辆号
    m_oPrintCarrySheet.SetCurrentObject PCI_LicenseTagNo2
    m_oPrintCarrySheet.LabelSetCaption poCarrySheet.VehicleLicense

'    '发车时间
'    m_oPrintCarrySheet.SetCurrentObject PCI_StartTime2
'    m_oPrintCarrySheet.LabelSetCaption Format(poCarrySheet.BusStartOffTime, "HH:mm")

    m_oPrintCarrySheet.SetCurrentObject PCI_CarryDetailItem  '装车清单

    rsAccept.MoveFirst
    For i = 1 To cnDetailItemCount
        If Not rsAccept.EOF Then
'            szAcceptType = GetLuggageTypeString(FormatDbValue(rsAccept!accept_type))    '得到第一条记录的受理类型
            szEndStation = FormatDbValue(rsAccept!des_station_name)
            szLuggageID = FormatDbValue(rsAccept!luggage_id)
            szLuggageName = FormatDbValue(rsAccept!luggage_name)
            nBaggageNumber = FormatDbValue(rsAccept!baggage_number)
            szPack = FormatDbValue(rsAccept!Pack)

            ReDim szDetail(1 To 5)
            szDetail(1) = szEndStation
            szDetail(2) = szLuggageID
            szDetail(3) = szLuggageName
            szDetail(4) = nBaggageNumber
            szDetail(5) = szPack


            For j = 1 To 5

                m_oPrintCarrySheet.GridGetCertainCell i, j
                m_oPrintCarrySheet.GridLabelSetCaption szDetail(j)
            Next j

            rsAccept.MoveNext

        Else

            '填充剩余的,用空串

            For j = 1 To 5
                m_oPrintCarrySheet.GridGetCertainCell i, j
                m_oPrintCarrySheet.GridLabelSetCaption ""
            Next j
        End If
    Next i

    m_oPrintCarrySheet.SetCurrentObject PCI_ShipperPhone
    m_oPrintCarrySheet.LabelSetCaption Trim(szSenderTel)
    
    m_oPrintCarrySheet.SetCurrentObject PCI_CarryTime3
    m_oPrintCarrySheet.LabelSetCaption poCarrySheet.OperateTime
    
    m_oPrintCarrySheet.SetCurrentObject PCI_TotalPrice2
    m_oPrintCarrySheet.LabelSetCaption FormatMoney(poCarrySheet.PrintSettlePrice)
    
    '新加“装卸工”打印,fpd
    m_oPrintCarrySheet.SetCurrentObject PCI_MoveWorker
    m_oPrintCarrySheet.LabelSetCaption poCarrySheet.MoveWorker
    
    m_oPrintCarrySheet.PrintA 100

#End If


    Exit Sub
ErrorHandle:
    ShowErrorMsg


End Sub

'行包退票打印
Public Sub PrintReturnAccept(poAcceptSheet As AcceptSheet, pszCredenceID As String, pdbReturnCharge As Double, pszOperator As String, pdtReturnTime As Date)
           
#If PRINT_SHEET <> 0 Then

'    m_oPrintTicket.ClearAll
'    m_oPrintTicket.ReadFormatFileA App.Path & "\ReturnAccept.bpf"
'
'    m_oPrintTicket.ClosePort
'    m_oPrintTicket.OpenPort
'        '受理时间
'    m_oPrintTicket.SetCurrentObject PRI_ReturnDate
'    m_oPrintTicket.LabelSetCaption Format(Date, "yyyy年MM月dd日")
'
'    '退运凭证号
'    m_oPrintTicket.SetCurrentObject PRI_CredenceID
'    m_oPrintTicket.LabelSetCaption pszCredenceID
'
'    '托运方式
'    m_oPrintTicket.SetCurrentObject PRI_TransType
'    m_oPrintTicket.LabelSetCaption poAcceptSheet.AcceptType
'
'    '长票号
'    m_oPrintTicket.SetCurrentObject PRI_LongLuggageID
'    m_oPrintTicket.LabelSetCaption poAcceptSheet.SheetID
'
'    '始发站
'    m_oPrintTicket.SetCurrentObject PRI_StartStation
'    m_oPrintTicket.LabelSetCaption poAcceptSheet.StartStationName
'
'    '到站名
'    m_oPrintTicket.SetCurrentObject PRI_EndStation
'    m_oPrintTicket.LabelSetCaption poAcceptSheet.DesStationName
'
'    '托运人
'    m_oPrintTicket.SetCurrentObject PRI_Shipper
'    m_oPrintTicket.LabelSetCaption poAcceptSheet.Shipper
'
'    '收件人
'    m_oPrintTicket.SetCurrentObject PRI_Picker
'    m_oPrintTicket.LabelSetCaption poAcceptSheet.Picker
'
'    '货名
'    m_oPrintTicket.SetCurrentObject PRI_LuggageName
'    m_oPrintTicket.LabelSetCaption poAcceptSheet.LuggageName
'
'    '标签号
'    m_oPrintTicket.SetCurrentObject PRI_LabelNo
'    m_oPrintTicket.LabelSetCaption IIf(poAcceptSheet.StartLabelID = poAcceptSheet.EndLabelID, poAcceptSheet.StartLabelID, poAcceptSheet.StartLabelID & "-" & poAcceptSheet.EndLabelID)
'
'    '件数
'    m_oPrintTicket.SetCurrentObject PRI_LuggageNumber
'    m_oPrintTicket.LabelSetCaption poAcceptSheet.Number
'
'    '退运手续费
'    m_oPrintTicket.SetCurrentObject PRI_ReturnCharge
'    m_oPrintTicket.LabelSetCaption pdbReturnCharge
'
'    '退运手续费名称
'    m_oPrintTicket.SetCurrentObject PRI_ReturnChargeName
'    m_oPrintTicket.LabelSetCaption "退运手续费"
'
'    m_oPrintTicket.SetCurrentObject PRI_CalWeight
'    m_oPrintTicket.LabelSetCaption poAcceptSheet.CalWeight
'
'    '金额大写
'    m_oPrintTicket.SetCurrentObject PRI_TotalPriceBig
'    m_oPrintTicket.LabelSetCaption GetNumber(poAcceptSheet.TotalPrice)
'
'    '金额小写
'    m_oPrintTicket.SetCurrentObject PRI_TotalPrice
'    m_oPrintTicket.LabelSetCaption poAcceptSheet.TotalPrice
'
'    '退运人
'    m_oPrintTicket.SetCurrentObject PRI_UserName
'    m_oPrintTicket.LabelSetCaption m_oAUser.UserName
'
'    '操作日期
'    m_oPrintTicket.SetCurrentObject PRI_OperationDate
'    m_oPrintTicket.LabelSetCaption Format(pdtReturnTime, cszDateStr)
'
'
'    '超重件数
'    m_oPrintTicket.SetCurrentObject PRI_OverNumber
'    m_oPrintTicket.LabelSetCaption poAcceptSheet.OverNumber
'
'    m_oPrintTicket.PrintA 100

'    m_oPrintTicket.ClosePort
#End If

End Sub

'托运方式状态转换
Public Function GetLuggageTypeString(pnType As Integer) As String
    Select Case pnType
        Case 0
            GetLuggageTypeString = "快件"
        Case 1
            GetLuggageTypeString = "随行"
    End Select
End Function
Public Function GetLuggageTypeInt(szType As String) As Integer
    Select Case szType
        Case "快件"
            GetLuggageTypeInt = 0
        Case "随行"
            GetLuggageTypeInt = 1
    End Select
    
End Function


Private Function AlignTextToCell(pszString As String, pnFixWidth As Integer) As String
    Dim nActLen As Integer
    nActLen = LenA(pszString)
    If nActLen > pnFixWidth Then
        AlignTextToCell = MidA(pszString, 1, pnFixWidth)
    Else
        AlignTextToCell = Space(pnFixWidth - nActLen) & pszString
    End If
End Function

Private Function GetSplitNum(paszNum() As String, pnNum As Integer) As String
    
    Dim nLen As Integer
    Const cszO = "¤"
    
    nLen = ArrayLength(paszNum)
    If pnNum > nLen Then
        GetSplitNum = cszO
    Else
        If paszNum(pnNum) = "" Then
            GetSplitNum = cszO
        
        Else
            GetSplitNum = paszNum(pnNum)
        End If
    End If
    
End Function

'得到配载车次的发车时间
Public Function GetAllotStationBusStartTime(ByVal pszBusID As String, ByVal pdtBusDate As Date) As Date
On Error GoTo ErrHandle

    Dim oLugSvr As New LuggageSvr
    Dim rsTemp As Recordset
    oLugSvr.Init m_oAUser
    Set rsTemp = oLugSvr.GetAllotStationBusStartTime(pszBusID, pdtBusDate)
    If rsTemp.RecordCount = 1 Then
        GetAllotStationBusStartTime = FormatDbValue(rsTemp!bus_date) & " " & FormatDbValue(rsTemp!bus_start_time)
    Else
        GetAllotStationBusStartTime = cdtEmptyDate
    End If
    
    Exit Function
ErrHandle:
    ShowErrorMsg
End Function
