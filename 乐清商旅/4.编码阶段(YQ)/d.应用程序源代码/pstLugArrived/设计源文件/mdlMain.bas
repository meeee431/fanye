Attribute VB_Name = "mdlMain"
Option Explicit



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

Const cszRecentSeller = "RecentSeller"
'====================================================================
'以下全局常量定义
'--------------------------------------------------------------------
Public Const CSZNoneString = "(全部)"
Public Const CPick_Normal = "未提"
Public Const CPick_Picked = "已提"
Public Const CPick_Canceled = "已废"

Public Const cnColor_Active = &HFF0000
Public Const cnColor_Normal = vbBlack
Public Const cnColor_Edited = vbRed
Public Const cszLongDateFormat = "yyyy年MM月dd日"

Private Declare Function GetComputerNameAPI Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'====================================================================
'以下全局变量定义
'--------------------------------------------------------------------
Public g_szSheetID   As String      '当前受理单号
Public g_oActUser As ActiveUser
Public g_oParam As New SystemParam
Public g_oPackageParam As New PackageParam
Public g_oPackageSvr As New PackageSvr
Public g_oShell As New STShell.CommShell
    '以下是加入短信后新增的全局变量
Public g_bSMSValid As Boolean
Public g_szSMSServer As String
Public g_szSMSUser As String
Public g_szSMSPassword As String
Public g_szSMSApiPort As String

Public Enum PrintPackageObjectIndexEnum
    PPI_PackageID = 1 '行包流水号
    PPI_SheetID = 2 '发票票据号
    PPI_PackageName = 3 '货名
    PPI_PackType = 5 '包装
    PPI_ArrivedDate = 6 ' 到达时间
    PPI_CalWeight = 7 '计重
    PPI_AreaType = 8 '地区类别
    PPI_StartStation = 9 '始发站
    PPI_Vehicle = 10 '车牌号
    PPI_PackageNumber = 11 '件数
    PPI_SavePosition = 12 '存放位置
    
    PPI_Shipper = 13 '托运人
    PPI_ShipperPhone = 14 '托运人电话
    PPI_ShipperUnit = 15 '托运人单位
    PPI_PickType = 16  '交付方式
    
    PPI_Picker = 17 '收件人
    PPI_PickerPhone = 18 '收件人电话
    PPI_PickerUnit = 19 '收件人单位
    PPI_PickerAddress = 20 '收件人地址
    PPI_PickerCredit = 21 '提取人身份证号
    PPI_PickTime = 22 '提件时间
    
    PPI_Operator = 23 '受理人
    PPI_OperationDate = 24 '操作日期
    PPI_UserName = 25 '操作用户
    PPI_Loader = 26 '装卸工
    
    PPI_TransCharge = 27 '代收运费
    
    PPI_TotalPriceBig = 28 '合计大写
    PPI_TotalPrice = 29 '合计（小写）
    PPI_PriceItem1 = 40  '票价1（装卸费）
    PPI_PriceItem2 = 41  '票价2（保管费）
    PPI_PriceItem3 = 42  '票价3（送货费）
    PPI_PriceItem4 = 43  '票价4（搬运费）
    PPI_PriceItem5 = 44  '票价5（其他费）
    
    PPI_Year = 51  '年
    PPI_Month = 52 '月
    PPI_Day = 53 '日
    
    '副联部分
    PPI_AreaType2 = 30 '地区类别
    PPI_StartStation2 = 31 '始发站
    PPI_PackageName2 = 32 '货名
    PPI_PackageNumber2 = 33 '件数
    PPI_PackType2 = 34 '包装
    
    PPI_TotalPrice2 = 35 '合计小写2
    
    
    PPI_Year2 = 54  '年
    PPI_Month2 = 55 '月
    PPI_Day2 = 56 '日
    
    
    '大写金额位数
    PAI_Cent = 61
    PAI_Jiao = 62
    PAI_Yuan = 63
    PAI_Ten = 64
    PAI_Hundred = 65
    PAI_Thousand = 66
    
    PPI_PackageID2 = 70 '货号流水号2
    PPI_Drawer = 71   '提件人
    PPI_DrawerPhone = 72   '提件人电话
    PPI_SheetID2 = 73   '票据号2
    
    PPI_ArrivedDate2 = 84 ' 到达时间2
    PPI_CalWeight2 = 85 '计重2
    PPI_CalWeight3 = 86 '计重3
    PPI_CalWeight4 = 87 '计重4
    PPI_Vehicle2 = 88 '车牌号2
    PPI_TransCharge2 = 89 '代收运费2
    PPI_TotalPriceBig2 = 90 '合计大写2
    PPI_Picker2 = 91 '收件人2
    PPI_Drawer2 = 92   '提件人2
    PPI_PickTime2 = 93 '提件时间2
    PPI_UserName2 = 94 '操作用户2
    
    PPI_Mark = 95 '到达标记
    PPI_Mark2 = 96 '到达标记2
    
End Enum

Dim m_oPrintTicket As FastPrint


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
        Set g_oActUser = g_oShell.ShowLogin()

    Else
        Set g_oActUser = New ActiveUser
        g_oActUser.Login GetLoginParam(szLoginCommandLine, cszUserName), GetLoginParam(szLoginCommandLine, cszUserPassword), GetComputerName()
    End If
    If Not g_oActUser Is Nothing Then
'        App.HelpFile = SetHTMLHelpStrings("SNSellTK.chm") '设定App.HelpFile
        
        g_oParam.init g_oActUser
        Date = g_oParam.NowDate
        Time = g_oParam.NowDateTime

        
        g_oPackageParam.init g_oActUser
        
        g_oPackageSvr.init g_oActUser
'        SetHTMLHelpStrings "pstLugDesk.chm"
        
        'frmSplash.Show

        Set m_oPrintTicket = New FastPrint
        '初始化打印签发单
        FileIsExist (App.Path & "\PackageSheet.ini")
        m_oPrintTicket.ReadFormatFile App.Path & "\PackageSheet.ini"
        
        
    '打开frmChgSheetNo窗体，设置起始的单据号， 放入g_szSheetID
    
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


'签发单号自增
Public Sub IncSheetID(Optional pnOffset As Integer = 1, Optional pbNoShow As Boolean = False)
 On Error GoTo ErrHandle
 
    Dim oReg As New CFreeReg
    Dim szConnectName As String
    
    szConnectName = "Luggage"
    
    oReg.init "RTStation", HKEY_LOCAL_MACHINE, "Software\RTSoft"
    g_szSheetID = FormatSheetID(g_szSheetID + pnOffset)
    If Not pbNoShow Then
        mdiMain.lblSheetNoName.Caption = "当前单据号:"
        mdiMain.lblSheetNo.Caption = g_szSheetID
    End If
    
    oReg.SaveSetting szConnectName, "CurrentSheetID", g_szSheetID
    
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub


Public Function FormatSheetID(pszCheckID As String)
    FormatSheetID = Format(IIf(pszCheckID <> "", pszCheckID, 0), String(g_oPackageParam.SheetIDNumberLen, "0"))
End Function

Public Sub GetAppSetting()
    Dim szLastTicketNo As String
    
    Dim szLastSheetID As String
    On Error GoTo Here
    szLastSheetID = g_oPackageSvr.GetLastSheetID(g_oActUser.UserID)
    
    g_szSheetID = FormatSheetID(szLastSheetID)
    
    IncSheetID , True
    
    Exit Sub
Here:
    ShowErrorMsg
    
    
End Sub

'楚门改的读注册表里的单号
Public Sub GetAppSetting2()

    Dim szLastSheetID As String
    Dim oReg As New CFreeReg
    Dim szConnectName As String

    
    oReg.init "RTStation", HKEY_LOCAL_MACHINE, "Software\RTSoft"
    
    szConnectName = "Luggage"
    
    On Error GoTo Here

    szLastSheetID = oReg.GetSetting(szConnectName, "CurrentSheetID", "")
    
    g_szSheetID = FormatSheetID(szLastSheetID) - 1

    IncSheetID , True
    
    Exit Sub
Here:
    ShowErrorMsg
End Sub

'打印行包受理单
Public Sub PrintAcceptSheet(poAcceptSheet As Package)
           
#If PRINT_SHEET <> 0 Then
    
'    m_oPrintTicket.ClearAll
'    m_oPrintTicket.ReadFormatFileA App.Path & "\AcceptSheet.bpf"
    
    m_oPrintTicket.ClosePort
    m_oPrintTicket.OpenPort
    
    
    '发票单号|2
    m_oPrintTicket.SetObject PPI_SheetID
    m_oPrintTicket.SetCaption poAcceptSheet.SheetID
    
    '起运站|9
    m_oPrintTicket.SetObject PPI_StartStation
    m_oPrintTicket.SetCaption IIf(poAcceptSheet.StartStationName <> "", poAcceptSheet.StartStationName, poAcceptSheet.AreaType)
    
    '品名|3
    m_oPrintTicket.SetObject PPI_PackageName
    m_oPrintTicket.SetCaption poAcceptSheet.PackageName
    
    '件数|11
    m_oPrintTicket.SetObject PPI_PackageNumber
    m_oPrintTicket.SetCaption poAcceptSheet.PackageNumber
    
    '计重|7
    m_oPrintTicket.SetObject PPI_CalWeight
    m_oPrintTicket.SetCaption poAcceptSheet.CalWeight
  
    '收件人|17
    m_oPrintTicket.SetObject PPI_Picker
    m_oPrintTicket.SetCaption poAcceptSheet.Picker
    
    '货号|1
    m_oPrintTicket.SetObject PPI_PackageID
    m_oPrintTicket.SetCaption poAcceptSheet.PackageID
    
    '装卸费|40
    m_oPrintTicket.SetObject PPI_PriceItem1
    m_oPrintTicket.SetCaption poAcceptSheet.LoadCharge
    
    '保管费|41
    m_oPrintTicket.SetObject PPI_PriceItem4
    m_oPrintTicket.SetCaption poAcceptSheet.KeepCharge
    
    '服务费(搬运费)|43
    m_oPrintTicket.SetObject PPI_PriceItem2
    m_oPrintTicket.SetCaption poAcceptSheet.MoveCharge
    
    '代收运费|27
    m_oPrintTicket.SetObject PPI_TransCharge
    m_oPrintTicket.SetCaption poAcceptSheet.TransitCharge
    
    '合计小写|29
    m_oPrintTicket.SetObject PPI_TotalPrice
    m_oPrintTicket.SetCaption poAcceptSheet.LoadCharge + poAcceptSheet.KeepCharge + poAcceptSheet.MoveCharge + poAcceptSheet.TransitCharge
    
    '合计大写|28
    m_oPrintTicket.SetObject PPI_TotalPriceBig
    m_oPrintTicket.SetCaption GetNumber(poAcceptSheet.LoadCharge + poAcceptSheet.KeepCharge + poAcceptSheet.MoveCharge + poAcceptSheet.TransitCharge)
    
    '工号|25
    m_oPrintTicket.SetObject PPI_UserName
    m_oPrintTicket.SetCaption poAcceptSheet.UserID
    
    '提件日期|22
    m_oPrintTicket.SetObject PPI_PickTime
    m_oPrintTicket.SetCaption Format(poAcceptSheet.PickTime, "YYYY-MM-DD HH:mm")
    
    '以下是副联部分
    
    '货号2|70
    m_oPrintTicket.SetObject PPI_PackageID2
    m_oPrintTicket.SetCaption poAcceptSheet.PackageID
    
    '件数2|33
    m_oPrintTicket.SetObject PPI_PackageNumber2
    m_oPrintTicket.SetCaption poAcceptSheet.PackageNumber
    
    '提件证件|21
    m_oPrintTicket.SetObject PPI_PickerCredit
    If poAcceptSheet.PickerCreditID <> "" Then
        m_oPrintTicket.SetCaption Left(poAcceptSheet.PickerCreditID, Len(poAcceptSheet.PickerCreditID) - 4) & "****"
    Else
        m_oPrintTicket.SetCaption ""
    End If
    
    '合计小写2|35
    m_oPrintTicket.SetObject PPI_TotalPrice2
    m_oPrintTicket.SetCaption poAcceptSheet.LoadCharge + poAcceptSheet.KeepCharge + poAcceptSheet.MoveCharge + poAcceptSheet.TransitCharge
    
    '工号2|94
    m_oPrintTicket.SetObject PPI_UserName2
    m_oPrintTicket.SetCaption poAcceptSheet.UserID
    
    '提件日期2|93
    m_oPrintTicket.SetObject PPI_PickTime2
    m_oPrintTicket.SetCaption Format(poAcceptSheet.PickTime, "YYYY-MM-DD HH:mm")
  
    
    DoEvents
    m_oPrintTicket.PrintFile

    m_oPrintTicket.ClosePort
    
#End If
End Sub



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
    
    Dim nlen As Integer
    Const cszO = "¤"
    
    nlen = ArrayLength(paszNum)
    If pnNum > nlen Then
        GetSplitNum = cszO
    Else
        If paszNum(pnNum) = "" Then
            GetSplitNum = cszO
        
        Else
            GetSplitNum = paszNum(pnNum)
        End If
    End If
    
End Function

Public Sub SetFlex(vsItem As VSFlexGrid, Optional pnRows As Integer = -1, Optional pnCols As Integer = -1)
    If pnRows <> -1 Then
        vsItem.Rows = pnRows
        If pnRows > 0 Then vsItem.FixedRows = 1
    End If
    If pnCols <> -1 Then
        vsItem.Cols = pnCols
        If pnCols > 0 Then vsItem.FixedCols = 1
    End If
 
End Sub

Public Sub SaveRecentSeller(pvaUser As Variant)
    Dim oReg As New CFreeReg
    oReg.init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany
    Dim nSellerCount As Integer
    Dim i As Integer
    Dim szRecentSeller As String
    nSellerCount = ArrayLength(pvaUser)
    If nSellerCount > 0 Then
        szRecentSeller = pvaUser(1)
        For i = 2 To nSellerCount
            szRecentSeller = szRecentSeller & "," & pvaUser(i)
        Next
        oReg.SaveSetting cszLuggageAccount, cszRecentSeller, szRecentSeller
    End If
End Sub

Public Function GetRecentSeller() As String
    Dim oReg As New CFreeReg
    oReg.init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany
    GetRecentSeller = oReg.GetSetting(cszLuggageAccount, cszRecentSeller)
End Function


Public Sub FillSellStation(cboSellStation As ComboBox)
    Dim oSystemMan As New SystemMan
    Dim atTemp() As TDepartmentInfo
    Dim i As Integer
    On Error GoTo Here
    '判断用户属于哪个上车站,如果为空则填充一个空行,再填充所有的上车站
    oSystemMan.init g_oActUser
    atTemp = oSystemMan.GetAllSellStation(g_oActUser.UserUnitID)
    If g_oActUser.SellStationID = "" Then
        cboSellStation.AddItem ""
        For i = 1 To ArrayLength(atTemp)
            cboSellStation.AddItem MakeDisplayString(atTemp(i).szSellStationID, atTemp(i).szSellStationName)
        Next i
    '否则只填充用户所属的上车站
    Else
        For i = 1 To ArrayLength(atTemp)
            If g_oActUser.SellStationID = atTemp(i).szSellStationID Then
               cboSellStation.AddItem MakeDisplayString(atTemp(i).szSellStationID, atTemp(i).szSellStationName)
               Exit For
            End If
        Next i
        cboSellStation.ListIndex = 0
    End If
    Exit Sub
Here:
    ShowErrorMsg
End Sub
Public Function BuildPacketID(plInputID As Long, Optional pszYear As String, Optional pszMonth As String) As Long
    If pszYear = "" Then
        pszYear = Format(Date, "yy")
    End If
    If pszMonth = "" Then
        pszMonth = Format(Date, "MM")
    End If
    BuildPacketID = Val(pszYear & pszMonth & Format(plInputID, "000000"))
End Function
Public Function UnBuildPacketID(plPacketID As Long) As Long
    UnBuildPacketID = Val(plPacketID Mod 10 ^ 6)
    
End Function

Public Sub InitSMS()
'    Dim oReg As New CFreeReg
'    oReg.init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany       'HKEY_LOCAL_MACHINE
'    '1先将默认值读出
'
'
'    g_szSMSServer = oReg.GetSetting("Package", "SMSServer")
'    g_szSMSUser = oReg.GetSetting("Package", "SMSUser")
'    g_szSMSPassword = oReg.GetSetting("Package", "SMSPwd")
'    g_szSMSApiPort = Val(oReg.GetSetting("Package", "SMSApiPort"))
'
'    If init(g_szSMSServer, g_szSMSUser, g_szSMSPassword, g_szSMSApiPort) = 0 Then
'        g_bSMSValid = True
'    Else
'        g_bSMSValid = False
'    End If

    If init("10.20.20.20", "xb", "xb123", 11) = 0 Then
        g_bSMSValid = True
    Else
        g_bSMSValid = False
    End If
End Sub
Public Sub SendSMS(pszPhone As String, pszMessage As String)
'    If sendSM(pszPhone, pszMessage, Val(g_szSMSApiPort)) = 0 Then
'        MsgBox "发送成功!", vbInformation, "提示"
'    Else
'        MsgBox "发送失败!", vbExclamation, "错误"
'    End If
        
    If sendSM(pszPhone, pszMessage, 11) = 0 Then
        MsgBox "发送成功!", vbInformation, "提示"
    Else
        MsgBox "发送失败!", vbExclamation, "错误"
    End If
End Sub
Public Sub ReleaseSMS()
    If g_bSMSValid Then
        release
    End If
End Sub

'刷新界面上的单据号
Public Sub RefreshCurrentSheetID()
    Dim szLastSheetID As String
    Dim oReg As New CFreeReg
    Dim szConnectName As String
    
    
    oReg.init "RTStation", HKEY_LOCAL_MACHINE, "Software\RTSoft"
    
    szConnectName = "Luggage"
    
    On Error GoTo Here
    szLastSheetID = oReg.GetSetting(szConnectName, "CurrentSheetID", "")
    
    g_szSheetID = FormatSheetID(szLastSheetID)
    
    mdiMain.lblSheetNoName.Caption = "当前单据号:"
    mdiMain.lblSheetNo.Caption = g_szSheetID
 
    oReg.SaveSetting szConnectName, "CurrentSheetID", g_szSheetID
       
    Exit Sub
Here:
    ShowErrorMsg
End Sub
