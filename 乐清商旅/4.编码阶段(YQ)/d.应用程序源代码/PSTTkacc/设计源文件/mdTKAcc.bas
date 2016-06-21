Attribute VB_Name = "mdTKAcc"
Option Explicit


Public Const cszByOperationTime = "按售票日期统计"
Public Const cszByBusDate = "按车次日期统计"

Public Const cszBySale = "按售票统计"
Public Const cszByCheck = "按检票统计"


Public Enum EBusStationType
    SNBusFromSale = 0
    SNBusFromCheck = 1
    SNVehicleFromCheck = 2
    
End Enum
Public Enum EFormStatus
    SNAddNew = 0
    SNModify = 1
    SNShow = 2
End Enum
Const cszDocumentDir = "DocumentDir"
Const cszDefDocumentDir = "C:\"
Const cszDefSimpleMsg = "准备!"
Const cszTicketAccount = "TkAcc"
Const cszRecentSeller = "RecentSeller"

Public m_oActiveUser As ActiveUser
Public m_oParam As New SystemParam
Public m_oShell As New CommDialog

Private m_aszUsedTicketItem() As String
Private m_nUsedTicketItemCount As Integer

Public m_rsPriceItem As Recordset
Public m_rsTicketType As Recordset
Public m_AllTicketType As Recordset


Public g_szUnitID As String
'====================================================================
'以下定义枚举
'--------------------------------------------------------------------
'主界面状态条字符串区域
Public Enum EStatusBarArea
    ESB_WorkingInfo = 1
    ESB_ResultCountInfo = 2
    ESB_UserInfo = 3
    ESB_LoginTime = 4
End Enum

Public Enum EBusStatMode
    ST_BySalerStationAndSaleTime = 0 '按售票员的所属车站及售票时间统计
    ST_ByBusStationAndBusDate = 1 '按车次上车站及车次日期统计
    ST_BySalerStationAndBusDate = 2 '按售票员的所属车站及车次日期统计
    
End Enum

Public Sub Main()
    Dim szLoginCommandLine As String
    Dim oPriceMan As New STPrice.TicketPriceMan
    Dim oShell As New CommShell
    
    
'    szLoginCommandLine = TransferLoginParam(Trim(Command()))
    
    On Error GoTo HelpFileErr
'    App.HelpFile = SetHTMLHelpStrings("SNTKAcc.chm")
    
GoOn:
    On Error GoTo 0
    
    
    m_nUsedTicketItemCount = -1
'    If szLoginCommandLine = "" Then
'        Set m_oActiveUser = oShell.ShowLogin()
'    Else
'        Set m_oActiveUser = New ActiveUser
'        m_oActiveUser.Login GetLoginParam(szLoginCommandLine, cszUserName), GetLoginParam(szLoginCommandLine, cszUserPassword), GetComputerName()
'        m_oShell.Init m_oActiveUser
'    End If
    
    Set m_oActiveUser = oShell.ShowLogin()
    If m_oActiveUser Is Nothing Then Exit Sub
'    oShell.ShowSplash "站务统计分析", "Station Business Account", LoadResPicture(101, 0), App.Major, App.Minor, App.Revision
    DoEvents
    
    oPriceMan.Init m_oActiveUser
    Set m_rsPriceItem = oPriceMan.GetAllTicketItemRS(TP_PriceItemUse)
    Set m_rsTicketType = m_oParam.GetAllTicketTypeRS(TP_TicketTypeValid)
    Set m_AllTicketType = m_oParam.GetAllTicketTypeRS(TP_TicketTypeAll)
    Set oPriceMan = Nothing
    
    
    m_oParam.Init m_oActiveUser
    
    g_szUnitID = m_oParam.UnitID
    SetHTMLHelpStrings "STTkAcc.chm"
    m_oShell.Init m_oActiveUser
    InitTicketItemInfo
    MDIMain.Show
    oShell.CloseSplash
    DoEvents
    Exit Sub
HelpFileErr:
    ShowMsg "不能找到帮助文件"
    Resume GoOn
End Sub


Private Sub InitTicketItemInfo()
    If m_nUsedTicketItemCount = -1 Then
'        Dim oPriceMan As New RoutePriceTable
        Dim oPriceM As New TicketPriceMan
        Dim oScheme As New RegularScheme
        Dim aszTemp() As String
'        oPriceMan.Init m_oActiveUser
        oScheme.Init m_oActiveUser
'        aszTemp = oScheme.GetRunPriceTable
'        oPriceMan.Identify aszTemp(1, 2)
        oPriceM.Init m_oActiveUser
        m_aszUsedTicketItem = oPriceM.GetAllTicketItem()
               
        m_nUsedTicketItemCount = ArrayLength(m_aszUsedTicketItem)
    End If
End Sub



Public Sub SetMouseBusy(pbBusy As Boolean)
    If pbBusy Then
        Screen.MousePointer = vbHourglass
        DoEvents
    Else
        Screen.MousePointer = vbDefault
    End If
    
End Sub


Public Function GetDocumentDir() As String
    Dim oReg As New CFreeReg
    Dim szFileDir As String
    On Error GoTo Error_Handle
    oReg.Init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany
    szFileDir = oReg.GetSetting(cszTicketAccount, cszDocumentDir, cszDefDocumentDir)
    szFileDir = IIf(szFileDir = "", cszDefDocumentDir, szFileDir)
    
    GetDocumentDir = szFileDir
    Exit Function
Error_Handle:
    GetDocumentDir = cszDefDocumentDir
End Function

Public Sub SaveDocumentDir(pszFullFileName As String)
    Dim oReg As New CFreeReg
    Dim szPath As String
    On Error Resume Next
    szPath = Left(pszFullFileName, InStrRev(pszFullFileName, "\") - 1)
    oReg.Init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany
    If szPath <> "" Then
        oReg.SaveSetting cszTicketAccount, cszDocumentDir, szPath
    Else
        oReg.SaveSetting cszTicketAccount, cszDocumentDir, cszDocumentDir
    End If
End Sub

Public Sub SaveRecentSeller(pvaUser As Variant)
    Dim oReg As New CFreeReg
    oReg.Init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany
    Dim nSellerCount As Integer
    Dim i As Integer
    Dim szRecentSeller As String
    nSellerCount = ArrayLength(pvaUser)
    If nSellerCount > 0 Then
        szRecentSeller = pvaUser(1)
        For i = 2 To nSellerCount
            szRecentSeller = szRecentSeller & "," & pvaUser(i)
        Next
        oReg.SaveSetting cszTicketAccount, cszRecentSeller, szRecentSeller
    End If
End Sub

Public Function GetRecentSeller() As String
    Dim oReg As New CFreeReg
    oReg.Init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany
    GetRecentSeller = oReg.GetSetting(cszTicketAccount, cszRecentSeller)
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
    With MDIMain
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
    With MDIMain.abMenu.Bands("statusBar")
        If pbVisual Then
            If Not .Tools("progressBar").Visible Then
                .Tools("progressBar").Visible = True
                .Tools("pnResultCountInfo").Caption = ""
                .Tools("pnResultCountInfo").Visible = False
                MDIMain.pbLoad.Max = 100
                MDIMain.abMenu.RecalcLayout
            End If
            nCurrProcess = Int(plCurrValue / plMaxValue * 100)
            MDIMain.pbLoad.Value = nCurrProcess
        Else
            .Tools("progressBar").Visible = False
            .Tools("pnResultCountInfo").Visible = True
        End If
    End With
End Sub


Public Sub FillSellStation(cboSellStation As ComboBox)
    Dim oSystemMan As New SystemMan
    Dim atTemp() As TDepartmentInfo
    Dim i As Integer
    On Error GoTo Here
    '判断用户属于哪个上车站,如果为空则填充一个空行,再填充所有的上车站
    oSystemMan.Init m_oActiveUser
    atTemp = oSystemMan.GetAllSellStation(g_szUnitID)
    If m_oActiveUser.SellStationID = "" Then
        cboSellStation.AddItem ""
        For i = 1 To ArrayLength(atTemp)
            cboSellStation.AddItem MakeDisplayString(atTemp(i).szSellStationID, atTemp(i).szSellStationName)
        Next i
    '否则只填充用户所属的上车站
    Else
        For i = 1 To ArrayLength(atTemp)
            If m_oActiveUser.SellStationID = atTemp(i).szSellStationID Then
               cboSellStation.AddItem MakeDisplayString(atTemp(i).szSellStationID, atTemp(i).szSellStationName)
               Exit For
            End If
        Next i
        cboSellStation.ListIndex = 0
    End If
    Exit Sub
Here:
    ShowMsg err.Description
End Sub


Public Function GetStatName(pnParam As EBusStatMode) As String
    If pnParam = ST_BySalerStationAndSaleTime Then
        GetStatName = "按售票员的所属车站及售票时间统计"
    ElseIf pnParam = ST_ByBusStationAndBusDate Then
        GetStatName = "按车次上车站及车次日期统计"
    Else
        GetStatName = "按售票员的所属车站及车次日期统计"
    End If
    
End Function

