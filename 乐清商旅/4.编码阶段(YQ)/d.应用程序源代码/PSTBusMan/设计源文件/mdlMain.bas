Attribute VB_Name = "mdlMain"
Option Explicit

' *******************************************************************
' *  Source File Name: mdlMain                                      *
' *  Brief Description: 系统主模块                                  *
' *******************************************************************
'====================================================================
'以下常量定义
'--------------------------------------------------------------------
'Public Const CnInternetCanSell = 0 '可互联售票
'Public Const CnInternetNotCanSell = 1   '不可互联售票
Public Const cszRegKeySystem = "RTBusMan" '本系统的注册表键
Public Const cvChangeColor = vbBlue
Public Const cszKeyPopMenu = 93
Public Const cnPreViewMaxDays = 30 '生成环境预览的最多天数

'票价中的常量声明
Public Const cnNotRunTable = 0 '未运行的票价表
Public Const cnRunTable = 1 '正在执行的票价表
Public Const cszItemBaseCarriage = "0000" '基本运价
Public Const cnAllBusType = 100 '所有车次类型，用于尾数处理
Public Const cszAllBusType = "所有类型" '所有车次类型，用于尾数处理


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
'对话框调用状态
Public Enum EFormStatus
    EFS_AddNew = 0
    EFS_Modify = 1
    EFS_Show = 2
    EFS_Delete = 3
End Enum

Public Enum ECheckStatus
    NormalTicket = 1
    ChangeTicket = 2
    MergeTicket = 3
End Enum

'====================================================================
'以下全局变量
'--------------------------------------------------------------------
Public g_oActiveUser As ActiveUser
'Public g_szExePlanID As String '执行的车次计划
Public g_nPreSell As Byte
Public g_szExePriceTable As String '执行的票价表
Public g_szLocalUnit As String      '当前单位
Public g_bStopAllRefundment As Boolean '是否全额退票
Public g_szLicenseForce As String '车牌前缀

Public g_szUserPassword As String '用户口令
Public g_szStationID As String '系统参数中的本站代码

'==================================================================
'以下变量票价部分用到
Public g_atTicketTypeValid() As TTicketType '可用的票种明细
Public g_nTicketCountValid As Integer   '可用的票种数目
Public g_atAllSellStation() As TDepartmentInfo  '所有的售票站点

Public Sub Main()
'===================================================
'Modify Date：2002-11-19
'Author:陆勇庆
'Reamrk:添加了全局变量g_atAllSellStation用于存放所有售票站点
'===================================================
    
    On Error GoTo ErrHandle
    Dim oShell As New CommShell
    Dim dtTemp As Date
    Dim oScheme As New RegularScheme
    Dim oBus As New BusProject
    Dim oSys As New SystemParam
'    If App.PrevInstance Then
'    MsgBox "应用程序已打开!", vbExclamation, "错误"
'    Exit Sub
'    End If
    
    
    
    
    Set g_oActiveUser = oShell.ShowLogin()
    
    If g_oActiveUser Is Nothing Then Exit Sub
'    oShell.ShowSplash "综合调度", "Bus Scheme", LoadResPicture(101, 0), App.Major, App.Minor, App.Revision
    DoEvents
    dtTemp = Now
    g_szUserPassword = oShell.UserPassword
    '初始调用，得到各项全局参数
    oScheme.Init g_oActiveUser
    oBus.Init g_oActiveUser
    oSys.Init g_oActiveUser
    g_bStopAllRefundment = True
    
    Time = oSys.NowDateTime
    Date = oSys.NowDate
    g_szLocalUnit = oSys.UnitID
    g_nPreSell = oSys.PreSellDate
    g_szStationID = oSys.StationID
    '得到可用的票种信息
    g_atTicketTypeValid = oSys.GetAllTicketType(TP_TicketTypeValid, False)
    g_nTicketCountValid = ArrayLength(g_atTicketTypeValid)
    
'    g_szExePlanID = oScheme.GetExecuteBusProject(Now).szProjectID
    '加入系统系统
    '    g_bStopFullReturn = True
    
    
    
    '设置帮助
    SetHTMLHelpStrings "stBusMan.chm"
    
    
    
    Dim oBase As New SystemMan
    oBase.Init g_oActiveUser
    g_atAllSellStation = oBase.GetAllSellStation '(g_szLocalUnit)
    
    oBus.Identify
    g_szExePriceTable = oBus.ExecutePriceTable
    Load MDIScheme
    MDIScheme.Show
'    DoEvents
'    Do
'    Loop While Second(Now - dtTemp) <= 3
    oShell.CloseSplash
    DoEvents
    

    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

'下拉框加入站点数据
Public Function AddCboStation(cboStationTemp As Object) As Boolean
    Dim oBusInfo As New BaseInfo
    Dim i As Integer
    Dim szaData() As String
    Dim nCount As Integer
    Dim cboStation As ComboBox
    Set cboStation = cboStationTemp
    oBusInfo.Init g_oActiveUser
    szaData = oBusInfo.GetStation(, cboStation.Text, cboStation.Text, cboStation.Text)
    Set oBusInfo = Nothing
    nCount = ArrayLength(szaData)
    If nCount > 0 Then
        cboStation.Clear
        For i = 1 To nCount
            cboStation.AddItem Trim(szaData(i, 1)) & "[" & Trim(szaData(i, 2)) & "]"
        Next
        cboStation.ListIndex = 0
    Else
        AddCboStation = False
        Beep
    End If
    AddCboStation = True
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
    With MDIScheme
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
    If pszShowInfo <> cszUnrepeatString Then ShowSBInfo pszShowInfo, ESB_WorkingInfo
    If plMaxValue = 0 And pbVisual = True Then Exit Sub
    Dim nCurrProcess As Integer
    With MDIScheme.abMenuTool.Bands("statusBar")
        If pbVisual Then
            If Not .Tools("progressBar").Visible Then
                .Tools("progressBar").Visible = True
                MDIScheme.pbLoad.Max = 100
                MDIScheme.abMenuTool.RecalcLayout
            End If
            nCurrProcess = Int(plCurrValue / plMaxValue * 100)
            MDIScheme.pbLoad.Value = nCurrProcess
        Else
            .Tools("progressBar").Visible = False
        End If
    End With
End Sub

'写主界面的标题栏
Public Sub WriteTitleBar(Optional pszFormName As String = "", Optional poIcon As StdPicture)
'    'pszFormName空时则清空
'    With MDIScheme
'    If pszFormName = "" Then
'        .lblInfoBar = ""
'        Set .imgInfoBar.Picture = Nothing
'    Else
'        .lblInfoBar = pszFormName
'        Set .imgInfoBar.Picture = poIcon
'    End If
'    End With
End Sub
'是否激活系统工具栏
Public Sub ActiveSystemToolBar(pbTrue As Boolean)
    With MDIScheme
        .abMenuTool.Bands("mnu_System").Tools("mnu_ExportFile").Enabled = pbTrue
        .abMenuTool.Bands("mnu_System").Tools("mnu_ExportFileOpen").Enabled = pbTrue
        .abMenuTool.Bands("mnu_System").Tools("mnu_system_print").Enabled = pbTrue
        .abMenuTool.Bands("mnu_System").Tools("mnu_system_printview").Enabled = pbTrue
        .abMenuTool.Bands("mnu_System").Tools("mnu_PageOption").Enabled = pbTrue
        .abMenuTool.Bands("mnu_System").Tools("mnu_PrintOption").Enabled = pbTrue
        .abMenuTool.Bands("tbn_system").Tools("tbn_system_export").Enabled = pbTrue
        .abMenuTool.Bands("tbn_system").Tools("tbn_system_exportopen").Enabled = pbTrue
        .abMenuTool.Bands("tbn_system").Tools("tbn_system_print").Enabled = pbTrue
        .abMenuTool.Bands("tbn_system").Tools("tbn_system_printview").Enabled = pbTrue
    End With
End Sub


Public Function MakeArray(ByRef tvBuArray() As String, szOther As String)
    Dim nCount As Integer
    Dim j As Integer
    Dim bflgTemp As Boolean
    Dim nCountTemp As Integer
    nCountTemp = ArrayLength(tvBuArray)
    If nCountTemp = 0 Then
        ReDim tvBuArray(1 To 1) As String
        nCountTemp = 1
    End If
    j = 1
    If nCountTemp = 1 Then
        If tvBuArray(1) = "" Then
            tvBuArray(1) = szOther
        Else
            For j = 1 To nCountTemp
                If Trim(tvBuArray(j)) = Trim(szOther) Then
                    bflgTemp = False
                    Exit For
                Else
                    bflgTemp = True
                End If
            Next
        End If
    Else
        For j = 1 To nCountTemp
            If Trim(tvBuArray(j)) = Trim(szOther) Then
                bflgTemp = False: Exit For
            Else
                bflgTemp = True
            End If
        Next
    End If
    If bflgTemp = True Then
        ReDim Preserve tvBuArray(1 To nCountTemp + 1)
        tvBuArray(nCountTemp + 1) = Trim(szOther)
    End If

End Function

Public Function IdentifyBusStatus(eStatuts As EREBusStatus) As Boolean
    Dim szMsgBusStatus As String
    IdentifyBusStatus = False
    If eStatuts = ST_BusNormal Then IdentifyBusStatus = True: Exit Function
    If eStatuts <> ST_BusSlitpStop And eStatuts <> ST_BusStopped And eStatuts <> ST_BusMergeStopped Then
        Select Case eStatuts
        Case 3
            szMsgBusStatus = "车次已经停检"
        Case 4
            szMsgBusStatus = "车次正在检票"
        Case 5
            szMsgBusStatus = "车次正在补检"
        Case 16
            szMsgBusStatus = "车次已顶班"
        Case Else
            If eStatuts >= 32 Then
                szMsgBusStatus = "车次可能正在顶班或拆分操作，不能停班"
                MsgBox szMsgBusStatus, vbInformation + vbOKOnly, "车次停班"
                Exit Function
            Else
                IdentifyBusStatus = True
                Exit Function
            End If
        End Select
        If szMsgBusStatus = "" Or MsgBox(szMsgBusStatus, vbInformation + vbYesNo, "车次停班") = vbYes Then
            IdentifyBusStatus = True
        End If
    Else
        Select Case eStatuts
        Case ST_BusSlitpStop
            szMsgBusStatus = "车次已拆分停班"
        Case ST_BusStopped
            szMsgBusStatus = "车次已停班,，不能再停班"
        Case ST_BusMergeStopped
            szMsgBusStatus = "车次并班停班"
        End Select
        MsgBox szMsgBusStatus, vbInformation + vbOKOnly, "车次停班"
    End If
End Function

'
'Public Sub ShowTBInfo(Optional pszMsgInfo As String, Optional pnCount As Integer, Optional pnIndex As Integer, Optional pbVisibled As Boolean)
'
'End Sub

Public Function ConvertTypeFromArray(paszBusID() As String, paszVehicleModel() As String, paszSeatType() As String) As TBusVehicleSeatType()
    '将车次代码、车型、座位类型数组转换为类型
    Dim nBus As Integer
    Dim nSeatType As Integer
    Dim nVehicleType As Integer
    Dim lTemp As Long
    Dim atBusVehicleSeat() As TBusVehicleSeatType
    Dim i As Integer
    Dim j As Integer
    Dim n As Integer
    
    nBus = ArrayLength(paszBusID)
    nSeatType = ArrayLength(paszSeatType)
    nVehicleType = ArrayLength(paszVehicleModel)
'    If nSeatType > 30 And nBus = nSeatType And nBus = nVehicleType Then
        '如果座位类型个数大于30 且车次数等于座位类型数及车型数则
        lTemp = nBus
        ReDim atBusVehicleSeat(1 To lTemp)
        For i = 1 To lTemp
            atBusVehicleSeat(i).szbusID = paszBusID(i)
            atBusVehicleSeat(i).szVehicleTypeCode = paszVehicleModel(i)
            atBusVehicleSeat(i).szSeatTypeID = paszSeatType(i)
        Next i
'    Else
'        lTemp = nBus * nSeatType * nVehicleType
'        ReDim atBusVehicleSeat(1 To lTemp)
'        lTemp = 0
'        For i = 1 To nBus
'            For j = 1 To nVehicleType
'                For n = 1 To nSeatType
'                    lTemp = lTemp + 1
'                    atBusVehicleSeat(lTemp).szBusID = paszBusID(i)
'                    atBusVehicleSeat(lTemp).szVehicleTypeCode = paszVehicleModel(j)
'                    atBusVehicleSeat(lTemp).szSeatTypeID = paszSeatType(n)
'                Next n
'            Next j
'        Next i
'    End If
    ConvertTypeFromArray = atBusVehicleSeat
End Function

'Public Function GetProjectExcutePriceTable(ProjectID As String) As String
'    '得到某计划的当前的执行票价表
'    On Error GoTo ErrorHandle
'    Dim oRegularScheme As New RegularScheme
'    Dim aszTable() As String
'    Dim i As Integer, nCount As Integer
'    Dim szTemp As String
'
'    oRegularScheme.Init g_oActiveUser
'    aszTable = oRegularScheme.ProjectExistTable(ProjectID)
'    nCount = ArrayLength(aszTable)
'    If nCount > 0 Then
'        For i = 1 To nCount
'             If aszTable(i, 6) <= Now Then
'               szTemp = aszTable(i, 2)
'               Exit For
'            End If
'        Next
'    End If
'    Set oRegularScheme = Nothing
'    GetProjectExcutePriceTable = szTemp
'    Exit Function
'
'ErrorHandle:
'    MsgBox "此计划无相应票价表"
'End Function

Public Function GetPriceTable(ThisDate As Date) As String()
    '得到票价表,并将其按一定顺序排列后,返回
    
    Dim aszRoutePriceTable() As String
    Dim i, j As Integer, nCount As Integer
    Dim szPriceTable As String
    Dim oRegularScheme As New RegularScheme
    Dim tTemp As TSchemeArrangement
    Dim szRunProject As String

    Dim szPriceTableTemp() As String
    Dim dtMaxDate As Date '可执行票价表的最大执行日期]
    Dim oTicketPriceMan As New TicketPriceMan

On Error GoTo ErrorHandle
    oRegularScheme.Init g_oActiveUser
    tTemp = oRegularScheme.GetExecuteBusProject(Now)
    szRunProject = tTemp.szProjectID
    oTicketPriceMan.Init g_oActiveUser
    aszRoutePriceTable = oTicketPriceMan.GetAllRoutePriceTable()
    nCount = ArrayLength(aszRoutePriceTable)
    dtMaxDate = Format("1900-01-01", cszDateStr)

    If nCount > 0 Then
       ReDim szPriceTableTemp(1 To nCount, 7)
       '把起始执行日期在现在日期前的票价表置为执行标记，其他均为不执行标记
       '但不是最终传出标记
        For i = 1 To nCount
            For j = 1 To 6
                szPriceTableTemp(i, j) = aszRoutePriceTable(i, j)
            Next
            If aszRoutePriceTable(i, 6) = szRunProject And Format(aszRoutePriceTable(i, 3), cszDateStr) <= Format(ThisDate, cszDateStr) Then
               szPriceTableTemp(i, 7) = cnRunTable
               If dtMaxDate < Format(aszRoutePriceTable(i, 3), cszDateStr) Then
                  dtMaxDate = Format(aszRoutePriceTable(i, 3), cszDateStr)
               End If
            Else
               szPriceTableTemp(i, 7) = cnNotRunTable
            End If
        Next
        '把起始执行日期在现在日期前、离现在日期最近的票价表置为执行，其他均为不执行，
        '保留唯一票价表置为执行票价表
        For i = 1 To nCount
            If szPriceTableTemp(i, 7) = cnRunTable Then
               If dtMaxDate > Format(aszRoutePriceTable(i, 3), cszDateStr) Then
                  szPriceTableTemp(i, 7) = cnNotRunTable
               End If
            End If
        Next
    End If

    GetPriceTable = szPriceTableTemp
    Exit Function
ErrorHandle:
    ShowErrorMsg
End Function

'文本框文本长度检查
Public Function TextLongValidate(nCharLong As Integer, szText As String) As Boolean
    Dim szTemp As String, szTemp1 As String, szTemp2 As String
    szTemp1 = CStr(nCharLong)
    If nCharLong Mod 2 = 0 Then
        szTemp2 = CStr(Int(nCharLong / 2))
    Else
        szTemp2 = CStr(Int(nCharLong / 2) + 0.5)
    End If
    szTemp = szText
    szTemp = StrConv(szTemp, vbFromUnicode)
    If LenB(szTemp) > nCharLong Then
        MsgBox "请输入" & szTemp1 & "个以下<英文字母>或" & szTemp2 & "个以下<汉字>,建议使用<英文字母>.", vbOKOnly + vbInformation, "系统管理"
        TextLongValidate = True
    End If

End Function
'
