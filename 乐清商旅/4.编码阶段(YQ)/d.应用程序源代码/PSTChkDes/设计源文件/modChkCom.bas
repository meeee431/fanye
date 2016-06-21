Attribute VB_Name = "mdlMain"

Option Explicit
'====================================================================
'以下定义常量
'--------------------------------------------------------------------
Public Const MaxLine = 5                       '同时进行的最大检票进程数
Public Const g_cszTitle_Info = "提示"
Public Const g_cszTitle_Warning = "警告"
Public Const g_cszTitle_Error = "错误"
Public Const g_cszTitle_Question = "问题"
Public Const g_cszTitleScollBus = "流水车次"

Public m_rsTicketType As Recordset
Public m_oActiveUser As ActiveUser
'--------------------------------------------------------------------
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
End Enum

Public Enum ECheckStatus
    ECS_CanotCheck = 1      '不能检
    ECS_CanCheck = 2        '能检票
    ECS_CanExtraCheck = 3   '能补检
    ECS_BeChecking = 4      '正在检票
    ECS_BeExtraChecking = 5 '正在补检
    ECS_Checked = 6         '已检
End Enum

Public Enum ETabAct             'TabStrip的行为
    Addone = 1
    CloseOne = 2
End Enum

Public Enum eEventId
    AddBus = 1              '添加车次
    AdjustTime = 2          '调整时间
    ChangeBusCheckGate = 3  '更改检票口
    ChangeBusSeat = 4       '更改座位
    ChangeBusStandCount = 5 '某车次的站票数改变
    ChangeBusTime = 6       '更改车次发车时间
    ChangeParam = 7         '更改参数
    ExStartCheckBus = 8     '补检车次
    MergeBus = 9            '车次并班
    RemoveBus = 10          '删除车次
    ResumeBus = 11          '车次复班
    StartCheckBus = 12      '开检车次
    StopBus = 13            '车次停班
    StopCheckBus = 14       '停检车次
End Enum

Type TEventSound                        '事件音效文件的路径信息结构
    InvalidTicket As String             '无效票
    CanceledTicket As String            '票已废
    ReturnedTicket As String            '票已退
    NoMatchedBus As String              '非当检车次
    CheckedTicket As String             '已检票
    CheckSucess As String               '检票成功
    CheckTimeOn As String               '检票时间已到
    StartupCheckTimeOn As String        '开检时间已到
    FreeTicket As String                '免票提示
    HalfTicket As String                '半票提示
    PreferentialTicket1 As String       '优惠票1提示
    PreferentialTicket2 As String       '优惠票2提示
    PreferentialTicket3 As String       '优惠票3提示
'    sndChanged As String                '正常改签
'    sndNormal As String                 '正常被检
'    sndHasBeChanged As String           '被改签
End Type


Type tCheckInfo
    CheckDate  As Date                              '检票日期
    '检票口信息
    CheckGateNo As String                           '当前检票口
    SellStationID As String                         '当前上车站代码
    SellStationName As String                       '当前上车站名称
    AutoPrint As Boolean                            '是否停检后直接打印路单
    CheckGateName As String                         '当前检票口
    CheckerId As String                             '检票人Id
    Checker As String                               '检票人
    CurrSheetNo As String                           '当前路单
    
    '检票车次信息
    BusID As String                                 '车次号
    EndStationName As String                        '终点站
    StartUpTime As Date                             '发车时间
    StartCheckTime As Date                          '开检时间
    StopCheckTime As Date                           '停检时间
    BusMode As EBusType                             '车次状态
    SellTickets As Integer                          '售票数
    SelfSellStationTickets As Integer            '用户所有上车站的票数
    SeatCount As Integer                            '座位数
    Owner As String                                 '车主
    Company As String                               '参营公司
    MergedBus As String                             '并入车次
    MergeType As Integer
    SplitSeat As Integer
    MergeInSells As Integer                         '并入车次售票数
    VehicleId As String
    Vehicle As String                               '车辆牌照
    VehicleMode As String                           '车辆类型
    SerialNo As Integer                             '车次序号
    CheckSheet As String                            '车次路单号
    '运行车次信息
    RunVehicle As M_TRunVehicle                     '当前运行车辆信息
End Type
Type TTicketInfo
    TicketID As String                              '车票号
    EndStation As String                            '终点站
    TicketStatus As ETicketStatus                   '车票状态
    g_tTicketType As ETicketType                       '车票类型
    TicketDate As Date                              '车票日期
End Type

Type TCheckLineFormInfo
'检票进程表单的基本信息，用于对当前同时执行的多道检票进程进行监控
    BusID As String
    ExCheck As Boolean
    SerialNo As Integer
End Type

Type TWillStopBusStack
    '停检车次堆栈
    Top As Integer
    MsgStyle(1 To MaxLine) As Integer
    ChkLine(1 To MaxLine) As Integer
End Type



'--------------------------------------------------------------------
'以下定义全局变量
'--------------------------------------------------------------------
Public g_oActiveUser As ActiveUser      '当前活动用户
Public g_oChkTicket As CheckTicket   '当前检票对象
Public g_oEnvBus As REBus         '当前环境车次
Public g_tCheckInfo As tCheckInfo   '当前系统的检票活动变量
Public g_tEventSoundPath As TEventSound                 '事件音效文件路径
Public g_cWillCheckBusList As BusCollection            '当天的车次列表集合
Public g_cCheckedBusList As BusCollection           '当天的车次列表集合
Public g_atCheckLine(1 To MaxLine) As TCheckLineFormInfo
Public g_aofrmCheckForm(1 To MaxLine) As frmCheckTicket

    '系统参数全局变量
Public g_nLatestExtraCheckTime As Integer
Public g_nBeginCheckTime As Integer
Public g_nExtraCheckTime As Integer
Public g_nCheckTicketTime As Integer
Public g_bAllowChangeRide As Boolean
Public g_szUnitID As String
Public g_szUnitName As String
Public g_tTicketType() As TTicketType
Public g_nCheckSheetLen As Integer
Public g_nCurrLineIndex As Integer                      '当前处于哪一个检票车次进程

Public g_bAllowStartChectNotRearchTime As Boolean '是否允许未到开检时间开检


Public g_szSellStationName As String
'--------------------------------------------------------------------
'以下定义本模块变量
'--------------------------------------------------------------------


'检票公共模块


'以下常量定义
'*************************************************************
'*      此处定义一些临时常量
'Public Const AheadTime = 0.0069                '预定义的提前开检时间
'Public Const cntCheckTime = 10                 '预定义的检票时间
Public Const m_cRegSystemKey = "ChkDes\CheckEnviroment"         '检票系统主键字符串
Public Const m_cRegSoundKey = "ChkDes\EventSound"      '检票音效主键字符串
Public Const m_cnTimeWindage = 1            '时间偏差数，用于协调下一班车次的时间（以分钟为单位）

Public szSeatBusID As String

'***********************************************************

'以下枚举定义


'Public m_sgAheadTime As Single                 '预定义的提前开检时间(以小时为单位)
'Public m_sgCheckTime As Single                  '预定义的检票时间(以分钟为单位)
Public m_szPrnFmtFile As String                  '路单打印格式文件的路径

Public m_dtAheadTime As Date                            '检票的提前时间
Public g_oNextEnvBus As REBus



'Public m_nPrevLineIndex As Integer                      '前一个使用的检票车次进程
Public g_szTitle As String
Public m_bIsFormActive As Boolean                       '是否是先激活了窗体
Public m_bCloseOne As Boolean                           '是否关闭了窗口
Public m_lErrorCode As Long                             '错误号

'以下声明一些需使用的API函数
Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'以下声明API函数用到的常数定义
Public Const SND_NOSTOP = &H10        '  don't stop any currently playing sound
Public Const SND_FILENAME = &H20000     '  name is a file name
Public Const SND_ASYNC = &H1         '  play asynchronously

Public Const HWND_TOPMOST = -1

'得到当前的检票进程数
Public Function CheckLineCount() As Integer
    If MDIMain.tbsBusList.Visible Then
        CheckLineCount = MDIMain.tbsBusList.Tabs.Count
    Else
        CheckLineCount = 0
    End If
End Function
'添加新的检票进程
Public Sub AddNewCheckLine(BusID As String, Optional ExCheck As Boolean = False, Optional IsChecking As Boolean = False, Optional szVehicleId As String, Optional oREBus As REBus)
'    ' 参数注释
'    ' *************************************
'    ' BusID:车次代码
'    ' ExCheck:可选。是否补检
'    ' IsChecking:可选。是否正在检票，（主要用于处理系统异常中断时起恢复作用）
'    ' szVehicleId:可选。滚动车次的运行车辆号
'    ' *************************************
On Error GoTo ErrHandle
    Dim nPrevLineIndex As Integer   'm_nCurrLineIdex的初值
    Dim nCheckLineCount As Integer
    Dim nErrorNum As Long
    
    nCheckLineCount = CheckLineCount
    If nCheckLineCount + 1 > MaxLine Then
        MsgboxEx "本系统最多支持同时进行" & Str(MaxLine) & "道检票进程！！！", vbInformation, g_cszTitle_Info
        Exit Sub
    End If
    nPrevLineIndex = g_nCurrLineIndex
    g_nCurrLineIndex = nCheckLineCount + 1
    
    
    ShowSBInfo "正在初始化检票窗体..."
    
    ResetEnvBusInfo BusID, nErrorNum, oREBus
    If nErrorNum <> 0 Then GoTo ErrHandle
    g_atCheckLine(g_nCurrLineIndex).BusID = g_tCheckInfo.BusID
    g_atCheckLine(g_nCurrLineIndex).ExCheck = ExCheck
    g_atCheckLine(g_nCurrLineIndex).SerialNo = g_tCheckInfo.SerialNo
    
    If Not IsChecking Then
        '以下调用中间层方法开检，并更改当天该车次的车次状态
        If ExCheck Then
            g_oChkTicket.ExtraStartCheckBus g_tCheckInfo.BusID, g_tCheckInfo.SerialNo
        Else
            If g_tCheckInfo.BusMode = TP_ScrollBus Then
                g_oChkTicket.StartCheckScrollBus g_tCheckInfo.BusID, g_tCheckInfo.SerialNo, szVehicleId
            Else
                If szVehicleId = "" Then
                    g_oChkTicket.StartCheckRegularBus g_tCheckInfo.BusID
                Else
                    g_oChkTicket.StartCheckRegularBus g_tCheckInfo.BusID, szVehicleId
                End If
            End If
        End If
        
        Dim nIndex As Integer           '更改缓冲区车次状态
        Dim tTmpBusInfo As tCheckBusLstInfo
        nIndex = g_cWillCheckBusList.FindItem(g_tCheckInfo.BusID)
        If nIndex > 0 Then
            tTmpBusInfo = g_cWillCheckBusList.Item(nIndex)
            tTmpBusInfo.Status = IIf(ExCheck, EREBusStatus.ST_BusExtraChecking, EREBusStatus.ST_BusChecking)
            g_cWillCheckBusList.UpdateOne tTmpBusInfo
            If frmBusList.IsShow Then frmBusList.UpdateWillCheckBusItem 2, tTmpBusInfo      '更改一行
        End If
    End If
    
    nCheckLineCount = nCheckLineCount + 1
    Dim szTabsString As String
    szTabsString = g_tCheckInfo.BusID & IIf(g_tCheckInfo.BusMode = TP_ScrollBus, "-" & g_tCheckInfo.SerialNo, "") & _
                                             g_tCheckInfo.EndStationName & "(&" & nCheckLineCount & ")"
    With MDIMain                   '设置检票车次进程标签
    If .tbsBusList.Visible = False Then
        .tbsBusList.Tabs(1).Caption = szTabsString
        .tbsBusList.Visible = True
    Else
        .tbsBusList.Tabs.Add g_nCurrLineIndex, , szTabsString
    End If
    End With
    
    '显示检票窗口
    Dim ofrmCheckTicket As New frmCheckTicket
    Set g_aofrmCheckForm(g_nCurrLineIndex) = ofrmCheckTicket
    Set g_aofrmCheckForm(g_nCurrLineIndex).m_oREBus = g_oEnvBus
    g_aofrmCheckForm(g_nCurrLineIndex).Show
    
    MDIMain.tbsBusList.Tabs(g_nCurrLineIndex).Selected = True
    ShowSBInfo ""
    Exit Sub
ErrHandle:
    If err.Number = ERR_ChkTkBusAlreadyExist Then
         '检票车次已存在时继续执行，主要用于非正常退出检票时恢复
        Resume Next
    Else
        ShowErrorMsg
    End If
    g_nCurrLineIndex = nPrevLineIndex   '返回初始值
    ShowSBInfo ""
End Sub
'关闭某一个的检票进程
Public Sub CloseOneCheckLine(nWhichOne As Integer)
    Dim i As Integer
    Dim nCheckLineCount As Integer
    nCheckLineCount = CheckLineCount
'    For i = g_nCurrLineIndex To nCheckLineCount - 1
    For i = nWhichOne To nCheckLineCount - 1
        g_atCheckLine(i).BusID = g_atCheckLine(i + 1).BusID
        g_atCheckLine(i).ExCheck = g_atCheckLine(i + 1).ExCheck
        Set g_aofrmCheckForm(i) = g_aofrmCheckForm(i + 1)
        MDIMain.tbsBusList.Tabs(i).Caption = _
            Left(MDIMain.tbsBusList.Tabs(i + 1).Caption, _
            Len(MDIMain.tbsBusList.Tabs(i + 1).Caption) - 4) _
            & "(&" & Trim(Str(i)) & ")"
        g_aofrmCheckForm(i).Tag = Str(i)
    Next i
    Set g_aofrmCheckForm(nCheckLineCount) = Nothing
    If nCheckLineCount = 1 Then
        g_nCurrLineIndex = 0
        MDIMain.tbsBusList.Visible = False
'        MDIMain.mnu_Query_Ticket.Enabled = False
'        MDIMain.mnu_Check_Bus.Enabled = False
    Else
        If g_nCurrLineIndex >= nWhichOne And g_nCurrLineIndex <> 1 Then g_nCurrLineIndex = g_nCurrLineIndex - 1
        MDIMain.tbsBusList.Tabs.Remove nCheckLineCount
        MDIMain.tbsBusList.Tabs.Item(g_nCurrLineIndex).Selected = True
    End If
End Sub
'根据车次号获取最新车次信息，放入系统活动变量g_tCheckInfo中，如有错误将错误返回
Public Sub ResetEnvBusInfo(szBusid As String, Optional ByRef ErrorCode As Long, Optional oREBus As REBus)
'    ' 参数注释
'    ' *************************************
'    ' szBusID:车次Id
'    ' ErrorCode:错误号,可选
'    ' *************************************
'    ' ****************************************************************
'    ' 把结果放入g_tCheckInfo活动变量中，如有错误将错误返回
'    ' ****************************************************************

On Error GoTo ErrHandle
    If oREBus Is Nothing Then
        g_oEnvBus.Identify szBusid, Date, g_tCheckInfo.CheckGateNo
    Else
        Set g_oEnvBus = oREBus
    End If
    g_tCheckInfo.BusID = UCase(Trim(g_oEnvBus.BusID))
    g_tCheckInfo.BusMode = g_oEnvBus.BusType
    g_tCheckInfo.Company = g_oEnvBus.CompanyName
    g_tCheckInfo.Owner = g_oEnvBus.OwnerName
    g_tCheckInfo.EndStationName = g_oEnvBus.EndStationName
    If g_oEnvBus.BusType <> TP_ScrollBus Then
            g_tCheckInfo.MergedBus = Trim(g_oEnvBus.BeMergedBus.szBusid)
            g_tCheckInfo.MergeType = g_oEnvBus.BeMergedBus.nMergeType
    End If
'    If g_oEnvBus.BusType = TP_ScrollBus Then
        
'    Else
    g_tCheckInfo.SeatCount = g_oEnvBus.TotalSeat
'    End If
    g_tCheckInfo.StartUpTime = g_oEnvBus.StartUpTime
    g_tCheckInfo.StartCheckTime = g_oEnvBus.StartCheckTime
    g_tCheckInfo.StopCheckTime = g_oEnvBus.StopCheckTime
    g_tCheckInfo.CheckSheet = Trim(g_oEnvBus.CheckSheet)
    g_tCheckInfo.VehicleId = g_oEnvBus.Vehicle
    g_tCheckInfo.Vehicle = g_oEnvBus.VehicleTag
    g_tCheckInfo.VehicleMode = g_oEnvBus.VehicleModelName
    ErrorCode = 0
    Exit Sub
ErrHandle:
    ShowErrorMsg
    ErrorCode = err.Number
End Sub
Public Sub PlayEventSound(szFileName As String)
    '播放音效
    PlaySound szFileName, 0, SND_FILENAME + SND_ASYNC
    
End Sub
'Public Sub ShowErrorMsg()
'    MsgboxEx err.Description, vbExclamation, "错误-" & err.Number
'End Sub
Public Sub Main()
    Dim oCommShell As CommShell
    
    If App.PrevInstance Then
        MsgBox "系统已运行!", vbExclamation, "警告"
        End
    End If
On Error GoTo ErrHandle
    Set oCommShell = New CommShell

TryLogin:
    Set g_oActiveUser = oCommShell.ShowLogin()
    If g_oActiveUser Is Nothing Then Exit Sub
    
    
'    m_szPrnFmtFile = App.Path & "\ChkSheet.bpf"
'    App.HelpFile = SetHTMLHelpStrings("SNChkSys.chm") '设定App.HelpFile
    
    InitSystemParam
    GetIniFile
    
    If g_tCheckInfo.CheckGateNo = "" Then
        MsgBox "第一次使用检票台,请指定当前检票口!", vbInformation, g_cszTitle_Info
        frmSetOption.Show vbModal
    End If
    If g_tCheckInfo.CheckGateNo <> "" Then
        frmChangeSheetNo.FirstLoad = True
        frmChangeSheetNo.Show vbModal
    End If
    
    
    '启动主窗体
'    oCommShell.ShowSplash "检票系统", "Check Ticket Desktop", LoadResPicture(101, 0), App.Major, App.Minor, App.Revision
    DoEvents
    
    '得到上车站代码
    Dim oCheckGate As New CheckGate
    oCheckGate.Init g_oActiveUser
    oCheckGate.Identify g_tCheckInfo.CheckGateNo
    g_tCheckInfo.SellStationID = oCheckGate.SellStationID
    g_tCheckInfo.SellStationName = oCheckGate.SellStationName
    
    g_szSellStationName = oCheckGate.SellStationName
    
    '初始化全局变量
    g_oChkTicket.CheckGateNo = g_tCheckInfo.CheckGateNo
    Set g_oEnvBus = New REBus
    g_oEnvBus.Init g_oActiveUser
    SetHTMLHelpStrings "STChkDes.chm"
    
'    Dim oPriceMan As New STPrice.TicketPriceMan
'    oPriceMan.Init m_oActiveUser
    
    
    
    Load MDIMain
    EvisibleCloseButton MDIMain
    MDIMain.Show
    oCommShell.CloseSplash
    Exit Sub
ErrHandle:
    If err.Number = 500 Then        '无注册表项
        SaveRegInitInfo
        Resume
    Else
        ShowErrorMsg
        Resume TryLogin
    End If
End Sub
'关闭当前正在显示的Modal窗口,窗口Tag标识为Modal的
Public Sub CloseModalForm()
    '利用Modal窗口的tag属性
On Error GoTo ErrHandle
    Do
        If Screen.ActiveForm.Tag <> "Modal" Then
            Exit Do
        End If
        If Screen.ActiveForm Is frmCheckSheet Then
            Exit Do
        Else
            Unload Screen.ActiveForm    '路单打印界面激活时，不要关闭，否则路单号会跳号
        End If
    Loop
    Exit Sub
ErrHandle:
    On Error Resume Next
End Sub
Public Sub WriteCheckGateInfo()
'得到检票口状态,并写入界面
    With MDIMain
        .lblChecker.Caption = g_tCheckInfo.Checker
        .lblCheckGate.Caption = g_tCheckInfo.CheckGateName
        .lblCurrentSheetNo.Caption = g_tCheckInfo.CurrSheetNo
        .moMessage.SellStation = g_tCheckInfo.SellStationID
    End With
End Sub
Public Sub WriteNextBus()
'得到待检车次,并写入界面
    Dim lHaveTime As Double
    Dim dtTmp As Date
    Dim dtStartUpTime As Date
    Dim dtStopCheckTime As Date
    
    On Error Resume Next
    
    With MDIMain
        Set g_oNextEnvBus = g_oChkTicket.GetNextCheckBus
        
        If Not (g_oNextEnvBus Is Nothing) Then
            dtStartUpTime = g_oNextEnvBus.StartUpTime
            dtStopCheckTime = g_oNextEnvBus.StopCheckTime
            .lblBusID.Caption = g_oNextEnvBus.BusID
            .lblStartupTime.Caption = Format(dtStartUpTime, "HH:MM:SS")
            .lblEndStation.Caption = g_oNextEnvBus.EndStationName
            .lblCompany.Caption = g_oNextEnvBus.CompanyName
            .lblOwner.Caption = g_oNextEnvBus.OwnerName
            .lblLicense.Caption = g_oNextEnvBus.Vehicle
            '********Need change******
            dtTmp = Now
            lHaveTime = DateDiff("s", dtTmp, DateAdd("n", -g_nBeginCheckTime, dtStartUpTime))
            
            lHaveTime = IIf(lHaveTime > 0, lHaveTime, 0)
            .rvtTime.Second = lHaveTime
            EnabledMDITimer True
            
            '将RevTimer1置有效，用于实时刷新下一待检车次（跳到没有开检的车次）
            '触发时间为下一待检车次发车时间后的m_cnTimeWindage分钟
            lHaveTime = DateDiff("s", dtTmp, _
                dtStartUpTime)
            lHaveTime = lHaveTime + 60 * m_cnTimeWindage
            lHaveTime = IIf(lHaveTime > 0, lHaveTime, 0)
            .RevTimer1.Second = lHaveTime
            .RevTimer1.Enabled = True
        Else
            .lblBusID.Caption = ""
            .lblStartupTime.Caption = ""
            .lblEndStation.Caption = ""
            .lblCompany.Caption = ""
            .lblOwner.Caption = ""
            .lblLicense.Caption = ""
            .rvtTime.Second = 0
'            .rvtTime.Enabled = False
            EnabledMDITimer False
        End If
    End With
    Exit Sub
End Sub
Public Sub WriteInitReg()
'保存当前路单号入注册表
    Dim oReg As New CFreeReg
    oReg.Init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany
    oReg.SaveSetting m_cRegSystemKey, "SheetNo", g_tCheckInfo.CurrSheetNo
    Set oReg = Nothing
End Sub


Public Sub EnabledMDITimer(bEnabled As Boolean)
'设置主界面的待检车次倒计时钟
    If bEnabled Then
        MDIMain.flblrevTime.Visible = False
        MDIMain.rvtTime.Enabled = True
        MDIMain.rvtTime.Visible = True
    Else
        MDIMain.flblrevTime.Visible = True
        MDIMain.rvtTime.Enabled = False
        MDIMain.rvtTime.Visible = False
    End If
End Sub
Public Function MsgboxEx(Optional Prompt As String = "", Optional Buttons As VbMsgBoxStyle = vbOKOnly, Optional Title As String = "", Optional showMode As FormShowConstants = vbModal) As VbMsgBoxResult
'自定义msgbox函数
    Dim ofrm As New frmMsgbox
    ofrm.Prompt = Prompt
    ofrm.Buttons = Buttons
    ofrm.Title = Title
    ofrm.Show showMode
    MsgboxEx = ofrm.Result
'    MsgboxEx = meMsgboxEx
End Function

'信息处理过程                              *

Public Sub RunMsgEvent(EventMode As eEventId, EventParam() As String)
'    ' 参数注释
'    ' *************************************
'    ' EventMode:消息类型（消息Id）
'    ' EventParam:参数数组
'    ' *************************************

    Dim nTmp As Integer
    Dim tTmpBusInfo As tCheckBusLstInfo
    Select Case EventMode
        Case eEventId.AddBus
            If Trim(EventParam(3)) = Trim(g_tCheckInfo.CheckGateNo) Then
                BuildBusCollection
                If frmBusList.IsShow Then
                    frmBusList.RefreshBus
                End If
            End If
        Case eEventId.AdjustTime
        Case eEventId.ChangeBusCheckGate
            If Trim(EventParam(3)) = Trim(g_tCheckInfo.CheckGateNo) Then
                BuildBusCollection
            Else
                nTmp = g_cWillCheckBusList.FindItem(EventParam(1))
                If nTmp > 0 Then
                    tTmpBusInfo = g_cWillCheckBusList.Item(nTmp)
                    g_cWillCheckBusList.RemoveOne nTmp
                Else
                    Exit Sub
                End If
            End If
            '刷新检票车次列表窗口
            If frmBusList.IsShow Then
                frmBusList.RefreshBus
            End If
            
            '刷新开检窗口
            '******************
            
            '刷新下一车次
            If Not g_oNextEnvBus Is Nothing Then
                If g_oNextEnvBus.BusID = EventParam(1) Then
                    WriteNextBus
                End If
            End If
        Case eEventId.ChangeBusSeat
                                                        
        Case eEventId.ChangeBusStandCount
        Case eEventId.ChangeBusTime
            nTmp = g_cWillCheckBusList.FindItem(EventParam(1))
            If nTmp > 0 Then
                tTmpBusInfo = g_cWillCheckBusList.Item(nTmp)
                tTmpBusInfo.StartUpTime = Format(EventParam(3), cszDateTimeStr)
                g_cWillCheckBusList.UpdateOne tTmpBusInfo
            Else
                Exit Sub
            End If
            
            '刷新检票车次列表窗口
            If frmBusList.IsShow Then
                frmBusList.RefreshBus
            End If
            
            '刷新开检窗口
            '******************
            
            '刷新下一车次
            If Not g_oNextEnvBus Is Nothing Then
                If g_oNextEnvBus.BusID = EventParam(1) Then
                    WriteNextBus
                End If
            End If
        Case eEventId.ChangeParam
        Case eEventId.MergeBus
            nTmp = g_cWillCheckBusList.FindItem(EventParam(1))
            If nTmp > 0 Then
                tTmpBusInfo = g_cWillCheckBusList.Item(nTmp)
                tTmpBusInfo.Status = EREBusStatus.ST_BusSlitpStop
                g_cWillCheckBusList.UpdateOne tTmpBusInfo
            Else
                Exit Sub
            End If
            
            '刷新检票车次列表窗口
            If frmBusList.IsShow Then
                frmBusList.RefreshBus
            End If
            
            '刷新开检窗口
            '******************
            
            '刷新下一车次
            If Not g_oNextEnvBus Is Nothing Then
                If g_oNextEnvBus.BusID = EventParam(1) Then
                    WriteNextBus
                End If
            End If
        Case eEventId.RemoveBus
            nTmp = g_cWillCheckBusList.FindItem(EventParam(1))
            If nTmp > 0 Then
                tTmpBusInfo = g_cWillCheckBusList.Item(nTmp)
                g_cWillCheckBusList.RemoveOne nTmp
            Else
                Exit Sub
            End If
            
            '刷新检票车次列表窗口
            If frmBusList.IsShow Then
                frmBusList.RefreshBus
            End If
            
            '刷新开检窗口
            '******************
            
            '刷新下一车次
            If Not g_oNextEnvBus Is Nothing Then
                If g_oNextEnvBus.BusID = EventParam(1) Then
                    WriteNextBus
                End If
            End If
        Case eEventId.ResumeBus
            nTmp = g_cWillCheckBusList.FindItem(EventParam(1))
            If nTmp > 0 Then
                tTmpBusInfo = g_cWillCheckBusList.Item(nTmp)
                tTmpBusInfo.Status = EREBusStatus.ST_BusNormal
                g_cWillCheckBusList.UpdateOne tTmpBusInfo
            Else
                Exit Sub
            End If
            
            '刷新检票车次列表窗口
            If frmBusList.IsShow Then
                frmBusList.RefreshBus
            End If
            
            '刷新开检窗口
            '******************
            
            '刷新下一车次
            If Not g_oNextEnvBus Is Nothing Then
                If g_oNextEnvBus.BusID = EventParam(1) Then
                   WriteNextBus
                Else
                   If tTmpBusInfo.StartUpTime < g_oNextEnvBus.StartUpTime Then
                     WriteNextBus
                   End If
                End If
            End If
        Case eEventId.StopBus
            nTmp = g_cWillCheckBusList.FindItem(EventParam(1))
            If nTmp > 0 Then
                tTmpBusInfo = g_cWillCheckBusList.Item(nTmp)
                tTmpBusInfo.Status = EREBusStatus.ST_BusStopped
                g_cWillCheckBusList.UpdateOne tTmpBusInfo
            Else
                Exit Sub
            End If
            
            '刷新检票车次列表窗口
            If frmBusList.IsShow Then
                frmBusList.RefreshBus
            End If
            
            '刷新开检窗口
            '******************
            
            '刷新下一车次
            If Not g_oNextEnvBus Is Nothing Then
                If g_oNextEnvBus.BusID = EventParam(1) Then
                    WriteNextBus
                End If
            End If
    End Select
End Sub
Public Sub BuildBusCollection()
    '取得当天的检票车次列表信息，存入g_cWillCheckBusList
    Dim atCheckSheetInfo() As tCheckBusLstInfo
    Dim szBusid As String, szCheckedBusID As String
    Dim nTmpIndex As Integer
    Dim i As Integer, nCount As Integer
    Dim j As Integer, l As Integer, n As Integer
    
    If g_cWillCheckBusList Is Nothing Then Set g_cWillCheckBusList = New BusCollection
    If g_cCheckedBusList Is Nothing Then Set g_cCheckedBusList = New BusCollection
    g_cWillCheckBusList.RemoveAll
    g_cCheckedBusList.RemoveAll
    
    '得到当天所有车次信息
    Dim rsBusInfo As Recordset
    Set rsBusInfo = g_oChkTicket.GetBusInfoRs(Date, g_tCheckInfo.CheckGateNo)
    
    '得到已检车次及路单信息
    Dim rsCheckedBus As Recordset
    Set rsCheckedBus = g_oChkTicket.GetBusCheckSheetRs(Date, g_tCheckInfo.CheckGateNo)
    
    Dim tTmpBusInfo As tCheckBusLstInfo '车次列表元素临时变量
    Dim tTmpBusInfo2 As tCheckBusLstInfo '车次列表元素临时变量
'    rsBusInfo.MoveFirst: rsCheckedBus.MoveFirst
'    Dim bStack As Boolean       '当True表示rsBusInfo未移动
    Do While Not rsBusInfo.EOF Or Not rsCheckedBus.EOF
        '以下判断是否有重复车次，将同一滚动车次并为一条记录
        If Not rsBusInfo.EOF Then
            szBusid = UCase(Trim(rsBusInfo("bus_id")))
            tTmpBusInfo.BusID = szBusid
            tTmpBusInfo.BusMode = rsBusInfo("bus_type")
            tTmpBusInfo.Company = rsBusInfo("transport_company_short_name")
            tTmpBusInfo.Vehicle = rsBusInfo("license_tag_no")
            tTmpBusInfo.StartUpTime = rsBusInfo("bus_start_time")
            tTmpBusInfo.EndStationName = rsBusInfo("end_station_name")
            tTmpBusInfo.Owner = rsBusInfo("owner_name")
            tTmpBusInfo.Status = rsBusInfo("status")
            tTmpBusInfo.BusSerial = 0
        Else
            szBusid = ""
        End If
        If Not rsCheckedBus.EOF Then
            szCheckedBusID = UCase(Trim(rsCheckedBus("bus_id")))
            tTmpBusInfo2.BusID = szCheckedBusID
            tTmpBusInfo2.BusSerial = rsCheckedBus("bus_serial_no")
            tTmpBusInfo2.BusMode = rsCheckedBus("bus_type")
            tTmpBusInfo2.Company = rsCheckedBus("transport_company_short_name")
            tTmpBusInfo2.Vehicle = rsCheckedBus("license_tag_no")
            tTmpBusInfo2.StartUpTime = rsCheckedBus("bus_start_time")
            tTmpBusInfo2.EndStationName = rsCheckedBus("end_station_name")
            tTmpBusInfo2.StartChkTime = rsCheckedBus("check_start_time")
            tTmpBusInfo2.StopChkTime = rsCheckedBus("check_end_time")
            tTmpBusInfo2.Owner = rsCheckedBus("owner_name")
            tTmpBusInfo2.Status = EREBusStatus.ST_BusStopCheck
            tTmpBusInfo2.CheckSheet = Trim(rsCheckedBus("check_sheet_id"))
        Else
            szCheckedBusID = ""
        End If
        '滚动车次既放在待检车次集合中，又放入已检车次集合中
        '固定车次要么在待检车次中，要么在已检车次中
        
        If szBusid = szCheckedBusID Then
            If szCheckedBusID <> "" Then
                g_cCheckedBusList.Addone tTmpBusInfo2
            End If
            If Not rsCheckedBus.EOF Then rsCheckedBus.MoveNext
            If tTmpBusInfo.BusMode = TP_ScrollBus Then
                g_cWillCheckBusList.Addone tTmpBusInfo
            End If
            rsBusInfo.MoveNext
        Else
            If szBusid <> "" Then
                g_cWillCheckBusList.Addone tTmpBusInfo
                rsBusInfo.MoveNext
            Else
                g_cCheckedBusList.Addone tTmpBusInfo2
                rsCheckedBus.MoveNext
            End If
        End If
    Loop
End Sub


Public Function GetCodeStr(szSource As String, nlen As Integer) As String
'按要求长度生成数字串
    Dim szNum As String
    Dim nZeroNum As Integer
    
    szNum = Trim(Str(Val(szSource)))
    nZeroNum = nlen - Len(szNum)
    If nZeroNum < 0 Then nZeroNum = 0
    GetCodeStr = String(nZeroNum, "0") & Left(szNum, nlen - nZeroNum)
End Function
'设置注册项初始信息
Private Sub SaveRegInitInfo()
On Error GoTo ErrHandle
    Dim oReg As New CFreeReg
    oReg.Init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany
    oReg.SaveSetting m_cRegSystemKey, "CheckGate", ""
    oReg.SaveSetting m_cRegSystemKey, "SheetNo", ""
    oReg.SaveSetting m_cRegSystemKey, "AutoPrint", ""
    
    oReg.SaveSetting m_cRegSoundKey, "CanceledTicket", ""
    oReg.SaveSetting m_cRegSoundKey, "CheckedTicket", ""
    oReg.SaveSetting m_cRegSoundKey, "CheckSucess", ""
    oReg.SaveSetting m_cRegSoundKey, "CheckTimeOn", ""
    oReg.SaveSetting m_cRegSoundKey, "InvalidTicket", ""
    oReg.SaveSetting m_cRegSoundKey, "NoMatchedBus", ""
    oReg.SaveSetting m_cRegSoundKey, "ReturnedTicket", ""
    oReg.SaveSetting m_cRegSoundKey, "StartupCheckTimeOn", ""
    oReg.SaveSetting m_cRegSoundKey, "FreeTicket", ""
    oReg.SaveSetting m_cRegSoundKey, "HalfTicket", ""
    oReg.SaveSetting m_cRegSoundKey, "PreferentialTicket1", ""
    oReg.SaveSetting m_cRegSoundKey, "PreferentialTicket2", ""
    oReg.SaveSetting m_cRegSoundKey, "PreferentialTicket3", ""
    Set oReg = Nothing
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

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
                .abMenu.Bands("statusBar").Tools("pnUserInfo").Caption = pszInfo
            Case EStatusBarArea.ESB_LoginTime
                If pszInfo <> "" Then pszInfo = "登录时间: " & pszInfo
                .abMenu.Bands("statusBar").Tools("pnLoginTime").Caption = pszInfo
        End Select
    End With
End Sub
'' *******************************************************************
'' *   Member Name: WriteProcessBar                                  *
'' *   Brief Description: 写系统进程条状态                           *
'' *   Engineer: 陆勇庆                                              *
'' *******************************************************************
'Public Sub WriteProcessBar(Optional pbVisual As Boolean = True, Optional ByVal plCurrValue As Variant = 0, Optional ByVal plMaxValue As Variant = 0, Optional pszShowInfo As String = cszUnrepeatString)
''参数注释
''*************************************
''plCurrValue(当前进度值)
''plMaxValue(最大进度值)
''*************************************
'    If plMaxValue = 0 And pbVisual = True Then Exit Sub
'    Dim nCurrProcess As Integer
'    With MDIMain.abMenu.Bands("statusBar").Tools("progressBar")
'        If pbVisual Then
'            If Not .Visible Then
'                .Visible = True
'                MDIMain.pbLoad.Max = 100
'            End If
'            nCurrProcess = Int(plCurrValue / plMaxValue * 100)
'            MDIMain.pbLoad.Value = nCurrProcess
'        Else
'            .Visible = False
'        End If
'    End With
'    If pszShowInfo <> cszUnrepeatString Then ShowSBInfo pszShowInfo, ESB_WorkingInfo
'End Sub
'设置初始化参数

Private Sub InitSystemParam()
    Dim oSystemParam As New SystemParam
    oSystemParam.Init g_oActiveUser
    '校正本地时间
    Date = oSystemParam.NowDate
    Time = oSystemParam.NowTime
    
    '读取初始化系统参数
    g_nBeginCheckTime = oSystemParam.BeginCheckTime
    g_nLatestExtraCheckTime = oSystemParam.LatestExtraCheckTime
    g_nExtraCheckTime = oSystemParam.ExtraCheckTime
    g_nCheckTicketTime = oSystemParam.CheckTicketTime
    g_szTitle = oSystemParam.RoadSheetTitle
    g_bAllowChangeRide = oSystemParam.AllowChangeRide
    g_szUnitID = oSystemParam.UnitID
    g_nCheckSheetLen = oSystemParam.CheckSheetLen
    g_tTicketType = oSystemParam.GetAllTicketType(1, True)
    
    g_bAllowStartChectNotRearchTime = oSystemParam.AllowStartChectNotRearchTime '是否允许未到开检时间开检
    
    
    
    Set m_rsTicketType = oSystemParam.GetAllTicketTypeRS(TP_TicketTypeValid)

    Set oSystemParam = Nothing


    '以下初始注册表设置
    Dim oReg As CFreeReg
    Set oReg = New CFreeReg
    
    Set g_oChkTicket = New CheckTicket
    g_oChkTicket.Init g_oActiveUser
    
    oReg.Init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany
    g_tCheckInfo.CheckGateNo = Trim(oReg.GetSetting(m_cRegSystemKey, "CheckGate"))
    g_tCheckInfo.CurrSheetNo = Format(Val(g_oChkTicket.GetLastCheckSheetID(g_oActiveUser.UserID)) + 1, String(g_nCheckSheetLen, "0"))  'Trim(oReg.GetSetting(m_cRegSystemKey, "SheetNo"))
    g_tCheckInfo.AutoPrint = IIf(Val(oReg.GetSetting(m_cRegSystemKey, "AutoPrint")) <> 0, True, False)
    g_tCheckInfo.CheckerId = g_oActiveUser.UserID
    g_tCheckInfo.Checker = g_oActiveUser.UserName
    g_tCheckInfo.CheckDate = Date


    '读取检票时各种音效文件路径
    g_tEventSoundPath.CanceledTicket = oReg.GetSetting(m_cRegSoundKey, "CanceledTicket")
    g_tEventSoundPath.CheckedTicket = oReg.GetSetting(m_cRegSoundKey, "CheckedTicket")
    g_tEventSoundPath.CheckSucess = oReg.GetSetting(m_cRegSoundKey, "CheckSucess")
    g_tEventSoundPath.CheckTimeOn = oReg.GetSetting(m_cRegSoundKey, "CheckTimeOn")
    g_tEventSoundPath.InvalidTicket = oReg.GetSetting(m_cRegSoundKey, "InvalidTicket")
    g_tEventSoundPath.NoMatchedBus = oReg.GetSetting(m_cRegSoundKey, "NoMatchedBus")
    g_tEventSoundPath.ReturnedTicket = oReg.GetSetting(m_cRegSoundKey, "ReturnedTicket")
    g_tEventSoundPath.StartupCheckTimeOn = oReg.GetSetting(m_cRegSoundKey, "StartupCheckTimeOn")
    g_tEventSoundPath.FreeTicket = oReg.GetSetting(m_cRegSoundKey, "FreeTicket")
    g_tEventSoundPath.HalfTicket = oReg.GetSetting(m_cRegSoundKey, "HalfTicket")
    g_tEventSoundPath.PreferentialTicket1 = oReg.GetSetting(m_cRegSoundKey, " PreferentialTicket1")
    g_tEventSoundPath.PreferentialTicket2 = oReg.GetSetting(m_cRegSoundKey, " PreferentialTicket2")
    g_tEventSoundPath.PreferentialTicket3 = oReg.GetSetting(m_cRegSoundKey, " PreferentialTicket3")
    
    Set oReg = Nothing
End Sub
Public Function GetStatusString(nStatus As Integer) As String
    Select Case nStatus
        Case EREBusStatus.ST_BusChecking
            GetStatusString = "正在检票"
        Case EREBusStatus.ST_BusExtraChecking
            GetStatusString = "正在补检"
        Case EREBusStatus.ST_BusMergeStopped, EREBusStatus.ST_BusSlitpStop
            GetStatusString = "并班停检"
        Case EREBusStatus.ST_BusNormal, EREBusStatus.ST_BusReplace
            GetStatusString = "未检"
        Case EREBusStatus.ST_BusStopCheck
            GetStatusString = "停检"
        Case EREBusStatus.ST_BusExtraChecking
            GetStatusString = "正在补检"
        Case EREBusStatus.ST_BusStopped
            GetStatusString = "车次停班"
    End Select
End Function
Public Function getCheckedTicketStatus(nStatus As Integer) As String
    Select Case nStatus
        Case ECheckedTicketStatus.NormalTicket
            getCheckedTicketStatus = "正常检入"
        Case ECheckedTicketStatus.ChangedTicket
            getCheckedTicketStatus = "改乘检入"
        Case ECheckedTicketStatus.MergedTicket
            getCheckedTicketStatus = "并班检入"
    End Select
End Function
