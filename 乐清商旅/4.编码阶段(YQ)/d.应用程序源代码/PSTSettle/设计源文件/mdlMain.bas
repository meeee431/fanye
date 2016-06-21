Attribute VB_Name = "mdlMain"
Option Explicit

Const cszDocumentDir = "DocumentDir"
Const cszDefDocumentDir = "C:\"
Const cszLuggageAccount = "LugAcc"
Const cszRecentSeller = "RecentSeller"

Public Const cszCompanyName = "参运公司"
Public Const cszVehicleName = "车辆"
Public Const cszBusName = "车次"

Public Const m_cRegSystemKey = "STSettle"         '结算系统主键字符串
Public g_oActiveUser As ActiveUser
Public g_oParam As New SystemParam
Public g_szUnitID As String
Public m_adbSplitItem() As Double   '用于修改时调用
'Public m_IsSave As Boolean
Public g_szFixFeeItem As String '结算项中的固定费用项.
Public g_bAllowSellteTotalNegative As Boolean '是否将上月结算的负数算到下月
Public g_szSettleNegativeSplitItem As String '将上月结算的负数项放在结算项的位置
Public Const g_cnSplitItemCount = 20
Public g_szIsFixFeeUpdateEachMonth As Boolean '是否固定费用是每个月更新的


'Public g_tEventSoundPath as TEventSound                 '事件音效文件路径

'以下声明一些需使用的API函数
Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'以下声明API函数用到的常数定义
Public Const SND_NOSTOP = &H10        '  don't stop any currently playing sound
Public Const SND_FILENAME = &H20000     '  name is a file name
Public Const SND_ASYNC = &H1         '  play asynchronously

Public Const HWND_TOPMOST = -1

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

Public Enum EFormStatus
    AddStatus = 0
    ModifyStatus = 1
    ShowStatus = 2
    NotStatus = 3
End Enum


Public Type TEventSound                        '事件音效文件的路径信息结构
    CheckSheetNotExist As String           '路单不存在
    CheckSheetCanceled As String           '已作废路单
    CheckSheetSettled As String            '路单已结算
    CheckSheetSelected As String           '路单已选择
    CheckSheetValid As String              '路单有效
    ObjectNotSame As String                '路单有效,但不在所要拆算时间或对象范围之内
End Type

Public g_tEventSoundPath As TEventSound


Public Const m_cRegSoundKey = "Settle\EventSound"      '音效主键字符串
Public g_atAllSellStation() As TDepartmentInfo  '所有的售票站点
Public g_szStationID As String '系统参数中的本站代码


Public Const cnPriceItemNum = 15


Public Sub Main()
    Dim oShell As New CommShell
    
    On Error GoTo ErrorHandle
    
     
    
    Set g_oActiveUser = oShell.ShowLogin()
    If g_oActiveUser Is Nothing Then Exit Sub
    
    InitSound
    
'    oShell.ShowSplash "财务结算管理", "Station Settle Management", LoadResPicture(101, 0), App.Major, App.Minor, App.Revision
    '稍作延时
    DoEvents
    mdiMain.Show
    
    
    g_oParam.Init g_oActiveUser
    g_szUnitID = g_oParam.UnitID
    Date = g_oParam.NowDate
    Time = g_oParam.NowTime

    g_szFixFeeItem = g_oParam.FixFeeItem '结算项中的固定费用项.
    g_bAllowSellteTotalNegative = g_oParam.AllowSettleTotalNegative '是否将上月结算的负数算到下月
    g_szSettleNegativeSplitItem = g_oParam.SettleNegativeSplitItem '将上月结算的负数项放在结算项的位置
    g_szIsFixFeeUpdateEachMonth = g_oParam.IsFixFeeUpdateEachMonth  '是否固定费用是每个月更新的
    
    
'    oShell.CloseSplash
    DoEvents
    Exit Sub
ErrorHandle:
    ShowErrorMsg
'    Resume GoOn
End Sub

Private Sub InitSound()

    '以下初始注册表设置
    Dim oReg As CFreeReg
    Set oReg = New CFreeReg
    
    
    oReg.Init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany
    
    '读取各种音效文件路径
    g_tEventSoundPath.CheckSheetCanceled = oReg.GetSetting(m_cRegSoundKey, "CheckSheetCanceled")
    g_tEventSoundPath.CheckSheetNotExist = oReg.GetSetting(m_cRegSoundKey, "CheckSheetNotExist")
    g_tEventSoundPath.CheckSheetSelected = oReg.GetSetting(m_cRegSoundKey, "CheckSheetSelected")
    g_tEventSoundPath.CheckSheetSettled = oReg.GetSetting(m_cRegSoundKey, "CheckSheetSettled")
    g_tEventSoundPath.CheckSheetValid = oReg.GetSetting(m_cRegSoundKey, "CheckSheetValid")
    g_tEventSoundPath.ObjectNotSame = oReg.GetSetting(m_cRegSoundKey, "ObjectNotSame")
    
    
    Set oReg = Nothing
End Sub







Public Sub PlayEventSound(szFileName As String)
    '播放音效
    PlaySound szFileName, 0, SND_FILENAME + SND_ASYNC
    
End Sub

Public Function GetDocumentDir() As String
    Dim oReg As New CFreeReg
    Dim szFileDir As String
    On Error GoTo Error_Handle
    oReg.Init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany
    szFileDir = oReg.GetSetting(cszLuggageAccount, cszDocumentDir, cszDefDocumentDir)
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
        oReg.SaveSetting cszLuggageAccount, cszDocumentDir, szPath
    Else
        oReg.SaveSetting cszLuggageAccount, cszDocumentDir, cszDocumentDir
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
        oReg.SaveSetting cszLuggageAccount, cszRecentSeller, szRecentSeller
    End If
End Sub

Public Function GetRecentSeller() As String
    Dim oReg As New CFreeReg
    oReg.Init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany
    GetRecentSeller = oReg.GetSetting(cszLuggageAccount, cszRecentSeller)
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
        .abMenu.Refresh
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


Public Sub FillSellStation(cboSellStation As ComboBox)
    Dim oSystemMan As New SystemMan
    Dim atTemp() As TDepartmentInfo
    Dim i As Integer
    On Error GoTo here
    '判断用户属于哪个上车站,如果为空则填充一个空行,再填充所有的上车站
    oSystemMan.Init g_oActiveUser
    atTemp = oSystemMan.GetAllSellStation(g_szUnitID)
    If g_oActiveUser.SellStationID = "" Then
        cboSellStation.AddItem ""
        For i = 1 To ArrayLength(atTemp)
            cboSellStation.AddItem MakeDisplayString(atTemp(i).szSellStationID, atTemp(i).szSellStationName)
        Next i
    '否则只填充用户所属的上车站
    Else
        For i = 1 To ArrayLength(atTemp)
            If g_oActiveUser.SellStationID = atTemp(i).szSellStationID Then
               cboSellStation.AddItem MakeDisplayString(atTemp(i).szSellStationID, atTemp(i).szSellStationName)
               Exit For
            End If
        Next i
        cboSellStation.ListIndex = 0
    End If
    Exit Sub
here:
    ShowErrorMsg
End Sub


'写主界面的标题栏
Public Sub WriteTitleBar(Optional pszFormName As String = "", Optional poIcon As StdPicture)
'    'pszFormName空时则清空
'    With mdiMain
'    If pszFormName = "" Then
'        .lblInfoBar = ""
'        Set .imgInfoBar.Picture = Nothing
'    Else
'        .lblInfoBar = pszFormName
'        Set .imgInfoBar.Picture = poIcon
'    End If
'    End With
End Sub






Public Function GetDefaultMark(ByVal nValue As Integer) As String
    Select Case nValue
        Case 0
            GetDefaultMark = "默认"
        Case 1
            GetDefaultMark = "非默认"
    End Select
End Function


Public Function GetSplitStatus(ByVal Value As Integer) As String '1-使用,0-未用
    Select Case Value
        Case 0
            GetSplitStatus = "未用"
        Case 1
            GetSplitStatus = "使用"
    End Select
End Function

Public Function GetSplitType(ByVal Value As Integer) As String '0-拆给对方公司,1-拆给站方,2-留给本公司
    Select Case Value
        Case 0
            GetSplitType = "拆给对方公司"
        Case 1
            GetSplitType = "拆给站方"
        Case 2
            GetSplitType = "留给本公司"
    End Select
End Function
Public Function GetAllowModify(ByVal Value As Integer) As String '0-不允许修改,1-允许修改
    Select Case Value
        Case 0
            GetAllowModify = "不允许修改"
        Case 1
            GetAllowModify = "允许修改"
        
    End Select
End Function

  '0-拆帐公司 1-车辆 2-参运公司 3-车主 4-车次
Public Function GetObjectTypeInt(szObject As String) As Integer
    Select Case szObject
           Case "车次"
            GetObjectTypeInt = CS_SettleByBus
           Case "车辆"
            GetObjectTypeInt = CS_SettleByVehicle
           Case "参运公司"
            GetObjectTypeInt = CS_SettleByTransportCompany
           Case "车主"
            GetObjectTypeInt = CS_SettleByOwner
           Case "拆账公司"
            GetObjectTypeInt = CS_SettleBySplitCompany
           Case "全部"
            GetObjectTypeInt = -1
    End Select
End Function

Public Function GetObjectTypeString(nObject As Integer) As String
    Select Case nObject
           Case CS_SettleByBus
            GetObjectTypeString = "车次"
           Case CS_SettleByVehicle
            GetObjectTypeString = "车辆"
           Case CS_SettleByTransportCompany
            GetObjectTypeString = "参运公司"
           Case CS_SettleByOwner
            GetObjectTypeString = "车主"
           Case CS_SettleBySplitCompany
            GetObjectTypeString = "拆账公司"
    End Select
End Function

''得到结算单状态代码
'Public Function GetSettleSheetStatusInt(szStatus As String) As Integer
'    Select Case szStatus
'        Case "未结"
'            GetSettleSheetStatusInt = CS_SettleSheetValid
'        Case "作废"
'            GetSettleSheetStatusInt = CS_SettleSheetInvalid
'        Case "已结"
'            GetSettleSheetStatusInt = CS_SettleSheetSettled
'    End Select
'End Function
'
''得到结算单状态
'Public Function GetSettleSheetStatusString(szStatus As Integer) As String
'    Select Case szStatus
'        Case CS_SettleSheetValid
'            GetSettleSheetStatusString = "未结"
'        Case CS_SettleSheetInvalid
'            GetSettleSheetStatusString = "作废"
'        Case CS_SettleSheetSettled
'            GetSettleSheetStatusString = "已结"
'    End Select
'End Function





Public Function GetSettleSheetStatusInt(szStatus As String) As Integer
    Select Case szStatus
        Case "未结"
            GetSettleSheetStatusInt = CS_SettleSheetValid
        Case "作废"
            GetSettleSheetStatusInt = CS_SettleSheetInvalid
        Case "已汇"
            GetSettleSheetStatusInt = CS_SettleSheetSettled
'        Case "应扣款未结清"
'            GetSettleSheetStatusInt = CS_SettleSheetNegativeNotPay
'        Case "应扣款已结清"
'            GetSettleSheetStatusInt = CS_SettleSheetNegativeHasPayed
        Case "全部"
            GetSettleSheetStatusInt = -1
    End Select
End Function

Public Function GetSettleSheetStatusString(szStatus As Integer) As String
    Select Case szStatus
        Case CS_SettleSheetValid
            GetSettleSheetStatusString = "未结"
        Case CS_SettleSheetInvalid
            GetSettleSheetStatusString = "作废"
        Case CS_SettleSheetSettled
            GetSettleSheetStatusString = "已汇"
        Case CS_SettleSheetNotInvalid
            GetSettleSheetStatusString = "非作废"
'        Case CS_SettleSheetNegativeHasPayed
'            GetSettleSheetStatusString = "应扣款已结清"
        Case -1
            GetSettleSheetStatusString = "全部"
    End Select
End Function




'得到路单站点的改并状态名称
Public Function GetSheetStationStatusName(pszStatus As Integer)
    Select Case pszStatus
    '0-正常，1-改签，2-并班
    Case 0
        GetSheetStationStatusName = "正常"
    Case 1
        GetSheetStationStatusName = "改签"
    Case 2
        GetSheetStationStatusName = "并班"
    End Select
    
    
End Function


Public Function GetQueryNegativeStatusString(nStatus As EQueryNegativeType) As String
    Select Case nStatus
        Case CS_QueryAll
            GetQueryNegativeStatusString = "全部"
        Case CS_QueryNegative
            GetQueryNegativeStatusString = "应结款为负"
        Case CS_QueryNotNegative
            GetQueryNegativeStatusString = "应结款为正"
    End Select
End Function


Public Function GetQueryNegativeStatusInt(szStatus As String) As Integer
    Select Case szStatus
        Case "全部"
            GetQueryNegativeStatusInt = CS_QueryAll
        Case "应结款为负"
            GetQueryNegativeStatusInt = CS_QueryNegative
        Case "应结款为正"
            GetQueryNegativeStatusInt = CS_QueryNotNegative
            
    End Select
End Function




'得到路单站点的改并状态名称
Public Function GetFixFeeStatusName(pnStatus As Integer)
    Select Case pnStatus
    Case -1
        GetFixFeeStatusName = "全部"
    Case 0
        GetFixFeeStatusName = "未扣"
    Case Else
        GetFixFeeStatusName = "已扣"
    End Select
    
    
End Function
