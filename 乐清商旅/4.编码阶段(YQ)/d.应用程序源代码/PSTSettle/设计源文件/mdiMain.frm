VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "actbar2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm mdiMain 
   BackColor       =   &H8000000C&
   Caption         =   "路单结算"
   ClientHeight    =   6015
   ClientLeft      =   3450
   ClientTop       =   2850
   ClientWidth     =   7395
   Icon            =   "mdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   WindowState     =   2  'Maximized
   Begin ActiveBar2LibraryCtl.ActiveBar2 abMenu 
      Align           =   1  'Align Top
      Height          =   6015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7395
      _LayoutVersion  =   1
      _ExtentX        =   13044
      _ExtentY        =   10610
      _DataPath       =   ""
      Bands           =   "mdiMain.frx":16AC2
      Begin MSComctlLib.ProgressBar pbLoad 
         Height          =   225
         Left            =   2910
         TabIndex        =   1
         Top             =   7170
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const cszCompanySettleDetail = "公司结算明细表.xls"
Const cszVehicleSettleDetail = "车辆结算明细表.xls"
Const cszBusSettleDetail = "车次结算明细表.xls"
Const cszSettleSheetStat = "路单结算明细表.xls"
Const cszCompanySettleStat = "公司结算汇总表.xls"
Const cszVehicleSettleStat = "车辆结算汇总表.xls"
Const cszBusSettleStat = "车次结算汇总表.xls"
Const cszVehicleSettleStatByMonth = "车辆结算汇总月报.xls"
Const cszVehicleFixFee = "车辆固定费用表.xls"
Const cszBusFixFee = "车次固定费用表.xls"


Const cszProtocol = "协议明细.xls"
Const cszVehicleProtocol = "车辆协议明细.xls"
Const cszCompanyProtocol = "公司协议明细.xls"
Const cszVehicleSettlePrice = "车辆结算价明细.xls"
Const cszCompanySettlePrice = "公司结算价明细.xls"


Private Sub MDIForm_Load()
    AddControlsToActBar
    '状态条
    ShowSBInfo "", ESB_WorkingInfo
    ShowSBInfo "", ESB_ResultCountInfo
    ShowSBInfo EncodeString(g_oActiveUser.UserID) & g_oActiveUser.UserName, ESB_UserInfo
    ShowSBInfo Format(g_oActiveUser.LoginTime, "HH:mm"), ESB_LoginTime
    SetPrintEnabled False
End Sub

'关联ActiveBar的控件
Private Sub AddControlsToActBar()
'    abMenu.Bands("bndTitleTop").Tools("tblTitleTop").Custom = ptTitleTop
    abMenu.Bands("statusBar").Tools("progressBar").Custom = pbLoad
End Sub
                     
'设置菜单的可用项
Public Sub SetPrintEnabled(bEnabled As Boolean)
    abMenu.Bands("mnu_System").Tools("mnu_ExportFile").Enabled = bEnabled
    abMenu.Bands("mnu_System").Tools("mnu_ExportFileOpen").Enabled = bEnabled
    abMenu.Bands("mnu_System").Tools("mnu_system_print").Enabled = bEnabled
    abMenu.Bands("mnu_System").Tools("mnu_system_printview").Enabled = bEnabled
    abMenu.Bands("mnu_System").Tools("mnu_PageOption").Enabled = bEnabled
    abMenu.Bands("mnu_System").Tools("mnu_PrintOption").Enabled = bEnabled
    
    abMenu.Bands("tbn_system").Tools("tbn_system_export").Enabled = bEnabled
    abMenu.Bands("tbn_system").Tools("tbn_system_exportopen").Enabled = bEnabled
    abMenu.Bands("tbn_system").Tools("tbn_system_print").Enabled = bEnabled
    abMenu.Bands("tbn_system").Tools("tbn_system_printview").Enabled = bEnabled
'    abMenu.Bands("tbn_system").Tools("tbn_system").Enabled = bEnabled
'    abMenu.Bands("tbn_system").Tools("tbn_system").Enabled = bEnabled
        
End Sub

Private Sub abMenu_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
On Error GoTo ErrHandle
Dim frmTemp As Form

Select Case Tool.name
    '系统
    Case "mnu_SetOption"
        frmSetOption.Show vbModal
    
    Case "mnu_BaseInfo"      ' 协议管理
        frmBaseInfo.ZOrder 0
        frmBaseInfo.Show

        
    Case "mnu_CompanyProtocol" '公司协议设置
        frmCompanyProtocol.ZOrder 0
        frmCompanyProtocol.Show

        
    Case "mnu_VehicleProtocol"  '车辆协议设置
        frmVehicleProtocol.ZOrder 0
        frmVehicleProtocol.Show
        
    Case "mnu_BusProtocol"  '车次协议设置
        frmBusProtocol.ZOrder 0
        frmBusProtocol.Show

        
    Case "mnu_CompanySettlePrice" '公司结算价设置
        frmCompanySettlePrice.ZOrder 0
        frmCompanySettlePrice.Show

        
    Case "mnu_VehicleSettlePrice" '车辆结算价设置
        frmVehicleSettlePrice.ZOrder 0
        frmVehicleSettlePrice.Show
        
    Case "mnu_BusSettlePrice" '车次结算价设置
        frmBusSettlePrice.ZOrder 0
        frmBusSettlePrice.Show

    Case "mnu_AllFixFee" '车辆固定费用
        
        Set frmTemp = New frmAllVehicleFixFee
        frmTemp.m_szFormStatus = EFS_Vehicle
        frmTemp.ZOrder 0
        frmTemp.Show
        
    Case "mnu_BusFixFee" '车次固定费用
        Set frmTemp = New frmAllVehicleFixFee
        frmTemp.m_szFormStatus = EFS_Bus
        frmTemp.ZOrder 0
        frmTemp.Show
'        frmAllVehicleFixFee.m_szFormStatus = EFS_Bus
'        frmAllVehicleFixFee.ZOrder 0
'        frmAllVehicleFixFee.Show
        
    Case "mnu_HalveCompany"  '加总平分公司
        frmHalve.ZOrder 0
        frmHalve.Show

    
    Case "mnu_AllStation"
        frmAllStation.ZOrder 0
        frmAllStation.Show
        
    Case "mnu_AllSection"
        frmAllSection.ZOrder 0
        frmAllSection.Show
        
    Case "mnu_AllRoute"
        frmAllRoute.ZOrder 0
        frmAllRoute.Show
    
    '报表统计
    Case "mnu_SettleSheetStat"  '路单结算明细表
        mnu_SettleSheetStat_Click
    Case "mnu_CompanyDetail" '公司结算明细表
        mnu_CompanyDetail_Click
    Case "mnu_VehicleDetail" '车辆结算明细表
        mnu_VehicleDetail_Click
    Case "mnu_BusDetail" '车次结算明细表
        mnu_BusDetail_Click
    Case "mnu_CompanySettleStat" '公司结算汇总表
        mnu_CompanySettleStat_Click
    Case "mnu_VehicleSettleStat" '车辆结算汇总表
        mnu_VehicleSettleStat_Click
    Case "mnu_BusSettleStat" '车次结算汇总表
        mnu_BusSettleStat_Click
    Case "mnu_VehicleSettleStatByMonth" '车辆结算汇总月报表
        mnu_VehicleSettleStatByMonth_Click
    
    Case "mnu_vehiclebalancestat" '车辆结算平衡表
        mnu_vehiclebalancestat_Click
    Case "mnu_busbalancestat" '车次结算平衡表
        mnu_busbalancestat_Click
    
    
    
    Case "mnu_protocol"
        '查询协议
        mnu_Protocol
    Case "mnu_vehicleprotocol"
        '车辆协议
        mnu_VehicleProtocol
    Case "mnu_companyprotocol"
        '公司协议
        mnu_CompanyProtocol
    Case "mnu_vehiclesettleprice"
        '车辆结算价
        mnu_VehicleSettlePrice
    Case "mnu_companysettleprice"
        '公司结算价
        mnu_CompanySettlePrice
    Case "mnu_VehicleFixFeeStat"
        '车辆固定费用报表
        mnu_VehicleFixFeeStat
    Case "mnu_BusFixFeeStat"
        '车次固定费用报表
        mnu_BusFixFeeStat
        
        
    '路单拆算
    Case "mi_SettleSheet"    '路单结算管理
        frmAllSettleSheets.ZOrder 0
        frmAllSettleSheets.Show

         
    Case "mi_NewSettleSheet"    '路单结算向导
'        frmWizSplitSettle.ZOrder 0
        frmWizSplitSettle.Show vbModal
    Case "mi_WizSplitSettleSheetManual"
'        frmWizSplitSettleBack.ZOrder 0
        frmWizSplitSettleBack.Show vbModal
         
    Case "mi_RePrintSettleSheet" '重打结算单
        frmRePrintSettleSheet.ZOrder 0
        frmRePrintSettleSheet.Show vbModal
    
    
    Case "mi_ViewCheckSheet"
        '察看路单

        frmAllSheet.m_dtStartDate = GetFirstMonthDay(Date)
        frmAllSheet.m_dtEndDate = GetLastMonthDay(Date)
        
        frmAllSheet.ZOrder 0
        frmAllSheet.Show
        
        
    Case "mi_ModifyCheckSheet"
        '修改路单
        frmModifySheet.Show vbModal
        
    Case "mi_MakeSheetStatTemp"
        MakeSheetStatTemp
        
    '窗口
    Case "mnu_TitleH"
'        mnu_TitleH_Click
    Case "mnu_TitleV"
'        mnu_TitleV_Click
    Case "mnu_Cascade"
'        mnu_Cascade_Click
    Case "mnu_ArrangeIcon"
'        mnu_ArrangeIcon_Click
    '帮助
    Case "mnu_HelpIndex"
'        frmWizSplitSettle.Show
    Case "mnu_HelpContent"
'        mnu_HelpContent_Click
    Case "mnu_About"
        mnu_About_Click
    
        '以下是系统部分
        Case "tbn_system_print"
            ActiveForm.PrintReport False
        Case "mnu_system_print"
            ActiveForm.PrintReport True
        Case "tbn_system_printview", "mnu_system_printview"
            ActiveForm.PreView
        Case "mnu_PageOption"
            '页面设置
            ActiveForm.PageSet
        Case "mnu_PrintOption"
            '打印设置
            ActiveForm.PrintSet
        Case "tbn_system_export", "mnu_ExportFile"
            ActiveForm.ExportFile
        Case "tbn_system_exportopen", "mnu_ExportFileOpen"
            ActiveForm.ExportFileOpen
        Case "mnu_ChgPassword"
            '修改口令
            ChangePassword
        Case "mnu_SysExit", "tbn_system_exit"
            ExitSystem
    End Select
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

'路单结算明细表
Private Sub mnu_SettleSheetStat_Click()
    On Error GoTo ErrHandle
    Dim lHelpContextID As Long
    Dim rsSettleDetail As Recordset
    Dim i As Integer
    Dim j As Integer
    Dim m_oReport As New Report
    Dim frmNewReport As New frmReport
    Dim vCustomData As Variant
    
    lHelpContextID = frmSettleSheetStat.HelpContextID
    
    frmSettleSheetStat.Show vbModal
    If frmSettleSheetStat.m_bOk Then
    
        m_oReport.Init g_oActiveUser
        
        '取得记录集
        Set rsSettleDetail = m_oReport.SettleSheetStat(frmSettleSheetStat.m_dtStartDate, DateAdd("d", 1, frmSettleSheetStat.m_dtEndDate))
        WriteProcessBar False, , , ""
        ReDim vCustomData(1 To 3, 1 To 2)
        vCustomData(1, 1) = "开始日期"
        vCustomData(1, 2) = ToDBDate(frmSettleSheetStat.m_dtStartDate)
        vCustomData(2, 1) = "结束日期"
        vCustomData(2, 2) = ToDBDate(frmSettleSheetStat.m_dtEndDate)
        vCustomData(3, 1) = "打印"
        vCustomData(3, 2) = g_oActiveUser.UserName
        
        
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport rsSettleDetail, cszSettleSheetStat, frmSettleSheetStat.Caption, vCustomData, 10
        
        
        WriteProcessBar False, , , ""
    End If
    
    
    
    Exit Sub
ErrHandle:
    WriteProcessBar False, , , ""
    ShowErrorMsg
End Sub

'公司结算明细表
Private Sub mnu_CompanyDetail_Click()
    On Error GoTo ErrHandle
    Dim lHelpContextID As Long
    Dim vCustomData As Variant
    Dim rsSettleDetail As Recordset
    Dim m_oReport As New Report
    Dim frmNewReport As New frmReport
    
    lHelpContextID = frmCompanySettleStat.HelpContextID
    frmCompanySettleStat.Show vbModal
    If frmCompanySettleStat.m_bOk Then
        m_oReport.Init g_oActiveUser
        
        
        '取得记录集
        Set rsSettleDetail = m_oReport.CompanySettleDetail(frmCompanySettleStat.m_dtStartDate, DateAdd("d", 1, frmCompanySettleStat.m_dtEndDate), ResolveDisplay(frmCompanySettleStat.m_szCompanyID))
        WriteProcessBar True, , , ""
        ReDim vCustomData(1 To 4, 1 To 2)
        vCustomData(1, 1) = "开始日期"
        vCustomData(1, 2) = ToDBDate(frmCompanySettleStat.m_dtStartDate)
        vCustomData(2, 1) = "结束日期"
        vCustomData(2, 2) = ToDBDate(frmCompanySettleStat.m_dtEndDate)
        vCustomData(3, 1) = "打印"
        vCustomData(3, 2) = g_oActiveUser.UserName
        vCustomData(4, 1) = "公司"
        vCustomData(4, 2) = IIf(frmCompanySettleStat.m_szCompanyID = "", "所有公司", frmCompanySettleStat.m_szCompanyID)
        
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport rsSettleDetail, cszCompanySettleDetail, frmCompanySettleStat.Caption, vCustomData, 10
        
        WriteProcessBar False, , , ""
    End If
    
    
    
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

'车辆结算明细表
Private Sub mnu_VehicleDetail_Click()
    On Error GoTo ErrHandle
    Dim lHelpContextID As Long
    Dim vCustomData As Variant
    Dim rsSettleDetail As Recordset
    Dim m_oReport As New Report
    Dim frmNewReport As New frmReport
    
    lHelpContextID = frmVehicleSettleStat.HelpContextID
    
    
    frmVehicleSettleStat.Show vbModal
    If frmVehicleSettleStat.m_bOk Then
        
        m_oReport.Init g_oActiveUser
        
        '取得记录集
        Set rsSettleDetail = m_oReport.VehicleSettleDetail(frmVehicleSettleStat.m_dtStartDate, DateAdd("d", 1, frmVehicleSettleStat.m_dtEndDate), ResolveDisplay(frmVehicleSettleStat.m_szVehicleID), ResolveDisplay(frmVehicleSettleStat.m_szCompanyID), frmVehicleSettleStat.m_nStatus, frmVehicleSettleStat.m_nQueryNegativeType, frmVehicleSettleStat.m_bStatBySettleDate)
        WriteProcessBar True, , , ""
        ReDim vCustomData(1 To 8, 1 To 2)
        vCustomData(1, 1) = "开始日期"
        vCustomData(1, 2) = ToDBDate(frmVehicleSettleStat.m_dtStartDate)
        vCustomData(2, 1) = "结束日期"
        vCustomData(2, 2) = ToDBDate(frmVehicleSettleStat.m_dtEndDate)
        vCustomData(3, 1) = "打印"
        vCustomData(3, 2) = g_oActiveUser.UserName
        vCustomData(4, 1) = "车辆"
        vCustomData(4, 2) = IIf(frmVehicleSettleStat.m_szVehicleID = "", "所有车辆", frmVehicleSettleStat.m_szVehicleTagNo)
        vCustomData(5, 1) = "公司"
        vCustomData(5, 2) = IIf(frmVehicleSettleStat.m_szCompanyID = "", "所有公司", frmVehicleSettleStat.m_szCompanyName)
        
        vCustomData(6, 1) = "类别"
        vCustomData(6, 2) = GetSettleSheetStatusString(frmVehicleSettleStat.m_nStatus)
        
        
        vCustomData(7, 1) = "负数类别"
        vCustomData(7, 2) = GetQueryNegativeStatusString(frmVehicleSettleStat.m_nQueryNegativeType)
        
        
        vCustomData(8, 1) = "统计方式"
        vCustomData(8, 2) = IIf(frmVehicleSettleStat.m_bStatBySettleDate, "按结算日期统计", "按车次日期统计")
        
        
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport rsSettleDetail, cszVehicleSettleDetail, frmVehicleSettleStat.Caption, vCustomData, 10
        
        WriteProcessBar False, , , ""
    End If
    
    
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

'车辆固定费用报表
Private Sub mnu_VehicleFixFeeStat()
    On Error GoTo ErrHandle
    Dim lHelpContextID As Long
    Dim vCustomData As Variant
    Dim rsVehicleFixFee As Recordset
    Dim m_oReport As New Report
    Dim frmNewReport As New frmReport
    
    lHelpContextID = frmVehicleFixFeeReport.HelpContextID
    
    
    frmVehicleFixFeeReport.Show vbModal
    If frmVehicleFixFeeReport.m_bOk Then
        
        m_oReport.Init g_oActiveUser
        
        '取得记录集
        Set rsVehicleFixFee = m_oReport.GetAllVehicleFixFee(ResolveDisplay(frmVehicleFixFeeReport.m_szVehicleID), ResolveDisplay(frmVehicleFixFeeReport.m_szCompanyID), frmVehicleFixFeeReport.m_dtStartDate, DateAdd("d", 1, frmVehicleFixFeeReport.m_dtEndDate), , frmVehicleFixFeeReport.m_nStatus)
        WriteProcessBar True, , , ""
        ReDim vCustomData(1 To 8, 1 To 2)
        vCustomData(1, 1) = "开始日期"
        vCustomData(1, 2) = ToDBDate(frmVehicleFixFeeReport.m_dtStartDate)
        vCustomData(2, 1) = "结束日期"
        vCustomData(2, 2) = ToDBDate(frmVehicleFixFeeReport.m_dtEndDate)
        vCustomData(3, 1) = "打印"
        vCustomData(3, 2) = g_oActiveUser.UserName
        vCustomData(4, 1) = "车辆"
        vCustomData(4, 2) = IIf(frmVehicleFixFeeReport.m_szVehicleID = "", "所有车辆", frmVehicleFixFeeReport.m_szVehicleTagNo)
        vCustomData(5, 1) = "公司"
        vCustomData(5, 2) = IIf(frmVehicleFixFeeReport.m_szCompanyID = "", "所有公司", frmVehicleFixFeeReport.m_szCompanyName)
        
        vCustomData(6, 1) = "类别"
        vCustomData(6, 2) = GetFixFeeStatusName(frmVehicleFixFeeReport.m_nStatus)
        
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport rsVehicleFixFee, cszVehicleFixFee, frmVehicleFixFeeReport.Caption, vCustomData, 10
        
        WriteProcessBar False, , , ""
    End If
    
    
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

'车次固定费用报表
Private Sub mnu_BusFixFeeStat()
    On Error GoTo ErrHandle
    Dim lHelpContextID As Long
    Dim vCustomData As Variant
    Dim rsBusFixFee As Recordset
    Dim m_oReport As New Report
    Dim frmNewReport As New frmReport
    
    lHelpContextID = frmBusFixFeeReport.HelpContextID
    
    
    frmBusFixFeeReport.Show vbModal
    If frmBusFixFeeReport.m_bOk Then
        
        m_oReport.Init g_oActiveUser
        
        '取得记录集
        Set rsBusFixFee = m_oReport.GetAllBusFixFee(ResolveDisplay(frmBusFixFeeReport.m_szBusID), ResolveDisplay(frmBusFixFeeReport.m_szCompanyID), frmBusFixFeeReport.m_dtStartDate, DateAdd("d", 1, frmBusFixFeeReport.m_dtEndDate), , frmBusFixFeeReport.m_nStatus)
        WriteProcessBar True, , , ""
        ReDim vCustomData(1 To 8, 1 To 2)
        vCustomData(1, 1) = "开始日期"
        vCustomData(1, 2) = ToDBDate(frmBusFixFeeReport.m_dtStartDate)
        vCustomData(2, 1) = "结束日期"
        vCustomData(2, 2) = ToDBDate(frmBusFixFeeReport.m_dtEndDate)
        vCustomData(3, 1) = "打印"
        vCustomData(3, 2) = g_oActiveUser.UserName
        vCustomData(4, 1) = "车次"
        vCustomData(4, 2) = IIf(frmBusFixFeeReport.m_szBusID = "", "所有车次", frmBusFixFeeReport.m_szBusTagNo)
        vCustomData(5, 1) = "公司"
        vCustomData(5, 2) = IIf(frmBusFixFeeReport.m_szCompanyID = "", "所有公司", frmBusFixFeeReport.m_szCompanyName)
        
        vCustomData(6, 1) = "类别"
        vCustomData(6, 2) = GetFixFeeStatusName(frmBusFixFeeReport.m_nStatus)
        
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport rsBusFixFee, cszBusFixFee, frmBusFixFeeReport.Caption, vCustomData, 10
        
        WriteProcessBar False, , , ""
    End If
    
    
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

'车次结算明细表
Private Sub mnu_BusDetail_Click()
    On Error GoTo ErrHandle
    Dim lHelpContextID As Long
    Dim vCustomData As Variant
    Dim rsSettleDetail As Recordset
    Dim m_oReport As New Report
    Dim frmNewReport As New frmReport
    
    lHelpContextID = frmBusSettleStat.HelpContextID
    
    
    frmBusSettleStat.Show vbModal
    If frmBusSettleStat.m_bOk Then
        
        m_oReport.Init g_oActiveUser
        
        '取得记录集
        Set rsSettleDetail = m_oReport.BusSettleDetail(frmBusSettleStat.m_dtStartDate, DateAdd("d", 1, frmBusSettleStat.m_dtEndDate), ResolveDisplay(frmBusSettleStat.m_szBusID), ResolveDisplay(frmBusSettleStat.m_szCompanyID), frmBusSettleStat.m_nStatus, frmBusSettleStat.m_nQueryNegativeType, frmBusSettleStat.m_bStatBySettleDate)
        WriteProcessBar True, , , ""
        ReDim vCustomData(1 To 8, 1 To 2)
        vCustomData(1, 1) = "开始日期"
        vCustomData(1, 2) = ToDBDate(frmBusSettleStat.m_dtStartDate)
        vCustomData(2, 1) = "结束日期"
        vCustomData(2, 2) = ToDBDate(frmBusSettleStat.m_dtEndDate)
        vCustomData(3, 1) = "打印"
        vCustomData(3, 2) = g_oActiveUser.UserName
        vCustomData(4, 1) = "车次"
        vCustomData(4, 2) = IIf(frmBusSettleStat.m_szBusID = "", "所有车次", frmBusSettleStat.m_szBusTagNo)
        vCustomData(5, 1) = "公司"
        vCustomData(5, 2) = IIf(frmBusSettleStat.m_szCompanyID = "", "所有公司", frmBusSettleStat.m_szCompanyName)
        
        vCustomData(6, 1) = "类别"
        vCustomData(6, 2) = GetSettleSheetStatusString(frmBusSettleStat.m_nStatus)
        
        
        vCustomData(7, 1) = "负数类别"
        vCustomData(7, 2) = GetQueryNegativeStatusString(frmBusSettleStat.m_nQueryNegativeType)
        
        
        vCustomData(8, 1) = "统计方式"
        vCustomData(8, 2) = IIf(frmBusSettleStat.m_bStatBySettleDate, "按结算日期统计", "按车次日期统计")
        
        
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport rsSettleDetail, cszBusSettleDetail, frmBusSettleStat.Caption, vCustomData, 10
        
        WriteProcessBar False, , , ""
    End If
    
    
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

'公司结算汇总表
Private Sub mnu_CompanySettleStat_Click()
    On Error GoTo ErrHandle
    
    Dim lHelpContextID As Long
    Dim frmNewReport As New frmReport
    Dim vCustomData As Variant
    Dim rsSettleDetail As Recordset

    Dim m_oReport As New Report
    
    lHelpContextID = frmCompanySettleStat.HelpContextID
    frmCompanySettleStat.Show vbModal
    If frmCompanySettleStat.m_bOk Then
        m_oReport.Init g_oActiveUser
        '取得记录集
        Set rsSettleDetail = m_oReport.CompanySettleStat(frmCompanySettleStat.m_dtStartDate, DateAdd("d", 1, frmCompanySettleStat.m_dtEndDate), ResolveDisplay(frmCompanySettleStat.m_szCompanyID))
        WriteProcessBar True, , , ""
        ReDim vCustomData(1 To 4, 1 To 2)
        vCustomData(1, 1) = "开始日期"
        vCustomData(1, 2) = ToDBDate(frmCompanySettleStat.m_dtStartDate)
        vCustomData(2, 1) = "结束日期"
        vCustomData(2, 2) = ToDBDate(frmCompanySettleStat.m_dtEndDate)
        vCustomData(3, 1) = "打印"
        vCustomData(3, 2) = g_oActiveUser.UserName
        vCustomData(4, 1) = "公司"
        vCustomData(4, 2) = IIf(frmCompanySettleStat.m_szCompanyID = "", "所有公司", frmCompanySettleStat.m_szCompanyID)
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport rsSettleDetail, cszCompanySettleStat, frmCompanySettleStat.Caption, vCustomData, 10
        WriteProcessBar False, , , ""
    End If
    
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

'车辆结算汇总表
Private Sub mnu_VehicleSettleStat_Click()
    On Error GoTo ErrHandle
    
    Dim lHelpContextID As Long
    Dim rsSettleDetail As Recordset

    Dim m_oReport As New Report
    Dim vCustomData As Variant
    Dim frmNewReport As New frmReport
    
    lHelpContextID = frmVehicleSettleStat.HelpContextID
    frmVehicleSettleStat.Show vbModal
    If frmVehicleSettleStat.m_bOk Then
        m_oReport.Init g_oActiveUser
        '取得记录集
        Set rsSettleDetail = m_oReport.VehicleSettleStat(frmVehicleSettleStat.m_dtStartDate, DateAdd("d", 1, frmVehicleSettleStat.m_dtEndDate), ResolveDisplay(frmVehicleSettleStat.m_szVehicleID), ResolveDisplay(frmVehicleSettleStat.m_szCompanyID), frmVehicleSettleStat.m_nQueryNegativeType, frmVehicleSettleStat.m_nStatus, frmVehicleSettleStat.m_bStatBySettleDate)
        WriteProcessBar True, , , ""
        ReDim vCustomData(1 To 7, 1 To 2)
        vCustomData(1, 1) = "开始日期"
        vCustomData(1, 2) = ToDBDate(frmVehicleSettleStat.m_dtStartDate)
        vCustomData(2, 1) = "结束日期"
        vCustomData(2, 2) = ToDBDate(frmVehicleSettleStat.m_dtEndDate)
        vCustomData(3, 1) = "打印"
        vCustomData(3, 2) = g_oActiveUser.UserName
        vCustomData(4, 1) = "车辆"
        vCustomData(4, 2) = IIf(frmVehicleSettleStat.m_szVehicleID = "", "所有车辆", frmVehicleSettleStat.m_szVehicleTagNo)
        
        vCustomData(5, 1) = "公司"
        vCustomData(5, 2) = IIf(frmVehicleSettleStat.m_szCompanyID = "", "所有公司", frmVehicleSettleStat.m_szCompanyName)
        
        vCustomData(6, 1) = "上车站"
        vCustomData(6, 2) = g_oActiveUser.UserUnitName
        
        vCustomData(7, 1) = "统计方式"
        vCustomData(7, 2) = IIf(frmVehicleSettleStat.m_bStatBySettleDate, "按结算日期统计", "按车次日期统计")
        
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport rsSettleDetail, cszVehicleSettleStat, frmVehicleSettleStat.Caption, vCustomData, 10
        
        WriteProcessBar False, , , ""
    End If
    
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

'车次结算汇总表
Private Sub mnu_BusSettleStat_Click()
    On Error GoTo ErrHandle
    
    Dim lHelpContextID As Long
    Dim rsSettleDetail As Recordset

    Dim m_oReport As New Report
    Dim vCustomData As Variant
    Dim frmNewReport As New frmReport
    
    lHelpContextID = frmBusSettleStat.HelpContextID
    frmBusSettleStat.Show vbModal
    If frmBusSettleStat.m_bOk Then
        m_oReport.Init g_oActiveUser
        '取得记录集
        Set rsSettleDetail = m_oReport.BusSettleStat(frmBusSettleStat.m_dtStartDate, DateAdd("d", 1, frmBusSettleStat.m_dtEndDate), ResolveDisplay(frmBusSettleStat.m_szBusID), ResolveDisplay(frmBusSettleStat.m_szCompanyID), , , frmBusSettleStat.m_bStatBySettleDate)
        WriteProcessBar True, , , ""
        ReDim vCustomData(1 To 7, 1 To 2)
        vCustomData(1, 1) = "开始日期"
        vCustomData(1, 2) = ToDBDate(frmBusSettleStat.m_dtStartDate)
        vCustomData(2, 1) = "结束日期"
        vCustomData(2, 2) = ToDBDate(frmBusSettleStat.m_dtEndDate)
        vCustomData(3, 1) = "打印"
        vCustomData(3, 2) = g_oActiveUser.UserName
        vCustomData(4, 1) = "车次"
        vCustomData(4, 2) = IIf(frmBusSettleStat.m_szBusID = "", "所有车次", frmBusSettleStat.m_szBusTagNo)
        
        vCustomData(5, 1) = "公司"
        vCustomData(5, 2) = IIf(frmBusSettleStat.m_szCompanyID = "", "所有公司", frmBusSettleStat.m_szCompanyName)
        
        vCustomData(6, 1) = "上车站"
        vCustomData(6, 2) = g_oActiveUser.UserUnitName
        
        vCustomData(7, 1) = "统计方式"
        vCustomData(7, 2) = IIf(frmBusSettleStat.m_bStatBySettleDate, "按结算日期统计", "按车次日期统计")
        
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport rsSettleDetail, cszBusSettleStat, frmBusSettleStat.Caption, vCustomData, 10
        
        WriteProcessBar False, , , ""
    End If
    
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub


'车辆结算汇总月报
Private Sub mnu_VehicleSettleStatByMonth_Click()
    On Error GoTo ErrHandle
    
    Dim lHelpContextID As Long
    Dim rsSettleDetail As Recordset

    Dim m_oReport As New Report
    Dim vCustomData As Variant
    Dim frmNewReport As New frmReport
    
    lHelpContextID = frmVehicleSettleStat.HelpContextID
    frmVehicleSettleStat.Show vbModal
    If frmVehicleSettleStat.m_bOk Then
        m_oReport.Init g_oActiveUser
        '取得记录集
        Set rsSettleDetail = m_oReport.VehicleSettleStatByMonth(frmVehicleSettleStat.m_dtStartDate, DateAdd("d", 1, frmVehicleSettleStat.m_dtEndDate), ResolveDisplay(frmVehicleSettleStat.m_szVehicleID), ResolveDisplay(frmVehicleSettleStat.m_szCompanyID), frmVehicleSettleStat.m_nQueryNegativeType, frmVehicleSettleStat.m_nStatus)
        WriteProcessBar True, , , ""
        ReDim vCustomData(1 To 6, 1 To 2)
        vCustomData(1, 1) = "开始日期"
        vCustomData(1, 2) = ToDBDate(frmVehicleSettleStat.m_dtStartDate)
        vCustomData(2, 1) = "结束日期"
        vCustomData(2, 2) = ToDBDate(frmVehicleSettleStat.m_dtEndDate)
        vCustomData(3, 1) = "打印"
        vCustomData(3, 2) = g_oActiveUser.UserName
        vCustomData(4, 1) = "车辆"
        vCustomData(4, 2) = IIf(frmVehicleSettleStat.m_szVehicleID = "", "所有车辆", frmVehicleSettleStat.m_szVehicleTagNo)
        
        vCustomData(5, 1) = "公司"
        vCustomData(5, 2) = IIf(frmVehicleSettleStat.m_szCompanyID = "", "所有公司", frmVehicleSettleStat.m_szCompanyName)
        
        vCustomData(6, 1) = "上车站"
        vCustomData(6, 2) = g_oActiveUser.UserUnitName
        
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport rsSettleDetail, cszVehicleSettleStatByMonth, frmVehicleSettleStat.Caption, vCustomData, 10
        
        WriteProcessBar False, , , ""
    End If
    
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub




'协议
Private Sub mnu_Protocol()
    On Error GoTo ErrHandle
    
    Dim lHelpContextID As Long
    
    Dim m_oReport As New Report
    Dim vCustomData As Variant
    Dim rsTemp As Recordset
    Dim frmNewReport As New frmReport
    
    m_oReport.Init g_oActiveUser
    '取得记录集
    Set rsTemp = m_oReport.GetAllProtocolRS()
    
    
    frmNewReport.m_lHelpContextID = lHelpContextID
    frmNewReport.ShowReport rsTemp, cszProtocol, "协议明细", vCustomData
    

Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

'车辆协议
Private Sub mnu_VehicleProtocol()
    On Error GoTo ErrHandle
    
    Dim lHelpContextID As Long
    
    Dim m_oReport As New Report
    Dim vCustomData As Variant
    Dim rsTemp As Recordset
    Dim frmNewReport As New frmReport
    
    m_oReport.Init g_oActiveUser
    '取得记录集
    Set rsTemp = m_oReport.GetVehicleProtocolRS()
    
    
    frmNewReport.m_lHelpContextID = lHelpContextID
    frmNewReport.ShowReport rsTemp, cszVehicleProtocol, "车辆协议明细", vCustomData
    

Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

'公司协议
Private Sub mnu_CompanyProtocol()
    On Error GoTo ErrHandle
    
    Dim lHelpContextID As Long
    
    Dim m_oReport As New Report
    Dim vCustomData As Variant
    Dim rsTemp As Recordset
    Dim frmNewReport As New frmReport
    
    m_oReport.Init g_oActiveUser
    '取得记录集
    Set rsTemp = m_oReport.GetAllCompanyProtocolRS()
    
    
    frmNewReport.m_lHelpContextID = lHelpContextID
    frmNewReport.ShowReport rsTemp, cszCompanyProtocol, "公司协议明细", vCustomData
    

Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

'公司结算价
Private Sub mnu_CompanySettlePrice()
    On Error GoTo ErrHandle
    
    Dim lHelpContextID As Long
    
    Dim m_oReport As New Report
    Dim vCustomData As Variant
    Dim rsTemp As Recordset
    Dim frmNewReport As New frmReport
    
    m_oReport.Init g_oActiveUser
    '取得记录集
    Set rsTemp = m_oReport.GetCompanySettlePriceLstRS()
    
    
    frmNewReport.m_lHelpContextID = lHelpContextID
    frmNewReport.ShowReport rsTemp, cszCompanySettlePrice, "公司结算价明细", vCustomData
    

Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub


'车辆结算价
Private Sub mnu_VehicleSettlePrice()
    On Error GoTo ErrHandle
    
    Dim lHelpContextID As Long
    
    Dim m_oReport As New Report
    Dim vCustomData As Variant
    Dim rsTemp As Recordset
    Dim frmNewReport As New frmReport
    
    m_oReport.Init g_oActiveUser
    '取得记录集
    Set rsTemp = m_oReport.GetVehicleSettlePriceLstRS()
    
    
    frmNewReport.m_lHelpContextID = lHelpContextID
    frmNewReport.ShowReport rsTemp, cszVehicleSettlePrice, "车辆结算价明细", vCustomData
    

Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub







Private Sub mnu_About_Click()
    Dim oShell As New CommShell
    oShell.ShowAbout App.ProductName, "Settle System", App.FileDescription, Me.Icon, App.Major, App.Minor, App.Revision
End Sub

Private Sub ChangePassword()
    Dim oShell As New CommDialog
    On Error GoTo ErrorHandle
    oShell.Init g_oActiveUser
    oShell.ShowUserInfo
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub
Private Sub mnu_HelpContent_Click()
    If Not ActiveForm Is Nothing Then
        DisplayHelp ActiveForm, content
    Else
        DisplayHelp Me
    End If
End Sub

Private Sub mnu_HelpIndex_Click()
    DisplayHelp Me, Index
End Sub

Private Sub cmdClose_Click()
On Error Resume Next
    If Not ActiveForm Is Nothing Then
'        ActiveToolBar "baseinfo", True
        Unload ActiveForm
    End If
End Sub
Private Sub ExitSystem()
    If MsgBox("您是否真的要退出本系统?", vbQuestion + vbYesNoCancel, "问题") = vbYes Then
        Unload Me
    End If
End Sub
Private Sub mnu_TitleH_Click()
    Arrange vbTileHorizontal
End Sub

Private Sub mnu_TitleV_Click()
    Arrange vbTileVertical
End Sub
Private Sub mnu_Cascade_Click()
    Arrange vbCascade
End Sub
Private Sub MDIForm_Resize()
    On Error Resume Next
'    cmdClose.Left = Me.Width - cmdClose.Width - 2000

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    End
End Sub

Private Sub mnu_vehiclebalancestat_Click()
    '车辆平衡表
    
    Dim lHelpContextID As Long
    lHelpContextID = frmVehicleBalance.HelpContextID
    
    frmVehicleBalance.Show vbModal, Me
    If frmVehicleBalance.m_bOk Then
        Dim frmTemp As IConditionForm
        Dim frmNewReport As New frmReport
        Set frmTemp = frmVehicleBalance
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport frmTemp.RecordsetData, frmTemp.FileName, "车辆平衡表", frmTemp.CustomData, 2
        
    End If
    
End Sub

Private Sub mnu_busbalancestat_Click()
    '车次平衡表
    
    Dim lHelpContextID As Long
    lHelpContextID = frmBusBalance.HelpContextID
    
    frmBusBalance.Show vbModal, Me
    If frmBusBalance.m_bOk Then
        Dim frmTemp As IConditionForm
        Dim frmNewReport As New frmReport
        Set frmTemp = frmBusBalance
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport frmTemp.RecordsetData, frmTemp.FileName, "车次平衡表", frmTemp.CustomData, 2
        
    End If
    
End Sub


Private Sub MakeSheetStatTemp()
    Dim oSplit As New Split
    On Error GoTo ErrorHandle
    SetBusy
    ShowSBInfo "正在生成当天结算的临时统计数据"
    abMenu.Refresh
    
    oSplit.Init g_oActiveUser
    oSplit.MakeSheetStatTemp Date
    SetNormal
    ShowSBInfo ""
    MsgBox "生成当天结算的临时统计数据完毕", vbInformation, Me.Caption
    Exit Sub
ErrorHandle:
    SetNormal
    ShowSBInfo ""
    ShowErrorMsg
End Sub


