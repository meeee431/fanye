VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "actbar2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm MDIScheme 
   BackColor       =   &H8000000C&
   Caption         =   "�೵����"
   ClientHeight    =   6660
   ClientLeft      =   2025
   ClientTop       =   2550
   ClientWidth     =   9285
   HelpContextID   =   2009201
   Icon            =   "MDIScheme.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   WindowState     =   2  'Maximized
   Begin ActiveBar2LibraryCtl.ActiveBar2 abMenuTool 
      Align           =   1  'Align Top
      Height          =   6660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9285
      _LayoutVersion  =   1
      _ExtentX        =   16378
      _ExtentY        =   11748
      _DataPath       =   ""
      Bands           =   "MDIScheme.frx":16AC2
      Begin MSComctlLib.ProgressBar pbLoad 
         Height          =   225
         Left            =   6120
         TabIndex        =   1
         Top             =   5010
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin MSComctlLib.ImageList ilBig 
      Left            =   2250
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIScheme.frx":27D4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIScheme.frx":28068
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIScheme.frx":28944
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIScheme.frx":28AA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIScheme.frx":28C04
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIScheme.frx":28D64
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIScheme.frx":28EC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIScheme.frx":297A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIScheme.frx":2A07C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgMdi 
      Left            =   3630
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   32
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIScheme.frx":2A958
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIScheme.frx":2AAB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIScheme.frx":2AF08
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIScheme.frx":2B35C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIScheme.frx":2B4B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIScheme.frx":2B614
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIScheme.frx":2B770
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIScheme.frx":2B8CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIScheme.frx":2BA28
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIScheme.frx":2BB84
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIScheme.frx":2BCE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIScheme.frx":2BE3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIScheme.frx":2BF98
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIScheme.frx":2C0F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIScheme.frx":2C250
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIScheme.frx":2C3AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIScheme.frx":2C508
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIScheme.frx":2C664
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIScheme.frx":2C7C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIScheme.frx":2C91C
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIScheme.frx":2CA78
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIScheme.frx":2CBD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIScheme.frx":2CD30
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIScheme.frx":2CE8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIScheme.frx":2CFE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIScheme.frx":2D144
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIScheme.frx":2D2A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIScheme.frx":2D3FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIScheme.frx":2D558
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIScheme.frx":2DE34
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIScheme.frx":2DF48
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIScheme.frx":2E05C
            Key             =   "AddNewREBus"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   2910
      Top             =   1470
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIScheme.frx":2E1B8
            Key             =   "Export"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIScheme.frx":2E314
            Key             =   "ExportOpen"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIScheme.frx":2E470
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIScheme.frx":2E584
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdPrintSetup 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "MDIScheme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private WithEvents moMessage As StNotify.MsgNotify  '��Ϣ���ն���
'Dim meEventMode As eEventId
'Dim aszEventParam(1 To 6) As String

Private Sub abMenuTool_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
On Error GoTo ErrHandle
    'ִ�ж�Ӧ�Ĳ˵�������ע��˵�����
    Select Case Tool.name
        Case "mnu_BaseInfo", "tbn_scheme_baseInfo"
            frmBaseInfo.ZOrder 0
            frmBaseInfo.Show
            Case "mnu_BaseMan_Add"
                frmBaseInfo.AddObject
            Case "mnu_BaseMan_BaseInfo"
                frmBaseInfo.EditObject
            Case "mnu_BaseMan_Del"
                frmBaseInfo.DeleteObject
        Case "mnu_BusPlanInfo", "tbn_scheme_busplan"
            frmBus.ZOrder 0
            frmBus.Show
            Case "mnu_BusPlanMan_Info"
                frmBus.EditBus
            Case "mnu_BusPlanMan_Price"
                frmBus.BusTicketPrice
            Case "mnu_BusPlanMan_Envir"
                frmBus.EnvPreview
            Case "mnu_BusPlanMan_Stop"
                frmBus.StopBus
            Case "mnu_BusPlanMan_Resume"
                frmBus.ResumeBus
            Case "mnu_BusPlanMan_Add"
                frmBus.AddBus
            Case "mnu_BusPlanMan_Copy"
                frmBus.CopyBus
            Case "mnu_BusPlanMan_Del"
                frmBus.DeleteBus
        Case "mnu_BusEnvInfo", "tbn_scheme_busenv"
            frmEnvBus.ZOrder 0
            frmEnvBus.Show
            Case "mnu_BusEnvMan_Info"
                frmEnvBus.EditBus
            Case "mnu_BusEnvMan_Price"
                frmEnvBus.BusTicketPrice
            Case "mnu_BusEnvMan_Stop"
                frmEnvBus.StopBus False
            Case "mnu_BusEnvMan_Resume"
                frmEnvBus.ResumeBus
            Case "mnu_BusEnvMan_Replace"
                frmEnvBus.ReplaceBus
            Case "mnu_BusEnvMan_Merge"
                frmEnvBus.MergeBus
            Case "mnu_BusEnvMan_Add"
                frmEnvBus.AddBus
            Case "mnu_BusEnvMan_Copy"
'                frmEnvBus.CopyBus
                frmCopyEvnBus.Show
            Case "mnu_BusEnvMan_Del"
                frmEnvBus.DeleteBus
            Case "mnu_BusEnvMan_Seat"
                frmEnvBus.EnvSeat
        Case "mnu_BusWizard", "tbn_scheme_buswizard"
            frmWizardAddBus.Show vbModal
        Case "mnu_BusBuildEnv"
            BuildEnv
        Case "mnu_ParamSet"
            frmOption.Show vbModal
        Case "mnu_StationInfo", "tbn_station_info"
            frmAllStation.ZOrder 0
            frmAllStation.Show
            Case "mnu_StationMan_Info"
                frmAllStation.EditStation
            Case "mnu_StationMan_Add"
                frmAllStation.AddStation
            Case "mnu_StationMan_Del"
                frmAllStation.DeleteStation

        Case "mnu_SectionInfo"
            frmAllSection.ZOrder 0
            frmAllSection.Show
        Case "mnu_RouteInfo", "tbn_route_info"
            frmAllRoute.ZOrder 0
            frmAllRoute.Show
            Case "mnu_RouteMan_Info"
                frmAllRoute.EditRoute
            Case "mnu_RouteMan_Section"
                frmAllRoute.RouteSection
            Case "mnu_RouteMan_Add"
                frmAllRoute.AddRoute
            Case "mnu_RouteMan_Copy"
                frmAllRoute.CopyRoute
            Case "mnu_RouteMan_Del"
                frmAllRoute.DeleteRoute
        Case "mnu_RouteRatio"
            '����
            frmShowRatio.Show vbModal

        '������Ʊ�۲���
        Case "mnu_TicketPriceInfo", "tbn_scheme_bustkprice"
            OpenBusPrice
        Case "mnu_TicketPriceMan_Save"
            '���泵��Ʊ��
            SaveBusPrice
        Case "mnu_TicketPriceMan_Open"
            '�򿪳���Ʊ��
            ShowDialog
        Case "mnu_TicketPriceMan_AddManual"
            '��������Ʊ��
            AddBusPriceManual
        Case "mnu_TicketPriceMan_AddAuto"
            AddBusPriceAuto
        Case "mnu_TicketPriceMan_Del"
            DeleteBusPrice
        Case "mnu_TicketPriceMan_Modify"
            ActiveForm.BatchModify
        Case "mnu_EnvirPriceInfo", "tbn_scheme_envtkprice"
            OpenEnvPriceInfo
        Case "mnu_PriceTableMan", "tbn_scheme_pricetable"
            frmPriceTableMan.Show vbModal
        Case "mnu_PriceItemSet"
            frmPriceItem.Show vbModal
        Case "mnu_PriceTableCopy", "tbn_scheme_tablecopy"
            frmCopyPriceTable.Show vbModal
        Case "mnu_PriceTableFormula"
            frmFormulaMan.Show vbModal
        Case "mnu_RoutePriceSet"
            frmSetRouteFormula.Show vbModal
        Case "mnu_HalfPriceSet"
            frmTicketTypeEdit.Show vbModal
        Case "mnu_PriceTailSet"
            frmMantissa.Show vbModal
        Case "mnu_PriceAreaTailSet"
            frmAreaTailMethod.Show vbModal
        Case "mnu_DirectModifyPrice"
            'ֱ�������޸����ݿ��е�Ʊ��
            frmDirectModifyPrice.Show vbModal
            
        '���¼�Ʊ����
        Case "mnu_CheckDoorMonitor", "tbn_scheme_chkmoniter"
            '��Ʊ��״̬
            frmGateMoniter.ZOrder 0
            frmGateMoniter.Show
        Case "mnu_CheckPersonSet"
            frmCheckerMan.Show vbModal
        Case "mnu_CheckUnregister"
            frmWriteOffCheck.Show vbModal

        '���²�ѯ��������
        Case "mnu_QuerySellBus", "tbn_scheme_bustkquery"
            QuerySellBus
        Case "mnu_QuerySellStation", "tbn_scheme_stantkquery"
            '��վ��Ʊ��ѯ
            frmQuerySellUnit.ZOrder 0
            frmQuerySellUnit.Show
        Case "mnu_BusReport"
            frmReportScheme.ZOrder 0
            frmReportScheme.Show
        Case "mnu_BusInfoReport" '�ƻ�������Ϣ����
            frmBusInfoReport.ZOrder 0
            frmBusInfoReport.Show
        Case "mnu_ByWayBusReport" ';����α�
            frmByWayBusReport.ZOrder 0
            frmByWayBusReport.Show
        Case "mnu_PriceReport_BusPrice"
            frmReportBus.ZOrder 0
            frmReportBus.Show
        Case "mnu_PriceReport_StationPrice"
            frmReportStation.ZOrder 0
            frmReportStation.Show
        Case "mnu_PriceReport_EnvBusPrice"
            frmReportEnv.ZOrder 0
            frmReportEnv.Show
        Case "mnu_QueryCheck"
            frmChkTkQuery.Show vbModal
        Case "mnu_CheckReport"
            frmReportCheck.ZOrder 0
            frmReportCheck.Show
    Case "mnu_ViewCheckSheet"
        '�쿴·��
        frmAllSheet.m_dtStartDate = GetFirstMonthDay(Date)
        frmAllSheet.m_dtEndDate = GetLastMonthDay(Date)
        frmAllSheet.ZOrder 0
        frmAllSheet.Show
        
    Case "mnu_QueryBusRealNameInfo"
        frmQUeryBusRealNameInfo.ZOrder 0
        frmQUeryBusRealNameInfo.Show

        '������ϵͳ����
        Case "tbn_system_print"
            If MsgBox("ȷ��Ҫ��ӡ��", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
                ActiveForm.PrintReport False
            End If
        Case "mnu_system_print"
            If MsgBox("ȷ��Ҫ��ӡ��", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
                ActiveForm.PrintReport True
            End If
        Case "tbn_system_printview", "mnu_system_printview"
            ActiveForm.PreView
        Case "mnu_PageOption"
            'ҳ������
            ActiveForm.PageSet
        Case "mnu_PrintOption"
            '��ӡ����
            ActiveForm.PrintSet
        Case "tbn_system_export", "mnu_ExportFile"
            ActiveForm.ExportFile
        Case "tbn_system_exportopen", "mnu_ExportFileOpen"
            ActiveForm.ExportFileOpen
        Case "mnu_SysExit", "tbn_system_exit"
            Unload Me
        Case "mnu_ChgPassword"
            '�޸Ŀ���
            ChangePassword
        Case "mnu_HelpAbout"
            '����
            ShowAbout
    '�����Ǵ�������
        Case "mnu_Horizontal"
            Me.Arrange 1
        Case "mnu_Vertical"
            Me.Arrange 2
        Case "mnu_Cascade"
            Me.Arrange 0
        Case "mnu_icon"
            Me.Arrange vbArrangeIcons
        Case "mnu_HelpContent"
            MDIScheme.HelpContextID = 10000730
            DisplayHelp Me
    End Select
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub MDIForm_Load()
    '��ʼ������
    AddControlsToActBar


    '��ʼ���˵���������
    ActiveSystemToolBar False
    ActiveToolBar "planbus", False
    ActiveToolBar "envbus", False
    ActiveToolBar "baseinfo", False
    ActiveToolBar "station", False
    ActiveToolBar "route", False
    ActiveToolBar "busprice", False

    WriteProcessBar False

    '״̬��
    ShowSBInfo "", ESB_WorkingInfo
    ShowSBInfo "", ESB_ResultCountInfo
    ShowSBInfo EncodeString(g_oActiveUser.UserID) & g_oActiveUser.UserName, ESB_UserInfo
    ShowSBInfo Format(g_oActiveUser.LoginTime, "HH:mm"), ESB_LoginTime

    '������
'    WriteTitleBar

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
'If frmSalelAndSlitpInfo.m_bIsShow = True Then
'   Unload frmSalelAndSlitpInfo
'End If
'End
End Sub





'����ActiveBar�Ŀؼ�
Private Sub AddControlsToActBar()
'    abMenuTool.Bands("bndTitleBar").Tools("ptTitle").Custom = ptTitle
    abMenuTool.Bands("statusBar").Tools("progressBar").Custom = pbLoad
End Sub
'�����Ӧ�Ĺ�����
Public Sub ActiveToolBar(pszToolGroupName As String, pbTrue As Boolean)
    Select Case LCase(pszToolGroupName)
        Case "planbus"
            abMenuTool.Bands("mnu_BusMan").Tools("mnu_BusPlanMan").Enabled = pbTrue
        Case "envbus"
            abMenuTool.Bands("mnu_BusMan").Tools("mnu_BusEnvMan").Enabled = pbTrue
        Case "baseinfo"
            abMenuTool.Bands("mnu_System").Tools("mnu_BaseMan").Enabled = pbTrue
        Case "station"
            abMenuTool.Bands("mnu_Route").Tools("mnu_StationMan").Enabled = pbTrue
        Case "route"
            abMenuTool.Bands("mnu_Route").Tools("mnu_RouteMan").Enabled = pbTrue
        Case "busprice"
'            abMenuTool.Bands("mnu_TicketPrice").Tools("mnu_TicketPriceMan").Enabled = pbTrue
    End Select
End Sub


Private Sub BuildEnv()
    '���ɻ���
    Dim oReg As New CFreeReg
    Dim szPosition As String
    On Error GoTo ErrorHandle
    oReg.Init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany
    szPosition = oReg.GetSetting("Scheme", "MakeEnPosition", """C:\Program Files\TEST\MakeEn\PSTMakeEn.exe""")
    Shell szPosition & " " & g_oActiveUser.UserID & "," & g_szUserPassword
ErrorHandle:

End Sub

Private Sub OpenBusPrice()
    '����Ʊ�۹���
    Dim ofrm As New frmModifyBusPrice
    ofrm.m_eFormStatus = EFS_Modify
'    oFrm.Show
    Load ofrm
End Sub

Private Sub SaveBusPrice()
    '���泵��Ʊ��
    ActiveForm.SaveBusPrice
End Sub

Private Sub ShowDialog()
    '�򿪳���Ʊ��
    ActiveForm.ShowOpenDialog
End Sub

Private Sub AddBusPriceManual()
    '�ֹ�����Ʊ��
    Dim ofrm As New frmModifyBusPrice
    ofrm.m_eFormStatus = EFS_AddNew
'    oFrm.Show
    Load ofrm
End Sub

Private Sub AddBusPriceAuto()
    '�Զ�����Ʊ��,ֱ��д�����ݿ�
    frmShowBus.m_bEnabledStop = True
    frmShowBus.m_eFormStatus = EFS_AddNew
    frmShowBus.Show vbModal
    If frmShowBus.m_bOk Then
        '�������OK
        frmMakeBusPrice.GetBusID = frmShowBus.GetBusID
        frmMakeBusPrice.GetIsOnlyStop = frmShowBus.IsOnlyStop
        frmMakeBusPrice.GetPriceTableID = frmShowBus.GetPriceTableID
        frmMakeBusPrice.GetSeatType = frmShowBus.GetSeatType
        frmMakeBusPrice.GetVehicleType = frmShowBus.GetVehicleType
        frmMakeBusPrice.Show vbModal
    End If
End Sub

Private Sub DeleteBusPrice()
    'ɾ������Ʊ��
    Dim oPriceTableID As New STPrice.RoutePriceTable
    Dim aszBusID() As String
    Dim aszVehicleType() As String
    Dim aszSeatType() As String
    Dim atBusVehicleSeatType() As TBusVehicleSeatType

    frmShowBus.m_eFormStatus = EFS_Delete
    frmShowBus.Show vbModal
    If frmShowBus.m_bOk Then
        '�������OK
        If MsgBox("ȷʵҪɾ����ѡ��ĳ��ε�Ʊ����", vbYesNo + vbDefaultButton2 + vbQuestion, "ɾ��Ʊ��") = vbYes Then
            '��ʼɾ��Ʊ��
            SetBusy
            ShowSBInfo "����ɾ������Ʊ��..."
            oPriceTableID.Init g_oActiveUser
            oPriceTableID.Identify frmShowBus.GetPriceTableID
            aszBusID = frmShowBus.GetBusID
            aszSeatType = frmShowBus.GetSeatType
            aszVehicleType = frmShowBus.GetVehicleType
            atBusVehicleSeatType = ConvertTypeFromArray(aszBusID, aszVehicleType, aszSeatType)
            oPriceTableID.DeleteBusPrice atBusVehicleSeatType
            MsgBox "ɾ������Ʊ�۳ɹ���", vbInformation, "ɾ��Ʊ��"
            ShowSBInfo ""
            SetNormal
        End If
    End If
    Exit Sub
    ShowSBInfo ""
    SetNormal
End Sub

Private Sub OpenEnvPriceInfo()
    '�򿪻���������Ϣ
    Dim ofrm As New frmModifyREPrice
    ofrm.m_eFormStatus = EFS_Modify
    ofrm.Show
End Sub

Public Sub SetPrintEnabled(pbEnabled As Boolean)
    '���ò˵��Ŀ�����
    With MDIScheme.abMenuTool
        .Bands("tbn_system").Tools("tbn_system_print").Enabled = pbEnabled
        .Bands("tbn_system").Tools("tbn_system_printview").Enabled = pbEnabled
        .Bands("tbn_system").Tools("tbn_system_export").Enabled = pbEnabled
        .Bands("tbn_system").Tools("tbn_system_exportopen").Enabled = pbEnabled
        .Bands("mnu_System").Tools("mnu_PageOption").Enabled = pbEnabled
        .Bands("mnu_System").Tools("mnu_PrintOption").Enabled = pbEnabled
        .Bands("mnu_System").Tools("mnu_system_print").Enabled = pbEnabled
        .Bands("mnu_System").Tools("mnu_system_printview").Enabled = pbEnabled
        .Bands("mnu_System").Tools("mnu_ExportFile").Enabled = pbEnabled
        .Bands("mnu_System").Tools("mnu_ExportFileOpen").Enabled = pbEnabled
    End With
End Sub

Private Sub QuerySellBus()
    '������Ʊ��ѯ ,Ϊ�˼Ӱ���
    frmQuerySellBus.ZOrder 0
    frmQuerySellBus.Show
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

Private Sub ShowAbout()
    Dim oShell As New CommShell
    On Error GoTo ErrorHandle
    oShell.ShowAbout "�೵����", "Bus Scheme System", "�೵����ϵͳ", Me.Icon, App.Major, App.Minor, App.Revision
    Set oShell = Nothing
    Exit Sub
ErrorHandle:
    ShowErrorMsg
    Set oShell = Nothing
End Sub

