VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "actbar2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.MDIForm MDIMain 
   BackColor       =   &H8000000C&
   Caption         =   "ͳ�Ʒ���"
   ClientHeight    =   5385
   ClientLeft      =   2565
   ClientTop       =   3180
   ClientWidth     =   8280
   HelpContextID   =   6000001
   Icon            =   "MDIMain.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   WindowState     =   2  'Maximized
   Begin ActiveBar2LibraryCtl.ActiveBar2 abMenu 
      Align           =   1  'Align Top
      Height          =   5385
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8280
      _LayoutVersion  =   1
      _ExtentX        =   14605
      _ExtentY        =   9499
      _DataPath       =   ""
      Bands           =   "MDIMain.frx":16AC2
      Begin VB.PictureBox ptTitleTop 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   -690
         Picture         =   "MDIMain.frx":27E46
         ScaleHeight     =   687.72
         ScaleMode       =   0  'User
         ScaleWidth      =   15405
         TabIndex        =   2
         Top             =   2370
         Width           =   15405
         Begin RTComctl3.CoolButton cmdClose 
            Height          =   390
            Left            =   8280
            TabIndex        =   3
            ToolTipText     =   "����"
            Top             =   210
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   688
            BTYPE           =   12
            TX              =   ""
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   3
            FOCUSR          =   0   'False
            BCOL            =   12632256
            BCOLO           =   12632256
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "MDIMain.frx":2B6FB
            PICN            =   "MDIMain.frx":2B717
            PICH            =   "MDIMain.frx":2C60C
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   1
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   5610
            TabIndex        =   4
            Top             =   360
            Width           =   120
         End
      End
      Begin MSComctlLib.ProgressBar pbLoad 
         Height          =   225
         Left            =   2490
         TabIndex        =   1
         Top             =   4020
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin MSComDlg.CommonDialog cdPrintSetup 
      Left            =   5220
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imglCell 
      Left            =   1800
      Top             =   2100
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":2D72F
            Key             =   "exporttofile"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":2D889
            Key             =   "open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":2D99B
            Key             =   "save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":2DAAD
            Key             =   "exporttofileandopen"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":2DC07
            Key             =   "printpreview"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":2DD61
            Key             =   "print"
         EndProperty
      EndProperty
   End
   Begin VB.Menu pmnu_Combine 
      Caption         =   "���"
      Visible         =   0   'False
      Begin VB.Menu pmnu_AddCombine 
         Caption         =   "����(&A)"
      End
      Begin VB.Menu pmnu_DeleteCombine 
         Caption         =   "ɾ��(&D)"
      End
   End
   Begin VB.Menu pmnu_BusTrans 
      Caption         =   "��������"
      Visible         =   0   'False
      Begin VB.Menu pmnu_SelectAll 
         Caption         =   "ѡ�����г���(&S)"
      End
      Begin VB.Menu pmnu_Query 
         Caption         =   "��ѯ(&Q)"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu pmnu_SelectSaler 
      Caption         =   "ѡ����ƱԱ"
      Visible         =   0   'False
      Begin VB.Menu pmnu_AddPreSaler 
         Caption         =   "�����ʼ��ƱԱ(&P)"
      End
      Begin VB.Menu pmnu_RemoveSelSaler 
         Caption         =   "�Ƴ�ѡ����ƱԱ(&S)"
      End
      Begin VB.Menu pmnu_RemoveAllSaler 
         Caption         =   "�Ƴ�ȫ��(&A)"
      End
   End
   Begin VB.Menu pmnu_SelectSaler2 
      Caption         =   "ѡ����ƱԱ"
      Visible         =   0   'False
      Begin VB.Menu pmnu_AddSinceSaler 
         Caption         =   "��ӽ�����ƱԱ(&N)"
      End
      Begin VB.Menu pmnu_RemoveSelSaler2 
         Caption         =   "�Ƴ�ѡ����ƱԱ(&S)"
      End
      Begin VB.Menu pmnu_RemoveAllSaler2 
         Caption         =   "�Ƴ�ȫ��(&A)"
      End
   End
End
Attribute VB_Name = "MDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const cszSellerSimple = "��ƱԱ��Ʊ�ձ�"
Const cszSellerSimpleMonth = "��ƱԱ��Ʊ��"
Const cszSellerInterSimple = "��ƱԱ������Ʊ��"
Const cszUnitProxySimple = "����λ��Ʊ��"
Const cszSellerPriceItemCon = "��ƱԱƱ�����"

Const cszUnitDaily = "��վ��ƱӪ���ձ�"
Const cszUnitMonthly = "��վ��ƱӪ���±�"
Const cszUnitInterTotal = "��վ������Ʊͳ�Ʊ���"
Const cszUnitInterSimple = "������λ��Ʊ��"
Const cszUnitInterSell = "������λ����ͳ�Ʊ���"
Const cszSellerEveryDay = "��ƱԱÿ�ս���"
Const cszCheckTicketDetail = "��Ʊ��ϸ��ѯ"
Const cszSellTicketDetail = "��ƱԱ��Ʊ��ϸ��ѯ"
Const cszChangeTicketDetail = "��ƱԱǩ֤��ϸ��ѯ"
Const cszReturnTicketDetail = "��ƱԱ��Ʊ��ϸ��ѯ"
Const cszCancelTicketDetail = "��ƱԱ��Ʊ��ϸ��ѯ"
Const cszSellerSomeDaysTotal = "��ƱԱ�ڼ���Ʊͳ��"
Const cszBusSationSellDetail = "����վ����Ʊ��ϸ"
Const cszBusSationSellCount = "����վ����Ʊͳ��"
Const cszBusSellCount = "������Ʊͳ��"
Const cszSationSellCount = "վ����Ʊͳ��"

Const cszStationMonthly = "վ����ƱӪ���±�"
Const cszBusMonthly = "������ƱӪ���±�"

Const cszUnitYearly = "��վ��ƱӪ���걨"
Const cszSellerSimpleYear = "��ƱԱ��Ʊ�걨"
Const cszStationYearly = "վ����ƱӪ���걨"
Const cszBusYearly = "������ƱӪ���걨"

Const cszUnitManSimple = "��վ��������ͳ��"


Const cszUnitInterSettle = "��������ͳ�ƽ��㱨��"


Const cszBusSellStation = "���θ��ϳ�վ��Ʊ��"

Const cszBusStationByBusDate = "���θ�;��վӪ�ռ�"
Const cszBusStationByBusDate1 = "����������������"
Const cszCheckerEveryMonth = "��ƱԱ��Ʊͳ���±�ģ��"

Public m_szMethod As Boolean '(False:����Ʊ��Ture:����Ʊ)




Private Sub abMenu_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
On Error GoTo ErrHandle
Select Case Tool.name
    
    '��ƱԱ����
    Case "mnu_SellerEveryDay"
        mnu_SellerEveryDay_Click
    Case "mnu_SellerEveryMonth"
        mnu_SellerEveryMonth_Click
'    Case "mnu_SellerSimpleDay"
'        mnu_SellerSimpleDay_Click
    Case "mnu_SellerDayCon"
        mnu_SellerDayCon_Click
    Case "mnu_SellerSimpleCon"
        mnu_SellerSimpleCon_Click
        
        
    Case "mnu_SellerPriceItemCon"
        mnu_SellerPriceItemCon_Click
    
    Case "mnu_SellerInterSimple"
        mnu_SellerInterSimple_Click
    Case "mnuSellerTicketDetail"
        mnuSellerTicketDetail_Click
    Case "mnuSellerTicketDetailFromAgent"
        mnuSellerTicketDetailFromAgent_Click
    Case "mnu_UnitProxySimple"
        mnu_UnitProxySimple_Click

    Case "mnuInternetDetailSellTime"
      mnu_InternetTkCount_Click True
    Case "mnuInternetDetailPrintTime" '������Ʊ��ϸ
       mnu_InternetTkCount_Click False
    Case "mnuInternetCoutSellTime"
      mnu_InternetTkDetail_Click True
        
    Case "mnuInternetCountPrintTime"  '������Ʊͳ��
     mnu_InternetTkDetail_Click False
        
    '��վ����
    Case "mnu_UnitDaily"
        mnu_UnitDaily_Click
    Case "mnu_UnitMonthly"
        mnu_UnitMonthly_Click
    Case "mnu_UnitYearly"
        mnu_UnitYearly_Click
    Case "mnu_UnitInterTotal"
        mnu_UnitInterTotal_Click
    Case "mnu_UnitInterSimple"
        mnu_UnitInterSimple_Click
    Case "mnu_UnitInterSell"
        mnu_UnitInterSell_Click
    Case "mnu_UnitInterSettle"
        mnu_UnitInterSettle_Click
    Case "mnu_StationDaily"
        mnu_StationDaily_Click
    Case "mnu_UnitManSimple"
        mnu_UnitManSimple_Click
    Case "mnu_UnitSalePriceSimplle" '��վ��ƱӪ�ռ�
        mnu_UnitSalePriceSimplle_Click
    Case "mnu_Stationsell"
     mnu_Sationsell_Click
    '��������
    Case "mnu_BusBySaleTime"
        mnu_BusBySaleTime_Click
    Case "mnu_BusByBusDate"
        mnu_BusByBusDate_Click
    Case "mnu_BusByBusDateAndSalerStation"
        mnu_BusByBusDateAndSalerStation_Click
    Case "mnu_BusSellStationBySellTime"
        BusSellStationBySaleTime
    Case "mnu_BusSellStationByBusDate"
        BusSellStationByBusDate
        
        
    Case "mnu_BusSellTicket"
        mnu_BusSellTicket_Click
'    Case "mnu_BusTransStat"
'        mnu_BusTransStat_Click
    Case "mnu_BusTransStatBySale" '��������ͳ�ư���Ʊ
        m_szMethod = False
        mnu_BusTransStatBySale_Click
    Case "mnu_BusTransStatByCheck" '��������ͳ�ư���Ʊ
        m_szMethod = True
        mnu_BusTransStatByCheck_Click
'    Case "mnu_BusSomeSum"
    Case "mnu_SomeSumBySellTime"
        mnu_SomeSumBySellTime_Click
    Case "mnu_SomeSumByBusDate"
        mnu_SomeSumByBusDate_Click
    Case "mnu_BusStation"
        mnu_BusStation_Click
    Case "mnu_PreSell" '����Ԥ��Ʊƽ���
        mnu_PreSell_Click
    Case "mnu_PreSellLst" '����Ԥ��Ʊƽ����ϸ
        mnu_PreSellLst_Click
    Case "mnu_BusStationBySellerStation"
        mnu_BusStationBySellerStation_Click
    Case "mnu_BusStationByBusStation"
        mnu_BusStationByBusStation_Click
        
        
    Case "mnu_VehicleSaleBySale"
        m_szMethod = False
        mnu_VehicleSaleBySale_Click
    Case "mnu_VehicleSaleByCheck"
        m_szMethod = True
        mnu_VehicleSaleByCheck_Click
    
    Case "mnu_CompanyBySaleTime"
        mnu_CompanyBySaleTime_Click
    Case "mnu_CompanyBusDate"
        mnu_CompanyBusDate_Click
'    Case "mnu_CompanyFloatSimply"
'        mnu_CompanyFloatSimply_Click
    Case "mnu_CompanyFloatSimplyBySale" '��˾����ͳ�ư���Ʊ
        m_szMethod = False
        mnu_CompanyFloatSimplyBySale_Click
    Case "mnu_CompanyFloatSimplyByCheck" '��˾����ͳ�ư���Ʊ
        m_szMethod = True
        mnu_CompanyFloatSimplyByCheck_Click
        
    Case "mnu_CompanyTransSimply"
        mnu_CompanyTransSimply_Click
    Case "mnu_StationMonthly"
        mnu_StationMonthly_Click
    Case "mnu_StationSellTicket"
        mnu_StationSellTicket_Click
    Case "mnu_RouteAreaSimply"
        mnu_RouteAreaSimply_Click
    Case "mnu_RouteTurnOver"
        mnu_RouteTurnOver_Click
'    Case "mnu_RouteTransport"
'        mnu_RouteTransport_Click
    Case "mnu_RouteTransportBySale" '��·����ͳ�ư���Ʊ
        m_szMethod = False
        mnu_RouteTransportBySale_Click
    Case "mnu_RouteTransportByCheck" '��·����ͳ�ư���Ʊ
        m_szMethod = True
        mnu_RouteTransportByCheck_Click
    Case "pmnu_TicketIssue"
        pmnu_TicketIssue_Click
    Case "mnu_BusCompanyCombineSet"
        mnu_BusCompanyCombineSet_Click
    Case "mnu_ModiyCompanyName"
        mnu_ModiyCompanyName_Click
    
    '��Ʊ����
    Case "mnu_CheckerSheetCon"
        CheckerSheetCon
    
    Case "mnu_CheckBusStationStat"
        mnu_CheckBusStationStat_Click
    Case "mnu_CheckVehicleStationStat"
        mnu_CheckVehicleStationStat_Click
    Case "mnu_CheckerEveryMonth" '��ƱԱ��Ʊͳ���±�
        mnu_CheckerEveryMonth_Click
        
        
    '����
    Case "mnu_TitleH"
        mnu_TitleH_Click
    Case "mnu_TitleV"
        mnu_TitleV_Click
    Case "mnu_Cascade"
        mnu_Cascade_Click
    Case "mnu_ArrangeIcon"
        mnu_ArrangeIcon_Click
    '����
    Case "mnu_HelpIndex"
        mnu_HelpIndex_Click
    Case "mnu_HelpContent"
        mnu_HelpContent_Click
    Case "mnu_About"
        mnu_About_Click
    
        '������ϵͳ����
        Case "tbn_system_print"
            ActiveForm.PrintReport False
        Case "mnu_system_print"
            ActiveForm.PrintReport True
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
        Case "mnu_ChgPassword"
            '�޸Ŀ���
            ChangePassword
        Case "mnu_SysExit", "tbn_system_exit"
            ExitSystem
    End Select
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub
Private Sub ChangePassword()
    Dim oShell As New CommDialog
    On Error GoTo ErrorHandle
    oShell.Init m_oActiveUser
    oShell.ShowUserInfo
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub cmdClose_Click()
On Error Resume Next
    If Not ActiveForm Is Nothing Then
        Unload ActiveForm
    End If
End Sub
Private Sub ExitSystem()
    If MsgBox("���Ƿ����Ҫ�˳���ϵͳ?", vbQuestion + vbYesNoCancel, "����") = vbYes Then
        Unload Me
    End If
End Sub
Private Sub MDIForm_Load()
    AddControlsToActBar
    '״̬��
    ShowSBInfo "", ESB_WorkingInfo
    ShowSBInfo "", ESB_ResultCountInfo
    ShowSBInfo EncodeString(m_oActiveUser.UserID) & m_oActiveUser.UserName, ESB_UserInfo
    ShowSBInfo Format(m_oActiveUser.LoginTime, "HH:mm"), ESB_LoginTime
    
    SetPrintEnabled False
End Sub

Private Sub mnu_About_Click()
    Dim oShell As New CommShell
    oShell.ShowAbout App.ProductName, "TJ", App.FileDescription, Me.Icon, App.Major, App.Minor, App.Revision
End Sub

Private Sub mnu_ArrangeIcon_Click()
    Arrange vbArrangeIcons
End Sub

Private Sub mnu_BusByBusDate_Click()

    Dim lHelpContextID As Long
    frmCompanyBusCon.m_nMode = ST_ByBusStationAndBusDate
    lHelpContextID = frmCompanyBusCon.HelpContextID
    
    frmCompanyBusCon.Show vbModal, Me
    If frmCompanyBusCon.m_bOk Then
        Dim frmTemp As IConditionForm
        Dim frmNewReport As New frmReport
        Set frmTemp = frmCompanyBusCon
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport frmTemp.RecordsetData, frmTemp.FileName, cszBusMonthly & "(���������ڻ���)", frmTemp.CustomData, 2
        
    End If
    
End Sub

Private Sub mnu_BusByBusDateAndSalerStation_Click()
    
    Dim lHelpContextID As Long
    frmCompanyBusCon.m_nMode = ST_BySalerStationAndBusDate
    lHelpContextID = frmCompanyBusCon.HelpContextID
    
    frmCompanyBusCon.Show vbModal, Me
    If frmCompanyBusCon.m_bOk Then
        Dim frmTemp As IConditionForm
        Dim frmNewReport As New frmReport
        Set frmTemp = frmCompanyBusCon
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport frmTemp.RecordsetData, frmTemp.FileName, cszBusMonthly & "(����ƱԱ������վ���������ڻ���)", frmTemp.CustomData, 2
        
    End If
    
    
End Sub



Private Sub mnu_BusBySaleTime_Click()
    Dim lHelpContextID As Long
    frmCompanyBusCon.m_nMode = ST_BySalerStationAndSaleTime
    lHelpContextID = frmCompanyBusCon.HelpContextID
    frmCompanyBusCon.Show vbModal, Me
    If frmCompanyBusCon.m_bOk Then
        Dim frmTemp As IConditionForm
        Dim frmNewReport As New frmReport
        Set frmTemp = frmCompanyBusCon
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport frmTemp.RecordsetData, frmTemp.FileName, cszBusMonthly & "(����Ʊ���ڻ���)", frmTemp.CustomData, 2
    End If
End Sub

Private Sub BusSellStationBySaleTime()
    Dim lHelpContextID As Long
    frmBusSellStationSellInfo.m_bBySaleTime = True
    lHelpContextID = frmBusSellStationSellInfo.HelpContextID
    frmBusSellStationSellInfo.Show vbModal, Me
    If frmBusSellStationSellInfo.m_bOk Then
        Dim frmTemp As IConditionForm
        Dim frmNewReport As New frmReport
        Set frmTemp = frmBusSellStationSellInfo
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport frmTemp.RecordsetData, frmTemp.FileName, cszBusSellStation & "(����Ʊ���ڻ���)", frmTemp.CustomData, 2
    End If
End Sub

Private Sub BusSellStationByBusDate()

    Dim lHelpContextID As Long
    frmBusSellStationSellInfo.m_bBySaleTime = False
    lHelpContextID = frmBusSellStationSellInfo.HelpContextID
    
    frmBusSellStationSellInfo.Show vbModal, Me
    If frmBusSellStationSellInfo.m_bOk Then
        Dim frmTemp As IConditionForm
        Dim frmNewReport As New frmReport
        Set frmTemp = frmBusSellStationSellInfo
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport frmTemp.RecordsetData, frmTemp.FileName, cszBusSellStation & "(���������ڻ���)", frmTemp.CustomData, 2
        
    End If
    
End Sub
Private Sub mnu_BusCompanyCombineSet_Click()
    frmCombine.ZOrder 0
    frmCombine.Show
End Sub

Private Sub mnu_BusSellTicket_Click()
    Dim oTicketAcc As New TicketBusDim
    Dim rsBusStationSell As Recordset
    Dim szbusID() As String
    Dim dtBusDate() As Date
    Dim frmNewReport As frmReport
    Dim vaCostumData As Variant
    
    On Error GoTo errorHander

    frmStationSellTicket.Show vbModal
    If frmStationSellTicket.m_bOk Then
        ReDim vaCostumData(1 To 3, 1 To 2)
        vaCostumData(1, 1) = "��ʼʱ��"
        vaCostumData(1, 2) = Format(frmStationSellTicket.m_dtStartDate, "YYYY-MM-DD")
        vaCostumData(2, 1) = "����ʱ��"
        vaCostumData(2, 2) = Format(frmStationSellTicket.m_dtEndDate, "YYYY-MM-DD")
        
        vaCostumData(3, 1) = "�Ʊ���"
        vaCostumData(3, 2) = m_oActiveUser.UserID
        
        Me.MousePointer = vbHourglass
        oTicketAcc.Init m_oActiveUser
        szbusID = frmStationSellTicket.m_szBus_Id
        dtBusDate = frmStationSellTicket.m_dtBus_Date
        Set frmNewReport = New frmReport
        If frmStationSellTicket.m_bnList = True Then
            Set rsBusStationSell = oTicketAcc.GetBusStationTickets(dtBusDate, szbusID, frmStationSellTicket.m_nOrder)
            frmNewReport.ShowReport rsBusStationSell, "����վ����Ʊ��ϸ��ѯ.xls", cszBusSationSellDetail, vaCostumData, 10
        Else
            If frmStationSellTicket.m_nCount = 0 Then
               Set rsBusStationSell = oTicketAcc.GetBusStationTicketsCount(dtBusDate, szbusID, frmStationSellTicket.m_nCount)
                frmNewReport.ShowReport rsBusStationSell, "������Ʊͳ��.xls", cszBusSellCount, vaCostumData, 10
            ElseIf frmStationSellTicket.m_nCount = 1 Then
               Set rsBusStationSell = oTicketAcc.GetBusStationTicketsCount(dtBusDate, szbusID, frmStationSellTicket.m_nCount)
                frmNewReport.ShowReport rsBusStationSell, "վ����Ʊͳ��.xls", cszSationSellCount, vaCostumData, 10
            Else
                Set rsBusStationSell = oTicketAcc.GetBusStationTicketsCount(dtBusDate, szbusID, frmStationSellTicket.m_nCount)
                frmNewReport.ShowReport rsBusStationSell, "����վ����Ʊͳ��.xls", cszBusSationSellCount, vaCostumData, 10
            End If
        End If
        Set rsBusStationSell = Nothing
        Set frmNewReport = Nothing
        Me.MousePointer = vbDefault
    End If
    Exit Sub
    
errorHander:
    Me.MousePointer = vbDefault
    MsgBox err.Description, vbCritical + vbOKOnly, "����"
End Sub


Private Sub mnu_BusTransStatBySale_Click()
    Dim lHelpContextID As Long
    lHelpContextID = frmBusTransStat.HelpContextID
    
    frmBusTransStat.ZOrder 0
    frmBusTransStat.Show
'    If frmBusTransStat.m_bOk Then
'        Dim frmTemp As IConditionForm
'        Dim frmNewReport As New frmReport
'        Set frmTemp = frmBusTransStat
'        frmNewReport.m_lHelpContextID = lHelpContextID
'        frmNewReport.ShowReport frmTemp.RecordsetData, frmTemp.FileName, "��������ͳ�Ƽ�", frmTemp.CustomData
'    End If
End Sub

Private Sub mnu_BusTransStatByCheck_Click()
    Dim lHelpContextID As Long
    lHelpContextID = frmBusTransStat.HelpContextID
    
    frmBusTransStat.ZOrder 0
    frmBusTransStat.Show
'    If frmBusTransStat.m_bOk Then
'        Dim frmTemp As IConditionForm
'        Dim frmNewReport As New frmReport
'        Set frmTemp = frmBusTransStat
'        frmNewReport.m_lHelpContextID = lHelpContextID
'        frmNewReport.ShowReport frmTemp.RecordsetData, frmTemp.FileName, "��������ͳ�Ƽ�", frmTemp.CustomData
'    End If
End Sub

'Private Sub mnu_BusYearly_Click()
'
'    Dim lHelpContextID As Long
'    lHelpContextID = frmCompanyBusCon.HelpContextID
'
'    frmCompanyBusConYear.Show vbModal, Me
'    If frmCompanyBusConYear.m_bOk Then
'        Dim frmTemp As IConditionForm
'        Dim frmNewReport As New frmReport
'        Set frmTemp = frmCompanyBusConYear
'        frmNewReport.m_lHelpContextID = lHelpContextID
'        frmNewReport.ShowReport frmTemp.RecordsetData, frmTemp.FileName, cszBusYearly, frmTemp.CustomData, 2
'
'    End If
'End Sub

Private Sub mnu_Cascade_Click()
    Arrange vbCascade
End Sub

Private Sub mnu_CompanyBusDate_Click()
    Dim lHelpContextID As Long
    frmSplitCompanySimpleCon.m_bBySaleTime = False
    lHelpContextID = frmSplitCompanySimpleCon.HelpContextID
    
    frmSplitCompanySimpleCon.Show vbModal, Me
    If frmSplitCompanySimpleCon.m_bOk Then
        Dim frmTemp As IConditionForm
        Dim frmNewReport As New frmReport
        Set frmTemp = frmSplitCompanySimpleCon
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport frmTemp.RecordsetData, frmTemp.FileName, "���ʹ�˾Ӫ�ռ�(���������ڻ���)", frmTemp.CustomData
    End If
End Sub

Private Sub mnu_CompanyBySaleTime_Click()
    Dim lHelpContextID As Long
    frmSplitCompanySimpleCon.m_bBySaleTime = True
    lHelpContextID = frmSplitCompanySimpleCon.HelpContextID
    frmSplitCompanySimpleCon.Show vbModal, Me
    If frmSplitCompanySimpleCon.m_bOk Then
        Dim frmTemp As IConditionForm
        Dim frmNewReport As New frmReport
        Set frmTemp = frmSplitCompanySimpleCon
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport frmTemp.RecordsetData, frmTemp.FileName, "���ʹ�˾Ӫ�ռ�(����Ʊ���ڻ���)", frmTemp.CustomData
    End If
End Sub

Private Sub mnu_CompanyFloatSimplyBySale_Click()
    Dim lHelpContextID As Long
    lHelpContextID = frmCompanyFloatSimply.HelpContextID
    frmCompanyFloatSimply.Show vbModal, Me
    If frmCompanyFloatSimply.m_bOk Then
        Dim frmTemp As IConditionForm
        Dim frmNewReport As New frmReport
        Set frmTemp = frmCompanyFloatSimply
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport frmTemp.RecordsetData, frmTemp.FileName, "��˾����ͳ�Ƽ�", frmTemp.CustomData
    End If
End Sub

Private Sub mnu_CompanyFloatSimplyByCheck_Click()
    Dim lHelpContextID As Long
    lHelpContextID = frmCompanyFloatSimply.HelpContextID
    frmCompanyFloatSimply.Show vbModal, Me
    If frmCompanyFloatSimply.m_bOk Then
        Dim frmTemp As IConditionForm
        Dim frmNewReport As New frmReport
        Set frmTemp = frmCompanyFloatSimply
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport frmTemp.RecordsetData, frmTemp.FileName, "��˾����ͳ�Ƽ�", frmTemp.CustomData
    End If
End Sub

Private Sub mnu_CompanyTransSimply_Click()
    Dim lHelpContextID As Long
    lHelpContextID = frmCompanyTranSimply.HelpContextID
    frmCompanyTranSimply.Show vbModal, Me
    If frmCompanyTranSimply.m_bOk Then
        Dim frmTemp As IConditionForm
        Dim frmNewReport As New frmReport
        Set frmTemp = frmCompanyTranSimply
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport frmTemp.RecordsetData, frmTemp.FileName, "��˾�����Ӱ�Ƚϱ�", frmTemp.CustomData
    End If

End Sub

Private Sub mnu_Exit_Click()
    Unload Me
End Sub

'Private Sub mnu_ExportAndOpen_Click()
'    Dim frmTemp As frmReport
'    Set frmTemp = Me.ActiveForm
'    frmTemp.ExportToFileAndOpen
'
'End Sub

Private Sub mnu_ExportFile_Click()
    Dim frmTemp As frmReport
    Set frmTemp = Me.ActiveForm
    frmTemp.ExportToFile
End Sub

Private Sub mnu_HelpContent_Click()
    MDIMain.HelpContextID = 60000340
    DisplayHelp Me
End Sub

Private Sub mnu_HelpIndex_Click()
    DisplayHelp Me, Index
End Sub

Private Sub mnu_ModiyCompanyName_Click()
    frmModifyCompany.Show vbModal
End Sub

'Private Sub mnu_OpenFile_Click()
'    On Error GoTo Error_Handle
'    cdPrintSetup.CancelError = True
'    cdPrintSetup.flags = cdlOFNFileMustExist
'    cdPrintSetup.Filter = "Excel�ļ�(*.xls)|*.xls"
'    cdPrintSetup.InitDir = GetDocumentDir()
'    cdPrintSetup.ShowOpen
'    Dim frmNewReport As New frmReport
'    frmNewReport.m_lHelpContextID = Me.HelpContextID
'    frmNewReport.OpenFile cdPrintSetup.FileName
'    SaveDocumentDir cdPrintSetup.FileName
'    Exit Sub
'Error_Handle:
'
'End Sub

Private Sub mnu_PageSet_Click()
'    ceMain.PageSetup
End Sub

Private Sub mnu_Print_Click()
    Dim frmTemp As frmReport
    Set frmTemp = Me.ActiveForm
    frmTemp.PrintReport
End Sub

Private Sub mnu_PrintPreview_Click()
    Dim frmTemp As frmReport
    Set frmTemp = Me.ActiveForm
    frmTemp.PreView
End Sub

Private Sub mnu_PrintSet_Click()
    On Error GoTo Error_Handle
    cdPrintSetup.flags = cdlPDPrintSetup
    cdPrintSetup.ShowPrinter
    Exit Sub
Error_Handle:
End Sub



Private Sub mnu_RouteAreaSimply_Click()
    Dim lHelpContextID As Long
    lHelpContextID = frmAreaRouteSimply.HelpContextID
    frmAreaRouteSimply.Show vbModal, Me
    If frmAreaRouteSimply.m_bOk Then
        Dim frmTemp As IConditionForm
        Dim frmNewReport As New frmReport
        Set frmTemp = frmAreaRouteSimply
        
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport frmTemp.RecordsetData, frmTemp.FileName, "������·ͳ�Ƽ�", frmTemp.CustomData
        
    End If

    
End Sub

Private Sub mnu_RouteTransportBySale_Click()
    Dim lHelpContextID As Long
    lHelpContextID = frmRouteTransport.HelpContextID
    frmRouteTransport.Show vbModal, Me
    If frmRouteTransport.m_bOk Then
        Dim frmTemp As IConditionForm
        Dim frmNewReport As New frmReport
        Set frmTemp = frmRouteTransport
        
        frmNewReport.m_lHelpContextID = lHelpContextID
        'frmNewReport.NeedMergeCol = 1
        frmNewReport.ShowReport frmTemp.RecordsetData, frmTemp.FileName, "��·����ͳ�Ƽ�", frmTemp.CustomData
        
    End If
End Sub

Private Sub mnu_RouteTransportByCheck_Click()
    Dim lHelpContextID As Long
    lHelpContextID = frmRouteTransport.HelpContextID
    frmRouteTransport.Show vbModal, Me
    If frmRouteTransport.m_bOk Then
        Dim frmTemp As IConditionForm
        Dim frmNewReport As New frmReport
        Set frmTemp = frmRouteTransport
        
        frmNewReport.m_lHelpContextID = lHelpContextID
        'frmNewReport.NeedMergeCol = 1
        frmNewReport.ShowReport frmTemp.RecordsetData, frmTemp.FileName, "��·����ͳ�Ƽ�", frmTemp.CustomData
        
    End If
End Sub

Private Sub mnu_RouteTurnOver_Click()
    Dim lHelpContextID As Long
    lHelpContextID = frmRouteTurnOver.HelpContextID
    frmRouteTurnOver.Show vbModal, Me
    If frmRouteTurnOver.m_bOk Then
        Dim frmTemp As IConditionForm
        Dim frmNewReport As New frmReport
        Set frmTemp = frmRouteTurnOver
        
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport frmTemp.RecordsetData, frmTemp.FileName, "��·Ӫ��ͳ�Ƽ�", frmTemp.CustomData
        
    End If

End Sub

Private Sub mnu_SellerEveryDay_Click()
    On Error GoTo Error_Handle
    Dim lHelpContextID As Long
    
    lHelpContextID = frmSellerEveryDayCon.HelpContextID
    
    frmSellerEveryDayCon.m_bCheck = False

    frmSellerEveryDayCon.Show vbModal, Me
    If frmSellerEveryDayCon.m_bOk Then
        
        Dim rsSellDetail As Recordset
        Dim rsDetailToShow As Recordset
        Dim adbOther() As Double
        Dim oDss As New TicketSellerDim
        Dim i As Integer, nUserCount As Integer
        
        Dim szLastTicketID As String
        Dim szBeginTicketID As String
        Dim arsData() As Recordset, vaCostumData As Variant
        
'        Dim lFullnumber As Long, lHalfnumber As Long, lFreenumber As Long
'        Dim dbFullAmount As Double, dbHalfAmount As Double, dbFreeAmount As Double
        Dim alNumber(TP_TicketTypeCount) As Long '����Ʊ�ֵ�����
        Dim adbAmount(TP_TicketTypeCount) As Double  '����Ʊ�ֵĽ��
        Dim j As Integer
        Dim aszAllSeller() As String
        Dim nAllSeller As Integer
        Dim k As Integer
        Dim l As Integer
        
        oDss.Init m_oActiveUser
        
        aszAllSeller = oDss.GetOperator(frmSellerEveryDayCon.m_dtWorkDate, frmSellerEveryDayCon.m_dtEndDate, ResolveDisplay(frmSellerEveryDayCon.cboSellStation))
        nAllSeller = ArrayLength(aszAllSeller)
        
        
        nUserCount = ArrayLength(frmSellerEveryDayCon.m_vaSeller)
        
        If nAllSeller > 0 Then
            
            ReDim arsData(1 To IIf(nAllSeller > nUserCount, nUserCount, nAllSeller))
            ReDim vaCostumData(1 To IIf(nAllSeller > nUserCount, nUserCount, nAllSeller), 1 To 11, 1 To 2)
            WriteProcessBar True, , nUserCount, "�����γɼ�¼��..."
            l = 0
            For i = 1 To nUserCount
                WriteProcessBar , i, nUserCount, "���ڵõ�" & EncodeString(frmSellerEveryDayCon.m_vaSeller(i)) & "������..."
                For k = 1 To nAllSeller
                    If LCase(Trim(ResolveDisplay(frmSellerEveryDayCon.m_vaSeller(i)))) = LCase(aszAllSeller(k)) Then
                        Exit For
                    End If
                Next k
                If k <= nAllSeller Then
                    l = l + 1
                    '��ʼ��
                    For j = 1 To TP_TicketTypeCount
                        alNumber(j) = 0
                        adbAmount(j) = 0
                    Next j
                
                    Set rsSellDetail = oDss.SellerEveryDaySellDetail(ResolveDisplay(frmSellerEveryDayCon.m_vaSeller(i)), frmSellerEveryDayCon.m_dtWorkDate, frmSellerEveryDayCon.m_dtEndDate)
                    Set rsDetailToShow = New Recordset
                    With rsDetailToShow.Fields
                        .Append "ticket_id_range", adChar, 30
                        '����¼�������ÿ��Ʊ�ֵ����������ֶ�
                        For j = 1 To TP_TicketTypeCount
                            .Append "number_ticket_type" & j, adInteger
                            .Append "amount_ticket_type" & j, adCurrency
                        Next j
                    End With
     
                    rsDetailToShow.Open
                    Dim nTicketNumberLen As Integer
                    Dim nTicketPrefixLen As Integer
                    nTicketNumberLen = m_oParam.TicketNumberLen
                    nTicketPrefixLen = m_oParam.TicketPrefixLen
                    
                    Do While Not rsSellDetail.EOF
                        If rsDetailToShow.RecordCount = 0 Or Not IsTicketIDSequence(szLastTicketID, RTrim(rsSellDetail!ticket_id), nTicketNumberLen, nTicketPrefixLen) Then
                            If rsDetailToShow.RecordCount <> 0 Then
                                rsDetailToShow!ticket_id_range = szBeginTicketID & "---" & szLastTicketID
                                
                                For j = 1 To TP_TicketTypeCount
                                    alNumber(j) = alNumber(j) + rsDetailToShow("number_ticket_type" & j)
                                    adbAmount(j) = adbAmount(j) + rsDetailToShow("amount_ticket_type" & j)
                                Next j
                            End If
    
                            szBeginTicketID = RTrim(rsSellDetail!ticket_id)
                            rsDetailToShow.AddNew
                        End If
                        rsDetailToShow("number_ticket_type" & rsSellDetail!ticket_type) = rsDetailToShow("number_ticket_type" & rsSellDetail!ticket_type) + 1
                        rsDetailToShow("amount_ticket_type" & rsSellDetail!ticket_type) = rsDetailToShow("amount_ticket_type" & rsSellDetail!ticket_type) + rsSellDetail!ticket_price
                        
                        szLastTicketID = RTrim(rsSellDetail!ticket_id)
                        
                        rsSellDetail.MoveNext
                    Loop
                    
                    If rsSellDetail.RecordCount > 0 Then
                        rsSellDetail.MoveLast
                        rsDetailToShow!ticket_id_range = szBeginTicketID & "---" & RTrim(rsSellDetail!ticket_id)
                        For j = 1 To TP_TicketTypeCount
                            alNumber(j) = alNumber(j) + rsDetailToShow("number_ticket_type" & j)
                            adbAmount(j) = adbAmount(j) + rsDetailToShow("amount_ticket_type" & j)
                        Next j
    
                        rsDetailToShow.AddNew
                        
                        rsDetailToShow!ticket_id_range = "�ϼ�"
                        For j = 1 To TP_TicketTypeCount
                            rsDetailToShow("number_ticket_type" & j) = alNumber(j)
                            rsDetailToShow("amount_ticket_type" & j) = adbAmount(j)
                        Next j
                        rsDetailToShow.Update
                    End If
                    Set arsData(l) = rsDetailToShow
                    adbOther = oDss.SellerEveryDayAnotherThing(ResolveDisplay(frmSellerEveryDayCon.m_vaSeller(i)), frmSellerEveryDayCon.m_dtWorkDate, frmSellerEveryDayCon.m_dtEndDate)
                    vaCostumData(l, 1, 1) = "��Ʊ"
                    vaCostumData(l, 1, 2) = "����=" & CInt(adbOther(1, 1)) & " ��  Ʊ��=" & adbOther(1, 2) & " Ԫ"
                    
                    vaCostumData(l, 2, 1) = "��Ʊ"
                    vaCostumData(l, 2, 2) = "����=" & CInt(adbOther(2, 1)) & " ��  Ʊ��=" & adbOther(2, 2) & " Ԫ  ������=" & adbOther(2, 3) & " Ԫ"
                    
                    
                    vaCostumData(l, 3, 1) = "ȫ����Ʊ"
                    vaCostumData(l, 3, 2) = "����=" & CInt(adbOther(4, 1)) & " ��  Ʊ��=" & adbOther(4, 2) & " Ԫ" '  ȫ����Ʊ������=" & adbOther(4, 3) & " Ԫ"
                    
                    
                    vaCostumData(l, 4, 1) = "��ǩ"
                    vaCostumData(l, 4, 2) = "����=" & CInt(adbOther(3, 1)) & " ��  Ʊ��=" & adbOther(3, 2) & " Ԫ  ������=" & adbOther(3, 3) & " Ԫ"
                    
                    Dim dbAmount As Double '��������Ʊ
                    Dim lNumber As Long '������Ʊ
                    lNumber = 0
                    dbAmount = 0
                    For j = 1 To TP_TicketTypeCount
                        If j <> TP_FreeTicket Then
                            dbAmount = dbAmount + adbAmount(j)
                        End If
                        lNumber = lNumber + alNumber(j)
                    Next j
                        
                    vaCostumData(l, 5, 1) = "Ӧ����"
                    vaCostumData(l, 5, 2) = dbAmount - adbOther(1, 2) - adbOther(2, 2) + adbOther(2, 3) - adbOther(4, 2) + adbOther(4, 3) - adbOther(3, 2) + adbOther(3, 3) & " Ԫ"
                    
                    vaCostumData(l, 6, 1) = "��Ʊ��"
                    vaCostumData(l, 6, 2) = lNumber & " ��"
                    
                    vaCostumData(l, 7, 1) = "��Ʊ��Ʊ"
                    vaCostumData(l, 7, 2) = lNumber - adbOther(1, 1) - adbOther(2, 1) - adbOther(4, 1) - adbOther(3, 1) & " ��"
                    
                    vaCostumData(l, 8, 1) = "�Ƶ�"
                    vaCostumData(l, 8, 2) = MakeDisplayString(m_oActiveUser.UserID, m_oActiveUser.UserName)
                    
                    vaCostumData(l, 9, 1) = "����"
                    vaCostumData(l, 9, 2) = ""
                    
                    vaCostumData(l, 10, 1) = "��ƱԱ"
                    vaCostumData(l, 10, 2) = frmSellerEveryDayCon.m_vaSeller(i)
                    
                    vaCostumData(l, 11, 1) = "��������"
                    vaCostumData(l, 11, 2) = Format(frmSellerEveryDayCon.m_dtWorkDate, "MM��DD�� hh:mm") & "��" & Format(frmSellerEveryDayCon.m_dtEndDate, "MM��DD�� hh:mm")
                    
                End If
            Next
            WriteProcessBar False, , , ""
            
            Dim frmNewReport As New frmReport
'            frmNewReport.Show
            Dim frmTemp As IConditionForm
            Set frmTemp = frmSellerEveryDayCon
            frmNewReport.m_lHelpContextID = lHelpContextID
            frmNewReport.ShowReport2 arsData, frmTemp.FileName, cszSellerEveryDay, vaCostumData, 10
            
            WriteProcessBar False, , , ""
        End If
    End If
    Exit Sub
Error_Handle:
    WriteProcessBar False, , , ""
    ShowErrorMsg
End Sub

'��վ��ƱӪ�ռ�
Private Sub mnu_UnitSalePriceSimplle_Click()
    Dim lHelpContextID As Long
    lHelpContextID = frmUnitSellSimple.HelpContextID
    
    frmUnitSellSimple.Show vbModal, Me
    If frmUnitSellSimple.m_bOk Then
        Dim frmTemp As IConditionForm
        Dim frmNewReport As New frmReport
        Set frmTemp = frmUnitSellSimple
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport frmTemp.RecordsetData, frmTemp.FileName, cszUnitInterTotal, frmTemp.CustomData
    End If
    
End Sub


'��ƱԱ��Ʊͳ���±�
Private Function mnu_CheckerEveryMonth_Click() As Recordset
    On Error GoTo Error_Handle
    Dim lHelpContextID As Long
    
    lHelpContextID = frmCheckerEveryMonth.HelpContextID

    frmCheckerEveryMonth.Show vbModal, Me
    If frmCheckerEveryMonth.m_bOk Then
        Dim rsChecked As Recordset
        Dim oCTReport As New STChkTk.CTReport
        Dim i As Integer, nUserCount As Integer
        Dim vaCostumData As Variant
        Dim aszAllSeller() As String
        
        oCTReport.Init m_oActiveUser
        
        nUserCount = ArrayLength(frmCheckerEveryMonth.m_vaSeller)
        ReDim arsData(1 To nUserCount)
        ReDim vaCostumData(1 To 3, 1 To 2)
        '�˴������޸�,��ʱ��һ��
        ReDim aszAllSeller(1 To nUserCount, 1 To 1)
        For i = 1 To nUserCount
            aszAllSeller(i, 1) = ResolveDisplay(frmCheckerEveryMonth.m_vaSeller(i))
        Next i
        
        Set rsChecked = oCTReport.GetCheckerEveryMonth(aszAllSeller, frmCheckerEveryMonth.m_dtWorkDate, frmCheckerEveryMonth.m_dtEndDate)
        
        If rsChecked.RecordCount <> 0 Then

            vaCostumData(1, 1) = "�Ʊ���"
            vaCostumData(1, 2) = MakeDisplayString(m_oActiveUser.UserID, m_oActiveUser.UserName)
            
            vaCostumData(2, 1) = "��ƱԱ"
            vaCostumData(2, 2) = "" ' frmCheckerEveryMonth.m_vaSeller(i)
            
            vaCostumData(3, 1) = "ͳ���·�"
            vaCostumData(3, 2) = Format(frmCheckerEveryMonth.m_dtWorkDate) & "��" & Format(frmCheckerEveryMonth.m_dtEndDate)
 
        End If
        Set mnu_CheckerEveryMonth_Click = rsChecked

        WriteProcessBar False, , , ""
        
        Dim frmNewReport As New frmReport
        Dim frmTemp As IConditionForm
        Set frmTemp = frmCheckerEveryMonth
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport mnu_CheckerEveryMonth_Click, frmTemp.FileName, cszCheckerEveryMonth, vaCostumData, 10
        Unload frmCheckerEveryMonth
        WriteProcessBar False, , , ""
    End If
    Exit Function
Error_Handle:
    WriteProcessBar False, , , ""
    ShowErrorMsg
End Function

Private Sub mnu_SellerEveryMonth_Click()
    Dim lHelpContextID As Long
    lHelpContextID = frmSellerEveryMonth.HelpContextID
    frmSellerEveryMonth.Show vbModal, Me
    If frmSellerEveryMonth.m_bOk Then
        Dim frmTemp As IConditionForm
        Dim frmNewReport As New frmReport
        Set frmTemp = frmSellerEveryMonth
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport frmTemp.RecordsetData, frmTemp.FileName, "��ƱԱ�����", frmTemp.CustomData
    End If
End Sub

Private Sub mnu_SellerInterSimple_Click()
    Dim lHelpContextID As Long
    lHelpContextID = frmSellerInterSimpleCon.HelpContextID
    frmSellerInterSimpleCon.Show vbModal, Me
    If frmSellerInterSimpleCon.m_bOk Then
        Dim frmTemp As IConditionForm
        Dim frmNewReport As New frmReport
        Set frmTemp = frmSellerInterSimpleCon
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport frmTemp.RecordsetData, frmTemp.FileName, cszSellerInterSimple, frmTemp.CustomData
    End If
    
End Sub

Private Sub mnu_SellerSimpleCon_Click()
    Dim lHelpContextID As Long
    lHelpContextID = frmSellerSimpleCon.HelpContextID

 frmSellerSimpleCon.Show vbModal, Me
   
    
    If frmSellerSimpleCon.m_bOk Then
        Dim frmTemp As IConditionForm
        Dim frmNewReport As New frmReport
        Set frmTemp = frmSellerSimpleCon
      
       
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport frmTemp.RecordsetData, frmTemp.FileName, cszSellerSimpleMonth, frmTemp.CustomData
       
    End If
End Sub

Private Sub mnu_Sationsell_Click()
 Dim lHelpContextID As Long
   
lHelpContextID = frmstationsell.HelpContextID
 
frmstationsell.Show vbModal, Me
    
    If frmstationsell.m_bOk Then
        Dim frmTemp As IConditionForm
        Dim frmNewReport As New frmReport
   
       Set frmTemp = frmstationsell
       
        frmNewReport.m_lHelpContextID = lHelpContextID
    
        frmNewReport.ShowReport frmTemp.RecordsetData, frmTemp.FileName, cszSellerSimpleMonth, frmTemp.CustomData
    End If
End Sub




Private Sub mnu_SellerDayCon_Click()
    Dim lHelpContextID As Long
    lHelpContextID = frmSellerDayCon.HelpContextID

    frmSellerDayCon.Show vbModal, Me
    If frmSellerDayCon.m_bOk Then
        Dim frmTemp As IConditionForm
        Dim frmNewReport As New frmReport
        Set frmTemp = frmSellerDayCon
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport frmTemp.RecordsetData, frmTemp.FileName, cszSellerSimple, frmTemp.CustomData
    End If
End Sub

Private Sub mnu_SellerPriceItemCon_Click()
    Dim lHelpContextID As Long
    lHelpContextID = frmSellerPriceItemCon.HelpContextID
    frmSellerPriceItemCon.Show vbModal, Me
    If frmSellerPriceItemCon.m_bOk Then
        Dim frmTemp As IConditionForm
        Dim frmNewReport As New frmReport
        Set frmTemp = frmSellerPriceItemCon
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport frmTemp.RecordsetData, frmTemp.FileName, cszSellerPriceItemCon, frmTemp.CustomData
    End If
    
End Sub


'Private Sub mnu_SellerSimpleDay_Click()
'Dim lHelpContextID As Long
'    lHelpContextID = frmSellerSimpleCon.HelpContextID
'
'    frmSellerSimpleCon.Show vbModal, Me
'    If frmSellerSimpleCon.m_bOk Then
'        Dim frmTemp As IConditionForm
'        Dim frmNewReport As New frmReport
'        Set frmTemp = frmSellerSimpleCon
'        frmNewReport.m_lHelpContextID = lHelpContextID
'        frmNewReport.ShowReport frmTemp.RecordsetData, frmTemp.FileName, cszSellerSimple, frmTemp.CustomData
'    End If
'End Sub


Private Sub mnu_SomeSumByBusDate_Click()
    Dim lHelpContextID As Long
    frmBusSomeSum.m_bBySaleTime = False
    lHelpContextID = frmBusSomeSum.HelpContextID
    
    frmBusSomeSum.Show vbModal, Me
    If frmBusSomeSum.m_bOk Then
        Dim frmTemp As IConditionForm
        Dim frmNewReport As New frmReport
        Set frmTemp = frmBusSomeSum
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport frmTemp.RecordsetData, frmTemp.FileName, "��������������С��" & "(���������ڻ���)", frmTemp.CustomData, 2
        
    End If
End Sub

Private Sub mnu_SomeSumBySellTime_Click()

    Dim lHelpContextID As Long
    frmBusSomeSum.m_bBySaleTime = True
    lHelpContextID = frmBusSomeSum.HelpContextID
    
    frmBusSomeSum.Show vbModal, Me
    If frmBusSomeSum.m_bOk Then
        Dim frmTemp As IConditionForm
        Dim frmNewReport As New frmReport
        Set frmTemp = frmBusSomeSum
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport frmTemp.RecordsetData, frmTemp.FileName, "��������������С��" & "(����Ʊ���ڻ���)", frmTemp.CustomData, 2
        
    End If
End Sub

Private Sub mnu_StationDaily_Click()
    Dim lHelpContextID As Long
    lHelpContextID = FrmStationDaily.HelpContextID
    FrmStationDaily.Show vbModal, Me
    If FrmStationDaily.m_bOk Then
        Dim frmTemp As IConditionForm
        Dim frmNewReport As New frmReport
        Set frmTemp = FrmStationDaily
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport frmTemp.RecordsetData, frmTemp.FileName, "��վվ���ձ�", frmTemp.CustomData
    End If
End Sub

Private Sub mnu_StationMonthly_Click()
    Dim lHelpContextID As Long
    lHelpContextID = frmStationCon.HelpContextID
    
    frmStationCon.Show vbModal, Me
    If frmStationCon.m_bOk Then
        Dim frmTemp As IConditionForm
        Dim frmNewReport As New frmReport
        Set frmTemp = frmStationCon
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport frmTemp.RecordsetData, frmTemp.FileName, cszStationMonthly, frmTemp.CustomData, 2
    End If
End Sub

Private Sub mnu_StationSellTicket_Click()
    Dim oTicketAcc As New TicketBusDim
    Dim rsStationSell As Recordset
    Dim szStationId() As String
    Dim frmNewReport As frmReport
    Dim vaCostumData As Variant
    
    On Error GoTo errorHander

    FrmStationSellCount.Show vbModal
    If FrmStationSellCount.m_bOk Then
        ReDim vaCostumData(1 To 3, 1 To 2)
        vaCostumData(1, 1) = "��ʼʱ��"
        vaCostumData(1, 2) = Format(FrmStationSellCount.m_dtStartDate, "YYYY-MM-DD")
        vaCostumData(2, 1) = "����ʱ��"
        vaCostumData(2, 2) = Format(FrmStationSellCount.m_dtEndDate, "YYYY-MM-DD")
            
        vaCostumData(3, 1) = "�Ʊ���"
        vaCostumData(3, 2) = m_oActiveUser.UserID
        Me.MousePointer = vbHourglass
        oTicketAcc.Init m_oActiveUser
        szStationId = FrmStationSellCount.m_szStationID
        Set frmNewReport = New frmReport
        If FrmStationSellCount.m_bnList = True Then
            Set rsStationSell = oTicketAcc.GetStationTickets(FrmStationSellCount.m_dtStartDate, FrmStationSellCount.m_dtEndDate, szStationId, FrmStationSellCount.m_nOrder)
            frmNewReport.ShowReport rsStationSell, "����վ����Ʊ��ϸ��ѯ.xls", cszBusSationSellDetail, vaCostumData, 10
        Else
            If FrmStationSellCount.m_nCount = 0 Then
               Set rsStationSell = oTicketAcc.GetStationTicketsCount(FrmStationSellCount.m_dtStartDate, FrmStationSellCount.m_dtEndDate, szStationId, FrmStationSellCount.m_nCount)
                frmNewReport.ShowReport rsStationSell, "������Ʊͳ��.xls", cszBusSellCount, vaCostumData, 10
            ElseIf FrmStationSellCount.m_nCount = 1 Then
               Set rsStationSell = oTicketAcc.GetStationTicketsCount(FrmStationSellCount.m_dtStartDate, FrmStationSellCount.m_dtEndDate, szStationId, FrmStationSellCount.m_nCount)
                frmNewReport.ShowReport rsStationSell, "վ����Ʊͳ��.xls", cszSationSellCount, vaCostumData, 10
            Else
                Set rsStationSell = oTicketAcc.GetStationTicketsCount(FrmStationSellCount.m_dtStartDate, FrmStationSellCount.m_dtEndDate, szStationId, FrmStationSellCount.m_nCount)
                frmNewReport.ShowReport rsStationSell, "����վ����Ʊͳ��.xls", cszBusSationSellCount, vaCostumData, 10
            End If
        End If
        Set rsStationSell = Nothing
        Set frmNewReport = Nothing
        Me.MousePointer = vbDefault
    End If
    Exit Sub
    
errorHander:
    Me.MousePointer = vbDefault
    MsgBox err.Description, vbCritical + vbOKOnly, "����"
End Sub



Private Sub mnu_TitleH_Click()
    Arrange vbTileHorizontal
End Sub

Private Sub mnu_TitleV_Click()
    Arrange vbTileVertical
End Sub

Private Sub mnu_UnitDaily_Click()
    Dim lHelpContextID As Long
    lHelpContextID = frmUnitDailyCon.HelpContextID
    
    frmUnitDailyCon.Show vbModal, Me
    If frmUnitDailyCon.m_bOk Then
        Dim frmTemp As IConditionForm
        Dim frmNewReport As New frmReport
        Set frmTemp = frmUnitDailyCon
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport frmTemp.RecordsetData, frmTemp.FileName, cszUnitDaily, frmTemp.CustomData, 1
    End If
    
End Sub

Private Sub mnu_UnitInterSell_Click()
    Dim lHelpContextID As Long
    lHelpContextID = frmUnitInterSellCon.HelpContextID
    
    frmUnitInterSellCon.Show vbModal, Me
    If frmUnitInterSellCon.m_bOk Then
        Dim frmTemp As IConditionForm
        Dim frmNewReport As New frmReport
        Set frmTemp = frmUnitInterSellCon
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport frmTemp.RecordsetData, frmTemp.FileName, cszUnitInterSell, frmTemp.CustomData
    End If
    
End Sub

Private Sub mnu_UnitInterSettle_Click()
    Dim lHelpContextID As Long
    lHelpContextID = frmUnitInterSettle.HelpContextID
    
    frmUnitInterSettle.Show vbModal, Me
    If frmUnitInterSettle.m_bOk Then
        Dim frmTemp As IConditionForm
        Dim frmNewReport As New frmReport
        Set frmTemp = frmUnitInterSettle
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport frmTemp.RecordsetData, frmTemp.FileName, cszUnitInterSettle, frmTemp.CustomData
    End If
    
End Sub

Private Sub mnu_UnitInterSimple_Click()
    Dim lHelpContextID As Long
    lHelpContextID = frmUnitInterSimpleCon.HelpContextID
    
    frmUnitInterSimpleCon.Show vbModal, Me
    If frmUnitInterSimpleCon.m_bOk Then
        Dim frmTemp As IConditionForm
        Dim frmNewReport As New frmReport
        Set frmTemp = frmUnitInterSimpleCon
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport frmTemp.RecordsetData, frmTemp.FileName, cszUnitInterSimple, frmTemp.CustomData
    End If
    
End Sub

Private Sub mnu_UnitInterTotal_Click()
    Dim lHelpContextID As Long
    lHelpContextID = frmUnitInterTotalCon.HelpContextID
    
    frmUnitInterTotalCon.Show vbModal, Me
    If frmUnitInterTotalCon.m_bOk Then
        Dim frmTemp As IConditionForm
        Dim frmNewReport As New frmReport
        Set frmTemp = frmUnitInterTotalCon
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport frmTemp.RecordsetData, frmTemp.FileName, cszUnitInterTotal, frmTemp.CustomData
    End If
    
End Sub

Private Sub mnu_UnitManSimple_Click()
    Dim lHelpContextID As Long
    lHelpContextID = frmUnitManSimple.HelpContextID
    
    frmUnitManSimple.Show vbModal, Me
    If frmUnitManSimple.m_bOk Then
        Dim frmTemp As IConditionForm
        Dim frmNewReport As New frmReport
        Set frmTemp = frmUnitManSimple
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport frmTemp.RecordsetData, frmTemp.FileName, cszUnitManSimple, frmTemp.CustomData, 2
    End If

End Sub

Private Sub mnu_UnitMonthly_Click()
    Dim lHelpContextID As Long
    lHelpContextID = frmUnitMonthlyCon.HelpContextID
    
    frmUnitMonthlyCon.Show vbModal, Me
    If frmUnitMonthlyCon.m_bOk Then
        Dim frmTemp As IConditionForm
        Dim frmNewReport As New frmReport
        Set frmTemp = frmUnitMonthlyCon
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport frmTemp.RecordsetData, frmTemp.FileName, cszUnitMonthly, frmTemp.CustomData, 2
    End If

End Sub

Private Sub mnu_UnitProxySimple_Click()
    Dim lHelpContextID As Long
    lHelpContextID = frmUnitProxySimpleCon.HelpContextID
    
    frmUnitProxySimpleCon.Show vbModal, Me
    If frmUnitProxySimpleCon.m_bOk Then
        Dim frmTemp As IConditionForm
        Dim frmNewReport As New frmReport
        Set frmTemp = frmUnitProxySimpleCon
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport frmTemp.RecordsetData, frmTemp.FileName, cszUnitProxySimple, frmTemp.CustomData
    End If
End Sub

Private Sub mnu_UnitYearly_Click()
Dim lHelpContextID As Long
    lHelpContextID = frmUnitYearlyCon.HelpContextID
    
    frmUnitYearlyCon.Show vbModal, Me
    If frmUnitYearlyCon.m_bOk Then
        Dim frmTemp As IConditionForm
        Dim frmNewReport As New frmReport
        Set frmTemp = frmUnitYearlyCon
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport frmTemp.RecordsetData, frmTemp.FileName, cszUnitYearly, frmTemp.CustomData, 2
    End If
End Sub

Private Sub mnu_UserProperty_Click()
    On Error GoTo ErrorHandle
    m_oShell.Init m_oActiveUser
    m_oShell.ShowUserInfo
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub mnuSellerTicketDetail_Click()
    
    Dim oTicketAcc As New TicketSellerDim
    Dim rsSellDetail As Recordset

    Dim i As Integer, nUserCount As Integer
    Dim vaCostumData As Variant
    Dim szFileName1 As String
    
    On Error GoTo errorHander
    
    frmSellerTicketDetail.m_nStatus2 = 1
    frmSellerTicketDetail.Caption = "����λ��Ʊ��ϸ��ѯ"
    frmSellerTicketDetail.Show vbModal, Me
    If frmSellerTicketDetail.m_bOk Then
        ReDim vaCostumData(1 To 3, 1 To 2)
'        vaCostumData(1, 1) = "��ʼʱ��"
'        vaCostumData(1, 2) = Format(frmSellerTicketDetail.m_dtStartDate, "YYYY-MM-DD")
'        vaCostumData(2, 1) = "����ʱ��"
'        vaCostumData(2, 2) = Format(frmSellerTicketDetail.m_dtEndDate, "YYYY-MM-DD")
'
        vaCostumData(1, 1) = "�Ʊ���"
        vaCostumData(1, 2) = m_oActiveUser.UserID
        
        Me.MousePointer = vbHourglass
        oTicketAcc.Init m_oActiveUser
        
        nUserCount = ArrayLength(frmSellerTicketDetail.m_vaSeller)
        'If nUserCount > 0 Then
            'Set rsSellDetail = oTicketAcc.SellerTicketDetail(frmSellerTicketDetail.m_vaSeller, frmSellerTicketDetail.m_dtBeginDateTime, frmSellerTicketDetail.m_dtEndDateTime, frmSellerTicketDetail.m_nStatus)
           
            Set rsSellDetail = oTicketAcc.SellerTicketDetail(frmSellerTicketDetail.m_vaSeller, frmSellerTicketDetail.m_dtBeginDateTime, frmSellerTicketDetail.m_dtEndDateTime, frmSellerTicketDetail.m_szFromTicketNo, frmSellerTicketDetail.m_szToTicketNo, frmSellerTicketDetail.m_nStatus, frmSellerTicketDetail.m_szBusId, frmSellerTicketDetail.m_szIDCardNo, frmSellerTicketDetail.m_szPersonName)
            
            Dim frmNewReport As New frmReport
            
            Dim frmTemp As IConditionForm
            
            Set frmTemp = frmSellerTicketDetail
            
            'frmNewReport.m_lHelpContextID = m_lHelpContextID
            If frmSellerTicketDetail.m_nStatus <= ST_TicketNormal Then
               szFileName1 = cszSellTicketDetail
            ElseIf frmSellerTicketDetail.m_nStatus = ST_TicketSellChange Or frmSellerTicketDetail.m_nStatus = ST_TicketChanged Then
               szFileName1 = cszChangeTicketDetail
            ElseIf frmSellerTicketDetail.m_nStatus = ST_TicketCanceled Then
               szFileName1 = cszCancelTicketDetail
            ElseIf frmSellerTicketDetail.m_nStatus = ST_TicketReturned Then
               szFileName1 = cszReturnTicketDetail
            ElseIf frmSellerTicketDetail.m_nStatus = ST_TicketChecked Then
               szFileName1 = cszCheckTicketDetail
            End If
            
            'frmNewReport.ShowReport rsSellDetail, frmTemp.FileName, cszSellTicketDetail, vaCostumData, 10
            frmNewReport.ShowReport rsSellDetail, szFileName1 & ".xls", szFileName1, vaCostumData, 10
            
        'End If
    End If
    Me.MousePointer = vbDefault
    Exit Sub
    
errorHander:
    Me.MousePointer = vbDefault
    MsgBox err.Description, vbCritical + vbOKOnly, "����"
End Sub



Private Sub mnuSellerTicketDetailFromAgent_Click()
    
    Dim oTicketAcc As New TicketSellerDim
    Dim rsSellDetail As Recordset

    Dim i As Integer, nUserCount As Integer
    Dim vaCostumData As Variant
    Dim szFileName1 As String
    
    On Error GoTo errorHander
    
    frmSellerTicketDetail.m_nStatus2 = 0
    frmSellerTicketDetail.Caption = "��ƱԱ��Ʊ��ϸ��ѯ"
    frmSellerTicketDetail.Show vbModal, Me
    If frmSellerTicketDetail.m_bOk Then
        Me.MousePointer = vbHourglass
        oTicketAcc.Init m_oActiveUser
        
        nUserCount = ArrayLength(frmSellerTicketDetail.m_vaSeller)
        'If nUserCount > 0 Then
            'Set rsSellDetail = oTicketAcc.SellerTicketDetail(frmSellerTicketDetail.m_vaSeller, frmSellerTicketDetail.m_dtBeginDateTime, frmSellerTicketDetail.m_dtEndDateTime, frmSellerTicketDetail.m_nStatus)
           
            Set rsSellDetail = oTicketAcc.SellerTicketDetailFromAgent(frmSellerTicketDetail.m_vaSeller, frmSellerTicketDetail.m_dtBeginDateTime, frmSellerTicketDetail.m_dtEndDateTime, frmSellerTicketDetail.m_szFromTicketNo, frmSellerTicketDetail.m_szToTicketNo, frmSellerTicketDetail.m_nStatus, frmSellerTicketDetail.m_szBusId, frmSellerTicketDetail.m_szIDCardNo, frmSellerTicketDetail.m_szPersonName)
            
            Dim frmNewReport As New frmReport
            
            Dim frmTemp As IConditionForm
            
            Set frmTemp = frmSellerTicketDetail
            
            'frmNewReport.m_lHelpContextID = m_lHelpContextID
            If frmSellerTicketDetail.m_nStatus <= ST_TicketNormal Then
               szFileName1 = cszSellTicketDetail
            ElseIf frmSellerTicketDetail.m_nStatus = ST_TicketSellChange Or frmSellerTicketDetail.m_nStatus = ST_TicketChanged Then
               szFileName1 = cszChangeTicketDetail
            ElseIf frmSellerTicketDetail.m_nStatus = ST_TicketCanceled Then
               szFileName1 = cszCancelTicketDetail
            ElseIf frmSellerTicketDetail.m_nStatus = ST_TicketReturned Then
               szFileName1 = cszReturnTicketDetail
            ElseIf frmSellerTicketDetail.m_nStatus = ST_TicketChecked Then
               szFileName1 = cszCheckTicketDetail
            End If
            
            'frmNewReport.ShowReport rsSellDetail, frmTemp.FileName, cszSellTicketDetail, vaCostumData, 10
            frmNewReport.ShowReport rsSellDetail, szFileName1 & ".xls", szFileName1, vaCostumData, 10
            
        'End If
    End If
    Me.MousePointer = vbDefault
    Exit Sub
    
errorHander:
    Me.MousePointer = vbDefault
    MsgBox err.Description, vbCritical + vbOKOnly, "����"
End Sub



Private Sub MDIForm_Resize()
    On Error Resume Next
    cmdClose.Left = Me.Width - cmdClose.Width - 800

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    End
End Sub

Private Sub pmnu_AddCombine_Click()
    Me.ActiveForm.AddCombine
    
End Sub

Private Sub pmnu_AddPreSaler_Click()
    frmSellerEveryMonth.AddPreSaler
End Sub

Private Sub pmnu_AddSinceSaler_Click()
    frmSellerEveryMonth.AddSinceSaler
End Sub

Private Sub pmnu_DeleteCombine_Click()
    Me.ActiveForm.DeleteCombine
    
End Sub

Private Sub pmnu_EditCombine_Click()
    Me.ActiveForm.ModifyCombine
End Sub

Private Sub pmnu_RemoveAllSaler_Click()
    frmSellerEveryMonth.RemovePreAll
End Sub

Private Sub pmnu_RemoveAllSaler2_Click()
    frmSellerEveryMonth.RemoveSinceAll
    
End Sub

Private Sub pmnu_RemoveSelSaler_Click()
    frmSellerEveryMonth.RemovePreSaler
End Sub

Private Sub pmnu_RemoveSelSaler2_Click()
    frmSellerEveryMonth.RemoveSinceSaler
End Sub

Private Sub pmnu_SelectAll_Click()
    frmBusTransStat.SelectAllBus
End Sub

Private Sub pmnu_TicketIssue_Click()
    Dim lHelpContextID As Long
    lHelpContextID = frmTicketIssueScheme.HelpContextID
    frmTicketIssueScheme.Show vbModal, Me
    If frmTicketIssueScheme.m_bOk Then
        Dim frmTemp As IConditionForm
        Dim frmNewReport As New frmReport
        Set frmTemp = frmTicketIssueScheme
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport frmTemp.RecordsetData, frmTemp.FileName, "��Ʊ���ۼƻ�", frmTemp.CustomData
    End If
End Sub

Private Function IsTicketIDSequence(pszFirstTicketID As String, pszSecondTicketID As String, nTicketNumberLen As Integer, nTicketPrefixLen As Integer) As Boolean
    Dim szTemp1 As String, szTemp2 As String
    On Error GoTo Error_Handle
    szTemp1 = UCase(Left(pszFirstTicketID, nTicketPrefixLen))
    szTemp2 = UCase(Left(pszSecondTicketID, nTicketPrefixLen))
    If szTemp1 = szTemp2 Then
        szTemp1 = Right(pszFirstTicketID, nTicketNumberLen)
        szTemp2 = Right(pszSecondTicketID, nTicketNumberLen)
        If CLng(szTemp1) + 1 = CLng(szTemp2) Then
            IsTicketIDSequence = True
        End If
    End If
    Exit Function
Error_Handle:
End Function

Public Sub SetPrintEnabled(pbEnabled As Boolean)
    '���ò˵��Ŀ�����
    With abMenu
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
'����ActiveBar�Ŀؼ�
Private Sub AddControlsToActBar()
    abMenu.Bands("bndTitleTop").Tools("tblTitleTop").Custom = ptTitleTop
    abMenu.Bands("statusBar").Tools("progressBar").Custom = pbLoad
End Sub


'��ƱԱ·��ͳ��
Private Sub CheckerSheetCon()
    On Error GoTo Error_Handle
'    Dim lHelpContextID As Long
'    lHelpContextID = frmSellerEveryDayCon.HelpContextID
    frmSellerEveryDayCon.m_bCheck = True
    frmSellerEveryDayCon.Show vbModal, Me
    If frmSellerEveryDayCon.m_bOk Then
        
        Dim rsDetail As Recordset
        Dim rsDetailToShow As Recordset
        Dim adbOther() As Double
        Dim i As Integer
        
        Dim szLastTicketID As String
        Dim szBeginTicketID As String
        Dim arsData() As Recordset, vaCostumData As Variant
        
'        Dim lFullnumber As Long, lHalfnumber As Long, lFreenumber As Long
'        Dim dbFullAmount As Double, dbHalfAmount As Double, dbFreeAmount As Double
        Dim lNumber As Long '����Ʊ�ֵ�����
        Dim dbAmount As Double  '����Ʊ�ֵĽ��
        Dim j As Integer
'        Dim k As Integer
'        Dim l As Integer
        
        Dim oCTReport As New CTReport
        Dim nTicketNumberLen As Integer
        Dim nTicketPrefixLen As Integer
        Dim nUserCount As Integer
        
        
        oCTReport.Init m_oActiveUser
        
        
        nUserCount = ArrayLength(frmSellerEveryDayCon.m_vaSeller)
        
        If nUserCount > 0 Then
            
            ReDim arsData(1 To nUserCount)
            ReDim vaCostumData(1 To nUserCount, 1 To 10, 1 To 2)
            WriteProcessBar True, , nUserCount, "�����γɼ�¼��..."
            For i = 1 To nUserCount
                WriteProcessBar , i, nUserCount, "���ڵõ�" & EncodeString(frmSellerEveryDayCon.m_vaSeller(i)) & "������..."

                    lNumber = 0
                    dbAmount = 0

                    Set rsDetail = oCTReport.GetCheckerSheetDetail(ResolveDisplay(frmSellerEveryDayCon.m_vaSeller(i)), frmSellerEveryDayCon.m_dtWorkDate, frmSellerEveryDayCon.m_dtEndDate)
                    Set rsDetailToShow = New Recordset
                    With rsDetailToShow.Fields
                        .Append "ticket_id_range", adChar, 30
                        .Append "number", adInteger
                        .Append "amount", adCurrency
                    End With
     
                    rsDetailToShow.Open
                    nTicketNumberLen = m_oParam.CheckSheetLen
                    nTicketPrefixLen = 0
                    
                    Do While Not rsDetail.EOF
                        If rsDetailToShow.RecordCount = 0 Or Not IsTicketIDSequence(szLastTicketID, RTrim(rsDetail!check_sheet_id), nTicketNumberLen, nTicketPrefixLen) Then
                            If rsDetailToShow.RecordCount <> 0 Then
                                rsDetailToShow!ticket_id_range = szBeginTicketID & "---" & szLastTicketID
                                
                                    lNumber = lNumber + rsDetailToShow("number")
                                    dbAmount = dbAmount + rsDetailToShow("amount")
                            End If
    
                            szBeginTicketID = RTrim(rsDetail!check_sheet_id)
                            rsDetailToShow.AddNew
                        End If
                        rsDetailToShow("number") = rsDetailToShow("number") + 1
'                        rsDetailToShow("amount") = rsDetailToShow("amount") + rsDetail!ticket_price
                        
                        szLastTicketID = RTrim(rsDetail!check_sheet_id)
                        
                        rsDetail.MoveNext
                    Loop
                    
                    If rsDetail.RecordCount > 0 Then
                        rsDetail.MoveLast
                        rsDetailToShow!ticket_id_range = szBeginTicketID & "---" & RTrim(rsDetail!check_sheet_id)
                            lNumber = lNumber + rsDetailToShow("number")
'                            dbAmount = dbAmount + rsDetailToShow("amount")
    
                        rsDetailToShow.AddNew
                        
                        rsDetailToShow!ticket_id_range = "�ϼ�"
                        For j = 1 To TP_TicketTypeCount
                            rsDetailToShow("number") = lNumber
'                            rsDetailToShow("amount") = dbAmount
                        Next j
                        rsDetailToShow.Update
                    End If
                    Set arsData(i) = rsDetailToShow
                    adbOther = oCTReport.GetCheckerSheetAnotherThing(ResolveDisplay(frmSellerEveryDayCon.m_vaSeller(i)), frmSellerEveryDayCon.m_dtWorkDate, frmSellerEveryDayCon.m_dtEndDate)
                    vaCostumData(i, 1, 1) = "����"
                    vaCostumData(i, 1, 2) = CInt(adbOther(1)) & " ��"
                    
                    vaCostumData(i, 2, 1) = "������"
                    vaCostumData(i, 2, 2) = lNumber & " ��"

                    vaCostumData(i, 3, 1) = "��Ʊ��Ʊ"
                    vaCostumData(i, 3, 2) = lNumber - adbOther(1) & " ��"

                    vaCostumData(i, 4, 1) = "�Ƶ�"
                    vaCostumData(i, 4, 2) = MakeDisplayString(m_oActiveUser.UserID, m_oActiveUser.UserName)


                    vaCostumData(i, 5, 1) = "��ƱԱ"
                    vaCostumData(i, 5, 2) = frmSellerEveryDayCon.m_vaSeller(i)

                    vaCostumData(i, 6, 1) = "��������"
                    vaCostumData(i, 6, 2) = Format(frmSellerEveryDayCon.m_dtWorkDate, "MM��DD�� hh:mm") & "��" & Format(frmSellerEveryDayCon.m_dtEndDate, "MM��DD�� hh:mm")
'
'                End If
            Next
            WriteProcessBar False, , , ""
            
            Dim frmNewReport As New frmReport
'            frmNewReport.Show
            Dim frmTemp As IConditionForm
            Set frmTemp = frmSellerEveryDayCon
'            frmNewReport.m_lHelpContextID = lHelpContextID
            frmNewReport.ShowReport2 arsData, frmTemp.FileName, cszSellerEveryDay, vaCostumData, 10
            
            WriteProcessBar False, , , ""
        End If
    End If
    Exit Sub
Error_Handle:
    WriteProcessBar False, , , ""
    ShowErrorMsg

    
End Sub




Private Sub mnu_BusStation_Click()
    Dim lHelpContextID As Long
    lHelpContextID = frmBusStationCon.HelpContextID
    
    frmBusStationCon.Show vbModal, Me
    If frmBusStationCon.m_bOk Then
        Dim frmTemp As IConditionForm
        Dim frmNewReport As New frmReport
        Set frmTemp = frmBusStationCon
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport frmTemp.RecordsetData, frmTemp.FileName, cszBusStationByBusDate, frmTemp.CustomData
    End If
End Sub


Private Sub mnu_BusStationBySellerStation_Click()

    Dim lHelpContextID As Long
    frmBusStationCon.m_bBySeller = True
    lHelpContextID = frmBusStationCon.HelpContextID
    frmBusStationCon.Show vbModal, Me
    If frmBusStationCon.m_bOk Then
        Dim frmTemp As IConditionForm
        Dim frmNewReport As New frmReport
        Set frmTemp = frmBusStationCon
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport frmTemp.RecordsetData, frmTemp.FileName, cszBusStationByBusDate, frmTemp.CustomData
    End If




End Sub
Private Sub mnu_PreSell_Click()
    Dim lHelpContextID As Long
'    frmPreSell.m_bBySeller = True
'    lHelpContextID = frmBusStationCon.HelpContextID
    frmPreSell.ZOrder 0
    frmPreSell.Show vbModal, Me
    If frmPreSell.m_bOk Then
        Dim frmTemp As IConditionForm
        Dim frmNewReport As New frmReport
        Set frmTemp = frmPreSell
'        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport frmTemp.RecordsetData, frmTemp.FileName, "Ԥ��Ʊƽ�ⱨ��.xls", frmTemp.CustomData
    End If
End Sub
Private Sub mnu_PreSellLst_Click()
    Dim lHelpContextID As Long
'    frmPreSell.m_bBySeller = True
'    lHelpContextID = frmBusStationCon.HelpContextID
    frmPreSellLst.ZOrder 0
    frmPreSellLst.Show vbModal, Me
    If frmPreSellLst.m_bOk Then
        Dim frmTemp As IConditionForm
        Dim frmNewReport As New frmReport
        Set frmTemp = frmPreSellLst
'        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport frmTemp.RecordsetData, frmTemp.FileName, "Ԥ��Ʊƽ����ϸ��.xls", frmTemp.CustomData
    End If
End Sub

Private Sub mnu_BusStationByBusStation_Click()

    Dim lHelpContextID As Long
    frmBusStationCon.m_bBySeller = False
    lHelpContextID = frmBusStationCon.HelpContextID
    frmBusStationCon.Show vbModal, Me
    If frmBusStationCon.m_bOk Then
        Dim frmTemp As IConditionForm
        Dim frmNewReport As New frmReport
        Set frmTemp = frmBusStationCon
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport frmTemp.RecordsetData, frmTemp.FileName, cszBusStationByBusDate, frmTemp.CustomData
    End If




End Sub


'����Ӫ�հ���Ʊ
Private Sub mnu_VehicleSaleBySale_Click()
    Dim lHelpContextID As Long
    lHelpContextID = frmSelectVehicleCompany.HelpContextID

    frmSelectVehicleCompany.Show vbModal, Me
    If frmSelectVehicleCompany.m_bOk Then
        Dim frmTemp As IConditionForm
        Dim frmNewReport As New frmReport
        Set frmTemp = frmSelectVehicleCompany
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport frmTemp.RecordsetData, frmTemp.FileName, cszBusStationByBusDate1, frmTemp.CustomData
    End If

End Sub

'����Ӫ�հ���Ʊ
Private Sub mnu_VehicleSaleByCheck_Click()
    Dim lHelpContextID As Long
    lHelpContextID = frmSelectVehicleCompany.HelpContextID

    frmSelectVehicleCompany.Show vbModal, Me
    If frmSelectVehicleCompany.m_bOk Then
        Dim frmTemp As IConditionForm
        Dim frmNewReport As New frmReport
        Set frmTemp = frmSelectVehicleCompany
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport frmTemp.RecordsetData, frmTemp.FileName, cszBusStationByBusDate1, frmTemp.CustomData
    End If
    
End Sub





'��Ʊ����;��վ����
Private Sub mnu_CheckBusStationStat_Click()
    Dim lHelpContextID As Long
    lHelpContextID = frmBusStationBusreport.HelpContextID
    frmBusStationBusreport.m_nIsCheck = SNBusFromCheck
    frmBusStationBusreport.Show vbModal, Me
    If frmBusStationBusreport.m_bOk Then
        Dim frmTemp As IConditionForm
        Dim frmNewReport As New frmReport
        Set frmTemp = frmBusStationBusreport
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport frmTemp.RecordsetData, frmTemp.FileName, cszBusStationByBusDate1, frmTemp.CustomData
    End If
End Sub


'��Ʊ����;��վ����
Private Sub mnu_CheckVehicleStationStat_Click()
    Dim lHelpContextID As Long
    lHelpContextID = frmBusStationBusreport.HelpContextID
    frmBusStationBusreport.m_nIsCheck = SNVehicleFromCheck
    frmBusStationBusreport.Show vbModal, Me
    If frmBusStationBusreport.m_bOk Then
        Dim frmTemp As IConditionForm
        Dim frmNewReport As New frmReport
        Set frmTemp = frmBusStationBusreport
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport frmTemp.RecordsetData, frmTemp.FileName, cszBusStationByBusDate1, frmTemp.CustomData
    End If
End Sub

Private Sub mnu_InternetTkCount_Click(bSellCount As Boolean)
    Dim lHelpContextID As Long
    lHelpContextID = frmInternetTkCount.HelpContextID
    frmInternetTkCount.bSellCount = bSellCount
    frmInternetTkCount.Show vbModal, Me
    If frmInternetTkCount.m_bOk Then
        Dim frmTemp As IConditionForm
        Dim frmNewReport As New frmReport
        Set frmTemp = frmInternetTkCount
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport frmTemp.RecordsetData, frmTemp.FileName, cszUnitInterSettle, frmTemp.CustomData
    End If
    
End Sub


Private Sub mnu_InternetTkDetail_Click(bSellCount As Boolean)
    Dim lHelpContextID As Long
    lHelpContextID = frmInternetTkDetail.HelpContextID
    frmInternetTkDetail.bSellCount = bSellCount
    frmInternetTkDetail.Show vbModal, Me
    If frmInternetTkDetail.m_bOk Then
        Dim frmTemp As IConditionForm
        Dim frmNewReport As New frmReport
        Set frmTemp = frmInternetTkDetail
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport frmTemp.RecordsetData, frmTemp.FileName, cszUnitInterSettle, frmTemp.CustomData
    End If
    
End Sub


