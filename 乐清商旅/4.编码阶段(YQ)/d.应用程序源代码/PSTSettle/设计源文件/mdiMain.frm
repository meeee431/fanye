VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "actbar2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm mdiMain 
   BackColor       =   &H8000000C&
   Caption         =   "·������"
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

Const cszCompanySettleDetail = "��˾������ϸ��.xls"
Const cszVehicleSettleDetail = "����������ϸ��.xls"
Const cszBusSettleDetail = "���ν�����ϸ��.xls"
Const cszSettleSheetStat = "·��������ϸ��.xls"
Const cszCompanySettleStat = "��˾������ܱ�.xls"
Const cszVehicleSettleStat = "����������ܱ�.xls"
Const cszBusSettleStat = "���ν�����ܱ�.xls"
Const cszVehicleSettleStatByMonth = "������������±�.xls"
Const cszVehicleFixFee = "�����̶����ñ�.xls"
Const cszBusFixFee = "���ι̶����ñ�.xls"


Const cszProtocol = "Э����ϸ.xls"
Const cszVehicleProtocol = "����Э����ϸ.xls"
Const cszCompanyProtocol = "��˾Э����ϸ.xls"
Const cszVehicleSettlePrice = "�����������ϸ.xls"
Const cszCompanySettlePrice = "��˾�������ϸ.xls"


Private Sub MDIForm_Load()
    AddControlsToActBar
    '״̬��
    ShowSBInfo "", ESB_WorkingInfo
    ShowSBInfo "", ESB_ResultCountInfo
    ShowSBInfo EncodeString(g_oActiveUser.UserID) & g_oActiveUser.UserName, ESB_UserInfo
    ShowSBInfo Format(g_oActiveUser.LoginTime, "HH:mm"), ESB_LoginTime
    SetPrintEnabled False
End Sub

'����ActiveBar�Ŀؼ�
Private Sub AddControlsToActBar()
'    abMenu.Bands("bndTitleTop").Tools("tblTitleTop").Custom = ptTitleTop
    abMenu.Bands("statusBar").Tools("progressBar").Custom = pbLoad
End Sub
                     
'���ò˵��Ŀ�����
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
    'ϵͳ
    Case "mnu_SetOption"
        frmSetOption.Show vbModal
    
    Case "mnu_BaseInfo"      ' Э�����
        frmBaseInfo.ZOrder 0
        frmBaseInfo.Show

        
    Case "mnu_CompanyProtocol" '��˾Э������
        frmCompanyProtocol.ZOrder 0
        frmCompanyProtocol.Show

        
    Case "mnu_VehicleProtocol"  '����Э������
        frmVehicleProtocol.ZOrder 0
        frmVehicleProtocol.Show
        
    Case "mnu_BusProtocol"  '����Э������
        frmBusProtocol.ZOrder 0
        frmBusProtocol.Show

        
    Case "mnu_CompanySettlePrice" '��˾���������
        frmCompanySettlePrice.ZOrder 0
        frmCompanySettlePrice.Show

        
    Case "mnu_VehicleSettlePrice" '�������������
        frmVehicleSettlePrice.ZOrder 0
        frmVehicleSettlePrice.Show
        
    Case "mnu_BusSettlePrice" '���ν��������
        frmBusSettlePrice.ZOrder 0
        frmBusSettlePrice.Show

    Case "mnu_AllFixFee" '�����̶�����
        
        Set frmTemp = New frmAllVehicleFixFee
        frmTemp.m_szFormStatus = EFS_Vehicle
        frmTemp.ZOrder 0
        frmTemp.Show
        
    Case "mnu_BusFixFee" '���ι̶�����
        Set frmTemp = New frmAllVehicleFixFee
        frmTemp.m_szFormStatus = EFS_Bus
        frmTemp.ZOrder 0
        frmTemp.Show
'        frmAllVehicleFixFee.m_szFormStatus = EFS_Bus
'        frmAllVehicleFixFee.ZOrder 0
'        frmAllVehicleFixFee.Show
        
    Case "mnu_HalveCompany"  '����ƽ�ֹ�˾
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
    
    '����ͳ��
    Case "mnu_SettleSheetStat"  '·��������ϸ��
        mnu_SettleSheetStat_Click
    Case "mnu_CompanyDetail" '��˾������ϸ��
        mnu_CompanyDetail_Click
    Case "mnu_VehicleDetail" '����������ϸ��
        mnu_VehicleDetail_Click
    Case "mnu_BusDetail" '���ν�����ϸ��
        mnu_BusDetail_Click
    Case "mnu_CompanySettleStat" '��˾������ܱ�
        mnu_CompanySettleStat_Click
    Case "mnu_VehicleSettleStat" '����������ܱ�
        mnu_VehicleSettleStat_Click
    Case "mnu_BusSettleStat" '���ν�����ܱ�
        mnu_BusSettleStat_Click
    Case "mnu_VehicleSettleStatByMonth" '������������±���
        mnu_VehicleSettleStatByMonth_Click
    
    Case "mnu_vehiclebalancestat" '��������ƽ���
        mnu_vehiclebalancestat_Click
    Case "mnu_busbalancestat" '���ν���ƽ���
        mnu_busbalancestat_Click
    
    
    
    Case "mnu_protocol"
        '��ѯЭ��
        mnu_Protocol
    Case "mnu_vehicleprotocol"
        '����Э��
        mnu_VehicleProtocol
    Case "mnu_companyprotocol"
        '��˾Э��
        mnu_CompanyProtocol
    Case "mnu_vehiclesettleprice"
        '���������
        mnu_VehicleSettlePrice
    Case "mnu_companysettleprice"
        '��˾�����
        mnu_CompanySettlePrice
    Case "mnu_VehicleFixFeeStat"
        '�����̶����ñ���
        mnu_VehicleFixFeeStat
    Case "mnu_BusFixFeeStat"
        '���ι̶����ñ���
        mnu_BusFixFeeStat
        
        
    '·������
    Case "mi_SettleSheet"    '·���������
        frmAllSettleSheets.ZOrder 0
        frmAllSettleSheets.Show

         
    Case "mi_NewSettleSheet"    '·��������
'        frmWizSplitSettle.ZOrder 0
        frmWizSplitSettle.Show vbModal
    Case "mi_WizSplitSettleSheetManual"
'        frmWizSplitSettleBack.ZOrder 0
        frmWizSplitSettleBack.Show vbModal
         
    Case "mi_RePrintSettleSheet" '�ش���㵥
        frmRePrintSettleSheet.ZOrder 0
        frmRePrintSettleSheet.Show vbModal
    
    
    Case "mi_ViewCheckSheet"
        '�쿴·��

        frmAllSheet.m_dtStartDate = GetFirstMonthDay(Date)
        frmAllSheet.m_dtEndDate = GetLastMonthDay(Date)
        
        frmAllSheet.ZOrder 0
        frmAllSheet.Show
        
        
    Case "mi_ModifyCheckSheet"
        '�޸�·��
        frmModifySheet.Show vbModal
        
    Case "mi_MakeSheetStatTemp"
        MakeSheetStatTemp
        
    '����
    Case "mnu_TitleH"
'        mnu_TitleH_Click
    Case "mnu_TitleV"
'        mnu_TitleV_Click
    Case "mnu_Cascade"
'        mnu_Cascade_Click
    Case "mnu_ArrangeIcon"
'        mnu_ArrangeIcon_Click
    '����
    Case "mnu_HelpIndex"
'        frmWizSplitSettle.Show
    Case "mnu_HelpContent"
'        mnu_HelpContent_Click
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

'·��������ϸ��
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
        
        'ȡ�ü�¼��
        Set rsSettleDetail = m_oReport.SettleSheetStat(frmSettleSheetStat.m_dtStartDate, DateAdd("d", 1, frmSettleSheetStat.m_dtEndDate))
        WriteProcessBar False, , , ""
        ReDim vCustomData(1 To 3, 1 To 2)
        vCustomData(1, 1) = "��ʼ����"
        vCustomData(1, 2) = ToDBDate(frmSettleSheetStat.m_dtStartDate)
        vCustomData(2, 1) = "��������"
        vCustomData(2, 2) = ToDBDate(frmSettleSheetStat.m_dtEndDate)
        vCustomData(3, 1) = "��ӡ"
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

'��˾������ϸ��
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
        
        
        'ȡ�ü�¼��
        Set rsSettleDetail = m_oReport.CompanySettleDetail(frmCompanySettleStat.m_dtStartDate, DateAdd("d", 1, frmCompanySettleStat.m_dtEndDate), ResolveDisplay(frmCompanySettleStat.m_szCompanyID))
        WriteProcessBar True, , , ""
        ReDim vCustomData(1 To 4, 1 To 2)
        vCustomData(1, 1) = "��ʼ����"
        vCustomData(1, 2) = ToDBDate(frmCompanySettleStat.m_dtStartDate)
        vCustomData(2, 1) = "��������"
        vCustomData(2, 2) = ToDBDate(frmCompanySettleStat.m_dtEndDate)
        vCustomData(3, 1) = "��ӡ"
        vCustomData(3, 2) = g_oActiveUser.UserName
        vCustomData(4, 1) = "��˾"
        vCustomData(4, 2) = IIf(frmCompanySettleStat.m_szCompanyID = "", "���й�˾", frmCompanySettleStat.m_szCompanyID)
        
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport rsSettleDetail, cszCompanySettleDetail, frmCompanySettleStat.Caption, vCustomData, 10
        
        WriteProcessBar False, , , ""
    End If
    
    
    
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

'����������ϸ��
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
        
        'ȡ�ü�¼��
        Set rsSettleDetail = m_oReport.VehicleSettleDetail(frmVehicleSettleStat.m_dtStartDate, DateAdd("d", 1, frmVehicleSettleStat.m_dtEndDate), ResolveDisplay(frmVehicleSettleStat.m_szVehicleID), ResolveDisplay(frmVehicleSettleStat.m_szCompanyID), frmVehicleSettleStat.m_nStatus, frmVehicleSettleStat.m_nQueryNegativeType, frmVehicleSettleStat.m_bStatBySettleDate)
        WriteProcessBar True, , , ""
        ReDim vCustomData(1 To 8, 1 To 2)
        vCustomData(1, 1) = "��ʼ����"
        vCustomData(1, 2) = ToDBDate(frmVehicleSettleStat.m_dtStartDate)
        vCustomData(2, 1) = "��������"
        vCustomData(2, 2) = ToDBDate(frmVehicleSettleStat.m_dtEndDate)
        vCustomData(3, 1) = "��ӡ"
        vCustomData(3, 2) = g_oActiveUser.UserName
        vCustomData(4, 1) = "����"
        vCustomData(4, 2) = IIf(frmVehicleSettleStat.m_szVehicleID = "", "���г���", frmVehicleSettleStat.m_szVehicleTagNo)
        vCustomData(5, 1) = "��˾"
        vCustomData(5, 2) = IIf(frmVehicleSettleStat.m_szCompanyID = "", "���й�˾", frmVehicleSettleStat.m_szCompanyName)
        
        vCustomData(6, 1) = "���"
        vCustomData(6, 2) = GetSettleSheetStatusString(frmVehicleSettleStat.m_nStatus)
        
        
        vCustomData(7, 1) = "�������"
        vCustomData(7, 2) = GetQueryNegativeStatusString(frmVehicleSettleStat.m_nQueryNegativeType)
        
        
        vCustomData(8, 1) = "ͳ�Ʒ�ʽ"
        vCustomData(8, 2) = IIf(frmVehicleSettleStat.m_bStatBySettleDate, "����������ͳ��", "����������ͳ��")
        
        
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport rsSettleDetail, cszVehicleSettleDetail, frmVehicleSettleStat.Caption, vCustomData, 10
        
        WriteProcessBar False, , , ""
    End If
    
    
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

'�����̶����ñ���
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
        
        'ȡ�ü�¼��
        Set rsVehicleFixFee = m_oReport.GetAllVehicleFixFee(ResolveDisplay(frmVehicleFixFeeReport.m_szVehicleID), ResolveDisplay(frmVehicleFixFeeReport.m_szCompanyID), frmVehicleFixFeeReport.m_dtStartDate, DateAdd("d", 1, frmVehicleFixFeeReport.m_dtEndDate), , frmVehicleFixFeeReport.m_nStatus)
        WriteProcessBar True, , , ""
        ReDim vCustomData(1 To 8, 1 To 2)
        vCustomData(1, 1) = "��ʼ����"
        vCustomData(1, 2) = ToDBDate(frmVehicleFixFeeReport.m_dtStartDate)
        vCustomData(2, 1) = "��������"
        vCustomData(2, 2) = ToDBDate(frmVehicleFixFeeReport.m_dtEndDate)
        vCustomData(3, 1) = "��ӡ"
        vCustomData(3, 2) = g_oActiveUser.UserName
        vCustomData(4, 1) = "����"
        vCustomData(4, 2) = IIf(frmVehicleFixFeeReport.m_szVehicleID = "", "���г���", frmVehicleFixFeeReport.m_szVehicleTagNo)
        vCustomData(5, 1) = "��˾"
        vCustomData(5, 2) = IIf(frmVehicleFixFeeReport.m_szCompanyID = "", "���й�˾", frmVehicleFixFeeReport.m_szCompanyName)
        
        vCustomData(6, 1) = "���"
        vCustomData(6, 2) = GetFixFeeStatusName(frmVehicleFixFeeReport.m_nStatus)
        
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport rsVehicleFixFee, cszVehicleFixFee, frmVehicleFixFeeReport.Caption, vCustomData, 10
        
        WriteProcessBar False, , , ""
    End If
    
    
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

'���ι̶����ñ���
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
        
        'ȡ�ü�¼��
        Set rsBusFixFee = m_oReport.GetAllBusFixFee(ResolveDisplay(frmBusFixFeeReport.m_szBusID), ResolveDisplay(frmBusFixFeeReport.m_szCompanyID), frmBusFixFeeReport.m_dtStartDate, DateAdd("d", 1, frmBusFixFeeReport.m_dtEndDate), , frmBusFixFeeReport.m_nStatus)
        WriteProcessBar True, , , ""
        ReDim vCustomData(1 To 8, 1 To 2)
        vCustomData(1, 1) = "��ʼ����"
        vCustomData(1, 2) = ToDBDate(frmBusFixFeeReport.m_dtStartDate)
        vCustomData(2, 1) = "��������"
        vCustomData(2, 2) = ToDBDate(frmBusFixFeeReport.m_dtEndDate)
        vCustomData(3, 1) = "��ӡ"
        vCustomData(3, 2) = g_oActiveUser.UserName
        vCustomData(4, 1) = "����"
        vCustomData(4, 2) = IIf(frmBusFixFeeReport.m_szBusID = "", "���г���", frmBusFixFeeReport.m_szBusTagNo)
        vCustomData(5, 1) = "��˾"
        vCustomData(5, 2) = IIf(frmBusFixFeeReport.m_szCompanyID = "", "���й�˾", frmBusFixFeeReport.m_szCompanyName)
        
        vCustomData(6, 1) = "���"
        vCustomData(6, 2) = GetFixFeeStatusName(frmBusFixFeeReport.m_nStatus)
        
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport rsBusFixFee, cszBusFixFee, frmBusFixFeeReport.Caption, vCustomData, 10
        
        WriteProcessBar False, , , ""
    End If
    
    
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

'���ν�����ϸ��
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
        
        'ȡ�ü�¼��
        Set rsSettleDetail = m_oReport.BusSettleDetail(frmBusSettleStat.m_dtStartDate, DateAdd("d", 1, frmBusSettleStat.m_dtEndDate), ResolveDisplay(frmBusSettleStat.m_szBusID), ResolveDisplay(frmBusSettleStat.m_szCompanyID), frmBusSettleStat.m_nStatus, frmBusSettleStat.m_nQueryNegativeType, frmBusSettleStat.m_bStatBySettleDate)
        WriteProcessBar True, , , ""
        ReDim vCustomData(1 To 8, 1 To 2)
        vCustomData(1, 1) = "��ʼ����"
        vCustomData(1, 2) = ToDBDate(frmBusSettleStat.m_dtStartDate)
        vCustomData(2, 1) = "��������"
        vCustomData(2, 2) = ToDBDate(frmBusSettleStat.m_dtEndDate)
        vCustomData(3, 1) = "��ӡ"
        vCustomData(3, 2) = g_oActiveUser.UserName
        vCustomData(4, 1) = "����"
        vCustomData(4, 2) = IIf(frmBusSettleStat.m_szBusID = "", "���г���", frmBusSettleStat.m_szBusTagNo)
        vCustomData(5, 1) = "��˾"
        vCustomData(5, 2) = IIf(frmBusSettleStat.m_szCompanyID = "", "���й�˾", frmBusSettleStat.m_szCompanyName)
        
        vCustomData(6, 1) = "���"
        vCustomData(6, 2) = GetSettleSheetStatusString(frmBusSettleStat.m_nStatus)
        
        
        vCustomData(7, 1) = "�������"
        vCustomData(7, 2) = GetQueryNegativeStatusString(frmBusSettleStat.m_nQueryNegativeType)
        
        
        vCustomData(8, 1) = "ͳ�Ʒ�ʽ"
        vCustomData(8, 2) = IIf(frmBusSettleStat.m_bStatBySettleDate, "����������ͳ��", "����������ͳ��")
        
        
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport rsSettleDetail, cszBusSettleDetail, frmBusSettleStat.Caption, vCustomData, 10
        
        WriteProcessBar False, , , ""
    End If
    
    
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

'��˾������ܱ�
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
        'ȡ�ü�¼��
        Set rsSettleDetail = m_oReport.CompanySettleStat(frmCompanySettleStat.m_dtStartDate, DateAdd("d", 1, frmCompanySettleStat.m_dtEndDate), ResolveDisplay(frmCompanySettleStat.m_szCompanyID))
        WriteProcessBar True, , , ""
        ReDim vCustomData(1 To 4, 1 To 2)
        vCustomData(1, 1) = "��ʼ����"
        vCustomData(1, 2) = ToDBDate(frmCompanySettleStat.m_dtStartDate)
        vCustomData(2, 1) = "��������"
        vCustomData(2, 2) = ToDBDate(frmCompanySettleStat.m_dtEndDate)
        vCustomData(3, 1) = "��ӡ"
        vCustomData(3, 2) = g_oActiveUser.UserName
        vCustomData(4, 1) = "��˾"
        vCustomData(4, 2) = IIf(frmCompanySettleStat.m_szCompanyID = "", "���й�˾", frmCompanySettleStat.m_szCompanyID)
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport rsSettleDetail, cszCompanySettleStat, frmCompanySettleStat.Caption, vCustomData, 10
        WriteProcessBar False, , , ""
    End If
    
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

'����������ܱ�
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
        'ȡ�ü�¼��
        Set rsSettleDetail = m_oReport.VehicleSettleStat(frmVehicleSettleStat.m_dtStartDate, DateAdd("d", 1, frmVehicleSettleStat.m_dtEndDate), ResolveDisplay(frmVehicleSettleStat.m_szVehicleID), ResolveDisplay(frmVehicleSettleStat.m_szCompanyID), frmVehicleSettleStat.m_nQueryNegativeType, frmVehicleSettleStat.m_nStatus, frmVehicleSettleStat.m_bStatBySettleDate)
        WriteProcessBar True, , , ""
        ReDim vCustomData(1 To 7, 1 To 2)
        vCustomData(1, 1) = "��ʼ����"
        vCustomData(1, 2) = ToDBDate(frmVehicleSettleStat.m_dtStartDate)
        vCustomData(2, 1) = "��������"
        vCustomData(2, 2) = ToDBDate(frmVehicleSettleStat.m_dtEndDate)
        vCustomData(3, 1) = "��ӡ"
        vCustomData(3, 2) = g_oActiveUser.UserName
        vCustomData(4, 1) = "����"
        vCustomData(4, 2) = IIf(frmVehicleSettleStat.m_szVehicleID = "", "���г���", frmVehicleSettleStat.m_szVehicleTagNo)
        
        vCustomData(5, 1) = "��˾"
        vCustomData(5, 2) = IIf(frmVehicleSettleStat.m_szCompanyID = "", "���й�˾", frmVehicleSettleStat.m_szCompanyName)
        
        vCustomData(6, 1) = "�ϳ�վ"
        vCustomData(6, 2) = g_oActiveUser.UserUnitName
        
        vCustomData(7, 1) = "ͳ�Ʒ�ʽ"
        vCustomData(7, 2) = IIf(frmVehicleSettleStat.m_bStatBySettleDate, "����������ͳ��", "����������ͳ��")
        
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport rsSettleDetail, cszVehicleSettleStat, frmVehicleSettleStat.Caption, vCustomData, 10
        
        WriteProcessBar False, , , ""
    End If
    
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

'���ν�����ܱ�
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
        'ȡ�ü�¼��
        Set rsSettleDetail = m_oReport.BusSettleStat(frmBusSettleStat.m_dtStartDate, DateAdd("d", 1, frmBusSettleStat.m_dtEndDate), ResolveDisplay(frmBusSettleStat.m_szBusID), ResolveDisplay(frmBusSettleStat.m_szCompanyID), , , frmBusSettleStat.m_bStatBySettleDate)
        WriteProcessBar True, , , ""
        ReDim vCustomData(1 To 7, 1 To 2)
        vCustomData(1, 1) = "��ʼ����"
        vCustomData(1, 2) = ToDBDate(frmBusSettleStat.m_dtStartDate)
        vCustomData(2, 1) = "��������"
        vCustomData(2, 2) = ToDBDate(frmBusSettleStat.m_dtEndDate)
        vCustomData(3, 1) = "��ӡ"
        vCustomData(3, 2) = g_oActiveUser.UserName
        vCustomData(4, 1) = "����"
        vCustomData(4, 2) = IIf(frmBusSettleStat.m_szBusID = "", "���г���", frmBusSettleStat.m_szBusTagNo)
        
        vCustomData(5, 1) = "��˾"
        vCustomData(5, 2) = IIf(frmBusSettleStat.m_szCompanyID = "", "���й�˾", frmBusSettleStat.m_szCompanyName)
        
        vCustomData(6, 1) = "�ϳ�վ"
        vCustomData(6, 2) = g_oActiveUser.UserUnitName
        
        vCustomData(7, 1) = "ͳ�Ʒ�ʽ"
        vCustomData(7, 2) = IIf(frmBusSettleStat.m_bStatBySettleDate, "����������ͳ��", "����������ͳ��")
        
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport rsSettleDetail, cszBusSettleStat, frmBusSettleStat.Caption, vCustomData, 10
        
        WriteProcessBar False, , , ""
    End If
    
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub


'������������±�
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
        'ȡ�ü�¼��
        Set rsSettleDetail = m_oReport.VehicleSettleStatByMonth(frmVehicleSettleStat.m_dtStartDate, DateAdd("d", 1, frmVehicleSettleStat.m_dtEndDate), ResolveDisplay(frmVehicleSettleStat.m_szVehicleID), ResolveDisplay(frmVehicleSettleStat.m_szCompanyID), frmVehicleSettleStat.m_nQueryNegativeType, frmVehicleSettleStat.m_nStatus)
        WriteProcessBar True, , , ""
        ReDim vCustomData(1 To 6, 1 To 2)
        vCustomData(1, 1) = "��ʼ����"
        vCustomData(1, 2) = ToDBDate(frmVehicleSettleStat.m_dtStartDate)
        vCustomData(2, 1) = "��������"
        vCustomData(2, 2) = ToDBDate(frmVehicleSettleStat.m_dtEndDate)
        vCustomData(3, 1) = "��ӡ"
        vCustomData(3, 2) = g_oActiveUser.UserName
        vCustomData(4, 1) = "����"
        vCustomData(4, 2) = IIf(frmVehicleSettleStat.m_szVehicleID = "", "���г���", frmVehicleSettleStat.m_szVehicleTagNo)
        
        vCustomData(5, 1) = "��˾"
        vCustomData(5, 2) = IIf(frmVehicleSettleStat.m_szCompanyID = "", "���й�˾", frmVehicleSettleStat.m_szCompanyName)
        
        vCustomData(6, 1) = "�ϳ�վ"
        vCustomData(6, 2) = g_oActiveUser.UserUnitName
        
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport rsSettleDetail, cszVehicleSettleStatByMonth, frmVehicleSettleStat.Caption, vCustomData, 10
        
        WriteProcessBar False, , , ""
    End If
    
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub




'Э��
Private Sub mnu_Protocol()
    On Error GoTo ErrHandle
    
    Dim lHelpContextID As Long
    
    Dim m_oReport As New Report
    Dim vCustomData As Variant
    Dim rsTemp As Recordset
    Dim frmNewReport As New frmReport
    
    m_oReport.Init g_oActiveUser
    'ȡ�ü�¼��
    Set rsTemp = m_oReport.GetAllProtocolRS()
    
    
    frmNewReport.m_lHelpContextID = lHelpContextID
    frmNewReport.ShowReport rsTemp, cszProtocol, "Э����ϸ", vCustomData
    

Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

'����Э��
Private Sub mnu_VehicleProtocol()
    On Error GoTo ErrHandle
    
    Dim lHelpContextID As Long
    
    Dim m_oReport As New Report
    Dim vCustomData As Variant
    Dim rsTemp As Recordset
    Dim frmNewReport As New frmReport
    
    m_oReport.Init g_oActiveUser
    'ȡ�ü�¼��
    Set rsTemp = m_oReport.GetVehicleProtocolRS()
    
    
    frmNewReport.m_lHelpContextID = lHelpContextID
    frmNewReport.ShowReport rsTemp, cszVehicleProtocol, "����Э����ϸ", vCustomData
    

Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

'��˾Э��
Private Sub mnu_CompanyProtocol()
    On Error GoTo ErrHandle
    
    Dim lHelpContextID As Long
    
    Dim m_oReport As New Report
    Dim vCustomData As Variant
    Dim rsTemp As Recordset
    Dim frmNewReport As New frmReport
    
    m_oReport.Init g_oActiveUser
    'ȡ�ü�¼��
    Set rsTemp = m_oReport.GetAllCompanyProtocolRS()
    
    
    frmNewReport.m_lHelpContextID = lHelpContextID
    frmNewReport.ShowReport rsTemp, cszCompanyProtocol, "��˾Э����ϸ", vCustomData
    

Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

'��˾�����
Private Sub mnu_CompanySettlePrice()
    On Error GoTo ErrHandle
    
    Dim lHelpContextID As Long
    
    Dim m_oReport As New Report
    Dim vCustomData As Variant
    Dim rsTemp As Recordset
    Dim frmNewReport As New frmReport
    
    m_oReport.Init g_oActiveUser
    'ȡ�ü�¼��
    Set rsTemp = m_oReport.GetCompanySettlePriceLstRS()
    
    
    frmNewReport.m_lHelpContextID = lHelpContextID
    frmNewReport.ShowReport rsTemp, cszCompanySettlePrice, "��˾�������ϸ", vCustomData
    

Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub


'���������
Private Sub mnu_VehicleSettlePrice()
    On Error GoTo ErrHandle
    
    Dim lHelpContextID As Long
    
    Dim m_oReport As New Report
    Dim vCustomData As Variant
    Dim rsTemp As Recordset
    Dim frmNewReport As New frmReport
    
    m_oReport.Init g_oActiveUser
    'ȡ�ü�¼��
    Set rsTemp = m_oReport.GetVehicleSettlePriceLstRS()
    
    
    frmNewReport.m_lHelpContextID = lHelpContextID
    frmNewReport.ShowReport rsTemp, cszVehicleSettlePrice, "�����������ϸ", vCustomData
    

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
    If MsgBox("���Ƿ����Ҫ�˳���ϵͳ?", vbQuestion + vbYesNoCancel, "����") = vbYes Then
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
    '����ƽ���
    
    Dim lHelpContextID As Long
    lHelpContextID = frmVehicleBalance.HelpContextID
    
    frmVehicleBalance.Show vbModal, Me
    If frmVehicleBalance.m_bOk Then
        Dim frmTemp As IConditionForm
        Dim frmNewReport As New frmReport
        Set frmTemp = frmVehicleBalance
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport frmTemp.RecordsetData, frmTemp.FileName, "����ƽ���", frmTemp.CustomData, 2
        
    End If
    
End Sub

Private Sub mnu_busbalancestat_Click()
    '����ƽ���
    
    Dim lHelpContextID As Long
    lHelpContextID = frmBusBalance.HelpContextID
    
    frmBusBalance.Show vbModal, Me
    If frmBusBalance.m_bOk Then
        Dim frmTemp As IConditionForm
        Dim frmNewReport As New frmReport
        Set frmTemp = frmBusBalance
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport frmTemp.RecordsetData, frmTemp.FileName, "����ƽ���", frmTemp.CustomData, 2
        
    End If
    
End Sub


Private Sub MakeSheetStatTemp()
    Dim oSplit As New Split
    On Error GoTo ErrorHandle
    SetBusy
    ShowSBInfo "�������ɵ���������ʱͳ������"
    abMenu.Refresh
    
    oSplit.Init g_oActiveUser
    oSplit.MakeSheetStatTemp Date
    SetNormal
    ShowSBInfo ""
    MsgBox "���ɵ���������ʱͳ���������", vbInformation, Me.Caption
    Exit Sub
ErrorHandle:
    SetNormal
    ShowSBInfo ""
    ShowErrorMsg
End Sub


