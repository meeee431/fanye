VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "actbar2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.MDIForm mdiMain 
   BackColor       =   &H8000000C&
   Caption         =   "行包管理"
   ClientHeight    =   8040
   ClientLeft      =   720
   ClientTop       =   2670
   ClientWidth     =   12405
   HelpContextID   =   7000201
   Icon            =   "mdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   WindowState     =   2  'Maximized
   Begin ActiveBar2LibraryCtl.ActiveBar2 abMenu 
      Align           =   1  'Align Top
      Height          =   8040
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12405
      _LayoutVersion  =   1
      _ExtentX        =   21881
      _ExtentY        =   14182
      _DataPath       =   ""
      Bands           =   "mdiMain.frx":16AC2
      Begin VB.PictureBox ptTitleTop 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   -240
         Picture         =   "mdiMain.frx":232C4
         ScaleHeight     =   687.72
         ScaleMode       =   0  'User
         ScaleWidth      =   15405
         TabIndex        =   2
         Top             =   2760
         Width           =   15405
         Begin RTComctl3.CoolButton cmdClose 
            Height          =   390
            Left            =   7830
            TabIndex        =   3
            ToolTipText     =   "返回"
            Top             =   240
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   688
            BTYPE           =   12
            TX              =   ""
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
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
            MICON           =   "mdiMain.frx":24903
            PICN            =   "mdiMain.frx":2491F
            PICH            =   "mdiMain.frx":25814
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
               Name            =   "宋体"
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
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************
'Last Modify By: 陆勇庆  2005-8-16
'Last Modify In: 报表的显示
'*******************************************************************************
Option Explicit
Public cszLuggageSaler As String
Public cszStation As String
Public cszSellstation As String

Const cszLugSalerDayTotal = "行包员每日结算.xls"
Const cszLugSplitList = "行包营收拆算一览表.xls"
Const cszLugCompanySplit = "行包营收拆算报表.xls"
Const cszAcceptSettle = "行包员每日结算_有票明细.xls"
Public szFromId As String


Private Sub MDIForm_Load()
    AddControlsToActBar
    '状态条
    ShowSBInfo "", ESB_WorkingInfo
    ShowSBInfo "", ESB_ResultCountInfo
    ShowSBInfo EncodeString(m_oAUser.UserID) & m_oAUser.UserName, ESB_UserInfo
    ShowSBInfo Format(m_oAUser.LoginTime, "HH:mm"), ESB_LoginTime
'    ActiveToolBar False
    
    SetPrintEnabled False
    
'    abMenu.Bands("mnu_System").Tools("mnu_OptionSet").Enabled = True
End Sub
Private Sub abMenu_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
On Error GoTo ErrHandle
Select Case Tool.name
    '系统
    Case "mnu_BaseInfo"      '
        frmBaseInfo.ZOrder 0
        frmBaseInfo.Show
    Case "mnu_itemformula"
        frmItemFormula.Show vbModal
    
    
    '行包统计
    Case "mi_LugSalerDayTotal"
         mnu_LugSalerDayTotal_Click
    
    Case "mi_LugAccepterSettle"
        mi_LugAccepterSettle_Click '每日结算_有票明细
    
    Case "mi_LugSalerDay"
        frmLuggagerDayStat.Caption = "行包员受理日报"
        szFromId = 1   '行包员受理日报
        
        mnu_LugSaler_Click
    Case "mi_LugSalerMonth"
        frmLuggagerDayStat.Caption = "行包员受理月报"
         szFromId = 2  '行包员受理月报
        mnu_LugSaler_Click
    Case "mi_LugSalerYear"
         frmLuggagerDayStat.Caption = "行包员受理年报"
         szFromId = 3   '行包员受理年报
        mnu_LugSaler_Click
    Case "mi_LugSalerReport"
      frmLuggagerDayStat.Caption = "行包员受理简报"
          szFromId = 9   '行包员受理简报
          mnu_LugSaler_Click
        
   
   
    
    Case "mi_LugStationDay"
        frmStationDayStat.Caption = "车站行包营收日报"
        szFromId = 4 '车站行包营收日报
        mnu_LugStation_Click
    Case "mi_LugStationMonth"
        frmStationDayStat.Caption = "车站行包营收月报"
        szFromId = 5 '车站行包营收月报
        mnu_LugStation_Click
    Case "mi_LugStationYear"
        frmStationDayStat.Caption = "车站行包营收年报"
         szFromId = 5 '车站行包营收年报
        mnu_LugStation_Click
    Case "mi_LugStationReport"
      frmStationDayStat.Caption = "车站行包营收简报"
       szFromId = 16 '车站行包营收简报
      mnu_LugStation_Click
        
        
        
    
     Case "mi_SellStationDayStat"
        frmSellStationDayStat.Caption = "售票站行包营收日报"
          
        szFromId = 6 '售票站行包营收日报
        mnu_LugSellStation_Click
    Case "mi_SellStationMonthStat"
        frmSellStationDayStat.Caption = "售票站行包营收月报"
        szFromId = 7 '售票站行包营收月报
        mnu_LugSellStation_Click
    Case "mi_SellStationYearStat"
        frmSellStationDayStat.Caption = "售票站行包营收年报"
         szFromId = 8 '售票站行包营收年报
        mnu_LugSellStation_Click
    Case "mi_SellStationReport"
        frmSellStationDayStat.Caption = "售票站行包营收简报"
         szFromId = 17 '售票站行包营收年报
        mnu_LugSellStation_Click
    
    
    Case "mi_SplitCompanyCheckStat"
        '拆帐公司签发简报
        mi_SplitCompanyCheckStat_Click
    Case "mi_VehicleCheckStat"
        '车辆签发简报
        mi_VehicleCheckStat_Click
    Case "mi_RouteCheckStat"
        '线路签发简报
        mi_RouteCheckStat_Click
'    Case "mi_StationCheckStat"
'        '站点签发简报
        
        
        
        
    '行包拆算
    Case "mi_FinanceSheet"
         frmAllFinSheets.Caption = "行包结算单"
         szFromId = 10 '行包结算单
         frmAllFinSheets.ZOrder 0
         frmAllFinSheets.Show
    Case "mi_NewFinSheet"
         szFromId = 11 '行包拆算
         frmWizSplitLuggage.ZOrder 0
         frmWizSplitLuggage.Show vbModal
         
    Case "mi_RePrintFinSheet" '重打结算单
         frmRePrintFinSheet.ZOrder 0
         frmRePrintFinSheet.Show vbModal
         
    Case "mi_FinSheetTotalList"
         szFromId = 12 '行包营收拆算一览表
        mnu_FinSheetTotalList_Click
    
    Case "mi_CompanyFinSheetStat"
         szFromId = 13 '行包营收拆帐明细报表
        mnu_CompanyFinSheetStat_Click
        
    Case "mnu_CarrySheet"
         szFromId = 14  '行包签发单
         frmQuerySheet.ZOrder 0
         frmQuerySheet.Show
    
    Case "mnu_AcceptSheet"
         szFromId = 15 '行包受理单
         frmQueryAccept.ZOrder 0
         frmQueryAccept.Show
    Case "mnu_ModifySheetVehicle"
         frmUpdateSheet.ZOrder 0
         frmUpdateSheet.Show
    
    '窗口
    Case "mnu_TitleH"
        mnu_TitleH_Click
    Case "mnu_TitleV"
        mnu_TitleV_Click
    Case "mnu_Cascade"
        mnu_Cascade_Click
    Case "mnu_ArrangeIcon"
        mnu_ArrangeIcon_Click
    '帮助
    Case "mnu_HelpIndex"
        mnu_HelpIndex_Click
    Case "mnu_HelpContent"
        mnu_HelpContent_Click
    Case "mnu_About"
        mnu_About_Click
    
        '以下是系统部分
        
        Case "mnu_OptionSet"
            frmSysParam.Show vbModal
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
Private Sub ChangePassword()
    Dim oShell As New CommDialog
    On Error GoTo ErrorHandle
    oShell.Init m_oAUser
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
    cmdClose.Left = Me.Width - cmdClose.Width - 2000

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    End
End Sub

Public Sub SetPrintEnabled(pbEnabled As Boolean)
    '设置菜单的可用性
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
'关联ActiveBar的控件
Private Sub AddControlsToActBar()
    abMenu.Bands("bndTitleTop").Tools("tblTitleTop").Custom = ptTitleTop
    abMenu.Bands("statusBar").Tools("progressBar").Custom = pbLoad
End Sub
Private Sub mnu_About_Click()
    Dim oShell As New CommShell
    oShell.ShowAbout App.ProductName, "Luggage Man", App.FileDescription, Me.Icon, App.Major, App.Minor, App.Revision
End Sub

Private Sub mnu_ArrangeIcon_Click()
    Arrange vbArrangeIcons
End Sub

Private Sub mnu_LugSaler_Click()
On Error GoTo Error_Handle
    Dim lHelpContextID As Long
    
    
    lHelpContextID = frmLuggagerDayStat.HelpContextID
    frmLuggagerDayStat.Show vbModal
    If frmLuggagerDayStat.m_bOk Then
        
        Dim rsSellDetail As Recordset
        Dim rsDetailToShow As Recordset
        Dim adbOther() As Double
       
        Dim oLugDss As New LugDataStatSvr
        oLugDss.Init m_oAUser
        
        Dim i As Integer, nUserCount As Integer
        
        Dim szLastTicketID As String
        Dim szBeginTicketID As String
        Dim arsData() As Recordset
        Dim cszFileName As String
        Dim j As Integer
        Dim aszAllSeller() As String
        Dim nAllSeller As Integer
        Dim k As Integer
        Dim l As Integer
        
        Dim oUnit As New Unit
        oUnit.Init m_oAUser
        oUnit.Identify g_oParam.UnitID
        aszAllSeller = oUnit.GetAllUserEX(, ResolveDisplay(frmLuggagerDayStat.cboSellStation))
        nAllSeller = ArrayLength(aszAllSeller)
        
        
        nUserCount = ArrayLength(frmLuggagerDayStat.m_vaSeller)
        
        If nAllSeller > 0 Then
            
            ReDim arsData(1 To IIf(nAllSeller > nUserCount, nUserCount, nAllSeller))
'            ReDim vaCostumData(1 To IIf(nAllSeller > nUserCount, nUserCount, nAllSeller), 1 To 7, 1 To 2)
            WriteProcessBar True, , nUserCount, "正在形成记录集..."
            l = 0
            
            Dim aszSelectUser() As String
            Dim frmNewReport As New frmReport
            
            If szFromId = 9 Then
                ReDim aszSelectUser(1 To nUserCount)
                For i = 1 To nUserCount
                    aszSelectUser(i) = ResolveDisplay(frmLuggagerDayStat.m_vaSeller(i))
                Next i
'            Else
'                For i = 1 To nUserCount
'                    WriteProcessBar , i, nUserCount, "正在得到" & EncodeString(frmLuggagerDayStat.m_vaSeller(i)) & "的数据..."
'                    For k = 1 To nAllSeller
'                        If LCase(Trim(ResolveDisplay(frmLuggagerDayStat.m_vaSeller(i)))) = LCase(aszAllSeller(k, 1)) Then
'                            Exit For
'                        End If
'                    Next k
'                    If k <= nAllSeller Then
'                        l = l + 1
'                        '初始化
'                        ReDim aszSelectUser(1 To 1) As String
'
                    

                Dim vaCostumData As Variant
                '创建自定义项目集tt
                ReDim vaCostumData(1 To 3, 1 To 2)
                vaCostumData(1, 1) = "结算月份"
                vaCostumData(1, 2) = IIf(Format(frmLuggagerDayStat.m_dtWorkDate, "YYYY年MM月DD日") = Format(frmLuggagerDayStat.m_dtEndDate, "YYYY年MM月DD日"), Format(frmLuggagerDayStat.m_dtWorkDate, "YYYY年MM月DD日"), Format(frmLuggagerDayStat.m_dtWorkDate, "YYYY年MM月DD日") & " - " & Format(frmLugSplitList.m_dtEndDate, "YYYY年MM月DD日"))
                vaCostumData(2, 1) = "制表人"
                vaCostumData(2, 2) = Trim(m_oAUser.UserName)
                vaCostumData(3, 1) = "制表日期"
                vaCostumData(3, 2) = Date
                        
                        Select Case szFromId
                                Case 1
                                    cszLuggageSaler = "行包员结算日报"
                                     cszFileName = "行包员结算日报.xls"
                                    Set rsSellDetail = oLugDss.LuggagerDataDayStat(ResolveDisplay(frmLuggagerDayStat.m_SellStation), frmLuggagerDayStat.m_dtWorkDate, frmLuggagerDayStat.m_dtEndDate, aszSelectUser, frmLuggagerDayStat.m_AcceptType)
                                Case 2
                                    cszLuggageSaler = "行包员结算月报"
                                     cszFileName = "行包员结算月报.xls"
                                    Set rsSellDetail = oLugDss.LuggagerDataMonthStat(ResolveDisplay(frmLuggagerDayStat.m_SellStation), frmLuggagerDayStat.m_dtWorkDate, frmLuggagerDayStat.m_dtEndDate, aszSelectUser, frmLuggagerDayStat.m_AcceptType)
                                Case 3
                                    cszLuggageSaler = "行包员结算年报"
                                    cszFileName = "行包员结算年报.xls"
                                    Set rsSellDetail = oLugDss.LuggagerDataYearStat(ResolveDisplay(frmLuggagerDayStat.m_SellStation), frmLuggagerDayStat.m_dtWorkDate, frmLuggagerDayStat.m_dtEndDate, aszSelectUser, frmLuggagerDayStat.m_AcceptType)
                                Case 9
                                    cszLuggageSaler = "行包员结算简报"
                                    cszFileName = "行包员结算简报.xls"
                                    Set rsSellDetail = oLugDss.LuggagerStat(frmLuggagerDayStat.m_dtWorkDate, frmLuggagerDayStat.m_dtEndDate, aszSelectUser, ResolveDisplay(frmLuggagerDayStat.m_SellStation), frmLuggagerDayStat.m_AcceptType)

                        End Select
'                        Set arsData(l) = rsSellDetail
'                   End If
                    
'                Next i
                
                
                WriteProcessBar False, , , ""
            
                frmNewReport.ShowReport rsSellDetail, cszFileName, cszLuggageSaler, vaCostumData
            
'                frmNewReport.ShowReport2 arsData, cszFileName, cszLuggageSaler
            
                WriteProcessBar False, , , ""
            End If
        End If
    End If
Exit Sub
Error_Handle:
    WriteProcessBar False, , , ""
    ShowErrorMsg
End Sub

Private Sub mi_LugAccepterSettle_Click()
    On Error GoTo Error_Handle
    Dim lHelpContextID As Long
    lHelpContextID = frmAccepterSettle.HelpContextID
    
    frmAccepterSettle.Show vbModal, Me
    If frmAccepterSettle.m_bOk Then
        Dim rsSellDetail As Recordset
        Dim rsDetailToShow As Recordset
        Dim adbOther() As Double
        Dim oCalculator As New LugDataStatSvr
        Dim i As Integer, nUserCount As Integer
        
        Dim szLastTicketID As String
        Dim szBeginTicketID As String
        Dim arsData() As Recordset, vaCostumData As Variant
        
'        Dim lFullnumber As Long, lHalfnumber As Long, lFreenumber As Long
'        Dim dbFullAmount As Double, dbHalfAmount As Double, dbFreeAmount As Double
        Dim alNumber As Long '各种票种的张数
'        Dim adbAmount As Double  '各种票种的金额
        Dim j As Integer
        Dim aszAllSeller() As String
        Dim nAllSeller As Integer
        Dim k As Integer
        'Dim l As Integer
        Dim adbPriceItem() As Double '票价项明细

        Dim nTicketNumberLen As Integer
        Dim nTicketPrefixLen As Integer
        nTicketNumberLen = g_oParam.LuggageIDNumberLen
        nTicketPrefixLen = g_oParam.LuggageIDPrefixLen
        
        oCalculator.Init m_oAUser
        

        
        nUserCount = ArrayLength(frmAccepterSettle.m_vaSeller)
        
            
            ReDim arsData(1 To nUserCount)
            ReDim vaCostumData(1 To nUserCount, 1 To 22, 1 To 2)
'            SetProgressRange nUserCount, "正在形成记录集..."
            
            For i = 1 To nUserCount
'                    For j = 1 To TP_TicketTypeCount
                    alNumber = 0
'                    adbAmount = 0
'                    Next j
                
                    Set rsSellDetail = oCalculator.AcceptEveryDaySellDetail(ResolveDisplay(frmAccepterSettle.m_vaSeller(i)), frmAccepterSettle.m_dtWorkDate, frmAccepterSettle.m_dtEndDate)
                    Set rsDetailToShow = New Recordset
                    With rsDetailToShow.Fields
                        .Append "ticket_id_range", adChar, 30
                        '往记录集中添加每种票种的数量与金额字段
                    
                        .Append "number_ticket", adInteger
                        .Append "amount_ticket", adCurrency
                        
                    End With
                    
                    rsDetailToShow.Open

                    
                    Do While Not rsSellDetail.EOF
                        If rsDetailToShow.RecordCount = 0 Or Not IsTicketIDSequence(szLastTicketID, RTrim(rsSellDetail!luggage_id), nTicketNumberLen, nTicketPrefixLen) Then
                            If rsDetailToShow.RecordCount <> 0 Then
                                rsDetailToShow!ticket_id_range = szBeginTicketID & "---" & szLastTicketID
                                
                            
                                alNumber = alNumber + rsDetailToShow("number_ticket")
'                                adbAmount = adbAmount + rsDetailToShow("amount_ticket")
                                
                            End If
    
                            szBeginTicketID = RTrim(rsSellDetail!luggage_id)
                            rsDetailToShow.AddNew
                        End If
                        rsDetailToShow("number_ticket") = rsDetailToShow("number_ticket") + 1
                        rsDetailToShow("amount_ticket") = rsDetailToShow("amount_ticket") + rsSellDetail!price_total
                        
                        szLastTicketID = RTrim(rsSellDetail!luggage_id)
                        
                        rsSellDetail.MoveNext
                    Loop
                    
                    If rsSellDetail.RecordCount > 0 Then
                        rsSellDetail.MoveLast
                        rsDetailToShow!ticket_id_range = szBeginTicketID & "---" & RTrim(rsSellDetail!luggage_id)
'                        For j = 1 To TP_TicketTypeCount
                        alNumber = alNumber + rsDetailToShow("number_ticket")
'                        adbAmount = adbAmount + rsDetailToShow("amount_ticket")
'                        Next j
    
'                        rsDetailToShow.AddNew
                        
'                        rsDetailToShow!ticket_id_range = "合计"
'                        For j = 1 To TP_TicketTypeCount
'                        rsDetailToShow("number_ticket") = alNumber
'                        rsDetailToShow("amount_ticket") = adbAmount
'                        Next j
'                        rsDetailToShow.Update
                    End If
                    vaCostumData(i, 22, 1) = "票号段"
                    
                    If rsDetailToShow.RecordCount > 0 Then rsDetailToShow.MoveFirst
                    For j = 1 To rsDetailToShow.RecordCount
                        vaCostumData(i, 22, 2) = vaCostumData(i, 22, 2) & rsDetailToShow!ticket_id_range & "   "
                        rsDetailToShow.MoveNext
                    Next j
                    
                    Set arsData(i) = rsDetailToShow
                    adbOther = oCalculator.AcceptEveryDayAnotherThing(ResolveDisplay(frmAccepterSettle.m_vaSeller(i)), frmAccepterSettle.m_dtWorkDate, frmAccepterSettle.m_dtEndDate)
                    vaCostumData(i, 1, 1) = "废单数"
                    vaCostumData(i, 1, 2) = CInt(adbOther(1, 1)) & " 张  票款=" & adbOther(1, 2) & " 元"
                    
                    vaCostumData(i, 2, 1) = "退单数"
                    vaCostumData(i, 2, 2) = CInt(adbOther(2, 1)) & " 张  票款=" & adbOther(2, 2) & " 元  手续费=" & adbOther(2, 3) & " 元"
                    

                    Dim dbAmount As Double '不包括免票
                    Dim lNumber As Long '包括免票
                    dbAmount = oCalculator.AcceptEveryDaySellTotal(ResolveDisplay(frmAccepterSettle.m_vaSeller(i)), frmAccepterSettle.m_dtWorkDate, frmAccepterSettle.m_dtEndDate)

                    lNumber = alNumber
                        
                    vaCostumData(i, 4, 1) = "应交款"
                    vaCostumData(i, 4, 2) = dbAmount - adbOther(1, 2) - adbOther(2, 2) + adbOther(2, 3) & " 元"
                    
                    vaCostumData(i, 5, 1) = "受理单数"
                    'vaCostumData(i, 5, 2) = lNumber & " 张"
                    vaCostumData(i, 5, 2) = lNumber + adbOther(1, 1) + adbOther(2, 1) & " 张"
                    
                    vaCostumData(i, 6, 1) = "正常票单数"
                    'vaCostumData(i, 6, 2) = lNumber - adbOther(1, 1) - adbOther(2, 1) & " 张"
                    vaCostumData(i, 6, 2) = lNumber & " 张"

                    vaCostumData(i, 7, 1) = "制单"
                    vaCostumData(i, 7, 2) = MakeDisplayString(m_oAUser.UserID, m_oAUser.UserName)
                    
                    vaCostumData(i, 8, 1) = "复核"
                    vaCostumData(i, 8, 2) = ""
                    
                    vaCostumData(i, 9, 1) = "受理员"
                    vaCostumData(i, 9, 2) = frmAccepterSettle.m_vaSeller(i)
                    
                    vaCostumData(i, 10, 1) = "结算日期"
                    vaCostumData(i, 10, 2) = Format(frmAccepterSettle.m_dtWorkDate, "MM月DD日 hh:mm") & "—" & Format(frmAccepterSettle.m_dtEndDate, "MM月DD日 hh:mm")
                    
                    
                    adbPriceItem = oCalculator.GetAccepterPriceDetail(ResolveDisplay(frmAccepterSettle.m_vaSeller(i)), frmAccepterSettle.m_dtWorkDate, frmAccepterSettle.m_dtEndDate)
                    Dim nPrice As Integer
                    
                    For j = 1 To 10
                        vaCostumData(i, 10 + j, 1) = "price_item_" & j
                        vaCostumData(i, 10 + j, 2) = adbPriceItem(j)
                    Next j
                    vaCostumData(i, 21, 1) = "ticket_price_total"
                    vaCostumData(i, 21, 2) = adbPriceItem(j)
                    
'                    For j = 1 To arsData(i).RecordCount
'
'                    Next j
                    
'                End If
'                SetProgressValue i
                
                
                
            Next
            
            Dim frmNewReport As New frmReport
'            Dim frmTemp As IConditionForm
'            Set frmTemp = frmAccepterSettle
'            frmNewReport.m_lHelpContextID = lHelpContextID
            frmNewReport.ShowReport2 arsData, cszAcceptSettle, "行包员每日结算", vaCostumData
        End If
'    End If
    Exit Sub
Error_Handle:
'    SetProgressVisible False
    ShowErrorMsg
    

End Sub



Private Sub mnu_LugSalerDayTotal_Click()
    On Error GoTo Error_Handle
    Dim lHelpContextID As Long
    lHelpContextID = frmLugSalerDayTotal.HelpContextID
    
    frmLugSalerDayTotal.Show vbModal, Me
    If frmLugSalerDayTotal.m_bOk Then
        
        Dim rsSellDetail As Recordset
        Dim rsDetailToShow As Recordset
        Dim adbOther() As Double
        Dim oDss As New LuggageSvr
        oDss.Init m_oAUser
        Dim i As Integer, nUserCount As Integer
        
        Dim szLastTicketID As String
        Dim szBeginTicketID As String
        Dim arsData() As Recordset, vaCostumData As Variant
        
        Dim j As Integer
        Dim aszAllSeller() As String
        Dim nAllSeller As Integer
        Dim k As Integer
        Dim l As Integer
        
        
        Dim oUnit As New Unit
        oUnit.Init m_oAUser
        oUnit.Identify g_oParam.UnitID
        aszAllSeller = oUnit.GetAllUserEX(, ResolveDisplay(frmLugSalerDayTotal.cboSellStation))
        nAllSeller = ArrayLength(aszAllSeller)
        
        
        nUserCount = ArrayLength(frmLugSalerDayTotal.m_vaSeller)
        
        If nAllSeller > 0 Then
            
            ReDim arsData(1 To IIf(nAllSeller > nUserCount, nUserCount, nAllSeller))
            ReDim vaCostumData(1 To IIf(nAllSeller > nUserCount, nUserCount, nAllSeller), 1 To 7, 1 To 2)
            WriteProcessBar True, , nUserCount, "正在形成记录集..."
            l = 0
            For i = 1 To nUserCount
                WriteProcessBar , i, nUserCount, "正在得到" & EncodeString(frmLugSalerDayTotal.m_vaSeller(i)) & "的数据..."
                For k = 1 To nAllSeller
                    If LCase(Trim(ResolveDisplay(frmLugSalerDayTotal.m_vaSeller(i)))) = LCase(aszAllSeller(k, 1)) Then
                        Exit For
                    End If
                Next k
                If k <= nAllSeller Then
                    l = l + 1
                    '初始化
                    Dim aszSelectUser(1 To 1) As String
                    aszSelectUser(1) = ResolveDisplay(frmLugSalerDayTotal.m_vaSeller(i))
                    Set rsSellDetail = oDss.TotalAcceptRS(aszSelectUser, frmLugSalerDayTotal.m_dtWorkDate, frmLugSalerDayTotal.m_dtEndDate)
                    Set arsData(l) = rsSellDetail   '记录集赋值
                    
                    '构建自义项目集
                    Dim dbCancelPrice As Double, dbReturnPrice As Double, dbReturnCharge As Double
                    Dim dbNeedPayMoney As Double
                    dbCancelPrice = 0: dbReturnCharge = 0: dbReturnPrice = 0: dbNeedPayMoney = 0
                    While Not rsSellDetail.EOF
                        dbCancelPrice = dbCancelPrice + FormatDbValue(rsSellDetail!cancel_price)
                        dbReturnPrice = dbReturnPrice + FormatDbValue(rsSellDetail!return_price)
                        dbReturnCharge = dbReturnCharge + FormatDbValue(rsSellDetail!return_charge)
                        dbNeedPayMoney = dbNeedPayMoney + FormatDbValue(rsSellDetail!price_total) - FormatDbValue(rsSellDetail!cancel_price) - FormatDbValue(rsSellDetail!return_price) + FormatDbValue(rsSellDetail!return_charge)
                        rsSellDetail.MoveNext
                    Wend
                    
                    
                    '创建自定义项目集
                    vaCostumData(l, 1, 1) = "统计开始时间"
                    vaCostumData(l, 1, 2) = Format(frmLugSalerDayTotal.m_dtWorkDate, "MM月DD日 HH:mm")
                    vaCostumData(l, 2, 1) = "统计结束时间"
                    vaCostumData(l, 2, 2) = Format(frmLugSalerDayTotal.m_dtEndDate, "MM月DD日 HH:mm")
                    vaCostumData(l, 3, 1) = "作废款"
                    vaCostumData(l, 3, 2) = dbCancelPrice & "元"
                    vaCostumData(l, 4, 1) = "退运款"
                    vaCostumData(l, 4, 2) = dbReturnPrice & "元"
                    vaCostumData(l, 5, 1) = "退运手续费"
                    vaCostumData(l, 5, 2) = dbReturnCharge & "元"
                    vaCostumData(l, 6, 1) = "应交款"
                    vaCostumData(l, 6, 2) = dbNeedPayMoney & "元"
                    vaCostumData(l, 7, 1) = "行包员"
                    vaCostumData(l, 7, 2) = frmLugSalerDayTotal.m_vaSeller(i)
                    
                End If
            Next
            WriteProcessBar False, , , ""
            
            Dim frmNewReport As New frmReport
'            frmNewReport.Show
            Dim frmTemp As IConditionForm
            Set frmTemp = frmLugSalerDayTotal
            frmNewReport.m_lHelpContextID = lHelpContextID
            frmNewReport.ShowReport2 arsData, frmTemp.FileName, cszLugSalerDayTotal, vaCostumData, 10
            
            WriteProcessBar False, , , ""
        End If
    End If
    Exit Sub
Error_Handle:
    WriteProcessBar False, , , ""
    ShowErrorMsg
End Sub

'行包营收拆算一览表
Private Sub mnu_FinSheetTotalList_Click()
  On Error GoTo ErrHandle
    Dim lHelpContextID As Long
    Dim rsSpliteDetail As Recordset
    Dim i As Integer
    Dim mStartDate As Date
    Dim mEndDate As Date
    Dim vaCostumData() As String
    
    lHelpContextID = frmLugSplitList.HelpContextID
    
    frmLugSplitList.Show vbModal, Me
     If frmLugSplitList.m_bOk Then

          WriteProcessBar True, , , "正在形成记录集..."
                  '初始化结算月份
                    mStartDate = CDate(Format(frmLugSplitList.m_dtStartDate, "yyyy-mm") & " -1 00:00:01")
                    Select Case Month(frmLugSplitList.m_dtEndDate)
                           Case 1, 3, 5, 7, 8, 10, 12
                            mEndDate = CDate(Format(frmLugSplitList.m_dtEndDate, "yyyy-mm") & "-31" & " 23:59:59")
                           Case 4, 6, 9, 11
                            mEndDate = CDate(Format(frmLugSplitList.m_dtEndDate, "yyyy-mm") & "-30" & " 23:59:59")
                           Case 2
                            mEndDate = CDate(Format(frmLugSplitList.m_dtEndDate, "yyyy-mm") & "-28" & " 23:59:59")
                    End Select
                    m_oLugSplitSvr.Init m_oAUser
                    If frmLugSplitList.cboAcceptType.Text = "" Then
                        Set rsSpliteDetail = m_oLugSplitSvr.LugFinanceStat(mStartDate, mEndDate, mSplitCompanyID, ResolveDisplay(Trim(frmLugSplitList.cboSellStation.Text)))
                    Else
                    Set rsSpliteDetail = m_oLugSplitSvr.LugFinanceStat(mStartDate, mEndDate, mSplitCompanyID, ResolveDisplay(Trim(frmLugSplitList.cboSellStation.Text)), GetLuggageTypeInt(frmLugSplitList.cboAcceptType.Text))
                    End If
                   
                    '创建自定义项目集
                    ReDim vaCostumData(1 To 3, 1 To 2)
                    vaCostumData(1, 1) = "结算月份"
                    vaCostumData(1, 2) = IIf(Format(frmLugSplitList.m_dtStartDate, "YYYY年MM月DD日") = Format(frmLugSplitList.m_dtEndDate, "YYYY年MM月DD日"), Format(frmLugSplitList.m_dtStartDate, "YYYY年MM月DD日"), Format(frmLugSplitList.m_dtStartDate, "YYYY年MM月DD日") & " - " & Format(frmLugSplitList.m_dtEndDate, "YYYY年MM月DD日"))
                    vaCostumData(2, 1) = "制表人"
                    vaCostumData(2, 2) = Trim(m_oAUser.UserName)
                    vaCostumData(3, 1) = "制表日期"
                    vaCostumData(3, 2) = Date


            WriteProcessBar False, , , ""
            WriteProcessBar True, , , "正在形成报表..."
            Dim frmNewReport As New frmReport
            frmNewReport.ShowReport rsSpliteDetail, cszLugSplitList, frmLugSplitList.Caption, vaCostumData, 10
            WriteProcessBar False, , , ""
        End If
    Set rsSpliteDetail = Nothing
 Exit Sub
ErrHandle:
ShowErrorMsg
End Sub

'行包营收拆帐明细报表
Private Sub mnu_CompanyFinSheetStat_Click()
  On Error GoTo ErrHandle
    Dim lHelpContextID As Long
    Dim rsSpliteDetail As Recordset
    Dim i As Integer
    Dim mStartDate As Date
    Dim mEndDate As Date
    Dim vaCostumData() As String
    Dim sCompany As String
    lHelpContextID = frmLugCompanySplite.HelpContextID
    
    frmLugCompanySplite.Show vbModal, Me
    If frmLugCompanySplite.m_bOk Then

          WriteProcessBar True, , , "正在形成记录集..."
                  '初始化结算月份
                    mStartDate = CDate(Format(frmLugCompanySplite.m_dtStartDate, "yyyy-MM") & " -1 00:00:01")
                    Select Case Month(frmLugCompanySplite.m_dtEndDate)
                           Case 1, 3, 5, 7, 8, 10, 12
                            mEndDate = CDate(Format(frmLugCompanySplite.m_dtEndDate, "yyyy-MM") & "-31" & " 23:59:59")
                           Case 4, 6, 9, 11
                            mEndDate = CDate(Format(frmLugCompanySplite.m_dtEndDate, "yyyy-MM") & "-30" & " 23:59:59")
                           Case 2
                            mEndDate = CDate(Format(frmLugCompanySplite.m_dtEndDate, "yyyy-MM") & "-28" & " 23:59:59")
                    End Select
                    m_oLugSplitSvr.Init m_oAUser
                    If frmLugSplitList.cboAcceptType.Text = "" Then
                        Set rsSpliteDetail = m_oLugSplitSvr.LugFinanceDetailStat(mStartDate, mEndDate, mSplitVehicleID)
                    Else
                        Set rsSpliteDetail = m_oLugSplitSvr.LugFinanceDetailStat(mStartDate, mEndDate, mSplitVehicleID, , GetLuggageTypeInt(frmLugCompanySplite.cboAcceptType.Text))
                    End If
              
                    '创建自定义项目集
                    ReDim vaCostumData(1 To 4, 1 To 2)
                    vaCostumData(1, 1) = "参运公司"
                    vaCostumData(1, 2) = frmLugCompanySplite.m_Company
                    vaCostumData(2, 1) = "结算月份"
                    vaCostumData(2, 2) = IIf(Format(frmLugCompanySplite.m_dtStartDate, "YYYY年MM月DD日") = Format(frmLugCompanySplite.m_dtEndDate, "YYYY年MM月DD日"), Format(frmLugCompanySplite.m_dtEndDate, "YYYY年MM月DD日"), Format(frmLugCompanySplite.m_dtStartDate, "YYYY年MM月DD日") & " - " & Format(frmLugCompanySplite.m_dtEndDate, "YYYY年MM月DD日"))
                    vaCostumData(3, 1) = "制表人"
                    vaCostumData(3, 2) = Trim(m_oAUser.UserName)
                    vaCostumData(4, 1) = "制表日期"
                    vaCostumData(4, 2) = Date


            WriteProcessBar False, , , ""
            WriteProcessBar True, , , "正在形成报表..."
            Dim frmNewReport As New frmReport
            frmNewReport.ShowReport rsSpliteDetail, cszLugCompanySplit, frmLugCompanySplite.Caption, vaCostumData, 10
            WriteProcessBar False, , , ""
        End If
    Set rsSpliteDetail = Nothing
 Exit Sub
ErrHandle:
ShowErrorMsg
End Sub
'激活对应的工具栏
Public Sub ActiveToolBar(pbTrue As Boolean)

            abMenu.Bands("mnu_System").Tools("mnu_BaseInfo").Enabled = pbTrue
            abMenu.Bands("mnu_System").Tools("mnu_ChgPassword").Enabled = pbTrue
            abMenu.Bands("mnu_System").Tools("mnu_OptionSet").Enabled = pbTrue
    
            
'        abMenu.Bands("mnu_System").Tools("mnu_ExportFileOpen").Enabled = pbTrue
End Sub
'调用站点行包营收窗口

Private Sub mnu_LugStation_Click()
    On Error GoTo Error_Handle
    Dim lHelpContextID As Long
        lHelpContextID = frmStationDayStat.HelpContextID
        frmStationDayStat.Show vbModal
  If frmStationDayStat.m_bOk Then
        Dim rsSellDetail As Recordset
      
        Dim oLugDss As New LugDataStatSvr
        oLugDss.Init m_oAUser
        Dim cszFileName As String
         
        Dim vaCostumData As Variant
        '创建自定义项目集tt
        ReDim vaCostumData(1 To 3, 1 To 2)
        vaCostumData(1, 1) = "结算月份"
        vaCostumData(1, 2) = IIf(Format(frmStationDayStat.m_dtWorkDate, "YYYY年MM月DD日") = Format(frmStationDayStat.m_dtEndDate, "YYYY年MM月DD日"), Format(frmStationDayStat.m_dtWorkDate, "YYYY年MM月DD日"), Format(frmStationDayStat.m_dtWorkDate, "YYYY年MM月DD日") & " - " & Format(frmStationDayStat.m_dtEndDate, "YYYY年MM月DD日"))
        vaCostumData(2, 1) = "制表人"
        vaCostumData(2, 2) = Trim(m_oAUser.UserName)
        vaCostumData(3, 1) = "制表日期"
        vaCostumData(3, 2) = Date
                
                Select Case szFromId
                        Case 4
                            cszStation = "车站行包营收日报"
                             cszFileName = "车站行包营收日报.xls"
                            Set rsSellDetail = oLugDss.StationDayDataStat(ResolveDisplay(frmStationDayStat.m_SellStation), frmStationDayStat.m_dtWorkDate, frmStationDayStat.m_dtEndDate, frmStationDayStat.m_szStation, frmStationDayStat.m_szAcceptType)
                          
                        Case 5
                            cszStation = "车站行包营收月报"
                             cszFileName = "车站行包营收月报.xls"
                            Set rsSellDetail = oLugDss.StationMonthDataStat(ResolveDisplay(frmStationDayStat.m_SellStation), frmStationDayStat.m_dtWorkDate, frmStationDayStat.m_dtEndDate, frmStationDayStat.m_szStation, frmStationDayStat.m_szAcceptType)
                            
                        Case 6
                            cszStation = "车站行包营收年报"
                            cszFileName = "车站行包营收年报.xls"
                            Set rsSellDetail = oLugDss.StationYearDataStat(ResolveDisplay(frmStationDayStat.m_SellStation), frmStationDayStat.m_dtWorkDate, frmStationDayStat.m_dtEndDate, frmStationDayStat.m_szStation, frmStationDayStat.m_szAcceptType)
                         Case 16
                            cszStation = "车站行包营收简报"
                            cszFileName = "车站行包营收简报.xls"
                            Set rsSellDetail = oLugDss.StationStat(ResolveDisplay(frmStationDayStat.m_SellStation), frmStationDayStat.m_dtWorkDate, frmStationDayStat.m_dtEndDate, frmStationDayStat.m_szStation, frmStationDayStat.m_szAcceptType)
                        
                End Select
               WriteProcessBar False, , , ""
            
            Dim frmNewReport As New frmReport
'            frmNewReport.Show
'            Dim frmTemp As IConditionForm
'            Set frmTemp = frmStationDayStat
'            frmNewReport.m_lHelpContextID = lHelpContextID
            
           frmNewReport.ShowReport rsSellDetail, cszFileName, cszStation, vaCostumData
                     
         End If
Exit Sub
Error_Handle:
  
    ShowErrorMsg
End Sub
'调用站点行包营收窗口
Private Sub mnu_LugSellStation_Click()
    On Error GoTo Error_Handle
    Dim lHelpContextID As Long
    lHelpContextID = frmSellStationDayStat.HelpContextID
    frmSellStationDayStat.Show vbModal
  If frmSellStationDayStat.m_bOk Then
        Dim rsSellDetail As Recordset
        Dim frmNewReport As New frmReport
        Dim oLugSellDss As New LugDataStatSvr
        oLugSellDss.Init m_oAUser
        Dim cszFileName As String
        Dim rsDate() As Recordset
         
        Dim vaCostumData As Variant
        '创建自定义项目集tt
        ReDim vaCostumData(1 To 3, 1 To 2)
        vaCostumData(1, 1) = "结算月份"
        vaCostumData(1, 2) = IIf(Format(frmSellStationDayStat.m_dtWorkDate, "YYYY年MM月DD日") = Format(frmSellStationDayStat.m_dtEndDate, "YYYY年MM月DD日"), Format(frmSellStationDayStat.m_dtWorkDate, "YYYY年MM月DD日"), Format(frmSellStationDayStat.m_dtWorkDate, "YYYY年MM月DD日") & " - " & Format(frmStationDayStat.m_dtEndDate, "YYYY年MM月DD日"))
        vaCostumData(2, 1) = "制表人"
        vaCostumData(2, 2) = Trim(m_oAUser.UserName)
        vaCostumData(3, 1) = "制表日期"
        vaCostumData(3, 2) = Date
                
                Select Case szFromId
                        Case 6
                            cszSellstation = "售票站行包营收日报"
                             cszFileName = "售票站行包营收日报.xls"
                            Set rsSellDetail = oLugSellDss.SellStationDayDataStat(ResolveDisplay(frmSellStationDayStat.m_SellStation), frmSellStationDayStat.m_dtWorkDate, frmSellStationDayStat.m_dtEndDate, frmSellStationDayStat.m_szAcceptType)
                          
                        Case 7
                            cszSellstation = "售票站行包营收月报"
                             cszFileName = "售票站行包营收月报.xls"
                            Set rsSellDetail = oLugSellDss.SellStationMonthDataStat(ResolveDisplay(frmSellStationDayStat.m_SellStation), frmSellStationDayStat.m_dtWorkDate, frmSellStationDayStat.m_dtEndDate, frmSellStationDayStat.m_szAcceptType)
                            
                        Case 8
                            cszSellstation = "售票站行包营收年报"
                            cszFileName = "售票站行包营收年报.xls"
                            Set rsSellDetail = oLugSellDss.SellStationYearDataStat(ResolveDisplay(frmSellStationDayStat.m_SellStation), frmSellStationDayStat.m_dtWorkDate, frmSellStationDayStat.m_dtEndDate, frmSellStationDayStat.m_szAcceptType)
                        Case 17
                            cszSellstation = "售票站行包营收简报"
                            cszFileName = "售票站行包营收简报.xls"
                            Set rsSellDetail = oLugSellDss.SellStationStat(ResolveDisplay(frmSellStationDayStat.m_SellStation), frmSellStationDayStat.m_dtWorkDate, frmSellStationDayStat.m_dtEndDate, frmSellStationDayStat.m_szAcceptType)
                                        
                End Select
               WriteProcessBar False, , , ""
            
        
'            frmNewReport.Show
'            Dim frmTemp As IConditionForm
'            Set frmTemp = frmSellStationDayStat
'            frmNewReport.m_lHelpContextID = lHelpContextID
            
           frmNewReport.ShowReport3 rsSellDetail, cszFileName, cszSellstation, vaCostumData
                     
         End If
Exit Sub
Error_Handle:
  
    ShowErrorMsg
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


Private Sub mi_SplitCompanyCheckStat_Click()
    Dim lHelpContextID As Long
    frmCheckStat.m_nStatType = UI_SplitCompanyCheckStat
    lHelpContextID = frmCheckStat.HelpContextID
    frmCheckStat.Show vbModal, Me
    If frmCheckStat.m_bOk Then
        
        Dim frmTemp As IConditionForm
        Dim frmNewReport As New frmReport
        Set frmTemp = frmCheckStat
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport frmTemp.RecordsetData, frmTemp.FileName, "", frmTemp.CustomData
        Unload frmCheckStat
    End If
End Sub


Private Sub mi_VehicleCheckStat_Click()
    Dim lHelpContextID As Long
    frmCheckStat.m_nStatType = UI_VehicleCheckStat
    lHelpContextID = frmCheckStat.HelpContextID
    frmCheckStat.Show vbModal, Me
    If frmCheckStat.m_bOk Then
        Dim frmTemp As IConditionForm
        Dim frmNewReport As New frmReport
        Set frmTemp = frmCheckStat
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport frmTemp.RecordsetData, frmTemp.FileName, "", frmTemp.CustomData
        Unload frmCheckStat
    End If
End Sub


Private Sub mi_RouteCheckStat_Click()
    Dim lHelpContextID As Long
    frmCheckStat.m_nStatType = UI_RouteCheckStat
    lHelpContextID = frmCheckStat.HelpContextID
    frmCheckStat.Show vbModal, Me
    If frmCheckStat.m_bOk Then
        Dim frmTemp As IConditionForm
        Dim frmNewReport As New frmReport
        Set frmTemp = frmCheckStat
        frmNewReport.m_lHelpContextID = lHelpContextID
        frmNewReport.ShowReport frmTemp.RecordsetData, frmTemp.FileName, "", frmTemp.CustomData
        Unload frmCheckStat
    End If
End Sub

