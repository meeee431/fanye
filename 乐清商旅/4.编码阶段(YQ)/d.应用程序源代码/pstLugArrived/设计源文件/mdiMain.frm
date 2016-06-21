VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "actbar2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.MDIForm mdiMain 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000C&
   Caption         =   "行包到达"
   ClientHeight    =   8175
   ClientLeft      =   1275
   ClientTop       =   2055
   ClientWidth     =   11235
   HelpContextID   =   7000001
   Icon            =   "mdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   WindowState     =   2  'Maximized
   Begin ActiveBar2LibraryCtl.ActiveBar2 abMenu 
      Align           =   1  'Align Top
      Height          =   8175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11235
      _LayoutVersion  =   1
      _ExtentX        =   19817
      _ExtentY        =   14420
      _DataPath       =   ""
      Bands           =   "mdiMain.frx":16AC2
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   480
         Top             =   2160
      End
      Begin MSComctlLib.ProgressBar pbLoad 
         Height          =   225
         Left            =   4920
         TabIndex        =   6
         Top             =   7020
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.PictureBox ptTitleTop 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   30
         ScaleHeight     =   450
         ScaleWidth      =   15360
         TabIndex        =   1
         Top             =   1140
         Width           =   15360
         Begin RTComctl3.CoolButton cmdClose 
            Height          =   390
            Left            =   11670
            TabIndex        =   2
            ToolTipText     =   "返回"
            Top             =   0
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
            MICON           =   "mdiMain.frx":1DA2A
            PICN            =   "mdiMain.frx":1DA46
            PICH            =   "mdiMain.frx":1E93B
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   1
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label lblSheetNo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   9960
            TabIndex        =   5
            Top             =   90
            Width           =   165
         End
         Begin VB.Label lblSheetNoName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "当前单据号:"
            Height          =   180
            Left            =   8940
            TabIndex        =   4
            Top             =   150
            Width           =   990
         End
         Begin VB.Label fblCurrentTime 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0:00:00"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   7140
            TabIndex        =   3
            Top             =   60
            Width           =   1185
         End
      End
      Begin VB.Label lblInStation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "途经站"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   540
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub abMenu_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
On Error GoTo ErrHandle
    Select Case Tool.name
        Case "mi_PackageMan"        '到达管理
            frmAllPackage.ZOrder 0
            frmAllPackage.Show
        Case "mi_Accept"        '正常受理
            AcceptPackage
        Case "mi_Reprint", "mi_ReprintSheet"
        
            ReprintSheet
        Case "mi_SheetNo"       '更改单据号
            RefreshNO
'            frmChgSheetNo.Show vbModal
        Case "mi_Param"
            frmSysParam.Show vbModal
        Case "mi_BaseCode"
            frmBaseDefine.ZOrder 0
            frmBaseDefine.Show
        Case "mi_ChgPassword"
            ChangePassword
        Case "mi_SysExit"
            Unload Me
        Case "mnu_HelpIndex"
            DisplayHelp Me, Index
        Case "mnu_HelpContent"
            If Not ActiveForm Is Nothing Then
                DisplayHelp ActiveForm, content
            End If
        Case "mnu_About"
            AboutMe

        '以下为统计部分

        Case "mi_LugAccepterSettle"
            mi_LugAccepterSettle_Click
        Case "mi_StatPackage"   '行单统计结算
            frmQueryAccept.ZOrder 0
            frmQueryAccept.Show
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
    End Select
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub AboutMe()
Dim oShell As New CommShell
    oShell.ShowAbout App.ProductName, "Package Manage", App.FileDescription, Me.Icon, App.Major, App.Minor, App.Revision
End Sub
Private Sub ReprintSheet()
    frmRePrintSheet.Show vbModal
End Sub

Private Sub ChangePassword()

 Dim oShell As New CommDialog
    On Error GoTo ErrorHandle
    oShell.init g_oActUser
    oShell.ShowUserInfo
    Set oShell = Nothing
    Exit Sub
ErrorHandle:
    Set oShell = Nothing
    ShowErrorMsg
End Sub
'刷新修改后的单受理单及签发单号
Private Sub RefreshNO()
    frmChgSheetNo.m_bNoCancel = False
    frmChgSheetNo.Show vbModal, Me
    If frmChgSheetNo.m_bOk Then
        lblSheetNo.Caption = g_szSheetID
        If ActiveForm.name = "frmArrived" Then frmArrived.lblSheetID.Caption = g_szSheetID
    End If
End Sub
'正常受理
Private Sub AcceptPackage()
    frmArrived.Status = EFS_AddNew
    frmArrived.ZOrder 0
    frmArrived.Show

End Sub

'关联ActiveBar的控件
Private Sub AddControlsToActBar()
    abMenu.Bands("bndTitleTop").Tools("tblTitleTop").Custom = ptTitleTop
    abMenu.Bands("statusBar").Tools("progressBar").Custom = pbLoad
End Sub

Private Sub cmdClose_Click()
On Error Resume Next
    If Not ActiveForm Is Nothing Then
        Unload ActiveForm
    End If
End Sub

Private Sub MDIForm_Load()
    AddControlsToActBar

    SetPrintEnabled False
    lblSheetNo.Caption = g_szSheetID

    '初始化主界面，如状态条等
    frmArrived.Status = EFS_AddNew
    frmArrived.RefreshForm
    frmArrived.Show     '缺省打开受理窗体
End Sub

Private Sub Timer1_Timer()
    fblCurrentTime.Caption = Format(Time, "HH:mm:ss")
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

Private Sub mi_LugAccepterSettle_Click()
    On Error GoTo Error_Handle
    Dim lHelpContextID As Long
    lHelpContextID = frmAccepterSettle.HelpContextID

    frmAccepterSettle.Show vbModal, Me
    If frmAccepterSettle.m_bOk Then
        Dim rsSellDetail As Recordset
        Dim rsDetailToShow As Recordset
        Dim adbOther() As Double
        Dim oCalculator As New PackageSvr
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
        nTicketNumberLen = g_oPackageParam.SheetIDNumberLen
        nTicketPrefixLen = 0

        oCalculator.init g_oActUser



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
                        If rsDetailToShow.RecordCount = 0 Or Not IsTicketIDSequence(szLastTicketID, RTrim(rsSellDetail!sheet_id), nTicketNumberLen, nTicketPrefixLen) Then
                            If rsDetailToShow.RecordCount <> 0 Then
                                rsDetailToShow!ticket_id_range = szBeginTicketID & "---" & szLastTicketID


                                alNumber = alNumber + rsDetailToShow("number_ticket")
'                                adbAmount = adbAmount + rsDetailToShow("amount_ticket")

                            End If

                            szBeginTicketID = RTrim(rsSellDetail!sheet_id)
                            rsDetailToShow.AddNew
                        End If
                        rsDetailToShow("number_ticket") = rsDetailToShow("number_ticket") + 1
                        rsDetailToShow("amount_ticket") = rsDetailToShow("amount_ticket") + rsSellDetail!price_total

                        szLastTicketID = RTrim(rsSellDetail!sheet_id)

                        rsSellDetail.MoveNext
                    Loop

                    If rsSellDetail.RecordCount > 0 Then
                        rsSellDetail.MoveLast
                        rsDetailToShow!ticket_id_range = szBeginTicketID & "---" & RTrim(rsSellDetail!sheet_id)
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

                    'zyw 2006-11-10 当日退提
                    vaCostumData(i, 2, 1) = "退提数"
                    vaCostumData(i, 2, 2) = CInt(adbOther(2, 1)) & " 张  票款=" & adbOther(2, 2) & " 元"

                    Dim dbAmount As Double '不包括免票
                    Dim lNumber As Long '包括免票
                    dbAmount = oCalculator.AcceptEveryDaySellTotal(ResolveDisplay(frmAccepterSettle.m_vaSeller(i)), frmAccepterSettle.m_dtWorkDate, frmAccepterSettle.m_dtEndDate)

                    lNumber = alNumber

                    vaCostumData(i, 4, 1) = "应交款"
                    vaCostumData(i, 4, 2) = dbAmount - adbOther(1, 2) - adbOther(2, 2) - adbOther(4, 2) & " 元" '要加上隔日退提的

                    vaCostumData(i, 5, 1) = "受理单数"
                    vaCostumData(i, 5, 2) = lNumber & " 张"

                    vaCostumData(i, 6, 1) = "正常票单数"
                    vaCostumData(i, 6, 2) = lNumber - IIf(adbOther(1, 1) < adbOther(5, 1), adbOther(5, 1) - adbOther(1, 1), adbOther(1, 1) - adbOther(5, 1)) - adbOther(2, 1) & " 张"

                    vaCostumData(i, 7, 1) = "制单"
                    vaCostumData(i, 7, 2) = MakeDisplayString(g_oActUser.UserID, g_oActUser.userName)

                    vaCostumData(i, 8, 1) = "复核"
                    vaCostumData(i, 8, 2) = ""

                    vaCostumData(i, 9, 1) = "受理员"
                    vaCostumData(i, 9, 2) = frmAccepterSettle.m_vaSeller(i)

                    vaCostumData(i, 10, 1) = "结算日期"
                    vaCostumData(i, 10, 2) = Format(frmAccepterSettle.m_dtWorkDate, "MM月DD日 hh:mm") & "—" & Format(frmAccepterSettle.m_dtEndDate, "MM月DD日 hh:mm")
                    
                    


'                    adbPriceItem = oCalculator.GetAccepterPriceDetail(ResolveDisplay(frmAccepterSettle.m_vaSeller(i)), frmAccepterSettle.m_dtWorkDate, frmAccepterSettle.m_dtEndDate)
'                    Dim nPrice As Integer
'
'                    For j = 1 To 10
'                        vaCostumData(i, 10 + j, 1) = "price_item_" & j
'                        vaCostumData(i, 10 + j, 2) = adbPriceItem(j)
'                    Next j
'                    vaCostumData(i, 21, 1) = "ticket_price_total"
'                    vaCostumData(i, 21, 2) = adbPriceItem(j)

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
            frmNewReport.ShowReport2 arsData, "行包员每日结算_有票明细.xls", "行包员每日结算", vaCostumData
        End If
'    End If
    Exit Sub
Error_Handle:
'    SetProgressVisible False
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
