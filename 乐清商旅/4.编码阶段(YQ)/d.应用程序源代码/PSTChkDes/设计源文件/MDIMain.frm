VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "actbar2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{A0123751-4698-48C1-A06C-A2482B5ED508}#2.0#0"; "RTComctl2.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.MDIForm MDIMain 
   BackColor       =   &H8000000C&
   Caption         =   "检票系统"
   ClientHeight    =   8280
   ClientLeft      =   1185
   ClientTop       =   2760
   ClientWidth     =   10575
   HelpContextID   =   4000001
   Icon            =   "MDIMain.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin ActiveBar2LibraryCtl.ActiveBar2 abMenu 
      Align           =   1  'Align Top
      Height          =   8280
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10575
      _LayoutVersion  =   1
      _ExtentX        =   18653
      _ExtentY        =   14605
      _DataPath       =   ""
      Bands           =   "MDIMain.frx":16AC2
      Begin VB.PictureBox ptTitle 
         BackColor       =   &H00FFFFFF&
         Height          =   465
         Left            =   -30
         ScaleHeight     =   405
         ScaleWidth      =   15300
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   960
         Width           =   15360
         Begin MSComctlLib.TabStrip tbsBusList 
            Height          =   450
            Left            =   4230
            TabIndex        =   17
            Top             =   0
            Visible         =   0   'False
            Width           =   11055
            _ExtentX        =   19500
            _ExtentY        =   794
            TabWidthStyle   =   2
            MultiRow        =   -1  'True
            TabFixedWidth   =   2646
            TabFixedHeight  =   616
            TabMinWidth     =   1587
            TabStyle        =   1
            _Version        =   393216
            BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
               NumTabs         =   1
               BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  ImageVarType    =   2
               EndProperty
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.PictureBox ptNextBusInfo 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   6120
         Left            =   0
         ScaleHeight     =   6120
         ScaleWidth      =   2550
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   1380
         Width           =   2550
         Begin VB.Frame fraCheckGate 
            BackColor       =   &H00E0E0E0&
            Caption         =   "检票口信息"
            ForeColor       =   &H00C00000&
            Height          =   1515
            Left            =   60
            TabIndex        =   12
            Top             =   4350
            Width           =   2430
            Begin VB.Label lblCheckGate 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "一号"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   300
               Left            =   990
               TabIndex        =   26
               Top             =   300
               Width           =   1320
            End
            Begin VB.Label lblChecker 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "001/陆勇庆"
               ForeColor       =   &H80000008&
               Height          =   300
               Left            =   990
               TabIndex        =   25
               Top             =   660
               Width           =   1320
            End
            Begin VB.Label lblCurrentSheetNo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "111112222"
               ForeColor       =   &H80000008&
               Height          =   300
               Left            =   990
               TabIndex        =   24
               Top             =   1020
               Width           =   1320
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "检票口:"
               Height          =   180
               Left            =   150
               TabIndex        =   15
               Top             =   360
               Width           =   630
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "检票员:"
               Height          =   180
               Left            =   150
               TabIndex        =   14
               Top             =   660
               Width           =   630
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "当前路单号:"
               Height          =   345
               Left            =   150
               TabIndex        =   13
               Top             =   960
               Width           =   570
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00E0E0E0&
            Caption         =   "下一检票车次"
            ForeColor       =   &H00C00000&
            Height          =   3645
            Left            =   60
            TabIndex        =   2
            Top             =   60
            Width           =   2430
            Begin RTComctl2.RevTimer rvtTime 
               Height          =   360
               Left            =   975
               Top             =   2805
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   635
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BackColor       =   14737632
               OutnerStyle     =   2
               NormTextColor   =   12582912
               Enabled         =   0   'False
               Second          =   0
            End
            Begin RTComctl3.FlatLabel flblrevTime 
               Height          =   360
               Left            =   975
               TabIndex        =   3
               Top             =   2805
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   635
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BackColor       =   14737632
               OutnerStyle     =   2
               VerticalAlignment=   1
               MarginLeft      =   0
               MarginTop       =   0
               NormTextColor   =   12582912
               Enabled         =   0   'False
               Caption         =   "00:00:00"
            End
            Begin VB.Label lblEndStation 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "杭州"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   300
               Left            =   990
               TabIndex        =   23
               Top             =   990
               Width           =   1320
            End
            Begin VB.Label lblLicense 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "浙D0001"
               ForeColor       =   &H00C00000&
               Height          =   300
               Left            =   990
               TabIndex        =   22
               Top             =   1365
               Width           =   1320
            End
            Begin VB.Label lblOwner 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "陆勇庆"
               ForeColor       =   &H00C00000&
               Height          =   300
               Left            =   990
               TabIndex        =   21
               Top             =   1725
               Width           =   1320
            End
            Begin VB.Label lblCompany 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "绍兴长运"
               ForeColor       =   &H00C00000&
               Height          =   300
               Left            =   990
               TabIndex        =   20
               Top             =   2085
               Width           =   1320
            End
            Begin VB.Label lblStartupTime 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "10:00"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   300
               Left            =   990
               TabIndex        =   19
               Top             =   630
               Width           =   1320
            End
            Begin VB.Label lblBusID 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "11900"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   300
               Left            =   990
               TabIndex        =   18
               Top             =   270
               Width           =   1320
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "请等候..."
               Height          =   180
               Left            =   960
               TabIndex        =   11
               Top             =   3285
               Width           =   810
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "离开始检票还有"
               Height          =   180
               Left            =   960
               TabIndex        =   10
               Top             =   2535
               Width           =   1260
            End
            Begin VB.Image Image1 
               Height          =   480
               Left            =   195
               Picture         =   "MDIMain.frx":1C2E4
               Top             =   2595
               Width           =   480
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "车牌号:"
               Height          =   180
               Left            =   150
               TabIndex        =   9
               Top             =   1455
               Width           =   630
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "车主:"
               Height          =   180
               Left            =   150
               TabIndex        =   8
               Top             =   1815
               Width           =   450
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "参营公司:"
               Height          =   180
               Left            =   150
               TabIndex        =   7
               Top             =   2175
               Width           =   810
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "终到站:"
               Height          =   180
               Left            =   150
               TabIndex        =   6
               Top             =   1080
               Width           =   630
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "发车时间:"
               Height          =   180
               Left            =   150
               TabIndex        =   5
               Top             =   720
               Width           =   810
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "车次:"
               Height          =   180
               Left            =   150
               TabIndex        =   4
               Top             =   345
               Width           =   450
            End
         End
         Begin VB.Timer tmrEventRaise 
            Enabled         =   0   'False
            Left            =   1650
            Top             =   2730
         End
         Begin RTComctl2.RevTimer RevTimer1 
            Height          =   405
            Left            =   0
            Top             =   500
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   714
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   0   'False
         End
         Begin RTComctl3.FlatLabel flbNowTime 
            Height          =   390
            Left            =   1020
            TabIndex        =   27
            Top             =   3810
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   688
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   14737632
            OutnerStyle     =   2
            NormTextColor   =   12582912
            Caption         =   "11:11:11"
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "当前时间:"
            Height          =   180
            Left            =   150
            TabIndex        =   28
            Top             =   3900
            Width           =   810
         End
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   870
      Top             =   4050
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1035
      Top             =   1950
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":1CBAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":1CD0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":1CE66
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "MDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lErrorCode As Long                  '错误号
Public WithEvents moMessage As STNotify.MsgNotify  '消息接收对象
Attribute moMessage.VB_VarHelpID = -1
Dim meEventMode As eEventId
Dim aszEventParam(1 To 6) As String
'事件参数    1---Paramvalue1
'            2---Paramvalue2
'            3---Paramvalue3


Private Sub CloseSystem()
    Dim S As Form
    
    WriteInitReg
  
  
  
'  For Each S In Forms
'      Unload S
'  Next
  Set S = Nothing
End Sub





Private Sub abMenu_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
Select Case Tool.name
    '系统
    Case "mnu_System_Option"
        mnu_System_Option_Click
    Case "mnu_System_ChangeSheetNo"
        mnu_System_ChangeSheetNo_Click
    Case "mnu_System_ChangeUserPwd"
        mnu_System_ChangeUserPwd_Click
    Case "mnu_System_Exit"
    
        '临时测试
        
'        mnu_System_Exit_Click
        Unload Me
    '检票
    Case "mnu_Check_CheckBus"
        mnu_Check_CheckBus_Click
    Case "mnu_Check_RefreshSeat"
        mnu_Check_RefreshSeat_Click
    Case "mnu_Check_Start"
        mnu_Check_Start_Click
    Case "mnu_Check_Extra"
        mnu_Check_Extra_Click
    Case "mnu_Check_Stop"
        mnu_Check_StopCheck_Click
    Case "mnu_Check_WriteOff"
        mnu_Check_WriteOff_Click
    Case "mnu_Check_ReprintSheet"
        mnu_Check_ReprintSheet_Click
    '补打路单
    Case "mnu_Check_ExPrintSheet"
        mnu_Check_ExPrintSheet_Click
        
    Case "mnu_Check_Query"
        frmChkTkQuery.Show vbModal
    Case "mnu_Check_OtherDay"
        FrmCheckExtraTicket.Show vbModal
        
    Case "mnuChangeTicketType"
        '全票特票互改
        ChangeTicketType
    
    '查询
'    Case "mnu_Check_Bus"
'        mnu_Check_Bus_Click
'    Case "mnu_Query_Ticket"
'        mnu_Query_Ticket_Click
    Case "mnu_hlp_content"
        mnu_hlp_content_Click
    Case "mnu_hlp_index"
        mnu_hlp_index_Click
    Case "mnu_hlp_about"
        mnu_hlp_about_Click
        
        
    '工具条
    Case "Checkbus"
        mnu_Check_CheckBus_Click
    Case "Startcheck"
        mnu_Check_Start_Click
    Case "Extracheck"
        mnu_Check_Extra_Click
    Case "ReprintSheet"
        mnu_Check_ReprintSheet_Click
    Case "Exit"
        Unload Me
    
    '帮助
    Case "mnu_hlp_content"
        DisplayHelp Me
    End Select
    
    
End Sub

Public Sub ChangeTicketType()
    frmChangeTicketType.Show vbModal
End Sub


Private Sub MDIForm_Activate()
    Static bNotFirstLoad As Boolean
    If Not bNotFirstLoad Then       '主界面初始化
        
        bNotFirstLoad = True
        
        Me.MousePointer = vbHourglass
    
        ShowSBInfo "正在读取检票口信息..."
        WriteCheckGateInfo
        
        '初始化待检车次
        ShowSBInfo "正在读取下一个待检车次信息..."
        WriteNextBus
        
        tbsBusList.Visible = False
        g_nCurrLineIndex = 0
        
        ShowSBInfo ""
        
'        frmBusList.Show
        Me.MousePointer = vbDefault
    End If
End Sub

Private Sub MDIForm_Load()
  On Error GoTo Here:
    Set moMessage = New MsgNotify
    moMessage.Unit = g_szUnitID
    g_tCheckInfo.CheckGateName = g_oChkTicket.CheckGateName
    
    
    Set g_cWillCheckBusList = New BusCollection
    Set g_cCheckedBusList = New BusCollection
    BuildBusCollection
    
    
    Set abMenu.Bands("bndTitle").Tools("tblTitle").Custom = ptTitle
    Set abMenu.Bands("bndNextBus").Tools("tblNextBus").Custom = ptNextBusInfo
    Exit Sub
Here:
    ShowErrorMsg
End Sub

Private Sub MDIForm_Resize()
    If Me.WindowState = vbNormal Then Me.WindowState = vbMaximized
'    ptNextBusInfo.Height = abMenu.Bands("bndNextBus").Height
'    fraCheckGate.Top = Me.ScaleHeight - fraCheckGate.Height - 200

End Sub

Private Sub MDIForm_Terminate()
  CloseComputer
    End
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    CloseSystem
End Sub

'Private Sub mnu_Check_Bus_Click()
''    If g_nCurrLineIndex > 0 Then
''        Dim oTmp As SNUIcom.CheckSysApp
''        Set oTmp = New CheckSysApp
''        oTmp.ShowCheckInfo g_oActiveUser, Date, g_atCheckLine(g_nCurrLineIndex).BusId, g_atCheckLine(g_nCurrLineIndex).SerialNo
''    End If
''    Dim oTmp As CTQuery
''    Set oTmp = New CTQuery
''    oTmp.ShowBusQuery g_oActiveUser
'    frmQueryBus.SelfUser = g_oActiveUser
'    frmQueryBus.Show , MDIMain
'End Sub

Private Sub mnu_Check_CheckBus_Click()
    frmBusList.ZOrder 0
    frmBusList.Show
End Sub

Private Sub mnu_Check_Extra_Click()
'    If g_oNextEnvBus Is Nothing Then
        frmStartCheck.SetProperty "", True
'    Else
'        frmStartCheck.SetProperty g_oNextEnvBus.BusId, True
'    End If
    frmStartCheck.Show vbModal
End Sub

Private Sub mnu_Check_RefreshSeat_Click()
    On Error Resume Next
    MDIMain.ActiveForm.RefreshSeat
End Sub

Private Sub mnu_Check_ReprintSheet_Click()
    frmRePrintSheet.Show vbModal
End Sub

'补打路单
Private Sub mnu_Check_ExPrintSheet_Click()
    If g_oChkTicket.SelectExPrintSheetIsValid Then
        frmExPrintSheet.Show vbModal
    Else
        MsgboxEx "当前用户没有补打路单的权限，请与管理员联系！", vbExclamation + vbOKOnly
    End If
End Sub

Private Sub mnu_Check_Start_Click()
    frmStartCheck.SetProperty ""
    frmStartCheck.Show vbModal
End Sub

Private Sub mnu_Check_StopCheck_Click()
    On Error Resume Next
    '调用当前检票窗体的停检方法
    MDIMain.ActiveForm.StopCheck
End Sub

Private Sub mnu_Check_WriteOff_Click()
   frmWriteOffCheck.Show vbModal
End Sub

Private Sub mnu_hlp_about_Click()
    Dim oShell As CommShell
    Set oShell = New CommShell
    oShell.ShowAbout "检票系统", "Check Ticket System", "检票系统", Me.Icon, App.Major, App.Minor, App.Revision
End Sub

Private Sub mnu_hlp_content_Click()
    MDIMain.HelpContextID = 20000120
    DisplayHelp Me, content
End Sub

Private Sub mnu_hlp_index_Click()
    DisplayHelp Me, Index
End Sub

'Private Sub mnu_Query_Ticket_Click()
''    Dim szTmpId As String
''    szTmpId = Trim(g_aofrmCheckForm(g_nCurrLineIndex).lblTicketID.Caption)
''    If szTmpId <> "" Then
''        Dim otemp As CheckSysApp
''        Set otemp = New CheckSysApp
''        otemp.ShowTicketCheckInfo g_oActiveUser, szTmpId
''    End If
''    Dim oTmp As CTQuery
''    Set oTmp = New CTQuery
''    oTmp.ShowTicketQuery g_oActiveUser
'    frmQueryTicket.SelfUser = g_oActiveUser
'    frmQueryTicket.Show , MDIMain
'End Sub

Private Sub mnu_System_ChangeSheetNo_Click()
    frmChangeSheetNo.FirstLoad = False
    frmChangeSheetNo.Show vbModal
End Sub

Private Sub mnu_System_ChangeUserPwd_Click()
    Dim oShell As New CommDialog
    On Error GoTo ErrorHandle
    oShell.Init g_oActiveUser
    oShell.ShowUserInfo
    Set oShell = Nothing
    Exit Sub
ErrorHandle:
    ShowErrorMsg
    Set oShell = Nothing
'
End Sub

Private Sub mnu_System_Exit_Click()
    Dim S As Form
    For Each S In Forms
        If S.name <> "MDIMain" Then
            Unload S
        End If
    Next
    Unload Me
    
    Set S = Nothing
    
    End
    
End Sub

Private Sub mnu_System_Option_Click()
    frmSetOption.Show vbModal
End Sub


'Private Sub moMessage_AddBus(ByVal szBusid As String, ByVal dtBusDate As Date)
'    If dtBusDate = Date Then
'        meEventMode = AddBus
'        aszEventParam(1) = Trim(szBusid)
'        aszEventParam(2) = dtBusDate
'        tmrEventRaise.Interval = 1
'        tmrEventRaise.Enabled = True
'    End If
'End Sub

Private Sub moMessage_AddBus1(ByVal szBusid As String, ByVal dtBusDate As Date, ByVal szCheckGate As String)
    If dtBusDate = Date Then
        meEventMode = AddBus
        aszEventParam(1) = Trim(szBusid)
        aszEventParam(2) = dtBusDate
        aszEventParam(3) = szCheckGate
        tmrEventRaise.Interval = 1
        tmrEventRaise.Enabled = True
    End If
End Sub

Private Sub moMessage_ChangeBusCheckGate(ByVal szBusid As String, ByVal dtBusDate As Date, ByVal szCheckGate As String)
    If dtBusDate = Date Then
        meEventMode = ChangeBusCheckGate
        aszEventParam(1) = Trim(szBusid)
        aszEventParam(2) = dtBusDate
        aszEventParam(3) = Trim(szCheckGate)
        tmrEventRaise.Interval = 1
        tmrEventRaise.Enabled = True
    End If
End Sub


Private Sub moMessage_ChangeBusTime(ByVal szBusid As String, ByVal dtBusDate As Date, ByVal dtOffTime As Date)
    If dtBusDate = Date Then
        meEventMode = ChangeBusTime
        aszEventParam(1) = Trim(szBusid)
        aszEventParam(2) = dtBusDate
        aszEventParam(3) = dtOffTime
        tmrEventRaise.Interval = 1
        tmrEventRaise.Enabled = True
    End If
End Sub


Private Sub moMessage_MergeBus(ByVal szBusid As String, ByVal dtBusDate As Date)
    If dtBusDate = Date Then
        meEventMode = MergeBus
        aszEventParam(1) = Trim(szBusid)
        aszEventParam(2) = dtBusDate
        tmrEventRaise.Interval = 1
        tmrEventRaise.Enabled = True
    End If
End Sub



Private Sub moMessage_RemoveBus(ByVal szBusid As String, ByVal dtBusDate As Date)
    If dtBusDate = Date Then
        meEventMode = RemoveBus
        aszEventParam(1) = Trim(szBusid)
        aszEventParam(2) = dtBusDate
        tmrEventRaise.Interval = 1
        tmrEventRaise.Enabled = True
    End If
End Sub


Private Sub moMessage_ResumeBus(ByVal szBusid As String, ByVal dtBusDate As Date)
    If dtBusDate = Date Then
        meEventMode = ResumeBus
        aszEventParam(1) = Trim(szBusid)
        aszEventParam(2) = dtBusDate
        tmrEventRaise.Interval = 1
        tmrEventRaise.Enabled = True
    End If
End Sub


Private Sub moMessage_StopBus(ByVal szBusid As String, ByVal dtBusDate As Date)
    If dtBusDate = Date Then
        meEventMode = StopBus
        aszEventParam(1) = Trim(szBusid)
        aszEventParam(2) = dtBusDate
        tmrEventRaise.Interval = 1
        tmrEventRaise.Enabled = True
    End If
End Sub

Private Sub RevTimer1_Timer()
    Dim lHaveTime As Long
    RevTimer1.Enabled = False
    
    If Not g_oNextEnvBus Is Nothing Then '当前车次是未开检车次则取当前车次
        lHaveTime = DateDiff("s", Now, g_oNextEnvBus.StartUpTime)
        If lHaveTime <= 0 Then '如果发车时间小于当前时间则取下一班车次的信息
            WriteNextBus
        End If
    Else
        WriteNextBus
    End If
End Sub

Private Sub rvtTime_Timer()
    '播放开检时间到的音效
    rvtTime.Second = 0
    EnabledMDITimer False
'    rvtTime.Enabled = False
'    RevTimer1.Second = (DateDiff("n", Now, DateAdd("n", -g_nLatestExtraCheckTime, g_oNextEnvBus.StartUpTime)) + m_cnTimeWindage) * 60
'    RevTimer1.Enabled = True
    
    PlayEventSound g_tEventSoundPath.StartupCheckTimeOn
        
    CloseModalForm
    
    If Not (Screen.ActiveForm Is frmCheckSheet) Then  '路单打印界面激活时，不显示，否则路单号会跳号
    
        If Trim(g_oActiveUser.SellStationID) <> "cm" Then
        '打开开检窗口
        frmStartCheck.SetProperty g_oNextEnvBus.BusID, False
        frmStartCheck.Timer1.Enabled = True
        frmStartCheck.Show vbModal
        End If
    End If
        
'    writeNextBus
End Sub

Private Sub tbsBusList_Click()
    If g_nCurrLineIndex <> tbsBusList.SelectedItem.Index Then
        g_nCurrLineIndex = tbsBusList.SelectedItem.Index
        g_aofrmCheckForm(g_nCurrLineIndex).ZOrder
        
'        szSeatBusID = g_aofrmCheckForm(g_nCurrLineIndex).LblszBusID
    End If
    
'    m_nPrevLineIndex = g_nCurrLineIndex
'    g_nCurrLineIndex = tbsBusList.SelectedItem.Index
'    If Not m_bIsFormActive Then
'        OpenCurrCheckLine
'    Else
'        ResetEnvBusInfo g_atCheckLine(g_nCurrLineIndex).BusId
'    End If
End Sub




Private Sub Timer2_Timer()
    flbNowTime.Caption = Format(Time, cszTimeStr)
End Sub

Private Sub tmrEventRaise_Timer()
    tmrEventRaise.Enabled = False
    RunMsgEvent meEventMode, aszEventParam
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "cmdCheckBus"
            mnu_Check_CheckBus_Click
        Case "cmdStartCheck"
            mnu_Check_Start_Click
        Case "cmdExCheck"
            mnu_Check_Extra_Click
        Case "cmdReprint"
            mnu_Check_ReprintSheet_Click
    End Select
End Sub
'Private Sub DisplayHelp(Optional HelpType As EHelpType = content)
'    Dim lActiveControl As Long
'
'    Select Case HelpType
'        Case content
'            lActiveControl = Me.ActiveControl.HelpContextID
'            If lActiveControl = 0 Then
'                TopicID = Me.HelpContextID
'                CallHTMLShowTopicID
'            Else
'                TopicID = lActiveControl
'                CallHTMLShowTopicID
'            End If
'        Case Index
'            CallHTMLHelpIndex
'        Case Support
'            TopicID = clSupportID
'            CallHTMLShowTopicID
'    End Select
'
'End Sub
'
'
