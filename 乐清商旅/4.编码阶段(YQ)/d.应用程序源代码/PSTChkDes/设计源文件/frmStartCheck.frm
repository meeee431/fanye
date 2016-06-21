VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmStartCheck 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "启动检票"
   ClientHeight    =   4440
   ClientLeft      =   3960
   ClientTop       =   3615
   ClientWidth     =   6030
   HelpContextID   =   2000050
   Icon            =   "frmStartCheck.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Tag             =   "Modal"
   Begin RTComctl3.CoolButton cmdFind 
      Height          =   285
      Left            =   4860
      TabIndex        =   31
      Top             =   2330
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   503
      BTYPE           =   3
      TX              =   "查询"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmStartCheck.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton CoolButton1 
      Height          =   345
      Left            =   1380
      TabIndex        =   30
      Top             =   3990
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "帮助"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmStartCheck.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdExit 
      Cancel          =   -1  'True
      Height          =   345
      Left            =   4260
      TabIndex        =   5
      Top             =   3990
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "取消"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmStartCheck.frx":0044
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtBusID 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1200
      TabIndex        =   0
      Top             =   840
      Width           =   1965
   End
   Begin VB.PictureBox ptTitle 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   6615
      TabIndex        =   26
      Top             =   0
      Width           =   6615
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   60
         Left            =   -30
         TabIndex        =   27
         Top             =   660
         Width           =   6885
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "以下车次已到检票时间，请开始检票"
         Height          =   180
         Left            =   120
         TabIndex        =   28
         Top             =   270
         Width           =   2880
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   855
      Left            =   2250
      TabIndex        =   4
      Top             =   2640
      Width           =   3390
      Begin VB.Label lblBusStop 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "已停车"
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   2640
         TabIndex        =   29
         Top             =   540
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Label label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车主:"
         Height          =   180
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   450
      End
      Begin VB.Label label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车型:"
         Height          =   180
         Left            =   1680
         TabIndex        =   10
         Top             =   240
         Width           =   450
      End
      Begin VB.Label label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "所属公司:"
         Height          =   180
         Left            =   120
         TabIndex        =   9
         Top             =   540
         Width           =   810
      End
      Begin VB.Label lblOwner 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   570
         TabIndex        =   8
         Top             =   240
         Width           =   90
      End
      Begin VB.Label lblVehicleType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   2160
         TabIndex        =   7
         Top             =   240
         Width           =   90
      End
      Begin VB.Label lblCompany2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   960
         TabIndex        =   6
         Top             =   510
         Width           =   90
      End
   End
   Begin VB.ComboBox CboVehicle 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3060
      TabIndex        =   2
      Text            =   "CboVehicle"
      ToolTipText     =   "按F8键查询车辆"
      Top             =   2310
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6180
      Top             =   2130
   End
   Begin RTComctl3.CoolButton cmdStartCheck 
      Height          =   345
      Left            =   2730
      TabIndex        =   3
      Top             =   3990
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "开始检票(&S)"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmStartCheck.frx":0060
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   " "
      Enabled         =   0   'False
      Height          =   750
      Left            =   -120
      TabIndex        =   25
      Top             =   3750
      Width           =   8745
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   1320
         Top             =   180
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmStartCheck.frx":007C
               Key             =   ""
               Object.Tag             =   "Check"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmStartCheck.frx":0207
               Key             =   ""
               Object.Tag             =   "ExtraCheck"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Image imgEnabled 
      Height          =   480
      Left            =   4725
      Picture         =   "frmStartCheck.frx":079A
      Top             =   1230
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCheckIcon 
      Height          =   630
      Left            =   4200
      Picture         =   "frmStartCheck.frx":1064
      Stretch         =   -1  'True
      Top             =   870
      Width           =   660
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   285
      Left            =   3510
      TabIndex        =   24
      Top             =   1710
      Width           =   2145
      WordWrap        =   -1  'True
   End
   Begin VB.Label label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车次类型:"
      Height          =   180
      Left            =   360
      TabIndex        =   23
      Top             =   2355
      Width           =   810
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "检票车次:"
      Height          =   180
      Left            =   360
      TabIndex        =   22
      Top             =   930
      Width           =   810
   End
   Begin VB.Label label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "售票数:"
      Height          =   180
      Left            =   360
      TabIndex        =   21
      Top             =   2790
      Width           =   630
   End
   Begin VB.Label label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "是否拆分:"
      Height          =   180
      Left            =   360
      TabIndex        =   20
      Top             =   3210
      Width           =   810
   End
   Begin VB.Label lblMergeBus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   180
      Left            =   1200
      TabIndex        =   19
      Top             =   3210
      Width           =   90
   End
   Begin VB.Label lblBusSerial 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3240
      TabIndex        =   18
      Top             =   885
      Width           =   300
   End
   Begin VB.Line Line1 
      X1              =   360
      X2              =   5670
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label lblEndStation 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "杭州"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   1200
      TabIndex        =   17
      Top             =   1710
      Width           =   1965
   End
   Begin VB.Label lblStartupTime 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "9:20"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   1200
      TabIndex        =   16
      Top             =   1290
      Width           =   1965
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "开往:"
      Height          =   180
      Left            =   360
      TabIndex        =   15
      Top             =   1800
      Width           =   450
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "运行车辆:"
      Height          =   180
      Left            =   2250
      TabIndex        =   1
      Top             =   2355
      Width           =   810
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "发车时间:"
      Height          =   180
      Left            =   360
      TabIndex        =   14
      Top             =   1380
      Width           =   810
   End
   Begin VB.Label lblBusMode 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   180
      Left            =   1200
      TabIndex        =   13
      Top             =   2355
      Width           =   90
   End
   Begin VB.Label lblSellTickets 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   180
      Left            =   1200
      TabIndex        =   12
      Top             =   2790
      Width           =   90
   End
End
Attribute VB_Name = "frmStartCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'显示模式

Private mszBusID As String                      '当前选择的车次号
Private mbExCheckMode As Boolean                '当前选择的车次是否是补检

Private mnCheckMode As Integer                  '当前所选车次的检票状态
Private matRunVehicle() As M_TRunVehicle        '可以选择的所有车辆信息数组
Private mnLastSearchIndex As Integer            '选择车辆时用到的定位索引
Private mbIsShow As Boolean

Private mnSerialNo As Integer                   '滚动车次的运行序号

'根据车次检票状态来布局界面
Private Sub LayoutCheckInfo()
    Select Case mnCheckMode
        Case ECS_CanotCheck
            imgEnabled.Visible = True
            lblMessage.Caption = "车次不能检票"
            cmdStartCheck.Enabled = False
            CboVehicle.Enabled = False
            cmdFind.Enabled = False
            txtBusID.SetFocus
        Case ECS_CanCheck
            imgEnabled.Visible = False
            lblMessage.Caption = "检票时间已到,请检票"
            cmdStartCheck.Enabled = True
            CboVehicle.Enabled = True
            cmdFind.Enabled = True
        Case ECS_BeChecking
            imgEnabled.Visible = False
            lblMessage.Caption = "车次正在检票"
            cmdStartCheck.Enabled = True
            CboVehicle.Enabled = False
            cmdFind.Enabled = False
        Case ECS_CanExtraCheck
            imgEnabled.Visible = False
            lblMessage.Caption = "车次已检,可以补检"
            cmdStartCheck.Enabled = True
            CboVehicle.Enabled = False
            cmdFind.Enabled = False
            mbExCheckMode = True
        Case ECS_BeExtraChecking
            imgEnabled.Visible = False
            lblMessage.Caption = "车次正在补检"
            cmdStartCheck.Enabled = True
            CboVehicle.Enabled = False
            cmdFind.Enabled = False
            mbExCheckMode = True
        Case ECS_Checked
            imgEnabled.Visible = True
            lblMessage.Caption = "车次已停检"
            cmdStartCheck.Enabled = False
            CboVehicle.Enabled = False
            cmdFind.Enabled = False
    End Select
    
End Sub

Private Function SetBusCheckStatus() As Integer
    '判断指定车次状态
    Dim dptCheckBus As Date
    Dim lHaveTime As Long
    Dim nResult As ECheckStatus
    If g_tCheckInfo.BusMode = TP_ScrollBus Then
        '取得该滚动车次的最后一次检票的状态，主要应付异常中断的情况
        'g_tCheckInfo.SerialNo中存放应检票的滚动车次序号
        g_tCheckInfo.SerialNo = g_oChkTicket.GetNextScrollNo(g_tCheckInfo.BusID)
        If g_tCheckInfo.SerialNo > 0 Then
'            nResult = g_oChkTicket.GetBusStatus(Date, g_tCheckInfo.BusID, g_tCheckInfo.SerialNo - 1)
            nResult = g_oEnvBus.busStatus
            Select Case nResult
                Case EREBusStatus.ST_BusChecking
                    nResult = ECS_BeChecking
                Case EREBusStatus.ST_BusExtraChecking
                    nResult = ECS_BeExtraChecking
                Case EREBusStatus.ST_BusNormal
                    nResult = ECS_CanCheck
                    '判断是否为补检
                    If mbExCheckMode And g_tCheckInfo.SerialNo > 1 Then
                        '设置被检的序号
                        If mnSerialNo = 0 Then
                            g_tCheckInfo.SerialNo = g_tCheckInfo.SerialNo - 1
                        Else
                            g_tCheckInfo.SerialNo = mnSerialNo ' g_tCheckInfo.SerialNo - 1
                        End If
                        nResult = ECS_CanExtraCheck
                    Else
                        nResult = ECS_CanCheck
                    End If
                Case Else
                    nResult = ECS_CanotCheck
            End Select
        Else
            nResult = ECS_CanCheck
        End If
    Else
        g_tCheckInfo.SerialNo = 0
'        nResult = g_oChkTicket.GetBusStatus(Date, g_tCheckInfo.BusID, 0)
        nResult = g_oEnvBus.busStatus
        Select Case nResult
            Case EREBusStatus.ST_BusChecking
                nResult = ECS_BeChecking
            Case EREBusStatus.ST_BusExtraChecking
                nResult = ECS_BeExtraChecking
            Case EREBusStatus.ST_BusReplace, EREBusStatus.ST_BusNormal
                nResult = ECS_CanCheck
            Case EREBusStatus.ST_BusMergeStopped, EREBusStatus.ST_BusSlitpStop, EREBusStatus.ST_BusStopped
                nResult = ECS_CanotCheck
            Case EREBusStatus.ST_BusStopped
                nResult = ECS_CanExtraCheck
        End Select
    
        dptCheckBus = Now
        lHaveTime = DateDiff("s", dptCheckBus, DateAdd("n", -g_nLatestExtraCheckTime, g_tCheckInfo.StartUpTime))
        If lHaveTime < 0 Then   '已过了最晚检票时间
            nResult = ECS_Checked
        Else
            If nResult = ECS_CanCheck Then
                lHaveTime = DateDiff("s", dptCheckBus, DateAdd("n", -g_nBeginCheckTime, g_tCheckInfo.StartUpTime))
                If lHaveTime > 0 Then       '还未到该检票时间
                    '根据系统参数来设置是否未到检票时间允许检票
                    If g_bAllowStartChectNotRearchTime Then
                        nResult = ECS_CanCheck
                    Else
                        nResult = ECS_CanotCheck
                    End If
                End If
            End If
        End If
    End If
    mnCheckMode = nResult
End Function


Private Sub RefreshBusInfo()
    Dim i As Integer
'    Dim TVehicle() As TBusInfo
    Dim nCount As Integer
    
    'mbExCheckMode = False
    SetBusCheckStatus
    LayoutCheckInfo
    
    '填充运行车辆
    CboVehicle.Clear
    matRunVehicle = g_oChkTicket.GetRunVehicle(g_tCheckInfo.BusID)
    For i = 1 To ArrayLength(matRunVehicle)
        CboVehicle.AddItem MakeDisplayString(matRunVehicle(i).VehicleId, matRunVehicle(i).Vehicle)
    Next i
    For i = 0 To CboVehicle.ListCount - 1
        If ResolveDisplay(CboVehicle.List(i)) = g_tCheckInfo.VehicleId Then
            CboVehicle.ListIndex = i
            Exit For
        End If
    Next i
    
    If i = CboVehicle.ListCount Then
        CboVehicle.AddItem MakeDisplayString(g_tCheckInfo.VehicleId, g_tCheckInfo.Vehicle)
        CboVehicle.ListIndex = i
    End If
    
    If g_tCheckInfo.BusMode = TP_ScrollBus Then
        '设置滚动车次信息
        lblEndStation.Caption = g_tCheckInfo.EndStationName
        lblStartupTime.Caption = ""
        lblBusMode.Caption = g_cszTitleScollBus
        lblSellTickets.Caption = ""
        lblMergeBus.Caption = "否"
        lblBusSerial.Caption = "-" & g_tCheckInfo.SerialNo
        
'        If CboVehicle.ListCount > 0 Then
'            CboVehicle.ListIndex = 0
'        End If
    Else
        '固定车次
        
        lblEndStation.Caption = g_tCheckInfo.EndStationName
        lblStartupTime.Caption = Format(g_tCheckInfo.StartUpTime, "HH:mm")
        lblBusMode.Caption = "固定车次"
        g_tCheckInfo.SellTickets = g_oEnvBus.GetNotCanSellCount(g_oActiveUser.SellStationID)
        lblSellTickets.Caption = g_tCheckInfo.SellTickets
        If g_tCheckInfo.MergeType = 1 Then
            lblMergeBus.Caption = "拆分"
        Else
            lblMergeBus.Caption = "无"
        End If
    End If
End Sub


'
Private Sub cboVehicle_Click()
    ShowVehicle ResolveDisplay(CboVehicle.Text)
'    lblOwner.Caption = matRunVehicle(CboVehicle.ListIndex + 1).Owner
'    lblVehicleType.Caption = matRunVehicle(CboVehicle.ListIndex + 1).VehicleType
'    lblCompany2.Caption = matRunVehicle(CboVehicle.ListIndex + 1).Company
'    lblBusStop.Visible = IIf(matRunVehicle(CboVehicle.ListIndex + 1).Status = 0, False, True)
End Sub


Private Sub CboVehicle_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then
'        SendKeys "{TAB}"
'    End If
    ShowVehicle ResolveDisplay(CboVehicle.Text)
End Sub

Private Sub CboVehicle_Validate(Cancel As Boolean)
'On Error GoTo ErrHandle
'    Dim i As Integer
'    Dim szInputVehicle As String
'    szInputVehicle = ResolveDisplay(CboVehicle.Text)
'    For i = 1 To ArrayLength(matRunVehicle)
'        If szInputVehicle = matRunVehicle(i).VehicleId Then
'            lblOwner.Caption = matRunVehicle(i).Owner
'            lblVehicleType.Caption = matRunVehicle(i).VehicleType
'            lblCompany2.Caption = matRunVehicle(i).Company
'            lblBusStop.Visible = IIf(matRunVehicle(i).Status = 0, False, True)
'            Exit For
'        End If
'    Next i
'    If i > ArrayLength(matRunVehicle) Then      '未在运行车辆中
'        If MsgboxEx("本车辆未在该车次安排中，是否真的用本车辆顶替?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
'            ShowVehicle szInputVehicle
'        Else
'            Cancel = True
'        End If
'    End If
'    Exit Sub
'ErrHandle:
'    ShowErrorMsg
'    Cancel = True
End Sub

Private Sub cmdExit_Click()
    If g_nCurrLineIndex > 0 Then
        MDIMain.tbsBusList.Tabs(g_nCurrLineIndex).Selected = True
    End If
    Unload Me
End Sub

Private Sub cmdStartCheck_Click()
On Error GoTo ErrHandle
'    Dim TVehicle() As TBusInfo
    Dim i As Integer
    Dim nCheckLineCount As Integer
    
    
    If lblBusStop.Visible Then
        MsgboxEx "该车已停车，请选择其他车辆！", vbExclamation, Me.Caption
        Exit Sub
    End If
    
    ShowSBInfo "正在准备开始检票..."
    nCheckLineCount = CheckLineCount
    For i = 1 To nCheckLineCount
        If g_atCheckLine(i).BusID = g_tCheckInfo.BusID And g_atCheckLine(i).SerialNo = g_tCheckInfo.SerialNo Then
            Exit For
        End If
    Next
    If i > nCheckLineCount Then
        '不管固定车次和滚动车次,把运行车辆信息放入g_tCheckInfo.RunVehicle变量中
        Dim szVehicleLicense As String
        
        '如果为顶班车次，则以环境里的车牌为准；否则以界面传入为准；FPD 2007-11-21
        ResetEnvBusInfo g_tCheckInfo.BusID, 0
        If g_oEnvBus.busStatus = ST_BusReplace Then
            g_tCheckInfo.RunVehicle.VehicleId = g_oEnvBus.Vehicle
            g_tCheckInfo.RunVehicle.Vehicle = g_oEnvBus.VehicleTag
        Else
            g_tCheckInfo.RunVehicle.VehicleId = ResolveDisplay(CboVehicle.Text, szVehicleLicense)
            g_tCheckInfo.RunVehicle.Vehicle = szVehicleLicense
        End If
        
        g_tCheckInfo.RunVehicle.Owner = lblOwner.Caption
        g_tCheckInfo.RunVehicle.VehicleType = lblVehicleType.Caption
        g_tCheckInfo.RunVehicle.Company = lblCompany2.Caption
        
        If mnCheckMode = ECS_BeChecking Or mnCheckMode = ECS_BeExtraChecking Then
            Unload Me
            AddNewCheckLine g_tCheckInfo.BusID, mbExCheckMode, True, , g_oEnvBus
        Else
            Unload Me
            AddNewCheckLine g_tCheckInfo.BusID, mbExCheckMode, False, g_tCheckInfo.RunVehicle.VehicleId, g_oEnvBus
        End If
    Else
        Unload Me
        MDIMain.tbsBusList.Tabs(i).Selected = True
    End If
    ShowSBInfo ""
    WriteNextBus
    Exit Sub
ErrHandle:
    ShowErrorMsg
    ShowSBInfo ""
End Sub




Private Sub CoolButton1_Click()
    DisplayHelp Me

End Sub

Private Sub cmdFind_Click()
    Dim oShell As New STShell.CommDialog
    Dim aszTemp() As String
    On Error GoTo ErrorHandle
        If CboVehicle.Enabled Then
            oShell.Init g_oActiveUser
            aszTemp = oShell.SelectVehicleEX
            If ArrayLength(aszTemp) > 0 Then CboVehicle.Text = MakeDisplayString(aszTemp(1, 1), aszTemp(1, 2))
            Set oShell = Nothing
        End If
        If cmdStartCheck.Enabled Then cmdStartCheck.SetFocus
    Exit Sub
ErrorHandle:
    Set oShell = Nothing
    ShowErrorMsg
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim oShell As New STShell.CommDialog
    Dim aszTemp() As String
    On Error GoTo ErrorHandle
    If KeyCode = vbKeyF8 And g_oChkTicket.SelectChangeBusBeforeCheetIsValid Then
        If CboVehicle.Enabled Then
'            frmSearchVechile.StartSearchIndex = mnLastSearchIndex
'            frmSearchVechile.Show vbModal
'            mnLastSearchIndex = CboVehicle.ListIndex
            oShell.Init g_oActiveUser
            aszTemp = oShell.SelectVehicleEX
            If ArrayLength(aszTemp) > 0 Then CboVehicle.Text = MakeDisplayString(aszTemp(1, 1), aszTemp(1, 2))
            Set oShell = Nothing
        End If
    End If
    
    If KeyCode = vbKeyF Then
        If CboVehicle.Enabled Then
            frmSearchVechile.Show vbModal
            cmdStartCheck.SetFocus
        End If
    End If
    Exit Sub
ErrorHandle:
    Set oShell = Nothing
    ShowErrorMsg
End Sub

Private Sub Form_Load()
'    txtBusID.Text = mszBusID
    AlignFormPos Me
    SelectChangeBusBeforeCheetValid
End Sub


Private Sub Form_Unload(Cancel As Integer)
    mbIsShow = False
    SaveFormPos Me
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
'    txtBusID_Validate False
    txtBusId_KeyPress (vbKeyReturn)
End Sub

Private Sub txtBusId_GotFocus()
    txtBusID.SelStart = 0
    txtBusID.SelLength = Len(txtBusID.Text)
End Sub

Private Sub txtBusId_KeyPress(KeyAscii As Integer)
On Error GoTo ErrHandle
    Dim lErrorCode As Long
    '取得车次信息
    '判断是否能够正常检票
    
    If KeyAscii = vbKeyReturn Then
        
        txtBusID.Text = Trim(txtBusID.Text)
        SetDefaultForm
        
        If txtBusID.Text = "" Then
            Exit Sub
        End If
        
        ShowSBInfo "正在读取车次信息..."
        Me.MousePointer = vbHourglass
        
        If txtBusID.Text <> "" Then
            m_lErrorCode = 0
            ResetEnvBusInfo txtBusID.Text, lErrorCode
            If lErrorCode <> 0 Then
                'Cancel = True
                txtBusID.SetFocus
                GoTo EndValidation
            End If
            mszBusID = txtBusID.Text
          
            RefreshBusInfo
            '设置方便的焦点
            If g_tCheckInfo.BusMode = TP_ScrollBus Then
                If CboVehicle.Enabled Then
                    CboVehicle.SetFocus
                Else
                    If cmdStartCheck.Enabled Then cmdStartCheck.SetFocus
                End If
            Else
                If cmdStartCheck.Enabled Then cmdStartCheck.SetFocus
            End If
        End If
        If cmdStartCheck.Enabled Then cmdStartCheck.SetFocus
EndValidation:
        ShowSBInfo ""
        Me.MousePointer = vbDefault
    End If
    Exit Sub
ErrHandle:
    ShowSBInfo ""
    Me.MousePointer = vbDefault
    ShowErrorMsg
    
End Sub


Public Sub SetProperty(BusID As String, Optional ExChecked As Boolean = False, Optional pnSerialNo As Integer = 0)
    mszBusID = BusID
    txtBusID.Text = mszBusID
    mbExCheckMode = ExChecked
    mnSerialNo = pnSerialNo
    Timer1.Enabled = True
End Sub

Public Property Get IsShow() As Boolean
    IsShow = mbIsShow
End Property

'Private Sub txtBusID_Validate(Cancel As Boolean)
Private Sub txtBusID_LostFocus()

End Sub

Public Property Get BusID() As String
    BusID = mszBusID
End Property
Public Property Get ExChecked() As Boolean
    ExChecked = mbExCheckMode
End Property



'显示车辆信息:车辆车牌、参运公司、车主
'##ModelId=38952F9F030C
Private Sub ShowVehicle(VehicleId As String)
    Dim oVehicle As New Vehicle
    On Error GoTo ErrorHandle
    oVehicle.Init g_oActiveUser
    oVehicle.Identify VehicleId
    lblOwner.Caption = oVehicle.Owner
    lblVehicleType.Caption = oVehicle.VehicleModelName
    lblCompany2.Caption = oVehicle.CompanyName
    lblBusStop.Visible = IIf(oVehicle.Status = 0, False, True)
    Set oVehicle = Nothing
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub





Private Sub SetFormCaption()
    If mbExCheckMode Then
        '正处在补检模式下
        Me.Caption = "启动检票(补检)"
        imgCheckIcon.Picture = ImageList1.ListImages("ExtraCheck").Picture
        lblCaption.Caption = "对以下车次进行补检"

    Else
        '正处在正常检票模式下
        Me.Caption = "启动检票"
        imgCheckIcon.Picture = ImageList1.ListImages("Check").Picture
        lblCaption.Caption = "以下车次已到检票时间，请开始检票"
    End If
End Sub

Private Sub SetDefaultForm()
    lblMessage.Caption = ""
    lblEndStation.Caption = ""
    lblStartupTime.Caption = ""
    lblBusMode.Caption = ""
    lblSellTickets.Caption = ""
    lblMergeBus.Caption = ""
    lblBusSerial.Caption = ""
    lblOwner.Caption = ""
    lblVehicleType.Caption = ""
    lblCompany2.Caption = ""
    lblBusStop.Caption = ""
    cmdStartCheck.Enabled = False
    CboVehicle.Clear
    CboVehicle.Enabled = True
End Sub
'
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

'是否有更改车辆的权限 开检前
'判断当前用户有更改车辆的权限 开检前的权限
Private Sub SelectChangeBusBeforeCheetValid()
    On Error GoTo Here
    If g_oChkTicket.SelectChangeBusBeforeCheetIsValid Then
        cmdFind.Enabled = True
    Else
        cmdFind.Enabled = False
    End If
    On Error GoTo 0
    Exit Sub
Here:
    ShowErrorMsg
End Sub

