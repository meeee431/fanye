VERSION 5.00
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Begin VB.Form frmEnvBusReplace 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "环境--车次顶班"
   ClientHeight    =   4080
   ClientLeft      =   4140
   ClientTop       =   3180
   ClientWidth     =   6165
   HelpContextID   =   10000140
   Icon            =   "frmREBusReplace.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   Begin RTComctl3.CoolButton cmdHelp 
      Height          =   345
      Left            =   4830
      TabIndex        =   10
      Top             =   1110
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "帮助(&H)"
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
      MICON           =   "frmREBusReplace.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   345
      Left            =   4830
      TabIndex        =   9
      Top             =   600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "关闭(&C)"
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
      MICON           =   "frmREBusReplace.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdOk 
      Height          =   345
      Left            =   4830
      TabIndex        =   8
      Top             =   90
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "顶班(&O)"
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
      MICON           =   "frmREBusReplace.frx":0044
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.CheckBox chkOutoSeatType 
      BackColor       =   &H00E0E0E0&
      Caption         =   "匹配座位类型"
      Height          =   285
      Left            =   2370
      TabIndex        =   7
      Top             =   3675
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.Frame frmNewBus 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   810
      TabIndex        =   20
      Top             =   2610
      Width           =   4080
      Begin VB.Label lblOwner 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Index           =   1
         Left            =   2745
         TabIndex        =   28
         Top             =   0
         Width           =   90
      End
      Begin VB.Label lblCorp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Index           =   1
         Left            =   2745
         TabIndex        =   27
         Top             =   285
         Width           =   90
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车主:"
         Height          =   180
         Left            =   1920
         TabIndex        =   26
         Top             =   0
         Width           =   450
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "参运公司:"
         Height          =   180
         Left            =   1920
         TabIndex        =   25
         Top             =   285
         Width           =   810
      End
      Begin VB.Label lblSeat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Index           =   1
         Left            =   660
         TabIndex        =   24
         Top             =   285
         Width           =   90
      End
      Begin VB.Label lblBusType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Index           =   1
         Left            =   660
         TabIndex        =   23
         Top             =   0
         Width           =   90
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "座位数:"
         Height          =   180
         Left            =   0
         TabIndex        =   22
         Top             =   285
         Width           =   630
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车型:"
         Height          =   180
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   450
      End
   End
   Begin VB.Frame fraOldBus 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "原车辆:"
      Height          =   555
      Left            =   795
      TabIndex        =   11
      Top             =   1440
      Width           =   3975
      Begin VB.Label lblOwner 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Index           =   0
         Left            =   2685
         TabIndex        =   19
         Top             =   15
         Width           =   90
      End
      Begin VB.Label lblCorp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Index           =   0
         Left            =   2685
         TabIndex        =   18
         Top             =   300
         Width           =   90
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车主:"
         Height          =   180
         Left            =   1815
         TabIndex        =   17
         Top             =   15
         Width           =   450
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "参运公司:"
         Height          =   180
         Left            =   1815
         TabIndex        =   16
         Top             =   300
         Width           =   810
      End
      Begin VB.Label lblSeat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Index           =   0
         Left            =   645
         TabIndex        =   15
         Top             =   300
         Width           =   90
      End
      Begin VB.Label lblBusType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Index           =   0
         Left            =   645
         TabIndex        =   14
         Top             =   15
         Width           =   90
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "座位数:"
         Height          =   180
         Left            =   0
         TabIndex        =   13
         Top             =   300
         Width           =   630
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车型:"
         Height          =   180
         Left            =   0
         TabIndex        =   12
         Top             =   15
         Width           =   450
      End
   End
   Begin VB.CheckBox chkAllRefundment 
      BackColor       =   &H00E0E0E0&
      Caption         =   "全额退票(&Q)"
      Height          =   225
      Left            =   825
      TabIndex        =   4
      Top             =   3405
      Width           =   1320
   End
   Begin VB.CheckBox chkAutoSeat 
      BackColor       =   &H00E0E0E0&
      Caption         =   "自动座位填充(&S)"
      Height          =   210
      Left            =   2385
      TabIndex        =   5
      Top             =   3405
      Value           =   1  'Checked
      Width           =   1860
   End
   Begin VB.CheckBox chkAutoTicketPrice 
      BackColor       =   &H00E0E0E0&
      Caption         =   "票价更新(&N)"
      Height          =   195
      Left            =   825
      TabIndex        =   6
      Top             =   3735
      Width           =   1320
   End
   Begin FText.asFlatTextBox txtBusID 
      Height          =   300
      Left            =   1920
      TabIndex        =   1
      Top             =   75
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonHotBackColor=   -2147483633
      ButtonPressedBackColor=   -2147483627
      Text            =   ""
      ButtonBackColor =   -2147483633
      ButtonVisible   =   -1  'True
      Registered      =   -1  'True
   End
   Begin FText.asFlatTextBox txtVehicle 
      Height          =   300
      Left            =   1920
      TabIndex        =   3
      Top             =   2235
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonHotBackColor=   -2147483633
      ButtonPressedBackColor=   -2147483627
      Text            =   ""
      ButtonBackColor =   -2147483633
      ButtonVisible   =   -1  'True
      Registered      =   -1  'True
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "顶班设置"
      Height          =   180
      Left            =   195
      TabIndex        =   40
      Top             =   3150
      Width           =   720
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FFFFFF&
      X1              =   960
      X2              =   4700
      Y1              =   3255
      Y2              =   3255
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00808080&
      X1              =   960
      X2              =   4700
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   960
      X2              =   4700
      Y1              =   2115
      Y2              =   2115
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      X1              =   960
      X2              =   4700
      Y1              =   2100
      Y2              =   2100
   End
   Begin VB.Label lblVehicleID 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   180
      Left            =   1620
      TabIndex        =   39
      Top             =   1200
      Width           =   90
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车辆代码:"
      Height          =   180
      Left            =   795
      TabIndex        =   38
      Top             =   1200
      Width           =   810
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   825
      X2              =   4700
      Y1              =   1095
      Y2              =   1095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   825
      X2              =   4700
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "顶班车辆(&B):"
      Height          =   180
      Left            =   795
      TabIndex        =   2
      Top             =   2295
      Width           =   1080
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "顶班车辆"
      Height          =   180
      Left            =   195
      TabIndex        =   37
      Top             =   2040
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "原车辆"
      Height          =   180
      Left            =   195
      TabIndex        =   36
      Top             =   1005
      Width           =   540
   End
   Begin VB.Label lblRoute 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   180
      Left            =   1620
      TabIndex        =   35
      Top             =   465
      Width           =   90
   End
   Begin VB.Label lblSellSeat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   180
      Left            =   3840
      TabIndex        =   34
      Top             =   465
      Width           =   90
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "已售或留票:"
      Height          =   180
      Left            =   2835
      TabIndex        =   33
      Top             =   465
      Width           =   1035
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "运行线路:"
      Height          =   180
      Left            =   795
      TabIndex        =   32
      Top             =   465
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "发车时间:"
      Height          =   180
      Left            =   795
      TabIndex        =   31
      Top             =   780
      Width           =   810
   End
   Begin VB.Label lblOffTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   180
      Left            =   1620
      TabIndex        =   30
      Top             =   780
      Width           =   90
   End
   Begin VB.Label lblBusID 
      AutoSize        =   -1  'True
      Caption         =   "0001"
      Height          =   180
      Left            =   1995
      TabIndex        =   29
      Top             =   150
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车次代码(&I):"
      Height          =   180
      Left            =   795
      TabIndex        =   0
      Top             =   135
      Width           =   1080
   End
End
Attribute VB_Name = "frmEnvBusReplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_szBusID As String
Public m_dtBusDate As Date
Public m_bIsParent As Boolean

Private m_oReBus As New REBus
Private m_oVehicle As New Vehicle
Private m_szVehicleSeatCount As Integer
Private WithEvents oRoutePrice   As RoutePriceTable
Attribute oRoutePrice.VB_VarHelpID = -1
Private m_szMakePriceIsNotSuccess As String
Private m_bQueryVehicle As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    DisplayHelp Me
End Sub

Private Sub cmdOk_Click()
    Dim blfg As Boolean
    blfg = RefreshVehicle(ResolveDisplay(txtVehicle.Text), 1)
    If blfg = True Then
        '车辆手工输入
        ReplaceVehicle
    End If
End Sub

Private Sub Form_Load()
    m_oReBus.Init g_oActiveUser
    m_oVehicle.Init g_oActiveUser
    If m_szBusID = "" Then
        m_dtBusDate = Date
    Else
        RefreshBus
    End If
'    Me.Caption = "环境--车次顶班[" & Format(m_dtBusDate, "YYYY年MM月DD日") & "]"

End Sub

Public Sub RefreshBus()
    On Error GoTo ErrorHandle
    m_oReBus.Identify m_szBusID, m_dtBusDate
    m_oReBus.ReBusSlipLock True
    txtBusID.Text = m_szBusID
    lblRoute.Caption = m_oReBus.Route
    lblOffTime.Caption = Format(m_oReBus.StartUpTime, "HH:mm")
    lblVehicleID.Caption = m_oReBus.VehicleTag
    RefreshVehicle m_oReBus.Vehicle, 0
    m_oVehicle.Identify m_oReBus.Vehicle
    lblSellSeat.Caption = m_oReBus.TotalSeat - m_oReBus.SaleSeat
Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Public Function RefreshVehicle(VehicleId As String, Index As Integer) As Boolean
On Error GoTo ErrorHandle
    m_oVehicle.Identify ResolveDisplay(VehicleId)
    If Index = 1 Then txtVehicle.Text = MakeDisplayString(VehicleId, Trim(m_oVehicle.LicenseTag))
    lblSeat(Index).Caption = m_oVehicle.SeatCount
    m_szVehicleSeatCount = m_oVehicle.SeatCount
    lblBusType(Index).Caption = m_oVehicle.VehicleModelName
    lblCorp(Index).Caption = m_oVehicle.CompanyName
    lblOwner(Index).Caption = m_oVehicle.OwnerName
    m_bQueryVehicle = True
    RefreshVehicle = True
Exit Function
ErrorHandle:
    m_bQueryVehicle = False
    RefreshVehicle = False
    ShowErrorMsg
End Function

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrorHandle
    m_oReBus.ReBusSlipLock False
    Set m_oReBus = Nothing
ErrorHandle:
End Sub

Private Sub oRoutePrice_SetMakeBusPriceStatus(ByVal lStatus As String)

    If Right(lStatus, 2) <> "成功" Then
        m_szMakePriceIsNotSuccess = lStatus
    Else
        m_szMakePriceIsNotSuccess = ""
    End If

End Sub



Private Sub txtBusId_Change()
    IsSave
End Sub

Private Sub txtBusID_ButtonClick()
    Dim oBus As New CommDialog
    Dim aszTemp() As String
    oBus.Init g_oActiveUser
    aszTemp = oBus.SelectREBus(m_dtBusDate)
    Set oBus = Nothing
    If ArrayLength(aszTemp) = 0 Then Exit Sub
    txtBusID.Text = aszTemp(1, 1)
    m_szBusID = txtBusID.Text
    RefreshBus
End Sub

Private Sub txtBusID_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
       m_szBusID = txtBusID.Text
       RefreshBus
    End If
End Sub

Private Sub txtVehicle_ButtonClick()
    Dim oBus As New CommDialog
    Dim aszTemp() As String
    oBus.Init g_oActiveUser
    aszTemp = oBus.SelectVehicleEX
    Set oBus = Nothing
    If ArrayLength(aszTemp) = 0 Then Exit Sub
    txtVehicle.Text = MakeDisplayString(aszTemp(1, 1), aszTemp(1, 2))
    RefreshVehicle aszTemp(1, 1), 1
    cmdOk.Enabled = True
    If lblBusType(0).Caption = lblBusType(1).Caption Then
        chkAutoTicketPrice.Enabled = False
    Else
        chkAutoTicketPrice.Enabled = True
    End If
End Sub

Private Sub txtVehicle_Change()
    IsSave
End Sub

Private Sub txtVehicle_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
    RefreshVehicle Trim(txtVehicle.Text), 1
    cmdOk.Enabled = True
    If lblBusType(0).Caption = lblBusType(1).Caption Then
        chkAutoTicketPrice.Enabled = False
    Else
        chkAutoTicketPrice.Enabled = True
    End If
    End If
End Sub


Private Sub IsSave()
    If txtBusID.Text = "" Or txtVehicle.Text = "" Then
        cmdOk.Enabled = False
    Else
        cmdOk.Enabled = True
    End If
End Sub

Private Function ReplaceVehicle()
    '进行顶班操作

    Dim nResult As VbMsgBoxResult
    Dim szMsg As String
    Dim szPriceTable As String
    Dim tVehicleSeatType() As TVehcileSeatType
    Dim nCount2 As Integer
    Dim bflgIsNot As Boolean
    Dim szProject As String
    Dim szReferBusid As String
    Dim tBusVehcileSeatInfo() As TBusVehicleSeatType
    Dim nCount As Integer
    Dim i As Integer
    Dim j As Integer
    Dim oBaseInfo As New BaseInfo
    On Error GoTo ErrorHandle

    oBaseInfo.Init g_oActiveUser
    
    If Val(lblSellSeat.Caption) > m_szVehicleSeatCount Then
        '如果已售座数大于车辆座位数
        szMsg = "车次已售或留票" & lblSellSeat.Caption & "张" & Chr(10)
        szMsg = szMsg & "而用来顶班车辆座位数只有[" & m_szVehicleSeatCount & "]个" & Chr(10)
        szMsg = szMsg & "不能顶班！"
        MsgBox szMsg, vbInformation + vbOKOnly, Me.Caption
        Exit Function
    End If

    If Val(lblSellSeat.Caption) > 0 Then
        szMsg = "车次已售或留票" & lblSellSeat.Caption & "张" & Chr(10)
    Else
        szMsg = ""
    End If

    nResult = MsgBox(szMsg & "是否用车辆[" & m_oVehicle.LicenseTag & "]顶班车次[" & Trim(txtBusID.Text) & "]", vbQuestion + vbYesNo + vbDefaultButton2, "环境")
    If nResult = vbNo Then Exit Function
    SetBusy
    m_oReBus.Identify txtBusID.Text, m_dtBusDate
    If m_oReBus.HaveLugge = True Then
        '如果被顶班车次已受理行包
        If MsgBox("顶班车次已受理行包是否顶班?", vbQuestion + vbYesNo + vbDefaultButton2, "环境") = vbNo Then Exit Function
    End If
    m_oReBus.ReBusSlipLock False


NextReplaceVehicle:

    '下面开始进行顶班操作
    ShowSBInfo "车次顶班......."
    If chkOutoSeatType.Value = 1 Then
        m_oReBus.ReplaceVehicle ResolveDisplay(txtVehicle.Text), chkAutoTicketPrice.Value, chkAutoSeat.Value, chkAllRefundment.Value
    Else
        m_oReBus.ReplaceVehicle ResolveDisplay(txtVehicle.Text), chkAutoTicketPrice.Value, chkAutoSeat.Value, chkAllRefundment.Value
    End If

    If m_bIsParent Then
        frmEnvBus.UpdateList m_oReBus.BusID, m_oReBus.RunDate
    End If
    MsgBox "顶班成功!", vbInformation, Me.Caption
    SetNormal
    Unload Me
    Exit Function
ErrorHandle:
    szReferBusid = m_oReBus.ReferenceBusID
    If err.Number = ERR_REBusNewVehicleNowRoutePrice Then
        '如果发现错误为:该车次车辆的车型无票价，而顶班时要求自动更新票价，结果顶班失败。
        '则进行生成票价
        szPriceTable = m_oReBus.GetPriceTable(m_dtBusDate)
        
        
        tVehicleSeatType = oBaseInfo.GetAllVehicleSeatTypeInfo(ResolveDisplay(txtVehicle.Text))
        nCount = ArrayLength(tVehicleSeatType)
        For i = 1 To nCount
            nCount2 = ArrayLength(tBusVehcileSeatInfo)
            j = 1
            If nCount2 = 0 Then
                ReDim tBusVehcileSeatInfo(1 To 1)
                tBusVehcileSeatInfo(1).szBusID = szReferBusid
                tBusVehcileSeatInfo(1).szSeatTypeID = tVehicleSeatType(i).szSeatTypeID
                tBusVehcileSeatInfo(1).szSeatTypeName = tVehicleSeatType(i).szSeatTypeName
                tBusVehcileSeatInfo(1).szVehicleID = ResolveDisplay(txtVehicle.Text)
                tBusVehcileSeatInfo(1).szVehicleTypeCode = ResolveDisplay(tVehicleSeatType(i).szVehcileTypeName)
            Else

                Do While Not tBusVehcileSeatInfo(j).szSeatTypeID = tVehicleSeatType(i).szSeatTypeID
                    j = j + 1
                    If j > nCount2 Then
                        bflgIsNot = True
                        ReDim Preserve tBusVehcileSeatInfo(1 To nCount2 + 1)
                        tBusVehcileSeatInfo(nCount2 + 1).szBusID = szReferBusid
                        tBusVehcileSeatInfo(nCount2 + 1).szSeatTypeID = tVehicleSeatType(i).szSeatTypeID
                        tBusVehcileSeatInfo(nCount2 + 1).szSeatTypeName = tVehicleSeatType(i).szSeatTypeName
                        tBusVehcileSeatInfo(nCount2 + 1).szVehicleID = ResolveDisplay(txtVehicle.Text)
                        tBusVehcileSeatInfo(nCount2 + 1).szVehicleTypeCode = ResolveDisplay(tVehicleSeatType(i).szVehcileTypeName)
                        Exit Do
                    End If
                Loop
            End If
        Next
        If nCount = 0 Then
            ReDim tBusVehcileSeatInfo(1 To 1)
            tBusVehcileSeatInfo(nCount2 + 1).szBusID = szReferBusid
            tBusVehcileSeatInfo(nCount2 + 1).szSeatTypeID = "01"
            tBusVehcileSeatInfo(nCount2 + 1).szVehicleID = ResolveDisplay(txtVehicle.Text)
            tBusVehcileSeatInfo(nCount2 + 1).szVehicleTypeCode = m_oVehicle.VehicleModel 'ResolveDisplay(tVehicleSeatType(i).szVehcileTypeName)
        End If
        ShowSBInfo "计算车次票价......."
        Set oRoutePrice = New RoutePriceTable
        oRoutePrice.Init g_oActiveUser
        oRoutePrice.Identify szPriceTable
        oRoutePrice.MakeBusPrice tBusVehcileSeatInfo, True, False
        If m_szMakePriceIsNotSuccess <> "" Then
            Dim msgErr As String
            msgErr = "顶班时以车次[" & szReferBusid & "]为参考生成车次票价，" & Chr(10) & "但是该车次" & m_szMakePriceIsNotSuccess & Chr(10) & " 顶班失败 !!"
            MsgBox msgErr, vbInformation, Me.Caption
            SetNormal
            Exit Function
        End If
        GoTo NextReplaceVehicle
    Else
        ShowErrorMsg
        m_oReBus.ReBusSlipLock False
        SetNormal
    End If
'
End Function


