VERSION 5.00
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmBusInfo 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "车次信息"
   ClientHeight    =   4260
   ClientLeft      =   3210
   ClientTop       =   2775
   ClientWidth     =   5610
   HelpContextID   =   4002201
   Icon            =   "frmBusInfo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   Tag             =   "Modal"
   Begin VB.PictureBox ptTitle 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   6615
      TabIndex        =   23
      Top             =   0
      Width           =   6615
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Height          =   60
         Left            =   -30
         TabIndex        =   24
         Top             =   660
         Width           =   6885
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车次信息:"
         Height          =   180
         Left            =   120
         TabIndex        =   25
         Top             =   270
         Width           =   810
      End
   End
   Begin VB.Timer tmStart 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   2625
      Left            =   420
      TabIndex        =   0
      Top             =   780
      Width           =   4815
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车次代码:"
         Height          =   180
         Left            =   360
         TabIndex        =   38
         Top             =   240
         Width           =   810
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "总座位数:"
         Height          =   180
         Left            =   360
         TabIndex        =   37
         Top             =   1014
         Width           =   810
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车型名称:"
         Height          =   180
         Left            =   360
         TabIndex        =   36
         Top             =   498
         Width           =   810
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "参运公司:"
         Height          =   180
         Left            =   360
         TabIndex        =   35
         Top             =   2310
         Width           =   810
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车次种类:"
         Height          =   180
         Left            =   360
         TabIndex        =   34
         Top             =   1530
         Width           =   810
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "检票口:"
         Height          =   180
         Left            =   360
         TabIndex        =   33
         Top             =   756
         Width           =   630
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "全额退票:"
         Height          =   180
         Left            =   360
         TabIndex        =   32
         Top             =   1788
         Width           =   810
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "总站票数:"
         Height          =   180
         Left            =   360
         TabIndex        =   31
         Top             =   1272
         Width           =   810
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "运行线路:"
         Height          =   180
         Left            =   360
         TabIndex        =   30
         Top             =   2040
         Width           =   810
      End
      Begin VB.Label lblBusDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1999-10-01"
         Height          =   180
         Left            =   3300
         TabIndex        =   29
         Top             =   240
         Width           =   900
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "日期:"
         Height          =   180
         Left            =   2280
         TabIndex        =   28
         Top             =   240
         Width           =   450
      End
      Begin VB.Label lblBusID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1111"
         Height          =   180
         Left            =   1230
         TabIndex        =   27
         Top             =   240
         Width           =   360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "发车时间:"
         Height          =   180
         Left            =   2280
         TabIndex        =   22
         Top             =   750
         Width           =   810
      End
      Begin VB.Label lblRoute 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "9999"
         Height          =   180
         Left            =   1230
         TabIndex        =   21
         Top             =   2046
         Width           =   360
      End
      Begin VB.Label lblStartTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "09：00：00"
         Height          =   180
         Left            =   3300
         TabIndex        =   20
         Top             =   750
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "可售座位数:"
         Height          =   180
         Left            =   2280
         TabIndex        =   19
         Top             =   1020
         Width           =   990
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "可售站票数:"
         Height          =   180
         Left            =   2280
         TabIndex        =   18
         Top             =   1275
         Width           =   990
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "状态:"
         Height          =   180
         Left            =   2280
         TabIndex        =   17
         Top             =   1530
         Width           =   450
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "终点站:"
         Height          =   180
         Left            =   2280
         TabIndex        =   16
         Top             =   1785
         Width           =   630
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车主:"
         Height          =   180
         Left            =   2280
         TabIndex        =   15
         Top             =   2040
         Width           =   450
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车辆代码:"
         Height          =   180
         Left            =   2280
         TabIndex        =   14
         Top             =   495
         Width           =   810
      End
      Begin VB.Label lblTypeName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "大巴"
         Height          =   180
         Left            =   1230
         TabIndex        =   13
         Top             =   498
         Width           =   360
      End
      Begin VB.Label lblCheckGate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "01"
         Height          =   180
         Left            =   1230
         TabIndex        =   12
         Top             =   756
         Width           =   180
      End
      Begin VB.Label lblTotalSeat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "50"
         Height          =   180
         Left            =   1230
         TabIndex        =   11
         Top             =   1014
         Width           =   180
      End
      Begin VB.Label lblSaleSeat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "40"
         Height          =   180
         Left            =   3300
         TabIndex        =   10
         Top             =   1020
         Width           =   180
      End
      Begin VB.Label lblTotalStand 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         Height          =   180
         Left            =   1230
         TabIndex        =   9
         Top             =   1272
         Width           =   180
      End
      Begin VB.Label lblSaleStand 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         Height          =   180
         Left            =   3300
         TabIndex        =   8
         Top             =   1275
         Width           =   180
      End
      Begin VB.Label lblBusType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "固定"
         Height          =   180
         Left            =   1230
         TabIndex        =   7
         Top             =   1530
         Width           =   360
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "停班"
         Height          =   180
         Left            =   3300
         TabIndex        =   6
         Top             =   1530
         Width           =   360
      End
      Begin VB.Label lblAllRefundment 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "否"
         Height          =   180
         Left            =   1230
         TabIndex        =   5
         Top             =   1788
         Width           =   180
      End
      Begin VB.Label lblEndStationName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "重庆"
         Height          =   180
         Left            =   3300
         TabIndex        =   4
         Top             =   1785
         Width           =   360
      End
      Begin VB.Label lblVehicleID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "川A 12345"
         Height          =   180
         Left            =   3300
         TabIndex        =   3
         Top             =   495
         Width           =   810
      End
      Begin VB.Label lblOwner 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0123"
         Height          =   180
         Left            =   3300
         TabIndex        =   2
         Top             =   2040
         Width           =   360
      End
      Begin VB.Label lblCompany 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0321"
         Height          =   180
         Left            =   1230
         TabIndex        =   1
         Top             =   2310
         Width           =   360
      End
   End
   Begin RTComctl3.CoolButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   4290
      TabIndex        =   39
      Top             =   3780
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "关闭"
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
      MICON           =   "frmBusInfo.frx":038A
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
      Height          =   780
      Left            =   -120
      TabIndex        =   26
      Top             =   3540
      Width           =   8745
   End
End
Attribute VB_Name = "frmBusInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mbIsLoaded As Boolean

Dim m_oUser As ActiveUser
Dim m_szBusID As String
Dim m_dBusDate As Date

Dim oREBus As New REBus
'Dim oChkTk As New CheckTicket

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Public Property Get SelfUser() As ActiveUser
    Set SelfUser = m_oUser
End Property

Public Property Let SelfUser(ByVal NewUser As ActiveUser)
    Set m_oUser = NewUser
End Property

Public Property Get BusID() As String
    BusID = m_szBusID
End Property

Public Property Let BusID(ByVal NewBusID As String)
    m_szBusID = NewBusID
End Property

Public Property Get BusDate() As Date
    BusDate = m_dBusDate
End Property

Public Property Let BusDate(ByVal NewBusDate As Date)
    m_dBusDate = NewBusDate
End Property

Private Sub Form_Load()
    RefreshForm
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mbIsLoaded = False
End Sub


Public Property Get IsLoaded() As Boolean
    IsLoaded = mbIsLoaded
End Property
Public Sub RefreshForm()

    oREBus.Init m_oUser
    oREBus.Identify m_szBusID, m_dBusDate

    Dim oVehicle As New Vehicle
    Dim oRoute As New Route
    oVehicle.Init m_oUser
    oRoute.Init m_oUser
    oVehicle.Identify Trim(oREBus.Vehicle)
    oRoute.Identify Trim(oREBus.Route)

    lblBusID.Caption = m_szBusID
    lblBusDate.Caption = Format(m_dBusDate, "YYYY-MM-DD")
    lblRoute.Caption = oRoute.RouteName
    lblCheckGate.Caption = oREBus.CheckGate
    lblStartTime.Caption = Format(oREBus.StartUpTime, "hh:mm:ss")
    lblTotalSeat.Caption = oREBus.TotalSeat
    lblSaleSeat.Caption = oREBus.SaleSeat
    lblTotalStand.Caption = oREBus.TotalStandSeat
    lblSaleStand.Caption = oREBus.SaleStandSeat
    Dim oChkTicket As CheckTicket
    Set oChkTicket = New CheckTicket
    oChkTicket.Init m_oUser
    lblBusType.Caption = oChkTicket.GetBusTypeName(oREBus.BusType)   '   IIf(oREBus.BusType = TP_RegularBus, "固定", "流水")
    lblAllRefundment.Caption = IIf(oREBus.AllRefundment, "是", "否")
    lblEndStationName.Caption = oREBus.EndStationName


    lblTypeName.Caption = oVehicle.VehicleModelName
    lblVehicleID.Caption = oVehicle.LicenseTag
    lblOwner.Caption = oVehicle.OwnerName
    lblCompany.Caption = oVehicle.CompanyName



    If oREBus.BusType <> TP_ScrollBus Then
        Select Case oREBus.busStatus
            Case EREBusStatus.ST_BusNormal
                lblStatus.Caption = "未检"
            Case EREBusStatus.ST_BusChecking
                lblStatus.Caption = "正检"
            Case EREBusStatus.ST_BusStopCheck
                lblStatus.Caption = "停检"
            Case Else
                lblStatus.Caption = "停班"
        End Select
    Else
        lblStatus.Caption = "无效"
    End If
    Set oRoute = Nothing
    Set oVehicle = Nothing
End Sub


