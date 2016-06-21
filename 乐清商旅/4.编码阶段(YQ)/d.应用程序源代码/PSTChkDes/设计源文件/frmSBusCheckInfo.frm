VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSBusCheckInfo 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "滚动车次检票信息"
   ClientHeight    =   4680
   ClientLeft      =   2160
   ClientTop       =   2655
   ClientWidth     =   6705
   Icon            =   "frmSBusCheckInfo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   6705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Tag             =   "Modal"
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   375
      TabIndex        =   5
      Top             =   525
      Width           =   4470
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1999/09/09"
         Height          =   180
         Left            =   810
         TabIndex        =   14
         Top             =   135
         Width           =   900
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "日期:"
         Height          =   180
         Left            =   240
         TabIndex        =   13
         Top             =   135
         Width           =   450
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "班"
         Height          =   180
         Left            =   1905
         TabIndex        =   12
         Top             =   375
         Width           =   180
      End
      Begin VB.Label lblTotalNum 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "12"
         Height          =   180
         Left            =   1425
         TabIndex        =   11
         Top             =   375
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "共检滚动车次:"
         Height          =   180
         Left            =   225
         TabIndex        =   10
         Top             =   375
         Width           =   1170
      End
      Begin VB.Label lblEndStation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "重庆"
         Height          =   180
         Left            =   2610
         TabIndex        =   9
         Top             =   135
         Width           =   600
      End
      Begin VB.Label lblCheckGate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "01"
         Height          =   180
         Left            =   4020
         TabIndex        =   8
         Top             =   135
         Width           =   180
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "检票口:"
         Height          =   195
         Left            =   3300
         TabIndex        =   7
         Top             =   135
         Width           =   675
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "终到站:"
         Height          =   180
         Left            =   1890
         TabIndex        =   6
         Top             =   135
         Width           =   630
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   315
      Left            =   5415
      TabIndex        =   4
      Top             =   4230
      Width           =   1185
   End
   Begin VB.CommandButton cmdCheckInfo 
      Caption         =   "班次检票信息(&B)"
      Height          =   315
      Left            =   3375
      TabIndex        =   3
      Top             =   4230
      Width           =   1845
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2565
      Left            =   90
      TabIndex        =   1
      Top             =   1575
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   4524
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "序号"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "发车时间"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "运行车辆"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "车主"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "参营公司"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "检票员"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "开检时间"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "停检时间"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Label lblBusID 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1122"
      Height          =   180
      Left            =   630
      TabIndex        =   15
      Top             =   210
      Width           =   360
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   6600
      Y1              =   495
      Y2              =   495
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   6600
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   6600
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   6615
      Y1              =   1185
      Y2              =   1185
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "检票车次列表(&L):"
      Height          =   255
      Left            =   105
      TabIndex        =   2
      Top             =   1335
      Width           =   1575
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车次:"
      Height          =   180
      Left            =   105
      TabIndex        =   0
      Top             =   210
      Width           =   450
   End
End
Attribute VB_Name = "frmSBusCheckInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_oUser As New ActiveUser '局部复制
Dim m_dBusDate As Date
Dim m_szBusID As String
Dim mbIsLoaded As Boolean


Dim nBusSerialNo As Integer

Private Sub cmdCheckInfo_Click()
    
    Dim oChkApp As New CheckSysApp
    oChkApp.ShowCheckInfo m_oUser, m_dBusDate, m_szBusID, nBusSerialNo
    Set oChkApp = Nothing
    
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    ShowSBInfo "正在读取滚动车次检票信息..."

    cmdCheckInfo.Enabled = False
On Error GoTo ErrHandle
    mbIsLoaded = True
    
    RefreshForm
    
    ShowSBInfo ""
    
    Exit Sub
ErrHandle:
    ShowErrorMsgLocal
    Unload Me
    ShowSBInfo ""
End Sub

Public Property Get BusDate() As Date
    BusDate = m_dBusDate
End Property

Public Property Let BusDate(ByVal NewBusDate As Date)
    m_dBusDate = NewBusDate
End Property

Public Property Get BusID() As String
    BusID = m_szBusID
End Property

Public Property Let BusID(ByVal NewBusID As String)
    m_szBusID = NewBusID
End Property

Public Property Get SelfUser() As ActiveUser
    Set SelfUser = m_oUser
End Property

Public Property Let SelfUser(ByVal NewUser As ActiveUser)
    Set m_oUser = NewUser
End Property

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    nBusSerialNo = CInt(Item.Text)
'    cmdCheckInfo.Enabled = True
End Sub
Public Sub RefreshForm()
    Dim oChkTk As New CheckTicket
    Dim oREBus As New REBus
    Dim oVehicle As New Vehicle
    Dim oOwner As New Owner
    Dim oCorp As New Company
    
    Dim nTotalNum As Integer
    Dim uCheckBus As TBusCheckInfo
    Dim nLoop As Integer

    oChkTk.Init m_oUser
    oREBus.Init m_oUser
    
        
    nTotalNum = oChkTk.GetNextScrollNo(m_szBusID, m_dBusDate)
    oREBus.Identify m_szBusID, m_dBusDate
    
    lblBusID.Caption = m_szBusID
    lblDate.Caption = m_dBusDate
    lblEndStation.Caption = oREBus.EndStation
    lblCheckGate.Caption = oREBus.CheckGate
    lblTotalNum.Caption = Str(nTotalNum - 1)
    
    ListView1.ListItems.Clear
    oVehicle.Init m_oUser
    oOwner.Init m_oUser
    oCorp.Init m_oUser
    
    For nLoop = 1 To nTotalNum - 1
        nBusSerialNo = nLoop
        uCheckBus = oChkTk.GetBusCheckInfo(m_dBusDate, m_szBusID, nBusSerialNo)
        ListView1.ListItems.Add , , uCheckBus.nBusSerialNo
        ListView1.ListItems(nLoop).SubItems(1) = Format(uCheckBus.dtStartUpTime, "hh:mm:ss")
        oVehicle.Identify Trim(uCheckBus.szVehicleId)
        ListView1.ListItems(nLoop).SubItems(2) = oVehicle.LicenseTag
        oOwner.Identify Trim(uCheckBus.szOwnerID)
        ListView1.ListItems(nLoop).SubItems(3) = oOwner.OwnerName
        oCorp.Identify Trim(uCheckBus.szCompanyID)
        ListView1.ListItems(nLoop).SubItems(4) = oCorp.CompanyShortName
        ListView1.ListItems(nLoop).SubItems(5) = Trim(uCheckBus.szChecker)
        ListView1.ListItems(nLoop).SubItems(6) = Format(uCheckBus.dtBeginCheckTime, "hh:mm:ss")
        ListView1.ListItems(nLoop).SubItems(7) = Format(uCheckBus.dtEndCheckTime, "hh:mm:ss")
    Next nLoop
    If ListView1.ListItems.Count > 0 Then
        cmdCheckInfo.Enabled = True
        ListView1.ListItems(1).Selected = True
        nBusSerialNo = CInt(ListView1.ListItems(1).Text)
    Else
        cmdCheckInfo.Enabled = False
    End If
    Set oVehicle = Nothing
    Set oOwner = Nothing
    Set oCorp = Nothing
    Set oChkTk = Nothing
    Set oREBus = Nothing
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mbIsLoaded = False
    Set m_oUser = Nothing
End Sub


Public Property Get IsLoaded() As Boolean
    IsLoaded = mbIsLoaded
End Property

