VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmCheckBusInfo 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "班次检票信息"
   ClientHeight    =   6435
   ClientLeft      =   2850
   ClientTop       =   2790
   ClientWidth     =   7230
   HelpContextID   =   2000030
   Icon            =   "frmCheckBusInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "Modal"
   Begin VB.ComboBox cboCheckSheet 
      Height          =   315
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   49
      Top             =   2850
      Width           =   1665
   End
   Begin RTComctl3.CoolButton cmdCheckSheet 
      Height          =   315
      Left            =   4500
      TabIndex        =   1
      Top             =   2850
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "路单(&S)"
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
      MICON           =   "frmCheckBusInfo.frx":014A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton imbCheckList 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   2850
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "检票明细(&D)>>"
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
      MICON           =   "frmCheckBusInfo.frx":0166
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   -1  'True
      VALUE           =   0   'False
   End
   Begin VB.Frame fraList 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   3105
      Left            =   60
      TabIndex        =   48
      Top             =   3270
      Width           =   7035
      Begin MSComctlLib.ListView lvCheckList 
         Height          =   2985
         Left            =   0
         TabIndex        =   4
         Top             =   105
         Width           =   5625
         _ExtentX        =   9922
         _ExtentY        =   5265
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "票号"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "检入方式"
            Object.Width           =   2645
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "检票时间"
            Object.Width           =   2540
         EndProperty
      End
      Begin RTComctl3.CoolButton cmdTicketInfo 
         Height          =   345
         Left            =   5730
         TabIndex        =   5
         Top             =   90
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "车票信息(&T)"
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
         MICON           =   "frmCheckBusInfo.frx":0182
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.Timer tmStart 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   2610
      Left            =   135
      TabIndex        =   2
      Top             =   105
      Width           =   6960
      Begin VB.Image Image1 
         Height          =   720
         Left            =   60
         Picture         =   "frmCheckBusInfo.frx":019E
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车次序号:"
         Height          =   180
         Left            =   2940
         TabIndex        =   47
         Top             =   330
         Width           =   810
      End
      Begin VB.Label lblSerialNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   3765
         TabIndex        =   46
         Top             =   330
         Width           =   90
      End
      Begin VB.Label lblStartTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   3765
         TabIndex        =   45
         Top             =   630
         Width           =   90
      End
      Begin VB.Label lblChkStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   5745
         TabIndex        =   44
         Top             =   330
         Width           =   90
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   1695
         TabIndex        =   43
         Top             =   630
         Width           =   90
      End
      Begin VB.Label lblBus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   1695
         TabIndex        =   42
         Top             =   330
         Width           =   90
      End
      Begin VB.Label lblChangeCheckSum 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   5745
         TabIndex        =   41
         Top             =   2310
         Width           =   90
      End
      Begin VB.Label lblNoCheckSum 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   5745
         TabIndex        =   40
         Top             =   2070
         Width           =   90
      End
      Begin VB.Label lblCheckSum 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   5745
         TabIndex        =   39
         Top             =   1830
         Width           =   90
      End
      Begin VB.Label lblSellSum 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   5745
         TabIndex        =   38
         Top             =   1590
         Width           =   90
      End
      Begin VB.Label lblMergeBus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   5745
         TabIndex        =   37
         Top             =   1320
         Width           =   90
      End
      Begin VB.Label lblCheckSheetNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   3765
         TabIndex        =   36
         Top             =   2280
         Width           =   90
      End
      Begin VB.Label lblStopCheckTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   3765
         TabIndex        =   35
         Top             =   2040
         Width           =   90
      End
      Begin VB.Label lblStartCheckTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   3765
         TabIndex        =   34
         Top             =   1800
         Width           =   90
      End
      Begin VB.Label lblCheckor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   3765
         TabIndex        =   33
         Top             =   1560
         Width           =   90
      End
      Begin VB.Label lblCheckGate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   3765
         TabIndex        =   32
         Top             =   1320
         Width           =   90
      End
      Begin VB.Label lblLicense 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   1695
         TabIndex        =   31
         Top             =   2280
         Width           =   90
      End
      Begin VB.Label lblOwner 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   1695
         TabIndex        =   30
         Top             =   2040
         Width           =   90
      End
      Begin VB.Label lblVehicle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   1695
         TabIndex        =   29
         Top             =   1800
         Width           =   90
      End
      Begin VB.Label lblCompany 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   1695
         TabIndex        =   28
         Top             =   1560
         Width           =   90
      End
      Begin VB.Label lblRoute 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   1695
         TabIndex        =   27
         Top             =   1320
         Width           =   90
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "检票状态:"
         Height          =   180
         Left            =   4920
         TabIndex        =   26
         Top             =   330
         Width           =   810
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "检票信息"
         Height          =   180
         Left            =   2910
         TabIndex        =   25
         Top             =   960
         Width           =   720
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "基本信息"
         Height          =   180
         Left            =   855
         TabIndex        =   24
         Top             =   975
         Width           =   720
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFFFF&
         X1              =   2925
         X2              =   6645
         Y1              =   1215
         Y2              =   1215
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         X1              =   840
         X2              =   2700
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00808080&
         X1              =   2910
         X2              =   6630
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   840
         X2              =   2700
         Y1              =   1215
         Y2              =   1215
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车主:"
         Height          =   180
         Left            =   855
         TabIndex        =   23
         Top             =   2040
         Width           =   450
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车牌照:"
         Height          =   180
         Left            =   855
         TabIndex        =   22
         Top             =   2280
         Width           =   630
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车辆:"
         Height          =   180
         Left            =   855
         TabIndex        =   21
         Top             =   1800
         Width           =   450
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "参营公司:"
         Height          =   180
         Left            =   855
         TabIndex        =   20
         Top             =   1560
         Width           =   810
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "运行线路:"
         Height          =   180
         Left            =   855
         TabIndex        =   19
         Top             =   1320
         Width           =   810
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "并入车次:"
         Height          =   180
         Left            =   4920
         TabIndex        =   18
         Top             =   1320
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车次:"
         Height          =   180
         Left            =   855
         TabIndex        =   17
         Top             =   315
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "日期:"
         Height          =   180
         Left            =   840
         TabIndex        =   16
         Top             =   630
         Width           =   450
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "发车时间:"
         Height          =   180
         Left            =   2925
         TabIndex        =   15
         Top             =   645
         Width           =   810
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "售票数:"
         Height          =   180
         Left            =   4920
         TabIndex        =   14
         Top             =   1590
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "检票数:"
         Height          =   180
         Left            =   4920
         TabIndex        =   13
         Top             =   1830
         Width           =   630
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "未检数:"
         Height          =   180
         Left            =   4920
         TabIndex        =   12
         Top             =   2070
         Width           =   630
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "检票口:"
         Height          =   180
         Left            =   2910
         TabIndex        =   11
         Top             =   1320
         Width           =   630
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "检票员:"
         Height          =   180
         Left            =   2910
         TabIndex        =   10
         Top             =   1560
         Width           =   630
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "开检时间:"
         Height          =   180
         Left            =   2910
         TabIndex        =   9
         Top             =   1800
         Width           =   810
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "停检时间:"
         Height          =   180
         Left            =   2910
         TabIndex        =   8
         Top             =   2040
         Width           =   810
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "改乘数:"
         Height          =   180
         Left            =   4920
         TabIndex        =   7
         Top             =   2310
         Width           =   630
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "路单编号:"
         Height          =   180
         Left            =   2910
         TabIndex        =   6
         Top             =   2280
         Width           =   810
      End
   End
   Begin RTComctl3.CoolButton cmdOK 
      Cancel          =   -1  'True
      Default         =   -1  'True
      Height          =   315
      Left            =   5835
      TabIndex        =   3
      Top             =   2850
      Width           =   1275
      _ExtentX        =   2249
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
      MICON           =   "frmCheckBusInfo.frx":1068
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Line Line4 
      X1              =   5460
      X2              =   6660
      Y1              =   1935
      Y2              =   1935
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   5460
      X2              =   6660
      Y1              =   1950
      Y2              =   1950
   End
End
Attribute VB_Name = "frmCheckBusInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const MaxWinHeight = 7000
Private Const NorWinHeight = 3640
Private szSelectedTicket As String '列表框中当前选中的票号

Public g_oActiveUser As ActiveUser
Public mdtBusDate As Date
Public mszBusID As String
Public mnBusSerialNo As Integer
Public moChkTicket As CheckTicket      '检票对象

Private mtBusCheckInfo() As TBusCheckInfo     '车次检票信息
'填充检票简要信息
Private Sub FillCheckInfo()
    '得到检票信息
    On Error GoTo ErrorHandle
    mtBusCheckInfo = moChkTicket.GetBusCheckInfo(mdtBusDate, mszBusID, mnBusSerialNo)
    lblBus.Caption = mszBusID
    lblDate.Caption = IIf(mtBusCheckInfo(1).dtDate = cdtEmptyDate, "", Format(mtBusCheckInfo(1).dtDate, cszDateStr))
    lblStartTime.Caption = Format(mtBusCheckInfo(1).dtStartUpTime, "HH:mm")
    lblStartCheckTime.Caption = Format(mtBusCheckInfo(1).dtBeginCheckTime, cszTimeStr)
    lblStopCheckTime.Caption = IIf(mtBusCheckInfo(1).dtEndCheckTime = cdtEmptyDateTime, "", Format(mtBusCheckInfo(1).dtEndCheckTime, cszTimeStr))
    lblCheckSheetNo.Caption = mtBusCheckInfo(1).szCheckSheet
    
    '得到车次基本信息
    Dim oREBus As REBus
    Set oREBus = New REBus
    oREBus.Init g_oActiveUser
    oREBus.Identify mszBusID, mdtBusDate, g_tCheckInfo.CheckGateNo
    If mtBusCheckInfo(1).szCheckSheet <> "" And oREBus.BusType = TP_ScrollBus Then
        lblChkStatus.Caption = GetStatusString(EREBusStatus.ST_BusStopCheck)
    Else
        lblChkStatus.Caption = GetStatusString(oREBus.busStatus)
    End If
    lblRoute.Caption = oREBus.RouteName
    lblCheckGate.Caption = oREBus.CheckGate
    lblMergeBus.Caption = oREBus.BeMergedBus.szBusid
    If oREBus.BusType = TP_ScrollBus Then
        lblSerialNo.Caption = mnBusSerialNo
    Else
        lblSerialNo.Caption = ""
    End If
    
    '设置车辆信息
    Dim oVehicle As Vehicle
    Set oVehicle = New Vehicle
    oVehicle.Init g_oActiveUser
    oVehicle.Identify mtBusCheckInfo(1).szVehicleId
    lblCompany.Caption = oVehicle.CompanyName
    lblOwner.Caption = oVehicle.OwnerName
    lblVehicle.Caption = oVehicle.VehicleId
    lblLicense.Caption = oVehicle.LicenseTag
    
    Dim atTicketsInfo() As TCheckedTicketInfo      '票信息数组
    atTicketsInfo = moChkTicket.GetBusCheckTicket(mdtBusDate, mszBusID, mnBusSerialNo)
    Dim nChangeCount As Integer
    nChangeCount = 0
    Dim i As Integer
    For i = 1 To ArrayLength(atTicketsInfo)
        If atTicketsInfo(i).nCheckTicketType = ECheckedTicketStatus.ChangedTicket Then
            nChangeCount = nChangeCount + 1
        End If
    Next i
    '得到售票数及检票数等
    lblSellSum.Caption = oREBus.GetNotCanSellCount
    lblCheckSum.Caption = ArrayLength(atTicketsInfo)
    lblChangeCheckSum.Caption = nChangeCount
    lblNoCheckSum.Caption = IIf(Val(lblSellSum) + nChangeCount - Val(lblCheckSum) > 0, Val(lblSellSum) + nChangeCount - Val(lblCheckSum), 0)
    
    '得到检票员名称
'    Dim szChecker As String
'    If mtBusCheckInfo(1).szChecker = g_oActiveUser.UserID Then
'        szChecker = MakeDisplayString(mtBusCheckInfo(1).szChecker, g_oActiveUser.UserName)
'    Else
'        Dim aszUsers() As String
'        aszUsers = moChkTicket.GetAllUser
'        For i = 1 To ArrayLength(aszUsers)
'            If aszUsers(i, 1) = mtBusCheckInfo(1).szChecker Then
'                szChecker = MakeDisplayString(mtBusCheckInfo(1).szChecker, aszUsers(i, 2))
'                Exit For
'            End If
'        Next i
'    End If
'    lblCheckor.Caption = szChecker
    lblCheckor.Caption = mtBusCheckInfo(1).szChecker
    
    '填充检票明细
    Dim oListItem As ListItem
    lvCheckList.ListItems.Clear
    For i = 1 To ArrayLength(atTicketsInfo)
        Set oListItem = lvCheckList.ListItems.Add(, , Trim(atTicketsInfo(i).szTicketID))  '票号
        oListItem.SubItems(1) = getCheckedTicketStatus(atTicketsInfo(i).nCheckTicketType) '检入方式
        oListItem.SubItems(2) = Format(atTicketsInfo(i).dtCheckTime, cszTimeStr) '检入时间
    Next i
    
    If mtBusCheckInfo(1).szCheckSheet = "" Then imbCheckList.Enabled = False
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub



Private Sub cmdCheckSheet_Click()
    Dim ofrmSheet As New frmCheckSheet
    ofrmSheet.mbViewMode = True
    ofrmSheet.mbNoPrintPrompt = True
    Set ofrmSheet.moChkTicket = moChkTicket
    Set ofrmSheet.g_oActiveUser = g_oActiveUser
    ofrmSheet.mszSheetID = lblCheckSheetNo.Caption
    ofrmSheet.Show vbModal
    Set ofrmSheet = Nothing
End Sub

Private Sub cmdOk_Click()
    Unload Me
End Sub

Private Sub cmdTicketInfo_Click()
    If Not lvCheckList.SelectedItem Is Nothing Then
        Dim ofrmTicket As frmTicketInfo
        Set ofrmTicket = New frmTicketInfo
        Set ofrmTicket.g_oActiveUser = g_oActiveUser
        ofrmTicket.TicketID = lvCheckList.SelectedItem.Text
        ofrmTicket.Show vbModal
        Set ofrmTicket = Nothing
    End If
End Sub

Public Sub RefreshForm()
    FillCheckInfo
End Sub

Private Sub CoolButton1_Click()
DisplayHelp Me
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandle
    If moChkTicket Is Nothing Then
        Set moChkTicket = New CheckTicket
        moChkTicket.Init g_oActiveUser
    End If
    
    '得到路单号码 有配载的时候，列出两张路单
    '得到检票信息
    Dim i As Integer
    mtBusCheckInfo = moChkTicket.GetBusCheckInfo(mdtBusDate, mszBusID, mnBusSerialNo)
    For i = 1 To ArrayLength(mtBusCheckInfo)
        cboCheckSheet.AddItem mtBusCheckInfo(i).szCheckSheet
    Next
    If ArrayLength(mtBusCheckInfo) > 0 Then cboCheckSheet.ListIndex = 0
    
    Call imbCheckList_Click
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub


Private Sub imbCheckList_Click()
    If imbCheckList.Value = False Then
        Me.Height = Me.Height - fraList.Height
    Else
        Me.Height = Me.Height + fraList.Height
    End If
End Sub

Private Sub lvCheckList_DblClick()
    Call cmdTicketInfo_Click
End Sub

Private Sub tmStart_Timer()
On Error GoTo ErrHandle
    tmStart.Enabled = False
    FillCheckInfo
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Function GetStatusString(nStatus As Integer) As String
    Select Case nStatus
        Case EREBusStatus.ST_BusChecking
            GetStatusString = "正在检票"
        Case EREBusStatus.ST_BusExtraChecking
            GetStatusString = "正在补检"
        Case EREBusStatus.ST_BusMergeStopped
            GetStatusString = "并班停检"
        Case EREBusStatus.ST_BusNormal
            GetStatusString = "未检"
        Case EREBusStatus.ST_BusStopCheck
            GetStatusString = "停检"
        Case EREBusStatus.ST_BusExtraChecking
            GetStatusString = "正在补检"
        Case EREBusStatus.ST_BusStopped
            GetStatusString = "车次停班"
    End Select
End Function

Private Function getCheckedTicketStatus(nStatus As Integer) As String
    Select Case nStatus
        Case ECheckedTicketStatus.NormalTicket
            getCheckedTicketStatus = "正常检入"
        Case ECheckedTicketStatus.ChangedTicket
            getCheckedTicketStatus = "改乘检入"
        Case ECheckedTicketStatus.MergedTicket
            getCheckedTicketStatus = "并班检入"
    End Select
End Function




Private Sub cboCheckSheet_Click()
        If cboCheckSheet.Text = "" Then Exit Sub
        Dim i As Integer
        i = cboCheckSheet.ListIndex + 1
        lblBus.Caption = mszBusID
        lblDate.Caption = IIf(mtBusCheckInfo(i).dtDate = cdtEmptyDate, "", Format(mtBusCheckInfo(i).dtDate, cszDateStr))
        lblStartTime.Caption = Format(mtBusCheckInfo(i).dtStartUpTime, "HH:mm")
        lblStartCheckTime.Caption = Format(mtBusCheckInfo(i).dtBeginCheckTime, cszTimeStr)
        lblStopCheckTime.Caption = IIf(mtBusCheckInfo(i).dtEndCheckTime = cdtEmptyDateTime, "", Format(mtBusCheckInfo(i).dtEndCheckTime, cszTimeStr))
        lblCheckSheetNo.Caption = mtBusCheckInfo(i).szCheckSheet
        
        '得到车次基本信息
        Dim oREBus As REBus
        Set oREBus = New REBus
        oREBus.Init g_oActiveUser
        oREBus.Identify mszBusID, mdtBusDate
        If mtBusCheckInfo(i).szCheckSheet <> "" And oREBus.BusType = TP_ScrollBus Then
            lblChkStatus.Caption = GetStatusString(EREBusStatus.ST_BusStopCheck)
        Else
            lblChkStatus.Caption = GetStatusString(oREBus.busStatus)
        End If
        lblRoute.Caption = oREBus.RouteName
        lblCheckGate.Caption = oREBus.CheckGate
        lblMergeBus.Caption = oREBus.BeMergedBus.szBusid
        If oREBus.BusType = TP_ScrollBus Then
            lblSerialNo.Caption = mnBusSerialNo
        Else
            lblSerialNo.Caption = ""
        End If
        
        '设置车辆信息
        Dim oVehicle As Vehicle
        Set oVehicle = New Vehicle
        oVehicle.Init g_oActiveUser
        oVehicle.Identify mtBusCheckInfo(i).szVehicleId
        lblCompany.Caption = oVehicle.CompanyName
        lblOwner.Caption = oVehicle.OwnerName
        lblVehicle.Caption = oVehicle.VehicleId
        lblLicense.Caption = oVehicle.LicenseTag
End Sub
