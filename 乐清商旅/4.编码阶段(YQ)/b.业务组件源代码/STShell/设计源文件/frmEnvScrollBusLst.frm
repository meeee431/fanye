VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmEnvScrollBusLst 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "滚动车次列表"
   ClientHeight    =   5340
   ClientLeft      =   2490
   ClientTop       =   1875
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   7455
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   915
      Left            =   135
      TabIndex        =   4
      Top             =   60
      Width           =   7275
      Begin VB.Label lblBusDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2006-1-12"
         Height          =   180
         Left            =   2520
         TabIndex        =   14
         Top             =   240
         Width           =   810
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "日期:"
         Height          =   180
         Left            =   2010
         TabIndex        =   13
         Top             =   240
         Width           =   450
      End
      Begin VB.Label lblBusID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "11011"
         Height          =   180
         Left            =   855
         TabIndex        =   12
         Top             =   240
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车次:"
         Height          =   180
         Left            =   330
         TabIndex        =   11
         Top             =   240
         Width           =   450
      End
      Begin VB.Label lblBusCount 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "20"
         Height          =   180
         Left            =   4485
         TabIndex        =   10
         Top             =   555
         Width           =   180
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "班次数:"
         Height          =   180
         Left            =   3780
         TabIndex        =   9
         Top             =   555
         Width           =   630
      End
      Begin VB.Label lblCheckGate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         Height          =   180
         Left            =   2730
         TabIndex        =   8
         Top             =   555
         Width           =   90
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "检票口:"
         Height          =   180
         Left            =   2010
         TabIndex        =   7
         Top             =   555
         Width           =   630
      End
      Begin VB.Label lblEndStation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "杭州"
         Height          =   180
         Left            =   1005
         TabIndex        =   6
         Top             =   555
         Width           =   360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "终点站:"
         Height          =   180
         Left            =   330
         TabIndex        =   5
         Top             =   555
         Width           =   630
      End
   End
   Begin RTComctl3.CoolButton cmdClose 
      Cancel          =   -1  'True
      Height          =   390
      Left            =   6000
      TabIndex        =   3
      Top             =   4845
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   688
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
      MICON           =   "frmEnvScrollBusLst.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdCheckInfo 
      Height          =   390
      Left            =   4560
      TabIndex        =   2
      Top             =   4845
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   688
      BTYPE           =   3
      TX              =   "检票信息(&B)"
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
      MICON           =   "frmEnvScrollBusLst.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ListView lvBus 
      Height          =   3195
      Left            =   75
      TabIndex        =   1
      Top             =   1515
      Width           =   7350
      _ExtentX        =   12965
      _ExtentY        =   5636
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
      NumItems        =   0
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "检票车次列表(L):"
      Height          =   180
      Left            =   90
      TabIndex        =   0
      Top             =   1170
      Width           =   1440
   End
End
Attribute VB_Name = "frmEnvScrollBusLst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public m_oActiveUser As ActiveUser
Public m_szBusID As String
Public m_dtBusDate As Date


Const cnSerial = 0
Const cnBusStartTime = 1
Const cnVehicle = 2
Const cnOwner = 3
Const cnTransportCompany = 4
Const cnChecker = 5
Const cnStartCheckTime = 6
Const cnEndCheckTime = 7




Private Sub cmdCheckInfo_Click()
    ShowCheckInfo
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Init
    RefreshBusInfo
End Sub

Private Sub lvBus_DblClick()
    ShowCheckInfo
End Sub


Private Sub ShowCheckInfo()
    On Error GoTo ErrHandle
    Dim oFrmCheckInfo As New frmCheckBusInfo
    Set oFrmCheckInfo.g_oActiveUser = m_oActiveUser
    If Not lvBus.SelectedItem Is Nothing Then
        oFrmCheckInfo.mszBusID = lblBusID
        oFrmCheckInfo.mdtBusDate = lblBusDate
        oFrmCheckInfo.mnBusSerialNo = lvBus.SelectedItem.Text
        Load oFrmCheckInfo
        oFrmCheckInfo.Show vbModal

    End If
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub


Private Sub Init()
    lblBusID.Caption = m_szBusID
    lblBusDate.Caption = m_dtBusDate
    lblCheckGate.Caption = ""
    lblEndStation.Caption = ""
    lblBusCount.Caption = 0
    
    
    '初始化listview
    With lvBus
        .ColumnHeaders.Add , , "序号", 540
        .ColumnHeaders.Add , , "发车时间", 929
        .ColumnHeaders.Add , , "运行车辆", 900
        .ColumnHeaders.Add , , "车主", 945
        .ColumnHeaders.Add , , "参运公司", 900
        .ColumnHeaders.Add , , "检票员", 720
        .ColumnHeaders.Add , , "开检时间", 959
        .ColumnHeaders.Add , , "停检时间", 959
        
    End With
    
    
    
End Sub
Private Sub RefreshBusInfo()
    Dim oREBus As New REBus
    Dim i As Integer
    Dim nCount As Integer
    Dim oCheckTicket As New CheckTicket
    Dim rsTemp As Recordset
    Dim liTemp As ListItem
    On Error GoTo ErrorHandle
    
    oREBus.Init m_oActiveUser
    oREBus.Identify m_szBusID, m_dtBusDate
    lblCheckGate.Caption = oREBus.CheckGate
    lblEndStation.Caption = oREBus.EndStationName
    
    
    '填充检票车次列表
    oCheckTicket.Init m_oActiveUser
    'oCheckTicket.i
    Set rsTemp = oCheckTicket.GetBusCheckInfoRS(m_dtBusDate, m_szBusID, -1)
    nCount = rsTemp.RecordCount
    

    lvBus.ListItems.Clear
    For i = 1 To nCount
        Set liTemp = lvBus.ListItems.Add(, , FormatDbValue(rsTemp!bus_serial_no))
        With liTemp
            .SubItems(cnBusStartTime) = Format(FormatDbValue(rsTemp!check_end_time), "hh:mm")
            .SubItems(cnVehicle) = FormatDbValue(rsTemp!license_tag_no)
            .SubItems(cnOwner) = FormatDbValue(rsTemp!owner_name)
            .SubItems(cnTransportCompany) = FormatDbValue(rsTemp!transport_company_short_name)
            .SubItems(cnChecker) = FormatDbValue(rsTemp!Checker)
            .SubItems(cnStartCheckTime) = Format(FormatDbValue(rsTemp!check_start_time), "hh:mm")
            .SubItems(cnEndCheckTime) = Format(FormatDbValue(rsTemp!check_end_time), "hh:mm")
        End With
        rsTemp.MoveNext
    Next i
        
    lblBusCount.Caption = lvBus.ListItems.Count
    
    Exit Sub
ErrorHandle:
    ShowErrorMsg
    
End Sub






