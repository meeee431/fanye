VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form FrmCheckExtraTicket 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "非当日车次检票"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10005
   HelpContextID   =   40000290
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   10005
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox txtBusID 
      Height          =   375
      Left            =   4440
      TabIndex        =   15
      Top             =   1470
      Width           =   1695
   End
   Begin RTComctl3.TextButtonBox txtVehicle 
      Height          =   375
      Left            =   1200
      TabIndex        =   14
      Top             =   1470
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   661
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
   Begin VB.TextBox txtBusCheck 
      Height          =   375
      Left            =   8280
      TabIndex        =   12
      Top             =   5400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox chkCheckChange 
      BackColor       =   &H00E0E0E0&
      Height          =   405
      HelpContextID   =   4000211
      Left            =   8280
      Picture         =   "FrmCheckExtraTicket.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "点击切换至改乘检入模式"
      Top             =   4920
      Visible         =   0   'False
      Width           =   1200
   End
   Begin MSComCtl2.DTPicker DTPStartDate 
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   22872065
      CurrentDate     =   38910
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   30
      Left            =   -120
      TabIndex        =   0
      Top             =   820
      Width           =   10000
   End
   Begin MSComCtl2.DTPicker DtpEndDate 
      Height          =   375
      Left            =   4440
      TabIndex        =   4
      Top             =   960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   22872065
      CurrentDate     =   38910
   End
   Begin MSComctlLib.ListView lvSeat 
      Height          =   2655
      Left            =   120
      TabIndex        =   7
      Top             =   3960
      Width           =   7500
      _ExtentX        =   13229
      _ExtentY        =   4683
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imgSeat"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "座位号"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "座位种类"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "状态"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "票号"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "到站"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "票价"
         Object.Width           =   1764
      EndProperty
   End
   Begin MSComctlLib.ListView lvBus 
      Height          =   1770
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   9780
      _ExtentX        =   17251
      _ExtentY        =   3122
      SortKey         =   3
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imlBusIcon"
      ForeColor       =   -2147483640
      BackColor       =   16777215
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
      NumItems        =   14
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "车次"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "车次种类"
         Object.Width           =   141
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "日期"
         Object.Width           =   1889
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "发车时间"
         Text            =   "发车时间"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "运行线路"
         Object.Width           =   3281
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "检票口"
         Object.Width           =   865
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "车牌"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "车型"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "全额退票"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "状态"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "终到站"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "总座位"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "余座"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "隶属公司"
         Object.Width           =   2540
      EndProperty
   End
   Begin RTComctl3.CoolButton cmdQuery 
      Height          =   375
      Left            =   6390
      TabIndex        =   16
      Top             =   990
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "查询(&Q)"
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
      MICON           =   "FrmCheckExtraTicket.frx":014A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdClose 
      Height          =   375
      Left            =   8250
      TabIndex        =   17
      Top             =   990
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "关闭(&E)"
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
      MICON           =   "FrmCheckExtraTicket.frx":0166
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdCheck 
      Height          =   375
      Left            =   8250
      TabIndex        =   18
      Top             =   3960
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "检票(&C)"
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
      MICON           =   "FrmCheckExtraTicket.frx":0182
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdMakeSheet 
      Height          =   375
      Left            =   8250
      TabIndex        =   19
      Top             =   4470
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "生成路单(&M)"
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
      MICON           =   "FrmCheckExtraTicket.frx":019E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdHelp 
      Height          =   375
      Left            =   8250
      TabIndex        =   20
      Top             =   6150
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   661
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
      MICON           =   "FrmCheckExtraTicket.frx":01BA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "用于检入非当日车次或已检车次的车票，对于已经参与结算的车票，不能补检。"
      ForeColor       =   &H8000000D&
      Height          =   180
      Left            =   420
      TabIndex        =   13
      Top             =   360
      Width           =   6300
   End
   Begin VB.Label lblBus 
      BackStyle       =   0  'Transparent
      Caption         =   "未检车次(&B)"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label lblSeatInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "座位信息:"
      Height          =   180
      Left            =   120
      TabIndex        =   8
      Top             =   3720
      Width           =   810
   End
   Begin VB.Label lblBusID 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车次(&B):"
      Height          =   180
      Left            =   3360
      TabIndex        =   6
      Top             =   1560
      Width           =   720
   End
   Begin VB.Label lblVehicle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车辆(&V):"
      Height          =   180
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   720
   End
   Begin VB.Label lblEndDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "结束日期(&E):"
      Height          =   180
      Left            =   3360
      TabIndex        =   3
      Top             =   1080
      Width           =   1080
   End
   Begin VB.Label lblStartDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "开始日期(&S):"
      Height          =   180
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   1080
   End
   Begin VB.Image imgBG 
      Appearance      =   0  'Flat
      Height          =   2385
      Left            =   7800
      Picture         =   "FrmCheckExtraTicket.frx":01D6
      Top             =   4440
      Visible         =   0   'False
      Width           =   2205
   End
End
Attribute VB_Name = "FrmCheckExtraTicket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const cnBusID = 0 '车次代码
Const cnDateTime = 1 '日期          发车时间
Const cnRoute = 2 '运行线路
Const cnEndStation = 3 '终到站
Const cnLicenseTag = 4 '车牌
Const cnCompany = 5 '隶属公司
Const cnSplitCompany = 6 '拆帐公司
Const cnVehicleType = 7 '车型
'Const cnAddStatus = 8 '是否加班

Const cnBusType = 8 '车次种类
Const cnStatus = 9 '状态
Const cnTotalSeats = 10 '总座位
Const cnSaleSeatQuantity = 11 '余座
'Const cnBusKind = 13 '是否滚动车次


Private m_oChkTicket  As New CheckTicket
Private m_nCheckStatus As EREBusCheckStatus '车次状态
Private m_szBusID As String '车次
Private m_dtBusDate As Date '日期
Private m_nBusSerialNo As Integer '车次序号
Private m_nBusKind As Integer '车次种类
Private m_bMakeSheetID As Boolean '是否打印路单

Private Sub chkCheckChange_Click()
On Error GoTo ErrorHandle
    If chkCheckChange.Value = vbChecked Then
        MsgboxEx "当前处于改乘模式,允许检入其它车次的车票!", vbExclamation + vbOKOnly
        chkCheckChange.ToolTipText = "点击切换至正常检入方式"
        chkCheckChange.BackColor = &HFF&
        txtBusCheck.Visible = True
    Else
        chkCheckChange.ToolTipText = "点击切换至改乘检入方式"
        chkCheckChange.BackColor = &HE0E0E0
        txtBusCheck.Visible = False
    End If
    
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

'检入此票
Private Sub cmdCheck_Click()
On Error GoTo ErrorHandle
Dim oREBus As New REBus
Dim i As Integer
Dim szVehicleId As String

Dim TTicketInfo As TInterfaceCheckTicket
            
    m_szBusID = txtBusCheck.Text
    m_dtBusDate = Format(lvBus.SelectedItem.ListSubItems(cnDateTime).Text, cszDateStr)
    m_nBusKind = ResolveDisplay(lvBus.SelectedItem.ListSubItems(cnBusType).Text)
    szVehicleId = ResolveDisplay(lvBus.SelectedItem.ListSubItems(cnLicenseTag).Text)
    
    oREBus.Init g_oActiveUser
    oREBus.Identify m_szBusID, m_dtBusDate
        
    m_nCheckStatus = oREBus.busStatus '车次状态
    
    If m_nCheckStatus = ST_BusNormal Then '未检
        If m_nBusKind = TP_ScrollBus Then
            m_nBusSerialNo = m_oChkTicket.GetNextScrollNo(m_szBusID)
            m_oChkTicket.StartCheckScrollBus m_szBusID, m_nBusSerialNo, szVehicleId, m_dtBusDate
        Else
            m_nBusSerialNo = 0
            m_oChkTicket.StartCheckRegularBus m_szBusID, szVehicleId, m_dtBusDate
        End If
    Else
        m_oChkTicket.ExtraStartCheckBus m_szBusID, m_nBusSerialNo, m_dtBusDate, True
    End If
    For i = 1 To lvSeat.ListItems.Count
        If lvSeat.ListItems(i).Selected = True And Trim(lvSeat.ListItems(i).ListSubItems(2).Text) = "未检" Then
        
            TTicketInfo = g_oChkTicket.GetOneTicketInfo(Trim(lvSeat.ListItems(i).ListSubItems(3).Text))
            If Trim(TTicketInfo.SellStationID) = "cm" And (Trim(g_oActiveUser.SellStationID) = "yh" Or Trim(g_oActiveUser.SellStationID) = "km") Then
                MsgBox "楚门的车票不允许检入！", vbExclamation, Me.Caption
                Exit Sub
            End If
            
            m_oChkTicket.CheckTicket m_szBusID, m_nBusSerialNo, Trim(lvSeat.ListItems(i).ListSubItems(3).Text), m_dtBusDate, 1
            lvSeat.ListItems(i).SubItems(2) = "已检"
            cmdMakeSheet.Enabled = True
        End If
    Next i
    FillSeat m_dtBusDate, m_szBusID, True
    Exit Sub
    
ErrorHandle:
    ShowSBInfo ""
    ShowErrorMsg
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
DisplayHelp Me
End Sub

Private Sub cmdMakeSheet_Click()
       
    On Error GoTo ErrHandle
        
    DoEvents
    '停检
    g_oChkTicket.StopCheckTicket m_szBusID, m_nBusSerialNo, False, m_dtBusDate
    
    '生成路单
    ShowSBInfo "正在生成路单..."
    
    If CreateSheet Then
        '显示路单表单
        
        FillSeat m_dtBusDate, m_szBusID, False
        
        Dim ofrmTmp As frmCheckSheet
        Set ofrmTmp = New frmCheckSheet
        Set ofrmTmp.g_oActiveUser = g_oActiveUser
        Set ofrmTmp.moChkTicket = g_oChkTicket
        ofrmTmp.mszSheetID = g_tCheckInfo.CurrSheetNo
        ofrmTmp.mbExitAfterPrint = True
        
        ofrmTmp.Show vbModal
    End If
    
    g_tCheckInfo.CurrSheetNo = Right(NumAdd(g_tCheckInfo.CurrSheetNo, 1), g_nCheckSheetLen)            '预置下一个路单号，不监测错误
    g_tCheckInfo.CheckSheet = g_tCheckInfo.CurrSheetNo
    
    WriteInitReg
    WriteCheckGateInfo
        
    Unload Me
    
    Exit Sub
ErrHandle:
    ShowSBInfo ""
    ShowErrorMsg
End Sub

'停检处理
Public Sub StopCheck(Optional StopCheckMode As Integer = 1)
    '显示停检对话框
'    Dim m_szBusid As String
'    Dim nIndex As Integer
'    Dim tTmpBusInfo As tCheckBusLstInfo
    

End Sub

Private Sub cmdQuery_Click()
On Error GoTo ErrHandle

    Dim aszBus() As String
    
    If CheckTicketCheckStatus Then
       MsgBox "选定的车次已有车票被补检，必须先生成路单", vbInformation, "提示："
       Exit Sub
    End If
    
    ShowSBInfo "正在查询指定的车次信息，请等待..."
    SetBusy
        
    cmdCheck.Enabled = False
    cmdMakeSheet.Enabled = False
        
    aszBus = m_oChkTicket.GetUnCheckBusInfo(DTPStartDate.Value, DtpEndDate.Value, Trim(ResolveDisplay(txtVehicle.Text)), Trim(txtBusID.Text))
    FillBusItem aszBus
        
    SetNormal
    Exit Sub
ErrHandle:
'    WriteProcessBar  False
    ShowSBInfo ""
    SetNormal
    ShowErrorMsg
    
End Sub

Private Sub DtpEndDate_Change()
    If DtpEndDate.Value >= Date Then DtpEndDate.Value = DateAdd("d", -3, Date)
    If DtpEndDate.Value < DateAdd("d", -3, Date) Then DtpEndDate.Value = DateAdd("d", -3, Date)
End Sub

Private Sub DTPStartDate_Change()
    If DTPStartDate.Value >= Date Then DTPStartDate.Value = DateAdd("d", -3, Date)
    If DTPStartDate.Value < DateAdd("d", -3, Date) Then DTPStartDate.Value = DateAdd("d", -3, Date)
End Sub

Private Sub DtpEndDate_Click()
    If DtpEndDate.Value >= Date Then DtpEndDate.Value = DateAdd("d", -3, Date)
    If DtpEndDate.Value < DateAdd("d", -3, Date) Then DtpEndDate.Value = DateAdd("d", -3, Date)
End Sub

Private Sub DTPStartDate_Click()
    If DTPStartDate.Value >= Date Then DTPStartDate.Value = DateAdd("d", -3, Date)
    If DTPStartDate.Value < DateAdd("d", -3, Date) Then DTPStartDate.Value = DateAdd("d", -3, Date)
End Sub

Private Sub Form_Load()
    InitLv
    
    AlignHeadWidth Me.name, lvBus
    AlignHeadWidth Me.name, lvSeat
        
    m_oChkTicket.Init g_oActiveUser
    m_oChkTicket.CheckGateNo = g_tCheckInfo.CheckGateNo
    m_oChkTicket.InitSystemParam g_oActiveUser, False, g_bAllowChangeRide
    cmdCheck.Enabled = False
    cmdMakeSheet.Enabled = False
    txtBusCheck.Visible = False
    DTPStartDate.Value = DateAdd("d", -1, Date)
    DtpEndDate.Value = DateAdd("d", -1, Date)
End Sub

'增加一个车次数组信息至Listview中
Private Sub FillBusItem(aszBus() As String)
    'pbIsUpdate  是否是更新,默认是新增
    
    '返回列表项的索引
    Dim i As Integer
    Dim oListItem As ListItem
    Dim szStopDateAndStartDateMsg As String
    Dim eStatus As EREBusStatus
    Dim nCount As Integer
    
    lvBus.ListItems.Clear
    nCount = ArrayLength(aszBus)
    If nCount = 0 Then Exit Sub
    For i = 1 To nCount
        Set oListItem = lvBus.ListItems.Add(, , aszBus(i, 1))
        
        oListItem.SubItems(cnBusType) = MakeDisplayString(aszBus(i, 8), aszBus(i, 14))
                
        oListItem.SubItems(cnDateTime) = aszBus(i, 2)
        oListItem.SubItems(cnRoute) = Trim(aszBus(i, 3))
        oListItem.SubItems(cnLicenseTag) = MakeDisplayString(Trim(aszBus(i, 16)), Trim(aszBus(i, 5)))
        oListItem.SubItems(cnVehicleType) = Trim(aszBus(i, 6))
        
        eStatus = Val(aszBus(i, 7))
        If eStatus = ST_BusStopped Or eStatus = ST_BusMergeStopped Or eStatus = ST_BusSlitpStop Then
            oListItem.SubItems(cnStatus) = "停班"
            oListItem.ListSubItems(cnStatus).ForeColor = vbRed
        Else
            oListItem.SubItems(cnStatus) = "运行"
        End If
        
        oListItem.SubItems(cnEndStation) = aszBus(i, 10)
        oListItem.SubItems(cnTotalSeats) = aszBus(i, 11)
        oListItem.SubItems(cnCompany) = aszBus(i, 12)
        oListItem.SubItems(cnSaleSeatQuantity) = aszBus(i, 15)
        
        oListItem.SubItems(cnSplitCompany) = aszBus(i, 13)
'        oListItem.SubItems(cnAddStatus) = IIf(Val(aszBus(i, 17)) = 0, "否", "是")
        
'        oListItem.SubItems(cnBusKind) = aszBus(i, 16)
        oListItem.Selected = False
    Next i
    If nCount > 1 Then
        lvBus.ListItems(1).Selected = True
        lvBus.ListItems(1).EnsureVisible
                
        FillSeat Format(lvBus.ListItems(1).ListSubItems(1).Text, cszDateStr), lvBus.ListItems(1).Text, False
    Else
        For i = 1 To lvBus.ListItems.Count
            lvBus.ListItems(i).Selected = False
        Next i
        oListItem.Selected = True
    End If
    
    ShowSBInfo ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '保存列头
    SaveHeadWidth Me.name, lvBus
    SaveHeadWidth Me.name, lvSeat
    
    If CheckTicketCheckStatus Then
       MsgBox "选定的车次已有车票被补检，必须先生成路单", vbInformation, "提示："
       Cancel = True
    End If
End Sub

Private Sub lvBus_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If CheckTicketCheckStatus Then
       MsgBox "选定的车次已有车票被补检，必须先生成路单", vbInformation, "提示："
       Exit Sub
    End If
    FillSeat Format(Item.ListSubItems(1).Text, cszDateStr), Item.Text, False
End Sub



Private Sub lvSeat_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If lvSeat.ListItems.Count > 0 Then
       cmdCheck.Enabled = True
    End If
End Sub


Private Sub txtBusCheck_GotFocus()
    txtBusCheck.SelStart = 0
    txtBusCheck.SelLength = Len(txtBusCheck.Text)
End Sub

Private Sub txtVehicle_Click()
    Dim oShell As New CommDialog
    Dim aszTemp() As String
    
    On Error GoTo ErrorHandle
    oShell.Init g_oActiveUser
    aszTemp = oShell.SelectVehicleEX(False)
             
    If ArrayLength(aszTemp) Then txtVehicle.Text = MakeDisplayString(aszTemp(1, 1), aszTemp(1, 2))
    
    Set oShell = Nothing
        
    Exit Sub
ErrorHandle:
    Set oShell = Nothing
    ShowErrorMsg
End Sub


'填充座位信息
Public Sub FillSeat(BusDate As Date, BusID As String, Optional RefreshSeat As Boolean)
    Dim tSeat() As TSeatInfoEx
    Dim nCount As Integer
    Dim liTemp As ListItem
    Dim i As Integer
    Dim nSlitp As Integer, nReplace As Integer
    Dim szSeatInfo As String
    Dim nSeatCount As Integer
    Dim oREBus As New REBus

    On Error GoTo ErrHandle
    MousePointer = vbHourglass
    
    If BusID <> "" Then
        ShowSBInfo "获得座位信息..."
        oREBus.Init g_oActiveUser
        oREBus.Identify BusID, BusDate
            
        m_nCheckStatus = oREBus.busStatus '车次状态

    End If
    If RefreshSeat = False Then
       lvSeat.ListItems.Clear
       tSeat = oREBus.GetSeatInfo
        nCount = ArrayLength(tSeat)
        For i = 1 To nCount
            If (tSeat(i).szSeatStatus = ST_SeatSold Or tSeat(i).szSeatStatus = ST_SeatReplace Or tSeat(i).szSeatStatus = ST_SeatSlitp) And ResolveDisplayEx(tSeat(i).szTicketNo) = tSeat(i).szTicketNo Then
                    Set liTemp = lvSeat.ListItems.Add(, , tSeat(i).szSeatNo)
                    liTemp.SubItems(1) = MakeDisplayString(Trim(tSeat(i).szSeatType), Trim(tSeat(i).szSeatTypeName))
                    liTemp.SubItems(2) = "未检"
                    liTemp.SubItems(3) = ResolveDisplay(tSeat(i).szTicketNo)
                    liTemp.SubItems(4) = tSeat(i).szDestName
                    liTemp.SubItems(5) = tSeat(i).szTicketPrice
            End If
            
            lblSeatInfo.Visible = True
        Next i
    End If
    txtBusCheck.Text = BusID
    Set oREBus = Nothing
    
    SetNormal
    Exit Sub
ErrHandle:
    Set oREBus = Nothing
    SetNormal
    ShowErrorMsg
End Sub



Private Sub InitLv()
    '初始化listview
    
    With lvBus
        
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "车次"
        .ColumnHeaders.Add , , "发车时间"
        .ColumnHeaders.Add , , "运行线路"
        .ColumnHeaders.Add , , "终到站"
        .ColumnHeaders.Add , , "车牌"
        .ColumnHeaders.Add , , "隶属公司"
        .ColumnHeaders.Add , , "拆帐公司"
        .ColumnHeaders.Add , , "车型"
'        .ColumnHeaders.Add , , "加班"
        .ColumnHeaders.Add , , "车次类型"
        .ColumnHeaders.Add , , "状态"
        .ColumnHeaders.Add , , "总座位"
        .ColumnHeaders.Add , , "余座"
'        .ColumnHeaders.Add , , "车次种类"
'        lvBus.ColumnHeaders(cnBusKind).Width = 0
        
    End With
    
End Sub

Private Function CreateSheet() As Boolean
'生成路单
    Dim tTmp As TCheckSheetInfo
    Dim szTempSheetID As String
    
On Error GoTo ErrHandle
    ShowSBInfo "正在生成路单..."
    Me.MousePointer = vbHourglass
     
    tTmp = g_oChkTicket.GetCheckSheetInfo(g_tCheckInfo.CurrSheetNo)
    
    '检查路单号是否已被使用
    While Not tTmp.szCheckSheet = ""
        MsgboxEx "此路单已存在,请修改当前路单号!", vbExclamation, g_cszTitle_Error
        frmChangeSheetNo.Show vbModal
        tTmp = g_oChkTicket.GetCheckSheetInfo(g_tCheckInfo.CurrSheetNo)
    Wend


    g_oChkTicket.MakeCheckSheet m_dtBusDate, m_szBusID, m_nBusSerialNo, g_tCheckInfo.CurrSheetNo

    
    ShowSBInfo ""
    Me.MousePointer = vbDefault
    
    CreateSheet = True
    
    Exit Function
ErrHandle:
    ShowSBInfo ""
    Me.MousePointer = vbDefault
    ShowErrorMsg
End Function

'检查选定的车次、选定的坐位是否有新检票，
'如果有，则必须先生成路单，否则不能更换车次或退出
Private Function CheckTicketCheckStatus() As Boolean
    Dim i As Integer
    
    For i = 1 To lvSeat.ListItems.Count
        If Trim(lvSeat.ListItems(i).ListSubItems(2).Text) = "已检" Then
           CheckTicketCheckStatus = True
           Exit Function
        End If
    Next i
End Function
