VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmEnvBusStation 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "环境-车次站点属性"
   ClientHeight    =   5790
   ClientLeft      =   1725
   ClientTop       =   2340
   ClientWidth     =   8790
   HelpContextID   =   2005801
   Icon            =   "frmEnvBusStation.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   8790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin RTComctl3.CoolButton cmdHelp 
      Height          =   375
      Left            =   7560
      TabIndex        =   18
      Top             =   2655
      Width           =   1110
      _ExtentX        =   1958
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
      MICON           =   "frmEnvBusStation.frx":000C
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
      Height          =   375
      Left            =   7560
      TabIndex        =   17
      Top             =   360
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "保存(&S)"
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
      MICON           =   "frmEnvBusStation.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdaddStation 
      Height          =   375
      Left            =   7560
      TabIndex        =   16
      Top             =   2175
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "新增站点"
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
      MICON           =   "frmEnvBusStation.frx":0044
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mfgPrice 
      Height          =   1680
      Left            =   315
      TabIndex        =   4
      Top             =   3420
      Width           =   7080
      _ExtentX        =   12488
      _ExtentY        =   2963
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Timer tmStart 
      Interval        =   500
      Left            =   1125
      Top             =   2580
   End
   Begin VB.ComboBox cboTime 
      Height          =   300
      ItemData        =   "frmEnvBusStation.frx":0060
      Left            =   1170
      List            =   "frmEnvBusStation.frx":0067
      TabIndex        =   1
      Top             =   2010
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.ComboBox cboTCount 
      Height          =   300
      ItemData        =   "frmEnvBusStation.frx":0071
      Left            =   1170
      List            =   "frmEnvBusStation.frx":0084
      TabIndex        =   0
      Top             =   1650
      Visible         =   0   'False
      Width           =   825
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mfgStation 
      Height          =   1620
      Left            =   315
      TabIndex        =   2
      Top             =   1395
      Width           =   7080
      _ExtentX        =   12488
      _ExtentY        =   2858
      _Version        =   393216
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin RTComctl3.TextButtonBox txtBusID 
      Height          =   300
      Left            =   1575
      TabIndex        =   7
      Top             =   135
      Width           =   2670
      _ExtentX        =   4710
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
   End
   Begin RTComctl3.CoolButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   7560
      TabIndex        =   19
      Top             =   810
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "关闭(&X)"
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
      MICON           =   "frmEnvBusStation.frx":00A7
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车次代码(&B):"
      Height          =   180
      Left            =   315
      TabIndex        =   15
      Top             =   195
      Width           =   1080
   End
   Begin VB.Label lblOffTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   180
      Left            =   5055
      TabIndex        =   14
      Top             =   600
      Width           =   1770
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "发车时间:"
      Height          =   180
      Left            =   4185
      TabIndex        =   13
      Top             =   615
      Width           =   810
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "运行线路: "
      Height          =   180
      Left            =   1635
      TabIndex        =   12
      Top             =   615
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "运行车辆 :"
      Height          =   195
      Left            =   1620
      TabIndex        =   11
      Top             =   855
      Width           =   810
   End
   Begin VB.Label lblVehicle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   180
      Left            =   2595
      TabIndex        =   10
      Top             =   840
      Width           =   1170
   End
   Begin VB.Label lblRoute 
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   2595
      TabIndex        =   9
      Top             =   615
      Width           =   1350
   End
   Begin VB.Label lblTotalSeat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "总座位数:"
      Height          =   180
      Left            =   4185
      TabIndex        =   8
      Top             =   855
      Width           =   810
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmEnvBusStation.frx":00C3
      ForeColor       =   &H00000000&
      Height          =   465
      Left            =   360
      TabIndex        =   6
      Top             =   5220
      Width           =   7050
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "站点票价(&T):"
      Height          =   195
      Left            =   315
      TabIndex        =   5
      Top             =   3150
      Width           =   1095
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "站点属性设定(&L):"
      Height          =   180
      Left            =   315
      TabIndex        =   3
      Top             =   1125
      Width           =   1440
   End
End
Attribute VB_Name = "frmEnvBusStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'**********************************************************
'* Source File Name:frmEnvBusStation.frm
'* Project Name:PSTBusMan.vbp
'* Engineer:
'* Data Generated:2005/12/23
'* Last Revision Date:2005/12/23
'* Brief Description:环境车次站点属性
'* Relational Document:
'**********************************************************
Option Explicit
Public m_dtRunDate As Date
Public m_szBusID As String
Private m_bAddStation As Boolean
Private m_szStationID As String
Private m_oREBus As New REBus
Private atTicketType() As TTicketType
Private m_oParSystem As New SystemParam
Private anCountType() As Integer

Private Sub cboTCount_Change()
    If mfgStation.Text = cboTCount.Text Then Exit Sub
    mfgStation.Col = 3
    mfgStation.Text = cboTCount.Text
    mfgStation.CellForeColor = cvChangeColor
    mfgStation.Col = 0
    mfgStation.CellForeColor = cvChangeColor
    mfgStation.Col = 1
    mfgStation.CellForeColor = cvChangeColor
    mfgStation.Col = 2
    mfgStation.CellForeColor = cvChangeColor
    mfgStation.Col = 3
    mfgStation.CellForeColor = cvChangeColor
    cmdok.Enabled = True
End Sub
Private Sub cboTCount_Click()
    If mfgStation.Text = cboTCount.Text Then Exit Sub
    mfgStation.Text = cboTCount.Text
    mfgStation.CellForeColor = cvChangeColor
    mfgStation.Col = 0
    mfgStation.CellForeColor = cvChangeColor
    mfgStation.Col = 1
    mfgStation.CellForeColor = cvChangeColor
    mfgStation.Col = 2
    mfgStation.CellForeColor = cvChangeColor
    mfgStation.Col = 3
    mfgStation.CellForeColor = cvChangeColor
    cmdok.Enabled = True
End Sub
Private Sub cboTime_Change()
    If mfgStation.Text = cboTime.Text Then Exit Sub
    mfgStation.Col = 4
    mfgStation.Text = cboTime.Text
    mfgStation.CellForeColor = cvChangeColor
    mfgStation.Col = 0
    mfgStation.CellForeColor = cvChangeColor
    mfgStation.Col = 1
    mfgStation.CellForeColor = cvChangeColor
    mfgStation.Col = 2
    mfgStation.CellForeColor = cvChangeColor
    mfgStation.Col = 3
    mfgStation.CellForeColor = cvChangeColor
    mfgStation.Col = 4
    mfgStation.CellForeColor = cvChangeColor
    cmdok.Enabled = True
End Sub
Private Sub cboTime_Click()
    If mfgStation.Text = cboTime.Text Then Exit Sub
    mfgStation.Text = cboTime.Text
    mfgStation.CellForeColor = cvChangeColor
    mfgStation.Col = 0
    mfgStation.CellForeColor = cvChangeColor
    mfgStation.Col = 1
    mfgStation.CellForeColor = cvChangeColor
    mfgStation.Col = 2
    mfgStation.CellForeColor = cvChangeColor
    mfgStation.Col = 3
    mfgStation.CellForeColor = cvChangeColor
    mfgStation.Col = 4
    mfgStation.CellForeColor = cvChangeColor
    cmdok.Enabled = True
End Sub
Private Sub cmdaddStation_Click()

  frmInsertReStation.m_szBusID = m_szBusID
  frmInsertReStation.m_dtRunDate = m_dtRunDate
  frmInsertReStation.Show vbModal
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    DisplayHelp Me
End Sub

Private Sub cmdOk_Click()
    If m_bAddStation = True Then
        mfgStation.ColSel = 2
        mfgStation.Sort = flexSortNumericAscending
    End If
    SaveDisk
End Sub

Private Sub Form_Activate()
    If txtBusID.Text = "" Then
        cmdaddStation.Enabled = False
    Else
        cmdaddStation.Enabled = True
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyEscape
        Unload Me
End Select
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim nCount As Integer
    Dim j As Integer
    Dim tTicketTypeInfo() As TTicketType
    
    On Error GoTo here
    tTicketTypeInfo = m_oParSystem.GetAllTicketType(1)
    MoveFreeTicket tTicketTypeInfo
    nCount = ArrayLength(atTicketType)
    
    mfgStation.Cols = 6
    mfgPrice.Cols = 4 + nCount - 1
    
    mfgStation.AllowUserResizing = flexResizeNone
    mfgStation.TextMatrix(0, 0) = "站点代码"
    mfgStation.TextMatrix(0, 1) = "站点名称"
    mfgStation.TextMatrix(0, 2) = "里程数"
    mfgStation.TextMatrix(0, 3) = "限售张数"
    mfgStation.TextMatrix(0, 4) = "该时间前不能售票"
    mfgStation.TextMatrix(0, 5) = "售票人数"
    
    mfgPrice.TextMatrix(0, 0) = "起点站"
    mfgPrice.TextMatrix(0, 1) = "座位类型"
    mfgPrice.TextMatrix(0, 2) = "站点名称"
'    mfgPrice.textmatrix(0,2) = "基本运价"
    ReDim anCountType(mfgPrice.Cols - 1)
    For i = 0 To nCount - 1
        mfgPrice.TextMatrix(0, 4 + i - 1) = atTicketType(i).szTicketTypename
        If atTicketType(i).nTicketTypeID > 3 Then
            anCountType(i + 3) = atTicketType(i).nTicketTypeID
        End If
    Next i
    mfgStation.ColWidth(2) = 800
    mfgStation.ColWidth(3) = 850
    mfgStation.ColWidth(4) = 1780
    mfgStation.ColWidth(5) = 800
    m_oREBus.Init g_oActiveUser
    If m_szBusID <> "" Then
        txtBusID.Text = m_szBusID
    Else
        m_dtRunDate = Date
    End If
    mfgStation.Row = 1
    Me.Caption = Me.Caption & "[" & Format(m_dtRunDate, "YYYY年MM月DD日") & "]"
    Exit Sub
here:
    ShowErrorMsg
End Sub

Private Sub FullBus()
    Dim tBusStation() As TREBusStationInfo
    Dim tBusStationTemp() As TREBusStationInfo
    Dim szaST() As String
    Dim i As Integer, nCount As Integer, j As Integer, nSCount As Integer, nTime As Integer, nPCont As Integer
    Dim szStationID As String
    Dim nTypeTicketNcount As Integer
    Dim nCols As Integer, nPCols As Integer
    Dim vLimitedSellTime As Variant
    On Error GoTo here
    If m_szBusID = "" Then Exit Sub
    SetBusy
    m_oREBus.Identify m_szBusID, m_dtRunDate
    lblRoute.Caption = m_oREBus.RouteName
    lblOffTime.Caption = Format(m_oREBus.StartUpTime, "HH:MM:SS")
    lblVehicle.Caption = m_oREBus.VehicleTag
    lblTotalSeat.Caption = "总座位数: " & m_oREBus.TotalSeat
    tBusStationTemp = m_oREBus.GetBusStation
    tBusStation = GetBusStation(tBusStationTemp) 'ljw
    szaST = m_oREBus.GetStationSeatInfo
    nCols = mfgStation.Cols
    nPCols = mfgPrice.Cols
    nCount = ArrayLength(tBusStation)
    nPCont = ArrayLength(tBusStationTemp)
    nSCount = ArrayLength(szaST)
    '填写表头
    mfgStation.Rows = nCount + 1
    nTime = DateDiff("h", Now, m_oREBus.StartUpTime)
    
    For i = 1 To nTime / 2
        cboTime.AddItem i * 2 & "小时"
    Next
    
    For i = 1 To nCount
        mfgStation.MergeRow(i) = True
        mfgStation.TextMatrix(i, 0) = tBusStation(i - 1).szStationID
        mfgStation.TextMatrix(i, 1) = tBusStation(i - 1).szStationName
        mfgStation.TextMatrix(i, 2) = tBusStation(i - 1).nMileage
        mfgStation.MergeCol(3) = False
        Select Case tBusStation(i - 1).nLimitedSellCount
            Case Is < 0: mfgStation.TextMatrix(i, 3) = "不限"
            Case 0: mfgStation.TextMatrix(i, 3) = "不可售"
            Case Else: mfgStation.TextMatrix(i, 3) = tBusStation(i - 1).nLimitedSellCount & "张"
        End Select
        mfgStation.MergeCol(4) = False
        Select Case tBusStation(i - 1).sgLimitedSellTime
            Case Is <= 0: mfgStation.TextMatrix(i, 4) = "不限"
            Case Else:
        
                If m_oREBus.BusType <> 1 Then
                    mfgStation.TextMatrix(i, 4) = TransferLimitedTime(CStr(tBusStation(i - 1).sgLimitedSellTime), m_oREBus.StartUpTime, True)
                Else
                    mfgStation.TextMatrix(i, 4) = TransferLimitedTime(CStr(tBusStation(i - 1).sgLimitedSellTime), m_oREBus.StartUpTime, False)
                End If
        End Select
        
        mfgStation.TextMatrix(i, 5) = 0
        
        
        For j = 1 To nSCount
            If Trim(szaST(j, 1)) = Trim(tBusStation(i - 1).szStationID) Then
                mfgStation.TextMatrix(i, 5) = szaST(j, 2)
            End If
        Next
    
    
    Next
    
    
    mfgStation.Redraw = True
    mfgStation.FixedRows = 1
    mfgStation.FixedCols = 2
    mfgPrice.Rows = nPCont + 1
    mfgPrice.FixedCols = 0
    
    
    For i = 1 To nPCont
        mfgPrice.MergeRow(i) = True
        mfgPrice.MergeCol(0) = True
        mfgPrice.TextMatrix(i, 0) = tBusStationTemp(i).szSellStationName
        mfgPrice.MergeCol(1) = True
        mfgPrice.TextMatrix(i, 1) = tBusStationTemp(i).szSeatTypeName
        mfgPrice.MergeCol(2) = True
        mfgPrice.TextMatrix(i, 2) = tBusStationTemp(i).szStationName
        mfgPrice.MergeCol(3) = True
        mfgPrice.TextMatrix(i, 3) = tBusStationTemp(i).sgFullPrice
        mfgPrice.MergeCol(4) = True
        mfgPrice.TextMatrix(i, 4) = tBusStationTemp(i).sgHalfPrice
        For j = 5 To nPCols - 1
            Select Case anCountType(j)
            Case TP_PreferentialTicket1
                mfgPrice.MergeCol(j) = True
                mfgPrice.TextMatrix(i, j) = tBusStationTemp(i).sgPreferentialPrice1
            Case TP_PreferentialTicket2
                mfgPrice.MergeCol(j) = True
                mfgPrice.TextMatrix(i, j) = tBusStationTemp(i).sgPreferentialPrice2
            Case TP_PreferentialTicket3
                mfgPrice.MergeCol(j) = True
                mfgPrice.TextMatrix(i, j) = tBusStationTemp(i).sgPreferentialPrice3
            End Select
        Next
        mfgPrice.MergeCells = flexMergeRestrictColumns
    Next
    mfgPrice.Redraw = True
    mfgPrice.FixedCols = 3
    SetNormal
    
    Exit Sub
here:
    SetNormal
    ShowErrorMsg
End Sub







Private Sub mfgStation_Click()
    Select Case mfgStation.Col
    Case 3
        cboTCount.Top = mfgStation.Top + mfgStation.CellTop - 40
        cboTCount.Left = mfgStation.Left + mfgStation.CellLeft - 20
        cboTCount.Visible = True
        cboTCount.Width = mfgStation.CellWidth
        cboTime.Visible = False
        cboTCount.Text = mfgStation.Text
        cboTCount.SetFocus
        Exit Sub
    Case 4
        cboTime.Top = mfgStation.Top + mfgStation.CellTop - 40
        cboTime.Left = mfgStation.Left + mfgStation.CellLeft - 20
        cboTime.Visible = True
        cboTime.Width = mfgStation.CellWidth
        cboTCount.Visible = False
        cboTime.Text = mfgStation.Text
        cboTime.SetFocus
        Exit Sub
    Case Else
        cboTCount.Visible = False
        cboTime.Visible = False
    End Select
End Sub

Private Sub mfgStation_Scroll()
    cboTCount.Visible = False
    cboTime.Visible = False
End Sub

Private Sub tmStart_Timer()
    FullBus
    tmStart.Enabled = False
End Sub
Private Sub txtBusID_Click()
    Dim oShell As New CommDialog
    Dim szaTemp() As String
    Dim nRowCount As Integer
    If m_bAddStation <> True Then
        oShell.Init g_oActiveUser
        szaTemp = oShell.SelectREBus(m_dtRunDate, False)
        Set oShell = Nothing
        If ArrayLength(szaTemp) = 0 Then Exit Sub
        txtBusID.Text = szaTemp(1, 1)
        m_szBusID = txtBusID.Text
        FullBus
    Else

    End If
End Sub

Private Sub txtBusID_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn
        m_szBusID = txtBusID.Text
        FullBus
    End Select
End Sub

Private Sub SaveDisk()
    Dim vbMsg As VbMsgBoxResult
    Dim nCols As Integer
    Dim tBusStation As TBusStationSellInfo
    Dim tBusTicketPrice As TRETicketPrice
    Dim i As Integer
    Dim nDateLimite As Integer
    Dim szStationSerial As Integer
    Dim szbusID As String
    Dim szStationID As String
    Dim szlimtTime As String
    Dim nLen As Integer
    On Error GoTo here
    
    If MsgBox("是否保存当前修改", vbQuestion + vbYesNo, "车次站点属性--插入站点") <> vbYes Then Exit Sub
    
    If DateDiff("d", m_dtRunDate, Now) > 0 Then MsgBox "开始日期小于当前日期，不能进行修改", vbExclamation, "站点属性修改": Exit Sub
    nCols = mfgStation.Cols
    mfgStation.Col = 0
    For i = 1 To mfgStation.Rows - 1
        mfgStation.Row = i
        mfgStation.Col = 0
        If mfgStation.CellForeColor = vbBlue Then
            mfgStation.Row = i
            
            tBusStation.szStationID = mfgStation.TextMatrix(i, 0)
            
            
            tBusStation.nMileage = Trim(mfgStation.TextMatrix(i, 2))
            
            If Trim(mfgStation.TextMatrix(i, 3)) = "不可售" Then
                tBusStation.nLimitedSellCount = 0
            End If
            
            
            If Trim(mfgStation.TextMatrix(i, 3)) = "不限" Then
                tBusStation.nLimitedSellCount = -1
            End If
            
            If Val(mfgStation.TextMatrix(i, 3)) > 0 Then
                tBusStation.nLimitedSellCount = Val(mfgStation.TextMatrix(i, 3))
            End If
            If Trim(mfgStation.TextMatrix(i, 4)) = "不限" Then
                tBusStation.sgLimitedSellTime = -1
            Else
            
                nLen = InStr(1, mfgStation.TextMatrix(i, 4), "小时")
                If nLen <> 0 Then
                    szlimtTime = Left(mfgStation.TextMatrix(i, 4), nLen - 1)
                Else
                    szlimtTime = mfgStation.TextMatrix(i, 4)
                End If
                mfgStation.Col = 4
                'If mfgStation.CellForeColor = vbBlue Then
                If IsNumeric(szlimtTime) = False And IsDate(szlimtTime) = False Then
                    MsgBox "输入有误。请按：A.B 小时格式输入。", vbExclamation, Me.Caption
                    Exit Sub
                End If
                If IsNumeric(szlimtTime) Then
                    szlimtTime = Format(szlimtTime, ".00")
                Else
                    szlimtTime = DateDiff("h", szlimtTime, m_dtRunDate & " " & lblOffTime)
                End If
            
                'tBusStation.nLimitedSellTime = DateDiff("h", CDate(mfgStation.textmatrix(i  + 4)), m_oREBus.StartupTime)
                tBusStation.sgLimitedSellTime = CSng(szlimtTime)
                'End If
            End If
            
            m_oREBus.ModifyBusStation tBusStation
            
            If tBusStation.sgLimitedSellTime <> -1 And tBusStation.sgLimitedSellTime <> 0 Then
                If m_oREBus.BusType = 1 Then
                    mfgStation.TextMatrix(i, 4) = TransferLimitedTime(CStr(tBusStation.sgLimitedSellTime), m_oREBus.StartUpTime, False)
                    cboTime.Text = mfgStation.TextMatrix(i, 4)
                Else
                    mfgStation.TextMatrix(i, 4) = TransferLimitedTime(CStr(tBusStation.sgLimitedSellTime), m_oREBus.StartUpTime, True)
                    cboTime.Text = mfgStation.TextMatrix(i, 4)
                End If
            End If
            
        End If
    Next
    cboTCount.Visible = False
    cboTime.Visible = False
    MsgBox "车次站点属性修改成功", vbInformation, "环境"
    cmdok.Enabled = False
    Exit Sub
here:
    ShowErrorMsg
End Sub
Public Function OpenTime()
    tmStart.Enabled = True
End Function
Private Function GetBusStation(tBusStationTemp() As TREBusStationInfo) As TREBusStationInfo()
    Dim i As Integer
    Dim j As Integer
    Dim nCount As Integer
    Dim nCountTemp As Integer
    Dim tBusStationTempBack() As TREBusStationInfo
    nCount = ArrayLength(tBusStationTemp)
    For i = 1 To nCount
        nCountTemp = ArrayLength(tBusStationTempBack)
        j = 0
        If nCountTemp <> 0 Then
            Do While tBusStationTempBack(j).szStationID <> tBusStationTemp(i).szStationID
                j = j + 1
                If j > nCountTemp - 1 Then
                    j = j - 1
                    Exit Do
                End If
            Loop
        Else
            ReDim tBusStationTempBack(0 To nCountTemp)
            tBusStationTempBack(nCountTemp).nLimitedSellCount = tBusStationTemp(i).nLimitedSellCount
            tBusStationTempBack(nCountTemp).sgLimitedSellTime = tBusStationTemp(i).sgLimitedSellTime
            tBusStationTempBack(nCountTemp).nMileage = tBusStationTemp(i).nMileage
            tBusStationTempBack(nCountTemp).szStationID = tBusStationTemp(i).szStationID
            tBusStationTempBack(nCountTemp).szStationName = tBusStationTemp(i).szStationName
            GoTo here
        End If
    If tBusStationTempBack(j).szStationID <> tBusStationTemp(i).szStationID Then
        ReDim Preserve tBusStationTempBack(0 To nCountTemp)
        tBusStationTempBack(nCountTemp).nLimitedSellCount = tBusStationTemp(i).nLimitedSellCount
        tBusStationTempBack(nCountTemp).sgLimitedSellTime = tBusStationTemp(i).sgLimitedSellTime
        tBusStationTempBack(nCountTemp).nMileage = tBusStationTemp(i).nMileage
        tBusStationTempBack(nCountTemp).szStationID = tBusStationTemp(i).szStationID
        tBusStationTempBack(nCountTemp).szStationName = tBusStationTemp(i).szStationName
    End If
here:
    Next
    GetBusStation = tBusStationTempBack
End Function
Private Function MoveFreeTicket(tTicketTypeInfo() As TTicketType)
    Dim i As Integer
    Dim nCount As Integer
    Dim nCountTemp As Integer
    Dim tTicketTypeInfoBack() As TTicketType
    nCount = ArrayLength(tTicketTypeInfo)
    For i = 1 To nCount
        If tTicketTypeInfo(i).nTicketTypeID <> TP_FreeTicket Then
            nCountTemp = nCountTemp + 1
            ReDim Preserve atTicketType(0 To nCountTemp - 1)
            atTicketType(nCountTemp - 1).nTicketTypeID = tTicketTypeInfo(i).nTicketTypeID
            atTicketType(nCountTemp - 1).szTicketTypename = tTicketTypeInfo(i).szTicketTypename
            atTicketType(nCountTemp - 1).szAnnotation = tTicketTypeInfo(i).szAnnotation
            atTicketType(nCountTemp - 1).nTicketTypeValid = tTicketTypeInfo(i).nTicketTypeValid
        End If
    Next
End Function
Private Function TransferLimitedTime(szLimitedSellTime As String, dtBusDate As Date, Optional pbIsRegular As Boolean = True) As String
    Dim vLimitedSellTime   As Variant
    If pbIsRegular Then
        vLimitedSellTime = TransferStopTime(CStr(szLimitedSellTime), True)
        TransferLimitedTime = Format(DateAdd("n", -CInt(vLimitedSellTime), dtBusDate), "YYYY-MM-DD HH:NN")
    Else
        TransferLimitedTime = CStr(TransferStopTime(szLimitedSellTime, False))
    End If
End Function

'转换停售时间
'/////////////////////////////////////////////
Public Function TransferStopTime(psgStopTime As String, Optional pbIsRegular As Boolean = True) As Variant
    Dim szPrefix As String
    Dim szSuffix As String
    Dim szStopTime As String
    Dim vTemp As Variant
    
    szStopTime = Format(psgStopTime, ".00")
    szPrefix = LeftAndRight(szStopTime, True, ".")
    szSuffix = LeftAndRight(szStopTime, False, ".")
    If szPrefix = "" Then
        szPrefix = "0"
    End If
    If pbIsRegular Then
        If Left(szSuffix, 1) <> "0" Then
            vTemp = CInt(szPrefix) * 60 + CInt(Val((Left(szSuffix, 1)))) * 10 + CInt(Val((Right(szSuffix, 1))))
        Else
            vTemp = CInt(szPrefix) * 60 + CInt(Val(szSuffix))
        End If
    Else
        vTemp = Format(szPrefix & ":" & szSuffix, "hh:mm")
    End If
    TransferStopTime = vTemp
End Function
