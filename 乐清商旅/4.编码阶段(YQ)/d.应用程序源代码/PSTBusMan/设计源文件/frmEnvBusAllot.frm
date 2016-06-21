VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmEnvBusAllot 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "环境车次配载信息"
   ClientHeight    =   4995
   ClientLeft      =   3075
   ClientTop       =   3105
   ClientWidth     =   7620
   HelpContextID   =   10000200
   Icon            =   "frmEnvBusAllot.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   7620
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox ptTitle 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   795
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   7665
      TabIndex        =   6
      Top             =   0
      Width           =   7665
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   60
         Left            =   -45
         TabIndex        =   7
         Top             =   735
         Width           =   7875
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "环境车次配载列表(&L):"
         Height          =   180
         Left            =   120
         TabIndex        =   8
         Top             =   330
         Width           =   1800
      End
   End
   Begin VB.ComboBox cboSellStation 
      Height          =   300
      ItemData        =   "frmEnvBusAllot.frx":038A
      Left            =   645
      List            =   "frmEnvBusAllot.frx":0394
      TabIndex        =   3
      Top             =   4335
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.ComboBox cboTCount 
      Height          =   300
      ItemData        =   "frmEnvBusAllot.frx":03A6
      Left            =   1365
      List            =   "frmEnvBusAllot.frx":03B0
      TabIndex        =   2
      Top             =   4530
      Visible         =   0   'False
      Width           =   1170
   End
   Begin RTComctl3.CoolButton cmdClose 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   6015
      TabIndex        =   0
      Top             =   4500
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
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
      MICON           =   "frmEnvBusAllot.frx":03C2
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
      Height          =   315
      Left            =   405
      TabIndex        =   1
      Top             =   4500
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
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
      MICON           =   "frmEnvBusAllot.frx":03DE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin FText.asFlatTextBox txtCheckGate 
      Height          =   300
      Left            =   885
      TabIndex        =   4
      Top             =   4650
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
   End
   Begin MSComCtl2.DTPicker dtpStartupTime 
      Height          =   300
      Left            =   435
      TabIndex        =   5
      Top             =   4680
      Visible         =   0   'False
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "HH:mm"
      Format          =   123994115
      UpDown          =   -1  'True
      CurrentDate     =   36392
   End
   Begin RTComctl3.CoolButton cmdAdd 
      Height          =   315
      Left            =   2505
      TabIndex        =   9
      Top             =   4500
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "新增(&A)"
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
      MICON           =   "frmEnvBusAllot.frx":03FA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdDelete 
      Height          =   315
      Left            =   3675
      TabIndex        =   10
      Top             =   4500
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "删除(&D)"
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
      MICON           =   "frmEnvBusAllot.frx":0416
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdSave 
      Default         =   -1  'True
      Height          =   315
      Left            =   4845
      TabIndex        =   11
      Top             =   4500
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
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
      MICON           =   "frmEnvBusAllot.frx":0432
      PICN            =   "frmEnvBusAllot.frx":044E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgBus 
      Height          =   3300
      Left            =   -15
      TabIndex        =   12
      Top             =   825
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   5821
      _Version        =   393216
      Rows            =   3
      Cols            =   5
      BackColorFixed  =   14737632
      BackColorBkg    =   14737632
      ScrollBars      =   2
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   " 设置"
      Enabled         =   0   'False
      Height          =   1200
      Left            =   -135
      TabIndex        =   13
      Top             =   4215
      Width           =   9705
      Begin RTComctl3.CoolButton cmdSellStation 
         Height          =   375
         Left            =   7965
         TabIndex        =   14
         Top             =   270
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "售票点"
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
         MICON           =   "frmEnvBusAllot.frx":07E8
         PICN            =   "frmEnvBusAllot.frx":0804
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
End
Attribute VB_Name = "frmEnvBusAllot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===================================================
'Reamrk:新增或修改车次配载情况
'===================================================
Option Explicit

Const cnMargin = 15

Const cnBusID = 0
Const cnSellStationID = 1
Const cnCheckGateID = 2
Const cnStartupTime = 3
Const cnCanSellCount = 4
Const cnHasSellCount = 5
Const cnStatus = 6
Const cnCols = 7



Const cnStationBusID = 0
Const cnStationSellStationID = 1
Const cnStationCanSellCount = 2
Const cnStationHasSellCount = 3
Const cnStationCols = 4


Public m_bIsAllot As Boolean


Public m_szBusID As String
Public m_dtEnvDate As Date
Public m_eStatus As EFormStatus

Dim m_oBus As New REBus
Dim nCount As Integer
Dim m_atTemp() As TReBusAllotInfo




Private Sub cboSellStation_Change()
    With hfgBus
        If .Text = cboSellStation.Text Then Exit Sub
        .Text = cboSellStation.Text
        .CellForeColor = cvChangeColor
        cmdSave.Enabled = True
    End With
End Sub

Private Sub cboSellStation_Click()
    With hfgBus
        If .Text = cboSellStation.Text Then Exit Sub
        .Text = cboSellStation.Text
        .CellForeColor = cvChangeColor
        cmdSave.Enabled = True
    End With
End Sub

Private Sub cboTCount_Click()
        With hfgBus
        If .Text = cboTCount.Text Then Exit Sub
        .Text = cboTCount.Text
        .CellForeColor = cvChangeColor
        cmdSave.Enabled = True
    End With
End Sub

Private Sub cmdAdd_Click()
    '新增一行
    With hfgBus
        
        .Rows = .Rows + 1
        .RowHeight(.Rows - 1) = 300
        hfgBus.ColWidth(cnSellStationID) = 2400
    End With
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub


Private Sub cmdDelete_Click()
    '删除一行
    Dim i As Integer
    With hfgBus
        If .Row > 1 Then
            .RemoveItem .Row
        End If
           If cboSellStation.Visible = True Then
                cboSellStation.Visible = False
            End If
            If cboTCount.Visible = True Then
                cboTCount.Visible = False
            End If
            If txtCheckGate.Visible = True Then
                txtCheckGate.Visible = False
            End If
            If dtpStartupTime.Visible = True Then
                dtpStartupTime.Visible = False
            End If
    End With
End Sub

Private Sub cmdSave_Click()
    SaveToDB
End Sub



Private Sub SaveToDB()
    '保存新增或修改到数据库中
    Dim i As Integer
    Dim nCount As Integer
    On Error GoTo ErrHandle
    Dim atTemp() As TReBusAllotInfo
    
    nCount = hfgBus.Rows - 1
    If nCount > 0 Then ReDim atTemp(1 To nCount)
    If m_bIsAllot Then
        
        For i = 1 To nCount
            hfgBus.Row = i
            hfgBus.Col = cnBusID
            atTemp(i).szbusID = hfgBus.Text
            hfgBus.Col = cnSellStationID
            atTemp(i).szSellStationID = ResolveDisplay(hfgBus.Text)
            hfgBus.Col = cnCheckGateID
            atTemp(i).szCheckGateID = ResolveDisplay(hfgBus.Text)
            hfgBus.Col = cnStartupTime
            atTemp(i).dtRunTime = ToDBDate(m_dtEnvDate) & " " & hfgBus.Text
            hfgBus.Col = cnCanSellCount
        '    atTemp(i).nCanSellQuantity = hfgBus.Text
            Select Case hfgBus.Text
                Case "不限"
                    atTemp(i).nCanSellQuantity = -1
                Case "不可售"
                    atTemp(i).nCanSellQuantity = 0
                Case Else
                    If Val(hfgBus.Text) > 0 Then
                        atTemp(i).nCanSellQuantity = Val(hfgBus.Text)
                    Else
                        atTemp(i).nCanSellQuantity = 0
                    End If
            End Select
            atTemp(i).dyBusDate = m_dtEnvDate
        Next
        
        m_oBus.SaveAllot atTemp
    Else
            
        For i = 1 To nCount
            hfgBus.Row = i
            hfgBus.Col = cnStationBusID
            atTemp(i).szbusID = hfgBus.Text
            hfgBus.Col = cnStationSellStationID
            atTemp(i).szSellStationID = ResolveDisplay(hfgBus.Text)
            hfgBus.Col = cnStationCanSellCount
            Select Case hfgBus.Text
                Case "不限"
                    atTemp(i).nCanSellQuantity = -1
                Case "不可售"
                    atTemp(i).nCanSellQuantity = 0
                Case Else
                    If Val(hfgBus.Text) > 0 Then
                        atTemp(i).nCanSellQuantity = Val(hfgBus.Text)
                    Else
                        atTemp(i).nCanSellQuantity = 0
                    End If
            End Select
            atTemp(i).dyBusDate = m_dtEnvDate
        Next i
        
        m_oBus.SaveSellStationInfo atTemp
    End If
    Unload Me
    Exit Sub
ErrHandle:
    ShowErrorMsg

End Sub




Private Sub CoolButton1_Click()
DisplayHelp Me
End Sub

Private Sub dtpStartupTime_Change()
    With hfgBus
        If .Text = Format(dtpStartupTime.Value, "HH:mm") Then Exit Sub
        .Text = Format(dtpStartupTime.Value, "HH:mm")
        .CellForeColor = cvChangeColor
        cmdSave.Enabled = True
    End With
End Sub

Private Sub Form_Load()
    AlignFormPos Me
    If m_bIsAllot Then
        Me.Caption = "环境车次配载信息"
        lblCaption.Caption = "环境车次配载列表(&L):"
        
    Else
        Me.Caption = "环境车次售票点管理"
        lblCaption.Caption = "环境车次售票点列表(&L):"
    End If
    FillSellStation
    RefreshBusAllot
End Sub

Private Sub RefreshBusAllot()
    '得到该车次的配载信息
    
    '如果车次的配载信息长度为0
    Dim szSellStation As String
    Dim i As Integer
    On Error GoTo ErrHandle
    ShowSBInfo "正在获得车次的配载信息..."
    
    m_oBus.Init g_oActiveUser
    m_oBus.Identify m_szBusID, m_dtEnvDate
    hfgBus.Redraw = False

    cboTCount.AddItem "5张"
    cboTCount.AddItem "10张"
    cboTCount.AddItem "15张"
    cboTCount.AddItem "20张"
    cboTCount.AddItem "25张"
    cboTCount.AddItem "30张"
            
    
    If m_bIsAllot Then
        m_atTemp = m_oBus.GetAllotInfo()
        
        
        nCount = ArrayLength(m_atTemp)
        
        
        hfgBus.Cols = cnCols
        
        hfgBus.Rows = nCount + 1
        hfgBus.ColWidth(cnBusID) = 800
        hfgBus.ColWidth(cnSellStationID) = 1800
        hfgBus.TextMatrix(0, cnBusID) = "车次代码"
        hfgBus.TextMatrix(0, cnSellStationID) = "上车站"
        hfgBus.TextMatrix(0, cnCheckGateID) = "检票口"
        hfgBus.TextMatrix(0, cnStartupTime) = "发车时间"
        hfgBus.TextMatrix(0, cnCanSellCount) = "可售张数"
        hfgBus.TextMatrix(0, cnHasSellCount) = "已售张数"
        hfgBus.TextMatrix(0, cnStatus) = "车次状态"
        
        If nCount = 0 Then
            hfgBus.Rows = 2
            hfgBus.RowHeight(1) = 300
            'hfgBus.ColWidth(cnSellStationID) = 2400
        Else
            
            For i = 1 To nCount
                ShowSBInfo ""
                hfgBus.TextMatrix(i, cnBusID) = m_atTemp(i).szbusID
                hfgBus.TextMatrix(i, cnSellStationID) = MakeDisplayString(m_atTemp(i).szSellStationID, m_atTemp(i).szSellStationName)
                hfgBus.TextMatrix(i, cnCheckGateID) = MakeDisplayString(m_atTemp(i).szCheckGateID, m_atTemp(i).szCheckGateName)
                hfgBus.TextMatrix(i, cnStartupTime) = Format(m_atTemp(i).dtRunTime, "HH:mm")
                Select Case m_atTemp(i).nCanSellQuantity
                   Case Is < 0: hfgBus.TextMatrix(i, cnCanSellCount) = "不限": hfgBus.Col = cnCanSellCount: hfgBus.Row = i: hfgBus.CellForeColor = vbBlack
                   Case 0: hfgBus.TextMatrix(i, cnCanSellCount) = "不可售": hfgBus.Col = cnCanSellCount: hfgBus.Row = i: hfgBus.CellForeColor = vbGrayText
                   Case Else: hfgBus.TextMatrix(i, cnCanSellCount) = m_atTemp(i).nCanSellQuantity
                End Select
                hfgBus.TextMatrix(i, cnHasSellCount) = m_atTemp(i).nSellQuantity
                Select Case m_atTemp(i).nStatus
                Case ST_BusStopped
                    hfgBus.TextMatrix(i, cnStatus) = "停班"
                Case ST_BusChecking
                    hfgBus.TextMatrix(i, cnStatus) = "检票中"
                Case ST_BusExtraChecking
                    hfgBus.TextMatrix(i, cnStatus) = "补检"
                Case ST_BusStopCheck
                    hfgBus.TextMatrix(i, cnStatus) = "已检"
                Case ST_BusReplace
                    hfgBus.TextMatrix(i, cnStatus) = "顶班"
                Case ST_BusSlitpStop
                    hfgBus.TextMatrix(i, cnStatus) = "并班"
                Case Else
                    hfgBus.TextMatrix(i, cnStatus) = "正常"
                End Select
            Next i
        End If
    Else
        
        m_atTemp = m_oBus.GetSellStationInfo()
        
        
        nCount = ArrayLength(m_atTemp)
        
        
        hfgBus.Cols = cnStationCols
        
        hfgBus.Rows = nCount + 1
        hfgBus.ColWidth(cnStationBusID) = 800
        hfgBus.ColWidth(cnStationSellStationID) = 1800
        hfgBus.TextMatrix(0, cnStationBusID) = "车次代码"
        hfgBus.TextMatrix(0, cnStationSellStationID) = "上车站"
        hfgBus.TextMatrix(0, cnStationCanSellCount) = "可售张数"
        hfgBus.TextMatrix(0, cnStationHasSellCount) = "已售张数"
        
        If nCount = 0 Then
            hfgBus.Rows = 1
            hfgBus.RowHeight(0) = 300
            'hfgBus.ColWidth(cnSellStationID) = 2400
        Else
            
            For i = 1 To nCount
                ShowSBInfo ""
                hfgBus.TextMatrix(i, cnStationBusID) = m_atTemp(i).szbusID
                hfgBus.TextMatrix(i, cnStationSellStationID) = MakeDisplayString(m_atTemp(i).szSellStationID, m_atTemp(i).szSellStationName)
                Select Case m_atTemp(i).nCanSellQuantity
                   Case Is < 0: hfgBus.TextMatrix(i, cnStationCanSellCount) = "不限": hfgBus.Col = cnStationCanSellCount: hfgBus.Row = i: hfgBus.CellForeColor = vbBlack
                   Case 0: hfgBus.TextMatrix(i, cnStationCanSellCount) = "不可售": hfgBus.Col = cnStationCanSellCount: hfgBus.Row = i: hfgBus.CellForeColor = vbGrayText
                   Case Else: hfgBus.TextMatrix(i, cnStationCanSellCount) = m_atTemp(i).nCanSellQuantity
                End Select
                hfgBus.TextMatrix(i, cnStationHasSellCount) = m_atTemp(i).nSellQuantity
            Next i
        End If
        '补上原先某些售票点未设置的售票点
        
        Dim oSysMan As New SystemMan
        Dim atSellStationInfo() As TDepartmentInfo
        Dim nSellStationCount As Integer
        Dim nRow As Integer
        Dim j As Integer
        oSysMan.Init g_oActiveUser
        atSellStationInfo = oSysMan.GetAllSellStation(g_oActiveUser.UserUnitID)
        nSellStationCount = ArrayLength(atSellStationInfo)
        For i = 1 To nSellStationCount
            For j = 1 To nCount
                If atSellStationInfo(i).szSellStationID = m_atTemp(j).szSellStationID Then
                    Exit For
                End If
            Next j
            If j > nCount Then
                '如果未找到
                hfgBus.Rows = hfgBus.Rows + 1
                nRow = hfgBus.Rows - 1
                hfgBus.TextMatrix(nRow, cnStationBusID) = m_szBusID
                
                hfgBus.TextMatrix(nRow, cnStationSellStationID) = MakeDisplayString(atSellStationInfo(i).szSellStationID, atSellStationInfo(i).szSellStationName)
                hfgBus.TextMatrix(nRow, cnStationCanSellCount) = "不可售"
                hfgBus.TextMatrix(i, cnStationHasSellCount) = 0
                hfgBus.Col = cnStationCanSellCount
                hfgBus.Row = nRow
                hfgBus.CellForeColor = vbGrayText

            End If
        Next i
        If hfgBus.Rows > 1 Then hfgBus.FixedRows = 1
    End If
    hfgBus.Redraw = True
    ShowSBInfo ""
    Exit Sub
ErrHandle:
    ShowSBInfo ""
    ShowErrorMsg

    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveFormPos Me
End Sub

Private Sub hfgBus_Click()
    SetInVisibled
    With hfgBus
        If m_bIsAllot Then
            Select Case .Col
            Case cnSellStationID
                cboSellStation.Width = .CellWidth
                cboSellStation.Top = .Top + .CellTop - cnMargin
                cboSellStation.Left = .Left + .CellLeft
                cboSellStation.Visible = True
                cboSellStation.Text = .Text
                cboSellStation.SetFocus
            Case cnCheckGateID
                txtCheckGate.Width = .CellWidth
                txtCheckGate.Top = .Top + .CellTop - cnMargin
                txtCheckGate.Left = .Left + .CellLeft
                txtCheckGate.Visible = True
                txtCheckGate.Text = .Text
                txtCheckGate.SetFocus
            Case cnStartupTime
                dtpStartupTime.Width = .CellWidth
                dtpStartupTime.Top = .Top + .CellTop - cnMargin
                dtpStartupTime.Left = .Left + .CellLeft
                dtpStartupTime.Visible = True
                If IsDate(.Text) Then
                    dtpStartupTime.Value = .Text
                End If
                dtpStartupTime.SetFocus
            
            Case cnCanSellCount
                cboTCount.Width = .CellWidth
                cboTCount.Top = .Top + .CellTop - cnMargin
                cboTCount.Left = .Left + .CellLeft
                cboTCount.Visible = True
                cboTCount.Text = .Text
                cboTCount.SetFocus
            End Select
        Else
            Select Case .Col
            Case cnStationSellStationID
                cboSellStation.Width = .CellWidth
                cboSellStation.Top = .Top + .CellTop - cnMargin
                cboSellStation.Left = .Left + .CellLeft
                cboSellStation.Visible = True
                cboSellStation.Text = .Text
                cboSellStation.SetFocus
            Case cnStationCanSellCount
                cboTCount.Width = .CellWidth
                cboTCount.Top = .Top + .CellTop - cnMargin
                cboTCount.Left = .Left + .CellLeft
                cboTCount.Visible = True
                cboTCount.Text = .Text
                cboTCount.SetFocus
            End Select
        End If
    End With
End Sub


Private Sub hfgBus_Scroll()
    SetInVisibled
End Sub

Private Sub SetInVisibled()
    '设置不可用
    cboSellStation.Visible = False
    txtCheckGate.Visible = False
    dtpStartupTime.Visible = False
    cboTCount.Visible = False
End Sub

'===================================================
'Modify Date：2002-11-19
'Author:fl
'Reamrk:填充售票站
'===================================================b

Private Sub FillSellStation()
    '填充售票站
    Dim nCount As Integer
    Dim i As Integer
    cboSellStation.Clear
    nCount = ArrayLength(g_atAllSellStation)
    For i = 1 To nCount
        cboSellStation.AddItem MakeDisplayString(g_atAllSellStation(i).szSellStationID, g_atAllSellStation(i).szSellStationName)
    Next i
    If nCount > 0 Then cboSellStation.ListIndex = 0
    
    '填充所有的售票站
End Sub

Private Sub cboTCount_Change()
''    FormatTextToNumeric cboTCount, False, False
     With hfgBus
        If .Text = cboTCount.Text Then Exit Sub
        .Text = cboTCount.Text
        .CellForeColor = cvChangeColor
        cmdSave.Enabled = True
    End With
End Sub

Private Sub txtCheckGate_ButtonClick()
    '选择检票口
    Dim oShell As New CommDialog
    Dim aszTmp() As String
    oShell.Init g_oActiveUser
    aszTmp = oShell.SelectCheckGate
    Set oShell = Nothing
    If ArrayLength(aszTmp) = 0 Then Exit Sub
    txtCheckGate.Text = MakeDisplayString(aszTmp(1, 1), aszTmp(1, 2))

End Sub

Private Sub txtCheckGate_Change()
    With hfgBus
        If .Text = txtCheckGate.Text Then Exit Sub
        .Text = txtCheckGate.Text
        .CellForeColor = cvChangeColor
        cmdSave.Enabled = True
    End With
End Sub


