VERSION 5.00
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Object = "{D11DF87C-EFEC-4838-B7E2-15462BF136FB}#1.4#0"; "RTReportlf.ocx"
Begin VB.Form frmReportStation 
   BackColor       =   &H80000009&
   Caption         =   "站点车型票价查询"
   ClientHeight    =   6720
   ClientLeft      =   2625
   ClientTop       =   3075
   ClientWidth     =   8535
   HelpContextID   =   1002401
   Icon            =   "frmReportStation.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6720
   ScaleWidth      =   8535
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   3660
      Top             =   2895
   End
   Begin VB.PictureBox ptQuery 
      BackColor       =   &H80000009&
      Height          =   7905
      Left            =   15
      ScaleHeight     =   7845
      ScaleWidth      =   2790
      TabIndex        =   14
      Top             =   0
      Width           =   2850
      Begin VB.CheckBox chkSellStation 
         BackColor       =   &H00FFFFFF&
         Caption         =   "所有的上车站(&T)"
         Height          =   180
         Left            =   960
         TabIndex        =   23
         Top             =   5430
         Width           =   1710
      End
      Begin VB.ListBox lstSellStation 
         Height          =   780
         Left            =   120
         MultiSelect     =   2  'Extended
         TabIndex        =   22
         Top             =   5745
         Width           =   2535
      End
      Begin VB.ListBox lstStation 
         Height          =   1500
         ItemData        =   "frmReportStation.frx":014A
         Left            =   90
         List            =   "frmReportStation.frx":014C
         MultiSelect     =   2  'Extended
         TabIndex        =   3
         Top             =   1305
         Width           =   2565
      End
      Begin VB.ListBox lstType 
         Height          =   960
         Left            =   90
         MultiSelect     =   2  'Extended
         TabIndex        =   6
         Top             =   3180
         Width           =   2565
      End
      Begin VB.ComboBox cboPriceTable 
         Height          =   300
         Left            =   90
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   660
         Width           =   2565
      End
      Begin VB.CheckBox chkAllStation 
         BackColor       =   &H80000009&
         Caption         =   "所有站点(&T)"
         Height          =   240
         Left            =   990
         TabIndex        =   4
         Top             =   1035
         Width           =   1470
      End
      Begin RTComctl3.CoolButton cmdCancel 
         Height          =   330
         Left            =   1410
         TabIndex        =   12
         Top             =   6645
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   582
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
         MICON           =   "frmReportStation.frx":014E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin RTComctl3.CoolButton cmdQuery 
         Default         =   -1  'True
         Height          =   330
         Left            =   135
         TabIndex        =   11
         Top             =   6645
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   582
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
         MICON           =   "frmReportStation.frx":016A
         PICN            =   "frmReportStation.frx":0186
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.CheckBox chkAllType 
         BackColor       =   &H80000009&
         Caption         =   "所有车型(&V)"
         Height          =   240
         Left            =   990
         TabIndex        =   7
         Top             =   2880
         Width           =   1440
      End
      Begin VB.ListBox lstSeatType 
         Height          =   780
         Left            =   90
         MultiSelect     =   2  'Extended
         TabIndex        =   9
         Top             =   4500
         Width           =   2535
      End
      Begin VB.CheckBox ChkAllSeatType 
         BackColor       =   &H00FFFFFF&
         Caption         =   "所有座位类型(&S)"
         Height          =   255
         Left            =   990
         TabIndex        =   10
         Top             =   4215
         Width           =   1740
      End
      Begin RTComctl3.CoolButton flbClose 
         Height          =   240
         Left            =   2505
         TabIndex        =   16
         Top             =   15
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   423
         BTYPE           =   8
         TX              =   "r"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Marlett"
            Size            =   9
            Charset         =   2
            Weight          =   500
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
         MICON           =   "frmReportStation.frx":0520
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin RTComctl3.FlatLabel FlatLabel1 
         Height          =   285
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   2790
         _ExtentX        =   4921
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
         VerticalAlignment=   1
         HorizontalAlignment=   1
         MarginLeft      =   0
         MarginTop       =   0
         Caption         =   "查询条件设定"
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "上车站:"
         Height          =   180
         Left            =   120
         TabIndex        =   24
         Top             =   5430
         Width           =   630
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "车型:"
         Height          =   180
         Left            =   90
         TabIndex        =   5
         Top             =   2895
         Width           =   450
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "票价表:"
         Height          =   195
         Left            =   90
         TabIndex        =   0
         Top             =   420
         Width           =   1365
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "计划:"
         Height          =   210
         Left            =   15
         TabIndex        =   20
         Top             =   0
         Width           =   1650
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "站点:"
         Height          =   180
         Left            =   90
         TabIndex        =   2
         Top             =   1035
         Width           =   450
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "座位类型:"
         Height          =   180
         Left            =   90
         TabIndex        =   8
         Top             =   4230
         Width           =   810
      End
   End
   Begin VB.PictureBox ptResult 
      BackColor       =   &H80000009&
      Height          =   6030
      Left            =   3150
      ScaleHeight     =   5970
      ScaleWidth      =   5145
      TabIndex        =   13
      Top             =   0
      Width           =   5205
      Begin RTReportLF.RTReport RTReport 
         Height          =   3675
         Left            =   105
         TabIndex        =   21
         Top             =   1320
         Width           =   3945
         _ExtentX        =   6959
         _ExtentY        =   6482
      End
      Begin VB.PictureBox ptQ 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1155
         Left            =   0
         Picture         =   "frmReportStation.frx":053C
         ScaleHeight     =   1155
         ScaleWidth      =   5640
         TabIndex        =   18
         Top             =   0
         Width           =   5640
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "站点车型票价报表情况"
            BeginProperty Font 
               Name            =   "黑体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   810
            TabIndex        =   19
            Top             =   750
            Width           =   2400
         End
         Begin VB.Image imgOpen 
            Height          =   240
            Left            =   0
            Picture         =   "frmReportStation.frx":1432
            Top             =   0
            Visible         =   0   'False
            Width           =   240
         End
      End
   End
   Begin RTComctl3.Spliter spQuery 
      Height          =   1320
      Left            =   3060
      TabIndex        =   15
      Top             =   1050
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   2328
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
Attribute VB_Name = "frmReportStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************
'* Source File Name:frmReportRoute.frm
'* Project Name:PSTBusMan.vbp
'* Engineer:陈峰
'* Date Generated:2002/09/03
'* Last Revision Date:2002/09/03
'* Brief Description:线路票价报表
'* Relational Document:
'**********************************************************


Option Explicit
Const cszTemplateFile = "站点车型票价报表模板.xls"
Const cnTop = 1200

Private m_lMoveLeft As Long
Private m_aszAllSeatType() As String

Private m_aszVehicleType() As String
Private m_aszStation() As String
    
Private m_atAllSellStation() As TDepartmentInfo


Private m_rsPriceItem As Recordset
Private m_rsTicketType As Recordset

Private m_lRange As Long '写进度条用
    

Private Sub chkSellStation_Click()
    '选择(或取消)所有的座位类型
    Dim i As Integer
    If chkSellStation.Value = vbChecked Then
        For i = 0 To lstSellStation.ListCount - 1
            lstSellStation.Selected(i) = False
        Next i
        lstSellStation.Enabled = False
        chkSellStation.Value = vbChecked
    Else
        lstSellStation.Enabled = True
    End If
    EnabledQuery
End Sub


Private Sub cmdCancel_Click()
    '关闭
    Unload Me
End Sub


Private Sub cmdQuery_Click()
    m_lRange = 0
    QueryPrice
    ShowSBInfo ""
    ShowSBInfo "共有" & m_lRange & "条记录", ESB_ResultCountInfo
End Sub

Private Sub Form_Activate()
    Form_Resize
    MDIScheme.SetPrintEnabled True
End Sub

Private Sub Form_Deactivate()
    MDIScheme.SetPrintEnabled False
End Sub

Private Sub Form_Load()
    '初始化
    Dim oParam As New SystemParam
    Dim oPriceMan As New TicketPriceMan
    On Error GoTo ErrorHandle
    
    spQuery.InitSpliter ptQuery, ptResult
    m_lMoveLeft = 0
    
    oPriceMan.Init g_oActiveUser
    oParam.Init g_oActiveUser
    Set m_rsPriceItem = oPriceMan.GetAllTicketItemRS(TP_PriceItemUse)
    Set m_rsTicketType = oParam.GetAllTicketTypeRS(TP_TicketTypeValid)
    Set oPriceMan = Nothing
    Set oParam = Nothing
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub QueryPrice()
    '查询票价信息
    Dim oPriceReport As New PriceSheetReport
'    Dim aszSelectedBus() As String
    Dim aszSelectStation() As String
    Dim aszSelectVehicleType() As String
    Dim aszSelectSeatType() As String
    Dim aszSellStation() As String
    
    
    Dim rsResult As Recordset
        
    Dim arsTemp As Variant
    Dim aszTemp As Variant
    
'    aszSelectedBus = GetSelectBus
    aszSelectSeatType = GetSelectSeatType
    aszSelectStation = GetSelectStation
    aszSelectVehicleType = GetSelectVehicleType
    aszSellStation = GetSelectSellStation
    oPriceReport.Init g_oActiveUser
    Set rsResult = oPriceReport.GetStationVehiclePriceRpt(ResolveDisplay(cboPriceTable.Text), aszSelectStation, aszSelectVehicleType, aszSelectSeatType, aszSellStation)
    
    ReDim aszTemp(1 To 2)
    ReDim arsTemp(1 To 2)
    '赋票种
    aszTemp(1) = "票种"
    Set arsTemp(1) = m_rsTicketType
    aszTemp(2) = "票价项"
    Set arsTemp(2) = m_rsPriceItem
    
    '填充票价记录集
    RTReport.LeftLabelVisual = True
    RTReport.TopLabelVisual = True
    RTReport.CustomStringCount = aszTemp
    RTReport.CustomString = arsTemp
    RTReport.TemplateFile = App.Path & "\" & cszTemplateFile
    RTReport.ShowReport rsResult
    '设置固定行,列可见性
    Set oPriceReport = Nothing
    WriteProcessBar False
    Exit Sub
Here:
    Set oPriceReport = Nothing
    WriteProcessBar False
    ShowErrorMsg
End Sub

Private Function GetSelectStation() As String()
    '得到选择的站点
    Dim aszStation() As String
    Dim i As Integer
    Dim nCount As Integer
    If lstStation.SelCount > 0 Then
        ReDim aszStation(1 To lstStation.SelCount)
        nCount = 0
        For i = 0 To lstStation.ListCount - 1
            If lstStation.Selected(i) Then
                nCount = nCount + 1
                aszStation(nCount) = m_aszStation(i + 1, 1)
            End If
        Next i
    ElseIf chkAllStation.Value = vbChecked Then
        ReDim aszStation(1 To lstStation.ListCount)
        For i = 0 To lstStation.ListCount - 1
            aszStation(i + 1) = m_aszStation(i + 1, 1)
            
        Next i
    End If
    GetSelectStation = aszStation
End Function
Private Function GetSelectVehicleType() As String()
    '得到选择的车型
    Dim aszVehicleType() As String
    Dim i As Integer
    Dim nCount As Integer
    If lstType.SelCount > 0 Then
        ReDim aszVehicleType(1 To lstType.SelCount)
        nCount = 0
        For i = 0 To lstType.ListCount - 1
            If lstType.Selected(i) Then
                nCount = nCount + 1
                aszVehicleType(nCount) = m_aszVehicleType(i + 1, 1)
            End If
        Next i
    ElseIf chkAllType.Value = vbChecked Then
        ReDim aszVehicleType(1 To lstType.ListCount)
        For i = 0 To lstType.ListCount - 1
            aszVehicleType(i + 1) = m_aszVehicleType(i + 1, 1)
            
        Next i
    End If
    GetSelectVehicleType = aszVehicleType
End Function

Private Function GetSelectSeatType() As String()
    '得到所有选择的车次
    Dim aszSeatType() As String
    Dim i As Integer
    Dim nCount As Integer
    If lstSeatType.SelCount > 0 Then
        ReDim aszSeatType(1 To lstSeatType.SelCount)
        nCount = 0
        For i = 0 To lstSeatType.ListCount - 1
            If lstSeatType.Selected(i) Then
                nCount = nCount + 1
                aszSeatType(nCount) = m_aszAllSeatType(i + 1, 1)
            End If
        Next i
    ElseIf ChkAllSeatType.Value = vbChecked Then
        ReDim aszSeatType(1 To lstSeatType.ListCount)
        For i = 0 To lstSeatType.ListCount - 1
            aszSeatType(i + 1) = m_aszAllSeatType(i + 1, 1)
            
        Next i
    End If
    GetSelectSeatType = aszSeatType
End Function
Private Function GetSelectSellStation() As String()
    '得到所有选择的上车站
    Dim aszSellStation() As String
    Dim i As Integer
    Dim nCount As Integer
    If lstSellStation.SelCount > 0 Then
        ReDim aszSellStation(1 To lstSellStation.SelCount)
        nCount = 0
        For i = 0 To lstSellStation.ListCount - 1
            If lstSellStation.Selected(i) Then
                nCount = nCount + 1
                aszSellStation(nCount) = m_atAllSellStation(i + 1).szSellStationID
            End If
        Next i
    ElseIf chkSellStation.Value = vbChecked Then
        ReDim aszSellStation(1 To lstSellStation.ListCount)
        For i = 0 To lstSellStation.ListCount - 1
            aszSellStation(i + 1) = m_atAllSellStation(i + 1).szSellStationID
            
        Next i
    End If
    GetSelectSellStation = aszSellStation
    
    
End Function

Private Sub FillPriceTable()
    '填充所有的票价表
    Dim i As Integer
    Dim oTicketPriceMan As New TicketPriceMan
    Dim nCount As Integer
    Dim aszAllRoutePriceTable() As String
        
    On Error GoTo ErrorHandle
    cboPriceTable.Clear
    oTicketPriceMan.Init g_oActiveUser
    aszAllRoutePriceTable = oTicketPriceMan.GetAllRoutePriceTable()
    nCount = ArrayLength(aszAllRoutePriceTable)
    For i = 1 To nCount
        cboPriceTable.AddItem MakeDisplayString(aszAllRoutePriceTable(i, 1), aszAllRoutePriceTable(i, 2))
    Next i
    If cboPriceTable.ListCount > 0 Then cboPriceTable.ListIndex = 0
    Set oTicketPriceMan = Nothing
    Exit Sub
ErrorHandle:
    Set oTicketPriceMan = Nothing
    ShowErrorMsg
End Sub

Private Sub FillStation()
    '填充所有的站点
    Dim i As Integer
    Dim oBaseInfo As New BaseInfo
    Dim nCount As Integer
    On Error GoTo ErrorHandle
    lstStation.Clear
    oBaseInfo.Init g_oActiveUser
    m_aszStation = oBaseInfo.GetStation
    nCount = ArrayLength(m_aszStation)
    For i = 1 To nCount
        lstStation.AddItem m_aszStation(i, 2)
    Next i
    Set oBaseInfo = Nothing
    Exit Sub
ErrorHandle:
    Set oBaseInfo = Nothing
    ShowErrorMsg
End Sub

Private Sub FillVehicleType()
    '填充所有的车型
    Dim i As Integer
    Dim oBaseInfo As New BaseInfo
    Dim nCount As Integer
    On Error GoTo ErrorHandle
    lstType.Clear
    oBaseInfo.Init g_oActiveUser
    m_aszVehicleType = oBaseInfo.GetAllVehicleModel
    nCount = ArrayLength(m_aszVehicleType)
    For i = 1 To nCount
        lstType.AddItem m_aszVehicleType(i, 2)
    Next i
    Set oBaseInfo = Nothing
    Exit Sub
ErrorHandle:
    Set oBaseInfo = Nothing
    ShowErrorMsg
End Sub

Private Sub FillSeatType()
    '往组合框中添加座位类型
    Dim nCount As Integer
    Dim i As Integer
    Dim oBaseInfo As New BaseInfo
    
    oBaseInfo.Init g_oActiveUser
    m_aszAllSeatType = oBaseInfo.GetAllSeatType
    nCount = ArrayLength(m_aszAllSeatType)
    For i = 1 To nCount
        lstSeatType.AddItem m_aszAllSeatType(i, 2)
    Next i
'    ChkAllSeatType.Value = vbChecked

End Sub
Private Sub FillSellStation()
'    '往组合框中添加座位类型
    Dim nCount As Integer
    Dim i As Integer
    Dim oSystemMan As New SystemMan
'    Dim atAllSellStation() As TDepartmentInfo
    

    oSystemMan.Init g_oActiveUser
    m_atAllSellStation = oSystemMan.GetAllSellStation
    nCount = ArrayLength(m_atAllSellStation)
    
    For i = 1 To nCount
        lstSellStation.AddItem m_atAllSellStation(i).szSellStationName
    Next i
End Sub
Private Sub Form_Unload(Cancel As Integer)
    MDIScheme.SetPrintEnabled False
End Sub

Private Sub lstSeatType_Click()
    EnabledQuery
End Sub



Private Sub lstSellStation_Click()
    EnabledQuery
End Sub

Private Sub lstStation_Click()
    EnabledQuery
End Sub

Private Sub lstType_Click()
EnabledQuery
End Sub

Private Sub RTReport_SetProgressRange(ByVal lRange As Variant)
    m_lRange = lRange
    
End Sub

Private Sub RTReport_SetProgressValue(ByVal lValue As Variant)
    WriteProcessBar True, lValue, m_lRange, "正在填充票价..."
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    FillPriceTable
    FillStation
    FillVehicleType
    FillSeatType
    FillSellStation
    
    EnabledQuery
End Sub

Private Sub Form_Resize()
    '重绘
    spQuery.LayoutIt
End Sub

Private Sub ptQuery_Resize()
    flbClose.Move ptQuery.ScaleWidth - 250
End Sub

Private Sub ptResult_Resize()
    Dim lTemp As Long
    lTemp = IIf((ptResult.ScaleHeight - cnTop) <= 0, lTemp, ptResult.ScaleHeight - cnTop)
    RTReport.Move 0, cnTop, ptResult.ScaleWidth, lTemp
    'lblTitle.Move 60 + m_lMoveLeft, 50, ptResult.ScaleWidth

    FlatLabel1.Width = ptQuery.ScaleWidth
    flbClose.Left = FlatLabel1.Left + FlatLabel1.Width - flbClose.Width - 30
End Sub

Private Sub flbClose_Click()
    '关闭左边部分
    ptQuery.Visible = False
    imgOpen.Visible = True
    m_lMoveLeft = 240
    'lblTitle.Move 60 + m_lMoveLeft
    spQuery.LayoutIt
End Sub

Private Sub imgOpen_Click()
    '打开左边部分
    ptQuery.Visible = True
    imgOpen.Visible = False
    m_lMoveLeft = 0
    'lblTitle.Move 60 + m_lMoveLeft
    spQuery.LayoutIt
End Sub

Private Sub ChkAllSeatType_Click()
    '选择(或取消)所有的座位类型
    Dim i As Integer
    If ChkAllSeatType.Value = vbChecked Then
        For i = 0 To lstSeatType.ListCount - 1
            lstSeatType.Selected(i) = False
        Next i
        lstSeatType.Enabled = False
        ChkAllSeatType.Value = vbChecked
    Else
        lstSeatType.Enabled = True
    End If
    EnabledQuery
End Sub

Private Sub chkAllStation_Click()
    '选择(或取消)所有的站点
    Dim i As Integer
    If chkAllStation.Value = vbChecked Then
        For i = 0 To lstStation.ListCount - 1
            lstStation.Selected(i) = False
        Next i
        lstStation.Enabled = False
        chkAllStation.Value = vbChecked
    Else
        lstStation.Enabled = True
    End If
    EnabledQuery
End Sub

Private Sub chkAllType_Click()
    '选择(或取消)所有的车型
    Dim i As Integer
    If chkAllType.Value = vbChecked Then
        For i = 0 To lstType.ListCount - 1
            lstType.Selected(i) = False
        Next i
        lstType.Enabled = False
        chkAllType.Value = vbChecked
    Else
        lstType.Enabled = True
    End If
    EnabledQuery
End Sub

Private Sub EnabledQuery()
    '查询按钮是否可用
    If cboPriceTable.Text <> "" And (lstSeatType.SelCount > 0 Or ChkAllSeatType.Value = vbChecked) _
    And (lstStation.SelCount > 0 Or chkAllStation.Value = vbChecked) And (lstType.SelCount > 0 Or chkAllType.Value = vbChecked) Then
        cmdQuery.Enabled = True
    Else
        cmdQuery.Enabled = False
    End If
End Sub

Public Sub PrintReport(Optional pbShowDialog As Boolean = False)
    On Error GoTo ErrorHandle
    RTReport.PrintReport pbShowDialog
    Exit Sub
ErrorHandle:
End Sub

Public Sub PreView()
    RTReport.PrintView
End Sub

Public Sub PageSet()
    RTReport.OpenDialog EDialogType.PAGESET_TYPE
End Sub

Public Sub PrintSet()
    RTReport.OpenDialog EDialogType.PRINT_TYPE
End Sub
'导出文件
Public Sub ExportFile()
    Dim szFileName As String
    szFileName = RTReport.OpenDialog(EDialogType.EXPORT_FILE)
End Sub
'导出文件并打开
Public Sub ExportFileOpen()
    Dim szFileName As String
    szFileName = RTReport.OpenDialog(EDialogType.EXPORT_FILE)
    If szFileName <> "" Then
        OpenLinkedFile szFileName
    End If
End Sub

