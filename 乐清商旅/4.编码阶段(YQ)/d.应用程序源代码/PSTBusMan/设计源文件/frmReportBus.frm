VERSION 5.00
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Object = "{D11DF87C-EFEC-4838-B7E2-15462BF136FB}#1.4#0"; "RTReportLF.ocx"
Begin VB.Form frmReportBus 
   BackColor       =   &H80000005&
   Caption         =   "车次总票价表查询"
   ClientHeight    =   6585
   ClientLeft      =   3075
   ClientTop       =   3285
   ClientWidth     =   9375
   HelpContextID   =   1003001
   Icon            =   "frmReportBus.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6585
   ScaleWidth      =   9375
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   4080
      Top             =   3045
   End
   Begin VB.PictureBox ptResult 
      BackColor       =   &H80000009&
      Height          =   6030
      Left            =   3375
      ScaleHeight     =   5970
      ScaleWidth      =   5355
      TabIndex        =   12
      Top             =   15
      Width           =   5415
      Begin VB.PictureBox ptQ 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1155
         Left            =   0
         Picture         =   "frmReportBus.frx":014A
         ScaleHeight     =   1155
         ScaleWidth      =   5640
         TabIndex        =   17
         Top             =   0
         Width           =   5640
         Begin VB.Image imgOpen 
            Height          =   240
            Left            =   0
            Picture         =   "frmReportBus.frx":1040
            Top             =   0
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "车次总票价报表情况"
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
            Left            =   735
            TabIndex        =   18
            Top             =   780
            Width           =   2160
         End
      End
      Begin RTReportLF.RTReport RTReport 
         Height          =   3225
         Left            =   75
         TabIndex        =   16
         Top             =   1275
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   5689
      End
   End
   Begin VB.PictureBox ptQuery 
      BackColor       =   &H80000009&
      Height          =   7965
      Left            =   315
      ScaleHeight     =   7905
      ScaleWidth      =   2790
      TabIndex        =   11
      Top             =   30
      Width           =   2850
      Begin VB.ListBox lstSellStation 
         Height          =   780
         Left            =   120
         MultiSelect     =   2  'Extended
         TabIndex        =   21
         Top             =   5070
         Width           =   2535
      End
      Begin VB.CheckBox chkSellStation 
         BackColor       =   &H00FFFFFF&
         Caption         =   "所有的上车站(&T)"
         Height          =   180
         Left            =   960
         TabIndex        =   20
         Top             =   4740
         Width           =   1710
      End
      Begin VB.CheckBox ChkRunStation 
         BackColor       =   &H80000009&
         Caption         =   "停靠站点(&P)"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   5970
         Width           =   1350
      End
      Begin VB.ListBox lstSeatType 
         Height          =   780
         Left            =   120
         MultiSelect     =   2  'Extended
         TabIndex        =   6
         Top             =   3915
         Width           =   2535
      End
      Begin VB.CheckBox ChkAllSeatType 
         BackColor       =   &H00FFFFFF&
         Caption         =   "所有座位类型(&S)"
         Height          =   255
         Left            =   960
         TabIndex        =   7
         Top             =   3570
         Width           =   1740
      End
      Begin VB.ComboBox cboPriceTable 
         Height          =   300
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   675
         Width           =   2565
      End
      Begin VB.CheckBox chkAllBus 
         BackColor       =   &H80000009&
         Caption         =   "所有车次(&B)"
         Height          =   240
         Left            =   960
         TabIndex        =   4
         Top             =   1035
         Width           =   1305
      End
      Begin RTComctl3.CoolButton cmdQuery 
         Default         =   -1  'True
         Height          =   330
         Left            =   135
         TabIndex        =   9
         Top             =   6375
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
         MICON           =   "frmReportBus.frx":118A
         PICN            =   "frmReportBus.frx":11A6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.ListBox lstBus 
         Height          =   2220
         Left            =   120
         MultiSelect     =   2  'Extended
         TabIndex        =   3
         Top             =   1290
         Width           =   2610
      End
      Begin RTComctl3.CoolButton cmdCancel 
         Height          =   330
         Left            =   1425
         TabIndex        =   10
         Top             =   6375
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
         MICON           =   "frmReportBus.frx":1540
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin RTComctl3.CoolButton flbClose 
         Height          =   240
         Left            =   2490
         TabIndex        =   14
         Top             =   15
         Width           =   285
         _ExtentX        =   503
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
         MICON           =   "frmReportBus.frx":155C
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
         TabIndex        =   15
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "上车站:"
         Height          =   180
         Left            =   120
         TabIndex        =   19
         Top             =   4740
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "座位类型:"
         Height          =   180
         Left            =   120
         TabIndex        =   5
         Top             =   3615
         Width           =   810
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "票价表:"
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   435
         Width           =   1365
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "车次:"
         Height          =   300
         Left            =   120
         TabIndex        =   2
         Top             =   1050
         Width           =   2340
      End
   End
   Begin RTComctl3.Spliter spQuery 
      Height          =   1320
      Left            =   3180
      TabIndex        =   13
      Top             =   1605
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
Attribute VB_Name = "frmReportBus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************
'* Source File Name:frmReportBus.frm
'* Project Name:PSTBusMan.vbp
'* Engineer:陈峰
'* Date Generated:2002/09/11
'* Last Revision Date:2002/09/11
'* Brief Description:车次票价报表
'* Relational Document:
'**********************************************************

Option Explicit
Const cszTemplateFile = "车次总票价报表模板.xls"
Const cnTop = 1200

Private m_lMoveLeft As Long
Private m_aszAllBus() As String
Private m_aszAllSeatType() As String

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
    Dim aszSelectedBus() As String
    Dim aszSelectSeatType() As String
    Dim aszSellStation() As String
    
    Dim rsResult As Recordset
        
    Dim arsTemp As Variant
    Dim aszTemp As Variant
    
    aszSelectedBus = GetSelectBus
    aszSelectSeatType = GetSelectSeatType
    aszSellStation = GetSelectSellStation
    
    oPriceReport.Init g_oActiveUser
    Set rsResult = oPriceReport.GetBusRoutePriceItem(ResolveDisplay(cboPriceTable.Text), g_szExePriceTable, aszSelectedBus, aszSelectSeatType, aszSellStation, IIf(ChkRunStation.Value = vbChecked, True, False))
    
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
    ShowSBInfo
    ShowSBInfo "共有" & m_lRange & "条记录", ESB_ResultCountInfo
    
    Exit Sub
Here:
    Set oPriceReport = Nothing
    WriteProcessBar False
    ShowSBInfo
    ShowErrorMsg
End Sub

Private Function GetSelectBus() As String()
    '得到所有选择的车次
    Dim aszBusID() As String
    Dim i As Integer
    Dim nCount As Integer
    If lstBus.SelCount > 0 Then
        '如果有选择
        ReDim aszBusID(1 To lstBus.SelCount)
        nCount = 0
        For i = 0 To lstBus.ListCount - 1
            If lstBus.Selected(i) Then
                nCount = nCount + 1
                aszBusID(nCount) = m_aszAllBus(i + 1, 1)
            End If
        Next i
    ElseIf chkAllBus.Value = vbChecked Then
        '如果车次全选
        ReDim aszBusID(1 To lstBus.ListCount)
'        nCount = 0
        For i = 0 To lstBus.ListCount - 1
'            nCount = nCount + 1
            aszBusID(i + 1) = m_aszAllBus(i + 1, 1)
        Next i
    End If
    GetSelectBus = aszBusID
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
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub FillBus()
    '填充所有的车次
    Dim i As Integer
    Dim nCount As Integer
    Dim oBusProject As New BusProject
    oBusProject.Init g_oActiveUser
    oBusProject.Identify
    m_aszAllBus = oBusProject.GetAllBus
    nCount = ArrayLength(m_aszAllBus)
    For i = 1 To nCount
        lstBus.AddItem m_aszAllBus(i, 1) & " " & Format(m_aszAllBus(i, 2), "HH:MM") & " " & m_aszAllBus(i, 4)
    Next i
    
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

Private Sub lstBus_Click()
    EnabledQuery
End Sub

Private Sub lstBus_ItemCheck(Item As Integer)
    EnabledQuery
End Sub

Private Sub lstSeatType_Click()
    EnabledQuery
End Sub

Private Sub lstSeatType_ItemCheck(Item As Integer)
    EnabledQuery
End Sub

Private Sub lstSellStation_Click()
EnabledQuery
End Sub

Private Sub lstSellStation_ItemCheck(Item As Integer)
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
    FillBus
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

Private Sub chkAllBus_Click()
    '选择(或取消选择)所有的车次
    Dim i As Integer
    If chkAllBus.Value = vbChecked Then
        For i = 0 To lstBus.ListCount - 1
            lstBus.Selected(i) = False
        Next i
        lstBus.Enabled = False
    Else
        lstBus.Enabled = True
    End If
    EnabledQuery
End Sub

Private Sub EnabledQuery()
    '查询按钮是否可用
    If cboPriceTable.Text <> "" And (lstBus.SelCount > 0 Or chkAllBus.Value = vbChecked) And (lstSeatType.SelCount > 0 Or ChkAllSeatType.Value = vbChecked) And (lstSellStation.SelCount > 0 Or chkSellStation.Value = vbChecked) Then
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
