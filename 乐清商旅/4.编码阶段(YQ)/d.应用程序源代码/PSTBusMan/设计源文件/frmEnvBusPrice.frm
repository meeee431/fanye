VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Object = "{D11DF87C-EFEC-4838-B7E2-15462BF136FB}#2.1#0"; "RTReportLF.ocx"
Begin VB.Form frmEnvBusPrice 
   BackColor       =   &H00E0E0E0&
   Caption         =   "环境车次票价"
   ClientHeight    =   4965
   ClientLeft      =   1230
   ClientTop       =   2880
   ClientWidth     =   9390
   HelpContextID   =   10000790
   Icon            =   "frmEnvBusPrice.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   9390
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraButton 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H00E0E0E0&
      Height          =   585
      Left            =   4110
      TabIndex        =   11
      Top             =   4290
      Width           =   5085
      Begin RTComctl3.CoolButton CoolButton1 
         Height          =   315
         Left            =   3930
         TabIndex        =   12
         Top             =   150
         Width           =   1125
         _ExtentX        =   1984
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
         MICON           =   "frmEnvBusPrice.frx":014A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin RTComctl3.CoolButton cmdRefresh 
         Height          =   315
         Left            =   150
         TabIndex        =   5
         Top             =   150
         Width           =   1125
         _ExtentX        =   0
         _ExtentY        =   0
         BTYPE           =   3
         TX              =   "刷新(&R)"
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
         MICON           =   "frmEnvBusPrice.frx":0166
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Timer tmStart 
         Interval        =   500
         Left            =   0
         Top             =   0
      End
      Begin RTComctl3.CoolButton cmdOk 
         Height          =   315
         Left            =   1440
         TabIndex        =   6
         Top             =   150
         Width           =   1125
         _ExtentX        =   0
         _ExtentY        =   0
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
         MICON           =   "frmEnvBusPrice.frx":0182
         PICN            =   "frmEnvBusPrice.frx":019E
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
         Height          =   315
         Left            =   2700
         TabIndex        =   7
         Top             =   150
         Width           =   1125
         _ExtentX        =   0
         _ExtentY        =   0
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
         MICON           =   "frmEnvBusPrice.frx":0538
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
   Begin VB.Frame fraInfo 
      BackColor       =   &H00E0E0E0&
      Height          =   705
      Left            =   60
      TabIndex        =   8
      Top             =   0
      Width           =   9165
      Begin FText.asFlatTextBox txtBusID 
         Height          =   300
         Left            =   1230
         TabIndex        =   1
         Top             =   240
         Width           =   990
         _ExtentX        =   1746
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
      Begin MSComCtl2.DTPicker dtpOffDate 
         Height          =   300
         Left            =   3450
         TabIndex        =   3
         Top             =   240
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   529
         _Version        =   393216
         Format          =   69337089
         CurrentDate     =   36396
      End
      Begin VB.Label lblOffDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车次日期(&D):"
         Height          =   180
         Left            =   2340
         TabIndex        =   2
         Top             =   300
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车次代码(&B):"
         Height          =   180
         Left            =   120
         TabIndex        =   0
         Top             =   300
         Width           =   1080
      End
      Begin VB.Label lblOffTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "发车时间:12:12"
         Height          =   180
         Left            =   4815
         TabIndex        =   10
         Top             =   300
         Width           =   1260
      End
      Begin VB.Label lblRoute 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "运行线路:"
         Height          =   180
         Left            =   6255
         TabIndex        =   9
         Top             =   300
         Width           =   810
      End
   End
   Begin RTReportLF.RTReport RTReport 
      Height          =   3525
      Left            =   90
      TabIndex        =   4
      Top             =   720
      Width           =   9105
      _ExtentX        =   16060
      _ExtentY        =   6218
   End
End
Attribute VB_Name = "frmEnvBusPrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const cszTemplateFile = "envbusprice.xls"
Const cnPriceItemStartCol = 6        '可编辑票价项的起始列
Const cnPriceItemStartRow = 2       '可编辑票价项的起始行
Const cnTotalCol = 5 '总计所在列

Public m_szBusID As String '存放选择的车次
Public m_dtEnvDate As Date '选择的车次日期
Public m_bDisplayOnly As Boolean '是否只显示


Private m_oRoutePriceTable As RoutePriceTable
Private m_oREScheme As New STReSch.REScheme
Private WithEvents F1Book As TTF160Ctl.F1Book
Attribute F1Book.VB_VarHelpID = -1
Private m_rsResultPrice As Recordset      '票价表结果记录集
Private m_rsAllTicketItem As Recordset    '所有的票种类型
'Private m_tHalfItemParam() As THalfTicketItemParam

Private m_atHalfItemParam() As THalfTicketItemParam '半票及优惠票票价项计算参数
Private m_oMantissa As New clMantissa '尾数处理对象

Private m_bChanged As Boolean '标志是否改变

Private m_abChanged() As Boolean '存放每一行是否修改的标志

Private Sub CoolButton1_Click()
DisplayHelp Me
End Sub

Private Sub F1Book_DblClick(ByVal nRow As Long, ByVal nCol As Long)
    F1Book.StartEdit False, True, False
End Sub

Private Sub cmdCancel_Click()
    m_szBusID = ""
    m_bChanged = False
    Unload Me
End Sub


Private Sub F1Book_EndEdit(EditString As String, Cancel As Integer)
    Dim szTicketTypeID As String
    Dim lRow As Long
    With F1Book
        If Not IsNumeric(EditString) Then
            '如果不是输入数字则出错
            Cancel = True
            MsgBox "请输入数字", vbInformation, Me.Caption
        Else
            '如果修改了值,则设置已修改的颜色
            If .Text <> EditString Then
                SetSaveEnabled True  '设置保存可用
                If .Col >= cnPriceItemStartCol Then
                    '如果是各票价项
                    .Text = EditString '此处赋值是为了适合SetTailCarry ,里面用的是.text
                    '进行尾数处理
                    m_oMantissa.SetTailCarry .Row, .Row, .Col, False
                    '此处由于用的是判断当前行,所以才会出现循环触发
                    lRow = .Row
                    szTicketTypeID = GetTicketTypeID(.Row)
                    If szTicketTypeID = TP_FullPrice Then
                        '如果此行为全票行,则修改相应的半票、优惠票等
                        ModifyHalfPrice .Row, .Col
                    End If
                    .Row = lRow
                    EditString = .Text '回赋,为了修改后的此过程退出时会自动回赋.Text=EditString
                    
                    SetChangeColor .Row, .Col
                End If
            End If
        End If
    End With
End Sub



Private Sub F1Book_SelChange()
    Dim lRow As Long, lCol As Long
    lRow = F1Book.Row
    lCol = F1Book.Col
    If lRow < cnPriceItemStartRow Or lCol < cnPriceItemStartCol Then
'        F1Book.ShowEditBar = False
        F1Book.AllowInCellEditing = False
    Else
'        F1Book.ShowEditBar = True
        F1Book.AllowInCellEditing = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrorHandle
    If Not QueryCancel Then
        Cancel = 0
    Else
        Cancel = 1
    End If
    If Cancel <> 1 Then
        Set m_oMantissa = Nothing
    End If
    Exit Sub
ErrorHandle:
    Cancel = 1
    ShowErrorMsg
End Sub


Private Sub Form_Load()
    Dim oHalfTicket As New HalfTicketPrice
    Dim oPriceMan As New stprice.TicketPriceMan
    On Error GoTo ErrorHandle
'    AlignFormPos Me
    lblOffDate.Visible = True
    dtpOffDate.Visible = True
    dtpOffDate.Value = m_dtEnvDate
    If m_bDisplayOnly Then
        txtBusId.Enabled = False
        dtpOffDate.Enabled = False
    Else
        txtBusId.Enabled = True
        dtpOffDate.Enabled = True
    End If
    
    txtBusId.Enabled = True
    Set m_oRoutePriceTable = CreateObject("stprice.RoutePriceTable")
    m_oRoutePriceTable.Init g_oActiveUser
    m_oRoutePriceTable.Identify g_szExePriceTable
    
    txtBusId.Text = m_szBusID
    SetSaveEnabled False
    Set F1Book = RTReport.CellObject
'    RTReport.Enabled = False
    '读取初始数据
    '得到所有票价项
    oPriceMan.Init g_oActiveUser
    Set m_rsAllTicketItem = oPriceMan.GetAllTicketItemRS(TP_PriceItemUse)
    '得到所有票种的半票计算参数
    oHalfTicket.Init g_oActiveUser
    m_atHalfItemParam = oHalfTicket.GetItemParam(0, g_szExePriceTable, TP_PriceItemUse)
    
    RefreshBus
    OpenBusPrice
    InitMantissa
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub OpenBusPrice()
    '打开车次票价
    Dim atDetailInfo() As TBusPriceDetailInfo
    Dim aszBusID(1 To 1, 1 To 1) As String
'    Dim aszSeatType() As String
'    Dim aszVehicleModel() As String
    Dim aszTemp(1 To 1) As String
    Dim arsTemp(1 To 1) As Recordset
'    Dim aszTempBusID()
    On Error GoTo ErrorHandle
    SetSaveEnabled False
    SetBusy
    aszBusID(1, 1) = m_szBusID
    Set m_rsResultPrice = m_oREScheme.GetBusTicketInfoRS(m_dtEnvDate, aszBusID)
    aszTemp(1) = "票价项"
    m_rsAllTicketItem.MoveFirst
    Set arsTemp(1) = m_rsAllTicketItem
    '填充票价记录集
    RTReport.LeftLabelVisual = True
    RTReport.TopLabelVisual = True
    RTReport.CustomStringCount = aszTemp
    RTReport.CustomString = arsTemp
    RTReport.TemplateFile = App.Path & "\" & cszTemplateFile
    RTReport.ShowReport m_rsResultPrice
    '设置固定行,列可见性
    Set F1Book = RTReport.CellObject
    F1Book.AllowInCellEditing = True
    F1Book.AllowDelete = False
    F1Book.Col = cnPriceItemStartCol
    If F1Book.LastRow >= 2 Then F1Book.Row = 2
    If F1Book.LastRow > 0 Then ReDim m_abChanged(1 To F1Book.LastRow)
    SetNormal
    Exit Sub
ErrorHandle:
    ShowErrorMsg
    SetNormal
End Sub
Public Sub SaveBusPrice()
    '保存车次票价
    Dim tDetailInfo(1 To 1) As TBusPriceDetailInfo
    Dim i As Long, k As Long
    Dim bModify As Boolean '标志此行是否被修改
    Dim szPriceItem As String
    Dim oRebus As New REBus
    Dim tPriceInfo As TRETicketPrice
    On Error GoTo ErrorHandle
    oRebus.Init g_oActiveUser
    
    With F1Book
        WriteProcessBar True, 1, .LastRow, "正在保存车次票价表"
        m_rsResultPrice.MoveFirst
        For i = cnPriceItemStartRow To .LastRow
            '得到修改状态
            bModify = GetModifyStatus(i)
            If bModify Then
                '如果为已修改或者为新增状态
                '车次
                oRebus.Identify m_rsResultPrice!bus_id, m_dtEnvDate
                
                '上车站
                tPriceInfo.szSellStationID = m_rsResultPrice!sell_station_id
                '站点
                tPriceInfo.szStationID = m_rsResultPrice!station_id
                '座位类型
                tPriceInfo.szSeatType = m_rsResultPrice!seat_type_id
                '票种
                tPriceInfo.nTicketType = m_rsResultPrice!ticket_type
                '总价
                tPriceInfo.sgTotal = .NumberRC(i, cnTotalCol)
                '基本运价
                tPriceInfo.sgBase = .NumberRC(i, cnPriceItemStartCol)
                '各票价项
                For k = cnPriceItemStartCol + 1 To .MaxCol
                    szPriceItem = GetPriceItem(k)
                    tPriceInfo.asgPrice(CInt(szPriceItem)) = .NumberRC(i, k)
                Next k
                oRebus.ModifyBusTicket tPriceInfo
                '设置修改状态
                MarkCellRowModifyStatus i
            End If
            WriteProcessBar , i, .LastRow, "正在保存车次票价表"
            m_rsResultPrice.MoveNext
        Next
        WriteProcessBar False
'        .DoRedrawAll
    End With
    WriteProcessBar False
    '设置保存不可用
    SetSaveEnabled False
    Exit Sub
ErrorHandle:
    ShowErrorMsg

End Sub


Private Sub cmdOk_Click()
    SaveBusPrice
    RTReport.SetFocus
End Sub
Public Sub PrintReport(Optional pbShowDialog As Boolean = False)
    '打印
    F1Book.FilePrint pbShowDialog
End Sub

Public Sub PreView()
    '打印预览
    RTReport.PrintView
End Sub

Public Sub PageSet()
    '页面设置
End Sub



Private Sub cmdRefresh_Click()
'    m_szBusID = txtBusID.Text
'    m_dtEnvDate = dtpOffDate.Value
    RefreshBus
    FillBusPriceRS
'    tmStart_Timer
End Sub

Private Sub InitMantissa()
''    初始化对象的属性
    m_oMantissa.MaxCol = F1Book.MaxCol
    m_oMantissa.oF1Book = RTReport.CellObject
    m_oMantissa.oPriceTable = m_oRoutePriceTable
    m_oMantissa.PriceItemStartCol = cnPriceItemStartCol
    m_oMantissa.PriceItemStartRow = cnPriceItemStartRow
    m_oMantissa.PriceRs = m_rsResultPrice
    m_oMantissa.TotalCol = cnPriceItemStartCol - 1
    m_oMantissa.UseItemRs = m_rsAllTicketItem
End Sub


Private Sub dtpOffDate_LostFocus()
    RefreshBusPrice
End Sub
'
'
''计算优惠票价
'Private Function CalHalfPrice(ByVal psgFullPrice As Single, ByVal pnTicketType As Integer, ByVal pnPriceItemIndex As Integer) As Single
'    Dim THalfTicketItemParam As THalfTicketItemParam
'    Dim i As Integer
'    '得到对应票种指定票价项的优惠计算参数
'    For i = 1 To ArrayLength(m_tHalfItemParam)
'        If Val(m_tHalfItemParam(i).szTicketType) = pnTicketType And Val(m_tHalfItemParam(i).szTicketItem) = pnPriceItemIndex Then
'            '找到了
'            THalfTicketItemParam = m_tHalfItemParam(i)
'            Exit For
'        End If
'    Next i
'    If i > ArrayLength(m_tHalfItemParam) Then Exit Function     '找不到
'
'
'    Dim sgResult As Single
'    sgResult = psgFullPrice * THalfTicketItemParam.sgParam1 + THalfTicketItemParam.sgParam2     '根据公式计算得到
'    Dim tTmp As TDealValue
'    tTmp = SetTailCarry(sgResult, Trim(m_rsResultPrice.Fields("station_id")), Trim(m_rsResultPrice.Fields("bus_type")), THalfTicketItemParam.szTicketItem)
'    sgResult = tTmp.sgValue
'
'    CalHalfPrice = sgResult
'End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub


'得到票价项名称
Private Function GetPriceItemName(pszIndex As String) As String
    If Val(pszIndex) = 0 Then
        GetPriceItemName = "base_carriage"
    Else
        GetPriceItemName = "price_item_" & Val(pszIndex)
    End If
End Function
'刷新车次信息
Private Sub RefreshBus()
On Error GoTo ErrHandle
'    If mbProjectBus Then
'        Dim oBus As New Bus
'        oBus.Init g_oActiveUser
'        oBus.Identify g_szExePlanID, m_szBusID
'
'        If oBus.BusType = TP_ScrollBus Then
'            lblOffTime.Caption = "间隔时间:" & oBus.ScrollBusCheckTime & " 分钟"
'        Else
'            lblOffTime.Caption = "发车时间:" & Format(oBus.StartupTime, "hh:mm")
'        End If
'        lblRoute.Caption = "运行线路:" & oBus.RouteName
'    Else
        Dim oRebus As New REBus
        oRebus.Init g_oActiveUser
        oRebus.Identify m_szBusID, m_dtEnvDate

        If oRebus.BusType = TP_ScrollBus Then
            lblOffTime.Caption = "间隔时间:" & oRebus.ScrollBusCheckTime & " 分钟"
        Else
            lblOffTime.Caption = "发车时间:" & Format(oRebus.StartUpTime, "hh:mm")
        End If
        lblRoute.Caption = "运行线路:" & oRebus.RouteName
'    End If
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Const cnMargin = 60
    fraInfo.Width = Me.ScaleWidth - 2 * cnMargin
    RTReport.Width = fraInfo.Width - cnMargin / 2
    'fraButton.Width = fraInfo.Width
    RTReport.Height = Me.ScaleHeight - fraInfo.Height - fraButton.Height
    fraButton.Top = RTReport.Top + RTReport.Height
    fraButton.Left = Me.ScaleWidth - fraButton.Width - 100
End Sub


'Private Sub Form_Unload(Cancel As Integer)
'    SaveFormPos Me
'    Set m_oReBus = Nothing
'End Sub



Private Sub RTReport_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        cmdCancel_Click
    End If
End Sub

Private Sub tmStart_Timer()
On Error GoTo ErrHandle
    SetBusy
    tmStart.Enabled = False
    
    '填充车次票价记录集
    FillBusPriceRS
    
    
'    RTReport.Enabled = True
    SetNormal
    Exit Sub
ErrHandle:
'    RTReport.Enabled = True
    SetNormal
    ShowErrorMsg
End Sub

Private Sub txtBusID_ButtonClick()
On Error GoTo ErrHandle
    Dim oShell As New CommDialog
    Dim aszTmp() As String
    oShell.Init g_oActiveUser
    aszTmp = oShell.SelectBus
    Set oShell = Nothing
    If ArrayLength(aszTmp) = 0 Then Exit Sub
    txtBusId.Text = aszTmp(1, 1)
    RefreshBusPrice
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

'得到车次票价记录集
Private Sub FillBusPriceRS()
    '设定车次
    Dim aszBusID(1 To 1) As String
    Dim i As Integer
    Dim aszTemp(1 To 1) As String
    Dim arsTemp(1 To 1) As Recordset
    
    aszBusID(1) = Trim(txtBusId.Text)
    Set m_rsResultPrice = m_oRoutePriceTable.GetEnvBusPriceRS(dtpOffDate.Value, aszBusID)
    aszTemp(1) = "票价项"
    m_rsAllTicketItem.MoveFirst
    Set arsTemp(1) = m_rsAllTicketItem
    '填充票价记录集
    RTReport.CustomStringCount = aszTemp
    RTReport.CustomString = arsTemp
    RTReport.TemplateFile = App.Path & "\envbusprice.xls"
    RTReport.ShowReport m_rsResultPrice
    F1Book.AllowInCellEditing = True
    F1Book.AllowDelete = False
    F1Book.FixedRows = 1
    RTReport.SetFocus
End Sub

Private Sub txtBusID_LostFocus()
    RefreshBusPrice
End Sub


Private Sub RefreshBusPrice()
    '如果当车次改变了以后,刷新车次票价,并将车次赋给变量m_szBusID
    If Trim(m_szBusID) <> Trim(txtBusId.Text) Or m_dtEnvDate <> Trim(dtpOffDate.Value) Then
        FillBusPriceRS
        m_szBusID = Trim(txtBusId.Text)
        m_dtEnvDate = dtpOffDate.Value
    End If
End Sub
Private Function SetSaveEnabled(Optional pbEnabled As Boolean = True)
    '设置保存是否可用
    cmdok.Enabled = pbEnabled
    m_bChanged = pbEnabled
End Function

Private Function GetTicketTypeID(plRow As Long)
    '得到该行的票种
    m_rsResultPrice.Move plRow - cnPriceItemStartRow, adBookmarkFirst
    GetTicketTypeID = FormatDbValue(m_rsResultPrice!ticket_type)
End Function
Private Sub ModifyHalfPrice(ByVal plRow As Long, ByVal plCol As Long)
    '如果此行为全票行,则修改相应的半票、优惠票等
    Dim nHalfItemCount As Integer
    Dim lRow As Long
    Dim i As Integer, j As Integer
    Dim nTicketType As Integer
    Dim szPriceItem As String '票价项
    
    szPriceItem = GetPriceItem(plCol)
    nHalfItemCount = ArrayLength(m_atHalfItemParam)
    lRow = plRow
    '移至下一条
    With F1Book
        For i = 1 To g_nTicketCountValid - 1
            '向下移
            lRow = lRow + 1
            m_rsResultPrice.Move lRow - cnPriceItemStartRow, adBookmarkFirst
            nTicketType = FormatDbValue(m_rsResultPrice!ticket_type)
            If nTicketType = TP_FullPrice Then
                '如果为全票,则完成
                Exit Sub
            End If
            '查找此票种的参数设置方法
            For j = 1 To nHalfItemCount
                If Val(m_atHalfItemParam(j).szTicketType) = nTicketType And Val(m_atHalfItemParam(j).szTicketItem) = szPriceItem Then
                    Exit For
                End If
            Next j
            If j <= nHalfItemCount Then
                '找到对应的票种
                .TextRC(lRow, plCol) = .TextRC(plRow, plCol) * m_atHalfItemParam(j).sgParam1 + m_atHalfItemParam(j).sgParam2
                '设置尾数处理
                m_oMantissa.SetTailCarry lRow, lRow, plCol
                SetChangeColor lRow, plCol
            End If
        Next i
    End With
End Sub


Private Sub SetChangeColor(plRow As Long, plCol As Long, Optional pbModify As Boolean = True)
    '设置某一格的颜色,表示该行已修改
'    Dim i As Integer
    Dim oCellFormat As F1CellFormat
'    Dim lCol As Long
'    Dim lRow As Long
    Dim lColor As OLE_COLOR
    If pbModify Then
        lColor = vbYellow  '红色
        
    Else
        lColor = vbWhite
    End If
    With F1Book
'        '备份原来的列
'        lRow = .Row
'        lCol = .Col
'        '设置背景
'        .Row = plRow
'        .Col = plCol
        Set oCellFormat = F1Book.GetCellFormat
        If oCellFormat.PatternFG <> lColor Then
            oCellFormat.PatternStyle = 1
            oCellFormat.PatternFG = lColor
            F1Book.SetCellFormat oCellFormat
        End If
'        .Col = cnTotalCol
'        Set oCellFormat = F1Book.GetCellFormat
'        If oCellFormat.PatternFG <> lColor Then
'            oCellFormat.PatternStyle = 1
'            oCellFormat.PatternFG = lColor
'            F1Book.SetCellFormat oCellFormat
'        End If
'        '回赋
'        .Col = lCol
'        .Row = lRow
    End With
    m_abChanged(plRow) = IIf(lColor = vbYellow, True, False)   '标志此行已被修改
    
End Sub


Private Function QueryCancel() As Boolean
    Dim nResult As VbMsgBoxResult
    Dim bCancel As Boolean
    Dim szMsg As String
    '如果修改了,则提示保存
    bCancel = False
    If m_bChanged Then
        szMsg = "票价已经修改,是否要保存？"
        nResult = MsgBox(szMsg, vbYesNoCancel, Me.Caption)
        If nResult = vbYes Then
            '保存票价
            SaveBusPrice
            bCancel = False
        ElseIf nResult = vbCancel Then
            bCancel = True
        ElseIf nResult = vbNo Then
            bCancel = False
        End If
    End If
    QueryCancel = bCancel
End Function


    
Private Function GetPriceItem(plCol As Long) As String
    '得到该列的票价项代码
    m_rsAllTicketItem.Move plCol - cnPriceItemStartCol, adBookmarkFirst
    GetPriceItem = FormatDbValue(m_rsAllTicketItem!price_item)
End Function



Private Function GetModifyStatus(plRow As Long) As Boolean
'    Dim i As Integer
'    Dim oCellFormat As F1CellFormat
'    Dim lCol As Long
'    Dim lRow As Long
'    With F1Book
'        '备份原来的列
'        lRow = .Row
'        .Row = plRow
'        lCol = .Col
'        .Col = cnTotalCol
'        Set oCellFormat = F1Book.GetCellFormat
'        If oCellFormat.PatternFG = vbRed Then  '红色
'            GetModifyStatus = True
'        Else
'            GetModifyStatus = False
'        End If
'    End With
    GetModifyStatus = m_abChanged(plRow)
End Function

Private Sub MarkCellRowModifyStatus(plRow As Long)
    '设置某一行的修改状态
    Dim i As Long
    Dim oCellFormat As F1CellFormat
    Dim lRow As Long
    Dim lCol As Long
    Dim lColor As OLE_COLOR
    
    With F1Book
        lColor = vbWhite
        lRow = .Row
        lCol = .Col
        .Row = plRow
        For i = cnPriceItemStartCol To .MaxCol
            .Col = i
            Set oCellFormat = F1Book.GetCellFormat
            If oCellFormat.PatternFG <> lColor Then
                oCellFormat.PatternStyle = 1
                oCellFormat.PatternFG = lColor
                F1Book.SetCellFormat oCellFormat
            End If
        .Row = lRow
        .Col = lCol
            'SetChangeColor plRow, i, False
        Next i
    End With
End Sub

