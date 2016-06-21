VERSION 5.00
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Object = "{D11DF87C-EFEC-4838-B7E2-15462BF136FB}#2.1#0"; "RTReportLF.ocx"
Begin VB.Form frmBusPrice 
   BackColor       =   &H00E0E0E0&
   Caption         =   "车次票价"
   ClientHeight    =   4920
   ClientLeft      =   2040
   ClientTop       =   2745
   ClientWidth     =   8895
   HelpContextID   =   10000760
   Icon            =   "frmBusPrice.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   8895
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraButton 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H00E0E0E0&
      Height          =   435
      Left            =   60
      TabIndex        =   8
      Top             =   4320
      Width           =   8685
      Begin RTComctl3.CoolButton CoolButton1 
         Height          =   315
         Left            =   3570
         TabIndex        =   10
         Top             =   120
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
         MICON           =   "frmBusPrice.frx":014A
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
         Left            =   4860
         TabIndex        =   3
         Top             =   120
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
         MICON           =   "frmBusPrice.frx":0166
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
         Left            =   6150
         TabIndex        =   4
         Top             =   120
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
         MICON           =   "frmBusPrice.frx":0182
         PICN            =   "frmBusPrice.frx":019E
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
         Left            =   7410
         TabIndex        =   5
         Top             =   120
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
         MICON           =   "frmBusPrice.frx":0538
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
      TabIndex        =   6
      Top             =   0
      Width           =   8775
      Begin FText.asFlatTextBox txtBusID 
         Height          =   300
         Left            =   1260
         TabIndex        =   1
         Top             =   240
         Width           =   1380
         _ExtentX        =   2434
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
      Begin VB.Label lblOffTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "发车时间:12:12"
         Height          =   180
         Left            =   2805
         TabIndex        =   9
         Top             =   300
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车次代码(&I):"
         Height          =   180
         Index           =   0
         Left            =   150
         TabIndex        =   0
         Top             =   300
         Width           =   1080
      End
      Begin VB.Label lblRoute 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "运行线路:"
         Height          =   180
         Left            =   4785
         TabIndex        =   7
         Top             =   300
         Width           =   810
      End
   End
   Begin RTReportLF.RTReport RTReport 
      Height          =   3525
      Left            =   90
      TabIndex        =   2
      Top             =   720
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   6218
   End
End
Attribute VB_Name = "frmBusPrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const cnPriceItemStartCol = 8        '可编辑票价项的起始列
Const cnPriceItemStartRow = 2       '可编辑票价项的起始行
Const cnTotalCol = 7 '总计所在列

Public m_szBusID As String
'Private m_bDisplayOnly As Boolean '是否只显示

Private WithEvents F1Book As TTF160Ctl.F1Book
Attribute F1Book.VB_VarHelpID = -1
Private m_rsResultPrice As Recordset      '票价表结果记录集
Private m_rsAllTicketItem As Recordset    '所有的票种类型
Private m_oRoutePriceTable As RoutePriceTable
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
    SaveFormPos Me
        On Error GoTo ErrorHandle
    If Not QueryCancel Then
        Cancel = 0
    Else
        Cancel = 1
    End If
    If Cancel <> 1 Then
        Set m_oRoutePriceTable = Nothing
        Set m_rsResultPrice = Nothing
        Set m_rsAllTicketItem = Nothing
        Set m_oMantissa = Nothing
    End If
    Exit Sub
ErrorHandle:
    Cancel = 1
    ShowErrorMsg

End Sub

Private Sub Form_Load()
'    Dim i As Integer
    Dim oHalfTicket As New HalfTicketPrice
    Dim oPriceMan As New STPrice.TicketPriceMan
'    AlignFormPos Me
    On Error GoTo ErrorHandle
    txtBusID.Enabled = True
    Set m_oRoutePriceTable = CreateObject("stprice.RoutePriceTable")
    m_oRoutePriceTable.Init g_oActiveUser
    m_oRoutePriceTable.Identify g_szExePriceTable
    
    txtBusID.Text = m_szBusID
    SetSaveEnabled False
'    cmdOk.Enabled = False
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
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub tmStart_Timer()
On Error GoTo ErrHandle
    tmStart.Enabled = False
    SetBusy
    '填充车次票价记录集
    FillBusPriceRS
'    RTReport.Enabled = True
    InitMantissa '初始化对象的属性
    SetNormal
    Exit Sub
ErrHandle:
'    RTReport.Enabled = True
    SetNormal
    ShowErrorMsg
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

Private Sub cmdOk_Click()
    SaveBusPrice
    RTReport.SetFocus
End Sub

Private Sub SaveBusPrice()
    '保存车次票价
    Dim tDetailInfo(1 To 1) As TBusPriceDetailInfo
    Dim i As Long, k As Long
    Dim bModify As Boolean '标志此行是否被修改
    Dim szPriceItem As String
    On Error GoTo ErrorHandle
    With F1Book
        WriteProcessBar True, 1, .LastRow, "正在保存车次票价表"
        F1Book.EndEdit
        m_rsResultPrice.MoveFirst
        For i = cnPriceItemStartRow To .LastRow
            '得到修改状态
            bModify = GetModifyStatus(i)
            
            If bModify Then
                '如果为已修改或者为新增状态
                '上车站
                tDetailInfo(1).szSellStationID = m_rsResultPrice!sell_station_id
                '车次
                tDetailInfo(1).szbusID = m_rsResultPrice!bus_id
                '距离
                tDetailInfo(1).sgMileage = m_rsResultPrice!Mileage
                '车型
                tDetailInfo(1).szVehicleModel = m_rsResultPrice!vehicle_type_code
                '座位类型
                tDetailInfo(1).szSeatTypeID = m_rsResultPrice!seat_type_id
                '站点
                tDetailInfo(1).szStationID = m_rsResultPrice!station_id
                '票种
                tDetailInfo(1).nTicketType = m_rsResultPrice!ticket_type
                '总价
                tDetailInfo(1).sgTotalPrice = .NumberRC(i, cnTotalCol)
                '基本运价
                tDetailInfo(1).sgBaseCarriage = .NumberRC(i, cnPriceItemStartCol)
'                '站点车次序号
                tDetailInfo(1).nSerialNo = m_rsResultPrice!station_serial_no
                '各票价项
                '由于其他不用项默认为0,所以只赋显示的票价项
                For k = cnPriceItemStartCol + 1 To .MaxCol
                    szPriceItem = GetPriceItem(k)
                    tDetailInfo(1).asgItem(CInt(szPriceItem)) = .NumberRC(i, k)
                Next
                m_oRoutePriceTable.ModifySpecifyBusPrice tDetailInfo
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
''更改计划车次票价表
'Private Sub ModifyPrice(padbNewPriceItem() As Double)
'    Dim tDetailInfo(1 To 1) As TBusPriceDetailInfo
'    Dim i As Integer
'    tDetailInfo(1).szBusID = m_rsResultPrice("bus_id").Value
'    tDetailInfo(1).szStationID = m_rsResultPrice("station_id").Value
'    tDetailInfo(1).nTicketType = m_rsResultPrice("ticket_type").Value
'    tDetailInfo(1).szVehicleModel = m_rsResultPrice("vehicle_type_code").Value
'    tDetailInfo(1).szSeatTypeID = m_rsResultPrice("seat_type_id").Value
'    tDetailInfo(1).nSerialNo = m_rsResultPrice("station_serial_no").Value
'    tDetailInfo(1).sgMileage = m_rsResultPrice("mileage").Value
'    tDetailInfo(1).sgTotalPrice = padbNewPriceItem(0)
'    tDetailInfo(1).sgBaseCarriage = padbNewPriceItem(1)
'    m_rsAllTicketItem.MoveFirst
'    m_rsAllTicketItem.MoveNext
'    i = 2
'    While Not m_rsAllTicketItem.EOF
'        tDetailInfo(1).asgItem(Val(m_rsAllTicketItem.Fields("price_item"))) = padbNewPriceItem(i)
'        m_rsAllTicketItem.MoveNext
'        i = i + 1
'    Wend
'
'
'    m_oRoutePriceTable.ModifySpecifyBusPrice tDetailInfo
'End Sub
Private Sub cmdRefresh_Click()
    m_szBusID = txtBusID.Text
    RefreshBus
    tmStart_Timer
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
Private Sub InitMantissa()
'    初始化对象的属性
    m_oMantissa.MaxCol = F1Book.MaxCol
    m_oMantissa.oF1Book = RTReport.CellObject
    m_oMantissa.oPriceTable = m_oRoutePriceTable
    m_oMantissa.PriceItemStartCol = cnPriceItemStartCol
    m_oMantissa.PriceItemStartRow = cnPriceItemStartRow
    m_oMantissa.PriceRs = m_rsResultPrice
    m_oMantissa.TotalCol = cnPriceItemStartCol - 1
    m_oMantissa.UseItemRs = m_rsAllTicketItem
End Sub

'Private Sub f1book_Click(ByVal nRow As Long, ByVal nCol As Long)
''可编辑时显示编辑条
'    If nRow < cnPriceItemStartRow Or nCol < cnPriceItemStartCol Then
'        F1Book.ShowEditBar = False
'    Else
'        F1Book.ShowEditBar = True
'    End If
'End Sub

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
    
Private Function GetPriceItem(plCol As Long) As String
    '得到该列的票价项代码
    m_rsAllTicketItem.Move plCol - cnPriceItemStartCol, adBookmarkFirst
    GetPriceItem = FormatDbValue(m_rsAllTicketItem!price_item)
End Function

Private Function SetSaveEnabled(Optional pbEnabled As Boolean = True)
    '设置保存是否可用
    cmdOk.Enabled = pbEnabled
    m_bChanged = pbEnabled
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


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

'刷新车次信息
Private Sub RefreshBus()
On Error GoTo ErrHandle
    Dim oBus As New Bus
    oBus.Init g_oActiveUser
    oBus.Identify m_szBusID
    
    If oBus.BusType = TP_ScrollBus Then
        lblOffTime.Caption = "间隔时间:" & oBus.ScrollBusCheckTime & " 分钟"
    Else
        lblOffTime.Caption = "发车时间:" & Format(oBus.StartUpTime, "hh:mm")
    End If
    lblRoute.Caption = "运行线路:" & oBus.RouteName
    

    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Const cnMargin = 30
    fraInfo.Width = Me.ScaleWidth - 2 * cnMargin
    RTReport.Width = fraInfo.Width - cnMargin / 2
    fraButton.Width = fraInfo.Width
    RTReport.Height = Me.ScaleHeight - fraInfo.Height - fraButton.Height - 2 * cnMargin
    fraButton.Top = RTReport.Top + RTReport.Height
End Sub


Private Sub txtBusID_ButtonClick()
On Error GoTo ErrHandle
    Dim oShell As New CommDialog
    Dim aszTmp() As String
    oShell.Init g_oActiveUser
    aszTmp = oShell.SelectBus
    Set oShell = Nothing
    If ArrayLength(aszTmp) = 0 Then Exit Sub
    txtBusID.Text = Trim(aszTmp(1, 1))
    txtBusID_LostFocus
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

'得到车次票价记录集
Private Sub FillBusPriceRS()
    '设定车次
    Dim aszBusID(1 To 1) As String
    Dim nTemp As Long
    Dim i As Integer
    Dim ttBusVehicleSeat() As TBusVehicleSeatType
    Dim m_oTicketPriceMan As New TicketPriceMan
    Dim aszTmp() As String
    Dim m_aszVehicleType() As String '存在的车型
    Dim m_aszSeatType() As String '存在的座位类型
    aszBusID(1) = Trim(txtBusID.Text)
    
        '选择了只打开存在的车型与座位类型
        m_oTicketPriceMan.Init g_oActiveUser
        ttBusVehicleSeat = m_oTicketPriceMan.GetAllBusVehicleTypeSeatType(aszBusID)
        nTemp = ArrayLength(ttBusVehicleSeat)
        If nTemp > 0 Then
           ReDim m_aszSeatType(1 To nTemp)
           ReDim m_aszVehicleType(1 To nTemp)
        End If
        For i = 1 To nTemp
            m_aszSeatType(i) = ttBusVehicleSeat(i).szSeatTypeID
            m_aszVehicleType(i) = ttBusVehicleSeat(i).szVehicleTypeCode
        Next i
        
        '得到计划车次票价
        Set m_rsResultPrice = m_oRoutePriceTable.GetSpecifyBusPriceRS(aszBusID, m_aszVehicleType, m_aszSeatType, aszTmp)
    
    Dim aszTemp(1 To 1) As String
    Dim arsTemp(1 To 1) As Recordset
    aszTemp(1) = "票价项"
    m_rsAllTicketItem.MoveFirst
    Set arsTemp(1) = m_rsAllTicketItem
    '填充票价记录集
    RTReport.LeftLabelVisual = True
    RTReport.TopLabelVisual = True
    RTReport.CustomStringCount = aszTemp
    RTReport.CustomString = arsTemp
    RTReport.TemplateFile = App.Path & "\busprice.xls"
    RTReport.ShowReport m_rsResultPrice
    F1Book.AllowInCellEditing = True
    F1Book.AllowDelete = False
    F1Book.Col = cnPriceItemStartCol
    F1Book.FixedRows = 1
    If F1Book.LastRow >= 2 Then F1Book.Row = 2
    If F1Book.MaxRow > 0 Then ReDim m_abChanged(1 To F1Book.MaxRow)
    RTReport.SetFocus
End Sub

Private Sub txtBusID_LostFocus()
    '如果当车次改变了以后,刷新车次票价,并将车次赋给变量m_szBusID
    If Trim(m_szBusID) <> Trim(txtBusID.Text) Then
        FillBusPriceRS
        m_szBusID = Trim(txtBusID.Text)
    End If
End Sub


