VERSION 5.00
Object = "{D11DF87C-EFEC-4838-B7E2-15462BF136FB}#2.1#0"; "RTReportLF.ocx"
Begin VB.Form frmModifyREPrice 
   Caption         =   "打开环境票价"
   ClientHeight    =   4080
   ClientLeft      =   3240
   ClientTop       =   4335
   ClientWidth     =   5490
   HelpContextID   =   1001601
   Icon            =   "frmModifyREPrice.frx":0000
   LinkTopic       =   "frmModifyREPrice"
   MDIChild        =   -1  'True
   ScaleHeight     =   4080
   ScaleWidth      =   5490
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   150
      Top             =   180
   End
   Begin RTReportLF.RTReport RTReport 
      Height          =   3495
      Left            =   75
      TabIndex        =   0
      Top             =   15
      Width           =   4590
      _ExtentX        =   8096
      _ExtentY        =   6165
   End
End
Attribute VB_Name = "frmModifyREPrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************
'* Source File Name:frmModifyREPrice.frm
'* Project Name:PSTBusMan.vbp
'* Engineer:陈峰
'* Date Generated:2002/09/03
'* Last Revision Date:2002/09/11
'* Brief Description:环境票价表窗口
'* Relational Document:
'**********************************************************

Option Explicit
Const cszTemplateFile = "环境车次票价模板.xls"

Const cnPriceItemStartCol = 10
Const cnPriceItemStartRow = 2
Const cnTotalCol = 9 '总计所在列

Public m_eFormStatus As EFormStatus

Private m_oREScheme As New STReSch.REScheme
Private m_oRoutePriceTable As New RoutePriceTable
Private WithEvents F1Book As TTF160Ctl.F1Book
Attribute F1Book.VB_VarHelpID = -1
Private m_rsAllTicketItem As Recordset '存放所有使用的票价项
Private m_rsResultPrice As Recordset '存放所有的打开的票价
Private m_lRange As Long '为了写进度条时用到
Private m_oMantissa As New clMantissa '尾数处理对象
Private m_atHalfItemParam() As THalfTicketItemParam '半票及优惠票票价项计算参数
Private m_bChanged As Boolean '标志是否改变

Private m_bCalHalfPrice As Boolean '标志是否需要计算半价

Private m_aszBusID() As String '存放选择的车次
Private m_dyBusDate As Date '选择的车次日期


Private m_abChanged() As Boolean '存放每一行是否修改的标志

Private Sub F1Book_DblClick(ByVal nRow As Long, ByVal nCol As Long)
    F1Book.StartEdit False, True, False
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
        F1Book.AllowInCellEditing = False
    Else
        F1Book.AllowInCellEditing = True
    End If
End Sub

Private Sub Form_Activate()
    SetMenuEnabled True
    MDIScheme.SetPrintEnabled True
End Sub

Private Sub Form_Deactivate()
    SetMenuEnabled False
    MDIScheme.SetPrintEnabled False
End Sub

Private Sub Form_Load()
    '初始化对象
'    RTReport.Enabled = False
End Sub

Private Sub Form_Resize()
    '设置控件位置
    RTReport.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrorHandle
    If Not QueryCancel Then
        Cancel = 0
    Else
        Cancel = 1
    End If
    If Cancel <> 1 Then
        SetMenuEnabled False
        MDIScheme.SetPrintEnabled False
        Set m_oMantissa = Nothing
    End If
    Exit Sub
ErrorHandle:
    Cancel = 1
    ShowErrorMsg
End Sub

Private Sub RTReport_SetProgressRange(ByVal lRange As Variant)
    m_lRange = lRange
End Sub

Private Sub RTReport_SetProgressValue(ByVal lValue As Variant)
    If lValue = m_lRange Then
        WriteProcessBar False, lValue, m_lRange, ""
    Else
        WriteProcessBar , lValue, m_lRange, "正在填充票价"
    End If
End Sub

Private Sub Timer1_Timer()
    '初始化
    Dim oPriceMan As New STPrice.TicketPriceMan
    Dim oHalfTicket As New STPrice.HalfTicketPrice
    On Error GoTo ErrorHandle
    Timer1.Enabled = False
'    Me.Hide
    '得到所有的票价项
    oPriceMan.Init g_oActiveUser
    Set m_rsAllTicketItem = oPriceMan.GetAllTicketItemRS(TP_PriceItemUse)
    m_oRoutePriceTable.Init g_oActiveUser
    m_oRoutePriceTable.Identify g_szExePriceTable
    '得到所有票种的半票计算参数
    oHalfTicket.Init g_oActiveUser
    m_atHalfItemParam = oHalfTicket.GetItemParam(0, g_szExePriceTable, TP_PriceItemUse)
    m_oREScheme.Init g_oActiveUser
    If m_eFormStatus = EFS_Modify Then
        '打开
        Dim bCancel As Boolean
        ShowOpenDialog True, bCancel
    End If
    If bCancel Then
        Unload Me
        Exit Sub
    End If
    If F1Book.MaxRow > 0 Then ReDim m_abChanged(1 To F1Book.MaxRow)
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Public Sub ShowOpenDialog(Optional pbFirst As Boolean = False, Optional pbCancel As Boolean)
    Dim oShell As New CommDialog
    Dim dyDate As Date
'    Dim aszBusID() As String
    dyDate = Date
    oShell.Init g_oActiveUser
    m_aszBusID = oShell.SelectREBus(dyDate, True, True)
    m_dyBusDate = dyDate
    If ArrayLength(m_aszBusID) > 0 Then
        If Not QueryCancel Then
'            Me.Show
            '是否按了确定
            OpenBusPrice
        End If
        If pbFirst Then
            InitMantissa
        End If
    Else
        pbCancel = True
    End If

End Sub

Private Function QueryCancel() As Boolean
    Dim nResult As VbMsgBoxResult
    Dim bCancel As Boolean
    Dim szMsg As String
    '如果修改了,则提示保存
    bCancel = False
    If m_bChanged Then
        If m_eFormStatus = EFS_AddNew Then
            szMsg = "新增的票价,是否要保存？"
        Else
            szMsg = "票价已经修改,是否要保存？"
        End If
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

Private Sub OpenBusPrice()
    '打开车次票价
    Dim atDetailInfo() As TBusPriceDetailInfo
    Dim aszBusID() As String
    Dim aszSeatType() As String
    Dim aszVehicleModel() As String
    Dim aszTemp(1 To 1) As String
    Dim arsTemp(1 To 1) As Recordset
    On Error GoTo ErrorHandle
    m_eFormStatus = EFS_Modify
    SetSaveEnabled False
    SetBusy
    Set m_rsResultPrice = m_oREScheme.GetBusTicketInfoRS(m_dyBusDate, m_aszBusID)
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
    F1Book.FixedRows = 1
    If F1Book.LastRow >= 2 Then F1Book.Row = 2
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
                oRebus.Identify m_rsResultPrice!bus_id, m_dyBusDate
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

Private Sub SetMenuEnabled(pbEnabled As Boolean)
    '设置菜单的可用性
    With MDIScheme.abMenuTool
        .Bands("mnu_TicketPrice").Tools("mnu_TicketPriceMan_Open").Enabled = pbEnabled
        .Bands("mnu_TicketPrice").Tools("mnu_TicketPriceMan_Modify").Enabled = pbEnabled
    End With
End Sub

Private Sub SetChangeColor(plRow As Long, plCol As Long, Optional pbModify As Boolean = True)
    '设置某一格的颜色,表示该行已修改
    Dim i As Integer
    Dim oCellFormat As F1CellFormat
    Dim lCol As Long
    Dim lRow As Long
    Dim lColor As OLE_COLOR
    If pbModify Then
        lColor = vbYellow '黄色，原来的 红色负数时会看不到
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
    m_abChanged(plRow) = IIf(lColor = vbYellow, True, False)  '标志此行已被修改
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
'    Dim lRowTemp As Long
'    Dim lColTemp As Long
'
    szPriceItem = GetPriceItem(plCol)
    nHalfItemCount = ArrayLength(m_atHalfItemParam)
    lRow = plRow
    '移至下一条
    With F1Book
'        lRowTemp = .Row
'        lColTemp = .Col
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
'                .Row = lRow
'                .Col = plCol
                '不用此过程中的注释的代码 , 是由于控件存在BUG, 一旦在EndEdit事件中改变.row或.col后, 按上下键就会无效
                '本来此过程中的注释的代码是为了弥补用.textRc时会循环触发EndEdit事件,而调用此过程时参数用的是.row .
'                .Text = .TextRC(plRow, plCol) * m_atHalfItemParam(j).sgParam1 + m_atHalfItemParam(j).sgParam2
                .TextRC(lRow, plCol) = .TextRC(plRow, plCol) * m_atHalfItemParam(j).sgParam1 + m_atHalfItemParam(j).sgParam2
                '设置尾数处理
                m_oMantissa.SetTailCarry lRow, lRow, plCol
                SetChangeColor lRow, plCol
            End If
        Next i
'        .Row = lRowTemp
'        .Col = lColTemp
        
    End With
End Sub
    
Private Function GetPriceItem(plCol As Long) As String
    '得到该列的票价项代码
    m_rsAllTicketItem.Move plCol - cnPriceItemStartCol, adBookmarkFirst
    GetPriceItem = m_rsAllTicketItem!price_item
End Function

Private Function SetSaveEnabled(Optional pbEnabled As Boolean = True)
    '设置保存是否可用
    MDIScheme.abMenuTool.Bands("mnu_TicketPrice").Tools("mnu_TicketPriceMan_Save").Enabled = pbEnabled
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
    With F1Book
    
    For i = cnTotalCol To .MaxCol
        SetChangeColor plRow, i, False
    Next i
    End With
End Sub



Public Sub BatchModify()
    '批量修改
    Dim aszTemp() As String
    Dim szKey1 As String
    Dim szKey2 As String
    Dim sgMul As Single
    Dim sgAdd As Single
    Dim i As Integer
    Dim j As Long
    Dim nCount As Integer
    Dim abTicketType() As Boolean
    
    
    Dim lIndex1 As Long
    Dim lIndex2 As Long
    Dim nThisTicketType  As Integer
    
    '嘉兴增加
    Dim aszBusID() As String
    aszBusID = m_aszBusID
    Dim szStationName As String

    '得到站点
    Dim moBusProject As New REBus
    Dim aszBus() As String
    moBusProject.Init g_oActiveUser
    
    moBusProject.Identify aszBusID(1, 1), CDate(aszBusID(1, 2))
    
    
    
'    aszBus = moBusProject.GetAllBus(Trim(aszBusID(1, 1)), , , True)
    
'    Dim m_oBus As New Bus
'
'    m_oBus.Init g_oActiveUser
'
'    m_oBus.Identify Trim(aszBusID(1, 1))
'
    frmGetFormula.m_szRouteID = moBusProject.Route
    
    
    frmGetFormula.Show vbModal
    If Not frmGetFormula.m_bOk Then Exit Sub
    '按了OK
    aszTemp = frmGetFormula.GetParam
    abTicketType = frmGetFormula.GetSelectTicketType
    nCount = ArrayLength(aszTemp)
    With F1Book
        For i = 1 To nCount
            '得到各参数值
            szKey1 = aszTemp(i, 1)
            szKey2 = aszTemp(i, 2)
            sgMul = aszTemp(i, 3)
            sgAdd = aszTemp(i, 4)
            
'            得到各票价项的列的位置
            lIndex1 = GetPriceItemEnablePosition(szKey1)
            If szKey1 <> szKey2 Then
                lIndex2 = GetPriceItemEnablePosition(szKey2)
            Else
                lIndex2 = lIndex1
            End If
            
            
            
            For j = cnPriceItemStartRow To .LastRow
                '得到票种
                nThisTicketType = GetTicketTypeID(j)
                If abTicketType(nThisTicketType) Then
                    .TextRC(j, lIndex1) = Round(.TextRC(j, lIndex2) * sgMul + sgAdd + 0.001, 2)
                    .Row = j
                    .Col = lIndex1
                    SetChangeColor j, lIndex1
                End If
'                If bModifyAll Or .IsSelectedCell(lIndex1, j - 1) Then
                '如果修改所有的或选择了所有的CELL

'                    HalfTicketItemParam = GetHalfTicketItemParam(nThisTicketType, m_tHalfItemParam)
'                    If nThisTicketType = TP_FullPrice Then
                        '如果票种为全票
'                        If szKey1 <> szKey2 Then
'                            .DoGetCellData lIndex1, j - 1, sgValue
'                        End If
'                        .DoGetCellData lIndex2, j - 1, sgValue1
'                        .DoGetCellData cnTotalPrice, j - 1, dbTempTotal
'
'                        sgResult = Round(sgValue1 * sgMul + sgAdd + 0.001, m_nDetailLength)
'                        If szKey1 <> szKey2 Then
'                            .DoSetCellData cnTotalPrice, j - 1, dbTempTotal + sgResult - sgValue
'                        Else
'                            .DoSetCellData cnTotalPrice, j - 1, dbTempTotal + sgResult - sgValue1
'                        End If
'                        .DoSetCellData lIndex1, j - 1, sgResult
'                        .DoSetCellColor lIndex1, j - 1, RGB(0, 0, 0), m_lModifiedColor
'                    ElseIf bnModifyTicketType(nThisTicketType) = True Then
'                        '如果为特殊票
'                        If lIndex1 < nRoutePriceItem + cnPriceItemStartCol Then
'                           dbParam1 = HalfTicketItemParam(lIndex1 - cnPriceItemStartCol + 1).sgParam1
'                           dbParam2 = HalfTicketItemParam(lIndex1 - cnPriceItemStartCol + 1).sgParam2
'                        Else
'                           dbParam1 = HalfTicketItemParam(lIndex1 - cnPriceItemStartCol + 13 - nRoutePriceItem).sgParam1
'                           dbParam2 = HalfTicketItemParam(lIndex1 - cnPriceItemStartCol + 13 - nRoutePriceItem).sgParam2
'                        End If
'                        dbHalf = Round(dbParam1 * sgResult + dbParam2 + 0.001, m_nDetailLength)
'                        .DoSetCellData lIndex1, j - 1, dbHalf
'                    End If
'                End If
            Next
        Next
        m_oMantissa.SetTailCarry cnPriceItemStartRow, .MaxRow, , True
        SetSaveEnabled True
    End With
End Sub


Private Function GetPriceItemEnablePosition(pszPriceItem As String) As Integer
    '得到票价项在可用票价项中的位置
    Dim i As Integer
    m_rsAllTicketItem.MoveFirst
    For i = 1 To m_rsAllTicketItem.RecordCount
        If FormatDbValue(m_rsAllTicketItem!price_item) = pszPriceItem Then
            Exit For
        End If
        m_rsAllTicketItem.MoveNext
    Next i
    If i <= m_rsAllTicketItem.RecordCount Then
        '如果找到则退出.
        GetPriceItemEnablePosition = m_rsAllTicketItem.Bookmark + cnPriceItemStartCol - 1
    End If
End Function


