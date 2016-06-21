VERSION 5.00
Object = "{D11DF87C-EFEC-4838-B7E2-15462BF136FB}#2.1#0"; "RTReportLF.ocx"
Begin VB.Form frmModifyBusPrice 
   Caption         =   "车次票价"
   ClientHeight    =   4395
   ClientLeft      =   4980
   ClientTop       =   4185
   ClientWidth     =   6060
   HelpContextID   =   1001401
   Icon            =   "frmModifyBusPrice.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4395
   ScaleWidth      =   6060
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   1740
      Top             =   1380
   End
   Begin RTReportLF.RTReport RTReport 
      Height          =   3495
      Left            =   480
      TabIndex        =   0
      Top             =   405
      Width           =   4590
      _ExtentX        =   8096
      _ExtentY        =   6165
   End
End
Attribute VB_Name = "frmModifyBusPrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************
'* Source File Name:frmModifyBusPrice.frm
'* Project Name:PSTBusMan.vbp
'* Engineer:陈峰
'* Date Generated:2002/09/03
'* Last Revision Date:2002/09/10
'* Brief Description:车次附加票价窗口
'* Relational Document:
'**********************************************************

Option Explicit
Const cszTemplateFile = "车次票价模板.xls"
Const cnPriceItemStartCol = 11
Const cnPriceItemStartRow = 2
Const cnTotalCol = 10 '总计所在列

Public m_eFormStatus As EFormStatus


Private m_oRoutePriceTable As New RoutePriceTable
Private WithEvents F1Book As TTF160Ctl.F1Book
Attribute F1Book.VB_VarHelpID = -1
Private m_rsAllTicketItem As Recordset '存放所有使用的票价项
Private m_rsResultPrice As Recordset '存放所有的打开的票价
Private m_lRange As Long '为了写进度条时用到
Private m_oMantissa As New clMantissa '尾数处理对象
Private m_atHalfItemParam() As THalfTicketItemParam '半票及优惠票票价项计算参数
Private m_szPriceTableID As String
Private m_bChanged As Boolean '标志是否改变
'Private m_bCalHalfPrice As Boolean '标志是否需要计算半价


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
'        F1Book.ShowEditBar = False
        F1Book.AllowInCellEditing = False
    Else
'        F1Book.ShowEditBar = True
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
    Dim oPriceMan As New stprice.TicketPriceMan
    Dim oHalfTicket As New stprice.HalfTicketPrice
    On Error GoTo ErrorHandle
    Timer1.Enabled = False
    '得到所有的票价项
    oPriceMan.Init g_oActiveUser
    Set m_rsAllTicketItem = oPriceMan.GetAllTicketItemRS(TP_PriceItemUse)
    '得到所有票种的半票计算参数
    oHalfTicket.Init g_oActiveUser
    m_atHalfItemParam = oHalfTicket.GetItemParam(0, g_szExePriceTable, TP_PriceItemUse)
    m_oRoutePriceTable.Init g_oActiveUser
    Select Case m_eFormStatus
    Case EFS_AddNew
        '新增车次
        ShowAddDialog True
    Case EFS_Modify
        '打开
        ShowOpenDialog True
        
    End Select
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Public Sub ShowAddDialog(Optional pbFirst As Boolean = False)
    frmShowBus.m_eFormStatus = EFS_AddNew
    frmShowBus.Show vbModal
    If frmShowBus.m_bOk Then
'        Me.Show
        '新增车次票价
        AddBusPrice
        If pbFirst Then
            InitMantissa
        End If
        If Not F1Book Is Nothing Then
            m_oMantissa.SetTailCarry cnPriceItemStartRow, F1Book.MaxRow, 0, True
        Else
            Unload Me
        End If
'        If Not QueryCancel Then
'            '是否按了确定
'            OpenBusPrice
'            InitMantissa
'        End If
    ElseIf pbFirst Then
        Unload Me
    End If

End Sub


Public Sub ShowOpenDialog(Optional pbFirst As Boolean = False)
    frmShowBus.m_eFormStatus = EFS_Show
    frmShowBus.Show vbModal
    If frmShowBus.m_bOk Then
        If Not QueryCancel Then
'            Me.Show
            '是否按了确定
            OpenBusPrice
        End If
        If pbFirst Then
            InitMantissa
        End If
    ElseIf pbFirst Then
        Unload Me
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

Private Sub AddBusPrice()
    '显示新增出来但未保存到数据库中的数据
    Dim atDetailInfo() As TBusPriceDetailInfo
    Dim aszBusID() As String
    Dim aszSeatType() As String
    Dim aszVehicleModel() As String
    Dim aszTemp(1 To 1) As String
    Dim arsTemp(1 To 1) As Recordset
    Dim i As Long
    Dim j As Integer
    Dim n As Integer
    Dim nTemp As Long
    Dim nCount As Integer
    Dim nBus As Integer
    Dim nSeatType As Integer
    Dim nVehicleType As Integer
    Dim atBusVehicleSeat() As TBusVehicleSeatType
    Dim bIsEmpty As Boolean '标志是否只生成空票价
    
    On Error GoTo ErrorHandle
    
    SetBusy
    aszBusID = frmShowBus.GetBusID
    aszVehicleModel = frmShowBus.GetVehicleType
    aszSeatType = frmShowBus.GetSeatType
    m_szPriceTableID = frmShowBus.GetPriceTableID
    '将各数组集合起来,并赋给类型
    atBusVehicleSeat = ConvertTypeFromArray(aszBusID, aszVehicleModel, aszSeatType)
    '将各数组合并为一个类型
    
    bIsEmpty = frmShowBus.IsEmpty
    m_oRoutePriceTable.Identify Trim(m_szPriceTableID)
    If bIsEmpty Then
        '只生成空票价,可以进行直接输入,不需要再录入基本运价率等等
        Set m_rsResultPrice = m_oRoutePriceTable.MakeEmptyBusPriceRS(atBusVehicleSeat)
    Else
        '按输入的条件进行生成票价
        Set m_rsResultPrice = m_oRoutePriceTable.MakeSpecifyBusPriceRS(atBusVehicleSeat)
    End If
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
'    For i = 2 To m_rsResultPrice.RecordCount + 1
'        SetChangeColor i, cnTotalCol
'    Next i
    
    m_bChanged = True
    F1Book.AllowInCellEditing = True
    F1Book.AllowDelete = False
    F1Book.Col = cnPriceItemStartCol
    F1Book.FixedRows = 1
    If F1Book.LastRow >= 2 Then F1Book.Row = 2
    m_bChanged = True
    SetSaveEnabled True
    If F1Book.MaxRow > 0 Then ReDim m_abChanged(1 To F1Book.MaxRow)

    SetNormal
    Exit Sub
ErrorHandle:
    ShowErrorMsg
    SetNormal
End Sub


Private Sub OpenBusPrice()
    '打开车次票价
    Dim atDetailInfo() As TBusPriceDetailInfo
    Dim aszBusID() As String
    Dim aszSeatType() As String
    Dim aszSellStation() As String
    Dim aszVehicleModel() As String
    Dim aszTemp(1 To 1) As String
    Dim arsTemp(1 To 1) As Recordset
    On Error GoTo ErrorHandle
    m_eFormStatus = EFS_Modify
    m_bChanged = False
    SetBusy
    aszBusID = frmShowBus.GetBusID
    aszVehicleModel = frmShowBus.GetVehicleType
    aszSeatType = frmShowBus.GetSeatType
    aszSellStation = frmShowBus.GetSellStation
    
    m_szPriceTableID = frmShowBus.GetPriceTableID
    m_oRoutePriceTable.Identify Trim(m_szPriceTableID)
    Set m_rsResultPrice = m_oRoutePriceTable.GetSpecifyBusPriceRS(aszBusID, aszVehicleModel, aszSeatType, aszSellStation)
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
    If F1Book.MaxRow > 0 Then ReDim m_abChanged(1 To F1Book.MaxRow)
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
    Dim atBusInfo() As TBusVehicleSeatType
    
    On Error GoTo ErrorHandle
    With F1Book
        WriteProcessBar True, 1, .LastRow, "正在保存车次票价表"
        m_rsResultPrice.MoveFirst
        '得到所有的东东
        If m_eFormStatus = EFS_AddNew Then
            '如果为新增
            If m_rsResultPrice.RecordCount > 0 Then
                ReDim atBusInfo(1 To m_rsResultPrice.RecordCount)
                For i = 1 To m_rsResultPrice.RecordCount
                    atBusInfo(i).szbusID = m_rsResultPrice!bus_id
                    atBusInfo(i).szVehicleTypeCode = m_rsResultPrice!vehicle_type_code
                    atBusInfo(i).szSeatTypeID = m_rsResultPrice!seat_type_id
                    atBusInfo(i).szSellStationID = m_rsResultPrice!sell_station_id
                Next i
                m_oRoutePriceTable.DeleteBusPrice atBusInfo
                m_rsResultPrice.MoveFirst
            End If
        End If
        '**********如果为新增状态,则需要先删除原先的数据,否则可能会出现站点不匹配
        
        For i = cnPriceItemStartRow To .LastRow
            '得到修改状态
            If m_eFormStatus = EFS_AddNew Then
                bModify = True
            Else
                bModify = GetModifyStatus(i)
            End If
            If bModify Then
                '如果为已修改或者为新增状态
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
                '发车站
                tDetailInfo(1).szSellStationID = m_rsResultPrice!sell_station_id
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
'        SetProgressBarVisible False
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
'    Dim i As Integer
    Dim oCellFormat As F1CellFormat
'    Dim lCol As Long
'    Dim lRow As Long
    Dim lColor As OLE_COLOR
    If pbModify Then
        lColor = vbYellow  '黄色
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
    If F1Book Is Nothing Then Exit Sub
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
    '
'    Dim lRowTemp As Long
'    Dim lColTemp As Long
    
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
'                .Text = .TextRC(plRow, plCol) * m_atHalfItemParam(j).sgParam1 + m_atHalfItemParam(j).sgParam2
                '不用此过程中的注释的代码 , 是由于控件存在BUG, 一旦在EndEdit事件中改变.row或.col后, 按上下键就会无效
                '本来此过程中的注释的代码是为了弥补用.textRc时会循环触发EndEdit事件,而调用此过程时参数用的是.row .
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
    GetPriceItem = FormatDbValue(m_rsAllTicketItem!price_item)
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


'Public Sub BatchModify()
'    '批量修改
'    Dim aszTemp() As String
'    Dim szKey1 As String
'    Dim szKey2 As String
'    Dim sgMul As Single
'    Dim sgAdd As Single
'    Dim sgMileage As Single
'    Dim i As Integer
'    Dim j As Long
'    Dim nCount As Integer
'    Dim abTicketType() As Boolean
'
'
'    Dim lIndex1 As Long
'    Dim lIndex2 As Long
'    Dim nThisTicketType  As Integer
'
'    '嘉兴增加
'    Dim aszBusID() As String
'    aszBusID = frmShowBus.GetBusID
'    Dim szStationName As String
'
'    '得到站点
'    Dim moBusProject As New BusProject
'    Dim aszBus() As String
'    moBusProject.Init g_oActiveUser
'
'    moBusProject.Identify
'
'    aszBus = moBusProject.GetAllBus(Trim(aszBusID(1)), , , True)
'
'    Dim m_oBus As New Bus
'
'    m_oBus.Init g_oActiveUser
'
'    m_oBus.Identify Trim(aszBusID(1))
'
'    frmGetFormula.m_szRouteID = m_oBus.Route
'
'    frmGetFormula.Show vbModal
'    If Not frmGetFormula.m_bOk Then Exit Sub
'    '按了OK
'    aszTemp = frmGetFormula.GetParam
'    abTicketType = frmGetFormula.GetSelectTicketType
'    nCount = ArrayLength(aszTemp)
'    With F1Book
'        For i = 1 To nCount
'            '得到各参数值
'            szKey1 = aszTemp(i, 1)
'            szKey2 = aszTemp(i, 2)
'            sgMul = aszTemp(i, 3)
'            sgAdd = aszTemp(i, 4)
'            szStationName = aszTemp(i, 5)
'            If aszTemp(i, 6) <> "" Then
'            sgMileage = aszTemp(i, 6)
'
''            得到各票价项的列的位置
'            lIndex1 = GetPriceItemEnablePosition(szKey1)
'            If szKey1 <> szKey2 Then
'                lIndex2 = GetPriceItemEnablePosition(szKey2)
'            Else
'                lIndex2 = lIndex1
'            End If
'
'
'
'            For j = cnPriceItemStartRow To .LastRow
'                '得到票种
'                nThisTicketType = GetTicketTypeID(j)
'
'                If abTicketType(nThisTicketType) And GetStationName(j) = szStationName And szKey2 <> "A000" Then
'                    If nThisTicketType = 1 Then
'                        .TextRC(j, lIndex1) = Round(.TextRC(j, lIndex2) * sgMul + sgAdd + 0.001, 2)
'                    Else
'                        .TextRC(j, lIndex1) = Round(.TextRC(j, lIndex2) * sgMul + sgAdd / 2 + 0.001, 2)
'                    End If
'                        .Row = j
'                        .Col = lIndex1
'                        SetChangeColor j, lIndex1
'                ElseIf abTicketType(nThisTicketType) And GetStationName(j) = szStationName And szKey2 = "A000" Then
'                    If nThisTicketType = 1 Then
'                        .TextRC(j, lIndex1) = Round(sgMileage * sgMul + 0.001, 2)
'                    Else
'                        .TextRC(j, lIndex1) = Round(sgMileage * sgMul / 2 + 0.001, 2)
'                    End If
'                        .Row = j
'                        .Col = lIndex1
'                        SetChangeColor j, lIndex1
'                End If
'                If abTicketType(nThisTicketType) And szStationName = "" Then
'                    If nThisTicketType = 1 Then
'                        .TextRC(j, lIndex1) = Round(.TextRC(j, lIndex2) * sgMul + sgAdd + 0.001, 2)
'                    Else
'                        .TextRC(j, lIndex1) = Round(.TextRC(j, lIndex2) * sgMul + sgAdd / 2 + 0.001, 2)
'                    End If
'                        .Row = j
'                        .Col = lIndex1
'                        SetChangeColor j, lIndex1
'                End If
'
''                If bModifyAll Or .IsSelectedCell(lIndex1, j - 1) Then
'                '如果修改所有的或选择了所有的CELL
'
''                    HalfTicketItemParam = GetHalfTicketItemParam(nThisTicketType, m_tHalfItemParam)
''                    If nThisTicketType = TP_FullPrice Then
'                        '如果票种为全票
''                        If szKey1 <> szKey2 Then
''                            .DoGetCellData lIndex1, j - 1, sgValue
''                        End If
''                        .DoGetCellData lIndex2, j - 1, sgValue1
''                        .DoGetCellData cnTotalPrice, j - 1, dbTempTotal
''
''                        sgResult = Round(sgValue1 * sgMul + sgAdd + 0.001, m_nDetailLength)
''                        If szKey1 <> szKey2 Then
''                            .DoSetCellData cnTotalPrice, j - 1, dbTempTotal + sgResult - sgValue
''                        Else
''                            .DoSetCellData cnTotalPrice, j - 1, dbTempTotal + sgResult - sgValue1
''                        End If
''                        .DoSetCellData lIndex1, j - 1, sgResult
''                        .DoSetCellColor lIndex1, j - 1, RGB(0, 0, 0), m_lModifiedColor
''                    ElseIf bnModifyTicketType(nThisTicketType) = True Then
''                        '如果为特殊票
''                        If lIndex1 < nRoutePriceItem + cnPriceItemStartCol Then
''                           dbParam1 = HalfTicketItemParam(lIndex1 - cnPriceItemStartCol + 1).sgParam1
''                           dbParam2 = HalfTicketItemParam(lIndex1 - cnPriceItemStartCol + 1).sgParam2
''                        Else
''                           dbParam1 = HalfTicketItemParam(lIndex1 - cnPriceItemStartCol + 13 - nRoutePriceItem).sgParam1
''                           dbParam2 = HalfTicketItemParam(lIndex1 - cnPriceItemStartCol + 13 - nRoutePriceItem).sgParam2
''                        End If
''                        dbHalf = Round(dbParam1 * sgResult + dbParam2 + 0.001, m_nDetailLength)
''                        .DoSetCellData lIndex1, j - 1, dbHalf
''                    End If
''                End If
'            Next
'        Next
'        m_oMantissa.SetTailCarry cnPriceItemStartRow, .MaxRow, , True
'        SetSaveEnabled True
'    End With
'End Sub
'
Public Sub BatchModify()
    '批量修改
    Dim aszTemp() As String
    Dim szKey1 As String
    Dim szKey2 As String
    Dim sgMul As Single
    Dim sgAdd As Single
    Dim sgMileage As Single
    Dim i As Integer
    Dim j As Long
    Dim nCount As Integer
    Dim abTicketType() As Boolean
    
    
    Dim lIndex1 As Long
    Dim lIndex2 As Long
    Dim nThisTicketType  As Integer
    Dim t As Integer
    Dim k As Long
    Dim nOtherTicketTypeComputeRatio As Double '不是全票的其它费的计算费率
    Dim nHalfItemCount As Integer
    
    '嘉兴增加
    Dim aszBusID() As String
    aszBusID = frmShowBus.GetBusID
    Dim szStationName As String

    '得到站点
    Dim moBusProject As New BusProject
    Dim aszBus() As String
    moBusProject.Init g_oActiveUser
    
    moBusProject.Identify
    
    nHalfItemCount = ArrayLength(m_atHalfItemParam)
    aszBus = moBusProject.GetAllBus(Trim(aszBusID(1)), , , True)
    
    Dim m_oBus As New Bus
    
    m_oBus.Init g_oActiveUser
    
    m_oBus.Identify Trim(aszBusID(1))
    
    frmGetFormula.m_szRouteID = m_oBus.Route
    
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
            szStationName = aszTemp(i, 5)
            If aszTemp(i, 6) <> "" Then
            sgMileage = aszTemp(i, 6)
            End If
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
                
                '查找此票种的参数设置方法 by zyw
                For t = 1 To nHalfItemCount
                    If Val(m_atHalfItemParam(t).szTicketType) = nThisTicketType And Val(m_atHalfItemParam(t).szTicketItem) = GetPriceItem(lIndex1) Then
                        Exit For
                    End If
                Next t
                If t <= nHalfItemCount Then nOtherTicketTypeComputeRatio = m_atHalfItemParam(t).sgParam1
                
                If abTicketType(nThisTicketType) And GetStationName(j) = szStationName And szKey2 <> "A000" Then
                    If nThisTicketType = 1 Then
                        .TextRC(j, lIndex1) = Round(.TextRC(j, lIndex2) * sgMul + sgAdd + 0.001, 2)
                        k = j
                    Else
                        .TextRC(j, lIndex1) = Round(.TextRC(k, lIndex2) * sgMul * nOtherTicketTypeComputeRatio + 0.001, 2)
                    End If
                        .Row = j
                        .Col = lIndex1
                        SetChangeColor j, lIndex1
                ElseIf abTicketType(nThisTicketType) And GetStationName(j) = szStationName And szKey2 = "A000" Then
                    If nThisTicketType = 1 Then
                        .TextRC(j, lIndex1) = Round(.TextRC(j, 7) * sgMul + 0.001, 2)
                        k = j
                    Else
                        .TextRC(j, lIndex1) = Round(.TextRC(k, 7) * sgMul * nOtherTicketTypeComputeRatio + 0.001, 2)
                    End If
                        .Row = j
                        .Col = lIndex1
                        SetChangeColor j, lIndex1
                End If
                If abTicketType(nThisTicketType) And szStationName = "" And szKey2 <> "A000" Then
                    If nThisTicketType = 1 Then
                        .TextRC(j, lIndex1) = Round(.TextRC(j, lIndex2) * sgMul + sgAdd + 0.001, 2)
                        k = j
                    Else
                        .TextRC(j, lIndex1) = Round(.TextRC(k, lIndex2) * sgMul * nOtherTicketTypeComputeRatio + 0.001, 2)
                    End If
                        .Row = j
                        .Col = lIndex1
                        SetChangeColor j, lIndex1
                End If
                
                If abTicketType(nThisTicketType) And szStationName = "" And szKey2 = "A000" Then
                    If nThisTicketType = 1 Then
                        .TextRC(j, lIndex1) = Round(.TextRC(j, 7) * sgMul + sgAdd + 0.001, 2)
                        k = j
                    Else
                        .TextRC(j, lIndex1) = Round(.TextRC(k, 7) * sgMul * nOtherTicketTypeComputeRatio + 0.001, 2)
                    End If
                        .Row = j
                        .Col = lIndex1
                        SetChangeColor j, lIndex1
                End If
                
            Next
        Next
        m_oMantissa.SetTailCarry cnPriceItemStartRow, .MaxRow, , True
        SetSaveEnabled True
    End With
End Sub

Private Function GetStationName(plRow As Long)
    '得到该行的票种
    m_rsResultPrice.Move plRow - cnPriceItemStartRow, adBookmarkFirst
    GetStationName = FormatDbValue(m_rsResultPrice!station_name)
End Function
