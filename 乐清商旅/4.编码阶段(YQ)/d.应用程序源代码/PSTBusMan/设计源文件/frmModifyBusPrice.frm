VERSION 5.00
Object = "{D11DF87C-EFEC-4838-B7E2-15462BF136FB}#2.1#0"; "RTReportLF.ocx"
Begin VB.Form frmModifyBusPrice 
   Caption         =   "����Ʊ��"
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
'* Engineer:�·�
'* Date Generated:2002/09/03
'* Last Revision Date:2002/09/10
'* Brief Description:���θ���Ʊ�۴���
'* Relational Document:
'**********************************************************

Option Explicit
Const cszTemplateFile = "����Ʊ��ģ��.xls"
Const cnPriceItemStartCol = 11
Const cnPriceItemStartRow = 2
Const cnTotalCol = 10 '�ܼ�������

Public m_eFormStatus As EFormStatus


Private m_oRoutePriceTable As New RoutePriceTable
Private WithEvents F1Book As TTF160Ctl.F1Book
Attribute F1Book.VB_VarHelpID = -1
Private m_rsAllTicketItem As Recordset '�������ʹ�õ�Ʊ����
Private m_rsResultPrice As Recordset '������еĴ򿪵�Ʊ��
Private m_lRange As Long 'Ϊ��д������ʱ�õ�
Private m_oMantissa As New clMantissa 'β���������
Private m_atHalfItemParam() As THalfTicketItemParam '��Ʊ���Ż�ƱƱ����������
Private m_szPriceTableID As String
Private m_bChanged As Boolean '��־�Ƿ�ı�
'Private m_bCalHalfPrice As Boolean '��־�Ƿ���Ҫ������


Private m_abChanged() As Boolean '���ÿһ���Ƿ��޸ĵı�־


Private Sub F1Book_DblClick(ByVal nRow As Long, ByVal nCol As Long)
    F1Book.StartEdit False, True, False
End Sub

Private Sub F1Book_EndEdit(EditString As String, Cancel As Integer)
    Dim szTicketTypeID As String
    Dim lRow As Long
    With F1Book
        If Not IsNumeric(EditString) Then
            '��������������������
            Cancel = True
            MsgBox "����������", vbInformation, Me.Caption
        Else
            '����޸���ֵ,���������޸ĵ���ɫ
            If .Text <> EditString Then
                SetSaveEnabled True  '���ñ������
                If .Col >= cnPriceItemStartCol Then
                    '����Ǹ�Ʊ����
                    .Text = EditString '�˴���ֵ��Ϊ���ʺ�SetTailCarry ,�����õ���.text
                    '����β������
                    m_oMantissa.SetTailCarry .Row, .Row, .Col, False
                    '�˴������õ����жϵ�ǰ��,���ԲŻ����ѭ������
                    lRow = .Row
                    szTicketTypeID = GetTicketTypeID(.Row)
                    If szTicketTypeID = TP_FullPrice Then
                        '�������ΪȫƱ��,���޸���Ӧ�İ�Ʊ���Ż�Ʊ��
                        ModifyHalfPrice .Row, .Col
                    End If
                    .Row = lRow
                    EditString = .Text '�ظ�,Ϊ���޸ĺ�Ĵ˹����˳�ʱ���Զ��ظ�.Text=EditString
                    
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
    '��ʼ������
'    RTReport.Enabled = False
End Sub



Private Sub Form_Resize()
    '���ÿؼ�λ��
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
        WriteProcessBar , lValue, m_lRange, "�������Ʊ��"
    End If
End Sub

Private Sub Timer1_Timer()
    '��ʼ��
    Dim oPriceMan As New stprice.TicketPriceMan
    Dim oHalfTicket As New stprice.HalfTicketPrice
    On Error GoTo ErrorHandle
    Timer1.Enabled = False
    '�õ����е�Ʊ����
    oPriceMan.Init g_oActiveUser
    Set m_rsAllTicketItem = oPriceMan.GetAllTicketItemRS(TP_PriceItemUse)
    '�õ�����Ʊ�ֵİ�Ʊ�������
    oHalfTicket.Init g_oActiveUser
    m_atHalfItemParam = oHalfTicket.GetItemParam(0, g_szExePriceTable, TP_PriceItemUse)
    m_oRoutePriceTable.Init g_oActiveUser
    Select Case m_eFormStatus
    Case EFS_AddNew
        '��������
        ShowAddDialog True
    Case EFS_Modify
        '��
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
        '��������Ʊ��
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
'            '�Ƿ���ȷ��
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
            '�Ƿ���ȷ��
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
    '����޸���,����ʾ����
    bCancel = False
    If m_bChanged Then
        If m_eFormStatus = EFS_AddNew Then
            szMsg = "������Ʊ��,�Ƿ�Ҫ���棿"
        Else
            szMsg = "Ʊ���Ѿ��޸�,�Ƿ�Ҫ���棿"
        End If
        nResult = MsgBox(szMsg, vbYesNoCancel, Me.Caption)
        If nResult = vbYes Then
            '����Ʊ��
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
    '��ʾ����������δ���浽���ݿ��е�����
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
    Dim bIsEmpty As Boolean '��־�Ƿ�ֻ���ɿ�Ʊ��
    
    On Error GoTo ErrorHandle
    
    SetBusy
    aszBusID = frmShowBus.GetBusID
    aszVehicleModel = frmShowBus.GetVehicleType
    aszSeatType = frmShowBus.GetSeatType
    m_szPriceTableID = frmShowBus.GetPriceTableID
    '�������鼯������,����������
    atBusVehicleSeat = ConvertTypeFromArray(aszBusID, aszVehicleModel, aszSeatType)
    '��������ϲ�Ϊһ������
    
    bIsEmpty = frmShowBus.IsEmpty
    m_oRoutePriceTable.Identify Trim(m_szPriceTableID)
    If bIsEmpty Then
        'ֻ���ɿ�Ʊ��,���Խ���ֱ������,����Ҫ��¼������˼��ʵȵ�
        Set m_rsResultPrice = m_oRoutePriceTable.MakeEmptyBusPriceRS(atBusVehicleSeat)
    Else
        '�������������������Ʊ��
        Set m_rsResultPrice = m_oRoutePriceTable.MakeSpecifyBusPriceRS(atBusVehicleSeat)
    End If
    aszTemp(1) = "Ʊ����"
    m_rsAllTicketItem.MoveFirst
    Set arsTemp(1) = m_rsAllTicketItem
    '���Ʊ�ۼ�¼��
    RTReport.LeftLabelVisual = True
    RTReport.TopLabelVisual = True
    RTReport.CustomStringCount = aszTemp
    RTReport.CustomString = arsTemp
    RTReport.TemplateFile = App.Path & "\" & cszTemplateFile
    RTReport.ShowReport m_rsResultPrice
    '���ù̶���,�пɼ���
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
    '�򿪳���Ʊ��
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
    aszTemp(1) = "Ʊ����"
    m_rsAllTicketItem.MoveFirst
    Set arsTemp(1) = m_rsAllTicketItem
    '���Ʊ�ۼ�¼��
    RTReport.LeftLabelVisual = True
    RTReport.TopLabelVisual = True
    RTReport.CustomStringCount = aszTemp
    RTReport.CustomString = arsTemp
    RTReport.TemplateFile = App.Path & "\" & cszTemplateFile
    RTReport.ShowReport m_rsResultPrice
    '���ù̶���,�пɼ���
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
    '���泵��Ʊ��
    Dim tDetailInfo(1 To 1) As TBusPriceDetailInfo
    Dim i As Long, k As Long
    Dim bModify As Boolean '��־�����Ƿ��޸�
    Dim szPriceItem As String
    Dim atBusInfo() As TBusVehicleSeatType
    
    On Error GoTo ErrorHandle
    With F1Book
        WriteProcessBar True, 1, .LastRow, "���ڱ��泵��Ʊ�۱�"
        m_rsResultPrice.MoveFirst
        '�õ����еĶ���
        If m_eFormStatus = EFS_AddNew Then
            '���Ϊ����
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
        '**********���Ϊ����״̬,����Ҫ��ɾ��ԭ�ȵ�����,������ܻ����վ�㲻ƥ��
        
        For i = cnPriceItemStartRow To .LastRow
            '�õ��޸�״̬
            If m_eFormStatus = EFS_AddNew Then
                bModify = True
            Else
                bModify = GetModifyStatus(i)
            End If
            If bModify Then
                '���Ϊ���޸Ļ���Ϊ����״̬
                '����
                tDetailInfo(1).szbusID = m_rsResultPrice!bus_id
                '����
                tDetailInfo(1).sgMileage = m_rsResultPrice!Mileage
                '����
                tDetailInfo(1).szVehicleModel = m_rsResultPrice!vehicle_type_code
                '��λ����
                tDetailInfo(1).szSeatTypeID = m_rsResultPrice!seat_type_id
                'վ��
                tDetailInfo(1).szStationID = m_rsResultPrice!station_id
                '����վ
                tDetailInfo(1).szSellStationID = m_rsResultPrice!sell_station_id
                'Ʊ��
                tDetailInfo(1).nTicketType = m_rsResultPrice!ticket_type
                '�ܼ�
                tDetailInfo(1).sgTotalPrice = .NumberRC(i, cnTotalCol)
                '�����˼�
                tDetailInfo(1).sgBaseCarriage = .NumberRC(i, cnPriceItemStartCol)
'                'վ�㳵�����
                tDetailInfo(1).nSerialNo = m_rsResultPrice!station_serial_no
                '��Ʊ����
                '��������������Ĭ��Ϊ0,����ֻ����ʾ��Ʊ����
                For k = cnPriceItemStartCol + 1 To .MaxCol
                    szPriceItem = GetPriceItem(k)
                    tDetailInfo(1).asgItem(CInt(szPriceItem)) = .NumberRC(i, k)
                Next
                m_oRoutePriceTable.ModifySpecifyBusPrice tDetailInfo
                MarkCellRowModifyStatus i
                
            End If
            
            WriteProcessBar , i, .LastRow, "���ڱ��泵��Ʊ�۱�"
            m_rsResultPrice.MoveNext
        Next
'        SetProgressBarVisible False
        WriteProcessBar False
'        .DoRedrawAll
    End With
    WriteProcessBar False
    '���ñ��治����
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
'�����ļ�
Public Sub ExportFile()
    Dim szFileName As String
    szFileName = RTReport.OpenDialog(EDialogType.EXPORT_FILE)
End Sub
'�����ļ�����
Public Sub ExportFileOpen()
    Dim szFileName As String
    szFileName = RTReport.OpenDialog(EDialogType.EXPORT_FILE)
    If szFileName <> "" Then
        OpenLinkedFile szFileName
    End If
End Sub

Private Sub SetMenuEnabled(pbEnabled As Boolean)
    '���ò˵��Ŀ�����
    With MDIScheme.abMenuTool
        .Bands("mnu_TicketPrice").Tools("mnu_TicketPriceMan_Open").Enabled = pbEnabled
        .Bands("mnu_TicketPrice").Tools("mnu_TicketPriceMan_Modify").Enabled = pbEnabled
    End With
End Sub

Private Sub SetChangeColor(plRow As Long, plCol As Long, Optional pbModify As Boolean = True)
    '����ĳһ�����ɫ,��ʾ�������޸�
'    Dim i As Integer
    Dim oCellFormat As F1CellFormat
'    Dim lCol As Long
'    Dim lRow As Long
    Dim lColor As OLE_COLOR
    If pbModify Then
        lColor = vbYellow  '��ɫ
    Else
        lColor = vbWhite
    End If
    With F1Book
'        '����ԭ������
'        lRow = .Row
'        lCol = .Col
'        '���ñ���
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
'        '�ظ�
'        .Col = lCol
'        .Row = lRow
    End With
    m_abChanged(plRow) = IIf(lColor = vbYellow, True, False)   '��־�����ѱ��޸�
    
End Sub

Private Sub InitMantissa()
'    ��ʼ�����������
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
    '�õ����е�Ʊ��
    m_rsResultPrice.Move plRow - cnPriceItemStartRow, adBookmarkFirst
    GetTicketTypeID = FormatDbValue(m_rsResultPrice!ticket_type)
End Function


Private Sub ModifyHalfPrice(ByVal plRow As Long, ByVal plCol As Long)
    '�������ΪȫƱ��,���޸���Ӧ�İ�Ʊ���Ż�Ʊ��
    Dim nHalfItemCount As Integer
    Dim lRow As Long
    Dim i As Integer, j As Integer
    Dim nTicketType As Integer
    Dim szPriceItem As String 'Ʊ����
    '
'    Dim lRowTemp As Long
'    Dim lColTemp As Long
    
    szPriceItem = GetPriceItem(plCol)
    nHalfItemCount = ArrayLength(m_atHalfItemParam)
    lRow = plRow
    '������һ��
    With F1Book
'        lRowTemp = .Row
'        lColTemp = .Col
        For i = 1 To g_nTicketCountValid - 1
            '������
            lRow = lRow + 1
            m_rsResultPrice.Move lRow - cnPriceItemStartRow, adBookmarkFirst
            nTicketType = FormatDbValue(m_rsResultPrice!ticket_type)
            If nTicketType = TP_FullPrice Then
                '���ΪȫƱ,�����
                Exit Sub
            End If
            '���Ҵ�Ʊ�ֵĲ������÷���
            For j = 1 To nHalfItemCount
                If Val(m_atHalfItemParam(j).szTicketType) = nTicketType And Val(m_atHalfItemParam(j).szTicketItem) = szPriceItem Then
                    Exit For
                End If
            Next j
            If j <= nHalfItemCount Then
                '�ҵ���Ӧ��Ʊ��
'                .Row = lRow
'                .Col = plCol
'                .Text = .TextRC(plRow, plCol) * m_atHalfItemParam(j).sgParam1 + m_atHalfItemParam(j).sgParam2
                '���ô˹����е�ע�͵Ĵ��� , �����ڿؼ�����BUG, һ����EndEdit�¼��иı�.row��.col��, �����¼��ͻ���Ч
                '�����˹����е�ע�͵Ĵ�����Ϊ���ֲ���.textRcʱ��ѭ������EndEdit�¼�,�����ô˹���ʱ�����õ���.row .
                .TextRC(lRow, plCol) = .TextRC(plRow, plCol) * m_atHalfItemParam(j).sgParam1 + m_atHalfItemParam(j).sgParam2
                
                '����β������
                m_oMantissa.SetTailCarry lRow, lRow, plCol
                SetChangeColor lRow, plCol
            End If
        Next i
'        .Row = lRowTemp
'        .Col = lColTemp
    End With
End Sub
    
Private Function GetPriceItem(plCol As Long) As String
    '�õ����е�Ʊ�������
    m_rsAllTicketItem.Move plCol - cnPriceItemStartCol, adBookmarkFirst
    GetPriceItem = FormatDbValue(m_rsAllTicketItem!price_item)
End Function

Private Function SetSaveEnabled(Optional pbEnabled As Boolean = True)
    '���ñ����Ƿ����
    MDIScheme.abMenuTool.Bands("mnu_TicketPrice").Tools("mnu_TicketPriceMan_Save").Enabled = pbEnabled
    m_bChanged = pbEnabled
End Function


Private Function GetModifyStatus(plRow As Long) As Boolean
'    Dim i As Integer
'    Dim oCellFormat As F1CellFormat
'    Dim lCol As Long
'    Dim lRow As Long
'    With F1Book
'        '����ԭ������
'        lRow = .Row
'        .Row = plRow
'        lCol = .Col
'        .Col = cnTotalCol
'        Set oCellFormat = F1Book.GetCellFormat
'        If oCellFormat.PatternFG = vbRed Then  '��ɫ
'            GetModifyStatus = True
'        Else
'            GetModifyStatus = False
'        End If
'    End With
    GetModifyStatus = m_abChanged(plRow)
End Function

Private Sub MarkCellRowModifyStatus(plRow As Long)
    '����ĳһ�е��޸�״̬
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
    '�õ�Ʊ�����ڿ���Ʊ�����е�λ��
    Dim i As Integer
    m_rsAllTicketItem.MoveFirst
    For i = 1 To m_rsAllTicketItem.RecordCount
        If FormatDbValue(m_rsAllTicketItem!price_item) = pszPriceItem Then
            Exit For
        End If
        m_rsAllTicketItem.MoveNext
    Next i
    If i <= m_rsAllTicketItem.RecordCount Then
        '����ҵ����˳�.
        GetPriceItemEnablePosition = m_rsAllTicketItem.Bookmark + cnPriceItemStartCol - 1
    End If
End Function


'Public Sub BatchModify()
'    '�����޸�
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
'    '��������
'    Dim aszBusID() As String
'    aszBusID = frmShowBus.GetBusID
'    Dim szStationName As String
'
'    '�õ�վ��
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
'    '����OK
'    aszTemp = frmGetFormula.GetParam
'    abTicketType = frmGetFormula.GetSelectTicketType
'    nCount = ArrayLength(aszTemp)
'    With F1Book
'        For i = 1 To nCount
'            '�õ�������ֵ
'            szKey1 = aszTemp(i, 1)
'            szKey2 = aszTemp(i, 2)
'            sgMul = aszTemp(i, 3)
'            sgAdd = aszTemp(i, 4)
'            szStationName = aszTemp(i, 5)
'            If aszTemp(i, 6) <> "" Then
'            sgMileage = aszTemp(i, 6)
'
''            �õ���Ʊ������е�λ��
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
'                '�õ�Ʊ��
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
'                '����޸����еĻ�ѡ�������е�CELL
'
''                    HalfTicketItemParam = GetHalfTicketItemParam(nThisTicketType, m_tHalfItemParam)
''                    If nThisTicketType = TP_FullPrice Then
'                        '���Ʊ��ΪȫƱ
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
''                        '���Ϊ����Ʊ
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
    '�����޸�
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
    Dim nOtherTicketTypeComputeRatio As Double '����ȫƱ�������ѵļ������
    Dim nHalfItemCount As Integer
    
    '��������
    Dim aszBusID() As String
    aszBusID = frmShowBus.GetBusID
    Dim szStationName As String

    '�õ�վ��
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
    '����OK
    aszTemp = frmGetFormula.GetParam
    abTicketType = frmGetFormula.GetSelectTicketType
    nCount = ArrayLength(aszTemp)
    With F1Book
        For i = 1 To nCount
            '�õ�������ֵ
            szKey1 = aszTemp(i, 1)
            szKey2 = aszTemp(i, 2)
            sgMul = aszTemp(i, 3)
            sgAdd = aszTemp(i, 4)
            szStationName = aszTemp(i, 5)
            If aszTemp(i, 6) <> "" Then
            sgMileage = aszTemp(i, 6)
            End If
'            �õ���Ʊ������е�λ��
            lIndex1 = GetPriceItemEnablePosition(szKey1)
            If szKey1 <> szKey2 Then
                lIndex2 = GetPriceItemEnablePosition(szKey2)
            Else
                lIndex2 = lIndex1
            End If
            
            
            
            For j = cnPriceItemStartRow To .LastRow
                '�õ�Ʊ��
                nThisTicketType = GetTicketTypeID(j)
                
                '���Ҵ�Ʊ�ֵĲ������÷��� by zyw
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
    '�õ����е�Ʊ��
    m_rsResultPrice.Move plRow - cnPriceItemStartRow, adBookmarkFirst
    GetStationName = FormatDbValue(m_rsResultPrice!station_name)
End Function
