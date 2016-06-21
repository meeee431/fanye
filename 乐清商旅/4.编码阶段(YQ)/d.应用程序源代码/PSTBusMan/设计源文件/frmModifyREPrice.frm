VERSION 5.00
Object = "{D11DF87C-EFEC-4838-B7E2-15462BF136FB}#2.1#0"; "RTReportLF.ocx"
Begin VB.Form frmModifyREPrice 
   Caption         =   "�򿪻���Ʊ��"
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
'* Engineer:�·�
'* Date Generated:2002/09/03
'* Last Revision Date:2002/09/11
'* Brief Description:����Ʊ�۱���
'* Relational Document:
'**********************************************************

Option Explicit
Const cszTemplateFile = "��������Ʊ��ģ��.xls"

Const cnPriceItemStartCol = 10
Const cnPriceItemStartRow = 2
Const cnTotalCol = 9 '�ܼ�������

Public m_eFormStatus As EFormStatus

Private m_oREScheme As New STReSch.REScheme
Private m_oRoutePriceTable As New RoutePriceTable
Private WithEvents F1Book As TTF160Ctl.F1Book
Attribute F1Book.VB_VarHelpID = -1
Private m_rsAllTicketItem As Recordset '�������ʹ�õ�Ʊ����
Private m_rsResultPrice As Recordset '������еĴ򿪵�Ʊ��
Private m_lRange As Long 'Ϊ��д������ʱ�õ�
Private m_oMantissa As New clMantissa 'β���������
Private m_atHalfItemParam() As THalfTicketItemParam '��Ʊ���Ż�ƱƱ����������
Private m_bChanged As Boolean '��־�Ƿ�ı�

Private m_bCalHalfPrice As Boolean '��־�Ƿ���Ҫ������

Private m_aszBusID() As String '���ѡ��ĳ���
Private m_dyBusDate As Date 'ѡ��ĳ�������


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
    Dim oPriceMan As New STPrice.TicketPriceMan
    Dim oHalfTicket As New STPrice.HalfTicketPrice
    On Error GoTo ErrorHandle
    Timer1.Enabled = False
'    Me.Hide
    '�õ����е�Ʊ����
    oPriceMan.Init g_oActiveUser
    Set m_rsAllTicketItem = oPriceMan.GetAllTicketItemRS(TP_PriceItemUse)
    m_oRoutePriceTable.Init g_oActiveUser
    m_oRoutePriceTable.Identify g_szExePriceTable
    '�õ�����Ʊ�ֵİ�Ʊ�������
    oHalfTicket.Init g_oActiveUser
    m_atHalfItemParam = oHalfTicket.GetItemParam(0, g_szExePriceTable, TP_PriceItemUse)
    m_oREScheme.Init g_oActiveUser
    If m_eFormStatus = EFS_Modify Then
        '��
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
            '�Ƿ���ȷ��
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

Private Sub OpenBusPrice()
    '�򿪳���Ʊ��
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
    Dim oRebus As New REBus
    Dim tPriceInfo As TRETicketPrice
    On Error GoTo ErrorHandle
    oRebus.Init g_oActiveUser
    
    With F1Book
        WriteProcessBar True, 1, .LastRow, "���ڱ��泵��Ʊ�۱�"
        m_rsResultPrice.MoveFirst
        For i = cnPriceItemStartRow To .LastRow
            '�õ��޸�״̬
            bModify = GetModifyStatus(i)
            If bModify Then
                '���Ϊ���޸Ļ���Ϊ����״̬
                '����
                oRebus.Identify m_rsResultPrice!bus_id, m_dyBusDate
                '�ϳ�վ
                tPriceInfo.szSellStationID = m_rsResultPrice!sell_station_id
                'վ��
                tPriceInfo.szStationID = m_rsResultPrice!station_id
                '��λ����
                tPriceInfo.szSeatType = m_rsResultPrice!seat_type_id
                'Ʊ��
                tPriceInfo.nTicketType = m_rsResultPrice!ticket_type
                '�ܼ�
                tPriceInfo.sgTotal = .NumberRC(i, cnTotalCol)
                '�����˼�
                tPriceInfo.sgBase = .NumberRC(i, cnPriceItemStartCol)
                '��Ʊ����
                For k = cnPriceItemStartCol + 1 To .MaxCol
                    szPriceItem = GetPriceItem(k)
                    tPriceInfo.asgPrice(CInt(szPriceItem)) = .NumberRC(i, k)
                Next k
                oRebus.ModifyBusTicket tPriceInfo
                '�����޸�״̬
                MarkCellRowModifyStatus i
            End If
            WriteProcessBar , i, .LastRow, "���ڱ��泵��Ʊ�۱�"
            m_rsResultPrice.MoveNext
        Next
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
    Dim i As Integer
    Dim oCellFormat As F1CellFormat
    Dim lCol As Long
    Dim lRow As Long
    Dim lColor As OLE_COLOR
    If pbModify Then
        lColor = vbYellow '��ɫ��ԭ���� ��ɫ����ʱ�ῴ����
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
    m_abChanged(plRow) = IIf(lColor = vbYellow, True, False)  '��־�����ѱ��޸�
End Sub

Private Sub InitMantissa()
''    ��ʼ�����������
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
'    Dim lRowTemp As Long
'    Dim lColTemp As Long
'
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
                '���ô˹����е�ע�͵Ĵ��� , �����ڿؼ�����BUG, һ����EndEdit�¼��иı�.row��.col��, �����¼��ͻ���Ч
                '�����˹����е�ע�͵Ĵ�����Ϊ���ֲ���.textRcʱ��ѭ������EndEdit�¼�,�����ô˹���ʱ�����õ���.row .
'                .Text = .TextRC(plRow, plCol) * m_atHalfItemParam(j).sgParam1 + m_atHalfItemParam(j).sgParam2
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
    GetPriceItem = m_rsAllTicketItem!price_item
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
    With F1Book
    
    For i = cnTotalCol To .MaxCol
        SetChangeColor plRow, i, False
    Next i
    End With
End Sub



Public Sub BatchModify()
    '�����޸�
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
    
    '��������
    Dim aszBusID() As String
    aszBusID = m_aszBusID
    Dim szStationName As String

    '�õ�վ��
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
                If abTicketType(nThisTicketType) Then
                    .TextRC(j, lIndex1) = Round(.TextRC(j, lIndex2) * sgMul + sgAdd + 0.001, 2)
                    .Row = j
                    .Col = lIndex1
                    SetChangeColor j, lIndex1
                End If
'                If bModifyAll Or .IsSelectedCell(lIndex1, j - 1) Then
                '����޸����еĻ�ѡ�������е�CELL

'                    HalfTicketItemParam = GetHalfTicketItemParam(nThisTicketType, m_tHalfItemParam)
'                    If nThisTicketType = TP_FullPrice Then
                        '���Ʊ��ΪȫƱ
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
'                        '���Ϊ����Ʊ
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


