VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmBusBalance 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����ƽ���"
   ClientHeight    =   3720
   ClientLeft      =   3930
   ClientTop       =   4440
   ClientWidth     =   5850
   Icon            =   "frmBusBalance.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   5850
   StartUpPosition =   1  '����������
   Begin VB.PictureBox ptTitle 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   795
      Left            =   -15
      ScaleHeight     =   795
      ScaleWidth      =   7185
      TabIndex        =   1
      Top             =   0
      Width           =   7185
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������ѯ����:"
         Height          =   180
         Left            =   180
         TabIndex        =   2
         Top             =   300
         Width           =   1350
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   -75
      TabIndex        =   0
      Top             =   780
      Width           =   7215
   End
   Begin MSComCtl2.DTPicker dtpEndDate 
      Height          =   315
      Left            =   2295
      TabIndex        =   3
      Top             =   1905
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   556
      _Version        =   393216
      Format          =   61800448
      CurrentDate     =   37725
   End
   Begin MSComCtl2.DTPicker dtpStartDate 
      Height          =   315
      Left            =   2295
      TabIndex        =   4
      Top             =   1365
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   556
      _Version        =   393216
      Format          =   61800448
      CurrentDate     =   37725
   End
   Begin RTComctl3.CoolButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   345
      Left            =   3975
      TabIndex        =   5
      Top             =   3180
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "ȡ��(&C)"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      MICON           =   "frmBusBalance.frx":000C
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
      Default         =   -1  'True
      Height          =   345
      Left            =   2490
      TabIndex        =   6
      Top             =   3180
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "ȷ��(&E)"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      MICON           =   "frmBusBalance.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   975
      Left            =   -60
      TabIndex        =   7
      Top             =   2895
      Width           =   6960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ʼ����:"
      Height          =   180
      Left            =   1335
      TabIndex        =   9
      Top             =   1425
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��������:"
      Height          =   180
      Left            =   1335
      TabIndex        =   8
      Top             =   1995
      Width           =   810
   End
End
Attribute VB_Name = "frmBusBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Implements IConditionForm
Const cszFileName = "����ƽ���.xls"

Public m_bOk As Boolean
Private m_rsData As Recordset
Private m_vaCustomData As Variant
Const cnNormalCheckStatus = 0

Private Sub cmdCancel_Click()
    Unload Me
    m_bOk = False
End Sub

Private Sub cmdok_Click()

    On Error GoTo Error_Handle
    '���ɼ�¼��
    Dim rsSale As Recordset '��Ʊ�Ӽ�¼��
    Dim rsCheck As Recordset '��Ʊ�Ӽ�¼��
    Dim rsSettle As Recordset '�����Ӽ�¼��
    
    Dim oDss As New STDss.TicketBusDim
    Dim oReport As New Report
    
    Dim rsTemp As Recordset
    Dim rsData As New Recordset
    Dim i As Integer
    Dim szSellStation As String

    
    oDss.Init g_oActiveUser
    oReport.Init g_oActiveUser
    
    Set rsSale = oDss.GetBusStatByBusDate("", dtpStartDate.Value, dtpEndDate.Value, TP_AllSold)
    Set rsCheck = oReport.GetCheckBusStatByBusDate("", dtpStartDate.Value, dtpEndDate.Value)
    
    Set rsSettle = oReport.BusSettleStat(dtpStartDate.Value, DateAdd("d", 1, dtpEndDate.Value), "", "", CS_QueryAll, CS_SettleSheetNotInvalid, False)
    
    Set rsTemp = MakeRecordset(rsSale, rsCheck, rsSettle)
    
    
    Set m_rsData = rsTemp
    
    ReDim m_vaCustomData(1 To 2, 1 To 2)
    
    m_vaCustomData(1, 1) = "��ʼ����"
    m_vaCustomData(1, 2) = Format(dtpStartDate.Value, "YYYY��MM��DD��")
    m_vaCustomData(2, 1) = "��������"
    m_vaCustomData(2, 2) = Format(dtpEndDate.Value, "YYYY��MM��DD��")
    
    
    m_bOk = True
    
    
    
    Set oDss = Nothing
    Set oReport = Nothing
    Set rsSale = Nothing
    Set rsCheck = Nothing
    Set rsSettle = Nothing
    Set rsTemp = Nothing
    
    
    Unload Me
    Exit Sub
Error_Handle:
    ShowErrorMsg
    
    Set oDss = Nothing
    Set oReport = Nothing
    Set rsSale = Nothing
    Set rsCheck = Nothing
    Set rsSettle = Nothing
    Set rsTemp = Nothing
End Sub

Private Sub Form_Load()
    AlignFormPos Me
    dtpStartDate.Value = GetFirstMonthDay(Date)
    dtpEndDate.Value = GetLastMonthDay(Date)
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    SaveFormPos Me
    Unload Me
End Sub

Private Function MakeRecordset(prsSale As Recordset, prsCheck As Recordset, prsSettle As Recordset) As Recordset
    '������Ʊ,��Ʊ,����ļ�¼��,�ϲ�Ϊһ����¼��
    '�ֹ����ɼ�¼��
    Dim rsTemp As New Recordset
    Dim i As Integer
    Dim j As Integer

    With rsTemp.Fields
        .Append "bus_id", adChar, 5 '���δ���
        .Append "transport_company_id", adChar, 12 '���˹�˾����
        .Append "transport_company_short_name", adVarChar, 10 '���˹�˾���
        .Append "route_id", adChar, 4 '��·����
        .Append "route_name", adChar, 16 '��·����
        .Append "bus_start_time", adChar, 5 '����ʱ��
        

        .Append "sale_quantity", adBigInt 'δ�ų��˵���Ʊ����
        .Append "sale_ticket_price", adCurrency 'δ�ų��˵���Ʊ���
        .Append "return_number", adBigInt '��Ʊ����
        .Append "return_charge", adCurrency '��Ʊ������
        .Append "sale_total_quantity", adBigInt '�ܵ���Ʊ����
        .Append "sale_total_ticket_price", adCurrency '�ܵ���Ʊ���

        .Append "check_quantity", adBigInt '������Ʊ��
        .Append "check_ticket_price", adCurrency '������Ʊ���
        .Append "check_change_number", adBigInt '�Ĳ�����
        .Append "check_change_price", adCurrency '�Ĳ����
        .Append "check_total_quantity", adBigInt '�ܵļ�Ʊ��
        .Append "check_total_ticket_price", adCurrency '�ܼ�Ʊ���

        .Append "settle_quantity", adBigInt '��������
        .Append "settle_price", adBigInt '������
        .Append "settle_station_price", adBigInt '�����վ
        .Append "settle_other_price", adBigInt '�������
'        .Append "settle_ticket_price", adCurrency '������
        For i = 1 To 20
            .Append "split_item_" & i, adCurrency '������
        Next i

        .Append "left_quantity", adBigInt 'ʣ������
        .Append "left_ticket_price", adCurrency 'ʣ����
    End With
'
'    rsTemp.Open
    Set prsSale = MergeSale(prsSale)
    Set prsCheck = MergeCheck(prsCheck)
    Set prsSettle = MergeSettle(prsSettle)
    If prsSale.RecordCount > 0 Then
        prsSale.MoveFirst
    End If
    If prsCheck.RecordCount > 0 Then
        prsCheck.MoveFirst
    End If
    '�ϲ���Ʊ����Ʊ
    Set prsSale = MergeMultiRecord(prsSale, prsCheck, "bus_id,transport_company_id")
    '���������кϲ�
    Set prsSale = MergeMultiRecord(prsSale, prsSettle, "bus_id,transport_company_id")
    
    '������ʣ�����������
    
    If prsSale Is Nothing Then Exit Function
    If prsSale.RecordCount = 0 Then Exit Function
    rsTemp.Open
    For i = 1 To prsSale.RecordCount
        rsTemp.AddNew
        rsTemp!bus_id = prsSale!bus_id
        rsTemp!route_id = prsSale!route_id
        rsTemp!route_name = prsSale!route_name
        rsTemp!bus_start_time = prsSale!bus_start_time
        
        
        rsTemp!transport_company_id = prsSale!transport_company_id
        rsTemp!transport_company_short_name = prsSale!transport_company_short_name
        
        rsTemp!sale_quantity = prsSale!sale_quantity
        rsTemp!sale_ticket_price = prsSale!sale_ticket_price
        rsTemp!return_number = prsSale!return_number
        rsTemp!return_charge = prsSale!return_charge
        rsTemp!sale_total_quantity = prsSale!sale_total_quantity
        rsTemp!sale_total_ticket_price = prsSale!sale_total_ticket_price
        
        
        rsTemp!check_quantity = prsSale!check_quantity
        rsTemp!check_ticket_price = prsSale!check_ticket_price
        rsTemp!check_change_number = prsSale!check_change_number
        rsTemp!check_change_price = prsSale!check_change_price
        rsTemp!check_total_quantity = prsSale!check_total_quantity
        rsTemp!check_total_ticket_price = prsSale!check_total_ticket_price
        
        
        rsTemp!settle_quantity = prsSale!settle_quantity
        rsTemp!settle_price = prsSale!settle_price
        rsTemp!settle_station_price = prsSale!settle_station_price
        rsTemp!settle_other_price = prsSale!settle_other_price
        
        
        For j = 1 To 20
            rsTemp.Fields("split_item_" & j).Value = FormatDbValue(prsSale.Fields("split_item_" & j))
        Next j
        
        rsTemp!left_quantity = prsSale!check_total_quantity - prsSale!settle_quantity
        rsTemp!left_ticket_price = prsSale!check_total_ticket_price - FormatDbValue(prsSale.Fields("split_item_1"))
        rsTemp.Update
        prsSale.MoveNext
    Next i
'    rsTemp.Close
    rsTemp.MoveFirst
    Set MakeRecordset = rsTemp
    
    
    
    
    
End Function


Private Function MergeSale(prsData As Recordset) As Recordset
    '����Ʊ�ļ�¼��,�ֶ���������
    
    Dim rsTemp As New Recordset
    Dim i As Integer
    Dim j As Integer

    With rsTemp.Fields
        .Append "bus_id", adChar, 5 '���δ���
        .Append "transport_company_id", adChar, 12 '���˹�˾����
        .Append "transport_company_short_name", adVarChar, 10 '���˹�˾���
        .Append "route_id", adChar, 4 '��·����
        .Append "route_name", adChar, 16 '��·����
        .Append "bus_start_time", adChar, 5 '����ʱ��
                
        .Append "sale_quantity", adBigInt 'δ�ų��˵���Ʊ����
        .Append "sale_ticket_price", adCurrency 'δ�ų��˵���Ʊ���
        .Append "return_number", adBigInt '��Ʊ����
        .Append "return_charge", adCurrency '��Ʊ������
        .Append "sale_total_quantity", adBigInt '�ܵ���Ʊ����
        .Append "sale_total_ticket_price", adCurrency '�ܵ���Ʊ���
    End With
    
    rsTemp.Open
    
    If prsData Is Nothing Then Set MergeSale = rsTemp:   Exit Function
    If prsData.RecordCount = 0 Then Set MergeSale = rsTemp: Exit Function
    '����Ʊ
    
    prsData.MoveFirst
    For i = 1 To prsData.RecordCount

        rsTemp.AddNew
        rsTemp!bus_id = FormatDbValue(prsData!bus_id)
        rsTemp!transport_company_id = FormatDbValue(prsData!transport_company_id)
        rsTemp!transport_company_short_name = FormatDbValue(prsData!transport_company_short_name)
        
        rsTemp!route_id = FormatDbValue(prsData!route_id)
        rsTemp!route_name = FormatDbValue(prsData!route_name)
        rsTemp!bus_start_time = Format(FormatDbValue(prsData!bus_start_time), "hh:mm")
        
        
        
        rsTemp!sale_quantity = FormatDbValue(prsData!passenger_number) + FormatDbValue(prsData!ticket_return_number)
        rsTemp!sale_ticket_price = FormatDbValue(prsData!ticket_price)
        rsTemp!return_number = FormatDbValue(prsData!ticket_return_number)
        rsTemp!return_charge = FormatDbValue(prsData!ticket_return_charge)
        rsTemp!sale_total_quantity = FormatDbValue(prsData!passenger_number)
        rsTemp!sale_total_ticket_price = FormatDbValue(prsData!total_ticket_price)
        rsTemp.Update
        
        prsData.MoveNext
        
        
    Next i
    Set MergeSale = rsTemp
    
End Function


Private Function MergeCheck(prsData As Recordset) As Recordset
    '����Ʊ�ļ�¼��,�ϲ���һ����һ����¼
    
    Dim rsTemp As New Recordset
    Dim i As Integer
    Dim j As Integer
    
    
    '�ݷ�����
    Dim szBusID As String
    Dim szRouteID As String
    Dim szRouteName As String
    Dim szBusStartTime As String
    Dim szTransportCompanyID As String
'    Dim szTransportCompanyName As String
    
    
    Dim lCheckQuantity As Long
    Dim dbCheckTicketPrice As Double
    Dim lCheckChangeNumber As Long
    Dim dbCheckChangePrice As Double
    Dim lCheckTotalQuantity As Long
    Dim dbCheckTotalTicketPrice As Double
    
    
    
    With rsTemp.Fields
        .Append "bus_id", adChar, 5 '���δ���
        .Append "transport_company_id", adChar, 12 '���˹�˾����
'        .Append "transport_company_short_name", adVarChar, 10 '���˹�˾���
        .Append "route_id", adChar, 4 '��·����
        .Append "route_name", adChar, 16 '��·����
        
        .Append "check_quantity", adBigInt '������Ʊ��
        .Append "check_ticket_price", adCurrency  '������Ʊ���
        .Append "check_change_number", adBigInt '�Ĳ�����
        .Append "check_change_price", adCurrency  '�Ĳ����
        .Append "check_total_quantity", adBigInt '�ܵļ�Ʊ��
        .Append "check_total_ticket_price", adCurrency '�ܼ�Ʊ���
    End With
    
    rsTemp.Open
    
    
    If prsData Is Nothing Then Set MergeCheck = rsTemp: Exit Function
    If prsData.RecordCount = 0 Then Set MergeCheck = rsTemp: Exit Function

    prsData.MoveFirst
    
    szBusID = FormatDbValue(prsData!bus_id)
    szRouteID = FormatDbValue(prsData!route_id)
    szRouteName = FormatDbValue(prsData!route_name)
    szTransportCompanyID = FormatDbValue(prsData!transport_company_id)
'    szTransportCompanyName = FormatDbValue(prsData!transport_company_short_name)
    
    For i = 1 To prsData.RecordCount
        If szBusID <> FormatDbValue(prsData!bus_id) Then
            '�����¼��
            rsTemp.AddNew
            rsTemp!bus_id = szBusID
            rsTemp!route_id = szRouteID
            rsTemp!route_name = szRouteName
            rsTemp!transport_company_id = szTransportCompanyID
'            rsTemp!transport_company_short_name = szTransportCompanyName
            
            rsTemp!check_quantity = lCheckQuantity
            rsTemp!check_ticket_price = dbCheckTicketPrice
            rsTemp!check_change_number = lCheckChangeNumber
            rsTemp!check_change_price = dbCheckChangePrice
            rsTemp!check_total_quantity = lCheckTotalQuantity
            rsTemp!check_total_ticket_price = dbCheckTotalTicketPrice
            
            rsTemp.Update
            
            lCheckQuantity = 0
            dbCheckTicketPrice = 0
            lCheckChangeNumber = 0
            dbCheckChangePrice = 0
            lCheckTotalQuantity = 0
            dbCheckTotalTicketPrice = 0
            
            '���ó��εĳ�ʼֵ
                    
            szBusID = FormatDbValue(prsData!bus_id)
            szRouteID = FormatDbValue(prsData!route_id)
            szRouteName = FormatDbValue(prsData!route_name)
            szTransportCompanyID = FormatDbValue(prsData!transport_company_id)
'            szTransportCompanyName = FormatDbValue(prsData!transport_company_short_name)
            If FormatDbValue(prsData!Status) = cnNormalCheckStatus Then
                lCheckQuantity = lCheckQuantity + FormatDbValue(prsData!Quantity)
                dbCheckTicketPrice = dbCheckTicketPrice + FormatDbValue(prsData!ticket_price)
                
            Else
                lCheckChangeNumber = lCheckChangeNumber + FormatDbValue(prsData!Quantity)
                dbCheckChangePrice = dbCheckChangePrice + FormatDbValue(prsData!ticket_price)
                
            End If
            lCheckTotalQuantity = lCheckTotalQuantity + FormatDbValue(prsData!Quantity)
            dbCheckTotalTicketPrice = dbCheckTotalTicketPrice + FormatDbValue(prsData!ticket_price)
            
        Else
            '��ͬ���ۼ�
            If FormatDbValue(prsData!Status) = cnNormalCheckStatus Then
                lCheckQuantity = lCheckQuantity + FormatDbValue(prsData!Quantity)
                dbCheckTicketPrice = dbCheckTicketPrice + FormatDbValue(prsData!ticket_price)
            Else
                lCheckChangeNumber = lCheckChangeNumber + FormatDbValue(prsData!Quantity)
                dbCheckChangePrice = dbCheckChangePrice + FormatDbValue(prsData!ticket_price)
            End If
            lCheckTotalQuantity = lCheckTotalQuantity + FormatDbValue(prsData!Quantity)
            dbCheckTotalTicketPrice = dbCheckTotalTicketPrice + FormatDbValue(prsData!ticket_price)
        End If
        prsData.MoveNext
    Next i


    rsTemp.AddNew
    rsTemp!bus_id = szBusID
    rsTemp!route_id = szRouteID
    rsTemp!route_name = szRouteName
    
    rsTemp!check_quantity = lCheckQuantity
    rsTemp!check_ticket_price = dbCheckTicketPrice
    rsTemp!check_change_number = lCheckChangeNumber
    rsTemp!check_change_price = dbCheckChangePrice
    rsTemp!check_total_quantity = lCheckTotalQuantity
    rsTemp!check_total_ticket_price = dbCheckTotalTicketPrice
    
    rsTemp!transport_company_id = szTransportCompanyID
'    rsTemp!transport_company_short_name = szTransportCompanyName
            
    rsTemp.Update
    
    Set MergeCheck = rsTemp
    
    
    
    
End Function

Private Function MergeSettle(prsData As Recordset) As Recordset
    '����Ʊ�ļ�¼��,�ֶ���������
    
    Dim rsTemp As New Recordset
    Dim i As Integer
    Dim j As Integer

    With rsTemp.Fields
        .Append "bus_id", adChar, 5 '���δ���
        
        .Append "transport_company_id", adChar, 12 '���˹�˾����
'        .Append "transport_company_short_name", adVarChar, 10 '���˹�˾���
        .Append "route_id", adChar, 4 '��·����
        .Append "route_name", adChar, 16 '��·����
        
        .Append "settle_quantity", adBigInt '��������
        .Append "settle_price", adCurrency '������
        .Append "settle_station_price", adCurrency '���㳵վ
        .Append "settle_other_price", adCurrency '���㳵��
'        .Append "settle_ticket_price", adCurrency '������
        For i = 1 To 20
            .Append "split_item_" & i, adCurrency '������
        Next i
        
    End With
    
    rsTemp.Open
    
    If prsData Is Nothing Then Set MergeSettle = rsTemp: Exit Function
    If prsData.RecordCount = 0 Then Set MergeSettle = rsTemp: Exit Function
    
    
    prsData.MoveFirst
    For i = 1 To prsData.RecordCount

        rsTemp.AddNew
        rsTemp!bus_id = FormatDbValue(prsData!bus_id)
        rsTemp!route_id = FormatDbValue(prsData!route_id)
        rsTemp!route_name = FormatDbValue(prsData!route_name)
        
        rsTemp!transport_company_id = FormatDbValue(prsData!transport_company_id)
'        rsTemp!transport_company_short_name = FormatDbValue(prsData!transport_company_short_name)
        
        rsTemp!settle_quantity = FormatDbValue(prsData!passenger_number)
        rsTemp!settle_price = FormatDbValue(prsData!settle_price)
        rsTemp!settle_station_price = FormatDbValue(prsData!settle_station_price)
        rsTemp!settle_other_price = FormatDbValue(prsData!settle_other_price)
        
        For j = 1 To 20
            rsTemp.Fields("split_item_" & j).Value = FormatDbValue(prsData.Fields("split_item_" & j))
        Next j
        rsTemp.Update
        
        prsData.MoveNext
        
        
    Next i
    Set MergeSettle = rsTemp
    
End Function


Private Property Get IConditionForm_CustomData() As Variant
    IConditionForm_CustomData = m_vaCustomData
End Property

Private Property Get IConditionForm_FileName() As String
    IConditionForm_FileName = cszFileName
End Property

Private Property Get IConditionForm_RecordsetData() As Recordset
    Set IConditionForm_RecordsetData = m_rsData
End Property



