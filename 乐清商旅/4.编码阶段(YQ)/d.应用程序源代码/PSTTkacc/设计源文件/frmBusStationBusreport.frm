VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmBusStationBusreport 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����������������"
   ClientHeight    =   4095
   ClientLeft      =   2325
   ClientTop       =   2595
   ClientWidth     =   6585
   Icon            =   "frmBusStationBusreport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   6585
   Begin VB.ComboBox cboSellStation 
      Height          =   300
      Left            =   1410
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2190
      Width           =   1635
   End
   Begin VB.ComboBox cboBusSection 
      Height          =   300
      Left            =   4440
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2220
      Width           =   1905
   End
   Begin VB.OptionButton optCombine 
      BackColor       =   &H00E0E0E0&
      Caption         =   "�����ζλ���"
      Height          =   255
      Left            =   3150
      TabIndex        =   7
      Top             =   1140
      Width           =   1455
   End
   Begin VB.OptionButton optCompany 
      BackColor       =   &H00E0E0E0&
      Caption         =   "�����ʹ�˾����"
      Height          =   285
      Left            =   3150
      TabIndex        =   6
      Top             =   1680
      Width           =   1605
   End
   Begin VB.ComboBox cboExtraStatus 
      Height          =   300
      Left            =   1470
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2985
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   -30
      TabIndex        =   4
      Top             =   690
      Width           =   6885
   End
   Begin VB.TextBox txtLike 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1410
      TabIndex        =   3
      Top             =   2700
      Width           =   1905
   End
   Begin RTComctl3.CoolButton CoolButton1 
      Height          =   345
      Left            =   2190
      TabIndex        =   9
      Top             =   3600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "����"
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
      MICON           =   "frmBusStationBusreport.frx":000C
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
      Height          =   345
      Left            =   5070
      TabIndex        =   10
      Top             =   3600
      Width           =   1215
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "ȡ��"
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
      MICON           =   "frmBusStationBusreport.frx":0028
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
      Left            =   3630
      TabIndex        =   11
      Top             =   3600
      Width           =   1215
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "ȷ��"
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
      MICON           =   "frmBusStationBusreport.frx":0044
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker dtpEndDate 
      Height          =   300
      Left            =   1410
      TabIndex        =   1
      Top             =   1665
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   529
      _Version        =   393216
      Format          =   62652416
      CurrentDate     =   36572
   End
   Begin MSComCtl2.DTPicker dtpBeginDate 
      Height          =   300
      Left            =   1410
      TabIndex        =   0
      Top             =   1110
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   529
      _Version        =   393216
      Format          =   62652416
      CurrentDate     =   36572
   End
   Begin FText.asFlatTextBox txtTransportCompanyID 
      Height          =   300
      Left            =   4440
      TabIndex        =   12
      Top             =   2220
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   ""
      ButtonVisible   =   -1  'True
      OfficeXPColors  =   -1  'True
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   " "
      Enabled         =   0   'False
      Height          =   3120
      Left            =   -120
      TabIndex        =   15
      Top             =   3360
      Width           =   8745
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   -30
      ScaleHeight     =   735
      ScaleWidth      =   6615
      TabIndex        =   13
      Top             =   0
      Width           =   6615
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ѡ������:"
         Height          =   180
         Left            =   270
         TabIndex        =   14
         Top             =   270
         Width           =   990
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ϳ�վ(&T):"
      Height          =   180
      Left            =   240
      TabIndex        =   22
      Top             =   2250
      Width           =   900
   End
   Begin VB.Label lblCompany 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���˹�˾(T):"
      Height          =   180
      Left            =   3150
      TabIndex        =   21
      Top             =   2280
      Width           =   1080
   End
   Begin VB.Label lblCombine 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���ζ����(&R):"
      Height          =   180
      Left            =   3150
      TabIndex        =   20
      Top             =   2265
      Width           =   1260
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��Ʊ״̬(&S):"
      Height          =   180
      Left            =   390
      TabIndex        =   19
      Top             =   3075
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��������(&E):"
      Height          =   180
      Left            =   240
      TabIndex        =   18
      Top             =   1725
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ʼ����(&B):"
      Height          =   180
      Left            =   240
      TabIndex        =   17
      Top             =   1170
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ģ������(&A):"
      Height          =   180
      Left            =   240
      TabIndex        =   16
      Top             =   2760
      Width           =   1080
   End
End
Attribute VB_Name = "frmBusStationBusreport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IConditionForm
'Const cszFileName = "����������������.xls"


Public m_bOk As Boolean
Public m_bBySaleTime As Boolean
Public m_nIsCheck As EBusStationType  '�Ƿ��ǴӼ�Ʊ�в�ѯ

Private m_rsData As Recordset
Private m_vaCustomData As Variant

Private m_aszTemp() As String
Private oDss As New TicketBusDim

Private m_szCode As String



Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
    On Error GoTo Error_Handle
    '���ɼ�¼��
    Dim rsTemp As Recordset
    Dim nCount As Integer
    Dim rsData As New Recordset
    Dim i As Integer
    Dim oCTReport As New STChkTk.CTReport
    If m_nIsCheck = SNBusFromSale Then
        Set rsTemp = oDss.GetBusStationStatByBusDate(m_szCode, dtpBeginDate.Value, dtpEndDate.Value, ResolveDisplay(cboExtraStatus.Text), txtLike.Text, ResolveDisplay(cboSellStation))
    ElseIf m_nIsCheck = SNBusFromCheck Then
        oCTReport.Init m_oActiveUser
        Set rsTemp = oCTReport.GetBusStationStatByBusDate(m_szCode, dtpBeginDate.Value, dtpEndDate.Value, txtLike.Text, ResolveDisplay(cboSellStation))
    ElseIf m_nIsCheck = SNVehicleFromCheck Then
        oCTReport.Init m_oActiveUser
        Set rsTemp = oCTReport.GetVehicleStationStatByBusDate(m_szCode, dtpBeginDate.Value, dtpEndDate.Value, txtLike.Text, ResolveDisplay(cboSellStation))
    End If
    If m_nIsCheck = SNBusFromSale Or m_nIsCheck = SNBusFromCheck Then
        nCount = MakeRecordSet(rsTemp)
    Else
        nCount = MakeVehicleRecordSet(rsTemp)
    End If
    ReDim m_vaCustomData(1 To 8, 1 To 2)
    m_vaCustomData(1, 1) = "�ϳ�վ"
    m_vaCustomData(1, 2) = ResolveDisplayEx(cboSellStation)
    
    m_vaCustomData(2, 1) = "��������"
    m_vaCustomData(2, 2) = nCount
    m_vaCustomData(3, 1) = "ͳ�ƿ�ʼ����"
    m_vaCustomData(3, 2) = Format(dtpBeginDate.Value, "YYYY��MM��DD��")
    m_vaCustomData(4, 1) = "ͳ�ƽ�������"
    m_vaCustomData(4, 2) = Format(dtpEndDate.Value, "YYYY��MM��DD��")
    m_vaCustomData(5, 1) = "״̬"
    m_vaCustomData(5, 2) = cboExtraStatus.Text
    m_vaCustomData(6, 1) = "ͳ�Ʒ�ʽ"
    If m_nIsCheck = SNBusFromSale Then
        m_vaCustomData(6, 2) = "����Ʊͳ��"
    ElseIf m_nIsCheck = SNBusFromCheck Or m_nIsCheck = SNVehicleFromCheck Then
        m_vaCustomData(6, 2) = "����Ʊͳ��"
    End If
    m_vaCustomData(7, 1) = "���ʹ�˾"
    m_vaCustomData(7, 2) = IIf(txtTransportCompanyID.Text <> "", txtTransportCompanyID.Text, "���й�˾")
    
    m_vaCustomData(8, 1) = "�Ʊ���"
    m_vaCustomData(8, 2) = m_oActiveUser.UserID
    
    If rsTemp.RecordCount = 0 Then
        m_bOk = False
    Else
        m_bOk = True
    End If
    
    
    
    Unload Me
    Exit Sub
Error_Handle:
    ShowErrorMsg
End Sub


Private Sub CoolButton1_Click()
DisplayHelp Me
End Sub

Private Sub Form_Load()
    oDss.Init m_oActiveUser
    m_szCode = ""
    m_bOk = False
    FillCombine
    dtpBeginDate.Value = DateAdd("d", -1, m_oParam.NowDate)
    dtpEndDate.Value = DateAdd("d", -1, m_oParam.NowDate)
    '����Ϊ�ϸ��µ�һ�ŵ�31
    Dim dyNow As Date
    dyNow = m_oParam.NowDate
'    dtpBeginDate.Value = Format(DateAdd("m", -1, dyNow), "yyyy-mm-01")
'    dtpEndDate.Value = DateAdd("d", -1, Format(dyNow, "yyyy-mm-01"))
    cboExtraStatus.AddItem "1[��Ʊ]"
    cboExtraStatus.AddItem "2[��Ʊ]"
    cboExtraStatus.AddItem "3[��Ʊ+��Ʊ]"
    
    cboExtraStatus.ListIndex = 2
    
    optCompany.Value = True
    SetVisible False
    FillSellStation cboSellStation
    If m_nIsCheck = SNBusFromSale Then
        
        If m_bBySaleTime Then
            Me.Caption = "����Ӫ�ռ�[����Ʊʱ�����]"
            lblCaption = "��������Ʊ����ֹ����:"
        Else
            Me.Caption = "����Ӫ�ռ�[���������ڻ���]"
            lblCaption = "�����복�ε���ֹ����:"
        End If
    ElseIf m_nIsCheck = SNBusFromCheck Then
        
        Me.Caption = "��Ʊ����;��վ����"
        lblCaption = "��������ֹ����:"
    ElseIf m_nIsCheck = SNVehicleFromCheck Then
        
        Me.Caption = "��Ʊ����;��վ����"
        lblCaption = "��������ֹ����:"
        
    End If
End Sub

Private Property Get IConditionForm_CustomData() As Variant
    IConditionForm_CustomData = m_vaCustomData
End Property

Private Property Get IConditionForm_FileName() As String
    If m_nIsCheck = SNBusFromSale Or m_nIsCheck = SNBusFromCheck Then
        IConditionForm_FileName = "����������������.xls"
    ElseIf m_nIsCheck = SNVehicleFromCheck Then
        IConditionForm_FileName = "����;��վ����.xls"
    End If
End Property

Private Property Get IConditionForm_RecordsetData() As Recordset
    Set IConditionForm_RecordsetData = m_rsData
End Property



Private Sub optCombine_Click()
    SetVisible True
End Sub

Private Sub optCompany_Click()
    SetVisible False
    
End Sub

Private Sub txtTransportCompanyID_ButtonClick()
    Dim aszTransportCompanyID() As String
    aszTransportCompanyID = m_oShell.SelectCompany
    txtTransportCompanyID.Text = TeamToString(aszTransportCompanyID, 2)
    
    m_szCode = TeamToString(aszTransportCompanyID, 1)
    
End Sub

Private Sub FillCombine()
    '���Ψһ�ĳ������
    Dim aszTemp() As String
    Dim i As Integer
    Dim nCount As Integer
    Dim oCompanyDim As New TicketCompanyDim
    oCompanyDim.Init m_oActiveUser
    aszTemp = oCompanyDim.GetUniqueCombine
    nCount = ArrayLength(aszTemp)
    For i = 1 To nCount
        cboBusSection.AddItem aszTemp(i)
    Next i
    If cboBusSection.ListCount > 0 Then cboBusSection.ListIndex = 0
    Set oCompanyDim = Nothing
End Sub


Private Sub SetVisible(pbVisible As Boolean)
    lblCombine.Visible = pbVisible
    cboBusSection.Visible = pbVisible
    lblCompany.Visible = Not pbVisible
    txtTransportCompanyID.Visible = Not pbVisible
    
End Sub




'Private Sub FillSellStation()
'    '�ж��û������ĸ��ϳ�վ,���Ϊ�������һ������,��������е��ϳ�վ
'
'    '����ֻ����û��������ϳ�վ
'End Sub


Private Function MakeRecordSet(prsData As Recordset) As Integer
    '�ֹ����ɼ�¼��
    Dim rsTemp As New Recordset
    Dim i As Integer
    Dim j As Integer
    Dim aszTicketType() As String
    Dim k As Integer
    ReDim aszTicketType(1 To m_AllTicketType.RecordCount, 1 To 2)
    
    '�ݷ�����
    Dim szbusID As String
    Dim szStationId As String
    Dim alNumber(TP_TicketTypeCount) As Long '����Ʊ�ֵ�����
    Dim adbAmount(TP_TicketTypeCount) As Double  '����Ʊ�ֵĽ��
    Dim szRouteName As String
    Dim szEndStationName As String
    Dim szBusStartTime As String
    Dim szTransportCompanyName As String
'    Dim lReturnNumber As Long
'    Dim dbReturnAmount As Double
'    Dim dbReturnCharge As Double
'    Dim lChangeNumber As Long
'    Dim dbChangeAmount As Double
'    Dim dbChangeCharge As Double
'    Dim lCancelNumber As Long '��Ʊ����
'    Dim dbCancelAmount As Double '��Ʊ�ܶ�
    Dim dbTotalPrice As Double '�ܶ�
    Dim lTotalNumber As Long '������
    Dim dbTotalTicketPrice As Double '������Ʊ�����ѵ��ܽ��
    Dim szStation As String
    Dim snTicketType As Integer
    Dim nCount As Integer
    
    Dim szStationName As String ';��վ��������
    Dim lStationNumber As Long  'ȫƱ����
    Dim lStationMoney As Double


    With rsTemp.Fields
        .Append "bus_id", adChar, 5
        .Append "route_name", adChar, 16
        .Append "end_station_name", adVarChar, 10 '�յ�վ��
        .Append "bus_start_time", adChar, 5
        .Append "transport_company_name", adVarChar, 10
        For i = 1 To TP_TicketTypeCount
            .Append "number_ticket_type" & i, adInteger
            .Append "amount_ticket_type" & i, adCurrency
        Next i
      
        .Append "total_number", adBigInt
        .Append "total_ticket_price", adCurrency
'        .Append "total_ticket_price", adCurrency
        .Append "station", adVarChar, 255 ';��վ��������
        .Append "ticket_type", adBigInt 'Ʊ������
        
        
    End With
    rsTemp.Open
    If prsData Is Nothing Then Exit Function
    If prsData.RecordCount = 0 Then Exit Function
    prsData.MoveFirst
    szbusID = FormatDbValue(prsData!bus_id)
    szStationId = FormatDbValue(prsData!station_id)
    szStationName = FormatDbValue(prsData!station_name)
'    lStationNumber = FormatDbValue(prsData!passenger_number)
    snTicketType = FormatDbValue(prsData!ticket_type)
    szRouteName = FormatDbValue(prsData!route_name)
    szEndStationName = FormatDbValue(prsData!end_station_name)
    szBusStartTime = FormatDbValue(prsData!bus_start_time)
    szTransportCompanyName = FormatDbValue(prsData!transport_company_short_name)
    
   
    
    
    
    
     m_AllTicketType.MoveFirst
    For k = 1 To m_AllTicketType.RecordCount
         aszTicketType(k, 2) = FormatDbValue(m_AllTicketType!ticket_type_name)
         m_AllTicketType.MoveNext
    Next k
        
    For i = 1 To prsData.RecordCount
        If szbusID <> FormatDbValue(prsData!bus_id) Then
            '�����¼��
            
            rsTemp.AddNew
            rsTemp!bus_id = szbusID
            rsTemp!route_name = szRouteName
            rsTemp!end_station_name = szEndStationName
            rsTemp!bus_start_time = szBusStartTime
            rsTemp!transport_company_name = szTransportCompanyName
            For j = 1 To TP_TicketTypeCount
                rsTemp.Fields("number_ticket_type" & j) = alNumber(j)
                rsTemp.Fields("amount_ticket_type" & j) = adbAmount(j)
            Next j
            rsTemp!total_ticket_price = dbTotalTicketPrice
            rsTemp!total_number = lTotalNumber
            rsTemp!Station = szStation
            rsTemp.Update
            nCount = nCount + 1
    
'            ���ԭֵ
            For j = 1 To TP_TicketTypeCount
                alNumber(j) = 0
                adbAmount(j) = 0
            Next j
            lStationNumber = 0
            lTotalNumber = 0
            dbTotalPrice = 0
            dbTotalTicketPrice = 0
            szStation = ""
                
            '���ó��εĳ�ʼֵ
            
            szbusID = FormatDbValue(prsData!bus_id)
            szRouteName = FormatDbValue(prsData!route_name)
            szEndStationName = FormatDbValue(prsData!end_station_name)
            szBusStartTime = FormatDbValue(prsData!bus_start_time)
            szTransportCompanyName = FormatDbValue(prsData!transport_company_short_name)
            

            szStationName = FormatDbValue(prsData!station_name)
            szStationId = FormatDbValue(prsData!station_id)
            lStationNumber = FormatDbValue(prsData!passenger_number)
            lStationMoney = FormatDbValue(prsData!ticket_price)
           
           If lStationNumber <> 0 Then
                 szStation = szStation & szStationName & "(" & Mid(aszTicketType(prsData!ticket_type, 2), 1, 1) & ")" & "(" & lStationNumber & ")" & "(" & lStationMoney & ")"
           End If
        ElseIf szStationId <> FormatDbValue(prsData!station_id) Then
            '�����ͬ
            
            '��¼�¸�վ�������
    
            szStationId = FormatDbValue(prsData!station_id)
            lStationNumber = FormatDbValue(prsData!passenger_number)
            If lStationNumber <> 0 Then
                lStationMoney = FormatDbValue(prsData!ticket_price)
                szStationName = FormatDbValue(prsData!station_name)
                szStation = szStation & szStationName & "(" & Mid(aszTicketType(prsData!ticket_type, 2), 1, 1) & ")" & "(" & lStationNumber & ")" & "(" & lStationMoney & ")"
            End If
        Else
            If snTicketType <> prsData!ticket_type Then '���Ʊ�����Ͳ�ͬ��ֿ����
                    alNumber(prsData!ticket_type) = FormatDbValue(prsData!passenger_number)
                    adbAmount(prsData!ticket_type) = FormatDbValue(prsData!ticket_price)
                    
                    If Val(alNumber(prsData!ticket_type)) <> 0 Then
                        If prsData!ticket_type = 1 Then
                           szStation = szStation & szStationName & "(" & Mid(aszTicketType(prsData!ticket_type, 2), 1, 1) & ")" & "(" & alNumber(prsData!ticket_type) & ")" & "(" & adbAmount(prsData!ticket_type) & ")"
                        Else
                            szStation = szStation & "(" & Mid(aszTicketType(prsData!ticket_type, 2), 1, 1) & ")" & "(" & alNumber(prsData!ticket_type) & ")" & "(" & adbAmount(prsData!ticket_type) & ")"
                        End If
                    End If
                                            
            Else
    '                '��ͬ���ۼ�
                   lStationNumber = lStationNumber + FormatDbValue(prsData!passenger_number)
                   lStationMoney = lStationMoney + FormatDbValue(prsData!ticket_price)
                   If lStationNumber <> 0 Then
                        If prsData!ticket_type = 1 Then
                             szStation = szStation & szStationName & "(" & Mid(aszTicketType(prsData!ticket_type, 2), 1, 1) & ")" & "(" & lStationNumber & ")" & "(" & lStationMoney & ")"
                        Else
                             szStation = szStation & "(" & Mid(aszTicketType(prsData!ticket_type, 2), 1, 1) & ")" & "(" & lStationNumber & ")" & "(" & lStationMoney & ")"
                        End If
                    End If
            End If
                    
        End If
        dbTotalTicketPrice = dbTotalTicketPrice + FormatDbValue(prsData!ticket_price)
        lTotalNumber = lTotalNumber + FormatDbValue(prsData!passenger_number)
        prsData.MoveNext
    Next i

    rsTemp.AddNew
    rsTemp!bus_id = szbusID
    rsTemp!route_name = szRouteName
    rsTemp!end_station_name = szEndStationName
    rsTemp!bus_start_time = szBusStartTime
    rsTemp!transport_company_name = szTransportCompanyName
    For j = 1 To TP_TicketTypeCount
        rsTemp.Fields("number_ticket_type" & j) = alNumber(j)
        rsTemp.Fields("amount_ticket_type" & j) = adbAmount(j)
    Next j
    szStation = szStation
'    dbTotalTicketPrice = dbTotalTicketPrice + FormatDbValue(prsData!ticket_price)
     rsTemp!total_ticket_price = dbTotalTicketPrice
     rsTemp!total_number = lTotalNumber
    rsTemp!Station = szStation
    rsTemp.Update
    
    Set m_rsData = rsTemp
    nCount = nCount + 1
  MakeRecordSet = nCount
End Function

Private Function MakeVehicleRecordSet(prsData As Recordset) As Integer
    '�ֹ����ɼ�¼��
    Dim rsTemp As New Recordset
    Dim i As Integer
    Dim j As Integer
    Dim aszTicketType() As String
    Dim k As Integer
    ReDim aszTicketType(1 To m_AllTicketType.RecordCount, 1 To 2)
    
    '�ݷ�����
    Dim szVehicleID As String
    Dim szStationId As String
    Dim alNumber(TP_TicketTypeCount) As Long '����Ʊ�ֵ�����
    Dim adbAmount(TP_TicketTypeCount) As Double  '����Ʊ�ֵĽ��
    Dim szRouteID As String
    Dim szRouteName As String
    
    Dim szLicenseTagNo As String
    Dim szEndStationName As String
    Dim szBusStartTime As String
    Dim szTransportCompanyName As String
'    Dim lReturnNumber As Long
'    Dim dbReturnAmount As Double
'    Dim dbReturnCharge As Double
'    Dim lChangeNumber As Long
'    Dim dbChangeAmount As Double
'    Dim dbChangeCharge As Double
'    Dim lCancelNumber As Long '��Ʊ����
'    Dim dbCancelAmount As Double '��Ʊ�ܶ�
    Dim dbTotalPrice As Double '�ܶ�
    Dim lTotalNumber As Long '������
    Dim dbTotalTicketPrice As Double '������Ʊ�����ѵ��ܽ��
    Dim szStation As String
    Dim snTicketType As Integer
    Dim nCount As Integer
    
    Dim szStationName As String ';��վ��������
    Dim lStationNumber As Long  'ȫƱ����
    Dim lStationMoney As Double


    With rsTemp.Fields
        .Append "vehicle_id", adChar, 5
        .Append "route_id", adChar, 4
        .Append "route_name", adChar, 16
        .Append "license_tag_no", adChar, 10
        .Append "end_station_name", adVarChar, 10 '�յ�վ��
        .Append "bus_start_time", adChar, 5
        .Append "transport_company_name", adVarChar, 10
        For i = 1 To TP_TicketTypeCount
            .Append "number_ticket_type" & i, adInteger
            .Append "amount_ticket_type" & i, adCurrency
        Next i
      
        .Append "total_number", adBigInt
        .Append "total_price", adCurrency
        .Append "total_ticket_price", adCurrency
        .Append "station", adVarChar, 255 ';��վ��������
        .Append "ticket_type", adBigInt 'Ʊ������
        
        
    End With
    rsTemp.Open
    If prsData Is Nothing Then Exit Function
    If prsData.RecordCount = 0 Then Exit Function
    prsData.MoveFirst
    szVehicleID = FormatDbValue(prsData!vehicle_id)
    
    szStationId = FormatDbValue(prsData!station_id)
    szStationName = FormatDbValue(prsData!station_name)
    szRouteID = FormatDbValue(prsData!route_id)
    szRouteName = FormatDbValue(prsData!route_name)
    
    snTicketType = FormatDbValue(prsData!ticket_type)
    szLicenseTagNo = FormatDbValue(prsData!license_tag_no)
    szTransportCompanyName = FormatDbValue(prsData!transport_company_short_name)
    
   
    
    
    
    
     m_AllTicketType.MoveFirst
    For k = 1 To m_AllTicketType.RecordCount
         aszTicketType(k, 2) = FormatDbValue(m_AllTicketType!ticket_type_name)
         m_AllTicketType.MoveNext
    Next k
        
    For i = 1 To prsData.RecordCount
        If szVehicleID <> FormatDbValue(prsData!vehicle_id) Or szRouteID <> FormatDbValue(prsData!route_id) Then
            '�����¼��
            
            rsTemp.AddNew
            rsTemp!vehicle_id = szVehicleID
            rsTemp!license_tag_no = szLicenseTagNo
            rsTemp!route_id = szRouteID
            rsTemp!route_name = szRouteName
            rsTemp!end_station_name = szEndStationName
            rsTemp!bus_start_time = szBusStartTime
            rsTemp!transport_company_name = szTransportCompanyName
            For j = 1 To TP_TicketTypeCount
                rsTemp.Fields("number_ticket_type" & j) = alNumber(j)
                rsTemp.Fields("amount_ticket_type" & j) = adbAmount(j)
            Next j
            rsTemp!total_ticket_price = dbTotalTicketPrice
            rsTemp!total_number = lTotalNumber
            rsTemp!Station = szStation
            rsTemp.Update
            nCount = nCount + 1
    
'            ���ԭֵ
            For j = 1 To TP_TicketTypeCount
                alNumber(j) = 0
                adbAmount(j) = 0
            Next j
            lStationNumber = 0
            lTotalNumber = 0
            dbTotalPrice = 0
            dbTotalTicketPrice = 0
            szStation = ""
                
            '���ó��εĳ�ʼֵ
            
            szVehicleID = FormatDbValue(prsData!vehicle_id)
            szLicenseTagNo = FormatDbValue(prsData!license_tag_no)
            szRouteID = FormatDbValue(prsData!route_id)
            szRouteName = FormatDbValue(prsData!route_name)
            szTransportCompanyName = FormatDbValue(prsData!transport_company_short_name)
            

            szStationName = FormatDbValue(prsData!station_name)
            szStationId = FormatDbValue(prsData!station_id)
            lStationNumber = FormatDbValue(prsData!passenger_number)
            lStationMoney = FormatDbValue(prsData!ticket_price)
           
           If lStationNumber <> 0 Then
                 szStation = szStation & szStationName & "(" & Mid(aszTicketType(prsData!ticket_type, 2), 1, 1) & ")" & "(" & lStationNumber & ")" & "(" & lStationMoney & ")"
           End If
        ElseIf szStationId <> FormatDbValue(prsData!station_id) Then
            '�����ͬ
            
            '��¼�¸�վ�������
    
            szStationId = FormatDbValue(prsData!station_id)
            lStationNumber = FormatDbValue(prsData!passenger_number)
            If lStationNumber <> 0 Then
                lStationMoney = FormatDbValue(prsData!ticket_price)
                szStationName = FormatDbValue(prsData!station_name)
                szStation = szStation & szStationName & "(" & Mid(aszTicketType(prsData!ticket_type, 2), 1, 1) & ")" & "(" & lStationNumber & ")" & "(" & lStationMoney & ")"
            End If
        Else
            If snTicketType <> prsData!ticket_type Then '���Ʊ�����Ͳ�ͬ��ֿ����
                    alNumber(prsData!ticket_type) = FormatDbValue(prsData!passenger_number)
                    adbAmount(prsData!ticket_type) = FormatDbValue(prsData!ticket_price)
                    
                    If Val(alNumber(prsData!ticket_type)) <> 0 Then
                        If prsData!ticket_type = 1 Then
                           szStation = szStation & szStationName & "(" & Mid(aszTicketType(prsData!ticket_type, 2), 1, 1) & ")" & "(" & alNumber(prsData!ticket_type) & ")" & "(" & adbAmount(prsData!ticket_type) & ")"
                        Else
                            szStation = szStation & "(" & Mid(aszTicketType(prsData!ticket_type, 2), 1, 1) & ")" & "(" & alNumber(prsData!ticket_type) & ")" & "(" & adbAmount(prsData!ticket_type) & ")"
                        End If
                    End If
                                            
            Else
    '                '��ͬ���ۼ�
                   lStationNumber = lStationNumber + FormatDbValue(prsData!passenger_number)
                   lStationMoney = lStationMoney + FormatDbValue(prsData!ticket_price)
                   If lStationNumber <> 0 Then
                        If prsData!ticket_type = 1 Then
                             szStation = szStation & szStationName & "(" & Mid(aszTicketType(prsData!ticket_type, 2), 1, 1) & ")" & "(" & lStationNumber & ")" & "(" & lStationMoney & ")"
                        Else
                             szStation = szStation & "(" & Mid(aszTicketType(prsData!ticket_type, 2), 1, 1) & ")" & "(" & lStationNumber & ")" & "(" & lStationMoney & ")"
                        End If
                    End If
            End If
                    
        End If
        dbTotalTicketPrice = dbTotalTicketPrice + FormatDbValue(prsData!ticket_price)
        lTotalNumber = lTotalNumber + FormatDbValue(prsData!passenger_number)
        prsData.MoveNext
    Next i

    rsTemp.AddNew
    rsTemp!vehicle_id = szVehicleID
    rsTemp!license_tag_no = szLicenseTagNo
    rsTemp!route_id = szRouteID
    rsTemp!route_name = szRouteName
    rsTemp!end_station_name = szEndStationName
    rsTemp!bus_start_time = szBusStartTime
    rsTemp!transport_company_name = szTransportCompanyName
    For j = 1 To TP_TicketTypeCount
        rsTemp.Fields("number_ticket_type" & j) = alNumber(j)
        rsTemp.Fields("amount_ticket_type" & j) = adbAmount(j)
    Next j
    szStation = szStation
'    dbTotalTicketPrice = dbTotalTicketPrice + FormatDbValue(prsData!ticket_price)
     rsTemp!total_ticket_price = dbTotalTicketPrice
     rsTemp!total_number = lTotalNumber
    rsTemp!Station = szStation
    rsTemp.Update
    
    Set m_rsData = rsTemp
    nCount = nCount + 1
  MakeVehicleRecordSet = nCount
End Function

