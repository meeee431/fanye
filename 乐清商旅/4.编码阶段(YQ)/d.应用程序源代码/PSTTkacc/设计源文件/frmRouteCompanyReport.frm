VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Begin VB.Form frmRouteCompanyReport 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��·��˾Ӫ�ձ���"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5220
   Icon            =   "frmRouteCompanyReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   5220
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   -180
      ScaleHeight     =   705
      ScaleWidth      =   7665
      TabIndex        =   4
      Top             =   0
      Width           =   7665
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ѡ���ѯ����:"
         Height          =   180
         Left            =   270
         TabIndex        =   5
         Top             =   240
         Width           =   1350
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   -240
      TabIndex        =   3
      Top             =   690
      Width           =   7725
   End
   Begin VB.Frame Frame1 
      Caption         =   "����˵��"
      Height          =   555
      Left            =   1020
      TabIndex        =   1
      Top             =   6210
      Width           =   6975
      Begin VB.Label Label3 
         Caption         =   "��Ʊ��ָ��ʱ��Σ�ͳ��Ʊ��������������ͳ����ƱԱ����Ʊ�����"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   6435
      End
   End
   Begin VB.ComboBox cboSellStation 
      Height          =   300
      Left            =   1890
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1920
      Width           =   2235
   End
   Begin RTComctl3.CoolButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   3300
      TabIndex        =   6
      Top             =   3300
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
      MICON           =   "frmRouteCompanyReport.frx":000C
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
      Height          =   315
      Left            =   1890
      TabIndex        =   7
      Top             =   1425
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   556
      _Version        =   393216
      Format          =   19791872
      CurrentDate     =   36572
   End
   Begin MSComCtl2.DTPicker dtpBeginDate 
      Height          =   315
      Left            =   1890
      TabIndex        =   8
      Top             =   930
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   556
      _Version        =   393216
      Format          =   19791872
      CurrentDate     =   36572
   End
   Begin RTComctl3.CoolButton cmdOk 
      Default         =   -1  'True
      Height          =   315
      Left            =   1890
      TabIndex        =   9
      Top             =   3300
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
      MICON           =   "frmRouteCompanyReport.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin FText.asFlatTextBox txtRouteID 
      Height          =   315
      Left            =   1860
      TabIndex        =   11
      Top             =   2400
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      OfficeXPColors  =   -1  'True
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   " RTStation"
      Enabled         =   0   'False
      Height          =   2760
      Left            =   -180
      TabIndex        =   10
      Top             =   3000
      Width           =   8745
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��������(&E):"
      Height          =   180
      Left            =   810
      TabIndex        =   15
      Top             =   1470
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ʼ����(&B):"
      Height          =   180
      Left            =   810
      TabIndex        =   14
      Top             =   990
      Width           =   1080
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ϳ�վ(&T):"
      Height          =   180
      Left            =   810
      TabIndex        =   13
      Top             =   1980
      Width           =   900
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��·����"
      Height          =   180
      Left            =   810
      TabIndex        =   12
      Top             =   2460
      Width           =   720
   End
End
Attribute VB_Name = "frmRouteCompanyReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IConditionForm
Const cszFileName = "��·��˾Ӫ�ձ���.xls"


Public m_bOk As Boolean
Private m_rsData As Recordset
Private m_vaCustomData As Variant

Dim m_aszTemp() As String
Dim oDss As New TicketBusDim
Dim rsTemp As Recordset
Dim m_szCode As String
Private m_szCompanyID As String
Private m_szStationID As String
Private m_szCompanyName As String
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdok_Click()
    On Error GoTo Error_Handle
    '���ɼ�¼��
    Dim rsData As New Recordset
    Dim i As Integer
    Set rsTemp = oDss.GetRouteCompanyByBusDate(dtpBeginDate.Value, dtpEndDate.Value, , IIf((cboSellStation.Text = ""), "", ResolveDisplay(cboSellStation)), , txtRouteID.Text)
    MakeRecordSet rsTemp
    
    '
    ReDim m_vaCustomData(1 To 3, 1 To 2)
    m_vaCustomData(1, 1) = "�ϳ�վ"
    If cboSellStation.Text = "" Then
        m_vaCustomData(1, 2) = ""
    Else
        m_vaCustomData(1, 2) = ResolveDisplayEx(cboSellStation)
    End If
    m_vaCustomData(2, 1) = "ͳ�ƿ�ʼ����"
    m_vaCustomData(2, 2) = Format(dtpBeginDate.Value, "YYYY��MM��DD��")
    m_vaCustomData(3, 1) = "ͳ�ƽ�������"
    m_vaCustomData(3, 2) = Format(dtpEndDate.Value, "YYYY��MM��DD��")
    
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

Private Sub Form_Load()
  AlignFormPos Me
    m_szCode = ""
    m_bOk = False
    FillSellStation cboSellStation
       
    oDss.Init m_oActiveUser
'    dtpBeginDate.Value = DateAdd("d", -1, m_oParam.NowDate)
'    dtpEndDate.Value = DateAdd("d", -1, m_oParam.NowDate)
    '����Ϊ�ϸ��µ�һ�ŵ�31
    Dim dyNow As Date
    dyNow = m_oParam.NowDate
    dtpBeginDate.Value = Format(DateAdd("m", -1, dyNow), "yyyy-mm-01")
    dtpEndDate.Value = DateAdd("d", -1, Format(dyNow, "yyyy-mm-01"))
'    dtpBeginDate.Value = dyNow
'    dtpEndDate.Value = dyNow
End Sub


Private Sub Form_Unload(Cancel As Integer)
SaveFormPos Me
End Sub

Private Property Get IConditionForm_CustomData() As Variant
    IConditionForm_CustomData = m_vaCustomData
End Property

Private Property Get IConditionForm_FileName() As String
    IConditionForm_FileName = cszFileName
End Property

Private Property Get IConditionForm_RecordsetData() As Recordset
    Set IConditionForm_RecordsetData = m_rsData
End Property
'Private Sub txtCompanyID_ButtonClick()
'   Dim oShell As New CommDialog
'    Dim szaTemp() As String
'     oShell.Init m_oActiveUser
'    szaTemp = oShell.SelectCompany(False)
'    Set oShell = Nothing
'    If ArrayLength(szaTemp) = 0 Then Exit Sub
'    txtCompanyID.Text = Trim(szaTemp(1, 1)) & "[" & Trim(szaTemp(1, 2)) & "]"
'    m_szCompanyID = Trim(szaTemp(1, 1))
'    m_szCompanyName = Trim(szaTemp(1, 2))
'    Set oShell = Nothing
'
'Exit Sub
'End Sub





Private Sub MakeRecordSet(prsData As Recordset)
    '�ֹ����ɼ�¼��
    Dim rsTemp As New Recordset
    Dim i As Integer
    Dim j As Integer
    
    '�ݷ�����


    Dim alNumber(TP_TicketTypeCount) As Long '����Ʊ�ֵ�����
    Dim adbAmount(TP_TicketTypeCount) As Double  '����Ʊ�ֵĽ��
    Dim szTransportCompanyName As String
    Dim szTransportCompanyID As String
    Dim lReturnNumber As Long
    Dim dbReturnAmount As Double
    Dim dbReturnCharge As Double
    Dim lChangeNumber As Long
    Dim dbChangeAmount As Double
    Dim dbChangeCharge As Double
    Dim lCancelNumber As Long '��Ʊ����
    Dim dbCancelAmount As Double '��Ʊ�ܶ�
    Dim dbTotalPrice As Double '�ܶ�
    Dim lTotalNumber As Long '������
    Dim dbTotalTicketPrice As Double '������Ʊ�����ѵ��ܽ��
    Dim szRouteName As String

    
    With rsTemp.Fields
      
        .Append "transport_company_name", adVarChar, 10
        For i = 1 To TP_TicketTypeCount
            .Append "number_ticket_type" & i, adInteger
            .Append "amount_ticket_type" & i, adCurrency
        Next i
        .Append "return_number", adBigInt
        .Append "return_amount", adCurrency
        .Append "return_charge", adCurrency
        .Append "change_number", adBigInt
        .Append "change_amount", adCurrency
        .Append "change_charge", adCurrency
        .Append "cancel_number", adBigInt
        .Append "cancel_amount", adCurrency
        .Append "total_number", adBigInt
        .Append "total_price", adCurrency
        .Append "total_ticket_price", adCurrency
        .Append "route_name", adVarChar, 20
 
    End With
    rsTemp.Open
    If prsData Is Nothing Then Exit Sub
    If prsData.RecordCount = 0 Then Exit Sub
    prsData.MoveFirst
    

    szTransportCompanyName = FormatDbValue(prsData!transport_company_short_name)
    szRouteName = FormatDbValue(prsData!route_name)
    
    For i = 1 To prsData.RecordCount
        If szTransportCompanyName <> FormatDbValue(prsData!transport_company_short_name) Then
        '�����¼��
            
            rsTemp.AddNew
        
            rsTemp!transport_company_name = szTransportCompanyName
            rsTemp!route_name = szRouteName
            For j = 1 To TP_TicketTypeCount
                rsTemp.Fields("number_ticket_type" & j) = alNumber(j)
                rsTemp.Fields("amount_ticket_type" & j) = adbAmount(j)
            Next j
            rsTemp!return_number = lReturnNumber
            rsTemp!return_amount = dbReturnAmount
            rsTemp!return_charge = dbReturnCharge
            rsTemp!change_number = lChangeNumber
            rsTemp!change_amount = dbChangeAmount
            rsTemp!change_charge = dbChangeCharge
            rsTemp!cancel_number = lCancelNumber
            rsTemp!cancel_amount = dbCancelAmount
            rsTemp!total_number = lTotalNumber
            rsTemp!total_price = dbTotalPrice
            rsTemp!total_ticket_price = dbTotalTicketPrice

            rsTemp.Update
            
            '���ԭֵ
            For j = 1 To TP_TicketTypeCount
                alNumber(j) = 0
                adbAmount(j) = 0
            Next j
            lReturnNumber = 0
            dbReturnAmount = 0
            dbReturnCharge = 0
            lChangeNumber = 0
            dbChangeAmount = 0
            dbChangeCharge = 0
            lCancelNumber = 0
            dbCancelAmount = 0
            lTotalNumber = 0
            dbTotalPrice = 0
            dbTotalTicketPrice = 0

            szTransportCompanyName = FormatDbValue(prsData!transport_company_short_name)
            
            
            szRouteName = FormatDbValue(prsData!route_name)
            alNumber(prsData!ticket_type) = alNumber(prsData!ticket_type) + FormatDbValue(prsData!passenger_number2)
            adbAmount(prsData!ticket_type) = adbAmount(prsData!ticket_type) + FormatDbValue(prsData!ticket_price2)
            lReturnNumber = lReturnNumber + FormatDbValue(prsData!ticket_return_number)
            dbReturnAmount = dbReturnAmount + FormatDbValue(prsData!ticket_return_amount)
            dbReturnCharge = dbReturnCharge + FormatDbValue(prsData!ticket_return_charge)
            lChangeNumber = lChangeNumber + FormatDbValue(prsData!ticket_change_number)
            dbChangeAmount = dbChangeAmount + FormatDbValue(prsData!ticket_change_charge)
            lCancelNumber = lCancelNumber + FormatDbValue(prsData!ticket_cancel_number)
            dbCancelAmount = dbCancelAmount + FormatDbValue(prsData!ticket_cancel_amount)
            lTotalNumber = lTotalNumber + FormatDbValue(prsData!passenger_number)
            dbTotalPrice = dbTotalPrice + FormatDbValue(prsData!ticket_price)
            dbTotalTicketPrice = dbTotalTicketPrice + FormatDbValue(prsData!total_ticket_price)
        Else
            alNumber(prsData!ticket_type) = alNumber(prsData!ticket_type) + FormatDbValue(prsData!passenger_number2)
            adbAmount(prsData!ticket_type) = adbAmount(prsData!ticket_type) + FormatDbValue(prsData!ticket_price2)
            lReturnNumber = lReturnNumber + FormatDbValue(prsData!ticket_return_number)
            dbReturnAmount = dbReturnAmount + FormatDbValue(prsData!ticket_return_amount)
            dbReturnCharge = dbReturnCharge + FormatDbValue(prsData!ticket_return_charge)
            lChangeNumber = lChangeNumber + FormatDbValue(prsData!ticket_change_number)
            dbChangeAmount = dbChangeAmount + FormatDbValue(prsData!ticket_change_charge)
            lCancelNumber = lCancelNumber + FormatDbValue(prsData!ticket_cancel_number)
            dbCancelAmount = dbCancelAmount + FormatDbValue(prsData!ticket_cancel_amount)
            lTotalNumber = lTotalNumber + FormatDbValue(prsData!passenger_number)
            dbTotalPrice = dbTotalPrice + FormatDbValue(prsData!ticket_price)
            dbTotalTicketPrice = dbTotalTicketPrice + FormatDbValue(prsData!total_ticket_price)
        End If
                    
        prsData.MoveNext
    Next i

    rsTemp.AddNew
    rsTemp!transport_company_name = szTransportCompanyName
    rsTemp!route_name = szRouteName
    For j = 1 To TP_TicketTypeCount
        rsTemp.Fields("number_ticket_type" & j) = alNumber(j)
        rsTemp.Fields("amount_ticket_type" & j) = adbAmount(j)
    Next j
    rsTemp!return_number = lReturnNumber
    rsTemp!return_amount = dbReturnAmount
    rsTemp!return_charge = dbReturnCharge
    rsTemp!change_number = lChangeNumber
    rsTemp!change_amount = dbChangeAmount
    rsTemp!change_charge = dbChangeCharge
    rsTemp!cancel_number = lCancelNumber
    rsTemp!cancel_amount = dbCancelAmount
    rsTemp!total_number = lTotalNumber
    rsTemp!total_price = dbTotalPrice
    rsTemp!total_ticket_price = dbTotalTicketPrice
    rsTemp.Update
    Set m_rsData = rsTemp
    
End Sub




Private Sub txtRouteID_ButtonClick()
    Dim aszRoute() As String
    aszRoute = m_oShell.SelectRoute(True)
    txtRouteID.Text = TeamToString(aszRoute, 2)

    m_szCode = TeamToString(aszRoute, 1)

End Sub


