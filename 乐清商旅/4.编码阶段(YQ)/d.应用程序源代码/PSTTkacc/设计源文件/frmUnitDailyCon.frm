VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmUnitDailyCon 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��վ��ƱӪ���ձ�"
   ClientHeight    =   3495
   ClientLeft      =   2925
   ClientTop       =   3390
   ClientWidth     =   5160
   HelpContextID   =   6001201
   Icon            =   "frmUnitDailyCon.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox txtSellStationID 
      Height          =   300
      Left            =   2295
      TabIndex        =   13
      Top             =   2430
      Width           =   1995
   End
   Begin RTComctl3.CoolButton CoolButton1 
      Height          =   315
      Left            =   1110
      TabIndex        =   12
      Top             =   3090
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
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
      MICON           =   "frmUnitDailyCon.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ComboBox cboSellStation 
      Height          =   300
      Left            =   2295
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   1980
      Width           =   1995
   End
   Begin RTComctl3.CoolButton cmdOk 
      Default         =   -1  'True
      Height          =   315
      Left            =   2475
      TabIndex        =   4
      Top             =   3090
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
      MICON           =   "frmUnitDailyCon.frx":0028
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
      Left            =   3810
      TabIndex        =   3
      Top             =   3090
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
      MICON           =   "frmUnitDailyCon.frx":0044
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   -60
      TabIndex        =   2
      Top             =   690
      Width           =   6885
   End
   Begin VB.PictureBox ptCaption 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   6615
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ѡ���ѯ����:"
         Height          =   180
         Left            =   270
         TabIndex        =   1
         Top             =   240
         Width           =   1350
      End
   End
   Begin MSComCtl2.DTPicker dtpBeginDate 
      Height          =   300
      Left            =   2310
      TabIndex        =   5
      Top             =   1080
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   529
      _Version        =   393216
      Format          =   62652416
      CurrentDate     =   36572
   End
   Begin MSComCtl2.DTPicker dtpEndDate 
      Height          =   300
      Left            =   2310
      TabIndex        =   6
      Top             =   1537
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   529
      _Version        =   393216
      Format          =   62652416
      CurrentDate     =   36572
   End
   Begin VB.Frame fraCaption 
      BackColor       =   &H00E0E0E0&
      Caption         =   " "
      Enabled         =   0   'False
      Height          =   3120
      Left            =   -120
      TabIndex        =   7
      Top             =   2880
      Width           =   8745
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ϳ�վ(&T):"
      Height          =   180
      Left            =   1050
      TabIndex        =   11
      Top             =   2055
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ʼ����(&B):"
      Height          =   180
      Left            =   1050
      TabIndex        =   9
      Top             =   1140
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��������(&E):"
      Height          =   180
      Left            =   1050
      TabIndex        =   8
      Top             =   1597
      Width           =   1080
   End
End
Attribute VB_Name = "frmUnitDailyCon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Implements IConditionForm
Const cszFileName = "��վ��ƱӪ���ձ�ģ��.xls"

Private m_rsData As Recordset
Public m_bOk As Boolean
Private m_vaCustomData As Variant

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
    On Error GoTo Error_Handle
    '���ɼ�¼��
    Dim rsTemp As Recordset
    Dim oDss As New TicketUnitDim
'    Dim rsData As New Recordset
'    Dim i As Integer
'
'    Dim j As Integer
    
    oDss.Init m_oActiveUser
    Set rsTemp = oDss.StationDateStat(dtpBeginDate.Value, dtpEndDate.Value, IIf(txtSellStationID.Text <> "", txtSellStationID.Text, ResolveDisplay(cboSellStation)))
'    With rsData.Fields
'        .Append "bus_date", rsTemp("bus_date").Type, rsTemp("bus_date").DefinedSize
'        For i = 1 To TP_TicketTypeCount
'            .Append "price_ticket_type_" & i, adCurrency
'            .Append "passenger_number_ticket_type_" & i, adInteger
'            .Append "base_price_ticket_type_" & i, adCurrency
'            For j = 1 To 15
'                .Append "price_item_" & j & "_ticket_type_" & i, adCurrency
'            Next j
'        Next i
'        .Append "total_price", adCurrency
'        .Append "total_passenger_number", adInteger
'        .Append "total_base_price", adCurrency
'        For j = 1 To 15
'            .Append "total_price_item_" & j, adCurrency
'        Next j
'
'    End With
'
'    rsData.Open
'    If rsTemp.RecordCount > 0 Then
'        rsTemp.MoveFirst
'        Dim dtLastDate As Date
'        Dim szFieldPrefix As String
'
'        Do While Not rsTemp.EOF
'            If dtLastDate <> rsTemp!bus_date Or rsTemp!ticket_type = TP_FullPrice Then
'                If rsData.RecordCount > 0 Then
'                    rsData.Update
'                End If
'                rsData.AddNew
'                dtLastDate = rsTemp!bus_date
'                rsData!bus_date = dtLastDate
'            End If
'            rsData("price_ticket_type_" & rsTemp!ticket_type) = rsTemp!ticket_price1
'            rsData("passenger_number_ticket_type_" & rsTemp!ticket_type) = rsTemp!passenger_number1
'            rsData("base_price_ticket_type_" & rsTemp!ticket_type) = rsTemp!base_price1
'
'            rsData!total_price = rsData!total_price + rsTemp!ticket_price1
'            rsData!total_passenger_number = rsData!total_passenger_number + rsTemp!passenger_number1
'            rsData!total_base_price = rsData!total_base_price + rsTemp!base_price1
'
'            On Error Resume Next
'            For i = 1 To 15
'                rsData("price_item_" & i & "_ticket_type_" & rsTemp!ticket_type) = rsTemp("price_item_" & i & "_1")
'                rsData("total_price_item_" & i) = rsData("total_price_item_" & i) + rsTemp("price_item_" & i & "_1")
'            Next
'            On Error GoTo Error_Handle
'
'            rsTemp.MoveNext
'        Loop
'        If rsData.RecordCount > 0 Then rsData.Update
'    End If
    Set m_rsData = rsTemp
    
    ReDim m_vaCustomData(1 To 4, 1 To 2)
    m_vaCustomData(1, 1) = "ͳ�ƿ�ʼ����"
    m_vaCustomData(1, 2) = Format(dtpBeginDate.Value, "YYYY��MM��DD��")
    
    m_vaCustomData(2, 1) = "ͳ�ƽ�������"
    m_vaCustomData(2, 2) = Format(dtpEndDate.Value, "YYYY��MM��DD��")
    
    Dim szSellStation As String
    ResolveDisplay cboSellStation, szSellStation
    m_vaCustomData(3, 1) = "�ϳ�վ"
    m_vaCustomData(3, 2) = szSellStation
    
    m_vaCustomData(4, 1) = "�Ʊ���"
    m_vaCustomData(4, 2) = m_oActiveUser.UserID
    m_bOk = True
    
    Unload Me
    Exit Sub
Error_Handle:
    ShowErrorMsg
End Sub

Private Sub CoolButton1_Click()
DisplayHelp Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If

End Sub

Private Sub Form_Load()
    m_bOk = False
    
    FillSellStation cboSellStation
    dtpBeginDate.Value = DateAdd("d", -1, m_oParam.NowDate)
    dtpEndDate.Value = DateAdd("d", -1, m_oParam.NowDate)
    
    If Trim(m_oActiveUser.SellStationID) <> "" Then
        txtSellStationID.Enabled = False
    End If

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

'Private Sub FillSellStation()
'    '�ж��û������ĸ��ϳ�վ,���Ϊ�������һ������,��������е��ϳ�վ
'
'    '����ֻ����û��������ϳ�վ
'End Sub



