VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Begin VB.Form frmStationConYear 
   Caption         =   "站点年报统计"
   ClientHeight    =   1665
   ClientLeft      =   2355
   ClientTop       =   1920
   ClientWidth     =   5865
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   5865
   Begin VB.Frame Frame1 
      Caption         =   "报表说明"
      Height          =   555
      Left            =   60
      TabIndex        =   2
      Top             =   990
      Width           =   5745
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "按指定参运公司，统计售出票的人数、金额和票价项各分项的统计值"
         Height          =   180
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   5400
      End
   End
   Begin RTComctl3.CoolButton cmdCancel 
      Cancel          =   -1  'True
      TX         =   "取消(&C)"
      Height          =   315
      Left            =   4440
      TabIndex        =   1
      Top             =   540
      Width           =   1215
   End
   Begin RTComctl3.CoolButton cmdOk 
      TX         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   315
      Left            =   4440
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dtpEndDate 
      Height          =   315
      Left            =   1170
      TabIndex        =   4
      Top             =   540
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   556
      _Version        =   393216
      Format          =   24510464
      CurrentDate     =   36572
   End
   Begin MSComCtl2.DTPicker dtpBeginDate 
      Height          =   315
      Left            =   1170
      TabIndex        =   5
      Top             =   120
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   556
      _Version        =   393216
      Format          =   24510464
      CurrentDate     =   36572
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "结束日期(&E)"
      Height          =   180
      Left            =   150
      TabIndex        =   7
      Top             =   600
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "开始日期(&B)"
      Height          =   180
      Left            =   150
      TabIndex        =   6
      Top             =   180
      Width           =   990
   End
End
Attribute VB_Name = "frmStationConYear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Implements IConditionForm
Const cszFileName = "站点售票营收年报模板.cll"

Private m_rsData As Recordset
Public m_bOk As Boolean
Private m_vaCustomData As Variant


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    On Error GoTo Error_Handle
    '生成记录集
    Dim rsTemp As Recordset
    Dim oCalculator As New SellerFinance
    Dim rsData As New Recordset
    Dim i As Integer
    
    oCalculator.Init m_oActiveUser
    Set rsTemp = oCalculator.StationCount(dtpBeginDate.Value, dtpEndDate.Value)
'    With rsData.Fields
'        .Append "bus_date", rsTemp("bus_date").Type, rsTemp("bus_date").DefinedSize
'
'        .Append "full_ticket_price", adCurrency
'        .Append "full_passenger_number", adInteger
'        .Append "full_base_price", adCurrency
'        For i = 1 To 14
'            .Append "full_price_item_" & i, adCurrency
'        Next
'
'
'        .Append "half_ticket_price", adCurrency
'        .Append "half_passenger_number", adInteger
'        .Append "half_base_price", adCurrency
'        For i = 1 To 14
'            .Append "half_price_item_" & i, adCurrency
'        Next
'
'        .Append "free_ticket_price", adCurrency
'        .Append "free_passenger_number", adInteger
'        .Append "free_base_price", adCurrency
'        For i = 1 To 14
'            .Append "free_price_item_" & i, adCurrency
'        Next
'
'        .Append "total_ticket_price", adCurrency
'        .Append "total_passenger_number", adInteger
'        .Append "total_base_price", adCurrency
'        For i = 1 To 14
'            .Append "total_price_item_" & i, adCurrency
'        Next
'
'    End With
'    rsData.Open
'    If rsTemp.RecordCount > 0 Then
'        rsTemp.MoveFirst
'        Dim dtLastDate As Date
'        Dim szFieldPrefix As String
'
'        Do While Not rsTemp.EOF
'            If dtLastDate <> rsTemp!bus_date Or rsTemp!Ticket_Type = TP_FullPrice Then
'                If rsData.RecordCount > 0 Then
'                    rsData.Update
'                End If
'                rsData.AddNew
'                dtLastDate = rsTemp!bus_date
'                rsData!bus_date = dtLastDate
'            End If
'            Select Case rsTemp!Ticket_Type
'                Case TP_FullPrice
'                    szFieldPrefix = "full_"
'                Case TP_HalfPrice
'                    szFieldPrefix = "half_"
'                Case TP_FreeTicket
'                    szFieldPrefix = "free_"
'            End Select
'
'            rsData(szFieldPrefix & "ticket_price") = rsTemp!ticket_price1
'            rsData(szFieldPrefix & "passenger_number") = rsTemp!passenger_number1
'            rsData(szFieldPrefix & "base_price") = rsTemp!base_price1
'
'            rsData!total_ticket_price = rsData!total_ticket_price + rsTemp!ticket_price1
'            rsData!total_passenger_number = rsData!total_passenger_number + rsTemp!passenger_number1
'            rsData!total_base_price = rsData!total_base_price + rsTemp!base_price1
'
'            On Error Resume Next
'            For i = 1 To 15
'                rsData(szFieldPrefix & "price_item_" & i) = rsTemp("price_item_" & i & "_1")
'                rsData("total_price_item_" & i) = rsData("total_price_item_" & i) + rsTemp("price_item_" & i & "_1")
'            Next
'            On Error GoTo Error_Handle
'
'            rsTemp.MoveNext
'        Loop
'        If rsData.RecordCount > 0 Then rsData.Update
'    End If
    Set m_rsData = rsTemp
    
    ReDim m_vaCustomData(1 To 2, 1 To 2)
    m_vaCustomData(1, 1) = "统计开始日期"
    m_vaCustomData(1, 2) = Format(dtpBeginDate.Value, "YYYY年MM月DD日")
    
    m_vaCustomData(2, 1) = "统计结束日期"
    m_vaCustomData(2, 2) = Format(dtpEndDate.Value, "YYYY年MM月DD日")
    
    
    m_bOk = True
    Unload Me
    Exit Sub
Error_Handle:
    ShowErrMsg
End Sub


Private Sub Form_Load()
    m_bOk = False
    
    dtpBeginDate.Value = m_oParam.NowDate
    dtpEndDate.Value = m_oParam.NowDate
    
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


