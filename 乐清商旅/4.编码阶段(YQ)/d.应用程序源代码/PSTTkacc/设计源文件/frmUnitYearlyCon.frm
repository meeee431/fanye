VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmUnitYearlyCon 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "车站售票营收年报"
   ClientHeight    =   3195
   ClientLeft      =   3510
   ClientTop       =   3930
   ClientWidth     =   4980
   HelpContextID   =   6000401
   Icon            =   "frmUnitYearlyCon.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox cboSellStation 
      Height          =   300
      Left            =   2025
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   2025
      Width           =   1995
   End
   Begin RTComctl3.CoolButton cmdOk 
      Default         =   -1  'True
      Height          =   315
      Left            =   2070
      TabIndex        =   4
      Top             =   2790
      Width           =   1215
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "确定"
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
      MICON           =   "frmUnitYearlyCon.frx":000C
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
      Left            =   3420
      TabIndex        =   3
      Top             =   2790
      Width           =   1215
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "取消"
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
      MICON           =   "frmUnitYearlyCon.frx":0028
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
         Caption         =   "请选择查询条件:"
         Height          =   180
         Left            =   270
         TabIndex        =   1
         Top             =   240
         Width           =   1350
      End
   End
   Begin MSComCtl2.DTPicker dtpBeginDate 
      Height          =   300
      Left            =   2040
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
      Left            =   2040
      TabIndex        =   6
      Top             =   1545
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
      Top             =   2580
      Width           =   8745
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "上车站(&T):"
      Height          =   180
      Left            =   750
      TabIndex        =   11
      Top             =   2085
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "开始日期(&B):"
      Height          =   180
      Left            =   780
      TabIndex        =   9
      Top             =   1140
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "结束日期(&E):"
      Height          =   180
      Left            =   780
      TabIndex        =   8
      Top             =   1620
      Width           =   1080
   End
End
Attribute VB_Name = "frmUnitYearlyCon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements IConditionForm
Const cszFileName = "车站售票营收年报模板.xls"

Public m_bOk As Boolean
Private m_rsData As Recordset
Private m_vaCustomData As Variant

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
    On Error GoTo Error_Handle
    '生成记录集
    
    Dim rsTemp As Recordset
    Dim oDss As New TicketUnitDim
    Dim rsData As New Recordset
    Dim i As Integer
    Dim j As Integer
    
    oDss.Init m_oActiveUser
    Set rsTemp = oDss.StationMonthStat(dtpBeginDate.Value, DateAdd("d", 1, dtpEndDate.Value), ResolveDisplay(cboSellStation))
'    With rsData.Fields
'        .Append "bus_date", rsTemp("bus_date1").Type, rsTemp("bus_date1").DefinedSize
'
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
'    rsData.Open
'    If rsTemp.RecordCount > 0 Then
'        rsTemp.MoveFirst
'        Dim szLastMonth As String
'        Dim szFieldPrefix As String
'
'        Do While Not rsTemp.EOF
'            If szLastMonth <> RTrim(rsTemp!bus_date1) Or rsTemp!ticket_type = TP_FullPrice Then
'                If rsData.RecordCount > 0 Then
'                    rsData.Update
'                End If
'                rsData.AddNew
'                szLastMonth = RTrim(rsTemp!bus_date1)
'                rsData!bus_date = szLastMonth
'            End If
'
'            rsData("price_ticket_type_" & rsTemp!ticket_type) = rsTemp!ticket_price1
'            rsData("passenger_number_ticket_type_" & rsTemp!ticket_type) = rsTemp!passenger_number1
'            rsData("base_price_ticket_type_" & rsTemp!ticket_type) = rsTemp!base_price1
'
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
    ReDim m_vaCustomData(1 To 3, 1 To 2)
    m_vaCustomData(1, 1) = "统计开始月份"
    
    m_vaCustomData(1, 2) = Format(dtpBeginDate.Value, "YYYY年MM月DD日")
    m_vaCustomData(2, 1) = "统计结束月份"
    If dtpEndDate.Value > m_oParam.NowDate Then
    'm_vaCustomData(2, 2) = Format(dtpEndDate.Value, "YYYY年MM月DD日")
        m_vaCustomData(2, 2) = Format(m_oParam.NowDate, "YYYY年MM月DD日")
    Else
         m_vaCustomData(2, 2) = Format(dtpEndDate.Value, "YYYY年MM月DD日")
    End If
    m_vaCustomData(3, 1) = "制表人"
    m_vaCustomData(3, 2) = m_oActiveUser.UserID
    m_bOk = True
    Unload Me
    Exit Sub
Error_Handle:
    ShowErrorMsg
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If

End Sub

Private Sub Form_Load()
    m_bOk = False

    dtpBeginDate.Value = DateAdd("d", -1, m_oParam.NowDate)
    dtpEndDate.Value = DateAdd("d", -1, m_oParam.NowDate)
    FillSellStation cboSellStation
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
'    '判断用户属于哪个上车站,如果为空则填充一个空行,再填充所有的上车站
'
'    '否则只填充用户所属的上车站
'End Sub


