VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmUnitManSimple 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "车站发出人数统计"
   ClientHeight    =   3615
   ClientLeft      =   4515
   ClientTop       =   1755
   ClientWidth     =   6180
   Icon            =   "frmUnitManSimple.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   6180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtSellStationID 
      Height          =   300
      Left            =   2250
      TabIndex        =   14
      Top             =   2520
      Width           =   1995
   End
   Begin RTComctl3.CoolButton cmdChart 
      Height          =   315
      Left            =   420
      TabIndex        =   13
      Top             =   3210
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "图表"
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
      MICON           =   "frmUnitManSimple.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton CoolButton1 
      Height          =   315
      Left            =   2160
      TabIndex        =   12
      Top             =   3210
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "帮助"
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
      MICON           =   "frmUnitManSimple.frx":0028
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
      Left            =   2250
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   2100
      Width           =   1995
   End
   Begin VB.PictureBox ptCaption 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   6615
      TabIndex        =   3
      Top             =   0
      Width           =   6615
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请选择查询条件:"
         Height          =   180
         Left            =   270
         TabIndex        =   4
         Top             =   240
         Width           =   1350
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   -60
      TabIndex        =   2
      Top             =   690
      Width           =   6885
   End
   Begin RTComctl3.CoolButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   4860
      TabIndex        =   1
      Top             =   3210
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
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
      MICON           =   "frmUnitManSimple.frx":0044
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
      Height          =   315
      Left            =   3510
      TabIndex        =   0
      Top             =   3210
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
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
      MICON           =   "frmUnitManSimple.frx":0060
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker dtpBeginDate 
      Height          =   300
      Left            =   2250
      TabIndex        =   5
      Top             =   990
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   529
      _Version        =   393216
      Format          =   62652416
      CurrentDate     =   36572
   End
   Begin MSComCtl2.DTPicker dtpEndDate 
      Height          =   300
      Left            =   2250
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
      Top             =   2985
      Width           =   8745
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "上车站(&T):"
      Height          =   180
      Left            =   1005
      TabIndex        =   11
      Top             =   2160
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "结束日期(&E):"
      Height          =   180
      Left            =   990
      TabIndex        =   9
      Top             =   1605
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "开始日期(&B):"
      Height          =   180
      Left            =   990
      TabIndex        =   8
      Top             =   1050
      Width           =   1080
   End
End
Attribute VB_Name = "frmUnitManSimple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Implements IConditionForm
Const cszFileName = "车站发出人数统计模板.xls"

Private m_rsData As Recordset
Public m_bOk As Boolean
Private m_vaCustomData As Variant
Dim m_aszTemp() As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdChart_Click()
    
    On Error GoTo Error_Handle
    '生成记录集
    Dim rsTemp As Recordset
    Dim oDss As New TicketUnitDim
    Dim i As Integer
    Dim frmTemp As frmChart
    oDss.Init m_oActiveUser
    Set rsTemp = oDss.GetCheckGateStatDetailSumByDate(dtpBeginDate.Value, dtpEndDate.Value, IIf(txtSellStationID.Text <> "", txtSellStationID.Text, ResolveDisplay(cboSellStation)))
    
    Dim rsData As New Recordset
    With rsData.Fields
        .Append "bus_date", adDate
        .Append "total_price", adBigInt
    End With
    rsData.Open
    For i = 1 To rsTemp.RecordCount
        rsData.AddNew
        rsData!bus_date = FormatDbValue(rsTemp!bus_date)
        rsData!total_price = FormatDbValue(rsTemp!total_price)
        rsTemp.MoveNext
        rsData.Update
    Next i
    
    Dim rsdata2 As New Recordset
    With rsdata2.Fields
        .Append "bus_date", adDate
        .Append "total_number", adBigInt
    End With
    rsdata2.Open
    rsTemp.MoveFirst
    For i = 1 To rsTemp.RecordCount
        rsdata2.AddNew
        rsdata2!bus_date = FormatDbValue(rsTemp!bus_date)
        rsdata2!total_number = FormatDbValue(rsTemp!passenger_number)
        rsTemp.MoveNext
        rsdata2.Update
    Next i
    
    Me.Hide
    Set frmTemp = New frmChart
    frmTemp.ClearChart
    frmTemp.AddChart "人数", rsdata2
    frmTemp.AddChart "金额", rsData
    frmTemp.ShowChart "车站发出人数统计"
    Set frmTemp = Nothing
    Unload Me

    Exit Sub
Error_Handle:
    Set frmTemp = Nothing
    ShowErrorMsg
    
End Sub

Private Sub cmdok_Click()
    On Error GoTo Error_Handle
    '生成记录集
    Dim rsTemp As Recordset
    Dim oDss As New TicketUnitDim
    Dim rsData As New Recordset
    Dim i As Integer
    
    oDss.Init m_oActiveUser
    Set rsTemp = oDss.GetCheckGateStatDetailSumByDate(dtpBeginDate.Value, dtpEndDate.Value, IIf(txtSellStationID.Text <> "", txtSellStationID.Text, ResolveDisplay(cboSellStation)))
    Set m_rsData = rsTemp
    
    ReDim m_vaCustomData(1 To 5, 1 To 2)
    m_vaCustomData(1, 1) = "复核"
    m_vaCustomData(1, 2) = m_oActiveUser.UserName
    
    m_vaCustomData(2, 1) = "统计开始日期"
    m_vaCustomData(2, 2) = Format(dtpBeginDate.Value, "YYYY年MM月DD日")
    
    m_vaCustomData(3, 1) = "统计结束日期"
    m_vaCustomData(3, 2) = Format(dtpEndDate.Value, "YYYY年MM月DD日")
    
    Dim szSellStation As String
    ResolveDisplay cboSellStation, szSellStation
    m_vaCustomData(4, 1) = "上车站"
    m_vaCustomData(4, 2) = szSellStation
    
    m_vaCustomData(5, 1) = "制表人"
    m_vaCustomData(5, 2) = m_oActiveUser.UserID
    m_bOk = True
    Unload Me
    Exit Sub
Error_Handle:
    ShowErrorMsg
End Sub


Private Sub CoolButton1_Click()
DisplayHelp Me
End Sub

Private Sub Form_Load()

    Dim nCount As Integer
    Dim i As Integer
    m_bOk = False
    

'    dtpBeginDate.Value = DateAdd("d", -1, m_oParam.NowDate)
'    dtpEndDate.Value = DateAdd("d", -1, m_oParam.NowDate)
    '设置为上个月的一号到31
    Dim dyNow As Date
    dyNow = m_oParam.NowDate
    dtpBeginDate.Value = Format(DateAdd("m", -1, dyNow), "yyyy-mm-01")
    dtpEndDate.Value = DateAdd("d", -1, Format(dyNow, "yyyy-mm-01"))
    FillSellStation cboSellStation
    
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
'    '判断用户属于哪个上车站,如果为空则填充一个空行,再填充所有的上车站
'
'    '否则只填充用户所属的上车站
'End Sub



