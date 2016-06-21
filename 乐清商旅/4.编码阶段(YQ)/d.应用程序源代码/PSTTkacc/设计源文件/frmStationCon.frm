VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmStationCon 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "站点简报统计"
   ClientHeight    =   3870
   ClientLeft      =   2865
   ClientTop       =   3870
   ClientWidth     =   5085
   HelpContextID   =   60000190
   Icon            =   "frmStationCon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtSellStationID 
      Height          =   300
      Left            =   2070
      TabIndex        =   16
      Top             =   2730
      Width           =   2010
   End
   Begin RTComctl3.CoolButton cmdChart 
      Height          =   315
      Left            =   30
      TabIndex        =   15
      Top             =   3360
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
      MICON           =   "frmStationCon.frx":000C
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
      Left            =   1260
      TabIndex        =   14
      Top             =   3360
      Width           =   1185
      _ExtentX        =   2090
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
      MICON           =   "frmStationCon.frx":0028
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
      Left            =   2070
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2325
      Width           =   2010
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   -30
      TabIndex        =   8
      Top             =   690
      Width           =   7200
   End
   Begin RTComctl3.CoolButton cmdOk 
      Default         =   -1  'True
      Height          =   315
      Left            =   2580
      TabIndex        =   2
      Top             =   3360
      Width           =   1185
      _ExtentX        =   2090
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
      MICON           =   "frmStationCon.frx":0044
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
      Left            =   3870
      TabIndex        =   1
      Top             =   3360
      Width           =   1185
      _ExtentX        =   2090
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
      MICON           =   "frmStationCon.frx":0060
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin FText.asFlatTextBox txtAreaCode 
      Height          =   315
      Left            =   2070
      TabIndex        =   0
      Top             =   1860
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
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
   Begin MSComCtl2.DTPicker dtpEndDate 
      Height          =   315
      Left            =   2070
      TabIndex        =   3
      Top             =   1410
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   556
      _Version        =   393216
      Format          =   62652416
      CurrentDate     =   36572
   End
   Begin MSComCtl2.DTPicker dtpBeginDate 
      Height          =   315
      Left            =   2085
      TabIndex        =   4
      Top             =   960
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   556
      _Version        =   393216
      Format          =   62652416
      CurrentDate     =   36572
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   -30
      ScaleHeight     =   735
      ScaleWidth      =   7200
      TabIndex        =   10
      Top             =   0
      Width           =   7200
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请选择条件:"
         Height          =   180
         Left            =   270
         TabIndex        =   11
         Top             =   270
         Width           =   990
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   " "
      Enabled         =   0   'False
      Height          =   3120
      Left            =   -120
      TabIndex        =   9
      Top             =   3120
      Width           =   8745
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "上车站(&T):"
      Height          =   180
      Left            =   975
      TabIndex        =   13
      Top             =   2385
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "开始日期(&B):"
      Height          =   180
      Left            =   975
      TabIndex        =   7
      Top             =   1020
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "结束日期(&E):"
      Height          =   180
      Left            =   975
      TabIndex        =   6
      Top             =   1470
      Width           =   1080
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "地区(&A):"
      Height          =   180
      Left            =   975
      TabIndex        =   5
      Top             =   1920
      Width           =   720
   End
End
Attribute VB_Name = "frmStationCon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Implements IConditionForm
Const cszFileName = "站点售票营收简报模板.xls"
Public m_bOk As Boolean

Private m_rsData As Recordset
Private m_vaCustomData As Variant

Dim m_szCode As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdChart_Click()
    
    On Error GoTo Error_Handle
    '生成记录集
    Dim rsTemp As Recordset
    Dim i As Integer
    Dim oObject As New TicketStationDim
    Dim frmTemp As frmChart
    oObject.Init m_oActiveUser
    Set rsTemp = oObject.StationCount(dtpBeginDate.Value, dtpEndDate.Value, m_szCode, IIf(txtSellStationID.Text <> "", txtSellStationID.Text, ResolveDisplay(cboSellStation)))
    
    Dim rsData As New Recordset
    With rsData.Fields
        .Append "station_name", adBSTR
        .Append "passenger_number", adBigInt
    End With
    rsData.Open
    rsTemp.MoveFirst
    For i = 1 To rsTemp.RecordCount
        rsData.AddNew
        rsData!station_name = FormatDbValue(rsTemp!station_name)
        rsData!passenger_number = FormatDbValue(rsTemp!passenger_number)
        rsTemp.MoveNext
        rsData.Update
    Next i
    
    Dim rsdata2 As New Recordset
    With rsdata2.Fields
        .Append "station_name", adBSTR
        .Append "total_ticket_price", adBigInt
    End With
    rsdata2.Open
    rsTemp.MoveFirst
    For i = 1 To rsTemp.RecordCount
        rsdata2.AddNew
        rsdata2!station_name = FormatDbValue(rsTemp!station_name)
        rsdata2!total_ticket_price = FormatDbValue(rsTemp!total_ticket_price)
        rsTemp.MoveNext
        rsdata2.Update
    Next i

    Me.Hide
    Set frmTemp = New frmChart
    frmTemp.ClearChart
    frmTemp.AddChart "实售张数", rsData
    frmTemp.AddChart "营收合计", rsdata2
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
    Dim adbTemp() As Double
    Dim oObject As New TicketStationDim
    
    oObject.Init m_oActiveUser
    oDss.Init m_oActiveUser
    Set rsTemp = oObject.StationCount(dtpBeginDate.Value, dtpEndDate.Value, m_szCode, IIf(txtSellStationID.Text <> "", txtSellStationID.Text, ResolveDisplay(cboSellStation)))
    Set m_rsData = rsTemp
    
    adbTemp = oDss.ProvinceSum(dtpBeginDate.Value, dtpEndDate.Value, m_szCode, IIf(txtSellStationID.Text <> "", txtSellStationID.Text, ResolveDisplay(cboSellStation)))
    
    
    
    ReDim m_vaCustomData(1 To 9, 1 To 2)
    m_vaCustomData(1, 1) = "统计开始日期"
    m_vaCustomData(1, 2) = Format(dtpBeginDate.Value, "YYYY年MM月DD日")
    
    m_vaCustomData(2, 1) = "统计结束日期"
    m_vaCustomData(2, 2) = Format(dtpEndDate.Value, "YYYY年MM月DD日")
    m_vaCustomData(3, 1) = "省内人数"
    m_vaCustomData(4, 1) = "省内总额"
    m_vaCustomData(5, 1) = "省内周转量"
    m_vaCustomData(6, 1) = "省外人数"
    m_vaCustomData(7, 1) = "省外总额"
    m_vaCustomData(8, 1) = "省外周转量"
    For i = 1 To ArrayLength(adbTemp)
        If adbTemp(i, 1) = EA_nInCity Or adbTemp(i, 1) = EA_nOutCity Then
            '省内
            m_vaCustomData(3, 2) = adbTemp(i, 2) + m_vaCustomData(3, 2)
            m_vaCustomData(4, 2) = adbTemp(i, 3) + m_vaCustomData(4, 2)
            m_vaCustomData(5, 2) = adbTemp(i, 4) + m_vaCustomData(5, 2)
            
        ElseIf adbTemp(i, 1) = EA_nOutProvince Then
            '省外
            
            m_vaCustomData(6, 2) = adbTemp(i, 2)
            m_vaCustomData(7, 2) = adbTemp(i, 3)
            m_vaCustomData(8, 2) = adbTemp(i, 4)
            
        End If
    Next i
    
    m_vaCustomData(9, 1) = "制表人"
    m_vaCustomData(9, 2) = m_oActiveUser.UserID
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
    m_bOk = False
    m_szCode = ""

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

Private Sub txtAreaCode_ButtonClick()
    Dim aszAreaCode() As String
    aszAreaCode = m_oShell.SelectArea
    txtAreaCode.Text = TeamToString(aszAreaCode, 2)
    m_szCode = TeamToString(aszAreaCode, 1)
End Sub



'Private Sub FillSellStation()
'    '判断用户属于哪个上车站,如果为空则填充一个空行,再填充所有的上车站
'
'    '否则只填充用户所属的上车站
'End Sub


