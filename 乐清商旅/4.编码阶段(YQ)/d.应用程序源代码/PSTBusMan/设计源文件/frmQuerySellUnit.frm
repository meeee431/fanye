VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Object = "{D11DF87C-EFEC-4838-B7E2-15462BF136FB}#1.4#0"; "RTReportlf.ocx"
Begin VB.Form frmQuerySellUnit 
   BackColor       =   &H00E0E0E0&
   Caption         =   "车站售票即时统计"
   ClientHeight    =   6180
   ClientLeft      =   2280
   ClientTop       =   2565
   ClientWidth     =   10185
   Icon            =   "frmQuerySellUnit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6180
   ScaleWidth      =   10185
   Begin VB.PictureBox ptQuery 
      BackColor       =   &H80000009&
      Height          =   6345
      Left            =   150
      ScaleHeight     =   6285
      ScaleWidth      =   2790
      TabIndex        =   4
      Top             =   90
      Width           =   2850
      Begin VB.ComboBox cboQueryType 
         Height          =   300
         ItemData        =   "frmQuerySellUnit.frx":000C
         Left            =   90
         List            =   "frmQuerySellUnit.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   2670
         Width           =   2625
      End
      Begin VB.OptionButton optSellTime 
         BackColor       =   &H00FFFFFF&
         Caption         =   "按售票时间(&L)"
         Height          =   345
         Left            =   90
         TabIndex        =   11
         Top             =   660
         Value           =   -1  'True
         Width           =   1605
      End
      Begin VB.OptionButton optBusDate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "按发车日期(&B)"
         Height          =   285
         Left            =   90
         TabIndex        =   10
         Top             =   420
         Width           =   1545
      End
      Begin RTComctl3.CoolButton flbClose 
         Height          =   240
         Left            =   2490
         TabIndex        =   5
         Top             =   15
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   423
         BTYPE           =   8
         TX              =   "r"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Marlett"
            Size            =   9
            Charset         =   2
            Weight          =   500
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
         MICON           =   "frmQuerySellUnit.frx":0010
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin RTComctl3.FlatLabel FlatLabel1 
         Height          =   285
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   2790
         _ExtentX        =   4921
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
         VerticalAlignment=   1
         HorizontalAlignment=   1
         MarginLeft      =   0
         MarginTop       =   0
         Caption         =   "查询条件设定"
      End
      Begin MSComCtl2.DTPicker dtpEndTime 
         Height          =   285
         Left            =   90
         TabIndex        =   8
         Top             =   1935
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   19726339
         CurrentDate     =   37512
      End
      Begin MSComCtl2.DTPicker dtpStartTime 
         Height          =   285
         Left            =   90
         TabIndex        =   9
         Top             =   1290
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   19726339
         CurrentDate     =   37512
      End
      Begin RTComctl3.CoolButton cmdQuery 
         Default         =   -1  'True
         Height          =   330
         Left            =   180
         TabIndex        =   17
         Top             =   3810
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   582
         BTYPE           =   3
         TX              =   "查询(&Q)"
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
         MICON           =   "frmQuerySellUnit.frx":002C
         PICN            =   "frmQuerySellUnit.frx":0048
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
         Height          =   330
         Left            =   1470
         TabIndex        =   18
         Top             =   3810
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   582
         BTYPE           =   3
         TX              =   "关闭(&C)"
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
         MICON           =   "frmQuerySellUnit.frx":03E2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin FText.asFlatTextBox txtObject 
         Height          =   285
         Left            =   90
         TabIndex        =   19
         Top             =   3345
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "检票口(&O):"
         Height          =   180
         Left            =   90
         TabIndex        =   16
         Top             =   3090
         Width           =   900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "查询类型(&T):"
         Height          =   180
         Left            =   90
         TabIndex        =   15
         Top             =   2370
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "起始时间(&R):"
         Height          =   180
         Left            =   90
         TabIndex        =   14
         Top             =   1020
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "结束日期(&E):"
         Height          =   180
         Left            =   105
         TabIndex        =   13
         Top             =   1665
         Width           =   1080
      End
   End
   Begin VB.PictureBox ptResult 
      BackColor       =   &H80000009&
      Height          =   6030
      Left            =   3990
      ScaleHeight     =   5970
      ScaleWidth      =   5355
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      Begin VB.PictureBox ptQ 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1155
         Left            =   0
         Picture         =   "frmQuerySellUnit.frx":03FE
         ScaleHeight     =   1155
         ScaleWidth      =   5640
         TabIndex        =   1
         Top             =   0
         Width           =   5640
         Begin VB.Image imgOpen 
            Height          =   240
            Left            =   0
            Picture         =   "frmQuerySellUnit.frx":12F4
            Top             =   0
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "售票即时查询及统计"
            BeginProperty Font 
               Name            =   "黑体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   840
            TabIndex        =   2
            Top             =   825
            Width           =   2160
         End
      End
      Begin RTReportLF.RTReport RTReport 
         Height          =   3225
         Left            =   240
         TabIndex        =   3
         Top             =   2205
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   5689
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   8850
      Top             =   3300
   End
   Begin MSComDlg.CommonDialog SaveDialogue 
      Left            =   6870
      Top             =   4230
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RTComctl3.Spliter spQuery 
      Height          =   1320
      Left            =   3510
      TabIndex        =   7
      Top             =   1920
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   2328
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmQuerySellUnit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'按照不同的统计方式,即时从售票中统计出售票数据

Const cnTop = 1200

Private m_lMoveLeft As Long
Private m_lRange As Long '写进度条用


Const cszBusIDQuery = "按车次查询.xls"
Const cszRouteIDQuery = "按线路查询.xls"
Const cszCheckGateIDQuery = "按检票口查询.xls"
Const cszStationIDQuery = "按站点查询.xls"
'Public m_bDefault As Boolean
'Public m_nQueryType As Integer
'Public m_szTitleID As String
'Public optbusdate As Boolean
'Private m_rsTitle As New Recordset
Private m_szTemplatePathName As String


Private Sub cboQueryType_Click()
    Select Case cboQueryType.ListIndex
        Case 0  '按线路查询
            lblTitle.Caption = "线路代码(&O):"
            m_szTemplatePathName = App.Path & "\" & cszRouteIDQuery
        Case 1  '按车次查询
            lblTitle.Caption = "车次代码(&O):"
            m_szTemplatePathName = App.Path & "\" & cszBusIDQuery
        Case 2  '按检票口查询
            lblTitle.Caption = "检票口代码(&O):"
            m_szTemplatePathName = App.Path & "\" & cszCheckGateIDQuery
        Case 3  '按站点查询
            lblTitle.Caption = "站点代码(&O):"
            m_szTemplatePathName = App.Path & "\" & cszStationIDQuery
    End Select
    txtObject.Text = ""
    
End Sub

Private Sub cmdQuery_Click()
    m_lRange = 0
    Query
End Sub



Private Sub Form_Load()
    spQuery.InitSpliter ptQuery, ptResult
    
    dtpStartTime.Value = Date & " 00:00"
    dtpEndTime.Value = Date & " 23:59:59"
    txtObject.Text = ""
    SetDateFormat
    InitCboQueryType
    cboQueryType_Click
End Sub


Private Sub Query()
On Error GoTo ErrHandle
'    '开始查询
    Dim rsTemp As Recordset
    Dim oTemp As New TicketUnitDim
    Dim i As Integer
    Dim aszTemp As Variant
    If DateDiff("d", dtpStartTime.Value, dtpEndTime.Value) >= 2 Then
        MsgBox "由于数据量太大,所以只能查两天内的数据" & Chr(10) & "如果多于两天,请到票款结算中去统计", vbInformation, Me.Caption
        dtpEndTime.SetFocus
        Exit Sub
    End If

    oTemp.Init g_oActiveUser
    Select Case cboQueryType.ListIndex
    Case 0  '按线路查询
        If optBusDate Then
            Set rsTemp = oTemp.GetBusDateRouteQuery(dtpStartTime.Value, dtpEndTime.Value, txtObject.Text)
        Else
            Set rsTemp = oTemp.GetSellRouteID(dtpStartTime.Value, dtpEndTime.Value, txtObject.Text)
        End If
    Case 1 '按车次查询
        If optBusDate Then
            Set rsTemp = oTemp.GetBusDateBusIDQuery(dtpStartTime.Value, dtpEndTime.Value, txtObject.Text)
        Else
            Set rsTemp = oTemp.GetSellBusID(dtpStartTime.Value, dtpEndTime.Value, txtObject.Text)
        End If
    Case 2  '按检票口查询
        If optBusDate Then
            Set rsTemp = oTemp.GetBusDateCheckGateQuery(dtpStartTime.Value, dtpEndTime.Value, txtObject.Text)
        Else
            Set rsTemp = oTemp.GetSellCheckGateID(dtpStartTime.Value, dtpEndTime.Value, txtObject.Text)
        End If
    Case 3  '按站点查询
        If optBusDate Then
            Set rsTemp = oTemp.GetBusDateStation(dtpStartTime.Value, dtpEndTime.Value, txtObject.Text)
        Else
            Set rsTemp = oTemp.GetSellStationID(dtpStartTime.Value, dtpEndTime.Value, txtObject.Text)
        End If
    End Select
    ReDim aszTemp(1 To 2, 1 To 2)
    aszTemp(1, 1) = "统计开始日期"
    aszTemp(1, 2) = dtpStartTime.Value
    aszTemp(2, 1) = "统计结束日期"
    aszTemp(2, 2) = dtpEndTime.Value

    '填充票价记录集
    '设置固定行,列可见性
    RTReport.LeftLabelVisual = True
    RTReport.TopLabelVisual = True
    RTReport.TemplateFile = m_szTemplatePathName
    RTReport.ShowReport rsTemp, aszTemp

    WriteProcessBar False
    ShowSBInfo
    ShowSBInfo "共有" & m_lRange & "条记录", ESB_ResultCountInfo
    Exit Sub
ErrHandle:
    WriteProcessBar False
    ShowSBInfo
    ShowErrorMsg
End Sub

Private Sub optBusDate_Click()
    SetDateFormat
End Sub

Private Sub SetDateFormat()
    If optBusDate.Value = True Then
        dtpEndTime.Format = dtpLongDate
        dtpStartTime.Format = dtpLongDate
    Else
        dtpEndTime.Format = dtpCustom
        dtpStartTime.Format = dtpCustom
        dtpEndTime.CustomFormat = "yyyy-MM-dd HH:mm"
        dtpStartTime.CustomFormat = "yyyy-MM-dd HH:mm"
    End If

End Sub

Private Sub optSellTime_Click()
    SetDateFormat
End Sub
'
Private Sub txtObject_ButtonClick()
    '选择对应的对象
    Dim oShell As New CommDialog
    Dim aszTemp() As String
    On Error GoTo ErrorHandle
    oShell.Init g_oActiveUser
    Select Case cboQueryType.ListIndex
    Case 0
        '线路
        aszTemp = oShell.SelectRoute(True)
    Case 1
        '车次
        aszTemp = oShell.SelectBus(True)
    Case 2
        '检票口
        aszTemp = oShell.SelectCheckGate(True)
    Case 3
        '站点
        aszTemp = oShell.SelectStation(, True)
    End Select
    txtObject.Text = TeamToString(aszTemp, 1)
    Set oShell = Nothing
    Exit Sub
ErrorHandle:
    Set oShell = Nothing
    ShowErrorMsg
End Sub

Private Sub InitCboQueryType()
    cboQueryType.AddItem "按线路查询"
    cboQueryType.AddItem "按车次查询"
    cboQueryType.AddItem "按检票口查询"
    cboQueryType.AddItem "按站点查询"
    cboQueryType.ListIndex = 0
End Sub


Private Sub Form_Activate()
    Form_Resize
    MDIScheme.SetPrintEnabled True
End Sub

Private Sub Form_Deactivate()
    MDIScheme.SetPrintEnabled False
End Sub

Private Sub Form_Resize()
    '重绘
    spQuery.LayoutIt
End Sub

Private Sub ptQuery_Resize()
    flbClose.Move ptQuery.ScaleWidth - 250
End Sub
Private Sub ptResult_Resize()
    Dim lTemp As Long
    lTemp = IIf((ptResult.ScaleHeight - cnTop) <= 0, lTemp, ptResult.ScaleHeight - cnTop)
    RTReport.Move 0, cnTop, ptResult.ScaleWidth, lTemp
    'lblTitle.Move 60 + m_lMoveLeft, 50, ptResult.ScaleWidth

    FlatLabel1.Width = ptQuery.ScaleWidth
    flbClose.Left = FlatLabel1.Left + FlatLabel1.Width - flbClose.Width - 30
End Sub

Private Sub cmdCancel_Click()
    '关闭
    Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    MDIScheme.SetPrintEnabled False
End Sub
Private Sub RTReport_SetProgressRange(ByVal lRange As Variant)
    m_lRange = lRange
    
End Sub

Private Sub RTReport_SetProgressValue(ByVal lValue As Variant)
    WriteProcessBar True, lValue, m_lRange, "正在填充统计信息..."
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    
End Sub
Private Sub flbClose_Click()
    '关闭左边部分
    ptQuery.Visible = False
    imgOpen.Visible = True
    m_lMoveLeft = 240
    'lblTitle.Move 60 + m_lMoveLeft
    spQuery.LayoutIt
End Sub

Private Sub imgOpen_Click()
    '打开左边部分
    ptQuery.Visible = True
    imgOpen.Visible = False
    m_lMoveLeft = 0
    'lblTitle.Move 60 + m_lMoveLeft
    spQuery.LayoutIt
End Sub
Private Sub EnabledQuery()
    '查询按钮是否可用
'    If txtObject.Text <> "" Then
'        cmdQuery.Enabled = True
'    Else
'        cmdQuery.Enabled = False
'    End If
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
'导出文件
Public Sub ExportFile()
    Dim szFileName As String
    szFileName = RTReport.OpenDialog(EDialogType.EXPORT_FILE)
End Sub
'导出文件并打开
Public Sub ExportFileOpen()
    Dim szFileName As String
    szFileName = RTReport.OpenDialog(EDialogType.EXPORT_FILE)
    If szFileName <> "" Then
        OpenLinkedFile szFileName
    End If
End Sub

Private Sub txtObject_Change()
    EnabledQuery
End Sub
