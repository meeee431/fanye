VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Object = "{D11DF87C-EFEC-4838-B7E2-15462BF136FB}#2.1#0"; "RTReportLF.ocx"
Begin VB.Form frmChkTkQuery 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "综合检票查询"
   ClientHeight    =   4605
   ClientLeft      =   2475
   ClientTop       =   3375
   ClientWidth     =   8610
   HelpContextID   =   10000370
   Icon            =   "frmChkTkQuery.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   8610
   ShowInTaskbar   =   0   'False
   Begin RTComctl3.CoolButton CoolButton1 
      Height          =   315
      Left            =   7290
      TabIndex        =   13
      Top             =   1860
      Width           =   1155
      _ExtentX        =   2037
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
      MICON           =   "frmChkTkQuery.frx":014A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTReportLF.RTReport RTReport 
      Height          =   3660
      Left            =   60
      TabIndex        =   8
      Top             =   885
      Width           =   6990
      _ExtentX        =   12330
      _ExtentY        =   6456
   End
   Begin MSComctlLib.ImageList imgObject 
      Left            =   8175
      Top             =   2940
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChkTkQuery.frx":0166
            Key             =   "CheckGate"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChkTkQuery.frx":02C0
            Key             =   "Station"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChkTkQuery.frx":041A
            Key             =   "Route"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChkTkQuery.frx":0576
            Key             =   "Bus"
         EndProperty
      EndProperty
   End
   Begin RTComctl3.TextButtonBox txtQuery 
      Height          =   315
      Left            =   4860
      TabIndex        =   5
      Top             =   45
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   556
      Enabled         =   0   'False
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
   Begin MSComctlLib.ImageCombo imgcbo 
      Height          =   315
      Left            =   1260
      TabIndex        =   4
      Top             =   30
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   556
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Locked          =   -1  'True
      ImageList       =   "imgObject"
   End
   Begin MSComCtl2.DTPicker dtpEndDate 
      Height          =   315
      Left            =   4875
      TabIndex        =   1
      ToolTipText     =   "查询结算日期"
      Top             =   450
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   123731971
      UpDown          =   -1  'True
      CurrentDate     =   36544
   End
   Begin MSComCtl2.DTPicker dtpStartDate 
      Height          =   315
      Left            =   1260
      TabIndex        =   0
      ToolTipText     =   "查询开始日期"
      Top             =   450
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   123731971
      UpDown          =   -1  'True
      CurrentDate     =   36544.6041666667
   End
   Begin RTComctl3.CoolButton cmdFind 
      Height          =   315
      Left            =   7260
      TabIndex        =   9
      Top             =   120
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "查询(&F)"
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
      MICON           =   "frmChkTkQuery.frx":06D0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdPrint 
      Height          =   315
      Left            =   7260
      TabIndex        =   10
      Top             =   945
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "打印(&P)"
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
      MICON           =   "frmChkTkQuery.frx":06EC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdPreview 
      Height          =   315
      Left            =   7260
      TabIndex        =   11
      Top             =   525
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "预览(&O)"
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
      MICON           =   "frmChkTkQuery.frx":0708
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
      Left            =   7260
      TabIndex        =   12
      Top             =   1380
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "关闭"
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
      MICON           =   "frmChkTkQuery.frx":0724
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblQuery 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "详细范围(&Y):"
      Height          =   180
      Left            =   3720
      TabIndex        =   7
      Top             =   120
      Width           =   1080
   End
   Begin VB.Label lblClass 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "查询范围(&I):"
      Height          =   180
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1080
   End
   Begin VB.Label lblEndDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "结束日期(&E):"
      Height          =   180
      Left            =   3720
      TabIndex        =   3
      Top             =   510
      Width           =   1080
   End
   Begin VB.Label lblStartDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "开始日期(&S):"
      Height          =   180
      Left            =   120
      TabIndex        =   2
      Top             =   510
      Width           =   1080
   End
End
Attribute VB_Name = "frmChkTkQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"38952F71032A"
Option Explicit
Private m_nCount As Integer
Private m_oReport As New CTReport
'Private rsStemp As Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub

'根据条件查询
Private Sub cmdFind_Click()
On Error GoTo Here
    Me.MousePointer = vbHourglass
    ShowSBInfo "开始查询..."
    Query
    ShowSBInfo ""
    Me.MousePointer = vbDefault
    cmdPrint.Enabled = True
    cmdPreview.Enabled = True
Exit Sub
Here:
    ShowSBInfo ""
    Me.MousePointer = vbDefault
    ShowErrorMsg
End Sub

Private Sub cmdPreview_Click()
  PrintPreview
End Sub

Private Sub cmdPrint_Click()
  PrintReport
End Sub

Private Sub CoolButton1_Click()
DisplayHelp Me
End Sub

Private Sub Form_Load()
    AlignFormPos Me
    imgcbo.ComboItems.Add(, "Bus", "车次", "Bus", , 0).Tag = "Bus"
    imgcbo.ComboItems.Add(, "Route", "线路", "Route", , 0).Tag = "Route"
    imgcbo.ComboItems.Add(, "Station", "站点", "Station", , 0).Tag = "Station"
    imgcbo.ComboItems.Add(, "CheckGate", "检票口", "CheckGate", , 0).Tag = "CheckGate"
    imgcbo.ComboItems.Item(1).Selected = True
    dtpStartDate.Enabled = True
    dtpEndDate.Enabled = True
    lblStartDate.Enabled = True
    lblEndDate.Enabled = True
    imgcbo.Enabled = True
    txtQuery.Enabled = True
    lblClass.Enabled = True
    lblQuery.Enabled = True
    dtpStartDate.Value = Date
    dtpEndDate.Value = Date & " " & "23:59"
    m_oReport.Init g_oActiveUser
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveFormPos Me
End Sub

Private Sub imgcbo_Change()
    imgcbo_Click
       
End Sub
Private Sub imgcbo_Click()
    If imgcbo.Text <> "" Then txtQuery.Enabled = True
End Sub


Private Sub txtQuery_Click()
    Dim oOpen As New CommDialog
    Dim aszTemp() As String
    oOpen.Init g_oActiveUser
    Select Case imgcbo.SelectedItem.Key
           Case "Bus"
           aszTemp = oOpen.SelectBus(True)
           Case "Route"
           aszTemp = oOpen.SelectRoute(True)
           Case "Station"
           aszTemp = oOpen.SelectStation(, True)
           Case "CheckGate"
           aszTemp = oOpen.SelectCheckGate(True)
    End Select
    If ArrayLength(aszTemp) < 1 Then Exit Sub
    txtQuery.Text = TeamToString(aszTemp) 'Trim(aszTemp(1, 1)) & "[" & Trim(aszTemp(1, 2)) & "]"
End Sub


Private Sub Query()
    Dim m_vaCustomData As Variant
    Dim m_rsData As Recordset
    Dim dtStart As Date
    Dim dtEnd As Date
    Dim szQuery As String
    dtStart = dtpStartDate.Value
    dtEnd = dtpEndDate.Value
    szQuery = txtQuery.Text
    On Error GoTo ErrorHandle
    'vaCustomData中为要显示的表头表尾信息
'    Set m_rsData = rsStemp
    ReDim m_vaCustomData(1 To 4, 1 To 2)
    m_vaCustomData(1, 1) = "统计开始月份"
    m_vaCustomData(1, 2) = Format(dtpStartDate.Value, "YYYY年MM月DD日")
    m_vaCustomData(2, 1) = "统计结束月份"
    m_vaCustomData(2, 2) = Format(dtpEndDate.Value, "YYYY年MM月DD日")
    m_vaCustomData(3, 1) = "制单日期"
    m_vaCustomData(3, 2) = Format(Date, "YYYY年MM月DD日")
    m_vaCustomData(4, 1) = "制单"
    m_vaCustomData(4, 2) = g_oActiveUser.UserName
'    m_vaCustomData(5, 1) = "检票人数"
'    m_vaCustomData(5, 2) = Val(LvObject.TextMatrix(LvObject.Rows - 1, 3))
'    m_vaCustomData(6, 1) = "检票金额"
'    m_vaCustomData(6, 2) = Format(Val(LvObject.TextMatrix(LvObject.Rows - 1, 4)), "0.00")
'    Select Case imgcbo.SelectedItem.Tag
'           Case "Bus"
'               Set m_rsData = m_oReport.GetBusTicketReport(dtStart, dtEnd, szQuery)
'           Case "Route"
'               Set m_rsData = m_oReport.GetRouteTicketReport(dtStart, dtEnd, szQuery)
'           Case "Station"
'               Set m_rsData = m_oReport.GetStationTicketReport(dtStart, dtEnd, szQuery)
'           Case "CheckGate"
'               Set m_rsData = m_oReport.GetCheckGateTicketReport(dtStart, dtEnd, szQuery)
'    End Select
    Select Case imgcbo.SelectedItem.Tag
       Case "Bus"
           Set m_rsData = m_oReport.GetBusTicketReport(dtStart, dtEnd, szQuery, g_oActiveUser.SellStationID)
       Case "Route"
           Set m_rsData = m_oReport.GetRouteTicketReport(dtStart, dtEnd, szQuery, g_oActiveUser.SellStationID)
       Case "Station"
           Set m_rsData = m_oReport.GetStationTicketReport(dtStart, dtEnd, szQuery, g_oActiveUser.SellStationID)
       Case "CheckGate"
           Set m_rsData = m_oReport.GetCheckGateTicketReport(dtStart, dtEnd, szQuery, g_oActiveUser.SellStationID)
    End Select
    Select Case imgcbo.SelectedItem.Index
       Case 1
           RTReport.TemplateFile = App.Path & "\车次检票统计报表模板.xls"   '模板名
       Case 2
            RTReport.TemplateFile = App.Path & "\线路检票统计报表模板.xls"  '模板名
       Case 3
            RTReport.TemplateFile = App.Path & "\站点检票统计报表模板.xls"  '模板名
       Case 4
           RTReport.TemplateFile = App.Path & "\检票口检票统计报表模板.xls"  '模板名
    End Select
    RTReport.ShowReport m_rsData, m_vaCustomData   '用记录集填充
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub PrintReport()
    '打印
    On Error Resume Next
    RTReport.PrintReport True
End Sub

Private Sub PrintPreview()
    '打印预览
    On Error Resume Next
    RTReport.PrintView
End Sub
