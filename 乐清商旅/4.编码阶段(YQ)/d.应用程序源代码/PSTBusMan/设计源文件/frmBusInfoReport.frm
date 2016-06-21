VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Object = "{D11DF87C-EFEC-4838-B7E2-15462BF136FB}#2.1#0"; "RTReportLF.ocx"
Begin VB.Form frmBusInfoReport 
   BackColor       =   &H80000009&
   Caption         =   "计划车次信息报表"
   ClientHeight    =   6420
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9030
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6420
   ScaleWidth      =   9030
   WindowState     =   2  'Maximized
   Begin VB.PictureBox ptQuery 
      BackColor       =   &H8000000E&
      Height          =   5805
      Left            =   0
      ScaleHeight     =   5745
      ScaleWidth      =   2775
      TabIndex        =   5
      Top             =   90
      Width           =   2835
      Begin VB.Frame fraQuery 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1935
         Left            =   105
         TabIndex        =   6
         Top             =   1050
         Width           =   2535
      End
      Begin RTComctl3.CoolButton flblClose 
         Height          =   225
         Left            =   2520
         TabIndex        =   7
         Top             =   15
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   397
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
         MICON           =   "frmBusInfoReport.frx":0000
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
         Left            =   105
         TabIndex        =   8
         Top             =   1860
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   71303168
         CurrentDate     =   36523
      End
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   300
         Left            =   105
         TabIndex        =   9
         Top             =   1305
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   71303168
         CurrentDate     =   36523
      End
      Begin RTComctl3.CoolButton cmdCancel 
         Cancel          =   -1  'True
         Height          =   330
         Left            =   1395
         TabIndex        =   10
         Top             =   3120
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
         MICON           =   "frmBusInfoReport.frx":001C
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
         Height          =   330
         Left            =   120
         TabIndex        =   11
         Top             =   3120
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   582
         BTYPE           =   3
         TX              =   "查询"
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
         MICON           =   "frmBusInfoReport.frx":0038
         PICN            =   "frmBusInfoReport.frx":0054
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
         TabIndex        =   12
         Top             =   0
         Width           =   2775
         _ExtentX        =   4895
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
         BackColor       =   -2147483644
         HorizontalAlignment=   1
         Caption         =   "查询条件设定"
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000009&
         Caption         =   "开始日期(&S):"
         Height          =   225
         Left            =   105
         TabIndex        =   15
         Top             =   1065
         Width           =   1170
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000009&
         Caption         =   "结束日期(&E):"
         Height          =   315
         Left            =   105
         TabIndex        =   14
         Top             =   1650
         Width           =   1455
      End
      Begin VB.Label lblBusID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车次(B):"
         Height          =   180
         Left            =   105
         TabIndex        =   13
         Top             =   2205
         Visible         =   0   'False
         Width           =   720
      End
   End
   Begin VB.PictureBox ptResult 
      BackColor       =   &H8000000E&
      Height          =   5835
      Left            =   3105
      ScaleHeight     =   5775
      ScaleWidth      =   5145
      TabIndex        =   0
      Top             =   0
      Width           =   5205
      Begin VB.PictureBox ptQ 
         BackColor       =   &H8000000E&
         BorderStyle     =   0  'None
         Height          =   1200
         Left            =   -15
         Picture         =   "frmBusInfoReport.frx":03EE
         ScaleHeight     =   1200
         ScaleWidth      =   5100
         TabIndex        =   2
         Top             =   45
         Width           =   5100
         Begin VB.Image imgOpen 
            Height          =   240
            Left            =   0
            Picture         =   "frmBusInfoReport.frx":12E4
            Top             =   0
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "计划车次信息查询"
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
            Left            =   780
            TabIndex        =   3
            Top             =   750
            Width           =   1920
         End
      End
      Begin RTReportLF.RTReport RTReport 
         Height          =   2880
         Left            =   90
         TabIndex        =   1
         Top             =   1335
         Width           =   4380
         _ExtentX        =   7726
         _ExtentY        =   5080
      End
   End
   Begin RTComctl3.Spliter spQuery 
      Height          =   1170
      Left            =   2910
      TabIndex        =   4
      Top             =   2520
      Width           =   165
      _ExtentX        =   291
      _ExtentY        =   2064
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
   Begin RTComctl3.Spliter Spliter1 
      Height          =   1170
      Left            =   2910
      TabIndex        =   16
      Top             =   2520
      Width           =   165
      _ExtentX        =   291
      _ExtentY        =   2064
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
Attribute VB_Name = "frmBusInfoReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const cszPtTop = 1200

Private lMoveLeft As Long
Private m_nSeatCount As Integer
Private m_nStartSeatNo As Integer
Private F1Book As TTF160Ctl.F1Book


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    On Error GoTo ErrorHandle
    SetBusy
    F1Book.DeleteRange F1Book.TopRow, F1Book.LeftCol, F1Book.MaxRow, F1Book.MaxCol, F1ShiftRows
    
        PlanBusInfo
        
'        RTReport.SetFocus
    SetNormal
    ShowSBInfo ""
'    ShowSBInfo "共有" & m_lRange & "条记录", ESB_ResultCountInfo
Exit Sub
ErrorHandle:
    ShowSBInfo ""
    SetNormal
    ShowErrorMsg
End Sub


Private Sub flblClose_Click()
    ptQuery.Visible = False
    imgOpen.Visible = True
    lMoveLeft = 240
    spQuery.LayoutIt
End Sub

Private Sub Form_Activate()
    MDIScheme.SetPrintEnabled True
End Sub
Private Sub Form_Deactivate()
    MDIScheme.SetPrintEnabled False
End Sub

Private Sub Form_Load()
    spQuery.InitSpliter ptQuery, ptResult
    lMoveLeft = 0
    dtpStartDate.Value = Date
    dtpEndDate.Value = Date
    Set F1Book = RTReport.CellObject
    F1Book.ShowColHeading = True
    F1Book.ShowRowHeading = True
    flblClose_Click
    cmdOk_Click
End Sub

Private Sub Form_Resize()
    spQuery.LayoutIt
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIScheme.SetPrintEnabled False
End Sub

Private Sub imgOpen_Click()
    ptQuery.Visible = True
    imgOpen.Visible = False
    lMoveLeft = 0
    spQuery.LayoutIt
End Sub


Private Sub ptResult_Resize()
    Dim lTemp As Long
    lTemp = IIf((ptResult.ScaleHeight - cszPtTop) <= 0, lTemp, ptResult.ScaleHeight - cszPtTop)
    RTReport.Move 0, cszPtTop, ptResult.ScaleWidth, lTemp
End Sub

Private Sub PlanBusInfo()
    Dim oPlan As New BusProject
    Dim rstemp As Recordset
    Dim i As Integer
On Error GoTo ErrorHandle
    ShowSBInfo "获得车次信息..."
    oPlan.Init g_oActiveUser
    oPlan.Identify
    Set rstemp = oPlan.GetAllBusReport
    F1Book.MaxCol = 6
    F1Book.MaxRow = rstemp.RecordCount + 1
    WriteProcessBar , , rstemp.RecordCount

    F1Book.TextRC(1, 1) = "车次代码"
    F1Book.TextRC(1, 2) = "运行线路"
    F1Book.TextRC(1, 3) = "发车时间"
    F1Book.TextRC(1, 4) = "检票口"
    F1Book.TextRC(1, 5) = "状态"
     F1Book.TextRC(1, 6) = "当天车辆情况"
    For i = 2 To rstemp.RecordCount + 1
        WriteProcessBar , i - 1, rstemp.RecordCount
        F1Book.TextRC(i, 1) = Trim(rstemp!bus_id)
        F1Book.TextRC(i, 2) = Trim(rstemp!route_name)
        F1Book.TextRC(i, 3) = Format(rstemp!bus_start_time, "HH:MM")
        F1Book.TextRC(i, 4) = Trim(rstemp!check_gate_name)
        If rstemp!stop_start_date <> CDate(cszEmptyDateStr) And rstemp!stop_start_date >= Date Then
            F1Book.TextRC(i, 5) = "停班在[" & Format(rstemp!stop_start_date, "YYYY-MM-DD") & "到" & Format(rstemp!stop_end_date, "YYYY-MM-DD") & "]时段停班"
        Else
            If rstemp!stop_end_date = "2050-1-1" Then
                F1Book.TextRC(i, 5) = "车次停班"
            Else
                F1Book.TextRC(i, 5) = "运行"
            End If
        End If
        F1Book.TextRC(i, 6) = Trim(rstemp!vehicle_id) & "[" & Trim(rstemp!license_tag_no) & "]"
        F1Book.ColWidth(2) = 3000
        F1Book.ColWidth(6) = 4000
    rstemp.MoveNext
    Next
    WriteProcessBar False
Exit Sub
ErrorHandle:
    SetNormal
    ShowErrorMsg
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
