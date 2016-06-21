VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmSelSheet 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "选择路单"
   ClientHeight    =   6195
   ClientLeft      =   1920
   ClientTop       =   2925
   ClientWidth     =   10590
   Icon            =   "frmSelSheet.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   10590
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox txtBusSerialNO 
      Height          =   285
      Left            =   7335
      TabIndex        =   15
      Top             =   510
      Width           =   1635
   End
   Begin MSComCtl2.DTPicker dtpEndDate 
      Height          =   285
      Left            =   4365
      TabIndex        =   9
      Top             =   105
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   503
      _Version        =   393216
      Format          =   62652416
      CurrentDate     =   38535
   End
   Begin MSComCtl2.DTPicker dtpStartDate 
      Height          =   285
      Left            =   1365
      TabIndex        =   7
      Top             =   105
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   503
      _Version        =   393216
      Format          =   62652416
      CurrentDate     =   38535
   End
   Begin RTComctl3.TextButtonBox txtBusID 
      Height          =   285
      Left            =   4350
      TabIndex        =   5
      Top             =   510
      Width           =   1530
      _ExtentX        =   2699
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
   End
   Begin RTComctl3.TextButtonBox txtVehicleID 
      Height          =   285
      Left            =   990
      TabIndex        =   4
      Top             =   510
      Width           =   1995
      _ExtentX        =   3519
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
   End
   Begin VB.TextBox txtSheetID 
      Height          =   285
      Left            =   7335
      TabIndex        =   1
      Top             =   105
      Width           =   1635
   End
   Begin RTComctl3.CoolButton cmdSel 
      Height          =   345
      Left            =   9300
      TabIndex        =   11
      Top             =   720
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "选择(&L)"
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
      MICON           =   "frmSelSheet.frx":000C
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
      Left            =   9300
      TabIndex        =   12
      Top             =   1755
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "取消(&C)"
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
      MICON           =   "frmSelSheet.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdFind 
      Default         =   -1  'True
      Height          =   345
      Left            =   9300
      TabIndex        =   13
      Top             =   180
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   609
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
      MICON           =   "frmSelSheet.frx":0044
      PICN            =   "frmSelSheet.frx":0060
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ListView lvSheet 
      Height          =   5190
      Left            =   105
      TabIndex        =   10
      Top             =   900
      Width           =   9060
      _ExtentX        =   15981
      _ExtentY        =   9155
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "注:当车次序号为-1或空时,表明为所有的车次序号"
      ForeColor       =   &H000000FF&
      Height          =   720
      Left            =   9210
      TabIndex        =   16
      Top             =   2475
      Width           =   1365
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车次序号(&R):"
      Height          =   180
      Left            =   6060
      TabIndex        =   14
      Top             =   540
      Width           =   1080
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "结束日期(&E):"
      Height          =   180
      Left            =   3180
      TabIndex        =   8
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "开始日期(&S):"
      Height          =   180
      Left            =   165
      TabIndex        =   6
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车辆(&V):"
      Height          =   180
      Left            =   165
      TabIndex        =   3
      Top             =   555
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车次代码(&B):"
      Height          =   180
      Left            =   3090
      TabIndex        =   2
      Top             =   555
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "路单号(&K):"
      Height          =   180
      Left            =   6105
      TabIndex        =   0
      Top             =   150
      Width           =   900
   End
   Begin VB.Menu pmnu_Select 
      Caption         =   "选择"
      Visible         =   0   'False
      Begin VB.Menu pmnu_AllSelect 
         Caption         =   "全选"
      End
      Begin VB.Menu pmnu_CancelSelect 
         Caption         =   "重选"
      End
   End
End
Attribute VB_Name = "frmSelSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public m_dtStartDate As Date
Public m_dtEndDate As Date
Public m_szSheetID As String

'Const cnCols = 6


Const cnSheetID = 0 '路单代码
Const cnDate = 1 '日期
Const cnBusID = 2 '车次
Const cnBusSerialNO = 3 '序号
Const cnCompany = 4 '参运公司
Const cnVehicle = 5 '车辆
Const cnValidMark = 6 '有效标志
Const cnStatus = 7 '有效标志



Private Sub cmdCancel_Click()

    Unload Me
    
End Sub

Private Sub cmdFind_Click()
    '列出路单
    Dim rsTemp As Recordset
    Dim i As Integer
    Dim oReport As New Report
    Dim pnSerial As Integer
    If Not IsNumeric(txtBusSerialNO.Text) Then
        pnSerial = -1
    Else
        pnSerial = txtBusSerialNO.Text
    End If
    oReport.Init g_oActiveUser
    Set rsTemp = oReport.GetAllCheckSheetRS(txtSheetID.Text, txtBusID.Text, ResolveDisplay(txtVehicleID.Text), dtpStartDate.Value, DateAdd("d", 1, dtpEndDate.Value), ECheckSheetValidMark.CS_CheckSheetValid, pnSerial)
    
    
    
    lvSheet.ListItems.Clear
    
    For i = 1 To rsTemp.RecordCount
        lvSheet.ListItems.Add
        lvSheet.ListItems(i).Text = FormatDbValue(rsTemp!check_sheet_id)
        lvSheet.ListItems(i).SubItems(cnDate) = ToDBDate(FormatDbValue(rsTemp!bus_date))
        lvSheet.ListItems(i).SubItems(cnBusID) = FormatDbValue(rsTemp!bus_id)
        lvSheet.ListItems(i).SubItems(cnBusSerialNO) = FormatDbValue(rsTemp!bus_serial_no)
        lvSheet.ListItems(i).SubItems(cnCompany) = FormatDbValue(rsTemp!transport_company_name)
        lvSheet.ListItems(i).SubItems(cnVehicle) = FormatDbValue(rsTemp!license_tag_no)
        lvSheet.ListItems(i).SubItems(cnValidMark) = FormatDbValue(rsTemp!valid_mark_name)
        lvSheet.ListItems(i).SubItems(cnStatus) = FormatDbValue(rsTemp!settlement_status_name)
        rsTemp.MoveNext

    Next i
    
    If lvSheet.ListItems.Count > 0 Then
        lvSheet.ListItems(1).Selected = True
    End If
    
End Sub

Private Sub cmdSel_Click()
    '选择lvSheet中选择的路单
    SetSel
    Unload Me
    
End Sub

Private Sub Form_Load()
    InitForm
    AlignHeadWidth Me.name, lvSheet
End Sub

Private Sub Form_Unload(Cancel As Integer)

    SaveHeadWidth Me.name, lvSheet
    
End Sub

Private Sub lvSheet_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    SortListView lvSheet, ColumnHeader.Index
End Sub

Private Sub lvSheet_DblClick()
    Dim oCommDialog As New CommDialog
    If lvSheet.SelectedItem Is Nothing Then Exit Sub
    
    oCommDialog.Init g_oActiveUser
    oCommDialog.ShowCheckSheet lvSheet.SelectedItem.Text
    
End Sub

Private Sub lvSheet_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu pmnu_Select
    End If
End Sub

Private Sub pmnu_AllSelect_Click()
    Dim i As Integer
    For i = 1 To lvSheet.ListItems.Count
        lvSheet.ListItems(i).Selected = True
    Next i
End Sub

Private Sub pmnu_CancelSelect_Click()
    Dim i As Integer
    For i = 1 To lvSheet.ListItems.Count
        lvSheet.ListItems(i).Selected = False
    Next i
End Sub

Private Sub txtBusID_Click()
    '选择车次
    Dim oCommDialog As New CommDialog
    Dim aszTemp() As String
    
    oCommDialog.Init g_oActiveUser
    aszTemp = oCommDialog.SelectBus()
    If ArrayLength(aszTemp) > 0 Then
        txtBusID.Text = aszTemp(1, 1)
        
    End If
    
    
End Sub


Private Sub txtVehicleID_Click()
    '选择车辆
    Dim oCommDialog As New CommDialog
    Dim aszTemp() As String
    oCommDialog.Init g_oActiveUser
    aszTemp = oCommDialog.SelectVehicleEX()
    If ArrayLength(aszTemp) > 0 Then
        txtVehicleID.Text = MakeDisplayString(aszTemp(1, 1), aszTemp(1, 2))
        
    End If
End Sub


Private Sub InitForm()
    '初始化窗口
    dtpStartDate.Value = m_dtStartDate
    dtpEndDate.Value = m_dtEndDate
    
    '初始化列表
    lvSheet.ColumnHeaders.Clear
    lvSheet.ColumnHeaders.Add , , "路单号"
    lvSheet.ColumnHeaders.Add , , "日期"
    lvSheet.ColumnHeaders.Add , , "车次"
    lvSheet.ColumnHeaders.Add , , "序号"
    lvSheet.ColumnHeaders.Add , , "参运公司"
    lvSheet.ColumnHeaders.Add , , "车辆"
    lvSheet.ColumnHeaders.Add , , "有效标志"
    lvSheet.ColumnHeaders.Add , , "结算标志"
    
    
'    cboValidMark.Clear
'    cboValidMark.AddItem MakeDisplayString("-1", "全部")
'    cboValidMark.AddItem MakeDisplayString(ECheckSheetValidMark.CS_CheckSheetValid, "有效")
'    cboValidMark.AddItem MakeDisplayString(ECheckSheetValidMark.CS_CheckSheetInvalid, "作废")
    
    
    
End Sub

Private Sub SetSel()
    '将选中的路单放到m_aszSheet中
    Dim i As Integer
    Dim j As Integer
    
    m_szSheetID = ""
    If lvSheet.ListItems.Count = 0 Then Exit Sub
    If lvSheet.SelectedItem Is Nothing Then Exit Sub
    For i = 1 To lvSheet.ListItems.Count
        If lvSheet.ListItems(i).Selected Then
            
            m_szSheetID = m_szSheetID & lvSheet.ListItems(i).Text & ","
        End If
    Next i
    If Len(m_szSheetID) > 0 Then m_szSheetID = Left(m_szSheetID, Len(m_szSheetID) - 1)
    
    
End Sub

