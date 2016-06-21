VERSION 5.00
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmREBus 
   BackColor       =   &H00E0E0E0&
   Caption         =   "环境管理"
   ClientHeight    =   6090
   ClientLeft      =   3480
   ClientTop       =   1770
   ClientWidth     =   7725
   HelpContextID   =   2005001
   Icon            =   "frmREBus.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11475
   ScaleWidth      =   23880
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cboStation 
      Height          =   300
      ItemData        =   "frmREBus.frx":030A
      Left            =   8946
      List            =   "frmREBus.frx":0311
      TabIndex        =   11
      Text            =   "(全部)"
      Top             =   45
      Width           =   1380
   End
   Begin VB.TextBox txtBusID 
      Height          =   315
      Left            =   5090
      MaxLength       =   5
      TabIndex        =   1
      Top             =   38
      Width           =   735
   End
   Begin RTComctl3.CoolButton cmdFind 
      Height          =   300
      Left            =   10365
      TabIndex        =   0
      Top             =   45
      Width           =   1020
      _ExtentX        =   0
      _ExtentY        =   0
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
      MICON           =   "frmREBus.frx":031D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5955
      Top             =   870
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmREBus.frx":0339
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmREBus.frx":0495
            Key             =   "FlowRun"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmREBus.frx":05F1
            Key             =   "FlowStop"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmREBus.frx":074D
            Key             =   "Merge"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmREBus.frx":08A9
            Key             =   "Replace"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmREBus.frx":0A05
            Key             =   "Checking"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmREBus.frx":0B61
            Key             =   "Checked"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmREBus.frx":0CBD
            Key             =   "ExCheck"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmREBus.frx":0E19
            Key             =   "Run"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvBus 
      Height          =   4890
      Left            =   135
      TabIndex        =   2
      Top             =   810
      Width           =   8400
      _ExtentX        =   14817
      _ExtentY        =   8625
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "车次代码"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "日期"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "发车时间"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "运行线路"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "检票口"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "车牌"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "车型"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "全额退票"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "状态"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "终到站"
         Object.Width           =   2540
      EndProperty
   End
   Begin RTComctl3.TextButtonBox txtRoute 
      Height          =   315
      Left            =   6613
      TabIndex        =   3
      Top             =   45
      Width           =   1380
      _ExtentX        =   2434
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
      Text            =   "(全部)"
   End
   Begin MSComCtl2.DTPicker dtpStartDate 
      Height          =   300
      Left            =   799
      TabIndex        =   4
      Top             =   45
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   529
      _Version        =   393216
      Format          =   22872064
      CurrentDate     =   36396
   End
   Begin MSComCtl2.DTPicker dtpEndDate 
      Height          =   300
      Left            =   2682
      TabIndex        =   8
      Top             =   45
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   529
      _Version        =   393216
      Format          =   22872064
      CurrentDate     =   36396
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "途经站(&D):"
      Height          =   180
      Left            =   8027
      TabIndex        =   10
      Top             =   105
      Width           =   885
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "->"
      Height          =   225
      Left            =   2453
      TabIndex        =   9
      Top             =   90
      Width           =   195
   End
   Begin VB.Label lblInputBusId 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车次(&C):"
      Height          =   180
      Left            =   4336
      TabIndex        =   7
      Top             =   105
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "线路(&R):"
      Height          =   180
      Left            =   5859
      TabIndex        =   6
      Top             =   105
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "日期(&D):"
      Height          =   180
      Left            =   45
      TabIndex        =   5
      Top             =   105
      Width           =   720
   End
End
Attribute VB_Name = "frmREBus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'**********************************************************
'* Source File Name:frmREBus.frm
'* Project Name:StationNet 2.0
'* Engineer:
'* Data Generated:2005/11/27
'* Last Revision Date:
'* Brief Description:
'* Relational Document:
'**********************************************************
Public m_szBusID As String
Public m_dtBus As Date
Public bIsShow As Boolean
Public bIsActive As Boolean
Private m_oStation As New Station
Private m_oREBus As New REBus
Private m_oREScheme As New REScheme
Dim WithEvents m_oMsg As MsgNotify
Attribute m_oMsg.VB_VarHelpID = -1
Private m_nIndex  As Integer
Private Type CopyBus
    BusID As String
    BusDate As Date
    StartupTime As String
    RouteName As String
    CheckGate As String
    BusStatus As String
    AllReturn As String
    DestStation As String
    VehicleType As String
    VehicleTag As String
    BusSmallIcon As String
    Tag As Variant
End Type
Private m_tCopyBus As CopyBus
Const cnRightNap = 50

Private Sub cboStation_KeyDown(KeyCode As Integer, Shift As Integer)
Static nIndex As Integer
On Error GoTo here
    If KeyCode = vbKeyReturn Then
        If Val(GetLString(cboStation.Text)) > 0 Then
        m_oStation.Identify GetLString(cboStation.Text)
        Else
        m_oStation.Identify , GetLString(cboStation.Text)
        End If
        cboStation.Text = m_oStation.StationID & "[" & m_oStation.StationName & "]"
        If Trim(cboStation.Text) <> "" Then
            For i = 1 To cboStation.ListCount
            If Trim(cboStation.Text) = GetLString(cboStation.List(i)) Then
                cmdFind_Click
                Exit Sub
            End If
            Next
            If cboStation.ListCount >= 10 Then
            cboStation.List(nIndex + 1) = m_oStation.StationID & "[" & m_oStation.StationName & "]"
            cboStation.Text = cboStation.List(nIndex + 1)
            nIndex = nIndex + 1
            If nIndex >= 9 Then
            nIndex = 0
            End If
            cboStation.Text = m_oStation.StationID & "[" & m_oStation.StationName & "]"
            cmdFind_Click
            Else
            cboStation.AddItem m_oStation.StationID & "[" & m_oStation.StationName & "]"
            cboStation.Text = m_oStation.StationID & "[" & m_oStation.StationName & "]"
            cmdFind_Click
            End If
        End If
    End If
Exit Sub
here:
    ShowErrorMsg
End Sub

Private Sub cmdFind_Click()
Dim szaBus() As String
Dim ltTemp As ListItem
Dim szDestStation As String, szRoute As String
Dim nCount As Integer, eBusStatus As EREBusStatus
Dim i As Integer, nDay As Integer, j As Integer
Dim vbMsg As VbMsgBoxResult
Dim dtDay As Date
On Error GoTo here

nDay = DateDiff("d", dtpStartDate.Value, dtpEndDate.Value)
szDestStation = GetLString(cboStation.Text)
szRoute = IIf(GetLString(txtRoute.Text) = "(全部)", "", GetLString(txtRoute.Text))
szDestStation = IIf(GetLString(szDestStation) = "(全部)", "", GetLString(szDestStation))

If nDay > 1 And txtBusID.Text = "" And szDestStation = "" And szRoute = "" Then
    vbMsg = MsgBox("是否查询[" & nDay & "]天内的车次", vbQuestion + vbYesNo + vbDefaultButton2, "环境")
    If vbMsg = vbNo Then Exit Sub
End If

Me.MousePointer = vbHourglass
lvBus.ListItems.Clear

For j = 0 To nDay
dtDay = DateAdd("d", j, dtpStartDate.Value)
'ShowTBInfo "获得" & Format(dtDay, "YYYY-MM-DD") & "的车次"
szaBus = m_oREScheme.GetBus(dtDay, txtBusID.Text, szRoute, szDestStation)
nCount = ArrayLength(szaBus)
If nCount <> 0 Then
    'ShowTBInfo , nCount, , True
End If
For i = 1 To nCount
    'ShowTBInfo "获得" & szaBus(i, 1) & "的车次信息", , i
    If Val(szaBus(i, 8)) <> TP_ScrollBus Then  '是否固定班次
        Set ltTemp = lvBus.ListItems.Add(, , szaBus(i, 1), , "Run")
        ltTemp.ListSubItems.Add , , Format(dtDay, "YYYY-MM-DD")
        ltTemp.ListSubItems.Add , , Format(szaBus(i, 2), "HH:MM:SS")
    Else
        Set ltTemp = lvBus.ListItems.Add(, , szaBus(i, 1), , "FlowRun")
        ltTemp.ListSubItems.Add , , Format(dtDay, "YYYY-MM-DD")
        ltTemp.ListSubItems.Add , , Format(szaBus(i, 2), "HH:MM:SS")
    End If
    ltTemp.ListSubItems.Add , , Trim(szaBus(i, 3))
    ltTemp.ListSubItems.Add , , szaBus(i, 4)
    ltTemp.ListSubItems.Add , , szaBus(i, 5)
    ltTemp.ListSubItems.Add , , szaBus(i, 6)
    If Val(szaBus(i, 9)) = 0 Then
        ltTemp.ListSubItems.Add , , "否"
    Else
        ltTemp.ListSubItems.Add , , "是"
        ltTemp.ListSubItems.Item(7).ForeColor = vbRed
    End If
    eBusStatus = Val(szaBus(i, 7))
    If eBusStatus = ST_BusStopped Or eBusStatus = ST_BusSlitpStop Or eBusStatus = ST_BusReplace Then
        ltTemp.Tag = "STOP"
        ltTemp.ListSubItems.Add , , "停班"
        ltTemp.ListSubItems.Item(8).ForeColor = vbRed
        ltTemp.ListSubItems.Item(8).ReportIcon = vbEmpty
        If Val(szaBus(i, 8)) <> TP_ScrollBus Then '固定班次
            Select Case Val(szaBus(i, 7))
                   Case ST_BusStopped: ltTemp.SmallIcon = "Stop"
                   Case ST_BusMergeStopped: ltTemp.SmallIcon = "Merge"
                   Case ST_BusReplace: ltTemp.SmallIcon = "Replace"
            End Select
        Else
            ltTemp.SmallIcon = "FlowStop"
        End If
    Else
        ltTemp.ListSubItems.Add , , "运行"
    End If
    Select Case eBusStatus
            Case ST_BusChecking
            ltTemp.SmallIcon = "Checking"
            Case ST_BusExtraChecking
            ltTemp.SmallIcon = "ExCheck"
            Case ST_BusStopCheck
            ltTemp.SmallIcon = "Checked"
    End Select
    ltTemp.ListSubItems.Add , , szaBus(i, 10)
Next
Next
    m_szBusID = ""
    'ShowTBInfo
    Me.MousePointer = vbDefault
Exit Sub
here:
    Me.MousePointer = vbDefault
    ShowErrorMsg
End Sub

Private Sub Form_Activate()
    bIsActive = True
    MDIBook.mnu_REBusPro.Enabled = True
    MDIBook.mnu_REBusSeat.Enabled = True
    MDIBook.Toolbar1.Buttons("Seat").Enabled = True
    MDIBook.Toolbar1.Buttons("REBus").Enabled = True
End Sub

Private Sub Form_Deactivate()
    bIsActive = False
    MDIBook.mnu_REBusPro.Enabled = False
    MDIBook.mnu_REBusSeat.Enabled = False
    MDIBook.Toolbar1.Buttons("Seat").Enabled = False
    MDIBook.Toolbar1.Buttons("REBus").Enabled = False
End Sub

Private Sub Form_Resize()
    Dim lTemp As Long
    On Error Resume Next
    lTemp = Me.ScaleHeight - 400
    lTemp = IIf(lTemp > 0, lTemp, 0)
    lvBus.Move 0, 500, Me.ScaleWidth - 50, lTemp
End Sub

Private Sub Form_Unload(Cancel As Integer)
    bIsActive = False
    bIsShow = False
    MDIBook.mnu_REBusPro.Enabled = False
    MDIBook.mnu_REBusSeat.Enabled = False
    MDIBook.Toolbar1.Buttons("Seat").Enabled = False
    MDIBook.Toolbar1.Buttons("REBus").Enabled = False
End Sub
Private Sub Form_Load()
'    Set m_oMsg = New MsgNotify
'    m_oMsg.Unit = g_szUnitID
'    bIsShow = True
'    Set MDIBook.CellExport1.ListViewSource = lvBus
'    m_oStation.Init m_oActiveUser
'    m_oREBus.Init m_oActiveUser
'    m_oREScheme.Init m_oActiveUser
'    dtpStartDate.Value = Date
'    dtpEndDate.Value = Date
'    MDIBook.mnu_REBusPro.Enabled = True
'    MDIBook.mnu_REBusSeat.Enabled = True
'    MDIBook.Toolbar1.Buttons("Seat").Enabled = True
'    MDIBook.Toolbar1.Buttons("REBus").Enabled = True

    dtpStartDate.Value = Date
    dtpEndDate.Value = Date
End Sub

Private Sub lvBus_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Static m_nUpColumn As Integer
lvBus.SortKey = ColumnHeader.Index - 1
If m_nUpColumn = ColumnHeader.Index - 1 Then
    lvBus.SortOrder = lvwDescending
    m_nUpColumn = ColumnHeader.Index
Else
    lvBus.SortOrder = lvwAscending
    m_nUpColumn = ColumnHeader.Index - 1
End If
lvBus.Sorted = True
End Sub

Public Sub lvBus_DblClick()
    MDIBook.mnu_REBusPro_Click
End Sub

Private Sub lvBus_ItemClick(ByVal Item As MSComctlLib.ListItem)
    m_szBusID = Item.Text
    m_dtBus = CDate(Item.ListSubItems(1).Text)
End Sub

Private Sub lvBus_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        PopupMenu MDIBook.mnu_popREBus
    End If
End Sub

Private Sub m_oMsg_ExStartCheckBus(ByVal szBusID As String, ByVal dtBusDate As Date, ByVal nBusSerialNo As Integer)
    On Error Resume Next
    UpList szBusID, dtBusDate, True
End Sub

Private Sub m_oMsg_StartCheckBus(ByVal szBusID As String, ByVal dtBusDate As Date, ByVal nBusSerialNo As Integer)
    On Error Resume Next
    UpList szBusID, dtBusDate, True
End Sub

Private Sub m_oMsg_StopCheckBus(ByVal szBusID As String, ByVal dtBusDate As Date, ByVal nBusSerialNo As Integer)
    On Error Resume Next
    UpList szBusID, dtBusDate, True
End Sub

Private Sub txtBusID_GotFocus()
    txtBusID.SelStart = 0
    txtBusID.SelLength = 100
End Sub

Private Sub txtBusID_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
       Case vbKeyReturn
       cmdFind_Click
End Select
End Sub

Private Sub txtRoute_Change()
Select Case KeyAscii
       Case vbKeyReturn
       cmdFind_Click
End Select
End Sub

Private Sub txtRoute_Click()
    Dim oShell As New CommDialog
    Dim szaTemp() As String
    oShell.Init m_oActiveUser
    szaTemp = oShell.SelectRoute(False)
    Set oShell = Nothing
    If ArrayLength(szaTemp) = 0 Then Exit Sub
    txtRoute.Text = szaTemp(1, 1) & "[" & szaTemp(1, 2) & "]"
End Sub



Private Sub txtRoute_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
       Case vbKeyReturn
       cmdFind_Click
End Select
End Sub

Public Sub AddList(BusID As String, RunDate As Date)
Dim szaBus() As String
Dim ltTemp As ListItem
Dim szDestStation As String
Dim nCount As String
Dim i As Integer, nDay As Integer, j As Integer
Dim dtDay As Date
Me.MousePointer = vbHourglass
On Error GoTo here
dtDay = RunDate
szaBus = m_oREScheme.GetBus(RunDate, BusID)
nCount = ArrayLength(szaBus)
For i = 1 To nCount
    If Val(szaBus(i, 8)) = TP_RegularBus Then '是否固定班次
        Set ltTemp = lvBus.ListItems.Add(, , szaBus(i, 1), , "Run")
        ltTemp.ListSubItems.Add , , Format(dtDay, "YYYY-MM-DD")
        ltTemp.ListSubItems.Add , , Format(szaBus(i, 2), "HH:MM:SS")
    Else
        Set ltTemp = lvBus.ListItems.Add(, , szaBus(i, 1), , "FlowRun")
        ltTemp.ListSubItems.Add , , Format(dtDay, "YYYY-MM-DD")
        ltTemp.ListSubItems.Add , , "流水车次"
    End If
    ltTemp.ListSubItems.Add , , Trim(szaBus(i, 3))
    ltTemp.ListSubItems.Add , , szaBus(i, 4)
    ltTemp.ListSubItems.Add , , szaBus(i, 5)
    ltTemp.ListSubItems.Add , , szaBus(i, 6)
    If Val(szaBus(i, 9)) = 0 Then
        ltTemp.ListSubItems.Add , , "否"
    Else
        ltTemp.ListSubItems.Add , , "是"
        ltTemp.ListSubItems.Item(7).ForeColor = vbRed
    End If
    If Val(szaBus(i, 7)) = ST_BusStopped Or Val(szaBus(i, 7)) = ST_BusMergeStopped Then
        ltTemp.Tag = "STOP"
        ltTemp.ListSubItems.Add , , "停班"
        ltTemp.ListSubItems.Item(8).ForeColor = vbRed
        ltTemp.ListSubItems.Item(8).ReportIcon = vbEmpty
        If Val(szaBus(i, 8)) = TP_RegularBus Then '固定班次
            Select Case Val(szaBus(i, 7))
                   Case ST_BusStopped: ltTemp.SmallIcon = "Stop"
                   Case ST_BusMergeStopped: ltTemp.SmallIcon = "Merge"
            End Select
        Else
            ltTemp.SmallIcon = "FlowStop"
        End If
        
    Else
        ltTemp.ListSubItems.Add , , "运行"
    End If
        ltTemp.ListSubItems.Add , , szaBus(i, 10)
Next
    m_szBusID = ""
    Me.MousePointer = vbDefault
Exit Sub
here:
    Me.MousePointer = vbDefault
    ShowErrorMsg
End Sub
Public Sub GetReport()
    ceReport.SourceSelect = ListViewControl
End Sub

Public Sub UpList(BusID As String, dtDay As Date, Optional RefreshStatus As Boolean = False)
Dim i As Integer, j As Integer
Dim ltTemp As ListItem
Dim bFind As Boolean
Dim szaBus() As String
On Error GoTo here
    szaBus = m_oREScheme.GetBus(dtDay, BusID)
    nCount = ArrayLength(szaBus)
    'ShowTBInfo "获得" & BusID & "的车次信息"
    bFind = False
    For i = 1 To lvBus.ListItems.Count
        Set ltTemp = lvBus.ListItems(i)
        If Trim(ltTemp.Text) = Trim(BusID) And (DateDiff("d", CDate(ltTemp.ListSubItems.Item(1).Text), dtDay) = 0) Then
            bFind = True
            Exit For
        End If
    Next
    If bFind = False Then Exit Sub
    If Val(szaBus(1, 8)) = TP_ScrollBus And RefreshStatus = True Then Exit Sub
    If RefreshStatus = False Then
        ltTemp.EnsureVisible
        ltTemp.Selected = True
    End If
    If Val(szaBus(1, 8)) <> TP_ScrollBus Then '是否固定班次
        ltTemp.SmallIcon = "Run"
        ltTemp.ListSubItems.Item(1).Text = Format(dtDay, "YYYY-MM-DD")
        ltTemp.ListSubItems.Item(2).Text = Format(szaBus(1, 2), "HH:MM:SS")
    Else
        ltTemp.SmallIcon = "FlowRun"
        ltTemp.ListSubItems.Item(1).Text = Format(dtDay, "YYYY-MM-DD")
        ltTemp.ListSubItems.Item(2).Text = Format(szaBus(1, 2), "HH:MM:SS")
    End If
    ltTemp.ListSubItems.Item(3).Text = Trim(szaBus(1, 3))
    ltTemp.ListSubItems.Item(4).Text = szaBus(1, 4)
    ltTemp.ListSubItems.Item(5).Text = szaBus(1, 5)
    ltTemp.ListSubItems.Item(6).Text = szaBus(1, 6)
    If Val(szaBus(1, 9)) = 0 Then
        ltTemp.ListSubItems.Item(7).Text = "否"
        ltTemp.ListSubItems.Item(7).ForeColor = vbBlack
    Else
        ltTemp.ListSubItems.Item(7).Text = "是"
        ltTemp.ListSubItems.Item(7).ForeColor = vbRed
    End If
    If Val(szaBus(1, 7)) = ST_BusStopped Or Val(szaBus(1, 7)) = ST_BusReplace Or Val(szaBus(1, 7)) = ST_BusSlitpStop Then
        ltTemp.Tag = "STOP"
        ltTemp.ListSubItems.Item(8).Text = "停班"
        ltTemp.ListSubItems.Item(8).ForeColor = vbRed
        ltTemp.ListSubItems.Item(8).ReportIcon = vbEmpty
        If Val(szaBus(1, 8)) <> TP_ScrollBus Then '固定班次
            Select Case Val(szaBus(1, 7))
                   Case ST_BusStopped: ltTemp.SmallIcon = "Stop"
                   Case ST_BusSlitpStop: ltTemp.SmallIcon = "Merge"
                   Case ST_BusReplace: ltTemp.SmallIcon = "Replace"
            End Select
        Else
            ltTemp.SmallIcon = "FlowStop"
        End If
    Else
        ltTemp.ListSubItems.Item(8).Text = "运行"
        ltTemp.ListSubItems.Item(8).ForeColor = vbBlack
    End If
        ltTemp.ListSubItems.Item(9).Text = szaBus(1, 10)
    eBusStatus = Val(szaBus(1, 7))
    Select Case eBusStatus
            Case ST_BusChecking
            ltTemp.SmallIcon = "Checking"
            Case ST_BusExtraChecking
            ltTemp.SmallIcon = "ExCheck"
            Case ST_BusStopCheck
            ltTemp.SmallIcon = "Checked"
    End Select
    'ShowTBInfo
Exit Sub
here:
    ShowErrorMsg
End Sub
