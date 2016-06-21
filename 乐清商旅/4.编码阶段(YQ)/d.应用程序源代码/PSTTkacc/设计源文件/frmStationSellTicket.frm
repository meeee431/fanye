VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Begin VB.Form frmStationSellTicket 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "车次站点售票"
   ClientHeight    =   6210
   ClientLeft      =   1800
   ClientTop       =   1905
   ClientWidth     =   8205
   HelpContextID   =   6000050
   Icon            =   "frmStationSellTicket.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   8205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin RTComctl3.CoolButton CoolButton1 
      Height          =   330
      Left            =   6900
      TabIndex        =   22
      Top             =   1470
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   582
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
      MICON           =   "frmStationSellTicket.frx":030A
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
      Height          =   330
      Left            =   6870
      TabIndex        =   16
      Top             =   180
      Width           =   1230
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
      MICON           =   "frmStationSellTicket.frx":0326
      PICN            =   "frmStationSellTicket.frx":0342
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
      BackColor       =   &H00E0E0E0&
      Height          =   1155
      Left            =   90
      TabIndex        =   20
      Top             =   1110
      Width           =   6585
      Begin VB.ComboBox cboStation 
         Height          =   300
         ItemData        =   "frmStationSellTicket.frx":06DC
         Left            =   4410
         List            =   "frmStationSellTicket.frx":06E3
         TabIndex        =   5
         Text            =   "(全部)"
         Top             =   660
         Width           =   2010
      End
      Begin VB.TextBox txtBusID 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   5490
         MaxLength       =   5
         TabIndex        =   1
         Top             =   210
         Width           =   915
      End
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   300
         Left            =   990
         TabIndex        =   7
         Top             =   210
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   529
         _Version        =   393216
         Format          =   67305472
         CurrentDate     =   36396
      End
      Begin MSComCtl2.DTPicker dtpEndDate 
         Height          =   300
         Left            =   2910
         TabIndex        =   8
         Top             =   210
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   529
         _Version        =   393216
         Format          =   67305472
         CurrentDate     =   36396
      End
      Begin FText.asFlatTextBox txtRoute 
         Height          =   300
         Left            =   990
         TabIndex        =   3
         Top             =   660
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   529
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
         ButtonVisible   =   -1  'True
         Registered      =   -1  'True
         OfficeXPColors  =   -1  'True
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "途经站(&D):"
         Height          =   180
         Left            =   3360
         TabIndex        =   4
         Top             =   720
         Width           =   885
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "->"
         Height          =   225
         Left            =   2670
         TabIndex        =   21
         Top             =   255
         Width           =   195
      End
      Begin VB.Label lblInputBusId 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车次(&C):"
         Height          =   180
         Left            =   4650
         TabIndex        =   0
         Top             =   270
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "日期(&D):"
         Height          =   180
         Left            =   150
         TabIndex        =   6
         Top             =   270
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "线路(&R):"
         Height          =   180
         Left            =   150
         TabIndex        =   2
         Top             =   720
         Width           =   720
      End
   End
   Begin VB.Frame FraCountMode 
      BackColor       =   &H00E0E0E0&
      Caption         =   "统计方式："
      Height          =   1005
      Left            =   90
      TabIndex        =   19
      Top             =   30
      Width           =   6585
      Begin VB.ComboBox CboCount 
         Height          =   300
         ItemData        =   "frmStationSellTicket.frx":06EF
         Left            =   2910
         List            =   "frmStationSellTicket.frx":06F1
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   570
         Width           =   3405
      End
      Begin VB.ComboBox CboOrder 
         Height          =   300
         ItemData        =   "frmStationSellTicket.frx":06F3
         Left            =   2910
         List            =   "frmStationSellTicket.frx":06F5
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   210
         Width           =   3405
      End
      Begin VB.OptionButton OptCount 
         BackColor       =   &H00E0E0E0&
         Caption         =   "统计求和"
         Height          =   300
         Left            =   480
         TabIndex        =   13
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton OptList 
         BackColor       =   &H00E0E0E0&
         Caption         =   "明细列表"
         Height          =   300
         Left            =   480
         TabIndex        =   10
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.Label lblCount 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "统计方式："
         Height          =   180
         Left            =   1830
         TabIndex        =   14
         Top             =   630
         Width           =   900
      End
      Begin VB.Label lblOrder 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "排序方式："
         Height          =   180
         Left            =   1830
         TabIndex        =   11
         Top             =   270
         Width           =   900
      End
   End
   Begin RTComctl3.CoolButton CmdCancel 
      Cancel          =   -1  'True
      Height          =   330
      Left            =   6870
      TabIndex        =   18
      Top             =   1020
      Width           =   1230
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
      MICON           =   "frmStationSellTicket.frx":06F7
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton CmdOK 
      Height          =   330
      Left            =   6870
      TabIndex        =   17
      Top             =   600
      Width           =   1230
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "确定(&O)"
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
      MICON           =   "frmStationSellTicket.frx":0713
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
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStationSellTicket.frx":072F
            Key             =   "Run"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStationSellTicket.frx":0889
            Key             =   "FlowRun"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvBus 
      Height          =   3750
      Left            =   90
      TabIndex        =   9
      Top             =   2340
      Width           =   8040
      _ExtentX        =   14182
      _ExtentY        =   6615
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
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "车次"
         Object.Width           =   1589
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "日期"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "时间"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "运行线路"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "车型"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "终到站"
         Object.Width           =   1270
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "座位"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "参运公司"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmStationSellTicket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_oREScheme As New REScheme

Const cnRightNap = 50
Public m_bOk As Boolean
Public m_szBus_Id As Variant
Public m_dtBus_Date As Variant
Public m_nOrder As Integer
Public m_bnList As Boolean
Public m_nCount As Integer
Public m_dtStartDate As Date
Public m_dtEndDate As Date

Private Sub CboCount_Click()
    m_nCount = CboCount.ListIndex
End Sub

Private Sub cmdCancel_Click()
    m_bOk = False
    Unload Me
End Sub

Private Sub cmdFind_Click()
Dim szaBus() As String
Dim ltTemp As ListItem
Dim szDestStation As String, szRoute As String
Dim nCount As Integer, eBusStatus As EREBusStatus
Dim i As Integer, nDay As Integer, j As Integer
Dim dtDay As Date
On Error GoTo Here

nDay = DateDiff("d", dtpStartDate.Value, dtpEndDate.Value)
szDestStation = GetLString(cboStation.Text)
szRoute = IIf(GetLString(txtRoute.Text) = "(全部)", "", GetLString(txtRoute.Text))
szDestStation = IIf(GetLString(szDestStation) = "(全部)", "", GetLString(szDestStation))

Me.MousePointer = vbHourglass
lvBus.ListItems.Clear

For j = 0 To nDay
dtDay = DateAdd("d", j, dtpStartDate.Value)
szaBus = m_oREScheme.GetBus(dtDay, txtBusID.Text, szRoute, szDestStation)
nCount = ArrayLength(szaBus)

For i = 1 To nCount
    If Val(szaBus(i, 8)) <> TP_ScrollBus Then  '是否固定班次
        Set ltTemp = lvBus.ListItems.Add(, , szaBus(i, 1), , "Run")
        ltTemp.ListSubItems.Add , , Format(dtDay, "YYYY-MM-DD")
        ltTemp.ListSubItems.Add , , Format(szaBus(i, 2), "HH:MM")
    Else
        Set ltTemp = lvBus.ListItems.Add(, , szaBus(i, 1), , "FlowRun")
        ltTemp.ListSubItems.Add , , Format(dtDay, "YYYY-MM-DD")
        ltTemp.ListSubItems.Add , , "流水车次"
    End If
    ltTemp.ListSubItems.Add , , Trim(szaBus(i, 3))
    ltTemp.ListSubItems.Add , , szaBus(i, 6)
    ltTemp.ListSubItems.Add , , Trim(szaBus(i, 10))
    ltTemp.ListSubItems.Add , , szaBus(i, 11)
     ltTemp.ListSubItems.Add , , szaBus(i, 12)
Next
Next
    Me.MousePointer = vbDefault
    lvBus.SetFocus
Exit Sub
Here:
    Me.MousePointer = vbDefault
    
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdok_Click()
Dim nSelectRow As Integer
Dim i As Integer
Dim szbusID() As String
Dim dtBusDate() As Date

On Error GoTo errorHander
m_dtStartDate = dtpStartDate.Value
m_dtEndDate = dtpEndDate.Value

If OptList.Value = True Then
   m_bnList = True
Else
   m_bnList = False
End If
For i = 1 To lvBus.ListItems.Count
    If lvBus.ListItems(i).Selected = True Then nSelectRow = nSelectRow + 1
Next i
ReDim szbusID(1 To nSelectRow)
ReDim dtBusDate(1 To nSelectRow)
nSelectRow = 0
For i = 1 To lvBus.ListItems.Count
    If lvBus.ListItems(i).Selected = True Then
       nSelectRow = nSelectRow + 1
       szbusID(nSelectRow) = lvBus.ListItems(i).Text
       dtBusDate(nSelectRow) = lvBus.ListItems(i).SubItems(1)
    End If
Next i
m_szBus_Id = szbusID
m_dtBus_Date = dtBusDate
m_nOrder = CboOrder.ListIndex
m_bOk = True
Unload Me
Exit Sub

errorHander:
End Sub

Private Sub CoolButton1_Click()
DisplayHelp Me
End Sub

Private Sub Form_Load()
'    Set ceReport.ListViewSource = lvBus

    dtpStartDate.Value = DateAdd("d", -1, m_oParam.NowDate)
    dtpEndDate.Value = DateAdd("d", -1, m_oParam.NowDate)
    m_bOk = False
    CboOrder.AddItem "票  号"
    CboOrder.AddItem "站  点"
    CboOrder.AddItem "座位号"
    CboOrder.ListIndex = 0
    CboCount.AddItem "车    次"
    CboCount.AddItem "站    点"
    CboCount.AddItem "车次站点"
    CboCount.ListIndex = 0
    CboCount.Enabled = False
    lblCount.Enabled = False
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

Private Sub lvBus_DblClick()
    cmdok_Click
End Sub

Private Sub OptCount_Click()
    CboOrder.Enabled = False
    lblOrder.Enabled = False
    CboCount.Enabled = True
    lblCount.Enabled = True
    m_bnList = False
End Sub

Private Sub OptList_Click()
    CboOrder.Enabled = True
    lblOrder.Enabled = True
    CboCount.Enabled = False
    lblCount.Enabled = False
    m_bnList = True
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

Private Sub txtRoute_ButtonClick()
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
Public Sub GetReport()
'    Set MDIMain.ceMain.ListViewSource = lvBus
'    MDIMain.ceMain.SourceSelect = 3
End Sub


