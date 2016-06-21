VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form FrmStationSellCount 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "站点售票查询"
   ClientHeight    =   5445
   ClientLeft      =   3165
   ClientTop       =   4365
   ClientWidth     =   7800
   HelpContextID   =   60000200
   Icon            =   "FrmStationSellCount.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   7800
   ShowInTaskbar   =   0   'False
   Begin RTComctl3.CoolButton CoolButton1 
      Height          =   345
      Left            =   6480
      TabIndex        =   25
      Top             =   2100
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
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
      MICON           =   "FrmStationSellCount.frx":014A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame FraCountMode 
      BackColor       =   &H00E0E0E0&
      Caption         =   "统计方式："
      Height          =   1095
      Left            =   3720
      TabIndex        =   24
      Top             =   90
      Width           =   3975
      Begin VB.OptionButton OptList 
         BackColor       =   &H00E0E0E0&
         Caption         =   "明细列表"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   300
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton OptCount 
         BackColor       =   &H00E0E0E0&
         Caption         =   "统计求和"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   660
         Width           =   1215
      End
      Begin VB.ComboBox CboOrder 
         Height          =   300
         ItemData        =   "FrmStationSellCount.frx":0166
         Left            =   2520
         List            =   "FrmStationSellCount.frx":0168
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   270
         Width           =   1215
      End
      Begin VB.ComboBox CboCount 
         Height          =   300
         ItemData        =   "FrmStationSellCount.frx":016A
         Left            =   2520
         List            =   "FrmStationSellCount.frx":016C
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   630
         Width           =   1215
      End
      Begin VB.Label lblOrder 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "排序方式："
         Height          =   180
         Left            =   1440
         TabIndex        =   18
         Top             =   330
         Width           =   900
      End
      Begin VB.Label lblCount 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "统计方式："
         Height          =   180
         Left            =   1440
         TabIndex        =   20
         Top             =   690
         Width           =   900
      End
   End
   Begin VB.Frame FraStation 
      BackColor       =   &H00E0E0E0&
      Caption         =   "站点选择："
      Height          =   1155
      Left            =   120
      TabIndex        =   23
      Top             =   1290
      Width           =   6225
      Begin RTComctl3.CoolButton cmdFind 
         Default         =   -1  'True
         Height          =   360
         Left            =   4860
         TabIndex        =   8
         Top             =   270
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   635
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
         MICON           =   "FrmStationSellCount.frx":016E
         PICN            =   "FrmStationSellCount.frx":018A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox txtInputStation 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   3600
         MaxLength       =   4
         TabIndex        =   3
         Top             =   285
         Width           =   1140
      End
      Begin VB.TextBox txtStationName 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1275
         TabIndex        =   5
         Top             =   705
         Width           =   1230
      End
      Begin VB.TextBox txtStationID 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1275
         TabIndex        =   1
         Top             =   285
         Width           =   1230
      End
      Begin FText.asFlatTextBox txtArea 
         Height          =   300
         Left            =   3600
         TabIndex        =   7
         Top             =   705
         Width           =   1140
         _ExtentX        =   2011
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
         Text            =   ""
         ButtonVisible   =   -1  'True
         Registered      =   -1  'True
         OfficeXPColors  =   -1  'True
      End
      Begin VB.Label lblInputRouteId 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "输入码(&M):"
         Height          =   180
         Left            =   2640
         TabIndex        =   2
         Top             =   345
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "地区(&D):"
         Height          =   180
         Left            =   2640
         TabIndex        =   6
         Top             =   765
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "站点名称(&N):"
         Height          =   180
         Left            =   180
         TabIndex        =   4
         Top             =   765
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "站点代码(&Z):"
         Height          =   180
         Left            =   180
         TabIndex        =   0
         Top             =   345
         Width           =   1080
      End
   End
   Begin VB.Frame FraStartEndDate 
      BackColor       =   &H00E0E0E0&
      Caption         =   "选择时间段："
      Height          =   1095
      Left            =   120
      TabIndex        =   22
      Top             =   90
      Width           =   3495
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   300
         Left            =   1440
         TabIndex        =   13
         Top             =   240
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   529
         _Version        =   393216
         Format          =   60030976
         CurrentDate     =   36396
      End
      Begin MSComCtl2.DTPicker dtpEndDate 
         Height          =   300
         Left            =   1440
         TabIndex        =   15
         Top             =   615
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   529
         _Version        =   393216
         Format          =   60030976
         CurrentDate     =   36396
      End
      Begin VB.Label lblEndDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "终止日期(&E)："
         Height          =   180
         Left            =   240
         TabIndex        =   14
         Top             =   720
         Width           =   1170
      End
      Begin VB.Label lblStartDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "起始日期(&S)："
         Height          =   180
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   1170
      End
   End
   Begin RTComctl3.CoolButton CmdOK 
      Height          =   345
      Left            =   6480
      TabIndex        =   10
      Top             =   1260
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
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
      MICON           =   "FrmStationSellCount.frx":0524
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton CmdCancel 
      Cancel          =   -1  'True
      Height          =   345
      Left            =   6480
      TabIndex        =   11
      Top             =   1680
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
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
      MICON           =   "FrmStationSellCount.frx":0540
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ImageList imgRoute 
      Left            =   585
      Top             =   3570
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
            Picture         =   "FrmStationSellCount.frx":055C
            Key             =   "Station"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmStationSellCount.frx":06B6
            Key             =   "NoSell"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvStation 
      Height          =   2775
      Left            =   120
      TabIndex        =   9
      Top             =   2550
      Width           =   7545
      _ExtentX        =   13309
      _ExtentY        =   4895
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imgRoute"
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "站点代码"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "站点名称"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "输入码"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "本地码"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "地区"
         Object.Width           =   3528
      EndProperty
   End
End
Attribute VB_Name = "FrmStationSellCount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_oBaseInfo As New BaseInfo '基本信息对象 BaseInfo
'Private m_oStation As New Station
'Public m_szStationID As String
Public m_bnList As Boolean
Public m_bOk As Boolean
Public m_szStationID As Variant
Public m_nOrder As Integer
Public m_nCount As Integer
Public m_dtStartDate As Date
Public m_dtEndDate As Date

Private Sub CboCount_Click()
    m_nCount = CboCount.ListIndex
End Sub

Private Sub CboOrder_Click()
  m_nOrder = CboOrder.ListIndex
End Sub

Private Sub cmdCancel_Click()
    m_bOk = False
    Unload Me
End Sub

Public Sub cmdFind_Click()
Dim szaStation() As String
Dim ltTemp As ListItem
Dim i As Integer, nCount As Integer
On Error GoTo Here
Me.MousePointer = vbHourglass
lvStation.ListItems.Clear
szaStation = m_oBaseInfo.GetStation(GetLString(txtArea.Text), txtStationName.Text, txtStationID.Text, txtInputStation.Text)
'--------------------------------------------------
nCount = ArrayLength(szaStation)
For i = 1 To nCount
    Set ltTemp = lvStation.ListItems.Add(, , szaStation(i, 1), , "Station")
    ltTemp.ListSubItems.Add , , szaStation(i, 2)
    ltTemp.ListSubItems.Add , , szaStation(i, 3)
    ltTemp.ListSubItems.Add , , szaStation(i, 5)
    ltTemp.ListSubItems.Add , , szaStation(i, 6)
Next
    Me.MousePointer = vbDefault
    lvStation.SetFocus
Exit Sub
Here:
    Me.MousePointer = vbDefault
End Sub

Private Sub cmdok_Click()
Dim nSelectRow As Integer
Dim i As Integer
Dim szStationId() As String

On Error GoTo errorHander
If OptList.Value = True Then
   m_bnList = True
Else
   m_bnList = False
End If
m_dtStartDate = dtpStartDate.Value
m_dtEndDate = dtpEndDate.Value
For i = 1 To lvStation.ListItems.Count
    If lvStation.ListItems(i).Selected = True Then nSelectRow = nSelectRow + 1
Next i
ReDim szStationId(1 To nSelectRow)
nSelectRow = 0
For i = 1 To lvStation.ListItems.Count
    If lvStation.ListItems(i).Selected = True Then
       nSelectRow = nSelectRow + 1
       szStationId(nSelectRow) = lvStation.ListItems(i).Text
    End If
Next i
m_szStationID = szStationId
m_bOk = True
Unload Me
Exit Sub

errorHander:
End Sub


Private Sub CoolButton1_Click()
DisplayHelp Me
End Sub

Private Sub lvStation_DblClick()
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

Private Sub Form_Load()
    m_oBaseInfo.Init m_oActiveUser
    'm_oStation.Init m_oActiveUser
    'Set ceReport.ListViewSource = lvStation

    dtpStartDate.Value = DateAdd("d", -1, m_oParam.NowDate)
    dtpEndDate.Value = DateAdd("d", -1, m_oParam.NowDate)
    
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

Private Sub txtRouteId_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
       Case vbKeyReturn
       cmdFind_Click
End Select
End Sub

Private Sub lvStation_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Static m_nUpColumn As Integer
lvStation.SortKey = ColumnHeader.Index - 1
If m_nUpColumn = ColumnHeader.Index - 1 Then
    lvStation.SortOrder = lvwDescending
    m_nUpColumn = ColumnHeader.Index
Else
    lvStation.SortOrder = lvwAscending
    m_nUpColumn = ColumnHeader.Index - 1
End If
lvStation.Sorted = True
End Sub

Private Sub txtArea_ButtonClick()
    Dim oShell As New CommDialog
    Dim szaTemp() As String
    oShell.Init m_oActiveUser
    szaTemp = oShell.SelectArea(False)
    Set oShell = Nothing
    If ArrayLength(szaTemp) = 0 Then Exit Sub
    txtArea.Text = szaTemp(1, 1) & "[" & szaTemp(1, 2) & "]"
End Sub

Private Sub txtStationID_GotFocus()
    txtStationID.SelStart = 0
    txtStationID.SelLength = 100
End Sub

Private Sub txtStationID_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
       Case vbKeyReturn
       cmdFind_Click
End Select
End Sub

Private Sub txtInputStation_GotFocus()
txtInputStation.SelStart = 0
txtInputStation.SelLength = 100
End Sub

Private Sub txtStationName_GotFocus()
    txtStationName.SelStart = 0
    txtStationName.SelLength = 100
End Sub
Public Sub GetReport()
'    Set MDIMain.ceMain.ListViewSource = lvStation
'    MDIMain.ceMain.SourceSelect = 3
End Sub

