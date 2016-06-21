VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Begin VB.Form frmSlitpInfo 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "车次拆分信息表"
   ClientHeight    =   6315
   ClientLeft      =   1080
   ClientTop       =   780
   ClientWidth     =   9105
   Icon            =   "frmSlitpInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   9105
   ShowInTaskbar   =   0   'False
   Begin RTComctl3.CoolButton cmdQuit 
      Height          =   375
      Left            =   6540
      TabIndex        =   3
      Top             =   5850
      Width           =   1005
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   2
      TX              =   ""
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
      MICON           =   "frmSlitpInfo.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ListView lvBus 
      Height          =   4605
      Left            =   240
      TabIndex        =   1
      Top             =   1170
      Width           =   8700
      _ExtentX        =   15346
      _ExtentY        =   8123
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imgSeat"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "目标车次"
         Object.Width           =   1589
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "被拆车次 "
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "现座号"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "原座号"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "票号"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "检票口"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "车 型"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "总痤位"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "终点站"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "发车时间"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Timer TimStart 
      Interval        =   250
      Left            =   3720
      Top             =   1680
   End
   Begin RTComctl3.CoolButton cmdexit 
      Height          =   345
      Left            =   7680
      TabIndex        =   2
      Top             =   5850
      Width           =   1065
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   2
      TX              =   ""
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
      MICON           =   "frmSlitpInfo.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdok 
      Height          =   375
      Left            =   5400
      TabIndex        =   0
      Top             =   5850
      Width           =   1035
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   2
      TX              =   ""
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
      MICON           =   "frmSlitpInfo.frx":0044
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
      Height          =   645
      Left            =   240
      TabIndex        =   4
      Top             =   180
      Width           =   8700
      Begin RTComctl3.TextButtonBox txtSlitpBusId 
         Height          =   315
         Left            =   1335
         TabIndex        =   5
         Top             =   225
         Width           =   1530
         _ExtentX        =   2699
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
      End
      Begin MSComCtl2.DTPicker dtBusDate 
         Height          =   315
         Left            =   3900
         TabIndex        =   6
         Top             =   225
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         _Version        =   393216
         Format          =   65536000
         CurrentDate     =   36950
      End
      Begin VB.Label lblSlitpBusid 
         Caption         =   "拆分车次代码"
         Height          =   225
         Left            =   225
         TabIndex        =   8
         Top             =   270
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "时  间"
         Height          =   225
         Left            =   3090
         TabIndex        =   7
         Top             =   270
         Width           =   570
      End
   End
   Begin VB.Label Label2 
      Caption         =   "拆分信息"
      Height          =   195
      Left            =   210
      TabIndex        =   9
      Top             =   900
      Width           =   780
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   1110
      X2              =   8910
      Y1              =   960
      Y2              =   960
   End
End
Attribute VB_Name = "frmSlitpInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public m_szBusID As String
Public m_dtBusDate As Date
Private m_oReBus As New REBus

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
m_szBusID = txtSlitpBusId.Text
m_dtBusDate = dtBusDate
RefreshList
End Sub



Private Sub cmdQuit_Click()
' Dim nResult As String
' If Not lvBus.SelectedItem Is Nothing Then
'    nResult = MsgBox("车次" & Trim(txtSlitpBusId.Text) & "将取消拆分", vbYesNo + vbInformation, "取消拆分")
'    If nResult = vbYes Then
'       Setbusy
'       m_oREBus.UnSlitp txtSlitpBusId.Text, m_dtBusDate
'       lvBus.ListItems.Clear
'       MsgBox "取消拆分成功", vbInformation + vbExclamation, "取消拆分"
''       frmRESlitpBus.lblSlitp(5).Caption = "正常"
''       frmRESlitpBus.txtAimReBusID(0).Enabled = True
''       frmRESlitpBus.m_bflg = False
'       frmEnvBus.updatelist txtSlitpBusId.Text, m_dtBusDate
'       Setnormal
'
'       If frmSalelAndSlitpInfo.m_bIsShow = True Then
'          frmSalelAndSlitpInfo.lblBusStatus.Caption = ""
'          frmSalelAndSlitpInfo.lblStatus.Caption = ""
'          frmSalelAndSlitpInfo.cmbFunction.ListIndex = 0
'         ' For i = 1 To frmSalelAndSlitpInfo.lvBusSale.ListItems.Count
'         '    frmSalelAndSlitpInfo.SetCorlorLvBusSale i, False
'         ' Next
'       End If
'       Unload Me
'    End If
' Else
'   MsgBox "车次" & Trim(txtSlitpBusId.Text) & "当前日期没被拆分", vbInformation, "取消拆分"
' End If
'Exit Sub
End Sub

Private Sub Form_Load()
Dim szaBusSlitpInfo() As String
Dim liTemp As ListItem
Dim nCount As Integer
Dim i As Integer
On Error GoTo ErrorHandle
Set m_oReBus = New REBus
    dtBusDate.Value = m_dtBusDate
If m_szBusID <> "" Then
   txtSlitpBusId.Text = m_szBusID
End If
m_oReBus.Init g_oActiveUser
If m_szBusID <> "" Then
   txtSlitpBusId.Text = m_szBusID
  End If

Exit Sub
ErrorHandle:
    SetNormal
End Sub
Public Sub RefreshList()
Dim szaBusSlitpInfo() As String
Dim liTemp As ListItem
Dim nCount As Integer
Dim i As Integer
On Error GoTo ErrorHandle
SetBusy
lvBus.ListItems.Clear
m_szBusID = txtSlitpBusId.Text
If m_szBusID <> "" Then
    If Not IsNull(dtBusDate.Value) Then
        m_dtBusDate = dtBusDate.Value
    Else
        m_dtBusDate = Date
    End If
    szaBusSlitpInfo = m_oReBus.GetSlitpInfo(m_szBusID, m_dtBusDate)
    nCount = ArrayLength(szaBusSlitpInfo)
    If nCount = 0 Then Exit Sub
       For i = 1 To nCount
            Set liTemp = lvBus.ListItems.Add(, , szaBusSlitpInfo(i, 1))
            liTemp.SubItems(1) = m_szBusID
            liTemp.SubItems(2) = szaBusSlitpInfo(i, 2)
            liTemp.SubItems(3) = szaBusSlitpInfo(i, 3)
            liTemp.SubItems(4) = szaBusSlitpInfo(i, 4)
            liTemp.SubItems(5) = szaBusSlitpInfo(i, 5)
            liTemp.SubItems(6) = szaBusSlitpInfo(i, 6)
            liTemp.SubItems(7) = szaBusSlitpInfo(i, 7)
            liTemp.SubItems(8) = szaBusSlitpInfo(i, 8)
            liTemp.SubItems(9) = szaBusSlitpInfo(i, 9)
         
    Next
    End If

    SetNormal
Exit Sub
ErrorHandle:
    SetNormal
    ShowErrorMsg
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

Private Sub TimStart_Timer()
 RefreshList
 TimStart.Enabled = False
End Sub

Private Sub txtSlitpBusId_Click()
 Dim oShell As New STShell.CommDialog
    Dim aszTemp() As String
    oShell.Init g_oActiveUser
    aszTemp = oShell.SelectREBus(m_dtBusDate)
    If ArrayLength(aszTemp) = 0 Then Exit Sub
    txtSlitpBusId.Text = aszTemp(1, 1)
    m_szBusID = Trim(txtSlitpBusId.Text)
End Sub

Private Sub txtSlitpBusID_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And txtSlitpBusId.Text <> "" Then
 RefreshList
End If
End Sub
Private Function QuitSlitp()
   
End Function
