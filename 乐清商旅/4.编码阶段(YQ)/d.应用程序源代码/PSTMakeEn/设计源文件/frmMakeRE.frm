VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMakeRE 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "生成环境"
   ClientHeight    =   5460
   ClientLeft      =   2820
   ClientTop       =   1560
   ClientWidth     =   8580
   HelpContextID   =   3000010
   Icon            =   "frmMakeRE.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   8580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Timer tmStart 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.ComboBox cboDate 
      Height          =   300
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   495
      Width           =   1905
   End
   Begin MSComctlLib.ImageList ilBig 
      Left            =   2445
      Top             =   1260
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakeRE.frx":16AC2
            Key             =   "Run"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakeRE.frx":16C1E
            Key             =   "Flow"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakeRE.frx":16D7A
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakeRE.frx":16ED6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakeRE.frx":171F2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   6525
      ScaleHeight     =   405
      ScaleWidth      =   1920
      TabIndex        =   21
      Top             =   15
      Width           =   1920
      Begin MSComctlLib.Toolbar tbView 
         Height          =   390
         Left            =   855
         TabIndex        =   15
         Top             =   30
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         ImageList       =   "imlToolbarIcons"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "小图标"
               Object.ToolTipText     =   "小图标"
               ImageKey        =   "View Small Icons"
               Style           =   2
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "列表"
               Object.ToolTipText     =   "列表"
               ImageKey        =   "View List"
               Style           =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "详细资料"
               Object.ToolTipText     =   "详细资料"
               ImageKey        =   "View Details"
               Style           =   2
               Value           =   1
            EndProperty
         EndProperty
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   315
      Left            =   7305
      TabIndex        =   13
      Top             =   5040
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   315
      Left            =   6015
      TabIndex        =   12
      Top             =   5040
      Width           =   1140
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "生成(&O)"
      Enabled         =   0   'False
      Height          =   315
      Left            =   4725
      TabIndex        =   11
      Top             =   5040
      Width           =   1140
   End
   Begin VB.CheckBox chkAllBus 
      BackColor       =   &H00FFFFC0&
      Caption         =   "全部车次(&A)"
      Height          =   300
      Left            =   7065
      TabIndex        =   10
      Top             =   4530
      Width           =   1320
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "生成设定"
      Height          =   2370
      Left            =   90
      TabIndex        =   18
      Top             =   2430
      Width           =   3015
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "重新生成站点票价(&P)"
         Enabled         =   0   'False
         Height          =   210
         Left            =   180
         TabIndex        =   6
         Top             =   1935
         Width           =   2550
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "重新生成座位(&S)"
         Enabled         =   0   'False
         Height          =   210
         Left            =   180
         TabIndex        =   5
         Top             =   1575
         Width           =   2535
      End
      Begin VB.CheckBox chkStopMake 
         BackColor       =   &H00FFFFC0&
         Caption         =   "生成停班车次(&Z)"
         Enabled         =   0   'False
         Height          =   255
         Left            =   180
         TabIndex        =   4
         Top             =   1215
         Width           =   2520
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Height          =   75
         Left            =   45
         TabIndex        =   22
         Top             =   1035
         Width           =   2910
      End
      Begin VB.CheckBox chkQuestion 
         BackColor       =   &H00FFFFC0&
         Caption         =   "出错时提示用户(&E)"
         Height          =   255
         Left            =   180
         TabIndex        =   3
         Top             =   630
         Width           =   2130
      End
      Begin VB.CheckBox chkOverlay 
         BackColor       =   &H00FFFFC0&
         Caption         =   "车次已存在则覆盖原车次(&B)"
         Enabled         =   0   'False
         Height          =   270
         Left            =   165
         TabIndex        =   2
         Top             =   315
         Value           =   1  'Checked
         Width           =   2745
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "计划信息"
      Height          =   1455
      Left            =   120
      TabIndex        =   14
      Top             =   870
      Width           =   3000
      Begin VB.Label lblEndDate 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "结束日期:1999年12月12日"
         Height          =   180
         Left            =   165
         TabIndex        =   20
         Top             =   855
         Width           =   2070
      End
      Begin VB.Label lblStartDate 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "开始日期:1999年12月11日"
         Height          =   180
         Left            =   165
         TabIndex        =   19
         Top             =   585
         Width           =   2070
      End
      Begin VB.Label lblSellBeforeDay 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "预售天数:"
         Height          =   180
         Left            =   165
         TabIndex        =   17
         Top             =   1140
         Width           =   810
      End
      Begin VB.Label lblPlanId 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "当前计划:"
         Height          =   180
         Left            =   165
         TabIndex        =   16
         Top             =   300
         Width           =   810
      End
   End
   Begin MSComctlLib.ImageList ilMakeRe 
      Left            =   1845
      Top             =   450
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakeRE.frx":1734E
            Key             =   "Run"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakeRE.frx":174AA
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakeRE.frx":17606
            Key             =   "Flow"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvBus 
      Height          =   3915
      Left            =   3240
      TabIndex        =   8
      Top             =   450
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   6906
      SortKey         =   3
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ilMakeRe"
      SmallIcons      =   "ilBig"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "车次代码"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "运行线路"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "发车时间"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "检票口"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "停班开始日期"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "停班结束日期"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   3645
      Top             =   2490
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
            Picture         =   "frmMakeRE.frx":17762
            Key             =   "View Large Icons"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakeRE.frx":17874
            Key             =   "View Small Icons"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakeRE.frx":17986
            Key             =   "View List"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakeRE.frx":17A98
            Key             =   "View Details"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdRefrech 
      Caption         =   "刷新车次列表(&R)"
      Height          =   315
      Left            =   3270
      TabIndex        =   9
      Top             =   4470
      Width           =   1770
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   75
      X2              =   8535
      Y1              =   4935
      Y2              =   4935
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   75
      X2              =   8535
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      Caption         =   "选择生成车次(&S):"
      Height          =   180
      Left            =   3270
      TabIndex        =   7
      Top             =   195
      Width           =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      Caption         =   "选择生成运行环境的车次日期(&D):"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   195
      Width           =   2700
   End
End
Attribute VB_Name = "frmMakeRE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private m_oScheme As New RegularScheme
Private m_oProject As New BusProject
Private m_dtNowDate As Date
Private m_nPreSell As Integer
Private m_tPlanID As TSchemeArrangement
Private m_oPara As New SystemParam
Private m_szOldPlan As String

Private Sub cboDate_Click()
    On Error GoTo here
    lvBus.ListItems.Clear
    m_dtNowDate = CDate(cboDate.Text)
    m_tPlanID = m_oScheme.GetExecuteBusProject(m_dtNowDate)
    m_oProject.Identify 'm_tPlanID.szProjectID
'    dtEndDate = m_oScheme.GetBusProjectEndDate(m_tPlanID.szProjectID, m_tPlanID.nSerialNo)
    If m_szOldPlan <> m_tPlanID.szProjectID Then
        lblPlanId.ForeColor = vbRed
    Else
        lblPlanId.ForeColor = vbBlack
    End If
    lblPlanId.Caption = "当前计划:" & m_tPlanID.szProjectID & "/" & m_oProject.ProjectName
    lblStartDate.Caption = "开始日期:" & Format(m_tPlanID.dtBeginDate, "YYYY年MM月DD日")
'    If dtEndDate = CDate(cszForeverDateStr) Then
'        lblEndDate.Caption = "结束日期:" & Format(dtEndDate, "YYYY年MM月DD日")
'    Else
        lblEndDate.Caption = "结束日期:计划将一直运行"
'    End If
    cmdOk.Enabled = False
    chkAllBus.Value = vbUnchecked
    Exit Sub
here:
        ShowErrorU err.Number
End Sub

Private Sub chkAllBus_Click()
    If chkAllBus.Value Then
        For i = 1 To lvBus.ListItems.Count
            lvBus.ListItems.Item(i).Selected = True
        Next
        lvBus.Enabled = False
        cmdOk.Enabled = True
    Else
        lvBus.Enabled = True
        If lvBus.ListItems.Count < 1 Then
            cmdOk.Enabled = False
        Else
            cmdOk.Enabled = True
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
DisplayHelp Me
End Sub

Private Sub cmdOk_Click()
    Dim szQuery As String
    Dim nCount As Integer
    Dim szPricetable() As String
    szPricetable = m_oScheme.GetPriceTableInfo(CDate(Format(CDate(cboDate.Text), "YYYY-MM-DD")))
    nCount = ArrayLength(szPricetable)
    szQuery = "生成计划[" & LeftAndRight(lblPlanId.Caption, False, ":") & "]的运行环境?" & vbCrLf & vbCrLf
    If nCount = 0 Then MsgBox "计划无执行票价表,不能生成环境", vbExclamation, "生成环境": Exit Sub
    szQuery = szQuery & "* 将生成的车次使用的票价表为[" & Trim(szPricetable(1, 2)) & "/" & Trim(szPricetable(1, 3)) & "]" & vbCrLf & vbCrLf
    nCount = 0
    If chkAllBus.Value = 1 Then
        szQuery = szQuery & "* 生成计划的全部车次"
    Else
        For i = 1 To lvBus.ListItems.Count
            If lvBus.ListItems.Item(i).Selected Then
                nCount = nCount + 1
            End If
        Next
        szQuery = szQuery & "* 已选中[" & nCount & "]个班车次生成"
    End If
    If MsgBox(szQuery, _
        vbQuestion + vbYesNoCancel) = vbYes Then
       MakeBus
    End If
End Sub

Private Sub cmdRefrech_Click()
Dim szaBus() As String
Dim i As Integer, nCount As Integer
Dim ltTemp As ListItem
On Error GoTo here
Me.MousePointer = vbHourglass
lvBus.ListItems.Clear
m_tPlanID = m_oScheme.GetExecuteBusProject(m_dtNowDate)
m_oProject.Identify 'm_tPlanID.szProjectID
szaBus = m_oProject.GetAllBus
nCount = ArrayLength(szaBus)
For i = 1 To nCount
    If Val(szaBus(i, 5)) = TP_RegularBus Then
        Set ltTemp = lvBus.ListItems.Add(, , Trim(szaBus(i, 1)), "Run", "Run")
    Else
        Set ltTemp = lvBus.ListItems.Add(, , Trim(szaBus(i, 1)), "Flow", "Flow")
    End If
    If DateDiff("d", CDate(szaBus(i, 6)), CDate(cboDate.Text)) >= 0 And DateDiff("d", CDate(szaBus(i, 7)), CDate(cboDate.Text)) <= 0 Then
        ltTemp.SmallIcon = "Stop"
        ltTemp.Icon = "Stop"
    End If
    ltTemp.ListSubItems.Add , , Trim(szaBus(i, 4))
    ltTemp.ListSubItems.Add , , Format(szaBus(i, 2), "HH:MM:SS")
    ltTemp.ListSubItems.Add , , szaBus(i, 3)
    If Format(CDate(szaBus(i, 6)), "YYYY-MM-DD") <> cszEmptyDateStr Then
        ltTemp.ListSubItems.Add , , Format(szaBus(i, 6), "YYYY-MM-DD")
    End If
    If Format(CDate(szaBus(i, 7)), "YYYY-MM-DD") <> cszEmptyDateStr Then
        ltTemp.ListSubItems.Add , , Format(szaBus(i, 7), "YYYY-MM-DD")
    End If
Next

    lvBus.SortKey = 3 - 1
    lvBus.SortOrder = lvwAscending
    m_nUpColumn = 3 - 1
    lvBus.Sorted = True

cmdOk.Enabled = True
Me.MousePointer = vbDefault
Exit Sub
here:
    Me.MousePointer = vbDefault
    ShowErrorU err.Number
End Sub


Private Sub Form_Load()
Dim oSysPra As New SystemParam
oSysPra.Init g_oActiveUser
 'chkStopMake.Enabled = True
  If oSysPra.MekeStopEnviroment = True Then
        chkStopMake.Value = 2
    Else
       chkStopMake.Value = 0
       chkStopMake.Enabled = False
    End If
    
Set oSysPra = Nothing
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

Private Sub tbView_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
    Select Case Button.Key
        Case "小图标"
            lvBus.View = lvwSmallIcon
        Case "列表"
            lvBus.View = lvwList
        Case "详细资料"
            lvBus.View = lvwReport
    End Select
End Sub

Private Sub MakeBus()
    Dim szExecuteApp As String
    Dim szDate As String
    Dim szBus As String
    Dim szQuery As String
    Dim szStopMake As String
    Dim szUser As String
    Dim szPassword As String
    
    On Error GoTo here
    szUser = g_oActiveUser.UserID
    szPassword = m_szPassword
   
    szDate = Format(CDate(cboDate.Text), "YYYY-MM-DD")
    If chkAllBus.Value Then
        szBus = ""
    Else
        For i = 1 To lvBus.ListItems.Count
            If lvBus.ListItems.Item(i).Selected Then
                szBus = szBus & lvBus.ListItems.Item(i).Text & ","
            End If
        Next
    szBus = Left(szBus, Len(szBus) - 1)
    szBus = "[" & szBus & "]"
    End If
       
    If chkQuestion.Value = 1 Then
        szQuery = "T"
    Else
        szQuery = "F"
    End If
  
    If chkStopMake.Value = 1 Then
        szStopMake = "T"
    Else
        szStopMake = "F"
    End If
    szExecuteApp = m_szExecute & " " & szUser & "," & szPassword & "," & szDate & "," & szBus & "," & szQuery & "," & "F," & "," & szStopMake & ",F"
    Debug.Print szExecuteApp & vbCrLf
    Shell szExecuteApp, vbNormalFocus
    Set oSysPra = Nothing
Exit Sub
here:
    ShowErrorU err.Number
End Sub

Private Sub tmStart_Timer()
Dim dtEndDate As Date
On Error GoTo here
    tmStart.Enabled = False
    Me.MousePointer = vbHourglass
    m_oPara.Init g_oActiveUser
    m_oScheme.Init g_oActiveUser
    m_oProject.Init g_oActiveUser
    m_dtNowDate = m_oPara.NowDate
    m_nPreSell = m_oPara.PreSellDate
    m_tPlanID = m_oScheme.GetExecuteBusProject(m_dtNowDate)
    m_oProject.Identify 'm_tPlanID.szProjectID
    m_szOldPlan = m_tPlanID.szProjectID
'    dtEndDate = m_oScheme.GetBusProjectEndDate(m_tPlanID.szProjectID, m_tPlanID.nSerialNo)
    lblPlanId.Caption = "当前计划:" & m_tPlanID.szProjectID & "/" & m_oProject.ProjectName
'    lblStartDate.Caption = "开始日期:" & Format(m_tPlanID.dtBeginDate, "YYYY年MM月DD日")
    lblEndDate.Caption = "结束日期:" & Format(dtEndDate, "YYYY年MM月DD日")
    lblSellBeforeDay.Caption = "预售天数:" & m_nPreSell
    For i = 0 To m_nPreSell + 1
        cboDate.AddItem Format(DateAdd("d", i, m_dtNowDate), "YYYY年MM月DD日")
    Next
    cboDate.ListIndex = 0
    Me.MousePointer = vbDefault
    Exit Sub
here:
    ShowErrorU err.Number
    Me.MousePointer = vbDefault
End Sub
