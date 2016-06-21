VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Begin VB.Form frmInsertSection 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "插入站点"
   ClientHeight    =   3390
   ClientLeft      =   2190
   ClientTop       =   2940
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "原线路信息"
      Height          =   1185
      Left            =   180
      TabIndex        =   14
      Top             =   135
      Width           =   2940
      Begin RTComctl3.TextButtonBox txtRouteID 
         Height          =   270
         Left            =   1305
         TabIndex        =   8
         Top             =   315
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   476
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
      Begin RTComctl3.TextButtonBox txtDelSectionID 
         Height          =   270
         Left            =   1305
         TabIndex        =   10
         Top             =   720
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   476
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
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "替换路段(T):"
         Height          =   180
         Left            =   135
         TabIndex        =   9
         Top             =   765
         Width           =   1080
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "线路(&R):"
         Height          =   180
         Left            =   135
         TabIndex        =   7
         Top             =   330
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "插入路段信息"
      Height          =   1185
      Left            =   180
      TabIndex        =   13
      Top             =   1575
      Width           =   2940
      Begin VB.TextBox txtSectionIDB 
         Enabled         =   0   'False
         Height          =   270
         Left            =   1110
         TabIndex        =   6
         Top             =   720
         Width           =   1665
      End
      Begin VB.TextBox txtSectionIDA 
         Enabled         =   0   'False
         Height          =   270
         Left            =   1110
         TabIndex        =   4
         Top             =   315
         Width           =   1665
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "路段&B:"
         Height          =   180
         Left            =   135
         TabIndex        =   5
         Top             =   765
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "路段&A:"
         Height          =   180
         Left            =   135
         TabIndex        =   3
         Top             =   360
         Width           =   540
      End
   End
   Begin RTComctl3.CoolButton cmdok 
      Height          =   375
      Left            =   3090
      TabIndex        =   11
      Top             =   2910
      Width           =   1245
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
      MICON           =   "frmInsertSection.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdExit 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   4530
      TabIndex        =   12
      Top             =   2910
      Width           =   1245
      _ExtentX        =   0
      _ExtentY        =   0
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
      MICON           =   "frmInsertSection.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.TextButtonBox txtInsertStationID 
      Height          =   285
      Left            =   3255
      TabIndex        =   1
      Top             =   270
      Width           =   2730
      _ExtentX        =   4815
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
   Begin MSComctlLib.ListView lvSection 
      Height          =   2205
      Left            =   3225
      TabIndex        =   2
      Top             =   585
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   3889
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "路段代码"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "路段名"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "起点站"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "终点站"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "里程数"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "公路等级"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "提示:选择路段后按回车"
      Height          =   180
      Left            =   315
      TabIndex        =   15
      Top             =   2970
      Width           =   1890
   End
   Begin VB.Label lblStation 
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      Caption         =   "插入站点(&S):"
      Height          =   180
      Left            =   3225
      TabIndex        =   0
      Top             =   60
      Width           =   1080
   End
End
Attribute VB_Name = "frmInsertSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_oRoute As New Route
Private szDelSection As String
Private szRouteID As String
Private m_szCheckText As String
Private m_oStation As New Station
Private m_oStationTemp As New Station
Private m_oSection As New Section
Private m_sbvbmsg  As Boolean


Private Sub cmdOk_Click()
    Dim liTemp As ListItem
    Dim szTemp As String
    Dim sgMileage As Single
    Dim nResult As Integer
    Dim szInsertSectionB As String
    Dim szInsertSectionA As String
    On Error GoTo ErrorHandle
    m_oRoute.Init g_oActiveUser
    szInsertSectionB = LeftAndRight(txtSectionIDB.Text, True, "[")
    szInsertSectionA = LeftAndRight(txtSectionIDA.Text, True, "[")
    If szInsertSectionA = "" Or szInsertSectionB = "" Or (szInsertSectionB = szInsertSectionA) Then
        MsgBox "终点站与起点站相同或没选中路段,不能插入", vbExclamation, "站点插入"
        Exit Sub
    End If
    nResult = MsgBox("确认插入站点" & Chr(10) & "插入开始路段" & txtSectionIDA.Text & Chr(10) & "插入结束路段" & txtSectionIDB.Text, vbYesNo + vbInformation, "站点插入")
    If nResult = vbNo Then Exit Sub
    szRouteID = LeftAndRight(txtRouteID.Text, True, "[")
    m_oRoute.Identify szRouteID
    szDelSection = LeftAndRight(txtDelSectionID.Text, True, "[")
    nResult = m_oRoute.InsertSection(szDelSection, szInsertSectionB, szInsertSectionA)
    If nResult = 1 Then
        frmArrangeSection.RefreshSection
        Unload Me
    End If
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub cmdExit_Click()
 Unload Me
End Sub

Private Sub Form_Load()
   m_oStation.Init g_oActiveUser
   m_oSection.Init g_oActiveUser
   m_oStationTemp.Init g_oActiveUser
   cmdOk.Enabled = False
   szRouteID = txtRouteID.Text
   If szRouteID = "" Then
   cmdOk.Enabled = False
   End If
End Sub

Private Sub lvSection_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim aszTemp() As String
    
    aszTemp = m_oSection.GetSectionStatAndEnd(ResolveDisplay(txtDelSectionID.Text))
    If ArrayLength(aszTemp) > 0 Then
        If aszTemp(1, 1) = ResolveDisplay(Item.ListSubItems(2)) Then
            txtSectionIDA.Text = Item.Text
        Else
            txtSectionIDB.Text = Item.Text
        End If
        IsSave
    End If
End Sub
Private Sub txtDelSectionID_Click()
    Dim oShell As New STShell.CommDialog
    Dim aszTemp() As String
    oShell.Init g_oActiveUser
    aszTemp = oShell.SelectSection("")
    Set oShell = Nothing
    If ArrayLength(aszTemp) = 0 Then Exit Sub
    txtDelSectionID.Text = aszTemp(1, 1) & "[" & aszTemp(1, 2) & "]"
    szDelSection = aszTemp(1, 1)
End Sub


Private Sub txtInsertStationID_Click()
    Dim aszTemp() As String
    Dim oShell As New STShell.CommDialog
    oShell.Init g_oActiveUser
    aszTemp = oShell.SelectStation()
    Set oShell = Nothing
'    frmSelVehicle.m_szfromstatus = "插入站点--站点"
'    aszTemp = SelectStation("", False, True)
    If ArrayLength(aszTemp) = 0 Then Exit Sub
    txtInsertStationID.Text = Trim(aszTemp(1, 1)) & "[" & Trim(aszTemp(1, 2)) & "]"
    selectctSection
    'If m_sbvbmsg = False Then
      '  selectctSection
   ' End If
'
    
End Sub
Private Sub txtInsertStationID_KeyDown(KeyCode As Integer, Shift As Integer)
 If txtInsertStationID.Text = "" Then Exit Sub
 If KeyCode <> vbKeyReturn Then Exit Sub
 
    selectctSection
    'selectctSection
End Sub

Public Sub selectctSection()
    Dim szaSection() As String
    Dim szaStarAndEndStationID() As String
    Dim oSystem As New SystemParam
    Dim liTemp As ListItem
    Dim nResult As VbMsgBoxResult
    Dim szStartStation As String, szEndStation As String
    Dim nCount As Integer
    Dim szInsertStationID As String
    Dim i As Integer
    Dim j As Integer
    Dim szStationName(1 To 2) As String
    Dim nErrFlg As Integer
'    Dim oFrmSection As Object
    Dim sgMileag As Double
    On Error GoTo Here
    nErrFlg = 0
    szInsertStationID = LeftAndRight(txtInsertStationID.Text, True, "[")
    szEndStation = Trim(txtInsertStationID.Text)
    '*************取得替换路段的首尾站点*******************
    szaStarAndEndStationID = m_oSection.GetSectionStatAndEnd(LeftAndRight(txtDelSectionID.Text, True, "["))
    '*****************************************************
    
here1:
    
    nCount = ArrayLength(szaStarAndEndStationID)
    
    If nCount <> 2 Then Exit Sub
    lvSection.ListItems.Clear
    For i = 1 To 2
    
        If i = 2 Then
            If szaStarAndEndStationID(i, 1) = szInsertStationID Then
                MsgBox "要插入站点起站和终点站相同", vbExclamation, "插入站点"
                Exit Sub
            End If
            szaSection = m_oSection.GetSESection(szInsertStationID, szaStarAndEndStationID(2, 1))
        Else
            If szaStarAndEndStationID(i, 1) = szInsertStationID Then
                MsgBox "要插入站点起站代码和终点站代码相同", vbExclamation, "插入站点"
                Exit Sub
            End If
            szaSection = m_oSection.GetSESection(szaStarAndEndStationID(1, 1), szInsertStationID)
        End If
        
        nCount = ArrayLength(szaSection)
        If nCount <> 0 Then
            lvSection.Visible = True
            For j = 1 To nCount
                m_oSection.Identify szaSection(j)
                Set liTemp = lvSection.ListItems.Add(, , szaSection(j))
                If j = 1 Then
                    sgMileag = m_oSection.Mileage
                    If txtSectionIDA.Text = "" Then
                        txtSectionIDA.Text = szaSection(j)
                    Else
                        txtSectionIDB.Text = szaSection(j)
                    End If
                End If
                liTemp.SubItems(1) = m_oSection.SectionName
                liTemp.SubItems(2) = m_oSection.BeginStationCode & "[" & m_oSection.BeginStationName & "]"
                liTemp.SubItems(3) = m_oSection.EndStationCode & "[" & m_oSection.EndStationName & "]"
                liTemp.SubItems(4) = m_oSection.Mileage
                liTemp.SubItems(5) = m_oSection.RoadLevelName
            Next j
        Else
            '取得错误号
            nErrFlg = i
            GoTo ErrorHandle
        End If
    Next i
'    Set oFrmSection = Nothing
Exit Sub
    

'以下是错误处理
ErrorHandle:
    
    On Error GoTo Here
    'm_oStationTemp.Identify szInsertStationID
    m_oSection.Identify LeftAndRight(txtDelSectionID.Text, True, "[")
    
    Select Case nErrFlg
    Case 1
        '新增第一个路段
        m_oStationTemp.Identify szaStarAndEndStationID(1, 1) '起点
        m_oStation.Identify szInsertStationID '终点
        szStationName(1) = LeftAndRight(LeftAndRight(txtInsertStationID.Text, True, "]"), False, "[")
        
        
        
        szStationName(2) = m_oStationTemp.StationName
        
        nResult = MsgBox("连接" & "'" & szStationName(2) & "'" & " 和 " & "'" & szStationName(1) & "'" & "两站点路段不存在,是否新增路段", vbQuestion + vbYesNo, "线路")
        
        If nResult = vbYes Then
'            Set oFrmSection = New frmSection
            frmSection.m_bIsInsertSection = True
            frmSection.m_eStatus = EFS_AddNew ' = EFS_AddNew
            '             frmSection.m_bOutFrmIsAdd = True
            If szaStarAndEndStationID(1, 2) <> "" Then
                frmSection.txtSectionID.Text = Left(Trim(m_oStationTemp.StationInputCode), 2) & Left(Trim(m_oStation.StationInputCode), 2)
            End If
            frmSection.txtArea.Text = Trim(m_oStation.AreaCode) & "[" & Trim(m_oStation.AreaName) & "]"
            frmSection.txtStartStation.Text = Trim(szaStarAndEndStationID(1, 1)) + "[" + Trim(m_oStationTemp.StationName) + "]"
            frmSection.txtEndStation.Text = Trim(txtInsertStationID.Text)
            frmSection.txtSectionName = Trim(Left(Trim(m_oStationTemp.StationName), 2)) & Trim(Left(Trim(m_oStation.StationName), 2))
            frmSection.txtRoadLevel.Text = Trim(m_oSection.RoadLevelCode) & "[" & Trim(m_oSection.RoadLevelName) & "]"
            frmSection.txtKm.Text = m_oSection.Mileage
            frmSection.Show vbModal
'            Set oFrmSection = Nothing
            GoTo here1
        Else
            m_sbvbmsg = True
            Exit Sub
        End If
    Case 2
        '新增第二个路段
        m_oStationTemp.Identify szaStarAndEndStationID(2, 1) '终点
        
        szStationName(1) = LeftAndRight(txtInsertStationID.Text, True, "]")
        szStationName(1) = LeftAndRight(szStationName(1), False, "[")
        
        m_oStation.Identify szInsertStationID  '起点
        szStationName(2) = m_oStationTemp.StationName
        nResult = MsgBox("连接" & "'" & szStationName(1) & "'" & " 和 " & "'" & szStationName(2) & "'" & "两站点路段不存在,是否新增路段", vbQuestion + vbYesNo, "线路")
        If nResult = vbYes Then
            frmSection.m_bIsInsertSection = True
            frmSection.m_eStatus = EFS_AddNew ' = EFS_AddNew
        
'            Set oFrmSection = New frmSection
            '             oFrmSection.Status = EFS_AddNew
            ' frmSection.Hide
            ' Load frmSection
            
            ' If szaStarAndEndStationID(2, 2) = "" Then
            frmSection.txtSectionID.Text = Left(Trim(m_oStation.StationInputCode), 2) & Left(Trim(szaStarAndEndStationID(2, 2)), 2)
            '   End If
            frmSection.txtArea.Text = Trim(m_oStationTemp.AreaCode) & "[" & Trim(m_oStationTemp.AreaName) & "]"
            frmSection.txtStartStation.Text = Trim(txtInsertStationID.Text)
            frmSection.txtEndStation.Text = Trim(szaStarAndEndStationID(2, 1)) & "[" & Trim(m_oStationTemp.StationName) & "]"
            frmSection.txtSectionName = Trim(Left(Trim(m_oStation.StationName), 2)) & Trim(Left(Trim(m_oStationTemp.StationName), 2))
            frmSection.txtRoadLevel.Text = Trim(m_oSection.RoadLevelCode) & "[" & Trim(m_oSection.RoadLevelName) & "]"
            frmSection.txtKm.Text = CStr(m_oSection.Mileage - sgMileag)
            '             frmSection.m_bOutFrmIsAdd = True
            
            frmSection.Show vbModal
'            Set oFrmSection = Nothing
            GoTo here1
        Else
            m_sbvbmsg = True
            Exit Sub
        End If
          
    End Select
    Exit Sub
    
Here:
    ShowErrorMsg
End Sub


Private Sub IsSave()
    
    If txtSectionIDA.Text = "" Or txtSectionIDB.Text = "" Or txtRouteID.Text = "" Or txtDelSectionID.Text = "" Then
        cmdOk.Enabled = False
    Else
        cmdOk.Enabled = True
    End If
    
    
End Sub

Private Sub txtSectionIDA_Change()
    IsSave
End Sub

Private Sub txtSectionIDB_Change()
    IsSave
End Sub
