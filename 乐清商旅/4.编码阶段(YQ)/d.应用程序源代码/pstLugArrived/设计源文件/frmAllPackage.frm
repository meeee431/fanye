VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmAllPackage 
   Caption         =   "行包查询"
   ClientHeight    =   7125
   ClientLeft      =   1035
   ClientTop       =   2160
   ClientWidth     =   11880
   HelpContextID   =   2000601
   Icon            =   "frmAllPackage.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7125
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.PictureBox ptShowInfo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1350
      Left            =   210
      ScaleHeight     =   1350
      ScaleWidth      =   11685
      TabIndex        =   20
      Top             =   90
      Width           =   11685
      Begin VB.TextBox txtStartStation 
         Height          =   300
         Left            =   8610
         TabIndex        =   17
         Text            =   "杭州"
         Top             =   990
         Width           =   945
      End
      Begin VB.ComboBox cboAreaType 
         Height          =   300
         Left            =   6990
         TabIndex        =   16
         Text            =   "省内"
         Top             =   990
         Width           =   1605
      End
      Begin VB.TextBox txtLicense 
         Height          =   315
         Left            =   2640
         TabIndex        =   8
         Top             =   563
         Width           =   1095
      End
      Begin VB.TextBox txtSheetID 
         Height          =   315
         Left            =   4920
         TabIndex        =   10
         Top             =   563
         Width           =   1290
      End
      Begin VB.ComboBox cboPackageName 
         Height          =   300
         Left            =   9600
         TabIndex        =   14
         Top             =   570
         Width           =   1395
      End
      Begin VB.TextBox txtPicker 
         Height          =   315
         Left            =   7200
         TabIndex        =   12
         Top             =   563
         Width           =   1155
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   315
         Left            =   4590
         TabIndex        =   3
         Top             =   173
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   393216
         Format          =   61538304
         CurrentDate     =   38630
      End
      Begin VB.ComboBox cboStatus 
         Height          =   300
         ItemData        =   "frmAllPackage.frx":014A
         Left            =   9600
         List            =   "frmAllPackage.frx":014C
         TabIndex        =   6
         Text            =   "未提"
         Top             =   180
         Width           =   1395
      End
      Begin VB.TextBox txtPackageID 
         Height          =   315
         Left            =   2310
         TabIndex        =   1
         Top             =   173
         Width           =   1425
      End
      Begin RTComctl3.CoolButton cmdFind 
         Default         =   -1  'True
         Height          =   345
         Left            =   9750
         TabIndex        =   18
         Top             =   930
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
         MICON           =   "frmAllPackage.frx":014E
         PICN            =   "frmAllPackage.frx":016A
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
         Height          =   315
         Left            =   6630
         TabIndex        =   4
         Top             =   173
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   393216
         Format          =   61538304
         CurrentDate     =   38630
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "->"
         Height          =   180
         Left            =   6330
         TabIndex        =   21
         Top             =   240
         Width           =   180
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "起运站(&S):"
         Height          =   180
         Left            =   6060
         TabIndex        =   15
         Top             =   1050
         Width           =   900
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "车号(&L):"
         Height          =   180
         Left            =   1920
         TabIndex        =   7
         Top             =   630
         Width           =   720
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "自编号(&O):"
         Height          =   180
         Left            =   1275
         TabIndex        =   0
         Top             =   240
         Width           =   900
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "收件人(&P):"
         Height          =   180
         Left            =   6270
         TabIndex        =   11
         Top             =   630
         Width           =   1080
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "日期(&D):"
         Height          =   180
         Left            =   3810
         TabIndex        =   2
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "提货状态(&P):"
         Height          =   180
         Left            =   8430
         TabIndex        =   5
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "单据号(&I):"
         Height          =   180
         Left            =   3810
         TabIndex        =   9
         Top             =   630
         Width           =   1080
      End
      Begin VB.Label lblInputRouteId 
         BackStyle       =   0  'Transparent
         Caption         =   "货物名称(&L):"
         Height          =   180
         Left            =   8460
         TabIndex        =   13
         Top             =   630
         Width           =   1080
      End
   End
   Begin MSComctlLib.ImageList imlBusIcon 
      Left            =   0
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAllPackage.frx":0504
            Key             =   "normal"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAllPackage.frx":089E
            Key             =   "picked"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAllPackage.frx":0C38
            Key             =   "canceled"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvPackage 
      Height          =   4635
      Left            =   1140
      TabIndex        =   19
      Top             =   2160
      Width           =   6585
      _ExtentX        =   11615
      _ExtentY        =   8176
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "imlBusIcon"
      SmallIcons      =   "imlBusIcon"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "自编号"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "单据号"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "车号"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "货名"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "发件人"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "收件人"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "计重"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "件数"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "存放位置"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "地区"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "到达时间"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Menu pmnu_Station 
      Caption         =   "站点"
      Visible         =   0   'False
      Begin VB.Menu pmnu_AddStation 
         Caption         =   "新增(&A)"
      End
      Begin VB.Menu pmnu_EditStation 
         Caption         =   "编辑(&E)"
         Enabled         =   0   'False
      End
      Begin VB.Menu pmnu_DeleteStation 
         Caption         =   "删除(&D)"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "frmAllPackage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************
'* Source File Name:frmAllPackage.frm
'* Project Name:PSTLugArrived
'* Engineer:
'* Data Generated:2002/08/30
'* Last Revision Date:2002/08/30
'* Brief Description:所有站点
'* Relational Document:
'**********************************************************
Const cnPackageID = 0
Const cnSheetID = 1
Const cnLicenseTagNO = 2
Const cnPackageName = 3
Const cnSendName = 4
Const cnPicker = 5
Const cnPickerPhone = 6
Const cnWeight = 7
Const cnPackageNumber = 8
Const cnSavePosition = 9
Const cnAreaType = 10
Const cnArriveTiem = 11
Const cnStatus = 12
    
'添加查询条件
Private Sub AddQueryCondition()
    Dim i As Integer
    dtpDate.Value = DateAdd("d", -6, Date)
    dtpEndDate.Value = Date
    
    Dim aszTmp() As String
    Dim aszTmpPackage() As String
    
    aszTmp = g_oPackageParam.ListBaseDefine(EDefineType.EDT_AreaType)
    cboAreaType.Clear
    cboAreaType.AddItem CSZNoneString
    For i = 1 To ArrayLength(aszTmp)
        cboAreaType.AddItem aszTmp(i, 3)
    Next i
    
    aszTmpPackage = g_oPackageParam.ListBaseDefine(EDefineType.EDT_PackageName)
    cboPackageName.Clear
    For i = 1 To ArrayLength(aszTmpPackage)
        cboPackageName.AddItem aszTmpPackage(i, 3)
    Next i
    
     
    txtStartStation.Text = ""
    cboStatus.AddItem CSZNoneString
    cboStatus.AddItem CPick_Normal
    cboStatus.ItemData(1) = EPS_Normal
    cboStatus.AddItem CPick_Picked
    cboStatus.ItemData(2) = EPS_Picked
    cboStatus.AddItem CPick_Canceled
    cboStatus.ItemData(3) = EPS_Cancel
    cboStatus.AddItem "代收"

    
    txtSheetID.Text = ""
    cboPackageName.Text = ""
    
    
     
    cboAreaType.ListIndex = 0
'    cboStatus.ListIndex = 0
    cboStatus.ListIndex = 1 '默认为未提状态
    
End Sub
Private Sub ListData()
    Dim rsTmp As Recordset
    Dim szSearch As String
    szSearch = " arrive_time>='" & Format(dtpDate.Value, "yyyy-MM-dd") & "'" & _
                " AND arrive_time<'" & Format(DateAdd("d", 1, dtpEndDate.Value), "yyyy-MM-dd") & "'"
    If Trim(cboAreaType.Text) <> "" And Trim(cboAreaType.Text) <> CSZNoneString Then
        szSearch = szSearch & " AND area_type='" & Trim(cboAreaType.Text) & "'"
    End If
    If Trim(txtStartStation.Text) <> "" And Trim(txtStartStation.Text) <> CSZNoneString Then
        szSearch = szSearch & " AND start_station_name LIKE '%" & Trim(cboAreaType.Text) & "%'"
    End If
    If Trim(cboStatus.Text) <> "" And Trim(cboStatus.Text) <> CSZNoneString And Trim(cboStatus.Text) <> "代收" Then
        szSearch = szSearch & " AND status=" & cboStatus.ItemData(cboStatus.ListIndex)
    ElseIf Trim(cboStatus.Text) = "代收" Then
        szSearch = szSearch & " AND transit_charge > 0 "
    End If
    If Trim(txtSheetID.Text) <> "" And Trim(txtSheetID.Text) <> CSZNoneString Then
        szSearch = szSearch & " AND sheet_id LIKE '%" & Trim(txtSheetID.Text) & "%'"
    End If
    If Trim(cboPackageName.Text) <> "" And Trim(cboPackageName.Text) <> CSZNoneString Then
        szSearch = szSearch & " AND package_name LIKE '%" & Trim(cboPackageName.Text) & "%'"
    End If
    If Trim(txtPicker.Text) <> "" And Trim(txtPicker.Text) <> CSZNoneString Then
        szSearch = szSearch & " AND picker LIKE '%" & Trim(txtPicker.Text) & "%'"
    End If
    If Trim(txtPackageID.Text) <> "" And Trim(txtPackageID.Text) <> CSZNoneString Then
        szSearch = szSearch & " AND CONVERT(varchar(10),package_id) LIKE '%" & Format(txtPackageID.Text, "000000") & "'"
    End If
    If Trim(txtLicense.Text) <> "" And Trim(txtLicense.Text) <> CSZNoneString Then
        szSearch = szSearch & " AND license_tag_no LIKE '%" & Trim(txtLicense.Text) & "%'"
    End If
    Set rsTmp = g_oPackageSvr.ListPackageRS(szSearch)
    
    'Fill Them
    WriteProcessBar True, 0, rsTmp.RecordCount
    lvPackage.ListItems.Clear
    Dim iCount As Integer
    Dim oListItem As ListItem
    While Not rsTmp.EOF
        iCount = iCount + 1
        WriteProcessBar , iCount, rsTmp.RecordCount
        Dim szIcon As String
        Select Case FormatDbValue(rsTmp!Status)
            Case EPS_Normal
                szIcon = "normal"
            Case EPS_Picked
                szIcon = "picked"
            Case EPS_Cancel
                szIcon = "canceled"
        End Select
        Set oListItem = lvPackage.ListItems.Add(, "k" & FormatDbValue(rsTmp!package_id), FormatDbValue(rsTmp!package_id), szIcon, szIcon)
        oListItem.SubItems(cnSheetID) = FormatDbValue(rsTmp!sheet_id)
        oListItem.SubItems(cnLicenseTagNO) = FormatDbValue(rsTmp!license_tag_no)
        oListItem.SubItems(cnPackageName) = FormatDbValue(rsTmp!package_name)
        oListItem.SubItems(cnSendName) = FormatDbValue(rsTmp!send_name)
        oListItem.SubItems(cnPicker) = FormatDbValue(rsTmp!Picker)
        oListItem.SubItems(cnPickerPhone) = FormatDbValue(rsTmp!picker_phone)
        oListItem.SubItems(cnWeight) = FormatDbValue(rsTmp!Weight)
        oListItem.SubItems(cnPackageNumber) = FormatDbValue(rsTmp!package_number)
        oListItem.SubItems(cnSavePosition) = FormatDbValue(rsTmp!save_position)
        oListItem.SubItems(cnAreaType) = FormatDbValue(rsTmp!area_type)
        oListItem.SubItems(cnArriveTiem) = Format(FormatDbValue(rsTmp!arrive_time), "yyyy-MM-dd HH:mm")
        
        oListItem.SubItems(cnStatus) = GetStatusName(FormatDbValue(rsTmp!Status))
        rsTmp.MoveNext
    Wend
    WriteProcessBar False
    

    
    
End Sub

'添加一个信息行
Public Sub AddOneInfoLine(plPackageID As Long)
On Error GoTo ErrHandle
    Dim oPackage As New Package
    oPackage.init g_oActUser
    oPackage.Identify plPackageID
    
    Dim oListItem As ListItem
    Dim szIcon As String
    Select Case oPackage.Status
        Case EPS_Normal
            szIcon = "normal"
        Case EPS_Picked
            szIcon = "picked"
        Case EPS_Cancel
            szIcon = "canceled"
    End Select
    
    Set oListItem = lvPackage.ListItems.Add(, "k" & plPackageID, oPackage.SheetID, szIcon, szIcon)
    oListItem.SubItems(cnSheetID) = oPackage.PackageName
    oListItem.SubItems(cnLicenseTagNO) = oPackage.AreaType
    oListItem.SubItems(cnPackageName) = Format(oPackage.ArrivedTime, "YYYY-MM-DD HH:mm")
    oListItem.SubItems(cnSendName) = oPackage.CalWeight
    oListItem.SubItems(cnPicker) = oPackage.PackageNumber
    oListItem.SubItems(cnWeight) = oPackage.SavePosition
    oListItem.SubItems(cnPackageNumber) = oPackage.Shipper
    oListItem.SubItems(cnSavePosition) = oPackage.Picker
    oListItem.SubItems(cnAreaType) = oPackage.AreaType
    oListItem.SubItems(cnArriveTiem) = Format(oPackage.ArrivedTime, "yyyy-MM-dd HH:mm")
    
    oListItem.SubItems(cnStatus) = GetStatusName(oPackage.Status)

    
    
    
    
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub
'添加一个信息行
Public Sub UpdOneInfoLine(plPackageID As Long)
On Error GoTo ErrHandle
    Dim oPackage As New Package
    oPackage.init g_oActUser
    oPackage.Identify plPackageID
    
    Dim oListItem As ListItem
    Dim szIcon As String
    Select Case oPackage.Status
        Case EPS_Normal
            szIcon = "normal"
        Case EPS_Picked
            szIcon = "picked"
        Case EPS_Cancel
            szIcon = "canceled"
    End Select
    
    Set oListItem = lvPackage.ListItems("k" & plPackageID)
    oListItem.Text = oPackage.SheetID
    oListItem.SmallIcon = szIcon
    oListItem.Icon = szIcon
    oListItem.SubItems(cnSheetID) = oPackage.PackageName
    oListItem.SubItems(cnLicenseTagNO) = oPackage.AreaType
    oListItem.SubItems(cnPackageName) = Format(oPackage.ArrivedTime, "YYYY-MM-DD HH:mm")
    oListItem.SubItems(cnSendName) = oPackage.CalWeight
    oListItem.SubItems(cnPicker) = oPackage.PackageNumber
    oListItem.SubItems(cnWeight) = oPackage.SavePosition
    oListItem.SubItems(cnPackageNumber) = oPackage.Shipper
    oListItem.SubItems(cnSavePosition) = oPackage.Picker
    oListItem.SubItems(cnAreaType) = oPackage.AreaType
    oListItem.SubItems(cnArriveTiem) = Format(oPackage.ArrivedTime, "yyyy-MM-dd HH:mm")
    
    oListItem.SubItems(cnStatus) = GetStatusName(oPackage.Status)


    
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub



Private Sub cmdFind_Click()
On Error GoTo ErrHandle
    SetBusy
    ShowSBInfo "正在查找,请稍等..."
    ListData
    SetNormal
    
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub Form_Load()
    InitLv
    AlignHeadWidth Me.name, lvPackage
    AddQueryCondition
End Sub

Private Sub Form_Resize()
On Error Resume Next
    ptShowInfo.Move 0, 0, Me.ScaleWidth, ptShowInfo.Height
    lvPackage.Move 0, ptShowInfo.Height, Me.ScaleWidth, Me.ScaleHeight - ptShowInfo.Height

End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveHeadWidth Me.name, lvPackage
End Sub

Private Sub lvPackage_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    SortListView lvPackage, ColumnHeader.Index

End Sub

Private Sub lvPackage_DblClick()
On Error GoTo ErrHandle
    If Not lvPackage.SelectedItem Is Nothing Then
        frmArrived.Status = EFS_Modify
        frmArrived.m_lPackageID = Right(lvPackage.SelectedItem.Key, Len(lvPackage.SelectedItem.Key) - 1)
        frmArrived.RefreshForm
        frmArrived.ZOrder 0
        frmArrived.Show
    End If

    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub InitLv()
    '初始化列表
    With lvPackage.ColumnHeaders
        .Clear
        .Add , , "自编号"
        .Add , , "单据号"
        .Add , , "车号"
        .Add , , "货名"
        .Add , , "发件人"
        .Add , , "收件人"
        .Add , , "收件人电话"
        .Add , , "计重"
        .Add , , "件数"
        .Add , , "存放位置"
        .Add , , "地区"
        .Add , , "到达时间"
        .Add , , "提货状态"
        
    End With
End Sub

Private Function GetStatusName(pnStatus As Integer) As String
    Dim szTemp As String
    Select Case pnStatus
    Case EPS_Normal
        szTemp = CPick_Normal
    Case EPS_Picked
        szTemp = CPick_Picked
    Case EPS_Cancel
        szTemp = CPick_Canceled
    End Select
    GetStatusName = szTemp
End Function
