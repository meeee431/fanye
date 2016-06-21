VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "mci32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmSetOption 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "选项"
   ClientHeight    =   4200
   ClientLeft      =   3690
   ClientTop       =   3825
   ClientWidth     =   6225
   HelpContextID   =   20000200
   Icon            =   "SetOption.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   Tag             =   "Modal"
   Begin VB.Frame fraOption 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3180
      Index           =   0
      Left            =   165
      TabIndex        =   15
      Top             =   390
      Width           =   5865
      Begin VB.Frame Frame2 
         Caption         =   "检票口设置"
         Height          =   2175
         Left            =   135
         TabIndex        =   16
         Top             =   510
         Width           =   5595
         Begin VB.ComboBox cboGateName 
            Height          =   300
            Left            =   990
            TabIndex        =   3
            Text            =   "cboGateName"
            Top             =   690
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.TextBox txtGateId 
            Height          =   270
            Left            =   990
            TabIndex        =   1
            Text            =   "01"
            Top             =   330
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.CommandButton cmdSelectGate 
            Caption         =   "选择检票口(&S)"
            Height          =   315
            Left            =   3480
            TabIndex        =   4
            Top             =   285
            Width           =   1935
         End
         Begin VB.Label lblCheckGateCode 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "01"
            Height          =   180
            Left            =   990
            TabIndex        =   21
            Top             =   360
            Width           =   180
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "代码(&C):"
            Height          =   180
            Left            =   195
            TabIndex        =   0
            Top             =   360
            Width           =   720
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "名称(&N):"
            Height          =   180
            Left            =   195
            TabIndex        =   2
            Top             =   720
            Width           =   720
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "说明:"
            Height          =   180
            Left            =   195
            TabIndex        =   23
            Top             =   1140
            Width           =   450
         End
         Begin VB.Label lblAnnotation 
            BackStyle       =   0  'Transparent
            Caption         =   "1号检票口"
            Height          =   435
            Left            =   960
            TabIndex        =   22
            Top             =   1140
            Width           =   4380
         End
         Begin VB.Label lblCheckGateName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "1"
            Height          =   180
            Left            =   990
            TabIndex        =   20
            Top             =   720
            Width           =   90
         End
      End
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   0
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   315
      Left            =   2340
      TabIndex        =   11
      Top             =   3705
      Width           =   1155
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   315
      Left            =   3570
      TabIndex        =   12
      Top             =   3705
      Width           =   1155
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   315
      Left            =   4860
      TabIndex        =   13
      Top             =   3705
      Width           =   1155
   End
   Begin VB.Frame fraOption 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3210
      Index           =   1
      Left            =   135
      TabIndex        =   17
      Top             =   360
      Width           =   5895
      Begin VB.Frame Frame3 
         Caption         =   "检票提示音"
         Height          =   2775
         Left            =   135
         TabIndex        =   18
         Top             =   195
         Width           =   5565
         Begin VB.CommandButton cmdBrowse 
            Caption         =   "浏览(&B)"
            Height          =   315
            Left            =   4320
            TabIndex        =   9
            Top             =   1245
            Width           =   1095
         End
         Begin MSComctlLib.Slider sldPlayer 
            Height          =   495
            Left            =   2850
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   2145
            Width           =   2685
            _ExtentX        =   4736
            _ExtentY        =   873
            _Version        =   393216
            Enabled         =   0   'False
            LargeChange     =   0
            SmallChange     =   0
         End
         Begin MCI.MMControl MMControl1 
            Height          =   375
            Left            =   2310
            TabIndex        =   10
            Top             =   1680
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   661
            _Version        =   393216
            AutoEnable      =   0   'False
            PrevVisible     =   0   'False
            NextVisible     =   0   'False
            PauseVisible    =   0   'False
            BackVisible     =   0   'False
            StepVisible     =   0   'False
            RecordVisible   =   0   'False
            EjectVisible    =   0   'False
            DeviceType      =   "WaveAudio"
            FileName        =   ""
         End
         Begin VB.TextBox txtSoundFile 
            Height          =   315
            Left            =   2310
            TabIndex        =   8
            Top             =   840
            Width           =   3105
         End
         Begin MSComctlLib.TreeView tvSoundEvent 
            Height          =   2115
            Left            =   120
            TabIndex        =   6
            Top             =   540
            Width           =   2025
            _ExtentX        =   3572
            _ExtentY        =   3731
            _Version        =   393217
            HideSelection   =   0   'False
            LabelEdit       =   1
            Style           =   7
            FullRowSelect   =   -1  'True
            Appearance      =   1
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "事件(&E)"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   300
            Width           =   855
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "声音文件名(&F):"
            Height          =   180
            Left            =   2310
            TabIndex        =   7
            Top             =   600
            Width           =   1260
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   2310
            Picture         =   "SetOption.frx":000C
            Top             =   2130
            Width           =   480
         End
      End
   End
   Begin MSComctlLib.TabStrip tabOption 
      Height          =   3540
      HelpContextID   =   4000811
      Left            =   105
      TabIndex        =   14
      Top             =   60
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   6244
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "检票(&G)"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "声音(&S)"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSetOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const fraLeft = 240
Private Const fraTop = 600
Dim m_nCurrentFrame As Integer

Dim aszGateInfo() As String                 '检票口信息数组
Dim szAnnotationInitValue As String         '检票信息初始化值
Dim nLastNodeIndex As Integer               '上一次选择的Node的Index
Dim bIsFirstLoaded As Boolean               '是否是第一次进入本系统,是则应对检票口进行初始化

'***************
Private Const cntShowCheckGate = 0          '显示检票口信息
Private Const cntSelectCheckGate = 1          '选择检票口信息


Private Const cntPosGateId = 1                    '字段位置
Private Const cntPosGateName = 2
Private Const cntPosGateAnnotation = 3
Private szCheckGateType As String
Private Sub WriteGateInfo(nMode As Integer)
    Dim i As Integer
    Dim oBaseInfo As BaseInfo
    On Error GoTo ErrorHandle
     
    If nMode = cntSelectCheckGate Then                       '选择检票口模式
        cmdSelectGate.Enabled = False
        lblCheckGateCode.Visible = False
        lblCheckGateName.Visible = False
        szAnnotationInitValue = lblAnnotation.Caption
        txtGateId.Text = lblCheckGateCode.Caption
        cboGateName.Text = lblCheckGateName.Caption
'        CboType.Text = LblType.Caption
        txtGateId.Visible = True
        cboGateName.Visible = True
'        CboType.Visible = True
        If cboGateName.ListCount = 0 Then
            Set oBaseInfo = New BaseInfo
            oBaseInfo.Init g_oActiveUser
            aszGateInfo = oBaseInfo.GetAllCheckGate(, g_oActiveUser.SellStationID)
            For i = 1 To ArrayLength(aszGateInfo)
                cboGateName.AddItem aszGateInfo(i, cntPosGateName) & "-" & aszGateInfo(i, 5)
            Next i
        End If
    Else                                    '显示模式
        txtGateId.Visible = False
        cboGateName.Visible = False
        lblCheckGateCode.Visible = True
        lblCheckGateName.Visible = True
'        LblType.Visible = True
        cmdSelectGate.Enabled = True
'        lblCheckGateCode.Visible = True
'        lblCheckGateName.Visible = True
    End If
    Set oBaseInfo = Nothing
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub
Private Sub cboGateName_Click()
    If cboGateName.ListIndex > -1 Then
        txtGateId.Text = aszGateInfo(cboGateName.ListIndex + 1, cntPosGateId)
        lblAnnotation.Caption = aszGateInfo(cboGateName.ListIndex + 1, cntPosGateAnnotation)
    End If
End Sub

Private Sub cboGateName_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub cmdBrowse_Click()
    Dim szFile As String
    dlgFile.Filter = "所有音效文件(*.wav)|*.wav"
    dlgFile.ShowOpen
    If dlgFile.FileName <> "" Then
        txtSoundFile.Text = dlgFile.FileName
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Function ValidateGateId() As Boolean
    Dim i As Integer
    Dim nLenArrayGate As Integer
    txtGateId.Text = Trim(txtGateId.Text)
    nLenArrayGate = UBound(aszGateInfo, 1)
    For i = 1 To nLenArrayGate
        If Trim(aszGateInfo(i, cntPosGateId)) = Trim(txtGateId.Text) Then
            cboGateName.ListIndex = i - 1
            cboGateName.Text = cboGateName.List(cboGateName.ListIndex)
            lblAnnotation.Caption = aszGateInfo(i, cntPosGateAnnotation)
            ValidateGateId = True
            Exit Function
        End If
    Next i
    ValidateGateId = False
End Function
Private Sub cmdCancelSelect_Click()
    WriteGateInfo cntShowCheckGate
    lblAnnotation.Caption = szAnnotationInitValue
'
'    txtGateId.Visible = False
'    cboGateName.Visible = False
'    lblCheckGateCode.Visible = True
'    lblCheckGateName.Visible = True
'    cmdSelectGate.Enabled = True
'    cmdSelected.Visible = False
'    cmdCancelSelect.Visible = False
'    lblCheckGateCode.Visible = True
'    lblCheckGateName.Visible = True
'    lblAnnotation.Caption = szAnnotationInitValue
End Sub

Private Sub cmdHelp_Click()
    Select Case tabOption.SelectedItem.Index
    Case 1
        frmSetOption.HelpContextID = 20000200
    Case 2
        frmSetOption.HelpContextID = 20000220
    End Select
    DisplayHelp Me
End Sub

Private Sub cmdOk_Click()
On Error GoTo ErrHandle
    Dim oFreeReg As CFreeReg
    
    Set oFreeReg = New CFreeReg
    oFreeReg.Init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany
    
    If Not cmdSelectGate.Enabled Then       '选择了检票口
        If Not ValidateGateId() Then
            tabOption.Tabs(1).Selected = True
            MsgboxEx "此检票口不存在！", vbExclamation, g_cszTitle_Error
            txtGateId.SetFocus
            Exit Sub
        End If
    
        If cboGateName.ListIndex > -1 Then
            lblCheckGateCode.Caption = txtGateId.Text
            lblCheckGateName.Caption = cboGateName.Text
'            LblType.Caption = CboType.Text
            oFreeReg.SaveSetting m_cRegSystemKey, "CheckGate", lblCheckGateCode.Caption
'            If CboType.Text = "全部" Then
'               CboType.ListIndex = 0
'            ElseIf CboType.Text = "单号" Then
'              CboType.ListIndex = 1
'            Else
'              CboType.ListIndex = 2
'            End If
'            oFreeReg.SaveSetting m_cRegSystemKey, "CheckGateType", CboType.ListIndex
            
            Dim oGate As New CheckGate
            oGate.Init g_oActiveUser
            oGate.Identify lblCheckGateCode.Caption
            g_tCheckInfo.SellStationID = oGate.SellStationID
            g_tCheckInfo.SellStationName = oGate.SellStationName
            g_tCheckInfo.CheckGateNo = lblCheckGateCode.Caption
            g_tCheckInfo.CheckGateName = lblCheckGateName.Caption
            
            
            If Not bIsFirstLoaded Then                  '不是第一次进入系统
                g_oChkTicket.CheckGateNo = g_tCheckInfo.CheckGateNo
                WriteCheckGateInfo
                WriteNextBus
            End If
        End If

        WriteGateInfo cntShowCheckGate
        BuildBusCollection
        If frmBusList.IsShow Then
            frmBusList.RefreshBus
        End If
    End If
    
    g_tEventSoundPath.CanceledTicket = tvSoundEvent.Nodes("CanceledTicket").Tag
    g_tEventSoundPath.CheckedTicket = tvSoundEvent.Nodes("CheckedTicket").Tag
    g_tEventSoundPath.CheckSucess = tvSoundEvent.Nodes("CheckSucess").Tag
    g_tEventSoundPath.CheckTimeOn = tvSoundEvent.Nodes("CheckTimeOn").Tag
    g_tEventSoundPath.InvalidTicket = tvSoundEvent.Nodes("InvalidTicket").Tag
    g_tEventSoundPath.NoMatchedBus = tvSoundEvent.Nodes("NoMatchedBus").Tag
    g_tEventSoundPath.ReturnedTicket = tvSoundEvent.Nodes("ReturnedTicket").Tag
    g_tEventSoundPath.StartupCheckTimeOn = tvSoundEvent.Nodes("StartupCheckTimeOn").Tag
    g_tEventSoundPath.HalfTicket = tvSoundEvent.Nodes("HalfTicket").Tag
    g_tEventSoundPath.FreeTicket = tvSoundEvent.Nodes("FreeTicket").Tag
    g_tEventSoundPath.PreferentialTicket1 = tvSoundEvent.Nodes("PreferentialTicket1").Tag
    g_tEventSoundPath.PreferentialTicket2 = tvSoundEvent.Nodes("PreferentialTicket2").Tag
    g_tEventSoundPath.PreferentialTicket3 = tvSoundEvent.Nodes("PreferentialTicket3").Tag
    
    oFreeReg.SaveSetting m_cRegSoundKey, "CanceledTicket", g_tEventSoundPath.CanceledTicket
    oFreeReg.SaveSetting m_cRegSoundKey, "CheckedTicket", g_tEventSoundPath.CheckedTicket
    oFreeReg.SaveSetting m_cRegSoundKey, "CheckSucess", g_tEventSoundPath.CheckSucess
    oFreeReg.SaveSetting m_cRegSoundKey, "CheckTimeOn", g_tEventSoundPath.CheckTimeOn
    oFreeReg.SaveSetting m_cRegSoundKey, "InvalidTicket", g_tEventSoundPath.InvalidTicket
    oFreeReg.SaveSetting m_cRegSoundKey, "NoMatchedBus", g_tEventSoundPath.NoMatchedBus
    oFreeReg.SaveSetting m_cRegSoundKey, "ReturnedTicket", g_tEventSoundPath.ReturnedTicket
    oFreeReg.SaveSetting m_cRegSoundKey, "StartupCheckTimeOn", g_tEventSoundPath.StartupCheckTimeOn
    oFreeReg.SaveSetting m_cRegSoundKey, "HalfTicket", g_tEventSoundPath.HalfTicket
    oFreeReg.SaveSetting m_cRegSoundKey, "FreeTicket", g_tEventSoundPath.FreeTicket
    oFreeReg.SaveSetting m_cRegSoundKey, "PreferentialTicket1", g_tEventSoundPath.PreferentialTicket1
    oFreeReg.SaveSetting m_cRegSoundKey, "PreferentialTicket2", g_tEventSoundPath.PreferentialTicket2
    oFreeReg.SaveSetting m_cRegSoundKey, "PreferentialTicket3", g_tEventSoundPath.PreferentialTicket3
'    tvSoundEvent.Nodes.Item("HalfTicket").Tag = g_tEventSoundPath.PreferentialTicket1
'    tvSoundEvent.Nodes.Item("HalfTicket").Tag = g_tEventSoundPath.PreferentialTicket2
'    tvSoundEvent.Nodes.Item("HalfTicket").Tag = g_tEventSoundPath.PreferentialTicket3
    
    If (Not bIsFirstLoaded) Or (Trim(g_tCheckInfo.CheckGateNo) = "") Then
        Unload Me
    Else
        Unload Me
'        MDIMain.Show
    End If
    Set oFreeReg = Nothing
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub cmdSelected_Click()
    Dim oFreeReg As CFreeReg
    On Error GoTo ErrorHandle
    If Not ValidateGateId() Then
        MsgboxEx "此检票口不存在！", vbExclamation, g_cszTitle_Error
        txtGateId.SetFocus
        Exit Sub
    End If
    If cboGateName.ListIndex > -1 Then
        lblCheckGateCode.Caption = txtGateId.Text
        lblCheckGateName.Caption = cboGateName.Text
        Set oFreeReg = New CFreeReg
        oFreeReg.Init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany
        oFreeReg.SaveSetting m_cRegSystemKey, "CheckGate", lblCheckGateCode.Caption
        g_tCheckInfo.CheckGateNo = lblCheckGateCode.Caption
        g_tCheckInfo.CheckGateName = lblCheckGateName.Caption
        If Not bIsFirstLoaded Then                  '不是第一次进入系统
            g_oChkTicket.CheckGateNo = g_tCheckInfo.CheckGateNo
            WriteCheckGateInfo
            WriteNextBus
        End If
    End If
    WriteGateInfo cntShowCheckGate
    
    '更改检票车次列表信息
    BuildBusCollection

'    txtGateId.Visible = False
'    cboGateName.Visible = False
'    lblCheckGateCode.Visible = True
'    lblCheckGateName.Visible = True
'    cmdSelectGate.Enabled = True
'    cmdSelected.Visible = False
'    cmdCancelSelect.Visible = False
'    lblCheckGateCode.Visible = True
'    lblCheckGateName.Visible = True
    Set oFreeReg = Nothing
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub cmdSelectGate_Click()
    WriteGateInfo cntSelectCheckGate
    txtGateId.SetFocus
'    cmdSelectGate.Enabled = False
'    cmdSelected.Visible = True
'    cmdCancelSelect.Visible = True
'    lblCheckGateCode.Visible = False
'    lblCheckGateName.Visible = False
'    szAnnotationInitValue = lblAnnotation.Caption
'    txtGateId.Text = lblCheckGateCode.Caption
'    cboGateName.Text = lblCheckGateName.Caption
'    txtGateId.Visible = True
'    cboGateName.Visible = True
'    If cboGateName.ListCount = 0 Then
'        Set oBaseInfo = New BaseInfo
'        oBaseInfo.Init g_oActiveUser
'        aszGateInfo = oBaseInfo.GetAllCheckGate
'        For i = 1 To UBound(aszGateInfo)
'            cboGateName.AddItem aszGateInfo(i, cntPosGateName)
'        Next i
'    End If
'    txtGateId.SetFocus
End Sub




Private Sub Form_Load()
    Dim n As Integer
    Dim fraTemp As Frame
    Dim oFreeReg As CFreeReg
 

On Error GoTo here
    AlignFormPos Me
    'CboType.ListIndex = 0
    Set oFreeReg = New CFreeReg
    oFreeReg.Init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany
    szCheckGateType = Trim(oFreeReg.GetSetting(m_cRegSystemKey, "CheckGateType"))
        For n = 0 To fraOption.Count - 1
            Set fraTemp = fraOption.Item(n)
            fraTemp.Visible = False
    '        fraTemp.Left = fraLeft
    '        fraTemp.Top = fraTop
        Next n
        fraOption.Item(0).Visible = True
        m_nCurrentFrame = 0
        
    '初始化fraOption(0)
        If g_tCheckInfo.CheckGateNo = "" Then               '没指定检票口号时必须进行初始化
            bIsFirstLoaded = True
            
            lblAnnotation.Caption = ""
            lblCheckGateCode.Caption = ""
            lblCheckGateName.Caption = ""
    '        LblType.Caption = ""
            WriteGateInfo cntSelectCheckGate        '首先选择检票口
        Else
            bIsFirstLoaded = False
            
            Dim oCheckTmp As CheckGate
            Set oCheckTmp = New CheckGate
            oCheckTmp.Init g_oActiveUser
            oCheckTmp.Identify g_tCheckInfo.CheckGateNo
            lblCheckGateCode.Caption = oCheckTmp.CheckGateCode
            lblCheckGateName.Caption = oCheckTmp.CheckGateName
            lblAnnotation.Caption = oCheckTmp.Annotation
    '        Select Case szCheckGateType
    '           Case "0"
    '              LblType.Caption = "全部"
    '           Case "1"
    '              LblType.Caption = "单号"
    '           Case "2"
    '              LblType.Caption = "双号"
    '        End Select
            
            WriteGateInfo cntShowCheckGate         '显示检票口
        End If
        
        
        
    '初始化fraOption(1)
        MMControl1.TimeFormat = mciFormatMilliseconds
        tvSoundEvent.Nodes.Add , , "Root", "检票事件"
        tvSoundEvent.Nodes.Add "Root", tvwChild, "InvalidTicket", "无效车票"
        tvSoundEvent.Nodes.Add "Root", tvwChild, "CanceledTicket", "票已废"
        tvSoundEvent.Nodes.Add "Root", tvwChild, "ReturnedTicket", "票已退"
        tvSoundEvent.Nodes.Add "Root", tvwChild, "NoMatchedBus", "非当检车次"
        tvSoundEvent.Nodes.Add "Root", tvwChild, "CheckedTicket", "已检票"
        tvSoundEvent.Nodes.Add "Root", tvwChild, "CheckSucess", "检票成功"
        tvSoundEvent.Nodes.Add "Root", tvwChild, "CheckTimeOn", "检票时间已到"
        tvSoundEvent.Nodes.Add "Root", tvwChild, "StartupCheckTimeOn", "开检时间已到"
        tvSoundEvent.Nodes.Add "Root", tvwChild, "FreeTicket", "免票提示"
        tvSoundEvent.Nodes.Add "Root", tvwChild, "HalfTicket", "半票提示"
        tvSoundEvent.Nodes.Add "Root", tvwChild, "PreferentialTicket1", "优惠票1提示"
        tvSoundEvent.Nodes.Add "Root", tvwChild, "PreferentialTicket2", "优惠票2提示"
        tvSoundEvent.Nodes.Add "Root", tvwChild, "PreferentialTicket3", "优惠票3提示"
       
        tvSoundEvent.Nodes.Item("InvalidTicket").Tag = g_tEventSoundPath.InvalidTicket
        tvSoundEvent.Nodes.Item("CanceledTicket").Tag = g_tEventSoundPath.CanceledTicket
        tvSoundEvent.Nodes.Item("ReturnedTicket").Tag = g_tEventSoundPath.ReturnedTicket
        tvSoundEvent.Nodes.Item("NoMatchedBus").Tag = g_tEventSoundPath.NoMatchedBus
        tvSoundEvent.Nodes.Item("CheckedTicket").Tag = g_tEventSoundPath.CheckedTicket
        tvSoundEvent.Nodes.Item("CheckSucess").Tag = g_tEventSoundPath.CheckSucess
        tvSoundEvent.Nodes.Item("CheckTimeOn").Tag = g_tEventSoundPath.CheckTimeOn
        tvSoundEvent.Nodes.Item("StartupCheckTimeOn").Tag = g_tEventSoundPath.StartupCheckTimeOn
        tvSoundEvent.Nodes.Item("FreeTicket").Tag = g_tEventSoundPath.FreeTicket
        tvSoundEvent.Nodes.Item("HalfTicket").Tag = g_tEventSoundPath.HalfTicket
        tvSoundEvent.Nodes.Item("PreferentialTicket1").Tag = g_tEventSoundPath.PreferentialTicket1
        tvSoundEvent.Nodes.Item("PreferentialTicket2").Tag = g_tEventSoundPath.PreferentialTicket2
        tvSoundEvent.Nodes.Item("PreferentialTicket3").Tag = g_tEventSoundPath.PreferentialTicket3
    
        tvSoundEvent.Nodes.Item("Root").Expanded = True
        tvSoundEvent.Nodes("InvalidTicket").Selected = True
        txtSoundFile.Text = tvSoundEvent.Nodes("InvalidTicket").Tag
        nLastNodeIndex = tvSoundEvent.SelectedItem.Index
            
    '    tvSoundEvent.Nodes.Add , , "Root", "检票事件"
    '    tvSoundEvent.Nodes.Add "Root", tvwChild, "keyHasChecked", "票已检"
    '    tvSoundEvent.Nodes.Add "Root", tvwChild, "keyHasReturned", "票已退"
    '    tvSoundEvent.Nodes.Add "Root", tvwChild, "keyHasBeChanged", "票已改签"
    '    tvSoundEvent.Nodes.Add "Root", tvwChild, "keyHasCanceled", "票已作废"
    '    tvSoundEvent.Nodes.Add "Root", tvwChild, "keyChanged", "正常改签"
    '    tvSoundEvent.Nodes.Add "Root", tvwChild, "keyNormal", "正常被检"
    '
    '    tvSoundEvent.Nodes.Item("keyHasChecked").Tag = g_tEventSoundPath.sndHasChecked
    '    tvSoundEvent.Nodes.Item("keyHasReturned").Tag = g_tEventSoundPath.sndHasReturned
    '    tvSoundEvent.Nodes.Item("keyHasBeChanged").Tag = g_tEventSoundPath.sndHasBeChanged
    '    tvSoundEvent.Nodes.Item("keyHasCanceled").Tag = g_tEventSoundPath.sndHasCanceled
    '    tvSoundEvent.Nodes.Item("keyChanged").Tag = g_tEventSoundPath.sndChanged
    '    tvSoundEvent.Nodes.Item("keyNormal").Tag = g_tEventSoundPath.sndNormal
    '
    '    tvSoundEvent.Nodes.Item("Root").Expanded = True
    '    tvSoundEvent.Nodes("keyHasChecked").Selected = True
    '    txtSoundFile.Text = tvSoundEvent.Nodes("keyHasChecked").Tag
    '    nLastNodeIndex = tvSoundEvent.SelectedItem.Index
    Exit Sub
here:
    ShowErrorMsg
End Sub
Private Sub RefreshFrame()
    Dim n As Integer
    For n = 0 To fraOption.Count - 1
        fraOption.Item(n).Visible = False
    Next n
    fraOption.Item(m_nCurrentFrame).Visible = True

End Sub




Private Sub Form_Unload(Cancel As Integer)
    SaveFormPos Me
End Sub

Private Sub MMControl1_Done(NotifyCode As Integer)
    sldPlayer.ClearSel
End Sub

Private Sub MMControl1_PlayClick(Cancel As Integer)
    MMControl1.FileName = txtSoundFile.Text
    MMControl1.Command = "open"
    MMControl1.UpdateInterval = MMControl1.Length / 10
    MMControl1.StopEnabled = True
    MMControl1.Command = "play"

End Sub

Private Sub MMControl1_StatusUpdate()
    If MMControl1.Position < MMControl1.Length Then
        sldPlayer.Value = (MMControl1.Position / MMControl1.Length) * 10 + 1
    Else
        MMControl1.StopEnabled = False
        MMControl1.Command = "close"
        MMControl1.UpdateInterval = 0
        sldPlayer.Value = 0
    End If
End Sub

Private Sub MMControl1_StopClick(Cancel As Integer)
    MMControl1.Command = "Close"
End Sub

Private Sub tabOption_Click()
    If Not m_nCurrentFrame = tabOption.SelectedItem.Index - 1 Then
        m_nCurrentFrame = tabOption.SelectedItem.Index - 1
        RefreshFrame
    End If
End Sub

Private Sub tvSoundEvent_NodeClick(ByVal Node As MSComctlLib.Node)
    txtSoundFile.Text = Trim(txtSoundFile.Text)
    If txtSoundFile.Text <> "" And Dir(txtSoundFile.Text) = "" Or Right(txtSoundFile.Text, 1) = "\" Then
        MsgboxEx "此文件不存在！", vbExclamation, g_cszTitle_Error
        Dim nTmpLastIndex As Integer
        nTmpLastIndex = nLastNodeIndex
        nLastNodeIndex = Node.Index
        tvSoundEvent.Nodes(nTmpLastIndex).Selected = True
                
        txtSoundFile.SelStart = 0
        txtSoundFile.SelLength = Len(txtSoundFile.Text)
        txtSoundFile.SetFocus
        Exit Sub
    End If
    If Node.Key <> "Root" Then
        If Not cmdBrowse.Enabled Then cmdBrowse.Enabled = True
'        tvSoundEvent.Nodes(nLastNodeIndex).Tag = txtSoundFile.Text
        txtSoundFile.Text = Node.Tag
        nLastNodeIndex = Node.Index
        MMControl1.StopEnabled = False
        If txtSoundFile.Text = "" Then
            MMControl1.PlayEnabled = False
        End If
    Else
        cmdBrowse.Enabled = False
        txtSoundFile.Text = ""
    End If
End Sub


Private Sub txtGateId_GotFocus()
    txtGateId.SelStart = 0
    txtGateId.SelLength = Len(txtGateId.Text)
End Sub

Private Sub txtGateId_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not ValidateGateId() Then
            MsgboxEx "此检票口不存在！", vbExclamation, g_cszTitle_Error
        End If
    End If
End Sub

Private Sub txtSoundFile_Change()
    MMControl1.PlayEnabled = False
    MMControl1.StopEnabled = False
    If txtSoundFile.Text <> "" And Dir(txtSoundFile.Text) <> "" Then
        If Right(Trim(txtSoundFile.Text), 1) = "\" Then     '不是一个路径
            Exit Sub
        Else
            MMControl1.PlayEnabled = True
            MMControl1.StopEnabled = True
            tvSoundEvent.SelectedItem.Tag = txtSoundFile.Text
        End If
    End If
    If txtSoundFile.Text = "" Then
        tvSoundEvent.SelectedItem.Tag = ""
    End If
    
End Sub
'Private Sub DisplayHelp(Optional HelpType As EHelpType = content)
'    Dim lActiveControl As Long
'
'    Select Case HelpType
'        Case content
'            lActiveControl = Me.ActiveControl.HelpContextID
'            If lActiveControl = 0 Then
'                TopicID = Me.HelpContextID
'                CallHTMLShowTopicID
'            Else
'                TopicID = lActiveControl
'                CallHTMLShowTopicID
'            End If
'        Case Index
'            CallHTMLHelpIndex
'        Case Support
'            TopicID = clSupportID
'            CallHTMLShowTopicID
'    End Select
'
'End Sub
