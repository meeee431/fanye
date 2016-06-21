VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "actbar2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDISellTicket 
   BackColor       =   &H8000000C&
   Caption         =   "售票台"
   ClientHeight    =   7395
   ClientLeft      =   825
   ClientTop       =   960
   ClientWidth     =   10800
   HelpContextID   =   4000001
   Icon            =   "MDISellTicket.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   WindowState     =   2  'Maximized
   Begin ActiveBar2LibraryCtl.ActiveBar2 abMenuTool 
      Align           =   1  'Align Top
      Height          =   7395
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10800
      _LayoutVersion  =   1
      _ExtentX        =   19050
      _ExtentY        =   13044
      _DataPath       =   ""
      Bands           =   "MDISellTicket.frx":0CCA
      Begin MSComctlLib.ProgressBar pbLoad 
         Height          =   225
         Left            =   3450
         TabIndex        =   6
         Top             =   6570
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.PictureBox ptTitle 
         BackColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   0
         ScaleHeight     =   375
         ScaleMode       =   0  'User
         ScaleWidth      =   21881.56
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   2850
         Width           =   15360
         Begin MSComctlLib.TabStrip tsUnit 
            Height          =   375
            Left            =   6495
            TabIndex        =   2
            Top             =   15
            Width           =   10305
            _ExtentX        =   18177
            _ExtentY        =   661
            Style           =   2
            HotTracking     =   -1  'True
            Separators      =   -1  'True
            TabMinWidth     =   1764
            ImageList       =   "ImageList2"
            _Version        =   393216
            BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
               NumTabs         =   1
               BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "本站 &1"
                  ImageVarType    =   2
                  ImageIndex      =   1
               EndProperty
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   11.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label lblLeaveNum 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   4980
            TabIndex        =   8
            Top             =   30
            Width           =   165
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "剩余张数:"
            Height          =   180
            Left            =   4140
            TabIndex        =   7
            Top             =   120
            Width           =   810
         End
         Begin VB.Label fblCurrentTime 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0:00:00"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   120
            TabIndex        =   5
            Top             =   30
            Width           =   1155
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "当前票号:"
            Height          =   180
            Left            =   1890
            TabIndex        =   4
            Top             =   120
            Width           =   810
         End
         Begin VB.Label lblTicketNo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   2760
            TabIndex        =   3
            Top             =   30
            Width           =   165
         End
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   2760
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDISellTicket.frx":7902
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   900
      Left            =   5640
      Top             =   2790
   End
End
Attribute VB_Name = "MDISellTicket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_bPing As Boolean

Const cszBankStat = "网点代售统计"
Const cszSellDetail = "售票明细查询"
Const cszSellerEveryDay = "网点代售统计详情"

'Private Sub abMenuTool_KeyDown(keycode As Integer, shift As Integer)
'    On Error Resume Next
'    me.ActiveForm.form_keydown(keycode,shift)
'End Sub

Private Sub abMenuTool_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    Select Case Tool.name
    Case "mnuSellTkt"
'        SellTkt
    Case "mnuExtraTkt"
'        ExtraTkt
    Case "mnuChangeTkt"
'        ChangeTkt
    Case "mnuReturnTkt"
'        ReturnTkt
    Case "mnuCancelTkt"
'        CancelTkt
        
        
    Case "mnuSellDiscountTkt"
        '售折扣票
        SellDiscountTkt
        
        
    Case "mnuCancelIns"
'        CancelInsurance
    Case "mnuChangeSeatType"
'        ChangeSeatType
    Case "mnuTicketQuery"
'        frmQuerySellTk.SelfUser = m_oAUser
'        frmQuerySellTk.Show
    Case "mnuRemoteLogin"
'        mnuRemoteLogin_Click
    Case "mnuChgTktStartNumber"
'        mnuChgTktStartNumber_Click
    Case "mnuChgPassword"
'        mnuChgPassword_Click
    Case "mnuExit"
        mnuExit_Click
    
    Case "mnuSellDetail"
        '售票明细查询
        SellDetail
    Case "mnuBankStat"
        '网点代售统计
        BankStat
    Case "mnuBankStatDetail"
        '网点代售统计详情表
        BankStatDetail
    
    
    
        '以下是系统部分
        Case "tbn_system_print"
            ActiveForm.PrintReport False
        Case "mnu_system_print"
            ActiveForm.PrintReport True
        Case "tbn_system_printview", "mnu_system_printview"
            ActiveForm.PreView
        Case "mnu_PageOption"
            '页面设置
            ActiveForm.PageSet
        Case "mnu_PrintOption"
            '打印设置
            ActiveForm.PrintSet
        Case "tbn_system_export", "mnu_ExportFile"
            ActiveForm.ExportFile
        Case "tbn_system_exportopen", "mnu_ExportFileOpen"
            ActiveForm.ExportFileOpen
    
    Case "mnuContents"
        mnuContents_Click
    Case "mnuIndex"
        mnuIndex_Click
    Case "mnuAbout"
        mnuAbout_Click
    Case Else
        If Left(Tool.name, 13) = "mnuRemoteUnit" Then
            '远程单位
            ChangeUnit Tool.TagVariant
    '        ChangeUnit mnuRemoteUnit(Index).Tag
        End If
    End Select
End Sub




Private Sub MDIForm_Load()
    AddControlsToActBar
    '状态条
    ShowSBInfo "", ESB_WorkingInfo
    ShowSBInfo "", ESB_ResultCountInfo
'    ShowSBInfo EncodeString(GetActiveUserID) & m_oAUser.UserName, ESB_UserInfo
    ShowSBInfo Format(Time, "hh:mm")
'    WriteProcessBar False
    
    
    
'    frmChgStartTktNumber.m_bNoCancel = True
'    frmChgStartTktNumber.Show vbModal, Me
    SetCaption
    
    m_bPing = False
    InitMDIForm
End Sub

Private Sub MDIForm_Resize()
    On Error Resume Next
    If Me.WindowState = vbNormal Then Me.WindowState = vbMaximized
    'abMenuTool.Width = Me.ScaleWidth
    
    abMenuTool.Bands("bndTitle").Width = Me.ScaleWidth
    ptTitle.Width = Me.ScaleWidth - 50
'    ptTitle.Refresh
    
'    fraSeller.Top = Me.ScaleHeight - fraSeller.Height - 200
End Sub


Private Sub MDIForm_Terminate()
'  CloseComputer
    End
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
'    SaveAppSetting
   
End Sub


Private Sub mnuAbout_Click()
'    Dim oShell As New STShell.CommShell
'    On Error GoTo ErrorHandle
'    oShell.ShowAbout "RTStation售票台", "RTStation Sell Ticket ", "RTStation瑞通客运售票台", Me.Icon, App.Major, App.Minor, App.Revision
'    Set oShell = Nothing
'    Exit Sub
'ErrorHandle:
'    Set oShell = Nothing
'    ShowErrorMsg
End Sub

'Private Sub mnuCancelQuery_Click()
'    frmSellQuery.m_nQueryType = TP_QueryCancel
'    frmSellQuery.Show vbModal, Me
'End Sub

Private Sub mnuCancelTkt_Click()
    CancelTkt
End Sub





'Private Sub mnuChangeQuery_Click()
'    frmSellQuery.m_nQueryType = TP_QueryChange
'    frmSellQuery.Show vbModal, Me
'End Sub

Private Sub mnuChangeTkt_Click()
    ChangeTkt
End Sub

Private Sub mnuChgPassword_Click()
    frmUserPwd.Show vbModal
    
    Exit Sub
ErrorHandle:
End Sub

Private Sub mnuChgTktStartNumber_Click()
    frmChgStartTktNumber.m_bNoCancel = False
    frmChgStartTktNumber.Show vbModal, Me
    If frmChgStartTktNumber.m_bOk Then
        lblTicketNo.Caption = GetTicketNo()
'        lblEndTicketNo.Caption = GetEndTicketNo()
        
        SetCaption
    End If
End Sub



Private Sub mnuContents_Click()
    If Not ActiveForm Is Nothing Then
       DisplayHelp ActiveForm, content
    End If
End Sub

Private Sub mnuExit_Click()
'    If MsgBox("退出售票系统吗？", vbYesNo Or vbQuestion Or vbDefaultButton2) = vbYes Then
        Unload Me
'    End If
End Sub

Private Sub mnuExtraTkt_Click()
    ExtraTkt
End Sub

Private Sub mnuIndex_Click()
    DisplayHelp Me, Index
End Sub
'
'Private Sub mnuOrderBy_Click(Index As Integer)
'    Dim i As Integer
'    Dim frmTemp As ISortKeyChanged
'    For i = 1 To 4
'        If i - 1 = Index Then
'            mnuOrderBy(i - 1).Checked = True
'        Else
'            mnuOrderBy(i - 1).Checked = False
'        End If
'    Next
'    On Error GoTo here
'    Set frmTemp = Me.ActiveForm
'    On Error GoTo 0
'    If Not frmTemp Is Nothing Then
'        Dim nTemp As Integer
'
'        Select Case Index + 1
'            Case SK_OffTime
'
'            nTemp = ID_OffTime + 1
'            Case SK_SeatCount
'            nTemp = ID_SeatCount + 1
'
'            Case SK_VehicleModel
'            nTemp = ID_VehicleModel + 1
'
'            Case SK_TicketPrice
'            nTemp = ID_FullPrice + 1
'
'        End Select
'        frmTemp.SortKeyChangedTo nTemp
'
'    End If
'    Exit Sub
'here:
'End Sub
'


Private Sub mnuRemoteLogin_Click()
'    Dim szTemp As String
'    frmRemoteConnect.Show vbModal, Me
'
'    If frmRemoteConnect.m_bOk Then
'        Dim oTab As MSComctlLib.Tab
'        Dim nTemp As Integer
'        nTemp = tsUnit.Tabs.count
'
'        On Error GoTo Here
'        abMenuTool.Bands("mnuRemote").Tools("mnuRemoteUnit" & CStr(nTemp)).Checked = False
'        abMenuTool.Bands("mnuRemote").Tools("mnuRemoteUnit" & CStr(nTemp)).TagVariant = frmRemoteConnect.m_szUnitID
'        abMenuTool.Bands("mnuRemote").Tools("mnuRemoteUnit" & CStr(nTemp)).Caption = frmRemoteConnect.m_szUnitName & "(&" & nTemp + 1 & ")"
'        abMenuTool.Bands("mnuRemote").Tools("mnuRemoteUnit" & CStr(nTemp)).Visible = True
''
''        mnuRemoteUnit(nTemp).Checked = False
''        mnuRemoteUnit(nTemp).Tag = frmRemoteConnect.m_szUnitID
''        mnuRemoteUnit(nTemp).Caption = frmRemoteConnect.m_szUnitName & "(&" & nTemp + 1 & ")"
''        mnuRemoteUnit(nTemp).Visible = True
''
'        szTemp = frmRemoteConnect.m_szUnitName & " &" & nTemp + 1
'
'        Set oTab = tsUnit.Tabs.Add(, GetEncodedKey(frmRemoteConnect.m_szUnitID), szTemp, 1)
'        oTab.Tag = frmRemoteConnect.m_szUnitID
'
'    End If
'    Exit Sub
'Here:
'    ShowErrorMsg
End Sub

Private Sub mnuRemoteUnit_Click(Index As Integer)
    
'    ChangeUnit mnuRemoteUnit(Index).Tag
End Sub

'Private Sub mnuReturnQuery_Click()
'    frmSellQuery.m_nQueryType = TP_QueryReturn
'    frmSellQuery.Show vbModal, Me
'End Sub

Private Sub mnuReturnTkt_Click()
    ReturnTkt
End Sub


Private Sub mnuSellTkt_Click()
    SellTkt
End Sub

Private Sub mnuTicketQuery_Click()
'    frmQuerySellTk.SelfUser = m_oAUser
'    frmQuerySellTk.Show , MDISellTicket
End Sub

'当前功能转变到售票
Public Sub SellTkt()
    Dim frmTemp As Form
    'Set m_clSell = frmSell
    
    Set frmTemp = GetObjecInCollection(GetEncodedKey(m_szCurrentUnitID), m_clSell)
    
    m_nCurrentTask = RT_SellTicket
    If frmTemp Is Nothing Then
        Set frmTemp = New frmSell
        frmTemp.Caption = MakeDisplayString("售票", GetUnitNameFromMenu(m_szCurrentUnitID))
        frmTemp.Tag = m_szCurrentUnitID
        m_clSell.Add frmTemp, GetEncodedKey(m_szCurrentUnitID)
        'lblSell.Visible = True
        frmTemp.Show
    Else
        frmTemp.ZOrder
    End If
    RestoreCheckLabel
    abMenuTool.Bands("mnuFunction").Tools("mnuSellTkt").Checked = True
'    mnuSellTkt.Checked = True
    
    EnableCanRemote True
    Set frmTemp = Nothing
End Sub



'当前功能转变到售票
Public Sub SellDiscountTkt()
'
'    frmSellDiscountTkt.Show
'    frmSellDiscountTkt.ZOrder 0
'
'    EnableCanRemote False
    
End Sub


'当前功能转变到退票
Public Sub ReturnTkt()
'    Dim frmTemp As Form
'    Set frmTemp = GetObjecInCollection(GetEncodedKey(m_szCurrentUnitID), m_clReturn)
'
'    m_nCurrentTask = RT_ReturnTicket
'    If frmTemp Is Nothing Then
'        Set frmTemp = New frmReturnTicket
'        frmTemp.Caption = MakeDisplayString("退票", GetUnitNameFromMenu(m_szCurrentUnitID))
'        frmTemp.Tag = m_szCurrentUnitID
'        m_clReturn.Add frmTemp, GetEncodedKey(m_szCurrentUnitID)
'        'lblReturn.Visible = True
'        frmTemp.Show
'    Else
'        frmTemp.ZOrder
'    End If
'
'    RestoreCheckLabel
'    abMenuTool.Bands("mnuFunction").Tools("mnuReturnTkt").Checked = True
''    mnuReturnTkt.Checked = True
'
'    EnableCanRemote True
'    Set frmTemp = Nothing
'
''    frmReturnTicket.Show
''    frmReturnTicket.ZOrder
''
''    m_nCurrentTask = RT_ReturnTicket
''
''    'lblCancel.Visible = True
''
''    RestoreCheckLabel
''    lblReturn.Value = vbChecked
''    mnuReturnTkt.Checked = True
''
''    EnableCanRemote False


End Sub

'当前功能转变到补票
Public Sub ExtraTkt()
'
'    frmExtraSell.Show
'    frmExtraSell.ZOrder
'
'    m_nCurrentTask = RT_ExtraSellTicket
'
'    RestoreCheckLabel
'    abMenuTool.Bands("mnuFunction").Tools("mnuExtraTkt").Checked = True
''    mnuExtraTkt.Checked = True
'
'    EnableCanRemote False
End Sub

'当前功能转变到改签
Public Sub ChangeTkt()
'
'    Dim frmTemp As Form
'    Set frmTemp = GetObjecInCollection(GetEncodedKey(m_szCurrentUnitID), m_clChange)
'
'    m_nCurrentTask = RT_ChangeTicket
'    If frmTemp Is Nothing Then
'        Set frmTemp = New frmChangeSell
'        frmTemp.Caption = MakeDisplayString("改签", GetUnitNameFromMenu(m_szCurrentUnitID))
'        frmTemp.Tag = m_szCurrentUnitID
'        m_clChange.Add frmTemp, GetEncodedKey(m_szCurrentUnitID)
'        frmTemp.Show
'    Else
'        frmTemp.ZOrder
'    End If
'
'    RestoreCheckLabel
'    abMenuTool.Bands("mnuFunction").Tools("mnuChangeTkt").Checked = True
''    mnuChangeTkt.Checked = True
'
'    EnableCanRemote True
'    Set frmTemp = Nothing
End Sub

'当前功能转变到废票
Public Sub CancelTkt()

'    Dim frmTemp As Form
'    Set frmTemp = GetObjecInCollection(GetEncodedKey(m_szCurrentUnitID), m_clCancel)
'
'    m_nCurrentTask = RT_CancelTicket
'    If frmTemp Is Nothing Then
'        Set frmTemp = New frmCancelTicket
'        frmTemp.Caption = MakeDisplayString("废票", GetUnitNameFromMenu(m_szCurrentUnitID))
'        frmTemp.Tag = m_szCurrentUnitID
'        m_clCancel.Add frmTemp, GetEncodedKey(m_szCurrentUnitID)
'        frmTemp.Show
'    Else
'        frmTemp.ZOrder
'    End If
'
'    RestoreCheckLabel
'    abMenuTool.Bands("mnuFunction").Tools("mnuCancelTkt").Checked = True
''    mnuCancelTkt.Checked = True
'
'    EnableCanRemote True
'    Set frmTemp = Nothing



    frmCancelTicket.ZOrder 0
    frmCancelTicket.Show

End Sub

Private Sub CancelInsurance()
'    '这样写,因为不需要远程的废保险
'
'    frmCancelInsurance.ZOrder 0
'    frmCancelInsurance.Show
'
'    m_nCurrentTask = RT_ExtraSellTicket
'
'    abMenuTool.Bands("mnuFunction").Tools("mnuCancelIns").Checked = True
'
'
'    EnableCanRemote False


End Sub


Private Sub Picture1_Resize()
    Dim lWidth As Long
    'Dim lHeight As Long
'    lWidth = Picture1.ScaleWidth - Picture2.Width
'    lWidth = IIf(lWidth > 0, lWidth, 0)
'    tsUnit.Move Picture2.Width, 0, lWidth
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case cszSellTicket
            SellTkt
        Case cszExtraSellTicket
            ExtraTkt
        
        Case cszChangeTicket
            ChangeTkt
        
        Case cszReturnTicket
            ReturnTkt
            
        Case cszCancelTicket
            CancelTkt
        Case cszHelp
            
        Case cszAbout
    
        Case cszExit
            End
    
    End Select

End Sub

Public Sub InitMDIForm()
'    Set abMenuTool.Bands("bndTitle").Tools("tblTitle").Custom = ptTitle
'
'
'    lblTicketNo.Caption = GetTicketNo()
''    lblEndTicketNo.Caption = GetEndTicketNo()
'    SetCaption
''    Dim oTab As MSComctlLib.Tab
''    tsUnit.Tabs.Clear
''    abMenuTool.Bands("mnuRemote").Tools("mnuRemoteUnit1").TagVariant = m_szCurrentUnitID
''    Set oTab = tsUnit.Tabs.Add(, GetEncodedKey(m_szCurrentUnitID), "本站 &1", _
''          1)
''    oTab.Tag = m_szCurrentUnitID
'    SellTkt

End Sub

Private Sub Timer1_Timer()
    fblCurrentTime.Caption = Time()
End Sub






Private Sub tsUnit_Click()
'
'    Dim echoReturn As ICMP_ECHO_REPLY
'    Dim szUnitIP As String
'    Dim frmTemp As Form
'On Error GoTo ErrorHandle
'If m_bPing Then
'   szUnitIP = m_oSell.GetUnitIP(tsUnit.SelectedItem.Tag)
''   IsConnected szUnitIP, echoReturn
''   If echoReturn.status <> IP_SUCCESS Then
''        MsgBox "[一般性网络错误]连接服务器失败！！", vbInformation, "错误！"
'
'        Select Case m_nCurrentTask
'            Case RT_SellTicket
'                Set frmTemp = GetObjecInCollection(GetEncodedKey(tsUnit.SelectedItem.Tag), m_clSell)
'                If Not frmTemp Is Nothing Then
''                   Unload frmTemp
'                    frmTemp.ZOrder 0
'                End If
'
'            Case RT_CancelTicket
'                'm_clCancel
'            Case RT_ChangeTicket
'                'm_clChange
'            Case RT_ReturnTicket
'               ' m_clReturn
'        End Select
''       ' Set Me.ActiveForm = Nothing
''       If tsUnit.Tabs.count > 1 Then
'''            tsUnit.Tabs.Remove tsUnit.SelectedItem.Index
'''            tsUnit.Tabs.Item(1).Selected = True
''
''            m_bSelfChangeUnitOrFun = False
''            m_szCurrentUnitID = tsUnit.SelectedItem.Tag
''       End If
'End If
'   ChangeUnit tsUnit.SelectedItem.Tag
'   m_bPing = True
'   Set frmTemp = Nothing
'   Exit Sub
'ErrorHandle:
'If tsUnit.Tabs.count > 1 Then
'   Set frmTemp = GetObjecInCollection(GetEncodedKey(tsUnit.SelectedItem.Tag), m_clSell)
'   If Not frmTemp Is Nothing Then
'         Unload frmTemp
'
'   End If
'   tsUnit.Tabs.Remove tsUnit.SelectedItem.Index
'   tsUnit.Tabs.Item(1).Selected = True
'   m_bSelfChangeUnitOrFun = False
'   m_szCurrentUnitID = tsUnit.SelectedItem.Tag
'   Set frmTemp = Nothing
'End If
End Sub

'按照当前的功能类型和提供票务服务所在的单位设置LABEL AND TABSCRIPT
'每当前窗口激活时都要调用

Public Sub SetFunAndUnit()

    Dim aForm As Variant
    Dim i As Integer
    On Error GoTo here
    If m_bSelfChangeUnitOrFun Then Exit Sub
    m_bSelfChangeUnitOrFun = True
    
    EnVisibleCheckLabel
    RestoreCheckLabel
    EnableCanRemote True
    Select Case m_nCurrentTask
        Case RT_SellTicket
        
        abMenuTool.Bands("mnuFunction").Tools("mnuSellTkt").Checked = True
'        mnuSellTkt.Checked = True
        ShowSBInfo "售票", ESB_WorkingInfo
        
        Case RT_ExtraSellTicket

        abMenuTool.Bands("mnuFunction").Tools("mnuExtraTkt").Checked = True
'        mnuExtraTkt.Checked = True
        ShowSBInfo "补票", ESB_WorkingInfo
        EnableCanRemote False
        
        Case RT_ReturnTicket

        abMenuTool.Bands("mnuFunction").Tools("mnuReturnTkt").Checked = True
'        mnuReturnTkt.Checked = True
        ShowSBInfo "退票", ESB_WorkingInfo
        EnableCanRemote True
        
        Case RT_ChangeTicket

        abMenuTool.Bands("mnuFunction").Tools("mnuChangeTkt").Checked = True
'        mnuChangeTkt.Checked = True
        ShowSBInfo "改签", ESB_WorkingInfo
        
        Case RT_CancelTicket

        abMenuTool.Bands("mnuFunction").Tools("mnuCancelTkt").Checked = True
'        mnuCancelTkt.Checked = True
        ShowSBInfo "废票", ESB_WorkingInfo
        
        EnableCanRemote True
        
    End Select
'    '******************************
'    '此处如果是折扣票，应该注释
'    If Not tsUnit.Tabs(GetEncodedKey(m_szCurrentUnitID)).Selected Then
'        tsUnit.Tabs(GetEncodedKey(m_szCurrentUnitID)).Selected = True
'    End If
'    '******************************
    For i = 1 To 9
        If abMenuTool.Bands("mnuRemote").Tools("mnuRemoteUnit" & CStr(i)).TagVariant = m_szCurrentUnitID Then
            abMenuTool.Bands("mnuRemote").Tools("mnuRemoteUnit" & CStr(i)).Checked = True
        Else
            abMenuTool.Bands("mnuRemote").Tools("mnuRemoteUnit" & CStr(i)).Checked = False
        End If
    Next
    ShowSBInfo MakeDisplayString(abMenuTool.Bands("statusBar").Tools("pnWorkingInfo").Caption, GetUnitName())
    '******************************
    '此处如果是折扣票，应该注释
    '=========================================
'    m_oSell.SellUnitCode = m_szCurrentUnitID
    '=========================================
    '******************************
    m_bSelfChangeUnitOrFun = False
    Exit Sub
here:
  ShowErrorMsg
End Sub

'改变单位
Public Sub ChangeUnit(pszUnitID As String)
    Dim szOldCurrentUnitID As String
    
    szOldCurrentUnitID = m_szCurrentUnitID
    
    If m_bSelfChangeUnitOrFun Then Exit Sub
    
    m_bSelfChangeUnitOrFun = True
On Error GoTo ErrorHandle:
    If m_szCurrentUnitID <> pszUnitID Then
        
        m_szCurrentUnitID = pszUnitID
        
        Select Case m_nCurrentTask
            Case RT_SellTicket
'            lblSell.Value = vbChecked
            SellTkt
            
            Case RT_ChangeTicket
'            lblChange.Value = vbChecked
            ChangeTkt
            
            Case RT_ReturnTicket
'            lblReturn.Value = vbChecked
            ReturnTkt

            Case RT_CancelTicket
'            lblCancel.Value = vbChecked
            CancelTkt
        End Select
        
        EnVisibleCheckLabel
    End If
   
    m_bSelfChangeUnitOrFun = False
Exit Sub
ErrorHandle:
    m_szCurrentUnitID = szOldCurrentUnitID
    
End Sub

'将当前所有可见的Check标签设为正常状态
Public Sub RestoreCheckLabel()
'    If lblSell.Enabled Then
'        If lblSell.Value = vbChecked Then
'            lblSell.Value = vbUnchecked
'            mnuSellTkt.Checked = False
'        End If
'    End If
'
'    If lblExtra.Enabled Then
'        If lblExtra.Value = vbChecked Then
'            lblExtra.Value = vbUnchecked
'            mnuExtraTkt.Checked = False
'        End If
'    End If
'
'    If lblReturn.Enabled Then
'        If lblReturn.Value = vbChecked Then
'            lblReturn.Value = vbUnchecked
'            mnuReturnTkt.Checked = False
'        End If
'    End If
'
'    If lblChange.Enabled Then
'        If lblChange.Value = vbChecked Then
'            lblChange.Value = vbUnchecked
'            mnuChangeTkt.Checked = False
'        End If
'    End If
'
'    If lblCancel.Enabled Then
'        If lblCancel.Value = vbChecked Then
'            lblCancel.Value = vbUnchecked
'            mnuCancelTkt.Checked = False
'        End If
'    End If
End Sub

'根据当前的状态设置好CheckLabel的可视状态
Public Sub EnVisibleCheckLabel()
    
    On Error GoTo here
'        Dim frmTemp As Form
'
'        Set frmTemp = GetObjecInCollection(GetEncodedKey(m_szCurrentUnitID), m_clSell)
'        If frmTemp Is Nothing Then
'            lblSell.Visible = False
'        Else
'            lblSell.Visible = True
'        End If
'
'        Set frmTemp = GetObjecInCollection(GetEncodedKey(m_szCurrentUnitID), m_clReturn)
'        If frmTemp Is Nothing Then
'            lblReturn.Visible = False
'        Else
'            lblReturn.Visible = True
'        End If
'
'        Set frmTemp = GetObjecInCollection(GetEncodedKey(m_szCurrentUnitID), m_clChange)
'        If frmTemp Is Nothing Then
'            lblChange.Visible = False
'        Else
'            lblChange.Visible = True
'        End If
        
        
        
        
        
'        Set frmTemp = GetObjecInCollection(GetEncodedKey(m_szCurrentUnitID), m_clExtra)
'        If frmTemp Is Nothing Then
'            lblExtra.Visible = False
'        Else
'            lblExtra.Visible = True
'        End If

'    If m_oParam.UnitID = m_szCurrentUnitID Or m_szCurrentUnitID = "" Then
        EnableExSell True
'    Else
'        EnableExSell False
'    End If
    Exit Sub
here:
    ShowErrorMsg
End Sub

'按要求设置可否进行远程操作(或远程的概念是否有意义)
Public Sub EnableCanRemote(pbCan As Boolean)
'    Dim i As Integer
    tsUnit.Enabled = pbCan
'    For i = 1 To mnuRemoteUnit.Count
'        mnuRemoteUnit(i - 1).Enabled = pbCan
'    Next
'    mnuRemoteLogin.Enabled = pbCan
End Sub

'得到当前的序方式
Public Function GetSortKey() As Integer
'    Dim i As Integer
    Dim nTemp As Integer
'    For i = 1 To 4
'        If mnuOrderBy(i - 1).Checked Then Exit For
'    Next
'
'    Select Case i
'        Case SK_OffTime
'
        nTemp = ID_OffTime + 1
'        Case SK_SeatCount
'        nTemp = ID_SeatCount + 1
'
'        Case SK_VehicleModel
'        nTemp = ID_VehicleModel + 1
'
'        Case SK_TicketPrice
'        nTemp = ID_FullPrice + 1
'
'    End Select
    GetSortKey = nTemp
End Function

'显示HTMLHELP,直接拷贝
Private Sub DisplayHelp(frmTemp As Form, Optional HelpType As EHelpType = content)
    Dim lActiveControl As Long
    
    Select Case HelpType
        Case content
            lActiveControl = frmTemp.ActiveControl.HelpContextID
            If lActiveControl = 0 Then
                TopicID = Me.HelpContextID
                CallHTMLShowTopicID
            Else
                TopicID = lActiveControl
                CallHTMLShowTopicID
            End If
        Case Index
            CallHTMLHelpIndex
        Case Support
            TopicID = clSupportID
            CallHTMLShowTopicID
    End Select
End Sub

'使能补票？废票
Public Function EnableExSell(ByVal pbEnable As Boolean) As Long
    abMenuTool.Bands("mnuFunction").Tools("mnuCancelTkt").Enabled = pbEnable
'    lblExtra.Enabled = pbEnable
    
    '--------------------
'    mnuCancelTkt.Enabled = pbEnable
'    lblCancel.Enabled = pbEnable
'    mnuReturnTkt.Enabled = pbEnable
'    lblReturn.Enabled = pbEnable
    
End Function

'从远程连接菜单中得到相应单位代码的单位名称
Private Function GetUnitNameFromMenu(pszUnitID As String) As String
    Dim i As Integer
    
    For i = 1 To 9 'mnuRemoteUnit.count
        If abMenuTool.Bands("mnuRemote").Tools("mnuRemoteUnit" & CStr(i)).TagVariant = pszUnitID Then
            GetUnitNameFromMenu = GetMenuUnitName(abMenuTool.Bands("mnuRemote").Tools("mnuRemoteUnit" & CStr(i)).Caption)
        End If
    Next
End Function

'使能选择排序和刷新菜单
Public Sub EnableSortAndRefresh(pbEnabled As Boolean)
'    Dim i As Integer
'    mnu_RefreshBus.Enabled = pbEnabled
'    mnu_RefreshStation.Enabled = pbEnabled
'    For i = 1 To mnuOrderBy.count
'        mnuOrderBy(i - 1).Enabled = pbEnabled
'    Next
End Sub

'得到当前单位的名称
Private Function GetUnitName() As String
    Dim i As Integer
    For i = 1 To 9
        If abMenuTool.Bands("mnuRemote").Tools("mnuRemoteUnit" & CStr(i)).Checked Then
            GetUnitName = abMenuTool.Bands("mnuRemote").Tools("mnuRemoteUnit" & CStr(i)).Caption
            Exit For
        End If
    Next
End Function


Private Sub ChangeSeatType()
    '切换座位类型
    On Error Resume Next
    '此处改成这样,是因为ActiveBar 将F9键给屏蔽了
'    If (ActiveForm Is frmSell) Or (ActiveForm Is frmExtraSell) Or (ActiveForm Is frmChangeSell) Then
        Me.ActiveForm.ChangeSeatType
        
'    End If
    
End Sub
'关联ActiveBar的控件
Private Sub AddControlsToActBar()
'    abMenuTool.Bands("bndTitleBar").Tools("ptTitle").Custom = ptTitle
    abMenuTool.Bands("statusBar").Tools("progressBar").Custom = pbLoad
End Sub

Private Sub SetCaption()


    Me.Caption = "售票台           " & Format(Date, "yyyy-mm-dd") & "  " & WeekdayName(Format(Date, "w")) ' & "    结束票号:" & GetEndTicketNo & "     "
    '剩余票数
    SetLeaveNum
    
End Sub

Public Sub SetLeaveNum()
    
    lblLeaveNum.Caption = m_lEndTicketNo - m_lTicketNo + 1
End Sub

Private Sub SellDetail()
    '售票明细查询
     frmSellDetail.Show vbModal, Me
   
    
    If frmSellDetail.m_bOk Then
        Dim frmTemp As IConditionForm
        Dim frmNewReport As New frmReport
        Set frmTemp = frmSellDetail
        frmNewReport.ShowReport frmTemp.RecordsetData, frmTemp.FileName, cszSellDetail, frmTemp.CustomData
       
    End If
End Sub

Private Sub BankStat()
    '网点代售统计
    frmBankStat.m_bDetail = False
     frmBankStat.Show vbModal, Me
   
    
    If frmBankStat.m_bOk And frmBankStat.m_bDetail = False Then
        Dim frmTemp As IConditionForm
        Dim frmNewReport As New frmReport
        Set frmTemp = frmBankStat
        frmNewReport.ShowReport frmTemp.RecordsetData, frmTemp.FileName, cszBankStat, frmTemp.CustomData
       
    End If

End Sub

Private Sub BankStatDetail()
    '网点代售统计详情
    On Error GoTo Error_Handle
      frmBankStat.m_bDetail = True
      frmBankStat.Show vbModal, Me
      If frmBankStat.m_bOk And frmBankStat.m_bDetail = True Then
        Dim rsSellDetail As Recordset
        Dim rsDetailToShow As Recordset
        Dim adbOther() As Double
'        Dim oDss As New TicketSellerDim
        Dim i As Integer, nUserCount As Integer
        
        Dim szLastTicketID As String
        Dim szBeginTicketID As String
        Dim arsData() As Recordset, vaCostumData As Variant
        
        Dim alNumber(TP_TicketTypeCount) As Long '各种票种的张数
        Dim adbAmount(TP_TicketTypeCount) As Double  '各种票种的金额
        Dim j As Integer
        Dim aszAllSeller() As String
        Dim nAllSeller As Integer
        Dim k As Integer
        Dim l As Integer
        
'        aszAllSeller = GetOperator()
'        nAllSeller = ArrayLength(aszAllSeller)
'
'
'        nUserCount = ArrayLength(aszAllSeller)
'
        Dim aszUserID() As String

        Dim pszUser As String
        nUserCount = ArrayLength(m_Sell)
        If nUserCount > 0 Then
            aszAllSeller = m_Sell
            nUserCount = ArrayLength(aszAllSeller)
            nAllSeller = ArrayLength(aszAllSeller)

        End If
        
        If nAllSeller > 0 Then
            
            ReDim arsData(1 To IIf(nAllSeller > nUserCount, nUserCount, nAllSeller))
            ReDim vaCostumData(1 To IIf(nAllSeller > nUserCount, nUserCount, nAllSeller), 1 To 8, 1 To 2)
            WriteProcessBar True, , nUserCount, "正在形成记录集..."
            l = 0
        
            
            For i = 1 To nUserCount
                k = nAllSeller
                If k <= nAllSeller Then
                    l = l + 1
                    '初始化
                    For j = 1 To TP_TicketTypeCount
                        alNumber(j) = 0
                        adbAmount(j) = 0
                    Next j
                
                    Set rsSellDetail = SellerEveryDaySellDetail(aszAllSeller(i), frmBankStat.m_dtWorkDate, frmBankStat.m_dtEndDate)
                    Set rsDetailToShow = New Recordset
                    With rsDetailToShow.Fields
                        .Append "ticket_id_range", adChar, 30
                        '往记录集中添加每种票种的数量与金额字段
                        For j = 1 To TP_TicketTypeCount
                            .Append "number_ticket_type" & j, adInteger
                            .Append "amount_ticket_type" & j, adCurrency
                        Next j
                    End With
     
                    rsDetailToShow.Open
                    Dim nTicketNumberLen As Integer
                    Dim nTicketPrefixLen As Integer
                    nTicketNumberLen = m_lTicketNoNumLen
                    nTicketPrefixLen = 0
                    
                    Do While Not rsSellDetail.EOF
                        If rsDetailToShow.RecordCount = 0 Or Not IsTicketIDSequence(szLastTicketID, RTrim(rsSellDetail!ID), nTicketNumberLen, nTicketPrefixLen) Then
                            If rsDetailToShow.RecordCount <> 0 Then
                                rsDetailToShow!ticket_id_range = szBeginTicketID & "---" & szLastTicketID
                                
                                For j = 1 To TP_TicketTypeCount
                                    alNumber(j) = alNumber(j) + rsDetailToShow("number_ticket_type" & j)
                                    adbAmount(j) = adbAmount(j) + rsDetailToShow("amount_ticket_type" & j)
                                Next j
                            End If
    
                            szBeginTicketID = RTrim(rsSellDetail!ID)
                            rsDetailToShow.AddNew
                        End If
                        rsDetailToShow("number_ticket_type" & rsSellDetail!TicketType) = rsDetailToShow("number_ticket_type" & rsSellDetail!TicketType) + 1
                        rsDetailToShow("amount_ticket_type" & rsSellDetail!TicketType) = rsDetailToShow("amount_ticket_type" & rsSellDetail!TicketType) + rsSellDetail!price
                        
                        szLastTicketID = RTrim(rsSellDetail!ID)
                        
                        rsSellDetail.MoveNext
                    Loop
                    
                    If rsSellDetail.RecordCount > 0 Then
                        rsSellDetail.MoveLast
                        rsDetailToShow!ticket_id_range = szBeginTicketID & "---" & RTrim(rsSellDetail!ID)
                        For j = 1 To TP_TicketTypeCount
                            alNumber(j) = alNumber(j) + rsDetailToShow("number_ticket_type" & j)
                            adbAmount(j) = adbAmount(j) + rsDetailToShow("amount_ticket_type" & j)
                        Next j
    
                        rsDetailToShow.AddNew
                        
                        rsDetailToShow!ticket_id_range = "合计"
                        For j = 1 To TP_TicketTypeCount
                            rsDetailToShow("number_ticket_type" & j) = alNumber(j)
                            rsDetailToShow("amount_ticket_type" & j) = adbAmount(j)
                        Next j
                        rsDetailToShow.Update
                    End If
                    Set arsData(l) = rsDetailToShow
                    adbOther = SellerEveryDayAnotherThing(aszAllSeller(i), frmBankStat.m_dtWorkDate, frmBankStat.m_dtEndDate)
                    vaCostumData(l, 1, 1) = "废票"
                    vaCostumData(l, 1, 2) = "张数=" & CInt(adbOther(1, 1)) & " 张  票款=" & adbOther(1, 2) & " 元"

                    
                    Dim dbAmount As Double '不包括免票
                    Dim lNumber As Long '包括免票
                    lNumber = 0
                    dbAmount = 0
                    For j = 1 To TP_TicketTypeCount
                        If j <> TP_FreeTicket Then
                            dbAmount = dbAmount + adbAmount(j)
                        End If
                        lNumber = lNumber + alNumber(j)
                    Next j
                        
                    vaCostumData(l, 2, 1) = "应交款"
                    vaCostumData(l, 2, 2) = dbAmount - adbOther(1, 2) & " 元"
                    
                    vaCostumData(l, 3, 1) = "总票数"
                    vaCostumData(l, 3, 2) = lNumber & " 张"
                    
                    vaCostumData(l, 4, 1) = "售票用票"
                    vaCostumData(l, 4, 2) = lNumber - adbOther(1, 1) & " 张"
                    
                    vaCostumData(l, 5, 1) = "制单"
                    vaCostumData(l, 5, 2) = MakeDisplayString(m_cszOperatorID, "")
                    
                    vaCostumData(l, 6, 1) = "复核"
                    vaCostumData(l, 6, 2) = ""
                    
                    vaCostumData(l, 7, 1) = "售票员"
                    vaCostumData(l, 7, 2) = aszAllSeller(i)
                    
                    vaCostumData(l, 8, 1) = "结算日期"
                    vaCostumData(l, 8, 2) = Format(frmBankStat.m_dtWorkDate, "yyyy年MM月DD日") & "—" & Format(frmBankStat.m_dtEndDate, "yyyy年MM月DD日")
                    
                End If
            Next
            
            Dim frmNewReport As New frmReport
            Dim frmTemp As IConditionForm
            Set frmTemp = frmBankStat
            frmNewReport.ShowReport2 arsData, "网点代售统计详情表模板.xls", cszSellerEveryDay, vaCostumData, 10
            
            WriteProcessBar False, , , ""
        End If
    End If
    Exit Sub
    Unload Me
Error_Handle:
    WriteProcessBar False, , , ""
    ShowErrorMsg

End Sub

Public Sub SetPrintEnabled(pbEnabled As Boolean)
    '设置菜单的可用性
    With abMenuTool
'        .Bands("tbn_system").Tools("tbn_system_print").Enabled = pbEnabled
'        .Bands("tbn_system").Tools("tbn_system_printview").Enabled = pbEnabled
'        .Bands("tbn_system").Tools("tbn_system_export").Enabled = pbEnabled
'        .Bands("tbn_system").Tools("tbn_system_exportopen").Enabled = pbEnabled
        .Bands("mnuSystem").Tools("mnu_PageOption").Enabled = pbEnabled
        .Bands("mnuSystem").Tools("mnu_PrintOption").Enabled = pbEnabled
        .Bands("mnuSystem").Tools("mnu_system_print").Enabled = pbEnabled
        .Bands("mnuSystem").Tools("mnu_system_printview").Enabled = pbEnabled
        .Bands("mnuSystem").Tools("mnu_ExportFile").Enabled = pbEnabled
        .Bands("mnuSystem").Tools("mnu_ExportFileOpen").Enabled = pbEnabled
    End With
End Sub


Public Function GetOperator() As String()

    Dim odb As New ADODB.Connection
    Dim rsTemp As Recordset
    Dim aszTemp() As String
    Dim szSql As String
    odb.ConnectionString = GetConnectionStr
    odb.CursorLocation = adUseClient
    odb.Open
    
    szSql = " select operatorid  as bank_id from tickets group by operatorid  "
    Set rsTemp = odb.Execute(szSql)
    aszTemp = GetUniqueTeam(rsTemp, aszTemp)
    GetOperator = aszTemp
    
End Function

'售票员环境时间售票详情
Public Function SellerEveryDaySellDetail(UserID As String, StartDate As Date, EndDate As Date) As Recordset
    Dim odb As New ADODB.Connection
    Dim rsTemp As Recordset
    Dim szSql As String
    odb.ConnectionString = GetConnectionStr
    odb.CursorLocation = adUseClient
    odb.Open
    
    szSql = "SELECT * FROM tickets t WHERE " _
    & " operatorid='" & UserID & "' and  t.status=1 AND " _
    & " selldate>='" & ToDBDateTime(StartDate) & "' AND " _
    & " selldate<'" & ToDBDateTime(DateAdd("d", 1, EndDate)) & "'" _
    & " ORDER BY id"
      
    
    Set rsTemp = odb.Execute(szSql)
    Set SellerEveryDaySellDetail = rsTemp
End Function

Public Function IsTicketIDSequence(pszFirstTicketID As String, pszSecondTicketID As String, nTicketNumberLen As Integer, nTicketPrefixLen As Integer) As Boolean
    Dim szTemp1 As String, szTemp2 As String
    On Error GoTo Error_Handle
    szTemp1 = UCase(Left(pszFirstTicketID, nTicketPrefixLen))
    szTemp2 = UCase(Left(pszSecondTicketID, nTicketPrefixLen))
    If szTemp1 = szTemp2 Then
        szTemp1 = Right(pszFirstTicketID, nTicketNumberLen)
        szTemp2 = Right(pszSecondTicketID, nTicketNumberLen)
        If CLng(szTemp1) + 1 = CLng(szTemp2) Then
            IsTicketIDSequence = True
        End If
    End If
    Exit Function
Error_Handle:
End Function


'数组相应为废、退、改
Public Function SellerEveryDayAnotherThing(UserID As String, StartDate As Date, EndDate As Date) As Double()
    Dim odb As New ADODB.Connection
    Dim adbResult(1 To 4, 1 To 3) As Double
    Dim rsTemp As Recordset
    Dim szSql As String
    odb.ConnectionString = GetConnectionStr
    odb.CursorLocation = adUseClient
    odb.Open

    szSql = "SELECT COUNT(*) AS countx ,SUM(price) AS total_ticket_price FROM " _
    & " tickets WHERE " _
    & " status =2  AND " _
    & " operatorid ='" & UserID & "' AND " _
    & " selldate>='" & ToDBDateTime(StartDate) & "' AND " _
    & " selldate<'" & ToDBDateTime(DateAdd("d", 1, EndDate)) & "'"
    
    
    Set rsTemp = odb.Execute(szSql)
    adbResult(1, 1) = rsTemp!countx
    adbResult(1, 2) = FormatDbValue(rsTemp!total_ticket_price)
    
    SellerEveryDayAnotherThing = adbResult
End Function


