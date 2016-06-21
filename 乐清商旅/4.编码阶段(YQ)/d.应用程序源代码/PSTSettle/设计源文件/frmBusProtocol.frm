VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "actbar2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Begin VB.Form frmBusProtocol 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "车次协议管理"
   ClientHeight    =   8400
   ClientLeft      =   540
   ClientTop       =   2235
   ClientWidth     =   12240
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8400
   ScaleWidth      =   12240
   WindowState     =   2  'Maximized
   Begin RTComctl3.CoolButton cmdQuery 
      Default         =   -1  'True
      Height          =   345
      Left            =   10170
      TabIndex        =   11
      Top             =   345
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
      MICON           =   "frmBusProtocol.frx":0000
      PICN            =   "frmBusProtocol.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.PictureBox ptShowInfo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1155
      Left            =   0
      ScaleHeight     =   1155
      ScaleWidth      =   15105
      TabIndex        =   1
      Top             =   0
      Width           =   15105
      Begin VB.ComboBox txtProtocol 
         Height          =   300
         Left            =   3315
         TabIndex        =   7
         Top             =   750
         Width           =   2610
      End
      Begin VB.TextBox txtBusID 
         Height          =   285
         Left            =   3315
         TabIndex        =   6
         Top             =   368
         Width           =   870
      End
      Begin FText.asFlatTextBox txtCompanyID 
         Height          =   285
         Left            =   5655
         TabIndex        =   2
         Top             =   375
         Width           =   1215
         _ExtentX        =   2143
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
         ButtonHotBackColor=   -2147483633
         ButtonPressedBackColor=   -2147483627
         Text            =   ""
         ButtonBackColor =   -2147483633
         ButtonVisible   =   -1  'True
         Registered      =   -1  'True
      End
      Begin FText.asFlatTextBox txtRoute 
         Height          =   315
         Left            =   7920
         TabIndex        =   9
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
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
         ButtonHotBackColor=   -2147483633
         ButtonPressedBackColor=   -2147483627
         Text            =   ""
         ButtonBackColor =   -2147483633
         ButtonVisible   =   -1  'True
         Registered      =   -1  'True
      End
      Begin VB.Label lblVehicleID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "线路(&R):"
         Height          =   180
         Left            =   7170
         TabIndex        =   10
         Top             =   420
         Width           =   720
      End
      Begin VB.Label lblProtocolID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "协议号(&P):"
         Height          =   180
         Left            =   2160
         TabIndex        =   8
         Top             =   810
         Width           =   900
      End
      Begin VB.Label lblBusID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车次(&B):"
         Height          =   180
         Left            =   2340
         TabIndex        =   5
         Top             =   420
         Width           =   720
      End
      Begin VB.Label lblCompanyID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "公司代码(&C):"
         Height          =   180
         Left            =   4545
         TabIndex        =   3
         Top             =   420
         Width           =   1080
      End
      Begin VB.Image Image2 
         Height          =   1275
         Left            =   60
         Picture         =   "frmBusProtocol.frx":03B6
         Top             =   150
         Width           =   2010
      End
   End
   Begin MSComctlLib.ListView lvBusProtocol 
      Height          =   5055
      Left            =   210
      TabIndex        =   0
      Top             =   1410
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   8916
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin ActiveBar2LibraryCtl.ActiveBar2 abAction 
      Height          =   5145
      Left            =   8400
      TabIndex        =   4
      Top             =   1350
      Width           =   1485
      _LayoutVersion  =   1
      _ExtentX        =   2619
      _ExtentY        =   9075
      _DataPath       =   ""
      Bands           =   "frmBusProtocol.frx":1889
   End
End
Attribute VB_Name = "frmBusProtocol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private m_oReport As New Report
Private m_oSplit As New Split
Private m_oFormular As New Formular

'界面排列
Private Sub AlignForm()
    On Error GoTo err
    ptShowInfo.Top = 0
    ptShowInfo.Left = 0
    ptShowInfo.Width = mdiMain.ScaleWidth
      
    lvBusProtocol.Top = ptShowInfo.Height + 50
    lvBusProtocol.Left = 50
    lvBusProtocol.Width = mdiMain.ScaleWidth - abAction.Width - 50
    lvBusProtocol.Height = mdiMain.ScaleHeight - ptShowInfo.Height - 50
    
    abAction.Top = lvBusProtocol.Top
    abAction.Left = lvBusProtocol.Width + 50
    abAction.Height = lvBusProtocol.Height
    Exit Sub
err:
    ShowErrorMsg
End Sub

Private Sub Form_Resize()
    AlignForm
End Sub

Private Sub lvBusProtocol_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    SortListView lvBusProtocol, ColumnHeader.Index
End Sub

Private Sub lvBusProtocol_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbRightButton Then
'        PopupMenu pmnu_Action
    End If
End Sub

Private Sub pmnu_Add_Click()

End Sub

Private Sub pmnu_All_Click()
    Dim i As Integer
    For i = 1 To lvBusProtocol.ListItems.Count
        lvBusProtocol.ListItems.Item(i).Checked = True
    Next i
End Sub

Private Sub pmnu_clear_Click()
ChangeProtocol
End Sub

Private Sub pmnu_edit_Click()
    frmEditBusProtocol.ZOrder 0
    frmEditBusProtocol.m_AllSet = False
    frmEditBusProtocol.Show vbModal
End Sub

Private Sub pmnu_notall_Click()
    Dim i As Integer
    For i = 1 To lvBusProtocol.ListItems.Count
        lvBusProtocol.ListItems.Item(i).Checked = False
    Next i
End Sub

Private Sub pmnu_set_Click()
    frmEditBusProtocol.ZOrder 0
    frmEditBusProtocol.m_AllSet = True
    frmEditBusProtocol.Show vbModal
End Sub

Private Sub txtBus_ButtonClick()
    On Error GoTo err
    Dim oShell As New STShell.CommDialog
    Dim aszTemp() As String
    On Error GoTo err
    oShell.Init g_oActiveUser
    aszTemp = oShell.SelectBus()
    Set oShell = Nothing
    If ArrayLength(aszTemp) = 0 Then Exit Sub
    txtBusID.Text = MakeDisplayString(Trim(aszTemp(1, 1)), Trim(aszTemp(1, 2)))
    Exit Sub
err:
    ShowErrorMsg
End Sub



Private Sub cmdQuery_Click()
    Query
End Sub
Private Sub Query()
    On Error GoTo err
    Dim nCount As Integer
    Dim i As Integer
    Dim lvTemp As ListItem
    Dim rsTemp As Recordset
    lvBusProtocol.ListItems.Clear
    If txtBusID.Text = "" And txtCompanyID.Text = "" And txtRoute.Text = "" And txtProtocol.Text = "" Then
        FillLvBusProtocol
    Else
        Set rsTemp = m_oReport.GetAllBusProtocol(ResolveDisplay(txtBusID.Text), ResolveDisplay(txtCompanyID.Text), ResolveDisplay(txtRoute.Text), ResolveDisplay(txtProtocol.Text))
        nCount = rsTemp.RecordCount
        If nCount <> 0 Then
            
            For i = 1 To nCount
                Set lvTemp = lvBusProtocol.ListItems.Add(, , FormatDbValue(rsTemp!bus_id))
                lvTemp.ListSubItems.Add , , "" ' FormatDbValue(rsTemp!route_name)
                lvTemp.ListSubItems.Add , , MakeDisplayString(FormatDbValue(rsTemp!transport_company_id), FormatDbValue(rsTemp!transport_company_short_name))
                lvTemp.ListSubItems.Add , , MakeDisplayString(FormatDbValue(rsTemp!protocol_id), FormatDbValue(rsTemp!protocol_name))
                lvTemp.ListSubItems.Add , , GetDefaultMark(FormatDbValue(rsTemp!default_mark))
                rsTemp.MoveNext
            Next i
        End If
        
        
        '填充其他的车次信息
        Set rsTemp = m_oReport.GetOtherBusProtocol(ResolveDisplay(txtBusID.Text), ResolveDisplay(txtCompanyID.Text), ResolveDisplay(txtProtocol.Text))
        nCount = rsTemp.RecordCount
        If nCount > 0 Then
            
            For i = 1 To nCount
                Set lvTemp = lvBusProtocol.ListItems.Add(, , FormatDbValue(rsTemp!bus_id))
                lvTemp.ListSubItems.Add , , ""
                lvTemp.ListSubItems.Add , , FormatDbValue(rsTemp!transport_company_id)
                lvTemp.ListSubItems.Add , , FormatDbValue(rsTemp!protocol_id)
                rsTemp.MoveNext
            Next i
        End If
        
    End If
    SetNormal
    lvBusProtocol.Refresh
    If lvBusProtocol.ListItems.Count > 0 Then lvBusProtocol.ListItems(1).Selected = True
    WriteProcessBar False
    ShowSBInfo "共" & nCount & "个对象", ESB_ResultCountInfo
    ShowSBInfo ""

    Exit Sub
err:
ShowErrorMsg
End Sub

Private Sub Form_Load()
On Error GoTo err
    m_oReport.Init g_oActiveUser
    AlignForm
    FillHead
    FillLvBusProtocol
    m_oFormular.Init g_oActiveUser
    AlignHeadWidth Me.name, lvBusProtocol
    SortListView lvBusProtocol, 1
    Exit Sub
err:
    ShowErrorMsg
End Sub

Private Sub FillLvBusProtocol(Optional pszBusID As String = "")
    On Error GoTo err
    Dim rsBusCompany As Recordset
    Dim lvTemp As ListItem
    Dim rsBusProtocol As Recordset
    Dim nBusCount As Long
    Dim i As Long
    Dim j As Long
    lvBusProtocol.ListItems.Clear
    
    m_oReport.Init g_oActiveUser
    Set rsBusCompany = m_oReport.GetAllBusCompany
    nBusCount = rsBusCompany.RecordCount
    If nBusCount = 0 Then Exit Sub
    
                
    Set rsBusProtocol = m_oReport.GetAllBusProtocol
    
    For i = 1 To nBusCount
        Set lvTemp = lvBusProtocol.ListItems.Add(, , FormatDbValue(rsBusCompany!bus_id))
        lvTemp.ListSubItems.Add , , MakeDisplayString(FormatDbValue(rsBusCompany!route_id), FormatDbValue(rsBusCompany!route_name))
        lvTemp.ListSubItems.Add , , MakeDisplayString(FormatDbValue(rsBusCompany!transport_company_id), FormatDbValue(rsBusCompany!transport_company_short_name))
        
        lvTemp.ListSubItems.Add , , ""
        lvTemp.ListSubItems.Add , , "未设置协议"
        rsBusCompany.MoveNext
    Next i
    
    For i = 1 To rsBusProtocol.RecordCount
        For j = 1 To lvBusProtocol.ListItems.Count
            With lvBusProtocol.ListItems(j)
                If .Text = FormatDbValue(rsBusProtocol!bus_id) And ResolveDisplay(.ListSubItems(2)) = FormatDbValue(rsBusProtocol!transport_company_id) Then
                    .ListSubItems(3) = MakeDisplayString(FormatDbValue(rsBusProtocol!protocol_id), FormatDbValue(rsBusProtocol!protocol_name))
                    .ListSubItems(4) = IIf(FormatDbValue(rsBusProtocol!bus_id) = "", "未设置协议", GetDefaultMark(FormatDbValue(rsBusProtocol!default_mark)))
                    Exit For
                End If
            End With
        Next j
        If j > lvBusProtocol.ListItems.Count Then
            '表示未找到该车次
            Set lvTemp = lvBusProtocol.ListItems.Add(, , FormatDbValue(rsBusProtocol!bus_id))
            lvTemp.ListSubItems.Add , , ""
            lvTemp.ListSubItems.Add , , MakeDisplayString(FormatDbValue(rsBusProtocol!transport_company_id), FormatDbValue(rsBusProtocol!transport_company_short_name))
            
            lvTemp.ListSubItems.Add , , MakeDisplayString(FormatDbValue(rsBusProtocol!protocol_id), FormatDbValue(rsBusProtocol!protocol_name))
            lvTemp.ListSubItems.Add , , IIf(FormatDbValue(rsBusProtocol!bus_id) = "", "未设置协议", GetDefaultMark(FormatDbValue(rsBusProtocol!default_mark)))
        End If
        rsBusProtocol.MoveNext
        
    Next i
    Dim rsTemp As New Recordset
    Dim nCount As Integer
    '填充其他的车次信息
    Set rsTemp = m_oReport.GetOtherBusProtocol()
    nCount = rsTemp.RecordCount
    If nCount > 0 Then
        
        For i = 1 To nCount
            Set lvTemp = lvBusProtocol.ListItems.Add(, , FormatDbValue(rsTemp!bus_id))
            lvTemp.ListSubItems.Add , , ""
            lvTemp.ListSubItems.Add , , FormatDbValue(rsTemp!transport_company_id)
            lvTemp.ListSubItems.Add , , FormatDbValue(rsTemp!protocol_id)
            rsTemp.MoveNext
        Next i
    End If
    
    
    WriteProcessBar False
    ShowSBInfo "共" & nBusCount & "个对象", ESB_ResultCountInfo
    ShowSBInfo ""
    Exit Sub
err:
    ShowErrorMsg
End Sub


Private Sub Form_Unload(Cancel As Integer)
    SaveHeadWidth Me.name, lvBusProtocol
    Unload Me
End Sub

Private Sub FillHead()
    On Error GoTo err
    Dim aszTemp() As String
    Dim i As Integer
    With lvBusProtocol.ColumnHeaders
        .Clear
        .Add , , "车次代码"
        .Add , , "线路名称"
        .Add , , "参运公司"
        .Add , , "协议名称"
        .Add , , "是否默认协议"
    End With
    With lvBusProtocol
        .ColumnHeaders(1).Width = 1000
        .ColumnHeaders(2).Width = 2500
        .ColumnHeaders(3).Width = 1500
        .ColumnHeaders(4).Width = 2000
        .ColumnHeaders(5).Width = 1500
    End With

    aszTemp = m_oReport.GetAllProtocol
    If ArrayLength(aszTemp) <> 0 Then
        For i = 1 To ArrayLength(aszTemp)
            txtProtocol.AddItem MakeDisplayString(aszTemp(i, 1), aszTemp(i, 2))
        Next i
    End If
    Exit Sub
err:
    ShowErrorMsg
End Sub

Private Sub abAction_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
On Error GoTo here
    Dim i As Integer
    If lvBusProtocol.ListItems.Count = 0 Then Exit Sub
    Select Case Tool.Caption
        Case "属性"
            ModifyBusProtocol False
        Case "协议设置"
            ModifyBusProtocol False
        Case "取消协议"
            ChangeProtocol
        Case "其他车次设置"
            ModifyBusProtocol True
    End Select
    
    Exit Sub
here:
ShowErrorMsg
End Sub
Private Sub ChangeProtocol()
    On Error GoTo err
    Dim vbYesOrNo As VbMsgBoxResult
    
        vbYesOrNo = MsgBox("取消协议：是将所选车次的协议取消，是否继续?", vbQuestion + vbYesNo + vbDefaultButton2, "协议管理")
        If vbYesOrNo = vbYes Then
            m_oSplit.DelBusProtocol lvBusProtocol.SelectedItem.Text
'            lvBusProtocol.SelectedItem.SubItems(3) = ""
'            lvBusProtocol.SelectedItem.SubItems(4) = "未设置协议"
            lvBusProtocol.ListItems.Remove lvBusProtocol.SelectedItem.Index
        End If

    Exit Sub
err:
ShowErrorMsg
End Sub
Private Sub lblProtocolName_Click()

End Sub

Private Sub lvBusProtocol_DblClick()
    ModifyBusProtocol False
    
    
End Sub

Private Sub txtCompanyID_ButtonClick()
    Dim oShell As New STShell.CommDialog
    Dim aszTemp() As String
    On Error GoTo err
    oShell.Init g_oActiveUser
    aszTemp = oShell.SelectCompany()
    Set oShell = Nothing
    If ArrayLength(aszTemp) = 0 Then Exit Sub
    txtCompanyID.Text = MakeDisplayString(Trim(aszTemp(1, 1)), Trim(aszTemp(1, 2)))
    
    Exit Sub
err:
    ShowErrorMsg
End Sub


Private Sub txtRoute_ButtonClick()
    Dim oShell As New STShell.CommDialog
    Dim aszTemp() As String
    On Error GoTo err
    oShell.Init g_oActiveUser
    aszTemp = oShell.SelectRoute(False)
    Set oShell = Nothing
    If ArrayLength(aszTemp) = 0 Then Exit Sub
    txtRoute.Text = aszTemp(1, 1) & "[" & Trim(aszTemp(1, 2)) & "]"
    Exit Sub
err:
    ShowErrorMsg
End Sub

Private Sub ModifyBusProtocol(pbOtherBus As Boolean)

'    Dim aszTemp() As String
    If lvBusProtocol.SelectedItem Is Nothing Then Exit Sub
    frmEditBusProtocol.m_bOtherBus = pbOtherBus
    frmEditBusProtocol.m_AllSet = False
    frmEditBusProtocol.Show vbModal
    If Not pbOtherBus Then
        ModifyList pbOtherBus, lvBusProtocol.SelectedItem.Text, ResolveDisplay(lvBusProtocol.SelectedItem.SubItems(2))
    Else
        
        ModifyList pbOtherBus, frmEditBusProtocol.m_szBusID
    End If
End Sub

Public Sub ModifyList(pbOtherBus As Boolean, pszBusID As String, Optional pszCompanyID As String)
    Dim lvTemp As ListItem
    Dim nCount As Integer
    Dim i As Integer
    Dim rsTemp As Recordset
    Set rsTemp = m_oReport.GetAllBusProtocol(pszBusID, pszCompanyID)
    nCount = rsTemp.RecordCount
    If nCount > 0 Then
        If Not pbOtherBus Then
            Set lvTemp = lvBusProtocol.SelectedItem
            lvTemp.SubItems(3) = MakeDisplayString(FormatDbValue(rsTemp!protocol_id), FormatDbValue(rsTemp!protocol_name))
            lvTemp.SubItems(4) = GetDefaultMark(FormatDbValue(rsTemp!default_mark))
        Else
            Set lvTemp = lvBusProtocol.ListItems.Add(, , FormatDbValue(rsTemp!bus_id))
            lvTemp.ListSubItems.Add , , "" 'FormatDbValue(rsTemp!route_name)
            lvTemp.ListSubItems.Add , , MakeDisplayString(FormatDbValue(rsTemp!transport_company_id), FormatDbValue(rsTemp!transport_company_short_name))
            lvTemp.ListSubItems.Add , , MakeDisplayString(FormatDbValue(rsTemp!protocol_id), FormatDbValue(rsTemp!protocol_name))
            lvTemp.ListSubItems.Add , , GetDefaultMark(FormatDbValue(rsTemp!default_mark))
        
            
        End If
    End If
    
    '填充其他的车次信息
    Set rsTemp = m_oReport.GetOtherBusProtocol(pszBusID)
    nCount = rsTemp.RecordCount
    If nCount > 0 And pbOtherBus Then
        
        For i = 1 To nCount
            Set lvTemp = lvBusProtocol.ListItems.Add(, , FormatDbValue(rsTemp!bus_id))
            lvTemp.ListSubItems.Add , , ""
            lvTemp.ListSubItems.Add , , FormatDbValue(rsTemp!transport_company_id)
            lvTemp.ListSubItems.Add , , FormatDbValue(rsTemp!protocol_id)
            rsTemp.MoveNext
        Next i
    End If
        

    Exit Sub
err:
    ShowErrorMsg
End Sub

