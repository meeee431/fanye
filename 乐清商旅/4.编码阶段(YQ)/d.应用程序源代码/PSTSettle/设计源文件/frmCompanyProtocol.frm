VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "actbar2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Begin VB.Form frmCompanyProtocol 
   BackColor       =   &H00E0E0E0&
   Caption         =   "公司协议管理"
   ClientHeight    =   7650
   ClientLeft      =   465
   ClientTop       =   2550
   ClientWidth     =   10245
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Picture         =   "frmCompanyProtocol.frx":0000
   ScaleHeight     =   7650
   ScaleWidth      =   10245
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView lvCompanyProtocol 
      Height          =   5055
      Left            =   150
      TabIndex        =   3
      Top             =   1410
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   8916
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.PictureBox ptShowInfo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1155
      Left            =   -60
      ScaleHeight     =   1155
      ScaleWidth      =   15105
      TabIndex        =   0
      Top             =   0
      Width           =   15105
      Begin FText.asFlatTextBox txtRoute 
         Height          =   255
         Left            =   8340
         TabIndex        =   8
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
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
      Begin VB.ComboBox txtProtocolID 
         Height          =   300
         Left            =   5580
         TabIndex        =   7
         Top             =   375
         Width           =   1425
      End
      Begin FText.asFlatTextBox txtCompanyID 
         Height          =   285
         Left            =   3030
         TabIndex        =   6
         Top             =   390
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
      Begin RTComctl3.CoolButton cmdQuery 
         Default         =   -1  'True
         Height          =   360
         Left            =   10350
         TabIndex        =   1
         Top             =   330
         Width           =   1065
         _ExtentX        =   1879
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
         MICON           =   "frmCompanyProtocol.frx":1272
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "线路："
         Height          =   195
         Left            =   7410
         TabIndex        =   9
         Top             =   420
         Width           =   735
      End
      Begin VB.Label lblProtocolID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "协议号:"
         Height          =   180
         Left            =   4650
         TabIndex        =   4
         Top             =   420
         Width           =   600
      End
      Begin VB.Label lblCompanyID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "公司代码:"
         Height          =   180
         Left            =   1980
         TabIndex        =   2
         Top             =   480
         Width           =   810
      End
      Begin VB.Image Image2 
         Height          =   1275
         Left            =   60
         Picture         =   "frmCompanyProtocol.frx":128E
         Top             =   150
         Width           =   2010
      End
   End
   Begin ActiveBar2LibraryCtl.ActiveBar2 abAction 
      Height          =   5145
      Left            =   8340
      TabIndex        =   5
      Top             =   1350
      Width           =   1485
      _LayoutVersion  =   1
      _ExtentX        =   2619
      _ExtentY        =   9075
      _DataPath       =   ""
      Bands           =   "frmCompanyProtocol.frx":2761
   End
   Begin VB.Menu pmnu_Action 
      Caption         =   "操作"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu pmnu_Edit 
         Caption         =   "属性(&E)"
      End
      Begin VB.Menu pmnu_set 
         Caption         =   "协议设置(&S)"
      End
   End
End
Attribute VB_Name = "frmCompanyProtocol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_oReport As New Report
Dim m_oSplit As New Split

Private Sub abAction_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
On Error GoTo here
    Dim i As Integer, j As Integer
    Dim aszTemp() As String
    
    Select Case Tool.Caption
        Case "属性"
            EditCompanyProtocol
        Case "全选"
            If lvCompanyProtocol.ListItems.Count = 0 Then Exit Sub
            For i = 1 To lvCompanyProtocol.ListItems.Count
                lvCompanyProtocol.ListItems.Item(i).Checked = True
            Next i
        Case "删除"
            DeleteProtocol
        Case "取消协议"
            If lvCompanyProtocol.ListItems.Count = 0 Then Exit Sub
            ChangeProtocol
        Case "协议设置"
            SetCompanyProtocol
        Case "回程设置"
            SetBackCompanyProtocol
    End Select
    
    Exit Sub
here:
ShowErrorMsg
End Sub
Private Sub DeleteProtocol()
    On Error GoTo err
    m_oSplit.DeleteCompanyProtocol lvCompanyProtocol.SelectedItem.Text, ResolveDisplay(lvCompanyProtocol.SelectedItem.SubItems(2))
    lvCompanyProtocol.ListItems.Remove lvCompanyProtocol.SelectedItem.Index
err:
    Exit Sub
End Sub
Private Sub ChangeProtocol()
    On Error GoTo err
    Dim i As Integer, j As Integer
    Dim aszTemp() As String
    Dim aszTemp1() As String
    Dim szProtocol As String
    Dim vbYesOrNo As VbMsgBoxResult
    vbYesOrNo = MsgBox("取消协议：是将所选公司的协议设为默认协议，是否继续?", vbQuestion + vbYesNo + vbDefaultButton2, "协议管理")
    
    If vbYesOrNo = vbYes Then
        
        For i = 1 To lvCompanyProtocol.ListItems.Count
            If lvCompanyProtocol.ListItems.Item(i).Checked = True Then
                j = j + 1
            End If
        Next i
        If j = 0 Then
            MsgBox "你没有选择要取消协议的公司，请在选择后再取消协议", vbInformation, "协议管理"
            Exit Sub
        End If
        ReDim aszTemp(1 To j)
        j = 0
        For i = 1 To lvCompanyProtocol.ListItems.Count
            If lvCompanyProtocol.ListItems.Item(i).Checked = True Then
                j = j + 1
                aszTemp(j) = lvCompanyProtocol.ListItems.Item(i).Text
            End If
        Next i
        j = 0
        m_oSplit.Init g_oActiveUser
        m_oReport.Init g_oActiveUser
        aszTemp1 = m_oReport.GetAllProtocol
        For i = 1 To ArrayLength(aszTemp1)
            If aszTemp1(i, 4) = 0 Then
                szProtocol = aszTemp1(i, 1)
            End If
        Next i
'        m_oSplit.SetCompanyProtocol aszTemp, szProtocol
        FilllvCompany
    End If
    Exit Sub
err:
    ShowErrorMsg
End Sub

'Private Sub asFlatTextBox1_ButtonClick()
'On Error GoTo err
'    Dim oShell As New STShell.CommDialog
'    Dim aszTemp() As String
'    oShell.Init g_oActiveUser
'    aszTemp = oShell.SelectRoute()
'    Set oShell = Nothing
'    If ArrayLength(aszTemp) = 0 Then Exit Sub
'    txtCompanyID.Text = MakeDisplayString(Trim(aszTemp(1, 1)), Trim(aszTemp(1, 2)))
'    Exit Sub
'err:
'    ShowErrorMsg
'End Sub


Private Sub cmdQuery_Click()
    Query
End Sub
Private Sub Query()
    On Error GoTo err
    Dim nCount As Integer
    Dim aszTemp() As String
    Dim i As Integer
    Dim lvTemp As ListItem
    aszTemp = m_oReport.GetAllCompanyProtocol(ResolveDisplay(txtCompanyID.Text), ResolveDisplay(txtRoute.Text), ResolveDisplay(txtProtocolID))
    lvCompanyProtocol.ListItems.Clear
    nCount = ArrayLength(aszTemp)
    If ArrayLength(aszTemp) <> 0 Then
        
        For i = 1 To ArrayLength(aszTemp)
            Set lvTemp = lvCompanyProtocol.ListItems.Add(, , aszTemp(i, 1))
            lvTemp.ListSubItems.Add , , aszTemp(i, 2)
            lvTemp.ListSubItems.Add , , MakeDisplayString(aszTemp(i, 5), aszTemp(i, 6))
            lvTemp.ListSubItems.Add , , MakeDisplayString(aszTemp(i, 3), aszTemp(i, 4))
        Next i
    End If
    SetNormal
    lvCompanyProtocol.Refresh
    If lvCompanyProtocol.ListItems.Count > 0 Then lvCompanyProtocol.ListItems(1).Selected = True
    WriteProcessBar False
    ShowSBInfo "共" & nCount & "个对象", ESB_ResultCountInfo
    ShowSBInfo ""

    Exit Sub
err:
    ShowErrorMsg
End Sub
Private Sub Form_Load()
    m_oReport.Init g_oActiveUser
    m_oSplit.Init g_oActiveUser
    AlignForm
    FillHead
    FilllvCompany
    AlignHeadWidth Me.name, lvCompanyProtocol
    SortListView lvCompanyProtocol, 1
End Sub
Public Sub FilllvCompany(Optional pszCompanyID As String = "", Optional pszRouteId As String = "")
    On Error GoTo err
    Dim lvTemp As ListItem
    Dim nCount As Integer
    Dim aszTemp() As String
    Dim i As Integer
    Dim j As Integer
    m_oReport.Init g_oActiveUser
    aszTemp = m_oReport.GetAllCompanyProtocol(pszCompanyID, pszRouteId)
    nCount = ArrayLength(aszTemp)
    If pszCompanyID = "" And pszRouteId = "" Then
        lvCompanyProtocol.ListItems.Clear
    End If
    If nCount <> 0 Then
        For i = 1 To nCount
            '移去原先显示的
            For j = lvCompanyProtocol.ListItems.Count To 1 Step -1
                If lvCompanyProtocol.ListItems(j).Text = aszTemp(i, 1) Then
                    lvCompanyProtocol.ListItems.Remove j
                End If
            Next j
            '
            Set lvTemp = lvCompanyProtocol.ListItems.Add(, , aszTemp(i, 1))
            lvTemp.ListSubItems.Add , , aszTemp(i, 2)
            lvTemp.ListSubItems.Add , , MakeDisplayString(aszTemp(i, 5), aszTemp(i, 6))
            lvTemp.ListSubItems.Add , , MakeDisplayString(aszTemp(i, 3), aszTemp(i, 4))
            '设置显示及当前行
            If i = nCount Then
                lvTemp.Selected = True
                lvTemp.EnsureVisible
            End If
        Next i
    End If
    SetNormal
    lvCompanyProtocol.Refresh
    If lvCompanyProtocol.ListItems.Count > 0 And pszCompanyID = "" And pszRouteId = "" Then
        lvCompanyProtocol.ListItems(1).Selected = True
        lvCompanyProtocol.ListItems(1).EnsureVisible
    End If
    WriteProcessBar False
    ShowSBInfo "共" & nCount & "个对象", ESB_ResultCountInfo
    ShowSBInfo ""
    Exit Sub
err:
    ShowErrorMsg
End Sub

'界面排列
Private Sub AlignForm()
    On Error GoTo err
    ptShowInfo.Top = 0
    ptShowInfo.Left = 0
    ptShowInfo.Width = mdiMain.ScaleWidth
    
    lblCompanyID.Left = Image2.Width + 200
    lblCompanyID.Top = ptShowInfo.Height / 3 + 100
    txtCompanyID.Left = lblCompanyID.Left + lblCompanyID.Width + 100
    txtCompanyID.Top = ptShowInfo.Height / 3
    
'    lblCompanyName.Top = lblCompanyID.Top
'    lblCompanyName.Left = txtCompanyID.Left + txtCompanyID.Width + 150
'    txtCompanyName.Top = txtCompanyID.Top
'    txtCompanyName.Left = lblCompanyName.Left + lblCompanyName.Width + 100
'
'    lblProtocolID.Top = lblCompanyID.Top
'    lblProtocolID.Left = txtCompanyName.Left + txtCompanyName.Width + 150
'    txtProtocolID.Top = txtCompanyID.Top
'    txtProtocolID.Left = lblProtocolID.Left + lblProtocolID.Width + 100
''
'    lblProtocolName.Top = lblCompanyID.Top
'    lblProtocolName.Left = txtProtocolID.Left + txtProtocolID.Width + 150
'    txtProtocolName.Top = txtCompanyID.Top
'    txtProtocolName.Left = lblProtocolName.Left + lblProtocolName.Width + 100
'
'    cmdQuery.Left = txtProtocolName.Left + txtProtocolName.Width + 500
'    cmdQuery.Top = ptShowInfo.Height / 3
    
    lvCompanyProtocol.Top = ptShowInfo.Height + 50
    lvCompanyProtocol.Left = 50
    lvCompanyProtocol.Width = mdiMain.ScaleWidth - abAction.Width - 50
    lvCompanyProtocol.Height = mdiMain.ScaleHeight - ptShowInfo.Height - 50
    
    abAction.Top = lvCompanyProtocol.Top
    abAction.Left = lvCompanyProtocol.Width + 50
    abAction.Height = lvCompanyProtocol.Height
    Exit Sub
err:
    ShowErrorMsg
End Sub

Private Sub FillHead()
    Dim aszTemp() As String
    Dim i As Integer
    With lvCompanyProtocol.ColumnHeaders
        .Clear
        .Add , , "公司代码"
        .Add , , "参运公司"
        .Add , , "线路"
        .Add , , "协议号"
        
    End With
    With lvCompanyProtocol
        .ColumnHeaders(1).Width = 1000
        .ColumnHeaders(2).Width = 1000
        .ColumnHeaders(3).Width = 1800
        .ColumnHeaders(4).Width = 1500
    End With
    aszTemp = m_oReport.GetAllProtocol
    If ArrayLength(aszTemp) <> 0 Then
        For i = 1 To ArrayLength(aszTemp)
            txtProtocolID.AddItem MakeDisplayString(aszTemp(i, 1), aszTemp(i, 2))
        Next i
    End If
End Sub

Private Sub Form_Resize()
    AlignForm
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveHeadWidth Me.name, lvCompanyProtocol
    Unload Me
End Sub

Private Sub lvCompanyProtocol_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    SortListView lvCompanyProtocol, ColumnHeader.Index
End Sub

Private Sub lvCompanyProtocol_DblClick()
    EditCompanyProtocol
End Sub

Private Sub lvCompanyProtocol_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu pmnu_Action
    End If
End Sub

Private Sub pmnu_All_Click()
    Dim i As Integer
    For i = 1 To lvCompanyProtocol.ListItems.Count
        lvCompanyProtocol.ListItems.Item(i).Checked = True
    Next i
End Sub

Private Sub pmnu_clear_Click()
    ChangeProtocol
End Sub

Private Sub pmnu_edit_Click()
'    frmEditCompanyProtocol.m_AllSet = False
'    frmEditCompanyProtocol.ZOrder 0
'    frmEditCompanyProtocol.Show vbModal

            frmSetCompanyProtocol.ZOrder 0
            frmSetCompanyProtocol.Show vbModal
End Sub

Private Sub pmnu_notall_Click()
    Dim i As Integer
    For i = 1 To lvCompanyProtocol.ListItems.Count
        lvCompanyProtocol.ListItems.Item(i).Checked = False
    Next i
End Sub

Private Sub pmnu_set_Click()
'    frmEditCompanyProtocol.m_AllSet = True
'    frmEditCompanyProtocol.ZOrder 0
'    frmEditCompanyProtocol.Show vbModal
    frmSetCompanyProtocol.ZOrder 0
    frmSetCompanyProtocol.Show vbModal
End Sub




Private Sub txtRoute_ButtonClick()
    On Error GoTo err
    Dim oShell As New STShell.CommDialog
    Dim aszTemp() As String
    oShell.Init g_oActiveUser
    aszTemp = oShell.SelectRoute()
    Set oShell = Nothing
    If ArrayLength(aszTemp) = 0 Then Exit Sub
    txtRoute.Text = MakeDisplayString(Trim(aszTemp(1, 1)), Trim(aszTemp(1, 2)))
    Exit Sub
err:
    ShowErrorMsg
End Sub


Private Sub EditCompanyProtocol()
    frmSetCompanyProtocol.m_eFormStatus = ModifyStatus
    If Not (lvCompanyProtocol.SelectedItem Is Nothing) Then
        frmSetCompanyProtocol.m_szCompanyID = MakeDisplayString(lvCompanyProtocol.SelectedItem.Text, lvCompanyProtocol.SelectedItem.SubItems(1))
    End If
    frmSetCompanyProtocol.Show vbModal
 
End Sub

Private Sub SetCompanyProtocol()
    frmSetCompanyProtocol.m_eFormStatus = AddStatus
    frmSetCompanyProtocol.m_bIsBack = False
    frmSetCompanyProtocol.Show vbModal
End Sub

Private Sub SetBackCompanyProtocol()
    frmSetCompanyProtocol.m_eFormStatus = AddStatus
    frmSetCompanyProtocol.m_bIsBack = True
    frmSetCompanyProtocol.Show vbModal
End Sub

