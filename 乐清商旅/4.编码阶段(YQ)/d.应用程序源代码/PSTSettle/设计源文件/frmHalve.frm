VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "actbar2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Begin VB.Form frmHalve 
   BackColor       =   &H00E0E0E0&
   Caption         =   "加总平分"
   ClientHeight    =   7860
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13005
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7860
   ScaleWidth      =   13005
   WindowState     =   2  'Maximized
   Begin VB.PictureBox ptShowInfo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1155
      Left            =   0
      ScaleHeight     =   1155
      ScaleWidth      =   15135
      TabIndex        =   0
      Top             =   -30
      Width           =   15135
      Begin FText.asFlatTextBox txtCompanyOther 
         Height          =   315
         Left            =   5910
         TabIndex        =   9
         Top             =   345
         Width           =   1125
         _ExtentX        =   1984
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
      Begin FText.asFlatTextBox txtRouteID 
         Height          =   315
         Left            =   8190
         TabIndex        =   7
         Top             =   345
         Width           =   1125
         _ExtentX        =   1984
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
      Begin FText.asFlatTextBox txtCompanyID 
         Height          =   315
         Left            =   3210
         TabIndex        =   6
         Top             =   345
         Width           =   1125
         _ExtentX        =   1984
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
      Begin RTComctl3.CoolButton cmdQuery 
         Default         =   -1  'True
         Height          =   345
         Left            =   10050
         TabIndex        =   1
         Top             =   330
         Width           =   1065
         _ExtentX        =   1879
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
         MICON           =   "frmHalve.frx":0000
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
         Caption         =   "对方公司"
         Height          =   225
         Left            =   4950
         TabIndex        =   8
         Top             =   420
         Width           =   915
      End
      Begin VB.Label lblRoute 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "线路(&N):"
         Height          =   180
         Left            =   7170
         TabIndex        =   3
         Top             =   480
         Width           =   720
      End
      Begin VB.Image Image2 
         Height          =   1275
         Left            =   60
         Picture         =   "frmHalve.frx":001C
         Top             =   150
         Width           =   2010
      End
      Begin VB.Label lbllicense 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "参运公司(&P):"
         Height          =   180
         Left            =   2070
         TabIndex        =   2
         Top             =   480
         Width           =   1080
      End
   End
   Begin MSComctlLib.ListView lvHavle 
      Height          =   5055
      Left            =   180
      TabIndex        =   4
      Top             =   1680
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   8916
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin ActiveBar2LibraryCtl.ActiveBar2 abAction 
      Height          =   4845
      Left            =   8160
      TabIndex        =   5
      Top             =   1470
      Width           =   1485
      _LayoutVersion  =   1
      _ExtentX        =   2619
      _ExtentY        =   8546
      _DataPath       =   ""
      Bands           =   "frmHalve.frx":14EF
   End
   Begin VB.Menu mnu_Action 
      Caption         =   "操作"
      Visible         =   0   'False
      Begin VB.Menu mnu_Action_Edit 
         Caption         =   "属性(&E)"
      End
      Begin VB.Menu mnu_Action_Add 
         Caption         =   "新增(&A)"
      End
      Begin VB.Menu mnu_Action_Delete 
         Caption         =   "删除(&D)"
      End
   End
End
Attribute VB_Name = "frmHalve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'界面排列
'Private Sub AlignForm()
'
'    ptShowInfo.Top = 0
'    ptShowInfo.Left = 0
'    ptShowInfo.Width = Screen.Width
'
'    lvHavle.Top = ptShowInfo.Height + 50
'    lvHavle.Left = 50
'    lvHavle.Width = Screen.Width - abAction.Width - 50
'    lvHavle.Height = Screen.Width - ptShowInfo.Height - 50
'
'    abAction.Top = lvHavle.Top
'    abAction.Left = lvHavle.Width + 50
'    abAction.Height = lvHavle.Height
'
'End Sub


Private Sub Form_Load()

'    AlignForm
    AlignHeadWidth Me.Caption, lvHavle
    FillHead
    If lvHavle.ListItems.Count > 0 Then
        SetActioinEnalbled True
    Else
        SetActioinEnalbled False
    End If
    FilllvHavle
End Sub

Private Sub SetActioinEnalbled(bEnabled As Boolean)
    
        abAction.Bands("bndActionTabs").ChildBands("actBaseMan").Tools("act_BaseMan_Add").Enabled = True
        abAction.Bands("bndActionTabs").ChildBands("actBaseMan").Tools("act_BaseMan_BaseInfo").Enabled = bEnabled
        abAction.Bands("bndActionTabs").ChildBands("actBaseMan").Tools("act_BaseMan_Del").Enabled = bEnabled
        
        mnu_Action_Add.Enabled = True
        mnu_Action_Edit.Enabled = bEnabled
        mnu_Action_Delete.Enabled = bEnabled
End Sub

'查询加总平分
Public Sub FilllvHavle()
On Error GoTo here
    Dim szTemp() As String
    Dim m_oReport As New Report
    Dim i As Integer
    Dim mCount As Integer
    Dim lvItem As ListItem
    Dim szCompanyID As String
    Dim szRouteID As String
    Dim szCompnayOther As String
    If txtCompanyID.Text <> "" Then
        szCompanyID = Trim(txtCompanyID.Text)
    Else
        szCompanyID = ""
    End If
   
    If txtRouteID.Text <> "" Then
        szRouteID = Trim(txtRouteID.Text)
    Else
        szRouteID = ""
    End If
    If txtCompanyOther.Text <> "" Then
        szCompnayOther = Trim(txtCompanyOther.Text)
    Else
        szCompnayOther = ""
    End If
        
    
    '查询接口
    m_oReport.Init g_oActiveUser
    szTemp = m_oReport.GetAllHalveCompany(ResolveDisplay(szRouteID), ResolveDisplay(szCompanyID), ResolveDisplay(szCompnayOther))
    '填充时以0001[****] 方式
     mCount = ArrayLength(szTemp)
     lvHavle.ListItems.Clear
     
    If mCount = 0 Then Exit Sub
    For i = 1 To mCount
        Set lvItem = lvHavle.ListItems.Add(, , MakeDisplayString(szTemp(i, 1), szTemp(i, 2)))
        lvItem.SubItems(1) = MakeDisplayString(szTemp(i, 3), szTemp(i, 4))
        lvItem.SubItems(2) = MakeDisplayString(szTemp(i, 5), szTemp(i, 6))
        lvItem.SubItems(3) = szTemp(i, 7)
    Next i
    
    If lvHavle.ListItems.Count > 0 Then
        SetActioinEnalbled True
    Else
        SetActioinEnalbled False
    End If
    SetNormal
    lvHavle.Refresh
    If lvHavle.ListItems.Count > 0 Then lvHavle.ListItems(1).Selected = True
    WriteProcessBar False
    ShowSBInfo "共" & mCount & "个对象", ESB_ResultCountInfo
    ShowSBInfo ""

    Exit Sub
here:
    ShowErrorMsg
End Sub

Private Sub Form_Resize()
    ptShowInfo.Top = 0
    ptShowInfo.Left = 0
    ptShowInfo.Width = Me.ScaleWidth

    lvHavle.Top = ptShowInfo.Height + 50
    lvHavle.Left = 0
    lvHavle.Height = Me.ScaleHeight - ptShowInfo.Height - 50
    lvHavle.Width = Me.ScaleWidth - abAction.Width - 50
    
    abAction.Top = lvHavle.Top
    abAction.Left = lvHavle.Width + 50
    abAction.Height = lvHavle.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveHeadWidth Me.name, lvHavle
    Unload Me
End Sub

Private Sub cmdQuery_Click()
On Error GoTo ErrHandle
        
    '查询并填充列表
    FilllvHavle
Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub FillHead()
    With lvHavle.ColumnHeaders
        .Clear
        .Add , , "线路"
        .Add , , "参运公司"
        .Add , , "对方公司"
        .Add , , "分成比率"
        
    End With
    With lvHavle
        .ColumnHeaders(1).Width = 2000
        .ColumnHeaders(2).Width = 1000
        .ColumnHeaders(3).Width = 2000
        .ColumnHeaders(4).Width = 1000
    End With
    AlignHeadWidth Me.name, lvHavle
End Sub
Private Sub abAction_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    Select Case Tool.Caption
        Case "属性"
            EditObject
        Case "新增"
            AddObject
        Case "删除"
            DeleteObject
    End Select
End Sub
Private Sub EditObject()
    
    frmHalveEdit.m_status = ModifyStatus
    frmHalveEdit.ZOrder 0
    frmHalveEdit.Show vbModal
End Sub

Private Sub AddObject()
    frmHalveEdit.m_status = AddStatus
    frmHalveEdit.ZOrder 0
    frmHalveEdit.Show vbModal
    FilllvHavle
End Sub

Private Sub DeleteObject()
    Dim m_Answer
    Dim m_oHalve As New HalveCompany
    m_Answer = MsgBox("你是否确认删除此加总平分的公司?", vbInformation + vbYesNo, Me.Caption)
    If m_Answer = vbYes Then
'        For i = 1 To lvHavle.SelectedItem.Index
            m_oHalve.Init g_oActiveUser
            m_oHalve.RouteID = ResolveDisplay(lvHavle.SelectedItem.Text)
            m_oHalve.CompanyID = ResolveDisplay(lvHavle.SelectedItem.SubItems(1))
            
            m_oHalve.Delete
'        Next i
        FilllvHavle
    End If
    
End Sub



Private Sub lvHavle_DblClick()
    EditObject
End Sub

Private Sub lvHavle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnu_Action
    End If
End Sub

Private Sub mnu_Action_Add_Click()
    AddObject
End Sub

Private Sub mnu_Action_Delete_Click()
    DeleteObject
End Sub

Private Sub mnu_Action_Edit_Click()
    EditObject
End Sub

Private Sub txtCompanyID_ButtonClick()
On Error GoTo ErrHandle
    Dim oShell As New STShell.CommDialog
    Dim aszTemp() As String
    oShell.Init g_oActiveUser
    aszTemp = oShell.SelectCompany()
    Set oShell = Nothing
    If ArrayLength(aszTemp) = 0 Then Exit Sub
    txtCompanyID.Text = MakeDisplayString(Trim(aszTemp(1, 1)), Trim(aszTemp(1, 2)))

Exit Sub
ErrHandle:
ShowErrorMsg
End Sub

Private Sub txtCompanyOther_ButtonClick()
    Dim oShell As New STShell.CommDialog
    Dim aszTemp() As String
    oShell.Init g_oActiveUser
    aszTemp = oShell.SelectCompany
    Set oShell = Nothing
    If ArrayLength(aszTemp) = 0 Then Exit Sub
    txtCompanyOther.Text = MakeDisplayString(Trim(aszTemp(1, 1)), Trim(aszTemp(1, 2)))
End Sub

Private Sub txtRouteID_ButtonClick()
On Error GoTo ErrHandle
    Dim oShell As New STShell.CommDialog
    Dim aszTemp() As String
    oShell.Init g_oActiveUser
    aszTemp = oShell.SelectRoute
    Set oShell = Nothing
    If ArrayLength(aszTemp) = 0 Then Exit Sub
    txtRouteID.Text = MakeDisplayString(Trim(aszTemp(1, 1)), Trim(aszTemp(1, 2)))

Exit Sub
ErrHandle:
ShowErrorMsg
End Sub

