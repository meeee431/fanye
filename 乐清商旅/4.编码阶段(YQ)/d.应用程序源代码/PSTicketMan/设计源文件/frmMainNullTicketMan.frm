VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "actbar2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmManNullTicketMan 
   BackColor       =   &H00E0E0E0&
   Caption         =   "空白票管理"
   ClientHeight    =   6570
   ClientLeft      =   1140
   ClientTop       =   2595
   ClientWidth     =   11205
   HelpContextID   =   2001801
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6570
   ScaleWidth      =   11205
   WindowState     =   2  'Maximized
   Begin VB.PictureBox ptShowInfo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1350
      Left            =   120
      ScaleHeight     =   1350
      ScaleWidth      =   10815
      TabIndex        =   3
      Top             =   30
      Width           =   10815
      Begin VB.TextBox txtSeller 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   7110
         MaxLength       =   5
         TabIndex        =   7
         Top             =   375
         Width           =   1980
      End
      Begin MSComCtl2.DTPicker dtpGetNullTicketEndDate 
         Height          =   315
         Left            =   3840
         TabIndex        =   5
         Top             =   825
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   556
         _Version        =   393216
         Format          =   22872064
         CurrentDate     =   37942
      End
      Begin MSComCtl2.DTPicker dtpGetNullTicketStartDate 
         Height          =   315
         Left            =   3840
         TabIndex        =   9
         Top             =   368
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   556
         _Version        =   393216
         Format          =   22872064
         CurrentDate     =   37942
      End
      Begin RTComctl3.CoolButton cmdFind 
         Default         =   -1  'True
         Height          =   345
         Left            =   9840
         TabIndex        =   2
         Top             =   360
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
         MICON           =   "frmMainNullTicketMan.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "空白票登记起始时间(&C):"
         Height          =   180
         Left            =   1860
         TabIndex        =   8
         Top             =   435
         Width           =   1980
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "空白票登记结束时间(&C):"
         Height          =   180
         Left            =   1860
         TabIndex        =   6
         Top             =   885
         Width           =   1980
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "售票员(&R):"
         Height          =   180
         Left            =   6210
         TabIndex        =   0
         Top             =   420
         Width           =   900
      End
      Begin VB.Image Image1 
         Height          =   1275
         Index           =   0
         Left            =   0
         Top             =   30
         Width           =   2010
      End
   End
   Begin MSComctlLib.ListView lvTicket 
      Height          =   4815
      Left            =   60
      TabIndex        =   1
      Top             =   1530
      Width           =   6780
      _ExtentX        =   11959
      _ExtentY        =   8493
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imlBusIcon"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList imlBusIcon 
      Left            =   6960
      Top             =   2370
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainNullTicketMan.frx":001C
            Key             =   "seller"
         EndProperty
      EndProperty
   End
   Begin ActiveBar2LibraryCtl.ActiveBar2 abAction 
      Height          =   5355
      Left            =   8400
      TabIndex        =   4
      Top             =   1200
      Width           =   1500
      _LayoutVersion  =   1
      _ExtentX        =   2646
      _ExtentY        =   9446
      _DataPath       =   ""
      Bands           =   "frmMainNullTicketMan.frx":0641
   End
   Begin VB.Menu pmnu_BusMan 
      Caption         =   "计划车次管理"
      Visible         =   0   'False
      Begin VB.Menu pmnu_BusPlanMan_Info 
         Caption         =   "车次属性"
         Enabled         =   0   'False
      End
      Begin VB.Menu pmnu_BusPlanMan_Allot 
         Caption         =   "车次配载"
      End
      Begin VB.Menu pmnu_BusPlanMan_Price 
         Caption         =   "车次票价信息"
         Enabled         =   0   'False
      End
      Begin VB.Menu pmnu_BusPlanMan_Envir 
         Caption         =   "环境预览"
         Enabled         =   0   'False
      End
      Begin VB.Menu pmnu_Break1 
         Caption         =   "-"
      End
      Begin VB.Menu pmnu_BusPlanMan_Stop 
         Caption         =   "车次停班"
         Enabled         =   0   'False
      End
      Begin VB.Menu pmnu_BusPlanMan_Resume 
         Caption         =   "车次复班"
         Enabled         =   0   'False
      End
      Begin VB.Menu pmnu_Break2 
         Caption         =   "-"
      End
      Begin VB.Menu pmnu_BusPlanMan_Add 
         Caption         =   "新增车次"
      End
      Begin VB.Menu pmnu_BusPlanMan_Copy 
         Caption         =   "复制车次"
         Enabled         =   0   'False
      End
      Begin VB.Menu pmnu_BusPlanMan_Del 
         Caption         =   "删除车次"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "frmManNullTicketMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'空白票管理 范鹏东 2013-07-05
Option Explicit
Const cnSeller = 0
Const cnGetnullTicketDate = 1
Const cnFirstnullTicketNo = 2
Const cnLastnullTicketNo = 3
Const cnMemo = 4

Private Sub abAction_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    Select Case Tool.name
        Case "act_GetNullTicket_Add"
            AddNullTicketTicketMan
        Case "act_DeleteNullTicketMan"
            DeleteNullTicketReocord
        Case "act_UpdateNullTicketMan"
            EditNullTicketMan
    End Select
End Sub

Private Sub cmdFind_Click()
    QueryGetNullTicketInfo dtpGetNullTicketStartDate.Value, dtpGetNullTicketEndDate.Value
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
           Case vbKeyReturn
           SendKeys "{TAB}"
    End Select
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Const cnMargin = 50
    ptShowInfo.Left = 0
    ptShowInfo.Top = 0
    ptShowInfo.Width = Me.ScaleWidth
    lvTicket.Left = cnMargin
    lvTicket.Top = ptShowInfo.Height + cnMargin
    lvTicket.Width = Me.ScaleWidth - IIf(abAction.Visible, abAction.Width, 0) - 2 * cnMargin
    lvTicket.Height = Me.ScaleHeight - ptShowInfo.Height - 2 * cnMargin
    '当操作条关闭时间处理
    If abAction.Visible Then
        abAction.Move lvTicket.Width + cnMargin, lvTicket.Top
        abAction.Height = lvTicket.Height
    End If
End Sub

Private Sub Form_Load()
    '初始化样式
    AlignHeadWidth Me.name, lvTicket
    AddLvTicketHeader
    Dim dyNow As Date
    dyNow = m_oParam.NowDate
    dtpGetNullTicketStartDate.Value = Format(DateAdd("m", 0, dyNow), "yyyy-mm-01")
    dtpGetNullTicketEndDate.Value = DateAdd("d", -1, DateAdd("m", 1, Format(dyNow, "yyyy-mm-01")))
    QueryGetNullTicketInfo
End Sub

Private Sub lvTicket_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    SortListView lvTicket, ColumnHeader.Index
End Sub

Public Sub AddList(szSellerID As String, dGetNullTicketDate As Date, szStartNo As String, szEndNo As String)
    On Error GoTo ErrHandle
    Dim i As Long, j As Integer
    Dim oTicketMan As New TicketMan
    Dim rsTemp As New Recordset
    oTicketMan.Init m_oActiveUser
    Set rsTemp = oTicketMan.GetNullTicketInfo(dGetNullTicketDate, dGetNullTicketDate, szSellerID, szStartNo, szEndNo)
    FillNullTicketInfo rsTemp
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Public Sub UpdateList(szSellerID As String, dGetNullTicketDate As Date, szStartNo As String, szEndNo As String)
    On Error GoTo ErrHandle
    Dim i As Long, j As Integer
    Dim oTicketMan As New TicketMan
    Dim rsTemp As New Recordset
    oTicketMan.Init m_oActiveUser
    Set rsTemp = oTicketMan.GetNullTicketInfo(dGetNullTicketDate, dGetNullTicketDate, szSellerID, szStartNo, szEndNo)
    FillNullTicketInfo rsTemp, True
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Public Sub QueryGetNullTicketInfo(Optional dtpGetNullTicketStat As Date = cszEmptyDateStr, Optional dtpGetNullTicketEnd As Date = cszForeverDateStr)
On Error GoTo ErrHandle
    Dim i As Integer, nCount As Integer
    Dim oListItem As ListItem
    Dim oTicketMan As New TicketMan
    Dim rsTmp As New Recordset
    lvTicket.ListItems.Clear
    oTicketMan.Init m_oActiveUser
    Set rsTmp = oTicketMan.GetNullTicketInfo(dtpGetNullTicketStat, dtpGetNullTicketEnd, txtSeller.Text)
    FillNullTicketInfo rsTmp
    Exit Sub
ErrHandle:
    SetNormal
    ShowErrorMsg
End Sub

Private Sub FillNullTicketInfo(prsInfo As Recordset, Optional isUpdate As Boolean = False)
'    填充信息
    Dim liTemp As ListItem
    Dim rsTemp As Recordset
    Dim i As Integer, nCount As Integer

    nCount = prsInfo.RecordCount
    If nCount = 0 Then Exit Sub

        For i = 1 To nCount
                If isUpdate = False Then
                    Set liTemp = lvTicket.ListItems.Add(, , MakeDisplayString(FormatDbValue(prsInfo!user_id), FormatDbValue(prsInfo!user_name)), , "seller")
                Else
                    Set liTemp = lvTicket.SelectedItem
                End If
                liTemp.SubItems(cnGetnullTicketDate) = Format(FormatDbValue(prsInfo!getnullticket_date), "YYYY-MM-DD")
                liTemp.SubItems(cnFirstnullTicketNo) = FormatDbValue(prsInfo!firstnullticket_no)
                liTemp.SubItems(cnLastnullTicketNo) = FormatDbValue(prsInfo!lastnullticket_no)
                liTemp.SubItems(cnMemo) = FormatDbValue(prsInfo!Memo)
            prsInfo.MoveNext
        Next i

    If nCount > 1 Then
        lvTicket.ListItems(1).Selected = True
        lvTicket.ListItems(1).EnsureVisible
    Else
        liTemp.Selected = True
        liTemp.EnsureVisible
    End If
   
End Sub

'填充列头
Private Sub AddLvTicketHeader()
    lvTicket.ColumnHeaders.Add , , "售票员", 2000
    lvTicket.ColumnHeaders.Add , , "登记空白票时间", 1800
    lvTicket.ColumnHeaders.Add , , "起始空白票号", 1800
    lvTicket.ColumnHeaders.Add , , "结束空白票号", 1800
    lvTicket.ColumnHeaders.Add , , "备注", 8000
End Sub

'删除空白票记录
Private Sub DeleteNullTicketReocord()
On Error GoTo ErrorHandle
    Dim nResult As VbMsgBoxResult
    Dim oTicketMan As New TicketMan
    Dim bIsDelete As Boolean
    oTicketMan.Init m_oActiveUser
    nResult = MsgBox("是否要删除" & Trim(lvTicket.SelectedItem.Text & "[" & Trim(lvTicket.SelectedItem.ListSubItems(1).Text) & "]" & "的空白票记录！"), vbQuestion + vbYesNo + vbDefaultButton2, "空白票管理")
    If nResult = vbYes Then
        bIsDelete = oTicketMan.DeleteNullTicketMan(ResolveDisplay(Trim(lvTicket.SelectedItem.Text)), Trim(lvTicket.SelectedItem.ListSubItems(1).Text), Trim(lvTicket.SelectedItem.ListSubItems(cnFirstnullTicketNo).Text), Trim(lvTicket.SelectedItem.ListSubItems(cnLastnullTicketNo).Text))
        If bIsDelete = True Then
            lvTicket.ListItems.Remove lvTicket.SelectedItem.Index
        End If
    End If
Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub lvTicket_DblClick()
    EditNullTicketMan
End Sub

'修改空白票
Private Sub EditNullTicketMan()
    If lvTicket.SelectedItem Is Nothing Then Exit Sub
    frmGetNullTicket.m_bIsParent = True
    frmGetNullTicket.m_User = Trim(lvTicket.SelectedItem.Text)
    frmGetNullTicket.m_GetNullTicketDate = CDate(lvTicket.SelectedItem.ListSubItems(1).Text)
    frmGetNullTicket.m_NullTicketStartNo = Trim(lvTicket.SelectedItem.ListSubItems(cnFirstnullTicketNo).Text)
    frmGetNullTicket.m_NullTicketEndNo = Trim(lvTicket.SelectedItem.ListSubItems(cnLastnullTicketNo).Text)
    frmGetNullTicket.Status = EFS_Modify
    frmGetNullTicket.Show vbModal
End Sub

'空白票登记
Private Sub AddNullTicketTicketMan()
    frmGetNullTicket.m_bIsParent = True
    frmGetNullTicket.Status = EFS_AddNew
    frmGetNullTicket.Show vbModal
End Sub
