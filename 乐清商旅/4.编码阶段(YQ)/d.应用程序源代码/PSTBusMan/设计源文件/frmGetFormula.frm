VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Object = "{BBF95DAA-F9CB-4CA9-A673-E0E9E0193752}#1.0#0"; "STSellCtl.ocx"
Begin VB.Form frmGetFormula 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "按公式修改"
   ClientHeight    =   3990
   ClientLeft      =   2655
   ClientTop       =   2790
   ClientWidth     =   6480
   HelpContextID   =   10000640
   Icon            =   "frmGetFormula.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "影响的票种"
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   135
      TabIndex        =   27
      Top             =   2775
      Width           =   6195
      Begin VB.CheckBox chkTicketType6 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "优惠票3"
         ForeColor       =   &H80000008&
         Height          =   400
         Left            =   4860
         TabIndex        =   18
         Top             =   210
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CheckBox chkTicketType5 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "优惠票2"
         ForeColor       =   &H80000008&
         Height          =   400
         Left            =   3711
         TabIndex        =   17
         Top             =   195
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.CheckBox chkTicketType4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "优惠票1"
         ForeColor       =   &H80000008&
         Height          =   400
         Left            =   2384
         TabIndex        =   16
         Top             =   180
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.CheckBox chkTicketType2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "半票"
         ForeColor       =   &H80000008&
         Height          =   400
         Left            =   1207
         TabIndex        =   15
         Top             =   195
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CheckBox chkTicketType1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "全票"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   400
         Left            =   120
         TabIndex        =   14
         Top             =   195
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1005
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "修改值"
      ForeColor       =   &H80000008&
      Height          =   2685
      Left            =   135
      TabIndex        =   23
      Top             =   30
      Width           =   6195
      Begin VB.ComboBox cboStationID 
         Height          =   315
         Left            =   4020
         TabIndex        =   28
         Top             =   788
         Width           =   1920
      End
      Begin RTComctl3.CoolButton cmdDelete 
         Height          =   315
         Left            =   4920
         TabIndex        =   11
         Top             =   1545
         Width           =   1080
         _ExtentX        =   0
         _ExtentY        =   0
         BTYPE           =   3
         TX              =   "删除(&D)"
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
         MICON           =   "frmGetFormula.frx":014A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin RTComctl3.CoolButton cmdAdd 
         Default         =   -1  'True
         Height          =   315
         Left            =   4920
         TabIndex        =   10
         Top             =   1155
         Width           =   1080
         _ExtentX        =   0
         _ExtentY        =   0
         BTYPE           =   3
         TX              =   "新增(&A)"
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
         MICON           =   "frmGetFormula.frx":0166
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.ComboBox cboItemSou 
         Height          =   300
         Left            =   1930
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   465
         Width           =   1575
      End
      Begin VB.ComboBox cboItem 
         Height          =   300
         Left            =   150
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   465
         Width           =   1440
      End
      Begin STSellCtl.ucNumTextBox txtMul 
         Height          =   300
         Left            =   3840
         TabIndex        =   5
         Top             =   465
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "1"
         MaxLength       =   6
         Alignment       =   1
         AllowNegative   =   -1  'True
      End
      Begin STSellCtl.ucNumTextBox txtAdd 
         Height          =   300
         Left            =   5160
         TabIndex        =   7
         Top             =   465
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   4
         Alignment       =   1
         AllowNegative   =   -1  'True
      End
      Begin MSComctlLib.ListView lvBrowse 
         Height          =   1500
         Left            =   150
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1110
         Width           =   4590
         _ExtentX        =   8096
         _ExtentY        =   2646
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "结果票价项"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "相关票价项"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "系数(乘)"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "附加值(加)"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "途径站"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblBusStationID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "途经站(&S):"
         Height          =   195
         Left            =   3000
         TabIndex        =   29
         Top             =   855
         Width           =   780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "=("
         Height          =   180
         Left            =   1670
         TabIndex        =   26
         Top             =   525
         Width           =   180
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "×"
         Height          =   180
         Left            =   3585
         TabIndex        =   25
         Top             =   525
         Width           =   180
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ")+"
         Height          =   180
         Left            =   4740
         TabIndex        =   24
         Top             =   525
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "修改票价项(&I)"
         Height          =   180
         Left            =   150
         TabIndex        =   0
         Top             =   210
         Width           =   1170
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "修改参照项(&F)"
         Height          =   180
         Left            =   1875
         TabIndex        =   2
         Top             =   210
         Width           =   1170
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "比率(&R)"
         Height          =   180
         Left            =   3750
         TabIndex        =   4
         Top             =   210
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "常数(&N)"
         Height          =   180
         Left            =   5250
         TabIndex        =   6
         Top             =   210
         Width           =   630
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "已添加的修改公式(&L):"
         Height          =   180
         Left            =   165
         TabIndex        =   8
         Top             =   855
         Width           =   1800
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "修改范围"
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   135
      TabIndex        =   22
      Top             =   5040
      Width           =   6195
      Begin VB.OptionButton optAll 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "修改所有打开的票价项(&O)"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   150
         TabIndex        =   12
         Top             =   255
         Value           =   -1  'True
         Width           =   2460
      End
      Begin VB.OptionButton optSelect 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "修改选择的票价项(&S)"
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   3540
         TabIndex        =   13
         Top             =   240
         Width           =   2205
      End
   End
   Begin RTComctl3.CoolButton cmdOK 
      Height          =   315
      Left            =   2370
      TabIndex        =   19
      Top             =   3570
      Width           =   1215
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
      MICON           =   "frmGetFormula.frx":0182
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
      Height          =   315
      Left            =   3675
      TabIndex        =   20
      Top             =   3570
      Width           =   1215
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
      MICON           =   "frmGetFormula.frx":019E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdHelp 
      Height          =   315
      HelpContextID   =   1001801
      Left            =   4995
      TabIndex        =   21
      Top             =   3570
      Width           =   1215
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "帮助(&H)"
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
      MICON           =   "frmGetFormula.frx":01BA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "frmGetFormula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************
'* Source File Name:frmGetFormula.frm
'* Project Name:PSTBusMan.vbp
'* Engineer:陈峰
'* Date Generated:2002/09/03
'* Last Revision Date:2002/09/10
'* Brief Description:得到修改参数设置
'* Relational Document:
'**********************************************************

Option Explicit

'Private m_aszItem1() As String '修改的票价项代码
'Private m_aszItem2() As String '修改的票价项名称
Public m_bOk As Boolean
Private m_rsPriceItem As Recordset

Private m_aszParam() As String
Private m_abTicketType() As Boolean  '所有需修改的票种

'--嘉兴批量修改
Public m_szRouteID As String

Public Property Get ModifyAll() As Boolean
    ModifyAll = optAll.Value
End Property

Public Property Get GetParam() As String()
    GetParam = m_aszParam
End Property

Public Property Get GetSelectTicketType() As Boolean()
    GetSelectTicketType = m_abTicketType
End Property



Private Sub cboItem_Click()
    EnableAdd
End Sub

Private Sub cboItemSou_Click()
    EnableAdd
End Sub

Private Sub cmdAdd_Click()
    Dim liTemp As ListItem
    Dim szTemp As String
    
    szTemp = cboItem.Text & Trim(cboStationID.Text)
    On Error GoTo ItemHaveExist
    Set liTemp = lvBrowse.ListItems.Add(, szTemp, cboItem.Text)
    On Error GoTo 0
    liTemp.ListSubItems.Add , cboItemSou.Text, cboItemSou.Text
    liTemp.ListSubItems.Add , , txtMul.Text
    liTemp.ListSubItems.Add , , txtAdd.Text
    
    liTemp.ListSubItems.Add , , cboStationID.Text
GoOn:
    EnableExecute
    EnableAdd
    EnableDelete
    Exit Sub
ItemHaveExist:
    If MsgBox("此票价项的修改公式已经存在，你要修改其参数吗？", vbYesNo Or vbQuestion Or vbDefaultButton2) = vbYes Then
        Set liTemp = lvBrowse.ListItems(szTemp)
        liTemp.ListSubItems(1).Text = cboItemSou.Text
        liTemp.ListSubItems(1).Key = cboItemSou.Text
        
        liTemp.ListSubItems(2).Text = txtMul.Text
        liTemp.ListSubItems(3).Text = txtAdd.Text
        
        Resume GoOn
    End If
End Sub

Private Sub cmdDelete_Click()
    Dim szTemp As String
    szTemp = lvBrowse.SelectedItem.Key
    lvBrowse.ListItems.Remove szTemp
    EnableExecute
    EnableDelete
End Sub

Private Sub cmdExit_Click()
    m_bOk = False
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    DisplayHelp Me
End Sub

Private Sub cmdOk_Click()
    m_bOk = True
    m_aszParam = GetModifyParam
    m_abTicketType = GetModifyTicketType
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()

    '--途径站列出来
    Dim m_oRoute As New Route
    m_oRoute.Init g_oActiveUser
    m_oRoute.Identify m_szRouteID
    Dim atSectionInfo() As TRouteSectionInfoEx
    atSectionInfo = m_oRoute.GetSectionInfoEx
    
    Dim j As Integer
    cboStationID.AddItem ""
    For j = 1 To ArrayLength(atSectionInfo)
        cboStationID.AddItem MakeDisplayString(atSectionInfo(j).sgEndStationMileage, atSectionInfo(j).szEndStationName)
    Next j
    
    cboStationID.ListIndex = 0
    
'    --嘉兴修改


    Dim nCount As Integer, i As Integer
    Dim atTicketType() As TTicketType
    Dim oParam As New SystemParam
    oParam.Init g_oActiveUser
    
    '得到所有的可用票种
    atTicketType = oParam.GetAllTicketType()
    nCount = ArrayLength(atTicketType)
    '设置票种的checkBox可用性
    For i = 1 To nCount
        Select Case atTicketType(i).nTicketTypeID
        Case TP_FullPrice
            chkTicketType1.Caption = Trim(atTicketType(i).szTicketTypename)
            If atTicketType(i).nTicketTypeValid Then chkTicketType1.Visible = True
        Case TP_HalfPrice
            chkTicketType2.Caption = Trim(atTicketType(i).szTicketTypename)
            If atTicketType(i).nTicketTypeValid Then chkTicketType2.Visible = True
        Case TP_PreferentialTicket1
            chkTicketType4.Caption = Trim(atTicketType(i).szTicketTypename)
            If atTicketType(i).nTicketTypeValid Then chkTicketType4.Visible = True
        Case TP_PreferentialTicket2
            chkTicketType5.Caption = Trim(atTicketType(i).szTicketTypename)
            If atTicketType(i).nTicketTypeValid Then chkTicketType5.Visible = True
        Case TP_PreferentialTicket3
            chkTicketType6.Caption = Trim(atTicketType(i).szTicketTypename)
            If atTicketType(i).nTicketTypeValid Then chkTicketType6.Visible = True
        End Select
    Next i
    m_bOk = False
'    Dim oTicketPriceMan As New STPrice.TicketPriceMan
'    oTicketPriceMan.Init g_oActiveUser
'    oTicketPriceMan.GetAllTicketItem
    GetAllUsePriceItem
    
    nCount = m_rsPriceItem.RecordCount
    cboItem.Clear
    For i = 1 To nCount
        cboItem.AddItem MakeDisplayString(m_rsPriceItem!price_item, m_rsPriceItem!chinese_name)
        cboItemSou.AddItem MakeDisplayString(m_rsPriceItem!price_item, m_rsPriceItem!chinese_name)
        m_rsPriceItem.MoveNext
    Next
    cboItemSou.AddItem "A000[里程数]"
    EnableAdd
    EnableDelete
    EnableExecute
End Sub


Private Sub EnableExecute()
    '修改是否可用
    cmdOk.Enabled = IIf(lvBrowse.ListItems.Count > 0, True, False)
End Sub

Private Sub EnableAdd()
    '添加是否可用
    If cboItem.ListIndex >= 0 And cboItem.ListIndex >= 0 Then
        cmdAdd.Enabled = True
    Else
        cmdAdd.Enabled = False
    End If
End Sub

Private Sub lvBrowse_ItemClick(ByVal Item As MSComctlLib.ListItem)
    EnableDelete
End Sub

Private Sub EnableDelete()
    '删除是否可用
    cmdDelete.Enabled = IIf(lvBrowse.SelectedItem Is Nothing, False, True)
End Sub

Private Function GetModifyParam() As String()
    '得到参数
    Dim aTemp() As String
    Dim i As Integer
    
    ReDim aTemp(1 To lvBrowse.ListItems.Count, 1 To 6)
    
    For i = 1 To lvBrowse.ListItems.Count
        aTemp(i, 1) = ResolveDisplay(lvBrowse.ListItems(i).Text)
        aTemp(i, 2) = ResolveDisplay(lvBrowse.ListItems(i).ListSubItems(1).Text)
        aTemp(i, 3) = lvBrowse.ListItems(i).ListSubItems(2).Text
        aTemp(i, 4) = lvBrowse.ListItems(i).ListSubItems(3).Text
        aTemp(i, 5) = ResolveDisplayEx(lvBrowse.ListItems(i).ListSubItems(4).Text)
        aTemp(i, 6) = ResolveDisplay(lvBrowse.ListItems(i).ListSubItems(4).Text)
    Next
    GetModifyParam = aTemp
End Function

Private Function GetModifyTicketType() As Boolean()

    '得到需修改的票种
    Dim abTemp() As Boolean
    Dim i As Integer
    Dim t As Integer
    ReDim abTemp(1 To TP_TicketTypeCount)
    
    For i = 1 To TP_TicketTypeCount
        Select Case i
        Case TP_FullPrice
'            t = t + 1
'            ReDim Preserve abTemp(1 To t)
            abTemp(TP_FullPrice) = IIf(chkTicketType1.Value = vbChecked, True, False)
        Case TP_HalfPrice
'            t = t + 1
'            ReDim Preserve abTemp(1 To t)
            abTemp(TP_HalfPrice) = IIf(chkTicketType2.Value = vbChecked, True, False)
        Case TP_PreferentialTicket1
'            t = t + 1
'            ReDim Preserve abTemp(1 To t)
            abTemp(TP_PreferentialTicket1) = IIf(chkTicketType4.Value = vbChecked, True, False)
        Case TP_PreferentialTicket2
'            t = t + 1
'            ReDim Preserve abTemp(1 To t)
            abTemp(TP_PreferentialTicket2) = IIf(chkTicketType5.Value = vbChecked, True, False)
        Case TP_PreferentialTicket3
'            t = t + 1
'            ReDim Preserve abTemp(1 To t)
            abTemp(TP_PreferentialTicket3) = IIf(chkTicketType6.Value = vbChecked, True, False)
        End Select
        
    Next i
    GetModifyTicketType = abTemp
End Function

Private Sub txtMul_LostFocus()
    If txtMul.Text > 100 Then
       MsgBox "比率不能大于100！", vbOKOnly + vbCritical, Me.Caption
       txtMul.SetFocus
    End If
End Sub

Private Sub GetAllUsePriceItem()
    '得到所有的可用的票种
    Dim i As Integer
    Dim nCount As Integer
    Dim rsTemp As Recordset
    Dim oPriceMan As New TicketPriceMan
    
    oPriceMan.Init g_oActiveUser
    Set m_rsPriceItem = oPriceMan.GetAllTicketItemRS(TP_TicketTypeValid)
    nCount = m_rsPriceItem.RecordCount
    
End Sub
