VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Begin VB.Form frmCopyPriceTable 
   BackColor       =   &H00E0E0E0&
   Caption         =   "票价表复制"
   ClientHeight    =   5040
   ClientLeft      =   4485
   ClientTop       =   2385
   ClientWidth     =   6585
   Icon            =   "FrmPriceTableCopy.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   6585
   Begin VB.CheckBox ChkSelectAll 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "选择所有车次(&A)"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1920
      TabIndex        =   6
      Top             =   975
      Width           =   1770
   End
   Begin RTComctl3.CoolButton cmdClose 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   5280
      TabIndex        =   8
      Top             =   570
      Width           =   1170
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "关闭(&X)"
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
      MICON           =   "FrmPriceTableCopy.frx":014A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdCopy 
      Height          =   315
      Left            =   5280
      TabIndex        =   7
      Top             =   165
      Width           =   1170
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "复制(&C)"
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
      MICON           =   "FrmPriceTableCopy.frx":0166
      PICN            =   "FrmPriceTableCopy.frx":0182
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ComboBox cboDesTable 
      Height          =   300
      Left            =   1425
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   525
      Width           =   3165
   End
   Begin VB.ComboBox cboSouTable 
      Height          =   300
      Left            =   1425
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   165
      Width           =   3165
   End
   Begin MSComctlLib.ListView LvBus 
      Height          =   3645
      Left            =   105
      TabIndex        =   5
      Top             =   1275
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   6429
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "车次"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "发车时间"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "车次类型"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "运行线路"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "终点站"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblBus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "计划所属车次(&B):"
      Height          =   180
      Left            =   120
      TabIndex        =   4
      Top             =   975
      Width           =   1440
   End
   Begin VB.Label LblDesTable 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "目的票价表(&D):"
      Height          =   180
      Left            =   120
      TabIndex        =   2
      Top             =   585
      Width           =   1260
   End
   Begin VB.Label LblSouTable 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "源票价表(&S):"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   225
      Width           =   1080
   End
End
Attribute VB_Name = "frmCopyPriceTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CboDesTable_Click()
    EnableCopy
End Sub

Private Sub CboSouTable_Click()
    EnableCopy
End Sub

Private Sub ChkSelectAll_Click()
    Dim i As Integer
    If ChkSelectAll.Value = vbChecked Then
        For i = 1 To lvBus.ListItems.Count
            lvBus.ListItems(i).Selected = True
'            LvBus.ListItems(i).Checked = True
        Next
    Else
        For i = 1 To lvBus.ListItems.Count
            lvBus.ListItems(i).Selected = False
'            LvBus.ListItems(i).Checked = False
        Next
        lvBus.ListItems(1).Selected = True
'        LvBus.ListItems(1).Checked = True
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub CmdCopy_Click()
On Error GoTo ErrorHandle
    Dim BusID() As String
    Dim nBus As Integer
    Dim i As Integer
    Dim j As Integer

    If MsgBox("如果您的目标票价表有相同车次的票价信息，将会被覆盖!" & vbCrLf & "是否要继续复制?", vbQuestion + vbYesNo, "注意") <> vbYes Then Exit Sub

    For i = 1 To lvBus.ListItems.Count
        If lvBus.ListItems(i).Selected = True Then nBus = nBus + 1
    Next
    If nBus = 0 Then Exit Sub

    SetBusy
    ReDim BusID(1 To nBus)
    For i = 1 To lvBus.ListItems.Count
        If lvBus.ListItems(i).Selected = True Then
           j = j + 1
           BusID(j) = lvBus.ListItems(i).Text
           If j = nBus Then Exit For
        End If
    Next

    Dim oTicketPriceMan As TicketPriceMan
    Set oTicketPriceMan = New TicketPriceMan
    oTicketPriceMan.Init g_oActiveUser

    oTicketPriceMan.CopyPriceTable ResolveDisplay(cboSouTable.Text), ResolveDisplay(cboDesTable.Text), BusID
    SetNormal

    MsgBox "票价表复制成功!", vbInformation
    Exit Sub

ErrorHandle:
    SetNormal
    ShowErrorMsg
End Sub

Private Sub Form_Load()
'    FillProject
    FillSouTable
    FillDesTable
    RefreshBus

End Sub


Private Sub FillSouTable()
    Dim szRoutePriceTable() As String
    Dim i As Integer, nCount As Integer
    Dim szPriceTable As String

On Error GoTo ErrorHandle
    szRoutePriceTable = GetAllPriceTable()
    nCount = ArrayLength(szRoutePriceTable)

    cboSouTable.Clear
    If nCount > 0 Then
        For i = 1 To nCount
'            If Trim(szRoutePriceTable(i, 6)) = Trim(ProjectID) Then
                szPriceTable = MakeDisplayString(szRoutePriceTable(i, 1), szRoutePriceTable(i, 2))
                cboSouTable.AddItem szPriceTable
'            End If
        Next
        cboSouTable.ListIndex = 0
    End If

    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub


Private Sub FillDesTable()
    Dim szRoutePriceTable() As String
    Dim i As Integer, nCount As Integer
    Dim szPriceTable As String

On Error GoTo ErrorHandle
    szRoutePriceTable = GetAllPriceTable()
    nCount = ArrayLength(szRoutePriceTable)

    cboDesTable.Clear
    If nCount > 0 Then
        For i = 1 To nCount
'            If Trim(szRoutePriceTable(i, 6)) = Trim(ProjectID) Then
                szPriceTable = MakeDisplayString(szRoutePriceTable(i, 1), szRoutePriceTable(i, 2))
                cboDesTable.AddItem szPriceTable
'            End If
        Next i
'        cboDesTable.ListIndex = 0
    End If

    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Function GetAllPriceTable() As String()
    Dim oTicketPriceMan As New TicketPriceMan
    On Error GoTo ErrorHandle
    oTicketPriceMan.Init g_oActiveUser
    GetAllPriceTable = oTicketPriceMan.GetAllRoutePriceTable()
    Exit Function
ErrorHandle:
    ShowErrorMsg
End Function

Private Sub EnableCopy()
    If cboDesTable.Text = "" Or cboSouTable.Text = "" Or cboDesTable.Text = cboSouTable.Text Or lvBus.ListItems.Count = 0 Then
        cmdCopy.Enabled = False
    Else
        cmdCopy.Enabled = True
    End If
End Sub

Private Sub RefreshBus()
    On Error GoTo ErrorHandle
    Dim oProject As BusProject
    Dim nDataCount As Integer, i As Integer
    Dim liTemp As ListItem
    Dim vaTemp As Variant
    Set oProject = New BusProject

    oProject.Init g_oActiveUser
    oProject.Identify
    vaTemp = oProject.GetAllBus()
    nDataCount = ArrayLength(vaTemp)
    lvBus.ListItems.Clear
    For i = 1 To nDataCount
        Set liTemp = lvBus.ListItems.Add(, GetEncodedKey(RTrim(vaTemp(i, 1))), RTrim(vaTemp(i, 1)))
        liTemp.ListSubItems.Add , , Format(vaTemp(i, 2), "HH:mm")
        liTemp.ListSubItems.Add , , Trim(vaTemp(i, 8))
        liTemp.ListSubItems.Add , , RTrim(RTrim(vaTemp(i, 4)))
        liTemp.ListSubItems.Add , , RTrim(RTrim(vaTemp(i, 12)))
    Next
    Set oProject = Nothing

    lvBus.ListItems(1).Selected = True
ErrorHandle:
End Sub

Private Sub lvbus_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    SortListView lvBus, ColumnHeader.Index
End Sub
'Private Sub CopyBusPrice()
'On Error GoTo ErrorHandle
'Dim BusID() As String
'Dim oTicketPriceMan As TicketPriceMan
'
'
'    Set oTicketPriceMan = New TicketPriceMan
'    oTicketPriceMan.Init g_oActiveUser
'
'    SetBusy
'    BusID = m_szBus
'    oTicketPriceMan.CopyPriceTable m_szSouPriceTableID, m_szDesPriceTableID, BusID
'    SetNormal
'    Exit Sub
'
'ErrorHandle:
'    MsgBox err.Description
'End Sub
