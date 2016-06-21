VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Begin VB.Form frmFormulaMan 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "票价公式管理"
   ClientHeight    =   3300
   ClientLeft      =   2910
   ClientTop       =   2910
   ClientWidth     =   7455
   HelpContextID   =   10000420
   Icon            =   "frmTicketFormulaMan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin RTComctl3.CoolButton cmdSetDefault 
      Height          =   315
      Left            =   6015
      TabIndex        =   5
      Top             =   1470
      Width           =   1215
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "设为缺省(&S)"
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
      MICON           =   "frmTicketFormulaMan.frx":014A
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
      Height          =   315
      Left            =   6015
      TabIndex        =   3
      Top             =   615
      Width           =   1215
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
      MICON           =   "frmTicketFormulaMan.frx":0166
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdProperty 
      Default         =   -1  'True
      Height          =   315
      Left            =   6015
      TabIndex        =   2
      Top             =   195
      Width           =   1215
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "属性(&P)"
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
      MICON           =   "frmTicketFormulaMan.frx":0182
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdDelete 
      Height          =   315
      Left            =   6015
      TabIndex        =   4
      Top             =   1050
      Width           =   1215
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
      MICON           =   "frmTicketFormulaMan.frx":019E
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
      Left            =   6015
      TabIndex        =   7
      Top             =   2520
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
      MICON           =   "frmTicketFormulaMan.frx":01BA
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
      Left            =   6015
      TabIndex        =   6
      Top             =   2070
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
      MICON           =   "frmTicketFormulaMan.frx":01D6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ListView lvFormula 
      Height          =   2805
      Left            =   75
      TabIndex        =   1
      Top             =   390
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   4948
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "FormulaName"
         Text            =   "公式名称"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "IsDefault"
         Text            =   "缺省执行标记"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "Annotation"
         Text            =   "注释"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "当前系统中所有的票价公式(&F):"
      Height          =   180
      Left            =   105
      TabIndex        =   0
      Top             =   135
      Width           =   2520
   End
End
Attribute VB_Name = "frmFormulaMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************
'* Source File Name:frmFormulaMan.frm
'* Project Name:PSTBusMan.vbp
'* Engineer:陈峰
'* Date Generated:2002/09/03
'* Last Revision Date:2002/09/03
'* Brief Description:票价公式管理
'* Relational Document:
'**********************************************************

Option Explicit

Public m_szPriceTableID As String


Private Sub cmdAdd_Click()
    AddFormula
End Sub

Private Sub cmdDelete_Click()
    DeleteFormula
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    DisplayHelp Me
End Sub

Private Sub cmdProperty_Click()
    EditFormula
End Sub

Private Sub cmdSetDefault_Click()
    SetDefault
End Sub


Private Sub Form_Load()
    FillFormula
End Sub

Private Sub FillFormula()
    '填充公式
    On Error GoTo ErrorHandle
    Dim i As Integer, nCount As Integer
    Dim aszFormula() As String
    Dim liTemp As ListItem
    Dim oTicketPriceMan As New TicketPriceMan
    
    oTicketPriceMan.Init g_oActiveUser
    aszFormula = oTicketPriceMan.GetAllTicketPriceFormula()
    nCount = ArrayLength(aszFormula)
    lvFormula.ListItems.Clear
    For i = 1 To nCount
        Set liTemp = lvFormula.ListItems.Add(, , RTrim(aszFormula(i, 1)))
        If CInt(aszFormula(i, 2)) <> 0 Then
            liTemp.SubItems(1) = "缺省公式"
        Else
            liTemp.SubItems(1) = "非缺省公式"
        End If
        liTemp.SubItems(2) = Trim(aszFormula(i, 3))
    Next
    Set oTicketPriceMan = Nothing
    SetEnabled
    Exit Sub
ErrorHandle:
    Set oTicketPriceMan = Nothing
    ShowErrorMsg
End Sub

Public Sub AddList(pszID As String)
    '新增后的刷新
        On Error GoTo ErrorHandle
    Dim i As Integer, nCount As Integer
    Dim aszFormula() As String
    Dim liTemp As ListItem
    Dim oFormula As New TicketPriceFormula
    
    oFormula.Init g_oActiveUser
    oFormula.Identify pszID
    Set liTemp = lvFormula.ListItems.Add(, , oFormula.FormulaName)
    If CInt(oFormula.IsDefault) <> 0 Then
        liTemp.SubItems(1) = "缺省公式"
    Else
        liTemp.SubItems(1) = "非缺省公式"
    End If
    liTemp.SubItems(2) = oFormula.Annotation
    Set oFormula = Nothing
    SetEnabled
    Exit Sub
ErrorHandle:
    Set oFormula = Nothing
    ShowErrorMsg

End Sub

Public Sub UpdateList(pszID As String)
    '修改后的刷新
    On Error GoTo ErrorHandle
    Dim i As Integer, nCount As Integer
    Dim aszFormula() As String
    Dim liTemp As ListItem
    Dim oFormula As New TicketPriceFormula
    
    oFormula.Init g_oActiveUser
    oFormula.Identify pszID
    Set liTemp = lvFormula.SelectedItem
    If CInt(oFormula.IsDefault) <> 0 Then
        liTemp.SubItems(1) = "缺省公式"
    Else
        liTemp.SubItems(1) = "非缺省公式"
    End If
    liTemp.SubItems(2) = oFormula.Annotation
    Set oFormula = Nothing
    SetEnabled
    Exit Sub
ErrorHandle:
    Set oFormula = Nothing
    ShowErrorMsg

End Sub

Private Sub SetEnabled()
    Dim liTemp As ListItem
    Set liTemp = lvFormula.SelectedItem
    cmdProperty.Enabled = IIf(liTemp Is Nothing, False, True)
    
    If liTemp Is Nothing Or lvFormula.ListItems.Count <= 1 Then
        cmdDelete.Enabled = False
    Else
        If liTemp.ListSubItems(1) = "缺省公式" Then
            cmdDelete.Enabled = False
        Else
            cmdDelete.Enabled = True
        End If
    End If
    
    If liTemp Is Nothing Then
        cmdSetDefault.Enabled = False
    Else
        If liTemp.ListSubItems(1) = "缺省公式" Then
            cmdSetDefault.Enabled = False
        Else
            cmdSetDefault.Enabled = True
        End If
    End If
    
End Sub

Private Sub lvFormula_Click()
    SetEnabled
End Sub

Private Sub DeleteFormula()
    '删除公式
    Dim oFormual As New TicketPriceFormula
    On Error GoTo ErrorHandle
    If MsgBox("你真的要删除选中的票价公式吗？", vbYesNo Or vbQuestion Or vbDefaultButton2) = vbYes Then
        oFormual.Init g_oActiveUser
        oFormual.Identify lvFormula.SelectedItem.Text
        oFormual.Delete
        lvFormula.ListItems.Remove lvFormula.SelectedItem.Index
    End If
    Set oFormual = Nothing
    Exit Sub
ErrorHandle:
    ShowErrorMsg
    Set oFormual = Nothing
End Sub

Private Sub AddFormula()
    '新增公式
    frmArrangeFormula.m_bIsParent = True
    frmArrangeFormula.m_eStatus = EFS_AddNew
    frmArrangeFormula.Show vbModal
End Sub

Private Sub EditFormula()
    '编辑公式
    frmArrangeFormula.m_bIsParent = True
    frmArrangeFormula.m_szFormulaID = lvFormula.SelectedItem.Text
    frmArrangeFormula.m_eStatus = EFS_Modify
    frmArrangeFormula.Show vbModal
End Sub

Private Sub SetDefault()
    '设置缺省
    Dim oFormula As New TicketPriceFormula
    On Error GoTo ErrorHandle
    
    oFormula.Init g_oActiveUser
    oFormula.Identify lvFormula.SelectedItem.Text
    oFormula.SetAsDefault
    Set oFormula = Nothing
    UpdateList lvFormula.SelectedItem.Text
    Exit Sub
ErrorHandle:
    Set oFormula = Nothing
    ShowErrorMsg
End Sub


Private Sub lvFormula_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    SortListView lvFormula, ColumnHeader.Index
End Sub

Private Sub lvFormula_DblClick()
    EditFormula
End Sub
