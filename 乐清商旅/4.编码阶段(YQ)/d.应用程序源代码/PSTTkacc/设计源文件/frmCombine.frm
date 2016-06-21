VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BBF95DAA-F9CB-4CA9-A673-E0E9E0193752}#1.0#0"; "STSellCtl.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Begin VB.Form frmCombine 
   BackColor       =   &H00E0E0E0&
   Caption         =   "车次公司组合设置"
   ClientHeight    =   4800
   ClientLeft      =   1920
   ClientTop       =   2205
   ClientWidth     =   9270
   Icon            =   "frmCombine.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4800
   ScaleWidth      =   9270
   WindowState     =   2  'Maximized
   Begin STSellCtl.ucUpDownText txtCombineSerial 
      Height          =   285
      Left            =   1200
      TabIndex        =   7
      Top             =   120
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   503
      SelectOnEntry   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Max             =   100
      Min             =   -1
      Value           =   "-1"
   End
   Begin VB.TextBox txtCompanyName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5850
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtCompanyID 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3480
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin RTComctl3.CoolButton cmdQuery 
      Default         =   -1  'True
      Height          =   345
      Left            =   7110
      TabIndex        =   4
      Top             =   90
      Width           =   1185
      _ExtentX        =   0
      _ExtentY        =   0
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
      MICON           =   "frmCombine.frx":000C
      PICN            =   "frmCombine.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ListView lvCombine 
      Height          =   2385
      Left            =   90
      TabIndex        =   0
      Top             =   510
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   4207
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "组合序号"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "起始车次"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "结束车次"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "公司代码"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "公司名称"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "公司名称(&N):"
      Height          =   180
      Left            =   4740
      TabIndex        =   5
      Top             =   165
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "公司代码(&C):"
      Height          =   180
      Left            =   2340
      TabIndex        =   2
      Top             =   165
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "组合序号(&S):"
      Height          =   180
      Left            =   120
      TabIndex        =   1
      Top             =   165
      Width           =   1080
   End
End
Attribute VB_Name = "frmCombine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oTkAcc As New TicketCompanyDim
Const cnStartBusID = 1
Const cnEndBusID = 2
Const cnCompanyID = 3
Const cnCompanyName = 4



Private Sub cmdQuery_Click()
    FillCombine
End Sub

Private Sub Form_Load()
    
    oTkAcc.Init m_oActiveUser
    
End Sub

Private Sub Form_Resize()
    Dim lTemp As Long
    On Error Resume Next
    lTemp = Me.ScaleHeight - 500
    lTemp = IIf(lTemp > 0, lTemp, 0)
    lvCombine.Move 0, 500, Me.ScaleWidth - 50, lTemp
    
End Sub

Private Sub lvCombine_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Static m_nUpColumn As Integer
lvCombine.SortKey = ColumnHeader.Index - 1
If m_nUpColumn = ColumnHeader.Index - 1 Then
    lvCombine.SortOrder = lvwDescending
    m_nUpColumn = ColumnHeader.Index
Else
    lvCombine.SortOrder = lvwAscending
    m_nUpColumn = ColumnHeader.Index - 1
End If
lvCombine.Sorted = True
End Sub

Public Sub FillCombine()
'填充查询
    Dim atCombineInfo() As tBusCompanyCombineInfo
    Dim nCount As Integer
    Dim i As Integer
    Dim liTemp As ListItem
    On Error GoTo Here
    lvCombine.ListItems.Clear
    
    atCombineInfo = oTkAcc.QueryCombile(txtCombineSerial.Value, txtCompanyID.Text, txtCompanyName.Text)
    nCount = ArrayLength(atCombineInfo)
    For i = 1 To nCount
        Set liTemp = lvCombine.ListItems.Add(, , atCombineInfo(i).CombineSerial)
        
        
        liTemp.ListSubItems.Add cnStartBusID, , atCombineInfo(i).StartBusId
        liTemp.ListSubItems.Add cnEndBusID, , atCombineInfo(i).EndBusID
        liTemp.ListSubItems.Add cnCompanyID, , atCombineInfo(i).TransportCompanyID
        liTemp.ListSubItems.Add cnCompanyName, , atCombineInfo(i).TransportCompanyName
        
    Next i
    
    Exit Sub
    
Here:
    ShowErrorMsg
    
End Sub

Public Sub DeleteCombine()
'删除组合
    On Error GoTo Here
    If lvCombine.SelectedItem.Text <> "" Then
        oTkAcc.DeleteCombileBusCompany lvCombine.SelectedItem.Text, lvCombine.SelectedItem.ListSubItems(cnStartBusID), lvCombine.SelectedItem.ListSubItems(cnEndBusID)
        lvCombine.ListItems.Remove lvCombine.SelectedItem.Index
    End If
    Exit Sub
Here:
    ShowErrorMsg
End Sub

Public Sub AddCombine()
'新增组合
    With frmAddCombine
        .Status = EFormStatus.SNAddNew
        .Show vbModal
        
    End With
End Sub

Public Sub ModifyCombine()
'修改组合
'    if lvCombine.ListItems.Count >0 then
'    If lvCombine.SelectedItem.Text <> "" Then
'
'    End If
End Sub

Private Sub lvCombine_DblClick()
    ModifyCombine
End Sub

Private Sub lvCombine_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 93 Then PopupMenu MDIMain.pmnu_Combine
End Sub

Private Sub lvCombine_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu MDIMain.pmnu_Combine
    End If
End Sub
