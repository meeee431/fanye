VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.0#0"; "RTComctl3.ocx"
Begin VB.Form frmProjectBook 
   Caption         =   "计划预订"
   ClientHeight    =   4725
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6645
   Icon            =   "frmProjectBook.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   6645
   StartUpPosition =   1  '所有者中心
   Begin RTComctl3.CoolButton CmdExit 
      Caption         =   "关 闭(&X)"
      Height          =   420
      Left            =   5490
      TabIndex        =   11
      Top             =   720
      Width           =   1050
   End
   Begin RTComctl3.CoolButton CmdSave 
      Caption         =   "保 存(&S)"
      Height          =   420
      Left            =   5490
      TabIndex        =   10
      Top             =   270
      Width           =   1050
   End
   Begin VB.Frame Frame2 
      Caption         =   "供用户设定预定信息"
      Height          =   1950
      Left            =   135
      TabIndex        =   0
      Top             =   135
      Width           =   5235
      Begin VB.TextBox txtRemarkInfo 
         Height          =   510
         Left            =   1260
         MaxLength       =   255
         TabIndex        =   13
         Top             =   1305
         Width           =   3660
      End
      Begin MSComCtl2.UpDown UpdStartSeatNo 
         Height          =   285
         Left            =   2296
         TabIndex        =   12
         Top             =   360
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "txtStartSeatNo"
         BuddyDispid     =   196617
         OrigLeft        =   2295
         OrigTop         =   360
         OrigRight       =   2535
         OrigBottom      =   645
         Max             =   99
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.DTPicker dtEndDate 
         Height          =   285
         Left            =   3690
         TabIndex        =   8
         Top             =   900
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   503
         _Version        =   393216
         Format          =   61931521
         CurrentDate     =   37238
      End
      Begin MSComCtl2.UpDown UpdSeatCount 
         Height          =   285
         Left            =   4726
         TabIndex        =   6
         Top             =   315
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "txtSeatCount"
         BuddyDispid     =   196616
         OrigLeft        =   4725
         OrigTop         =   315
         OrigRight       =   4965
         OrigBottom      =   600
         Max             =   100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtSeatCount 
         Height          =   285
         Left            =   3690
         MaxLength       =   3
         TabIndex        =   5
         Text            =   "10"
         Top             =   315
         Width           =   1035
      End
      Begin VB.TextBox txtStartSeatNo 
         Height          =   285
         Left            =   1260
         MaxLength       =   2
         TabIndex        =   3
         Text            =   "1"
         Top             =   360
         Width           =   1035
      End
      Begin MSComCtl2.DTPicker dtStartDate 
         Height          =   285
         Left            =   1260
         TabIndex        =   1
         Top             =   900
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   503
         _Version        =   393216
         Format          =   61931521
         CurrentDate     =   37238
      End
      Begin VB.Label Label5 
         Caption         =   "备    注："
         Height          =   240
         Left            =   180
         TabIndex        =   14
         Top             =   1305
         Width           =   915
      End
      Begin VB.Label Label4 
         Caption         =   "结束时间："
         Height          =   240
         Left            =   2655
         TabIndex        =   9
         Top             =   945
         Width           =   915
      End
      Begin VB.Label Label3 
         Caption         =   "起始时间："
         Height          =   285
         Left            =   180
         TabIndex        =   7
         Top             =   900
         Width           =   1005
      End
      Begin VB.Label Label2 
         Caption         =   "总预订座数："
         Height          =   240
         Left            =   2610
         TabIndex        =   4
         Top             =   360
         Width           =   1185
      End
      Begin VB.Label Label1 
         Caption         =   "起始座号："
         Height          =   285
         Left            =   180
         TabIndex        =   2
         Top             =   405
         Width           =   915
      End
   End
   Begin MSComctlLib.ListView lstBusBookInfo 
      Height          =   2010
      Left            =   135
      TabIndex        =   16
      Top             =   2475
      Width           =   5250
      _ExtentX        =   9260
      _ExtentY        =   3545
      SortKey         =   1
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ilBus"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "车次代码"
         Object.Width           =   1942
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "运行线路"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Label Label6 
      Caption         =   "预订车次列表："
      Height          =   240
      Left            =   135
      TabIndex        =   15
      Top             =   2115
      Width           =   1410
   End
End
Attribute VB_Name = "frmProjectBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_oBusProject As New BusProject
Private bNoBusBookflg As Boolean




Private Sub cmdExit_Click()
Unload Me

End Sub

Private Sub CmdSave_Click()
Dim szBookInfo(0 To 2) As String
Dim m_szBusInfo() As String
Dim i As Integer
On Error GoTo ErrorHandle
Dim nCount As Integer
szBookInfo(0) = txtStartSeatNo.Text
szBookInfo(1) = txtSeatCount.Text
szBookInfo(2) = txtRemarkInfo.Text
nCount = lstBusBookInfo.ListItems.Count

ReDim m_szBusInfo(0 To nCount - 1)

'取得车次
For i = 0 To nCount - 1
    If lstBusBookInfo.ListItems(i + 1).Checked = True Then
        m_szBusInfo(i) = lstBusBookInfo.ListItems(i + 1)
    End If
Next

'计划预订开始
If MsgBox("是否要计划预订？", vbInformation + vbYesNo) = vbYes Then
   If m_oBusProject.ProjectBook(dtStartdate.Value, dtEnddate, szBookInfo, m_szBusInfo) = True Then
     MsgBox "计划预订完成", vbInformation + vbOKOnly, Me.Caption
   End If
End If

Exit Sub

ErrorHandle:
 ShowErrorMsg
End Sub

Private Sub dtEnddate_Change()
    IsSaveOk
End Sub



Private Sub dtStartdate_Change()
    IsSaveOk
End Sub

Private Sub Form_Load()
' m_oBusProject.Init g_oActiveUser
' m_oBusProject.Identify g_szPlanID
' dtEnddate.Value = Date
' dtStartdate.Value = Date
'
' FilllstBookInfo
End Sub

Private Sub IsSaveOk()
 If bNoBusBookflg = False Then
    If txtStartSeatNo.Text <> "" And txtSeatCount.Text <> "" And DateDiff("d", dtStartdate.Value, dtEnddate.Value) >= 0 Then
       If IsNumeric(txtStartSeatNo.Text) And IsNumeric(txtSeatCount.Text) Then
           cmdSave.Enabled = True
       End If
    Else
       cmdSave.Enabled = False
    End If
 Else
     cmdSave.Enabled = False
 End If
End Sub



Private Sub Form_Unload(Cancel As Integer)
Set m_oBusProject = Nothing
End Sub

'Private Sub lstBusBookInfo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'
'If Button = 2 Then
'
' PopupMenu Mnu_DeletBusId
'
'End If
'End Sub

'Private Sub pMnu_DeletBusId_Click()
'Dim i As Integer
'Dim j As Integer
'If Button = 2 Then
' For i = 1 To lstBusBookInfo.ListItems.Count - j
'   On Error Resume Next
'   If lstBusBookInfo.ListItems(i).Selected = True Then
'      lstBusBookInfo.ListItems.Remove (i)
'      j = j + 1
'   End If
' Next
'End If
'End Sub

Private Sub txtSeatCount_Change()
    IsSaveOk
End Sub

Private Sub txtStartSeatNo_Change()
    IsSaveOk
End Sub
Private Function FilllstBookInfo()
'Dim i As Integer
'Dim liTemp As ListItem
'Dim bIsAdd As Boolean
'With frmBus
'If .bIsShow = True Then
''bNoBusBookflg = True
'cmdSave.Enabled = False
'If .lvBus.ListItems.Count <> 0 Then
'
'   cmdSave.Enabled = True
'
'   For i = 1 To .lvBus.ListItems.Count
'      If .lvBus.ListItems(i).Selected = True Then
'        Set liTemp = lstBusBookInfo.ListItems.Add(, , .lvBus.ListItems(i))
'        liTemp.Checked = True
'        liTemp.subitems()= .lvBus.ListItems(i).ListSubItems(2)
'
'      End If
'   Next
'    'listtem.ListSubItems.Remove listtem.ListSubItems.Count - 1
'End If
'
'End If
'End With
End Function
