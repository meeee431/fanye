VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Begin VB.Form frmBookquery 
   Caption         =   "计划预订信息"
   ClientHeight    =   5115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8535
   Icon            =   "frmBookquery.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   8535
   StartUpPosition =   1  '所有者中心
   Begin MSComctlLib.ListView lvBookInfo 
      Height          =   3255
      Left            =   270
      TabIndex        =   12
      Top             =   1755
      Width           =   6945
      _ExtentX        =   12250
      _ExtentY        =   5741
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "车次代码"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "线路代码"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "预订起始时间"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "预订结束时间"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "预订起始座位"
         Object.Width           =   2559
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "预订总座位"
         Object.Width           =   1942
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "备  注"
         Object.Width           =   2540
      EndProperty
   End
   Begin RTComctl3.CoolButton cmdExit 
      Height          =   465
      Left            =   7380
      TabIndex        =   8
      Top             =   1440
      Width           =   1050
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
      MICON           =   "frmBookquery.frx":0442
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
      Height          =   465
      Left            =   7365
      TabIndex        =   7
      Top             =   855
      Width           =   1050
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
      MICON           =   "frmBookquery.frx":045E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdQuery 
      Height          =   465
      Left            =   7380
      TabIndex        =   6
      Top             =   270
      Width           =   1050
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
      MICON           =   "frmBookquery.frx":047A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame1 
      Caption         =   "查询条件"
      Height          =   1230
      Left            =   270
      TabIndex        =   0
      Top             =   135
      Width           =   6900
      Begin VB.ComboBox cboFunction 
         Height          =   300
         ItemData        =   "frmBookquery.frx":0496
         Left            =   5085
         List            =   "frmBookquery.frx":04A0
         TabIndex        =   11
         Text            =   "具体时间查询"
         Top             =   315
         Width           =   1500
      End
      Begin MSComCtl2.DTPicker dtStartdate 
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Top             =   315
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   503
         _Version        =   393216
         Format          =   64749569
         CurrentDate     =   37240
      End
      Begin MSComCtl2.DTPicker dtEnddate 
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   765
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   64749569
         CurrentDate     =   37240
      End
      Begin FText.asFlatTextBox txtBus 
         Height          =   285
         Left            =   5085
         TabIndex        =   13
         Top             =   765
         Width           =   1500
         _ExtentX        =   2646
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
         Text            =   "(全部)"
         ButtonBackColor =   -2147483633
         ButtonVisible   =   -1  'True
      End
      Begin VB.Label Label5 
         Caption         =   "查询方式："
         Height          =   195
         Left            =   4095
         TabIndex        =   10
         Top             =   360
         Width           =   1050
      End
      Begin VB.Label Label3 
         Caption         =   "查询车次："
         Height          =   240
         Left            =   4095
         TabIndex        =   5
         Top             =   810
         Width           =   915
      End
      Begin VB.Label Label2 
         Caption         =   "结束时间："
         Height          =   240
         Left            =   180
         TabIndex        =   2
         Top             =   810
         Width           =   1005
      End
      Begin VB.Label Label1 
         Caption         =   "起始时间："
         Height          =   240
         Left            =   180
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Label Label4 
      Caption         =   "订划预信息："
      Height          =   195
      Left            =   270
      TabIndex        =   9
      Top             =   1485
      Width           =   1680
   End
End
Attribute VB_Name = "frmBookquery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_oBusProject As New BusProject
Private bNoBusBookflg As Boolean

Private Sub cboFunction_Click()
  If cboFunction.ListIndex = 1 Then
      dtEnddate.Enabled = True
  Else
     dtEnddate.Enabled = False
  End If
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdQuery_Click()
lvBookInfo.ListItems.Clear
FilllstBookInfo
End Sub







Private Sub dtEnddate_Change()
IsSave
End Sub

Private Sub dtStartdate_Change()
IsSave
End Sub

Private Sub Form_Load()
'    m_oBusProject.Init g_oActiveUser
'    m_oBusProject.Identify g_szPlanID
'    dtEnddate.Value = Date
'    dtStartdate.Value = Date
'    FilllstBookInfo
End Sub
Private Function FilllstBookInfo()
Dim i As Integer
Dim nCount As Integer
Dim szBookInfo() As String
Dim liTemp As ListItem
Dim szBusID As String


If txtBus.Text <> "" And txtBus.Text <> "(全部)" Then
   szBusID = txtBus.Text
End If

If cboFunction.ListIndex = 1 Then
    szBookInfo = m_oBusProject.ProjectBookQuery(dtStartdate.Value, dtEnddate.Value, szBusID, True)
Else
    szBookInfo = m_oBusProject.ProjectBookQuery(dtStartdate.Value, dtStartdate.Value, szBusID, True)

End If
nCount = ArrayLength(szBookInfo)

If nCount = 0 Then CmdDelete.Enabled = False: Exit Function
lvBookInfo.ListItems.Clear
For i = 0 To nCount - 1
        Set liTemp = lvBookInfo.ListItems.Add(, , Trim(szBookInfo(i, 0)))
        liTemp.Checked = False
       
        liTemp.SubItems(1) = szBookInfo(i, 1)
        liTemp.SubItems(2) = szBookInfo(i, 2)
        liTemp.SubItems(3) = szBookInfo(i, 3)
        liTemp.SubItems(4) = szBookInfo(i, 4)
        liTemp.SubItems(5) = szBookInfo(i, 5)
        liTemp.SubItems(6) = szBookInfo(i, 6)
Next

'    szBookInfo(i, 0) = rsTemp!bus_id
'    szBookInfo(i, 1) = rsTemp!route_name
'    szBookInfo(i, 2) = rsTemp!start_bus_date
'    szBookInfo(i, 3) = rsTemp!End_bus_date
'    szBookInfo(i, 4) = rsTemp!start_seat_no
'    szBookInfo(i, 5) = rsTemp!seat_book_count
'    szBookInfo(i, 6) = rsTemp!remark_info
End Function

Private Sub CmdDelete_Click()
Dim i As Integer
Dim nCount As Integer
Dim szBookInfo() As String
Dim dtStartdate As Date
Dim dtEnddate As Date
Dim szBusID(0 To 0) As String
Dim bflgIsDelete As Boolean
On Error GoTo ErrorHandle
If (MsgBox("是否删除你选择的计划预定信息?", vbInformation + vbYesNo, Me.Caption)) = vbYes Then

With lvBookInfo
  nCount = .ListItems.Count
  For i = 1 To nCount
       
       If .ListItems(i).Checked = True Then
            dtStartdate = CDate(.ListItems(i).SubItems(2))
            dtEnddate = CDate(.ListItems(i).SubItems(3))
            szBusID(0) = .ListItems(i)
            m_oBusProject.ProjectBookDelete dtStartdate, dtEnddate, szBusID(), True
            .ListItems.Remove i
            bflgIsDelete = True
            i = i - 1
            nCount = nCount - 1
       End If
Next

If bflgIsDelete = True Then
   MsgBox "删除完成", vbInformation + vbOKOnly, Me.Caption
End If

End With

End If


Exit Sub
ErrorHandle:
  If err.Number = 35600 Then Exit Sub
  ShowErrorMsg
End Sub


Private Sub Form_Unload(Cancel As Integer)
Set m_oBusProject = Nothing
End Sub

Private Sub lvBookInfo_ItemCheck(ByVal Item As MSComctlLib.ListItem)

If Item.Checked = True Then
   CmdDelete.Enabled = True
End If

End Sub

Private Sub txtBus_Click()
'  Dim szTemp() As String
'
'  szTemp = selectAllBus()
'  On Error Resume Next
'  txtBus.Text = szTemp(1, 1)
End Sub

Private Function IsSave()
   If cboFunction.ListIndex = 1 Then
      If DateDiff("d", dtStartdate.Value, dtEnddate.Value) <= 0 Then
         dtEnddate.Value = dtStartdate.Value
      End If
   Else
     dtEnddate.Value = dtStartdate.Value
   End If
End Function
