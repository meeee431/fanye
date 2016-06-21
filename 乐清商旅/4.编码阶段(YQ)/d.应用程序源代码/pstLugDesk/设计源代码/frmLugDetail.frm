VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Begin VB.Form frmLugDetail 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "行包单明细"
   ClientHeight    =   4110
   ClientLeft      =   6060
   ClientTop       =   3330
   ClientWidth     =   7005
   Icon            =   "frmLugDetail.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   7005
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox ptTitle 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   795
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   7185
      TabIndex        =   2
      Top             =   0
      Width           =   7185
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   60
         Left            =   0
         TabIndex        =   3
         Top             =   750
         Width           =   7215
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "行包列表(&L):"
         Height          =   180
         Left            =   120
         TabIndex        =   6
         Top             =   330
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "行包单代码:"
         Height          =   180
         Left            =   1920
         TabIndex        =   5
         Top             =   330
         Width           =   990
      End
      Begin VB.Label lblAcceptID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00001"
         Height          =   180
         Left            =   2925
         TabIndex        =   4
         Top             =   330
         Width           =   450
      End
      Begin VB.Image Image1 
         Height          =   855
         Left            =   1470
         Picture         =   "frmLugDetail.frx":038A
         Top             =   0
         Width           =   5925
      End
   End
   Begin MSComctlLib.ListView lvInfo 
      Height          =   2475
      Left            =   15
      TabIndex        =   0
      Top             =   855
      Width           =   6960
      _ExtentX        =   12277
      _ExtentY        =   4366
      SortKey         =   1
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "标签号"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "物品名称"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "类型"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "件数"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "实重"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "计重"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "体积"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "包装形式"
         Object.Width           =   2540
      EndProperty
   End
   Begin RTComctl3.CoolButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   5820
      TabIndex        =   1
      Top             =   3630
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
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
      MICON           =   "frmLugDetail.frx":1874
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdAddNew 
      Height          =   315
      Left            =   2130
      TabIndex        =   7
      Top             =   3630
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "新增(&R)"
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
      MICON           =   "frmLugDetail.frx":1890
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdEdit 
      Height          =   315
      Left            =   3360
      TabIndex        =   8
      Top             =   3630
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "修改(&R)"
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
      MICON           =   "frmLugDetail.frx":18AC
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
      Left            =   4560
      TabIndex        =   9
      Top             =   3630
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "删除(&R)"
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
      MICON           =   "frmLugDetail.frx":18C8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   " RTStation"
      Enabled         =   0   'False
      Height          =   780
      Left            =   -120
      TabIndex        =   10
      Top             =   3360
      Width           =   8745
   End
End
Attribute VB_Name = "frmLugDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public LuggageID As String
Private Sub OKButton_Click()

End Sub

Private Sub cmdAddNew_Click()
'需要更改
On Error GoTo ErrHandle
    frmLuggage.LuggageItemID = LuggageID
    frmLuggage.Status = EFS_AddNew
    frmLuggage.Show vbModal
    RefreshlvInfo
 Exit Sub
ErrHandle:
 ShowErrorMsg
End Sub

Private Sub cmdCancel_Click()
Dim i As Integer
Dim mBagNum As Integer
Dim mCalWeight As Double
Dim mActWeight As Double
On Error GoTo ErrHandle
 
 If lvInfo.ListItems.Count > 0 Then
   For i = 1 To lvInfo.ListItems.Count
    mBagNum = mBagNum + CInt(lvInfo.ListItems(i).SubItems(3)) '件数
    mCalWeight = mCalWeight + CDbl(lvInfo.ListItems(i).SubItems(5)) '计重
    mActWeight = mActWeight + CDbl(lvInfo.ListItems(i).SubItems(4)) '实重
    
   Next i
    frmAccept.txtBagNum.Text = CStr(mBagNum) '件数
    frmAccept.txtCalWeight.Text = CStr(mCalWeight)  '计重
    frmAccept.txtActWeight.Text = CStr(mActWeight)  '实重
    frmAccept.txtStartLabel.Text = lvInfo.ListItems(1).Text  '起始标签号
    frmAccept.cboLuggageName.Text = lvInfo.ListItems(1).SubItems(1) '行包名称
    frmAccept.txtEndLabel.Text = NumAdd(lvInfo.ListItems(1).Text, mBagNum - 1) '结束标签号
'    frmAccept.txtPicker.SetFocus
 End If
 
 Unload Me
Exit Sub
ErrHandle:
ShowErrorMsg
End Sub

Private Sub cmdDelete_Click()
Dim szLabelID As String
On Error GoTo ErrHandle
   szLabelID = Trim(lvInfo.SelectedItem.Text)
   moAcceptSheet.DeleteLugItem szLabelID
   lvInfo.ListItems.Remove (lvInfo.SelectedItem.Index)
   RefreshlvInfo
   Exit Sub
ErrHandle:
 ShowErrorMsg
End Sub

Private Sub cmdEdit_Click()
On Error GoTo ErrHandle
frmLuggage.Status = EFS_Modify
frmLuggage.LuggageItemID = LuggageID
frmLuggage.Show vbModal
  RefreshlvInfo
  Exit Sub
ErrHandle:
 ShowErrorMsg
End Sub

Private Sub Form_Load()
     AlignFormPos Me
    lblAcceptID.Caption = LuggageID
    moAcceptSheet.Init m_oAUser
    RefreshlvInfo
End Sub
Private Sub RefreshlvInfo()
    On Error GoTo ErrHandle
    Dim sLugItem() As TLuggageItemInfo
    Dim i As Integer
    Dim nLen As Integer
    Dim lvs As ListItem
    Dim j As Integer
    nLen = ArrayLength(moAcceptSheet.GetLugItemDetail)
    If nLen > 0 Then
        cmdEdit.Enabled = True
        cmdDelete.Enabled = True
        ReDim sLugItem(1 To nLen)
        sLugItem = moAcceptSheet.GetLugItemDetail
        For i = 1 To nLen
          If lvInfo.ListItems.Count > 0 Then
            For j = 1 To lvInfo.ListItems.Count
             If Trim(lvInfo.ListItems(j).Text) = sLugItem(i).LabelID Then GoTo Nexthere
            Next j
          End If
            Set lvs = lvInfo.ListItems.Add(, , sLugItem(i).LabelID)
            lvs.SubItems(1) = sLugItem(i).LuggageName
            lvs.SubItems(2) = sLugItem(i).LuggageTypeName
            lvs.SubItems(3) = sLugItem(i).Number
            lvs.SubItems(4) = sLugItem(i).ActWeight
            lvs.SubItems(5) = sLugItem(i).CalWeight
            lvs.SubItems(6) = sLugItem(i).luggage_bulk
            lvs.SubItems(7) = sLugItem(i).PackType
Nexthere:
        Next i
    Else
        cmdEdit.Enabled = False
        cmdDelete.Enabled = False
    End If
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveFormPos Me
    cmdCancel_Click
End Sub

Private Sub lvInfo_DblClick()
    frmLuggage.Status = EFS_Show
    frmLuggage.LuggageItemID = LuggageID
    frmLuggage.Show vbModal
End Sub
