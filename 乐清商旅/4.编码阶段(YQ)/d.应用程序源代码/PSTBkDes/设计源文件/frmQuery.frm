VERSION 5.00
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmQuery 
   BackColor       =   &H00E0E0E0&
   Caption         =   "预定信息"
   ClientHeight    =   5550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7065
   Icon            =   "frmQuery.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5550
   ScaleWidth      =   7065
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   390
      Top             =   4350
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQuery.frx":08CA
            Key             =   "Book"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox ptToolBar 
      Align           =   1  'Align Top
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   7065
      TabIndex        =   0
      Top             =   0
      Width           =   7065
      Begin VB.TextBox txtStation 
         Height          =   285
         Left            =   6990
         TabIndex        =   8
         Top             =   60
         Width           =   1365
      End
      Begin VB.TextBox txtBusID 
         Height          =   285
         Left            =   3840
         TabIndex        =   7
         Top             =   75
         Width           =   1065
      End
      Begin VB.CheckBox chkAll 
         BackColor       =   &H00E0E0E0&
         Caption         =   "全部(&A)"
         Height          =   285
         Left            =   1890
         TabIndex        =   5
         Top             =   75
         Value           =   1  'Checked
         Width           =   1035
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Height          =   45
         Left            =   30
         TabIndex        =   3
         Top             =   450
         Width           =   6675
      End
      Begin RTComctl3.CoolButton cmdRefresh 
         Height          =   315
         Left            =   8550
         TabIndex        =   2
         Top             =   60
         Width           =   1155
         _ExtentX        =   0
         _ExtentY        =   0
         BTYPE           =   3
         TX              =   "刷新(&R)"
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
         MICON           =   "frmQuery.frx":171E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComCtl2.DTPicker dtpBusDate 
         Height          =   315
         Left            =   60
         TabIndex        =   1
         Top             =   60
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   393216
         Format          =   62783488
         CurrentDate     =   36693
      End
      Begin VB.Label lblStation 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "站点代码/名称/简拼:"
         Height          =   180
         Left            =   5250
         TabIndex        =   9
         Top             =   105
         Width           =   1710
      End
      Begin VB.Label lblBusID 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "车次代码:"
         Height          =   180
         Left            =   3030
         TabIndex        =   6
         Top             =   127
         Width           =   810
      End
   End
   Begin MSComctlLib.ListView lvSeatInfo 
      Height          =   4155
      Left            =   720
      TabIndex        =   4
      Top             =   900
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   7329
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   14
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "bookeventid"
         Text            =   "流水号"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "busid"
         Text            =   "车次代码"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "startuptime"
         Text            =   "发车时间"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "seatno"
         Text            =   "座位号"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "booknumber"
         Text            =   "预定号"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "status"
         Text            =   "状态"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Key             =   "stationid"
         Text            =   "到站"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Key             =   "bookman"
         Text            =   "姓名"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Key             =   "userid"
         Text            =   "操作员"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Key             =   "telephone"
         Text            =   "电话"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Key             =   "email"
         Text            =   "地址"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Key             =   "canceluserid"
         Text            =   "取消预定操作员"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "操作时间"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Key             =   "annotation"
         Text            =   "备注"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdRefresh_Click()
    FillBookInfo
End Sub

Private Sub Form_Activate()
    MDIBook.EnableExportAndPrint True
End Sub

Private Sub Form_Deactivate()
    MDIBook.EnableExportAndPrint False
End Sub

Private Sub Form_Load()
    InitListView lvSeatInfo
    dtpBusDate.Value = m_oParam.NowDate
    FillBookInfo
End Sub

Private Sub Form_Resize()
    Dim lTemp As Long
    lTemp = Me.ScaleHeight - ptToolBar.Height
    lTemp = IIf(lTemp > 0, lTemp, 0)
    lvSeatInfo.Move 0, ptToolBar.Height, Me.ScaleWidth, lTemp
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveListView lvSeatInfo
    
    MDIBook.EnableExportAndPrint False
    
End Sub

Private Sub lvSeatInfo_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    SortListView lvSeatInfo, ColumnHeader.Index
End Sub

Public Sub lvSeatInfo_DblClick()
    On Error Resume Next
    With lvSeatInfo.SelectedItem
        frmBookAttr.lblAdd.Caption = "地址:" & .ListSubItems("email").Text
        frmBookAttr.lblBookID.Caption = "预定号:" & .ListSubItems("booknumber").Text
        frmBookAttr.lblBus.Caption = "车次代码:" & .ListSubItems("busid").Text
        If Trim(.ListSubItems("canceluserid").Text) = "" Then
            frmBookAttr.lblCancelOperation.Caption = ""
        Else
            frmBookAttr.lblCancelOperation.Caption = "取消操作员:" & .ListSubItems("canceluserid").Text
        End If
        frmBookAttr.lblDestStation.Caption = "终点站:" & .ListSubItems("stationid").Text
        frmBookAttr.lblID.Caption = "流水号:" & .Text
        frmBookAttr.lblMemo.Caption = "备注:" & .ListSubItems("annotation").Text
        frmBookAttr.lblName.Caption = "姓名:" & .ListSubItems("bookman").Text
        frmBookAttr.lblOperation.Caption = "操作员:" & .ListSubItems("userid").Text
        frmBookAttr.lblStartTime.Caption = "发车时间:" & .ListSubItems("startuptime").Text
        frmBookAttr.lblSeatNumber.Caption = "座位号:" & .ListSubItems("seatno").Text
        frmBookAttr.lblStauts.Caption = "状态:" & .ListSubItems("status").Text
        frmBookAttr.lblTelephone.Caption = "电话:" & .ListSubItems("telephone").Text
        frmBookAttr.lblOperateTime.Caption = "操作时间:" & .ListSubItems("operate_time").Text
    End With
    frmBookAttr.Show vbModal
End Sub

Private Sub lvSeatInfo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbKeyRButton Then
        PopupMenu MDIBook.mnu_Function2, , , , MDIBook.mnu_Refresh
    End If
End Sub

Private Sub mnu_Refresh_Click()
    FillBookInfo
End Sub

Private Sub ptToolBar_Resize()
    Frame1.Width = ptToolBar.ScaleWidth
End Sub

Public Sub FillBookInfo()
    Dim i As Long, lRecordCount As Long
    Dim rsTemp As Recordset
    Dim aszTemp() As String
    Dim liTemp As ListItem
    
    On Error GoTo Error_Handle
    lvSeatInfo.ListItems.Clear
    
    If Trim(txtBusID.Text) <> "" Then
        ReDim aszTemp(1 To 1)
        aszTemp(1) = Trim(txtBusID.Text)
    End If
            
    Set rsTemp = m_oBook.GetBookedSeat(dtpBusDate.Value, aszTemp, Trim(txtStation.Text))
    lRecordCount = rsTemp.RecordCount
    SetTaskLong lRecordCount
    For i = 1 To lRecordCount
        If chkAll.Value = vbChecked Or rsTemp!Status = ST_BOOKED Then
            Set liTemp = lvSeatInfo.ListItems.Add(, "A" & rsTemp!book_event_id, rsTemp!book_event_id, , "Book")
            liTemp.ListSubItems.Add , "busid", Trim(rsTemp!bus_id)
            liTemp.ListSubItems.Add , "startuptime", ToDBDateTime(rsTemp!bus_start_time)
            liTemp.ListSubItems.Add , "seatno", Trim(rsTemp!seat_no)
            liTemp.ListSubItems.Add , "booknumber", Trim(rsTemp!book_number)
            If rsTemp!Status <> ST_BOOKED Then
                 liTemp.ListSubItems.Add(, "status", GetBookStatusStr(rsTemp!Status)).ForeColor = vbRed
            Else
                 liTemp.ListSubItems.Add(, "status", GetBookStatusStr(rsTemp!Status)).ForeColor = vbBlack
            End If
            liTemp.ListSubItems.Add , "stationid", MakeDisplayString(Trim(rsTemp!station_id), Trim(rsTemp!station_name))
            liTemp.ListSubItems.Add , "bookman", Trim(rsTemp!book_man)
            liTemp.ListSubItems.Add , "userid", MakeDisplayString(Trim(rsTemp!user_id), Trim(rsTemp!user_name))
            liTemp.ListSubItems.Add , "telephone", Trim(rsTemp!telephone)
            liTemp.ListSubItems.Add , "email", Trim(rsTemp!email)
            
            If IsNull(rsTemp.Fields("cancel_user_name").Value) Then
                liTemp.ListSubItems.Add , "canceluserid", ""
            Else
                liTemp.ListSubItems.Add , "canceluserid", MakeDisplayString(Trim(rsTemp!cancel_user_id), Trim(rsTemp!cancel_user_name))
            End If
            liTemp.ListSubItems.Add , "operate_time", ToDBDateTime(rsTemp!operate_time)
            liTemp.ListSubItems.Add , "annotation", Trim(rsTemp!Annotation)
        End If
        SetTaskValue i
        rsTemp.MoveNext
    Next
    Exit Sub
Error_Handle:
    ShowErrorMsg
End Sub

Public Sub DeleteBookInfo()
    Dim i As Integer, j As Integer
    Dim aszTemp() As Long
    Dim lResult As Long
    If MsgBox("你真的要删除所选中的预定吗？", vbYesNo Or vbDefaultButton2, "询问") <> vbYes Then Exit Sub
    
    On Error GoTo Error_Handle
    If Not lvSeatInfo.SelectedItem Is Nothing Then
        ReDim aszTemp(1 To lvSeatInfo.ListItems.Count)
        j = 0
        For i = 1 To lvSeatInfo.ListItems.Count
            If lvSeatInfo.ListItems.Item(i).Selected Then
                j = j + 1
                aszTemp(j) = lvSeatInfo.ListItems.Item(i).Text
            End If
        Next
        
        ReDim Preserve aszTemp(1 To j)
        lResult = m_oBook.DeleteBookRec(aszTemp)
        If lResult = j Then
            SetTaskLong j, "正在删除..."
            For i = 1 To j
                SetTaskValue i
                lvSeatInfo.ListItems.Remove "A" & aszTemp(i)
            Next
        Else
            FillBookInfo
        End If
    End If
    Synchro
    Exit Sub
Error_Handle:
    SetMousePointer False
    ShowErrorMsg

End Sub

Public Sub UnBook()
    Dim i As Integer, nCount As Integer
    Dim aszTemp(1 To 1) As String
    Dim szBusID As String
    If MsgBox("你真的要取消所选中的预定吗？", vbYesNo Or vbDefaultButton2, "询问") <> vbYes Then Exit Sub
    
    On Error GoTo Error_Handle
    
    If Not lvSeatInfo.SelectedItem Is Nothing Then
        SetTaskLong lvSeatInfo.ListItems.Count, "正在取消预定..."
        For i = 1 To lvSeatInfo.ListItems.Count
            If lvSeatInfo.ListItems.Item(i).Selected Then
                aszTemp(1) = lvSeatInfo.ListItems.Item(i).ListSubItems("seatno").Text
                szBusID = lvSeatInfo.ListItems.Item(i).ListSubItems("busid").Text
                m_oBook.UnBook szBusID, dtpBusDate.Value, aszTemp
                lvSeatInfo.ListItems.Item(i).ListSubItems("status").Text = GetBookStatusStr(ST_BOOKCANCELED)
                lvSeatInfo.ListItems.Item(i).ListSubItems("status").ForeColor = vbRed
                lvSeatInfo.ListItems.Item(i).ListSubItems("canceluserid").Text = MakeDisplayString(m_oActiveUser.UserID, m_oActiveUser.UserName)
                
GoOnDoIt:
            End If
            SetTaskValue i
        Next
    End If
    Synchro
    Exit Sub
Error_Handle:
    SetMousePointer False
    ShowErrorMsg
    Resume GoOnDoIt
End Sub

Private Sub Synchro()
    Dim frmTemp As Form
    For Each frmTemp In Forms
        If TypeName(frmTemp) = "frmQuery" Then
            If Not frmTemp Is Me Then
                frmTemp.FillBookInfo
            End If
        End If
    Next
End Sub

'处理导出
Public Sub ExportFile(pbOpen As Boolean)
'    Set MDIBook.CellExport1.ListViewSource = lvSeatInfo
'    MDIBook.CellExport1.SourceSelect = ListViewControl
'    MDIBook.CellExport1.InitDir = GetDocumentDir()
'    MDIBook.CellExport1.ExportFile pbOpen
''
'    SaveDocumentDir MDIBook.CellExport1.InitDir

End Sub

Private Sub txtBusID_GotFocus()
    With txtBusID
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
