VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "actbar2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmAllFinSheets 
   BackColor       =   &H00E0E0E0&
   Caption         =   "行包结算单"
   ClientHeight    =   8550
   ClientLeft      =   540
   ClientTop       =   2610
   ClientWidth     =   13215
   HelpContextID   =   7000390
   Icon            =   "frmAllFinSheets.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8550
   ScaleWidth      =   13215
   WindowState     =   2  'Maximized
   Begin VB.PictureBox ptMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7335
      Left            =   30
      ScaleHeight     =   7335
      ScaleWidth      =   10995
      TabIndex        =   3
      Top             =   1110
      Width           =   10995
      Begin VB.PictureBox ptFinSheet 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   5535
         Left            =   90
         ScaleHeight     =   5505
         ScaleWidth      =   10635
         TabIndex        =   4
         Top             =   180
         Width           =   10665
         Begin ActiveBar2LibraryCtl.ActiveBar2 abAction 
            Height          =   5280
            Left            =   9090
            TabIndex        =   5
            Top             =   60
            Width           =   1530
            _LayoutVersion  =   1
            _ExtentX        =   2699
            _ExtentY        =   9313
            _DataPath       =   ""
            Bands           =   "frmAllFinSheets.frx":000C
         End
         Begin MSComctlLib.ListView lvFinSheets 
            Height          =   4635
            Left            =   180
            TabIndex        =   6
            Top             =   90
            Width           =   8340
            _ExtentX        =   14711
            _ExtentY        =   8176
            View            =   3
            Arrange         =   1
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            SmallIcons      =   "imglv"
            ForeColor       =   -2147483640
            BackColor       =   16777215
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
            NumItems        =   0
         End
      End
      Begin MSComctlLib.ListView lvLuggageSheet 
         Height          =   1125
         Left            =   720
         TabIndex        =   7
         Top             =   6180
         Width           =   7440
         _ExtentX        =   13123
         _ExtentY        =   1984
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imglv"
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
         NumItems        =   0
      End
      Begin RTComctl3.Spliter Spliter1 
         Height          =   45
         Left            =   2010
         TabIndex        =   19
         Top             =   5910
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   79
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IsVertical      =   -1  'True
      End
   End
   Begin VB.PictureBox ptShowInfo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   990
      Left            =   0
      ScaleHeight     =   990
      ScaleWidth      =   12525
      TabIndex        =   0
      Top             =   30
      Width           =   12525
      Begin VB.ComboBox cboStatus 
         Height          =   300
         Left            =   8520
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   510
         Width           =   1335
      End
      Begin VB.ComboBox cboSellStation 
         Height          =   300
         ItemData        =   "frmAllFinSheets.frx":4050
         Left            =   2910
         List            =   "frmAllFinSheets.frx":4052
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   120
         Width           =   1605
      End
      Begin RTComctl3.CoolButton cmdFind 
         Default         =   -1  'True
         Height          =   345
         Left            =   10380
         TabIndex        =   1
         Top             =   240
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
         MICON           =   "frmAllFinSheets.frx":4054
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   315
         Left            =   5850
         TabIndex        =   8
         Top             =   120
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   62390275
         CurrentDate     =   37646
      End
      Begin FText.asFlatTextBox txtObject 
         Height          =   315
         Left            =   2910
         TabIndex        =   10
         Top             =   540
         Width           =   1575
         _ExtentX        =   2778
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
         Text            =   "全部"
         ButtonVisible   =   -1  'True
      End
      Begin MSComctlLib.ImageCombo imgcbo 
         Height          =   330
         Left            =   5850
         TabIndex        =   11
         Top             =   510
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Locked          =   -1  'True
         ImageList       =   "imglv"
      End
      Begin MSComCtl2.DTPicker dtpEndDate 
         Height          =   315
         Left            =   8520
         TabIndex        =   21
         Top             =   120
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   62390275
         CurrentDate     =   37646
      End
      Begin VB.ComboBox cboAcceptType 
         Height          =   300
         Left            =   8520
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   120
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "结束日期(&F):"
         Height          =   180
         Left            =   7380
         TabIndex        =   20
         Top             =   180
         Width           =   1080
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "车站(&S):"
         Height          =   180
         Left            =   1980
         TabIndex        =   16
         Top             =   180
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "状态(&E):"
         Height          =   180
         Left            =   7740
         TabIndex        =   14
         Top             =   570
         Width           =   720
      End
      Begin VB.Label lblType 
         BackStyle       =   0  'Transparent
         Caption         =   "对象(&O):"
         Height          =   180
         Left            =   1950
         TabIndex        =   13
         Top             =   570
         Width           =   810
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "查询类型(&P):"
         Height          =   180
         Left            =   4680
         TabIndex        =   12
         Top             =   570
         Width           =   1110
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "托运方式(&T):"
         Height          =   180
         Left            =   7380
         TabIndex        =   9
         Top             =   150
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "开始日期(&D):"
         Height          =   180
         Left            =   4680
         TabIndex        =   2
         Top             =   150
         Width           =   1080
      End
      Begin VB.Image Image1 
         Height          =   1275
         Left            =   0
         Picture         =   "frmAllFinSheets.frx":4070
         Top             =   30
         Width           =   2010
      End
   End
   Begin MSComctlLib.ImageList imglv 
      Left            =   12270
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAllFinSheets.frx":5543
            Key             =   "splitcompany"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAllFinSheets.frx":569D
            Key             =   "company"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAllFinSheets.frx":57F7
            Key             =   "vehicle"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAllFinSheets.frx":5951
            Key             =   "bus"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAllFinSheets.frx":5AAB
            Key             =   "busowner"
         EndProperty
      EndProperty
   End
   Begin VB.Menu pmnu_SettleLG 
      Caption         =   "行包结算单"
      Visible         =   0   'False
      Begin VB.Menu pmnu_Open 
         Caption         =   "打开(&O)"
      End
      Begin VB.Menu pmnu_Property 
         Caption         =   "属性(&P)"
      End
      Begin VB.Menu pmnu_Detail 
         Caption         =   "明细信息(&T)"
      End
      Begin VB.Menu pmnu_Break1 
         Caption         =   "-"
      End
      Begin VB.Menu pmnu_Delete 
         Caption         =   "删除(&D)"
      End
      Begin VB.Menu pmnu_Cancel 
         Caption         =   "作废(&C)"
      End
      Begin VB.Menu pmnu_Break2 
         Caption         =   "-"
      End
      Begin VB.Menu pmnu_NewWizard 
         Caption         =   "新增向导(&W)"
      End
      Begin VB.Menu pmnu_NewExtra 
         Caption         =   "非同行结算(&E)"
      End
   End
End
Attribute VB_Name = "frmAllFinSheets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mSheetID() As String '准备设置的结算单数组
Dim mSheetNum As Integer '准备设置的结算单总数
'Private m_oFinSheet As New FinanceSheet

Private Sub abAction_BandClose(ByVal Band As ActiveBar2LibraryCtl.Band)
    '设置控件的位置
    lvFinSheets.Left = 0 ' 50
    lvFinSheets.Top = 0 'Me.Height + 50
    lvFinSheets.Width = ptMain.Width - IIf(abAction.Visible, abAction.Width, 0) - 20
    lvFinSheets.Height = ptMain.Height - 20
    lvFinSheets.Refresh
    '当操作条关闭时间处理
    If abAction.Visible Then
        abAction.Move lvFinSheets.Width + 20, lvFinSheets.Top
        abAction.Height = lvFinSheets.Height
    End If
    
End Sub

'获得结算单数组
Private Sub GetSheetID()
 On Error GoTo ErrHandle
   Dim i As Integer
   Dim nlen As Integer
   Dim j As Integer
   mSheetNum = 0
   If lvFinSheets.ListItems.Count = 0 Then Exit Sub
   For i = 1 To lvFinSheets.ListItems.Count
       If lvFinSheets.ListItems.Item(i).Checked = True Then
          mSheetNum = mSheetNum + 1
       End If
   Next i
   If mSheetNum = 0 Then Exit Sub
   
   ReDim mSheetID(1 To mSheetNum)
   j = 1
   For i = 1 To lvFinSheets.ListItems.Count
       If lvFinSheets.ListItems.Item(i).Checked = True Then
          mSheetID(j) = Trim(lvFinSheets.ListItems(i).Text)
          j = j + 1
       End If
   Next i
   
 Exit Sub
ErrHandle:
 ShowErrorMsg
End Sub


Private Sub abAction_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
    Form_Resize
    '获得结算单数组
    GetSheetID
    Select Case Tool.name
    Case "Property"  '属性
        ShowFinSheet
        cmdFind_Click
    Case "FinSheetlst" '车辆结算明细
        ShowFinSheetlst
    Case "LuggageSheet" '行包签发单
        ShowSheetInfo
    Case "NewWizard"  '新增向导
        ShowWiz
        cmdFind_Click
    Case "SetPayed" '设置已结
        ShowSetPayed
        cmdFind_Click
    Case "Cancel" '作废
        ShowSheetCancel
        cmdFind_Click
    Case "Delete"  '删除
        ShowDeleteSheet
        cmdFind_Click
    End Select
    WriteProcessBar False
End Sub
 '打开行包结算单属性
Public Sub ShowFinSheet()
    
    Dim oCommDialog As New STShell.CommDialog
    On Error GoTo ErrorHandle
    
    If lvFinSheets.SelectedItem Is Nothing Then Exit Sub
    oCommDialog.Init m_oAUser
    oCommDialog.ShowLugFinSheet lvFinSheets.SelectedItem.Text, Trim(lvFinSheets.SelectedItem.SubItems(1))
    Set oCommDialog = Nothing
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub
'填充车辆结算信息lvLuggageSheet列首
Private Sub FilllvLuggageBusHead()
        With lvLuggageSheet.ColumnHeaders
          .clear
          .Add , , "结算单号"
          .Add , , "车辆牌号"
          .Add , , "车主姓名"
          .Add , , "上车站"
          .Add , , "参营公司"
          .Add , , "拆帐公司"
          .Add , , "线路名称"
          .Add , , "协议名称"
          .Add , , "结算款"
          .Add , , "应拆金额"
          .Add , , "结算月份"
     End With
End Sub
'车辆结算明细
Public Sub ShowFinSheetlst()
On Error GoTo ErrHandle
  Dim i As Integer
  Dim nlen As Integer
  Dim rsTemp As Recordset
  Dim lvItem As ListItem

     lvLuggageSheet.Visible = True
     
     Spliter1.Visible = True
     

     Form_Resize
     ptMain_Resize
     '填充lvLuggageSheet
     FilllvLuggageBusHead
     
     lvLuggageSheet.ListItems.clear
     If lvFinSheets.ListItems.Count = 0 Then Exit Sub
     m_oFinanceSheet.SheetID = Trim(lvFinSheets.SelectedItem.Text)
     Set rsTemp = m_oFinanceSheet.GetFinSheetDetailRS
     If rsTemp.RecordCount = 0 Then Exit Sub
      
     For i = 1 To rsTemp.RecordCount
         Set lvItem = lvLuggageSheet.ListItems.Add(, , FormatDbValue(rsTemp!fin_sheet_id))
             lvItem.SubItems(1) = FormatDbValue(rsTemp!license_tag_no)
             lvItem.SubItems(2) = FormatDbValue(rsTemp!owner_name)
             lvItem.SubItems(3) = FormatDbValue(rsTemp!sell_station_id)  '转换成名称
             lvItem.SubItems(4) = FormatDbValue(rsTemp!transport_company_short_name)
             lvItem.SubItems(5) = FormatDbValue(rsTemp!split_company_name)
             lvItem.SubItems(6) = FormatDbValue(rsTemp!route_name)
             lvItem.SubItems(7) = FormatDbValue(rsTemp!protocol_name)
             lvItem.SubItems(8) = FormatDbValue(rsTemp!total_price)
             lvItem.SubItems(9) = FormatDbValue(rsTemp!need_split_out)
             lvItem.SubItems(10) = Format(FormatDbValue(rsTemp!settle_month), "yyyy-mm")
         rsTemp.MoveNext
     Next i
     
Exit Sub
ErrHandle:
 ShowErrorMsg
End Sub

'行包签发单
Public Sub ShowSheetInfo()
On Error GoTo ErrHandle
  Dim i As Integer
  Dim nlen As Integer
  Dim rsTemp As Recordset
  Dim lvItem As ListItem
 
     lvLuggageSheet.Visible = True
     Spliter1.Visible = True
'     Spliter1.InitSpliter ptFinSheet, lvLuggageSheet
     Form_Resize
     ptMain_Resize
     '填充lvLuggageSheet
     FilllvLuggageSheetHead
     '填充记录
     If lvFinSheets.ListItems.Count = 0 Then Exit Sub
     m_oFinanceSheet.SheetID = Trim(lvFinSheets.SelectedItem.Text)
     Set rsTemp = m_oFinanceSheet.GetLuggageSheetRS
     lvLuggageSheet.ListItems.clear
     If rsTemp.RecordCount = 0 Then Exit Sub
      
     For i = 1 To rsTemp.RecordCount
         Set lvItem = lvLuggageSheet.ListItems.Add(, , FormatDbValue(rsTemp!sheet_id))
             lvItem.SubItems(1) = FormatDbValue(rsTemp!bus_id)
             lvItem.SubItems(2) = FormatDbValue(rsTemp!bus_date)
             lvItem.SubItems(3) = FormatDbValue(rsTemp!baggage_number)
             lvItem.SubItems(4) = FormatDbValue(rsTemp!luggage_number)
             lvItem.SubItems(5) = FormatDbValue(rsTemp!cal_weight)
             lvItem.SubItems(6) = FormatDbValue(rsTemp!fact_weight)
             lvItem.SubItems(7) = FormatDbValue(rsTemp!over_number)
             lvItem.SubItems(8) = FormatDbValue(rsTemp!price_item_1)
             lvItem.SubItems(9) = FormatDbValue(rsTemp!total_price)
         rsTemp.MoveNext
     Next i
     
Exit Sub
ErrHandle:
 ShowErrorMsg
End Sub
'填充签发单lvLuggageSheet列首
Private Sub FilllvLuggageSheetHead()
        With lvLuggageSheet.ColumnHeaders
          .clear
          .Add , , "签发单号"
          .Add , , "车次"
          .Add , , "车次日期"
          .Add , , "受理单数"
          .Add , , "总件数"
          .Add , , "计重"
          .Add , , "实重"
          .Add , , "超重件数"
          .Add , , "总运费"
          .Add , , "总价"
     End With
End Sub

'显示向导
Public Sub ShowWiz()
   frmWizSplitLuggage.Show vbModal
End Sub
'设置已结
Public Sub ShowSetPayed()
On Error GoTo ErrHandle
  Dim i As Integer
 
  For i = 1 To mSheetNum
      m_oFinanceSheet.Identify mSheetID(i)
      m_oFinanceSheet.SetPayed
      WriteProcessBar True
  Next i
Exit Sub
ErrHandle:
ShowErrorMsg
End Sub

'行包结算单作废
Public Sub ShowSheetCancel()
  On Error GoTo ErrHandle
  Dim i As Integer
  If mSheetNum = 0 Then
    MsgBox "请打勾选择要作废的结算单!", vbInformation, Me.Caption
  Else
    If MsgBox("是否作废行包结算单?", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) <> vbYes Then Exit Sub
  End If
  For i = 1 To mSheetNum
      m_oFinanceSheet.Identify mSheetID(i)
      m_oFinanceSheet.Cancel
      WriteProcessBar True
  Next i
Exit Sub
ErrHandle:
ShowErrorMsg
End Sub

 '行包结算单删除
Public Sub ShowDeleteSheet()
  On Error GoTo ErrHandle
  Dim i As Integer
  If MsgBox("是否删除行包结算单?", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) <> vbYes Then Exit Sub
  For i = 1 To mSheetNum
      m_oFinanceSheet.Identify mSheetID(i)
      m_oFinanceSheet.Delete
      WriteProcessBar True
  Next i
Exit Sub
ErrHandle:
ShowErrorMsg
 
End Sub

Private Sub InitlvFinSheets()
    '设置列首
    lvFinSheets.ColumnHeaders.clear
    With lvFinSheets.ColumnHeaders
        .Add , , "结算单编号", 1600
        .Add , , "状态"
        .Add , , "上车站"
        .Add , , "托运方式"
        .Add , , "结算款"
        .Add , , "应拆金额"
        .Add , , "实拆金额"
        .Add , , "结算月份"
        .Add , , "结算开始日期"
        .Add , , "结算结束日期"
        .Add , , "拆算对象名称"
        .Add , , "拆帐人"
        .Add , , "拆帐日期"
        .Add , , "备注"
    End With
    AlignHeadWidth Me.name, lvFinSheets
End Sub

Private Sub cmdFind_Click()
On Error GoTo ErrHandle
   Dim i As Integer
   Dim j As Integer
   Dim rsTemp As Recordset
   Dim lvItem As ListItem
   Dim szSellstation As String
   Dim szAcceptType As String
   Dim szObject As String
   Dim szStatus As String
   Dim StartDate As Date
   Dim EndDate As Date
'   StartDate = CDate(CStr(Format(dtpDate.Value, "yyyy-mm")) + "-01")
''   EndDate = CDate(CStr(Format(dtpDate.Value, "yyyy-mm")) + "-31")
'   Select Case Month(dtpDate.Value)
'          Case 1, 3, 5, 7, 8, 10, 12
'            EndDate = CDate(Format(dtpDate.Value, "yyyy-mm") & "-31")
'          Case 4, 6, 9, 11
'            EndDate = CDate(Format(dtpDate.Value, "yyyy-mm") & "-30")
'          Case 2
'            EndDate = CDate(Format(dtpDate.Value, "yyyy-mm") & "-28")
'    End Select
    StartDate = dtpStartDate.Value
    EndDate = dtpEndDate.Value
    If Trim(cboSellStation.Text) = "全部" Then
        szSellstation = ""
    Else
       szSellstation = cboSellStation.Text
    End If
    If Trim(cboAcceptType.Text) = "全部" Then
       szAcceptType = ""
    Else
       szAcceptType = cboAcceptType.Text
    End If
    If Trim(txtObject.Text) = "全部" Then
       szObject = ""
    Else
       szObject = txtObject.Text
    End If
    If Trim(cboStatus.Text) = "全部" Then
      szStatus = "全部"
    Else
      szStatus = cboStatus.Text
    End If
    WriteProcessBar True, , , ""
   SetBusy
   lvFinSheets.ListItems.clear  '先清空lvFinsheets
   ShowSBInfo "获得结算单..."
   m_oLugSplitSvr.Init m_oAUser
   Set rsTemp = m_oLugSplitSvr.GetFinanceSheetRS(StartDate, EndDate, Trim(ResolveDisplay(szSellstation)), GetFinTypeInt(Trim(szStatus)), GetLuggageTypeInt(Trim(szAcceptType)), ResolveDisplay(szObject))
   If rsTemp.RecordCount = 0 Then
     SetNormal
     ShowSBInfo ""
     Exit Sub
   End If
   For i = 1 To rsTemp.RecordCount
        Set lvItem = lvFinSheets.ListItems.Add(, , FormatDbValue(rsTemp!fin_sheet_id))

        lvItem.SubItems(1) = GetFinTypeString(FormatDbValue(rsTemp!Status))
        lvItem.SubItems(2) = FormatDbValue(rsTemp!sell_station_name)
        lvItem.SubItems(3) = GetLuggageTypeString(FormatDbValue(rsTemp!accept_type))
        lvItem.SubItems(4) = FormatDbValue(rsTemp!total_price)
        lvItem.SubItems(5) = FormatDbValue(rsTemp!need_split_out)
        lvItem.SubItems(6) = FormatDbValue(rsTemp!act_split_out)
        lvItem.SubItems(7) = Format(rsTemp!settle_month, "yyyy-mm")
        lvItem.SubItems(8) = Format(rsTemp!settlement_start_time, "yyyy-mm-dd")
        lvItem.SubItems(9) = Format(rsTemp!settlement_end_time, "yyyy-mm-dd")
        lvItem.SubItems(10) = FormatDbValue(rsTemp!split_object_name)
        lvItem.SubItems(11) = FormatDbValue(rsTemp!Operator)
        lvItem.SubItems(12) = FormatDbValue(rsTemp!operate_date)
        lvItem.SubItems(13) = FormatDbValue(rsTemp!Annotation)
        
        If GetFinTypeString(FormatDbValue(rsTemp!Status)) = "已结" Then
             SetListViewLineColor lvFinSheets, lvItem.Index, vbBlack
        
        ElseIf GetFinTypeString(FormatDbValue(rsTemp!Status)) = "作废" Then
             SetListViewLineColor lvFinSheets, lvItem.Index, vbRed
        Else
             SetListViewLineColor lvFinSheets, lvItem.Index, vbBlack
        End If
        rsTemp.MoveNext
   Next i
   SetNormal
   ShowSBInfo ""
   WriteProcessBar False, , , ""
   Set rsTemp = Nothing
Exit Sub
ErrHandle:
  SetNormal
  ShowSBInfo ""
  WriteProcessBar False, , , ""
  Set rsTemp = Nothing
  
ShowErrorMsg
End Sub



Private Sub Form_Activate()
    mdiMain.lblTitle = "行包结算单管理"

    Form_Resize
End Sub

Private Sub Form_Deactivate()
    '设置菜单不可用
'    MDIFinance.lblTitle = ""
   
End Sub

Private Sub Form_Load()
 Dim szTemp() As String
 Dim i As Integer
 Dim nlen As Integer
    '初始化
    AlignHeadWidth Me.name, lvFinSheets
'    m_oFinSheet.Init m_oAUser
    '填充查询条件
    '托运方式
    With cboAcceptType
      .clear
      .AddItem "全部"
      .AddItem szAcceptTypeGeneral
      .AddItem szAcceptTypeMan
      .ListIndex = 0
    End With
    '状态
    With cboStatus
      .clear
      .AddItem "全部"
'      .AddItem mStatusNo
      .AddItem mStatusReal
      .AddItem mStatusCancel
      .ListIndex = 0
    End With
    '车站
    cboSellStation.clear
    cboSellStation.AddItem "全部"
    FillSellStation cboSellStation
    cboSellStation.ListIndex = 0
   '查询类型
   '0-拆帐公司 1-车辆 2-参运公司 3-车主 4-车次
    With imgcbo
      .ComboItems.clear
      .ComboItems.Add , , "全部"
      .ComboItems.Add , , "拆帐公司", 1
      .ComboItems.Add , , "车辆", 3
      .ComboItems.Add , , "参运公司", 2
      .ComboItems.Add , , "车主", 5
      .ComboItems.Add , , "车次", 4
      .Text = "全部"
    End With
    Spliter1.InitSpliter ptFinSheet, lvLuggageSheet
    Spliter1.LayoutIt
    dtpStartDate.Value = GetFirstMonthDay(Date)
    dtpEndDate.Value = GetLastMonthDay(Date)
    '初始化lvFinSheets列首
    InitlvFinSheets
    '隐藏lvLuggageSheet控件
    lvLuggageSheet.Visible = False
'    Spliter1.Visible = False
    
End Sub

Private Sub Form_Resize()
    If Not Me.WindowState = vbMaximized Then Me.WindowState = vbMaximized
    '设置查询区位置
    ptShowInfo.Top = 0
    ptShowInfo.Left = 0
    ptShowInfo.Width = frmAllFinSheets.Width
'    cmdFind.Left = ptShowInfo.Width - 3000
   
    '操作区位置设置
    ptMain.Top = ptShowInfo.Height + 20
    ptMain.Left = 0
    ptMain.Width = frmAllFinSheets.Width
    ptMain.Height = frmAllFinSheets.Height
    ptFinSheet.Top = 0
    ptFinSheet.Left = ptMain.Left + 10
    ptFinSheet.Width = ptMain.Width
    ptFinSheet.Height = ptMain.Height - 1350
    ptFinSheet.Refresh
    
    abAction.Top = 0
    abAction.Left = ptFinSheet.Width - abAction.Width - 120
    abAction.Height = ptFinSheet.Height - 50
     '----设置控件的位置
    lvFinSheets.Left = ptFinSheet.Left
    lvFinSheets.Top = 0
    lvFinSheets.Width = ptFinSheet.Width - IIf(abAction.Visible, abAction.Width, 0) - 180
    lvFinSheets.Height = ptFinSheet.Height - 50
    lvFinSheets.Refresh
    
'    '当操作条关闭时间处理
'    If abAction.Visible Then
'        abAction.Move lvFinSheets.Width + 20, lvFinSheets.Top
'        abAction.Height = lvFinSheets.Height
'    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveHeadWidth Me.name, lvFinSheets
    Unload Me
'    mdiMain.lblTitle = ""

End Sub

Private Sub imgcbo_Change()
  If imgcbo.Text = "全部" Then
    txtObject.Text = "全部"
  End If
End Sub

Private Sub imgcbo_Click()
 imgcbo_Change
End Sub

Private Sub lvFinSheets_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
 If lvFinSheets.SortOrder = lvwAscending Then
    lvFinSheets.SortOrder = lvwDescending
 Else
    lvFinSheets.SortOrder = lvwAscending
 End If
    lvFinSheets.SortKey = ColumnHeader.Index - 1
    lvFinSheets.Sorted = True
End Sub

Private Sub lvFinSheets_DblClick()
    ShowFinSheet
End Sub

Private Sub lvFinSheets_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim rsTemp As Recordset
Dim i As Integer
Dim lvItem As ListItem

    If lvLuggageSheet.Visible = True Then
        If lvLuggageSheet.ColumnHeaders.Item(1).Text = "签发单号" Then
            If lvFinSheets.ListItems.Count = 0 Then Exit Sub
            m_oFinanceSheet.SheetID = Trim(lvFinSheets.SelectedItem.Text)
            Set rsTemp = m_oFinanceSheet.GetLuggageSheetRS
            lvLuggageSheet.ListItems.clear
            If rsTemp.RecordCount = 0 Then Exit Sub
             
            For i = 1 To rsTemp.RecordCount
                Set lvItem = lvLuggageSheet.ListItems.Add(, , FormatDbValue(rsTemp!sheet_id))
                    lvItem.SubItems(1) = FormatDbValue(rsTemp!bus_id)
                    lvItem.SubItems(2) = FormatDbValue(rsTemp!bus_date)
                    lvItem.SubItems(3) = FormatDbValue(rsTemp!baggage_number)
                    lvItem.SubItems(4) = FormatDbValue(rsTemp!luggage_number)
                    lvItem.SubItems(5) = FormatDbValue(rsTemp!cal_weight)
                    lvItem.SubItems(6) = FormatDbValue(rsTemp!fact_weight)
                    lvItem.SubItems(7) = FormatDbValue(rsTemp!over_number)
                    lvItem.SubItems(8) = FormatDbValue(rsTemp!price_item_1)
                    lvItem.SubItems(9) = FormatDbValue(rsTemp!total_price)
                rsTemp.MoveNext
            Next i
        Else
            lvLuggageSheet.ListItems.clear
            If lvFinSheets.ListItems.Count = 0 Then Exit Sub
            m_oFinanceSheet.SheetID = Trim(lvFinSheets.SelectedItem.Text)
            Set rsTemp = m_oFinanceSheet.GetFinSheetDetailRS
            If rsTemp.RecordCount = 0 Then Exit Sub
             
            For i = 1 To rsTemp.RecordCount
                Set lvItem = lvLuggageSheet.ListItems.Add(, , FormatDbValue(rsTemp!fin_sheet_id))
                    lvItem.SubItems(1) = FormatDbValue(rsTemp!license_tag_no)
                    lvItem.SubItems(2) = FormatDbValue(rsTemp!owner_name)
                    lvItem.SubItems(3) = FormatDbValue(rsTemp!sell_station_id)  '转换成名称
                    lvItem.SubItems(4) = FormatDbValue(rsTemp!transport_company_short_name)
                    lvItem.SubItems(5) = FormatDbValue(rsTemp!split_company_name)
                    lvItem.SubItems(6) = FormatDbValue(rsTemp!route_name)
                    lvItem.SubItems(7) = FormatDbValue(rsTemp!protocol_name)
                    lvItem.SubItems(8) = FormatDbValue(rsTemp!total_price)
                    lvItem.SubItems(9) = FormatDbValue(rsTemp!need_split_out)
                    lvItem.SubItems(10) = Format(FormatDbValue(rsTemp!settle_month), "yyyy-mm")
                rsTemp.MoveNext
            Next i
        End If
    End If
End Sub

Private Sub lvLuggageSheet_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
 If lvLuggageSheet.SortOrder = lvwAscending Then
    lvLuggageSheet.SortOrder = lvwDescending
 Else
    lvLuggageSheet.SortOrder = lvwAscending
 End If
    lvLuggageSheet.SortKey = ColumnHeader.Index - 1
    lvLuggageSheet.Sorted = True
End Sub

Private Sub ptFinSheet_Resize()
  If lvLuggageSheet.Visible = True Then
    abAction.Top = 0
    abAction.Left = ptFinSheet.Width - abAction.Width - 120
    abAction.Height = ptFinSheet.Height - 50
     '----设置控件的位置
    lvFinSheets.Left = ptFinSheet.Left
    lvFinSheets.Top = 0
    lvFinSheets.Width = ptFinSheet.Width - IIf(abAction.Visible, abAction.Width, 0) - 180
    lvFinSheets.Height = ptFinSheet.Height - 50
    lvFinSheets.Refresh
  End If
End Sub

Private Sub ptMain_Resize()
 If lvLuggageSheet.Visible = True Then
   
   lvLuggageSheet.Height = 2450
   ptFinSheet.Height = ptFinSheet.Height - lvLuggageSheet.Height - Spliter1.Height
   ptFinSheet_Resize
   
   Spliter1.Top = ptFinSheet.Height + 10
   Spliter1.Left = 0
   Spliter1.Width = ptMain.Width
   
   lvLuggageSheet.Left = ptMain.Left
   lvLuggageSheet.Width = ptMain.Width - 50
   lvLuggageSheet.Top = ptFinSheet.Height + Spliter1.Height
  

 End If
End Sub

Private Sub txtObject_ButtonClick()
On Error GoTo ErrHandle
    Dim oShell As New STShell.CommDialog
    Dim aszTemp() As String

    oShell.Init m_oAUser
    If Trim(imgcbo.Text) = "拆帐公司" Or Trim(imgcbo.Text) = "参运公司" Then
       aszTemp = oShell.SelectCompany(False)
    ElseIf Trim(imgcbo.Text) = "车辆" Then
       aszTemp = oShell.SelectVehicleEX()
    ElseIf Trim(imgcbo.Text) = "车主" Then
       aszTemp = oShell.SelectOwner(, False)
    ElseIf Trim(imgcbo.Text) = "车次" Then
       aszTemp = oShell.SelectBus(False)
    End If
    
    Set oShell = Nothing
    If ArrayLength(aszTemp) = 0 Then Exit Sub
    txtObject.Text = Trim(aszTemp(1, 1))
    
Exit Sub
ErrHandle:
ShowErrorMsg
End Sub

