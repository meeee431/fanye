VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "actbar2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Begin VB.Form frmAllSettleSheets 
   BackColor       =   &H00E0E0E0&
   Caption         =   "路单结算"
   ClientHeight    =   8145
   ClientLeft      =   525
   ClientTop       =   2565
   ClientWidth     =   11850
   HelpContextID   =   7000390
   Icon            =   "frmAllSettleSheets.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8145
   ScaleWidth      =   11850
   WindowState     =   2  'Maximized
   Begin VB.PictureBox ptMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7335
      Left            =   0
      ScaleHeight     =   7335
      ScaleWidth      =   10995
      TabIndex        =   12
      Top             =   1140
      Width           =   10995
      Begin VB.PictureBox ptFinSheet 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   5535
         Left            =   90
         ScaleHeight     =   5505
         ScaleWidth      =   10635
         TabIndex        =   14
         Top             =   180
         Width           =   10665
         Begin ActiveBar2LibraryCtl.ActiveBar2 abAction 
            Height          =   5280
            Left            =   9090
            TabIndex        =   15
            Top             =   60
            Width           =   1530
            _LayoutVersion  =   1
            _ExtentX        =   2699
            _ExtentY        =   9313
            _DataPath       =   ""
            Bands           =   "frmAllSettleSheets.frx":000C
         End
         Begin MSComctlLib.ListView lvSettleSheets 
            Height          =   4635
            Left            =   180
            TabIndex        =   16
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
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
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
      ScaleWidth      =   14145
      TabIndex        =   10
      Top             =   0
      Width           =   14145
      Begin VB.TextBox txtCheckSheetID 
         Height          =   270
         Left            =   8565
         TabIndex        =   23
         Top             =   142
         Width           =   1185
      End
      Begin VB.TextBox txtObjectName 
         Height          =   270
         Left            =   8580
         TabIndex        =   21
         Top             =   540
         Width           =   1200
      End
      Begin VB.TextBox txtSettleSheet 
         Height          =   270
         Left            =   6435
         TabIndex        =   19
         Top             =   142
         Width           =   1230
      End
      Begin VB.ComboBox cboStatus 
         Height          =   300
         Left            =   6450
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   540
         Width           =   1155
      End
      Begin VB.ComboBox cboSellStation 
         Height          =   300
         ItemData        =   "frmAllSettleSheets.frx":3886
         Left            =   11685
         List            =   "frmAllSettleSheets.frx":3888
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   705
         Visible         =   0   'False
         Width           =   1155
      End
      Begin RTComctl3.CoolButton cmdFind 
         Default         =   -1  'True
         Height          =   345
         Left            =   9870
         TabIndex        =   13
         Top             =   90
         Width           =   1125
         _ExtentX        =   1984
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
         MICON           =   "frmAllSettleSheets.frx":388A
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
         Left            =   3690
         TabIndex        =   1
         Top             =   120
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy年M月dd日"
         Format          =   73465856
         CurrentDate     =   37704
      End
      Begin FText.asFlatTextBox txtObject 
         Height          =   315
         Left            =   11940
         TabIndex        =   9
         Top             =   345
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
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
         ButtonHotBackColor=   14737632
         ButtonPressedBackColor=   14737632
         Text            =   "全部"
         ButtonBackColor =   14737632
         ButtonVisible   =   -1  'True
         Registered      =   -1  'True
      End
      Begin MSComctlLib.ImageCombo imgcbo 
         Height          =   315
         Left            =   11940
         TabIndex        =   3
         Top             =   -75
         Visible         =   0   'False
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Locked          =   -1  'True
         ImageList       =   "imglv"
      End
      Begin MSComCtl2.DTPicker dtpEndDate 
         Height          =   315
         Left            =   3690
         TabIndex        =   7
         Top             =   540
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy年M月dd日"
         Format          =   73465856
         CurrentDate     =   37704
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "路单号:"
         Height          =   180
         Left            =   7740
         TabIndex        =   22
         Top             =   187
         Width           =   630
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "对象名:"
         Height          =   180
         Left            =   7740
         TabIndex        =   20
         Top             =   570
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "结算单号:"
         Height          =   180
         Left            =   5595
         TabIndex        =   18
         Top             =   187
         Width           =   810
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "车站(&S):"
         Height          =   180
         Left            =   10845
         TabIndex        =   4
         Top             =   765
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "状态(&E):"
         Height          =   180
         Left            =   5625
         TabIndex        =   17
         Top             =   600
         Width           =   720
      End
      Begin VB.Label lblType 
         BackStyle       =   0  'Transparent
         Caption         =   "对象(&O):"
         Height          =   180
         Left            =   10860
         TabIndex        =   8
         Top             =   435
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "查询类型(&P):"
         Height          =   180
         Left            =   10830
         TabIndex        =   2
         Top             =   15
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "结算结束日期(&T):"
         Height          =   180
         Left            =   2100
         TabIndex        =   6
         Top             =   600
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "结算开始日期(&D):"
         Height          =   180
         Left            =   2130
         TabIndex        =   0
         Top             =   187
         Width           =   1440
      End
      Begin VB.Image Image1 
         Height          =   1275
         Left            =   0
         Picture         =   "frmAllSettleSheets.frx":38A6
         Top             =   30
         Width           =   2010
      End
   End
   Begin MSComctlLib.ImageList imglv 
      Left            =   11220
      Top             =   3435
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
            Picture         =   "frmAllSettleSheets.frx":4D79
            Key             =   "cancel"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAllSettleSheets.frx":5113
            Key             =   "company"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAllSettleSheets.frx":526D
            Key             =   "vehicle"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAllSettleSheets.frx":53C7
            Key             =   "valid"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAllSettleSheets.frx":5761
            Key             =   "settled"
         EndProperty
      EndProperty
   End
   Begin VB.Menu pmnu_SettleSheet 
      Caption         =   "行包结算单"
      Visible         =   0   'False
      Begin VB.Menu pmnu_Property 
         Caption         =   "属性(&P)"
      End
      Begin VB.Menu pmnu_Cancel 
         Caption         =   "作废(&C)"
      End
      Begin VB.Menu pmnu_ViewSettleSheet 
         Caption         =   "结算单(&T)"
      End
      Begin VB.Menu pmnu_Break1 
         Caption         =   "-"
      End
      Begin VB.Menu pmnu_Settle 
         Caption         =   "汇款(&S)"
      End
      Begin VB.Menu pmnu_Break2 
         Caption         =   "-"
      End
      Begin VB.Menu pmnu_NewWizard 
         Caption         =   "新增向导(&W)"
      End
   End
End
Attribute VB_Name = "frmAllSettleSheets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'列表subItem常数
Const PI_SettleSheetID = 0
Const PI_Status = 1
Const PI_SheetQuantity = 2
Const PI_TotalTicketPrice = 3
Const PI_TotalQuantity = 4
Const PI_SettlePrice = 5
Const PI_SettleStationPrice = 6
Const PI_SettleObject = 7
Const PI_SettleObjectName = 8
Const PI_TransportCompanyName = 9
Const PI_RouteName = 10
Const PI_Settler = 11
Const PI_Checker = 12
Const PI_SettleDate = 13
Const PI_StartDate = 14
Const PI_EndDate = 15
Const PI_UnitName = 16
Const PI_Annotation = 17

Const cszTemplateFile = "银行结算报表.xls"
Const cszTemplateFileName = "银行结算报表(车次).xls"



Private Sub abAction_BandClose(ByVal Band As ActiveBar2LibraryCtl.Band)
    '设置控件的位置
    lvSettleSheets.Left = 0 ' 50
    lvSettleSheets.Top = 0 'Me.Height + 50
    lvSettleSheets.Width = ptMain.Width - IIf(abAction.Visible, abAction.Width, 0) - 20
    lvSettleSheets.Height = ptMain.Height - 20
    lvSettleSheets.Refresh
    '当操作条关闭时间处理
    If abAction.Visible Then
        abAction.Move lvSettleSheets.Width + 20, lvSettleSheets.Top
        abAction.Height = lvSettleSheets.Height
    End If

End Sub



Private Sub abAction_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)

    Select Case Tool.name
    Case "Property"  '属性
        ShowSettleSheet
'        RefreshFind
    Case "NewWizard"  '新增向导
        ShowWiz
'        RefreshFind
    Case "Cancel" '作废
        CancelSettleSheet
    Case "SetRemit" '汇款
        SetRemit
'        RefreshFind
    Case "ViewSettleSheet" '结算单
        ViewSettleSheet
    Case "CancelRemit" '汇款作废
        CancelRemit
    End Select
End Sub

'作废结算单
Private Sub CancelSettleSheet()
On Error GoTo here
    Dim m_oSettleSheet As New SettleSheet

    
    If lvSettleSheets.SelectedItem Is Nothing Then Exit Sub
    
    m_oSettleSheet.Init g_oActiveUser
    m_oSettleSheet.Identify Trim(lvSettleSheets.SelectedItem.Text)
    If MsgBox("是否作废此结算单!", vbInformation + vbYesNo, Me.Caption) = vbYes Then
        m_oSettleSheet.CancelSettleSheet Trim(lvSettleSheets.SelectedItem.Text), lvSettleSheets.SelectedItem.SubItems(PI_SettleObject)
    End If
    UpdateList
    Exit Sub
here:
    ShowErrorMsg
End Sub

Private Sub UpdateList()
    If lvSettleSheets.SelectedItem Is Nothing Then Exit Sub
    Dim liTemp As ListItem
    Dim oReport As New Report
    Dim TSettleSheet() As TSettleSheet
    Dim nCount As Integer
    Dim i As Integer
    
    
    Set liTemp = lvSettleSheets.SelectedItem
    
    oReport.Init g_oActiveUser
   TSettleSheet = oReport.GetSettleSheetInfo(, , , , , lvSettleSheets.SelectedItem.Text)
   nCount = ArrayLength(TSettleSheet)
    If nCount = 0 Then
        Exit Sub
    End If
    i = 1
'   For i = 1 To nCount
'        If TSettleSheet(i).Status = CS_SettleSheetValid Then
'            Set liTemp = lvSettleSheets.ListItems.Add(, , TSettleSheet(i).SettleSheetID, , "valid")
'        ElseIf TSettleSheet(i).Status = CS_SettleSheetInvalid Then
'            Set liTemp = lvSettleSheets.ListItems.Add(, , TSettleSheet(i).SettleSheetID, , "cancel")
'        ElseIf TSettleSheet(i).Status = CS_SettleSheetSettled Then
'            Set liTemp = lvSettleSheets.ListItems.Add(, , TSettleSheet(i).SettleSheetID, , "settled")
'        Else
'            Set liTemp = lvSettleSheets.ListItems.Add(, , TSettleSheet(i).SettleSheetID, , "valid")
'        End If
        liTemp.SubItems(PI_Status) = GetSettleSheetStatusString(CInt(TSettleSheet(i).Status))    '转化
        liTemp.SubItems(PI_SheetQuantity) = TSettleSheet(i).CheckSheetCount
        liTemp.SubItems(PI_TotalTicketPrice) = TSettleSheet(i).TotalTicketPrice
        liTemp.SubItems(PI_TotalQuantity) = TSettleSheet(i).TotalQuantity
        liTemp.SubItems(PI_SettlePrice) = FormatMoney(TSettleSheet(i).SettleOtherCompanyPrice - TSettleSheet(i).SettleStationPrice)  '实结金额
        liTemp.SubItems(PI_SettleStationPrice) = TSettleSheet(i).SettleStationPrice
        liTemp.SubItems(PI_SettleObject) = GetObjectTypeString(CInt(TSettleSheet(i).SettleObject)) ' TSettleSheet(i).ObjectID  ' GetObjectTypeString(CInt(TSettleSheet(i).SettleObject))
        
        
        liTemp.SubItems(PI_SettleObjectName) = TSettleSheet(i).ObjectName
        liTemp.SubItems(PI_TransportCompanyName) = TSettleSheet(i).TransportCompanyName
        liTemp.SubItems(PI_Settler) = TSettleSheet(i).Settler
        liTemp.SubItems(PI_Checker) = TSettleSheet(i).Checker
        liTemp.SubItems(PI_SettleDate) = TSettleSheet(i).SettleDate
        liTemp.SubItems(PI_StartDate) = Format(TSettleSheet(i).SettleStartDate, "yyyy-MM-dd")
        liTemp.SubItems(PI_EndDate) = Format(TSettleSheet(i).SettleEndDate, "yyyy-MM-dd")
        liTemp.SubItems(PI_UnitName) = TSettleSheet(i).UnitName
        liTemp.SubItems(PI_Annotation) = TSettleSheet(i).Annotation
        
        liTemp.SubItems(PI_RouteName) = TSettleSheet(i).RouteName
        
        
'        If TSettleSheet(i).Status = CS_SettleSheetInvalid Then
'            SetListViewLineColor lvSettleSheets, liTemp.Index, vbRed
'        Else
'            SetListViewLineColor lvSettleSheets, liTemp.Index, vbBlack
'        End If
    
End Sub



'属性
Private Sub ShowSettleSheet()
    If lvSettleSheets.ListItems.Count = 0 Then Exit Sub
    
    frmSettleAttrib.m_szSettleSheetID = Trim(lvSettleSheets.SelectedItem.Text)
    frmSettleAttrib.ZOrder 0
    frmSettleAttrib.Show vbModal

End Sub
'填充结算信息列首
Private Sub FilllvSettleSheetsHead()
    With lvSettleSheets.ColumnHeaders
          .Clear
          .Add , , "结算单号"
          .Add , , "状态"
          .Add , , "路单数"
          .Add , , "总票款"
          .Add , , "人数"
          .Add , , "应结票款"
          .Add , , "结给车站"
          .Add , , "结算对象"
          .Add , , "对象名称"
          
          .Add , , "参运公司"
          .Add , , "线路"
          
          .Add , , "拆帐人"
          .Add , , "复核人", 0
          .Add , , "拆帐日期"
          .Add , , "开始时间"
          .Add , , "结束时间"
          .Add , , "单位名称", 0
          .Add , , "注释"
          
     End With
     AlignHeadWidth Me.name, lvSettleSheets
End Sub



'显示向导
Public Sub ShowWiz()
    frmWizSplitSettle.ZOrder 0
    frmWizSplitSettle.Show vbModal
End Sub



Private Sub cmdFind_Click()
On Error GoTo ErrHandle
   RefreshFind
Exit Sub
ErrHandle:
ShowErrorMsg
End Sub
'查询结算单
Public Sub RefreshFind()
On Error GoTo ErrHandle
   Dim i As Integer
   Dim m_oReport As New Report
   Dim TSettleSheet() As TSettleSheet
   Dim nCount As Integer
   Dim lvItem As ListItem
   Dim dtStartDate As Date
   Dim dtEndDate As Date
   Dim eObjectType As ESettleObjectType
   Dim szObject As String
   Dim szSellStation As String
   Dim eStatus As ESettleSheetStatus
   m_oReport.Init g_oActiveUser
   dtStartDate = dtpStartDate.Value
   dtEndDate = DateAdd("d", 1, dtpEndDate.Value)
   If Trim(imgcbo.Text) <> "全部" Then
        eObjectType = GetObjectTypeInt(Trim(imgcbo.Text))
   Else
        eObjectType = GetObjectTypeInt(Trim(imgcbo.Text))
   End If
   If txtObject.Text = "全部" Or txtObject.Text = "" Then
        szObject = ""
   Else
        szObject = ResolveDisplay(txtObject.Text)
   End If
   If cboSellStation.Text <> "" Then
        szSellStation = ResolveDisplay(cboSellStation.Text)
   Else
        szSellStation = ""
   End If
   If cboStatus.Text <> "" Then
        eStatus = GetSettleSheetStatusInt(Trim(cboStatus.Text))
   Else
        eStatus = -1
   End If
   
   TSettleSheet = m_oReport.GetSettleSheetInfo(eStatus, dtStartDate, dtEndDate, eObjectType, szObject, txtSettleSheet.Text, txtObjectName.Text, txtCheckSheetID.Text)
   nCount = ArrayLength(TSettleSheet)
    If nCount = 0 Then
        lvSettleSheets.ListItems.Clear
        Exit Sub
    End If
   lvSettleSheets.ListItems.Clear
   For i = 1 To nCount
        If TSettleSheet(i).Status = CS_SettleSheetValid Then
            Set lvItem = lvSettleSheets.ListItems.Add(, , TSettleSheet(i).SettleSheetID, , "valid")
        ElseIf TSettleSheet(i).Status = CS_SettleSheetInvalid Then
            Set lvItem = lvSettleSheets.ListItems.Add(, , TSettleSheet(i).SettleSheetID, , "cancel")
        ElseIf TSettleSheet(i).Status = CS_SettleSheetSettled Then
            Set lvItem = lvSettleSheets.ListItems.Add(, , TSettleSheet(i).SettleSheetID, , "settled")
        Else
            Set lvItem = lvSettleSheets.ListItems.Add(, , TSettleSheet(i).SettleSheetID, , "valid")
        End If
        lvItem.SubItems(PI_Status) = GetSettleSheetStatusString(CInt(TSettleSheet(i).Status))    '转化
        lvItem.SubItems(PI_SheetQuantity) = TSettleSheet(i).CheckSheetCount
        lvItem.SubItems(PI_TotalTicketPrice) = TSettleSheet(i).TotalTicketPrice
        lvItem.SubItems(PI_TotalQuantity) = TSettleSheet(i).TotalQuantity
        lvItem.SubItems(PI_SettlePrice) = FormatMoney(TSettleSheet(i).SettleOtherCompanyPrice - TSettleSheet(i).SettleStationPrice)  '实结金额
        lvItem.SubItems(PI_SettleStationPrice) = TSettleSheet(i).SettleStationPrice
        lvItem.SubItems(PI_SettleObject) = GetObjectTypeString(CInt(TSettleSheet(i).SettleObject)) ' TSettleSheet(i).ObjectID  ' GetObjectTypeString(CInt(TSettleSheet(i).SettleObject))
        
        
        lvItem.SubItems(PI_SettleObjectName) = TSettleSheet(i).ObjectName
        lvItem.SubItems(PI_TransportCompanyName) = TSettleSheet(i).TransportCompanyName
        lvItem.SubItems(PI_Settler) = TSettleSheet(i).Settler
        lvItem.SubItems(PI_Checker) = TSettleSheet(i).Checker
        lvItem.SubItems(PI_SettleDate) = TSettleSheet(i).SettleDate
        lvItem.SubItems(PI_StartDate) = Format(TSettleSheet(i).SettleStartDate, "yyyy-MM-dd")
        lvItem.SubItems(PI_EndDate) = Format(TSettleSheet(i).SettleEndDate, "yyyy-MM-dd")
        lvItem.SubItems(PI_UnitName) = TSettleSheet(i).UnitName
        lvItem.SubItems(PI_Annotation) = TSettleSheet(i).Annotation
        
        lvItem.SubItems(PI_RouteName) = TSettleSheet(i).RouteName
        
        
'        If TSettleSheet(i).Status = CS_SettleSheetInvalid Then
'            SetListViewLineColor lvSettleSheets, lvItem.Index, vbRed
'        Else
'            SetListViewLineColor lvSettleSheets, lvItem.Index, vbBlack
'        End If
   Next i
   
Exit Sub
ErrHandle:
ShowErrorMsg
End Sub




Private Sub Form_Activate()
Form_Resize
End Sub

Private Sub Form_Load()
 Dim szTemp() As String
 Dim i As Integer
 Dim nLen As Integer
    '初始化
    
    FormSize
    '填充查询条件
    '状态
    With cboStatus
        .Clear
        .AddItem "全部"

        
        
        .AddItem GetSettleSheetStatusString(CS_SettleSheetValid) '"未结"
        .AddItem GetSettleSheetStatusString(CS_SettleSheetInvalid)  '"作废"
        .AddItem GetSettleSheetStatusString(CS_SettleSheetSettled)   '"已结"
'        .AddItem GetSettleSheetStatusString(CS_SettleSheetNegativeNotPay)    '应扣款未结清
'        .AddItem GetSettleSheetStatusString(CS_SettleSheetNegativeHasPayed)    '应扣款已结清
        
        .ListIndex = 0
    End With
    '车站
    cboSellStation.Clear
    cboSellStation.AddItem "全部"
    FillSellStation cboSellStation
    cboSellStation.ListIndex = 0
    '   '查询类型
    '0-拆帐公司 1-车辆 2-参运公司 3-车主 4-车次
    With imgcbo
        .ComboItems.Clear
        .ComboItems.Add , , "全部"
        .ComboItems.Add , , "参运公司", "company"
        .ComboItems.Add , , "车辆", "vehicle"
        .Text = "全部"
    End With
    
    dtpStartDate.Value = GetFirstMonthDay(Date)
    dtpEndDate.Value = GetLastMonthDay(Date)
    
    FilllvSettleSheetsHead
    
End Sub

'初始化界面
Private Sub FormSize()
    ptShowInfo.Top = 0
    ptShowInfo.Left = 0
    ptShowInfo.Width = Me.ScaleWidth
    
    ptMain.Top = ptShowInfo.Height
    ptMain.Left = 0
    ptMain.Width = Me.ScaleWidth
    ptMain.Height = mdiMain.ScaleHeight - ptShowInfo.Height - 50
    
    ptFinSheet.Top = 0
    ptFinSheet.Left = ptMain.Left
    ptFinSheet.Height = ptMain.Height
    ptFinSheet.Width = ptMain.Width
    
    lvSettleSheets.Top = 0
    lvSettleSheets.Left = ptFinSheet.Left
    lvSettleSheets.Height = ptFinSheet.Height
    lvSettleSheets.Width = ptFinSheet.Width - abAction.Width - 50
    
    abAction.Top = lvSettleSheets.Top
    abAction.Left = lvSettleSheets.Width + 50
    abAction.Height = lvSettleSheets.Height
End Sub

Private Sub Form_Resize()
    FormSize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveHeadWidth Me.name, lvSettleSheets
    Unload Me
End Sub


Private Sub imgcbo_Click()
    If imgcbo.Text = "全部" Then
        txtObject.Text = "全部"
    End If
End Sub

Private Sub lvSettleSheets_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
 If lvSettleSheets.SortOrder = lvwAscending Then
    lvSettleSheets.SortOrder = lvwDescending
 Else
    lvSettleSheets.SortOrder = lvwAscending
 End If
    lvSettleSheets.SortKey = ColumnHeader.Index - 1
    lvSettleSheets.Sorted = True
End Sub

Private Sub lvSettleSheets_DblClick()
    ShowSettleSheet
End Sub


Private Sub lvSettleSheets_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    '需要循环判断
'    If Item.Checked Then
'        pmnu_Settle.Enabled = True
'    Else
'        pmnu_Settle.Enabled = False
'    End If
End Sub

Private Sub lvSettleSheets_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu pmnu_SettleSheet
        
    End If
End Sub

Private Sub pmnu_Cancel_Click()
    CancelSettleSheet
End Sub

Private Sub pmnu_NewWizard_Click()
    ShowWiz
    
End Sub

Private Sub pmnu_Property_Click()
    ShowSettleSheet
End Sub

Private Sub pmnu_Settle_Click()
    '汇款
    SetRemit
    
    
    
    
End Sub

Private Sub pmnu_ViewSettleSheet_Click()
    ViewSettleSheet
End Sub

Private Sub txtObject_ButtonClick()
On Error GoTo ErrHandle
    Dim oShell As New STShell.CommDialog
    Dim aszTemp() As String
    oShell.Init g_oActiveUser
    If Trim(imgcbo.Text) = "参运公司" Then
       aszTemp = oShell.SelectCompany()
    ElseIf Trim(imgcbo.Text) = "车辆" Then
       aszTemp = oShell.SelectVehicle()
    End If

    Set oShell = Nothing
    If ArrayLength(aszTemp) = 0 Then Exit Sub
    txtObject.Text = MakeDisplayString(Trim(aszTemp(1, 1)), Trim(aszTemp(1, 2)))

Exit Sub
ErrHandle:
ShowErrorMsg
End Sub




Private Sub SetRemit()
    '设置汇款
    On Error GoTo ErrHandle
    
    
    Dim oSplit As New Split
    Dim aszTemp() As String
    Dim i As Integer
    Dim nCount As Integer
    nCount = 0
    
    If MsgBox("是否确实要对已选择的结算单进行汇款？", vbQuestion + vbYesNo, "汇款") = vbNo Then Exit Sub
    
    
    For i = 1 To lvSettleSheets.ListItems.Count
        If lvSettleSheets.ListItems(i).Checked Then
            nCount = nCount + 1
            ReDim Preserve aszTemp(1 To nCount)
            aszTemp(nCount) = lvSettleSheets.ListItems(i).Text
        End If
    Next i
    
    If nCount > 0 Then
        '如果有选择的结算单
        
        oSplit.Init g_oActiveUser
        oSplit.SetRemit aszTemp
        
        ShowRemitLst aszTemp
    End If
     
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub CancelRemit()
    '作废汇款
    On Error GoTo ErrHandle
    
    
    Dim oSplit As New Split
    Dim aszTemp() As String
    Dim i As Integer
    Dim nCount As Integer
    nCount = 0
    
    If MsgBox("是否确实要对已选择的结算单进行汇款作废？", vbQuestion + vbYesNo, "汇款作废") = vbNo Then Exit Sub
    
    
    For i = 1 To lvSettleSheets.ListItems.Count
        If lvSettleSheets.ListItems(i).Checked Then
            nCount = nCount + 1
            ReDim Preserve aszTemp(1 To nCount)
            aszTemp(nCount) = lvSettleSheets.ListItems(i).Text
        End If
    Next i
    
    If nCount > 0 Then
        '如果有选择的结算单
        
        oSplit.Init g_oActiveUser
        oSplit.CancelRemit aszTemp
        
    End If
     
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub


'显示汇款报表
Private Sub ShowRemitLst(paszTemp() As String)
    On Error GoTo ErrHandle
    
    Dim lHelpContextID As Long
    Dim rsTemp As Recordset

    Dim oReport As New Report
    Dim vCustomData As Variant
    Dim frmNewReport As New frmReport
    
    oReport.Init g_oActiveUser
    
    If lvSettleSheets.SelectedItem.SubItems(PI_SettleObject) = "车次" Then
'        '取得记录集
'        Set rsTemp = oReport.BusSettleDetailEX(paszTemp)
'        WriteProcessBar True, , , ""
'        ReDim vCustomData(1 To 4, 1 To 2)
'        vCustomData(1, 1) = "开始日期"
'        vCustomData(2, 1) = "结束日期"
'        If rsTemp.RecordCount > 0 Then
'            vCustomData(1, 2) = FormatDbValue(rsTemp!start_date)
'
'            vCustomData(2, 2) = FormatDbValue(rsTemp!end_date)
'        End If
'        vCustomData(3, 1) = "打印"
'        vCustomData(3, 2) = g_oActiveUser.UserName
'        vCustomData(4, 1) = "单位"
'        vCustomData(4, 2) = g_oActiveUser.UserUnitName
'
'        frmNewReport.ShowReport rsTemp, cszTemplateFileName, "银行汇款清单(车次)", vCustomData, 10
    Else
        '取得记录集
        Set rsTemp = oReport.VehicleSettleDetailEX(paszTemp)
        WriteProcessBar True, , , ""
        ReDim vCustomData(1 To 4, 1 To 2)
        vCustomData(1, 1) = "开始日期"
        vCustomData(2, 1) = "结束日期"
        If rsTemp.RecordCount > 0 Then
            vCustomData(1, 2) = FormatDbValue(rsTemp!start_date)
            
            vCustomData(2, 2) = FormatDbValue(rsTemp!end_date)
        End If
        vCustomData(3, 1) = "打印"
        vCustomData(3, 2) = g_oActiveUser.UserName
        vCustomData(4, 1) = "单位"
        vCustomData(4, 2) = g_oActiveUser.UserUnitName
        
        frmNewReport.ShowReport rsTemp, cszTemplateFile, "银行汇款清单", vCustomData, 10
    End If
    
    WriteProcessBar False, , , ""
    
    
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub




Private Sub ViewSettleSheet()
    '打印结算单
    If lvSettleSheets.SelectedItem Is Nothing Then Exit Sub
    
    frmPrintFinSheet.m_SheetID = Trim(lvSettleSheets.SelectedItem.Text)
    frmPrintFinSheet.m_szLugSettleSheetID = ""
    frmPrintFinSheet.m_bRePrint = False
'    frmPrintFinSheet.m_bNeedPrint = False
    
    frmPrintFinSheet.ZOrder 0
    frmPrintFinSheet.Show vbModal
End Sub
