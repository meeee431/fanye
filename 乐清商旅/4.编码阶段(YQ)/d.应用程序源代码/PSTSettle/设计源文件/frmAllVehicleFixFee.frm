VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "actbar2.ocx"
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Begin VB.Form frmAllVehicleFixFee 
   Caption         =   "固定费用"
   ClientHeight    =   7845
   ClientLeft      =   2205
   ClientTop       =   2145
   ClientWidth     =   11475
   Icon            =   "frmAllVehicleFixFee.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7845
   ScaleWidth      =   11475
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   8610
      Top             =   2235
   End
   Begin VB.PictureBox ptShowInfo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1155
      Left            =   0
      ScaleHeight     =   1155
      ScaleWidth      =   15135
      TabIndex        =   1
      Top             =   0
      Width           =   15135
      Begin VB.ComboBox cboIsDec 
         Height          =   300
         Left            =   8325
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   255
         Width           =   1365
      End
      Begin VB.ComboBox cboLicenseTagNO 
         Height          =   300
         Left            =   2850
         TabIndex        =   13
         Top             =   255
         Width           =   1650
      End
      Begin MSComCtl2.DTPicker dtpEndDate 
         Height          =   330
         Left            =   5895
         TabIndex        =   10
         Top             =   735
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   582
         _Version        =   393216
         Format          =   23789568
         CurrentDate     =   38553
      End
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   300
         Left            =   2850
         TabIndex        =   8
         Top             =   750
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   529
         _Version        =   393216
         Format          =   23789568
         CurrentDate     =   38553
      End
      Begin FText.asFlatTextBox txtVehicle 
         Height          =   300
         Left            =   5580
         TabIndex        =   2
         Top             =   255
         Width           =   1695
         _ExtentX        =   2990
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
         ButtonHotBackColor=   -2147483633
         ButtonPressedBackColor=   -2147483627
         Text            =   ""
         ButtonBackColor =   -2147483633
         ButtonVisible   =   -1  'True
         Registered      =   -1  'True
      End
      Begin FText.asFlatTextBox txtCompany 
         Height          =   285
         Left            =   8790
         TabIndex        =   3
         Top             =   750
         Width           =   2235
         _ExtentX        =   3942
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
         Text            =   ""
         ButtonBackColor =   -2147483633
         ButtonVisible   =   -1  'True
         Registered      =   -1  'True
      End
      Begin RTComctl3.CoolButton cmdQuery 
         Default         =   -1  'True
         Height          =   375
         Left            =   9945
         TabIndex        =   4
         Top             =   180
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   661
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
         MICON           =   "frmAllVehicleFixFee.frx":000C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "状态:"
         Height          =   180
         Left            =   7425
         TabIndex        =   14
         Top             =   315
         Width           =   450
      End
      Begin VB.Label lblLicense 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车牌号:"
         Height          =   180
         Left            =   1905
         TabIndex        =   12
         Top             =   315
         Width           =   630
      End
      Begin VB.Label lblEndDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "结束日期:"
         Height          =   180
         Left            =   4770
         TabIndex        =   9
         Top             =   810
         Width           =   810
      End
      Begin VB.Label lblStartDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "起始日期:"
         Height          =   180
         Left            =   1905
         TabIndex        =   7
         Top             =   810
         Width           =   810
      End
      Begin VB.Label lblCompany 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "参运公司:"
         Height          =   180
         Left            =   7890
         TabIndex        =   6
         Top             =   795
         Width           =   810
      End
      Begin VB.Label lblVehicleID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车辆代码:"
         Height          =   180
         Left            =   4650
         TabIndex        =   5
         Top             =   315
         Width           =   810
      End
      Begin VB.Image Image2 
         Height          =   1275
         Left            =   60
         Picture         =   "frmAllVehicleFixFee.frx":0028
         Top             =   150
         Width           =   2010
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid vsFixFee 
      Height          =   5175
      Left            =   930
      TabIndex        =   0
      Top             =   1275
      Width           =   7035
      _cx             =   12409
      _cy             =   9128
      _ConvInfo       =   -1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmAllVehicleFixFee.frx":14FB
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin ActiveBar2LibraryCtl.ActiveBar2 abAction 
      Height          =   4845
      Left            =   9480
      TabIndex        =   11
      Top             =   1365
      Width           =   1485
      _LayoutVersion  =   1
      _ExtentX        =   2619
      _ExtentY        =   8546
      _DataPath       =   ""
      Bands           =   "frmAllVehicleFixFee.frx":15D0
   End
End
Attribute VB_Name = "frmAllVehicleFixFee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'所有的固定费用项是从系统参数中读取的,用逗号分隔的。
'并需要判断，该项是否是上个月的负数自动统计到本月的项目，哪个项目是上个月的负数存放项目，从系统参数中读取

Const cnVehicle = 1
Const cnDate = 2
Const cnCompany = 3
Const cnIsDec = 4 '是否已扣
Const cnSplitItemStart = 5

Public m_szFormStatus As FixFeeStatus

Private m_aszFixFeeItem() As String
Private m_atTemp() As TSplitItemInfo

Private m_oReport As New Report

'界面排列
Private Sub AlignForm()

    ptShowInfo.Top = 0
    ptShowInfo.Left = 0
    ptShowInfo.Width = mdiMain.ScaleWidth
    
    vsFixFee.Top = ptShowInfo.Height + 50
    vsFixFee.Left = 50
    vsFixFee.Width = mdiMain.ScaleWidth - abAction.Width - 50
    vsFixFee.Height = mdiMain.ScaleHeight - 50

    abAction.Top = vsFixFee.Top
    abAction.Left = vsFixFee.Width + 50
    abAction.Height = vsFixFee.Height
End Sub


Private Sub cmdQuery_Click()
    QueryFixFee
End Sub

Private Sub Form_Load()


    
    m_oReport.Init g_oActiveUser
    
    
    AlignForm
    GetSplitItem
    FillHead
    AlignHeadWidth Me.name, vsFixFee
    
    cboIsDec.AddItem MakeDisplayString(-1, GetFixFeeStatusName(-1))
    cboIsDec.AddItem MakeDisplayString(0, GetFixFeeStatusName(0))
    cboIsDec.AddItem MakeDisplayString(1, GetFixFeeStatusName(1))
    
    cboIsDec.ListIndex = 1
    
    
End Sub

'得到所有的结算项
Private Sub GetSplitItem()
    
    m_atTemp = m_oReport.GetSplitItemInfo(, True)
    
End Sub


Private Sub FillHead()


    Dim i As Integer
    Dim nCols As Integer
    Dim j As Integer
    '得到固定费用项
    m_aszFixFeeItem = Split(g_szFixFeeItem, ",") '将项目分解出来
    nCols = ArrayLength(m_aszFixFeeItem)
    
    
    With vsFixFee
        .Cols = cnSplitItemStart + nCols
        .Rows = 1
        .AllowUserResizing = flexResizeColumns
        '设置合并

        .ExplorerBar = flexExSortShowAndMove  '设置允许点列头排序
        .MergeCells = flexMergeRestrictColumns
        .MergeCol(cnVehicle) = True
        .MergeCol(cnCompany) = True
        .MergeCol(cnDate) = True
        
        
        '显示列头的值
        
        
        .TextMatrix(0, cnCompany) = "公司"
        If m_szFormStatus = EFS_Vehicle Then
          .TextMatrix(0, cnDate) = "日期"
          .TextMatrix(0, cnVehicle) = "车辆"
        Else
          .TextMatrix(0, cnVehicle) = "车次"
          .TextMatrix(0, cnDate) = "发车时间"
        End If
        .TextMatrix(0, cnIsDec) = "状态"
        
        '用split函数分解出来的数组是从0开始的
        '提取出固定项的项目名称
        For i = 0 To nCols - 1
            For j = 1 To ArrayLength(m_atTemp)
                If Val(m_aszFixFeeItem(i)) = Val(m_atTemp(j).SplitItemID) Then
                    .TextMatrix(0, cnSplitItemStart + i) = m_atTemp(j).SplitItemName
                    Exit For
                
                End If
            Next j
        Next i
        
    End With
    With vsFixFee
        .ColWidth(0) = 100
        .ColWidth(cnVehicle) = 1500
        .ColWidth(cnCompany) = 1080
        .ColWidth(cnDate) = 1170
        
    End With

End Sub


Private Sub Form_Resize()
    AlignForm
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveHeadWidth Me.name, vsFixFee
    Unload Me
End Sub


Private Sub Timer1_Timer()
    Timer1.Enabled = False
    InitForm
End Sub

Private Sub txtVehicle_ButtonClick()
    On Error GoTo err
    Dim oShell As New STShell.CommDialog
    Dim aszTemp() As String
    oShell.Init g_oActiveUser
    
    Select Case m_szFormStatus
    Case FixFeeStatus.EFS_Vehicle
        aszTemp = oShell.SelectVehicleEX
        Set oShell = Nothing
        If ArrayLength(aszTemp) = 0 Then Exit Sub
        txtVehicle.Text = MakeDisplayString(Trim(aszTemp(1, 1)), Trim(aszTemp(1, 2)))
        Exit Sub
    Case FixFeeStatus.EFS_Bus
        aszTemp = oShell.SelectBus
        Set oShell = Nothing
        If ArrayLength(aszTemp) = 0 Then Exit Sub
        txtVehicle.Text = Trim(aszTemp(1, 1))
        Exit Sub
    End Select
    
err:
ShowErrorMsg
End Sub

Private Sub txtCompany_ButtonClick()
On Error GoTo ErrHandle
    Dim oShell As New STShell.CommDialog
    Dim aszTemp() As String
    oShell.Init g_oActiveUser
    aszTemp = oShell.SelectCompany
    Set oShell = Nothing
    If ArrayLength(aszTemp) = 0 Then Exit Sub
    txtCompany.Text = MakeDisplayString(Trim(aszTemp(1, 1)), Trim(aszTemp(1, 2)))

Exit Sub
ErrHandle:
ShowErrorMsg
End Sub



Private Sub abAction_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)

Select Case Tool.Caption
    Case "属性"
        EditObject
    Case "新增"
        AddObject
    Case "删除"
        DeleteObject
End Select
End Sub

Private Sub EditObject()
    '修改
    If vsFixFee.Rows <= 1 Then Exit Sub
    If vsFixFee.Row < 1 Then Exit Sub
    If m_szFormStatus = EFS_Vehicle Then
        frmVehicleFixFee.m_eFormStatus = ModifyStatus
        frmVehicleFixFee.m_szVehicleID = ResolveDisplay(vsFixFee.TextMatrix(vsFixFee.Row, cnVehicle))
        frmVehicleFixFee.m_dtDate = vsFixFee.TextMatrix(vsFixFee.Row, cnDate)
        frmVehicleFixFee.m_bIsParent = True
        frmVehicleFixFee.Show vbModal
    Else
        
        frmAddBusFixFee.m_eFormStatus = ModifyStatus
        frmAddBusFixFee.m_szBusID = ResolveDisplay(vsFixFee.TextMatrix(vsFixFee.Row, cnVehicle))
        frmAddBusFixFee.m_dtDate = vsFixFee.TextMatrix(vsFixFee.Row, cnDate)
        frmAddBusFixFee.m_bIsParent = True
        frmAddBusFixFee.Show vbModal
    End If
End Sub

Private Sub AddObject()
    '新增
    If m_szFormStatus = EFS_Vehicle Then
        frmVehicleFixFee.m_bIsParent = True
        frmVehicleFixFee.m_eFormStatus = AddStatus
        frmVehicleFixFee.Show vbModal
    Else
        frmAddBusFixFee.m_bIsParent = True
        frmAddBusFixFee.m_eFormStatus = AddStatus
        frmAddBusFixFee.Show vbModal
    End If
End Sub

Private Sub DeleteObject()
    On Error GoTo err
    Dim i As Integer
    Dim m_Answer As VbMsgBoxResult
    Dim oSplit As New Split
       
    Select Case m_szFormStatus
    Case FixFeeStatus.EFS_Vehicle
        With vsFixFee
        If .Rows <= 1 Then Exit Sub
        If .Row < 1 Then Exit Sub
            m_Answer = MsgBox("你是否确认删除此车辆的固定费用？", vbInformation + vbYesNo, Me.Caption)
            If m_Answer = vbYes Then
                oSplit.Init g_oActiveUser
                oSplit.DelVehicleFixFee ResolveDisplay(.TextMatrix(.Row, cnVehicle)), .TextMatrix(.Row, cnDate)
                .RemoveItem .Row
            
            End If
        End With
    
        Case FixFeeStatus.EFS_Bus
            With vsFixFee
            If .Rows <= 1 Then Exit Sub
            If .Row < 1 Then Exit Sub
                m_Answer = MsgBox("你是否确认删除此车次的固定费用？", vbInformation + vbYesNo, Me.Caption)
                If m_Answer = vbYes Then
                    oSplit.Init g_oActiveUser
                    oSplit.DelBusFixFee .TextMatrix(.Row, cnVehicle), .TextMatrix(.Row, cnDate)
                    .RemoveItem .Row
                End If
            End With
        End Select
            
    Exit Sub
err:
    ShowErrorMsg
End Sub

Private Sub QueryFixFee()
    '查询固定费
    Dim rsTemp As Recordset
    Dim i As Integer
    Dim j As Integer
    Dim nCols As Integer
    
    
    On Error GoTo ErrorHandle
    vsFixFee.Clear
    FillHead
    
    Select Case m_szFormStatus
    Case FixFeeStatus.EFS_Vehicle
        Set rsTemp = m_oReport.GetAllVehicleFixFee(txtVehicle.Text, ResolveDisplay(txtCompany.Text), dtpStartDate.Value, dtpEndDate.Value, cboLicenseTagNO.Text, ResolveDisplay(cboIsDec.Text))
        
        nCols = ArrayLength(m_aszFixFeeItem)
        vsFixFee.Rows = rsTemp.RecordCount + 1
        For i = 1 To rsTemp.RecordCount
            With vsFixFee
                .TextMatrix(i, cnVehicle) = MakeDisplayString(FormatDbValue(rsTemp!vehicle_id), FormatDbValue(rsTemp!license_tag_no))
                .TextMatrix(i, cnDate) = ToDBDate(FormatDbValue(rsTemp!bus_date))
                .TextMatrix(i, cnCompany) = FormatDbValue(rsTemp!transport_company_name)
                .TextMatrix(i, cnIsDec) = GetFixFeeStatusName(FormatDbValue(rsTemp!is_dec))
                For j = 0 To nCols - 1
                    .TextMatrix(i, cnSplitItemStart + j) = FormatDbValue(rsTemp.Fields("split_item_" & m_aszFixFeeItem(j)))
                Next j
            End With
            rsTemp.MoveNext
        Next i
        
    Case FixFeeStatus.EFS_Bus
        Set rsTemp = m_oReport.GetAllBusFixFee(ResolveDisplay(txtVehicle.Text), ResolveDisplay(txtCompany.Text), IIf(g_szIsFixFeeUpdateEachMonth = False, cszEmptyDateStr, dtpStartDate.Value), IIf(g_szIsFixFeeUpdateEachMonth = False, cszForeverDateStr, dtpEndDate.Value), , ResolveDisplay(cboIsDec.Text))
        
        nCols = ArrayLength(m_aszFixFeeItem)
        vsFixFee.Rows = rsTemp.RecordCount + 1
        For i = 1 To rsTemp.RecordCount
            With vsFixFee
                .TextMatrix(i, cnVehicle) = FormatDbValue(rsTemp!bus_id)
                .TextMatrix(i, cnDate) = ToDBDate(FormatDbValue(rsTemp!bus_date))
                .TextMatrix(i, cnCompany) = FormatDbValue(rsTemp!transport_company_name)
                .TextMatrix(i, cnIsDec) = GetFixFeeStatusName(FormatDbValue(rsTemp!is_dec))
                For j = 0 To nCols - 1
                    .TextMatrix(i, cnSplitItemStart + j) = FormatDbValue(rsTemp.Fields("split_item_" & m_aszFixFeeItem(j)))
                Next j
            End With
            rsTemp.MoveNext
        Next i
        
    End Select
    
    Exit Sub
ErrorHandle:
    ShowErrorMsg
    
End Sub

Public Sub AddList(pszVehicleID As String, Optional pdtDate As Date)
    '刷新新增的信息
    Dim rsTemp As Recordset
    Dim i As Integer
    Dim j As Integer
    Dim nCols As Integer
    On Error GoTo ErrorHandle
    
    m_oReport.Init g_oActiveUser
    If m_szFormStatus = EFS_Vehicle Then
        Set rsTemp = m_oReport.GetAllVehicleFixFee(ResolveDisplay(pszVehicleID), "", pdtDate, DateAdd("d", 1, pdtDate), "", -1)
        
        nCols = ArrayLength(m_aszFixFeeItem)
        vsFixFee.Rows = vsFixFee.Rows + rsTemp.RecordCount
        For i = 1 To rsTemp.RecordCount
            With vsFixFee
                .TextMatrix(vsFixFee.Rows - rsTemp.RecordCount + i - 1, cnVehicle) = MakeDisplayString(FormatDbValue(rsTemp!vehicle_id), FormatDbValue(rsTemp!license_tag_no))
                .TextMatrix(vsFixFee.Rows - rsTemp.RecordCount + i - 1, cnDate) = ToDBDate(FormatDbValue(rsTemp!bus_date))
                .TextMatrix(vsFixFee.Rows - rsTemp.RecordCount + i - 1, cnCompany) = FormatDbValue(rsTemp!transport_company_name)
                .TextMatrix(vsFixFee.Rows - rsTemp.RecordCount + i - 1, cnIsDec) = GetFixFeeStatusName(FormatDbValue(rsTemp!is_dec))
                For j = 0 To nCols - 1
                    .TextMatrix(vsFixFee.Rows - rsTemp.RecordCount + i - 1, cnSplitItemStart + j) = FormatDbValue(rsTemp.Fields("split_item_" & m_aszFixFeeItem(j)))
                Next j
            End With
            rsTemp.MoveNext
        Next i
    End If
    
    If m_szFormStatus = EFS_Bus Then
        Set rsTemp = m_oReport.GetAllBusFixFee(ResolveDisplay(pszVehicleID), "", IIf(g_szIsFixFeeUpdateEachMonth = False, cszEmptyDateStr, pdtDate), DateAdd("d", 1, pdtDate), "", -1)
        
        nCols = ArrayLength(m_aszFixFeeItem)
        vsFixFee.Rows = vsFixFee.Rows + rsTemp.RecordCount
        For i = 1 To rsTemp.RecordCount
            With vsFixFee
                .TextMatrix(vsFixFee.Rows - rsTemp.RecordCount + i - 1, cnVehicle) = FormatDbValue(rsTemp!bus_id)
                .TextMatrix(vsFixFee.Rows - rsTemp.RecordCount + i - 1, cnDate) = ToDBDate(FormatDbValue(rsTemp!bus_date))
                .TextMatrix(vsFixFee.Rows - rsTemp.RecordCount + i - 1, cnCompany) = FormatDbValue(rsTemp!transport_company_name)
                .TextMatrix(vsFixFee.Rows - rsTemp.RecordCount + i - 1, cnIsDec) = GetFixFeeStatusName(FormatDbValue(rsTemp!is_dec))
                For j = 0 To nCols - 1
                    .TextMatrix(vsFixFee.Rows - rsTemp.RecordCount + i - 1, cnSplitItemStart + j) = FormatDbValue(rsTemp.Fields("split_item_" & m_aszFixFeeItem(j)))
                Next j
            End With
            rsTemp.MoveNext
        Next i
    End If
    
    Exit Sub
ErrorHandle:
    ShowErrorMsg

    
End Sub

Public Sub UpdateList(pszVehicleID As String, pdtDate As Date)
    '刷新更新的信息
    Dim rsTemp As Recordset
    Dim i As Integer
    Dim j As Integer
    Dim nCols As Integer
    
    
    On Error GoTo ErrorHandle
    nCols = ArrayLength(m_aszFixFeeItem)
    
    If m_szFormStatus = EFS_Vehicle Then
        Set rsTemp = m_oReport.GetAllVehicleFixFee(ResolveDisplay(pszVehicleID), "", pdtDate, DateAdd("d", 1, pdtDate), "", -1)
        If rsTemp.RecordCount = 1 Then
            
    '        vsFixFee.Rows = vsFixFee.Rows + rsTemp.RecordCount
    '        For i = 1 To rsTemp.RecordCount
                With vsFixFee
                    .TextMatrix(vsFixFee.Row, cnVehicle) = MakeDisplayString(FormatDbValue(rsTemp!vehicle_id), FormatDbValue(rsTemp!license_tag_no))
                    .TextMatrix(vsFixFee.Row, cnDate) = ToDBDate(FormatDbValue(rsTemp!bus_date))
                    .TextMatrix(vsFixFee.Row, cnCompany) = FormatDbValue(rsTemp!transport_company_name)
                    .TextMatrix(vsFixFee.Row, cnIsDec) = GetFixFeeStatusName(FormatDbValue(rsTemp!is_dec))
                    For j = 0 To nCols - 1
                        .TextMatrix(vsFixFee.Row, cnSplitItemStart + j) = FormatDbValue(rsTemp.Fields("split_item_" & m_aszFixFeeItem(j)))
                    Next j
                End With
                rsTemp.MoveNext
    '        Next i
        End If
    End If
    
    If m_szFormStatus = EFS_Bus Then
        Set rsTemp = m_oReport.GetAllBusFixFee(ResolveDisplay(pszVehicleID), "", IIf(g_szIsFixFeeUpdateEachMonth = False, cszEmptyDateStr, pdtDate), DateAdd("d", 1, pdtDate), "", -1)
        If rsTemp.RecordCount = 1 Then
            
    '        vsFixFee.Rows = vsFixFee.Rows + rsTemp.RecordCount
    '        For i = 1 To rsTemp.RecordCount
                With vsFixFee
                    .TextMatrix(vsFixFee.Row, cnVehicle) = FormatDbValue(rsTemp!bus_id)
                    .TextMatrix(vsFixFee.Row, cnDate) = ToDBDate(FormatDbValue(rsTemp!bus_date))
                    .TextMatrix(vsFixFee.Row, cnCompany) = FormatDbValue(rsTemp!transport_company_name)
                    .TextMatrix(vsFixFee.Row, cnIsDec) = GetFixFeeStatusName(FormatDbValue(rsTemp!is_dec))
                    For j = 0 To nCols - 1
                        .TextMatrix(vsFixFee.Row, cnSplitItemStart + j) = FormatDbValue(rsTemp.Fields("split_item_" & m_aszFixFeeItem(j)))
                    Next j
                End With
                rsTemp.MoveNext
    '        Next i
        End If
    End If
    
    Exit Sub
ErrorHandle:
    ShowErrorMsg

    
    
End Sub







Private Sub vsFixFee_DblClick()
    
    EditObject
End Sub



Private Sub InitForm()


   Select Case m_szFormStatus
   
      Case FixFeeStatus.EFS_Vehicle
         Me.Caption = "车辆固定费用"
         dtpStartDate.Value = GetFirstMonthDay(Date)
         dtpEndDate.Value = GetLastMonthDay(Date)
         
      Case FixFeeStatus.EFS_Bus
         Me.Caption = "车次固定费用"
         lblVehicleID.Caption = "车次代码:"
         cboLicenseTagNO.Visible = False
         lblLicense.Visible = False
         
         If Not g_szIsFixFeeUpdateEachMonth Then
            dtpStartDate.Visible = False
            dtpEndDate.Visible = False
            lblStartDate.Visible = False
            lblEndDate.Visible = False
         Else
            dtpStartDate.Value = GetFirstMonthDay(Date)
            dtpEndDate.Value = GetLastMonthDay(Date)
         End If
         
    End Select
         
   
End Sub
