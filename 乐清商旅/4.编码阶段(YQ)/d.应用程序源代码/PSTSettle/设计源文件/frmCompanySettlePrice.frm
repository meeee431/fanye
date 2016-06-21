VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "actbar2.ocx"
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Begin VB.Form frmCompanySettlePrice 
   BackColor       =   &H00E0E0E0&
   Caption         =   "公司结算价"
   ClientHeight    =   6945
   ClientLeft      =   1050
   ClientTop       =   2970
   ClientWidth     =   12885
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6945
   ScaleWidth      =   12885
   WindowState     =   2  'Maximized
   Begin VSFlex7LCtl.VSFlexGrid vsSettlePrice 
      Height          =   3165
      Left            =   1380
      TabIndex        =   12
      Top             =   1455
      Width           =   5400
      _cx             =   9525
      _cy             =   5583
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
      FormatString    =   ""
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
      TabIndex        =   0
      Top             =   0
      Width           =   15135
      Begin FText.asFlatTextBox txtSellStation 
         Height          =   330
         Left            =   5670
         TabIndex        =   13
         Top             =   675
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   582
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
      Begin VB.ComboBox txtVehicleType 
         Height          =   300
         Left            =   8490
         TabIndex        =   11
         Top             =   390
         Width           =   1215
      End
      Begin FText.asFlatTextBox txtStation 
         Height          =   330
         Left            =   5670
         TabIndex        =   10
         Top             =   195
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   582
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
         Height          =   330
         Left            =   3210
         TabIndex        =   9
         Top             =   195
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   582
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
      Begin FText.asFlatTextBox txtRoute 
         Height          =   330
         Left            =   3210
         TabIndex        =   8
         Top             =   675
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   582
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
         Left            =   10170
         TabIndex        =   1
         Top             =   360
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
         MICON           =   "frmCompanySettlePrice.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblStation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "站点:"
         Height          =   180
         Left            =   5040
         TabIndex        =   7
         Top             =   270
         Width           =   450
      End
      Begin VB.Label lblSellStation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "上车站:"
         Height          =   180
         Left            =   5010
         TabIndex        =   6
         Top             =   750
         Width           =   630
      End
      Begin VB.Label lblRoute 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "线路:"
         Height          =   180
         Left            =   2310
         TabIndex        =   5
         Top             =   750
         Width           =   450
      End
      Begin VB.Label lblVehicleType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车型:"
         Height          =   180
         Left            =   7890
         TabIndex        =   4
         Top             =   480
         Width           =   450
      End
      Begin VB.Image Image2 
         Height          =   1275
         Left            =   60
         Picture         =   "frmCompanySettlePrice.frx":001C
         Top             =   150
         Width           =   2010
      End
      Begin VB.Label lblCompany 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "参运公司:"
         Height          =   240
         Left            =   2280
         TabIndex        =   2
         Top             =   240
         Width           =   810
      End
   End
   Begin ActiveBar2LibraryCtl.ActiveBar2 abAction 
      Height          =   4845
      Left            =   8160
      TabIndex        =   3
      Top             =   1470
      Width           =   1485
      _LayoutVersion  =   1
      _ExtentX        =   2619
      _ExtentY        =   8546
      _DataPath       =   ""
      Bands           =   "frmCompanySettlePrice.frx":14EF
   End
   Begin VB.Menu pmnu_Action 
      Caption         =   "操作"
      Visible         =   0   'False
      Begin VB.Menu pmnu_edit 
         Caption         =   "属性(&E)"
      End
      Begin VB.Menu pmnu_Add 
         Caption         =   "新增(&A)"
      End
      Begin VB.Menu pmnu_delete 
         Caption         =   "删除(&D)"
      End
      Begin VB.Menu pmnu_delete_route 
         Caption         =   "删除此线路"
      End
   End
End
Attribute VB_Name = "frmCompanySettlePrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const cnSellStationName = 1

Const cnCompanyName = 2
Const cnRouteName = 3
Const cnVehicleTypeName = 4
Const cnStationName = 5
Const cnMileage = 6
Const cnPassCharge = 7
Const cnSettleFullPrice = 8
Const cnSettleHalfPrice = 9
Const cnHalveFullPrice = 10
Const cnHalveHalfPrice = 11
Const cnServiceFullPrice = 12
Const cnServiceHalfPrice = 13
Const cnSpringFullPrice = 14
Const cnSpringHalfPrice = 15

Const cnRouteID = 16
Const cnSellStationID = 17
Const cnVehicleTypeCode = 18
Const cnStationID = 19
Const cnCompanyID = 20
Const cnAnnotation = 21

Const cnCols = 22



Private m_oReport As New Report
Private m_oCompanySettlePrice As New CompanySettlePrice



'界面排列
Private Sub AlignForm()
    On Error GoTo err
    ptShowInfo.Top = 0
    ptShowInfo.Left = 0
    ptShowInfo.Width = mdiMain.ScaleWidth
    
    vsSettlePrice.Top = ptShowInfo.Height + 50
    vsSettlePrice.Left = 50
    vsSettlePrice.Width = mdiMain.ScaleWidth - abAction.Width - 50
    vsSettlePrice.Height = mdiMain.ScaleHeight - ptShowInfo.Height - 50
    
    abAction.Top = vsSettlePrice.Top
    abAction.Left = vsSettlePrice.Width + 50
    abAction.Height = vsSettlePrice.Height
    Exit Sub
err:
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
    Case "删除此线路"
        DeleteRoute
    End Select
End Sub
Private Sub DeleteRoute()
    Dim m_Answer
    m_Answer = MsgBox("你是否确认删除此线路结算价", vbInformation + vbYesNo, Me.Caption)
    If m_Answer = vbYes Then
        m_oCompanySettlePrice.DeleteRoute vsSettlePrice.TextMatrix(vsSettlePrice.Row, cnCompanyID), vsSettlePrice.TextMatrix(vsSettlePrice.Row, cnRouteID)
    End If
    FillvsSettlePrice
End Sub
Private Sub EditObject()
    frmEditSettlePrice.szTitle = "公司结算价"
    
    frmEditSettlePrice.m_eFormStatus = ModifyStatus
    With vsSettlePrice
        frmEditSettlePrice.m_szCompany = MakeDisplayString(vsSettlePrice.TextMatrix(.Row, cnCompanyID), vsSettlePrice.TextMatrix(.Row, cnCompanyName))
        frmEditSettlePrice.m_szRoute = MakeDisplayString(vsSettlePrice.TextMatrix(.Row, cnRouteID), vsSettlePrice.TextMatrix(.Row, cnRouteName))
        frmEditSettlePrice.m_szVehicleType = MakeDisplayString(vsSettlePrice.TextMatrix(.Row, cnVehicleTypeCode), vsSettlePrice.TextMatrix(.Row, cnVehicleTypeName))
    End With
    frmEditSettlePrice.ZOrder 0
    frmEditSettlePrice.Show vbModal
End Sub

Private Sub AddObject()
    frmEditSettlePrice.szTitle = "公司结算价"
    frmEditSettlePrice.m_eFormStatus = AddStatus
    frmEditSettlePrice.ZOrder 0
    frmEditSettlePrice.Show vbModal
End Sub

Private Sub DeleteObject()
    On Error GoTo err
    Dim m_Answer, i As Integer
    m_Answer = MsgBox("你是否确认删除此结算价", vbInformation + vbYesNo, Me.Caption)
    If m_Answer = vbYes Then
        m_oCompanySettlePrice.Init g_oActiveUser
        For i = 1 To vsSettlePrice.Rows
            m_oCompanySettlePrice.CompanyID = vsSettlePrice.TextMatrix(vsSettlePrice.Row, cnCompanyID)
            m_oCompanySettlePrice.RouteID = vsSettlePrice.TextMatrix(vsSettlePrice.Row, cnRouteID)
            m_oCompanySettlePrice.StationID = vsSettlePrice.TextMatrix(vsSettlePrice.Row, cnStationID)
            m_oCompanySettlePrice.SellStationID = vsSettlePrice.TextMatrix(vsSettlePrice.Row, cnSellStationID)
            m_oCompanySettlePrice.VehicleTypeCode = vsSettlePrice.TextMatrix(vsSettlePrice.Row, cnVehicleTypeCode)
            m_oCompanySettlePrice.Delete
        Next i
        FillvsSettlePrice
    End If
    Exit Sub
err:
    ShowErrorMsg
End Sub

Private Sub cmdQuery_Click()
    FillvsSettlePrice
End Sub


Private Sub Form_Load()
    
    m_oReport.Init g_oActiveUser
    AlignForm
    FillHead
    FillvsSettlePrice
    FillBaseInfo
    
    AlignHeadWidth Me.name, vsSettlePrice
    
End Sub
Private Sub FillBaseInfo()
    On Error GoTo err
    Dim oBase As New BaseInfo
    Dim i As Integer
    Dim aszTemp() As String
    
'    Dim oSellStation As New SystemMan
'    Dim atTemp() As TDepartmentInfo
'    oSellStation.Init g_oActiveUser
'    atTemp = oSellStation.GetAllSellStation
'    txtSellStation.AddItem ""
'    For i = 1 To ArrayLength(atTemp)
'        txtSellStation.AddItem MakeDisplayString(atTemp(i).szSellStationID, atTemp(i).szSellStationName)
'    Next i
    
    oBase.Init g_oActiveUser
    txtVehicleType.AddItem ""
    aszTemp = oBase.GetAllVehicleModel
    For i = 1 To ArrayLength(aszTemp)
        txtVehicleType.AddItem MakeDisplayString(aszTemp(i, 1), aszTemp(i, 2))
    Next i
    Exit Sub
err:
    ShowErrorMsg
End Sub

Public Sub FillvsSettlePrice(Optional pszCompanyID As String, Optional pszRouteId As String, Optional pszVehicleType As String)
Attribute FillvsSettlePrice.VB_Description = "fule编写"
On Error GoTo here
    '填充列表
    Dim nCount As Integer
    Dim atTemp() As TCompanySettlePrice, atCompanySettlePrice() As TCompanySettlePrice
    Dim lvTemp As ListItem, i As Integer
    Dim szCompanyID As String
    Dim szRouteID As String
    Dim szVehicleType As String
    
    If pszCompanyID = "" Then
        szCompanyID = ResolveDisplay(txtCompany.Text)
    Else
        szCompanyID = pszCompanyID
    End If
    If pszRouteId = "" Then
        szRouteID = ResolveDisplay(txtRoute.Text)
    Else
        szRouteID = pszRouteId
    End If
    If pszVehicleType = "" Then
        szVehicleType = ResolveDisplay(txtVehicleType.Text)
    Else
        szVehicleType = pszVehicleType
    End If
    
    
    m_oReport.Init g_oActiveUser
'    VsSettlePrice.Clear
    atCompanySettlePrice = m_oReport.GetCompanySettlePriceLst(szCompanyID, szVehicleType, szRouteID)
    nCount = ArrayLength(atCompanySettlePrice)
    vsSettlePrice.Rows = nCount + 1
    If nCount <> 0 Then
        For i = 1 To ArrayLength(atCompanySettlePrice)
            
            vsSettlePrice.TextMatrix(i, cnCompanyName) = atCompanySettlePrice(i).CompanyName
            vsSettlePrice.TextMatrix(i, cnRouteName) = atCompanySettlePrice(i).RouteName
            vsSettlePrice.TextMatrix(i, cnSellStationName) = atCompanySettlePrice(i).SellStationName
            vsSettlePrice.TextMatrix(i, cnVehicleTypeName) = atCompanySettlePrice(i).VehicleTypeName
            vsSettlePrice.TextMatrix(i, cnStationName) = atCompanySettlePrice(i).StationName
            
            vsSettlePrice.TextMatrix(i, cnMileage) = atCompanySettlePrice(i).Mileage
            vsSettlePrice.TextMatrix(i, cnPassCharge) = atCompanySettlePrice(i).PassCharge
            vsSettlePrice.TextMatrix(i, cnSettleFullPrice) = atCompanySettlePrice(i).SettlefullPrice
            vsSettlePrice.TextMatrix(i, cnSettleHalfPrice) = atCompanySettlePrice(i).SettleHalfPrice
            vsSettlePrice.TextMatrix(i, cnHalveFullPrice) = atCompanySettlePrice(i).HalveFullPrice
            vsSettlePrice.TextMatrix(i, cnHalveHalfPrice) = atCompanySettlePrice(i).HalveHalfPrice
            vsSettlePrice.TextMatrix(i, cnServiceFullPrice) = atCompanySettlePrice(i).ServiceFullPrice
            vsSettlePrice.TextMatrix(i, cnServiceHalfPrice) = atCompanySettlePrice(i).ServiceHalfPrice
            vsSettlePrice.TextMatrix(i, cnSpringFullPrice) = atCompanySettlePrice(i).SpringFullPrice
            vsSettlePrice.TextMatrix(i, cnSpringHalfPrice) = atCompanySettlePrice(i).SpringHalfPrice
            
            vsSettlePrice.TextMatrix(i, cnRouteID) = atCompanySettlePrice(i).RouteID
            vsSettlePrice.TextMatrix(i, cnVehicleTypeCode) = atCompanySettlePrice(i).VehicleTypeCode
            vsSettlePrice.TextMatrix(i, cnSellStationID) = atCompanySettlePrice(i).SellStationID
            vsSettlePrice.TextMatrix(i, cnStationID) = atCompanySettlePrice(i).StationID
            vsSettlePrice.TextMatrix(i, cnCompanyID) = atCompanySettlePrice(i).CompanyID
            vsSettlePrice.TextMatrix(i, cnAnnotation) = atCompanySettlePrice(i).Annotation
        Next i
    End If
    SetNormal
'    VsSettlePrice.Refresh
'    If VsSettlePrice.ListItems.Count > 0 Then VsSettlePrice.ListItems(1).Selected = True
    WriteProcessBar False
    ShowSBInfo "共" & nCount & "个对象", ESB_ResultCountInfo
    ShowSBInfo ""

Exit Sub
here:
    ShowErrorMsg
End Sub

Private Sub FillHead()

    With vsSettlePrice
        .Cols = cnCols
        .Rows = 2
        .AllowUserResizing = flexResizeColumns
        '设置合并
        
        .MergeCells = flexMergeRestrictColumns
        .MergeCol(cnCompanyName) = True
        .MergeCol(cnRouteName) = True
        .MergeCol(cnSellStationName) = True
        .MergeCol(cnVehicleTypeName) = True
        .MergeCol(cnStationName) = True
        .MergeCol(cnMileage) = True
        
        .TextMatrix(0, cnSellStationName) = "上车站"
        .TextMatrix(0, cnCompanyName) = "公司"
        .TextMatrix(0, cnRouteName) = "线路"
        .TextMatrix(0, cnVehicleTypeName) = "车型"
        .TextMatrix(0, cnStationName) = "站点"
        .TextMatrix(0, cnMileage) = "里程"
        .TextMatrix(0, cnPassCharge) = "通行费"
        .TextMatrix(0, cnSettleFullPrice) = "结算全"
        .TextMatrix(0, cnSettleHalfPrice) = "结算半"
        .TextMatrix(0, cnHalveFullPrice) = "平分全"
        .TextMatrix(0, cnHalveHalfPrice) = "平分半"
        .TextMatrix(0, cnServiceFullPrice) = "劳务费全"
        .TextMatrix(0, cnServiceHalfPrice) = "劳务费半"
        .TextMatrix(0, cnSpringFullPrice) = "春运费全"
        .TextMatrix(0, cnSpringHalfPrice) = "春运费半"
        .TextMatrix(0, cnRouteID) = "线路代码"
        .TextMatrix(0, cnSellStationID) = "上车站代码"
        .TextMatrix(0, cnVehicleTypeCode) = "车型代码"
        .TextMatrix(0, cnStationID) = "站点代码"
        .TextMatrix(0, cnCompanyID) = "公司代码"
        .TextMatrix(0, cnAnnotation) = "备注"
    End With
    With vsSettlePrice
        .ColWidth(0) = 100
        .ColWidth(cnSellStationName) = 720
        .ColWidth(cnCompanyName) = 1080
        .ColWidth(cnRouteName) = 1170
        .ColWidth(cnVehicleTypeName) = 720
        .ColWidth(cnStationName) = 900
        .ColWidth(cnMileage) = 540
        .ColWidth(cnSettleFullPrice) = 720
        .ColWidth(cnSettleHalfPrice) = 720
        .ColWidth(cnHalveFullPrice) = 720
        .ColWidth(cnHalveHalfPrice) = 720
        .ColWidth(cnServiceFullPrice) = 720
        .ColWidth(cnServiceHalfPrice) = 720
        .ColWidth(cnSpringFullPrice) = 720
        .ColWidth(cnSpringHalfPrice) = 720
        .ColWidth(cnRouteID) = 0
        .ColWidth(cnSellStationID) = 0
        .ColWidth(cnVehicleTypeCode) = 0
        .ColWidth(cnStationID) = 0
        .ColWidth(cnCompanyID) = 0
        .ColWidth(cnAnnotation) = 0
    End With
    '1:720         2:1080        3:1170        4:720         5:900         6:540         7:720         8:720         9:720         10:720        11:720        12:900        13:900        14:0          15:0          16:0          17:0          18:0          19:0
    
End Sub

Private Sub Form_Resize()
    AlignForm
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveHeadWidth Me.name, vsSettlePrice
    Unload Me
End Sub

Private Sub vsSettlePrice_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    SortListView vsSettlePrice, ColumnHeader.Index
End Sub

Private Sub txtSellStation_ButtonClick()
On Error GoTo ErrHandle
    Dim oShell As New STShell.CommDialog
    Dim aszTemp() As String
    oShell.Init g_oActiveUser
    aszTemp = oShell.SelectStation
    Set oShell = Nothing
    If ArrayLength(aszTemp) = 0 Then Exit Sub
    txtSellStation.Text = MakeDisplayString(Trim(aszTemp(1, 1)), Trim(aszTemp(1, 2)))

Exit Sub
ErrHandle:
ShowErrorMsg
End Sub

Private Sub vsSettlePrice_DblClick()
    EditObject
End Sub

Private Sub vsSettlePrice_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete And vsSettlePrice.ListItems.Count > 0 Then
        DeleteObject
    End If
End Sub

Private Sub vsSettlePrice_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyDelete Then
        DeleteObject
    End If
End Sub

Private Sub vsSettlePrice_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu pmnu_Action
    End If
End Sub

Private Sub pmnu_Add_Click()
AddObject
End Sub

Private Sub pmnu_delete_Click()
DeleteObject
End Sub

Private Sub pmnu_delete_route_Click()
    DeleteRoute
End Sub

Private Sub pmnu_edit_Click()
    EditObject
End Sub

Private Sub txtCompany_ButtonClick()
On Error GoTo ErrHandle
    Dim oShell As New STShell.CommDialog
    Dim aszTemp() As String
    oShell.Init g_oActiveUser
    aszTemp = oShell.SelectCompany()
    Set oShell = Nothing
    If ArrayLength(aszTemp) = 0 Then Exit Sub
    txtCompany.Text = MakeDisplayString(Trim(aszTemp(1, 1)), Trim(aszTemp(1, 2)))

Exit Sub
ErrHandle:
ShowErrorMsg
End Sub

Private Sub txtRoute_ButtonClick()
On Error GoTo ErrHandle
    Dim oShell As New STShell.CommDialog
    Dim aszTemp() As String
    oShell.Init g_oActiveUser
    aszTemp = oShell.SelectRoute
    Set oShell = Nothing
    If ArrayLength(aszTemp) = 0 Then Exit Sub
    txtRoute.Text = MakeDisplayString(Trim(aszTemp(1, 1)), Trim(aszTemp(1, 2)))

Exit Sub
ErrHandle:
ShowErrorMsg
End Sub



Private Sub txtStation_ButtonClick()
On Error GoTo ErrHandle
    Dim oShell As New STShell.CommDialog
    Dim aszTemp() As String
    oShell.Init g_oActiveUser
    aszTemp = oShell.SelectStation
    Set oShell = Nothing
    If ArrayLength(aszTemp) = 0 Then Exit Sub
    txtStation.Text = MakeDisplayString(Trim(aszTemp(1, 1)), Trim(aszTemp(1, 2)))

Exit Sub
ErrHandle:
ShowErrorMsg
End Sub

