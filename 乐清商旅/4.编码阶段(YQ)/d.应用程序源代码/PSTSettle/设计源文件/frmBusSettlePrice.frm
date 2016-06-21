VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "actbar2.ocx"
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Begin VB.Form frmBusSettlePrice 
   Caption         =   "车次结算价"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11445
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8220
   ScaleWidth      =   11445
   WindowState     =   2  'Maximized
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
      Begin FText.asFlatTextBox txtSellStation 
         Height          =   330
         Left            =   7665
         TabIndex        =   2
         Top             =   660
         Width           =   1875
         _ExtentX        =   3307
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
      End
      Begin FText.asFlatTextBox txtBus 
         Height          =   345
         Left            =   2865
         TabIndex        =   3
         Top             =   203
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   609
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
      End
      Begin FText.asFlatTextBox txtStation 
         Height          =   330
         Left            =   7665
         TabIndex        =   4
         Top             =   210
         Width           =   1875
         _ExtentX        =   3307
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
      End
      Begin FText.asFlatTextBox txtCompany 
         Height          =   330
         Left            =   5325
         TabIndex        =   5
         Top             =   210
         Width           =   1575
         _ExtentX        =   2778
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
      End
      Begin RTComctl3.CoolButton cmdQuery 
         Default         =   -1  'True
         Height          =   375
         Left            =   10170
         TabIndex        =   6
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
         MICON           =   "frmBusSettlePrice.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblBusID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车次代码:"
         Height          =   180
         Left            =   1875
         TabIndex        =   7
         Top             =   285
         Width           =   840
      End
      Begin VB.Label lblStation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "站点:"
         Height          =   180
         Left            =   7020
         TabIndex        =   10
         Top             =   285
         Width           =   450
      End
      Begin VB.Label lblSellStation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "上车站:"
         Height          =   180
         Left            =   7020
         TabIndex        =   9
         Top             =   735
         Width           =   630
      End
      Begin VB.Image Image2 
         Height          =   1275
         Left            =   60
         Picture         =   "frmBusSettlePrice.frx":001C
         Top             =   150
         Width           =   2010
      End
      Begin VB.Label lblCompany 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "参运公司:"
         Height          =   180
         Left            =   4425
         TabIndex        =   8
         Top             =   285
         Width           =   810
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid vsSettlePrice 
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
   Begin ActiveBar2LibraryCtl.ActiveBar2 abAction 
      Height          =   4845
      Left            =   8160
      TabIndex        =   11
      Top             =   1470
      Width           =   1590
      _LayoutVersion  =   1
      _ExtentX        =   2805
      _ExtentY        =   8546
      _DataPath       =   ""
      Bands           =   "frmBusSettlePrice.frx":14EF
   End
End
Attribute VB_Name = "frmBusSettlePrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const cnSellStationName = 1
Const cnRouteName = 2
Const cnBusID = 3
Const cnTransportCompany = 4
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
Const cnSellStationID = 16
Const cnStationID = 17
Const cnAnnotation = 18

Const cnCols = 19

Private m_oRepot As New Report
Private m_oBusSettlePrice As New BusSettlePrice

'界面排列
Private Sub AlignForm()

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
            DetletRoute
    End Select
End Sub
Private Sub DetletRoute()
    Dim m_Answer
    m_Answer = MsgBox("你是否确认删除此线路结算价", vbInformation + vbYesNo, Me.Caption)
    If m_Answer = vbYes Then
        m_oBusSettlePrice.DeleteRoute vsSettlePrice.TextMatrix(vsSettlePrice.Row, cnBusID), ResolveDisplay(vsSettlePrice.TextMatrix(vsSettlePrice.Row, cnTransportCompany))
    End If
    QueryBusSettlePrice
End Sub
Private Sub EditObject()
    frmEditSettlePrice.szTitle = Me.Caption

    frmEditSettlePrice.m_eFormStatus = ModifyStatus
    frmEditSettlePrice.m_szBus = vsSettlePrice.TextMatrix(vsSettlePrice.Row, cnBusID)
    frmEditSettlePrice.m_szRoute = vsSettlePrice.TextMatrix(vsSettlePrice.Row, cnRouteName)
    frmEditSettlePrice.m_szTransportCompany = vsSettlePrice.TextMatrix(vsSettlePrice.Row, cnTransportCompany)
    frmEditSettlePrice.ZOrder 0
    frmEditSettlePrice.Show vbModal
End Sub

Private Sub AddObject()
    frmEditSettlePrice.szTitle = Me.Caption
    frmEditSettlePrice.m_eFormStatus = AddStatus
    frmEditSettlePrice.ZOrder 0
    frmEditSettlePrice.Show vbModal
End Sub

Private Sub DeleteObject()
    On Error GoTo err
    Dim i As Integer
    Dim m_Answer
    m_Answer = MsgBox("你是否确认删除此结算价", vbInformation + vbYesNo, Me.Caption)
    If m_Answer = vbYes Then
        m_oBusSettlePrice.Init g_oActiveUser
        For i = 1 To vsSettlePrice.Rows
            m_oBusSettlePrice.BusID = vsSettlePrice.TextMatrix(vsSettlePrice.Row, cnBusID)
            m_oBusSettlePrice.TransportCompanyID = ResolveDisplay(vsSettlePrice.TextMatrix(vsSettlePrice.Row, cnTransportCompany))
            m_oBusSettlePrice.StationID = vsSettlePrice.TextMatrix(vsSettlePrice.Row, cnStationID)
            m_oBusSettlePrice.SellStationID = vsSettlePrice.TextMatrix(vsSettlePrice.Row, cnSellStationID)
            m_oBusSettlePrice.Delete
        Next i
        QueryBusSettlePrice

    End If
    Exit Sub
err:
ShowErrorMsg
End Sub

Private Sub cmdQuery_Click()
    QueryBusSettlePrice
End Sub
Public Sub QueryBusSettlePrice(Optional pszBusID As String, Optional pszTransportCompanyID As String)
    On Error GoTo err
    Dim atBusSettlePrice() As TBusSettlePrice
    Dim lvTemp As ListItem
    Dim i As Integer, nCount As Integer

    Dim szBusID As String
    Dim szTransportCompanyID As String


    If pszBusID = "" Then
        szBusID = ResolveDisplay(txtBus.Text)
    Else
        szBusID = pszBusID
    End If
    If pszTransportCompanyID = "" Then
        szTransportCompanyID = ResolveDisplay(txtCompany.Text)
    Else
        szTransportCompanyID = pszTransportCompanyID
    End If

    m_oRepot.Init g_oActiveUser

    atBusSettlePrice = m_oRepot.GetBusSettlePriceLst(szBusID, ResolveDisplay(txtCompany.Text), ResolveDisplay(txtSellStation.Text), ResolveDisplay(txtStation.Text))
    nCount = ArrayLength(atBusSettlePrice)
    vsSettlePrice.Rows = nCount + 1
    If ArrayLength(atBusSettlePrice) <> 0 Then
        For i = 1 To ArrayLength(atBusSettlePrice)
            vsSettlePrice.TextMatrix(i, cnBusID) = atBusSettlePrice(i).BusID
            vsSettlePrice.TextMatrix(i, cnTransportCompany) = MakeDisplayString(atBusSettlePrice(i).TransportCompanyID, atBusSettlePrice(i).TransportCompanyName)
            vsSettlePrice.TextMatrix(i, cnRouteName) = MakeDisplayString(atBusSettlePrice(i).RouteID, atBusSettlePrice(i).RouteName)
            vsSettlePrice.TextMatrix(i, cnSellStationID) = atBusSettlePrice(i).SellStationID
            vsSettlePrice.TextMatrix(i, cnSellStationName) = atBusSettlePrice(i).SellStationName
            vsSettlePrice.TextMatrix(i, cnStationID) = atBusSettlePrice(i).StationID
            vsSettlePrice.TextMatrix(i, cnStationName) = atBusSettlePrice(i).StationName

            vsSettlePrice.TextMatrix(i, cnMileage) = atBusSettlePrice(i).Mileage
            vsSettlePrice.TextMatrix(i, cnPassCharge) = atBusSettlePrice(i).PassCharge
            vsSettlePrice.TextMatrix(i, cnSettleFullPrice) = atBusSettlePrice(i).SettlefullPrice
            vsSettlePrice.TextMatrix(i, cnSettleHalfPrice) = atBusSettlePrice(i).SettleHalfPrice
            vsSettlePrice.TextMatrix(i, cnHalveFullPrice) = atBusSettlePrice(i).HalveFullPrice
            vsSettlePrice.TextMatrix(i, cnHalveHalfPrice) = atBusSettlePrice(i).HalveHalfPrice
            vsSettlePrice.TextMatrix(i, cnServiceFullPrice) = atBusSettlePrice(i).ServiceFullPrice
            vsSettlePrice.TextMatrix(i, cnServiceHalfPrice) = atBusSettlePrice(i).ServiceHalfPrice
            vsSettlePrice.TextMatrix(i, cnSpringFullPrice) = atBusSettlePrice(i).SpringFullPrice
            vsSettlePrice.TextMatrix(i, cnSpringHalfPrice) = atBusSettlePrice(i).SpringHalfPrice
            vsSettlePrice.TextMatrix(i, cnAnnotation) = atBusSettlePrice(i).Annotation

        Next i
    End If
    SetNormal
    WriteProcessBar False
    ShowSBInfo "共" & nCount & "个对象", ESB_ResultCountInfo
    ShowSBInfo ""

    Exit Sub
err:
ShowErrorMsg
End Sub



Private Sub Form_Load()
    m_oRepot.Init g_oActiveUser
    AlignForm
    FillHead
    QueryBusSettlePrice
    AlignHeadWidth Me.name, vsSettlePrice
End Sub

Private Sub FillHead()
    Dim oSellStation As New SystemMan
    Dim i As Integer
    Dim atTemp() As TDepartmentInfo


    With vsSettlePrice
        .Cols = cnCols
        .Rows = 2
        .AllowUserResizing = flexResizeColumns
        '设置合并

        .MergeCells = flexMergeRestrictColumns
        .MergeCol(cnTransportCompany) = True
        .MergeCol(cnBusID) = True
        .MergeCol(cnRouteName) = True
        .MergeCol(cnSellStationName) = True
        .MergeCol(cnStationName) = True
        .MergeCol(cnMileage) = True

        .TextMatrix(0, cnSellStationName) = "上车站"
        .TextMatrix(0, cnRouteName) = "线路"
        .TextMatrix(0, cnBusID) = "车次代码"
        .TextMatrix(0, cnTransportCompany) = "参运公司"
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
        .TextMatrix(0, cnSellStationID) = "上车站代码"
        .TextMatrix(0, cnStationID) = "站点代码"
        .TextMatrix(0, cnAnnotation) = "计算说明"
    End With
    With vsSettlePrice
        .ColWidth(0) = 100
        .ColWidth(cnSellStationName) = 720
        .ColWidth(cnBusID) = 700
        .ColWidth(cnTransportCompany) = 1600
        .ColWidth(cnRouteName) = 1170
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
        .ColWidth(cnSellStationID) = 0
        .ColWidth(cnStationID) = 0
        .ColWidth(cnAnnotation) = 0
    End With

End Sub



Private Sub Form_Resize()
    AlignForm
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveHeadWidth Me.name, vsSettlePrice
    Unload Me
End Sub

Private Sub txtSellStation_ButtonClick()
    On Error GoTo err
    Dim oShell As New STShell.CommDialog
    Dim aszTemp() As String
    oShell.Init g_oActiveUser
    aszTemp = oShell.SelectStation
    Set oShell = Nothing
    If ArrayLength(aszTemp) = 0 Then Exit Sub
    txtSellStation.Text = MakeDisplayString(Trim(aszTemp(1, 1)), Trim(aszTemp(1, 2)))
    Exit Sub
err:
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

'Private Sub vsSettlePrice_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
'    If Button = vbRightButton Then
'        PopupMenu pmnu_Action
'    End If
'End Sub

Private Sub pmnu_Add_Click()
AddObject
End Sub

Private Sub pmnu_delete_Click()
DeleteObject
End Sub

Private Sub pmnu_delete_route_Click()
    DetletRoute
End Sub

Private Sub pmnu_edit_Click()
EditObject
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

Private Sub txtStation_ButtonClick()
    On Error GoTo err
    Dim oShell As New STShell.CommDialog
    Dim aszTemp() As String
    oShell.Init g_oActiveUser
    aszTemp = oShell.SelectStation
    Set oShell = Nothing
    If ArrayLength(aszTemp) = 0 Then Exit Sub
    txtStation.Text = MakeDisplayString(Trim(aszTemp(1, 1)), Trim(aszTemp(1, 2)))
    Exit Sub
err:
ShowErrorMsg
End Sub

Private Sub txtBus_ButtonClick()
    On Error GoTo err
    Dim oShell As New STShell.CommDialog
    Dim aszTemp() As String
        oShell.Init g_oActiveUser
        aszTemp = oShell.SelectBus
        Set oShell = Nothing
        If ArrayLength(aszTemp) = 0 Then Exit Sub
        txtBus.Text = Trim(aszTemp(1, 1))
    Exit Sub
err:
ShowErrorMsg
End Sub


