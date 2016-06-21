VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Begin VB.Form frmInsertStation 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "插入站点"
   ClientHeight    =   3180
   ClientLeft      =   4530
   ClientTop       =   4815
   ClientWidth     =   5280
   Icon            =   "frmInsertStation.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraPrice 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "站点票价"
      Height          =   1695
      Left            =   0
      TabIndex        =   13
      Top             =   1470
      Width           =   5745
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mfgStation 
         Height          =   1560
         Left            =   150
         TabIndex        =   8
         Top             =   60
         Width           =   4980
         _ExtentX        =   8784
         _ExtentY        =   2752
         _Version        =   393216
         FixedCols       =   0
         BackColorFixed  =   14737632
         BackColorBkg    =   14737632
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.TextBox txtMileage 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1260
      TabIndex        =   7
      Top             =   990
      Width           =   1275
   End
   Begin VB.ComboBox cobStationId 
      Height          =   300
      Left            =   1260
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   570
      Width           =   1290
   End
   Begin VB.OptionButton optNext 
      BackColor       =   &H00E0E0E0&
      Caption         =   "后"
      Height          =   195
      Left            =   2670
      TabIndex        =   4
      Top             =   630
      Value           =   -1  'True
      Width           =   450
   End
   Begin VB.OptionButton optPrev 
      BackColor       =   &H00E0E0E0&
      Caption         =   "前"
      Height          =   195
      Left            =   3240
      TabIndex        =   5
      Top             =   630
      Width           =   480
   End
   Begin RTComctl3.CoolButton CancelButton 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   3960
      TabIndex        =   10
      Top             =   570
      Width           =   1185
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
      MICON           =   "frmInsertStation.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton OKButton 
      Default         =   -1  'True
      Height          =   315
      Left            =   3960
      TabIndex        =   9
      Top             =   150
      Width           =   1185
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "确定(&O)"
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
      MICON           =   "frmInsertStation.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin FText.asFlatTextBox txtBusStationID 
      Height          =   300
      Left            =   1260
      TabIndex        =   1
      Top             =   150
      Width           =   2445
      _ExtentX        =   4313
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
   Begin RTComctl3.CoolButton cmdPrice 
      Height          =   315
      Left            =   3960
      TabIndex        =   11
      Top             =   990
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   ""
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
      MICON           =   "frmInsertStation.frx":0044
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   -1  'True
      VALUE           =   0   'False
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "公里"
      Height          =   180
      Left            =   2670
      TabIndex        =   12
      Top             =   1050
      Width           =   360
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "到站里程(&R):"
      Height          =   180
      Left            =   150
      TabIndex        =   6
      Top             =   1050
      Width           =   1080
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "插入位置(&P):"
      Height          =   180
      Left            =   150
      TabIndex        =   2
      Top             =   630
      Width           =   1080
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "插入站点(&N):"
      Height          =   180
      Left            =   150
      TabIndex        =   0
      Top             =   210
      Width           =   1080
   End
End
Attribute VB_Name = "frmInsertStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'Option Explicit
'Private m_oReBus As REBus
'Private mtTicketTypes() As TTicketType
'Dim matReTicketPrice() As TRETicketPriceEx
'
'Public Function Init(poREBus As REBus)
'    Set m_oReBus = poREBus
'End Function
'
'Private Sub CancelButton_Click()
'    Unload Me
'End Sub
'Private Sub EnableOKButton()
'    If ArrayLength(matReTicketPrice) = 0 Or ResolveDisplay(txtBusStationID.Text) = "" Or cobStationId.Text = "" Or Val(txtMileage.Text) <= 0 Then
'        OKButton.Enabled = False
'    Else
'        OKButton.Enabled = True
'    End If
'End Sub
'Private Sub cmdPrice_Click()
'    If Not cmdPrice.Value Then
'        cmdPrice.Caption = "站点票价>>"
'        Me.Height = Me.Height - fraPrice.Height
'        fraPrice.Visible = False
'        Exit Sub
'    Else
'        cmdPrice.Caption = "站点票价<<"
'        Me.Height = Me.Height + fraPrice.Height
'        fraPrice.Visible = True
'    End If
'End Sub
'
'Private Sub cobStationId_Change()
'    EnableOKButton
'End Sub
'
'Private Sub cobStationId_Click()
'    EnableOKButton
'End Sub
'
'Private Sub Form_Load()
'On Error GoTo ErrHandle
'    AlignFormPos Me
'
'    '添加站点
'    Dim i As Integer
'    Dim nCountSeatType As Integer
'    nCountSeatType = frmEnvBusRoute.hfgRouteStation.Rows
'    For i = 1 To nCountSeatType - 1
'      cobStationId.AddItem frmEnvBusRoute.hfgRouteStation.TextArray(i * frmEnvBusRoute.hfgRouteStation.Cols + 1)
'    Next
'
'    mfgStation.Cols = 3
'    mfgStation.TextArray(0) = "座位类型"
'    mfgStation.TextArray(1) = "票  种"
'    mfgStation.TextArray(2) = "票  价"
'
'    Dim oParSystem As New SystemParam
'    oParSystem.Init g_oActiveUser
'    mtTicketTypes = oParSystem.GetAllTicketType(1, False)
'
'
'    cmdPrice_Click
'    EnableOKButton
'    Exit Sub
'ErrHandle:
'    ShowErrorMsg
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'    SaveFormPos Me
'End Sub
'
'Private Sub OKButton_Click()
'
'
'    On Error GoTo ErrHandle
'    Dim nCountRow As Integer
'    Dim nSerial As Integer
'    Dim szMsg As String
'    nCountRow = cobStationId.ListCount
'    '取得站点序号
'    If optNext.Value = True Then
'        nSerial = cobStationId.ListIndex + 2
'        szMsg = szMsg & "将站点插入到" & cobStationId.Text & "站点后"
'    Else
'        nSerial = cobStationId.ListIndex + 1
'        szMsg = szMsg & "将站点插入到" & cobStationId.Text & "站点前"
'    End If
'    Dim bFlgEndStation As Boolean
'
'    '是否为终点站
'    If cobStationId.ListCount < nSerial Then
'     bFlgEndStation = True
'    End If
'
'    If txtBusStationID.Text = "" Or Val(txtMileage.Text) = 0 Or cobStationId.Text = "" Then
'        MsgBox "请输入要插入的站点和里程!", vbExclamation, "错误"
'        Exit Sub
'    End If
'
'    szMsg = szMsg & ",此操作将影响检票和售票!" & Chr(10) & "是否真的进行处理？"
'    Dim nResult As Integer
'    nResult = MsgBox(szMsg, vbQuestion + vbYesNoCancel, "车次站点--插入站点")
'    If nResult = vbCancel Then Exit Sub
'    If nResult = vbNo Then Unload Me
'
'    SetBusy
'
'    '形成传入接口的数据格式
'    Dim atInsTkPrice() As TRETicketPriceEx
'    ReDim atInsTkPrice(1 To ArrayLength(matReTicketPrice) - 1)
'    Dim i As Integer
'    For i = 1 To ArrayLength(matReTicketPrice) - 1
'        atInsTkPrice(i).sgMileage = Val(txtMileage.Text)
'        atInsTkPrice(i).szSeatType = matReTicketPrice(i).szSeatType
'        atInsTkPrice(i).sgTotal = matReTicketPrice(i).sgTotal
'        atInsTkPrice(i).sgBase = matReTicketPrice(i).sgBase
'        atInsTkPrice(i).szStationID = matReTicketPrice(i).szStationID
'        atInsTkPrice(i).nTicketType = matReTicketPrice(i).nTicketType
'    Next i
'
'    m_oReBus.EnInsertStation nSerial, atInsTkPrice, bFlgEndStation     '添加站点及对应票价
'
'    '在站点设置表单中添加新站点
'    frmEnvBusRoute.FillStation
'
'    SetNormal
'    MsgBox " 环境插入站点及缺省票价成功!", vbInformation, "信息"
'    Unload Me
'
'    Exit Sub
'ErrHandle:
'    SetNormal
'    ShowErrorMsg
''    MsgBox " 环境插入站点失败" & Chr(10) & err.Description, vbInformation, "车次站点--插入站点"
'End Sub
'
'Private Sub txtBusStationID_ButtonClick()
''On Error GoTo ErrHandle
''    Dim oShell As New CommDialog
''    Dim aszTmp() As String
''    oShell.Init g_oActiveUser
''    aszTmp = oShell.SelectStation
''    Set oShell = Nothing
''    If ArrayLength(aszTmp) = 0 Then Exit Sub
''    txtBusStationID.Text = MakeDisplayString(aszTmp(1, 1), aszTmp(1, 2))
''    FillBusPriceTable
''
''    Exit Sub
''ErrHandle:
''    ShowErrorMsg
'End Sub
''填充票价
'Private Sub FillBusPriceTable()
'    Dim szStationID As String
'    szStationID = ResolveDisplay(txtBusStationID.Text)
'    If ArrayLength(matReTicketPrice) > 0 Then      '判断是否已经填充
'        If matReTicketPrice(1).szStationID = szStationID Then
'            Exit Sub
'        End If
'    End If
'
'    Dim aszSeatType() As String
'    aszSeatType = m_oReBus.GetReBusSeatType       '得到座位
'    matReTicketPrice = m_oReBus.GetStationPrice(szStationID, aszSeatType)       '得到初始票价
'
'    Dim i As Integer
'    Dim nCount As Integer
'    Dim j As Integer
'    nCount = ArrayLength(matReTicketPrice) - 1
'    If nCount = -1 Then Exit Sub
'    Dim bTreFlg As Boolean
'    bTreFlg = True
'    txtMileage.Text = matReTicketPrice(1).sgMileage
'     With mfgStation
'        .Redraw = True
'        .FixedCols = 0
'        .Rows = nCount + 1
'        .MergeCol(0) = True
'        For i = 1 To nCount
'
'           .MergeRow(i) = True
'
'           .TextArray(i * .Cols + 0) = MakeDisplayString(Trim(matReTicketPrice(i).szSeatType), Trim(matReTicketPrice(i).szSeatTypeName))
'           .TextArray(i * .Cols + 1) = FindTicketTypeName(matReTicketPrice(i).nTicketType)
'           .TextArray(i * .Cols + 2) = Round(matReTicketPrice(i).sgTotal, 1)
'           .MergeCells = flexMergeRestrictColumns
'        Next
'        .FixedCols = 1
'        .Redraw = True
'    End With
'End Sub
''取得票种名称
'Private Function FindTicketTypeName(szType As Integer) As String
'Dim i As Integer
'Dim nCount As Integer
'   i = 1
'   nCount = ArrayLength(mtTicketTypes)
'   Do While Not Trim(mtTicketTypes(i).nTicketTypeID) = Trim(szType)
'   i = i + 1
'   Loop
'   FindTicketTypeName = mtTicketTypes(i).szTicketTypeName
'End Function
'
'Private Sub txtBusStationID_Change()
'    EnableOKButton
'End Sub
'
''Private Sub txtBusStationID_KeyPress(KeyAscii As Integer)
''On Error GoTo ErrHandle
''    If KeyAscii = vbEnter Then
''        FillBusPriceTable
''    End If
''    Exit Sub
''ErrHandle:
''    ShowErrorMsg
''End Sub
'
'Private Sub txtBusStationID_Validate(Cancel As Boolean)
'    Dim szStationID As String
'    szStationID = ResolveDisplay(txtBusStationID.Text)
'    If szStationID = "" Then Exit Sub
'    If IfStationExist(szStationID) Then
'       MsgBox "本线路已存在该站点!", vbExclamation, "错误"
'       Cancel = True
'       Exit Sub
'    End If
'    FillBusPriceTable
'End Sub
''判断是否已存在该站点
'Private Function IfStationExist(szStaionId As String) As Boolean
'   Dim nRows As Integer
'   Dim i As Integer
'   Dim nCols As Integer
'   Dim sgData As String
'   nRows = frmEnvBusRoute.hfgRouteStation.Rows
'   nCols = frmEnvBusRoute.hfgRouteStation.Cols
'   i = 0
'   Do While Trim(frmEnvBusRoute.hfgRouteStation.TextArray(i * nCols + 0)) <> szStaionId
'    i = i + 1
'    If i >= nRows Then
'       IfStationExist = False
'       Exit Function
'    End If
'   Loop
'   IfStationExist = True
'End Function
'
'Private Sub txtMileage_Change()
'    FormatTextToNumeric txtMileage, False
'    EnableOKButton
'End Sub
'
