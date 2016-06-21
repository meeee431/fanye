VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Begin VB.Form frmInsertReStation 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "环境站点插入"
   ClientHeight    =   6000
   ClientLeft      =   3585
   ClientTop       =   2895
   ClientWidth     =   8760
   Icon            =   "frmInsertReStation.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   8760
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAband 
      Caption         =   "放弃(&A)"
      Height          =   330
      Left            =   7245
      TabIndex        =   15
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox txtInfo 
      Height          =   225
      Left            =   1800
      TabIndex        =   0
      ToolTipText     =   "您可以输入数值，然后按下〈计算〉按钮修改该站点票价。"
      Top             =   2280
      Visible         =   0   'False
      Width           =   945
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mfgStation 
      Height          =   3660
      Left            =   315
      TabIndex        =   14
      Top             =   1710
      Width           =   7950
      _ExtentX        =   14023
      _ExtentY        =   6456
      _Version        =   393216
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdCaula 
      Caption         =   "计 算(&C)"
      Height          =   330
      Left            =   7245
      TabIndex        =   13
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "保存(&S)"
      Enabled         =   0   'False
      Height          =   330
      Left            =   7245
      TabIndex        =   8
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "退出(&X)"
      Height          =   285
      Left            =   7245
      TabIndex        =   7
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   1290
      Left            =   360
      TabIndex        =   1
      Top             =   45
      Width           =   6750
      Begin VB.PictureBox txtBusStationIDold 
         Height          =   315
         Left            =   4950
         ScaleHeight     =   255
         ScaleWidth      =   1515
         TabIndex        =   18
         Top             =   750
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.OptionButton Option2 
         Caption         =   "前"
         Height          =   195
         Left            =   4185
         TabIndex        =   17
         Top             =   825
         Width           =   510
      End
      Begin VB.OptionButton op1 
         Caption         =   "后"
         Height          =   285
         Left            =   3420
         TabIndex        =   16
         Top             =   780
         Value           =   -1  'True
         Width           =   690
      End
      Begin VB.ComboBox cobStationId 
         Height          =   300
         Left            =   4500
         TabIndex        =   3
         Top             =   240
         Width           =   1275
      End
      Begin VB.TextBox txtMileage 
         Height          =   315
         Left            =   945
         TabIndex        =   2
         Top             =   765
         Width           =   1575
      End
      Begin FText.asFlatTextBox txtBusStationID 
         Height          =   300
         Left            =   945
         TabIndex        =   19
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
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
      End
      Begin VB.Label Label1 
         Caption         =   "新增站点"
         Height          =   195
         Left            =   90
         TabIndex        =   6
         Top             =   300
         Width           =   795
      End
      Begin VB.Label Label4 
         Caption         =   "插入到站点"
         Height          =   210
         Left            =   3420
         TabIndex        =   5
         Top             =   292
         Width           =   945
      End
      Begin VB.Label Label3 
         Caption         =   "到站里程"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   795
         Width           =   735
      End
   End
   Begin VB.Timer TimStar 
      Interval        =   500
      Left            =   1440
      Top             =   2160
   End
   Begin VB.Label Label2 
      Caption         =   "环境票价项"
      Height          =   225
      Left            =   360
      TabIndex        =   12
      Top             =   1485
      Width           =   945
   End
   Begin VB.Label lblStationName 
      Height          =   225
      Left            =   840
      TabIndex        =   11
      Top             =   3480
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Label lblMileage 
      Height          =   225
      Left            =   2490
      TabIndex        =   10
      Top             =   3480
      Visible         =   0   'False
      Width           =   2385
   End
   Begin VB.Label Label5 
      Caption         =   "默认：新增站点插入到站点后"
      Height          =   255
      Left            =   150
      TabIndex        =   9
      Top             =   5640
      Width           =   2415
   End
End
Attribute VB_Name = "frmInsertReStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Public m_dtRunDate As Date
Private m_oRoutePrice As New RoutePriceTable
Private m_oRegularScheme As New RegularScheme
Private szPriceItem As Recordset
Private m_oParSystem As New SystemParam
Private m_nCountCol As Integer
Private m_nCountRows As Integer
Private m_oREBus As New REBus
Public m_szBusID As String
Private tTicketTpye() As TTicketType
Private tTicketTpye1() As TTicketType
Private m_szStationName As String '
Private m_sgFistMileage() As String
Private m_szStation As String
Private bTreFlg  As Boolean
Private tTReTicketPrice() As TRETicketPriceEx
Private m_szInsertStationId As String
Private szSeatType() As String
Private m_szExePriceTable As String
'Dim m_nCol As Integer
Dim nCol As Integer
Dim oPriceMan As New stprice.TicketPriceMan
Public m_oActiveUser As ActiveUser
Private Sub cmdAband_Click()
Dim vbMsg As VbMsgBoxResult
Dim nResult As Integer
On Error GoTo Here

 If txtBusStationID.Text <> "" Then
    If FindStationName(m_szStationName) = True Then
         vbMsg = MsgBox("是否将车次" & m_szBusID & "站点" & txtBusStationID.Text & "[" & m_szStationName & "]" & "删除", vbQuestion + vbYesNo, Me.Caption)
         If vbMsg = vbYes Then
               Me.MousePointer = vbHourglass
               m_oREBus.EnInsertStationAbandon txtBusStationID.Text
               frmEnvBusStation.OpenTime
               MsgBox "站点已被删除", vbInformation, Me.Caption
               Me.MousePointer = vbDefault
         End If
        Else
          MsgBox "该车次不经过该站", vbExclamation, Me.Caption
        End If
 Else
    MsgBox "请选取要删除车次站点", vbQuestion + vbYesNo, Me.Caption
 End If

 Exit Sub
Here:
   Me.MousePointer = vbDefault
'   ShowErrorU err.Number
End Sub

Private Sub cmdCaula_Click()
  MakeMfgData
  mfgStation.Col = nCol
End Sub
Private Sub cmdOk_Click()
  Dim i As Integer
  Dim sgData() As Single
  Dim nCount As Integer
  Dim nRows As Integer
  Dim nCols As Integer
  Dim vbMsg As String
  Dim j As Integer
  Dim bflg As Boolean
  Dim sgFullAndHalfTicket(1 To 2) As Single
  Dim nSerial As Integer
  Dim szStationIdName As String
  Dim szMsg As String
  nRows = mfgStation.Rows
  Dim nCountRow As Integer
  Dim bFlgEndStation As Boolean

  Dim szPreStationId As String
  Dim tTReTicketPriceTemp() As TRETicketPriceEx

  On Error GoTo Here
  
  MakeMfgData
  
  nCountRow = cobStationId.ListCount
  '取得站点序号
  If op1.Value = True Then
  nSerial = cobStationId.ListIndex + 2
  Else
   nSerial = cobStationId.ListIndex + 1
  End If
  '是否为终点站
  If cobStationId.ListCount < nSerial Then
   bFlgEndStation = True
  End If
  If txtBusStationID.Text = "" Or txtMileage.Text = "" Then
  MsgBox "请输入要插入的站点和里程", vbQuestion + vbYesNo, Me.Caption
  Exit Sub
  End If
  If op1.Value = False Then
  szMsg = szMsg & "将站点插入到" & cobStationId.Text & "站点前"
  Else
  szMsg = szMsg & "将站点插入到" & cobStationId.Text & "站点后"
  End If
  szMsg = szMsg & Chr(10) & "此操作将影响检票和售票....." & Chr(10) & "*是否保存数据"
  If m_szBusID = "" Then Exit Sub
  vbMsg = MsgBox(szMsg, vbQuestion + vbYesNo, "车次站点--插入站点")
  If vbMsg = vbNo Then Exit Sub
  Me.MousePointer = vbHourglass
'  MakeMfgData '计算票价
  tTReTicketPriceTemp = GetDateFromMfg '从界面得到数据
   m_oREBus.EnInsertStation nSerial, tTReTicketPriceTemp, bFlgEndStation
   frmEnvBusStation.mfgStation.AddItem Trim(txtBusStationID.Text)
   frmEnvBusStation.OpenTime
   cmdOk.Enabled = False
   Me.MousePointer = vbDefault
   MsgBox " 环境插入站点成功", vbInformation, "车次站点--插入站点"
   cmdAband.Enabled = True
   Me.MousePointer = vbDefault
   Exit Sub
Here:
  Me.MousePointer = vbDefault
  MsgBox " 环境插入站点失败" & Chr(10) & err.Description, vbInformation, "车次站点--插入站点"
End Sub

 Private Sub cmdExit_Click()
 Set m_oRoutePrice = Nothing
 Set m_oREBus = Nothing
 Set m_oParSystem = Nothing
 Unload Me
 End Sub





Private Sub Form_Load()
Dim i As Integer, j As Integer
Dim szaExePriceTable As String
On Error GoTo Here
    m_oRoutePrice.Init g_oActiveUser
    m_oRegularScheme.Init g_oActiveUser

    m_oREBus.Init g_oActiveUser
    m_oREBus.Identify m_szBusID, m_dtRunDate
    szaExePriceTable = m_oRegularScheme.GetRunPriceTableEx(m_dtRunDate) '得到使用票价表
    m_szExePriceTable = szaExePriceTable '得到使用票价表
    m_oRoutePrice.Identify m_szExePriceTable
    '
    oPriceMan.Init g_oActiveUser
    Set szPriceItem = oPriceMan.GetAllTicketItemRS(TP_PriceItemUse)
'    szPriceItem = m_oRoutePrice.GetUseTicketItem '得到使用的票价项

    m_oParSystem.Init g_oActiveUser

    tTicketTpye = m_oParSystem.GetAllTicketType(1, False)
    tTicketTpye1 = m_oParSystem.GetAllTicketType(1, False)


    m_nCountCol = szPriceItem.RecordCount
    If m_nCountCol = 0 Then Exit Sub
    mfgStation.Cols = m_nCountCol + 4
    mfgStation.AllowUserResizing = flexResizeNone
    mfgStation.TextArray(0) = "上车站"
    mfgStation.TextArray(1) = "座位类型"
    mfgStation.TextArray(2) = "票  种"
    mfgStation.TextArray(3) = "票  价"
    For i = 4 To m_nCountCol + 3
      mfgStation.TextArray(i) = szPriceItem!chinese_name
      szPriceItem.MoveNext
    Next
    mfgStation.Row = 0
    Me.Caption = Me.Caption & "[" & Format(m_dtRunDate, "YYYY年MM月DD日") & "]"
Exit Sub
Here:

   ' ShowErrorUerr.Number
End Sub
Private Sub FullBus()
 Dim i As Integer, j As Integer, k As Integer, h As Integer
 Dim nCountSeatType As Integer
 Dim nCountType As Integer
 Dim nRows As Integer
 Dim nCols As Integer
 Dim nTemp As Integer
 nCountType = ArrayLength(tTicketTpye)
 szSeatType = m_oREBus.GetReBusSeatType
 nCountSeatType = ArrayLength(szSeatType)
 Dim szSellStationID() As TReBusAllotInfo
 Dim nCountSellStation As Integer
 Dim m As Integer
 szSellStationID = m_oREBus.GetAllotInfo()
 nCountSellStation = ArrayLength(szSellStationID)
 With mfgStation
     .Rows = nCountSeatType * nCountType * nCountSellStation + 1
     nCols = .Cols
     .FixedCols = 0
     .Redraw = True
     .MergeCol(0) = True
     .MergeCol(1) = True
     
   For h = 1 To .Rows - 1
   For m = 1 To nCountSellStation
     For i = 1 To nCountSeatType
         .Row = h
         .MergeRow(i) = True
         .TextArray(h * nCols + 0) = MakeDisplayString(Trim(szSellStationID(m).szSellStationID), Trim(szSellStationID(m).szSellStationName))
         .TextArray(h * nCols + 1) = MakeDisplayString(Trim(szSeatType(i, 1)), Trim(szSeatType(i, 2)))
          For j = 2 To nCountType + 1
                If tTicketTpye(j - 1).nTicketTypeID = 1 Then
                    SetColor h, vbGreen
                End If
                For k = 3 To nCols - 1
                .TextArray(h * nCols + k) = 0
                Next
                .TextArray(h * nCols + 2) = tTicketTpye(j - 1).szTicketTypename
                .TextArray(h * nCols + 1) = MakeDisplayString(Trim(szSeatType(i, 1)), Trim(szSeatType(i, 2)))
                .TextArray(h * nCols + 0) = MakeDisplayString(Trim(szSellStationID(m).szSellStationID), Trim(szSellStationID(m).szSellStationName))
          h = h + 1
          .MergeCells = flexMergeRestrictColumns
          Next j
     Next i
    Next m
  Next h
     .FixedCols = 4
     .Redraw = True
 End With
 nCountSeatType = frmEnvBusStation.mfgStation.Rows
 For i = 1 To nCountSeatType - 1
   cobStationId.AddItem frmEnvBusStation.mfgStation.TextArray(i * frmEnvBusStation.mfgStation.Cols + 1)
 Next
 cobStationId.ListIndex = nCountSeatType - 2
 Set m_oParSystem = Nothing
 End Sub
Private Sub mfgStation_Scroll()
   txtInfo.Visible = False
End Sub
Private Sub TimStar_Timer()
Dim szItemp() As String
TimStar.Enabled = False
FullBus
End Sub
Private Sub mfgStation_Click()
Dim nCountTypePrice As Integer
Dim nCountSeatType As Integer
Dim nRow As Integer
Dim nCount As Integer

On Error GoTo Here
nCount = ArrayLength(tTicketTpye)

If mfgStation.Col = 1 Or mfgStation.Col = 2 Then
       txtInfo.Visible = False
Exit Sub
txtInfo.Visible = False
End If
If mfgStation.Row = 1 Or ((mfgStation.Row - 1) Mod nCount) = 0 Then
            txtInfo.Visible = True
            txtInfo.Top = mfgStation.Top + mfgStation.CellTop - 40
            txtInfo.Left = mfgStation.Left + mfgStation.CellLeft - 20
            txtInfo.Width = mfgStation.CellWidth
            txtInfo.Height = mfgStation.CellHeight
            txtInfo.Text = mfgStation.Text
            txtInfo.SetFocus
Else
            txtInfo.Visible = False
End If
Here:
End Sub
Private Sub txtBusStationID_ButtonClick()
    Dim szaTemp() As String
    Dim bflg As Boolean
    Dim szSeatTypeInfo() As String

    szSeatTypeInfo = m_oREBus.GetReBusSeatType '得到座位信息

    'frmSelVehicle.m_szfromstatus = "车次站点--插入站点"

'    szaTemp = ShowSelStation()
    Dim oShell As New CommDialog
    Dim aszTmp() As String
    oShell.Init g_oActiveUser
            szaTemp = oShell.SelectStation
    Set oShell = Nothing
    If ArrayLength(szaTemp) = 0 Then Exit Sub
    txtBusStationID.Text = MakeDisplayString(szaTemp(1, 1), szaTemp(1, 2))

    If ArrayLength(szaTemp) = 0 Then Exit Sub
    txtBusStationID.Text = Trim(szaTemp(1, 1))

    bflg = GetStationIDFromMfg(Trim(txtBusStationID.Text))

    m_szStationName = Trim(szaTemp(1, 2))
    m_szInsertStationId = Trim(szaTemp(1, 1))
    
    If bflg = False Then
       MsgBox "本线路已存在该站点", vbExclamation, "车次站点--插入站点"
       Exit Sub
    End If


    lblStationName.Caption = "插入站点 : " & m_szStationName
'    m_oREBus.Init g_oActiveUser
'    m_oREBus.Identify m_szBusID, m_dtRunDate

    '取得站点票价
    tTReTicketPrice = m_oREBus.GetStationPrice(szaTemp(1, 1), szSeatTypeInfo)

    '界面显示处理
    FullBusEx tTReTicketPrice

    cmdOk.Enabled = True
End Sub
Private Sub txtInfo_Change()
 Dim nRowCount As Integer
 Dim InputData As Single

nCol = mfgStation.Col
If txtInfo.Text <> "" And txtInfo.Text <> "." Then
   On Error GoTo Here
   InputData = CSng(txtInfo.Text)
End If
 If mfgStation.Text = txtInfo.Text Then Exit Sub
    mfgStation.CellForeColor = cvChangeColor
    
    mfgStation.Text = txtInfo.Text
    mfgStation.Col = 0
    mfgStation.CellForeColor = cvChangeColor
    mfgStation.Col = 1
    mfgStation.CellForeColor = cvChangeColor
    mfgStation.Col = 2
    mfgStation.CellForeColor = cvChangeColor

    mfgStation.Col = nCol

    cmdCaula.Enabled = True

    If mfgStation.Text <> "" Then
    End If
    cmdOk.Enabled = True
    Exit Sub
Here:
  txtInfo.Text = ""
End Sub

'得到当天票价表
Private Function GetExecuteTable() As String
    Dim oRegularScheme As New RegularScheme
    Dim szExitTable() As String
    Dim nExit As String
    'Dim tTemp As TSchemeArrangement
    Dim tTemp() As String
    oRegularScheme.Init g_oActiveUser
    tTemp = oRegularScheme.GetPriceTableInfo(Now)
     'tTemp = oRegularScheme.GetExecuteBusProject(Now)
     'szExitTable = oRegularScheme.ProjectExistTable(tTemp.szProjectID)
'    nExit = szExitTable(1, 2)
    nExit = tTemp(1, 2)
    GetExecuteTable = nExit
  End Function
 'nTicketNo---票种
Private Function GetTicketItemParam() As THalfTicketItemParamEX()
Dim szPriceTableID As String
On Error GoTo errorHander
szPriceTableID = Trim(GetExecuteTable)
    GetTicketItemParam = m_oREBus.GetItemParam(szPriceTableID)
     Exit Function
errorHander:
'   ShowErrorU err.Number
End Function
Private Function GetStationIDFromMfg(szStaionId As String) As Boolean
   Dim nRows As Integer
   Dim i As Integer
   Dim nCols As Integer
   Dim sgData As String
   nRows = frmEnvBusStation.mfgStation.Rows
   nCols = frmEnvBusStation.mfgStation.Cols
   i = 0
   Do While Trim(frmEnvBusStation.mfgStation.TextArray(i * nCols + 0)) <> szStaionId
    i = i + 1
    If i >= nRows Then
       GetStationIDFromMfg = True
       Exit Function
    End If
   Loop
   GetStationIDFromMfg = False
End Function
Private Function FullBusEx(tReTicket() As TRETicketPriceEx)
Dim i As Integer
Dim nCount As Integer
Dim j As Integer
nCount = ArrayLength(tReTicket) - 1
ReDim szSeatType(0 To 0)
bTreFlg = True
If nCount = -1 Then Exit Function
txtMileage.Text = tReTicket(1).sgMileage
 With mfgStation
     .Redraw = True
     .FixedCols = 0
     .Rows = nCount + 1
     .MergeCol(0) = True
     .MergeCol(1) = True
    For i = 1 To nCount

     On Error GoTo Here

       .MergeRow(i) = True

        If tReTicket(i).nTicketType = 1 Then
           SetColor i, vbGreen
        End If

       .TextArray(i * .Cols + 0) = MakeDisplayString(Trim(tReTicket(i).szSellStationID), Trim(tReTicket(i).szSellStationName))
       .TextArray(i * .Cols + 1) = MakeDisplayString(Trim(tReTicket(i).szSeatType), Trim(tReTicket(i).szSeatTypeName))
       .TextArray(i * .Cols + 2) = FindTicketTypeName(tReTicket(i).nTicketType)
       .TextArray(i * .Cols + 3) = Round(tReTicket(i).sgTotal, 1)
       .TextArray(i * .Cols + 4) = Round(tReTicket(i).sgBase, 2)
       For j = 5 To .Cols - 1
            If j <> .Cols - 1 Then
                .TextArray(i * .Cols + j) = Round(FindFullPlace(tReTicket(i).asgPrice, j - 4), 2)
            Else
                .TextArray(i * .Cols + j) = Round(FindFullPlace(tReTicket(i).asgPrice, 15), 2)
            End If
       Next
       .MergeCells = flexMergeRestrictColumns
Here:

    Next
       .FixedCols = 2
       .Redraw = True
End With
End Function
'取得票种名称
Private Function FindTicketTypeName(szType As Integer) As String
Dim i As Integer
Dim nCount As Integer
   i = 1
   nCount = ArrayLength(tTicketTpye)
   Do While Not Trim(tTicketTpye1(i).nTicketTypeID) = Trim(szType)
   i = i + 1
   Loop
   FindTicketTypeName = tTicketTpye1(i).szTicketTypename
End Function
Private Function FindTicketTypeID(szTicketTypename As String) As String
Dim i As Integer
Dim nCount As Integer
   i = 1
   nCount = ArrayLength(tTicketTpye)
   Do While Not Trim(tTicketTpye1(i).szTicketTypename) = Trim(szTicketTypename)
   i = i + 1
   Loop
   FindTicketTypeID = tTicketTpye1(i).nTicketTypeID
End Function
'取得票价项
Private Function FindFullPlace(dDataTemp() As Double, nI As Integer) As Double
Dim i As Integer
'nCount = szPriceItem.RecordCount
'If nI < nCount Then
'nCount = Val(szPriceItem(nI + 1, 1))
'FindFullPlace = dDataTemp(nCount)
FindFullPlace = dDataTemp(nI)
'End If
End Function
'取得数据
Private Function GetDateFromMfg() As TRETicketPriceEx()
Dim tTReTicketPriceTemp() As TRETicketPriceEx
Dim i As Integer
Dim nCount As Integer
Dim nRows As Integer
Dim nCols As Integer
Dim j As Integer
With mfgStation
nRows = .Rows
nCols = .Cols
ReDim tTReTicketPriceTemp(1 To nRows - 1)
For i = 1 To nRows - 1
    tTReTicketPriceTemp(i).szSellStationID = ResolveDisplay(GetLString(.TextArray(i * nCols + 0)))
    tTReTicketPriceTemp(i).szSellStationName = ResolveDisplayEx(GetLString(.TextArray(i * nCols + 0)))
    tTReTicketPriceTemp(i).szSeatType = GetLString(.TextArray(i * nCols + 1))
    tTReTicketPriceTemp(i).sgTotal = Val(.TextArray(i * nCols + 3))
    tTReTicketPriceTemp(i).sgBase = Val(.TextArray(i * nCols + 4))
    tTReTicketPriceTemp(i).szStationID = Trim(txtBusStationID.Text)
    tTReTicketPriceTemp(i).nTicketType = FindTicketTypeID(.TextArray(i * nCols + 2))
    tTReTicketPriceTemp(i).sgMileage = Val(txtMileage.Text)
    For j = 5 To nCols - 1
        tTReTicketPriceTemp(i).asgPrice(FindFullPlaceEx(j)) = Val(.TextArray(i * nCols + j))
    Next
Next
GetDateFromMfg = tTReTicketPriceTemp
End With
End Function
Public Function FindFullPlaceEx(nI As Integer) As Integer
Dim i As Integer
Dim nCount As Integer
nCount = szPriceItem.RecordCount
i = 1
'Do While Not Trim(szPriceItem(i, 2)) = Trim(mfgStation.TextArray(nI))
szPriceItem.MoveFirst
Do While Not Trim(szPriceItem!chinese_name) = Trim(mfgStation.TextArray(nI))
i = i + 1
szPriceItem.MoveNext
Loop
'nCount = Val(szPriceItem(i, 1))
nCount = Val(szPriceItem!price_item)
FindFullPlaceEx = nCount
End Function
Private Function MakeMfgData()
Dim i As Integer
Dim daItemPrice As Double   '分项票价
Dim daParam() As THalfTicketItemParamEX '计算票价项参数
'Dim daParam As Recordset
Dim dTolPrice As Double
Dim tReTicket() As TRETicketPriceEx
Dim nCountSeatType As Integer

Dim nRows As Integer
Dim nCols As Integer
Dim nCountType As Integer
Dim nFullPace As Integer
Dim FullPriceTol As Double
Dim k As Integer
Dim ncountTyp As Integer
Dim nCount As Integer
Dim dPriceTolalTemp As Double
Dim j As Integer
'取得座位数
nCountSeatType = ArrayLength(szSeatType)
ncountTyp = ArrayLength(tTicketTpye)
'得到所有可用票价项
daParam = GetTicketItemParam
' Set daParam = oPriceMan.GetAllTicketItemRS(TP_PriceItemAll)
nCount = ArrayLength(daParam)
'nCount = daParam.RecordCount
With mfgStation
  nRows = .Rows
  nCols = .Cols
  For i = 1 To nRows - 1
      .Row = i
      .Col = 0
      If .CellForeColor = vbBlue Then
          For j = 1 To ncountTyp - 1
              '按票种计算
              For k = 4 To nCols - 2
                '列计算
                nFullPace = FindFullPlaceEx(k) + 1
                daItemPrice = Val(.TextArray(i * nCols + k)) * daParam(j).asgParam1(nFullPace) + daParam(j).asgParam2(nFullPace)
                dTolPrice = dTolPrice + daItemPrice
                .TextArray((i + j) * nCols + k) = Round(daItemPrice, 2)
                If j = 1 Then
                  FullPriceTol = FullPriceTol + Val(.TextArray(i * nCols + k))
                End If
              Next
               '非全票填充
               dPriceTolalTemp = dTolPrice
'               .TextArray((i + j) * nCols + 3) = Round(dTolPrice, 1)  改为四舍五入到元
                .TextArray((i + j) * nCols + 3) = IIf(dTolPrice - Int(dTolPrice) >= 0.5, Int(dTolPrice) + 1, Int(dTolPrice))
               dPriceTolalTemp = Format(Val(.TextArray((i + j) * nCols + 3)) - dPriceTolalTemp, "0.00")
               .TextArray((i + j) * nCols + nCols - 1) = dPriceTolalTemp
               dTolPrice = 0
               '全票填充
                If j = 1 Then
                  dPriceTolalTemp = FullPriceTol
'                  .TextArray(i * nCols + 3) = Round(FullPriceTol, 1)  改为四舍五入到元
                   .TextArray((i) * nCols + 3) = IIf(FullPriceTol - Int(FullPriceTol) >= 0.5, Int(FullPriceTol) + 1, Int(FullPriceTol))
                  dPriceTolalTemp = Format(Val(.TextArray(i * nCols + 3) - dPriceTolalTemp), "0.00")
                  .TextArray(i * nCols + nCols - 1) = dPriceTolalTemp
                   FullPriceTol = 0
                   dPriceTolalTemp = 0
                End If
         Next
         i = ncountTyp + i - 1
      End If
  Next

End With
End Function
Private Function GetSeatType(szSeatTypeID As String)
Dim nCount As Integer
Dim i As Integer
nCount = ArrayLength(szSeatType)
i = 0
Do While Trim(szSeatType(i)) = Trim(szSeatTypeID)
  i = i + 1
  If i >= nCount Then
  ReDim Preserve szSeatType(0 To nCount)
  Exit Do
  End If
Loop
 szSeatType(i) = Trim(szSeatTypeID)
End Function
Private Function SetColor(nI As Integer, Color As Double)
Dim nCols As Integer
Dim nCol As Integer
Dim nRow As Integer
Dim i As Integer
With mfgStation
    nCols = .Cols
    nRow = .Row
    nCol = .Col
    .Row = nI
    For i = 1 To nCols - 1
    .Col = i
     .CellBackColor = Color
    Next
    .Row = nRow
    .Col = nCol
End With

End Function
Private Function FindStationName(szStationName As String) As Boolean
 Dim i As Integer
 Dim nCount As Integer
 nCount = cobStationId.ListCount
 Do While szStationName <> cobStationId.List(i)
 i = i + 1
  If i >= nCount Then
    FindStationName = False
    Exit Function
  End If
 Loop
 FindStationName = True
End Function



