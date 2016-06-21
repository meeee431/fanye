VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VsFlex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Begin VB.Form frmQuerySellSeat 
   BackColor       =   &H00E0E0E0&
   Caption         =   "车次站点已售票数统计"
   ClientHeight    =   4290
   ClientLeft      =   3705
   ClientTop       =   3330
   ClientWidth     =   5565
   Icon            =   "frmQuerySellSeat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   5565
   StartUpPosition =   1  '所有者中心
   Begin MSComCtl2.DTPicker dtRunData 
      Height          =   285
      Left            =   3840
      TabIndex        =   2
      Top             =   180
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   503
      _Version        =   393216
      Format          =   23658496
      CurrentDate     =   37237
      MinDate         =   -108324
   End
   Begin RTComctl3.CoolButton cmdExit 
      Cancel          =   -1  'True
      Height          =   330
      Left            =   4050
      TabIndex        =   5
      Top             =   3855
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
      MICON           =   "frmQuerySellSeat.frx":000C
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
      Height          =   330
      Left            =   2715
      TabIndex        =   4
      Top             =   3855
      Width           =   1185
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
      MICON           =   "frmQuerySellSeat.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VSFlex7LCtl.VSFlexGrid VsQuery 
      Height          =   3090
      Left            =   135
      TabIndex        =   3
      Top             =   630
      Width           =   5295
      _cx             =   9340
      _cy             =   5450
      _ConvInfo       =   -1
      Appearance      =   0
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
      BackColorFixed  =   14737632
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   14737632
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
   Begin FText.asFlatTextBox txtBusID 
      Height          =   285
      Left            =   1290
      TabIndex        =   6
      Top             =   180
      Visible         =   0   'False
      Width           =   1185
      _ExtentX        =   2090
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "发车日期(&D):"
      Height          =   180
      Left            =   2700
      TabIndex        =   1
      Top             =   225
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车次代号(&B):"
      Height          =   180
      Left            =   180
      TabIndex        =   0
      Top             =   225
      Width           =   1080
   End
End
Attribute VB_Name = "frmQuerySellSeat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''Option Explicit
''Public m_szBusID As String
''Public m_dyBusData As Date
''
''Private ticket() As TTicketType
''Private m_oReSheme As New REBus
''
''
''
''Private Sub cmdExit_Click()
''    Unload Me
''End Sub
''
''Private Sub cmdQuery_Click()
''    VsQuery.Rows = 2
''    FillVsQuery
''End Sub
''
''Private Sub Form_Load()
''    m_oReSheme.Init g_oActiveUser
''    txtBusID.Text = m_szBusID
''    dtRunData.Value = CDate(m_dyBusData)
''    InitVsQuery
''    FillVsQuery
''End Sub
''
''Private Sub InitVsQuery()
''    Dim i As Integer
''
''    VsQuery.Editable = flexEDNone
''    VsQuery.FixedRows = 1
''    VsQuery.MergeRow(0) = True
''    VsQuery.ColWidth(-1) = 1300
''    VsQuery.ColWidth(0) = 600
''    VsQuery.RowHeight(-1) = 300
''    VsQuery.Cols = 6
''    VsQuery.ColWidth(0) = 600
''    VsQuery.ColWidth(1) = 1000
''    VsQuery.ColWidth(2) = 900
''    VsQuery.ColWidth(3) = 900
''    VsQuery.ColWidth(4) = 900
''    VsQuery.ColWidth(5) = 900
''
''    VsQuery.TextMatrix(0, 0) = "序号"
''    VsQuery.TextMatrix(0, 1) = "站名"
''    VsQuery.TextMatrix(0, 2) = "全票张数"
''    VsQuery.TextMatrix(0, 3) = "半票张数"
''    VsQuery.TextMatrix(0, 4) = "其它票数"
''    VsQuery.TextMatrix(0, 5) = "总票数"
''
''End Sub
''Private Sub FillVsQuery()
''    Dim szBusInfo() As String
''    Dim nCount As Integer
''    Dim nCountTemp As Integer
''    Dim i As Integer
''    Dim j As Integer
''    Dim bflgSaleCount As Boolean
''    With VsQuery
''        .Rows = 1
''
''        If txtBusID.Text <> "" Then
''            szBusInfo = m_oReSheme.GetBusStatinSellInfo(ResolveDisplay(txtBusID.Text), CDate(dtRunData.Value))
''            nCount = ArrayLength(szBusInfo)
''            If nCount = 0 Then Exit Sub
''        End If
''
''        .MergeCells = flexMergeFixedOnly
''        '加一行
''        For i = 0 To nCount - 1
''            .Rows = .Rows + 1
''            j = .Rows - 1
''            bflgSaleCount = False
''            If i > 0 And CStr(.TextMatrix(j - 1, 1)) = CStr(szBusInfo(i, 1)) Then
''                bflgSaleCount = True
''                '如果站点不同，且不是第一次循环， 则VS加一行
''                .Rows = .Rows - 1
''                j = .Rows - 1
''            End If
''            .TextMatrix(j, 0) = j
''            .TextMatrix(j, 1) = szBusInfo(i, 1)
''            If .TextMatrix(j, 2) = "" Then .TextMatrix(j, 2) = 0
''            If .TextMatrix(j, 3) = "" Then .TextMatrix(j, 3) = 0
''            If .TextMatrix(j, 4) = "" Then .TextMatrix(j, 4) = 0
''            If szBusInfo(i, 2) <> 0 Then
''            '票种
''                Select Case CInt(szBusInfo(i, 3))
''                Case 1
''                    '全票
''                    .TextMatrix(j, 2) = CStr(CInt(.TextMatrix(j, 2)) + CInt(szBusInfo(i, 2)))
''                    nCountTemp = nCountTemp + CInt(.TextMatrix(j, 2))
''                Case 2
''                    '票
''                    .TextMatrix(j, 3) = CStr(CInt(.TextMatrix(j, 3)) + CInt(szBusInfo(i, 2)))
''                    nCountTemp = nCountTemp + CInt(.TextMatrix(j, 3))
''                Case Else
''                    '其它票
''                    .TextMatrix(j, 4) = CStr(CInt(.TextMatrix(j, 4)) + CInt(szBusInfo(i, 2)))
''                    nCountTemp = nCountTemp + CInt(.TextMatrix(j, 4))
''                End Select
''            End If
''            .TextMatrix(j, 5) = CStr(CInt(.TextMatrix(j, 3)) + CInt(.TextMatrix(j, 2)) + CInt(.TextMatrix(j, 4)))
''        Next i
''        Dim bflg As Boolean
''        bflg = False
''        On Error GoTo ErrorHandle
''        For i = 1 To .Rows
''            If i > .Rows Then Exit For
''            If bflg = True Then bflg = False: i = i - 1
''                .TextMatrix(i, 0) = i
''                If .TextMatrix(i, 5) = "0" Then
''                bflg = True
''                .RemoveItem i
''                .Refresh
''            End If
''        Next
''ErrorHandle:
''        If .Rows > 1 Then
''            .Rows = .Rows + 1
''            .MergeCol(.Cols - 1) = False
''            .MergeCells = flexMergeRestrictRows
''            .MergeRow(.Rows - 1) = True
''            For i = 1 To 5
''                .TextMatrix(.Rows - 1, i) = CStr(nCountTemp)
''                .MergeCol(i) = True
''            Next
''            .TextMatrix(.Rows - 1, 0) = "合计"
''        End If
''    End With
''End Sub
''
''
''
''
''Private Sub txtBusID_Click()
'''    Dim aszTemp() As String
'''
'''    g_dtdateSellQuery = CDate(dtRunData.Value)
'''    aszTemp = selectAllBus(, , True)
'''    If ArrayLength(aszTemp) = 0 Then Exit Sub
'''    txtBusID.Text = aszTemp(1, 1)
'''    m_szBusID = txtBusID.Text
''End Sub
