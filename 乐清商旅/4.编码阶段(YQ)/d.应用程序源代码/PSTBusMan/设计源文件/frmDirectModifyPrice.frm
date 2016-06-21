VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Object = "{BBF95DAA-F9CB-4CA9-A673-E0E9E0193752}#1.0#0"; "STSellCtl.ocx"
Begin VB.Form frmDirectModifyPrice 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "直接批量修改数据库中的票价"
   ClientHeight    =   5940
   ClientLeft      =   2295
   ClientTop       =   2640
   ClientWidth     =   8400
   Icon            =   "frmDirectModifyPrice.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   8400
   Begin VB.ListBox lstType 
      Appearance      =   0  'Flat
      Columns         =   2
      Height          =   2370
      Left            =   3720
      MultiSelect     =   2  'Extended
      TabIndex        =   8
      Top             =   2490
      Width           =   4455
   End
   Begin VB.ListBox lstSeatType 
      Appearance      =   0  'Flat
      Columns         =   2
      Height          =   930
      Left            =   6075
      MultiSelect     =   2  'Extended
      TabIndex        =   6
      Top             =   1260
      Width           =   2100
   End
   Begin VB.ListBox lstTicketType 
      Appearance      =   0  'Flat
      Height          =   930
      Left            =   3720
      MultiSelect     =   2  'Extended
      TabIndex        =   4
      Top             =   1260
      Width           =   2190
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "修改值"
      Height          =   930
      Left            =   150
      TabIndex        =   14
      Top             =   15
      Width           =   8025
      Begin VB.ComboBox cboItem 
         Height          =   300
         Left            =   765
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   465
         Width           =   1440
      End
      Begin VB.ComboBox cboItemSou 
         Height          =   300
         Left            =   2550
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   465
         Width           =   1575
      End
      Begin STSellCtl.ucNumTextBox txtRatio 
         Height          =   300
         Left            =   4455
         TabIndex        =   17
         Top             =   480
         Width           =   735
         _ExtentX        =   1296
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
         Text            =   ".1"
         MaxLength       =   6
         Alignment       =   1
      End
      Begin STSellCtl.ucNumTextBox txtAdd 
         Height          =   300
         Left            =   5775
         TabIndex        =   18
         Top             =   465
         Width           =   780
         _ExtentX        =   1376
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
         MaxLength       =   4
         Alignment       =   1
         AllowNegative   =   -1  'True
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "常数(&N)"
         Height          =   180
         Left            =   5865
         TabIndex        =   25
         Top             =   210
         Width           =   630
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "比率(&R)"
         Height          =   180
         Left            =   4365
         TabIndex        =   24
         Top             =   210
         Width           =   630
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "修改参照项(&F)"
         Height          =   180
         Left            =   2490
         TabIndex        =   23
         Top             =   210
         Width           =   1170
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "修改票价项(&I)"
         Height          =   180
         Left            =   765
         TabIndex        =   22
         Top             =   210
         Width           =   1170
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ")+"
         Height          =   180
         Left            =   5430
         TabIndex        =   21
         Top             =   525
         Width           =   180
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "×"
         Height          =   180
         Left            =   4200
         TabIndex        =   20
         Top             =   525
         Width           =   180
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "=("
         Height          =   180
         Left            =   2280
         TabIndex        =   19
         Top             =   525
         Width           =   180
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   3585
      Top             =   3225
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "同时进行尾数处理"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   3765
      TabIndex        =   9
      Top             =   5010
      Value           =   1  'Checked
      Width           =   1770
   End
   Begin MSComctlLib.ListView lvStation 
      Height          =   4065
      Left            =   120
      TabIndex        =   12
      Top             =   1770
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   7170
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imgStation"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "票种(&T):"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3720
      TabIndex        =   3
      Top             =   1020
      Width           =   2190
   End
   Begin VB.Frame framSelVehicleType 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "车型选择(&V):"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3720
      TabIndex        =   7
      Top             =   2250
      Width           =   4455
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "座位类型选择(&A):"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   6090
      TabIndex        =   5
      Top             =   1020
      Width           =   2100
   End
   Begin VB.ComboBox cboPriceTable 
      Height          =   300
      Left            =   1335
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1050
      Width           =   2220
   End
   Begin RTComctl3.CoolButton cmdOk 
      Height          =   315
      Left            =   5610
      TabIndex        =   11
      Top             =   4950
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "保存(&S)"
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
      MICON           =   "frmDirectModifyPrice.frx":000C
      PICN            =   "frmDirectModifyPrice.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   6975
      TabIndex        =   13
      Top             =   4950
      Width           =   1215
      _ExtentX        =   2143
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
      MICON           =   "frmDirectModifyPrice.frx":03C2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ImageList imgStation 
      Left            =   1725
      Top             =   1395
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDirectModifyPrice.frx":03DE
            Key             =   "Station"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDirectModifyPrice.frx":0538
            Key             =   "NoSell"
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   3705
      Picture         =   "frmDirectModifyPrice.frx":0AD2
      Top             =   5340
      Width           =   480
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "    注意：用此方法来修改票价需要极度小心,一旦修改无法恢复原票价,所以修改之前需要备份一下票价表或数据库。"
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   4200
      TabIndex        =   10
      Top             =   5340
      Width           =   3870
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "选择站点(&S):"
      Height          =   180
      Left            =   150
      TabIndex        =   2
      Top             =   1500
      Width           =   1080
   End
   Begin VB.Label lblExcuteTable 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "票价表(&P):"
      Height          =   180
      Left            =   150
      TabIndex        =   1
      Top             =   1110
      Width           =   900
   End
End
Attribute VB_Name = "frmDirectModifyPrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_oTicketPriceMan As New TicketPriceMan
Dim m_oBaseInfo As New BaseInfo
Dim m_oSystemParam As New SystemParam
Dim m_oRoutePriceTable As New RoutePriceTable



Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Dim aszStationID() As String
    Dim anTicketType() As Integer
    Dim aszSeatType() As String
    Dim aszVehicleModel() As String
    Dim dbAdd As Double
    Dim dbRatio As Double
    Dim szSourceItem As String
    Dim szDesItem As String
    Dim i As Integer, j As Integer
    
    On Error GoTo ErrorHandle
    dbAdd = txtAdd.Text
    dbRatio = txtRatio.Text
    szSourceItem = ResolveDisplay(cboItemSou.Text)
    szDesItem = ResolveDisplay(cboItem.Text)
    
    m_oRoutePriceTable.Init g_oActiveUser
    Dim szPriceTable As String
    m_oRoutePriceTable.Identify ResolveDisplay(cboPriceTable.Text, szPriceTable)
    
    
    If MsgBox("确实要对" & EncodeString(szPriceTable) & "进行票价的批量修改吗？", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
        SetBusy
        j = 0
        For i = 1 To lvStation.ListItems.Count
            If lvStation.ListItems(i).Selected = True Then
                j = j + 1
                ReDim Preserve aszStationID(1 To j)
                aszStationID(j) = lvStation.ListItems(i).Text
            End If
        Next i
        
        j = 0
        For i = 1 To lstSeatType.ListCount
            If lstSeatType.Selected(i - 1) = True Then
                j = j + 1
                ReDim Preserve aszSeatType(1 To j)
                aszSeatType(j) = ResolveDisplay(lstSeatType.List(i - 1))
            End If
        Next i
                
        j = 0
        For i = 1 To lstTicketType.ListCount
            If lstTicketType.Selected(i - 1) = True Then
                j = j + 1
                ReDim Preserve anTicketType(1 To j)
                anTicketType(j) = ResolveDisplay(lstTicketType.List(i - 1))
            End If
        Next i
        j = 0
        For i = 1 To lstType.ListCount
            If lstType.Selected(i - 1) = True Then
                j = j + 1
                ReDim Preserve aszVehicleModel(1 To j)
                aszVehicleModel(j) = ResolveDisplay(lstType.List(i - 1))
            End If
        Next i
        
            
        m_oRoutePriceTable.DirectModifyPrice aszStationID, anTicketType, aszSeatType, aszVehicleModel, True, dbRatio, szSourceItem, szDesItem, dbAdd
        '进行修改操作
        
        '应该可以根据数据库中返回的进度，显示修改的进度条
        SetNormal
        MsgBox "票价批量修改成功!", vbInformation
    End If
    Unload Me
    Exit Sub
ErrorHandle:
    SetNormal
End Sub

Private Sub Form_Load()
    FillLvHead
    '见计时器
End Sub
'===================================================
'Modify Date：2002-11-19
'Author:fl
'Reamrk:填充所有的票价表
'===================================================b
Private Sub FillPriceTable()
    '填充所有的票价表
    '默认选择当前的执行的票价表
    Dim i As Integer
    Dim nCount As Integer
    Dim szTemp As Integer
    Dim aszRoutePriceTable() As String
    On Error GoTo ErrorHandle
    m_oTicketPriceMan.Init g_oActiveUser
    aszRoutePriceTable = m_oTicketPriceMan.GetAllRoutePriceTable()
    nCount = ArrayLength(aszRoutePriceTable)
    For i = 1 To nCount
        cboPriceTable.AddItem MakeDisplayString(aszRoutePriceTable(i, 1), aszRoutePriceTable(i, 2))
    Next i
    If nCount > 0 Then cboPriceTable.ListIndex = 0
    
    Set m_oTicketPriceMan = Nothing
    Exit Sub
    
ErrorHandle:
    Set m_oTicketPriceMan = Nothing
    ShowErrorMsg
    
End Sub
'===================================================
'Modify Date：2002-11-19
'Author:fl
'Reamrk:填充所有的车型
'===================================================b
Private Sub FillVehicleType()
    '填充所有的车型
    '默认选择所有的车型
    Dim i As Integer
    Dim nCount As Integer
    Dim aszAllVehicleModel() As String
    On Error GoTo ErrorHandle
    m_oBaseInfo.Init g_oActiveUser
    aszAllVehicleModel = m_oBaseInfo.GetAllVehicleModel '无参数
    nCount = ArrayLength(aszAllVehicleModel)
    For i = 1 To nCount
        lstType.AddItem MakeDisplayString(Trim(aszAllVehicleModel(i, 1)), Trim(aszAllVehicleModel(i, 2)))
    Next i
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub


'===================================================
'Modify Date：2002-11-19
'Author:fl
'Reamrk:填充所有的票种
'===================================================b
Private Sub FillTicketType()
    '填充所有的票种
    '默认选择所有的票种
    Dim i As Integer
    Dim nCount As Integer
    Dim atAllTicketType() As TTicketType
    On Error GoTo ErrorHandle
    
    lstTicketType.Clear
    m_oSystemParam.Init g_oActiveUser
    atAllTicketType = m_oSystemParam.GetAllTicketType
    nCount = ArrayLength(atAllTicketType)
    For i = 1 To nCount
        lstTicketType.AddItem MakeDisplayString(Trim(atAllTicketType(i).nTicketTypeID), Trim(atAllTicketType(i).szTicketTypeName))
    Next i
    Set m_oSystemParam = Nothing
    Exit Sub
ErrorHandle:
    Set m_oSystemParam = Nothing
    ShowErrorMsg
    

    
End Sub
'===================================================
'Modify Date：2002-11-19
'Author:fl
'Reamrk:填充所有的座位类型
'===================================================b
Private Sub FillSeatType()
    '填充所有的座位类型
    '默认选择所有的座位类型
    Dim i As Integer
    Dim nCount As Integer
    Dim aszAllSeatType() As String
    On Error GoTo ErrorHandle
    m_oBaseInfo.Init g_oActiveUser
    aszAllSeatType = m_oBaseInfo.GetAllSeatType
    nCount = ArrayLength(aszAllSeatType)
    For i = 1 To nCount
        lstSeatType.AddItem MakeDisplayString(Trim(aszAllSeatType(i, 1)), Trim(aszAllSeatType(i, 2)))
    Next i
    Set m_oBaseInfo = Nothing
    Exit Sub
ErrorHandle:
    Set m_oBaseInfo = Nothing
    ShowErrorMsg
    
    
End Sub
'===================================================
'Modify Date：2002-11-19
'Author:fl
'Reamrk:填充所有的站点
'===================================================b
Private Sub FillStation()
    '填充所有的站点
    '默认选择所有的站点
    Dim i As Integer
    Dim liTemp As ListItem
    Dim nCount As Integer
    Dim aszStation() As String
    On Error GoTo ErrorHandle
    m_oBaseInfo.Init g_oActiveUser
    aszStation = m_oBaseInfo.GetStation
    nCount = ArrayLength(aszStation)
    lvStation.ListItems.Clear
    '还无参数
    For i = 1 To nCount
        WriteProcessBar True, i, nCount, "正在刷新站点"
        Set liTemp = lvStation.ListItems.Add(, GetEncodedKey(RTrim(aszStation(i, 1))), RTrim(aszStation(i, 1)))
        liTemp.ListSubItems.Add , , aszStation(i, 2)
        liTemp.ListSubItems.Add , , aszStation(i, 3)
        liTemp.ListSubItems.Add , , aszStation(i, 6)
    Next i
    WriteProcessBar False
    If lvStation.ListItems.Count > 0 Then lvStation.ListItems(1).Selected = True
    Set m_oBaseInfo = Nothing
    Exit Sub
ErrorHandle:
    Set m_oBaseInfo = Nothing
    ShowErrorMsg
    
End Sub

'===================================================
'Modify Date：2002-11-19
'Author:fl
'Reamrk:填充所有的票价表
'===================================================b'
Private Sub FillPriceItem()
    '填充所有的票价项
    '填充两个组合框
    '默认选择第一个票价项
    Dim nCount As Integer
    Dim i As Integer
    Dim aszAllRoutePriceTable() As String
    On Error GoTo ErrorHandle
    m_oTicketPriceMan.Init g_oActiveUser
    aszAllRoutePriceTable = m_oTicketPriceMan.GetAllTicketItem
    cboItem.Clear
    cboItemSou.Clear
    nCount = ArrayLength(aszAllRoutePriceTable)
    For i = 1 To nCount
        If aszAllRoutePriceTable(i, 3) = TP_PriceItemUse Then
            cboItemSou.AddItem MakeDisplayString(Trim(aszAllRoutePriceTable(i, 1)), Trim(aszAllRoutePriceTable(i, 2)))
            cboItem.AddItem MakeDisplayString(Trim(aszAllRoutePriceTable(i, 1)), Trim(aszAllRoutePriceTable(i, 2)))
        End If
    Next i
    If nCount > 0 Then cboItem.ListIndex = 0
    If nCount > 0 Then cboItemSou.ListIndex = 0
    
    Set m_oTicketPriceMan = Nothing
   
    Exit Sub
    
ErrorHandle:
    Set m_oTicketPriceMan = Nothing
    ShowErrorMsg
End Sub

Private Sub Label1_Click()

End Sub

Private Sub lvStation_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    SortListView lvStation, ColumnHeader.Index
End Sub

Private Sub Timer1_Timer()
    SetBusy
    Timer1.Enabled = False
    FillPriceTable
    FillVehicleType
    FillTicketType
    FillSeatType
    FillStation
    FillPriceItem
    SetNormal
End Sub

Private Sub txtMul_Change()

End Sub

'Private Sub txtRatio_Change()
'    FormatTextToNumeric txtRatio, False
'End Sub



Private Sub txtRatio_Validate(Cancel As Boolean)
    If txtRatio.Text < 0 Or txtRatio.Text > 1 Then
        MsgBox "比率必须大于等于0且小于等于1"
        Cancel = True
    End If
End Sub

Private Sub FillLvHead()
    '填充所有的列首
    lvStation.ColumnHeaders.Add , , "代码", 800
    lvStation.ColumnHeaders.Add , , "名称", 1000
    lvStation.ColumnHeaders.Add , , "输入码", 800
    lvStation.ColumnHeaders.Add , , "地区名称", 900
End Sub
