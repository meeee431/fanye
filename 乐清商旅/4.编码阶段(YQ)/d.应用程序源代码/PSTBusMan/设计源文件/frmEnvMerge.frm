VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmEnvBusMerge 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "环境—车次并班"
   ClientHeight    =   6045
   ClientLeft      =   2490
   ClientTop       =   2700
   ClientWidth     =   9960
   HelpContextID   =   1000010
   Icon            =   "frmEnvMerge.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   9960
   Begin VB.CheckBox chkUnSplitMode 
      BackColor       =   &H00E0E0E0&
      Caption         =   "取消并班模式"
      ForeColor       =   &H8000000D&
      Height          =   345
      Left            =   6990
      TabIndex        =   30
      Top             =   810
      Width           =   1395
   End
   Begin MSComCtl2.DTPicker dtpBusDate 
      Height          =   285
      Left            =   3810
      TabIndex        =   3
      Top             =   173
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   503
      _Version        =   393216
      Format          =   25690113
      CurrentDate     =   37497
   End
   Begin RTComctl3.CoolButton cmdHelp 
      Height          =   345
      Left            =   8490
      TabIndex        =   14
      Top             =   630
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "帮助(&H)"
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
      MICON           =   "frmEnvMerge.frx":038A
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
      Height          =   345
      Left            =   8490
      TabIndex        =   13
      Top             =   180
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   609
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
      MICON           =   "frmEnvMerge.frx":03A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   4845
      Left            =   6870
      TabIndex        =   22
      Top             =   1110
      Width           =   2955
      Begin RTComctl3.CoolButton cmdMerge 
         Height          =   345
         Left            =   1320
         TabIndex        =   12
         Top             =   4380
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "并入(&M)"
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
         MICON           =   "frmEnvMerge.frx":03C2
         PICN            =   "frmEnvMerge.frx":03DE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin RTComctl3.CoolButton cmdBusInfo 
         Height          =   300
         Left            =   1650
         TabIndex        =   11
         Top             =   3840
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   529
         BTYPE           =   3
         TX              =   "车次信息"
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
         MICON           =   "frmEnvMerge.frx":0778
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin FText.asFlatTextBox txtNewBusID 
         Height          =   300
         Left            =   120
         TabIndex        =   10
         Top             =   3840
         Width           =   1440
         _ExtentX        =   2540
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
      Begin FText.asFlatTextBox txtStationID 
         Height          =   300
         Left            =   120
         TabIndex        =   6
         Top             =   420
         Width           =   2700
         _ExtentX        =   4763
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
      Begin MSComctlLib.ListView lvBusID 
         Height          =   2505
         Left            =   120
         TabIndex        =   8
         Top             =   1020
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   4419
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "车次"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "车型"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "普通"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "卧铺"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "加座"
            Object.Width           =   1058
         EndProperty
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   30
         X2              =   2910
         Y1              =   4275
         Y2              =   4275
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         X1              =   30
         X2              =   2910
         Y1              =   4260
         Y2              =   4260
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "途经站点车次列表(&L):"
         Height          =   180
         Left            =   120
         TabIndex        =   7
         Top             =   780
         Width           =   1800
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "并入车次(&B):"
         Height          =   180
         Left            =   120
         TabIndex        =   9
         Top             =   3600
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "途径站点(&S):"
         Height          =   180
         Left            =   120
         TabIndex        =   5
         Top             =   180
         Width           =   1080
      End
   End
   Begin FText.asFlatTextBox txtOldBusID 
      Height          =   300
      Left            =   1320
      TabIndex        =   1
      Top             =   165
      Width           =   1140
      _ExtentX        =   2011
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
   Begin MSComctlLib.ListView lvSeat 
      Height          =   4755
      Left            =   180
      TabIndex        =   4
      Top             =   1200
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   8387
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "票号"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "到站"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "原座号"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "票种"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "票价"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "并入车次"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "新座号"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblOffTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "07:00"
      Height          =   180
      Left            =   6630
      TabIndex        =   29
      Top             =   225
      Width           =   450
   End
   Begin VB.Label lblAllRefundment 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "是"
      Height          =   180
      Left            =   6630
      TabIndex        =   28
      Top             =   885
      Width           =   180
   End
   Begin VB.Label lblVehicleType 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "豪大11"
      Height          =   180
      Left            =   6270
      TabIndex        =   27
      Top             =   555
      Width           =   540
   End
   Begin VB.Label lblSellSeats 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      Height          =   180
      Left            =   3990
      TabIndex        =   26
      Top             =   885
      Width           =   180
   End
   Begin VB.Label lblVehicleLisenceTag 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "浙A.30300"
      Height          =   180
      Left            =   3990
      TabIndex        =   25
      Top             =   555
      Width           =   810
   End
   Begin VB.Label lblTotalSeats 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "30"
      Height          =   180
      Left            =   840
      TabIndex        =   24
      Top             =   885
      Width           =   180
   End
   Begin VB.Label lblRouteName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "甬杭州线"
      Height          =   180
      Left            =   1050
      TabIndex        =   23
      Top             =   555
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车次日期(&D):"
      Height          =   180
      Left            =   2610
      TabIndex        =   2
      Top             =   225
      Width           =   1080
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "总座位:"
      Height          =   180
      Left            =   210
      TabIndex        =   21
      Top             =   885
      Width           =   630
   End
   Begin VB.Label lblsdfsd 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车型:"
      Height          =   180
      Left            =   5760
      TabIndex        =   20
      Top             =   555
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车次代码(&I):"
      Height          =   180
      Index           =   0
      Left            =   210
      TabIndex        =   0
      Top             =   225
      Width           =   1080
   End
   Begin VB.Label lblkclcx 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "运行线路:"
      Height          =   180
      Left            =   210
      TabIndex        =   19
      Top             =   555
      Width           =   810
   End
   Begin VB.Label lbldsfsd 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "运行车辆:"
      Height          =   180
      Left            =   3180
      TabIndex        =   18
      Top             =   555
      Width           =   810
   End
   Begin VB.Label lblsdf 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "已售座数:"
      Height          =   180
      Left            =   3180
      TabIndex        =   17
      Top             =   885
      Width           =   810
   End
   Begin VB.Label lblrewrwe 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "全额退票:"
      Height          =   180
      Left            =   5760
      TabIndex        =   16
      Top             =   885
      Width           =   810
   End
   Begin VB.Label lblStartupTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "发车时间:"
      Height          =   180
      Left            =   5760
      TabIndex        =   15
      Top             =   225
      Width           =   810
   End
End
Attribute VB_Name = "frmEnvBusMerge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*********此处LvSeat中需要加入一个 并入车次与新座位号,以将并入后的车次有一个清楚的察看.
'*********在以后中完善
'ListView Seat  Columns Position
Const cnSerial = 0
Const cnTicketStatus = 2
Const cnStationName = 3
Const cnSeatTypeName = 4
Const cnSeatNO = 1
Const cnTicketTypeName = 5
Const cnTicketID = 6
Const cnTicketPrice = 7
Const cnOperationTime = 8
Const cnUserName = 9


'Const SplitSeat Color
Const cnSplitColor = vbRed
Const cnUnSplitColor = vbBlack


Public m_szBusID As String  '有外窗体传进的BusID
Public m_dtBusDate As Date
Public m_bIsParent As Boolean

Private m_oREBus As New REBus

Private m_aszSeatInfo() As String


Private Sub chkUnSplitMode_Click()
    If chkUnSplitMode.Value Then
        If SeekSplitedBus Then
            cmdMerge.Caption = "取消并入(&U)"
        Else
            chkUnSplitMode.Value = 0
        End If
    Else
        txtNewBusID.Text = ""
        cmdMerge.Caption = "并入(&M)"
    End If
End Sub

'查找原来被并入的车次信息
Private Function SeekSplitedBus() As Boolean
On Error GoTo ErrHandle
    
    Dim oRebus As New REBus
    Dim rsTemp As New Recordset
    Dim bCanUnSplit As Boolean
    Dim i As Integer
    Dim bFindMegerInfo As Boolean
    
    For i = 1 To lvSeat.ListItems.Count
        If lvSeat.ListItems(i).ForeColor = vbRed Then
            bCanUnSplit = True
        End If
    Next i
    
    If Not bCanUnSplit Then
        ShowMsg "没有可以取消并班的票号！"
        Exit Function
    End If
    
    oRebus.Init g_oActiveUser
    Set rsTemp = oRebus.GetMegerBusInfo(txtOldBusID.Text, dtpBusDate.Value)
    
    If rsTemp.RecordCount = 0 Then
        ShowMsg "找到到本车次有过的并班信息，不能取消并班！"
        bFindMegerInfo = False
    ElseIf rsTemp.RecordCount > 1 Then
        ShowMsg "找到本车次当日有过对次并入信息，将自动填充其中一个班次号到[并入班次]选择框里！"
        bFindMegerInfo = True
        txtNewBusID.Text = FormatDbValue(rsTemp!bus_id)
    Else
        bFindMegerInfo = True
        txtNewBusID.Text = FormatDbValue(rsTemp!bus_id)
    End If
    
    SeekSplitedBus = bFindMegerInfo

    Exit Function
ErrHandle:
    ShowErrorMsg
End Function

Private Sub cmdBusInfo_Click()
    ShowBusForm
End Sub

Private Sub ShowBusForm()
    '弹出显示车次的窗体
    Dim oShell As New CommDialog
    oShell.Init g_oActiveUser
    oShell.ShowBusInfo m_dtBusDate, txtNewBusID.Text
    
End Sub


Private Sub cmdCancel_Click()
    m_szBusID = ""
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    '显示帮助
    DisplayHelp Me

End Sub

Private Sub cmdMerge_Click()
    If Trim(txtNewBusID.Text) = "" Then
        ShowMsg "尚未选择并入车次!"
        Exit Sub
    End If
    
    If chkUnSplitMode.Value Then
        If MsgBox("请确保在取消并班之前，原先并到" & txtNewBusID.Text & "车次里的票已全部选中，否则取消并班将失败，是否继续取消并班操作！", vbInformation + vbYesNo, Me.Caption) = vbYes Then
            UnMerge '取消并班
        End If
    Else
        MergeBus '并班
    End If
End Sub

Private Sub dtpBusDate_LostFocus()
    If dtpBusDate.Value <> m_dtBusDate Then
        m_dtBusDate = dtpBusDate.Value
        RefreshBus
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()

    dtpBusDate.Value = m_dtBusDate
    txtOldBusID.Text = m_szBusID
    SetDefault
    FillLvBus
    FillLvSeat
    m_oREBus.Init g_oActiveUser
    m_oREBus.Identify txtOldBusID.Text, dtpBusDate.Value
    m_oREBus.Init g_oActiveUser
    RefreshBus

End Sub

Private Sub lvBusID_ItemClick(ByVal Item As MSComctlLib.ListItem)
    '将选择的车次赋予txtNewBusID
    txtNewBusID.Text = lvBusID.SelectedItem.Text
End Sub

Private Sub txtNewBusID_ButtonClick()
    '选择车次
    On Error GoTo ErrorHandle

    Dim aszTemp() As String
    Dim oShell As New CommDialog
    oShell.Init g_oActiveUser
    aszTemp = oShell.SelectBus(False)
    If ArrayLength(aszTemp) > 0 Then
        txtNewBusID.Text = aszTemp(1, 1)
    End If
    Set oShell = Nothing

    '刷新车次信息
    RefreshBus
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub txtOldBusID_ButtonClick()
    '选择车次
    On Error GoTo ErrorHandle
    Dim aszTemp() As String
    Dim oShell As New CommDialog
    oShell.Init g_oActiveUser
    aszTemp = oShell.SelectBus(False)
    If ArrayLength(aszTemp) > 0 Then
        txtOldBusID.Text = aszTemp(1, 1)
    End If
    Set oShell = Nothing

    '刷新车次信息
    If m_szBusID <> txtOldBusID.Text Then
        m_szBusID = txtOldBusID.Text
        RefreshBus
    End If
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub txtOldBusID_LostFocus()
    '刷新车次信息
    If m_szBusID <> txtOldBusID.Text Then
        m_szBusID = txtOldBusID.Text
        RefreshBus
    End If
End Sub

Private Sub txtStationID_ButtonClick()
    '选择站点
    Dim oShell As New CommDialog
    Dim aszTemp() As String
    oShell.Init g_oActiveUser
    aszTemp = oShell.SelectStation
    If ArrayLength(aszTemp) > 0 Then
        txtStationID.Text = MakeDisplayString(aszTemp(1, 1), aszTemp(1, 2))
        RefreshStationBus
    End If

End Sub

Private Sub RefreshStationBus()
    '刷新到该站的车次
    Dim i As Integer
    Dim nCount As Integer
    Dim liTemp As ListItem
    Dim aszBusID() As String

    On Error GoTo ErrorHandle
    If txtStationID.Text = "" Then Exit Sub
    aszBusID = m_oREBus.GetRemainSeatInfo(ResolveDisplay(txtStationID.Text))
    nCount = ArrayLength(aszBusID)
    lvBusID.ListItems.Clear
    For i = 1 To nCount
        Set liTemp = lvBusID.ListItems.Add(, , aszBusID(i, 1))
        liTemp.SubItems(1) = aszBusID(i, 2) 'vehicle_type_name
        liTemp.SubItems(2) = aszBusID(i, 8) 'Seat_count
        liTemp.SubItems(3) = aszBusID(i, 4) ' route_name
        liTemp.SubItems(4) = aszBusID(i, 9) 'bed_seat_count
        liTemp.SubItems(5) = aszBusID(i, 10) 'add_seat_count

    Next i
    If nCount > 0 Then
        lvBusID.ListItems(1).Selected = True
        txtNewBusID.Text = lvBusID.ListItems(1).Text
    End If
    Exit Sub
ErrorHandle:
    ShowErrorMsg

End Sub

Private Sub RefreshBus()
    '刷新车次信息
    On Error GoTo ErrorHandle
    m_oREBus.Identify txtOldBusID.Text, dtpBusDate.Value
    lblOffTime.Caption = m_oREBus.StartUpTime
    lblRouteName.Caption = m_oREBus.RouteName
    lblVehicleLisenceTag.Caption = m_oREBus.VehicleTag
    lblVehicleType.Caption = m_oREBus.VehicleModelName
    lblTotalSeats.Caption = m_oREBus.TotalSeat
    lblSellSeats.Caption = m_oREBus.SaledSeatCount
   '王候记改
    If m_oREBus.AllRefundment = False Then
        lblAllRefundment.Caption = "否"
    Else
        lblAllRefundment.Caption = "是"
    End If
    RefreshSeat
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub


Private Sub RefreshSeat()
    '刷新车次的座位信息
    Dim i As Integer
    Dim nCount As Integer
    Dim liTemp As ListItem
    m_aszSeatInfo = m_oREBus.GetBusSaleSeatInfo
    nCount = ArrayLength(m_aszSeatInfo)
    For i = 1 To nCount
        Set liTemp = lvSeat.ListItems.Add(, , i)
        liTemp.SubItems(cnTicketStatus) = GetTicketStatusStr(CInt(m_aszSeatInfo(i, 11)))  '票状态
        liTemp.SubItems(cnStationName) = Trim(m_aszSeatInfo(i, 1)) 'station_name
        liTemp.SubItems(cnSeatTypeName) = Trim(m_aszSeatInfo(i, 4)) 'seat_type_name
        liTemp.SubItems(cnSeatNO) = Trim(m_aszSeatInfo(i, 5)) 'seat_no
        liTemp.SubItems(cnTicketTypeName) = Trim(m_aszSeatInfo(i, 7)) 'ticket_type_name
        liTemp.SubItems(cnTicketID) = Trim(m_aszSeatInfo(i, 8)) 'ticket_id
        liTemp.SubItems(cnTicketPrice) = Trim(m_aszSeatInfo(i, 6)) 'ticket_price
'        liTemp.SubItems(cnOperationTime) = Format(m_aszSeatInfo(i, 3), "YYYY-MM-DD HH:MM:SS") 'operation_time
'        liTemp.SubItems(cnUserName) = Trim(m_aszSeatInfo(i, 2)) 'user_name

    Next
'    m_aszSeatInfo = m_oReBus.GetSlitpInfo(txtOldBusID.Text, dtpBusDate.Value)
'    nCount = ArrayLength(m_aszSeatInfo)
'    If nCount = 0 Then Exit Sub
'       For i = 1 To nCount
'            Set liTemp = lvSeat.ListItems.Add(, , m_aszSeatInfo(i, 4))
''            liTemp.subitems()= txtOldBusID.Text
'            liTemp.subitems()= m_aszSeatInfo(i, 8)
'            liTemp.subitems()= m_aszSeatInfo(i, 2)
'            liTemp.subitems()= m_aszSeatInfo(i, 4)
'            liTemp.subitems()= m_aszSeatInfo(i, 5)
'            liTemp.subitems()= m_aszSeatInfo(i, 6)
'            liTemp.subitems()= m_aszSeatInfo(i, 7)
'            liTemp.subitems()= m_aszSeatInfo(i, 8)
'            liTemp.subitems()= m_aszSeatInfo(i, 9)
'
'    Next
    SetLvColor
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub FillLvBus()
    '车次:599.811  车型:750.0473 普通:555.0236 卧铺:555.0236 加座:569.7638
    lvBusID.ColumnHeaders.Clear
    lvBusID.ColumnHeaders.Add , , "车次", 600
    lvBusID.ColumnHeaders.Add , , "车型", 750
    lvBusID.ColumnHeaders.Add , , "普通", 555
    lvBusID.ColumnHeaders.Add , , "线路", 1000
    lvBusID.ColumnHeaders.Add , , "卧铺", 555
    lvBusID.ColumnHeaders.Add , , "加座", 555
End Sub

Private Sub FillLvSeat()
    lvSeat.ColumnHeaders.Clear
    '序号:374.7402 票状态:900.2835             终点站:720.0001             座位类型:540.2835           座位号:540.2835             票种:540.2835 票号:900.2835 票价:540.2835 售票时间:1890.142           售票员:799.9371

    lvSeat.ColumnHeaders.Add , , "序", 374
    lvSeat.ColumnHeaders.Add , , "座位", 540
    lvSeat.ColumnHeaders.Add , , "状态", 900
    lvSeat.ColumnHeaders.Add , , "终点站", 720
    lvSeat.ColumnHeaders.Add , , "位型", 540
    lvSeat.ColumnHeaders.Add , , "票种", 540
    lvSeat.ColumnHeaders.Add , , "票号", 900
    lvSeat.ColumnHeaders.Add , , "票价", 540
'    lvSeat.ColumnHeaders.Add , , "售票时间", 1890
'    lvSeat.ColumnHeaders.Add , , "售票员", 800



'    '票号:1769.953 到站:824.882  原座号:720.0001             票种:689.9528 票价:629.8583 并入车次:1005.165           新座号:794.8347
'    lvSeat.ColumnHeaders.Add , , "票号", 1769
'    lvSeat.ColumnHeaders.Add , , "到站", 824
'    lvSeat.ColumnHeaders.Add , , "原座号", 720
'    lvSeat.ColumnHeaders.Add , , "票种", 689
'    lvSeat.ColumnHeaders.Add , , "票价", 629
'    lvSeat.ColumnHeaders.Add , , "并入车次", 1005
'    lvSeat.ColumnHeaders.Add , , "新座号", 794

End Sub


Private Sub MergeBus()
    '拆分
'    Dim i As Integer
    Dim m_aszSeatInfo() As String
    Dim nCount As Integer
    Dim szMsg As String



    If m_oREBus.busStatus <> EREBusStatus.ST_BusStopped And m_oREBus.busStatus <> EREBusStatus.ST_BusMergeStopped And m_oREBus.busStatus <> EREBusStatus.ST_BusSlitpStop Then
        '如果车次不是停班状态,则出错
        MsgBox "车次不处于停班状态,请先停班再进行拆分", vbInformation, Me.Caption
        Exit Sub
    End If
    If txtNewBusID.Text = txtOldBusID.Text Then
        MsgBox "被拆车次和目标车次相同，不能拆分", vbInformation, Me.Caption
        Exit Sub
    End If
    '取的拆分数据
    m_aszSeatInfo = GetSelectSeat
    nCount = ArrayLength(m_aszSeatInfo)
    If nCount = 0 Then Exit Sub
    On Error GoTo ErrorHandle

    szMsg = "是否将车次[" & txtOldBusID.Text & "]的[" & nCount & "]个座位拆分到[" & txtNewBusID.Text & "]车次" & Chr(10) & "目标车次[" & txtNewBusID.Text & "]可售座位"
    If MsgBox(szMsg, vbInformation + vbYesNo, Me.Caption) = vbNo Then Exit Sub
    '拆分
    m_oREBus.Identify txtOldBusID.Text, dtpBusDate.Value
    m_aszSeatInfo = m_oREBus.MegerBusAndSlitpBus(m_aszSeatInfo, txtNewBusID.Text)


    nCount = ArrayLength(m_aszSeatInfo)
    If nCount <> 0 Then
        MsgBox "拆分成功", vbInformation, Me.Caption
        If m_bIsParent = True Then
            frmEnvBus.UpdateList txtOldBusID.Text, dtpBusDate.Value
        End If
        SetLvColor
    End If
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Function GetSelectSeat() As String()
    '得到选择的座位信息
    Dim bAllowSplit As Boolean
    Dim oParam As New SystemParam
    Dim i As Integer
    Dim nCount As Integer
    Dim aszStationID() As String
    Dim j As Integer
    Dim szMsg As String
    Dim nStationCount As Integer
    Dim aszTemp() As String
    Dim nStatus As Integer '车次状态

'    Dim nCountTemp As Integer
'    Dim szBusInfo() As String
'    ReDim m_naIndexSlitpSeatInfo(1 To 1)
'    ReDim szBusInfo(1 To 1)
'    m_oReBus.Identify txtOldBusID.Text, dtpBusDate.Value
    On Error GoTo ErrorHandle

    oParam.Init g_oActiveUser
    '得到是否允许拆分
    bAllowSplit = oParam.AllowSlitp
    '得到该车次经过的站点
    aszStationID = m_oREBus.GetEnBusStationInfo
    nStationCount = ArrayLength(aszStationID)
    nCount = 0
    For i = 1 To lvSeat.ListItems.Count
        With lvSeat.ListItems(i)
            If .Selected Then
                '判断座位的有效性
                If Trim(.SubItems(cnTicketStatus)) = GetTicketStatusStr(ETicketStatus.ST_TicketChecked) Then
                    szMsg = szMsg & "第" & i & "行选取有误(座号为" & .SubItems(cnSeatNO) & ":)。原因:" & Chr(10)
                    szMsg = szMsg & "目标车次[" & txtOldBusID.Text & "]该座号[已检]"
                    szMsg = szMsg & Chr(10)
                    MsgBox szMsg & "拆分失败", vbInformation, Me.Caption
                    Set oParam = Nothing
                    Exit Function
                End If
                '该车是否经过该站点
                For j = 1 To nStationCount
                    If Trim(m_aszSeatInfo(i, 10)) = Trim(aszStationID(j, 2)) Then
                        Exit For
                    End If
                Next j
                If j > nStationCount Then
                    szMsg = szMsg & "第" & i & "行选取有误(座号为" & m_aszSeatInfo(i, 5) & ":).原因:" & Chr(10)
                    szMsg = szMsg & "目标车次" & txtOldBusID.Text & "不经过站点[" & m_aszSeatInfo(i, 1) & "]"
                    szMsg = szMsg & Chr(10)
                    If bAllowSplit = False Then '允许拆分不经过站点
                        MsgBox szMsg & "拆分失败", vbInformation, Me.Caption
                        szMsg = ""
                        Set oParam = Nothing
                        Exit Function
                    End If
                End If
                '判断每一张票是否已拆分,如果拆分则出错

                If .ForeColor = cnSplitColor Then
                    MsgBox .SubItems(cnSeatNO) & "座位已拆分不能重复拆分", vbInformation, Me.Caption
                    MsgBox .SubItems(cnSeatNO) & "座位已拆分不能重复拆分", vbInformation, Me.Caption
                    Exit Function
                End If
                nCount = nCount + 1
                ReDim Preserve aszTemp(1 To nCount)
                aszTemp(nCount) = .SubItems(cnTicketID)
                .ForeColor = cnSplitColor
            End If
        End With
    Next
    Set oParam = Nothing
    If szMsg <> "" Then
        If MsgBox(szMsg & "是否拆分？", vbYesNo + vbInformation, Me.Caption) = vbNo Then
            Exit Function
        End If
    End If
    GetSelectSeat = aszTemp
    Exit Function
ErrorHandle:
    ShowErrorMsg
End Function

Private Function GetUnMergeSelectSeat() As String()
    '得到选择的座位信息
    Dim bAllowSplit As Boolean
    Dim oParam As New SystemParam
    Dim i As Integer
    Dim nCount As Integer
    Dim aszStationID() As String
    Dim szMsg As String
    Dim nStationCount As Integer
    Dim aszTemp() As String
    Dim nStatus As Integer '车次状态

    On Error GoTo ErrorHandle

    oParam.Init g_oActiveUser
    '得到是否允许拆分
    bAllowSplit = oParam.AllowSlitp
    '得到该车次经过的站点
    aszStationID = m_oREBus.GetEnBusStationInfo
    nStationCount = ArrayLength(aszStationID)
    nCount = 0
    For i = 1 To lvSeat.ListItems.Count
        With lvSeat.ListItems(i)
            If .Selected Then
                '判断座位的有效性
                If Trim(.SubItems(cnTicketStatus)) = GetTicketStatusStr(ETicketStatus.ST_TicketChecked) Then
                    szMsg = szMsg & "第" & i & "行选取有误(座号为" & .SubItems(cnSeatNO) & ":)。原因:" & Chr(10)
                    szMsg = szMsg & "目标车次[" & txtOldBusID.Text & "]该座号[已检]"
                    szMsg = szMsg & Chr(10)
                    MsgBox szMsg & "取消拆分失败", vbInformation, Me.Caption
                    Set oParam = Nothing
                    Exit Function
                End If
   
                '判断每一张票是否已取消拆分,如果已取消拆分则出错
                If .ForeColor = cnUnSplitColor Then
                    MsgBox .SubItems(cnSeatNO) & "座位未经过拆分，不能重复取消拆分", vbInformation, Me.Caption
                    Exit Function
                End If
                nCount = nCount + 1
                ReDim Preserve aszTemp(1 To nCount)
                aszTemp(nCount) = .SubItems(cnTicketID)
                .ForeColor = cnUnSplitColor
            End If
        End With
    Next
    Set oParam = Nothing
    If szMsg <> "" Then
        If MsgBox(szMsg & "是否取消拆分？", vbYesNo + vbInformation, Me.Caption) = vbNo Then
            Exit Function
        End If
    End If
    GetUnMergeSelectSeat = aszTemp
    Exit Function
ErrorHandle:
    ShowErrorMsg
End Function


Private Sub UnMerge()
    '取消并班
    Dim m_aszSeatInfo() As String
    Dim nCount As Integer
    Dim szMsg As String


    If m_oREBus.busStatus <> EREBusStatus.ST_BusStopped And m_oREBus.busStatus <> EREBusStatus.ST_BusMergeStopped And m_oREBus.busStatus <> EREBusStatus.ST_BusSlitpStop Then
        '如果车次不是停班状态,则出错
        MsgBox "车次不处于停班状态,请先停班再进行取消并班", vbInformation, Me.Caption
        Exit Sub
    End If
'    If txtNewBusID.Text = txtOldBusID.Text Then
'        MsgBox "被拆车次和目标车次相同，不能拆分", vbInformation, Me.Caption
'        Exit Sub
'    End If
    
    '取的拆分数据
    m_aszSeatInfo = GetUnMergeSelectSeat
    nCount = ArrayLength(m_aszSeatInfo)
    If nCount = 0 Then Exit Sub
    On Error GoTo ErrorHandle

    szMsg = "是否取消对车次[" & txtOldBusID.Text & "]的[" & nCount & "]个座位的拆分操作"
    If MsgBox(szMsg, vbInformation + vbYesNo, Me.Caption) = vbNo Then Exit Sub
    '拆分
    m_oREBus.Identify txtOldBusID.Text, dtpBusDate.Value
    m_oREBus.UnMegerBusAndSlitpBus m_aszSeatInfo, txtNewBusID.Text


    nCount = ArrayLength(m_aszSeatInfo)
    If nCount <> 0 Then
        MsgBox "取消拆分成功", vbInformation, Me.Caption
        If m_bIsParent = True Then
            frmEnvBus.UpdateList txtOldBusID.Text, dtpBusDate.Value
        End If
        SetLvColor
    End If
    Exit Sub
ErrorHandle:
    ShowErrorMsg

End Sub

Private Sub SetDefault()
    '设置默认信息
    lblOffTime.Caption = ""
    lblRouteName.Caption = ""
    lblVehicleLisenceTag.Caption = ""
    lblVehicleType.Caption = ""
    lblTotalSeats.Caption = ""
    lblSellSeats.Caption = ""
    lblAllRefundment.Caption = ""
End Sub

Private Sub SetLvColor()
    '设置已拆分的座位的颜色
    Dim i As Integer
    Dim aszSplitTicket() As String
    Dim j As Integer
    Dim nCount As Integer

    If m_oREBus.busStatus = ST_BusSlitpStop Then
        '得到该车次的已拆分的票信息
        aszSplitTicket = m_oREBus.GetSlitpBusTicketNo
    Else
        Exit Sub
    End If
    nCount = ArrayLength(aszSplitTicket)
    For i = 1 To lvSeat.ListItems.Count
        For j = 1 To nCount
            With lvSeat.ListItems(i)
                If .SubItems(cnTicketID) = aszSplitTicket(j, 2) Then
                '如果发现此票为拆分票 ,则改变颜色
                    SetListViewLineColor lvSeat, i, cnSplitColor
                    Exit For
                End If
            End With
        Next j
    Next i
End Sub

Private Sub txtStationID_LostFocus()
    RefreshStationBus
End Sub
