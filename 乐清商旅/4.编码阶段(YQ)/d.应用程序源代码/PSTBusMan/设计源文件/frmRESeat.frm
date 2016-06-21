VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmEnvReserveSeat 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "环境--座位管理"
   ClientHeight    =   5250
   ClientLeft      =   2445
   ClientTop       =   3420
   ClientWidth     =   9450
   HelpContextID   =   10000810
   Icon            =   "frmRESeat.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   9450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   1245
      Left            =   150
      TabIndex        =   5
      Top             =   30
      Width           =   7620
      Begin VB.Label lblBusID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   945
         TabIndex        =   13
         Top             =   270
         Width           =   90
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车次代码:"
         Height          =   180
         Left            =   120
         TabIndex        =   12
         Top             =   270
         Width           =   810
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "发车时间:"
         Height          =   180
         Left            =   120
         TabIndex        =   11
         Top             =   585
         Width           =   810
      End
      Begin VB.Label lblStartupTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   945
         TabIndex        =   10
         Top             =   585
         Width           =   90
      End
      Begin VB.Label lblSell 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "已售座位数:"
         Height          =   180
         Left            =   120
         TabIndex        =   9
         Top             =   900
         Width           =   990
      End
      Begin VB.Label lblReserve 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "预留座位数:"
         Height          =   180
         Left            =   3810
         TabIndex        =   8
         Top             =   900
         Width           =   990
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "总座位数:"
         Height          =   180
         Left            =   3810
         TabIndex        =   7
         Top             =   585
         Width           =   810
      End
      Begin VB.Label lblBusDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车次日期:"
         Height          =   180
         Left            =   3810
         TabIndex        =   6
         Top             =   270
         Width           =   810
      End
   End
   Begin RTComctl3.CoolButton cmdRefresh 
      Height          =   345
      Left            =   8070
      TabIndex        =   3
      Top             =   555
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "刷新(&R)"
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
      MICON           =   "frmRESeat.frx":014A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Timer tmStart 
      Interval        =   500
      Left            =   5820
      Top             =   3420
   End
   Begin RTComctl3.CoolButton cmdHelp 
      Height          =   345
      Left            =   8070
      TabIndex        =   4
      Top             =   990
      Width           =   1245
      _ExtentX        =   2196
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
      MICON           =   "frmRESeat.frx":0166
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ImageList imgSeat 
      Left            =   2955
      Top             =   2205
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRESeat.frx":0182
            Key             =   "reserved"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRESeat.frx":02DE
            Key             =   "normal"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRESeat.frx":043A
            Key             =   "edit"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRESeat.frx":078D
            Key             =   "del"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRESeat.frx":0AE0
            Key             =   "add"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRESeat.frx":0E33
            Key             =   "sell"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRESeat.frx":0F8D
            Key             =   "booked"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRESeat.frx":12A7
            Key             =   "projectBooked"
         EndProperty
      EndProperty
   End
   Begin RTComctl3.CoolButton cmdOk 
      Cancel          =   -1  'True
      Default         =   -1  'True
      Height          =   345
      Left            =   8070
      TabIndex        =   2
      Top             =   165
      Width           =   1245
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
      MICON           =   "frmRESeat.frx":15C1
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.Toolbar tbSeatManage 
      Height          =   360
      Left            =   5865
      TabIndex        =   0
      Top             =   1350
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "imgSeat"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "AddSeat"
            Object.ToolTipText     =   "新增座位"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "EditSeat"
            Object.ToolTipText     =   "修改座位"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DelSeat"
            Object.ToolTipText     =   "删除座位"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ReserveSeat"
            Object.ToolTipText     =   "预留座位"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "UnReserveSeat"
            Object.ToolTipText     =   "取消预留座位"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ListView lvSeat 
      Height          =   3120
      Left            =   150
      TabIndex        =   15
      Top             =   1695
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   5503
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imgSeat"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "座位号"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "座位种类"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "状态"
         Object.Width           =   177
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "票号"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "到站"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "票价"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "备注"
         Object.Width           =   14111
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "证件类型"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "姓名"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "证件号"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblSeatInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "其它:"
      Height          =   255
      Left            =   150
      TabIndex        =   14
      Top             =   4995
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车次座位列表(&L):"
      Height          =   180
      Left            =   120
      TabIndex        =   1
      Top             =   1455
      Width           =   1440
   End
End
Attribute VB_Name = "frmEnvReserveSeat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'车次座位管理

Public m_szBusID As String
Public m_dtEnvDate As Date
Private m_oReBus As New REBus
Private maszSeatType() As String
Private mbEdited As Boolean     '是否已经变更

Private Sub cmdHelp_Click()
DisplayHelp Me
End Sub
'修改座位
Private Sub EditSeat()
    
    frmSeat.m_szSeatNo = lvSeat.SelectedItem
    frmSeat.m_SeatType = maszSeatType
    frmSeat.efrmSeat = AddSeatType
    Set frmSeat.m_oReBus = m_oReBus
    frmSeat.Show vbModal
    FullSeat
End Sub
Private Sub cmdOk_Click()
    Unload Me
End Sub

Private Sub cmdRefresh_Click()
    RefreshBus
    FullSeat
End Sub

Private Sub Form_Load()
   AlignFormPos Me

    Dim tSeat() As TSeatInfoEx
    Dim nCount As Integer
    Dim liTemp As ListItem
    Dim i As Integer
    
    m_oReBus.Init g_oActiveUser
    
    If m_szBusID = "" Then m_dtEnvDate = Date
    
    lblBusDate.Caption = "车次日期:" & Format(m_dtEnvDate, "YYYY年MM月DD日")
    Me.Caption = "环境车次座位"
    
    If m_szBusID <> "" Then
       m_oReBus.Identify m_szBusID, m_dtEnvDate
       maszSeatType = m_oReBus.GetReBusSeatType
    End If
    
     AlignHeadWidth Me.name, lvSeat

'    tbSeatManage.Buttons("DelSeat").Enabled = False
'    tbSeatManage.Buttons("EditSeat").Enabled = False
'    tbSeatManage.Buttons("ReserveSeat").Enabled = False
'    tbSeatManage.Buttons("UnReserveSeat").Enabled = False
End Sub
'填充座位信息
Public Function FullSeat(Optional RefreshSeat As Boolean) As Boolean
    Dim tSeat() As TSeatInfoEx2
    Dim nCount As Integer
    Dim liTemp As ListItem
    Dim i As Integer
    Dim nSlitp As Integer, nReplace As Integer
    Dim szSeatInfo As String
    Dim nSeatCount As Integer
    On Error GoTo ErrHandle
    MousePointer = vbHourglass
    lvSeat.ListItems.Clear
    If m_szBusID <> "" Then
        ShowSBInfo "获得座位信息..."
        '        m_oREBus.Identify m_szBusID, m_dtEnvDate
        tSeat = m_oReBus.GetSeatInfoEx
        nCount = ArrayLength(tSeat)
        For i = 1 To nCount
            Select Case tSeat(i).szSeatStatus
            
            Case ST_SeatReserved
                Set liTemp = lvSeat.ListItems.Add(, , tSeat(i).szSeatNo, , "reserved")
                liTemp.SubItems(1) = MakeDisplayString(Trim(tSeat(i).szSeatType), Trim(tSeat(i).szSeatTypeName))
                liTemp.SubItems(2) = "预留"
                liTemp.SubItems(6) = tSeat(i).szRemark
            Case ST_SeatCanSell
                Set liTemp = lvSeat.ListItems.Add(, , tSeat(i).szSeatNo, , "normal")
                liTemp.SubItems(1) = MakeDisplayString(Trim(tSeat(i).szSeatType), Trim(tSeat(i).szSeatTypeName))
                liTemp.SubItems(2) = "可售"
            Case ST_SeatSold
                Set liTemp = lvSeat.ListItems.Add(, , tSeat(i).szSeatNo, , "sell")
                liTemp.SubItems(1) = MakeDisplayString(Trim(tSeat(i).szSeatType), Trim(tSeat(i).szSeatTypeName))
                liTemp.SubItems(2) = "已售"
                liTemp.SubItems(3) = tSeat(i).szTicketNo
                liTemp.SubItems(4) = tSeat(i).szDestName
                liTemp.SubItems(5) = tSeat(i).szTicketPrice
                liTemp.SubItems(7) = tSeat(i).szCardType
                liTemp.SubItems(8) = tSeat(i).szPersonName
                liTemp.SubItems(9) = tSeat(i).szIDCardNo
            Case ST_SeatSlitp
                Set liTemp = lvSeat.ListItems.Add(, , tSeat(i).szSeatNo, , "normal")
                liTemp.SubItems(1) = MakeDisplayString(Trim(tSeat(i).szSeatType), Trim(tSeat(i).szSeatTypeName))
                liTemp.SubItems(2) = "拆分得到"
                liTemp.SubItems(3) = tSeat(i).szTicketNo
                liTemp.SubItems(4) = tSeat(i).szDestName
                liTemp.SubItems(5) = tSeat(i).szTicketPrice
                liTemp.SubItems(7) = tSeat(i).szCardType
                liTemp.SubItems(8) = tSeat(i).szPersonName
                liTemp.SubItems(9) = tSeat(i).szIDCardNo
                nSlitp = nSlitp + 1
            Case ST_SeatReplace
                Set liTemp = lvSeat.ListItems.Add(, , tSeat(i).szSeatNo, , "normal")
                liTemp.SubItems(1) = MakeDisplayString(Trim(tSeat(i).szSeatType), Trim(tSeat(i).szSeatTypeName))
                liTemp.SubItems(2) = "顶班得到"
                
                liTemp.SubItems(3) = tSeat(i).szTicketNo
                liTemp.SubItems(4) = tSeat(i).szDestName
                liTemp.SubItems(5) = tSeat(i).szTicketPrice
                liTemp.SubItems(7) = tSeat(i).szCardType
                liTemp.SubItems(8) = tSeat(i).szPersonName
                liTemp.SubItems(9) = tSeat(i).szIDCardNo
                nReplace = nReplace + 1
            Case ST_SeatBooked
                Set liTemp = lvSeat.ListItems.Add(, , tSeat(i).szSeatNo, , "booked")
                liTemp.SubItems(1) = MakeDisplayString(Trim(tSeat(i).szSeatType), Trim(tSeat(i).szSeatTypeName))
                liTemp.SubItems(2) = "预定"
            Case ST_SeatProjectBooked
                Set liTemp = lvSeat.ListItems.Add(, , tSeat(i).szSeatNo, , "projectBooked")
                liTemp.SubItems(1) = MakeDisplayString(Trim(tSeat(i).szSeatType), Trim(tSeat(i).szSeatTypeName))
                liTemp.SubItems(2) = "计划预定"
            End Select
            If nReplace <> 0 Then
            szSeatInfo = "由于顶班得到座位[" & nReplace & "]个"
            End If
            If nSlitp <> 0 Then
                szSeatInfo = "由于拆分得到座位[" & nSlitp & "]个"
            End If
            lblSeatInfo.Visible = True
            lblSeatInfo.Caption = szSeatInfo
        Next i
    End If
    FullSeat = True
    SetNormal
    Exit Function
ErrHandle:
    SetNormal
    ShowErrorMsg
End Function

Public Sub DeleteSeat()
    On Error GoTo ErrHandle
    Dim ErrStr As String
    Dim szSeatNo As String
    Dim szSeatNoLog As String
    Dim i As Integer
    Dim oItem As ListItem
    MousePointer = vbHourglass
    
    For i = lvSeat.ListItems.Count To 1 Step -1
        Set oItem = lvSeat.ListItems(i)
        If oItem.Selected Then
            szSeatNo = szSeatNo & "'" & Format(oItem.Text, "00") & "',"
            szSeatNoLog = szSeatNoLog & "[" & Format(oItem.Text, "00") & "]"
        End If
    Next
    
    
    If szSeatNo <> "" Then
    szSeatNo = Left(szSeatNo, Len(szSeatNo) - 1)
    szSeatNo = szSeatNo & szSeatNoLog
        m_oReBus.DeleteSeat 1, szSeatNo
        
        For i = lvSeat.ListItems.Count To 1 Step -1
         Set oItem = lvSeat.ListItems(i)
            If oItem.Selected Then
                ShowSBInfo "删除座位" & oItem.Text
                'DeleteSeatEx oItem
                lvSeat.ListItems.Remove oItem.Index
            End If
        Next

    End If
    SetNormal
    Exit Sub
ErrHandle:
    SetNormal
    ShowErrorMsg
End Sub


'预留座位
Private Sub ReserveSeat()
    On Error GoTo ErrHandle
    Dim ErrStr As String
    Dim szSeatNo As String
    Dim szSeatNoLog As String
    Dim oItem As ListItem
    Dim szRemark As String
    Dim i As Integer
    MousePointer = vbHourglass
    For Each oItem In lvSeat.ListItems
        If oItem.Selected Then
             ShowSBInfo "预留座位" & oItem.Text
             szSeatNo = szSeatNo & "'" & oItem.Text & "',"
             szSeatNoLog = szSeatNoLog & "[" & oItem.Text & "]"
        End If
    Next
    
    If szSeatNo <> "" Then
        szSeatNo = Left(szSeatNo, Len(szSeatNo) - 1)
        szSeatNo = szSeatNo & szSeatNoLog
        szRemark = "此座位由" & g_oActiveUser.UserName & "于" & ToDBDateTime(Now) & "预留"
        m_oReBus.ReserveSeat szSeatNo, szRemark
        For i = 1 To lvSeat.ListItems.Count
            
            Set oItem = lvSeat.ListItems(i)
            If oItem.Selected Then
                 If oItem.SmallIcon = "normal" Then
                  oItem.SubItems(2) = "预留"
                  oItem.SubItems(6) = szRemark
                 End If
                 oItem.SmallIcon = "reserved"
            End If
            
        Next i
    End If
    
    SetNormal
    Exit Sub
ErrHandle:
    SetNormal
    ShowErrorMsg
End Sub


Private Sub UnReserveSeat()
    On Error GoTo ErrHandle
    Dim ErrStr As String
    Dim szSeatNo As String
    Dim szSeatNoLog As String

    Dim oItem As ListItem
    MousePointer = vbHourglass
    For Each oItem In lvSeat.ListItems
        If oItem.Selected Then
            ShowSBInfo "取消预留" & oItem.Text
            If oItem.SmallIcon = "reserved" Or oItem.SmallIcon = "projectBooked" Then
              ' UnReserveSeatEx oItem
              szSeatNo = szSeatNo & "'" & oItem.Text & "',"
              szSeatNoLog = szSeatNoLog & "[" & oItem.Text & "]"
            End If
        End If
    Next
    
    If szSeatNo <> "" Then
       szSeatNo = Left(szSeatNo, Len(szSeatNo) - 1)
       szSeatNo = szSeatNo & szSeatNoLog
       m_oReBus.UnReserveSeat szSeatNo
       
       For Each oItem In lvSeat.ListItems
        If oItem.Selected Then
            If oItem.SmallIcon = "reserved" Or oItem.SmallIcon = "projectBooked" Then
               oItem.ListSubItems(2).Text = "可售"
               oItem.ListSubItems(6).Text = ""
               oItem.SmallIcon = "normal"
            End If
        End If
      Next
    
    End If
    SetNormal
    Exit Sub
ErrHandle:
    SetNormal
    ShowErrorMsg
End Sub



Private Sub Form_Unload(Cancel As Integer)
    If mbEdited Then
        frmEnvBus.UpdateList m_szBusID, m_dtEnvDate
        mbEdited = False
    End If
    
    Set m_oReBus = Nothing
    SaveFormPos Me
    SaveHeadWidth Me.name, lvSeat
End Sub

Private Sub lvSeat_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    SortListView lvSeat, ColumnHeader.Index
End Sub

Private Sub lvSeat_DblClick()
    Dim oSncom As New STShell.CommDialog
    Dim szTicketID As String
On Error GoTo ErrHandle
     
    If lvSeat.SelectedItem Is Nothing Then Exit Sub
    If lvSeat.SelectedItem.ListSubItems(2).Text = "已售" Or lvSeat.SelectedItem.ListSubItems(2).Text = "顶班得到" Or lvSeat.SelectedItem.ListSubItems(2).Text = "拆分得到" Then
        szTicketID = ResolveDisplay(lvSeat.SelectedItem.ListSubItems(3).Text)
    End If
    If szTicketID = "" Then Exit Sub
    oSncom.Init g_oActiveUser
    oSncom.ShowTicketInfo szTicketID
    Set oSncom = Nothing
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub tbSeatManage_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo ErrHandle
    Dim nResult As Integer
    Select Case Button.Key
        Case "AddSeat"
            AddSeat
            mbEdited = True
        Case "DelSeat"
            If lvSeat.SelectedItem Is Nothing Then Exit Sub
            nResult = MsgBox("是否删除选择的座位?", vbQuestion + vbYesNo + vbDefaultButton2, "问题")
            If nResult = vbYes Then
                DeleteSeat
                mbEdited = True
            End If
        Case "EditSeat"
            If lvSeat.SelectedItem Is Nothing Then Exit Sub
            EditSeat
            mbEdited = True
        Case "ReserveSeat"
            If lvSeat.SelectedItem Is Nothing Then Exit Sub
            ReserveSeat
            mbEdited = True
        Case "UnReserveSeat"
            If lvSeat.SelectedItem Is Nothing Then Exit Sub
            UnReserveSeat
            mbEdited = True
            
        Case Else
            Exit Sub
    End Select
    RefreshBus
    ShowSBInfo
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub AddSeat()
    frmSeat.m_SeatType = maszSeatType
    Set frmSeat.m_oReBus = m_oReBus
    On Error GoTo ErrHandle
    If lvSeat.ListItems.Count <> 0 Then
       frmSeat.m_szSeatNo = lvSeat.ListItems(lvSeat.ListItems.Count)
    End If
    frmSeat.efrmSeat = eSeat.AddSeat
    frmSeat.Show vbModal
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub tmStart_Timer()
On Error GoTo ErrHandle
    tmStart.Enabled = False
    SetBusy
    If RefreshBus Then
    FullSeat
    End If
    SetNormal
Exit Sub
ErrHandle:
    tmStart.Enabled = False
    SetNormal
    ShowErrorMsg
End Sub
'补充车次信息
Private Function RefreshBus() As Boolean
    On Error GoTo ErrHandle
    Dim liTemp As ListItem
    Dim tSeat() As TSeatInfoEx
    Dim i As Integer, nCount As Integer
    lblBusID.Caption = m_szBusID
    m_oReBus.Identify m_szBusID, m_dtEnvDate
    If m_oReBus.Vehicle = "" Then Exit Function
    If m_oReBus.BusType = TP_ScrollBus Then
        lblStartupTime.Caption = m_oReBus.ScrollBusCheckTime & "分钟一班"
    Else
        lblStartupTime.Caption = Format(m_oReBus.StartUpTime, "hh:mm")
    End If
    lblReserve.Caption = "预留座位数:" & m_oReBus.ReserveSeatCount
    lblTotal.Caption = "总座位数:" & m_oReBus.TotalSeat
    lblSell.Caption = "已售座位数:" & m_oReBus.SaledSeatCount
    RefreshBus = True
    Exit Function
ErrHandle:
    ShowErrorMsg
End Function

