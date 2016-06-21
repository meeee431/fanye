VERSION 5.00
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmRESeat 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "环境--座位管理"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7755
   HelpContextID   =   2006401
   Icon            =   "frmRESeat.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   7755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin RTComctl3.CoolButton cmdCancelBook 
      Height          =   315
      Left            =   6360
      TabIndex        =   17
      Top             =   2010
      Width           =   1245
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "取消预定(&B)"
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
      MICON           =   "frmRESeat.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdTicket 
      Height          =   315
      Left            =   6375
      TabIndex        =   16
      Top             =   2385
      Width           =   1245
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "车票信息(&T)"
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
      MICON           =   "frmRESeat.frx":0028
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
      Left            =   5445
      Top             =   630
   End
   Begin RTComctl3.CoolButton cmdHelp 
      Height          =   315
      Left            =   6375
      TabIndex        =   10
      Top             =   2760
      Width           =   1245
      _ExtentX        =   0
      _ExtentY        =   0
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
      MICON           =   "frmRESeat.frx":0044
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
      Left            =   3075
      Top             =   2970
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRESeat.frx":0060
            Key             =   "ResSeat"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRESeat.frx":01BC
            Key             =   "SellSeat"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRESeat.frx":0318
            Key             =   "Normal"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRESeat.frx":0474
            Key             =   "Book"
         EndProperty
      EndProperty
   End
   Begin RTComctl3.CoolButton cmdCancelReserve 
      Height          =   315
      Left            =   6375
      TabIndex        =   9
      Top             =   1635
      Width           =   1245
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "取消预留(&E)"
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
      MICON           =   "frmRESeat.frx":12C8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdReserve 
      Height          =   315
      Left            =   6375
      TabIndex        =   8
      Top             =   1260
      Width           =   1245
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "预留(&R)"
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
      MICON           =   "frmRESeat.frx":12E4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdDelete 
      Height          =   315
      Left            =   6375
      TabIndex        =   7
      Top             =   885
      Width           =   1245
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "删除(&D)"
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
      MICON           =   "frmRESeat.frx":1300
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdAddNew 
      Height          =   315
      Left            =   6375
      TabIndex        =   6
      Top             =   510
      Width           =   1245
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "新增(&A)"
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
      MICON           =   "frmRESeat.frx":131C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ListView lvSeat 
      Height          =   3015
      Left            =   120
      TabIndex        =   4
      Top             =   1350
      Width           =   6120
      _ExtentX        =   10795
      _ExtentY        =   5318
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
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "座位号"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "状态"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "票号"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "到站"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "票价"
         Object.Width           =   1764
      EndProperty
   End
   Begin RTComctl3.CoolButton cmdOk 
      Height          =   315
      Left            =   6375
      TabIndex        =   5
      Top             =   135
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
      MICON           =   "frmRESeat.frx":1338
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.TextButtonBox txtBusID 
      Height          =   300
      Left            =   1920
      TabIndex        =   15
      Top             =   135
      Width           =   1425
      _ExtentX        =   2514
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
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   1170
      X2              =   6195
      Y1              =   1185
      Y2              =   1185
   End
   Begin VB.Line Line1 
      X1              =   1170
      X2              =   6195
      Y1              =   1170
      Y2              =   1170
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmRESeat.frx":1354
      Stretch         =   -1  'True
      Top             =   225
      Width           =   480
   End
   Begin VB.Label lblBusDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车次日期:"
      Height          =   180
      Left            =   3795
      TabIndex        =   2
      Top             =   180
      Width           =   810
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车次座位(&S)"
      Height          =   180
      Left            =   120
      TabIndex        =   3
      Top             =   1095
      Width           =   990
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "总座位数:"
      Height          =   180
      Left            =   3795
      TabIndex        =   1
      Top             =   495
      Width           =   810
   End
   Begin VB.Label lblReserve 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "预留座位数:"
      Height          =   180
      Left            =   3795
      TabIndex        =   14
      Top             =   825
      Width           =   990
   End
   Begin VB.Label lblSell 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "已售座位数:"
      Height          =   180
      Left            =   825
      TabIndex        =   13
      Top             =   825
      Width           =   990
   End
   Begin VB.Label lblStartupTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   180
      Left            =   1890
      TabIndex        =   12
      Top             =   525
      Width           =   90
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "发车时间:"
      Height          =   180
      Left            =   825
      TabIndex        =   11
      Top             =   525
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车次代码(&I):"
      Height          =   180
      Left            =   825
      TabIndex        =   0
      Top             =   180
      Width           =   1080
   End
End
Attribute VB_Name = "frmRESeat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************
'                   TOP GROUP INC.
'* Copyright(C)1999 TOP GROUP INC.
'*
'* All rights reserved.No part of this program or publication
'* may be reproduced,transmitted,transcribed,stored in a
'* retrieval system,or translated intoany language or compute
'* language,in any form or by any means,electronic,mechanical,
'* magnetic,optical,chemical,biological,or otherwise,without
'* the prior written permission.
'*********************************************************
'
'**********************************************************
'* Source File Name:frmRESeat.frm
'* Project Name:StationNet 2.0
'* Engineer:魏宏旭
'* Data Generated:1999/8/28
'* Last Revision Date:1999/9/3
'* Brief Description:座位管理
'* Relational Document:UI_BS_SM_029.DOC
'**********************************************************
Public m_szBusID As String
Public m_dtBusDate As Date
Private m_oREBus As New REBus

Private Sub cmdAddNew_Click()
    Set frmSeat.m_oREBus = m_oREBus
    frmSeat.efrmSeat = AddSeat
    frmSeat.Show vbModal
    FullBus
End Sub

Private Sub cmdCancelBook_Click()
    Me.MousePointer = vbHourglass
    UnBookSeat
    FullBus
    Me.MousePointer = vbDefault
End Sub

Private Sub cmdCancelReserve_Click()
    Me.MousePointer = vbHourglass
    UnReserveSeat
    FullBus
    'ShowTBInfo
    Me.MousePointer = vbDefault
End Sub

Private Sub cmdDelete_Click()
Dim vbMsg As VbMsgBoxResult
vbMsg = MsgBox("是否删除选择的座位?", vbQuestion + vbYesNo + vbDefaultButton2, "环境")
If vbMsg = vbYes Then
    Me.MousePointer = vbHourglass
    DeleteSeat
    FullBus
    'ShowTBInfo
    Me.MousePointer = vbDefault
End If
End Sub

Private Sub cmdHelp_Click()
'DisplayHelp Me
End Sub

Private Sub cmdOk_Click()
    Unload Me
End Sub

Private Sub cmdReserve_Click()
    Me.MousePointer = vbHourglass
    ReserveSeat
    FullBus
    'ShowTBInfo
    Me.MousePointer = vbDefault
End Sub

Private Sub cmdTicket_Click()
    On Error Resume Next
    lvSeat_DblClick
End Sub

Private Sub Form_Load()
Dim tSeat() As TSeatInfo
Dim nCount As Integer
Dim ltTemp As ListItem
Dim i As Integer
m_oREBus.Init m_oActiveUser
If m_szBusID = "" Then m_dtBusDate = Date
lblBusDate.Caption = "车次日期: " & Format(m_dtBusDate, "YYYY-MM-DD")
Me.Caption = "环境--车次座位" & "[" & Format(m_dtBusDate, "YYYY-MM-DD") & "]"
'SetListViewWidth lvSeat
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
       Case vbKeyEscape
       Unload Me
End Select
End Sub

Public Sub FullSeat(Optional RefreshSeat As Boolean)
Dim tSeat() As TSeatInfo
Dim nCount As Integer
Dim ltTemp As ListItem
Dim i As Integer
On Error GoTo here
lvSeat.ListItems.Clear
If m_szBusID <> "" Then
    'ShowTBInfo "获得座位信息..."
    m_oREBus.Identify m_szBusID, m_dtBusDate
    tSeat = m_oREBus.GetSeatInfo
    nCount = ArrayLength(tSeat)
    If nCount <= 0 Then Exit Sub
    'ShowTBInfo , nCount, , True
    For i = 1 To nCount
        'ShowTBInfo "填写座位" & tSeat(i).szSeatNo, , i
        If tSeat(i).szSeatStatus = ST_SeatReserved Then
            Set ltTemp = lvSeat.ListItems.Add(, , tSeat(i).szSeatNo, , "ResSeat")
            ltTemp.ListSubItems.Add , , "预留"
        End If
        If tSeat(i).szSeatStatus = ST_SeatBooked Then
            Set ltTemp = lvSeat.ListItems.Add(, , tSeat(i).szSeatNo, , "Book")
            ltTemp.ListSubItems.Add , , "预定"
        End If
        If tSeat(i).szSeatStatus = ST_SeatCanSell Then
            Set ltTemp = lvSeat.ListItems.Add(, , tSeat(i).szSeatNo, , "Normal")
            ltTemp.ListSubItems.Add , , "可售"
        End If
        If tSeat(i).szSeatStatus = ST_SeatSold Then
            Set ltTemp = lvSeat.ListItems.Add(, , tSeat(i).szSeatNo, , "SellSeat")
            ltTemp.ListSubItems.Add , , "已售"
            ltTemp.ListSubItems.Add , , tSeat(i).szTicketNo
            ltTemp.ListSubItems.Add , , tSeat(i).szDestName
            ltTemp.ListSubItems.Add , , tSeat(i).szTicketPrice
        End If
        
    Next
End If
'ShowTBInfo
Exit Sub
here:
    'ShowTBInfo
    ShowErrorMsg
End Sub

Public Sub DeleteSeat()
On Error GoTo err
Dim ErrStr As String
    Dim i As Integer
    Dim oItem As ListItem
    For i = lvSeat.ListItems.Count To 1 Step -1
        Set oItem = lvSeat.ListItems(i)
        If oItem.Selected Then
     '       ShowTBInfo "删除座位" & oItem.Text
            DeleteSeatEx oItem
        End If
    Next
    If ErrStr <> "" Then MsgBox ErrStr
    Exit Sub
err:
    ErrStr = ErrStr & "[座位" & oItem.Text & "]" & err.Description & vbCrLf
    Resume Next
End Sub

'删除座位
Private Sub DeleteSeatEx(oInListItem As ListItem)
m_oREBus.DeleteSeat 1, Format(oInListItem.Text, "00")
lvSeat.ListItems.Remove oInListItem.Index
End Sub

Private Sub ReserveSeat()
On Error GoTo err
Dim ErrStr As String
    Dim oItem As ListItem
    For Each oItem In lvSeat.ListItems
        If oItem.Selected Then
             'ShowTBInfo "预留座位" & oItem.Text
             ReserveSeatEx oItem
        End If
    Next
    If ErrStr <> "" Then MsgBox ErrStr
    Exit Sub
err:
    ErrStr = ErrStr & "[座位" & oItem.Text & "]" & err.Description & vbCrLf
    Resume Next
End Sub

Private Sub ReserveSeatEx(oInListItem As ListItem)
m_oREBus.ReserveSeat oInListItem.Text
oInListItem.ListSubItems(1).Text = "预留"
oInListItem.SmallIcon = "ResSeat"
End Sub

Private Sub UnReserveSeat()
On Error GoTo err
Dim ErrStr As String
    Dim oItem As ListItem
    For Each oItem In lvSeat.ListItems
        If oItem.Selected Then
      '      ShowTBInfo "取消预留" & oItem.Text
            UnReserveSeatEx oItem
        End If
    Next
    If ErrStr <> "" Then MsgBox ErrStr
    Exit Sub
err:
    ErrStr = ErrStr & "[座位" & oItem.Text & "]" & err.Description & vbCrLf
    Resume Next
End Sub

Private Sub UnReserveSeatEx(oInListItem As ListItem)
m_oREBus.UnReserveSeat oInListItem.Text
oInListItem.ListSubItems(1).Text = "可售"
oInListItem.SmallIcon = "Normal"
End Sub

Private Sub Form_Unload(Cancel As Integer)
'SaveListViewWidth lvSeat
End Sub

Private Sub lvSeat_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error Resume Next
Static m_nUpColumn As Integer
lvSeat.SortKey = ColumnHeader.Index - 1
If m_nUpColumn = ColumnHeader.Index - 1 Then
    lvSeat.SortOrder = lvwDescending
    m_nUpColumn = ColumnHeader.Index
Else
    lvSeat.SortOrder = lvwAscending
    m_nUpColumn = ColumnHeader.Index - 1
End If
lvSeat.Sorted = True
End Sub

Private Sub lvSeat_DblClick()
'    Dim oSncom As New CheckSysApp
'    Dim szTicketID As String
'On Error GoTo here
'    szTicketID = GetLString(lvSeat.SelectedItem.ListSubItems(2).Text)
'    If szTicketID = "" Then Exit Sub
'    oSncom.ShowTicketInfo m_oActiveUser, szTicketID
'    Set oSncom = Nothing
'Exit Sub
'here:
End Sub

Private Sub lvSeat_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo here
    If lvSeat.SelectedItem.ListSubItems(2).Text <> "" Then
        cmdTicket.Enabled = True
    Else
        cmdTicket.Enabled = False
    End If
Exit Sub
here:
If err.Number = 35600 Then cmdTicket.Enabled = False
End Sub

Private Sub tmStart_Timer()
On Error GoTo here
    Me.MousePointer = vbHourglass
    'SetListViewWidth lvSeat
    FullBus
    FullSeat
    tmStart.Enabled = False
Exit Sub
here:
    tmStart.Enabled = False
    Me.MousePointer = vbDefault
    ShowErrorMsg
End Sub

Private Sub txtBusID_Click()
    Dim oShell As New CommDialog
    Dim szaTemp() As String
On Error GoTo here
    oShell.Init m_oActiveUser
    szaTemp = oShell.SelectREBus(m_dtBusDate, False)
    Set oShell = Nothing
    If ArrayLength(szaTemp) = 0 Then Exit Sub
    txtBusID.Text = szaTemp(1, 1)
    m_szBusID = txtBusID.Text
    Me.MousePointer = vbHourglass
    FullBus
    FullSeat
Exit Sub
here:
    Me.MousePointer = vbDefault
    ShowErrorMsg
End Sub

Private Sub txtBusID_KeyPress(KeyAscii As Integer)
On Error GoTo here
Select Case KeyAscii
       Case vbKeyReturn
       m_dtBusDate = m_dtBusDate
       m_szBusID = txtBusID.Text
       Me.MousePointer = vbHourglass
       FullBus
       FullSeat
End Select
Exit Sub
here:
    Me.MousePointer = vbDefault
    ShowErrorMsg
End Sub

Private Sub FullBus()
Dim ltTemp As ListItem
Dim tSeat() As TSeatInfo
Dim i As Integer, nCount As Integer
Me.MousePointer = vbHourglass
txtBusID.Text = m_szBusID
If m_szBusID <> "" Then
    m_oREBus.Identify m_szBusID, m_dtBusDate
    lblStartupTime.Caption = Format(m_oREBus.StartupTime, "HH:MM:SS")
    lblReserve.Caption = "预留座位数:" & m_oREBus.ReserveSeatCount
    lblTotal.Caption = "总座位数: " & m_oREBus.TotalSeat
    lblSell.Caption = "已售座位数: " & m_oREBus.SaledSeatCount
End If
cmdAddNew.Enabled = True
cmdDelete.Enabled = True
cmdReserve.Enabled = True
cmdCancelBook.Enabled = True
cmdCancelReserve.Enabled = True
cmdTicket.Enabled = False
Me.MousePointer = vbDefault
End Sub

Private Sub UnBookSeat()
Dim aszTemp(1 To 1) As String
On Error GoTo err
Dim ErrStr As String
    Dim oItem As ListItem
    For Each oItem In lvSeat.ListItems
        If oItem.Selected Then
            aszTemp(1) = oItem.Text
            m_oBook.UnBook m_szBusID, m_dtBusDate, aszTemp
            oItem.SmallIcon = "Normal"
            oItem.ListSubItems(1).Text = "可售"
        End If
    Next
    If ErrStr <> "" Then MsgBox ErrStr
    Exit Sub
err:
    ErrStr = ErrStr & "[座位" & oItem.Text & "]" & err.Description & vbCrLf
    Resume Next
End Sub
