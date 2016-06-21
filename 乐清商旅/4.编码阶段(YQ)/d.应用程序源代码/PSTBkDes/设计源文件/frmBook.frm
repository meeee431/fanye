VERSION 5.00
Object = "{61C3A787-42A5-4F09-9AD8-C9DE75BAD364}#1.0#0"; "STSeatpad.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Object = "{BBF95DAA-F9CB-4CA9-A673-E0E9E0193752}#1.0#0"; "STSellCtl.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBook 
   BackColor       =   &H00E0E0E0&
   Caption         =   "预定窗口"
   ClientHeight    =   7170
   ClientLeft      =   2460
   ClientTop       =   2445
   ClientWidth     =   10215
   Icon            =   "frmBook.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7170
   ScaleWidth      =   10215
   StartUpPosition =   1  '所有者中心
   Begin RTComctl3.CoolButton cmdBook 
      Height          =   585
      Left            =   7380
      TabIndex        =   17
      Top             =   6390
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   1032
      BTYPE           =   3
      TX              =   "预定(F2)"
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
      MICON           =   "frmBook.frx":0E42
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.TextButtonBox txtTel 
      Height          =   315
      Left            =   3690
      TabIndex        =   12
      Top             =   5910
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   556
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
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   9570
      Top             =   3360
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   75
      Left            =   90
      TabIndex        =   18
      Top             =   5250
      Width           =   10155
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   75
      Left            =   60
      TabIndex        =   22
      Top             =   435
      Width           =   10095
   End
   Begin VB.CommandButton cmdFocusBook 
      Caption         =   "Command1"
      Height          =   255
      Left            =   6300
      TabIndex        =   19
      Top             =   -2000
      Width           =   375
   End
   Begin VB.TextBox txtSeat 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3675
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   5415
      Width           =   2925
   End
   Begin VB.TextBox txtBookMan 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7770
      TabIndex        =   10
      Top             =   5415
      Width           =   2265
   End
   Begin VB.TextBox txtEmail 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7770
      TabIndex        =   14
      Top             =   5865
      Width           =   2265
   End
   Begin VB.TextBox txtAnnotation 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3675
      TabIndex        =   16
      Top             =   6345
      Width           =   2925
   End
   Begin MSComCtl2.DTPicker dtpBusDate 
      Height          =   375
      Left            =   4080
      TabIndex        =   21
      Top             =   60
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   22937600
      CurrentDate     =   36692
   End
   Begin STSellCtl.ucUpDownText txtPreDate 
      Height          =   360
      Left            =   1635
      TabIndex        =   1
      Top             =   60
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   635
      SelectOnEntry   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Max             =   100
      Value           =   "0"
   End
   Begin MSComctlLib.ListView lvBus 
      Height          =   2520
      Left            =   2640
      TabIndex        =   4
      Top             =   840
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   4445
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "BusID"
         Text            =   "车次"
         Object.Width           =   1676
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "OffTime"
         Text            =   "发车时间"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "EndStation"
         Text            =   "终到站"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "SeatCount"
         Text            =   "座位"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "BusModel"
         Text            =   "车型"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "FullPrice"
         Text            =   "全价"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Key             =   "HalfPrice"
         Text            =   "半价"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Key             =   "mbookc"
         Text            =   "此站张数"
         Object.Width           =   2540
      EndProperty
   End
   Begin STSellCtl.ucSuperCombo scmbStation 
      Height          =   4275
      Left            =   120
      TabIndex        =   3
      Top             =   825
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   7541
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin STSeatPad.SeatPad spSeat 
      Height          =   1515
      Left            =   2640
      TabIndex        =   6
      Top             =   3720
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   2672
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      Caption         =   "SeatPad1"
      GridNum         =   0
      RowGrids        =   17
   End
   Begin MSComctlLib.ListView lvSellStation 
      Height          =   1605
      Left            =   30
      TabIndex        =   31
      Top             =   5550
      Width           =   2430
      _ExtentX        =   4286
      _ExtentY        =   2831
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "上车站"
         Object.Width           =   1942
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   " 时间"
         Object.Width           =   1696
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "上车站代码"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "检票门"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label lblsellstation 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "上车站(&O):"
      Height          =   180
      Left            =   30
      TabIndex        =   32
      Top             =   5370
      Width           =   900
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "剩余特票数:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   6060
      TabIndex        =   30
      Top             =   3420
      Width           =   1320
   End
   Begin VB.Label lblLeftSpecial 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   7590
      TabIndex        =   29
      Top             =   3420
      Width           =   120
   End
   Begin VB.Label lblLeftHalf 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   5520
      TabIndex        =   28
      Top             =   3420
      Width           =   120
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "剩余半票数:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   4020
      TabIndex        =   27
      Top             =   3420
      Width           =   1320
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车次(&B):"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2685
      TabIndex        =   26
      Top             =   540
      Width           =   960
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "红色表示预定；蓝色表示已出售的特票；棕色表示已出售的半票；黄色表示已出售"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   3690
      TabIndex        =   25
      Top             =   570
      Width           =   6480
   End
   Begin VB.Label flblChinaDate 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "丙戌年(狗)五月初一"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6450
      TabIndex        =   24
      Top             =   120
      Width           =   1830
   End
   Begin VB.Label flblSellWeek 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "星期六"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8520
      TabIndex        =   23
      Top             =   120
      Width           =   630
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "座位(&I):"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2670
      TabIndex        =   5
      Top             =   3420
      Width           =   960
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "到站(&S):"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   2
      Top             =   540
      Width           =   960
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车次日期(&D):"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2610
      TabIndex        =   20
      Top             =   120
      Width           =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车次日期(&P):"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1440
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "姓名(&M):"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6750
      TabIndex        =   9
      Top             =   5475
      Width           =   960
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "电话(&L):"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2580
      TabIndex        =   11
      Top             =   5925
      Width           =   960
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "地址(&E):"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6750
      TabIndex        =   13
      Top             =   5925
      Width           =   960
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "备注(&A):"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2580
      TabIndex        =   15
      Top             =   6375
      Width           =   960
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "座位(&T):"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2580
      TabIndex        =   7
      Top             =   5475
      Width           =   960
   End
End
Attribute VB_Name = "frmBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_oShell As New SellTicketClient
Dim m_nCanSellDay As Integer


Private Sub cmdBook_Click()
    Dim i As Integer, nSelectSeat As Integer
    Dim aszSeat() As String
    Dim szBookNumber As String
    Dim frmQueryTemp As Form
    Dim szTemp As String
    
    On Error GoTo Error_Handle
    If txtSeat.Text <> "" And scmbStation.BoundText <> "" And Not lvBus.SelectedItem Is Nothing Then
        nSelectSeat = 0
        ReDim aszSeat(1 To spSeat.GridNum)
        For i = 1 To spSeat.GridNum
            If spSeat.PadGrids.Item(i).Value = vbChecked Then
                nSelectSeat = nSelectSeat + 1
                aszSeat(nSelectSeat) = spSeat.PadGrids.Item(i).Caption
            End If
        Next
        
        
        ReDim Preserve aszSeat(1 To nSelectSeat)
        
        szBookNumber = m_oBook.Book(lvBus.SelectedItem.Text, dtpBusDate.Value, "", scmbStation.BoundText, aszSeat, txtBookMan.Text, txtTel.Text, txtEmail.Text, txtAnnotation.Text)

        For i = 1 To spSeat.GridNum
            If spSeat.PadGrids.Item(i).Value = vbChecked Then
                spSeat.PadGrids.Item(i).Value = vbUnchecked
                spSeat.PadGrids.Item(i).Enabled = False
            End If
        Next
        
        spSeat.Refresh

        For Each frmQueryTemp In Forms
            If TypeName(frmQueryTemp) = "frmQuery" Then
                If frmQueryTemp.dtpBusDate = dtpBusDate.Value Then
                    frmQueryTemp.FillBookInfo
                End If
            End If
        Next
        szTemp = "预定号为:" & szBookNumber
        ShowMsg szTemp
        scmbStation.SetFocus
'        txtSeat.Text = ""
'        txtBookMan.Text = ""
'        txtTel.Text = ""
'        txtEmail.Text = ""
'        txtAnnotation.Text = ""
        
    End If
    Exit Sub
Error_Handle:
    ShowErrorMsg
End Sub

Private Sub cmdFocusBook_GotFocus()
    scmbStation.SetFocus
End Sub

Private Sub dtpBusDate_Change()
    Dim nTemp  As Integer
    nTemp = DateDiff("d", m_oParam.NowDate, dtpBusDate.Value)
    If nTemp <> txtPreDate.Value Then
        txtPreDate.Value = nTemp
    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        cmdBook_Click
    End If
End Sub
Private Sub InitForm()

        txtSeat.Text = ""
        txtBookMan.Text = ""
        txtTel.Text = ""
        txtEmail.Text = ""
        txtAnnotation.Text = ""
        lblLeftHalf.Caption = 0
        lblLeftSpecial.Caption = 0
        lvSellStation.ListItems.Clear
'        txtPreDate.Value = 0
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And Not Me.ActiveControl Is scmbStation Then
        SendKeys "{TAB}"
    ElseIf KeyAscii = vbKeyEscape Then
'        Unload Me
        InitForm
        txtPreDate.Value = 0
        scmbStation.SetFocus
        
    End If
End Sub

Private Sub Form_Load()
    InitListView lvBus
    InitForm
    
    m_oShell.Init m_oActiveUser
    txtPreDate.Value = 0
    flblChinaDate.Caption = ResolveDisplay(GetChinaDate(ToDBDate(dtpBusDate.Value)))
    flblSellWeek.Caption = ResolveDisplayEx(GetChinaDate(ToDBDate(dtpBusDate.Value)))
    
    FillStation
    SortListView lvBus, 2
    
    Dim oSysMan As New User
    oSysMan.Init m_oActiveUser
    oSysMan.Identify m_oActiveUser.UserID
    m_nCanSellDay = oSysMan.CanSellDay

    EnableBook
End Sub

Private Sub FillStation()
    Dim rsTemp As Recordset
    Set rsTemp = m_oShell.GetAllStationRs()
    With scmbStation
        Set .RowSource = rsTemp
        .BoundField = "station_id"
        .ListFields = "station_input_code:3,station_name:4"
        .AppendWithFields "station_id:9,station_name"
    End With
End Sub

Private Sub FillBus()
    lvBus.ListItems.Clear
    If scmbStation.BoundText <> "" Then
        Dim rsTemp As Recordset
        Dim lRecordsetCount As Long, i As Long
        Dim liTemp As ListItem
        Dim j As Integer
        Dim bFound As Boolean
        Set rsTemp = m_oShell.GetBusRs(dtpBusDate.Value, scmbStation.BoundText)
        lRecordsetCount = rsTemp.RecordCount
        If lRecordsetCount > 0 Then
            For i = 1 To lRecordsetCount
                If rsTemp!Status = ST_BusStopped Or rsTemp!Status = ST_BusMergeStopped Or rsTemp!Status = ST_BusSlitpStop Then
                    '停班的车次不显示
                    'Debug.Print 1
                Else
                
                        
'                If rsTemp!bus_type = TP_RegularBus And rsTemp!Status = ST_BusNormal Then
                    bFound = False
                    For j = 1 To lvBus.ListItems.Count
                        If FormatDbValue(rsTemp!bus_id) = lvBus.ListItems(j).Text Then
                            bFound = True
                            Exit For
                        End If
                    Next j
                    
                    If Not bFound Then
                        Set liTemp = lvBus.ListItems.Add(, "A" & Trim(rsTemp!bus_id), Trim(rsTemp!bus_id))
                        liTemp.ListSubItems.Add , , Format(rsTemp!bus_start_time, "HH:MM") '车次发车时间
                        liTemp.ListSubItems.Add , , RTrim(rsTemp!end_station_name) '车次终到站名
                        liTemp.ListSubItems.Add , , rsTemp!sale_seat_quantity '可售座位数
                        liTemp.ListSubItems.Add , , rsTemp!vehicle_type_name '车次名称
                        liTemp.ListSubItems.Add , , rsTemp!full_price '全票价
                        liTemp.ListSubItems.Add , , rsTemp!half_price '半票价
'                        If rsTemp!sale_ticket_quantity >= 0 Then
'                            liTemp.ListSubItems.Add , , rsTemp!sale_ticket_quantity - rsTemp!book_count
'                        Else
                            liTemp.ListSubItems.Add , , "不限"
'                        End If
                    End If
'                End If
                End If
                rsTemp.MoveNext
            Next
            If lvBus.ListItems.Count > 0 Then
                lvBus.ListItems.Item(1).Selected = True
            End If
        End If
    End If
    FillSeat
    
    If lvBus.ListItems.Count > 0 Then
      RefreshSellStation
   Else
      lvSellStation.ListItems.Clear
   End If
End Sub

Private Sub FillSeat()
    Dim rsTemp As Recordset
    Dim nSeatCount As Integer, i As Integer
    Dim oPad As PadGrid
    Dim bSpecialTicketTypeRatio As Double
    Dim bHalfTicketTypeRatio As Double
    Dim nSpecialSold As Integer  '已售特票数
    Dim nHalfSold As Integer    '已售半票数
    
    nSpecialSold = 0
    nHalfSold = 0
    bSpecialTicketTypeRatio = m_oParam.SpecialTicketTypeRatio
    bHalfTicketTypeRatio = m_oParam.HalfTicketTypeRatio
    
    If Not lvBus.SelectedItem Is Nothing Then
        spSeat.Enabled = True
        
        Set rsTemp = m_oShell.GetSeatRs(CDate(dtpBusDate.Value), lvBus.SelectedItem.Text)

        nSeatCount = rsTemp.RecordCount
        spSeat.GridNum = nSeatCount
        If nSeatCount > 0 Then rsTemp.MoveFirst
        
        For i = 1 To spSeat.GridNum
            Set oPad = spSeat.PadGrids.Item(i)
            oPad.Caption = rsTemp!seat_no
            oPad.BackColor = &HE0E0E0
            oPad.Enabled = True

            Select Case rsTemp!Status
            Case ST_SeatCanSell
                oPad.Value = vbUnchecked
            Case ST_SeatBooked
                oPad.Value = vbUnchecked
                oPad.BackColor = RGB(255, 0, 0)
            Case ST_SeatProjectBooked
                oPad.Value = vbUnchecked
                oPad.BackColor = RGB(0, 255, 0)
            Case Else
                If rsTemp!ticket_type = TP_PreferentialTicket2 Then '已售特票蓝色显示
                    oPad.BackColor = vbBlue
                    nSpecialSold = nSpecialSold + 1
                ElseIf rsTemp!ticket_type = TP_HalfPrice Then '已售半票棕色显示
                    oPad.BackColor = RGB(251, 149, 3)
                    nHalfSold = nHalfSold + 1
                Else '已售其他票种黄色显示
                    oPad.BackColor = vbYellow
                End If
                oPad.Enabled = False
            End Select
            
            rsTemp.MoveNext
        Next
        
        lblLeftHalf.Caption = Int(bHalfTicketTypeRatio * nSeatCount / 100) - nHalfSold
        lblLeftSpecial.Caption = Int(bSpecialTicketTypeRatio * nSeatCount / 100) - nSpecialSold
            
    Else
        spSeat.Enabled = False
    End If
    spSeat.Refresh
    txtSeat.Text = ""
    EnableBook
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveListView lvBus
End Sub

Private Sub lvBus_GotFocus()
    If lvBus.ListItems.Count = 0 Then scmbStation.SetFocus
End Sub

Private Sub lvBus_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Timer1.Enabled = False
    Timer1.Enabled = True
    RefreshSellStation
    'FillSeat
End Sub

Private Sub RefreshSellStation()
  Dim i As Integer
  Dim lvS As ListItem

  On Error GoTo err:
    lvSellStation.Sorted = False
    lvSellStation.ListItems.Clear
    lvSellStation.Refresh
    If lvBus.ListItems.Count = 0 Then
       Exit Sub
    End If
    
    Dim rsTemp As New Recordset
    Set rsTemp = m_oBook.GetEnvBusAllotInfo(CDate(ToDBDate(dtpBusDate.Value)), lvBus.SelectedItem.Text)
    If rsTemp.RecordCount <= 0 Then Exit Sub
    
    rsTemp.MoveFirst
    For i = 1 To rsTemp.RecordCount
        Set lvS = lvSellStation.ListItems.Add(, , Trim(rsTemp!sell_station_name))
            lvS.SubItems(1) = Trim(Format(rsTemp!bus_start_time, "hh:mm"))
            lvS.SubItems(2) = Trim(rsTemp!sell_station_id)
            lvS.SubItems(3) = Trim(rsTemp!check_gate_id)
            rsTemp.MoveNext
    Next i
    
    For i = 1 To lvSellStation.ListItems.Count
        If UCase(Trim(m_oActiveUser.SellStationID)) = UCase(Trim(lvSellStation.ListItems(i).SubItems(2))) Then

            lvSellStation.ListItems(i).Selected = True
            lvSellStation.ListItems(i).EnsureVisible
            Exit For
        End If
    Next i
    
    Set rsTemp = Nothing
    Exit Sub
    
err:
    Set rsTemp = Nothing
    MsgBox err.Description
End Sub

Private Sub scmbStation_GotFocus()
''    lvBus.ListItems.Clear
    'spSeat.GridNum = 0
    
    'txtSeat.Text = ""
    EnableBook
End Sub

Private Sub scmbStation_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyLeft
            KeyCode = 0
            If Val(txtPreDate.Value) > 0 Then
                txtPreDate.Value = Val(txtPreDate.Value) - 1
            End If
'            m_bNotRefresh = False
        Case vbKeyRight
            KeyCode = 0
'            If Val(txtPreDate.Value) < m_nCanSellDay Then
            
                txtPreDate.Value = Val(txtPreDate.Value) + 1
'            End If
'            m_bNotRefresh = False
        Case Else
'            m_bNotRefresh = False
    End Select
'    If m_bPreClear Then
'        lvPreSell.ListItems.Clear
'        flblTotalPrice.Caption = 0#
'        txtReceivedMoney.Text = ""
'        flblRestMoney.Caption = ""
'        m_bPreClear = False
'    End If
End Sub

Private Sub scmbStation_LostFocus()
    InitForm
    FillBus
End Sub

Private Sub spSeat_GridClick(Index As Integer)
    Dim i As Integer
    Dim szTemp As String
    For i = 1 To spSeat.GridNum
        If spSeat.PadGrids.Item(i).Value = vbChecked Then
            If szTemp = "" Then
                szTemp = spSeat.PadGrids.Item(i).Caption
            Else
                szTemp = szTemp & "," & spSeat.PadGrids.Item(i).Caption
            End If
        End If
    Next
    txtSeat.Text = szTemp
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    FillSeat
End Sub

Private Sub txtPreDate_Change()
    Dim dtDate As Date
    
    If Val(txtPreDate.Value) > m_nCanSellDay Then txtPreDate.Value = m_nCanSellDay
    
    dtDate = DateAdd("d", txtPreDate.Value, m_oParam.NowDate)
    If dtDate <> dtpBusDate.Value Then
        dtpBusDate.Value = dtDate
    End If
    lvBus.ListItems.Clear
    spSeat.GridNum = 0
    flblChinaDate.Caption = ResolveDisplay(GetChinaDate(ToDBDate(dtpBusDate.Value)))
    flblSellWeek.Caption = ResolveDisplayEx(GetChinaDate(ToDBDate(dtpBusDate.Value)))
End Sub

Private Sub EnableBook()
    If scmbStation.BoundText <> "" And Not lvBus.SelectedItem Is Nothing And txtSeat.Text <> "" Then
        cmdBook.Enabled = True
    Else
        cmdBook.Enabled = False
    End If
End Sub

Private Sub txtSeat_Change()
    EnableBook
End Sub
