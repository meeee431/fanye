VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BBF95DAA-F9CB-4CA9-A673-E0E9E0193752}#1.0#0"; "STSellCtl.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmCancelTicket 
   BackColor       =   &H8000000C&
   Caption         =   "废票"
   ClientHeight    =   7065
   ClientLeft      =   -120
   ClientTop       =   1950
   ClientWidth     =   9900
   HelpContextID   =   4000220
   Icon            =   "frmCancelTicket.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7065
   ScaleWidth      =   9900
   WindowState     =   2  'Maximized
   Begin VB.Frame fraOutLine 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   7005
      Left            =   360
      TabIndex        =   14
      Top             =   180
      Width           =   9900
      Begin VB.TextBox txtEndTicketNo 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   1650
         MaxLength       =   10
         TabIndex        =   3
         Top             =   690
         Width           =   1950
      End
      Begin VB.ComboBox cboSeat 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8040
         TabIndex        =   11
         Text            =   "所有座位"
         Top             =   180
         Width           =   1305
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   400
         Left            =   6360
         TabIndex        =   10
         Top             =   180
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   714
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
         Format          =   61669377
         CurrentDate     =   37877
      End
      Begin VB.TextBox txtBusID 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   5400
         TabIndex        =   9
         Top             =   180
         Width           =   915
      End
      Begin VB.TextBox txtTicketNo 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   1650
         MaxLength       =   10
         TabIndex        =   1
         Top             =   180
         Width           =   1950
      End
      Begin RTComctl3.CoolButton cmdResumeCancel 
         Height          =   435
         Left            =   7890
         TabIndex        =   5
         Top             =   1920
         Visible         =   0   'False
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   767
         BTYPE           =   3
         TX              =   "取消废票(&D)"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
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
         MICON           =   "frmCancelTicket.frx":014A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin RTComctl3.CoolButton cmdCancelTicket 
         Height          =   435
         Left            =   7920
         TabIndex        =   4
         Top             =   1350
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   767
         BTYPE           =   3
         TX              =   "废票(&T)"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
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
         MICON           =   "frmCancelTicket.frx":0166
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Frame fraTktInfoChange 
         BackColor       =   &H00E0E0E0&
         Caption         =   "车票信息"
         Height          =   2775
         Left            =   150
         TabIndex        =   15
         Top             =   1260
         Width           =   7245
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "票号:"
            Height          =   180
            Left            =   135
            TabIndex        =   39
            Top             =   240
            Width           =   450
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "发车时间:"
            Height          =   180
            Left            =   135
            TabIndex        =   38
            Top             =   1335
            Width           =   810
         End
         Begin VB.Label lblOperatorChange 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "售票员:"
            Height          =   180
            Left            =   135
            TabIndex        =   37
            Top             =   1710
            Width           =   630
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "起站:"
            Height          =   180
            Index           =   1
            Left            =   135
            TabIndex        =   36
            Top             =   570
            Width           =   450
         End
         Begin VB.Label lblScheduleChange 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "车次:"
            Height          =   180
            Left            =   135
            TabIndex        =   35
            Top             =   945
            Width           =   450
         End
         Begin VB.Label lblTimeChange 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "售票时间:"
            Height          =   180
            Left            =   135
            TabIndex        =   34
            Top             =   2460
            Width           =   810
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "票价:"
            Height          =   180
            Left            =   135
            TabIndex        =   33
            Top             =   2085
            Width           =   450
         End
         Begin VB.Label lblTicketID 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "A0000134590"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   1170
            TabIndex        =   32
            Tag             =   "lblCurrentTktNum"
            Top             =   210
            Width           =   1320
         End
         Begin VB.Label lblSeatNo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "01"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   4890
            TabIndex        =   31
            Top             =   1680
            Width           =   240
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "座位号:"
            Height          =   180
            Left            =   3975
            TabIndex        =   30
            Top             =   1710
            Width           =   630
         End
         Begin VB.Label lblTypeChange 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "票种:"
            Height          =   180
            Left            =   3975
            TabIndex        =   29
            Top             =   1335
            Width           =   450
         End
         Begin VB.Label label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "到站:"
            Height          =   180
            Left            =   3975
            TabIndex        =   28
            Top             =   570
            Width           =   450
         End
         Begin VB.Label lblStateChange 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "状态:"
            Height          =   180
            Left            =   3975
            TabIndex        =   27
            Top             =   2085
            Width           =   450
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "日期:"
            Height          =   180
            Left            =   3975
            TabIndex        =   26
            Top             =   945
            Width           =   450
         End
         Begin VB.Label lblBusID 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "25101"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   1170
            TabIndex        =   25
            Top             =   915
            Width           =   600
         End
         Begin VB.Label lblEndStation 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "杭州"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   4890
            TabIndex        =   24
            Top             =   540
            Width           =   480
         End
         Begin VB.Label lblStartStation 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "宁波南站"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   1170
            TabIndex        =   23
            Top             =   540
            Width           =   960
         End
         Begin VB.Label lblSeller 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "张三"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   1170
            TabIndex        =   22
            Top             =   1680
            Width           =   480
         End
         Begin VB.Label lblTicketType 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "全票"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   4890
            TabIndex        =   21
            Top             =   1305
            Width           =   480
         End
         Begin VB.Label lblSellTime 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "2002-07-15 07:00:00"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   1170
            TabIndex        =   20
            Top             =   2445
            Width           =   2280
         End
         Begin VB.Label lblDate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "2002-07-15"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   4890
            TabIndex        =   19
            Top             =   915
            Width           =   1200
         End
         Begin VB.Label lblTicketPrice 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "37.50"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   1170
            TabIndex        =   18
            Top             =   2055
            Width           =   600
         End
         Begin VB.Label lblStatus 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "正常售出"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   4890
            TabIndex        =   17
            Top             =   2055
            Width           =   960
         End
         Begin VB.Label lblOffTime 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "10:00"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   1170
            TabIndex        =   16
            Top             =   1305
            Width           =   600
         End
      End
      Begin MSComctlLib.ListView lvTicketInfo 
         Height          =   2475
         Left            =   150
         TabIndex        =   13
         Top             =   4410
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   4366
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin STSellCtl.ucUpDownText txtCount 
         Height          =   405
         Left            =   3120
         TabIndex        =   7
         Top             =   180
         Visible         =   0   'False
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   714
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
         Value           =   "1"
      End
      Begin RTComctl3.CoolButton cmdQuery 
         Height          =   315
         Left            =   9360
         TabIndex        =   12
         Top             =   180
         Width           =   375
         _ExtentX        =   661
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
         MICON           =   "frmCancelTicket.frx":0182
         PICN            =   "frmCancelTicket.frx":019E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblEndOldTktNum 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "结束票号(&E):"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   150
         TabIndex        =   2
         Top             =   750
         Width           =   1440
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   4380
         TabIndex        =   8
         Top             =   255
         Width           =   960
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "张(&N)"
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
         Left            =   3900
         TabIndex        =   40
         Top             =   255
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "待处理车票列表(&I):"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   150
         TabIndex        =   6
         Top             =   4140
         Width           =   1890
      End
      Begin VB.Label lblOldTktNum 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "废票票号(&Z):"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   150
         TabIndex        =   0
         Top             =   240
         Width           =   1440
      End
   End
End
Attribute VB_Name = "frmCancelTicket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'废票中用的枚举
Private Enum CancelTicketInfo
    CT_BusID = 1
    CT_StartStation = 2
    CT_EndStation = 3
    CT_Date = 4
    CT_OffTime = 5
    CT_SeatNo = 6
    CT_Status = 7
    CT_SellTime = 8
    CT_TicketPrice = 9
    CT_TicketType = 10
    CT_Seller = 11
End Enum

'--需要删除,为了南昌演示
Private Sub cboSeat_GotFocus()
    cboSeat.SelStart = 0
    cboSeat.SelLength = Len(cboSeat.Text)
End Sub

Private Sub cmdCancelTicket_Click()
    Dim aszCancelTicket() As String
    On Error GoTo here
        If txtTicketNo.Text = "" Then Exit Sub
        If lvTicketInfo.ListItems.count = 0 Then Exit Sub
        
        aszCancelTicket = GetAllTickets
        
        If MsgBox("是否确认废除这些票？", vbYesNo, "提示") = vbYes Then
        m_oSell.CancelTicket aszCancelTicket
        SerialCancelTkt
        ShowMsg "正常废票成功！"
        EnableCancelButton
        lvTicketInfo.ListItems.Clear
        SetDefaultValue
        txtTicketNo.SetFocus
        txtEndTicketNo.Text = ""
        End If
'    End If
    Exit Sub
forcecancel:
    On Error GoTo here
        m_oSell.ForceCancelTicket aszCancelTicket
        SerialCancelTkt
        ShowMsg "强行废票成功！"
        lvTicketInfo.ListItems.Clear
        SetDefaultValue
        txtTicketNo.SetFocus
    Exit Sub
here:
    If err.Number = ERR_CancelTicketTimeOut Then
        If MsgBox("已过正常废票时间,要否强行废票？", vbYesNo Or vbQuestion Or vbDefaultButton2) = vbYes Then
            Resume forcecancel
        End If
    ElseIf err.Number = ERR_UserIsNotTicketSeller Then
        If MsgBox("您不是此票的售票员,要否强行废票？", vbYesNo Or vbQuestion Or vbDefaultButton2) = vbYes Then
            Resume forcecancel
        End If
    Else
        ShowErrorMsg
    End If
End Sub



Private Sub cmdQuery_Click()
 On Error GoTo ErrHandle
    Dim oEnvBus As New REBus
   
    
    oEnvBus.Init m_oAUser
    oEnvBus.Identify txtBusID.Text, dtpDate.Value
    Dim atSeat() As TSeatInfo
    atSeat = oEnvBus.GetSeatInfo

    If ArrayLength(atSeat) > 0 Then
        lvTicketInfo.ListItems.Clear
        Dim i As Integer
        For i = 1 To ArrayLength(atSeat)
            If atSeat(i).szTicketNo <> "" Then
                If cboSeat.Text <> "所有座位" And cboSeat.Text <> "" Then
                    If atSeat(i).szSeatNo = Trim(cboSeat.Text) Then
                        FillLvTicket atSeat(i).szTicketNo
                    End If
                Else
                    FillLvTicket atSeat(i).szTicketNo
                End If
            End If
        Next i
    End If
Exit Sub
ErrHandle:
    MsgBox "你输入的车次有误，请输入要废票的车次！", vbOKOnly + vbExclamation, Me.Caption: txtTicketNo.SetFocus
End Sub

Private Sub cmdResumeCancel_Click()
Dim aszTicketID() As String
On Error GoTo here
    If txtTicketNo.Text = "" Then Exit Sub
    aszTicketID = GetAllTickets
    
    If MsgBox("是否确认取消这些废票？", vbYesNo, "提示") = vbYes Then
        m_oSell.ResumeCancelTicket aszTicketID
        
        SerialCancelTkt
        ShowMsg "正常取消废票成功！"
        lvTicketInfo.ListItems.Clear
        SetDefaultValue
        EnableCancelButton
        txtTicketNo.SetFocus
    End If
    
    Exit Sub
here:
    ShowErrorMsg
End Sub

Private Sub cmdRefresh_Click()
On Error GoTo here
'
    SerialCancelTkt
On Error GoTo 0
Exit Sub
here:
    ShowErrorMsg
End Sub

Private Sub Form_Activate()
On Error GoTo here
    m_szCurrentUnitID = m_oParam.UnitID
    m_nCurrentTask = RT_CancelTicket
    txtTicketNo.Text = GetTicketNo(-1)
    MDISellTicket.SetFunAndUnit
Exit Sub
here:
    ShowErrorMsg
'-------------------
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo Error_Handle
    If KeyAscii = vbKeyReturn And (Not Me.ActiveControl Is txtTicketNo) And (Not Me.ActiveControl Is txtCount) And (Not Me.ActiveControl Is txtEndTicketNo) Then
        SendKeys "{TAB}"
    ElseIf KeyAscii = vbKeyEscape Then
        lvTicketInfo.ListItems.Clear
        txtEndTicketNo.Text = ""
        txtTicketNo.SetFocus
        EnableCancelButton
    ElseIf KeyAscii = Asc("+") Then
        '如果按了加号
        '则继续可以废下一张
        
    End If
    Exit Sub
Error_Handle:
    ShowErrorMsg
    
End Sub

Private Sub Form_Load()
'--需要删除,为了南昌演示
dtpDate.Value = Date
txtBusID.Text = ""
cboSeat.Text = "所有座位"
cboSeat.AddItem "所有座位"

On Error GoTo here
    txtTicketNo.MaxLength = 10
    FillColumnHeader
    EnableCancelButton
    SetDefaultValue
    
    On Error GoTo 0
Exit Sub
here:
    ShowErrorMsg
End Sub

Private Sub Form_Resize()
    If MDISellTicket.ActiveForm Is Me Then
        If Not Me.WindowState = vbMaximized Then Me.WindowState = vbMaximized
        fraOutLine.Left = (Me.ScaleWidth - fraOutLine.Width) / 2
        fraOutLine.Top = (Me.ScaleHeight - fraOutLine.Height) / 2
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    MDISellTicket.lblCancel.Value = vbUnchecked
    MDISellTicket.abMenuTool.Bands("mnuFunction").Tools("mnuCancelTkt").Checked = False
'    MDISellTicket.mnuCancelTkt.Checked = False
End Sub




Private Sub lvTicketInfo_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    SortListView lvTicketInfo, ColumnHeader.Index
End Sub

Private Sub lvTicketInfo_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If lvTicketInfo.ListItems.count = 0 Then SetDefaultValue
    If Not Item Is Nothing Then RefreshTicketInfo 'Item.Text
End Sub

Private Sub lvTicketInfo_KeyPress(KeyAscii As Integer)

    If Not lvTicketInfo.SelectedItem Is Nothing Then
        If KeyAscii = vbKeyBack Then
            lvTicketInfo.ListItems.Remove lvTicketInfo.SelectedItem.Index
        End If
    End If
    Exit Sub
End Sub



Private Sub RefreshTicketInfo()
    If lvTicketInfo.SelectedItem Is Nothing Then Exit Sub
    With lvTicketInfo.SelectedItem
        
        lblTicketID.Caption = .Text
        lblStartStation.Caption = .SubItems(CT_StartStation)
        lblEndStation.Caption = .SubItems(CT_EndStation)
        lblBusID.Caption = .SubItems(CT_BusID)
        lblDate.Caption = .SubItems(CT_Date)
        lblOffTime.Caption = .SubItems(CT_OffTime)
        lblTicketType.Caption = .SubItems(CT_TicketType)
        lblSeller.Caption = .SubItems(CT_Seller)
        lblSeatNo.Caption = .SubItems(CT_SeatNo)
        lblTicketPrice.Caption = .SubItems(CT_TicketPrice)
        lblStatus.Caption = .SubItems(CT_Status)
        lblSellTime.Caption = .SubItems(CT_SellTime)
        
        
    End With
'    Dim oTicket As ServiceTicket
'    Dim oCTicket As ClientTicket
'    Dim oREBus As REBus
'    On Error GoTo Here
'    If pszTicketID <> "" Then
'        Set oTicket = m_oSell.GetTicket(pszTicketID)
'        Set oCTicket = m_oSell.GetTicketClient(pszTicketID)
'
'        If Not oCTicket Is Nothing Then
'            If Trim(oCTicket.UnitID) = Trim(m_oAUser.UserUnitID) Then
'
'                Set oREBus = m_oSell.CreateServiceObject("SNRunEnv.REBus")
'                oREBus.Init m_oAUser
'                oREBus.Identify oTicket.REBusID, oTicket.REBusDate
'                If oREBus.BusType <> TP_ScrollBus Then
'                    lblOffTime.Caption = ToStandardTimeStr(oREBus.StartupTime)
'                Else
'                    lblOffTime.Caption = cszScrollBus
'                End If
'            Else
'                lblOffTime.Caption = "远程车票..."
'            End If
'            lblStartStation.Caption = oCTicket.StartStaionName
'        Else
'            Set oREBus = m_oSell.CreateServiceObject("SNRunEnv.REBus")
'            oREBus.Init m_oAUser
'            oREBus.Identify oTicket.REBusID, oTicket.REBusDate
'            If oREBus.BusType <> TP_ScrollBus Then
'                lblOffTime.Caption = ToStandardTimeStr(oREBus.StartupTime)
'            Else
'                lblOffTime.Caption = cszScrollBus
'            End If
'            lblStartStation.Caption = m_oSell.SellUnitShortName
'        End If
'
'        lblBusID.Caption = oTicket.REBusID
'        lblDate.Caption = ToStandardDateStr(oTicket.REBusDate)
'        lblEndStation.Caption = oTicket.ToStationName
'        lblTicketType.Caption = GetTicketTypeStr2(oTicket.TicketType)
'        lblTicketPrice.Caption = FormatMoney(oTicket.TicketPrice)
'        lblSeller.Caption = oTicket.Operator
'        lblSellTime.Caption = ToStandardDateTimeStr(oTicket.SellTime)
'        lblStatus.Caption = GetTicketStatusStr(oTicket.TicketStatus)
'        lblSeatNo.Caption = oTicket.SeatNo
'        lblTicketID.Caption = pszTicketID
'    End If
'    Set oTicket = Nothing
'    Set oCTicket = Nothing
'    Set oREBus = Nothing
'    On Error GoTo 0
'    Exit Sub
'Here:
'    Set oTicket = Nothing
'    Set oCTicket = Nothing
'    Set oREBus = Nothing
'    SetDefaultValue
'    ShowErrorMsg
End Sub
'显示HTMLHELP,直接拷贝
Private Sub DisplayHelp(Optional HelpType As EHelpType = content)
    Dim lActiveControl As Long
    
    Select Case HelpType
        Case content
            lActiveControl = Me.ActiveControl.HelpContextID
            If lActiveControl = 0 Then
                TopicID = Me.HelpContextID
                CallHTMLShowTopicID
            Else
                TopicID = lActiveControl
                CallHTMLShowTopicID
            End If
        Case Index
            CallHTMLHelpIndex
        Case Support
            TopicID = clSupportID
            CallHTMLShowTopicID
    End Select
End Sub

Private Sub TicketNumberAddOne()
    Dim count As Integer
    Dim TxtLenth As Integer
    Dim TicketNumber As String
    Dim ZeroNumber As Integer
    
    TxtLenth = Len(txtTicketNo.Text)
    For count = 1 To TxtLenth
       If Asc(Mid(txtTicketNo.Text, count, 1)) >= 48 And Asc(Mid(txtTicketNo.Text, count, 1)) <= 57 Then
          TicketNumber = Right(txtTicketNo.Text, TxtLenth - count + 1) + 1
          Do While Len(Right(txtTicketNo.Text, TxtLenth - count + 1)) > Len(TicketNumber)
             TicketNumber = "0" & TicketNumber
          Loop
          txtTicketNo.Text = Left(txtTicketNo.Text, count - 1) & TicketNumber
          Exit For
       End If
    Next count
End Sub

'////////////////////////////////////
'设置废票信息
Private Sub FillColumnHeader()
    Dim liTemp As ListItem
    With lvTicketInfo.ColumnHeaders
        .Add , , "票号", 1200
        .Add , , "车次", 950
        .Add , , "起站", 0
        .Add , , "到站", 1200
        .Add , , "日期", 1400
        .Add , , "时间", 1100
        .Add , , "座号", 800
        .Add , , "状态", 2100
        .Add , , "售票时间", 0
        .Add , , "票价", 1000
        .Add , , "票种", 850
        .Add , , "售票员", 1100
    End With
End Sub
'//////////////////////////////////////
'得到废票信息状态
Private Function FillLvTicket(TicketID As String) As Boolean
    
    Dim oTicket As ServiceTicket
    Dim liTemp As ListItem
    Dim oCTicket As ClientTicket
    Dim oREBus As REBus
On Error GoTo here
    
    Set oTicket = m_oSell.GetTicket(TicketID)
    Set liTemp = lvTicketInfo.ListItems.Add(, , TicketID)
    Set oCTicket = m_oSell.GetTicketClient(TicketID)
    
    
    
    With liTemp
        If Not oCTicket Is Nothing Then
            If Trim(oCTicket.UnitID) = Trim(m_oAUser.UserUnitID) Then
                Set oREBus = m_oSell.CreateServiceObject("STReSch.REBus")
                oREBus.Init m_oAUser
                oREBus.Identify oTicket.REBusID, oTicket.REBusDate
                If oREBus.BusType <> TP_ScrollBus Then
                    .SubItems(CT_OffTime) = Format(ToStandardTimeStr(oREBus.StartUpTime), "hh:mm")
                Else
                    .SubItems(CT_OffTime) = cszScrollBus
                End If
            Else
                .SubItems(CT_OffTime) = "远程车票..."
            End If
            .SubItems(CT_StartStation) = oCTicket.StartStaionName
        Else
            Set oREBus = m_oSell.CreateServiceObject("STReSch.REBus")
            oREBus.Init m_oAUser
            oREBus.Identify oTicket.REBusID, oTicket.REBusDate
            If oREBus.BusType <> TP_ScrollBus Then
               .SubItems(CT_OffTime) = Format(ToStandardTimeStr(oREBus.StartUpTime), "hh:mm")
            Else
               .SubItems(CT_OffTime) = cszScrollBus
            End If
            .SubItems(CT_StartStation) = m_oSell.SellUnitShortName
        End If
        .SubItems(CT_SeatNo) = oTicket.SeatNo
        
        .SubItems(CT_BusID) = oTicket.REBusID
        .SubItems(CT_Date) = ToStandardDateStr(oTicket.REBusDate)
        .SubItems(CT_EndStation) = oTicket.ToStationName
        .SubItems(CT_TicketType) = GetTicketTypeStr2(oTicket.TicketType)
        .SubItems(CT_TicketPrice) = FormatMoney(oTicket.TicketPrice)
        .SubItems(CT_Seller) = oTicket.Operator
        .SubItems(CT_SellTime) = ToStandardDateTimeStr(oTicket.SellTime)
        .SubItems(CT_Status) = GetTicketStatusStr(oTicket.TicketStatus)
    End With
    Set oCTicket = Nothing
    Set oTicket = Nothing
    Set liTemp = Nothing
    FillLvTicket = True
    Exit Function
here:
    Set oCTicket = Nothing
    Set oTicket = Nothing
    Set liTemp = Nothing
    ShowErrorMsg
    FillLvTicket = False
End Function

'///////////////////////////
'成批废票中得到车票信息显示于车票信息ListView中
Private Sub SerialCancelTkt()
    Dim lTemp1 As Long, lTemp2 As Long, lTemp3 As Long
    Dim szTemp As String
    Dim lCount As Long
    On Error GoTo here
    lvTicketInfo.ListItems.Clear
    lTemp1 = Right(txtTicketNo.Text, TicketNoNumLen())
'    lTemp2 = lTemp1 + txtCount.Value - 1
    If txtEndTicketNo.Text <> "" Then
        lTemp3 = Right(txtEndTicketNo.Text, TicketNoNumLen())
        lTemp2 = lTemp3
    Else
        lTemp2 = lTemp1
    End If
    
    If lTemp3 - lTemp1 + 1 <= 100 Then
        lvTicketInfo.ListItems.Clear
        szTemp = Left(txtTicketNo.Text, Len(txtTicketNo.Text) - TicketNoNumLen())
        For lCount = lTemp1 To lTemp2
            If Not FillLvTicket(szTemp & String(TicketNoNumLen() - Len(CStr(lCount)), "0") & lCount) Then
                Exit Sub
            End If
        Next lCount
        If lvTicketInfo.ListItems.count > 0 Then lvTicketInfo.ListItems(lvTicketInfo.ListItems.count).Selected = True
    '        If lCount > lTemp1 Then
    '            RefreshTicketInfo szTemp & String(TicketNoNumLen() - Len(CStr(lCount - 1)), "0") & lCount - 1
    '        End If
    '    End If
        
        RefreshTicketInfo
    '    ShowReturnInfo
    '    GetReturnMoney
    Else
        MsgBox "为保证系统运行效率，废票张数应在100张以内！", vbInformation, "注意"
    End If
    EnableCancelButton
    On Error GoTo 0
Exit Sub
here:
    
    ShowErrorMsg
End Sub


'--需要删除,为了南昌演示
Private Sub txtBusID_GotFocus()
    txtBusID.SelStart = 0
    txtBusID.SelLength = Len(txtBusID.Text)
End Sub

Private Sub txtCount_Change()
'    If txtCount.value = 0 Then txtCount.value = 1
    EnableCancelButton
'    lvTicketInfo.ListItems.Clear
'    SetDefaultValue
End Sub

Private Sub txtCount_GotFocus()
    txtCount.SelStart = 0
    txtCount.SelLength = 100
End Sub

Private Sub txtCount_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtTicketNo_KeyPress KeyAscii
    End If
End Sub

Private Sub txtEndTicketNo_Change()
    EnableCancelButton
End Sub

Private Sub txtEndTicketNo_GotFocus()
        txtEndTicketNo.SelStart = 0
        txtEndTicketNo.SelLength = 100
End Sub

Private Sub txtEndTicketNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim szTemp As String
    Dim lTemp As Long
    On Error GoTo Error_Handel
    If Len(txtEndTicketNo.Text) >= TicketNoNumLen() Then
        If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
            lTemp = Right(txtEndTicketNo.Text, TicketNoNumLen())
            szTemp = Left(txtEndTicketNo.Text, Len(txtEndTicketNo.Text) - TicketNoNumLen())
            
            If KeyCode = vbKeyUp Then
                lTemp = lTemp + 1
            Else
                lTemp = lTemp - 1
                lTemp = IIf(lTemp > 0, lTemp, 0)
            End If
            txtEndTicketNo.Text = MakeTicketNo(lTemp, szTemp)
            KeyCode = 0
        End If
    End If
    Exit Sub
Error_Handel:
End Sub

Private Sub txtEndTicketNo_KeyPress(KeyAscii As Integer)
On Error GoTo here
    If KeyAscii = 13 And txtTicketNo.Text <> "" And txtCount.Value > 0 Then
        SerialCancelTkt
        cmdCancelTicket.SetFocus
    End If
On Error GoTo 0
Exit Sub
here:
  
  ShowErrorMsg
End Sub

Private Sub txtEndTicketNo_LostFocus()
    If txtEndTicketNo <> "" Then
        If Val(Right(txtEndTicketNo.Text, TicketNoNumLen())) < Val(Right(txtTicketNo.Text, TicketNoNumLen())) Then
            MsgBox "结束票号应大于起始票号！", vbInformation, "出错"
        End If
    End If
End Sub

Private Sub txtTicketNo_Change()
    EnableCancelButton
'    lvTicketInfo.ListItems.Clear
'    SetDefaultValue
End Sub

Private Sub txtTicketNo_GotFocus()
        txtTicketNo.SelStart = 0
        txtTicketNo.SelLength = 100 'Len(txtTicketNo.Text)
End Sub

Private Sub txtTicketNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim szTemp As String
    Dim lTemp As Long
    On Error GoTo Error_Handel
    If Len(txtTicketNo.Text) >= TicketNoNumLen() Then
        If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
            lTemp = Right(txtTicketNo.Text, TicketNoNumLen())
            szTemp = Left(txtTicketNo.Text, Len(txtTicketNo.Text) - TicketNoNumLen())
            
            If KeyCode = vbKeyUp Then
                lTemp = lTemp + 1
            Else
                lTemp = lTemp - 1
                lTemp = IIf(lTemp > 0, lTemp, 0)
            End If
            txtTicketNo.Text = MakeTicketNo(lTemp, szTemp)
            KeyCode = 0
        End If
    End If
    Exit Sub
Error_Handel:
End Sub

'设置退票按钮状态
Private Sub EnableCancelButton()
'    设置废票按钮
    If txtTicketNo.Text <> "" And txtCount.Value > 0 And lvTicketInfo.ListItems.count > 0 Then
        cmdCancelTicket.Enabled = True
        cmdResumeCancel.Enabled = True
    Else
        cmdCancelTicket.Enabled = False
        cmdResumeCancel.Enabled = False
    End If
    
End Sub

Private Sub txtTicketNo_KeyPress(KeyAscii As Integer)
On Error GoTo here
    If KeyAscii = 13 And txtTicketNo.Text <> "" And txtCount.Value > 0 Then
        SerialCancelTkt
        If cmdCancelTicket.Enabled Then cmdCancelTicket.SetFocus
    End If
On Error GoTo 0
Exit Sub
here:
  
  ShowErrorMsg
End Sub

Private Sub SetDefaultValue()
    '设置默认的控件值
    lblStartStation.Caption = ""
    lblEndStation.Caption = ""
    lblBusID.Caption = ""
    lblDate.Caption = ""
    lblOffTime.Caption = ""
    lblTicketType.Caption = ""
    lblSeller.Caption = ""
    lblSeatNo.Caption = ""
    lblTicketPrice.Caption = ""
    lblStatus.Caption = ""
    lblSellTime.Caption = ""
    lblTicketID.Caption = ""
    
End Sub

Private Function GetAllTickets() As String()
    '根据txtTicketNo 和txtCount 得到所有的票
    Dim lTemp1 As Long
    Dim lTemp2 As Long
    Dim szTemp As String
    Dim lCount As Long
    Dim aszTemp() As String
'    lTemp1 = Right(txtTicketNo.Text, TicketNoNumLen())
'    lTemp2 = lTemp1 + txtCount.Value - 1
'    szTemp = Left(txtTicketNo.Text, Len(txtTicketNo.Text) - TicketNoNumLen())
'    ReDim aszTemp(1 To txtCount.Value)
'    For lCount = lTemp1 To lTemp2
'        aszTemp(lCount - lTemp1 + 1) = szTemp & String(TicketNoNumLen() - Len(CStr(lCount)), "0") & lCount
'    Next lCount
'    GetAllTickets = aszTemp
    If lvTicketInfo.ListItems.count > 0 Then
        ReDim aszTemp(1 To lvTicketInfo.ListItems.count)
        For lCount = 1 To lvTicketInfo.ListItems.count
            aszTemp(lCount) = lvTicketInfo.ListItems(lCount).Text
        Next lCount
        GetAllTickets = aszTemp
    End If
    
End Function





