VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{695ABF14-B2D8-11D2-A5ED-DE08DCF33612}#1.2#0"; "asfcombo.ocx"
Object = "{BBF95DAA-F9CB-4CA9-A673-E0E9E0193752}#1.0#0"; "STSellCtl.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmExtraSell 
   BackColor       =   &H00929292&
   Caption         =   "补票"
   ClientHeight    =   7710
   ClientLeft      =   675
   ClientTop       =   2505
   ClientWidth     =   11400
   HelpContextID   =   4000140
   Icon            =   "frmExtraSell.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7710
   ScaleWidth      =   11400
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1245
      HelpContextID   =   3000411
      Left            =   6300
      TabIndex        =   42
      Top             =   8670
      Visible         =   0   'False
      Width           =   2880
      Begin VB.CheckBox chkSetSeat 
         BackColor       =   &H00FFFFFF&
         Caption         =   "定座(&P)"
         Height          =   270
         HelpContextID   =   3000411
         Left            =   120
         TabIndex        =   43
         Top             =   -30
         Width           =   975
      End
      Begin VB.Label lblSetSeat 
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   180
         TabIndex        =   44
         Top             =   60
         Width           =   435
      End
   End
   Begin VB.Timer Timer1 
      Left            =   5820
      Top             =   3840
   End
   Begin VB.CommandButton cmdPreSell 
      Caption         =   "预售(&C)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   7350
      TabIndex        =   19
      Top             =   1980
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.Frame fraDiscountTicket 
      Caption         =   "折扣票"
      Height          =   615
      Left            =   3210
      TabIndex        =   16
      Top             =   9420
      Visible         =   0   'False
      Width           =   1665
      Begin VB.TextBox txtDiscount 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   1080
         TabIndex        =   17
         Text            =   "1"
         Top             =   330
         Width           =   1155
      End
      Begin VB.Label lblDiscount 
         Caption         =   "折扣(&F):"
         Height          =   195
         Left            =   270
         TabIndex        =   18
         Top             =   300
         Width           =   795
      End
   End
   Begin VB.Frame fraOutLine 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   7095
      Left            =   600
      TabIndex        =   15
      Top             =   210
      Width           =   11760
      Begin VB.CheckBox chkInsurance 
         BackColor       =   &H00E0E0E0&
         Caption         =   "保险(F12)"
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   4185
         TabIndex        =   70
         Top             =   4980
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.Timer Timer2 
         Interval        =   10
         Left            =   3495
         Top             =   2055
      End
      Begin VB.PictureBox ptExtraByBus 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         FillColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   2670
         Left            =   135
         ScaleHeight     =   2670
         ScaleWidth      =   2550
         TabIndex        =   52
         Top             =   435
         Width           =   2550
         Begin VB.TextBox txtBusID 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   1275
            TabIndex        =   11
            Top             =   120
            Width           =   1185
         End
         Begin RTComctl3.FlatLabel flblOffTimeExtra 
            Height          =   315
            Left            =   1275
            TabIndex        =   53
            Top             =   675
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   14737632
            OutnerStyle     =   2
            Caption         =   ""
         End
         Begin RTComctl3.FlatLabel flblSeatCount 
            Height          =   315
            Left            =   1275
            TabIndex        =   54
            Top             =   1575
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   14737632
            OutnerStyle     =   2
            Caption         =   "0"
         End
         Begin RTComctl3.FlatLabel flblBusTypeExtra 
            Height          =   315
            Left            =   1275
            TabIndex        =   55
            Top             =   1125
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   14737632
            OutnerStyle     =   2
            Caption         =   ""
         End
         Begin RTComctl3.FlatLabel flblStandCount 
            Height          =   315
            Left            =   1275
            TabIndex        =   56
            Top             =   2010
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   14737632
            OutnerStyle     =   2
            Caption         =   "0"
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "站票数:"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   90
            TabIndex        =   62
            Top             =   2100
            Width           =   735
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            Index           =   1
            X1              =   60
            X2              =   3135
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Label lblBusTypeExtra 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "车型:"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   90
            TabIndex        =   61
            Top             =   1185
            Width           =   525
         End
         Begin VB.Label lblOffTimeExtra 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "发车时间:"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   90
            TabIndex        =   60
            Top             =   720
            Width           =   945
         End
         Begin VB.Label lblAvailableSeatsNumExtra 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "可售座位数:"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   90
            TabIndex        =   59
            Top             =   1650
            Width           =   1155
         End
         Begin VB.Label lblScheduleExtra 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "补票车次(&Q):"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   90
            TabIndex        =   10
            Top             =   225
            Width           =   1260
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00000000&
            Index           =   0
            X1              =   60
            X2              =   2550
            Y1              =   585
            Y2              =   585
         End
         Begin VB.Label lblCanExtraSell 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   1500
            TabIndex        =   58
            Top             =   2550
            Width           =   120
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "是否允许补票:"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   90
            TabIndex        =   57
            Top             =   2430
            Width           =   1365
         End
      End
      Begin VB.PictureBox ptExtraSellByStation 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   2715
         Left            =   90
         ScaleHeight     =   2715
         ScaleWidth      =   2640
         TabIndex        =   63
         Top             =   465
         Visible         =   0   'False
         Width           =   2640
         Begin MSComctlLib.ImageList ImageList1 
            Left            =   3210
            Top             =   1125
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   1
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmExtraSell.frx":014A
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin STSellCtl.ucSuperCombo cboEndStation 
            Height          =   2400
            Left            =   60
            TabIndex        =   1
            Top             =   240
            Width           =   2475
            _ExtentX        =   4366
            _ExtentY        =   4233
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
         Begin VB.Label lblToStation 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "到站(&Z):"
            Height          =   180
            Left            =   60
            TabIndex        =   0
            Top             =   0
            Width           =   720
         End
      End
      Begin RTComctl3.CoolButton cmdSell 
         Height          =   555
         Left            =   90
         TabIndex        =   7
         Top             =   5100
         Width           =   2550
         _ExtentX        =   4498
         _ExtentY        =   979
         BTYPE           =   3
         TX              =   "售出(&P)"
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
         MICON           =   "frmExtraSell.frx":02A4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox txtReceivedMoney 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1140
         TabIndex        =   9
         Text            =   "0.00"
         Top             =   6210
         Width           =   1575
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   6390
         Top             =   2010
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483633
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   16777215
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmExtraSell.frx":02C0
               Key             =   "bus"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmExtraSell.frx":041C
               Key             =   "station"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lvStationPreSell 
         Height          =   1725
         Left            =   5880
         TabIndex        =   40
         Top             =   5280
         Width           =   5730
         _ExtentX        =   10107
         _ExtentY        =   3043
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
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lvBusPreSell 
         Height          =   1725
         Left            =   5880
         TabIndex        =   41
         Top             =   5280
         Width           =   5730
         _ExtentX        =   10107
         _ExtentY        =   3043
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
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lvBus 
         Height          =   3930
         Left            =   2760
         TabIndex        =   4
         Top             =   840
         Width           =   8745
         _ExtentX        =   15425
         _ExtentY        =   6932
         SortKey         =   1
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   16777215
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
         NumItems        =   12
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
            Key             =   "LimitedCount"
            Text            =   "限售张数"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Key             =   "LimitedTime"
            Text            =   "限售时间"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Key             =   "BusType"
            Text            =   "车次类型"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Key             =   "CheckGate"
            Text            =   "检票口"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Key             =   "StandCount"
            Text            =   "站票"
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.ListView lvStation 
         Height          =   3930
         Left            =   2790
         TabIndex        =   12
         Top             =   840
         Width           =   8745
         _ExtentX        =   15425
         _ExtentY        =   6932
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         Sorted          =   -1  'True
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
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "EndStationID"
            Text            =   "终到站代码"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "EndStation"
            Text            =   "终到站"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Key             =   "WTktPrice"
            Text            =   "全价"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Key             =   "STktPrice"
            Text            =   "半价"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Key             =   "LimitedCount"
            Text            =   "限售张数"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Key             =   "LimitedTime"
            Text            =   "限售时间"
            Object.Width           =   0
         EndProperty
      End
      Begin RTComctl3.CoolButton cmdBus 
         Height          =   270
         Left            =   120
         TabIndex        =   46
         Top             =   105
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   476
         BTYPE           =   9
         TX              =   "按车次补票(&B)"
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
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmExtraSell.frx":0578
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   -1  'True
         VALUE           =   -1  'True
      End
      Begin RTComctl3.CoolButton cmdStation 
         Height          =   270
         Left            =   1720
         TabIndex        =   45
         Top             =   105
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   476
         BTYPE           =   3
         TX              =   "按到站补票(&S)"
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
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmExtraSell.frx":0594
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   -1  'True
         VALUE           =   -1  'True
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1635
         Left            =   90
         TabIndex        =   47
         Top             =   3120
         Width           =   2655
         Begin FCmbo.asFlatCombo cboSeatType 
            Height          =   330
            Left            =   1440
            TabIndex        =   48
            Top             =   1245
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   582
            ButtonDisabledForeColor=   12632256
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   16711680
            ButtonHotBackColor=   8421504
            ButtonPressedBackColor=   0
            Text            =   ""
            ButtonBackColor =   8421504
            Style           =   1
            Registered      =   -1  'True
            OfficeXPColors  =   -1  'True
         End
         Begin STSellCtl.ucNumTextBox txtFullSell 
            Height          =   300
            Left            =   1425
            TabIndex        =   6
            Top             =   150
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   2
            Alignment       =   1
         End
         Begin FCmbo.asFlatCombo cboPreferentialTicket 
            Height          =   330
            Left            =   195
            TabIndex        =   49
            Top             =   870
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   582
            ButtonDisabledForeColor=   12632256
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonHotBackColor=   8421504
            ButtonPressedBackColor=   0
            Text            =   ""
            ButtonBackColor =   8421504
            Style           =   1
            Registered      =   -1  'True
            OfficeXPColors  =   -1  'True
         End
         Begin STSellCtl.ucNumTextBox txtPreferentialSell 
            Height          =   300
            Left            =   1425
            TabIndex        =   50
            Top             =   870
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   2
            Alignment       =   1
         End
         Begin STSellCtl.ucNumTextBox txtHalfSell 
            Height          =   300
            Left            =   1425
            TabIndex        =   14
            Top             =   510
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   2
            Alignment       =   1
         End
         Begin MSComCtl2.UpDown upFull 
            Height          =   300
            Left            =   2266
            TabIndex        =   71
            TabStop         =   0   'False
            Top             =   150
            Width           =   255
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            BuddyControl    =   "txtFullSell"
            BuddyDispid     =   196649
            OrigLeft        =   2370
            OrigTop         =   3090
            OrigRight       =   2625
            OrigBottom      =   3405
            Max             =   100
            SyncBuddy       =   -1  'True
            BuddyProperty   =   1745027080
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown upHalf 
            Height          =   300
            Left            =   2266
            TabIndex        =   72
            TabStop         =   0   'False
            Top             =   510
            Width           =   255
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            BuddyControl    =   "txtHalfSell"
            BuddyDispid     =   196652
            OrigLeft        =   2370
            OrigTop         =   3090
            OrigRight       =   2625
            OrigBottom      =   3405
            Max             =   100
            SyncBuddy       =   -1  'True
            BuddyProperty   =   1745027080
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown ucPreferential 
            Height          =   300
            Left            =   2266
            TabIndex        =   73
            TabStop         =   0   'False
            Top             =   870
            Width           =   255
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            BuddyControl    =   "txtPreferentialSell"
            BuddyDispid     =   196651
            OrigLeft        =   2370
            OrigTop         =   3090
            OrigRight       =   2625
            OrigBottom      =   3405
            Max             =   100
            SyncBuddy       =   -1  'True
            BuddyProperty   =   1745027080
            Enabled         =   -1  'True
         End
         Begin VB.Line lnTicketType 
            X1              =   75
            X2              =   2805
            Y1              =   0
            Y2              =   0
         End
         Begin VB.Label lblSeatType 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "座位类型(&R):"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   195
            TabIndex        =   51
            Top             =   1290
            Width           =   1260
         End
         Begin VB.Label lblHalfSell 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "半票(&X):"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   195
            TabIndex        =   13
            Top             =   570
            Width           =   840
         End
         Begin VB.Label lblFullSell 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "全票(&A):"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   195
            TabIndex        =   5
            Top             =   240
            Width           =   840
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            Index           =   2
            X1              =   60
            X2              =   3135
            Y1              =   -45
            Y2              =   -45
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         Height          =   4605
         Left            =   60
         TabIndex        =   64
         Top             =   270
         Width           =   11595
         Begin VB.CommandButton cmdSetSeat 
            Caption         =   "定座(&G)"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            HelpContextID   =   3000411
            Left            =   9435
            TabIndex        =   69
            Top             =   195
            Width           =   1230
         End
         Begin VB.TextBox txtSeat 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            HelpContextID   =   3000411
            Left            =   7830
            TabIndex        =   67
            Top             =   180
            Width           =   1395
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "座位号(&T):"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   6780
            TabIndex        =   68
            Top             =   270
            Width           =   1050
         End
         Begin VB.Label lblBus 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "车次列表(&V):"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2805
            TabIndex        =   2
            Top             =   240
            Width           =   1260
         End
         Begin VB.Label lblStation 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "站点列表(&N):"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2805
            TabIndex        =   3
            Top             =   255
            Width           =   1260
         End
      End
      Begin MSComctlLib.ListView lvSellStation 
         Height          =   1725
         Left            =   2790
         TabIndex        =   65
         Top             =   5280
         Width           =   3030
         _ExtentX        =   5345
         _ExtentY        =   3043
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "上车站"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   " 时间"
            Object.Width           =   1695
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   " 票价"
            Object.Width           =   1676
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "上车站代码"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "检票门"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label lblSellStation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "上车站(&O):"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2790
         TabIndex        =   66
         Top             =   5010
         Width           =   1050
      End
      Begin VB.Label flblTotalPrice 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   345
         Left            =   2040
         TabIndex        =   37
         Top             =   5850
         Width           =   585
      End
      Begin VB.Label flblRestMoney 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   2040
         TabIndex        =   36
         Top             =   6690
         Width           =   585
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "应找:"
         Height          =   180
         Left            =   90
         TabIndex        =   27
         Top             =   6750
         Width           =   450
      End
      Begin VB.Line Line4 
         X1              =   30
         X2              =   2730
         Y1              =   6690
         Y2              =   6690
      End
      Begin VB.Label lblReceivedMoney 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "实收(&Q):"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   90
         TabIndex        =   8
         Top             =   6300
         Width           =   840
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "累计总票价:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   90
         TabIndex        =   26
         Top             =   5970
         Width           =   1155
      End
      Begin VB.Line Line3 
         X1              =   30
         X2              =   2760
         Y1              =   4950
         Y2              =   4950
      End
      Begin VB.Label lblSinglePrice 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00/0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   330
         Left            =   1410
         TabIndex        =   25
         Top             =   5550
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车票单价:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   90
         TabIndex        =   24
         Top             =   5640
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label lblPreBus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "预售车次列表信息(&E):"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   5895
         TabIndex        =   23
         Top             =   5010
         Width           =   2100
      End
      Begin VB.Label lblSellMsg 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   8070
         TabIndex        =   22
         Top             =   4920
         Width           =   2115
      End
   End
   Begin RTComctl3.FlatLabel flblLimitedTime2 
      Height          =   375
      Left            =   11580
      TabIndex        =   28
      Top             =   8640
      Visible         =   0   'False
      Width           =   2550
      _ExtentX        =   4498
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12632256
      OutnerStyle     =   2
      Caption         =   ""
   End
   Begin RTComctl3.FlatLabel flblLimitedCount2 
      Height          =   375
      Left            =   9120
      TabIndex        =   29
      Top             =   8610
      Visible         =   0   'False
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12632256
      OutnerStyle     =   2
      Caption         =   ""
   End
   Begin RTComctl3.FlatLabel flblLimitedCount 
      Height          =   345
      Left            =   2580
      TabIndex        =   32
      Top             =   8310
      Visible         =   0   'False
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12632256
      OutnerStyle     =   2
      Caption         =   ""
   End
   Begin RTComctl3.FlatLabel flblLimitedTime 
      Height          =   375
      Left            =   5280
      TabIndex        =   33
      Top             =   8310
      Visible         =   0   'False
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12632256
      OutnerStyle     =   2
      Caption         =   ""
   End
   Begin RTComctl3.FlatLabel flblStandCount2 
      Height          =   375
      Left            =   8700
      TabIndex        =   38
      Top             =   9090
      Visible         =   0   'False
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12632256
      OutnerStyle     =   2
      Caption         =   "0"
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "站票:"
      Height          =   180
      Left            =   8010
      TabIndex        =   39
      Top             =   9180
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "限售时间:"
      Height          =   195
      Left            =   4050
      TabIndex        =   35
      Top             =   8400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "限售张数:"
      Height          =   225
      Left            =   1080
      TabIndex        =   34
      Top             =   8430
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.Label Label7 
      Caption         =   "限售张数:"
      Height          =   225
      Left            =   8040
      TabIndex        =   31
      Top             =   8700
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.Label Label6 
      Caption         =   "限售时间:"
      Height          =   195
      Left            =   10590
      TabIndex        =   30
      Top             =   8700
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "总票价:"
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
      Left            =   6120
      TabIndex        =   21
      Top             =   9060
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Label lblTotalMoney 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7110
      TabIndex        =   20
      Top             =   9000
      Visible         =   0   'False
      Width           =   585
   End
End
Attribute VB_Name = "frmExtraSell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const cnActiveStatus = [Office XP]
Const cnNotActiveStatus = [Windows XP]

'补票用的枚举
Private Enum ExtraTicketInfo
    EI_BusType = 1
    EI_OffTime = 2
    EI_EndStation = 3
    EI_TotalNum = 4
    EI_TotalPrice = 5
    EI_VehicleModel = 6
    EI_FullPrice = 7
    EI_FullNum = 8
    EI_HalfPrice = 9
    EI_HalfNum = 10
    EI_FreePrice = 11
    EI_FreeNum = 12
    EI_PreferentialPrice1 = 13
    EI_PreferentialNum1 = 14
    EI_PreferentialPrice2 = 15
    EI_PreferentialNum2 = 16
    EI_PreferentialPrice3 = 17
    EI_PreferentialNum3 = 18
    EI_Discount = 19
    EI_OrderSeat = 20
    EI_StandTicket = 21
    EI_CheckGate = 22
    EI_EndStationCode = 23
    EI_LimitedCount = 24
    EI_SeatStatus1 = 25
    EI_SeatStatus2 = 26
    EI_SeatNo = 27
    EI_AllTicketPrice = 28
    EI_AllTicketType = 29
    EI_SeatType = 30
    EI_TerminateName = 31
    EI_SumTicketNum = 32
    EI_RealName = 33
End Enum


'补票用的枚举
Private Enum EBusStationInfoIndex
'    ID_StationID = 1
    ID_StationName = 1
    ID_BusType2 = 2
    ID_FullPrice2 = 3
    ID_HalfPrice2 = 4
    ID_FreePrice2 = 5
    ID_PreferentialPrice21 = 6
    ID_PreferentialPrice22 = 7
    ID_PreferentialPrice23 = 8
    ID_BedFullPrice2 = 9
    ID_BedHalfPrice2 = 10
    ID_BedFreePrice2 = 11
    ID_BedPreferential21 = 12
    ID_BedPreferential22 = 13
    ID_BedPreferential23 = 14
    ID_AdditionalFullPrice2 = 15
    ID_AdditionalHalfPrice2 = 16
    ID_AdditionalFreePrice2 = 17
    ID_AdditionalPreferential21 = 18
    ID_AdditionalPreferential22 = 19
    ID_AdditionalPreferential23 = 20
    ID_LimitedCount2 = 21
    ID_LimitedTime2 = 22
    ID_TerminateName2 = 23
    ID_RealName = 24
End Enum
Const cszAllowExtraSell = "允许"
Const cszNotAllowExtraSell = "不允许"


Private m_bBusInfoDirty As Boolean
Private m_szCheckGate As String
Private m_bPointCount As Boolean
Private m_sgTotalMoney As Single '记录上一次售票的金额
Private m_tbSeatTypeBus As TMultiSeatTypeBus '得到不同座位类型的车次
Private m_aszSeatType() As String
Private m_dbTotalPrice As Single
Private m_atTicketType() As TTicketType
Private rsCountTemp As Recordset

Private m_aszInsurce() As String '乘意险

Private m_aTReBusAllotInfo() As TReBusAllotInfo

Private m_aszRealNameInfo() As TCardInfo

Private Sub ClearInfo()
    Erase m_aszRealNameInfo
End Sub


Private Sub cboEndStation_Change()
    If lvBus.ListItems.count > 0 Then
        lvBus.ListItems.Clear
        lvSellStation.ListItems.Clear
        DoThingWhenBusChange
    End If
    SetPreSellButton
    EnableSellButton
End Sub

Private Sub cboEndStation_GotFocus()
    lblToStation.ForeColor = clActiveColor
End Sub

Private Sub cboEndStation_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If lvBus.ListItems.count = 0 Then
          cboEndStation.SetFocus
       End If
    End If
End Sub



Private Sub cboEndStation_LostFocus()
    lblToStation.ForeColor = 0
    RefreshBus True
    SetDefaultSellTicket
    If lvBus.ListItems.count = 0 Then
       lvSellStation.ListItems.Clear
    End If
    txtReceivedMoney.Text = ""
    EnableSellButton
End Sub

Private Sub cboPreferentialTicket_Change()
    txtPreferentialSell.Text = 0
End Sub

Private Sub cboPreferentialTicket_Click()
    txtPreferentialSell.Text = 0
End Sub

Private Sub cboPreferentialTicket_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
        Case vbKeyRight
            If cboPreferentialTicket.ListIndex <> 0 Then
                cboPreferentialTicket.ListIndex = cboPreferentialTicket.ListIndex - 1
            End If
              txtPreferentialSell.SetFocus
        Case vbKeyLeft
                If cboPreferentialTicket.ListIndex <> cboPreferentialTicket.ListCount - 1 Then
                    cboPreferentialTicket.ListIndex = cboPreferentialTicket.ListIndex + 1
                End If
               txtHalfSell.SetFocus
    End Select
End Sub



Private Sub cmdForFocusMoney_GotFocus()
    If ptExtraByBus.Visible Then
        txtBusID.SetFocus
    Else
        cboEndStation.SetFocus
    End If
End Sub


Private Sub cboSeatType_Change()
    If lvBus.Visible = True Then
      If lvSellStation.ListItems.count > 0 Then
       RefreshBusStation rsCountTemp, Trim(lvSellStation.SelectedItem.SubItems(3)), cboSeatType.ListIndex + 1
      End If
    End If
    If lvStation.Visible = True Then
      If lvSellStation.ListItems.count > 0 Then
       RefreshBusStation rsCountTemp, Trim(lvSellStation.SelectedItem.SubItems(3)), cboSeatType.ListIndex + 1
      End If
    End If
End Sub

Private Sub cboSeatType_GotFocus()
    lblSeatType.ForeColor = clActiveColor
End Sub

Private Sub cboSeatType_LostFocus()
    lblSeatType.ForeColor = 0
End Sub



Private Sub chkInsurance_Click()
    DealPrice
End Sub

Private Sub cmdBus_Click()
    SetBus
End Sub

Private Sub cmdPreSell_Click()
Dim nSameIndex As Integer
    If txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text = 0 Then Exit Sub
   cmdPreSell.Enabled = False
   If BusIsActive Then
        GetBusPreSellInfo
   Else

        GetStationPreSellInfo
   End If
   txtFullSell.Text = 0
   txtHalfSell.Text = 0
   txtPreferentialSell.Text = 0
   txtSeat.Text = ""
   cmdPreSell.Enabled = True
End Sub

'补票处理过程
Private Sub cmdSell_Click()
    
    
    Dim k As Long
    Dim m As Long
    Dim i As Integer
    m = 0
'根据绍兴的情况，新加入对输入了票数，但未在lvPreSell中体现记录的情况，使其能在直接输入数字后点击打印按钮后卖票
    If lvStationPreSell.ListItems.count = 0 Then
        Call txtFullSell_KeyPress(vbKeyReturn)
    End If
    
    For i = 1 To lvBusPreSell.ListItems.count
        m = m + lvBusPreSell.ListItems(i).SubItems(EI_SumTicketNum)
    Next i
    For i = 1 To lvStationPreSell.ListItems.count
        m = m + lvStationPreSell.ListItems(i).SubItems(EI_SumTicketNum)
    Next i
    If m_lEndTicketNoOld = 0 Then
        ShowMsg "售票不成功，用户还未领票，请先去领票！"
        Exit Sub
    End If
    If m + Val(txtFullSell.Text) + Val(txtHalfSell.Text) + Val(txtPreferentialSell.Text) + Val(m_lTicketNo) - 1 > Val(m_lEndTicketNo) Then
        k = Val(m_lEndTicketNo) - Val(m_lTicketNo) + 1
        MsgBox "打印机上的票已不够！" & vbCrLf & "车票只剩 " & k & "张", vbInformation, "售票台"
    Else
        ExtraSell
    End If
End Sub


Private Sub cmdSetSeat_Click()
    Dim rsTemp As Recordset
    On Error GoTo here
    If ptExtraByBus.Visible Then
        Set rsTemp = m_oSell.GetSeatRs(m_oParam.NowDate, txtBusID.Text)
    Else
        Set rsTemp = m_oSell.GetSeatRs(m_oParam.NowDate, lvBus.SelectedItem.Text)
    End If
    
    Set frmOrderSeats.m_rsSeat = rsTemp
    frmOrderSeats.m_szSeatNumber = PreOrderSeat
    frmOrderSeats.Show vbModal
    If frmOrderSeats.m_bOk Then
        txtSeat = frmOrderSeats.m_szSeat
    End If
    Unload frmOrderSeats
    
    Set rsTemp = Nothing
    On Error GoTo 0
    Exit Sub
here:
    Set rsTemp = Nothing
    ShowErrorMsg
End Sub

Private Sub cmdStation_Click()
    SetStation
End Sub

Private Sub flblTotalPrice_Change()
    CalReceiveMoney
End Sub

Private Sub CalReceiveMoney()
    Dim dbTemp As Double
    dbTemp = Val(txtReceivedMoney.Text) - Val(flblTotalPrice.Caption)
    If Val(dbTemp) <= 0 Then
        flblRestMoney.Caption = ""
    Else
        flblRestMoney.Caption = FormatMoney(dbTemp)
    End If
End Sub
Private Sub Form_Activate()
On Error GoTo here
    m_szCurrentUnitID = m_oParam.UnitID
    m_nCurrentTask = RT_ExtraSellTicket
    SetBusPreSellInfo
    SetStationPreSellInfo
    MDISellTicket.SetFunAndUnit
'    If tsExtraType.SelectedItem.Index = 1 Then
'        MDISellTicket.EnableSortAndRefresh True
'    End If
'    MDISellTicket.abMenuTool.Bands("mnuFunction").Tools("mnuChangeSeatType").Enabled = True
    
    On Error GoTo 0
    Exit Sub
here:
    ShowErrorMsg
End Sub

Private Sub Form_Deactivate()
    MDISellTicket.EnableSortAndRefresh False
'    MDISellTicket.abMenuTool.Bands("mnuFunction").Tools("mnuChangeSeatType").Enabled = False
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF8 Then
        '定座位
        If cmdSetSeat.Enabled = True Then
            cmdSetSeat_Click
        End If
    ElseIf KeyCode = vbKeyF12 Then
''暂时去掉2006-1-11
''        '选中需要保险
''        If chkInsurance.Value = vbChecked Then
''            chkInsurance.Value = vbUnchecked
''        Else
''            chkInsurance.Value = vbChecked
''        End If
    End If
    
    If g_bIsUseInsurance Then
        If KeyCode = vbKeyF9 And Shift = 0 Then 'F9
            '售保险
            If ArrayLength(m_aszInsurce) <> 0 Then
                g_oCommDialog.InitInfo m_oAUser.UserID, m_oAUser.UserName, m_oAUser.SellStationID
                g_oCommDialog.PrintInsurance m_oAUser.UserID, m_aszInsurce
                Dim aszNull() As String
                m_aszInsurce = aszNull
            Else
                MsgBox "没有客票信息，不能售保！", vbOKOnly + vbExclamation, "错误"
                cboEndStation.SetFocus
                Exit Sub
            End If
        End If
        If KeyCode = vbKeyF9 And Shift = 2 Then 'Ctrl+F7
            '补保险
            g_oCommDialog.InitInfo m_oAUser.UserID, m_oAUser.UserName, m_oAUser.SellStationID
            g_oCommDialog.RecruitInsurance m_oAUser.UserID
        End If
        If KeyCode = vbKeyF11 And Shift = 0 Then 'F11
            '快速退保险
            If MsgBox("是否快速退保？", vbInformation + vbYesNo, "快速退保") = vbYes Then
                g_oCommDialog.InitInfo m_oAUser.UserID, m_oAUser.UserName, m_oAUser.SellStationID
                Dim bIsReturned As Boolean
                bIsReturned = g_oCommDialog.FastReturnInsurance(m_oAUser.UserID)
                If bIsReturned = True Then
                    MsgBox "快速退保成功！", vbInformation, "快速退保"
                    cboEndStation.SetFocus
                Else
                    MsgBox "快速退保失败！", vbInformation, "快速退保"
                    cboEndStation.SetFocus
                End If
            End If
        End If
        If KeyCode = vbKeyF11 And Shift = 2 Then 'Ctrl+F11
            '按保单号退保险
            g_oCommDialog.InitInfo m_oAUser.UserID, m_oAUser.UserName, m_oAUser.SellStationID
            g_oCommDialog.ReturnInsurance m_oAUser.UserID
        End If
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo Error_Handle
    If KeyAscii = vbKeyEscape Then
        SetDefaultSellTicket
        If ptExtraByBus.Visible Then
            lvBusPreSell.ListItems.Clear
            
            txtBusID.SetFocus
        Else
            lvStationPreSell.ListItems.Clear
            cboEndStation.Text = ""
            cboEndStation.SetFocus
        End If
        Erase m_aszRealNameInfo
        txtSeat.Text = ""
    End If
    If KeyAscii = 45 Then
        If lvBus.ListItems.count > 0 Then
            If lvBus.SelectedItem.SubItems(ID_OffTime) = cszScrollBus Then
                Exit Sub
            End If
        End If
        KeyAscii = 0
        ChangeSeatType
        Exit Sub
    End If
    If KeyAscii = 13 And (Me.ActiveControl Is lvStation) Then
        If lvStation.ListItems.count = 0 Then
            MsgBox "没有指定的环境车次!", , "补票"
            txtBusID.SetFocus
            Exit Sub
        End If
    End If
    If KeyAscii = 13 And (Me.ActiveControl Is lvBus) And (lvSellStation.Enabled = True) Then
        lvSellStation.SetFocus
        Exit Sub
    End If
    If KeyAscii = 13 And (Me.ActiveControl Is lvStation) And (lvSellStation.Enabled = True) Then
        lvSellStation.SetFocus
        Exit Sub
    End If
    If KeyAscii = 13 And (Me.ActiveControl Is lvSellStation) Then
        txtFullSell.SetFocus
        Exit Sub
    End If
    If KeyAscii = 13 And (Me.ActiveControl Is txtReceivedMoney) Then
        txtReceivedMoney.SetFocus
        cmdSell.Enabled = False
        Exit Sub
    End If
    If KeyAscii = 13 And ((Me.ActiveControl Is txtFullSell) Or (Me.ActiveControl Is txtHalfSell)) Then
        txtReceivedMoney.SetFocus
        Exit Sub
    ElseIf KeyAscii = vbKeyReturn And (Not Me.ActiveControl Is txtReceivedMoney) And Not (Me.ActiveControl Is cboEndStation) And Not (Me.ActiveControl Is lvStation) _
        And Not (Me.ActiveControl Is txtHalfSell) And Not (Me.ActiveControl Is txtPreferentialSell) _
        And Not (Me.ActiveControl Is txtFullSell) Then
    
        SendKeys "{TAB}"
    ElseIf KeyAscii = 43 Then
        
        If Not ptExtraByBus.Visible Then
            cboEndStation.SetFocus
        Else
            txtBusID.SetFocus

        End If
    End If
    
    Exit Sub
    
Error_Handle:
    ShowErrorMsg
    End Sub

Private Sub Form_Load()
    If m_bSellStationCanSellEachOther Then
      lvSellStation.Enabled = True
    Else
      lvSellStation.Enabled = False
    End If
End Sub

Private Sub Form_Resize()
    If MDISellTicket.ActiveForm Is Me Then
        If Not Me.WindowState = vbMaximized Then Me.WindowState = vbMaximized
        fraOutLine.Left = (Me.ScaleWidth - fraOutLine.Width) / 2
        fraOutLine.Top = (Me.ScaleHeight - fraOutLine.Height) / 2
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    WriteReg cszSubKey_ExtraSellType, IIf(BusIsActive, 1, 2)
'    MDISellTicket.lblExtra.Value = vbUnchecked
    MDISellTicket.abMenuTool.Bands("mnuFunction").Tools("mnuExtraTkt").Checked = False
'    MDISellTicket.mnuExtraTkt.Checked = False
    MDISellTicket.EnableSortAndRefresh False
End Sub

Private Sub ShowBusInfo()
    On Error GoTo here
    Dim i As Integer
    
    If txtBusID.Text <> "" Then
        Dim oBus As REBus
        Set oBus = CreateObject("STReSch.REBus")
        oBus.Init m_oAUser
        oBus.Identify txtBusID.Text, m_oParam.NowDate
        
        m_aTReBusAllotInfo = oBus.GetAllotInfo(m_oAUser.SellStationID)
'
        If ArrayLength(m_aTReBusAllotInfo) > 0 Then
            m_szCheckGate = m_aTReBusAllotInfo(1).szCheckGateName
            flblOffTimeExtra.Caption = ToStandardTimeStr(m_aTReBusAllotInfo(1).dtRunTime)
        Else
            MsgBox "对应的上车站的信息不存在.", vbInformation, Me.Caption
            Exit Sub
        End If
        flblBusTypeExtra.Caption = oBus.VehicleModelName
        flblSeatCount.Caption = oBus.SaleSeat
        flblStandCount.Caption = oBus.SaleStandSeat
        If oBus.SaleSeat + oBus.SaleStandSeat > 0 Then
            lblCanExtraSell.Caption = cszAllowExtraSell
        Else
            lblCanExtraSell.Caption = cszNotAllowExtraSell
        End If
        If oBus.BusType <> TP_ScrollBus Then
            lblCanExtraSell.Caption = cszAllowExtraSell
'            Dim rsTemp As Recordset
            Dim liTemp As ListItem
            
            Set rsCountTemp = m_oSell.GetStationFromBusRs(m_oParam.NowDate, txtBusID.Text)
            
'            If rsTemp.RecordCount <> 0 Then
'               RefreshSellStation rsTemp
'            End If
                       
            lvStation.ListItems.Clear
            lvStation.Sorted = False
            Do While Not rsCountTemp.EOF
                For i = lvStation.ListItems.count To 1 Step -1
                    If lvStation.ListItems(i) = (RTrim(rsCountTemp!station_id)) Then
                        Select Case Trim(rsCountTemp!seat_type_id)
                            Case cszSeatType
                                liTemp.SubItems(ID_FullPrice2) = FormatMoney(rsCountTemp!full_price)
                                liTemp.SubItems(ID_HalfPrice2) = FormatMoney(rsCountTemp!half_price)
                                liTemp.SubItems(ID_FreePrice2) = 0
                                liTemp.SubItems(ID_PreferentialPrice21) = FormatMoney(rsCountTemp!preferential_ticket1)
                                liTemp.SubItems(ID_PreferentialPrice22) = FormatMoney(rsCountTemp!preferential_ticket2)
                                liTemp.SubItems(ID_PreferentialPrice23) = FormatMoney(rsCountTemp!preferential_ticket3)
                            Case cszBedType
                                liTemp.SubItems(ID_BedFullPrice2) = FormatMoney(rsCountTemp!full_price)
                                liTemp.SubItems(ID_BedHalfPrice2) = FormatMoney(rsCountTemp!half_price)
                                liTemp.SubItems(ID_BedFreePrice2) = 0
                                liTemp.SubItems(ID_BedPreferential21) = FormatMoney(rsCountTemp!preferential_ticket1)
                                liTemp.SubItems(ID_BedPreferential22) = FormatMoney(rsCountTemp!preferential_ticket2)
                                liTemp.SubItems(ID_BedPreferential23) = FormatMoney(rsCountTemp!preferential_ticket3)
                            Case cszAdditionalType
                                liTemp.SubItems(ID_AdditionalFullPrice2) = FormatMoney(rsCountTemp!full_price)
                                liTemp.SubItems(ID_AdditionalHalfPrice2) = FormatMoney(rsCountTemp!half_price)
                                liTemp.SubItems(ID_AdditionalFreePrice2) = 0
                                liTemp.SubItems(ID_AdditionalPreferential21) = FormatMoney(rsCountTemp!preferential_ticket1)
                                liTemp.SubItems(ID_AdditionalPreferential22) = FormatMoney(rsCountTemp!preferential_ticket2)
                                liTemp.SubItems(ID_AdditionalPreferential23) = FormatMoney(rsCountTemp!preferential_ticket3)
                        End Select
                        GoTo nextstep
                    End If
                    
                Next i
                Set liTemp = lvStation.ListItems.Add(, GetEncodedKey(RTrim(rsCountTemp!station_id)), RTrim(rsCountTemp!station_id))
                liTemp.ListSubItems.Add , , RTrim(rsCountTemp!station_name)
                liTemp.ListSubItems.Add , , Trim(rsCountTemp!bus_type)
                Select Case Trim(rsCountTemp!seat_type_id)
                    Case cszSeatType
                        liTemp.SubItems(ID_FullPrice2) = FormatMoney(rsCountTemp!full_price)
                        liTemp.SubItems(ID_HalfPrice2) = FormatMoney(rsCountTemp!half_price)
                        liTemp.SubItems(ID_FreePrice2) = 0
                        liTemp.SubItems(ID_PreferentialPrice21) = FormatMoney(rsCountTemp!preferential_ticket1)
                        liTemp.SubItems(ID_PreferentialPrice22) = FormatMoney(rsCountTemp!preferential_ticket2)
                        liTemp.SubItems(ID_PreferentialPrice23) = FormatMoney(rsCountTemp!preferential_ticket3)
                        liTemp.SubItems(ID_BedFullPrice2) = 0
                        liTemp.SubItems(ID_BedHalfPrice2) = 0
                        liTemp.SubItems(ID_BedFreePrice2) = 0
                        liTemp.SubItems(ID_BedPreferential21) = 0
                        liTemp.SubItems(ID_BedPreferential22) = 0
                        liTemp.SubItems(ID_BedPreferential23) = 0
                        liTemp.SubItems(ID_AdditionalFullPrice2) = 0
                        liTemp.SubItems(ID_AdditionalHalfPrice2) = 0
                        liTemp.SubItems(ID_AdditionalFreePrice2) = 0
                        liTemp.SubItems(ID_AdditionalPreferential21) = 0
                        liTemp.SubItems(ID_AdditionalPreferential22) = 0
                        liTemp.SubItems(ID_AdditionalPreferential23) = 0
                    Case cszBedType
                        liTemp.SubItems(ID_FullPrice2) = 0
                        liTemp.SubItems(ID_HalfPrice2) = 0
                        liTemp.SubItems(ID_FreePrice2) = 0
                        liTemp.SubItems(ID_PreferentialPrice21) = 0
                        liTemp.SubItems(ID_PreferentialPrice22) = 0
                        liTemp.SubItems(ID_PreferentialPrice23) = 0
                        liTemp.SubItems(ID_BedFullPrice2) = FormatMoney(rsCountTemp!full_price)
                        liTemp.SubItems(ID_BedHalfPrice2) = FormatMoney(rsCountTemp!half_price)
                        liTemp.SubItems(ID_BedFreePrice2) = 0
                        liTemp.SubItems(ID_BedPreferential21) = FormatMoney(rsCountTemp!preferential_ticket1)
                        liTemp.SubItems(ID_BedPreferential22) = FormatMoney(rsCountTemp!preferential_ticket2)
                        liTemp.SubItems(ID_BedPreferential23) = FormatMoney(rsCountTemp!preferential_ticket3)
                        liTemp.SubItems(ID_AdditionalFullPrice2) = 0
                        liTemp.SubItems(ID_AdditionalHalfPrice2) = 0
                        liTemp.SubItems(ID_AdditionalFreePrice2) = 0
                        liTemp.SubItems(ID_AdditionalPreferential21) = 0
                        liTemp.SubItems(ID_AdditionalPreferential22) = 0
                        liTemp.SubItems(ID_AdditionalPreferential23) = 0
                    Case cszAdditionalType
                        liTemp.SubItems(ID_FullPrice2) = 0
                        liTemp.SubItems(ID_HalfPrice2) = 0
                        liTemp.SubItems(ID_FreePrice2) = 0
                        liTemp.SubItems(ID_PreferentialPrice21) = 0
                        liTemp.SubItems(ID_PreferentialPrice22) = 0
                        liTemp.SubItems(ID_PreferentialPrice23) = 0
                        liTemp.SubItems(ID_BedFullPrice2) = 0
                        liTemp.SubItems(ID_BedHalfPrice2) = 0
                        liTemp.SubItems(ID_BedFreePrice2) = 0
                        liTemp.SubItems(ID_BedPreferential21) = 0
                        liTemp.SubItems(ID_BedPreferential22) = 0
                        liTemp.SubItems(ID_BedPreferential23) = 0
                        liTemp.SubItems(ID_AdditionalFullPrice2) = FormatMoney(rsCountTemp!full_price)
                        liTemp.SubItems(ID_AdditionalHalfPrice2) = FormatMoney(rsCountTemp!half_price)
                        liTemp.SubItems(ID_AdditionalFreePrice2) = 0
                        liTemp.SubItems(ID_AdditionalPreferential21) = FormatMoney(rsCountTemp!preferential_ticket1)
                        liTemp.SubItems(ID_AdditionalPreferential22) = FormatMoney(rsCountTemp!preferential_ticket2)
                        liTemp.SubItems(ID_AdditionalPreferential23) = FormatMoney(rsCountTemp!preferential_ticket3)
                End Select
                liTemp.Tag = MakeDisplayString(Trim(rsCountTemp!sell_station_id), Trim(rsCountTemp!sell_station_name))
                liTemp.ListSubItems.Add , , rsCountTemp!sale_ticket_quantity
                liTemp.ListSubItems.Add , , rsCountTemp!stop_sale_time
                liTemp.ListSubItems.Add , , rsCountTemp!end_station_name
nextstep:
                rsCountTemp.MoveNext
            Loop
            lvStation.Sorted = True
        Else
            lblCanExtraSell.Caption = cszNotAllowExtraSell
        End If
                
        m_bBusInfoDirty = False
    End If
    
    If lvStation.ListItems.count > 0 Then
        RefreshSellStation rsCountTemp, m_aTReBusAllotInfo
    Else
        lvSellStation.ListItems.Clear
    End If
    DoThingWhenBusChange
    Set oBus = Nothing
    On Error GoTo 0
    Exit Sub
here:
    ShowErrorMsg
    txtBusID.SetFocus
    Set oBus = Nothing
End Sub





Private Sub lvBus_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    SortListView lvBus, ColumnHeader.Index
End Sub
Private Sub lvBus_DblClick()
    If Not lvBus.SelectedItem Is Nothing Then
        Call txtFullSell_KeyPress(vbKeyReturn)
    End If
End Sub
Private Sub lvBus_GotFocus()
    lblBus.ForeColor = clActiveColor
    ShowRightSeatType
    On Error Resume Next
    If lvBus.ListItems.count = 0 Then cboEndStation.SetFocus
    
End Sub

Private Sub lvBus_ItemClick(ByVal Item As MSComctlLib.ListItem)
'    DoThingWhenBusChange
On Error GoTo here
      
        RefreshSellStation rsCountTemp, m_aTReBusAllotInfo
        ShowRightSeatType
        DoThingWhenBusChange
'        DealPrice
        flblTotalPrice.Caption = FormatMoney(m_dbTotalPrice + TotalInsurace)
On Error GoTo 0
Exit Sub
here:
    ShowErrorMsg
End Sub

Private Sub lvBus_LostFocus()
    lblBus.ForeColor = 0
    Debug.Print "lostfocus"
End Sub

Private Sub lvBusPreSell_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    SortListView lvBusPreSell, ColumnHeader.Index
End Sub

Private Sub lvBusPreSell_GotFocus()
    lblPreBus.ForeColor = clActiveColor
End Sub

Private Sub lvBusPreSell_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        If Not lvBusPreSell.SelectedItem Is Nothing Then
            lvBusPreSell.ListItems.Remove lvBusPreSell.SelectedItem.Index
        End If
    End If
End Sub

Private Sub lvBusPreSell_LostFocus()
    lblPreBus.ForeColor = 0
End Sub

Private Sub lvSellStation_ItemClick(ByVal Item As MSComctlLib.ListItem)
   cboSeatType_Change
   flblTotalPrice.Caption = FormatMoney(lvSellStation.SelectedItem.SubItems(2) + TotalInsurace)
   DealPrice
End Sub

Private Sub lvStation_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    SortListView lvStation, ColumnHeader.Index
End Sub

Private Sub lvStation_GotFocus()
    lblStation.ForeColor = clActiveColor
    ShowRightSeatType
End Sub

Private Sub lvStation_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo here
      
        RefreshSellStation rsCountTemp, m_aTReBusAllotInfo
        ShowRightSeatType
        DoThingWhenBusChange
'        DealPrice
        flblTotalPrice.Caption = FormatMoney(m_dbTotalPrice + TotalInsurace)
On Error GoTo 0
Exit Sub
here:
    ShowErrorMsg
End Sub

Private Sub lvStation_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If lvSellStation.Enabled Then
            lvSellStation.SetFocus
        Else
            txtFullSell.SetFocus
        End If
    End If
End Sub

Private Sub lvStation_LostFocus()
    lblStation.ForeColor = 0
End Sub

Private Sub lvStationPreSell_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    SortListView lvStationPreSell, ColumnHeader.Index
End Sub

Private Sub lvStationPreSell_GotFocus()
    lblPreBus.ForeColor = clActiveColor
End Sub

Private Sub lvStationPreSell_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        If Not lvStationPreSell.SelectedItem Is Nothing Then
            lvStationPreSell.ListItems.Remove lvStationPreSell.SelectedItem.Index
        End If
    End If
End Sub


Private Sub Timer1_Timer()
    Timer1.Enabled = False
'    If BusisActive Then txtBusID.SetFocus


End Sub

Private Sub lvStationPreSell_LostFocus()
    lblPreBus.ForeColor = 0
End Sub

'Private Sub tsExtraType_Click()
'    If BusisActive Then
'        ptExtraSellByStation.Visible = False
'        ptExtraByBus.Visible = True
'        picStationInfo.Visible = False
'        picBusInfo.Visible = True
'        MDISellTicket.EnableSortAndRefresh False
'        If Visible Then txtBusID.SetFocus
'    Else
'        ptExtraByBus.Visible = False
'        ptExtraSellByStation.Visible = True
'        picStationInfo.Visible = True
'        picBusInfo.Visible = False
'        MDISellTicket.EnableSortAndRefresh True
'        If Visible Then cboEndStation.SetFocus
'    End If
'    DoThingWhenBusChange
'    EnableSeatAndStand
'    SetPreSellButton
'    'EnableSellButton
'    txtReceivedMoney.Text = ""
'End Sub

Private Sub tsExtraType_TabActivate(Index As Integer)
    If Index = 1 Then
        ptExtraSellByStation.Visible = False
        ptExtraByBus.Visible = True
        lvStationPreSell.Visible = False
        lvBusPreSell.Visible = True
        MDISellTicket.EnableSortAndRefresh False
        If Visible Then txtBusID.SetFocus
        lblStation.Visible = True
        lvStation.Visible = True
        lblBus.Visible = False
        lvBus.Visible = False
        lvStation.Refresh
        
    Else
        ptExtraByBus.Visible = False
        ptExtraSellByStation.Visible = True
        lvStationPreSell.Visible = True
        lvBusPreSell.Visible = False
        MDISellTicket.EnableSortAndRefresh True
        If Visible Then cboEndStation.SetFocus
        lblStation.Visible = False
        lvStation.Visible = False
        lblBus.Visible = True
        lvBus.Visible = True
        lvBus.Refresh
    End If
    DoThingWhenBusChange
    EnableSeatAndStand
    SetPreSellButton
    'EnableSellButton
    txtReceivedMoney.Text = ""
    
    lblFullSell.Refresh
    lblHalfSell.Refresh
    lblSeatType.Refresh
    lnTicketType.Refresh
    
End Sub

Private Sub Timer2_Timer()
    Dim nExSellType As Integer
On Error GoTo here
    Timer2.Enabled = False

'    lblUser.Caption = m_oAUser.UserID & "/" & m_oAUser.UserName
    SetPreSellButton
    EnableSellButton
    RefreshPreferentialTicket
    DealDiscountAndSeat
    RefreshStation2
    EnableSeatAndStand
    SetDefaultSellTicket
    
    m_dbTotalPrice = 0
    flblTotalPrice.Caption = FormatMoney(m_dbTotalPrice + TotalInsurace)
    m_tbSeatTypeBus = m_oSell.GetMultiSeatTypeBus
    nExSellType = Val(ReadReg(cszSubKey_ExtraSellType))
    If nExSellType <> 2 Then nExSellType = 1
    '如果为车次
    If nExSellType = 1 Then
        SetBus
    Else
        SetStation
    End If
    
'    On Error GoTo 0
Exit Sub
here:
    ShowErrorMsg
End Sub

Private Sub txtBusID_Change()
    m_bBusInfoDirty = True
    txtReceivedMoney.Text = ""
End Sub


'变换售票(应为补票按钮状态)
Private Sub EnableSellButton()
    If BusIsActive Then
        If txtBusID.Text = "" Or (lvStation.SelectedItem Is Nothing) Then
            cmdSell.Enabled = False
        Else
            cmdSell.Enabled = True
        End If
    Else
        If cboEndStation.Text = "" Or (lvBus.SelectedItem Is Nothing) Then
            cmdSell.Enabled = False
        Else
            cmdSell.Enabled = True
        End If
    End If
End Sub

'变换定座相关控件的状态
Private Sub EnableSeatAndStand()
    Dim liTemp As ListItem
    Dim szStationID As String
    szStationID = cboEndStation.BoundText

    If ptExtraByBus.Visible Then '按车次补票
        If txtBusID.Text <> "" And CInt(flblSeatCount.Caption) > 0 Then
            cmdSetSeat.Enabled = True
'            chkSetSeat.Enabled = True
        Else
            cmdSetSeat.Enabled = False
'            chkSetSeat.Enabled = False
'            chkSetSeat.Value = vbUnchecked
        End If

    Else ' 按到站补票
        Set liTemp = lvBus.SelectedItem
        If liTemp Is Nothing Or szStationID = "" Then
            cmdSetSeat.Enabled = False
'            chkSetSeat.Enabled = False
'            chkSetSeat.Value = vbUnchecked

        Else
            If liTemp.SubItems(ID_SeatCount) > 0 Then
                cmdSetSeat.Enabled = True
'                chkSetSeat.Enabled = True
            Else
                cmdSetSeat.Enabled = False
'                chkSetSeat.Enabled = False
'                chkSetSeat.Value = vbUnchecked
            End If
        End If
    End If
    Set liTemp = Nothing
End Sub

'处理票价 (计算总票价？算出找钱)
Private Sub DealPrice()
    Dim sgTemp As Double
    Dim sgvalue As Double
    Dim dSum As Double
    Dim aszSeatNo() As String
    Dim TicketType() As ETicketType
    Dim TicketPrice() As Single
    Dim nSeat As Integer
    Dim i As Integer
    Dim nSameIndex As Integer
    Dim nLength As Integer
    nLength = 0
   On Error GoTo here

'    dSum = 0
    dSum = GetDealTotalPrice()
    sgTemp = 0
    If (Not lvStation.SelectedItem Is Nothing And ptExtraByBus.Visible) Or (Not lvBus.SelectedItem Is Nothing And (Not ptExtraByBus.Visible)) Then
        Dim liTemp As ListItem
        If txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text = 0 Then
            sgTemp = 0
               
            flblTotalPrice.Caption = FormatMoney(0 + dSum + TotalInsurace)
            m_dbTotalPrice = FormatMoney(0 + dSum)
            lblTotalMoney.Caption = FormatMoney(0 + m_sgTotalMoney)
        Else
            If ptExtraByBus.Visible Then
                Set liTemp = lvStation.SelectedItem
                Select Case Trim(m_aszSeatType(cboSeatType.ListIndex + 1, 1))
                    Case cszSeatType
                        For i = 1 To txtFullSell.Text
                            sgTemp = sgTemp + CDbl(liTemp.SubItems(ID_FullPrice2))
                        Next i
                        For i = 1 To txtHalfSell.Text
                            sgTemp = sgTemp + liTemp.SubItems(ID_HalfPrice2)
                        Next i
                        For i = 1 To txtPreferentialSell.Text
                            sgTemp = GetPreferentialPrice + sgTemp
                        Next i

                    Case cszBedType
                        For i = 1 To txtFullSell.Text
                            sgTemp = sgTemp + CDbl(liTemp.SubItems(ID_BedFullPrice2))
                        Next i
                        For i = 1 To txtHalfSell.Text
                            sgTemp = sgTemp + liTemp.SubItems(ID_BedHalfPrice2)
                        Next i
                        For i = 1 To txtPreferentialSell.Text
                            sgTemp = GetPreferentialPrice + sgTemp
                        Next i
                    Case cszAdditionalType
                        For i = 1 To txtFullSell.Text
                            sgTemp = sgTemp + CDbl(liTemp.SubItems(ID_AdditionalFullPrice2))
                        Next i
                        For i = 1 To txtHalfSell.Text
                            sgTemp = sgTemp + liTemp.SubItems(ID_AdditionalHalfPrice2)
                        Next i
                        For i = 1 To txtPreferentialSell.Text
                            sgTemp = GetPreferentialPrice + sgTemp
                        Next i
                End Select
            Else
                Set liTemp = lvBus.SelectedItem
                Select Case Trim(m_aszSeatType(cboSeatType.ListIndex + 1, 1))
                Case cszSeatType
                    For i = 1 To txtFullSell.Text
                        sgTemp = sgTemp + liTemp.SubItems(ID_FullPrice)
                    Next i
                    For i = 1 To txtHalfSell.Text
                        sgTemp = sgTemp + liTemp.SubItems(ID_HalfPrice)
                    Next i
                    For i = 1 To txtPreferentialSell.Text
                        sgTemp = sgTemp + GetPreferentialPrice(False)
                    Next i
                Case cszBedType
                    For i = 1 To txtFullSell.Text
                        sgTemp = sgTemp + liTemp.SubItems(ID_BedFullPrice)
                    Next i
                    For i = 1 To txtHalfSell.Text
                        sgTemp = sgTemp + liTemp.SubItems(ID_BedHalfPrice)
                    Next i
                    For i = 1 To txtPreferentialSell.Text
                        sgTemp = sgTemp + GetPreferentialPrice(False)
                    Next i
                Case cszAdditionalType
                    For i = 1 To txtFullSell.Text
                        sgTemp = sgTemp + liTemp.SubItems(ID_AdditionalFullPrice)
                    Next i
                    For i = 1 To txtHalfSell.Text
                        sgTemp = sgTemp + liTemp.SubItems(ID_AdditionalHalfPrice)
                    Next i
                    For i = 1 To txtPreferentialSell.Text
                        sgTemp = sgTemp + GetPreferentialPrice(False)
                    Next i
                End Select

            End If
            Set liTemp = Nothing
            sgTemp = sgTemp * CDbl(txtDiscount.Text)
            flblTotalPrice.Caption = FormatMoney(sgTemp + dSum + TotalInsurace)
            m_dbTotalPrice = FormatMoney(sgTemp + dSum)
            lblTotalMoney.Caption = FormatMoney(sgTemp + m_sgTotalMoney)
            
        
        End If
    Else
        flblTotalPrice.Caption = FormatMoney(0 + dSum + TotalInsurace)
        m_dbTotalPrice = FormatMoney(0 + dSum)
        lblTotalMoney.Caption = FormatMoney(0 + m_sgTotalMoney)

    End If
    'flblRestMoney.Caption = FormatMoney(CDbl(txtReceivedMoney.Text) - CDbl(lblTotalPrice))
    If txtReceivedMoney.Text = "0" And Not Me.ActiveControl Is txtReceivedMoney Then txtReceivedMoney.Text = ""
    If Left(txtReceivedMoney.Text, 1) = "." Then txtReceivedMoney.Text = "0" & txtReceivedMoney.Text
    If txtReceivedMoney.Text = "" Then
       sgvalue = 0
    Else
       sgvalue = CDbl(txtReceivedMoney.Text)
    End If
    If sgvalue <> 0 And sgvalue - CDbl(lblTotalMoney) >= 0 And CDbl(lblTotalMoney.Caption) >= 0 Then
       flblRestMoney.Caption = FormatMoney(CDbl(txtReceivedMoney.Text) - CDbl(flblTotalPrice.Caption))
       cmdSell.Enabled = True
    Else
       flblRestMoney.Caption = ""
    End If
    On Error GoTo 0
    Exit Sub
here:
    ShowErrorMsg
End Sub


'当选中的车次改变之后要做的事
Private Sub DoThingWhenBusChange()
On Error GoTo here
    If (Not lvStation.SelectedItem Is Nothing And ptExtraByBus.Visible) Or (Not lvBus.SelectedItem Is Nothing And (Not ptExtraByBus.Visible)) Then
        Dim liTemp As ListItem
        'Set liTemp = lvStation.SelectedItem
        
        If ptExtraByBus.Visible Then
            Set liTemp = lvStation.SelectedItem
            lblSinglePrice.Caption = liTemp.SubItems(ID_FullPrice2) & "/" & liTemp.SubItems(ID_HalfPrice2)
            flblLimitedCount.Caption = GetStationLimitedCountStr(liTemp.SubItems(ID_LimitedCount2))
            flblLimitedTime.Caption = GetStationLimitedTimeStr(liTemp.SubItems(ID_LimitedTime2), CDate(flblOffTimeExtra.Caption), m_oParam.NowDate)
        Else
            Set liTemp = lvBus.SelectedItem
            lblSinglePrice.Caption = liTemp.SubItems(ID_FullPrice) & "/" & liTemp.SubItems(ID_HalfPrice)
            flblLimitedCount2.Caption = GetStationLimitedCountStr(liTemp.SubItems(ID_LimitedCount))
            flblLimitedTime2.Caption = GetStationLimitedTimeStr(liTemp.SubItems(ID_LimitedTime), CDate(liTemp.SubItems(ID_OffTime)), m_oParam.NowDate)
            'flblStandCount2.Caption = liTemp.subitems(ID_StandCount)
        End If
        Set liTemp = Nothing
    Else
    
        lblSinglePrice.Caption = FormatMoney(0) & "/" & FormatMoney(0)
        If ptExtraByBus.Visible Then
            flblLimitedCount.Caption = ""
            flblLimitedTime.Caption = ""
        Else
            flblLimitedCount2.Caption = ""
            flblLimitedTime2.Caption = ""
            flblStandCount2.Caption = ""
        End If
    End If
    
    DealPrice
    EnableSeatAndStand
    'EnableSellButton
    On Error GoTo 0
    Exit Sub
    
here:
    ShowErrorMsg
    
End Sub



Private Sub txtBusID_GotFocus()
    lblScheduleExtra.ForeColor = clActiveColor
    txtBusID.SelStart = 0
    txtBusID.SelLength = Len(txtBusID.Text)
End Sub

Private Sub txtBusID_LostFocus()
    lblScheduleExtra.ForeColor = 0
    ShowBusInfo
    EnableSellButton
End Sub

Private Sub txtDiscount_KeyPress(KeyAscii As Integer)
    If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = 8 Or (KeyAscii = 46 And InStr(txtDiscount.Text, ".") = 0) Then
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtDiscount_LostFocus()
    lblDiscount.ForeColor = 0
    If txtDiscount.Text <> "" Then
        If Trim(txtDiscount.Text) = "." Then
            txtDiscount.SetFocus
            Exit Sub
        End If
        If Left(txtDiscount.Text, 1) = "." And Len(txtDiscount.Text) > 1 Then
            txtDiscount.Text = "0" & txtDiscount.Text
        End If
        If CSng(txtDiscount.Text) > 1 Then
            MsgBox "折扣率不能大于 1", vbInformation, "提示"
            txtDiscount.SetFocus
        End If
    Else
        txtDiscount.Text = 1
    End If
    DealPrice
End Sub

Private Sub txtDiscount_Validate(Cancel As Boolean)
If txtDiscount.Text = "" Then
    Cancel = True
Else
    If CSng(txtDiscount.Text) > 1 Then
        Cancel = True
    End If
End If
End Sub



Private Sub txtFullSell_GotFocus()
    lblFullSell.ForeColor = clActiveColor
End Sub

Private Sub txtFullSell_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        txtHalfSell.SetFocus
    End If
End Sub

Private Sub txtFullSell_KeyPress(KeyAscii As Integer)

On Error GoTo here
    If KeyAscii = 13 Then
        If txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text <> 0 And Not (lvBus.SelectedItem Is Nothing) Then
            If lvBus.SelectedItem.SubItems(ID_OffTime) <> cszScrollBus Then

'                If txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text > lvBus.SelectedItem.SubItems(ID_SeatCount) Then
'                    MsgBox "该车次座位已不够！" & vbCrLf & "座位只剩 " & lvBus.SelectedItem.SubItems(ID_SeatCount), vbInformation, "售票台"
'                ElseIf cboSeatType.ListIndex = 0 And txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text > lvBus.SelectedItem.SubItems(ID_SeatTypeCount) Then
'                    MsgBox "该车次[普通]座位已不够！" & vbCrLf & "  [普通]座位只剩 " & lvBus.SelectedItem.SubItems(ID_SeatTypeCount), vbInformation, "售票台"
'                ElseIf cboSeatType.ListIndex = 1 And txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text > lvBus.SelectedItem.SubItems(ID_BedTypeCount) Then
'                    MsgBox "该车次[卧铺]座位已不够！" & vbCrLf & "  [卧铺]座位只剩 " & lvBus.SelectedItem.SubItems(ID_BedTypeCount), vbInformation, "售票台"
'                ElseIf cboSeatType.ListIndex = 2 And txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text > lvBus.SelectedItem.SubItems(ID_AdditionalCount) Then
'                    MsgBox "该车次[加座]座位已不够！" & vbCrLf & "  [加座]座位只剩 " & lvBus.SelectedItem.SubItems(ID_AdditionalCount), vbInformation, "售票台"
'                    txtFullSell.SetFocus
'                    KeyAscii = 0
'                Else
                    cmdPreSell_Click
'                    cmdSell.SetFocus
                    'cmdPreSell.SetFocus
                    txtReceivedMoney.SetFocus
'                End If
            Else
                cmdPreSell_Click
'                    cmdSell.SetFocus
                txtReceivedMoney.SetFocus
            End If
        End If
    End If
    
    Exit Sub
here:
   ShowErrorMsg
End Sub



Private Sub txtFullSell_LostFocus()
    lblFullSell.ForeColor = 0
End Sub

Private Sub txtHalfSell_GotFocus()
    lblHalfSell.ForeColor = clActiveColor
End Sub

Private Sub txtHalfSell_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyUp
            txtFullSell.SetFocus
        Case vbKeyDown
            If cboPreferentialTicket.Enabled Then cboPreferentialTicket.SetFocus
    End Select
End Sub

Private Sub txtHalfSell_KeyPress(KeyAscii As Integer)
'On Error GoTo Here
'    If KeyAscii = 13 And (Me.ActiveControl Is lvBus) Then
'      If txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text > lvBus.SelectedItem.SubItems(ID_SeatCount) Then
'                MsgBox "该车次座位已不够！" & vbCrLf & "座位只剩 " & lvBus.SelectedItem.SubItems(ID_SeatCount), vbInformation, "售票台"
'             ElseIf cboSeatType.ListIndex = 0 And txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text > lvBus.SelectedItem.SubItems(ID_SeatTypeCount) Then
'                MsgBox "该车次[普通]座位已不够！" & vbCrLf & "  [普通]座位只剩 " & lvBus.SelectedItem.SubItems(ID_SeatTypeCount), vbInformation, "售票台"
'             ElseIf cboSeatType.ListIndex = 1 And txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text > lvBus.SelectedItem.SubItems(ID_BedTypeCount) Then
'                MsgBox "该车次[卧铺]座位已不够！" & vbCrLf & "  [卧铺]座位只剩 " & lvBus.SelectedItem.SubItems(ID_BedTypeCount), vbInformation, "售票台"
'             ElseIf cboSeatType.ListIndex = 2 And txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text > lvBus.SelectedItem.SubItems(ID_AdditionalCount) Then
'                MsgBox "该车次[加座]座位已不够！" & vbCrLf & "  [加座]座位只剩 " & lvBus.SelectedItem.SubItems(ID_AdditionalCount), vbInformation, "售票台"
'             txtHalfSell.SetFocus
'             KeyAscii = 0
'      Else
'          cmdPreSell_Click
'          txtReceivedMoney.SetFocus
'      End If
'    End If
'
'On Error GoTo 0
'Exit Sub
'Here:
'   ShowErrorMsg
On Error GoTo here
    If KeyAscii = 13 Then
        If txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text <> 0 And Not (lvBus.SelectedItem Is Nothing) Then
            If lvBus.SelectedItem.SubItems(ID_OffTime) <> cszScrollBus Then
'                If txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text > lvBus.SelectedItem.SubItems(ID_SeatCount) Then
'                    MsgBox "该车次座位已不够！" & vbCrLf & "座位只剩 " & lvBus.SelectedItem.SubItems(ID_SeatCount), vbInformation, "售票台"
'                ElseIf cboSeatType.ListIndex = 0 And txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text > lvBus.SelectedItem.SubItems(ID_SeatTypeCount) Then
'                    MsgBox "该车次[普通]座位已不够！" & vbCrLf & "  [普通]座位只剩 " & lvBus.SelectedItem.SubItems(ID_SeatTypeCount), vbInformation, "售票台"
'                ElseIf cboSeatType.ListIndex = 1 And txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text > lvBus.SelectedItem.SubItems(ID_BedTypeCount) Then
'                    MsgBox "该车次[卧铺]座位已不够！" & vbCrLf & "  [卧铺]座位只剩 " & lvBus.SelectedItem.SubItems(ID_BedTypeCount), vbInformation, "售票台"
'                ElseIf cboSeatType.ListIndex = 2 And txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text > lvBus.SelectedItem.SubItems(ID_AdditionalCount) Then
'                    MsgBox "该车次[加座]座位已不够！" & vbCrLf & "  [加座]座位只剩 " & lvBus.SelectedItem.SubItems(ID_AdditionalCount), vbInformation, "售票台"
'                    txtHalfSell.SetFocus
'                    KeyAscii = 0
'                End If
                If KeyAscii = 13 Then
                    cmdPreSell_Click
                    txtReceivedMoney.SetFocus
'                    cmdSell.SetFocus
                End If
            Else
                cmdPreSell_Click
'                    cmdSell.SetFocus
                txtReceivedMoney.SetFocus
            End If
        End If
    End If
    
Exit Sub
here:
   ShowErrorMsg
End Sub

Private Sub txtHalfSell_LostFocus()
    lblHalfSell.ForeColor = 0
End Sub

Private Sub txtPreferentialSell_Change()
    'EnableSellButton
    EnableSeatAndStand
    SetPreSellButton
    DealPrice
    flblTotalPrice.Caption = FormatMoney(m_dbTotalPrice + TotalInsurace)
End Sub

Private Sub txtFullSell_Change()
    'EnableSellButton
    SetPreSellButton
    EnableSeatAndStand
    DealPrice
    flblTotalPrice.Caption = FormatMoney(m_dbTotalPrice + TotalInsurace)
End Sub

Private Sub txtHalfSell_Change()
    'EnableSellButton
    EnableSeatAndStand
    SetPreSellButton
    DealPrice
    flblTotalPrice.Caption = FormatMoney(m_dbTotalPrice + TotalInsurace)
End Sub

Private Sub txtPreferentialSell_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyUp
            txtHalfSell.SetFocus
        Case vbKeyLeft
            cboPreferentialTicket.SetFocus
    End Select
End Sub

Private Sub txtPreferentialSell_KeyPress(KeyAscii As Integer)
On Error GoTo here
    If KeyAscii = 13 Then
        If txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text <> 0 And Not (lvBus.SelectedItem Is Nothing) Then
            If lvBus.SelectedItem.SubItems(ID_OffTime) <> cszScrollBus Then

'                If txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text > lvBus.SelectedItem.SubItems(ID_SeatCount) Then
'                    MsgBox "该车次座位已不够！" & vbCrLf & "座位只剩 " & lvBus.SelectedItem.SubItems(ID_SeatCount), vbInformation, "售票台"
'                ElseIf cboSeatType.ListIndex = 0 And txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text > lvBus.SelectedItem.SubItems(ID_SeatTypeCount) Then
'                    MsgBox "该车次[普通]座位已不够！" & vbCrLf & "  [普通]座位只剩 " & lvBus.SelectedItem.SubItems(ID_SeatTypeCount), vbInformation, "售票台"
'                ElseIf cboSeatType.ListIndex = 1 And txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text > lvBus.SelectedItem.SubItems(ID_BedTypeCount) Then
'                    MsgBox "该车次[卧铺]座位已不够！" & vbCrLf & "  [卧铺]座位只剩 " & lvBus.SelectedItem.SubItems(ID_BedTypeCount), vbInformation, "售票台"
'                ElseIf cboSeatType.ListIndex = 2 And txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text > lvBus.SelectedItem.SubItems(ID_AdditionalCount) Then
'                    MsgBox "该车次[加座]座位已不够！" & vbCrLf & "  [加座]座位只剩 " & lvBus.SelectedItem.SubItems(ID_AdditionalCount), vbInformation, "售票台"
'                    txtPreferentialSell.SetFocus
'                    KeyAscii = 0
'                Else
                cmdPreSell_Click
                    txtReceivedMoney.SetFocus
'                cmdSell.SetFocus
'                End If
            Else
                cmdPreSell_Click
                txtReceivedMoney.SetFocus
'                cmdSell.SetFocus
            End If
        End If
    End If
    Exit Sub
here:
   ShowErrorMsg
End Sub

Private Sub txtReceivedMoney_Change()
'    Dim sgvalue As Double
'On Error GoTo Here
'
'
'    DealPrice
'    CalReceiveMoney
'    Exit Sub
'Here:
'    ShowErrorMsg

 Dim sgvalue As Double
On Error GoTo here
    If txtReceivedMoney.Text = "0" And Not Me.ActiveControl Is txtReceivedMoney Then txtReceivedMoney.Text = ""
    If Left(txtReceivedMoney.Text, 1) = "." Then txtReceivedMoney.Text = "0" & txtReceivedMoney.Text
    If txtReceivedMoney.Text = "" Then
       sgvalue = 0
    Else
       sgvalue = CDbl(txtReceivedMoney.Text)
    End If
    If sgvalue <> 0 Then
       flblRestMoney.Caption = FormatMoney(CDbl(txtReceivedMoney.Text) - CDbl(flblTotalPrice.Caption))
       cmdSell.Enabled = True
    Else
       flblRestMoney.Caption = ""
    End If
    CalReceiveMoney
    
On Error GoTo 0
Exit Sub
here:
    ShowErrorMsg
    
End Sub

'设置缺省的补票张数设置
Private Sub SetDefaultSellTicket()
    txtFullSell.Text = 1
    txtHalfSell.Text = 0
    txtPreferentialSell.Text = 0
    
    If chkSetSeat.Enabled Then chkSetSeat.Value = 0 '不定座位
    
    If txtReceivedMoney.Enabled Then '所收钞票为0
        txtReceivedMoney.Text = ""
        DealPrice
    End If
    chkInsurance.Value = vbUnchecked
End Sub

'得到此次售票的相应序号的座号
Private Function SelfGetSeatNo(pnIndex As Integer) As String
    If chkSetSeat.Enabled = False Then '如果站票选中,则为站票
        SelfGetSeatNo = "ST"
    ElseIf chkSetSeat.Enabled And chkSetSeat.Value = 1 Then '如果定座选中,则得到相应的座号
        SelfGetSeatNo = GetSeatNo(txtSeat.Text, pnIndex)
    Else '否则为自动座位号
        SelfGetSeatNo = ""
    End If
End Function
Private Function SelfGetSeatNo12(pnIndex As Integer, SetSeatEnable As Boolean, SetSeatValue As Integer, pszSeatNo As String) As String
    If SetSeatEnable = False Then '如果站票选中,则为站票
        SelfGetSeatNo12 = "ST"
    ElseIf SetSeatEnable And SetSeatValue = 1 Then '如果定座选中,则得到相应的座号
        SelfGetSeatNo12 = GetSeatNo(pszSeatNo, pnIndex)
    Else '否则为自动座位号
        SelfGetSeatNo12 = ""
    End If
End Function

Private Sub txtReceivedMoney_GotFocus()
lblReceivedMoney.ForeColor = clActiveColor
End Sub

Private Sub txtReceivedMoney_KeyDown(KeyCode As Integer, Shift As Integer)
Dim nCount As Integer

For nCount = 1 To Len(txtReceivedMoney.Text)
    If Mid(txtReceivedMoney.Text, nCount, 1) = "." Then
       m_bPointCount = True
       Exit For
    End If
Next nCount
End Sub

Private Sub txtReceivedMoney_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 8 Then
    If KeyAscii < 48 Or KeyAscii > 57 Then
       If m_bPointCount = True And KeyAscii = 46 Then
          KeyAscii = 0
       ElseIf KeyAscii <> 46 Then
          KeyAscii = 0
       End If
    End If
Else
    If KeyAscii = 13 Then
        If lvStationPreSell.ListItems.count > 0 Then
           cmdSell_Click
        End If
        If ptExtraByBus.Visible Then
            txtBusID.SetFocus
        Else
            cboEndStation.SetFocus
        End If
    End If
End If

m_bPointCount = False
End Sub

Private Sub txtReceivedMoney_KeyUp(KeyCode As Integer, Shift As Integer)
Dim sgvalue As Double
    
    If txtReceivedMoney.Text = "" Then
       sgvalue = 0
    Else
       sgvalue = CDbl(txtReceivedMoney.Text)
    End If
    If sgvalue - CDbl(flblTotalPrice) <= 0 Then txtReceivedMoney.SetFocus
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


'根据当天和到站代码刷新车次信息
'pbForce表示是否强行刷新 (不管预售天数和到站是否改变)
Private Sub RefreshBus(Optional pbForce As Boolean = False, Optional pszSellStationID As String = "")
    Dim szStationID As String
'    Dim rsTemp As Recordset
    Dim liTemp As ListItem
    Dim lForeColor As OLE_COLOR
    Dim nBusType As EBusType
    Dim i As Integer
'    Dim varBookmark As Variant

    On Error GoTo here
    szStationID = RTrim(cboEndStation.BoundText)
    
    If cboEndStation.Changed Or pbForce Then
        
        If szStationID <> "" Then
            If m_szRegValue = "" Then
                Set rsCountTemp = m_oSell.GetBusRs(m_oParam.NowDate, szStationID, True)
            Else
                Set rsCountTemp = m_oSell.GetBusRsEx(m_oParam.NowDate, szStationID, m_szRegValue, True)
            End If

            lvBus.ListItems.Clear
            lvBus.Refresh
            lvBus.Sorted = False
            Do While Not rsCountTemp.EOF
                For i = lvBus.ListItems.count To 1 Step -1
                    If RTrim(rsCountTemp!bus_id) = lvBus.ListItems(i) Then ' And Format(rsCountTemp!bus_date, "yyyy-mm-dd") = CDate(m_oParam.NowDate) Then
                        GoTo nextstep
                    End If
                Next i
                If rsCountTemp!status = ST_BusStopped Or rsCountTemp!status = ST_BusMergeStopped Or rsCountTemp!status = ST_BusSlitpStop Then
                    lForeColor = m_lStopBusColor
                    
                Else
                    lForeColor = m_lNormalBusColor
                    Set liTemp = lvBus.ListItems.Add(, "A" & RTrim(rsCountTemp!bus_id), RTrim(rsCountTemp!bus_id))   '车次代码"A" & RTrim(rsCountTemp！bus_id)
                End If
                
                nBusType = rsCountTemp!bus_type
                
                                
                If lForeColor <> m_lStopBusColor Then
                    liTemp.ForeColor = lForeColor
'                 varBookmark = rsCountTemp.Bookmark
'                   If pszSellStationID = "" Then
'                    If rsCountTemp.RecordCount <> 0 Then
'                       RefreshSellStation rsCountTemp
'                    End If
'                   End If
'                 rsCountTemp.Bookmark = varBookmark
                    If nBusType <> TP_ScrollBus Then
                        liTemp.SubItems(ID_BusType) = Trim(rsCountTemp!bus_type)
                        liTemp.SubItems(ID_OffTime) = Format(rsCountTemp!BusStartTime, "hh:mm")
                        
                    Else
                        liTemp.SubItems(ID_VehicleModel) = cszScrollBus
                        liTemp.SubItems(ID_OffTime) = cszScrollBus
                        
                    End If
                    liTemp.SubItems(ID_RouteName) = RTrim(rsCountTemp!route_name)
                    liTemp.SubItems(ID_EndStation) = RTrim(rsCountTemp!end_station_name)
                    liTemp.SubItems(ID_TotalSeat) = rsCountTemp!total_seat
                            If IsDate(liTemp.SubItems(ID_OffTime)) Then
                                If g_bIsBookValid And DateAdd("n", -g_nBookTime, liTemp.SubItems(ID_OffTime)) < Time Then
                                    '如果车次日期为当天,且已过预定时限,则将预定人数加到可售张数上面.
                                    liTemp.SubItems(ID_BookCount) = 0
                                    liTemp.SubItems(ID_SeatCount) = rsCountTemp!sale_seat_quantity + rsCountTemp!book_count
                                    
                                Else
                                    liTemp.SubItems(ID_BookCount) = rsCountTemp!book_count
                                    liTemp.SubItems(ID_SeatCount) = rsCountTemp!sale_seat_quantity
                                End If
                            Else
                            
                                liTemp.SubItems(ID_BookCount) = rsCountTemp!book_count
                                liTemp.SubItems(ID_SeatCount) = rsCountTemp!sale_seat_quantity
                            End If
                    liTemp.SubItems(ID_SeatTypeCount) = rsCountTemp!seat_remain
                    liTemp.SubItems(ID_BedTypeCount) = rsCountTemp!bed_remain
                    liTemp.SubItems(ID_AdditionalCount) = rsCountTemp!additional_remain
                    liTemp.SubItems(ID_VehicleModel) = rsCountTemp!vehicle_type_name
                    
                    Select Case Trim(rsCountTemp!seat_type_id)
                    Case cszSeatType
                        liTemp.SubItems(ID_FullPrice) = FormatMoney(rsCountTemp!full_price)
                        liTemp.SubItems(ID_HalfPrice) = FormatMoney(rsCountTemp!half_price)
                        liTemp.SubItems(ID_FreePrice) = 0
                        liTemp.SubItems(ID_PreferentialPrice1) = FormatMoney(rsCountTemp!preferential_ticket1)
                        liTemp.SubItems(ID_PreferentialPrice2) = FormatMoney(rsCountTemp!preferential_ticket2)
                        liTemp.SubItems(ID_PreferentialPrice3) = FormatMoney(rsCountTemp!preferential_ticket3)
                        liTemp.SubItems(ID_BedFullPrice) = 0
                        liTemp.SubItems(ID_BedHalfPrice) = 0
                        liTemp.SubItems(ID_BedFreePrice) = 0
                        liTemp.SubItems(ID_BedPreferentialPrice1) = 0
                        liTemp.SubItems(ID_BedPreferentialPrice2) = 0
                        liTemp.SubItems(ID_BedPreferentialPrice3) = 0
                        liTemp.SubItems(ID_AdditionalFullPrice) = 0
                        liTemp.SubItems(ID_AdditionalHalfPrice) = 0
                        liTemp.SubItems(ID_AdditionalFreePrice) = 0
                        liTemp.SubItems(ID_AdditionalPreferential1) = 0
                        liTemp.SubItems(ID_AdditionalPreferential2) = 0
                        liTemp.SubItems(ID_AdditionalPreferential3) = 0
                    Case cszBedType
                        liTemp.SubItems(ID_FullPrice) = 0
                        liTemp.SubItems(ID_HalfPrice) = 0
                        liTemp.SubItems(ID_FreePrice) = 0
                        liTemp.SubItems(ID_PreferentialPrice1) = 0
                        liTemp.SubItems(ID_PreferentialPrice2) = 0
                        liTemp.SubItems(ID_PreferentialPrice3) = 0
                        liTemp.SubItems(ID_BedFullPrice) = FormatMoney(rsCountTemp!full_price)
                        liTemp.SubItems(ID_BedHalfPrice) = FormatMoney(rsCountTemp!half_price)
                        liTemp.SubItems(ID_BedFreePrice) = 0
                        liTemp.SubItems(ID_BedPreferentialPrice1) = FormatMoney(rsCountTemp!preferential_ticket1)
                        liTemp.SubItems(ID_BedPreferentialPrice2) = FormatMoney(rsCountTemp!preferential_ticket2)
                        liTemp.SubItems(ID_BedPreferentialPrice3) = FormatMoney(rsCountTemp!preferential_ticket3)
                        liTemp.SubItems(ID_AdditionalFullPrice) = 0
                        liTemp.SubItems(ID_AdditionalHalfPrice) = 0
                        liTemp.SubItems(ID_AdditionalFreePrice) = 0
                        liTemp.SubItems(ID_AdditionalPreferential1) = 0
                        liTemp.SubItems(ID_AdditionalPreferential2) = 0
                        liTemp.SubItems(ID_AdditionalPreferential3) = 0
                    Case cszAdditionalType
                        liTemp.SubItems(ID_FullPrice) = 0
                        liTemp.SubItems(ID_HalfPrice) = 0
                        liTemp.SubItems(ID_FreePrice) = 0
                        liTemp.SubItems(ID_PreferentialPrice1) = 0
                        liTemp.SubItems(ID_PreferentialPrice2) = 0
                        liTemp.SubItems(ID_PreferentialPrice3) = 0
                        liTemp.SubItems(ID_BedFullPrice) = 0
                        liTemp.SubItems(ID_BedHalfPrice) = 0
                        liTemp.SubItems(ID_BedFreePrice) = 0
                        liTemp.SubItems(ID_BedPreferentialPrice1) = 0
                        liTemp.SubItems(ID_BedPreferentialPrice2) = 0
                        liTemp.SubItems(ID_BedPreferentialPrice3) = 0
                        liTemp.SubItems(ID_AdditionalFullPrice) = FormatMoney(rsCountTemp!full_price)
                        liTemp.SubItems(ID_AdditionalHalfPrice) = FormatMoney(rsCountTemp!half_price)
                        liTemp.SubItems(ID_AdditionalFreePrice) = 0
                        liTemp.SubItems(ID_AdditionalPreferential1) = FormatMoney(rsCountTemp!preferential_ticket1)
                        liTemp.SubItems(ID_AdditionalPreferential2) = FormatMoney(rsCountTemp!preferential_ticket2)
                        liTemp.SubItems(ID_AdditionalPreferential3) = FormatMoney(rsCountTemp!preferential_ticket3)
                    End Select
                    '以下几列不显示出来，只是将其存储，以备后用
                    liTemp.SubItems(ID_LimitedCount) = rsCountTemp!sale_ticket_quantity
                    liTemp.SubItems(ID_LimitedTime) = rsCountTemp!stop_sale_time
                    liTemp.SubItems(ID_BusType1) = nBusType
                    liTemp.SubItems(ID_CheckGate) = rsCountTemp!check_gate_id
                    
                    liTemp.SubItems(ID_RealName) = Trim(rsCountTemp!id_card)
                    
                    'liTemp.SubItems(ID_StandCount) = rsCountTemp！sale_stand_seat_quantity
                  liTemp.Tag = MakeDisplayString(Trim(rsCountTemp!sell_station_id), Trim(rsCountTemp!sell_station_name))
                End If
                
nextstep:
                rsCountTemp.MoveNext
            Loop
'        If lvBus.ListItems.count > 0 Then
'           RefreshSellStation lvBus.SelectedItem.Text
'        Else
'           lvSellStation.ListItems.Clear
'        End If
            lvBus.Sorted = True
        Else
            lvBus.ListItems.Clear
        End If
    End If

    If lvBus.ListItems.count > 0 Then
        lvBus.SortKey = MDISellTicket.GetSortKey() - 1
        lvBus.Sorted = True
        lvBus.ListItems(1).Selected = True
        lvBus.ListItems(1).EnsureVisible
    End If

'    调用车次改变要进行相应操作的方法
   If lvBus.ListItems.count > 0 Then
      RefreshSellStation rsCountTemp, m_aTReBusAllotInfo
   Else
      lvSellStation.ListItems.Clear
   End If
    DoThingWhenBusChange
    Set liTemp = Nothing
'    Set rsTemp = Nothing

    On Error GoTo 0
    Exit Sub
here:
    ShowErrorMsg
    Set liTemp = Nothing
    Set rsCountTemp = Nothing
End Sub

Private Sub lvSellStation_GotFocus()
   lblSellStation.ForeColor = clActiveColor
   cboSeatType_Change
   If lvSellStation.ListItems.count > 0 Then
     flblTotalPrice.Caption = FormatMoney(lvSellStation.SelectedItem.SubItems(2) + TotalInsurace)
   End If
   DealPrice
End Sub

Private Sub lvSellStation_LostFocus()
    lblSellStation.ForeColor = 0
End Sub

'座位类型改变时,刷新相应的票价
Private Sub RefreshBusStation(rsTemp As Recordset, SellStationID As String, SeatTypeID As String)
  Dim i As Integer
  Dim szBusID As String
  On Error GoTo err:
    If lvBus.Visible = True Then
      szBusID = Trim(lvBus.SelectedItem.Text)
    End If
    If lvStation.Visible = True Then
      szBusID = Trim(txtBusID.Text)
    End If
     If rsTemp.RecordCount = 0 Then
        Exit Sub
     End If
    rsTemp.MoveFirst
    For i = 1 To rsTemp.RecordCount
    If lvBus.Visible = True Then
     If Trim(rsTemp!bus_id) = szBusID Then
        If Trim(rsTemp!sell_station_id) = Trim(SellStationID) Then
           If Trim("0" + SeatTypeID) = Trim(rsTemp!seat_type_id) Then
                  Select Case Trim(rsTemp!seat_type_id)
                  Case cszSeatType
                      lvBus.SelectedItem.SubItems(ID_FullPrice) = FormatMoney(rsTemp!full_price)
                      lvBus.SelectedItem.SubItems(ID_HalfPrice) = FormatMoney(rsTemp!half_price)
                      lvBus.SelectedItem.SubItems(ID_PreferentialPrice1) = FormatMoney(rsTemp!preferential_ticket1)
                      lvBus.SelectedItem.SubItems(ID_PreferentialPrice2) = FormatMoney(rsTemp!preferential_ticket2)
                      lvBus.SelectedItem.SubItems(ID_PreferentialPrice3) = FormatMoney(rsTemp!preferential_ticket3)
                  Case cszBedType
                      lvBus.SelectedItem.SubItems(ID_BedFullPrice) = FormatMoney(rsTemp!full_price)
                      lvBus.SelectedItem.SubItems(ID_BedHalfPrice) = FormatMoney(rsTemp!half_price)
                      lvBus.SelectedItem.SubItems(ID_BedPreferentialPrice1) = FormatMoney(rsTemp!preferential_ticket1)
                      lvBus.SelectedItem.SubItems(ID_BedPreferentialPrice2) = FormatMoney(rsTemp!preferential_ticket2)
                      lvBus.SelectedItem.SubItems(ID_BedPreferentialPrice3) = FormatMoney(rsTemp!preferential_ticket3)
                  Case cszAdditionalType
                      lvBus.SelectedItem.SubItems(ID_AdditionalFullPrice) = FormatMoney(rsTemp!full_price)
                      lvBus.SelectedItem.SubItems(ID_AdditionalHalfPrice) = FormatMoney(rsTemp!half_price)
                      lvBus.SelectedItem.SubItems(ID_AdditionalPreferential1) = FormatMoney(rsTemp!preferential_ticket1)
                      lvBus.SelectedItem.SubItems(ID_AdditionalPreferential2) = FormatMoney(rsTemp!preferential_ticket2)
                      lvBus.SelectedItem.SubItems(ID_AdditionalPreferential3) = FormatMoney(rsTemp!preferential_ticket3)
                  End Select
                End If
              End If
            End If
         Else
                  Select Case Trim(rsTemp!seat_type_id)
                  Case cszSeatType
                      lvStation.SelectedItem.SubItems(ID_FullPrice) = FormatMoney(rsTemp!full_price)
                      lvStation.SelectedItem.SubItems(ID_HalfPrice) = FormatMoney(rsTemp!half_price)
                      lvStation.SelectedItem.SubItems(ID_PreferentialPrice1) = FormatMoney(rsTemp!preferential_ticket1)
                      lvStation.SelectedItem.SubItems(ID_PreferentialPrice2) = FormatMoney(rsTemp!preferential_ticket2)
                      lvStation.SelectedItem.SubItems(ID_PreferentialPrice3) = FormatMoney(rsTemp!preferential_ticket3)
                  Case cszBedType
                      lvStation.SelectedItem.SubItems(ID_BedFullPrice) = FormatMoney(rsTemp!full_price)
                      lvStation.SelectedItem.SubItems(ID_BedHalfPrice) = FormatMoney(rsTemp!half_price)
                      lvStation.SelectedItem.SubItems(ID_BedPreferentialPrice1) = FormatMoney(rsTemp!preferential_ticket1)
                      lvStation.SelectedItem.SubItems(ID_BedPreferentialPrice2) = FormatMoney(rsTemp!preferential_ticket2)
                      lvStation.SelectedItem.SubItems(ID_BedPreferentialPrice3) = FormatMoney(rsTemp!preferential_ticket3)
                  Case cszAdditionalType
                      lvStation.SelectedItem.SubItems(ID_AdditionalFullPrice) = FormatMoney(rsTemp!full_price)
'                      lvStation.SelectedItem.SubItems(ID_AdditionalHalfPrice) = rsTemp!half_price
'                      lvStation.SelectedItem.SubItems(ID_AdditionalPreferential1) = rsTemp!preferential_ticket1
'                      lvStation.SelectedItem.SubItems(ID_AdditionalPreferential2) = rsTemp!preferential_ticket2
'                      lvStation.SelectedItem.SubItems(ID_AdditionalPreferential3) = rsTemp!preferential_ticket3
                  End Select
          End If
                If lvBus.Visible = True Then
                 lvBus.SelectedItem.SubItems(ID_CheckGate) = Trim(rsTemp!check_gate_id)
                End If
        rsTemp.MoveNext
    Next i
       If lvBus.Visible = True Then
            lvBus.SelectedItem.Tag = MakeDisplayString(lvSellStation.SelectedItem.SubItems(3), lvSellStation.SelectedItem.Text)
            lvBus.SelectedItem.SubItems(ID_OffTime) = lvSellStation.SelectedItem.SubItems(1)
            lvBus.SelectedItem.SubItems(ID_FullPrice) = FormatMoney(lvSellStation.SelectedItem.SubItems(2))
            lvBus.SelectedItem.SubItems(ID_CheckGate) = lvSellStation.SelectedItem.SubItems(4)
       Else
            lvStation.SelectedItem.Tag = MakeDisplayString(lvSellStation.SelectedItem.SubItems(3), lvSellStation.SelectedItem.Text)
            lvStation.SelectedItem.SubItems(ID_OffTime) = lvSellStation.SelectedItem.SubItems(1)
            lvStation.SelectedItem.SubItems(ID_FullPrice) = FormatMoney(lvSellStation.SelectedItem.SubItems(2))
            lvStation.SelectedItem.SubItems(ID_CheckGate) = lvSellStation.SelectedItem.SubItems(4)
       End If
    Exit Sub
err:
   MsgBox err.Description
End Sub
'刷新某车次的上车站信息
'刷新某车次的上车站信息
Private Sub RefreshSellStation(rsTemp As Recordset, aTReBusAllotInfo() As TReBusAllotInfo)
    Dim i As Integer
    Dim lvS As ListItem
    Dim szTemp As String
    Dim nBusType As EBusType
    Dim szBusID As String
    
    
    Dim j As Integer
    
    On Error GoTo err:
    lvSellStation.Sorted = False
    lvSellStation.ListItems.Clear
    lvSellStation.Refresh
    szTemp = ""
    
    If lvBus.Visible = True Then
        szBusID = Trim(lvBus.SelectedItem.Text)
    End If
    If lvStation.Visible = True Then
        szBusID = Trim(txtBusID.Text)
    End If
    '     If rsTemp.RecordCount = 0 Then
    '        Exit Sub
    '     End If
    rsTemp.MoveFirst
    For i = 1 To rsTemp.RecordCount
        If lvBus.Visible = True Then
             If Trim(rsTemp!bus_id) = Trim(lvBus.SelectedItem.Text) Then
              If Trim(rsTemp!sell_station_id) <> szTemp Then
                If m_oAUser.SellStationID = Trim(rsTemp!sell_station_id) Then '玉环,应对在选择上车站时售票员选其他站,以此要票低价卖出
                    lvSellStation.ListItems.Clear
                    szTemp = Trim(rsTemp!sell_station_id)
                    Set lvS = lvSellStation.ListItems.Add(, , Trim(rsTemp!sell_station_name))
                    nBusType = rsTemp!bus_type
                    If nBusType <> TP_ScrollBus Then
                       lvS.SubItems(1) = Trim(Format(rsTemp!BusStartTime, "hh:mm"))
                    Else
                       lvS.SubItems(1) = cszScrollBus
                    End If
                    lvS.SubItems(3) = Trim(rsTemp!sell_station_id)
                    lvS.SubItems(2) = Trim(rsTemp!full_price)
                    lvS.SubItems(4) = Trim(rsTemp!sell_check_gate_id)
                    GoTo begin
        '            End If
                ElseIf m_oAUser.SellStationID = "km" And Trim(rsTemp!sell_station_id) = "cm" Then
                    GoTo begin
                Else
                    szTemp = Trim(rsTemp!sell_station_id)
                    Set lvS = lvSellStation.ListItems.Add(, , Trim(rsTemp!sell_station_name))
                    nBusType = rsTemp!bus_type
                    If nBusType <> TP_ScrollBus Then
                       lvS.SubItems(1) = Trim(Format(rsTemp!BusStartTime, "hh:mm"))
                    Else
                       lvS.SubItems(1) = cszScrollBus
                    End If
                    lvS.SubItems(3) = Trim(rsTemp!sell_station_id)
                    lvS.SubItems(2) = Trim(rsTemp!full_price)
                    lvS.SubItems(4) = Trim(rsTemp!sell_check_gate_id)
                End If
              End If
             End If
        End If
        
        If ptExtraByBus.Visible Then
            If lvStation.ListItems.count > 0 Then
             If Trim(lvStation.SelectedItem.Text) = Trim(rsTemp!station_id) Then
              If Trim(rsTemp!sell_station_id) <> szTemp Then
                If m_oAUser.SellStationID = Trim(rsTemp!sell_station_id) Then '玉环,应对在选择上车站时售票员选其他站,以此要票低价卖出
                    lvSellStation.ListItems.Clear
                    szTemp = Trim(rsTemp!sell_station_id)
                    Set lvS = lvSellStation.ListItems.Add(, , Trim(rsTemp!sell_station_name))
                    nBusType = rsTemp!bus_type
                    If nBusType <> TP_ScrollBus Then
                            '查找发车时间
                            For j = 1 To ArrayLength(aTReBusAllotInfo)
                                If txtBusID.Text = aTReBusAllotInfo(j).szBusID And Trim(rsTemp!sell_station_id) = Trim(aTReBusAllotInfo(j).szSellStationID) Then
                                    Exit For
                                End If
                            Next j
                            If j > ArrayLength(aTReBusAllotInfo) Then
                                lvS.SubItems(1) = Trim(Format(flblOffTimeExtra.Caption, "hh:mm"))
                            Else
                                lvS.SubItems(1) = Format(aTReBusAllotInfo(j).dtRunTime, "hh:mm")
                            End If
                    Else
                       lvS.SubItems(1) = cszScrollBus
                    End If
                    lvS.SubItems(3) = Trim(rsTemp!sell_station_id)
                    lvS.SubItems(2) = FormatMoney(Trim(rsTemp!full_price))
                    lvS.SubItems(4) = Trim(rsTemp!sell_check_gate_id)
                    GoTo begin
        '            End If
                ElseIf m_oAUser.SellStationID = "km" And Trim(rsTemp!sell_station_id) = "cm" Then
                    GoTo begin
                Else
                    szTemp = Trim(rsTemp!sell_station_id)
                    Set lvS = lvSellStation.ListItems.Add(, , Trim(rsTemp!sell_station_name))
                    nBusType = rsTemp!bus_type
                    If nBusType <> TP_ScrollBus Then
                            '查找发车时间
                            For j = 1 To ArrayLength(aTReBusAllotInfo)
                                If txtBusID.Text = aTReBusAllotInfo(j).szBusID And Trim(rsTemp!sell_station_id) = Trim(aTReBusAllotInfo(j).szSellStationID) Then
                                    Exit For
                                End If
                            Next j
                            If j > ArrayLength(aTReBusAllotInfo) Then
                                lvS.SubItems(1) = Trim(Format(flblOffTimeExtra.Caption, "hh:mm"))
                            Else
                                lvS.SubItems(1) = Format(aTReBusAllotInfo(j).dtRunTime, "hh:mm")
                            End If
                    Else
                       lvS.SubItems(1) = cszScrollBus
                    End If
                    lvS.SubItems(3) = Trim(rsTemp!sell_station_id)
                    lvS.SubItems(2) = FormatMoney(Trim(rsTemp!full_price))
                    lvS.SubItems(4) = Trim(rsTemp!sell_check_gate_id)
                End If
              End If
             End If
            End If
        End If
begin:
        rsTemp.MoveNext
    Next i
    
    '玉环
    For i = 1 To lvSellStation.ListItems.count
        If m_oAUser.SellStationID = lvSellStation.ListItems(i).SubItems(3) Then

            lvSellStation.ListItems(i).Selected = True
            lvSellStation.ListItems(i).EnsureVisible
            cboSeatType_Change
            Exit For
        End If
    Next i
    Exit Sub
err:
   MsgBox err.Description
End Sub

Private Sub RefreshStation2()
    '填充站点
    Dim rsTemp As Recordset
    Dim szTemp As String
    On Error GoTo here
    szTemp = m_oSell.SellUnitCode
    m_oSell.SellUnitCode = m_szCurrentUnitID
    Set rsTemp = m_oSell.GetAllStationRs()
    m_oSell.SellUnitCode = szTemp
    With cboEndStation
        Set .RowSource = rsTemp
        'station_id:到站代码
        'station_input_code:车站输入码
        'station_name:车次名称
        
        
        .BoundField = "station_id"
        .ListFields = "station_input_code:4,station_name:4" ',station_id:5"
        .AppendWithFields "station_id:9,station_name"
    End With
    
    '因为站点已变，所以当前显示的车次信息无效，将其清空
    lvBus.ListItems.Clear
    
    '调用车次改变要进行相应操作的方法
    DoThingWhenBusChange
    Set rsTemp = Nothing
    On Error GoTo 0
Exit Sub
here:
    Set rsTemp = Nothing
    ShowErrorMsg
End Sub

Private Sub txtReceivedMoney_LostFocus()
    lblReceivedMoney.ForeColor = 0
End Sub

Private Sub txtReceivedMoney_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = vbEnter Then txtReceivedMoney.SetFocus
End Sub



'读取优惠票信息
Private Sub RefreshPreferentialTicket()
    Dim atTicketType() As TTicketType
    Dim aszSeatType() As String
    Dim szHeadText As String
    Dim sgWidth As Single
    Dim nCount As Integer
    Dim nlen As Integer
    Dim i As Integer, j As Integer
    
    Dim nUsedPerential As Integer
    
    On Error GoTo here
    nlen = 0
    
    '得到所有的票种
    atTicketType = m_oSell.GetAllTicketType()
    aszSeatType = m_oSell.GetAllSeatType
    nlen = ArrayLength(aszSeatType)
    nCount = ArrayLength(atTicketType)
    sgWidth = 800
    lvBus.ColumnHeaders.Clear
    '添加LvBus列头
    With lvBus.ColumnHeaders
        .Add , , "车次", 950.1733 '"BusID"
        .Add , , "车次类型", 0 '"BusType"
        .Add , , "时间", 1000 '"OffTime"
        .Add , , "线路名称", 0
        .Add , , "终到站", 1500 '"EndStation"
        .Add , , "总", 500
        .Add , , "订", 440
        .Add , , "座位", 700 '"SeatCount"
        .Add , , "座", 0
        .Add , , "卧", 0 '500
        .Add , , "加", 0 '500
        .Add , , "车型", 1200 '"BusModel"
        '添加票种,不可用的则宽度设为0
        For i = 1 To nCount
            If atTicketType(i).nTicketTypeID = TP_FullPrice Then
                .Add , , "座全", IIf(atTicketType(i).nTicketTypeValid = TP_TicketTypeValid, sgWidth, 0)
            Else
                .Add , , Trim(atTicketType(i).szTicketTypeName), IIf(atTicketType(i).nTicketTypeValid = TP_TicketTypeValid, sgWidth, 0)
                If atTicketType(i).nTicketTypeID = TP_HalfPrice Then lblHalfSell.Caption = Trim(atTicketType(i).szTicketTypeName) & "(&X)" & ":"
            End If
        Next i
        For i = 1 To nCount
            If atTicketType(i).nTicketTypeID = TP_FullPrice Then
                .Add , , "卧全", IIf(atTicketType(i).nTicketTypeValid = TP_TicketTypeValid, 0, 0)
            Else
                .Add , , Trim(atTicketType(i).szTicketTypeName), IIf(atTicketType(i).nTicketTypeValid = TP_TicketTypeValid, 0, 0)
            End If
        Next i
        For i = 1 To nCount
            If atTicketType(i).nTicketTypeID = TP_FullPrice Then
                .Add , , "加全", IIf(atTicketType(i).nTicketTypeValid = TP_TicketTypeValid, 0, 0)
            Else
                .Add , , Trim(atTicketType(i).szTicketTypeName), IIf(atTicketType(i).nTicketTypeValid = TP_TicketTypeValid, 0, 0)
            End If
        Next i
        .Add , , "限售张数", 0 '"LimitedCount"
        .Add , , "限售时间", 0 '"LimitedTime"
        .Add , , "车次类型", 0 '"BusType"
        .Add , , "检票口", 0 '"CheckGate"
        .Add , , "站票", 0 '"StandCount"
        .Add , , "是否实名制", 0
    End With
    '添加lvStation列头
    lvStation.ColumnHeaders.Clear
    With lvStation.ColumnHeaders
        .Add , , "终到站代码", 1299.969
        .Add , , "终到站", 1440
        .Add , , "车次类型", 0
        '添加票种,不可用的则宽度设为0
        For i = 1 To nCount
            If atTicketType(i).nTicketTypeID = TP_FullPrice Then
                .Add , , "座全", IIf(atTicketType(i).nTicketTypeValid = TP_TicketTypeValid, sgWidth, 0)
            Else
                .Add , , Trim(atTicketType(i).szTicketTypeName), IIf(atTicketType(i).nTicketTypeValid = TP_TicketTypeValid, sgWidth, 0)
            End If
        Next i
        For i = 1 To nCount
            If atTicketType(i).nTicketTypeID = TP_FullPrice Then
                .Add , , "卧全", IIf(atTicketType(i).nTicketTypeValid = TP_TicketTypeValid, 0, 0)
            Else
                .Add , , Trim(atTicketType(i).szTicketTypeName), IIf(atTicketType(i).nTicketTypeValid = TP_TicketTypeValid, 0, 0)
            End If
        Next i
        For i = 1 To nCount
            If atTicketType(i).nTicketTypeID = TP_FullPrice Then
                .Add , , "加全", IIf(atTicketType(i).nTicketTypeValid = TP_TicketTypeValid, 0, 0)
            Else
                .Add , , Trim(atTicketType(i).szTicketTypeName), IIf(atTicketType(i).nTicketTypeValid = TP_TicketTypeValid, 0, 0)
            End If
        Next i
        .Add , , "限售张数", 0
        .Add , , "限售时间", 0
        .Add , , "终点站名", 0
    End With
    
    '设置座位类型
    If nlen <> 0 Then
        ReDim m_aszSeatType(1 To nlen, 1 To 3)
        For i = 1 To nlen
            cboSeatType.AddItem aszSeatType(i, 2)
            m_aszSeatType(i, 1) = aszSeatType(i, 1)
            m_aszSeatType(i, 2) = aszSeatType(i, 2)
            m_aszSeatType(i, 3) = aszSeatType(i, 3)
        Next
    End If
    If cboSeatType.ListCount > 0 Then
        cboSeatType.ListIndex = 0
    End If
    '设置ComboBox和优惠票是否可用
    nUsedPerential = 0
    For i = 1 To nCount
        If atTicketType(i).nTicketTypeID = TP_HalfPrice And atTicketType(i).nTicketTypeValid = TP_TicketTypeValid Then
            txtHalfSell.Visible = True
            lblHalfSell.Visible = True
        ElseIf atTicketType(i).nTicketTypeID > TP_HalfPrice And atTicketType(i).nTicketTypeValid = TP_TicketTypeValid Then
            cboPreferentialTicket.AddItem Trim(atTicketType(i).szTicketTypeName)
            nUsedPerential = nUsedPerential + 1
        End If
    Next i
    
    If cboPreferentialTicket.ListCount < 1 Then
        txtPreferentialSell.Enabled = False
        cboPreferentialTicket.Enabled = False
        cboPreferentialTicket.Text = ""
    Else
        txtPreferentialSell.Enabled = True
        cboPreferentialTicket.Enabled = True
        cboPreferentialTicket.ListIndex = 0
    End If
    '将组合框中的票种代码与票种名称放到数组m_atTicketType 中
    If nUsedPerential > 0 Then ReDim m_atTicketType(1 To nUsedPerential)
    j = 0
    For i = 1 To nCount
        If atTicketType(i).nTicketTypeID > TP_HalfPrice And atTicketType(i).nTicketTypeValid = TP_TicketTypeValid Then
            j = j + 1
            m_atTicketType(j) = atTicketType(i)
        End If
    Next i
    On Error GoTo 0
 Exit Sub
here:
    ShowErrorMsg
End Sub

'得到对应的优惠票种的对应的票价
Private Function GetPreferentialPrice(Optional pbIsSell As Boolean = False) As Double
Dim liTemp As ListItem
Dim dbTemp As Double
    '如果是用通过车次补票
    
    If ptExtraByBus.Visible Then
        Set liTemp = lvStation.SelectedItem
        Select Case Trim(m_aszSeatType(cboSeatType.ListIndex + 1, 1))
            Case cszSeatType
                If cboPreferentialTicket.ListCount > 0 Then
                    Select Case m_atTicketType(cboPreferentialTicket.ListIndex + 1).nTicketTypeID
                    Case TP_FreeTicket
                        dbTemp = CDbl(liTemp.SubItems(IIf(pbIsSell, ID_FullPrice2, ID_FreePrice2)))
                    Case TP_PreferentialTicket1
                        dbTemp = CDbl(liTemp.SubItems(ID_PreferentialPrice21))
                    Case TP_PreferentialTicket2
                        dbTemp = CDbl(liTemp.SubItems(ID_PreferentialPrice22))
                    Case TP_PreferentialTicket3
                        dbTemp = CDbl(liTemp.SubItems(ID_PreferentialPrice23))
                    End Select
                End If

            Case cszBedType
                If cboPreferentialTicket.ListCount > 0 Then
                    Select Case m_atTicketType(cboPreferentialTicket.ListIndex + 1).nTicketTypeID
                    Case TP_FreeTicket
                        dbTemp = CDbl(liTemp.SubItems(IIf(pbIsSell, ID_BedFullPrice2, ID_FreePrice2)))
                    Case TP_PreferentialTicket1
                        dbTemp = CDbl(liTemp.SubItems(ID_BedPreferential21))
                    Case TP_PreferentialTicket2
                        dbTemp = CDbl(liTemp.SubItems(ID_BedPreferential22))
                    Case TP_PreferentialTicket3
                        dbTemp = CDbl(liTemp.SubItems(ID_BedPreferential23))
                    End Select
                End If
                
            Case cszAdditionalType
                If cboPreferentialTicket.ListCount > 0 Then
                    Select Case m_atTicketType(cboPreferentialTicket.ListIndex + 1).nTicketTypeID
                    Case TP_FreeTicket
                        dbTemp = CDbl(liTemp.SubItems(IIf(pbIsSell, ID_AdditionalFullPrice2, ID_FreePrice2)))
                    Case TP_PreferentialTicket1
                        dbTemp = CDbl(liTemp.SubItems(ID_AdditionalPreferential21))
                    Case TP_PreferentialTicket2
                        dbTemp = CDbl(liTemp.SubItems(ID_AdditionalPreferential22))
                    Case TP_PreferentialTicket3
                        dbTemp = CDbl(liTemp.SubItems(ID_AdditionalPreferential23))
                    End Select
                End If
                
        End Select
                
    Else
        Set liTemp = lvBus.SelectedItem
        Select Case Trim(m_aszSeatType(cboSeatType.ListIndex + 1, 1))
        Case cszSeatType
            If cboPreferentialTicket.ListCount > 0 Then
                Select Case m_atTicketType(cboPreferentialTicket.ListIndex + 1).nTicketTypeID
                    Case TP_FreeTicket
                        dbTemp = CDbl(liTemp.SubItems(IIf(pbIsSell, ID_FullPrice, ID_FreePrice)))
                    Case TP_PreferentialTicket1
                        dbTemp = CDbl(liTemp.SubItems(ID_PreferentialPrice1))
                    Case TP_PreferentialTicket2
                        dbTemp = CDbl(liTemp.SubItems(ID_PreferentialPrice2))
                    Case TP_PreferentialTicket3
                        dbTemp = CDbl(liTemp.SubItems(ID_PreferentialPrice3))
                End Select
            End If
        Case cszBedType
            If cboPreferentialTicket.ListCount > 0 Then
                Select Case m_atTicketType(cboPreferentialTicket.ListIndex + 1).nTicketTypeID
                    Case TP_FreeTicket
                        dbTemp = CDbl(liTemp.SubItems(IIf(pbIsSell, ID_BedFullPrice, ID_BedFreePrice)))
                    Case TP_PreferentialTicket1
                        dbTemp = CDbl(liTemp.SubItems(ID_BedPreferentialPrice1))
                    Case TP_PreferentialTicket2
                        dbTemp = CDbl(liTemp.SubItems(ID_BedPreferentialPrice2))
                    Case TP_PreferentialTicket3
                        dbTemp = CDbl(liTemp.SubItems(ID_BedPreferentialPrice3))
                End Select
            End If
        Case cszAdditionalType
            If cboPreferentialTicket.ListCount > 0 Then
                Select Case m_atTicketType(cboPreferentialTicket.ListIndex + 1).nTicketTypeID
                    Case TP_FreeTicket
                        dbTemp = CDbl(liTemp.SubItems(IIf(pbIsSell, ID_AdditionalFullPrice, ID_FreePrice)))
                    Case TP_PreferentialTicket1
                        dbTemp = CDbl(liTemp.SubItems(ID_AdditionalPreferential1))
                    Case TP_PreferentialTicket2
                        dbTemp = CDbl(liTemp.SubItems(ID_AdditionalPreferential2))
                    Case TP_PreferentialTicket3
                        dbTemp = CDbl(liTemp.SubItems(ID_AdditionalPreferential3))
                End Select
            End If
    End Select
    End If
    GetPreferentialPrice = dbTemp
    Set liTemp = Nothing
End Function

'按到站预售所需信息
Private Sub SetStationPreSellInfo()
Dim atTicketType() As TTicketType
Dim nCount As Integer
Dim i As Integer

atTicketType = m_oSell.GetAllTicketType()
nCount = ArrayLength(atTicketType)
   lvBusPreSell.ColumnHeaders.Clear
   With lvBusPreSell.ColumnHeaders
        .Add , , "车次", 950  '"BusID"
        .Add , , "车次类型", 0 '"BusType"
        .Add , , "发车时间", 1200  '"BusDate"
        .Add , , "终到站", 899   '"EndStation"
        .Add , , "总票数", 899  '"TotalTicketNo"
        .Add , , "总票价", 899 '"TotalPrice"
        .Add , , "车型", 899  '"VehicleModel"
        For i = 1 To nCount
             .Add , , Trim(atTicketType(i).szTicketTypeName), IIf(atTicketType(i).nTicketTypeValid = TP_TicketTypeValid, 0, 0)
             .Add , , Trim(atTicketType(i).szTicketTypeName) & "票数", IIf(atTicketType(i).nTicketTypeValid = TP_TicketTypeValid, 1000, 0)
        Next i
        .Add , , "折扣率", 0  '"Discount"
        .Add , , "定座", 0  '"OrderSeat"
        .Add , , "站票", 0  '"StandTicket"
        .Add , , "检票口", 0 '"CheckGate"
        .Add , , "终到站代码", 0 '"EndStationCode"
        .Add , , "限售张数", 0  '"LimitedCount"
        .Add , , "座位状态1", 0 '"SeatStatus1"
        .Add , , "座位状态2", 0 '"SeatStatus2"
        .Add , , "座位号", 0   '"SeatNo"
        .Add , , "明细票价", 0 '"AllTicketPrice"
        .Add , , "明细票种", 0 '"AllTicketPrice"
        .Add , , "座位类型", 0
        .Add , , "终点站名", 0
        .Add , , "是否实名制", 0
   End With
End Sub

'按车次预售所需信息
Private Sub SetBusPreSellInfo()
Dim atTicketType() As TTicketType
Dim nCount As Integer
Dim i As Integer

atTicketType = m_oSell.GetAllTicketType()
nCount = ArrayLength(atTicketType)
   lvStationPreSell.ColumnHeaders.Clear
   With lvStationPreSell.ColumnHeaders
        .Add , , "车次", 950 '"BusID"
        .Add , , "车次类型", 0 '"BusType"
        .Add , , "发车时间", 1200 '"BusDate"
        .Add , , "终到站", 899  '"EndStation"
        .Add , , "总票数", 899  '"TotalTicketNo"
        .Add , , "总票价", 899 '"TotalPrice"
        .Add , , "车型", 899  '"VehicleModel"
        For i = 1 To nCount
             .Add , , Trim(atTicketType(i).szTicketTypeName), IIf(atTicketType(i).nTicketTypeValid = TP_TicketTypeValid, 0, 0)
             .Add , , Trim(atTicketType(i).szTicketTypeName) & "票数", IIf(atTicketType(i).nTicketTypeValid = TP_TicketTypeValid, 1000, 0)
        Next i
        .Add , , "折扣率", 0  '"Discount"
        .Add , , "定座", 0  '"OrderSeat"
        .Add , , "站票", 0  '"StandTicket"
        .Add , , "检票口", 0 '"CheckGate"
        .Add , , "终到站代码", 0 '"EndStationCode"
        .Add , , "限售张数", 0  '"LimitedCount"
        .Add , , "座位状态1", 0 '"SeatStatus1"
        .Add , , "座位状态2", 0 '"SeatStatus2"
        .Add , , "座位号", 0   '"SeatNo"
        .Add , , "明细票价", 0 '"AllTicketPrice"
        .Add , , "明细票种", 0 '"AllTicketPrice"
        .Add , , "座位类型", 0
        .Add , , "终点站名", 0
        .Add , , "总张数", 0
        .Add , , "是否实名制", 0
   End With
End Sub

'按到站得到预售所需信息
Private Sub GetStationPreSellInfo()
    Dim liPreSell As ListItem
    Dim liBus As ListItem
    Dim i As Integer
    Dim szPrice As String
    Dim szTicketType As String
    If Not lvBus.SelectedItem Is Nothing Then
        Set liBus = lvBus.SelectedItem
        Set liPreSell = lvStationPreSell.ListItems.Add(, , liBus.Text)
        With liPreSell
            .Tag = lvBus.SelectedItem.Tag
            .SubItems(EI_BusType) = liBus.SubItems(ID_BusType)
            .SubItems(EI_OffTime) = liBus.SubItems(ID_OffTime)
            .SubItems(EI_EndStation) = GetStationNameInCbo(cboEndStation.Text)
            .SubItems(EI_TotalNum) = txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text
            .SubItems(EI_VehicleModel) = liBus.SubItems(ID_VehicleModel)
            
            .SubItems(EI_SumTicketNum) = txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text
           
            Select Case Trim(m_aszSeatType(cboSeatType.ListIndex + 1, 1))
                Case cszSeatType
                     .SubItems(EI_FullPrice) = FormatMoney(liBus.SubItems(ID_FullPrice))
                     .SubItems(EI_FullNum) = txtFullSell.Text
                     .SubItems(EI_HalfPrice) = FormatMoney(liBus.SubItems(ID_HalfPrice))
                     .SubItems(EI_HalfNum) = txtHalfSell.Text
                    If cboPreferentialTicket.ListCount > 0 Then
                        Select Case m_atTicketType(cboPreferentialTicket.ListIndex + 1).nTicketTypeID
                            Case TP_FreeTicket
                                .SubItems(EI_FreePrice) = 0
                                .SubItems(EI_FreeNum) = txtPreferentialSell.Text
                            Case TP_PreferentialTicket1
                                .SubItems(EI_PreferentialPrice1) = FormatMoney(liBus.SubItems(ID_PreferentialPrice1))
                                .SubItems(EI_PreferentialNum1) = txtPreferentialSell.Text
                            Case TP_PreferentialTicket2
                                .SubItems(EI_PreferentialPrice2) = FormatMoney(liBus.SubItems(ID_PreferentialPrice2))
                                .SubItems(EI_PreferentialNum2) = txtPreferentialSell.Text
                            Case TP_PreferentialTicket3
                                .SubItems(EI_PreferentialPrice3) = FormatMoney(liBus.SubItems(ID_PreferentialPrice3))
                                .SubItems(EI_PreferentialNum3) = txtPreferentialSell.Text
                        End Select
                    End If
                Case cszBedType
                     .SubItems(EI_FullPrice) = FormatMoney(liBus.SubItems(ID_BedFullPrice))
                     .SubItems(EI_FullNum) = txtFullSell.Text
                     .SubItems(EI_HalfPrice) = FormatMoney(liBus.SubItems(ID_BedHalfPrice))
                     .SubItems(EI_HalfNum) = txtHalfSell.Text
                    If cboPreferentialTicket.ListCount > 0 Then
                        Select Case m_atTicketType(cboPreferentialTicket.ListIndex + 1).nTicketTypeID
                            Case TP_FreeTicket
                                .SubItems(EI_FreePrice) = 0
                                .SubItems(EI_FreeNum) = txtPreferentialSell.Text
                            Case TP_PreferentialTicket1
                                .SubItems(EI_PreferentialPrice1) = FormatMoney(liBus.SubItems(ID_BedPreferentialPrice1))
                                .SubItems(EI_PreferentialNum1) = txtPreferentialSell.Text
                            Case TP_PreferentialTicket2
                                .SubItems(EI_PreferentialPrice2) = FormatMoney(liBus.SubItems(ID_BedPreferentialPrice2))
                                .SubItems(EI_PreferentialNum2) = txtPreferentialSell.Text
                            Case TP_PreferentialTicket3
                                .SubItems(EI_PreferentialPrice3) = FormatMoney(liBus.SubItems(ID_BedPreferentialPrice3))
                                .SubItems(EI_PreferentialNum3) = txtPreferentialSell.Text
                        End Select
                    End If
                Case cszAdditionalType
                     .SubItems(EI_FullPrice) = FormatMoney(liBus.SubItems(ID_AdditionalFullPrice))
                     .SubItems(EI_FullNum) = txtFullSell.Text
                     .SubItems(EI_HalfPrice) = FormatMoney(liBus.SubItems(ID_AdditionalHalfPrice))
                     .SubItems(EI_HalfNum) = txtHalfSell.Text
                    If cboPreferentialTicket.ListCount > 0 Then
                        Select Case m_atTicketType(cboPreferentialTicket.ListIndex + 1).nTicketTypeID
                            Case TP_FreeTicket
                                .SubItems(EI_FreePrice) = 0
                                .SubItems(EI_FreeNum) = txtPreferentialSell.Text
                            Case TP_PreferentialTicket1
                                .SubItems(EI_PreferentialPrice1) = FormatMoney(liBus.SubItems(ID_AdditionalPreferential1))
                                .SubItems(EI_PreferentialNum1) = txtPreferentialSell.Text
                            Case TP_PreferentialTicket2
                                .SubItems(EI_PreferentialPrice2) = FormatMoney(liBus.SubItems(ID_AdditionalPreferential2))
                                .SubItems(EI_PreferentialNum2) = txtPreferentialSell.Text
                            Case TP_PreferentialTicket3
                                .SubItems(EI_PreferentialPrice3) = FormatMoney(liBus.SubItems(ID_AdditionalPreferential3))
                                .SubItems(EI_PreferentialNum3) = txtPreferentialSell.Text
                        End Select
                    End If
            End Select
            .SubItems(EI_TotalPrice) = FormatMoney(Val(.SubItems(EI_FullNum)) * Val(.SubItems(EI_FullPrice)) + _
                                        Val(.SubItems(EI_HalfNum)) * Val(.SubItems(EI_HalfPrice)) + _
                                        Val(.SubItems(EI_PreferentialNum1)) * Val(.SubItems(EI_PreferentialPrice1)) + _
                                        Val(.SubItems(EI_PreferentialNum2)) * Val(.SubItems(EI_PreferentialPrice2)) + _
                                        Val(.SubItems(EI_PreferentialNum3)) * Val(.SubItems(EI_PreferentialPrice3)))
            .SubItems(EI_Discount) = CSng(txtDiscount.Text)
            .SubItems(EI_CheckGate) = liBus.SubItems(ID_CheckGate)
            .SubItems(EI_EndStationCode) = cboEndStation.BoundText
            .SubItems(EI_SeatStatus1) = True
            .SubItems(EI_SeatStatus2) = vbChecked
            .SubItems(EI_SeatNo) = ""
            .SubItems(EI_SeatType) = m_aszSeatType(cboSeatType.ListIndex + 1, 1)
            .SubItems(EI_TerminateName) = liBus.SubItems(ID_EndStation)
            
            .SubItems(EI_RealName) = liBus.SubItems(ID_RealName)
            SetRealNameInfo .SubItems(EI_RealName), Val(.SubItems(EI_TotalNum))
           
        End With
        Set liPreSell = Nothing
        Set liBus = Nothing
    End If
End Sub

'按车次预售得到所需信息
Private Sub GetBusPreSellInfo()
   Dim liPreSell As ListItem
   Dim liStation As ListItem
   Dim i As Integer
   Dim szPrice As String
   Dim szTicketType As String
   On Error GoTo here
   If Not lvStation.SelectedItem Is Nothing Then
        Set liStation = lvStation.SelectedItem
        Set liPreSell = lvBusPreSell.ListItems.Add(, , txtBusID.Text)
        With liPreSell
            .Tag = lvStation.SelectedItem.Tag
            .SubItems(EI_BusType) = liStation.SubItems(ID_BusType2)
            .SubItems(EI_OffTime) = flblOffTimeExtra.Caption
            .SubItems(EI_EndStation) = liStation.SubItems(ID_StationName)
            .SubItems(EI_TotalNum) = txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text
            .SubItems(EI_VehicleModel) = flblBusTypeExtra.Caption
            
            .SubItems(EI_SumTicketNum) = txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text
            
            Select Case Trim(m_aszSeatType(cboSeatType.ListIndex + 1, 1))
                Case cszSeatType
                    .SubItems(EI_FullPrice) = FormatMoney(lvStation.SelectedItem.SubItems(ID_FullPrice2))
                    .SubItems(EI_FullNum) = txtFullSell.Text
                    .SubItems(EI_HalfPrice) = FormatMoney(lvStation.SelectedItem.SubItems(ID_HalfPrice2))
                    .SubItems(EI_HalfNum) = txtHalfSell.Text
                    If cboPreferentialTicket.ListCount > 0 Then
                        Select Case m_atTicketType(cboPreferentialTicket.ListIndex + 1).nTicketTypeID
                            Case TP_FreeTicket
                                .SubItems(EI_FreePrice) = 0
                                .SubItems(EI_FreeNum) = txtPreferentialSell.Text
                            Case TP_PreferentialTicket1
                                .SubItems(EI_PreferentialPrice1) = FormatMoney(lvStation.SelectedItem.SubItems(ID_PreferentialPrice21))
                                .SubItems(EI_PreferentialNum1) = txtPreferentialSell.Text
                            Case TP_PreferentialTicket2
                                .SubItems(EI_PreferentialPrice2) = FormatMoney(lvStation.SelectedItem.SubItems(ID_PreferentialPrice22))
                                .SubItems(EI_PreferentialNum2) = txtPreferentialSell.Text
                            Case TP_PreferentialTicket3
                                .SubItems(EI_PreferentialPrice3) = FormatMoney(lvStation.SelectedItem.SubItems(ID_PreferentialPrice23))
                                .SubItems(EI_PreferentialNum3) = txtPreferentialSell.Text
                        End Select
                    End If
                Case cszBedType
                    .SubItems(EI_FullPrice) = FormatMoney(lvStation.SelectedItem.SubItems(ID_BedFullPrice2))
                    .SubItems(EI_FullNum) = txtFullSell.Text
                    .SubItems(EI_HalfPrice) = FormatMoney(lvStation.SelectedItem.SubItems(ID_BedHalfPrice2))
                    .SubItems(EI_HalfNum) = txtHalfSell.Text
                    If cboPreferentialTicket.ListCount > 0 Then
                        Select Case m_atTicketType(cboPreferentialTicket.ListIndex + 1).nTicketTypeID
                            Case TP_FreeTicket
                                .SubItems(EI_FreePrice) = 0
                                .SubItems(EI_FreeNum) = txtPreferentialSell.Text
                            Case TP_PreferentialTicket1
                                .SubItems(EI_PreferentialPrice1) = FormatMoney(lvStation.SelectedItem.SubItems(ID_BedPreferential21))
                                .SubItems(EI_PreferentialNum1) = txtPreferentialSell.Text
                            Case TP_PreferentialTicket2
                                .SubItems(EI_PreferentialPrice2) = FormatMoney(lvStation.SelectedItem.SubItems(ID_BedPreferential22))
                                .SubItems(EI_PreferentialNum2) = txtPreferentialSell.Text
                            Case TP_PreferentialTicket3
                                .SubItems(EI_PreferentialPrice3) = FormatMoney(lvStation.SelectedItem.SubItems(ID_BedPreferential23))
                                .SubItems(EI_PreferentialNum3) = txtPreferentialSell.Text
                        End Select
                    End If
                
                Case cszAdditionalType
                    .SubItems(EI_FullPrice) = FormatMoney(lvStation.SelectedItem.SubItems(ID_AdditionalFullPrice2))
                    .SubItems(EI_FullNum) = txtFullSell.Text
                    .SubItems(EI_HalfPrice) = FormatMoney(lvStation.SelectedItem.SubItems(ID_AdditionalHalfPrice2))
                    .SubItems(EI_HalfNum) = txtHalfSell.Text
                    If cboPreferentialTicket.ListCount > 0 Then
                        Select Case m_atTicketType(cboPreferentialTicket.ListIndex + 1).nTicketTypeID
                            Case TP_FreeTicket
                                .SubItems(EI_FreePrice) = 0
                                .SubItems(EI_FreeNum) = txtPreferentialSell.Text
                            Case TP_PreferentialTicket1
                                .SubItems(EI_PreferentialPrice1) = FormatMoney(lvStation.SelectedItem.SubItems(ID_AdditionalPreferential21))
                                .SubItems(EI_PreferentialNum1) = txtPreferentialSell.Text
                            Case TP_PreferentialTicket2
                                .SubItems(EI_PreferentialPrice2) = FormatMoney(lvStation.SelectedItem.SubItems(ID_AdditionalPreferential22))
                                .SubItems(EI_PreferentialNum2) = txtPreferentialSell.Text
                            Case TP_PreferentialTicket3
                                .SubItems(EI_PreferentialPrice3) = FormatMoney(lvStation.SelectedItem.SubItems(ID_AdditionalPreferential3))
                                .SubItems(EI_PreferentialNum3) = txtPreferentialSell.Text
                        End Select
                    End If
                
            End Select
            .SubItems(EI_TotalPrice) = FormatMoney(Val(.SubItems(EI_FullNum)) * Val(.SubItems(EI_FullPrice)) + _
                                        Val(.SubItems(EI_HalfNum)) * Val(.SubItems(EI_HalfPrice)) + _
                                        Val(.SubItems(EI_PreferentialNum1)) * Val(.SubItems(EI_PreferentialPrice1)) + _
                                        Val(.SubItems(EI_PreferentialNum2)) * Val(.SubItems(EI_PreferentialPrice2)) + _
                                        Val(.SubItems(EI_PreferentialNum3)) * Val(.SubItems(EI_PreferentialPrice3)))
                                        
            .SubItems(EI_Discount) = CSng(txtDiscount.Text)
            .SubItems(EI_CheckGate) = m_szCheckGate
            .SubItems(EI_EndStationCode) = lvStation.SelectedItem.Text
            .SubItems(EI_SeatStatus1) = chkSetSeat.Enabled
            .SubItems(EI_SeatStatus2) = chkSetSeat.Value
            .SubItems(EI_SeatNo) = txtSeat.Text
            .SubItems(EI_SeatType) = m_aszSeatType(cboSeatType.ListIndex + 1, 1)
            .SubItems(EI_TerminateName) = Trim(lvStation.SelectedItem.SubItems(ID_TerminateName2))
            
            .SubItems(EI_RealName) = lvStation.SelectedItem.SubItems(ID_RealName)
            SetRealNameInfo .SubItems(EI_RealName), Val(.SubItems(EI_TotalNum))
           
        End With
        
   End If
   Set liStation = Nothing
   Set liPreSell = Nothing
   On Error GoTo 0
   Exit Sub
here:
     Set liStation = Nothing
     Set liPreSell = Nothing
     ShowErrorMsg
End Sub

Private Sub SetRealNameInfo(bIsRealName As Boolean, nSellNums As Integer)
On Error GoTo here
    Dim aszRealNameInfoTemp() As TCardInfo
    Dim i As Integer
    Dim nCountOld As Integer
    Dim nCountAdd As Integer
    Dim nCountNew As Integer
    
    If bIsRealName = True Then  '需要实名制
        frmRealNameReg.m_nSellCount = nSellNums
        frmRealNameReg.Show vbModal
        If frmRealNameReg.m_bOk Then
            aszRealNameInfoTemp = frmRealNameReg.m_vaCardInfo
            nCountOld = ArrayLength(m_aszRealNameInfo)
            nCountAdd = ArrayLength(aszRealNameInfoTemp)
            nCountNew = nCountOld + nCountAdd
            If nCountOld = 0 Then
                ReDim m_aszRealNameInfo(1 To nCountNew)
            Else
                ReDim Preserve m_aszRealNameInfo(1 To nCountNew)
            End If
            For i = 1 To nCountAdd
                m_aszRealNameInfo(nCountOld + i).szCardType = aszRealNameInfoTemp(i).szCardType
                m_aszRealNameInfo(nCountOld + i).szIDCardNo = aszRealNameInfoTemp(i).szIDCardNo
                m_aszRealNameInfo(nCountOld + i).szPersonName = aszRealNameInfoTemp(i).szPersonName
                m_aszRealNameInfo(nCountOld + i).szSex = aszRealNameInfoTemp(i).szSex
                m_aszRealNameInfo(nCountOld + i).szNation = aszRealNameInfoTemp(i).szNation
                m_aszRealNameInfo(nCountOld + i).szAddress = aszRealNameInfoTemp(i).szAddress
                m_aszRealNameInfo(nCountOld + i).szPersonPicture = aszRealNameInfoTemp(i).szPersonPicture
            Next i
        Else
            nCountOld = ArrayLength(m_aszRealNameInfo)
            nCountAdd = nSellNums
            nCountNew = nCountOld + nCountAdd
            If nCountOld = 0 Then
                ReDim m_aszRealNameInfo(1 To nCountNew)
            Else
                ReDim Preserve m_aszRealNameInfo(1 To nCountNew)
            End If
            For i = 1 To nCountAdd
                m_aszRealNameInfo(nCountOld + i).szCardType = ""
                m_aszRealNameInfo(nCountOld + i).szIDCardNo = ""
                m_aszRealNameInfo(nCountOld + i).szPersonName = ""
                m_aszRealNameInfo(nCountOld + i).szSex = ""
                m_aszRealNameInfo(nCountOld + i).szNation = ""
                m_aszRealNameInfo(nCountOld + i).szAddress = ""
                m_aszRealNameInfo(nCountOld + i).szPersonPicture = ""
            Next i
        End If
    Else    '不需要实名制
        nCountOld = ArrayLength(m_aszRealNameInfo)
        nCountAdd = nSellNums
        nCountNew = nCountOld + nCountAdd
        If nCountOld = 0 Then
            ReDim m_aszRealNameInfo(1 To nCountNew)
        Else
            ReDim Preserve m_aszRealNameInfo(1 To nCountNew)
        End If
        For i = 1 To nCountAdd
            m_aszRealNameInfo(nCountOld + i).szCardType = ""
            m_aszRealNameInfo(nCountOld + i).szIDCardNo = ""
            m_aszRealNameInfo(nCountOld + i).szPersonName = ""
            m_aszRealNameInfo(nCountOld + i).szSex = ""
            m_aszRealNameInfo(nCountOld + i).szNation = ""
            m_aszRealNameInfo(nCountOld + i).szAddress = ""
            m_aszRealNameInfo(nCountOld + i).szPersonPicture = ""
        Next i
    End If

    Exit Sub
here:
    ShowErrorMsg
End Sub

Private Function GetDealTotalPrice() As Double  '得到总票价
    Dim iCount As Integer
    Dim dbTotal As Double
    
    dbTotal = 0
    If BusIsActive Then
        If lvBusPreSell.ListItems.count <> 0 Then
            For iCount = 1 To lvBusPreSell.ListItems.count
                dbTotal = dbTotal + lvBusPreSell.ListItems(iCount).SubItems(EI_TotalPrice)
            Next iCount
        End If
        GetDealTotalPrice = FormatMoney(dbTotal)
    Else
        If lvStationPreSell.ListItems.count <> 0 Then
            For iCount = 1 To lvStationPreSell.ListItems.count
                dbTotal = dbTotal + lvStationPreSell.ListItems(iCount).SubItems(EI_TotalPrice)
            Next iCount
        End If
        GetDealTotalPrice = FormatMoney(dbTotal)
    End If
End Function

'设置预售按钮状态
Private Sub SetPreSellButton()
'    If txtFullSell.text = 0 And txtHalfSell.text = 0 And txtPreferentialSell.Text = 0 Then
'        cmdPreSell.Enabled = False
'    Else
'        cmdPreSell.Enabled = True
'    End If
End Sub

'处理折扣票与定座
Private Sub DealDiscountAndSeat()
   '判断是否有售折扣票权限
On Error GoTo here
   If m_oSell.DiscountIsValid Then
        txtDiscount.Enabled = False
        fraDiscountTicket.Enabled = False
   End If
   If m_oSell.OrderSeatIsValid Then
        chkSetSeat.Value = 0
        chkSetSeat.Visible = False
        lblSetSeat.Enabled = False
   End If
   On Error GoTo 0
   Exit Sub
here:
    ShowErrorMsg
End Sub

'预售用订座
Private Function PreOrderSeat() As String
Dim i As Integer
Dim szTemp As String
Dim liTemp As ListItem
If ptExtraByBus.Visible Then
    If lvBusPreSell.ListItems.count <> 0 Then
        For i = 1 To lvBusPreSell.ListItems.count
            Set liTemp = lvBusPreSell.ListItems(i)
            If txtBusID.Text = liTemp.Text Then
                If liTemp.SubItems(EI_SeatNo) <> "" Then
                    szTemp = szTemp & "," & liTemp.SubItems(EI_SeatNo)
                End If
            End If
        Next i
    Else
        szTemp = ""
    End If
Else
    If lvStationPreSell.ListItems.count <> 0 Then
        For i = 1 To lvStationPreSell.ListItems.count
            Set liTemp = lvStationPreSell.ListItems(i)
            If lvBus.SelectedItem.Text = liTemp.Text Then
                If liTemp.SubItems(EI_SeatNo) <> "" Then
                    szTemp = szTemp & "," & liTemp.SubItems(EI_SeatNo)
                End If
            End If
        Next i
    Else
        szTemp = ""
    End If
End If
PreOrderSeat = szTemp
End Function

'返回相同车次索引
Private Function GetSameBusIndex(lvPreSell As ListView, lvSell As ListView, Optional pbBus As Boolean = False) As Integer
Dim i As Integer
Dim liTemp As ListItem
Dim liSelected As ListItem
If lvPreSell.ListItems.count <> 0 And (Not lvSell.SelectedItem Is Nothing) Then
    Set liSelected = lvSell.SelectedItem
    For i = 1 To lvPreSell.ListItems.count
        Set liTemp = lvPreSell.ListItems(i)
        If pbBus Then
            If liTemp.Text = liSelected.Text And liTemp.SubItems(EI_EndStationCode) = cboEndStation.BoundText Then
                GetSameBusIndex = i
                Exit Function
            End If
        Else
            If liTemp.Text = txtBusID.Text And liTemp.SubItems(EI_EndStationCode) = liSelected.Text Then
                GetSameBusIndex = i
                Exit Function
            End If
        End If
    Next i
End If
GetSameBusIndex = 0
End Function

'合并相同车次信息
Private Sub MergeSameBusInfo(nSameIndex As Integer)
Dim liTemp As ListItem
Set liTemp = lvBusPreSell.ListItems(nSameIndex)
Dim szPrice As String
Dim szTicketType As String
Dim i As Integer
Dim sgTemp As Single
sgTemp = 0
With liTemp

    .SubItems(EI_SumTicketNum) = Val(.SubItems(EI_SumTicketNum)) + txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text
    .SubItems(EI_TotalNum) = Val(.SubItems(EI_TotalNum)) + txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text
    .SubItems(EI_FullNum) = Val(.SubItems(EI_FullNum)) + txtFullSell.Text
    .SubItems(EI_HalfNum) = Val(.SubItems(EI_HalfNum)) + txtHalfSell.Text
    .SubItems(EI_SeatNo) = Trim(.SubItems(EI_SeatNo)) & "," & Trim(txtSeat.Text)
    .SubItems(EI_SeatNo) = Trim(.SubItems(EI_SeatNo)) & "," & Trim("")
    
    If cboPreferentialTicket.ListCount > 0 Then
        Select Case m_atTicketType(cboPreferentialTicket.ListIndex + 1).nTicketTypeID
        Case TP_FreeTicket
            .SubItems(EI_FreeNum) = txtPreferentialSell.Text + Val(.SubItems(EI_FreeNum))
        Case TP_PreferentialTicket1
            .SubItems(EI_PreferentialNum1) = txtPreferentialSell.Text + Val(.SubItems(EI_PreferentialNum1))
        Case TP_PreferentialTicket2
            .SubItems(EI_PreferentialNum2) = txtPreferentialSell.Text + Val(.SubItems(EI_PreferentialNum2))
        Case TP_PreferentialTicket3
            .SubItems(EI_PreferentialNum3) = txtPreferentialSell.Text + Val(.SubItems(EI_PreferentialNum3))
        End Select
    End If
    
    
    .SubItems(EI_TotalPrice) = Val(.SubItems(EI_TotalPrice)) + _
                                txtFullSell.Text * lvStation.SelectedItem.SubItems(ID_FullPrice2) + _
                                txtHalfSell.Text * lvStation.SelectedItem.SubItems(ID_HalfPrice2) + _
                                txtPreferentialSell.Text * GetPreferentialPrice
                                
    
End With
End Sub

'合并相同车次信息
Private Sub MergeSameStationInfo(nSameIndex As Integer)
Dim liTemp As ListItem
Set liTemp = lvStationPreSell.ListItems(nSameIndex)
Dim szPrice As String
Dim szTicketType As String
Dim i As Integer
Dim sgTemp As Single
sgTemp = 0
With liTemp
    
    .SubItems(EI_SumTicketNum) = Val(.SubItems(EI_SumTicketNum)) + txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text
    .SubItems(EI_TotalNum) = Val(.SubItems(EI_TotalNum)) + txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text
    .SubItems(EI_FullNum) = Val(.SubItems(EI_FullNum)) + txtFullSell.Text
    .SubItems(EI_HalfNum) = Val(.SubItems(EI_HalfNum)) + txtHalfSell.Text
    .SubItems(EI_SeatNo) = Trim(.SubItems(EI_SeatNo)) & "," & Trim(txtSeat.Text)
    .SubItems(EI_SeatNo) = Trim(.SubItems(EI_SeatNo)) & "," & Trim("")
    If cboPreferentialTicket.ListCount > 0 Then
        Select Case m_atTicketType(cboPreferentialTicket.ListIndex + 1).nTicketTypeID
        Case TP_FreeTicket
            .SubItems(EI_FreeNum) = txtPreferentialSell.Text + Val(.SubItems(EI_FreeNum))
        Case TP_PreferentialTicket1
            .SubItems(EI_PreferentialNum1) = txtPreferentialSell.Text + Val(.SubItems(EI_PreferentialNum1))
        Case TP_PreferentialTicket2
            .SubItems(EI_PreferentialNum2) = txtPreferentialSell.Text + Val(.SubItems(EI_PreferentialNum2))
        Case TP_PreferentialTicket3
            .SubItems(EI_PreferentialNum3) = txtPreferentialSell.Text + Val(.SubItems(EI_PreferentialNum3))
        End Select
    End If
    
   
    .SubItems(EI_TotalPrice) = Val(.SubItems(EI_TotalPrice)) + _
                                txtFullSell.Text * lvBus.SelectedItem.SubItems(ID_FullPrice) + _
                                txtHalfSell.Text * lvBus.SelectedItem.SubItems(ID_HalfPrice) + _
                                txtPreferentialSell.Text * GetPreferentialPrice
End With
End Sub

Private Function GetSeatTypeName(pszSeatTypeID As String) As String
    Dim i As Integer
    Dim nlen As Integer
    Dim szTemp As String
    nlen = ArrayLength(m_aszSeatType)
    For i = 1 To nlen
        If m_aszSeatType(i, 1) = pszSeatTypeID Then
            szTemp = Space(1) & Trim(m_aszSeatType(i, 2))
            Exit For
        End If
    Next
    GetSeatTypeName = szTemp
End Function
Private Sub ShowRightSeatType()
    Dim liTemp As ListItem
    If Not ptExtraByBus.Visible Then
        If cboSeatType.ListCount = 0 Then Exit Sub
        If Not lvBus.SelectedItem Is Nothing And Me.ActiveControl Is lvBus Then
            Set liTemp = lvBus.SelectedItem
            If liTemp.SubItems(ID_FullPrice) <> 0 Then
                cboSeatType.ListIndex = 0
            Else
                If liTemp.SubItems(ID_BedFullPrice) <> 0 Then
                    cboSeatType.ListIndex = 1
                Else
                    cboSeatType.ListIndex = 2
                End If
            End If
        End If
    Else
        If cboSeatType.ListCount = 0 Then Exit Sub
        If Not lvBus.SelectedItem Is Nothing And Me.ActiveControl Is lvBus Then
            Set liTemp = lvStation.SelectedItem
            If liTemp.SubItems(ID_FullPrice2) <> 0 Then
                cboSeatType.ListIndex = 0
            Else
                If liTemp.SubItems(ID_BedFullPrice2) <> 0 Then
                    cboSeatType.ListIndex = 1
                Else
                    cboSeatType.ListIndex = 2
                End If
            End If
        End If
    End If
End Sub

Private Sub ExtraSell()
    Dim i As Integer
    Dim nCount As Integer
    Dim nTicketCount As Integer
    Dim nlen As Integer
    Dim aspSellTicket() As TSellTicketParam
    Dim dBusDate() As Date
    Dim szBusID() As String
    Dim szDesStationID() As String
    Dim szDesStationName() As String
    Dim szSellStationID() As String
    Dim szSellStationName() As String
    Dim szStartStationName As String
    
    Dim asrSellResult() As TSellTicketResult
    Dim psgDiscount() As Single
    Dim apiTicketInfo() As TPrintTicketParam
    Dim pszBusDate() As String
    Dim pnTicketCount() As Integer
    Dim pszEndStation() As String
    Dim pszOffTime() As String
    Dim pszBusID() As String
    Dim pszVehicleType() As String
    Dim pszCheckGate() As String
    Dim pbSaleChange() As Boolean
    Dim pszTerminateName() As String
    
    Dim anInsurance() As Integer
    Dim panInsurance() As Integer
    
    Dim liTemp As ListItem
    Dim nTemp As Integer
    Dim dbTotalPrice As Double
    Dim dbRealTotalPrice As Double
    
    Dim nTotal As Integer
    Dim n As Integer
    
    dbTotalPrice = 0
    dbRealTotalPrice = 0
    dbTotalPrice = FormatMoney(flblTotalPrice.Caption)
    nTicketCount = 0
    nlen = 0
    
    If txtDiscount.Text > 1 Then
        MsgBox "折扣率不能大于1", vbInformation, "提示"
        txtDiscount.SetFocus
        Exit Sub
    End If
    On Error GoTo here
'以下是真正的售票处理
'-------------------------------------------------------------------------------------
    'ShowStatusInMDI "正在处理补票"
    lblSellMsg.Caption = "正在处理补票"
    DoEvents
    
    cmdSell.Enabled = False
    
    If ptExtraByBus.Visible Then
        If lvBusPreSell.ListItems.count = 0 Then
            If txtPreferentialSell.Text + txtHalfSell.Text + txtFullSell.Text = 0 Then
                lblSellMsg.Caption = ""
                cmdSell.Enabled = True
                Exit Sub
            End If
            ReDim asrSellResult(1 To 1)
            ReDim dBusDate(1 To 1)
            ReDim szBusID(1 To 1)
            ReDim szDesStationID(1 To 1)
            ReDim szDesStationName(1 To 1)
            ReDim szSellStationID(1 To 1)
            ReDim szSellStationName(1 To 1)
            ReDim anInsurance(1 To 1)
            
            ReDim psgDiscount(1 To 1)
            ReDim aspSellTicket(1 To 1)
            ReDim aspSellTicket(1).BuyTicketInfo(1 To txtPreferentialSell.Text + txtHalfSell.Text + txtFullSell.Text)
            ReDim aspSellTicket(1).pasgSellTicketPrice(1 To txtPreferentialSell.Text + txtHalfSell.Text + txtFullSell.Text)
            
            Set liTemp = lvStation.SelectedItem
            
            For i = 1 To txtFullSell.Text
                aspSellTicket(1).BuyTicketInfo(i).nTicketType = TP_FullPrice
                aspSellTicket(1).BuyTicketInfo(i).szTicketNo = GetTicketNo(i - 1)
                aspSellTicket(1).BuyTicketInfo(i).szSeatNo = SelfGetSeatNo(i)
                aspSellTicket(1).pasgSellTicketPrice(i) = CDbl(liTemp.SubItems(ID_FullPrice2)) * CSng(txtDiscount.Text)
                aspSellTicket(1).BuyTicketInfo(i).szSeatTypeID = m_aszSeatType(cboSeatType.ListIndex + 1, 1)
                aspSellTicket(1).BuyTicketInfo(i).szSeatTypeName = GetSeatTypeName(cboSeatType.Text)
            Next
            
            For i = 1 To txtHalfSell.Text
                aspSellTicket(1).BuyTicketInfo(i + txtFullSell.Text).nTicketType = TP_HalfPrice
                aspSellTicket(1).BuyTicketInfo(i + txtFullSell.Text).szTicketNo = GetTicketNo(i + txtFullSell.Text - 1)
                aspSellTicket(1).BuyTicketInfo(i + txtFullSell.Text).szSeatNo = SelfGetSeatNo(i + txtFullSell.Text)
                aspSellTicket(1).pasgSellTicketPrice(i + txtFullSell.Text) = CDbl(liTemp.SubItems(ID_HalfPrice2)) * CSng(txtDiscount.Text)
                aspSellTicket(1).BuyTicketInfo(i + txtFullSell.Text).szSeatTypeID = m_aszSeatType(cboSeatType.ListIndex + 1, 1)
                aspSellTicket(1).BuyTicketInfo(i + txtFullSell.Text).szSeatTypeName = GetSeatTypeName(cboSeatType.Text)
            Next
            
            For i = 1 To txtPreferentialSell.Text
                aspSellTicket(1).BuyTicketInfo(i + txtFullSell.Text + txtHalfSell.Text).nTicketType = m_atTicketType(cboPreferentialTicket.ListIndex + 1).nTicketTypeID
                aspSellTicket(1).BuyTicketInfo(i + txtFullSell.Text + txtHalfSell.Text).szTicketNo = GetTicketNo(i + txtFullSell.Text + txtHalfSell.Text - 1)
                aspSellTicket(1).BuyTicketInfo(i + txtFullSell.Text + txtHalfSell.Text).szSeatNo = SelfGetSeatNo(i + txtFullSell.Text + txtHalfSell.Text)
                aspSellTicket(1).pasgSellTicketPrice(i + txtFullSell.Text + txtHalfSell.Text) = CDbl(GetPreferentialPrice(True)) * CSng(txtDiscount.Text)
                aspSellTicket(1).BuyTicketInfo(i + txtFullSell.Text + txtHalfSell.Text).szSeatTypeID = m_aszSeatType(cboSeatType.ListIndex + 1, 1)
                aspSellTicket(1).BuyTicketInfo(i + txtFullSell.Text + txtHalfSell.Text).szSeatTypeName = GetSeatTypeName(cboSeatType.Text)
            Next
            dBusDate(1) = Date
            szBusID(1) = txtBusID.Text
            szDesStationID(1) = lvStation.SelectedItem.Text
            szDesStationName(1) = ""
            psgDiscount(1) = CSng(txtDiscount.Text)
            
            szSellStationID(1) = ResolveDisplay(lvStation.SelectedItem.Tag, szStartStationName)
            szSellStationName(1) = szStartStationName
            
            anInsurance(1) = IIf(chkInsurance.Value = vbChecked, 2, 0)   '如果选中,则赋为1,否则为0
            
            nTotal = txtPreferentialSell.Text + txtHalfSell.Text + txtFullSell.Text
            '判断售票数是否与实名制张数一致
            If nTotal <> ArrayLength(m_aszRealNameInfo) Then
                MsgBox "证件数[" & ArrayLength(m_aszRealNameInfo) & "]张与售票数[" & nTotal & "]张不符！", vbExclamation, App.Title
                GoTo out
            End If
            
            asrSellResult = m_oSell.ExtraSellTicket(dBusDate, szBusID, szSellStationID, szDesStationID, szDesStationName, aspSellTicket, anInsurance, m_aszRealNameInfo)
          
            IncTicketNo txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text
            
            '以下处理打印车票
            '-----------------------------------------------
            ReDim apiTicketInfo(1 To 1)
            ReDim pszBusDate(1 To 1)
            ReDim pnTicketCount(1 To 1)
            ReDim pszEndStation(1 To 1)
            ReDim pszOffTime(1 To 1)
            ReDim pszBusID(1 To 1)
            ReDim pszVehicleType(1 To 1)
            ReDim pszCheckGate(1 To 1)
            ReDim pbSaleChange(1 To 1)
            ReDim apiTicketInfo(1).aptPrintTicketInfo(1 To txtPreferentialSell.Text + txtHalfSell.Text + txtFullSell.Text)
            
            ReDim pszTerminateName(1 To 1)
            ReDim panInsurance(1 To 1)
            
            lblSellMsg.Refresh
            pszBusDate(1) = m_oParam.NowDateTime
            pnTicketCount(1) = txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text
            pszEndStation(1) = lvStation.SelectedItem.SubItems(ID_StationName)
            pszOffTime(1) = Format(CDate(flblOffTimeExtra.Caption), "hh:mm")
            pszBusID(1) = txtBusID.Text
            pszVehicleType(1) = flblBusTypeExtra.Caption
            pszCheckGate(1) = m_szCheckGate
            pszTerminateName(1) = Trim(lvStation.SelectedItem.SubItems(ID_TerminateName2))
            
            pbSaleChange(1) = False
            For i = 1 To txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text
                apiTicketInfo(1).aptPrintTicketInfo(i).nTicketType = aspSellTicket(1).BuyTicketInfo(i).nTicketType
                apiTicketInfo(1).aptPrintTicketInfo(i).sgTicketPrice = asrSellResult(1).asgTicketPrice(i)
                apiTicketInfo(1).aptPrintTicketInfo(i).szSeatNo = asrSellResult(1).aszSeatNo(i)
                apiTicketInfo(1).aptPrintTicketInfo(i).szTicketNo = aspSellTicket(1).BuyTicketInfo(i).szTicketNo
                
                If aspSellTicket(1).BuyTicketInfo(i).nTicketType <> TP_FreeTicket Then
                    dbRealTotalPrice = dbRealTotalPrice + asrSellResult(1).asgTicketPrice(i)
                End If
            Next i
            
            panInsurance(1) = IIf(chkInsurance.Value = vbChecked, 2, 0)   '如果选中,则赋为1,否则为0
            
            lblSellMsg.Caption = "正在打印车票"
            lblSellMsg.Refresh
            PrintTicket apiTicketInfo, pszBusDate, pnTicketCount, pszEndStation, pszOffTime, pszBusID, pszVehicleType, pszCheckGate, pbSaleChange, pszTerminateName, szSellStationName, panInsurance, m_aszRealNameInfo
            m_aszInsurce = CombInsurance(apiTicketInfo, szBusID, pnTicketCount, pszBusDate, szDesStationID, pszEndStation, pszOffTime, szSellStationID, szSellStationName, m_aszRealNameInfo)
            SaveInsurance m_aszInsurce
            
            ShowBusInfo
        Else
            nlen = lvBusPreSell.ListItems.count
    
            ReDim dBusDate(1 To nlen)
            ReDim szBusID(1 To nlen)
            ReDim szDesStationID(1 To nlen)
            ReDim szDesStationName(1 To nlen)
            ReDim psgDiscount(1 To nlen)
            ReDim asrSellResult(1 To nlen)
            ReDim aspSellTicket(1 To nlen)
            ReDim szSellStationID(1 To nlen)
            ReDim szSellStationName(1 To nlen)
            ReDim anInsurance(1 To nlen)
            
            
            For nCount = 1 To lvBusPreSell.ListItems.count
                With lvBusPreSell.ListItems(nCount)
                    ReDim aspSellTicket(nCount).BuyTicketInfo(1 To .SubItems(EI_TotalNum))
                    ReDim aspSellTicket(nCount).pasgSellTicketPrice(1 To .SubItems(EI_TotalNum))
                    nTemp = 0
                    For i = 1 To Val(.SubItems(EI_FullNum))
                        aspSellTicket(nCount).BuyTicketInfo(i).nTicketType = TP_FullPrice
                        aspSellTicket(nCount).BuyTicketInfo(i).szTicketNo = GetTicketNo(i - 1 + nTicketCount)
                        aspSellTicket(nCount).BuyTicketInfo(i).szSeatNo = SelfGetSeatNo12(i, .SubItems(EI_SeatStatus1), .SubItems(EI_SeatStatus2), .SubItems(EI_SeatNo))
                        aspSellTicket(nCount).pasgSellTicketPrice(i) = CDbl(.SubItems(EI_FullPrice))
                        aspSellTicket(nCount).BuyTicketInfo(i).szSeatTypeID = .SubItems(EI_SeatType)
                        aspSellTicket(nCount).BuyTicketInfo(i).szSeatTypeName = GetSeatTypeName(.SubItems(EI_SeatType))
                    Next i
                    
                    nTemp = nTemp + Val(.SubItems(EI_FullNum))
                    For i = 1 To Val(.SubItems(EI_HalfNum))
                        aspSellTicket(nCount).BuyTicketInfo(i + nTemp).nTicketType = TP_HalfPrice
                        aspSellTicket(nCount).BuyTicketInfo(i + nTemp).szTicketNo = GetTicketNo(i - 1 + nTemp + nTicketCount)
                        aspSellTicket(nCount).BuyTicketInfo(i + nTemp).szSeatNo = SelfGetSeatNo12(i + nTemp, .SubItems(EI_SeatStatus1), .SubItems(EI_SeatStatus2), .SubItems(EI_SeatNo))
                        aspSellTicket(nCount).pasgSellTicketPrice(i + nTemp) = CDbl(.SubItems(EI_HalfPrice))
                        aspSellTicket(nCount).BuyTicketInfo(i + nTemp).szSeatTypeID = .SubItems(EI_SeatType)
                        aspSellTicket(nCount).BuyTicketInfo(i + nTemp).szSeatTypeName = GetSeatTypeName(.SubItems(EI_SeatType))
                    Next i
                    
                    nTemp = nTemp + Val(.SubItems(EI_HalfNum))
                    For i = 1 To Val(.SubItems(EI_FreeNum))
                        aspSellTicket(nCount).BuyTicketInfo(i + nTemp).nTicketType = TP_FreeTicket
                        aspSellTicket(nCount).BuyTicketInfo(i + nTemp).szTicketNo = GetTicketNo(i - 1 + nTemp + nTicketCount)
                        aspSellTicket(nCount).BuyTicketInfo(i + nTemp).szSeatNo = SelfGetSeatNo12(i + nTemp, .SubItems(EI_SeatStatus1), .SubItems(EI_SeatStatus2), .SubItems(EI_SeatNo))
                        aspSellTicket(nCount).pasgSellTicketPrice(i + nTemp) = CDbl(.SubItems(EI_FullPrice))
                        aspSellTicket(nCount).BuyTicketInfo(i + nTemp).szSeatTypeID = .SubItems(EI_SeatType)
                        aspSellTicket(nCount).BuyTicketInfo(i + nTemp).szSeatTypeName = GetSeatTypeName(.SubItems(EI_SeatType))
                    Next i
                    
                    nTemp = nTemp + Val(.SubItems(EI_FreeNum))
                    For i = 1 To Val(.SubItems(EI_PreferentialNum1))
                        aspSellTicket(nCount).BuyTicketInfo(i + nTemp).nTicketType = TP_PreferentialTicket1
                        aspSellTicket(nCount).BuyTicketInfo(i + nTemp).szTicketNo = GetTicketNo(i - 1 + nTemp + nTicketCount)
                        aspSellTicket(nCount).BuyTicketInfo(i + nTemp).szSeatNo = SelfGetSeatNo12(i + nTemp, .SubItems(EI_SeatStatus1), .SubItems(EI_SeatStatus2), .SubItems(EI_SeatNo))
                        aspSellTicket(nCount).pasgSellTicketPrice(i + nTemp) = CDbl(.SubItems(EI_PreferentialPrice1))
                        aspSellTicket(nCount).BuyTicketInfo(i + nTemp).szSeatTypeID = .SubItems(EI_SeatType)
                        aspSellTicket(nCount).BuyTicketInfo(i + nTemp).szSeatTypeName = GetSeatTypeName(.SubItems(EI_SeatType))
                    Next i
                    
                    nTemp = nTemp + Val(.SubItems(EI_PreferentialNum1))
                    For i = 1 To Val(.SubItems(EI_PreferentialNum2))
                        aspSellTicket(nCount).BuyTicketInfo(i + nTemp).nTicketType = TP_PreferentialTicket2
                        aspSellTicket(nCount).BuyTicketInfo(i + nTemp).szTicketNo = GetTicketNo(i - 1 + nTemp + nTicketCount)
                        aspSellTicket(nCount).BuyTicketInfo(i + nTemp).szSeatNo = SelfGetSeatNo12(i + nTemp, .SubItems(EI_SeatStatus1), .SubItems(EI_SeatStatus2), .SubItems(EI_SeatNo))
                        aspSellTicket(nCount).pasgSellTicketPrice(i + nTemp) = CDbl(.SubItems(EI_PreferentialPrice2))
                        aspSellTicket(nCount).BuyTicketInfo(i + nTemp).szSeatTypeID = .SubItems(EI_SeatType)
                        aspSellTicket(nCount).BuyTicketInfo(i + nTemp).szSeatTypeName = GetSeatTypeName(.SubItems(EI_SeatType))
                    Next i
                    
                    nTemp = nTemp + Val(.SubItems(EI_PreferentialNum2))
                    For i = 1 To Val(.SubItems(EI_PreferentialNum3))
                        aspSellTicket(nCount).BuyTicketInfo(i + nTemp).nTicketType = TP_PreferentialTicket3
                        aspSellTicket(nCount).BuyTicketInfo(i + nTemp).szTicketNo = GetTicketNo(i - 1 + nTemp + nTicketCount)
                        aspSellTicket(nCount).BuyTicketInfo(i + nTemp).szSeatNo = SelfGetSeatNo12(i + nTemp, .SubItems(EI_SeatStatus1), .SubItems(EI_SeatStatus2), .SubItems(EI_SeatNo))
                        aspSellTicket(nCount).pasgSellTicketPrice(i + nTemp) = CDbl(.SubItems(EI_PreferentialPrice3))
                        aspSellTicket(nCount).BuyTicketInfo(i + nTemp).szSeatTypeID = .SubItems(EI_SeatType)
                        aspSellTicket(nCount).BuyTicketInfo(i + nTemp).szSeatTypeName = GetSeatTypeName(.SubItems(EI_SeatType))
                    Next i
                    
                    dBusDate(nCount) = m_oParam.NowDate
                    szBusID(nCount) = .Text
                    szDesStationID(nCount) = .SubItems(EI_EndStationCode)
                    szDesStationName(nCount) = ""
                    psgDiscount(nCount) = CSng(.SubItems(EI_Discount))
                    nTicketCount = nTicketCount + .SubItems(EI_TotalNum)
                    
                    anInsurance(nCount) = IIf(chkInsurance.Value = vbChecked, 2, 0)    '如果选中,则赋为1,否则为0
            
                End With
                If lvStation.ListItems.count < nCount Then
                    szSellStationID(nCount) = ResolveDisplay(lvStation.ListItems(lvStation.ListItems.count).Tag, szStartStationName)
                    szSellStationName(nCount) = szStartStationName
                Else
                    szSellStationID(nCount) = ResolveDisplay(lvStation.ListItems(nCount).Tag, szStartStationName)
                    szSellStationName(nCount) = szStartStationName
                 End If
            Next nCount
            
            nTotal = nTicketCount
            '判断售票数是否与实名制张数一致
            If nTotal <> ArrayLength(m_aszRealNameInfo) Then
                MsgBox "证件数[" & ArrayLength(m_aszRealNameInfo) & "]张与售票数[" & nTotal & "]张不符！", vbExclamation, App.Title
                GoTo out
            End If
            
            asrSellResult = m_oSell.ExtraSellTicket(dBusDate, szBusID, szSellStationID, szDesStationID, szDesStationName, aspSellTicket, anInsurance, m_aszRealNameInfo)
          
            IncTicketNo nTicketCount
            
            '以下处理打印车票
            '-----------------------------------------------------------------
            ReDim apiTicketInfo(1 To nlen)
            ReDim pszBusDate(1 To nlen)
            ReDim pnTicketCount(1 To nlen)
            ReDim pszEndStation(1 To nlen)
            ReDim pszOffTime(1 To nlen)
            ReDim pszBusID(1 To nlen)
            ReDim pszVehicleType(1 To nlen)
            ReDim pszCheckGate(1 To nlen)
            ReDim pbSaleChange(1 To nlen)
            ReDim pszTerminateName(1 To nlen)
            ReDim panInsurance(1 To nlen)
            
            
            For nCount = 1 To lvBusPreSell.ListItems.count
                ReDim apiTicketInfo(nCount).aptPrintTicketInfo(1 To lvBusPreSell.ListItems(nCount).SubItems(EI_TotalNum))
                With lvBusPreSell.ListItems(nCount)
                    pszBusDate(nCount) = m_oParam.NowDateTime
                    pnTicketCount(nCount) = .SubItems(EI_TotalNum)
                    pszEndStation(nCount) = .SubItems(EI_EndStation)
                    pszOffTime(nCount) = .SubItems(EI_OffTime)
                    pszBusID(nCount) = .Text
                    pszVehicleType(nCount) = .SubItems(EI_VehicleModel)
                    pszCheckGate(nCount) = .SubItems(EI_CheckGate)
                    pszTerminateName(nCount) = Trim(.SubItems(EI_TerminateName))
                    pbSaleChange(nCount) = False
                    
                    panInsurance(nCount) = IIf(chkInsurance.Value = vbChecked, 2, 0)    '如果选中,则赋为1,否则为0
            
                    For i = 1 To .SubItems(EI_TotalNum)
                        apiTicketInfo(nCount).aptPrintTicketInfo(i).nTicketType = aspSellTicket(nCount).BuyTicketInfo(i).nTicketType
                        apiTicketInfo(nCount).aptPrintTicketInfo(i).sgTicketPrice = asrSellResult(nCount).asgTicketPrice(i)
                        apiTicketInfo(nCount).aptPrintTicketInfo(i).szSeatNo = asrSellResult(nCount).aszSeatNo(i)
                        apiTicketInfo(nCount).aptPrintTicketInfo(i).szTicketNo = aspSellTicket(nCount).BuyTicketInfo(i).szTicketNo
                        If aspSellTicket(nCount).BuyTicketInfo(i).nTicketType <> TP_FreeTicket Then
                            dbRealTotalPrice = dbRealTotalPrice + asrSellResult(nCount).asgTicketPrice(i)
                        End If
                    Next i
                End With
            Next nCount
            lblSellMsg.Caption = "正在打印车票"
            lblSellMsg.Refresh
            
            PrintTicket apiTicketInfo, pszBusDate, pnTicketCount, pszEndStation, pszOffTime, pszBusID, pszVehicleType, pszCheckGate, pbSaleChange, pszTerminateName, szSellStationName, panInsurance, m_aszRealNameInfo
            
            m_aszInsurce = CombInsurance(apiTicketInfo, szBusID, pnTicketCount, pszBusDate, szDesStationID, pszEndStation, pszOffTime, szSellStationID, szSellStationName, m_aszRealNameInfo)
            SaveInsurance m_aszInsurce
            
            lvBusPreSell.ListItems.Clear
            ShowBusInfo
        End If
    Else
        If lvStationPreSell.ListItems.count = 0 Then
            If txtPreferentialSell.Text + txtHalfSell.Text + txtFullSell.Text = 0 Then
                lblSellMsg.Caption = ""
                cmdSell.Enabled = True
                Exit Sub
            End If
            ReDim szSellStationID(1 To 1)
            ReDim szSellStationName(1 To 1)
            
            
            ReDim asrSellResult(1 To 1)
            ReDim dBusDate(1 To 1)
            ReDim szBusID(1 To 1)
            ReDim szDesStationID(1 To 1)
            ReDim szStationName(1 To 1)
            ReDim psgDiscount(1 To 1)
            ReDim aspSellTicket(1 To 1)
            ReDim aspSellTicket(1).BuyTicketInfo(1 To txtPreferentialSell.Text + txtHalfSell.Text + txtFullSell.Text)
            ReDim aspSellTicket(1).pasgSellTicketPrice(1 To txtPreferentialSell.Text + txtHalfSell.Text + txtFullSell.Text)
            ReDim anInsurance(1 To 1)
            Set liTemp = lvBus.SelectedItem
            
            For i = 1 To txtFullSell.Text
                aspSellTicket(1).BuyTicketInfo(i).nTicketType = TP_FullPrice
                aspSellTicket(1).BuyTicketInfo(i).szTicketNo = GetTicketNo(i - 1)
                aspSellTicket(1).BuyTicketInfo(i).szSeatNo = SelfGetSeatNo(i)
                aspSellTicket(1).pasgSellTicketPrice(i) = CDbl(liTemp.SubItems(ID_FullPrice)) * CSng(txtDiscount.Text)
                aspSellTicket(1).BuyTicketInfo(i).szSeatTypeID = m_aszSeatType(cboSeatType.ListIndex + 1, 1)
                aspSellTicket(1).BuyTicketInfo(i).szSeatTypeName = GetSeatTypeName(cboSeatType.Text)
            Next
            
            For i = 1 To txtHalfSell.Text
                aspSellTicket(1).BuyTicketInfo(i + txtFullSell.Text).nTicketType = TP_HalfPrice
                aspSellTicket(1).BuyTicketInfo(i + txtFullSell.Text).szTicketNo = GetTicketNo(i + txtFullSell.Text - 1)
                aspSellTicket(1).BuyTicketInfo(i + txtFullSell.Text).szSeatNo = SelfGetSeatNo(i + txtFullSell.Text)
                aspSellTicket(1).pasgSellTicketPrice(i + txtFullSell.Text) = CDbl(liTemp.SubItems(ID_HalfPrice)) * CSng(txtDiscount.Text)
                aspSellTicket(1).BuyTicketInfo(i + txtFullSell.Text).szSeatTypeID = m_aszSeatType(cboSeatType.ListIndex + 1, 1)
                aspSellTicket(1).BuyTicketInfo(i + txtFullSell.Text).szSeatTypeName = GetSeatTypeName(cboSeatType.Text)
            Next
            
            For i = 1 To txtPreferentialSell.Text
                aspSellTicket(1).BuyTicketInfo(i + txtFullSell.Text + txtHalfSell.Text).nTicketType = m_atTicketType(cboPreferentialTicket.ListIndex + 1).nTicketTypeID
                aspSellTicket(1).BuyTicketInfo(i + txtFullSell.Text + txtHalfSell.Text).szTicketNo = GetTicketNo(i + txtFullSell.Text + txtHalfSell.Text - 1)
                aspSellTicket(1).BuyTicketInfo(i + txtFullSell.Text + txtHalfSell.Text).szSeatNo = SelfGetSeatNo(i + txtFullSell.Text + txtHalfSell.Text)
                aspSellTicket(1).pasgSellTicketPrice(i + txtFullSell.Text + txtHalfSell.Text) = GetPreferentialPrice(True) * CSng(txtDiscount.Text)
                aspSellTicket(1).BuyTicketInfo(i + txtFullSell.Text + txtHalfSell.Text).szSeatTypeID = m_aszSeatType(cboSeatType.ListIndex + 1, 1)
                aspSellTicket(1).BuyTicketInfo(i + txtFullSell.Text + txtHalfSell.Text).szSeatTypeName = GetSeatTypeName(cboSeatType.Text)
            Next
            dBusDate(1) = m_oParam.NowDate
            szBusID(1) = lvBus.SelectedItem.Text
            szDesStationID(1) = cboEndStation.BoundText
            szStationName(1) = ""
            
            anInsurance(1) = IIf(chkInsurance.Value = vbChecked, 2, 0)   '如果选中,则赋为1,否则为0
            
            psgDiscount(1) = CSng(txtDiscount.Text)
                szSellStationID(1) = ResolveDisplay(lvBus.SelectedItem.Tag, szStartStationName)
                szSellStationName(1) = szStartStationName
                
                nTotal = txtPreferentialSell.Text + txtHalfSell.Text + txtFullSell.Text
                '判断售票数是否与实名制张数一致
                If nTotal <> ArrayLength(m_aszRealNameInfo) Then
                    MsgBox "证件数[" & ArrayLength(m_aszRealNameInfo) & "]张与售票数[" & nTotal & "]张不符！", vbExclamation, App.Title
                    GoTo out
                End If
                
                asrSellResult = m_oSell.ExtraSellTicket(dBusDate, szBusID, szSellStationID, szDesStationID, szDesStationName, aspSellTicket, anInsurance, m_aszRealNameInfo)
            
            IncTicketNo txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text
            
            '以下处理打印车票
            '------------------------------------------------------------------------
            ReDim apiTicketInfo(1 To 1)
            ReDim pszBusDate(1 To 1)
            ReDim pnTicketCount(1 To 1)
            ReDim pszEndStation(1 To 1)
            ReDim pszOffTime(1 To 1)
            ReDim pszBusID(1 To 1)
            ReDim pszVehicleType(1 To 1)
            ReDim pszCheckGate(1 To 1)
            ReDim pbSaleChange(1 To 1)
            ReDim apiTicketInfo(1).aptPrintTicketInfo(1 To txtPreferentialSell.Text + txtHalfSell.Text + txtFullSell.Text)
            ReDim pszTerminateName(1 To 1)
            ReDim panInsurance(1 To 1)
            
            lblSellMsg.Refresh
            pszBusDate(1) = m_oParam.NowDateTime
            pnTicketCount(1) = txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text
            pszEndStation(1) = GetStationNameInCbo(cboEndStation.Text)
            pszOffTime(1) = Format(CDate(lvBus.SelectedItem.SubItems(ID_OffTime)), "hh:mm")
            pszBusID(1) = lvBus.SelectedItem.Text
            pszVehicleType(1) = lvBus.SelectedItem.SubItems(ID_VehicleModel)
            pszCheckGate(1) = GetCheckName(lvBus.SelectedItem.SubItems(ID_CheckGate))
            pszTerminateName(1) = lvBus.SelectedItem.SubItems(ID_EndStation)
            
            panInsurance(1) = IIf(chkInsurance.Value = vbChecked, 2, 0)   '如果选中,则赋为1,否则为0
            
            pbSaleChange(1) = False
            For i = 1 To txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text
                apiTicketInfo(1).aptPrintTicketInfo(i).nTicketType = aspSellTicket(1).BuyTicketInfo(i).nTicketType
                apiTicketInfo(1).aptPrintTicketInfo(i).sgTicketPrice = asrSellResult(1).asgTicketPrice(i)
                apiTicketInfo(1).aptPrintTicketInfo(i).szSeatNo = asrSellResult(1).aszSeatNo(i)
                apiTicketInfo(1).aptPrintTicketInfo(i).szTicketNo = aspSellTicket(1).BuyTicketInfo(i).szTicketNo
                If aspSellTicket(1).BuyTicketInfo(i).nTicketType <> TP_FreeTicket Then
                    dbRealTotalPrice = dbRealTotalPrice + asrSellResult(1).asgTicketPrice(i)
                End If
            Next i
            
            lblSellMsg.Caption = "正在打印车票"
            lblSellMsg.Refresh
            
            PrintTicket apiTicketInfo, pszBusDate, pnTicketCount, pszEndStation, pszOffTime, pszBusID, pszVehicleType, pszCheckGate, pbSaleChange, pszTerminateName, szSellStationName, panInsurance, m_aszRealNameInfo
            m_aszInsurce = CombInsurance(apiTicketInfo, szBusID, pnTicketCount, pszBusDate, szDesStationID, pszEndStation, pszOffTime, szSellStationID, szSellStationName, m_aszRealNameInfo)
            SaveInsurance m_aszInsurce
            
            RefreshBus True
        Else
            nlen = lvStationPreSell.ListItems.count
            
            ReDim szSellStationID(1 To nlen)
            ReDim szSellStationName(1 To nlen)
            ReDim dBusDate(1 To nlen)
            ReDim szBusID(1 To nlen)
            ReDim szDesStationID(1 To nlen)
            ReDim szDesStationName(1 To nlen)
            ReDim psgDiscount(1 To nlen)
            ReDim asrSellResult(1 To nlen)
            ReDim aspSellTicket(1 To nlen)
            ReDim panInsurance(1 To nlen)
            For nCount = 1 To lvStationPreSell.ListItems.count
                With lvStationPreSell.ListItems(nCount)
                    ReDim aspSellTicket(nCount).BuyTicketInfo(1 To .SubItems(EI_TotalNum))
                    ReDim aspSellTicket(nCount).pasgSellTicketPrice(1 To .SubItems(EI_TotalNum))
                    
                    nTemp = 0
                    For i = 1 To Val(.SubItems(EI_FullNum))
                        aspSellTicket(nCount).BuyTicketInfo(i).nTicketType = TP_FullPrice
                        aspSellTicket(nCount).BuyTicketInfo(i).szTicketNo = GetTicketNo(i - 1 + nTicketCount)
                        aspSellTicket(nCount).BuyTicketInfo(i).szSeatNo = SelfGetSeatNo12(i, .SubItems(EI_SeatStatus1), .SubItems(EI_SeatStatus2), .SubItems(EI_SeatNo))
                        aspSellTicket(nCount).pasgSellTicketPrice(i) = CDbl(.SubItems(EI_FullPrice))
                        aspSellTicket(nCount).BuyTicketInfo(i).szSeatTypeID = .SubItems(EI_SeatType)
                        aspSellTicket(nCount).BuyTicketInfo(i).szSeatTypeName = GetSeatTypeName(.SubItems(EI_SeatType))
                    Next i
                    
                    nTemp = nTemp + Val(.SubItems(EI_FullNum))
                    For i = 1 To Val(.SubItems(EI_HalfNum))
                        aspSellTicket(nCount).BuyTicketInfo(i + nTemp).nTicketType = TP_HalfPrice
                        aspSellTicket(nCount).BuyTicketInfo(i + nTemp).szTicketNo = GetTicketNo(i - 1 + nTemp + nTicketCount)
                        aspSellTicket(nCount).BuyTicketInfo(i + nTemp).szSeatNo = SelfGetSeatNo12(i + nTemp, .SubItems(EI_SeatStatus1), .SubItems(EI_SeatStatus2), .SubItems(EI_SeatNo))
                        aspSellTicket(nCount).pasgSellTicketPrice(i + nTemp) = CDbl(.SubItems(EI_HalfPrice))
                        aspSellTicket(nCount).BuyTicketInfo(i + nTemp).szSeatTypeID = .SubItems(EI_SeatType)
                        aspSellTicket(nCount).BuyTicketInfo(i + nTemp).szSeatTypeName = GetSeatTypeName(.SubItems(EI_SeatType))
                    Next i
                    
                    nTemp = nTemp + Val(.SubItems(EI_HalfNum))
                    For i = 1 To Val(.SubItems(EI_FreeNum))
                        aspSellTicket(nCount).BuyTicketInfo(i + nTemp).nTicketType = TP_FreeTicket
                        aspSellTicket(nCount).BuyTicketInfo(i + nTemp).szTicketNo = GetTicketNo(i - 1 + nTemp + nTicketCount)
                        aspSellTicket(nCount).BuyTicketInfo(i + nTemp).szSeatNo = SelfGetSeatNo12(i + nTemp, .SubItems(EI_SeatStatus1), .SubItems(EI_SeatStatus2), .SubItems(EI_SeatNo))
                        aspSellTicket(nCount).pasgSellTicketPrice(i + nTemp) = CDbl(.SubItems(EI_FullPrice))
                        aspSellTicket(nCount).BuyTicketInfo(i + nTemp).szSeatTypeID = .SubItems(EI_SeatType)
                        aspSellTicket(nCount).BuyTicketInfo(i + nTemp).szSeatTypeName = GetSeatTypeName(.SubItems(EI_SeatType))
                    Next i
                    
                    nTemp = nTemp + Val(.SubItems(EI_FreeNum))
                    For i = 1 To Val(.SubItems(EI_PreferentialNum1))
                        aspSellTicket(nCount).BuyTicketInfo(i + nTemp).nTicketType = TP_PreferentialTicket1
                        aspSellTicket(nCount).BuyTicketInfo(i + nTemp).szTicketNo = GetTicketNo(i - 1 + nTemp + nTicketCount)
                        aspSellTicket(nCount).BuyTicketInfo(i + nTemp).szSeatNo = SelfGetSeatNo12(i + nTemp, .SubItems(EI_SeatStatus1), .SubItems(EI_SeatStatus2), .SubItems(EI_SeatNo))
                        aspSellTicket(nCount).pasgSellTicketPrice(i + nTemp) = CDbl(.SubItems(EI_PreferentialPrice1))
                        aspSellTicket(nCount).BuyTicketInfo(i + nTemp).szSeatTypeID = .SubItems(EI_SeatType)
                        aspSellTicket(nCount).BuyTicketInfo(i + nTemp).szSeatTypeName = GetSeatTypeName(.SubItems(EI_SeatType))
                    Next i
                    
                    nTemp = nTemp + Val(.SubItems(EI_PreferentialNum1))
                    For i = 1 To Val(.SubItems(EI_PreferentialNum2))
                        aspSellTicket(nCount).BuyTicketInfo(i + nTemp).nTicketType = TP_PreferentialTicket2
                        aspSellTicket(nCount).BuyTicketInfo(i + nTemp).szTicketNo = GetTicketNo(i - 1 + nTemp + nTicketCount)
                        aspSellTicket(nCount).BuyTicketInfo(i + nTemp).szSeatNo = SelfGetSeatNo12(i + nTemp, .SubItems(EI_SeatStatus1), .SubItems(EI_SeatStatus2), .SubItems(EI_SeatNo))
                        aspSellTicket(nCount).pasgSellTicketPrice(i + nTemp) = CDbl(.SubItems(EI_PreferentialPrice2))
                        aspSellTicket(nCount).BuyTicketInfo(i + nTemp).szSeatTypeID = .SubItems(EI_SeatType)
                        aspSellTicket(nCount).BuyTicketInfo(i + nTemp).szSeatTypeName = GetSeatTypeName(.SubItems(EI_SeatType))
                    Next i
                    
                    nTemp = nTemp + Val(.SubItems(EI_PreferentialNum2))
                    For i = 1 To Val(.SubItems(EI_PreferentialNum3))
                        aspSellTicket(nCount).BuyTicketInfo(i + nTemp).nTicketType = TP_PreferentialTicket3
                        aspSellTicket(nCount).BuyTicketInfo(i + nTemp).szTicketNo = GetTicketNo(i - 1 + nTemp + nTicketCount)
                        aspSellTicket(nCount).BuyTicketInfo(i + nTemp).szSeatNo = SelfGetSeatNo12(i + nTemp, .SubItems(EI_SeatStatus1), .SubItems(EI_SeatStatus2), .SubItems(EI_SeatNo))
                        aspSellTicket(nCount).pasgSellTicketPrice(i + nTemp) = CDbl(.SubItems(EI_PreferentialPrice3))
                        aspSellTicket(nCount).BuyTicketInfo(i + nTemp).szSeatTypeID = .SubItems(EI_SeatType)
                        aspSellTicket(nCount).BuyTicketInfo(i + nTemp).szSeatTypeName = GetSeatTypeName(.SubItems(EI_SeatType))
                    Next i
                    dBusDate(nCount) = Date
                    szBusID(nCount) = .Text
                    szDesStationID(nCount) = .SubItems(EI_EndStationCode)
                    szDesStationName(nCount) = ""
                    psgDiscount(nCount) = CSng(.SubItems(EI_Discount))
                    nTicketCount = nTicketCount + .SubItems(EI_TotalNum)
                End With
                
                If lvBus.ListItems.count < nCount Then
                   szSellStationID(nCount) = ResolveDisplay(lvBus.ListItems(lvBus.ListItems.count).Tag, szStartStationName)
                   szSellStationName(nCount) = szStartStationName
                Else
                   szSellStationID(nCount) = ResolveDisplay(lvBus.ListItems(nCount).Tag, szStartStationName)
                   szSellStationName(nCount) = szStartStationName
                End If
                
                panInsurance(nCount) = IIf(chkInsurance.Value = vbChecked, 2, 0)    '如果选中,则赋为1,否则为0
            
            Next nCount
            
            nTotal = nTicketCount
            '判断售票数是否与实名制张数一致
            If nTotal <> ArrayLength(m_aszRealNameInfo) Then
                MsgBox "证件数[" & ArrayLength(m_aszRealNameInfo) & "]张与售票数[" & nTotal & "]张不符！", vbExclamation, App.Title
                GoTo out
            End If
            
            asrSellResult = m_oSell.ExtraSellTicket(dBusDate, szBusID, szSellStationID, szDesStationID, szDesStationName, aspSellTicket, panInsurance, m_aszRealNameInfo)
          
            IncTicketNo nTicketCount
             
             '以下处理打印车票
            '-----------------------------------------------------------------
            ReDim apiTicketInfo(1 To lvStationPreSell.ListItems.count)
            ReDim pszBusDate(1 To lvStationPreSell.ListItems.count)
            ReDim pnTicketCount(1 To lvStationPreSell.ListItems.count)
            ReDim pszEndStation(1 To lvStationPreSell.ListItems.count)
            ReDim pszOffTime(1 To lvStationPreSell.ListItems.count)
            ReDim pszBusID(1 To lvStationPreSell.ListItems.count)
            ReDim pszVehicleType(1 To lvStationPreSell.ListItems.count)
            ReDim pszCheckGate(1 To lvStationPreSell.ListItems.count)
            ReDim pbSaleChange(1 To lvStationPreSell.ListItems.count)
            ReDim pszTerminateName(1 To lvStationPreSell.ListItems.count)
            ReDim panInsurance(1 To lvStationPreSell.ListItems.count)
            For nCount = 1 To lvStationPreSell.ListItems.count
                ReDim apiTicketInfo(nCount).aptPrintTicketInfo(1 To lvStationPreSell.ListItems(nCount).SubItems(EI_TotalNum))
                With lvStationPreSell.ListItems(nCount)
                    pszBusDate(nCount) = Date
                    pnTicketCount(nCount) = .SubItems(EI_TotalNum)
                    pszEndStation(nCount) = .SubItems(EI_EndStation)
                    pszOffTime(nCount) = .SubItems(EI_OffTime)
                    pszBusID(nCount) = .Text
                    pszVehicleType(nCount) = .SubItems(EI_VehicleModel)
                    pszCheckGate(nCount) = GetCheckName(.SubItems(EI_CheckGate))
                    
                    pszTerminateName(nCount) = .SubItems(EI_TerminateName)
                    pbSaleChange(nCount) = False
                    
                    panInsurance(nCount) = IIf(chkInsurance.Value = vbChecked, 2, 0)    '如果选中,则赋为1,否则为0
            
                    For i = 1 To Val(.SubItems(EI_TotalNum))
                        apiTicketInfo(nCount).aptPrintTicketInfo(i).nTicketType = aspSellTicket(nCount).BuyTicketInfo(i).nTicketType
                        apiTicketInfo(nCount).aptPrintTicketInfo(i).sgTicketPrice = asrSellResult(nCount).asgTicketPrice(i)
                        apiTicketInfo(nCount).aptPrintTicketInfo(i).szSeatNo = asrSellResult(nCount).aszSeatNo(i)
                        apiTicketInfo(nCount).aptPrintTicketInfo(i).szTicketNo = aspSellTicket(nCount).BuyTicketInfo(i).szTicketNo
                        If aspSellTicket(nCount).BuyTicketInfo(i).nTicketType <> TP_FreeTicket Then
                            dbRealTotalPrice = dbRealTotalPrice + asrSellResult(nCount).asgTicketPrice(i)
                        End If
                    Next i
                End With
            Next nCount
            lblSellMsg.Caption = "正在打印车票"
            lblSellMsg.Refresh
            
            PrintTicket apiTicketInfo, pszBusDate, pnTicketCount, pszEndStation, pszOffTime, pszBusID, pszVehicleType, pszCheckGate, pbSaleChange, pszTerminateName, szSellStationName, panInsurance, m_aszRealNameInfo
            
            m_aszInsurce = CombInsurance(apiTicketInfo, szBusID, pnTicketCount, pszBusDate, szDesStationID, pszEndStation, pszOffTime, szSellStationID, szSellStationName, m_aszRealNameInfo)
            SaveInsurance m_aszInsurce
            
            lvStationPreSell.ListItems.Clear
'            RefreshBus True
        End If
    End If

'    '**********************************************************************
'   If Abs(dbRealTotalPrice - dbTotalPrice) > 0.01 Then
'        frmPriceInfo.m_sngTotalPrice = dbRealTotalPrice
'        frmPriceInfo.Show vbModal
'   End If
    'DoEvents
    '进行票款累加
    
    If IsNumeric(txtReceivedMoney.Text) Then
    
        If txtReceivedMoney.Text = 0 Then
            m_sgTotalMoney = lblTotalMoney.Caption
        Else
            m_sgTotalMoney = 0#
        End If
    End If
   

'-------------------------------------------------------------------------------------

    lblSellMsg.Caption = ""
'    SetDefaultSellTicket
    cmdSell.Enabled = True
    On Error GoTo 0
    
    Erase m_aszRealNameInfo
    
    Exit Sub
    
out:
        lblSellMsg.Caption = ""
        cmdSell.Enabled = True
        
        flblRestMoney.Caption = ""
        frmOrderSeats.m_szBookNumber = ""
        txtSeat.Text = ""

        SetPreSellButton

        'fpd
        Dim szRestMoney As String
        Dim szTotaMoney As String
        szRestMoney = flblRestMoney.Caption
        szTotaMoney = flblTotalPrice.Caption
        
        ClearInfo

        flblRestMoney.Caption = Format(szRestMoney, "0.00")
        flblTotalPrice.Caption = Format(szTotaMoney, "0.00")

    Exit Sub
    
here:
    lblSellMsg.Caption = ""
    cmdSell.Enabled = True
    ShowErrorMsg


End Sub

Public Sub ChangeSeatType()
    If cboSeatType.ListIndex = cboSeatType.ListCount - 1 Then
        cboSeatType.ListIndex = 0
    Else
        cboSeatType.ListIndex = cboSeatType.ListIndex + 1
    End If
End Sub


Private Sub SetBus()
    '显示车次补票
    cmdBus.ButtonType = cnActiveStatus
    cmdStation.ButtonType = cnNotActiveStatus
    cmdBus.Value = True
    
    ptExtraSellByStation.Visible = False
    ptExtraByBus.Visible = True
    lvStationPreSell.Visible = False
    lvBusPreSell.Visible = True
    MDISellTicket.EnableSortAndRefresh False
    
    lblStation.Visible = True
    lvStation.Visible = True
    lblBus.Visible = False
    lvBus.Visible = False
'    lvStation.Refresh
    DoThingWhenBusChange
    EnableSeatAndStand
    SetPreSellButton
    'EnableSellButton
    txtBusID.Text = ""
    txtReceivedMoney.Text = ""
    lvSellStation.ListItems.Clear
    If txtBusID.Visible Then txtBusID.SetFocus
End Sub

Private Sub SetStation()
    '显示站点补票
    cmdStation.ButtonType = cnActiveStatus
    cmdBus.ButtonType = cnNotActiveStatus
    cmdStation.Value = True
    
    ptExtraByBus.Visible = False
    ptExtraSellByStation.Visible = True
    lvStationPreSell.Visible = True
    lvBusPreSell.Visible = False
    MDISellTicket.EnableSortAndRefresh True
    lblStation.Visible = False
    lvStation.Visible = False
    lblBus.Visible = True
    lvBus.Visible = True
    '    lvBus.Refresh
    
    DoThingWhenBusChange
    EnableSeatAndStand
    SetPreSellButton
    'EnableSellButton
    txtReceivedMoney.Text = ""
    lvSellStation.ListItems.Clear
    If cboEndStation.Visible Then cboEndStation.SetFocus

End Sub

Private Function BusIsActive() As Boolean
    If cmdBus.ButtonType = cnActiveStatus Then
        BusIsActive = True
    Else
        BusIsActive = False
    End If
End Function


Private Function TotalInsurace() As Double
    '汇总保险费
    Dim i As Integer
    Dim nCount As Integer
    If chkInsurance.Value = vbChecked Then
        nCount = 0
        For i = 1 To lvStationPreSell.ListItems.count
            nCount = nCount + lvStationPreSell.ListItems(i).SubItems(EI_SumTicketNum)
        Next i
        For i = 1 To lvBusPreSell.ListItems.count
            nCount = nCount + lvBusPreSell.ListItems(i).SubItems(EI_SumTicketNum)
        Next i
        nCount = nCount + Val(txtFullSell.Text) + Val(txtHalfSell.Text) + Val(txtPreferentialSell.Text)
        '保险费设为每张2元
        
        TotalInsurace = nCount * 2
    Else
        TotalInsurace = 0
    End If
End Function







