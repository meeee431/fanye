VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{695ABF14-B2D8-11D2-A5ED-DE08DCF33612}#1.2#0"; "asfcombo.ocx"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{BBF95DAA-F9CB-4CA9-A673-E0E9E0193752}#1.0#0"; "STSellCtl.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmChangeSell 
   BackColor       =   &H00929292&
   Caption         =   "改签"
   ClientHeight    =   8130
   ClientLeft      =   750
   ClientTop       =   2280
   ClientWidth     =   12060
   HelpContextID   =   4000210
   Icon            =   "frmChangeSell.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8130
   ScaleWidth      =   12060
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1245
      HelpContextID   =   3000411
      Left            =   6060
      TabIndex        =   69
      Top             =   7710
      Visible         =   0   'False
      Width           =   2880
      Begin VB.CheckBox chkSetSeat 
         BackColor       =   &H00FFFFFF&
         Caption         =   "定座(&P)"
         Height          =   270
         HelpContextID   =   3000411
         Left            =   120
         TabIndex        =   70
         Top             =   -30
         Width           =   975
      End
      Begin VB.Label lblSetSeat 
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   180
         TabIndex        =   71
         Top             =   60
         Width           =   435
      End
   End
   Begin VB.Frame fraOutLine 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   7095
      Left            =   840
      TabIndex        =   18
      Top             =   300
      Width           =   11955
      Begin VB.ComboBox cboInsurance 
         Height          =   300
         Left            =   5010
         Style           =   2  'Dropdown List
         TabIndex        =   77
         Top             =   5490
         Width           =   1005
      End
      Begin VB.CommandButton cmdSetSeat 
         Caption         =   "定座(&G)"
         Enabled         =   0   'False
         Height          =   315
         HelpContextID   =   3000411
         Left            =   8160
         TabIndex        =   75
         Top             =   1110
         Width           =   870
      End
      Begin VB.TextBox txtSeat 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         HelpContextID   =   3000411
         Left            =   6675
         TabIndex        =   74
         Top             =   1125
         Width           =   1395
      End
      Begin RTComctl3.CoolButton cmdSell 
         Height          =   525
         Left            =   150
         TabIndex        =   13
         Top             =   6390
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   926
         BTYPE           =   3
         TX              =   "改签(&P)"
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
         MICON           =   "frmChangeSell.frx":0442
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
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
         Height          =   525
         Left            =   8670
         TabIndex        =   14
         Top             =   2250
         Visible         =   0   'False
         Width           =   2220
      End
      Begin FCmbo.asFlatCombo cboSeatType 
         Height          =   270
         Left            =   9810
         TabIndex        =   64
         Top             =   5070
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   476
         ButtonDisabledForeColor=   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
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
      Begin FText.asFlatSpinEdit txtPrevDate 
         Height          =   345
         Left            =   150
         TabIndex        =   3
         Top             =   1530
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonHotBackColor=   -2147483633
         ButtonPressedBackColor=   -2147483627
         Text            =   "0"
         ButtonBackColor =   -2147483633
         Registered      =   -1  'True
         OfficeXPColors  =   -1  'True
      End
      Begin MSComctlLib.ListView lvPreSell 
         Height          =   1185
         Left            =   6150
         TabIndex        =   59
         Top             =   5790
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   2090
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
      Begin VB.Frame fraTicketTypeChange 
         BackColor       =   &H00E0E0E0&
         Caption         =   "票种"
         Height          =   645
         HelpContextID   =   3000419
         Left            =   2910
         TabIndex        =   19
         Top             =   4830
         Width           =   5445
         Begin FCmbo.asFlatCombo cboPreferentialTicket 
            Height          =   270
            Left            =   3840
            TabIndex        =   68
            Top             =   240
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   476
            ButtonDisabledForeColor=   12632256
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
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
         Begin VB.OptionButton optFullTicket 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "全票(&A)"
            ForeColor       =   &H80000008&
            Height          =   255
            HelpContextID   =   3000419
            Left            =   150
            TabIndex        =   8
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optHalfTicket 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "半票(&Q)"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   285
            HelpContextID   =   3000419
            Left            =   1320
            TabIndex        =   15
            Top             =   240
            Width           =   1065
         End
         Begin VB.OptionButton optPreferentialTicket 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "优惠票(&X)"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   285
            HelpContextID   =   3000419
            Left            =   2550
            TabIndex        =   16
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.TextBox txtReceivedMoney 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1050
         TabIndex        =   10
         Text            =   "0.00"
         Top             =   5160
         Width           =   1695
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   3090
         Top             =   1770
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
               Picture         =   "frmChangeSell.frx":045E
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txtOldTktNum 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         HelpContextID   =   3000415
         Left            =   180
         MaxLength       =   10
         TabIndex        =   1
         Top             =   390
         Width           =   2415
      End
      Begin VB.TextBox txtSheetID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   12
         Top             =   6000
         Width           =   1590
      End
      Begin STSellCtl.ucSuperCombo cboEndStation 
         Height          =   1155
         Left            =   150
         TabIndex        =   5
         Top             =   2190
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   2037
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
      Begin RTComctl3.FlatLabel flblSellDate 
         Height          =   375
         Left            =   960
         TabIndex        =   30
         Top             =   1530
         Width           =   1875
         _ExtentX        =   3307
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
         BackColor       =   14737632
         OutnerStyle     =   2
         Caption         =   ""
      End
      Begin MSComctlLib.ListView lvBus 
         Height          =   3285
         Left            =   2910
         TabIndex        =   7
         Top             =   1470
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   5794
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
         ColHdrIcons     =   "ImageList1"
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
      Begin RTComctl3.FlatLabel flblStandCount 
         Height          =   375
         Left            =   8685
         TabIndex        =   31
         Top             =   3210
         Width           =   780
         _ExtentX        =   1376
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
      Begin RTComctl3.FlatLabel flblLimitedCount 
         Height          =   375
         Left            =   3240
         TabIndex        =   32
         Top             =   3210
         Width           =   1320
         _ExtentX        =   2328
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
      Begin RTComctl3.FlatLabel flblLimitedTime 
         Height          =   375
         Left            =   5490
         TabIndex        =   33
         Top             =   3210
         Width           =   2670
         _ExtentX        =   4710
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
      Begin VB.CommandButton cmdForFocusSeat 
         Caption         =   "Seat"
         Height          =   315
         Left            =   1800
         TabIndex        =   17
         Top             =   2580
         Width           =   615
      End
      Begin VB.Frame fraDiscountTicket 
         Caption         =   "折扣票:"
         Height          =   915
         Left            =   1650
         TabIndex        =   58
         Top             =   2280
         Width           =   1005
         Begin VB.TextBox txtDiscount 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   90
            TabIndex        =   60
            Text            =   "1"
            Top             =   540
            Width           =   795
         End
         Begin VB.Label Label18 
            Caption         =   "折扣(&F):"
            Height          =   255
            Left            =   120
            TabIndex        =   61
            Top             =   240
            Width           =   765
         End
      End
      Begin MSComctlLib.ListView lvSellStation 
         Height          =   1185
         Left            =   2970
         TabIndex        =   72
         Top             =   5790
         Width           =   3030
         _ExtentX        =   5345
         _ExtentY        =   2090
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
            Text            =   "票价"
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
      Begin VB.Label lblInsurance 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "保险(F12):"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   4050
         TabIndex        =   78
         Top             =   5520
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "座位号(&T):"
         Height          =   180
         Left            =   5640
         TabIndex        =   76
         Top             =   1230
         Width           =   900
      End
      Begin VB.Label lblSellStation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "上车站(&O):"
         Height          =   180
         Left            =   2970
         TabIndex        =   73
         Top             =   5520
         Width           =   900
      End
      Begin VB.Label lblCredence 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "凭证号(&C):"
         Height          =   180
         Left            =   270
         TabIndex        =   11
         Top             =   6060
         Width           =   900
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
         Left            =   2100
         TabIndex        =   67
         Top             =   5640
         Width           =   585
      End
      Begin VB.Label lblOffTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "10:00"
         Height          =   180
         Left            =   7485
         TabIndex        =   66
         Top             =   585
         Width           =   450
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "发车时间:"
         Height          =   180
         Left            =   6480
         TabIndex        =   65
         Top             =   585
         Width           =   810
      End
      Begin VB.Label lblSeatType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "座位类型(&R):"
         Height          =   180
         Left            =   8580
         TabIndex        =   63
         Top             =   5130
         Width           =   1080
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "当前票改签信息:"
         Height          =   180
         Left            =   180
         TabIndex        =   62
         Top             =   3420
         Width           =   1350
      End
      Begin VB.Label lblPreBus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "预售车次列表(&E):"
         Height          =   180
         Left            =   6150
         TabIndex        =   57
         Top             =   5520
         Width           =   1440
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "票信息:"
         Height          =   180
         Left            =   2760
         TabIndex        =   56
         Top             =   60
         Width           =   630
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00808080&
         X1              =   2730
         X2              =   11700
         Y1              =   270
         Y2              =   270
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "日期:"
         Height          =   180
         Left            =   6480
         TabIndex        =   55
         Top             =   330
         Width           =   450
      End
      Begin VB.Label lblBusID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "21050"
         Height          =   180
         Left            =   5265
         TabIndex        =   54
         Top             =   330
         Width           =   450
      End
      Begin VB.Label lblEndStation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "杭州"
         Height          =   180
         Left            =   3480
         TabIndex        =   53
         Top             =   585
         Width           =   360
      End
      Begin VB.Label lblStartStation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "宁波南站"
         Height          =   180
         Left            =   3480
         TabIndex        =   52
         Top             =   330
         Width           =   720
      End
      Begin VB.Label lblOperatorChange 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "售票员:"
         Height          =   180
         Left            =   9090
         TabIndex        =   51
         Top             =   585
         Width           =   630
      End
      Begin VB.Label lblStateChange 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "状态:"
         Height          =   180
         Left            =   9090
         TabIndex        =   50
         Top             =   330
         Width           =   450
      End
      Begin VB.Label label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "到站:"
         Height          =   180
         Left            =   2850
         TabIndex        =   49
         Top             =   585
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "起站:"
         Height          =   180
         Index           =   1
         Left            =   2850
         TabIndex        =   48
         Top             =   330
         Width           =   450
      End
      Begin VB.Label lblTypeChange 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "票种:"
         Height          =   180
         Left            =   4785
         TabIndex        =   47
         Top             =   585
         Width           =   450
      End
      Begin VB.Label lblScheduleChange 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车次:"
         Height          =   180
         Left            =   4785
         TabIndex        =   46
         Top             =   330
         Width           =   450
      End
      Begin VB.Label lblTimeChange 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "售票时间:"
         Height          =   180
         Left            =   6480
         TabIndex        =   45
         Top             =   825
         Width           =   810
      End
      Begin VB.Label lblSeller 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "张三"
         Height          =   180
         Left            =   9795
         TabIndex        =   44
         Top             =   585
         Width           =   360
      End
      Begin VB.Label lblTicketType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "全票"
         Height          =   180
         Left            =   5265
         TabIndex        =   43
         Top             =   585
         Width           =   360
      End
      Begin VB.Label lblSellTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2002-07-14 10:15"
         Height          =   180
         Left            =   7485
         TabIndex        =   42
         Top             =   825
         Width           =   1440
      End
      Begin VB.Label lblBusDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2002-07-15"
         Height          =   180
         Left            =   7485
         TabIndex        =   41
         Top             =   330
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "票价:"
         Height          =   180
         Left            =   2850
         TabIndex        =   40
         Top             =   825
         Width           =   450
      End
      Begin VB.Label lblPrice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "37.5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3480
         TabIndex        =   39
         Top             =   750
         Width           =   570
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "正常售出"
         Height          =   180
         Left            =   9795
         TabIndex        =   38
         Top             =   330
         Width           =   720
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "限售时间:"
         Height          =   180
         Left            =   4650
         TabIndex        =   37
         Top             =   3270
         Width           =   765
      End
      Begin VB.Label lblChangeMsg 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   9300
         TabIndex        =   36
         Top             =   5460
         Width           =   165
      End
      Begin VB.Label Label1 
         Caption         =   "限售张数:"
         Height          =   225
         Index           =   0
         Left            =   3000
         TabIndex        =   35
         Top             =   3270
         Width           =   855
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "站票:"
         Height          =   180
         Left            =   8205
         TabIndex        =   34
         Top             =   3330
         Width           =   450
      End
      Begin VB.Label lblEndChange 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "到站(&Z):"
         Height          =   180
         Left            =   150
         TabIndex        =   4
         Top             =   1950
         Width           =   720
      End
      Begin VB.Label lblDaysChange 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "预售天数(&D):"
         Height          =   180
         Left            =   180
         TabIndex        =   2
         Top             =   1260
         Width           =   1080
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "新车票:"
         Height          =   180
         Left            =   165
         TabIndex        =   29
         Top             =   855
         Width           =   630
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   4
         X1              =   150
         X2              =   11700
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label lblOldTktNum 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "改签原票号(&E):"
         Height          =   180
         Left            =   210
         TabIndex        =   0
         Top             =   120
         Width           =   1260
      End
      Begin VB.Label lblBusList 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车次列表(&B):"
         Height          =   180
         Left            =   2910
         TabIndex        =   6
         Top             =   1260
         Width           =   1080
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "原票价:"
         Height          =   180
         Left            =   270
         TabIndex        =   28
         Top             =   3780
         Width           =   630
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "手续费:"
         Height          =   180
         Left            =   270
         TabIndex        =   27
         Top             =   4080
         Width           =   630
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "改签票票价"
         Height          =   180
         Left            =   270
         TabIndex        =   26
         Top             =   4410
         Width           =   900
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         X1              =   180
         X2              =   2820
         Y1              =   3690
         Y2              =   3690
      End
      Begin VB.Label lblOldTicketPrice 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2100
         TabIndex        =   25
         Top             =   3720
         Width           =   570
      End
      Begin VB.Label lblChangeCharge 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2100
         TabIndex        =   24
         Top             =   4020
         Width           =   570
      End
      Begin VB.Label lblTotalPrice 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2100
         TabIndex        =   23
         Top             =   4350
         Width           =   570
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         X1              =   180
         X2              =   2790
         Y1              =   4740
         Y2              =   4740
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "应收:"
         Height          =   180
         Left            =   270
         TabIndex        =   22
         Top             =   4830
         Width           =   450
      End
      Begin VB.Label lblFactIn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "实收(&M):"
         Height          =   180
         Left            =   270
         TabIndex        =   9
         Top             =   5220
         Width           =   720
      End
      Begin VB.Label lblRightMoney 
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
         Left            =   2100
         TabIndex        =   21
         Top             =   4770
         Width           =   585
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00808080&
         X1              =   180
         X2              =   2820
         Y1              =   5640
         Y2              =   5640
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "应找:"
         Height          =   180
         Left            =   270
         TabIndex        =   20
         Top             =   5700
         Width           =   450
      End
   End
End
Attribute VB_Name = "frmChangeSell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'改签用的枚举
Private Enum ChangeTicketInfo
    CC_BusType = 1
    CC_BusDate = 2
    CC_EndStation = 3
    CC_OffTime = 4
    CC_VehicleModel = 5
    CC_FullPrice = 6
    CC_FullNum = 7
    CC_HalfPrice = 8
    CC_HalfNum = 9
    CC_FreePrice = 10
    CC_FreeNum = 11
    CC_PreferentialPrice1 = 12
    CC_PreferentialNum1 = 13
    CC_PreferentialPrice2 = 14
    CC_PreferentialNum2 = 15
    CC_PreferentialPrice3 = 16
    CC_PreferentialNum3 = 17
    CC_EndStationCode = 18
    CC_OriginalTicket = 19
    CC_ChangeSheet = 20
    CC_CheckGate = 21
    CC_Discount = 22
    CC_ChangeFees = 23
    CC_OriginalPrice = 24
    CC_SeatStatus1 = 25
    CC_SeatStatus2 = 26
    CC_SeatNo = 27
    CC_AllTicketPrice = 28
    CC_AllTicketType = 29
    CC_TotalNum = 30
    CC_SeatType = 31
    CC_TerminateName = 32
End Enum



Private blPointCount As Boolean
Private m_bTicketInfoDirty As Boolean

Private m_atbSeatTypeBus As TMultiSeatTypeBus '得到不同座位类型的车次
Private m_TicketPrice() As Single
Private m_TicketTypeDetail() As ETicketType

Dim m_aszSeatType() As String

Dim m_atTicketType() As TTicketType
Private rsCountTemp As Recordset
Private m_aszInsurce() As String '乘意险

Public m_vaCardInfo As Variant     '实名信息
Private m_aszCardInfo() As TCardInfo    '实名信息

Private Sub cboEndStation_Change()
On Error GoTo here
    If lvBus.ListItems.count > 0 Then
       DoThingWhenBusChange
    Else
       cboEndStation.SetFocus
    End If
    DealPrice
    cmdPreSell.Enabled = True
'On Error GoTo 0
Exit Sub
here:
  ShowErrorMsg
End Sub
Private Sub cboEndStation_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyLeft
            KeyCode = 0
            If Val(txtPrevDate.Text) > 0 Then
                txtPrevDate.Text = Val(txtPrevDate.Text) - 1
            End If
        Case vbKeyRight
            KeyCode = 0
            If Val(txtPrevDate.Text) < m_nCanSellDay Then
            
                txtPrevDate.Text = Val(txtPrevDate.Text) + 1
            End If
    End Select
'    If m_bPreClear Then
'        lvPreSell.ListItems.Clear
'        flblTotalPrice.Caption = 0#
'        txtReceivedMoney.Text = ""
'        flblRestMoney.Caption = ""
'        m_bPreClear = False
'    End If
   
End Sub

Private Sub cboEndStation_GotFocus()
    lblEndChange.ForeColor = clActiveColor
End Sub

Private Sub cboPreferentialTicket_Change()
    DealPrice
End Sub

Private Sub cboPreferentialTicket_Click()
    DealPrice
End Sub

Private Sub cmdForFocusMoney_GotFocus()
    txtOldTktNum.SetFocus
End Sub

Private Sub cboPreferentialTicket_GotFocus()
    optPreferentialTicket_GotFocus
End Sub

Private Sub cboPreferentialTicket_LostFocus()
    optPreferentialTicket_LostFocus
End Sub

Private Sub cboSeatType_Change()
 If lvSellStation.ListItems.count > 0 Then
  RefreshBusStation rsCountTemp, Trim(lvSellStation.SelectedItem.SubItems(3)), cboSeatType.ListIndex + 1
 End If
End Sub

Private Sub cboSeatType_GotFocus()
    lblSeatType.ForeColor = clActiveColor
End Sub

Private Sub cboSeatType_LostFocus()
    lblSeatType.ForeColor = 0
End Sub

Private Sub cboInsurance_Click()
    '改签,暂不累加保险费.
    '因为要加的话,还要考虑到原来的票是否有保险
End Sub

Private Sub cmdForFocusSeat_GotFocus()
    If cmdSell.Enabled Then
        cmdSell.SetFocus
    Else
        txtOldTktNum.SetFocus
    End If
End Sub

Private Sub cmdPreSell_Click()
    Dim nSameIndex As Integer
    GetPreSellInfo
    cmdPreSell.Enabled = False
    txtSeat.Text = ""
    
End Sub


Private Sub cmdPreSell_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdPreSell_Click
    End If
End Sub

Private Sub Form_Activate()
On Error GoTo here
    m_nCurrentTask = RT_ChangeTicket
    m_szCurrentUnitID = Me.Tag
    SetPreSellInfo
    MDISellTicket.SetFunAndUnit
    lvBus.SortKey = MDISellTicket.GetSortKey()
'------------------------------------
    MDISellTicket.EnableSortAndRefresh True
    
    'MDISellTicket.abMenuTool.Bands("mnuFunction").Tools("mnuChangeSeatType").Enabled = True
    
    Exit Sub
here:
    ShowErrorMsg
End Sub

Private Sub Form_Deactivate()
    MDISellTicket.EnableSortAndRefresh False
    'MDISellTicket.abMenuTool.Bands("mnuFunction").Tools("mnuChangeSeatType").Enabled = False
    
End Sub

Private Sub Form_Resize()
    If MDISellTicket.ActiveForm Is Me Then
        If Not Me.WindowState = vbMaximized Then Me.WindowState = vbMaximized
        fraOutLine.Left = (Me.ScaleWidth - fraOutLine.Width) / 2
        fraOutLine.Top = (Me.ScaleHeight - fraOutLine.Height) / 2
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    m_clChange.Remove GetEncodedKey(Me.Tag)
    MDISellTicket.abMenuTool.Bands("mnuFunction").Tools("mnuChangeTkt").Checked = False
'    MDISellTicket.mnuChangeTkt.Checked = False
    
    
    MDISellTicket.EnableSortAndRefresh False
End Sub

'------------------------------------------------------

Private Sub cboEndStation_LostFocus()
    lblEndChange.ForeColor = 0
    RefreshBus True
    DealPrice
    DealPreButtonStatus
End Sub

Private Sub chkSetSeat_Click()
'    If chkSetSeat.Value = vbChecked And chkSetSeat.Enabled Then
'        txtSeat.Enabled = True
'        cmdSetSeat.Enabled = True
'        txtSeat.SetFocus
'    Else
'        txtSeat.Enabled = False
'        cmdSetSeat.Enabled = False
'        txtSeat.Text = ""
'      End If
End Sub


'改签售票
Private Sub cmdSell_Click()

    
    Dim k As Long
    Dim m As Long
    Dim i As Integer
    m = 0
    For i = 1 To lvPreSell.ListItems.count
        m = m + 1 'lvPreSell.ListItems(i).SubItems(IT_SumTicketNum)
    Next i
    If m_lEndTicketNoOld = 0 Then
        ShowMsg "售票不成功，用户还未领票，请先去领票！"
        Exit Sub
    End If
    If m + Val(m_lTicketNo) - 1 > Val(m_lEndTicketNo) Then
        k = Val(m_lEndTicketNo) - Val(m_lTicketNo) + 1
        MsgBox "打印机上的票已不够！" & vbCrLf & "车票只剩 " & k & "张", vbInformation, "售票台"
    Else
        ChangeSell
    End If
    
    
End Sub

Private Sub cmdSetSeat_Click()
    Dim rsTemp As Recordset
    If lvBus.SelectedItem Is Nothing Then Exit Sub
    Set rsTemp = m_oSell.GetSeatRs(CDate(flblSellDate.Caption), lvBus.SelectedItem.Text)
    Set frmOrderSeats.m_rsSeat = rsTemp
    frmOrderSeats.m_szSeatNumber = PreOrderSeat
    frmOrderSeats.Show vbModal
    If frmOrderSeats.m_bOk Then
        txtSeat = frmOrderSeats.m_szSeat
    End If
    Set rsTemp = Nothing
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   On Error GoTo here
   If KeyAscii = 45 Then
         If lvBus.ListItems.count > 0 Then
            If lvBus.SelectedItem.SubItems(ID_OffTime) = cszScrollBus Then
                Exit Sub
            End If
        End If
        KeyAscii = 0
        ChangeSeatType
        DealPrice
        Exit Sub
    End If
    If KeyAscii = 27 Then
'            lvPreSell.ListItems.Clear
        lvPreSell.ListItems.Clear
        txtPrevDate.Text = 0
        cboEndStation.Text = ""
        cboEndStation.SetFocus
        txtSeat.Text = ""
        ElseIf KeyAscii = 13 And (lvSellStation.Enabled) And (Me.ActiveControl Is lvBus) Then
        lvSellStation.SetFocus
        Exit Sub
        ElseIf KeyAscii = 13 And (lvSellStation.Enabled) And (Me.ActiveControl Is lvSellStation) Then
        txtReceivedMoney.SetFocus
        Exit Sub
        ElseIf KeyAscii = 13 And (lvSellStation.Enabled) And (Me.ActiveControl Is txtReceivedMoney) Then
        txtSheetID.SetFocus
        Exit Sub
        cboInsurance.ListIndex = 0
    ElseIf KeyAscii = vbKeyReturn And (Not Me.ActiveControl Is cboEndStation) _
            And (Not Me.ActiveControl Is optHalfTicket) And (Not Me.ActiveControl Is optPreferentialTicket) _
            And (Not Me.ActiveControl Is txtReceivedMoney) Then
        SendKeys "{TAB}"
    End If
Exit Sub
here:
    ShowErrorMsg
End Sub

Private Sub Form_Load()
On Error GoTo here
    SetDefaultLabel
    txtOldTktNum.MaxLength = 10
'    lblUser.Caption = m_oAUser.UserID & "/" & m_oAUser.UserName
    m_bTicketInfoDirty = True
    flblSellDate.Caption = ToStandardDateStr(m_oParam.NowDate)
    txtPrevDate.Value = 0
    
    RefreshPreferentialTicket
    DealDiscountAndSeat
    DealPreButtonStatus
    RefreshStation2
    SetDefaultSellTicket
    EnableSeatAndStand
    EnableSellButton
    
    If m_bSellStationCanSellEachOther Then
      lvSellStation.Enabled = True
      Else
      lvSellStation.Enabled = False
    End If
    
    '初始化保险信息列表 zyw 2008-01-07
    cboInsurance.AddItem "无保险"
    cboInsurance.AddItem "1元"
    cboInsurance.AddItem "2元"
    cboInsurance.ListIndex = 0
    
    m_atbSeatTypeBus = m_oSell.GetMultiSeatTypeBus
Exit Sub
here:
    ShowErrorMsg
End Sub


'Private Sub ISortKeyChanged_RefreshBus()
'    RefreshBus True
'End Sub




Private Sub lvBus_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    SortListView lvBus, ColumnHeader.Index
End Sub


Private Sub lvBus_GotFocus()
    lblBusList.ForeColor = clActiveColor
    ShowRightSeatType
End Sub

Private Sub lvBus_ItemClick(ByVal Item As MSComctlLib.ListItem)
  If lvBus.ListItems.count > 0 Then
    RefreshSellStation rsCountTemp
  End If
    DoThingWhenBusChange
End Sub
Private Sub DoThingWhenBusChange()
    On Error GoTo here
    If Not lvBus.SelectedItem Is Nothing Then
        Dim liTemp As ListItem
        Set liTemp = lvBus.SelectedItem
        If liTemp.SubItems(ID_BusType) = TP_ScrollBus Then
            
            flblLimitedCount.Caption = ""
            flblLimitedTime.Caption = ""
            flblStandCount.Caption = ""
        Else
            flblLimitedCount.Caption = GetStationLimitedCountStr(CInt(liTemp.SubItems(ID_LimitedCount)))
            flblLimitedTime.Caption = GetStationLimitedTimeStr(CInt(liTemp.SubItems(ID_LimitedTime)), CDate(flblSellDate.Caption), CDate(liTemp.SubItems(ID_OffTime)))
            'flblStandCount.Caption = liTemp.subitems(ID_StandCount)
        End If
    Else

        flblLimitedCount.Caption = ""
        flblLimitedTime.Caption = ""
        flblStandCount.Caption = ""

    End If
    DealPrice
    EnableSeatAndStand
    EnableSellButton
    On Error GoTo 0
    Exit Sub
here:
    ShowErrorMsg
End Sub

Private Sub lvBus_LostFocus()
   lblBusList.ForeColor = 0
End Sub

Private Sub lvPreSell_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    SortListView lvPreSell, ColumnHeader.Index
End Sub

Private Sub lvPreSell_GotFocus()
    lblPreBus.ForeColor = clActiveColor
End Sub

Private Sub lvPreSell_KeyPress(KeyAscii As Integer)
    If Not lvPreSell.SelectedItem Is Nothing Then
        If KeyAscii = 8 Then
            lvPreSell.ListItems.Remove lvPreSell.SelectedItem.Index
        End If
    End If
End Sub

Private Sub lvPreSell_LostFocus()
    lblPreBus.ForeColor = 0
End Sub

Private Sub lvSellStation_ItemClick(ByVal Item As MSComctlLib.ListItem)
    cboSeatType_Change
    lblTotalPrice.Caption = FormatMoney(lvSellStation.SelectedItem.SubItems(2))
    DealPrice
End Sub

Private Sub lvSellStation_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdPreSell_Click
        txtReceivedMoney.SetFocus
    End If
End Sub

Private Sub optFullTicket_GotFocus()
    optFullTicket.ForeColor = clActiveColor
End Sub

Private Sub optFullTicket_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
 
         txtReceivedMoney.SetFocus
    End If
End Sub

Private Sub optFullTicket_LostFocus()
    optFullTicket.ForeColor = 0
End Sub

Private Sub optHalfTicket_GotFocus()
    optHalfTicket.ForeColor = clActiveColor
End Sub

Private Sub optHalfTicket_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
         txtReceivedMoney.SetFocus
    End If
End Sub

Private Sub optHalfTicket_LostFocus()
    optHalfTicket.ForeColor = 0
End Sub

Private Sub optPreferentialTicket_Click()
    EnableSellButton
    'EnableSeatAndStand
    DealPrice
    DealPreButtonStatus
End Sub

Private Sub optFullTicket_Click()
    EnableSellButton
    'EnableSeatAndStand
    DealPrice
    DealPreButtonStatus
End Sub

Private Sub optHalfTicket_Click()
    EnableSellButton
    'EnableSeatAndStand
    DealPrice
    DealPreButtonStatus
End Sub


Private Sub optPreferentialTicket_GotFocus()
    optPreferentialTicket.ForeColor = clActiveColor
    
End Sub


Private Sub optPreferentialTicket_KeyPress(KeyAscii As Integer)
     If KeyAscii = vbKeyReturn Then
         txtReceivedMoney.SetFocus
    End If
End Sub

Private Sub optPreferentialTicket_LostFocus()
    optPreferentialTicket.ForeColor = 0
End Sub

Private Sub txtDiscount_Change()
    DealPreButtonStatus
End Sub

Private Sub txtDiscount_KeyPress(KeyAscii As Integer)
    If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = 8 Or (KeyAscii = 46 And InStr(txtDiscount.Text, ".") = 0) Then
    Else
        KeyAscii = 0
    End If
End Sub
Private Sub txtDiscount_LostFocus()
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

Private Sub txtOldTktNum_Change()
    m_bTicketInfoDirty = True
    txtReceivedMoney.Text = ""
    DealPreButtonStatus
End Sub

Private Sub txtOldTktNum_GotFocus()
    lblOldTktNum.ForeColor = clActiveColor
End Sub

Private Sub txtOldTktNum_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim szTemp As String
    Dim lTemp As Long
    On Error GoTo Error_Handel
    If (Len(txtOldTktNum.Text) >= TicketNoNumLen()) Then
        If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
            lTemp = Right(txtOldTktNum.Text, TicketNoNumLen())
            szTemp = Left(txtOldTktNum.Text, Len(txtOldTktNum.Text) - TicketNoNumLen())
            
            If KeyCode = vbKeyUp Then
                lTemp = lTemp + 1
            Else
                lTemp = lTemp - 1
            End If
            txtOldTktNum.Text = MakeTicketNo(lTemp, szTemp)
            KeyCode = 0
        End If
    End If
    Exit Sub
Error_Handel:
End Sub

Private Sub txtOldTktNum_LostFocus()
    lblOldTktNum.ForeColor = 0
    ShowOldTicketInfo
    DealPrice
End Sub

Private Sub txtPrevDate_Change()
    On Error Resume Next
    
    If txtPrevDate.Value > m_nCanSellDay Then txtPrevDate.Value = m_nCanSellDay
    flblSellDate.Caption = ToStandardDateStr(DateAdd("d", txtPrevDate.Value, m_oParam.NowDate))
End Sub

Private Sub DealPrice()
    Dim sgTemp As Double
    Dim sgvalue As Double
    Dim aszSeatNo() As String
    Dim nSeat As Integer
    Dim nSameIndex As Integer
    Dim liTemp1 As ListItem
    Dim nLength As Integer
    
    Dim sgSum As Single
    sgSum = 0
    sgTemp = 0
    
    On Error GoTo here
    '计算票价
    If Not lvBus.SelectedItem Is Nothing Then
        Dim liTemp As ListItem
        Set liTemp = lvBus.SelectedItem
        Select Case Trim(m_aszSeatType(cboSeatType.ListIndex + 1, 1))
            Case cszSeatType
                If optFullTicket.Value = True Then
                    sgTemp = liTemp.SubItems(ID_FullPrice)
                Else
                    If optHalfTicket.Value = True Then
                        sgTemp = liTemp.SubItems(ID_HalfPrice)
                    Else
                        sgTemp = GetPreferentialPrice
                    End If
                End If
                
            Case cszBedType
                If optFullTicket.Value = True Then
                    sgTemp = liTemp.SubItems(ID_BedFullPrice)
                Else
                    If optHalfTicket.Value = True Then
                        sgTemp = liTemp.SubItems(ID_BedHalfPrice)
                    Else
                        sgTemp = GetPreferentialPrice
                    End If
                End If
                
            Case cszAdditionalType
                If optFullTicket.Value = True Then
                    sgTemp = liTemp.SubItems(ID_AdditionalFullPrice)
                Else
                    If optHalfTicket.Value = True Then
                        sgTemp = liTemp.SubItems(ID_AdditionalHalfPrice)
                    Else
                        sgTemp = GetPreferentialPrice
                    End If
                End If
               
        End Select
        Set liTemp = Nothing
        lblTotalPrice.Caption = FormatMoney(sgTemp)
        'lblRightMoney.Caption = FormatMoney(CDbl(lblTotalPrice.Caption) + CDbl(lblChangeCharge.Caption) - CDbl(lblOldTicketPrice.Caption))
    Else
        'lblOldTicketPrice.Caption = FormatMoney(0)
        lblChangeCharge.Caption = FormatMoney(0)
        lblTotalPrice.Caption = FormatMoney(0)
    End If
    If txtReceivedMoney.Text = "0" And Not Me.ActiveControl Is txtReceivedMoney Then txtReceivedMoney.Text = ""
    If Left(txtReceivedMoney.Text, 1) = "." Then txtReceivedMoney.Text = "0" & txtReceivedMoney.Text
    
    If txtReceivedMoney.Text = "" Then
       sgvalue = 0
    Else
       sgvalue = CDbl(txtReceivedMoney.Text)
    End If

    
    '应收票款=改签票价-原票价+手续费 +保险费
    GetRightMoney
    flblRestMoney.Caption = FormatMoney(sgvalue - CDbl(lblRightMoney))
    '如果应收票款>=0,且原票价>=0,
'    If CDbl(flblRestMoney.Caption) >= 0 And CDbl(lblOldTicketPrice.Caption) > 0 Then   'CDbl(lblRightMoney.Caption) >= 0 And
'       cmdSell.Enabled = True
'    Else
'       cmdSell.Enabled = False
'    End If
Exit Sub
here:
    ShowErrorMsg
End Sub


Private Sub txtPrevDate_GotFocus()
    lblDaysChange.ForeColor = clActiveColor

End Sub

Private Sub txtPrevDate_LostFocus()
    lblDaysChange.ForeColor = 0
    On Error GoTo here
'    lblPrevDate.ForeColor = 0
    If txtPrevDate.Text = "" Then txtPrevDate.Text = 0
    RefreshBus True
On Error GoTo 0
Exit Sub
here:
    ShowErrorMsg
End Sub

Private Sub txtReceivedMoney_Change()
    DealPrice
    DealPreButtonStatus
End Sub

Private Sub EnableSellButton()
    Dim szStationID As String
    Dim sgvalue As Double
    
    szStationID = cboEndStation.BoundText
    If txtReceivedMoney.Text = "" Then sgvalue = 0
    flblRestMoney.Caption = FormatMoney(sgvalue - CDbl(lblRightMoney))
    
    If szStationID = "" Or lvBus.SelectedItem Is Nothing Then
        cmdSell.Enabled = False
    Else
        cmdSell.Enabled = True
    End If

End Sub

Private Sub EnableSeatAndStand()
    Dim szStationID As String
    szStationID = cboEndStation.BoundText
    If szStationID = "" Or lvBus.SelectedItem Is Nothing Then '当前无车次
        cmdSetSeat.Enabled = False
'        chkSetSeat.Value = 0
'        chkSetSeat.Enabled = False
    Else
        Dim liTemp As ListItem
        Set liTemp = lvBus.SelectedItem
        
        If liTemp.SubItems(ID_BusType) = TP_ScrollBus Then '是流水车次的话定座和站票无意义
            cmdSetSeat.Enabled = False
'            chkStandSeat.Value = 0
'            chkStandSeat.Enabled = False
        Else
            If liTemp.SubItems(ID_SeatCount) > 0 Then '可售座位数大于0

                cmdSetSeat.Enabled = True
'                chkSetSeat.Enabled = True

            Else '无可售座位数（则这时可售站票肯定大于0，不然就不会将此车次查出来）
                cmdSetSeat.Enabled = False
'                chkSetSeat.Enabled = False
'                chkSetSeat.Value = 0

            End If
            
        End If
    End If

End Sub

Private Sub SetDefaultSellTicket()
    optFullTicket.Value = True
    
    If chkSetSeat.Enabled Then
        chkSetSeat.Value = 0
    End If
   
    If txtReceivedMoney.Enabled Then
        txtReceivedMoney.Text = 0
        DealPrice
    End If
End Sub

Private Sub lvSellStation_GotFocus()
   lblSellStation.ForeColor = clActiveColor
   cboSeatType_Change
   If lvSellStation.ListItems.count > 0 Then
   lblTotalPrice.Caption = FormatMoney(lvSellStation.SelectedItem.SubItems(2))
   End If
   DealPrice
End Sub

Private Sub lvSellStation_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo here
   If KeyCode = 13 Then
      txtReceivedMoney.SetFocus
   End If
Exit Sub
here:
    ShowErrorMsg
  
End Sub

Private Sub lvSellStation_LostFocus()
    lblSellStation.ForeColor = 0
End Sub

'座位类型改变时 , 刷新相应的票价
Private Sub RefreshBusStation(rsTemp As Recordset, SellStationID As String, SeatTypeID As String)
  Dim i As Integer
  On Error GoTo err:
     If rsTemp.RecordCount = 0 Then
        Exit Sub
     End If
    rsTemp.MoveFirst
    For i = 1 To rsTemp.RecordCount
      If Trim(rsTemp!bus_id) = Trim(lvBus.SelectedItem.Text) Then
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
        lvBus.SelectedItem.SubItems(ID_CheckGate) = Trim(rsTemp!check_gate_id)
        End If
      End If
     End If
        rsTemp.MoveNext
    Next i
        lvBus.SelectedItem.Tag = MakeDisplayString(lvSellStation.SelectedItem.SubItems(3), lvSellStation.SelectedItem.Text)
        lvBus.SelectedItem.SubItems(ID_OffTime) = lvSellStation.SelectedItem.SubItems(1)
        lvBus.SelectedItem.SubItems(ID_FullPrice) = FormatMoney(lvSellStation.SelectedItem.SubItems(2))
        lvBus.SelectedItem.SubItems(ID_CheckGate) = lvSellStation.SelectedItem.SubItems(4)
    Exit Sub
err:
   MsgBox err.Description
End Sub

'刷新某车次的上车站信息

Private Sub RefreshSellStation(rsTemp As Recordset)
  Dim i As Integer
  Dim lvS As ListItem
  Dim szTemp As String
  Dim nBusType As EBusType
  On Error GoTo err:
    lvSellStation.Sorted = False
    lvSellStation.ListItems.Clear
    lvSellStation.Refresh
    szTemp = ""
     If rsTemp.RecordCount = 0 Then
        Exit Sub
     End If
    rsTemp.MoveFirst
    For i = 1 To rsTemp.RecordCount
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

Private Sub RefreshBus(Optional pbForce As Boolean = False, Optional pszSellStationID As String = "")
    Dim szStationID As String
    Dim liTemp As ListItem
    Dim lForeColor As OLE_COLOR
    Dim nBusType As EBusType
    Dim i As Integer
'    Dim varBookmark As Variant

   On Error GoTo here
    szStationID = RTrim(cboEndStation.BoundText)
'        If pszSellStationID = "" Then
'         lvSellStation.ListItems.Clear
'        End If
    If cboEndStation.Changed Or pbForce Then
        
        If szStationID <> "" Then
            Set rsCountTemp = m_oSell.GetBusRs(CDate(flblSellDate.Caption), szStationID)
            
            lvBus.ListItems.Clear
            Do While Not rsCountTemp.EOF
            
                For i = lvBus.ListItems.count To 1 Step -1
                    
                    If RTrim(rsCountTemp!bus_id) = lvBus.ListItems(i) And Format(rsCountTemp!bus_date, "yyyy-mm-dd") = CDate(flblSellDate.Caption) Then
'                        Select Case Trim(rsCountTemp!seat_type_id)
'                            Case cszSeatType
'                                liTemp.SubItems(ID_FullPrice) = rsCountTemp!full_price
'                                liTemp.SubItems(ID_HalfPrice) = rsCountTemp!half_price
'                                liTemp.SubItems(ID_PreferentialPrice1) = rsCountTemp!preferential_ticket1
'                                liTemp.SubItems(ID_PreferentialPrice2) = rsCountTemp!preferential_ticket2
'                                liTemp.SubItems(ID_PreferentialPrice3) = rsCountTemp!preferential_ticket3
'                            Case cszBedType
'                                liTemp.SubItems(ID_BedFullPrice) = rsCountTemp!full_price
'                                liTemp.SubItems(ID_BedHalfPrice) = rsCountTemp!half_price
'                                liTemp.SubItems(ID_BedPreferentialPrice1) = rsCountTemp!preferential_ticket1
'                                liTemp.SubItems(ID_BedPreferentialPrice2) = rsCountTemp!preferential_ticket2
'                                liTemp.SubItems(ID_BedPreferentialPrice3) = rsCountTemp!preferential_ticket3
'                            Case cszAdditionalType
'                                liTemp.SubItems(ID_AdditionalFullPrice) = rsCountTemp!full_price
'                                liTemp.SubItems(ID_AdditionalHalfPrice) = rsCountTemp!half_price
'                                liTemp.SubItems(ID_AdditionalPreferential1) = rsCountTemp!preferential_ticket1
'                                liTemp.SubItems(ID_AdditionalPreferential2) = rsCountTemp!preferential_ticket2
'                                liTemp.SubItems(ID_AdditionalPreferential3) = rsCountTemp!preferential_ticket3
'
'                        End Select
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
'                  varBookmark = rsCountTemp.Bookmark
'                    If rsCountTemp.RecordCount <> 0 Then
'                       RefreshSellStation rsCountTemp
'                    End If
'                  rsCountTemp.Bookmark = varBookmark
                    If nBusType <> TP_ScrollBus Then
                        liTemp.SubItems(ID_BusType) = Trim(rsCountTemp!bus_type)
                        liTemp.SubItems(ID_OffTime) = Format(rsCountTemp!BusStartTime, "hh:mm")
                        
                    Else
                        liTemp.SubItems(ID_VehicleModel) = cszScrollBus
                        liTemp.SubItems(ID_OffTime) = cszScrollBus
                        
                    End If
                    liTemp.SubItems(ID_RouteName) = Trim(rsCountTemp!route_name)
                    liTemp.SubItems(ID_EndStation) = RTrim(rsCountTemp!end_station_name)
                    liTemp.SubItems(ID_TotalSeat) = rsCountTemp!total_seat
                            If IsDate(liTemp.SubItems(ID_OffTime)) Then
                                If g_bIsBookValid And DateAdd("n", -g_nBookTime, liTemp.SubItems(ID_OffTime)) < Time And ToDBDate(flblSellDate.Caption) = ToDBDate(Date) Then
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
                   ' liTemp.SubItems(ID_StandCount) = rsCountTemp！sale_stand_seat_quantity
                   liTemp.Tag = MakeDisplayString(Trim(rsCountTemp!sell_station_id), Trim(rsCountTemp!sell_station_name))
                End If
                
nextstep:
                rsCountTemp.MoveNext
            Loop
        Else

            lvBus.ListItems.Clear
        End If

    End If
'    Set rsTemp = Nothing
    
    '设定某个站点？车次显示在第一行   ********以下陈峰加***********
    Dim nStationCount As Integer
    Dim nBusCount As Integer
    Dim nIndex As Integer
    Dim nListIndex  As Integer
    Dim nIndex2 As Integer
    Dim aszPreviousBusProperty() As String
    nStationCount = ArrayLength(m_aszFirstStation)
    nBusCount = ArrayLength(m_aszFirstBus)
    For nIndex = 1 To nStationCount
        If Trim(szStationID) = Trim(m_aszFirstStation(nIndex)) Then
            Exit For
        End If
    Next
    If nIndex <= nStationCount Then
        lvBus.Sorted = False
        For nIndex = 1 To nBusCount
            For nListIndex = 1 To lvBus.ListItems.count
                If Trim(lvBus.ListItems(nListIndex)) = Trim(m_aszFirstBus(nIndex)) Then
                    With lvBus.ListItems(nListIndex)
                        lForeColor = .ForeColor
                        ReDim aszPreviousBusProperty(1 To .ListSubItems.count + 1)
                        aszPreviousBusProperty(1) = .Text
                        For nIndex2 = 1 To .ListSubItems.count
                            aszPreviousBusProperty(nIndex2 + 1) = .SubItems(nIndex2)
                        Next nIndex2
                    End With
                    lvBus.ListItems.Remove nListIndex
                    Set liTemp = lvBus.ListItems.Add(1, "A" & aszPreviousBusProperty(1), aszPreviousBusProperty(1))
                    
                    For nIndex2 = 2 To ArrayLength(aszPreviousBusProperty)
                        liTemp.ListSubItems.Add(, , aszPreviousBusProperty(nIndex2)).ForeColor = lForeColor
                    Next nIndex2
                    
                    
                End If
                    
            Next
        
        Next nIndex
        If lvBus.ListItems.count > 0 Then
            lvBus.ListItems(1).Selected = True
            lvBus.ListItems(1).EnsureVisible
        End If
    Else
        If lvBus.ListItems.count > 0 Then
            lvBus.SortKey = MDISellTicket.GetSortKey() - 1
            lvBus.Sorted = True
            lvBus.ListItems(1).Selected = True
            lvBus.ListItems(1).EnsureVisible
        End If
    End If
    '设定某个站点？车次显示在第一行   ********以上陈峰加***********
    
    
    '设定某个站点？车次显示在第一行   ********以下 陈峰 注释*********
'    If lvBus.ListItems.Count > 0 Then
'        lvBus.SortKey = MDISellTicket.GetSortKey() - 1
'        'lvBus.Sorted = True
'        lvBus.ListItems(1).Selected = True
'        lvBus.ListItems(1).EnsureVisible
'    End If
    '设定某个站点？车次显示在第一行   ********以上 陈峰 注释***********
    '调用车次改变要进行相应操作的方法
    If lvBus.ListItems.count > 0 Then
        RefreshSellStation rsCountTemp
    Else
        lvSellStation.ListItems.Clear
    End If
    DoThingWhenBusChange

    Exit Sub
here:
    ShowErrorMsg

End Sub


Private Sub ShowOldTicketInfo()
    On Error GoTo here
    If txtOldTktNum.Text <> "" And m_bTicketInfoDirty Then
    
        Erase m_aszCardInfo
        Set m_vaCardInfo = Nothing
    
        Dim oTicket As ServiceTicket
        Dim oParam As New SystemParam
        Set oTicket = m_oSell.GetServerTicketAtCurrentUnit(txtOldTktNum.Text)
        lblStartStation.Caption = oTicket.SellStationName
        lblBusID.Caption = oTicket.REBusID
        lblBusDate.Caption = ToStandardDateStr(oTicket.REBusDate)
        lblStatus.Caption = GetTicketStatusStr(oTicket.TicketStatus)
        lblEndStation.Caption = oTicket.ToStationName
        lblPrice.Caption = FormatMoney(oTicket.TicketPrice)
        lblOldTicketPrice.Caption = FormatMoney(oTicket.TicketPrice)
'        Dim oUser As New User
'        oUser.Init m_oAUser
'        oUser.Identify oTicket.Operator
        lblSeller.Caption = oTicket.Operator ' MakeDisplayString(oUser.UserID, oUser.FullName)
        
        lblSellTime.Caption = ToStandardDateTimeStr(oTicket.SellTime)
        lblTicketType.Caption = GetTicketTypeStr2(oTicket.TicketType)
        lblOffTime.Caption = ToStandardTimeStr(oTicket.dtBusStartUpTime)
        
        ReDim Preserve m_aszCardInfo(1 To 1)
        m_aszCardInfo(1).szCardType = Trim(oTicket.CardType)
        m_aszCardInfo(1).szIDCardNo = Trim(oTicket.IDCardNo)
        m_aszCardInfo(1).szPersonName = Trim(oTicket.PersonName)
        m_aszCardInfo(1).szSex = Trim(oTicket.Sex)
        m_aszCardInfo(1).szPersonPicture = Trim(oTicket.PersonPicture)
        
        oParam.Init m_oAUser
        lblChangeCharge.Caption = oParam.ChangeCharge
        
        m_bTicketInfoDirty = False
    End If
    
    m_vaCardInfo = m_aszCardInfo
    
    Set oParam = Nothing
    Set oTicket = Nothing
    Exit Sub
here:
    SetDefaultLabel
    ShowErrorMsg
End Sub

Private Sub txtReceivedMoney_GotFocus()
    lblFactIn.ForeColor = clActiveColor
End Sub

Private Sub txtReceivedMoney_KeyDown(KeyCode As Integer, Shift As Integer)
Dim nCount As Integer

For nCount = 1 To Len(txtReceivedMoney.Text)
    If Mid(txtReceivedMoney.Text, nCount, 1) = "." Then
       blPointCount = True
       Exit For
    End If
Next nCount
End Sub

Private Sub txtReceivedMoney_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 8 Then
    If KeyAscii < 48 Or KeyAscii > 57 Then
       If blPointCount = True And KeyAscii = 46 Then
          KeyAscii = 0
       ElseIf KeyAscii <> 46 Then
          KeyAscii = 0
       End If
    End If
Else
    If KeyAscii = 13 Then
        'txtOldTktNum.SetFocus
'        cmdSell_Click
         txtSheetID.SetFocus
    End If
End If

blPointCount = False
End Sub

Private Sub txtReceivedMoney_KeyUp(KeyCode As Integer, Shift As Integer)
Dim sgvalue As Double
    
    If txtReceivedMoney.Text = "" Then
       sgvalue = 0
    Else
       sgvalue = CDbl(txtReceivedMoney.Text)
    End If
    If sgvalue - CDbl(lblRightMoney) <= 0 Then txtReceivedMoney.SetFocus
End Sub


'得到此次售票的相应序号的座号
Private Function SelfGetSeatNo(pnIndex As Integer) As String
    If chkSetSeat.Enabled = False Then  '不可选座位,则为站票
        SelfGetSeatNo = "ST"
    ElseIf chkSetSeat.Enabled And txtSeat.Text <> "" Then 'And chkSetSeat.Value = 1 Then '如果定座选中,则得到相应的座号
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

'Private Sub txtSheetID_Change()
'    EnableSellButton
'End Sub

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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    DealWithChildKeyDown KeyCode, Shift
    If KeyCode = vbKeyF8 Then
        '定座位
        If cmdSetSeat.Enabled = True Then
            cmdSetSeat_Click
        End If
    ElseIf KeyCode = vbKeyF12 Then
        '选中需要保险
        If cboInsurance.ListIndex < cboInsurance.ListCount - 1 Then
            cboInsurance.ListIndex = cboInsurance.ListIndex + 1
        Else
            cboInsurance.ListIndex = 0
        End If
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

Private Sub RefreshStation2()
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
        .ListFields = "station_input_code:4,station_name"
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
    ShowErrorMsg
End Sub

Private Sub txtReceivedMoney_LostFocus()
    lblFactIn.ForeColor = 0
End Sub

Private Sub txtReceivedMoney_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = vbEnter Then txtReceivedMoney.SetFocus
End Sub

Private Sub ShowNullTicketInfo()
On Error GoTo here

    lblStartStation.Caption = ""
    lblBusID.Caption = ""
    lblBusDate.Caption = ""
    lblStatus.Caption = ""
    lblEndStation.Caption = ""
    lblPrice.Caption = FormatMoney("0")
    lblOldTicketPrice.Caption = FormatMoney("0")
    lblSeller.Caption = ""
    lblChangeCharge.Caption = FormatMoney("0")
    lblSellTime.Caption = ""
    lblTicketType.Caption = ""
    lblOffTime.Caption = ""
    
    Exit Sub
here:
    ShowErrorMsg
End Sub

'读取优惠票信息
Private Sub RefreshPreferentialTicket()
    Dim atTicketType() As TTicketType
    Dim szHeadText As String
    Dim sgWidth As Single
    Dim nCount As Integer
    Dim i As Integer, j As Integer
    
    Dim aszSeatType() As String
    Dim nlen As Integer
    Dim nUsedPerential As Integer
    On Error GoTo here
    '得到所有的票种
    atTicketType = m_oSell.GetAllTicketType()
    aszSeatType = m_oSell.GetAllSeatType
    nlen = ArrayLength(aszSeatType)
    nCount = ArrayLength(atTicketType)
    sgWidth = 690
    lvBus.ColumnHeaders.Clear
    '添加ListView列头
    With lvBus.ColumnHeaders
        .Add , , "车次", 950 '"BusID"
        .Add , , "车次类型", 0 '"BusType"
        .Add , , "时间", 850 '"OffTime"
        .Add , , "线路名称", 1900
        .Add , , "终到站", 930 '"EndStation"
        .Add , , "总", 440
        .Add , , "订", 440
        .Add , , "座位", 700 '"SeatCount"
        .Add , , "座", 0
        .Add , , "卧", 0 '440
        .Add , , "加", 0 '440
        .Add , , "车型", 900 '"BusModel"
        '添加票种,不可用的则宽度设为0
        For i = 1 To nCount     '座位票价
            If atTicketType(i).nTicketTypeID = TP_FullPrice Then
                .Add , , "座全", IIf(atTicketType(i).nTicketTypeValid = TP_TicketTypeValid, sgWidth, 0)
            Else
                .Add , , Trim(atTicketType(i).szTicketTypeName), IIf(atTicketType(i).nTicketTypeValid = TP_TicketTypeValid, sgWidth, 0)
                If atTicketType(i).nTicketTypeID = TP_HalfPrice Then optHalfTicket.Caption = Trim(atTicketType(i).szTicketTypeName) & "(&X)" & ":"
            End If
        Next i
        For i = 1 To nCount   '卧铺票价
            If atTicketType(i).nTicketTypeID = TP_FullPrice Then
                .Add , , "卧全", IIf(atTicketType(i).nTicketTypeValid = TP_TicketTypeValid, 0, 0)
            Else
                .Add , , Trim(atTicketType(i).szTicketTypeName), IIf(atTicketType(i).nTicketTypeValid = TP_TicketTypeValid, 0, 0)
            End If
        Next i
        For i = 1 To nCount  '加座票价
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
        If cboSeatType.ListCount > 0 Then
            cboSeatType.ListIndex = 0
        End If
    End If
    '设置ComboBox和优惠票是否可用
    nUsedPerential = 0
    For i = 1 To nCount
    
        If atTicketType(i).nTicketTypeID = TP_HalfPrice And atTicketType(i).nTicketTypeValid = TP_TicketTypeValid Then
            optHalfTicket.Enabled = True
        ElseIf atTicketType(i).nTicketTypeID > TP_HalfPrice And atTicketType(i).nTicketTypeValid = TP_TicketTypeValid Then
            cboPreferentialTicket.AddItem Trim(atTicketType(i).szTicketTypeName)
            nUsedPerential = nUsedPerential + 1
        End If
    Next i
    
    If cboPreferentialTicket.ListCount < 1 Then
        optPreferentialTicket.Enabled = False
        cboPreferentialTicket.Enabled = False
        cboPreferentialTicket.Text = ""
    Else
        optPreferentialTicket.Enabled = True
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
    Exit Sub
here:
    ShowErrorMsg
End Sub

'得到对应的优惠票种的对应的票价
Private Function GetPreferentialPrice(Optional pbIsSell As Boolean = False) As Double
Dim liTemp As ListItem
Dim dbTemp As Double
    Set liTemp = lvBus.SelectedItem
    Select Case Trim(m_aszSeatType(cboSeatType.ListIndex, 1))
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
                    dbTemp = CDbl(liTemp.SubItems(IIf(pbIsSell, ID_AdditionalFullPrice, ID_AdditionalFreePrice)))
                Case TP_PreferentialTicket1
                    dbTemp = CDbl(liTemp.SubItems(ID_AdditionalPreferential1))
                Case TP_PreferentialTicket2
                    dbTemp = CDbl(liTemp.SubItems(ID_AdditionalPreferential2))
                Case TP_PreferentialTicket3
                    dbTemp = CDbl(liTemp.SubItems(ID_AdditionalPreferential3))
                End Select
            End If
    End Select

    GetPreferentialPrice = dbTemp
    Set liTemp = Nothing
End Function


'设置预售信息
Private Sub SetPreSellInfo()
    Dim atTicketType() As TTicketType
    Dim nCount As Integer
    Dim i As Integer
    atTicketType = m_oSell.GetAllTicketType()
    nCount = ArrayLength(atTicketType)
    
    lvPreSell.ColumnHeaders.Clear
    With lvPreSell.ColumnHeaders
        .Add , , "车次", 950 '"BusID"
        .Add , , "车次类型", 0 '"BusType"
        .Add , , "乘车日期", 1200 '"BusDate"
        .Add , , "终到站", 899 '"EndStation"
        .Add , , "时间", 1000 '"OffTime"
        .Add , , "车型", 899 '"VehicleModel"
        For i = 1 To nCount
            .Add , , Trim(atTicketType(i).szTicketTypeName), IIf(atTicketType(i).nTicketTypeValid = TP_TicketTypeValid, 0, 0)
            .Add , , Trim(atTicketType(i).szTicketTypeName) & "票数", IIf(atTicketType(i).nTicketTypeValid = TP_TicketTypeValid, 1200, 0)
        Next i
        .Add , , "到站代码", 0    '"EndStationCode"
        .Add , , "原票号", 950 '"OriginalTicket"
        .Add , , "改签凭证", 1200   '"ChangeSheet"
        .Add , , "检票口", 750 '"CheckGate"
        .Add , , "折扣率", 750 '"Discount"
        .Add , , "改签费", 750   '"ChangeFees"
        .Add , , "原票价", 750   '"OriginalPrice"
        .Add , , "座位状态1", 0    '"SeatStatus1"
        .Add , , "座位状态2", 0    '"SeatStatus2"
        .Add , , "座位号", 0     '"SeatNo"
        .Add , , "票价明细", 0  '"AllTicketPrice"
        .Add , , "票种明细", 0 '"AllTicketType"
        .Add , , "总票数", 0 '"TotalNum"
        .Add , , "座位类型", 0
        .Add , , "终点站", 0
    End With
End Sub

'得到预售信息
Private Sub GetPreSellInfo()
    Dim liPreSell As ListItem
    Dim liBus As ListItem
    If Not lvBus.SelectedItem Is Nothing Then
        If IsSameOldTicket Then
            MsgBox "有相同的改签原票号！", vbInformation + vbOKOnly, "提示"
        Else
            Set liBus = lvBus.SelectedItem
            Set liPreSell = lvPreSell.ListItems.Add(, , lvBus.SelectedItem.Text)
            With liPreSell
                .Tag = lvBus.SelectedItem.Tag
                .SubItems(CC_BusType) = lvBus.SelectedItem.SubItems(ID_BusType)
                .SubItems(CC_BusDate) = CDate(flblSellDate.Caption)
                .SubItems(CC_EndStation) = GetStationNameInCbo(cboEndStation.Text)
                .SubItems(CC_OffTime) = lvBus.SelectedItem.SubItems(ID_OffTime)
                .SubItems(CC_VehicleModel) = lvBus.SelectedItem.SubItems(ID_VehicleModel)
                Select Case Trim(m_aszSeatType(cboSeatType.ListIndex + 1, 1))
                    Case cszSeatType
                        If optFullTicket.Value Then
                           .SubItems(CC_FullPrice) = FormatMoney(lvBus.SelectedItem.SubItems(ID_FullPrice))
                           .SubItems(CC_FullNum) = 1
                        ElseIf optHalfTicket.Value Then
                           
                           .SubItems(CC_HalfPrice) = FormatMoney(lvBus.SelectedItem.SubItems(ID_HalfPrice))
                           .SubItems(CC_HalfNum) = 1
                        Else
                           If cboPreferentialTicket.ListCount > 0 Then
                               Select Case m_atTicketType(cboPreferentialTicket.ListIndex + 1).nTicketTypeID
                               Case TP_FreeTicket
                                   .SubItems(CC_FreePrice) = 0
                                   .SubItems(CC_FreeNum) = 1
                               Case TP_PreferentialTicket1
                                   .SubItems(CC_PreferentialPrice1) = FormatMoney(lvBus.SelectedItem.SubItems(ID_PreferentialPrice1))
                                   .SubItems(CC_PreferentialNum1) = 1
                               Case TP_PreferentialTicket2
                                   .SubItems(CC_PreferentialPrice2) = FormatMoney(lvBus.SelectedItem.SubItems(ID_PreferentialPrice2))
                                   .SubItems(CC_PreferentialNum2) = 1
                               Case TP_PreferentialTicket3
                                   .SubItems(CC_PreferentialPrice3) = FormatMoney(lvBus.SelectedItem.SubItems(ID_PreferentialPrice3))
                                   .SubItems(CC_PreferentialNum3) = 1
                               End Select
                           End If
                          
                        End If
                    Case cszBedType
                        If optFullTicket.Value Then
                           .SubItems(CC_FullPrice) = FormatMoney(lvBus.SelectedItem.SubItems(ID_BedFullPrice))
                           .SubItems(CC_FullNum) = 1
                        ElseIf optHalfTicket.Value Then
                           
                           .SubItems(CC_HalfPrice) = FormatMoney(lvBus.SelectedItem.SubItems(ID_BedHalfPrice))
                           .SubItems(CC_HalfNum) = 1
                        Else
                           If cboPreferentialTicket.ListCount > 0 Then
                               Select Case m_atTicketType(cboPreferentialTicket.ListIndex + 1).nTicketTypeID
                               Case TP_FreeTicket
                                   .SubItems(CC_FreePrice) = 0
                                   .SubItems(CC_FreeNum) = 1
                               Case TP_PreferentialTicket1
                                   .SubItems(CC_PreferentialPrice1) = FormatMoney(lvBus.SelectedItem.SubItems(ID_BedPreferentialPrice1))
                                   .SubItems(CC_PreferentialNum1) = 1
                               Case TP_PreferentialTicket2
                                   .SubItems(CC_PreferentialPrice2) = FormatMoney(lvBus.SelectedItem.SubItems(ID_BedPreferentialPrice2))
                                   .SubItems(CC_PreferentialNum2) = 1
                               Case TP_PreferentialTicket3
                                   .SubItems(CC_PreferentialPrice3) = FormatMoney(lvBus.SelectedItem.SubItems(ID_BedPreferentialPrice3))
                                   .SubItems(CC_PreferentialNum3) = 1
                               End Select
                           End If
                          
                        End If
                    Case cszAdditionalType
                        If optFullTicket.Value Then
                           .SubItems(CC_FullPrice) = FormatMoney(lvBus.SelectedItem.SubItems(ID_AdditionalFullPrice))
                           .SubItems(CC_FullNum) = 1
                        ElseIf optHalfTicket.Value Then
                           
                           .SubItems(CC_HalfPrice) = FormatMoney(lvBus.SelectedItem.SubItems(ID_AdditionalHalfPrice))
                           .SubItems(CC_HalfNum) = 1
                        Else
                           If cboPreferentialTicket.ListCount > 0 Then
                               Select Case m_atTicketType(cboPreferentialTicket.ListIndex + 1).nTicketTypeID
                               Case TP_FreeTicket
                                   .SubItems(CC_FreePrice) = 0
                                   .SubItems(CC_FreeNum) = 1
                               Case TP_PreferentialTicket1
                                   .SubItems(CC_PreferentialPrice1) = FormatMoney(lvBus.SelectedItem.SubItems(ID_AdditionalPreferential1))
                                   .SubItems(CC_PreferentialNum1) = 1
                               Case TP_PreferentialTicket2
                                   .SubItems(CC_PreferentialPrice2) = FormatMoney(lvBus.SelectedItem.SubItems(ID_AdditionalPreferential2))
                                   .SubItems(CC_PreferentialNum2) = 1
                               Case TP_PreferentialTicket3
                                   .SubItems(CC_PreferentialPrice3) = FormatMoney(lvBus.SelectedItem.SubItems(ID_AdditionalPreferential3))
                                   .SubItems(CC_PreferentialNum3) = 1
                               End Select
                           End If
                          
                        End If
                End Select
                
                .SubItems(CC_EndStationCode) = cboEndStation.BoundText
                .SubItems(CC_OriginalTicket) = txtOldTktNum.Text
                .SubItems(CC_ChangeSheet) = txtSheetID.Text
                .SubItems(CC_CheckGate) = lvBus.SelectedItem.SubItems(ID_CheckGate)
                .SubItems(CC_Discount) = txtDiscount.Text
                .SubItems(CC_ChangeFees) = lblChangeCharge.Caption
                .SubItems(CC_OriginalPrice) = lblPrice.Caption
                .SubItems(CC_SeatStatus1) = chkSetSeat.Enabled
                .SubItems(CC_SeatStatus2) = chkSetSeat.Value
                .SubItems(CC_SeatNo) = txtSeat.Text
                .SubItems(CC_AllTicketPrice) = FormatMoney(Val(.SubItems(CC_FullNum)) * Val(.SubItems(CC_FullPrice)) + _
                                                Val(.SubItems(CC_HalfNum)) * Val(.SubItems(CC_HalfPrice)) + _
                                                Val(.SubItems(CC_PreferentialNum1)) * Val(.SubItems(CC_PreferentialPrice1)) + _
                                                Val(.SubItems(CC_PreferentialNum2)) * Val(.SubItems(CC_PreferentialPrice2)) + _
                                                Val(.SubItems(CC_PreferentialNum3)) * Val(.SubItems(CC_PreferentialPrice3)))
                
                .SubItems(CC_TotalNum) = 1
                .SubItems(CC_SeatType) = m_aszSeatType(cboSeatType.ListIndex + 1, 1)
                .SubItems(CC_TerminateName) = lvBus.SelectedItem.SubItems(ID_EndStation)
            End With
        End If
    End If
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

'处理预售按钮状态
Private Sub DealPreButtonStatus()
If lvBus.ListItems.count <> 0 And (Not lvBus.SelectedItem Is Nothing) Then
    cmdPreSell.Enabled = True
Else
    cmdPreSell.Enabled = False
End If
End Sub

'得到应收总票款
Private Sub GetRightMoney()
    Dim iCount As Integer
    Dim RightMoney As Double
    RightMoney = 0
    If lvPreSell.ListItems.count <> 0 Then
        For iCount = 1 To lvPreSell.ListItems.count
            With lvPreSell.ListItems(iCount)
                RightMoney = RightMoney + Val(.SubItems(CC_AllTicketPrice)) + Val(.SubItems(CC_ChangeFees)) - Val(.SubItems(CC_OriginalPrice))
            End With
        Next iCount
        lblRightMoney.Caption = FormatMoney(RightMoney)
    Else
        lblRightMoney.Caption = FormatMoney(CDbl(lblTotalPrice.Caption) - CDbl(lblOldTicketPrice.Caption) + CDbl(lblChangeCharge.Caption))
    End If
End Sub

'判断是否有相同的改签原票
Private Function IsSameOldTicket() As Boolean
    Dim iCount As Integer
    If lvPreSell.ListItems.count <> 0 Then
        For iCount = 1 To lvPreSell.ListItems.count
            If txtOldTktNum.Text = lvPreSell.ListItems(iCount).SubItems(CC_OriginalTicket) Then
                IsSameOldTicket = True
                Exit Function
            End If
        Next iCount
    End If
    IsSameOldTicket = False
End Function

Private Sub txtSheetID_Change()
    DealPreButtonStatus
End Sub

'预售用订座
Private Function PreOrderSeat() As String
Dim i As Integer
Dim szTemp As String
Dim liTemp As ListItem
If lvPreSell.ListItems.count <> 0 Then
    For i = 1 To lvPreSell.ListItems.count
        Set liTemp = lvPreSell.ListItems(i)
        If CDate(flblSellDate.Caption) = CDate(liTemp.SubItems(CC_BusDate)) And lvBus.SelectedItem.Text = liTemp.Text Then
            If liTemp.SubItems(CC_SeatNo) <> "" Then
                szTemp = szTemp & "," & liTemp.SubItems(CC_SeatNo)
            End If
        End If
    Next i
Else
    szTemp = ""
End If
PreOrderSeat = szTemp
End Function

'返回相同车次索引
Private Function GetSameBusIndex() As Integer
Dim i As Integer
Dim liTemp As ListItem
Dim liSelected As ListItem
If lvPreSell.ListItems.count <> 0 And (Not lvBus.SelectedItem Is Nothing) Then
    Set liSelected = lvBus.SelectedItem
    For i = 1 To lvPreSell.ListItems.count
        Set liTemp = lvPreSell.ListItems(i)
        If liTemp.Text = liSelected.Text And liTemp.SubItems(CC_BusDate) = CDate(flblSellDate.Caption) And liTemp.SubItems(IT_BoundText) = cboEndStation.BoundText Then
            GetSameBusIndex = i
            Exit Function
        End If
    Next i
End If
GetSameBusIndex = 0
End Function

'合并相同车次信息
Private Sub MergeSameBusInfo(nSameIndex As Integer)
Dim i As Integer
Dim szTicketPrice As String
Dim szTicketType As String
Dim liTemp As ListItem
Set liTemp = lvPreSell.ListItems(nSameIndex)
With liTemp
    If optFullTicket.Value Then
       .SubItems(CC_FullNum) = 1 + .SubItems(CC_FullNum)
    ElseIf optHalfTicket.Value Then
       .SubItems(CC_HalfNum) = 1 + .SubItems(CC_HalfNum)
    Else
       If cboPreferentialTicket.ListCount > 0 Then
           Select Case m_atTicketType(cboPreferentialTicket.ListIndex + 1).nTicketTypeID
           Case TP_FreeTicket
               .SubItems(CC_FreeNum) = 1 + .SubItems(CC_FreeNum)
           Case TP_PreferentialTicket1
               .SubItems(CC_PreferentialNum1) = 1 + .SubItems(CC_PreferentialNum1)
           Case TP_PreferentialTicket2
               .SubItems(CC_PreferentialNum2) = 1 + .SubItems(CC_PreferentialNum2)
           Case TP_PreferentialTicket3
               .SubItems(CC_PreferentialNum3) = 1 + .SubItems(CC_PreferentialNum3)
           End Select
       End If
    End If
    .SubItems(CC_OriginalPrice) = .SubItems(CC_OriginalPrice) & "," & lblPrice.Caption
    .SubItems(CC_OriginalTicket) = .SubItems(CC_OriginalTicket) & "," & txtOldTktNum.Text
    .SubItems(CC_ChangeSheet) = .SubItems(CC_ChangeSheet) & "," & Trim(txtSheetID.Text)
    If txtSeat.Text <> "" Then
        .SubItems(CC_SeatNo) = .SubItems(CC_SeatNo) & "," & txtSeat.Text
    End If
    .SubItems(CC_TotalNum) = .SubItems(CC_TotalNum) + 1
    For i = 1 To .SubItems(CC_TotalNum)
        If i = .SubItems(CC_TotalNum) Then
            szTicketPrice = szTicketPrice & m_TicketPrice(i)
            szTicketType = szTicketType & m_TicketTypeDetail(i)
        Else
            szTicketPrice = szTicketPrice & m_TicketPrice(i) & ","
            szTicketType = szTicketType & m_TicketTypeDetail(i) & ","
        End If
    Next i
    .SubItems(CC_AllTicketPrice) = szTicketPrice
    .SubItems(CC_AllTicketType) = szTicketType
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
End Sub


Private Sub SetDefaultLabel()
    lblStartStation.Caption = ""
    lblBusID.Caption = ""
    lblBusDate.Caption = ""
    lblStatus.Caption = ""
    lblEndStation.Caption = ""
    lblTicketType.Caption = ""
    lblOffTime.Caption = ""
    lblSeller.Caption = ""
    lblTotalPrice.Caption = ""
    lblSellTime.Caption = ""
    lblPrice.Caption = ""
End Sub

Private Sub txtSheetID_GotFocus()
    lblCredence.ForeColor = clActiveColor

End Sub

Private Sub txtSheetID_LostFocus()
    lblCredence.ForeColor = 0
End Sub

Private Sub ChangeSell()

    Dim i As Integer
   
    
    '处理改签售票
    Dim srSellResult() As TSellTicketResult
    Dim dBusDate() As Date
    Dim szBusID() As String
    Dim szDesStationID() As String
    Dim btTicketInfoParam() As TSellTicketParam
    Dim psgDiscount() As Single
    Dim orgTicketID() As String
    Dim szSheetID() As String
    Dim szSellStationID() As String
    Dim szSellStationName() As String
    Dim szStartStationName As String
    
    '处理打印售票
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
    Dim anInsurance() As Integer '售票用
    Dim panInsurance() As Integer '打印用
    
    Dim liTemp As ListItem
    Dim nTemp As Integer
    
    Dim dbRealTotalPrice As Double
    Dim dbTotalPrice As Double
    Dim dbTotalChangeFees As Double
    Dim dbTotalOrgPrice As Double
    Dim dbTotalNewPrice As Double
    dbTotalNewPrice = 0
    dbRealTotalPrice = 0
    dbTotalPrice = 0
    dbTotalChangeFees = 0
    dbTotalOrgPrice = 0
    dbTotalPrice = FormatMoney(lblRightMoney.Caption)
    
    If txtDiscount.Text > 1 Then
        MsgBox "折扣率不能大于1", vbInformation, "提示"
        txtDiscount.SetFocus
        Exit Sub
    End If
    If txtOldTktNum.Text = "" Then
        txtOldTktNum.SetFocus
        Exit Sub
    End If
    
    On Error GoTo here
    
'以下是真正的改签售票处理
'-------------------------------------------------------------------------------------
    If lvPreSell.ListItems.count = 0 Then
        '没售多站票
        ReDim srSellResult(1 To 1)
        ReDim dBusDate(1 To 1)
        ReDim szBusID(1 To 1)
        ReDim szDesStationID(1 To 1)
        ReDim btTicketInfoParam(1 To 1)
        ReDim psgDiscount(1 To 1)
        ReDim orgTicketID(1 To 1)
        ReDim szSheetID(1 To 1)
        ReDim pszTerminateName(1 To 1)
        ReDim szSellStationID(1 To 1)
        ReDim szSellStationName(1 To 1)
        ReDim anInsurance(1 To 1)
        
        lblChangeMsg.Caption = "正在处理改签"
        lblChangeMsg.Refresh
        ReDim btTicketInfoParam(1).BuyTicketInfo(1 To 1)
        ReDim btTicketInfoParam(1).pasgSellTicketPrice(1 To 1)
        ReDim btTicketInfoParam(1).aszOrgTicket(1 To 1)
        ReDim btTicketInfoParam(1).aszChangeSheetID(1 To 1)
        
        Set liTemp = lvBus.SelectedItem
        If optFullTicket.Value Then
            btTicketInfoParam(1).BuyTicketInfo(1).nTicketType = TP_FullPrice
            btTicketInfoParam(1).pasgSellTicketPrice(1) = CDbl(liTemp.SubItems(ID_FullPrice)) * CSng(txtDiscount.Text)
    
        ElseIf optHalfTicket.Value Then
            btTicketInfoParam(1).BuyTicketInfo(1).nTicketType = TP_HalfPrice
            btTicketInfoParam(1).pasgSellTicketPrice(1) = CDbl(liTemp.SubItems(ID_HalfPrice)) * CSng(txtDiscount.Text)
    
        Else
            btTicketInfoParam(1).BuyTicketInfo(1).nTicketType = m_atTicketType(cboPreferentialTicket.ListIndex + 1).nTicketTypeID
            btTicketInfoParam(1).pasgSellTicketPrice(1) = GetPreferentialPrice(True) * CSng(txtDiscount.Text)
        End If
        btTicketInfoParam(1).BuyTicketInfo(1).szTicketNo = GetTicketNo()
        btTicketInfoParam(1).BuyTicketInfo(1).szSeatNo = SelfGetSeatNo(1)
        btTicketInfoParam(1).BuyTicketInfo(1).szSeatTypeID = m_aszSeatType(cboSeatType.ListIndex + 1, 1)
        btTicketInfoParam(1).BuyTicketInfo(1).szSeatTypeName = cboSeatType.Text
        dBusDate(1) = CDate(flblSellDate.Caption)
        szBusID(1) = lvBus.SelectedItem.Text
        szDesStationID(1) = cboEndStation.BoundText
        psgDiscount(1) = CSng(txtDiscount.Text)
        btTicketInfoParam(1).aszOrgTicket(1) = txtOldTktNum.Text
        btTicketInfoParam(1).aszChangeSheetID(1) = txtSheetID.Text
        
        
        
        dbTotalChangeFees = Format(lblChangeCharge.Caption)
        dbTotalOrgPrice = FormatMoney(lblOldTicketPrice.Caption)
        szSellStationID(1) = ResolveDisplay(lvBus.SelectedItem.Tag, szStartStationName)
        szSellStationName(1) = szStartStationName
        
        anInsurance(1) = Val(cboInsurance.Text)
        
        If m_oAUser.SellStationID <> szSellStationID(1) Then
            If MsgBox("此票目前只能改签成" & szSellStationName(1) & "始发的票，是否要继续?", vbYesNo + vbInformation + vbDefaultButton2, "注意") = vbNo Then
                Exit Sub
            End If
        End If
            
        
        srSellResult = m_oSell.ChangeTicket(dBusDate, szBusID, szSellStationID, szDesStationID, btTicketInfoParam, anInsurance, m_vaCardInfo)
        
        IncTicketNo
        RestoreStatusInMDI
        If chkSetSeat.Enabled Then
            DecBusListViewSeatInfo lvBus, 1, True
        Else
            DecBusListViewSeatInfo lvBus, 1, False
        End If
        If lvBus.SelectedItem.SubItems(ID_LimitedCount) > 0 Then
            lvBus.SelectedItem.SubItems(ID_LimitedCount) = lvBus.SelectedItem.SubItems(ID_LimitedCount) - 1
            flblLimitedCount.Caption = GetStationLimitedCountStr(CInt(lvBus.SelectedItem.SubItems(ID_LimitedCount)))
        End If
        
        '以下处理打印票
        '-----------------------------------------------------------
        ReDim apiTicketInfo(1 To 1)
        ReDim pszBusDate(1 To 1)
        ReDim pnTicketCount(1 To 1)
        ReDim pszEndStation(1 To 1)
        ReDim pszOffTime(1 To 1)
        ReDim pszBusID(1 To 1)
        ReDim pszVehicleType(1 To 1)
        ReDim pszCheckGate(1 To 1)
        ReDim pbSaleChange(1 To 1)
        ReDim apiTicketInfo(1).aptPrintTicketInfo(1 To 1)
        ReDim panInsurance(1 To 1)
        
        
        
        pszBusDate(1) = CDate(flblSellDate.Caption)
        pnTicketCount(1) = 1
        pszEndStation(1) = GetStationNameInCbo(cboEndStation.Text)
        pszOffTime(1) = lvBus.SelectedItem.SubItems(ID_OffTime)
        pszBusID(1) = lvBus.SelectedItem.Text
        pszCheckGate(1) = lvBus.SelectedItem.SubItems(ID_CheckGate)
        pszVehicleType(1) = lvBus.SelectedItem.SubItems(ID_VehicleModel)
        
        pszTerminateName(1) = lvBus.SelectedItem.SubItems(ID_EndStation)
        
        panInsurance(1) = Val(cboInsurance.Text)
            
        pbSaleChange(1) = True
       
        apiTicketInfo(1).aptPrintTicketInfo(1).nTicketType = btTicketInfoParam(1).BuyTicketInfo(1).nTicketType
        apiTicketInfo(1).aptPrintTicketInfo(1).sgTicketPrice = srSellResult(1).asgTicketPrice(1)
        apiTicketInfo(1).aptPrintTicketInfo(1).szSeatNo = srSellResult(1).aszSeatNo(1)
        apiTicketInfo(1).aptPrintTicketInfo(1).szTicketNo = btTicketInfoParam(1).BuyTicketInfo(1).szTicketNo
        
        lblChangeMsg.Caption = "正在打印车票"
        
        PrintTicket apiTicketInfo, pszBusDate, pnTicketCount, pszEndStation, pszOffTime, pszBusID, pszVehicleType, pszCheckGate, pbSaleChange, pszTerminateName, szSellStationName, panInsurance, m_vaCardInfo
        
        m_aszInsurce = CombInsurance(apiTicketInfo, szBusID, pnTicketCount, pszBusDate, szDesStationID, pszEndStation, pszOffTime, szSellStationID, szSellStationName, m_aszCardInfo)
        SaveInsurance m_aszInsurce
        
        If btTicketInfoParam(1).BuyTicketInfo(1).nTicketType <> TP_FreeTicket Then
            dbTotalNewPrice = srSellResult(1).asgTicketPrice(1)
        End If
    Else
        '同时出售多个站的票
        Dim iLen As Integer
        Dim iCount As Integer
        Dim nOffset As Integer
        iLen = 0
        nOffset = 0
        iLen = lvPreSell.ListItems.count
        ReDim srSellResult(1 To iLen)
        ReDim dBusDate(1 To iLen)
        ReDim szBusID(1 To iLen)
        ReDim szDesStationID(1 To iLen)
        ReDim btTicketInfoParam(1 To iLen)
        ReDim psgDiscount(1 To iLen)
        ReDim orgTicketID(1 To iLen)
        ReDim szSheetID(1 To iLen)
        ReDim szSellStationID(1 To iLen)
        ReDim szSellStationName(1 To iLen)
        
        ReDim anInsurance(1 To iLen)
        
        lblChangeMsg.Caption = "正在处理改签"
        lblChangeMsg.Refresh
        For iCount = 1 To iLen
            
            With lvPreSell.ListItems(iCount)
                ReDim btTicketInfoParam(iCount).BuyTicketInfo(1 To .SubItems(CC_TotalNum))
                ReDim btTicketInfoParam(iCount).pasgSellTicketPrice(1 To .SubItems(CC_TotalNum))
                ReDim btTicketInfoParam(iCount).aszOrgTicket(1 To .SubItems(CC_TotalNum))
                ReDim btTicketInfoParam(iCount).aszChangeSheetID(1 To .SubItems(CC_TotalNum))
                
                nTemp = 0
                For i = 1 To Val(.SubItems(CC_FullNum))
                    btTicketInfoParam(iCount).BuyTicketInfo(i).nTicketType = TP_FullPrice
                    btTicketInfoParam(iCount).BuyTicketInfo(i).szTicketNo = GetTicketNo(i - 1 + nOffset)
                    btTicketInfoParam(iCount).BuyTicketInfo(i).szSeatNo = SelfGetSeatNo12(1, .SubItems(CC_SeatStatus1), .SubItems(CC_SeatStatus2), .SubItems(CC_SeatNo))
                    btTicketInfoParam(iCount).pasgSellTicketPrice(i) = CSng(.SubItems(CC_FullPrice)) * CSng(.SubItems(CC_Discount))
                    btTicketInfoParam(iCount).aszOrgTicket(i) = Trim(GetSeatNo(.SubItems(CC_OriginalTicket), i))
                    btTicketInfoParam(iCount).aszChangeSheetID(i) = Trim(GetSeatNo(.SubItems(CC_ChangeSheet), i))
                    btTicketInfoParam(iCount).BuyTicketInfo(i).szSeatTypeID = .SubItems(CC_SeatType)
                    btTicketInfoParam(iCount).BuyTicketInfo(i).szSeatTypeName = GetSeatTypeName(.SubItems(CC_SeatType))
                    
                    
                    dbTotalChangeFees = dbTotalChangeFees + CSng(.SubItems(CC_ChangeFees))
                    dbTotalOrgPrice = dbTotalOrgPrice + CSng(.SubItems(CC_OriginalPrice))
                    
                Next i
                
                nTemp = nTemp + Val(.SubItems(CC_FullNum))
                For i = 1 To Val(.SubItems(CC_HalfNum))
                    btTicketInfoParam(iCount).BuyTicketInfo(i + nTemp).nTicketType = TP_HalfPrice
                    btTicketInfoParam(iCount).BuyTicketInfo(i + nTemp).szTicketNo = GetTicketNo(i - 1 + nTemp + nOffset)
                    btTicketInfoParam(iCount).BuyTicketInfo(i + nTemp).szSeatNo = SelfGetSeatNo12(1, .SubItems(CC_SeatStatus1), .SubItems(CC_SeatStatus2), .SubItems(CC_SeatNo))
                    btTicketInfoParam(iCount).pasgSellTicketPrice(i + nTemp) = CSng(.SubItems(CC_HalfPrice)) * CSng(.SubItems(CC_Discount))
                    btTicketInfoParam(iCount).aszOrgTicket(i + nTemp) = Trim(GetSeatNo(.SubItems(CC_OriginalTicket), i))
                    btTicketInfoParam(iCount).aszChangeSheetID(i + nTemp) = Trim(GetSeatNo(.SubItems(CC_ChangeSheet), i))
                    btTicketInfoParam(iCount).BuyTicketInfo(i + nTemp).szSeatTypeID = .SubItems(CC_SeatType)
                    btTicketInfoParam(iCount).BuyTicketInfo(i + nTemp).szSeatTypeName = GetSeatTypeName(.SubItems(CC_SeatType))
                    
                    dbTotalChangeFees = dbTotalChangeFees + CSng(.SubItems(CC_ChangeFees))
                    dbTotalOrgPrice = dbTotalOrgPrice + CSng(.SubItems(CC_OriginalPrice))
                Next i
                
                nTemp = nTemp + Val(.SubItems(CC_HalfNum))
                For i = 1 To Val(.SubItems(CC_FreeNum))
                    btTicketInfoParam(iCount).BuyTicketInfo(i + nTemp).nTicketType = TP_FreeTicket
                    btTicketInfoParam(iCount).BuyTicketInfo(i + nTemp).szTicketNo = GetTicketNo(i - 1 + nTemp + nOffset)
                    btTicketInfoParam(iCount).BuyTicketInfo(i + nTemp).szSeatNo = SelfGetSeatNo12(1, .SubItems(CC_SeatStatus1), .SubItems(CC_SeatStatus2), .SubItems(CC_SeatNo))
                    btTicketInfoParam(iCount).pasgSellTicketPrice(i + nTemp) = CSng(.SubItems(CC_FullPrice)) * CSng(.SubItems(CC_Discount))
                    btTicketInfoParam(iCount).aszOrgTicket(i + nTemp) = Trim(GetSeatNo(.SubItems(CC_OriginalTicket), i))
                    btTicketInfoParam(iCount).aszChangeSheetID(i + nTemp) = Trim(GetSeatNo(.SubItems(CC_ChangeSheet), i))
                    btTicketInfoParam(iCount).BuyTicketInfo(i + nTemp).szSeatTypeID = .SubItems(CC_SeatType)
                    btTicketInfoParam(iCount).BuyTicketInfo(i + nTemp).szSeatTypeName = GetSeatTypeName(.SubItems(CC_SeatType))
                    
                    dbTotalChangeFees = dbTotalChangeFees + CSng(.SubItems(CC_ChangeFees))
                    dbTotalOrgPrice = dbTotalOrgPrice + CSng(.SubItems(CC_OriginalPrice))
                Next i
                
                nTemp = nTemp + Val(.SubItems(CC_FreeNum))
                For i = 1 To Val(.SubItems(CC_PreferentialNum1))
                    btTicketInfoParam(iCount).BuyTicketInfo(i + nTemp).nTicketType = TP_PreferentialTicket1
                    btTicketInfoParam(iCount).BuyTicketInfo(i + nTemp).szTicketNo = GetTicketNo(i - 1 + nTemp + nOffset)
                    btTicketInfoParam(iCount).BuyTicketInfo(i + nTemp).szSeatNo = SelfGetSeatNo12(1, .SubItems(CC_SeatStatus1), .SubItems(CC_SeatStatus2), .SubItems(CC_SeatNo))
                    btTicketInfoParam(iCount).pasgSellTicketPrice(i + nTemp) = CSng(.SubItems(CC_PreferentialPrice1)) * CSng(.SubItems(CC_Discount))
                    btTicketInfoParam(iCount).aszOrgTicket(i + nTemp) = Trim(GetSeatNo(.SubItems(CC_OriginalTicket), i))
                    btTicketInfoParam(iCount).aszChangeSheetID(i + nTemp) = Trim(GetSeatNo(.SubItems(CC_ChangeSheet), i))
                    btTicketInfoParam(iCount).BuyTicketInfo(i + nTemp).szSeatTypeID = .SubItems(CC_SeatType)
                    btTicketInfoParam(iCount).BuyTicketInfo(i + nTemp).szSeatTypeName = GetSeatTypeName(.SubItems(CC_SeatType))
                    
                    dbTotalChangeFees = dbTotalChangeFees + CSng(.SubItems(CC_ChangeFees))
                    dbTotalOrgPrice = dbTotalOrgPrice + CSng(.SubItems(CC_OriginalPrice))
                Next i
                
                nTemp = nTemp + Val(.SubItems(CC_PreferentialNum1))
                For i = 1 To Val(.SubItems(CC_PreferentialNum2))
                    btTicketInfoParam(iCount).BuyTicketInfo(i + nTemp).nTicketType = TP_PreferentialTicket2
                    btTicketInfoParam(iCount).BuyTicketInfo(i + nTemp).szTicketNo = GetTicketNo(i - 1 + nTemp + nOffset)
                    btTicketInfoParam(iCount).BuyTicketInfo(i + nTemp).szSeatNo = SelfGetSeatNo12(1, .SubItems(CC_SeatStatus1), .SubItems(CC_SeatStatus2), .SubItems(CC_SeatNo))
                    btTicketInfoParam(iCount).pasgSellTicketPrice(i + nTemp) = CSng(.SubItems(CC_PreferentialPrice2)) * CSng(.SubItems(CC_Discount))
                    btTicketInfoParam(iCount).aszOrgTicket(i + nTemp) = Trim(GetSeatNo(.SubItems(CC_OriginalTicket), i))
                    btTicketInfoParam(iCount).aszChangeSheetID(i + nTemp) = Trim(GetSeatNo(.SubItems(CC_ChangeSheet), i))
                    btTicketInfoParam(iCount).BuyTicketInfo(i + nTemp).szSeatTypeID = .SubItems(CC_SeatType)
                    btTicketInfoParam(iCount).BuyTicketInfo(i + nTemp).szSeatTypeName = GetSeatTypeName(.SubItems(CC_SeatType))
                    
                    dbTotalChangeFees = dbTotalChangeFees + CSng(.SubItems(CC_ChangeFees))
                    dbTotalOrgPrice = dbTotalOrgPrice + CSng(.SubItems(CC_OriginalPrice))
                Next i
                
                nTemp = nTemp + Val(.SubItems(CC_PreferentialNum2))
                For i = 1 To Val(.SubItems(CC_PreferentialNum3))
                    btTicketInfoParam(iCount).BuyTicketInfo(i + nTemp).nTicketType = TP_PreferentialTicket3
                    btTicketInfoParam(iCount).BuyTicketInfo(i + nTemp).szTicketNo = GetTicketNo(i - 1 + nTemp + nOffset)
                    btTicketInfoParam(iCount).BuyTicketInfo(i + nTemp).szSeatNo = SelfGetSeatNo12(1, .SubItems(CC_SeatStatus1), .SubItems(CC_SeatStatus2), .SubItems(CC_SeatNo))
                    btTicketInfoParam(iCount).pasgSellTicketPrice(i + nTemp) = CSng(.SubItems(CC_PreferentialPrice3)) * CSng(.SubItems(CC_Discount))
                    btTicketInfoParam(iCount).aszOrgTicket(i + nTemp) = Trim(GetSeatNo(.SubItems(CC_OriginalTicket), i))
                    btTicketInfoParam(iCount).aszChangeSheetID(i + nTemp) = Trim(GetSeatNo(.SubItems(CC_ChangeSheet), i))
                    
                    dbTotalChangeFees = dbTotalChangeFees + CSng(.SubItems(CC_ChangeFees))
                    dbTotalOrgPrice = dbTotalOrgPrice + CSng(.SubItems(CC_OriginalPrice))
                Next i
                
                dBusDate(iCount) = .SubItems(CC_BusDate)
                szBusID(iCount) = .Text
                szDesStationID(iCount) = .SubItems(CC_EndStationCode)
                
                psgDiscount(iCount) = .SubItems(CC_Discount)
                nOffset = .SubItems(CC_TotalNum) + nOffset
            End With
            
            If lvPreSell.ListItems.count < iCount Then
               szSellStationID(iCount) = ResolveDisplay(lvPreSell.ListItems(lvPreSell.ListItems.count).Tag, szStartStationName)
               
               szSellStationName(iCount) = szStartStationName
            Else
               szSellStationID(iCount) = ResolveDisplay(lvPreSell.ListItems(iCount).Tag, szStartStationName)
               szSellStationName(iCount) = szStartStationName
            End If
            
            anInsurance(iCount) = Val(cboInsurance.Text)
            
        Next iCount
        
        srSellResult = m_oSell.ChangeTicket(dBusDate, szBusID, szSellStationID, szDesStationID, btTicketInfoParam, anInsurance, m_vaCardInfo)
        IncTicketNo iLen
        RestoreStatusInMDI
        
         '以下是打印车票
        '-----------------------------------------------------------------------------------
        ReDim apiTicketInfo(1 To lvPreSell.ListItems.count)
        ReDim pszBusDate(1 To lvPreSell.ListItems.count)
        ReDim pnTicketCount(1 To lvPreSell.ListItems.count)
        ReDim pszEndStation(1 To lvPreSell.ListItems.count)
        ReDim pszOffTime(1 To lvPreSell.ListItems.count)
        ReDim pszBusID(1 To lvPreSell.ListItems.count)
        ReDim pszVehicleType(1 To lvPreSell.ListItems.count)
        ReDim pszCheckGate(1 To lvPreSell.ListItems.count)
        ReDim pbSaleChange(1 To lvPreSell.ListItems.count)
        ReDim pszTerminateName(1 To lvPreSell.ListItems.count)
        ReDim panInsurance(1 To lvPreSell.ListItems.count)
        
        For iCount = 1 To lvPreSell.ListItems.count
            With lvPreSell.ListItems(iCount)
                ReDim apiTicketInfo(iCount).aptPrintTicketInfo(1 To .SubItems(CC_TotalNum))
                For i = 1 To .SubItems(CC_TotalNum)
                    apiTicketInfo(iCount).aptPrintTicketInfo(i).nTicketType = srSellResult(iCount).aszTicketType(i)
                    apiTicketInfo(iCount).aptPrintTicketInfo(i).sgTicketPrice = srSellResult(iCount).asgTicketPrice(i)
                    apiTicketInfo(iCount).aptPrintTicketInfo(i).szSeatNo = srSellResult(iCount).aszSeatNo(i)
                    apiTicketInfo(iCount).aptPrintTicketInfo(i).szTicketNo = btTicketInfoParam(iCount).BuyTicketInfo(i).szTicketNo
                    
                    If srSellResult(iCount).aszTicketType(i) <> TP_FreeTicket Then
                        dbTotalNewPrice = dbTotalNewPrice + srSellResult(iCount).asgTicketPrice(i)
                    End If
                    
                Next i
                pszBusDate(iCount) = .SubItems(CC_BusDate)
                pnTicketCount(iCount) = .SubItems(CC_TotalNum)
                pszEndStation(iCount) = .SubItems(CC_EndStation)
                pszOffTime(iCount) = .SubItems(CC_OffTime)
                pszBusID(iCount) = .Text
                pszVehicleType(iCount) = .SubItems(CC_VehicleModel)
                pszCheckGate(iCount) = GetCheckName(.SubItems(CC_CheckGate))
                pszTerminateName(iCount) = .SubItems(CC_TerminateName)
                pbSaleChange(iCount) = True
                
                panInsurance(iCount) = Val(cboInsurance.Text)
            
            End With
        Next iCount
        
        lblChangeMsg.Caption = "正在打印车票"
        
        
        PrintTicket apiTicketInfo, pszBusDate, pnTicketCount, pszEndStation, pszOffTime, pszBusID, pszVehicleType, pszCheckGate, pbSaleChange, pszTerminateName, szSellStationName, panInsurance, m_vaCardInfo
        
        m_aszInsurce = CombInsurance(apiTicketInfo, szBusID, pnTicketCount, pszBusDate, szDesStationID, pszEndStation, pszOffTime, szSellStationID, szSellStationName, m_aszCardInfo)
        SaveInsurance m_aszInsurce
        
        lvPreSell.ListItems.Clear
    End If
    
   
    '******************************************************
    
'-------------------------------------------------------------------------------------
    
'如果票价变了
'--------------------------
'    Dim sgTemp As Double
'    If btTicketInfo.nTicketType <> TP_FreeTicket Then
'        sgTemp = srSellResult.asgTicketPrice(1)
'        If Abs(sgTemp - CDbl(lblTotalPrice.Caption)) > 0.01 Then
'            If optHalfTicket.Value Then
'                lvBus.SelectedItem.subitems(ID_HalfPrice) = sgTemp
'            Else
'                lvBus.SelectedItem.subitems(ID_FullPrice) = sgTemp
'            End If
'            DealPrice
'
'            ShowMsg "票价已经变了,请自己刷新车次信息,总票价应为:" & sgTemp & "元"
'        End If
'    End If
    
    dbRealTotalPrice = dbTotalNewPrice + dbTotalChangeFees - dbTotalOrgPrice
'    If Abs(dbRealTotalPrice - dbTotalPrice) > 0.01 Then
'        frmPriceInfo.m_sngTotalPrice = dbRealTotalPrice
'        frmPriceInfo.Show vbModal
'    End If
    
    ShowOldTicketInfo
    txtOldTktNum.SetFocus
    txtOldTktNum.Text = ""
    ShowNullTicketInfo
    lblChangeMsg.Caption = ""
    optFullTicket.Value = True
    cmdSell.Enabled = False
    Exit Sub
here:
    lblChangeMsg.Caption = ""
    
    
    ShowErrorMsg
End Sub

Public Sub ChangeSeatType()
    If cboSeatType.ListIndex = cboSeatType.ListCount - 1 Then
        cboSeatType.ListIndex = 0
    Else
        cboSeatType.ListIndex = cboSeatType.ListIndex + 1
    End If
    cboSeatType_Change
End Sub



Private Function TotalInsurace() As Double
'    '汇总保险费
'    Dim i As Integer
'    Dim nCount As Integer
'    If chkInsurance.Value = vbChecked Then
'        nCount = 0
'        For i = 1 To lvPreSell.ListItems.count
'            nCount = nCount + lvPreSell.ListItems(i).SubItems(IT_SumTicketNum)
'        Next i
'        nCount = nCount + 1
'        '保险费设为每张2元
'
'        TotalInsurace = nCount * 2
'    Else
'        TotalInsurace = 0
'    End If
End Function
