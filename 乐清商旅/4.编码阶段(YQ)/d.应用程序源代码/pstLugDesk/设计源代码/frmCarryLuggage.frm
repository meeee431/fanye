VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmCarryLuggage 
   BackColor       =   &H8000000C&
   Caption         =   "签发行包"
   ClientHeight    =   7680
   ClientLeft      =   3060
   ClientTop       =   2220
   ClientWidth     =   11445
   ControlBox      =   0   'False
   HelpContextID   =   7000050
   Icon            =   "frmCarryLuggage.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7680
   ScaleWidth      =   11445
   WindowState     =   2  'Maximized
   Begin VB.Frame fraOutLine 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6975
      Left            =   390
      TabIndex        =   0
      Top             =   120
      Width           =   10635
      Begin VB.ComboBox cboWorker 
         Height          =   300
         Left            =   1500
         Sorted          =   -1  'True
         TabIndex        =   17
         Top             =   2685
         Width           =   1575
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   3315
         Top             =   4455
      End
      Begin VB.ComboBox cboRatio 
         Height          =   300
         Left            =   9090
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   2595
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.ComboBox cboAcceptType 
         Height          =   300
         Left            =   1410
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   165
         Width           =   1155
      End
      Begin RTComctl3.CoolButton cmdSheetBus 
         Height          =   315
         Left            =   8670
         TabIndex        =   36
         Top             =   120
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "签发车次填写(&D)>>"
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
         MICON           =   "frmCarryLuggage.frx":076A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin RTComctl3.CoolButton cmdRefresh 
         Height          =   345
         Left            =   5220
         TabIndex        =   2
         Top             =   2670
         Width           =   1185
         _ExtentX        =   2090
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
         MICON           =   "frmCarryLuggage.frx":0786
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "签发车次信息"
         Height          =   1050
         Left            =   150
         TabIndex        =   10
         Top             =   555
         Width           =   6285
         Begin RTComctl3.TextButtonBox txtBusID 
            Height          =   300
            Left            =   975
            TabIndex        =   43
            Top             =   210
            Width           =   1305
            _ExtentX        =   2302
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
            Locked          =   -1  'True
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "发车时间:"
            Height          =   180
            Index           =   11
            Left            =   4440
            TabIndex        =   48
            Top             =   270
            Width           =   810
         End
         Begin VB.Label lblBusStratTime 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "8:30"
            Height          =   180
            Left            =   5445
            TabIndex        =   47
            Top             =   270
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "车次日期:"
            Height          =   180
            Index           =   10
            Left            =   2400
            TabIndex        =   15
            Top             =   270
            Width           =   810
         End
         Begin VB.Label lblBusDate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "2020-01-02"
            Height          =   180
            Left            =   3210
            TabIndex        =   14
            Top             =   270
            Width           =   900
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "车次(&B):"
            Height          =   180
            Index           =   18
            Left            =   210
            TabIndex        =   13
            Top             =   270
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "途经站点:"
            Height          =   180
            Index           =   17
            Left            =   210
            TabIndex        =   12
            Top             =   540
            Width           =   810
         End
         Begin VB.Label lblStations 
            BackStyle       =   0  'Transparent
            Caption         =   "中华人民共和国、中华人民共和国、中华人民共和国、中华人民共和国、中华人民共和国、中华人民共和国、中华人民共和国"
            ForeColor       =   &H00C00000&
            Height          =   360
            Left            =   1050
            TabIndex        =   11
            Top             =   540
            Width           =   5100
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "承运车辆信息"
         Height          =   945
         Left            =   150
         TabIndex        =   4
         Top             =   1710
         Width           =   6285
         Begin VB.ComboBox cboCarryVehicle 
            Height          =   300
            Left            =   1365
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   44
            ToolTipText     =   "F5键选择其他车辆"
            Top             =   225
            Width           =   1575
         End
         Begin RTComctl3.CoolButton cmdSearchVehicle 
            Height          =   315
            Left            =   3000
            TabIndex        =   46
            Top             =   540
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "全"
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
            MICON           =   "frmCarryLuggage.frx":07A2
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label lblBus 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "确认车辆(F5):"
            Height          =   180
            Left            =   210
            TabIndex        =   45
            Top             =   285
            Width           =   1170
         End
         Begin VB.Label lblProtocal 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "9：1"
            Height          =   180
            Left            =   4230
            TabIndex        =   31
            Top             =   570
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "拆算协议:"
            Height          =   180
            Index           =   12
            Left            =   3420
            TabIndex        =   30
            Top             =   570
            Width           =   810
         End
         Begin VB.Label lblSplitCompany 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "陈峰"
            Height          =   180
            Left            =   1020
            TabIndex        =   27
            Top             =   570
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "拆帐公司:"
            Height          =   180
            Index           =   2
            Left            =   210
            TabIndex        =   26
            Top             =   570
            Width           =   810
         End
         Begin VB.Label lblCompany 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "国务院"
            Height          =   180
            Left            =   5295
            TabIndex        =   9
            Top             =   285
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "参运公司:"
            Height          =   180
            Index           =   6
            Left            =   4455
            TabIndex        =   8
            Top             =   285
            Width           =   810
         End
         Begin VB.Label lblOwner 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "高兴"
            Height          =   180
            Left            =   3615
            TabIndex        =   7
            Top             =   285
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "车主:"
            Height          =   180
            Index           =   5
            Left            =   3165
            TabIndex        =   6
            Top             =   285
            Width           =   450
         End
      End
      Begin MSComctlLib.ListView lvLuggage 
         Height          =   3750
         Left            =   210
         TabIndex        =   1
         Top             =   3030
         Width           =   7605
         _ExtentX        =   13414
         _ExtentY        =   6615
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
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
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lvCarryed 
         Height          =   1920
         Left            =   6570
         TabIndex        =   5
         Top             =   615
         Width           =   3825
         _ExtentX        =   6747
         _ExtentY        =   3387
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   0   'False
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
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin RTComctl3.CoolButton cmdCarry 
         Height          =   705
         Left            =   7890
         TabIndex        =   19
         Top             =   6060
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   1244
         BTYPE           =   3
         TX              =   "签发(&P)"
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
         MICON           =   "frmCarryLuggage.frx":07BE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "装卸工(&W):"
         Height          =   180
         Index           =   0
         Left            =   360
         TabIndex        =   16
         Top             =   2760
         Width           =   900
      End
      Begin VB.Label lblSettlePrice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Left            =   9090
         TabIndex        =   42
         Top             =   2940
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label lblSettlePriceCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "应结运费:"
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
         Left            =   8010
         TabIndex        =   41
         Top             =   2940
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label lblRatio 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "应结费率:"
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
         Left            =   8010
         TabIndex        =   39
         Top             =   2595
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label lblAcceptType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "托运方式(&T):"
         Height          =   180
         Left            =   210
         TabIndex        =   37
         Top             =   225
         Width           =   1080
      End
      Begin VB.Label lblCalWeight 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2313"
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
         Left            =   9090
         TabIndex        =   3
         Top             =   4680
         Width           =   480
      End
      Begin VB.Line Line4 
         Index           =   1
         X1              =   7890
         X2              =   10350
         Y1              =   5955
         Y2              =   5955
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "托运总价:"
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
         Left            =   8010
         TabIndex        =   35
         Top             =   3405
         Width           =   1080
      End
      Begin VB.Line Line4 
         Index           =   0
         X1              =   7890
         X2              =   10350
         Y1              =   3285
         Y2              =   3285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "已装行包列表:"
         Height          =   180
         Index           =   14
         Left            =   6570
         TabIndex        =   34
         Top             =   330
         Width           =   1170
      End
      Begin VB.Label lblOverNumber 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "123123"
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
         Left            =   9090
         TabIndex        =   33
         Top             =   5535
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "超重件数:"
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
         Index           =   13
         Left            =   8010
         TabIndex        =   32
         Top             =   5535
         Width           =   1080
      End
      Begin VB.Label lblActWeight 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "12312"
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
         Left            =   9090
         TabIndex        =   29
         Top             =   5115
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "行包单数:"
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
         Index           =   3
         Left            =   8010
         TabIndex        =   28
         Top             =   3825
         Width           =   1080
      End
      Begin VB.Label lblBagNumber 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "12323"
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
         Left            =   9090
         TabIndex        =   25
         Top             =   4260
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "总实重:"
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
         Index           =   9
         Left            =   8010
         TabIndex        =   24
         Top             =   5115
         Width           =   840
      End
      Begin VB.Label lblBillNumber 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "123123"
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
         Left            =   9090
         TabIndex        =   23
         Top             =   3825
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "总计重:"
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
         Index           =   8
         Left            =   8010
         TabIndex        =   22
         Top             =   4680
         Width           =   840
      End
      Begin VB.Label lblTotalPrice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "234324"
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
         Left            =   9090
         TabIndex        =   21
         Top             =   3405
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "签发件数:"
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
         Index           =   7
         Left            =   8010
         TabIndex        =   20
         Top             =   4260
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "待受理行包列表(&L):"
         Height          =   180
         Index           =   1
         Left            =   3180
         TabIndex        =   18
         Top             =   2760
         Width           =   1620
      End
   End
End
Attribute VB_Name = "frmCarryLuggage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************
'Last Modify By: 陆勇庆  2005-8-16
'Last Modify In: 加入了新的装卸工字段，关联的所有进行了更改
'*******************************************************************************
Option Explicit

Private szTempStationName() As String               '途经站数组
Private mnLastSearchIndex As Integer            '选择车辆时用到的定位索引

Private m_aszVehicle() As String

Const cnAcceptNum = 4


'lvLuggage 列头
Const cnLuggageID = 0
Const cnEndStationName = 1
Const cnShipper = 2
Const cnPicket = 3
Const cnBusDate = 4
Const cnBusStartTime = 5
Const cnBusID = 6
Const cnBaggageNumber = 7
Const cnAcceptType = 8
Const cnMileage = 9
Const cnCalWeight = 10
Const cnFactWeight = 11
Const cnStartLabelID = 12
Const cnOverWeightNumber = 13
Const cnPriceTotal = 14
Const cnPriceItem1 = 15
Const cnHiddenChecked = 16
    
    
Private m_nSeconds As Integer '刷新的秒数


Private Sub FormClear()
  txtBusID.Text = ""
'  cboCarryVehicle.Clear
  lblBusStratTime.Caption = ""
  txtBusID.Text = ""
  lblBusDate.Caption = ""
'  lblStartTime.Caption = ""
  lblStations.Caption = ""
  
  lblOwner.Caption = ""
  lblCompany.Caption = ""
  lblSplitCompany.Caption = ""
  lblProtocal.Caption = ""
  lblTotalPrice.Caption = ""
  lblBillNumber.Caption = ""
  lblBagNumber.Caption = ""
  lblCalWeight.Caption = ""
  lblActWeight.Caption = ""
  lblOverNumber.Caption = ""
  lblSettlePrice.Caption = ""
  cboWorker.Text = ""
  cboRatio.ListIndex = 0
'  cmdCarry.Enabled = False
End Sub


'在lvLuggage,lvCarryed中移去签发完成的受理信息
Private Sub MoveSheet()
On Error GoTo ErrHandle
  Dim i As Integer
  Dim nCount As Integer
    nCount = 1
      Do While (lvLuggage.ListItems.Count > 0)
         If lvLuggage.ListItems(nCount).Checked = True Then
            lvLuggage.ListItems.Remove (nCount)
            nCount = nCount - 1
         End If
         nCount = nCount + 1
         If nCount > lvLuggage.ListItems.Count Then Exit Do
      Loop
      lvCarryed.ListItems.Clear
 Exit Sub
ErrHandle:
ShowErrorMsg
End Sub

Private Sub cboAcceptType_Change()
cmdRefresh_Click
End Sub

Private Sub cboAcceptType_Click()
cmdRefresh_Click
End Sub
'
''此处为指定车辆,发车时间,签发
'Private Sub cboBusStartTime_Change()
'    Dim i As Integer
'    Dim j As Integer
'    Dim nlen As Integer
'    Dim szTemp() As String
'    Dim szaTemp() As String
'    Dim Count As Integer
'On Error GoTo ErrHandle
'  If cboBusStartTime.Text <> "" Then
'         moCarrySheet.BusDate = Date
''        moCarrySheet.BusID = Trim(txtBusID.Text)
''        moCarrySheet.VehicleID = ResolveDisplay(cboCarryVehicle.Text)
'        moCarrySheet.RefreshBusInfo Trim(ResolveDisplayEx(cboCarryVehicle.Text)), CDate(CStr(Format(Date, "yy-mm-dd")) + " " + Trim(cboBusStartTime.Text) + ":00"), Trim(cboAcceptType.Text)
'        lblBusDate = moCarrySheet.BusDate
'        txtBusID.text = moCarrySheet.BusID
'        lblVehicle.Caption = moCarrySheet.VehicleLicense
'        lblOwner.Caption = moCarrySheet.BusOwnerName
'        lblCompany.Caption = moCarrySheet.CompanyName
'        lblSplitCompany.Caption = moCarrySheet.SplitCompanyName
'        lblProtocal.Caption = moCarrySheet.ProtocolName
'
'            szaTemp = moLugSvr.GetBusStationNames(Date, Trim(txtBusID.text))
'            Count = ArrayLength(szaTemp)
'            ReDim szTempStationName(1 To Count)
'            szTempStationName = szaTemp
'            lblStations = ""
'            For i = 1 To Count
'                lblStations.Caption = lblStations.Caption + " " + szTempStationName(i)
'            Next
'
''            nLen = ArrayLength(moLugSvr.GetBusRunVehicles(Trim(txtBusID.text)))
''
''            If nLen > 0 Then
''                ReDim szTemp(1 To nLen)
''                szTemp = moLugSvr.GetBusRunVehicles(Trim(txtBusID.Text))
''                cboCarryVehicle.Clear
''                For i = 1 To nLen
''                    cboCarryVehicle.AddItem MakeDisplayString(szTemp(i, 1), szTemp(i, 2))
''                Next i
''                 If nLen <> 0 Then cboCarryVehicle.ListIndex = 0
''            End If
'
''            lvCarryed.ListItems.Clear
'            moCarrySheet.BusDate = Date
'            moCarrySheet.BusID = Trim(txtBusID.text)
'            nlen = ArrayLength(moCarrySheet.GetBusPreLoadLuggage(GetLuggageTypeInt(Trim(cboAcceptType.Text))))
'            If nlen > 0 Then
'                ReDim szaTemp(1 To nlen)
'                szaTemp = moCarrySheet.GetBusPreLoadLuggage(GetLuggageTypeInt(Trim(cboAcceptType.Text)))
'                For i = 1 To lvLuggage.ListItems.Count
'                  For j = 1 To nlen
'                    If Trim(szaTemp(j)) = Trim(lvLuggage.ListItems(i).Text) Then
'                        lvLuggage.ListItems(i).Checked = True  '打勾填信息
'                        lvLuggage_ItemCheck lvLuggage.ListItems.Item(i) '向lvCarryed
'                    End If
'
'                  Next j
'                Next i
'            End If
'  End If
''  cboBusStartTime.Visible = False
'  Exit Sub
'ErrHandle:
'    ShowErrorMsg
'End Sub

'Private Sub cboBusStartTime_Click()
'cboBusStartTime_Change
'End Sub

'Private Sub cboBusStartTime_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then
'   cboBusStartTime_Change
'   lvLuggage.SetFocus
'End If
'End Sub

'Private Sub cboBusStartTime_LostFocus()
'lblStartTime.Caption = Trim(cboBusStartTime.Text)
'cboBusStartTime.Visible = False
'End Sub
Private Sub ClearBusInfo()
  txtBusID.Text = ""
'  cboCarryVehicle.Clear
  txtBusID.Text = ""
  lblBusDate.Caption = ""
  lblBusStratTime.Caption = ""
  lblStations.Caption = ""
'  lblVehicle.Caption = ""
  lblOwner.Caption = ""
  lblCompany.Caption = ""
  lblSplitCompany.Caption = ""
  lblProtocal.Caption = ""
'  cboBusStartTime.Clear
'  cboBusStartTime.Visible = False
End Sub
Private Sub cboCarryVehicle_Change()
Dim rsTemp As Recordset
Dim i As Integer

On Error GoTo ErrHandle
'    ClearBusInfo
'    If cboCarryVehicle.Text <> "" Then
'        moCarrySheet.BusDate = Date
''        moCarrySheet.BusID = Trim(txtBusID.Text)
''        moCarrySheet.VehicleID = ResolveDisplay(cboCarryVehicle.Text)
'       Set rsTemp = moCarrySheet.RefreshBusStartTime(cboCarryVehicle.Text)
'          cboBusStartTime.Clear
'       If rsTemp.RecordCount > 0 Then
'         For i = 1 To rsTemp.RecordCount
'          cboBusStartTime.AddItem Format(rsTemp!bus_start_time, "hh:mm")
'          rsTemp.MoveNext
'         Next i
'          cboBusStartTime.ListIndex = 0
'       End If
''        lblVehicle.Caption = moCarrySheet.VehicleLicense
''        lblOwner.Caption = moCarrySheet.BusOwnerName
''        lblCompany.Caption = moCarrySheet.CompanyName
''        lblSplitCompany.Caption = moCarrySheet.SplitCompanyName
''        lblProtocal.Caption = moCarrySheet.ProtocolName
'    End If
'    Set rsTemp = Nothing
    If cboCarryVehicle.Text = "" Then Exit Sub
'    cboBusStartTime.Visible = True
'    cboBusStartTime.SetFocus
'    lblVehicle.Caption = Trim(ResolveDisplayEx(cboCarryVehicle.Text))
    '同时修改车主,拆帐公司,参运公司
    Set rsTemp = moLugSvr.GetBusOtherInfo(Trim(ResolveDisplay(cboCarryVehicle.Text)))
    If rsTemp.RecordCount = 1 Then
        moCarrySheet.BusOwnerID = FormatDbValue(rsTemp!owner_id)
        moCarrySheet.BusOwnerName = FormatDbValue(rsTemp!owner_name)
        moCarrySheet.SplitCompanyID = FormatDbValue(rsTemp!split_company_id)
        moCarrySheet.SplitCompanyName = FormatDbValue(rsTemp!transport_company_short_name)
        moCarrySheet.CompanyId = FormatDbValue(rsTemp!transport_company_id)
        moCarrySheet.CompanyName = FormatDbValue(rsTemp!transport_company_short_name)
        lblOwner.Caption = moCarrySheet.BusOwnerName
        lblCompany.Caption = moCarrySheet.CompanyName
        lblSplitCompany.Caption = moCarrySheet.SplitCompanyName
    End If
Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub cboCarryVehicle_Click()
  cboCarryVehicle_Change
End Sub

Private Sub cboCarryVehicle_GotFocus()
lblBus.ForeColor = clActiveColor
End Sub

Private Sub cboCarryVehicle_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   cboCarryVehicle_Change
'   If txtBusID.text <> "" Then
'     cboBusStartTime.Visible = True
'     cboBusStartTime.SetFocus
'   Else
'     cboBusStartTime.Visible = False
'     lvLuggage.SetFocus
'   End If
End If
End Sub

Private Sub cboCarryVehicle_LostFocus()
lblBus.ForeColor = 0
If cboCarryVehicle.Text <> "" Then
'   cboBusStartTime.Visible = True
'   cboBusStartTime.SetFocus
End If
End Sub

Private Sub ClearCarry()
  
  If txtBusID.Text = "" Then
     
  End If
End Sub

Private Sub cboRatio_Change()
    SumAccept
End Sub

Private Sub cboRatio_Click()
    SumAccept
End Sub



Private Sub cmdAllVehicle_Click()
FillAllVehicle
End Sub

Public Sub FillAllVehicle()
Dim i As Integer
Dim nLen As Integer
  '得到所有车辆信息
  cboCarryVehicle.Clear
  cboCarryVehicle.AddItem "空"
  nLen = ArrayLength(m_aszVehicle)
  If nLen > 0 Then
   For i = 1 To nLen
     cboCarryVehicle.AddItem MakeDisplayString(m_aszVehicle(i, 1), m_aszVehicle(i, 2))
    
   Next i
  End If
End Sub

Private Sub cmdCarry_Click()
    Dim nCount As Integer
    Dim i As Integer
    Dim dbPrintSettlePrice As Double
    On Error GoTo ErrHandle
    If txtBusID.Text = "" Then
        MsgBox "请选择签发车次!", vbInformation, Me.Caption
        Exit Sub
    End If
    
    If lblBusStratTime.Caption = "" Then
        MsgBox "无效车次！不能签发！", vbInformation, Me.Caption
        Exit Sub
    End If
    
    
    
    If MsgBox("是否确认签发?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then Exit Sub
    
    If Trim(cboCarryVehicle.Text) = "空" Then
        moCarrySheet.VehicleID = ""
        moCarrySheet.VehicleLicense = ""
    Else
        moCarrySheet.VehicleID = ResolveDisplay(cboCarryVehicle.Text)
        moCarrySheet.VehicleLicense = ResolveDisplayEx(cboCarryVehicle.Text)
    End If
    
    nCount = 0
    dbPrintSettlePrice = 0
    moCarrySheet.ClearLuggage
    '每4张受理单作一张签发单打印
    For i = 1 To lvLuggage.ListItems.Count
        If lvLuggage.ListItems(i).Checked = True Then
            moCarrySheet.AddLuggage Trim(lvLuggage.ListItems(i).Text), GetLuggageTypeInt(Trim(cboAcceptType.Text))
            nCount = nCount + 1
            dbPrintSettlePrice = dbPrintSettlePrice + lvLuggage.ListItems(i).SubItems(cnPriceItem1) '应结运费
            If nCount Mod cnAcceptNum = 0 Then
                If g_szCarrySheetID <> Trim(mdiMain.lblSheetNo) Then
                    g_szCarrySheetID = Trim(mdiMain.lblSheetNo)
                End If
                moCarrySheet.SheetID = g_szCarrySheetID
                moCarrySheet.PrintSettleRatio = cboRatio.Text
                moCarrySheet.PrintSettlePrice = Val(Format(dbPrintSettlePrice * cboRatio.Text / 100, "0.00")) ')dbPrintSettlePrice)
                dbPrintSettlePrice = 0
                moLugSvr.CarryLuggage moCarrySheet, GetLuggageTypeInt(Trim(cboAcceptType.Text))
                moCarrySheet.BusID = txtBusID.Text
                moCarrySheet.BusDate = lblBusDate.Caption
                
                moCarrySheet.MoveWorker = cboWorker.Text
           
                
                If Trim(cboCarryVehicle.Text) = "空" Then
                    moCarrySheet.VehicleID = ""
                    moCarrySheet.VehicleLicense = ""
                Else
                    moCarrySheet.VehicleID = ResolveDisplay(cboCarryVehicle.Text)
                    moCarrySheet.VehicleLicense = ResolveDisplayEx(cboCarryVehicle.Text)
                End If
                '以下打印签发单
                ShowSBInfo "正在打印签发单..."
                PrintCarrySheet moCarrySheet
                ShowSBInfo ""
                moCarrySheet.ClearLuggage
'                lblBus_ID.ForeColor = clActiveColor
'                moCarrySheet.AddNew
                IncSheetID
                nCount = 0
            End If
        End If
    Next i
    If nCount < cnAcceptNum And nCount > 0 Then
        '
        If g_szCarrySheetID <> Trim(mdiMain.lblSheetNo) Then
            g_szCarrySheetID = Trim(mdiMain.lblSheetNo)
        End If
        moCarrySheet.SheetID = g_szCarrySheetID
        moCarrySheet.BusID = txtBusID.Text
        moCarrySheet.BusDate = lblBusDate.Caption
        
        
        
        moCarrySheet.MoveWorker = cboWorker.Text
        moCarrySheet.PrintSettleRatio = cboRatio.Text
        moCarrySheet.PrintSettlePrice = Val(Format(dbPrintSettlePrice * cboRatio.Text / 100, "0.00")) 'Val(lblSettlePrice.Caption)
        moLugSvr.CarryLuggage moCarrySheet, GetLuggageTypeInt(Trim(cboAcceptType.Text))
        '以下打印签发单
        ShowSBInfo "正在打印签发单..."
        PrintCarrySheet moCarrySheet
        ShowSBInfo ""
        moCarrySheet.ClearLuggage
'        lblBus_ID.ForeColor = clActiveColor
'        moCarrySheet.AddNew
        IncSheetID
        nCount = 0
    
        '2005-7-13 lyq changed
        RefreshLuggage
    
    End If
    
    '记忆新加入的装卸工名字
    For i = 0 To cboWorker.ListCount - 1
        If cboWorker.Text = cboWorker.List(i) Then Exit For
    Next i
    If i = cboWorker.ListCount And Trim(cboWorker.Text) <> "" Then
        cboWorker.AddItem cboWorker.Text
    End If
    
    
    FormClear
    MoveSheet
    g_szCarrySheetID = Trim(mdiMain.lblSheetNo)
    moCarrySheet.SheetID = g_szCarrySheetID
    '界面处理
    FillAllVehicle
    
'    cboBusStartTime.Clear
'    cboBusStartTime.Visible = False
    
    
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub cmdRefresh_Click()
    RefreshLuggage
End Sub

Private Sub cmdSearchVehicle_Click()
    cmdAllVehicle_Click
    frmSearchVechile.mFormNum = 0
    frmSearchVechile.StartSearchIndex = mnLastSearchIndex
    frmSearchVechile.Show vbModal
    mnLastSearchIndex = frmCarryLuggage.cboCarryVehicle.ListIndex
End Sub

Private Sub cmdSheetBus_Click()
    frmUpdateSheet.Show vbModal
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF12 Then
   If cboCarryVehicle.Enabled Then
'    cboBusStartTime.Visible = False
    frmSearchVechile.mFormNum = 0
    frmSearchVechile.StartSearchIndex = mnLastSearchIndex
    frmSearchVechile.Show vbModal
    mnLastSearchIndex = cboCarryVehicle.ListIndex
    
   End If
ElseIf KeyCode = vbKeyF3 Then '托运方式切换
    If cboAcceptType.ListIndex = 0 Then
       cboAcceptType.ListIndex = 1
    Else
       cboAcceptType.ListIndex = 0
    End If
    cboAcceptType_Change
ElseIf KeyCode = vbKeyF5 Then
    '选择车辆
    Dim oCommDialog As New CommDialog
    Dim aszTemp() As String
    Dim i As Integer
    Dim bFound As Boolean
    
    oCommDialog.Init m_oAUser
    aszTemp = oCommDialog.SelectVehicleEX()
    If ArrayLength(aszTemp) > 0 Then
        bFound = False
        For i = 0 To cboCarryVehicle.ListCount - 1
            If ResolveDisplay(cboCarryVehicle.List(i)) = aszTemp(1, 1) Then
                bFound = True
                Exit For
            End If
        Next i
        If Not bFound Then
            cboCarryVehicle.AddItem MakeDisplayString(aszTemp(1, 1), aszTemp(1, 2))
'            cboCarryVehicle.ListIndex = cboCarryVehicle.ListCount - 1
            cboCarryVehicle.Text = MakeDisplayString(aszTemp(1, 1), aszTemp(1, 2))
        Else
            cboCarryVehicle.ListIndex = i
        End If
    End If
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If Me.ActiveControl Is lvLuggage And KeyAscii = 13 Then
       cmdCarry.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Dim szaTemp() As String
    Dim i As Integer
    Dim nLen As Integer
    AlignHeadWidth Me.name, lvLuggage
    FillLvLuggageColumnHead
    With cboRatio
        cboRatio.AddItem "100"
        cboRatio.AddItem "80"
        cboRatio.AddItem "65"
        cboRatio.AddItem "60"
        cboRatio.AddItem "40"
        cboRatio.AddItem "35"
        cboRatio.AddItem "30"
        cboRatio.AddItem "20"
        cboRatio.AddItem "10"
        cboRatio.AddItem "5"
        cboRatio.AddItem "0"
        cboRatio.ListIndex = 0
    End With
    LvCarryHead
    With cboAcceptType
        .AddItem szAcceptTypeGeneral
        .AddItem szAcceptTypeMan
        .Text = szAcceptTypeGeneral
    End With
    '得到所有的车辆
    m_oBase.Init m_oAUser
    m_aszVehicle = m_oBase.GetVehicle()
    
    '得到所有车辆信息
    FillAllVehicle
    
    moCarrySheet.Init m_oAUser
    FormClear
    RefreshLuggage
'    cboBusStartTime.Visible = False
    moCarrySheet.AddNew
    moCarrySheet.SheetID = g_szCarrySheetID
'    lblVehicle.ForeColor = vbBlue
    lblStations.ForeColor = vbBlue
    
    '装卸工信息
    LoadWorkerInfo
End Sub

'加载装卸工信息
Private Sub LoadWorkerInfo()
On Error GoTo ErrHandle
    Dim oLuggageParam As New LuggageParam
    Dim aszTmp() As String
    Dim i As Integer
    
    aszTmp = oLuggageParam.ListBaseDefine(62)
    cboWorker.Clear
    
    For i = 1 To ArrayLength(aszTmp)
        cboWorker.AddItem aszTmp(i, 3)
    Next i
    
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Public Sub RefreshLuggage()
    On Error GoTo ErrHandle
    Dim rsTemp As Recordset
    Dim lvl As ListItem
    Dim i As Integer
    Dim j As Integer
    Dim szaTemp() As String
    Dim nLen As Integer
    Dim szStartDate As Date
    Dim szEndDate As Date
    
    '定义未签发的操作时间
    szStartDate = cszEmptyDateStr ' CDate(CStr(Date) + " 00:00:01")
    szEndDate = cszForeverDateStr ' CDate(CStr(Date) + " 23:59:59")
    
    lvLuggage.ListItems.Clear
    lvCarryed.ListItems.Clear
    
    '得到当天未受理的所有行包受理单信息
    Set rsTemp = moLugSvr.GetAcceptSheetRS(szStartDate, szEndDate, 0, , , , , True)  '时间叁数,0表示状态为未签发
    If rsTemp.RecordCount > 0 Then
        For i = 1 To rsTemp.RecordCount
            '         If GetLuggageTypeInt(Trim(cboAcceptType.Text)) = FormatDbValue(rsTemp!accept_type) Then
            Set lvl = lvLuggage.ListItems.Add(, , FormatDbValue(rsTemp!luggage_id)) '受理单号
            lvl.SubItems(cnEndStationName) = FormatDbValue(rsTemp!des_station_name) '到达站
            lvl.SubItems(cnShipper) = FormatDbValue(rsTemp!Shipper) '托运人
            lvl.SubItems(cnPicket) = FormatDbValue(rsTemp!Picker) '收件人
            lvl.SubItems(cnBusDate) = FormatDbValue(rsTemp!bus_date) '车次日期
            lvl.SubItems(cnBusStartTime) = Format(rsTemp!bus_start_time, "HH:mm") '发车时间
            lvl.SubItems(cnBusID) = FormatDbValue(rsTemp!bus_id) '车次
            lvl.SubItems(cnBaggageNumber) = FormatDbValue(rsTemp!baggage_number) '件数
            lvl.SubItems(cnAcceptType) = GetLuggageTypeString(FormatDbValue(rsTemp!accept_type)) '托运方式
            lvl.SubItems(cnMileage) = FormatDbValue(rsTemp!Mileage) '里程
            lvl.SubItems(cnCalWeight) = FormatDbValue(rsTemp!cal_weight) '计重
            lvl.SubItems(cnFactWeight) = FormatDbValue(rsTemp!fact_weight) '实重
            lvl.SubItems(cnStartLabelID) = FormatDbValue(rsTemp!start_label_id) '标签号
            lvl.SubItems(cnOverWeightNumber) = FormatDbValue(rsTemp!over_weight_number) '超重件数
            lvl.SubItems(cnPriceTotal) = FormatDbValue(rsTemp!price_total) '托运费
            lvl.SubItems(cnPriceItem1) = FormatDbValue(rsTemp!price_item_1) '运费
            
            
            '         End If
            rsTemp.MoveNext
        Next i
    End If
    Set rsTemp = Nothing
    '得到已指定该车次签发的行包受理单
    Dim bChecked As Boolean
    If txtBusID.Text <> "" Then
        moCarrySheet.BusDate = Date
        moCarrySheet.BusID = Trim(txtBusID.Text)
        szaTemp = moCarrySheet.GetBusPreLoadLuggage(-1)
        nLen = ArrayLength(szaTemp)
        If nLen <> 0 Then
            ReDim szaTemp(1 To nLen)
            For i = 1 To lvLuggage.ListItems.Count
                For j = 1 To nLen
                    If Trim(szaTemp(j)) = Trim(lvLuggage.ListItems(i).Text) Then
''                        lvLuggage.ListItems(i).Checked = True  '打勾
                        lvLuggage_ItemCheck lvLuggage.ListItems.Item(i)
                    
                    End If
                Next j
            Next i
        End If
    End If
    
        
    SumAccept
    ClearBusInfo
    '得到所有车辆信息
    FillAllVehicle
    
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub Form_Resize()
    If mdiMain.ActiveForm Is Me Then
        If Not Me.WindowState = vbMaximized Then Me.WindowState = vbMaximized
        fraOutLine.Left = (Me.ScaleWidth - fraOutLine.Width) / 2
        fraOutLine.Top = (Me.ScaleHeight - fraOutLine.Height) / 2
    End If
End Sub

Private Sub Form_Activate()
    SetSheetNoLabel False, g_szCarrySheetID
End Sub

Private Sub Form_Deactivate()
    HideSheetNoLabel
End Sub
Private Sub Form_Unload(Cancel As Integer)
    SaveHeadWidth Me.name, lvLuggage
    HideSheetNoLabel
End Sub

Private Sub SumAccept()
    Dim i As Integer
    Dim dbTotalPrice As Double
    Dim dbBillNum As Integer
    Dim dbBagNum As Integer
    Dim dbCal As Double
    Dim dbAct As Double
    Dim dbOver As Integer
    Dim dbSettlePrice As Double
    Dim dbBasePrice As Double
    For i = 1 To lvLuggage.ListItems.Count
        If lvLuggage.ListItems(i).Checked = True Then
            dbTotalPrice = dbTotalPrice + CDbl(lvLuggage.ListItems(i).SubItems(cnPriceTotal))
            
            dbBillNum = dbBillNum + 1
            dbBasePrice = dbBasePrice + Val(lvLuggage.ListItems(i).SubItems(cnPriceItem1))
            
            dbBagNum = dbBagNum + CInt(lvLuggage.ListItems(i).SubItems(cnBaggageNumber))
            dbCal = dbCal + CDbl(lvLuggage.ListItems(i).SubItems(cnCalWeight))
            dbAct = dbAct + CDbl(lvLuggage.ListItems(i).SubItems(cnFactWeight))
            dbOver = dbOver + CInt(lvLuggage.ListItems(i).SubItems(cnOverWeightNumber))
            If dbBillNum Mod cnAcceptNum = 0 Then
                '精确到元
                '每4单精确一次,怕按总的汇总,累加再四舍五入与每4单四舍五入的金额会不对.
                dbSettlePrice = dbSettlePrice + Format(dbBasePrice * cboRatio.Text / 100, "0.00")
                dbBasePrice = 0
            End If
        End If
    Next i
    If dbBillNum Mod cnAcceptNum <> 0 Then
        dbSettlePrice = dbSettlePrice + Format(dbBasePrice * cboRatio.Text / 100, "0.00")
    End If
    
    lblTotalPrice.Caption = CStr(dbTotalPrice)
    lblBillNumber.Caption = CStr(dbBillNum)
    lblBagNumber.Caption = CStr(dbBagNum)
    lblCalWeight.Caption = CStr(dbCal)
    lblActWeight.Caption = CStr(dbAct)
    lblOverNumber.Caption = CStr(dbOver)
    lblSettlePrice.Caption = CStr(dbSettlePrice)
    If dbBillNum > 0 And txtBusID.Text <> "" Then
        cmdCarry.Enabled = True
    End If
End Sub
Private Sub LvCarryHead()
  lvCarryed.ColumnHeaders.Clear
    '添加ListView列头
    With lvCarryed.ColumnHeaders
        .Add , , "标签号", 1540
        .Add , , "行包名称", 1500
        .Add , , "计重", 800
        .Add , , "实重", 800
        .Add , , "件数", 800
        .Add , , "类型", 800
        .Add , , "受理单号", 0
   End With
End Sub

Private Sub lvLuggage_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    SortListView lvLuggage, ColumnHeader.Index
End Sub

Private Sub lvLuggage_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo ErrHandle
    Dim i As Integer
    Dim tLugItem() As TLuggageItemInfo
    Dim nLen As Integer
    Dim j As Integer
    Dim lvc As ListItem
    Dim nCount As Integer
    Dim mHave As Boolean
    Dim rsTemp As Recordset
    Dim szaTemp() As String
    
    If Item.Checked = True Then
        Item.Selected = True
        '判断指定车次的途经站是否包括受理信息中的到达站点
        '        mHave = False
        '        For i = 1 To ArrayLength(szTempStationName)
        '            If szTempStationName(i) = Trim(Item.SubItems(2)) Then
        '               mHave = True
        '            End If
        '        Next i
        '        If mHave = False Then
        '          If MsgBox("指定车次的途经站不包括此受理单的到达站," & Chr(13) & Chr(10) & "            您是否要强行签发?", vbInformation + vbYesNo, Me.Caption) = vbNo Then
        '             Item.Checked = False
        '             Exit Sub
        '          End If
        '        End If
        
        '新增========================================================
        '不同车次的选择
        
        For i = 1 To lvLuggage.ListItems.Count
            If lvLuggage.ListItems.Item(i).Checked = True Then
                If Trim(Item.SubItems(cnBusID)) <> Trim(lvLuggage.ListItems.Item(i).SubItems(cnBusID)) Then
                    lvLuggage.ListItems.Item(i).Checked = False
                     '清除该受理单对应的元素
                   For nCount = lvCarryed.ListItems.Count To 1 Step -1
                        If Trim(lvLuggage.ListItems.Item(i).Text) = Trim(lvCarryed.ListItems(nCount).SubItems(6)) Then
                            lvCarryed.ListItems.Remove (nCount)
                        End If
                    Next nCount
                 End If
            End If
        Next i
        
        '同一车次选中操作
        Dim bChecked As Boolean
        Dim k As Integer
''        lvCarryed.ListItems.Clear
''        For k = 1 To lvLuggage.ListItems.Count
''            If Trim(Item.SubItems(cnBusID)) = Trim(lvLuggage.ListItems.Item(k).SubItems(cnBusID)) Then
''''                lvLuggage.ListItems.Item(k).Checked = True
''                          '向lvCarryed 填受理单信息
''
''              LvCarryHead
''              moAcceptSheet.Identify Trim(lvLuggage.ListItems.Item(k).Text)
''              Set lvc = lvCarryed.ListItems.Add(, , Trim(moAcceptSheet.StartLabelID) + "-" + Trim(moAcceptSheet.EndLabelID))
''              lvc.SubItems(1) = moAcceptSheet.LuggageName
''              lvc.SubItems(2) = moAcceptSheet.CalWeight
''              lvc.SubItems(3) = moAcceptSheet.ActWeight
''              lvc.SubItems(4) = moAcceptSheet.Number
''              lvc.SubItems(5) = moAcceptSheet.AcceptType
''              lvc.SubItems(6) = moAcceptSheet.SheetID
''            End If
''        Next k

    
        '    Dim rsTemp1 As Recordset
         moCarrySheet.RefreshBusInfoEX Trim(Item.SubItems(cnBusID)), CDate(CStr(Format(Item.SubItems(cnBusDate), "yy-MM-dd")) + " " + Trim(Item.SubItems(cnBusStartTime)) + ":00"), Trim(cboAcceptType.Text)
        '    If rsTemp1.RecordCount = 0 Then Exit Sub
        lblBusDate = moCarrySheet.BusDate
        txtBusID.Text = moCarrySheet.BusID
        lblBusStratTime.Caption = Format(moCarrySheet.BusStartOffTime, "hh:mm")
'        lblVehicle.Caption = moCarrySheet.VehicleLicense
        lblOwner.Caption = moCarrySheet.BusOwnerName
        lblCompany.Caption = moCarrySheet.CompanyName
        lblSplitCompany.Caption = moCarrySheet.SplitCompanyName
        lblProtocal.Caption = moCarrySheet.ProtocolName
        
        Dim szaTemp1() As String
        Dim count1 As Integer
        Dim szTempStationName1() As String
        szaTemp1 = moLugSvr.GetBusStationNames(Date, Trim(txtBusID.Text))
        count1 = ArrayLength(szaTemp1)
        ReDim szTempStationName1(1 To count1)
        szTempStationName1 = szaTemp1
        lblStations = ""
        For i = 1 To count1
            lblStations.Caption = lblStations.Caption + " " + szTempStationName1(i)
        Next
        '显示些车次的所有车辆
        Dim rsTemp1 As Recordset
        Set rsTemp1 = moLugSvr.GetBusVehicle(Trim(Item.SubItems(cnBusID)))
        If rsTemp1.RecordCount > 0 Then
            cboCarryVehicle.Clear
            cboCarryVehicle.AddItem "空"
            For i = 1 To rsTemp1.RecordCount
                cboCarryVehicle.AddItem MakeDisplayString(FormatDbValue(rsTemp1!vehicle_id), FormatDbValue(rsTemp1!license_tag_no))
                rsTemp1.MoveNext
            Next i
        End If
        If cboCarryVehicle.ListCount > 0 Then
            Dim ListIndex1 As Integer
            ListIndex1 = 0
'            cboCarryVehicle.ListIndex = 0
            For i = 0 To cboCarryVehicle.ListCount
'                If Trim(ResolveDisplayEx(cboCarryVehicle.List(i))) = Trim(lblVehicle.Caption) Then

                If Trim(ResolveDisplayEx(cboCarryVehicle.List(i))) = Trim(moCarrySheet.VehicleLicense) Then
                    
                    cboCarryVehicle.ListIndex = i
                End If
'                ListIndex1 = ListIndex1 + 1
            Next i
        End If
    
    
        moAcceptSheet.SheetID = Trim(Item.Text)
        nLen = ArrayLength(moAcceptSheet.GetLugItemDetail)
        If nLen > 0 Then
            LvCarryHead
            ReDim tLugItem(1 To nLen)
            tLugItem = moAcceptSheet.GetLugItemDetail
            For j = 1 To nLen
                Set lvc = lvCarryed.ListItems.Add(, , tLugItem(j).LabelID)    '向lvCarryed 填行包明细
                lvc.SubItems(1) = tLugItem(j).LuggageName
                lvc.SubItems(2) = tLugItem(j).CalWeight
                lvc.SubItems(3) = tLugItem(j).ActWeight
                lvc.SubItems(4) = tLugItem(j).Number
                lvc.SubItems(5) = tLugItem(j).LuggageTypeName
                lvc.SubItems(6) = tLugItem(j).LuggageID
            Next j
        
        Else
''            '向lvCarryed 填受理单信息
''            If lvCarryed.ListItems.Count > 0 Then
''                For k = 1 To lvCarryed.ListItems.Count
''                    If Trim(lvCarryed.ListItems.Item(k).SubItems(6)) = Trim(Item.Text) Then
''                        GoTo Skip:
''                    End If
''                Next k
''            End If
            LvCarryHead
            moAcceptSheet.Identify Trim(Item.Text)
            Set lvc = lvCarryed.ListItems.Add(, , Trim(moAcceptSheet.StartLabelID) + "-" + Trim(moAcceptSheet.EndLabelID))
            lvc.SubItems(1) = moAcceptSheet.LuggageName
            lvc.SubItems(2) = moAcceptSheet.CalWeight
            lvc.SubItems(3) = moAcceptSheet.ActWeight
            lvc.SubItems(4) = moAcceptSheet.Number
            lvc.SubItems(5) = moAcceptSheet.AcceptType
            lvc.SubItems(6) = moAcceptSheet.SheetID
        End If
Skip:

  
    Else
        '清除该受理单对应的元素
        nCount = 1
        Do While (lvCarryed.ListItems.Count > 0)
            If Trim(Item.Text) = Trim(lvCarryed.ListItems(nCount).SubItems(6)) Then
                lvCarryed.ListItems.Remove (nCount)
                nCount = nCount - 1
            End If
            nCount = nCount + 1
            If nCount > lvCarryed.ListItems.Count Then Exit Do
        Loop
        
    End If
    SumAccept
    If lvCarryed.ListItems.Count = 0 Then
        ClearBusInfo
        FillAllVehicle
    End If
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub


Private Sub Timer1_Timer()
    m_nSeconds = m_nSeconds + 1
    If m_nSeconds >= 120 Then
        '每隔2分钟刷新一次
        RefreshLuggage
        m_nSeconds = 0
    End If
End Sub

Private Sub txtBusID_Click()
    Dim oCommDialog As New CommDialog
    Dim aszTemp() As String
    Dim dtBusDate As Date
    Dim dtBusStartTime As Date
    Dim nCount As Integer
    Dim i As Integer
    
    ShowMsg "此处车次更该只能更改为与原来相同线路的车次，若要更改其它线路的车次请点击[签发车次填写]按钮进行修改！"
    
    oCommDialog.Init m_oAUser
    dtBusDate = IIf(lblBusDate.Caption = "", m_oParam.NowDate, lblBusDate.Caption)
    aszTemp = oCommDialog.SelectREBus(dtBusDate, True, False)
    
    If ArrayLength(aszTemp) > 0 Then
        txtBusID.Text = aszTemp(1, 1)
        lblBusDate.Caption = CDate(Format(aszTemp(1, 2), "YYYY-MM-dd"))
        moCarrySheet.BusDate = CDate(lblBusDate.Caption)
        moCarrySheet.BusID = aszTemp(1, 1)
        
        '发车时间
        dtBusStartTime = GetAllotStationBusStartTime(txtBusID.Text, CDate(lblBusDate.Caption))
        If dtBusStartTime <> cdtEmptyDate Then
            lblBusStratTime.Caption = Format(dtBusStartTime, "HH:mm")
        Else
            lblBusStratTime.Caption = ""
            ShowMsg "无效车次或车次已停班！"
            Exit Sub
        End If
        
        
        '将该车次的默认的车辆设置到车辆框中
        aszTemp = moLugSvr.GetEnvBusRunVehicles(CDate(lblBusDate.Caption), txtBusID.Text)
        
        nCount = ArrayLength(aszTemp)
        
        If nCount > 0 Then
        
            moCarrySheet.VehicleLicense = aszTemp(1, 1)
            
            
            cboCarryVehicle.Clear
            cboCarryVehicle.AddItem "空"
            For i = 1 To nCount
                cboCarryVehicle.AddItem MakeDisplayString(aszTemp(i, 1), aszTemp(i, 2))
                
            Next i
        End If
        
        If cboCarryVehicle.ListCount > 0 Then
            Dim ListIndex1 As Integer
            ListIndex1 = 0
'            cboCarryVehicle.ListIndex = 0
            For i = 0 To cboCarryVehicle.ListCount
'                If Trim(ResolveDisplayEx(cboCarryVehicle.List(i))) = Trim(lblVehicle.Caption) Then

                If Trim(ResolveDisplay(cboCarryVehicle.List(i))) = Trim(moCarrySheet.VehicleLicense) Then
                    
                    cboCarryVehicle.ListIndex = i
                End If
'                ListIndex1 = ListIndex1 + 1
            Next i
        End If
        
        
        cboCarryVehicle_Change
        
'        RefreshBusInfo
        
        moCarrySheet.RefreshBusInfoEX txtBusID.Text, CDate(lblBusDate.Caption & " " & lblBusStratTime.Caption), Trim(cboAcceptType.Text)
    End If
    
End Sub


'Private Sub RefreshBusInfo(pszBusID As String, pdtBusDate As Date)
'
'End Sub





Private Sub txtBusID_GotFocus()
'lblBus_ID.ForeColor = clActiveColor
End Sub

'此处为输入指定车次,签发
'Private Sub txtBusID_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo ErrHandle
'    Dim i As Integer
'    Dim j As Integer
'    Dim nLen As Integer
'    Dim szTemp() As String
'    Dim szaTemp() As String
'    Dim Count As Integer
'    '---fl--
''    Dim szTempStationName() As String
'    '---fl--
'    If KeyCode = vbKeyReturn Then
'        If txtBusID.Text <> "" Then
'            moCarrySheet.BusDate = ToDBDate(Date)
'            moCarrySheet.BusID = Trim(txtBusID.Text)
'            moCarrySheet.RefreshBusVehicle
'            txtBusID.text = moCarrySheet.BusID
'            lblBusDate.Caption = CStr(moCarrySheet.BusDate)
'            lblStartTime.Caption = CStr(Format(moCarrySheet.BusStartOffTime, "hh:mm"))
'
'            '=====fl
'            szaTemp = moLugSvr.GetBusStationNames(Date, Trim(txtBusID.Text))
'            Count = ArrayLength(szaTemp)
'            ReDim szTempStationName(1 To Count)
'            szTempStationName = szaTemp
'            lblStations = ""
'            For i = 1 To Count
'                lblStations.Caption = lblStations.Caption + " " + szTempStationName(i)
'            Next
'
'            nLen = ArrayLength(moLugSvr.GetBusRunVehicles(Trim(txtBusID.Text)))
'
'            If nLen > 0 Then
'                ReDim szTemp(1 To nLen)
'                szTemp = moLugSvr.GetBusRunVehicles(Trim(txtBusID.Text))
'                cboCarryVehicle.Clear
'                For i = 1 To nLen
'                    cboCarryVehicle.AddItem MakeDisplayString(szTemp(i, 1), szTemp(i, 2))
'                Next i
'                 If nLen <> 0 Then cboCarryVehicle.ListIndex = 0
'            End If
'
'            lvCarryed.ListItems.Clear
'            moCarrySheet.BusDate = Date
'            moCarrySheet.BusID = Trim(txtBusID.Text)
'            nLen = ArrayLength(moCarrySheet.GetBusPreLoadLuggage(GetLuggageTypeInt(Trim(cboAcceptType.Text))))
'            If nLen > 0 Then
'                ReDim szaTemp(1 To nLen)
'                szaTemp = moCarrySheet.GetBusPreLoadLuggage(GetLuggageTypeInt(Trim(cboAcceptType.Text)))
'                For i = 1 To lvLuggage.ListItems.Count
'                  For j = 1 To nLen
'                    If Trim(szaTemp(j)) = Trim(lvLuggage.ListItems(i).Text) Then
'                        lvLuggage.ListItems(i).Checked = True  '打勾填信息
'                        lvLuggage_ItemCheck lvLuggage.ListItems.Item(i) '向lvCarryed
'                    End If
'
'                  Next j
'                Next i
'            End If
'
'            cboCarryVehicle.SetFocus
'        End If
'
'    End If
'    Exit Sub
'ErrHandle:
'    ShowErrorMsg
'End Sub

Private Sub txtBusID_LostFocus()
 Dim i As Integer
' lblBus_ID.ForeColor = 0
 If lvLuggage.ListItems.Count > 0 Then
  For i = 1 To lvLuggage.ListItems.Count
    If lvLuggage.ListItems(i).Checked = True Then
    cmdCarry.Enabled = True
    End If
  Next i
 End If
 
End Sub

'车辆查询
'Private Sub txtVehicle_ButtonClick()
'On Error GoTo ErrHandle
'    Dim oShell As New STShell.CommDialog
'    Dim aszTemp() As String
'    oShell.Init m_oAUser
'    aszTemp = oShell.SelectVehicleEX()
'    Set oShell = Nothing
'    If ArrayLength(aszTemp) = 0 Then Exit Sub
'    txtVehicle.Text = aszTemp(1, 1) & "[" & Trim(aszTemp(1, 2)) & "]"
'
'Exit Sub
'ErrHandle:
'ShowErrorMsg
'End Sub
Private Sub FillLvLuggageColumnHead()
    '填充列表的列头
    lvLuggage.ColumnHeaders.Clear
    With lvLuggage.ColumnHeaders
        .Add , , "受理单号", 1400.31
        .Add , , "到站", 799.93
        .Add , , "托运人", 1050.71
        .Add , , "收件人", 1050.71
        .Add , , "车次日期", 1150.06
        .Add , , "时间", 850.06
        .Add , , "车次", 700.72
        .Add , , "件数", 700.15
        .Add , , "托运方式", 1000.06
        .Add , , "里程", 700
        .Add , , "计重", 700
        .Add , , "实重", 700
        .Add , , "标签号", 1440
        .Add , , "超重", 700
        .Add , , "托运费", 799
        .Add , , "运费", 0
        
    End With
End Sub
