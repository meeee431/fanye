VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmArrangeBus 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "计划--车次属性"
   ClientHeight    =   5370
   ClientLeft      =   1770
   ClientTop       =   3540
   ClientWidth     =   9495
   HelpContextID   =   10000320
   Icon            =   "frmArrangeBus.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   9495
   ShowInTaskbar   =   0   'False
   Begin RTComctl3.CoolButton cmdSellStation 
      Height          =   375
      Left            =   7785
      TabIndex        =   39
      Top             =   4890
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "售票点"
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
      MICON           =   "frmArrangeBus.frx":038A
      PICN            =   "frmArrangeBus.frx":03A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton imbSeatLeave 
      Height          =   375
      Left            =   3465
      TabIndex        =   34
      Top             =   4890
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "座位预留"
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
      MICON           =   "frmArrangeBus.frx":0740
      PICN            =   "frmArrangeBus.frx":075C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton imbTkPrice 
      Height          =   375
      Left            =   4830
      TabIndex        =   35
      Top             =   4890
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "车次票价"
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
      MICON           =   "frmArrangeBus.frx":0AF6
      PICN            =   "frmArrangeBus.frx":0B12
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton imbStationSet 
      Height          =   375
      Left            =   6195
      TabIndex        =   36
      Top             =   4890
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "车次站点"
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
      MICON           =   "frmArrangeBus.frx":0C6C
      PICN            =   "frmArrangeBus.frx":0C88
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdEnvPreview 
      Height          =   375
      Left            =   8220
      TabIndex        =   28
      Top             =   240
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "环境预览(&W)"
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
      MICON           =   "frmArrangeBus.frx":0DE2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton imbBusVehicle 
      Height          =   375
      Left            =   2085
      TabIndex        =   33
      Top             =   4890
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "车次车辆"
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
      MICON           =   "frmArrangeBus.frx":0DFE
      PICN            =   "frmArrangeBus.frx":0E1A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame fraBusInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "基本信息"
      ForeColor       =   &H80000008&
      Height          =   2235
      Left            =   150
      TabIndex        =   30
      Top             =   150
      Width           =   7875
      Begin VB.ComboBox cmbBusType 
         Height          =   300
         ItemData        =   "frmArrangeBus.frx":11B4
         Left            =   4470
         List            =   "frmArrangeBus.frx":11B6
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   270
         Width           =   1920
      End
      Begin VB.OptionButton OptInterNet1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "可售(&Y)"
         Height          =   285
         Left            =   1350
         TabIndex        =   21
         Top             =   1860
         Width           =   1005
      End
      Begin VB.OptionButton OptInterNet2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "不可售(&N)"
         Height          =   285
         Left            =   2430
         TabIndex        =   22
         Top             =   1860
         Width           =   1185
      End
      Begin FText.asFlatSpinEdit txtCycleStart 
         Height          =   300
         Left            =   4470
         TabIndex        =   11
         Top             =   1050
         Width           =   855
         _ExtentX        =   1508
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
         Text            =   "1"
         ButtonBackColor =   -2147483633
         Value           =   1
         Registered      =   -1  'True
      End
      Begin MSComCtl2.DTPicker dtpOffTime 
         Height          =   300
         Left            =   1320
         TabIndex        =   9
         Top             =   1050
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "HH:mm"
         Format          =   74514435
         UpDown          =   -1  'True
         CurrentDate     =   36392
      End
      Begin FText.asFlatTextBox txtRouteID 
         Height          =   300
         Left            =   1320
         TabIndex        =   5
         Top             =   660
         Width           =   1905
         _ExtentX        =   3360
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
      Begin FText.asFlatTextBox txtCheckGate 
         Height          =   300
         Left            =   4470
         TabIndex        =   7
         Top             =   660
         Width           =   1905
         _ExtentX        =   3360
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
      Begin FText.asFlatTextBox txtBusID 
         Height          =   315
         Left            =   1320
         TabIndex        =   1
         Top             =   270
         Width           =   1905
         _ExtentX        =   3360
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
         ButtonHotBackColor=   -2147483633
         ButtonPressedBackColor=   -2147483627
         Text            =   ""
         ButtonBackColor =   -2147483633
         ButtonVisible   =   -1  'True
         Registered      =   -1  'True
      End
      Begin FText.asFlatSpinEdit txtCycle 
         Height          =   300
         Left            =   6570
         TabIndex        =   13
         Top             =   1050
         Width           =   1125
         _ExtentX        =   1984
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
         Text            =   "1"
         ButtonBackColor =   -2147483633
         Value           =   1
         Registered      =   -1  'True
      End
      Begin FText.asFlatSpinEdit txtCheckTime 
         Height          =   300
         Left            =   2010
         TabIndex        =   15
         Top             =   1440
         Width           =   765
         _ExtentX        =   1349
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
         Text            =   "0"
         ButtonBackColor =   -2147483633
         Registered      =   -1  'True
      End
      Begin MSComCtl2.DTPicker dtpFirstBus 
         Height          =   300
         Left            =   4470
         TabIndex        =   17
         Top             =   1440
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "HH:mm"
         Format          =   74514435
         UpDown          =   -1  'True
         CurrentDate     =   36392
      End
      Begin MSComCtl2.DTPicker dtpLastBus 
         Height          =   300
         Left            =   6570
         TabIndex        =   19
         Top             =   1440
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "HH:mm"
         Format          =   74514435
         UpDown          =   -1  'True
         CurrentDate     =   36392
      End
      Begin VB.Label lblCycleInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "排班基准日期:"
         Height          =   255
         Left            =   4440
         TabIndex        =   38
         Top             =   1920
         Width           =   3135
      End
      Begin VB.Label lblRoute 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "运行线路(&I):"
         Height          =   180
         Left            =   180
         TabIndex        =   4
         Top             =   720
         Width           =   1080
      End
      Begin VB.Label lblCheckGate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "检票口(&C):"
         Height          =   180
         Left            =   3525
         TabIndex        =   6
         Top             =   720
         Width           =   900
      End
      Begin VB.Label lblCircly 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "运行周期(&Z):"
         Height          =   180
         Left            =   5445
         TabIndex        =   12
         Top             =   1110
         Width           =   1080
      End
      Begin VB.Label lblStartNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "起始序号(&Q):"
         Height          =   180
         Left            =   3345
         TabIndex        =   10
         ToolTipText     =   "周期起始序号"
         Top             =   1110
         Width           =   1080
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "发车时间(&M):"
         Height          =   180
         Left            =   180
         TabIndex        =   8
         Top             =   1110
         Width           =   1080
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "车次种类(&Y):"
         Height          =   180
         Left            =   3345
         TabIndex        =   2
         Top             =   330
         Width           =   1080
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "互联售票(&P):"
         Height          =   195
         Left            =   180
         TabIndex        =   20
         Top             =   1905
         Width           =   1185
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "滚动车次相隔时间(&G):"
         Height          =   180
         Left            =   180
         TabIndex        =   14
         Top             =   1500
         Width           =   1800
      End
      Begin VB.Label lblBusID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车次代码(&B):"
         Height          =   180
         Left            =   180
         TabIndex        =   0
         Top             =   330
         Width           =   1170
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "始班车(&F):"
         Enabled         =   0   'False
         Height          =   180
         Left            =   3525
         TabIndex        =   16
         Top             =   1500
         Width           =   900
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "末班车(&D):"
         Enabled         =   0   'False
         Height          =   180
         Left            =   5625
         TabIndex        =   18
         Top             =   1500
         Width           =   900
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "分钟"
         Height          =   180
         Left            =   2865
         TabIndex        =   31
         Top             =   1500
         Width           =   360
      End
   End
   Begin VB.Timer tmStart 
      Interval        =   500
      Left            =   8430
      Top             =   4470
   End
   Begin RTComctl3.CoolButton cmdHelp 
      Height          =   375
      Left            =   8220
      TabIndex        =   27
      Top             =   1980
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   661
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
      MICON           =   "frmArrangeBus.frx":11B8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ImageList ilsToolBar 
      Left            =   8400
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArrangeBus.frx":11D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArrangeBus.frx":156E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArrangeBus.frx":16C8
            Key             =   "del"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArrangeBus.frx":1A1B
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArrangeBus.frx":1B75
            Key             =   "Run"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArrangeBus.frx":1CCF
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmArrangeBus.frx":2B21
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin RTComctl3.CoolButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   8220
      TabIndex        =   26
      Top             =   1200
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "取消(&C)"
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
      MICON           =   "frmArrangeBus.frx":3973
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdOk 
      Default         =   -1  'True
      Height          =   375
      Left            =   8220
      TabIndex        =   25
      Top             =   720
      Width           =   1140
      _ExtentX        =   0
      _ExtentY        =   0
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
      MICON           =   "frmArrangeBus.frx":398F
      PICN            =   "frmArrangeBus.frx":39AB
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ListView lvVehicle 
      Height          =   1665
      Left            =   150
      TabIndex        =   24
      Top             =   2820
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   2937
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ilsToolBar"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "序号"
         Object.Width           =   1270
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "车辆代码"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "车牌"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "车型"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "座位数"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "参运公司"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "车主"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "起始座位"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "营运日期"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "停班开始时间"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "停班结束时间"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbBusManage 
      Height          =   360
      Left            =   5355
      TabIndex        =   29
      Top             =   2490
      Width           =   2670
      _ExtentX        =   4710
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ilsToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "VehicleInfo"
            Object.ToolTipText     =   "车次车辆"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "NewVehicle"
            Object.ToolTipText     =   "新增车辆"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DelVehicle"
            Object.ToolTipText     =   "删除车辆"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "BusVehicleStop"
            Object.ToolTipText     =   "车辆停班"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "BusVehicleResume"
            Object.ToolTipText     =   "车辆复班"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "BusVehicleMoveUp"
            Object.ToolTipText     =   "上移位置"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "BusVehicleMoveDown"
            Object.ToolTipText     =   "下移位置"
            ImageIndex      =   6
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin RTComctl3.CoolButton cmdAllot 
      Height          =   375
      Left            =   675
      TabIndex        =   37
      Top             =   4890
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "车次配载"
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
      MICON           =   "frmArrangeBus.frx":3D45
      PICN            =   "frmArrangeBus.frx":3D61
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   " 高级设置"
      Enabled         =   0   'False
      Height          =   1200
      Left            =   -120
      TabIndex        =   32
      Top             =   4620
      Width           =   9705
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "运行车辆列表(&L):"
      Height          =   180
      Left            =   150
      TabIndex        =   23
      Top             =   2520
      Width           =   1440
   End
End
Attribute VB_Name = "frmArrangeBus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Public m_bIsParent As Boolean '是否是父窗体打开
'Public m_bShow As Boolean
Public m_szBusID As String   '当前车次代码

Private m_bflg As Boolean '车次车辆车型改变
Private m_oBus As New Bus '车次对象的引用 Bus
Private szSaveFormName As String

Private m_oBusType As New BusType
Private m_oVehicle As New Vehicle '车辆对象 Vehicle
Private m_bVehicleModeIdChang As Boolean
Private m_nSerial As Integer '车辆的序号
Private m_bAddNew As Boolean '是否是新增状态
Private m_bChange As Boolean '是否修改车次序号
Private m_bChangeSerial As Boolean '是否修改车次序号
Private m_szOldRoute As String

Private m_atBusVehicle() As TBusVehicleInfo  '车次车辆结构
Private m_atBusVehicleAddTemp() As TBusVehicleInfo
Private m_aszBusType() As String
Private m_anOldSerial() As Integer '旧的车次车辆序号数组
Private m_anOldVehicleModelID() As String  '旧的车次车辆车型数组


'以下变量定义
Private Sub cmbBusType_Change()
    If ResolveDisplay(cmbBusType.Text) = TP_ScrollBus Then
       txtCheckTime.Enabled = True
    Else
       txtCheckTime.Text = 0
    End If
    IsChanged
End Sub
Private Sub cmbBusType_Click()
    If ResolveDisplay(cmbBusType.Text) = TP_ScrollBus Then
        txtCheckTime.Enabled = True
    Else
        txtCheckTime.Text = 0
       txtCheckTime.Enabled = False
    End If
    IsChanged
End Sub


Private Sub cmdAllot_Click()
    frmBusAllot.m_bIsAllot = True
    frmBusAllot.m_szBusID = m_szBusID
    frmBusAllot.Show vbModal
    
    '刷新检票口及发车时间信息
    m_oBus.Identify m_szBusID
    Dim oBase As New CheckGate
    oBase.Init g_oActiveUser
    oBase.Identify m_oBus.CheckGate
    txtCheckGate.Text = MakeDisplayString(m_oBus.CheckGate, Trim(oBase.CheckGateName))
    dtpOffTime.Value = m_oBus.StartUpTime
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdHelp_Click()
    DisplayHelp Me
End Sub
Private Sub cmdOk_Click()
    On Error GoTo ErrHandle
    ModifyBus
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub cmdSellStation_Click()
    frmBusAllot.m_bIsAllot = False
    frmBusAllot.m_szBusID = m_szBusID
    frmBusAllot.Show vbModal
End Sub

Private Sub dtpOffTime_Change()
    IsChanged
End Sub
Private Sub Form_Load()
    On Error GoTo ErrHandle
    AlignFormPos Me
    
    Dim liTemp As ListItem
    Dim oBase As New CheckGate
    m_oVehicle.Init g_oActiveUser
    m_oBus.Init g_oActiveUser
    m_oBus.Identify m_szBusID
    m_oBusType.Init g_oActiveUser
    oBase.Init g_oActiveUser
    m_bChange = False
    GetBusType
    txtCycle.Text = m_oBus.RunCycle
    txtCycleStart.Text = m_oBus.CycleStartSerialNo
    oBase.Identify m_oBus.CheckGate
    txtCheckGate.Text = MakeDisplayString(m_oBus.CheckGate, Trim(oBase.CheckGateName))
    txtRouteID.Text = MakeDisplayString(m_oBus.Route, Trim(m_oBus.RouteName))
    m_szOldRoute = m_oBus.Route
    txtBusID.Text = m_szBusID
    IndenfiyBusType m_oBus.BusType
    txtCheckTime.Text = CStr(m_oBus.ScrollBusCheckTime)
    If m_oBus.InternetStatus = CnInternetCanSell Then
        Optinternet1.Value = True
    Else
        Optinternet2.Value = True
    End If
    dtpOffTime.Value = m_oBus.StartUpTime
    txtBusID.Enabled = False
    cmdOk.Enabled = False
'    m_bShow = True
Exit Sub
ErrHandle:
    tmStart.Enabled = False
    ShowErrorMsg
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim nResult As VbMsgBoxResult
'    If cmdOk.Enabled = True Then
'        nResult = MsgBox("是否保存车次的修改", vbQuestion + vbYesNoCancel, "计划")
'        If nResult = vbYes Then
'            ModifyBus
'        End If
'        If nResult = vbCancel Then Cancel = 1: Exit Sub
'    End If
'    m_bShow = False
    SaveFormPos Me
    m_bIsParent = False
    
End Sub

Private Sub imbBusVehicle_Click()
    If IsSave = False Then Exit Sub
    frmBusVehicleMan.Init m_oBus
    frmBusVehicleMan.Show vbModal
End Sub


Private Sub cmdEnvPreview_Click()
    BusPreview
End Sub

Private Sub imbSeatLeave_Click()
    frmReserveSeat.m_bIsParent = True
    frmReserveSeat.Init m_oBus
    frmReserveSeat.Show vbModal

End Sub

Private Sub imbStationSet_Click()
    BusStation
End Sub

Private Sub imbTkPrice_Click()
    If IsSave = False Then Exit Sub
    ShowPrice
End Sub

Private Sub lvVehicle_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    SortListView lvVehicle, ColumnHeader.Index
End Sub

Private Sub lvVehicle_DblClick()
'    If Not lvVehicle.SelectedItem Is Nothing Then
'        frmVehicle.mszVehicleId = Trim(lvVehicle.SelectedItem.ListSubItems(1).Text)
'        frmVehicle.Status = EFS_Modify
''        'frmVehicle.m_szfromstatus = "计划-车次属性"
''        frmVehicle.m_szBusID = ResolveDisplay(txtBusID.Text)
'        frmVehicle.Show vbModal
'    End If
    VehicleInfo
    
End Sub

Private Sub lvVehicle_ItemClick(ByVal Item As MSComctlLib.ListItem)
    m_nSerial = Val(Item.Index)
    If Item.SmallIcon = "Stop" Then
        tbBusManage.Buttons("BusVehicleStop").Enabled = False
        tbBusManage.Buttons("BusVehicleResume").Enabled = True
    Else
        tbBusManage.Buttons("BusVehicleStop").Enabled = True
        tbBusManage.Buttons("BusVehicleResume").Enabled = False
    End If
    If Item.ForeColor = vbBlue Then
        tbBusManage.Buttons("BusVehicleStop").Enabled = False
        tbBusManage.Buttons("BusVehicleResume").Enabled = False
    End If
End Sub

Private Sub lvVehicle_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            lvVehicle_DblClick
    End Select
End Sub


Private Sub mnu_BusPreview_Click()
    BusPreview
End Sub

Private Sub BusStation()
    frmBusRoute.Init m_oBus
    frmBusRoute.Show vbModal
End Sub

Private Sub mnu_BusTicketPrice_Click()
    ShowPrice
End Sub

Private Sub BusVehicleResume()
    If lvVehicle.SelectedItem Is Nothing Then
        MsgBox "请选择车辆!", vbExclamation, "提示"
        Exit Sub
    End If

Dim oScheme As New REScheme
    If m_nSerial = 0 Then Exit Sub

   m_oBus.BusVehicleRun Trim(lvVehicle.ListItems(m_nSerial).ListSubItems(1)), cszEmptyDateStr, cszEmptyDateStr

   If MsgBox("计划车辆复班成功" & Chr(10) & "是否影响环境？", vbInformation + vbYesNo, Me.Caption) = vbYes Then

      oScheme.Init g_oActiveUser
      oScheme.StopOrResumBusVehile m_oBus.BusID, lvVehicle.ListItems(m_nSerial).ListSubItems(1), True
      MsgBox "环境车次车辆复班成功!", vbInformation, Me.Caption
      Set oScheme = Nothing

   End If

   RefreshVehicle
   If m_bIsParent = True Then
      frmBus.UpdateList m_oBus.BusID
   End If

   'cmdOk.Enabled = True
End Sub

Private Sub BusVehicleStop()
    If lvVehicle.SelectedItem Is Nothing Then
        MsgBox "请选择车辆!", vbExclamation, "提示"
        Exit Sub
    End If

    Dim frmStop As frmBusStop

    Set frmStop = New frmBusStop


    If lvVehicle.ListItems.Count = 0 Then Exit Sub
    If m_nSerial = 0 Then m_nSerial = 1


    frmStop.m_szVehicle = Trim(lvVehicle.ListItems(m_nSerial).ListSubItems(1))

    frmStop.m_szBusID = m_szBusID
    frmStop.m_szPlanID = g_szExePriceTable
    frmStop.Caption = "车次车辆停班"
    frmStop.lblCheck = "车辆: " & Trim(lvVehicle.ListItems(m_nSerial).ListSubItems(1))
    frmStop.chStop.Visible = False

    frmStop.cmdAllInfo.Visible = False
    frmStop.cmdHelp.Visible = False
    frmStop.Show vbModal
    Set frmStop = Nothing
    RefreshVehicle
    If m_bIsParent Then
       frmBus.UpdateList m_oBus.BusID
    End If
End Sub


Private Sub DelVehicle()
On Error GoTo ErrHandle
    If lvVehicle.SelectedItem Is Nothing Then
        MsgBox "请选择车辆!", vbExclamation, "提示"
        Exit Sub
    End If
    With lvVehicle.SelectedItem
        m_oBus.DeleteRunVehicle Val(lvVehicle.ListItems(.Index).Text)
        lvVehicle.ListItems.Remove .Index
        m_atBusVehicle = m_oBus.GetAllVehicleEx
        RefreshVehicle
        cmdOk.Enabled = True
    End With
    Exit Sub
ErrHandle:
    If err.Number = 14434 Then
        lvVehicle.ListItems.Remove lvVehicle.SelectedItem.Index
    Else
        ShowErrorMsg
    End If
End Sub

Private Sub AddVehicle()
    frmQueryVehicle.Show vbModal
    If frmQueryVehicle.IsCancel Then
        Unload frmQueryVehicle
        Exit Sub
    End If
    Dim oShell As New CommDialog
    oShell.Init g_oActiveUser
    Dim aszTmp() As String
    With frmQueryVehicle
    aszTmp = oShell.SelectVehicle(Trim(.txtVehicle.Text), Trim(ResolveDisplay(.txtCompany.Text)), Trim(ResolveDisplay(.txtBusOwner.Text)), _
                                  Trim(ResolveDisplay(.txtVehicleType.Text)), Trim(.txtLicense.Text), True)
    End With
    If ArrayLength(aszTmp) = 0 Then Exit Sub
    
    '添加车辆至列表中
    Dim i As Integer
    For i = 1 To ArrayLength(aszTmp)
        AddBusVehicleItem aszTmp(i, 1)
    Next i
    lvVehicle.Refresh
    
End Sub

Private Sub OptScroll_Click()
    dtpOffTime.Enabled = False
    IsChanged
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
       Case vbKeyEscape
       Unload Me
       Case vbKeyReturn
       SendKeys "{TAB}"
End Select
End Sub

Private Sub mnu_BusVehicleStop_Click()

End Sub

Private Sub OptInterNet1_Click()
    IsChanged
End Sub

Private Sub OptInterNet2_Click()
    IsChanged
End Sub



Private Sub tbBusManage_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo ErrHandle
    Select Case Button.Key
        Case "VehicleInfo"
            VehicleInfo
        Case "NewVehicle"
            AddVehicle
        Case "DelVehicle"
            DelVehicle
        Case "BusVehicleStop"
            BusVehicleStop
        Case "BusVehicleResume"
            BusVehicleResume
        Case "BusVehicleMoveUp"
            BusVehicleMoveUp
        Case "BusVehicleMoveDown"
            BusVehicleMoveDown
    End Select
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

'车辆信息查看及更改
Private Sub VehicleInfo()
On Error GoTo ErrHandle
    If lvVehicle.SelectedItem Is Nothing Then
        MsgBox "请选择车辆!", vbExclamation, "提示"
        Exit Sub
    End If
    
    frmVehicle.mszVehicleId = lvVehicle.SelectedItem.SubItems(1)
    frmVehicle.Status = EFS_Modify
    frmVehicle.m_bIsParent = False
    frmVehicle.m_bIsArrayBus = True
    frmVehicle.Show vbModal
    
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub tmStart_Timer()
    SetBusy
    tmStart.Enabled = False
    RefreshVehicle
    SetNormal
End Sub

Private Sub txtBusID_ButtonClick()
On Error GoTo ErrHandle
    Dim oShell As New CommDialog
    Dim aszTmp() As String
    oShell.Init g_oActiveUser
    aszTmp = oShell.SelectBus
    Set oShell = Nothing
    If ArrayLength(aszTmp) = 0 Then Exit Sub
    txtBusID.Text = aszTmp(1, 1)
    RefreshVehicle
    m_szBusID = ResolveDisplay(txtBusID.Text)
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub txtCheckGate_Change()
    IsChanged
End Sub
Private Sub txtCheckGate_ButtonClick()
    Dim oShell As New CommDialog
    Dim aszTmp() As String
    oShell.Init g_oActiveUser
    aszTmp = oShell.SelectCheckGate
    Set oShell = Nothing
    If ArrayLength(aszTmp) = 0 Then Exit Sub
    txtCheckGate.Text = MakeDisplayString(aszTmp(1, 1), aszTmp(1, 2))
End Sub

Private Sub txtCheckTime_Change()
  IsChanged
  FormatTextToNumeric txtCheckTime, False, False
End Sub

Private Sub txtCycle_Change()
    IsChanged
    FormatTextToNumeric txtCycle, False, False
End Sub

Private Sub txtCycleStart_Change()
    IsChanged
    FormatTextToNumeric txtCycleStart, False, False
End Sub

Private Sub txtRouteID_Change()
    IsChanged
End Sub

Private Sub txtRouteID_ButtonClick()
    Dim oShell As New CommDialog
    Dim aszTmp() As String
    oShell.Init g_oActiveUser
    aszTmp = oShell.SelectRoute
    Set oShell = Nothing
    If ArrayLength(aszTmp) = 0 Then Exit Sub
    txtRouteID.Text = MakeDisplayString(aszTmp(1, 1), aszTmp(1, 2))
End Sub

Private Sub udCirclyStart_Change()

End Sub

'修改车次序号1
Private Sub BusVehicleMoveDown()
    On Error GoTo ErrHandle
    lvVehicle.SelectedItem.Text = lvVehicle.SelectedItem.Text + 1
    m_atBusVehicle(lvVehicle.SelectedItem.Index).nSerialNo = lvVehicle.SelectedItem.Text
    m_bChangeSerial = True
    IsChanged
Exit Sub
ErrHandle:
End Sub
'修改车次序号2
Private Sub BusVehicleMoveUp()
    On Error GoTo ErrHandle
    If (lvVehicle.SelectedItem.Text - 1) = 0 Then
        MsgBox "车次车辆序号不能为0", vbExclamation + vbOKOnly, "车次管理"
    Else
        lvVehicle.SelectedItem.Text = lvVehicle.SelectedItem.Text - 1
        m_atBusVehicle(lvVehicle.SelectedItem.Index).nSerialNo = lvVehicle.SelectedItem.Text
    End If
    m_bChangeSerial = True
    IsChanged
Exit Sub
ErrHandle:
End Sub

Private Sub BusPreview()
    frmBusPreview.RealTimeInit m_szBusID, m_atBusVehicle, True, Val(txtCycle.Text), Val(txtCycleStart.Text)
    frmBusPreview.Show vbModal
End Sub

Public Sub RefreshVehicle()
    '获得车次运行的车辆
    Dim liTemp As ListItem
    Dim nVehicleCount As Integer, i As Integer, j As Integer
    Dim nCountTemp As Integer
    Dim bflg As Boolean
    Dim bflg2 As Boolean
    Dim oScheme As New REScheme
    Dim nExeVehicleSeralNo As Integer
    On Error GoTo ErrHandle
    oScheme.Init g_oActiveUser
    ReDim m_anOldVehicleModelID(1) As String
    If txtBusID.Text <> "" Then
        m_oBus.Identify Trim(txtBusID.Text)
    End If
    nExeVehicleSeralNo = oScheme.GetExecuteVehicleSerialNo(m_oBus.RunCycle, m_oBus.CycleStartSerialNo, Now)
    Set oScheme = Nothing
    m_atBusVehicle = m_oBus.GetAllVehicleEx
    nVehicleCount = ArrayLength(m_atBusVehicle)
    If nVehicleCount = 0 Then Exit Sub
    ReDim m_anOldSerial(1 To nVehicleCount) As Integer
    '显示该车次的所有车次车辆
    lvVehicle.ListItems.Clear
    For i = 1 To nVehicleCount
        Set liTemp = lvVehicle.ListItems.Add(, , m_atBusVehicle(i).nSerialNo)
        If SelfGetBusStatus(m_atBusVehicle(i).dtBeginStopDate, m_atBusVehicle(i).dtEndStopDate, CDate(Date)) = ST_BusStopped Then
            liTemp.SmallIcon = "Stop"
            bflg2 = True
        End If
        m_oVehicle.Identify m_atBusVehicle(i).szVehicleID
        If m_oVehicle.Status = ST_VehicleRun Then
            If bflg2 = False Then
                liTemp.SmallIcon = "Run"
            End If
        Else
            liTemp.SmallIcon = "Stop"
            bflg2 = False
        End If
        liTemp.SubItems(1) = m_atBusVehicle(i).szVehicleID
        m_anOldSerial(i) = m_atBusVehicle(i).nSerialNo
        liTemp.SubItems(2) = m_oVehicle.LicenseTag
        WriteProcessBar , i, nVehicleCount, "车辆" & m_oVehicle.LicenseTag
        liTemp.SubItems(3) = m_oVehicle.VehicleModelName
        liTemp.SubItems(4) = m_oVehicle.SeatCount
        liTemp.SubItems(5) = m_oVehicle.CompanyName
        liTemp.SubItems(6) = m_oVehicle.OwnerName
        liTemp.SubItems(7) = m_oVehicle.StartSeatNo
        If m_oBus.BusType <> TP_ScrollBus Then
            If liTemp.SmallIcon = "Run" Then
                liTemp.SubItems(8) = GetVehicleRunDate(nVehicleCount, nExeVehicleSeralNo, m_atBusVehicle(i).nSerialNo, True)
            Else
    
                liTemp.SubItems(8) = GetVehicleRunDate(nVehicleCount, nExeVehicleSeralNo, m_atBusVehicle(i).nSerialNo, False)
            End If
        Else
            If liTemp.SmallIcon = "Run" Then
                liTemp.SubItems(8) = Format(Now, "YYYY-MM-DD")
            Else
                liTemp.SubItems(8) = Format(Now, "YYYY-MM-DD") & "(停)"
            End If
        End If
        If bflg2 = True Then
           liTemp.SubItems(9) = m_atBusVehicle(i).dtBeginStopDate
           liTemp.SubItems(10) = m_atBusVehicle(i).dtEndStopDate
           bflg2 = False
        End If
        MakeArray m_anOldVehicleModelID, m_oVehicle.VehicleModel
    Next
    WriteProcessBar False
Exit Sub
ErrHandle:
    WriteProcessBar False
    ShowErrorMsg
End Sub
'编辑车次
Private Sub ModifyBus(Optional flg As Boolean)
    Dim szReCode() As String
    Dim nCount As Integer
    Dim bflg As Boolean
    Dim nCase As Integer
    Dim szMsg As String
    Dim nCountTemp As Integer
    Dim tBusV() As TBusVehicleInfo
    On Error GoTo ErrHandle
        Dim i As Integer

    m_atBusVehicle = GetBusVehicleFromLvLst

    If m_bflg Then
        '车型改变
        szMsg = "车辆的车型,保存后请重新生成车次票价" & Chr(10)
    End If

    If m_szOldRoute <> ResolveDisplay(txtRouteID.Text) Then
        '线路改变
        szMsg = szMsg & "您修改了线路，保存后请重新生成车次票价" & Chr(10)

    End If
    Dim nResult As VbMsgBoxResult
    If szMsg <> "" Then
        nResult = MsgBox(szMsg & "是否保存？", vbQuestion + vbYesNo, Me.Caption)
    Else
        nResult = vbYes ' MsgBox("是否保存？", vbQuestion + vbYesNo, Me.Caption)
    End If

    If nResult = vbNo Then Exit Sub

    If m_bChangeSerial Then
        szReCode = ReCode
        If szReCode(1) <> "" Then
            MsgBox "车次车辆代码有重复" & szReCode(1) & "--" & szReCode(2), vbExclamation + vbOKOnly, "车次车辆"
            Exit Sub
        End If
        SetBusy
        m_oBus.DeleteAllRunVehicle
        nCount = ArrayLength(m_atBusVehicle)
        For i = 1 To nCount
            m_oBus.AddRunVehicle m_atBusVehicle(i)
        Next
        m_oBus.Route = ResolveDisplay(txtRouteID.Text)
        m_oBus.StartUpTime = dtpOffTime.Value
        m_oBus.BusType = ResolveDisplay(cmbBusType.Text)
        m_oBus.CheckGate = ResolveDisplay(txtCheckGate.Text)
        m_oBus.RunCycle = Val(txtCycle.Text)
        m_oBus.CycleStartSerialNo = Val(txtCycleStart.Text)
        If m_oBus.BusType <> 1 Then
          m_oBus.ScrollBusCheckTime = 0
        Else
          m_oBus.ScrollBusCheckTime = CInt(Val((txtCheckTime.Text)))
        End If
        If Optinternet2.Value = True Then
          m_oBus.InternetStatus = CnInternetNotCanSell
        Else
          m_oBus.InternetStatus = CnInternetCanSell
        End If
        m_oBus.Update
       If m_szOldRoute <> ResolveDisplay(txtRouteID.Text) Then
    '          m_oBus.DeleteBusPrice
    '          ncount = ArrayLength(m_atBusVehicle)
    '          For i = 1 To ncount
    '                m_oBus.MakeBusPrice m_atBusVehicle(i)
    '          Next

       End If
    End If

    m_oBus.Route = ResolveDisplay(txtRouteID.Text)
    m_oBus.StartUpTime = dtpOffTime.Value
    m_oBus.BusType = ResolveDisplay(cmbBusType.Text)
    m_oBus.CheckGate = ResolveDisplay(txtCheckGate.Text)
    m_oBus.RunCycle = Val(txtCycle.Text)
    m_oBus.CycleStartSerialNo = Val(txtCycleStart.Text)
    If m_oBus.BusType <> 1 Then
      m_oBus.ScrollBusCheckTime = 0
    Else
      m_oBus.ScrollBusCheckTime = CInt(Val((txtCheckTime.Text)))
    End If
    If Optinternet2.Value = True Then
      m_oBus.InternetStatus = CnInternetNotCanSell
    Else
      m_oBus.InternetStatus = CnInternetCanSell
    End If
    m_oBus.Update
    If m_bIsParent Then
        frmBus.UpdateList m_oBus.BusID
    End If
    
    SetNormal
    If flg = False Then
        cmdOk.Enabled = False
    End If
    m_bChange = False
    m_bChangeSerial = False
    If m_bIsParent Then
       frmBus.UpdateList Trim(txtBusID.Text)
    End If

    For i = 1 To lvVehicle.ListItems.Count
        SetListViewLineColor lvVehicle, i, &H80000008
    Next i
    lvVehicle.Refresh
    Exit Sub
ErrHandle:
    SetNormal
    ShowErrorMsg
End Sub

'查询车次车辆代码中的重复序号
Public Function ReCode() As String()
Dim nCount As Integer, j As Integer, i As Integer
Dim nTemp As Integer
Dim szReCode(1 To 2) As String
nCount = ArrayLength(m_atBusVehicle)
    For i = 1 To nCount
        nTemp = m_atBusVehicle(i).nSerialNo
        For j = 1 To nCount
         If nTemp = m_atBusVehicle(j).nSerialNo And i <> j Then
              szReCode(1) = m_atBusVehicle(i).szVehicleID
              szReCode(2) = m_atBusVehicle(j).szVehicleID
              ReCode = szReCode
              Exit Function
         End If
        Next
    Next
szReCode(1) = ""
szReCode(2) = ""
ReCode = szReCode
End Function

'添加车辆至排班
Private Sub AddBusVehicleItem(VehicleId As String)
On Error GoTo ErrHandle
    Dim nCount As Integer, nCountTemp As Integer
    Dim liTemp As ListItem
    Dim szVehicleModel As String
    m_oVehicle.Identify VehicleId
    
    nCount = ArrayLength(m_atBusVehicle)
    
    
    If nCount = 0 Then
        ReDim m_atBusVehicle(1) As TBusVehicleInfo
        szVehicleModel = ""
    Else
        ReDim Preserve m_atBusVehicle(1 To nCount + 1) As TBusVehicleInfo
        szVehicleModel = Trim(lvVehicle.ListItems.Item(1).ListSubItems(3).Text)
        m_bflg = True
    End If
    
    Set liTemp = lvVehicle.ListItems.Add(, , nCount + 1, , "Run")
        If m_oVehicle.Status = ST_VehicleRun Then
            liTemp.SmallIcon = "Run"
        Else
            liTemp.SmallIcon = "Stop"
        End If
        liTemp.SubItems(1) = VehicleId
        liTemp.SubItems(2) = m_oVehicle.LicenseTag
        liTemp.SubItems(3) = m_oVehicle.VehicleModelName
        liTemp.SubItems(4) = m_oVehicle.SeatCount
        liTemp.SubItems(5) = m_oVehicle.CompanyName
        liTemp.SubItems(6) = m_oVehicle.OwnerName
        
        SetListViewLineColor lvVehicle, liTemp.Index, vbBlue
    
    m_bChangeSerial = True
    
    m_atBusVehicle(nCount + 1).szVehicleID = VehicleId
    m_atBusVehicle(nCount + 1).nStandTicketCount = 0
    m_atBusVehicle(nCount + 1).nSerialNo = nCount + 1
    m_atBusVehicle(nCount + 1).dtBeginStopDate = CDate(cszEmptyDateStr)
    m_atBusVehicle(nCount + 1).dtEndStopDate = CDate(cszEmptyDateStr)
    
    
    IsChanged
    
    Exit Sub

ErrHandle:
    ShowErrorMsg
End Sub

Private Sub ShowPrice()
    frmBusPrice.m_szBusID = Trim(txtBusID.Text)
'    frmBusPrice.mbProjectBus = True
    frmBusPrice.Show vbModal

End Sub

Private Sub IsChanged()
m_bChange = True
cmdOk.Enabled = True
End Sub

Public Sub ChangeVehicle(tChangeBusVehicle() As TBusVehicleInfo)
    m_atBusVehicle = tChangeBusVehicle
End Sub
Public Function IsSave() As Boolean
   If lvVehicle.ListItems.Count = 0 And cmdOk.Enabled = True Then
     MsgBox "可能无车辆,请新增车辆后再试试", vbInformation, "计划--车次属性"
     IsSave = False
     Exit Function
   End If
   IsSave = True
End Function
Private Sub GetBusType()
   Dim i As Integer
   Dim nCount As Integer
   m_oBusType.ObjStatus = ST_NormalObj
   m_aszBusType = m_oBusType.GetAllBusType
   nCount = ArrayLength(m_aszBusType)
'   ReDim m_aszBusType(1 To nCount, 1 To 3)
   For i = 1 To nCount
   cmbBusType.AddItem MakeDisplayString(m_aszBusType(i, 1), m_aszBusType(i, 2))
   Next
End Sub
Private Sub IndenfiyBusType(szBusType As String)
   Dim i As Integer
   Dim nCount As Integer
   nCount = ArrayLength(m_aszBusType)
   nCount = cmbBusType.ListCount
   Do While szBusType <> CInt(Val(ResolveDisplay(cmbBusType.List(i))))
    i = i + 1
    If i > nCount Then GoTo ErrHandle
   Loop
   cmbBusType.ListIndex = i
   Exit Sub
ErrHandle:
   MsgBox "车次种类设置有误", vbInformation, "提示"
End Sub
Private Function GetBusVehicleFromLvLst() As TBusVehicleInfo()
    Dim nCount As Integer, nCountVehicleMode As Integer
    Dim szVehicle As String
    Dim i As Integer, j As Integer
    Dim tBusVehicleTemp() As TBusVehicleInfo
    Dim bflgTemp  As Boolean
    nCount = lvVehicle.ListItems.Count
    If nCount = 0 Then Exit Function
    ReDim tBusVehicleTemp(1 To nCount)
    nCountVehicleMode = ArrayLength(m_anOldVehicleModelID)
    For i = 1 To nCount
        szVehicle = Trim(lvVehicle.ListItems(i).ListSubItems(1))
        m_oVehicle.Identify szVehicle
        tBusVehicleTemp(i).szVehicleID = szVehicle
        tBusVehicleTemp(i).nStandTicketCount = 0
        tBusVehicleTemp(i).nSerialNo = CInt(lvVehicle.ListItems(i).Text)
        tBusVehicleTemp(i).dtBeginStopDate = CDate(cszEmptyDateStr)
        tBusVehicleTemp(i).dtEndStopDate = CDate(cszEmptyDateStr)
        '检测车型改变
        If FindOtherInArray(m_anOldVehicleModelID, Trim(m_oVehicle.VehicleModel)) = False Then
         '车型改变
          m_bflg = True
        End If
    Next
    GetBusVehicleFromLvLst = tBusVehicleTemp
End Function
Private Function SelfGetBusStatus(pdtBeginStopDate As Date, pdtEndStopDate As Date, pdtRunDate As Date) As EREBusStatus
    SelfGetBusStatus = ST_BusNormal
    If pdtRunDate >= pdtBeginStopDate And pdtRunDate <= pdtEndStopDate Then
       SelfGetBusStatus = ST_BusStopped
    End If
End Function
'''''''
'得到车辆运行时间
Private Function GetVehicleRunDate(nCountVehicle As Integer, nExeVehicleSeralNo As Integer, nVehicleSeralNo As Integer, bflg As Boolean) As String
    Dim szVehicleDate As String
    Dim nDate As Integer
    Dim szTemp As String
    If bflg = False Then
        szTemp = "(停)"
    End If
    If nExeVehicleSeralNo = nVehicleSeralNo Then
        szVehicleDate = Format(Now, "YYYY-MM-DD")
    Else
    If nExeVehicleSeralNo = nCountVehicle Then
    
        szVehicleDate = Format(DateAdd("d", nVehicleSeralNo, Now), "YYYY-MM-DD")
    
    Else
        If nExeVehicleSeralNo > nVehicleSeralNo Then
            nDate = nCountVehicle - nExeVehicleSeralNo + nVehicleSeralNo
        Else
            nDate = nVehicleSeralNo - nExeVehicleSeralNo
        
        End If
    szVehicleDate = Format(DateAdd("d", nDate, Now), "YYYY-MM-DD")
    End If
    End If
    GetVehicleRunDate = szVehicleDate & szTemp
End Function



Private Function FindOtherInArray(ByRef aszTemp() As String, szOther As String) As Boolean
    Dim nCountTemp As Integer
    Dim j As Integer
    Dim bflgTemp As Boolean
    
    nCountTemp = ArrayLength(aszTemp)
    j = 1
    Do While Not Trim(aszTemp(j)) = Trim(szOther)
        j = j + 1
        If j > nCountTemp Then bflgTemp = True: Exit Do
    Loop
    If bflgTemp = True Then
        FindOtherInArray = False
    Else
        FindOtherInArray = True
    End If
End Function
