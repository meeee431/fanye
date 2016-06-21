VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmWizardAddBus 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "新增车次向导"
   ClientHeight    =   4590
   ClientLeft      =   2880
   ClientTop       =   3510
   ClientWidth     =   7155
   HelpContextID   =   10000600
   Icon            =   "frmWizardAddBus.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   7155
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox ptLeftLogo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3660
      Left            =   210
      Picture         =   "frmWizardAddBus.frx":038A
      ScaleHeight     =   242
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   131
      TabIndex        =   58
      Top             =   180
      Width           =   1995
      Begin VB.Label lblStepTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "用户信息"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   180
         Left            =   165
         TabIndex        =   60
         Top             =   120
         Width           =   780
      End
      Begin VB.Label lblStepDetail 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "    本步骤用于输入新建用户的基本信息及口令"
         ForeColor       =   &H00808000&
         Height          =   1200
         Left            =   165
         TabIndex        =   59
         Top             =   465
         Width           =   1605
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraStep1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   3900
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   7095
      Begin VB.CheckBox chkHavePreSell 
         BackColor       =   &H00E0E0E0&
         Caption         =   "所有预售天内(&V)"
         Height          =   285
         Left            =   2745
         TabIndex        =   22
         Top             =   2700
         Width           =   2205
      End
      Begin MSComCtl2.DTPicker dtpEnvBus 
         Height          =   315
         Left            =   3855
         TabIndex        =   21
         Top             =   2340
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   123994113
         CurrentDate     =   37468
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   1125
         Index           =   0
         Left            =   2430
         TabIndex        =   55
         Top             =   210
         Width           =   4545
         Begin VB.Line Line3 
            BorderColor     =   &H00808080&
            Index           =   0
            X1              =   0
            X2              =   3360
            Y1              =   930
            Y2              =   930
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00FFFFFF&
            Index           =   0
            X1              =   4380
            X2              =   600
            Y1              =   135
            Y2              =   135
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            Index           =   0
            X1              =   4350
            X2              =   600
            Y1              =   120
            Y2              =   120
         End
         Begin VB.Label lblStep 
            BackStyle       =   0  'Transparent
            Caption         =   "第一步"
            ForeColor       =   &H00808000&
            Height          =   180
            Index           =   0
            Left            =   0
            TabIndex        =   57
            Top             =   0
            Width           =   1170
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "新增车次按类型可分为（1）调度计划中新增车次。（2）运行环境内新增车次，在运行环境内新增的车次将直接影响售票、检票等实时运行的系统。"
            Height          =   615
            Index           =   0
            Left            =   0
            TabIndex        =   56
            Top             =   270
            Width           =   4515
         End
      End
      Begin VB.OptionButton optAddEnvBus 
         BackColor       =   &H00E0E0E0&
         Caption         =   "新增环境车次(&E)"
         Height          =   180
         Left            =   2430
         TabIndex        =   19
         Top             =   2070
         Width           =   1965
      End
      Begin VB.OptionButton optAddBus 
         BackColor       =   &H00E0E0E0&
         Caption         =   "新增计划车次(&B)"
         Height          =   180
         Left            =   2430
         TabIndex        =   17
         Top             =   1380
         Width           =   1965
      End
      Begin VB.CheckBox chkRE 
         BackColor       =   &H00E0E0E0&
         Caption         =   "运行环境内新增车次(&R)"
         Height          =   285
         Left            =   2745
         TabIndex        =   18
         Top             =   1665
         Visible         =   0   'False
         Width           =   2205
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车次日期(&D):"
         Height          =   180
         Left            =   2745
         TabIndex        =   20
         Top             =   2400
         Width           =   1080
      End
   End
   Begin VB.Frame fraStep2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "第二步/共五步"
      Height          =   3900
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Width           =   7095
      Begin FText.asFlatSpinEdit txtCheckTime 
         Height          =   300
         Left            =   4230
         TabIndex        =   11
         Top             =   2580
         Width           =   915
         _ExtentX        =   1614
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
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   1125
         Index           =   1
         Left            =   2400
         TabIndex        =   61
         Top             =   210
         Width           =   4545
         Begin VB.Label lblStep 
            BackStyle       =   0  'Transparent
            Caption         =   "第二步"
            ForeColor       =   &H00808000&
            Height          =   180
            Index           =   6
            Left            =   0
            TabIndex        =   72
            Top             =   0
            Width           =   1170
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            Index           =   6
            X1              =   4350
            X2              =   600
            Y1              =   120
            Y2              =   120
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00FFFFFF&
            Index           =   6
            X1              =   4380
            X2              =   600
            Y1              =   135
            Y2              =   135
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "要新增车次必须填写车次代码、发车时间、运行线路、检票口和运行车辆（常规调度中运行车辆可能有多辆），车次代码不允许重复。"
            Height          =   765
            Left            =   0
            TabIndex        =   62
            Top             =   270
            Width           =   4500
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            Index           =   1
            X1              =   4320
            X2              =   1230
            Y1              =   120
            Y2              =   120
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00FFFFFF&
            Index           =   1
            X1              =   4320
            X2              =   1230
            Y1              =   135
            Y2              =   135
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00808080&
            Index           =   1
            X1              =   0
            X2              =   3360
            Y1              =   1110
            Y2              =   1110
         End
      End
      Begin VB.OptionButton optInterNet2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "不可售"
         Height          =   180
         Left            =   4830
         TabIndex        =   16
         Top             =   3390
         Width           =   915
      End
      Begin VB.OptionButton optInterNet1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "可售 "
         Height          =   180
         Left            =   3900
         TabIndex        =   15
         Top             =   3390
         Value           =   -1  'True
         Width           =   870
      End
      Begin VB.ComboBox cboBusType 
         Height          =   300
         Left            =   5940
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1410
         Width           =   1035
      End
      Begin MSComCtl2.DTPicker dtpStartupTime 
         Height          =   300
         Left            =   3510
         TabIndex        =   9
         Top             =   2190
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "HH:mm"
         Format          =   123994115
         UpDown          =   -1  'True
         CurrentDate     =   36398
      End
      Begin VB.TextBox txtBusID 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   3510
         MaxLength       =   5
         TabIndex        =   1
         Top             =   1410
         Width           =   1230
      End
      Begin FText.asFlatTextBox txtCheckGate 
         Height          =   300
         Left            =   5940
         TabIndex        =   7
         Top             =   1800
         Width           =   1035
         _ExtentX        =   1826
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
      End
      Begin FText.asFlatTextBox txtRouteID 
         Height          =   300
         Left            =   3510
         TabIndex        =   5
         Top             =   1800
         Width           =   1230
         _ExtentX        =   2170
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
      End
      Begin FText.asFlatTextBox txtVehicleId 
         Height          =   300
         Left            =   3510
         TabIndex        =   13
         Top             =   2970
         Width           =   1230
         _ExtentX        =   2170
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
      End
      Begin VB.Label lblVehicle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "运行车辆(&B):"
         Height          =   180
         Left            =   2400
         TabIndex        =   12
         Top             =   3030
         Width           =   1080
      End
      Begin VB.Label lblScollBusNote1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "分钟"
         Height          =   180
         Left            =   5220
         TabIndex        =   63
         Top             =   2640
         Width           =   360
      End
      Begin VB.Label lblScollBusNote2 
         BackStyle       =   0  'Transparent
         Caption         =   "滚动车次间隔时间(&R):"
         Height          =   195
         Left            =   2400
         TabIndex        =   10
         Top             =   2655
         Width           =   1950
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "可否互联售票(&I):"
         Height          =   240
         Left            =   2400
         TabIndex        =   14
         Top             =   3390
         Width           =   1545
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "车次类型(&B):"
         Height          =   180
         Left            =   4830
         TabIndex        =   2
         Top             =   1470
         Width           =   1080
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "运行线路(&R):"
         Height          =   180
         Left            =   2400
         TabIndex        =   4
         Top             =   1845
         Width           =   1080
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "检票口(&E):"
         Height          =   180
         Left            =   4830
         TabIndex        =   6
         Top             =   1845
         Width           =   900
      End
      Begin VB.Label lblRegularStartTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "发车时间(&T):"
         Height          =   180
         Left            =   2400
         TabIndex        =   8
         Top             =   2265
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车次代码(&C):"
         Height          =   180
         Left            =   2400
         TabIndex        =   0
         Top             =   1470
         Width           =   1080
      End
   End
   Begin RTComctl3.CoolButton cmdHelp 
      Height          =   315
      Left            =   150
      TabIndex        =   47
      Top             =   4170
      Width           =   1155
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
      MICON           =   "frmWizardAddBus.frx":1B23
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
      Left            =   2250
      TabIndex        =   45
      Top             =   4170
      Width           =   1155
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "取消"
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
      MICON           =   "frmWizardAddBus.frx":1B3F
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdFinish 
      Height          =   315
      Left            =   5895
      TabIndex        =   44
      Top             =   4170
      Visible         =   0   'False
      Width           =   1155
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "完成(&F)"
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
      MICON           =   "frmWizardAddBus.frx":1B5B
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdDownStep 
      Default         =   -1  'True
      Height          =   315
      Left            =   4680
      TabIndex        =   43
      Top             =   4170
      Width           =   1155
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "下一步(&N)"
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
      MICON           =   "frmWizardAddBus.frx":1B77
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdUpStep 
      Height          =   315
      Left            =   3465
      TabIndex        =   42
      Top             =   4170
      Width           =   1155
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "上一步(&U)"
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
      MICON           =   "frmWizardAddBus.frx":1B93
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
      BackColor       =   &H00E0E0E0&
      Caption         =   " "
      Enabled         =   0   'False
      Height          =   750
      Left            =   -120
      TabIndex        =   54
      Top             =   3900
      Width           =   8745
   End
   Begin VB.Frame fraStepEnv3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "第三步/共五步"
      Height          =   3900
      Left            =   30
      TabIndex        =   69
      Top             =   0
      Width           =   7095
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   1200
         Index           =   5
         Left            =   2400
         TabIndex        =   70
         Top             =   210
         Width           =   4545
         Begin VB.Line Line2 
            BorderColor     =   &H00FFFFFF&
            Index           =   5
            X1              =   3375
            X2              =   -15
            Y1              =   1125
            Y2              =   1125
         End
         Begin VB.Label lblStep 
            BackStyle       =   0  'Transparent
            Caption         =   "第三步"
            ForeColor       =   &H00808000&
            Height          =   180
            Index           =   3
            Left            =   0
            TabIndex        =   75
            Top             =   0
            Width           =   1170
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            Index           =   3
            X1              =   4350
            X2              =   600
            Y1              =   120
            Y2              =   120
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00FFFFFF&
            Index           =   3
            X1              =   4380
            X2              =   600
            Y1              =   135
            Y2              =   135
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00808080&
            Index           =   3
            X1              =   0
            X2              =   3360
            Y1              =   1110
            Y2              =   1110
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "选择需要生成票价的站点，以便于进行售票处理。在需要生成的站点前打勾进行选中。"
            Height          =   795
            Left            =   0
            TabIndex        =   71
            Top             =   300
            Width           =   4500
         End
      End
      Begin MSComctlLib.ListView lvStationInfo 
         Height          =   1695
         Left            =   2460
         TabIndex        =   40
         Top             =   1710
         Width           =   4410
         _ExtentX        =   7779
         _ExtentY        =   2990
         View            =   2
         LabelEdit       =   1
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ilBus"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "stationID"
            Object.Width           =   4410
         EndProperty
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "站点列表(&T):"
         Height          =   180
         Left            =   2460
         TabIndex        =   39
         Top             =   1470
         Width           =   1080
      End
   End
   Begin VB.Frame fraStep5 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   3900
      Left            =   0
      TabIndex        =   48
      Top             =   0
      Width           =   7095
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   225
         Index           =   4
         Left            =   2430
         TabIndex        =   68
         Top             =   240
         Width           =   4545
         Begin VB.Line Line2 
            BorderColor     =   &H00FFFFFF&
            Index           =   4
            X1              =   4470
            X2              =   840
            Y1              =   135
            Y2              =   135
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            Index           =   4
            X1              =   4470
            X2              =   810
            Y1              =   120
            Y2              =   120
         End
         Begin VB.Label lblStep 
            BackStyle       =   0  'Transparent
            Caption         =   "最后一步"
            ForeColor       =   &H00808000&
            Height          =   180
            Index           =   4
            Left            =   0
            TabIndex        =   67
            Top             =   0
            Width           =   720
         End
      End
      Begin VB.ListBox lstReport 
         Height          =   2040
         Left            =   2430
         TabIndex        =   41
         Top             =   1755
         Width           =   4530
      End
      Begin VB.Label lblReport 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "生成报告(&R):"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   2430
         TabIndex        =   38
         Top             =   1530
         Width           =   1080
      End
      Begin VB.Label lblStartupTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "发车时间:"
         Height          =   180
         Left            =   2430
         TabIndex        =   52
         Top             =   1215
         Width           =   810
      End
      Begin VB.Label lblRoute 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "运行线路:"
         Height          =   180
         Left            =   2430
         TabIndex        =   51
         Top             =   930
         Width           =   810
      End
      Begin VB.Label lblBus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "生成车次:"
         Height          =   210
         Left            =   2430
         TabIndex        =   50
         Top             =   630
         Width           =   810
      End
   End
   Begin VB.Frame fraStep4 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "第四步/共五步"
      Height          =   3900
      Left            =   0
      TabIndex        =   46
      Top             =   0
      Width           =   7095
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   225
         Index           =   3
         Left            =   2400
         TabIndex        =   66
         Top             =   240
         Width           =   4545
         Begin VB.Label lblStep 
            BackStyle       =   0  'Transparent
            Caption         =   "第四步"
            ForeColor       =   &H00808000&
            Height          =   180
            Index           =   2
            Left            =   0
            TabIndex        =   74
            Top             =   0
            Width           =   1170
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            Index           =   2
            X1              =   4350
            X2              =   600
            Y1              =   120
            Y2              =   120
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00FFFFFF&
            Index           =   2
            X1              =   4380
            X2              =   600
            Y1              =   135
            Y2              =   135
         End
      End
      Begin RTComctl3.CoolButton cmdPreview 
         Height          =   315
         Left            =   2400
         TabIndex        =   34
         Top             =   1290
         Width           =   1080
         _ExtentX        =   0
         _ExtentY        =   0
         BTYPE           =   3
         TX              =   "预览(&P)"
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
         MICON           =   "frmWizardAddBus.frx":1BAF
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   315
         Left            =   3510
         TabIndex        =   35
         Top             =   1290
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   556
         _Version        =   393216
         Format          =   123994112
         CurrentDate     =   36462
      End
      Begin VB.ListBox lstVehicle 
         Appearance      =   0  'Flat
         Height          =   1830
         Left            =   2415
         TabIndex        =   37
         Top             =   1680
         Width           =   4575
      End
      Begin MSComCtl2.DTPicker dtpEndDate 
         Height          =   315
         Left            =   5400
         TabIndex        =   36
         Top             =   1290
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   556
         _Version        =   393216
         Format          =   123994112
         CurrentDate     =   36462
      End
      Begin FText.asFlatSpinEdit txtCycleStart 
         Height          =   300
         Left            =   3510
         TabIndex        =   31
         Top             =   900
         Width           =   1005
         _ExtentX        =   1773
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
      End
      Begin FText.asFlatSpinEdit txtCycle 
         Height          =   300
         Left            =   5970
         TabIndex        =   33
         Top             =   900
         Width           =   1005
         _ExtentX        =   1773
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
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "至"
         Height          =   180
         Left            =   5160
         TabIndex        =   53
         Top             =   1350
         Width           =   180
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "运行周期(&Y):"
         Height          =   180
         Left            =   4860
         TabIndex        =   32
         Top             =   960
         Width           =   1080
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "起始序号(&S):"
         Height          =   180
         Left            =   2400
         TabIndex        =   30
         Top             =   960
         Width           =   1080
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "新增车次情况如下:"
         Height          =   180
         Left            =   2400
         TabIndex        =   49
         Top             =   600
         Width           =   1530
      End
   End
   Begin VB.Frame fraStep3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "第三步/共五步"
      Height          =   3900
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   7095
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   1125
         Index           =   2
         Left            =   2400
         TabIndex        =   64
         Top             =   210
         Width           =   4545
         Begin VB.Label lblStep 
            BackStyle       =   0  'Transparent
            Caption         =   "第三步"
            ForeColor       =   &H00808000&
            Height          =   180
            Index           =   1
            Left            =   0
            TabIndex        =   73
            Top             =   0
            Width           =   1170
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            Index           =   7
            X1              =   4350
            X2              =   600
            Y1              =   120
            Y2              =   120
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00FFFFFF&
            Index           =   7
            X1              =   4380
            X2              =   600
            Y1              =   135
            Y2              =   135
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00808080&
            Index           =   2
            X1              =   0
            X2              =   3360
            Y1              =   1110
            Y2              =   1110
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmWizardAddBus.frx":1BCB
            Height          =   765
            Left            =   0
            TabIndex        =   65
            Top             =   300
            Width           =   4500
         End
      End
      Begin MSComctlLib.ListView lvVehicle 
         Height          =   1695
         Left            =   2430
         TabIndex        =   24
         Top             =   1710
         Width           =   3210
         _ExtentX        =   5662
         _ExtentY        =   2990
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "序号"
            Object.Width           =   1323
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "车辆代码"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "车牌"
            Object.Width           =   3528
         EndProperty
      End
      Begin RTComctl3.CoolButton cmdDeleteVehicle 
         Height          =   315
         Left            =   5790
         TabIndex        =   26
         Top             =   2130
         Width           =   1125
         _ExtentX        =   0
         _ExtentY        =   0
         BTYPE           =   3
         TX              =   " 删除(&D)"
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
         MICON           =   "frmWizardAddBus.frx":1C56
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin RTComctl3.CoolButton cmdAddVehicle 
         Height          =   315
         Left            =   5790
         TabIndex        =   25
         Top             =   1740
         Width           =   1110
         _ExtentX        =   0
         _ExtentY        =   0
         BTYPE           =   3
         TX              =   " 新增(&A)"
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
         MICON           =   "frmWizardAddBus.frx":1C72
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblVehicleTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "选中车辆(&H):"
         Height          =   180
         Left            =   2460
         TabIndex        =   23
         Top             =   1470
         Width           =   1080
      End
   End
End
Attribute VB_Name = "frmWizardAddBus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public m_nWizardType As Integer   '向导类型 0-未定义 1-计划车次 2-环境车次
Public m_bIsParent As Boolean


Private Const nfraTop = 240
Private Const nfraLeft = 315
Private byCurrentlySetp As Byte
Private m_oVehicle As New Vehicle
Private m_oREScheme As New REScheme
Private m_oBusInfo As New BaseInfo



Private Sub cboBusType_Click()
   If ResolveDisplay(cboBusType.Text) = TP_ScrollBus Then
'        lblScollBusNote1.Visible = True
'        lblScollBusNote2.Visible = True
        lblRegularStartTime.Caption = "末班时间(&T):"
        txtCheckTime.Enabled = True
        ''支持以后需要修改
    Else
        lblRegularStartTime.Caption = "发车时间(&T):"
        txtCheckTime.Enabled = False
    End If
End Sub


Private Sub chkHavePreSell_Click()
    If chkHavePreSell.Value = vbChecked Then
        dtpEnvBus.Enabled = False
    Else
        dtpEnvBus.Enabled = True
    End If
End Sub


'Private Sub chkScroll_Click()
'    If chkScroll.Value Then
'        dtpStartupTime.Enabled = False
'    Else
'        dtpStartupTime.Enabled = True
'    End If
'End Sub

Private Sub cmdAddVehicle_Click()
On Error GoTo ErrHandle
'    frmQueryVehicle.Show vbModal
'    If frmQueryVehicle.IsCancel Then
'        Unload frmQueryVehicle
'        Exit Sub
'    End If
    Dim oShell As New CommDialog
    oShell.Init g_oActiveUser
    Dim aszTmp() As String
'    With frmQueryVehicle
    '    aszTmp = oShell.SelectVehicle(Trim(.txtVehicle.Text), Trim(ResolveDisplay(.txtCompany.Text)), Trim(ResolveDisplay(.txtBusOwner.Text)), _
                                  Trim(ResolveDisplay(.txtVehicleType.Text)), Trim(.txtLicense.Text), True)
        aszTmp = oShell.SelectVehicleEX(True)
'    End With
    If ArrayLength(aszTmp) = 0 Then Exit Sub
    
    '添加车辆至列表中
    Dim i As Integer
    Dim oListItem As ListItem
    Dim nCount As Integer
    nCount = lvVehicle.ListItems.Count
    For i = 1 To ArrayLength(aszTmp)
        Set oListItem = lvVehicle.ListItems.Add(, , i + nCount)
        oListItem.SubItems(1) = aszTmp(i, 1)
        oListItem.SubItems(2) = aszTmp(i, 2)
    Next i
    cmdDownStep.Enabled = True
    
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDeleteVehicle_Click()
    With lvVehicle
    If .SelectedItem Is Nothing Then Exit Sub
    Dim i As Integer
    For i = .SelectedItem.Index + 1 To .ListItems.Count
        .ListItems(i).Text = Val(.ListItems(i).Text) - 1
    Next i
    .ListItems.Remove .SelectedItem.Index
    If .ListItems.Count = 0 Then cmdDownStep.Enabled = False
    End With
End Sub

Private Sub cmdDownStep_Click()
    byCurrentlySetp = byCurrentlySetp + 1
    Select Case byCurrentlySetp
        Case 0
            ShowStep1
        Case 1
            ShowStep2
        Case 2
            If m_nWizardType = 1 Then
                ShowStep3
            Else
                ShowStepEnv3
            End If
        Case 3
            If m_nWizardType = 1 Then
                ShowStep4
            Else
                ShowStep5
            End If
        Case 4
            ShowStep5
    End Select
End Sub

Private Sub cmdFinish_Click()
    Dim bTmp As Boolean
    If optAddBus.Value = True Then
        bTmp = MakeBus
    Else
        bTmp = MakeEnvBus
    End If
    If bTmp Then
        cmdUpStep.Enabled = False
        cmdDownStep.Enabled = False
        cmdFinish.Enabled = False
        cmdCancel.Caption = "关闭(&C)"
    End If
End Sub

Private Sub cmdHelp_Click()
DisplayHelp Me
End Sub

Private Sub cmdPreview_Click()
    FullCycleVehicle
End Sub

Private Sub cmdUpStep_Click()
    byCurrentlySetp = byCurrentlySetp - 1
    Select Case byCurrentlySetp
        Case 0
            cmdUpStep.Enabled = False
            ShowStep1
        Case 1
            ShowStep2
        Case 2
            If m_nWizardType = 1 Then
                ShowStep3
            Else
                ShowStepEnv3
            End If
        Case 3
            
            ShowStep4
            
        Case 5
            ShowStep5
                
    End Select
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            Unload Me
        Case vbKeyReturn
            SendKeys "{TAB}"
    End Select
End Sub

Private Sub Form_Load()
    On Error GoTo ErrHandle
    AlignFormPos Me
    
    Dim i As Byte
    Dim szBusType() As String
    Dim nCount As Integer
    byCurrentlySetp = 0
    m_oVehicle.Init g_oActiveUser
    m_oREScheme.Init g_oActiveUser
    m_oBusInfo.Init g_oActiveUser
''    g_nPreSell = g_nPreSell
    cmdDownStep.Enabled = True
    
    szBusType = m_oBusInfo.GetAllBusType
    nCount = ArrayLength(szBusType)
    For i = 1 To nCount
        cboBusType.AddItem MakeDisplayString(szBusType(i, 1), szBusType(i, 2))
    Next
    If nCount <> 0 Then
        cboBusType.ListIndex = 0
    End If
    
    ShowStep1
    dtpEnvBus.Value = Date
    If m_nWizardType = 0 Or m_nWizardType = 1 Then
        m_nWizardType = 1
        optAddBus.Value = True
        optAddBus_Click
    Else
        optAddEnvBus.Value = True
        optAddEnvBus_Click
    End If
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Public Sub ShowStep1()
    cmdUpStep.Enabled = False
    cmdDownStep.Enabled = True
    cmdFinish.Visible = False
    fraStep1.Left = 0
    fraStep2.Left = -10000
    fraStep3.Left = -10000
    fraStep4.Left = -10000
    fraStep5.Left = -10000
    fraStepEnv3.Left = -10000
    
    lblStepTitle.Caption = "车次类别"
    lblStepDetail.Caption = "    本步骤用于选择需要新增车次的类别"
End Sub
Public Sub ShowStep2()
    cmdUpStep.Enabled = True
    If Step2CanDown Then
        cmdDownStep.Enabled = True
    Else
        cmdDownStep.Enabled = False
    End If
    cmdFinish.Visible = False
    fraStep1.Left = -10000
    fraStep2.Left = 0
    fraStep3.Left = -10000
    fraStep4.Left = -10000
    fraStep5.Left = -10000
    fraStepEnv3.Left = -10000

    If m_nWizardType = 1 Then        '计划车次
        lblVehicle.Visible = False
        txtVehicleID.Visible = False
    Else                            '环境车次
        lblVehicle.Visible = True
        txtVehicleID.Visible = True
    End If
    
    lblStepTitle.Caption = "车次信息"
    lblStepDetail.Caption = "    本步骤用于登记需要新增的计划车次基本信息"
    
    Call cboBusType_Click
End Sub
Public Sub ShowStep3()
    cmdUpStep.Enabled = True
    If lvVehicle.ListItems.Count = 0 Then
        cmdDownStep.Enabled = False
    Else
        cmdDownStep.Enabled = True
    End If
    cmdFinish.Visible = False
    fraStep1.Left = -10000
    fraStep2.Left = -10000
    fraStep3.Left = 0
    fraStep4.Left = -10000
    fraStep5.Left = -10000
    fraStepEnv3.Left = -10000

    lblStepTitle.Caption = "车辆信息"
    lblStepDetail.Caption = "    本步骤用于新增车次的运行车辆"
    lblVehicleTitle.Caption = "车次[" & txtBusID.Text & "]的运行车辆(&L):"

    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Public Sub ShowStepEnv3()
    cmdUpStep.Enabled = True
    If lvStationInfo.ListItems.Count = 0 Then
        cmdDownStep.Enabled = False
    Else
        cmdDownStep.Enabled = True
    End If
    cmdFinish.Visible = False
    fraStep1.Left = -10000
    fraStep2.Left = -10000
    fraStep3.Left = -10000
    fraStep4.Left = -10000
    fraStep5.Left = -10000
    fraStepEnv3.Left = 0
    
    lblStepTitle.Caption = "选择站点"
    lblStepDetail.Caption = "    本步骤选择需要生成票价的站点"
End Sub
Public Sub ShowStep4()
    cmdUpStep.Enabled = True
    cmdDownStep.Enabled = True
    cmdFinish.Visible = False
    fraStep1.Left = -10000
    fraStep2.Left = -10000
    fraStep3.Left = -10000
    fraStep4.Left = 0
    fraStep5.Left = -10000
    fraStepEnv3.Left = -10000

    dtpStartDate.Value = Date
    dtpEndDate.Value = DateAdd("d", g_nPreSell, Date)
    
    lblStepTitle.Caption = "车辆排班"
    lblStepDetail.Caption = "    本步骤用于对车次的运行车辆进行安排"

End Sub

Public Sub ShowStep5()
    cmdUpStep.Enabled = True
    cmdDownStep.Enabled = False
    cmdFinish.Visible = True
    fraStep1.Left = -10000
    fraStep2.Left = -10000
    fraStep3.Left = -10000
    fraStep4.Left = -10000
    fraStep5.Left = 0
    fraStepEnv3.Left = -10000

    lblBus.Caption = "车次代码:" & txtBusID.Text
    lblRoute.Caption = "线路代码:" & txtRouteID.Text
    lblStartupTime.Caption = IIf(ResolveDisplay(cboBusType.Text) = TP_ScrollBus, "滚动车次间隔时间:" & txtCheckTime.Text, "发车时间:" & Format(dtpStartupTime.Value, "hh:mm"))
    lblReport.Caption = "生成报告预览(&R):"
    MakeReport
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m_oVehicle = Nothing
    Set m_oREScheme = Nothing
    Set m_oBusInfo = Nothing

    SaveFormPos Me
End Sub

Private Sub optAddBus_Click()
    chkRE.Enabled = True
    dtpEnvBus.Enabled = False
    chkHavePreSell.Enabled = False
    m_nWizardType = 1
End Sub

Private Sub optAddEnvBus_Click()
    chkRE.Enabled = False
    If chkHavePreSell.Value = vbChecked Then
        dtpEnvBus.Enabled = False
    Else
        dtpEnvBus.Enabled = True
    End If
    chkHavePreSell.Enabled = True
    m_nWizardType = 2
End Sub

Private Sub txtBusId_Change()
    cmdDownStep.Enabled = Step2CanDown
End Sub

Private Sub txtCheckGate_Change()
    cmdDownStep.Enabled = Step2CanDown
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


Private Sub txtRouteID_Change()
    cmdDownStep.Enabled = Step2CanDown
End Sub



Private Function Step2CanDown() As Boolean
    If Trim(txtRouteID.Text) = "" Or Trim(txtCheckGate.Text) = "" Or Trim(txtBusID.Text) = "" Then
        Step2CanDown = False
    Else
        Step2CanDown = True
    End If
End Function


Public Sub FullCycleVehicle()
    Dim i As Integer
    Dim nSerial As Integer
    Dim nCount As Integer
    On Error GoTo ErrHandle
    nCount = DateDiff("d", dtpStartDate.Value, dtpEndDate.Value)
    lstVehicle.Clear
    For i = 0 To nCount
        nSerial = m_oREScheme.GetExecuteVehicleSerialNo(Val(txtCycle.Text), Val(txtCycleStart.Text), DateAdd("d", i, dtpStartDate.Value))
        lstVehicle.AddItem Format(DateAdd("d", i, dtpStartDate.Value), "YYYY年MM月DD日") & "  序号:" & lvVehicle.ListItems(nSerial).Text & "  车辆:" & lvVehicle.ListItems(nSerial).ListSubItems(1).Text & "  车牌:" & lvVehicle.ListItems(nSerial).ListSubItems(2).Text
NextStep:
    Next
    Exit Sub
ErrHandle:
    If err.Number = 35600 Then
        lstVehicle.AddItem Format(DateAdd("d", i, dtpStartDate.Value), "YYYY年MM月DD日") & "  无运行车辆"
        Resume NextStep
    End If
End Sub

Private Sub MakeReport()
    Dim i As Integer
    Dim dtmake As Date
    lblReport.ForeColor = vbBlack
    lstReport.Clear
    If (optAddBus.Value And chkRE.Value = vbChecked) Or (optAddEnvBus.Value And chkHavePreSell.Value = vbChecked) Then
        For i = 1 To g_nPreSell
            dtmake = DateAdd("d", i - 1, Date)
            lstReport.AddItem "环境中将生成:" & Format(dtmake, "YYYY年MM月DD日") & "的车次"
        Next
    End If
    If optAddBus.Value Then
        lstReport.AddItem "计划中生成车次" & txtBusID.Text & "..."
    Else
        If optAddEnvBus.Value And chkHavePreSell.Value <> vbChecked Then
            lstReport.AddItem "环境中生成" & Format(dtpEnvBus.Value, "YYYY年MM月DD日") & "的车次[" & txtBusID.Text & "]..."
        End If
    End If
End Sub

Private Function MakeBus() As Boolean
On Error GoTo ErrHandle
    Dim oBus As New Bus
    Dim tBusVehicle As TBusVehicleInfo
    Dim dtReMake As Date
    Dim szMsgShow As String
    Dim bBusAddFinished As Boolean
    Dim i As Integer
    Dim atVehicleSeat() As TBusVehicleSeatType
    Dim oTicketPriceMan As New TicketPriceMan
    Dim aszBusID(1 To 1) As String
    Dim oRoutePriceTable As New RoutePriceTable
    
    
    lblReport.Caption = "新增计划车次结果(&R):"
    SetBusy
    oBus.Init g_oActiveUser
    lstReport.Clear
    lstReport.AddItem "新增车次" & EncodeString(txtBusID.Text) & "开始"
    lstReport.AddItem "-------------------------------------------------------------"
    
    oBus.AddNew
    oBus.BusID = txtBusID.Text
    oBus.CheckGate = ResolveDisplay(txtCheckGate.Text)
    oBus.CycleStartSerialNo = Val(txtCycleStart.Text)
    oBus.RunCycle = Val(txtCycle.Text)
  '  oBus.ProjectID = g_szExePlanID
    oBus.Route = ResolveDisplay(txtRouteID.Text)
    oBus.StartUpTime = dtpStartupTime.Value
    oBus.BusType = ResolveDisplay(cboBusType.Text)
    ''以后加入始班车和末班车的处理
    If oBus.BusType <> TP_ScrollBus Then
        oBus.ScrollBusCheckTime = 0
    Else
        oBus.ScrollBusCheckTime = CInt(Val(txtCheckTime.Text))
    End If
    If Optinternet1.Value = True Then
        oBus.InternetStatus = CnInternetCanSell
    Else
        oBus.InternetStatus = CnInternetNotCanSell
    End If
    
    oBus.Update
    bBusAddFinished = True
    
    On Error Resume Next
    '添加车辆
    For i = 1 To lvVehicle.ListItems.Count
        tBusVehicle.nSerialNo = Val(lvVehicle.ListItems.Item(i).Text)
        tBusVehicle.szVehicleID = Trim(lvVehicle.ListItems.Item(i).ListSubItems(1).Text)
        tBusVehicle.nStandTicketCount = 0
        tBusVehicle.dtBeginStopDate = CDate(cszEmptyDateStr)
        tBusVehicle.dtEndStopDate = CDate(cszEmptyDateStr)
        oBus.AddRunVehicle tBusVehicle
        If err Then
            lstReport.AddItem "计划中车次[" & txtBusID.Text & "]新增车次车辆失败!错误号[" & err.Number & "]:" & err.Description
        Else
            lstReport.AddItem "计划中车次[" & txtBusID.Text & "]新增车次车辆" & tBusVehicle.nSerialNo & EncodeString(tBusVehicle.szVehicleID)
        End If
    Next
    
    '自动添加票价
    oTicketPriceMan.Init g_oActiveUser
    aszBusID(1) = txtBusID.Text
    atVehicleSeat = oTicketPriceMan.GetAllBusVehicleTypeSeatType(aszBusID)
    oRoutePriceTable.Init g_oActiveUser
    oRoutePriceTable.Identify g_szExePriceTable
    oRoutePriceTable.MakeBusPrice atVehicleSeat
    '    If err Then
    '        lstReport.AddItem "计划中车次[" & txtBusID.Text & "]新增车次票价失败!错误号[" & err.Number & "]:" & err.Description
    '    Else
    '        lstReport.AddItem "计划中车次[" & txtBusID.Text & "]新增票价成功!"
    '    End If
    
    '生成环境
    If chkRE.Value = vbChecked Then
        For i = 1 To g_nPreSell
            dtReMake = DateAdd("d", i - 1, Date)
            m_oREScheme.MakeRunEvironment dtReMake, txtBusID.Text
            If err Then
                lstReport.AddItem "错误号[" & err.Number & "]:" & Format(dtReMake, "YYYY年MM月DD日") & " 车次[" & txtBusID.Text & "]" & err.Description
            Else
                lstReport.AddItem "环境中新增" & Format(dtReMake, "YYYY年MM月DD日") & "的车次[" & txtBusID.Text & "]成功..."
            End If
        Next i
    End If
    SetNormal
    lstReport.AddItem "-------------------------------------------------------------"
    lstReport.AddItem "新增车次完成!"
    MsgBox "计划车次" & EncodeString(txtBusID.Text) & "新增成功", vbInformation + vbOKOnly, "信息"
    If m_bIsParent Then
        '将新增的车次刷新出来
'        If m_nWizardType = 1 Then
            frmBus.AddList txtBusID.Text
'        Else
'            frmEnvBus.AddList txtBusID.Text, dtpEnvBus.Value
'        End If
    End If
    MakeBus = True
    Exit Function
ErrHandle:
    szMsgShow = "新增计划车次" & EncodeString(txtBusID.Text) & "失败!错误号[" & err.Number & "]:" & err.Description
    lstReport.AddItem "-------------------------------------------------------------"
    lstReport.AddItem "新增车次中止!"
    SetNormal
    MsgBox szMsgShow, vbExclamation, "错误"
    MakeBus = False
    Exit Function
End Function
'新增环境车次
Private Function MakeEnvBus() As Boolean
On Error GoTo ErrHandle
    lblReport.Caption = "新增环境车次结果(&R):"
    Dim oRebus As New REBus
    Dim dtReMake As Date, nCount As Integer
    oRebus.Init g_oActiveUser
    
    If chkHavePreSell.Value = vbChecked Then        '包括所有预售天数
        dtReMake = Date
        nCount = g_nPreSell
    Else                                            '指定日期
        dtReMake = dtpEnvBus.Value
        nCount = 1
    End If
    
    lstReport.Clear
    lstReport.AddItem "新增车次" & EncodeString(txtBusID.Text) & "开始"
    lstReport.AddItem "-------------------------------------------------------------"

    Dim i As Integer
    MousePointer = vbHourglass
    Dim oREScheme As New REScheme
    Dim oRoutePrice As New RoutePriceTable
    Dim oRegularScheme As New RegularScheme
    Dim aszStationInfo() As String
    Dim szPriceTable As String
    Dim tBusPrice() As TBusPriceDetailInfo
    
    oRoutePrice.Init g_oActiveUser
    aszStationInfo = GetStationItem
    oRegularScheme.Init g_oActiveUser
    
    
    oREScheme.Init g_oActiveUser
    For i = 1 To nCount
        dtReMake = DateAdd("d", i - 1, dtReMake)
    
        If oREScheme.IdentifyNotBusId(txtBusID.Text, dtReMake, g_szExePriceTable) = False Then
            lstReport.AddItem Format(dtReMake, "YYYY年MM月DD日") & "的车次" & EncodeString(txtBusID.Text) & "在环境或计划中已存在，不能新增!"
            GoTo LoopNext
        End If

        szPriceTable = oRegularScheme.GetRunPriceTableEx(dtReMake)
        oRoutePrice.Identify szPriceTable
        tBusPrice = oRoutePrice.MakeEnvBusPrice(ResolveDisplay(txtRouteID.Text), txtBusID.Text, ResolveDisplay(txtVehicleID.Text), dtpStartupTime.Value, ResolveDisplay(txtCheckGate.Text), aszStationInfo)
        If ArrayLength(tBusPrice) = 0 Then
            lstReport.AddItem Format(dtReMake, "YYYY年MM月DD日") & "的车次" & EncodeString(txtBusID.Text) & "没有生成票价!"
            GoTo LoopNext
        End If
        
        '支持始发班和末班车时，应更改
        oREScheme.AddEnviromentBus dtReMake, dtpStartupTime.Value, Trim(txtBusID.Text), ResolveDisplay(txtVehicleID.Text), ResolveDisplay(txtCheckGate.Text), ResolveDisplay(cboBusType.Text), IIf(Optinternet1.Value, 1, 0), False, Val(txtCheckTime.Text), tBusPrice
        
        lstReport.AddItem "环境中新增" & Format(dtReMake, "YYYY年MM月DD日") & "的车次成功..."
LoopNext:
    Next i

    lstReport.AddItem "-------------------------------------------------------------"
    lstReport.AddItem "新增车次完成!"
    If m_bIsParent Then
        '将新增的车次刷新出来
        frmEnvBus.AddList txtBusID.Text, dtpEnvBus.Value
    End If
    SetNormal
    
    MsgBox "环境车次" & EncodeString(txtBusID.Text) & "新增成功", vbInformation + vbOKOnly, "信息"
    MakeEnvBus = True
    Exit Function
ErrHandle:
    SetNormal
    lstReport.AddItem "错误号[" & err.Number & "]" & Format(dtReMake, "YYYY年MM月DD日") & "车次[" & txtBusID.Text & "]" & err.Description
    lstReport.AddItem "-------------------------------------------------------------"
    lstReport.AddItem "新增车次中止!"
    MsgBox "新增车次失败!", vbExclamation, "提示"
    MakeEnvBus = False
End Function
Private Sub txtRouteID_ButtonClick()
On Error GoTo ErrHandle
    Dim oShell As New CommDialog
    Dim aszTmp() As String
    oShell.Init g_oActiveUser
    aszTmp = oShell.SelectRoute
    Set oShell = Nothing
    If ArrayLength(aszTmp) = 0 Then Exit Sub
    txtRouteID.Text = MakeDisplayString(aszTmp(1, 1), aszTmp(1, 2))
    
    
    If optAddEnvBus.Value Then      '环境车次
        Dim oRoute As Route
        Set oRoute = New Route
        oRoute.Init g_oActiveUser
        oRoute.Identify aszTmp(1, 1)
        Dim szStationInfo() As String
        szStationInfo = oRoute.RouteStationEx
        AddStationItems szStationInfo
        Set oRoute = Nothing
    End If
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub
Private Sub AddStationItems(szSationInfo() As String)
    Dim nCount As Integer
    Dim i As Integer
    Dim oListItem As ListItem
    Dim szTemp As String
    nCount = ArrayLength(szSationInfo)
    lvStationInfo.ListItems.Clear
    lvStationInfo.Sorted = False
    For i = 1 To nCount
        szTemp = Trim(szSationInfo(i, 1)) & "[" & Trim(szSationInfo(i, 3)) & "]"
        Set oListItem = lvStationInfo.ListItems.Add(, , szTemp)
        
        If i = nCount Then oListItem.Checked = True
    Next
End Sub
Private Function GetStationItem() As String()
    Dim i As Integer, nCount As Integer
    Dim szTemp() As String
    
    For i = 1 To lvStationInfo.ListItems.Count
        If lvStationInfo.ListItems(i).Checked = True Then
        nCount = ArrayLength(szTemp) + 1
        ReDim Preserve szTemp(1 To nCount)
        szTemp(nCount) = ResolveDisplay(lvStationInfo.ListItems(i).Text)
        End If
    Next
    
    GetStationItem = szTemp

End Function

Private Sub txtVehicleId_ButtonClick()
'    frmQueryVehicle.Show vbModal
'    If frmQueryVehicle.IsCancel Then
'        Unload frmQueryVehicle
'        Exit Sub
'    End If
    Dim oShell As New CommDialog
    oShell.Init g_oActiveUser
    Dim aszTmp() As String
'    With frmQueryVehicle
'        aszTmp = oShell.SelectVehicle(Trim(.txtVehicle.Text), Trim(ResolveDisplay(.txtBusOwner.Text)), Trim(ResolveDisplay(.txtBusOwner.Text)), _
'                                      Trim(.txtLicense.Text), Trim(ResolveDisplay(.txtVehicleType.Text)), False)
'    End With
    aszTmp = oShell.SelectVehicleEX(False)
    If ArrayLength(aszTmp) = 0 Then Exit Sub
    txtVehicleID.Text = MakeDisplayString(aszTmp(1, 1), aszTmp(1, 2))
End Sub

