VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmArrangeEnvBus 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "环境车次"
   ClientHeight    =   4515
   ClientLeft      =   2130
   ClientTop       =   2655
   ClientWidth     =   7605
   HelpContextID   =   10000230
   Icon            =   "frmArrangeEnvBus.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   7605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Visible         =   0   'False
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "是否晚点"
      Height          =   570
      Left            =   195
      TabIndex        =   50
      Top             =   3240
      Width           =   2835
      Begin VB.OptionButton optNormal 
         BackColor       =   &H00E0E0E0&
         Caption         =   "正常(&N)"
         Height          =   195
         Left            =   1380
         TabIndex        =   52
         Top             =   255
         Width           =   1320
      End
      Begin VB.OptionButton optDelay 
         BackColor       =   &H00E0E0E0&
         Caption         =   "晚点(&D)"
         Height          =   180
         Left            =   150
         TabIndex        =   51
         Top             =   255
         Width           =   1350
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "是否全额退票"
      Height          =   900
      Left            =   195
      TabIndex        =   47
      Top             =   2265
      Width           =   2835
      Begin VB.OptionButton optAllRe 
         BackColor       =   &H00E0E0E0&
         Caption         =   "全额退票(&R)"
         Height          =   180
         Left            =   150
         TabIndex        =   49
         Top             =   255
         Width           =   1350
      End
      Begin VB.OptionButton optNoR 
         BackColor       =   &H00E0E0E0&
         Caption         =   "不全额退票(&N)"
         Height          =   195
         Left            =   150
         TabIndex        =   48
         Top             =   555
         Width           =   1560
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "互联售票(I)"
      Height          =   900
      Left            =   3180
      TabIndex        =   44
      Top             =   2265
      Width           =   2850
      Begin VB.OptionButton Optinternet2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "不可售"
         Height          =   195
         Left            =   165
         TabIndex        =   46
         Top             =   540
         Value           =   -1  'True
         Width           =   960
      End
      Begin VB.OptionButton Optinternet1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "可售"
         Height          =   195
         Left            =   165
         TabIndex        =   45
         Top             =   240
         Width           =   735
      End
   End
   Begin RTComctl3.CoolButton cmdSellStation 
      Height          =   375
      Left            =   5685
      TabIndex        =   32
      Top             =   3975
      Width           =   1125
      _ExtentX        =   1984
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
      MICON           =   "frmArrangeEnvBus.frx":038A
      PICN            =   "frmArrangeEnvBus.frx":03A6
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
      Height          =   360
      Left            =   6210
      TabIndex        =   25
      Top             =   540
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   635
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
      MICON           =   "frmArrangeEnvBus.frx":0740
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame fraVehicle 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "车辆属性"
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   540
      TabIndex        =   33
      Top             =   6060
      Width           =   7275
      Begin VB.Label lblBusOwnerA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车主:"
         Height          =   180
         Left            =   1980
         TabIndex        =   37
         Top             =   315
         Width           =   450
      End
      Begin VB.Label lblBusTypeA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车型:"
         Height          =   180
         Left            =   555
         TabIndex        =   36
         Top             =   315
         Width           =   450
      End
      Begin VB.Image imgStop 
         Height          =   240
         Left            =   105
         Picture         =   "frmArrangeEnvBus.frx":075C
         Top             =   270
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgRun 
         Height          =   240
         Left            =   105
         Picture         =   "frmArrangeEnvBus.frx":0AE6
         Top             =   270
         Width           =   240
      End
      Begin VB.Label lblBusSeatA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "座位数:"
         Height          =   180
         Left            =   3240
         TabIndex        =   35
         Top             =   315
         Width           =   630
      End
      Begin VB.Label lblCorpA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "参运公司:"
         Height          =   180
         Left            =   4770
         TabIndex        =   34
         Top             =   315
         Width           =   810
      End
   End
   Begin FText.asFlatTextBox txtCheckGate 
      Height          =   300
      Left            =   4320
      TabIndex        =   7
      Top             =   510
      Width           =   1680
      _ExtentX        =   2963
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
   Begin VB.ComboBox cboBusType 
      Height          =   300
      Left            =   4320
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox txtBusId 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1335
      MaxLength       =   5
      TabIndex        =   1
      Top             =   120
      Width           =   1680
   End
   Begin RTComctl3.CoolButton cmdOK 
      Default         =   -1  'True
      Height          =   360
      Left            =   6210
      TabIndex        =   24
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   635
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
      MICON           =   "frmArrangeEnvBus.frx":1728
      PICN            =   "frmArrangeEnvBus.frx":1744
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdHelp 
      Height          =   360
      Left            =   6210
      TabIndex        =   26
      Top             =   960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   635
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
      MICON           =   "frmArrangeEnvBus.frx":1ADE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker dtpStartupTime 
      Height          =   300
      Left            =   4320
      TabIndex        =   11
      Top             =   900
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "HH:mm"
      Format          =   92930051
      UpDown          =   -1  'True
      CurrentDate     =   36396
   End
   Begin MSComCtl2.DTPicker dtpOffDate 
      Height          =   300
      Left            =   1335
      TabIndex        =   9
      Top             =   900
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   529
      _Version        =   393216
      Format          =   92930048
      CurrentDate     =   36396
   End
   Begin FText.asFlatTextBox txtRouteID 
      Height          =   300
      Left            =   1335
      TabIndex        =   5
      Top             =   510
      Width           =   1680
      _ExtentX        =   2963
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
   Begin MSComCtl2.DTPicker dtpFirstBus 
      Height          =   300
      Left            =   2430
      TabIndex        =   13
      Top             =   4815
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
      CustomFormat    =   "HH:mm"
      Format          =   92930051
      UpDown          =   -1  'True
      CurrentDate     =   36392
   End
   Begin MSComCtl2.DTPicker dtpLastBus 
      Height          =   300
      Left            =   5415
      TabIndex        =   15
      Top             =   4815
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
      CustomFormat    =   "HH:mm"
      Format          =   92930051
      UpDown          =   -1  'True
      CurrentDate     =   36392
   End
   Begin FText.asFlatSpinEdit txtCheckTime 
      Height          =   300
      Left            =   3330
      TabIndex        =   17
      Top             =   5190
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
   Begin FText.asFlatTextBox txtOwner 
      Height          =   315
      Left            =   4320
      TabIndex        =   23
      Top             =   1785
      Width           =   1695
      _ExtentX        =   2990
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
   Begin FText.asFlatTextBox txtSplitCompanyID 
      Height          =   315
      Left            =   1335
      TabIndex        =   21
      Top             =   1830
      Width           =   1680
      _ExtentX        =   2963
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
   Begin RTComctl3.CoolButton imbSeatLeave 
      Height          =   375
      Left            =   1635
      TabIndex        =   29
      Top             =   3975
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "座位更改"
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
      MICON           =   "frmArrangeEnvBus.frx":1AFA
      PICN            =   "frmArrangeEnvBus.frx":1B16
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
      Left            =   2850
      TabIndex        =   30
      Top             =   3975
      Width           =   1125
      _ExtentX        =   1984
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
      MICON           =   "frmArrangeEnvBus.frx":20B0
      PICN            =   "frmArrangeEnvBus.frx":20CC
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
      Left            =   4095
      TabIndex        =   31
      Top             =   3975
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "站点设置"
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
      MICON           =   "frmArrangeEnvBus.frx":2226
      PICN            =   "frmArrangeEnvBus.frx":2242
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton imbCheckInfo 
      Height          =   360
      Left            =   6210
      TabIndex        =   27
      Top             =   2355
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   635
      BTYPE           =   3
      TX              =   "检票信息"
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
      MICON           =   "frmArrangeEnvBus.frx":239C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin FText.asFlatTextBox txtVehicleId 
      Height          =   300
      Left            =   1335
      TabIndex        =   19
      Top             =   1380
      Width           =   1680
      _ExtentX        =   2963
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
   Begin RTComctl3.CoolButton cmdAllot 
      Height          =   375
      Left            =   375
      TabIndex        =   28
      Top             =   3975
      Width           =   1125
      _ExtentX        =   1984
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
      MICON           =   "frmArrangeEnvBus.frx":23B8
      PICN            =   "frmArrangeEnvBus.frx":23D4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ListView lvStationInfo 
      Height          =   1680
      Left            =   525
      TabIndex        =   41
      Top             =   7170
      Width           =   7260
      _ExtentX        =   12806
      _ExtentY        =   2963
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
   Begin VB.Label lblTotalSeat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "总座位数:"
      Height          =   180
      Left            =   4620
      TabIndex        =   43
      Top             =   1440
      Width           =   810
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "请选择需要生成票价的站点(&L):"
      Height          =   180
      Left            =   630
      TabIndex        =   42
      Top             =   6915
      Width           =   2520
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "拆帐公司(&I):"
      Height          =   180
      Left            =   195
      TabIndex        =   20
      Top             =   1890
      Width           =   1080
   End
   Begin VB.Label lblOwner 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车主(&W):"
      Height          =   180
      Left            =   3180
      TabIndex        =   22
      Top             =   1845
      Width           =   720
   End
   Begin VB.Label lblSellMoney 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "已售票款:"
      Height          =   180
      Left            =   3750
      TabIndex        =   40
      Top             =   5790
      Width           =   810
   End
   Begin VB.Label lblSellSeat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "已售座数:"
      Height          =   180
      Left            =   3180
      TabIndex        =   39
      Top             =   1440
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "运行车辆(&B):"
      Height          =   180
      Left            =   195
      TabIndex        =   18
      Top             =   1440
      Width           =   1080
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "末班车(&T):"
      Enabled         =   0   'False
      Height          =   180
      Left            =   4275
      TabIndex        =   14
      Top             =   4845
      Width           =   900
   End
   Begin VB.Label lblOffTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "发车时间(&T):"
      Height          =   180
      Left            =   3180
      TabIndex        =   10
      Top             =   945
      Width           =   1170
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "始班车(&T):"
      Enabled         =   0   'False
      Height          =   180
      Left            =   1290
      TabIndex        =   12
      Top             =   4845
      Width           =   900
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "请选取站点以便生成票价："
      Height          =   240
      Left            =   615
      TabIndex        =   38
      Top             =   5775
      Width           =   2220
   End
   Begin VB.Label lblScroll 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "滚动班次间隔时间(&F):"
      Height          =   180
      Left            =   1290
      TabIndex        =   16
      Top             =   5265
      Width           =   1800
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "车次类型(&M):"
      Height          =   195
      Left            =   3180
      TabIndex        =   2
      Top             =   180
      Width           =   1095
   End
   Begin VB.Label lblOffDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车次日期(&D):"
      Height          =   180
      Left            =   195
      TabIndex        =   8
      Top             =   945
      Width           =   1080
   End
   Begin VB.Label lblCheckGate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "检票口(&C):"
      Height          =   180
      Left            =   3180
      TabIndex        =   6
      Top             =   570
      Width           =   990
   End
   Begin VB.Label lblRoute 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "运行线路(&R):"
      Height          =   180
      Left            =   195
      TabIndex        =   4
      Top             =   570
      Width           =   1080
   End
   Begin VB.Label lblBusId 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车次代码(&I):"
      Height          =   180
      Left            =   195
      TabIndex        =   0
      Top             =   180
      Width           =   1080
   End
End
Attribute VB_Name = "frmArrangeEnvBus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 1
Option Explicit

Public Status As EFormStatus

Public m_szBusID As String
Public m_dtBusDate As Date

Private moReSheme As New REScheme
Private m_oREBus As New REBus
Private m_bChanged As Boolean

Private Sub cboBusType_Click()
    If ResolveDisplay(cboBusType.Text) = TP_ScrollBus Then
'       lblScroll.Visible = True
       txtCheckTime.Enabled = True
       ''支持后需要更改
       dtpFirstBus.Enabled = False
       dtpLastBus.Enabled = False
       
'       lblOffTime.Visible = False
       dtpStartupTime.Enabled = True
    Else
'       lblScroll.Visible = False
       txtCheckTime.Enabled = False
       dtpFirstBus.Enabled = False
       dtpLastBus.Enabled = False
    
'       lblOffTime.Visible = True
       dtpStartupTime.Enabled = True
    End If
    AddEnabled
End Sub
Public Sub RefreshBus()
    If Status = EFS_AddNew Then
        txtBusID.Text = ""
        cboBusType.ListIndex = 0
        txtRouteID.Text = ""
        dtpOffDate.Value = Date
        txtVehicleID.Text = ""
        OptInterNet1.Value = True
        lvStationInfo.Enabled = True
    Else
        m_oREBus.Identify m_szBusID, m_dtBusDate
        txtBusID.Text = m_szBusID
        cboBusType.ListIndex = SeekListIndex(cboBusType, MakeDisplayString(m_oREBus.BusType, m_oREBus.BusTypeName))
        txtRouteID.Text = MakeDisplayString(m_oREBus.Route, m_oREBus.RouteName)
        txtCheckGate.Text = m_oREBus.CheckGate
        dtpOffDate.Value = m_dtBusDate
'        If m_oReBus.BusType = TP_ScrollBus Then
            txtCheckTime.Value = m_oREBus.ScrollBusCheckTime
'        Else
            dtpStartupTime.Value = m_oREBus.StartUpTime
'        End If
        txtVehicleID.Text = MakeDisplayString(m_oREBus.Vehicle, m_oREBus.VehicleTag)
        If m_oREBus.InternetStatus = CnInternetCanSell Then
            OptInterNet1.Value = True
        Else
            OptInterNet2.Value = True
        End If
        
        If m_oREBus.AllRefundment Then
            optAllRe.Value = True
        Else
            optNoR.Value = True
        End If
        '是否晚点
    If m_oREBus.DelayStatus Then
        optDelay.Value = True
    Else
        optNormal.Value = True
    End If
        lblSellSeat.Caption = "已售座数:" & m_oREBus.SaledSeatCount
        lblTotalSeat.Caption = "总座位数:" & m_oREBus.TotalSeat
        
        '需要补充已售票款
        lblSellMoney.Caption = "已售票款:"
        
        txtOwner.Text = MakeDisplayString(m_oREBus.Owner, m_oREBus.OwnerName)
        txtSplitCompanyID.Text = MakeDisplayString(m_oREBus.SplitCompanyID, m_oREBus.SplitCompanyName)
    End If
    RefreshVehicle
    m_bChanged = False
End Sub
Private Sub LayoutForm()
    If Status = EFS_AddNew Then
        Me.Caption = "环境车次—新增"
        cmdOk.Caption = "新增(&A)"
        txtBusID.Enabled = True
        cboBusType.Enabled = True
        dtpOffDate.Enabled = True
'        fraStation.Visible = True
'        fraEditSet.Visible = False
'        fraStation.Visible = False
'        fraEditSet.Visible = True
        
        
        imbCheckInfo.Visible = False
        Me.Height = 3570
        
    Else
        Me.Caption = "环境车次—编辑"
        cmdOk.Caption = "保存(&S)"
        cmdOk.Default = True
        txtBusID.Enabled = False
        dtpOffDate.Enabled = False
        cboBusType.Enabled = True 'False
'        fraStation.Visible = False
'        fraEditSet.Visible = True
    End If
End Sub

Private Sub cmdAllot_Click()
    frmEnvBusAllot.m_bIsAllot = True
    frmEnvBusAllot.m_szBusID = m_szBusID
    frmEnvBusAllot.m_dtEnvDate = m_dtBusDate
    frmEnvBusAllot.Show vbModal
    
    '刷新检票口及发车时间信息
    m_oREBus.Identify m_szBusID, m_dtBusDate
    txtCheckGate.Text = m_oREBus.CheckGate
    dtpStartupTime.Value = m_oREBus.StartUpTime
    
End Sub

Private Sub cmdHelp_Click()
DisplayHelp Me
End Sub

Private Sub cmdOk_Click()
On Error GoTo ErrHandle
    SetBusy
    Dim szMsg As String
    Dim nOldStatus As Integer
    nOldStatus = Status
    If Status = EFS_AddNew Then
        szMsg = AddNewEnvBus
    Else
        szMsg = ModifyBus
    End If
    SetNormal
    If szMsg <> "" Then
        MsgBox szMsg, vbExclamation, "错误"
    Else
        If nOldStatus = EFS_AddNew Then
            frmEnvBus.AddList m_szBusID, m_dtBusDate
        Else
            frmEnvBus.UpdateList m_szBusID, m_dtBusDate
        End If
        MsgBox "车次处理成功!", vbInformation, "信息"
        Unload Me
    End If
    
    Exit Sub
ErrHandle:
    SetNormal
    ShowErrorMsg
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub



Private Sub cmdSellStation_Click()
    frmEnvBusAllot.m_bIsAllot = False
    frmEnvBusAllot.m_szBusID = m_szBusID
    frmEnvBusAllot.m_dtEnvDate = m_dtBusDate
    frmEnvBusAllot.Show vbModal
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
           SendKeys "{TAB}"
    End Select
End Sub

Private Sub Form_Load()
    On Error GoTo ErrHandle
    AlignFormPos Me
    
    m_oREBus.Init g_oActiveUser
    Dim oBaseInfo As New BaseInfo
    oBaseInfo.Init g_oActiveUser
    Dim aszTmp() As String
    aszTmp = oBaseInfo.GetAllBusType()
    cboBusType.Clear
    Dim i As Integer
    For i = 1 To ArrayLength(aszTmp)
        cboBusType.AddItem MakeDisplayString(aszTmp(i, 1), aszTmp(i, 2))
    Next i
    
    RefreshBus
    LayoutForm
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveFormPos Me
    Set m_oREBus = Nothing
    Set moReSheme = Nothing
End Sub


Private Sub imbCheckInfo_Click()
    On Error GoTo ErrHandle
    Dim szbusID As String, dtEnvDate As Date
    szbusID = m_szBusID
    dtEnvDate = m_dtBusDate
    
    Dim oCheckSheet As New STShell.CommDialog
    oCheckSheet.Init g_oActiveUser
    If ResolveDisplay(cboBusType.Text) = TP_ScrollBus Then
        oCheckSheet.ShowEnvScrollBusList dtEnvDate, szbusID
        
    Else
        oCheckSheet.ShowCheckInfo dtEnvDate, szbusID
    End If
    
    Set oCheckSheet = Nothing
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub imbSeatLeave_Click()
    frmEnvReserveSeat.m_szBusID = m_szBusID
    frmEnvReserveSeat.m_dtEnvDate = m_dtBusDate
    frmEnvReserveSeat.Show vbModal
    
    m_oREBus.Identify m_szBusID, m_dtBusDate

    lblSellSeat.Caption = "已售座数:" & m_oREBus.SaledSeatCount
    lblTotalSeat.Caption = "总座位数:" & m_oREBus.TotalSeat
End Sub

Private Sub imbStationSet_Click()
    frmEnvBusStation.m_dtRunDate = m_dtBusDate
    frmEnvBusStation.m_szBusID = m_szBusID
    frmEnvBusStation.Show vbModal
    
    
    
'    frmEnvBusRoute.Init m_oREBus
'    frmEnvBusRoute.Show vbModal
End Sub

Private Sub imbTkPrice_Click()
On Error GoTo ErrHandle
    frmEnvBusPrice.m_szBusID = m_szBusID
    frmEnvBusPrice.m_dtEnvDate = m_dtBusDate
    frmEnvBusPrice.Show vbModal
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub



Private Sub txtBusId_Change()
    AddEnabled
End Sub

Private Sub txtCheckGate_Change()
    AddEnabled
End Sub

Private Sub txtCheckGate_ButtonClick()
    Dim oShell As New CommDialog
    Dim aszTmp() As String
    oShell.Init g_oActiveUser
    aszTmp = oShell.SelectCheckGate
    Set oShell = Nothing
    If ArrayLength(aszTmp) = 0 Then Exit Sub
    txtCheckGate.Text = MakeDisplayString(aszTmp(1, 1), aszTmp(1, 2))
    AddEnabled
  End Sub


Private Sub txtOwner_ButtonClick()
    Dim oShell As New CommDialog
    Dim aszTmp() As String
    oShell.Init g_oActiveUser
    aszTmp = oShell.SelectOwner
    Set oShell = Nothing
    If ArrayLength(aszTmp) = 0 Then Exit Sub
    txtOwner.Text = MakeDisplayString(aszTmp(1, 1), aszTmp(1, 2))
End Sub


Private Sub txtRouteID_ButtonClick()
On Error GoTo ErrHandle
    Dim oShell As New CommDialog
    Dim aszTmp() As String
    oShell.Init g_oActiveUser
    aszTmp = oShell.SelectRoute
    Set oShell = Nothing
    If ArrayLength(aszTmp) = 0 Then Exit Sub
    If ResolveDisplay(txtRouteID.Text) = aszTmp(1, 1) Then
        txtRouteID.Text = MakeDisplayString(aszTmp(1, 1), aszTmp(1, 2))
        Exit Sub
    End If
    txtRouteID.Text = MakeDisplayString(aszTmp(1, 1), aszTmp(1, 2))
    AddEnabled
    
    '填充站点
    AddStationItems
    
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub txtRouteID_Change()
    AddEnabled
End Sub

Private Sub txtRouteID_KeyPress(KeyAscii As Integer)
On Error GoTo ErrHandle
    If KeyAscii = vbKeyReturn Then
        If Trim(txtRouteID.Text) = "" Then Exit Sub
        AddStationItems
    End If
    
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub txtSplitCompanyID_ButtonClick()
    Dim oShell As New CommDialog
    Dim aszTmp() As String
    oShell.Init g_oActiveUser
    aszTmp = oShell.SelectCompany
    Set oShell = Nothing
    If ArrayLength(aszTmp) = 0 Then Exit Sub
    txtSplitCompanyID.Text = MakeDisplayString(aszTmp(1, 1), aszTmp(1, 2))
End Sub

Private Sub txtVehicleID_Change()
    AddEnabled
End Sub

Private Sub txtVehicleId_ButtonClick()
    Dim oShell As New CommDialog
    oShell.Init g_oActiveUser
    Dim aszTmp() As String
    aszTmp = oShell.SelectVehicleEX(False)
    If ArrayLength(aszTmp) = 0 Then Exit Sub
    txtVehicleID.Text = MakeDisplayString(aszTmp(1, 1), aszTmp(1, 2))
    
    AddEnabled
    RefreshVehicle
End Sub
Private Sub txtVehicleId_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        RefreshVehicle
    End If
End Sub

Private Sub RefreshVehicle()
   Dim oVehicle As New Vehicle
    Dim szVehicleID As String
    szVehicleID = ResolveDisplay(txtVehicleID.Text)
    If szVehicleID <> "" Then
        oVehicle.Init g_oActiveUser
        oVehicle.Identify szVehicleID
        lblBusTypeA.Caption = "车型:" & oVehicle.VehicleModelName
        lblBusOwnerA.Caption = "车主:" & oVehicle.OwnerName
        lblBusSeatA.Caption = "座位数:" & oVehicle.SeatCount
        lblCorpA.Caption = "参运公司:" & oVehicle.CompanyName
        txtVehicleID.Text = MakeDisplayString(oVehicle.VehicleId, oVehicle.LicenseTag)
        If oVehicle.Status = ST_VehicleStop Then
            imgStop.Visible = True
            imgRun.Visible = False
        Else
            imgStop.Visible = False
            imgRun.Visible = True
        End If
    Else
        lblBusTypeA.Caption = "车型:"
        lblBusOwnerA.Caption = "车主:"
        lblBusSeatA.Caption = "座位数:"
        lblCorpA.Caption = "参运公司:"
        imgRun.Visible = True
        imgStop.Visible = False
    End If
End Sub

Public Sub AddEnabled()
    If txtBusID.Text = "" Or Len(txtVehicleID.Text) <= 0 _
       Or Len(txtCheckGate.Text) <= 0 Or Len(txtRouteID.Text) <= 0 Or cboBusType.Text = "" Then
        cmdOk.Enabled = False
    Else
        cmdOk.Enabled = True
    End If
End Sub
Private Sub AddStationItems()
    Dim oRoute As New Route
    Dim szStationInfo() As String
    oRoute.Init g_oActiveUser
    oRoute.Identify ResolveDisplay(txtRouteID.Text)
    szStationInfo = oRoute.RouteStationEx
    
    Dim nCount As Integer
    Dim i As Integer
    Dim oListItem As ListItem
    Dim szTemp As String
    nCount = ArrayLength(szStationInfo)
    lvStationInfo.ListItems.Clear
    lvStationInfo.Sorted = False
    For i = 1 To nCount
      szTemp = Trim(szStationInfo(i, 1)) & "[" & Trim(szStationInfo(i, 3)) & "]"
      Set oListItem = lvStationInfo.ListItems.Add(, , szTemp)
    
      If i = nCount Then oListItem.Checked = True
    Next
'lvStationInfo
'lvStationInfo.View = lvwIcon

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
'新增环境车次,返回错误信息
Private Function AddNewEnvBus() As String
    Dim oREScheme As New REScheme
    Dim oRoutePrice As New STPrice.RoutePriceTable
    Dim aszStationInfo() As String
    Dim szPriceTable As String
    Dim tBusPrice() As TBusPriceDetailInfo
    
    oRoutePrice.Init g_oActiveUser
    aszStationInfo = GetStationItem
    oREScheme.Init g_oActiveUser

    m_szBusID = Trim(txtBusID.Text)
    m_dtBusDate = dtpOffDate.Value
    If oREScheme.IdentifyNotBusId(m_szBusID, m_dtBusDate, g_szExePriceTable) = False Then
        AddNewEnvBus = Format(m_dtBusDate, "YYYY年MM月DD日") & "的车次" & EncodeString(m_szBusID) & "在环境或计划中已存在，不能新增!"
        Exit Function
    End If

    oRoutePrice.Identify g_szExePriceTable
    tBusPrice = oRoutePrice.MakeEnvBusPrice(ResolveDisplay(txtRouteID.Text), txtBusID.Text, ResolveDisplay(txtVehicleID.Text), dtpStartupTime.Value, ResolveDisplay(txtCheckGate.Text), aszStationInfo)
    If ArrayLength(tBusPrice) = 0 Then
        AddNewEnvBus = Format(m_dtBusDate, "YYYY年MM月DD日") & "的车次" & EncodeString(txtBusID.Text) & "没有生成票价!"
        Exit Function
    End If
    '支持始发班和末班车时，应更改
    oREScheme.AddEnviromentBus m_dtBusDate, dtpStartupTime.Value, Trim(txtBusID.Text), ResolveDisplay(txtVehicleID.Text), ResolveDisplay(txtCheckGate.Text), ResolveDisplay(cboBusType.Text), IIf(OptInterNet1.Value, 1, 0), optDelay.Value, Val(txtCheckTime.Text), tBusPrice
    Status = EFS_Modify
    
End Function
'更改环境车次
Private Function ModifyBus() As String
    m_oREBus.Identify m_szBusID, m_dtBusDate
    If optNoR.Value Then
        m_oREBus.AllRefundment = False
    Else
        m_oREBus.AllRefundment = True
    End If
    '是否晚点
    If optDelay.Value Then
        m_oREBus.DelayStatus = True
    Else
        m_oREBus.DelayStatus = False
    End If
    m_oREBus.CheckGate = ResolveDisplay(txtCheckGate.Text)
'    If m_oReBus.BusType = TP_ScrollBus Then
        m_oREBus.ScrollBusCheckTime = Val(txtCheckTime.Text)
'    Else
        m_oREBus.StartUpTime = dtpStartupTime.Value
'    End If
    m_oREBus.BusType = ResolveDisplay(cboBusType.Text)
    m_oREBus.Owner = ResolveDisplay(txtOwner.Text)
    m_oREBus.SplitCompanyID = ResolveDisplay(txtSplitCompanyID.Text)
    m_oREBus.Vehicle = ResolveDisplay(txtVehicleID.Text)
    If OptInterNet1.Value = True Then
        m_oREBus.InternetStatus = CnInternetCanSell
    Else
        m_oREBus.InternetStatus = CnInternetNotCanSell
    End If
    m_oREBus.Update
End Function
