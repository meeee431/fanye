VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Begin VB.Form frmSalelAndSlitpInfo 
   Caption         =   "售票信息及拆分"
   ClientHeight    =   7395
   ClientLeft      =   2565
   ClientTop       =   1980
   ClientWidth     =   9900
   Icon            =   "frmSalelAndSlitpInfo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   9900
   Begin MSComctlLib.ListView lvBusSale 
      Height          =   5565
      Left            =   150
      TabIndex        =   0
      Top             =   1410
      Width           =   6465
      _ExtentX        =   11404
      _ExtentY        =   9816
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "序号"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "车票状态"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "到站"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "座型"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "座号"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "票号"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "票种"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "票价"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "售票时间"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "售票人员"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "目标车次信息"
      Height          =   5595
      Left            =   6600
      TabIndex        =   5
      Top             =   1380
      Width           =   3195
      Begin RTComctl3.CoolButton cmdSlitpInfo 
         Height          =   345
         Left            =   330
         TabIndex        =   47
         Top             =   5040
         Width           =   1245
         _ExtentX        =   0
         _ExtentY        =   0
         BTYPE           =   2
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
         MICON           =   "frmSalelAndSlitpInfo.frx":000C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   4545
         Left            =   180
         TabIndex        =   14
         Top             =   360
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   8017
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         Tab             =   1
         TabHeight       =   520
         TabCaption(0)   =   "车次代码"
         TabPicture(0)   =   "frmSalelAndSlitpInfo.frx":0028
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "lblSaleSeat"
         Tab(0).Control(1)=   "lblBusToalSeat"
         Tab(0).Control(2)=   "lblBusVehicle"
         Tab(0).Control(3)=   "Label17"
         Tab(0).Control(4)=   "Label18"
         Tab(0).Control(5)=   "lblBusStatus"
         Tab(0).Control(6)=   "txtBusID"
         Tab(0).Control(7)=   "lvBus"
         Tab(0).Control(8)=   "Frame3(0)"
         Tab(0).ControlCount=   9
         TabCaption(1)   =   "站点代码"
         TabPicture(1)   =   "frmSalelAndSlitpInfo.frx":0044
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Label1"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Label5"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Label8"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "Label15"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "Label16"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "txtStationID"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "lvStation"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "faCanSale"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).ControlCount=   8
         Begin VB.Frame faCanSale 
            Caption         =   "可售座位"
            Height          =   705
            Left            =   180
            TabIndex        =   25
            Top             =   3300
            Width           =   2505
            Begin VB.Label Label14 
               Caption         =   "加座"
               Height          =   195
               Left            =   1680
               TabIndex        =   38
               Top             =   330
               Width           =   375
            End
            Begin VB.Label Label13 
               Caption         =   "卧铺"
               Height          =   195
               Left            =   900
               TabIndex        =   37
               Top             =   330
               Width           =   375
            End
            Begin VB.Label Label12 
               Caption         =   "普通"
               Height          =   195
               Left            =   150
               TabIndex        =   36
               Top             =   330
               Width           =   375
            End
            Begin VB.Label lblStationBusAdd 
               AutoSize        =   -1  'True
               Caption         =   "0"
               Height          =   180
               Left            =   2220
               TabIndex        =   28
               Top             =   330
               Width           =   90
            End
            Begin VB.Label lblStationBusBedseat 
               AutoSize        =   -1  'True
               Caption         =   "0"
               Height          =   180
               Left            =   1410
               TabIndex        =   27
               Top             =   330
               Width           =   90
            End
            Begin VB.Label lblStatioBusSeat 
               AutoSize        =   -1  'True
               Caption         =   "0"
               Height          =   180
               Left            =   630
               TabIndex        =   26
               Top             =   330
               Width           =   90
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "可售座位"
            Height          =   735
            Index           =   0
            Left            =   -74790
            TabIndex        =   21
            Top             =   3300
            Width           =   2415
            Begin VB.Label Label21 
               Caption         =   "加座"
               Height          =   195
               Left            =   1650
               TabIndex        =   45
               Top             =   360
               Width           =   405
            End
            Begin VB.Label Label20 
               Caption         =   "卧铺"
               Height          =   195
               Left            =   870
               TabIndex        =   44
               Top             =   360
               Width           =   465
            End
            Begin VB.Label Label19 
               Caption         =   "普通"
               Height          =   285
               Left            =   120
               TabIndex        =   43
               Top             =   360
               Width           =   465
            End
            Begin VB.Label lblAddSeat 
               AutoSize        =   -1  'True
               Caption         =   "0"
               Height          =   180
               Left            =   2100
               TabIndex        =   35
               Top             =   360
               Width           =   90
            End
            Begin VB.Label lblBedSeat 
               AutoSize        =   -1  'True
               Caption         =   "0"
               Height          =   180
               Left            =   1350
               TabIndex        =   34
               Top             =   360
               Width           =   90
            End
            Begin VB.Label lblSeat 
               AutoSize        =   -1  'True
               Caption         =   "0"
               Height          =   180
               Left            =   600
               TabIndex        =   33
               Top             =   360
               Width           =   90
            End
            Begin VB.Label lblBusAddSeat 
               Height          =   225
               Left            =   2010
               TabIndex        =   24
               Top             =   330
               Width           =   315
            End
            Begin VB.Label lblBusBed 
               Height          =   225
               Left            =   1350
               TabIndex        =   23
               Top             =   330
               Width           =   435
            End
            Begin VB.Label lblBusSeat 
               Height          =   225
               Left            =   480
               TabIndex        =   22
               Top             =   360
               Width           =   375
            End
         End
         Begin MSComctlLib.ListView lvStation 
            Height          =   2025
            Left            =   180
            TabIndex        =   18
            Top             =   1260
            Width           =   2505
            _ExtentX        =   4419
            _ExtentY        =   3572
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   8
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "车次"
               Object.Width           =   1235
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "车型"
               Object.Width           =   1235
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "普通"
               Object.Width           =   1058
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "卧铺"
               Object.Width           =   1058
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "加座"
               Object.Width           =   1058
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "线路"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Object.Width           =   1235
            EndProperty
         End
         Begin RTComctl3.TextButtonBox txtStationID 
            Height          =   345
            Left            =   180
            TabIndex        =   17
            Top             =   570
            Width           =   2475
            _ExtentX        =   4366
            _ExtentY        =   609
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
         Begin MSComctlLib.ListView lvBus 
            Height          =   1995
            Left            =   -74790
            TabIndex        =   16
            Top             =   1290
            Width           =   2475
            _ExtentX        =   4366
            _ExtentY        =   3519
            SortKey         =   1
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "简码"
               Object.Width           =   1059
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "序号"
               Object.Width           =   1059
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "站名"
               Object.Width           =   1059
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "状态"
               Object.Width           =   1941
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Object.Width           =   2540
            EndProperty
         End
         Begin RTComctl3.TextButtonBox txtBusID 
            Height          =   345
            Left            =   -74790
            TabIndex        =   15
            Top             =   540
            Width           =   2475
            _ExtentX        =   4366
            _ExtentY        =   609
            ForeColor       =   -2147483642
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
         Begin VB.Label lblBusStatus 
            AutoSize        =   -1  'True
            Caption         =   "车次状态:"
            ForeColor       =   &H000000C0&
            Height          =   180
            Left            =   -73680
            TabIndex        =   46
            Top             =   1020
            Visible         =   0   'False
            Width           =   810
         End
         Begin VB.Label Label18 
            Caption         =   "总座位:"
            Height          =   195
            Left            =   -73440
            TabIndex        =   42
            Top             =   4200
            Width           =   660
         End
         Begin VB.Label Label17 
            Caption         =   "已售座位:"
            Height          =   255
            Left            =   -74730
            TabIndex        =   41
            Top             =   4170
            Width           =   915
         End
         Begin VB.Label Label16 
            Caption         =   "总座位:"
            Height          =   255
            Left            =   1680
            TabIndex        =   40
            Top             =   4140
            Width           =   645
         End
         Begin VB.Label Label15 
            Caption         =   "已售座位："
            Height          =   255
            Left            =   270
            TabIndex        =   39
            Top             =   4140
            Width           =   915
         End
         Begin VB.Label Label8 
            Caption         =   "经过改站车次信息:"
            Height          =   225
            Left            =   210
            TabIndex        =   32
            Top             =   1020
            Width           =   1575
         End
         Begin VB.Label lblBusVehicle 
            AutoSize        =   -1  'True
            Caption         =   "车型:"
            Height          =   180
            Left            =   -74850
            TabIndex        =   31
            Top             =   1020
            Width           =   450
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   180
            Left            =   2460
            TabIndex        =   30
            Top             =   4140
            Width           =   90
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   180
            Left            =   1290
            TabIndex        =   29
            Top             =   4140
            Width           =   90
         End
         Begin VB.Label lblBusToalSeat 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   195
            Left            =   -72720
            TabIndex        =   20
            Top             =   4200
            Width           =   90
         End
         Begin VB.Label lblSaleSeat 
            Caption         =   "0"
            Height          =   225
            Left            =   -73800
            TabIndex        =   19
            Top             =   4170
            Width           =   165
         End
      End
      Begin RTComctl3.CoolButton CmdExe 
         Height          =   345
         Left            =   1740
         TabIndex        =   6
         Top             =   5040
         Width           =   1215
         _ExtentX        =   0
         _ExtentY        =   0
         BTYPE           =   2
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
         MICON           =   "frmSalelAndSlitpInfo.frx":0060
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.Frame framBusInfo 
      Caption         =   "待处理车次信息"
      Height          =   1215
      Left            =   90
      TabIndex        =   1
      Top             =   90
      Width           =   9735
      Begin VB.ComboBox cmbFunction 
         Height          =   300
         ItemData        =   "frmSalelAndSlitpInfo.frx":007C
         Left            =   7020
         List            =   "frmSalelAndSlitpInfo.frx":0086
         TabIndex        =   52
         Top             =   300
         Width           =   1035
      End
      Begin VB.TextBox txtOldBusID 
         Height          =   345
         Left            =   4140
         TabIndex        =   9
         Top             =   270
         Width           =   1665
      End
      Begin RTComctl3.CoolButton cmdRefish 
         Height          =   345
         Left            =   8340
         TabIndex        =   4
         Top             =   750
         Width           =   1125
         _ExtentX        =   0
         _ExtentY        =   0
         BTYPE           =   2
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
         MICON           =   "frmSalelAndSlitpInfo.frx":0096
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin RTComctl3.CoolButton cmdQuit 
         Height          =   345
         Left            =   8340
         TabIndex        =   3
         Top             =   300
         Width           =   1125
         _ExtentX        =   0
         _ExtentY        =   0
         BTYPE           =   2
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
         MICON           =   "frmSalelAndSlitpInfo.frx":00B2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComCtl2.DTPicker dtRunDate 
         Height          =   345
         Left            =   1260
         TabIndex        =   2
         Top             =   270
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   609
         _Version        =   393216
         Format          =   19726337
         CurrentDate     =   37078
      End
      Begin VB.Label Label7 
         Caption         =   "功能选择(&F):"
         Height          =   225
         Left            =   5880
         TabIndex        =   51
         Top             =   360
         Width           =   1125
      End
      Begin VB.Label lblOldBusVehicle 
         Caption         =   "车型:"
         Height          =   165
         Left            =   5040
         TabIndex        =   48
         Top             =   840
         Width           =   1425
      End
      Begin VB.Label lblSeatSale 
         Height          =   165
         Left            =   3990
         TabIndex        =   13
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "已售座位"
         Height          =   165
         Left            =   2910
         TabIndex        =   12
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblSeatCount 
         Height          =   165
         Left            =   2340
         TabIndex        =   11
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "总座位(&S):"
         Height          =   165
         Left            =   1350
         TabIndex        =   10
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "车次代码(&M)"
         Height          =   225
         Left            =   3000
         TabIndex        =   8
         Top             =   360
         Width           =   1005
      End
      Begin VB.Label Label2 
         Caption         =   "日期时间(&T)"
         Height          =   225
         Left            =   180
         TabIndex        =   7
         Top             =   330
         Width           =   1005
      End
   End
   Begin VB.Label lblOldBus 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   360
      TabIndex        =   50
      Top             =   7080
      Width           =   90
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   5280
      TabIndex        =   49
      Top             =   7110
      Width           =   90
   End
End
Attribute VB_Name = "frmSalelAndSlitpInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
'Public m_bIsShow As Boolean
'
'Public m_szBusID As String  '有外窗体传进的BusID
'Public m_dtRunDate As Date
'
'Private m_szOldBusID As String 'txtBusId =--old
'Private m_szOldSlitpBusID As String   'txtbusId --new
'Private m_szStationIDSlitpBusID As String ' lvstation
'
'Private m_szExeBusID As String ' SlitpBusID
'
'Private m_nCount As Integer   '目标车次可售座位数
'
'Private m_flg As Boolean  '座位重复选取
'Private m_bflgStationIsNo As Boolean  '座位重复选取
'
'Private m_naIndexSlitpSeatInfo() As Integer
'Private m_szaBusSaleInfo() As String
'Private m_oBus As New REBus
'Private m_szStationIDbus() As String '按站点得到的车次数组
'
'
'Private Sub cmbFunction_Click()
'  Dim oREScheme As New REScheme
'  If txtOldBusID.Text <> "" Then
'      FillSeatInfo
'  End If
'
'
'
'  lblStatus.Caption = ""
'  lblStatus.Caption = ""
'
'
'  If cmbFunction.ListIndex = 1 Then
'     If lvStation.ListItems.Count <> 0 Then
'        lvStation_ItemClick lvStation.SelectedItem
'     End If
'     FullLvBusInfo
'  Else
'    FullLvBusInfo
'    CmdExe.Enabled = False
'    oREScheme.Init g_oActiveUser
'    oREScheme.EnBusRollbackTrans
'    Set oREScheme = Nothing
'  End If
'End Sub
'
'Private Sub cmdExe_Click()
'On Error GoTo ErrorHandle
'
' If cmbFunction.ListIndex = 1 Then
'   If lvBusSale.ListItems.Count = 0 Then
'     MsgBox "车次无售票", vbInformation, Me.Caption
'     Exit Sub
'   End If
'
'    ExeSlitp
' Else
'    MsgBox "请选取[拆分]功能", vbInformation, Me.Caption
' End If
'
'Exit Sub
'ErrorHandle:
'  ShowErrorMsg
'End Sub
'
'Private Sub cmdQuit_Click()
'  Unload Me
'End Sub
'
'Private Sub cmdRefish_Click()
'  If txtOldBusID.Text <> "" Then
'     FillSeatInfo
'  End If
'End Sub
'
'Private Function FillSeatInfo()
'  Dim i As Integer
' ' Dim szBusSaleInfo() As String
'  Dim nCount As Integer
'  Dim liTemp As ListItem
'  Dim nSeatCount As Integer
'  Dim szMsg As String
'  Dim nStatus As Integer
'  Dim szData() As String
'  Dim nIndex As Integer
'  Dim szBusDateTemp() As String
'  On Error GoTo ErrorHandle
'  lvBusSale.ListItems.Clear
'  lblOldBus.Caption = ""
'  If txtOldBusID.Text = "" Then Exit Function
'  If m_szOldSlitpBusID <> txtOldBusID.Text And m_szOldSlitpBusID <> "" Then
'    m_oBus.Identify txtOldBusID.Text, dtRunDate.Value
'    m_oBus.ReBusSlipLock False
'    szMsg = "车次[" & m_szOldSlitpBusID & "]允许售票"
'    lblOldBus.Caption = szMsg
'  End If
'  If txtOldBusID.Text = "" Then Exit Function
'  m_szOldSlitpBusID = txtOldBusID.Text
'  m_oBus.Identify txtOldBusID.Text, dtRunDate.Value
'  If cmbFunction.ListIndex = 1 Then
'     m_oBus.ReBusSlipLock True
'     szMsg = szMsg & "车次[" & txtOldBusID.Text & "]暂停售票"
'     lblOldBus.Caption = szMsg
'  End If
'  szBusDateTemp = m_oBus.GetBusSaleSeatInfo
'  m_szaBusSaleInfo = szBusDateTemp
'  nSeatCount = m_oBus.TotalSeat
'  nCount = ArrayLength(m_szaBusSaleInfo)
'  lblSeatCount.Caption = nSeatCount
'  lblSeatSale.Caption = nCount
'  lblOldBusVehicle.Caption = "车型:" & m_oBus.VehicleModelName
'  nStatus = m_oBus.busStatus
'  If nStatus >= CnSlitpStatus Then
'    nStatus = nStatus - CnSlitpStatus
'  End If
'  If nStatus = ST_BusSlitpStop Then
'     szData = m_oBus.GetSlitpBusTicketNo
'  End If
'  For i = 1 To nCount
'        Set liTemp = lvBusSale.ListItems.Add(, , i)
'        Select Case CInt(m_szaBusSaleInfo(i, 11))
'        Case 1
'        liTemp.SubItems(1) = "正常售出" 'station_name
'        Case 33, 34
'        liTemp.SubItems(1) = " 已检" 'station_name
'        Case 2
'        liTemp.SubItems(1) = " 车票改签售出 " 'station_name
'        End Select
'        liTemp.SubItems(2) = Trim(m_szaBusSaleInfo(i, 1)) 'station_name
'        liTemp.SubItems(3) = Trim(m_szaBusSaleInfo(i, 4)) 'seat_type_name
'        liTemp.SubItems(4) = Trim(m_szaBusSaleInfo(i, 5)) ' seat_no
'
'        liTemp.SubItems(5) = Trim(m_szaBusSaleInfo(i, 8)) ' ticket_type_name
'        liTemp.SubItems(6) = Trim(m_szaBusSaleInfo(i, 7)) 'ticket_id
'        liTemp.SubItems(7) = Trim(m_szaBusSaleInfo(i, 6)) 'ticket_price
'        liTemp.SubItems(8) = Format(m_szaBusSaleInfo(i, 3), cszDateTimeStr)  'operation_time
'        liTemp.SubItems(9) = Trim(m_szaBusSaleInfo(i, 2)) 'user_name
'        If nStatus = ST_BusSlitpStop Then
'
'        If GetIndex(m_szaBusSaleInfo(i, 5), szData) Then
'        SetCorlorLvBusSale lvBusSale.ListItems.Count, True
'        End If
'
'        End If
'  Next
'
'   'SetCorlorLvBusSale Val(szData(i, 3)), True
'
'ErrorHandle:
'
'End Function
'
'
'
'Private Sub cmdSlitpInfo_Click()
'Dim ofrm As New frmSlitpInfo
'ofrm.m_szBusID = txtOldBusID.Text
'ofrm.m_dtBusDate = dtRunDate.Value
'ofrm.Show
'End Sub
'
'
'
'Private Sub Form_Load()
'
''    m_oBus.Init g_oActiveUser
''    m_bIsShow = True
''    cmbFunction.ListIndex = 0
''
''    If frmEnvBus.bIsShow = True Then
''
''    dtRunDate.Value = frmEnvBus.dtpStartDate.Value
''
''
''    If m_szBusID <> "" Then
''
''    txtOldBusID.Text = m_szBusID
''    dtRunDate.Value = m_dtRunDate
''    m_oBus.Identify m_szBusID, m_dtRunDate
''
''    If m_oBus.BusType = TP_ScrollBus Then
''    cmbFunction.Clear
''    cmbFunction.AddItem "查询"
''    End If
''
''
''    End If
''
''    FillSeatInfo
''
''    Else
''
''    dtRunDate.Value = Now
''
''    End If
'End Sub
'
'Private Function FulllvStationID()
'  Dim i As Integer
'  Dim nCount As Integer
'  Dim liTemp As ListItem
'  On Error GoTo ErrorHandle
'  If txtStationID.Text = "" Then Exit Function
'  m_szStationIDbus = m_oBus.GetFromEnvBusBusID(ResolveDisplay(txtStationID.Text))
'  nCount = ArrayLength(m_szStationIDbus)
'  lvStation.ListItems.Clear
'  For i = 1 To nCount
'        Set liTemp = lvStation.ListItems.Add(, , m_szStationIDbus(i, 1))
'       'liTemp.subitems()= Trim(m_szStationIDbus(i, 1)) 'bus_id
'        liTemp.SubItems(1) = m_szStationIDbus(i, 2) 'vehicle_type_name
'        liTemp.SubItems(2) = m_szStationIDbus(i, 8) 'Seat_count
'        liTemp.SubItems(3) = m_szStationIDbus(i, 9) 'bed_seat_count
'        liTemp.SubItems(4) = m_szStationIDbus(i, 10) 'add_seat_count
'        liTemp.SubItems(5) = m_szStationIDbus(i, 4) ' route_name
'
'  Next
'   Exit Function
'ErrorHandle:
'    ShowErrorMsg
'End Function
'
'
'
'Private Sub Form_Unload(Cancel As Integer)
'   Dim oReshcme As New REScheme
'   m_szOldBusID = ""
'   m_szOldSlitpBusID = ""
'   m_szStationIDSlitpBusID = ""
'   m_szExeBusID = ""
'   oReshcme.Init g_oActiveUser
'   oReshcme.EnBusRollbackTrans
'   Set m_oBus = Nothing
'   Set oReshcme = Nothing
'End Sub
'
'
'
'
'
'
'
'
'
'
'
'Private Sub lvStation_DblClick()
'
'   cmdExe_Click
'End Sub
'
'Private Sub lvStation_ItemClick(ByVal Item As MSComctlLib.ListItem)
'  Dim nIndex As Integer
'  Dim szMsg As String
'  On Error GoTo ErrorHandle
'  nIndex = Item.Index
'  lblStatus.Caption = ""
'
'  If m_szExeBusID <> "" And m_szExeBusID <> Item.Text Then
'     m_oBus.Identify m_szStationIDSlitpBusID, dtRunDate.Value
'     m_oBus.ReBusSlipLock False
'
'     If cmbFunction.ListIndex = 1 Then
'
'        szMsg = "车次[" & m_szStationIDSlitpBusID & "]允许售票"
'        lblStatus.Caption = szMsg
'        CmdExe.Enabled = True
'
'     End If
'  End If
'
'
'  If Item.Text = "" Then Exit Sub
'
'   m_szStationIDSlitpBusID = Item.Text
'   m_szExeBusID = Item.Text
'
'
'  If cmbFunction.ListIndex = 1 Then
'     CmdExe.Enabled = True
'     m_oBus.Identify Item.Text, dtRunDate.Value
'     m_oBus.ReBusSlipLock True
'     szMsg = szMsg & "  车次[" & Item.Text & "]暂停售票"
'     lblStatus.Caption = szMsg
'  End If
'
'  RefreshBusSeatInfo
''   lblStatioBusSeat.Caption = m_szStationIDbus(nIndex, 8)
''   lblStationBusBedseat.Caption = m_szStationIDbus(nIndex, 9)
''   lblStationBusAdd.Caption = m_szStationIDbus(nIndex, 10)
''   Label1.Caption = m_szStationIDbus(nIndex, 6) - m_szStationIDbus(nIndex, 7)
''   Label5.Caption = m_szStationIDbus(nIndex, 6)
''   m_nCount = Val(lblStatioBusSeat.Caption) + Val(lblStationBusBedseat.Caption) + Val(lblStationBusAdd.Caption)
'  Exit Sub
'ErrorHandle:
'  ShowErrorMsg
'End Sub
'
'Private Sub SSTab1_Click(PreviousTab As Integer)
' On Error GoTo ErrorHandle
' Dim szMsg As String
' lblStatus.Caption = ""
' If PreviousTab = 1 And txtBusId.Text <> "" Then
'
'     If m_szStationIDSlitpBusID <> "" And txtBusId.Text <> m_szStationIDSlitpBusID Then
'
'
'           m_oBus.Identify m_szStationIDSlitpBusID, dtRunDate.Value
'
'           m_oBus.ReBusSlipLock False
'
'           szMsg = "车次[" & m_szStationIDSlitpBusID & "]允许售票"
'
'           lblStatus.Caption = szMsg
'
'     End If
'
'
'     If cmbFunction.ListIndex = 1 Then
'        If txtBusId.Text <> "" Then
'
'          m_szExeBusID = txtBusId.Text  'save slitp bus id
'
'          m_oBus.Identify txtBusId.Text, dtRunDate.Value
'          m_oBus.ReBusSlipLock True
'
'          szMsg = szMsg & "  车次[" & txtBusId.Text & "]暂停售票"
'
'          lblStatus.Caption = szMsg
'
'          ' m_szExeBusID= txtBusID.Text
'
'          m_nCount = Val(lblSeat.Caption) + Val(lblBedSeat.Caption) + Val(lblAddSeat.Caption)
'
'        End If
'          CmdExe.Enabled = True
'     Else
'    '      m_szExeBusID = ""
'          CmdExe.Enabled = False
'     End If
'
'   Else
'
'     szMsg = ""
'     If cmbFunction.ListIndex = 1 Then
'
'        If m_szStationIDSlitpBusID <> "" And txtBusId.Text <> m_szStationIDSlitpBusID Then
'
'             m_oBus.Identify m_szStationIDSlitpBusID, dtRunDate.Value
'             m_oBus.ReBusSlipLock True
'
'             szMsg = "车次[" & m_szStationIDSlitpBusID & "]暂停售票."
'             lblStatus.Caption = szMsg
'
'        End If
'
'     End If
'
'
'
'      If cmbFunction.ListIndex = 1 Then
'
'        m_szExeBusID = m_szStationIDSlitpBusID 'save slitp bus id
'
'        If txtBusId.Text <> "" Then
'            m_oBus.Identify txtBusId.Text, dtRunDate.Value
'            m_oBus.ReBusSlipLock False
'            szMsg = szMsg & "车次[" & txtBusId.Text & "]允许售票."
'            lblStatus.Caption = szMsg
'        End If
'
'     End If
'     CmdExe.Enabled = False
' End If
'    Exit Sub
'ErrorHandle:
'   ShowErrorMsg
'End Sub
'
'
'Private Sub txtBusID_Click()
''    Dim szTemp() As String
''    g_dtdateSellQuery = dtRunDate.Value
''    szTemp = selectAllBus(, , True)
''
''    If ArrayLength(szTemp) <> 0 Then
''        txtBusId.Text = szTemp(1, 1)
''        FullLvBusInfo
''    Else
''        txtBusId.Text = ""
''    End If
'End Sub
'Private Sub txtBusID_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then
'    FullLvBusInfo
'End If
'
'End Sub
'
'Private Sub txtBusID_LostFocus()
'
'If txtBusId.Text <> "" Then
'    FullLvBusInfo
'End If
'
'End Sub
'
'Private Sub txtOldBusID_KeyDown(KeyCode As Integer, Shift As Integer)
'  If KeyCode = 13 Then
'    FillSeatInfo
'  End If
'End Sub
'Private Function FullLvBusInfo()
'  Dim i As Integer
'  Dim nCount As Integer
'  Dim liTemp As ListItem
'  Dim szMsg As String
'  Dim nStatus As Integer
'  Dim szBusStation() As String
'  On Error GoTo ErrorHandle
'
'  lblBusStatus.Visible = False
'
'
'  If txtBusId.Text = "" Then Exit Function
'
'  If m_szOldBusID <> txtBusId.Text And m_szOldBusID <> "" Then
'
'    m_oBus.Identify m_szOldBusID, dtRunDate.Value
'    m_oBus.ReBusSlipLock False
'    szMsg = "车次[" & m_szOldBusID & "]允许售票"
'    m_szOldBusID = txtBusId.Text
'  End If
'
'  If m_szStationIDSlitpBusID <> "" And m_szStationIDSlitpBusID <> txtBusId.Text Then
'
'    m_oBus.Identify m_szStationIDSlitpBusID, dtRunDate.Value
'    m_oBus.ReBusSlipLock False
'    szMsg = "车次[" & m_szStationIDSlitpBusID & "]允许售票"
'
'  End If
'
'  m_szExeBusID = txtBusId.Text 'save slitp bus id
'
'  lvBus.ListItems.Clear
'  m_oBus.Identify txtBusId.Text, dtRunDate.Value
'
'  nStatus = m_oBus.busStatus
'  If nStatus >= CnSlitpStatus Then
'    nStatus = nStatus - CnSlitpStatus
'  End If
'  lblStatus.Caption = ""
'  If nStatus <> ST_BusNormal Then
'
'    lblBusStatus.Visible = True
'    lblBusStatus.Caption = "车次状态:停班或.."
'  Else
'    lblBusStatus.Caption = "车次状态: 正常"
'  End If
'
'
'
'
'  If cmbFunction.ListIndex = 1 Then
'     m_oBus.ReBusSlipLock True
'     szMsg = szMsg & "  车次[" & txtBusId.Text & "]暂停售票"
'     lblStatus.Caption = szMsg
'     CmdExe.Enabled = True
'  End If
'
'  szBusStation = m_oBus.GetEnBusStationInfo
'  nCount = ArrayLength(szBusStation)
'
'  If nCount <> 0 And cmbFunction.ListIndex = 1 Then
'     CmdExe.Enabled = True
'  End If
'  For i = 1 To nCount
'
'    Set liTemp = lvBus.ListItems.Add(, , szBusStation(i, 5))
'        liTemp.SubItems(1) = i
'        liTemp.SubItems(2) = szBusStation(i, 3)
'        If szBusStation(i, 4) <> 0 Then
'         If szBusStation(i, 4) <> -1 Then
'            liTemp.SubItems(3) = "限售[" & szBusStation(i, 4) & "]张"
'         Else
'             liTemp.SubItems(3) = "可售(不限)"
'         End If
'       Else
'             liTemp.SubItems(3) = "不可售"
'       End If
'  Next
'
' RefreshBusSeatInfo
'
''  lblBusVehicle.Caption = "车型：" & m_oBus.VehicleModelName
''  lblSeat.Caption = m_oBus.SeatRemainCount
''  lblBedSeat.Caption = m_oBus.BedSeatCount
''  lblAddSeat.Caption = m_oBus.AdditionalSeatCount
''  lblSaleSeat.Caption = m_oBus.TotalSeat - m_oBus.SaleSeat
''  lblBusToalSeat.Caption = m_oBus.TotalSeat
''
''  m_nCount = Val(lblSeat.Caption) + Val(lblBedSeat.Caption) + Val(lblAddSeat.Caption)
' '  End If
'
'
'  Exit Function
'ErrorHandle:
'   CmdExe.Enabled = False
'  ShowErrorMsg
'End Function
'
'Private Sub txtStationID_Click()
'' Dim aszTemp() As String
''    aszTemp = selectStation(, False)
''    If ArrayLength(aszTemp) = 0 Then Exit Sub
''    txtStationID.Text = aszTemp(1, 1) & "[" & aszTemp(1, 2)
''    FulllvStationID
'End Sub
'
'Private Sub txtStationID_KeyDown(KeyCode As Integer, Shift As Integer)
' If KeyCode = 13 Then
'     FulllvStationID
'  End If
'End Sub
''取的拆分数据
'Private Function GetBusSaleFromLvBusSale() As String()
'Dim i As Integer
'Dim nCount As Integer
'Dim nCountTemp As Integer
'Dim szBusInfo() As String
''Dim szBusStaionID() As String
'''''
'Dim bflg As Boolean
'Dim OSysPra As New SystemParam
'Dim szaStaionID() As String
'Dim j As Integer
'Dim nCountBusStation As Integer
'Dim szMsg As String
'
' OSysPra.Init g_oActiveUser
' bflg = OSysPra.AllowSlitp
' ReDim m_naIndexSlitpSeatInfo(1 To 1)
' ReDim szBusInfo(1 To 1)
'
'
' If m_oBus.BusID <> m_szExeBusID And m_szExeBusID <> "" Then
'
'    m_oBus.Identify m_szExeBusID, dtRunDate.Value
'
' End If
'
' szaStaionID = m_oBus.GetEnBusStationInfo
'
' nCountBusStation = ArrayLength(szaStaionID)
'
' With lvBusSale
'
'      nCount = .ListItems.Count
'
'      For i = 1 To nCount
'
'          If .ListItems(i).Selected Then
'
'                        If Trim(.ListItems(i).ListSubItems(1)) = "已检" Then
'                             szMsg = szMsg & "第" & i & "行选取有误(座号为" & m_szaBusSaleInfo(i, 5) & ":)。原因:" & Chr(10)
'                             szMsg = szMsg & "目标车次[" & m_szExeBusID & "]该座号[已检]"
'                             szMsg = szMsg & Chr(10)
'                             MsgBox szMsg & "拆分失败", vbInformation, Me.Caption
'                             m_bflgStationIsNo = True
'                             Set OSysPra = Nothing
'                             Exit Function
'                        End If
'
'                       '该车是否经过该站点
'                        j = 1
'                        Do While Trim(m_szaBusSaleInfo(i, 10)) <> Trim(szaStaionID(j, 2))
'                           j = j + 1
'                           If j > nCountBusStation Then m_bflgStationIsNo = True: Exit Do
'                        Loop
'
'                        If m_bflgStationIsNo = True Then
'                             szMsg = szMsg & "第" & i & "行选取有误(座号为" & m_szaBusSaleInfo(i, 5) & ":).原因:" & Chr(10)
'                             szMsg = szMsg & "目标车次" & m_szExeBusID & "不经过站点[" & m_szaBusSaleInfo(i, 1) & "]"
'                             szMsg = szMsg & Chr(10)
'
'                             If bflg = False Then '允许拆分不经过站点
'
'                                 MsgBox szMsg & "拆分失败", vbInformation, Me.Caption
'                                 szMsg = ""
'                                 Set OSysPra = Nothing
'                                 Exit Function
'
'                             End If
'                        End If
'
'                  '取票号，保存索引---i
'
'                  nCountTemp = ArrayLength(szBusInfo)
'                  If nCountTemp = 1 And szBusInfo(1) = "" Then
'                          If .ListItems(i).ListSubItems(2).ForeColor = &HFF0000 Then
'                               MsgBox .ListItems(i).ListSubItems(4) & "座位已拆分不能重复拆分"
'                               m_flg = True
'                               Exit Function
'                          End If
'
'                          szBusInfo(1) = .ListItems(i).ListSubItems(5)
'                          m_naIndexSlitpSeatInfo(1) = i
'                  Else
'                        nCountTemp = nCountTemp + 1
'
'                         If .ListItems(i).ListSubItems(2).ForeColor = &HFF0000 Then
'                              MsgBox .ListItems(i).ListSubItems(4) & "座位已拆分不能重复拆分"
'                              m_flg = True
'                              Exit Function
'                         End If
'
'                         ReDim Preserve szBusInfo(1 To nCountTemp)
'                         ReDim Preserve m_naIndexSlitpSeatInfo(1 To nCountTemp)
'                         szBusInfo(nCountTemp) = .ListItems(i).ListSubItems(5)
'                         'szBusInfo(nCountTemp, 2) = .ListItems(i).ListSubItems(8)
'                         m_naIndexSlitpSeatInfo(nCountTemp) = i
'
'
'                End If
'
'            ' Private m_naIndexSlitpSeatInfo() As Integer
'             'Public m_bIsShow As Boolean
'             'Private m_szaBusSaleInfo() As String
'
'         End If
'      Next
' End With
' Set OSysPra = Nothing
' If szMsg <> "" Then
'   If MsgBox(szMsg & "是否拆分？", vbYesNo + vbInformation, Me.Caption) = vbNo Then
'      Exit Function
'   Else
'      m_bflgStationIsNo = False
'   End If
' End If
' GetBusSaleFromLvBusSale = szBusInfo
'End Function
'
'Private Sub UpdateLvBussale(Optional szBudinfo As Variant)
'   Dim nCount As Integer
'   Dim i As Integer
'   nCount = ArrayLength(szBudinfo)
'   With lvBusSale
'   For i = 1 To nCount
'       .ListItems(szBudinfo(i, 3)).ListSubItems(13) = "已拆"
'       SetCorlorLvBusSale CInt(szBudinfo(i, 3))
'   Next
'   End With
'End Sub
'Public Sub SetCorlorLvBusSale(nIdex As Integer, Optional bflg As Boolean)
' Dim i As Integer
'
'   With lvBusSale
'       If bflg = True Then
'            For i = 1 To 9
'               .ListItems(nIdex).ListSubItems(i).ForeColor = &HFF0000
'            Next
'       Else
'            If .ListItems(nIdex).ListSubItems(2).ForeColor = &HFF0000 Then
'               For i = 1 To 9
'                  .ListItems(nIdex).ListSubItems(i).ForeColor = vbDefault
'               Next
'            End If
'       End If
'   End With
'
'End Sub
'Private Function ExeSlitp()
''    Dim i As Integer
''    Dim szBusSaleInfo() As String
''    Dim szBusRuturnInfo() As String
''    Dim nCount As Integer
''    Dim nCount1 As Integer
''    Dim szMsg As String
''
''
''    If m_szExeBusID = txtOldBusID.Text Then
''    MsgBox "被拆车次和目标车次相同，不能拆分", vbInformation, Me.Caption
''    Exit Function
''    End If
''
''    '取的拆分数据
''    szBusSaleInfo = GetBusSaleFromLvBusSale
''    If m_bflgStationIsNo = True Then m_bflgStationIsNo = False: Exit Function
''    If m_flg = True Then m_flg = False: Exit Function '座位重复选取
''
''    nCount = ArrayLength(szBusSaleInfo)
''
''    If nCount = 0 Then
''    MsgBox "请选取座位", vbInformation, Me.Caption
''    Exit Function
''    End If
''    szMsg = "是否将车次[" & txtOldBusID.Text & "]的[" & nCount & "]个座位拆分到[" & m_szExeBusID & "]车次" & Chr(10)
''    szMsg = szMsg & "目标车次[" & m_szExeBusID & "]可售座位" & m_nCount
''    If MsgBox(szMsg, vbInformation + vbYesNo, Me.Caption) = vbNo Then Exit Function
''    '拆分
''    m_oBus.Identify txtOldBusID.Text, dtRunDate.Value
''    szBusRuturnInfo = m_oBus.MegerBusAndSlitpBus(szBusSaleInfo, m_szExeBusID)
''    nCount = ArrayLength(szBusRuturnInfo)
''    '
''    If nCount <> 0 Then
''    lblStatus.Caption = "车次[" & m_szExeBusID & "]允许售票"
''    lblOldBus.Caption = "车次[" & txtOldBusID.Text & "]拆分停班"
''    MsgBox "拆分成功", vbInformation, Me.Caption
''    If frmEnvBus.bIsShow = True Then
''    frmEnvBus.UpdateList txtOldBusID.Text, dtRunDate.Value
''    End If
''
''    For i = 1 To nCount
''    SetCorlorLvBusSale m_naIndexSlitpSeatInfo(i), True    '设置色
''    Next
''
''    RefreshBusSeatInfo
''
''    End If
'
'End Function
'Private Function GetIndex(szSeatNo As String, szData() As String) As Boolean
'Dim nCount As Integer
'Dim i As Integer
'On Error GoTo ErrorHandle
'nCount = ArrayLength(szData)
'If nCount = 0 Then Exit Function
'i = 1
'
'Do While Trim(szSeatNo) <> Trim(szData(i, 3))
'   i = i + 1
'   If i > nCount Then GetIndex = False: Exit Function
'Loop
'
'GetIndex = True
'
'ErrorHandle:
'
'End Function
'Private Function RefreshBusSeatInfo()
'  If m_oBus.BusID <> m_szExeBusID Then
'   m_oBus.Identify m_szExeBusID, m_dtRunDate
'  End If
'  If SSTab1.Tab = 1 Then
'
'      lblStatioBusSeat.Caption = m_oBus.SeatRemainCount
'      lblStationBusBedseat.Caption = m_oBus.BedSeatCount
'      lblStationBusAdd.Caption = m_oBus.AdditionalSeatCount
'      Label1.Caption = m_oBus.SaledSeatCount
'      Label5.Caption = m_oBus.TotalSeat
'      m_nCount = Val(lblStatioBusSeat.Caption) + Val(lblStationBusBedseat.Caption) + Val(lblStationBusAdd.Caption)
'
'  Else
'
'     lblBusVehicle.Caption = "车型：" & m_oBus.VehicleModelName
'     lblSeat.Caption = m_oBus.SeatRemainCount
'     lblBedSeat.Caption = m_oBus.BedSeatCount
'     lblAddSeat.Caption = m_oBus.AdditionalSeatCount
'     lblSaleSeat.Caption = m_oBus.TotalSeat - m_oBus.SaleSeat
'     lblBusToalSeat.Caption = m_oBus.TotalSeat
'     m_nCount = Val(lblSeat.Caption) + Val(lblBedSeat.Caption) + Val(lblAddSeat.Caption)
'
'  End If
'
'End Function
'
