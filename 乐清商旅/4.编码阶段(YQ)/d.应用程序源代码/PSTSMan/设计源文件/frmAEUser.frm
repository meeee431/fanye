VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmAEUser 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "UserProperty"
   ClientHeight    =   4260
   ClientLeft      =   3000
   ClientTop       =   3870
   ClientWidth     =   7800
   HelpContextID   =   5000090
   Icon            =   "frmAEUser.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   7800
   ShowInTaskbar   =   0   'False
   Begin RTComctl3.CoolButton cmdAEUser 
      Height          =   315
      Left            =   3675
      TabIndex        =   20
      Top             =   3855
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   556
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
      MICON           =   "frmAEUser.frx":038A
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
      Height          =   3630
      Left            =   120
      TabIndex        =   37
      Top             =   150
      Width           =   7500
      _ExtentX        =   13229
      _ExtentY        =   6403
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   14737632
      TabCaption(0)   =   "用户基本属性"
      TabPicture(0)   =   "frmAEUser.frx":03A6
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "  登录限制  "
      TabPicture(1)   =   "frmAEUser.frx":03C2
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "chkLoginCount"
      Tab(1).Control(1)=   "fraWorkStation"
      Tab(1).Control(2)=   "fraWeek"
      Tab(1).Control(3)=   "fraLoginTime"
      Tab(1).Control(4)=   "Frame2"
      Tab(1).ControlCount=   5
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   3280
         Left            =   15
         TabIndex        =   57
         Top             =   315
         Width           =   7455
         Begin VB.CheckBox chkEnable 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "此用户账号可用(&E)"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   3885
            TabIndex        =   17
            Top             =   2760
            Value           =   1  'Checked
            Width           =   1860
         End
         Begin VB.CheckBox chkInternet 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "此用户为Internet用户(&I)"
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   1140
            TabIndex        =   16
            Top             =   2715
            Width           =   2370
         End
         Begin VB.TextBox txtAnno 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   1395
            TabIndex        =   13
            Top             =   2235
            Width           =   1665
         End
         Begin VB.TextBox txtUserName 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   1395
            TabIndex        =   3
            Top             =   525
            Width           =   4515
         End
         Begin VB.TextBox txtUserID 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   1395
            TabIndex        =   1
            Top             =   120
            Width           =   4515
         End
         Begin VB.TextBox txtPass 
            Appearance      =   0  'Flat
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   1395
            PasswordChar    =   "*"
            TabIndex        =   5
            Text            =   "Text1"
            Top             =   945
            Width           =   1665
         End
         Begin VB.TextBox txtRePass 
            Appearance      =   0  'Flat
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   4200
            PasswordChar    =   "*"
            TabIndex        =   7
            Text            =   "Text2"
            Top             =   945
            Width           =   1695
         End
         Begin VB.TextBox txtCanSellDay 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   4200
            TabIndex        =   15
            Text            =   "3"
            Top             =   2235
            Width           =   1695
         End
         Begin VB.ComboBox cboStation 
            Height          =   300
            ItemData        =   "frmAEUser.frx":03DE
            Left            =   1395
            List            =   "frmAEUser.frx":03E0
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   1770
            Width           =   4515
         End
         Begin RTComctl3.CoolButton cmdGroup 
            Height          =   360
            Left            =   6165
            TabIndex        =   18
            Top             =   120
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   635
            BTYPE           =   3
            TX              =   "组(&G)"
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
            MICON           =   "frmAEUser.frx":03E2
            PICN            =   "frmAEUser.frx":03FE
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin RTComctl3.CoolButton cmdUserRight 
            Height          =   360
            Left            =   6165
            TabIndex        =   32
            Top             =   540
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   635
            BTYPE           =   3
            TX              =   "权限(&R)"
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
            MICON           =   "frmAEUser.frx":0798
            PICN            =   "frmAEUser.frx":07B4
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin RTComctl3.CoolButton cmdClear 
            Height          =   360
            Left            =   6165
            TabIndex        =   19
            Top             =   960
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   635
            BTYPE           =   3
            TX              =   "清除密码(&M)"
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
            MICON           =   "frmAEUser.frx":090E
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.ComboBox cboUnit 
            Height          =   300
            Left            =   1395
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   1365
            Width           =   4515
         End
         Begin VB.Label lblUserID 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   1395
            TabIndex        =   59
            Top             =   195
            Width           =   90
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "用户注释(&Z):"
            Height          =   180
            Left            =   240
            TabIndex        =   12
            Top             =   2295
            Width           =   1080
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "用户单位(&D):"
            Height          =   180
            Left            =   240
            TabIndex        =   8
            Top             =   1410
            Width           =   1080
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "用户名称(&N):"
            Height          =   180
            Left            =   240
            TabIndex        =   2
            Top             =   600
            Width           =   1080
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "用户代码(&U):"
            Height          =   180
            Left            =   240
            TabIndex        =   0
            Top             =   195
            Width           =   1080
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "用户密码(P):"
            Height          =   180
            Left            =   240
            TabIndex        =   4
            Top             =   1005
            Width           =   1080
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "确认密码(&F):"
            Height          =   180
            Left            =   3135
            TabIndex        =   6
            Top             =   1005
            Width           =   1080
         End
         Begin VB.Label Label17 
            BackColor       =   &H00E0E0E0&
            Caption         =   "可售天数(&Y):"
            Height          =   180
            Left            =   3135
            TabIndex        =   14
            Top             =   2295
            Width           =   1080
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "所属车站(&S):"
            Height          =   180
            Left            =   240
            TabIndex        =   10
            Top             =   1815
            Width           =   1080
         End
      End
      Begin VB.CheckBox chkLoginCount 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "允许此用户同时多次登录(&M)"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -74685
         TabIndex        =   45
         Top             =   3255
         Value           =   1  'Checked
         Width           =   3015
      End
      Begin VB.Frame fraWorkStation 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   1290
         Left            =   -74700
         TabIndex        =   49
         Top             =   1860
         Width           =   7050
         Begin VB.CheckBox chkWorkStation 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "指定登录工作站(&G)"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   0
            TabIndex        =   38
            Top             =   0
            Width           =   1890
         End
         Begin VB.TextBox txtHostName 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   300
            Index           =   5
            Left            =   5385
            TabIndex        =   44
            Top             =   930
            Width           =   1590
         End
         Begin VB.TextBox txtHostName 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   300
            Index           =   4
            Left            =   3000
            TabIndex        =   43
            Top             =   930
            Width           =   1590
         End
         Begin VB.TextBox txtHostName 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   300
            Index           =   3
            Left            =   735
            TabIndex        =   42
            Top             =   930
            Width           =   1590
         End
         Begin VB.TextBox txtHostName 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   300
            Index           =   2
            Left            =   5385
            TabIndex        =   41
            Top             =   570
            Width           =   1590
         End
         Begin VB.TextBox txtHostName 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   300
            Index           =   1
            Left            =   3000
            TabIndex        =   40
            Top             =   570
            Width           =   1590
         End
         Begin VB.TextBox txtHostName 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   300
            Index           =   0
            Left            =   735
            TabIndex        =   39
            Top             =   570
            Width           =   1590
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "6(&6):"
            Enabled         =   0   'False
            Height          =   180
            Left            =   4890
            TabIndex        =   56
            Top             =   975
            Width           =   450
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "5(&5):"
            Enabled         =   0   'False
            Height          =   180
            Left            =   2520
            TabIndex        =   55
            Top             =   975
            Width           =   450
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "4(&4):"
            Enabled         =   0   'False
            Height          =   180
            Left            =   270
            TabIndex        =   54
            Top             =   975
            Width           =   450
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "3(&3):"
            Enabled         =   0   'False
            Height          =   180
            Left            =   4890
            TabIndex        =   53
            Top             =   615
            Width           =   450
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "2(&2):"
            Enabled         =   0   'False
            Height          =   180
            Left            =   2520
            TabIndex        =   52
            Top             =   615
            Width           =   450
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "1(&1):"
            Enabled         =   0   'False
            Height          =   180
            Left            =   270
            TabIndex        =   51
            Top             =   615
            Width           =   450
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "允许用户登录的工作站"
            Height          =   180
            Left            =   240
            TabIndex        =   50
            Top             =   330
            Width           =   1800
         End
         Begin VB.Line Line6 
            BorderColor     =   &H00808080&
            X1              =   1260
            X2              =   6870
            Y1              =   60
            Y2              =   60
         End
         Begin VB.Line Line7 
            BorderColor     =   &H00FFFFFF&
            X1              =   1275
            X2              =   6855
            Y1              =   75
            Y2              =   75
         End
      End
      Begin VB.Frame fraWeek 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   570
         Left            =   -74700
         TabIndex        =   48
         Top             =   1215
         Width           =   7005
         Begin VB.CheckBox chkWeek 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "指定登录日(&W)"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   0
            TabIndex        =   28
            Top             =   0
            Width           =   1530
         End
         Begin VB.CheckBox chkWorkDate 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "星期三"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   3
            Left            =   3225
            TabIndex        =   33
            Top             =   300
            Width           =   855
         End
         Begin VB.CheckBox chkWorkDate 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "星期六"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   6
            Left            =   6090
            TabIndex        =   36
            Top             =   300
            Width           =   855
         End
         Begin VB.CheckBox chkWorkDate 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "星期二"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   2220
            TabIndex        =   31
            Top             =   300
            Width           =   855
         End
         Begin VB.CheckBox chkWorkDate 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "星期五"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   5
            Left            =   5175
            TabIndex        =   35
            Top             =   300
            Width           =   855
         End
         Begin VB.CheckBox chkWorkDate 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "星期一"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   1245
            TabIndex        =   30
            Top             =   300
            Width           =   840
         End
         Begin VB.CheckBox chkWorkDate 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "星期四"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   4
            Left            =   4215
            TabIndex        =   34
            Top             =   300
            Width           =   855
         End
         Begin VB.CheckBox chkWorkDate 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "星期日"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   270
            TabIndex        =   29
            Top             =   300
            Width           =   855
         End
         Begin VB.Line Line4 
            BorderColor     =   &H00808080&
            X1              =   1215
            X2              =   6855
            Y1              =   75
            Y2              =   75
         End
         Begin VB.Line Line5 
            BorderColor     =   &H00FFFFFF&
            X1              =   1530
            X2              =   6870
            Y1              =   90
            Y2              =   90
         End
      End
      Begin VB.Frame fraLoginTime 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   645
         Left            =   -74760
         TabIndex        =   46
         Top             =   600
         Width           =   7050
         Begin VB.CheckBox chkLoginTime 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "时间限制(&T)"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   90
            TabIndex        =   23
            Top             =   15
            Width           =   1395
         End
         Begin MSComCtl2.DTPicker dtpEndTime 
            Height          =   315
            Left            =   5850
            TabIndex        =   27
            Top             =   285
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   74973186
            CurrentDate     =   36409.9999884259
         End
         Begin MSComCtl2.DTPicker dtpBeginTime 
            Height          =   315
            Left            =   3420
            TabIndex        =   25
            Top             =   285
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   74973186
            CurrentDate     =   36409.0000115741
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "结束时间(&E):"
            Enabled         =   0   'False
            Height          =   180
            Left            =   4635
            TabIndex        =   26
            Top             =   345
            Width           =   1080
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "开始时间(&S):"
            Enabled         =   0   'False
            Height          =   180
            Left            =   2295
            TabIndex        =   24
            Top             =   345
            Width           =   1080
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "允许用户登录的时间段"
            Height          =   180
            Left            =   330
            TabIndex        =   47
            Top             =   345
            Width           =   1800
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   3255
         Left            =   -74990
         TabIndex        =   58
         Top             =   330
         Width           =   7455
      End
   End
   Begin RTComctl3.CoolButton cmdClose 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   5025
      TabIndex        =   21
      Top             =   3855
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "关闭"
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
      MICON           =   "frmAEUser.frx":092A
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
      Height          =   315
      Left            =   6390
      TabIndex        =   22
      Top             =   3855
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   556
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
      MICON           =   "frmAEUser.frx":0946
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
Attribute VB_Name = "frmAEUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *******************************************************************
' *  Source File Name  : frmAEUser                                  *
' *  Project Name: PSTSMan                                          *
' *  Engineer:                                                      *
' *  Date Generated:   2002/08/19                                   *
' *  Last Revision Date : 2002/08/19                                *
' *  Brief Description   : 添加用户或编辑用户属性                   *
' *******************************************************************


Option Explicit
Option Base 1
Const cGray = &H80000012
Const cGrayForHost = &HE0E0E0

Public bEdit As Boolean
Public bRightRead As Boolean '标注权限的读取,即为最新数据
Public bGroupRead As Boolean '标注组信息的读取,即为最新数据,新增时用


Dim bInnerUser As Boolean
Dim nUnitCount As Integer '所有单位数
Dim nWN(0 To 6) As Integer '星期
Dim bWeek As Boolean 'true=所有星期
Dim szIP As String 'IP地址
Dim dtStartTime As Date
Dim dtEndTime As Date
Dim szUserID As String
Dim szSellStationID As String
Dim szUserName As String
Dim nMultLogin  As Integer
Dim szUnitID As String
Dim nWeek As Integer
Dim nQuery As Integer
Dim szAnno As String
Dim bIsInternetUser As Boolean

'可售天数
Dim m_nCanSellDay As Integer

Private Sub cboUnit_Click()
    Dim oSysMan As New SystemMan
    Dim tTemp() As TDepartmentInfo
    Dim tTemp1() As TDepartmentInfo
    Dim szUnitID As String
    Dim nCount As Integer
    Dim nCount1 As Integer
    Dim i As Integer
    Dim ListIndex As Integer
    oSysMan.Init g_oActUser
    cboStation.Clear
    szUnitID = ResolveDisplay(cboUnit.Text)
    tTemp = oSysMan.GetAllSellStation(szUnitID)
    tTemp1 = oSysMan.GetAllSellStation(szUnitID, frmSMCMain.lvDetail.SelectedItem.Text)
    nCount = ArrayLength(tTemp)
    nCount1 = ArrayLength(tTemp1)
    cboStation.AddItem ""
    For i = 1 To nCount
         cboStation.AddItem MakeDisplayString(tTemp(i).szSellStationID, tTemp(i).szSellStationName)
         If nCount1 = 0 Then
              ListIndex = 0
         Else
           If tTemp(i).szSellStationID = tTemp1(1).szSellStationID Then
              szSellStationID = tTemp(i).szSellStationID
              ListIndex = i
           End If
         End If
    Next i
    
    If nCount > 0 Then cboStation.ListIndex = ListIndex


End Sub

Private Sub chkEnable_Click()
    '如果是对自己不可用则出错
    If chkEnable.Value = Checked Then
        chkLoginCount.Enabled = True
    Else
        If g_oActUser.UserID = szUserID Then
            MsgBox "不能禁用自己!", vbInformation, cszMsg
            chkEnable.Value = Checked
            Exit Sub
        Else
            chkLoginCount.Enabled = False
        End If
    End If
        
End Sub

Private Sub chkLoginTime_Click()
    EnableContainer fraLoginTime, IIf(chkLoginTime.Value = 1, True, False), chkLoginTime
End Sub


Private Sub chkWeek_Click()
    EnableContainer fraWeek, IIf(chkWeek.Value = 1, True, False), chkWeek
End Sub

Private Sub chkWorkStation_Click()
    DisPlayIPControl
    
End Sub

Private Sub cmdAEUser_Click()
    '进行新增或修改
    Dim oUserTemp As New User
    Dim narrLenOld As Integer
    Dim narrLen As Integer
    Dim i As Integer, j As Integer, bShouldDel As Boolean, bShouldAdd As Boolean
    Dim nAddCount As Integer, nDelCount As Integer
    Dim aszDel() As String
    Dim aszAdd() As String
    
    If cboUnit.Text = "" Then
        MsgBox "请选择单位.", vbInformation, cszMsg
        Exit Sub
    End If
    
    
    GetDateFormUI
    
    If dtStartTime > dtEndTime Or dtStartTime = dtEndTime Then
        MsgBox "登录时间限制,开始时间必须小于结束时间.", vbInformation, cszMsg
        Exit Sub
    End If
    
    If bEdit = True Then
        On Error GoTo there '修改
        SetBusy
        oUserTemp.Init g_oActUser
        oUserTemp.Identify szUserID
        oUserTemp.SellStationID = ResolveDisplay(cboStation.Text)
        oUserTemp.FullName = szUserName
        oUserTemp.LoginHost = szIP
        oUserTemp.CanSellDay = m_nCanSellDay
        oUserTemp.MultiLogin = nMultLogin
        oUserTemp.UnitID = szUnitID
        oUserTemp.WorkBeginTime = dtStartTime
        oUserTemp.WorkEndTime = dtEndTime
        oUserTemp.WorkWeekDay = nWeek
        oUserTemp.InternetUser = bIsInternetUser
        oUserTemp.Annotation = szAnno
        
        
        oUserTemp.Update
        
        
        '修改后数据刷新
        j = ArrayLength(g_atUserInfo)
        For i = 1 To j
            If g_atUserInfo(i).UserID = szUserID Then
                g_atUserInfo(i).SellStationID = szSellStationID 'FL ADD
                g_atUserInfo(i).EndTime = dtEndTime
                g_atUserInfo(i).InnerUser = False
                g_atUserInfo(i).LoginIP = szIP
                g_atUserInfo(i).CanSellDay = m_nCanSellDay
                g_atUserInfo(i).MultLogin = nMultLogin
                g_atUserInfo(i).StartTime = dtStartTime
                g_atUserInfo(i).UnitID = szUnitID
                g_atUserInfo(i).UserName = szUserName
                g_atUserInfo(i).IsInternetUser = bIsInternetUser
                g_atUserInfo(i).Week = nWeek
                g_atUserInfo(i).Annotation = szAnno
            End If
        Next i
        frmStoreMenu.DisplayUserInfo (j)
    Else '新增用户
    
    
        '验证密码
        If txtUserID.Text = "" Then
            MsgBox "请输入新用户代码,重试.", vbInformation, cszMsg
            Exit Sub
        End If
        If Trim(txtPass.Text) <> Trim(txtRePass.Text) Then
            MsgBox "两次输入的密码不相同,重试.", vbInformation, cszMsg
            Exit Sub
        End If
        On Error GoTo ErrorHandle
        '新增
        SetBusy
        
        oUserTemp.Init g_oActUser
        oUserTemp.AddNew
        oUserTemp.UserID = szUserID
        oUserTemp.SellStationID = szSellStationID
        oUserTemp.FullName = szUserName
        oUserTemp.LoginHost = szIP
        oUserTemp.CanSellDay = m_nCanSellDay
        oUserTemp.MultiLogin = nMultLogin
        oUserTemp.UnitID = szUnitID
        oUserTemp.WorkBeginTime = dtStartTime
        oUserTemp.WorkEndTime = dtEndTime
        oUserTemp.WorkWeekDay = nWeek
        oUserTemp.InternetUser = bIsInternetUser
        oUserTemp.Annotation = szAnno
        oUserTemp.PassWord = Trim(txtPass.Text)
        oUserTemp.Update
        
        ''加权限
        narrLen = 0
        On Error Resume Next
        
        narrLen = UBound(g_atInBrowse)
        On Error GoTo 0
        
        If (narrLen <> 0) And (g_bBrowseNull = False) Then
            If g_atInBrowse(1).FunID <> "" Then
            For i = 1 To narrLen
            
                On Error GoTo ErrorHandle '新增
                oUserTemp.Identify szUserID
                oUserTemp.AddFunction (g_atInBrowse(i).FunID)
                
                
            Next i
            End If
        End If
        
        ''加用户组
        narrLen = 0
        
        On Error Resume Next
        narrLen = UBound(g_atBelongGroup)
        On Error GoTo 0
        
        
        If narrLen <> 0 Then
            If g_atBelongGroup(1).GroupID <> "" Then
            Dim oGroup As New UserGroup
            oGroup.Init g_oActUser
            For i = 1 To narrLen
                On Error GoTo ErrorHandle
                oGroup.Identify g_atBelongGroup(i).GroupID
                oGroup.AddUser szUserID
                
            Next i
            End If
        End If
        '新增后数据的刷新
        i = ArrayLength(g_atUserInfo) + 1
        If i > 1 Then
            ReDim Preserve g_atUserInfo(1 To i)
        Else
            ReDim g_atUserInfo(1 To i)
        End If
        g_atUserInfo(i).EndTime = dtEndTime
        g_atUserInfo(i).InnerUser = False
        g_atUserInfo(i).LoginIP = szIP
        g_atUserInfo(i).CanSellDay = m_nCanSellDay
        g_atUserInfo(i).MultLogin = nMultLogin
        g_atUserInfo(i).StartTime = dtStartTime
        g_atUserInfo(i).UnitID = szUnitID
        g_atUserInfo(i).UserID = szUserID
        g_atUserInfo(i).UserName = szUserName
        g_atUserInfo(i).Week = nWeek
        g_atUserInfo(i).IsInternetUser = bIsInternetUser
        frmStoreMenu.DisplayUserInfo (i)
    
    
    End If
   
    '设定选择光条
    For i = 1 To frmSMCMain.lvDetail.ListItems.Count
        If frmSMCMain.lvDetail.ListItems.Item(i).Key = "A" & oUserTemp.UserID Then
            frmSMCMain.lvDetail.ListItems.Item(i).Selected = True
    '        j = i
        Else
            frmSMCMain.lvDetail.ListItems.Item(i).Selected = False
        End If
    Next i
    
ResetDate:
    SetNormal
    Unload Me
    '*****清空下列值
    ReDim g_atBelongGroup(1)
    ReDim g_atBelongGroupOld(1)
    g_bLeftNull = False
    g_bRightNull = False
    ReDim g_atUnBelongGroup(1)
    ReDim g_atAllGroup(1)
    ''''''
    ReDim g_atAuthored(1)
    ReDim g_atAddBrowse(1)
    ReDim g_atInBrowse(1)
    ReDim g_aszFunOld(1)
    g_bBrowseNull = False
    bRightRead = False
    
    Exit Sub
there:
        ShowErrorMsg
        GoTo ResetDate
    
    Exit Sub
ErrorHandle:
    ShowErrorMsg
    GoTo ResetDate
End Sub

Private Sub cmdClear_Click()
    '清空口令
    Dim nTemp As Integer
    Dim oUserTemp As New User
    
    On Error GoTo ErrorHandle
    oUserTemp.Init g_oActUser
    oUserTemp.Identify lblUserID
    
    
    nTemp = MsgBox("确认清除此用户密码?", vbYesNo + vbQuestion, cszMsg)
    If nTemp = vbYes Then
        oUserTemp.BlankPassword
        MsgBox "用户的密码已清除,此用户下次登录的密码为空.", vbInformation, cszMsg
        cmdClear.Enabled = False
    End If
Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub cmdClose_Click()
    Unload Me
    
    
    ReDim g_atBelongGroup(1)
    ReDim g_atBelongGroupOld(1)
    ReDim g_atUnBelongGroup(1)
    ReDim g_atAllGroup(1)
    ReDim g_atAuthored(1)
    ReDim g_atAddBrowse(1)
    ReDim g_atInBrowse(1)
    ReDim g_aszFunOld(1)
    g_bBrowseNull = False
    g_bLeftNull = False
    g_bRightNull = False

End Sub

Private Sub cmdGroup_Click()
    If bEdit = False Then
        If txtUserID.Text = "" Then
            MsgBox "请输入用户代码,重试.", vbInformation, cszMsg
        Else
            frmUserBeGroup.Show vbModal, Me
        End If
    Else
        frmUserBeGroup.Show vbModal, Me
    End If
End Sub

Private Sub cmdHelp_Click()
    DisplayHelp Me
End Sub

Private Sub cmdUserRight_Click()
    If bEdit = False Then
        If txtUserID.Text = "" Then
            MsgBox "请输入用户代码,重试.", vbInformation, cszMsg
        Else
            frmUser_GroupRight.m_bUser = True
            frmUser_GroupRight.Show vbModal, Me
        End If
    Else
            frmUser_GroupRight.m_bUser = True
            frmUser_GroupRight.Show vbModal, Me
    End If
End Sub

Private Sub cmdUserRight_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub
'===================================================
'Modify Date：2002-11-22
'Author:陆勇庆
'Reamrk:功能说明
'===================================================
Private Sub Form_Load()
   
    Dim i As Integer
    
'    bClearPassWord = False
'    bRightChange = False
    bRightRead = False
    bGroupRead = False
'    bGroupChange = False
    g_bRightNull = False
    g_bLeftNull = False
    g_bBrowseNull = True
    bInnerUser = False

    
    Me.Top = (Screen.Height - Me.ScaleHeight) / 2
    Me.Left = (Screen.Width - Me.ScaleWidth) / 2

    ClearTextBox Me
    cboUnit.Clear
    
    
    nUnitCount = 0
    
    nUnitCount = ArrayLength(g_atAllUnit)
    If nUnitCount <> 0 Then
        For i = 1 To nUnitCount
            
            cboUnit.AddItem g_atAllUnit(i).szUnitID & "[" & g_atAllUnit(i).szUnitFullName & "]"
            If g_atAllUnit(i).szUnitID = g_oSysParam.UnitID Then
                cboUnit.ListIndex = i - 1
            End If
        Next
    Else
        ''''''
    End If
   If nUnitCount > 0 Then cboUnit.ListIndex = 0
  
    If bEdit Then
        Me.Caption = "修改用户属性"
        txtUserID.Visible = False
        lblUserID.Visible = True
        
        cmdAEUser.Caption = "确定(&O)"
        txtPass.Enabled = False
        txtRePass.Enabled = False
        txtRePass.BackColor = cGreyColor
        txtPass.BackColor = cGreyColor
        frmAEUser.HelpContextID = 50000100
        LoadData
    Else
        Me.Caption = "新增用户"
        lblUserID.Visible = False
        txtUserID.Visible = True
        cmdClear.Enabled = False
        frmAEUser.HelpContextID = 5000090
    End If
    
    If bInnerUser = True Then
        Dim oTemp As Object
        For Each oTemp In Me.Controls
            If TypeName(oTemp) <> "Line" Then
                oTemp.Enabled = False
            End If
        Next
        cmdClear.Enabled = True
        cmdClose.Enabled = True
        cmdHelp.Enabled = True
    End If
End Sub


Private Sub txtAnno_Validate(Cancel As Boolean)
    If TextLongValidate(255, txtHostName(Index).Text) Then Cancel = True
End Sub

Private Sub txtHostName_Validate(Index As Integer, Cancel As Boolean)
    If TextLongValidate(120, txtHostName(Index).Text) Then Cancel = True
End Sub

Private Sub txtPass_Validate(Cancel As Boolean)
    If TextLongValidate(8, txtPass.Text) Then Cancel = True
End Sub

Private Sub txtUserID_Validate(Cancel As Boolean)
    Dim nLen As Integer, i As Integer
    Dim nMsg As Integer

    
    If TextLongValidate(40, txtUserID.Text) Then Cancel = True
    If SpacialStrValid(txtUserID.Text, "[") Then Cancel = True
    If SpacialStrValid(txtUserID.Text, ",") Then Cancel = True
    If SpacialStrValid(txtUserID.Text, "]") Then Cancel = True

    
    
    nLen = ArrayLength(g_atUserInfoDelTag)
    
    If nLen > 0 Then
        For i = 1 To nLen
            If g_atUserInfoDelTag(i).UserID = txtUserID.Text Then
                nMsg = MsgBox("此用户代码被一个已删除的用户使用,此用户可恢复,是否恢复?", vbQuestion + vbYesNo, cszMsg)
                If nMsg = vbYes Then
                    RecoverUser
                    Exit Sub
                Else
                    Cancel = True
                End If
                Exit For
            End If
        Next i
    End If

    
    nLen = ArrayLength(g_atUserInfo)
    
    If nLen > 0 Then
        For i = 1 To nLen
            If g_atUserInfo(i).UserID = txtUserID.Text Then
                MsgBox "此用户代码已使用,请另选一个重试!", vbInformation, cszMsg
                Cancel = True
                Exit Sub
            End If
        Next i
    End If
    
    
    
    nLen = ArrayLength(g_atAllUserInfo)
    
    If nLen > 0 Then
        For i = 1 To nLen
            If g_atAllUserInfo(i).UserID = txtUserID.Text Then
                MsgBox "此用户代码被一个已删除的用户使用,且此用户所属单位已删除,单位代码为" & g_atAllUserInfo(i).UnitID _
                    & "要恢复此用户必须先恢复所属单位!", vbExclamation, cszMsg
                Cancel = True
                Exit For
            End If
        Next i
    End If
                    

    

    
End Sub

Private Sub txtUserName_Validate(Cancel As Boolean)
    If TextLongValidate(100, txtUserName.Text) Then Cancel = True
    If SpacialStrValid(txtUserName.Text, "[") Then Cancel = True
    If SpacialStrValid(txtUserName.Text, ",") Then Cancel = True
    If SpacialStrValid(txtUserName.Text, "]") Then Cancel = True

End Sub


Private Sub LoadData()
    Dim i As Integer, j As Integer, k As Integer, nInStrStart As Integer
    lblUserID.Caption = g_alvItemText(1)
    szUserID = g_alvItemText(1)
    For i = 1 To ArrayLength(g_atUserInfo)
        If g_atUserInfo(i).UserID = g_alvItemText(1) Then '取得内存中对应用户的属性
            If g_atUserInfo(i).InnerUser = True Then
                bInnerUser = True
            End If
            
            chkInternet.Value = IIf(g_atUserInfo(i).IsInternetUser, Checked, Unchecked)
            txtUserName.Text = g_atUserInfo(i).UserName '用户名
            txtAnno = g_atUserInfo(i).Annotation '注释
            
            
            '使CboUnit显示此用户所属单位
            For j = 0 To cboUnit.ListCount - 1
                nInStrStart = 0
                 nInStrStart = InStr(1, cboUnit.List(j), g_atUserInfo(i).UnitID)
                If nInStrStart = 1 Then '使CboUnit显示此用户所属单位
                    cboUnit.ListIndex = j
                End If
            Next j
            
            ChangWeek (i) '处理内存中的星期数据
            If bWeek = True Then
                chkWeek.Value = Unchecked
            Else
                chkWeek.Value = Checked
                For k = 0 To 6
                    If nWN(k) = 1 Then
                        chkWorkDate(k).Value = Checked
                    Else
                        chkWorkDate(k).Value = Unchecked
                    End If
                Next k
            End If
            
            If g_atUserInfo(i).MultLogin = -1 Then '处理多次登录显示,及是否登录
                chkEnable.Value = Checked
                chkLoginCount.Enabled = True
                chkLoginCount.Value = Checked
            ElseIf g_atUserInfo(i).MultLogin = 0 Then
                chkEnable.Value = Unchecked
                chkLoginCount.Enabled = False
                chkLoginCount.Value = Unchecked
            ElseIf g_atUserInfo(i).MultLogin = 1 Then
                chkEnable.Value = Checked
                chkLoginCount.Enabled = True
                chkLoginCount.Value = Unchecked
            End If
            
            '用户登录时间
            If g_atUserInfo(i).StartTime = "0:00:00" And g_atUserInfo(i).EndTime = "0:00:00" Then
                chkLoginTime.Value = Unchecked
            Else
                chkLoginTime.Value = Checked
                If g_atUserInfo(i).StartTime = "0:00:00" Then
                    dtpBeginTime.Enabled = True
                    dtpBeginTime.Value = g_atUserInfo(i).StartTime + "0:00:01"
                Else
                    dtpBeginTime.Enabled = True
                    dtpBeginTime.Value = g_atUserInfo(i).StartTime
                End If
                If g_atUserInfo(i).EndTime = "0:00:00" Then
                    dtpEndTime.Enabled = True
                    dtpEndTime.Value = "23:59:59"
                Else
                    dtpEndTime.Enabled = True
                    dtpEndTime.Value = g_atUserInfo(i).EndTime
                End If
            End If

            '用户登录IP
            If g_atUserInfo(i).LoginIP <> "" Then
                chkWorkStation.Value = Checked
                DisPlayIPControl
                ''''****数据
                GetAndDisPlayIp (g_atUserInfo(i).LoginIP)
            Else
                chkWorkStation.Value = Unchecked
                DisPlayIPControl
            End If
            txtCanSellDay.Text = g_atUserInfo(i).CanSellDay
            Exit For
        End If
    Next i
End Sub

Private Sub ChangWeek(nTemp As Integer)
    Dim nWeekNum As Integer
    nWeekNum = g_atUserInfo(nTemp).Week
    Dim i As Integer
    If nWeekNum = 0 Then
        bWeek = True
    Else
        bWeek = False
        For i = 0 To 6
            nWN(i) = 0
        Next
        
        If nWeekNum < 128 Then
            For i = 6 To 0 Step -1
               If nWeekNum <= DenbiFunction(i) And nWeekNum > DenbiFunction(i - 1) Then
                    nWN(i) = 1
                    nWeekNum = nWeekNum - (2 ^ i)
                
                End If
            Next i
        Else
            For i = 0 To 6
                nWN(i) = 1
            Next
        End If
    End If
End Sub

Private Function DenbiFunction(n As Integer)
    Dim i As Integer
    Dim nTemp As Integer
    nTemp = 0
    If n > -1 Then
        For i = 0 To n
            nTemp = nTemp + (2 ^ i)
        Next
    Else
        nTemp = 0
    End If
    DenbiFunction = nTemp
End Function

Private Sub DisPlayIPControl()
    Dim i As Integer
    EnableContainer fraWorkStation, IIf(chkWorkStation.Value = 1, True, False), chkWorkStation
    If chkWorkStation.Value = Checked Then
        For i = 0 To 5
            txtHostName(i).BackColor = vbWhite
        Next i
    Else
        For i = 0 To 5
            txtHostName(i).BackColor = cGrayForHost
        Next i
    End If
End Sub

Private Sub GetAndDisPlayIp(IPs As String)
    Dim aszTemp() As String, nIPCount As Integer, i As Integer, j As Integer
    szIP = IPs
    Dim szIPPart As String
    Dim aszIPPart() As String
On Error GoTo ErrorHandle
    aszTemp = GetIPString(szIP)
    nIPCount = ArrayLength(aszTemp)
    If (nIPCount > 0) And (nIPCount < 7) Then
        For i = 0 To nIPCount - 1
            szIPPart = aszTemp(i + 1)
            txtHostName(i).Text = szIPPart
        Next i
        
    ElseIf nIPCount > 6 Then
        For i = 1 To 6
            szIPPart = aszTemp(i)
            txtHostName(i - 1).Text = szIPPart
        Next i
    Else
    End If
Exit Sub
ErrorHandle:
    MsgBox "数据库中存在非法IP地址,与数据库管理员联系.", vbExclamation, cszMsg
    
End Sub

Private Sub GetDateFormUI()
    Dim i As Integer
    Dim aszTemp() As String
    
    bIsInternetUser = IIf(chkInternet.Value = Checked, True, False)
    
    '得到IPAddress
    
    If chkWorkStation = Unchecked Then
        szIP = ""
    Else
        szIP = ""
        For i = 0 To 5
                If szIP <> "" Then
                    If txtHostName(i).Text <> "" Then
                        szIP = szIP & "," & txtHostName(i).Text
                    End If
                Else
                    If txtHostName(i).Text <> "" Then
                        szIP = txtHostName(i).Text
                    End If
                End If
        Next i
    End If
    
    '得到注释
    szAnno = txtAnno
    
    
    '得到时间
    If chkLoginTime = Unchecked Then
        dtStartTime = "0:00:01"
        dtEndTime = "23:59:59"
    Else
        dtStartTime = dtpBeginTime.Value
        dtEndTime = dtpEndTime.Value
    End If
    
    '得到用户代码
    If bEdit = True Then
        szUserID = lblUserID
    Else
        szUserID = txtUserID
    End If
    
    '得到用户名
    szUserName = txtUserName.Text
    
    '得到多次登录及是否登录
    If chkEnable = Checked Then
        If chkLoginCount = Checked Then
            nMultLogin = -1
        Else
            nMultLogin = 1
        End If
    Else
        nMultLogin = 0
    End If
    
    '得到UnitId
    szUnitID = GetUnitID(cboUnit.Text)
    szSellStationID = ResolveDisplay(cboStation.Text)
    '得到week
    If chkWeek = Unchecked Then
        nWeek = 0
    Else
        nWeek = 0
        For i = 0 To 6
            If chkWorkDate(i) = Checked Then
                nWN(i) = 1
            Else
                nWN(i) = 0
            End If
        Next i
        For i = 0 To 6
            nWeek = nWeek + nWN(i) * (2 ^ i)
        Next i
    End If
    '得到可售天数
    If IsNumeric(txtCanSellDay.Text) Then
        m_nCanSellDay = txtCanSellDay.Text
    Else
        m_nCanSellDay = 0
    End If
    
    
End Sub

Private Function GetUnitID(cboText As String) As String
    Dim nTemp As Integer
    nTemp = InStr(1, cboText, "[")
    GetUnitID = Left(cboText, nTemp - 1)
End Function

Private Sub RecoverUser()
    Dim oUser As New User
    Dim nTemp As Integer
    
    On Error GoTo ErrorHandle
    oUser.Init g_oActUser
    oUser.Identify txtUserID
    oUser.ReCover
    frmStoreMenu.LoadCommonData
    nTemp = ArrayLength(g_atUserInfo)
    
    frmStoreMenu.DisplayUserInfo nTemp
'    g_alvItemText(1) = txtUserID
    
    txtUserID = ""
    bEdit = True
    Me.Caption = "修改用户属性"
    txtUserID.Visible = False
    lblUserID.Visible = True
    cmdAEUser.Caption = "确定(&O)"
    LoadData
Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub
