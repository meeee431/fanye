VERSION 5.00
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Begin VB.Form frmReturnAccept 
   BackColor       =   &H00808080&
   Caption         =   "退受理单"
   ClientHeight    =   5775
   ClientLeft      =   1575
   ClientTop       =   4035
   ClientWidth     =   9525
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   15
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   7000030
   Icon            =   "fraReturnAccept.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5775
   ScaleWidth      =   9525
   Begin VB.Frame fraOutLine 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5265
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   8895
      Begin VB.ComboBox cboFeesRatio 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "fraReturnAccept.frx":179A
         Left            =   7590
         List            =   "fraReturnAccept.frx":179C
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox txtCredenceID 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7380
         TabIndex        =   35
         Top             =   3960
         Width           =   1185
      End
      Begin VB.TextBox txtLuggageID 
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
         Left            =   1830
         TabIndex        =   32
         Top             =   225
         Width           =   2490
      End
      Begin VB.Frame fraTktInfoChange 
         BackColor       =   &H00E0E0E0&
         Caption         =   "行包票信息"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4275
         Left            =   210
         TabIndex        =   1
         Top             =   750
         Width           =   5925
         Begin VB.Label lblLuggageName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   300
            Left            =   1380
            TabIndex        =   46
            Top             =   300
            Width           =   150
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "行包名称:"
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
            Left            =   180
            TabIndex        =   45
            Top             =   360
            Width           =   1080
         End
         Begin VB.Label Label13 
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
            Left            =   2790
            TabIndex        =   31
            Top             =   2055
            Width           =   1080
         End
         Begin VB.Label lblOverNumber 
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
            Left            =   3900
            TabIndex        =   30
            Top             =   2055
            Width           =   120
         End
         Begin VB.Label lblAcceptType 
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
            Left            =   1350
            TabIndex        =   29
            Top             =   750
            Width           =   120
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "托运方式:"
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
            Left            =   180
            TabIndex        =   28
            Top             =   750
            Width           =   1080
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "件数:"
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
            Left            =   210
            TabIndex        =   27
            Top             =   2055
            Width           =   600
         End
         Begin VB.Label lblBagNumber 
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
            Left            =   1080
            TabIndex        =   26
            Top             =   2050
            Width           =   120
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "标签号:"
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
            Left            =   210
            TabIndex        =   25
            Top             =   3300
            Width           =   840
         End
         Begin VB.Label lblLabelID 
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
            Left            =   1050
            TabIndex        =   24
            Top             =   3270
            Width           =   120
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "实重:"
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
            Left            =   2790
            TabIndex        =   23
            Top             =   1620
            Width           =   600
         End
         Begin VB.Label lblActWeight 
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
            Left            =   3900
            TabIndex        =   22
            Top             =   1620
            Width           =   120
         End
         Begin VB.Label lblMileage 
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
            Height          =   300
            Left            =   3420
            TabIndex        =   21
            Top             =   690
            Width           =   120
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "里程:"
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
            Left            =   2760
            TabIndex        =   20
            Top             =   720
            Width           =   600
         End
         Begin VB.Label lblPicker 
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
            Left            =   3900
            TabIndex        =   19
            Tag             =   "提货人"
            Top             =   2910
            Width           =   120
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "提取人:"
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
            Left            =   2790
            TabIndex        =   18
            Top             =   2910
            Width           =   840
         End
         Begin VB.Label lblShipper 
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
            Left            =   1080
            TabIndex        =   17
            Top             =   2910
            Width           =   120
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "托运人:"
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
            Left            =   210
            TabIndex        =   16
            Top             =   2910
            Width           =   840
         End
         Begin VB.Label lblCalWeight 
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
            Left            =   1080
            TabIndex        =   15
            Top             =   1620
            Width           =   120
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "计重:"
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
            Left            =   210
            TabIndex        =   14
            Top             =   1620
            Width           =   600
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "票价:"
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
            Left            =   240
            TabIndex        =   13
            Top             =   3750
            Width           =   600
         End
         Begin VB.Label lblTimeChange 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "受理时间:"
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
            Left            =   2790
            TabIndex        =   12
            Top             =   2475
            Width           =   1080
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "起点站:"
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
            Index           =   1
            Left            =   180
            TabIndex        =   11
            Top             =   1170
            Width           =   840
         End
         Begin VB.Label label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "到站:"
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
            Left            =   2790
            TabIndex        =   10
            Top             =   1170
            Width           =   600
         End
         Begin VB.Label lblStateChange 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "状态:"
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
            Left            =   2790
            TabIndex        =   9
            Top             =   3750
            Width           =   600
         End
         Begin VB.Label lblOperatorChange 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "操作员:"
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
            Left            =   210
            TabIndex        =   8
            Top             =   2475
            Width           =   840
         End
         Begin VB.Label lblEndStation 
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
            Left            =   3510
            TabIndex        =   7
            Top             =   1170
            Width           =   120
         End
         Begin VB.Label lblStartStation 
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
            Height          =   300
            Left            =   1200
            TabIndex        =   6
            Top             =   1140
            Width           =   195
         End
         Begin VB.Label lblOperater 
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
            Left            =   1080
            TabIndex        =   5
            Top             =   2480
            Width           =   120
         End
         Begin VB.Label lblOperationTime 
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
            Left            =   3900
            TabIndex        =   4
            Top             =   2475
            Width           =   120
         End
         Begin VB.Label lblTicketPrice 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
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
            Height          =   240
            Left            =   1150
            TabIndex        =   3
            Top             =   3750
            Width           =   495
         End
         Begin VB.Label lblStatus 
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
            Left            =   3570
            TabIndex        =   2
            Top             =   3750
            Width           =   120
         End
      End
      Begin RTComctl3.CoolButton cmdCancelAccept 
         Height          =   615
         Left            =   6300
         TabIndex        =   33
         Top             =   4410
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "退票(&R)"
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
         MICON           =   "fraReturnAccept.frx":179E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         X1              =   6240
         X2              =   8700
         Y1              =   3780
         Y2              =   3780
      End
      Begin VB.Line Line4 
         Index           =   0
         X1              =   6240
         X2              =   8700
         Y1              =   2580
         Y2              =   2580
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "应退票款:"
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
         Left            =   6330
         TabIndex        =   44
         Top             =   3495
         Width           =   945
      End
      Begin VB.Label lblCredenceID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "凭证号(&C):"
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
         Left            =   6300
         TabIndex        =   43
         Top             =   4020
         Width           =   1050
      End
      Begin VB.Label lblReturnCharge 
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
         Left            =   7980
         TabIndex        =   42
         Top             =   2985
         Width           =   585
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "退款手续费:"
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
         Left            =   6330
         TabIndex        =   41
         Top             =   3060
         Width           =   1155
      End
      Begin VB.Label lblFree 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "手续费比例(%):"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   6330
         TabIndex        =   40
         Top             =   2250
         Width           =   1260
      End
      Begin VB.Label lblCurrectTicketPrice 
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
         Height          =   345
         Left            =   7980
         TabIndex        =   39
         Top             =   2670
         Width           =   585
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "票价:"
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
         Left            =   6330
         TabIndex        =   38
         Top             =   2745
         Width           =   525
      End
      Begin VB.Label lblTotalReturnMoney 
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
         Left            =   7980
         TabIndex        =   37
         Top             =   3420
         Width           =   585
      End
      Begin VB.Line Line1 
         X1              =   6240
         X2              =   8700
         Y1              =   3390
         Y2              =   3390
      End
      Begin VB.Label lblOldTktNum 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "退行包单号(&N):"
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
         Left            =   210
         TabIndex        =   34
         Top             =   300
         Width           =   1470
      End
   End
End
Attribute VB_Name = "frmReturnAccept"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboFeesRatio_Change()
'On Error GoTo ErrHandle
'
'   lblCurrectTicketPrice.Caption = lblTicketPrice.Caption
'   lblReturnCharge.Caption = FormatMoney(lblCurrectTicketPrice.Caption * cboFeesRatio.Text / 100)
'   lblTotalReturnMoney.Caption = FormatMoney(lblCurrectTicketPrice - lblReturnCharge.Caption)
'
'Exit Sub
'ErrHandle:
'ShowErrorMsg
End Sub

Private Sub cboFeesRatio_Click()
On Error GoTo ErrHandle
    If txtLuggageID.Text = "" Then Exit Sub
    
   lblCurrectTicketPrice.Caption = lblTicketPrice.Caption
   lblReturnCharge.Caption = Round(FormatMoney(lblCurrectTicketPrice.Caption * cboFeesRatio.Text / 100))
   lblTotalReturnMoney.Caption = Round(FormatMoney(lblCurrectTicketPrice - lblReturnCharge.Caption))
'   txtCredenceID.SetFocus
Exit Sub
ErrHandle:
ShowErrorMsg
End Sub

Private Sub cboFeesRatio_GotFocus()
lblFree.ForeColor = clActiveColor
End Sub

Private Sub cboFeesRatio_LostFocus()
lblFree.ForeColor = 0
End Sub

Private Sub cmdCancelAccept_Click()
On Error GoTo ErrHandle
Dim sAnswer, sAnswer1
Dim tPriceValue() As TLuggagePriceItem
Dim nlen As Integer
Dim i As Integer
Dim mdbShippePrice As Double
Dim mdbSerPrice As Double
Dim mdbPickPrice As Double
If txtCredenceID.Text = "" Then
      MsgBox "  凭证号不能为空，请输入凭证号！", vbInformation + vbOKOnly, "行包退运"
      txtCredenceID.SetFocus
      
Else
   sAnswer = MsgBox("  您确认要退此行包?", vbInformation + vbYesNo, "行包退运")
   If sAnswer = vbYes Then
    moLugSvr.ReturnAcceptSheet Trim(txtLuggageID.Text), Trim(txtCredenceID.Text), lblReturnCharge
    lblStatus.ForeColor = vbRed
    lblStatus.Caption = "已退"
    cmdCancelAccept.Enabled = False
    txtCredenceID.Enabled = False
    txtLuggageID.SetFocus
    txtCredenceID.Text = ""
    '打印退单处理
    If moSysParam.EnabledPrintReturnSheet Then
        PrintReturnAccept moAcceptSheet, Trim(txtLuggageID.Text), Val(lblReturnCharge), m_oAUser.UserName, Now
        IncTicketNo
    End If
    moAcceptSheet.AddNew
'    moAcceptSheet.SheetID = g_szAcceptSheetID
 End If
End If
 
Exit Sub
ErrHandle:
 ShowErrorMsg
End Sub

Private Sub Form_Activate()
SetSheetNoLabel True, g_szAcceptSheetID
FormClear
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
     txtLuggageID.Text = ""
     txtLuggageID.SetFocus
     FormClear
  End If
End Sub
Private Sub FormClear()
    txtLuggageID.Text = ""
    
    lblLuggageName.Caption = ""
    lblStartStation.Caption = ""
    lblEndStation.Caption = ""
    lblMileage.Caption = ""
    lblAcceptType.Caption = ""
    lblLabelID.Caption = ""
    lblCalWeight.Caption = ""
    lblActWeight.Caption = ""
    lblBagNumber.Caption = ""
    lblOverNumber.Caption = ""
    lblOperater.Caption = ""
    lblOperationTime.Caption = ""
    lblShipper.Caption = ""
    lblPicker.Caption = ""
    lblTicketPrice.Caption = ""
    lblStatus.Caption = ""
    
    lblCurrectTicketPrice.Caption = "0.00"
    lblReturnCharge.Caption = "0.00"
    lblTotalReturnMoney.Caption = "0.00"
    
    txtCredenceID.Text = ""
End Sub

Private Sub Form_Load()

    AlignFormPos Me
    HideSheetNoLabel
    FormClear
    cmdCancelAccept.Enabled = False
    
   
         '得到费率
     GetFeesRatio
End Sub

Private Sub Form_Resize()
    If mdiMain.ActiveForm Is Me Then
        If Not Me.WindowState = vbMaximized Then Me.WindowState = vbMaximized
        fraOutLine.Left = (Me.ScaleWidth - fraOutLine.Width) / 2
        fraOutLine.Top = (Me.ScaleHeight - fraOutLine.Height) / 2
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveFormPos Me
End Sub

Private Sub txtCredenceID_GotFocus()
lblCredenceID.ForeColor = clActiveColor

End Sub

Private Sub txtCredenceID_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtCredenceID.Text <> "" Then
        If KeyCode = vbKeyReturn Then
           cmdCancelAccept_Click
            
        End If
    Else
        
        txtCredenceID.SetFocus
    End If
    
        
End Sub

Private Sub txtCredenceID_LostFocus()
lblCredenceID.ForeColor = 0
End Sub

Private Sub txtLuggageID_GotFocus()
lblOldTktNum.ForeColor = clActiveColor
lblCredenceID.ForeColor = 0
End Sub

Private Sub txtLuggageID_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrHandle

    If KeyCode = vbKeyReturn Then
        If txtLuggageID.Text = "" Then Exit Sub
        moAcceptSheet.Identify Trim(txtLuggageID.Text)
        If moAcceptSheet.Status = 0 Then
            cmdCancelAccept.Enabled = True
            txtCredenceID.Enabled = True
            txtCredenceID.SetFocus
        Else
            txtCredenceID.Enabled = False
            cmdCancelAccept.Enabled = False
        End If
        RefreshAccept
        '计算退运费
        cboFeesRatio_Click
    End If

Exit Sub
ErrHandle:
ShowErrorMsg
End Sub
Private Sub GetFeesRatio()
On Error GoTo ErrHandle
 Dim mValue As String
 Dim nlen As Integer
 Dim i As Integer
 '以后需要更改,现在只处理一个费率项
'      nLen = ArrayLength(moSysParam.LuggageReturnRatio1)
'
'     If nLen > 0 Then
'        ReDim mValue(1 To nLen)
        mValue = moSysParam.LuggageReturnRatio1
'       For i = 1 To nLen
        cboFeesRatio.AddItem mValue * 100
        cboFeesRatio.AddItem 0
        cboFeesRatio.ListIndex = 0
'        cboFeesRatio.Index = 0
'       Next i
'     End If
'        For i = 1 To nLen
'           If mValue(i).sgReturnRate = 0 Then Exit For
'        Next i
'        If i > nLen Then cboFeesRatio.AddItem "0"
'
'        For i = 1 To nLen
'            If mValue(i).sgReturnRate = 100 Then Exit For
'        Next i
'        If i > nLen Then cboFeesRatio.AddItem "100"
'
 
 Exit Sub
ErrHandle:
 ShowErrorMsg
End Sub
Private Sub RefreshAccept()
 On Error GoTo ErrHandle
    lblLuggageName.Caption = Trim(moAcceptSheet.LuggageName)
    lblStartStation.Caption = Trim(moAcceptSheet.StartStationName)
    lblEndStation.Caption = Trim(moAcceptSheet.DesStationName)
    lblMileage.Caption = CStr(moAcceptSheet.Mileage)
    lblAcceptType.Caption = Trim(moAcceptSheet.AcceptType)
    lblLabelID.Caption = Trim(moAcceptSheet.StartLabelID)
    lblCalWeight.Caption = CStr(moAcceptSheet.CalWeight)
    lblActWeight.Caption = CStr(moAcceptSheet.ActWeight)
    lblBagNumber.Caption = CStr(moAcceptSheet.Number)
    lblOverNumber.Caption = CStr(moAcceptSheet.OverNumber)
    lblOperater.Caption = Trim(moAcceptSheet.Operator)
    lblOperationTime.Caption = CStr(Format(moAcceptSheet.OperateTime, "yyyy-MM-dd hh:mm"))
    lblShipper.Caption = Trim(moAcceptSheet.Shipper)
    lblPicker.Caption = Trim(moAcceptSheet.Picker)
    lblTicketPrice.Caption = CStr(moAcceptSheet.TotalPrice)
    If moAcceptSheet.Status <> 0 Then
       lblStatus.ForeColor = vbRed
    Else
       lblStatus.ForeColor = 0
    End If
    lblStatus.Caption = Trim(moAcceptSheet.StatusString)
    
Exit Sub
ErrHandle:
ShowErrorMsg
End Sub

Private Sub txtLuggageID_LostFocus()
lblOldTktNum.ForeColor = 0
End Sub
