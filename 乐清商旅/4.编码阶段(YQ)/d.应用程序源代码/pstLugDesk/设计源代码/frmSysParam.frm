VERSION 5.00
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Begin VB.Form frmSysParam 
   BackColor       =   &H00E0E0E0&
   Caption         =   "设置系统参数"
   ClientHeight    =   2700
   ClientLeft      =   4680
   ClientTop       =   1560
   ClientWidth     =   6075
   Icon            =   "frmSysParam.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2700
   ScaleWidth      =   6075
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "行包单/签发单号长度设置"
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   5805
      Begin VB.TextBox txtCheckLuggageIDLen 
         Height          =   315
         Left            =   1560
         TabIndex        =   8
         Top             =   1410
         Width           =   1485
      End
      Begin VB.TextBox txtLuggageIDNumber 
         Height          =   315
         Left            =   4320
         TabIndex        =   7
         Top             =   255
         Width           =   1215
      End
      Begin VB.TextBox txtLuggageIDPrefix 
         Height          =   315
         Left            =   1560
         TabIndex        =   6
         Top             =   255
         Width           =   1485
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   300
         X2              =   5490
         Y1              =   1305
         Y2              =   1305
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   300
         X2              =   5490
         Y1              =   1290
         Y2              =   1290
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "前缀部分用于标识车票印刷的批号,一般为3位,如(99A),数字部分为票号的阿拉伯数字部分,一般为7位。如:     99A-0123456 (前辍-数字)"
         Height          =   615
         Left            =   750
         TabIndex        =   5
         Top             =   630
         Width           =   4830
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "说明:"
         Height          =   180
         Left            =   270
         TabIndex        =   4
         Top             =   630
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "前缀部分(P):"
         Height          =   180
         Left            =   270
         TabIndex        =   3
         Top             =   330
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "数字部分(&N):"
         Height          =   180
         Left            =   3210
         TabIndex        =   2
         Top             =   330
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "签发单长度(&C):"
         Height          =   180
         Left            =   270
         TabIndex        =   1
         Top             =   1470
         Width           =   1260
      End
   End
   Begin RTComctl3.CoolButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   345
      Left            =   4800
      TabIndex        =   9
      Top             =   2220
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   609
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
      MICON           =   "frmSysParam.frx":000C
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
      Height          =   345
      Left            =   3540
      TabIndex        =   10
      Top             =   2220
      Width           =   1125
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "确定"
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
      MICON           =   "frmSysParam.frx":0028
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
Attribute VB_Name = "frmSysParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

