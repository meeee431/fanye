VERSION 5.00
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmBookAttr 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "预定信息"
   ClientHeight    =   5415
   ClientLeft      =   3990
   ClientTop       =   2775
   ClientWidth     =   5250
   Icon            =   "frmBookAttr.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   4140
      Left            =   315
      ScaleHeight     =   4140
      ScaleWidth      =   4515
      TabIndex        =   2
      Top             =   570
      Width           =   4515
      Begin VB.Label lblOperateTime 
         AutoSize        =   -1  'True
         Caption         =   "操作时间:"
         Height          =   180
         Left            =   285
         TabIndex        =   16
         Top             =   2484
         Width           =   810
      End
      Begin VB.Label lblCancelOperation 
         AutoSize        =   -1  'True
         Caption         =   "取消:"
         Height          =   180
         Left            =   270
         TabIndex        =   15
         Top             =   3840
         Width           =   450
      End
      Begin VB.Label lblOperation 
         AutoSize        =   -1  'True
         Caption         =   "操作员:"
         Height          =   180
         Left            =   270
         TabIndex        =   14
         Top             =   3555
         Width           =   630
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000005&
         X1              =   60
         X2              =   4485
         Y1              =   3390
         Y2              =   3390
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000003&
         X1              =   60
         X2              =   4485
         Y1              =   3375
         Y2              =   3375
      End
      Begin VB.Label lblMemo 
         Caption         =   "备注:"
         Height          =   480
         Left            =   270
         TabIndex        =   13
         Top             =   2865
         Width           =   4185
      End
      Begin VB.Label lblAdd 
         AutoSize        =   -1  'True
         Caption         =   "地址:"
         Height          =   180
         Left            =   2670
         TabIndex        =   12
         Top             =   2103
         Width           =   450
      End
      Begin VB.Label lblTelephone 
         AutoSize        =   -1  'True
         Caption         =   "电话:"
         Height          =   180
         Left            =   270
         TabIndex        =   11
         Top             =   2103
         Width           =   450
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "姓名:"
         Height          =   180
         Left            =   2670
         TabIndex        =   10
         Top             =   1722
         Width           =   450
      End
      Begin VB.Label lblDestStation 
         AutoSize        =   -1  'True
         Caption         =   "到站:"
         Height          =   180
         Left            =   270
         TabIndex        =   9
         Top             =   1722
         Width           =   450
      End
      Begin VB.Label lblStauts 
         AutoSize        =   -1  'True
         Caption         =   "状态:"
         Height          =   180
         Left            =   2670
         TabIndex        =   8
         Top             =   1341
         Width           =   450
      End
      Begin VB.Label lblBookID 
         AutoSize        =   -1  'True
         Caption         =   "预定号:"
         Height          =   180
         Left            =   2670
         TabIndex        =   7
         Top             =   240
         Width           =   630
      End
      Begin VB.Label lblSeatNumber 
         AutoSize        =   -1  'True
         Caption         =   "座位号:"
         Height          =   180
         Left            =   270
         TabIndex        =   6
         Top             =   1341
         Width           =   630
      End
      Begin VB.Label lblStartTime 
         AutoSize        =   -1  'True
         Caption         =   "发车时间:"
         Height          =   180
         Left            =   2670
         TabIndex        =   5
         Top             =   960
         Width           =   810
      End
      Begin VB.Label lblBus 
         AutoSize        =   -1  'True
         Caption         =   "代码:"
         Height          =   180
         Left            =   270
         TabIndex        =   4
         Top             =   960
         Width           =   450
      End
      Begin VB.Label lblID 
         AutoSize        =   -1  'True
         Caption         =   "流水号:"
         Height          =   180
         Left            =   1245
         TabIndex        =   3
         Top             =   225
         Width           =   630
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         X1              =   45
         X2              =   6105
         Y1              =   795
         Y2              =   795
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         X1              =   30
         X2              =   6120
         Y1              =   780
         Y2              =   780
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   210
         Picture         =   "frmBookAttr.frx":000C
         Top             =   105
         Width           =   240
      End
   End
   Begin RTComctl3.CoolButton cmdClose 
      Cancel          =   -1  'True
      Default         =   -1  'True
      Height          =   315
      Left            =   4005
      TabIndex        =   1
      Top             =   4965
      Width           =   1065
      _ExtentX        =   0
      _ExtentY        =   0
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
      MICON           =   "frmBookAttr.frx":0E4E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4695
      Left            =   105
      TabIndex        =   0
      Top             =   135
      Width           =   5010
      _ExtentX        =   8837
      _ExtentY        =   8281
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "预定属性"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmBookAttr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
Unload Me
End Sub

