VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BBF95DAA-F9CB-4CA9-A673-E0E9E0193752}#1.0#0"; "STSellCtl.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmAccept 
   BackColor       =   &H8000000C&
   Caption         =   "行包受理"
   ClientHeight    =   7080
   ClientLeft      =   1395
   ClientTop       =   2490
   ClientWidth     =   12465
   ControlBox      =   0   'False
   HelpContextID   =   7000020
   Icon            =   "frmAccept.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7080
   ScaleWidth      =   12465
   WindowState     =   2  'Maximized
   Begin VB.Frame fradetail 
      BackColor       =   &H00E0E0E0&
      Caption         =   "运费明细"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3915
      Left            =   13275
      TabIndex        =   81
      Top             =   4680
      Visible         =   0   'False
      Width           =   2325
      Begin MSComctlLib.ListView lvprice 
         Height          =   3495
         Left            =   120
         TabIndex        =   82
         Top             =   270
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   6165
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   " 计重"
            Object.Width           =   1694
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   " 运费"
            Object.Width           =   1711
         EndProperty
      End
   End
   Begin VB.Frame FraPrice 
      BackColor       =   &H00E0E0E0&
      Height          =   2295
      Left            =   13275
      TabIndex        =   74
      Top             =   2310
      Visible         =   0   'False
      Width           =   2325
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "运费计算"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1395
         Left            =   120
         TabIndex        =   75
         Top             =   750
         Width           =   2055
         Begin VB.TextBox txtPrice 
            Height          =   315
            Left            =   1080
            TabIndex        =   77
            Text            =   "0"
            Top             =   720
            Width           =   825
         End
         Begin VB.TextBox txtCal 
            Height          =   315
            Left            =   120
            TabIndex        =   76
            Text            =   "0"
            Top             =   720
            Width           =   825
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "运费"
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
            Left            =   1200
            TabIndex        =   79
            Top             =   360
            Width           =   480
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "计重"
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
            TabIndex        =   78
            Top             =   360
            Width           =   480
         End
      End
      Begin RTComctl3.CoolButton cmdOver 
         Height          =   345
         Left            =   360
         TabIndex        =   80
         Top             =   270
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "<<完成运费计算(&E)"
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
         MICON           =   "frmAccept.frx":076A
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
   Begin VB.Frame fraOutLine 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6735
      Left            =   120
      TabIndex        =   51
      Top             =   120
      Width           =   12165
      Begin VB.Frame Frame5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "   联网受理"
         Height          =   675
         Left            =   6360
         TabIndex        =   89
         Top             =   2640
         Visible         =   0   'False
         Width           =   3165
         Begin VB.ComboBox cboLinkAccept 
            Enabled         =   0   'False
            Height          =   300
            Left            =   960
            TabIndex        =   92
            Text            =   "Combo1"
            Top             =   240
            Width           =   2115
         End
         Begin VB.CheckBox chkLinkAccept 
            BackColor       =   &H00E0E0E0&
            Height          =   255
            Left            =   120
            TabIndex        =   91
            Top             =   0
            Width           =   255
         End
         Begin VB.ComboBox cboSellStation 
            Enabled         =   0   'False
            Height          =   300
            Left            =   960
            TabIndex        =   90
            Top             =   600
            Visible         =   0   'False
            Width           =   2115
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "所属车站:"
            Height          =   255
            Left            =   120
            TabIndex        =   94
            Top             =   645
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "所属单位:"
            Height          =   180
            Left            =   120
            TabIndex        =   93
            Top             =   285
            Width           =   810
         End
      End
      Begin VB.Frame fraSettle 
         BackColor       =   &H00E0E0E0&
         Caption         =   "应结运费"
         Height          =   1530
         Left            =   1350
         TabIndex        =   84
         Top             =   180
         Visible         =   0   'False
         Width           =   2760
         Begin VB.ComboBox cboSettleRatio 
            Height          =   300
            Left            =   1545
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   300
            Width           =   1035
         End
         Begin VB.CheckBox chkPrintSettlePrice 
            BackColor       =   &H00E0E0E0&
            Caption         =   "打印应结运费(&P)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   180
            TabIndex        =   45
            Top             =   1065
            Value           =   1  'Checked
            Width           =   2160
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "应结费率(&E):"
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
            Left            =   195
            TabIndex        =   43
            Top             =   345
            Width           =   1260
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "应结运费:"
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
            Left            =   195
            TabIndex        =   86
            Top             =   705
            Width           =   945
         End
         Begin VB.Label lblSettlePrice 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   1515
            TabIndex        =   85
            Top             =   720
            Width           =   105
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Height          =   1500
         Left            =   3360
         TabIndex        =   83
         Top             =   5115
         Width           =   5790
         Begin VB.TextBox txtAnnotation2 
            Height          =   510
            Left            =   1545
            TabIndex        =   42
            Text            =   "结算备注"
            Top             =   840
            Width           =   4080
         End
         Begin VB.TextBox txtAnnotation1 
            Height          =   525
            Left            =   1545
            TabIndex        =   40
            Text            =   "发票备注"
            Top             =   240
            Width           =   4080
         End
         Begin VB.Label lblAnnotation2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "结算联备注："
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
            Left            =   195
            TabIndex        =   41
            Top             =   930
            Width           =   1260
         End
         Begin VB.Label lblAnnotation1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "发票联备注："
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
            Left            =   195
            TabIndex        =   39
            Top             =   300
            Width           =   1260
         End
      End
      Begin MSComctlLib.ListView lvBus 
         Height          =   1590
         Left            =   90
         TabIndex        =   4
         Top             =   3480
         Width           =   3165
         _ExtentX        =   5583
         _ExtentY        =   2805
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin pstLugDesk.ucSuperCombo cboEndStation 
         Height          =   2385
         Left            =   90
         TabIndex        =   1
         Top             =   570
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   4207
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
      Begin RTComctl3.CoolButton cmdBagDetail 
         Height          =   345
         Left            =   6960
         TabIndex        =   68
         Top             =   405
         Visible         =   0   'False
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "行包明细>>"
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
         MICON           =   "frmAccept.frx":0786
         PICN            =   "frmAccept.frx":07A2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.ComboBox cboAcceptType 
         Height          =   300
         ItemData        =   "frmAccept.frx":0B3C
         Left            =   7440
         List            =   "frmAccept.frx":0B3E
         Style           =   2  'Dropdown List
         TabIndex        =   53
         Top             =   120
         Width           =   1575
      End
      Begin VB.TextBox txtReceivedMoney 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   10425
         TabIndex        =   49
         Text            =   "0.00"
         Top             =   3285
         Width           =   1365
      End
      Begin VB.Frame fraCustomerInfo 
         BackColor       =   &H00E0E0E0&
         Height          =   1845
         Left            =   3360
         TabIndex        =   59
         Top             =   3270
         Width           =   5790
         Begin VB.TextBox txtShippePhone 
            Height          =   315
            Left            =   4440
            TabIndex        =   30
            Text            =   "托电话"
            Top             =   210
            Width           =   1185
         End
         Begin VB.ComboBox cboPickType 
            Height          =   300
            ItemData        =   "frmAccept.frx":0B40
            Left            =   4440
            List            =   "frmAccept.frx":0B42
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Top             =   1050
            Width           =   1185
         End
         Begin VB.TextBox txtPickerAddress 
            Height          =   315
            Left            =   240
            TabIndex        =   36
            Text            =   "收地址"
            Top             =   1410
            Width           =   5385
         End
         Begin VB.TextBox txtPhone 
            Height          =   315
            Left            =   4440
            TabIndex        =   34
            Text            =   "收电话"
            Top             =   630
            Width           =   1185
         End
         Begin VB.TextBox txtPicker 
            Height          =   315
            Left            =   1560
            TabIndex        =   32
            Text            =   "收件人"
            Top             =   630
            Width           =   1185
         End
         Begin VB.TextBox txtShipper 
            Height          =   315
            Left            =   1560
            TabIndex        =   28
            Text            =   "托运人"
            Top             =   210
            Width           =   1185
         End
         Begin VB.Label lblShippePhone 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "托运人电话(&V):"
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
            Left            =   2925
            TabIndex        =   29
            Top             =   255
            Width           =   1470
         End
         Begin VB.Label lblPickType 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "交付方式(&H):"
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
            Left            =   2925
            TabIndex        =   37
            Top             =   1050
            Width           =   1260
         End
         Begin VB.Label lblShipper 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "托运人(&C):"
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
            TabIndex        =   27
            Top             =   255
            Width           =   1050
         End
         Begin VB.Label lblPickerAddress 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "收件人地址(&R):"
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
            Left            =   225
            TabIndex        =   35
            Top             =   1050
            Width           =   1470
         End
         Begin VB.Label lblPhone 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "收件人电话(&O):"
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
            Left            =   2925
            TabIndex        =   33
            Top             =   675
            Width           =   1470
         End
         Begin VB.Label lblPicker 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "收件人(&K):"
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
            Left            =   225
            TabIndex        =   31
            Top             =   675
            Width           =   1050
         End
      End
      Begin VB.Frame fraLuggage 
         BackColor       =   &H00E0E0E0&
         Caption         =   "行包信息"
         Height          =   2700
         Left            =   3360
         TabIndex        =   58
         Top             =   465
         Width           =   5775
         Begin VB.TextBox txtActWeight 
            Height          =   315
            Left            =   1560
            TabIndex        =   10
            Text            =   "4"
            Top             =   660
            Width           =   1185
         End
         Begin VB.TextBox txtOverNum 
            Height          =   315
            Left            =   4440
            TabIndex        =   16
            Text            =   "1"
            Top             =   1065
            Width           =   1185
         End
         Begin VB.TextBox txtBagNum 
            Height          =   315
            Left            =   1560
            TabIndex        =   14
            Text            =   "3"
            Top             =   1065
            Width           =   1185
         End
         Begin VB.TextBox txtCalWeight 
            Height          =   315
            Left            =   4440
            TabIndex        =   12
            Text            =   "10"
            Top             =   660
            Width           =   1185
         End
         Begin VB.TextBox txtInsuranceID 
            Height          =   315
            Left            =   4440
            TabIndex        =   20
            Text            =   "111111111"
            Top             =   1455
            Width           =   1185
         End
         Begin VB.TextBox txtTransTicketID 
            Height          =   315
            Left            =   1560
            TabIndex        =   18
            Text            =   "000000"
            Top             =   1455
            Width           =   1185
         End
         Begin VB.ComboBox cboPack 
            Height          =   300
            Left            =   1560
            TabIndex        =   26
            Top             =   2250
            Width           =   1200
         End
         Begin VB.ComboBox cboLuggageName 
            Height          =   300
            Left            =   1560
            TabIndex        =   8
            Top             =   270
            Width           =   4065
         End
         Begin VB.TextBox txtStartLabel 
            Height          =   315
            Left            =   1560
            TabIndex        =   22
            Text            =   "111"
            Top             =   1845
            Width           =   1185
         End
         Begin VB.TextBox txtEndLabel 
            Height          =   315
            Left            =   4440
            TabIndex        =   24
            Text            =   "113"
            Top             =   1860
            Width           =   1185
         End
         Begin VB.Label lblInsuranceID 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "保险单号(&J):"
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
            Left            =   3120
            TabIndex        =   19
            Top             =   1530
            Width           =   1260
         End
         Begin VB.Label lblTransTicketID 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "运输单号(&I):"
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
            Left            =   240
            TabIndex        =   17
            Top             =   1500
            Width           =   1260
         End
         Begin VB.Label lblPack 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "包装(&B):"
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
            Left            =   660
            TabIndex        =   25
            Top             =   2295
            Width           =   840
         End
         Begin VB.Label lblLuggageName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "行包名称(&M):"
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
            Left            =   240
            TabIndex        =   7
            Top             =   315
            Width           =   1260
         End
         Begin VB.Label lblEndLabel 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "结束标签(&L):"
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
            Left            =   3120
            TabIndex        =   23
            Top             =   1935
            Width           =   1260
         End
         Begin VB.Label lblOverNum 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "超重件数(&G):"
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
            Left            =   3120
            TabIndex        =   15
            Top             =   1129
            Width           =   1260
         End
         Begin VB.Label lblActWeight 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "实重(&A):"
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
            Left            =   660
            TabIndex        =   9
            Top             =   720
            Width           =   840
         End
         Begin VB.Label lblStartLabel 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "起始标签(&U):"
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
            Left            =   240
            TabIndex        =   21
            Top             =   1890
            Width           =   1260
         End
         Begin VB.Label lblBagNum 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "件数(&N):"
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
            Left            =   645
            TabIndex        =   13
            Top             =   1101
            Width           =   840
         End
         Begin VB.Label lblCalWeight 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "计重(&W):"
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
            Left            =   3540
            TabIndex        =   11
            Top             =   720
            Width           =   840
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Height          =   1410
         Left            =   105
         TabIndex        =   55
         Top             =   5025
         Width           =   3150
         Begin VB.ComboBox cboVehicle 
            Height          =   300
            Left            =   1335
            Style           =   2  'Dropdown List
            TabIndex        =   6
            ToolTipText     =   "F5键选择其他车辆"
            Top             =   150
            Width           =   1380
         End
         Begin VB.Label lblCompanyName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "绍兴汽运"
            Height          =   180
            Left            =   1335
            TabIndex        =   72
            Top             =   1170
            Width           =   720
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "参运公司:"
            Height          =   180
            Left            =   180
            TabIndex        =   71
            Top             =   1170
            Width           =   810
         End
         Begin VB.Label lblBusDate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "2004-9-11"
            Height          =   180
            Left            =   1335
            TabIndex        =   70
            Top             =   570
            Width           =   810
         End
         Begin VB.Label lblStartTime 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "10:30"
            Height          =   180
            Left            =   1335
            TabIndex        =   69
            Top             =   870
            Width           =   450
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "发车时间:"
            Height          =   180
            Left            =   180
            TabIndex        =   57
            Top             =   850
            Width           =   810
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "车次日期:"
            Height          =   180
            Left            =   180
            TabIndex        =   56
            Top             =   530
            Width           =   810
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "车辆(F5):"
            Height          =   180
            Left            =   180
            TabIndex        =   5
            Top             =   210
            Width           =   810
         End
      End
      Begin VSFlex7LCtl.VSFlexGrid vsPriceItem 
         Height          =   2115
         Left            =   9300
         TabIndex        =   47
         Top             =   450
         Width           =   2655
         _cx             =   4683
         _cy             =   3731
         _ConvInfo       =   -1
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   14737632
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16777215
         ForeColorSel    =   0
         BackColorBkg    =   14737632
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   5
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   3
         GridLinesFixed  =   5
         GridLineWidth   =   1
         Rows            =   10
         Cols            =   2
         FixedRows       =   0
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   -1  'True
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin STSellCtl.ucUpDownText txtPreSell 
         Height          =   315
         Left            =   1230
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   3075
         Width           =   555
         _ExtentX        =   979
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
      End
      Begin RTComctl3.CoolButton cmdAccept 
         Height          =   615
         Left            =   9300
         TabIndex        =   50
         Top             =   5220
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "受理(F2)"
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
         MICON           =   "frmAccept.frx":0B44
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin RTComctl3.FlatLabel flblSellDate 
         Height          =   315
         Left            =   2055
         TabIndex        =   66
         Top             =   3075
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   14737632
         OutnerStyle     =   2
         Caption         =   ""
      End
      Begin VB.Label lblBasePrice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Left            =   11610
         TabIndex        =   88
         Top             =   180
         Width           =   105
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "计算运费:"
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
         Left            =   10620
         TabIndex        =   87
         Top             =   180
         Width           =   945
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   90
         X2              =   3285
         Y1              =   2970
         Y2              =   2970
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   90
         X2              =   3330
         Y1              =   2955
         Y2              =   2955
      End
      Begin VB.Label lblInStation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "途经站"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   180
         TabIndex        =   67
         Top             =   6480
         Width           =   540
      End
      Begin VB.Label lblPriceItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "托运费用(F):"
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
         Left            =   9300
         TabIndex        =   46
         Top             =   180
         Width           =   1260
      End
      Begin VB.Label lblAcceptType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "托运方式(&T):"
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
         Left            =   6090
         TabIndex        =   52
         Top             =   180
         Width           =   1260
      End
      Begin VB.Label lblPreSellDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "预售天数(&D):"
         Height          =   180
         Left            =   105
         TabIndex        =   2
         Top             =   3135
         Width           =   1080
      End
      Begin VB.Label lblEndStation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "杭州"
         Height          =   180
         Left            =   1110
         TabIndex        =   65
         Top             =   195
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblMileage 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "29公里"
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
         Left            =   4050
         TabIndex        =   64
         Top             =   180
         Width           =   630
      End
      Begin VB.Label lblRestMoney 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "应找:"
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
         Left            =   9315
         TabIndex        =   63
         Top             =   3945
         Width           =   525
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00000000&
         X1              =   9210
         X2              =   11850
         Y1              =   3795
         Y2              =   3795
      End
      Begin VB.Label lblReceivedMoney 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "实收(&Q):"
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
         Left            =   9300
         TabIndex        =   48
         Top             =   3450
         Width           =   840
      End
      Begin VB.Label lblTotalPrice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "总价:"
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
         Left            =   9300
         TabIndex        =   62
         Top             =   2955
         Width           =   525
      End
      Begin VB.Label flblTotalPrice 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   10170
         TabIndex        =   61
         Top             =   2895
         Width           =   1575
      End
      Begin VB.Label flblRestMoney 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   10965
         TabIndex        =   60
         Top             =   3855
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "里程:"
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
         Index           =   1
         Left            =   3480
         TabIndex        =   54
         Top             =   180
         Width           =   525
      End
      Begin VB.Label lblToStation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "到站(&Z):"
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
         Left            =   120
         TabIndex        =   0
         Top             =   180
         Width           =   840
      End
   End
   Begin RTComctl3.CoolButton cmdbegin 
      Height          =   345
      Left            =   13215
      TabIndex        =   73
      Top             =   2115
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "启用运费计算(&S)>>"
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
      MICON           =   "frmAccept.frx":0B60
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
Attribute VB_Name = "frmAccept"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************
'Last Modify By: 陆勇庆  2005-8-16
'Last Modify In:
'*******************************************************************************

Option Explicit

Const cnMargin = 15
Dim mbCalItemChanged As Boolean '判断用以计算票价的项目是否有更改
Dim sumprice As Double
Dim sumint As Integer
Dim sumcal As Double
Dim ListCount As Integer
Dim ListMove As Integer

Dim DBServer As String
Dim Database As String
Dim Password As String
Dim User As String

Dim oFreeReg As New CFreeReg
'Dim mbFillNums As Boolean '判断站点列表是否已填充过
Dim mbLoad As Boolean '判断是否已登陆对方站点


Private Sub cboAcceptType_Change()
    On Error GoTo ErrHandle
    Dim i As Integer
    Dim nLen As Integer
    Dim tPriceItem() As TLuggagePriceItem
    If CDbl(Trim(flblTotalPrice.Caption)) > 0 And lblMileage.Caption <> "" Then
        '运费运算区计算
        For i = 1 To lvprice.ListItems.Count
            moAcceptSheet.SheetID = g_szAcceptSheetID
            moAcceptSheet.AcceptType = Trim(cboAcceptType.Text)
            moAcceptSheet.DesStationID = ResolveDisplay(cboEndStation.BoundText)
            moAcceptSheet.CalWeight = CDbl(Trim(lvprice.ListItems(i).Text))
            moAcceptSheet.ActWeight = CDbl(Trim(lvprice.ListItems(i).Text))
            moAcceptSheet.Number = 1
            CalPrice
            If moAcceptSheet.TotalPrice > 0 Then
                nLen = ArrayLength(moAcceptSheet.PriceItems)
                If nLen > 0 Then
'                    ReDim tPriceItem(1 To vsPriceItem.Rows)
                    tPriceItem = moAcceptSheet.PriceItems
                    If Trim(tPriceItem(1).PriceName) = Trim(vsPriceItem.TextMatrix(0, 0)) Then
                        lvprice.ListItems(i).SubItems(1) = tPriceItem(1).PriceValue
                    End If
                End If
            End If
        Next i
        
        
        sumprice = 0
        For i = 1 To lvprice.ListItems.Count
            sumprice = sumprice + CDbl(Trim(lvprice.ListItems(i).SubItems(1)))
        Next i
        mbCalItemChanged = True
        txtBagNum_LostFocus
        vsPriceItem.TextMatrix(0, 1) = sumprice
        CalSumPrice
        
        '平均重量计算
        moAcceptSheet.AcceptType = Trim(cboAcceptType.Text)
        mbCalItemChanged = True
        txtBagNum_LostFocus
    End If
    
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub cboAcceptType_Click()
    cboAcceptType_Change
End Sub

Private Sub cboAcceptType_GotFocus()
    lblAcceptType.ForeColor = clActiveColor
End Sub





Private Sub cboAcceptType_LostFocus()
    lblAcceptType.ForeColor = 0
 
End Sub

Private Sub lvBus_Change()
    FillBusInfo
End Sub

Private Sub cboEndStation_KeyPress(KeyAscii As Integer)
    On Error GoTo here
    
    If KeyAscii = vbKeyReturn Then
        lblToStation.ForeColor = 0
        RefreshBus
        RefreshStationMileageInfo
        FillBusInfo
        CalPrice
    End If
Exit Sub
here:
    ShowErrorMsg
    
End Sub

Private Sub cboSellStation_Click()
    RefreshStation
End Sub

Private Sub cboSettleRatio_Change()
    '重新计算应结运费
    CalSettlePrice
End Sub

Private Sub cboSettleRatio_Click()
    '重新计算应结运费
    CalSettlePrice
End Sub

Private Sub Form_Paint()
    If mdiMain.ActiveForm Is frmAccept Then RefreshCurrentSheetID
End Sub

Private Sub lvBus_Click()
    FillBusInfo
End Sub

Private Sub FillBusInfo()
    On Error GoTo ErrHandle
    Dim rsTemp As Recordset
    Dim szBusID As String
    Dim i As Integer
    Dim szaTemp1() As String
    Dim count1 As Integer
    Dim szTempStationName1() As String
    
    If lvBus.SelectedItem Is Nothing Then Exit Sub
    
    szBusID = cboEndStation.BoundText
    If szBusID = "" Then
        cboVehicle.Text = ""
        
        lblBusDate.Caption = ""
        lblStartTime.Caption = ""
        lblInStation.Caption = ""
        Exit Sub
    End If
    Set rsTemp = moLugSvr.GetToStationBusRS(szBusID, CDate(flblSellDate.Caption))
    If rsTemp.RecordCount > 0 Then
        For i = 1 To rsTemp.RecordCount
            If Trim(rsTemp!bus_id) = RTrim(lvBus.SelectedItem.Text) Then
'                cbovehicle.text = FormatDbValue(rsTemp!license_tag_no)
                lblBusDate.Caption = FormatDbValue(rsTemp!bus_date)
                lblStartTime.Caption = Format(rsTemp!bus_start_time, "hh:mm")
                lblCompanyName.Caption = FormatDbValue(rsTemp!transport_company_short_name)
                Exit For
            End If
            rsTemp.MoveNext
        Next i
    End If
    
    '显示途经站
    szaTemp1 = moLugSvr.GetBusStationNames(CDate(flblSellDate.Caption), Trim(lvBus.SelectedItem.Text))
    count1 = ArrayLength(szaTemp1)
    ReDim szTempStationName1(1 To count1)
    szTempStationName1 = szaTemp1
    lblInStation.Caption = ""
    For i = 1 To count1
        lblInStation.Caption = lblInStation.Caption + " " + szTempStationName1(i)
    Next
    
    '填充车辆
    Dim rsVehicle As Recordset
    Set rsVehicle = moLugSvr.GetBusVehicle(lvBus.SelectedItem.Text)
    cboVehicle.Clear
    cboVehicle.AddItem ""
    If rsVehicle.RecordCount > 0 Then
        For i = 1 To rsVehicle.RecordCount
            cboVehicle.AddItem FormatDbValue(rsVehicle!license_tag_no)
            rsVehicle.MoveNext
        Next i
    End If
    '把车牌赋为当前环境车次的车牌
    If cboVehicle.ListCount > 1 And rsTemp.RecordCount > 0 Then
        If Not rsTemp.EOF Then
            If FormatDbValue(rsTemp!bus_id) = Trim(lvBus.SelectedItem.Text) Then
                For i = 0 To cboVehicle.ListCount - 1
                    If Trim(cboVehicle.List(i)) = FormatDbValue(rsTemp!license_tag_no) Then
                        cboVehicle.ListIndex = i
                        Exit For
                    End If
                Next i
            End If
        End If
    End If

    If cboVehicle.Text <> FormatDbValue(rsTemp!license_tag_no) Then
        cboVehicle.AddItem FormatDbValue(rsTemp!license_tag_no)
        cboVehicle.ListIndex = cboVehicle.ListCount - 1
    End If

    
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub lvBus_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    SortListView lvBus, ColumnHeader.Index
End Sub

Private Sub lvBus_GotFocus()
    If lvBus.ListItems.Count = 0 Then
        cboEndStation.SetFocus
    End If
End Sub

Private Sub lvBus_ItemClick(ByVal Item As MSComctlLib.ListItem)
    FillBusInfo
End Sub



Private Sub cboEndStation_GotFocus()
    lblToStation.ForeColor = clActiveColor
    
End Sub


Private Sub cboEndStation_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
    Case vbKeyLeft
        KeyCode = 0
        If Val(txtPreSell.Value) > 0 Then
            txtPreSell.Value = Val(txtPreSell.Value) - 1
        End If
    Case vbKeyRight
        KeyCode = 0
        txtPreSell.Value = Val(txtPreSell.Value) + 1
    Case Else
    End Select

End Sub



Private Sub cboPack_GotFocus()
    lblPack.ForeColor = clActiveColor
End Sub



Private Sub cboPack_LostFocus()
    lblPack.ForeColor = 0
End Sub

Private Sub cboPickType_GotFocus()
    lblPickType.ForeColor = clActiveColor
End Sub



Private Sub cboPickType_KeyPress(KeyAscii As Integer)
'    Select Case KeyAscii
'    Case vbKeyReturn
'        KeyAscii = 0
'        txtPickerAddress.SetFocus
'    Case vbKeyRight
'        KeyAscii = 0
'        txtPickerAddress.SetFocus
'    Case vbKeyLeft
'        KeyAscii = 0
'        txtShippePhone.SetFocus
'    End Select
End Sub

Private Sub cboPickType_LostFocus()
    lblPickType.ForeColor = 0
End Sub


Private Sub AcceptIdentity()
    On Error GoTo ErrHandle
    moAcceptSheet.SheetID = g_szAcceptSheetID
    moAcceptSheet.AcceptType = Trim(cboAcceptType.Text)
    moAcceptSheet.LuggageName = Trim(cboLuggageName.Text)
    moAcceptSheet.DesStationID = ResolveDisplay(cboEndStation.BoundText)
    
    '**************
    If lvBus.SelectedItem Is Nothing Then
        moAcceptSheet.BusID = "" 'Trim(flbCarryBus.Caption)
    Else
        moAcceptSheet.BusID = lvBus.SelectedItem.Text
    End If
    '*****************
    
    moAcceptSheet.Number = CInt(txtBagNum.Text)
    moAcceptSheet.CalWeight = Trim(txtCalWeight.Text)
    moAcceptSheet.ActWeight = Trim(txtActWeight.Text)
    
    
    moAcceptSheet.StartLabelID = Trim(txtStartLabel.Text)
    moAcceptSheet.EndLabelID = Trim(txtEndLabel.Text)
    
    moAcceptSheet.OverNumber = CInt(txtOverNum.Text)
    moAcceptSheet.Shipper = Trim(txtShipper.Text)
    moAcceptSheet.Picker = Trim(txtPicker.Text)
    moAcceptSheet.PickerPhone = Trim(txtPhone.Text)
    moAcceptSheet.LuggageShipperPhone = Trim(txtShippePhone.Text)
    moAcceptSheet.PickerAddress = Trim(txtPickerAddress.Text)
    moAcceptSheet.PickType = Trim(cboPickType.Text)
    
    
    
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub



Private Sub cmdAccept_Click()
    
    If txtCalWeight.Text = 0 Then
        MsgBox "计重不能为 0", vbExclamation, Me.Caption
        Exit Sub
    End If
    If txtBagNum.Text = 0 Then
        MsgBox "件数不能为 0", vbExclamation, Me.Caption
        Exit Sub
    End If
    If Val(txtOverNum.Text) > Val(txtBagNum.Text) Then
        MsgBox "超重件数不能大于件数", vbExclamation, Me.Caption
        Exit Sub
    End If
    If lblMileage.Caption = "" Then
        MsgBox "无里程,请选择到达站!", vbInformation, Me.Caption
        cboEndStation.SetFocus
        Exit Sub
    End If
    If cboVehicle.Text = "" Then
        MsgBox "车次不能为空,请指定车次!", vbInformation, Me.Caption
        Exit Sub
    End If
    If txtShipper.Text = "" Then
        MsgBox "托运人不能为空,请指定托运人!", vbInformation, Me.Caption
        Exit Sub
    End If
    If txtPicker.Text = "" Then
        MsgBox "收件人不能为空,请指定收件人!", vbInformation, Me.Caption
        Exit Sub
    End If
    
    
    On Error GoTo ErrHandle
    g_szAcceptSheetID = Trim(mdiMain.lblSheetNo.Caption)
    moAcceptSheet.SheetID = g_szAcceptSheetID
    moAcceptSheet.AcceptType = Trim(cboAcceptType.Text)
    moAcceptSheet.LuggageName = Trim(cboLuggageName.Text)
    moAcceptSheet.DesStationID = cboEndStation.BoundText
    If lvBus.SelectedItem Is Nothing Then
        moAcceptSheet.BusID = "" 'Trim(flbCarryBus.Caption)
    Else
        moAcceptSheet.BusID = lvBus.SelectedItem.Text
    End If
    If lblBusDate.Caption <> "" Then
        moAcceptSheet.BusDate = CDate(lblBusDate.Caption)
    End If
    moAcceptSheet.CalWeight = Trim(txtCalWeight.Text)
    moAcceptSheet.ActWeight = Trim(txtActWeight.Text)
    moAcceptSheet.StartLabelID = Trim(txtStartLabel.Text)
    moAcceptSheet.EndLabelID = Trim(txtEndLabel.Text)
    moAcceptSheet.Number = CInt(txtBagNum.Text)
    moAcceptSheet.OverNumber = CInt(txtOverNum.Text)
    moAcceptSheet.Shipper = Trim(txtShipper.Text)
    moAcceptSheet.Picker = Trim(txtPicker.Text)
    moAcceptSheet.PickerPhone = Trim(txtPhone.Text)
    moAcceptSheet.PickerAddress = Trim(txtPickerAddress.Text)
    moAcceptSheet.LuggageShipperPhone = Trim(txtShippePhone.Text)
    moAcceptSheet.PickType = Trim(cboPickType.Text)
    moAcceptSheet.LicenseTagNo = Trim(cboVehicle.Text)
    moAcceptSheet.Pack = cboPack.Text
    moAcceptSheet.CalBasePrice = Val(lblBasePrice.Caption)
    
    moAcceptSheet.TransTicketID = Trim(txtTransTicketID.Text)
    moAcceptSheet.InsuranceID = Trim(txtInsuranceID.Text)
    moAcceptSheet.SettleRatio = Val(cboSettleRatio.Text) / 100
    moAcceptSheet.SettlePrice = Val(lblSettlePrice.Caption)
    moAcceptSheet.Annotation1 = Trim(txtAnnotation1.Text)
    moAcceptSheet.Annotation2 = Trim(txtAnnotation2.Text)
    
    
    
    moLugSvr.AcceptLuggage moAcceptSheet
    
    '打印受理单
    ShowSBInfo "正在打印受理单..."
'    PrintAcceptSheet moAcceptSheet, lblBusDate.Caption & " " & lblStartTime.Caption
    frmSheet.PrintSheetReport moAcceptSheet, lblBusDate.Caption & " " & lblStartTime.Caption
    ShowSBInfo ""
    
    '当前行包单号加1
    IncTicketNo
    RefreshClear
    cboEndStation.SetFocus
    g_szAcceptSheetID = Trim(mdiMain.lblSheetNo.Caption)
    ' 等待下一张受理单
    moAcceptSheet.AddNew
    moAcceptSheet.SheetID = g_szAcceptSheetID
    '计算运费初始化
    '/************************
    FraPrice.Visible = False
    fradetail.Visible = False
    lvprice.ListItems.Clear
    sumprice = 0
    sumint = 0
    sumcal = 0
    txtCal.Text = 0
    txtPrice.Text = 0
'    fraOutLine.Width = 9315
    Form_Resize
    '/************************
    Exit Sub
ErrHandle:
    ShowSBInfo ""
    ShowErrorMsg
End Sub

Private Sub cmdBagDetail_Click()
    '需要更改
    frmLugDetail.LuggageID = g_szAcceptSheetID
    frmLugDetail.Show vbModal
End Sub

Private Sub cmdbegin_Click()
    
'    fraOutLine.Width = 11865
    Form_Resize
    FraPrice.Visible = True
    fradetail.Visible = True
    txtCal.Enabled = True
    txtPrice.Enabled = False
    cmdbegin.Enabled = False
    txtCal.SetFocus
End Sub

Private Sub cmdover_Click()
    Dim i As Integer
    
'    fraOutLine.Width = 9315
    Form_Resize
    If sumint = 0 Then
        cmdbegin.Enabled = True
        FraPrice.Visible = False
        txtBagNum.Enabled = True
        cboLuggageName.SetFocus
        fradetail.Visible = False
        txtCalWeight.Text = 0
        txtActWeight.Text = 0
        txtBagNum.Text = 0
        txtOverNum.Text = 0
        Exit Sub
    End If
    sumcal = 0
    sumprice = 0
    For i = 1 To lvprice.ListItems.Count
        sumcal = sumcal + CDbl(Trim(lvprice.ListItems(i).Text))
        sumprice = sumprice + CDbl(Trim(lvprice.ListItems(i).SubItems(1)))
    Next i
    txtCalWeight.Text = sumcal
    txtActWeight.Text = sumcal
    txtBagNum.Text = sumint
    mbCalItemChanged = True
    txtBagNum_LostFocus
    vsPriceItem.TextMatrix(0, 1) = sumprice
    CalSumPrice
    txtBagNum.Enabled = False
    FraPrice.Visible = False
    cmdbegin.Enabled = True
    txtStartLabel.SetFocus

End Sub



Private Sub flblTotalPrice_Change()
    If Val(flblTotalPrice.Caption) > 0 Then
        cmdAccept.Enabled = True
    Else
        cmdAccept.Enabled = False
    End If
End Sub

Private Sub Form_Activate()

    SetSheetNoLabel True, g_szAcceptSheetID
    
End Sub
 
Private Sub Form_Deactivate()
    HideSheetNoLabel
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error GoTo ErrorHandle
    If KeyCode = vbKeyF9 Then '修改票价
        vsPriceItem.SetFocus
        If vsPriceItem.Cols > 0 Then
            vsPriceItem.Select 0, 1
        End If
        
    ElseIf KeyCode = vbKeyF2 Then
        If cmdAccept.Enabled Then
            cboEndStation.SetFocus
            cmdAccept_Click
        End If
    ElseIf KeyCode = vbKeyF3 Then  '托运类型切换
        If cboAcceptType.ListIndex = 0 And cboAcceptType.ListCount > 1 Then
        
            cboAcceptType.ListIndex = 1
        Else
            cboAcceptType.ListIndex = 0
        End If
    ElseIf KeyCode = vbKeyF5 Then
        '选择车辆
        Dim oCommDialog As New CommDialog
        Dim aszTemp() As String
        Dim i As Integer
        Dim bFound As Boolean
        If Not (lvBus.SelectedItem Is Nothing) Then
            oCommDialog.Init m_oAUser
            aszTemp = oCommDialog.SelectVehicleEX()
            If ArrayLength(aszTemp) > 0 Then
                bFound = False
                For i = 0 To cboVehicle.ListCount - 1
                    If cboVehicle.List(i) = aszTemp(1, 2) Then
                        Exit For
                        bFound = True
                    End If
                Next i
                If Not bFound Then
                    cboVehicle.AddItem aszTemp(1, 2)
                    cboVehicle.ListIndex = cboVehicle.ListCount - 1
                Else
                    cboVehicle.ListIndex = i
                End If
            End If
        End If
    End If
    
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrHandle
    If KeyAscii = vbKeyTab Then Exit Sub
'    If KeyAscii = vbKeyF1 Then
'        DisplayHelp Me
'    End If
    If KeyAscii = vbKeyReturn And (Me.ActiveControl Is cboAcceptType) Then
        txtCalWeight.SetFocus
        KeyAscii = 0
    ElseIf KeyAscii = vbKeyReturn And (Me.ActiveControl Is txtOverNum) Then
        cboPack.SetFocus
        KeyAscii = 0
    ElseIf KeyAscii = vbKeyReturn And (Me.ActiveControl Is txtPhone) Then
        vsPriceItem.SetFocus
        KeyAscii = 0
    ElseIf KeyAscii = vbKeyReturn Then 'And Not (Me.ActiveControl Is txtAnnotation1) And Not (Me.ActiveControl Is txtAnnotation2) Then
        SendTab
    ElseIf KeyAscii = 27 Then
        lblMileage.Caption = ""
        cboEndStation.SetFocus
        RefreshClear
    ElseIf KeyAscii = 43 Or KeyAscii = 61 Then
        If txtPreSell.Enabled = True Then
            txtPreSell.Value = Val(txtPreSell.Value) + 1
        End If
    ElseIf KeyAscii = 45 Then
        If txtPreSell.Enabled = True And Val(txtPreSell.Value) > 0 Then
            txtPreSell.Value = Val(txtPreSell.Value) - 1
        End If
    End If
    
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub Form_Load()
    Dim tPriceItem() As TLuggagePriceItem
    Dim i As Integer
    Dim rsTemp As Recordset
    On Error GoTo here
    
    ReadReg '读注册表内容
    InitLvBus
    AlignFormPos Me
    AlignHeadWidth Me.name, lvBus
'    fraOutLine.Width = 9315
    Form_Resize
    '计算运费初始化
    FraPrice.Visible = False
    fradetail.Visible = False
    lvprice.ListItems.Clear
    sumprice = 0
    sumint = 0
    sumcal = 0
    txtCal.Text = 0
    txtPrice.Text = 0
    moAcceptSheet.Init m_oAUser
    InitGrid
    RefreshClear
    cmdAccept.Enabled = False
    moAcceptSheet.AddNew
    g_szAcceptSheetID = GetTicketNo
    moAcceptSheet.SheetID = g_szAcceptSheetID
    SetSheetNoLabel True, g_szAcceptSheetID
    RefreshStation
    FillLuggageName
    flblSellDate.Caption = ToStandardDateStr(m_oParam.NowDate)
'    If chkSelectBus.Value <> 1 Then
       txtPreSell.Enabled = True
'    End If
    '填充受理类型
    With cboAcceptType
       .AddItem szAcceptTypeGeneral
       .AddItem szAcceptTypeMan
       .Text = szAcceptTypeGeneral
    End With
    '填充交付方式
    With cboPickType
       .AddItem szPickTypeGeneral
       .AddItem szPickTypeEms
       .Text = szPickTypeGeneral
    End With
    '填充包装
    With cboPack
        .AddItem "纸箱"
        .AddItem "木箱"
        .AddItem "编织袋"
        .AddItem "麻袋"
        .AddItem "筐"
        .AddItem "桶"
        .AddItem "布袋"
        
        .AddItem "袋"
        .AddItem "包"
        .AddItem "箱"
        
    End With
    With cboSettleRatio
        .AddItem "100%"
        .AddItem "80%"
        .AddItem "70%"
        .AddItem "65%"
        .AddItem "60%"
        .AddItem "50%"
        .AddItem "40%"
        .AddItem "30%"
        .AddItem "20%"
        
    End With
    cboSettleRatio.ListIndex = 0
    
    Station '联网受理站点
    
Exit Sub
here:
ShowErrorMsg
End Sub

Private Sub FillLuggageName()
    Dim rsTemp As Recordset
    Dim i As Integer
    On Error GoTo ErrorHandle
    
    Set rsTemp = moLugSvr.GetLuggageKinds
    For i = 1 To rsTemp.RecordCount
        cboLuggageName.AddItem FormatDbValue(rsTemp!kinds_name)
        rsTemp.MoveNext
    Next i
'    If rsTemp.RecordCount > 0 Then cboLuggageName.ListIndex = 0
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub InitGrid()
    On Error GoTo ErrorHandle
    Dim atPriceItem() As TLuggagePriceItem
    Dim oTmp As New LuggageParam
    oTmp.Init m_oAUser
    atPriceItem = oTmp.GetPriceItemInfo(GetLuggageTypeInt(Trim(cboAcceptType.Text)))
    vsPriceItem.Cols = 2
    vsPriceItem.FixedCols = 1
    vsPriceItem.Rows = ArrayLength(atPriceItem)
'    vsPriceItem.FixedRows = 1
    
    vsPriceItem.ColWidth(0) = 1000 'vsPriceItem.Width * 0.6
    vsPriceItem.ColWidth(1) = 1000 ' vsPriceItem.Width * 0.4
'    vsPriceItem.ColWidth(2) = 1000
    
    Dim i As Integer
    For i = 0 To ArrayLength(atPriceItem) - 1
        vsPriceItem.TextMatrix(i, 0) = atPriceItem(i + 1).PriceName
        vsPriceItem.TextMatrix(i, 1) = atPriceItem(i + 1).PriceValue
    Next i
    vsPriceItem.Row = 0
    vsPriceItem.Col = 1
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub
'清空界面
Private Sub RefreshClear()
  Dim i As Integer
    lblInStation.Caption = ""
'    cboEndStation.Clear
    lvBus.ListItems.Clear
    lblPreSellDate.ForeColor = 0

    cboLuggageName.Text = ""
    cboVehicle.Clear
    cboVehicle.AddItem ""
    cboVehicle.ListIndex = 0
    txtActWeight.Text = 0
    txtCalWeight.Text = 0
    txtOverNum.Text = 0
    txtBagNum.Text = 0
    txtStartLabel.Text = ""
    txtEndLabel.Text = ""
    cboPack.Text = ""
    txtTransTicketID.Text = ""
    txtInsuranceID.Text = ""
    txtAnnotation1.Text = ""
    txtAnnotation2.Text = ""

    txtPicker.Text = ""
    txtPhone.Text = ""
    txtPickerAddress.Text = ""
    txtShipper.Text = ""
    txtShippePhone.Text = ""
'    chkSelectBus.Value = 0
'    chkSelectBus.Enabled = true
    txtPreSell.Enabled = True
    If cboAcceptType.ListCount > 0 Then cboAcceptType.ListIndex = 0
'    cboAcceptType.Clear
'    cboPickType.Clear
    lblMileage.Caption = ""
    txtPreSell.Value = 0
    lblBusDate.Caption = ""
    lblStartTime.Caption = ""
    If cboSettleRatio.ListCount > 0 Then cboSettleRatio.ListIndex = 0
    chkPrintSettlePrice.Value = vbChecked
    'vsPriceItem.Clear
    For i = 1 To vsPriceItem.Rows
        vsPriceItem.TextMatrix(i - 1, 1) = 0
    Next i
    flblTotalPrice.Caption = "0.00"
    txtReceivedMoney.Text = "0.00"
    flblRestMoney.Caption = "0.00"
    txtBagNum.Enabled = True
    fradetail.Visible = False
    lvprice.ListItems.Clear
    




End Sub
Private Sub RefreshStation()
On Error GoTo here
    Dim rsTemp As Recordset
    Dim szTemp As String
    Set rsTemp = moLugSvr.GetAllStationRS()
    With cboEndStation
        Set .RowSource = rsTemp
        'station_id:到站代码
        'station_input_code:车站输入码
        'station_name:车次名称
        
        .BoundField = "station_id"
        .ListFields = "station_input_code:4,station_name:4"
        .AppendWithFields "station_id:9,station_name"
    End With
    
    Set rsTemp = Nothing
    On Error GoTo 0
    Exit Sub
here:
ShowErrorMsg
End Sub

Private Sub RefreshStationMileageInfo()
On Error GoTo here
    Dim szStationID As String
    szStationID = Trim(ResolveDisplay(cboEndStation.BoundText))
    moAcceptSheet.DesStationID = szStationID
    lblMileage.Caption = CStr(moAcceptSheet.Mileage) + " 公里"
    
    On Error GoTo 0
    Exit Sub
here:
        MsgBox err.Description, vbExclamation, "错误-RefreshStationMileageInfo " & err.Number
End Sub

'刷新车次信息
Private Sub RefreshBus()
    On Error GoTo here
    Dim rsTemp As Recordset
    Dim szBusID As String
    Dim liTemp As ListItem
    Dim i As Integer
    cboVehicle.Clear
    lblBusDate.Caption = ""
    lblStartTime.Caption = ""
    
    szBusID = ResolveDisplay(cboEndStation.BoundText)
    If cboEndStation.BoundText = "" Then Exit Sub
    lvBus.ListItems.Clear
    Set rsTemp = moLugSvr.GetToStationBusRS(szBusID, CDate(flblSellDate.Caption))
    If rsTemp.RecordCount > 0 Then

        With lvBus
            rsTemp.MoveFirst
            For i = 1 To rsTemp.RecordCount
                Set liTemp = .ListItems.Add(, , FormatDbValue(rsTemp!bus_id))
                liTemp.SubItems(1) = IIf(FormatDbValue(rsTemp!bus_type) = TP_ScrollBus, "滚动车次", Format(FormatDbValue(rsTemp!bus_start_time), "hh:mm"))
                
                liTemp.SubItems(2) = FormatDbValue(rsTemp!end_station_name)
                liTemp.SubItems(3) = FormatDbValue(rsTemp!vehicle_type_name)
                liTemp.SubItems(4) = FormatDbValue(rsTemp!vehicle_type_code)
                If FormatDbValue(rsTemp!transport_company_id) = g_szOurCompany Then
                    SetListViewLineColor lvBus, i, vbBlue
                End If
                rsTemp.MoveNext
            Next i
            .ListItems(1).Selected = True
        End With
        
    End If
    Set rsTemp = Nothing
    On Error GoTo 0
    Exit Sub
here:
        MsgBox err.Description, vbExclamation, "错误-RefreshBus " & err.Number
End Sub

Private Sub Form_LostFocus()
    lblMileage.Caption = ""
    RefreshClear
    cboEndStation.Text = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveFormPos Me  '保存本次窗口位置
    SaveHeadWidth Me.name, lvBus
    HideSheetNoLabel
    SaveOldReg
End Sub
Private Sub Form_Resize()
    If mdiMain.ActiveForm Is Me Then
        If Not Me.WindowState = vbMaximized Then Me.WindowState = vbMaximized
        fraOutLine.Left = (Me.ScaleWidth - fraOutLine.Width) / 2
        fraOutLine.Top = (Me.ScaleHeight - fraOutLine.Height) / 2
    End If
End Sub


Private Sub lvprice_DblClick()
    txtPrice.Enabled = True
    txtCal.Enabled = False
    txtPrice.SetFocus
    txtCal.Text = Trim(lvprice.SelectedItem.Text)
    txtPrice.Text = Trim(lvprice.SelectedItem.SubItems(1))
    ListCount = lvprice.SelectedItem.Index
End Sub

Private Sub lvprice_ItemClick(ByVal Item As MSComctlLib.ListItem)
   ListMove = Item.Index
End Sub



Private Sub lvprice_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyDelete And ListMove <> 0 Then
        lvprice.ListItems.Remove ListMove
        sumint = sumint - 1
        ListMove = 0
    End If
End Sub

Private Sub txtActWeight_Change()
    On Error GoTo ErrHandle
    FormatTextToNumeric txtActWeight, False, False
    AcceptIdentity
    CalPrice
    If moAcceptSheet.TotalPrice > 0 Then
        cmdAccept.Enabled = True
        '      RefreshPrice
        flblTotalPrice.Caption = CStr(moAcceptSheet.TotalPrice)
    End If
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub txtActWeight_GotFocus()
    lblActWeight.ForeColor = clActiveColor
    
    txtActWeight.SelStart = 0
    txtActWeight.SelLength = 100
End Sub

Private Sub txtActWeight_KeyDown(KeyCode As Integer, Shift As Integer)
'    Select Case KeyCode
'    Case vbKeyReturn
'        KeyCode = 0
'        If txtBagNum.Enabled Then
'            txtBagNum.SetFocus
'        Else
'            txtOverNum.SetFocus
'        End If
'    Case vbKeyRight
'        KeyCode = 0
'        txtBagNum.SetFocus
'    Case vbKeyDown
'        KeyCode = 0
'        txtOverNum.SetFocus
'    Case vbKeyLeft
'        KeyCode = 0
'        txtCalWeight.SetFocus
'    Case vbKeyUp
'        KeyCode = 0
'        txtCalWeight.SetFocus
'    End Select
End Sub



Private Sub txtActWeight_LostFocus()
    lblActWeight.ForeColor = 0
    txtCalWeight.Text = txtActWeight.Text
End Sub

Private Sub txtActWeight_Validate(Cancel As Boolean)
    If Not IsNumeric(txtActWeight.Text) Then
        '如果不为数字
        MsgBox "实重必须输入数字", vbExclamation, Me.Caption
        Cancel = True
    Else
        If txtActWeight.Text < 0 Then
            MsgBox "实重必须大于0", vbExclamation, Me.Caption
            Cancel = True
        End If
    End If
End Sub

Private Sub txtAnnotation1_GotFocus()
    lblAnnotation1.ForeColor = clActiveColor
End Sub

Private Sub txtAnnotation1_LostFocus()
    lblAnnotation1.ForeColor = 0
End Sub

Private Sub txtAnnotation2_GotFocus()
    lblAnnotation2.ForeColor = clActiveColor
End Sub

Private Sub txtAnnotation2_LostFocus()
    lblAnnotation2.ForeColor = 0
End Sub

Private Sub txtBagNum_Change()
    mbCalItemChanged = True
    FormatTextToNumeric txtBagNum, False, False
End Sub

Private Sub txtBagNum_GotFocus()
    mbCalItemChanged = False
    lblBagNum.ForeColor = clActiveColor
    txtBagNum.SelStart = 0
    txtBagNum.SelLength = 100
End Sub

Private Sub txtBagNum_KeyDown(KeyCode As Integer, Shift As Integer)
'    Select Case KeyCode
'    Case vbKeyReturn
'        KeyCode = 0
'        txtOverNum.SetFocus
'    Case vbKeyRight
'        KeyCode = 0
'        txtOverNum.SetFocus
'    Case vbKeyDown
'        KeyCode = 0
'        txtStartLabel.SetFocus
'    Case vbKeyLeft
'        KeyCode = 0
'        txtActWeight.SetFocus
'    Case vbKeyUp
'        KeyCode = 0
'        txtCalWeight.SetFocus
'    End Select
End Sub




Private Sub txtBagNum_LostFocus()

    lblBagNum.ForeColor = 0
    txtCalWeight = Val(txtCalWeight) * Val(txtBagNum) '谢晓冬07-07-06改
    '玉环计重直接输平均重量，让我们自动算出总重
    On Error GoTo ErrHandle
    AcceptIdentity
    CalPrice
    
    If moAcceptSheet.TotalPrice > 0 Then
        cmdAccept.Enabled = True
'        RefreshPrice
        flblTotalPrice.Caption = CStr(moAcceptSheet.TotalPrice)
    End If
'    RefreshPrice
    flblTotalPrice.Caption = CStr(moAcceptSheet.TotalPrice)
    Exit Sub
ErrHandle:
ShowErrorMsg

End Sub


Private Sub txtBagNum_Validate(Cancel As Boolean)
    If Not IsNumeric(txtBagNum.Text) Then
        '如果不为数字
        MsgBox "件数必须输入数字", vbExclamation, Me.Caption
        Cancel = True
    Else
        If txtBagNum.Text < 0 Then
            MsgBox "件数必须大于0", vbExclamation, Me.Caption
            Cancel = True
        End If
    End If
End Sub

Private Sub txtCal_GotFocus()
    txtCal.Text = ""
End Sub

Private Sub txtCal_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Dim nLen As Integer
    Dim tPriceItem() As TLuggagePriceItem
    Dim lvitem As ListItem
    On Error GoTo ErrHandle
    If KeyCode = 13 Then
        If txtCal.Text = "" Or Val(txtCal.Text) = 0 Then
            Exit Sub
        End If
    
        If txtCal.Text <> "" And Val(txtCal.Text) <> 0 Then
            moAcceptSheet.SheetID = g_szAcceptSheetID
            moAcceptSheet.AcceptType = Trim(cboAcceptType.Text)
            moAcceptSheet.DesStationID = ResolveDisplay(cboEndStation.BoundText)
            moAcceptSheet.CalWeight = CDbl(Trim(txtCal.Text))
            moAcceptSheet.ActWeight = CDbl(Trim(txtCal.Text))
            moAcceptSheet.Number = 1
            '  moAcceptSheet.OverNumber = CInt(txtOverNum.Text)
            '    If lblMileage.Caption = "" Then
            '        MsgBox "无里程,请选择到达站!", vbInformation, Me.Caption
            '        cboEndStation.SetFocus
            '        Exit Sub
            '    End If
            CalPrice
            If moAcceptSheet.TotalPrice > 0 Then
            
                nLen = ArrayLength(moAcceptSheet.PriceItems)
                If nLen > 0 Then
'                    ReDim tPriceItem(1 To vsPriceItem.Rows)
                    tPriceItem = moAcceptSheet.PriceItems
                    If Trim(tPriceItem(1).PriceName) = Trim(vsPriceItem.TextMatrix(0, 0)) Then
                        txtPrice.Text = tPriceItem(1).PriceValue
                        '        sumprice = sumprice + CDbl(txtPrice.Text)
                        '        sumcal = sumcal + CDbl(txtCal.Text)
                        sumint = sumint + 1
                    End If
                End If
                
                Set lvitem = lvprice.ListItems.Add(, , "  " & Trim(txtCal.Text))
                lvitem.SubItems(1) = Trim(txtPrice.Text)
            
            End If
        End If
        txtCal.Text = ""
        txtPrice.Text = "0"
        txtCal.SetFocus
    End If
    
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub txtCal_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = vbKeyBack Or KeyAscii = 46) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtCalWeight_Change()
    mbCalItemChanged = True
    FormatTextToNumeric txtCalWeight, False, False
End Sub
'Private Sub RefreshPrice()
'    On Error GoTo ErrHandle
'    Dim i As Integer
'    Dim nlen As Integer
'    Dim tPriceItem() As TLuggagePriceItem
'
'    nlen = ArrayLength(moAcceptSheet.PriceItems)
'    If nlen > 0 Then
'        ReDim tPriceItem(1 To nlen)
'        tPriceItem = moAcceptSheet.PriceItems
'        For i = 1 To nlen
'            If Trim(tPriceItem(i).PriceName) = Trim(vsPriceItem.TextMatrix(i - 1, 0)) Then
'                vsPriceItem.TextMatrix(i - 1, 1) = tPriceItem(i).PriceValue
'            End If
'        Next i
'    End If
'
'    Exit Sub
'ErrHandle:
'    ShowErrorMsg
'End Sub
Private Sub txtCalWeight_GotFocus()
    mbCalItemChanged = False
    lblCalWeight.ForeColor = clActiveColor
    txtCalWeight.SelStart = 0
    txtCalWeight.SelLength = 100
End Sub

Private Sub txtCalWeight_KeyDown(KeyCode As Integer, Shift As Integer)
'    Select Case KeyCode
'    Case vbKeyReturn
'        KeyCode = 0
'        txtActWeight.SetFocus
'    Case vbKeyRight
'        KeyCode = 0
'        txtActWeight.SetFocus
'    Case vbKeyDown
'        KeyCode = 0
'        txtBagNum.SetFocus
'    Case vbKeyLeft
'        KeyCode = 0
'        cboAcceptType.SetFocus
'    Case vbKeyUp
'        KeyCode = 0
'        cboLuggageName.SetFocus
'    Case vbKeyTab
'        KeyCode = 0
'    End Select
End Sub



Private Sub txtCalWeight_LostFocus()
    
    lblCalWeight.ForeColor = 0
    On Error GoTo ErrHandle
    AcceptIdentity
    CalPrice
    If moAcceptSheet.TotalPrice > 0 Then
        cmdAccept.Enabled = True
    End If
    flblTotalPrice.Caption = CStr(moAcceptSheet.TotalPrice)
    Exit Sub
ErrHandle:
    ShowErrorMsg
    
End Sub

Private Sub txtCalWeight_Validate(Cancel As Boolean)
    If Not IsNumeric(txtCalWeight.Text) Then
        '如果不为数字
        MsgBox "计重必须输入数字", vbExclamation, Me.Caption
        Cancel = True
    Else
        If txtCalWeight.Text < 0 Then
            MsgBox "计重必须大于0", vbExclamation, Me.Caption
            Cancel = True
        End If
    End If
End Sub

Private Sub txtEndLabel_GotFocus()
    lblEndLabel.ForeColor = clActiveColor
    txtEndLabel.SelStart = 0
    txtEndLabel.SelLength = 100
End Sub

Private Sub txtEndLabel_KeyDown(KeyCode As Integer, Shift As Integer)
'Select Case KeyCode
'        Case vbKeyReturn
'            KeyCode = 0
'            txtPicker.SetFocus
'        Case vbKeyRight
'            KeyCode = 0
'            txtPicker.SetFocus
'        Case vbKeyDown
'            KeyCode = 0
'            txtShipper.SetFocus
'        Case vbKeyLeft
'            KeyCode = 0
'            txtOverNum.SetFocus
'        Case vbKeyUp
'            KeyCode = 0
'            txtOverNum.SetFocus
'    End Select
End Sub

Private Sub txtEndLabel_LostFocus()
    lblEndLabel.ForeColor = 0
End Sub

Private Sub cboLuggageName_GotFocus()
    lblLuggageName.ForeColor = clActiveColor
End Sub





Private Sub cboLuggageName_LostFocus()
    lblLuggageName.ForeColor = 0
End Sub

Private Sub txtInsuranceID_GotFocus()
    lblInsuranceID.ForeColor = clActiveColor
End Sub

Private Sub txtInsuranceID_LostFocus()
    lblInsuranceID.ForeColor = 0
End Sub

Private Sub txtOverNum_Change()
    mbCalItemChanged = True
    FormatTextToNumeric txtOverNum, False, False
End Sub

Private Sub txtOverNum_GotFocus()
    mbCalItemChanged = False
    lblOverNum.ForeColor = clActiveColor
    txtOverNum.SelStart = 0
    txtOverNum.SelLength = 100
End Sub

Private Sub txtOverNum_KeyDown(KeyCode As Integer, Shift As Integer)
''    On Error Resume Next
'
'    Select Case KeyCode
'    Case vbKeyReturn
'        KeyCode = 0
'        txtStartLabel.SetFocus
'
'    Case vbKeyRight
'        KeyCode = 0
'    Case vbKeyDown
'        KeyCode = 0
'        txtEndLabel.SetFocus
'    Case vbKeyLeft
'        KeyCode = 0
'        txtBagNum.SetFocus
'    Case vbKeyUp
'        KeyCode = 0
'        txtActWeight.SetFocus
'    End Select
End Sub


Private Sub txtOverNum_LostFocus()
    lblOverNum.ForeColor = 0
    
    If mbCalItemChanged = False Then Exit Sub
    
    On Error GoTo ErrHandle
    AcceptIdentity
    '    If lblMileage.Caption = "" Then
    '      MsgBox "无里程,请选择到达站!", vbInformation, Me.Caption
    '      cboEndStation.SetFocus
    '      Exit Sub
    '    End If
    CalPrice
    If moAcceptSheet.TotalPrice > 0 Then
        cmdAccept.Enabled = True
    End If
    '    RefreshPrice
    flblTotalPrice.Caption = CStr(moAcceptSheet.TotalPrice)
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub txtOverNum_Validate(Cancel As Boolean)
    If Not IsNumeric(txtOverNum.Text) Then
        '如果不为数字
        MsgBox "超重件数必须输入数字", vbExclamation, Me.Caption
        Cancel = True
    Else
        If txtOverNum.Text < 0 Then
            MsgBox "超重件数必须大于0", vbExclamation, Me.Caption
            Cancel = True
        End If
    End If
End Sub

Private Sub txtPhone_GotFocus()
    lblPhone.ForeColor = clActiveColor
End Sub



Private Sub txtPhone_KeyDown(KeyCode As Integer, Shift As Integer)
'    Select Case KeyCode
'    Case vbKeyReturn
'        KeyCode = 0
'        txtShipper.SetFocus
'    Case vbKeyRight
'        KeyCode = 0
'        txtShipper.SetFocus
'    Case vbKeyDown
'        KeyCode = 0
'        txtShippePhone.SetFocus
'    Case vbKeyLeft
'        KeyCode = 0
'        txtPicker.SetFocus
'    Case vbKeyUp
'        KeyCode = 0
'        txtEndLabel.SetFocus
'    End Select
End Sub

Private Sub txtPhone_LostFocus()
    lblPhone.ForeColor = 0
End Sub

Private Sub txtPicker_GotFocus()
    lblPicker.ForeColor = clActiveColor
End Sub



Private Sub txtPicker_KeyDown(KeyCode As Integer, Shift As Integer)
'    Select Case KeyCode
'    Case vbKeyReturn
'        KeyCode = 0
'        txtPhone.SetFocus
'    Case vbKeyRight
'        KeyCode = 0
'        txtPhone.SetFocus
'    Case vbKeyDown
'        KeyCode = 0
'        txtShipper.SetFocus
'    Case vbKeyLeft
'        KeyCode = 0
'        txtEndLabel.SetFocus
'    Case vbKeyUp
'        KeyCode = 0
'        txtStartLabel.SetFocus
'    End Select
End Sub

Private Sub txtPicker_LostFocus()
    lblPicker.ForeColor = 0
End Sub

Private Sub txtPickerAddress_GotFocus()
    lblPickerAddress.ForeColor = clActiveColor
End Sub

Private Sub txtPickerAddress_KeyDown(KeyCode As Integer, Shift As Integer)
'    Select Case KeyCode
'    Case vbKeyReturn
'        KeyCode = 0
'        vsPriceItem.SetFocus
'    Case vbKeyRight
'        KeyCode = 0
'        cboPickType.SetFocus
'    Case vbKeyDown
'        KeyCode = 0
'        vsPriceItem.SetFocus
'    Case vbKeyLeft
'        KeyCode = 0
'        txtShippePhone.SetFocus
'    Case vbKeyUp
'        KeyCode = 0
'        txtShipper.SetFocus
'    End Select
End Sub

Private Sub txtPickerAddress_LostFocus()
    lblPickerAddress.ForeColor = 0
End Sub

Private Sub txtPreSell_Change()
    On Error Resume Next
    Dim m_nCanSellDay  As Integer
    m_nCanSellDay = m_oParam.PreSellDate
    If Val(txtPreSell.Value) > m_nCanSellDay Then txtPreSell.Value = m_nCanSellDay
    flblSellDate.Caption = ToStandardDateStr(DateAdd("d", txtPreSell.Value, m_oParam.NowDate))
End Sub

Private Sub txtPreSell_GotFocus()
    lblPreSellDate.ForeColor = clActiveColor
End Sub



Private Sub txtPreSell_LostFocus()
    lblPreSellDate.ForeColor = 0
    RefreshBus

    RefreshStationMileageInfo
    FillBusInfo
    CalPrice
End Sub


Private Sub txtPrice_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
       lvprice.ListItems(ListCount).SubItems(1) = Trim(txtPrice.Text)
       txtPrice.Enabled = False
       txtCal.Enabled = True
       txtCal.Text = ""
    End If
End Sub

Private Sub txtPrice_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = vbKeyBack Or KeyAscii = 46) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtReceivedMoney_Change()
    If txtReceivedMoney.Text <> "" Then
        flblRestMoney.Caption = Str(Round(Val(txtReceivedMoney.Text) - Val(flblTotalPrice.Caption), 2))
    End If
End Sub

Private Sub txtReceivedMoney_GotFocus()
    lblReceivedMoney.ForeColor = clActiveColor
    txtReceivedMoney.Text = ""
End Sub





Private Sub txtReceivedMoney_LostFocus()
    lblReceivedMoney.ForeColor = 0
End Sub

Private Sub txtShippePhone_GotFocus()
    lblShippePhone.ForeColor = clActiveColor
End Sub

Private Sub txtShippePhone_KeyDown(KeyCode As Integer, Shift As Integer)
'    Select Case KeyCode
'    Case vbKeyReturn
'        KeyCode = 0
'        cboPickType.SetFocus
'    Case vbKeyRight
'        KeyCode = 0
'        cboPickType.SetFocus
'    Case vbKeyDown
'        KeyCode = 0
'        cboPickType.SetFocus
'    Case vbKeyLeft
'        KeyCode = 0
'        txtShipper.SetFocus
'    Case vbKeyUp
'        KeyCode = 0
'        txtPhone.SetFocus
'    End Select
End Sub

Private Sub txtShippePhone_LostFocus()
    lblShippePhone.ForeColor = 0
End Sub

Private Sub txtShipper_GotFocus()
    lblShipper.ForeColor = clActiveColor
End Sub

Private Sub txtShipper_KeyDown(KeyCode As Integer, Shift As Integer)
'    Select Case KeyCode
'    Case vbKeyReturn
'        KeyCode = 0
'        txtShippePhone.SetFocus
'    Case vbKeyRight
'        KeyCode = 0
'        txtShippePhone.SetFocus
'    Case vbKeyDown
'        KeyCode = 0
'        txtPickerAddress.SetFocus
'    Case vbKeyLeft
'        KeyCode = 0
'        txtPhone.SetFocus
'    Case vbKeyUp
'        KeyCode = 0
'        txtPicker.SetFocus
'    End Select
End Sub

Private Sub txtShipper_LostFocus()
    lblShipper.ForeColor = 0
End Sub

Private Sub txtStartLabel_GotFocus()
    lblStartLabel.ForeColor = clActiveColor
End Sub

Private Sub txtStartLabel_KeyDown(KeyCode As Integer, Shift As Integer)
'    Select Case KeyCode
'    Case vbKeyReturn
'        KeyCode = 0
'        txtPicker.SetFocus
'    Case vbKeyRight
'        KeyCode = 0
'        txtPicker.SetFocus
'    Case vbKeyDown
'        KeyCode = 0
'        txtPicker.SetFocus
'    Case vbKeyLeft
'        KeyCode = 0
'        txtOverNum.SetFocus
'    Case vbKeyUp
'        KeyCode = 0
'        txtBagNum.SetFocus
'    End Select
End Sub

Private Sub txtStartLabel_LostFocus()
    Dim mszLabel As String
    lblStartLabel.ForeColor = 0
    If txtStartLabel.Text <> "" Then
        mszLabel = Val(txtStartLabel.Text) + txtBagNum.Text - 1
        txtEndLabel.Text = mszLabel
    End If
End Sub

Private Sub txtTransTicketID_GotFocus()
    lblTransTicketID.ForeColor = clActiveColor
End Sub

Private Sub txtTransTicketID_LostFocus()
    lblTransTicketID.ForeColor = 0
End Sub

Private Sub vsPriceItem_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    On Error GoTo ErrHandle:
    'CalSumPrice
    Dim i As Integer
    Dim tPriceItem() As TLuggagePriceItem
    Dim nLen As Integer
    '改变票价后同时改变属性PriceItem
    nLen = ArrayLength(moAcceptSheet.PriceItems)
    If nLen > 0 Then
'        ReDim tPriceItem(1 To vsPriceItem.Rows)
        tPriceItem = moAcceptSheet.PriceItems
        For i = 1 To nLen
            If Trim(vsPriceItem.TextMatrix(i - 1, 0)) = Trim(tPriceItem(i).PriceName) Then
                If Not IsNumeric(vsPriceItem.TextMatrix(i - 1, 1)) Then
                    vsPriceItem.TextMatrix(i - 1, 1) = 0
                End If
                tPriceItem(i).PriceValue = CDbl(vsPriceItem.TextMatrix(i - 1, 1)) '写票价
                
            End If
        Next i
    End If
    moAcceptSheet.PriceItems = tPriceItem
    CalSettlePrice
    CalSumPrice
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub vsPriceItem_GotFocus()
    lblPriceItem.ForeColor = clActiveColor
    vsPriceItem_SelChange
End Sub





Private Sub vsPriceItem_LostFocus()
    lblPriceItem.ForeColor = 0
End Sub

Private Sub vsPriceItem_SelChange()
    With vsPriceItem
'        If (.Row = 0 Or .Row = 1 Or .Row = 2 Or .Row = 3 Or .Row = 4) And .Col = 1 Then
        If .Col = 1 Then
            '如果为运费项
            .Editable = flexEDKbdMouse
        Else
            .Editable = flexEDNone
        End If
    End With
End Sub



Private Sub CalSumPrice()
    On Error GoTo ErrHandle
    '算出总价 vsPriceItem.subTotal
    Dim i As Integer
    Dim Num As Integer
    Dim Sum As Double
    Num = vsPriceItem.Rows
    For i = 1 To Num
        Sum = Sum + Val(vsPriceItem.TextMatrix(i - 1, 1))
    Next i
    flblTotalPrice.Caption = Sum
    
    
    Exit Sub
ErrHandle:
        MsgBox err.Description, vbExclamation, "错误-CalSumPrice " & err.Number

End Sub


Private Sub CalPrice()
    Dim i As Integer
    Dim nLen As Integer
    Dim tPriceItem() As TLuggagePriceItem
    Dim szVehicleType As String
    
'    On Error GoTo ErrorHandle
    If cboEndStation.BoundText <> "" Then
        
        If m_bIsRelationWithVehicleType Then
            '如果与车型有关联
            szVehicleType = ""
            If Not (lvBus.SelectedItem Is Nothing) Then
                szVehicleType = lvBus.SelectedItem.SubItems(4)
            End If
            moAcceptSheet.CalculatePrice GetLuggageTypeInt(Trim(cboAcceptType.Text)), cboEndStation.BoundText, szVehicleType
            
        Else
            
            moAcceptSheet.CalculatePrice GetLuggageTypeInt(Trim(cboAcceptType.Text)), cboEndStation.BoundText, ""
            
        End If
        
        nLen = ArrayLength(moAcceptSheet.PriceItems)
        If nLen > 0 Then
'            ReDim tPriceItem(1 To vsPriceItem.Rows)
            tPriceItem = moAcceptSheet.PriceItems
            For i = 1 To nLen
                If Trim(tPriceItem(i).PriceName) = Trim(vsPriceItem.TextMatrix(i - 1, 0)) Then
                    vsPriceItem.TextMatrix(i - 1, 1) = tPriceItem(i).PriceValue
                    
                    If i = 1 Then
                        lblBasePrice.Caption = tPriceItem(i).PriceValue
                    End If
                End If
            Next i
        End If
        '计算应结运费
        CalSettlePrice
    End If
    Exit Sub
'ErrorHandle:
'        MsgBox err.Description, vbExclamation, "错误-CalPrice " & err.Number
End Sub

Private Sub CalSettlePrice()
    '计算应结运费
    If vsPriceItem.Rows > 0 Then
        lblSettlePrice.Caption = Format(vsPriceItem.TextMatrix(0, 1) * Val(cboSettleRatio.Text) / 100, "0.0")
    End If
End Sub


Private Sub InitLvBus()
    '初始化车次列表的列头
    With lvBus
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "车次", 600
        .ColumnHeaders.Add , , "时间", 700
        .ColumnHeaders.Add , , "终点站", 800
        .ColumnHeaders.Add , , "车型", 800
        .ColumnHeaders.Add , , "车型代码", 0
'        .ColumnHeaders.Add , , "公司", 0
    End With
End Sub

Private Sub chkLinkAccept_Click()
    If chkLinkAccept.Value = 1 Then
        cboLinkAccept.Enabled = True
        cboSellStation.Enabled = True
    Else
        cboLinkAccept.Enabled = False
        cboSellStation.Enabled = False
        SaveOldReg '保存旧的注册表
        ReFillInfo
    End If
End Sub


Private Sub cboLinkAccept_Click() '单位互联
    Dim oSysMan As New SystemMan
    Dim tTemp() As TDepartmentInfo
    Dim szUnitID As String
    Dim nCount As Integer
    Dim i As Integer
    Dim aConnectUnit() As TConnectUnit
    Dim UserPassWord As String

    mbLoad = False
On Error GoTo here
    If mbLoad = False Then
        SetBusy
        oFreeReg.Init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany
        oSysMan.Init m_oAUser
        szUnitID = ResolveDisplay(cboLinkAccept.Text)
    
        If szUnitID <> m_oParam.UnitID Then
            aConnectUnit = moLugSvr.GetConnectUnit
            nCount = ArrayLength(aConnectUnit)
            For i = 1 To nCount
                If aConnectUnit(i).szUnitID = szUnitID Then
                    oFreeReg.SaveSetting m_cRegParamKey, "DBServer", aConnectUnit(i).szIPAddress
                    oFreeReg.SaveSetting m_cRegParamKey, "Database", aConnectUnit(i).szDatabase
                    oFreeReg.SaveSetting m_cRegParamKey, "User", aConnectUnit(i).szDBUser
                    oFreeReg.SaveSetting m_cRegParamKey, "Password", aConnectUnit(i).szUserPassword
                End If
            Next i
        End If
        Set oFreeReg = Nothing
        SetNormal
    
        Dim m_oAUserTmp As New ActiveUser
        m_oAUserTmp.Login m_oAUser.UserID, "", GetComputerName
        
        
        ReFillInfo
        mbLoad = True
    
    End If
    Exit Sub

here:

      MsgBox err.Description, vbCritical + vbOKOnly, "提示:"
   Set oFreeReg = Nothing
   SetNormal
'    tTemp = oSysMan.GetAllSellStation(szUnitID)

'    nCount = ArrayLength(tTemp)
'    cboSellStation.Clear
'    cboSellStation.AddItem ""
'    For i = 1 To nCount
'         cboSellStation.AddItem MakeDisplayString(tTemp(i).szSellStationID, tTemp(i).szSellStationName)
'
'    Next i
End Sub

Private Sub Station() '新增联网受理站点
    Dim oLuggagesvr As New LuggageSvr
    Dim tTemp() As TDepartmentInfo
'    Dim g_atAllUnit() As TConnectUnit
    Dim szUnitID As String
    Dim nCount As Integer
    Dim nUnitCount As Integer
    Dim i As Integer
    Dim j As Integer
    Dim g_oSysParam  As New SystemParam
    
    g_oSysParam.Init m_oAUser
    oLuggagesvr.Init m_oAUser
    cboLinkAccept.Clear
'    g_atAllUnit = oLuggagesvr.GetConnectUnit
'    nUnitCount = ArrayLength(g_atAllUnit)
'    cboLinkAccept.AddItem ""
'    For i = 1 To nUnitCount
'        cboLinkAccept.AddItem MakeDisplayString(g_atAllUnit(i).szUnitID, g_atAllUnit(i).szUnitFullName)
'        If g_atAllUnit(i).szUnitID = g_oSysParam.UnitID Then
'            szUnitID = g_atAllUnit(i).szUnitID
'            cboLinkAccept.Text = cboLinkAccept.List(i)
'        End If
'    Next i

'    tTemp = oSysMan.GetAllSellStation(szUnitID)
'
'    nCount = ArrayLength(tTemp)
'
'    cboSellStation.AddItem ""
'    For i = 1 To nCount
'         cboSellStation.AddItem MakeDisplayString(tTemp(i).szSellStationID, tTemp(i).szSellStationName)
'
'    Next i
End Sub

Private Sub SaveOldReg() '保存旧的注册表
    oFreeReg.Init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany
    oFreeReg.SaveSetting m_cRegParamKey, "DBServer", DBServer
    oFreeReg.SaveSetting m_cRegParamKey, "Database", Database
    oFreeReg.SaveSetting m_cRegParamKey, "Password", Password
    oFreeReg.SaveSetting m_cRegParamKey, "User", User
    Set oFreeReg = Nothing
End Sub

Private Sub ReFillInfo() '联网后重新载入信息
    GetParam
    '得到起始的行包单号和签发单号
    GetAppSetting
'    mdiMain.lblEndSheetNo.Visible = True
    frmChgSheetNo.Show vbModal
    If Not frmChgSheetNo.m_bOk Then
        End
    End If
    SetSheetNoLabel True, g_szAcceptSheetID '设置主界面上的标签号

    RefreshStation '刷新站点
    InitGrid '填充收费项
End Sub


Private Sub ReadReg() '读注册表
On Error GoTo ErrorHandle

    oFreeReg.Init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany
    
    DBServer = oFreeReg.GetSetting(m_cRegParamKey, "DBServer")
    Database = oFreeReg.GetSetting(m_cRegParamKey, "Database")
    User = oFreeReg.GetSetting(m_cRegParamKey, "User")
    Password = oFreeReg.GetSetting(m_cRegParamKey, "Password")
    
    Set oFreeReg = Nothing
Exit Sub

ErrorHandle:
   ShowMsg err.Description
End Sub

Public Sub GetParam()
         Dim oFreeReg As CFreeReg
      
         '系统参数
         moSysParam.Init m_oAUser
         m_oParam.Init m_oAUser
         Date = m_oParam.NowDate
         Time = m_oParam.NowDateTime

         m_bIsRelationWithVehicleType = False 'm_oParam.IsRelationWithVehicleType '行包的公式是否与车型有关系
         m_bIsDispSettlePriceInAccept = False 'm_oParam.IsDispSettlePriceInAccept '是否在受理时显示应结费用
         m_bIsDispSettlePriceInCheck = False 'm_oParam.IsDispSettlePriceInCheck '是否在签发时显示应结费用
         m_bIsSettlePriceFromAcceptInCheck = False 'm_oParam.IsSettlePriceFromAcceptInCheck '结算运费是不是从受理的结算运费中汇总得到
         m_bIsPrintCheckSheet = False 'm_oParam.IsPrintCheckSheet '是否打印签发单
        
         SetHTMLHelpStrings "pstLugDesk.chm"
         
         '读出注册表里的注意事项
         '行包注意事项
         
         Set oFreeReg = New CFreeReg
         oFreeReg.Init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany
         m_szCustom = oFreeReg.GetSetting(cszLuggageAccount, "CareContent") '自定义信息

End Sub
