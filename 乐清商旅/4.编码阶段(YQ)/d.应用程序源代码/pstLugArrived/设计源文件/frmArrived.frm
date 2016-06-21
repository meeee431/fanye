VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmArrived 
   BackColor       =   &H8000000C&
   Caption         =   "行包到达受理"
   ClientHeight    =   7080
   ClientLeft      =   1395
   ClientTop       =   2490
   ClientWidth     =   11400
   ControlBox      =   0   'False
   HelpContextID   =   7000020
   Icon            =   "frmArrived.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7080
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtOwnID 
      Height          =   300
      Left            =   3405
      TabIndex        =   83
      Top             =   7680
      Visible         =   0   'False
      Width           =   915
   End
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
      TabIndex        =   71
      Top             =   4680
      Visible         =   0   'False
      Width           =   2325
      Begin MSComctlLib.ListView lvprice 
         Height          =   3495
         Left            =   120
         TabIndex        =   72
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
      TabIndex        =   64
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
         TabIndex        =   65
         Top             =   750
         Width           =   2055
         Begin VB.TextBox txtPrice 
            Height          =   315
            Left            =   1080
            TabIndex        =   67
            Text            =   "0"
            Top             =   720
            Width           =   825
         End
         Begin VB.TextBox txtCal 
            Height          =   315
            Left            =   120
            TabIndex        =   66
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
            TabIndex        =   69
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
            TabIndex        =   68
            Top             =   360
            Width           =   480
         End
      End
      Begin RTComctl3.CoolButton cmdOver 
         Height          =   345
         Left            =   360
         TabIndex        =   70
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
         MICON           =   "frmArrived.frx":076A
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
      Caption         =   "&H00C0C0C0&"
      Height          =   6735
      Left            =   360
      TabIndex        =   60
      Top             =   120
      Width           =   11325
      Begin VB.ComboBox cboPickOperator 
         Enabled         =   0   'False
         Height          =   300
         Left            =   5040
         TabIndex        =   85
         Top             =   5610
         Width           =   1395
      End
      Begin VB.ComboBox txtSMS 
         Enabled         =   0   'False
         Height          =   300
         Left            =   150
         TabIndex        =   48
         Top             =   4110
         Width           =   6615
      End
      Begin VB.Frame fraCustomerInfo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "收件信息"
         Height          =   1140
         Left            =   165
         TabIndex        =   62
         Top             =   2865
         Width           =   8025
         Begin VB.TextBox txtShipperPhone 
            Height          =   315
            Left            =   4125
            MaxLength       =   30
            TabIndex        =   30
            Text            =   "托电话"
            Top             =   270
            Width           =   2445
         End
         Begin VB.TextBox txtShipper 
            Height          =   315
            Left            =   1515
            MaxLength       =   20
            TabIndex        =   28
            Text            =   "托运人"
            Top             =   270
            Width           =   1395
         End
         Begin VB.TextBox txtShipperUnit 
            Height          =   315
            Left            =   1725
            MaxLength       =   20
            TabIndex        =   32
            Text            =   "托运人"
            Top             =   660
            Width           =   4845
         End
         Begin VB.Label lblShippePhone 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "电话(&H):"
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
            Left            =   3015
            TabIndex        =   29
            Top             =   315
            Width           =   840
         End
         Begin VB.Label lblShipper 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "发件人(&E):"
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
            TabIndex        =   27
            Top             =   315
            Width           =   1050
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "发件人单位(&U):"
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
            TabIndex        =   31
            Top             =   705
            Width           =   1470
         End
      End
      Begin VB.ComboBox cboOperator 
         Height          =   300
         Left            =   1725
         TabIndex        =   42
         Top             =   5610
         Width           =   1395
      End
      Begin VB.Frame fraPickBill 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   6495
         Left            =   8310
         TabIndex        =   74
         Top             =   210
         Width           =   3030
         Begin VB.TextBox txtDrawer 
            Height          =   315
            Left            =   210
            TabIndex        =   51
            Top             =   2310
            Width           =   2655
         End
         Begin VB.TextBox txtDrawerPhone 
            Height          =   315
            Left            =   210
            TabIndex        =   52
            Text            =   "(电话)"
            Top             =   2670
            Width           =   2655
         End
         Begin VB.TextBox txtPickerCreditID 
            Height          =   315
            Left            =   210
            TabIndex        =   54
            Top             =   3300
            Width           =   2655
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
            Left            =   1425
            TabIndex        =   57
            Text            =   "0.0"
            Top             =   4740
            Width           =   1365
         End
         Begin VSFlex7LCtl.VSFlexGrid vsPriceItem 
            Height          =   1455
            Left            =   210
            TabIndex        =   50
            Top             =   570
            Width           =   2655
            _cx             =   4683
            _cy             =   2566
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
         Begin RTComctl3.CoolButton cmdAccept 
            Height          =   615
            Left            =   300
            TabIndex        =   58
            Top             =   5730
            Width           =   2505
            _ExtentX        =   4419
            _ExtentY        =   1085
            BTYPE           =   3
            TX              =   "结算(F2)"
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
            MICON           =   "frmArrived.frx":0786
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSComCtl2.DTPicker dtpPickTime 
            Height          =   300
            Left            =   180
            TabIndex        =   56
            Top             =   3960
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy年MM月dd日 HH:mm"
            Format          =   117899267
            UpDown          =   -1  'True
            CurrentDate     =   38646
         End
         Begin VB.Label lblTransCharge 
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
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   1650
            TabIndex        =   88
            Top             =   4440
            Width           =   135
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "提件人姓名及电话:"
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
            TabIndex        =   87
            Top             =   2070
            Width           =   1785
         End
         Begin VB.Label lblSheetID 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "1111111111"
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
            Left            =   990
            TabIndex        =   81
            Top             =   30
            Width           =   1050
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "单据号:"
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
            TabIndex        =   80
            Top             =   30
            Width           =   735
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "提件时间(&T):"
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
            TabIndex        =   55
            Top             =   3690
            Width           =   1260
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "提件人身份证(&I):"
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
            TabIndex        =   53
            Top             =   3060
            Width           =   1680
         End
         Begin VB.Label flblRestMoney 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0.0"
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
            Left            =   2160
            TabIndex        =   78
            Top             =   5310
            Width           =   585
         End
         Begin VB.Label flblTotalPrice 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0.0"
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
            Left            =   2160
            TabIndex        =   77
            Top             =   4350
            Width           =   585
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
            Left            =   195
            TabIndex        =   76
            Top             =   4440
            Width           =   525
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
            Left            =   195
            TabIndex        =   59
            Top             =   4905
            Width           =   840
         End
         Begin VB.Line Line4 
            BorderColor     =   &H00000000&
            X1              =   210
            X2              =   2850
            Y1              =   5250
            Y2              =   5250
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
            Left            =   195
            TabIndex        =   75
            Top             =   5400
            Width           =   525
         End
         Begin VB.Label lblPriceItem 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "托运费用(F9):"
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
            TabIndex        =   49
            Top             =   330
            Width           =   1365
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Height          =   1125
         Left            =   150
         TabIndex        =   73
         Top             =   4440
         Width           =   8025
         Begin VB.TextBox txtRemark 
            Height          =   315
            Left            =   1560
            TabIndex        =   40
            Top             =   630
            Width           =   6285
         End
         Begin VB.TextBox txtTransCharge 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   7050
            TabIndex        =   38
            Top             =   218
            Width           =   795
         End
         Begin VB.ComboBox cboSavePosition 
            Height          =   300
            Left            =   4320
            TabIndex        =   36
            Top             =   225
            Width           =   1290
         End
         Begin VB.ComboBox cboLoader 
            Height          =   300
            Left            =   1560
            TabIndex        =   34
            Top             =   225
            Width           =   1395
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "备注(&M):"
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
            TabIndex        =   39
            Top             =   675
            Width           =   840
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "代收运费(&Q):"
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
            Left            =   5730
            TabIndex        =   37
            Top             =   270
            Width           =   1260
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "存放位置(&T):"
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
            Left            =   3030
            TabIndex        =   35
            Top             =   270
            Width           =   1260
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "装卸工(&J):"
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
            TabIndex        =   33
            Top             =   270
            Width           =   1050
         End
      End
      Begin VB.Frame fraLuggage 
         BackColor       =   &H00E0E0E0&
         Caption         =   "行包信息"
         Height          =   2700
         Left            =   150
         TabIndex        =   61
         Top             =   150
         Width           =   8025
         Begin VB.TextBox txtPackageID 
            Enabled         =   0   'False
            Height          =   315
            Left            =   990
            Locked          =   -1  'True
            TabIndex        =   1
            Top             =   263
            Width           =   900
         End
         Begin VB.TextBox txtPicker 
            Height          =   315
            Left            =   1560
            MaxLength       =   20
            TabIndex        =   18
            Text            =   "收件人"
            Top             =   1425
            Width           =   1395
         End
         Begin VB.TextBox txtPickerPhone 
            Height          =   315
            Left            =   4140
            MaxLength       =   30
            TabIndex        =   20
            Text            =   "13357567557"
            Top             =   1455
            Width           =   1425
         End
         Begin VB.TextBox txtPickerAddress 
            Height          =   315
            Left            =   1545
            MaxLength       =   30
            TabIndex        =   26
            Text            =   "收地址"
            Top             =   2280
            Visible         =   0   'False
            Width           =   5055
         End
         Begin VB.ComboBox cboPickType 
            Height          =   300
            ItemData        =   "frmArrived.frx":07A2
            Left            =   6630
            List            =   "frmArrived.frx":07AC
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   1425
            Width           =   1245
         End
         Begin VB.TextBox txtPickerUnit 
            Height          =   315
            Left            =   2370
            MaxLength       =   30
            TabIndex        =   24
            Text            =   "收件人"
            Top             =   1890
            Width           =   4245
         End
         Begin MSComCtl2.DTPicker dtpArriveTime 
            Height          =   300
            Left            =   5595
            TabIndex        =   10
            Top             =   645
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy年MM月dd日 HH:mm"
            Format          =   117899267
            UpDown          =   -1  'True
            CurrentDate     =   38646
         End
         Begin VB.ComboBox cboWeight 
            Height          =   300
            Left            =   6600
            TabIndex        =   16
            Top             =   1020
            Width           =   1260
         End
         Begin VB.TextBox txtStartStation 
            Height          =   300
            Left            =   2970
            MaxLength       =   10
            TabIndex        =   8
            Text            =   "杭州"
            Top             =   645
            Width           =   1185
         End
         Begin VB.ComboBox cboAreaType 
            Height          =   300
            Left            =   1560
            TabIndex        =   7
            Text            =   "省内"
            Top             =   645
            Width           =   1395
         End
         Begin VB.TextBox txtNums 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4170
            TabIndex        =   14
            Text            =   "10"
            Top             =   1020
            Width           =   1425
         End
         Begin VB.TextBox txtLicense 
            Height          =   315
            Left            =   1560
            TabIndex        =   12
            Text            =   "000000"
            Top             =   1020
            Width           =   1395
         End
         Begin VB.ComboBox cboPack 
            Height          =   300
            Left            =   6630
            TabIndex        =   5
            Top             =   270
            Width           =   1260
         End
         Begin VB.ComboBox cboPackageName 
            Height          =   300
            Left            =   3285
            TabIndex        =   3
            Top             =   270
            Width           =   2445
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "自编号:"
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
            TabIndex        =   0
            Top             =   285
            Width           =   735
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
            TabIndex        =   17
            Top             =   1470
            Width           =   1050
         End
         Begin VB.Label lblPhone 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "电话(&O):"
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
            Left            =   3060
            TabIndex        =   19
            Top             =   1500
            Width           =   840
         End
         Begin VB.Label lblPickerAddress 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "地址(&R):"
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
            TabIndex        =   25
            Top             =   2325
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.Label lblPickType 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "交付(&Y):"
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
            Left            =   5730
            TabIndex        =   21
            Top             =   1470
            Width           =   840
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "收件人单位或地址(&I):"
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
            TabIndex        =   23
            Top             =   1935
            Width           =   2100
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "到达时间(&T):"
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
            Left            =   4290
            TabIndex        =   9
            Top             =   690
            Width           =   1260
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "起运站(&F):"
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
            TabIndex        =   6
            Top             =   690
            Width           =   1050
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
            Left            =   5760
            TabIndex        =   4
            Top             =   315
            Width           =   840
         End
         Begin VB.Label lblLuggageName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "行包名称(&N):"
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
            Left            =   1965
            TabIndex        =   2
            Top             =   315
            Width           =   1230
         End
         Begin VB.Label lblActWeight 
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
            Left            =   5760
            TabIndex        =   15
            Top             =   1072
            Width           =   840
         End
         Begin VB.Label lblBagNum 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "件数(&M):"
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
            Left            =   3060
            TabIndex        =   13
            Top             =   1065
            Width           =   840
         End
         Begin VB.Label lblCalWeight 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "车牌号(&L):"
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
            TabIndex        =   11
            Top             =   1072
            Width           =   1050
         End
      End
      Begin RTComctl3.CoolButton cmdSave 
         Height          =   375
         Left            =   3824
         TabIndex        =   43
         Top             =   6180
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "保存(&S)"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
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
         MICON           =   "frmArrived.frx":07BC
         PICN            =   "frmArrived.frx":07D8
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
         Height          =   375
         Left            =   5286
         TabIndex        =   46
         Top             =   6180
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "作废(&D)"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
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
         MICON           =   "frmArrived.frx":0A56
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin RTComctl3.CoolButton cmdClose 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   6750
         TabIndex        =   47
         Top             =   6180
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "关闭"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
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
         MICON           =   "frmArrived.frx":0A72
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin RTComctl3.CoolButton cmdPick 
         Height          =   375
         Left            =   2362
         TabIndex        =   44
         Top             =   6180
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "提货(&P)"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
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
         MICON           =   "frmArrived.frx":0A8E
         PICN            =   "frmArrived.frx":0AAA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin RTComctl3.CoolButton cmdSMS 
         Height          =   315
         Left            =   6840
         TabIndex        =   82
         Top             =   4110
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "短信发送"
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
         MICON           =   "frmArrived.frx":0E44
         PICN            =   "frmArrived.frx":0E60
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin RTComctl3.CoolButton cmdAddNew 
         Height          =   375
         Left            =   900
         TabIndex        =   45
         Top             =   6180
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "新增(&A)"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
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
         MICON           =   "frmArrived.frx":11FA
         PICN            =   "frmArrived.frx":1216
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblPickOperator 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "提货受理人(&O):"
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
         Left            =   3600
         TabIndex        =   86
         Top             =   5655
         Width           =   1470
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "到达受理人(&C):"
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
         Left            =   270
         TabIndex        =   41
         Top             =   5655
         Width           =   1470
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   180
         X2              =   8160
         Y1              =   6015
         Y2              =   6015
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   180
         X2              =   8160
         Y1              =   6000
         Y2              =   6000
      End
      Begin VB.Label lblStatus 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "已结"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   315
         Left            =   7470
         TabIndex        =   79
         Top             =   5610
         Width           =   660
      End
   End
   Begin RTComctl3.CoolButton cmdbegin 
      Height          =   345
      Left            =   13215
      TabIndex        =   63
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
      MICON           =   "frmArrived.frx":14BC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "自编号(&O):"
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
      Left            =   2055
      TabIndex        =   84
      Top             =   7710
      Visible         =   0   'False
      Width           =   1050
   End
End
Attribute VB_Name = "frmArrived"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'*******************************************************************************
'Last Modify By: 陆勇庆  2005-8-16
'Last Modify In:
'remark by lyq at 2006-1-25:添加行包编号规则为yymm000000，但输入时只输尾部数字，其他自动生成
'*******************************************************************************

Option Explicit

Const cnMargin = 15

Public Status As EFormStatus
Public m_lPackageID As Long
Public m_bIsParent As Boolean
Private Const cszDrawerPhone = "(电话)"

Private m_oPackage As New Package
Private mbSizeExpand As Boolean      '界面尺寸是否已经缩小
'以下定义常用短信标识
Const cSMSDefine_Picker = "[收件人]"
Const cSMSDefine_Shipper = "[发件人]"
Const cSMSDefine_ShartStation = "[起运站]"
Const cSMSDefine_PackageName = "[行包名称]"
Const cSMSDefine_PackageID = "[编号]"
Const cSMSDefine_PackageNumber = "[件数]"





Private Sub cboAreaType_LostFocus()
    SetBaseDefine cboAreaType, EDefineType.EDT_AreaType
'    SynSMSString

End Sub

Private Sub cboLoader_LostFocus()
    SetBaseDefine cboLoader, EDefineType.EDT_LoadWorker

End Sub

Private Sub cboOperator_LostFocus()
    SetBaseDefine cboOperator, EDefineType.EDT_Operator
End Sub

Private Sub cboPack_LostFocus()
    SetBaseDefine cboPack, EDefineType.EDT_PackType
End Sub

Private Sub cboPackageName_Change()
'    SynSMSString
End Sub

Private Sub cboPackageName_LostFocus()
    SetBaseDefine cboPackageName, EDefineType.EDT_PackageName

End Sub

Private Sub cboSavePosition_LostFocus()
    SetBaseDefine cboSavePosition, EDefineType.EDT_SavePosition

End Sub

Private Sub cboWeight_LostFocus()
    If mbSizeExpand And cboWeight.Text <> m_oPackage.CalWeight Then MsgBox "重量被更改，请先按[保存]按钮进行保存，再提货！", vbInformation
End Sub

Private Sub cmdAccept_Click()
On Error GoTo ErrHandle
    m_oPackage.PackageID = m_lPackageID
    m_oPackage.SheetID = lblSheetID.Caption
    m_oPackage.LoadCharge = Val(vsPriceItem.TextMatrix(0, 1))
    m_oPackage.KeepCharge = Val(vsPriceItem.TextMatrix(1, 1))
    m_oPackage.SendCharge = Val(vsPriceItem.TextMatrix(2, 1))
    m_oPackage.MoveCharge = Val(vsPriceItem.TextMatrix(3, 1))
    m_oPackage.OtherCharge = Val(vsPriceItem.TextMatrix(4, 1))
    m_oPackage.PickerCreditID = txtPickerCreditID.Text
    m_oPackage.PickTime = dtpPickTime.Value
    m_oPackage.Drawer = txtDrawer.Text
    m_oPackage.DrawerPhone = IIf(txtDrawerPhone.Text = cszDrawerPhone, "", txtDrawerPhone.Text)
    m_oPackage.Shipper = Trim(txtShipper.Text)
    g_oPackageSvr.PickPackage m_oPackage
    
    
    IncSheetID
    cmdAccept.Enabled = False
    lblStatus.Caption = CPick_Picked
    lblStatus.ForeColor = vbRed
    
    txtNums.Enabled = False
    cboWeight.Enabled = False
    txtTransCharge.Enabled = False
    vsPriceItem.Enabled = False
    
    'PrintAcceptSheet m_oPackage
    
    frmSheet.PrintSheetReport m_oPackage

   
    
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub cmdAddNew_Click()
On Error GoTo ErrHandle
    Status = EFS_AddNew
    RefreshForm
    cmdSave.Enabled = True
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub cmdCancel_Click()
On Error GoTo ErrHandle
    m_oPackage.Remark = txtRemark.Text
    g_oPackageSvr.CancelPackage m_oPackage
    RefreshForm
            
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPick_Click()
    
    If Not mbSizeExpand Then
        fraOutLine.Width = fraOutLine.Width + fraPickBill.Width
        fraPickBill.Visible = True
        Form_Resize
        mbSizeExpand = True
    
        
        lblSheetID.Caption = g_szSheetID
        dtpPickTime.Value = Format(Now, "yyyy-MM-dd HH:mm")
        CalKeepCharge
        CalLoadCharge
        cmdPick.Enabled = False
'***********************************fpd
'温州金额为0也允许提货
        cmdAccept.Enabled = True
'***********************************
        lblPickOperator.Visible = True
        cboPickOperator.Visible = True
        cboPickOperator.Text = g_oActUser.UserID
        txtDrawer.Text = txtPicker.Text
        txtDrawerPhone.Text = cszDrawerPhone
        txtPickerCreditID.Text = ""
        
        '自动计保管费
'        CalKeepCharge
        
        vsPriceItem.SetFocus
    End If
    
End Sub
'自动计算保管费
Private Sub CalKeepCharge()
On Error GoTo ErrHandle
    Dim dbCharge As Double
       
    '当天不要收
    '
    dbCharge = (DateDiff("d", DateAdd("d", g_oPackageParam.KeepFeeDays, dtpArriveTime.Value), dtpPickTime.Value) + 1) * Val(txtNums.Text) * g_oPackageParam.NormalKeepCharge
    If dbCharge > 0 Then
        vsPriceItem.TextMatrix(1, 1) = dbCharge
    Else
        vsPriceItem.TextMatrix(1, 1) = 0
    End If

    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub
'自动计算装卸费
Private Sub CalLoadCharge()
On Error GoTo ErrHandle
    Dim dbCharge As Double
    Dim aszTmp() As String
    Dim i As Integer
    Dim MyPos As Integer
    Dim sz1 As String
    Dim sz2 As String
    aszTmp = g_oPackageParam.ListLoadChargeCode()
    For i = 1 To ArrayLength(aszTmp) '谢晓冬07-07-06改
        MyPos = InStr(1, aszTmp(i, 2), "-", 1)
        If MyPos > 0 Then
            sz1 = Val(Left(aszTmp(i, 2), MyPos - 1))
            sz2 = Val(Right(aszTmp(i, 2), Len(aszTmp(i, 2)) - MyPos))
            If Val(cboWeight.Text) >= sz1 And Val(cboWeight.Text) <= sz2 Then
                dbCharge = Val(aszTmp(i, 3))
                Exit For
            End If
    
    
        Else
    
            If Val(cboWeight.Text) >= Val(aszTmp(i, 2)) Then
                dbCharge = Val(aszTmp(i, 3))
                Exit For
            End If
        End If
    Next
    
    For i = 1 To ArrayLength(aszTmp)
        If Trim(cboWeight.Text) = Trim(aszTmp(i, 2)) Then
            dbCharge = Val(aszTmp(i, 3))
            Exit For
        End If
    Next
    If dbCharge > 0 Then
        vsPriceItem.TextMatrix(0, 1) = dbCharge * Val(txtNums)
    Else
         vsPriceItem.TextMatrix(0, 1) = 0
    End If

    '上面是以前的算法，下面是嵊州的算法，按件数来的，默认每件一元，超重的让他们自己手动改
'    vsPriceItem.TextMatrix(0, 1) = Val(txtNums) * 2
    
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub
Private Sub cmdSave_Click()
On Error GoTo ErrHandle
    If cboAreaType.Text = "" Then
        MsgBox "必须输入起运站！", vbExclamation, Me.Caption
        cboAreaType.SetFocus
        Exit Sub
    End If
    If Val(txtNums.Text) = 0 Then
        MsgBox "必须输入件数！", vbExclamation, Me.Caption
        txtNums.SetFocus
        Exit Sub
    End If
    If txtPicker.Text = "" Then
        MsgBox "必须收件人！", vbExclamation, Me.Caption
        txtPicker.SetFocus
        Exit Sub
    End If


    SetBusy
    Select Case Status
        Case EFormStatus.EFS_AddNew
            AddNewPackage
            
        Case EFormStatus.EFS_Modify
            UpdatePackage
        
    End Select
    LayoutForm
    cmdSave.Enabled = False
    SetNormal
    
    
    Exit Sub
ErrHandle:
    SetNormal
    ShowErrorMsg
End Sub
Private Sub UpdatePackage()
    m_oPackage.PackageID = m_lPackageID
'    m_oPackage.OwnID = txtOwnID.Text
    m_oPackage.PackageName = cboPackageName.Text
    m_oPackage.PackType = cboPack.Text
    m_oPackage.AreaType = cboAreaType.Text
    m_oPackage.StartStationName = txtStartStation.Text
    m_oPackage.PackageNumber = txtNums.Text
    m_oPackage.ArrivedTime = dtpArriveTime.Value
    m_oPackage.CalWeight = cboWeight.Text
    m_oPackage.Shipper = txtShipper.Text
    m_oPackage.ShipperPhone = txtShipperPhone.Text
    m_oPackage.ShipperUnit = txtShipperUnit.Text
    m_oPackage.Picker = txtPicker.Text
    m_oPackage.PickerAddress = txtPickerAddress.Text
    m_oPackage.PickerPhone = txtPickerPhone.Text
    m_oPackage.PickerUnit = txtPickerUnit.Text
    m_oPackage.PickType = cboPickType.Text
    m_oPackage.Loader = cboLoader.Text
    m_oPackage.SavePosition = cboSavePosition.Text
    m_oPackage.TransitCharge = Val(txtTransCharge.Text)
    m_oPackage.Operator = cboOperator.Text
    m_oPackage.Remark = txtRemark.Text
    m_oPackage.LicenseTagNo = txtLicense.Text
    
    m_oPackage.LoadCharge = m_oPackage.LoadCharge
    m_oPackage.KeepCharge = m_oPackage.KeepCharge
    m_oPackage.SendCharge = m_oPackage.SendCharge
    m_oPackage.MoveCharge = m_oPackage.MoveCharge
    m_oPackage.OtherCharge = m_oPackage.OtherCharge
    m_oPackage.PickerCreditID = txtPickerCreditID.Text
    m_oPackage.PickTime = IIf(m_oPackage.Status <> EPS_Picked, cdtEmptyDate, dtpPickTime.Value)
    m_oPackage.Drawer = txtDrawer.Text
    m_oPackage.DrawerPhone = IIf(txtDrawerPhone.Text = cszDrawerPhone, "", txtDrawerPhone.Text)
    
    m_oPackage.Update (ST_EditObj)

End Sub
Private Sub AddNewPackage()
    m_oPackage.AddNew
    m_oPackage.PackageID = BuildPacketID(Val(txtPackageID.Text))
    m_oPackage.PackageName = cboPackageName.Text
    m_oPackage.PackType = cboPack.Text
    m_oPackage.AreaType = cboAreaType.Text
    m_oPackage.StartStationName = txtStartStation.Text
    m_oPackage.PackageNumber = txtNums.Text
    m_oPackage.ArrivedTime = dtpArriveTime.Value
    m_oPackage.CalWeight = cboWeight.Text
    m_oPackage.Shipper = txtShipper.Text
    m_oPackage.ShipperPhone = txtShipperPhone.Text
    m_oPackage.ShipperUnit = txtShipperUnit.Text
    m_oPackage.Picker = txtPicker.Text
    m_oPackage.PickerAddress = txtPickerAddress.Text
    m_oPackage.PickerPhone = txtPickerPhone.Text
    m_oPackage.PickerUnit = txtPickerUnit.Text
    m_oPackage.PickType = cboPickType.Text
    m_oPackage.Loader = cboLoader.Text
    m_oPackage.SavePosition = cboSavePosition.Text
    m_oPackage.TransitCharge = Val(txtTransCharge.Text)
    m_oPackage.Operator = cboOperator.Text
    m_oPackage.Remark = txtRemark.Text
    m_oPackage.LicenseTagNo = txtLicense.Text


    
'    m_oPackage.Update
    g_oPackageSvr.AcceptPackage m_oPackage
    m_lPackageID = m_oPackage.PackageID
    txtPackageID.Text = UnBuildPacketID(m_lPackageID)
    
    Status = EFS_Modify
End Sub

Private Sub cmdSMS_Click()
    '判断是否为手机号
    Dim szPhone As String
    szPhone = Trim(txtPickerPhone.Text)
    szPhone = IIf(Len(szPhone) > 11, Left(szPhone, 11), szPhone)
    Select Case Left(szPhone, 2)
        Case "13"
        Case Else
            MsgBox "收件人的手机号码不准确，请检查后重试!", vbExclamation, "错误"
            Exit Sub
    End Select
    SendSMS szPhone, txtSMS.Text
End Sub

Private Sub dtpArriveTime_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then
'        KeyCode = 0
'        Debug.Print "dtpPickTime_KeyDown 13"
'        txtLicense.SetFocus
'    End If

End Sub

Private Sub dtpPickTime_Change()
    CalKeepCharge
End Sub

Private Sub dtpPickTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        
        txtReceivedMoney.SetFocus
    End If

End Sub

Private Sub flblTotalPrice_Change()
    If Val(flblTotalPrice.Caption) >= 0 Then
        cmdAccept.Enabled = True
    Else
        cmdAccept.Enabled = False
    End If
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
            cmdAccept_Click
        End If
    ElseIf KeyCode = vbKeyReturn And (Me.ActiveControl Is cboPickType) Then
        cboLoader.SetFocus
    ElseIf KeyCode = vbKeyReturn Then
'        KeyCode = 0
        SendKeys "{TAB}"


'        Debug.Print "Form_KeyDown 13"
    End If

    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub


'填充数据字典
Private Sub FillBaseInfo()
    Dim aszTmp() As String
    aszTmp = g_oPackageParam.ListBaseDefine(EDefineType.EDT_PackageName)
    cboPackageName.Clear
    Dim i As Integer
    For i = 1 To ArrayLength(aszTmp)
        cboPackageName.AddItem aszTmp(i, 3)
    Next i
    aszTmp = g_oPackageParam.ListBaseDefine(EDefineType.EDT_AreaType)
    cboAreaType.Clear
    For i = 1 To ArrayLength(aszTmp)
        cboAreaType.AddItem aszTmp(i, 3)
    Next i
    aszTmp = g_oPackageParam.ListBaseDefine(EDefineType.EDT_PackType)
    cboPack.Clear
    For i = 1 To ArrayLength(aszTmp)
        cboPack.AddItem aszTmp(i, 3)
    Next i


    aszTmp = g_oPackageParam.ListBaseDefine(EDefineType.EDT_LoadWorker)
    cboLoader.Clear
    For i = 1 To ArrayLength(aszTmp)
        cboLoader.AddItem aszTmp(i, 3)
    Next i
    aszTmp = g_oPackageParam.ListBaseDefine(EDefineType.EDT_Operator)
    cboOperator.Clear
    For i = 1 To ArrayLength(aszTmp)
        cboOperator.AddItem aszTmp(i, 3)
    Next i

    aszTmp = g_oPackageParam.ListBaseDefine(EDefineType.EDT_SavePosition)
    cboSavePosition.Clear
    For i = 1 To ArrayLength(aszTmp)
        cboSavePosition.AddItem aszTmp(i, 3)
    Next i

    aszTmp = g_oPackageParam.ListBaseDefine(EDefineType.EDT_Other1)
    txtSMS.Clear
    For i = 1 To ArrayLength(aszTmp)
        txtSMS.AddItem aszTmp(i, 4)
    
    Next i
    
    aszTmp = g_oPackageParam.ListLoadChargeCode
    cboWeight.Clear
    For i = 1 To ArrayLength(aszTmp)
        cboWeight.AddItem aszTmp(i, 2)
    Next i


End Sub


Private Sub Form_Load()
    '加入短信初始化
    On Error Resume Next
    InitSMS
    On Error GoTo Here
    
    '嵊州不用短信
    'If Not g_bSMSValid Then
        'MsgBox "短信息设备初始化不成功，无法使用该功能!", vbExclamation
        cmdSMS.Enabled = False
    'Else
        'cmdSMS.Enabled = True
    'End If
    
    mbSizeExpand = True
    
    m_oPackage.init g_oActUser

    FillBaseInfo
    
    vsPriceItem.Cols = 2
    vsPriceItem.FixedCols = 1
    vsPriceItem.Rows = 5
    vsPriceItem.ColWidth(0) = 1000 'vsPriceItem.Width * 0.6
    vsPriceItem.ColWidth(1) = 1000 ' vsPriceItem.Width * 0.4
    vsPriceItem.TextMatrix(0, 0) = "装卸费"
    vsPriceItem.TextMatrix(1, 0) = "保管费"
    vsPriceItem.TextMatrix(2, 0) = "送货费"
    vsPriceItem.TextMatrix(3, 0) = "搬运费"
    vsPriceItem.TextMatrix(4, 0) = "其他费"
    
    LayoutForm
    If Me.Status = EFS_AddNew Then
        RefreshForm
    End If
    Exit Sub
Here:
ShowErrorMsg
End Sub
Private Sub LayoutForm()
  
  Select Case Status
    Case EFormStatus.EFS_AddNew
        If mbSizeExpand Then
            fraPickBill.Visible = False
            fraOutLine.Width = fraOutLine.Width - fraPickBill.Width
            mbSizeExpand = False
        End If
    
        cmdAddNew.Enabled = False
        cmdPick.Enabled = False
        cmdSave.Enabled = True
        cmdCancel.Enabled = False
        cmdAccept.Enabled = False
        lblPickOperator.Visible = False
        cboPickOperator.Visible = False
        
        txtDrawer.Enabled = True
        dtpPickTime.Enabled = True
        cboAreaType.Enabled = True
        txtStartStation.Enabled = True
        cboOperator.Enabled = True
        cboPickOperator.Enabled = True
        dtpArriveTime.Enabled = True
        
    Case EFormStatus.EFS_Modify
        Select Case m_oPackage.Status
            Case EPS_Normal
                If mbSizeExpand Then
                    fraPickBill.Visible = False
                    fraOutLine.Width = fraOutLine.Width - fraPickBill.Width
                    mbSizeExpand = False
                End If
                cmdAddNew.Enabled = True
                cmdPick.Enabled = True
                cmdSave.Enabled = True
                cmdCancel.Enabled = True
                cmdAccept.Enabled = False
                lblPickOperator.Visible = False
                cboPickOperator.Visible = False
                txtPackageID.Enabled = False
            Case EPS_Picked
                If Not mbSizeExpand Then
                    fraPickBill.Visible = True
                    fraOutLine.Width = fraOutLine.Width + fraPickBill.Width
                    mbSizeExpand = True
                End If
                cmdAddNew.Enabled = True
                cmdPick.Enabled = False
                cmdSave.Enabled = True
                cmdCancel.Enabled = True
                cmdAccept.Enabled = False
                lblPickOperator.Visible = True
                cboPickOperator.Visible = True
            Case EPS_Cancel
                '判断是在受理时废的，还是已经提件了以后再废的
                If Trim(m_oPackage.SheetID) = "" Then
                    If mbSizeExpand Then
                        fraPickBill.Visible = False
                        fraOutLine.Width = fraOutLine.Width - fraPickBill.Width
                        mbSizeExpand = False
                        lblPickOperator.Visible = False
                        cboPickOperator.Visible = False
                    End If
                Else
                    If Not mbSizeExpand Then
                        fraPickBill.Visible = True
                        fraOutLine.Width = fraOutLine.Width + fraPickBill.Width
                        mbSizeExpand = True
                        lblPickOperator.Visible = True
                        cboPickOperator.Visible = True
                    End If
                End If
                cmdAddNew.Enabled = True
                cmdPick.Enabled = False
                cmdSave.Enabled = False
                cmdCancel.Enabled = False
                cmdAccept.Enabled = False
        End Select
    End Select
    cmdSMS.Enabled = g_bSMSValid
End Sub
'刷新界面
Public Sub RefreshForm()
On Error GoTo ErrHandle
  
  txtNums.Enabled = True
  cboWeight.Enabled = True
  txtTransCharge.Enabled = True
  vsPriceItem.Enabled = True
  
  Select Case Status
    Case EFormStatus.EFS_AddNew
        m_lPackageID = 0
            
        cboPackageName.Text = ""
        cboPack.Text = ""
        cboAreaType.Text = ""
        txtStartStation.Text = ""
        dtpArriveTime.Value = Format(Now, "yyyy-MM-dd HH:mm")
        txtLicense.Text = ""
        txtNums.Text = ""
        cboWeight.Text = ""
        cboPickType.ListIndex = 0
        txtShipper.Text = ""
        txtShipperPhone.Text = ""
        txtShipperUnit.Text = ""
        txtPicker.Text = ""
        txtPickerAddress.Text = ""
        txtPickerPhone.Text = ""
        txtPickerUnit.Text = ""
        txtSMS.Text = ""
        cboLoader.Text = ""
        cboSavePosition.Text = ""
        txtTransCharge.Text = ""
        txtRemark.Text = ""
        cboOperator.Text = g_oActUser.userName
        
        lblSheetID.Caption = ""
        vsPriceItem.TextMatrix(0, 1) = ""
        vsPriceItem.TextMatrix(1, 1) = ""
        vsPriceItem.TextMatrix(2, 1) = ""
        vsPriceItem.TextMatrix(3, 1) = ""
        vsPriceItem.TextMatrix(4, 1) = ""
        txtPickerCreditID.Text = ""
        txtReceivedMoney.Text = "0.0"
        
        
        lblStatus.ForeColor = vbBlue
        lblStatus.Caption = CPick_Normal
        cboPickOperator.Text = ""
        
        txtDrawer.Text = ""
        txtDrawerPhone.Text = cszDrawerPhone
        
'        txtPackageID.Enabled = True
        txtPackageID.Text = ""
        
        
    Case EFormStatus.EFS_Modify
        m_oPackage.Identify m_lPackageID
        
        txtPackageID.Text = UnBuildPacketID(m_lPackageID) 'm_oPackage.PackageID
        txtPackageID.Enabled = False
        cboPackageName.Text = m_oPackage.PackageName
        cboPack.Text = m_oPackage.PackType
        cboAreaType.Text = m_oPackage.AreaType
        txtStartStation.Text = m_oPackage.StartStationName
        dtpArriveTime.Value = m_oPackage.ArrivedTime
        txtLicense.Text = m_oPackage.LicenseTagNo
        txtNums.Text = m_oPackage.PackageNumber
        cboWeight.Text = m_oPackage.CalWeight
        cboPickType.ListIndex = SeekListIndex(cboPickType, m_oPackage.PickType)
        txtShipper.Text = m_oPackage.Shipper
        txtShipperPhone.Text = m_oPackage.ShipperPhone
        txtShipperUnit.Text = m_oPackage.ShipperUnit
        txtPicker.Text = m_oPackage.Picker
        txtPickerAddress.Text = m_oPackage.PickerAddress
        txtPickerPhone.Text = m_oPackage.PickerPhone
        txtPickerUnit.Text = m_oPackage.PickerUnit
        txtSMS.Text = ""
        cboLoader.Text = m_oPackage.Loader
        cboSavePosition.Text = m_oPackage.SavePosition
        txtTransCharge.Text = m_oPackage.TransitCharge
        txtRemark.Text = m_oPackage.Remark
        cboOperator.Text = m_oPackage.Operator
        
        '根据单据的状态设置提货信息区域是显示还是隐藏
        Select Case m_oPackage.Status
            Case EPS_Normal
            
                lblStatus.ForeColor = vbBlue
                lblStatus.Caption = CPick_Normal
                cboPickOperator.Text = ""
            
            Case EPS_Picked
            
                lblStatus.ForeColor = vbRed
                lblStatus.Caption = CPick_Picked
            
                lblSheetID.Caption = m_oPackage.SheetID
                vsPriceItem.TextMatrix(0, 1) = m_oPackage.LoadCharge
                vsPriceItem.TextMatrix(1, 1) = m_oPackage.KeepCharge
                vsPriceItem.TextMatrix(2, 1) = m_oPackage.SendCharge
                vsPriceItem.TextMatrix(3, 1) = m_oPackage.MoveCharge
                vsPriceItem.TextMatrix(4, 1) = m_oPackage.OtherCharge
                txtPickerCreditID.Text = m_oPackage.PickerCreditID
                dtpPickTime.Value = m_oPackage.PickTime
                cboPickOperator.Text = m_oPackage.PickOperator
                txtDrawer.Text = m_oPackage.Drawer
                txtDrawerPhone.Text = m_oPackage.DrawerPhone
                
                txtNums.Enabled = False
                cboWeight.Enabled = False
                txtTransCharge.Enabled = False
                vsPriceItem.Enabled = False
                txtDrawer.Enabled = False
                dtpPickTime.Enabled = False
                cboAreaType.Enabled = False
                txtStartStation.Enabled = False
                cboOperator.Enabled = False
                cboPickOperator.Enabled = False
                dtpArriveTime.Enabled = False
                
            Case EPS_Cancel
                '判断是在受理时废的，还是已经提件了以后再废的
                If Trim(m_oPackage.SheetID) = "" Then
                cboPickOperator.Text = ""

                Else

                    lblSheetID.Caption = m_oPackage.SheetID
                    vsPriceItem.TextMatrix(0, 1) = m_oPackage.LoadCharge
                    vsPriceItem.TextMatrix(1, 1) = m_oPackage.KeepCharge
                    vsPriceItem.TextMatrix(2, 1) = m_oPackage.SendCharge
                    vsPriceItem.TextMatrix(3, 1) = m_oPackage.MoveCharge
                    vsPriceItem.TextMatrix(4, 1) = m_oPackage.OtherCharge
                    txtPickerCreditID.Text = m_oPackage.PickerCreditID
                    cboPickOperator.Text = m_oPackage.UserID
                    txtDrawer.Text = m_oPackage.Drawer
                    txtDrawerPhone.Text = m_oPackage.DrawerPhone
                End If
                lblStatus.ForeColor = &H808080
                lblStatus.Caption = CPick_Canceled
            
        End Select
    
    
    
    End Select
    LayoutForm
    
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub Form_Paint()
    RefreshCurrentSheetID
End Sub

Private Sub Form_Resize()
    If mdiMain.ActiveForm Is Me Then
        If Not Me.WindowState = vbMaximized Then Me.WindowState = vbMaximized
        fraOutLine.Left = (Me.ScaleWidth - fraOutLine.Width) / 2
        fraOutLine.Top = (Me.ScaleHeight - fraOutLine.Height) / 2
    End If
End Sub


'当某一个常用定义combox输入了新的条目后，可利用本函数将新条目加入至combox中
Private Sub SetBaseDefine(cboObject As ComboBox, peDefineType As EDefineType)
On Error GoTo ErrHandle
'利comboBox的tag存放常用定义类型
    Dim szText As String
    szText = Trim(cboObject.Text)
    If szText = "" Then Exit Sub
    
    Dim i As Integer
    For i = 1 To cboObject.ListCount
        If szText = cboObject.List(i - 1) Then
            Exit Sub
        End If
    Next i
    
    Dim nResult As Integer
    nResult = MsgBox("该条目不存在，是否添加该条目？", vbYesNo + vbQuestion + vbDefaultButton2)
    If nResult = vbNo Then Exit Sub
    
    g_oPackageParam.AddBaseDefine peDefineType, szText, ""
    cboObject.AddItem szText
    
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseSMS
End Sub

Private Sub txtDrawer_GotFocus()
    txtDrawer.SelStart = 0
    txtDrawer.SelLength = Len(txtDrawer.Text)
End Sub

Private Sub txtDrawerPhone_GotFocus()
    txtDrawerPhone.SelStart = 0
    txtDrawerPhone.SelLength = Len(txtDrawerPhone.Text)
End Sub

Private Sub txtLicense_Change()
    FormatTextBoxBySize txtLicense, 10
End Sub

Private Sub txtNums_Change()
    FormatTextToNumeric txtNums, False, False
'    SynSMSString
End Sub

Private Sub txtNums_GotFocus()
    With txtNums
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtNums_LostFocus()
     If mbSizeExpand And Val(txtNums) <> m_oPackage.PackageNumber Then MsgBox "件数被更改，请先按[保存]按钮进行保存，再提货！", vbInformation
End Sub

Private Sub txtOwnID_GotFocus()
    With txtOwnID
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtPackageID_Change()
    FormatTextToNumeric txtPackageID, False, False
    FormatTextBoxBySize txtPackageID, 6
'    SynSMSString
End Sub

Private Sub txtPackageID_GotFocus()
    With txtPackageID
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub


Private Sub txtPicker_Change()
'    SynSMSString
    
End Sub

Private Sub txtPickerCreditID_Change()
    FormatTextBoxBySize txtPickerCreditID, 18
End Sub

Private Sub txtPickerCreditID_LostFocus()
    If txtPickerCreditID.Text <> "" And Len(txtPickerCreditID.Text) < 5 Then
        ShowMsg "身份证号码格式不对，请重输！"
        txtPickerCreditID.SetFocus
        txtPickerCreditID.SelStart = 0
        txtPickerCreditID.SelLength = Len(txtPickerCreditID.Text)
    End If
End Sub

Private Sub txtPickerUnit_Change()
'    SynSMSString

End Sub

Private Sub txtShipper_Change()
'    SynSMSString

End Sub

Private Sub txtSMS_Click()
    SynSMSString

End Sub

Private Sub txtSMS_LostFocus()
'    SynSMSString

End Sub

Private Sub txtStartStation_Change()
'    SynSMSString

End Sub

Private Sub txtTransCharge_Change()
    FormatTextToNumeric txtTransCharge, False, False
    lblTransCharge.Caption = txtTransCharge.Text
End Sub
Private Sub txtPrice_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = vbKeyBack Or KeyAscii = 46) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtReceivedMoney_Change()
    If txtReceivedMoney.Text <> "" Then
        flblRestMoney.Caption = Str(Round(Val(txtReceivedMoney.Text) - Val(flblTotalPrice.Caption) - Val(lblTransCharge.Caption), 2))
    End If
End Sub

Private Sub txtReceivedMoney_GotFocus()
    lblReceivedMoney.ForeColor = cnColor_Active
    txtReceivedMoney.Text = ""
End Sub

Private Sub txtReceivedMoney_KeyDown(KeyCode As Integer, Shift As Integer)
'    Select Case KeyCode
'    Case vbKeyReturn
'        KeyCode = 0
'        If cmdAccept.Enabled = True Then
'            cmdAccept.SetFocus
'        End If
'    End Select
End Sub

Private Sub txtReceivedMoney_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = vbKeyBack Or KeyAscii = 46) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtReceivedMoney_LostFocus()
    lblReceivedMoney.ForeColor = 0
End Sub

Private Sub txtTransCharge_GotFocus()
    With txtTransCharge
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub vsPriceItem_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    On Error GoTo ErrHandle:
    Dim i As Integer

    CalSumPrice
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub vsPriceItem_CellChanged(ByVal Row As Long, ByVal Col As Long)
    On Error GoTo ErrHandle:
    Dim i As Integer

    CalSumPrice
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub vsPriceItem_GotFocus()
    lblPriceItem.ForeColor = cnColor_Active
    vsPriceItem_SelChange
End Sub

Private Sub vsPriceItem_KeyDown(KeyCode As Integer, Shift As Integer)
'    Select Case KeyCode
'    Case vbKeyReturn
'        KeyCode = 0
'        txtPickerCreditID.SetFocus
'    End Select
End Sub



Private Sub vsPriceItem_LostFocus()
    lblPriceItem.ForeColor = 0
End Sub

Private Sub vsPriceItem_SelChange()
    With vsPriceItem
        If (.Row = 0 Or .Row = 1 Or .Row = 2 Or .Row = 3 Or .Row = 4) And .Col = 1 Then
            '如果为运费项
            .Editable = flexEDKbdMouse
        Else
            .Editable = flexEDNone
        End If
    End With
End Sub
Private Sub CalSumPrice()
    '算出总价 vsPriceItem.subTotal
    Dim i As Integer
    Dim Num As Integer
    Dim Sum As Double
    Num = vsPriceItem.Rows
    For i = 1 To Num
        Sum = Sum + Val(vsPriceItem.TextMatrix(i - 1, 1))
    Next i
'    Sum = Sum + Val(txtTransCharge.Text)
    flblTotalPrice.Caption = Sum
    
End Sub

'同步要发送的短信息内容
Private Sub SynSMSString()
    Dim szTmp As String
    szTmp = txtSMS.Text
    szTmp = Replace(szTmp, cSMSDefine_PackageID, txtPackageID.Text)
    szTmp = Replace(szTmp, cSMSDefine_PackageName, cboPackageName.Text)
    szTmp = Replace(szTmp, cSMSDefine_Picker, txtPicker.Text)
    szTmp = Replace(szTmp, cSMSDefine_ShartStation, IIf(Trim(txtStartStation.Text) = "", cboAreaType.Text, txtStartStation.Text))
    szTmp = Replace(szTmp, cSMSDefine_Shipper, txtShipper.Text)
    szTmp = Replace(szTmp, cSMSDefine_PackageNumber, txtNums.Text)
    txtSMS.Text = szTmp
End Sub


