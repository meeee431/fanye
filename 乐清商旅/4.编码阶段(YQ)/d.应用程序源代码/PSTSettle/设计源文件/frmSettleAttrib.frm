VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmSettleAttrib 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "属性"
   ClientHeight    =   6930
   ClientLeft      =   3375
   ClientTop       =   2460
   ClientWidth     =   9150
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   9150
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox ptExtra 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4380
      Left            =   435
      ScaleHeight     =   4380
      ScaleWidth      =   8235
      TabIndex        =   42
      Top             =   1305
      Width           =   8235
      Begin VSFlex7LCtl.VSFlexGrid vsExtra 
         Height          =   4050
         Left            =   180
         TabIndex        =   43
         Top             =   135
         Width           =   7890
         _cx             =   13917
         _cy             =   7144
         _ConvInfo       =   -1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
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
   End
   Begin RTComctl3.CoolButton cmdShowCheckSheet 
      Height          =   345
      Left            =   3690
      TabIndex        =   9
      Top             =   6480
      Visible         =   0   'False
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "路单(&S)"
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
      MICON           =   "frmSettleAttrib.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdStation 
      Height          =   345
      Left            =   4800
      TabIndex        =   8
      Top             =   6480
      Visible         =   0   'False
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "站点(&T)"
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
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmSettleAttrib.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdTotalFixFee 
      Height          =   345
      Left            =   225
      TabIndex        =   7
      Top             =   6480
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "汇总应扣款(&V)"
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
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmSettleAttrib.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   9165
      TabIndex        =   0
      Top             =   0
      Width           =   9165
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Height          =   60
         Left            =   -60
         TabIndex        =   4
         Top             =   690
         Width           =   9315
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "作废"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   6390
         TabIndex        =   3
         Top             =   270
         Width           =   450
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "路单结算属性"
         Height          =   180
         Left            =   330
         TabIndex        =   2
         Top             =   300
         Width           =   1080
      End
      Begin VB.Image Image1 
         Height          =   855
         Left            =   3225
         Picture         =   "frmSettleAttrib.frx":0054
         Top             =   0
         Width           =   5925
      End
   End
   Begin RTComctl3.CoolButton cmdViewSettleSheet 
      Height          =   345
      Left            =   1950
      TabIndex        =   5
      Top             =   6480
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "结算单(&V)"
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
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmSettleAttrib.frx":153E
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
      Height          =   330
      Left            =   7410
      TabIndex        =   6
      Top             =   6480
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   582
      BTYPE           =   3
      TX              =   "关闭(&E)"
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
      MICON           =   "frmSettleAttrib.frx":155A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   " "
      Height          =   945
      Left            =   -180
      TabIndex        =   1
      Top             =   6210
      Width           =   9690
   End
   Begin VB.PictureBox ptCheckSheet 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   4380
      Left            =   435
      ScaleHeight     =   4380
      ScaleWidth      =   8235
      TabIndex        =   30
      Top             =   1305
      Width           =   8235
      Begin MSComctlLib.ListView lvCheckSheet 
         Height          =   4200
         Left            =   45
         TabIndex        =   31
         Top             =   60
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   7408
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
   End
   Begin VB.PictureBox ptStation 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   4380
      Left            =   435
      ScaleHeight     =   4380
      ScaleWidth      =   8130
      TabIndex        =   28
      Top             =   1305
      Width           =   8130
      Begin VSFlex7LCtl.VSFlexGrid vsStation 
         Height          =   4335
         Left            =   75
         TabIndex        =   29
         Top             =   105
         Width           =   8115
         _cx             =   14314
         _cy             =   7646
         _ConvInfo       =   -1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
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
   End
   Begin MSComctlLib.TabStrip tsInfo 
      Height          =   5175
      Left            =   315
      TabIndex        =   10
      Top             =   900
      Width           =   8490
      _ExtentX        =   14975
      _ExtentY        =   9128
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "基本信息"
            Key             =   "baseinfo"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   " 结算项 "
            Key             =   "splititem"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "  站点  "
            Key             =   "station"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "  路单  "
            Key             =   "checksheet"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "补票人数"
            Key             =   "extra"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   420
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettleAttrib.frx":1576
            Key             =   "checksheet"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox ptBaseInfo 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   4380
      Left            =   465
      ScaleHeight     =   4380
      ScaleWidth      =   8130
      TabIndex        =   11
      Top             =   1320
      Width           =   8130
      Begin VB.TextBox txtChecker 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   6180
         TabIndex        =   40
         Text            =   "999"
         Top             =   255
         Width           =   1200
      End
      Begin VB.TextBox txtSettleDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   300
         Left            =   5220
         TabIndex        =   41
         Text            =   "100"
         Top             =   2160
         Width           =   2730
      End
      Begin VB.TextBox txtExecutive 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   300
         Left            =   1290
         TabIndex        =   39
         Text            =   "888"
         Top             =   2130
         Width           =   1425
      End
      Begin VB.TextBox txtSettleStationPrice 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   300
         Left            =   1725
         TabIndex        =   37
         Text            =   "666"
         Top             =   1650
         Width           =   2310
      End
      Begin VB.TextBox txtSettlePrice 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   300
         Left            =   5385
         TabIndex        =   36
         Text            =   "555"
         Top             =   1170
         Width           =   2385
      End
      Begin VB.TextBox txtTotalTicketPrice 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   300
         Left            =   1290
         TabIndex        =   35
         Text            =   "444"
         Top             =   1139
         Width           =   2655
      End
      Begin VB.TextBox txtSheetCount 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   300
         Left            =   5475
         TabIndex        =   34
         Text            =   "333"
         Top             =   675
         Width           =   2400
      End
      Begin VB.TextBox txtFinDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   300
         Left            =   1410
         TabIndex        =   33
         Text            =   "222"
         Top             =   705
         Width           =   2520
      End
      Begin VB.TextBox txtSettleSheetID 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   330
         Left            =   1605
         TabIndex        =   32
         Text            =   "111"
         Top             =   240
         Width           =   3315
      End
      Begin VB.TextBox txtAnnotation 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   1545
         Left            =   1200
         TabIndex        =   12
         Top             =   2655
         Width           =   6825
      End
      Begin VB.TextBox txtQuautity 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   300
         Left            =   5250
         TabIndex        =   38
         Text            =   "77"
         Top             =   1665
         Width           =   2565
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "复核人:"
         Height          =   180
         Left            =   5385
         TabIndex        =   14
         Top             =   270
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "备注:"
         Height          =   180
         Left            =   495
         TabIndex        =   23
         Top             =   2640
         Width           =   450
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "总票款:"
         Height          =   180
         Left            =   465
         TabIndex        =   22
         Top             =   1199
         Width           =   630
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "应结票款:"
         Height          =   180
         Left            =   4200
         TabIndex        =   21
         Top             =   1199
         Width           =   810
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "结算单代码:"
         Height          =   180
         Index           =   0
         Left            =   465
         TabIndex        =   20
         Top             =   255
         Width           =   990
      End
      Begin VB.Label lblStateChange 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "总路单数:"
         Height          =   180
         Left            =   4200
         TabIndex        =   19
         Top             =   690
         Width           =   810
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "总人数:"
         Height          =   180
         Left            =   4200
         TabIndex        =   18
         Top             =   1671
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "结给车站票款:"
         Height          =   180
         Left            =   465
         TabIndex        =   17
         Top             =   1671
         Width           =   1140
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "结算日期:"
         Height          =   180
         Left            =   465
         TabIndex        =   16
         Top             =   727
         Width           =   810
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "拆帐人:"
         Height          =   180
         Left            =   465
         TabIndex        =   15
         Top             =   2145
         Width           =   630
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "拆帐日期:"
         Height          =   180
         Left            =   4185
         TabIndex        =   13
         Top             =   2190
         Width           =   810
      End
   End
   Begin VB.PictureBox ptSplitItem 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   4380
      Left            =   435
      ScaleHeight     =   4380
      ScaleWidth      =   8130
      TabIndex        =   24
      Top             =   1305
      Width           =   8130
      Begin MSComctlLib.ListView lvVehicle 
         Height          =   1305
         Left            =   150
         TabIndex        =   25
         Top             =   1425
         Width           =   7905
         _ExtentX        =   13944
         _ExtentY        =   2302
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
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lvCompany 
         Height          =   1290
         Left            =   150
         TabIndex        =   26
         Top             =   60
         Width           =   7905
         _ExtentX        =   13944
         _ExtentY        =   2275
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lvBus 
         Height          =   1590
         Left            =   150
         TabIndex        =   27
         Top             =   2820
         Width           =   7905
         _ExtentX        =   13944
         _ExtentY        =   2805
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
End
Attribute VB_Name = "frmSettleAttrib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'此窗口内的代码,均为垃圾代码,全部需要重新写过
'2005-07-22陈峰注



Public m_szSettleSheetID As String

Dim m_atSplitItem() As TSplitItemInfo
Dim m_nSplitItemCount As Integer

Dim m_oSettleSheet As New SettleSheet





Private Sub cmdCancel_Click()
    Unload Me
End Sub





Private Sub cmdShowCheckSheet_Click()
    frmShowCheckSheet.m_szSettleSheetID = txtSettleSheetID.Text
    frmShowCheckSheet.ZOrder 0
    frmShowCheckSheet.Show vbModal
End Sub



Private Sub cmdStation_Click()
    frmSettleStationInfo.m_szSheetID = txtSettleSheetID.Text
    frmSettleStationInfo.Show vbModal
End Sub

Private Sub cmdTotalFixFee_Click()
    '汇总应扣款
    Dim oSplit As New Split
    On Error GoTo ErrorHandle
    oSplit.Init g_oActiveUser
    oSplit.ReTotalVehicleFixFee Trim(txtSettleSheetID.Text)
    FillFormInfo
    Exit Sub
ErrorHandle:
    ShowErrorMsg
    
End Sub

Private Sub cmdViewSettleSheet_Click()
    '打印结算单
    frmPrintFinSheet.m_SheetID = Trim(txtSettleSheetID.Text)
    frmPrintFinSheet.m_szLugSettleSheetID = ""
    frmPrintFinSheet.m_bRePrint = False
'    frmPrintFinSheet.m_bNeedPrint = False
    
    
    frmPrintFinSheet.ZOrder 0
    frmPrintFinSheet.Show vbModal
End Sub



Private Sub Form_Load()
    AlignFormPos Me
    
    FillFormInfo
    AlignHeadWidth Me.name, lvCompany
    AlignHeadWidth Me.name, lvVehicle
    AlignHeadWidth Me.name, lvBus
    
    
    tsInfo_Click
End Sub

Private Sub FillFormInfo()
On Error GoTo err:
    Dim i As Integer
    m_oSettleSheet.Init g_oActiveUser
    m_oSettleSheet.Identify m_szSettleSheetID
    
    txtSettleSheetID.Text = m_oSettleSheet.SettleSheetID
    txtFinDate.Text = Format(m_oSettleSheet.SettleStartDate, "yyyy-mm-dd") & " 至 " & Format(m_oSettleSheet.SettleEndDate, "yyyy-mm-dd")
    txtSheetCount.Text = m_oSettleSheet.CheckSheetCount
    
    txtTotalTicketPrice.Text = m_oSettleSheet.TotalTicketPrice
    txtSettlePrice.Text = FormatMoney(m_oSettleSheet.SettleOtherCompanyPrice - m_oSettleSheet.SettleStationPrice)  '应结票款
    txtSettleStationPrice.Text = m_oSettleSheet.SettleStationPrice
    txtQuautity.Text = m_oSettleSheet.TotalQuantity
    
    txtExecutive.Text = m_oSettleSheet.Settler
    txtChecker.Text = m_oSettleSheet.Checker
    txtSettleDate.Text = m_oSettleSheet.SettleDate
    If m_oSettleSheet.Status = CS_SettleSheetInvalid Then
        lblStatus.ForeColor = vbRed
        cmdShowCheckSheet.Enabled = False
    Else
        lblStatus.ForeColor = 0
        cmdShowCheckSheet.Enabled = True
    End If
    lblStatus.Caption = GetSettleSheetStatusString(m_oSettleSheet.Status)    '转换
    txtAnnotation.Text = m_oSettleSheet.Annotation
'    txtOwner.text = m_oSettleSheet.OwnerID
    '填充列表
    FillList
    
    
Exit Sub
err:
    ShowErrorMsg
End Sub
Private Sub FillList()
On Error GoTo here
Dim m_oReport As New Report
Dim TSplitResult As TSplitResult
'Dim Count As Integer
Dim i As Integer
Dim j As Integer
    With lvCompany.ColumnHeaders
        .Clear
        .Add , , "参运公司"
        .Add , , "协议"
        .Add , , "应结票款"
        .Add , , "结给车站"
        .Add , , "人数"
        .Add , , "人公里"
    End With
    
    With lvVehicle.ColumnHeaders
        .Clear
        .Add , , "车辆"
        .Add , , "协议"
        .Add , , "应结票款"
        .Add , , "结给车站"
        .Add , , "人数"
        .Add , , "人公里"
    End With
    
    With lvBus.ColumnHeaders
        .Clear
        .Add , , "车次"
        .Add , , "协议"
        .Add , , "应结票款"
        .Add , , "结给车站"
        .Add , , "人数"
        .Add , , "人公里"
    End With
    '取得使用的拆算款项
    m_oReport.Init g_oActiveUser
    m_atSplitItem = m_oReport.GetSplitItemInfo(, True)
    m_nSplitItemCount = ArrayLength(m_atSplitItem)
    If m_nSplitItemCount = 0 Then Exit Sub
    For i = 1 To m_nSplitItemCount
'        If m_atSplitItem(i).SplitItemName <> "" Then
            lvCompany.ColumnHeaders.Add , , m_atSplitItem(i).SplitItemName '
            lvVehicle.ColumnHeaders.Add , , m_atSplitItem(i).SplitItemName '同时增加车辆列表的拆算项  "k" & Val(m_atSplitItem(i).SplitItemID)
            lvBus.ColumnHeaders.Add , , m_atSplitItem(i).SplitItemName
'        End If
    Next i
        
    '填充列表内容
    Dim lvItem As ListItem
    Dim rsTemp As Recordset
    Dim lvItem1 As ListItem
    Dim rsTemp1 As Recordset
    Dim lvItem2 As ListItem
    Dim rsTemp2 As Recordset
    
    
    Set rsTemp = m_oReport.GetSettleCompanyLst(m_szSettleSheetID)
    lvCompany.ListItems.Clear
    If rsTemp.RecordCount > 0 Then
        For j = 1 To rsTemp.RecordCount
            Set lvItem = lvCompany.ListItems.Add(, , FormatDbValue(rsTemp!transport_company_short_name))
                lvItem.SubItems(1) = FormatDbValue(rsTemp!protocol_name)
                lvItem.SubItems(2) = FormatMoney(FormatDbValue(rsTemp!settle_price) - FormatDbValue(rsTemp!settle_station_price))
                lvItem.SubItems(3) = FormatDbValue(rsTemp!settle_station_price)
                lvItem.SubItems(4) = FormatDbValue(rsTemp!passenger_number)
                lvItem.SubItems(5) = FormatDbValue(rsTemp!Mileage)
                
                For i = 1 To m_nSplitItemCount
                    lvItem.SubItems(i + 5) = rsTemp("split_item_" & Val(m_atSplitItem(i).SplitItemID)).Value  '拆算项处理
                Next i
              rsTemp.MoveNext
        Next j
    End If
    
    
    Set rsTemp1 = m_oReport.GetSettleVehicleLst(m_szSettleSheetID)
    lvVehicle.ListItems.Clear
    If rsTemp1.RecordCount > 0 Then
        For j = 1 To rsTemp1.RecordCount
            Set lvItem1 = lvVehicle.ListItems.Add(, , FormatDbValue(rsTemp1!object_name))
            lvItem1.SubItems(1) = FormatDbValue(rsTemp1!protocol_name)
            lvItem1.SubItems(2) = FormatMoney(FormatDbValue(rsTemp1!settle_price) - FormatDbValue(rsTemp1!settle_station_price))
            lvItem1.SubItems(3) = FormatDbValue(rsTemp1!settle_station_price)
            lvItem1.SubItems(4) = FormatDbValue(rsTemp1!passenger_number)
            lvItem1.SubItems(5) = FormatDbValue(rsTemp1!Mileage)


            For i = 1 To m_nSplitItemCount
                lvItem1.SubItems(i + 5) = rsTemp1("split_item_" & Val(m_atSplitItem(i).SplitItemID)).Value
            Next i
             rsTemp1.MoveNext
        Next j
    End If
    
    Set rsTemp2 = m_oReport.GetSettleBusLstSimple(m_szSettleSheetID)
    lvBus.ListItems.Clear
    If rsTemp2.RecordCount > 0 Then
        For j = 1 To rsTemp2.RecordCount
            Set lvItem2 = lvBus.ListItems.Add(, , FormatDbValue(rsTemp2!bus_id))
            lvItem2.SubItems(1) = FormatDbValue(rsTemp2!protocol_name)
            lvItem2.SubItems(2) = FormatMoney(FormatDbValue(rsTemp2!settle_price) - FormatDbValue(rsTemp2!settle_station_price))
            lvItem2.SubItems(3) = FormatDbValue(rsTemp2!settle_station_price)
            lvItem2.SubItems(4) = FormatDbValue(rsTemp2!passenger_number)
            lvItem2.SubItems(5) = FormatDbValue(rsTemp2!Mileage)


            For i = 1 To m_nSplitItemCount
                lvItem2.SubItems(i + 5) = rsTemp2("split_item_" & Val(m_atSplitItem(i).SplitItemID)).Value
            Next i
             rsTemp2.MoveNext
        Next j
    End If
Exit Sub
here:
    ShowErrorMsg

End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveFormPos Me
    
    SaveHeadWidth Me.name, lvCompany
    SaveHeadWidth Me.name, lvVehicle
    SaveHeadWidth Me.name, lvBus
'    SaveHeadWidth Me.name, lvCheckSheet
    Unload Me
End Sub



Private Sub lvCheckSheet_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
 If lvCheckSheet.SortOrder = lvwAscending Then
    lvCheckSheet.SortOrder = lvwDescending
 Else
    lvCheckSheet.SortOrder = lvwAscending
 End If
    lvCheckSheet.SortKey = ColumnHeader.Index - 1
    lvCheckSheet.Sorted = True
End Sub
Private Sub lvCheckSheet_DblClick()
    
    Dim oCommDialog As New STShell.CommDialog
    On Error GoTo here
    oCommDialog.Init g_oActiveUser
    oCommDialog.ShowCheckSheet lvCheckSheet.SelectedItem.Text
    Set oCommDialog = Nothing
    Exit Sub
here:
    ShowErrorMsg
End Sub




Private Sub lvBus_DblClick()
    ModifyBusSplitItem
End Sub

Private Sub lvCompany_DblClick()
    ModifyCompanySplitItem
End Sub


Private Sub lvVehicle_DblClick()
    ModifyVehicleSplitItem
End Sub


Private Sub UpdateTotalPrice()
    
    Dim dbSettleOtherCompanyPrice As Double
    Dim dbSettleStationPrice As Double
    Dim i As Integer
    
    dbSettleOtherCompanyPrice = 0
    dbSettleStationPrice = 0
    For i = 1 To lvCompany.ListItems.Count
        dbSettleOtherCompanyPrice = dbSettleOtherCompanyPrice + lvCompany.SelectedItem.SubItems(2)
        dbSettleStationPrice = dbSettleOtherCompanyPrice + lvCompany.SelectedItem.SubItems(3)
    Next i
    For i = 1 To lvVehicle.ListItems.Count
        dbSettleOtherCompanyPrice = dbSettleOtherCompanyPrice + lvVehicle.SelectedItem.SubItems(2)
        dbSettleStationPrice = dbSettleOtherCompanyPrice + lvVehicle.SelectedItem.SubItems(3)
    Next i
    For i = 1 To lvBus.ListItems.Count
        dbSettleOtherCompanyPrice = dbSettleOtherCompanyPrice + lvBus.SelectedItem.SubItems(2)
        dbSettleStationPrice = dbSettleOtherCompanyPrice + lvBus.SelectedItem.SubItems(3)
    Next i
    txtSettlePrice.Text = dbSettleOtherCompanyPrice - dbSettleStationPrice
    txtSettleStationPrice.Text = dbSettleStationPrice
End Sub

Private Sub tsInfo_Click()
    ptBaseInfo.Visible = False
    ptCheckSheet.Visible = False
    ptStation.Visible = False
    ptSplitItem.Visible = False
    
    If tsInfo.Tabs(1).Selected Then
        ptBaseInfo.ZOrder 0
        ptBaseInfo.Visible = True
        
    ElseIf tsInfo.Tabs(2).Selected Then
        ptSplitItem.ZOrder 0
        ptSplitItem.Visible = True
    ElseIf tsInfo.Tabs(3).Selected Then
        ptStation.ZOrder 0
        ptStation.Visible = True
        FillStation
    ElseIf tsInfo.Tabs(4).Selected Then
        ptCheckSheet.ZOrder 0
        ptCheckSheet.Visible = True
        FillCheckSheet
    ElseIf tsInfo.Tabs(5).Selected Then
        '显示手工补票信息
        FillExtra
    End If
End Sub



'填充结算单的站点信息
Private Sub FillStation()
    
    '站点人数的列位置
    Const VI_Route = 1
    Const VI_Station = 2
    Const VI_TicketType = 3
    Const VI_Quantity = 4


    vsStation.Clear
    
    vsStation.Cols = 5 '12
    vsStation.FixedCols = 1
    vsStation.Rows = 2
    vsStation.FixedRows = 1
    
    
'    AlignHeadWidth Me.name, lvCheckSheet

    With vsStation
        .TextMatrix(0, VI_Route) = "线路"
        .TextMatrix(0, VI_Station) = "站点"
        .TextMatrix(0, VI_TicketType) = "票种"
        .TextMatrix(0, VI_Quantity) = "人数"
    
    End With
    
    
    Dim rsTemp As Recordset
    Dim oReport As New Report
    Dim i As Integer
    On Error GoTo ErrorHandle
    oReport.Init g_oActiveUser
    
    Set rsTemp = oReport.GetSettleRouteQuantity(txtSettleSheetID.Text)
    If rsTemp.RecordCount = 0 Then Exit Sub
    With vsStation
        .Rows = rsTemp.RecordCount + 1
        For i = 1 To rsTemp.RecordCount
'            .TextMatrix(i, VI_SellStation) = FormatDbValue(rsTemp!sell_station_name)
            .TextMatrix(i, VI_Route) = FormatDbValue(rsTemp!route_name)
'            .TextMatrix(i, VI_VehicleType) = FormatDbValue(rsTemp!vehicle_type_name)
            .TextMatrix(i, VI_Station) = FormatDbValue(rsTemp!station_name)
'            .TextMatrix(i, vi_bus) = FormatDbValue(rsTemp!bus_id)
            .TextMatrix(i, VI_TicketType) = FormatDbValue(rsTemp!ticket_type_name)
            .TextMatrix(i, VI_Quantity) = FormatDbValue(rsTemp!Quantity)
'            .TextMatrix(i, VI_PassCharge) = FormatDbValue(rsTemp!pass_charge)
'            .TextMatrix(i, VI_SettlePrice) = FormatDbValue(rsTemp!settle_price)
'            .TextMatrix(i, VI_HalvePrice) = FormatDbValue(rsTemp!halve_price)
'            .TextMatrix(i, VI_ServicePrice) = FormatDbValue(rsTemp!service_price)
'            .TextMatrix(i, VI_SpringPrice) = FormatDbValue(rsTemp!spring_price)
            
            rsTemp.MoveNext
            
        Next i
    End With
    vsStation.MergeCells = flexMergeRestrictColumns
'    vsStation.MergeCol(VI_SellStation) = True
    vsStation.MergeCol(VI_Route) = True
'    vsStation.MergeCol(VI_VehicleType) = True
    vsStation.MergeCol(VI_Station) = True
    vsStation.MergeCol(VI_TicketType) = True

    vsStation.AllowUserResizing = flexResizeColumns
    
    
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

'填充路单信息
Private Sub FillCheckSheet()
    
    '路单信息的列位置
    Const PI_CheckSheetID = 0
    Const PI_BusDate = 1
    Const PI_BusID = 2
    Const PI_BusSerialNO = 3
    Const PI_LicenseTagNo = 4
    Const PI_CompanyName = 5
    Const PI_RouteID = 6
    Const PI_VehicleType = 7
    Const PI_Owner = 8
    Const PI_Checker = 9

    With lvCheckSheet.ColumnHeaders
        .Add , , "路单代码"
        .Add , , "日期"
        .Add , , "车次"
        .Add , , "序号"
        .Add , , "车辆"
        .Add , , "参运公司"
        .Add , , "线路"
        .Add , , "车型"
        .Add , , "车主"
        .Add , , "检票员"
    End With
On Error GoTo here
    Dim i As Integer
    Dim rsTemp As Recordset
    Dim lvItem As ListItem
    Dim m_oReport As New Report
    m_oReport.Init g_oActiveUser
    Set rsTemp = m_oReport.GetCheckSheetInfo(m_szSettleSheetID)
    If rsTemp.RecordCount = 0 Then Exit Sub
    For i = 1 To rsTemp.RecordCount
       Set lvItem = lvCheckSheet.ListItems.Add(, , FormatDbValue(rsTemp!check_sheet_id), , "checksheet")
       lvItem.SubItems(PI_BusDate) = Format(FormatDbValue(rsTemp!bus_date), "yyyy-MM-dd")
       lvItem.SubItems(PI_BusID) = FormatDbValue(rsTemp!bus_id)
       lvItem.SubItems(PI_BusSerialNO) = FormatDbValue(rsTemp!bus_serial_no)
       lvItem.SubItems(PI_LicenseTagNo) = FormatDbValue(rsTemp!license_tag_no)
       lvItem.SubItems(PI_CompanyName) = FormatDbValue(rsTemp!transport_company_short_name)
       lvItem.SubItems(PI_RouteID) = FormatDbValue(rsTemp!route_name)
       lvItem.SubItems(PI_VehicleType) = FormatDbValue(rsTemp!vehicle_type_name)
       lvItem.SubItems(PI_Owner) = FormatDbValue(rsTemp!owner_name)
       lvItem.SubItems(PI_Checker) = FormatDbValue(rsTemp!Checker)
       rsTemp.MoveNext
    Next i
    
    Exit Sub
here:
    ShowErrorMsg
End Sub

'填充手工补票人数信息
Private Sub FillExtra()
    '手工补票表格的列首位置
    Const VIE_NO = 0
    Const VIE_Quantity = 1
    Const VIE_TotalTicketPrice = 2
    Const VIE_Ratio = 3
    Const VIE_ServicePrice = 4
    Const VIE_SettleOutPrice = 5
    
    
    Dim i As Integer
    Dim rsTemp As Recordset
    Dim m_oReport As New Report
    
    On Error GoTo ErrorHandle
    
    '初始化手工补票的表格
    With vsExtra
        .Cols = 6
        .Rows = 1
        .Clear
        
        .TextMatrix(0, VIE_NO) = "序"
        .TextMatrix(0, VIE_Quantity) = "人数"
        .TextMatrix(0, VIE_TotalTicketPrice) = "票款"
        .TextMatrix(0, VIE_Ratio) = "劳务费率"
        .TextMatrix(0, VIE_ServicePrice) = "劳务费"
        .TextMatrix(0, VIE_SettleOutPrice) = "拆出金额"
        
        m_oReport.Init g_oActiveUser
        Set rsTemp = m_oReport.GetExtraInfo(m_szSettleSheetID)
        If rsTemp.RecordCount = 0 Then Exit Sub
        
        .Rows = rsTemp.RecordCount + 1
        
        '填充序
        For i = 1 To rsTemp.RecordCount
            .TextMatrix(i, VIE_NO) = i
            .TextMatrix(i, VIE_Quantity) = FormatDbValue(rsTemp!passenger_number)
            .TextMatrix(i, VIE_TotalTicketPrice) = FormatDbValue(rsTemp!total_ticket_price)
            .TextMatrix(i, VIE_Ratio) = FormatDbValue(rsTemp!Ratio)
            .TextMatrix(i, VIE_ServicePrice) = FormatDbValue(rsTemp!service_price)
            .TextMatrix(i, VIE_SettleOutPrice) = FormatDbValue(rsTemp!settle_out_price)
            rsTemp.MoveNext
        Next i
    End With
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub


Private Sub ModifyBusSplitItem()

    Dim i As Integer
    Dim nIndex As Integer
    Dim atBusInfo() As TBusSettle
    Dim atSplitItem() As TSplitItemInfo
    Dim oReport As New Report
    Dim dbSettlePrice As Double
    Dim dbSettleStationPrice As Double
    
    On Error GoTo ErrorHandle
    
    oReport.Init g_oActiveUser
    atSplitItem = oReport.GetSplitItemInfo()

    
    If lvBus.ListItems.Count = 0 Then Exit Sub
    '调用修改
    
    frmSplitItemModify.m_szObject = Trim(lvBus.SelectedItem.Text)
    frmSplitItemModify.m_szProtocol = Trim(lvBus.SelectedItem.SubItems(1))
    frmSplitItemModify.m_dbSettlePrice = Trim(lvBus.SelectedItem.SubItems(2))
    frmSplitItemModify.m_dbStationPrice = Trim(lvBus.SelectedItem.SubItems(3))
    frmSplitItemModify.m_nQuantity = Trim(lvBus.SelectedItem.SubItems(4))
    frmSplitItemModify.m_nType = 2
  
    frmSplitItemModify.m_eSettleObject = m_oSettleSheet.SettleObject
    frmSplitItemModify.m_szObjectID = m_oSettleSheet.ObjectID
    frmSplitItemModify.m_szSettleSheetID = m_oSettleSheet.SettleSheetID
  
    frmSplitItemModify.Show vbModal
    
    If frmSplitItemModify.m_bIsSave = False Then Exit Sub
    
    
    nIndex = lvBus.SelectedItem.Index

    dbSettlePrice = 0
    dbSettleStationPrice = 0
    For i = 1 To m_nSplitItemCount
        If lvBus.ColumnHeaders.Item(i + 6).Text = Trim(m_atSplitItem(i).SplitItemName) Then
            lvBus.SelectedItem.SubItems(i + 5) = m_adbSplitItem(Val(m_atSplitItem(i).SplitItemID)) ' rsTemp1("split_item_" & Val(m_atSplitItem(i).SplitItemID)).Value
            
            If m_atSplitItem(i).SplitType = CS_SplitOtherCompany Then
                dbSettlePrice = dbSettlePrice + lvBus.SelectedItem.SubItems(i + 5)
            ElseIf m_atSplitItem(i).SplitType = CS_SplitStation Then
                dbSettleStationPrice = dbSettleStationPrice + lvBus.SelectedItem.SubItems(i + 5)
            End If
        End If
    Next i
    lvBus.SelectedItem.SubItems(2) = dbSettlePrice - dbSettleStationPrice
    lvBus.SelectedItem.SubItems(3) = dbSettleStationPrice
    
    UpdateTotalPrice
    Exit Sub
ErrorHandle:
    ShowErrorMsg
 
End Sub

Private Sub ModifyCompanySplitItem()

    Dim i As Integer
    Dim nIndex As Integer
    Dim atCompanyInfo() As TCompnaySettle
    Dim atSplitItem() As TSplitItemInfo
    Dim oReport As New Report
    Dim dbSettlePrice As Double
    Dim dbSettleStationPrice As Double
    
    On Error GoTo ErrorHandle
    
    
    oReport.Init g_oActiveUser
    atSplitItem = oReport.GetSplitItemInfo()

    
    If lvCompany.ListItems.Count = 0 Then Exit Sub
    '调用修改
    
    frmSplitItemModify.m_szObject = Trim(lvCompany.SelectedItem.Text)
    frmSplitItemModify.m_szProtocol = Trim(lvCompany.SelectedItem.SubItems(1))
    frmSplitItemModify.m_dbSettlePrice = Trim(lvCompany.SelectedItem.SubItems(2))
    frmSplitItemModify.m_dbStationPrice = Trim(lvCompany.SelectedItem.SubItems(3))
    frmSplitItemModify.m_nQuantity = Trim(lvCompany.SelectedItem.SubItems(4))
    frmSplitItemModify.m_nType = 2
    
    frmSplitItemModify.Show vbModal
    
    If frmSplitItemModify.m_bIsSave = False Then Exit Sub
    
    
    nIndex = lvCompany.SelectedItem.Index

    dbSettlePrice = 0
    dbSettleStationPrice = 0
    For i = 1 To m_nSplitItemCount
        If lvCompany.ColumnHeaders.Item(i + 6).Text = Trim(m_atSplitItem(i).SplitItemName) Then
            lvCompany.SelectedItem.SubItems(i + 5) = m_adbSplitItem(Val(m_atSplitItem(i).SplitItemID)) ' rsTemp1("split_item_" & Val(m_atSplitItem(i).SplitItemID)).Value
            
            If m_atSplitItem(i).SplitType = CS_SplitOtherCompany Then
                dbSettlePrice = dbSettlePrice + lvCompany.SelectedItem.SubItems(i + 5)
            ElseIf m_atSplitItem(i).SplitType = CS_SplitStation Then
                dbSettleStationPrice = dbSettleStationPrice + lvCompany.SelectedItem.SubItems(i + 5)
            End If
        End If
    Next i
    lvCompany.SelectedItem.SubItems(2) = dbSettlePrice - dbSettleStationPrice
    lvCompany.SelectedItem.SubItems(3) = dbSettleStationPrice
    
    UpdateTotalPrice
    
    Exit Sub
ErrorHandle:
    ShowErrorMsg
 
End Sub

Private Sub ModifyVehicleSplitItem()

    Dim i As Integer
    Dim nIndex As Integer
    Dim atVehicleInfo() As TVehilceSettle
    Dim atSplitItem() As TSplitItemInfo
    Dim oReport As New Report
    Dim dbSettlePrice As Double
    Dim dbSettleStationPrice As Double
    
    On Error GoTo ErrorHandle
    
    oReport.Init g_oActiveUser
    atSplitItem = oReport.GetSplitItemInfo()

    
    If lvVehicle.ListItems.Count = 0 Then Exit Sub
    '调用修改
    
    frmSplitItemModify.m_szObject = Trim(lvVehicle.SelectedItem.Text)
    frmSplitItemModify.m_szProtocol = Trim(lvVehicle.SelectedItem.SubItems(1))
    frmSplitItemModify.m_dbSettlePrice = Trim(lvVehicle.SelectedItem.SubItems(2))
    frmSplitItemModify.m_dbStationPrice = Trim(lvVehicle.SelectedItem.SubItems(3))
    frmSplitItemModify.m_nQuantity = Trim(lvVehicle.SelectedItem.SubItems(4))
    frmSplitItemModify.m_nType = 2
  
    frmSplitItemModify.m_eSettleObject = m_oSettleSheet.SettleObject
    frmSplitItemModify.m_szObjectID = m_oSettleSheet.ObjectID
    frmSplitItemModify.m_szSettleSheetID = m_oSettleSheet.SettleSheetID
  
    frmSplitItemModify.Show vbModal
    
    If frmSplitItemModify.m_bIsSave = False Then Exit Sub
    
    
    nIndex = lvVehicle.SelectedItem.Index

    dbSettlePrice = 0
    dbSettleStationPrice = 0
    For i = 1 To m_nSplitItemCount
        If lvVehicle.ColumnHeaders.Item(i + 6).Text = Trim(m_atSplitItem(i).SplitItemName) Then
            lvVehicle.SelectedItem.SubItems(i + 5) = m_adbSplitItem(Val(m_atSplitItem(i).SplitItemID)) ' rsTemp1("split_item_" & Val(m_atSplitItem(i).SplitItemID)).Value
            
            If m_atSplitItem(i).SplitType = CS_SplitOtherCompany Then
                dbSettlePrice = dbSettlePrice + lvVehicle.SelectedItem.SubItems(i + 5)
            ElseIf m_atSplitItem(i).SplitType = CS_SplitStation Then
                dbSettleStationPrice = dbSettleStationPrice + lvVehicle.SelectedItem.SubItems(i + 5)
            End If
        End If
    Next i
    lvVehicle.SelectedItem.SubItems(2) = dbSettlePrice - dbSettleStationPrice
    lvVehicle.SelectedItem.SubItems(3) = dbSettleStationPrice
    
    UpdateTotalPrice
    
    Exit Sub
ErrorHandle:
    ShowErrorMsg
 
End Sub
