VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Begin VB.Form frmWizSplitSettle 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "路单结算向导"
   ClientHeight    =   7470
   ClientLeft      =   2580
   ClientTop       =   1620
   ClientWidth     =   10950
   HelpContextID   =   7000280
   Icon            =   "frmWizSplitSettle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   10950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fraWizard 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "第一步"
      Height          =   5700
      Index           =   1
      Left            =   210
      TabIndex        =   19
      Top             =   930
      Width           =   10755
      Begin VB.CheckBox chkIsToday 
         BackColor       =   &H00E0E0E0&
         Caption         =   "当天结算"
         Height          =   315
         Left            =   1905
         TabIndex        =   101
         Top             =   5175
         Width           =   2685
      End
      Begin VB.ComboBox cboCompany 
         Height          =   300
         Left            =   6330
         TabIndex        =   100
         Top             =   3390
         Width           =   1785
      End
      Begin VB.TextBox txtSheetID 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3075
         TabIndex        =   9
         Top             =   3390
         Width           =   1620
      End
      Begin FText.asFlatTextBox txtObject 
         Height          =   315
         Left            =   6330
         TabIndex        =   3
         Top             =   2250
         Width           =   1785
         _ExtentX        =   3149
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
         Text            =   ""
         ButtonVisible   =   -1  'True
         OfficeXPColors  =   -1  'True
      End
      Begin MSComCtl2.DTPicker dtpEndDate 
         Height          =   315
         Left            =   6330
         TabIndex        =   7
         Top             =   2850
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   556
         _Version        =   393216
         Format          =   61800448
         CurrentDate     =   37642
      End
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   300
         Left            =   3075
         TabIndex        =   5
         Top             =   2850
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   529
         _Version        =   393216
         Format          =   61800448
         CurrentDate     =   37622
      End
      Begin MSComctlLib.ImageCombo imgcbo 
         Height          =   330
         Left            =   3075
         TabIndex        =   1
         Top             =   2235
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Locked          =   -1  'True
         ImageList       =   "imglv1"
      End
      Begin FText.asFlatTextBox txtRouteID 
         Height          =   300
         Left            =   6330
         TabIndex        =   11
         Top             =   3390
         Width           =   1785
         _ExtentX        =   3149
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
         OfficeXPColors  =   -1  'True
      End
      Begin VB.Label lblCompany 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "参运公司(&C):"
         Height          =   180
         Left            =   5055
         TabIndex        =   91
         Top             =   3435
         Width           =   1080
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "结算单编号(&S):"
         Height          =   180
         Left            =   1770
         TabIndex        =   8
         Top             =   3450
         Width           =   1260
      End
      Begin VB.Label lblRoute 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "线路(&R):"
         Height          =   180
         Left            =   5085
         TabIndex        =   10
         Top             =   3420
         Width           =   720
      End
      Begin VB.Line Line6 
         BorderColor     =   &H8000000E&
         X1              =   1080
         X2              =   9090
         Y1              =   1545
         Y2              =   1545
      End
      Begin VB.Line Line5 
         BorderColor     =   &H8000000C&
         X1              =   1095
         X2              =   9060
         Y1              =   1500
         Y2              =   1500
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "类型(&S):"
         Height          =   180
         Index           =   1
         Left            =   1770
         TabIndex        =   0
         Top             =   2310
         Width           =   720
      End
      Begin VB.Label label2 
         BackStyle       =   0  'Transparent
         Caption         =   "对象(&O):"
         Height          =   180
         Index           =   1
         Left            =   5070
         TabIndex        =   2
         Top             =   2310
         Width           =   810
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "开始日期(&B):"
         Height          =   225
         Left            =   1770
         TabIndex        =   4
         Top             =   2895
         Width           =   1380
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "结束日期(&E):"
         Height          =   180
         Left            =   5070
         TabIndex        =   6
         Top             =   2910
         Width           =   1080
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmWizSplitSettle.frx":014A
         Height          =   840
         Left            =   1740
         TabIndex        =   20
         Top             =   660
         Width           =   6165
      End
   End
   Begin VB.Frame fraWizard 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "第四步"
      Height          =   5670
      Index           =   3
      Left            =   105
      TabIndex        =   83
      Top             =   990
      Width           =   10725
      Begin VSFlex7LCtl.VSFlexGrid vsStationList 
         Height          =   4935
         Left            =   180
         TabIndex        =   89
         Top             =   540
         Width           =   10380
         _cx             =   18309
         _cy             =   8705
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
      Begin VSFlex7LCtl.VSFlexGrid vsStationTotal 
         Height          =   4935
         Left            =   180
         TabIndex        =   85
         Top             =   540
         Width           =   10380
         _cx             =   18309
         _cy             =   8705
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
      Begin VSFlex7LCtl.VSFlexGrid vsStationDayList 
         Height          =   4950
         Left            =   180
         TabIndex        =   86
         Top             =   540
         Width           =   10365
         _cx             =   18283
         _cy             =   8731
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
      Begin MSComctlLib.TabStrip tsStation 
         Height          =   5520
         Left            =   45
         TabIndex        =   84
         Top             =   120
         Width           =   10620
         _ExtentX        =   18733
         _ExtentY        =   9737
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   3
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "站点汇总"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "每日站点汇总"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "站点人数清单"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fraWizard 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "第三步"
      Height          =   5670
      Index           =   2
      Left            =   210
      TabIndex        =   21
      Top             =   960
      Width           =   10725
      Begin VB.CheckBox chkSingle 
         BackColor       =   &H00E0E0E0&
         Caption         =   "按单行"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3120
         TabIndex        =   81
         Top             =   240
         Width           =   915
      End
      Begin VB.TextBox txtCheckSheetID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
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
         Height          =   255
         Left            =   1020
         TabIndex        =   52
         Top             =   180
         Width           =   1245
      End
      Begin RTComctl3.FlatLabel FlatLabel1 
         Height          =   435
         Left            =   930
         TabIndex        =   51
         Top             =   90
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   767
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         InnerStyle      =   2
         Caption         =   ""
      End
      Begin MSComctlLib.ListView lvObject 
         Height          =   4965
         Left            =   15
         TabIndex        =   54
         Top             =   660
         Width           =   10665
         _ExtentX        =   18812
         _ExtentY        =   8758
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ColHdrIcons     =   "imgObject"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.Image imgEnabled 
         Height          =   480
         Left            =   2550
         Picture         =   "frmWizSplitSettle.frx":0235
         Top             =   90
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label lblTotalQuantity 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   4740
         TabIndex        =   58
         Top             =   270
         Width           =   105
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "总人数:"
         Height          =   180
         Left            =   4035
         TabIndex        =   57
         Top             =   270
         Width           =   630
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "路单张数:"
         Height          =   180
         Left            =   7125
         TabIndex        =   50
         Top             =   270
         Width           =   810
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "路单有效张数:"
         Height          =   180
         Left            =   5385
         TabIndex        =   49
         Top             =   270
         Width           =   1170
      End
      Begin VB.Label lblEnableCount 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   6600
         TabIndex        =   48
         Top             =   270
         Width           =   105
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "有效"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   2550
         TabIndex        =   46
         Top             =   210
         Width           =   540
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "路单号:"
         Height          =   180
         Left            =   240
         TabIndex        =   23
         Top             =   240
         Width           =   630
      End
      Begin VB.Label lblSettleSheetCount 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   8040
         TabIndex        =   22
         Top             =   270
         Width           =   105
      End
   End
   Begin VB.Frame fraWizard 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "第五步"
      Height          =   5730
      Index           =   5
      Left            =   90
      TabIndex        =   35
      Top             =   900
      Width           =   10785
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "备注"
         Height          =   1755
         Left            =   300
         TabIndex        =   102
         Top             =   3330
         Width           =   10275
         Begin FText.asFlatMemo txtAnnotation 
            Height          =   1245
            Left            =   210
            TabIndex        =   103
            Top             =   300
            Width           =   9825
            _ExtentX        =   17330
            _ExtentY        =   2196
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonHotForeColor=   -2147483628
            ButtonHotBackColor=   -2147483632
         End
      End
      Begin VB.Frame fraBus 
         BackColor       =   &H00E0E0E0&
         Caption         =   "车次结算明细表"
         Height          =   1935
         Left            =   300
         TabIndex        =   92
         Top             =   1170
         Width           =   10275
         Begin MSComctlLib.ListView lvBus 
            Height          =   1425
            Left            =   210
            TabIndex        =   93
            Top             =   300
            Width           =   9825
            _ExtentX        =   17330
            _ExtentY        =   2514
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
            Appearance      =   0
            NumItems        =   0
         End
      End
      Begin VB.Frame fraCompany 
         BackColor       =   &H00E0E0E0&
         Caption         =   "公司结算明细表"
         Height          =   1935
         Left            =   300
         TabIndex        =   47
         Top             =   1170
         Width           =   10275
         Begin MSComctlLib.ListView lvCompany 
            Height          =   1425
            Left            =   180
            TabIndex        =   56
            Top             =   300
            Width           =   9825
            _ExtentX        =   17330
            _ExtentY        =   2514
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
            Appearance      =   0
            NumItems        =   0
         End
      End
      Begin VB.Frame fraVehicle 
         BackColor       =   &H00E0E0E0&
         Caption         =   "车辆结算明细表"
         Height          =   1935
         Left            =   300
         TabIndex        =   41
         Top             =   1170
         Width           =   10275
         Begin MSComctlLib.ListView lvVehicle 
            Height          =   1425
            Left            =   180
            TabIndex        =   55
            Top             =   300
            Width           =   9810
            _ExtentX        =   17304
            _ExtentY        =   2514
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
            Appearance      =   0
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "路单结算总汇"
         Height          =   675
         Left            =   300
         TabIndex        =   36
         Top             =   180
         Width           =   10275
         Begin VB.TextBox txtAdditionPrice 
            Height          =   300
            Left            =   8340
            TabIndex        =   94
            Top             =   240
            Width           =   1005
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "补差款:"
            Height          =   180
            Left            =   7650
            TabIndex        =   95
            Top             =   300
            Width           =   630
         End
         Begin VB.Label lblTotalPrice 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "100"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   840
            TabIndex        =   45
            Top             =   300
            Width           =   495
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "总票款:"
            Height          =   180
            Left            =   180
            TabIndex        =   44
            Top             =   300
            Width           =   630
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "路单张数:"
            Height          =   180
            Left            =   5850
            TabIndex        =   43
            Top             =   300
            Width           =   810
         End
         Begin VB.Label lblSheetCount 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "2张"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   6720
            TabIndex        =   42
            Top             =   300
            Width           =   270
         End
         Begin VB.Label lblNeedSplitMoney 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "220.3"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   2520
            TabIndex        =   40
            Top             =   300
            Width           =   450
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "应结票款:"
            Height          =   180
            Left            =   1650
            TabIndex        =   39
            Top             =   300
            Width           =   810
         End
         Begin VB.Label lblTotalQuautity 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "220.3"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   4590
            TabIndex        =   38
            Top             =   300
            Width           =   450
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "总人数:"
            Height          =   180
            Left            =   3900
            TabIndex        =   37
            Top             =   300
            Width           =   630
         End
      End
   End
   Begin VB.Frame fraWizard 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   5520
      Index           =   4
      Left            =   30
      TabIndex        =   96
      Top             =   990
      Width           =   10815
      Begin VSFlex7LCtl.VSFlexGrid vsExtra 
         Height          =   4755
         Left            =   255
         TabIndex        =   98
         Top             =   525
         Width           =   10395
         _cx             =   18336
         _cy             =   8387
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
         Editable        =   2
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
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   4335
         Top             =   30
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmWizSplitSettle.frx":0AFF
               Key             =   "add"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmWizSplitSettle.frx":0DA5
               Key             =   "del"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tbSection 
         Height          =   360
         Left            =   2085
         TabIndex        =   99
         Top             =   30
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   635
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "add"
               Object.ToolTipText     =   "新增"
               ImageKey        =   "add"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "删除"
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "del"
               ImageKey        =   "del"
            EndProperty
         EndProperty
         BorderStyle     =   1
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "手工补票信息:"
         Height          =   270
         Left            =   360
         TabIndex        =   97
         Top             =   60
         Width           =   1890
      End
   End
   Begin RTComctl3.CoolButton cmdNext 
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   315
      Left            =   7080
      TabIndex        =   12
      Top             =   6975
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
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
      MICON           =   "frmWizSplitSettle.frx":0F26
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
      Left            =   7080
      TabIndex        =   34
      Top             =   6975
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
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
      MICON           =   "frmWizSplitSettle.frx":0F42
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdAddSheet 
      Height          =   315
      Left            =   1800
      TabIndex        =   87
      Top             =   6975
      Visible         =   0   'False
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "添加路单(&A)"
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
      MICON           =   "frmWizSplitSettle.frx":0F5E
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
      Height          =   315
      Left            =   3165
      TabIndex        =   88
      Top             =   6975
      Visible         =   0   'False
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "清除所有"
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
      MICON           =   "frmWizSplitSettle.frx":0F7A
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
      Left            =   8940
      TabIndex        =   14
      Top             =   6975
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
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
      MICON           =   "frmWizSplitSettle.frx":0F96
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdPrevious 
      Height          =   315
      Left            =   5820
      TabIndex        =   13
      Top             =   6975
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "上一步(&P)"
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
      MICON           =   "frmWizSplitSettle.frx":0FB2
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
      Left            =   255
      TabIndex        =   90
      Top             =   6975
      Width           =   1125
      _ExtentX        =   1984
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
      MICON           =   "frmWizSplitSettle.frx":0FCE
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
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   -30
      TabIndex        =   80
      Top             =   840
      Width           =   11025
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   -120
      TabIndex        =   17
      Top             =   -300
      Width           =   8775
   End
   Begin VB.PictureBox ptTitle 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   10965
      TabIndex        =   31
      Top             =   0
      Width           =   10965
      Begin VB.Label lblContent 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请选择路单结算的方式。"
         Height          =   180
         Left            =   360
         TabIndex        =   33
         Top             =   450
         Width           =   1980
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "结算方式"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   0
         Left            =   180
         TabIndex        =   32
         Top             =   150
         Width           =   780
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   " "
      Height          =   990
      Left            =   -30
      TabIndex        =   18
      Top             =   6645
      Width           =   11265
      Begin MSComctlLib.ProgressBar pbFill 
         Height          =   285
         Left            =   1620
         TabIndex        =   82
         Top             =   300
         Visible         =   0   'False
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
      End
   End
   Begin MSComctlLib.ImageList imglv1 
      Left            =   8610
      Top             =   900
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWizSplitSettle.frx":0FEA
            Key             =   "splitcompany"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWizSplitSettle.frx":1144
            Key             =   "company"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWizSplitSettle.frx":129E
            Key             =   "vehicle"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWizSplitSettle.frx":13F8
            Key             =   "bus"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWizSplitSettle.frx":1552
            Key             =   "busowner"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraWizard 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "最后一步"
      Height          =   5685
      Index           =   8
      Left            =   180
      TabIndex        =   15
      Top             =   945
      Width           =   10770
      Begin MSComctlLib.ProgressBar CreateProgressBar 
         Height          =   300
         Left            =   660
         Negotiate       =   -1  'True
         TabIndex        =   53
         Top             =   4350
         Width           =   7710
         _ExtentX        =   13600
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.ListBox lstCreateInfo 
         Appearance      =   0  'Flat
         Height          =   3630
         ItemData        =   "frmWizSplitSettle.frx":16AC
         Left            =   270
         List            =   "frmWizSplitSettle.frx":16AE
         MultiSelect     =   2  'Extended
         TabIndex        =   28
         Top             =   450
         Width           =   8115
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "执行情况:"
         Height          =   255
         Left            =   300
         TabIndex        =   30
         Top             =   180
         Width           =   1095
      End
      Begin VB.Label lblProgress 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   29
         Top             =   4380
         Width           =   615
      End
   End
   Begin VB.Frame fraWizard 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   5730
      Index           =   7
      Left            =   75
      TabIndex        =   68
      Top             =   1035
      Width           =   10800
      Begin MSComctlLib.ListView lvLugSheet 
         Height          =   3015
         Left            =   960
         TabIndex        =   69
         Top             =   780
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   5318
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "行包结算单"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblLugProtocol 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   5460
         TabIndex        =   79
         Top             =   2490
         Width           =   105
      End
      Begin VB.Label lblLugNeedSplit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   5460
         TabIndex        =   78
         Top             =   2010
         Width           =   105
      End
      Begin VB.Label lblLugTotalPrice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   5460
         TabIndex        =   77
         Top             =   1500
         Width           =   105
      End
      Begin VB.Label lblLubObject 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   5460
         TabIndex        =   76
         Top             =   1080
         Width           =   105
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "行包结算单:"
         Height          =   180
         Left            =   960
         TabIndex        =   75
         Top             =   420
         Width           =   990
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "行包结算汇总信息:"
         Height          =   180
         Left            =   4200
         TabIndex        =   74
         Top             =   450
         Width           =   1530
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "拆算协议:"
         Height          =   180
         Left            =   4230
         TabIndex        =   73
         Top             =   2490
         Width           =   810
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "拆算对象:"
         Height          =   180
         Left            =   4200
         TabIndex        =   72
         Top             =   1050
         Width           =   810
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "应拆金额:"
         Height          =   180
         Left            =   4230
         TabIndex        =   71
         Top             =   2010
         Width           =   810
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "行包运费:"
         Height          =   180
         Left            =   4200
         TabIndex        =   70
         Top             =   1560
         Width           =   810
      End
   End
   Begin VB.Frame fraWizard 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "第二步"
      Height          =   5700
      Index           =   9
      Left            =   105
      TabIndex        =   16
      Top             =   975
      Width           =   10785
      Begin VB.TextBox txtOperator 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   540
         TabIndex        =   24
         Top             =   3330
         Visible         =   0   'False
         Width           =   1785
      End
      Begin MSComCtl2.DTPicker dtpSplitDate 
         Height          =   300
         Left            =   510
         TabIndex        =   27
         Top             =   3810
         Visible         =   0   'False
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   529
         _Version        =   393216
         Format          =   61800448
         CurrentDate     =   37642
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "拆帐日期(&D):"
         Height          =   180
         Left            =   840
         TabIndex        =   26
         Top             =   4305
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "拆帐人(&P):"
         Height          =   180
         Left            =   270
         TabIndex        =   25
         Top             =   2460
         Visible         =   0   'False
         Width           =   900
      End
   End
   Begin VB.Frame fraWizard 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   5715
      Index           =   6
      Left            =   60
      TabIndex        =   59
      Top             =   1005
      Width           =   10785
      Begin RTComctl3.CoolButton cmdDetele 
         Height          =   345
         Left            =   3210
         TabIndex        =   66
         Top             =   1710
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "移除<<"
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
         MICON           =   "frmWizSplitSettle.frx":16B0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin RTComctl3.CoolButton cmdAdd 
         Height          =   345
         Left            =   3210
         TabIndex        =   65
         Top             =   1140
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "添加>>"
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
         MICON           =   "frmWizSplitSettle.frx":16CC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox txtLugSheetID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   915
         TabIndex        =   61
         Top             =   1140
         Width           =   1995
      End
      Begin MSComctlLib.ListView lvLugSheetID 
         Height          =   3255
         Left            =   4440
         TabIndex        =   60
         Top             =   1110
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   5741
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "行包结算单"
            Object.Width           =   4410
         EndProperty
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "所选的结算单:"
         Height          =   180
         Left            =   4440
         TabIndex        =   67
         Top             =   780
         Width           =   1170
      End
      Begin VB.Label lblSplitObject 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "苏州公司"
         Height          =   180
         Left            =   1980
         TabIndex        =   64
         Top             =   390
         Width           =   720
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "拆算对象:"
         Height          =   180
         Left            =   930
         TabIndex        =   63
         Top             =   390
         Width           =   810
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "行包结算单号:"
         Height          =   180
         Left            =   930
         TabIndex        =   62
         Top             =   780
         Width           =   1170
      End
   End
   Begin VB.Menu pmnu_Select 
      Caption         =   "选择"
      Visible         =   0   'False
      Begin VB.Menu pmnu_AllSelect 
         Caption         =   "全选(&S)"
      End
      Begin VB.Menu pmnu_AllUnSelect 
         Caption         =   "重选(&U)"
      End
      Begin VB.Menu pmnu_SelectCheck 
         Caption         =   "指定全选(&C)"
      End
      Begin VB.Menu pmnu_UnSelectCheck 
         Caption         =   "指定重选(&N)"
      End
   End
End
Attribute VB_Name = "frmWizSplitSettle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3888FDA40140"

Option Explicit

'imgcbo 常数
Const cnBus = 3
Const cnVehicle = 2
Const cnCompany = 1


'路单列首常数
Const PI_CheckSheetID = 0
Const PI_Station = 1
Const PI_BusID = 2
Const PI_Date = 3
Const PI_BusSerialNO = 4
Const PI_TicketPrice = 5
Const PI_LicenseTagNo = 6
Const PI_CompanyName = 7
Const PI_Quantity = 8
Const PI_Mileage = 9
Const PI_Route = 10
Const PI_CheckGate = 11
Const PI_VehicleID = 12
Const PI_CompanyID = 13


'路单站点汇总的列首位置
Const VI_SellStation = 1
Const VI_Route = 2
Const VI_Bus = 3
Const VI_VehicleType = 4
Const VI_Station = 5
Const VI_TicketType = 6
Const VI_Quantity = 7
Const VI_AreaRatio = 8
Const VI_RouteID = 9
Const VI_VehicleTypeID = 10
Const VI_SellStationID = 11
Const VI_StationID = 12
Const VI_TicketTypeID = 13


    
'路单站点清单的列首位置
Const VIL_CheckSheetID = 1
Const VIL_BusDate = 2
Const VIL_BusID = 3
Const VIL_SellStationName = 4
Const VIL_StationName = 5
Const VIL_TicketTypeName = 6
Const VIL_PriceIdentify = 7
Const VIL_SellStationID = 8
Const VIL_StationID = 9
Const VIL_TicketType = 10
Const VIL_StatusName = 11 '改并状态
Const VIL_RouteID = 12
Const VIL_VehicleTypeID = 13
Const VIL_SeatTypeID = 14
Const VIL_BusSerialNO = 15
Const VIL_StationSerial = 16
Const VIL_AreaRatio = 17
'Const VIL_RouteName = 18
Const VIL_Quantity = 18
Const VIL_Mileage = 19
Const VIL_TicketPrice = 20 '单价
Const VIL_TotalTicketPrice = 21 '总价
Const VIL_StatusCode = 22 '改并状态
Const VIL_BasePrice = 23

'手工补票表格的列首位置
Const VIE_NO = 0
Const VIE_Quantity = 1
Const VIE_TotalTicketPrice = 2
Const VIE_Ratio = 3
Const VIE_ServicePrice = 4
Const VIE_SettleOutPrice = 5



Dim m_nSheetCount As Integer  ' 统计路单总数
Dim nValibleCount As Integer  '有效的路单数
Dim nTotalQuantity As Long '总人数
Dim m_aszCheckSheetID() As String '路单组数
Dim m_szVehicleID As String  '车辆数组
Dim m_szCompanyID As String  '公司ID
Dim m_szBusID As String '车次ID
Dim m_szAdditionPrice As Integer  '补差款
Dim m_bLogFileValid As Boolean '日志文件
Dim m_bPromptWhenError As Boolean '是否提示错误
Dim CancelHasPress As Boolean
Dim TSplitResult As TSplitResult  '预览结果
Dim m_nSplitItenCount As Integer '使用的拆算项数
Dim tSplitItem() As TSplitItemInfo
Dim AdditionPrice As String

Dim m_szTransportCompanyName As String
Dim m_szTransportCompanyID As String


Public m_bIsManualSettle As Boolean '是否是手工结算,即汇总出来的人数及站点信息,是否允许修改
Dim m_rsStationQuantity As Recordset '站点人数汇总信息
Dim m_rsStationInfo As Recordset '每日汇总,的站点信息

Dim oPriceMan As New STPrice.TicketPriceMan
Dim m_rsPriceItem As Recordset  '可用的票价项
Dim m_rsStationList As Recordset '站点人数明细信息

Dim m_nQuantity As Integer '保存修改前的人数


Dim m_oSplit As New STSettle.Split
Dim m_atExtraInfo() As TSettleExtraInfo
Dim szSplitType As String

Dim szLicenseTagNO As String



Private Sub cmdAdd_Click()
    AddLugSheet
End Sub

Private Sub AddLugSheet()
 On Error GoTo ErrHandle
    Dim i As Integer
    Dim m_oReport As New Report
    Dim rsTemp As Recordset
    '判断行包单是否有效,拆算对象是否为所指定的对象
    m_oReport.Init g_oActiveUser
    
    '满足条件
    If txtLugSheetID.Text <> "" Then
        Set rsTemp = m_oReport.GetLugSheet(Trim(txtLugSheetID.Text))
        If rsTemp.RecordCount = 0 Then
            MsgBox "此行包结算单无效!", vbExclamation, Me.Caption
            Exit Sub
        End If
        lvLugSheetID.ListItems.Add , , Trim(txtLugSheetID.Text)
        txtLugSheetID.Text = ""
        txtLugSheetID.SetFocus
        cmdDetele.Enabled = True
    End If
    Exit Sub
ErrHandle:
 ShowErrorMsg

End Sub

Private Sub cmdAddSheet_Click()
    '添加其他的路单，这些路单为检票时，将车牌设置错误，导致未出来的路单。
    Dim aszTemp() As String
    Dim m_oReport As New Report
    Dim rsTemp As Recordset
    Dim nCount As Integer
    Dim i As Integer
    Dim lvItem As ListItem
    Dim j As Integer
    frmSelSheet.m_dtStartDate = dtpStartDate.Value
    frmSelSheet.m_dtEndDate = dtpEndDate.Value
    frmSelSheet.Show vbModal
    If frmSelSheet.m_szSheetID <> "" Then
        '显示选择的路单信息
        SetBusy
        aszTemp = StringToTeam(frmSelSheet.m_szSheetID)
        
        m_oReport.Init g_oActiveUser
        Set rsTemp = m_oReport.GetNeedSplitCheckSheet(CS_SettleByVehicle, ResolveDisplay(Trim(txtObject.Text)), dtpStartDate.Value, dtpEndDate.Value, ResolveDisplay(txtRouteID.Text), aszTemp, IIf(chkIsToday.Value = vbChecked, True, False))
        
        nCount = rsTemp.RecordCount
        
        If nCount = 0 Then
            SetNormal
            Exit Sub
        End If
        pbFill.Visible = True
        pbFill.Max = nCount
        
        For i = 1 To nCount
            For j = 1 To lvObject.ListItems.Count
                If lvObject.ListItems(j) = FormatDbValue(rsTemp!check_sheet_id) Then
                    '如果找到存在该路单,则忽略不添加
                    Exit For
                End If
            
            Next j
            '未找到该路单在列表中存在
            If j > lvObject.ListItems.Count Then
        
                Set lvItem = lvObject.ListItems.Add(, , FormatDbValue(rsTemp!check_sheet_id))
                lvItem.SubItems(PI_Station) = FormatDbValue(rsTemp!station_name)
                lvItem.SubItems(PI_CheckGate) = FormatDbValue(rsTemp!check_gate_id)
                lvItem.SubItems(PI_Route) = FormatDbValue(rsTemp!route_name)
                lvItem.SubItems(PI_BusID) = FormatDbValue(rsTemp!bus_id)
                lvItem.SubItems(PI_BusSerialNO) = FormatDbValue(rsTemp!bus_serial_no)
                lvItem.SubItems(PI_LicenseTagNo) = FormatDbValue(rsTemp!license_tag_no)
                lvItem.SubItems(PI_CompanyName) = FormatDbValue(rsTemp!transport_company_short_name)
                lvItem.SubItems(PI_Quantity) = FormatDbValue(rsTemp!Quantity)
                lvItem.SubItems(PI_Mileage) = FormatDbValue(rsTemp!Mileage)    '里程
                lvItem.SubItems(PI_TicketPrice) = FormatDbValue(rsTemp!ticket_price)
                lvItem.SubItems(PI_Date) = Format(FormatDbValue(rsTemp!bus_date), "yy-mm-dd")
                lvItem.SubItems(PI_VehicleID) = FormatDbValue(rsTemp!vehicle_id)
                lvItem.SubItems(PI_CompanyID) = FormatDbValue(rsTemp!transport_company_id)
                pbFill.Value = i
                lvItem.Checked = True
                lvObject_ItemCheck lvItem
            End If
            rsTemp.MoveNext
            
        Next i
        pbFill.Visible = False
        SetNormal
    End If
    frmSelSheet.m_szSheetID = ""
    
End Sub

Private Sub cmdClear_Click()
    lvObject.ListItems.Clear
    
    nValibleCount = 0
    nTotalQuantity = 0
    lblEnableCount.Caption = nValibleCount
    lblTotalQuantity.Caption = nTotalQuantity
    
    m_nSheetCount = lvObject.ListItems.Count
    lblSettleSheetCount.Caption = CStr(m_nSheetCount)
End Sub

Private Sub cmdDetele_Click()
 On Error GoTo ErrHandle
  If lvLugSheetID.ListItems.Count = 0 Then
     cmdDetele.Enabled = False
     Exit Sub
  End If
  lvLugSheetID.ListItems.Remove (lvLugSheetID.SelectedItem.Index)
 
 Exit Sub
ErrHandle:
 ShowErrorMsg
End Sub

Private Sub cmdFinish_Click()
    On Error GoTo ErrHandle
    Dim i As Integer
'    Dim m_oSplit As New STSettle.Split
    Dim mAnswer As VbMsgBoxResult
    Dim szTemp As String
    SetBusy
    
    '权限验证
    cmdFinish.Visible = True
    cmdFinish.Enabled = False
    cmdNext.Visible = False
'    fraWizard(1).Visible = False
'    fraWizard(2).Visible = False
'    fraWizard(3).Visible = False
'    fraWizard(4).Visible = False
'    fraWizard(5).Visible = False
'    fraWizard(6).Visible = False
'    fraWizard(7).Visible = False
'    fraWizard(8).Visible = True
'    m_oSplit.Init g_oActiveUser
    
    '开始生成结算单 填写lstCreateInfo信息
    CreateFinanceSheetRs
    '打印结算单
    frmPrintFinSheet.m_SheetID = Trim(txtSheetID.Text)
    For i = 1 To lvLugSheet.ListItems.Count
        If i <> lvLugSheet.ListItems.Count Then
            szTemp = szTemp & lvLugSheet.ListItems(i).Text & ","
        Else
            szTemp = szTemp & lvLugSheet.ListItems(i).Text
        End If
    Next i
    frmPrintFinSheet.m_szLugSettleSheetID = szTemp
    frmPrintFinSheet.m_bRePrint = False
'    frmPrintFinSheet.m_bNeedPrint = True
    
    frmPrintFinSheet.ZOrder 0
    frmPrintFinSheet.Show vbModal
    
    SaveReg
    Unload Me
    
    frmWizSplitSettle.ZOrder 0
    frmWizSplitSettle.Show vbModal
    
    SetNormal
    Exit Sub
ErrHandle:
    SetNormal
    ShowErrorMsg
End Sub
'
''//***********************************    生成日志区  **************************************
''生成结算单日志
'Private Sub CreateFinanceInfo()
'    On Error GoTo ErrHandle
'
'    CreateFinanceSheetRs
'
'
'    Exit Sub
'ErrHandle:
'    ShowErrorMsg
'End Sub

Private Sub SaveReg()
Dim oFreeReg As CFreeReg
    Set oFreeReg = New CFreeReg
    oFreeReg.Init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany
    oFreeReg.SaveSetting m_cRegSystemKey, "SplitType", imgcbo.Text
    Set oFreeReg = Nothing
End Sub

Private Sub GetReg()
    Dim oReg As CFreeReg
    Set oReg = New CFreeReg

    oReg.Init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany
    szSplitType = Trim(oReg.GetSetting(m_cRegSystemKey, "SplitType"))
    Set oReg = Nothing
End Sub


'生成某结算单
Private Sub CreateFinanceSheetRs()
    Dim aszSheetID() As String   '为了适应接口数组类型
    Dim oVehicle As New Vehicle
    Dim oCompany As New Company
    Dim oBus As New Bus
    Dim szObjectName As String
    Dim szLuggageSettleIDs As String
    Dim i As Integer
    Dim rsTemp As Recordset
    
    On Error GoTo here
    
    
    
    '****取得对象名称
    If imgcbo.Text = cszCompanyName Then
        oCompany.Init g_oActiveUser
        oCompany.Identify m_szCompanyID
        szObjectName = oCompany.CompanyShortName
        '考虑到有补差款，对TSplitResult.SettleSheetInfo.SettleStationPrice的值进行更新
        TSplitResult.SettleSheetInfo.SettleStationPrice = lvCompany.SelectedItem.SubItems(3)
        '考虑到有补差款，将补差款添加到相应的费用项中
        TSplitResult.CompanyInfo(1).SplitItem(m_oSplit.m_nServiceItem) = TSplitResult.CompanyInfo(1).SplitItem(m_oSplit.m_nServiceItem) + Val(txtAdditionPrice.Text)
        
    ElseIf imgcbo.Text = "车辆" Then
        oVehicle.Init g_oActiveUser
        oVehicle.Identify m_szVehicleID
        szObjectName = szLicenseTagNO 'oVehicle.LicenseTag
        '考虑到有补差款，对TSplitResult.SettleSheetInfo.SettleStationPrice的值进行更新
        TSplitResult.SettleSheetInfo.SettleStationPrice = lvVehicle.SelectedItem.SubItems(3)
        '考虑到有补差款，将补差款添加到相应的费用项中
        TSplitResult.VehicleInfo(1).SplitItem(m_oSplit.m_nServiceItem) = TSplitResult.VehicleInfo(1).SplitItem(m_oSplit.m_nServiceItem) + Val(txtAdditionPrice.Text)
        
    ElseIf imgcbo.Text = "车次" Then
'        oBus.Init g_oActiveUser
'        oBus.Identify m_szBusID
        szObjectName = m_szBusID
        TSplitResult.SettleSheetInfo.TransportCompanyID = m_szTransportCompanyID
        TSplitResult.SettleSheetInfo.TransportCompanyName = m_szTransportCompanyName
        m_szCompanyID = m_szTransportCompanyID
        '考虑到有补差款，对TSplitResult.SettleSheetInfo.SettleStationPrice的值进行更新
        TSplitResult.SettleSheetInfo.SettleStationPrice = lvBus.SelectedItem.SubItems(3)
        '考虑到有补差款，将补差款添加到相应的费用项中
        TSplitResult.BusInfo(1).SplitItem(m_oSplit.m_nServiceItem) = TSplitResult.BusInfo(1).SplitItem(m_oSplit.m_nServiceItem) + Val(txtAdditionPrice.Text)
    End If
    TSplitResult.SettleSheetInfo.ObjectName = szObjectName

    '****取得行包相关信息
    szLuggageSettleIDs = ""
    
    For i = 1 To lvLugSheet.ListItems.Count
        szLuggageSettleIDs = szLuggageSettleIDs & lvLugSheet.ListItems(i)
    Next i
    TSplitResult.SettleSheetInfo.LuggageSettleIDs = szLuggageSettleIDs
    TSplitResult.SettleSheetInfo.LuggageTotalBaseCarriage = Val(lblLugTotalPrice.Caption)
    TSplitResult.SettleSheetInfo.LuggageTotalSettlePrice = Val(lblLugNeedSplit.Caption)
    TSplitResult.SettleSheetInfo.LuggageProtocolName = lblLugProtocol.Caption
    TSplitResult.SettleExtraInfo = m_atExtraInfo
    TSplitResult.SheetStationInfo = GetStationQuantity()
    TSplitResult.SettleSheetInfo.Annotation = txtAnnotation.Text
    
'    '开始结算
'    If QuantityHasBeChanged Then
'        m_oSplit.SplitCheckSheetManual Trim(txtSheetID.Text), m_szCompanyID, m_szVehicleID, m_aszCheckSheetID, TSplitResult, dtpStartDate.Value, dtpEndDate.Value
'    Else
'        m_oSplit.SplitCheckSheet Trim(txtSheetID.Text), m_szCompanyID, m_szVehicleID, m_aszCheckSheetID, TSplitResult, dtpStartDate.Value, dtpEndDate.Value
'    End If
    m_szAdditionPrice = Val(txtAdditionPrice.Text)
    Set rsTemp = GetSheetStationInfoRS
    m_oSplit.SplitCheckSheet Trim(txtSheetID.Text), m_szCompanyID, m_szVehicleID, m_aszCheckSheetID, TSplitResult, dtpStartDate.Value, dtpEndDate.Value, rsTemp, m_szAdditionPrice, m_szBusID, IIf(chkIsToday.Value = vbChecked, True, False)
    
    Exit Sub
    
    
here:
    ShowErrorMsg
End Sub

Private Sub CloseLogFile()
    On Error Resume Next
    If m_bLogFileValid Then
        Close #1
    End If
End Sub

Private Sub RecordLog(log As String)
    With lstCreateInfo
        .AddItem log
        .ListIndex = .ListCount - 1
        .Refresh
    End With
    If m_bLogFileValid Then
        AddLogToFile log
    End If
End Sub

Private Sub AddLogToFile(log As String)
    On Error Resume Next
    If m_bLogFileValid Then
        Print #1, log
    End If
End Sub

Private Sub cmdHelp_Click()
    DisplayHelp Me
End Sub

Private Sub cmdNext_Click()
    On Error GoTo ErrHandle
    Dim rsTemp As Recordset
    Dim szLastSheetID As String
    Dim i As Integer
'    Dim m_oSplit As New STSettle.Split
    SetBusy
    pbFill.Visible = False
'    m_oSplit.Init g_oActiveUser
    cmdPrevious.Enabled = True
    If fraWizard(1).Visible = True Then
        '判断拆算对象有没有选
        If txtObject.Text = "" Then
            MsgBox " 请选择拆算对象!", vbInformation, Me.Caption
            cmdPrevious.Enabled = False
            SetNormal
            Exit Sub
        End If

        '判断结算单号是否为空
        If txtSheetID.Text = "" Then
            MsgBox " 结算单编号不能为空!", vbInformation, Me.Caption
            SetNormal
            Exit Sub
        End If
        If chkIsToday.Value = vbChecked Then
            '如果是当天结算,则日期不能大于一天
            If ToDBDate(dtpStartDate.Value) <> ToDBDate(dtpEndDate.Value) Then
                MsgBox "如果为当天结算,则开始日期与结束日期必须相同", vbExclamation, Me.Caption
                SetNormal
                Exit Sub
            End If
        
        
        End If
        lvObject.ListItems.Clear
        cmdNext.Enabled = True
        cmdFinish.Visible = False
        cmdFinish.Enabled = True
        cmdNext.Visible = True
        fraWizard(1).Visible = False
        fraWizard(2).Visible = True
        fraWizard(3).Visible = False
        fraWizard(4).Visible = False
        fraWizard(5).Visible = False
        fraWizard(6).Visible = False
        fraWizard(7).Visible = False
        fraWizard(8).Visible = False
        fraWizard(9).Visible = False
        
        cmdNext.Default = False
        fraWizard(2).Refresh
        
        
        
        
        
        '填充要拆算的路单记录
        FilllvObject
        '显示签发单总数
        m_nSheetCount = lvObject.ListItems.Count
        lblSettleSheetCount.Caption = CStr(m_nSheetCount)
        nValibleCount = 0 '扫描路单的有效张数
        nTotalQuantity = 0
        lblEnableCount.Caption = nValibleCount
        lblTotalQuantity.Caption = nTotalQuantity
        txtCheckSheetID.Text = ""
        '等待扫描路单
        lblStatus.Visible = False
        txtCheckSheetID.Text = ""
        txtCheckSheetID.SetFocus
        
        cmdAddSheet.Visible = True
        cmdClear.Visible = True
        
    ElseIf fraWizard(2).Visible = True Then
        If lvObject.ListItems.Count > 0 Then
            If nValibleCount = 0 Then
                MsgBox "请扫描路单或选择打勾！", vbExclamation, Me.Caption
                SetNormal
                Exit Sub
            End If
        Else
            SetNormal
            Exit Sub
        End If
        tsStation.Tabs(1).Selected = True
        tsStation_Click
        cmdAddSheet.Visible = False
        cmdClear.Visible = False
        cmdFinish.Enabled = False
        cmdFinish.Visible = False
        cmdNext.Visible = True
        
        fraWizard(1).Visible = False
        fraWizard(2).Visible = False
        fraWizard(3).Visible = True
        fraWizard(4).Visible = False
        fraWizard(5).Visible = False
        fraWizard(6).Visible = False
        fraWizard(7).Visible = False
        fraWizard(8).Visible = False
        fraWizard(9).Visible = False
        
        
        SaveHeadWidth Me.name, lvObject
        '      填充路单数组
        FillSheetID
        '填充路单的站点统计信息
        
        vsStationDayList.Visible = False
        vsStationList.Visible = False
        vsStationTotal.Visible = True
        FillCheckSheetStationTotal
        '显示每日的路单统计
        FillCheckSheetStationDayList
        FillCheckSheetStationList
    ElseIf fraWizard(3).Visible = True Then
        '手工票信息录入
        
        FillVsExtra
        
        fraWizard(1).Visible = False
        fraWizard(2).Visible = False
        fraWizard(3).Visible = False
        fraWizard(4).Visible = True
        fraWizard(5).Visible = False
        fraWizard(6).Visible = False
        fraWizard(7).Visible = False
        fraWizard(8).Visible = False
        fraWizard(9).Visible = False
        
        cmdFinish.Enabled = False
        cmdFinish.Visible = False
        cmdNext.Visible = True

    ElseIf fraWizard(4).Visible = True Then
        If Not ValidateVsExtra Then
            SetNormal
            Exit Sub
        End If
        '填充路单预览信息
        FillPreSheetInfo

        
        
        fraWizard(1).Visible = False
        fraWizard(2).Visible = False
        fraWizard(3).Visible = False
        fraWizard(4).Visible = False
        fraWizard(5).Visible = True
        fraWizard(6).Visible = False
        fraWizard(7).Visible = False
        fraWizard(8).Visible = False
        fraWizard(9).Visible = False
        
        
        cmdFinish.Enabled = True
        cmdFinish.Visible = True
        cmdNext.Visible = False

    End If
    SetNormal
    Exit Sub
ErrHandle:
    SetNormal
    ShowErrorMsg
End Sub

'处理行包结算信息
Private Sub HandleLugInfo()
    lblSplitObject.Caption = ResolveDisplayEx(txtObject.Text)
    txtLugSheetID.Text = ""
    lvLugSheetID.ListItems.Clear
    cmdDetele.Enabled = False
    
End Sub

'处理行包结算统计
Private Sub StatLugInfo()
On Error GoTo here
    Dim rsTemp As Recordset
    Dim m_oReport As New Report
    Dim szaLugSheetID() As String
    Dim i As Integer
    lblLubObject.Caption = ""
    lblLugTotalPrice.Caption = ""
    lblLugNeedSplit.Caption = ""
    lblLugProtocol.Caption = ""
    
    If lvLugSheetID.ListItems.Count = 0 Then Exit Sub
    lvLugSheet.ListItems.Clear
    For i = 1 To lvLugSheetID.ListItems.Count
        lvLugSheet.ListItems.Add , , lvLugSheetID.ListItems.Item(i).Text
    Next i
    If lvLugSheet.ListItems.Count = 0 Then Exit Sub
    ReDim szaLugSheetID(1 To lvLugSheet.ListItems.Count)
    For i = 1 To lvLugSheet.ListItems.Count
        szaLugSheetID(i) = Trim(lvLugSheet.ListItems.Item(i).Text)
    Next i
    m_oReport.Init g_oActiveUser
    Set rsTemp = m_oReport.PreLugFinSheet(szaLugSheetID)
    If rsTemp.RecordCount > 0 Then
   
        lblLubObject.Caption = FormatDbValue(rsTemp!split_object_name)
        lblLugTotalPrice.Caption = FormatDbValue(rsTemp!total_price)
        lblLugNeedSplit.Caption = FormatDbValue(rsTemp!need_split_out)
        lblLugProtocol.Caption = FormatDbValue(rsTemp!protocol_name)
    End If
    
    Exit Sub
here:
    ShowErrorMsg
End Sub

'填充要拆算的路单记录
Private Sub FilllvObject()
    On Error GoTo ErrHandle
    Dim i As Integer
    Dim rsTemp As Recordset
    Dim nCount As Integer
    Dim lvItem As ListItem
    Dim m_oReport As New Report
    Dim aszTemp() As String
    SetBusy
    
    m_oReport.Init g_oActiveUser
    '填充路单信息
    If imgcbo.ComboItems(cnCompany).Selected Then      '需转化
        Set rsTemp = m_oReport.GetNeedSplitCheckSheet(CS_SettleByTransportCompany, ResolveDisplay(Trim(txtObject.Text)), dtpStartDate.Value, dtpEndDate.Value, ResolveDisplay(txtRouteID.Text), aszTemp, , IIf(chkIsToday.Value = vbChecked, True, False))
    ElseIf imgcbo.ComboItems(cnVehicle).Selected Then
        Set rsTemp = m_oReport.GetNeedSplitCheckSheet(CS_SettleByVehicle, ResolveDisplay(Trim(txtObject.Text)), dtpStartDate.Value, dtpEndDate.Value, ResolveDisplay(txtRouteID.Text), aszTemp, , IIf(chkIsToday.Value = vbChecked, True, False))
    ElseIf imgcbo.ComboItems(cnBus).Selected Then
        Set rsTemp = m_oReport.GetNeedSplitCheckSheet(CS_SettleByBus, ResolveDisplay(Trim(txtObject.Text)), dtpStartDate.Value, dtpEndDate.Value, ResolveDisplay(txtRouteID.Text), aszTemp, ResolveDisplay(cboCompany.Text), IIf(chkIsToday.Value = vbChecked, True, False))
        
    End If
    FilllvObjectHead '填充列首
    nCount = rsTemp.RecordCount
    
    If nCount = 0 Then
        SetNormal
        Exit Sub
    End If
    pbFill.Visible = True
    pbFill.Max = nCount
    lvObject.ListItems.Clear
    
    For i = 1 To nCount
        Set lvItem = lvObject.ListItems.Add(, , FormatDbValue(rsTemp!check_sheet_id))
        lvItem.SubItems(PI_Station) = FormatDbValue(rsTemp!station_name)
        lvItem.SubItems(PI_CheckGate) = FormatDbValue(rsTemp!check_gate_id)
        lvItem.SubItems(PI_Route) = FormatDbValue(rsTemp!route_name)
        lvItem.SubItems(PI_BusID) = FormatDbValue(rsTemp!bus_id)
        lvItem.SubItems(PI_BusSerialNO) = FormatDbValue(rsTemp!bus_serial_no)
        lvItem.SubItems(PI_LicenseTagNo) = FormatDbValue(rsTemp!license_tag_no)
        lvItem.SubItems(PI_CompanyName) = FormatDbValue(rsTemp!transport_company_short_name)
        lvItem.SubItems(PI_Quantity) = FormatDbValue(rsTemp!Quantity)
        lvItem.SubItems(PI_Mileage) = FormatDbValue(rsTemp!Mileage)    '里程
        lvItem.SubItems(PI_TicketPrice) = FormatDbValue(rsTemp!ticket_price)
        lvItem.SubItems(PI_Date) = Format(FormatDbValue(rsTemp!bus_date), "yy-mm-dd")
        lvItem.SubItems(PI_VehicleID) = FormatDbValue(rsTemp!vehicle_id)
        lvItem.SubItems(PI_CompanyID) = FormatDbValue(rsTemp!transport_company_id)
        m_szTransportCompanyID = FormatDbValue(rsTemp!transport_company_id)
        m_szTransportCompanyName = FormatDbValue(rsTemp!transport_company_short_name)
        pbFill.Value = i
        rsTemp.MoveNext
    Next i
    
    cmdAddSheet.Visible = True
    cmdClear.Visible = True
    pbFill.Visible = False
    SetNormal
    
    Exit Sub
ErrHandle:
    SetNormal
    ShowErrorMsg

End Sub

' 填充列首
Private Sub FilllvObjectHead()

   With lvObject.ColumnHeaders
         .Clear
         .Add , , "路单号", 1200
         .Add , , "终站点", 900
         .Add , , "车次", 800
         .Add , , "日期", 700
         .Add , , "序号", 700
         .Add , , "总票价", 800
         .Add , , "车牌号", 900
         .Add , , "参运公司", 900
         .Add , , "人数", 800
         .Add , , "里程数", 800
         .Add , , "线路", 900
         .Add , , "检票口", 0
         .Add , , "车辆代码", 0
         .Add , , "参运公司代码", 0
   End With
       AlignHeadWidth Me.name, lvObject
End Sub
'填充路单,车辆,公司数组
Private Sub FillSheetID()
 On Error GoTo ErrHandle
   Dim i As Integer
   Dim j As Integer
   
   ReDim m_aszCheckSheetID(1 To nValibleCount)
   j = 1
   For i = 1 To lvObject.ListItems.Count
      If lvObject.ListItems.Item(i).Checked = True Then
         m_aszCheckSheetID(j) = Trim(lvObject.ListItems(i).Text)
         j = j + 1
      End If
   Next i
   If imgcbo.Text = cszCompanyName Then   '转换常量
        m_szCompanyID = ResolveDisplay(txtObject.Text)
        m_szVehicleID = ""
        m_szBusID = ""
   ElseIf imgcbo.Text = "车辆" Then
        m_szVehicleID = ResolveDisplay(txtObject.Text)
        m_szCompanyID = ""
        m_szBusID = ""
   ElseIf imgcbo.Text = "车次" Then
        m_szBusID = ResolveDisplay(txtObject.Text)
        m_szCompanyID = ""
        m_szVehicleID = ""
   End If
 Exit Sub
ErrHandle:
 ShowErrorMsg
End Sub
Private Sub cmdPrevious_Click()
    On Error GoTo ErrHandle
    pbFill.Visible = False
    If fraWizard(5).Visible = True Then

        cmdFinish.Visible = False
        cmdFinish.Enabled = False
        cmdNext.Visible = True
        cmdNext.Enabled = True
        cmdPrevious.Enabled = True
        fraWizard(1).Visible = False
        fraWizard(2).Visible = False
        fraWizard(3).Visible = False
        fraWizard(4).Visible = True
        fraWizard(5).Visible = False
        fraWizard(6).Visible = False
        fraWizard(7).Visible = False
        fraWizard(8).Visible = False
        fraWizard(9).Visible = False
        
    ElseIf fraWizard(4).Visible = True Then
        
        cmdAddSheet.Visible = False
        cmdClear.Visible = False
        
        cmdFinish.Visible = False
        cmdFinish.Enabled = False
        cmdNext.Visible = True
        cmdNext.Enabled = True
        cmdPrevious.Enabled = True
        fraWizard(1).Visible = False
        fraWizard(2).Visible = False
        fraWizard(3).Visible = True
        fraWizard(4).Visible = False
        fraWizard(5).Visible = False
        fraWizard(6).Visible = False
        fraWizard(7).Visible = False
        fraWizard(8).Visible = False
        fraWizard(9).Visible = False
        
        txtAdditionPrice.Text = ""
    ElseIf fraWizard(3).Visible = True Then
        
        cmdAddSheet.Visible = True
        cmdClear.Visible = True
        
        cmdFinish.Visible = False
        cmdFinish.Enabled = False
        cmdNext.Visible = True
        cmdNext.Enabled = True
        cmdPrevious.Enabled = True
        fraWizard(1).Visible = False
        fraWizard(2).Visible = True
        fraWizard(3).Visible = False
        fraWizard(4).Visible = False
        fraWizard(5).Visible = False
        fraWizard(6).Visible = False
        fraWizard(7).Visible = False
        fraWizard(8).Visible = False
        fraWizard(9).Visible = False
        
    ElseIf fraWizard(2).Visible = True Then
    
        cmdAddSheet.Visible = False
        cmdClear.Visible = False
        
        cmdFinish.Visible = False
        cmdFinish.Enabled = False
        cmdNext.Visible = True
        cmdNext.Enabled = True
        cmdPrevious.Enabled = False
        fraWizard(1).Visible = True
        fraWizard(2).Visible = False
        fraWizard(3).Visible = False
        fraWizard(4).Visible = False
        fraWizard(5).Visible = False
        fraWizard(6).Visible = False
        fraWizard(7).Visible = False
        fraWizard(8).Visible = False
        fraWizard(9).Visible = False
        
    
    End If
    
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete And (ActiveControl Is lvLugSheetID) Then
        If lvLugSheetID.ListItems.Count = 0 Then
            cmdDetele.Enabled = False
            Exit Sub
        End If
        lvLugSheetID.ListItems.Remove (lvLugSheetID.SelectedItem.Index)
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()
    Dim nLen As Integer
    Dim szTemp() As String
    Dim i As Integer
    Dim szLastSheetID As String
    
    AlignFormPos Me
    cmdPrevious.Enabled = False
    cmdFinish.Visible = False
    cmdNext.Visible = True
    
    GetReg
    
    oPriceMan.Init g_oActiveUser
    '得到所有使用的票价项
    Set m_rsPriceItem = oPriceMan.GetAllTicketItemRS(TP_PriceItemUse)
    
    
    
    fraWizard(1).Visible = True
    fraWizard(2).Visible = False
    fraWizard(3).Visible = False
    fraWizard(4).Visible = False
    fraWizard(5).Visible = False
    fraWizard(6).Visible = False
    fraWizard(7).Visible = False
    fraWizard(8).Visible = False
    
    dtpStartDate.Value = Format(DateAdd("m", -1, Date), "yyyy-mm-01") 'GetFirstMonthDay(Date) & " 00:00:01"

    dtpEndDate.Value = DateAdd("d", -1, Format(Date, "yyyy-mm-01")) 'GetLastMonthDay(Date) & " 23:59:59"
    
    
    dtpSplitDate.Value = Now
    
    '  类型 初始化cbo对象
    '   1-车辆 2-参运公司
    With imgcbo
        .ComboItems.Clear
        .ComboItems.Add , , "参运公司", 1
        .ComboItems.Add , , "车辆", 3
        .ComboItems.Add , , "车次", 4
        .Locked = True
        If szSplitType <> "" Then
            If szSplitType = "参运公司" Then
                .ComboItems(1).Selected = True
            ElseIf szSplitType = "车辆" Then
                .ComboItems(2).Selected = True
            ElseIf szSplitType = "车次" Then
                .ComboItems(3).Selected = True
            End If
        Else
            .ComboItems(2).Selected = True
        End If
    End With
    
    '   自动生成结算单号 YYYYMM0001格式
    m_oSplit.Init g_oActiveUser
    szLastSheetID = m_oSplit.GetLastSettleSheetID
    If szLastSheetID = "0" Then
        txtSheetID.Text = CStr(Year(Now)) + CStr(IIf(Len(Month(Now)) = 2, Month(Now), 0 & Month(Now))) + "0001"
    Else
        txtSheetID.Text = szLastSheetID
    End If
    
End Sub

Private Sub imgcbo_Change()
    If imgcbo.ComboItems(cnBus).Selected Then
        '车次选中
        lblCompany.Visible = True
        cboCompany.Visible = True
        lblRoute.Visible = False
        txtRouteID.Visible = False
        lvBus.Visible = True
        fraBus.Visible = True
        lvVehicle.Visible = False
        fraVehicle.Visible = False
        lvCompany.Visible = False
        fraCompany.Visible = False
    ElseIf imgcbo.ComboItems(cnCompany).Selected Then
        '公司选中
        lblCompany.Visible = False
        cboCompany.Visible = False
        lblRoute.Visible = True
        txtRouteID.Visible = True
        lvCompany.Visible = True
        fraCompany.Visible = True
        lvBus.Visible = False
        fraBus.Visible = False
        lvVehicle.Visible = False
        fraVehicle.Visible = False
    ElseIf imgcbo.ComboItems(cnVehicle).Selected Then
        '车辆选中
        lblCompany.Visible = False
        cboCompany.Visible = False
        lblRoute.Visible = True
        txtRouteID.Visible = True
        lvVehicle.Visible = True
        fraVehicle.Visible = True
        lvCompany.Visible = False
        fraCompany.Visible = False
        lvBus.Visible = False
        fraBus.Visible = False
    End If
End Sub

Private Sub imgcbo_Click()
    imgcbo_Change
End Sub

Private Sub lstCreateInfo_DblClick()
 MsgBox lstCreateInfo.Text, vbInformation + vbOKOnly, "生成信息"
End Sub

Private Sub lvCompany_DblClick()
    Dim i As Integer
    Dim nIndex As Integer
    If lvCompany.ListItems.Count = 0 Then Exit Sub
    frmSplitItemModify.m_szObject = Trim(lvCompany.SelectedItem.Text)
    frmSplitItemModify.m_szProtocol = Trim(lvCompany.SelectedItem.SubItems(1))
    frmSplitItemModify.m_dbSettlePrice = Trim(lvCompany.SelectedItem.SubItems(2))
    frmSplitItemModify.m_dbStationPrice = Trim(lvCompany.SelectedItem.SubItems(3))
    frmSplitItemModify.m_nQuantity = Trim(lvCompany.SelectedItem.SubItems(4))
    
    frmSplitItemModify.m_nType = 1
'    frmSplitItemModify.ZOrder 0
    frmSplitItemModify.Show vbModal
                                                                                                           
    If frmSplitItemModify.m_bIsSave = False Then Exit Sub
    nIndex = lvCompany.SelectedItem.Index
    TSplitResult.CompanyInfo(nIndex).SettlePrice = 0
    TSplitResult.CompanyInfo(nIndex).SettleStationPrice = 0
    For i = 1 To g_cnSplitItemCount
        TSplitResult.CompanyInfo(nIndex).SplitItem(i) = m_adbSplitItem(i)
    
        If tSplitItem(i).SplitType = CS_SplitOtherCompany Then
            TSplitResult.CompanyInfo(nIndex).SettlePrice = TSplitResult.CompanyInfo(nIndex).SettlePrice + TSplitResult.CompanyInfo(nIndex).SplitItem(i)
        ElseIf tSplitItem(i).SplitType = CS_SplitStation Then
            TSplitResult.CompanyInfo(nIndex).SettleStationPrice = TSplitResult.CompanyInfo(nIndex).SettleStationPrice + TSplitResult.CompanyInfo(nIndex).SplitItem(i)
        End If
    Next i
    TSplitResult.SettleSheetInfo.SettleOtherCompanyPrice = 0
    TSplitResult.SettleSheetInfo.SettleStationPrice = 0
    For i = 1 To ArrayLength(TSplitResult.CompanyInfo)
        TSplitResult.SettleSheetInfo.SettleOtherCompanyPrice = TSplitResult.SettleSheetInfo.SettleOtherCompanyPrice + TSplitResult.CompanyInfo(i).SettlePrice
        TSplitResult.SettleSheetInfo.SettleStationPrice = TSplitResult.SettleSheetInfo.SettleStationPrice + TSplitResult.CompanyInfo(i).SettleStationPrice
    Next i
    lblNeedSplitMoney.Caption = FormatMoney(TSplitResult.CompanyInfo(nIndex).SettlePrice - TSplitResult.CompanyInfo(nIndex).SettleStationPrice)
    FilllvCompany
End Sub

Private Sub lvLugSheet_DblClick()
    ShowLugSheetInfo
End Sub

'显示行包结算单属性
Private Sub ShowLugSheetInfo()
    Dim oCommDialog As New STShell.CommDialog
    On Error GoTo ErrorHandle
    oCommDialog.Init g_oActiveUser
    If Not lvLugSheet.SelectedItem Is Nothing Then
        oCommDialog.ShowLugFinSheet lvLugSheet.SelectedItem.Text, Trim(lvLugSheet.SelectedItem.Text)
    ElseIf Not lvLugSheetID.SelectedItem Is Nothing Then
        oCommDialog.ShowLugFinSheet lvLugSheetID.SelectedItem.Text, Trim(lvLugSheetID.SelectedItem.Text)
    End If
    Set oCommDialog = Nothing
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub lvLugSheetID_DblClick()
    ShowLugSheetInfo
End Sub

Private Sub lvobject_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    SortListView lvObject, ColumnHeader.Index
End Sub

Private Sub lvObject_DblClick()
    Dim oCommDialog As New STShell.CommDialog
'    On Error Resume Next
    oCommDialog.Init g_oActiveUser
    oCommDialog.ShowCheckSheet lvObject.SelectedItem.Text
    Set oCommDialog = Nothing
End Sub

Private Sub lvObject_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.Checked = True Then
        DisplayAdd Item.Index, True
    Else
        DisplayAdd Item.Index, False
    End If
End Sub

Private Sub lvObject_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu pmnu_Select
    End If
End Sub

Private Sub lvVehicle_DblClick()
    Dim nIndex As Integer
    Dim i As Integer
    If lvVehicle.ListItems.Count = 0 Then Exit Sub
    frmSplitItemModify.m_szObject = Trim(lvVehicle.SelectedItem.Text)
    frmSplitItemModify.m_szProtocol = Trim(lvVehicle.SelectedItem.SubItems(1))
    frmSplitItemModify.m_dbSettlePrice = Trim(lvVehicle.SelectedItem.SubItems(2))
    frmSplitItemModify.m_dbStationPrice = Trim(lvVehicle.SelectedItem.SubItems(3))
    frmSplitItemModify.m_nQuantity = Trim(lvVehicle.SelectedItem.SubItems(4))
    
    frmSplitItemModify.m_nType = 1
    
    frmSplitItemModify.Show vbModal
    
    If frmSplitItemModify.m_bIsSave = False Then Exit Sub
    nIndex = lvVehicle.SelectedItem.Index
    TSplitResult.VehicleInfo(nIndex).SettlePrice = 0
    TSplitResult.VehicleInfo(nIndex).SettleStationPrice = 0
    For i = 1 To g_cnSplitItemCount
        TSplitResult.VehicleInfo(nIndex).SplitItem(i) = m_adbSplitItem(i)
    
        If tSplitItem(i).SplitType = CS_SplitOtherCompany Then
            TSplitResult.VehicleInfo(nIndex).SettlePrice = TSplitResult.VehicleInfo(nIndex).SettlePrice + TSplitResult.VehicleInfo(nIndex).SplitItem(i)
        ElseIf tSplitItem(i).SplitType = CS_SplitStation Then
            TSplitResult.VehicleInfo(nIndex).SettleStationPrice = TSplitResult.VehicleInfo(nIndex).SettleStationPrice + TSplitResult.VehicleInfo(nIndex).SplitItem(i)
        End If
    Next i
    TSplitResult.SettleSheetInfo.SettleOtherCompanyPrice = TSplitResult.VehicleInfo(nIndex).SettlePrice
    TSplitResult.SettleSheetInfo.SettleStationPrice = TSplitResult.VehicleInfo(nIndex).SettleStationPrice
    lblNeedSplitMoney.Caption = FormatMoney(TSplitResult.VehicleInfo(nIndex).SettlePrice - TSplitResult.VehicleInfo(nIndex).SettleStationPrice)
    FilllvVehicle
End Sub

Private Sub lvBus_DblClick()
    Dim nIndex As Integer
    Dim i As Integer
    If lvBus.ListItems.Count = 0 Then Exit Sub
    frmSplitItemModify.m_szObject = Trim(lvBus.SelectedItem.Text)
    frmSplitItemModify.m_szProtocol = Trim(lvBus.SelectedItem.SubItems(1))
    frmSplitItemModify.m_dbSettlePrice = Trim(lvBus.SelectedItem.SubItems(2))
    frmSplitItemModify.m_dbStationPrice = Trim(lvBus.SelectedItem.SubItems(3))
    frmSplitItemModify.m_nQuantity = Trim(lvBus.SelectedItem.SubItems(4))
    
    frmSplitItemModify.m_nType = 1
    
    frmSplitItemModify.Show vbModal
    
    If frmSplitItemModify.m_bIsSave = False Then Exit Sub
    nIndex = lvBus.SelectedItem.Index
    TSplitResult.BusInfo(nIndex).SettlePrice = 0
    TSplitResult.BusInfo(nIndex).SettleStationPrice = 0
    For i = 1 To g_cnSplitItemCount
        TSplitResult.BusInfo(nIndex).SplitItem(i) = m_adbSplitItem(i)
    
        If tSplitItem(i).SplitType = CS_SplitOtherCompany Then
            TSplitResult.BusInfo(nIndex).SettlePrice = TSplitResult.BusInfo(nIndex).SettlePrice + TSplitResult.BusInfo(nIndex).SplitItem(i)
        ElseIf tSplitItem(i).SplitType = CS_SplitStation Then
            TSplitResult.BusInfo(nIndex).SettleStationPrice = TSplitResult.BusInfo(nIndex).SettleStationPrice + TSplitResult.BusInfo(nIndex).SplitItem(i)
        End If
    Next i
    TSplitResult.SettleSheetInfo.SettleOtherCompanyPrice = TSplitResult.BusInfo(nIndex).SettlePrice
    TSplitResult.SettleSheetInfo.SettleStationPrice = TSplitResult.BusInfo(nIndex).SettleStationPrice
    lblNeedSplitMoney.Caption = FormatMoney(TSplitResult.BusInfo(nIndex).SettlePrice - TSplitResult.BusInfo(nIndex).SettleStationPrice)
    FilllvBus
End Sub

Private Sub pmnu_AllSelect_Click()
    AllSelect
End Sub

Private Sub pmnu_AllUnSelect_Click()
    UnAllSelect
End Sub

Private Sub pmnu_SelectCheck_Click()
    SelectCheck
End Sub

Private Sub pmnu_UnSelectCheck_Click()
    UnSelectCheck
End Sub



Private Sub tsStation_Click()
    If tsStation.Tabs(1).Selected Then
        '显示汇总信息
        vsStationTotal.Visible = True
        vsStationDayList.Visible = False
        vsStationList.Visible = False
        
    ElseIf tsStation.Tabs(2).Selected Then
        
        '显示明细信息
        vsStationTotal.Visible = False
        vsStationDayList.Visible = True
        vsStationList.Visible = False
    ElseIf tsStation.Tabs(3).Selected Then
        
        '显示明细信息
        vsStationTotal.Visible = False
        vsStationDayList.Visible = False
        vsStationList.Visible = True
    End If
End Sub

Private Sub txtAdditionPrice_Change()
Dim i As Integer
   If imgcbo.Text = cszCompanyName Then   '转换常量
        For i = 1 To lvCompany.ColumnHeaders.Count
            If Trim(tSplitItem(m_oSplit.m_nServiceItem).SplitItemName) = Trim(lvCompany.ColumnHeaders.Item(i)) Then
                lvCompany.SelectedItem.SubItems(i - 1) = AdditionPrice + Val(txtAdditionPrice.Text)
            End If
        Next i
        lvCompany.SelectedItem.SubItems(3) = TSplitResult.SettleSheetInfo.SettleStationPrice + Val(txtAdditionPrice.Text)
        lvCompany.SelectedItem.SubItems(2) = TSplitResult.BusInfo(1).SettlePrice - lvCompany.SelectedItem.SubItems(3)
        lblNeedSplitMoney.Caption = lvCompany.SelectedItem.SubItems(2)
   ElseIf imgcbo.Text = "车辆" Then
        For i = 1 To lvVehicle.ColumnHeaders.Count
            If Trim(tSplitItem(m_oSplit.m_nServiceItem).SplitItemName) = Trim(lvVehicle.ColumnHeaders.Item(i)) Then
                lvVehicle.SelectedItem.SubItems(i - 1) = AdditionPrice + Val(txtAdditionPrice.Text)
            End If
        Next i
        lvVehicle.SelectedItem.SubItems(3) = TSplitResult.SettleSheetInfo.SettleStationPrice + Val(txtAdditionPrice.Text)
        lvVehicle.SelectedItem.SubItems(2) = TSplitResult.BusInfo(1).SettlePrice - lvVehicle.SelectedItem.SubItems(3)
        lblNeedSplitMoney.Caption = lvVehicle.SelectedItem.SubItems(2)
   ElseIf imgcbo.Text = "车次" Then
        For i = 1 To lvBus.ColumnHeaders.Count
            If Trim(tSplitItem(m_oSplit.m_nServiceItem).SplitItemName) = Trim(lvBus.ColumnHeaders.Item(i)) Then
                lvBus.SelectedItem.SubItems(i - 1) = AdditionPrice + Val(txtAdditionPrice.Text)
            End If
        Next i
        lvBus.SelectedItem.SubItems(3) = TSplitResult.SettleSheetInfo.SettleStationPrice + Val(txtAdditionPrice.Text)
        lvBus.SelectedItem.SubItems(2) = TSplitResult.BusInfo(1).SettlePrice - lvBus.SelectedItem.SubItems(3)
        lblNeedSplitMoney.Caption = lvBus.SelectedItem.SubItems(2)
   End If
End Sub

Private Sub SelectCompany()
'On Error GoTo ErrHandle
'    Dim oShell As New STShell.CommDialog
'    Dim aszTemp() As String
'    oShell.Init g_oActiveUser
'    aszTemp = oShell.SelectCompany
'    Set oShell = Nothing
'    If ArrayLength(aszTemp) = 0 Then Exit Sub
'    cboCompany.Text = MakeDisplayString(Trim(aszTemp(1, 1)), Trim(aszTemp(1, 2)))
'
'Exit Sub
'ErrHandle:
'ShowErrorMsg
End Sub

Private Sub txtLugSheetID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        AddLugSheet
    End If
End Sub

Private Sub txtObject_ButtonClick()
    On Error GoTo ErrHandle
    Dim oShell As New STShell.CommDialog
    Dim aszTemp() As String
    oShell.Init g_oActiveUser
    If Trim(imgcbo.Text) = cszCompanyName Then
        aszTemp = oShell.SelectCompany()
    ElseIf Trim(imgcbo.Text) = "车辆" Then
        aszTemp = oShell.SelectVehicleEX()
    ElseIf Trim(imgcbo.Text) = "车次" Then
        aszTemp = oShell.SelectBus()
    End If
    
    Set oShell = Nothing
    If ArrayLength(aszTemp) = 0 Then Exit Sub
    
    If Trim(imgcbo.Text) = "车次" Then
        txtObject.Text = Trim(aszTemp(1, 1))
    Else
        txtObject.Text = MakeDisplayString(Trim(aszTemp(1, 1)), Trim(aszTemp(1, 2)))
    End If
    
    Exit Sub
    
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveFormPos Me
    SaveHeadWidth Me.name, lvObject
    SaveHeadWidth Me.name, lvCompany
    SaveHeadWidth Me.name, lvVehicle
    SaveHeadWidth Me.name, lvBus
    SaveHeadWidth Me.name, vsStationTotal
    SaveHeadWidth Me.name, vsStationDayList
    SaveHeadWidth Me.name, vsStationList
    
    Unload Me
       
End Sub


Private Sub cmdCancel_Click()
    Unload Me


End Sub




Private Sub FillTOlvObject()
On Error GoTo here
    Dim i As Integer
    Dim m_oCheckSheet As New CheckSheet
    Dim rsTemp As Recordset
    
    If lvObject.ListItems.Count = 0 Then Exit Sub
    
    For i = 1 To lvObject.ListItems.Count
        If lvObject.ListItems.Item(i).Text = Trim(txtCheckSheetID.Text) Then
            Exit For
        End If
    Next i
    If i > lvObject.ListItems.Count Then
        m_oCheckSheet.Init g_oActiveUser
        Set rsTemp = m_oCheckSheet.CheckSheetAvailable(Trim(txtCheckSheetID.Text))   '验证路单有效性,如果无效声音提示
        If rsTemp.RecordCount = 0 Then
            PlayEventSound g_tEventSoundPath.CheckSheetNotExist '此路单不存在
            MsgBox "此路单不存在", vbExclamation, Me.Caption
            imgEnabled.Visible = True
            Exit Sub
        ElseIf FormatDbValue(rsTemp!valid_mark) = 0 Then
            
        
            MsgBox "此路单已废", vbExclamation, Me.Caption
            PlayEventSound g_tEventSoundPath.CheckSheetCanceled  '此路单已废
            txtCheckSheetID.SelStart = 0
            txtCheckSheetID.SelLength = Len(txtCheckSheetID.Text)
            Exit Sub
        ElseIf FormatDbValue(rsTemp!settlement_status) = 1 Then
            MsgBox "此路单已结算", vbExclamation, Me.Caption
            PlayEventSound g_tEventSoundPath.CheckSheetSettled '此路单已结算
            txtCheckSheetID.SelStart = 0
            txtCheckSheetID.SelLength = Len(txtCheckSheetID.Text)
            Exit Sub
        Else
            PlayEventSound g_tEventSoundPath.ObjectNotSame   '此路单有效,但不在所要拆算时间或对象范围之内
            MsgBox "此路单有效,但不在所要拆算时间或对象范围之内", vbExclamation, Me.Caption
            txtCheckSheetID.SelStart = 0
            txtCheckSheetID.SelLength = Len(txtCheckSheetID.Text)
            Exit Sub
        End If
        
    Else
        If lvObject.ListItems(i).Checked Then
            PlayEventSound g_tEventSoundPath.CheckSheetSelected '此路单已选
'            MsgBox "此路单已选！", vbInformation, Me.Caption
            txtCheckSheetID.SelStart = 0
            txtCheckSheetID.SelLength = Len(txtCheckSheetID.Text)
            
            Exit Sub
        End If
        
        
    End If
    
    '扫描路单成功
    PlayEventSound g_tEventSoundPath.CheckSheetValid '有效路单
    '列表中打勾有效路单
    imgEnabled.Visible = False
    For i = 1 To lvObject.ListItems.Count
        If Trim(txtCheckSheetID.Text) = Trim(lvObject.ListItems.Item(i).Text) Then
            lvObject.ListItems.Item(i).EnsureVisible
            lvObject.ListItems.Item(i).Selected = True
            If lvObject.ListItems.Item(i).Checked = False Then
                lvObject.ListItems.Item(i).Checked = True
                    '扫描一张,有效总数加一
                    nValibleCount = nValibleCount + 1
                    '总人数累加
                nTotalQuantity = nTotalQuantity + CInt(lvObject.ListItems.Item(i).SubItems(PI_Quantity))
                lblTotalQuantity.Caption = nTotalQuantity
            End If
        End If
    Next i
    
    lblEnableCount.Caption = nValibleCount
    txtCheckSheetID.Text = ""
    txtCheckSheetID.SetFocus
    Exit Sub
here:
    ShowErrorMsg
    
End Sub

Private Sub txtObject_LostFocus()
    '如果对象为车次,则填充所有的该车次的公司
    Dim oBus As New Bus
    Dim nCount As Integer
    Dim aszCompany() As String
    Dim i As Integer
    On Error GoTo ErrorHandle
    If Trim(imgcbo.Text) = "车次" Then
        cboCompany.Clear
        cboCompany.AddItem ""
        oBus.Init g_oActiveUser
        oBus.Identify txtObject.Text
        aszCompany = oBus.GetAllCompany
        nCount = ArrayLength(aszCompany)
        For i = 1 To nCount
            cboCompany.AddItem MakeDisplayString(aszCompany(i, 1), aszCompany(i, 2))
        Next i
        
    End If
    Exit Sub
ErrorHandle:
    
End Sub

Private Sub txtRouteID_ButtonClick()
    SelectRoute
End Sub

Private Sub txtCheckSheetID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FillTOlvObject
    End If
End Sub

Private Sub txtSheetID_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
      txtOperator.SetFocus
   End If
End Sub


'填充填充路单预览信息
Public Sub FillPreSheetInfo()
    Dim i As Integer
    Dim nCount As Integer
    Dim lvItem As ListItem
    
    Dim atStationQuantity() As TSettleSheetStation
    
    
    Dim rsPassageNumber As Recordset
    
    
    
    
    '生成人数的记录集
    If imgcbo.Text = cszCompanyName Then
        '按公司结算
        Set rsPassageNumber = MakeRecordSetListByCompany
    ElseIf imgcbo.Text = cszVehicleName Then
        Set rsPassageNumber = MakeRecordSetListByVehicle
    ElseIf imgcbo.Text = cszBusName Then
        Set rsPassageNumber = MakeRecordSetListByBus
    End If
    
    
    '拆算预览
    m_oSplit.Init g_oActiveUser
    TSplitResult = m_oSplit.PreviewSplitCheckSheetEx(m_szCompanyID, m_szVehicleID, m_aszCheckSheetID, rsPassageNumber, dtpStartDate.Value, DateAdd("d", 1, dtpEndDate.Value), m_atExtraInfo, IIf(chkSingle.Value = vbChecked, True, False), m_szBusID)

    lblTotalPrice.Caption = IIf(TSplitResult.SettleSheetInfo.TotalTicketPrice = 0, "无", TSplitResult.SettleSheetInfo.TotalTicketPrice)
    
    
    lblNeedSplitMoney.Caption = Format(TSplitResult.SettleSheetInfo.SettleOtherCompanyPrice - TSplitResult.SettleSheetInfo.SettleStationPrice, "0.00")
    lblTotalQuautity.Caption = TSplitResult.SettleSheetInfo.TotalQuantity
    lblSheetCount.Caption = TSplitResult.SettleSheetInfo.CheckSheetCount
    
    '填充列首
    lvHead
    '填充公司列表
    FilllvCompany
    '填充车辆列表
    FilllvVehicle
    '填充车次列表
    FilllvBus


End Sub



Private Sub FilllvCompany()
On Error GoTo here
    Dim i As Integer
    Dim j As Integer
    Dim lvItem As ListItem
    Dim nCount As Integer
    Dim k As Integer
    nCount = ArrayLength(TSplitResult.CompanyInfo)
    lvCompany.ListItems.Clear
    If nCount = 0 Then Exit Sub
    
    For j = 1 To nCount
        Set lvItem = lvCompany.ListItems.Add(, , TSplitResult.CompanyInfo(j).CompanyName)
        lvItem.SubItems(1) = TSplitResult.CompanyInfo(j).ProtocolName
        lvItem.SubItems(2) = FormatMoney(TSplitResult.CompanyInfo(j).SettlePrice - TSplitResult.CompanyInfo(j).SettleStationPrice)
        lvItem.SubItems(3) = TSplitResult.CompanyInfo(j).SettleStationPrice
        lvItem.SubItems(4) = TSplitResult.CompanyInfo(j).PassengerNumber
        lvItem.SubItems(5) = TSplitResult.CompanyInfo(j).Mileage
        k = 5
        For i = 1 To g_cnSplitItemCount
            If tSplitItem(i).SplitStatus <> CS_SplitItemNotUse Then
                k = k + 1
                lvItem.SubItems(k) = TSplitResult.CompanyInfo(j).SplitItem(i) '拆算项处理
                                
            End If
        Next i
        For i = 1 To lvCompany.ColumnHeaders.Count
            If Trim(tSplitItem(m_oSplit.m_nServiceItem).SplitItemName) = Trim(lvCompany.ColumnHeaders.Item(i)) Then
                AdditionPrice = lvItem.SubItems(i - 1)
            End If
        Next i
    Next j
    
Exit Sub
here:
    ShowErrorMsg
End Sub

Private Sub FilllvVehicle()
On Error GoTo here
    Dim i As Integer
    Dim j As Integer
    Dim lvItem As ListItem
    Dim nCount As Integer
    Dim k As Integer
    nCount = ArrayLength(TSplitResult.VehicleInfo)
    lvVehicle.ListItems.Clear
    If nCount = 0 Then Exit Sub
    
    For j = 1 To nCount
        Set lvItem = lvVehicle.ListItems.Add(, , TSplitResult.VehicleInfo(j).LicenseTagNo)
        lvItem.SubItems(1) = TSplitResult.VehicleInfo(j).ProtocolName
        lvItem.SubItems(2) = FormatMoney(TSplitResult.VehicleInfo(j).SettlePrice - TSplitResult.VehicleInfo(j).SettleStationPrice)
        lvItem.SubItems(3) = TSplitResult.VehicleInfo(j).SettleStationPrice
        lvItem.SubItems(4) = TSplitResult.VehicleInfo(j).PassengerNumber
        lvItem.SubItems(5) = TSplitResult.VehicleInfo(j).Mileage
        k = 5
        For i = 1 To g_cnSplitItemCount
            If tSplitItem(i).SplitStatus <> CS_SplitItemNotUse Then
                k = k + 1
                lvItem.SubItems(k) = TSplitResult.VehicleInfo(j).SplitItem(i) '拆算项处理
            End If
        Next i
        
        For i = 1 To lvVehicle.ColumnHeaders.Count
            If Trim(tSplitItem(m_oSplit.m_nServiceItem).SplitItemName) = Trim(lvVehicle.ColumnHeaders.Item(i)) Then
                AdditionPrice = lvItem.SubItems(i - 1)
            End If
        Next i
    Next j
Exit Sub
here:
    ShowErrorMsg
End Sub

Private Sub FilllvBus()
On Error GoTo here
    Dim i As Integer
    Dim j As Integer
    Dim lvItem As ListItem
    Dim nCount As Integer
    Dim k As Integer
    nCount = ArrayLength(TSplitResult.BusInfo)
    lvBus.ListItems.Clear
    If nCount = 0 Then Exit Sub
    
    For j = 1 To nCount
        Set lvItem = lvBus.ListItems.Add(, , TSplitResult.BusInfo(j).BusID)
        lvItem.SubItems(1) = TSplitResult.BusInfo(j).ProtocolName
        lvItem.SubItems(2) = FormatMoney(TSplitResult.BusInfo(j).SettlePrice - TSplitResult.BusInfo(j).SettleStationPrice)
        lvItem.SubItems(3) = TSplitResult.BusInfo(j).SettleStationPrice
        lvItem.SubItems(4) = TSplitResult.BusInfo(j).PassengerNumber
        lvItem.SubItems(5) = TSplitResult.BusInfo(j).Mileage
        k = 5
        For i = 1 To g_cnSplitItemCount
            If tSplitItem(i).SplitStatus <> CS_SplitItemNotUse Then
                k = k + 1
                lvItem.SubItems(k) = TSplitResult.BusInfo(j).SplitItem(i) '拆算项处理
            End If
        Next i
        For i = 1 To lvBus.ColumnHeaders.Count
            If Trim(tSplitItem(m_oSplit.m_nServiceItem).SplitItemName) = Trim(lvBus.ColumnHeaders.Item(i)) Then
                AdditionPrice = lvItem.SubItems(i - 1)
            End If
        Next i
    Next j

Exit Sub
here:
    ShowErrorMsg
End Sub

Private Sub lvHead()
On Error GoTo here
Dim m_oReport As New Report

Dim nSplitItenCount As Integer
Dim i As Integer
    With lvCompany.ColumnHeaders
        .Clear
        .Add , , "参运公司"
        .Add , , "协议"
        .Add , , "应结票款"
        .Add , , "结给车站票款"
        .Add , , "人数"
        .Add , , "人公里"
        
    End With
    
    With lvVehicle.ColumnHeaders
        .Clear
        .Add , , "车辆"
        .Add , , "协议"
        .Add , , "应结票款"
        .Add , , "结给车站票款"
        .Add , , "人数"
        .Add , , "人公里"
        
    End With
    
    With lvBus.ColumnHeaders
        .Clear
        .Add , , "车次"
        .Add , , "协议"
        .Add , , "应结票款"
        .Add , , "结给车站票款"
        .Add , , "人数"
        .Add , , "人公里"
        
    End With
    

    '取得使用的拆算款项
    m_oReport.Init g_oActiveUser
    tSplitItem = m_oReport.GetSplitItemInfo()
    nSplitItenCount = ArrayLength(tSplitItem)
    m_nSplitItenCount = 0
    If nSplitItenCount = 0 Then Exit Sub
    For i = 1 To nSplitItenCount
        If tSplitItem(i).SplitStatus <> CS_SplitItemNotUse Then
            lvCompany.ColumnHeaders.Add , , tSplitItem(i).SplitItemName
            lvVehicle.ColumnHeaders.Add , , tSplitItem(i).SplitItemName '同时增加车辆列表的拆算项
            lvBus.ColumnHeaders.Add , , tSplitItem(i).SplitItemName '同时增加车次列表的拆算项
            m_nSplitItenCount = m_nSplitItenCount + 1
        End If
    Next i
    
    
    AlignHeadWidth Me.name, lvCompany
    AlignHeadWidth Me.name, lvVehicle
    AlignHeadWidth Me.name, lvBus
    Exit Sub
here:
    ShowErrorMsg
End Sub

Private Sub SelectCheck()
    On Error GoTo ErrHandle
    Dim i As Integer
    If lvObject.ListItems.Count = 0 Then Exit Sub
    For i = 1 To lvObject.ListItems.Count
        If lvObject.ListItems(i).Selected Then
            If lvObject.ListItems.Item(i).Checked = False Then
                lvObject.ListItems.Item(i).Checked = True
                DisplayAdd i, True
            End If
        End If
    Next i
    
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub


Private Sub UnSelectCheck()
    On Error GoTo ErrHandle
    Dim i As Integer
    If lvObject.ListItems.Count = 0 Then Exit Sub
    For i = 1 To lvObject.ListItems.Count
        If lvObject.ListItems(i).Selected Then
            If lvObject.ListItems.Item(i).Checked = True Then
                lvObject.ListItems.Item(i).Checked = False
                DisplayAdd i, False
            End If
        End If
    Next i
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub


Private Sub UnAllSelect()
    On Error GoTo ErrHandle
    Dim i As Integer
    If lvObject.ListItems.Count = 0 Then Exit Sub
    For i = 1 To lvObject.ListItems.Count
        If lvObject.ListItems.Item(i).Checked = True Then
            lvObject.ListItems.Item(i).Checked = False
            DisplayAdd i, False
        End If
    Next i
    
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub AllSelect()
    On Error GoTo ErrHandle
    Dim i As Integer
    If lvObject.ListItems.Count = 0 Then Exit Sub
    For i = 1 To lvObject.ListItems.Count
        If lvObject.ListItems.Item(i).Checked = False Then
            lvObject.ListItems.Item(i).Checked = True
            DisplayAdd i, True
        End If
    Next i
    
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub DisplayAdd(pnIndex As Integer, pbIsAdd As Boolean)
    If pbIsAdd Then
        nValibleCount = nValibleCount + 1
        nTotalQuantity = nTotalQuantity + CInt(lvObject.ListItems(pnIndex).SubItems(PI_Quantity))
    Else
        
        nTotalQuantity = IIf(nTotalQuantity - CInt(lvObject.ListItems(pnIndex).SubItems(PI_Quantity)) >= 0, nTotalQuantity - CInt(lvObject.ListItems(pnIndex).SubItems(PI_Quantity)), 0)
        nValibleCount = IIf(nValibleCount - 1 >= 0, nValibleCount - 1, 0)
    End If
    lblEnableCount.Caption = nValibleCount
    lblTotalQuantity.Caption = nTotalQuantity
    '显示路单数
    m_nSheetCount = lvObject.ListItems.Count
    lblSettleSheetCount.Caption = CStr(m_nSheetCount)
End Sub





Private Sub SelectRoute()
On Error GoTo ErrHandle
    Dim oShell As New STShell.CommDialog
    Dim aszTemp() As String
    oShell.Init g_oActiveUser
       aszTemp = oShell.SelectRoute()

    Set oShell = Nothing
    If ArrayLength(aszTemp) = 0 Then Exit Sub
    txtRouteID.Text = MakeDisplayString(Trim(aszTemp(1, 1)), Trim(aszTemp(1, 2)))

Exit Sub
ErrHandle:
ShowErrorMsg
End Sub





Private Function MakeRecordSetDayList(prsData As Recordset, prsStation As Recordset) As Recordset
    '手工生成记录集
    'nStationCount 为总共有的站点数目.
    
    Dim rsTemp As New Recordset
    Dim i As Integer
    Dim j As Integer
    
    
    Dim aszStation() As String
    Dim nStationCount As Integer '站点票种数
    
    Dim szBusDate As String
    Dim szStationID As String '站点
'    Dim szStationName As String '站名
    Dim szTicketTypeID As String '票种
'    Dim szTicketTypeName As String '票种名
    Dim alNumber() As Long '人数
    Dim lTotalNum As Long
    nStationCount = prsStation.RecordCount
    '存放字段的名称,为比较用.
    ReDim aszStation(1 To nStationCount, 1 To 2)
    ReDim alNumber(1 To nStationCount)
    prsStation.MoveFirst
    For i = 1 To nStationCount
        aszStation(i, 1) = FormatDbValue(prsStation!station_id)
        aszStation(i, 2) = FormatDbValue(prsStation!ticket_type)
        prsStation.MoveNext
    Next i
    
    
    
    With rsTemp.Fields
        .Append "bus_date", adChar, 10
        For i = 1 To nStationCount
            .Append "station_" & i, adBigInt
        Next i
        .Append "total_num", adBigInt
        
    End With
    rsTemp.Open
    If prsData Is Nothing Then Exit Function
    If prsData.RecordCount = 0 Then Exit Function
    prsData.MoveFirst
    
    
    szBusDate = FormatDbValue(prsData!bus_date)
    szStationID = FormatDbValue(prsData!station_id)
'    szStationName = FormatDbValue(prsData!station_name)
    
    szTicketTypeID = FormatDbValue(prsData!ticket_type)
'    szTicketTypeName = FormatDbValue(prsData!ticket_type_name)
    
    
    
    For i = 1 To prsData.RecordCount
        If szBusDate <> FormatDbValue(prsData!bus_date) Then
            '赋予记录集
            
            rsTemp.AddNew
            rsTemp!bus_date = szBusDate
            lTotalNum = 0
            For j = 1 To nStationCount
                rsTemp.Fields("station_" & j) = alNumber(j)
                lTotalNum = lTotalNum + alNumber(j)
                
            Next j
            rsTemp.Fields("total_num") = lTotalNum
            
            rsTemp.Update
            
            '清空原值
            For j = 1 To nStationCount
                alNumber(j) = 0
            Next j

            '赋该车次的初始值
                    
            szBusDate = FormatDbValue(prsData!bus_date)
            szStationID = FormatDbValue(prsData!station_id)
        '    szStationName = FormatDbValue(prsData!station_name)
            
            szTicketTypeID = FormatDbValue(prsData!ticket_type)
        '    szTicketTypeName = FormatDbValue(prsData!ticket_type_name)
            For j = 1 To nStationCount
                If FormatDbValue(prsData!station_id) = aszStation(j, 1) And FormatDbValue(prsData!ticket_type) = aszStation(j, 2) Then
                    alNumber(j) = alNumber(j) + FormatDbValue(prsData!Quantity)
                    Exit For
                End If
            Next j
            

        Else
            '如果不同
            
            For j = 1 To nStationCount
                If FormatDbValue(prsData!station_id) = aszStation(j, 1) And FormatDbValue(prsData!ticket_type) = aszStation(j, 2) Then
                    alNumber(j) = alNumber(j) + FormatDbValue(prsData!Quantity)
                    Exit For
                End If
            Next j
        End If
        prsData.MoveNext
    Next i


    rsTemp.AddNew
    rsTemp!bus_date = szBusDate
    lTotalNum = 0
    For j = 1 To nStationCount
        rsTemp.Fields("station_" & j) = alNumber(j)
        lTotalNum = lTotalNum + alNumber(j)
    Next j
    rsTemp.Fields("total_num") = lTotalNum

    rsTemp.Update
            
    Set MakeRecordSetDayList = rsTemp
    
End Function


Private Sub vsStationList_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    

    Dim i As Integer
    Dim j As Integer
    
    
    
    With vsStationList
'        'debug.Print "new value: " & vsStationList.Text
        If Col = VIL_Quantity Then
        
            If vsStationList.Text <> m_nQuantity Then
            
                .CellBackColor = vbRed
                '修改站点汇总的人数
                For i = 1 To vsStationTotal.Rows - 1
                    '找到对就的行,将汇总人数进行修改
                    If vsStationTotal.TextMatrix(i, VI_SellStationID) = .TextMatrix(Row, VIL_SellStationID) _
                            And vsStationTotal.TextMatrix(i, VI_RouteID) = .TextMatrix(Row, VIL_RouteID) _
                            And vsStationTotal.TextMatrix(i, VI_VehicleTypeID) = .TextMatrix(Row, VIL_VehicleTypeID) _
                            And vsStationTotal.TextMatrix(i, VI_StationID) = .TextMatrix(Row, VIL_StationID) _
                            And vsStationTotal.TextMatrix(i, VI_TicketTypeID) = .TextMatrix(Row, VIL_TicketType) _
                            Then
                            
                        '找到了,将数量进行调整
                        vsStationTotal.TextMatrix(i, VI_Quantity) = Val(vsStationTotal.TextMatrix(i, VI_Quantity)) + Val(.Text) - Val(m_nQuantity)
                        vsStationTotal.Row = i
                        vsStationTotal.Col = VI_Quantity
                        vsStationTotal.CellBackColor = vbRed
                        '修改总价
                        .TextMatrix(Row, VIL_TotalTicketPrice) = Val(.TextMatrix(Row, VIL_Quantity)) * Val(.TextMatrix(Row, VIL_TicketPrice))
                        Exit For
                    End If
                Next i
                
                
                '修改站点每日汇总的人数
                For i = 1 To vsStationDayList.Rows - 1
                    If ToDBDate(vsStationDayList.TextMatrix(i, 0)) = ToDBDate(.TextMatrix(Row, VIL_BusDate)) Then
                        '从记录集中查找出站点代码及票种来进行比较
                        m_rsStationInfo.MoveFirst
                        For j = 1 To m_rsStationInfo.RecordCount
                            If FormatDbValue(m_rsStationInfo!station_id) = .TextMatrix(Row, VIL_StationID) And FormatDbValue(m_rsStationInfo!ticket_type) = .TextMatrix(Row, VIL_TicketType) Then
                                vsStationDayList.TextMatrix(i, j) = vsStationDayList.TextMatrix(i, j) + Val(.Text) - Val(m_nQuantity)
                                vsStationDayList.Row = i
                                vsStationDayList.Col = j
                                vsStationDayList.CellBackColor = vbRed
                                
                                Exit For
                                
                            End If
                            m_rsStationInfo.MoveNext
                        Next j
                    End If
                Next i
                
            End If
        End If
    End With
    
End Sub

Private Sub vsStationList_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '保存修改前的人数
    If Col = VIL_Quantity Then
        m_nQuantity = vsStationList.Text
    End If
End Sub

Private Sub vsStationTotal_BeforeSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long, Cancel As Boolean)
    '汇总处不允许修改人数
    
    
'    If NewColSel = VI_Quantity Then
'        '如果为数量,则允许修改
''        If m_bIsManualSettle Then
'            vsStationTotal.Editable = flexEDKbdMouse
''        Else
''            vsStationTotal.Editable = flexEDNone
''        End If
'    Else
'        vsStationTotal.Editable = flexEDNone
'    End If
End Sub

Private Sub vsStationList_BeforeSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long, Cancel As Boolean)
'    If NewColSel = VIL_Quantity Then
'        '如果为数量,则允许修改
'        vsStationList.Editable = flexEDKbdMouse
'    Else
'        vsStationList.Editable = flexEDNone
'    End If
End Sub

Private Function QuantityHasBeChanged() As Boolean
    '站点人数信息是否被改变过
    Dim i As Integer
    vsStationTotal.Col = 1
    m_rsStationQuantity.MoveFirst
    For i = 1 To m_rsStationQuantity.RecordCount
        If vsStationTotal.TextMatrix(i, VI_Quantity) <> FormatDbValue(m_rsStationQuantity!Quantity) Then
            Exit For
        End If
        m_rsStationQuantity.MoveNext
    Next i
    If i > m_rsStationQuantity.RecordCount Then
        '未被改变过
        QuantityHasBeChanged = False
    Else
        '被改动过
        QuantityHasBeChanged = True
    End If
    
End Function

Private Function GetStationQuantity() As TSettleSheetStation()
    '得到vsStationTotal中的站点人数信息
    Dim atStationQuantity() As TSettleSheetStation
    Dim i As Integer
    ReDim atStationQuantity(1 To m_rsStationQuantity.RecordCount)
    
    '汇总站点人数
    '********此处应加入把不同的车次汇总成同一个.
    m_rsStationQuantity.MoveFirst
    For i = 1 To m_rsStationQuantity.RecordCount
        atStationQuantity(i).RouteID = FormatDbValue(m_rsStationQuantity!route_id)
        atStationQuantity(i).RouteName = FormatDbValue(m_rsStationQuantity!route_name)
        atStationQuantity(i).SellSationID = FormatDbValue(m_rsStationQuantity!sell_station_id)
        atStationQuantity(i).SellStationName = FormatDbValue(m_rsStationQuantity!sell_station_name)
        atStationQuantity(i).StationID = FormatDbValue(m_rsStationQuantity!station_id)
        atStationQuantity(i).StationName = FormatDbValue(m_rsStationQuantity!station_name)
        atStationQuantity(i).TicketType = FormatDbValue(m_rsStationQuantity!ticket_type)
        atStationQuantity(i).TicketTypeName = FormatDbValue(m_rsStationQuantity!ticket_type_name)
        atStationQuantity(i).VehicleTypeCode = FormatDbValue(m_rsStationQuantity!vehicle_type_code)
        atStationQuantity(i).VehicleTypeName = FormatDbValue(m_rsStationQuantity!vehicle_type_name)
        atStationQuantity(i).Quantity = vsStationTotal.TextMatrix(i, VI_Quantity)
        atStationQuantity(i).AreaRatio = vsStationTotal.TextMatrix(i, VI_AreaRatio)
        m_rsStationQuantity.MoveNext
    Next i
    GetStationQuantity = atStationQuantity
    
End Function




Private Sub FillCheckSheetStationTotal()
    '填充路单站点汇总信息
    Dim oSplit As New Split
'    Dim m_rsStationQuantity As Recordset
    On Error GoTo ErrorHandle
    Dim i As Integer
    Dim nCount As Integer
    Dim j As Integer
    SetBusy
    ShowSBInfo "正在填充路单站点汇总信息"
'路单站点汇总的列首位置
    
    
    With vsStationTotal
        .Clear
        .Rows = 2
        .Cols = 14
        .FixedCols = 1
        .FixedRows = 1
        .ColWidth(0) = 100
        '设置合并
        .MergeCells = flexMergeRestrictColumns
        .MergeCol(VI_SellStation) = True
        .MergeCol(VI_Route) = True
        .MergeCol(VI_Bus) = True
        .MergeCol(VI_VehicleType) = True
        .MergeCol(VI_Station) = True
        .MergeCol(VI_TicketType) = False
        
        
        .AllowUserResizing = flexResizeColumns
        
        
        .TextMatrix(0, VI_SellStation) = "上车站"
        .TextMatrix(0, VI_Route) = "线路"
        .TextMatrix(0, VI_Bus) = "车次"
        .TextMatrix(0, VI_VehicleType) = "车型"
        .TextMatrix(0, VI_Station) = "站点"
        .TextMatrix(0, VI_TicketType) = "票种"
        .TextMatrix(0, VI_Quantity) = "人数"
        .TextMatrix(0, VI_AreaRatio) = "区域费率"
        
        '下面的信息纯粹只在汇总时用到
        .TextMatrix(0, VI_RouteID) = "线路代码"
        .TextMatrix(0, VI_VehicleTypeID) = "车型代码"
        .TextMatrix(0, VI_SellStationID) = "上车站代码"
        .TextMatrix(0, VI_StationID) = "站点代码"
        .TextMatrix(0, VI_TicketTypeID) = "票种代码"
        
        AlignHeadWidth Me.name, vsStationTotal
        .ColWidth(VI_AreaRatio) = 0
        .ColWidth(VI_Bus) = 0
        '下面的信息纯粹只在汇总时用到
        .ColWidth(VI_RouteID) = 0
        .ColWidth(VI_VehicleTypeID) = 0
        .ColWidth(VI_SellStationID) = 0
        .ColWidth(VI_StationID) = 0
        .ColWidth(VI_TicketTypeID) = 0
    End With
    oSplit.Init g_oActiveUser
    FillSheetID
    'debug.Print "totalCheckSheetStationInfo start:" & Time
    Set m_rsStationQuantity = oSplit.TotalCheckSheetStationInfo(m_aszCheckSheetID, IIf(chkIsToday.Value = vbChecked, True, False))
    
    'debug.Print "totalCheckSheetStationInfo end:" & Time
    
    
    nCount = m_rsStationQuantity.RecordCount
    If nCount > 0 Then
        vsStationTotal.Rows = nCount + 1
    End If
    With vsStationTotal
        For i = 1 To nCount
            .TextMatrix(i, VI_SellStation) = FormatDbValue(m_rsStationQuantity!sell_station_name)
            .TextMatrix(i, VI_Route) = FormatDbValue(m_rsStationQuantity!route_name)
            '车次不显示
            '.TextMatrix(i, VI_Bus) = FormatDbValue(m_rsStationQuantity!bus_id)
            .TextMatrix(i, VI_Station) = FormatDbValue(m_rsStationQuantity!station_name)
            .TextMatrix(i, VI_TicketType) = FormatDbValue(m_rsStationQuantity!ticket_type_name)
            .TextMatrix(i, VI_VehicleType) = FormatDbValue(m_rsStationQuantity!vehicle_type_name)
            .TextMatrix(i, VI_Quantity) = FormatDbValue(m_rsStationQuantity!Quantity)
            .TextMatrix(i, VI_AreaRatio) = Val(FormatDbValue(m_rsStationQuantity!Annotation))
            
            
            '下面的信息纯粹只在汇总时用到
            
            .TextMatrix(i, VI_RouteID) = FormatDbValue(m_rsStationQuantity!route_id)
            .TextMatrix(i, VI_VehicleTypeID) = FormatDbValue(m_rsStationQuantity!vehicle_type_code)
            .TextMatrix(i, VI_SellStationID) = FormatDbValue(m_rsStationQuantity!sell_station_id)
            .TextMatrix(i, VI_StationID) = FormatDbValue(m_rsStationQuantity!station_id)
            .TextMatrix(i, VI_TicketTypeID) = FormatDbValue(m_rsStationQuantity!ticket_type)
            m_rsStationQuantity.MoveNext
        Next i
    
    End With
    SetNormal
    ShowSBInfo ""
    
    Exit Sub
ErrorHandle:
    SetNormal
    ShowSBInfo ""
    ShowErrorMsg
End Sub


Private Sub FillCheckSheetStationDayList()

    '填充路单站点每日人数信息
    '需将数据库中查出来的记录集转为列头是到站，行头是日期，内容是人数。
    Dim oSplit As New Split
    Dim rsStationQuantityList As Recordset
    On Error GoTo ErrorHandle
    Dim i As Integer
    Dim nCount As Integer
    Dim j As Integer
    Dim rsTemp As Recordset
'    Dim m_rsStationInfo As Recordset
    Dim nCols As Integer
    
    Dim lTemp As Long
    SetBusy
    ShowSBInfo "正在填充路单站点每日人数信息"
    oSplit.Init g_oActiveUser
    FillSheetID
    
    'debug.Print "GetCheckSheetDistinctStation start:" & Time
    Set m_rsStationInfo = oSplit.GetCheckSheetDistinctStation(m_aszCheckSheetID, IIf(chkIsToday.Value = vbChecked, True, False))
    
    'debug.Print "GetCheckSheetDistinctStation end:" & Time
    nCols = m_rsStationInfo.RecordCount
    
    
    'debug.Print "TotalCheckSheetStationInfoEx start:" & Time
    Set rsStationQuantityList = oSplit.TotalCheckSheetStationInfoEx(m_aszCheckSheetID, IIf(chkIsToday.Value = vbChecked, True, False))
    
    'debug.Print "TotalCheckSheetStationInfoEx end:" & Time
    '将记录集进行转换
    Set rsTemp = MakeRecordSetDayList(rsStationQuantityList, m_rsStationInfo)
    
    
    nCount = rsTemp.RecordCount
    
    With vsStationDayList
    
    
        .Rows = nCount + 1
        .Cols = nCols + 1 + 1 '多了个小计
        
        .FixedCols = 1
        .FixedRows = 1
        .ColWidth(0) = 1000
        
        .Clear
        
        .AllowUserResizing = flexResizeColumns
        '填充站名
        .TextMatrix(0, 0) = "日期\人数\到站"
        m_rsStationInfo.MoveFirst
        For i = 1 To nCols
            .TextMatrix(0, i) = FormatDbValue(m_rsStationInfo!station_name) & "(" & GetUnicodeBySize(FormatDbValue(m_rsStationInfo!ticket_type_name), 2) & ")"
            m_rsStationInfo.MoveNext
        Next i
        .TextMatrix(0, i) = "小计"
        
        AlignHeadWidth Me.name, vsStationDayList
'
        rsTemp.MoveFirst
        For i = 1 To nCount
            .TextMatrix(i, 0) = FormatDbValue(rsTemp!bus_date)
            For j = 1 To nCols
                lTemp = FormatDbValue(rsTemp.Fields("station_" & j))
                .TextMatrix(i, j) = IIf(lTemp = 0, "", lTemp)
            Next j
            .TextMatrix(i, j) = IIf(FormatDbValue(rsTemp.Fields("total_num")) = 0, "", FormatDbValue(rsTemp.Fields("total_num")))
            rsTemp.MoveNext
            
'            .TextMatrix(i, VI_SellStation) = FormatDbValue(m_rsStationQuantity!sell_station_name)
'            .TextMatrix(i, VI_Route) = FormatDbValue(m_rsStationQuantity!route_name)
'            '车次不显示
'            '.TextMatrix(i, VI_Bus) = FormatDbValue(m_rsStationQuantity!bus_id)
'            .TextMatrix(i, VI_Station) = FormatDbValue(m_rsStationQuantity!station_name)
'            .TextMatrix(i, VI_TicketType) = FormatDbValue(m_rsStationQuantity!ticket_type_name)
'            .TextMatrix(i, VI_VehicleType) = FormatDbValue(m_rsStationQuantity!vehicle_type_name)
'            .TextMatrix(i, VI_Quantity) = FormatDbValue(m_rsStationQuantity!Quantity)
'            .TextMatrix(i, VI_AreaRatio) = Val(FormatDbValue(m_rsStationQuantity!Annotation))
'            m_rsStationQuantity.MoveNext
        Next i
'
    End With
    
    ShowSBInfo ""
        
    SetNormal
    
    Exit Sub
ErrorHandle:
    SetNormal
    
    ShowSBInfo ""
    ShowErrorMsg
End Sub



Private Sub FillCheckSheetStationList()
    '填充路单站点明细
    Dim oSplit As New Split
'    Dim m_rsStationList As Recordset
    On Error GoTo ErrorHandle
    Dim i As Integer
    Dim nCount As Integer
    Dim j As Integer
    SetBusy

    ShowSBInfo "正在填充路单站点明细"
    
    
    With vsStationList
        '1:825,2:1080,3:570,4:645,5:720,6:540,7:0,8:0,9:0,10:0,11:615,12:0,13:0,14:0,15:0,16:0,17:0,18:480,19:540,20:585,21:690,22:0,23:720,24:720,25:720,26:720,27:720,28:720,29:720,30:750,31:720,32:720,

        'quantity mileage  ticket_price base_carriage price_item_1 price_item_2 price_item_3 price_item_4 price_item_5 price_item_6 price_item_7 price_item_8 price_item_9 price_item_10 price_item_11 price_item_12 price_item_13 price_item_14 price_item_15 sell_station_name ticket_type_name seat_type_name
        '设置允许点击列首排序

        .Rows = 2
        .Cols = VIL_BasePrice + m_rsPriceItem.RecordCount
        
        
        .ExplorerBar = flexExSortShowAndMove  '设置允许点列头排序
        
        .FrozenCols = 0 '设置冻结的列
        .AllowUserFreezing = flexFreezeColumns '设置允许调整冻结列的位置
        .AllowUserResizing = flexResizeColumns
        
        .FixedCols = 1
        .FixedRows = 1
        .ColWidth(0) = 100
        .Clear
        '设置合并
        .MergeCells = flexMergeRestrictColumns
        
        
        .MergeCol(VIL_CheckSheetID) = True
        .MergeCol(VIL_BusDate) = True
        .MergeCol(VIL_BusID) = True
        .MergeCol(VIL_SellStationName) = True
        .MergeCol(VIL_StationName) = True
        
        
        
    
    
    
        .TextMatrix(0, VIL_CheckSheetID) = "路单号"
        .TextMatrix(0, VIL_BusDate) = "日期"
        .TextMatrix(0, VIL_BusID) = "车次"
        .TextMatrix(0, VIL_SellStationName) = "上车站"
        .TextMatrix(0, VIL_StationName) = "站点"
        .TextMatrix(0, VIL_TicketTypeName) = "票种"
        
        .TextMatrix(0, VIL_PriceIdentify) = "序"
        .TextMatrix(0, VIL_SellStationID) = "上车站代码"
        .TextMatrix(0, VIL_StationID) = "站点代码"
        .TextMatrix(0, VIL_TicketType) = "票种代码"
        .TextMatrix(0, VIL_StatusName) = "改并状态"
        
        .TextMatrix(0, VIL_Quantity) = "人数"
        .TextMatrix(0, VIL_Mileage) = "里程"
        .TextMatrix(0, VIL_TicketPrice) = "单价"
        .TextMatrix(0, VIL_TotalTicketPrice) = "总价"
        
           
        '下面的信息纯粹只在汇总时用到
        .TextMatrix(0, VIL_RouteID) = "线路代码"
        .TextMatrix(0, VIL_VehicleTypeID) = "车型代码"
        
     
        '下面信息在保存站点明细时用到
        
        .TextMatrix(0, VIL_SeatTypeID) = "位型"
        .TextMatrix(0, VIL_BusSerialNO) = "车次序号"
        .TextMatrix(0, VIL_StationSerial) = "站点序号"
        '在计算应结费用时用到
        .TextMatrix(0, VIL_AreaRatio) = "地区费率"
        .TextMatrix(0, VIL_StatusCode) = "改并状态代码"
        
        m_rsPriceItem.MoveFirst
        For i = 1 To m_rsPriceItem.RecordCount
            .TextMatrix(0, VIL_BasePrice + i - 1) = FormatDbValue(m_rsPriceItem!chinese_name)
            .ColWidth(VIL_BasePrice + i - 1) = 720
            m_rsPriceItem.MoveNext
        Next i
        
        
        
'        AlignHeadWidth Me.name, vsStationList
        .ColWidth(VIL_CheckSheetID) = 825
        .ColWidth(VIL_BusDate) = 1080
        .ColWidth(VIL_BusID) = 570
        .ColWidth(VIL_SellStationName) = 645
        .ColWidth(VIL_StationName) = 720
        .ColWidth(VIL_TicketTypeName) = 540
        .ColWidth(VIL_PriceIdentify) = 0
        .ColWidth(VIL_SellStationID) = 0
        .ColWidth(VIL_StationID) = 0
        .ColWidth(VIL_TicketType) = 0
        .ColWidth(VIL_StatusName) = 615
        .ColWidth(VIL_RouteID) = 0
        .ColWidth(VIL_VehicleTypeID) = 0
        .ColWidth(VIL_SeatTypeID) = 0
        .ColWidth(VIL_BusSerialNO) = 0
        .ColWidth(VIL_StationSerial) = 0
        .ColWidth(VIL_AreaRatio) = 0
        .ColWidth(VIL_Quantity) = 480
        .ColWidth(VIL_Mileage) = 540
        .ColWidth(VIL_TicketPrice) = 585
        .ColWidth(VIL_TotalTicketPrice) = 690
        .ColWidth(VIL_StatusCode) = 0
        .ColWidth(VIL_BasePrice) = 720
        
        
        
        
        
    End With
    oSplit.Init g_oActiveUser
    FillSheetID
    
    'debug.Print "GetCheckSheetStationList start:" & Time
    Set m_rsStationList = oSplit.GetCheckSheetStationList(m_aszCheckSheetID, IIf(chkIsToday.Value = vbChecked, True, False))
    
    Dim rsLicenseTagNo As Recordset
    Set rsLicenseTagNo = oSplit.GetLicenseTagNo(m_aszCheckSheetID, IIf(chkIsToday.Value = vbChecked, True, False))
    
    szLicenseTagNO = FormatDbValue(rsLicenseTagNo!license_tag_no)
    rsLicenseTagNo.MoveNext
    For i = 2 To rsLicenseTagNo.RecordCount
            szLicenseTagNO = szLicenseTagNO & "," & FormatDbValue(rsLicenseTagNo!license_tag_no)
            rsLicenseTagNo.MoveNext
    Next i
    
    'debug.Print "GetCheckSheetStationList end:" & Time
    nCount = m_rsStationList.RecordCount
    If nCount > 0 Then
        vsStationList.Rows = nCount + 1
    End If
    With vsStationList
        WriteProcessBar True, 0, nCount
        For i = 1 To nCount
            
            WriteProcessBar True, i, nCount
            .TextMatrix(i, VIL_CheckSheetID) = FormatDbValue(m_rsStationList!check_sheet_id)
            .TextMatrix(i, VIL_BusDate) = ToDBDate(FormatDbValue(m_rsStationList!bus_date))
            .TextMatrix(i, VIL_BusID) = FormatDbValue(m_rsStationList!bus_id)
            .TextMatrix(i, VIL_SellStationName) = FormatDbValue(m_rsStationList!sell_station_name)
            .TextMatrix(i, VIL_StationName) = FormatDbValue(m_rsStationList!station_name)
            .TextMatrix(i, VIL_TicketTypeName) = FormatDbValue(m_rsStationList!ticket_type_name)
            .TextMatrix(i, VIL_PriceIdentify) = FormatDbValue(m_rsStationList!price_identify)
            .TextMatrix(i, VIL_SellStationID) = FormatDbValue(m_rsStationList!sell_station_id2) '显示的是站点代码
        
        
            .TextMatrix(i, VIL_StationID) = FormatDbValue(m_rsStationList!station_id)
            .TextMatrix(i, VIL_TicketType) = FormatDbValue(m_rsStationList!ticket_type)
            .TextMatrix(i, VIL_StatusName) = GetSheetStationStatusName(FormatDbValue(m_rsStationList!Status))
            .TextMatrix(i, VIL_StatusCode) = FormatDbValue(m_rsStationList!Status)
            '如果允许结算多次,则显示剩余的人数,否则显示所有的人数
            If Not g_oParam.AllowSplitBySomeTimes Then
                .TextMatrix(i, VIL_Quantity) = FormatDbValue(m_rsStationList!Quantity)
            Else
                .TextMatrix(i, VIL_Quantity) = FormatDbValue(m_rsStationList!Quantity) - FormatDbValue(m_rsStationList!fact_quantity)
            End If
            .TextMatrix(i, VIL_Mileage) = Val(FormatDbValue(m_rsStationList!Mileage))
            .TextMatrix(i, VIL_TicketPrice) = Val(FormatDbValue(m_rsStationList!ticket_price))
            .TextMatrix(i, VIL_TotalTicketPrice) = .TextMatrix(i, VIL_Quantity) * .TextMatrix(i, VIL_TicketPrice)
            .TextMatrix(i, VIL_BasePrice) = Val(FormatDbValue(m_rsStationList!base_carriage))
            
            
            '下面的信息纯粹只在汇总时用到
            .TextMatrix(i, VIL_RouteID) = FormatDbValue(m_rsStationList!route_id)
            
'            .TextMatrix(i, VIL_RouteName) = FormatDbValue(m_rsStationList!route_name)
            
            .TextMatrix(i, VIL_VehicleTypeID) = FormatDbValue(m_rsStationList!vehicle_type_code)
        
            
            '下面信息在保存站点明细时用到
            .TextMatrix(i, VIL_SeatTypeID) = FormatDbValue(m_rsStationList!seat_type_id)
            .TextMatrix(i, VIL_BusSerialNO) = FormatDbValue(m_rsStationList!bus_serial_no)
            .TextMatrix(i, VIL_StationSerial) = FormatDbValue(m_rsStationList!station_serial)
            '在计算应结费用时用到
            .TextMatrix(i, VIL_AreaRatio) = FormatDbValue(m_rsStationList!Annotation)
        
        
            m_rsPriceItem.MoveFirst
            For j = 1 To m_rsPriceItem.RecordCount
                If Val(FormatDbValue(m_rsPriceItem!price_item)) = 0 Then
                    '如果为基本运价,则略过
                Else
                    .TextMatrix(i, VIL_BasePrice + j - 1) = Val(FormatDbValue(m_rsStationList.Fields("price_item_" & Val(FormatDbValue(m_rsPriceItem!price_item)))))
                    
                End If
                m_rsPriceItem.MoveNext
            Next j
            m_rsStationList.MoveNext
        Next i
    
        WriteProcessBar False
    End With
    
    ShowSBInfo ""
    SetNormal
    Exit Sub
ErrorHandle:
    ShowSBInfo ""
    SetNormal
    ShowErrorMsg
End Sub

'生成按公司拆算的记录集
Private Function MakeRecordSetListByCompany() As Recordset

    Dim rsTemp As New Recordset
    Dim i As Integer
    Dim j As Integer
    
    Dim oCompany As New Company
    
    
    
    '保存临时数据
    Dim szCompanyID As String
    Dim szVehicleTypeID As String
    Dim szRouteID As String
    Dim szSellStationID As String
    Dim szStationID As String
    Dim szTicketType As String
    Dim szCompanyName As String
    Dim szRouteName As String
    Dim szSellStationName As String
    Dim szStationName As String
    Dim szVehicleTypeName As String
    Dim szTicketTypeName As String
    Dim szAnnotation As String
    Dim lQuantity As Long
    Dim dbMileage As Double
    Dim dbTicketPrice As Double
    Dim dbBasePrice As Double
    Dim adbPriceItem(1 To cnPriceItemNum) As Double
    
    Dim lSheetCount As Long
    Dim szCheckSheetID As String
    
    On Error GoTo ErrorHandle
    If vsStationList.Rows <= 1 Then Exit Function
    
        With rsTemp.Fields
        .Append "transport_company_id", adChar, 12
        .Append "vehicle_type_code", adChar, 3
        .Append "route_id", adChar, 4
        .Append "sell_station_id", adChar, 9
        .Append "station_id", adChar, 9
        .Append "ticket_type", adInteger
        .Append "transport_company_short_name", adVarChar, 10
'        .Append "route_name", adChar, 16
        .Append "sell_station_name", adChar, 10
        .Append "station_name", adVarChar, 10
        .Append "vehicle_type_short_name", adVarChar, 10
        .Append "ticket_type_name", adChar, 12
        .Append "annotation", adVarChar, 255   '存放的是地区的费率
        .Append "quantity", adBigInt
        .Append "mileage", adDouble
        .Append "ticket_price", adCurrency
        .Append "base_carriage", adCurrency
        
        For i = 1 To cnPriceItemNum
            .Append "price_item_" & i, adCurrency
        Next i
        
        .Append "check_sheet_count", adBigInt '路单数
    End With
    
    rsTemp.Open
    '排好序,再进行汇总
    With vsStationList
        .ColSort(VIL_VehicleTypeID) = flexSortGenericAscending
        .ColSort(VIL_RouteID) = flexSortGenericAscending
        .ColSort(VIL_SellStationID) = flexSortGenericAscending
        .ColSort(VIL_StationID) = flexSortGenericAscending
        .ColSort(VIL_TicketType) = flexSortGenericAscending
        .Select 1, 1, .Rows - 1, .Cols - 1
        .Sort = flexSortUseColSort
    End With
    
    
    
    With vsStationList
        
        '得到公司代码及名称
        '因为是按公司拆算,所以对象内的就是公司代码
        szCompanyID = ResolveDisplay(txtObject.Text)
        oCompany.Init g_oActiveUser
        oCompany.Identify szCompanyID
        szCompanyName = oCompany.CompanyShortName
        
        i = 1
        szVehicleTypeID = .TextMatrix(i, VIL_VehicleTypeID)
        szRouteID = .TextMatrix(i, VIL_RouteID)
        szSellStationID = .TextMatrix(i, VIL_SellStationID)
        szStationID = .TextMatrix(i, VIL_StationID)
        szTicketType = .TextMatrix(i, VIL_TicketType)
        'szRouteName = .TextMatrix(i,vil_route  '暂时不传入
        szSellStationName = .TextMatrix(i, VIL_SellStationName)
        szStationName = .TextMatrix(i, VIL_StationName)
'        szVehicleTypeName = .TextMatrix(i,vil_ve '暂时不传入
        szTicketTypeName = .TextMatrix(i, VIL_TicketTypeName)
        szAnnotation = .TextMatrix(i, VIL_AreaRatio)
'        lQuantity = .TextMatrix(i, VIL_Quantity)
'        dbMileage = .TextMatrix(i, VIL_Mileage) * lQuantity '里程数还要乘以人数
'        dbTicketPrice = .TextMatrix(i, VIL_TotalTicketPrice) '取总票价,已乘了人数
        
    
        For i = 1 To .Rows - 1
            '赋予记录集
            
            
            '不同时,新建一记录
            'c.vehicle_type_code , c.route_id , i.station_id , s.station_id , s.ticket_type
            If szVehicleTypeID <> .TextMatrix(i, VIL_VehicleTypeID) Or szRouteID <> .TextMatrix(i, VIL_RouteID) _
                Or szStationID <> .TextMatrix(i, VIL_StationID) Or szTicketType <> .TextMatrix(i, VIL_TicketType) _
                Or szSellStationID <> .TextMatrix(i, VIL_SellStationID) Then
                    
                rsTemp.AddNew
                
                rsTemp!transport_company_id = szCompanyID
                rsTemp!vehicle_type_code = szVehicleTypeID
                rsTemp!route_id = szRouteID
                rsTemp!sell_station_id = szSellStationID
                rsTemp!station_id = szStationID
                rsTemp!ticket_type = szTicketType
                rsTemp!transport_company_short_name = szCompanyName
'                rsTemp!route_name = szRouteName
                rsTemp!sell_station_name = szSellStationName
                rsTemp!station_name = szStationName
                rsTemp!vehicle_type_short_name = szVehicleTypeName
                rsTemp!ticket_type_name = szTicketTypeName
                rsTemp!Annotation = szAnnotation
                rsTemp!Quantity = lQuantity
                rsTemp!Mileage = dbMileage
                rsTemp!ticket_price = dbTicketPrice
                rsTemp!base_carriage = dbBasePrice
                For j = 1 To cnPriceItemNum
                    rsTemp.Fields("price_item_" & j) = adbPriceItem(j)
                Next j

                rsTemp!CHECK_SHEET_COUNT = lSheetCount
                
                rsTemp.Update
        
            
                
                '清空原值
                lQuantity = 0
                dbMileage = 0
                dbTicketPrice = 0
                dbBasePrice = 0
                For j = 1 To cnPriceItemNum
                    adbPriceItem(j) = 0
                Next j

                lSheetCount = 0
                szCheckSheetID = .TextMatrix(i, VIL_CheckSheetID)
'                lSheetCount = lSheetCount + 1


                szVehicleTypeID = .TextMatrix(i, VIL_VehicleTypeID)
                szRouteID = .TextMatrix(i, VIL_RouteID)
                szSellStationID = .TextMatrix(i, VIL_SellStationID)
                szStationID = .TextMatrix(i, VIL_StationID)
                szTicketType = .TextMatrix(i, VIL_TicketType)
                'szRouteName = .TextMatrix(i,vil_route  '暂时不传入
                szSellStationName = .TextMatrix(i, VIL_SellStationName)
                szStationName = .TextMatrix(i, VIL_StationName)
        '        szVehicleTypeName = .TextMatrix(i,vil_ve '暂时不传入
                szTicketTypeName = .TextMatrix(i, VIL_TicketTypeName)
                szAnnotation = .TextMatrix(i, VIL_AreaRatio)
                
                
                
                
                '除去总票价,其他均要乘以人数
                lQuantity = lQuantity + .TextMatrix(i, VIL_Quantity)
                dbMileage = dbMileage + .TextMatrix(i, VIL_Mileage) * .TextMatrix(i, VIL_Quantity)
                dbTicketPrice = dbTicketPrice + .TextMatrix(i, VIL_TicketPrice)
                dbBasePrice = dbBasePrice + .TextMatrix(i, VIL_BasePrice) * .TextMatrix(i, VIL_Quantity)
                
                '赋票价项
                m_rsPriceItem.MoveFirst
                For j = 1 To m_rsPriceItem.RecordCount
                    If Val(FormatDbValue(m_rsPriceItem!price_item)) = 0 Then
                        '如果为基本运价,则略过
                    Else
                        adbPriceItem(Val(FormatDbValue(m_rsPriceItem!price_item))) = adbPriceItem(Val(FormatDbValue(m_rsPriceItem!price_item))) + .TextMatrix(i, VIL_BasePrice + j - 1) * .TextMatrix(i, VIL_Quantity)
                    End If
                    m_rsPriceItem.MoveNext
                Next j
                
            Else
                '如果不同
                
                '除去总票价,其他均要乘以人数
                lQuantity = lQuantity + .TextMatrix(i, VIL_Quantity)
                dbMileage = dbMileage + .TextMatrix(i, VIL_Mileage) * .TextMatrix(i, VIL_Quantity)
                dbTicketPrice = dbTicketPrice + .TextMatrix(i, VIL_TicketPrice)
                dbBasePrice = dbBasePrice + .TextMatrix(i, VIL_BasePrice) * .TextMatrix(i, VIL_Quantity)
                
'                If .TextMatrix(i, VIL_CheckSheetID) <> szCheckSheetID Then
'                    szCheckSheetID = .TextMatrix(i, VIL_CheckSheetID)
'                    lSheetCount = lSheetCount + 1
'                End If
                
                Dim bFind As Boolean
                '因为未按路单排序,所以需向前搜索.
                
                bFind = False
                For j = i - 1 To 1 Step -1
                    If .TextMatrix(j, VIL_CheckSheetID) = .TextMatrix(i, VIL_CheckSheetID) Then
                        bFind = True
                        Exit For
                    End If
                Next j
                If Not bFind Then
                    '如果未找到,则累加
                    lSheetCount = lSheetCount + 1
                End If
                '赋票价项
                m_rsPriceItem.MoveFirst
                For j = 1 To m_rsPriceItem.RecordCount
                    If Val(FormatDbValue(m_rsPriceItem!price_item)) = 0 Then
                        '如果为基本运价,则略过
                    Else
                        adbPriceItem(Val(FormatDbValue(m_rsPriceItem!price_item))) = adbPriceItem(Val(FormatDbValue(m_rsPriceItem!price_item))) + .TextMatrix(i, VIL_BasePrice + j - 1) * .TextMatrix(i, VIL_Quantity)
                    End If
                    m_rsPriceItem.MoveNext
                Next j
            End If
            
        Next i
        
        
        
        
        rsTemp.AddNew
        rsTemp!transport_company_id = szCompanyID
        rsTemp!vehicle_type_code = szVehicleTypeID
        rsTemp!route_id = szRouteID
        rsTemp!sell_station_id = szSellStationID
        rsTemp!station_id = szStationID
        rsTemp!ticket_type = szTicketType
        rsTemp!transport_company_short_name = szCompanyName
'        rsTemp!route_name = szRouteName
        rsTemp!sell_station_name = szSellStationName
        rsTemp!station_name = szStationName
        rsTemp!vehicle_type_short_name = szVehicleTypeName
        rsTemp!ticket_type_name = szTicketTypeName
        rsTemp!Annotation = szAnnotation
        rsTemp!Quantity = lQuantity
        rsTemp!Mileage = dbMileage
        rsTemp!ticket_price = dbTicketPrice
        rsTemp!base_carriage = dbBasePrice
        For j = 1 To cnPriceItemNum
            rsTemp.Fields("price_item_" & j) = adbPriceItem(j)
        Next j
        rsTemp!CHECK_SHEET_COUNT = lSheetCount
        rsTemp.Update
        
        
        
    End With
    
    Set MakeRecordSetListByCompany = rsTemp
    Exit Function
ErrorHandle:
    ShowErrorMsg
    
End Function



'生成按车辆拆算的记录集
Private Function MakeRecordSetListByVehicle() As Recordset

    Dim rsTemp As New Recordset
    Dim i As Integer
    Dim j As Integer
    
    
    Dim oVehicle As New Vehicle
    
    
    '保存临时数据
    Dim szVehicleID As String
    Dim szRouteID As String
    Dim szSellStationID As String
    Dim szStationID As String
    Dim szTicketType As String
    Dim szLicenseTagNO As String
    
    Dim szRouteName As String
    Dim szSellStationName As String
    Dim szStationName As String
    Dim szVehicleTypeID As String
    Dim szVehicleTypeName As String
    Dim szTicketTypeName As String
    
    Dim szCompanyID As String
    Dim szCompanyName As String
    
    Dim szAnnotation As String
    Dim lQuantity As Long
    Dim dbMileage As Double
    Dim dbTicketPrice As Double
    Dim dbBasePrice As Double
    Dim adbPriceItem(1 To cnPriceItemNum) As Double
    
    Dim lSheetCount As Long
    Dim szCheckSheetID As String
    
    On Error GoTo ErrorHandle
    
    If vsStationList.Rows <= 1 Then Exit Function
    
    With rsTemp.Fields
        .Append "vehicle_id", adChar, 5
        
        .Append "vehicle_type_code", adChar, 3
        .Append "route_id", adChar, 4
        .Append "sell_station_id", adChar, 9
        .Append "station_id", adChar, 9
        .Append "ticket_type", adInteger
        
        
        .Append "license_tag_no", adChar, 10
        
        .Append "route_name", adChar, 16
        .Append "sell_station_name", adChar, 10
        .Append "station_name", adVarChar, 10
        .Append "vehicle_type_short_name", adVarChar, 10
        .Append "ticket_type_name", adChar, 12
        .Append "transport_company_id", adChar, 12
        .Append "transport_company_short_name", adVarChar, 10
        .Append "annotation", adVarChar, 255   '存放的是地区的费率
        .Append "quantity", adBigInt
        .Append "mileage", adDouble
        .Append "ticket_price", adCurrency
        .Append "base_carriage", adCurrency
        
        For i = 1 To cnPriceItemNum
            .Append "price_item_" & i, adCurrency
        Next i
    
        .Append "check_sheet_count", adBigInt '路单数
        
    End With

    rsTemp.Open
    '排好序,再进行汇总
    'c.route_id , i.station_id , s.station_id , s.ticket_type
    With vsStationList
        .ColSort(VIL_RouteID) = flexSortGenericAscending
        .ColSort(VIL_SellStationID) = flexSortGenericAscending
        .ColSort(VIL_StationID) = flexSortGenericAscending
        .ColSort(VIL_TicketType) = flexSortGenericAscending
        .Select 1, 1, .Rows - 1, .Cols - 1
        .Sort = flexSortUseColSort
    End With
    
    
        
    With vsStationList
        
        '得到公司代码及名称
        '因为是按公司拆算,所以对象内的就是公司代码
        szVehicleID = ResolveDisplay(txtObject.Text)
        oVehicle.Init g_oActiveUser
        oVehicle.Identify szVehicleID
        szLicenseTagNO = oVehicle.LicenseTag
        szCompanyID = oVehicle.Company
        szCompanyName = oVehicle.CompanyName
        szVehicleTypeID = oVehicle.VehicleModel
        i = 1
        
'        szVehicleTypeID = .TextMatrix(i, VIL_VehicleTypeID)
        szRouteID = .TextMatrix(i, VIL_RouteID)
'        szRouteName = .TextMatrix(i, VIL_RouteName)
        szSellStationID = .TextMatrix(i, VIL_SellStationID)
        szStationID = .TextMatrix(i, VIL_StationID)
        szTicketType = .TextMatrix(i, VIL_TicketType)
'        szRouteName = .TextMatrix(i, VI_Route) '暂时不传入
        szSellStationName = .TextMatrix(i, VIL_SellStationName)
        szStationName = .TextMatrix(i, VIL_StationName)
'        szVehicleTypeName = .TextMatrix(i,vil_ve '暂时不传入
        szTicketTypeName = .TextMatrix(i, VIL_TicketTypeName)
        szAnnotation = .TextMatrix(i, VIL_AreaRatio)
        
        
        For i = 1 To .Rows - 1
            '赋予记录集
            
            
            '不同时,新建一记录
            '
            If szRouteID <> .TextMatrix(i, VIL_RouteID) Or szStationID <> .TextMatrix(i, VIL_StationID) _
                Or szTicketType <> .TextMatrix(i, VIL_TicketType) Or szSellStationID <> .TextMatrix(i, VIL_SellStationID) Then
                
                rsTemp.AddNew
                rsTemp!vehicle_id = szVehicleID
                rsTemp!license_tag_no = szLicenseTagNO
                rsTemp!transport_company_id = szCompanyID
                rsTemp!vehicle_type_code = szVehicleTypeID
                rsTemp!route_id = szRouteID
                rsTemp!sell_station_id = szSellStationID
                rsTemp!station_id = szStationID
                rsTemp!ticket_type = szTicketType
                rsTemp!transport_company_short_name = szCompanyName
'                rsTemp!route_name = szRouteName
                rsTemp!sell_station_name = szSellStationName
                rsTemp!station_name = szStationName
                rsTemp!vehicle_type_short_name = szVehicleTypeName
                rsTemp!ticket_type_name = szTicketTypeName
                rsTemp!Annotation = szAnnotation
                rsTemp!Quantity = lQuantity
                rsTemp!Mileage = dbMileage
                rsTemp!ticket_price = dbTicketPrice
                rsTemp!base_carriage = dbBasePrice
                For j = 1 To cnPriceItemNum
                    rsTemp.Fields("price_item_" & j) = adbPriceItem(j)
                Next j

                rsTemp!CHECK_SHEET_COUNT = lSheetCount
                
                rsTemp.Update
        
            
                '清空原值
                lQuantity = 0
                dbMileage = 0
                dbTicketPrice = 0
                dbBasePrice = 0
                For j = 1 To cnPriceItemNum
                    adbPriceItem(j) = 0
                Next j
                
                lSheetCount = 0
                


                szVehicleTypeID = .TextMatrix(i, VIL_VehicleTypeID)
                szRouteID = .TextMatrix(i, VIL_RouteID)
                szSellStationID = .TextMatrix(i, VIL_SellStationID)
                szStationID = .TextMatrix(i, VIL_StationID)
                szTicketType = .TextMatrix(i, VIL_TicketType)
                'szRouteName = .TextMatrix(i,vil_route  '暂时不传入
                szSellStationName = .TextMatrix(i, VIL_SellStationName)
                szStationName = .TextMatrix(i, VIL_StationName)
        '        szVehicleTypeName = .TextMatrix(i,vil_ve '暂时不传入
                szTicketTypeName = .TextMatrix(i, VIL_TicketTypeName)
                szAnnotation = .TextMatrix(i, VIL_AreaRatio)
                
                
                szCheckSheetID = .TextMatrix(i, VIL_CheckSheetID)
'                lSheetCount = lSheetCount + 1
                
                lQuantity = lQuantity + .TextMatrix(i, VIL_Quantity)
                dbMileage = dbMileage + .TextMatrix(i, VIL_Mileage) * .TextMatrix(i, VIL_Quantity) '里程数还要乘以人数
                dbTicketPrice = dbTicketPrice + .TextMatrix(i, VIL_TotalTicketPrice) '取总票价,已乘了人数
                dbBasePrice = dbBasePrice + .TextMatrix(i, VIL_BasePrice) * .TextMatrix(i, VIL_Quantity)
                
                '赋票价项
                m_rsPriceItem.MoveFirst
                For j = 1 To m_rsPriceItem.RecordCount
                    If Val(FormatDbValue(m_rsPriceItem!price_item)) = 0 Then
                        '如果为基本运价,则略过
                    Else
                        adbPriceItem(Val(FormatDbValue(m_rsPriceItem!price_item))) = adbPriceItem(Val(FormatDbValue(m_rsPriceItem!price_item))) + .TextMatrix(i, VIL_BasePrice + j - 1) * .TextMatrix(i, VIL_Quantity)
                    End If
                    m_rsPriceItem.MoveNext
                Next j
                
                
            Else
                '如果不同
                lQuantity = lQuantity + .TextMatrix(i, VIL_Quantity)
                dbMileage = dbMileage + .TextMatrix(i, VIL_Mileage) * .TextMatrix(i, VIL_Quantity) '里程数还要乘以人数
                dbTicketPrice = dbTicketPrice + .TextMatrix(i, VIL_TotalTicketPrice) '取总票价,已乘了人数
                dbBasePrice = dbBasePrice + .TextMatrix(i, VIL_BasePrice) * .TextMatrix(i, VIL_Quantity)
                
'                If .TextMatrix(i, VIL_CheckSheetID) <> szCheckSheetID Then
'                    szCheckSheetID = .TextMatrix(i, VIL_CheckSheetID)
'                    lSheetCount = lSheetCount + 1
'                End If
                
                Dim bFind As Boolean
                '因为未按路单排序,所以需向前搜索.
                
                bFind = False
                For j = i - 1 To 1 Step -1
                    If .TextMatrix(j, VIL_CheckSheetID) = .TextMatrix(i, VIL_CheckSheetID) Then
                        bFind = True
                        Exit For
                    End If
                Next j
                If Not bFind Then
                    '如果未找到,则累加
                    lSheetCount = lSheetCount + 1
                End If
                
                
                
                '赋票价项
                m_rsPriceItem.MoveFirst
                For j = 1 To m_rsPriceItem.RecordCount
                    If Val(FormatDbValue(m_rsPriceItem!price_item)) = 0 Then
                        '如果为基本运价,则略过
                    Else
                        adbPriceItem(Val(FormatDbValue(m_rsPriceItem!price_item))) = adbPriceItem(Val(FormatDbValue(m_rsPriceItem!price_item))) + .TextMatrix(i, VIL_BasePrice + j - 1) * .TextMatrix(i, VIL_Quantity)
                    End If
                    m_rsPriceItem.MoveNext
                Next j
                
                
                
            End If
            
        Next i
        
        
        rsTemp.AddNew
        rsTemp!vehicle_id = szVehicleID
        rsTemp!license_tag_no = szLicenseTagNO

        rsTemp!transport_company_id = szCompanyID
        rsTemp!vehicle_type_code = szVehicleTypeID
        rsTemp!route_id = szRouteID
        rsTemp!route_name = szRouteName
        rsTemp!sell_station_id = szSellStationID
        rsTemp!station_id = szStationID
        rsTemp!ticket_type = szTicketType
        rsTemp!transport_company_short_name = szCompanyName
'        rsTemp!route_name = szRouteName
        rsTemp!sell_station_name = szSellStationName
        rsTemp!station_name = szStationName
        rsTemp!vehicle_type_short_name = szVehicleTypeName
        rsTemp!ticket_type_name = szTicketTypeName
        rsTemp!Annotation = szAnnotation
        rsTemp!Quantity = lQuantity
        rsTemp!Mileage = dbMileage
        rsTemp!ticket_price = dbTicketPrice
        rsTemp!base_carriage = dbBasePrice
        For j = 1 To cnPriceItemNum
            rsTemp.Fields("price_item_" & j) = adbPriceItem(j)
        Next j

        rsTemp!CHECK_SHEET_COUNT = lSheetCount
        rsTemp.Update
        
        
        
    End With
    
    Set MakeRecordSetListByVehicle = rsTemp
    
    Exit Function
ErrorHandle:
    ShowErrorMsg
End Function

'生成按车次拆算的记录集
Private Function MakeRecordSetListByBus() As Recordset

    Dim rsTemp As New Recordset
    Dim i As Integer
    Dim j As Integer

'    Dim oBus As New Bus
    
    
    '保存临时数据
    Dim szBusID As String
    Dim szVehicleID As String
    Dim szRouteID As String
    Dim szSellStationID As String
    Dim szStationID As String
    Dim szTicketType As String
    Dim szLicenseTagNO As String
    
    Dim szRouteName As String
    Dim szSellStationName As String
    Dim szStationName As String
    Dim szVehicleTypeID As String
    Dim szVehicleTypeName As String
    Dim szTicketTypeName As String
    
    Dim szCompanyID As String
    Dim szCompanyName As String
    
    Dim szAnnotation As String
    Dim lQuantity As Long
    Dim dbMileage As Double
    Dim dbTicketPrice As Double
    Dim dbBasePrice As Double
    Dim adbPriceItem(1 To cnPriceItemNum) As Double
    
    Dim lSheetCount As Long
    Dim szCheckSheetID As String
    
    On Error GoTo ErrorHandle
    
    If vsStationList.Rows <= 1 Then Exit Function
    
    With rsTemp.Fields
        .Append "bus_id", adChar, 5
        
        .Append "vehicle_type_code", adChar, 3
        .Append "route_id", adChar, 4
        .Append "sell_station_id", adChar, 9
        .Append "station_id", adChar, 9
        .Append "ticket_type", adInteger
        
        
        .Append "license_tag_no", adChar, 10
        
        .Append "route_name", adChar, 16
        .Append "sell_station_name", adChar, 10
        .Append "station_name", adVarChar, 10
        .Append "vehicle_type_short_name", adVarChar, 10
        .Append "ticket_type_name", adChar, 12
        .Append "transport_company_id", adChar, 12
        .Append "transport_company_short_name", adVarChar, 10
        .Append "annotation", adVarChar, 255   '存放的是地区的费率
        .Append "quantity", adBigInt
        .Append "mileage", adDouble
        .Append "ticket_price", adCurrency
        .Append "base_carriage", adCurrency
        
        For i = 1 To cnPriceItemNum
            .Append "price_item_" & i, adCurrency
        Next i
        
        .Append "check_sheet_count", adBigInt '路单数
        
    End With

    rsTemp.Open
    '排好序,再进行汇总
    'c.route_id , i.station_id , s.station_id , s.ticket_type
    With vsStationList
        .ColSort(VIL_RouteID) = flexSortGenericAscending
        .ColSort(VIL_SellStationID) = flexSortGenericAscending
        .ColSort(VIL_StationID) = flexSortGenericAscending
        .ColSort(VIL_TicketType) = flexSortGenericAscending
        .Select 1, 1, .Rows - 1, .Cols - 1
        .Sort = flexSortUseColSort
    End With
    
    
        
    With vsStationList
        
        '得到公司代码及名称
        '因为是按公司拆算,所以对象内的就是公司代码
        szBusID = ResolveDisplay(txtObject.Text)
'        oBus.Init g_oActiveUser
'        oBus.Identify szBusID
'        szLicenseTagNO = oVehicle.LicenseTag
        szCompanyID = ResolveDisplay(cboCompany.Text)
'        szCompanyName = oVehicle.CompanyName
'        szVehicleTypeID = oVehicle.VehicleModel
        i = 1
        
'        szVehicleTypeID = .TextMatrix(i, VIL_VehicleTypeID)
        szRouteID = .TextMatrix(i, VIL_RouteID)
'        szRouteName = .TextMatrix(i, VIL_RouteName)
        szSellStationID = .TextMatrix(i, VIL_SellStationID)
        szStationID = .TextMatrix(i, VIL_StationID)
        szTicketType = .TextMatrix(i, VIL_TicketType)
'        szRouteName = .TextMatrix(i, VI_Route) '暂时不传入
        szSellStationName = .TextMatrix(i, VIL_SellStationName)
        szStationName = .TextMatrix(i, VIL_StationName)
'        szVehicleTypeName = .TextMatrix(i,vil_ve '暂时不传入
        szTicketTypeName = .TextMatrix(i, VIL_TicketTypeName)
        szAnnotation = .TextMatrix(i, VIL_AreaRatio)
        
        
        For i = 1 To .Rows - 1
            '赋予记录集
            
            
            '不同时,新建一记录
            '
            If szRouteID <> .TextMatrix(i, VIL_RouteID) Or szStationID <> .TextMatrix(i, VIL_StationID) _
                Or szTicketType <> .TextMatrix(i, VIL_TicketType) Or szSellStationID <> .TextMatrix(i, VIL_SellStationID) Then
                
                rsTemp.AddNew
                rsTemp!bus_id = szBusID
'                rsTemp!license_tag_no = szLicenseTagNO
                rsTemp!transport_company_id = szCompanyID
                rsTemp!vehicle_type_code = szVehicleTypeID
                rsTemp!route_id = szRouteID
                rsTemp!sell_station_id = szSellStationID
                rsTemp!station_id = szStationID
                rsTemp!ticket_type = szTicketType
                rsTemp!transport_company_short_name = szCompanyName
'                rsTemp!route_name = szRouteName
                rsTemp!sell_station_name = szSellStationName
                rsTemp!station_name = szStationName
                rsTemp!vehicle_type_short_name = szVehicleTypeName
                rsTemp!ticket_type_name = szTicketTypeName
                rsTemp!Annotation = szAnnotation
                rsTemp!Quantity = lQuantity
                rsTemp!Mileage = dbMileage
                rsTemp!ticket_price = dbTicketPrice
                rsTemp!base_carriage = dbBasePrice
                For j = 1 To cnPriceItemNum
                    rsTemp.Fields("price_item_" & j) = adbPriceItem(j)
                Next j
                rsTemp!CHECK_SHEET_COUNT = lSheetCount
                
                rsTemp.Update
        
            
                '清空原值
                lQuantity = 0
                dbMileage = 0
                dbTicketPrice = 0
                dbBasePrice = 0
                For j = 1 To cnPriceItemNum
                    adbPriceItem(j) = 0
                Next j
                lSheetCount = 0
                

                szVehicleTypeID = .TextMatrix(i, VIL_VehicleTypeID)
                szRouteID = .TextMatrix(i, VIL_RouteID)
                szSellStationID = .TextMatrix(i, VIL_SellStationID)
                szStationID = .TextMatrix(i, VIL_StationID)
                szTicketType = .TextMatrix(i, VIL_TicketType)
                'szRouteName = .TextMatrix(i,vil_route  '暂时不传入
                szSellStationName = .TextMatrix(i, VIL_SellStationName)
                szStationName = .TextMatrix(i, VIL_StationName)
        '        szVehicleTypeName = .TextMatrix(i,vil_ve '暂时不传入
                szTicketTypeName = .TextMatrix(i, VIL_TicketTypeName)
                szAnnotation = .TextMatrix(i, VIL_AreaRatio)
                
                szCheckSheetID = .TextMatrix(i, VIL_CheckSheetID)
'                lSheetCount = lSheetCount + 1
                lQuantity = lQuantity + .TextMatrix(i, VIL_Quantity)
                dbMileage = dbMileage + .TextMatrix(i, VIL_Mileage) * .TextMatrix(i, VIL_Quantity) '里程数还要乘以人数
                dbTicketPrice = dbTicketPrice + .TextMatrix(i, VIL_TotalTicketPrice) '取总票价,已乘了人数
                dbBasePrice = dbBasePrice + .TextMatrix(i, VIL_BasePrice) * .TextMatrix(i, VIL_Quantity)
                
                '赋票价项
                m_rsPriceItem.MoveFirst
                For j = 1 To m_rsPriceItem.RecordCount
                    If Val(FormatDbValue(m_rsPriceItem!price_item)) = 0 Then
                        '如果为基本运价,则略过
                    Else
                        adbPriceItem(Val(FormatDbValue(m_rsPriceItem!price_item))) = adbPriceItem(Val(FormatDbValue(m_rsPriceItem!price_item))) + .TextMatrix(i, VIL_BasePrice + j - 1) * .TextMatrix(i, VIL_Quantity)
                    End If
                    m_rsPriceItem.MoveNext
                Next j
                
                
            Else
                '如果不同
                lQuantity = lQuantity + .TextMatrix(i, VIL_Quantity)
                dbMileage = dbMileage + .TextMatrix(i, VIL_Mileage) * .TextMatrix(i, VIL_Quantity) '里程数还要乘以人数
                dbTicketPrice = dbTicketPrice + .TextMatrix(i, VIL_TotalTicketPrice) '取总票价,已乘了人数
                dbBasePrice = dbBasePrice + .TextMatrix(i, VIL_BasePrice) * .TextMatrix(i, VIL_Quantity)
                
'                If .TextMatrix(i, VIL_CheckSheetID) <> szCheckSheetID Then
'                    szCheckSheetID = .TextMatrix(i, VIL_CheckSheetID)
'                    lSheetCount = lSheetCount + 1
'                End If
                Dim bFind As Boolean
                '因为未按路单排序,所以需向前搜索.
                
                bFind = False
                For j = i - 1 To 1 Step -1
                    If .TextMatrix(j, VIL_CheckSheetID) = .TextMatrix(i, VIL_CheckSheetID) Then
                        bFind = True
                        Exit For
                    End If
                Next j
                If Not bFind Then
                    '如果未找到,则累加
                    lSheetCount = lSheetCount + 1
                End If
                
                '赋票价项
                m_rsPriceItem.MoveFirst
                For j = 1 To m_rsPriceItem.RecordCount
                    If Val(FormatDbValue(m_rsPriceItem!price_item)) = 0 Then
                        '如果为基本运价,则略过
                    Else
                        adbPriceItem(Val(FormatDbValue(m_rsPriceItem!price_item))) = adbPriceItem(Val(FormatDbValue(m_rsPriceItem!price_item))) + .TextMatrix(i, VIL_BasePrice + j - 1) * .TextMatrix(i, VIL_Quantity)
                    End If
                    m_rsPriceItem.MoveNext
                Next j
            End If
            
        Next i
        
        
        rsTemp.AddNew
        rsTemp!bus_id = szBusID
'        rsTemp!license_tag_no = szLicenseTagNO
'
        rsTemp!transport_company_id = szCompanyID
        rsTemp!vehicle_type_code = szVehicleTypeID
        rsTemp!route_id = szRouteID
        rsTemp!route_name = szRouteName
        rsTemp!sell_station_id = szSellStationID
        rsTemp!station_id = szStationID
        rsTemp!ticket_type = szTicketType
        rsTemp!transport_company_short_name = szCompanyName
'        rsTemp!route_name = szRouteName
        rsTemp!sell_station_name = szSellStationName
        rsTemp!station_name = szStationName
        rsTemp!vehicle_type_short_name = szVehicleTypeName
        rsTemp!ticket_type_name = szTicketTypeName
        rsTemp!Annotation = szAnnotation
        rsTemp!Quantity = lQuantity
        rsTemp!Mileage = dbMileage
        rsTemp!ticket_price = dbTicketPrice
        rsTemp!base_carriage = dbBasePrice
        For j = 1 To cnPriceItemNum
            rsTemp.Fields("price_item_" & j) = adbPriceItem(j)
        Next j
        
        rsTemp!CHECK_SHEET_COUNT = lSheetCount
        rsTemp.Update
        
        
        
    End With
    
    Set MakeRecordSetListByBus = rsTemp
    
    Exit Function
ErrorHandle:
    ShowErrorMsg
End Function

Private Function GetSheetStationInfoRS() As Recordset
    Dim rsTemp As New Recordset
    
    Dim i As Integer
    Dim j As Integer
    If vsStationList.Rows <= 1 Then Exit Function
    
    
    With rsTemp.Fields
        .Append "check_sheet_id", adChar, 10
        
        .Append "sell_station_id", adChar, 9
        .Append "station_id", adChar, 9
        .Append "price_identify", adSmallInt
        .Append "ticket_type", adSmallInt
        .Append "seat_type_id", adChar, 3
        .Append "fact_quantity", adSmallInt
        .Append "status", adSmallInt
        .Append "bus_date", adDate
        .Append "bus_id", adChar, 5
        .Append "bus_serial_no", adSmallInt
        .Append "station_serial", adSmallInt
        .Append "station_name", adChar, 10
'        .Append "quantity", adSmallInt
        .Append "mileage", adDouble
        .Append "ticket_price", adDouble
        .Append "base_carriage", adDouble
        For i = 1 To 15
            .Append "price_item_" & i, adDouble
        Next i
        
        
    End With
    '重新排序
    With vsStationList
        .ColSort(VIL_CheckSheetID) = flexSortGenericAscending
        .ColSort(VIL_SellStationID) = flexSortGenericAscending
        .ColSort(VIL_StationID) = flexSortGenericAscending
        .ColSort(VIL_PriceIdentify) = flexSortGenericAscending
        .ColSort(VIL_TicketType) = flexSortGenericAscending
        .ColSort(VIL_SeatTypeID) = flexSortGenericAscending
        .ColSort(VIL_Quantity) = flexSortGenericAscending
        
        .Select 1, 1, .Rows - 1, .Cols - 1
        .Sort = flexSortUseColSort
    End With
    
    rsTemp.Open
    With vsStationList
        For i = 1 To .Rows - 1
            rsTemp.AddNew
            rsTemp!check_sheet_id = .TextMatrix(i, VIL_CheckSheetID)
            rsTemp!sell_station_id = .TextMatrix(i, VIL_SellStationID)
            rsTemp!station_id = .TextMatrix(i, VIL_StationID)
            rsTemp!price_identify = .TextMatrix(i, VIL_PriceIdentify)
            rsTemp!ticket_type = .TextMatrix(i, VIL_TicketType)
            rsTemp!seat_type_id = .TextMatrix(i, VIL_SeatTypeID)
            rsTemp!fact_quantity = .TextMatrix(i, VIL_Quantity)
                
            rsTemp!Status = .TextMatrix(i, VIL_StatusCode)
            rsTemp!bus_date = .TextMatrix(i, VIL_BusDate)
            rsTemp!bus_id = .TextMatrix(i, VIL_BusID)
            rsTemp!bus_serial_no = .TextMatrix(i, VIL_BusSerialNO)
            rsTemp!station_serial = .TextMatrix(i, VIL_StationSerial)
            rsTemp!station_name = .TextMatrix(i, VIL_StationName)
'            rsTemp!Quantity = .TextMatrix(i, VIL_Quantity)
            rsTemp!Mileage = .TextMatrix(i, VIL_Mileage)
            rsTemp!ticket_price = .TextMatrix(i, VIL_TicketPrice)
            rsTemp!base_carriage = .TextMatrix(i, VIL_BasePrice)
            '初始化
            For j = 1 To 15
                rsTemp.Fields("price_item_" & j) = 0
            Next j
            
            '赋票价项
            m_rsPriceItem.MoveFirst
            For j = 1 To m_rsPriceItem.RecordCount
                If Val(FormatDbValue(m_rsPriceItem!price_item)) = 0 Then
                    '如果为基本运价,则略过
                Else
                    rsTemp.Fields("price_item_" & Val(FormatDbValue(m_rsPriceItem!price_item))) = Val(.TextMatrix(i, VIL_BasePrice + j - 1))
                End If
                m_rsPriceItem.MoveNext
            Next j
            
            
            rsTemp.Update
            
        Next i
    End With
    
    Set GetSheetStationInfoRS = rsTemp
    
    
End Function

'
'
'Private Function TotalSheetStationInfoRS() As Recordset
'    Dim rsTemp As New Recordset
'
'    Dim i As Integer
'
'    If vsStationList.Rows <= 1 Then Exit Function
'
'    With rsTemp.Fields
'        .Append "sell_station_id", adChar, 9
'        .Append "bus_id", adChar, 5
'        .Append "route_id", adChar, 4
'        .Append "station_id", adChar, 9
'        .Append "ticket_type_id", adSmallInt
'        .Append "transport_company_id", adChar, 12
'        .Append "vehicle_type_code", adChar, 3
'        .Append "vehicle_id", adChar, 5
'        .Append "route_name", adChar, 16
'        .Append "quantity", adSmallInt
'
'    End With
'    '重新排序
'    With vsStationList
'        .ColSort(VIL_CheckSheetID) = flexSortGenericAscending
'        .ColSort(VIL_SellStationID) = flexSortGenericAscending
'        .ColSort(VIL_StationID) = flexSortGenericAscending
'        .ColSort(VIL_PriceIdentify) = flexSortGenericAscending
'        .ColSort(VIL_SeatTypeID) = flexSortGenericAscending
'        .ColSort(VIL_Quantity) = flexSortGenericAscending
'
'        .Select 1, 1, .Rows - 1, .Cols - 1
'        .Sort = flexSortUseColSort
'    End With
'
'
'    With vsStationList
'        For i = 1 To .Rows - 1
'            rsTemp.AddNew
'            rsTemp!check_sheet_id = .TextMatrix(i, VIL_CheckSheetID)
'            rsTemp!sell_station_id = .TextMatrix(i, VIL_SellStationID)
'            rsTemp!station_id = .TextMatrix(i, VIL_StationID)
'            rsTemp!price_identify = .TextMatrix(i, VIL_PriceIdentify)
'            rsTemp!seat_type_id = .TextMatrix(i, VIL_SeatTypeID)
'            rsTemp!fact_quantity = .TextMatrix(i, VIL_Quantity)
''            rsTemp!bus_id = .TextMatrix(i, VIL_BusID)
''            rsTemp!route_id = .TextMatrix(i, VIL_RouteID)
'            rsTemp.Update
'
'        Next i
'    End With
'
'    Set GetSheetStationInfoRS = rsTemp
'
'
'End Function


Private Sub FillVsExtra()
    '初始化手工补票的表格
    Dim i As Integer
    With vsExtra
        .Cols = 6
        .Rows = 8
        .Clear
        
        .TextMatrix(0, VIE_NO) = "序"
        .TextMatrix(0, VIE_Quantity) = "人数"
        .TextMatrix(0, VIE_TotalTicketPrice) = "票款"
        .TextMatrix(0, VIE_Ratio) = "劳务费比率"
        .TextMatrix(0, VIE_ServicePrice) = "劳务费"
        .TextMatrix(0, VIE_SettleOutPrice) = "拆出金额"
        '填充序
        For i = 1 To .Rows - 1
            .TextMatrix(i, 0) = i
        Next i
    End With
    
End Sub

Private Function ValidateVsExtra() As Boolean
    '验证表格的有效性   ,并赋值
    Dim i As Integer
    Dim nCount As Integer
    Dim atExtraInfo() As TSettleExtraInfo
    
    With vsExtra
        
        '验证表格的有效性
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, VIE_Quantity)) <> 0 Or Val(.TextMatrix(i, VIE_TotalTicketPrice)) <> 0 Then
                '有数据输入
                If Val(.TextMatrix(i, VIE_Quantity)) = 0 Or Val(.TextMatrix(i, VIE_TotalTicketPrice)) = 0 Then
                    MsgBox "第" & i & "行数据输入有误,人数、金额有一个数为0", vbExclamation, Me.Caption
                    '有无效的数据
                    ValidateVsExtra = False
                    Exit Function
                    
                End If
                nCount = nCount + 1
            End If
        Next i
        '赋值
        If nCount > 0 Then
            ReDim atExtraInfo(1 To nCount)
            nCount = 0
            For i = 1 To .Rows - 1
                If Val(.TextMatrix(i, VIE_Quantity)) <> 0 Or Val(.TextMatrix(i, VIE_TotalTicketPrice)) <> 0 Then
                    nCount = nCount + 1
                    atExtraInfo(nCount).PassengerNumber = Val(.TextMatrix(i, VIE_Quantity))
                    atExtraInfo(nCount).TotalTicketPrice = Val(.TextMatrix(i, VIE_TotalTicketPrice))
                    atExtraInfo(nCount).Ratio = Val(.TextMatrix(i, VIE_Ratio))
                    atExtraInfo(nCount).ServicePrice = Val(.TextMatrix(i, VIE_ServicePrice))
                    atExtraInfo(nCount).SettleOutPrice = Val(.TextMatrix(i, VIE_SettleOutPrice))
                End If
            Next i
        End If
    End With
    
    m_atExtraInfo = atExtraInfo
    '输入有效
    ValidateVsExtra = True
        
End Function


Private Sub vsExtra_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    '进行允许输入验证,及自动更新
    With vsExtra
        If Col = VIE_Quantity Or Col = VIE_Ratio Or Col = VIE_TotalTicketPrice Then

'            If .TextMatrix(Row, VIE_Quantity) = "" Then .TextMatrix(Row, VIE_Quantity) = 0
'            If .TextMatrix(Row, VIE_Ratio) = "" Then .TextMatrix(Row, VIE_Ratio) = 0
'            If .TextMatrix(Row, VIE_TotalTicketPrice) = "" Then .TextMatrix(Row, VIE_TotalTicketPrice) = 0
            
            '精确到元
            .TextMatrix(Row, VIE_ServicePrice) = FormatMoney(Val(.TextMatrix(Row, VIE_TotalTicketPrice)) * Val(.TextMatrix(Row, VIE_Ratio)) / 100)
            
            .TextMatrix(Row, VIE_SettleOutPrice) = FormatMoney(Val(.TextMatrix(Row, VIE_TotalTicketPrice)) - Val(.TextMatrix(Row, VIE_ServicePrice)))
            
        End If
    End With
End Sub


Private Sub tbSection_ButtonClick(ByVal Button As MSComctlLib.Button)
    '对手工补票人数输入的表格进行操作
    Dim i As Integer
    Select Case Button.Key
    Case "add"
        '新增一行
        vsExtra.Rows = vsExtra.Rows + 1
        vsExtra.TextMatrix(vsExtra.Rows - 1, 0) = vsExtra.Rows - 1
        
    Case "del"
        '删除一行
        If vsExtra.Rows = 1 Then Exit Sub
        
        If (MsgBox("确认要删除第" & vsExtra.Rows - 1 & "行吗？", vbQuestion + vbYesNo, Me.Caption) = vbYes) Then
            
            vsExtra.RemoveItem vsExtra.Row
            For i = 1 To vsExtra.Rows - 1
                vsExtra.TextMatrix(i, 0) = i
            Next i
            
        End If
    End Select
    
End Sub




Private Sub vsExtra_BeforeSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long, Cancel As Boolean)
    '控制输入，错的，不应该是NewRowSel，而是NewColSel，FPD
    If NewColSel <> VIE_ServicePrice And NewColSel <> VIE_SettleOutPrice Then
        vsExtra.Editable = flexEDKbdMouse
    Else
        vsExtra.Editable = flexEDNone
    End If
End Sub


Private Sub vsExtra_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '判断是否为数字
    With vsExtra
        If Col > 0 And .TextMatrix(Row, Col) <> "" Then
            If Not IsNumeric(.TextMatrix(Row, Col)) Then
                '如果不是数字,则出错
                MsgBox "必须输入数字", vbExclamation, Me.Caption
            Else
                If .TextMatrix(Row, Col) < 0 Then
                    '如果小于0,则出错
                    MsgBox "必须大于0", vbExclamation, Me.Caption
                End If
                If IsNumeric(.TextMatrix(Row, VIE_Ratio)) Then
                    If .TextMatrix(Row, VIE_Ratio) > 100 Then
                        MsgBox "费率必须小于100", vbExclamation, Me.Caption
                    End If
                End If
            End If
        End If
    End With
End Sub









