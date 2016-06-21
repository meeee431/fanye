VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Begin VB.Form frmWizSplitSettleBack 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "回程路单结算向导"
   ClientHeight    =   6435
   ClientLeft      =   1530
   ClientTop       =   3015
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   8655
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox ptTitle 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   8655
      TabIndex        =   38
      Top             =   15
      Width           =   8655
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
         TabIndex        =   40
         Top             =   150
         Width           =   780
      End
      Begin VB.Label lblContent 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请选择路单结算的方式。"
         Height          =   180
         Left            =   360
         TabIndex        =   39
         Top             =   450
         Width           =   1980
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   -120
      TabIndex        =   33
      Top             =   -285
      Width           =   8775
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   -30
      TabIndex        =   32
      Top             =   855
      Width           =   8685
   End
   Begin RTComctl3.CoolButton cmdFinish 
      Height          =   315
      Left            =   6000
      TabIndex        =   34
      Top             =   6015
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
      MICON           =   "frmWizSplitSettleBack.frx":0000
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
      Left            =   7290
      TabIndex        =   35
      Top             =   6015
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
      MICON           =   "frmWizSplitSettleBack.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdNext 
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   315
      Left            =   6000
      TabIndex        =   36
      Top             =   6015
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
      MICON           =   "frmWizSplitSettleBack.frx":0038
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
      Left            =   4770
      TabIndex        =   37
      Top             =   6015
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
      MICON           =   "frmWizSplitSettleBack.frx":0054
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ImageList imglv1 
      Left            =   8610
      Top             =   915
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
            Picture         =   "frmWizSplitSettleBack.frx":0070
            Key             =   "splitcompany"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWizSplitSettleBack.frx":01CA
            Key             =   "company"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWizSplitSettleBack.frx":0324
            Key             =   "vehicle"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWizSplitSettleBack.frx":047E
            Key             =   "bus"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWizSplitSettleBack.frx":05D8
            Key             =   "busowner"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   " "
      Height          =   990
      Left            =   -30
      TabIndex        =   41
      Top             =   5745
      Width           =   9465
      Begin MSComctlLib.ProgressBar pbFill 
         Height          =   285
         Left            =   1620
         TabIndex        =   42
         Top             =   300
         Visible         =   0   'False
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
      End
      Begin RTComctl3.CoolButton cmdHelp 
         Height          =   315
         Left            =   300
         TabIndex        =   43
         Top             =   300
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
         MICON           =   "frmWizSplitSettleBack.frx":0732
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
   Begin VB.Frame fraWizard 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "第三步"
      Height          =   4755
      Index           =   3
      Left            =   105
      TabIndex        =   16
      Top             =   960
      Width           =   8625
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   3555
         Top             =   150
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmWizSplitSettleBack.frx":074E
               Key             =   "DELETE"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmWizSplitSettleBack.frx":0AA1
               Key             =   "INSERT"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmWizSplitSettleBack.frx":0DF4
               Key             =   "APPEND"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tbOperation 
         Height          =   390
         Left            =   180
         TabIndex        =   45
         Top             =   15
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         Wrappable       =   0   'False
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "APPEND"
               Description     =   "追加行"
               Object.ToolTipText     =   "追加行"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "DELETE"
               Description     =   "删除行"
               Object.ToolTipText     =   "删除行"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "INSERT"
               Description     =   "插入行"
               Object.ToolTipText     =   "插入行"
               ImageIndex      =   2
            EndProperty
         EndProperty
      End
      Begin VSFlex7LCtl.VSFlexGrid vsStation 
         Height          =   4275
         Left            =   60
         TabIndex        =   17
         Top             =   450
         Width           =   8325
         _cx             =   14684
         _cy             =   7541
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
   Begin VB.Frame fraWizard 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "第二步"
      Height          =   4755
      Index           =   2
      Left            =   105
      TabIndex        =   10
      Top             =   960
      Width           =   8625
      Begin VB.TextBox txtSheetID 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2880
         TabIndex        =   11
         Top             =   1380
         Width           =   1785
      End
      Begin FText.asFlatMemo txtAnnotation 
         Height          =   1440
         Left            =   2880
         TabIndex        =   12
         Top             =   1860
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   2540
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
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "填写结算单编号和其他结算单信息，根据结算的条件不同，对某些结算单信息可以不用填写。"
         Height          =   405
         Left            =   1470
         TabIndex        =   15
         Top             =   690
         Width           =   5445
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "结算单编号(&S):"
         Height          =   180
         Left            =   1500
         TabIndex        =   14
         Top             =   1440
         Width           =   1260
      End
      Begin VB.Label lblMemo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "备注(&R):"
         Height          =   180
         Left            =   1530
         TabIndex        =   13
         Top             =   1920
         Width           =   720
      End
   End
   Begin VB.Frame fraWizard 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "第一步"
      Height          =   4755
      Index           =   1
      Left            =   105
      TabIndex        =   0
      Top             =   960
      Width           =   8625
      Begin VB.CheckBox chkSingle 
         BackColor       =   &H00E0E0E0&
         Caption         =   "按单行"
         Height          =   270
         Left            =   1155
         TabIndex        =   44
         Top             =   3120
         Width           =   1695
      End
      Begin FText.asFlatTextBox txtObject 
         Height          =   315
         Left            =   5460
         TabIndex        =   1
         Top             =   1950
         Width           =   1575
         _ExtentX        =   2778
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
         Height          =   300
         Left            =   5460
         TabIndex        =   2
         Top             =   2460
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   529
         _Version        =   393216
         Format          =   61800448
         CurrentDate     =   37642
      End
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   300
         Left            =   2235
         TabIndex        =   3
         Top             =   2520
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   529
         _Version        =   393216
         Format          =   61800448
         CurrentDate     =   37622
      End
      Begin MSComctlLib.ImageCombo imgcbo 
         Height          =   330
         Left            =   2235
         TabIndex        =   4
         Top             =   1950
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Locked          =   -1  'True
         ImageList       =   "imglv1"
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmWizSplitSettleBack.frx":1147
         Height          =   840
         Left            =   1110
         TabIndex        =   9
         Top             =   630
         Width           =   6165
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "结束日期(&E):"
         Height          =   180
         Left            =   4320
         TabIndex        =   8
         Top             =   2580
         Width           =   1080
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "开始日期(&B):"
         Height          =   225
         Left            =   1170
         TabIndex        =   7
         Top             =   2595
         Width           =   1380
      End
      Begin VB.Label label2 
         BackStyle       =   0  'Transparent
         Caption         =   "对象(&O):"
         Height          =   180
         Index           =   1
         Left            =   4320
         TabIndex        =   6
         Top             =   2040
         Width           =   810
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "类型(&S):"
         Height          =   180
         Index           =   1
         Left            =   1170
         TabIndex        =   5
         Top             =   2010
         Width           =   720
      End
      Begin VB.Line Line5 
         BorderColor     =   &H8000000C&
         X1              =   1095
         X2              =   7020
         Y1              =   1530
         Y2              =   1530
      End
      Begin VB.Line Line6 
         BorderColor     =   &H8000000E&
         X1              =   1080
         X2              =   6990
         Y1              =   1545
         Y2              =   1545
      End
   End
   Begin VB.Frame fraWizard 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "第四步"
      Height          =   4755
      Index           =   4
      Left            =   105
      TabIndex        =   18
      Top             =   960
      Width           =   8625
      Begin VB.Frame Frame5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "路单结算总汇"
         Height          =   675
         Left            =   300
         TabIndex        =   23
         Top             =   180
         Width           =   8025
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "总人数:"
            Height          =   180
            Left            =   3900
            TabIndex        =   31
            Top             =   300
            Width           =   630
         End
         Begin VB.Label lblTotalQuautity 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "220.3"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   4590
            TabIndex        =   30
            Top             =   300
            Width           =   450
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "应结票款:"
            Height          =   180
            Left            =   1650
            TabIndex        =   29
            Top             =   300
            Width           =   810
         End
         Begin VB.Label lblNeedSplitMoney 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "220.3"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   2520
            TabIndex        =   28
            Top             =   300
            Width           =   450
         End
         Begin VB.Label lblSheetCount 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "2张"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   6720
            TabIndex        =   27
            Top             =   300
            Width           =   270
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "路单张数:"
            Height          =   180
            Left            =   5850
            TabIndex        =   26
            Top             =   300
            Width           =   810
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "总票款:"
            Height          =   180
            Left            =   180
            TabIndex        =   25
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
            TabIndex        =   24
            Top             =   300
            Width           =   495
         End
      End
      Begin VB.Frame fraVehicle 
         BackColor       =   &H00E0E0E0&
         Caption         =   "车辆结算明细表"
         Height          =   1845
         Left            =   300
         TabIndex        =   21
         Top             =   2850
         Width           =   8025
         Begin MSComctlLib.ListView lvVehicle 
            Height          =   1365
            Left            =   180
            TabIndex        =   22
            Top             =   300
            Width           =   7590
            _ExtentX        =   13388
            _ExtentY        =   2408
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
         Height          =   1845
         Left            =   300
         TabIndex        =   19
         Top             =   930
         Width           =   8025
         Begin MSComctlLib.ListView lvCompany 
            Height          =   1365
            Left            =   180
            TabIndex        =   20
            Top             =   300
            Width           =   7605
            _ExtentX        =   13414
            _ExtentY        =   2408
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
   End
End
Attribute VB_Name = "frmWizSplitSettleBack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'路单列首常数
Const PI_CheckSheetID = 0
Const PI_Station = 1
Const PI_CheckGate = 10
Const PI_Route = 9
Const PI_BusID = 2
Const PI_LicenseTagNo = 5
Const PI_CompanyName = 6
Const PI_Quantity = 7
Const PI_Mileage = 8
Const PI_TicketPrice = 4
Const PI_Date = 3
Const PI_VehicleID = 11
Const PI_CompanyID = 12


'网格的列首
Const VI_SellStation = 1
Const VI_Route = 2
'Const VI_Bus = 3
Const VI_VehicleType = 3
Const VI_Station = 4
Const VI_TicketType = 5
Const VI_Quantity = 6
Const VI_SellStationID = 7
Const VI_RouteID = 8
Const VI_VehicleTypeID = 9
Const VI_StationID = 10
'Const VI_TicketTypeID = 11


Const VI_Cols = 11

Dim nSheetCount As Integer  ' 统计路单总数
Dim nValibleCount As Integer  '有效的路单数
Dim nTotalQuantity As Long '总人数
Dim m_aszCheckSheetID() As String '路单组数
Dim m_szVehicleID As String  '车辆数组
Dim m_szCompanyID As String  '公司ID
Dim m_bLogFileValid As Boolean '4日志文件
Dim m_bPromptWhenError As Boolean '2是否提示错误
Dim CancelHasPress As Boolean
Dim TSplitResult As TSplitResult  '预览结果
Dim m_nSplitItenCount As Integer '使用的拆算项数
Dim tSplitItem() As TSplitItemInfo


Public m_bIsManualSettle As Boolean '是否是手工结算,即汇总出来的人数及站点信息,是否允许修改
Dim m_rsStationQuantity As Recordset '站点人数汇总信息

Dim m_szTicketType As String '票种



Private Sub cmdFinish_Click()
    On Error GoTo ErrHandle
    Dim i As Integer
    Dim m_oSplit As New STSettle.Split
    Dim mAnswer As VbMsgBoxResult
    Dim szTemp As String
    SetBusy
    
    '权限验证
    cmdFinish.Visible = True
    cmdFinish.Enabled = False
    cmdNext.Visible = False
    cmdPrevious.Enabled = False
    
    m_oSplit.Init g_oActiveUser
    
    '开始生成结算单 填写lstCreateInfo信息
    CreateFinanceInfo
    '打印结算单
    frmPrintFinSheet.m_SheetID = Trim(txtSheetID.Text)
'    For i = 1 To lvLugSheet.ListItems.Count
'        If i <> lvLugSheet.ListItems.Count Then
'            szTemp = szTemp & lvLugSheet.ListItems(i).Text & ","
'        Else
'            szTemp = szTemp & lvLugSheet.ListItems(i).Text
'        End If
'    Next i
    frmPrintFinSheet.m_szLugSettleSheetID = szTemp
    frmPrintFinSheet.m_bRePrint = False
    
    
'    frmPrintFinSheet.m_dbTotalPrice = Val(lblLugTotalPrice.Caption)
'    frmPrintFinSheet.m_dbNeedSplitPrice = Val(lblLugNeedSplit.Caption)
'    frmPrintFinSheet.m_szProtocol = lblLugProtocol.Caption
    
    
    
'    frmPrintFinSheet.m_bNeedPrint = True
    
    frmPrintFinSheet.ZOrder 0
    frmPrintFinSheet.Show vbModal
    SetNormal
    Exit Sub
ErrHandle:
    SetNormal
    ShowErrorMsg
End Sub

'//***********************************    生成日志区  **************************************
'生成结算单日志
Private Sub CreateFinanceInfo()
  On Error GoTo ErrHandle
    
    CreateFinanceSheetRs
    
    Exit Sub
    
ErrHandle:
  ShowErrorMsg
End Sub

'生成某结算单并记录生成 1成功，2错误，3重试
Private Function CreateFinanceSheetRs() As Integer
    Dim vbMsg As VbMsgBoxResult
    Dim ErrString As String
    Dim bCreateOk As Integer
    Dim nErrNumber As Long
    Dim szErrDescription As String
    Dim aszSheetID() As String   '为了适应接口数组类型
    Dim m_oSplit As New STSettle.Split
    
    Static nHasPrompt, nPromptTime As Integer
    
    Dim oVehicle As New Vehicle
    Dim oCompany As New Company
    Dim szObjectName As String
    
    Dim szLuggageSettleIDs As String
    Dim i As Integer
    
    
    On Error GoTo here
 '   初始化设置开始提醒时的错误次数
    If nPromptTime < 2 Then nPromptTime = 2
    ReDim aszSheetID(1 To 1)
'    aszSheetID(1) = m_aszCheckSheetID(Index)
    bCreateOk = 1
    ErrString = " 结算成功    "
    m_oSplit.Init g_oActiveUser
    
    '****取得对象名称
    If imgcbo.Text = cszCompanyName Then
        oCompany.Init g_oActiveUser
        oCompany.Identify m_szCompanyID
        szObjectName = oCompany.CompanyShortName
        
    Else
        oVehicle.Init g_oActiveUser
        oVehicle.Identify m_szVehicleID
        szObjectName = oVehicle.LicenseTag
        
    End If
    TSplitResult.SettleSheetInfo.ObjectName = szObjectName
    '****取得行包相关信息
    szLuggageSettleIDs = ""
    TSplitResult.SettleSheetInfo.LuggageSettleIDs = szLuggageSettleIDs
    TSplitResult.SettleSheetInfo.LuggageTotalBaseCarriage = 0 ' Val(lblLugTotalPrice.Caption)
    TSplitResult.SettleSheetInfo.LuggageTotalSettlePrice = 0 'Val(lblLugNeedSplit.Caption)
    TSplitResult.SettleSheetInfo.LuggageProtocolName = 0 ' lblLugProtocol.Caption
    
    '开始结算
'    If QuantityHasBeChanged Then
    TSplitResult.SheetStationInfo = GetStationQuantity()
    m_oSplit.SplitCheckSheetManual Trim(txtSheetID.Text), m_szCompanyID, m_szVehicleID, m_aszCheckSheetID, TSplitResult, dtpStartDate.Value, dtpEndDate.Value
'    Else
'        m_oSplit.SplitCheckSheet Trim(txtSheetID.Text), m_szCompanyID, m_szVehicleID, m_aszCheckSheetID, TSplitResult, dtpStartDate.Value, dtpEndDate.Value
'    End If
ErrContinue:
    DoEvents
    CreateFinanceSheetRs = bCreateOk
    Exit Function
here:
    bCreateOk = 2
    nErrNumber = err.Number
    szErrDescription = err.Description
    ErrString = m_aszCheckSheetID(Index) & "[  未生成]" & _
        " * 错误描述:(" & Trim(Str(nErrNumber)) & ")" & Trim(szErrDescription) & " *"
    
    If m_bPromptWhenError Then
        ErrString = "结算单" & m_aszCheckSheetID(Index) & "未生成！" & vbCrLf & _
            Trim(szErrDescription) & "(" & Trim(Str(nErrNumber)) & ")"
        vbMsg = MsgBox(ErrString, vbExclamation + vbAbortRetryIgnore + vbDefaultButton3)
        Select Case vbMsg
               Case vbAbort
                   CancelHasPress = True
               Case vbRetry
                   CreateFinanceSheetRs = 3
                   Exit Function
               Case vbIgnore
                   If nHasPrompt >= nPromptTime - 1 Then
                        If MsgBox("以后不再提示生成错误？", vbQuestion + vbYesNo) = vbYes Then
                            m_bPromptWhenError = False
                        End If
                        nHasPrompt = 0
                        nPromptTime = nPromptTime + 1
                   End If
                   nHasPrompt = nHasPrompt + 1
               Exit Function
        End Select
    Else
        GoTo ErrContinue
    End If
End Function

Private Sub CloseLogFile()
    On Error Resume Next
    If m_bLogFileValid Then
        Close #1
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
    Dim m_oSplit As New STSettle.Split
    SetBusy
    pbFill.Visible = False
    m_oSplit.Init g_oActiveUser
    cmdPrevious.Enabled = True
    If fraWizard(1).Visible = True Then
        '      判断拆算对象有没有选
        If txtObject.Text = "" Then
            MsgBox " 请选择拆算对象!", vbInformation, Me.Caption
            cmdPrevious.Enabled = False
            SetNormal
            Exit Sub
        End If
        cmdNext.Enabled = True
        fraWizard(1).Visible = False
        fraWizard(2).Visible = True
        fraWizard(3).Visible = False
        fraWizard(4).Visible = False
        '       清空第二页界面
        txtSheetID.Text = ""
        
        txtAnnotation.Text = ""
        '   自动生成结算单号 YYYYMM0001格式
        szLastSheetID = m_oSplit.GetLastSettleSheetID
        If szLastSheetID = "0" Then
            txtSheetID.Text = CStr(Year(Now)) + CStr(Month(Now)) + "0001"
        Else
            txtSheetID.Text = szLastSheetID
        End If
        '      txtOperator.SetFocus
    
    ElseIf fraWizard(2).Visible = True Then
        
        cmdFinish.Enabled = False
        cmdFinish.Visible = False
        cmdNext.Visible = True
        
        fraWizard(1).Visible = False
        fraWizard(2).Visible = False
        fraWizard(3).Visible = True
        fraWizard(4).Visible = False
        
        
        
        
    ElseIf fraWizard(3).Visible = True Then
    
        If Not ValidateVSInput Then
            SetNormal
            Exit Sub
        End If
        '填充路单预览信息
        FillPreSheetInfo
        
        
        
        fraWizard(1).Visible = False
        fraWizard(2).Visible = False
        fraWizard(3).Visible = False
        fraWizard(4).Visible = True
        
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

Private Sub cmdPrevious_Click()
    On Error GoTo ErrHandle
    pbFill.Visible = False
    If fraWizard(3).Visible = True Then
        cmdFinish.Visible = False
        cmdFinish.Enabled = False
        cmdNext.Visible = True
        cmdNext.Enabled = True
        cmdPrevious.Enabled = True
        fraWizard(1).Visible = False
        fraWizard(2).Visible = True
        fraWizard(3).Visible = False
        fraWizard(4).Visible = False
        
    ElseIf fraWizard(2).Visible = True Then
        cmdFinish.Visible = False
        cmdFinish.Enabled = False
        cmdNext.Visible = True
        cmdNext.Enabled = True
        cmdPrevious.Enabled = False
        fraWizard(1).Visible = True
        fraWizard(2).Visible = False
        fraWizard(3).Visible = False
        fraWizard(4).Visible = False
        
    
    End If
    
    Exit Sub
ErrHandle:
    ShowErrorMsg
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
'    Me.StartUpPosition = 1
    AlignFormPos Me
    cmdPrevious.Enabled = False
    cmdFinish.Visible = False
    cmdNext.Visible = True
    
    fraWizard(1).Visible = True
    fraWizard(2).Visible = False
    fraWizard(3).Visible = False
    fraWizard(4).Visible = False
    
    dtpStartDate.Value = CDate(Format(Date, "yyyy-mm") & "-1" & " 00:00:01")
    Select Case Month(Date)
    Case 1, 3, 5, 7, 8, 10, 12
        dtpEndDate.Value = CDate(Format(Date, "yyyy-mm") & "-31" & " 23:59:59")
    Case 4, 6, 9, 11
        dtpEndDate.Value = CDate(Format(Date, "yyyy-mm") & "-30" & " 23:59:59")
    Case 2
        dtpEndDate.Value = CDate(Format(Date, "yyyy-mm") & "-28" & " 23:59:59")
    End Select
    
    
    '  类型 初始化cbo对象
    '   1-车辆 2-参运公司
    With imgcbo
        .ComboItems.Clear
        .ComboItems.Add , , "参运公司", 1
        .ComboItems.Add , , "车辆", 3
        .Locked = True
        .ComboItems(1).Selected = True
    End With
    
    GetTicketType
    
    '填充路单的站点统计信息
    
    FillCheckSheetStation
    
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
    lblNeedSplitMoney.Caption = TSplitResult.SettleSheetInfo.SettleOtherCompanyPrice
    
    FilllvCompany
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
    lblNeedSplitMoney.Caption = TSplitResult.SettleSheetInfo.SettleOtherCompanyPrice
    FilllvVehicle
End Sub




Private Sub txtObject_ButtonClick()
    On Error GoTo ErrHandle
    Dim oShell As New STShell.CommDialog
    Dim aszTemp() As String
    oShell.Init g_oActiveUser
    If Trim(imgcbo.Text) = cszCompanyName Then
        aszTemp = oShell.SelectCompany()
    ElseIf Trim(imgcbo.Text) = "车辆" Then
        aszTemp = oShell.SelectVehicle()
    End If
    
    Set oShell = Nothing
    If ArrayLength(aszTemp) = 0 Then Exit Sub
    txtObject.Text = MakeDisplayString(Trim(aszTemp(1, 1)), Trim(aszTemp(1, 2)))
    
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub
Private Sub Form_Unload(Cancel As Integer)
    SaveFormPos Me
    SaveHeadWidth Me.name, lvCompany
    SaveHeadWidth Me.name, lvVehicle
'    SaveHeadWidth Me.name, vsStation
    Unload Me
    
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub



'填充填充路单预览信息
Public Sub FillPreSheetInfo()
    Dim i As Integer
    Dim nCount As Integer
    Dim lvItem As ListItem
    Dim m_oSplit As New STSettle.Split
    
    Dim atStationQuantity() As TSettleSheetStation
    
    '拆算预览
    m_oSplit.Init g_oActiveUser
'    If QuantityHasBeChanged Then
        '已被改变过
    atStationQuantity = GetStationQuantity
        
    If imgcbo.Text = cszCompanyName Then   '转换常量
         m_szCompanyID = ResolveDisplay(txtObject.Text)
         m_szVehicleID = ""
    Else
         m_szVehicleID = ResolveDisplay(txtObject.Text)
         m_szCompanyID = ""
    End If
    TSplitResult = m_oSplit.PreviewSplitCheckSheetManual(m_szCompanyID, m_szVehicleID, m_aszCheckSheetID, atStationQuantity, IIf(chkSingle.Value = vbChecked, True, False))
'    Else
        '未被改变过
'        TSplitResult = m_oSplit.PreviewSplitCheckSheet(m_szCompanyID, m_szVehicleID, m_aszCheckSheetID, IIf(chkSingle.Value = vbChecked, True, False))
'    End If
    '填充信息区
    'tSplitResult.SettleSheetInfo.
    lblTotalPrice.Caption = IIf(TSplitResult.SettleSheetInfo.TotalTicketPrice = 0, "无", TSplitResult.SettleSheetInfo.TotalTicketPrice)
    
    
    lblNeedSplitMoney.Caption = TSplitResult.SettleSheetInfo.SettleOtherCompanyPrice - TSplitResult.SettleSheetInfo.SettleStationPrice
    lblTotalQuautity.Caption = TSplitResult.SettleSheetInfo.TotalQuantity
    lblSheetCount.Caption = TSplitResult.SettleSheetInfo.CheckSheetCount
    
    '填充列首
    lvHead
    '填充公司列表
    FilllvCompany
    '填充车辆列表
    FilllvVehicle


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
        lvItem.SubItems(2) = TSplitResult.CompanyInfo(j).SettlePrice - TSplitResult.CompanyInfo(j).SettleStationPrice
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
        lvItem.SubItems(2) = TSplitResult.VehicleInfo(j).SettlePrice - TSplitResult.VehicleInfo(j).SettleStationPrice
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
            m_nSplitItenCount = m_nSplitItenCount + 1
        End If
    Next i
    
    
    AlignHeadWidth Me.name, lvCompany
    AlignHeadWidth Me.name, lvVehicle
    Exit Sub
here:
    ShowErrorMsg
End Sub




Private Sub FillCheckSheetStation()
    '填充路单站点汇总信息
    Dim oSplit As New Split
'    Dim m_rsStationQuantity As Recordset
    On Error GoTo ErrorHandle
    Dim i As Integer
    Dim nCount As Integer
    Dim j As Integer
    

    With vsStation
        .Rows = 2
        .Cols = VI_Cols
        .FixedCols = 1
        .FixedRows = 1
        .Editable = flexEDKbdMouse
        .ColWidth(0) = 100
        '设置合并
'        .MergeCells = flexMergeRestrictColumns
'        .MergeCol(VI_SellStation) = True
'        .MergeCol(VI_Route) = True
''        .MergeCol(VI_Bus) = True
'        .MergeCol(VI_VehicleType) = True
'        .MergeCol(VI_Station) = True
'        .MergeCol(VI_TicketType) = True
        
        .AllowUserResizing = flexResizeColumns
        
        
        .TextMatrix(0, VI_SellStation) = "上车站"
        .TextMatrix(0, VI_Route) = "线路"
'        .TextMatrix(0, VI_Bus) = "车次"
        .TextMatrix(0, VI_VehicleType) = "车型"
        .TextMatrix(0, VI_Station) = "站点"
        .TextMatrix(0, VI_TicketType) = "票种"
        .TextMatrix(0, VI_Quantity) = "人数"
        .TextMatrix(0, VI_SellStationID) = "上车站代码"
        .TextMatrix(0, VI_RouteID) = "线路代码"
        .TextMatrix(0, VI_VehicleTypeID) = "车型代码"
        .TextMatrix(0, VI_StationID) = "站点代码"
'        .TextMatrix(0, VI_TicketTypeID) = "票种代码"
        
        
        .ColWidth(VI_SellStation) = 1245
        .ColWidth(VI_Route) = 1845
'        .ColWidth(VI_Bus) = 675
        .ColWidth(VI_VehicleType) = 1155
        .ColWidth(VI_Station) = 1215
        .ColWidth(VI_TicketType) = 1000
        .ColWidth(VI_Quantity) = 900
        
        .ColWidth(VI_SellStationID) = 0
        .ColWidth(VI_RouteID) = 0
        .ColWidth(VI_VehicleTypeID) = 0
        .ColWidth(VI_StationID) = 0
'        .ColWidth(VI_TicketTypeID) = 0
        '1:1245        2:1845        3:675         4:1155        5:1215        6:855         7:900         8:0           9:0           10:0          11:0          12:0
        
        
        .ColComboList(VI_SellStation) = "..."
        .ColComboList(VI_Route) = "..."
        .ColComboList(VI_VehicleType) = "..."
        .ColComboList(VI_Station) = "..."
        .ColComboList(VI_TicketType) = m_szTicketType '组合框
        
    End With
    
'    oSplit.Init g_oActiveUser
    
    'Set m_rsStationQuantity = oSplit.TotalCheckSheetStationInfo(m_aszCheckSheetID)
    
'    nCount = m_rsStationQuantity.RecordCount
'    If nCount > 0 Then
'        vsStation.Rows = nCount + 1
'    End If
'    With vsStation
'        For i = 1 To nCount
'            .TextMatrix(i, VI_SellStation) = FormatDbValue(m_rsStationQuantity!sell_station_name)
'            .TextMatrix(i, VI_Route) = FormatDbValue(m_rsStationQuantity!route_name)
'            .TextMatrix(i, VI_Bus) = FormatDbValue(m_rsStationQuantity!bus_id)
'            .TextMatrix(i, VI_Station) = FormatDbValue(m_rsStationQuantity!station_name)
'            .TextMatrix(i, VI_TicketType) = FormatDbValue(m_rsStationQuantity!ticket_type_name)
'            .TextMatrix(i, VI_VehicleType) = FormatDbValue(m_rsStationQuantity!vehicle_type_name)
'            .TextMatrix(i, VI_Quantity) = FormatDbValue(m_rsStationQuantity!Quantity)
'
'            m_rsStationQuantity.MoveNext
'        Next i
'
'    End With
'
    
    
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub vsStation_BeforeSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long, Cancel As Boolean)
    If NewColSel = VI_Quantity Then
        '如果为数量,则允许修改
        With vsStation
            If OldColSel = VI_Quantity Then
                If IsNumeric(.TextMatrix(OldRowSel, VI_Quantity)) Then
                    If .TextMatrix(OldRowSel, VI_Quantity) <= 0 Then
                        MsgBox "数量必须大于零", vbInformation, Me.Caption
                    End If
                Else
                    MsgBox "数量必须输入数字", vbInformation, Me.Caption
                End If
            End If
        End With
    End If

    
End Sub

Private Function QuantityHasBeChanged() As Boolean
    '站点人数信息是否被改变过
    Dim i As Integer
    vsStation.Col = 1
'    If Not m_rsStationQuantity Is Nothing Then
'        m_rsStationQuantity.MoveFirst
'        For i = 1 To m_rsStationQuantity.RecordCount
'            If vsStation.TextMatrix(i, VI_Quantity) <> FormatDbValue(m_rsStationQuantity!Quantity) Then
'                Exit For
'            End If
'            m_rsStationQuantity.MoveNext
'        Next i
'        If i > m_rsStationQuantity.RecordCount Then
'            '未被改变过
'            QuantityHasBeChanged = False
'        Else
'            '被改动过
'            QuantityHasBeChanged = True
'        End If
'    End If
    QuantityHasBeChanged = True
End Function

Private Function GetStationQuantity() As TSettleSheetStation()
    '得到vsStation中的站点人数信息
    Dim atStationQuantity() As TSettleSheetStation
    Dim i As Integer
    Dim j As Integer
    Dim nCount As Integer
    Dim szTicketTypeName As String
    With vsStation
        nCount = 0
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 1) <> "" Then
                nCount = nCount + 1
            End If
        Next i
    
        If nCount = 0 Then Exit Function
        '先算出有多少行站点人数
        ReDim atStationQuantity(1 To nCount)
        
        '汇总站点人数
        '********此处应加入把不同的车次汇总成同一个.
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 1) <> "" Then
                
                atStationQuantity(i).RouteID = .TextMatrix(i, VI_RouteID)
                atStationQuantity(i).RouteName = .TextMatrix(i, VI_Route)
                atStationQuantity(i).SellSationID = .TextMatrix(i, VI_SellStationID)
                atStationQuantity(i).SellStationName = .TextMatrix(i, VI_SellStation)
                atStationQuantity(i).StationID = .TextMatrix(i, VI_StationID)
                atStationQuantity(i).StationName = .TextMatrix(i, VI_Station)
                atStationQuantity(i).TicketType = ResolveDisplay(.TextMatrix(i, VI_TicketType), szTicketTypeName)
                atStationQuantity(i).TicketTypeName = szTicketTypeName
                atStationQuantity(i).VehicleTypeCode = .TextMatrix(i, VI_VehicleTypeID)
                atStationQuantity(i).VehicleTypeName = .TextMatrix(i, VI_VehicleType)
                atStationQuantity(i).Quantity = .TextMatrix(i, VI_Quantity)
                
            End If
        Next i
    End With
    GetStationQuantity = atStationQuantity
    
End Function


Private Sub vsStation_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim oCommDialog As New CommDialog
    Dim aszTemp() As String
    Dim i As Integer
    Dim nCount As Integer
    oCommDialog.Init g_oActiveUser
    
    With vsStation
        Select Case Col
        Case VI_SellStation
            '选择站点   因回程车的上车站就是站点
            aszTemp = oCommDialog.SelectStation()
            If ArrayLength(aszTemp) > 0 Then
                .TextMatrix(Row, VI_SellStation) = aszTemp(1, 2)
                .TextMatrix(Row, VI_SellStationID) = aszTemp(1, 1)
            End If
        Case VI_Route
            '选择回程线路
            aszTemp = oCommDialog.SelectBackRoute(, True)
            If ArrayLength(aszTemp) > 0 Then
                .TextMatrix(Row, VI_Route) = aszTemp(1, 2)
                .TextMatrix(Row, VI_RouteID) = aszTemp(1, 1)
            End If
            
        Case VI_VehicleType
            '选择车型
            aszTemp = oCommDialog.SelectVehicleType()
            If ArrayLength(aszTemp) > 0 Then
                .TextMatrix(Row, VI_VehicleType) = aszTemp(1, 2)
                .TextMatrix(Row, VI_VehicleTypeID) = aszTemp(1, 1)
            End If
        Case VI_Station
            '选择站点
            aszTemp = oCommDialog.SelectStation()
            If ArrayLength(aszTemp) > 0 Then
                .TextMatrix(Row, VI_Station) = aszTemp(1, 2)
                .TextMatrix(Row, VI_StationID) = aszTemp(1, 1)
            End If
        End Select
    End With
        
End Sub



Private Sub tbOperation_ButtonClick(ByVal Button As MSComctlLib.Button)
    With vsStation
        Select Case Button.Key
        Case "INSERT"
            '插入行
            .AddItem "", .Row
            If .Rows > 2 Then
                SetColValue .Row, .Row + 1
            End If
        Case "DELETE"
            '删除行
            .RemoveItem .Row
        Case "APPEND"
            '追加行
            .Rows = .Rows + 1
            
            If .Rows > 2 Then
                SetColValue .Rows - 1, .Rows - 2
            End If
        End Select
    End With
End Sub

Private Sub SetColValue(nDestRow As Integer, nSourceRow)

    With vsStation
        .TextMatrix(nDestRow, VI_SellStation) = .TextMatrix(nSourceRow, VI_SellStation)
        .TextMatrix(nDestRow, VI_Route) = .TextMatrix(nSourceRow, VI_Route)
        .TextMatrix(nDestRow, VI_VehicleType) = .TextMatrix(nSourceRow, VI_VehicleType)
        .TextMatrix(nDestRow, VI_SellStationID) = .TextMatrix(nSourceRow, VI_SellStationID)
        .TextMatrix(nDestRow, VI_RouteID) = .TextMatrix(nSourceRow, VI_RouteID)
        .TextMatrix(nDestRow, VI_VehicleTypeID) = .TextMatrix(nSourceRow, VI_VehicleTypeID)

        
        
    End With
End Sub

Private Function ValidateVSInput() As Boolean
    '验证输入是否有效
    Dim i As Integer
    Dim nCount As Integer
    Dim j As Integer
    ValidateVSInput = False
    With vsStation
        For i = 1 To .Rows - 1
            For j = 1 To .Cols - 1
                nCount = 0
                If .TextMatrix(i, j) = "" Then
                    nCount = nCount + 1
                End If
                
            Next j
            If nCount > 0 And nCount <> VI_Cols Then
                '出错
                MsgBox "第" & i & "行输入不完整", vbInformation, Me.Caption
                .Row = i
                Exit Function
            End If
            If IsNumeric(.TextMatrix(i, VI_Quantity)) Then
                If .TextMatrix(i, VI_Quantity) <= 0 Then
                    MsgBox "数量必须大于零", vbInformation, Me.Caption
                    .Row = i
                    Exit Function
                End If
            Else
                MsgBox "数量必须输入数字", vbInformation, Me.Caption
                .Row = i
                Exit Function
            End If

        Next i
    End With
    ValidateVSInput = True
End Function

Private Sub GetTicketType()
    '得到票种信息
    Dim oSystemParam As New SystemParam
    Dim atTicketType() As TTicketType
    Dim nCount As Integer
    Dim i As Integer
    atTicketType = oSystemParam.GetAllTicketType(1, False)
    nCount = ArrayLength(atTicketType)
    m_szTicketType = ""
    For i = 1 To nCount
        m_szTicketType = m_szTicketType & MakeDisplayString(atTicketType(i).nTicketTypeID, Trim(atTicketType(i).szTicketTypeName)) & "|"
    Next i
    
End Sub
