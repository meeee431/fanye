VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Begin VB.Form frmWizSplitSettle 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "·��������"
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
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame fraWizard 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "��һ��"
      Height          =   5700
      Index           =   1
      Left            =   210
      TabIndex        =   19
      Top             =   930
      Width           =   10755
      Begin VB.CheckBox chkIsToday 
         BackColor       =   &H00E0E0E0&
         Caption         =   "�������"
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
            Name            =   "����"
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
            Name            =   "����"
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
         Caption         =   "���˹�˾(&C):"
         Height          =   180
         Left            =   5055
         TabIndex        =   91
         Top             =   3435
         Width           =   1080
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���㵥���(&S):"
         Height          =   180
         Left            =   1770
         TabIndex        =   8
         Top             =   3450
         Width           =   1260
      End
      Begin VB.Label lblRoute 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��·(&R):"
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
         Caption         =   "����(&S):"
         Height          =   180
         Index           =   1
         Left            =   1770
         TabIndex        =   0
         Top             =   2310
         Width           =   720
      End
      Begin VB.Label label2 
         BackStyle       =   0  'Transparent
         Caption         =   "����(&O):"
         Height          =   180
         Index           =   1
         Left            =   5070
         TabIndex        =   2
         Top             =   2310
         Width           =   810
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "��ʼ����(&B):"
         Height          =   225
         Left            =   1770
         TabIndex        =   4
         Top             =   2895
         Width           =   1380
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "��������(&E):"
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
      Caption         =   "���Ĳ�"
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
            Name            =   "����"
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
            Name            =   "����"
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
            Name            =   "����"
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
               Caption         =   "վ�����"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "ÿ��վ�����"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "վ�������嵥"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fraWizard 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "������"
      Height          =   5670
      Index           =   2
      Left            =   210
      TabIndex        =   21
      Top             =   960
      Width           =   10725
      Begin VB.CheckBox chkSingle 
         BackColor       =   &H00E0E0E0&
         Caption         =   "������"
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
            Name            =   "����"
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
            Name            =   "����"
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
            Name            =   "����"
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
         Caption         =   "������:"
         Height          =   180
         Left            =   4035
         TabIndex        =   57
         Top             =   270
         Width           =   630
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "·������:"
         Height          =   180
         Left            =   7125
         TabIndex        =   50
         Top             =   270
         Width           =   810
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "·����Ч����:"
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
            Name            =   "����"
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
         Caption         =   "��Ч"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "·����:"
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
            Name            =   "����"
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
      Caption         =   "���岽"
      Height          =   5730
      Index           =   5
      Left            =   90
      TabIndex        =   35
      Top             =   900
      Width           =   10785
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "��ע"
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
               Name            =   "����"
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
         Caption         =   "���ν�����ϸ��"
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
         Caption         =   "��˾������ϸ��"
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
         Caption         =   "����������ϸ��"
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
         Caption         =   "·�������ܻ�"
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
            Caption         =   "�����:"
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
            Caption         =   "��Ʊ��:"
            Height          =   180
            Left            =   180
            TabIndex        =   44
            Top             =   300
            Width           =   630
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "·������:"
            Height          =   180
            Left            =   5850
            TabIndex        =   43
            Top             =   300
            Width           =   810
         End
         Begin VB.Label lblSheetCount 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "2��"
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
            Caption         =   "Ӧ��Ʊ��:"
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
            Caption         =   "������:"
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
            Name            =   "����"
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
               Object.ToolTipText     =   "����"
               ImageKey        =   "add"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "ɾ��"
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
         Caption         =   "�ֹ���Ʊ��Ϣ:"
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
      TX              =   "��һ��(&N)"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      TX              =   "���(&F)"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      TX              =   "���·��(&A)"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      TX              =   "�������"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      TX              =   "�ر�(&C)"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      TX              =   "��һ��(&P)"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      TX              =   "����(&H)"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
         Caption         =   "��ѡ��·������ķ�ʽ��"
         Height          =   180
         Left            =   360
         TabIndex        =   33
         Top             =   450
         Width           =   1980
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���㷽ʽ"
         BeginProperty Font 
            Name            =   "����"
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
      Caption         =   "���һ��"
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
         Caption         =   "ִ�����:"
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
            Text            =   "�а����㵥"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblLugProtocol 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "����"
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
            Name            =   "����"
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
            Name            =   "����"
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
            Name            =   "����"
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
         Caption         =   "�а����㵥:"
         Height          =   180
         Left            =   960
         TabIndex        =   75
         Top             =   420
         Width           =   990
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�а����������Ϣ:"
         Height          =   180
         Left            =   4200
         TabIndex        =   74
         Top             =   450
         Width           =   1530
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����Э��:"
         Height          =   180
         Left            =   4230
         TabIndex        =   73
         Top             =   2490
         Width           =   810
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������:"
         Height          =   180
         Left            =   4200
         TabIndex        =   72
         Top             =   1050
         Width           =   810
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ӧ����:"
         Height          =   180
         Left            =   4230
         TabIndex        =   71
         Top             =   2010
         Width           =   810
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�а��˷�:"
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
      Caption         =   "�ڶ���"
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
         Caption         =   "��������(&D):"
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
         Caption         =   "������(&P):"
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
         TX              =   "�Ƴ�<<"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
         TX              =   "���>>"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
            Name            =   "����"
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
            Name            =   "����"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "�а����㵥"
            Object.Width           =   4410
         EndProperty
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ѡ�Ľ��㵥:"
         Height          =   180
         Left            =   4440
         TabIndex        =   67
         Top             =   780
         Width           =   1170
      End
      Begin VB.Label lblSplitObject 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ݹ�˾"
         Height          =   180
         Left            =   1980
         TabIndex        =   64
         Top             =   390
         Width           =   720
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������:"
         Height          =   180
         Left            =   930
         TabIndex        =   63
         Top             =   390
         Width           =   810
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�а����㵥��:"
         Height          =   180
         Left            =   930
         TabIndex        =   62
         Top             =   780
         Width           =   1170
      End
   End
   Begin VB.Menu pmnu_Select 
      Caption         =   "ѡ��"
      Visible         =   0   'False
      Begin VB.Menu pmnu_AllSelect 
         Caption         =   "ȫѡ(&S)"
      End
      Begin VB.Menu pmnu_AllUnSelect 
         Caption         =   "��ѡ(&U)"
      End
      Begin VB.Menu pmnu_SelectCheck 
         Caption         =   "ָ��ȫѡ(&C)"
      End
      Begin VB.Menu pmnu_UnSelectCheck 
         Caption         =   "ָ����ѡ(&N)"
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

'imgcbo ����
Const cnBus = 3
Const cnVehicle = 2
Const cnCompany = 1


'·�����׳���
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


'·��վ����ܵ�����λ��
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


    
'·��վ���嵥������λ��
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
Const VIL_StatusName = 11 '�Ĳ�״̬
Const VIL_RouteID = 12
Const VIL_VehicleTypeID = 13
Const VIL_SeatTypeID = 14
Const VIL_BusSerialNO = 15
Const VIL_StationSerial = 16
Const VIL_AreaRatio = 17
'Const VIL_RouteName = 18
Const VIL_Quantity = 18
Const VIL_Mileage = 19
Const VIL_TicketPrice = 20 '����
Const VIL_TotalTicketPrice = 21 '�ܼ�
Const VIL_StatusCode = 22 '�Ĳ�״̬
Const VIL_BasePrice = 23

'�ֹ���Ʊ��������λ��
Const VIE_NO = 0
Const VIE_Quantity = 1
Const VIE_TotalTicketPrice = 2
Const VIE_Ratio = 3
Const VIE_ServicePrice = 4
Const VIE_SettleOutPrice = 5



Dim m_nSheetCount As Integer  ' ͳ��·������
Dim nValibleCount As Integer  '��Ч��·����
Dim nTotalQuantity As Long '������
Dim m_aszCheckSheetID() As String '·������
Dim m_szVehicleID As String  '��������
Dim m_szCompanyID As String  '��˾ID
Dim m_szBusID As String '����ID
Dim m_szAdditionPrice As Integer  '�����
Dim m_bLogFileValid As Boolean '��־�ļ�
Dim m_bPromptWhenError As Boolean '�Ƿ���ʾ����
Dim CancelHasPress As Boolean
Dim TSplitResult As TSplitResult  'Ԥ�����
Dim m_nSplitItenCount As Integer 'ʹ�õĲ�������
Dim tSplitItem() As TSplitItemInfo
Dim AdditionPrice As String

Dim m_szTransportCompanyName As String
Dim m_szTransportCompanyID As String


Public m_bIsManualSettle As Boolean '�Ƿ����ֹ�����,�����ܳ�����������վ����Ϣ,�Ƿ������޸�
Dim m_rsStationQuantity As Recordset 'վ������������Ϣ
Dim m_rsStationInfo As Recordset 'ÿ�ջ���,��վ����Ϣ

Dim oPriceMan As New STPrice.TicketPriceMan
Dim m_rsPriceItem As Recordset  '���õ�Ʊ����
Dim m_rsStationList As Recordset 'վ��������ϸ��Ϣ

Dim m_nQuantity As Integer '�����޸�ǰ������


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
    '�ж��а����Ƿ���Ч,��������Ƿ�Ϊ��ָ���Ķ���
    m_oReport.Init g_oActiveUser
    
    '��������
    If txtLugSheetID.Text <> "" Then
        Set rsTemp = m_oReport.GetLugSheet(Trim(txtLugSheetID.Text))
        If rsTemp.RecordCount = 0 Then
            MsgBox "���а����㵥��Ч!", vbExclamation, Me.Caption
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
    '���������·������Щ·��Ϊ��Ʊʱ�����������ô��󣬵���δ������·����
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
        '��ʾѡ���·����Ϣ
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
                    '����ҵ����ڸ�·��,����Բ����
                    Exit For
                End If
            
            Next j
            'δ�ҵ���·�����б��д���
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
                lvItem.SubItems(PI_Mileage) = FormatDbValue(rsTemp!Mileage)    '���
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
    
    'Ȩ����֤
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
    
    '��ʼ���ɽ��㵥 ��дlstCreateInfo��Ϣ
    CreateFinanceSheetRs
    '��ӡ���㵥
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
''//***********************************    ������־��  **************************************
''���ɽ��㵥��־
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


'����ĳ���㵥
Private Sub CreateFinanceSheetRs()
    Dim aszSheetID() As String   'Ϊ����Ӧ�ӿ���������
    Dim oVehicle As New Vehicle
    Dim oCompany As New Company
    Dim oBus As New Bus
    Dim szObjectName As String
    Dim szLuggageSettleIDs As String
    Dim i As Integer
    Dim rsTemp As Recordset
    
    On Error GoTo here
    
    
    
    '****ȡ�ö�������
    If imgcbo.Text = cszCompanyName Then
        oCompany.Init g_oActiveUser
        oCompany.Identify m_szCompanyID
        szObjectName = oCompany.CompanyShortName
        '���ǵ��в�����TSplitResult.SettleSheetInfo.SettleStationPrice��ֵ���и���
        TSplitResult.SettleSheetInfo.SettleStationPrice = lvCompany.SelectedItem.SubItems(3)
        '���ǵ��в������������ӵ���Ӧ�ķ�������
        TSplitResult.CompanyInfo(1).SplitItem(m_oSplit.m_nServiceItem) = TSplitResult.CompanyInfo(1).SplitItem(m_oSplit.m_nServiceItem) + Val(txtAdditionPrice.Text)
        
    ElseIf imgcbo.Text = "����" Then
        oVehicle.Init g_oActiveUser
        oVehicle.Identify m_szVehicleID
        szObjectName = szLicenseTagNO 'oVehicle.LicenseTag
        '���ǵ��в�����TSplitResult.SettleSheetInfo.SettleStationPrice��ֵ���и���
        TSplitResult.SettleSheetInfo.SettleStationPrice = lvVehicle.SelectedItem.SubItems(3)
        '���ǵ��в������������ӵ���Ӧ�ķ�������
        TSplitResult.VehicleInfo(1).SplitItem(m_oSplit.m_nServiceItem) = TSplitResult.VehicleInfo(1).SplitItem(m_oSplit.m_nServiceItem) + Val(txtAdditionPrice.Text)
        
    ElseIf imgcbo.Text = "����" Then
'        oBus.Init g_oActiveUser
'        oBus.Identify m_szBusID
        szObjectName = m_szBusID
        TSplitResult.SettleSheetInfo.TransportCompanyID = m_szTransportCompanyID
        TSplitResult.SettleSheetInfo.TransportCompanyName = m_szTransportCompanyName
        m_szCompanyID = m_szTransportCompanyID
        '���ǵ��в�����TSplitResult.SettleSheetInfo.SettleStationPrice��ֵ���и���
        TSplitResult.SettleSheetInfo.SettleStationPrice = lvBus.SelectedItem.SubItems(3)
        '���ǵ��в������������ӵ���Ӧ�ķ�������
        TSplitResult.BusInfo(1).SplitItem(m_oSplit.m_nServiceItem) = TSplitResult.BusInfo(1).SplitItem(m_oSplit.m_nServiceItem) + Val(txtAdditionPrice.Text)
    End If
    TSplitResult.SettleSheetInfo.ObjectName = szObjectName

    '****ȡ���а������Ϣ
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
    
'    '��ʼ����
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
        '�жϲ��������û��ѡ
        If txtObject.Text = "" Then
            MsgBox " ��ѡ��������!", vbInformation, Me.Caption
            cmdPrevious.Enabled = False
            SetNormal
            Exit Sub
        End If

        '�жϽ��㵥���Ƿ�Ϊ��
        If txtSheetID.Text = "" Then
            MsgBox " ���㵥��Ų���Ϊ��!", vbInformation, Me.Caption
            SetNormal
            Exit Sub
        End If
        If chkIsToday.Value = vbChecked Then
            '����ǵ������,�����ڲ��ܴ���һ��
            If ToDBDate(dtpStartDate.Value) <> ToDBDate(dtpEndDate.Value) Then
                MsgBox "���Ϊ�������,��ʼ������������ڱ�����ͬ", vbExclamation, Me.Caption
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
        
        
        
        
        
        '���Ҫ�����·����¼
        FilllvObject
        '��ʾǩ��������
        m_nSheetCount = lvObject.ListItems.Count
        lblSettleSheetCount.Caption = CStr(m_nSheetCount)
        nValibleCount = 0 'ɨ��·������Ч����
        nTotalQuantity = 0
        lblEnableCount.Caption = nValibleCount
        lblTotalQuantity.Caption = nTotalQuantity
        txtCheckSheetID.Text = ""
        '�ȴ�ɨ��·��
        lblStatus.Visible = False
        txtCheckSheetID.Text = ""
        txtCheckSheetID.SetFocus
        
        cmdAddSheet.Visible = True
        cmdClear.Visible = True
        
    ElseIf fraWizard(2).Visible = True Then
        If lvObject.ListItems.Count > 0 Then
            If nValibleCount = 0 Then
                MsgBox "��ɨ��·����ѡ��򹴣�", vbExclamation, Me.Caption
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
        '      ���·������
        FillSheetID
        '���·����վ��ͳ����Ϣ
        
        vsStationDayList.Visible = False
        vsStationList.Visible = False
        vsStationTotal.Visible = True
        FillCheckSheetStationTotal
        '��ʾÿ�յ�·��ͳ��
        FillCheckSheetStationDayList
        FillCheckSheetStationList
    ElseIf fraWizard(3).Visible = True Then
        '�ֹ�Ʊ��Ϣ¼��
        
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
        '���·��Ԥ����Ϣ
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

'�����а�������Ϣ
Private Sub HandleLugInfo()
    lblSplitObject.Caption = ResolveDisplayEx(txtObject.Text)
    txtLugSheetID.Text = ""
    lvLugSheetID.ListItems.Clear
    cmdDetele.Enabled = False
    
End Sub

'�����а�����ͳ��
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

'���Ҫ�����·����¼
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
    '���·����Ϣ
    If imgcbo.ComboItems(cnCompany).Selected Then      '��ת��
        Set rsTemp = m_oReport.GetNeedSplitCheckSheet(CS_SettleByTransportCompany, ResolveDisplay(Trim(txtObject.Text)), dtpStartDate.Value, dtpEndDate.Value, ResolveDisplay(txtRouteID.Text), aszTemp, , IIf(chkIsToday.Value = vbChecked, True, False))
    ElseIf imgcbo.ComboItems(cnVehicle).Selected Then
        Set rsTemp = m_oReport.GetNeedSplitCheckSheet(CS_SettleByVehicle, ResolveDisplay(Trim(txtObject.Text)), dtpStartDate.Value, dtpEndDate.Value, ResolveDisplay(txtRouteID.Text), aszTemp, , IIf(chkIsToday.Value = vbChecked, True, False))
    ElseIf imgcbo.ComboItems(cnBus).Selected Then
        Set rsTemp = m_oReport.GetNeedSplitCheckSheet(CS_SettleByBus, ResolveDisplay(Trim(txtObject.Text)), dtpStartDate.Value, dtpEndDate.Value, ResolveDisplay(txtRouteID.Text), aszTemp, ResolveDisplay(cboCompany.Text), IIf(chkIsToday.Value = vbChecked, True, False))
        
    End If
    FilllvObjectHead '�������
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
        lvItem.SubItems(PI_Mileage) = FormatDbValue(rsTemp!Mileage)    '���
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

' �������
Private Sub FilllvObjectHead()

   With lvObject.ColumnHeaders
         .Clear
         .Add , , "·����", 1200
         .Add , , "��վ��", 900
         .Add , , "����", 800
         .Add , , "����", 700
         .Add , , "���", 700
         .Add , , "��Ʊ��", 800
         .Add , , "���ƺ�", 900
         .Add , , "���˹�˾", 900
         .Add , , "����", 800
         .Add , , "�����", 800
         .Add , , "��·", 900
         .Add , , "��Ʊ��", 0
         .Add , , "��������", 0
         .Add , , "���˹�˾����", 0
   End With
       AlignHeadWidth Me.name, lvObject
End Sub
'���·��,����,��˾����
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
   If imgcbo.Text = cszCompanyName Then   'ת������
        m_szCompanyID = ResolveDisplay(txtObject.Text)
        m_szVehicleID = ""
        m_szBusID = ""
   ElseIf imgcbo.Text = "����" Then
        m_szVehicleID = ResolveDisplay(txtObject.Text)
        m_szCompanyID = ""
        m_szBusID = ""
   ElseIf imgcbo.Text = "����" Then
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
    '�õ�����ʹ�õ�Ʊ����
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
    
    '  ���� ��ʼ��cbo����
    '   1-���� 2-���˹�˾
    With imgcbo
        .ComboItems.Clear
        .ComboItems.Add , , "���˹�˾", 1
        .ComboItems.Add , , "����", 3
        .ComboItems.Add , , "����", 4
        .Locked = True
        If szSplitType <> "" Then
            If szSplitType = "���˹�˾" Then
                .ComboItems(1).Selected = True
            ElseIf szSplitType = "����" Then
                .ComboItems(2).Selected = True
            ElseIf szSplitType = "����" Then
                .ComboItems(3).Selected = True
            End If
        Else
            .ComboItems(2).Selected = True
        End If
    End With
    
    '   �Զ����ɽ��㵥�� YYYYMM0001��ʽ
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
        '����ѡ��
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
        '��˾ѡ��
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
        '����ѡ��
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
 MsgBox lstCreateInfo.Text, vbInformation + vbOKOnly, "������Ϣ"
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

'��ʾ�а����㵥����
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
        '��ʾ������Ϣ
        vsStationTotal.Visible = True
        vsStationDayList.Visible = False
        vsStationList.Visible = False
        
    ElseIf tsStation.Tabs(2).Selected Then
        
        '��ʾ��ϸ��Ϣ
        vsStationTotal.Visible = False
        vsStationDayList.Visible = True
        vsStationList.Visible = False
    ElseIf tsStation.Tabs(3).Selected Then
        
        '��ʾ��ϸ��Ϣ
        vsStationTotal.Visible = False
        vsStationDayList.Visible = False
        vsStationList.Visible = True
    End If
End Sub

Private Sub txtAdditionPrice_Change()
Dim i As Integer
   If imgcbo.Text = cszCompanyName Then   'ת������
        For i = 1 To lvCompany.ColumnHeaders.Count
            If Trim(tSplitItem(m_oSplit.m_nServiceItem).SplitItemName) = Trim(lvCompany.ColumnHeaders.Item(i)) Then
                lvCompany.SelectedItem.SubItems(i - 1) = AdditionPrice + Val(txtAdditionPrice.Text)
            End If
        Next i
        lvCompany.SelectedItem.SubItems(3) = TSplitResult.SettleSheetInfo.SettleStationPrice + Val(txtAdditionPrice.Text)
        lvCompany.SelectedItem.SubItems(2) = TSplitResult.BusInfo(1).SettlePrice - lvCompany.SelectedItem.SubItems(3)
        lblNeedSplitMoney.Caption = lvCompany.SelectedItem.SubItems(2)
   ElseIf imgcbo.Text = "����" Then
        For i = 1 To lvVehicle.ColumnHeaders.Count
            If Trim(tSplitItem(m_oSplit.m_nServiceItem).SplitItemName) = Trim(lvVehicle.ColumnHeaders.Item(i)) Then
                lvVehicle.SelectedItem.SubItems(i - 1) = AdditionPrice + Val(txtAdditionPrice.Text)
            End If
        Next i
        lvVehicle.SelectedItem.SubItems(3) = TSplitResult.SettleSheetInfo.SettleStationPrice + Val(txtAdditionPrice.Text)
        lvVehicle.SelectedItem.SubItems(2) = TSplitResult.BusInfo(1).SettlePrice - lvVehicle.SelectedItem.SubItems(3)
        lblNeedSplitMoney.Caption = lvVehicle.SelectedItem.SubItems(2)
   ElseIf imgcbo.Text = "����" Then
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
    ElseIf Trim(imgcbo.Text) = "����" Then
        aszTemp = oShell.SelectVehicleEX()
    ElseIf Trim(imgcbo.Text) = "����" Then
        aszTemp = oShell.SelectBus()
    End If
    
    Set oShell = Nothing
    If ArrayLength(aszTemp) = 0 Then Exit Sub
    
    If Trim(imgcbo.Text) = "����" Then
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
        Set rsTemp = m_oCheckSheet.CheckSheetAvailable(Trim(txtCheckSheetID.Text))   '��֤·����Ч��,�����Ч������ʾ
        If rsTemp.RecordCount = 0 Then
            PlayEventSound g_tEventSoundPath.CheckSheetNotExist '��·��������
            MsgBox "��·��������", vbExclamation, Me.Caption
            imgEnabled.Visible = True
            Exit Sub
        ElseIf FormatDbValue(rsTemp!valid_mark) = 0 Then
            
        
            MsgBox "��·���ѷ�", vbExclamation, Me.Caption
            PlayEventSound g_tEventSoundPath.CheckSheetCanceled  '��·���ѷ�
            txtCheckSheetID.SelStart = 0
            txtCheckSheetID.SelLength = Len(txtCheckSheetID.Text)
            Exit Sub
        ElseIf FormatDbValue(rsTemp!settlement_status) = 1 Then
            MsgBox "��·���ѽ���", vbExclamation, Me.Caption
            PlayEventSound g_tEventSoundPath.CheckSheetSettled '��·���ѽ���
            txtCheckSheetID.SelStart = 0
            txtCheckSheetID.SelLength = Len(txtCheckSheetID.Text)
            Exit Sub
        Else
            PlayEventSound g_tEventSoundPath.ObjectNotSame   '��·����Ч,��������Ҫ����ʱ������Χ֮��
            MsgBox "��·����Ч,��������Ҫ����ʱ������Χ֮��", vbExclamation, Me.Caption
            txtCheckSheetID.SelStart = 0
            txtCheckSheetID.SelLength = Len(txtCheckSheetID.Text)
            Exit Sub
        End If
        
    Else
        If lvObject.ListItems(i).Checked Then
            PlayEventSound g_tEventSoundPath.CheckSheetSelected '��·����ѡ
'            MsgBox "��·����ѡ��", vbInformation, Me.Caption
            txtCheckSheetID.SelStart = 0
            txtCheckSheetID.SelLength = Len(txtCheckSheetID.Text)
            
            Exit Sub
        End If
        
        
    End If
    
    'ɨ��·���ɹ�
    PlayEventSound g_tEventSoundPath.CheckSheetValid '��Ч·��
    '�б��д���Ч·��
    imgEnabled.Visible = False
    For i = 1 To lvObject.ListItems.Count
        If Trim(txtCheckSheetID.Text) = Trim(lvObject.ListItems.Item(i).Text) Then
            lvObject.ListItems.Item(i).EnsureVisible
            lvObject.ListItems.Item(i).Selected = True
            If lvObject.ListItems.Item(i).Checked = False Then
                lvObject.ListItems.Item(i).Checked = True
                    'ɨ��һ��,��Ч������һ
                    nValibleCount = nValibleCount + 1
                    '�������ۼ�
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
    '�������Ϊ����,��������еĸó��εĹ�˾
    Dim oBus As New Bus
    Dim nCount As Integer
    Dim aszCompany() As String
    Dim i As Integer
    On Error GoTo ErrorHandle
    If Trim(imgcbo.Text) = "����" Then
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


'������·��Ԥ����Ϣ
Public Sub FillPreSheetInfo()
    Dim i As Integer
    Dim nCount As Integer
    Dim lvItem As ListItem
    
    Dim atStationQuantity() As TSettleSheetStation
    
    
    Dim rsPassageNumber As Recordset
    
    
    
    
    '���������ļ�¼��
    If imgcbo.Text = cszCompanyName Then
        '����˾����
        Set rsPassageNumber = MakeRecordSetListByCompany
    ElseIf imgcbo.Text = cszVehicleName Then
        Set rsPassageNumber = MakeRecordSetListByVehicle
    ElseIf imgcbo.Text = cszBusName Then
        Set rsPassageNumber = MakeRecordSetListByBus
    End If
    
    
    '����Ԥ��
    m_oSplit.Init g_oActiveUser
    TSplitResult = m_oSplit.PreviewSplitCheckSheetEx(m_szCompanyID, m_szVehicleID, m_aszCheckSheetID, rsPassageNumber, dtpStartDate.Value, DateAdd("d", 1, dtpEndDate.Value), m_atExtraInfo, IIf(chkSingle.Value = vbChecked, True, False), m_szBusID)

    lblTotalPrice.Caption = IIf(TSplitResult.SettleSheetInfo.TotalTicketPrice = 0, "��", TSplitResult.SettleSheetInfo.TotalTicketPrice)
    
    
    lblNeedSplitMoney.Caption = Format(TSplitResult.SettleSheetInfo.SettleOtherCompanyPrice - TSplitResult.SettleSheetInfo.SettleStationPrice, "0.00")
    lblTotalQuautity.Caption = TSplitResult.SettleSheetInfo.TotalQuantity
    lblSheetCount.Caption = TSplitResult.SettleSheetInfo.CheckSheetCount
    
    '�������
    lvHead
    '��乫˾�б�
    FilllvCompany
    '��䳵���б�
    FilllvVehicle
    '��䳵���б�
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
                lvItem.SubItems(k) = TSplitResult.CompanyInfo(j).SplitItem(i) '�������
                                
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
                lvItem.SubItems(k) = TSplitResult.VehicleInfo(j).SplitItem(i) '�������
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
                lvItem.SubItems(k) = TSplitResult.BusInfo(j).SplitItem(i) '�������
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
        .Add , , "���˹�˾"
        .Add , , "Э��"
        .Add , , "Ӧ��Ʊ��"
        .Add , , "�����վƱ��"
        .Add , , "����"
        .Add , , "�˹���"
        
    End With
    
    With lvVehicle.ColumnHeaders
        .Clear
        .Add , , "����"
        .Add , , "Э��"
        .Add , , "Ӧ��Ʊ��"
        .Add , , "�����վƱ��"
        .Add , , "����"
        .Add , , "�˹���"
        
    End With
    
    With lvBus.ColumnHeaders
        .Clear
        .Add , , "����"
        .Add , , "Э��"
        .Add , , "Ӧ��Ʊ��"
        .Add , , "�����վƱ��"
        .Add , , "����"
        .Add , , "�˹���"
        
    End With
    

    'ȡ��ʹ�õĲ������
    m_oReport.Init g_oActiveUser
    tSplitItem = m_oReport.GetSplitItemInfo()
    nSplitItenCount = ArrayLength(tSplitItem)
    m_nSplitItenCount = 0
    If nSplitItenCount = 0 Then Exit Sub
    For i = 1 To nSplitItenCount
        If tSplitItem(i).SplitStatus <> CS_SplitItemNotUse Then
            lvCompany.ColumnHeaders.Add , , tSplitItem(i).SplitItemName
            lvVehicle.ColumnHeaders.Add , , tSplitItem(i).SplitItemName 'ͬʱ���ӳ����б�Ĳ�����
            lvBus.ColumnHeaders.Add , , tSplitItem(i).SplitItemName 'ͬʱ���ӳ����б�Ĳ�����
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
    '��ʾ·����
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
    '�ֹ����ɼ�¼��
    'nStationCount Ϊ�ܹ��е�վ����Ŀ.
    
    Dim rsTemp As New Recordset
    Dim i As Integer
    Dim j As Integer
    
    
    Dim aszStation() As String
    Dim nStationCount As Integer 'վ��Ʊ����
    
    Dim szBusDate As String
    Dim szStationID As String 'վ��
'    Dim szStationName As String 'վ��
    Dim szTicketTypeID As String 'Ʊ��
'    Dim szTicketTypeName As String 'Ʊ����
    Dim alNumber() As Long '����
    Dim lTotalNum As Long
    nStationCount = prsStation.RecordCount
    '����ֶε�����,Ϊ�Ƚ���.
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
            '�����¼��
            
            rsTemp.AddNew
            rsTemp!bus_date = szBusDate
            lTotalNum = 0
            For j = 1 To nStationCount
                rsTemp.Fields("station_" & j) = alNumber(j)
                lTotalNum = lTotalNum + alNumber(j)
                
            Next j
            rsTemp.Fields("total_num") = lTotalNum
            
            rsTemp.Update
            
            '���ԭֵ
            For j = 1 To nStationCount
                alNumber(j) = 0
            Next j

            '���ó��εĳ�ʼֵ
                    
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
            '�����ͬ
            
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
                '�޸�վ����ܵ�����
                For i = 1 To vsStationTotal.Rows - 1
                    '�ҵ��Ծ͵���,���������������޸�
                    If vsStationTotal.TextMatrix(i, VI_SellStationID) = .TextMatrix(Row, VIL_SellStationID) _
                            And vsStationTotal.TextMatrix(i, VI_RouteID) = .TextMatrix(Row, VIL_RouteID) _
                            And vsStationTotal.TextMatrix(i, VI_VehicleTypeID) = .TextMatrix(Row, VIL_VehicleTypeID) _
                            And vsStationTotal.TextMatrix(i, VI_StationID) = .TextMatrix(Row, VIL_StationID) _
                            And vsStationTotal.TextMatrix(i, VI_TicketTypeID) = .TextMatrix(Row, VIL_TicketType) _
                            Then
                            
                        '�ҵ���,���������е���
                        vsStationTotal.TextMatrix(i, VI_Quantity) = Val(vsStationTotal.TextMatrix(i, VI_Quantity)) + Val(.Text) - Val(m_nQuantity)
                        vsStationTotal.Row = i
                        vsStationTotal.Col = VI_Quantity
                        vsStationTotal.CellBackColor = vbRed
                        '�޸��ܼ�
                        .TextMatrix(Row, VIL_TotalTicketPrice) = Val(.TextMatrix(Row, VIL_Quantity)) * Val(.TextMatrix(Row, VIL_TicketPrice))
                        Exit For
                    End If
                Next i
                
                
                '�޸�վ��ÿ�ջ��ܵ�����
                For i = 1 To vsStationDayList.Rows - 1
                    If ToDBDate(vsStationDayList.TextMatrix(i, 0)) = ToDBDate(.TextMatrix(Row, VIL_BusDate)) Then
                        '�Ӽ�¼���в��ҳ�վ����뼰Ʊ�������бȽ�
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
    '�����޸�ǰ������
    If Col = VIL_Quantity Then
        m_nQuantity = vsStationList.Text
    End If
End Sub

Private Sub vsStationTotal_BeforeSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long, Cancel As Boolean)
    '���ܴ��������޸�����
    
    
'    If NewColSel = VI_Quantity Then
'        '���Ϊ����,�������޸�
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
'        '���Ϊ����,�������޸�
'        vsStationList.Editable = flexEDKbdMouse
'    Else
'        vsStationList.Editable = flexEDNone
'    End If
End Sub

Private Function QuantityHasBeChanged() As Boolean
    'վ��������Ϣ�Ƿ񱻸ı��
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
        'δ���ı��
        QuantityHasBeChanged = False
    Else
        '���Ķ���
        QuantityHasBeChanged = True
    End If
    
End Function

Private Function GetStationQuantity() As TSettleSheetStation()
    '�õ�vsStationTotal�е�վ��������Ϣ
    Dim atStationQuantity() As TSettleSheetStation
    Dim i As Integer
    ReDim atStationQuantity(1 To m_rsStationQuantity.RecordCount)
    
    '����վ������
    '********�˴�Ӧ����Ѳ�ͬ�ĳ��λ��ܳ�ͬһ��.
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
    '���·��վ�������Ϣ
    Dim oSplit As New Split
'    Dim m_rsStationQuantity As Recordset
    On Error GoTo ErrorHandle
    Dim i As Integer
    Dim nCount As Integer
    Dim j As Integer
    SetBusy
    ShowSBInfo "�������·��վ�������Ϣ"
'·��վ����ܵ�����λ��
    
    
    With vsStationTotal
        .Clear
        .Rows = 2
        .Cols = 14
        .FixedCols = 1
        .FixedRows = 1
        .ColWidth(0) = 100
        '���úϲ�
        .MergeCells = flexMergeRestrictColumns
        .MergeCol(VI_SellStation) = True
        .MergeCol(VI_Route) = True
        .MergeCol(VI_Bus) = True
        .MergeCol(VI_VehicleType) = True
        .MergeCol(VI_Station) = True
        .MergeCol(VI_TicketType) = False
        
        
        .AllowUserResizing = flexResizeColumns
        
        
        .TextMatrix(0, VI_SellStation) = "�ϳ�վ"
        .TextMatrix(0, VI_Route) = "��·"
        .TextMatrix(0, VI_Bus) = "����"
        .TextMatrix(0, VI_VehicleType) = "����"
        .TextMatrix(0, VI_Station) = "վ��"
        .TextMatrix(0, VI_TicketType) = "Ʊ��"
        .TextMatrix(0, VI_Quantity) = "����"
        .TextMatrix(0, VI_AreaRatio) = "�������"
        
        '�������Ϣ����ֻ�ڻ���ʱ�õ�
        .TextMatrix(0, VI_RouteID) = "��·����"
        .TextMatrix(0, VI_VehicleTypeID) = "���ʹ���"
        .TextMatrix(0, VI_SellStationID) = "�ϳ�վ����"
        .TextMatrix(0, VI_StationID) = "վ�����"
        .TextMatrix(0, VI_TicketTypeID) = "Ʊ�ִ���"
        
        AlignHeadWidth Me.name, vsStationTotal
        .ColWidth(VI_AreaRatio) = 0
        .ColWidth(VI_Bus) = 0
        '�������Ϣ����ֻ�ڻ���ʱ�õ�
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
            '���β���ʾ
            '.TextMatrix(i, VI_Bus) = FormatDbValue(m_rsStationQuantity!bus_id)
            .TextMatrix(i, VI_Station) = FormatDbValue(m_rsStationQuantity!station_name)
            .TextMatrix(i, VI_TicketType) = FormatDbValue(m_rsStationQuantity!ticket_type_name)
            .TextMatrix(i, VI_VehicleType) = FormatDbValue(m_rsStationQuantity!vehicle_type_name)
            .TextMatrix(i, VI_Quantity) = FormatDbValue(m_rsStationQuantity!Quantity)
            .TextMatrix(i, VI_AreaRatio) = Val(FormatDbValue(m_rsStationQuantity!Annotation))
            
            
            '�������Ϣ����ֻ�ڻ���ʱ�õ�
            
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

    '���·��վ��ÿ��������Ϣ
    '�轫���ݿ��в�����ļ�¼��תΪ��ͷ�ǵ�վ����ͷ�����ڣ�������������
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
    ShowSBInfo "�������·��վ��ÿ��������Ϣ"
    oSplit.Init g_oActiveUser
    FillSheetID
    
    'debug.Print "GetCheckSheetDistinctStation start:" & Time
    Set m_rsStationInfo = oSplit.GetCheckSheetDistinctStation(m_aszCheckSheetID, IIf(chkIsToday.Value = vbChecked, True, False))
    
    'debug.Print "GetCheckSheetDistinctStation end:" & Time
    nCols = m_rsStationInfo.RecordCount
    
    
    'debug.Print "TotalCheckSheetStationInfoEx start:" & Time
    Set rsStationQuantityList = oSplit.TotalCheckSheetStationInfoEx(m_aszCheckSheetID, IIf(chkIsToday.Value = vbChecked, True, False))
    
    'debug.Print "TotalCheckSheetStationInfoEx end:" & Time
    '����¼������ת��
    Set rsTemp = MakeRecordSetDayList(rsStationQuantityList, m_rsStationInfo)
    
    
    nCount = rsTemp.RecordCount
    
    With vsStationDayList
    
    
        .Rows = nCount + 1
        .Cols = nCols + 1 + 1 '���˸�С��
        
        .FixedCols = 1
        .FixedRows = 1
        .ColWidth(0) = 1000
        
        .Clear
        
        .AllowUserResizing = flexResizeColumns
        '���վ��
        .TextMatrix(0, 0) = "����\����\��վ"
        m_rsStationInfo.MoveFirst
        For i = 1 To nCols
            .TextMatrix(0, i) = FormatDbValue(m_rsStationInfo!station_name) & "(" & GetUnicodeBySize(FormatDbValue(m_rsStationInfo!ticket_type_name), 2) & ")"
            m_rsStationInfo.MoveNext
        Next i
        .TextMatrix(0, i) = "С��"
        
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
'            '���β���ʾ
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
    '���·��վ����ϸ
    Dim oSplit As New Split
'    Dim m_rsStationList As Recordset
    On Error GoTo ErrorHandle
    Dim i As Integer
    Dim nCount As Integer
    Dim j As Integer
    SetBusy

    ShowSBInfo "�������·��վ����ϸ"
    
    
    With vsStationList
        '1:825,2:1080,3:570,4:645,5:720,6:540,7:0,8:0,9:0,10:0,11:615,12:0,13:0,14:0,15:0,16:0,17:0,18:480,19:540,20:585,21:690,22:0,23:720,24:720,25:720,26:720,27:720,28:720,29:720,30:750,31:720,32:720,

        'quantity mileage  ticket_price base_carriage price_item_1 price_item_2 price_item_3 price_item_4 price_item_5 price_item_6 price_item_7 price_item_8 price_item_9 price_item_10 price_item_11 price_item_12 price_item_13 price_item_14 price_item_15 sell_station_name ticket_type_name seat_type_name
        '������������������

        .Rows = 2
        .Cols = VIL_BasePrice + m_rsPriceItem.RecordCount
        
        
        .ExplorerBar = flexExSortShowAndMove  '�����������ͷ����
        
        .FrozenCols = 0 '���ö������
        .AllowUserFreezing = flexFreezeColumns '����������������е�λ��
        .AllowUserResizing = flexResizeColumns
        
        .FixedCols = 1
        .FixedRows = 1
        .ColWidth(0) = 100
        .Clear
        '���úϲ�
        .MergeCells = flexMergeRestrictColumns
        
        
        .MergeCol(VIL_CheckSheetID) = True
        .MergeCol(VIL_BusDate) = True
        .MergeCol(VIL_BusID) = True
        .MergeCol(VIL_SellStationName) = True
        .MergeCol(VIL_StationName) = True
        
        
        
    
    
    
        .TextMatrix(0, VIL_CheckSheetID) = "·����"
        .TextMatrix(0, VIL_BusDate) = "����"
        .TextMatrix(0, VIL_BusID) = "����"
        .TextMatrix(0, VIL_SellStationName) = "�ϳ�վ"
        .TextMatrix(0, VIL_StationName) = "վ��"
        .TextMatrix(0, VIL_TicketTypeName) = "Ʊ��"
        
        .TextMatrix(0, VIL_PriceIdentify) = "��"
        .TextMatrix(0, VIL_SellStationID) = "�ϳ�վ����"
        .TextMatrix(0, VIL_StationID) = "վ�����"
        .TextMatrix(0, VIL_TicketType) = "Ʊ�ִ���"
        .TextMatrix(0, VIL_StatusName) = "�Ĳ�״̬"
        
        .TextMatrix(0, VIL_Quantity) = "����"
        .TextMatrix(0, VIL_Mileage) = "���"
        .TextMatrix(0, VIL_TicketPrice) = "����"
        .TextMatrix(0, VIL_TotalTicketPrice) = "�ܼ�"
        
           
        '�������Ϣ����ֻ�ڻ���ʱ�õ�
        .TextMatrix(0, VIL_RouteID) = "��·����"
        .TextMatrix(0, VIL_VehicleTypeID) = "���ʹ���"
        
     
        '������Ϣ�ڱ���վ����ϸʱ�õ�
        
        .TextMatrix(0, VIL_SeatTypeID) = "λ��"
        .TextMatrix(0, VIL_BusSerialNO) = "�������"
        .TextMatrix(0, VIL_StationSerial) = "վ�����"
        '�ڼ���Ӧ�����ʱ�õ�
        .TextMatrix(0, VIL_AreaRatio) = "��������"
        .TextMatrix(0, VIL_StatusCode) = "�Ĳ�״̬����"
        
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
            .TextMatrix(i, VIL_SellStationID) = FormatDbValue(m_rsStationList!sell_station_id2) '��ʾ����վ�����
        
        
            .TextMatrix(i, VIL_StationID) = FormatDbValue(m_rsStationList!station_id)
            .TextMatrix(i, VIL_TicketType) = FormatDbValue(m_rsStationList!ticket_type)
            .TextMatrix(i, VIL_StatusName) = GetSheetStationStatusName(FormatDbValue(m_rsStationList!Status))
            .TextMatrix(i, VIL_StatusCode) = FormatDbValue(m_rsStationList!Status)
            '������������,����ʾʣ�������,������ʾ���е�����
            If Not g_oParam.AllowSplitBySomeTimes Then
                .TextMatrix(i, VIL_Quantity) = FormatDbValue(m_rsStationList!Quantity)
            Else
                .TextMatrix(i, VIL_Quantity) = FormatDbValue(m_rsStationList!Quantity) - FormatDbValue(m_rsStationList!fact_quantity)
            End If
            .TextMatrix(i, VIL_Mileage) = Val(FormatDbValue(m_rsStationList!Mileage))
            .TextMatrix(i, VIL_TicketPrice) = Val(FormatDbValue(m_rsStationList!ticket_price))
            .TextMatrix(i, VIL_TotalTicketPrice) = .TextMatrix(i, VIL_Quantity) * .TextMatrix(i, VIL_TicketPrice)
            .TextMatrix(i, VIL_BasePrice) = Val(FormatDbValue(m_rsStationList!base_carriage))
            
            
            '�������Ϣ����ֻ�ڻ���ʱ�õ�
            .TextMatrix(i, VIL_RouteID) = FormatDbValue(m_rsStationList!route_id)
            
'            .TextMatrix(i, VIL_RouteName) = FormatDbValue(m_rsStationList!route_name)
            
            .TextMatrix(i, VIL_VehicleTypeID) = FormatDbValue(m_rsStationList!vehicle_type_code)
        
            
            '������Ϣ�ڱ���վ����ϸʱ�õ�
            .TextMatrix(i, VIL_SeatTypeID) = FormatDbValue(m_rsStationList!seat_type_id)
            .TextMatrix(i, VIL_BusSerialNO) = FormatDbValue(m_rsStationList!bus_serial_no)
            .TextMatrix(i, VIL_StationSerial) = FormatDbValue(m_rsStationList!station_serial)
            '�ڼ���Ӧ�����ʱ�õ�
            .TextMatrix(i, VIL_AreaRatio) = FormatDbValue(m_rsStationList!Annotation)
        
        
            m_rsPriceItem.MoveFirst
            For j = 1 To m_rsPriceItem.RecordCount
                If Val(FormatDbValue(m_rsPriceItem!price_item)) = 0 Then
                    '���Ϊ�����˼�,���Թ�
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

'���ɰ���˾����ļ�¼��
Private Function MakeRecordSetListByCompany() As Recordset

    Dim rsTemp As New Recordset
    Dim i As Integer
    Dim j As Integer
    
    Dim oCompany As New Company
    
    
    
    '������ʱ����
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
        .Append "annotation", adVarChar, 255   '��ŵ��ǵ����ķ���
        .Append "quantity", adBigInt
        .Append "mileage", adDouble
        .Append "ticket_price", adCurrency
        .Append "base_carriage", adCurrency
        
        For i = 1 To cnPriceItemNum
            .Append "price_item_" & i, adCurrency
        Next i
        
        .Append "check_sheet_count", adBigInt '·����
    End With
    
    rsTemp.Open
    '�ź���,�ٽ��л���
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
        
        '�õ���˾���뼰����
        '��Ϊ�ǰ���˾����,���Զ����ڵľ��ǹ�˾����
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
        'szRouteName = .TextMatrix(i,vil_route  '��ʱ������
        szSellStationName = .TextMatrix(i, VIL_SellStationName)
        szStationName = .TextMatrix(i, VIL_StationName)
'        szVehicleTypeName = .TextMatrix(i,vil_ve '��ʱ������
        szTicketTypeName = .TextMatrix(i, VIL_TicketTypeName)
        szAnnotation = .TextMatrix(i, VIL_AreaRatio)
'        lQuantity = .TextMatrix(i, VIL_Quantity)
'        dbMileage = .TextMatrix(i, VIL_Mileage) * lQuantity '�������Ҫ��������
'        dbTicketPrice = .TextMatrix(i, VIL_TotalTicketPrice) 'ȡ��Ʊ��,�ѳ�������
        
    
        For i = 1 To .Rows - 1
            '�����¼��
            
            
            '��ͬʱ,�½�һ��¼
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
        
            
                
                '���ԭֵ
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
                'szRouteName = .TextMatrix(i,vil_route  '��ʱ������
                szSellStationName = .TextMatrix(i, VIL_SellStationName)
                szStationName = .TextMatrix(i, VIL_StationName)
        '        szVehicleTypeName = .TextMatrix(i,vil_ve '��ʱ������
                szTicketTypeName = .TextMatrix(i, VIL_TicketTypeName)
                szAnnotation = .TextMatrix(i, VIL_AreaRatio)
                
                
                
                
                '��ȥ��Ʊ��,������Ҫ��������
                lQuantity = lQuantity + .TextMatrix(i, VIL_Quantity)
                dbMileage = dbMileage + .TextMatrix(i, VIL_Mileage) * .TextMatrix(i, VIL_Quantity)
                dbTicketPrice = dbTicketPrice + .TextMatrix(i, VIL_TicketPrice)
                dbBasePrice = dbBasePrice + .TextMatrix(i, VIL_BasePrice) * .TextMatrix(i, VIL_Quantity)
                
                '��Ʊ����
                m_rsPriceItem.MoveFirst
                For j = 1 To m_rsPriceItem.RecordCount
                    If Val(FormatDbValue(m_rsPriceItem!price_item)) = 0 Then
                        '���Ϊ�����˼�,���Թ�
                    Else
                        adbPriceItem(Val(FormatDbValue(m_rsPriceItem!price_item))) = adbPriceItem(Val(FormatDbValue(m_rsPriceItem!price_item))) + .TextMatrix(i, VIL_BasePrice + j - 1) * .TextMatrix(i, VIL_Quantity)
                    End If
                    m_rsPriceItem.MoveNext
                Next j
                
            Else
                '�����ͬ
                
                '��ȥ��Ʊ��,������Ҫ��������
                lQuantity = lQuantity + .TextMatrix(i, VIL_Quantity)
                dbMileage = dbMileage + .TextMatrix(i, VIL_Mileage) * .TextMatrix(i, VIL_Quantity)
                dbTicketPrice = dbTicketPrice + .TextMatrix(i, VIL_TicketPrice)
                dbBasePrice = dbBasePrice + .TextMatrix(i, VIL_BasePrice) * .TextMatrix(i, VIL_Quantity)
                
'                If .TextMatrix(i, VIL_CheckSheetID) <> szCheckSheetID Then
'                    szCheckSheetID = .TextMatrix(i, VIL_CheckSheetID)
'                    lSheetCount = lSheetCount + 1
'                End If
                
                Dim bFind As Boolean
                '��Ϊδ��·������,��������ǰ����.
                
                bFind = False
                For j = i - 1 To 1 Step -1
                    If .TextMatrix(j, VIL_CheckSheetID) = .TextMatrix(i, VIL_CheckSheetID) Then
                        bFind = True
                        Exit For
                    End If
                Next j
                If Not bFind Then
                    '���δ�ҵ�,���ۼ�
                    lSheetCount = lSheetCount + 1
                End If
                '��Ʊ����
                m_rsPriceItem.MoveFirst
                For j = 1 To m_rsPriceItem.RecordCount
                    If Val(FormatDbValue(m_rsPriceItem!price_item)) = 0 Then
                        '���Ϊ�����˼�,���Թ�
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



'���ɰ���������ļ�¼��
Private Function MakeRecordSetListByVehicle() As Recordset

    Dim rsTemp As New Recordset
    Dim i As Integer
    Dim j As Integer
    
    
    Dim oVehicle As New Vehicle
    
    
    '������ʱ����
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
        .Append "annotation", adVarChar, 255   '��ŵ��ǵ����ķ���
        .Append "quantity", adBigInt
        .Append "mileage", adDouble
        .Append "ticket_price", adCurrency
        .Append "base_carriage", adCurrency
        
        For i = 1 To cnPriceItemNum
            .Append "price_item_" & i, adCurrency
        Next i
    
        .Append "check_sheet_count", adBigInt '·����
        
    End With

    rsTemp.Open
    '�ź���,�ٽ��л���
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
        
        '�õ���˾���뼰����
        '��Ϊ�ǰ���˾����,���Զ����ڵľ��ǹ�˾����
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
'        szRouteName = .TextMatrix(i, VI_Route) '��ʱ������
        szSellStationName = .TextMatrix(i, VIL_SellStationName)
        szStationName = .TextMatrix(i, VIL_StationName)
'        szVehicleTypeName = .TextMatrix(i,vil_ve '��ʱ������
        szTicketTypeName = .TextMatrix(i, VIL_TicketTypeName)
        szAnnotation = .TextMatrix(i, VIL_AreaRatio)
        
        
        For i = 1 To .Rows - 1
            '�����¼��
            
            
            '��ͬʱ,�½�һ��¼
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
        
            
                '���ԭֵ
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
                'szRouteName = .TextMatrix(i,vil_route  '��ʱ������
                szSellStationName = .TextMatrix(i, VIL_SellStationName)
                szStationName = .TextMatrix(i, VIL_StationName)
        '        szVehicleTypeName = .TextMatrix(i,vil_ve '��ʱ������
                szTicketTypeName = .TextMatrix(i, VIL_TicketTypeName)
                szAnnotation = .TextMatrix(i, VIL_AreaRatio)
                
                
                szCheckSheetID = .TextMatrix(i, VIL_CheckSheetID)
'                lSheetCount = lSheetCount + 1
                
                lQuantity = lQuantity + .TextMatrix(i, VIL_Quantity)
                dbMileage = dbMileage + .TextMatrix(i, VIL_Mileage) * .TextMatrix(i, VIL_Quantity) '�������Ҫ��������
                dbTicketPrice = dbTicketPrice + .TextMatrix(i, VIL_TotalTicketPrice) 'ȡ��Ʊ��,�ѳ�������
                dbBasePrice = dbBasePrice + .TextMatrix(i, VIL_BasePrice) * .TextMatrix(i, VIL_Quantity)
                
                '��Ʊ����
                m_rsPriceItem.MoveFirst
                For j = 1 To m_rsPriceItem.RecordCount
                    If Val(FormatDbValue(m_rsPriceItem!price_item)) = 0 Then
                        '���Ϊ�����˼�,���Թ�
                    Else
                        adbPriceItem(Val(FormatDbValue(m_rsPriceItem!price_item))) = adbPriceItem(Val(FormatDbValue(m_rsPriceItem!price_item))) + .TextMatrix(i, VIL_BasePrice + j - 1) * .TextMatrix(i, VIL_Quantity)
                    End If
                    m_rsPriceItem.MoveNext
                Next j
                
                
            Else
                '�����ͬ
                lQuantity = lQuantity + .TextMatrix(i, VIL_Quantity)
                dbMileage = dbMileage + .TextMatrix(i, VIL_Mileage) * .TextMatrix(i, VIL_Quantity) '�������Ҫ��������
                dbTicketPrice = dbTicketPrice + .TextMatrix(i, VIL_TotalTicketPrice) 'ȡ��Ʊ��,�ѳ�������
                dbBasePrice = dbBasePrice + .TextMatrix(i, VIL_BasePrice) * .TextMatrix(i, VIL_Quantity)
                
'                If .TextMatrix(i, VIL_CheckSheetID) <> szCheckSheetID Then
'                    szCheckSheetID = .TextMatrix(i, VIL_CheckSheetID)
'                    lSheetCount = lSheetCount + 1
'                End If
                
                Dim bFind As Boolean
                '��Ϊδ��·������,��������ǰ����.
                
                bFind = False
                For j = i - 1 To 1 Step -1
                    If .TextMatrix(j, VIL_CheckSheetID) = .TextMatrix(i, VIL_CheckSheetID) Then
                        bFind = True
                        Exit For
                    End If
                Next j
                If Not bFind Then
                    '���δ�ҵ�,���ۼ�
                    lSheetCount = lSheetCount + 1
                End If
                
                
                
                '��Ʊ����
                m_rsPriceItem.MoveFirst
                For j = 1 To m_rsPriceItem.RecordCount
                    If Val(FormatDbValue(m_rsPriceItem!price_item)) = 0 Then
                        '���Ϊ�����˼�,���Թ�
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

'���ɰ����β���ļ�¼��
Private Function MakeRecordSetListByBus() As Recordset

    Dim rsTemp As New Recordset
    Dim i As Integer
    Dim j As Integer

'    Dim oBus As New Bus
    
    
    '������ʱ����
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
        .Append "annotation", adVarChar, 255   '��ŵ��ǵ����ķ���
        .Append "quantity", adBigInt
        .Append "mileage", adDouble
        .Append "ticket_price", adCurrency
        .Append "base_carriage", adCurrency
        
        For i = 1 To cnPriceItemNum
            .Append "price_item_" & i, adCurrency
        Next i
        
        .Append "check_sheet_count", adBigInt '·����
        
    End With

    rsTemp.Open
    '�ź���,�ٽ��л���
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
        
        '�õ���˾���뼰����
        '��Ϊ�ǰ���˾����,���Զ����ڵľ��ǹ�˾����
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
'        szRouteName = .TextMatrix(i, VI_Route) '��ʱ������
        szSellStationName = .TextMatrix(i, VIL_SellStationName)
        szStationName = .TextMatrix(i, VIL_StationName)
'        szVehicleTypeName = .TextMatrix(i,vil_ve '��ʱ������
        szTicketTypeName = .TextMatrix(i, VIL_TicketTypeName)
        szAnnotation = .TextMatrix(i, VIL_AreaRatio)
        
        
        For i = 1 To .Rows - 1
            '�����¼��
            
            
            '��ͬʱ,�½�һ��¼
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
        
            
                '���ԭֵ
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
                'szRouteName = .TextMatrix(i,vil_route  '��ʱ������
                szSellStationName = .TextMatrix(i, VIL_SellStationName)
                szStationName = .TextMatrix(i, VIL_StationName)
        '        szVehicleTypeName = .TextMatrix(i,vil_ve '��ʱ������
                szTicketTypeName = .TextMatrix(i, VIL_TicketTypeName)
                szAnnotation = .TextMatrix(i, VIL_AreaRatio)
                
                szCheckSheetID = .TextMatrix(i, VIL_CheckSheetID)
'                lSheetCount = lSheetCount + 1
                lQuantity = lQuantity + .TextMatrix(i, VIL_Quantity)
                dbMileage = dbMileage + .TextMatrix(i, VIL_Mileage) * .TextMatrix(i, VIL_Quantity) '�������Ҫ��������
                dbTicketPrice = dbTicketPrice + .TextMatrix(i, VIL_TotalTicketPrice) 'ȡ��Ʊ��,�ѳ�������
                dbBasePrice = dbBasePrice + .TextMatrix(i, VIL_BasePrice) * .TextMatrix(i, VIL_Quantity)
                
                '��Ʊ����
                m_rsPriceItem.MoveFirst
                For j = 1 To m_rsPriceItem.RecordCount
                    If Val(FormatDbValue(m_rsPriceItem!price_item)) = 0 Then
                        '���Ϊ�����˼�,���Թ�
                    Else
                        adbPriceItem(Val(FormatDbValue(m_rsPriceItem!price_item))) = adbPriceItem(Val(FormatDbValue(m_rsPriceItem!price_item))) + .TextMatrix(i, VIL_BasePrice + j - 1) * .TextMatrix(i, VIL_Quantity)
                    End If
                    m_rsPriceItem.MoveNext
                Next j
                
                
            Else
                '�����ͬ
                lQuantity = lQuantity + .TextMatrix(i, VIL_Quantity)
                dbMileage = dbMileage + .TextMatrix(i, VIL_Mileage) * .TextMatrix(i, VIL_Quantity) '�������Ҫ��������
                dbTicketPrice = dbTicketPrice + .TextMatrix(i, VIL_TotalTicketPrice) 'ȡ��Ʊ��,�ѳ�������
                dbBasePrice = dbBasePrice + .TextMatrix(i, VIL_BasePrice) * .TextMatrix(i, VIL_Quantity)
                
'                If .TextMatrix(i, VIL_CheckSheetID) <> szCheckSheetID Then
'                    szCheckSheetID = .TextMatrix(i, VIL_CheckSheetID)
'                    lSheetCount = lSheetCount + 1
'                End If
                Dim bFind As Boolean
                '��Ϊδ��·������,��������ǰ����.
                
                bFind = False
                For j = i - 1 To 1 Step -1
                    If .TextMatrix(j, VIL_CheckSheetID) = .TextMatrix(i, VIL_CheckSheetID) Then
                        bFind = True
                        Exit For
                    End If
                Next j
                If Not bFind Then
                    '���δ�ҵ�,���ۼ�
                    lSheetCount = lSheetCount + 1
                End If
                
                '��Ʊ����
                m_rsPriceItem.MoveFirst
                For j = 1 To m_rsPriceItem.RecordCount
                    If Val(FormatDbValue(m_rsPriceItem!price_item)) = 0 Then
                        '���Ϊ�����˼�,���Թ�
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
    '��������
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
            '��ʼ��
            For j = 1 To 15
                rsTemp.Fields("price_item_" & j) = 0
            Next j
            
            '��Ʊ����
            m_rsPriceItem.MoveFirst
            For j = 1 To m_rsPriceItem.RecordCount
                If Val(FormatDbValue(m_rsPriceItem!price_item)) = 0 Then
                    '���Ϊ�����˼�,���Թ�
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
'    '��������
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
    '��ʼ���ֹ���Ʊ�ı��
    Dim i As Integer
    With vsExtra
        .Cols = 6
        .Rows = 8
        .Clear
        
        .TextMatrix(0, VIE_NO) = "��"
        .TextMatrix(0, VIE_Quantity) = "����"
        .TextMatrix(0, VIE_TotalTicketPrice) = "Ʊ��"
        .TextMatrix(0, VIE_Ratio) = "����ѱ���"
        .TextMatrix(0, VIE_ServicePrice) = "�����"
        .TextMatrix(0, VIE_SettleOutPrice) = "������"
        '�����
        For i = 1 To .Rows - 1
            .TextMatrix(i, 0) = i
        Next i
    End With
    
End Sub

Private Function ValidateVsExtra() As Boolean
    '��֤������Ч��   ,����ֵ
    Dim i As Integer
    Dim nCount As Integer
    Dim atExtraInfo() As TSettleExtraInfo
    
    With vsExtra
        
        '��֤������Ч��
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, VIE_Quantity)) <> 0 Or Val(.TextMatrix(i, VIE_TotalTicketPrice)) <> 0 Then
                '����������
                If Val(.TextMatrix(i, VIE_Quantity)) = 0 Or Val(.TextMatrix(i, VIE_TotalTicketPrice)) = 0 Then
                    MsgBox "��" & i & "��������������,�����������һ����Ϊ0", vbExclamation, Me.Caption
                    '����Ч������
                    ValidateVsExtra = False
                    Exit Function
                    
                End If
                nCount = nCount + 1
            End If
        Next i
        '��ֵ
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
    '������Ч
    ValidateVsExtra = True
        
End Function


Private Sub vsExtra_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    '��������������֤,���Զ�����
    With vsExtra
        If Col = VIE_Quantity Or Col = VIE_Ratio Or Col = VIE_TotalTicketPrice Then

'            If .TextMatrix(Row, VIE_Quantity) = "" Then .TextMatrix(Row, VIE_Quantity) = 0
'            If .TextMatrix(Row, VIE_Ratio) = "" Then .TextMatrix(Row, VIE_Ratio) = 0
'            If .TextMatrix(Row, VIE_TotalTicketPrice) = "" Then .TextMatrix(Row, VIE_TotalTicketPrice) = 0
            
            '��ȷ��Ԫ
            .TextMatrix(Row, VIE_ServicePrice) = FormatMoney(Val(.TextMatrix(Row, VIE_TotalTicketPrice)) * Val(.TextMatrix(Row, VIE_Ratio)) / 100)
            
            .TextMatrix(Row, VIE_SettleOutPrice) = FormatMoney(Val(.TextMatrix(Row, VIE_TotalTicketPrice)) - Val(.TextMatrix(Row, VIE_ServicePrice)))
            
        End If
    End With
End Sub


Private Sub tbSection_ButtonClick(ByVal Button As MSComctlLib.Button)
    '���ֹ���Ʊ��������ı����в���
    Dim i As Integer
    Select Case Button.Key
    Case "add"
        '����һ��
        vsExtra.Rows = vsExtra.Rows + 1
        vsExtra.TextMatrix(vsExtra.Rows - 1, 0) = vsExtra.Rows - 1
        
    Case "del"
        'ɾ��һ��
        If vsExtra.Rows = 1 Then Exit Sub
        
        If (MsgBox("ȷ��Ҫɾ����" & vsExtra.Rows - 1 & "����", vbQuestion + vbYesNo, Me.Caption) = vbYes) Then
            
            vsExtra.RemoveItem vsExtra.Row
            For i = 1 To vsExtra.Rows - 1
                vsExtra.TextMatrix(i, 0) = i
            Next i
            
        End If
    End Select
    
End Sub




Private Sub vsExtra_BeforeSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long, Cancel As Boolean)
    '�������룬��ģ���Ӧ����NewRowSel������NewColSel��FPD
    If NewColSel <> VIE_ServicePrice And NewColSel <> VIE_SettleOutPrice Then
        vsExtra.Editable = flexEDKbdMouse
    Else
        vsExtra.Editable = flexEDNone
    End If
End Sub


Private Sub vsExtra_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '�ж��Ƿ�Ϊ����
    With vsExtra
        If Col > 0 And .TextMatrix(Row, Col) <> "" Then
            If Not IsNumeric(.TextMatrix(Row, Col)) Then
                '�����������,�����
                MsgBox "������������", vbExclamation, Me.Caption
            Else
                If .TextMatrix(Row, Col) < 0 Then
                    '���С��0,�����
                    MsgBox "�������0", vbExclamation, Me.Caption
                End If
                If IsNumeric(.TextMatrix(Row, VIE_Ratio)) Then
                    If .TextMatrix(Row, VIE_Ratio) > 100 Then
                        MsgBox "���ʱ���С��100", vbExclamation, Me.Caption
                    End If
                End If
            End If
        End If
    End With
End Sub









