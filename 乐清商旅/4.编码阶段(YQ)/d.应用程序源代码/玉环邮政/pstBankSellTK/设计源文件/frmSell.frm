VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{695ABF14-B2D8-11D2-A5ED-DE08DCF33612}#1.2#0"; "asfcombo.ocx"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{BBF95DAA-F9CB-4CA9-A673-E0E9E0193752}#1.0#0"; "STSellCtl.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmSell 
   BackColor       =   &H00929292&
   Caption         =   "售票"
   ClientHeight    =   8160
   ClientLeft      =   2265
   ClientTop       =   1785
   ClientWidth     =   11880
   HelpContextID   =   4000040
   Icon            =   "frmSell.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8160
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrConnected 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   3150
      Top             =   7290
   End
   Begin MSWinsockLib.Winsock wsClient 
      Left            =   2520
      Top             =   7335
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1245
      HelpContextID   =   3000411
      Left            =   4650
      TabIndex        =   41
      Top             =   7500
      Visible         =   0   'False
      Width           =   2880
      Begin VB.CheckBox chkSetSeat 
         BackColor       =   &H00FFFFFF&
         Caption         =   "定座(&P)"
         Height          =   270
         HelpContextID   =   3000411
         Left            =   120
         TabIndex        =   42
         Top             =   -30
         Width           =   975
      End
      Begin VB.Label lblSetSeat 
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   180
         TabIndex        =   43
         Top             =   60
         Width           =   435
      End
   End
   Begin VB.Frame fraOutLine 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   7035
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   11835
      Begin PSTBankSellTK.ucSuperCombo cboEndStation 
         Height          =   1755
         Left            =   90
         TabIndex        =   5
         Top             =   1620
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   3096
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.ComboBox cboStartStation 
         Height          =   300
         Left            =   945
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   930
         Width           =   1575
      End
      Begin STSellCtl.ucNumTextBox txtPreferentialSell 
         Height          =   390
         Left            =   6030
         TabIndex        =   16
         Top             =   2970
         Visible         =   0   'False
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   2
         Alignment       =   1
      End
      Begin FCmbo.asFlatCombo cboPreferentialTicket 
         Height          =   330
         Left            =   4740
         TabIndex        =   15
         Top             =   3000
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         ButtonDisabledForeColor=   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonHotBackColor=   8421504
         ButtonPressedBackColor=   0
         Text            =   ""
         ButtonBackColor =   8421504
         Style           =   1
         Registered      =   -1  'True
         OfficeXPColors  =   -1  'True
      End
      Begin MSComCtl2.UpDown ucPreferential 
         Height          =   390
         Left            =   6900
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   2970
         Visible         =   0   'False
         Width           =   255
         _ExtentX        =   423
         _ExtentY        =   688
         _Version        =   393216
         BuddyControl    =   "txtPreferentialSell"
         BuddyDispid     =   196650
         OrigLeft        =   2370
         OrigTop         =   3090
         OrigRight       =   2625
         OrigBottom      =   3405
         Max             =   100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   1745027080
         Enabled         =   -1  'True
      End
      Begin VB.CheckBox chkInsurance 
         BackColor       =   &H00E0E0E0&
         Caption         =   "保险(F11)"
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   3825
         TabIndex        =   23
         Top             =   5025
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.TextBox txtSeat 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         HelpContextID   =   3000411
         Left            =   7470
         TabIndex        =   21
         Top             =   45
         Width           =   2865
      End
      Begin VB.CommandButton cmdSetSeat 
         Caption         =   "定座(&G)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         HelpContextID   =   3000411
         Left            =   10545
         TabIndex        =   22
         Top             =   75
         Width           =   990
      End
      Begin VB.TextBox txtReceivedMoney 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   1170
         TabIndex        =   12
         Text            =   "0.0"
         Top             =   5730
         Width           =   1455
      End
      Begin VB.Timer Timer1 
         Interval        =   60000
         Left            =   4800
         Top             =   510
      End
      Begin RTComctl3.CoolButton cmdSell 
         Height          =   705
         Left            =   120
         TabIndex        =   10
         Top             =   4500
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   1244
         BTYPE           =   3
         TX              =   "售出(&P)"
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
         MICON           =   "frmSell.frx":014A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin STSellCtl.ucNumTextBox txtHalfSell 
         Height          =   390
         Left            =   1410
         TabIndex        =   14
         Top             =   3870
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   2
         Alignment       =   1
      End
      Begin STSellCtl.ucNumTextBox txtFullSell 
         Height          =   390
         Left            =   1410
         TabIndex        =   9
         Top             =   3435
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   2
         Alignment       =   1
      End
      Begin FText.asFlatSpinEdit txtPrevDate 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   510
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonHotBackColor=   -2147483633
         ButtonPressedBackColor=   -2147483627
         Text            =   "asFlatSpinEdit1"
         ButtonBackColor =   -2147483633
         Registered      =   -1  'True
         OfficeXPColors  =   -1  'True
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   3420
         Top             =   1920
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
               Picture         =   "frmSell.frx":0166
               Key             =   "StopBus"
            EndProperty
         EndProperty
      End
      Begin RTComctl3.FlatLabel flblSellDate 
         Height          =   345
         Left            =   930
         TabIndex        =   26
         Top             =   510
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   14737632
         OutnerStyle     =   2
         Caption         =   ""
      End
      Begin RTComctl3.FlatLabel flblLimitedTime 
         Height          =   315
         Left            =   8010
         TabIndex        =   27
         Top             =   4620
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   12632256
         OutnerStyle     =   2
         Caption         =   ""
      End
      Begin FCmbo.asFlatCombo cboSeatType 
         Height          =   330
         Left            =   3600
         TabIndex        =   19
         Top             =   3210
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         ButtonDisabledForeColor=   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
         ButtonHotBackColor=   8421504
         ButtonPressedBackColor=   0
         Text            =   ""
         ButtonBackColor =   8421504
         Style           =   1
         Registered      =   -1  'True
         OfficeXPColors  =   -1  'True
      End
      Begin RTComctl3.FlatLabel flblLimitedCount 
         Height          =   315
         Left            =   300
         TabIndex        =   28
         Top             =   6795
         Visible         =   0   'False
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   14737632
         OutnerStyle     =   2
         NormTextColor   =   16711680
         Caption         =   ""
      End
      Begin MSComctlLib.ListView lvSellStation 
         Height          =   1605
         Left            =   5145
         TabIndex        =   25
         Top             =   6450
         Visible         =   0   'False
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   2831
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "上车站"
            Object.Width           =   1942
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   " 时间"
            Object.Width           =   1696
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   " 票价"
            Object.Width           =   1677
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "上车站代码"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "检票门"
            Object.Width           =   0
         EndProperty
      End
      Begin MSComCtl2.UpDown upFull 
         Height          =   390
         Left            =   2280
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   3435
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   688
         _Version        =   393216
         BuddyControl    =   "txtFullSell"
         BuddyDispid     =   196655
         OrigLeft        =   2370
         OrigTop         =   3090
         OrigRight       =   2625
         OrigBottom      =   3405
         Max             =   100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   1745027080
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown upHalf 
         Height          =   390
         Left            =   2280
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   3870
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   688
         _Version        =   393216
         BuddyControl    =   "txtHalfSell"
         BuddyDispid     =   196654
         OrigLeft        =   2370
         OrigTop         =   3090
         OrigRight       =   2625
         OrigBottom      =   3405
         Max             =   100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   1745027080
         Enabled         =   -1  'True
      End
      Begin MSComctlLib.ListView lvBus 
         Height          =   6060
         Left            =   2640
         TabIndex        =   7
         Top             =   510
         Width           =   9090
         _ExtentX        =   16034
         _ExtentY        =   10689
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ColHdrIcons     =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "起点站:"
         Height          =   180
         Left            =   120
         TabIndex        =   2
         Top             =   990
         Width           =   630
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "座位号(&T):"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5895
         TabIndex        =   20
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label lblsellstation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "上车站(&O):"
         Height          =   180
         Left            =   7680
         TabIndex        =   24
         Top             =   6105
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label flblRestMoney 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.0"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   435
         Left            =   1845
         TabIndex        =   40
         Top             =   6450
         Width           =   720
      End
      Begin VB.Label flblTotalPrice 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.0"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   435
         Left            =   1845
         TabIndex        =   39
         Top             =   5250
         Width           =   720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         Visible         =   0   'False
         X1              =   2790
         X2              =   5430
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Label lblSeatType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "座型"
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
         Left            =   2880
         TabIndex        =   18
         Top             =   3255
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label lblHalfSell 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "半票(&X):"
         Enabled         =   0   'False
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
         Left            =   120
         TabIndex        =   13
         Top             =   3975
         Width           =   960
      End
      Begin VB.Label lblFullSell 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "全票(&A):"
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
         Left            =   120
         TabIndex        =   8
         Top             =   3540
         Width           =   960
      End
      Begin VB.Label lblSellMsg 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
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
         Height          =   255
         Left            =   9330
         TabIndex        =   38
         Top             =   6660
         Width           =   2220
      End
      Begin VB.Label lblTotalMoney 
         AutoSize        =   -1  'True
         Caption         =   "120.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9720
         TabIndex        =   37
         Top             =   2460
         Width           =   1320
      End
      Begin VB.Line Line5 
         X1              =   8760
         X2              =   11250
         Y1              =   3180
         Y2              =   3180
      End
      Begin VB.Line Line2 
         X1              =   8730
         X2              =   11220
         Y1              =   2250
         Y2              =   2250
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "总票价:"
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
         Left            =   8640
         TabIndex        =   36
         Top             =   2610
         Width           =   840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "站票:"
         Height          =   180
         Left            =   5190
         TabIndex        =   35
         Top             =   2880
         Width           =   705
      End
      Begin VB.Label Label1 
         Caption         =   "限售张数:"
         Height          =   225
         Left            =   300
         TabIndex        =   34
         Top             =   1770
         Width           =   1110
      End
      Begin VB.Label lblToStation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "到站(&Z):"
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
         Left            =   120
         TabIndex        =   4
         Top             =   1290
         Width           =   960
      End
      Begin VB.Label lblPrevDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "预售天数(&D):"
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
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Width           =   1440
      End
      Begin VB.Label lblBus 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "车次列表(&V):"
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
         Left            =   2640
         TabIndex        =   6
         Top             =   135
         Width           =   1440
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "车票单价:"
         Height          =   180
         Left            =   5970
         TabIndex        =   33
         Top             =   2310
         Width           =   810
      End
      Begin VB.Label lblSinglePrice 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0.00/0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   300
         Left            =   7560
         TabIndex        =   32
         Top             =   2130
         Width           =   1215
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C0C0C0&
         Visible         =   0   'False
         X1              =   30
         X2              =   2670
         Y1              =   4380
         Y2              =   4380
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "总票价:"
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
         Left            =   90
         TabIndex        =   31
         Top             =   5340
         Width           =   840
      End
      Begin VB.Label lblReceivedMoney 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "实收(&Q):"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   90
         TabIndex        =   11
         Top             =   5880
         Width           =   1170
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00000000&
         X1              =   30
         X2              =   2670
         Y1              =   6390
         Y2              =   6390
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "应找票款:"
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
         Left            =   90
         TabIndex        =   30
         Top             =   6540
         Width           =   1080
      End
      Begin VB.Label lblmileate 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1200
         TabIndex        =   29
         Top             =   900
         Visible         =   0   'False
         Width           =   180
      End
   End
   Begin VB.Label lblTotalPrice 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
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
      ForeColor       =   &H80000007&
      Height          =   525
      Left            =   0
      TabIndex        =   44
      Top             =   -15
      Visible         =   0   'False
      Width           =   645
   End
End
Attribute VB_Name = "frmSell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*******************************************************************
'# Last Modify:2005-12-2 By 陆勇庆
'# Last Modify At:加入了对全鼠标操作的支持，体现为：
'#                a.增加了upFull、upHalf、upPreferential控件
'#                b.新增lvBus_DblClick的支持，以及增加了upFull、upHalf、upPreferential控件
'#                c.当控件上有数量而待售票列表中没有时，则自动出售控件上数量的车票
'#                d.将操作流程重新设计了一个，将售票按钮提前（相关的txtReceviMoney.setfocus更改，更符合绍兴的操作习惯）
'#                e.恢复了停班车次的列出，并整行变红
'*******************************************************************
'Private m_bPrint As Boolean
Private m_blPointCount As Boolean
Private m_bSumPriceIsEmpty As Boolean   '总票价为0
Private m_nCount As Integer '隔一段时间读取服务器时间的自加一的变量
Private m_sgTotalMoney As Single '记录上一次售票的金额
Private m_atTicketType() As TTicketType
Private m_dbTotalPrice As Double
Private m_aszSeatType() As String
'Private m_atbSeatTypeBus As TMultiSeatTypeBus '得到不同座位类型的车次
Private m_TicketPrice() As Single '存储票价
Private m_TicketTypeDetail() As ETicketType '存储票种
Private m_bPreClear As Boolean
Private m_bSetFocus As Boolean
Private m_bPreSellFocus As Boolean
Private m_rsBusInfo As Recordset
Private m_atbBusOrder() As TBusOrderCount

Private m_bNotRefresh As Boolean '是否需要刷新,主要是在设置查询车次时间时用到.
'Private m_rsBusInfo As Recordset
'Private m_nSellCount As Integer '出售张数

Private m_szSend As String
Private m_szAllSend As String

Private m_szStartStationID As String '起点站
Public rsCheckGate As Recordset '检票口


Private Sub cboEndStation_Change()
On Error GoTo Here
    If lvBus.ListItems.count > 0 Then
        DoThingWhenBusChange
    Else
       lvSellStation.ListItems.Clear
    End If
    DealPrice
    
'    cmdPreSell.Enabled = True
On Error GoTo 0
Exit Sub
Here:
  ShowErrorMsg
End Sub

Private Sub cboEndStation_GotFocus()
On Error GoTo Here
    lblToStation.ForeColor = clActiveColor
    DealPrice
'    '********************
'    '语音显示器
'        SetUser GetActiveUserID
'        'SetWhere
'    '********************

On Error GoTo 0
Exit Sub
Here:
 ShowErrorMsg
End Sub



Private Sub cboEndStation_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
    Case vbKeyLeft
        KeyCode = 0
        If Val(txtPrevDate.Text) > 0 Then
            txtPrevDate.Text = Val(txtPrevDate.Text) - 1
        End If
        m_bNotRefresh = False
    Case vbKeyRight
        KeyCode = 0
        If Val(txtPrevDate.Text) < m_nCanSellDay Then
        
            txtPrevDate.Text = Val(txtPrevDate.Text) + 1
        End If
        m_bNotRefresh = False
    Case vbKeyReturn
        'lvBus.SetFocus
        SendTab
    Case Else
        m_bNotRefresh = False
    End Select
    If m_bPreClear Then
        flblTotalPrice.Caption = 0#
        txtReceivedMoney.Text = ""
        flblRestMoney.Caption = ""
        m_bPreClear = False
    End If
   
Exit Sub
Here:
    ShowErrorMsg
End Sub

Private Sub cboEndStation_LostFocus()

    On Error GoTo Here
    Dim nIndex As Integer
        lblToStation.ForeColor = 0
        If m_bNotRefresh Then Exit Sub '如果是跳到了输入时间处,则不刷新
        
        txtReceivedMoney.Text = ""
        SendBusRequest True
'
'            If m_bPreClear Then
'                flblTotalPrice.Caption = 0#
'                txtReceivedMoney.Text = ""
'                flblRestMoney.Caption = ""
'                m_bPreClear = False
'            End If
'            DealPrice
'            If cboEndStation.Text <> "" Then
'                SetInsurance
'            End If
    On Error GoTo 0
Exit Sub
Here:
    ShowErrorMsg
End Sub

Private Sub cboPreferentialTicket_Change()
    txtPreferentialSell.Text = 0
'    cmdPreSell.Enabled = True
End Sub

Private Sub cboPreferentialTicket_Click()
    txtPreferentialSell.Text = 0
End Sub


Private Sub cboPreferentialTicket_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyRight
              txtPreferentialSell.SetFocus
              KeyCode = 0
        Case vbKeyLeft
              txtHalfSell.SetFocus
              KeyCode = 0
    End Select
    
End Sub




Private Sub cboSeatType_Change()
  If lvSellStation.ListItems.count > 0 Then
   RefreshBusStation m_rsBusInfo, Trim(lvSellStation.SelectedItem.SubItems(3)), cboSeatType.ListIndex + 1
  End If
End Sub

Private Sub cboSeatType_GotFocus()
    lblSeatType.ForeColor = clActiveColor
End Sub

Private Sub cboSeatType_KeyPress(KeyAscii As Integer)
'    lvBus.SetFocus
End Sub

Private Sub cboSeatType_LostFocus()
    lblSeatType.ForeColor = 0
End Sub

Private Sub cboStartStation_GotFocus()
    m_szStartStationID = ResolveDisplay(cboStartStation.Text)
    
End Sub

Private Sub cboStartStation_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyLeft
            KeyCode = 0
            If Val(txtPrevDate.Text) > 0 Then
                txtPrevDate.Text = Val(txtPrevDate.Text) - 1
            End If
            m_bNotRefresh = False
        Case vbKeyRight
            KeyCode = 0
            If Val(txtPrevDate.Text) < m_nCanSellDay Then
            
                txtPrevDate.Text = Val(txtPrevDate.Text) + 1
            End If
            m_bNotRefresh = False
'        Case 106 'Asc("*")
            '+号则跳到输入时间处
'            KeyCode = 0
'            txtTime.SetFocus
'            m_bNotRefresh = True
        Case Else
            m_bNotRefresh = False
    End Select
   
End Sub

Private Sub cboStartStation_LostFocus()
    '刷新站点信息及车次信息
    '如果起点站更改,则刷新站点
    If m_szStartStationID <> ResolveDisplay(cboStartStation.Text) Then
        RefreshStation2
    End If
End Sub

Private Sub chkInsurance_Click()
    DealPrice
End Sub

Private Sub cmdPreSell_Click()
'    On Error GoTo Here
'    Dim nSameBusIndex As Integer
'    If txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text = 0 Then Exit Sub
''    cmdPreSell.Enabled = False
''If nSameBusIndex = 0 Then
''    GetPreSellTicketInfo
''Else
''    MergeSameBusInfo nSameBusIndex
''End If
'    txtFullSell.Text = 0
'    txtHalfSell.Text = 0
'    txtPreferentialSell.Text = 0
'    SetPreSellButton
'    DealPrice
''    cmdPreSell.Enabled = True
'Exit Sub
'Here:
'    ShowErrorMsg
End Sub

'售票
Private Sub cmdSell_Click()
    Dim k As Integer
    
    If Val(txtFullSell.Text) + Val(txtHalfSell.Text) + Val(txtPreferentialSell.Text) + Val(m_lTicketNo) - 1 > Val(m_lEndTicketNo) Then
        k = Val(m_lEndTicketNo) - Val(m_lTicketNo) + 1
        MsgBox "打印机上的票已不够！" & vbCrLf & "车票只剩 " & k & "张", vbInformation, "售票台"
    Else
        SellTicket
    End If
'        SellTicket
End Sub



'
'Private Sub Command1_Click()
'    frmNotify.Show vbModal
'End Sub


Private Sub flblTotalPrice_Change()
    CalReceiveMoney
End Sub

Private Sub CalReceiveMoney()
    Dim dbTemp As Double
    dbTemp = Val(txtReceivedMoney.Text) - Val(flblTotalPrice.Caption)
    If Val(dbTemp) <= 0 Then
        flblRestMoney.Caption = ""
    Else
        flblRestMoney.Caption = FormatMoney(dbTemp)
    End If
End Sub

Private Sub Form_Activate()
On Error GoTo Here
    m_nCurrentTask = RT_SellTicket
    m_szCurrentUnitID = Me.Tag
    SetPreSellButton
    MDISellTicket.SetFunAndUnit
'    m_oSell.SellUnitCode = Me.Tag

'    lvBus.SortKey = MDISellTicket.GetSortKey()
    SetDefaultSellTicket
    DealDiscountAndSeat
'--------------------------
    MDISellTicket.EnableSortAndRefresh True
    cboEndStation.SetFocus
    On Error GoTo 0
'    lblTicketNo.Caption = MDISellTicket.lblTicketNo.Caption
    
'    'MDISellTicket.abMenuTool.Bands("mnuFunction").Tools("mnuChangeSeatType").Enabled = True
    
Exit Sub
Here:
    ShowErrorMsg
End Sub

Private Sub Form_Deactivate()
On Error GoTo Here
    MDISellTicket.EnableSortAndRefresh False
'    'MDISellTicket.abMenuTool.Bands("mnuFunction").Tools("mnuChangeSeatType").Enabled = False
On Error GoTo 0
Exit Sub
Here:
    ShowErrorMsg
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    DealWithChildKeyDown KeyCode, Shift
'    If KeyCode = vbKeyF9 Then
'        ChangeSeatType
'    End If
    If KeyCode = vbKeyF8 Then
        '定座位
        If cmdSetSeat.Enabled = True Then
            cmdSetSeat_Click
        End If
        
    ElseIf KeyCode = vbKeyCapital And Shift Then
        If lvBus.GridLines = True Then
            lvBus.GridLines = False
        Else
            lvBus.GridLines = True
        End If
    
    ElseIf KeyCode = vbKeyF11 Then
''由于绍兴的保险与温州不同，先暂时去掉
''        '选中需要保险
''        If chkInsurance.Value = vbChecked Then
''            chkInsurance.Value = vbUnchecked
''        Else
''            chkInsurance.Value = vbChecked
''        End If
        
    End If
    
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 45 Then
        
        Exit Sub
    ElseIf KeyAscii = 27 Then
        SetDefaultSellTicket
        txtPrevDate.Text = 0
'        txtTime.Text = 0
        lblmileate.Caption = ""
        txtSeat.Text = ""
        cboEndStation.Text = ""
        cboEndStation.SetFocus
    ElseIf KeyAscii = 13 And (lvSellStation.Enabled) And (Me.ActiveControl Is lvBus) Then
        lvSellStation.SetFocus
        Exit Sub
    ElseIf KeyAscii = 13 And Not (Me.ActiveControl Is cboEndStation) And Not (Me.ActiveControl Is txtReceivedMoney) _
                And Not (Me.ActiveControl Is txtHalfSell) And Not (Me.ActiveControl Is txtPreferentialSell) _
                And Not (Me.ActiveControl Is txtFullSell) Then   '
            SendKeys "{TAB}"
    ElseIf KeyAscii = 43 Then
        txtPrevDate.Text = 0
        cboEndStation.SetFocus
    ElseIf KeyAscii = Asc("*") Then
        '如果输入*,则跳到起点站选择处
        If cboStartStation.Enabled Then
            cboStartStation.SetFocus
        End If
    End If
    
End Sub

'初始化winsock
Private Sub InitSock()
    
    wsClient.Close
    wsClient.RemoteHost = m_szRemoteHost
    wsClient.RemotePort = m_szRemotePort
    wsClient.Connect
    
End Sub
Private Sub Form_Load()
On Error GoTo Here
    '===============================
    '初始化winsock控件
    '===============================
    InitValue
    InitSock
    '===============================
    
    FillStartStation
    
    flblSellDate.Caption = ToStandardDateStr(Date)
'    flblSellDate.Caption = ToStandardDateStr(m_oParam.NowDate)
    txtPrevDate.Text = 0
'    txtTime.Text = 0
    m_dbTotalPrice = 0
'    m_bPrint = False
    RefreshPreferentialTicket '读取优惠票信息
'    GetPreSellBus  '显示预售状态信息
'    RefreshStation2
    
    flblTotalPrice.Caption = FormatMoney(m_dbTotalPrice + TotalInsurace)
    EnableSeatAndStand
    EnableSellButton
    
    m_bPreClear = False
    m_bSetFocus = False
    m_bPreSellFocus = True
    
    If m_bSellStationCanSellEachOther Then
        lvSellStation.Enabled = True
    Else
        lvSellStation.Enabled = False
    End If
    
    '对齐列表头
'    AlignHeadWidth Me.name, lvBus
'    AlignHeadWidth Me.name, lvSellStation

Exit Sub
Here:
    ShowErrorMsg


End Sub

Private Sub Form_Resize()
    If MDISellTicket.ActiveForm Is Me Then
        If Not Me.WindowState = vbMaximized Then Me.WindowState = vbMaximized
        fraOutLine.Left = (Me.ScaleWidth - fraOutLine.Width) / 2
        fraOutLine.Top = (Me.ScaleHeight - fraOutLine.Height) / 2
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Here
    
    
'    SaveHeadWidth Me.name, lvBus
'    SaveHeadWidth Me.name, lvSellStation
    m_clSell.Remove GetEncodedKey(Me.Tag)
'    MDISellTicket.lblSell.Value = vbUnchecked
    MDISellTicket.abMenuTool.Bands("mnuFunction").Tools("mnuSellTkt").Checked = False
    MDISellTicket.EnableSortAndRefresh False
    
    
        '***************
        '语音显示器
        WriteReg cszComPort, CStr(g_lComPort)
    
On Error GoTo 0
Exit Sub
Here:
    ShowErrorMsg
End Sub






Private Sub lvBus_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    SortListView lvBus, ColumnHeader.Index
End Sub

Private Sub lvBus_DblClick()
    If Not lvBus.SelectedItem Is Nothing Then
        Call txtFullSell_KeyPress(vbKeyReturn)
    End If
End Sub

Private Sub lvBus_GotFocus()
    lblBus.ForeColor = clActiveColor
    ShowRightSeatType

'
'    '********************
'    '语音显示器
'        SetUser GetActiveUserID
'        'SetWhere
'    '********************
End Sub

Private Sub lvBus_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo Here
      
        RefreshSellStation m_rsBusInfo
        ShowRightSeatType
        DoThingWhenBusChange
        DealPrice
        
On Error GoTo 0
Exit Sub
Here:
    ShowErrorMsg
End Sub

'当前选中的车次改变了要做的事
'更新全/半票价
'显示此车次到指定车站的限售时间及限售张数
'处理票价
'设置站票及选座按钮的状态
'设置售票按钮的状态
Private Sub DoThingWhenBusChange()
    On Error GoTo Here
    If Not lvBus.SelectedItem Is Nothing Then
        Dim liTemp As ListItem
        Set liTemp = lvBus.SelectedItem
        lblSinglePrice.Caption = FormatMoney(liTemp.SubItems(ID_FullPrice)) & "/" & FormatMoney(liTemp.SubItems(ID_HalfPrice))
        'flblStandCount.Caption = liTemp.subitems(ID_StandCount)
        If liTemp.SubItems(ID_BusType1) = TP_ScrollBus Then
            flblLimitedCount.Caption = ""
            flblLimitedTime.Caption = ""
'            flblStandCount.Caption = ""
        Else
            flblLimitedCount.Caption = GetStationLimitedCountStr(CInt(liTemp.SubItems(ID_LimitedCount)))
            flblLimitedTime.Caption = GetStationLimitedTimeStr(CInt(liTemp.SubItems(ID_LimitedTime)), CDate(flblSellDate.Caption), CDate(liTemp.SubItems(ID_OffTime)))
           ' flblStandCount.Caption = liTemp.subitems(ID_StandCount)
'           flblStandCount.Caption = 0
        End If
    Else
        lblSinglePrice.Caption = FormatMoney(0) & "/" & FormatMoney(0)
        flblLimitedCount.Caption = ""
        flblLimitedTime.Caption = ""
'        flblStandCount.Caption = ""
    End If
'    DealPrice   ' 处理票价
    EnableSeatAndStand  '设置站票及选座按钮的状态
    EnableSellButton    '设置售票按钮状态
    On Error GoTo 0
    Exit Sub
Here:
    ShowErrorMsg
End Sub

Private Sub lvBus_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown And Shift = 0 Then
        If lvBus.SelectedItem Is Nothing Or lvBus.ListItems.count < 1 Then Exit Sub
        If (lvBus.SelectedItem.Index = lvBus.ListItems.count - 2) Or (lvBus.SelectedItem.Index = lvBus.ListItems.count - 1) Or (lvBus.SelectedItem.Index = lvBus.ListItems.count) Then
            RefreshNextScreen
        End If
    End If
    If KeyCode = vbKeyPageDown And Shift = 0 Then
        RefreshNextScreen
    End If
    If KeyCode = vbKeyEnd And Shift = 0 Then
        RefreshAllScreen
    End If
End Sub

Private Sub lvBus_LostFocus()
    lblBus.ForeColor = 0
    DispStation
    RefreshCheckGate
End Sub


Private Sub mnu_changeseattype_Click()
    If cboSeatType.ListIndex = cboSeatType.ListCount - 1 Then
        cboSeatType.ListIndex = 0
    Else
        cboSeatType.ListIndex = cboSeatType.ListIndex + 1
    End If
End Sub

Private Sub lvSellStation_GotFocus()

   lblsellstation.ForeColor = clActiveColor
   cboSeatType_Change
   If lvSellStation.ListItems.count > 0 Then
        flblTotalPrice.Caption = FormatMoney(lvSellStation.SelectedItem.SubItems(2) + TotalInsurace)
        lvSellStation.Tag = lvSellStation.SelectedItem.Text
    '    lvSellStation.ListItems(m_nSellCount).Tag = lvSellStation.SelectedItem.Text
   End If
   DealPrice
End Sub

Private Sub lvSellStation_ItemClick(ByVal Item As MSComctlLib.ListItem)
   cboSeatType_Change
   flblTotalPrice.Caption = FormatMoney(lvSellStation.SelectedItem.SubItems(2) + TotalInsurace)
   DealPrice
End Sub

Private Sub lvSellStation_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Here
   If KeyCode = 13 Then
     txtFullSell.SetFocus
   End If
Exit Sub
Here:
    ShowErrorMsg
  
End Sub

Private Sub lvSellStation_LostFocus()
    lblsellstation.ForeColor = 0
End Sub

Private Sub Timer1_Timer()
'    RefreshBusSeats True
    On Error GoTo Here
    '隔40秒取一下服务器时间
    If m_nCount Mod 20 = 0 Then
'        Date = m_oParam.NowDate
'        Time = m_oParam.NowDateTime
        m_nCount = 0
    End If
    m_nCount = m_nCount + 1
    Exit Sub
    On Error GoTo 0
Here:
    ShowMsg err.Description
End Sub




Private Sub Timer2_Timer()

End Sub

Private Sub tmrConnected_Timer()
    '说明连接成功
    
    tmrConnected.Enabled = False
    '填充站点
    RefreshStation2
    
End Sub



Private Sub txtFullSell_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        txtHalfSell.SetFocus
    End If
End Sub

Private Sub txtFullSell_KeyPress(KeyAscii As Integer)
On Error GoTo Here
    If KeyAscii = 13 Then
        txtFullSell.Text = CInt(txtFullSell.Text)
        If txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text <> 0 And Not (lvBus.SelectedItem Is Nothing) Then
            If lvBus.SelectedItem.SubItems(ID_OffTime) <> cszScrollBus Then
'                If txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text > lvBus.SelectedItem.SubItems(ID_SeatCount) Then
'                    MsgBox "该车次座位已不够！" & vbCrLf & "座位只剩 " & lvBus.SelectedItem.SubItems(ID_SeatCount), vbInformation, "售票台"
'                ElseIf cboSeatType.ListIndex = 0 And txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text > lvBus.SelectedItem.SubItems(ID_SeatTypeCount) Then
'                    MsgBox "该车次[普通]座位已不够！" & vbCrLf & "  [普通]座位只剩 " & lvBus.SelectedItem.SubItems(ID_SeatTypeCount), vbInformation, "售票台"
'                ElseIf cboSeatType.ListIndex = 1 And txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text > lvBus.SelectedItem.SubItems(ID_BedTypeCount) Then
'                    MsgBox "该车次[卧铺]座位已不够！" & vbCrLf & "  [卧铺]座位只剩 " & lvBus.SelectedItem.SubItems(ID_BedTypeCount), vbInformation, "售票台"
'                ElseIf cboSeatType.ListIndex = 2 And txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text > lvBus.SelectedItem.SubItems(ID_AdditionalCount) Then
'                    MsgBox "该车次[加座]座位已不够！" & vbCrLf & "  [加座]座位只剩 " & lvBus.SelectedItem.SubItems(ID_AdditionalCount), vbInformation, "售票台"
'                    txtFullSell.SetFocus
'                    KeyAscii = 0
'                Else
                    cmdPreSell_Click
'                    txtReceivedMoney.SetFocus
                    cmdSell.SetFocus
'                End If
            Else
                cmdPreSell_Click
'                txtReceivedMoney.SetFocus
                cmdSell.SetFocus
            End If
        End If
    End If
    
On Error GoTo 0
Exit Sub
Here:
   ShowErrorMsg
End Sub


Private Sub txtHalfSell_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            If cboPreferentialTicket.ListCount > 0 Then cboPreferentialTicket.SetFocus
        Case vbKeyUp
            txtFullSell.SetFocus
    End Select
    
End Sub

Private Sub txtHalfSell_KeyPress(KeyAscii As Integer)
On Error GoTo Here
    If KeyAscii = 13 Then
        txtFullSell.Text = CInt(txtFullSell.Text)
        If txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text <> 0 And Not (lvBus.SelectedItem Is Nothing) Then
            If lvBus.SelectedItem.SubItems(ID_OffTime) <> cszScrollBus Then
'                If txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text > lvBus.SelectedItem.SubItems(ID_SeatCount) Then
'                    MsgBox "该车次座位已不够！" & vbCrLf & "座位只剩 " & lvBus.SelectedItem.SubItems(ID_SeatCount), vbInformation, "售票台"
'                 ElseIf cboSeatType.ListIndex = 0 And txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text > lvBus.SelectedItem.SubItems(ID_SeatTypeCount) Then
'                    MsgBox "该车次[普通]座位已不够！" & vbCrLf & "  [普通]座位只剩 " & lvBus.SelectedItem.SubItems(ID_SeatTypeCount), vbInformation, "售票台"
'                 ElseIf cboSeatType.ListIndex = 1 And txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text > lvBus.SelectedItem.SubItems(ID_BedTypeCount) Then
'                    MsgBox "该车次[卧铺]座位已不够！" & vbCrLf & "  [卧铺]座位只剩 " & lvBus.SelectedItem.SubItems(ID_BedTypeCount), vbInformation, "售票台"
'                 ElseIf cboSeatType.ListIndex = 2 And txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text > lvBus.SelectedItem.SubItems(ID_AdditionalCount) Then
'                    MsgBox "该车次[加座]座位已不够！" & vbCrLf & "  [加座]座位只剩 " & lvBus.SelectedItem.SubItems(ID_AdditionalCount), vbInformation, "售票台"
'                    txtHalfSell.SetFocus
'                    KeyAscii = 0
'                Else
                    cmdPreSell_Click
'                    txtReceivedMoney.SetFocus
                    cmdSell.SetFocus
'                End If
            Else
                cmdPreSell_Click
                cmdSell.SetFocus
'                txtReceivedMoney.SetFocus
            End If
        End If
    End If
    
On Error GoTo 0
Exit Sub
Here:
    ShowErrorMsg
End Sub

Private Sub txtPreferentialSell_GotFocus()
    txtPreferentialSell.SelStart = 0
    txtPreferentialSell.SelLength = 2
End Sub

Private Sub txtPreferentialSell_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyLeft
            
            cboPreferentialTicket.SetFocus
        Case vbKeyUp
            txtHalfSell.SetFocus
    End Select
    
End Sub

Private Sub txtPreferentialSell_KeyPress(KeyAscii As Integer)
On Error GoTo Here
    If KeyAscii = 13 Then
        If txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text <> 0 And Not (lvBus.SelectedItem Is Nothing) Then
            If lvBus.SelectedItem.SubItems(ID_OffTime) <> cszScrollBus Then
'                If txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text > lvBus.SelectedItem.SubItems(ID_SeatCount) Then
'                    MsgBox "该车次座位已不够！" & vbCrLf & "座位只剩 " & lvBus.SelectedItem.SubItems(ID_SeatCount), vbInformation, "售票台"
'                 ElseIf cboSeatType.ListIndex = 0 And txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text > lvBus.SelectedItem.SubItems(ID_SeatTypeCount) Then
'                    MsgBox "该车次[普通]座位已不够！" & vbCrLf & "  [普通]座位只剩 " & lvBus.SelectedItem.SubItems(ID_SeatTypeCount), vbInformation, "售票台"
'                 ElseIf cboSeatType.ListIndex = 1 And txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text > lvBus.SelectedItem.SubItems(ID_BedTypeCount) Then
'                    MsgBox "该车次[卧铺]座位已不够！" & vbCrLf & "  [卧铺]座位只剩 " & lvBus.SelectedItem.SubItems(ID_BedTypeCount), vbInformation, "售票台"
'                 ElseIf cboSeatType.ListIndex = 2 And txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text > lvBus.SelectedItem.SubItems(ID_AdditionalCount) Then
'                    MsgBox "该车次[加座]座位已不够！" & vbCrLf & "  [加座]座位只剩 " & lvBus.SelectedItem.SubItems(ID_AdditionalCount), vbInformation, "售票台"
'                    txtPreferentialSell.SetFocus
'                    KeyAscii = 0
'                Else
                    cmdPreSell_Click
                   'cmdPreSell.SetFocus
    
                      'txtReceivedMoney.SetFocus
                    cmdSell.SetFocus
'                End If
            Else
                cmdPreSell_Click
                cmdSell.SetFocus
'                txtReceivedMoney.SetFocus
            End If
        End If
        
    End If
    
On Error GoTo 0
Exit Sub
Here:
    ShowErrorMsg
End Sub

Private Sub txtPreferentialSell_LostFocus()
'    fraPreferentialTicket.ForeColor = 0
    DispStationAndNum
End Sub

Private Sub txtFullSell_GotFocus()
    lblFullSell.ForeColor = clActiveColor
'    txtFullSell.SelectOnEntry = True
    txtFullSell.SelStart = 0
    txtFullSell.SelLength = 2
'    txtFullSell.SelText = txtFullSell.text
    
End Sub

Private Sub txtFullSell_LostFocus()
    lblFullSell.ForeColor = 0
    DispStationAndNum
End Sub

Private Sub txtHalfSell_GotFocus()
    lblHalfSell.ForeColor = clActiveColor
    txtHalfSell.SelStart = 0
    txtHalfSell.SelLength = 2
End Sub

Private Sub txtHalfSell_LostFocus()
    lblHalfSell.ForeColor = 0
    DispStationAndNum
 End Sub

Private Sub txtPrevDate_Change()
    On Error Resume Next
    
    If Val(txtPrevDate.Text) > m_nCanSellDay Then txtPrevDate.Text = m_nCanSellDay
'    flblSellDate.Caption = ToStandardDateStr(DateAdd("d", txtPrevDate.Text, m_oParam.NowDate))
    flblSellDate.Caption = ToStandardDateStr(DateAdd("d", txtPrevDate.Text, Date))
    
End Sub

'处理票价（计算总票价？算出找钱)
Private Sub DealPrice()
    Dim i As Integer
    Dim sgTemp As Double  '计算总票价值
    Dim sgvalue As Double
    Dim dbSum As Double
    Dim aszSeatNo() As String
    Dim nSeat As Integer
    Dim nLength As Integer
    nLength = 0
    sgTemp = 0
On Error GoTo Here
    dbSum = GetDealTotalPrice()
    If Not lvBus.SelectedItem Is Nothing Then
        Dim liTemp As ListItem
        Set liTemp = lvBus.SelectedItem
        If txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text = 0 Then
            sgTemp = 0
        Else
            Select Case Trim(m_aszSeatType(1, 1)) 'cboSeatType.ListIndex + 1, 1))
                Case cszSeatType
                    For i = 1 To txtFullSell.Text
                        sgTemp = sgTemp + liTemp.SubItems(ID_FullPrice)
                    Next i
                    For i = 1 To txtHalfSell.Text
                        sgTemp = sgTemp + liTemp.SubItems(ID_HalfPrice)
                    Next i
                    For i = 1 To txtPreferentialSell.Text
                        sgTemp = sgTemp + GetPreferentialPrice(False)
                    Next i
                Case cszBedType
                    For i = 1 To txtFullSell.Text
                        sgTemp = sgTemp + liTemp.SubItems(ID_BedFullPrice)
                    Next i
                    For i = 1 To txtHalfSell.Text
                        sgTemp = sgTemp + liTemp.SubItems(ID_BedHalfPrice)
                    Next i
                    For i = 1 To txtPreferentialSell.Text
                        sgTemp = sgTemp + GetPreferentialPrice(False)
                    Next i
                Case cszAdditionalType
                    For i = 1 To txtFullSell.Text
                        sgTemp = sgTemp + liTemp.SubItems(ID_AdditionalFullPrice)
                    Next i
                    For i = 1 To txtHalfSell.Text
                        sgTemp = sgTemp + liTemp.SubItems(ID_AdditionalHalfPrice)
                    Next i
                    For i = 1 To txtPreferentialSell.Text
                        sgTemp = sgTemp + GetPreferentialPrice(False)
                    Next i
            End Select
            
        End If
        lblTotalPrice.Caption = FormatMoney(sgTemp)
        m_dbTotalPrice = FormatMoney(sgTemp + dbSum)
        lblTotalMoney.Caption = FormatMoney(sgTemp + m_sgTotalMoney)
    Else
        lblTotalPrice.Caption = FormatMoney(0)
        m_dbTotalPrice = FormatMoney(0 + dbSum)
        lblTotalMoney.Caption = FormatMoney(0 + m_sgTotalMoney)
    End If
    If txtReceivedMoney.Text = "0" And Not Me.ActiveControl Is txtReceivedMoney Then txtReceivedMoney.Text = ""
    If Left(txtReceivedMoney.Text, 1) = "." Then txtReceivedMoney.Text = "0" & txtReceivedMoney.Text
    If txtReceivedMoney.Text = "" Then
       sgvalue = 0
    Else
       sgvalue = CDbl(txtReceivedMoney.Text)
    End If
    If sgvalue <> 0 Then
       flblRestMoney.Caption = FormatMoney(CDbl(txtReceivedMoney.Text) - CDbl(m_dbTotalPrice))
       cmdSell.Enabled = True
    Else
       flblRestMoney.Caption = ""
    End If
    
    flblTotalPrice.Caption = FormatMoney(m_dbTotalPrice + TotalInsurace)
On Error GoTo 0
Exit Sub
Here:
    ShowErrorMsg
End Sub

Private Sub txtPreferentialSell_Change()
On Error GoTo Here
    
    
    
    EnableSellButton
    EnableSeatAndStand
    DealPrice
    SetPreSellButton
On Error GoTo 0
Exit Sub
Here:
    ShowErrorMsg
End Sub

Private Sub txtFullSell_Change()
On Error GoTo Here
    
    
    EnableSellButton
    EnableSeatAndStand
    DealPrice
    SetPreSellButton
On Error GoTo 0
Exit Sub
Here:
    ShowErrorMsg
End Sub

Private Sub txtHalfSell_Change()
On Error GoTo Here
    EnableSellButton
    EnableSeatAndStand
    DealPrice
    
    SetPreSellButton
On Error GoTo 0
Exit Sub
Here:
    ShowErrorMsg
End Sub


Private Sub txtPrevDate_GotFocus()
    lblPrevDate.ForeColor = clActiveColor
''    '********************
''    '语音显示器
'        SetUser GetActiveUserID
''    '********************
End Sub

Private Sub txtPrevDate_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyUp
            txtPrevDate.Text = Val(txtPrevDate.Text) + 1
        Case vbKeyDown
            If Val(txtPrevDate.Text) > 0 Then
                txtPrevDate.Text = Val(txtPrevDate.Text) - 1
            End If
    End Select
End Sub

Private Sub txtPrevDate_KeyPress(KeyAscii As Integer)
    If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = 8 Then
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtPrevDate_LostFocus()
On Error GoTo Here
    lblPrevDate.ForeColor = 0
    If txtPrevDate.Text = "" Then txtPrevDate.Text = 0
    

    SendBusRequest True
On Error GoTo 0
Exit Sub
Here:
    ShowErrorMsg
End Sub

Private Sub txtReceivedMoney_Change()
    Dim sgvalue As Double
On Error GoTo Here
    If txtReceivedMoney.Text = "0" And Not Me.ActiveControl Is txtReceivedMoney Then txtReceivedMoney.Text = ""
    If Left(txtReceivedMoney.Text, 1) = "." Then txtReceivedMoney.Text = "0" & txtReceivedMoney.Text
    If txtReceivedMoney.Text = "" Then
       sgvalue = 0
    Else
       sgvalue = CDbl(txtReceivedMoney.Text)
    End If
    If sgvalue <> 0 Then
       flblRestMoney.Caption = FormatMoney(CDbl(txtReceivedMoney.Text) - CDbl(flblTotalPrice.Caption))
       cmdSell.Enabled = True
    Else
       flblRestMoney.Caption = ""
    End If
    CalReceiveMoney
    
On Error GoTo 0
Exit Sub
Here:
    ShowErrorMsg
End Sub


'根据当前信息设置售票按钮的状态
Private Sub EnableSellButton()
    Dim szStationID As String
    szStationID = cboEndStation.BoundText
    If szStationID = "" Or lvBus.SelectedItem Is Nothing Then
        cmdSell.Enabled = False
    Else
        cmdSell.Enabled = True
    End If
End Sub

'根据当前的信息设置站票Check按钮和定座按钮的状态
Private Sub EnableSeatAndStand()
    Dim szStationID As String
    szStationID = cboEndStation.BoundText
        If szStationID = "" Or lvBus.SelectedItem Is Nothing Then  '当前无车次
            cmdSetSeat.Enabled = False
'            chkSetSeat.Value = 0
'            chkSetSeat.Enabled = False
        Else
            Dim liTemp As ListItem
            Set liTemp = lvBus.SelectedItem
            
            If liTemp.SubItems(ID_BusType1) = TP_ScrollBus Then  '是流水车次的话定座和站票无意义
                cmdSetSeat.Enabled = False
'                chkSetSeat.Value = 0
'                chkSetSeat.Enabled = False
            Else
                If liTemp.SubItems(ID_SeatCount) > 0 Then '可售座位数大于0
                    If (txtFullSell.Text = 0 And txtHalfSell.Text = 0 And txtPreferentialSell.Text = 0) Then
                        cmdSetSeat.Enabled = False
'                        chkSetSeat.Value = 0
'                        chkSetSeat.Enabled = False
                    Else
                        cmdSetSeat.Enabled = True
'                        chkSetSeat.Enabled = True
                    End If
                    
                Else '无可售座位数（则这时可售站票肯定大于0，不然就不会将此车次查出来）
                    cmdSetSeat.Enabled = True
'                    chkSetSeat.Enabled = True
'                    chkSetSeat.Value = 0
                End If
                
            End If
        End If

End Sub

'设置缺省的售票状态
Private Sub SetDefaultSellTicket()
    txtFullSell.Text = 1 '售全票张数为1
    txtHalfSell.Text = 0 '售半票张数为0
    txtPreferentialSell.Text = 0 '售免票张数为0
'    txtRightSell.Value = 0 '售免票张数为0
    If chkSetSeat.Enabled Then chkSetSeat.Value = 0 '不定座位
'    If chkStandSeat.Enabled Then chkStandSeat.Value = 0 '不售站票
    
    If txtReceivedMoney.Enabled Then  '所收钞票为0
        'txtReceivedMoney.Text = 0
'        DealPrice
    End If
    chkInsurance.Value = vbUnchecked
End Sub

'刷新某车次的上车站信息
'lvSellStation 列 1 上车站 2 发车时间 3 全价 4 上车站代码

'Private Sub RefreshSellStation(BusID As String)
'  Dim i As Integer
'  Dim lvS As ListItem
'  Dim szTemp As String
'  Dim rsTemp As Recordset
'  Dim szStationID As String
'  Dim nBusType As EBusType
'  On Error GoTo err:
'    lvSellStation.Sorted = False
'    lvSellStation.ListItems.Clear
'    lvSellStation.Refresh
'    szTemp = ""
''    lvSellStation.Enabled = True
'    szStationID = RTrim(cboEndStation.BoundText)
'    If szStationID <> "" Then
'            If m_szRegValue = "" Then
'                Set rsTemp = m_oSell.GetBusRs(CDate(flblSellDate.Caption), szStationID, , BusID)
'            Else
'                Set rsTemp = m_oSell.GetBusRsEx(CDate(flblSellDate.Caption), szStationID, m_szRegValue, , BusID)
'            End If
'    End If
'     If rsTemp.RecordCount = 0 Then
'        Exit Sub
'     End If
'    rsTemp.MoveFirst
'    For i = 1 To rsTemp.RecordCount
'        If Trim(rsTemp!sell_station_id) <> szTemp Then
'        szTemp = Trim(rsTemp!sell_station_id)
'        Set lvS = lvSellStation.ListItems.Add(, , Trim(rsTemp!sell_station_name))
'        nBusType = rsTemp!bus_type
'        If nBusType <> TP_ScrollBus Then
'           lvS.SubItems(1) = Trim(Format(rsTemp!busstarttime, "hh:mm"))
'        Else
'           lvS.SubItems(1) = cszScrollBus
'        End If
'        lvS.SubItems(3) = Trim(rsTemp!sell_station_id)
'        lvS.SubItems(2) = Trim(rsTemp!full_price)
'        End If
'        rsTemp.MoveNext
'    Next i
'
''     If lvSellStation.Enabled = True Then
''        lvSellStation.SetFocus
''        lvSellStation.ListItems(1).Selected = True
''     End If
'
'    Exit Sub
'err:
'   MsgBox err.Description
'End Sub

Private Sub RefreshSellStation(rsTemp As Recordset)
  Dim i As Integer
  Dim lvS As ListItem
  Dim szTemp As String
  Dim szStationID As String
  Dim nBusType As EBusType
  On Error GoTo err:
    lvSellStation.Sorted = False
    lvSellStation.ListItems.Clear
    lvSellStation.Refresh
    szTemp = ""
'    lvSellStation.Enabled = True
    If lvBus.ListItems.count = 0 Then
       Exit Sub
    End If
    rsTemp.MoveFirst
    For i = 1 To rsTemp.RecordCount
     If Trim(rsTemp!bus_id) = Trim(lvBus.SelectedItem.Text) Then
      If Trim(rsTemp!sell_station_id) <> szTemp Then
        szTemp = Trim(rsTemp!sell_station_id)
        Set lvS = lvSellStation.ListItems.Add(, , Trim(rsTemp!sell_station_name))
        nBusType = rsTemp!bus_type
        If nBusType <> TP_ScrollBus Then
           lvS.SubItems(1) = Trim(Format(rsTemp!busstarttime, "hh:mm"))
        Else
           lvS.SubItems(1) = cszScrollBus
        End If
        lvS.SubItems(3) = Trim(rsTemp!sell_station_id)
        lvS.SubItems(2) = Trim(rsTemp!full_price)
        lvS.SubItems(4) = Trim(rsTemp!sell_check_gate_id)
      End If
     End If
        rsTemp.MoveNext
    Next i

    Exit Sub
err:
   MsgBox err.Description
End Sub

'座位类型改变时,刷新相应的票价
Private Sub RefreshBusStation(rsTemp As Recordset, SellStationID As String, SeatTypeID As String)
  Dim i As Integer
  Dim szStationID As String
  On Error GoTo err:
     If rsTemp.RecordCount = 0 Then
        Exit Sub
     End If
    rsTemp.MoveFirst
    For i = 1 To rsTemp.RecordCount
     If Trim(rsTemp!bus_id) = Trim(lvBus.SelectedItem.Text) Then
        If Trim(rsTemp!sell_station_id) = Trim(SellStationID) Then
           If Trim("0" + SeatTypeID) = Trim(rsTemp!seat_type_id) Then
        Select Case Trim(rsTemp!seat_type_id)
        Case cszSeatType
            lvBus.SelectedItem.SubItems(ID_FullPrice) = rsTemp!full_price
            lvBus.SelectedItem.SubItems(ID_HalfPrice) = rsTemp!half_price
            lvBus.SelectedItem.SubItems(ID_PreferentialPrice1) = rsTemp!preferential_ticket1
            lvBus.SelectedItem.SubItems(ID_PreferentialPrice2) = rsTemp!preferential_ticket2
            lvBus.SelectedItem.SubItems(ID_PreferentialPrice3) = rsTemp!preferential_ticket3
        Case cszBedType
            lvBus.SelectedItem.SubItems(ID_BedFullPrice) = rsTemp!full_price
            lvBus.SelectedItem.SubItems(ID_BedHalfPrice) = rsTemp!half_price
            lvBus.SelectedItem.SubItems(ID_BedPreferentialPrice1) = rsTemp!preferential_ticket1
            lvBus.SelectedItem.SubItems(ID_BedPreferentialPrice2) = rsTemp!preferential_ticket2
            lvBus.SelectedItem.SubItems(ID_BedPreferentialPrice3) = rsTemp!preferential_ticket3
        Case cszAdditionalType
            lvBus.SelectedItem.SubItems(ID_AdditionalFullPrice) = rsTemp!full_price
            lvBus.SelectedItem.SubItems(ID_AdditionalHalfPrice) = rsTemp!half_price
            lvBus.SelectedItem.SubItems(ID_AdditionalPreferential1) = rsTemp!preferential_ticket1
            lvBus.SelectedItem.SubItems(ID_AdditionalPreferential2) = rsTemp!preferential_ticket2
            lvBus.SelectedItem.SubItems(ID_AdditionalPreferential3) = rsTemp!preferential_ticket3
        End Select
        lvBus.SelectedItem.SubItems(ID_CheckGate) = Trim(rsTemp!check_gate_id)
        End If
      End If
     End If
        rsTemp.MoveNext
    Next i
        lvBus.SelectedItem.Tag = MakeDisplayString(lvSellStation.SelectedItem.SubItems(3), lvSellStation.SelectedItem.Text)
        lvBus.SelectedItem.SubItems(ID_OffTime) = lvSellStation.SelectedItem.SubItems(1)
        lvBus.SelectedItem.SubItems(ID_FullPrice) = lvSellStation.SelectedItem.SubItems(2)
        lvBus.SelectedItem.SubItems(ID_CheckGate) = lvSellStation.SelectedItem.SubItems(4)
    Exit Sub
err:
   MsgBox err.Description
End Sub

'Private Sub RefreshBusStation(BusID As String, SellStationID As String, SeatTypeID As String)
'  Dim i As Integer
'  Dim rsTemp As Recordset
'  Dim szStationID As String
'  On Error GoTo err:
'    szStationID = RTrim(cboEndStation.BoundText)
'    If szStationID <> "" Then
'            If m_szRegValue = "" Then
'                Set rsTemp = m_oSell.GetBusRs(CDate(flblSellDate.Caption), szStationID, , BusID)
'            Else
'                Set rsTemp = m_oSell.GetBusRsEx(CDate(flblSellDate.Caption), szStationID, m_szRegValue, , BusID)
'            End If
'    End If
'     If rsTemp.RecordCount = 0 Then
'        Exit Sub
'     End If
'    rsTemp.MoveFirst
'    For i = 1 To rsTemp.RecordCount
'        If Trim(rsTemp!sell_station_id) = Trim(SellStationID) Then
'           If Trim("0" + SeatTypeID) = Trim(rsTemp!seat_type_id) Then
'        Select Case Trim(rsTemp!seat_type_id)
'        Case cszSeatType
'            lvBus.SelectedItem.SubItems(ID_FullPrice) = rsTemp!full_price
'            lvBus.SelectedItem.SubItems(ID_HalfPrice) = rsTemp!half_price
'            lvBus.SelectedItem.SubItems(ID_PreferentialPrice1) = rsTemp!preferential_ticket1
'            lvBus.SelectedItem.SubItems(ID_PreferentialPrice2) = rsTemp!preferential_ticket2
'            lvBus.SelectedItem.SubItems(ID_PreferentialPrice3) = rsTemp!preferential_ticket3
'        Case cszBedType
'            lvBus.SelectedItem.SubItems(ID_BedFullPrice) = rsTemp!full_price
'            lvBus.SelectedItem.SubItems(ID_BedHalfPrice) = rsTemp!half_price
'            lvBus.SelectedItem.SubItems(ID_BedPreferentialPrice1) = rsTemp!preferential_ticket1
'            lvBus.SelectedItem.SubItems(ID_BedPreferentialPrice2) = rsTemp!preferential_ticket2
'            lvBus.SelectedItem.SubItems(ID_BedPreferentialPrice3) = rsTemp!preferential_ticket3
'        Case cszAdditionalType
'            lvBus.SelectedItem.SubItems(ID_AdditionalFullPrice) = rsTemp!full_price
'            lvBus.SelectedItem.SubItems(ID_AdditionalHalfPrice) = rsTemp!half_price
'            lvBus.SelectedItem.SubItems(ID_AdditionalPreferential1) = rsTemp!preferential_ticket1
'            lvBus.SelectedItem.SubItems(ID_AdditionalPreferential2) = rsTemp!preferential_ticket2
'            lvBus.SelectedItem.SubItems(ID_AdditionalPreferential3) = rsTemp!preferential_ticket3
'        End Select
'        lvBus.SelectedItem.SubItems(ID_CheckGate) = Trim(rsTemp!check_gate_id)
'        End If
'      End If
'        rsTemp.MoveNext
'    Next i
'        lvBus.SelectedItem.Tag = lvSellStation.SelectedItem.SubItems(3)
'        lvBus.SelectedItem.SubItems(ID_OffTime) = lvSellStation.SelectedItem.SubItems(1)
'        lvBus.SelectedItem.SubItems(ID_FullPrice) = lvSellStation.SelectedItem.SubItems(2)
'    Exit Sub
'err:
'   MsgBox err.Description
'End Sub

'根据指定的日期（当前日期加预售天数）和到站代码刷新车次信息
'pbForce表示是否强行刷新（不管预售天数和到站是否改变）


'得到此次售票的相应序号的座号
Private Function SelfGetSeatNo(pnIndex As Integer) As String
    If chkSetSeat.Enabled = False Then '如果站票选中,则为站票
        SelfGetSeatNo = "ST"
    ElseIf chkSetSeat.Enabled And txtSeat.Text <> "" Then  '如果定座选中,则得到相应的座号
        SelfGetSeatNo = GetSeatNo(txtSeat.Text, pnIndex)
    Else '否则为自动座位号
        SelfGetSeatNo = ""
    End If
End Function

Private Function SelfGetSeatNo12(SetSeatEnable As Boolean, SetSeatValue As Integer, pnIndex As Integer, pszSeatNo As String) As String
    If SetSeatEnable = False Then '如果站票选中,则为站票
        SelfGetSeatNo12 = "ST"
    ElseIf SetSeatEnable And txtSeat.Text <> "" Then  '如果定座选中,则得到相应的座号
        SelfGetSeatNo12 = GetSeatNo(pszSeatNo, pnIndex)
    Else '否则为自动座位号
        SelfGetSeatNo12 = ""
    End If
End Function


Private Sub txtReceivedMoney_GotFocus()
    lblReceivedMoney.ForeColor = clActiveColor
    
    '********************
    '语音显示器
    SetPay flblTotalPrice.Caption
    '********************
End Sub


Private Sub txtReceivedMoney_KeyDown(KeyCode As Integer, Shift As Integer)
Dim nCount As Integer

For nCount = 1 To Len(txtReceivedMoney.Text)
    If Mid(txtReceivedMoney.Text, nCount, 1) = "." Then
       m_blPointCount = True
       Exit For
    End If
Next nCount

End Sub

Private Sub txtReceivedMoney_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 And KeyAscii <> 8 Then
        If KeyAscii < 48 Or KeyAscii > 57 Then
            If KeyAscii = 43 Then
                cboEndStation.SetFocus
            End If
            If m_blPointCount = True And KeyAscii = 46 Then
                KeyAscii = 0
            ElseIf KeyAscii <> 46 Then
                KeyAscii = 0
            End If
        End If
    End If
    m_blPointCount = False
    If KeyAscii = 13 Then
        m_bSetFocus = True
'        cmdSell_Click
        cboEndStation.SetFocus
    End If
End Sub

Private Sub txtReceivedMoney_LostFocus()
'Dim nResult As VbMsgBoxResult
'If IsNumeric(txtReceivedMoney.Text) Then
'    If txtReceivedMoney.Text = 0 Then
'        nResult = MsgBox("是否与下一张票进行票款累加", vbInformation + vbYesNo, Me.Caption)
'        If nResult = vbYes Then
'            cmdSell.Enabled = True
'            cmdSell_Click
'        End If
'    'if
'    End If
'End If
lblReceivedMoney.ForeColor = 0

    '********************
    '语音显示器
        If Val(txtReceivedMoney.Text) <> 0 Then
            SetReceive txtReceivedMoney.Text
            If lblTotalPrice.Caption = txtReceivedMoney.Text Then
                SetThanks
            Else
                SetReturn Val(flblRestMoney.Caption)
            End If
        End If
    '********************
    'Timer1.Enabled = True
        '********************
    '语音显示器
''        SetPay flblTotalPrice.Caption
''        If txtReceivedMoney.Text <> 0 Then
''            SetReceive txtReceivedMoney.Text
''            SetReturn flblRestMoney.Caption
''        End If
'        SetCheck
'        SetThanks
    '********************
  
End Sub



'显示HTMLHELP,直接拷贝
Private Sub DisplayHelp(Optional HelpType As EHelpType = content)
    Dim lActiveControl As Long
    
    Select Case HelpType
        Case content
            lActiveControl = Me.ActiveControl.HelpContextID
            If lActiveControl = 0 Then
                TopicID = Me.HelpContextID
                CallHTMLShowTopicID
            Else
                TopicID = lActiveControl
                CallHTMLShowTopicID
            End If
        Case Index
            CallHTMLHelpIndex
        Case Support
            TopicID = clSupportID
            CallHTMLShowTopicID
    End Select
End Sub

Private Sub RefreshCheckGate()

    On Error GoTo Here

    '发送请求检票口的报文
    
    If wsClient.State = sckConnected Then
        m_szSend = GetCheckGateRequestStr(ResolveDisplay(cboStartStation.Text))
        wsClient.SendData m_szSend
    Else
        InitSock
    End If
    Exit Sub
Here:
    ShowErrorMsg
End Sub

Private Sub RefreshStation2()
    On Error GoTo Here
    
'    Dim szSend As String
    '发送请求站点的报文
    
    If wsClient.State = sckConnected Then
'        SetStartStationID ResolveDisplay(cboStartStation.Text)
        
        m_szSend = GetStationRequestStr(ResolveDisplay(cboStartStation.Text))
        wsClient.SendData m_szSend
    Else
        InitSock
'
'        m_szSend = GetStationRequestStr
'        wsClient.SendData m_szSend
    End If

'
'
'    Dim rsTemp As Recordset
'    Dim sztemp As String
''    szTemp = m_oSell.SellUnitCode
''    m_oSell.SellUnitCode = m_szCurrentUnitID
'    Set rsTemp = GetAllStationRs()
''    m_oSell.SellUnitCode = szTemp
'
'    With cboEndStation
'        Set .RowSource = rsTemp
'        'station_id:到站代码
'        'station_input_code:车站输入码
'        'station_name:车次名称
'
'        .BoundField = "station_id"
'        .ListFields = "station_input_code:4,station_name:4"
'        .AppendWithFields "station_id:9,station_name"
'    End With
'
'    '因为站点已变，所以当前显示的车次信息无效，将其清空
'    lvBus.ListItems.Clear
'
'    '调用车次改变要进行相应操作的方法
'    DoThingWhenBusChange
'    DealPrice
'
'    Set rsTemp = Nothing
'    On Error GoTo 0
    Exit Sub
Here:
'    Set rsTemp = Nothing
    ShowErrorMsg
End Sub


'读取优惠票信息
Private Sub RefreshPreferentialTicket()
    Dim atTicketType() As TTicketType
    Dim aszSeatType() As String
    Dim szHeadText As String
    Dim sgWidth As Single
    Dim nCount As Integer
    Dim i As Integer, j As Integer
    Dim nLen As Integer
    Dim nUsedPerential As Integer
    Dim szTemp As String
    On Error GoTo Here
    
'    szTemp = m_oSell.SellUnitCode
'    m_oSell.SellUnitCode = m_szCurrentUnitID

    '得到所有的票种
    atTicketType = GetAllTicketType
    aszSeatType = GetAllSeatType
'    m_oSell.SellUnitCode = szTemp
    
    nCount = GetTicketTypeCount
    nLen = GetSeatTypeCount
    
    
    sgWidth = 690
    lvBus.ColumnHeaders.Clear
    '添加ListView列头
    With lvBus.ColumnHeaders
        .Add , , "车次", 950 '"BusID"
        .Add , , "车次类型", 0 '"BusType"
        .Add , , "时间", 850 '"OffTime"
        .Add , , "线路名称", 0
        .Add , , "终到站", 1500 '"EndStation"
        .Add , , "总", 700
        .Add , , "订", 440
        .Add , , "座位", 440 '"SeatCount"
        .Add , , "座", 0
        .Add , , "卧", 0 '440
        .Add , , "加", 0 '440
          .Add , , "车型", 1200 '"BusModel"
        '添加票种,不可用的则宽度设为0
        For i = 1 To nCount     '座位票价
            If atTicketType(i).nTicketTypeID = TP_FullPrice Then
                .Add , , "座全", IIf(atTicketType(i).nTicketTypeValid = TP_TicketTypeValid, sgWidth, 0)
            Else
                .Add , , Trim(atTicketType(i).szTicketTypeName), IIf(atTicketType(i).nTicketTypeValid = TP_TicketTypeValid, sgWidth, 0)
                If atTicketType(i).nTicketTypeID = TP_HalfPrice Then lblHalfSell.Caption = Trim(atTicketType(i).szTicketTypeName) & "(&X)" & ":"
            End If
        Next i
        For i = 1 To nCount   '卧铺票价
            If atTicketType(i).nTicketTypeID = TP_FullPrice Then
                .Add , , "卧全", IIf(atTicketType(i).nTicketTypeValid = TP_TicketTypeValid, 0, 0)
            Else
                .Add , , Trim(atTicketType(i).szTicketTypeName), IIf(atTicketType(i).nTicketTypeValid = TP_TicketTypeValid, 0, 0)
            End If
        Next i
        For i = 1 To nCount  '加座票价
            If atTicketType(i).nTicketTypeID = TP_FullPrice Then
                .Add , , "加全", IIf(atTicketType(i).nTicketTypeValid = TP_TicketTypeValid, 0, 0)
            Else
                .Add , , Trim(atTicketType(i).szTicketTypeName), IIf(atTicketType(i).nTicketTypeValid = TP_TicketTypeValid, 0, 0)
            End If
        Next i
        .Add , , "限售张数", 0 '"LimitedCount"
        .Add , , "限售时间", 0 '"LimitedTime"
        .Add , , "车次类型", 0 '"BusType"
        .Add , , "检票口", 0 '"CheckGate"
        .Add , , "站票", 0 '"StandCount"
    End With
    
    '设置ComboBox和优惠票是否可用
    nUsedPerential = 0
    For i = 1 To nCount
        If atTicketType(i).nTicketTypeID = TP_HalfPrice And atTicketType(i).nTicketTypeValid = TP_TicketTypeValid Then
            txtHalfSell.Enabled = True
            lblHalfSell.Enabled = True
        ElseIf atTicketType(i).nTicketTypeID > TP_HalfPrice And atTicketType(i).nTicketTypeValid = TP_TicketTypeValid Then
            cboPreferentialTicket.AddItem Trim(atTicketType(i).szTicketTypeName)
            nUsedPerential = nUsedPerential + 1
        End If
    Next i
    
    '设置座位类型
    If nLen <> 0 Then
        ReDim m_aszSeatType(1 To nLen, 1 To 3)
        For i = 1 To nLen
            cboSeatType.AddItem aszSeatType(i, 2)
            m_aszSeatType(i, 1) = aszSeatType(i, 1)
            m_aszSeatType(i, 2) = aszSeatType(i, 2)
            m_aszSeatType(i, 3) = aszSeatType(i, 3)
        Next
        If cboSeatType.ListCount > 0 Then
            cboSeatType.ListIndex = 0
        End If
    End If
    If cboPreferentialTicket.ListCount < 1 Then
        cboPreferentialTicket.Enabled = False
        txtPreferentialSell.Enabled = False
        cboPreferentialTicket.Text = ""
    Else
        cboPreferentialTicket.Enabled = True
        txtPreferentialSell.Enabled = True
        cboPreferentialTicket.ListIndex = 0
    End If
    '将组合框中的票种代码与票种名称放到数组m_atTicketType 中
    If nUsedPerential > 0 Then ReDim m_atTicketType(1 To nUsedPerential)
    j = 0
    For i = 1 To nCount
        If atTicketType(i).nTicketTypeID > TP_HalfPrice And atTicketType(i).nTicketTypeValid = TP_TicketTypeValid Then
            j = j + 1
            m_atTicketType(j) = atTicketType(i)
        End If
    Next i
    On Error GoTo 0
   Exit Sub
Here:
    ShowErrorMsg
End Sub

'得到对应的优惠票种的对应的票价
Private Function GetPreferentialPrice(Optional pbIsSell As Boolean = False) As Double
Dim liTemp As ListItem
Dim dbTemp As Double
    Set liTemp = lvBus.SelectedItem
    Select Case Trim(m_aszSeatType(cboSeatType.ListIndex + 1, 1))
        Case cszSeatType
            If cboPreferentialTicket.ListCount > 0 Then
                Select Case m_atTicketType(cboPreferentialTicket.ListIndex + 1).nTicketTypeID
                    Case TP_FreeTicket
                        dbTemp = CDbl(liTemp.SubItems(IIf(pbIsSell, ID_FullPrice, ID_FreePrice)))
                    Case TP_PreferentialTicket1
                        dbTemp = CDbl(liTemp.SubItems(ID_PreferentialPrice1))
                    Case TP_PreferentialTicket2
                        dbTemp = CDbl(liTemp.SubItems(ID_PreferentialPrice2))
                    Case TP_PreferentialTicket3
                        dbTemp = CDbl(liTemp.SubItems(ID_PreferentialPrice3))
                End Select
            End If
        Case cszBedType
            If cboPreferentialTicket.ListCount > 0 Then
                Select Case m_atTicketType(cboPreferentialTicket.ListIndex + 1).nTicketTypeID
                    Case TP_FreeTicket
                        dbTemp = CDbl(liTemp.SubItems(IIf(pbIsSell, ID_BedFullPrice, ID_BedFreePrice)))
                    Case TP_PreferentialTicket1
                        dbTemp = CDbl(liTemp.SubItems(ID_BedPreferentialPrice1))
                    Case TP_PreferentialTicket2
                        dbTemp = CDbl(liTemp.SubItems(ID_BedPreferentialPrice2))
                    Case TP_PreferentialTicket3
                        dbTemp = CDbl(liTemp.SubItems(ID_BedPreferentialPrice3))
                End Select
            End If
        Case cszAdditionalType
            If cboPreferentialTicket.ListCount > 0 Then
                Select Case m_atTicketType(cboPreferentialTicket.ListIndex + 1).nTicketTypeID
                    Case TP_FreeTicket
                        dbTemp = CDbl(liTemp.SubItems(IIf(pbIsSell, ID_AdditionalFullPrice, ID_FreePrice)))
                    Case TP_PreferentialTicket1
                        dbTemp = CDbl(liTemp.SubItems(ID_AdditionalPreferential1))
                    Case TP_PreferentialTicket2
                        dbTemp = CDbl(liTemp.SubItems(ID_AdditionalPreferential2))
                    Case TP_PreferentialTicket3
                        dbTemp = CDbl(liTemp.SubItems(ID_AdditionalPreferential3))
                End Select
            End If
    End Select
            
    GetPreferentialPrice = dbTemp
End Function
Private Function GetDealTotalPrice() As Double  '得到总票价
    Dim iCount As Integer
    Dim dbTotal As Double
    dbTotal = 0
'    If lvPreSell.ListItems.Count <> 0 Then
'        For iCount = 1 To lvPreSell.ListItems.Count
'            dbTotal = dbTotal + Val(lvPreSell.ListItems(iCount).SubItems(IT_SumPrice))
'        Next iCount
'    End If
    GetDealTotalPrice = dbTotal
End Function
'///////////////////////////////////
'设置预售按钮状态
Private Sub SetPreSellButton()
'    If txtFullSell.Text = 0 And txtHalfSell.Text = 0 And txtPreferentialSell.Text = 0 Then
'        cmdPreSell.Enabled = False
'    Else
'        cmdPreSell.Enabled = True
'    End If
End Sub
'/////////////////////////////////
'处理折扣票与定座
Private Sub DealDiscountAndSeat()
   '判断是否有售折扣票权限
   On Error GoTo Here
'   If m_oSell.DiscountIsValid Then
'        txtDiscount.Enabled = False
'        fraDiscountTicket.Enabled = False
'   End If
'   If m_oSell.OrderSeatIsValid Then
        chkSetSeat.Value = 0
        chkSetSeat.Visible = False
        lblSetSeat.Enabled = False
'   End If
   On Error GoTo 0
   Exit Sub
Here:
   ShowErrorMsg
End Sub
'//////////////////////////
'预售用订座
Private Function PreOrderSeat() As String
Dim i As Integer
Dim szTemp As String
Dim liTemp As ListItem
On Error GoTo Here
'If lvPreSell.ListItems.Count <> 0 Then
'    For i = 1 To lvPreSell.ListItems.Count
'        Set liTemp = lvPreSell.ListItems(i)
'        If CDate(flblSellDate.Caption) = CDate(liTemp.SubItems(IT_BusDate)) And lvBus.SelectedItem.Text = liTemp.Text Then
'            If liTemp.SubItems(IT_SeatNo) <> "" Then
'                sztemp = sztemp & "," & liTemp.SubItems(IT_SeatNo)
'            End If
'        End If
'    Next i
'Else
    szTemp = ""
'End If
PreOrderSeat = szTemp
On Error GoTo 0
Exit Function
Here:
    PreOrderSeat = ""
    ShowErrorMsg
End Function
'//////////////////////////////////////
'返回相同车次索引
Private Function GetSameBusIndex() As Integer
Dim i As Integer
Dim liTemp As ListItem
Dim liSelected As ListItem
'If lvPreSell.ListItems.Count <> 0 And (Not lvBus.SelectedItem Is Nothing) Then
'    Set liSelected = lvBus.SelectedItem
'    For i = 1 To lvPreSell.ListItems.Count
'        Set liTemp = lvPreSell.ListItems(i)
'        If liTemp.Text = liSelected.Text And liTemp.SubItems(IT_BusDate) = CDate(flblSellDate.Caption) And liTemp.SubItems(IT_BoundText) = cboEndStation.BoundText Then
'            GetSameBusIndex = i
'            Exit Function
'        End If
'    Next i
'End If
GetSameBusIndex = 0
End Function
'/////////////////////////////////////
'合并相同车次信息
'Private Sub MergeSameBusInfo(nSameIndex As Integer)
'Dim liTemp As ListItem
'Set liTemp = lvPreSell.ListItems(nSameIndex)
'Dim szPrice As String
'Dim szTicketType As String
'Dim i As Integer
'Dim sgTemp As Single
'sgTemp = 0
'With liTemp
'
'    .SubItems(IT_SumTicketNum) = Val(.SubItems(IT_SumTicketNum)) + txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text
'    .SubItems(IT_FullNum) = Val(.SubItems(IT_FullNum)) + txtFullSell.Text
'    .SubItems(IT_HalfNum) = Val(.SubItems(IT_HalfNum)) + txtHalfSell.Text
'    .SubItems(IT_SeatNo) = Trim(.SubItems(IT_SeatNo)) & "," & Trim(txtSeat.Text)
'    If cboPreferentialTicket.ListCount > 0 Then
'        Select Case m_atTicketType(cboPreferentialTicket.ListIndex + 1).nTicketTypeID
'        Case TP_FreeTicket
'            .SubItems(IT_FreeNum) = txtPreferentialSell.Text + Val(.SubItems(IT_FreeNum))
'        Case TP_PreferentialTicket1
'            .SubItems(IT_PreferentialNum1) = txtPreferentialSell.Text + Val(.SubItems(IT_PreferentialNum1))
'        Case TP_PreferentialTicket2
'            .SubItems(IT_PreferentialNum2) = txtPreferentialSell.Text + Val(.SubItems(IT_PreferentialNum2))
'        Case TP_PreferentialTicket3
'            .SubItems(IT_PreferentialNum3) = txtPreferentialSell.Text + Val(.SubItems(IT_PreferentialNum3))
'        End Select
'    End If
'
'    .SubItems(IT_SumPrice) = Val(.SubItems(IT_SumPrice)) + _
'                                txtFullSell.Text * lvBus.SelectedItem.SubItems(ID_FullPrice) + _
'                                txtHalfSell.Text * lvBus.SelectedItem.SubItems(ID_HalfPrice) + _
'                                txtPreferentialSell.Text * GetPreferentialPrice
'
'
'End With
'End Sub


''判断是否有滚动车次
'Private Function IsHaveScrollBus() As Boolean
''    Dim i As Integer
''    If lvBus.ListItems.count > 0 Then
''        For i = 1 To lvBus.ListItems.count
''            If lvBus.ListItems(i).SubItems(ID_OffTime) = cszScrollBus Then
''                IsHaveScrollBus = True
''                Exit Function
''            End If
''        Next i
''    End If
''    IsHaveScrollBus = False
'End Function
'''初始化站点车次顺序数组
'Private Sub InitScrollBusOrder()
'    Dim i As Integer
'    Dim nCurLen As Integer
'
'    nCurLen = ArrayLength(m_atbBusOrder)
'    If nCurLen = 0 Then
'        ReDim m_atbBusOrder(1 To 1)
'    Else
'        ReDim Preserve m_atbBusOrder(1 To nCurLen + 1)
'    End If
'
'    m_atbBusOrder(nCurLen + 1).szStatioinID = Trim(cboEndStation.BoundText)
'    m_atbBusOrder(nCurLen + 1).dbCount = 1
'End Sub
'判断滚动站点是否存在于数组当中
'Private Function IsExitInTeam(pszStationID As String) As Integer
'    Dim i As Integer
'    Dim nLen As Integer
'    nLen = ArrayLength(m_atbBusOrder)
'    If nLen > 300 Then
'        ReDim m_atbBusOrder(1 To 1)
'        m_atbBusOrder(1).szStatioinID = ""
'        m_atbBusOrder(1).dbCount = 1
'        IsExitInTeam = 0
'        Exit Function
'    End If
'    For i = 1 To nLen
'        If pszStationID = m_atbBusOrder(i).szStatioinID Then
'            IsExitInTeam = i
'            Exit Function
'        End If
'    Next i
'    IsExitInTeam = 0
'End Function
'给数组顺序最小索引加值
Private Sub AddValueToIndex(pnIndex As Integer)
    If m_atbBusOrder(pnIndex).dbCount > 1000 Then
        m_atbBusOrder(pnIndex).dbCount = 1
    Else
        m_atbBusOrder(pnIndex).dbCount = m_atbBusOrder(pnIndex).dbCount + 1
    End If
End Sub
''lvBus中显示正确的车次顺序
'Private Sub SetCorrectBusOrder(pszStationID As String)
'    Dim nIndex As Integer
'    Dim dbTemp As Double
'    Dim aszSaveTemp() As String
'    Dim j As Integer
'    Dim liTemp As ListItem
'    Dim nCount As Integer
'    nIndex = IsExitInTeam(pszStationID)
'    If lvBus.ListItems.count <> 0 Then
'        nCount = (m_atbBusOrder(nIndex).dbCount Mod lvBus.ListItems.count) + 1
'
'        ReDim aszSaveTemp(1 To lvBus.ListItems(nCount).ListSubItems.count)
'        aszSaveTemp(1) = lvBus.ListItems(nCount)
'        For j = 2 To lvBus.ListItems(nCount).ListSubItems.count
'            aszSaveTemp(j) = lvBus.ListItems(nCount).SubItems(j - 1)
'        Next j
'        lvBus.ListItems.Remove nCount
'
'        Set liTemp = lvBus.ListItems.Add(1, , aszSaveTemp(1))
'        For j = 1 To ArrayLength(aszSaveTemp) - 1
'            liTemp.ListSubItems.Add , , aszSaveTemp(j + 1)
'
'        Next j
'        liTemp.Selected = True
'    End If
'   If lvBus.ListItems.count > 0 Then
'     RefreshSellStation m_rsBusInfo
'   Else
'     lvSellStation.ListItems.Clear
'   End If
'    Set liTemp = Nothing
'End Sub

Private Function GetSeatTypeName(pszSeatTypeID As String) As String
    Dim i As Integer
    Dim nLen As Integer
    Dim szTemp As String
    nLen = ArrayLength(m_aszSeatType)
    For i = 1 To nLen
        If m_aszSeatType(i, 1) = pszSeatTypeID Then
            szTemp = Space(1) & Trim(m_aszSeatType(i, 2))
            Exit For
        End If
    Next
    GetSeatTypeName = szTemp
End Function

Private Sub ShowRightSeatType()
    Dim liTemp As ListItem
    If cboSeatType.ListCount = 0 Then Exit Sub
    If Not lvBus.SelectedItem Is Nothing And Me.ActiveControl Is lvBus Then
        Set liTemp = lvBus.SelectedItem
'        If liTemp.SubItems(ID_FullPrice) <> 0 Then
            cboSeatType.ListIndex = 0
'        Else
'            If liTemp.SubItems(ID_BedFullPrice) <> 0 Then
'                cboSeatType.ListIndex = 1
'            Else
'                cboSeatType.ListIndex = 2
'            End If
'        End If
    End If
End Sub

Private Sub upPreDate_DownClick()
    If Val(txtPrevDate) > 0 Then
        txtPrevDate.Text = Val(txtPrevDate.Text) - 1
    End If
End Sub

Private Sub upPreDate_UpClick()
    txtPrevDate.Text = Val(txtPrevDate.Text) + 1
End Sub

'///////////////////////////////////////
'显示下一屏
Private Sub RefreshNextScreen()
    Dim i As Integer
    Dim j As Integer
    Dim liTemp As ListItem
    Dim lForeColor As OLE_COLOR
    Dim nBusType As EBusType
    j = 0
    If m_rsBusInfo Is Nothing Then Exit Sub
    If Not m_rsBusInfo.EOF Then m_rsBusInfo.MoveNext
    Do While Not m_rsBusInfo.EOF
       j = j + 1
       If j > m_ISellScreenShow Then
         m_rsBusInfo.MovePrevious
         Exit Sub
       End If
       For i = lvBus.ListItems.count To 1 Step -1

            If RTrim(m_rsBusInfo!bus_id) = lvBus.ListItems(i) And Format(m_rsBusInfo!bus_date, "yyyy-mm-dd") = CDate(flblSellDate.Caption) Then
                If liTemp Is Nothing Then Set liTemp = lvBus.ListItems(i)
                Select Case Trim(m_rsBusInfo!seat_type_id)
                    Case cszSeatType
                        liTemp.SubItems(ID_FullPrice) = m_rsBusInfo!full_price
                        liTemp.SubItems(ID_HalfPrice) = m_rsBusInfo!half_price
                        liTemp.SubItems(ID_PreferentialPrice1) = m_rsBusInfo!preferential_ticket1
                        liTemp.SubItems(ID_PreferentialPrice2) = m_rsBusInfo!preferential_ticket2
                        liTemp.SubItems(ID_PreferentialPrice3) = m_rsBusInfo!preferential_ticket3
                    Case cszBedType
                        liTemp.SubItems(ID_BedFullPrice) = m_rsBusInfo!full_price
                        liTemp.SubItems(ID_BedHalfPrice) = m_rsBusInfo!half_price
                        liTemp.SubItems(ID_BedPreferentialPrice1) = m_rsBusInfo!preferential_ticket1
                        liTemp.SubItems(ID_BedPreferentialPrice2) = m_rsBusInfo!preferential_ticket2
                        liTemp.SubItems(ID_BedPreferentialPrice3) = m_rsBusInfo!preferential_ticket3
                    Case cszAdditionalType
                        liTemp.SubItems(ID_AdditionalFullPrice) = m_rsBusInfo!full_price
                        liTemp.SubItems(ID_AdditionalHalfPrice) = m_rsBusInfo!half_price
                        liTemp.SubItems(ID_AdditionalPreferential1) = m_rsBusInfo!preferential_ticket1
                        liTemp.SubItems(ID_AdditionalPreferential2) = m_rsBusInfo!preferential_ticket2
                        liTemp.SubItems(ID_AdditionalPreferential3) = m_rsBusInfo!preferential_ticket3

                End Select
                GoTo nextstep
            End If
            Exit For
        Next i
        If m_rsBusInfo!status = ST_BusStopped Or m_rsBusInfo!status = ST_BusMergeStopped Or m_rsBusInfo!status = ST_BusSlitpStop Then
            lForeColor = m_lStopBusColor

        Else
            lForeColor = m_lNormalBusColor
            Set liTemp = lvBus.ListItems.Add(, "A" & RTrim(m_rsBusInfo!bus_id), RTrim(m_rsBusInfo!bus_id))   '车次代码"A" & RTrim(m_rsBusInfo！bus_id)
        End If

        nBusType = m_rsBusInfo!bus_type


        If lForeColor <> m_lStopBusColor Then
            liTemp.ForeColor = lForeColor
            If nBusType <> TP_ScrollBus Then
                liTemp.SubItems(ID_BusType) = Trim(m_rsBusInfo!bus_type)
                liTemp.SubItems(ID_OffTime) = Format(m_rsBusInfo!busstarttime, "hh:mm")

            Else
                liTemp.SubItems(ID_VehicleModel) = cszScrollBus
                liTemp.SubItems(ID_OffTime) = cszScrollBus

            End If
            liTemp.SubItems(ID_RouteName) = Trim(m_rsBusInfo!route_name)
            liTemp.SubItems(ID_EndStation) = RTrim(m_rsBusInfo!end_station_name)
            liTemp.SubItems(ID_TotalSeat) = m_rsBusInfo!total_seat
            liTemp.SubItems(ID_SeatCount) = m_rsBusInfo!sale_seat_quantity
            liTemp.SubItems(ID_SeatTypeCount) = m_rsBusInfo!seat_remain
            liTemp.SubItems(ID_BedTypeCount) = m_rsBusInfo!bed_remain
            liTemp.SubItems(ID_AdditionalCount) = m_rsBusInfo!additional_remain
            liTemp.SubItems(ID_VehicleModel) = m_rsBusInfo!vehicle_type_name

            Select Case Trim(m_rsBusInfo!seat_type_id)

                Case cszSeatType
                    liTemp.SubItems(ID_FullPrice) = m_rsBusInfo!full_price
                    liTemp.SubItems(ID_HalfPrice) = m_rsBusInfo!half_price
                    liTemp.SubItems(ID_FreePrice) = 0
                    liTemp.SubItems(ID_PreferentialPrice1) = m_rsBusInfo!preferential_ticket1
                    liTemp.SubItems(ID_PreferentialPrice2) = m_rsBusInfo!preferential_ticket2
                    liTemp.SubItems(ID_PreferentialPrice3) = m_rsBusInfo!preferential_ticket3
                    liTemp.SubItems(ID_BedFullPrice) = 0
                    liTemp.SubItems(ID_BedHalfPrice) = 0
                    liTemp.SubItems(ID_BedFreePrice) = 0
                    liTemp.SubItems(ID_BedPreferentialPrice1) = 0
                    liTemp.SubItems(ID_BedPreferentialPrice2) = 0
                    liTemp.SubItems(ID_BedPreferentialPrice3) = 0
                    liTemp.SubItems(ID_AdditionalFullPrice) = 0
                    liTemp.SubItems(ID_AdditionalHalfPrice) = 0
                    liTemp.SubItems(ID_AdditionalFreePrice) = 0
                    liTemp.SubItems(ID_AdditionalPreferential1) = 0
                    liTemp.SubItems(ID_AdditionalPreferential2) = 0
                    liTemp.SubItems(ID_AdditionalPreferential3) = 0
                Case cszBedType
                    liTemp.SubItems(ID_FullPrice) = 0
                    liTemp.SubItems(ID_HalfPrice) = 0
                    liTemp.SubItems(ID_FreePrice) = 0
                    liTemp.SubItems(ID_PreferentialPrice1) = 0
                    liTemp.SubItems(ID_PreferentialPrice2) = 0
                    liTemp.SubItems(ID_PreferentialPrice3) = 0
                    liTemp.SubItems(ID_BedFullPrice) = m_rsBusInfo!full_price
                    liTemp.SubItems(ID_BedHalfPrice) = m_rsBusInfo!half_price
                    liTemp.SubItems(ID_BedFreePrice) = 0
                    liTemp.SubItems(ID_BedPreferentialPrice1) = m_rsBusInfo!preferential_ticket1
                    liTemp.SubItems(ID_BedPreferentialPrice2) = m_rsBusInfo!preferential_ticket2
                    liTemp.SubItems(ID_BedPreferentialPrice3) = m_rsBusInfo!preferential_ticket3
                    liTemp.SubItems(ID_AdditionalFullPrice) = 0
                    liTemp.SubItems(ID_AdditionalHalfPrice) = 0
                    liTemp.SubItems(ID_AdditionalFreePrice) = 0
                    liTemp.SubItems(ID_AdditionalPreferential1) = 0
                    liTemp.SubItems(ID_AdditionalPreferential2) = 0
                    liTemp.SubItems(ID_AdditionalPreferential3) = 0
                Case cszAdditionalType
                    liTemp.SubItems(ID_FullPrice) = 0
                    liTemp.SubItems(ID_HalfPrice) = 0
                    liTemp.SubItems(ID_FreePrice) = 0
                    liTemp.SubItems(ID_PreferentialPrice1) = 0
                    liTemp.SubItems(ID_PreferentialPrice2) = 0
                    liTemp.SubItems(ID_PreferentialPrice3) = 0
                    liTemp.SubItems(ID_BedFullPrice) = 0
                    liTemp.SubItems(ID_BedHalfPrice) = 0
                    liTemp.SubItems(ID_BedFreePrice) = 0
                    liTemp.SubItems(ID_BedPreferentialPrice1) = 0
                    liTemp.SubItems(ID_BedPreferentialPrice2) = 0
                    liTemp.SubItems(ID_BedPreferentialPrice3) = 0
                    liTemp.SubItems(ID_AdditionalFullPrice) = m_rsBusInfo!full_price
                    liTemp.SubItems(ID_AdditionalHalfPrice) = m_rsBusInfo!half_price
                    liTemp.SubItems(ID_AdditionalFreePrice) = 0
                    liTemp.SubItems(ID_AdditionalPreferential1) = m_rsBusInfo!preferential_ticket1
                    liTemp.SubItems(ID_AdditionalPreferential2) = m_rsBusInfo!preferential_ticket2
                    liTemp.SubItems(ID_AdditionalPreferential3) = m_rsBusInfo!preferential_ticket3

            End Select

            '以下几列不显示出来，只是将其存储，以备后用
            liTemp.SubItems(ID_LimitedCount) = m_rsBusInfo!sale_ticket_quantity
            liTemp.SubItems(ID_LimitedTime) = m_rsBusInfo!stop_sale_time
            liTemp.SubItems(ID_BusType1) = nBusType
            liTemp.SubItems(ID_CheckGate) = m_rsBusInfo!check_gate_id
            liTemp.SubItems(ID_StandCount) = m_rsBusInfo!sale_stand_seat_quantity

        End If
nextstep:
        m_rsBusInfo.MoveNext
    Loop
End Sub
'显示下一屏
Private Sub RefreshAllScreen()
    Dim i As Integer
    Dim j As Integer
    Dim liTemp As ListItem
    Dim lForeColor As OLE_COLOR
    Dim nBusType As EBusType

    If m_rsBusInfo Is Nothing Then Exit Sub
    If Not m_rsBusInfo.EOF Then m_rsBusInfo.MoveNext
    Do While Not m_rsBusInfo.EOF

       For i = lvBus.ListItems.count To 1 Step -1

            If RTrim(m_rsBusInfo!bus_id) = lvBus.ListItems(i) And Format(m_rsBusInfo!bus_date, "yyyy-mm-dd") = CDate(flblSellDate.Caption) Then
                If liTemp Is Nothing Then Set liTemp = lvBus.ListItems(i)
                Select Case Trim(m_rsBusInfo!seat_type_id)
                    Case cszSeatType
                        liTemp.SubItems(ID_FullPrice) = m_rsBusInfo!full_price
                        liTemp.SubItems(ID_HalfPrice) = m_rsBusInfo!half_price
                        liTemp.SubItems(ID_PreferentialPrice1) = m_rsBusInfo!preferential_ticket1
                        liTemp.SubItems(ID_PreferentialPrice2) = m_rsBusInfo!preferential_ticket2
                        liTemp.SubItems(ID_PreferentialPrice3) = m_rsBusInfo!preferential_ticket3
                    Case cszBedType
                        liTemp.SubItems(ID_BedFullPrice) = m_rsBusInfo!full_price
                        liTemp.SubItems(ID_BedHalfPrice) = m_rsBusInfo!half_price
                        liTemp.SubItems(ID_BedPreferentialPrice1) = m_rsBusInfo!preferential_ticket1
                        liTemp.SubItems(ID_BedPreferentialPrice2) = m_rsBusInfo!preferential_ticket2
                        liTemp.SubItems(ID_BedPreferentialPrice3) = m_rsBusInfo!preferential_ticket3
                    Case cszAdditionalType
                        liTemp.SubItems(ID_AdditionalFullPrice) = m_rsBusInfo!full_price
                        liTemp.SubItems(ID_AdditionalHalfPrice) = m_rsBusInfo!half_price
                        liTemp.SubItems(ID_AdditionalPreferential1) = m_rsBusInfo!preferential_ticket1
                        liTemp.SubItems(ID_AdditionalPreferential2) = m_rsBusInfo!preferential_ticket2
                        liTemp.SubItems(ID_AdditionalPreferential3) = m_rsBusInfo!preferential_ticket3

                End Select
                GoTo nextstep
            End If
            Exit For
        Next i
        If m_rsBusInfo!status = ST_BusStopped Or m_rsBusInfo!status = ST_BusMergeStopped Or m_rsBusInfo!status = ST_BusSlitpStop Then
            lForeColor = m_lStopBusColor

        Else
            lForeColor = m_lNormalBusColor
            Set liTemp = lvBus.ListItems.Add(, "A" & RTrim(m_rsBusInfo!bus_id), RTrim(m_rsBusInfo!bus_id))   '车次代码"A" & RTrim(m_rsBusInfo！bus_id)
        End If

        nBusType = m_rsBusInfo!bus_type


        If lForeColor <> m_lStopBusColor Then
            liTemp.ForeColor = lForeColor
            If nBusType <> TP_ScrollBus Then
                liTemp.SubItems(ID_BusType) = Trim(m_rsBusInfo!bus_type)
                liTemp.SubItems(ID_OffTime) = Format(m_rsBusInfo!busstarttime, "hh:mm")

            Else
                liTemp.SubItems(ID_VehicleModel) = cszScrollBus
                liTemp.SubItems(ID_OffTime) = cszScrollBus

            End If
            liTemp.SubItems(ID_RouteName) = Trim(m_rsBusInfo!route_name)
            liTemp.SubItems(ID_EndStation) = RTrim(m_rsBusInfo!end_station_name)
            liTemp.SubItems(ID_TotalSeat) = m_rsBusInfo!total_seat
            liTemp.SubItems(ID_SeatCount) = m_rsBusInfo!sale_seat_quantity
            liTemp.SubItems(ID_SeatTypeCount) = m_rsBusInfo!seat_remain
            liTemp.SubItems(ID_BedTypeCount) = m_rsBusInfo!bed_remain
            liTemp.SubItems(ID_AdditionalCount) = m_rsBusInfo!additional_remain
            liTemp.SubItems(ID_VehicleModel) = m_rsBusInfo!vehicle_type_name

            Select Case Trim(m_rsBusInfo!seat_type_id)

                Case cszSeatType
                    liTemp.SubItems(ID_FullPrice) = m_rsBusInfo!full_price
                    liTemp.SubItems(ID_HalfPrice) = m_rsBusInfo!half_price
                    liTemp.SubItems(ID_FreePrice) = 0
                    liTemp.SubItems(ID_PreferentialPrice1) = m_rsBusInfo!preferential_ticket1
                    liTemp.SubItems(ID_PreferentialPrice2) = m_rsBusInfo!preferential_ticket2
                    liTemp.SubItems(ID_PreferentialPrice3) = m_rsBusInfo!preferential_ticket3
                    liTemp.SubItems(ID_BedFullPrice) = 0
                    liTemp.SubItems(ID_BedHalfPrice) = 0
                    liTemp.SubItems(ID_BedFreePrice) = 0
                    liTemp.SubItems(ID_BedPreferentialPrice1) = 0
                    liTemp.SubItems(ID_BedPreferentialPrice2) = 0
                    liTemp.SubItems(ID_BedPreferentialPrice3) = 0
                    liTemp.SubItems(ID_AdditionalFullPrice) = 0
                    liTemp.SubItems(ID_AdditionalHalfPrice) = 0
                    liTemp.SubItems(ID_AdditionalFreePrice) = 0
                    liTemp.SubItems(ID_AdditionalPreferential1) = 0
                    liTemp.SubItems(ID_AdditionalPreferential2) = 0
                    liTemp.SubItems(ID_AdditionalPreferential3) = 0
                Case cszBedType
                    liTemp.SubItems(ID_FullPrice) = 0
                    liTemp.SubItems(ID_HalfPrice) = 0
                    liTemp.SubItems(ID_FreePrice) = 0
                    liTemp.SubItems(ID_PreferentialPrice1) = 0
                    liTemp.SubItems(ID_PreferentialPrice2) = 0
                    liTemp.SubItems(ID_PreferentialPrice3) = 0
                    liTemp.SubItems(ID_BedFullPrice) = m_rsBusInfo!full_price
                    liTemp.SubItems(ID_BedHalfPrice) = m_rsBusInfo!half_price
                    liTemp.SubItems(ID_BedFreePrice) = 0
                    liTemp.SubItems(ID_BedPreferentialPrice1) = m_rsBusInfo!preferential_ticket1
                    liTemp.SubItems(ID_BedPreferentialPrice2) = m_rsBusInfo!preferential_ticket2
                    liTemp.SubItems(ID_BedPreferentialPrice3) = m_rsBusInfo!preferential_ticket3
                    liTemp.SubItems(ID_AdditionalFullPrice) = 0
                    liTemp.SubItems(ID_AdditionalHalfPrice) = 0
                    liTemp.SubItems(ID_AdditionalFreePrice) = 0
                    liTemp.SubItems(ID_AdditionalPreferential1) = 0
                    liTemp.SubItems(ID_AdditionalPreferential2) = 0
                    liTemp.SubItems(ID_AdditionalPreferential3) = 0
                Case cszAdditionalType
                    liTemp.SubItems(ID_FullPrice) = 0
                    liTemp.SubItems(ID_HalfPrice) = 0
                    liTemp.SubItems(ID_FreePrice) = 0
                    liTemp.SubItems(ID_PreferentialPrice1) = 0
                    liTemp.SubItems(ID_PreferentialPrice2) = 0
                    liTemp.SubItems(ID_PreferentialPrice3) = 0
                    liTemp.SubItems(ID_BedFullPrice) = 0
                    liTemp.SubItems(ID_BedHalfPrice) = 0
                    liTemp.SubItems(ID_BedFreePrice) = 0
                    liTemp.SubItems(ID_BedPreferentialPrice1) = 0
                    liTemp.SubItems(ID_BedPreferentialPrice2) = 0
                    liTemp.SubItems(ID_BedPreferentialPrice3) = 0
                    liTemp.SubItems(ID_AdditionalFullPrice) = m_rsBusInfo!full_price
                    liTemp.SubItems(ID_AdditionalHalfPrice) = m_rsBusInfo!half_price
                    liTemp.SubItems(ID_AdditionalFreePrice) = 0
                    liTemp.SubItems(ID_AdditionalPreferential1) = m_rsBusInfo!preferential_ticket1
                    liTemp.SubItems(ID_AdditionalPreferential2) = m_rsBusInfo!preferential_ticket2
                    liTemp.SubItems(ID_AdditionalPreferential3) = m_rsBusInfo!preferential_ticket3

            End Select

            '以下几列不显示出来，只是将其存储，以备后用
            liTemp.SubItems(ID_LimitedCount) = m_rsBusInfo!sale_ticket_quantity
            liTemp.SubItems(ID_LimitedTime) = m_rsBusInfo!stop_sale_time
            liTemp.SubItems(ID_BusType1) = nBusType
            liTemp.SubItems(ID_CheckGate) = m_rsBusInfo!check_gate_id
            liTemp.SubItems(ID_StandCount) = m_rsBusInfo!sale_stand_seat_quantity

        End If
nextstep:
        m_rsBusInfo.MoveNext
    Loop
End Sub


Public Sub ChangeSeatType()
    If cboSeatType.ListIndex = cboSeatType.ListCount - 1 Then
        cboSeatType.ListIndex = 0
    Else
        cboSeatType.ListIndex = cboSeatType.ListIndex + 1
    End If
    cboSeatType_Change
End Sub

Private Function TotalInsurace() As Double
    '汇总保险费
    Dim i As Integer
    Dim nCount As Integer
    If chkInsurance.Value = vbChecked Then
        nCount = 0
'        For i = 1 To lvPreSell.ListItems.Count
'            nCount = nCount + lvPreSell.ListItems(i).SubItems(IT_SumTicketNum)
'        Next i
        nCount = nCount + Val(txtFullSell.Text) + Val(txtHalfSell.Text) + Val(txtPreferentialSell.Text)
        '保险费设为每张2元
        
        TotalInsurace = nCount * 2
    Else
        TotalInsurace = 0
    End If
End Function

Private Sub DispStationAndNum()
    '重计算张数
    Dim m As Long
    Dim i As Integer
    m = 0
'    For i = 1 To lvPreSell.ListItems.Count
'        m = m + lvPreSell.ListItems(i).SubItems(IT_SumTicketNum)
'    Next i
    
        If Not lvBus.SelectedItem Is Nothing Then
            Dim liTemp As ListItem
            Set liTemp = lvBus.SelectedItem
            
'                If lvPreSell.ListItems.Count > 1 Then
'                    '如果预售的站点数大于一个，就不显示
'                    Exit Sub
'                End If
'                If Not lvPreSell.SelectedItem Is Nothing Then
'                    If GetStationNameInCbo(cboEndStation.Text) <> lvPreSell.SelectedItem.SubItems(IT_EndStation) Then
'                        '如果当前的站点与预售中的站点不同，则不显示
'                        Exit Sub
'
'                    End If
'                End If
            '********************
            '语音显示器
            
                'SetStationAndTime liTemp.ListSubItems(ID_EndStation), flblSellDate.Caption & " " & liTemp.ListSubItems(ID_OffTime), txtFreeSell.Value + txtFullSell.Value + txtHalfSell.Value
                SetStationAndTime GetStationNameInCbo(cboEndStation.Text), flblSellDate.Caption & " " & liTemp.ListSubItems(ID_OffTime), m + Val(txtFullSell.Text) + Val(txtHalfSell.Text) + Val(txtPreferentialSell.Text)
            '********************
        End If
End Sub

Private Sub DispStation()
    
    If Not lvBus.SelectedItem Is Nothing Then
        Dim liTemp As ListItem
        Set liTemp = lvBus.SelectedItem
'            If lvPreSell.ListItems.Count > 1 Then
'                '如果预售的站点数大于一个，就不显示
'                Exit Sub
'            End If
'            If Not lvPreSell.SelectedItem Is Nothing Then
'                If GetStationNameInCbo(cboEndStation.Text) <> lvPreSell.SelectedItem.SubItems(IT_EndStation) Then
'                    '如果当前的站点与预售中的站点不同，则不显示
'                    Exit Sub
'
'                End If
'            End If
        '********************
        '语音显示器
'            SetStationAndTime liTemp.ListSubItems(ID_EndStation), flblSellDate.Caption & " " & liTemp.ListSubItems(ID_OffTime)
            SetStationAndTime GetStationNameInCbo(cboEndStation.Text), flblSellDate.Caption & " " & liTemp.ListSubItems(ID_OffTime)
        '********************
    End If
End Sub

'填充起点站
Private Sub FillStartStation()
    Dim aszTemp() As String
    Dim i As Integer
    
    aszTemp = g_aszAllStartStation
    cboStartStation.Clear
    With cboStartStation
        For i = 1 To ArrayLength(aszTemp)
            .AddItem MakeDisplayString(aszTemp(i, 1), aszTemp(i, 3))
        Next i
        If .ListCount > 0 Then
            .ListIndex = 0
        End If
    End With
    
    
End Sub


Private Sub cmdSetSeat_Click()
On Error GoTo Here
'    Dim rsTemp As Recordset
    Dim szTemp As String
    If lvBus.SelectedItem Is Nothing Then
'        Set rsTemp = Nothing
        Exit Sub
    End If
    szTemp = GetSeatRequestStr(ResolveDisplay(cboStartStation.Text), CDate(flblSellDate.Caption), lvBus.SelectedItem.Text)
    wsClient.SendData szTemp
    
    
On Error GoTo 0
Exit Sub
Here:
    ShowErrorMsg
End Sub


Private Sub SendBusRequest(Optional pbForce As Boolean = False)
'    Dim szSend As String
    
    Dim i As Integer
    
    Dim szStationID As String
    
    On Error GoTo Here
    szStationID = RTrim(cboEndStation.BoundText)
    
    Set m_rsBusInfo = Nothing
    
    If cboEndStation.Changed Or pbForce Then
        
        If szStationID <> "" Then
                    
                    
            If wsClient.State = sckConnected Then
                m_szSend = GetBusRequestStr(ResolveDisplay(cboStartStation.Text), CDate(flblSellDate.Caption), szStationID)
                wsClient.SendData m_szSend
            Else
                InitSock
            End If
        Else
            lvBus.ListItems.Clear
            
        End If
    End If

    On Error GoTo 0
    Exit Sub
Here:
    ShowErrorMsg
    Set m_rsBusInfo = Nothing
'    Set liTemp = Nothing
End Sub




Private Sub wsClient_Connect()
    tmrConnected.Enabled = True
End Sub

Private Sub wsClient_DataArrival(ByVal bytesTotal As Long)
    '数据到达处理
    Dim szReceive As String
'    Debug.Print "DataArrival" & "," & bytesTotal
    If bytesTotal >= FIXPKGLEN Then
        wsClient.GetData szReceive
        
        '下面这个条件处理是针对有时包会转成两个来发送的.处理情况
        If Right(szReceive, 1) = cszPackageEnd And Left(szReceive, 1) = cszPackageBegin Then
            '如果数据包以"B"开头,以"@"结尾
'            Debug.Print szReceive, bytesTotal
            RxPkgProcess szReceive
        Else
            
            If Left(szReceive, 1) = cszPackageBegin Then
                '为第一串包
                m_szAllSend = szReceive
            Else
                '为后续包,则进行合并
                m_szAllSend = m_szAllSend & szReceive
            End If
            If Right(szReceive, 1) = cszPackageEnd Then
                
'                Debug.Print m_szAllSend, bytesTotal
                RxPkgProcess m_szAllSend
            End If
        End If
    End If
End Sub

Private Sub wsClient_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox "Error(" & "," & Number & "," & Description & "," & Scode & "," & Source & "," & HelpFile & "," & HelpContext & "," & CancelDisplay

End Sub


Private Sub RxPkgProcess(pszReceive As String)
    '对返回的数据进行处理
    
    Dim szTradeID As String
    Dim szTemp As String
    Dim szStationInfo As String
    Dim szBusInfo As String
    Dim nLen As Integer
    Dim rsTemp As Recordset
    
    
    szTradeID = GetTradeID(pszReceive)
    nLen = GetLen(pszReceive)
    
    Select Case szTradeID
    Case GETSTATIONSID
        '得到站点信息
        If LenA(pszReceive) > FIXPKGLEN + 1 Then
            On Error GoTo jerr
'            szStationInfo = Mid(pszReceive, nLen, Len(pszReceive) - (nLen + 1))
'            Debug.Print pszReceive & vbCr & szStationInfo

        
            Set rsTemp = ConvertStringToStationRS(pszReceive)
            
            If rsTemp.RecordCount = 0 Then Exit Sub
            rsTemp.MoveFirst
            With cboEndStation
                .Clear
                Set .RowSource = rsTemp
                'station_id:到站代码
                'station_input_code:车站输入码
                'station_name:车次名称
                
                .BoundField = "station_id"
                .ListFields = "station_input_code:4,station_name:4"
                .AppendWithFields "station_id:9,station_name"
            End With
            
            '因为站点已变，所以当前显示的车次信息无效，将其清空
            lvBus.ListItems.Clear
            
            '调用车次改变要进行相应操作的方法
            DoThingWhenBusChange
            DealPrice
            
            Set rsTemp = Nothing
            On Error GoTo 0
        End If
    Case GETSCHEDULESID
        '得到车次信息
        
        If Len(pszReceive) > FIXPKGLEN + 1 Then
            On Error GoTo jerr
            

            Set rsTemp = ConvertStringToBusRS(pszReceive)
            If rsTemp Is Nothing Then
                lvBus.ListItems.Clear
                Exit Sub
            Else
                RefreshBus rsTemp
            End If
            DealPrice
            On Error GoTo 0
        End If
    Case BUYTICKETSID
        '得到车票信息
        Debug.Print pszReceive
        On Error GoTo jerr
            If GetRetCode(pszReceive) <> cszCorrectRetCode Then
                MsgBox MidA(pszReceive, cnPosReserved, LenA(pszReceive) - cnPosReserved) & vbCr & "售票不成功", vbOKOnly, "错误" & GetRetCode(pszReceive)
                '清空标签消息.
                lblSellMsg.Caption = ""
                lblSellMsg.Refresh
            Else
                ToPrintTicket pszReceive
            End If
        
        On Error GoTo 0
        
    Case CANCELTICKETSID
'            szTemp = sxBus.CancelTicket(szReceive, m_szSend)
'            wsClient.SendData m_szSend
    Case GETSEATSID
        '得到座位信息
        On Error GoTo jerr
        
        Set rsTemp = ConvertStringToSeatRS(pszReceive)
        Set frmOrderSeats.m_rsSeat = rsTemp
        frmOrderSeats.Show vbModal
        If frmOrderSeats.m_bOk Then
            txtSeat.Text = frmOrderSeats.m_szSeat
        End If
        On Error GoTo 0
    Case GETCHECKGATE
        '得到检票口信息
            On Error GoTo jerr
            Set rsCheckGate = Nothing
            Set rsCheckGate = ConvertStringToCheckGateRS(pszReceive)
            On Error GoTo 0
        
    Case GETACCOUNTLISTID
'            On Error GoTo jerr
'            szTemp = sxBus.GetAccountList(szReceive, m_szSend)
'            wsClient.SendData m_szSend
    Case GETTKPRICEID
'            szTemp = sxBus.GetTkPrice(szReceive, m_szSend)
'            wsClient.SendData m_szSend
    Case Else
    End Select
    Exit Sub
    
jerr:
    ShowErrorMsg
End Sub


Private Sub SellTicket()
    Dim szTemp As String
    
    
    cmdSell.Enabled = False
    m_bSumPriceIsEmpty = True
    On Error GoTo Here
    
    '以下是真正的售票处理

          
    If txtPreferentialSell.Text + txtHalfSell.Text + txtFullSell.Text <> 0 Then
        lblSellMsg.Caption = "正在发送售票请求"
        lblSellMsg.Refresh
        
        If Val(txtFullSell.Text) > 0 And Val(txtHalfSell.Text) = 0 Then
            szTemp = GetSellTicketRequestStr(GetTicketNo, TP_FullPrice, ResolveDisplay(cboStartStation.Text), cboEndStation.BoundText, lvBus.SelectedItem.Text, txtFullSell.Text, CDate(flblSellDate.Caption), txtSeat.Text, lvBus.SelectedItem.SubItems(ID_OffTime))
            wsClient.SendData szTemp
        ElseIf Val(txtHalfSell.Text) > 0 And Val(txtFullSell.Text) = 0 Then
            szTemp = GetSellTicketRequestStr(GetTicketNo, TP_HalfPrice, ResolveDisplay(cboStartStation.Text), cboEndStation.BoundText, lvBus.SelectedItem.Text, txtHalfSell.Text, CDate(flblSellDate.Caption), txtSeat.Text, lvBus.SelectedItem.SubItems(ID_OffTime))
            wsClient.SendData szTemp
        ElseIf Val(txtFullSell.Text) > 0 And Val(txtHalfSell.Text) > 0 Then
            '当为全票\半票同时出售时,选座位无效.
            szTemp = GetSellTicketRequestStr(GetTicketNo, TP_FullPrice, ResolveDisplay(cboStartStation.Text), cboEndStation.BoundText, lvBus.SelectedItem.Text, txtFullSell.Text, CDate(flblSellDate.Caption), "", lvBus.SelectedItem.SubItems(ID_OffTime))
            wsClient.SendData szTemp
            szTemp = GetSellTicketRequestStr(GetTicketNo(txtFullSell.Text - 1), TP_HalfPrice, ResolveDisplay(cboStartStation.Text), cboEndStation.BoundText, lvBus.SelectedItem.Text, txtHalfSell.Text, CDate(flblSellDate.Caption), "", lvBus.SelectedItem.SubItems(ID_OffTime))
            wsClient.SendData szTemp
        End If
        
        Debug.Print "售票数据发送时间:" & Time
    
    End If
    Exit Sub
Here:
    ShowErrorMsg
End Sub

Private Sub ToPrintTicket(pszReceive As String)
    Dim szTemp As String
    Dim i As Integer
'    Dim szBookNumber() As String
'    Dim dbTotalMoney As Double  '总票价
'    Dim dbRealTotalMoney As Double '实际票价
'    Dim aSellTicket() As TSellTicketParam
'    Dim dyBusDate() As Date
'    Dim szBusID() As String
'    Dim szDesStationID() As String
'    Dim szDesStationName() As String
'    Dim szSellStationID() As String
    Dim szSellStationName() As String
    Dim szStartStationName As String
'
'    Dim srSellResult() As TSellTicketResult
'    Dim psgDiscount() As Single
    Dim apiTicketInfo() As TPrintTicketParam
    Dim pszBusDate() As String
    Dim pnTicketCount() As Integer
    Dim pszEndStation() As String
    Dim pszOffTime() As String
    Dim pszBusID() As String
    Dim pszVehicleType() As String
    Dim pszCheckGate() As String
    Dim pbSaleChange() As Boolean
    Dim aszTerminateName() As String
    Dim anInsurance() As Integer '售票用
    Dim panInsurance() As Integer '打印用


'    Dim liTemp As ListItem
'
'    Dim nCount As Integer
'    Dim nLen As Integer
'    Dim nTicketOffset As Integer
'    Dim nLength As Integer
'    Dim nTemp As Integer
'    Dim szTemp As String
    Dim szSeatNo As String
    
    
    Debug.Print "售票数据返回时间:" & Time
    
    
    '处理返回的结果集
    
    lblSellMsg.Caption = "处理返回的结果集"
    lblSellMsg.Refresh
    IncTicketNo txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text
    
    '刷新座位信息
    If chkSetSeat.Enabled Then
        DecBusListViewSeatInfo lvBus, txtFullSell.Text + txtPreferentialSell.Text + txtHalfSell.Text, True
    Else
        DecBusListViewSeatInfo lvBus, txtFullSell.Text + txtPreferentialSell.Text + txtHalfSell.Text, False
    End If
'            flblStandCount.Caption = lvBus.SelectedItem.SubItems(ID_StandCount)
    If lvBus.SelectedItem.SubItems(ID_LimitedCount) > 0 Then
        lvBus.SelectedItem.SubItems(ID_LimitedCount) = lvBus.SelectedItem.SubItems(ID_LimitedCount) - 1
        flblLimitedCount.Caption = GetStationLimitedCountStr(CInt(lvBus.SelectedItem.SubItems(ID_LimitedCount)))
    End If

    '以下是打印票的代码
    ReDim apiTicketInfo(1 To 1)
    ReDim pszBusDate(1 To 1)
    ReDim pnTicketCount(1 To 1)
    ReDim pszEndStation(1 To 1)
    ReDim pszOffTime(1 To 1)
    ReDim pszBusID(1 To 1)
    ReDim pszVehicleType(1 To 1)
    ReDim pszCheckGate(1 To 1)
    ReDim pbSaleChange(1 To 1)
    ReDim apiTicketInfo(1).aptPrintTicketInfo(1 To txtPreferentialSell.Text + txtHalfSell.Text + txtFullSell.Text)
    ReDim aszTerminateName(1 To 1)
    ReDim panInsurance(1 To 1)
    ReDim szSellStationName(1 To 1)
    lblSellMsg.Refresh
    pnTicketCount(1) = txtPreferentialSell.Text + txtHalfSell.Text + txtFullSell.Text
    pszEndStation(1) = GetStationNameInCbo(cboEndStation.Text)
    
    pszVehicleType(1) = lvBus.SelectedItem.SubItems(ID_VehicleModel)
    pszCheckGate(1) = GetCheckName(lvBus.SelectedItem.SubItems(ID_CheckGate))
    pbSaleChange(1) = False
    pszBusDate(1) = flblSellDate.Caption
    pszOffTime(1) = Trim(GetBusOffTime(pszReceive))  'lvBus.SelectedItem.SubItems(ID_OffTime) fpd
    pszBusID(1) = Trim(GetBusID2(pszReceive)) 'lvBus.SelectedItem.Text fpd
    aszTerminateName(1) = lvBus.SelectedItem.SubItems(ID_EndStation)

    panInsurance(1) = IIf(chkInsurance.Value = vbChecked, 2, 0)   '如果选中,则赋为1,否则为0
    
    szSeatNo = GetSeatID(pszReceive)
    
    Dim szTicketNo As String
    szTicketNo = ""
    
    For i = 1 To pnTicketCount(1)
    
        If i > 1 Then
            szTicketNo = GetTicketNo(-(pnTicketCount(1) - i + 1))
        Else
            szTicketNo = GetTicketID(pszReceive)
        End If
    
        apiTicketInfo(1).aptPrintTicketInfo(i).nTicketType = GetTicketType(pszReceive)   'aSellTicket(1).BuyTicketInfo(i).nTicketType
        apiTicketInfo(1).aptPrintTicketInfo(i).sgTicketPrice = PackageToMoney(GetTicketPrice(pszReceive))    ' srSellResult(1).asgTicketPrice(i)
        '下面需要考虑多张票的情况
        
        apiTicketInfo(1).aptPrintTicketInfo(i).szSeatNo = MidA(szSeatNo, i * 2 - 1, 2)  'srSellResult(1).aszSeatNo(i)
        apiTicketInfo(1).aptPrintTicketInfo(i).szTicketNo = szTicketNo ' aSellTicket(1).BuyTicketInfo(i).szTicketNo
        
'        '取得实际总票价
'        If apiTicketInfo(1).aptPrintTicketInfo(i).nTicketType <> TP_FreeTicket Then
'            dbRealTotalMoney = apiTicketInfo(1).aptPrintTicketInfo(i).sgTicketPrice + dbRealTotalMoney
'        End If
    Next
    
    ResolveDisplay lvBus.SelectedItem.Tag, szStartStationName
    szSellStationName(1) = szStartStationName

    '进行票款累加
    If IsNumeric(txtReceivedMoney.Text) Then

        If txtReceivedMoney.Text = 0 Then
            m_sgTotalMoney = lblTotalMoney.Caption
        Else
            m_sgTotalMoney = 0#
        End If
    End If
    lblSellMsg.Caption = "正在打印车票"
    lblSellMsg.Refresh


    PrintTicket apiTicketInfo, pszBusDate, pnTicketCount, pszEndStation, pszOffTime, pszBusID, pszVehicleType, pszCheckGate, pbSaleChange, aszTerminateName, szSellStationName, panInsurance

'    dbTotalMoney = CDbl(flblTotalPrice.Caption)




    m_bPreClear = True
    lblSellMsg.Caption = ""
    cmdSell.Enabled = True
    
    txtPrevDate.Text = 0
    lblmileate.Caption = ""
    chkInsurance.Value = vbUnchecked
    flblRestMoney.Caption = szTemp
    frmOrderSeats.m_szBookNumber = ""
    txtSeat.Text = ""
'    cboEndStation.SetFocus
    txtReceivedMoney.SetFocus
    m_bSetFocus = False
    Exit Sub
Here:
    frmOrderSeats.m_szBookNumber = ""
    lblSellMsg.Caption = ""
'    m_bPrint = False
    If err.Number = 91 Then
        MsgBox "该天该站点已无车次！", vbInformation + vbOKOnly, "提示"
    Else
        frmNotify.m_szErrorDescription = err.Description
        frmNotify.Show vbModal
    End If
    txtPrevDate.Text = 0
    SetPreSellButton
    cboEndStation.SetFocus
End Sub



Private Sub RefreshBus(prsBus As Recordset)
    '处理接收到记录集后的操作
    Dim i As Integer
    Dim szStationID As String
    Dim rsTemp As Recordset
    Dim liTemp As ListItem
    Dim lForeColor As OLE_COLOR
    Dim nBusType As EBusType
    Dim j As Integer
    Dim k As Integer
    Dim lvS As ListItem
    Dim szScrollBus As String
'


'    If prsBus.RecordCount <> 0 Then
        'lblmileate = m_rsBusInfo!end_station_mileage & "公里"
'    End If
    Set m_rsBusInfo = prsBus.Clone '克隆该记录集
    '如果无车次,则返回到站点输入处
    If m_rsBusInfo.RecordCount = 0 Then
        If ActiveControl Is lvBus Then
            '只有到车次列表处时,无车次才返回到站点输入
            cboEndStation.SetFocus
        End If
    End If
    lvBus.Sorted = False
    lvBus.ListItems.Clear
    lvBus.Refresh
    For j = 1 To m_rsBusInfo.RecordCount
        If lvBus.ListItems.count > m_ISellScreenShow And m_ISellScreenShow <> 0 Then
            With lvBus
                If .ListItems.count > 0 Then
                    .SortKey = MDISellTicket.GetSortKey() - 1
                    .Sorted = True
                    For i = 1 To .ListItems.count
                    '如果车次不是停班而且(车次有座位或站票),则让该车次选中
                        If .ListItems(i).ForeColor <> m_lStopBusColor And (.ListItems(i).SubItems(ID_SeatCount) > 0) Then

                            .ListItems(i).Selected = True
                            .ListItems(i).EnsureVisible
                            Exit For
                        End If
                    Next i
                    If i > .ListItems.count Then
                        .ListItems(1).Selected = True
                        .ListItems(1).EnsureVisible
                    End If
                End If
            End With
'            m_rsBusInfo.MovePrevious
            m_rsBusInfo.MovePrevious
'                    Set rsTemp = Nothing
            Set liTemp = Nothing
            Exit Sub
        Else
        '可根据参数选择是否加入已售完车次
        If m_bListNoSeatBus Or m_rsBusInfo!sale_seat_quantity > 0 Then '+ m_rsBusInfo!sale_stand_seat_quantity

'                    If Hour(m_rsBusInfo!busstarttime) >= txtTime.Text Then
                '如果车次时间大于查询的时间
                For i = lvBus.ListItems.count To 1 Step -1
                    If RTrim(m_rsBusInfo!bus_id) = lvBus.ListItems(i) And Format(m_rsBusInfo!bus_date, "yyyy-mm-dd") = CDate(flblSellDate.Caption) Then
'                                Select Case Trim(m_rsBusInfo!seat_type_id)
'                                Case cszSeatType
'                                    liTemp.SubItems(ID_FullPrice) = m_rsBusInfo!full_price
'                                    liTemp.SubItems(ID_HalfPrice) = m_rsBusInfo!half_price
'                                    liTemp.SubItems(ID_PreferentialPrice1) = m_rsBusInfo!preferential_ticket1
'                                    liTemp.SubItems(ID_PreferentialPrice2) = m_rsBusInfo!preferential_ticket2
'                                    liTemp.SubItems(ID_PreferentialPrice3) = m_rsBusInfo!preferential_ticket3
'                                Case cszBedType
'                                    liTemp.SubItems(ID_BedFullPrice) = m_rsBusInfo!full_price
'                                    liTemp.SubItems(ID_BedHalfPrice) = m_rsBusInfo!half_price
'                                    liTemp.SubItems(ID_BedPreferentialPrice1) = m_rsBusInfo!preferential_ticket1
'                                    liTemp.SubItems(ID_BedPreferentialPrice2) = m_rsBusInfo!preferential_ticket2
'                                    liTemp.SubItems(ID_BedPreferentialPrice3) = m_rsBusInfo!preferential_ticket3
'                                Case cszAdditionalType
'                                    liTemp.SubItems(ID_AdditionalFullPrice) = m_rsBusInfo!full_price
'                                    liTemp.SubItems(ID_AdditionalHalfPrice) = m_rsBusInfo!half_price
'                                    liTemp.SubItems(ID_AdditionalPreferential1) = m_rsBusInfo!preferential_ticket1
'                                    liTemp.SubItems(ID_AdditionalPreferential2) = m_rsBusInfo!preferential_ticket2
'                                    liTemp.SubItems(ID_AdditionalPreferential3) = m_rsBusInfo!preferential_ticket3
'                                End Select
                        GoTo nextstep
                    End If
'                            Exit For
                Next i
                If m_rsBusInfo!status = ST_BusStopped Or m_rsBusInfo!status = ST_BusMergeStopped Or m_rsBusInfo!status = ST_BusSlitpStop Then
'                    lForeColor = m_lStopBusColor
'                    Set liTemp = lvBus.ListItems.Add(, "A" & FormatDbValue(m_rsBusInfo!bus_id), FormatDbValue(m_rsBusInfo!bus_id))
'                    liTemp.SmallIcon = "StopBus"
                    GoTo nextstep
                Else
                    lForeColor = m_lNormalBusColor
                    Set liTemp = lvBus.ListItems.Add(, "A" & FormatDbValue(m_rsBusInfo!bus_id), FormatDbValue(m_rsBusInfo!bus_id))
                    '车次代码"A" & RTrim(m_rsBusInfo！bus_id)
                End If
                nBusType = m_rsBusInfo!bus_type
'                        If lForeColor <> m_lStopBusColor Then
                    liTemp.ForeColor = lForeColor

'                         varBookmark = m_rsBusInfo.Bookmark
'                                If m_rsBusInfo.RecordCount > 0 Then
'                                   RefreshSellStation m_rsBusInfo
'                                End If
'                           m_rsBusInfo.Bookmark = varBookmark

                    If nBusType <> TP_ScrollBus Then
                        liTemp.SubItems(ID_BusType) = FormatDbValue(m_rsBusInfo!bus_type)
                        liTemp.SubItems(ID_OffTime) = Format(m_rsBusInfo!busstarttime, "hh:mm")
                    Else
                        liTemp.SubItems(ID_VehicleModel) = cszScrollBus
                        liTemp.SubItems(ID_OffTime) = cszScrollBus
                    End If
                    liTemp.SubItems(ID_RouteName) = FormatDbValue(m_rsBusInfo!route_name)
                    liTemp.SubItems(ID_EndStation) = FormatDbValue(m_rsBusInfo!end_station_name)
                    liTemp.SubItems(ID_TotalSeat) = FormatDbValue(m_rsBusInfo!total_seat)
                    If IsDate(liTemp.SubItems(ID_OffTime)) Then
                        If g_bIsBookValid And DateAdd("n", -g_nBookTime, liTemp.SubItems(ID_OffTime)) < Time And ToDBDate(flblSellDate.Caption) = ToDBDate(Date) Then
                            '如果车次日期为当天,且已过预定时限,则将预定人数加到可售张数上面.
                            liTemp.SubItems(ID_BookCount) = 0
                            liTemp.SubItems(ID_SeatCount) = FormatDbValue(m_rsBusInfo!sale_seat_quantity) + FormatDbValue(m_rsBusInfo!book_count)

                        Else
                            liTemp.SubItems(ID_BookCount) = FormatDbValue(m_rsBusInfo!book_count)
                            liTemp.SubItems(ID_SeatCount) = FormatDbValue(m_rsBusInfo!sale_seat_quantity)
                        End If
                    Else

                        liTemp.SubItems(ID_BookCount) = FormatDbValue(m_rsBusInfo!book_count)
                        liTemp.SubItems(ID_SeatCount) = FormatDbValue(m_rsBusInfo!sale_seat_quantity)
                    End If
                    liTemp.SubItems(ID_SeatTypeCount) = FormatDbValue(m_rsBusInfo!seat_remain)
                    liTemp.SubItems(ID_BedTypeCount) = FormatDbValue(m_rsBusInfo!bed_remain)
                    liTemp.SubItems(ID_AdditionalCount) = FormatDbValue(m_rsBusInfo!additional_remain)
                    liTemp.SubItems(ID_VehicleModel) = FormatDbValue(m_rsBusInfo!vehicle_type_name)
                    Select Case FormatDbValue(m_rsBusInfo!seat_type_id)
                    Case cszSeatType
                        liTemp.SubItems(ID_FullPrice) = FormatDbValue(m_rsBusInfo!full_price)
                        liTemp.SubItems(ID_HalfPrice) = FormatDbValue(m_rsBusInfo!half_price)
                        liTemp.SubItems(ID_FreePrice) = 0
                        liTemp.SubItems(ID_PreferentialPrice1) = FormatDbValue(m_rsBusInfo!preferential_ticket1)
                        liTemp.SubItems(ID_PreferentialPrice2) = FormatDbValue(m_rsBusInfo!preferential_ticket2)
                        liTemp.SubItems(ID_PreferentialPrice3) = FormatDbValue(m_rsBusInfo!preferential_ticket3)
                        liTemp.SubItems(ID_BedFullPrice) = 0
                        liTemp.SubItems(ID_BedHalfPrice) = 0
                        liTemp.SubItems(ID_BedFreePrice) = 0
                        liTemp.SubItems(ID_BedPreferentialPrice1) = 0
                        liTemp.SubItems(ID_BedPreferentialPrice2) = 0
'                        liTemp.SubItems(ID_BedPreferentialPrice3) = 0
'                        liTemp.SubItems(ID_AdditionalFullPrice) = 0
'                        liTemp.SubItems(ID_AdditionalHalfPrice) = 0
'                        liTemp.SubItems(ID_AdditionalFreePrice) = 0
'                        liTemp.SubItems(ID_AdditionalPreferential1) = 0
'                        liTemp.SubItems(ID_AdditionalPreferential2) = 0
'                        liTemp.SubItems(ID_AdditionalPreferential3) = 0
                    Case cszBedType
                        liTemp.SubItems(ID_FullPrice) = 0
                        liTemp.SubItems(ID_HalfPrice) = 0
                        liTemp.SubItems(ID_FreePrice) = 0
                        liTemp.SubItems(ID_PreferentialPrice1) = 0
                        liTemp.SubItems(ID_PreferentialPrice2) = 0
                        liTemp.SubItems(ID_PreferentialPrice3) = 0
                        liTemp.SubItems(ID_BedFullPrice) = FormatDbValue(m_rsBusInfo!full_price)
                        liTemp.SubItems(ID_BedHalfPrice) = FormatDbValue(m_rsBusInfo!half_price)
                        liTemp.SubItems(ID_BedFreePrice) = 0
                        liTemp.SubItems(ID_BedPreferentialPrice1) = FormatDbValue(m_rsBusInfo!preferential_ticket1)
                        liTemp.SubItems(ID_BedPreferentialPrice2) = FormatDbValue(m_rsBusInfo!preferential_ticket2)
                        liTemp.SubItems(ID_BedPreferentialPrice3) = FormatDbValue(m_rsBusInfo!preferential_ticket3)
                        liTemp.SubItems(ID_AdditionalFullPrice) = 0
                        liTemp.SubItems(ID_AdditionalHalfPrice) = 0
                        liTemp.SubItems(ID_AdditionalFreePrice) = 0
                        liTemp.SubItems(ID_AdditionalPreferential1) = 0
                        liTemp.SubItems(ID_AdditionalPreferential2) = 0
                        liTemp.SubItems(ID_AdditionalPreferential3) = 0
                    Case cszAdditionalType
                        liTemp.SubItems(ID_FullPrice) = 0
                        liTemp.SubItems(ID_HalfPrice) = 0
                        liTemp.SubItems(ID_FreePrice) = 0
                        liTemp.SubItems(ID_PreferentialPrice1) = 0
                        liTemp.SubItems(ID_PreferentialPrice2) = 0
                        liTemp.SubItems(ID_PreferentialPrice3) = 0
                        liTemp.SubItems(ID_BedFullPrice) = 0
                        liTemp.SubItems(ID_BedHalfPrice) = 0
                        liTemp.SubItems(ID_BedFreePrice) = 0
                        liTemp.SubItems(ID_BedPreferentialPrice1) = 0
                        liTemp.SubItems(ID_BedPreferentialPrice2) = 0
                        liTemp.SubItems(ID_BedPreferentialPrice3) = 0
                        liTemp.SubItems(ID_AdditionalFullPrice) = FormatDbValue(m_rsBusInfo!full_price)
                        liTemp.SubItems(ID_AdditionalHalfPrice) = FormatDbValue(m_rsBusInfo!half_price)
                        liTemp.SubItems(ID_AdditionalFreePrice) = 0
                        liTemp.SubItems(ID_AdditionalPreferential1) = FormatDbValue(m_rsBusInfo!preferential_ticket1)
                        liTemp.SubItems(ID_AdditionalPreferential2) = FormatDbValue(m_rsBusInfo!preferential_ticket2)
                        liTemp.SubItems(ID_AdditionalPreferential3) = FormatDbValue(m_rsBusInfo!preferential_ticket3)
                    End Select
                    '以下几列不显示出来，只是将其存储，以备后用
                    liTemp.SubItems(ID_LimitedCount) = FormatDbValue(m_rsBusInfo!sale_ticket_quantity)
                    liTemp.SubItems(ID_LimitedTime) = FormatDbValue(m_rsBusInfo!stop_sale_time)
                    liTemp.SubItems(ID_BusType1) = nBusType
                    liTemp.SubItems(ID_CheckGate) = FormatDbValue(m_rsBusInfo!sell_check_gate_id)
                    liTemp.SubItems(ID_StandCount) = 0 ' m_rsBusInfo!sale_stand_seat_quantity
                    liTemp.Tag = MakeDisplayString(FormatDbValue(m_rsBusInfo!sell_station_id), FormatDbValue(m_rsBusInfo!sell_station_name))

'以下一句：停班车整行变色
                    If lForeColor = m_lStopBusColor Then
                        Dim oSubLtems As ListSubItem
                        For Each oSubLtems In liTemp.ListSubItems
                            oSubLtems.ForeColor = lForeColor
                        Next
                    End If
'                        End If
            End If

        End If


nextstep:
        m_rsBusInfo.MoveNext
'        m_rsBusInfo.MoveNext
    Next j
    
    lvBus.Sorted = True

    
   If lvBus.ListItems.count > 0 Then
      RefreshSellStation m_rsBusInfo
   Else
      lvSellStation.ListItems.Clear
   End If
    DoThingWhenBusChange
    
    
'    Set liTemp = Nothing
    On Error GoTo 0
    Exit Sub
Here:
    ShowErrorMsg
    Set m_rsBusInfo = Nothing
'    Set liTemp = Nothing
End Sub

'得到检票口名称和代码
Public Function GetCheckName(pszCheckGateID As String) As String
    Dim i As Integer
    Dim szResult As String
    Dim nLen As Integer
    nLen = 0
    nLen = rsCheckGate.RecordCount
    szResult = ""
    rsCheckGate.MoveFirst
    For i = 1 To nLen
        If Trim(rsCheckGate!check_gate_id) = Trim(pszCheckGateID) Then
            szResult = Trim(rsCheckGate!check_gate_name)
            Exit For
        End If
        rsCheckGate.MoveNext
    Next i
    GetCheckName = szResult

End Function


