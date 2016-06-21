VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{695ABF14-B2D8-11D2-A5ED-DE08DCF33612}#1.2#0"; "asfcombo.ocx"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{BBF95DAA-F9CB-4CA9-A673-E0E9E0193752}#1.0#0"; "STSellCtl.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmSellDiscountTkt 
   BackColor       =   &H00929292&
   Caption         =   "售折扣票"
   ClientHeight    =   7710
   ClientLeft      =   240
   ClientTop       =   1935
   ClientWidth     =   12015
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7710
   ScaleWidth      =   12015
   Begin VB.Frame fraOutLine 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   7035
      Left            =   105
      TabIndex        =   20
      Top             =   30
      Width           =   11835
      Begin FCmbo.asFlatCombo cboSeatType 
         Height          =   330
         Left            =   960
         TabIndex        =   29
         Top             =   2520
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
      Begin VB.Frame fraDiscountTicket 
         Caption         =   "折扣票"
         Height          =   585
         Left            =   7380
         TabIndex        =   32
         Top             =   5655
         Visible         =   0   'False
         Width           =   3075
         Begin VB.TextBox txtDiscount 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1680
            TabIndex        =   33
            Text            =   "1"
            Top             =   180
            Width           =   1185
         End
         Begin VB.Label lblDiscount 
            Caption         =   "折扣(&F):"
            Height          =   225
            Left            =   360
            TabIndex        =   34
            Top             =   225
            Width           =   795
         End
      End
      Begin VB.CommandButton cmdPreSell 
         Caption         =   "预售"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   9150
         TabIndex        =   31
         Top             =   3960
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Timer Timer1 
         Interval        =   60000
         Left            =   4800
         Top             =   510
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
         Height          =   480
         Left            =   1020
         TabIndex        =   9
         Text            =   "0.00"
         Top             =   5040
         Width           =   1575
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
         Left            =   10515
         TabIndex        =   22
         Top             =   75
         Width           =   990
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
         Left            =   8685
         TabIndex        =   14
         Top             =   45
         Width           =   1755
      End
      Begin VB.CheckBox chkInsurance 
         BackColor       =   &H00FFC0C0&
         Caption         =   "保险(F12)"
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   3825
         TabIndex        =   21
         Top             =   5025
         Width           =   1290
      End
      Begin STSellCtl.ucNumTextBox txtTime 
         Height          =   330
         Left            =   5790
         TabIndex        =   12
         Top             =   75
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   582
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RTComctl3.CoolButton cmdSell 
         Height          =   615
         Left            =   60
         TabIndex        =   10
         Top             =   6300
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   1085
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
         MICON           =   "frmSellDiscountTkt.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin STSellCtl.ucNumTextBox txtPreferentialSell 
         Height          =   390
         Left            =   1440
         TabIndex        =   7
         Top             =   3900
         Width           =   840
         _ExtentX        =   1482
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
      Begin STSellCtl.ucNumTextBox txtHalfSell 
         Height          =   390
         Left            =   1410
         TabIndex        =   23
         Top             =   3360
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
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
         TabIndex        =   24
         Top             =   2925
         Visible         =   0   'False
         Width           =   1350
         _ExtentX        =   2381
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
         Enabled         =   0   'False
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
      Begin MSComctlLib.ListView lvPreSell 
         Height          =   1635
         Left            =   5760
         TabIndex        =   25
         Top             =   5280
         Width           =   5925
         _ExtentX        =   10451
         _ExtentY        =   2884
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
         NumItems        =   0
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
               Picture         =   "frmSellDiscountTkt.frx":001C
               Key             =   ""
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
      Begin MSComctlLib.ListView lvBus 
         Height          =   4470
         Left            =   2640
         TabIndex        =   5
         Top             =   495
         Width           =   9105
         _ExtentX        =   16060
         _ExtentY        =   7885
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
      Begin RTComctl3.FlatLabel flblLimitedTime 
         Height          =   315
         Left            =   3240
         TabIndex        =   27
         Top             =   2820
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
      Begin RTComctl3.FlatLabel flblStandCount 
         Height          =   315
         Left            =   5880
         TabIndex        =   28
         Top             =   2820
         Width           =   825
         _ExtentX        =   1455
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
         Caption         =   "0"
      End
      Begin FCmbo.asFlatCombo cboPreferentialTicket 
         Height          =   330
         Left            =   120
         TabIndex        =   6
         Top             =   3900
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         ButtonDisabledForeColor=   12632256
         Enabled         =   0   'False
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
      Begin RTComctl3.FlatLabel flblLimitedCount 
         Height          =   315
         Left            =   1440
         TabIndex        =   30
         Top             =   1560
         Visible         =   0   'False
         Width           =   930
         _ExtentX        =   1640
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
      Begin MSComctlLib.ListView lvSellStation 
         Height          =   1635
         Left            =   2640
         TabIndex        =   16
         Top             =   5280
         Width           =   3030
         _ExtentX        =   5345
         _ExtentY        =   2884
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
         NumItems        =   4
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
      End
      Begin STSellCtl.ucSuperCombo cboEndStation 
         Height          =   2580
         Left            =   90
         TabIndex        =   3
         Top             =   1200
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   4551
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.UpDown ucPreferential 
         Height          =   390
         Left            =   2280
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   3885
         Width           =   255
         _ExtentX        =   423
         _ExtentY        =   688
         _Version        =   393216
         BuddyControl    =   "txtPreferentialSell"
         BuddyDispid     =   196656
         OrigLeft        =   2370
         OrigTop         =   3090
         OrigRight       =   2625
         OrigBottom      =   3405
         Max             =   100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   1745027080
         Enabled         =   -1  'True
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
         Left            =   180
         TabIndex        =   38
         Top             =   2565
         Visible         =   0   'False
         Width           =   480
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
         TabIndex        =   40
         Top             =   3030
         Visible         =   0   'False
         Width           =   960
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
         TabIndex        =   52
         Top             =   900
         Visible         =   0   'False
         Width           =   180
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
         Left            =   0
         TabIndex        =   51
         Top             =   5880
         Width           =   1080
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00000000&
         X1              =   30
         X2              =   2670
         Y1              =   5550
         Y2              =   5550
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
         Left            =   0
         TabIndex        =   8
         Top             =   5160
         Width           =   1170
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
         Left            =   0
         TabIndex        =   50
         Top             =   4620
         Width           =   840
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C0C0C0&
         Visible         =   0   'False
         X1              =   2550
         X2              =   5190
         Y1              =   4410
         Y2              =   4410
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
         TabIndex        =   49
         Top             =   2130
         Width           =   1215
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "车票单价:"
         Height          =   180
         Left            =   5970
         TabIndex        =   48
         Top             =   2310
         Width           =   810
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
         TabIndex        =   4
         Top             =   135
         Width           =   1440
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
         TabIndex        =   2
         Top             =   900
         Width           =   960
      End
      Begin VB.Label Label1 
         Caption         =   "限售张数:"
         Height          =   225
         Left            =   300
         TabIndex        =   47
         Top             =   1650
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "限售时间:"
         Height          =   195
         Left            =   3120
         TabIndex        =   46
         Top             =   2850
         Width           =   855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "站票:"
         Height          =   180
         Left            =   5190
         TabIndex        =   45
         Top             =   2880
         Width           =   450
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
         TabIndex        =   44
         Top             =   2610
         Width           =   840
      End
      Begin VB.Line Line2 
         X1              =   8730
         X2              =   11220
         Y1              =   2250
         Y2              =   2250
      End
      Begin VB.Line Line5 
         X1              =   8760
         X2              =   11250
         Y1              =   3180
         Y2              =   3180
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
         TabIndex        =   43
         Top             =   2460
         Width           =   1320
      End
      Begin VB.Label lblPreSell 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "预售车次列表(F12):"
         Height          =   180
         Left            =   5805
         TabIndex        =   42
         Top             =   5040
         Width           =   1620
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
         Left            =   9480
         TabIndex        =   41
         Top             =   5010
         Width           =   1875
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
         TabIndex        =   39
         Top             =   3465
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         Visible         =   0   'False
         X1              =   2550
         X2              =   5190
         Y1              =   4920
         Y2              =   4920
      End
      Begin VB.Label flblTotalPrice 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
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
         Left            =   1605
         TabIndex        =   37
         Top             =   4530
         Width           =   960
      End
      Begin VB.Label flblRestMoney 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
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
         Left            =   1605
         TabIndex        =   36
         Top             =   5790
         Width           =   960
      End
      Begin VB.Label lblTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "发车时间(&S):"
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
         Left            =   4200
         TabIndex        =   11
         Top             =   120
         Width           =   1440
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "点以后"
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
         Left            =   6405
         TabIndex        =   35
         Top             =   120
         Width           =   720
      End
      Begin VB.Label lblsellstation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "上车站(&O):"
         Height          =   180
         Left            =   2655
         TabIndex        =   15
         Top             =   5040
         Width           =   900
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
         Left            =   7395
         TabIndex        =   13
         Top             =   120
         Width           =   1200
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1245
      HelpContextID   =   3000411
      Left            =   4650
      TabIndex        =   17
      Top             =   7515
      Visible         =   0   'False
      Width           =   2880
      Begin VB.CheckBox chkSetSeat 
         BackColor       =   &H00FFFFFF&
         Caption         =   "定座(&P)"
         Height          =   270
         HelpContextID   =   3000411
         Left            =   120
         TabIndex        =   18
         Top             =   -30
         Width           =   975
      End
      Begin VB.Label lblSetSeat 
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   180
         TabIndex        =   19
         Top             =   60
         Width           =   435
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
      TabIndex        =   53
      Top             =   0
      Visible         =   0   'False
      Width           =   645
   End
End
Attribute VB_Name = "frmSellDiscountTkt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_bPrint As Boolean
Private m_blPointCount As Boolean
Private m_bSumPriceIsEmpty As Boolean   '总票价为0
Private m_nCount As Integer '隔一段时间读取服务器时间的自加一的变量
Private m_sgTotalMoney As Single '记录上一次售票的金额
Private m_atTicketType() As TTicketType
Private m_dbTotalPrice As Double
Private m_aszSeatType() As String
Private m_atbSeatTypeBus As TMultiSeatTypeBus '得到不同座位类型的车次
Private m_TicketPrice() As Single '存储票价
Private m_TicketTypeDetail() As ETicketType '存储票种
Private m_bPreClear As Boolean
Private m_bSetFocus As Boolean
Private m_bPreSellFocus As Boolean
Private m_rsBusInfo As Recordset
Private m_atbBusOrder() As TBusOrderCount
Private m_aszCheckGateInfo() As String
Private m_bNotRefresh As Boolean '是否需要刷新,主要是在设置查询车次时间时用到.
Private rsCountTemp As Recordset
Private nSellCount As Integer '出售张数
Private m_oSellDiscount As New SellTicketClient

Private Sub cboEndStation_Change()
On Error GoTo Here
    If lvBus.ListItems.count > 0 Then
        DoThingWhenBusChange
    Else
       lvSellStation.ListItems.Clear
    End If
    DealPrice
    
    
    cmdPreSell.Enabled = True
On Error GoTo 0
Exit Sub
Here:
  ShowErrorMsg
End Sub

Private Sub cboEndStation_GotFocus()
On Error GoTo Here
    lblToStation.ForeColor = clActiveColor
    DealPrice
    
On Error GoTo 0
Exit Sub
Here:
 ShowErrorMsg
End Sub



Private Sub cboEndStation_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyLeft
'            KeyCode = 0
'            If Val(txtPrevDate.Text) > 0 Then
'                txtPrevDate.Text = Val(txtPrevDate.Text) - 1
'            End If
'            m_bNotRefresh = False
        Case vbKeyRight
'            KeyCode = 0
'            If Val(txtPrevDate.Text) < m_nCanSellDay Then
'
'                txtPrevDate.Text = Val(txtPrevDate.Text) + 1
'            End If
'            m_bNotRefresh = False
        Case 106 'Asc("*")
            '+号则跳到输入时间处
            KeyCode = 0
            txtTime.SetFocus
            m_bNotRefresh = True
        Case Else
            m_bNotRefresh = False
    End Select
    If m_bPreClear Then
        lvPreSell.ListItems.Clear
        flblTotalPrice.Caption = 0#
        txtReceivedMoney.Text = ""
        flblRestMoney.Caption = ""
        m_bPreClear = False
    End If
   
End Sub

Private Sub cboEndStation_LostFocus()
On Error GoTo Here
Dim nIndex As Integer
    lblToStation.ForeColor = 0
    If m_bNotRefresh Then Exit Sub '如果是跳到了输入时间处,则不刷新
    
    DoThingWhenBusChange
    SetDefaultSellTicket
    txtReceivedMoney.Text = ""
    RefreshBus True
    
'    If IsHaveScrollBus Then  '判断是否有滚动车次
'        nIndex = IsExitInTeam(Trim(cboEndStation.BoundText))
'        If nIndex = 0 Then
'             InitScrollBusOrder
'        Else
'             AddValueToIndex nIndex
'             SetCorrectBusOrder Trim(cboEndStation.BoundText)
'        End If
'    End If
    If m_bPreClear Then
        lvPreSell.ListItems.Clear
        flblTotalPrice.Caption = 0#
        txtReceivedMoney.Text = ""
        flblRestMoney.Caption = ""
        m_bPreClear = False
    End If
    DealPrice
    
On Error GoTo 0
Exit Sub
Here:
    ShowErrorMsg
End Sub

Private Sub cboPreferentialTicket_Change()
    txtPreferentialSell.Text = 0
    cmdPreSell.Enabled = True
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
   RefreshBusStation rsCountTemp, Trim(lvSellStation.SelectedItem.SubItems(3)), cboSeatType.ListIndex + 1
  End If
End Sub

Private Sub cboSeatType_GotFocus()
    lblSeatType.ForeColor = clActiveColor
End Sub

Private Sub cboSeatType_KeyPress(KeyAscii As Integer)
    lvBus.SetFocus
End Sub

Private Sub cboSeatType_LostFocus()
    lblSeatType.ForeColor = 0
End Sub

Private Sub chkInsurance_Click()
    DealPrice
End Sub

Private Sub cmdPreSell_Click()
    On Error GoTo Here
    Dim nSameBusIndex As Integer
    If txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text = 0 Then Exit Sub
    cmdPreSell.Enabled = False
'If nSameBusIndex = 0 Then
    GetPreSellTicketInfo
'Else
'    MergeSameBusInfo nSameBusIndex
'End If
    txtFullSell.Text = 0
    txtHalfSell.Text = 0
    txtPreferentialSell.Text = 0
    SetPreSellButton
    DealPrice
    cmdPreSell.Enabled = True
Exit Sub
Here:
    ShowErrorMsg
End Sub

'售票
Private Sub cmdSell_Click()
Dim k As Long
Dim m As Long
Dim i As Integer

m = 0
    For i = 1 To lvPreSell.ListItems.count
        m = m + lvPreSell.ListItems(i).SubItems(IT_SumTicketNum)
    Next i
    If m_lEndTicketNoOld = 0 Then
        ShowMsg "售票不成功，用户还未领票，请先去领票！"
        Exit Sub
    End If
    If m + Val(txtFullSell.Text) + Val(txtHalfSell.Text) + Val(txtPreferentialSell.Text) + Val(m_lTicketNo) - 1 > Val(m_lEndTicketNo) Then
        k = Val(m_lEndTicketNo) - Val(m_lTicketNo) + 1
        MsgBox "打印机上的票已不够！" & vbCrLf & "车票只剩 " & k & "张", vbInformation, "售票台"
    Else
        SellTicket
        
    End If
End Sub

Private Sub cmdSetSeat_Click()
On Error GoTo Here
    Dim rsTemp As Recordset
    If lvBus.SelectedItem Is Nothing Then
        Set rsTemp = Nothing
        Exit Sub
    End If
    Set rsTemp = m_oSellDiscount.GetSeatRs(CDate(flblSellDate.Caption), lvBus.SelectedItem.Text)
    Set frmOrderSeats.m_rsSeat = rsTemp
    Set rsTemp = m_oSell.GetBookRs(CDate(flblSellDate.Caption), lvBus.SelectedItem.Text)
    Set frmOrderSeats.m_rsBook = rsTemp
    frmOrderSeats.m_szSeatNumber = PreOrderSeat
    frmOrderSeats.Show vbModal
    If frmOrderSeats.m_bOk Then
        txtSeat.Text = frmOrderSeats.m_szSeat
    End If
    Set rsTemp = Nothing
On Error GoTo 0
Exit Sub
Here:
    Set rsTemp = Nothing
    ShowErrorMsg
End Sub



Private Sub Command1_Click()
    frmNotify.Show vbModal
End Sub


Private Sub Form_Activate()
On Error GoTo Here
    
    SetPreSellButton
'    MDISellTicket.SetFunAndUnit
    lvBus.SortKey = MDISellTicket.GetSortKey()
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
'    ElseIf KeyCode = vbKeyF12 And lvPreSell.ListItems.count <> 0 Then
'        lvPreSell.SetFocus
    ElseIf KeyCode = vbKeyCapital And Shift Then
        If lvBus.GridLines = True Then
            lvBus.GridLines = False
        Else
            lvBus.GridLines = True
        End If
    ElseIf KeyCode = vbKeyF12 Then
        '选中需要保险
        If chkInsurance.Value = vbChecked Then
            chkInsurance.Value = vbUnchecked
        Else
            chkInsurance.Value = vbChecked
        End If
        
    End If
    
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 45 Then
     If lvBus.ListItems.count > 0 Then
      If lvBus.SelectedItem.SubItems(ID_OffTime) = cszScrollBus Then
         Exit Sub
      End If
    End If
        KeyAscii = 0
        ChangeSeatType
        Exit Sub
    ElseIf KeyAscii = 27 Then
        lvPreSell.ListItems.Clear
        SetDefaultSellTicket
        txtPrevDate.Text = 0
        txtTime.Text = 0
        lblmileate.Caption = ""
        txtSeat.Text = ""
        cboEndStation.Text = ""
        cboEndStation.SetFocus
        
    ElseIf KeyAscii = 13 And (lvSellStation.Enabled) And (Me.ActiveControl Is lvBus) Then
        lvSellStation.SetFocus
        Exit Sub
    ElseIf KeyAscii = 13 And Not (Me.ActiveControl Is cboEndStation) And Not (Me.ActiveControl Is lvPreSell) And Not (Me.ActiveControl Is txtReceivedMoney) _
                And Not (Me.ActiveControl Is txtHalfSell) And Not (Me.ActiveControl Is txtPreferentialSell) And Not (Me.ActiveControl Is txtTime) _
                And Not (Me.ActiveControl Is txtFullSell) Then
            SendKeys "{TAB}"
    ElseIf KeyAscii = 43 Then
        txtPrevDate.Text = 0
        txtTime.Text = 0
        cboEndStation.SetFocus
    ElseIf KeyAscii = Asc("*") And Not (Me.ActiveControl Is cboEndStation) Then
        txtTime.SetFocus
    End If
    
End Sub

Private Sub Form_Load()
On Error GoTo Here
    flblSellDate.Caption = ToStandardDateStr(m_oParam.NowDate)
    m_oSellDiscount.Init m_oAUser
    txtPrevDate.Text = 0
    txtTime.Text = 0
    m_dbTotalPrice = 0
    m_bPrint = False
    RefreshPreferentialTicket '读取优惠票信息
    GetPreSellBus  '显示预售状态信息
    RefreshStation2
    GetInitCheckGate
    
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
'
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
'    m_clSell.Remove GetEncodedKey(Me.Tag)
'    MDISellTicket.lblSell.Value = vbUnchecked
'    MDISellTicket.abMenuTool.Bands("mnuFunction").Tools("mnuSellTkt").Checked = False
    MDISellTicket.EnableSortAndRefresh False
On Error GoTo 0
Exit Sub
Here:
    ShowErrorMsg
End Sub






Private Sub lvBus_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    SortListView lvBus, ColumnHeader.Index
End Sub



Private Sub lvBus_GotFocus()
    lblBus.ForeColor = clActiveColor
    ShowRightSeatType
    If lvBus.ListItems.count = 0 Then
        cboEndStation.SetFocus
        Exit Sub
    End If
    
End Sub

Private Sub lvBus_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo Here
      
        RefreshSellStation rsCountTemp
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
            flblStandCount.Caption = ""
        Else
            flblLimitedCount.Caption = GetStationLimitedCountStr(CInt(liTemp.SubItems(ID_LimitedCount)))
            flblLimitedTime.Caption = GetStationLimitedTimeStr(CInt(liTemp.SubItems(ID_LimitedTime)), CDate(flblSellDate.Caption), CDate(liTemp.SubItems(ID_OffTime)))
           ' flblStandCount.Caption = liTemp.subitems(ID_StandCount)
           flblStandCount.Caption = 0
        End If
    Else
        lblSinglePrice.Caption = FormatMoney(0) & "/" & FormatMoney(0)
        flblLimitedCount.Caption = ""
        flblLimitedTime.Caption = ""
        flblStandCount.Caption = ""
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
'    If KeyCode = vbKeyDown And Shift = 0 Then
'        If lvBus.SelectedItem Is Nothing Or lvBus.ListItems.count < 1 Then Exit Sub
'        If (lvBus.SelectedItem.Index = lvBus.ListItems.count - 2) Or (lvBus.SelectedItem.Index = lvBus.ListItems.count - 1) Or (lvBus.SelectedItem.Index = lvBus.ListItems.count) Then
'            RefreshNextScreen
'        End If
'    End If
'    If KeyCode = vbKeyPageDown And Shift = 0 Then
'        RefreshNextScreen
'    End If
'    If KeyCode = vbKeyEnd And Shift = 0 Then
'        RefreshAllScreen
'    End If
End Sub

Private Sub lvBus_LostFocus()
    lblBus.ForeColor = 0
    
End Sub


Private Sub lvPreSell_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    SortListView lvPreSell, ColumnHeader.Index
End Sub

Private Sub lvPreSell_GotFocus()
    lblPreSell.ForeColor = clActiveColor
End Sub
Private Sub lvPreSell_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Then
        If lvPreSell.ListItems.count <> 0 And Not lvPreSell.SelectedItem Is Nothing Then
            lvPreSell.ListItems.Remove lvPreSell.SelectedItem.Index
            DealPrice
            
        End If
    End If
    If KeyAscii = vbKeyReturn Then
        txtReceivedMoney.SetFocus
    End If
End Sub

Private Sub lvPreSell_LostFocus()
    lblPreSell.ForeColor = 0
End Sub

Private Sub mnu_changeseattype_Click()
    If cboSeatType.ListIndex = cboSeatType.ListCount - 1 Then
        cboSeatType.ListIndex = 0
    Else
        cboSeatType.ListIndex = cboSeatType.ListIndex + 1
    End If
End Sub

Private Sub lvSellStation_GotFocus()

   lblSellStation.ForeColor = clActiveColor
   cboSeatType_Change
   If lvSellStation.ListItems.count > 0 Then
        flblTotalPrice.Caption = FormatMoney(lvSellStation.SelectedItem.SubItems(2) + TotalInsurace)
        lvSellStation.Tag = lvSellStation.SelectedItem.Text
    '    lvSellStation.ListItems(nSellCount).Tag = lvSellStation.SelectedItem.Text
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
    lblSellStation.ForeColor = 0
End Sub

Private Sub Timer1_Timer()
'    RefreshBusSeats True
    On Error GoTo Here
    '隔40秒取一下服务器时间
    If m_nCount Mod 20 = 0 Then
        Date = m_oParam.NowDate
        Time = m_oParam.NowDateTime
        m_nCount = 0
    End If
    m_nCount = m_nCount + 1
    Exit Sub
    On Error GoTo 0
Here:
    ShowMsg err.Description
End Sub




Private Sub txtDiscount_GotFocus()
    lblDiscount.ForeColor = clActiveColor
End Sub

Private Sub txtDiscount_KeyPress(KeyAscii As Integer)
    If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = 8 Or (KeyAscii = 46 And InStr(txtDiscount.Text, ".") = 0) Then
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtDiscount_LostFocus()
On Error GoTo Here
    lblDiscount.ForeColor = 0
    If txtDiscount.Text <> "" Then
        If Trim(txtDiscount.Text) = "." Then
            txtDiscount.SetFocus
            Exit Sub
        End If
        If Left(txtDiscount.Text, 1) = "." And Len(txtDiscount.Text) > 1 Then
            txtDiscount.Text = "0" & txtDiscount.Text
        End If
        If CSng(txtDiscount.Text) > 1 Then
            MsgBox "折扣率不能大于 1", vbInformation, "提示"
            txtDiscount.SetFocus
        End If
    Else
        txtDiscount.Text = 1
    End If
    DealPrice
    
On Error GoTo 0
Exit Sub
Here:
    ShowErrorMsg
End Sub

Private Sub txtDiscount_Validate(Cancel As Boolean)
    If txtDiscount.Text = "" Then
        Cancel = True
    Else
        If CSng(txtDiscount.Text) > 1 Then
            Cancel = True
        End If
    End If
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
                    txtReceivedMoney.SetFocus
'                End If
            Else
                cmdPreSell_Click
                txtReceivedMoney.SetFocus
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
                    txtReceivedMoney.SetFocus
'                End If
            Else
                cmdPreSell_Click
                txtReceivedMoney.SetFocus
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
'    Select Case KeyCode
'        Case vbKeyLeft
'
'            cboPreferentialTicket.SetFocus
'        Case vbKeyUp
'            txtHalfSell.SetFocus
'    End Select
    
End Sub

Private Sub txtPreferentialSell_KeyPress(KeyAscii As Integer)
On Error GoTo Here
    If KeyAscii = 13 Then
        If txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text <> 0 And Not (lvBus.SelectedItem Is Nothing) Then
'            If lvBus.SelectedItem.SubItems(ID_OffTime) <> cszScrollBus Then
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
'                    cmdPreSell_Click
                   'cmdPreSell.SetFocus
    
'                      txtReceivedMoney.SetFocus
            
'                End If
'            Else
                cmdPreSell_Click
                txtReceivedMoney.SetFocus
'            End If
        End If
        
    End If
    
On Error GoTo 0
Exit Sub
Here:
    ShowErrorMsg
End Sub

Private Sub txtPreferentialSell_LostFocus()
'    fraPreferentialTicket.ForeColor = 0

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

End Sub

Private Sub txtHalfSell_GotFocus()
    lblHalfSell.ForeColor = clActiveColor
    txtHalfSell.SelStart = 0
    txtHalfSell.SelLength = 2
End Sub

Private Sub txtHalfSell_LostFocus()
    lblHalfSell.ForeColor = 0
 End Sub

Private Sub txtPrevDate_Change()
    On Error Resume Next
    
    If Val(txtPrevDate.Text) > m_nCanSellDay Then txtPrevDate.Text = m_nCanSellDay
    flblSellDate.Caption = ToStandardDateStr(DateAdd("d", txtPrevDate.Text, m_oParam.NowDate))
    
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
            Select Case Trim(m_aszSeatType(cboSeatType.ListIndex + 1, 1))
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
    RefreshBus True
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
    txtFullSell.Text = 0 '售全票张数为1
    txtHalfSell.Text = 0 '售半票张数为0
    txtPreferentialSell.Text = 1 '售免票张数为0
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
'                Set rsTemp = m_oSellDiscount.GetBusRs(CDate(flblSellDate.Caption), szStationID, , BusID)
'            Else
'                Set rsTemp = m_oSellDiscount.GetBusRsEx(CDate(flblSellDate.Caption), szStationID, m_szRegValue, , BusID)
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
        lvBus.SelectedItem.Tag = lvSellStation.SelectedItem.SubItems(3)
        lvBus.SelectedItem.SubItems(ID_OffTime) = lvSellStation.SelectedItem.SubItems(1)
        lvBus.SelectedItem.SubItems(ID_FullPrice) = lvSellStation.SelectedItem.SubItems(2)
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
'                Set rsTemp = m_oSellDiscount.GetBusRs(CDate(flblSellDate.Caption), szStationID, , BusID)
'            Else
'                Set rsTemp = m_oSellDiscount.GetBusRsEx(CDate(flblSellDate.Caption), szStationID, m_szRegValue, , BusID)
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

Private Sub RefreshBus(Optional pbForce As Boolean = False)
    Dim szStationID As String
'    Dim rsTemp As Recordset
    Dim liTemp As ListItem
    Dim lForeColor As OLE_COLOR
    Dim nBusType As EBusType
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim lvS As ListItem
    Dim szScrollBus As String
 
    
    On Error GoTo Here
    szStationID = RTrim(cboEndStation.BoundText)
    
    Set m_rsBusInfo = Nothing
    
    If cboEndStation.Changed Or pbForce Then
        
        If szStationID <> "" Then
            If m_szRegValue = "" Then
                Set rsCountTemp = m_oSellDiscount.GetBusRs(CDate(flblSellDate.Caption), szStationID, , , True)
            Else
                Set rsCountTemp = m_oSellDiscount.GetBusRsEx(CDate(flblSellDate.Caption), szStationID, m_szRegValue)
            End If
            
'            If pszSellStationID = "" Then
'              lvSellStation.ListItems.Clear
'            End If
            
            If rsCountTemp.RecordCount <> 0 Then
'                lblmileate = rsCountTemp!end_station_mileage & "公里"
            End If
            
            lvBus.Sorted = False
            lvBus.ListItems.Clear
            lvBus.Refresh
            For j = 1 To rsCountTemp.RecordCount
'                If lvBus.ListItems.count > 17 Then
'                    With lvBus
'                        If .ListItems.count > 0 Then
'                            .SortKey = MDISellTicket.GetSortKey() - 1
'                            .Sorted = True
'                            For i = 1 To .ListItems.count
'                            '如果车次不是停班而且(车次有座位或站票),则让该车次选中
'                                If .ListItems(i).ForeColor <> m_lStopBusColor And (.ListItems(i).SubItems(ID_SeatCount) > 0) Then
'
'                                    .ListItems(i).Selected = True
'                                    .ListItems(i).EnsureVisible
'                                    Exit For
'                                End If
'                            Next i
'                            If i > .ListItems.count Then
'                                .ListItems(1).Selected = True
'                                .ListItems(1).EnsureVisible
'                            End If
'                        End If
'                    End With
'                    rsCountTemp.MovePrevious
'                    Set m_rsBusInfo = rsCountTemp
''                    Set rsTemp = Nothing
'                    Set liTemp = Nothing
'                    Exit Sub
'                Else
                    If Hour(rsCountTemp!busstarttime) >= txtTime.Text Then
                        '如果车次时间大于查询的时间
                        For i = lvBus.ListItems.count To 1 Step -1
                            If RTrim(rsCountTemp!bus_id) = lvBus.ListItems(i) And Format(rsCountTemp!bus_date, "yyyy-mm-dd") = CDate(flblSellDate.Caption) Then
'                                Select Case Trim(rsCountTemp!seat_type_id)
'                                Case cszSeatType
'                                    liTemp.SubItems(ID_FullPrice) = rsCountTemp!full_price
'                                    liTemp.SubItems(ID_HalfPrice) = rsCountTemp!half_price
'                                    liTemp.SubItems(ID_PreferentialPrice1) = rsCountTemp!preferential_ticket1
'                                    liTemp.SubItems(ID_PreferentialPrice2) = rsCountTemp!preferential_ticket2
'                                    liTemp.SubItems(ID_PreferentialPrice3) = rsCountTemp!preferential_ticket3
'                                Case cszBedType
'                                    liTemp.SubItems(ID_BedFullPrice) = rsCountTemp!full_price
'                                    liTemp.SubItems(ID_BedHalfPrice) = rsCountTemp!half_price
'                                    liTemp.SubItems(ID_BedPreferentialPrice1) = rsCountTemp!preferential_ticket1
'                                    liTemp.SubItems(ID_BedPreferentialPrice2) = rsCountTemp!preferential_ticket2
'                                    liTemp.SubItems(ID_BedPreferentialPrice3) = rsCountTemp!preferential_ticket3
'                                Case cszAdditionalType
'                                    liTemp.SubItems(ID_AdditionalFullPrice) = rsCountTemp!full_price
'                                    liTemp.SubItems(ID_AdditionalHalfPrice) = rsCountTemp!half_price
'                                    liTemp.SubItems(ID_AdditionalPreferential1) = rsCountTemp!preferential_ticket1
'                                    liTemp.SubItems(ID_AdditionalPreferential2) = rsCountTemp!preferential_ticket2
'                                    liTemp.SubItems(ID_AdditionalPreferential3) = rsCountTemp!preferential_ticket3
'                                End Select
                                GoTo nextstep
                            End If
'                            Exit For
                        Next i
                        If rsCountTemp!status = ST_BusStopped Or rsCountTemp!status = ST_BusMergeStopped Or rsCountTemp!status = ST_BusSlitpStop Then
                            lForeColor = m_lStopBusColor
                        Else
                            lForeColor = m_lNormalBusColor
                            Set liTemp = lvBus.ListItems.Add(, "A" & RTrim(rsCountTemp!bus_id), RTrim(rsCountTemp!bus_id))
                            '车次代码"A" & RTrim(rsCountTemp！bus_id)
                        End If
                        nBusType = rsCountTemp!bus_type
                        If lForeColor <> m_lStopBusColor Then
                            liTemp.ForeColor = lForeColor
                            
'                         varBookmark = rsCountTemp.Bookmark
'                                If rsCountTemp.RecordCount > 0 Then
'                                   RefreshSellStation rsCountTemp
'                                End If
'                           rsCountTemp.Bookmark = varBookmark

                            If nBusType <> TP_ScrollBus Then
                                liTemp.SubItems(ID_BusType) = Trim(rsCountTemp!bus_type)
                                liTemp.SubItems(ID_OffTime) = Format(rsCountTemp!busstarttime, "hh:mm")
                            Else
                                liTemp.SubItems(ID_VehicleModel) = cszScrollBus
                                liTemp.SubItems(ID_OffTime) = cszScrollBus
                            End If
                            liTemp.SubItems(ID_RouteName) = Trim(rsCountTemp!route_name)
                            liTemp.SubItems(ID_EndStation) = RTrim(rsCountTemp!end_station_name)
                            liTemp.SubItems(ID_TotalSeat) = rsCountTemp!total_seat
                            If IsDate(liTemp.SubItems(ID_OffTime)) Then
                                If g_bIsBookValid And DateAdd("n", -g_nBookTime, liTemp.SubItems(ID_OffTime)) < Time And ToDBDate(flblSellDate.Caption) = ToDBDate(Date) Then
                                    '如果车次日期为当天,且已过预定时限,则将预定人数加到可售张数上面.
                                    liTemp.SubItems(ID_BookCount) = 0
                                    liTemp.SubItems(ID_SeatCount) = rsCountTemp!sale_seat_quantity + rsCountTemp!book_count
                                    
                                Else
                                    liTemp.SubItems(ID_BookCount) = rsCountTemp!book_count
                                    liTemp.SubItems(ID_SeatCount) = rsCountTemp!sale_seat_quantity
                                End If
                            Else
                            
                                liTemp.SubItems(ID_BookCount) = rsCountTemp!book_count
                                liTemp.SubItems(ID_SeatCount) = rsCountTemp!sale_seat_quantity
                            End If
                            liTemp.SubItems(ID_SeatTypeCount) = rsCountTemp!seat_remain
                            liTemp.SubItems(ID_BedTypeCount) = rsCountTemp!bed_remain
                            liTemp.SubItems(ID_AdditionalCount) = rsCountTemp!additional_remain
                            liTemp.SubItems(ID_VehicleModel) = rsCountTemp!vehicle_type_name
                            Select Case Trim(rsCountTemp!seat_type_id)
                            Case cszSeatType
                                liTemp.SubItems(ID_FullPrice) = rsCountTemp!full_price
                                liTemp.SubItems(ID_HalfPrice) = rsCountTemp!half_price
                                liTemp.SubItems(ID_FreePrice) = 0
                                liTemp.SubItems(ID_PreferentialPrice1) = rsCountTemp!preferential_ticket1
                                liTemp.SubItems(ID_PreferentialPrice2) = rsCountTemp!preferential_ticket2
                                liTemp.SubItems(ID_PreferentialPrice3) = rsCountTemp!preferential_ticket3
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
                                liTemp.SubItems(ID_BedFullPrice) = rsCountTemp!full_price
                                liTemp.SubItems(ID_BedHalfPrice) = rsCountTemp!half_price
                                liTemp.SubItems(ID_BedFreePrice) = 0
                                liTemp.SubItems(ID_BedPreferentialPrice1) = rsCountTemp!preferential_ticket1
                                liTemp.SubItems(ID_BedPreferentialPrice2) = rsCountTemp!preferential_ticket2
                                liTemp.SubItems(ID_BedPreferentialPrice3) = rsCountTemp!preferential_ticket3
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
                                liTemp.SubItems(ID_AdditionalFullPrice) = rsCountTemp!full_price
                                liTemp.SubItems(ID_AdditionalHalfPrice) = rsCountTemp!half_price
                                liTemp.SubItems(ID_AdditionalFreePrice) = 0
                                liTemp.SubItems(ID_AdditionalPreferential1) = rsCountTemp!preferential_ticket1
                                liTemp.SubItems(ID_AdditionalPreferential2) = rsCountTemp!preferential_ticket2
                                liTemp.SubItems(ID_AdditionalPreferential3) = rsCountTemp!preferential_ticket3
                            End Select
                            '以下几列不显示出来，只是将其存储，以备后用
                            liTemp.SubItems(ID_LimitedCount) = rsCountTemp!sale_ticket_quantity
                            liTemp.SubItems(ID_LimitedTime) = rsCountTemp!stop_sale_time
                            liTemp.SubItems(ID_BusType1) = nBusType
                            liTemp.SubItems(ID_CheckGate) = rsCountTemp!check_gate_id
                            liTemp.SubItems(ID_StandCount) = rsCountTemp!sale_stand_seat_quantity
                            liTemp.Tag = MakeDisplayString(Trim(rsCountTemp!sell_station_id), Trim(rsCountTemp!sell_station_name))

                        End If
                    End If
'                End If
                
                
nextstep:
                rsCountTemp.MoveNext
            Next j
            ' Loop
'    If lvBus.ListItems.count > 0 Then
'       RefreshSellStation lvBus.SelectedItem.Text
'    Else
'       lvSellStation.ListItems.Clear
'    End If
            lvBus.Sorted = True
        Else
            lvBus.ListItems.Clear
            
        End If
    End If


    
    '设定某个站点？车次显示在第一行
    With lvBus
        If .ListItems.count > 0 Then
'             szScrollBus = .ListItems(i).SubItems(ID_OffTime)
            .SortKey = MDISellTicket.GetSortKey() - 1
            .Sorted = True
            For i = 1 To .ListItems.count

                '如果车次不是停班而且(车次有座位或站票),则让该车次选中   是否为滚动车次
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
        '设定某个站点？车次显示在第一行
        '调用车次改变要进行相应操作的方法
    End With
    
   If lvBus.ListItems.count > 0 Then
      RefreshSellStation rsCountTemp
   Else
      lvSellStation.ListItems.Clear
   End If
    DoThingWhenBusChange
'    Set rsTemp = Nothing
    Set liTemp = Nothing
    On Error GoTo 0
    Exit Sub
Here:
    ShowErrorMsg
    Set rsCountTemp = Nothing
    Set liTemp = Nothing
End Sub

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
        cmdSell_Click
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


Private Sub RefreshStation2()
    Dim rsTemp As Recordset
    Dim szTemp As String
    On Error GoTo Here
'    szTemp = m_oSellDiscount.SellUnitCode
'    m_oSellDiscount.SellUnitCode = m_szCurrentUnitID
    Set rsTemp = m_oSellDiscount.GetAllStationRs()
'    m_oSellDiscount.SellUnitCode = szTemp

    With cboEndStation
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
    Exit Sub
Here:
    Set rsTemp = Nothing
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
    
'    szTemp = m_oSellDiscount.SellUnitCode
'    m_oSellDiscount.SellUnitCode = m_szCurrentUnitID

    '得到所有的票种
    atTicketType = m_oParam.GetAllTicketType()
    aszSeatType = m_oSellDiscount.GetAllSeatType
'    m_oSellDiscount.SellUnitCode = szTemp
    
    nCount = ArrayLength(atTicketType)
    nLen = ArrayLength(aszSeatType)
    
    
    sgWidth = 690
    lvBus.ColumnHeaders.Clear
    '添加ListView列头
    With lvBus.ColumnHeaders
        .Add , , "车次", 950 '"BusID"
        .Add , , "车次类型", 0 '"BusType"
        .Add , , "时间", 850 '"OffTime"
        .Add , , "线路名称", 1900
        .Add , , "终到站", 930 '"EndStation"
        .Add , , "总", 440
        .Add , , "订", 440
        .Add , , "座余", 700 '"SeatCount"
        .Add , , "座", 0
        .Add , , "卧", 0 '440
        .Add , , "加", 0 '440
          .Add , , "车型", 900 '"BusModel"
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
    
    
    '将组合框中的票种代码与票种名称放到数组m_atTicketType 中
    If nUsedPerential > 0 Then ReDim m_atTicketType(1 To nUsedPerential)
    j = 0
    For i = 1 To nCount
        If atTicketType(i).nTicketTypeID > TP_HalfPrice And atTicketType(i).nTicketTypeValid = TP_TicketTypeValid Then
            j = j + 1
            m_atTicketType(j) = atTicketType(i)
        End If
    Next i
    If cboPreferentialTicket.ListCount < 1 Then
        cboPreferentialTicket.Enabled = False
        txtPreferentialSell.Enabled = False
        cboPreferentialTicket.Text = ""
    Else
        cboPreferentialTicket.Enabled = False
        txtPreferentialSell.Enabled = True
        
        
        '设置折扣票的票种
        For i = 1 To cboPreferentialTicket.ListCount
            If m_atTicketType(i).nTicketTypeID = g_nDiscountTicketInTicketTypePosition Then
                cboPreferentialTicket.ListIndex = i - 1
                Exit For
            End If
        Next i
        If i > cboPreferentialTicket.ListCount Then
            cboPreferentialTicket.ListIndex = 0
        End If
    End If
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
Private Sub GetPreSellBus() '设置预售票状态信息
Dim i As Integer
Dim nCount As Integer
Dim atTicketType() As TTicketType

atTicketType = m_oSellDiscount.GetAllTicketType()
    nCount = ArrayLength(atTicketType)
    With lvPreSell.ColumnHeaders
        .Add , , "车次", 950   '"BusID"
        .Add , , "车次类型", 0 '"BusType"
        .Add , , "时间", 900 '"OffTime"
        .Add , , "乘车日期", 1450 '"BusDate"
        .Add , , "起始站", 0  '"StartStation"
        .Add , , "终到站", 899 '"EndStation"
        .Add , , "车型", 899  '"VehicleModel"
        .Add , , "总票数", 899  '"SumTicketNum"
        .Add , , "总票价", 0  '"SumPrice"
        .Add , , "定座", 0   '"OrderSeat"
        For i = 1 To nCount
            .Add , , Trim(atTicketType(i).szTicketTypeName) & "票价", 0  '票种
            .Add , , Trim(atTicketType(i).szTicketTypeName), 900
        Next i
        .Add , , "折扣票价", 0  'DiscountPrice
        .Add , , "折扣率", 0   '"Discount"
        .Add , , "站票", 0     '"StandCount"
        .Add , , "检票口", 700   '"CheckGate"
        .Add , , "限售张数", 0 '"LimitedCount"
        .Add , , "终到站代码", 0 '"BoundText"
        .Add , , "座位状态1", 0 '"SetSeatEnable"
        .Add , , "座位状态2", 0 '"SetSeatValue"
        .Add , , "座位号", 0  '"SeatNo"
        .Add , , "票价明细", 0 '"TicketPrice"
        .Add , , "票钟明细", 0 ' "TicketType"
        .Add , , "座位类型", 0
        .Add , , "终点站", 0
    End With
End Sub
Private Sub GetPreSellTicketInfo()  '得到预售暂放票的信息
    Dim liPreSell As ListItem
    Dim liSelected As ListItem
    Dim i As Integer
    Dim szPrice As String
    Dim szTicketType As String
    On Error GoTo Here
    If lvBus.ListItems.count <> 0 Then
        Set liSelected = lvBus.SelectedItem
        Set liPreSell = lvPreSell.ListItems.Add(, , liSelected.Text)
        With liPreSell
            .Tag = lvBus.SelectedItem.Tag
            .SubItems(IT_BusType) = liSelected.SubItems(ID_BusType)
            .SubItems(IT_OffTime) = liSelected.SubItems(ID_OffTime)
            .SubItems(IT_BusDate) = flblSellDate.Caption
            .SubItems(IT_StartStation) = lvSellStation.SelectedItem.Text
            .SubItems(IT_EndStation) = GetStationNameInCbo(cboEndStation.Text)
            .SubItems(IT_VehicleModel) = liSelected.SubItems(ID_VehicleModel)
            .SubItems(IT_SumTicketNum) = txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text
            
            .SubItems(IT_OrderSeat) = frmOrderSeats.m_szBookNumber
            Select Case Trim(m_aszSeatType(cboSeatType.ListIndex + 1, 1))
                Case cszSeatType
                    .SubItems(IT_FullPrice) = liSelected.SubItems(ID_FullPrice)
                    
                    .SubItems(IT_FullNum) = txtFullSell.Text & " 座"
                    
                    .SubItems(IT_HalfPrice) = liSelected.SubItems(ID_HalfPrice)
                    .SubItems(IT_HalfNum) = txtHalfSell.Text & " 座"
                    

                    If cboPreferentialTicket.ListCount > 0 Then
                        Select Case m_atTicketType(cboPreferentialTicket.ListIndex + 1).nTicketTypeID
                        Case TP_FreeTicket
                            .SubItems(IT_FreeType) = liSelected.SubItems(ID_FreePrice)
                            .SubItems(IT_FreeNum) = txtPreferentialSell.Text & " 座"
                        Case TP_PreferentialTicket1
                            .SubItems(IT_PreferentialType1) = liSelected.SubItems(ID_PreferentialPrice1)
                            .SubItems(IT_PreferentialNum1) = txtPreferentialSell.Text
                        Case TP_PreferentialTicket2
                            .SubItems(IT_PreferentialType2) = liSelected.SubItems(ID_PreferentialPrice2)
                            .SubItems(IT_PreferentialNum2) = txtPreferentialSell.Text & " 座"
                        Case TP_PreferentialTicket3
                            .SubItems(IT_PreferentialType3) = liSelected.SubItems(ID_PreferentialPrice3)
                            .SubItems(IT_PreferentialNum3) = txtPreferentialSell.Text & " 座"
                        End Select
                    End If
                Case cszBedType
                    .SubItems(IT_FullPrice) = liSelected.SubItems(ID_BedFullPrice)
                    .SubItems(IT_FullNum) = txtFullSell.Text & " 卧"
                    .SubItems(IT_HalfPrice) = liSelected.SubItems(ID_BedHalfPrice)
                    .SubItems(IT_HalfNum) = txtHalfSell.Text & " 卧"
                    If cboPreferentialTicket.ListCount > 0 Then
                        Select Case m_atTicketType(cboPreferentialTicket.ListIndex + 1).nTicketTypeID
                        Case TP_FreeTicket
                            .SubItems(IT_FreeType) = liSelected.SubItems(ID_BedFreePrice)
                            .SubItems(IT_FreeNum) = txtPreferentialSell.Text & " 卧"
                        Case TP_PreferentialTicket1
                            .SubItems(IT_PreferentialType1) = liSelected.SubItems(ID_BedPreferentialPrice1)
                            .SubItems(IT_PreferentialNum1) = txtPreferentialSell.Text & " 卧"
                        Case TP_PreferentialTicket2
                            .SubItems(IT_PreferentialType2) = liSelected.SubItems(ID_BedPreferentialPrice2)
                            .SubItems(IT_PreferentialNum2) = txtPreferentialSell.Text & " 卧"
                        Case TP_PreferentialTicket3
                            .SubItems(IT_PreferentialType3) = liSelected.SubItems(ID_BedPreferentialPrice3)
                            .SubItems(IT_PreferentialNum3) = txtPreferentialSell.Text & " 卧"
                        End Select
                    End If
                Case cszAdditionalType
                    .SubItems(IT_FullPrice) = liSelected.SubItems(ID_AdditionalFullPrice)
                    .SubItems(IT_FullNum) = txtFullSell.Text & " 加"
                    .SubItems(IT_HalfPrice) = liSelected.SubItems(ID_AdditionalHalfPrice)
                    .SubItems(IT_HalfNum) = txtHalfSell.Text & " 加"
                    If cboPreferentialTicket.ListCount > 0 Then
                        Select Case m_atTicketType(cboPreferentialTicket.ListIndex + 1).nTicketTypeID
                        Case TP_FreeTicket
                            .SubItems(IT_FreeType) = liSelected.SubItems(ID_AdditionalFreePrice)
                            .SubItems(IT_FreeNum) = txtPreferentialSell.Text & " 加"
                        Case TP_PreferentialTicket1
                            .SubItems(IT_PreferentialType1) = liSelected.SubItems(ID_AdditionalPreferential1)
                            .SubItems(IT_PreferentialNum1) = txtPreferentialSell.Text & " 加"
                        Case TP_PreferentialTicket2
                            .SubItems(IT_PreferentialType2) = liSelected.SubItems(ID_AdditionalPreferential2)
                            .SubItems(IT_PreferentialNum2) = txtPreferentialSell.Text & " 加"
                        Case TP_PreferentialTicket3
                            .SubItems(IT_PreferentialType3) = liSelected.SubItems(ID_AdditionalPreferential3)
                            .SubItems(IT_PreferentialNum3) = txtPreferentialSell.Text & " 加"
                        End Select
                    End If
            End Select
            .SubItems(IT_SumPrice) = Val(.SubItems(IT_FullNum)) * Val(.SubItems(IT_FullPrice)) + _
                                                Val(.SubItems(IT_HalfNum)) * Val(.SubItems(IT_HalfPrice)) + _
                                                Val(.SubItems(IT_PreferentialNum1)) * Val(.SubItems(IT_PreferentialType1)) + _
                                                Val(.SubItems(IT_PreferentialNum2)) * Val(.SubItems(IT_PreferentialType2)) + _
                                                Val(.SubItems(IT_PreferentialNum3)) * Val(.SubItems(IT_PreferentialType3))
'            .SubItems(IT_DiscountPrice) = CDbl(txtDiscount.Text) * CDbl(liSelected.subitems(ID_FullPrice))
            .SubItems(IT_Discount) = txtDiscount.Text
            '.SubItems(IT_StandCount) = liSelected.subitems(ID_StandCount)
            .SubItems(IT_CheckGate) = liSelected.SubItems(ID_CheckGate)
            .SubItems(IT_LimitedCount) = liSelected.SubItems(ID_LimitedCount)
            .SubItems(IT_BoundText) = cboEndStation.BoundText
            .SubItems(IT_SetSeatEnable) = chkSetSeat.Enabled
            .SubItems(IT_SetSeatValue) = chkSetSeat.Value
            .SubItems(IT_SeatNo) = txtSeat.Text
            .SubItems(IT_SeatType) = m_aszSeatType(cboSeatType.ListIndex + 1, 1)
            .SubItems(IT_TerminateName) = liSelected.SubItems(ID_EndStation)
        End With
    End If
    Set liPreSell = Nothing
    Set liSelected = Nothing
On Error GoTo 0
Exit Sub
Here:
     Set liPreSell = Nothing
    Set liSelected = Nothing
    ShowErrorMsg
End Sub
Private Function GetDealTotalPrice() As Double  '得到总票价
    Dim iCount As Integer
    Dim dbTotal As Double
    dbTotal = 0
    If lvPreSell.ListItems.count <> 0 Then
        For iCount = 1 To lvPreSell.ListItems.count
            dbTotal = dbTotal + Val(lvPreSell.ListItems(iCount).SubItems(IT_SumPrice))
        Next iCount
    End If
    GetDealTotalPrice = dbTotal
End Function
'///////////////////////////////////
'设置预售按钮状态
Private Sub SetPreSellButton()
    If txtFullSell.Text = 0 And txtHalfSell.Text = 0 And txtPreferentialSell.Text = 0 Then
        cmdPreSell.Enabled = False
    Else
        cmdPreSell.Enabled = True
    End If
End Sub
'/////////////////////////////////
'处理折扣票与定座
Private Sub DealDiscountAndSeat()
   '判断是否有售折扣票权限
   On Error GoTo Here
   If m_oSellDiscount.DiscountIsValid Then
        txtDiscount.Enabled = False
        fraDiscountTicket.Enabled = False
   End If
   If m_oSellDiscount.OrderSeatIsValid Then
        chkSetSeat.Value = 0
        chkSetSeat.Visible = False
        lblSetSeat.Enabled = False
   End If
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
If lvPreSell.ListItems.count <> 0 Then
    For i = 1 To lvPreSell.ListItems.count
        Set liTemp = lvPreSell.ListItems(i)
        If CDate(flblSellDate.Caption) = CDate(liTemp.SubItems(IT_BusDate)) And lvBus.SelectedItem.Text = liTemp.Text Then
            If liTemp.SubItems(IT_SeatNo) <> "" Then
                szTemp = szTemp & "," & liTemp.SubItems(IT_SeatNo)
            End If
        End If
    Next i
Else
    szTemp = ""
End If
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
If lvPreSell.ListItems.count <> 0 And (Not lvBus.SelectedItem Is Nothing) Then
    Set liSelected = lvBus.SelectedItem
    For i = 1 To lvPreSell.ListItems.count
        Set liTemp = lvPreSell.ListItems(i)
        If liTemp.Text = liSelected.Text And liTemp.SubItems(IT_BusDate) = CDate(flblSellDate.Caption) And liTemp.SubItems(IT_BoundText) = cboEndStation.BoundText Then
            GetSameBusIndex = i
            Exit Function
        End If
    Next i
End If
GetSameBusIndex = 0
End Function
'/////////////////////////////////////
'合并相同车次信息
Private Sub MergeSameBusInfo(nSameIndex As Integer)
Dim liTemp As ListItem
Set liTemp = lvPreSell.ListItems(nSameIndex)
Dim szPrice As String
Dim szTicketType As String
Dim i As Integer
Dim sgTemp As Single
sgTemp = 0
With liTemp

    .SubItems(IT_SumTicketNum) = Val(.SubItems(IT_SumTicketNum)) + txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text
    .SubItems(IT_FullNum) = Val(.SubItems(IT_FullNum)) + txtFullSell.Text
    .SubItems(IT_HalfNum) = Val(.SubItems(IT_HalfNum)) + txtHalfSell.Text
    .SubItems(IT_SeatNo) = Trim(.SubItems(IT_SeatNo)) & "," & Trim(txtSeat.Text)
    If cboPreferentialTicket.ListCount > 0 Then
        Select Case m_atTicketType(cboPreferentialTicket.ListIndex + 1).nTicketTypeID
        Case TP_FreeTicket
            .SubItems(IT_FreeNum) = txtPreferentialSell.Text + Val(.SubItems(IT_FreeNum))
        Case TP_PreferentialTicket1
            .SubItems(IT_PreferentialNum1) = txtPreferentialSell.Text + Val(.SubItems(IT_PreferentialNum1))
        Case TP_PreferentialTicket2
            .SubItems(IT_PreferentialNum2) = txtPreferentialSell.Text + Val(.SubItems(IT_PreferentialNum2))
        Case TP_PreferentialTicket3
            .SubItems(IT_PreferentialNum3) = txtPreferentialSell.Text + Val(.SubItems(IT_PreferentialNum3))
        End Select
    End If
 
    .SubItems(IT_SumPrice) = Val(.SubItems(IT_SumPrice)) + _
                                txtFullSell.Text * lvBus.SelectedItem.SubItems(ID_FullPrice) + _
                                txtHalfSell.Text * lvBus.SelectedItem.SubItems(ID_HalfPrice) + _
                                txtPreferentialSell.Text * GetPreferentialPrice
                                
 
End With
End Sub


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
''初始化站点车次顺序数组
Private Sub InitScrollBusOrder()
    Dim i As Integer
    Dim nCurLen As Integer

    nCurLen = ArrayLength(m_atbBusOrder)
    If nCurLen = 0 Then
        ReDim m_atbBusOrder(1 To 1)
    Else
        ReDim Preserve m_atbBusOrder(1 To nCurLen + 1)
    End If

    m_atbBusOrder(nCurLen + 1).szStatioinID = Trim(cboEndStation.BoundText)
    m_atbBusOrder(nCurLen + 1).dbCount = 1
End Sub
'判断滚动站点是否存在于数组当中
Private Function IsExitInTeam(pszStationID As String) As Integer
    Dim i As Integer
    Dim nLen As Integer
    nLen = ArrayLength(m_atbBusOrder)
    If nLen > 300 Then
        ReDim m_atbBusOrder(1 To 1)
        m_atbBusOrder(1).szStatioinID = ""
        m_atbBusOrder(1).dbCount = 1
        IsExitInTeam = 0
        Exit Function
    End If
    For i = 1 To nLen
        If pszStationID = m_atbBusOrder(i).szStatioinID Then
            IsExitInTeam = i
            Exit Function
        End If
    Next i
    IsExitInTeam = 0
End Function
'给数组顺序最小索引加值
Private Sub AddValueToIndex(pnIndex As Integer)
    If m_atbBusOrder(pnIndex).dbCount > 1000 Then
        m_atbBusOrder(pnIndex).dbCount = 1
    Else
        m_atbBusOrder(pnIndex).dbCount = m_atbBusOrder(pnIndex).dbCount + 1
    End If
End Sub
'lvBus中显示正确的车次顺序
Private Sub SetCorrectBusOrder(pszStationID As String)
    Dim nIndex As Integer
    Dim dbTemp As Double
    Dim aszSaveTemp() As String
    Dim j As Integer
    Dim liTemp As ListItem
    Dim nCount As Integer
    nIndex = IsExitInTeam(pszStationID)
    If lvBus.ListItems.count <> 0 Then
        nCount = (m_atbBusOrder(nIndex).dbCount Mod lvBus.ListItems.count) + 1
        
        ReDim aszSaveTemp(1 To lvBus.ListItems(nCount).ListSubItems.count)
        aszSaveTemp(1) = lvBus.ListItems(nCount)
        For j = 2 To lvBus.ListItems(nCount).ListSubItems.count
            aszSaveTemp(j) = lvBus.ListItems(nCount).SubItems(j - 1)
        Next j
        lvBus.ListItems.Remove nCount
        
        Set liTemp = lvBus.ListItems.Add(1, , aszSaveTemp(1))
        For j = 1 To ArrayLength(aszSaveTemp) - 1
            liTemp.ListSubItems.Add , , aszSaveTemp(j + 1)
       
        Next j
        liTemp.Selected = True
    End If
   If lvBus.ListItems.count > 0 Then
     RefreshSellStation rsCountTemp
   Else
     lvSellStation.ListItems.Clear
   End If
    Set liTemp = Nothing
End Sub

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
'    Dim i As Integer
'    Dim j As Integer
'    Dim liTemp As ListItem
'    Dim lForeColor As OLE_COLOR
'    Dim nBusType As EBusType
'    j = 0
'    If m_rsBusInfo Is Nothing Then Exit Sub
'    If Not m_rsBusInfo.EOF Then m_rsBusInfo.MoveNext
'    Do While Not m_rsBusInfo.EOF
'       j = j + 1
'       If j > 17 Then
'         m_rsBusInfo.MovePrevious
'         Exit Sub
'       End If
'       For i = lvBus.ListItems.count To 1 Step -1
'
'            If RTrim(m_rsBusInfo!bus_id) = lvBus.ListItems(i) And Format(m_rsBusInfo!bus_date, "yyyy-mm-dd") = CDate(flblSellDate.Caption) Then
'                If liTemp Is Nothing Then Set liTemp = lvBus.ListItems(i)
'                Select Case Trim(m_rsBusInfo!seat_type_id)
'                    Case cszSeatType
'                        liTemp.SubItems(ID_FullPrice) = m_rsBusInfo!full_price
'                        liTemp.SubItems(ID_HalfPrice) = m_rsBusInfo!half_price
'                        liTemp.SubItems(ID_PreferentialPrice1) = m_rsBusInfo!preferential_ticket1
'                        liTemp.SubItems(ID_PreferentialPrice2) = m_rsBusInfo!preferential_ticket2
'                        liTemp.SubItems(ID_PreferentialPrice3) = m_rsBusInfo!preferential_ticket3
'                    Case cszBedType
'                        liTemp.SubItems(ID_BedFullPrice) = m_rsBusInfo!full_price
'                        liTemp.SubItems(ID_BedHalfPrice) = m_rsBusInfo!half_price
'                        liTemp.SubItems(ID_BedPreferentialPrice1) = m_rsBusInfo!preferential_ticket1
'                        liTemp.SubItems(ID_BedPreferentialPrice2) = m_rsBusInfo!preferential_ticket2
'                        liTemp.SubItems(ID_BedPreferentialPrice3) = m_rsBusInfo!preferential_ticket3
'                    Case cszAdditionalType
'                        liTemp.SubItems(ID_AdditionalFullPrice) = m_rsBusInfo!full_price
'                        liTemp.SubItems(ID_AdditionalHalfPrice) = m_rsBusInfo!half_price
'                        liTemp.SubItems(ID_AdditionalPreferential1) = m_rsBusInfo!preferential_ticket1
'                        liTemp.SubItems(ID_AdditionalPreferential2) = m_rsBusInfo!preferential_ticket2
'                        liTemp.SubItems(ID_AdditionalPreferential3) = m_rsBusInfo!preferential_ticket3
'
'                End Select
'                GoTo nextstep
'            End If
'            Exit For
'        Next i
'        If m_rsBusInfo!status = ST_BusStopped Or m_rsBusInfo!status = ST_BusMergeStopped Or m_rsBusInfo!status = ST_BusSlitpStop Then
'            lForeColor = m_lStopBusColor
'
'        Else
'            lForeColor = m_lNormalBusColor
'            Set liTemp = lvBus.ListItems.Add(, "A" & RTrim(m_rsBusInfo!bus_id), RTrim(m_rsBusInfo!bus_id))   '车次代码"A" & RTrim(m_rsBusInfo！bus_id)
'        End If
'
'        nBusType = m_rsBusInfo!bus_type
'
'
'        If lForeColor <> m_lStopBusColor Then
'            liTemp.ForeColor = lForeColor
'            If nBusType <> TP_ScrollBus Then
'                liTemp.SubItems(ID_BusType) = Trim(m_rsBusInfo!bus_type)
'                liTemp.SubItems(ID_OffTime) = Format(m_rsBusInfo!busstarttime, "hh:mm")
'
'            Else
'                liTemp.SubItems(ID_VehicleModel) = cszScrollBus
'                liTemp.SubItems(ID_OffTime) = cszScrollBus
'
'            End If
'            liTemp.SubItems(ID_RouteName) = Trim(m_rsBusInfo!route_name)
'            liTemp.SubItems(ID_EndStation) = RTrim(m_rsBusInfo!end_station_name)
'            liTemp.SubItems(ID_TotalSeat) = m_rsBusInfo!total_seat
'            liTemp.SubItems(ID_SeatCount) = m_rsBusInfo!sale_seat_quantity
'            liTemp.SubItems(ID_SeatTypeCount) = m_rsBusInfo!seat_remain
'            liTemp.SubItems(ID_BedTypeCount) = m_rsBusInfo!bed_remain
'            liTemp.SubItems(ID_AdditionalCount) = m_rsBusInfo!additional_remain
'            liTemp.SubItems(ID_VehicleModel) = m_rsBusInfo!vehicle_type_name
'
'            Select Case Trim(m_rsBusInfo!seat_type_id)
'
'                Case cszSeatType
'                    liTemp.SubItems(ID_FullPrice) = m_rsBusInfo!full_price
'                    liTemp.SubItems(ID_HalfPrice) = m_rsBusInfo!half_price
'                    liTemp.SubItems(ID_FreePrice) = 0
'                    liTemp.SubItems(ID_PreferentialPrice1) = m_rsBusInfo!preferential_ticket1
'                    liTemp.SubItems(ID_PreferentialPrice2) = m_rsBusInfo!preferential_ticket2
'                    liTemp.SubItems(ID_PreferentialPrice3) = m_rsBusInfo!preferential_ticket3
'                    liTemp.SubItems(ID_BedFullPrice) = 0
'                    liTemp.SubItems(ID_BedHalfPrice) = 0
'                    liTemp.SubItems(ID_BedFreePrice) = 0
'                    liTemp.SubItems(ID_BedPreferentialPrice1) = 0
'                    liTemp.SubItems(ID_BedPreferentialPrice2) = 0
'                    liTemp.SubItems(ID_BedPreferentialPrice3) = 0
'                    liTemp.SubItems(ID_AdditionalFullPrice) = 0
'                    liTemp.SubItems(ID_AdditionalHalfPrice) = 0
'                    liTemp.SubItems(ID_AdditionalFreePrice) = 0
'                    liTemp.SubItems(ID_AdditionalPreferential1) = 0
'                    liTemp.SubItems(ID_AdditionalPreferential2) = 0
'                    liTemp.SubItems(ID_AdditionalPreferential3) = 0
'                Case cszBedType
'                    liTemp.SubItems(ID_FullPrice) = 0
'                    liTemp.SubItems(ID_HalfPrice) = 0
'                    liTemp.SubItems(ID_FreePrice) = 0
'                    liTemp.SubItems(ID_PreferentialPrice1) = 0
'                    liTemp.SubItems(ID_PreferentialPrice2) = 0
'                    liTemp.SubItems(ID_PreferentialPrice3) = 0
'                    liTemp.SubItems(ID_BedFullPrice) = m_rsBusInfo!full_price
'                    liTemp.SubItems(ID_BedHalfPrice) = m_rsBusInfo!half_price
'                    liTemp.SubItems(ID_BedFreePrice) = 0
'                    liTemp.SubItems(ID_BedPreferentialPrice1) = m_rsBusInfo!preferential_ticket1
'                    liTemp.SubItems(ID_BedPreferentialPrice2) = m_rsBusInfo!preferential_ticket2
'                    liTemp.SubItems(ID_BedPreferentialPrice3) = m_rsBusInfo!preferential_ticket3
'                    liTemp.SubItems(ID_AdditionalFullPrice) = 0
'                    liTemp.SubItems(ID_AdditionalHalfPrice) = 0
'                    liTemp.SubItems(ID_AdditionalFreePrice) = 0
'                    liTemp.SubItems(ID_AdditionalPreferential1) = 0
'                    liTemp.SubItems(ID_AdditionalPreferential2) = 0
'                    liTemp.SubItems(ID_AdditionalPreferential3) = 0
'                Case cszAdditionalType
'                    liTemp.SubItems(ID_FullPrice) = 0
'                    liTemp.SubItems(ID_HalfPrice) = 0
'                    liTemp.SubItems(ID_FreePrice) = 0
'                    liTemp.SubItems(ID_PreferentialPrice1) = 0
'                    liTemp.SubItems(ID_PreferentialPrice2) = 0
'                    liTemp.SubItems(ID_PreferentialPrice3) = 0
'                    liTemp.SubItems(ID_BedFullPrice) = 0
'                    liTemp.SubItems(ID_BedHalfPrice) = 0
'                    liTemp.SubItems(ID_BedFreePrice) = 0
'                    liTemp.SubItems(ID_BedPreferentialPrice1) = 0
'                    liTemp.SubItems(ID_BedPreferentialPrice2) = 0
'                    liTemp.SubItems(ID_BedPreferentialPrice3) = 0
'                    liTemp.SubItems(ID_AdditionalFullPrice) = m_rsBusInfo!full_price
'                    liTemp.SubItems(ID_AdditionalHalfPrice) = m_rsBusInfo!half_price
'                    liTemp.SubItems(ID_AdditionalFreePrice) = 0
'                    liTemp.SubItems(ID_AdditionalPreferential1) = m_rsBusInfo!preferential_ticket1
'                    liTemp.SubItems(ID_AdditionalPreferential2) = m_rsBusInfo!preferential_ticket2
'                    liTemp.SubItems(ID_AdditionalPreferential3) = m_rsBusInfo!preferential_ticket3
'
'            End Select
'
'            '以下几列不显示出来，只是将其存储，以备后用
'            liTemp.SubItems(ID_LimitedCount) = m_rsBusInfo!sale_ticket_quantity
'            liTemp.SubItems(ID_LimitedTime) = m_rsBusInfo!stop_sale_time
'            liTemp.SubItems(ID_BusType1) = nBusType
'            liTemp.SubItems(ID_CheckGate) = m_rsBusInfo!check_gate_id
'            liTemp.SubItems(ID_StandCount) = m_rsBusInfo!sale_stand_seat_quantity
'
'        End If
'nextstep:
'        m_rsBusInfo.MoveNext
'    Loop
End Sub
'显示下一屏
Private Sub RefreshAllScreen()
'    Dim i As Integer
'    Dim j As Integer
'    Dim liTemp As ListItem
'    Dim lForeColor As OLE_COLOR
'    Dim nBusType As EBusType
'
'    If m_rsBusInfo Is Nothing Then Exit Sub
'    If Not m_rsBusInfo.EOF Then m_rsBusInfo.MoveNext
'    Do While Not m_rsBusInfo.EOF
'
'       For i = lvBus.ListItems.count To 1 Step -1
'
'            If RTrim(m_rsBusInfo!bus_id) = lvBus.ListItems(i) And Format(m_rsBusInfo!bus_date, "yyyy-mm-dd") = CDate(flblSellDate.Caption) Then
'                If liTemp Is Nothing Then Set liTemp = lvBus.ListItems(i)
'                Select Case Trim(m_rsBusInfo!seat_type_id)
'                    Case cszSeatType
'                        liTemp.SubItems(ID_FullPrice) = m_rsBusInfo!full_price
'                        liTemp.SubItems(ID_HalfPrice) = m_rsBusInfo!half_price
'                        liTemp.SubItems(ID_PreferentialPrice1) = m_rsBusInfo!preferential_ticket1
'                        liTemp.SubItems(ID_PreferentialPrice2) = m_rsBusInfo!preferential_ticket2
'                        liTemp.SubItems(ID_PreferentialPrice3) = m_rsBusInfo!preferential_ticket3
'                    Case cszBedType
'                        liTemp.SubItems(ID_BedFullPrice) = m_rsBusInfo!full_price
'                        liTemp.SubItems(ID_BedHalfPrice) = m_rsBusInfo!half_price
'                        liTemp.SubItems(ID_BedPreferentialPrice1) = m_rsBusInfo!preferential_ticket1
'                        liTemp.SubItems(ID_BedPreferentialPrice2) = m_rsBusInfo!preferential_ticket2
'                        liTemp.SubItems(ID_BedPreferentialPrice3) = m_rsBusInfo!preferential_ticket3
'                    Case cszAdditionalType
'                        liTemp.SubItems(ID_AdditionalFullPrice) = m_rsBusInfo!full_price
'                        liTemp.SubItems(ID_AdditionalHalfPrice) = m_rsBusInfo!half_price
'                        liTemp.SubItems(ID_AdditionalPreferential1) = m_rsBusInfo!preferential_ticket1
'                        liTemp.SubItems(ID_AdditionalPreferential2) = m_rsBusInfo!preferential_ticket2
'                        liTemp.SubItems(ID_AdditionalPreferential3) = m_rsBusInfo!preferential_ticket3
'
'                End Select
'                GoTo nextstep
'            End If
'            Exit For
'        Next i
'        If m_rsBusInfo!status = ST_BusStopped Or m_rsBusInfo!status = ST_BusMergeStopped Or m_rsBusInfo!status = ST_BusSlitpStop Then
'            lForeColor = m_lStopBusColor
'
'        Else
'            lForeColor = m_lNormalBusColor
'            Set liTemp = lvBus.ListItems.Add(, "A" & RTrim(m_rsBusInfo!bus_id), RTrim(m_rsBusInfo!bus_id))   '车次代码"A" & RTrim(m_rsBusInfo！bus_id)
'        End If
'
'        nBusType = m_rsBusInfo!bus_type
'
'
'        If lForeColor <> m_lStopBusColor Then
'            liTemp.ForeColor = lForeColor
'            If nBusType <> TP_ScrollBus Then
'                liTemp.SubItems(ID_BusType) = Trim(m_rsBusInfo!bus_type)
'                liTemp.SubItems(ID_OffTime) = Format(m_rsBusInfo!busstarttime, "hh:mm")
'
'            Else
'                liTemp.SubItems(ID_VehicleModel) = cszScrollBus
'                liTemp.SubItems(ID_OffTime) = cszScrollBus
'
'            End If
'            liTemp.SubItems(ID_RouteName) = Trim(m_rsBusInfo!route_name)
'            liTemp.SubItems(ID_EndStation) = RTrim(m_rsBusInfo!end_station_name)
'            liTemp.SubItems(ID_TotalSeat) = m_rsBusInfo!total_seat
'            liTemp.SubItems(ID_SeatCount) = m_rsBusInfo!sale_seat_quantity
'            liTemp.SubItems(ID_SeatTypeCount) = m_rsBusInfo!seat_remain
'            liTemp.SubItems(ID_BedTypeCount) = m_rsBusInfo!bed_remain
'            liTemp.SubItems(ID_AdditionalCount) = m_rsBusInfo!additional_remain
'            liTemp.SubItems(ID_VehicleModel) = m_rsBusInfo!vehicle_type_name
'
'            Select Case Trim(m_rsBusInfo!seat_type_id)
'
'                Case cszSeatType
'                    liTemp.SubItems(ID_FullPrice) = m_rsBusInfo!full_price
'                    liTemp.SubItems(ID_HalfPrice) = m_rsBusInfo!half_price
'                    liTemp.SubItems(ID_FreePrice) = 0
'                    liTemp.SubItems(ID_PreferentialPrice1) = m_rsBusInfo!preferential_ticket1
'                    liTemp.SubItems(ID_PreferentialPrice2) = m_rsBusInfo!preferential_ticket2
'                    liTemp.SubItems(ID_PreferentialPrice3) = m_rsBusInfo!preferential_ticket3
'                    liTemp.SubItems(ID_BedFullPrice) = 0
'                    liTemp.SubItems(ID_BedHalfPrice) = 0
'                    liTemp.SubItems(ID_BedFreePrice) = 0
'                    liTemp.SubItems(ID_BedPreferentialPrice1) = 0
'                    liTemp.SubItems(ID_BedPreferentialPrice2) = 0
'                    liTemp.SubItems(ID_BedPreferentialPrice3) = 0
'                    liTemp.SubItems(ID_AdditionalFullPrice) = 0
'                    liTemp.SubItems(ID_AdditionalHalfPrice) = 0
'                    liTemp.SubItems(ID_AdditionalFreePrice) = 0
'                    liTemp.SubItems(ID_AdditionalPreferential1) = 0
'                    liTemp.SubItems(ID_AdditionalPreferential2) = 0
'                    liTemp.SubItems(ID_AdditionalPreferential3) = 0
'                Case cszBedType
'                    liTemp.SubItems(ID_FullPrice) = 0
'                    liTemp.SubItems(ID_HalfPrice) = 0
'                    liTemp.SubItems(ID_FreePrice) = 0
'                    liTemp.SubItems(ID_PreferentialPrice1) = 0
'                    liTemp.SubItems(ID_PreferentialPrice2) = 0
'                    liTemp.SubItems(ID_PreferentialPrice3) = 0
'                    liTemp.SubItems(ID_BedFullPrice) = m_rsBusInfo!full_price
'                    liTemp.SubItems(ID_BedHalfPrice) = m_rsBusInfo!half_price
'                    liTemp.SubItems(ID_BedFreePrice) = 0
'                    liTemp.SubItems(ID_BedPreferentialPrice1) = m_rsBusInfo!preferential_ticket1
'                    liTemp.SubItems(ID_BedPreferentialPrice2) = m_rsBusInfo!preferential_ticket2
'                    liTemp.SubItems(ID_BedPreferentialPrice3) = m_rsBusInfo!preferential_ticket3
'                    liTemp.SubItems(ID_AdditionalFullPrice) = 0
'                    liTemp.SubItems(ID_AdditionalHalfPrice) = 0
'                    liTemp.SubItems(ID_AdditionalFreePrice) = 0
'                    liTemp.SubItems(ID_AdditionalPreferential1) = 0
'                    liTemp.SubItems(ID_AdditionalPreferential2) = 0
'                    liTemp.SubItems(ID_AdditionalPreferential3) = 0
'                Case cszAdditionalType
'                    liTemp.SubItems(ID_FullPrice) = 0
'                    liTemp.SubItems(ID_HalfPrice) = 0
'                    liTemp.SubItems(ID_FreePrice) = 0
'                    liTemp.SubItems(ID_PreferentialPrice1) = 0
'                    liTemp.SubItems(ID_PreferentialPrice2) = 0
'                    liTemp.SubItems(ID_PreferentialPrice3) = 0
'                    liTemp.SubItems(ID_BedFullPrice) = 0
'                    liTemp.SubItems(ID_BedHalfPrice) = 0
'                    liTemp.SubItems(ID_BedFreePrice) = 0
'                    liTemp.SubItems(ID_BedPreferentialPrice1) = 0
'                    liTemp.SubItems(ID_BedPreferentialPrice2) = 0
'                    liTemp.SubItems(ID_BedPreferentialPrice3) = 0
'                    liTemp.SubItems(ID_AdditionalFullPrice) = m_rsBusInfo!full_price
'                    liTemp.SubItems(ID_AdditionalHalfPrice) = m_rsBusInfo!half_price
'                    liTemp.SubItems(ID_AdditionalFreePrice) = 0
'                    liTemp.SubItems(ID_AdditionalPreferential1) = m_rsBusInfo!preferential_ticket1
'                    liTemp.SubItems(ID_AdditionalPreferential2) = m_rsBusInfo!preferential_ticket2
'                    liTemp.SubItems(ID_AdditionalPreferential3) = m_rsBusInfo!preferential_ticket3
'
'            End Select
'
'            '以下几列不显示出来，只是将其存储，以备后用
'            liTemp.SubItems(ID_LimitedCount) = m_rsBusInfo!sale_ticket_quantity
'            liTemp.SubItems(ID_LimitedTime) = m_rsBusInfo!stop_sale_time
'            liTemp.SubItems(ID_BusType1) = nBusType
'            liTemp.SubItems(ID_CheckGate) = m_rsBusInfo!check_gate_id
'            liTemp.SubItems(ID_StandCount) = m_rsBusInfo!sale_stand_seat_quantity
'
'        End If
'nextstep:
'        m_rsBusInfo.MoveNext
'    Loop
End Sub

'得到检票口名称和代码
Private Function GetCheckName(pszCheckGateID As String) As String
    Dim i As Integer
    Dim szResult As String
    Dim nLen As Integer
    nLen = 0
    nLen = ArrayLength(m_aszCheckGateInfo)
    szResult = ""
    For i = 1 To nLen
        If Trim(m_aszCheckGateInfo(i, 1)) = Trim(pszCheckGateID) Then
            szResult = Trim(m_aszCheckGateInfo(i, 2))
            Exit For
        End If
    Next i
    GetCheckName = szResult

End Function
Private Sub GetInitCheckGate()
    Dim szTemp As String

'    szTemp = m_oSellDiscount.SellUnitCode
'    m_oSellDiscount.SellUnitCode = m_szCurrentUnitID
    
    m_aszCheckGateInfo = m_oSellDiscount.GetAllCheckGate()
    
'    m_oSellDiscount.SellUnitCode = szTemp
    
End Sub


Private Sub SellTicket()
    Dim i As Integer
    Dim szBookNumber() As String
    Dim dbTotalMoney As Double  '总票价
    Dim dbRealTotalMoney As Double '实际票价
    Dim aSellTicket() As TSellTicketParam
    Dim dyBusDate() As Date
    Dim szBusID() As String
    Dim szDesStationID() As String
    Dim szDesStationName() As String
    Dim szSellStationID() As String
    Dim szSellStationName() As String
    Dim szStartStationName As String
    
    Dim srSellResult() As TSellTicketResult
    Dim psgDiscount() As Single
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
    
    
    Dim liTemp As ListItem
    
    Dim nCount As Integer
    Dim nLen As Integer
    Dim nTicketOffset As Integer
    Dim nLength As Integer
    Dim nTemp As Integer
    Dim szTemp As String
    
    
    If m_bPrint Then Exit Sub
    
    If m_oSellDiscount.DiscountIsValid Then
        MsgBox "无售折扣票的权限", vbExclamation, Me.Caption
        Exit Sub
    End If
    
    dbTotalMoney = 0
    dbRealTotalMoney = 0
    
    If txtDiscount.Text > 1 Then
       MsgBox "折扣率不能大于1", vbInformation, "提示"
       txtDiscount.SetFocus
       Exit Sub
    End If
    szTemp = flblRestMoney.Caption
    cmdSell.Enabled = False
    m_bSumPriceIsEmpty = True
    m_bPrint = True
    On Error GoTo Here
    
    '以下是真正的售票处理
    '////////////////////
    
    '-------------------------------------------------------------------------------------
    If lvPreSell.ListItems.count = 0 Then
          
        If txtPreferentialSell.Text + txtHalfSell.Text + txtFullSell.Text <> 0 Then
            lblSellMsg.Caption = "正在处理售票"
            lblSellMsg.Refresh
            ReDim srSellResult(1 To 1)
            ReDim dyBusDate(1 To 1)
            ReDim szBusID(1 To 1)
            ReDim szDesStationID(1 To 1)
            ReDim szSellStationID(1 To 1)
            ReDim szSellStationName(1 To 1)
            ReDim szStationName(1 To 1)
            ReDim psgDiscount(1 To 1)
            ReDim aSellTicket(1 To 1)
            ReDim aSellTicket(1).BuyTicketInfo(1 To txtPreferentialSell.Text + txtHalfSell.Text + txtFullSell.Text)
            ReDim aSellTicket(1).pasgSellTicketPrice(1 To txtPreferentialSell.Text + txtHalfSell.Text + txtFullSell.Text)
            ReDim szBookNumber(1 To 1)
            ReDim anInsurance(1 To 1)
            '-------------------------------------------------------------------------------------
            Set liTemp = lvBus.SelectedItem
            szBookNumber(1) = frmOrderSeats.m_szBookNumber
            
            For i = 1 To txtFullSell.Text
                aSellTicket(1).BuyTicketInfo(i).nTicketType = TP_FullPrice
                aSellTicket(1).BuyTicketInfo(i).szTicketNo = GetTicketNo(i - 1)
                aSellTicket(1).BuyTicketInfo(i).szSeatNo = SelfGetSeatNo(i)
                aSellTicket(1).pasgSellTicketPrice(i) = CDbl(liTemp.SubItems(ID_FullPrice)) * CSng(txtDiscount.Text)
                aSellTicket(1).BuyTicketInfo(i).szReserved = szBookNumber(1)
                aSellTicket(1).BuyTicketInfo(i).szSeatTypeID = m_aszSeatType(cboSeatType.ListIndex + 1, 1)
                aSellTicket(1).BuyTicketInfo(i).szSeatTypeName = m_aszSeatType(cboSeatType.ListIndex + 1, 2)
            Next
            
            For i = 1 To txtHalfSell.Text
                aSellTicket(1).BuyTicketInfo(i + txtFullSell.Text).nTicketType = TP_HalfPrice
                aSellTicket(1).BuyTicketInfo(i + txtFullSell.Text).szTicketNo = GetTicketNo(i + txtFullSell.Text - 1)
                
                aSellTicket(1).BuyTicketInfo(i + txtFullSell.Text).szSeatNo = SelfGetSeatNo(i + txtFullSell.Text)
                '半价的票价
                aSellTicket(1).pasgSellTicketPrice(i + txtFullSell.Text) = CDbl(liTemp.SubItems(ID_HalfPrice)) * CSng(txtDiscount.Text)
                
                aSellTicket(1).BuyTicketInfo(i + txtFullSell.Text).szReserved = szBookNumber(1)
                
                aSellTicket(1).BuyTicketInfo(i + txtFullSell.Text).szSeatTypeID = m_aszSeatType(cboSeatType.ListIndex + 1, 1)
                aSellTicket(1).BuyTicketInfo(i + txtFullSell.Text).szSeatTypeName = m_aszSeatType(cboSeatType.ListIndex + 1, 2)
            Next
            
            For i = 1 To txtPreferentialSell.Text
                aSellTicket(1).BuyTicketInfo(i + txtFullSell.Text + txtHalfSell.Text).nTicketType = m_atTicketType(cboPreferentialTicket.ListIndex + 1).nTicketTypeID
                aSellTicket(1).BuyTicketInfo(i + txtFullSell.Text + txtHalfSell.Text).szTicketNo = GetTicketNo(i + txtFullSell.Text + txtHalfSell.Text - 1)
                
                aSellTicket(1).BuyTicketInfo(i + txtFullSell.Text + txtHalfSell.Text).szSeatNo = SelfGetSeatNo(i + txtFullSell.Text + txtHalfSell.Text)
                aSellTicket(1).pasgSellTicketPrice(i + txtFullSell.Text + txtHalfSell.Text) = CDbl(GetPreferentialPrice(True)) * CSng(txtDiscount.Text)
                
                aSellTicket(1).BuyTicketInfo(i + txtFullSell.Text + txtHalfSell.Text).szReserved = szBookNumber(1)
                aSellTicket(1).BuyTicketInfo(i + txtFullSell.Text + txtHalfSell.Text).szSeatTypeID = m_aszSeatType(cboSeatType.ListIndex + 1, 1)
                aSellTicket(1).BuyTicketInfo(i + txtFullSell.Text + txtHalfSell.Text).szSeatTypeName = m_aszSeatType(cboSeatType.ListIndex + 1, 2)
            Next
            dyBusDate(1) = CDate(flblSellDate.Caption)
            szBusID(1) = lvBus.SelectedItem.Text
            szDesStationID(1) = cboEndStation.BoundText
            szStationName(1) = ""
            psgDiscount(1) = CSng(txtDiscount.Text)
            
            szSellStationID(1) = ResolveDisplay(lvBus.SelectedItem.Tag, szStartStationName)
            szSellStationName(1) = szStartStationName
            
            anInsurance(1) = IIf(chkInsurance.Value = vbChecked, 2, 0)   '如果选中,则赋为1,否则为0
            
            
            
            srSellResult = m_oSellDiscount.SellTicket(dyBusDate, szBusID, szSellStationID, szDesStationID, szDesStationName, aSellTicket, anInsurance)
            
            IncTicketNo txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text
            
            '刷新座位信息
            If chkSetSeat.Enabled Then
                DecBusListViewSeatInfo lvBus, txtFullSell.Text + txtPreferentialSell.Text + txtHalfSell.Text, True
            Else
                DecBusListViewSeatInfo lvBus, txtFullSell.Text + txtPreferentialSell.Text + txtHalfSell.Text, False
            End If
            flblStandCount.Caption = lvBus.SelectedItem.SubItems(ID_StandCount)
            If lvBus.SelectedItem.SubItems(ID_LimitedCount) > 0 Then
                lvBus.SelectedItem.SubItems(ID_LimitedCount) = lvBus.SelectedItem.SubItems(ID_LimitedCount) - 1
                flblLimitedCount.Caption = GetStationLimitedCountStr(CInt(lvBus.SelectedItem.SubItems(ID_LimitedCount)))
            End If
    
            
    
            '以下应该是打印票的代码
            '-------------------------------------------------------------------------------------
    
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
            
            lblSellMsg.Refresh
            pnTicketCount(1) = txtPreferentialSell.Text + txtHalfSell.Text + txtFullSell.Text
            pszEndStation(1) = GetStationNameInCbo(cboEndStation.Text)
    
            pszVehicleType(1) = lvBus.SelectedItem.SubItems(ID_VehicleModel)
            pszCheckGate(1) = GetCheckName(lvBus.SelectedItem.SubItems(ID_CheckGate))
            pbSaleChange(1) = False
            pszBusDate(1) = flblSellDate.Caption
            pszOffTime(1) = lvBus.SelectedItem.SubItems(ID_OffTime)
            pszBusID(1) = lvBus.SelectedItem.Text
            aszTerminateName(1) = lvBus.SelectedItem.SubItems(ID_EndStation)
            
            panInsurance(1) = IIf(chkInsurance.Value = vbChecked, 2, 0)   '如果选中,则赋为1,否则为0
            
            For i = 1 To pnTicketCount(1)
                apiTicketInfo(1).aptPrintTicketInfo(i).nTicketType = aSellTicket(1).BuyTicketInfo(i).nTicketType
                apiTicketInfo(1).aptPrintTicketInfo(i).sgTicketPrice = srSellResult(1).asgTicketPrice(i)
                apiTicketInfo(1).aptPrintTicketInfo(i).szSeatNo = srSellResult(1).aszSeatNo(i)
                apiTicketInfo(1).aptPrintTicketInfo(i).szTicketNo = aSellTicket(1).BuyTicketInfo(i).szTicketNo
                
                '取得实际总票价
                If srSellResult(1).aszTicketType(i) <> TP_FreeTicket Then
                    dbRealTotalMoney = srSellResult(1).asgTicketPrice(i) + dbRealTotalMoney
                End If
            Next
            
            
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
            
            If ArrayLength(srSellResult) = 0 Then
                frmNotify.Show vbModal
            End If
            
            dbTotalMoney = CDbl(flblTotalPrice.Caption)
            
        End If
    Else

        nTicketOffset = 0
        nLength = 0
        lblSellMsg.Caption = "正在处理售票"
        lblSellMsg.Refresh
        
        nLen = lvPreSell.ListItems.count
        
        ReDim dyBusDate(1 To nLen)
        ReDim szBusID(1 To nLen)
        ReDim szDesStationID(1 To nLen)
        ReDim szDesStationName(1 To nLen)
        ReDim psgDiscount(1 To nLen)
        ReDim srSellResult(1 To nLen)
        ReDim szBookNumber(1 To nLen)
        ReDim aSellTicket(1 To nLen)
        ReDim szSellStationID(1 To nLen)
        ReDim szSellStationName(1 To nLen)
        ReDim anInsurance(1 To nLen)
        
        
        For nCount = 1 To lvPreSell.ListItems.count
              '-------------------------------------------------------------------------------------
            nTemp = 0
            With lvPreSell.ListItems(nCount)
                szBookNumber(nCount) = .SubItems(IT_OrderSeat)
                ReDim aSellTicket(nCount).BuyTicketInfo(1 To .SubItems(IT_SumTicketNum))
                ReDim aSellTicket(nCount).pasgSellTicketPrice(1 To .SubItems(IT_SumTicketNum))
                For i = 1 To Val(.SubItems(IT_FullNum))
                    aSellTicket(nCount).BuyTicketInfo(i).nTicketType = TP_FullPrice
                    aSellTicket(nCount).BuyTicketInfo(i).szTicketNo = GetTicketNo(i - 1 + nTicketOffset)
                    aSellTicket(nCount).BuyTicketInfo(i).szSeatNo = SelfGetSeatNo12(.SubItems(IT_SetSeatEnable), .SubItems(IT_SetSeatValue), i, .SubItems(IT_SeatNo))
                    aSellTicket(nCount).pasgSellTicketPrice(i) = CDbl(.SubItems(IT_FullPrice)) * CSng(.SubItems(IT_Discount))
                    aSellTicket(nCount).BuyTicketInfo(i).szReserved = szBookNumber(nCount)
                    aSellTicket(nCount).BuyTicketInfo(i).szSeatTypeID = .SubItems(IT_SeatType)
                    aSellTicket(nCount).BuyTicketInfo(i).szSeatTypeName = GetSeatTypeName(.SubItems(IT_SeatType))
                Next i
                
                nTemp = Val(.SubItems(IT_FullNum))
                For i = 1 To Val(.SubItems(IT_HalfNum))
                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).nTicketType = TP_HalfPrice
                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).szTicketNo = GetTicketNo(i - 1 + nTemp + nTicketOffset)
                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).szSeatNo = SelfGetSeatNo12(.SubItems(IT_SetSeatEnable), .SubItems(IT_SetSeatValue), i + nTemp, .SubItems(IT_SeatNo))
                    aSellTicket(nCount).pasgSellTicketPrice(i + nTemp) = CDbl(.SubItems(IT_HalfPrice)) * CSng(.SubItems(IT_Discount))
                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).szReserved = szBookNumber(nCount)
                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).szSeatTypeID = .SubItems(IT_SeatType)
                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).szSeatTypeName = GetSeatTypeName(.SubItems(IT_SeatType))
                Next i
                
                nTemp = nTemp + Val(.SubItems(IT_HalfNum))
                For i = 1 To Val(.SubItems(IT_FreeNum))
                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).nTicketType = TP_FreeTicket
                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).szTicketNo = GetTicketNo(i - 1 + nTemp + nTicketOffset)
                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).szSeatNo = SelfGetSeatNo12(.SubItems(IT_SetSeatEnable), .SubItems(IT_SetSeatValue), i + nTemp, .SubItems(IT_SeatNo))
                    aSellTicket(nCount).pasgSellTicketPrice(i + nTemp) = CDbl(.SubItems(IT_FullPrice)) * CSng(.SubItems(IT_Discount))
                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).szReserved = szBookNumber(nCount)
                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).szSeatTypeID = .SubItems(IT_SeatType)
                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).szSeatTypeName = GetSeatTypeName(.SubItems(IT_SeatType))
                Next i
                
                nTemp = nTemp + Val(.SubItems(IT_FreeNum))
                For i = 1 To Val(.SubItems(IT_PreferentialNum1))
                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).nTicketType = TP_PreferentialTicket1
                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).szTicketNo = GetTicketNo(i - 1 + nTemp + nTicketOffset)
                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).szSeatNo = SelfGetSeatNo12(.SubItems(IT_SetSeatEnable), .SubItems(IT_SetSeatValue), i + nTemp, .SubItems(IT_SeatNo))
                    aSellTicket(nCount).pasgSellTicketPrice(i + nTemp) = CDbl(.SubItems(IT_PreferentialType1)) * CSng(.SubItems(IT_Discount))
                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).szReserved = szBookNumber(nCount)
                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).szSeatTypeID = .SubItems(IT_SeatType)
                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).szSeatTypeName = GetSeatTypeName(.SubItems(IT_SeatType))
                Next i
                nTemp = nTemp + Val(.SubItems(IT_PreferentialNum1))
                For i = 1 To Val(.SubItems(IT_PreferentialNum2))
                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).nTicketType = TP_PreferentialTicket2
                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).szTicketNo = GetTicketNo(i - 1 + nTemp + nTicketOffset)
                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).szSeatNo = SelfGetSeatNo12(.SubItems(IT_SetSeatEnable), .SubItems(IT_SetSeatValue), i + nTemp, .SubItems(IT_SeatNo))
                    aSellTicket(nCount).pasgSellTicketPrice(i + nTemp) = CDbl(.SubItems(IT_PreferentialType2)) * CSng(.SubItems(IT_Discount))
                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).szReserved = szBookNumber(nCount)
                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).szSeatTypeID = .SubItems(IT_SeatType)
                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).szSeatTypeName = GetSeatTypeName(.SubItems(IT_SeatType))
                Next i
                nTemp = nTemp + Val(.SubItems(IT_PreferentialNum2))
                For i = 1 To Val(.SubItems(IT_PreferentialNum3))
                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).nTicketType = TP_PreferentialTicket3
                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).szTicketNo = GetTicketNo(i - 1 + nTemp + nTicketOffset)
                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).szSeatNo = SelfGetSeatNo12(.SubItems(IT_SetSeatEnable), .SubItems(IT_SetSeatValue), i + nTemp, .SubItems(IT_SeatNo))
                    aSellTicket(nCount).pasgSellTicketPrice(i + nTemp) = CDbl(.SubItems(IT_PreferentialType3)) * CSng(.SubItems(IT_Discount))
                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).szReserved = szBookNumber(nCount)
                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).szSeatTypeID = .SubItems(IT_SeatType)
                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).szSeatTypeName = GetSeatTypeName(.SubItems(IT_SeatType))
                Next i
                
               
                dyBusDate(nCount) = CDate(.SubItems(IT_BusDate))
                szBusID(nCount) = .Text
                szDesStationID(nCount) = .SubItems(IT_BoundText)
                szDesStationName(nCount) = ""
                psgDiscount(nCount) = CSng(.SubItems(IT_Discount))
                nTicketOffset = Val(.SubItems(IT_SumTicketNum)) + nTicketOffset
            End With
            
            If lvPreSell.ListItems.count < nCount Then
               szSellStationID(nCount) = ResolveDisplay(lvPreSell.ListItems(lvPreSell.ListItems.count).Tag, szStartStationName)
               
               szSellStationName(nCount) = szStartStationName
            Else
               szSellStationID(nCount) = ResolveDisplay(lvPreSell.ListItems(nCount).Tag, szStartStationName)
               szSellStationName(nCount) = szStartStationName
            End If
            
            anInsurance(nCount) = IIf(chkInsurance.Value = vbChecked, 2, 0)   '如果选中,则赋为1,否则为0
        Next nCount
        
            srSellResult = m_oSellDiscount.SellTicket(dyBusDate, szBusID, szSellStationID, szDesStationID, szDesStationName, aSellTicket, anInsurance)
            IncTicketNo nTicketOffset
            
            For nCount = 1 To lvPreSell.ListItems.count
                With lvPreSell.ListItems(nCount)
                'flblStandCount.Caption = .subitems(IT_StandCount)
                If .SubItems(IT_LimitedCount) > 0 Then
                    .SubItems(IT_LimitedCount) = .SubItems(IT_LimitedCount) - 1
                    flblLimitedCount.Caption = GetStationLimitedCountStr(CInt(.SubItems(IT_LimitedCount)))
                End If
                End With
            Next nCount
    
            '以下应该是打印票的代码
            '-------------------------------------------------------------------------------------
            ReDim apiTicketInfo(1 To lvPreSell.ListItems.count)
            ReDim pszBusDate(1 To lvPreSell.ListItems.count)
            ReDim pnTicketCount(1 To lvPreSell.ListItems.count)
            ReDim pszEndStation(1 To lvPreSell.ListItems.count)
            ReDim pszOffTime(1 To lvPreSell.ListItems.count)
            ReDim pszBusID(1 To lvPreSell.ListItems.count)
            ReDim pszVehicleType(1 To lvPreSell.ListItems.count)
            ReDim pszCheckGate(1 To lvPreSell.ListItems.count)
            ReDim pbSaleChange(1 To lvPreSell.ListItems.count)
            ReDim aszTerminateName(1 To lvPreSell.ListItems.count)
            ReDim panInsurance(1 To lvPreSell.ListItems.count)
            For nCount = 1 To lvPreSell.ListItems.count
                With lvPreSell.ListItems(nCount)
                ReDim apiTicketInfo(nCount).aptPrintTicketInfo(1 To .SubItems(IT_SumTicketNum))
                
                    pnTicketCount(nCount) = .SubItems(IT_SumTicketNum)
                    pszEndStation(nCount) = .SubItems(IT_EndStation)
                    pszVehicleType(nCount) = .SubItems(IT_VehicleModel)
                    pszCheckGate(nCount) = GetCheckName(.SubItems(IT_CheckGate))
                    pbSaleChange(nCount) = False
                    pszBusDate(nCount) = .SubItems(IT_BusDate)
                    pszOffTime(nCount) = .SubItems(IT_OffTime)
                    pszBusID(nCount) = lvPreSell.ListItems(nCount)
                    aszTerminateName(nCount) = .SubItems(IT_TerminateName)
                    
                    panInsurance(1) = IIf(chkInsurance.Value = vbChecked, 2, 0)   '如果选中,则赋为1,否则为0
            
                    For i = 1 To Val(.SubItems(IT_SumTicketNum))
                        apiTicketInfo(nCount).aptPrintTicketInfo(i).nTicketType = aSellTicket(nCount).BuyTicketInfo(i).nTicketType
                        apiTicketInfo(nCount).aptPrintTicketInfo(i).sgTicketPrice = srSellResult(nCount).asgTicketPrice(i)
                        apiTicketInfo(nCount).aptPrintTicketInfo(i).szSeatNo = srSellResult(nCount).aszSeatNo(i)
                        apiTicketInfo(nCount).aptPrintTicketInfo(i).szTicketNo = aSellTicket(nCount).BuyTicketInfo(i).szTicketNo
                        
                    Next
                End With
            Next nCount
            
             '取得实际总票价
            For nCount = 1 To ArrayLength(srSellResult)
                For i = 1 To ArrayLength(srSellResult(nCount).asgTicketPrice)
                   If srSellResult(nCount).aszTicketType(i) <> TP_FreeTicket Then
                            dbRealTotalMoney = srSellResult(nCount).asgTicketPrice(i) + dbRealTotalMoney
                   End If
                Next i
            Next nCount
            lblSellMsg.Caption = "正在打印车票"
            lblSellMsg.Refresh
            
           PrintTicket apiTicketInfo, pszBusDate, pnTicketCount, pszEndStation, pszOffTime, pszBusID, pszVehicleType, pszCheckGate, pbSaleChange, aszTerminateName, szSellStationName, anInsurance
            
           
           dbTotalMoney = CDbl(flblTotalPrice.Caption)
            '如果票价变了
        
            
            If Abs(dbRealTotalMoney - dbTotalMoney) > 0.01 Then
                frmPriceInfo.m_sngTotalPrice = dbRealTotalMoney
                frmPriceInfo.Show vbModal
            End If
           
           If IsNumeric(txtReceivedMoney.Text) Then
            
                If txtReceivedMoney.Text = 0 Then
                    m_sgTotalMoney = lblTotalMoney.Caption
                Else
                    m_sgTotalMoney = 0#
                End If
            End If
        
        '进行票款累加
           
          If ArrayLength(srSellResult) = 0 Then
                frmNotify.m_szErrorDescription = "[一般性网络错误！]  请注意对一下票号"
                frmNotify.Show vbModal
          End If
            
        End If
        
        
        m_bPreClear = True
        lblSellMsg.Caption = ""
        cmdSell.Enabled = True
        m_bPrint = False
        txtPrevDate.Text = 0
        txtTime.Text = 0
        lblmileate.Caption = ""
         '   RefreshBus True
        
        chkInsurance.Value = vbUnchecked
        
        flblRestMoney.Caption = szTemp
        frmOrderSeats.m_szBookNumber = ""
        txtSeat.Text = ""
        SetPreSellButton
        
'        lvPreSell.ListItems.Clear
        cboEndStation.SetFocus
        m_bSetFocus = False
    Exit Sub
Here:
    frmOrderSeats.m_szBookNumber = ""
    lblSellMsg.Caption = ""
    m_bPrint = False
    If err.Number = 91 Then
       MsgBox "该天该站点已无车次！", vbInformation + vbOKOnly, "提示"
    Else
        frmNotify.m_szErrorDescription = err.Description
        frmNotify.Show vbModal
      ' ShowErrorMsg
    End If
    
    txtPrevDate.Text = 0
    txtTime.Text = 0
    SetPreSellButton
    lvPreSell.ListItems.Clear
    cboEndStation.SetFocus
End Sub

Public Sub ChangeSeatType()
    If cboSeatType.ListIndex = cboSeatType.ListCount - 1 Then
        cboSeatType.ListIndex = 0
    Else
        cboSeatType.ListIndex = cboSeatType.ListIndex + 1
    End If
    cboSeatType_Change
End Sub

Private Sub txtTime_Change()
    If txtTime.Text > 24 Then txtTime.Text = 0
End Sub

Private Sub txtTime_GotFocus()
    lblTime.ForeColor = clActiveColor
    txtTime.SelStart = 0
    txtTime.SelLength = 100
End Sub

Private Sub txtTime_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        lvBus.SetFocus
    End If
End Sub


Private Sub txtTime_LostFocus()
    lblTime.ForeColor = 0
    m_bNotRefresh = False
    cboEndStation_LostFocus
End Sub




Private Function TotalInsurace() As Double
    '汇总保险费
    Dim i As Integer
    Dim nCount As Integer
    If chkInsurance.Value = vbChecked Then
        nCount = 0
        For i = 1 To lvPreSell.ListItems.count
            nCount = nCount + lvPreSell.ListItems(i).SubItems(IT_SumTicketNum)
        Next i
        nCount = nCount + Val(txtFullSell.Text) + Val(txtHalfSell.Text) + Val(txtPreferentialSell.Text)
        '保险费设为每张2元
        
        TotalInsurace = nCount * 2
    Else
        TotalInsurace = 0
    End If
End Function







