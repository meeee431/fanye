VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Object = "{A0123751-4698-48C1-A06C-A2482B5ED508}#2.0#0"; "RTComctl2.ocx"
Begin VB.Form frmSystemParam 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����ϵͳ����"
   ClientHeight    =   7155
   ClientLeft      =   1920
   ClientTop       =   2340
   ClientWidth     =   6945
   Icon            =   "frmSystemParam.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox ptMain 
      BorderStyle     =   0  'None
      Height          =   6195
      Index           =   1
      Left            =   180
      ScaleHeight     =   6195
      ScaleWidth      =   6585
      TabIndex        =   110
      Top             =   360
      Width           =   6585
      Begin VB.Frame Frame5 
         Caption         =   "��Ʊ����:"
         Height          =   2760
         Left            =   60
         TabIndex        =   143
         Top             =   3450
         Width           =   6405
         Begin VB.CheckBox chkScrollBusMode 
            Caption         =   "��ӡ��������(&S)"
            Height          =   195
            Left            =   3090
            TabIndex        =   152
            Top             =   1110
            Value           =   1  'Checked
            Width           =   1755
         End
         Begin VB.TextBox txtBusIDLen 
            Alignment       =   1  'Right Justify
            Height          =   270
            Left            =   1710
            MaxLength       =   1
            TabIndex        =   151
            Text            =   "0"
            Top             =   1035
            Width           =   555
         End
         Begin VB.TextBox txtNumber 
            Alignment       =   1  'Right Justify
            Height          =   270
            Left            =   3720
            TabIndex        =   150
            Text            =   "0"
            Top             =   1710
            Width           =   465
         End
         Begin VB.TextBox txtPreFix 
            Alignment       =   1  'Right Justify
            Height          =   270
            Left            =   1365
            TabIndex        =   149
            Text            =   "0"
            Top             =   1710
            Width           =   465
         End
         Begin VB.CheckBox chkPrintBarCode 
            Caption         =   "��ӡ����(&B)"
            Height          =   210
            Left            =   195
            TabIndex        =   148
            Top             =   270
            Value           =   1  'Checked
            Width           =   1515
         End
         Begin VB.CheckBox chkWantDirectSheetPrint 
            Caption         =   "��Ʊ�Ƿ�ʹ�ÿ��ٴ�ӡ"
            Height          =   285
            Left            =   3090
            TabIndex        =   147
            Top             =   210
            Width           =   2385
         End
         Begin VB.CheckBox chkPrintZeroValueReturnSheet 
            Caption         =   "��ӡȫ����Ʊ����Ʊ������"
            Height          =   195
            Left            =   3090
            TabIndex        =   146
            Top             =   660
            Width           =   2595
         End
         Begin VB.CheckBox chkPrintReturnChangeSheet 
            Caption         =   "��ӡ��Ʊ�����ѵ���"
            Height          =   210
            Left            =   210
            TabIndex        =   145
            Top             =   660
            Width           =   2355
         End
         Begin VB.CheckBox chkPintAid 
            Caption         =   "��ӡ����"
            Height          =   210
            Left            =   1800
            TabIndex        =   144
            Top             =   270
            Value           =   1  'Checked
            Width           =   1035
         End
         Begin MSComCtl2.UpDown UpDown17 
            Height          =   270
            Left            =   2265
            TabIndex        =   153
            Top             =   1035
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   476
            _Version        =   393216
            BuddyControl    =   "chkPrintBarCode"
            BuddyDispid     =   196647
            OrigLeft        =   2280
            OrigTop         =   1050
            OrigRight       =   2520
            OrigBottom      =   1305
            Max             =   5
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDown9 
            Height          =   270
            Left            =   1860
            TabIndex        =   154
            Top             =   1710
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   476
            _Version        =   393216
            BuddyControl    =   "chkPrintZeroValueReturnSheet"
            BuddyDispid     =   196645
            OrigLeft        =   3120
            OrigTop         =   960
            OrigRight       =   3390
            OrigBottom      =   1245
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDown8 
            Height          =   270
            Left            =   4215
            TabIndex        =   155
            Top             =   1710
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   476
            _Version        =   393216
            BuddyControl    =   "chkWantDirectSheetPrint"
            BuddyDispid     =   196646
            OrigLeft        =   3120
            OrigTop         =   960
            OrigRight       =   3390
            OrigBottom      =   1245
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.Label Label26 
            Caption         =   "ǰ׺�������ڱ�ʶ��Ʊӡˢ������,һ��Ϊ3λ,��(99A),���ֲ���ΪƱ�ŵİ��������ֲ���,һ��Ϊ7λ����:"
            Height          =   450
            Left            =   690
            TabIndex        =   165
            Top             =   2100
            Width           =   5490
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "˵��:"
            Height          =   180
            Left            =   195
            TabIndex        =   164
            Top             =   2115
            Width           =   450
         End
         Begin VB.Label Label59 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���δ�ӡ����(&L):"
            Height          =   180
            Left            =   195
            TabIndex        =   163
            Top             =   1080
            Width           =   1440
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "λ"
            Height          =   180
            Left            =   4515
            TabIndex        =   162
            Top             =   1755
            Width           =   180
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "λ"
            Height          =   180
            Left            =   2100
            TabIndex        =   161
            Top             =   1755
            Width           =   180
         End
         Begin VB.Label Label33 
            Caption         =   "-"
            Height          =   210
            Left            =   2340
            TabIndex        =   160
            Top             =   1740
            Width           =   165
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "99A-0123456      (ǰ�-����)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   690
            TabIndex        =   159
            Top             =   2460
            Width           =   2280
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "���ֲ���(&N):"
            Height          =   180
            Left            =   2535
            TabIndex        =   158
            Top             =   1755
            Width           =   1080
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00FFFFFF&
            X1              =   1410
            X2              =   5910
            Y1              =   1515
            Y2              =   1515
         End
         Begin VB.Line Line1 
            X1              =   1410
            X2              =   5910
            Y1              =   1500
            Y2              =   1500
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "��ƱƱ�ų���"
            Height          =   180
            Left            =   195
            TabIndex        =   157
            Top             =   1410
            Width           =   1080
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "ǰꡲ���(&F):"
            Height          =   180
            Left            =   195
            TabIndex        =   156
            Top             =   1755
            Width           =   1080
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "��Ʊ�������:"
         Height          =   3255
         Left            =   90
         TabIndex        =   111
         Top             =   30
         Width           =   6405
         Begin VB.TextBox txtBefore 
            Alignment       =   1  'Right Justify
            Height          =   270
            Left            =   1920
            TabIndex        =   126
            Text            =   "3"
            Top             =   954
            Width           =   675
         End
         Begin VB.TextBox txtAfter 
            Alignment       =   1  'Right Justify
            Height          =   270
            Left            =   4770
            TabIndex        =   125
            Text            =   "3"
            Top             =   954
            Width           =   735
         End
         Begin VB.TextBox txtInternetBusInfo 
            Alignment       =   1  'Right Justify
            Height          =   270
            Left            =   3000
            TabIndex        =   124
            Text            =   "0"
            Top             =   1728
            Width           =   735
         End
         Begin VB.TextBox txtChangeCharge 
            Alignment       =   1  'Right Justify
            Height          =   270
            Left            =   4740
            TabIndex        =   123
            Text            =   "0"
            Top             =   180
            Width           =   600
         End
         Begin VB.TextBox txtCancelTicket 
            Alignment       =   1  'Right Justify
            Height          =   270
            Left            =   5115
            TabIndex        =   122
            Text            =   "0"
            Top             =   567
            Width           =   585
         End
         Begin VB.TextBox txtStopSale 
            Alignment       =   1  'Right Justify
            Height          =   270
            Left            =   1785
            TabIndex        =   121
            Text            =   "0"
            ToolTipText     =   "����ʱ����ͣ��ʱ��Ĳ�ֵ"
            Top             =   567
            Width           =   585
         End
         Begin VB.TextBox txtPreSale 
            Alignment       =   1  'Right Justify
            Height          =   270
            Left            =   1335
            TabIndex        =   120
            Text            =   "0"
            Top             =   180
            Width           =   585
         End
         Begin VB.TextBox txtScrollBusTime_RT 
            Alignment       =   1  'Right Justify
            Height          =   270
            Left            =   3000
            TabIndex        =   119
            Text            =   "0"
            Top             =   2115
            Width           =   735
         End
         Begin VB.CheckBox chkWantListNoSeatBus 
            Caption         =   "�г�����λ�ĳ���"
            Height          =   255
            Left            =   4050
            TabIndex        =   118
            Top             =   2123
            Width           =   1785
         End
         Begin VB.CheckBox chkSellStationCanSellEachOther 
            Caption         =   "�ϳ�վ���Ƿ��໥��Ʊ"
            Height          =   225
            Left            =   210
            TabIndex        =   117
            Top             =   2547
            Width           =   2145
         End
         Begin VB.CheckBox chkAllowScrollBusSaleForever 
            Caption         =   "��������Ƿ���Զ����"
            Height          =   225
            Left            =   3060
            TabIndex        =   116
            Top             =   2547
            Width           =   2175
         End
         Begin VB.TextBox txtDiscountTicketInTicketTypePosition 
            Alignment       =   1  'Right Justify
            Height          =   270
            Left            =   2220
            TabIndex        =   115
            Text            =   "4"
            Top             =   1341
            Width           =   645
         End
         Begin VB.TextBox txtAllowSellScreenShow 
            Alignment       =   1  'Right Justify
            Height          =   270
            Left            =   4770
            TabIndex        =   114
            Text            =   "1"
            Top             =   1341
            Width           =   705
         End
         Begin VB.TextBox txtHalfTicketTypeRatio 
            Alignment       =   1  'Right Justify
            Height          =   270
            Left            =   5430
            TabIndex        =   113
            Text            =   "1"
            Top             =   1728
            Width           =   435
         End
         Begin VB.TextBox txtCardBuyTicketInterval 
            Alignment       =   1  'Right Justify
            Height          =   270
            Left            =   3015
            TabIndex        =   112
            Text            =   "0"
            Top             =   2940
            Width           =   735
         End
         Begin MSComCtl2.UpDown UpDown6 
            Height          =   270
            Left            =   5700
            TabIndex        =   127
            Top             =   570
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   476
            _Version        =   393216
            BuddyControl    =   "txtCancelTicket"
            BuddyDispid     =   196621
            OrigLeft        =   5700
            OrigTop         =   747
            OrigRight       =   5940
            OrigBottom      =   1017
            Max             =   120
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDown5 
            Height          =   270
            Left            =   2370
            TabIndex        =   128
            ToolTipText     =   "����ʱ����ͣ��ʱ��Ĳ�ֵ"
            Top             =   570
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   476
            _Version        =   393216
            BuddyControl    =   "txtStopSale"
            BuddyDispid     =   196620
            OrigLeft        =   2340
            OrigTop         =   747
            OrigRight       =   2580
            OrigBottom      =   1017
            Max             =   120
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDown4 
            Height          =   270
            Left            =   1920
            TabIndex        =   129
            Top             =   180
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   476
            _Version        =   393216
            BuddyControl    =   "txtPreSale"
            BuddyDispid     =   196619
            OrigLeft        =   1920
            OrigTop         =   330
            OrigRight       =   2160
            OrigBottom      =   600
            Max             =   40
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.Label Label61 
            AutoSize        =   -1  'True
            Caption         =   "��Ʊ��ʾǰ������:"
            Height          =   180
            Left            =   195
            TabIndex        =   142
            Top             =   999
            Width           =   1530
         End
         Begin VB.Label Label62 
            AutoSize        =   -1  'True
            Caption         =   "��Ʊ��ʾ�󼸷���:"
            Height          =   180
            Left            =   3060
            TabIndex        =   141
            Top             =   999
            Width           =   1530
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            Caption         =   "����վ�۱�վ����ʱ�ļ��\Сʱ:"
            Height          =   180
            Left            =   195
            TabIndex        =   140
            Top             =   1773
            Width           =   2700
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "��ǩ����\Ԫ(&G):"
            Height          =   180
            Left            =   3060
            TabIndex        =   139
            Top             =   225
            Width           =   1350
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "������Ʊʱ��\����(&D):"
            Height          =   180
            Left            =   3060
            TabIndex        =   138
            Top             =   612
            Width           =   1890
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "ͣ��ʱ��\����(&T):"
            Height          =   180
            Left            =   210
            TabIndex        =   137
            ToolTipText     =   "����ʱ����ͣ��ʱ��Ĳ�ֵ"
            Top             =   612
            Width           =   1530
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Ԥ������(&P):"
            Height          =   180
            Left            =   180
            TabIndex        =   136
            Top             =   225
            Width           =   1080
         End
         Begin VB.Label Label51 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������������ϳ�ʱ���\Сʱ:"
            Height          =   180
            Left            =   195
            TabIndex        =   135
            Top             =   2160
            Width           =   2520
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ۿ�Ʊ��Ʊ�����λ��:"
            Height          =   180
            Left            =   195
            TabIndex        =   134
            Top             =   1386
            Width           =   1890
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "һ����ʾ�ĳ�����:"
            Height          =   180
            Left            =   3060
            TabIndex        =   133
            Top             =   1386
            Width           =   1530
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            Caption         =   "��Ʊ���۰ٷֱ�:"
            Height          =   180
            Left            =   4050
            TabIndex        =   132
            Top             =   1773
            Width           =   1350
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   180
            Left            =   5940
            TabIndex        =   131
            Top             =   1773
            Width           =   90
         End
         Begin VB.Label Label52 
            AutoSize        =   -1  'True
            Caption         =   "ͬһ֤����Ʊ����ʱ����(����):"
            Height          =   180
            Left            =   210
            TabIndex        =   130
            Top             =   2985
            Width           =   2790
         End
      End
   End
   Begin VB.PictureBox ptMain 
      BorderStyle     =   0  'None
      Height          =   6195
      Index           =   0
      Left            =   180
      ScaleHeight     =   6195
      ScaleWidth      =   6585
      TabIndex        =   1
      Top             =   390
      Width           =   6585
      Begin VB.Frame Frame15 
         Caption         =   "Ʊ�۹���"
         Height          =   2505
         Left            =   90
         TabIndex        =   2
         Top             =   2820
         Width           =   6315
         Begin VB.TextBox txtPriceDetailKeepBit 
            Height          =   285
            Left            =   2160
            TabIndex        =   9
            Top             =   330
            Width           =   510
         End
         Begin VB.TextBox TxtSpeed 
            Height          =   270
            Left            =   5430
            TabIndex        =   8
            Top             =   277
            Width           =   585
         End
         Begin VB.TextBox TxtAdvanceDistance1 
            Height          =   285
            Left            =   2415
            TabIndex        =   7
            Top             =   735
            Width           =   480
         End
         Begin VB.TextBox TxtAdvanceDistance2 
            Height          =   270
            Left            =   5430
            TabIndex        =   6
            Top             =   675
            Width           =   585
         End
         Begin VB.TextBox txtFarDistanceAddChargeItem 
            Height          =   285
            Left            =   2040
            TabIndex        =   5
            Top             =   1560
            Width           =   870
         End
         Begin VB.TextBox txtRoadBuildChargeItem 
            Height          =   270
            Left            =   4980
            TabIndex        =   4
            Top             =   1500
            Width           =   1035
         End
         Begin VB.TextBox txtSpringChargeItem 
            Height          =   285
            Left            =   2010
            TabIndex        =   3
            Top             =   1950
            Width           =   900
         End
         Begin MSComCtl2.DTPicker dtpNightShiftTime2 
            Height          =   285
            Left            =   4980
            TabIndex        =   10
            Top             =   1080
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   503
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "00:00:00"
            Format          =   97058818
            CurrentDate     =   37039
         End
         Begin MSComCtl2.DTPicker dtpNightShiftTime1 
            Height          =   315
            Left            =   1800
            TabIndex        =   11
            Top             =   1125
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   97058818
            CurrentDate     =   37039
         End
         Begin MSComCtl2.UpDown UpDown18 
            Height          =   285
            Left            =   2670
            TabIndex        =   12
            Top             =   330
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            Value           =   1
            BuddyControl    =   "txtPriceDetailKeepBit"
            BuddyDispid     =   196611
            OrigLeft        =   2655
            OrigTop         =   315
            OrigRight       =   2895
            OrigBottom      =   615
            Max             =   2
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.Label Label60 
            AutoSize        =   -1  'True
            Caption         =   "Ʊ����С��λ������:"
            Height          =   180
            Left            =   180
            TabIndex        =   21
            Top             =   375
            Width           =   1710
         End
         Begin VB.Label LblSpeed 
            AutoSize        =   -1  'True
            Caption         =   "����ƽ���ٶ�(����/Сʱ):"
            Height          =   180
            Left            =   3210
            TabIndex        =   20
            Top             =   315
            Width           =   2160
         End
         Begin VB.Label LblNightShiftTime1 
            AutoSize        =   -1  'True
            Caption         =   "ҹ��ѵ�һʱ���:"
            Height          =   180
            Left            =   180
            TabIndex        =   19
            Top             =   1185
            Width           =   1530
         End
         Begin VB.Label LblNightShiftTime2 
            AutoSize        =   -1  'True
            Caption         =   "ҹ��ѵڶ�ʱ���:"
            Height          =   180
            Left            =   3210
            TabIndex        =   18
            Top             =   1125
            Width           =   1530
         End
         Begin VB.Label Label66 
            AutoSize        =   -1  'True
            Caption         =   "���˷ѵ�һ�����(����):"
            Height          =   180
            Left            =   180
            TabIndex        =   17
            Top             =   780
            Width           =   2070
         End
         Begin VB.Label Label67 
            AutoSize        =   -1  'True
            Caption         =   "���˷ѵڶ������(����):"
            Height          =   180
            Left            =   3210
            TabIndex        =   16
            Top             =   720
            Width           =   2070
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "250K�ӳ�Ʊ����λ��:"
            Height          =   180
            Left            =   180
            TabIndex        =   15
            Top             =   1605
            Width           =   1710
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "������Ʊ����λ��:"
            Height          =   180
            Left            =   3210
            TabIndex        =   14
            Top             =   1545
            Width           =   1530
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "���˷�Ʊ����λ��:"
            Height          =   180
            Left            =   180
            TabIndex        =   13
            Top             =   1995
            Width           =   1530
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "ϵͳ��ʶ:"
         Height          =   1035
         Left            =   90
         TabIndex        =   64
         Top             =   90
         Width           =   6345
         Begin VB.TextBox txtLocalStation 
            Height          =   285
            Left            =   1830
            TabIndex        =   66
            Top             =   615
            Width           =   3600
         End
         Begin VB.TextBox txtLocalUnit 
            Height          =   285
            Left            =   1830
            TabIndex        =   65
            Top             =   255
            Width           =   3600
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "����վվ�����(&S):"
            Height          =   180
            Left            =   150
            TabIndex        =   68
            Top             =   645
            Width           =   1620
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "����λ����(&U):"
            Height          =   180
            Left            =   165
            TabIndex        =   67
            Top             =   300
            Width           =   1260
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "���ι���"
         Height          =   1455
         Left            =   90
         TabIndex        =   22
         Top             =   1260
         Width           =   6315
         Begin VB.TextBox txtAddReBus 
            Height          =   270
            Left            =   1260
            MaxLength       =   2
            TabIndex        =   27
            Top             =   285
            Width           =   1320
         End
         Begin VB.CheckBox ChkAllowSlitp 
            Caption         =   "�����ֲ�ͣ��վ"
            Height          =   195
            Left            =   3240
            TabIndex        =   26
            Top             =   1043
            Width           =   1770
         End
         Begin VB.CheckBox chkAllowMakeSaleBus 
            Caption         =   "��������������Ʊ���λ���"
            Height          =   255
            Left            =   3240
            TabIndex        =   25
            Top             =   668
            Width           =   2505
         End
         Begin VB.CheckBox chkMakeSotpEniroment 
            Caption         =   "����������ͣ�೵��"
            Height          =   240
            Left            =   150
            TabIndex        =   24
            Top             =   675
            Value           =   1  'Checked
            Width           =   1980
         End
         Begin VB.CheckBox chkEndStationCanSale 
            Caption         =   "�Ӱ೵Ĭ�����վ����"
            Height          =   240
            Left            =   150
            TabIndex        =   23
            Top             =   1020
            Width           =   2220
         End
         Begin VB.Label lblReAddBus 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�Ӱ೵ǰ׺:"
            Height          =   180
            Left            =   150
            TabIndex        =   28
            Top             =   330
            Width           =   990
         End
      End
   End
   Begin VB.PictureBox ptMain 
      BorderStyle     =   0  'None
      Height          =   6195
      Index           =   2
      Left            =   900
      ScaleHeight     =   6195
      ScaleWidth      =   6585
      TabIndex        =   30
      Top             =   600
      Width           =   6585
      Begin RTComctl3.CoolButton cmdDeleteLine 
         Height          =   315
         Left            =   5190
         TabIndex        =   101
         Top             =   900
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "ɾ��"
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
         MICON           =   "frmSystemParam.frx":038A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin RTComctl3.CoolButton cmdAddLine 
         Height          =   315
         Left            =   5190
         TabIndex        =   100
         Top             =   480
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "����"
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
         MICON           =   "frmSystemParam.frx":03A6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Frame Frame11 
         Height          =   1035
         Left            =   180
         TabIndex        =   31
         Top             =   3180
         Width           =   5985
         Begin VB.CheckBox chkScorll_RT 
            Caption         =   "�������������Ʊ(&S)"
            Height          =   255
            Left            =   180
            TabIndex        =   32
            Top             =   0
            Value           =   1  'Checked
            Width           =   2025
         End
         Begin VB.TextBox txtScorll_RT_Charge 
            Alignment       =   1  'Right Justify
            Height          =   270
            Left            =   2520
            TabIndex        =   33
            Text            =   "0"
            Top             =   420
            Width           =   735
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "�������ε���Ʊ����(&4):	"
            Height          =   180
            Left            =   300
            TabIndex        =   34
            Top             =   465
            Width           =   2520
         End
      End
      Begin VSFlex7LCtl.VSFlexGrid vsReturnRatio 
         Height          =   1965
         Left            =   180
         TabIndex        =   35
         Top             =   480
         Width           =   4815
         _cx             =   8493
         _cy             =   3466
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
         AllowUserResizing=   3
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmSystemParam.frx":03C2
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
         Begin RTComctl2.RevTimer RevTimer1 
            Height          =   30
            Left            =   4620
            Top             =   2280
            Width           =   30
            _ExtentX        =   53
            _ExtentY        =   53
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������Ʊ��������:"
         Height          =   180
         Left            =   360
         TabIndex        =   38
         Top             =   180
         Width           =   1530
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "˵��:"
         Height          =   180
         Left            =   210
         TabIndex        =   37
         Top             =   2715
         Width           =   450
      End
      Begin VB.Label Label28 
         Caption         =   "����Ϊ0.00-1.00,ָ��վ��ȡ����������,���һ��Ʊʱ��ķ���һ��Ϊ0.2,���˸��ÿͳ�Ʊ�۵�80%��"
         Height          =   435
         Left            =   705
         TabIndex        =   36
         Top             =   2595
         Width           =   4335
      End
   End
   Begin VB.PictureBox ptMain 
      BorderStyle     =   0  'None
      Height          =   6195
      Index           =   3
      Left            =   1200
      ScaleHeight     =   6195
      ScaleWidth      =   6585
      TabIndex        =   29
      Top             =   330
      Width           =   6585
      Begin VB.Frame Frame10 
         Caption         =   "·��:"
         Height          =   1980
         Left            =   180
         TabIndex        =   58
         Top             =   3030
         Width           =   6165
         Begin VB.TextBox txtRoldTitle 
            Height          =   285
            Left            =   3300
            TabIndex        =   104
            Top             =   1080
            Width           =   1965
         End
         Begin VB.CheckBox chkPrintSheetTitle 
            Caption         =   "�Ƿ��ӡ����"
            Height          =   285
            Left            =   270
            TabIndex        =   103
            Top             =   1110
            Value           =   1  'Checked
            Width           =   1650
         End
         Begin VB.CheckBox chkPrintSheetNum 
            Caption         =   "�Ƿ��ӡ·����"
            Height          =   285
            Left            =   270
            TabIndex        =   102
            Top             =   660
            Value           =   1  'Checked
            Width           =   1650
         End
         Begin VB.TextBox txtRoldSheetNum 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1755
            TabIndex        =   60
            Text            =   "0"
            Top             =   270
            Width           =   375
         End
         Begin VB.CheckBox chkTD 
            Caption         =   "��Ʊ�Ƿ�ʹ�ÿ��ٴ�ӡ"
            Height          =   285
            Left            =   2190
            TabIndex        =   59
            Top             =   660
            Value           =   1  'Checked
            Width           =   4500
         End
         Begin MSComCtl2.UpDown UpDown12 
            Height          =   285
            Left            =   2175
            TabIndex        =   61
            Top             =   270
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtRoldSheetNum"
            BuddyDispid     =   196709
            OrigLeft        =   3075
            OrigTop         =   795
            OrigRight       =   3345
            OrigBottom      =   1065
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.Label Label20 
            Caption         =   "·������(&T):	"
            Height          =   240
            Left            =   2190
            TabIndex        =   105
            Top             =   1140
            Width           =   1080
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            Caption         =   "λ"
            Height          =   180
            Left            =   2520
            TabIndex        =   63
            Top             =   360
            Width           =   180
         End
         Begin VB.Label Label19 
            Caption         =   "·���ų���(&L):	"
            Height          =   240
            Left            =   300
            TabIndex        =   62
            Top             =   330
            Width           =   1320
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "�ĳ�:"
         Height          =   705
         Left            =   180
         TabIndex        =   55
         Top             =   2265
         Width           =   6390
         Begin VB.CheckBox chkLowChangeRide 
            Caption         =   "�Ƿ�����ͼ۳�Ʊ�ĳ�(&D)"
            Enabled         =   0   'False
            Height          =   210
            Left            =   2970
            TabIndex        =   57
            Top             =   315
            Width           =   2490
         End
         Begin VB.CheckBox chkChangeRide 
            Caption         =   "�Ƿ�����ĳ�(&R)"
            Height          =   240
            Left            =   270
            TabIndex        =   56
            Top             =   285
            Value           =   1  'Checked
            Width           =   2130
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "ʱ�����/����:"
         Height          =   2025
         Left            =   180
         TabIndex        =   39
         Top             =   120
         Width           =   6150
         Begin VB.CheckBox chkAllowStartChectNotRearchTime 
            Caption         =   "�Ƿ�����δ������ʱ��ֱ�ӿ���"
            Height          =   285
            Left            =   720
            TabIndex        =   44
            Top             =   1500
            Width           =   3285
         End
         Begin VB.TextBox txtExChkTime 
            Alignment       =   1  'Right Justify
            Height          =   270
            Left            =   4785
            TabIndex        =   43
            Text            =   "0"
            ToolTipText     =   "ʱ��γ���"
            Top             =   1095
            Width           =   360
         End
         Begin VB.TextBox txtChkTime 
            Alignment       =   1  'Right Justify
            Height          =   270
            Left            =   4785
            TabIndex        =   42
            Text            =   "0"
            ToolTipText     =   "ʱ��γ���"
            Top             =   600
            Width           =   360
         End
         Begin VB.TextBox txtExChkStartTime 
            Alignment       =   1  'Right Justify
            Height          =   270
            Left            =   2880
            TabIndex        =   41
            Text            =   "0"
            ToolTipText     =   "�෢��ǰ������?"
            Top             =   1110
            Width           =   360
         End
         Begin VB.TextBox txtChkStartTime 
            Alignment       =   1  'Right Justify
            Height          =   270
            Left            =   2880
            TabIndex        =   40
            Text            =   "0"
            ToolTipText     =   "�෢��ǰ������?"
            Top             =   600
            Width           =   360
         End
         Begin MSComCtl2.UpDown UpDown13 
            Height          =   270
            Left            =   3210
            TabIndex        =   45
            ToolTipText     =   "�෢��ǰ������?"
            Top             =   600
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   476
            _Version        =   393216
            BuddyControl    =   "txtChkStartTime"
            BuddyDispid     =   196723
            OrigLeft        =   3120
            OrigTop         =   225
            OrigRight       =   3390
            OrigBottom      =   510
            Max             =   120
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDown14 
            Height          =   270
            Left            =   3240
            TabIndex        =   46
            ToolTipText     =   "�෢��ǰ������?"
            Top             =   1110
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   476
            _Version        =   393216
            BuddyControl    =   "txtExChkStartTime"
            BuddyDispid     =   196722
            OrigLeft        =   3120
            OrigTop         =   225
            OrigRight       =   3390
            OrigBottom      =   510
            Max             =   120
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDown15 
            Height          =   270
            Left            =   5145
            TabIndex        =   47
            ToolTipText     =   "ʱ��γ���"
            Top             =   600
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   476
            _Version        =   393216
            BuddyControl    =   "txtChkTime"
            BuddyDispid     =   196721
            OrigLeft        =   3120
            OrigTop         =   225
            OrigRight       =   3390
            OrigBottom      =   510
            Max             =   120
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDown16 
            Height          =   270
            Left            =   5145
            TabIndex        =   48
            ToolTipText     =   "ʱ��γ���"
            Top             =   1095
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   476
            _Version        =   393216
            BuddyControl    =   "txtExChkTime"
            BuddyDispid     =   196720
            OrigLeft        =   3120
            OrigTop         =   225
            OrigRight       =   3390
            OrigBottom      =   510
            Max             =   120
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "����"
            Height          =   180
            Left            =   270
            TabIndex        =   54
            Top             =   840
            Width           =   360
         End
         Begin VB.Line Line6 
            BorderColor     =   &H00FFFFFF&
            X1              =   705
            X2              =   5570
            Y1              =   960
            Y2              =   960
         End
         Begin VB.Line Line5 
            BorderColor     =   &H00808080&
            X1              =   705
            X2              =   5570
            Y1              =   945
            Y2              =   945
         End
         Begin VB.Line Line4 
            BorderColor     =   &H00FFFFFF&
            X1              =   1110
            X2              =   5580
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00808080&
            X1              =   1110
            X2              =   5580
            Y1              =   465
            Y2              =   465
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "������Ʊ"
            Height          =   180
            Left            =   270
            TabIndex        =   53
            Top             =   375
            Width           =   720
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "��Ʊʱ��(&E):"
            Height          =   180
            Left            =   3645
            TabIndex        =   52
            ToolTipText     =   "ʱ��γ���"
            Top             =   1140
            Width           =   1080
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "������ʱ��\����ǰ(&W):"
            Height          =   180
            Left            =   720
            TabIndex        =   51
            ToolTipText     =   "�෢��ǰ������?"
            Top             =   1125
            Width           =   2070
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "��Ʊʱ��(&J):"
            Height          =   180
            Left            =   3630
            TabIndex        =   50
            ToolTipText     =   "ʱ��γ���"
            Top             =   645
            Width           =   1080
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "����ʱ��\����ǰ(&S):"
            Height          =   180
            Left            =   720
            TabIndex        =   49
            ToolTipText     =   "�෢��ǰ������?"
            Top             =   645
            Width           =   1710
         End
      End
   End
   Begin VB.PictureBox ptMain 
      BorderStyle     =   0  'None
      Height          =   6195
      Index           =   4
      Left            =   420
      ScaleHeight     =   6195
      ScaleWidth      =   6585
      TabIndex        =   81
      Top             =   390
      Width           =   6585
      Begin VB.Frame Frame6 
         Caption         =   "����"
         Height          =   2445
         Left            =   60
         TabIndex        =   86
         Top             =   1500
         Width           =   6405
         Begin VB.CheckBox chkAllowSettleTotalNegative 
            Caption         =   "�Ƿ����½���ĸ����㵽����"
            Height          =   210
            Left            =   180
            TabIndex        =   92
            Top             =   1155
            Width           =   2955
         End
         Begin VB.TextBox txtSettleNegativeSplitItem 
            Height          =   270
            Left            =   3420
            TabIndex        =   91
            Top             =   1515
            Width           =   585
         End
         Begin VB.TextBox txtServiceItemPosition 
            Height          =   270
            Left            =   1860
            TabIndex        =   90
            Top             =   1890
            Width           =   585
         End
         Begin VB.TextBox txtCarriageItemPosition 
            Height          =   270
            Left            =   4440
            TabIndex        =   89
            Top             =   1890
            Width           =   585
         End
         Begin VB.CheckBox chkAllowSplitBySomeTimes 
            Caption         =   "�Ƿ�����һ��·�����������"
            Height          =   210
            Left            =   180
            TabIndex        =   88
            Top             =   360
            Width           =   2925
         End
         Begin VB.CheckBox chkAllowSplitAboveFactQuantity 
            Caption         =   "�Ƿ����������������ʵ��·��վ������"
            Height          =   210
            Left            =   180
            TabIndex        =   87
            Top             =   765
            Width           =   4785
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "�����½���ĸ�������ڽ������λ��:"
            Height          =   180
            Left            =   180
            TabIndex        =   95
            Top             =   1560
            Width           =   3150
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "����ѵ���Ŀλ��:"
            Height          =   180
            Left            =   180
            TabIndex        =   94
            Top             =   1935
            Width           =   1530
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "�˷ѵ���Ŀλ��:"
            Height          =   180
            Left            =   2850
            TabIndex        =   93
            Top             =   1935
            Width           =   1350
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "�̶�����"
         Height          =   1365
         Left            =   60
         TabIndex        =   82
         Top             =   60
         Width           =   6405
         Begin VB.CheckBox chkFixFeeUpdateEachMonth 
            Caption         =   "�Ƿ�̶�������ÿ���¸��µ�"
            Height          =   210
            Left            =   180
            TabIndex        =   84
            Top             =   975
            Width           =   4095
         End
         Begin VB.TextBox txtFixFeeItem 
            Height          =   285
            Left            =   2010
            TabIndex        =   83
            Top             =   270
            Width           =   3945
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            Caption         =   "˵�����̶�������֮�����ö��Ÿ���"
            Height          =   180
            Left            =   180
            TabIndex        =   109
            Top             =   630
            Width           =   3030
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "����Ĺ̶�������:"
            Height          =   180
            Left            =   180
            TabIndex        =   85
            Top             =   330
            Width           =   1530
         End
      End
   End
   Begin VB.PictureBox ptMain 
      BorderStyle     =   0  'None
      Height          =   6195
      Index           =   5
      Left            =   1560
      ScaleHeight     =   6195
      ScaleWidth      =   6585
      TabIndex        =   69
      Top             =   300
      Width           =   6585
      Begin VB.Frame Frame12 
         Height          =   2985
         Left            =   300
         TabIndex        =   70
         Top             =   90
         Width           =   6405
         Begin VB.TextBox txtBookLong 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1935
            TabIndex        =   106
            Text            =   "0"
            Top             =   690
            Width           =   375
         End
         Begin VB.CheckBox chkUseBook 
            Caption         =   "ʹ��Ԥ������"
            Height          =   195
            Left            =   120
            TabIndex        =   73
            Top             =   0
            Width           =   1515
         End
         Begin VB.Frame Frame13 
            Height          =   75
            Left            =   960
            TabIndex        =   72
            Top             =   1260
            Width           =   4335
         End
         Begin VB.TextBox txtBook 
            Enabled         =   0   'False
            Height          =   270
            Left            =   3360
            TabIndex        =   71
            Text            =   "30"
            Top             =   240
            Width           =   360
         End
         Begin MSComCtl2.UpDown udBook 
            Height          =   270
            Left            =   3750
            TabIndex        =   74
            Top             =   240
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   476
            _Version        =   393216
            Value           =   30
            BuddyControl    =   "txtBook"
            BuddyDispid     =   196757
            OrigLeft        =   3720
            OrigTop         =   240
            OrigRight       =   3960
            OrigBottom      =   510
            Max             =   120
            Enabled         =   -1  'True
         End
         Begin VB.Label Label38 
            Caption         =   "Ԥ���ų���(&L):	"
            Height          =   240
            Left            =   480
            TabIndex        =   108
            Top             =   750
            Width           =   1320
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "λ"
            Height          =   180
            Left            =   2490
            TabIndex        =   107
            Top             =   780
            Width           =   180
         End
         Begin VB.Label Label48 
            AutoSize        =   -1  'True
            Caption         =   "˵����"
            Height          =   180
            Left            =   300
            TabIndex        =   80
            Top             =   1200
            Width           =   540
         End
         Begin VB.Label Label47 
            Caption         =   "      3��Ԥ��ƱʧЧʱ�������趨�ڷ���ǰ��ã���ƱԱ����Ԥ��Ʊ�������˿͡�ʱ�����Ϊ0��120���ӡ�"
            Height          =   360
            Left            =   360
            TabIndex        =   79
            Top             =   1920
            Width           =   4800
         End
         Begin VB.Label Label45 
            Caption         =   "      1����ѡ��ʹ�á�Ԥ�����ܡ�����Ԥ����ϵͳ��ʧЧ���������������ϵͳ����ӦЧ�ʡ�"
            Height          =   360
            Left            =   360
            TabIndex        =   78
            Top             =   1470
            Width           =   4800
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            Caption         =   "����"
            Height          =   180
            Left            =   4080
            TabIndex        =   77
            Top             =   330
            Width           =   360
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            Caption         =   "����ǰ"
            Height          =   240
            Left            =   2700
            TabIndex        =   76
            Top             =   330
            Width           =   540
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "Ԥ��ƱʧЧʱ��:"
            Height          =   180
            Left            =   480
            TabIndex        =   75
            Top             =   330
            Width           =   1590
         End
      End
   End
   Begin MSComctlLib.TabStrip tsMain 
      Height          =   6645
      Left            =   120
      TabIndex        =   0
      Top             =   30
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   11721
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   6
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "����(&1)"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "��Ʊ(&2)"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "��Ʊ(&3)"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "��Ʊ(&4)"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "����(&5)"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "����(&6)"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin RTComctl3.CoolButton cmdHelp 
      Height          =   300
      Left            =   5070
      TabIndex        =   99
      Top             =   6750
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   529
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
      MICON           =   "frmSystemParam.frx":0444
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdColse 
      Height          =   300
      Left            =   3900
      TabIndex        =   98
      Top             =   6750
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   529
      BTYPE           =   3
      TX              =   "ȡ��(&C)"
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
      MICON           =   "frmSystemParam.frx":0460
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdOK 
      Height          =   300
      Left            =   2730
      TabIndex        =   97
      Top             =   6750
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   529
      BTYPE           =   3
      TX              =   "ȷ��(&O)"
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
      MICON           =   "frmSystemParam.frx":047C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdApply 
      Height          =   300
      Left            =   1560
      TabIndex        =   96
      Top             =   6750
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   529
      BTYPE           =   3
      TX              =   "Ӧ��(&A)"
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
      MICON           =   "frmSystemParam.frx":0498
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
Attribute VB_Name = "frmSystemParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************
'* Source File Name:frmSystemParam
'* Project Name:PSTSMan.vbp
'* Engineer:
'* Data Generated:2002-08-15
'* Last Revision Date:2002-08-15
'* Brief Description:ϵͳ����
'* Remark:2006��3�·�������������
'**********************************************************
Option Explicit
Option Base 1
Dim g_oSysParam As New SystemParam

Const cnScheme = 1
Const cnSale = 2
Const cnReturn = 3
Const cnCheck = 4
Const cnSettle = 5
Const cnOther = 6

Const cnTabs = 6

'/*����*/
'ϵͳ��ʶ
Dim m_szLocalUnitID  '����λ����
Dim m_szLocalStationID  '��վվ����
'���ι���
Dim m_szAdditionBusPreFix As String   '�Ӱ೵�ε�ǰ׺
Dim m_szMakeStopEnviroment As Boolean  '�������Ƿ�����ͣ�೵��
Dim m_szAllowMakeEnviromentSaleBus As Boolean  '�Ƿ�������������Ʊ���εĻ���
Dim m_szEndStatinCanSale As Boolean  '�Ƿ�Ӱ೵ֻ���յ�վ
Dim m_szAllowSlitpNotPassStation As Boolean  '�Ƿ�����ͣ��վ����
Dim m_szImmediatelyQuery  '��վ���ѯ
'Ʊ�۹���
Dim m_szPriceDetailKeepBit As Integer 'Ʊ����С�����λ��
Dim m_szBusSpeed As Integer   '����ƽ���ٶ�
Dim m_szAdvanceDistance1 As Single  '���˼Ӽ۾���1
Dim m_szAdvanceDistance2 As Single   '���˼Ӽ۾���2
Dim m_szNightShiftTime1 As Date  'ҹ��ѵ�һʱ���
Dim m_szNightShiftTime2 As Date  'ҹ��ѵڶ�ʱ���
Dim m_szFarDistanceAddChargeItem As Integer  '250K�ӳɷ�Ϊ�ڼ���Ʊ����
Dim m_szRoadBuildChargeItem As Integer  '������Ϊ�ڼ���Ʊ����
Dim m_szSpringChargeItem As Integer   '���˷�Ϊ�ڼ���Ʊ����


'/*��Ʊ*/
'��Ʊ�������
Dim m_szPreSellDate As Integer   'Ԥ������
Dim m_szChangeCharge As Single   '��ǩ����
Dim m_szStopSellTime As Integer   'ͣ��ʱ��
Dim m_szQueryTime As Integer   '��ѯʱ��
Dim m_szCancelTicketTime As Integer  '��Ʊʱ��
Dim m_szExtraBusShowBefore As Double   '��Ʊ������ʾǰ������
Dim m_szExtraBusShowAfter As Double    '��Ʊ������ʾ�󼸷���
Dim m_szDiscountTicketInTicketTypePosition As Integer  '�ۿ�Ʊ��Ʊ�����λ��
'Const cszAllowSellScreenShow = "AllowSellScreenShow" '������ʾ��С��0��ʹ�÷�����ʾ������0���Ƿ�����ʾ����������������
Dim m_szAllowSellScreenShow As Integer
Dim m_szInternetBusShow  As Double '������Ʊ������ʾʱ��
Dim m_szScrollBusLatestTime As Integer      '������������ϳ�ʱ��
'Const cszWantListNoSeatBus = "WantListNoSeatBus" 'WantListNoSeatBus�Ƿ��г�����λ�ĳ��Σ�ֵΪ1��ʾ�г�����λ�ĳ���,0��ʾ���г�
Dim m_szWantListNoSeatBus As Boolean
Dim m_szSellStationCanSellEachOther As Boolean   '��Ʊվ֮���Ƿ�����໥��Ʊ
Dim m_szAllowScrollBusSaleForever As Boolean  '��������Ƿ���Զ����
Dim m_szHalfTicketTypeRatio As Double    '��Ʊ���۰ٷֱ�
'��Ʊ����
Dim m_szPrintBarCode As Boolean   '��ƱҪ���ӡ������
Dim m_szPrintAid As Boolean   '��ƱҪ���ӡ����
'Const cszWantDirectSheetPrint = "WantDirectSheetPrint" '��Ʊʱ�Ƿ�ʹ�ÿ��ٴ�ӡ��ֵΪ1��ʾ���ٴ�ӡ,0��ʾʹ����ͨ��ӡ
Dim m_szWantDirectSheetPrint As Boolean
Dim m_szIsPrintReturnChangeSheet As Boolean   '�Ƿ��ӡ��Ʊ�����ѵ���
Dim m_szIsPrintZeroValueReturnSheet As Boolean  '�Ƿ��ӡ������Ϊ0����Ʊ������
Dim m_szPrintBusIDLen As Integer  '���δ�ӡ����
Dim m_szPrintScrollBusMode As Boolean   '�������η���ʱ���ӡ��ʽ
Dim m_szTicketPrefixLen As Integer  'Ʊ��ǰꡲ��ֳ�
Dim m_szTicketNumberLen As Integer   'Ʊ�����ֲ��ֳ�
Dim m_nCardBuyTicketInterval As Integer   'ͬһ֤����Ʊ����ʱ���������ӣ�

'/*��Ʊ*/
Dim m_szScrollBusReturnRatio As Single  '��ˮ���ε���Ʊ����
Dim m_szScrollBusReturnTime As Single  '��ˮ������Ʊʱ���
Dim m_szScrollBusCanReturnTicket As Boolean   '��ˮ�����ܷ���Ʊ


'/*��Ʊ*/
'ʱ�����
Dim m_szBeginCheckTime As Integer   '��ʼ��Ʊʱ��һʱ����
Dim m_szCheckTicketTime As Integer  '��Ʊʱ��(����)
Dim m_szLatestExtraCheckTime As Integer  '����ʱ��������ʱ��(����)
Dim m_szExtraCheckTime As Integer  '����ʱ��(����)
Dim m_szAllowStartChectNotRearchTime As Boolean '����δ������ʱ��ֱ�ӿ���
'�ĳ�
Dim m_szAllowChangeRide As Boolean  '�Ƿ�����ĳ�
Dim m_szAllowChangeRideLowerPrice As Boolean  '�Ƿ�����ӵ�Ʊ�۵ĳ��θĳ˵���Ʊ�۵ĳ���
'·��
Dim m_szCheckSheetLen As Integer  '·����ų���
Dim m_szWantTDPrintType As Boolean  '·���Ƿ��״�ʽ
Dim m_szWantPrintRoadSheetID As Boolean  '�Ƿ��ӡ·����
Dim m_szWantPrintRoadSheetTitle As Boolean   '�Ƿ��ӡ����
Dim m_szRoadSheetTitle As String   '·������


'/*����*/
Dim m_szFixFeeItem As String   '����Ĺ̶�������
Dim m_szIsFixFeeUpdateEachMonth As Boolean  '�Ƿ�̶�������ÿ���¸��µ�
Dim m_szAllowSplitBySomeTimes As Boolean  '�Ƿ�����·���ּ��������㣺 0������,1����
Dim m_szAllowSplitAboveFactQuantity As Boolean   '�Ƿ�����·��������������ʵ��������0������,1����
Dim m_szAllowSettleTotalNegative As Boolean  '�Ƿ����½���ĸ����㵽���£� 0������,1����
Dim m_szSettleNegativeSplitItem As Integer  '�����½���ĸ�������ڽ������λ��
Dim m_szServiceItemPosition As Integer  '����ѵ���Ŀλ��
Dim m_szCarriageItemPosition As Integer  '�˷ѵ���Ŀλ��


'/*����*/
Dim m_szIsBookValid As Boolean   '�Ƿ���Ԥ��ϵͳ
Dim m_szBookTime As Integer  'Ԥ��ʱ��
Dim m_szBookNumberLen As Integer  'Ԥ���ų���

Private Sub chkChangeRide_Click()
    If chkChangeRide.Value = vbChecked Then
        chkLowChangeRide.Enabled = True
    Else
        chkLowChangeRide.Enabled = False
    End If
End Sub

Private Sub chkPrintSheetTitle_Click()
    If chkPrintSheetTitle.Value = vbChecked Then
        txtRoldTitle.Enabled = True
        txtRoldTitle.BackColor = vbWhite
    Else
        txtRoldTitle.Enabled = False
        txtRoldTitle.BackColor = cGreyColor
    End If
End Sub

Private Sub chkScorll_RT_Click()
    If chkScorll_RT.Value = vbChecked Then
        txtScorll_RT_Charge.Enabled = True
        txtScorll_RT_Charge.BackColor = vbWhite
    Else
        txtScorll_RT_Charge.Enabled = False
        txtScorll_RT_Charge.BackColor = cGreyColor
    End If
End Sub

Private Sub Form_Load()
    Me.Top = (Screen.Height - Me.ScaleHeight) / 2
    Me.Left = (Screen.Width - Me.ScaleWidth) / 2
    On Error GoTo ErrorHandle
    SetPtVisible (0)
    g_oSysParam.Init g_oActUser
    LoadInfo
    GetReturnTktRatio
    DisplayInfoInit
Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub SetPtVisible(pnIndex)
    '����ͼƬ��ɼ�
    Dim i As Integer
    For i = 0 To cnTabs - 1
        ptMain(i).Visible = False
        ptMain(i).Left = ptMain(0).Left
        ptMain(i).Top = ptMain(0).Top
    Next i
    ptMain(pnIndex).Visible = True
End Sub

Private Sub tsMain_Click()
    SetPtVisible (tsMain.SelectedItem.Index - 1)
End Sub

Private Sub chkUseBook_Click()
    DoUseBook
End Sub

Private Sub cmdAddLine_Click()
    AddLine
End Sub

Private Sub cmdApply_Click()
    Me.MousePointer = vbHourglass
    GetInfo
'    ModifySysParam
    SetReturnTktRatio
    Me.MousePointer = vbDefault
End Sub

Private Sub cmdColse_Click()
    Unload Me
End Sub

Private Sub cmdDeleteLine_Click()
    DeleteLine
End Sub

Private Sub cmdHelp_Click()
    Select Case tsMain.SelectedItem.Index
        Case 1
            frmSystemParam.HelpContextID = 5000020
        Case 2
            frmSystemParam.HelpContextID = 5000030
        Case 3
            frmSystemParam.HelpContextID = 5000040
        Case 4
            frmSystemParam.HelpContextID = 5000050
        Case 5
            frmSystemParam.HelpContextID = 5000060
        Case 6
            frmSystemParam.HelpContextID = 5000070
    End Select
    DisplayHelp Me
End Sub

Private Sub cmdOK_Click()
    Me.MousePointer = vbHourglass
    GetInfo
'    ModifySysParam
    Me.MousePointer = vbDefault
    SetReturnTktRatio
    Unload Me
End Sub

Private Sub LoadInfo()
    On Error GoTo ErrorHandle
    
    '/*����*/
    'ϵͳ��ʶ
    m_szLocalUnitID = g_oSysParam.UnitID
    m_szLocalStationID = g_oSysParam.StationID
    '���ι���
    m_szAdditionBusPreFix = g_oSysParam.AdditionBusPreFix
    m_szMakeStopEnviroment = g_oSysParam.MekeStopEnviroment
    m_szAllowMakeEnviromentSaleBus = g_oSysParam.MakeEnviromentSaleBus
    m_szEndStatinCanSale = g_oSysParam.EndStationCanSale
    m_szAllowSlitpNotPassStation = g_oSysParam.AllowSlitp
    'Ʊ�۹���
    m_szPriceDetailKeepBit = g_oSysParam.PriceDetailKeepBit
    m_szBusSpeed = g_oSysParam.BusSpeed
    m_szAdvanceDistance1 = g_oSysParam.AdvanceDistance1
    m_szAdvanceDistance2 = g_oSysParam.AdvanceDistance2
    m_szNightShiftTime1 = g_oSysParam.NightShiftTime1
    m_szNightShiftTime2 = g_oSysParam.NightShiftTime2
    m_szFarDistanceAddChargeItem = g_oSysParam.FarDistanceAddChargeItem
    m_szRoadBuildChargeItem = g_oSysParam.RoadBuildChargeItem
    m_szSpringChargeItem = g_oSysParam.SpringChargeItem

    '/*��Ʊ*/
    '��Ʊ�������
    m_szPreSellDate = g_oSysParam.PreSellDate
    m_szChangeCharge = g_oSysParam.ChangeCharge
    m_szStopSellTime = g_oSysParam.StopSellTime
    m_szQueryTime = g_oSysParam.QueryTime
    m_szCancelTicketTime = g_oSysParam.CancelTicketTime
    m_szExtraBusShowBefore = g_oSysParam.ExtraBusShowBefore
    m_szExtraBusShowAfter = g_oSysParam.ExtraBusShowAfter
    m_szDiscountTicketInTicketTypePosition = g_oSysParam.DiscountTicketInTicketTypePosition
    m_szAllowSellScreenShow = g_oSysParam.AllowSellScreenShow
    m_szInternetBusShow = g_oSysParam.InternetBusShow
    m_szScrollBusLatestTime = g_oSysParam.ScrollBusLatestTime
    m_szWantListNoSeatBus = g_oSysParam.WantListNoSeatBus
    m_szSellStationCanSellEachOther = g_oSysParam.SellStationCanSellEachOther
    m_szAllowScrollBusSaleForever = g_oSysParam.AllowScrollBusSaleForever
    m_szHalfTicketTypeRatio = g_oSysParam.HalfTicketTypeRatio
    m_nCardBuyTicketInterval = g_oSysParam.CardBuyTicketInterval
    
    '��Ʊ����
    m_szPrintBarCode = g_oSysParam.PrintBarCode
    m_szPrintAid = g_oSysParam.PrintAid
    m_szWantDirectSheetPrint = g_oSysParam.WantDirectSheetPrint
    m_szIsPrintReturnChangeSheet = g_oSysParam.IsPrintReturnChangeSheet
    m_szIsPrintZeroValueReturnSheet = g_oSysParam.IsPrintZeroValueReturnSheet
    m_szPrintBusIDLen = g_oSysParam.PrintBusIDLen
    m_szPrintScrollBusMode = g_oSysParam.PrintScrollBusMode
    m_szTicketPrefixLen = g_oSysParam.TicketPrefixLen
    m_szTicketNumberLen = g_oSysParam.TicketNumberLen
    
    '/*��Ʊ*/
    m_szScrollBusReturnRatio = g_oSysParam.ScrollBusReturnRatio
    m_szScrollBusReturnTime = g_oSysParam.ScrollBusReturnTime
    m_szScrollBusReturnTime = g_oSysParam.ScrollBusCanReturnTicket
    
    '/*��Ʊ*/
    'ʱ�����
    m_szBeginCheckTime = g_oSysParam.BeginCheckTime
    m_szCheckTicketTime = g_oSysParam.CheckTicketTime
    m_szLatestExtraCheckTime = g_oSysParam.LatestExtraCheckTime
    m_szExtraCheckTime = g_oSysParam.ExtraCheckTime
    m_szAllowStartChectNotRearchTime = g_oSysParam.AllowStartChectNotRearchTime
    '�ĳ�
    m_szAllowChangeRide = g_oSysParam.AllowChangeRide
    m_szAllowChangeRideLowerPrice = g_oSysParam.AllowChangeRideLowerPrice
    '·��
    m_szCheckSheetLen = g_oSysParam.CheckSheetLen
    m_szWantTDPrintType = g_oSysParam.WantTDPrintType
    m_szWantPrintRoadSheetID = g_oSysParam.WantPrintRoadSheetID
    m_szWantPrintRoadSheetTitle = g_oSysParam.WantPrintRoadSheetTitle
    m_szRoadSheetTitle = g_oSysParam.RoadSheetTitle
    
    '/*����*/
    m_szFixFeeItem = g_oSysParam.FixFeeItem
    m_szIsFixFeeUpdateEachMonth = g_oSysParam.IsFixFeeUpdateEachMonth
    m_szAllowSplitBySomeTimes = g_oSysParam.AllowSplitBySomeTimes
    m_szAllowSplitAboveFactQuantity = g_oSysParam.AllowSplitAboveFactQuantity
    m_szAllowSettleTotalNegative = g_oSysParam.AllowSettleTotalNegative
    m_szSettleNegativeSplitItem = g_oSysParam.SettleNegativeSplitItem
    m_szServiceItemPosition = g_oSysParam.ServiceItemPosition
    m_szCarriageItemPosition = g_oSysParam.CarriageItemPosition
    
    '/*����*/
    m_szIsBookValid = g_oSysParam.IsBookValid
    m_szBookTime = g_oSysParam.BookTime
    m_szBookNumberLen = g_oSysParam.BookNumberLen

Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub DisplayInfoInit()

    '/*����*/
    'ϵͳ��ʶ
    txtLocalUnit.Text = m_szLocalUnitID
    txtLocalStation.Text = m_szLocalStationID
    '���ι���
    If m_szAdditionBusPreFix <> "" Then
            txtAddReBus.Text = m_szAdditionBusPreFix
    End If
    If m_szMakeStopEnviroment = True Then
        chkMakeSotpEniroment.Value = Checked
    Else
        chkMakeSotpEniroment.Value = Unchecked
    End If
    If m_szAllowMakeEnviromentSaleBus = True Then
        chkAllowMakeSaleBus.Value = vbChecked
    Else
        chkAllowMakeSaleBus.Value = Unchecked
    End If
    If m_szEndStatinCanSale = True Then
        chkEndStationCanSale.Value = vbChecked
    Else
        chkEndStationCanSale.Value = Unchecked
    End If
    If m_szAllowSlitpNotPassStation = True Then
        ChkAllowSlitp.Value = vbChecked
    Else
        ChkAllowSlitp.Value = Unchecked
    End If
    'Ʊ�۹���
    txtPriceDetailKeepBit.Text = CInt(Val(m_szPriceDetailKeepBit))
    TxtSpeed.Text = CInt(Val(m_szBusSpeed))
    TxtAdvanceDistance1.Text = m_szAdvanceDistance1
    TxtAdvanceDistance2.Text = m_szAdvanceDistance2
    dtpNightShiftTime1.Value = m_szNightShiftTime1
    dtpNightShiftTime2.Value = m_szNightShiftTime2
    txtFarDistanceAddChargeItem.Text = m_szFarDistanceAddChargeItem
    txtRoadBuildChargeItem.Text = m_szRoadBuildChargeItem
    txtSpringChargeItem.Text = m_szSpringChargeItem

    '/*��Ʊ*/
    '��Ʊ�������
    txtPreSale.Text = CInt(Val(m_szPreSellDate))
    txtChangeCharge.Text = Val(m_szChangeCharge)
    txtStopSale.Text = CInt(Val(m_szStopSellTime))
    txtCancelTicket.Text = CInt(Val(m_szCancelTicketTime))
    txtBefore.Text = m_szExtraBusShowBefore
    txtAfter.Text = m_szExtraBusShowAfter
    txtDiscountTicketInTicketTypePosition.Text = m_szDiscountTicketInTicketTypePosition
    txtAllowSellScreenShow.Text = m_szAllowSellScreenShow
    txtInternetBusInfo.Text = Trim(m_szInternetBusShow)
    txtScrollBusTime_RT.Text = m_szScrollBusLatestTime
    If m_szWantListNoSeatBus = True Then
        chkWantListNoSeatBus.Value = vbChecked
    Else
        chkWantListNoSeatBus.Value = Unchecked
    End If
    If m_szSellStationCanSellEachOther = True Then
        chkSellStationCanSellEachOther.Value = vbChecked
    Else
        chkSellStationCanSellEachOther.Value = Unchecked
    End If
    If m_szAllowScrollBusSaleForever = True Then
        chkAllowScrollBusSaleForever.Value = vbChecked
    Else
        chkAllowScrollBusSaleForever.Value = Unchecked
    End If
    txtHalfTicketTypeRatio.Text = CInt(Val(m_szHalfTicketTypeRatio))
    txtCardBuyTicketInterval.Text = Val(m_nCardBuyTicketInterval)
    
    '��Ʊ����
    If m_szPrintBarCode = True Then
        chkPrintBarCode.Value = Checked
    ElseIf m_szPrintBarCode = False Then
        chkPrintBarCode.Value = Unchecked
    Else
        MsgBox "ϵͳ����[�Ƿ��ӡ����]����Ϊ1 �� 0 ,���ݿ����.", vbExclamation, cszMsg
    End If
    If m_szPrintAid = True Then
        chkPintAid.Value = vbChecked
    Else
        chkPintAid.Value = Unchecked
    End If
    If m_szWantDirectSheetPrint = True Then
        chkWantDirectSheetPrint.Value = vbChecked
    Else
        chkWantDirectSheetPrint.Value = Unchecked
    End If
    If m_szIsPrintReturnChangeSheet = True Then
        chkPrintReturnChangeSheet.Value = Checked
    ElseIf m_szIsPrintReturnChangeSheet = False Then
        chkPrintReturnChangeSheet.Value = Unchecked
    Else
        MsgBox "ϵͳ����[�Ƿ��ӡ��Ʊ�����ѵ���]����Ϊ1 �� 0 ,���ݿ����.", vbExclamation, cszMsg
    End If
    If m_szIsPrintZeroValueReturnSheet = True Then
        chkPrintZeroValueReturnSheet.Value = Checked
    ElseIf m_szIsPrintZeroValueReturnSheet = False Then
        chkPrintZeroValueReturnSheet.Value = Unchecked
    Else
        MsgBox "ϵͳ����[�Ƿ��ӡȫ����Ʊ����Ʊ������]����Ϊ1 �� 0 ,���ݿ����.", vbExclamation, cszMsg
    End If
    txtBusIDLen.Text = CInt(Val(m_szPrintBusIDLen))
    chkScrollBusMode.Value = IIf(m_szPrintScrollBusMode = True, vbChecked, vbUnchecked)
    txtPreFix.Text = CInt(Val(m_szTicketPrefixLen))
    txtNumber.Text = CInt(Val(m_szTicketNumberLen))
    
    '/*��Ʊ*/
    txtScorll_RT_Charge.Text = Val(m_szScrollBusReturnRatio)
    If m_szScrollBusCanReturnTicket = True Then
        chkScorll_RT.Value = Checked
    ElseIf m_szScrollBusCanReturnTicket = False Then
        chkScorll_RT.Value = Unchecked
    Else
        MsgBox "ϵͳ����[���������ܷ���Ʊ]����Ϊ1 �� 0 ,���ݿ����.", vbExclamation, cszMsg
    End If
    If chkScorll_RT.Value = Checked Then
        txtScorll_RT_Charge.Enabled = True
    Else
        txtScorll_RT_Charge.Enabled = False
    End If
    
    '/*��Ʊ*/
    'ʱ�����
    txtChkStartTime.Text = CInt(Val(m_szBeginCheckTime))
    txtChkTime.Text = CInt(Val(m_szCheckTicketTime))
    txtExChkStartTime.Text = CInt(Val(m_szLatestExtraCheckTime))
    txtExChkTime.Text = CInt(Val(m_szExtraCheckTime))
    chkAllowStartChectNotRearchTime.Value = IIf(m_szAllowStartChectNotRearchTime = True, vbChecked, vbUnchecked)
    '�ĳ�
    If m_szAllowChangeRide = True Then
        chkChangeRide.Value = Checked
    ElseIf m_szAllowChangeRide = False Then
        chkChangeRide.Value = Unchecked
    Else
        MsgBox "ϵͳ����[�Ƿ�����ĳ�]����Ϊ1 �� 0 ,���ݿ����.", vbExclamation, cszMsg
    End If
    If m_szAllowChangeRideLowerPrice = True Then
        chkLowChangeRide.Value = Checked
    ElseIf m_szAllowChangeRideLowerPrice = False Then
        chkLowChangeRide.Value = Unchecked
    Else
        MsgBox "ϵͳ����[�Ƿ�����ͼ۳�Ʊ�ĳ�]����Ϊ1 �� 0 ,���ݿ����.", vbExclamation, cszMsg
    End If
    If chkChangeRide.Value = Checked Then
        chkLowChangeRide.Enabled = True
    Else
        chkLowChangeRide.Enabled = False
    End If
    '·��
    txtRoldSheetNum.Text = CInt(Val(m_szCheckSheetLen))
    If m_szWantPrintRoadSheetID = True Then
        chkPrintSheetNum.Value = Checked
    ElseIf m_szWantPrintRoadSheetID = False Then
        chkPrintSheetNum.Value = Unchecked
    Else
        MsgBox "ϵͳ����[�Ƿ��ӡ·����]����Ϊ1 �� 0 ,���ݿ����.", vbExclamation, cszMsg
    End If
    If m_szWantTDPrintType = True Then
        chkTD.Value = Checked
    ElseIf m_szWantTDPrintType = False Then
        chkTD.Value = Unchecked
    Else
        MsgBox "ϵͳ����[�Ƿ�Ʊ���״�ʽ]����Ϊ1 �� 0 ,���ݿ����.", vbExclamation, cszMsg
    End If
    If m_szWantPrintRoadSheetTitle = True Then
        chkPrintSheetTitle.Value = Checked
    ElseIf m_szWantPrintRoadSheetTitle = False Then
        chkPrintSheetTitle.Value = Unchecked
    Else
        MsgBox "ϵͳ����[�Ƿ��ӡ·������]����Ϊ1 �� 0 ,���ݿ����.", vbExclamation, cszMsg
    End If
    txtRoldTitle.Text = m_szRoadSheetTitle
    If chkPrintSheetTitle.Value = Checked Then
        txtRoldTitle.Enabled = True
    Else
        txtRoldTitle.Enabled = False
    End If
    
    '/*����*/
    txtFixFeeItem.Text = m_szFixFeeItem
    If m_szIsFixFeeUpdateEachMonth = True Then
        chkFixFeeUpdateEachMonth.Value = Checked
    Else
        chkFixFeeUpdateEachMonth.Value = Unchecked
    End If
    If m_szAllowSplitBySomeTimes = True Then
        chkAllowSplitBySomeTimes.Value = Checked
    Else
        chkAllowSplitBySomeTimes.Value = Unchecked
    End If
    If m_szAllowSplitAboveFactQuantity = True Then
        chkAllowSplitAboveFactQuantity.Value = Checked
    Else
        chkAllowSplitAboveFactQuantity.Value = Unchecked
    End If
    If m_szAllowSettleTotalNegative = True Then
        chkAllowSettleTotalNegative.Value = Checked
    Else
        chkAllowSettleTotalNegative.Value = Unchecked
    End If
    txtSettleNegativeSplitItem.Text = m_szSettleNegativeSplitItem
    txtServiceItemPosition.Text = m_szServiceItemPosition
    txtCarriageItemPosition.Text = m_szCarriageItemPosition
    
    '/*����*/
    chkUseBook.Value = IIf(m_szIsBookValid = True, vbChecked, vbUnchecked)
    txtBook.Text = m_szBookTime
    txtBookLong.Text = m_szBookNumberLen
    DoUseBook
            
Exit Sub
ErrorHandle:
    MsgBox "���ݿ����Ҳ���ǹ���Ԥ��ϵͳ��ϵͳ����ȱ�ٻ����ݲ��ԣ�", vbExclamation + vbOKOnly, cszMsg
End Sub

Private Sub GetInfo()
    '/*����*/
    'ϵͳ��ʶ
     If m_szLocalUnitID <> txtLocalUnit.Text Then
         m_szLocalUnitID = txtLocalUnit.Text
         g_oSysParam.UnitID = m_szLocalUnitID
     End If
     If m_szLocalStationID <> txtLocalStation.Text Then
         m_szLocalStationID = txtLocalStation.Text
         g_oSysParam.StationID = m_szLocalStationID
     End If
     '���ι���
     If m_szAdditionBusPreFix <> Trim(txtAddReBus.Text) Then
         If Right(Trim(txtAddReBus.Text), 1) <> "%" Then
             m_szAdditionBusPreFix = Trim(txtAddReBus.Text) & "%"
         Else
             m_szAdditionBusPreFix = Trim(txtAddReBus.Text)
         End If
         g_oSysParam.AdditionBusPreFix = m_szAdditionBusPreFix
     End If
     If m_szMakeStopEnviroment <> chkMakeSotpEniroment.Value Then
         m_szMakeStopEnviroment = chkMakeSotpEniroment.Value
         g_oSysParam.MekeStopEnviroment = m_szMakeStopEnviroment
     End If
     If m_szAllowMakeEnviromentSaleBus <> chkAllowMakeSaleBus.Value Then
         m_szAllowMakeEnviromentSaleBus = chkAllowMakeSaleBus.Value
         g_oSysParam.MakeEnviromentSaleBus = m_szAllowMakeEnviromentSaleBus
     End If
     If m_szEndStatinCanSale <> chkEndStationCanSale.Value Then
         m_szEndStatinCanSale = chkEndStationCanSale.Value
         g_oSysParam.EndStationCanSale = m_szEndStatinCanSale
     End If
     If m_szAllowSlitpNotPassStation <> ChkAllowSlitp.Value Then
         m_szAllowSlitpNotPassStation = ChkAllowSlitp.Value
         g_oSysParam.AllowSlitp = m_szAllowSlitpNotPassStation
     End If
     'Ʊ�۹���
     If m_szPriceDetailKeepBit <> CInt(Val(txtPriceDetailKeepBit.Text)) Then
         m_szPriceDetailKeepBit = CInt(Val(txtPriceDetailKeepBit.Text))
         g_oSysParam.PriceDetailKeepBit = m_szPriceDetailKeepBit
     End If
     If m_szBusSpeed <> Val(TxtSpeed.Text) Then
         m_szBusSpeed = Val(TxtSpeed.Text)
         g_oSysParam.BusSpeed = m_szBusSpeed
     End If
     If m_szAdvanceDistance1 <> Val(TxtAdvanceDistance1.Text) Then
         m_szAdvanceDistance1 = Val(TxtAdvanceDistance1.Text)
         g_oSysParam.AdvanceDistance1 = m_szAdvanceDistance1
     End If
     If m_szAdvanceDistance2 <> Val(TxtAdvanceDistance2.Text) Then
         m_szAdvanceDistance2 = Val(TxtAdvanceDistance2.Text)
         g_oSysParam.AdvanceDistance2 = m_szAdvanceDistance2
     End If
     If m_szNightShiftTime1 <> dtpNightShiftTime1.Value Then
         m_szNightShiftTime1 = dtpNightShiftTime1.Value
         g_oSysParam.NightShiftTime1 = m_szNightShiftTime1
     End If
     If m_szNightShiftTime2 <> dtpNightShiftTime2.Value Then
         m_szNightShiftTime2 = dtpNightShiftTime2.Value
         g_oSysParam.NightShiftTime2 = m_szNightShiftTime2
     End If
     If m_szFarDistanceAddChargeItem <> CInt(Val(txtFarDistanceAddChargeItem.Text)) Then
         m_szFarDistanceAddChargeItem = CInt(Val(txtFarDistanceAddChargeItem.Text))
         g_oSysParam.FarDistanceAddChargeItem = m_szFarDistanceAddChargeItem
     End If
     If m_szRoadBuildChargeItem <> CInt(Val(txtRoadBuildChargeItem.Text)) Then
         m_szRoadBuildChargeItem = CInt(Val(txtRoadBuildChargeItem.Text))
         g_oSysParam.RoadBuildChargeItem = m_szRoadBuildChargeItem
     End If
     If m_szSpringChargeItem <> CInt(Val(txtSpringChargeItem.Text)) Then
         m_szSpringChargeItem = CInt(Val(txtSpringChargeItem.Text))
         g_oSysParam.SpringChargeItem = m_szSpringChargeItem
     End If

    '/*��Ʊ*/
    '��Ʊ�������
     If m_szPreSellDate <> CInt(Val(txtPreSale.Text)) Then
         m_szPreSellDate = CInt(Val(txtPreSale.Text))
         g_oSysParam.PreSellDate = m_szPreSellDate
     End If
     If m_szChangeCharge <> Val(txtChangeCharge.Text) Then
         m_szChangeCharge = Val(txtChangeCharge.Text)
         g_oSysParam.ChangeCharge = m_szChangeCharge
     End If
     If m_szStopSellTime <> CInt(Val(txtStopSale.Text)) Then
         m_szStopSellTime = CInt(Val(txtStopSale.Text))
         g_oSysParam.StopSellTime = m_szStopSellTime
     End If
     If m_szCancelTicketTime <> CInt(Val(txtCancelTicket.Text)) Then
         m_szCancelTicketTime = CInt(Val(txtCancelTicket.Text))
         g_oSysParam.CancelTicketTime = m_szCancelTicketTime
     End If
     If m_szExtraBusShowBefore <> txtBefore.Text Then
         m_szExtraBusShowBefore = txtBefore.Text
         g_oSysParam.ExtraBusShowBefore = m_szExtraBusShowBefore
     End If
     If m_szExtraBusShowAfter <> txtAfter.Text Then
         m_szExtraBusShowAfter = txtAfter.Text
         g_oSysParam.ExtraBusShowAfter = m_szExtraBusShowAfter
     End If
     If m_szDiscountTicketInTicketTypePosition <> CInt(Val(txtDiscountTicketInTicketTypePosition.Text)) Then
         m_szDiscountTicketInTicketTypePosition = CInt(Val(txtDiscountTicketInTicketTypePosition.Text))
         g_oSysParam.DiscountTicketInTicketTypePosition = m_szDiscountTicketInTicketTypePosition
     End If
     If m_szAllowSellScreenShow <> Val(txtAllowSellScreenShow.Text) Then
         m_szAllowSellScreenShow = Val(txtAllowSellScreenShow.Text)
         g_oSysParam.AllowSellScreenShow = m_szAllowSellScreenShow
     End If
     If m_szInternetBusShow <> Val(txtInternetBusInfo.Text) Then
         m_szInternetBusShow = Val(txtInternetBusInfo.Text)
         g_oSysParam.InternetBusShow = m_szInternetBusShow
     End If
     If m_szScrollBusLatestTime <> Val(txtScrollBusTime_RT.Text) Then
         m_szScrollBusLatestTime = Val(txtScrollBusTime_RT.Text)
         g_oSysParam.ScrollBusLatestTime = m_szScrollBusLatestTime
     End If
     If m_szWantListNoSeatBus <> chkWantListNoSeatBus.Value Then
         If chkWantListNoSeatBus.Value = vbChecked Then
             m_szWantListNoSeatBus = True
         Else
             m_szWantListNoSeatBus = False
         End If
         g_oSysParam.WantListNoSeatBus = m_szWantListNoSeatBus
     End If
     If m_szSellStationCanSellEachOther <> chkSellStationCanSellEachOther.Value Then
         If chkSellStationCanSellEachOther.Value = vbChecked Then
             m_szSellStationCanSellEachOther = True
         Else
             m_szSellStationCanSellEachOther = False
         End If
         g_oSysParam.SellStationCanSellEachOther = m_szSellStationCanSellEachOther
     End If
     If m_szAllowScrollBusSaleForever <> chkAllowScrollBusSaleForever.Value Then
         If chkAllowScrollBusSaleForever.Value = vbChecked Then
             m_szAllowScrollBusSaleForever = True
         Else
             m_szAllowScrollBusSaleForever = False
         End If
         g_oSysParam.AllowScrollBusSaleForever = m_szAllowScrollBusSaleForever
     End If
     If m_szHalfTicketTypeRatio <> CInt(Val(txtHalfTicketTypeRatio.Text)) Then
         m_szHalfTicketTypeRatio = CInt(Val(txtHalfTicketTypeRatio.Text))
         g_oSysParam.HalfTicketTypeRatio = m_szHalfTicketTypeRatio
     End If
     If m_nCardBuyTicketInterval <> Val(txtCardBuyTicketInterval.Text) Then
        m_nCardBuyTicketInterval = Val(txtCardBuyTicketInterval.Text)
        g_oSysParam.CardBuyTicketInterval = m_nCardBuyTicketInterval
    End If
    
     '��Ʊ����
     If m_szPrintBarCode <> chkPrintBarCode.Value Then
         If chkPrintBarCode.Value = vbChecked Then
             m_szPrintBarCode = True
         Else
             m_szPrintBarCode = False
         End If
         g_oSysParam.PrintBarCode = m_szPrintBarCode
     End If
     If m_szPrintAid <> chkPintAid.Value Then
         If chkPintAid.Value = vbChecked Then
             m_szPrintAid = True
         Else
             m_szPrintAid = False
         End If
         g_oSysParam.PrintAid = m_szPrintAid
     End If
     If m_szWantDirectSheetPrint <> chkWantDirectSheetPrint.Value Then
         If chkWantDirectSheetPrint.Value = vbChecked Then
             m_szWantDirectSheetPrint = True
         Else
             m_szWantDirectSheetPrint = False
         End If
         g_oSysParam.WantDirectSheetPrint = m_szWantDirectSheetPrint
     End If
     If m_szIsPrintReturnChangeSheet <> chkPrintReturnChangeSheet.Value Then
         If chkPrintReturnChangeSheet.Value = vbChecked Then
             m_szIsPrintReturnChangeSheet = True
         Else
             m_szIsPrintReturnChangeSheet = False
         End If
         g_oSysParam.IsPrintReturnChangeSheet = m_szIsPrintReturnChangeSheet
     End If
     If m_szIsPrintZeroValueReturnSheet <> chkPrintZeroValueReturnSheet.Value Then
         If chkPrintZeroValueReturnSheet.Value = vbChecked Then
             m_szIsPrintZeroValueReturnSheet = True
         Else
             m_szIsPrintZeroValueReturnSheet = False
         End If
         g_oSysParam.IsPrintZeroValueReturnSheet = m_szIsPrintZeroValueReturnSheet
     End If
     If m_szPrintBusIDLen <> Val(txtBusIDLen.Text) Then
         m_szPrintBusIDLen = Val(txtBusIDLen.Text)
         g_oSysParam.PrintBusIDLen = m_szPrintBusIDLen
     End If
     If m_szPrintScrollBusMode <> chkScrollBusMode.Value Then
         If chkScrollBusMode.Value = vbChecked Then
             m_szPrintScrollBusMode = True
         Else
             m_szPrintScrollBusMode = False
         End If
         g_oSysParam.PrintScrollBusMode = m_szPrintScrollBusMode
     End If
     If m_szTicketPrefixLen <> CInt(Val(txtPreFix.Text)) Then
         m_szTicketPrefixLen = CInt(Val(txtPreFix.Text))
         g_oSysParam.TicketPrefixLen = m_szTicketPrefixLen
     End If
     If m_szTicketNumberLen <> CInt(Val(txtNumber.Text)) Then
         m_szTicketNumberLen = CInt(Val(txtNumber.Text))
         g_oSysParam.TicketNumberLen = m_szTicketNumberLen
     End If

     '/*��Ʊ*/
     If m_szScrollBusCanReturnTicket <> chkScorll_RT.Value Then
         If chkScorll_RT.Value = vbChecked Then
             m_szScrollBusCanReturnTicket = True
         Else
             m_szScrollBusCanReturnTicket = False
         End If
         g_oSysParam.ScrollBusCanReturnTicket = m_szScrollBusCanReturnTicket
     End If
     If m_szScrollBusReturnRatio <> Val(txtScorll_RT_Charge.Text) Then
         m_szScrollBusReturnRatio = Val(txtScorll_RT_Charge.Text)
         g_oSysParam.ScrollBusReturnRatio = m_szScrollBusReturnRatio
     End If
     If m_szScrollBusReturnTime <> Val(txtScrollBusTime_RT.Text) Then
         m_szScrollBusReturnTime = Val(txtScrollBusTime_RT.Text)
         g_oSysParam.ScrollBusReturnTime = m_szScrollBusReturnTime
     End If

    '/*��Ʊ*/
    'ʱ�����
     If m_szBeginCheckTime <> CInt(Val(txtChkStartTime.Text)) Then
         m_szBeginCheckTime = CInt(Val(txtChkStartTime.Text))
         g_oSysParam.BeginCheckTime = m_szBeginCheckTime
     End If
     If m_szCheckTicketTime <> CInt(Val(txtChkTime.Text)) Then
         m_szCheckTicketTime = CInt(Val(txtChkTime.Text))
         g_oSysParam.CheckTicketTime = m_szCheckTicketTime
     End If
     If m_szLatestExtraCheckTime <> CInt(Val(txtExChkStartTime.Text)) Then
         m_szLatestExtraCheckTime = CInt(Val(txtExChkStartTime.Text))
         g_oSysParam.LatestExtraCheckTime = m_szLatestExtraCheckTime
     End If
     If m_szExtraCheckTime <> CInt(Val(txtExChkTime.Text)) Then
         m_szExtraCheckTime = CInt(Val(txtExChkTime.Text))
         g_oSysParam.ExtraCheckTime = m_szExtraCheckTime
     End If
     If m_szAllowStartChectNotRearchTime <> chkAllowStartChectNotRearchTime.Value Then
         If chkAllowStartChectNotRearchTime.Value = vbChecked Then
             m_szAllowStartChectNotRearchTime = True
         Else
             m_szAllowStartChectNotRearchTime = False
         End If
         g_oSysParam.AllowStartChectNotRearchTime = m_szAllowStartChectNotRearchTime
     End If
     '�ĳ�
     If m_szAllowChangeRide <> chkChangeRide.Value Then
         If chkChangeRide.Value = vbChecked Then
             m_szAllowChangeRide = True
         Else
             m_szAllowChangeRide = False
         End If
         g_oSysParam.AllowChangeRide = m_szAllowChangeRide
     End If
     If m_szAllowChangeRideLowerPrice <> chkLowChangeRide.Value Then
         If chkLowChangeRide.Value = vbChecked Then
             m_szAllowChangeRideLowerPrice = True
         Else
             m_szAllowChangeRideLowerPrice = False
         End If
         g_oSysParam.AllowChangeRideLowerPrice = m_szAllowChangeRideLowerPrice
     End If
     '·��
     If m_szCheckSheetLen <> Val(txtRoldSheetNum.Text) Then
         m_szCheckSheetLen = Val(txtRoldSheetNum.Text)
         g_oSysParam.CheckSheetLen = m_szCheckSheetLen
     End If
     If m_szWantPrintRoadSheetID <> chkPrintSheetNum.Value Then
         If chkPrintSheetNum.Value = vbChecked Then
             m_szWantPrintRoadSheetID = True
         Else
             m_szWantPrintRoadSheetID = False
         End If
         g_oSysParam.WantPrintRoadSheetID = m_szWantPrintRoadSheetID
     End If
     If m_szWantTDPrintType <> chkTD.Value Then
         If chkTD.Value = vbChecked Then
             m_szWantTDPrintType = True
         Else
             m_szWantTDPrintType = False
         End If
         g_oSysParam.WantTDPrintType = m_szWantTDPrintType
     End If
     If m_szWantPrintRoadSheetTitle <> chkPrintSheetTitle.Value Then
         If chkPrintSheetTitle.Value = vbChecked Then
             m_szWantPrintRoadSheetTitle = True
         Else
             m_szWantPrintRoadSheetTitle = False
         End If
         g_oSysParam.WantPrintRoadSheetTitle = m_szWantPrintRoadSheetTitle
     End If
     If m_szRoadSheetTitle <> txtRoldTitle.Text Then
         m_szRoadSheetTitle = txtRoldTitle.Text
         g_oSysParam.RoadSheetTitle = m_szRoadSheetTitle
     End If

     '/*����*/
     If m_szFixFeeItem <> Trim(txtFixFeeItem.Text) Then
         m_szFixFeeItem = Trim(txtFixFeeItem.Text)
         g_oSysParam.FixFeeItem = m_szFixFeeItem
     End If
     If m_szIsFixFeeUpdateEachMonth <> chkFixFeeUpdateEachMonth.Value Then
         If chkFixFeeUpdateEachMonth.Value = vbChecked Then
             m_szIsFixFeeUpdateEachMonth = True
         Else
             m_szIsFixFeeUpdateEachMonth = False
         End If
         g_oSysParam.IsFixFeeUpdateEachMonth = m_szIsFixFeeUpdateEachMonth
     End If
     If m_szAllowSplitBySomeTimes <> chkAllowSplitBySomeTimes.Value Then
         If chkAllowSplitBySomeTimes.Value = vbChecked Then
             m_szAllowSplitBySomeTimes = True
         Else
             m_szAllowSplitBySomeTimes = False
         End If
         g_oSysParam.AllowSlitp = m_szAllowSplitBySomeTimes
     End If
     If m_szAllowSplitAboveFactQuantity <> chkAllowSplitAboveFactQuantity.Value Then
         If chkAllowSplitAboveFactQuantity.Value = vbChecked Then
             m_szAllowSplitAboveFactQuantity = True
         Else
             m_szAllowSplitAboveFactQuantity = False
         End If
         g_oSysParam.AllowSplitAboveFactQuantity = m_szAllowSplitAboveFactQuantity
     End If
     If m_szAllowSettleTotalNegative <> chkAllowSettleTotalNegative.Value Then
         If chkAllowSettleTotalNegative.Value = vbChecked Then
             m_szAllowSettleTotalNegative = True
         Else
             m_szAllowSettleTotalNegative = False
         End If
         g_oSysParam.AllowSettleTotalNegative = m_szAllowSettleTotalNegative
     End If
     If m_szSettleNegativeSplitItem <> Val(txtSettleNegativeSplitItem.Text) Then
         m_szSettleNegativeSplitItem = Val(txtSettleNegativeSplitItem.Text)
         g_oSysParam.SettleNegativeSplitItem = m_szSettleNegativeSplitItem
     End If
     If m_szServiceItemPosition <> Val(txtServiceItemPosition.Text) Then
         m_szServiceItemPosition = Val(txtServiceItemPosition.Text)
         g_oSysParam.ServiceItemPosition = m_szServiceItemPosition
     End If
         If m_szCarriageItemPosition <> Val(txtCarriageItemPosition.Text) Then
         m_szCarriageItemPosition = Val(txtCarriageItemPosition.Text)
         g_oSysParam.CarriageItemPosition = m_szCarriageItemPosition
     End If

     '/*����*/
     If m_szIsBookValid <> chkUseBook.Value Then
         If chkUseBook.Value = vbChecked Then
             m_szIsBookValid = True
         Else
             m_szIsBookValid = False
         End If
         g_oSysParam.IsBookValid = m_szIsBookValid
     End If
     If m_szBookTime <> CInt(Val(txtBook.Text)) Then
         m_szBookTime = CInt(Val(txtBook.Text))
         g_oSysParam.BookTime = m_szBookTime
     End If
     If m_szBookNumberLen <> CInt(Val(txtBookLong.Text)) Then
         m_szBookNumberLen = CInt(Val(txtBookLong.Text))
         g_oSysParam.BookNumberLen = m_szBookNumberLen
     End If
End Sub

Private Sub txtAfter_KeyPress(KeyAscii As Integer)
     If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = 46 Or KeyAscii = 8 Then
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtAfter_LostFocus()
    If txtAfter.Text = "" Then txtAfter.Text = 3
End Sub

Private Sub txtBefore_KeyPress(KeyAscii As Integer)
    If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = 46 Or KeyAscii = 8 Then
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtBefore_LostFocus()
    If txtBefore.Text = "" Then txtBefore.Text = 3
End Sub

Private Sub txtBook_KeyPress(KeyAscii As Integer)
    KeyAscii = NumberText(KeyAscii, txtBook.Text, txtBook.Seltext, txtBook.SelStart, False, False)
End Sub

Private Sub txtBook_Validate(Cancel As Boolean)
    If txtBook = "" Then
        MsgBox "Ԥ��ƱʧЧʱ�䳤�ȱ�����0��120֮�䣡", vbOKOnly + vbInformation, cszMsg
        txtBook.Text = "30"
        Exit Sub
    End If
    If CInt(txtBook.Text) > 120 Or CInt(txtBook.Text) < 0 Then
        MsgBox "Ԥ��ƱʧЧʱ�䳤�ȱ�����0��120֮�䣡", vbOKOnly + vbInformation, cszMsg
        txtBook.Text = "30"
    End If
End Sub

Private Sub txtCancelTicket_KeyPress(KeyAscii As Integer)
    KeyAscii = NumberText(KeyAscii, txtCancelTicket.Text, txtCancelTicket.Seltext, txtCancelTicket.SelStart, False, False)
End Sub

Private Sub txtChangeCharge_KeyPress(KeyAscii As Integer)
    KeyAscii = NumberText(KeyAscii, txtChangeCharge.Text, txtChangeCharge.Seltext, txtChangeCharge.SelStart, True, True)
End Sub

Private Sub txtChkStartTime_KeyPress(KeyAscii As Integer)
   KeyAscii = NumberText(KeyAscii, txtChkStartTime.Text, txtChkStartTime.Seltext, txtChkStartTime.SelStart, False, False)
End Sub

Private Sub txtChkTime_KeyPress(KeyAscii As Integer)
     KeyAscii = NumberText(KeyAscii, txtChkTime.Text, txtChkTime.Seltext, txtChkTime.SelStart, False, False)
End Sub

Private Sub txtExChkStartTime_KeyPress(KeyAscii As Integer)
    KeyAscii = NumberText(KeyAscii, txtExChkStartTime.Text, txtExChkStartTime.Seltext, txtExChkStartTime.SelStart, False, False)
End Sub

Private Sub txtExChkTime_KeyPress(KeyAscii As Integer)
    KeyAscii = NumberText(KeyAscii, txtExChkTime.Text, txtExChkTime.Seltext, txtExChkTime.SelStart, False, False)
End Sub

Private Sub TxtExtraCsRatio_KeyPress(KeyAscii As Integer)
    If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = 46 Or KeyAscii = 8 Then
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub TxtExtraLsRatio_KeyPress(KeyAscii As Integer)
    If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = 46 Or KeyAscii = 8 Then
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtHalfTicketTypeRatio_Change()
    FormatTextToNumeric txtHalfTicketTypeRatio, False, False
End Sub

Private Sub txtHalfTicketTypeRatio_LostFocus()
    If Val(txtHalfTicketTypeRatio.Text) > 100 Then
        MsgBox "��Ʊ���۰ٷֱȲ��ܴ���100%", vbInformation, "��ʾ"
        txtHalfTicketTypeRatio.SetFocus
    End If
End Sub

Private Sub txtInternetBusInfo_KeyPress(KeyAscii As Integer)
    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtNumber_KeyPress(KeyAscii As Integer)
    KeyAscii = NumberText(KeyAscii, txtNumber.Text, txtNumber.Seltext, txtNumber.SelStart, False, False)
End Sub

Private Sub txtPreFix_KeyPress(KeyAscii As Integer)
     KeyAscii = NumberText(KeyAscii, txtPreFix.Text, txtPreFix.Seltext, txtPreFix.SelStart, False, False)
End Sub

Private Sub txtPreFix_Change()
    If Val(txtPreFix.Text) > UpDown9.Max Then
         txtPreFix.Text = UpDown9.Max
    End If
End Sub

'Private Sub txtPreSale_Change()
'    If Val(txtPreSale.Text) > UpDown8.Max Then
'         txtPreSale.Text = UpDown8.Max
'    End If
'End Sub

Private Sub txtPreSale_KeyPress(KeyAscii As Integer)
    KeyAscii = NumberText(KeyAscii, txtPreSale.Text, txtPreSale.Seltext, txtPreSale.SelStart, False, False)
End Sub

Private Sub txtPriceDetailKeepBit_Change()
    If Val(txtPriceDetailKeepBit.Text) > UpDown18.Max Then
         txtPriceDetailKeepBit.Text = UpDown18.Max
    End If
End Sub

Private Sub txtPriceDetailKeepBit_KeyPress(KeyAscii As Integer)
  If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Then
     If KeyAscii = vbKeyBack Then Exit Sub
  Else
     KeyAscii = 0
  End If
End Sub

Private Sub txtRoldSheetNum_KeyPress(KeyAscii As Integer)
    KeyAscii = NumberText(KeyAscii, txtRoldSheetNum.Text, txtRoldSheetNum.Seltext, txtRoldSheetNum.SelStart, False, False)
End Sub

Private Sub txtScorll_RT_Charge_KeyPress(KeyAscii As Integer)
     KeyAscii = NumberText(KeyAscii, txtScorll_RT_Charge.Text, txtScorll_RT_Charge.Seltext, txtScorll_RT_Charge.SelStart)
End Sub

Private Sub txtScrollBusTime_RT_KeyPress(KeyAscii As Integer)
    If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = 8 Or (KeyAscii = 46 And InStr(txtScrollBusTime_RT.Text, ".") = 0) Then
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtScrollBusTime_RT_LostFocus()
   If txtScrollBusTime_RT.Text = "" Then
    MsgBox "������������ϳ�ʱ��β���Ϊ�գ�", vbInformation, "��ʾ��"
    txtScrollBusTime_RT.SetFocus
   End If
End Sub

Private Sub TxtSpeed_KeyPress(KeyAscii As Integer)
    KeyAscii = NumberText(KeyAscii, TxtSpeed.Text, TxtSpeed.Seltext, TxtSpeed.SelStart, True, False)
End Sub

Private Sub TxtSpeed_KeyUp(KeyCode As Integer, Shift As Integer)
    If Val(TxtSpeed.Text) > 500 Then TxtSpeed.Text = 500
End Sub

Private Sub txtStopSale_KeyPress(KeyAscii As Integer)
    KeyAscii = NumberText(KeyAscii, txtStopSale.Text, txtStopSale.Seltext, txtStopSale.SelStart, False, False)
End Sub

Private Sub UpDown8_change()
    If UpDown8.Value >= 11 - UpDown9.Value Then
        UpDown8.Value = 0
    End If
End Sub

Private Sub UpDown9_change()
    If UpDown9.Value >= 11 - UpDown8.Value Then
        UpDown9.Value = 0
    End If
End Sub

Private Sub DoUseBook()
    If chkUseBook.Value = vbChecked Then
        udBook.Enabled = True
        txtBook.Enabled = True
        txtBook.BackColor = vbWhite
        txtBookLong.Enabled = True
        txtBookLong.BackColor = vbWhite
    Else
        udBook.Enabled = False
        txtBook.Enabled = False
        txtBook.BackColor = cGreyColor
        txtBookLong.Enabled = False
        txtBookLong.BackColor = cGreyColor
    End If
End Sub

'������Ʊʱ���
Private Sub AddLine()
    vsReturnRatio.AddItem "��" & vsReturnRatio.Row + 1 & "ʱ���", vsReturnRatio.Row + 1
    ShowReturnTktNum
End Sub

'ɾ����Ʊʱ���
Private Sub DeleteLine()
    If vsReturnRatio.Rows > 1 Then
        If vsReturnRatio.Row <> 0 Then
            vsReturnRatio.RemoveItem vsReturnRatio.Row
            ShowReturnTktNum
        End If
    End If
End Sub

'��ʾ��Ʊʱ������
Private Sub ShowReturnTktNum()
    Dim iCount As Integer
    With vsReturnRatio
        If .Rows > 1 Then
            For iCount = 1 To .Rows - 1
                .TextMatrix(iCount, 0) = "��" & iCount & "ʱ���"
            Next iCount
        End If
    End With
End Sub

'�õ���Ʊ����
Private Sub GetReturnTktRatio()
    Dim rfTemp() As RETURNFEES
    Dim iLen, iArrayLength As Integer
    Dim iCount As Integer

    vsReturnRatio.Rows = 1
    iLen = ArrayLength(g_oSysParam.GetReturnFees)
    If iLen <> 0 Then
        ReDim rfTemp(1 To iLen)
        rfTemp = g_oSysParam.GetReturnFees
        For iCount = 1 To iLen
            vsReturnRatio.AddItem rfTemp(iCount).iReturnNum & vbTab & rfTemp(iCount).iReturnTime & vbTab & _
            rfTemp(iCount).sgReturnRate & vbTab & rfTemp(iCount).sgLeastMoney, iCount
        Next iCount
    End If
End Sub

'������Ʊ����
Private Sub SetReturnTktRatio()
    Dim rfTemp() As RETURNFEES
    Dim iCount As Integer
    DeleteEmptyLine
    If vsReturnRatio.Rows > 1 Then
        ReDim rfTemp(1 To vsReturnRatio.Rows - 1)
        With vsReturnRatio
            For iCount = 1 To vsReturnRatio.Rows - 1
                .TextMatrix(iCount, 1) = CInt(Val(.TextMatrix(iCount, 1)))
                rfTemp(iCount).iReturnNum = iCount
                rfTemp(iCount).iReturnTime = Val(.TextMatrix(iCount, 1))
                rfTemp(iCount).sgReturnRate = Val(.TextMatrix(iCount, 2))
                rfTemp(iCount).sgLeastMoney = Val(.TextMatrix(iCount, 3))
            Next iCount
            g_oSysParam.SetReturnFees .Rows - 1, rfTemp
        End With
    Else
        g_oSysParam.SetReturnFees 0, rfTemp
    End If
End Sub

'ȥ����Ʊ�����п���
Private Sub DeleteEmptyLine()
    Dim iCount As Integer
    If vsReturnRatio.Rows > 1 Then
        For iCount = 1 To vsReturnRatio.Rows - 1
            If vsReturnRatio.TextMatrix(iCount, 1) = "" Then
                vsReturnRatio.RemoveItem iCount
            End If
        Next iCount
    End If
    ShowReturnTktNum
End Sub

'��Ʊ���������ʽ����
Private Function InputControl(KeyAscii As Integer) As Boolean
    If (KeyAscii > vbKey9 Or KeyAscii < vbKey0) And KeyAscii <> 8 Then
        InputControl = True
    Else
        InputControl = False
    End If
End Function

Private Sub vsReturnRatio_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Select Case Col
        Case 1
            If InputControl(KeyAscii) And KeyAscii <> 45 Then KeyAscii = 0
        Case 2, 3
            If (Not InputControl(KeyAscii)) Or (KeyAscii = 46 And InStr(vsReturnRatio.EditText, ".") = 0) Then
            Else
                KeyAscii = 0
            End If
        Case Else
    End Select
End Sub
