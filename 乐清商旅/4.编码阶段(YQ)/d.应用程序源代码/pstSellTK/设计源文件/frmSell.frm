VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{695ABF14-B2D8-11D2-A5ED-DE08DCF33612}#1.2#0"; "asfcombo.ocx"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{BBF95DAA-F9CB-4CA9-A673-E0E9E0193752}#1.0#0"; "STSellCtl.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmSell 
   BackColor       =   &H00929292&
   Caption         =   "��Ʊ"
   ClientHeight    =   7635
   ClientLeft      =   -2955
   ClientTop       =   1890
   ClientWidth     =   11880
   HelpContextID   =   4000040
   Icon            =   "frmSell.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7635
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1245
      HelpContextID   =   3000411
      Left            =   4650
      TabIndex        =   51
      Top             =   7500
      Visible         =   0   'False
      Width           =   2880
      Begin VB.CheckBox chkSetSeat 
         BackColor       =   &H00FFFFFF&
         Caption         =   "����(&P)"
         Height          =   270
         HelpContextID   =   3000411
         Left            =   120
         TabIndex        =   52
         Top             =   -30
         Width           =   975
      End
      Begin VB.Label lblSetSeat 
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   180
         TabIndex        =   53
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
      TabIndex        =   15
      Top             =   120
      Width           =   11835
      Begin VB.ComboBox cboInsurance 
         Height          =   300
         Left            =   4680
         Style           =   2  'Dropdown List
         TabIndex        =   61
         Top             =   4980
         Width           =   1005
      End
      Begin VB.CommandButton cmdFSale 
         Caption         =   "ǿ��(&F)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         HelpContextID   =   3000411
         Left            =   10830
         TabIndex        =   58
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
         Left            =   7935
         TabIndex        =   21
         Top             =   45
         Width           =   1755
      End
      Begin VB.CommandButton cmdSetSeat 
         Caption         =   "����(&G)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         HelpContextID   =   3000411
         Left            =   9765
         TabIndex        =   22
         Top             =   75
         Width           =   990
      End
      Begin VB.TextBox txtReceivedMoney 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "����"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   1110
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
      Begin STSellCtl.ucNumTextBox txtTime 
         Height          =   330
         Left            =   5040
         TabIndex        =   19
         Top             =   90
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   582
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RTComctl3.CoolButton cmdSell 
         Height          =   405
         Left            =   120
         TabIndex        =   14
         Top             =   4860
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   714
         BTYPE           =   3
         TX              =   "�۳�(&P)"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
      Begin STSellCtl.ucNumTextBox txtPreferentialSell 
         Height          =   390
         Left            =   1410
         TabIndex        =   11
         Top             =   4230
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
         TabIndex        =   9
         Top             =   3810
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
         TabIndex        =   7
         Top             =   3375
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   688
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
            Name            =   "����"
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
         TabIndex        =   29
         Top             =   5280
         Width           =   5955
         _ExtentX        =   10504
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
            Name            =   "����"
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
               Picture         =   "frmSell.frx":0166
               Key             =   "StopBus"
            EndProperty
         EndProperty
      End
      Begin RTComctl3.FlatLabel flblSellDate 
         Height          =   315
         Left            =   930
         TabIndex        =   31
         Top             =   510
         Width           =   1590
         _ExtentX        =   2805
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
         Caption         =   "2006-05-27"
      End
      Begin MSComctlLib.ListView lvBus 
         Height          =   4470
         Left            =   2640
         TabIndex        =   5
         Top             =   510
         Width           =   9090
         _ExtentX        =   16034
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
         BackColor       =   16761087
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
         Left            =   8010
         TabIndex        =   32
         Top             =   4620
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
         TabIndex        =   33
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
         TabIndex        =   10
         Top             =   4260
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         ButtonDisabledForeColor=   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
      Begin FCmbo.asFlatCombo cboSeatType 
         Height          =   330
         Left            =   3600
         TabIndex        =   17
         Top             =   3210
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         ButtonDisabledForeColor=   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
         Left            =   8370
         TabIndex        =   34
         Top             =   4980
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
      Begin STSellCtl.ucSuperCombo cboEndStation 
         Height          =   1635
         Left            =   90
         TabIndex        =   3
         Top             =   1560
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   2884
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.ListView lvSellStation 
         Height          =   1635
         Left            =   2640
         TabIndex        =   25
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
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "�ϳ�վ"
            Object.Width           =   1942
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   " ʱ��"
            Object.Width           =   1696
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   " Ʊ��"
            Object.Width           =   1677
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "�ϳ�վ����"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "��Ʊ��"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.CommandButton cmdPreSell 
         Caption         =   "Ԥ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   9150
         TabIndex        =   35
         Top             =   3960
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CommandButton cmdForFocusFreeTicket 
         Caption         =   "����Ʊ֮��õ�����"
         Height          =   360
         Left            =   8880
         TabIndex        =   36
         Top             =   2490
         Width           =   615
      End
      Begin VB.CommandButton cmdForFocustxtReceiveMoney 
         Caption         =   "��ʵ��Ʊ��֮��õ�����"
         Height          =   345
         Left            =   8310
         TabIndex        =   37
         Top             =   1620
         Width           =   645
      End
      Begin VB.Frame fraDiscountTicket 
         Caption         =   "�ۿ�Ʊ"
         Height          =   585
         Left            =   6900
         TabIndex        =   24
         Top             =   5730
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
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1680
            TabIndex        =   26
            Text            =   "1"
            Top             =   180
            Width           =   1185
         End
         Begin VB.Label lblDiscount 
            Caption         =   "�ۿ�(&F):"
            Height          =   225
            Left            =   360
            TabIndex        =   28
            Top             =   225
            Width           =   795
         End
      End
      Begin VB.CommandButton cmdForFocustxtSeat 
         Caption         =   "����λ��֮��õ�����"
         Height          =   180
         Left            =   7950
         TabIndex        =   30
         Top             =   6090
         Visible         =   0   'False
         Width           =   975
      End
      Begin MSComCtl2.UpDown upFull 
         Height          =   390
         Left            =   2280
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   3375
         Width           =   255
         _ExtentX        =   423
         _ExtentY        =   688
         _Version        =   393216
         BuddyControl    =   "txtFullSell"
         BuddyDispid     =   196623
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
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   3810
         Width           =   255
         _ExtentX        =   423
         _ExtentY        =   688
         _Version        =   393216
         BuddyControl    =   "txtHalfSell"
         BuddyDispid     =   196622
         OrigLeft        =   2370
         OrigTop         =   3090
         OrigRight       =   2625
         OrigBottom      =   3405
         Max             =   100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   1745027080
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown ucPreferential 
         Height          =   390
         Left            =   2280
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   4230
         Width           =   255
         _ExtentX        =   423
         _ExtentY        =   688
         _Version        =   393216
         BuddyControl    =   "txtPreferentialSell"
         BuddyDispid     =   196621
         OrigLeft        =   2370
         OrigTop         =   3090
         OrigRight       =   2625
         OrigBottom      =   3405
         Max             =   100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   1745027080
         Enabled         =   -1  'True
      End
      Begin VB.Label lblInsurance 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����(F12):"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   3720
         TabIndex        =   62
         Top             =   5040
         Width           =   900
      End
      Begin VB.Label flblSellWeek 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
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
         Left            =   1950
         TabIndex        =   60
         Top             =   1230
         Width           =   540
      End
      Begin VB.Label flblChinaDate 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������(��)���³�һ"
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
         Left            =   930
         TabIndex        =   59
         Top             =   930
         Width           =   1560
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "��λ��(&T):"
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
         Left            =   6645
         TabIndex        =   20
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label lblsellstation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ϳ�վ(&O):"
         Height          =   180
         Left            =   2655
         TabIndex        =   23
         Top             =   5040
         Width           =   900
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���Ժ�"
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
         Left            =   5655
         TabIndex        =   50
         Top             =   120
         Width           =   720
      End
      Begin VB.Label lblTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ʱ��(&S):"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3450
         TabIndex        =   18
         Top             =   120
         Width           =   1440
      End
      Begin VB.Label flblRestMoney 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.0"
         BeginProperty Font 
            Name            =   "����"
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
         TabIndex        =   49
         Top             =   6450
         Width           =   720
      End
      Begin VB.Label flblTotalPrice 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.0"
         BeginProperty Font 
            Name            =   "����"
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
         TabIndex        =   48
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
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2880
         TabIndex        =   16
         Top             =   3255
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label lblHalfSell 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ʊ(&X):"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
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
         Top             =   3915
         Width           =   960
      End
      Begin VB.Label lblFullSell 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ȫƱ(&A):"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   6
         Top             =   3480
         Width           =   960
      End
      Begin VB.Label lblSellMsg 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   9780
         TabIndex        =   47
         Top             =   5010
         Width           =   1875
      End
      Begin VB.Label lblPreSell 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "���۳����б�:"
         Height          =   180
         Left            =   5745
         TabIndex        =   27
         Top             =   5040
         Width           =   1170
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
         TabIndex        =   46
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
         Caption         =   "��Ʊ��:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   8640
         TabIndex        =   45
         Top             =   2610
         Width           =   840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "վƱ:"
         Height          =   180
         Left            =   5190
         TabIndex        =   44
         Top             =   2880
         Width           =   705
      End
      Begin VB.Label Label1 
         Caption         =   "��������:"
         Height          =   225
         Left            =   300
         TabIndex        =   43
         Top             =   1650
         Width           =   1110
      End
      Begin VB.Label lblToStation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��վ(&Z):"
         BeginProperty Font 
            Name            =   "����"
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
         Top             =   1260
         Width           =   960
      End
      Begin VB.Label lblPrevDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ԥ������(&D):"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "�����б�(&V):"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1890
         TabIndex        =   4
         Top             =   135
         Width           =   1440
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "��Ʊ����:"
         Height          =   180
         Left            =   5970
         TabIndex        =   42
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
         TabIndex        =   41
         Top             =   2130
         Width           =   1215
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C0C0C0&
         Visible         =   0   'False
         X1              =   30
         X2              =   2670
         Y1              =   4740
         Y2              =   4740
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ʊ��:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   90
         TabIndex        =   40
         Top             =   5340
         Width           =   840
      End
      Begin VB.Label lblReceivedMoney 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ʵ��(&Q):"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   90
         TabIndex        =   13
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
         Caption         =   "Ӧ��Ʊ��:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   90
         TabIndex        =   39
         Top             =   6540
         Width           =   1080
      End
      Begin VB.Label lblmileate 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "����"
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
         TabIndex        =   38
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
      TabIndex        =   54
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
'# Last Modify:2005-12-2 By ½����
'# Last Modify At:�����˶�ȫ��������֧�֣�����Ϊ��
'#                a.������upFull��upHalf��upPreferential�ؼ�
'#                b.����lvBus_DblClick��֧�֣��Լ�������upFull��upHalf��upPreferential�ؼ�
'#                c.���ؼ���������������Ʊ�б���û��ʱ�����Զ����ۿؼ��������ĳ�Ʊ
'#                d.�������������������һ��������Ʊ��ť��ǰ����ص�txtReceviMoney.setfocus���ģ����������˵Ĳ���ϰ�ߣ�
'#                e.�ָ���ͣ�೵�ε��г��������б��
'*******************************************************************
Private m_bPrint As Boolean
Private m_blPointCount As Boolean
Private m_bSumPriceIsEmpty As Boolean   '��Ʊ��Ϊ0
Private m_nCount As Integer '��һ��ʱ���ȡ������ʱ����Լ�һ�ı���
Private m_sgTotalMoney As Single '��¼��һ����Ʊ�Ľ��
Private m_atTicketType() As TTicketType
Private m_dbTotalPrice As Double
Private m_aszSeatType() As String
Private m_atbSeatTypeBus As TMultiSeatTypeBus '�õ���ͬ��λ���͵ĳ���
Private m_TicketPrice() As Single '�洢Ʊ��
Private m_TicketTypeDetail() As ETicketType '�洢Ʊ��
Private m_bPreClear As Boolean
Private m_bSetFocus As Boolean
Private m_bPreSellFocus As Boolean
Private m_rsBusInfo As Recordset
Private m_atbBusOrder() As TBusOrderCount

Private m_bNotRefresh As Boolean '�Ƿ���Ҫˢ��,��Ҫ�������ò�ѯ����ʱ��ʱ�õ�.
Private rsCountTemp As Recordset
Private nSellCount As Integer '��������
Private m_szStatus As Boolean
Private nTicketCount As Integer
Private nKey As Integer

Private m_aszInsurce() As String '������

Private szSeatStatus As String

Private m_aszRealNameInfo() As TCardInfo


Private Sub cboEndStation_Change()
On Error GoTo here
'    SetUser m_oAUser.UserID
    If lvBus.ListItems.count > 0 Then
        DoThingWhenBusChange
    Else
       lvSellStation.ListItems.Clear
    End If
    'fpd
    cboInsurance.ListIndex = 0
    DealPrice
    
    
    cmdPreSell.Enabled = True
On Error GoTo 0
Exit Sub
here:
  ShowErrorMsg
End Sub

Private Sub cboEndStation_GotFocus()
On Error GoTo here
    lblToStation.ForeColor = clActiveColor
'    DealPrice
    
    '********************
    '������ʾ��
'        SetUser m_oAUser.UserID
        'SetWhere
    '********************
On Error GoTo 0
Exit Sub
here:
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
        Case 106 'Asc("*")
            '+������������ʱ�䴦
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
On Error GoTo here
Dim nIndex As Integer
    lblToStation.ForeColor = 0
    If m_bNotRefresh Then Exit Sub '���������������ʱ�䴦,��ˢ��
    
    DoThingWhenBusChange
    SetDefaultSellTicket
    txtReceivedMoney.Text = ""
    RefreshBus True
    
'    If IsHaveScrollBus Then  '�ж��Ƿ��й�������
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
    If cboEndStation.Text <> "" Then
        SetInsurance
    End If
    
On Error GoTo 0
Exit Sub
here:
    ShowErrorMsg
End Sub


Private Sub cboInsurance_Click()
    DealPrice
    If Me.ActiveControl Is txtReceivedMoney And lvPreSell.ListItems.count <> 0 Then
        SetPay2 flblTotalPrice.Caption
    End If
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

'Private Sub chkInsurance_Click()
'    DealPrice
'    If Me.ActiveControl Is txtReceivedMoney And lvPreSell.ListItems.count <> 0 Then
'        SetPay2 flblTotalPrice.Caption
'    End If
'End Sub



Private Sub cmdPreSell_Click()
    On Error GoTo here
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
here:
    ShowErrorMsg
End Sub

'��Ʊ
Private Sub cmdSell_Click()
Dim k As Long
Dim m As Long
Dim i As Integer
'�������˵�������¼����������Ʊ������δ��lvPreSell�����ּ�¼�������ʹ������ֱ���������ֺ�����ӡ��ť����Ʊ
    If lvPreSell.ListItems.count = 0 Then
        Call txtFullSell_KeyPress(vbKeyReturn)
    End If
    m = 0
    For i = 1 To lvPreSell.ListItems.count
        m = m + lvPreSell.ListItems(i).SubItems(IT_SumTicketNum)
    Next i
    
    If m_lEndTicketNoOld = 0 Then
        ShowMsg "��Ʊ���ɹ����û���δ��Ʊ������ȥ��Ʊ��"
        Exit Sub
    End If
'    If Int(m_lEndTicketNoOld) < Val(MDISellTicket.lblTicketNo.Caption) Then
'        ShowMsg "��Ʊû��Ʊ֤�ǼǷ�Χ,����Ʊ!"
'        Exit Sub
'    End If
    If m + Val(txtFullSell.Text) + Val(txtHalfSell.Text) + Val(txtPreferentialSell.Text) + Val(m_lTicketNo) - 1 > Val(m_lEndTicketNo) Then
        k = Val(m_lEndTicketNo) - Val(m_lTicketNo) + 1
        MsgBox "��ӡ���ϵ�Ʊ�Ѳ�����" & vbCrLf & "��Ʊֻʣ " & k & "��", vbInformation, "��Ʊ̨"
    Else
        SellTicket
    End If
End Sub

Private Sub cmdSetSeat_Click()
On Error GoTo here
    Dim rsTemp As Recordset
    szSeatStatus = ""
    frmOrderSeats.m_szStatus = True
    If lvBus.SelectedItem Is Nothing Then
        Set rsTemp = Nothing
        Exit Sub
    End If
    Set rsTemp = m_oSell.GetSeatRs(CDate(flblSellDate.Caption), lvBus.SelectedItem.Text)
    Set frmOrderSeats.m_rsSeat = rsTemp
    Set rsTemp = m_oSell.GetBookRs(CDate(flblSellDate.Caption), lvBus.SelectedItem.Text)
    Set frmOrderSeats.m_rsBook = rsTemp
    frmOrderSeats.m_szSeatNumber = PreOrderSeat
    frmOrderSeats.Show vbModal
    If frmOrderSeats.m_bOk Then
        txtSeat.Text = frmOrderSeats.m_szSeat
        szSeatStatus = frmOrderSeats.m_szSeatStatus
    End If
    Set rsTemp = Nothing
On Error GoTo 0
Exit Sub
here:
    Set rsTemp = Nothing
    ShowErrorMsg
End Sub

'ǿ��
Private Sub cmdFSale_Click()
On Error GoTo here
    Dim rsTemp As Recordset
    frmOrderSeats.m_szStatus = False
    If lvBus.SelectedItem Is Nothing Then
        Set rsTemp = Nothing
        Exit Sub
    End If
    Set rsTemp = m_oSell.GetSeatRs(CDate(flblSellDate.Caption), lvBus.SelectedItem.Text)
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
here:
    Set rsTemp = Nothing
    ShowErrorMsg
End Sub



Private Sub Command1_Click()
    frmNotify.Show vbModal
End Sub

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
On Error GoTo here
    m_nCurrentTask = RT_SellTicket
    m_szCurrentUnitID = Me.Tag
    SetPreSellButton
    MDISellTicket.SetFunAndUnit
'    m_oSell.SellUnitCode = Me.Tag
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
here:
    ShowErrorMsg
End Sub

Private Sub Form_Deactivate()
On Error GoTo here
    MDISellTicket.EnableSortAndRefresh False
'    'MDISellTicket.abMenuTool.Bands("mnuFunction").Tools("mnuChangeSeatType").Enabled = False
On Error GoTo 0
Exit Sub
here:
    ShowErrorMsg
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    DealWithChildKeyDown KeyCode, Shift
'    If KeyCode = vbKeyF9 Then
'        ChangeSeatType
'    End If
    If KeyCode = vbKeyF8 Then
        '����λ
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
        '�ָĳɰ�һ��F12Ϊ1Ԫ��������Ϊ2Ԫ zyw 2008-01-07
        If cboInsurance.ListIndex < cboInsurance.ListCount - 1 Then
            cboInsurance.ListIndex = cboInsurance.ListIndex + 1
        Else
            cboInsurance.ListIndex = 0
        End If
        
    ElseIf KeyCode = vbKeyF12 Then
        'ǿ��
        If cmdFSale.Enabled = True Then
            cmdFSale_Click
        End If
        
    ElseIf KeyCode = 219 Then '��������
        SetQueue
        
    ElseIf KeyCode = 221 Then '��������
        SetProtect
        
    End If
    
    If g_bIsUseInsurance Then
        If KeyCode = vbKeyF9 And Shift = 0 Then 'F9
            '�۱���
            If ArrayLength(m_aszInsurce) <> 0 Then
                g_oCommDialog.InitInfo m_oAUser.UserID, m_oAUser.UserName, m_oAUser.SellStationID
                g_oCommDialog.PrintInsurance m_oAUser.UserID, m_aszInsurce
                Dim aszNull() As String
                m_aszInsurce = aszNull
            Else
                MsgBox "û�п�Ʊ��Ϣ�������۱���", vbOKOnly + vbExclamation, "����"
                cboEndStation.SetFocus
                Exit Sub
            End If
        End If
        If KeyCode = vbKeyF9 And Shift = 2 Then 'Ctrl+F7
            '������
            g_oCommDialog.InitInfo m_oAUser.UserID, m_oAUser.UserName, m_oAUser.SellStationID
            g_oCommDialog.RecruitInsurance m_oAUser.UserID
        End If
        If KeyCode = vbKeyF11 And Shift = 0 Then 'F11
            '�����˱���
            If MsgBox("�Ƿ�����˱���", vbInformation + vbYesNo, "�����˱�") = vbYes Then
                g_oCommDialog.InitInfo m_oAUser.UserID, m_oAUser.UserName, m_oAUser.SellStationID
                Dim bIsReturned As Boolean
                bIsReturned = g_oCommDialog.FastReturnInsurance(m_oAUser.UserID)
                If bIsReturned = True Then
                    MsgBox "�����˱��ɹ���", vbInformation, "�����˱�"
                    cboEndStation.SetFocus
                Else
                    MsgBox "�����˱�ʧ�ܣ�", vbInformation, "�����˱�"
                    cboEndStation.SetFocus
                End If
            End If
        End If
        If KeyCode = vbKeyF11 And Shift = 2 Then 'Ctrl+F11
            '���������˱���
            g_oCommDialog.InitInfo m_oAUser.UserID, m_oAUser.UserName, m_oAUser.SellStationID
            g_oCommDialog.ReturnInsurance m_oAUser.UserID
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
        Erase m_aszRealNameInfo
        
        cboEndStation.SetFocus
        
        SetUser m_oAUser.UserID
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
    nKey = KeyAscii
    
End Sub


Private Sub Form_Load()
On Error GoTo here
    m_szStatus = m_oSell.IsForceSell

    flblSellDate.Caption = ToStandardDateStr(m_oParam.NowDate)
    flblChinaDate.Caption = ResolveDisplay(GetChinaDate(flblSellDate.Caption))
    flblSellWeek.Caption = ResolveDisplayEx(GetChinaDate(flblSellDate.Caption))
    txtPrevDate.Text = 0
    txtTime.Text = 0
    m_dbTotalPrice = 0
    m_bPrint = False
    RefreshPreferentialTicket '��ȡ�Ż�Ʊ��Ϣ
    GetPreSellBus  '��ʾԤ��״̬��Ϣ
    RefreshStation2
    
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
    '�����б�ͷ
    AlignHeadWidth Me.name, lvPreSell
'    AlignHeadWidth Me.name, lvBus
'    AlignHeadWidth Me.name, lvSellStation

    '��ʼ��������Ϣ�б� zyw 2008-01-07
    cboInsurance.AddItem "�ޱ���"
    cboInsurance.AddItem "1Ԫ"
    cboInsurance.AddItem "2Ԫ"
    cboInsurance.AddItem "3Ԫ"
    cboInsurance.AddItem "5Ԫ"
    cboInsurance.ListIndex = 0
    

SetUser m_oAUser.UserID

Exit Sub
here:
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
On Error GoTo here
    
    
    SaveHeadWidth Me.name, lvPreSell
'    SaveHeadWidth Me.name, lvBus
'    SaveHeadWidth Me.name, lvSellStation
    m_clSell.Remove GetEncodedKey(Me.Tag)
'    MDISellTicket.lblSell.Value = vbUnchecked
    MDISellTicket.abMenuTool.Bands("mnuFunction").Tools("mnuSellTkt").Checked = False
    MDISellTicket.EnableSortAndRefresh False
    
    
        '***************
        '������ʾ��
        WriteReg cszComPort, CStr(g_lComPort)
    
On Error GoTo 0
Exit Sub
here:
    ShowErrorMsg
End Sub




Private Sub Label4_Click()

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
    If lvBus.ListItems.count = 0 Then
        cboEndStation.SetFocus
        Exit Sub
    End If
'
'    '********************
'    '������ʾ��
'        SetUser m_oAUser.UserID
'        'SetWhere
'    '********************
End Sub

Private Sub lvBus_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo here
      
        RefreshSellStation rsCountTemp
        ShowRightSeatType
        DoThingWhenBusChange
        DealPrice
        
On Error GoTo 0
Exit Sub
here:
    ShowErrorMsg
End Sub

'��ǰѡ�еĳ��θı���Ҫ������
'����ȫ/��Ʊ��
'��ʾ�˳��ε�ָ����վ������ʱ�估��������
'����Ʊ��
'����վƱ��ѡ����ť��״̬
'������Ʊ��ť��״̬
Private Sub DoThingWhenBusChange()
    On Error GoTo here
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
'    DealPrice   ' ����Ʊ��
    EnableSeatAndStand  '����վƱ��ѡ����ť��״̬
    EnableSellButton    '������Ʊ��ť״̬
    On Error GoTo 0
    Exit Sub
here:
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

   lblsellstation.ForeColor = clActiveColor
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
On Error GoTo here
   If KeyCode = 13 Then
     txtFullSell.SetFocus
   End If
Exit Sub
here:
    ShowErrorMsg
  
End Sub

Private Sub lvSellStation_LostFocus()
    lblsellstation.ForeColor = 0
End Sub

Private Sub Timer1_Timer()
'    RefreshBusSeats True
    On Error GoTo here
    '��40��ȡһ�·�����ʱ��
    If m_nCount Mod 20 = 0 Then
        Date = m_oParam.NowDate
        Time = m_oParam.NowDateTime
        m_nCount = 0
    End If
    m_nCount = m_nCount + 1
    Exit Sub
    On Error GoTo 0
here:
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
On Error GoTo here
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
            MsgBox "�ۿ��ʲ��ܴ��� 1", vbInformation, "��ʾ"
            txtDiscount.SetFocus
        End If
    Else
        txtDiscount.Text = 1
    End If
    DealPrice
    
On Error GoTo 0
Exit Sub
here:
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
On Error GoTo here
    If KeyAscii = 13 Then
        txtFullSell.Text = CInt(txtFullSell.Text)
        If txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text <> 0 And Not (lvBus.SelectedItem Is Nothing) Then
            If lvBus.SelectedItem.SubItems(ID_OffTime) <> cszScrollBus Then
'                If txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text > lvBus.SelectedItem.SubItems(ID_SeatCount) Then
'                    MsgBox "�ó�����λ�Ѳ�����" & vbCrLf & "��λֻʣ " & lvBus.SelectedItem.SubItems(ID_SeatCount), vbInformation, "��Ʊ̨"
'                ElseIf cboSeatType.ListIndex = 0 And txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text > lvBus.SelectedItem.SubItems(ID_SeatTypeCount) Then
'                    MsgBox "�ó���[��ͨ]��λ�Ѳ�����" & vbCrLf & "  [��ͨ]��λֻʣ " & lvBus.SelectedItem.SubItems(ID_SeatTypeCount), vbInformation, "��Ʊ̨"
'                ElseIf cboSeatType.ListIndex = 1 And txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text > lvBus.SelectedItem.SubItems(ID_BedTypeCount) Then
'                    MsgBox "�ó���[����]��λ�Ѳ�����" & vbCrLf & "  [����]��λֻʣ " & lvBus.SelectedItem.SubItems(ID_BedTypeCount), vbInformation, "��Ʊ̨"
'                ElseIf cboSeatType.ListIndex = 2 And txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text > lvBus.SelectedItem.SubItems(ID_AdditionalCount) Then
'                    MsgBox "�ó���[����]��λ�Ѳ�����" & vbCrLf & "  [����]��λֻʣ " & lvBus.SelectedItem.SubItems(ID_AdditionalCount), vbInformation, "��Ʊ̨"
'                    txtFullSell.SetFocus
'                    KeyAscii = 0
'                Else
                    cmdPreSell_Click
                    txtReceivedMoney.SetFocus
'                    cmdSell.SetFocus
'                End If
            Else
                cmdPreSell_Click
                txtReceivedMoney.SetFocus
'                cmdSell.SetFocus
            End If
        End If
    End If
    
On Error GoTo 0
Exit Sub
here:
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
On Error GoTo here
    If KeyAscii = 13 Then
        txtFullSell.Text = CInt(txtFullSell.Text)
        If txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text <> 0 And Not (lvBus.SelectedItem Is Nothing) Then
            If lvBus.SelectedItem.SubItems(ID_OffTime) <> cszScrollBus Then
'                If txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text > lvBus.SelectedItem.SubItems(ID_SeatCount) Then
'                    MsgBox "�ó�����λ�Ѳ�����" & vbCrLf & "��λֻʣ " & lvBus.SelectedItem.SubItems(ID_SeatCount), vbInformation, "��Ʊ̨"
'                 ElseIf cboSeatType.ListIndex = 0 And txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text > lvBus.SelectedItem.SubItems(ID_SeatTypeCount) Then
'                    MsgBox "�ó���[��ͨ]��λ�Ѳ�����" & vbCrLf & "  [��ͨ]��λֻʣ " & lvBus.SelectedItem.SubItems(ID_SeatTypeCount), vbInformation, "��Ʊ̨"
'                 ElseIf cboSeatType.ListIndex = 1 And txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text > lvBus.SelectedItem.SubItems(ID_BedTypeCount) Then
'                    MsgBox "�ó���[����]��λ�Ѳ�����" & vbCrLf & "  [����]��λֻʣ " & lvBus.SelectedItem.SubItems(ID_BedTypeCount), vbInformation, "��Ʊ̨"
'                 ElseIf cboSeatType.ListIndex = 2 And txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text > lvBus.SelectedItem.SubItems(ID_AdditionalCount) Then
'                    MsgBox "�ó���[����]��λ�Ѳ�����" & vbCrLf & "  [����]��λֻʣ " & lvBus.SelectedItem.SubItems(ID_AdditionalCount), vbInformation, "��Ʊ̨"
'                    txtHalfSell.SetFocus
'                    KeyAscii = 0
'                Else
                    cmdPreSell_Click
                    txtReceivedMoney.SetFocus
'                    cmdSell.SetFocus
'                End If
            Else
                cmdPreSell_Click
'                cmdSell.SetFocus
                txtReceivedMoney.SetFocus
            End If
        End If
    End If
    
On Error GoTo 0
Exit Sub
here:
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
On Error GoTo here
    If KeyAscii = 13 Then
        If txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text <> 0 And Not (lvBus.SelectedItem Is Nothing) Then
            If lvBus.SelectedItem.SubItems(ID_OffTime) <> cszScrollBus Then
'                If txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text > lvBus.SelectedItem.SubItems(ID_SeatCount) Then
'                    MsgBox "�ó�����λ�Ѳ�����" & vbCrLf & "��λֻʣ " & lvBus.SelectedItem.SubItems(ID_SeatCount), vbInformation, "��Ʊ̨"
'                 ElseIf cboSeatType.ListIndex = 0 And txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text > lvBus.SelectedItem.SubItems(ID_SeatTypeCount) Then
'                    MsgBox "�ó���[��ͨ]��λ�Ѳ�����" & vbCrLf & "  [��ͨ]��λֻʣ " & lvBus.SelectedItem.SubItems(ID_SeatTypeCount), vbInformation, "��Ʊ̨"
'                 ElseIf cboSeatType.ListIndex = 1 And txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text > lvBus.SelectedItem.SubItems(ID_BedTypeCount) Then
'                    MsgBox "�ó���[����]��λ�Ѳ�����" & vbCrLf & "  [����]��λֻʣ " & lvBus.SelectedItem.SubItems(ID_BedTypeCount), vbInformation, "��Ʊ̨"
'                 ElseIf cboSeatType.ListIndex = 2 And txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text > lvBus.SelectedItem.SubItems(ID_AdditionalCount) Then
'                    MsgBox "�ó���[����]��λ�Ѳ�����" & vbCrLf & "  [����]��λֻʣ " & lvBus.SelectedItem.SubItems(ID_AdditionalCount), vbInformation, "��Ʊ̨"
'                    txtPreferentialSell.SetFocus
'                    KeyAscii = 0
'                Else
                    cmdPreSell_Click
                   'cmdPreSell.SetFocus
    
                      txtReceivedMoney.SetFocus
'                    cmdSell.SetFocus
'                End If
            Else
                cmdPreSell_Click
'                cmdSell.SetFocus
                txtReceivedMoney.SetFocus
            End If
        End If
        
    End If
    
On Error GoTo 0
Exit Sub
here:
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
    flblSellDate.Caption = ToStandardDateStr(DateAdd("d", txtPrevDate.Text, m_oParam.NowDate))
    flblChinaDate.Caption = ResolveDisplay(GetChinaDate(flblSellDate.Caption))
    flblSellWeek.Caption = ResolveDisplayEx(GetChinaDate(flblSellDate.Caption))
    
End Sub

'����Ʊ�ۣ�������Ʊ�ۣ������Ǯ)
Private Sub DealPrice()
    Dim i As Integer
    Dim sgTemp As Double  '������Ʊ��ֵ
    Dim sgvalue As Double
    Dim dbSum As Double
    Dim aszSeatNo() As String
    Dim nSeat As Integer
    Dim nLength As Integer
    nLength = 0
    sgTemp = 0
On Error GoTo here
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
here:
    ShowErrorMsg
End Sub

Private Sub txtPreferentialSell_Change()
On Error GoTo here
    
    
    
    EnableSellButton
    EnableSeatAndStand
    DealPrice
    SetPreSellButton
On Error GoTo 0
Exit Sub
here:
    ShowErrorMsg
End Sub

Private Sub txtFullSell_Change()
On Error GoTo here
    
    
    EnableSellButton
    EnableSeatAndStand
    DealPrice
    SetPreSellButton
On Error GoTo 0
Exit Sub
here:
    ShowErrorMsg
End Sub

Private Sub txtHalfSell_Change()
On Error GoTo here
    EnableSellButton
    EnableSeatAndStand
    DealPrice
    
    SetPreSellButton
On Error GoTo 0
Exit Sub
here:
    ShowErrorMsg
End Sub


Private Sub txtPrevDate_GotFocus()
    lblPrevDate.ForeColor = clActiveColor
'    '********************
'    '������ʾ��
'        SetUser m_oAUser.UserID
'    '********************
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
On Error GoTo here
    lblPrevDate.ForeColor = 0
    If txtPrevDate.Text = "" Then txtPrevDate.Text = 0
    

    RefreshBus True
On Error GoTo 0
Exit Sub
here:
    ShowErrorMsg
End Sub

Private Sub txtReceivedMoney_Change()
    Dim sgvalue As Double
On Error GoTo here
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
here:
    ShowErrorMsg
End Sub


'���ݵ�ǰ��Ϣ������Ʊ��ť��״̬
Private Sub EnableSellButton()
    Dim szStationID As String
    szStationID = cboEndStation.BoundText
    If szStationID = "" Or lvBus.SelectedItem Is Nothing Then
        cmdSell.Enabled = False
    Else
        cmdSell.Enabled = True
    End If
End Sub

'���ݵ�ǰ����Ϣ����վƱCheck��ť�Ͷ�����ť��״̬
Private Sub EnableSeatAndStand()
    Dim szStationID As String
    szStationID = cboEndStation.BoundText
        If szStationID = "" Or lvBus.SelectedItem Is Nothing Then  '��ǰ�޳���
            cmdSetSeat.Enabled = False
            cmdFSale.Enabled = False
'            chkSetSeat.Value = 0
'            chkSetSeat.Enabled = False
        Else
            Dim liTemp As ListItem
            Set liTemp = lvBus.SelectedItem
            
            If liTemp.SubItems(ID_BusType1) = TP_ScrollBus Then  '����ˮ���εĻ�������վƱ������
                cmdSetSeat.Enabled = False
                cmdFSale.Enabled = False
'                chkSetSeat.Value = 0
'                chkSetSeat.Enabled = False
            Else
                If liTemp.SubItems(ID_SeatCount) > 0 Then '������λ������0
                    If (txtFullSell.Text = 0 And txtHalfSell.Text = 0 And txtPreferentialSell.Text = 0) Then
                        cmdSetSeat.Enabled = False
                        cmdFSale.Enabled = False
'                        chkSetSeat.Value = 0
'                        chkSetSeat.Enabled = False
                    Else
                        cmdSetSeat.Enabled = True
                        If m_szStatus = True Then
                            cmdFSale.Enabled = True
                        Else
                            cmdFSale.Enabled = False
                        End If
'                        chkSetSeat.Enabled = True
                    End If
                    
                Else '�޿�����λ��������ʱ����վƱ�϶�����0����Ȼ�Ͳ��Ὣ�˳��β������
                    cmdSetSeat.Enabled = True
                    If m_szStatus = True Then
                        cmdFSale.Enabled = True
                    Else
                        cmdFSale.Enabled = False
                    End If
'                    chkSetSeat.Enabled = True
'                    chkSetSeat.Value = 0
                End If
                
            End If
        End If

End Sub

'����ȱʡ����Ʊ״̬
Private Sub SetDefaultSellTicket()
    txtFullSell.Text = 1 '��ȫƱ����Ϊ1
    txtHalfSell.Text = 0 '�۰�Ʊ����Ϊ0
    txtPreferentialSell.Text = 0 '����Ʊ����Ϊ0
'    txtRightSell.Value = 0 '����Ʊ����Ϊ0
    If chkSetSeat.Enabled Then chkSetSeat.Value = 0 '������λ
'    If chkStandSeat.Enabled Then chkStandSeat.Value = 0 '����վƱ
    
    If txtReceivedMoney.Enabled Then  '���ճ�ƱΪ0
        'txtReceivedMoney.Text = 0
'        DealPrice
    End If
    cboInsurance.ListIndex = 0
End Sub

'ˢ��ĳ���ε��ϳ�վ��Ϣ
'lvSellStation �� 1 �ϳ�վ 2 ����ʱ�� 3 ȫ�� 4 �ϳ�վ����

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
        If m_oAUser.SellStationID = Trim(rsTemp!sell_station_id) Then '��,Ӧ����ѡ���ϳ�վʱ��ƱԱѡ����վ,�Դ�ҪƱ�ͼ�����
'            lvSellStation.ListItems.Clear
            szTemp = Trim(rsTemp!sell_station_id)
            Set lvS = lvSellStation.ListItems.Add(, , Trim(rsTemp!sell_station_name))
            nBusType = rsTemp!bus_type
            If nBusType <> TP_ScrollBus Then
               lvS.SubItems(1) = Trim(Format(rsTemp!BusStartTime, "hh:mm"))
            Else
               lvS.SubItems(1) = cszScrollBus
            End If
            lvS.SubItems(3) = Trim(rsTemp!sell_station_id)
            lvS.SubItems(2) = Trim(rsTemp!full_price)
            lvS.SubItems(4) = Trim(rsTemp!sell_check_gate_id)
            GoTo begin
'            End If
        ElseIf m_oAUser.SellStationID = "km" And Trim(rsTemp!sell_station_id) = "cm" Then
            GoTo begin
        Else
            szTemp = Trim(rsTemp!sell_station_id)
            Set lvS = lvSellStation.ListItems.Add(, , Trim(rsTemp!sell_station_name))
            nBusType = rsTemp!bus_type
            If nBusType <> TP_ScrollBus Then
               lvS.SubItems(1) = Trim(Format(rsTemp!BusStartTime, "hh:mm"))
            Else
               lvS.SubItems(1) = cszScrollBus
            End If
            lvS.SubItems(3) = Trim(rsTemp!sell_station_id)
            lvS.SubItems(2) = Trim(rsTemp!full_price)
            lvS.SubItems(4) = Trim(rsTemp!sell_check_gate_id)
        End If
      End If
     End If
begin:
        rsTemp.MoveNext
    Next i
    
    '��
    For i = 1 To lvSellStation.ListItems.count
        If m_oAUser.SellStationID = lvSellStation.ListItems(i).SubItems(3) Then

            lvSellStation.ListItems(i).Selected = True
            lvSellStation.ListItems(i).EnsureVisible
            cboSeatType_Change
            Exit For
        End If
    Next i
    Exit Sub
err:
   MsgBox err.Description
End Sub

'��λ���͸ı�ʱ,ˢ����Ӧ��Ʊ��
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
            lvBus.SelectedItem.SubItems(ID_FullPrice) = FormatMoney(rsTemp!full_price)
            lvBus.SelectedItem.SubItems(ID_HalfPrice) = FormatMoney(rsTemp!half_price)
            lvBus.SelectedItem.SubItems(ID_PreferentialPrice1) = FormatMoney(rsTemp!preferential_ticket1)
            lvBus.SelectedItem.SubItems(ID_PreferentialPrice2) = FormatMoney(rsTemp!preferential_ticket2)
            lvBus.SelectedItem.SubItems(ID_PreferentialPrice3) = FormatMoney(rsTemp!preferential_ticket3)
        Case cszBedType
            lvBus.SelectedItem.SubItems(ID_BedFullPrice) = FormatMoney(rsTemp!full_price)
            lvBus.SelectedItem.SubItems(ID_BedHalfPrice) = FormatMoney(rsTemp!half_price)
            lvBus.SelectedItem.SubItems(ID_BedPreferentialPrice1) = FormatMoney(rsTemp!preferential_ticket1)
            lvBus.SelectedItem.SubItems(ID_BedPreferentialPrice2) = FormatMoney(rsTemp!preferential_ticket2)
            lvBus.SelectedItem.SubItems(ID_BedPreferentialPrice3) = FormatMoney(rsTemp!preferential_ticket3)
        Case cszAdditionalType
            lvBus.SelectedItem.SubItems(ID_AdditionalFullPrice) = FormatMoney(rsTemp!full_price)
            lvBus.SelectedItem.SubItems(ID_AdditionalHalfPrice) = FormatMoney(rsTemp!half_price)
            lvBus.SelectedItem.SubItems(ID_AdditionalPreferential1) = FormatMoney(rsTemp!preferential_ticket1)
            lvBus.SelectedItem.SubItems(ID_AdditionalPreferential2) = FormatMoney(rsTemp!preferential_ticket2)
            lvBus.SelectedItem.SubItems(ID_AdditionalPreferential3) = FormatMoney(rsTemp!preferential_ticket3)
        End Select
        lvBus.SelectedItem.SubItems(ID_CheckGate) = Trim(rsTemp!check_gate_id)
        End If
      End If
     End If
        rsTemp.MoveNext
    Next i
        lvBus.SelectedItem.Tag = MakeDisplayString(lvSellStation.SelectedItem.SubItems(3), lvSellStation.SelectedItem.Text)
        lvBus.SelectedItem.SubItems(ID_OffTime) = lvSellStation.SelectedItem.SubItems(1)
        lvBus.SelectedItem.SubItems(ID_FullPrice) = FormatMoney(lvSellStation.SelectedItem.SubItems(2))
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

'����ָ�������ڣ���ǰ���ڼ�Ԥ���������͵�վ����ˢ�³�����Ϣ
'pbForce��ʾ�Ƿ�ǿ��ˢ�£�����Ԥ�������͵�վ�Ƿ�ı䣩

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
 
    
    On Error GoTo here
    szStationID = RTrim(cboEndStation.BoundText)
    
    Set m_rsBusInfo = Nothing
    
    If cboEndStation.Changed Or pbForce Then
        
        If szStationID <> "" Then
            If m_szRegValue = "" Then
                Set rsCountTemp = m_oSell.GetBusRs(CDate(flblSellDate.Caption), szStationID)
            Else
                Set rsCountTemp = m_oSell.GetBusRsEx(CDate(flblSellDate.Caption), szStationID, m_szRegValue)
            End If
            
'            If pszSellStationID = "" Then
'              lvSellStation.ListItems.Clear
'            End If
            
            If rsCountTemp.RecordCount <> 0 Then
'                lblmileate = rsCountTemp!end_station_mileage & "����"
            End If
            
            Set m_rsBusInfo = rsCountTemp.Clone
            
            lvBus.Sorted = False
            lvBus.ListItems.Clear
            lvBus.Refresh
            For j = 1 To rsCountTemp.RecordCount
                If lvBus.ListItems.count > m_ISellScreenShow And m_ISellScreenShow <> 0 Then
                    With lvBus
                        If .ListItems.count > 0 Then
                            .SortKey = MDISellTicket.GetSortKey() - 1
                            .Sorted = True
                            For i = 1 To .ListItems.count
                            '������β���ͣ�����(��������λ��վƱ),���øó���ѡ��
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
                    rsCountTemp.MovePrevious
                    m_rsBusInfo.MovePrevious
'                    Set rsTemp = Nothing
                    Set liTemp = Nothing
                    Exit Sub
                Else
                'If DateAdd("n", 30, Now) < rsCountTemp!busstarttime And FormatDbValue(rsCountTemp!sell_station_id) <> "cm" Then
                '�ɸ��ݲ���ѡ���Ƿ���������공��
                If m_bListNoSeatBus Or rsCountTemp!sale_seat_quantity + rsCountTemp!sale_stand_seat_quantity > 0 Then
                
                    If Hour(rsCountTemp!BusStartTime) >= txtTime.Text Then
                        '�������ʱ����ڲ�ѯ��ʱ��
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
                            Set liTemp = lvBus.ListItems.Add(, "A" & RTrim(rsCountTemp!bus_id), RTrim(rsCountTemp!bus_id))
                            liTemp.SmallIcon = "StopBus"
                        Else
                            lForeColor = m_lNormalBusColor
                            Set liTemp = lvBus.ListItems.Add(, "A" & RTrim(rsCountTemp!bus_id), RTrim(rsCountTemp!bus_id))
                            '���δ���"A" & RTrim(rsCountTemp��bus_id)
                        End If
                        nBusType = rsCountTemp!bus_type
'                        If lForeColor <> m_lStopBusColor Then
                            liTemp.ForeColor = lForeColor
                            
'                         varBookmark = rsCountTemp.Bookmark
'                                If rsCountTemp.RecordCount > 0 Then
'                                   RefreshSellStation rsCountTemp
'                                End If
'                           rsCountTemp.Bookmark = varBookmark

                            If nBusType <> TP_ScrollBus Then
                                liTemp.SubItems(ID_BusType) = Trim(rsCountTemp!bus_type)
                                liTemp.SubItems(ID_OffTime) = Format(rsCountTemp!BusStartTime, "hh:mm")
                            Else
                                liTemp.SubItems(ID_VehicleModel) = cszScrollBus
                                liTemp.SubItems(ID_OffTime) = cszScrollBus
                            End If
                            liTemp.SubItems(ID_RouteName) = Trim(rsCountTemp!route_name)
                            liTemp.SubItems(ID_EndStation) = RTrim(rsCountTemp!end_station_name)
                            liTemp.SubItems(ID_TotalSeat) = rsCountTemp!total_seat
                            If IsDate(liTemp.SubItems(ID_OffTime)) Then
                                If g_bIsBookValid And DateAdd("n", -g_nBookTime, liTemp.SubItems(ID_OffTime)) < Time And ToDBDate(flblSellDate.Caption) = ToDBDate(Date) Then
                                    '�����������Ϊ����,���ѹ�Ԥ��ʱ��,��Ԥ�������ӵ�������������.
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
                                liTemp.SubItems(ID_FullPrice) = FormatMoney(rsCountTemp!full_price)
                                liTemp.SubItems(ID_HalfPrice) = FormatMoney(rsCountTemp!half_price)
                                liTemp.SubItems(ID_FreePrice) = 0
                                liTemp.SubItems(ID_PreferentialPrice1) = FormatMoney(rsCountTemp!preferential_ticket1)
                                liTemp.SubItems(ID_PreferentialPrice2) = FormatMoney(rsCountTemp!preferential_ticket2)
                                liTemp.SubItems(ID_PreferentialPrice3) = FormatMoney(rsCountTemp!preferential_ticket3)
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
                                liTemp.SubItems(ID_BedFullPrice) = FormatMoney(rsCountTemp!full_price)
                                liTemp.SubItems(ID_BedHalfPrice) = FormatMoney(rsCountTemp!half_price)
                                liTemp.SubItems(ID_BedFreePrice) = 0
                                liTemp.SubItems(ID_BedPreferentialPrice1) = FormatMoney(rsCountTemp!preferential_ticket1)
                                liTemp.SubItems(ID_BedPreferentialPrice2) = FormatMoney(rsCountTemp!preferential_ticket2)
                                liTemp.SubItems(ID_BedPreferentialPrice3) = FormatMoney(rsCountTemp!preferential_ticket3)
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
                                liTemp.SubItems(ID_AdditionalFullPrice) = FormatMoney(rsCountTemp!full_price)
                                liTemp.SubItems(ID_AdditionalHalfPrice) = FormatMoney(rsCountTemp!half_price)
                                liTemp.SubItems(ID_AdditionalFreePrice) = 0
                                liTemp.SubItems(ID_AdditionalPreferential1) = FormatMoney(rsCountTemp!preferential_ticket1)
                                liTemp.SubItems(ID_AdditionalPreferential2) = FormatMoney(rsCountTemp!preferential_ticket2)
                                liTemp.SubItems(ID_AdditionalPreferential3) = FormatMoney(rsCountTemp!preferential_ticket3)
                            End Select
                            '���¼��в���ʾ������ֻ�ǽ���洢���Ա�����
                            liTemp.SubItems(ID_LimitedCount) = rsCountTemp!sale_ticket_quantity
                            liTemp.SubItems(ID_LimitedTime) = rsCountTemp!stop_sale_time
                            liTemp.SubItems(ID_BusType1) = nBusType
                            liTemp.SubItems(ID_CheckGate) = rsCountTemp!check_gate_id
                            liTemp.SubItems(ID_StandCount) = rsCountTemp!sale_stand_seat_quantity
                            
                            liTemp.SubItems(ID_RealName) = Trim(rsCountTemp!id_card)
                            
                            liTemp.Tag = MakeDisplayString(Trim(rsCountTemp!sell_station_id), Trim(rsCountTemp!sell_station_name))

'����һ�䣺ͣ�೵���б�ɫ
                            If lForeColor = m_lStopBusColor Then
                                Dim oSubLtems As ListSubItem
                                For Each oSubLtems In liTemp.ListSubItems
                                    oSubLtems.ForeColor = lForeColor
                                Next
                            End If
                        End If
                    End If
                    
                    End If
                
                'End If
                
                
nextstep:
                rsCountTemp.MoveNext
                m_rsBusInfo.MoveNext
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


    
    '�趨ĳ��վ�㣿������ʾ�ڵ�һ��
    With lvBus
        If .ListItems.count > 0 Then
'             szScrollBus = .ListItems(i).SubItems(ID_OffTime)
            .SortKey = MDISellTicket.GetSortKey() - 1
            .Sorted = True
            For i = 1 To .ListItems.count

                '������β���ͣ�����(��������λ��վƱ),���øó���ѡ��   �Ƿ�Ϊ��������
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
        '�趨ĳ��վ�㣿������ʾ�ڵ�һ��
        '���ó��θı�Ҫ������Ӧ�����ķ���
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
here:
    ShowErrorMsg
    Set rsCountTemp = Nothing
    Set liTemp = Nothing
End Sub

'�õ��˴���Ʊ����Ӧ��ŵ�����
Private Function SelfGetSeatNo(pnIndex As Integer) As String
    If chkSetSeat.Enabled = False Then '���վƱѡ��,��ΪվƱ
        SelfGetSeatNo = "ST"
    ElseIf chkSetSeat.Enabled And txtSeat.Text <> "" Then  '�������ѡ��,��õ���Ӧ������
        SelfGetSeatNo = GetSeatNo(txtSeat.Text, pnIndex)
    Else '����Ϊ�Զ���λ��
        SelfGetSeatNo = ""
    End If
End Function

Private Function SelfGetSeatNo12(SetSeatEnable As Boolean, SetSeatValue As Integer, pnIndex As Integer, pszSeatNo As String) As String
    If SetSeatEnable = False Then '���վƱѡ��,��ΪվƱ
        SelfGetSeatNo12 = "ST"
    ElseIf SetSeatEnable And txtSeat.Text <> "" Then  '�������ѡ��,��õ���Ӧ������
        SelfGetSeatNo12 = GetSeatNo(pszSeatNo, pnIndex)
    Else '����Ϊ�Զ���λ��
        SelfGetSeatNo12 = ""
    End If
End Function

'�õ��˴���Ʊ����Ӧ��ŵ����ŵ�״̬
Private Function SelfGetSeatStatus(pnIndex As Integer) As String
    If chkSetSeat.Enabled = False Then '���վƱѡ��,��ΪվƱ
        SelfGetSeatStatus = "ST"
    ElseIf chkSetSeat.Enabled And txtSeat.Text <> "" Then  '�������ѡ��,��õ���Ӧ������
        SelfGetSeatStatus = GetSeatNo(szSeatStatus, pnIndex)
    Else '����Ϊ�Զ���λ��
        SelfGetSeatStatus = ""
    End If
End Function

Private Function SelfGetSeatStatus12(SetSeatEnable As Boolean, SetSeatValue As Integer, pnIndex As Integer, pszSeatNo As String) As String
    If SetSeatEnable = False Then '���վƱѡ��,��ΪվƱ
        SelfGetSeatStatus12 = "ST"
    ElseIf SetSeatEnable And txtSeat.Text <> "" Then  '�������ѡ��,��õ���Ӧ������
        SelfGetSeatStatus12 = GetSeatNo(szSeatStatus, pnIndex)
    Else '����Ϊ�Զ���λ��
        SelfGetSeatStatus12 = ""
    End If
End Function


Private Sub txtReceivedMoney_GotFocus()
    lblReceivedMoney.ForeColor = clActiveColor
    
    '********************
    '������ʾ��
    SetPay flblTotalPrice.Caption, GetStationNameInCbo(cboEndStation.Text), flblSellDate.Caption & " " & lvBus.SelectedItem.ListSubItems(ID_OffTime), nTicketCount, Val(cboInsurance.Text)
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
        cmdSell_Click
        cboEndStation.SetFocus
    End If
End Sub

Private Sub txtReceivedMoney_LostFocus()
'Dim nResult As VbMsgBoxResult
'If IsNumeric(txtReceivedMoney.Text) Then
'    If txtReceivedMoney.Text = 0 Then
'        nResult = MsgBox("�Ƿ�����һ��Ʊ����Ʊ���ۼ�", vbInformation + vbYesNo, Me.Caption)
'        If nResult = vbYes Then
'            cmdSell.Enabled = True
'            cmdSell_Click
'        End If
'    'if
'    End If
'End If
lblReceivedMoney.ForeColor = 0

    '********************
    '������ʾ��
        If Val(txtReceivedMoney.Text) <> 0 Then
            SetReceive txtReceivedMoney.Text
            If Val(txtReceivedMoney.Text) > Val(flblTotalPrice.Caption) Then
                SetReturn Val(flblRestMoney.Caption)
            Else
                SetThanks
            End If
        End If
    '********************
    'Timer1.Enabled = True
        '********************
    '������ʾ��
''        SetPay flblTotalPrice.Caption
''        If txtReceivedMoney.Text <> 0 Then
''            SetReceive txtReceivedMoney.Text
''            SetReturn flblRestMoney.Caption
''        End If
'        SetCheck
'        SetThanks
    '********************
  
End Sub



'��ʾHTMLHELP,ֱ�ӿ���
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
    On Error GoTo here
    szTemp = m_oSell.SellUnitCode
    m_oSell.SellUnitCode = m_szCurrentUnitID
    Set rsTemp = m_oSell.GetAllStationRs()
    m_oSell.SellUnitCode = szTemp

    With cboEndStation
        Set .RowSource = rsTemp
        'station_id:��վ����
        'station_input_code:��վ������
        'station_name:��������
        
        .BoundField = "station_id"
        .ListFields = "station_input_code:5,station_name:5"
        .AppendWithFields "station_id:9,station_name"
    End With
    
    '��Ϊվ���ѱ䣬���Ե�ǰ��ʾ�ĳ�����Ϣ��Ч���������
    lvBus.ListItems.Clear
    
    '���ó��θı�Ҫ������Ӧ�����ķ���
    DoThingWhenBusChange
    DealPrice
    
    Set rsTemp = Nothing
    On Error GoTo 0
    Exit Sub
here:
    Set rsTemp = Nothing
    ShowErrorMsg
End Sub


'��ȡ�Ż�Ʊ��Ϣ
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
    On Error GoTo here
    
    szTemp = m_oSell.SellUnitCode
    m_oSell.SellUnitCode = m_szCurrentUnitID

    '�õ����е�Ʊ��
    atTicketType = m_oParam.GetAllTicketType()
    aszSeatType = m_oSell.GetAllSeatType
    m_oSell.SellUnitCode = szTemp
    
    nCount = ArrayLength(atTicketType)
    nLen = ArrayLength(aszSeatType)
    
    
    sgWidth = 900
    lvBus.ColumnHeaders.Clear
    '���ListView��ͷ
    With lvBus.ColumnHeaders
        .Add , , "����", 950 '"BusID"
        .Add , , "��������", 0 '"BusType"
        .Add , , "ʱ��", 850 '"OffTime"
        .Add , , "��·����", 1650
        .Add , , "�յ�վ", 900 '"EndStation"
        .Add , , "��", 420
'        .Add , , "��", 440
        .Add , , "����", 670 '"SeatCount"
        .Add , , "��", 0
        .Add , , "��", 0 '440
        .Add , , "��", 0 '440
          .Add , , "����", 900 '"BusModel"
          .Add , , "��", 420
        '���Ʊ��,�����õ�������Ϊ0
        For i = 1 To nCount     '��λƱ��
            If atTicketType(i).nTicketTypeID = TP_FullPrice Then
                .Add , , "��ȫ", IIf(atTicketType(i).nTicketTypeValid = TP_TicketTypeValid, sgWidth, 0)
            Else
                .Add , , Trim(atTicketType(i).szTicketTypeName), IIf(atTicketType(i).nTicketTypeValid = TP_TicketTypeValid, sgWidth, 0)
                If atTicketType(i).nTicketTypeID = TP_HalfPrice Then lblHalfSell.Caption = Trim(atTicketType(i).szTicketTypeName) & "(&X)" & ":"
            End If
        Next i
        For i = 1 To nCount   '����Ʊ��
            If atTicketType(i).nTicketTypeID = TP_FullPrice Then
                .Add , , "��ȫ", IIf(atTicketType(i).nTicketTypeValid = TP_TicketTypeValid, 0, 0)
            Else
                .Add , , Trim(atTicketType(i).szTicketTypeName), IIf(atTicketType(i).nTicketTypeValid = TP_TicketTypeValid, 0, 0)
            End If
        Next i
        For i = 1 To nCount  '����Ʊ��
            If atTicketType(i).nTicketTypeID = TP_FullPrice Then
                .Add , , "��ȫ", IIf(atTicketType(i).nTicketTypeValid = TP_TicketTypeValid, 0, 0)
            Else
                .Add , , Trim(atTicketType(i).szTicketTypeName), IIf(atTicketType(i).nTicketTypeValid = TP_TicketTypeValid, 0, 0)
            End If
        Next i
        .Add , , "��������", 0 '"LimitedCount"
        .Add , , "����ʱ��", 0 '"LimitedTime"
        .Add , , "��������", 0 '"BusType"
        .Add , , "��Ʊ��", 0 '"CheckGate"
        .Add , , "վƱ", 0 '"StandCount"
        .Add , , "�Ƿ�ʵ����", 0
    End With
    
    '����ComboBox���Ż�Ʊ�Ƿ����
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
    
    '������λ����
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
    '����Ͽ��е�Ʊ�ִ�����Ʊ�����Ʒŵ�����m_atTicketType ��
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
here:
    ShowErrorMsg
End Sub

'�õ���Ӧ���Ż�Ʊ�ֵĶ�Ӧ��Ʊ��
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
Private Sub GetPreSellBus() '����Ԥ��Ʊ״̬��Ϣ
Dim i As Integer
Dim nCount As Integer
Dim atTicketType() As TTicketType

atTicketType = m_oSell.GetAllTicketType()
    nCount = ArrayLength(atTicketType)
    With lvPreSell.ColumnHeaders
        .Add , , "����", 950   '"BusID"
        .Add , , "��������", 0 '"BusType"
        .Add , , "ʱ��", 900 '"OffTime"
        .Add , , "�˳�����", 1450 '"BusDate"
        .Add , , "��ʼվ", 0  '"StartStation"
        .Add , , "�յ�վ", 899 '"EndStation"
        .Add , , "����", 899  '"VehicleModel"
        .Add , , "��Ʊ��", 899  '"SumTicketNum"
        .Add , , "��Ʊ��", 0  '"SumPrice"
        .Add , , "����", 0   '"OrderSeat"
        For i = 1 To nCount
            .Add , , Trim(atTicketType(i).szTicketTypeName) & "Ʊ��", 0  'Ʊ��
            .Add , , Trim(atTicketType(i).szTicketTypeName), 900
        Next i
        .Add , , "�ۿ�Ʊ��", 0  'DiscountPrice
        .Add , , "�ۿ���", 0   '"Discount"
        .Add , , "վƱ", 0     '"StandCount"
        .Add , , "��Ʊ��", 700   '"CheckGate"
        .Add , , "��������", 0 '"LimitedCount"
        .Add , , "�յ�վ����", 0 '"BoundText"
        .Add , , "��λ״̬1", 0 '"SetSeatEnable"
        .Add , , "��λ״̬2", 0 '"SetSeatValue"
        .Add , , "��λ��", 0  '"SeatNo"
        .Add , , "Ʊ����ϸ", 0 '"TicketPrice"
        .Add , , "Ʊ����ϸ", 0 ' "TicketType"
        .Add , , "��λ����", 0
        .Add , , "�յ�վ", 0
        .Add , , "�Ƿ�ʵ����", 0
    End With
End Sub
Private Sub GetPreSellTicketInfo()  '�õ�Ԥ���ݷ�Ʊ����Ϣ
    Dim liPreSell As ListItem
    Dim liSelected As ListItem
    Dim i As Integer
    Dim szPrice As String
    Dim szTicketType As String
    On Error GoTo here
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
                    .SubItems(IT_FullPrice) = FormatMoney(liSelected.SubItems(ID_FullPrice))
                    
                    .SubItems(IT_FullNum) = txtFullSell.Text & " ��"
                    
                    .SubItems(IT_HalfPrice) = FormatMoney(liSelected.SubItems(ID_HalfPrice))
                    .SubItems(IT_HalfNum) = txtHalfSell.Text & " ��"
                    

                    If cboPreferentialTicket.ListCount > 0 Then
                        Select Case m_atTicketType(cboPreferentialTicket.ListIndex + 1).nTicketTypeID
                        Case TP_FreeTicket
                            .SubItems(IT_FreeType) = FormatMoney(liSelected.SubItems(ID_FreePrice))
                            .SubItems(IT_FreeNum) = txtPreferentialSell.Text & " ��"
                        Case TP_PreferentialTicket1
                            .SubItems(IT_PreferentialType1) = FormatMoney(liSelected.SubItems(ID_PreferentialPrice1))
                            .SubItems(IT_PreferentialNum1) = txtPreferentialSell.Text
                        Case TP_PreferentialTicket2
                            .SubItems(IT_PreferentialType2) = FormatMoney(liSelected.SubItems(ID_PreferentialPrice2))
                            .SubItems(IT_PreferentialNum2) = txtPreferentialSell.Text & " ��"
                        Case TP_PreferentialTicket3
                            .SubItems(IT_PreferentialType3) = FormatMoney(liSelected.SubItems(ID_PreferentialPrice3))
                            .SubItems(IT_PreferentialNum3) = txtPreferentialSell.Text & " ��"
                        End Select
                    End If
                Case cszBedType
                    .SubItems(IT_FullPrice) = FormatMoney(liSelected.SubItems(ID_BedFullPrice))
                    .SubItems(IT_FullNum) = txtFullSell.Text & " ��"
                    .SubItems(IT_HalfPrice) = FormatMoney(liSelected.SubItems(ID_BedHalfPrice))
                    .SubItems(IT_HalfNum) = txtHalfSell.Text & " ��"
                    If cboPreferentialTicket.ListCount > 0 Then
                        Select Case m_atTicketType(cboPreferentialTicket.ListIndex + 1).nTicketTypeID
                        Case TP_FreeTicket
                            .SubItems(IT_FreeType) = FormatMoney(liSelected.SubItems(ID_BedFreePrice))
                            .SubItems(IT_FreeNum) = txtPreferentialSell.Text & " ��"
                        Case TP_PreferentialTicket1
                            .SubItems(IT_PreferentialType1) = FormatMoney(liSelected.SubItems(ID_BedPreferentialPrice1))
                            .SubItems(IT_PreferentialNum1) = txtPreferentialSell.Text & " ��"
                        Case TP_PreferentialTicket2
                            .SubItems(IT_PreferentialType2) = FormatMoney(liSelected.SubItems(ID_BedPreferentialPrice2))
                            .SubItems(IT_PreferentialNum2) = txtPreferentialSell.Text & " ��"
                        Case TP_PreferentialTicket3
                            .SubItems(IT_PreferentialType3) = FormatMoney(liSelected.SubItems(ID_BedPreferentialPrice3))
                            .SubItems(IT_PreferentialNum3) = txtPreferentialSell.Text & " ��"
                        End Select
                    End If
                Case cszAdditionalType
                    .SubItems(IT_FullPrice) = FormatMoney(liSelected.SubItems(ID_AdditionalFullPrice))
                    .SubItems(IT_FullNum) = txtFullSell.Text & " ��"
                    .SubItems(IT_HalfPrice) = FormatMoney(liSelected.SubItems(ID_AdditionalHalfPrice))
                    .SubItems(IT_HalfNum) = txtHalfSell.Text & " ��"
                    If cboPreferentialTicket.ListCount > 0 Then
                        Select Case m_atTicketType(cboPreferentialTicket.ListIndex + 1).nTicketTypeID
                        Case TP_FreeTicket
                            .SubItems(IT_FreeType) = liSelected.SubItems(ID_AdditionalFreePrice)
                            .SubItems(IT_FreeNum) = txtPreferentialSell.Text & " ��"
                        Case TP_PreferentialTicket1
                            .SubItems(IT_PreferentialType1) = liSelected.SubItems(ID_AdditionalPreferential1)
                            .SubItems(IT_PreferentialNum1) = txtPreferentialSell.Text & " ��"
                        Case TP_PreferentialTicket2
                            .SubItems(IT_PreferentialType2) = liSelected.SubItems(ID_AdditionalPreferential2)
                            .SubItems(IT_PreferentialNum2) = txtPreferentialSell.Text & " ��"
                        Case TP_PreferentialTicket3
                            .SubItems(IT_PreferentialType3) = liSelected.SubItems(ID_AdditionalPreferential3)
                            .SubItems(IT_PreferentialNum3) = txtPreferentialSell.Text & " ��"
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
            
            .SubItems(IT_RealName) = liSelected.SubItems(ID_RealName)
            
            '��ǩ����Ҫ��ʾ��Ч֤������ԭ��Ʊ��Ϣ��
            SetRealNameInfo .SubItems(IT_RealName), Val(.SubItems(IT_SumTicketNum))
        End With
    End If
    Set liPreSell = Nothing
    Set liSelected = Nothing
On Error GoTo 0
Exit Sub
here:
     Set liPreSell = Nothing
    Set liSelected = Nothing
    ShowErrorMsg
End Sub

Private Sub SetRealNameInfo(bIsRealName As Boolean, nSellNums As Integer)
On Error GoTo here
    Dim aszRealNameInfoTemp() As TCardInfo
    Dim i As Integer
    Dim nCountOld As Integer
    Dim nCountAdd As Integer
    Dim nCountNew As Integer
    
    If bIsRealName = True Then  '��Ҫʵ����
        frmRealNameReg.m_nSellCount = nSellNums
        frmRealNameReg.Show vbModal
        If frmRealNameReg.m_bOk Then
            aszRealNameInfoTemp = frmRealNameReg.m_vaCardInfo
            nCountOld = ArrayLength(m_aszRealNameInfo)
            nCountAdd = ArrayLength(aszRealNameInfoTemp)
            nCountNew = nCountOld + nCountAdd
            If nCountOld = 0 Then
                ReDim m_aszRealNameInfo(1 To nCountNew)
            Else
                ReDim Preserve m_aszRealNameInfo(1 To nCountNew)
            End If
            For i = 1 To nCountAdd
                m_aszRealNameInfo(nCountOld + i).szCardType = aszRealNameInfoTemp(i).szCardType
                m_aszRealNameInfo(nCountOld + i).szIDCardNo = aszRealNameInfoTemp(i).szIDCardNo
                m_aszRealNameInfo(nCountOld + i).szPersonName = aszRealNameInfoTemp(i).szPersonName
                m_aszRealNameInfo(nCountOld + i).szSex = aszRealNameInfoTemp(i).szSex
                m_aszRealNameInfo(nCountOld + i).szNation = aszRealNameInfoTemp(i).szNation
                m_aszRealNameInfo(nCountOld + i).szAddress = aszRealNameInfoTemp(i).szAddress
                m_aszRealNameInfo(nCountOld + i).szPersonPicture = aszRealNameInfoTemp(i).szPersonPicture
            Next i
        Else
            nCountOld = ArrayLength(m_aszRealNameInfo)
            nCountAdd = nSellNums
            nCountNew = nCountOld + nCountAdd
            If nCountOld = 0 Then
                ReDim m_aszRealNameInfo(1 To nCountNew)
            Else
                ReDim Preserve m_aszRealNameInfo(1 To nCountNew)
            End If
            For i = 1 To nCountAdd
                m_aszRealNameInfo(nCountOld + i).szCardType = ""
                m_aszRealNameInfo(nCountOld + i).szIDCardNo = ""
                m_aszRealNameInfo(nCountOld + i).szPersonName = ""
                m_aszRealNameInfo(nCountOld + i).szSex = ""
                m_aszRealNameInfo(nCountOld + i).szNation = ""
                m_aszRealNameInfo(nCountOld + i).szAddress = ""
                m_aszRealNameInfo(nCountOld + i).szPersonPicture = ""
            Next i
        End If
    Else    '����Ҫʵ����
        nCountOld = ArrayLength(m_aszRealNameInfo)
        nCountAdd = nSellNums
        nCountNew = nCountOld + nCountAdd
        If nCountOld = 0 Then
            ReDim m_aszRealNameInfo(1 To nCountNew)
        Else
            ReDim Preserve m_aszRealNameInfo(1 To nCountNew)
        End If
        For i = 1 To nCountAdd
            m_aszRealNameInfo(nCountOld + i).szCardType = ""
            m_aszRealNameInfo(nCountOld + i).szIDCardNo = ""
            m_aszRealNameInfo(nCountOld + i).szPersonName = ""
            m_aszRealNameInfo(nCountOld + i).szSex = ""
            m_aszRealNameInfo(nCountOld + i).szNation = ""
            m_aszRealNameInfo(nCountOld + i).szAddress = ""
            m_aszRealNameInfo(nCountOld + i).szPersonPicture = ""
        Next i
    End If

    Exit Sub
here:
    ShowErrorMsg
End Sub

Private Function GetDealTotalPrice() As Double  '�õ���Ʊ��
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
'����Ԥ�۰�ť״̬
Private Sub SetPreSellButton()
    If txtFullSell.Text = 0 And txtHalfSell.Text = 0 And txtPreferentialSell.Text = 0 Then
        cmdPreSell.Enabled = False
    Else
        cmdPreSell.Enabled = True
    End If
End Sub
'/////////////////////////////////
'�����ۿ�Ʊ�붨��
Private Sub DealDiscountAndSeat()
   '�ж��Ƿ������ۿ�ƱȨ��
   On Error GoTo here
   If m_oSell.DiscountIsValid Then
        txtDiscount.Enabled = False
        fraDiscountTicket.Enabled = False
   End If
   If m_oSell.OrderSeatIsValid Then
        chkSetSeat.Value = 0
        chkSetSeat.Visible = False
        lblSetSeat.Enabled = False
   End If
   On Error GoTo 0
   Exit Sub
here:
   ShowErrorMsg
End Sub
'//////////////////////////
'Ԥ���ö���
Private Function PreOrderSeat() As String
Dim i As Integer
Dim szTemp As String
Dim liTemp As ListItem
On Error GoTo here
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
here:
    PreOrderSeat = ""
    ShowErrorMsg
End Function
'//////////////////////////////////////
'������ͬ��������
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
'�ϲ���ͬ������Ϣ
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
 
    .SubItems(IT_SumPrice) = FormatMoney(Val(.SubItems(IT_SumPrice)) + _
                                txtFullSell.Text * lvBus.SelectedItem.SubItems(ID_FullPrice) + _
                                txtHalfSell.Text * lvBus.SelectedItem.SubItems(ID_HalfPrice) + _
                                txtPreferentialSell.Text * GetPreferentialPrice)
                                
 
End With
End Sub


''�ж��Ƿ��й�������
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
''��ʼ��վ�㳵��˳������
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
'�жϹ���վ���Ƿ���������鵱��
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
'������˳����С������ֵ
Private Sub AddValueToIndex(pnIndex As Integer)
    If m_atbBusOrder(pnIndex).dbCount > 1000 Then
        m_atbBusOrder(pnIndex).dbCount = 1
    Else
        m_atbBusOrder(pnIndex).dbCount = m_atbBusOrder(pnIndex).dbCount + 1
    End If
End Sub
'lvBus����ʾ��ȷ�ĳ���˳��
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
'��ʾ��һ��
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
                        liTemp.SubItems(ID_FullPrice) = FormatMoney(m_rsBusInfo!full_price)
                        liTemp.SubItems(ID_HalfPrice) = FormatMoney(m_rsBusInfo!half_price)
                        liTemp.SubItems(ID_PreferentialPrice1) = FormatMoney(m_rsBusInfo!preferential_ticket1)
                        liTemp.SubItems(ID_PreferentialPrice2) = FormatMoney(m_rsBusInfo!preferential_ticket2)
                        liTemp.SubItems(ID_PreferentialPrice3) = FormatMoney(m_rsBusInfo!preferential_ticket3)
                    Case cszBedType
                        liTemp.SubItems(ID_BedFullPrice) = FormatMoney(m_rsBusInfo!full_price)
                        liTemp.SubItems(ID_BedHalfPrice) = FormatMoney(m_rsBusInfo!half_price)
                        liTemp.SubItems(ID_BedPreferentialPrice1) = FormatMoney(m_rsBusInfo!preferential_ticket1)
                        liTemp.SubItems(ID_BedPreferentialPrice2) = FormatMoney(m_rsBusInfo!preferential_ticket2)
                        liTemp.SubItems(ID_BedPreferentialPrice3) = FormatMoney(m_rsBusInfo!preferential_ticket3)
                    Case cszAdditionalType
                        liTemp.SubItems(ID_AdditionalFullPrice) = FormatMoney(m_rsBusInfo!full_price)
                        liTemp.SubItems(ID_AdditionalHalfPrice) = FormatMoney(m_rsBusInfo!half_price)
                        liTemp.SubItems(ID_AdditionalPreferential1) = FormatMoney(m_rsBusInfo!preferential_ticket1)
                        liTemp.SubItems(ID_AdditionalPreferential2) = FormatMoney(m_rsBusInfo!preferential_ticket2)
                        liTemp.SubItems(ID_AdditionalPreferential3) = FormatMoney(m_rsBusInfo!preferential_ticket3)

                End Select
                GoTo nextstep
            End If
            Exit For
        Next i
        If m_rsBusInfo!status = ST_BusStopped Or m_rsBusInfo!status = ST_BusMergeStopped Or m_rsBusInfo!status = ST_BusSlitpStop Then
            lForeColor = m_lStopBusColor

        Else
            lForeColor = m_lNormalBusColor
            Set liTemp = lvBus.ListItems.Add(, "A" & RTrim(m_rsBusInfo!bus_id), RTrim(m_rsBusInfo!bus_id))   '���δ���"A" & RTrim(m_rsBusInfo��bus_id)
        End If

        nBusType = m_rsBusInfo!bus_type


        If lForeColor <> m_lStopBusColor Then
            liTemp.ForeColor = lForeColor
            If nBusType <> TP_ScrollBus Then
                liTemp.SubItems(ID_BusType) = Trim(m_rsBusInfo!bus_type)
                liTemp.SubItems(ID_OffTime) = Format(m_rsBusInfo!BusStartTime, "hh:mm")

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
                    liTemp.SubItems(ID_FullPrice) = FormatMoney(m_rsBusInfo!full_price)
                    liTemp.SubItems(ID_HalfPrice) = FormatMoney(m_rsBusInfo!half_price)
                    liTemp.SubItems(ID_FreePrice) = 0
                    liTemp.SubItems(ID_PreferentialPrice1) = FormatMoney(m_rsBusInfo!preferential_ticket1)
                    liTemp.SubItems(ID_PreferentialPrice2) = FormatMoney(m_rsBusInfo!preferential_ticket2)
                    liTemp.SubItems(ID_PreferentialPrice3) = FormatMoney(m_rsBusInfo!preferential_ticket3)
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
                    liTemp.SubItems(ID_BedFullPrice) = FormatMoney(m_rsBusInfo!full_price)
                    liTemp.SubItems(ID_BedHalfPrice) = FormatMoney(m_rsBusInfo!half_price)
                    liTemp.SubItems(ID_BedFreePrice) = 0
                    liTemp.SubItems(ID_BedPreferentialPrice1) = FormatMoney(m_rsBusInfo!preferential_ticket1)
                    liTemp.SubItems(ID_BedPreferentialPrice2) = FormatMoney(m_rsBusInfo!preferential_ticket2)
                    liTemp.SubItems(ID_BedPreferentialPrice3) = FormatMoney(m_rsBusInfo!preferential_ticket3)
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
                    liTemp.SubItems(ID_AdditionalFullPrice) = FormatMoney(m_rsBusInfo!full_price)
                    liTemp.SubItems(ID_AdditionalHalfPrice) = FormatMoney(m_rsBusInfo!half_price)
                    liTemp.SubItems(ID_AdditionalFreePrice) = 0
                    liTemp.SubItems(ID_AdditionalPreferential1) = FormatMoney(m_rsBusInfo!preferential_ticket1)
                    liTemp.SubItems(ID_AdditionalPreferential2) = FormatMoney(m_rsBusInfo!preferential_ticket2)
                    liTemp.SubItems(ID_AdditionalPreferential3) = FormatMoney(m_rsBusInfo!preferential_ticket3)

            End Select

            '���¼��в���ʾ������ֻ�ǽ���洢���Ա�����
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
'��ʾ��һ��
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
                        liTemp.SubItems(ID_FullPrice) = FormatMoney(m_rsBusInfo!full_price)
                        liTemp.SubItems(ID_HalfPrice) = FormatMoney(m_rsBusInfo!half_price)
                        liTemp.SubItems(ID_PreferentialPrice1) = FormatMoney(m_rsBusInfo!preferential_ticket1)
                        liTemp.SubItems(ID_PreferentialPrice2) = FormatMoney(m_rsBusInfo!preferential_ticket2)
                        liTemp.SubItems(ID_PreferentialPrice3) = FormatMoney(m_rsBusInfo!preferential_ticket3)
                    Case cszBedType
                        liTemp.SubItems(ID_BedFullPrice) = FormatMoney(m_rsBusInfo!full_price)
                        liTemp.SubItems(ID_BedHalfPrice) = FormatMoney(m_rsBusInfo!half_price)
                        liTemp.SubItems(ID_BedPreferentialPrice1) = FormatMoney(m_rsBusInfo!preferential_ticket1FormatMoney)
                        liTemp.SubItems(ID_BedPreferentialPrice2) = FormatMoney(m_rsBusInfo!preferential_ticket2)
                        liTemp.SubItems(ID_BedPreferentialPrice3) = FormatMoney(m_rsBusInfo!preferential_ticket3)
                    Case cszAdditionalType
                        liTemp.SubItems(ID_AdditionalFullPrice) = FormatMoney(m_rsBusInfo!full_price)
                        liTemp.SubItems(ID_AdditionalHalfPrice) = FormatMoney(m_rsBusInfo!half_price)
                        liTemp.SubItems(ID_AdditionalPreferential1) = FormatMoney(m_rsBusInfo!preferential_ticket1)
                        liTemp.SubItems(ID_AdditionalPreferential2) = FormatMoney(m_rsBusInfo!preferential_ticket2)
                        liTemp.SubItems(ID_AdditionalPreferential3) = FormatMoney(m_rsBusInfo!preferential_ticket3)

                End Select
                GoTo nextstep
            End If
            Exit For
        Next i
        If m_rsBusInfo!status = ST_BusStopped Or m_rsBusInfo!status = ST_BusMergeStopped Or m_rsBusInfo!status = ST_BusSlitpStop Then
            lForeColor = m_lStopBusColor

        Else
            lForeColor = m_lNormalBusColor
            Set liTemp = lvBus.ListItems.Add(, "A" & RTrim(m_rsBusInfo!bus_id), RTrim(m_rsBusInfo!bus_id))   '���δ���"A" & RTrim(m_rsBusInfo��bus_id)
        End If

        nBusType = m_rsBusInfo!bus_type


        If lForeColor <> m_lStopBusColor Then
            liTemp.ForeColor = lForeColor
            If nBusType <> TP_ScrollBus Then
                liTemp.SubItems(ID_BusType) = Trim(m_rsBusInfo!bus_type)
                liTemp.SubItems(ID_OffTime) = Format(m_rsBusInfo!BusStartTime, "hh:mm")

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
                    liTemp.SubItems(ID_FullPrice) = FormatMoney(m_rsBusInfo!full_price)
                    liTemp.SubItems(ID_HalfPrice) = FormatMoney(m_rsBusInfo!half_price)
                    liTemp.SubItems(ID_FreePrice) = 0
                    liTemp.SubItems(ID_PreferentialPrice1) = FormatMoney(m_rsBusInfo!preferential_ticket1)
                    liTemp.SubItems(ID_PreferentialPrice2) = FormatMoney(m_rsBusInfo!preferential_ticket2)
                    liTemp.SubItems(ID_PreferentialPrice3) = FormatMoney(m_rsBusInfo!preferential_ticket3)
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
                    liTemp.SubItems(ID_BedFullPrice) = FormatMoney(m_rsBusInfo!full_price)
                    liTemp.SubItems(ID_BedHalfPrice) = FormatMoney(m_rsBusInfo!half_price)
                    liTemp.SubItems(ID_BedFreePrice) = 0
                    liTemp.SubItems(ID_BedPreferentialPrice1) = FormatMoney(m_rsBusInfo!preferential_ticket1)
                    liTemp.SubItems(ID_BedPreferentialPrice2) = FormatMoney(m_rsBusInfo!preferential_ticket2)
                    liTemp.SubItems(ID_BedPreferentialPrice3) = FormatMoney(m_rsBusInfo!preferential_ticket3)
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
                    liTemp.SubItems(ID_AdditionalFullPrice) = FormatMoney(m_rsBusInfo!full_price)
                    liTemp.SubItems(ID_AdditionalHalfPrice) = FormatMoney(m_rsBusInfo!half_price)
                    liTemp.SubItems(ID_AdditionalFreePrice) = 0
                    liTemp.SubItems(ID_AdditionalPreferential1) = FormatMoney(m_rsBusInfo!preferential_ticket1)
                    liTemp.SubItems(ID_AdditionalPreferential2) = FormatMoney(m_rsBusInfo!preferential_ticket2)
                    liTemp.SubItems(ID_AdditionalPreferential3) = FormatMoney(m_rsBusInfo!preferential_ticket3)

            End Select

            '���¼��в���ʾ������ֻ�ǽ���洢���Ա�����
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

Private Sub SellTicket()
    Dim i As Integer
    Dim szBookNumber() As String
    Dim dbTotalMoney As Double  '��Ʊ��
    Dim dbRealTotalMoney As Double 'ʵ��Ʊ��
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
    Dim anInsurance() As Integer '��Ʊ��
    Dim panInsurance() As Integer '��ӡ��
    
    Dim nTotal As Integer
    Dim n As Integer
    
    Dim liTemp As ListItem
    
    Dim nCount As Integer
    Dim nLen As Integer
    Dim nTicketOffset As Integer
    Dim nLength As Integer
    Dim nTemp As Integer
    Dim szTemp As String
    
    
    If m_bPrint Then Exit Sub
    
    
    dbTotalMoney = 0
    dbRealTotalMoney = 0
    
    If txtDiscount.Text > 1 Then
       MsgBox "�ۿ��ʲ��ܴ���1", vbInformation, "��ʾ"
       txtDiscount.SetFocus
       Exit Sub
    End If
    szTemp = flblRestMoney.Caption
    cmdSell.Enabled = False
    m_bSumPriceIsEmpty = True
    m_bPrint = True
    On Error GoTo here
    
    '��������������Ʊ����
    '////////////////////
    
    '-------------------------------------------------------------------------------------
    If lvPreSell.ListItems.count = 0 Then
          
        If txtPreferentialSell.Text + txtHalfSell.Text + txtFullSell.Text <> 0 Then
            lblSellMsg.Caption = "���ڴ�����Ʊ"
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
'                aSellTicket(1).BuyTicketInfo(i).szReserved = szBookNumber(1)
                aSellTicket(1).BuyTicketInfo(i).szReserved = SelfGetSeatStatus(i)
                aSellTicket(1).BuyTicketInfo(i).szSeatTypeID = m_aszSeatType(cboSeatType.ListIndex + 1, 1)
                aSellTicket(1).BuyTicketInfo(i).szSeatTypeName = m_aszSeatType(cboSeatType.ListIndex + 1, 2)
            Next
            
            For i = 1 To txtHalfSell.Text
                aSellTicket(1).BuyTicketInfo(i + txtFullSell.Text).nTicketType = TP_HalfPrice
                aSellTicket(1).BuyTicketInfo(i + txtFullSell.Text).szTicketNo = GetTicketNo(i + txtFullSell.Text - 1)
                
                aSellTicket(1).BuyTicketInfo(i + txtFullSell.Text).szSeatNo = SelfGetSeatNo(i + txtFullSell.Text)
                '��۵�Ʊ��
                aSellTicket(1).pasgSellTicketPrice(i + txtFullSell.Text) = CDbl(liTemp.SubItems(ID_HalfPrice)) * CSng(txtDiscount.Text)
                
'                aSellTicket(1).BuyTicketInfo(i + txtFullSell.Text).szReserved = szBookNumber(1)
                aSellTicket(1).BuyTicketInfo(i + txtFullSell.Text).szReserved = SelfGetSeatStatus(i + txtFullSell.Text)
                
                aSellTicket(1).BuyTicketInfo(i + txtFullSell.Text).szSeatTypeID = m_aszSeatType(cboSeatType.ListIndex + 1, 1)
                aSellTicket(1).BuyTicketInfo(i + txtFullSell.Text).szSeatTypeName = m_aszSeatType(cboSeatType.ListIndex + 1, 2)
            Next
            
            For i = 1 To txtPreferentialSell.Text
                aSellTicket(1).BuyTicketInfo(i + txtFullSell.Text + txtHalfSell.Text).nTicketType = m_atTicketType(cboPreferentialTicket.ListIndex + 1).nTicketTypeID
                aSellTicket(1).BuyTicketInfo(i + txtFullSell.Text + txtHalfSell.Text).szTicketNo = GetTicketNo(i + txtFullSell.Text + txtHalfSell.Text - 1)
                
                aSellTicket(1).BuyTicketInfo(i + txtFullSell.Text + txtHalfSell.Text).szSeatNo = SelfGetSeatNo(i + txtFullSell.Text + txtHalfSell.Text)
                aSellTicket(1).pasgSellTicketPrice(i + txtFullSell.Text + txtHalfSell.Text) = CDbl(GetPreferentialPrice(True)) * CSng(txtDiscount.Text)
                
'                aSellTicket(1).BuyTicketInfo(i + txtFullSell.Text + txtHalfSell.Text).szReserved = szBookNumber(1)
                aSellTicket(1).BuyTicketInfo(i + txtFullSell.Text + txtHalfSell.Text).szReserved = SelfGetSeatStatus(i + txtFullSell.Text + txtHalfSell.Text)
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
            
            anInsurance(1) = Val(cboInsurance.Text)
            
            nTotal = Val(txtFullSell.Text) + Val(txtHalfSell.Text) + Val(txtPreferentialSell.Text)
   
            '�ж���Ʊ���Ƿ���ʵ��������һ��
            If nTotal <> ArrayLength(m_aszRealNameInfo) Then
                MsgBox "֤����[" & ArrayLength(m_aszRealNameInfo) & "]������Ʊ��[" & nTotal & "]�Ų�����", vbExclamation, App.Title
                GoTo out
            End If
            
            srSellResult = m_oSell.SellTicket(dyBusDate, szBusID, szSellStationID, szDesStationID, szDesStationName, aSellTicket, anInsurance)
            
            IncTicketNo txtFullSell.Text + txtHalfSell.Text + txtPreferentialSell.Text
            
            'ˢ����λ��Ϣ
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
    
            
    
            '����Ӧ���Ǵ�ӡƱ�Ĵ���
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
            
            panInsurance(1) = Val(cboInsurance.Text)
            
            For i = 1 To pnTicketCount(1)
                apiTicketInfo(1).aptPrintTicketInfo(i).nTicketType = aSellTicket(1).BuyTicketInfo(i).nTicketType
                apiTicketInfo(1).aptPrintTicketInfo(i).sgTicketPrice = srSellResult(1).asgTicketPrice(i)
                apiTicketInfo(1).aptPrintTicketInfo(i).szSeatNo = srSellResult(1).aszSeatNo(i)
                apiTicketInfo(1).aptPrintTicketInfo(i).szTicketNo = aSellTicket(1).BuyTicketInfo(i).szTicketNo
                
                'ȡ��ʵ����Ʊ��
                If srSellResult(1).aszTicketType(i) <> TP_FreeTicket Then
                    dbRealTotalMoney = srSellResult(1).asgTicketPrice(i) + dbRealTotalMoney
                End If
            Next
            
            
            '����Ʊ���ۼ�
            If IsNumeric(txtReceivedMoney.Text) Then
    
                If txtReceivedMoney.Text = 0 Then
                    m_sgTotalMoney = lblTotalMoney.Caption
                Else
                    m_sgTotalMoney = 0#
                End If
            End If
            lblSellMsg.Caption = "���ڴ�ӡ��Ʊ"
            lblSellMsg.Refresh
            
            PrintTicket apiTicketInfo, pszBusDate, pnTicketCount, pszEndStation, pszOffTime, pszBusID, pszVehicleType, pszCheckGate, pbSaleChange, aszTerminateName, szSellStationName, panInsurance, m_aszRealNameInfo
            
            m_aszInsurce = CombInsurance(apiTicketInfo, szBusID, pnTicketCount, pszBusDate, szDesStationID, pszEndStation, pszOffTime, szSellStationID, szSellStationName, m_aszRealNameInfo)
            SaveInsurance m_aszInsurce
            
            If ArrayLength(srSellResult) = 0 Then
                frmNotify.Show vbModal
            End If
            
            dbTotalMoney = CDbl(flblTotalPrice.Caption)
            
        End If
    Else

        nTicketOffset = 0
        nLength = 0
        lblSellMsg.Caption = "���ڴ�����Ʊ"
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
'                    aSellTicket(nCount).BuyTicketInfo(i).szReserved = szBookNumber(nCount)
                    aSellTicket(nCount).BuyTicketInfo(i).szReserved = SelfGetSeatStatus12(.SubItems(IT_SetSeatEnable), .SubItems(IT_SetSeatValue), i, .SubItems(IT_SeatNo))
                    aSellTicket(nCount).BuyTicketInfo(i).szSeatTypeID = .SubItems(IT_SeatType)
                    aSellTicket(nCount).BuyTicketInfo(i).szSeatTypeName = GetSeatTypeName(.SubItems(IT_SeatType))
                Next i
                
                nTemp = Val(.SubItems(IT_FullNum))
                For i = 1 To Val(.SubItems(IT_HalfNum))
                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).nTicketType = TP_HalfPrice
                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).szTicketNo = GetTicketNo(i - 1 + nTemp + nTicketOffset)
                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).szSeatNo = SelfGetSeatNo12(.SubItems(IT_SetSeatEnable), .SubItems(IT_SetSeatValue), i + nTemp, .SubItems(IT_SeatNo))
                    aSellTicket(nCount).pasgSellTicketPrice(i + nTemp) = CDbl(.SubItems(IT_HalfPrice)) * CSng(.SubItems(IT_Discount))
'                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).szReserved = szBookNumber(nCount)
                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).szReserved = SelfGetSeatStatus12(.SubItems(IT_SetSeatEnable), .SubItems(IT_SetSeatValue), i + nTemp, .SubItems(IT_SeatNo))
                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).szSeatTypeID = .SubItems(IT_SeatType)
                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).szSeatTypeName = GetSeatTypeName(.SubItems(IT_SeatType))
                Next i
                
                nTemp = nTemp + Val(.SubItems(IT_HalfNum))
                For i = 1 To Val(.SubItems(IT_FreeNum))
                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).nTicketType = TP_FreeTicket
                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).szTicketNo = GetTicketNo(i - 1 + nTemp + nTicketOffset)
                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).szSeatNo = SelfGetSeatNo12(.SubItems(IT_SetSeatEnable), .SubItems(IT_SetSeatValue), i + nTemp, .SubItems(IT_SeatNo))
                    aSellTicket(nCount).pasgSellTicketPrice(i + nTemp) = CDbl(.SubItems(IT_FullPrice)) * CSng(.SubItems(IT_Discount))
'                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).szReserved = szBookNumber(nCount)
                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).szReserved = SelfGetSeatStatus12(.SubItems(IT_SetSeatEnable), .SubItems(IT_SetSeatValue), i + nTemp, .SubItems(IT_SeatNo))
                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).szSeatTypeID = .SubItems(IT_SeatType)
                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).szSeatTypeName = GetSeatTypeName(.SubItems(IT_SeatType))
                Next i
                
                nTemp = nTemp + Val(.SubItems(IT_FreeNum))
                For i = 1 To Val(.SubItems(IT_PreferentialNum1))
                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).nTicketType = TP_PreferentialTicket1
                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).szTicketNo = GetTicketNo(i - 1 + nTemp + nTicketOffset)
                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).szSeatNo = SelfGetSeatNo12(.SubItems(IT_SetSeatEnable), .SubItems(IT_SetSeatValue), i + nTemp, .SubItems(IT_SeatNo))
                    aSellTicket(nCount).pasgSellTicketPrice(i + nTemp) = CDbl(.SubItems(IT_PreferentialType1)) * CSng(.SubItems(IT_Discount))
'                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).szReserved = szBookNumber(nCount)
                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).szReserved = SelfGetSeatStatus12(.SubItems(IT_SetSeatEnable), .SubItems(IT_SetSeatValue), i + nTemp, .SubItems(IT_SeatNo))
                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).szSeatTypeID = .SubItems(IT_SeatType)
                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).szSeatTypeName = GetSeatTypeName(.SubItems(IT_SeatType))
                Next i
                nTemp = nTemp + Val(.SubItems(IT_PreferentialNum1))
                For i = 1 To Val(.SubItems(IT_PreferentialNum2))
                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).nTicketType = TP_PreferentialTicket2
                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).szTicketNo = GetTicketNo(i - 1 + nTemp + nTicketOffset)
                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).szSeatNo = SelfGetSeatNo12(.SubItems(IT_SetSeatEnable), .SubItems(IT_SetSeatValue), i + nTemp, .SubItems(IT_SeatNo))
                    aSellTicket(nCount).pasgSellTicketPrice(i + nTemp) = CDbl(.SubItems(IT_PreferentialType2)) * CSng(.SubItems(IT_Discount))
'                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).szReserved = szBookNumber(nCount)
                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).szReserved = SelfGetSeatStatus12(.SubItems(IT_SetSeatEnable), .SubItems(IT_SetSeatValue), i + nTemp, .SubItems(IT_SeatNo))
                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).szSeatTypeID = .SubItems(IT_SeatType)
                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).szSeatTypeName = GetSeatTypeName(.SubItems(IT_SeatType))
                Next i
                nTemp = nTemp + Val(.SubItems(IT_PreferentialNum2))
                For i = 1 To Val(.SubItems(IT_PreferentialNum3))
                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).nTicketType = TP_PreferentialTicket3
                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).szTicketNo = GetTicketNo(i - 1 + nTemp + nTicketOffset)
                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).szSeatNo = SelfGetSeatNo12(.SubItems(IT_SetSeatEnable), .SubItems(IT_SetSeatValue), i + nTemp, .SubItems(IT_SeatNo))
                    aSellTicket(nCount).pasgSellTicketPrice(i + nTemp) = CDbl(.SubItems(IT_PreferentialType3)) * CSng(.SubItems(IT_Discount))
'                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).szReserved = szBookNumber(nCount)
                    aSellTicket(nCount).BuyTicketInfo(i + nTemp).szReserved = SelfGetSeatStatus12(.SubItems(IT_SetSeatEnable), .SubItems(IT_SetSeatValue), i + nTemp, .SubItems(IT_SeatNo))
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
            
            anInsurance(nCount) = Val(cboInsurance.Text)
        Next nCount
        
            nTotal = nTicketOffset
            '�ж���Ʊ���Ƿ���ʵ��������һ��
            If nTotal <> ArrayLength(m_aszRealNameInfo) Then
                MsgBox "֤����[" & ArrayLength(m_aszRealNameInfo) & "]������Ʊ��[" & nTotal & "]�Ų�����", vbExclamation, App.Title
                GoTo out
            End If
        
            srSellResult = m_oSell.SellTicket(dyBusDate, szBusID, szSellStationID, szDesStationID, szDesStationName, aSellTicket, anInsurance, m_aszRealNameInfo)
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
    
            '����Ӧ���Ǵ�ӡƱ�Ĵ���
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
                    
                    panInsurance(1) = Val(cboInsurance.Text)
            
                    For i = 1 To Val(.SubItems(IT_SumTicketNum))
                        apiTicketInfo(nCount).aptPrintTicketInfo(i).nTicketType = aSellTicket(nCount).BuyTicketInfo(i).nTicketType
                        apiTicketInfo(nCount).aptPrintTicketInfo(i).sgTicketPrice = srSellResult(nCount).asgTicketPrice(i)
                        apiTicketInfo(nCount).aptPrintTicketInfo(i).szSeatNo = srSellResult(nCount).aszSeatNo(i)
                        apiTicketInfo(nCount).aptPrintTicketInfo(i).szTicketNo = aSellTicket(nCount).BuyTicketInfo(i).szTicketNo
                        
                    Next
                End With
            Next nCount
            
             'ȡ��ʵ����Ʊ��
            For nCount = 1 To ArrayLength(srSellResult)
                For i = 1 To ArrayLength(srSellResult(nCount).asgTicketPrice)
                   If srSellResult(nCount).aszTicketType(i) <> TP_FreeTicket Then
                            dbRealTotalMoney = srSellResult(nCount).asgTicketPrice(i) + dbRealTotalMoney
                   End If
                Next i
            Next nCount
            lblSellMsg.Caption = "���ڴ�ӡ��Ʊ"
            lblSellMsg.Refresh
            
           PrintTicket apiTicketInfo, pszBusDate, pnTicketCount, pszEndStation, pszOffTime, pszBusID, pszVehicleType, pszCheckGate, pbSaleChange, aszTerminateName, szSellStationName, anInsurance, m_aszRealNameInfo
           
           m_aszInsurce = CombInsurance(apiTicketInfo, szBusID, pnTicketCount, pszBusDate, szDesStationID, pszEndStation, pszOffTime, szSellStationID, szSellStationName, m_aszRealNameInfo)
           SaveInsurance m_aszInsurce
           
           dbTotalMoney = CDbl(flblTotalPrice.Caption)
'            '���Ʊ�۱���
'
'
'            If Abs(dbRealTotalMoney - dbTotalMoney) > 0.01 Then
'                frmPriceInfo.m_sngTotalPrice = dbRealTotalMoney
'                frmPriceInfo.Show vbModal
'            End If
           
           If IsNumeric(txtReceivedMoney.Text) Then
            
                If txtReceivedMoney.Text = 0 Then
                    m_sgTotalMoney = lblTotalMoney.Caption
                Else
                    m_sgTotalMoney = 0#
                End If
            End If
        
        '����Ʊ���ۼ�
           
          If ArrayLength(srSellResult) = 0 Then
                frmNotify.m_szErrorDescription = "[һ�����������]  ��ע���һ��Ʊ��"
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
        
        'chkInsurance.Value = vbUnchecked fpd
        
        flblRestMoney.Caption = szTemp
        frmOrderSeats.m_szBookNumber = ""
        txtSeat.Text = ""
        SetPreSellButton
        
        Erase m_aszRealNameInfo
        
'        lvPreSell.ListItems.Clear
'        cboEndStation.SetFocus
        txtReceivedMoney.SetFocus
        m_bSetFocus = False
    Exit Sub
out:
        m_bPreClear = True
        lblSellMsg.Caption = ""
        cmdSell.Enabled = True
        m_bPrint = False
        txtPrevDate.Value = 0
        txtTime.Text = 0
        
        flblRestMoney.Caption = szTemp
        frmOrderSeats.m_szBookNumber = ""
        txtSeat.Text = ""

        SetPreSellButton

        'fpd
        Dim szRestMoney As String
        Dim szTotaMoney As String
        szRestMoney = flblRestMoney.Caption
        szTotaMoney = flblTotalPrice.Caption
        
        ClearInfo

        flblRestMoney.Caption = Format(szRestMoney, "0.00")
        flblTotalPrice.Caption = Format(szTotaMoney, "0.00")

    Exit Sub
    
here:
    frmOrderSeats.m_szBookNumber = ""
    lblSellMsg.Caption = ""
    m_bPrint = False
    If err.Number = 91 Then
       MsgBox "�����վ�����޳��Σ�", vbInformation + vbOKOnly, "��ʾ"
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

Private Sub ClearInfo()
    Erase m_aszRealNameInfo
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
    '���ܱ��շ�
    Dim i As Integer
    Dim nCount As Integer
    If cboInsurance.ListIndex > 0 Then
        nCount = 0
        For i = 1 To lvPreSell.ListItems.count
            nCount = nCount + lvPreSell.ListItems(i).SubItems(IT_SumTicketNum)
        Next i
        nCount = nCount + Val(txtFullSell.Text) + Val(txtHalfSell.Text) + Val(txtPreferentialSell.Text)
        
        TotalInsurace = nCount * Val(cboInsurance.Text)
    Else
        TotalInsurace = 0
    End If
End Function





Private Sub DispStationAndNum()
    '�ؼ�������
    Dim m As Long
    Dim i As Integer
    m = 0
    For i = 1 To lvPreSell.ListItems.count
        m = m + lvPreSell.ListItems(i).SubItems(IT_SumTicketNum)
    Next i
    
        If Not lvBus.SelectedItem Is Nothing Then
            Dim liTemp As ListItem
            Set liTemp = lvBus.SelectedItem
            
                If lvPreSell.ListItems.count > 1 Then
                    '���Ԥ�۵�վ��������һ�����Ͳ���ʾ
                    Exit Sub
                End If
                If Not lvPreSell.SelectedItem Is Nothing Then
                    If GetStationNameInCbo(cboEndStation.Text) <> lvPreSell.SelectedItem.SubItems(IT_EndStation) Then
                        '�����ǰ��վ����Ԥ���е�վ�㲻ͬ������ʾ
                        Exit Sub
                        
                    End If
                End If
            '********************
            '������ʾ��
            
                'SetStationAndTime liTemp.ListSubItems(ID_EndStation), flblSellDate.Caption & " " & liTemp.ListSubItems(ID_OffTime), txtFreeSell.Value + txtFullSell.Value + txtHalfSell.Value
                If nKey <> 27 Then
                SetStationAndTime GetStationNameInCbo(cboEndStation.Text), flblSellDate.Caption & " " & liTemp.ListSubItems(ID_OffTime), m + Val(txtFullSell.Text) + Val(txtHalfSell.Text) + Val(txtPreferentialSell.Text)
                End If
                nTicketCount = m + Val(txtFullSell.Text) + Val(txtHalfSell.Text) + Val(txtPreferentialSell.Text)
            '********************
        End If
End Sub

Private Sub DispStation()
    
    If Not lvBus.SelectedItem Is Nothing Then
        Dim liTemp As ListItem
        Set liTemp = lvBus.SelectedItem
            If lvPreSell.ListItems.count > 1 Then
                '���Ԥ�۵�վ��������һ�����Ͳ���ʾ
                Exit Sub
            End If
            If Not lvPreSell.SelectedItem Is Nothing Then
                If GetStationNameInCbo(cboEndStation.Text) <> lvPreSell.SelectedItem.SubItems(IT_EndStation) Then
                    '�����ǰ��վ����Ԥ���е�վ�㲻ͬ������ʾ
                    Exit Sub
                    
                End If
            End If
        '********************
        '������ʾ��
'            SetStationAndTime liTemp.ListSubItems(ID_EndStation), flblSellDate.Caption & " " & liTemp.ListSubItems(ID_OffTime)
Dim q As Integer
If nKey <> 27 Then
            SetStationAndTime GetStationNameInCbo(cboEndStation.Text), flblSellDate.Caption & " " & liTemp.ListSubItems(ID_OffTime)
End If
        '********************
    End If
End Sub
