VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{695ABF14-B2D8-11D2-A5ED-DE08DCF33612}#1.2#0"; "asfcombo.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmUpdateSheet 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7035
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   10815
   ControlBox      =   0   'False
   HelpContextID   =   7000060
   Icon            =   "frmUpdateSheet.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   10815
   Begin VB.ComboBox cboAllVehicle 
      Height          =   300
      Left            =   8490
      TabIndex        =   54
      Text            =   "Combo1"
      Top             =   840
      Visible         =   0   'False
      Width           =   1545
   End
   Begin RTComctl3.CoolButton cmdcancel 
      Height          =   315
      Left            =   7920
      TabIndex        =   45
      Top             =   6420
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "取 消(&C)"
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
      MICON           =   "frmUpdateSheet.frx":08CA
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
      Height          =   315
      Left            =   9060
      TabIndex        =   38
      Top             =   6420
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "提 交(&T)"
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
      MICON           =   "frmUpdateSheet.frx":08E6
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
      BackColor       =   &H00FFFFFF&
      Caption         =   "签发单信息"
      Height          =   555
      Left            =   540
      TabIndex        =   25
      Top             =   5820
      Width           =   9555
      Begin VB.Label lblOwner 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "罗汉"
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
         Left            =   8400
         TabIndex        =   50
         Top             =   270
         Width           =   390
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车主:"
         Height          =   180
         Left            =   7800
         TabIndex        =   49
         Top             =   270
         Width           =   450
      End
      Begin VB.Label lblBusID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "11023"
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
         Left            =   7080
         TabIndex        =   48
         Top             =   270
         Width           =   525
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车次:"
         Height          =   180
         Left            =   6480
         TabIndex        =   47
         Top             =   270
         Width           =   510
      End
      Begin VB.Label lblCalWeight 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2313"
         Height          =   180
         Left            =   1020
         TabIndex        =   33
         Top             =   270
         Width           =   360
      End
      Begin VB.Label lblOverNumber 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "123123"
         Height          =   180
         Left            =   4050
         TabIndex        =   32
         Top             =   270
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "超重件数:"
         Height          =   180
         Index           =   13
         Left            =   3120
         TabIndex        =   31
         Top             =   270
         Width           =   810
      End
      Begin VB.Label lblActWeight 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "12312"
         Height          =   180
         Left            =   2460
         TabIndex        =   30
         Top             =   270
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "行包单数:"
         Height          =   180
         Index           =   3
         Left            =   4740
         TabIndex        =   29
         Top             =   270
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "总实重:"
         Height          =   180
         Index           =   9
         Left            =   1710
         TabIndex        =   28
         Top             =   270
         Width           =   630
      End
      Begin VB.Label lblBillNumber 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "123123"
         Height          =   180
         Left            =   5640
         TabIndex        =   27
         Top             =   270
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "总计重:"
         Height          =   180
         Index           =   8
         Left            =   240
         TabIndex        =   26
         Top             =   270
         Width           =   630
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4635
      Left            =   540
      TabIndex        =   5
      Top             =   1140
      Width           =   9555
      Begin FCmbo.asFlatCombo cboBusStartTime 
         Height          =   270
         Left            =   4470
         TabIndex        =   53
         Top             =   150
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   476
         ButtonDisabledForeColor=   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
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
         ListSelectOnly  =   -1  'True
         Style           =   1
         Registered      =   -1  'True
         OfficeXPColors  =   -1  'True
      End
      Begin FText.asFlatTextBox txtEndStation 
         Height          =   285
         Left            =   1350
         TabIndex        =   41
         Top             =   150
         Visible         =   0   'False
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   503
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
         Text            =   "杭州"
         ButtonBackColor =   -2147483633
         ButtonVisible   =   -1  'True
         Registered      =   -1  'True
      End
      Begin FText.asFlatTextBox txtVehicle 
         Height          =   285
         Left            =   7920
         TabIndex        =   39
         Top             =   150
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
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
         Text            =   "浙B02237"
         ButtonBackColor =   -2147483633
         ButtonVisible   =   -1  'True
         Registered      =   -1  'True
      End
      Begin MSComctlLib.ListView LvInfo 
         Height          =   2835
         Left            =   0
         TabIndex        =   21
         Top             =   1110
         Width           =   9555
         _ExtentX        =   16854
         _ExtentY        =   5001
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "标签号"
            Object.Width           =   1941
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "托运人"
            Object.Width           =   1536
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "电 话"
            Object.Width           =   1747
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "收件人(单位)"
            Object.Width           =   2558
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "电 话"
            Object.Width           =   2401
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "地 址"
            Object.Width           =   2964
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "交付方式"
            Object.Width           =   1147
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "件数"
            Object.Width           =   1147
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "运 费"
            Object.Width           =   1325
         EndProperty
      End
      Begin VB.Label lblVehicle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "浙B10203"
         Height          =   180
         Left            =   8040
         TabIndex        =   52
         Top             =   210
         Width           =   720
      End
      Begin VB.Label lblEndStation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "杭州"
         Height          =   180
         Left            =   1470
         TabIndex        =   51
         Top             =   210
         Width           =   360
      End
      Begin VB.Label lblBagNumber 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "6"
         Height          =   180
         Left            =   8400
         TabIndex        =   46
         Top             =   4020
         Width           =   90
      End
      Begin VB.Label lblTotalPrice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "45"
         Height          =   180
         Left            =   9030
         TabIndex        =   44
         Top             =   4020
         Width           =   180
      End
      Begin VB.Label lblTotalPriceBig 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "叁拾伍元"
         Height          =   210
         Left            =   3030
         TabIndex        =   43
         Top             =   4350
         Width           =   720
      End
      Begin VB.Label lblBusStartTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "03-15 14:30"
         Height          =   180
         Left            =   4500
         TabIndex        =   40
         Top             =   180
         Width           =   990
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "合计人民币(大写)"
         Height          =   180
         Left            =   60
         TabIndex        =   23
         Top             =   4350
         Width           =   1440
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "合                    计"
         Height          =   180
         Left            =   3330
         TabIndex        =   22
         Top             =   3990
         Width           =   2160
      End
      Begin VB.Line Line18 
         X1              =   8730
         X2              =   8730
         Y1              =   3930
         Y2              =   4260
      End
      Begin VB.Line Line17 
         X1              =   8130
         X2              =   8130
         Y1              =   3930
         Y2              =   4260
      End
      Begin VB.Line Line16 
         X1              =   30
         X2              =   9540
         Y1              =   4620
         Y2              =   4620
      End
      Begin VB.Line Line15 
         X1              =   30
         X2              =   9540
         Y1              =   4260
         Y2              =   4260
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "运  费"
         Height          =   180
         Left            =   8850
         TabIndex        =   20
         Top             =   720
         Width           =   540
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "件数"
         Height          =   180
         Left            =   8220
         TabIndex        =   19
         Top             =   720
         Width           =   360
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "方式"
         Height          =   180
         Left            =   7620
         TabIndex        =   18
         Top             =   870
         Width           =   360
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "交付"
         Height          =   180
         Left            =   7620
         TabIndex        =   17
         Top             =   570
         Width           =   360
      End
      Begin VB.Line Line14 
         X1              =   8640
         X2              =   8640
         Y1              =   480
         Y2              =   1110
      End
      Begin VB.Line Line13 
         X1              =   8100
         X2              =   8100
         Y1              =   480
         Y2              =   1110
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "地      址"
         Height          =   180
         Left            =   6180
         TabIndex        =   16
         Top             =   870
         Width           =   900
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "电  话"
         Height          =   180
         Left            =   4800
         TabIndex        =   15
         Top             =   870
         Width           =   540
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "收件人(单位)"
         Height          =   180
         Left            =   3120
         TabIndex        =   14
         Top             =   870
         Width           =   1080
      End
      Begin VB.Line Line12 
         X1              =   5820
         X2              =   5820
         Y1              =   810
         Y2              =   1110
      End
      Begin VB.Line Line11 
         X1              =   4410
         X2              =   4410
         Y1              =   810
         Y2              =   1110
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "收         件         方"
         Height          =   180
         Left            =   4050
         TabIndex        =   13
         Top             =   540
         Width           =   2160
      End
      Begin VB.Line Line10 
         X1              =   2910
         X2              =   7470
         Y1              =   810
         Y2              =   810
      End
      Begin VB.Line Line9 
         X1              =   1890
         X2              =   1890
         Y1              =   810
         Y2              =   1110
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "电  话"
         Height          =   180
         Left            =   2130
         TabIndex        =   12
         Top             =   870
         Width           =   540
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "托运人"
         Height          =   180
         Left            =   1230
         TabIndex        =   11
         Top             =   870
         Width           =   540
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "托  运  方"
         Height          =   180
         Left            =   1470
         TabIndex        =   10
         Top             =   570
         Width           =   900
      End
      Begin VB.Line Line8 
         X1              =   1110
         X2              =   2880
         Y1              =   810
         Y2              =   810
      End
      Begin VB.Line Line7 
         X1              =   7470
         X2              =   7470
         Y1              =   90
         Y2              =   1110
      End
      Begin VB.Line Line6 
         X1              =   5820
         X2              =   5820
         Y1              =   90
         Y2              =   480
      End
      Begin VB.Line Line5 
         X1              =   4410
         X2              =   4410
         Y1              =   90
         Y2              =   480
      End
      Begin VB.Line Line4 
         X1              =   2910
         X2              =   2910
         Y1              =   90
         Y2              =   1140
      End
      Begin VB.Line Line3 
         X1              =   1080
         X2              =   1080
         Y1              =   90
         Y2              =   1110
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "标签号"
         Height          =   180
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   540
      End
      Begin VB.Line Line2 
         X1              =   30
         X2              =   9540
         Y1              =   1110
         Y2              =   1110
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车     号"
         Height          =   180
         Left            =   6210
         TabIndex        =   8
         Top             =   210
         Width           =   810
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "发车时间"
         Height          =   180
         Left            =   3240
         TabIndex        =   7
         Top             =   210
         Width           =   720
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "到达站"
         Height          =   180
         Left            =   240
         TabIndex        =   6
         Top             =   210
         Width           =   540
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   9540
         Y1              =   480
         Y2              =   480
      End
   End
   Begin FCmbo.asFlatCombo cboSheetID 
      Height          =   300
      Left            =   8460
      TabIndex        =   1
      Top             =   390
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   529
      ButtonDisabledForeColor=   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   255
      ButtonHotBackColor=   8421504
      ButtonPressedBackColor=   0
      Text            =   ""
      ButtonBackColor =   8421504
      Registered      =   -1  'True
      OfficeXPColors  =   -1  'True
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2003 年 01 月 15 日"
      Height          =   180
      Left            =   4470
      TabIndex        =   42
      Top             =   960
      Width           =   1710
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "操作员："
      Height          =   180
      Left            =   3660
      TabIndex        =   37
      Top             =   6450
      Width           =   720
   End
   Begin VB.Label lblChecker 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "wjb"
      Height          =   180
      Left            =   4410
      TabIndex        =   36
      Top             =   6420
      Width           =   270
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "操作时间："
      Height          =   180
      Left            =   630
      TabIndex        =   35
      Top             =   6450
      Width           =   900
   End
   Begin VB.Label lblOprationTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "12:50"
      Height          =   180
      Left            =   1590
      TabIndex        =   34
      Top             =   6450
      Width           =   450
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NO."
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   7980
      TabIndex        =   24
      Top             =   360
      Width           =   540
   End
   Begin VB.Label lblStartStation 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "北站"
      Height          =   180
      Left            =   1290
      TabIndex        =   4
      Top             =   960
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "始发站:"
      Height          =   180
      Left            =   660
      TabIndex        =   3
      Top             =   960
      Width           =   630
   End
   Begin VB.Label lblAcceptType 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "快件"
      Height          =   180
      Left            =   1110
      TabIndex        =   2
      Top             =   660
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "行包装车交接清单"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   3810
      TabIndex        =   0
      Top             =   330
      Width           =   3000
   End
End
Attribute VB_Name = "frmUpdateSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboBusStartTime_Change()
    Dim i As Integer
    Dim j As Integer
    Dim nlen As Integer
    Dim szTemp() As String
    Dim szaTemp() As String
    Dim Count As Integer
On Error GoTo ErrHandle
  If cboBusStartTime.Text <> "" Then
        moCarrySheet.BusDate = Date
        moCarrySheet.RefreshBusInfo Trim(txtVehicle.Text), CDate(CStr(Format(Date, "yy-mm-dd")) + " " + CStr(Format(Trim(cboBusStartTime.Text), "hh:mm")) + ":00"), Trim(lblAcceptType.Caption)
        lblOwner.Caption = moCarrySheet.BusOwnerName
        lblBusID.Caption = moCarrySheet.BusID
   End If
  Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub cboBusStartTime_Click()
cboBusStartTime_Change
End Sub

Private Sub cboBusStartTime_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then
    txtEndStation.SetFocus
 End If
 
End Sub

Private Sub cboSheetID_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        RefreshPreSheet
    End If
End Sub

Private Sub RefreshPreSheet()
 Dim rsTemp As Recordset
 Dim lvItem As ListItem
 Dim i As Integer
 Dim mSumPrice As Double
  On Error GoTo here
' If KeyCode = 13 Then
  If Trim(cboSheetID.Text) = "" Then
     MsgBox "已签发的签发单号不能为空！", vbExclamation, Me.Caption
     cboSheetID.SetFocus
     Exit Sub
  End If
  Set rsTemp = moLugSvr.GetOldSheetRs(Trim(cboSheetID.Text))
  If rsTemp.RecordCount = 0 Then
     MsgBox "没有此签发单号！", vbExclamation, Me.Caption
     cboSheetID.SetFocus
     Exit Sub
  End If
  If Trim(rsTemp!vehicle_id) <> "" Then
'     MsgBox "此签发单已有车次，不允许修改！", vbInformation, Me.Caption
     cmdOk.Enabled = True
     lblOwner.Caption = Trim(rsTemp!owner_name)
     lblBusID.Caption = Trim(rsTemp!BusID)
     lblBusStartTime.Caption = CStr(Format(rsTemp!bus_start_time, "yy-mm-dd hh:mm"))
     lblEndStation.Caption = Trim(rsTemp!des_station_name)
     txtVehicle.Text = Trim(rsTemp!license_tag_no)
     lblEndStation.Visible = True
     lblVehicle.Visible = False
     txtEndStation.Visible = False
     txtVehicle.Visible = True
     lblBusStartTime.Visible = True
     cboBusStartTime.Visible = False
   Else
'     lblEndStation.Visible = False
'     lblVehicle.Visible = False
'     txtEndStation.Visible = True
'     lblOwner.Caption = Trim(rsTemp!Owner_name)
     lblBusID.Caption = Trim(rsTemp!BusID)
     lblBusStartTime.Caption = CStr(Format(rsTemp!bus_start_time, "yy-mm-dd hh:mm"))
     lblEndStation.Caption = Trim(rsTemp!des_station_name)
     txtVehicle.Text = ""
     lblOwner.Caption = ""
     txtVehicle.Visible = True
'     cboBusStartTime.Visible = True
     lblBusStartTime.Visible = True
     cmdOk.Enabled = True
  End If
'  moCarrySheet.AddNew
'  lblSheetID.Caption = Trim(rsTemp!sheet_id)
  lblCalWeight.Caption = rsTemp!cal_weight
  lblActWeight.Caption = rsTemp!fact_weight
  lblOverNumber.Caption = rsTemp!over_number
  lblBagNumber.Caption = rsTemp!luggage_number
  lblBillNumber.Caption = rsTemp!baggage_number
  lblChecker.Caption = rsTemp!checker
  lblOprationTime.Caption = rsTemp!sheet_make_time
  lblStartStation.Caption = Trim(rsTemp!start_station_name)
  lblAcceptType.Caption = GetLuggageTypeString(Trim(rsTemp!accept_type))
  '填定lvInfo
  mSumPrice = 0
  LvInfo.ListItems.clear
  For i = 1 To rsTemp.RecordCount
   Set lvItem = LvInfo.ListItems.Add(, , Trim(rsTemp!start_label_id) + "-" + Trim(rsTemp!end_label_id))
      lvItem.SubItems(1) = Trim(rsTemp!Shipper)
      lvItem.SubItems(2) = Trim(rsTemp!shipper_phone)
      lvItem.SubItems(3) = Trim(rsTemp!Picker)
      lvItem.SubItems(4) = Trim(rsTemp!picker_phone)
      lvItem.SubItems(5) = Trim(rsTemp!picker_address)
      lvItem.SubItems(6) = Trim(rsTemp!pick_type)
      lvItem.SubItems(7) = Trim(rsTemp!baggage_number)
      lvItem.SubItems(8) = Trim(rsTemp!shipperprice)
      mSumPrice = mSumPrice + CDbl(lvItem.SubItems(8))
      rsTemp.MoveNext
   Next i
  lblTotalPrice.Caption = CStr(mSumPrice)
  lblTotalPriceBig.Caption = GetNumber(mSumPrice)
'      txtVehicle.SetFocus
' End If
  Exit Sub
here:
ShowErrorMsg
End Sub

Private Sub cboSheetID_LostFocus()
'    RefreshPreSheet
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
'Dim mAnswer
'On Error GoTo here
'    If txtVehicle.Text = "" Or (Len(txtVehicle.Text) < 7) Then
'       MsgBox "请正确填写车牌信息！", vbExclamation, Me.Caption
'       Exit Sub
'    End If
'    mAnswer = MsgBox("您输入的车号是  " & Trim(txtVehicle.Text) & "   是否确定更改？", vbInformation + vbYesNo, Me.Caption)
'    If mAnswer = vbYes Then
''    moCarrySheet.SheetID = Trim(cboSheetID.Text)
'    moLugSvr.UpdateCarryLuggage Trim(cboSheetID.Text), ResolveDisplay(txtVehicle.Text), ResolveDisplayEx(txtVehicle.Text)
'
'    Unload Me
'    Else
'    Exit Sub
'    End If
'Exit Sub
'here:
'ShowErrorMsg
End Sub


Private Sub Form_Load()
  Dim rsTemp As Recordset
  Dim i As Integer
  
     AlignFormPos Me
     Top = (mdiMain.Height - Me.Height) / 4
     Left = (mdiMain.Width - Me.Width) / 2
'     Height = frmCarryLuggage.fraOutLine.Height
'     Width = frmCarryLuggage.fraOutLine.Width
     lblDate.Caption = CStr(Year(Date)) + " 年 " + CStr(Month(Date)) + " 月 " + CStr(Day(Date)) + " 日"
     FormClear
     '取得没有填车次的签发单
     Set rsTemp = moLugSvr.PreCarryLuggage()
     If rsTemp.RecordCount > 0 Then
       For i = 1 To rsTemp.RecordCount
        cboSheetID.AddItem Trim(rsTemp!sheet_id)
        rsTemp.MoveNext
       Next i
       cboSheetID.ListIndex = 0
     End If
     GetAllVehicle
End Sub

Public Sub GetAllVehicle()
Dim szaTemp() As String
Dim i As Integer
Dim nlen As Integer
'Dim m_obase As BaseInfo
  '得到所有车辆信息
  cboAllVehicle.clear
  
'  m_obase.Init m_oAUser
  szaTemp = m_obase.GetVehicle()
  nlen = ArrayLength(szaTemp)
  If nlen > 0 Then
   For i = 1 To nlen
     cboAllVehicle.AddItem MakeDisplayString(szaTemp(i, 1), szaTemp(i, 2))
    
   Next i
  End If
End Sub

Private Sub FormClear()
    LvInfo.ListItems.clear
    lblOwner.Caption = ""
    lblBusID.Caption = ""
    lblBusStartTime.Caption = ""
    lblEndStation.Caption = ""
    lblVehicle.Caption = ""
    lblEndStation.Visible = True
    lblVehicle.Visible = True
'    txtEndStation.Visible = True
    txtVehicle.Visible = True
    lblBusStartTime.Visible = True
'    cboBusStartTime.Visible = False
    txtEndStation.Text = ""
    txtVehicle.Text = ""
  lblCalWeight.Caption = ""
  lblActWeight.Caption = ""
  lblOverNumber.Caption = ""
  lblBagNumber.Caption = ""
  lblBillNumber.Caption = ""
  lblTotalPrice.Caption = ""
  lblTotalPriceBig.Caption = ""
  lblChecker.Caption = ""
  lblOprationTime.Caption = ""
  lblStartStation.Caption = ""
  lblAcceptType.Caption = ""
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Unload Me
End Sub

Private Sub txtEndStation_ButtonClick()
On Error GoTo ErrHandle
    Dim oShell As New STShell.CommDialog
    Dim aszTemp() As String
    oShell.Init m_oAUser
    aszTemp = oShell.SelectStation()
    Set oShell = Nothing
    If ArrayLength(aszTemp) = 0 Then Exit Sub
    txtEndStation.Text = Trim(aszTemp(1, 2))
    
Exit Sub
ErrHandle:
ShowErrorMsg
End Sub

Private Sub txtEndStation_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   cmdOk.SetFocus
End If
End Sub

Private Sub txtVehicle_ButtonClick()
   If txtVehicle.Enabled Then
'    cboBusStartTime.Visible = False
    frmSearchVechile.mFormNum = 1
    frmSearchVechile.StartSearchIndex = mnLastSearchIndex
    frmSearchVechile.Show vbModal
    mnLastSearchIndex = cboAllVehicle.ListIndex
    
   End If
End Sub

Private Sub txtVehicle_Change()
    GetVehicleOwner
End Sub

Private Sub txtVehicle_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cmdOk.SetFocus
    End If
End Sub

Private Sub GetVehicleOwner()
    Dim rsTemp As Recordset
Dim i As Integer

On Error GoTo ErrHandle
'  If KeyCode = 13 Then
    If txtVehicle.Text <> "" Then
'        moCarrySheet.BusDate = Date
'       Set rsTemp = moCarrySheet.RefreshBusStartTime(txtVehicle.Text)
'          cboBusStartTime.Clear
'       If rsTemp.RecordCount > 0 Then
'         For i = 1 To rsTemp.RecordCount
'          cboBusStartTime.AddItem Format(rsTemp!bus_start_time, "mm-dd hh:mm")
'          rsTemp.MoveNext
'         Next i
'          cboBusStartTime.ListIndex = 0
'       End If
'    End If
'    Set rsTemp = Nothing
'    cboBusStartTime.Visible = True
'    cboBusStartTime.SetFocus
'得到车辆,车主
    Set rsTemp = moLugSvr.GetOwnerName(ResolveDisplay(txtVehicle.Text))
    If rsTemp.RecordCount > 0 Then
        lblOwner.Caption = FormatDbValue(rsTemp!owner_name)
    End If
'    cmdOK.SetFocus
  End If
' End If
Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub txtVehicle_LostFocus()
    GetVehicleOwner
End Sub
