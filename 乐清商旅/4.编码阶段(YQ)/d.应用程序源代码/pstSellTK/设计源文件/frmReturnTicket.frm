VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmReturnTicket 
   BackColor       =   &H8000000C&
   Caption         =   "��Ʊ"
   ClientHeight    =   7500
   ClientLeft      =   825
   ClientTop       =   2700
   ClientWidth     =   11250
   ForeColor       =   &H00000000&
   HelpContextID   =   4000180
   Icon            =   "frmReturnTicket.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7500
   ScaleWidth      =   11250
   WindowState     =   2  'Maximized
   Begin VB.Frame fraOutLine 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   7095
      Left            =   480
      TabIndex        =   12
      Top             =   540
      Width           =   10275
      Begin VB.CheckBox chkNoRatio 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ȫ����Ʊ(A)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5355
         TabIndex        =   51
         Top             =   180
         Width           =   1755
      End
      Begin VB.TextBox txtEndTicketNo 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   3
         Top             =   630
         Width           =   2490
      End
      Begin RTComctl3.CoolButton cmdReturnTkt 
         Height          =   405
         Left            =   7500
         TabIndex        =   4
         Top             =   3090
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   714
         BTYPE           =   3
         TX              =   "��Ʊ(&T)"
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
         MICON           =   "frmReturnTicket.frx":014A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox txtTicketNo 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   1
         Top             =   120
         Width           =   2490
      End
      Begin VB.Frame fraTktInfoChange 
         BackColor       =   &H00E0E0E0&
         Caption         =   "��Ʊ��Ϣ"
         Height          =   2775
         Left            =   120
         TabIndex        =   24
         Top             =   1170
         Width           =   7155
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   3975
            TabIndex        =   50
            Top             =   2475
            Visible         =   0   'False
            Width           =   525
         End
         Begin VB.Label lblVehicleType 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   4890
            TabIndex        =   49
            Top             =   2445
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ʊ��:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   135
            TabIndex        =   48
            Top             =   240
            Width           =   525
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����ʱ��:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   135
            TabIndex        =   47
            Top             =   1335
            Width           =   945
         End
         Begin VB.Label lblOperatorChange 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ƱԱ:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   135
            TabIndex        =   46
            Top             =   1710
            Width           =   735
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��վ:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   135
            TabIndex        =   45
            Top             =   570
            Width           =   525
         End
         Begin VB.Label lblScheduleChange 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   135
            TabIndex        =   44
            Top             =   945
            Width           =   525
         End
         Begin VB.Label lblTimeChange 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��Ʊʱ��:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   135
            TabIndex        =   43
            Top             =   2460
            Width           =   945
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ʊ��:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   135
            TabIndex        =   42
            Top             =   2085
            Width           =   525
         End
         Begin VB.Label lblTicketID 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "A000013459"
            BeginProperty Font 
               Name            =   "����"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   1170
            TabIndex        =   41
            Tag             =   "lblCurrentTktNum"
            Top             =   210
            Width           =   1500
         End
         Begin VB.Label lblSeatNo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "01"
            BeginProperty Font 
               Name            =   "����"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   4890
            TabIndex        =   40
            Top             =   1680
            Width           =   300
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��λ��:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   3975
            TabIndex        =   39
            Top             =   1710
            Width           =   735
         End
         Begin VB.Label lblTypeChange 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ʊ��:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   3975
            TabIndex        =   38
            Top             =   1335
            Width           =   525
         End
         Begin VB.Label label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��վ:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   3975
            TabIndex        =   37
            Top             =   570
            Width           =   525
         End
         Begin VB.Label lblStateChange 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "״̬:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   3975
            TabIndex        =   36
            Top             =   2085
            Width           =   525
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   3975
            TabIndex        =   35
            Top             =   945
            Width           =   525
         End
         Begin VB.Label lblBusID 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "25101"
            BeginProperty Font 
               Name            =   "����"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   1170
            TabIndex        =   34
            Top             =   915
            Width           =   750
         End
         Begin VB.Label lblEndStation 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   4890
            TabIndex        =   33
            Top             =   540
            Width           =   570
         End
         Begin VB.Label lblStartStation 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������վ"
            BeginProperty Font 
               Name            =   "����"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   1170
            TabIndex        =   32
            Top             =   540
            Width           =   1140
         End
         Begin VB.Label lblSeller 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   1170
            TabIndex        =   31
            Top             =   1680
            Width           =   570
         End
         Begin VB.Label lblTicketType 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ȫƱ"
            BeginProperty Font 
               Name            =   "����"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   4920
            TabIndex        =   30
            Top             =   1320
            Width           =   570
         End
         Begin VB.Label lblSellTime 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "2002-07-15 07:00:00"
            BeginProperty Font 
               Name            =   "����"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   1170
            TabIndex        =   29
            Top             =   2445
            Width           =   2850
         End
         Begin VB.Label lblDate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "2002-07-15"
            BeginProperty Font 
               Name            =   "����"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   4890
            TabIndex        =   28
            Top             =   915
            Width           =   1500
         End
         Begin VB.Label lblTicketPrice 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "37.50"
            BeginProperty Font 
               Name            =   "����"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   1170
            TabIndex        =   27
            Top             =   2055
            Width           =   750
         End
         Begin VB.Label lblStatus 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�����۳�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   4890
            TabIndex        =   26
            Top             =   2055
            Width           =   1140
         End
         Begin VB.Label lblOffTime 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "10:00"
            BeginProperty Font 
               Name            =   "����"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   1170
            TabIndex        =   25
            Top             =   1305
            Width           =   750
         End
      End
      Begin VB.TextBox txtCredenceID 
         Alignment       =   1  'Right Justify
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
         Left            =   8580
         TabIndex        =   7
         Top             =   2580
         Width           =   1455
      End
      Begin VB.ComboBox cboFeesRatio 
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
         ItemData        =   "frmReturnTicket.frx":0166
         Left            =   8820
         List            =   "frmReturnTicket.frx":0168
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   150
         Width           =   1215
      End
      Begin MSComctlLib.ListView lvTicketInfo 
         Height          =   2625
         Left            =   120
         TabIndex        =   11
         Top             =   4290
         Width           =   10065
         _ExtentX        =   17754
         _ExtentY        =   4630
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin RTComctl3.CoolButton cmdResumeReturn 
         Height          =   405
         Left            =   7500
         TabIndex        =   5
         Top             =   3570
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   714
         BTYPE           =   3
         TX              =   "ȡ����Ʊ(&D)"
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
         MICON           =   "frmReturnTicket.frx":016A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Line Line2 
         Visible         =   0   'False
         X1              =   4770
         X2              =   7230
         Y1              =   1860
         Y2              =   1860
      End
      Begin VB.Label lblEndTktNum 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����Ʊ��(&E):"
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
         Top             =   720
         Width           =   1440
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������Ʊ�б�(&I):"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   150
         TabIndex        =   10
         Top             =   4020
         Width           =   1890
      End
      Begin VB.Line Line1 
         X1              =   7560
         X2              =   10020
         Y1              =   1950
         Y2              =   1950
      End
      Begin VB.Label lblTotalReturnMoney 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   9060
         TabIndex        =   23
         Top             =   2100
         Width           =   690
      End
      Begin VB.Label lblTotalFees 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   9060
         TabIndex        =   22
         Top             =   1485
         Width           =   690
      End
      Begin VB.Label lblTotalPrice 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   405
         Left            =   9060
         TabIndex        =   21
         Top             =   1005
         Width           =   690
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ʊ��Ϣ:"
         Height          =   180
         Left            =   3060
         TabIndex        =   20
         Top             =   2940
         Width           =   810
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ƱƱ��:"
         Height          =   180
         Left            =   4770
         TabIndex        =   19
         Top             =   1275
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.Label lblCurrectTicketPrice 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6390
         TabIndex        =   18
         Top             =   1200
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����ѱ���:(%)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   7320
         TabIndex        =   8
         Top             =   240
         Width           =   1470
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�˿�������:"
         Height          =   180
         Left            =   4770
         TabIndex        =   17
         Top             =   1590
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label lblReturnCharge 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6390
         TabIndex        =   16
         Top             =   1515
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ʊƾ֤��:"
         Height          =   180
         Left            =   7560
         TabIndex        =   6
         Top             =   2670
         Width           =   990
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ӧ��Ʊ��:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   7560
         TabIndex        =   15
         Top             =   2175
         Width           =   945
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ʊ��:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   7560
         TabIndex        =   14
         Top             =   1095
         Width           =   735
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   7560
         TabIndex        =   13
         Top             =   1560
         Width           =   945
      End
      Begin VB.Line Line4 
         Index           =   0
         X1              =   7305
         X2              =   10150
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label lblOldTktNum 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ƱƱ��(&Z):"
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
         Top             =   210
         Width           =   1440
      End
   End
End
Attribute VB_Name = "frmReturnTicket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'SINGLE_RETURN =1��ʾֻ������Ʊ


'��Ʊ�õ�ö��
Private Enum ReturnTicketInfo
    RT_BusID = 1
    RT_StartStation = 2
    RT_EndStation = 3
    RT_OffTime = 4
    RT_SeatNo = 5
    RT_Status = 6
    RT_ReturnRatio = 7
    RT_ReturnMoney = 8
    RT_TicketType = 9
    RT_TicketPrice = 10
    RT_Date = 11
    RT_SellTime = 12
    RT_Seller = 13
End Enum




Private Sub chkForce_Click()
    EnableReturnButton
End Sub




Private Sub cmdRefreshSomeTkt_Click()
On Error GoTo here
    SerialReturnTkt
    EnableReturnButton
    If cmdReturnTkt.Enabled Then cmdReturnTkt.SetFocus
On Error GoTo 0
Exit Sub
here:
    ShowErrorMsg
End Sub

Private Sub chkNoRatio_Click()
    If chkNoRatio.Value = vbChecked Then
        cboFeesRatio.Text = 0
    Else
        cmdRefreshSomeTkt_Click
    End If
End Sub

Private Sub cmdResumeReturn_Click()
    Dim szTicketID() As String
    Dim iCount As Integer
    On Error GoTo here
    If MsgBox("ȷ��Ҫȡ����Ʊ��", vbInformation + vbYesNo, "ȡ����Ʊ") = vbYes Then

        szTicketID = GetAllTickets
        
        m_oSell.ResumeReturnTicket szTicketID
        '����״̬Ϊ����
        For iCount = 1 To lvTicketInfo.ListItems.count
            lvTicketInfo.ListItems(iCount).SubItems(RT_Status) = GetTicketStatusStr(ETicketStatus.ST_TicketNormal)
        Next iCount
        lblStatus.Caption = GetTicketStatusStr(ETicketStatus.ST_TicketNormal)
        
        
        ShowMsg "����ȡ����Ʊ�ɹ�"
        EnableReturnButton
        lvTicketInfo.ListItems.Clear
        SetDefaultLabel
        txtTicketNo.SetFocus
    End If
    Exit Sub
here:
    ShowErrorMsg
End Sub

Private Sub cmdReturnTkt_Click()
    '****��Ҫ��,�������е���Ʊ����ǿ����Ʊ
    Dim sgReturnCharge()  As Double
    Dim sgReturnRatio() As Double
    Dim TicketID() As String
    Dim sgTicketPrice As Double
    Dim iCount As Integer
    Dim szReturnSheetID As String
    Dim szCurrentTicketNo As String  '��ǰ��ӡƱ��
    Dim szReturnCount As String '��Ʊ����
    Dim sgTotalReturnCharge As Double
    Dim sgTotalTicketPrice As Double
    Dim g_oSysParam As New SystemParam
    Dim rfTemp() As RETURNFEES
    Dim iLen As Integer
    Dim szStartTime As Date
    
    
    On Error GoTo here
    szCurrentTicketNo = GetTicketNo()
    '������Ʊ
    If lvTicketInfo.ListItems.count = 0 Then
       Exit Sub
    End If
    If MsgBox("ȷ��Ҫ��Ʊ��", vbInformation + vbYesNo, "��Ʊ") = vbYes Then
    
        '�������ָ������Ʊʱ�䣬��û��ѡ����Ʊ���ʵ�Ȩ�ޣ�������Ʊ
        iLen = ArrayLength(g_oSysParam.GetReturnFees)
        ReDim rfTemp(1 To iLen)
        rfTemp = g_oSysParam.GetReturnFees
        If ResolveDisplay(lblOffTime.Caption) = cszScrollBus Then
            szStartTime = lblDate.Caption & " " & ResolveDisplayEx(lblOffTime.Caption)
        Else
            szStartTime = lblDate.Caption & " " & lblOffTime.Caption
        End If
'        If cboFeesRatio.Text = 0 Then
'            If DateAdd("s", Abs(rfTemp(iLen).iReturnTime), szStartTime) < Now Then
'                If m_oSell.SelectReturnIsValid Then
'                    MsgBox "�ѹ�ָ������Ʊʱ�䣬���û���[ѡ����Ʊ����]��Ȩ�ޣ����Բ�������Ʊ��", vbInformation, "ע��"
'                    Exit Sub
'                End If
'            End If
'        End If

        sgTotalReturnCharge = 0
        sgTotalTicketPrice = 0
        ReDim sgReturnCharge(1 To Val(lvTicketInfo.ListItems.count))
        ReDim TicketID(1 To Val(lvTicketInfo.ListItems.count))
        ReDim sgReturnRatio(1 To Val(lvTicketInfo.ListItems.count))
        For iCount = 1 To Val(lvTicketInfo.ListItems.count)
            With lvTicketInfo.ListItems(iCount)
                
                sgReturnCharge(iCount) = FormatMoney(FormatTail(.SubItems(RT_TicketPrice) * CSng(cboFeesRatio.Text) / 100))
                sgTotalReturnCharge = sgTotalReturnCharge + sgReturnCharge(iCount)
                TicketID(iCount) = .Text
                sgReturnRatio(iCount) = CDbl(cboFeesRatio.Text)
                sgTotalTicketPrice = FormatMoney(sgTotalTicketPrice + .SubItems(RT_TicketPrice))
            End With
        Next iCount
        szReturnSheetID = szCurrentTicketNo
        szReturnCount = CStr(Val(lvTicketInfo.ListItems.count))
        ShowStatusInMDI "���ڴ�����Ʊ"
        
        'ȫ����Ʊ ƾ֤��Ϊ��
        If Val(sgTotalReturnCharge) = 0 Then
            szReturnSheetID = ""
        End If
        
        If chkNoRatio.Value Then
            m_oSell.ForceReturnTicket TicketID, szReturnSheetID, sgReturnCharge, sgReturnRatio
        Else
            m_oSell.ReturnTicket TicketID, szReturnSheetID, sgReturnCharge, sgReturnRatio
        End If
        If m_oParam.IsPrintReturnChangeSheet Then
            If sgTotalReturnCharge > 0 Then
                
                '�����Ʊ�����Ѵ���0,�ʹ�ӡ
                #If SINGLE_RETURN = 1 Then
                Dim i As Integer
                For i = 1 To lvTicketInfo.ListItems.count
'                    PrintSingleReturnSheet TicketID(1), txtCredenceID.Text, lblReturnCharge.Caption, lblTicketPrice.Caption, IIf(lblOffTime.Caption = cszScrollBus, lblDate.Caption, lblDate.Caption & " " & lblOffTime.Caption), lblBusID.Caption, lblEndStation.Caption, lblTicketType.Caption, lblVehicleType.Caption, lblSeatNo.Caption, lblStartStation.Caption, lblDate.Caption
                    IncTicketNo
                    PrintSingleReturnSheet lvTicketInfo.ListItems(i).Text, MDISellTicket.lblTicketNo.Caption, FormatMoney(lvTicketInfo.ListItems(i).SubItems(RT_TicketPrice) - lvTicketInfo.ListItems(i).SubItems(RT_ReturnMoney)), FormatMoney(lvTicketInfo.ListItems(i).SubItems(RT_TicketPrice)), IIf(lvTicketInfo.ListItems(i).SubItems(RT_OffTime) = cszScrollBus, lvTicketInfo.ListItems(i).SubItems(RT_Date), lvTicketInfo.ListItems(i).SubItems(RT_Date) & " " & lvTicketInfo.ListItems(i).SubItems(RT_OffTime)), lvTicketInfo.ListItems(i).SubItems(RT_BusID), lvTicketInfo.ListItems(i).SubItems(RT_EndStation), lvTicketInfo.ListItems(i).SubItems(RT_TicketType), "", lvTicketInfo.ListItems(i).SubItems(RT_SeatNo), lvTicketInfo.ListItems(i).SubItems(RT_StartStation), lvTicketInfo.ListItems(i).SubItems(RT_Date)
                Next i
'                    PrintSingleReturnSheet txtCredenceID.Text, lblReturnCharge.Caption, lblTicketPrice.Caption, lblOffTime.Caption, lblBusID.Caption, lblEndStation.Caption, lblTicketType.Caption, lblVehicleType.Caption, lblSeatNo.Caption, lblStartStation.Caption, lblDate.Caption
                #Else
                    If lvTicketInfo.ListItems.count = 1 Then
                    PrintSingleReturnSheet TicketID(1), txtCredenceID.Text, lblReturnCharge.Caption, lblTicketPrice.Caption, IIf(lblOffTime.Caption = cszScrollBus, lblDate.Caption, lblDate.Caption & " " & lblOffTime.Caption), lblBusID.Caption, lblEndStation.Caption, lblTicketType.Caption, lblVehicleType.Caption, lblSeatNo.Caption, lblStartStation.Caption, lblDate.Caption
                        
 '                       PrintSingleReturnSheet txtCredenceID.Text, lblReturnCharge.Caption, lblTicketPrice.Caption, lblOffTime.Caption, lblBusID.Caption, lblEndStation.Caption, lblTicketType.Caption, lblVehicleType.Caption, lblSeatNo.Caption, lblStartStation.Caption, lblDate.Caption
                    Else
                        PrintReturnSheet szCurrentTicketNo, txtCredenceID.Text, sgTotalReturnCharge, sgTotalTicketPrice, lvTicketInfo.ListItems(1).SubItems(RT_OffTime), lvTicketInfo.ListItems(1).SubItems(RT_BusID), lblEndStation.Caption, szReturnCount
                    End If
                #End If
                
                'Ʊ�ż�һ
'                IncTicketNo
                
            End If
        End If
        RestoreStatusInMDI
        '����״̬Ϊ����
        For iCount = 1 To lvTicketInfo.ListItems.count
            lvTicketInfo.ListItems(iCount).SubItems(RT_Status) = GetTicketStatusStr(ETicketStatus.ST_TicketReturned + ETicketStatus.ST_TicketNormal)
        Next iCount
        lblStatus.Caption = GetTicketStatusStr(ETicketStatus.ST_TicketReturned + ETicketStatus.ST_TicketNormal)
        
        ShowMsg "��Ʊ�ɹ���"
        chkNoRatio.Value = vbUnchecked
        'lvTicketInfo.ListItems.Clear
        'SetDefaultLabel
        EnableReturnButton
        '����ƾ֤��Ϊ��ǰƱ��
        txtCredenceID.Text = GetTicketNo
        txtEndTicketNo.Text = ""
        txtTicketNo.SetFocus
        
    End If
    Exit Sub
here:
    ShowErrorMsg
End Sub

Private Sub Form_Activate()
On Error GoTo here
    m_nCurrentTask = RT_ReturnTicket
    m_szCurrentUnitID = Me.Tag
    MDISellTicket.SetFunAndUnit
    '����ƾ֤��Ϊ��ǰƱ��
    txtCredenceID.Text = GetTicketNo
    Exit Sub
here:
    ShowErrorMsg
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And Not (Me.ActiveControl Is txtTicketNo) And Not (Me.ActiveControl Is txtEndTicketNo) Then
        SendKeys "{TAB}"
    ElseIf KeyAscii = 27 Then
        lvTicketInfo.ListItems.Clear
        SetDefaultLabel
        txtEndTicketNo.Text = ""
        txtTicketNo.SetFocus
    ElseIf KeyAscii = Asc("+") Then
        '����һ��Ʊ�����ۼ�
'        txtCount.Value = 1
        txtTicketNo.SetFocus
        
    End If
    
End Sub

Private Sub Form_Load()
On Error GoTo here
Dim szStatus As Boolean
    
    txtTicketNo.MaxLength = 10
    FillHeaderColumn
    EnableReturnButton
    SelectReturnRatioValid
    InitReturnRatioValue
    SetDefaultLabel
    
    szStatus = m_oSell.IsAllReturn
    If szStatus = True Then
        chkNoRatio.Enabled = True
    Else
        chkNoRatio.Enabled = False
    End If
    
    #If SINGLE_RETURN = 1 Then
'        txtCount.Enabled = False
    #End If
    
Exit Sub
here:
    If szStatus = False Then
        chkNoRatio.Enabled = False
    Else
        ShowErrorMsg
    End If
End Sub

Private Sub Form_Resize()
    If MDISellTicket.ActiveForm Is Me Then
        If Not Me.WindowState = vbMaximized Then Me.WindowState = vbMaximized
        fraOutLine.Left = (Me.ScaleWidth - fraOutLine.Width) / 2
        fraOutLine.Top = (Me.ScaleHeight - fraOutLine.Height) / 2
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    m_clReturn.Remove GetEncodedKey(Me.Tag)
'    MDISellTicket.lblReturn.Value = vbUnchecked
    MDISellTicket.abMenuTool.Bands("mnuFunction").Tools("mnuReturnTkt").Checked = False
'    If Forms.Count = 2 Then
'        MDISellTicket.EnVisibleCheckLabel
'    End If
    
End Sub
'//////////////////////////////////
'������Ʊ��ť״̬
Private Sub EnableReturnButton()
    '������Ʊ��ť
    If lvTicketInfo.ListItems.count > 0 Then
        cmdReturnTkt.Enabled = True
        cmdResumeReturn.Enabled = True
    Else
        cmdReturnTkt.Enabled = False
        cmdResumeReturn.Enabled = False
    End If
End Sub

Private Sub lvTicketInfo_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    SortListView lvTicketInfo, ColumnHeader.Index
End Sub

Private Sub lvTicketInfo_ItemClick(ByVal Item As MSComctlLib.ListItem)
    ShowReturnInfo
End Sub

Private Sub lvTicketInfo_KeyPress(KeyAscii As Integer)
    If Not lvTicketInfo.SelectedItem Is Nothing Then
        If KeyAscii = vbKeyBack Then
            lvTicketInfo.ListItems.Remove lvTicketInfo.SelectedItem.Index
        End If
    End If
End Sub



'
'Private Sub txtCount_GotFocus()
'    txtCount.SelStart = 0
'    txtCount.SelLength = 100
'End Sub

'Private Sub txtCount_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn And txtTicketNo.Text <> "" And txtCount.Value > 0 Then
'        cmdRefreshSomeTkt_Click
'    End If
'
'End Sub

Private Sub txtCredenceID_Change()
    EnableReturnButton
End Sub


Private Sub cboFeesRatio_Click()
    Dim i As Integer
    
    If lvTicketInfo.ListItems.count = 0 Then
       Exit Sub
    End If
    '��������Ʊ����Ʊ������
    For i = 1 To lvTicketInfo.ListItems.count
        With lvTicketInfo.ListItems(i)
            .SubItems(RT_ReturnRatio) = cboFeesRatio.Text
            .SubItems(RT_ReturnMoney) = FormatMoney(.SubItems(RT_TicketPrice) - FormatTail(.SubItems(RT_ReturnRatio) * .SubItems(RT_TicketPrice) / 100))
            lblReturnCharge.Caption = FormatMoney(.SubItems(RT_TicketPrice) - .SubItems(RT_ReturnMoney))
            
        End With
    Next i
    GetReturnMoney
    EnableReturnButton
End Sub


Private Function ShowReturnInfo() As Boolean
    On Error GoTo here
    If Not lvTicketInfo.SelectedItem Is Nothing Then
        With lvTicketInfo.SelectedItem
            lblTicketID.Caption = .Text
            lblStartStation.Caption = .SubItems(RT_StartStation)
            lblBusID.Caption = .SubItems(RT_BusID)
            lblDate.Caption = .SubItems(RT_Date)
            lblStatus.Caption = .SubItems(RT_Status)
            lblEndStation.Caption = .SubItems(RT_EndStation)
            lblTicketType.Caption = .SubItems(RT_TicketType)
            lblTicketPrice.Caption = .SubItems(RT_TicketPrice)
            
            lblSeller.Caption = .SubItems(RT_Seller)
            lblSeatNo.Caption = .SubItems(RT_SeatNo)
            lblSellTime.Caption = .SubItems(RT_SellTime)
            lblOffTime.Caption = .SubItems(RT_OffTime)
            lblReturnCharge.Caption = .SubItems(RT_ReturnMoney)
            ShowReturnRatio
            
            '���ñ�ǩ
            lblCurrectTicketPrice.Caption = FormatMoney(.SubItems(RT_TicketPrice))
            lblReturnCharge.Caption = FormatMoney(FormatTail(lblCurrectTicketPrice.Caption * cboFeesRatio.Text / 100))
            
            
            
        End With
    End If
    ShowReturnInfo = True
    Exit Function
here:
    SetDefaultLabel
    ShowErrorMsg
    ShowReturnInfo = False

End Function

Private Sub txtCount_Change()
    EnableReturnButton
End Sub

Private Sub txtEndTicketNo_GotFocus()
    txtEndTicketNo.SelStart = 0
    txtEndTicketNo.SelLength = Len(txtEndTicketNo.Text)
End Sub

Private Sub txtEndTicketNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim szTemp As String
    Dim lTemp As Long
    On Error GoTo Error_Handel
    If Len(txtEndTicketNo.Text) >= TicketNoNumLen() Then
        If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
            lTemp = Right(txtEndTicketNo.Text, TicketNoNumLen())
            szTemp = Left(txtEndTicketNo.Text, Len(txtEndTicketNo.Text) - TicketNoNumLen())
            
            If KeyCode = vbKeyUp Then
                lTemp = lTemp + 1
            Else
                lTemp = lTemp - 1
                lTemp = IIf(lTemp > 0, lTemp, 0)
            End If
            txtEndTicketNo.Text = MakeEndTicketNo(lTemp, szTemp)
            KeyCode = 0
        End If
    End If
    Exit Sub
Error_Handel:
End Sub

Private Sub txtEndTicketNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And txtTicketNo.Text <> "" Then
        cmdRefreshSomeTkt_Click
    End If
End Sub



Private Sub txtEndTicketNo_Validate(Cancel As Boolean)
    If txtEndTicketNo <> "" Then
        If Val(Right(txtEndTicketNo.Text, TicketNoNumLen())) < Val(Right(txtTicketNo.Text, TicketNoNumLen())) Then
            MsgBox "����Ʊ��Ӧ������ʼƱ�ţ�", vbInformation, "����"
            Cancel = True
        End If
    End If
End Sub

Private Sub txtTicketNo_GotFocus()
    txtTicketNo.SelStart = 0
    txtTicketNo.SelLength = Len(txtTicketNo.Text)
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



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    DealWithChildKeyDown KeyCode, Shift
End Sub



Private Sub txtTicketNo_Change()
    EnableReturnButton
End Sub


Private Sub txtTicketNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim szTemp As String
    Dim lTemp As Long
    On Error GoTo Error_Handel
    If Len(txtTicketNo.Text) >= TicketNoNumLen() Then
        If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
            lTemp = Right(txtTicketNo.Text, TicketNoNumLen())
            szTemp = Left(txtTicketNo.Text, Len(txtTicketNo.Text) - TicketNoNumLen())
            
            If KeyCode = vbKeyUp Then
                lTemp = lTemp + 1
            Else
                lTemp = lTemp - 1
                lTemp = IIf(lTemp > 0, lTemp, 0)
            End If
            txtTicketNo.Text = MakeTicketNo(lTemp, szTemp)
            KeyCode = 0
        End If
    End If
    Exit Sub
Error_Handel:
End Sub
'////////////////////////////
'���ó�Ʊ��Ϣ�б�ͷ
Private Sub FillHeaderColumn()
    With lvTicketInfo.ColumnHeaders
        .Add , , "Ʊ��", 1200
        .Add , , "����", 810
        .Add , , "��վ", 1100
        .Add , , "��վ", 1050
        .Add , , "����ʱ��", 2500
        .Add , , "��λ", 629
        .Add , , "״̬", 1964
        .Add , , "����(%)", 1100
        .Add , , "��ƱƱ��", 0
        .Add , , "Ʊ��", 950
        .Add , , "Ʊ��", 900
        .Add , , "����", 2000
        .Add , , "��Ʊʱ��", 0
        .Add , , "��ƱԱ", 1100
    End With
End Sub


'��ʾ��Ʊ��Ϣ
Private Sub GetReturnTktInfo(TicketID As String)
    Dim liTemp As ListItem
    Dim oTicket As ServiceTicket
    Dim oREBus As REBus
    
    On Error GoTo handle
    Set oTicket = m_oSell.GetServerTicketAtCurrentUnit(TicketID)
    Set liTemp = lvTicketInfo.ListItems.Add(, , TicketID)
    With liTemp
        .SubItems(RT_BusID) = Trim(oTicket.REBusID)
        


        .SubItems(RT_StartStation) = Trim(oTicket.SellStationName)
        .SubItems(RT_EndStation) = Trim(oTicket.ToStationName)
        On Error GoTo 0
        
        On Error Resume Next
        '�����������������۳���Ʊ���Ǳ����ε������
        Set oREBus = m_oSell.CreateServiceObject("STReSch.REBus")
        oREBus.Init m_oAUser
        oREBus.Identify oTicket.REBusID, oTicket.REBusDate
        If oREBus.BusType <> TP_ScrollBus Then
            .SubItems(RT_OffTime) = Format(ToStandardTimeStr(oTicket.dtBusStartUpTime), "hh:mm")
        Else
            .SubItems(RT_OffTime) = MakeDisplayString(cszScrollBus, Format(ToStandardTimeStr(oTicket.dtBusStartUpTime), "hh:mm"))
        End If
        On Error GoTo 0
        If .SubItems(RT_OffTime) = "" Then .SubItems(RT_OffTime) = "Զ�̳�Ʊ"
        On Error GoTo handle
        
'        .SubItems(RT_OffTime) = Trim(oTicket.dtBusStartUpTime)
        .SubItems(RT_SeatNo) = Trim(oTicket.SeatNo)
        .SubItems(RT_Status) = Trim(GetTicketStatusStr(oTicket.TicketStatus))
        
        If (oTicket.TicketStatus And ST_TicketReturned) = 0 Then
            If chkNoRatio.Value = vbChecked Then
                '���ȫ����Ʊ��,��������Ϊ0
                .SubItems(RT_ReturnRatio) = 0
            Else
                .SubItems(RT_ReturnRatio) = Round(oTicket.ReturnRatio, 2)
            End If
        Else
            .SubItems(RT_ReturnRatio) = Round(oTicket.ReturnedInfo.sgReturnCharge / CDbl(oTicket.TicketPrice) * 100, 2)
        End If
        
        .SubItems(RT_ReturnMoney) = FormatMoney(oTicket.TicketPrice - FormatTail(.SubItems(RT_ReturnRatio) * oTicket.TicketPrice / 100))
        .SubItems(RT_TicketType) = Trim(GetTicketTypeStr2(oTicket.TicketType))
        .SubItems(RT_TicketPrice) = FormatMoney(oTicket.TicketPrice)
        .SubItems(RT_Date) = Trim(ToStandardDateStr(oTicket.REBusDate))
        .SubItems(RT_SellTime) = Trim(ToStandardDateTimeStr(oTicket.SellTime))
        .SubItems(RT_Seller) = Trim(oTicket.Operator)
        
        
        
    End With
    Set liTemp = Nothing
    Set oTicket = Nothing
    Exit Sub
handle:
    Set liTemp = Nothing
    Set oTicket = Nothing
End Sub

Private Sub SerialReturnTkt()
'�õ���Ʊ��Ϣ��ʾ�ڳ�Ʊ��ϢListView��
    Dim lTemp1 As Long, lTemp2 As Long, lTemp3 As Long
    Dim szTemp As String
    Dim lCount As Long
    Dim szTicketNo As String
    Dim i As Integer
    
    
    lvTicketInfo.ListItems.Clear
    
    On Error Resume Next
    lTemp1 = Right(txtTicketNo.Text, TicketNoNumLen())
    lTemp3 = Right(txtEndTicketNo.Text, TicketNoNumLen())
    On Error GoTo 0
    On Error GoTo here
'    lTemp2 = lTemp1 + txtCount.Value - 1
    If txtEndTicketNo.Text <> "" Then
        lTemp2 = lTemp3
    Else
        lTemp2 = lTemp1
    End If
    
    If lTemp3 - lTemp1 + 1 <= 100 Then
        If lTemp1 <= lTemp2 Then
            If Len(txtTicketNo.Text) - TicketNoNumLen() > 0 Then szTemp = Left(txtTicketNo.Text, Len(txtTicketNo.Text) - TicketNoNumLen())
            For lCount = lTemp1 To lTemp2
                '���listView���ڴ�Ʊ��Ϣ��ˢ��,�����������ˢ��
                szTicketNo = szTemp & String(TicketNoNumLen() - Len(CStr(lCount)), "0") & lCount
                For i = 1 To lvTicketInfo.ListItems.count
                    If lvTicketInfo.ListItems(i) = szTicketNo Then
                        Exit For
                    End If
                Next i
                If i > lvTicketInfo.ListItems.count Or lvTicketInfo.ListItems.count = 0 Then GetReturnTktInfo szTicketNo
            Next lCount
        End If
        lCount = lvTicketInfo.ListItems.count
        If lCount > 0 Then lvTicketInfo.ListItems(lCount).Selected = True
        ShowReturnInfo
        GetReturnMoney
    Else
        MsgBox "Ϊ��֤ϵͳ����Ч�ʣ���Ʊ����Ӧ��100�����ڣ�", vbInformation, "ע��"
    End If
    
    Exit Sub
here:
    ShowErrorMsg
End Sub

Private Sub txtTicketNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And txtTicketNo.Text <> "" Then
        cmdRefreshSomeTkt_Click
    End If
End Sub



Private Sub GetReturnMoney()
    '�õ��ܵ�Ӧ��Ʊ��
    Dim dbTotalPrice As Double '��Ʊ��
    Dim dbTotalMoney As Double '���˽��
    Dim i As Integer
    
    If lvTicketInfo.ListItems.count = 0 Then Exit Sub
    
    dbTotalPrice = 0
    dbTotalMoney = 0
    For i = 1 To lvTicketInfo.ListItems.count
        With lvTicketInfo.ListItems(i)
            dbTotalPrice = dbTotalPrice + .SubItems(RT_TicketPrice)
            dbTotalMoney = dbTotalMoney + .SubItems(RT_ReturnMoney)
        End With
    Next i
    lblTotalPrice.Caption = FormatMoney(dbTotalPrice)
    lblTotalFees = FormatMoney(dbTotalPrice - dbTotalMoney)
    lblTotalReturnMoney = FormatMoney(dbTotalMoney)
End Sub


'��ʼ���õ���ѡ�����Ʊ����
Private Sub InitReturnRatioValue()
    Dim arfValue() As RETURNFEES
    Dim iLen As Integer
    Dim i As Integer
    On Error GoTo here
    Dim szTemp As String

    szTemp = m_oSell.SellUnitCode
    m_oSell.SellUnitCode = m_szCurrentUnitID
    
    arfValue = m_oSell.GetReturnRatioValue
    m_oSell.SellUnitCode = szTemp
    
    iLen = ArrayLength(arfValue)
    ReDim m_asgLeastMoney(1 To iLen)
    ReDim m_asgReturnRatio(1 To iLen)
    If iLen <> 0 Then
        For i = 1 To iLen
            cboFeesRatio.AddItem Round(arfValue(i).sgReturnRate, 2)
            m_asgLeastMoney(i) = arfValue(i).sgLeastMoney
            m_asgReturnRatio(i) = arfValue(i).sgReturnRate
        Next i
        For i = 1 To iLen
           If arfValue(i).sgReturnRate = 0 Then Exit For
        Next i
        If i > iLen Then cboFeesRatio.AddItem "0"
        
        For i = 1 To iLen
            If arfValue(i).sgReturnRate = 100 Then Exit For
        Next i
        If i > iLen Then cboFeesRatio.AddItem "100"
        
    End If
    On Error GoTo 0
    Exit Sub
here:
    ShowErrorMsg
End Sub

'��ʾ�ʵ�����Ʊ����
Private Sub ShowReturnRatio()
    Dim i As Integer
    Dim dbReturnRatio As Double
    If lvTicketInfo.SelectedItem Is Nothing Then Exit Sub
    
    On Error GoTo here
    dbReturnRatio = lvTicketInfo.SelectedItem.SubItems(RT_ReturnRatio)
    For i = 0 To cboFeesRatio.ListCount - 2
        If Abs(cboFeesRatio.List(i) - Round(dbReturnRatio, 2)) < 2 Then
            cboFeesRatio.ListIndex = i
            Exit Sub
        End If
    Next i
    '���δ�ҵ�,���趨Ϊȫ����Ʊ
    cboFeesRatio.ListIndex = cboFeesRatio.ListCount - 1
    Exit Sub
here:
    MsgBox "��Ʊ����δ���ã�", vbInformation, "����"
End Sub

'�жϵ�ǰ�û��з�ѡ����Ʊ���ʵ�Ȩ��
Private Sub SelectReturnRatioValid()
    On Error GoTo here
    If m_oSell.SelectReturnIsValid Then
        cboFeesRatio.Enabled = False
    Else
        cboFeesRatio.Enabled = True
    End If
    On Error GoTo 0
    Exit Sub
here:
    ShowErrorMsg
End Sub

Private Sub SetDefaultLabel()
'�ָ���ʼֵ
    lblStartStation.Caption = ""
    
    
    lblBusID.Caption = ""
    lblDate.Caption = ""
    lblStatus.Caption = ""
    lblEndStation.Caption = ""
    lblTicketType.Caption = ""
    lblTicketPrice.Caption = ""
    lblSeller.Caption = ""
    lblSeatNo.Caption = ""
    lblSellTime.Caption = ""
    lblOffTime.Caption = ""
    lblTicketID.Caption = ""
    lblVehicleType.Caption = ""
    
    lblCurrectTicketPrice.Caption = "0.00"
    lblReturnCharge.Caption = "0.00"
    lblTotalPrice.Caption = "0.00"
    lblTotalFees.Caption = "0.00"
    lblTotalReturnMoney.Caption = "0.00"
    
    
    
End Sub

Private Function GetAllTickets() As String()
    '����txtTicketNo ��txtCount �õ����е�Ʊ
    Dim lTemp1 As Long
    Dim lTemp2 As Long
    Dim lTemp3 As Long
    Dim szTemp As String
    Dim lCount As Long
    Dim aszTemp() As String
    If lvTicketInfo.ListItems.count > 0 Then
        ReDim aszTemp(1 To lvTicketInfo.ListItems.count)
    Else
        Exit Function
    End If
    For lTemp1 = 1 To lvTicketInfo.ListItems.count
        aszTemp(lTemp1) = lvTicketInfo.ListItems(lTemp1).Text
    Next
    
    GetAllTickets = aszTemp
    
End Function

Private Function FormatTail(pdbValue As Double) As Double
    '����β������
    '0-2��0,3-7��5,8-9��10
    Dim dbTemp As Double
    dbTemp = pdbValue - Int(pdbValue)
    If dbTemp >= 0 And dbTemp < 0.3 Then
        '0-2��0
        FormatTail = Int(pdbValue)
    ElseIf dbTemp >= 0.3 And dbTemp < 0.8 Then
        '3-7��5
        FormatTail = FormatMoney(Int(pdbValue) + 0.5)
    Else
        '8-9��10
        FormatTail = FormatMoney(Int(pdbValue) + 1)
    End If
    
    
End Function

