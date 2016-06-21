VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{A0123751-4698-48C1-A06C-A2482B5ED508}#2.0#0"; "RTComctl2.ocx"
Object = "{61C3A787-42A5-4F09-9AD8-C9DE75BAD364}#1.0#0"; "STSeatpad.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmCheckTicket 
   BackColor       =   &H00808080&
   ClientHeight    =   7470
   ClientLeft      =   1710
   ClientTop       =   1995
   ClientWidth     =   9630
   FillColor       =   &H00808080&
   HelpContextID   =   4000201
   Icon            =   "frmCheckTicket.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7470
   ScaleWidth      =   9630
   WindowState     =   2  'Maximized
   Begin VB.Frame fraOutLine 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6510
      Left            =   270
      TabIndex        =   5
      Top             =   660
      Width           =   9075
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6255
         Left            =   180
         TabIndex        =   6
         Top             =   210
         Width           =   8730
         Begin VB.Frame fraTicketInfo 
            BackColor       =   &H00E0E0E0&
            Caption         =   "车票信息"
            Height          =   1080
            Left            =   2550
            TabIndex        =   56
            Top             =   1590
            Width           =   6120
            Begin VB.Label lblTicketID2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "票号:"
               Height          =   180
               Left            =   120
               TabIndex        =   72
               Top             =   240
               Width           =   450
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "到达站:"
               Height          =   180
               Left            =   1860
               TabIndex        =   71
               Top             =   240
               Width           =   630
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "座位号:"
               Height          =   180
               Left            =   120
               TabIndex        =   70
               Top             =   510
               Width           =   630
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "票状态:"
               Height          =   180
               Left            =   1860
               TabIndex        =   69
               Top             =   510
               Width           =   630
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "票种:"
               Height          =   180
               Left            =   120
               TabIndex        =   68
               Top             =   780
               Width           =   450
            End
            Begin VB.Label lblTicketID 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "43194319"
               Height          =   180
               Left            =   600
               TabIndex        =   67
               Top             =   240
               Width           =   720
            End
            Begin VB.Label lblSeatNo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "12"
               Height          =   180
               Left            =   750
               TabIndex        =   66
               Top             =   510
               Width           =   180
            End
            Begin VB.Label lblTicketType 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "全票/半票/免票"
               Height          =   180
               Left            =   600
               TabIndex        =   65
               Top             =   780
               Width           =   1260
            End
            Begin VB.Label lblEndStation2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "重庆"
               Height          =   180
               Left            =   2505
               TabIndex        =   64
               Top             =   240
               Width           =   360
            End
            Begin VB.Label lblTicketStatus 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "正常/改签/作废/.."
               Height          =   180
               Left            =   2505
               TabIndex        =   63
               Top             =   510
               Width           =   1530
            End
            Begin VB.Label lblPersonName 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "张三李四王五"
               Height          =   180
               Left            =   4440
               TabIndex        =   62
               Top             =   510
               Width           =   1080
            End
            Begin VB.Label Label30 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "姓名:"
               Height          =   180
               Left            =   3960
               TabIndex        =   61
               Top             =   510
               Width           =   450
            End
            Begin VB.Label lblCardType 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "身份证"
               Height          =   180
               Left            =   4440
               TabIndex        =   60
               Top             =   240
               Width           =   540
            End
            Begin VB.Label Label29 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "证件类型:"
               Height          =   180
               Left            =   3600
               TabIndex        =   59
               Top             =   240
               Width           =   810
            End
            Begin VB.Label Label27 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "证件号:"
               Height          =   180
               Left            =   3780
               TabIndex        =   58
               Top             =   780
               Width           =   630
            End
            Begin VB.Label lblIDCardNo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "3306231983****2253"
               Height          =   180
               Left            =   4440
               TabIndex        =   57
               Top             =   780
               Width           =   1620
            End
         End
         Begin RTComctl3.CoolButton cmdDetailInfo 
            Height          =   285
            Left            =   7590
            TabIndex        =   55
            Top             =   1380
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   503
            BTYPE           =   3
            TX              =   "详细信息(&D)"
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
            MICON           =   "frmCheckTicket.frx":014A
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.CommandButton cmdLedShow 
            Caption         =   "插播条屏(&I)"
            Height          =   435
            Left            =   7470
            TabIndex        =   53
            Top             =   3990
            Width           =   1155
         End
         Begin RTComctl3.CoolButton cmdFind 
            Height          =   255
            Left            =   1890
            TabIndex        =   52
            Top             =   1440
            Width           =   465
            _ExtentX        =   820
            _ExtentY        =   450
            BTYPE           =   3
            TX              =   "查询"
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
            MICON           =   "frmCheckTicket.frx":0166
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.CheckBox chkCheckChange 
            BackColor       =   &H00E0E0E0&
            Height          =   645
            HelpContextID   =   4000211
            Left            =   2670
            Picture         =   "frmCheckTicket.frx":0182
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "点击切换至改乘检入模式"
            Top             =   510
            Width           =   795
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00E0E0E0&
            Height          =   1005
            Left            =   2550
            TabIndex        =   31
            Top             =   2670
            Width           =   6135
            Begin RTComctl3.CoolButton cmdRefreshSeat 
               Height          =   345
               Left            =   4530
               TabIndex        =   3
               Top             =   570
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   609
               BTYPE           =   3
               TX              =   " 刷新(&R)"
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
               MICON           =   "frmCheckTicket.frx":06EF
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "并入车次售票数:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   300
               Left            =   2730
               TabIndex        =   49
               Top             =   240
               Width           =   1740
            End
            Begin VB.Label lblMergeInSells 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "23"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   300
               Left            =   4530
               TabIndex        =   48
               Top             =   240
               Width           =   270
            End
            Begin VB.Label lblUncheckSum 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "23"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   300
               Left            =   2250
               TabIndex        =   39
               Top             =   600
               Width           =   270
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "未检数:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   300
               Left            =   1410
               TabIndex        =   45
               Top             =   600
               Width           =   780
            End
            Begin VB.Label lblCheckedSum 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "23"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   300
               Left            =   960
               TabIndex        =   44
               Top             =   600
               Width           =   270
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "已检数:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   300
               Left            =   120
               TabIndex        =   43
               Top             =   600
               Width           =   780
            End
            Begin VB.Label lblChangeSum 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "23"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   300
               Left            =   3840
               TabIndex        =   37
               Top             =   600
               Width           =   270
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "改乘票数:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   300
               Left            =   2730
               TabIndex        =   36
               Top             =   600
               Width           =   1020
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "总座数:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   300
               Left            =   120
               TabIndex        =   35
               Top             =   240
               Width           =   780
            End
            Begin VB.Label lblSeatSum 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "32"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   300
               Left            =   960
               TabIndex        =   34
               Top             =   240
               Width           =   270
            End
            Begin VB.Label lblTicketSells 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "23"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   300
               Left            =   2250
               TabIndex        =   33
               Top             =   240
               Width           =   270
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "已售数:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   300
               Left            =   1410
               TabIndex        =   32
               Top             =   240
               Width           =   780
            End
         End
         Begin MSComctlLib.ImageList ImageList1 
            Left            =   7440
            Top             =   3600
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   8
            ImageHeight     =   8
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   2
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmCheckTicket.frx":070B
                  Key             =   "CheckSeat"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmCheckTicket.frx":07CB
                  Key             =   "NoCheckSeat"
               EndProperty
            EndProperty
         End
         Begin STSeatPad.SeatPad SeatPad1 
            Height          =   2115
            Left            =   2550
            TabIndex        =   23
            Top             =   3990
            Width           =   4755
            _ExtentX        =   8387
            _ExtentY        =   3731
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   14737632
            Caption         =   "SeatPad1"
            GridNum         =   20
            RowGrids        =   10
            SeatPadStyle    =   1
         End
         Begin VB.CommandButton cmdStopCheck 
            BackColor       =   &H00E0E0E0&
            Caption         =   "停止检票 F9"
            Height          =   960
            Left            =   7470
            Picture         =   "frmCheckTicket.frx":087F
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   5130
            Width           =   1155
         End
         Begin VB.TextBox txtTicketID 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   465
            IMEMode         =   2  'OFF
            Left            =   3555
            TabIndex        =   1
            Text            =   "1234567"
            Top             =   600
            Width           =   3300
         End
         Begin VB.Frame fraBusInfo 
            BackColor       =   &H00E0E0E0&
            Caption         =   "车次信息"
            Height          =   6120
            Left            =   0
            TabIndex        =   7
            Top             =   0
            Width           =   2400
            Begin VB.CheckBox chkExtraCheck 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               ForeColor       =   &H80000008&
               Height          =   180
               Left            =   1005
               TabIndex        =   46
               Top             =   4740
               Width           =   225
            End
            Begin RTComctl2.RevTimer rvtLostTime 
               Height          =   330
               Left            =   1020
               Top             =   5580
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   582
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BackColor       =   16777215
               OutnerStyle     =   2
               Enabled         =   0   'False
               Second          =   0
               HourNoShowIfZero=   -1  'True
            End
            Begin RTComctl3.CoolButton cmdAllotStationTicketsInfo 
               Height          =   345
               Left            =   240
               TabIndex        =   54
               Top             =   3000
               Width           =   1965
               _ExtentX        =   3466
               _ExtentY        =   609
               BTYPE           =   3
               TX              =   "配载站售/检票信息(&P)"
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
               MICON           =   "frmCheckTicket.frx":0FB0
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "补检状态:"
               Height          =   180
               Left            =   150
               TabIndex        =   47
               Top             =   4740
               Width           =   810
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00808080&
               Index           =   1
               X1              =   30
               X2              =   2370
               Y1              =   3375
               Y2              =   3375
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00FFFFFF&
               Index           =   0
               X1              =   30
               X2              =   2370
               Y1              =   3390
               Y2              =   3390
            End
            Begin VB.Label Label28 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "途经:"
               Height          =   180
               Left            =   150
               TabIndex        =   42
               Top             =   2010
               Width           =   450
            End
            Begin VB.Label lblStationInRoad 
               BackStyle       =   0  'Transparent
               Caption         =   "绍兴、杭州"
               ForeColor       =   &H00FF0000&
               Height          =   720
               Left            =   150
               TabIndex        =   41
               Top             =   2280
               Width           =   2130
               WordWrap        =   -1  'True
            End
            Begin VB.Label lblEndStation 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "杭州"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   330
               Left            =   960
               TabIndex        =   40
               Top             =   660
               Width           =   1305
            End
            Begin VB.Label lblBusID 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "1231-1"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   330
               Left            =   960
               TabIndex        =   38
               Top             =   270
               Width           =   1305
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "开检时间:"
               Height          =   180
               Left            =   150
               TabIndex        =   30
               Top             =   5040
               Width           =   810
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "停检时间:"
               Height          =   180
               Left            =   150
               TabIndex        =   29
               Top             =   5340
               Width           =   810
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "剩余时间:"
               Height          =   180
               Left            =   150
               TabIndex        =   28
               Top             =   5670
               Width           =   810
            End
            Begin VB.Label lblBeginCheckTime 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "00:00:00"
               Height          =   180
               Left            =   1020
               TabIndex        =   27
               Top             =   5040
               Width           =   720
            End
            Begin VB.Label lblEndCheckTime 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "00:00:00"
               Height          =   180
               Left            =   1020
               TabIndex        =   26
               Top             =   5340
               Width           =   720
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "并入车次:"
               Height          =   180
               Left            =   150
               TabIndex        =   25
               Top             =   1740
               Width           =   810
            End
            Begin VB.Label lblMergeIn 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "1234"
               ForeColor       =   &H00FF0000&
               Height          =   180
               Left            =   1020
               TabIndex        =   24
               Top             =   1740
               Width           =   360
            End
            Begin VB.Label lblVehicleType 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "豪华大巴"
               Height          =   180
               Left            =   1020
               TabIndex        =   21
               Top             =   3840
               Width           =   720
            End
            Begin VB.Label lblOwner 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "某某某"
               Height          =   180
               Left            =   1020
               TabIndex        =   20
               Top             =   4440
               Width           =   540
            End
            Begin VB.Label lblVehicle 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "川A13884"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   210
               Left            =   1020
               TabIndex        =   19
               Top             =   1470
               Width           =   840
            End
            Begin VB.Label lblCompany 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "客运"
               Height          =   180
               Left            =   1020
               TabIndex        =   18
               Top             =   4140
               Width           =   360
            End
            Begin VB.Label lblBusType 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "固定/流水"
               Height          =   180
               Left            =   1020
               TabIndex        =   17
               Top             =   3540
               Width           =   810
            End
            Begin VB.Label lblStartupTime 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "10:10"
               BeginProperty Font 
                  Name            =   "黑体"
                  Size            =   15
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   330
               Left            =   960
               TabIndex        =   16
               Top             =   1050
               Width           =   1305
            End
            Begin VB.Label Label21 
               BackStyle       =   0  'Transparent
               Caption         =   "开往"
               Height          =   255
               Left            =   150
               TabIndex        =   15
               Top             =   735
               Width           =   495
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "车次类型:"
               Height          =   180
               Left            =   150
               TabIndex        =   14
               Top             =   3540
               Width           =   810
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "车型:"
               Height          =   180
               Left            =   150
               TabIndex        =   13
               Top             =   3840
               Width           =   450
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "车主:"
               Height          =   180
               Left            =   150
               TabIndex        =   12
               Top             =   4440
               Width           =   450
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "车辆(F2):"
               Height          =   180
               Left            =   150
               TabIndex        =   11
               Top             =   1470
               Width           =   810
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "参营公司:"
               Height          =   180
               Left            =   150
               TabIndex        =   10
               Top             =   4140
               Width           =   810
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "发车时间"
               Height          =   180
               Left            =   150
               TabIndex        =   9
               Top             =   1140
               Width           =   720
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "检票车次"
               Height          =   180
               Left            =   150
               TabIndex        =   8
               Top             =   315
               Width           =   720
            End
         End
         Begin RTComctl3.FlatLabel lblCheckInfo 
            Height          =   315
            Left            =   2550
            TabIndex        =   50
            Top             =   1230
            Width           =   5010
            _ExtentX        =   8837
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
            BackColor       =   14737632
            OutnerStyle     =   2
            Caption         =   ""
         End
         Begin RTComctl3.FlatLabel FlatLabel1 
            Height          =   645
            Left            =   3480
            TabIndex        =   51
            Top             =   510
            Width           =   3465
            _ExtentX        =   6112
            _ExtentY        =   1138
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   14737632
            InnerStyle      =   2
            Caption         =   ""
         End
         Begin VB.Image imgEnabled 
            Height          =   405
            Left            =   7830
            Picture         =   "frmCheckTicket.frx":0FCC
            Stretch         =   -1  'True
            Top             =   645
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "请输入车票号码(&T):"
            Height          =   240
            Left            =   2550
            TabIndex        =   0
            Top             =   180
            Width           =   3450
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "座位检票情况:"
            Height          =   180
            Left            =   2640
            TabIndex        =   22
            Top             =   3750
            Width           =   1170
         End
         Begin VB.Image Image1 
            Height          =   720
            Left            =   7365
            Picture         =   "frmCheckTicket.frx":1896
            Top             =   240
            Width           =   720
         End
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00004000&
         BorderWidth     =   3
         Height          =   6495
         Left            =   0
         Top             =   0
         Width           =   9075
      End
   End
End
Attribute VB_Name = "frmCheckTicket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'***************************************************************************************
'    注释:
'    g_tCheckInfo的车次信息和g_oEnvBus的信息在初始化窗体时才能使用 ,
'    将检票需用的信息在初始时传入窗体变量
'    初始化过程有：
'       Form_Load    主要在窗体初始化时获取检票过程中需用的信息传入窗体变量
'       ShowInitInfo   主要初始化界面，信息主要从g_tCheckInfo和g_oEnvBus获取
'       InitSeatPadInfo    主要初始化座位面板，信息主要从g_tCheckInfo和g_oEnvBus获取
'***************************************************************************************
Const cszTitleInfo = "检票信息"
Const cszNeedMakeSheet = "是否生成路单？"
'以下窗体变量存放检票过程中需用到的参数
Private mszBusID As String                             '车次ID
Private mnBusSerialNo As Integer                       '车次序号
Private mbExCheck As Boolean                          '是否是补检
'控制变量
Private m_bStopSuccess As Boolean                       '是否停检成功
Private m_nCheckedTickets   As Integer                                '已检票数
Private m_nOldSoldCount As Integer
Private m_nNewSoldCount As Integer
Private m_nOldSelftStationSoldCount As Integer       '本上车站的票数，用于统计已检，未检数
Private m_nNewSelfStationSoldCount As Integer

Private m_nOtherStationCheckCount As Integer     '别的站检本上车站的票数
Private m_nCheckOtherStationCount As Integer     '本站检别的上车站的票数
Private m_nOldOtherStationCheckCount As Integer     '别的站检本上车站的票数
Private m_nOldCheckOtherStationCount As Integer     '本站检别的上车站的票数


Private m_MergeBusInfo() As String
Private m_nSellTickets As Integer '售出票数

Private m_MergeBusCheckInfo() As String
Private m_nMergeInChecked As Integer

Private m_nSelfSellStationTickets As Integer '本上车站的票数
Private m_nOtherTickets As Integer '改乘票数


Private mnBusMode As EBusType                           '车次类型
Private TTicketInfo As TInterfaceCheckTicketEx


Public m_oREBus As REBus

Private Sub CheckTicket(ByRef bCheckSuccess As Boolean)
        '车票验证
        '如果有效检入车票
    Static nSeatNo As Integer
    
    Dim n As Integer
    Dim szTemp As String
    Dim szTicketBusid As String     '车票所属的车次号
    Dim nStatus As Integer
On Error GoTo ErrHandle
    '取得当前车票信息
    TTicketInfo = g_oChkTicket.GetOneTicketInfoEx(txtTicketID.Text)
    TTicketInfo.BusID = UCase(Trim(TTicketInfo.BusID))
    
    
    WriteTicketInfo TTicketInfo, bCheckSuccess
    
    If Not bCheckSuccess Then              '车票无效
        Exit Sub
    End If
    
    szTicketBusid = Trim(TTicketInfo.BusID)
    If mnBusMode = TP_ScrollBus Then
        If szTicketBusid <> mszBusID Then
            lblCheckInfo.Caption = "本车票不属于当检车次"
            PlayEventSound g_tEventSoundPath.NoMatchedBus
            bCheckSuccess = False
            Exit Sub
        End If
        If m_nCheckedTickets + 1 > SeatPad1.GridNum Then
            bCheckSuccess = False
            MsgBox "座位数已满,不允许再检入", vbExclamation, Me.Caption
            Exit Sub
        End If
    Else
        If m_nCheckedTickets + 1 > SeatPad1.GridNum Then
            bCheckSuccess = False
            MsgBox "座位数已满,不允许再检入", vbExclamation, Me.Caption
            Exit Sub
        End If
        
        '以下判断是否为车票是否属于当前车次，或属于并班车次，或是改乘检入
        If szTicketBusid <> mszBusID And IsMergeBus(szTicketBusid, lblMergeIn.Caption, txtTicketID.Text) = False Then
            If chkCheckChange.Value = vbChecked Then
                    m_nOtherTickets = m_nOtherTickets + 1
'                If MsgboxEx("是否将该车票进行改乘处理？", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
'                    m_nOtherTickets = m_nOtherTickets + 1
'                Else
'                    lblCheckInfo.Caption = "本车票不属于当检车次"
'                    PlayEventSound g_tEventSoundPath.NoMatchedBus
'                    bCheckSuccess = False
'                    Exit Sub
'                End If
            Else
                lblCheckInfo.Caption = "本车票不属于当检车次"
                PlayEventSound g_tEventSoundPath.NoMatchedBus
                bCheckSuccess = False
                Exit Sub
            End If
        End If
    End If
    
    '玉环站和坎门站不允许检入楚门的车票
    If Trim(TTicketInfo.SellStationID) = "cm" And (Trim(g_oActiveUser.SellStationID) = "yh" Or Trim(g_oActiveUser.SellStationID) = "km") Then
        bCheckSuccess = False
        MsgBox "楚门的车票不允许检入！", vbExclamation, Me.Caption
        Exit Sub
    End If
    
    
    
   If Trim(TTicketInfo.SellStationID) <> "cm" And (Trim(g_oActiveUser.SellStationID) = "cm") Then
        bCheckSuccess = False
        MsgBox "非楚门站上车的车票不允许检入！", vbExclamation, Me.Caption
        Exit Sub
    End If
    
    '检入此票
    g_oChkTicket.CheckTicket mszBusID, mnBusSerialNo, txtTicketID.Text, TTicketInfo.BusDate
    If mnBusMode <> TP_ScrollBus Then
        '如果为固定班次
        If szTicketBusid <> mszBusID And IsMergeBus(szTicketBusid, lblMergeIn.Caption, txtTicketID.Text) = False Then
            '车次与当前检票车次不同
            If chkCheckChange.Value = vbChecked Then
                '为改乘
                '每检一张票,把改乘设回去
                chkCheckChange.Value = vbUnchecked
                chkCheckChange_Click
            End If
        End If
    End If
    
    
    '刷新检票进度
    m_nCheckedTickets = m_nCheckedTickets + 1
   
    If szTicketBusid = mszBusID Then             '是本车次的票，不是并入票
        '刷新进度条和座位板
        If mnBusMode = TP_ScrollBus Then
            If m_nCheckedTickets > SeatPad1.GridNum Then

                g_tCheckInfo.SeatCount = m_nCheckedTickets
                SeatPad1.GridNum = m_nCheckedTickets
            End If
            SeatPad1.PadGrids.Item(m_nCheckedTickets).BackColor = &HC0C0E0
            SeatPad1.PadGrids.Item(m_nCheckedTickets).Enabled = True
            SeatPad1.PadGrids.Item(m_nCheckedTickets).MiniIcon = ImageList1.ListImages("CheckSeat").Picture
            SeatPad1.RefreshGrid (m_nCheckedTickets)
        Else
            If TTicketInfo.SeatNo <> 0 Then
                n = 1
                szTemp = Trim(Str(TTicketInfo.SeatNo))
                If Len(szTemp) < 2 Then szTemp = "0" & szTemp
                Do
                   If SeatPad1.PadGrids.Item(n).Caption = szTemp Then
                       SeatPad1.PadGrids.Item(n).BackColor = &HC0C0E0
                       SeatPad1.PadGrids.Item(n).Enabled = True
                       SeatPad1.PadGrids.Item(n).MiniIcon = ImageList1.ListImages("CheckSeat").Picture
                       SeatPad1.RefreshGrid n
                       Exit Do
                   End If
                   n = n + 1
                Loop Until n > SeatPad1.GridNum
            End If
        End If
    Else                                               '是并入票
           If mnBusMode = TP_ScrollBus Then
                If m_nCheckedTickets > SeatPad1.GridNum Then
                    g_tCheckInfo.SeatCount = m_nCheckedTickets
                    SeatPad1.GridNum = m_nCheckedTickets
                End If
                SeatPad1.PadGrids.Item(m_nCheckedTickets).BackColor = &HC0C0E0
                SeatPad1.PadGrids.Item(m_nCheckedTickets).Enabled = True
                SeatPad1.PadGrids.Item(m_nCheckedTickets).MiniIcon = ImageList1.ListImages("CheckSeat").Picture
                SeatPad1.RefreshGrid (m_nCheckedTickets)
           Else
               If g_tCheckInfo.SplitSeat <> 0 Then
                    n = 1
                    szTemp = Trim(Str(g_tCheckInfo.SplitSeat))  'g_tCheckInfo.SplitSeat
                    If Len(szTemp) < 2 Then szTemp = "0" & szTemp
                    Do
                       If SeatPad1.PadGrids.Item(n).Caption = szTemp Then
                           SeatPad1.PadGrids.Item(n).BackColor = &HC0C0E0
                           SeatPad1.PadGrids.Item(n).Enabled = True
                           SeatPad1.PadGrids.Item(n).MiniIcon = ImageList1.ListImages("CheckSeat").Picture
                           SeatPad1.RefreshGrid n
                           Exit Do
                       End If
                       n = n + 1
                    Loop Until n > SeatPad1.GridNum
                 End If
           End If
    End If
    RefreshCheckCountInfo
    lblTicketStatus.Caption = "检票成功"
    bCheckSuccess = True
    Exit Sub
ErrHandle:
    If err.Number = ERR_TicketNoExist Or err.Number = ERR_ChkTkTicketStatusError Or err.Number = ERR_ChkTkTicketNotChecked Then
        lblCheckInfo.Caption = "无效车票，检入失败"
        PlayEventSound g_tEventSoundPath.InvalidTicket
    Else
        ShowErrorMsg
    End If
    bCheckSuccess = False
End Sub

Private Sub StopCheckTicket()
On Error GoTo ErrHandle
    g_oChkTicket.StopCheckTicket mszBusID
    
    m_bStopSuccess = True
    Exit Sub
ErrHandle:
    ShowErrorMsg
    m_bStopSuccess = False
End Sub
Private Sub InitSeatPadInfo()
    '初始化座位信息及检票信息
    Dim i As Integer
    Dim aSeatInfo() As TSeatInfoEx
    Dim aSeatCheckInfo() As TSeatStatus
    Dim aSelfStationSeatInfo() As TSeatInfoEx      '本站上车的票
    Dim aOtherStationSeatInfo() As TSeatInfoEx   '其它站上车的票
    Dim nSeatCount As Integer
    Dim oVehicle As Object
    On Error GoTo ErrorHandle
    If g_tCheckInfo.BusMode = TP_ScrollBus Then
        Set oVehicle = CreateObject("STBase.Vehicle")
        oVehicle.Init g_oActiveUser
        oVehicle.Identify g_tCheckInfo.RunVehicle.VehicleId
        
        nSeatCount = oVehicle.SeatCount
        g_tCheckInfo.SeatCount = nSeatCount
        '刷新检票人数
        m_nCheckedTickets = ArrayLength(g_oChkTicket.GetBusCheckTicket(Date, mszBusID, mnBusSerialNo, g_tCheckInfo.CheckGateNo))
        
        g_tCheckInfo.SellTickets = nSeatCount
        m_nSellTickets = nSeatCount
        SeatPad1.GridNum = 0
        SeatPad1.GridNum = nSeatCount
        For i = 1 To nSeatCount
            If i <= m_nCheckedTickets Then
                '设置已检
                SeatPad1.PadGrids.Item(i).BackColor = &HC0C0E0
                SeatPad1.PadGrids.Item(i).MiniIcon = _
                ImageList1.ListImages("CheckSeat").Picture
                
            Else
                SeatPad1.PadGrids.Item(i).BackColor = &HFFFFFF
                SeatPad1.PadGrids.Item(i).Caption = Trim(Str(i))
                
                SeatPad1.PadGrids.Item(i).Enabled = False
            End If
        Next i
        Set oVehicle = Nothing
    Else
        aSeatCheckInfo = g_oChkTicket.GetBusSeatCheckInfo(mszBusID, , g_tCheckInfo.CheckGateNo)
        g_oEnvBus.Identify mszBusID, Date, g_tCheckInfo.CheckGateNo
        aSeatInfo = g_oEnvBus.GetSeatInfo()
        aSelfStationSeatInfo = g_oEnvBus.GetOtherSeatInfo(True, g_oActiveUser.SellStationID)
        aOtherStationSeatInfo = g_oEnvBus.GetOtherSeatInfo(False, g_oActiveUser.SellStationID)
        nSeatCount = ArrayLength(aSeatInfo)
        SeatPad1.GridNum = nSeatCount
        
        m_nNewSoldCount = 0
        m_nNewSelfStationSoldCount = 0
        m_nCheckOtherStationCount = 0
        m_nOtherStationCheckCount = 0
        
        
        For i = 1 To nSeatCount
            '根据座位的检票状态和售票状态刷新座位板
           Select Case aSeatInfo(i).szSeatStatus
                Case ST_SeatSold, 4, 5                   '此座位已被售出
                    If ResolveDisplayEx(aSeatInfo(i).szTicketNo) = "已检" Then
                        '已检   本站检的
                        If aSeatCheckInfo(i).szTicketID <> "" Then
                            '淡紫色
                            SeatPad1.PadGrids.Item(i).BackColor = &HC0C0E0
                            SeatPad1.PadGrids.Item(i).MiniIcon = _
                                ImageList1.ListImages("CheckSeat").Picture
                                
                            '判断是否本站检本上车站的票
                            If UCase(aSeatCheckInfo(i).szTicketID) = UCase(aOtherStationSeatInfo(i).szTicketNo) Then
                                m_nCheckOtherStationCount = m_nCheckOtherStationCount + 1
                            End If
                            
                        Else  '其它站检的
                            '灰黑色
                            SeatPad1.PadGrids.Item(i).BackColor = &H808080
                            SeatPad1.PadGrids.Item(i).MiniIcon = _
                                ImageList1.ListImages("CheckSeat").Picture
                                
                            If UCase(ResolveDisplay(aSeatInfo(i).szTicketNo)) = UCase(aSelfStationSeatInfo(i).szTicketNo) Then
                                m_nOtherStationCheckCount = m_nOtherStationCheckCount + 1
                            End If
                            
                        End If
                    Else
                        '未检  本站上车的
                        If UCase(ResolveDisplay(aSeatInfo(i).szTicketNo)) = UCase(ResolveDisplay(aSelfStationSeatInfo(i).szTicketNo)) Then
                            '灰白色
                            SeatPad1.PadGrids.Item(i).BackColor = &HFFFFFF
                            SeatPad1.PadGrids.Item(i).MiniIcon = _
                                ImageList1.ListImages("NoCheckSeat").Picture
                        Else   '不是本站上车的
                            '蓝绿色
                            SeatPad1.PadGrids.Item(i).BackColor = &HFFFF80
                            SeatPad1.PadGrids.Item(i).MiniIcon = _
                                ImageList1.ListImages("NoCheckSeat").Picture
                        End If
                    End If
                    
                    m_nNewSoldCount = m_nNewSoldCount + 1
                    If aSelfStationSeatInfo(i).szTicketNo <> "" Then m_nNewSelfStationSoldCount = m_nNewSelfStationSoldCount + 1
                Case ST_SeatReserved '预留
                
                    If InStr(1, aSeatInfo(i).szRemark, "网购") > 0 Then
                        '网购预留 橘色
                        SeatPad1.PadGrids.Item(i).Enabled = False
                        SeatPad1.PadGrids.Item(i).BackColor = &H80FF&
                        SeatPad1.PadGrids.Item(i).MiniIcon = Nothing
                    Else
                        '普通预留 黄色
                        SeatPad1.PadGrids.Item(i).Enabled = False
                        SeatPad1.PadGrids.Item(i).BackColor = &HFFFF&
                        SeatPad1.PadGrids.Item(i).MiniIcon = Nothing
                    End If
                Case Else
                    '未售出
                    SeatPad1.PadGrids.Item(i).Enabled = False
                    SeatPad1.PadGrids.Item(i).BackColor = &HE0E0E0
                    SeatPad1.PadGrids.Item(i).MiniIcon = Nothing
                 
            End Select
            SeatPad1.PadGrids.Item(i).Caption = Trim(aSeatInfo(i).szSeatNo)
        Next i
        
        If m_nOldSoldCount >= 0 Then
            m_nSellTickets = m_nSellTickets + m_nNewSoldCount - m_nOldSoldCount
        End If
        
        If m_nOldSelftStationSoldCount >= 0 Then
            m_nSelfSellStationTickets = m_nSelfSellStationTickets + m_nNewSelfStationSoldCount - m_nOldSelftStationSoldCount
        End If
        
        If m_nOldCheckOtherStationCount >= 0 Or m_nOldOtherStationCheckCount >= 0 Then
            m_nSelfSellStationTickets = m_nSelfSellStationTickets + m_nCheckOtherStationCount - m_nOtherStationCheckCount - m_nOldCheckOtherStationCount + m_nOldOtherStationCheckCount
        End If
        
        m_nOldCheckOtherStationCount = m_nCheckOtherStationCount
        m_nOldOtherStationCheckCount = m_nOtherStationCheckCount
        m_nOldSoldCount = m_nNewSoldCount
        m_nOldSelftStationSoldCount = m_nNewSelfStationSoldCount
    End If
    
    '显示在界面上
    SeatPad1.Refresh
    Exit Sub
ErrorHandle:
    ShowErrorMsg
    
End Sub
'刷新检票汇总信息
Private Sub RefreshCheckCountInfo()
    If g_tCheckInfo.BusMode <> TP_ScrollBus Then
        lblSeatSum.Caption = g_tCheckInfo.SeatCount
        lblTicketSells.Caption = m_nSellTickets
        
        lblCheckedSum.Caption = m_nCheckedTickets
        Dim nUnChecks As Integer
        nUnChecks = m_nSelfSellStationTickets - m_nCheckedTickets + m_nOtherTickets + g_tCheckInfo.MergeInSells - m_nMergeInChecked
        If nUnChecks > Val(lblTicketSells.Caption) - Val(lblCheckedSum.Caption) Then
            lblUncheckSum.Caption = Val(lblTicketSells.Caption) - Val(lblCheckedSum.Caption)
        Else
            lblUncheckSum.Caption = nUnChecks
        End If
    Else
        lblSeatSum.Caption = g_tCheckInfo.SeatCount
        lblTicketSells.Caption = ""
        lblCheckedSum.Caption = m_nCheckedTickets
        lblUncheckSum.Caption = ""
    End If
    If m_nOtherTickets > 0 Then
        lblChangeSum.Caption = m_nOtherTickets
    Else
        lblChangeSum.Caption = ""
    End If
End Sub

'显示初始信息
Private Sub ShowInitInfo()
    Dim dtTmp As Date
    Dim lHaveTime As Long
    Dim nTicketsSellCount As Integer          '已售票数
    Dim nMergedSellCount As Integer           '并入车次的售票数
    Dim i As Integer
    On Error GoTo ErrorHandle
    
    lblBusID.Caption = mszBusID & IIf(g_tCheckInfo.BusMode = TP_ScrollBus, "-" & Format(g_tCheckInfo.SerialNo), "")
    lblEndStation.Caption = g_tCheckInfo.EndStationName
    lblVehicle.Caption = g_tCheckInfo.RunVehicle.Vehicle
    lblStartupTime.Caption = IIf(g_tCheckInfo.BusMode <> TP_ScrollBus, Format(g_tCheckInfo.StartUpTime, "HH:mm"), "")
    '途经站点
    Dim aTmp() As TREBusStationInfo
    aTmp = g_oEnvBus.GetBusStation
    lblStationInRoad.Caption = ""
    If ArrayLength(aTmp) > 0 Then
        For i = 1 To ArrayLength(aTmp) - 1
            If aTmp(i).szSellStationID = g_oActiveUser.SellStationID Then
                '仅显示本车站的站点信息
                lblStationInRoad.Caption = lblStationInRoad.Caption & aTmp(i).szStationName & "、"
            End If
        Next i
        '去掉尾部
        If lblStationInRoad.Caption <> "" Then lblStationInRoad.Caption = Left(lblStationInRoad, Len(lblStationInRoad) - 1)
    End If
            
    lblBusType.Caption = IIf(g_tCheckInfo.BusMode = TP_ScrollBus, g_cszTitleScollBus, "固定车次")
    lblVehicleType.Caption = g_tCheckInfo.RunVehicle.VehicleType
    lblCompany.Caption = g_tCheckInfo.Company
    lblOwner.Caption = g_tCheckInfo.RunVehicle.Owner
    lblBeginCheckTime.Caption = Format(g_tCheckInfo.StartCheckTime, cszTimeStr)
    If mbExCheck Then
        lblEndCheckTime.Caption = Format(g_tCheckInfo.StopCheckTime, cszTimeStr)
    Else
        lblEndCheckTime.Caption = ""
    End If
    lblMergeIn.Caption = ""
    lblMergeInSells.Caption = ""
    If g_tCheckInfo.BusMode <> TP_ScrollBus Then
        '得到合并车次信息
        m_MergeBusInfo = g_oChkTicket.GetMergeSeatInfo(mszBusID, Date)
        g_tCheckInfo.MergedBus = g_oChkTicket.GetMergeBus(mszBusID, g_tCheckInfo.StartUpTime)
        g_tCheckInfo.MergeInSells = ArrayLength(m_MergeBusInfo)
        lblMergeIn.Caption = UCase(Trim(g_tCheckInfo.MergedBus))
        If lblMergeIn.Caption <> "" Then
           lblMergeInSells.Caption = g_tCheckInfo.MergeInSells
        End If
        
        m_MergeBusCheckInfo = g_oChkTicket.GetMergeSeatCheckInfo(mszBusID, Date)
        m_nMergeInChecked = ArrayLength(m_MergeBusCheckInfo)
    End If
    
    '玉环，售票总数要所有的票数，而检票数只要用户所有上车站的检票数，座位面板里的信息只要本站的就行了 2007-02-09 by zyw
'    g_tCheckInfo.SellTickets = g_oEnvBus.GetNotCanSellCount(g_oActiveUser.SellStationID)
    g_tCheckInfo.SellTickets = g_oEnvBus.GetNotCanSellCount()
    g_tCheckInfo.SelfSellStationTickets = g_oEnvBus.GetNotCanSellCount(g_oActiveUser.SellStationID)
    m_nSellTickets = g_tCheckInfo.SellTickets
    m_nSelfSellStationTickets = g_tCheckInfo.SelfSellStationTickets
    m_nOldSoldCount = -1
    m_nOldSelftStationSoldCount = -1
'    m_nOldCheckOtherStationCount = -1
'    m_nOldOtherStationCheckCount = -1
    lblTicketSells.Caption = g_tCheckInfo.SellTickets
    
    '初始化座位板信息
'    InitSeatPadInfo
'    RefreshCheckCountInfo
    RefreshSeat
    
    If g_tCheckInfo.BusMode = TP_ScrollBus Then
        If mbExCheck Then         '设置参考检票时间
            rvtLostTime.Second = g_nExtraCheckTime * 60
        Else
            rvtLostTime.Second = g_nCheckTicketTime * 60
        End If
    Else
        If mbExCheck Then         '设置参考检票时间
            rvtLostTime.Second = DateDiff("s", Time, Format(g_tCheckInfo.StartUpTime, "hh:mm:ss")) + g_nExtraCheckTime * 60   ''GetSecondOfTime(CDate(lblStartupTime.Caption)) - GetSecondOfTime(Time) +
        Else
            rvtLostTime.Second = DateDiff("s", Time, Format(g_tCheckInfo.StartUpTime, "hh:mm:ss")) + 3 * 60 'GetSecondOfTime(CDate(lblStartupTime.Caption)) - GetSecondOfTime(Time) + 3 * 60 '+ g_nCheckTicketTime * 60
        End If
    End If
    EnabledTimer True
    txtTicketID.Text = ""
    lblEndStation2.Caption = ""
    lblTicketID.Caption = ""
    lblTicketStatus.Caption = ""
    lblTicketType.Caption = ""
    lblSeatNo.Caption = ""
    
    lblCardType.Caption = ""
    lblPersonName.Caption = ""
    lblIDCardNo.Caption = ""
    
    If mbExCheck Then
        lblCheckInfo.Caption = "车次补检"
    Else
        lblCheckInfo.Caption = "车次检票"
    End If
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub WriteTicketInfo(TTicketInfo As TInterfaceCheckTicketEx, ByRef bTicketValid As Boolean)
    Dim lResult As Long
    Dim szStatus As String
    Dim nStatus As Integer
    Dim i As Integer
    Dim dyDate As Date
    On Error GoTo ErrorHandle
    nStatus = TTicketInfo.TicketStatus
    
    If TTicketInfo.BusDate <> Date Then
        lblCheckInfo.Caption = "不是今天的票，无法检入!"
        bTicketValid = False
        Exit Sub
    End If
    
    lblTicketID.Caption = txtTicketID.Text
    Select Case TTicketInfo.TicketType
        Case TP_FreeTicket
            lblTicketType.Caption = "免票"
            lblTicketType.ForeColor = vbRed
            PlayEventSound g_tEventSoundPath.FreeTicket
        Case TP_FullPrice
            lblTicketType.Caption = GetTkName(TP_FullPrice)
            lblTicketType.ForeColor = vbBlack
        Case TP_HalfPrice
            lblTicketType.Caption = GetTkName(TP_HalfPrice)
            lblTicketType.ForeColor = vbRed
            PlayEventSound g_tEventSoundPath.HalfTicket
        Case TP_PreferentialTicket1
            lblTicketType.Caption = GetTkName(TP_PreferentialTicket1)
            lblTicketType.ForeColor = vbRed
            PlayEventSound g_tEventSoundPath.PreferentialTicket1
        Case TP_PreferentialTicket2
            lblTicketType.Caption = GetTkName(TP_PreferentialTicket2)
            lblTicketType.ForeColor = vbRed
            PlayEventSound g_tEventSoundPath.PreferentialTicket2
        Case TP_PreferentialTicket3
            lblTicketType.Caption = GetTkName(TP_PreferentialTicket2)
            lblTicketType.ForeColor = vbRed
            PlayEventSound g_tEventSoundPath.PreferentialTicket3
    End Select
    lblSeatNo.Caption = TTicketInfo.SeatNo
    lblEndStation2.Caption = TTicketInfo.StationName
    
    lblCardType.Caption = TTicketInfo.CardType
    lblPersonName.Caption = TTicketInfo.PersonName
    lblIDCardNo.Caption = TTicketInfo.IDCardNo
    
    For i = 1 To 6
        If nStatus And (2 ^ (i - 1)) Then
            Select Case i
                Case 1
                    szStatus = szStatus & "/正常售出"
                Case 2
                    szStatus = szStatus & "/改签售出"
                Case 3
                    szStatus = szStatus & "/废票"
                Case 4
                    szStatus = szStatus & "/被改签"
                Case 5
                    szStatus = szStatus & "/退票"
                Case 6
                    szStatus = szStatus & "/已检"
            End Select
        End If
    Next i
    lblTicketStatus.Caption = szStatus
    If nStatus >= ST_TicketChecked Then
        lblCheckInfo.Caption = "本车票已经被检"
        PlayEventSound g_tEventSoundPath.CheckedTicket
    ElseIf nStatus >= ST_TicketReturned Then
        lblCheckInfo.Caption = "本车票已经被退"
        PlayEventSound g_tEventSoundPath.ReturnedTicket
    ElseIf nStatus >= ST_TicketChanged Then
        lblCheckInfo.Caption = "本车票已经被改签"
        PlayEventSound g_tEventSoundPath.InvalidTicket
    ElseIf nStatus >= ST_TicketCanceled Then
        lblCheckInfo.Caption = "本车票已经作废"
        PlayEventSound g_tEventSoundPath.CanceledTicket
    End If
    If nStatus > ST_TicketSellChange + ST_TicketNormal Then
        bTicketValid = False
    Else
        bTicketValid = True
    End If
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub chkCheckChange_Click()
On Error GoTo ErrorHandle
    If chkCheckChange.Value = vbChecked Then
        MsgboxEx "当前处于改乘模式,允许检入其它车次的车票!", vbExclamation + vbOKOnly
        chkCheckChange.ToolTipText = "点击切换至正常检入方式"
        chkCheckChange.BackColor = &HFF&
    Else
        chkCheckChange.ToolTipText = "点击切换至改乘检入方式"
        chkCheckChange.BackColor = &HE0E0E0
    End If
    txtTicketID.SetFocus
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub cmdAllotStationTicketsInfo_Click()
    frmAllotStationTicketsInfo.m_szBusID = lblBusID.Caption
    frmAllotStationTicketsInfo.m_dtBusDate = Date
    frmAllotStationTicketsInfo.m_nChangeCount = Val(lblChangeSum.Caption)
    frmAllotStationTicketsInfo.m_nMergeCount = Val(lblMergeInSells.Caption)
    frmAllotStationTicketsInfo.Show vbModal
    txtTicketID.SetFocus
End Sub

Private Sub cmdDetailInfo_Click()
    DisPlayTicketInfo lblTicketID.Caption
    txtTicketID.SetFocus
End Sub

Private Sub DisPlayTicketInfo(ByVal pszTicketID As String)

    Dim oTemp As frmTicketInfo
    
    On Error GoTo ErrorHandle
        
    If pszTicketID <> "" Then
        Set oTemp = New frmTicketInfo
        Set oTemp.g_oActiveUser = g_oActiveUser
        oTemp.TicketID = pszTicketID
        oTemp.Show vbModal
        Set oTemp = Nothing
    End If
    
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub


Private Sub cmdFind_Click()

    ChangeVehicle
End Sub

Private Sub cmdRefreshSeat_Click()
    Dim lErrorCode As Long
    ResetEnvBusInfo mszBusID, lErrorCode
    RefreshSeat
    txtTicketID.SetFocus
End Sub

Private Sub cmdStopCheck_Click()
'    RefreshSeat
    StopCheck
End Sub
'停检处理
Public Sub StopCheck(Optional StopCheckMode As Integer = 1)
    '显示停检对话框
    Dim szBusid As String
    Dim frmTemp As New frmStopCheck
    Dim bAutoPrint As Boolean
    Dim nIndex As Integer
    Dim tTmpBusInfo As tCheckBusLstInfo
    
On Error GoTo ErrHandle
    
    
    frmTemp.BusID = mszBusID
    frmTemp.MessageStyle = StopCheckMode
    frmTemp.RefreshForm
    frmTemp.Show vbModal, MDIMain
    If Not frmTemp.ClickButton = vbYes Then
        Exit Sub
    End If
    bAutoPrint = frmTemp.AutoPrint
    Set frmTemp = Nothing
    DoEvents
    '停检
    ShowSBInfo "正在停检车次" & EncodeString(mszBusID) & "..."
    g_oChkTicket.StopCheckTicket mszBusID, mnBusSerialNo
    
    '生成路单
    ShowSBInfo "正在生成路单..."
    If CreateSheet Then
        If bAutoPrint Then       '直接打印路单
            Set frmCheckSheet.g_oActiveUser = g_oActiveUser
            Set frmCheckSheet.moChkTicket = g_oChkTicket
            frmCheckSheet.mszSheetID = g_tCheckInfo.CurrSheetNo
            frmCheckSheet.GetCheckSheetInfo
            frmCheckSheet.PrintSheetReport
        Else        '显示路单表单
            Dim ofrmTmp As frmCheckSheet
            Set ofrmTmp = New frmCheckSheet
            Set ofrmTmp.g_oActiveUser = g_oActiveUser
            Set ofrmTmp.moChkTicket = g_oChkTicket
            ofrmTmp.mbExitAfterPrint = True
            ofrmTmp.mszSheetID = g_tCheckInfo.CurrSheetNo
            ofrmTmp.Show vbModal
        End If
    End If
    
    '更改车次状态为停检
    nIndex = g_cWillCheckBusList.FindItem(mszBusID)
    If nIndex > 0 Then
        tTmpBusInfo = g_cWillCheckBusList.Item(nIndex)
        If tTmpBusInfo.BusMode = TP_ScrollBus Then
            tTmpBusInfo.Status = EREBusStatus.ST_BusNormal
            g_cWillCheckBusList.UpdateOne tTmpBusInfo
            If frmBusList.IsShow Then frmBusList.UpdateWillCheckBusItem 2, tTmpBusInfo  '更新
        Else
            g_cWillCheckBusList.RemoveOne nIndex
            If frmBusList.IsShow Then frmBusList.UpdateWillCheckBusItem 3, tTmpBusInfo  '删除
        End If
        tTmpBusInfo.BusSerial = mnBusSerialNo
        tTmpBusInfo.Status = EREBusStatus.ST_BusStopCheck
        tTmpBusInfo.CheckSheet = g_tCheckInfo.CurrSheetNo
        tTmpBusInfo.StartChkTime = g_tCheckInfo.StartCheckTime
        tTmpBusInfo.StopChkTime = Now
        tTmpBusInfo.Company = g_tCheckInfo.RunVehicle.Company
        tTmpBusInfo.Vehicle = g_tCheckInfo.RunVehicle.Vehicle
        tTmpBusInfo.Owner = g_tCheckInfo.RunVehicle.Owner
        If Not mbExCheck Then       '新增
            g_cCheckedBusList.Addone tTmpBusInfo
            If frmBusList.IsShow Then frmBusList.UpdateCheckedBusItem 1, tTmpBusInfo
        Else
            g_cCheckedBusList.UpdateOne tTmpBusInfo
            If frmBusList.IsShow Then frmBusList.UpdateCheckedBusItem 2, tTmpBusInfo
        End If
    End If
    '更改当前系统路单号，并刷新主界面，将路单号存入注册表
    g_tCheckInfo.CurrSheetNo = Right(NumAdd(g_tCheckInfo.CurrSheetNo, 1), g_nCheckSheetLen)            '预置下一个路单号，不监测错误
    WriteInitReg
    WriteCheckGateInfo
    ShowSBInfo ""
    '关闭主窗体检票窗体栏的显示
    CloseOneCheckLine Val(Me.Tag)
    '退出检票窗体
    WriteNextBus
    Unload Me
    Exit Sub
ErrHandle:
    ShowSBInfo ""
    ShowErrorMsg
End Sub


Private Function CreateSheet() As Boolean
'生成路单
    Dim tTmp As TCheckSheetInfo
    
On Error GoTo ErrHandle
    ShowSBInfo "正在生成路单..."
    Me.MousePointer = vbHourglass
    tTmp = g_oChkTicket.GetCheckSheetInfo(g_tCheckInfo.CurrSheetNo)
    
    '检查路单号是否已被使用
    While Not tTmp.szCheckSheet = ""
        MsgboxEx "此路单已存在,请修改当前路单号!", vbExclamation, g_cszTitle_Error
        frmChangeSheetNo.Show vbModal
        tTmp = g_oChkTicket.GetCheckSheetInfo(g_tCheckInfo.CurrSheetNo)
    Wend
    
    g_oChkTicket.MakeCheckSheet Date, mszBusID, _
         mnBusSerialNo, g_tCheckInfo.CurrSheetNo
    
    ShowSBInfo ""
    Me.MousePointer = vbDefault
    
    CreateSheet = True
    
    Exit Function
ErrHandle:
    ShowSBInfo ""
    Me.MousePointer = vbDefault
    ShowErrorMsg
End Function



Private Sub Form_Activate()
    Dim nFormIndex As Integer
    nFormIndex = Val(Me.Tag)
    g_nCurrLineIndex = nFormIndex
    MDIMain.tbsBusList.Tabs(g_nCurrLineIndex).Selected = True
    MDIMain.abMenu.Bands("mnu_Check").Tools("mnu_Check_Stop").Enabled = True
    txtTicketID.SetFocus
End Sub

Private Sub Form_Deactivate()
    MDIMain.abMenu.Bands("mnu_Check").Tools("mnu_Check_Stop").Enabled = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 4 Then
        Dim nKeyPressed   As Integer
        nKeyPressed = KeyCode - 48
        If nKeyPressed > 0 And nKeyPressed <= MDIMain.tbsBusList.Tabs.Count Then
            MDIMain.tbsBusList.Tabs.Item(nKeyPressed).Selected = True
        End If
        If (Chr(KeyCode) = "I" Or Chr(KeyCode) = "i") Then
            txtTicketID.SetFocus
        ElseIf (Chr(KeyCode) = "G" Or Chr(KeyCode) = "g") Then
            If chkCheckChange.Value = vbChecked Then
                chkCheckChange.Value = vbUnchecked
                chkCheckChange.BackColor = &HE0E0E0
            Else
                chkCheckChange.Value = vbChecked
                chkCheckChange.BackColor = &HFF&
            End If
        End If
    End If
    If KeyCode = vbKeyF2 Then
        If mbExCheck Then     '补检
            If g_oChkTicket.SelectChangeBusAfterCheetIsValid Then
                
            Else
                ChangeVehicle
            End If
        Else
            If g_oChkTicket.SelectChangeBusBeforeCheetIsValid Then
                
            Else
                ChangeVehicle
            End If
        End If
            
    End If
End Sub

Private Sub Form_Load()
    
    
    On Error GoTo ErrorHandle
    
    SelectChangeBusValid  '是否有修改车辆的权限
    
    Me.Tag = Str(g_nCurrLineIndex)
    
    mszBusID = UCase(g_tCheckInfo.BusID)
    mnBusSerialNo = g_tCheckInfo.SerialNo
    mnBusMode = g_tCheckInfo.BusMode
    mbExCheck = g_atCheckLine(g_nCurrLineIndex).ExCheck
    If Not mbExCheck Then
        '正常检票
        chkExtraCheck.Value = vbUnchecked
        g_tCheckInfo.StartCheckTime = Now
    Else
        '补检
        chkExtraCheck.Value = vbChecked
    End If
    g_oChkTicket.CheckTicketBeforInitBus m_oREBus.BusID, m_oREBus.RunDate, m_oREBus
    g_oChkTicket.InitSystemParam g_oActiveUser, False, g_bAllowChangeRide
    m_bStopSuccess = False
    
    m_nCheckedTickets = 0
    m_nOtherTickets = 0
    m_nMergeInChecked = 0
'    If mbExCheck Then
        '如果是补检就才读车次检票张数,改乘人数
        m_nCheckedTickets = ArrayLength(g_oChkTicket.GetBusCheckTicket(Date, mszBusID, mnBusSerialNo))
        m_nOtherTickets = g_oChkTicket.GetChangeTicket(mszBusID, Date)
'    End If

    If mnBusMode = TP_ScrollBus Then
        cmdAllotStationTicketsInfo.Enabled = False
    Else
        cmdAllotStationTicketsInfo.Enabled = True
    End If
    
    
    ShowInitInfo
    If g_bAllowChangeRide = False Or m_oREBus.BusType = TP_ScrollBus Then
        chkCheckChange.Enabled = False
    End If
    On Error GoTo 0
    On Error Resume Next
    
    
    g_oChkTicket.RightChangeBus
    If err.Number <> 0 Then
'        chkCheckChange.Enabled = False
    End If
    On Error GoTo 0
       
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> vbFormCode Then
        MsgboxEx "请先停止检票！", vbExclamation, cszTitleInfo
        Cancel = True
    End If
End Sub

Private Sub Form_Resize()
    If Not Me.WindowState = vbMaximized And MDIMain.ActiveForm Is Me Then Me.WindowState = vbMaximized
    fraOutLine.Left = (Me.ScaleWidth - fraOutLine.Width) / 2
    fraOutLine.Top = (Me.ScaleHeight - fraOutLine.Height) / 2
End Sub



Private Sub rvtLostTime_Timer()
    '播放停检时间到的音效,
    rvtLostTime.Second = 0
    EnabledTimer False
'    PlayEventSound g_tEventSoundPath.CheckTimeOn
'    CloseModalForm
'    StopCheck 0
End Sub
Public Sub EnabledTimer(bEnabled As Boolean)
    If bEnabled Then
'        FlatLabel1.Visible = False
        rvtLostTime.Enabled = True
        rvtLostTime.Visible = True
    Else
'        FlatLabel1.Visible = True
        rvtLostTime.Enabled = False
        rvtLostTime.Visible = False
    End If
End Sub




Private Sub SeatPad1_GridClick(Index As Integer)
'    Dim TSeatInfo As TSeatInfo
'    Dim szTicketID As String
'    On Error GoTo ErrorHandle
'
'    TSeatInfo = g_oChkTicket.GetBusSeatInfo(Format(Date, cszDateStr), lblBusID.Caption, SeatPad1.PadGrids.Item(Index).Caption)
'
'    szTicketID = TSeatInfo.szTicketNo
'
'    DisPlayTicketInfo szTicketID
'
'    txtTicketID.SetFocus
'
'    Exit Sub
'ErrorHandle:
'    ShowErrorMsg
End Sub

Private Sub txtTicketID_Change()
    If Len(txtTicketID.Text) >= 10 Then
        txtTicketID.Text = Left(txtTicketID.Text, 10)
    End If
End Sub

Private Sub txtTicketID_GotFocus()
    txtTicketID.SelStart = 0
    txtTicketID.SelLength = Len(txtTicketID.Text)
End Sub


Private Sub txtTicketID_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        txtTicketID.Text = NumSub(txtTicketID.Text, 1)
    ElseIf KeyCode = vbKeyUp Then
        txtTicketID.Text = NumAdd(txtTicketID.Text, 1)
    End If
    
End Sub

Private Sub txtTicketID_KeyPress(KeyAscii As Integer)
    Dim bCheckSucced As Boolean     '检票是否成功
    If KeyAscii = vbKeyReturn Then
        lblCheckInfo.Caption = "正在检查车票的有效性..."
        CheckTicket bCheckSucced
        If bCheckSucced Then
            lblCheckInfo.Caption = "车票成功检入"
            imgEnabled.Visible = False
            If lblTicketType.Caption = GetTkName(TP_FullPrice) Then
                PlayEventSound g_tEventSoundPath.CheckSucess
            End If
        Else
            imgEnabled.Visible = True
        End If
        txtTicketID.SelStart = 0
        txtTicketID.SelLength = Len(txtTicketID.Text)
    Else
        If KeyAscii = 32 Then
            KeyAscii = 0
        End If
    End If
End Sub

'Private Sub MakeCheckSheet()
''生成路单
''如果是一般停检，直接生成路单
''如果补检停检，提示是否生成路单
'Dim szOldSheet As String
'On Error GoTo ErrHandle
'
'    g_oChkTicket.MakeCheckSheet Date, mszBusID, mnBusSerialNo, g_tCheckInfo.CurrSheetNo
'
'    szOldSheet = g_tCheckInfo.CurrSheetNo
'    g_tCheckInfo.CurrSheetNo = NumAdd(g_tCheckInfo.CurrSheetNo, 1)
'    WriteInitReg
'    WriteCheckGateInfo
'
'    Dim oTemp As CheckSysApp
'    Set oTemp = New CheckSysApp
'    oTemp.ShowCheckSheet g_oActiveUser, szOldSheet, False, False, False, False
'    Exit Sub
'ErrHandle:
'    ShowErrorMsg
'End Sub

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
Public Sub RefreshSeat()
    '重新得到检票数及改乘数
    m_nCheckedTickets = ArrayLength(g_oChkTicket.GetBusCheckTicket(Date, mszBusID, mnBusSerialNo, g_tCheckInfo.CheckGateNo))
    m_nOtherTickets = g_oChkTicket.GetChangeTicket(mszBusID, Date)
    
    m_MergeBusCheckInfo = g_oChkTicket.GetMergeSeatCheckInfo(mszBusID, Date)
    m_nMergeInChecked = ArrayLength(m_MergeBusCheckInfo)
    '
    InitSeatPadInfo
    RefreshCheckCountInfo
End Sub
Private Function GetTkName(TicketTypeID As Integer) As String     '根据票种类型得到票种名称
    Dim nCount As Integer
    Dim szChar As String
    Dim i As Integer

    nCount = ArrayLength(g_tTicketType)
    For i = 1 To nCount
        If g_tTicketType(i).nTicketTypeID = TicketTypeID Then
              szChar = g_tTicketType(i).szTicketTypeName
              Exit For
        End If
    Next i
    GetTkName = szChar
End Function
'显示并班车次的信息
Private Sub ShowBusMergeInfo()
    Dim i As Integer
    Dim nCount As Integer
    On Error GoTo ErrorHandle
    
    m_MergeBusInfo = g_oChkTicket.GetMergeSeatInfo(mszBusID, Date)
    g_tCheckInfo.MergedBus = g_oChkTicket.GetMergeBus(mszBusID, g_tCheckInfo.StartUpTime)
    g_tCheckInfo.MergeInSells = ArrayLength(m_MergeBusInfo)
    lblMergeIn.Caption = UCase(Trim(g_tCheckInfo.MergedBus))
    If lblMergeIn.Caption <> "" Then
       lblMergeInSells.Caption = g_tCheckInfo.MergeInSells
    Else
       lblMergeInSells.Caption = ""
    End If
    
    m_MergeBusCheckInfo = g_oChkTicket.GetMergeSeatCheckInfo(mszBusID, Date)
    m_nMergeInChecked = ArrayLength(m_MergeBusCheckInfo)
    
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub
'判断本车票是否是并班车次
Private Function IsMergeBus(szBusid As String, szChar As String, szTicketID As String) As Boolean
    Dim i As Integer, nCount As Integer
    Dim szTemp() As String
    Dim bStatus As Boolean
    szTemp = StringToTeam(szChar)
    nCount = ArrayLength(szTemp)
    For i = 1 To nCount
      If szTemp(i) = szBusid Then
        bStatus = True
        Exit For
      Else
        IsMergeBus = False
      End If
    Next
    If bStatus = True Then
'      m_MergeBusInfo = g_oChkTicket.GetMergeSeatInfo(mszBusID, Date)
      nCount = ArrayLength(m_MergeBusInfo)
      If nCount = 0 Then IsMergeBus = False: Exit Function
      For i = 1 To nCount
        If Trim(m_MergeBusInfo(i, 3)) = txtTicketID.Text Then
            g_tCheckInfo.SplitSeat = m_MergeBusInfo(i, 4)
            IsMergeBus = True
            Exit For
        Else
            IsMergeBus = False
        End If
      Next i
    End If
End Function


Private Sub ChangeVehicle()
    Dim oShell As New STShell.CommDialog
    Dim aszTemp() As String
    Dim oVehicle As New Vehicle
    On Error GoTo ErrorHandle
'    If Not mbExCheck Then
        oShell.Init g_oActiveUser
        aszTemp = oShell.SelectVehicleEX
        If ArrayLength(aszTemp) > 0 Then
            oVehicle.Init g_oActiveUser
            oVehicle.Identify aszTemp(1, 1)
            If oVehicle.SeatCount < m_nCheckedTickets And g_tCheckInfo.BusMode = TP_ScrollBus Then
                '如果为滚动车次且,总座位数小于已检票数则.
                MsgBox "更改的车辆的座位数不能小于已检票数,无法更改", vbExclamation, Me.Caption
                Exit Sub
            
            End If
            g_oChkTicket.ChangeVehicle aszTemp(1, 1), mnBusSerialNo, mszBusID, Date
            g_tCheckInfo.RunVehicle.VehicleId = aszTemp(1, 1)
            
            lblVehicle.Caption = aszTemp(1, 2)
            InitSeatPadInfo
            RefreshCheckCountInfo
            MsgBox "车辆已改为" & aszTemp(1, 2)
            
        End If
        txtTicketID.SetFocus
'    End If
    Set oShell = Nothing
    Exit Sub
ErrorHandle:
    Set oShell = Nothing
    ShowErrorMsg
End Sub
Private Sub cmdLedShow_Click()
On Error GoTo ErrHandle
    If MsgBox("您是否将本班次检票信息插播到检票条屏？", vbQuestion + vbYesNo, "检票") = vbNo Then Exit Sub
    
    cmdLedShow.Enabled = False
    g_oChkTicket.SetLED mszBusID, Date, 1
    
    
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

'是否有更改车辆的权限 开检后  分补检和不补检两种权限
'判断当前用户有更改车辆的权限 开检前的权限
Private Sub SelectChangeBusValid()
    On Error GoTo Here
    If mbExCheck Then     '补检
        If g_oChkTicket.SelectChangeBusAfterCheetIsValid Then
            cmdFind.Enabled = True
        Else
            cmdFind.Enabled = False
        End If
    Else
        If g_oChkTicket.SelectChangeBusBeforeCheetIsValid Then
            cmdFind.Enabled = True
        Else
            cmdFind.Enabled = False
        End If
    End If
    On Error GoTo 0
    Exit Sub
Here:
    ShowErrorMsg
End Sub
