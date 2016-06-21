VERSION 5.00
Object = "{BBF95DAA-F9CB-4CA9-A673-E0E9E0193752}#1.0#0"; "STSellCtl.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Begin VB.Form frmSheetBus 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "签发车次填写"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7890
   Icon            =   "frmSheetBus.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   7890
   StartUpPosition =   3  '窗口缺省
   Begin RTComctl3.CoolButton cmdcancel 
      Height          =   375
      Left            =   4860
      TabIndex        =   7
      Top             =   7410
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "取消(&C)"
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
      MICON           =   "frmSheetBus.frx":038A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdOk 
      Height          =   375
      Left            =   6390
      TabIndex        =   6
      Top             =   7410
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "完成(&H)"
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
      MICON           =   "frmSheetBus.frx":03A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ComboBox cboCarryVehicle 
      Height          =   300
      Left            =   5730
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2970
      Width           =   1995
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "已签发的单子信息"
      Height          =   1935
      Left            =   60
      TabIndex        =   30
      Top             =   840
      Width           =   7725
      Begin VB.Label lblOprationTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "12:50"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4260
         TabIndex        =   48
         Top             =   1590
         Width           =   525
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "操作时间："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3090
         TabIndex        =   47
         Top             =   1590
         Width           =   1050
      End
      Begin VB.Label lblChecker 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "wjb"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1140
         TabIndex        =   46
         Top             =   1560
         Width           =   315
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "操作员："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   45
         Top             =   1560
         Width           =   840
      End
      Begin VB.Label lblSheetID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1234567"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2190
         TabIndex        =   44
         Top             =   390
         Width           =   735
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "已签发的签发单号："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   43
         Top             =   390
         Width           =   1890
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "签发件数:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   7
         Left            =   240
         TabIndex        =   42
         Top             =   1140
         Width           =   945
      End
      Begin VB.Label lblTotalPrice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "234324"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   6030
         TabIndex        =   41
         Top             =   1200
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "总计重:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   8
         Left            =   240
         TabIndex        =   40
         Top             =   750
         Width           =   735
      End
      Begin VB.Label lblBillNumber 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "123123"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3420
         TabIndex        =   39
         Top             =   1170
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "总实重:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   9
         Left            =   2370
         TabIndex        =   38
         Top             =   780
         Width           =   735
      End
      Begin VB.Label lblBagNumber 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "12323"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1230
         TabIndex        =   37
         Top             =   1140
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "行包单数:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   2370
         TabIndex        =   36
         Top             =   1170
         Width           =   945
      End
      Begin VB.Label lblActWeight 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "12312"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3240
         TabIndex        =   35
         Top             =   780
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "超重件数:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   13
         Left            =   4890
         TabIndex        =   34
         Top             =   780
         Width           =   945
      End
      Begin VB.Label lblOverNumber 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "123123"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   6060
         TabIndex        =   33
         Top             =   780
         Width           =   630
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "托运总价:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4860
         TabIndex        =   32
         Top             =   1200
         Width           =   945
      End
      Begin VB.Label lblCalWeight 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2313"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1110
         TabIndex        =   31
         Top             =   750
         Width           =   420
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "承运车辆信息"
      Height          =   945
      Left            =   120
      TabIndex        =   19
      Top             =   6300
      Width           =   7665
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车牌号:"
         Height          =   180
         Index           =   4
         Left            =   210
         TabIndex        =   29
         Top             =   270
         Width           =   630
      End
      Begin VB.Label lblVehicle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "浙BA123456"
         Height          =   180
         Left            =   840
         TabIndex        =   28
         Top             =   270
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车主:"
         Height          =   180
         Index           =   5
         Left            =   1980
         TabIndex        =   27
         Top             =   270
         Width           =   450
      End
      Begin VB.Label lblOwner 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "江泽民"
         Height          =   180
         Left            =   2430
         TabIndex        =   26
         Top             =   270
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "参运公司:"
         Height          =   180
         Index           =   6
         Left            =   3420
         TabIndex        =   25
         Top             =   270
         Width           =   810
      End
      Begin VB.Label lblCompany 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "国务院"
         Height          =   180
         Left            =   4230
         TabIndex        =   24
         Top             =   270
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "拆帐公司:"
         Height          =   180
         Index           =   0
         Left            =   210
         TabIndex        =   23
         Top             =   570
         Width           =   810
      End
      Begin VB.Label lblSplitCompany 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RT"
         Height          =   180
         Left            =   1020
         TabIndex        =   22
         Top             =   570
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "拆算协议:"
         Height          =   180
         Index           =   12
         Left            =   3420
         TabIndex        =   21
         Top             =   570
         Width           =   810
      End
      Begin VB.Label lblProtocal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "9：1"
         Height          =   180
         Left            =   4230
         TabIndex        =   20
         Top             =   570
         Width           =   360
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "签发车次信息"
      Height          =   975
      Left            =   120
      TabIndex        =   10
      Top             =   5220
      Width           =   7665
      Begin VB.Label lblStations 
         BackStyle       =   0  'Transparent
         Caption         =   "中华人民共和国、中华人民共和国、中华人民共和国、中华人民共和国、中华人民共和国、中华人民共和国、中华人民共和国"
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   1050
         TabIndex        =   18
         Top             =   540
         Width           =   5100
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "途经站点:"
         Height          =   180
         Index           =   17
         Left            =   210
         TabIndex        =   17
         Top             =   540
         Width           =   810
      End
      Begin VB.Label lblBusID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "123123"
         Height          =   180
         Left            =   660
         TabIndex        =   16
         Top             =   270
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车次:"
         Height          =   180
         Index           =   18
         Left            =   210
         TabIndex        =   15
         Top             =   270
         Width           =   450
      End
      Begin VB.Label lblBusDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2020-01-02"
         Height          =   180
         Left            =   2760
         TabIndex        =   14
         Top             =   270
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车次日期:"
         Height          =   180
         Index           =   10
         Left            =   1950
         TabIndex        =   13
         Top             =   270
         Width           =   810
      End
      Begin VB.Label lblStartTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2020-02-02"
         Height          =   180
         Left            =   4800
         TabIndex        =   12
         Top             =   270
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "发车时间:"
         Height          =   180
         Index           =   11
         Left            =   3990
         TabIndex        =   11
         Top             =   270
         Width           =   810
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   7875
      TabIndex        =   0
      Top             =   0
      Width           =   7875
      Begin RTComctl3.CoolButton cmdSheet 
         Height          =   345
         Left            =   5640
         TabIndex        =   2
         Top             =   180
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "确认(&D)"
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
         MICON           =   "frmSheetBus.frx":03C2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox txtSheetID 
         Appearance      =   0  'Flat
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1980
         TabIndex        =   1
         Top             =   210
         Width           =   2505
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "已签发的签发单号："
         Height          =   180
         Left            =   210
         TabIndex        =   8
         Top             =   300
         Width           =   1620
      End
      Begin VB.Image Image3 
         Height          =   855
         Left            =   2820
         Picture         =   "frmSheetBus.frx":03DE
         Top             =   -60
         Width           =   5925
      End
   End
   Begin FText.asFlatTextBox txtEndStation 
      Height          =   315
      Left            =   1530
      TabIndex        =   3
      Top             =   3000
      Width           =   2490
      _ExtentX        =   4392
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
      ButtonHotBackColor=   -2147483633
      ButtonPressedBackColor=   -2147483627
      Text            =   ""
      ButtonBackColor =   -2147483633
      ButtonVisible   =   -1  'True
   End
   Begin STSellCtl.ucSuperCombo cboBus 
      Height          =   1635
      Left            =   1530
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3480
      Width           =   6225
      _ExtentX        =   10980
      _ExtentY        =   2884
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "承运车辆:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4410
      TabIndex        =   50
      Top             =   3030
      Width           =   945
   End
   Begin VB.Label lblToStation 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "到站:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   330
      TabIndex        =   49
      Top             =   3030
      Width           =   525
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "承运车次:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   330
      TabIndex        =   9
      Top             =   3510
      Width           =   945
   End
End
Attribute VB_Name = "frmSheetBus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboBus_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   cboCarryVehicle.SetFocus
End If
End Sub

Private Sub cboBus_LostFocus()
On Error GoTo here
    Dim i As Integer
    Dim j As Integer
    Dim nlen As Integer
    Dim szTempStationName() As String
    Dim szTemp() As String
    Dim szaTemp() As String
    Dim Count As Integer
           If cboBus.BoundText <> "" Then
            moCarrySheet.BusDate = ToDBDate(Date)
            moCarrySheet.BusID = Trim(cboBus.BoundText)
            moCarrySheet.RefreshBusVehicle
            lblBusID.Caption = moCarrySheet.BusID
            lblBusDate.Caption = CStr(moCarrySheet.BusDate)
            lblStartTime.Caption = CStr(Format(moCarrySheet.BusStartOffTime, "hh:mm"))
            
            szaTemp = moLugSvr.GetBusStationNames(Date, Trim(cboBus.BoundText))
            Count = ArrayLength(szaTemp)
            ReDim szTempStationName(1 To Count)
            szTempStationName = szaTemp
            lblStations = ""
            For i = 1 To Count
                lblStations.Caption = lblStations.Caption + " " + szTempStationName(i)
            Next
            
            
            nlen = ArrayLength(moLugSvr.GetBusRunVehicles(Trim(cboBus.BoundText)))
            If nlen > 0 Then
                ReDim szTemp(1 To nlen)
                szTemp = moLugSvr.GetBusRunVehicles(Trim(cboBus.BoundText))
                cboCarryVehicle.Clear
                For i = 1 To nlen
                    cboCarryVehicle.AddItem MakeDisplayString(szTemp(i, 1), szTemp(i, 2))
                Next i
                 If nlen <> 0 Then cboCarryVehicle.ListIndex = 0
            End If
  
  End If

Exit Sub
here:
ShowErrorMsg
End Sub

Private Sub cboCarryVehicle_Click()
cboCarryVehicle_Change
End Sub

Private Sub cboCarryVehicle_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And cboCarryVehicle.Text <> "" Then
   cmdOK.Enabled = True
   cmdOK.SetFocus
End If
End Sub

Private Sub cboCarryVehicle_Change()
On Error GoTo ErrHandle
  
    If cboBus.BoundText <> "" And cboCarryVehicle.Text <> "" Then
        cmdOK.Enabled = True
        moCarrySheet.BusDate = Date
        moCarrySheet.BusID = Trim(cboBus.BoundText)
        moCarrySheet.VehicleID = ResolveDisplay(cboCarryVehicle.Text)
        moCarrySheet.VehicleLicense = ResolveDisplayEx(cboCarryVehicle.Text)
        moCarrySheet.RefreshBusVehicle
        lblVehicle.Caption = moCarrySheet.VehicleLicense
        lblOwner.Caption = moCarrySheet.BusOwnerName
        lblCompany.Caption = moCarrySheet.CompanyName
        lblSplitCompany.Caption = moCarrySheet.SplitCompanyName
        lblProtocal.Caption = moCarrySheet.ProtocolName
    End If
Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub cboCarryVehicle_LostFocus()
cboCarryVehicle_Change
End Sub

Private Sub cmdCancel_Click()
 Unload Me
End Sub

Private Sub cmdOk_Click()
Dim mAnswer
On Error GoTo here
    If lblVehicle.Caption = "" Then
       MsgBox "请正确填写完整所有信息！", vbExclamation, Me.Caption
       Exit Sub
    End If
    mAnswer = MsgBox("您输入的车号是  " & Trim(lblVehicle.Caption) & "   是否确认？", vbInformation + vbYesNo, Me.Caption)
    If mAnswer = vbYes Then
    moCarrySheet.SheetID = Trim(lblSheetID.Caption)
    moLugSvr.UpdateCarryLuggage moCarrySheet
   
    Unload Me
    Else
    Exit Sub
    End If
Exit Sub
here:
ShowErrorMsg
End Sub
Private Sub FormClear()
  lblChecker.Caption = ""
  lblSheetID.Caption = ""
  lblOprationTime.Caption = ""
  txtEndStation.Text = ""
  cboBus.Clear
  cboCarryVehicle.Clear
  lblBusID.Caption = ""
  lblBusDate.Caption = ""
  lblStartTime.Caption = ""
  lblStations.Caption = ""
  lblVehicle.Caption = ""
  lblOwner.Caption = ""
  lblCompany.Caption = ""
  lblSplitCompany.Caption = ""
  lblProtocal.Caption = ""
  lblTotalPrice.Caption = ""
  lblBillNumber.Caption = ""
  lblBagNumber.Caption = ""
  lblCalWeight.Caption = ""
  lblActWeight.Caption = ""
  lblOverNumber.Caption = ""
End Sub
Private Sub cmdSheet_Click()
 Dim rsTemp As Recordset
 
  On Error GoTo here
  If Trim(txtSheetID.Text) = "" Then
     MsgBox "已签发的签发单号不能为空！", vbExclamation, Me.Caption
     txtSheetID.SetFocus
     Exit Sub
  End If
  Set rsTemp = moLugSvr.GetOldSheetRs(Trim(txtSheetID.Text))
  If rsTemp.RecordCount = 0 Then
     MsgBox "没有此签发单号！", vbExclamation, Me.Caption
     txtSheetID.SetFocus
     Exit Sub
  End If
  If Trim(rsTemp!license_tag_no) <> "" Then
     MsgBox "此签发单已有车次，不允许修改！", vbInformation, Me.Caption
     Exit Sub
  End If
'  moCarrySheet.AddNew
  lblSheetID.Caption = Trim(rsTemp!sheet_id)
  lblCalWeight.Caption = rsTemp!cal_weight
  lblActWeight.Caption = rsTemp!fact_weight
  lblOverNumber.Caption = rsTemp!over_number
  lblBagNumber.Caption = rsTemp!luggage_number
  lblBillNumber.Caption = rsTemp!baggage_number
  lblTotalPrice.Caption = rsTemp!total_price
  lblChecker.Caption = rsTemp!checker
  lblOprationTime.Caption = rsTemp!sheet_make_time
  Me.Height = 8265
  Me.Top = Screen.Height / 5
  Exit Sub
here:
ShowErrorMsg
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 And (Me.ActiveControl Is cboBus) Then
'   cboCarryVehicle.SetFocus
'ElseIf KeyAscii = 13 And (Me.ActiveControl Is txtEndStation) Then
'   cboBus.SetFocus
'ElseIf KeyAscii = 13 And (Me.ActiveControl Is cboCarryVehicle) Then
'   cmdOk.SetFocus
'End If
End Sub

Private Sub Form_Load()
   Me.Top = Screen.Height / 3
   Me.Left = (Screen.Width - Me.Width) / 2
   Me.Height = 3285
   FormClear
   
End Sub

Private Sub txtEndStation_ButtonClick()
On Error GoTo ErrHandle
    Dim oShell As New STShell.CommDialog
    Dim aszTemp() As String
    oShell.Init m_oAUser
    aszTemp = oShell.SelectStation()
    Set oShell = Nothing
    If ArrayLength(aszTemp) = 0 Then Exit Sub
    txtEndStation.Text = aszTemp(1, 1) & "[" & Trim(aszTemp(1, 2)) & "]"
    
Exit Sub
ErrHandle:
ShowErrorMsg
End Sub

Private Sub RefreshBus()
On Error GoTo here
    Dim rsTemp As Recordset
    Dim szBusEndStation As String
    szBusEndStation = ResolveDisplay(txtEndStation.Text)
    Set rsTemp = moLugSvr.GetToStationBusRS(szBusEndStation, Date)
    If rsTemp.RecordCount > 0 Then
      With cboBus
        Set .RowSource = rsTemp
        'bus_id:车次代码
        'bus_start_time:发车日期
        'vehicle_type_name:车型
        
        .BoundField = "bus_id"
'        .ListFields = "bus_id:4,vehicle_type_name:4"
        .AppendWithFields "bus_id:8,bus_start_time:19,vehicle_type_name"
      End With
     Else
      MsgBox " 当前时间没有车次信息!", vbInformation, Me.Caption
      
     End If
    Set rsTemp = Nothing
    On Error GoTo 0
    Exit Sub
here:
ShowErrorMsg
End Sub



Private Sub txtEndStation_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   cboBus.SetFocus
End If
End Sub

Private Sub txtEndStation_LostFocus()
On Error GoTo ErrHandle
 If txtEndStation.Text <> "" Then
    RefreshBus
 End If
 
 Exit Sub
ErrHandle:
 ShowErrorMsg
End Sub

Private Sub txtSheetID_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   cmdSheet.SetFocus
End If
End Sub
