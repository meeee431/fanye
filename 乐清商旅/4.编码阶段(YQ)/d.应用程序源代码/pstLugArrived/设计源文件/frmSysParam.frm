VERSION 5.00
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmSysParam 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "设置系统参数"
   ClientHeight    =   4095
   ClientLeft      =   5910
   ClientTop       =   2220
   ClientWidth     =   6075
   HelpContextID   =   7000201
   Icon            =   "frmSysParam.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   6075
   Begin VB.TextBox txtKeepFeeDays 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1920
      TabIndex        =   14
      Top             =   1440
      Width           =   1245
   End
   Begin VB.TextBox txtKeepCharge 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1680
      TabIndex        =   5
      Top             =   1020
      Width           =   1485
   End
   Begin VB.TextBox txtSheetIDLen 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Top             =   180
      Width           =   1485
   End
   Begin VB.TextBox txtTransRatio 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1680
      TabIndex        =   3
      Top             =   600
      Width           =   1485
   End
   Begin RTComctl3.CoolButton cmdExit 
      Cancel          =   -1  'True
      Height          =   345
      Left            =   4800
      TabIndex        =   8
      Top             =   3630
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "取消"
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
      MICON           =   "frmSysParam.frx":000C
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
      BackColor       =   &H00E0E0E0&
      Caption         =   "短信息设置"
      Height          =   1545
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   5805
   End
   Begin RTComctl3.CoolButton cmdSave 
      Default         =   -1  'True
      Height          =   345
      Left            =   3570
      TabIndex        =   7
      Top             =   3630
      Width           =   1125
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "确定"
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
      MICON           =   "frmSysParam.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   " "
      Height          =   3120
      Left            =   -90
      TabIndex        =   10
      Top             =   3360
      Width           =   8745
      Begin RTComctl3.CoolButton cmdHelp 
         Height          =   315
         Left            =   240
         TabIndex        =   12
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
         MICON           =   "frmSysParam.frx":0044
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
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "天"
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
      Left            =   3240
      TabIndex        =   15
      Top             =   1440
      Width           =   240
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "保管费免费天数(&M):"
      Height          =   180
      Left            =   360
      TabIndex        =   13
      Top             =   1507
      Width           =   1620
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "保管费(&K):"
      Height          =   180
      Left            =   390
      TabIndex        =   4
      Top             =   1080
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "元/天件"
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
      Left            =   3240
      TabIndex        =   11
      Top             =   1050
      Width           =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "票据单长度(&C):"
      Height          =   180
      Left            =   390
      TabIndex        =   0
      Top             =   240
      Width           =   1260
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "%"
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
      Left            =   3240
      TabIndex        =   9
      Top             =   630
      Width           =   120
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "代收手续费(&R):"
      Height          =   180
      Left            =   390
      TabIndex        =   2
      Top             =   660
      Width           =   1260
   End
End
Attribute VB_Name = "frmSysParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'行包参数
Dim nSheetIDLen As Integer '行包单号数字部分长
Dim dbKeepCharge As Double  '签发单数字部分长度
Dim dbTransRatio As Double '行包受理单退运费率1
Dim dbKeepFeeDays As Integer '保管费免费天数

Private Sub cmdExit_Click()
    Unload Me

End Sub

Private Sub cmdHelp_Click()
    DisplayHelp Me
End Sub

Private Sub cmdSave_Click()
    On Error GoTo Here

    If nSheetIDLen <> txtSheetIDLen.Text Then g_oPackageParam.SheetIDNumberLen = txtSheetIDLen.Text
    nSheetIDLen = txtSheetIDLen.Text
    If dbKeepCharge <> txtKeepCharge.Text Then g_oPackageParam.NormalKeepCharge = txtKeepCharge.Text
    dbKeepCharge = txtKeepCharge.Text
    If dbTransRatio <> txtTransRatio.Text Then g_oPackageParam.TransitChargeRatio = Val(txtTransRatio.Text) / 100
    dbTransRatio = txtTransRatio.Text
    
    If dbKeepFeeDays <> txtKeepFeeDays.Text Then g_oPackageParam.KeepFeeDays = txtKeepFeeDays.Text
    dbKeepFeeDays = txtKeepFeeDays.Text
    
'    '行包注意事项
'    Dim oFreeReg As CFreeReg
'    Set oFreeReg = New CFreeReg
'    oFreeReg.Init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany
'    oFreeReg.SaveSetting cszLuggageAccount, "CareContent", Trim(txtCare.Text)
    
    
    
    Unload Me
    Exit Sub
Here:

    ShowErrorMsg
End Sub

Private Sub Form_Load()
    On Error GoTo Here
    AlignFormPos Me



    nSheetIDLen = g_oPackageParam.SheetIDNumberLen
    dbKeepCharge = g_oPackageParam.NormalKeepCharge
    dbTransRatio = g_oPackageParam.TransitChargeRatio * 100
    dbKeepFeeDays = g_oPackageParam.KeepFeeDays
    
    txtSheetIDLen.Text = nSheetIDLen
    txtKeepCharge.Text = dbKeepCharge
    txtTransRatio.Text = dbTransRatio
    txtKeepFeeDays.Text = dbKeepFeeDays
    
'    行包注意事项
'    Dim oFreeReg As CFreeReg
'    Set oFreeReg = New CFreeReg
'    oFreeReg.Init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany
'    txtCare.Text = oFreeReg.GetSetting(cszLuggageAccount, "CareContent")
    
    Exit Sub
Here:
    ShowErrorMsg
End Sub

Private Sub Form_Unload(Cancel As Integer)
   SaveFormPos Me
End Sub

Private Sub txtKeepCharge_Change()
    FormatTextToNumeric txtKeepCharge, False, True

End Sub

Private Sub txtSheetIDLen_Change()
    FormatTextToNumeric txtSheetIDLen, False, False

End Sub


Private Sub txtTransRatio_Change()
    FormatTextToNumeric txtTransRatio, False, True
End Sub
