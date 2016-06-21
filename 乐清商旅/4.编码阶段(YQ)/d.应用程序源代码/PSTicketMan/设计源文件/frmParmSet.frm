VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmParmSet 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "票证系统参数设置"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   -240
      TabIndex        =   5
      Top             =   690
      Width           =   7845
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   -180
      ScaleHeight     =   705
      ScaleWidth      =   7785
      TabIndex        =   3
      Top             =   0
      Width           =   7785
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "系统参数设置"
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   480
         TabIndex        =   4
         Top             =   240
         Width           =   90
      End
   End
   Begin VB.TextBox txtTicketNum 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      MaxLength       =   2
      TabIndex        =   1
      Text            =   "0"
      Top             =   960
      Width           =   1935
   End
   Begin MSComctlLib.ImageList imgRoute 
      Left            =   0
      Top             =   0
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
            Picture         =   "frmParmSet.frx":0000
            Key             =   "Station"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParmSet.frx":015A
            Key             =   "NoSell"
         EndProperty
      EndProperty
   End
   Begin RTComctl3.CoolButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   3420
      TabIndex        =   7
      Top             =   1680
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   556
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
      MICON           =   "frmParmSet.frx":02B6
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
      Height          =   315
      Left            =   2070
      TabIndex        =   8
      Top             =   1680
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   556
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
      MICON           =   "frmParmSet.frx":02D2
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
      Enabled         =   0   'False
      Height          =   3240
      Left            =   -120
      TabIndex        =   9
      Top             =   1440
      Width           =   8745
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "张"
      Height          =   255
      Left            =   4200
      TabIndex        =   2
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "在车票剩余几张时提示"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   1935
   End
End
Attribute VB_Name = "frmParmSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  Dim oTicketMan As New TicketMan

Private Sub cmdCancel_Click()
    Unload Me
  
End Sub

Private Sub cmdOk_Click()
Dim bupdate As Boolean
oTicketMan.Init m_oActiveUser
bupdate = oTicketMan.upDateTicketParm("HowTicketts", Label1.Caption, CInt(txtTicketNum.Text), "")
If bupdate = True Then
    MsgBox "系统参数修改成功!", vbOKOnly + vbInformation, "提示"
Else
    MsgBox "系统参数修改失败!,请检查是否输入正确！", vbOKOnly + vbInformation, "提示"
End If
End Sub

Private Sub Form_Load()
    AlignFormPos Me
   txtTicketNum.Text = GetTicketNum
End Sub

Private Sub Form_Unload(Cancel As Integer)
  SaveFormPos Me
End Sub
