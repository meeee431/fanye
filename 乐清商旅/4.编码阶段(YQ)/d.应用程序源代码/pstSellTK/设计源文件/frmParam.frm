VERSION 5.00
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmParam 
   BackColor       =   &H00E0E0E0&
   Caption         =   "售票参数"
   ClientHeight    =   3375
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   4920
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtStopSellTime 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   1800
      TabIndex        =   1
      Top             =   930
      Width           =   1005
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   " "
      Height          =   840
      Left            =   -30
      TabIndex        =   0
      Top             =   2640
      Width           =   8745
      Begin RTComctl3.CoolButton cmdCancel 
         Cancel          =   -1  'True
         Height          =   345
         Left            =   3600
         TabIndex        =   4
         Top             =   270
         Width           =   1125
         _ExtentX        =   0
         _ExtentY        =   0
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
         MICON           =   "frmParam.frx":0000
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
         Default         =   -1  'True
         Height          =   345
         Left            =   2280
         TabIndex        =   5
         Top             =   270
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "确定(&O)"
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
         MICON           =   "frmParam.frx":001C
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
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "售票参数设置"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   180
      TabIndex        =   3
      Top             =   240
      Width           =   1080
   End
   Begin VB.Label Label56 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "停售时间\分钟(&T):"
      Height          =   180
      Left            =   180
      TabIndex        =   2
      Top             =   1005
      Width           =   1530
   End
End
Attribute VB_Name = "frmParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Dim oReg As New CFreeReg
    On Error GoTo here
    
    oReg.Init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany
    
    oReg.SaveSetting cszPrimaryKey, "StopSellTime", Val(txtStopSellTime.Text)

    Set oReg = Nothing
       
    ShowMsg "保存成功"
    Unload Me
    
    Exit Sub
    
here:
    Set oReg = Nothing
    ShowErrorMsg
End Sub

Private Sub Form_Load()
    Dim oReg As New CFreeReg
    
    On Error GoTo here
    
    Dim oParam As New SystemParam
    Dim nStopSellTime As String
    oParam.Init m_oAUser
    nStopSellTime = oParam.StopSellTime
    
    oReg.Init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany
    
    txtStopSellTime.Text = oReg.GetSetting(cszPrimaryKey, "StopSellTime", nStopSellTime)
    
    Set oReg = Nothing
    
    Exit Sub
    
here:
    Set oReg = Nothing
    ShowErrorMsg
End Sub

Private Sub txtStopSellTime_Change()
    FormatTextToNumeric txtStopSellTime, False, False
End Sub

Private Sub txtStopSellTime_GotFocus()
    With txtStopSellTime
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
