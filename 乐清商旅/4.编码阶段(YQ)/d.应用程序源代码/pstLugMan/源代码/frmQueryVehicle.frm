VERSION 5.00
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmQueryVehicle 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "查询车辆"
   ClientHeight    =   3225
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6600
   Icon            =   "frmQueryVehicle.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox cboAcceptType 
      Appearance      =   0  'Flat
      Height          =   300
      ItemData        =   "frmQueryVehicle.frx":038A
      Left            =   4440
      List            =   "frmQueryVehicle.frx":038C
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   960
      Width           =   1740
   End
   Begin VB.PictureBox ptTitle 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   7215
      TabIndex        =   13
      Top             =   0
      Width           =   7215
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   60
         Left            =   0
         TabIndex        =   14
         Top             =   660
         Width           =   7245
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请输入查询车辆的条件"
         Height          =   180
         Left            =   480
         TabIndex        =   15
         Top             =   240
         Width           =   1800
      End
   End
   Begin FText.asFlatTextBox txtCompany 
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Top             =   1440
      Width           =   1740
      _ExtentX        =   3069
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
      Text            =   ""
      ButtonBackColor =   -2147483633
      ButtonVisible   =   -1  'True
   End
   Begin VB.TextBox txtVehicle 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   968
      Width           =   1740
   End
   Begin VB.TextBox txtLicense 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Top             =   1920
      Width           =   1740
   End
   Begin RTComctl3.CoolButton CancelButton 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   5160
      TabIndex        =   10
      Top             =   2640
      Width           =   1065
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
      MICON           =   "frmQueryVehicle.frx":038E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton OKButton 
      Default         =   -1  'True
      Height          =   315
      Left            =   3840
      TabIndex        =   9
      Top             =   2640
      Width           =   1065
      _ExtentX        =   0
      _ExtentY        =   0
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
      MICON           =   "frmQueryVehicle.frx":03AA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin FText.asFlatTextBox txtBusOwner 
      Height          =   285
      Left            =   4440
      TabIndex        =   7
      Top             =   1920
      Width           =   1740
      _ExtentX        =   3069
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
      Text            =   ""
      ButtonBackColor =   -2147483633
      ButtonVisible   =   -1  'True
   End
   Begin FText.asFlatTextBox txtSplitCompany 
      Height          =   285
      Left            =   4440
      TabIndex        =   11
      Top             =   1440
      Width           =   1740
      _ExtentX        =   3069
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
      Text            =   ""
      ButtonBackColor =   -2147483633
      ButtonVisible   =   -1  'True
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   " "
      Enabled         =   0   'False
      Height          =   960
      Left            =   -120
      TabIndex        =   17
      Top             =   2400
      Width           =   7425
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "拆帐公司(&S):"
      Height          =   180
      Left            =   3240
      TabIndex        =   12
      Top             =   1492
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "托运方式 (&T):"
      Height          =   180
      Left            =   3240
      TabIndex        =   8
      Top             =   1020
      Width           =   1170
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "参运公司(&Z):"
      Height          =   180
      Left            =   270
      TabIndex        =   4
      Top             =   1492
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车主(&W):"
      Height          =   180
      Left            =   3240
      TabIndex        =   6
      Top             =   1972
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "代码(&N):"
      Height          =   180
      Left            =   270
      TabIndex        =   0
      Top             =   1020
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车牌(&P):"
      Height          =   180
      Left            =   240
      TabIndex        =   2
      Top             =   1972
      Width           =   720
   End
End
Attribute VB_Name = "frmQueryVehicle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public IsCancel As Boolean  '是否取消查询

Private Sub CancelButton_Click()
    Unload Me
    IsCancel = True
End Sub

Private Sub Form_Load()
    IsCancel = True
      With cboAcceptType
     .AddItem ""
     .AddItem szAcceptTypeGeneral
     .AddItem szAcceptTypeMan
     .ListIndex = 0
      End With
    txtVehicle.Text = ""
    txtCompany.Text = ""
    txtSplitCompany.Text = ""
    txtBusOwner.Text = ""
    txtLicense.Text = ""
End Sub

Private Sub OKButton_Click()
    IsCancel = False
    Me.Hide
End Sub


Private Sub txtBusOwner_ButtonClick()
On Error GoTo ErrHandle
    Dim oShell As New CommDialog
    Dim aszTmp() As String
    oShell.Init m_oAUser
    aszTmp = oShell.SelectOwner(ResolveDisplay(Trim(txtCompany.Text)))
    Set oShell = Nothing
    If ArrayLength(aszTmp) = 0 Then Exit Sub
    txtBusOwner.Text = MakeDisplayString(aszTmp(1, 1), aszTmp(1, 2))
    Exit Sub
ErrHandle:
    ShowErrorMsg

End Sub

Private Sub txtCompany_ButtonClick()
On Error GoTo ErrHandle
    Dim oShell As New CommDialog
    Dim aszTmp() As String
    oShell.Init m_oAUser
    aszTmp = oShell.SelectCompany
    Set oShell = Nothing
    If ArrayLength(aszTmp) = 0 Then Exit Sub
    txtCompany.Text = MakeDisplayString(aszTmp(1, 1), aszTmp(1, 2))
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub


Private Sub txtVehicleType_ButtonClick()
'On Error GoTo ErrHandle
'    Dim oShell As New CommDialog
'    Dim aszTmp() As String
'    oShell.Init m_oAUser
'    aszTmp = oShell.SelectVehicleType
'    Set oShell = Nothing
'    If ArrayLength(aszTmp) = 0 Then Exit Sub
'    txtVehicleType.Text = MakeDisplayString(aszTmp(1, 1), aszTmp(1, 2))
'    Exit Sub
'ErrHandle:
'    ShowErrorMsg

End Sub

Private Sub txtSplitCompany_ButtonClick()
On Error GoTo ErrHandle
    Dim oShell As New STShell.CommDialog
    Dim aszTemp() As String
    oShell.Init m_oAUser
    aszTemp = oShell.SelectCompany()
    Set oShell = Nothing
    If ArrayLength(aszTemp) = 0 Then Exit Sub
     txtSplitCompany.Text = MakeDisplayString(aszTemp(1, 1), aszTemp(1, 2))

    
Exit Sub
ErrHandle:
ShowErrorMsg
End Sub

