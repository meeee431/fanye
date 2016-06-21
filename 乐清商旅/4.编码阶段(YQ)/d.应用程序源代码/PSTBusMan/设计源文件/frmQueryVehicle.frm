VERSION 5.00
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.0#0"; "RTComctl3.ocx"
Begin VB.Form frmQueryVehicle 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "查询车辆"
   ClientHeight    =   2070
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4725
   Icon            =   "frmQueryVehicle.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   4725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin FText.asFlatTextBox txtCompany 
      Height          =   285
      Left            =   1020
      TabIndex        =   5
      Top             =   884
      Width           =   2100
      _ExtentX        =   3704
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
      Registered      =   -1  'True
   End
   Begin VB.TextBox txtVehicle 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1020
      TabIndex        =   1
      Top             =   150
      Width           =   2100
   End
   Begin VB.TextBox txtLicense 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1020
      TabIndex        =   3
      Top             =   517
      Width           =   2100
   End
   Begin RTComctl3.CoolButton CancelButton 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   3390
      TabIndex        =   11
      Top             =   600
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
      MICON           =   "frmQueryVehicle.frx":038A
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
      Left            =   3390
      TabIndex        =   10
      Top             =   180
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
      MICON           =   "frmQueryVehicle.frx":03A6
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
      Left            =   1020
      TabIndex        =   7
      Top             =   1275
      Width           =   2100
      _ExtentX        =   3704
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
      Registered      =   -1  'True
   End
   Begin FText.asFlatTextBox txtVehicleType 
      Height          =   285
      Left            =   1020
      TabIndex        =   9
      Top             =   1650
      Width           =   2100
      _ExtentX        =   3704
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
      Registered      =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车型(&T):"
      Height          =   180
      Left            =   270
      TabIndex        =   8
      Top             =   1695
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "公司(&Z):"
      Height          =   180
      Left            =   270
      TabIndex        =   4
      Top             =   952
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车主(&W):"
      Height          =   180
      Left            =   270
      TabIndex        =   6
      Top             =   1323
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "代码(&N):"
      Height          =   180
      Left            =   270
      TabIndex        =   0
      Top             =   210
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车牌(&P):"
      Height          =   180
      Left            =   270
      TabIndex        =   2
      Top             =   581
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
End Sub

Private Sub Form_Load()
    IsCancel = True
End Sub

Private Sub OKButton_Click()
    IsCancel = False
    Me.Hide
End Sub


Private Sub txtBusOwner_ButtonClick()
On Error GoTo ErrHandle
    Dim oShell As New CommDialog
    Dim aszTmp() As String
    oShell.Init g_oActiveUser
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
    oShell.Init g_oActiveUser
    aszTmp = oShell.SelectCompany
    Set oShell = Nothing
    If ArrayLength(aszTmp) = 0 Then Exit Sub
    txtCompany.Text = MakeDisplayString(aszTmp(1, 1), aszTmp(1, 2))
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub


Private Sub txtVehicleType_ButtonClick()
On Error GoTo ErrHandle
    Dim oShell As New CommDialog
    Dim aszTmp() As String
    oShell.Init g_oActiveUser
    aszTmp = oShell.SelectVehicleType
    Set oShell = Nothing
    If ArrayLength(aszTmp) = 0 Then Exit Sub
    txtVehicleType.Text = MakeDisplayString(aszTmp(1, 1), aszTmp(1, 2))
    Exit Sub
ErrHandle:
    ShowErrorMsg

End Sub
