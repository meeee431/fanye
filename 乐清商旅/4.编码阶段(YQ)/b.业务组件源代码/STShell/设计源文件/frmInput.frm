VERSION 5.00
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.0#0"; "RTComctl3.ocx"
Begin VB.Form frmInput 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "输入对话框"
   ClientHeight    =   1485
   ClientLeft      =   4380
   ClientTop       =   3975
   ClientWidth     =   4785
   Icon            =   "frmInput.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtInput 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   930
      TabIndex        =   1
      Top             =   600
      Width           =   3525
   End
   Begin RTComctl3.CoolButton cmdOk 
      Default         =   -1  'True
      Height          =   315
      Left            =   2280
      TabIndex        =   2
      Top             =   1020
      Width           =   1020
      _ExtentX        =   1799
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
      MICON           =   "frmInput.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   3435
      TabIndex        =   3
      Top             =   1020
      Width           =   1020
      _ExtentX        =   1799
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
      MICON           =   "frmInput.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Image imgInformation 
      Height          =   645
      Left            =   150
      Top             =   180
      Width           =   660
   End
   Begin VB.Label lblMsg 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "请输入所要求的值(&I)："
      Height          =   180
      Left            =   930
      TabIndex        =   0
      Top             =   330
      Width           =   1890
   End
End
Attribute VB_Name = "frmInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mbNumberOnly As Boolean
Public mszResult As String

Private Sub cmdCancel_Click()
    mszResult = ""
    Unload Me
End Sub

Private Sub cmdOk_Click()
    mszResult = txtInput.Text
    Unload Me
End Sub



Private Sub Form_Load()
    Set imgInformation.Picture = LoadResPicture(101, 0)
End Sub

Private Sub txtInput_Change()
    If mbNumberOnly Then
        FormatTextToNumeric txtInput, True, True
    End If
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
'    If mbNumberOnly Then
'        If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46 Or KeyAscii = vbKeyBack Then
'            If KeyAscii = 46 Then
'                If InStr(1, txtInput.Text, ".") > 0 And InStr(1, txtInput.SelText, ".") = 0 Then KeyAscii = 0
'            End If
'        Else
'            KeyAscii = 0
'        End If
'    End If
End Sub
