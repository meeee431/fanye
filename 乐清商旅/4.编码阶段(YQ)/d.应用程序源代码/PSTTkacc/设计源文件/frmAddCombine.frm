VERSION 5.00
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmAddCombine 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "车次公司组合"
   ClientHeight    =   3660
   ClientLeft      =   4545
   ClientTop       =   3090
   ClientWidth     =   5070
   Icon            =   "frmAddCombine.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   5070
   StartUpPosition =   1  '所有者中心
   Begin RTComctl3.CoolButton CoolButton1 
      Height          =   345
      Left            =   570
      TabIndex        =   16
      Top             =   3210
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "帮助"
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
      MICON           =   "frmAddCombine.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin FText.asFlatSpinEdit txtCombineSerial 
      Height          =   300
      Left            =   1920
      TabIndex        =   1
      Top             =   900
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   529
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
      Text            =   "0"
      ButtonBackColor =   -2147483633
   End
   Begin FText.asFlatTextBox txtCompanyID 
      Height          =   300
      Left            =   1920
      TabIndex        =   7
      Top             =   2100
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   529
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
      OfficeXPColors  =   -1  'True
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1920
      TabIndex        =   9
      Top             =   2520
      Width           =   2535
   End
   Begin RTComctl3.CoolButton cmdOk 
      Default         =   -1  'True
      Height          =   345
      Left            =   2205
      TabIndex        =   10
      Top             =   3210
      Width           =   1215
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "保存(&S)"
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
      MICON           =   "frmAddCombine.frx":0028
      PICN            =   "frmAddCombine.frx":0044
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
      Height          =   345
      Left            =   3630
      TabIndex        =   11
      Top             =   3210
      Width           =   1215
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "关闭(&C)"
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
      MICON           =   "frmAddCombine.frx":03DE
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
      Height          =   3120
      Left            =   -150
      TabIndex        =   15
      Top             =   2970
      Width           =   8745
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   -30
      TabIndex        =   14
      Top             =   660
      Width           =   6885
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   6615
      TabIndex        =   12
      Top             =   0
      Width           =   6615
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请修改或新增车次公司组合:"
         Height          =   180
         Left            =   270
         TabIndex        =   13
         Top             =   270
         Width           =   2250
      End
   End
   Begin VB.TextBox txtStartBusID 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1920
      MaxLength       =   5
      TabIndex        =   3
      Top             =   1305
      Width           =   2535
   End
   Begin VB.TextBox txtEndBusID 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1920
      MaxLength       =   5
      TabIndex        =   5
      Top             =   1695
      Width           =   2535
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "组合说明(&N):"
      Height          =   180
      Left            =   600
      TabIndex        =   8
      Top             =   2565
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "参运公司(&D):"
      Height          =   180
      Left            =   600
      TabIndex        =   6
      Top             =   2130
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "起始车次(&B):"
      Height          =   180
      Left            =   600
      TabIndex        =   2
      Top             =   1350
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "组合序号(&C):"
      Height          =   180
      Left            =   600
      TabIndex        =   0
      Top             =   960
      Width           =   1080
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "结束车次(&E):"
      Height          =   180
      Left            =   600
      TabIndex        =   4
      Top             =   1770
      Width           =   1080
   End
End
Attribute VB_Name = "frmAddCombine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Status As EFormStatus
Public m_nCombineSerial As Integer
Public m_szStartBusID As String
Public m_szEndBusID As String
Public m_szCompanyID As String
Public m_szCompanyName As String
Private m_oTkAcc As New TicketCompanyDim

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
    On Error GoTo Here
    If Status = EFormStatus.SNAddNew Then
        m_oTkAcc.AddBusCompanyCombine txtCombineSerial.Value, LeftAndRight(txtCompanyID.Text, False, "["), txtStartBusID.Text, txtEndBusID.Text, LeftAndRight(txtCompanyID.Text, True, "[")
        txtStartBusID.Text = ""
        txtEndBusID.Text = ""
        txtCompanyID.Text = ""
        cmdOk.Enabled = False
        frmCombine.FillCombine
        txtStartBusID.SetFocus
        
        
    ElseIf Status = EFormStatus.SNModify Then
        
    End If
    Exit Sub
Here:
    ShowErrorMsg
End Sub

Private Sub CoolButton1_Click()
DisplayHelp Me
End Sub

Private Sub Form_Load()
    On Error GoTo Here
    m_oTkAcc.Init m_oActiveUser
    
    Select Case Status
    Case EFormStatus.SNAddNew
       
        cmdOk.Caption = "新增(&A)"
        
    Case EFormStatus.SNModify
        
        cmdOk.Caption = "修改(&S)"
        txtCombineSerial.Value = m_nCombineSerial
        txtStartBusID.Text = m_szStartBusID
        txtEndBusID.Text = m_szEndBusID
        txtCompanyID.Text = MakeDisplayString(m_szCompanyID, m_szCompanyName)
        
    End Select
    cmdOk.Enabled = False
    Exit Sub
Here:
    ShowErrorMsg
    
End Sub


Private Sub txtCombineSerial_Change()
    IsSave
    
End Sub

Private Sub IsSave()
    If txtStartBusID.Text <> "" And txtEndBusID.Text <> "" _
        And txtCompanyID.Text <> "" Then
        cmdOk.Enabled = True
    End If
End Sub

Private Sub txtCompanyID_ButtonClick()
    Dim aszCompany() As String
    On Error GoTo Here
    aszCompany = m_oShell.SelectCompany
    If ArrayLength(aszCompany) > 0 Then txtCompanyID.Text = MakeDisplayString(aszCompany(1, 1), aszCompany(1, 2))
    Exit Sub
Here:
    ShowErrorMsg
End Sub

Private Sub txtCompanyID_Change()
    IsSave
End Sub

Private Sub txtEndBusID_Change()
    IsSave
End Sub

Private Sub txtEndBusID_GotFocus()
    txtEndBusID.SelStart = 0
    txtEndBusID.SelLength = 100
    
End Sub

Private Sub txtStartBusID_Change()
    IsSave
End Sub

Private Sub txtStartBusID_GotFocus()
    txtStartBusID.SelStart = 0
    txtStartBusID.SelLength = 100
    
End Sub
