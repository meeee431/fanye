VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Begin VB.Form frmModifyCompany 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "统计中修改参运公司"
   ClientHeight    =   2160
   ClientLeft      =   3780
   ClientTop       =   4125
   ClientWidth     =   5310
   Icon            =   "frmMordyCompany.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox ckBusID 
      BackColor       =   &H00E0E0E0&
      Caption         =   "模糊车次查询"
      Height          =   195
      Left            =   2940
      TabIndex        =   5
      Top             =   960
      Width           =   1770
   End
   Begin RTComctl3.CoolButton cmdExit 
      Cancel          =   -1  'True
      Height          =   345
      Left            =   3930
      TabIndex        =   3
      Top             =   1725
      Width           =   1110
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
      MICON           =   "frmMordyCompany.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdok 
      Height          =   345
      Left            =   2700
      TabIndex        =   2
      Top             =   1725
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   609
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
      MICON           =   "frmMordyCompany.frx":0028
      PICN            =   "frmMordyCompany.frx":0044
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin FText.asFlatTextBox txtCompanyID 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   1305
      Width           =   3915
      _ExtentX        =   6906
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
      Text            =   ""
      ButtonVisible   =   -1  'True
      Registered      =   -1  'True
      OfficeXPColors  =   -1  'True
   End
   Begin FText.asFlatTextBox txtArea 
      Height          =   285
      Left            =   1230
      TabIndex        =   4
      Top             =   90
      Width           =   1605
      _ExtentX        =   2831
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
      Text            =   ""
      ButtonVisible   =   -1  'True
      Registered      =   -1  'True
      OfficeXPColors  =   -1  'True
   End
   Begin MSComCtl2.DTPicker dtpEndDate 
      Height          =   285
      Left            =   3510
      TabIndex        =   6
      Top             =   495
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   503
      _Version        =   393216
      Format          =   24051712
      CurrentDate     =   36958
   End
   Begin FText.asFlatTextBox txtStationID 
      Height          =   285
      Left            =   3510
      TabIndex        =   7
      Top             =   90
      Width           =   1605
      _ExtentX        =   2831
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
      Text            =   ""
      ButtonVisible   =   -1  'True
      Registered      =   -1  'True
      OfficeXPColors  =   -1  'True
   End
   Begin MSComCtl2.DTPicker dtpStartDate 
      Height          =   285
      Left            =   1200
      TabIndex        =   8
      Top             =   495
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   503
      _Version        =   393216
      Format          =   24051712
      CurrentDate     =   36958
   End
   Begin FText.asFlatTextBox txtBusId 
      Height          =   285
      Left            =   1200
      TabIndex        =   9
      Top             =   915
      Width           =   1605
      _ExtentX        =   2831
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
      Text            =   ""
      ButtonVisible   =   -1  'True
      Registered      =   -1  'True
      OfficeXPColors  =   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "终点站"
      Height          =   180
      Left            =   2940
      TabIndex        =   14
      Top             =   135
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "时间从"
      Height          =   180
      Left            =   135
      TabIndex        =   13
      Top             =   540
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "到"
      Height          =   180
      Left            =   2910
      TabIndex        =   12
      Top             =   570
      Width           =   180
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车次代码"
      Height          =   180
      Left            =   135
      TabIndex        =   11
      Top             =   960
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "地区代码"
      Height          =   180
      Left            =   135
      TabIndex        =   10
      Top             =   135
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "参运公司"
      Height          =   180
      Left            =   135
      TabIndex        =   1
      Top             =   1350
      Width           =   720
   End
End
Attribute VB_Name = "frmModifyCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_szCompanyID As String
Private m_szStationID As String
Private m_szCompanyName As String
Private Sub cmdExit_Click()
  Unload Me
End Sub

Private Sub cmdok_Click()
  Dim szbusID As String
  Dim oCompanyDim As New TicketCompanyDim
  Dim bflg As Boolean
  
  szbusID = LeftAndRight(txtBusID.Text, True, "[")
  If ckBusID.Value = 1 Then
  bflg = True
  End If
  oCompanyDim.Init m_oActiveUser
  oCompanyDim.ModifyCompanyName dtpStartDate.Value, m_szCompanyID, m_szCompanyName, m_szStationID, szbusID, bflg, dtpEndDate.Value
  MsgBox "统计表中修改参运公司成功", vbInformation, "修改参运公司"
  End Sub

Private Sub Form_Load()
  ckBusID.Value = 0
  cmdOk.Enabled = False

    dtpStartDate.Value = DateAdd("d", -1, m_oParam.NowDate)
    dtpEndDate.Value = DateAdd("d", -1, m_oParam.NowDate)
    
End Sub

Private Sub txtArea_ButtonClick()
 Dim szaTemp() As String
  Dim oShell As New CommDialog
   oShell.Init m_oActiveUser
    szaTemp = oShell.SelectArea(False)
    If ArrayLength(szaTemp) = 0 Then Exit Sub
    txtArea.Text = Trim(szaTemp(1, 1)) & "[" & szaTemp(1, 2) & "]"
  
   Set oShell = Nothing
End Sub

Private Sub txtBusId_ButtonClick()
 Dim oBus As New CommDialog
    Dim szaTemp() As String
    oBus.Init m_oActiveUser
    szaTemp = oBus.SelectBus()
    Set oBus = Nothing
    If ArrayLength(szaTemp) = 0 Then Exit Sub
    txtBusID.Text = szaTemp(1, 1)
    IsSave
End Sub

Private Sub txtCompanyID_ButtonClick()
   Dim oShell As New CommDialog
    Dim szaTemp() As String
     oShell.Init m_oActiveUser
    szaTemp = oShell.SelectCompany(False)
    Set oShell = Nothing
    If ArrayLength(szaTemp) = 0 Then Exit Sub
    txtCompanyID.Text = Trim(szaTemp(1, 1)) & "[" & Trim(szaTemp(1, 2)) & "]"
    m_szCompanyID = Trim(szaTemp(1, 1))
    m_szCompanyName = Trim(szaTemp(1, 2))
    Set oShell = Nothing
    IsSave
Exit Sub
End Sub

Private Sub txtStationID_ButtonClick()
' Dim szaTemp() As String
' Dim oShell As New CommDialog
'    oShell.Init m_oActiveUser
'    If txtArea.Text = "" Then
'       szaTemp = oShell.selectStation(, False)
'    Else
'       szaTemp = oShell.selectStation(GetLString(txtArea.Text), False)
'    End If
'    Set oShell = Nothing
'    If ArrayLength(szaTemp) = 0 Then Exit Sub
'    txtStationID.Text = szaTemp(1, 1) & "[" & szaTemp(1, 2) & "]"
'    m_szStationID = szaTemp(1, 1)
'    IsSave
'Exit Sub
End Sub

Public Sub IsSave()
   If txtCompanyID.Text <> "" And txtStationID.Text <> "" And txtBusID.Text <> "" Then
    cmdOk.Enabled = True
   End If
End Sub

