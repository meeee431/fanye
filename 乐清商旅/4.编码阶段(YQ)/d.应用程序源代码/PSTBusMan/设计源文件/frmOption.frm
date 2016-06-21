VERSION 5.00
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Begin VB.Form frmOption 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "设定"
   ClientHeight    =   2865
   ClientLeft      =   3750
   ClientTop       =   3195
   ClientWidth     =   4890
   HelpContextID   =   2004001
   Icon            =   "frmOption.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   4890
   StartUpPosition =   1  '所有者中心
   Begin RTComctl3.CoolButton cmdClose 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   3450
      TabIndex        =   8
      Top             =   2400
      Width           =   1095
      _ExtentX        =   0
      _ExtentY        =   0
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
      MICON           =   "frmOption.frx":038A
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
      Height          =   315
      Left            =   2280
      TabIndex        =   9
      Top             =   2400
      Width           =   1095
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
      MICON           =   "frmOption.frx":03A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.CheckBox chkBusStionIDCanSale 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "查询可售途经站"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   120
      TabIndex        =   3
      Top             =   1710
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "车牌前缀设定"
      Height          =   705
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   4605
      Begin VB.TextBox txtFChar 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   150
         TabIndex        =   2
         Text            =   "浙"
         Top             =   270
         Width           =   1260
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车辆查询时系统自动设定的车牌前缀"
         Height          =   180
         Left            =   1575
         TabIndex        =   6
         Top             =   315
         Width           =   2880
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "全额退票设定"
      Height          =   660
      Left            =   120
      TabIndex        =   4
      Top             =   90
      Width           =   4605
      Begin VB.CheckBox chkAll 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "全额退票(&A)"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   150
         TabIndex        =   0
         Top             =   315
         Value           =   2  'Grayed
         Width           =   1410
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "停班车次自动设定为全额退票"
         Height          =   240
         Left            =   1575
         TabIndex        =   5
         Top             =   330
         Width           =   2385
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "RTStation"
      Enabled         =   0   'False
      Height          =   1530
      Left            =   -105
      TabIndex        =   7
      Top             =   2100
      Width           =   5385
   End
End
Attribute VB_Name = "frmOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
     
Private Sub SaveSystemOption()
    Dim oReg As New CFreeReg
    Dim szTemp As String
    oReg.Init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany
    Select Case chkAll.Value
           Case 0: oReg.SaveSetting "Scheme\Option", "StopAllRefundment", "否":    g_bStopAllRefundment = False
           Case Else: oReg.SaveSetting "Scheme\Option", "StopAllRefundment", "是": g_bStopAllRefundment = True
    End Select

     oReg.SaveSetting "Scheme\Option", "LicenseForce", txtFChar.Text & "[]": g_szLicenseForce = txtFChar.Text
'     oReg.SaveSetting "Scheme\Option", "ExeFilePosition", txtFile.Text & "[]": g_szExeFilePosition = txtFile.Text
     
End Sub

Private Sub chkAll_Click()
    IsSave
End Sub

Private Sub chkBusStionIDCanSale_Click()
cmdOK.Enabled = True
End Sub

Private Sub CmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    SaveSystemOption
    cmdOK.Enabled = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
       Case vbKeyEscape
       Unload Me
End Select
End Sub

Private Sub Form_Load()
On Error Resume Next
    AlignFormPos Me
    LoadSystemOption
    txtFChar.Text = g_szLicenseForce
    If g_bStopAllRefundment = True Then
        chkAll.Value = 2
    Else
        chkAll.Value = 0
    End If
    cmdOK.Enabled = False
End Sub

Private Sub IsSave()
    cmdOK.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveFormPos Me
End Sub

Private Sub txtFChar_Change()
    IsSave
End Sub

Private Sub txtFile_Change()
    IsSave
End Sub

Public Sub LoadSystemOption()
    Dim oReg As New CFreeReg
    Dim szTemp As String
    Dim i As Integer, nCount As Integer
    Dim aszTemp As Variant
    On Error GoTo ErrorHandle
    oReg.Init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany
    szTemp = oReg.GetSetting("Scheme\Option", "StopAllRefundment", "是")
'    g_szExeFilePosition = GetLString(oReg.GetSetting("Scheme\Option", "ExeFilePosition", App.Path))
    If szTemp = "否" Then
        g_bStopAllRefundment = False
    Else
        g_bStopAllRefundment = True
    End If
    szTemp = GetLString(oReg.GetSetting("Scheme\Option", "LicenseForce", "浙"))
    g_szLicenseForce = szTemp
    ' sztemp = oReg.GetSetting("SNScheme\Option", "BusStationCanSale")
    ' If sztemp = "否" Then
    '     g_bFlgBusStationIdCanSale = False
    ' Else
    '    g_bFlgBusStationIdCanSale = True
    ' End If
Exit Sub
ErrorHandle:
    If err.Number = 500 Then
        oReg.SaveSetting "Scheme\Option", "StopAllRefundment", "是"
        oReg.SaveSetting "Scheme\Option", "LicenseForce", "浙"
'        oReg.SaveSetting "Scheme\Option", "ExeFilePosition", "C:\Program Files\RTSoft\RTStation\PSTPriMan\PSTPriMan.exe[]"
        'oReg.SaveSetting "SNScheme\Option", "BusStationCanSale", "是"
        Resume Next
    End If
    ShowErrorMsg
End Sub


