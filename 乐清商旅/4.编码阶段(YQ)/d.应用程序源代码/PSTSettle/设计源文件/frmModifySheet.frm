VERSION 5.00
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Begin VB.Form frmModifySheet 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "修改路单所属车辆"
   ClientHeight    =   4980
   ClientLeft      =   3375
   ClientTop       =   1935
   ClientWidth     =   6630
   Icon            =   "frmModifySheet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   6630
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "注意"
      Height          =   990
      Left            =   525
      TabIndex        =   25
      Top             =   870
      Width           =   5535
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "  修改路单所属车辆及公司将原来检票时，错误的车辆或公司改为正确的，请在路单错误时才使用此功能，正常时请勿使用。"
         Height          =   555
         Left            =   960
         TabIndex        =   26
         Top             =   270
         Width           =   4470
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   225
         Picture         =   "frmModifySheet.frx":000C
         Top             =   285
         Width           =   480
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "路单摘要信息"
      Height          =   1845
      Left            =   525
      TabIndex        =   10
      Top             =   2280
      Width           =   5535
      Begin FText.asFlatTextBox txtVehicle 
         Height          =   285
         Left            =   1050
         TabIndex        =   3
         Top             =   885
         Width           =   1815
         _ExtentX        =   3201
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车次代码:"
         Height          =   180
         Left            =   225
         TabIndex        =   24
         Top             =   360
         Width           =   810
      End
      Begin VB.Label lblBusID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1234"
         Height          =   180
         Left            =   1125
         TabIndex        =   23
         Top             =   360
         Width           =   360
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "日期:"
         Height          =   180
         Left            =   225
         TabIndex        =   22
         Top             =   630
         Width           =   450
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2005-03-25"
         Height          =   180
         Left            =   1125
         TabIndex        =   21
         Top             =   630
         Width           =   900
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车次序号:"
         Height          =   180
         Left            =   2865
         TabIndex        =   20
         Top             =   360
         Width           =   810
      End
      Begin VB.Label lblBusSerial 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "12"
         Height          =   180
         Left            =   3705
         TabIndex        =   19
         Top             =   360
         Width           =   180
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车辆:"
         Height          =   180
         Left            =   225
         TabIndex        =   2
         Top             =   930
         Width           =   450
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "公司:"
         Height          =   180
         Left            =   3225
         TabIndex        =   18
         Top             =   930
         Width           =   450
      End
      Begin VB.Label lblCompany 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "绍兴汽运集团"
         Height          =   180
         Left            =   3705
         TabIndex        =   17
         Top             =   930
         Width           =   1080
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "线路:"
         Height          =   180
         Left            =   3225
         TabIndex        =   16
         Top             =   660
         Width           =   450
      End
      Begin VB.Label lblRoute 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "绍兴富阳线"
         Height          =   180
         Left            =   3705
         TabIndex        =   15
         Top             =   660
         Width           =   900
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "检票员:"
         Height          =   180
         Left            =   225
         TabIndex        =   14
         Top             =   1215
         Width           =   630
      End
      Begin VB.Label lblCheckor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "陈峰"
         Height          =   180
         Left            =   1125
         TabIndex        =   13
         Top             =   1215
         Width           =   360
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "制单时间:"
         Height          =   180
         Left            =   2865
         TabIndex        =   12
         Top             =   1215
         Width           =   810
      End
      Begin VB.Label lblMakeTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2005-03-25 12:30:23"
         Height          =   180
         Left            =   3705
         TabIndex        =   11
         Top             =   1215
         Width           =   1710
      End
   End
   Begin VB.TextBox txtSheetID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1830
      TabIndex        =   1
      Text            =   "0214865"
      Top             =   1950
      Width           =   1410
   End
   Begin VB.PictureBox ptTitle 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   795
      Left            =   15
      ScaleHeight     =   795
      ScaleWidth      =   7185
      TabIndex        =   7
      Top             =   0
      Width           =   7185
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Height          =   60
         Left            =   0
         TabIndex        =   8
         Top             =   750
         Width           =   7215
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请输入需要修改的路单编号:"
         Height          =   180
         Left            =   120
         TabIndex        =   9
         Top             =   300
         Width           =   2250
      End
   End
   Begin RTComctl3.CoolButton cmdPreView 
      Height          =   345
      Left            =   525
      TabIndex        =   6
      Top             =   4500
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "路单预览(&V)"
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
      MICON           =   "frmModifySheet.frx":08D6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdExit 
      Cancel          =   -1  'True
      Height          =   345
      Left            =   4875
      TabIndex        =   5
      Top             =   4500
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   609
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
      MICON           =   "frmModifySheet.frx":08F2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdSave 
      Height          =   345
      Left            =   3345
      TabIndex        =   4
      Top             =   4500
      Width           =   1395
      _ExtentX        =   2461
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
      MICON           =   "frmModifySheet.frx":090E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   " "
      Enabled         =   0   'False
      Height          =   990
      Left            =   -45
      TabIndex        =   27
      Top             =   4230
      Width           =   8745
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "原路单编号(&N):"
      Height          =   180
      Left            =   540
      TabIndex        =   0
      Top             =   1995
      Width           =   1260
   End
End
Attribute VB_Name = "frmModifySheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mtSheetInfo As TCheckSheetInfo       '原始路单信息
Private m_oChkTicket As New CheckTicket

Private Sub InitForm()
    txtSheetID.Text = ""
    
    
    lblBusID.Caption = ""
    lblBusSerial.Caption = ""
    lblDate.Caption = ""
    lblCheckor.Caption = ""
    txtVehicle.Text = ""
    
    lblCompany.Caption = ""
    lblMakeTime.Caption = ""
    lblRoute.Caption = ""

End Sub




Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPreView_Click()
    '显示路单窗口
        Dim oCommDialog As New STShell.CommDialog
        Dim szCheckSheetID As String
        szCheckSheetID = txtSheetID.Text
        If szCheckSheetID <> "" Then
            oCommDialog.Init g_oActiveUser
            oCommDialog.ShowCheckSheet szCheckSheetID
        End If
        
        Set oCommDialog = Nothing
End Sub

Private Sub cmdSave_Click()
    '保存设置
    Dim oSplit As New Split
    On Error GoTo ErrorHandle
    SetBusy
    oSplit.Init g_oActiveUser
    oSplit.ChangeSheetVehicle txtSheetID.Text, ResolveDisplay(txtVehicle.Text)
    ShowMsg "车辆已改为[" & txtVehicle.Text & "]"
    SetNormal
    Exit Sub
ErrorHandle:
    SetNormal
    ShowErrorMsg
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendTab
    End If
End Sub

Private Sub Form_Load()
    InitForm
    m_oChkTicket.Init g_oActiveUser
    
End Sub


Private Sub txtSheetID_LostFocus()

    'InitForm
    
    RefreshCheckSheet
    
End Sub

Private Sub RefreshCheckSheet()
    '刷新路单信息(如车次、公司等)
    
    If txtSheetID.Text <> "" Then
    
        SetBusy
        mtSheetInfo = m_oChkTicket.GetCheckSheetInfo(txtSheetID.Text)
        If mtSheetInfo.szCheckSheet = "" Then
            ShowMsg "该路单不存在！"
        Else
            WriteSheetInfo
        End If
        
        SetNormal
    End If

End Sub



Private Sub WriteSheetInfo()
    Dim oVehicle As New Vehicle
    Dim oRoute As New Route
    oVehicle.Init g_oActiveUser
    oVehicle.Identify mtSheetInfo.szVehicleID
    oRoute.Init g_oActiveUser
    oRoute.Identify mtSheetInfo.szRouteID
    
    lblBusID.Caption = mtSheetInfo.szBusID
    lblBusSerial.Caption = mtSheetInfo.nBusSerialNo
    lblCheckor.Caption = mtSheetInfo.szMakeSheetUser
    lblCompany.Caption = MakeDisplayString(Trim(mtSheetInfo.szCompanyID), Trim(oVehicle.CompanyName))
    lblDate.Caption = Format(mtSheetInfo.dtDate, "YYYY-MM-DD")
    txtVehicle.Text = MakeDisplayString(oVehicle.VehicleID, oVehicle.LicenseTag)
    lblMakeTime.Caption = Format(mtSheetInfo.dtMakeSheetDateTime, "YYYY-MM-DD HH:MM:SS")
    lblRoute.Caption = oRoute.RouteName
    Set oVehicle = Nothing
    Set oRoute = Nothing
End Sub

Private Sub txtVehicle_ButtonClick()

    '显示车辆
    On Error GoTo ErrHandle
    Dim oShell As New STShell.CommDialog
    Dim aszTemp() As String
    Dim oVehicle As New Vehicle
    SetBusy
    oShell.Init g_oActiveUser

    aszTemp = oShell.SelectVehicleEX()
    
    Set oShell = Nothing
    If ArrayLength(aszTemp) = 0 Then Exit Sub
    txtVehicle.Text = MakeDisplayString(Trim(aszTemp(1, 1)), Trim(aszTemp(1, 2)))
    
    oVehicle.Init g_oActiveUser
    oVehicle.Identify mtSheetInfo.szVehicleID
    
    lblCompany.Caption = MakeDisplayString(Trim(mtSheetInfo.szCompanyID), Trim(oVehicle.CompanyName))
    
    SetNormal
    
    Exit Sub
ErrHandle:
    SetNormal
    ShowErrorMsg
    
End Sub

