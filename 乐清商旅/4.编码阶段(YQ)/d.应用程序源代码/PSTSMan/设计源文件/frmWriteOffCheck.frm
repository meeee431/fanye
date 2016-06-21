VERSION 5.00
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Begin VB.Form frmWriteOffCheck 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "注销车票"
   ClientHeight    =   3900
   ClientLeft      =   2220
   ClientTop       =   2085
   ClientWidth     =   6075
   HelpContextID   =   5003801
   Icon            =   "frmWriteOffCheck.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   Begin RTComctl3.CoolButton cmdClose 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   4575
      TabIndex        =   28
      Top             =   3525
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "关闭(&C)__"
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
      MICON           =   "frmWriteOffCheck.frx":014A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdWriteOff 
      Height          =   315
      Left            =   3075
      TabIndex        =   2
      Top             =   3525
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "注销(&W)__"
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
      MICON           =   "frmWriteOffCheck.frx":0166
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "说明"
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   120
      TabIndex        =   9
      Top             =   60
      Width           =   5805
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "      注销车票用于系统调试或特殊情况的处理，注销车票将消除所有相关的车票纪录。正常情况下，请不要使用此功能。"
         Height          =   540
         Left            =   1230
         TabIndex        =   11
         Top             =   495
         Width           =   4440
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "注意：使用前请仔细阅读以下文字"
         Height          =   180
         Left            =   1215
         TabIndex        =   10
         Top             =   240
         Width           =   2700
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "车票信息"
      ForeColor       =   &H80000008&
      Height          =   1755
      Left            =   120
      TabIndex        =   3
      Top             =   1650
      Width           =   5805
      Begin RTComctl3.FloatLabel lblBusInfo 
         Height          =   255
         Left            =   2505
         TabIndex        =   25
         Top             =   210
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NormTextColor   =   -2147483635
         Caption         =   "车次详细信息"
         NormUnderline   =   -1  'True
      End
      Begin VB.Label lblSaler 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label18"
         Height          =   180
         Left            =   1005
         TabIndex        =   27
         Top             =   1260
         Width           =   630
      End
      Begin VB.Label lblBusID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label18"
         Height          =   180
         Left            =   1035
         TabIndex        =   26
         Top             =   255
         Width           =   630
      End
      Begin VB.Label lblSaleTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label17"
         Height          =   180
         Left            =   3420
         TabIndex        =   24
         Top             =   1260
         Width           =   630
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "售票时间:"
         Height          =   180
         Left            =   2565
         TabIndex        =   23
         Top             =   1260
         Width           =   810
      End
      Begin VB.Label lblPrice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label15"
         Height          =   180
         Left            =   1035
         TabIndex        =   22
         Top             =   1005
         Width           =   630
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label14"
         Height          =   180
         Left            =   3450
         TabIndex        =   21
         Top             =   1005
         Width           =   630
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "操作人员:"
         Height          =   180
         Left            =   150
         TabIndex        =   20
         Top             =   1260
         Width           =   810
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车票种类:"
         Height          =   180
         Left            =   2565
         TabIndex        =   19
         Top             =   1005
         Width           =   810
      End
      Begin VB.Label lblSeverUnit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label11"
         Height          =   180
         Left            =   1005
         TabIndex        =   18
         Top             =   510
         Width           =   630
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车票状态:"
         Height          =   180
         Left            =   150
         TabIndex        =   17
         Top             =   1500
         Width           =   810
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "发车日期:"
         Height          =   180
         Left            =   2565
         TabIndex        =   16
         Top             =   510
         Width           =   810
      End
      Begin VB.Label lblStartTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "12:01:01"
         Height          =   180
         Left            =   3420
         TabIndex        =   15
         Top             =   510
         Width           =   720
      End
      Begin VB.Label lblEndStation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0123"
         Height          =   180
         Left            =   3420
         TabIndex        =   14
         Top             =   750
         Width           =   360
      End
      Begin VB.Label lblStartStation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "绍兴"
         Height          =   180
         Left            =   1035
         TabIndex        =   13
         Top             =   750
         Width           =   360
      End
      Begin VB.Label lblTicketStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "正常"
         Height          =   180
         Left            =   1020
         TabIndex        =   12
         Top             =   1500
         Width           =   360
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车票价格:"
         Height          =   180
         Left            =   150
         TabIndex        =   8
         Top             =   1005
         Width           =   810
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "到达车站:"
         Height          =   180
         Left            =   2565
         TabIndex        =   7
         Top             =   750
         Width           =   810
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "起点车站:"
         Height          =   180
         Left            =   150
         TabIndex        =   6
         Top             =   750
         Width           =   810
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "服务单位:"
         Height          =   180
         Left            =   150
         TabIndex        =   5
         Top             =   510
         Width           =   810
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车次代码:"
         Height          =   180
         Left            =   150
         TabIndex        =   4
         Top             =   255
         Width           =   810
      End
   End
   Begin VB.TextBox txtTicketID 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   2775
      TabIndex        =   1
      Top             =   1290
      Width           =   3150
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "请输入需要注销车票票号(&I):"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   1335
      Width           =   2340
   End
End
Attribute VB_Name = "frmWriteOffCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oChkApp As New CommDialog
Dim oClientTicket As New ClientTicket
Dim oSellTicketClient As New SellTicketClient


Private Sub cmdClose_Click()
    Unload Me
End Sub


Private Sub cmdWriteOff_Click()
    If Len(txtTicketID) > 0 Then
        If MsgBox("注销车票会丢失车票所有信息，注销吗?", _
            vbExclamation + vbYesNo + vbDefaultButton2) = vbYes Then
                DoWriteOff
                txtTicketID.SetFocus
        End If
    Else
        MsgBox "票号不能为空!", vbInformation + vbOKOnly, cszMsg
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandle
    oClientTicket.Init g_oActUser
    
    lblBusID.Caption = ""
    lblSaler.Caption = ""
    lblSeverUnit.Caption = ""
    lblTicketStatus = ""
    lblStartStation = ""
    lblEndStation = ""
    lblPrice = ""
    lblType = ""
    lblSaleTime = ""
    lblSaleTime = ""
Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub


Private Sub lblBusInfo_Click()
    Dim dtBusDate As Date
    dtBusDate = CDate(Format(lblStartTime.Caption, "yy-mm-dd"))
    If Len(lblBusID.Caption) > 0 Then
        oChkApp.ShowBusInfo dtBusDate, CStr(lblBusID.Caption)
    End If
    lblBusInfo.NormTextColor = &H8000000D
    
End Sub

Private Sub txtTicketID_GotFocus()
    txtTicketID.SelStart = 0
    txtTicketID.SelLength = Len(txtTicketID.Text)
End Sub

Private Sub txtTicketID_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtTicketID_Validate True    '显示车票信息
        cmdWriteOff.SetFocus
    End If
End Sub


Private Sub txtTicketID_Validate(Cancel As Boolean)
    Dim szTemp As String
    Dim oUnit As New Unit
    Dim oUser As New User
    Dim uStatus As ETicketStatus
    Dim uType As ETicketType
    
    On Error GoTo ErrorHandle
    oUnit.Init g_oActUser
    oUser.Init g_oActUser
    
  
  If Len(txtTicketID.Text) > 0 Then
        oClientTicket.Identify Trim(txtTicketID.Text)
        lblBusID.Caption = oClientTicket.REBusID
        szTemp = oClientTicket.Operator
        oUser.Identify szTemp
        szTemp = szTemp & "[" & oUser.FullName & "]"
        lblSaler.Caption = szTemp
        szTemp = oClientTicket.UnitID
        oUnit.Identify szTemp
        lblSeverUnit.Caption = szTemp & "[" & oUnit.UnitShortName & "]"
        
        
        uStatus = oClientTicket.TicketStatus
        If (uStatus And ST_TicketNormal) <> 0 Then
            szTemp = "正常售出"
        Else
            szTemp = "改签售出"
        End If
        If (uStatus And ST_TicketCanceled) = ST_TicketCanceled Then
            szTemp = szTemp & "\已废"
        End If
        If (uStatus And ST_TicketChanged) = ST_TicketChanged Then
            szTemp = szTemp & "\已改签"
        End If
        If (uStatus And ST_TicketChecked) = ST_TicketChecked Then
            szTemp = szTemp & "\已检"
        End If
        If (uStatus And ST_TicketReturned) = ST_TicketReturned Then
            szTemp = szTemp & "\已退"
        End If
        lblTicketStatus.Caption = szTemp
        
        lblStartStation.Caption = oClientTicket.StartStationID & "[" & oClientTicket.StartStaionName & "]"
        lblEndStation.Caption = oClientTicket.ToStationID & "[" & oClientTicket.ToStationName & "]"
        lblPrice.Caption = oClientTicket.TicketPrice & "(元)"
        
        uType = oClientTicket.TicketType
        Select Case uType
            Case TP_FreeTicket
                lblType.Caption = "免票"
            Case TP_FullPrice
                lblType.Caption = "全票"
            Case TP_HalfPrice
                lblType.Caption = "半票"
        End Select
        
        lblSaleTime.Caption = Format(oClientTicket.SellTime, "YYYY-MM-DD HH:MM:SS")
        lblStartTime.Caption = Format(oClientTicket.REBusDate, "YYYY-MM-DD HH:MM:SS")
        
    End If
Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub DoWriteOff()
    Dim bOnlyLocal As Boolean
    
    On Error GoTo ErrorHandle
    oSellTicketClient.Init g_oActUser
    oSellTicketClient.WriteOffTicket Trim(txtTicketID)
    MsgBox "车票[" & Trim(txtTicketID) & "]注销成功!", vbOKOnly + vbInformation, cszMsg
hereback:
    If bOnlyLocal = True Then
        oSellTicketClient.WriteOffTicket Trim(txtTicketID), True
        MsgBox "车票[" & Trim(txtTicketID) & "]只注销本地成功!", vbOKOnly + vbInformation, cszMsg
    End If
Exit Sub
ErrorHandle:
    If err.Number = 11623 Then
        If MsgBox("此车票是代售车票,注销时远程连接失败!" & vbCrLf & _
            "是否只注销本地信息?", vbQuestion + vbYesNo + vbDefaultButton2, cszMsg) = vbYes Then
            bOnlyLocal = True
            GoTo hereback
        End If
    Else
        ShowErrorMsg
    End If
End Sub
