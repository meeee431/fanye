VERSION 5.00
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmWriteOffCheck 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "注销检票"
   ClientHeight    =   4710
   ClientLeft      =   3795
   ClientTop       =   3750
   ClientWidth     =   5985
   HelpContextID   =   4001801
   Icon            =   "frmWriteOffCheck.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox ptTitle 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   7185
      TabIndex        =   24
      Top             =   0
      Width           =   7185
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Height          =   60
         Left            =   0
         TabIndex        =   25
         Top             =   660
         Width           =   7215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ALT+T注销单张检票与注销路单检票切换"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   2670
         TabIndex        =   28
         Top             =   240
         Width           =   3195
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车票号"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1350
         TabIndex        =   27
         Top             =   240
         Width           =   540
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请输入待注销:"
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   1125
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "注意"
      Height          =   1005
      Left            =   510
      TabIndex        =   13
      Top             =   810
      Width           =   5085
      Begin VB.Image Image1 
         Height          =   480
         Left            =   195
         Picture         =   "frmWriteOffCheck.frx":038A
         Top             =   270
         Width           =   480
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmWriteOffCheck.frx":0C54
         Height          =   525
         Left            =   960
         TabIndex        =   14
         Top             =   270
         Width           =   3960
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "检票信息"
      Height          =   1575
      Left            =   510
      TabIndex        =   6
      Top             =   2310
      Width           =   5085
      Begin RTComctl3.CoolButton lblBusCheckInfo 
         Height          =   255
         Left            =   3420
         TabIndex        =   3
         Top             =   1155
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   450
         BTYPE           =   8
         TX              =   "车次检票信息(&C)"
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   12582912
         FCOLO           =   12582912
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmWriteOffCheck.frx":0CDC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin RTComctl3.CoolButton lblTicketSellInfo 
         Height          =   255
         Left            =   1830
         TabIndex        =   2
         Top             =   1155
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   450
         BTYPE           =   8
         TX              =   "车票详细信息(&T)"
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   12582912
         FCOLO           =   12582912
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmWriteOffCheck.frx":0CF8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblBusID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1234"
         Height          =   180
         Left            =   990
         TabIndex        =   22
         Top             =   285
         Width           =   360
      End
      Begin VB.Label lblSerial 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "固定车次"
         Height          =   180
         Left            =   2955
         TabIndex        =   21
         Top             =   285
         Width           =   720
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车次序号:"
         Height          =   180
         Left            =   2130
         TabIndex        =   20
         Top             =   285
         Width           =   810
      End
      Begin VB.Label lblCheckTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "12:01:01"
         Height          =   180
         Left            =   990
         TabIndex        =   19
         Top             =   1215
         Width           =   720
      End
      Begin VB.Label lblCheckor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0123"
         Height          =   180
         Left            =   2955
         TabIndex        =   18
         Top             =   885
         Width           =   360
      End
      Begin VB.Label lblCheckGate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "检票口一"
         Height          =   180
         Left            =   990
         TabIndex        =   17
         Top             =   885
         Width           =   720
      End
      Begin VB.Label lblCheckInMode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "正常/改乘/并班"
         Height          =   180
         Left            =   2955
         TabIndex        =   16
         Top             =   585
         Width           =   1260
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2000-01-01"
         Height          =   180
         Left            =   765
         TabIndex        =   15
         Top             =   585
         Width           =   900
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "检票时间:"
         Height          =   180
         Left            =   180
         TabIndex        =   12
         Top             =   1215
         Width           =   810
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "检票员:"
         Height          =   180
         Left            =   2130
         TabIndex        =   11
         Top             =   885
         Width           =   630
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "检票口:"
         Height          =   180
         Left            =   180
         TabIndex        =   10
         Top             =   885
         Width           =   630
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "日期:"
         Height          =   180
         Left            =   180
         TabIndex        =   9
         Top             =   585
         Width           =   450
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "检票方式:"
         Height          =   180
         Left            =   2130
         TabIndex        =   8
         Top             =   585
         Width           =   810
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "检入车次:"
         Height          =   180
         Left            =   180
         TabIndex        =   7
         Top             =   285
         Width           =   810
      End
   End
   Begin VB.TextBox txtTicketID 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   3075
      MaxLength       =   10
      TabIndex        =   1
      Top             =   1920
      Width           =   2490
   End
   Begin RTComctl3.CoolButton cmdWriteOff 
      Default         =   -1  'True
      Height          =   315
      Left            =   3450
      TabIndex        =   4
      Top             =   4245
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "注销(&W)"
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
      MICON           =   "frmWriteOffCheck.frx":0D14
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
      Height          =   315
      Left            =   4680
      TabIndex        =   5
      Top             =   4245
      Width           =   1065
      _ExtentX        =   1879
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
      MICON           =   "frmWriteOffCheck.frx":0D30
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
      Height          =   3120
      Left            =   -120
      TabIndex        =   23
      Top             =   3990
      Width           =   8745
   End
   Begin VB.Label lblType1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "已检票号(&I):"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   2010
      TabIndex        =   29
      Top             =   1995
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "请输入需要注销的"
      Height          =   195
      Left            =   510
      TabIndex        =   0
      Top             =   1995
      Width           =   1440
   End
End
Attribute VB_Name = "frmWriteOffCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mtOriSheetInfo As TCheckSheetInfo       '原始路单信息

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdWriteOff_Click()
On Error GoTo Err_Done
    Dim nBusSerialNo As Integer
    Dim tCheckTicket As TCheckedTicketInfo
    If Len(txtTicketID) > 0 Then
        '两种，注销检票和删除路单
        If lblType.Caption = "车票号" Then
            tCheckTicket = g_oChkTicket.GetTicketCheckInfo(txtTicketID.Text)
            If Len(tCheckTicket.szBusid) = 0 Then
                MsgboxEx "该车次还未检票,不能注销检票[" & txtTicketID.Text & "]"
                Exit Sub
            End If
            
            lblTicketSellInfo.Enabled = True
            lblBusCheckInfo.Enabled = True
            lblBusID.Caption = tCheckTicket.szBusid
            lblDate.Caption = Format(tCheckTicket.dtBusDate, "YYYY年MM月DD日")
            nBusSerialNo = tCheckTicket.nBusSerialNo
            Select Case tCheckTicket.nCheckTicketType
                Case 1
                    lblCheckInMode.Caption = "正常检入"
                Case 2
                    lblCheckInMode.Caption = "改乘检入"
                Case 3
                    lblCheckInMode.Caption = "并班检入"
            End Select
            lblCheckGate.Caption = Trim(tCheckTicket.szCheckGateName)
            lblCheckor.Caption = Trim(tCheckTicket.szCheckerName)
            lblCheckTime.Caption = Format(tCheckTicket.dtCheckTime, "hh:mm:ss")
            If nBusSerialNo > 0 Then
                lblSerial.Caption = nBusSerialNo
            Else
                lblSerial.Caption = "固定车次"
            End If
            
            If MsgboxEx("注销车票会丢失车票的检票数据，注销吗?", _
                vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                    g_oChkTicket.WriteOffCheckTicket txtTicketID
                    MsgboxEx "注销成功", vbInformation, "信息"
                    cmdWriteOff.Enabled = False
                    
                    If txtTicketID.Enabled Then txtTicketID.SetFocus
            End If
        
        Else
            '判断路单是否存在
            If Not ValidateSheetId Then
                txtTicketID.SetFocus
                Exit Sub
            End If
            
            lblBusCheckInfo.Enabled = True
            
            If MsgboxEx("注销车票会丢失车票的检票数据，注销吗?", _
                vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                
                    Dim moChkTicket As New CheckTicket       '检票对象
                    Dim atTicketsInfo() As TCheckedTicketInfo      '票信息数组
                    Dim i As Integer
                    moChkTicket.Init g_oActiveUser
                    atTicketsInfo = moChkTicket.GetBusCheckTicket(Trim(lblDate.Caption), Trim(lblBusID.Caption), Trim(lblSerial.Caption), Trim(lblCheckGate.Tag))
                    For i = 1 To ArrayLength(atTicketsInfo)
                        g_oChkTicket.WriteOffCheckTicket atTicketsInfo(i).szTicketID
                    Next i
                    
                    MsgboxEx "注销成功", vbInformation, "信息"
                    cmdWriteOff.Enabled = False
                    
                    If txtTicketID.Enabled Then txtTicketID.SetFocus
            End If
        End If
    End If
    Exit Sub
Err_Done:
    ShowErrorMsg
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Chr(KeyCode) = "T" Or Chr(KeyCode) = "T") Then
        If lblType.Caption = "车票号" Then
            lblType.Caption = "路单号"
            lblType1.Caption = "已打路单号(&I):"
        Else
            lblType.Caption = "车票号"
            lblType1.Caption = "已检票号(&I):"
        End If
    End If
End Sub

Private Sub Form_Load()
    AlignFormPos Me
    lblBusID.Caption = ""
    lblDate.Caption = ""
    lblCheckInMode.Caption = ""
    lblCheckor.Caption = ""
    lblCheckGate.Caption = ""
    lblCheckTime.Caption = ""
    lblSerial.Caption = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveFormPos Me
'    Set g_oChkTicket = Nothing
End Sub

Private Sub lblBusCheckInfo_Click()
    If lblBusID.Caption <> "" Then
        Dim oFrmCheckInfo As New frmCheckBusInfo
        Set oFrmCheckInfo.g_oActiveUser = g_oActiveUser
        oFrmCheckInfo.mszBusID = lblBusID
        oFrmCheckInfo.mnBusSerialNo = Val(lblSerial.Caption)
        oFrmCheckInfo.mdtBusDate = lblDate.Caption
        oFrmCheckInfo.Show vbModal
        Set oFrmCheckInfo = Nothing
    End If
End Sub
Private Sub lblTicketSellInfo_Click()
    If Trim(txtTicketID.Text) <> "" Then
        Dim ofrmTicket As frmTicketInfo
        Set ofrmTicket = New frmTicketInfo
        ofrmTicket.TicketID = Trim(txtTicketID.Text)
        ofrmTicket.Show vbModal
        Set ofrmTicket = Nothing
    End If
    
End Sub

Private Sub txtTicketID_Change()
    If Len(txtTicketID.Text) = 0 Then
        cmdWriteOff.Enabled = False
    Else
        cmdWriteOff.Enabled = True
    End If
End Sub

Private Sub txtTicketID_GotFocus()
    txtTicketID.SelStart = 0
    txtTicketID.SelLength = 100
End Sub

Private Function ValidateSheetId() As Boolean
    If txtTicketID.Text <> "" Then
'        If txtOriSheetNo.Text <> lblCurSheetNo.Caption Then
            mtOriSheetInfo = g_oChkTicket.GetCheckSheetInfo(txtTicketID.Text)
            If mtOriSheetInfo.szCheckSheet = "" Then
                MsgboxEx "该路单不存在！", vbCritical, "错误"
                ValidateSheetId = False
            Else
                writeSheetInfo
                ValidateSheetId = True
            End If
        End If
 '   End If
End Function

'填路单信息
Private Sub writeSheetInfo()
    Dim oVehicle As New Vehicle
    Dim oRoute As New Route
    On Error GoTo ErrorHandle
    oVehicle.Init g_oActiveUser
    oVehicle.Identify mtOriSheetInfo.szVehicleId
    oRoute.Init g_oActiveUser
    oRoute.Identify mtOriSheetInfo.szRouteID
    
    lblBusID.Caption = mtOriSheetInfo.szBusid
    lblSerial.Caption = mtOriSheetInfo.nBusSerialNo
    lblCheckor.Caption = mtOriSheetInfo.szMakeSheetUser
    lblDate.Caption = Format(mtOriSheetInfo.dtDate, "YYYY-MM-DD")
    lblCheckGate.Caption = mtOriSheetInfo.szCheckGateName
    lblCheckGate.Tag = mtOriSheetInfo.szCheckGateID
    lblCheckTime.Caption = Format(mtOriSheetInfo.dtStartUpTime, "HH:MM:SS")

    Set oVehicle = Nothing
    Set oRoute = Nothing
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub
