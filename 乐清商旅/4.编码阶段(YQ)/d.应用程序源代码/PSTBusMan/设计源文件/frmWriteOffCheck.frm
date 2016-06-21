VERSION 5.00
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmWriteOffCheck 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "注销检票"
   ClientHeight    =   4725
   ClientLeft      =   3225
   ClientTop       =   2130
   ClientWidth     =   6120
   HelpContextID   =   4001801
   Icon            =   "frmWriteOffCheck.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtTicketID 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   3195
      TabIndex        =   22
      Top             =   1920
      Width           =   2490
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "检票信息"
      Height          =   1575
      Left            =   630
      TabIndex        =   5
      Top             =   2310
      Width           =   5085
      Begin RTComctl3.CoolButton lblBusCheckInfo 
         Height          =   255
         Left            =   3420
         TabIndex        =   6
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
         MICON           =   "frmWriteOffCheck.frx":000C
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
         TabIndex        =   7
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
         MICON           =   "frmWriteOffCheck.frx":0028
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "检入车次:"
         Height          =   180
         Left            =   180
         TabIndex        =   21
         Top             =   285
         Width           =   810
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "检票方式:"
         Height          =   180
         Left            =   2130
         TabIndex        =   20
         Top             =   585
         Width           =   810
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "日期:"
         Height          =   180
         Left            =   180
         TabIndex        =   19
         Top             =   585
         Width           =   450
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "检票口:"
         Height          =   180
         Left            =   180
         TabIndex        =   18
         Top             =   885
         Width           =   630
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "检票员:"
         Height          =   180
         Left            =   2130
         TabIndex        =   17
         Top             =   885
         Width           =   630
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "检票时间:"
         Height          =   180
         Left            =   180
         TabIndex        =   16
         Top             =   1215
         Width           =   810
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2000-01-01"
         Height          =   180
         Left            =   990
         TabIndex        =   15
         Top             =   585
         Width           =   900
      End
      Begin VB.Label lblCheckInMode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "正常/改乘/并班"
         Height          =   180
         Left            =   2955
         TabIndex        =   14
         Top             =   585
         Width           =   1260
      End
      Begin VB.Label lblCheckGate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "检票口一"
         Height          =   180
         Left            =   990
         TabIndex        =   13
         Top             =   885
         Width           =   720
      End
      Begin VB.Label lblCheckor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0123"
         Height          =   180
         Left            =   2955
         TabIndex        =   12
         Top             =   885
         Width           =   360
      End
      Begin VB.Label lblCheckTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "12:01:01"
         Height          =   180
         Left            =   990
         TabIndex        =   11
         Top             =   1215
         Width           =   720
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车次序号:"
         Height          =   180
         Left            =   2130
         TabIndex        =   10
         Top             =   285
         Width           =   810
      End
      Begin VB.Label lblSerial 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "固定车次"
         Height          =   180
         Left            =   2955
         TabIndex        =   9
         Top             =   285
         Width           =   720
      End
      Begin VB.Label lblBusID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1234"
         Height          =   180
         Left            =   990
         TabIndex        =   8
         Top             =   285
         Width           =   360
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "注意"
      Height          =   1005
      Left            =   630
      TabIndex        =   3
      Top             =   810
      Width           =   5085
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmWriteOffCheck.frx":0044
         Height          =   525
         Left            =   960
         TabIndex        =   4
         Top             =   270
         Width           =   3960
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   195
         Picture         =   "frmWriteOffCheck.frx":00CC
         Top             =   270
         Width           =   480
      End
   End
   Begin VB.PictureBox ptTitle 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   7185
      TabIndex        =   0
      Top             =   0
      Width           =   7185
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Height          =   60
         Left            =   0
         TabIndex        =   1
         Top             =   660
         Width           =   7215
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请输入待注销车票号:"
         Height          =   180
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1710
      End
   End
   Begin RTComctl3.CoolButton cmdWriteOff 
      Default         =   -1  'True
      Height          =   315
      Left            =   3570
      TabIndex        =   23
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
      MICON           =   "frmWriteOffCheck.frx":0996
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
      Left            =   4800
      TabIndex        =   24
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
      MICON           =   "frmWriteOffCheck.frx":09B2
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
      TabIndex        =   25
      Top             =   3990
      Width           =   8745
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "请输入需要注销的已检票号(&I):"
      Height          =   180
      Left            =   630
      TabIndex        =   26
      Top             =   1995
      Width           =   2520
   End
End
Attribute VB_Name = "frmWriteOffCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oChkTk As New CheckTicket

Dim nBusSerialNo As Integer

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdWriteOff_Click()
On Error GoTo Err_Done
    If Len(txtTicketID) > 0 Then
        ShowTicket
        If MsgBox("注销车票会丢失车票的检票数据，注销吗?", _
            vbQuestion + vbYesNoCancel + vbDefaultButton2) = vbYes Then

                oChkTk.WriteOffCheckTicket txtTicketID
                MsgBox "注销成功", vbInformation, "信息"
                cmdWriteOff.Enabled = False
        End If
    Else
        MsgBox "票号不能为空", vbExclamation, "警告"
    End If
    Exit Sub
Err_Done:
    MsgBox err.Description, vbExclamation, "错误 -- " & err.Number
End Sub

Private Sub Form_Load()

    oChkTk.Init g_oActiveUser


    lblBusID.Caption = ""
    lblDate.Caption = ""
    lblCheckInMode.Caption = ""
    lblCheckor.Caption = ""
    lblCheckGate.Caption = ""
    lblCheckTime.Caption = ""
    lblSerial.Caption = ""
End Sub

Private Sub lblBusCheckInfo_Click()
    Dim oChkApp As New CommDialog
    oChkApp.Init g_oActiveUser
    oChkApp.ShowCheckInfo lblDate.Caption, lblBusID.Caption, nBusSerialNo
    Set oChkApp = Nothing

End Sub

Private Sub lblBusId_Click()

'    If Len(lblBusID.Caption) > 0 Then
'        oChkApp.ShowBusInfo m_oActiveUser, lblDate.Caption, Trim(lblBusID.Caption)
'    End If
'    lblBusID.NormTextColor = &H8000000D

End Sub

Private Sub lblTicketSellInfo_Click()
    Dim oChkApp As New CommDialog
    oChkApp.Init g_oActiveUser
    oChkApp.ShowTicketInfo txtTicketID.Text
    Set oChkApp = Nothing
End Sub

Private Sub txtTicketID_Change()
    If Len(txtTicketID.Text) = 0 Then
        cmdWriteOff.Enabled = False
    Else
        cmdWriteOff.Enabled = True
    End If
End Sub

Private Sub ShowTicket()
    Dim tCheckTicket As TCheckedTicketInfo
    If Len(txtTicketID.Text) > 0 Then
        tCheckTicket = oChkTk.GetTicketCheckInfo(txtTicketID.Text)
        If Len(tCheckTicket.szbusID) = 0 Then
            err.Raise 23302, , "该车次还未检票,不能注销检票[" & txtTicketID.Text & "]"
        End If
        lblTicketSellInfo.Enabled = True
        lblBusCheckInfo.Enabled = True
        lblBusID.Caption = tCheckTicket.szbusID
        lblDate.Caption = Format(tCheckTicket.dtBusDate, "YYYY年MM月DD日")
        nBusSerialNo = tCheckTicket.nBusSerialNo
        Select Case tCheckTicket.nCheckTicketType
            Case ECheckStatus.NormalTicket
                lblCheckInMode.Caption = "正常检入"
            Case ECheckStatus.ChangeTicket
                lblCheckInMode.Caption = "改乘检入"
            Case ECheckStatus.MergeTicket
                lblCheckInMode.Caption = "并班检入"
        End Select
        lblCheckGate.Caption = Trim(tCheckTicket.szCheckGateName)
        lblCheckor.Caption = Trim(tCheckTicket.szCheckerID)
        lblCheckTime.Caption = Format(tCheckTicket.dtCheckTime, cszTimeStr)
        If nBusSerialNo > 0 Then
            lblSerial.Caption = nBusSerialNo
        Else
            lblSerial.Caption = "固定车次"
        End If
    End If
End Sub
