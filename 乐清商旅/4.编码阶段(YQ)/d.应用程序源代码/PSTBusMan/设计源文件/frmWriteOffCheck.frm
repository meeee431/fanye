VERSION 5.00
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmWriteOffCheck 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ע����Ʊ"
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
   StartUpPosition =   1  '����������
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
      Caption         =   "��Ʊ��Ϣ"
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
         TX              =   "���μ�Ʊ��Ϣ(&C)"
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
         TX              =   "��Ʊ��ϸ��Ϣ(&T)"
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
         Caption         =   "���복��:"
         Height          =   180
         Left            =   180
         TabIndex        =   21
         Top             =   285
         Width           =   810
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ʊ��ʽ:"
         Height          =   180
         Left            =   2130
         TabIndex        =   20
         Top             =   585
         Width           =   810
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����:"
         Height          =   180
         Left            =   180
         TabIndex        =   19
         Top             =   585
         Width           =   450
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ʊ��:"
         Height          =   180
         Left            =   180
         TabIndex        =   18
         Top             =   885
         Width           =   630
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ƱԱ:"
         Height          =   180
         Left            =   2130
         TabIndex        =   17
         Top             =   885
         Width           =   630
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ʊʱ��:"
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
         Caption         =   "����/�ĳ�/����"
         Height          =   180
         Left            =   2955
         TabIndex        =   14
         Top             =   585
         Width           =   1260
      End
      Begin VB.Label lblCheckGate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ʊ��һ"
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
         Caption         =   "�������:"
         Height          =   180
         Left            =   2130
         TabIndex        =   10
         Top             =   285
         Width           =   810
      End
      Begin VB.Label lblSerial 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�̶�����"
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
      Caption         =   "ע��"
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
         Caption         =   "�������ע����Ʊ��:"
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
      TX              =   "ע��(&W)"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      TX              =   "ȡ��"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      Caption         =   "��������Ҫע�����Ѽ�Ʊ��(&I):"
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
        If MsgBox("ע����Ʊ�ᶪʧ��Ʊ�ļ�Ʊ���ݣ�ע����?", _
            vbQuestion + vbYesNoCancel + vbDefaultButton2) = vbYes Then

                oChkTk.WriteOffCheckTicket txtTicketID
                MsgBox "ע���ɹ�", vbInformation, "��Ϣ"
                cmdWriteOff.Enabled = False
        End If
    Else
        MsgBox "Ʊ�Ų���Ϊ��", vbExclamation, "����"
    End If
    Exit Sub
Err_Done:
    MsgBox err.Description, vbExclamation, "���� -- " & err.Number
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
            err.Raise 23302, , "�ó��λ�δ��Ʊ,����ע����Ʊ[" & txtTicketID.Text & "]"
        End If
        lblTicketSellInfo.Enabled = True
        lblBusCheckInfo.Enabled = True
        lblBusID.Caption = tCheckTicket.szbusID
        lblDate.Caption = Format(tCheckTicket.dtBusDate, "YYYY��MM��DD��")
        nBusSerialNo = tCheckTicket.nBusSerialNo
        Select Case tCheckTicket.nCheckTicketType
            Case ECheckStatus.NormalTicket
                lblCheckInMode.Caption = "��������"
            Case ECheckStatus.ChangeTicket
                lblCheckInMode.Caption = "�ĳ˼���"
            Case ECheckStatus.MergeTicket
                lblCheckInMode.Caption = "�������"
        End Select
        lblCheckGate.Caption = Trim(tCheckTicket.szCheckGateName)
        lblCheckor.Caption = Trim(tCheckTicket.szCheckerID)
        lblCheckTime.Caption = Format(tCheckTicket.dtCheckTime, cszTimeStr)
        If nBusSerialNo > 0 Then
            lblSerial.Caption = nBusSerialNo
        Else
            lblSerial.Caption = "�̶�����"
        End If
    End If
End Sub
