VERSION 5.00
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmTicketInfo 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��Ʊ��Ϣ"
   ClientHeight    =   4875
   ClientLeft      =   2415
   ClientTop       =   2160
   ClientWidth     =   6645
   HelpContextID   =   20000100
   Icon            =   "frmTicketInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   Tag             =   "Modal"
   Begin RTComctl3.CoolButton CoolButton1 
      Height          =   345
      Left            =   3330
      TabIndex        =   42
      Top             =   4440
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "����"
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
      MICON           =   "frmTicketInfo.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1050
      Left            =   300
      TabIndex        =   26
      Top             =   840
      Width           =   6090
      Begin VB.Label lblSellStation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   3480
         TabIndex        =   43
         Top             =   210
         Width           =   360
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   195
         Picture         =   "frmTicketInfo.frx":0028
         Top             =   210
         Width           =   480
      End
      Begin VB.Label lblSerialNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         Height          =   180
         Left            =   4500
         TabIndex        =   36
         Top             =   480
         Width           =   90
      End
      Begin VB.Label lblSerialNoHead 
         BackStyle       =   0  'Transparent
         Caption         =   "�������:"
         Height          =   180
         Left            =   3480
         TabIndex        =   35
         Top             =   480
         Width           =   810
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Left            =   3900
         TabIndex        =   34
         Top             =   210
         Width           =   180
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ʱ��:"
         Height          =   180
         Left            =   840
         TabIndex        =   33
         Top             =   750
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ʊ��:"
         Height          =   180
         Left            =   840
         TabIndex        =   32
         Top             =   210
         Width           =   630
      End
      Begin VB.Label lblTicket 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "U000001"
         Height          =   195
         Left            =   1785
         TabIndex        =   31
         Top             =   210
         Width           =   630
      End
      Begin VB.Label lblStartTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1999-11-04 13:20:21"
         Height          =   180
         Left            =   1785
         TabIndex        =   30
         Top             =   750
         Width           =   1710
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����:"
         Height          =   180
         Index           =   1
         Left            =   840
         TabIndex        =   29
         Top             =   480
         Width           =   450
      End
      Begin VB.Label lblBus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "01025"
         Height          =   180
         Left            =   1785
         TabIndex        =   28
         Top             =   480
         Width           =   450
      End
      Begin VB.Label lblToStation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   4170
         TabIndex        =   27
         Top             =   210
         Width           =   360
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   2235
      Left            =   300
      TabIndex        =   1
      Top             =   1890
      Width           =   6090
      Begin RTComctl3.CoolButton flbInfoHead 
         Height          =   195
         Index           =   0
         Left            =   780
         TabIndex        =   41
         Top             =   1020
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   344
         BTYPE           =   8
         TX              =   "������Ϣ"
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
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmTicketInfo.frx":08F2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "��ƱԱ:"
         Height          =   180
         Left            =   3495
         TabIndex        =   25
         Top             =   495
         Width           =   630
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "��Ʊʱ��:"
         Height          =   180
         Index           =   0
         Left            =   3495
         TabIndex        =   24
         Top             =   210
         Width           =   810
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ʊ����:"
         Height          =   180
         Left            =   855
         TabIndex        =   23
         Top             =   750
         Width           =   810
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��λ��:"
         Height          =   180
         Left            =   855
         TabIndex        =   22
         Top             =   180
         Width           =   630
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Ʊ��:"
         Height          =   180
         Left            =   855
         TabIndex        =   21
         Top             =   465
         Width           =   450
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "��Ʊ״̬:"
         Height          =   180
         Left            =   3495
         TabIndex        =   20
         Top             =   780
         Width           =   810
      End
      Begin VB.Label lblChkStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "δ��"
         Height          =   180
         Left            =   4485
         TabIndex        =   19
         Top             =   780
         Width           =   360
      End
      Begin VB.Label lblTicketType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ȫƱ/��Ʊ"
         Height          =   240
         Left            =   1785
         TabIndex        =   18
         Top             =   750
         Width           =   810
      End
      Begin VB.Label lblSellTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "12-30 13:20:21"
         Height          =   180
         Left            =   4485
         TabIndex        =   17
         Top             =   210
         Width           =   1260
      End
      Begin VB.Label lblSeatNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "12"
         Height          =   240
         Left            =   1785
         TabIndex        =   16
         Top             =   180
         Width           =   180
      End
      Begin VB.Label lblTktPrice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "111.23"
         Height          =   180
         Left            =   1785
         TabIndex        =   15
         Top             =   495
         Width           =   540
      End
      Begin VB.Label lblSeller 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "½����"
         Height          =   180
         Left            =   4485
         TabIndex        =   14
         Top             =   495
         Width           =   540
      End
      Begin VB.Label lblInfoValue 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������Ʊ"
         Height          =   180
         Index           =   5
         Left            =   4500
         TabIndex        =   13
         Top             =   1920
         Width           =   720
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "ԭƱ��:"
         Height          =   180
         Index           =   5
         Left            =   3510
         TabIndex        =   12
         Top             =   1920
         Width           =   630
      End
      Begin VB.Label lblInfoValue 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������Ʊ"
         Height          =   180
         Index           =   4
         Left            =   4500
         TabIndex        =   11
         Top             =   1650
         Width           =   720
      End
      Begin VB.Label lblInfoValue 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "000001"
         Height          =   180
         Index           =   3
         Left            =   4500
         TabIndex        =   10
         Top             =   1380
         Width           =   540
      End
      Begin VB.Label lblInfoValue 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "½����"
         Height          =   180
         Index           =   2
         Left            =   1845
         TabIndex        =   9
         Top             =   1920
         Width           =   540
      End
      Begin VB.Label lblInfoValue 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "12.33"
         Height          =   180
         Index           =   1
         Left            =   1845
         TabIndex        =   8
         Top             =   1650
         Width           =   450
      End
      Begin VB.Label lblInfoValue 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1999-3-12"
         Height          =   180
         Index           =   0
         Left            =   1845
         TabIndex        =   7
         Top             =   1380
         Width           =   810
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "��Ʊ��ʽ:"
         Height          =   180
         Index           =   4
         Left            =   3510
         TabIndex        =   6
         Top             =   1650
         Width           =   810
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ƱԱ:"
         Height          =   180
         Index           =   2
         Left            =   825
         TabIndex        =   5
         Top             =   1920
         Width           =   630
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "ƾ֤���ݺ�:"
         Height          =   180
         Index           =   3
         Left            =   3510
         TabIndex        =   4
         Top             =   1380
         Width           =   990
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ǩ������:"
         Height          =   180
         Index           =   1
         Left            =   840
         TabIndex        =   3
         Top             =   1650
         Width           =   990
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ʊʱ��:"
         Height          =   180
         Index           =   0
         Left            =   840
         TabIndex        =   2
         Top             =   1380
         Width           =   810
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         X1              =   840
         X2              =   5790
         Y1              =   1275
         Y2              =   1275
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         X1              =   840
         X2              =   5775
         Y1              =   1290
         Y2              =   1290
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   210
         Picture         =   "frmTicketInfo.frx":090E
         Top             =   285
         Width           =   480
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   -30
      TabIndex        =   0
      Top             =   690
      Width           =   6885
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   6735
      TabIndex        =   37
      Top             =   0
      Width           =   6735
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ʊ��Ϣ:"
         Height          =   180
         Left            =   270
         TabIndex        =   38
         Top             =   270
         Width           =   810
      End
   End
   Begin RTComctl3.CoolButton cmdOK 
      Cancel          =   -1  'True
      Default         =   -1  'True
      Height          =   345
      Left            =   4740
      TabIndex        =   40
      Top             =   4440
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "ȷ��"
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
      MICON           =   "frmTicketInfo.frx":0C18
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
      TabIndex        =   39
      Top             =   4200
      Width           =   8745
   End
End
Attribute VB_Name = "frmTicketInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mszTicketID As String       'Ʊ��
Public g_oActiveUser As ActiveUser
Public moSeviceTicket As ServiceTicket
Dim maszAllUsers() As String
Dim mbIsLoaded As Boolean
Const LabelInfoHeadSepWidth = 200   '

'Private Sub cmdCheckInfo_Click()
'    Dim oTmp As CheckSysApp
'    Set oTmp = New CheckSysApp
'    oTmp.ShowCheckInfo g_oActiveUser, moSeviceTicket.REBusDate, lblBus.Caption, Val(lblSerialNo.Caption)
'    Unload Me
'End Sub

Private Sub cmdOk_Click()
    Unload Me
End Sub

Private Sub CoolButton1_Click()
DisplayHelp Me
End Sub

Private Sub flbInfoHead_Click(Index As Integer)
    On Error Resume Next
    Dim i As Integer
    For i = 1 To flbInfoHead.Count - 1
        flbInfoHead(i).BackColor = &HE0E0E0
    Next i
    flbInfoHead(Index).BackColor = &HFFFFFF
    Select Case flbInfoHead(Index).Tag
        Case "Returned"
            showReturned
        Case "Checked"
            showChecked
        Case "Changed"
            showChanged
        Case "Canceled"
            showCanceled
        Case "BeChanged"
            showBeChanged
    End Select
End Sub

Private Sub showReturned()
'*************************************************
'��ʾ��Ʊ��Ϣ
'*************************************************
    Dim tReturnedTkInfo As TReturnedTicketInfo
    Dim i As Integer
    For i = 0 To 5
        lblInfo(i).Visible = False
        lblInfoValue(i).Visible = False
    Next i
'    Label50.Caption = "������Ϣ"
    
    tReturnedTkInfo = moSeviceTicket.ReturnedInfo
    lblInfo(0).Caption = "��Ʊ����:"
    lblInfo(1).Caption = "��Ʊ������:"
    lblInfo(2).Caption = "��ƱԱ:"
    lblInfo(3).Caption = "ƾ֤���ݺ�:"
    lblInfo(4).Caption = "��Ʊ��ʽ:"
    lblInfoValue(0).Left = lblInfo(0).Left + lblInfo(0).Width + 100
    lblInfoValue(1).Left = lblInfo(1).Left + lblInfo(1).Width + 100
    lblInfoValue(2).Left = lblInfo(2).Left + lblInfo(2).Width + 100
    lblInfoValue(3).Left = lblInfo(3).Left + lblInfo(3).Width + 100
    lblInfoValue(4).Left = lblInfo(4).Left + lblInfo(4).Width + 100
    lblInfoValue(0).Caption = Format(tReturnedTkInfo.dtReturnTime, "YYYY-MM-DD HH:MM:DD")
    lblInfoValue(1).Caption = Format(tReturnedTkInfo.sgReturnCharge, "#0.00")
    
    
    lblInfoValue(2).Caption = Trim(tReturnedTkInfo.szOperatorID)
    If lblInfoValue(2).Caption = g_oActiveUser.UserID Then
        lblInfoValue(2).Caption = "[" & lblInfoValue(2).Caption & "]" & g_oActiveUser.UserName
    Else
        For i = 1 To ArrayLength(maszAllUsers)
            If Trim(maszAllUsers(i, 1)) = lblInfoValue(2).Caption Then
                lblInfoValue(2).Caption = "[" & lblInfoValue(2).Caption & "]" & Trim(maszAllUsers(i, 2))
                Exit For
            End If
        Next i
    End If
    
    
    lblInfoValue(3).Caption = tReturnedTkInfo.szCredenceID
    lblInfoValue(4).Caption = IIf(tReturnedTkInfo.nReturnType = 0, "������Ʊ", "ǿ����Ʊ")
    For i = 0 To 4
        lblInfo(i).Visible = True
        lblInfoValue(i).Visible = True
    Next i
 '   Label50.Caption = "��Ʊ��Ϣ"
End Sub
Private Sub showCanceled()
'*************************************************
'��ʾ��Ʊ��Ϣ
'*************************************************
    Dim tCanceledTicket As TCanceledTicketInfo
    Dim i As Integer
    For i = 0 To 5
        lblInfo(i).Visible = False
        lblInfoValue(i).Visible = False
    Next i
 '   Label50.Caption = "������Ϣ"
    
    tCanceledTicket = moSeviceTicket.CanceledInfo
    lblInfo(0).Caption = "��Ʊʱ��:"
    lblInfo(1).Caption = "��ƱԱ:"
    lblInfo(2).Caption = "��Ʊ��ʽ:"
    lblInfoValue(0).Left = lblInfo(0).Left + lblInfo(0).Width + 100
    lblInfoValue(1).Left = lblInfo(1).Left + lblInfo(1).Width + 100
    lblInfoValue(2).Left = lblInfo(2).Left + lblInfo(2).Width + 100
    
    lblInfoValue(0).Caption = Format(tCanceledTicket.dtCancelTime, "YYYY-MM-DD HH:MM:SS")
    lblInfoValue(1).Caption = Trim(tCanceledTicket.szOperatorID)
    If lblInfoValue(1).Caption = g_oActiveUser.UserID Then
        lblInfoValue(1).Caption = "[" & lblInfoValue(1).Caption & "]" & g_oActiveUser.UserName
    Else
        For i = 1 To ArrayLength(maszAllUsers)
            If Trim(maszAllUsers(i, 1)) = lblInfoValue(1).Caption Then
                lblInfoValue(1).Caption = "[" & lblInfoValue(1).Caption & "]" & Trim(maszAllUsers(i, 2))
                Exit For
            End If
        Next i
    End If
    
    lblInfoValue(2).Caption = IIf(tCanceledTicket.nCancelType = 0, "������Ʊ", "ǿ�Ʒ�Ʊ")
    For i = 0 To 2
        lblInfo(i).Visible = True
        lblInfoValue(i).Visible = True
    Next i
'    Label50.Caption = "��Ʊ��Ϣ"
    Exit Sub
End Sub
Private Sub showChanged()
'*************************************************
'��ʾ��ǩƱ��Ϣ
'*************************************************
    Dim tChangedTkInfo As TChangedTicketInfo
    Dim i As Integer
    For i = 0 To 5
        lblInfo(i).Visible = False
        lblInfoValue(i).Visible = False
    Next i
'    Label50.Caption = "������Ϣ"


    tChangedTkInfo = moSeviceTicket.ChangedInfo
    lblInfo(0).Caption = "��ǩ������:"
    lblInfo(1).Caption = "ƾ֤���ݺ�:"
    lblInfo(2).Caption = "ԭƱ��:"
    lblInfo(3).Caption = "ԭƱ��:"
    lblInfoValue(0).Left = lblInfo(0).Left + lblInfo(0).Width + 100
    lblInfoValue(1).Left = lblInfo(1).Left + lblInfo(1).Width + 100
    lblInfoValue(2).Left = lblInfo(2).Left + lblInfo(2).Width + 100
    lblInfoValue(3).Left = lblInfo(3).Left + lblInfo(3).Width + 100
    
    lblInfoValue(0).Caption = Format(tChangedTkInfo.sgChangeCharge, "#0.00")
    lblInfoValue(1).Caption = tChangedTkInfo.szCredenceID
    lblInfoValue(2).Caption = tChangedTkInfo.szTicketID
    lblInfoValue(3).Caption = tChangedTkInfo.sgTicketPrice
    For i = 0 To 3
        lblInfo(i).Visible = True
        lblInfoValue(i).Visible = True
    Next i
 '   Label50.Caption = "��ǩ��Ϣ"
End Sub

Private Sub ShowTicketInfo(szTicketID As String)
'*************************************************
'��ʾ��Ʊ��Ϣ
'*************************************************
    Dim tCheckInfo As TCheckedTicketInfo
'On Error Resume Next
'    mbSuccess = True
    Dim i As Integer
    On Error GoTo ErrorHandle
    moSeviceTicket.Identify szTicketID
    
    tCheckInfo = moSeviceTicket.CheckedInfo
    lblBus.Caption = moSeviceTicket.REBusID
    lblSeatNo.Caption = moSeviceTicket.SeatNo
    lblSeller.Caption = Trim(moSeviceTicket.Operator)
        
    If lblSeller.Caption = g_oActiveUser.UserID Then
        lblSeller.Caption = MakeDisplayString(lblSeller.Caption, g_oActiveUser.UserName)
    Else
        Dim oUser As User
        Set oUser = New User
        oUser.Init g_oActiveUser
        oUser.Identify lblSeller.Caption
        lblSeller.Caption = lblSeller.Caption & "/" & oUser.FullName
    End If

'�����˹̶����εķ���ʱ�䲻����Ϊ0��©�������жϳ�������
    lblSerialNo.Caption = tCheckInfo.nBusSerialNo
    lblStartTime.Caption = Format(moSeviceTicket.dtBusStartUpTime, cszDateStr)
    If tCheckInfo.szBusid = "" Then     'δ��
        If TimeValue(moSeviceTicket.dtBusStartUpTime) = 0 Then  '��������
            lblSerialNoHead.Caption = "��������"
            lblSerialNo.Caption = "��������"
        Else
            lblStartTime.Caption = lblStartTime.Caption & " " & Format(moSeviceTicket.dtBusStartUpTime, "HH:mm")
        End If
    Else
        If tCheckInfo.nBusSerialNo = 0 Then     '�̶�����
            lblStartTime.Caption = lblStartTime.Caption & " " & Format(moSeviceTicket.dtBusStartUpTime, "HH:mm")
        End If
    End If
        

'    lblStartTime.Caption = Format(tCheckInfo.dtBusDate, "YYYY-MM-DD HH:MM:SS")
'    lblStartTime.Caption = Format(moSeviceTicket.REBusDate, "YYYY-MM-DD HH:MM:SS")
    lblSellTime.Caption = Format(moSeviceTicket.SellTime, "MM-DD HH:mm")
    lblTicket.Caption = szTicketID
    Select Case moSeviceTicket.TicketType
        Case TP_FullPrice
            lblTicketType.Caption = "ȫƱ"
        Case TP_HalfPrice
            lblTicketType.Caption = "��Ʊ"
        Case TP_FreeTicket
            lblTicketType.Caption = "��Ʊ"
        Case TP_PreferentialTicket1
            lblTicketType.Caption = "ѧ��Ʊ"
        Case TP_PreferentialTicket2
            lblTicketType.Caption = "��Ʊ"
    End Select
    
    lblSellStation.Caption = moSeviceTicket.SellStationName
    lblToStation.Caption = moSeviceTicket.ToStationName
    lblTktPrice.Caption = moSeviceTicket.TicketPrice

    
    Dim nTmpIndex As Integer
    lblChkStatus.Caption = ""
    If moSeviceTicket.TicketStatus And ST_TicketChecked Then
        nTmpIndex = flbInfoHead.Count
        Load flbInfoHead(nTmpIndex)
        flbInfoHead(nTmpIndex).Left = flbInfoHead(nTmpIndex - 1).Left + flbInfoHead(nTmpIndex - 1).Width + LabelInfoHeadSepWidth
        'flbInfoHead(nTmpIndex).Visible = True
        flbInfoHead(nTmpIndex).Tag = "Checked"
        flbInfoHead(nTmpIndex).Caption = "��Ʊ��Ϣ"
        lblChkStatus.Caption = lblChkStatus.Caption & "/�Ѽ�"
    End If
    If moSeviceTicket.TicketStatus And ST_TicketReturned Then
        nTmpIndex = flbInfoHead.Count
        Load flbInfoHead(nTmpIndex)
        flbInfoHead(nTmpIndex).Left = flbInfoHead(nTmpIndex - 1).Left + flbInfoHead(nTmpIndex - 1).Width + LabelInfoHeadSepWidth
        'flbInfoHead(nTmpIndex).Visible = True
        flbInfoHead(nTmpIndex).Tag = "Returned"
        flbInfoHead(nTmpIndex).Caption = "��Ʊ��Ϣ"
        lblChkStatus.Caption = lblChkStatus.Caption & "/����"
    End If
    If moSeviceTicket.TicketStatus And ST_TicketCanceled Then
        nTmpIndex = flbInfoHead.Count
        Load flbInfoHead(nTmpIndex)
        flbInfoHead(nTmpIndex).Left = flbInfoHead(nTmpIndex - 1).Left + flbInfoHead(nTmpIndex - 1).Width + LabelInfoHeadSepWidth
        'flbInfoHead(nTmpIndex).Visible = True
        flbInfoHead(nTmpIndex).Tag = "Canceled"
        flbInfoHead(nTmpIndex).Caption = "��Ʊ��Ϣ"
        lblChkStatus.Caption = lblChkStatus.Caption & "/��Ʊ"
    End If
    If moSeviceTicket.TicketStatus And ST_TicketSellChange Then
        nTmpIndex = flbInfoHead.Count
        Load flbInfoHead(nTmpIndex)
        flbInfoHead(nTmpIndex).Left = flbInfoHead(nTmpIndex - 1).Left + flbInfoHead(nTmpIndex - 1).Width + LabelInfoHeadSepWidth
        'flbInfoHead(nTmpIndex).Visible = True
        flbInfoHead(nTmpIndex).Tag = "Changed"
        flbInfoHead(nTmpIndex).Caption = "��ǩ��Ϣ"
        lblChkStatus.Caption = lblChkStatus.Caption & "/��ǩ�۳�"
    End If
    If moSeviceTicket.TicketStatus And ST_TicketChanged Then
        nTmpIndex = flbInfoHead.Count
        Load flbInfoHead(nTmpIndex)
        flbInfoHead(nTmpIndex).Left = flbInfoHead(nTmpIndex - 1).Left + flbInfoHead(nTmpIndex - 1).Width + LabelInfoHeadSepWidth
        'flbInfoHead(nTmpIndex).Visible = True
        flbInfoHead(nTmpIndex).Tag = "BeChanged"
        flbInfoHead(nTmpIndex).Caption = "����ǩ��Ϣ"
        lblChkStatus.Caption = lblChkStatus.Caption & "/����ǩ"
    End If
    If lblChkStatus.Caption = "" Then
        lblChkStatus.Caption = "/�����۳�"
    End If
    
    For i = 0 To 5      '���
        lblInfo(i).Visible = False
        lblInfoValue(i).Visible = False
    Next i
    If flbInfoHead.Count > 1 Then    '��ʾ��һ��
        flbInfoHead_Click 1
    End If
    Exit Sub
ErrorHandle:
    ShowErrorMsg
'ErrorPos:
'    MsgBox Err.Description, vbCritical, cszErrTitle & Trim(Str(Err.Number))
'    mbSuccess = False
End Sub

Public Property Get TicketID() As String
    TicketID = mszTicketID
End Property

Public Property Let TicketID(ByVal vNewValue As String)
    mszTicketID = vNewValue
End Property


Private Sub showChecked()
'*************************************************
'��ʾ��Ʊ��Ϣ
'*************************************************
    
    Dim tCheckTicket As TCheckedTicketInfo
    Dim i As Integer
    For i = 0 To 5
        lblInfo(i).Visible = False
        lblInfoValue(i).Visible = False
    Next i
'    Label50.Caption = "������Ϣ"
    
    
    
    tCheckTicket = moSeviceTicket.CheckedInfo
    lblInfo(0).Caption = "��Ʊʱ��:"
    lblInfo(1).Caption = "��Ʊ��ʽ:"
    lblInfoValue(0).Left = lblInfo(0).Left + lblInfo(0).Width + 100
    lblInfoValue(1).Left = lblInfo(1).Left + lblInfo(1).Width + 100
    
    lblInfoValue(0).Caption = Format(tCheckTicket.dtCheckTime, "YYYY-MM-DD HH:mm:ss")
    lblInfoValue(1).Caption = getCheckedTicketStatus(tCheckTicket.nCheckTicketType)
    For i = 0 To 1
        lblInfo(i).Visible = True
        lblInfoValue(i).Visible = True
    Next i
'    Label50.Caption = "��Ʊ��Ϣ"
End Sub
Public Sub RefreshForm()
    ShowTicketInfo mszTicketID
'    If Not bSucced Then GoTo ErrorPos
'    cmdCheckInfo.Enabled = IIf(lblBus.Caption <> "", True, False)
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandle
    mbIsLoaded = True
    If moSeviceTicket Is Nothing Then
        Set moSeviceTicket = New ServiceTicket
        moSeviceTicket.Init g_oActiveUser
    End If
    
    RefreshForm
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mbIsLoaded = False
End Sub


Public Property Get IsLoaded() As Boolean
    IsLoaded = mbIsLoaded
End Property

Private Sub showBeChanged()
'*************************************************
'��ʾ��ǩƱ��Ϣ
'*************************************************
    Dim i As Integer
    For i = 0 To 5
        lblInfo(i).Visible = False
        lblInfoValue(i).Visible = False
    Next i
'    Label50.Caption = "������Ϣ"


'    tChangedTkInfo = moSeviceTicket.ChangedInfo
    lblInfo(0).Caption = "��Ʊ��:"
    lblInfoValue(0).Left = lblInfo(0).Left + lblInfo(0).Width + 100
    
    lblInfoValue(0).Caption = moSeviceTicket.BeChanedToTicket
    For i = 0 To 0
        lblInfo(i).Visible = True
        lblInfoValue(i).Visible = True
    Next i
 '   Label50.Caption = "��ǩ��Ϣ"
End Sub

