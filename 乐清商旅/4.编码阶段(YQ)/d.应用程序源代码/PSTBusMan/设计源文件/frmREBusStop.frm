VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmEnvBusStop 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����--����ͣ��"
   ClientHeight    =   4215
   ClientLeft      =   3270
   ClientTop       =   4860
   ClientWidth     =   7110
   HelpContextID   =   10000240
   Icon            =   "frmREBusStop.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fraStop 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "ͣ�෽ʽ"
      ForeColor       =   &H80000008&
      Height          =   945
      Left            =   120
      TabIndex        =   19
      Top             =   1230
      Width           =   5355
      Begin VB.CheckBox cbStopBus 
         BackColor       =   &H00E0E0E0&
         Caption         =   "����ͣ��"
         Height          =   285
         Left            =   3105
         TabIndex        =   6
         Top             =   570
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   300
         Left            =   870
         TabIndex        =   3
         Top             =   240
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   529
         _Version        =   393216
         Format          =   60555264
         CurrentDate     =   36392
      End
      Begin MSComCtl2.DTPicker dtpEndDate 
         Height          =   300
         Left            =   3105
         TabIndex        =   5
         Top             =   240
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   529
         _Version        =   393216
         Format          =   60555264
         CurrentDate     =   36392
      End
      Begin VB.Label lblStartDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��(&K):"
         Height          =   180
         Left            =   300
         TabIndex        =   2
         Top             =   300
         Width           =   540
      End
      Begin VB.Label lblEndDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��(&E):"
         Height          =   180
         Left            =   2535
         TabIndex        =   4
         Top             =   300
         Width           =   540
      End
   End
   Begin VB.Frame fraEnvir 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   1785
      Left            =   150
      TabIndex        =   18
      Top             =   2280
      Width           =   6705
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgBusInfo 
         Height          =   1515
         Left            =   30
         TabIndex        =   12
         Top             =   270
         Width           =   6660
         _ExtentX        =   11748
         _ExtentY        =   2672
         _Version        =   393216
         Rows            =   4
         Cols            =   6
         BackColorFixed  =   14737632
         BackColorBkg    =   14737632
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         X1              =   1170
         X2              =   6750
         Y1              =   105
         Y2              =   105
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         X1              =   1170
         X2              =   6720
         Y1              =   90
         Y2              =   90
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ϸ���(&Z):"
         Height          =   180
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   1080
      End
   End
   Begin RTComctl3.CoolButton cmdHelp 
      Height          =   315
      Left            =   5730
      TabIndex        =   9
      Top             =   880
      Width           =   1185
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "����(&H)"
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
      MICON           =   "frmREBusStop.frx":014A
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
      Left            =   5730
      TabIndex        =   8
      Top             =   515
      Width           =   1185
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "ȡ��(&C)"
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
      MICON           =   "frmREBusStop.frx":0166
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
      Left            =   5730
      TabIndex        =   7
      Top             =   150
      Width           =   1185
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "ȷ��(&O)"
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
      MICON           =   "frmREBusStop.frx":0182
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin FText.asFlatTextBox txtBusID 
      Height          =   300
      Left            =   1050
      TabIndex        =   1
      Top             =   150
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      Registered      =   -1  'True
   End
   Begin RTComctl3.CoolButton cmdAllInfo 
      Height          =   345
      Left            =   5730
      TabIndex        =   10
      Top             =   1830
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "��ϸ>>"
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
      MICON           =   "frmREBusStop.frx":019E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   -1  'True
      VALUE           =   0   'False
   End
   Begin VB.Label lblStartupTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����ʱ��:"
      Height          =   180
      Left            =   2730
      TabIndex        =   20
      Top             =   225
      Width           =   810
   End
   Begin VB.Label lblAllRefundment 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ȫ����Ʊ:"
      Height          =   180
      Left            =   4380
      TabIndex        =   17
      Top             =   885
      Width           =   810
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����״̬:"
      Height          =   180
      Left            =   210
      TabIndex        =   16
      Top             =   885
      Width           =   810
   End
   Begin VB.Label lblSellSeat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��������:"
      Height          =   180
      Left            =   2730
      TabIndex        =   15
      Top             =   885
      Width           =   810
   End
   Begin VB.Label lblVehicle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���г���:"
      Height          =   180
      Left            =   2730
      TabIndex        =   14
      Top             =   555
      Width           =   810
   End
   Begin VB.Label lblRoute 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������·:"
      Height          =   180
      Left            =   210
      TabIndex        =   13
      Top             =   555
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���δ���:"
      Height          =   180
      Left            =   210
      TabIndex        =   0
      Top             =   225
      Width           =   810
   End
End
Attribute VB_Name = "frmEnvBusStop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_szBusID As String '���δ���
Public m_dtBusDate As Date '������������
'Public m_bStopResume As Boolean '��ͣ�໹�Ǹ���


Private m_oReBus As New REBus '��������
Private mbShowEnvir As Boolean


Private Sub cmdAllInfo_Click()
    On Error GoTo ErrHandle
    Dim i As Integer
    If Not cmdAllInfo.Value Then
        cmdAllInfo.Caption = "��ϸ>>"
        Me.Height = Me.Height - fraEnvir.Height
        fraEnvir.Visible = False
        Exit Sub
    Else
        cmdAllInfo.Caption = "��ϸ<<"
        Me.Height = Me.Height + fraEnvir.Height
        fraEnvir.Visible = True
    End If
    If Not mbShowEnvir Then
        Dim oVehicle As Vehicle
        Set oVehicle = New Vehicle
    
        SetBusy
        m_oReBus.Init g_oActiveUser
        oVehicle.Init g_oActiveUser
        hfgBusInfo.Redraw = False
        hfgBusInfo.Cols = 7
        hfgBusInfo.Rows = g_nPreSell + 2
        hfgBusInfo.Row = 0
        hfgBusInfo.Col = 0
        hfgBusInfo.Text = "����"
        hfgBusInfo.Col = 1
        hfgBusInfo.Text = "״̬"
        hfgBusInfo.Col = 2
        hfgBusInfo.Text = "������λ��"
        hfgBusInfo.Col = 3
        hfgBusInfo.Text = "���г���"
        hfgBusInfo.Col = 4
        hfgBusInfo.Text = "����"
        hfgBusInfo.Col = 5
        hfgBusInfo.Text = "������λ��"
        hfgBusInfo.Col = 6
        hfgBusInfo.Text = "����"
        For i = 0 To g_nPreSell
            hfgBusInfo.Row = i + 1
            hfgBusInfo.Col = 0
            hfgBusInfo.Text = Format(DateAdd("d", i, Now), "YYYY-MM-DD")
            hfgBusInfo.Col = 1
            hfgBusInfo.Text = "�޳�������"
            m_oReBus.Identify m_szBusID, DateAdd("d", i, Now)
            Select Case m_oReBus.busStatus
                   Case ST_BusMergeStopped
                        hfgBusInfo.Text = "������"
                        hfgBusInfo.CellForeColor = vbBlue
                   Case ST_BusNormal
                        hfgBusInfo.Text = "����"
                        hfgBusInfo.CellForeColor = vbBlack
                   Case ST_BusStopCheck
                        hfgBusInfo.Text = "ͣ��"
                        hfgBusInfo.CellForeColor = vbBlack
                   Case ST_BusStopped
                        hfgBusInfo.Text = "ͣ��"
                        hfgBusInfo.CellForeColor = vbRed
            End Select
            hfgBusInfo.Col = 2
            hfgBusInfo.Text = m_oReBus.SaledSeatCount
            hfgBusInfo.Col = 3
            hfgBusInfo.Text = m_oReBus.Vehicle
            hfgBusInfo.Col = 4
            oVehicle.Identify m_oReBus.Vehicle
            hfgBusInfo.Text = oVehicle.LicenseTag
            hfgBusInfo.Col = 5
            hfgBusInfo.Text = m_oReBus.SaleSeat
            hfgBusInfo.Col = 6
            hfgBusInfo.Text = m_oReBus.VehicleModelName
NextBus:
        Next
        hfgBusInfo.Redraw = True
        mbShowEnvir = True
        SetNormal
    End If
    Exit Sub

ErrHandle:
    Select Case err.Number
           Case ERR_REBusNotExist: Resume NextBus
           Case Else: SetNormal: ShowErrorMsg
    End Select
End Sub

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdHelp_Click()
DisplayHelp Me
End Sub

Public Sub cmdOk_Click()
    Const cnPreViewMaxDays = 50
    If dtpEndDate <= DateAdd("D", cnPreViewMaxDays, dtpStartDate) Then
        If DateDiff("D", dtpEndDate, dtpStartDate) > 0 Then
           MsgBox "ͣ����ʼʱ�䲻�ܴ��ڽ���ʱ��!", vbExclamation, Me.Caption
           Exit Sub
        End If
        If cbStopBus.Value = 1 And frmEnvBus.lvBus.ListItems.Count > 0 Then
           '����ͣ
            frmEnvBus.SelectedStopBus dtpStartDate, dtpEndDate, True
            Unload Me
            Exit Sub
        End If
          
        If MsgBox("�Ƿ�ͣ��[" & Format(dtpStartDate.Value, "YYYY��MM��DD��") & "]��[" & Format(dtpEndDate.Value, "YYYY��MM��DD��") & "]��ָ�����Σ�", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
            StopBus Trim(txtBusID.Text)
        End If
    Else
         MsgBox "����ͣ���������ܳ���" & cnPreViewMaxDays & "��", vbExclamation, Me.Caption
    End If
End Sub

Private Sub Form_Load()
    AlignFormPos Me
    
    m_oReBus.Init g_oActiveUser
    If m_szBusID <> "" Then
        dtpStartDate.Value = m_dtBusDate
        dtpEndDate.Value = m_dtBusDate
        FullREBusInfo m_szBusID, m_dtBusDate
        cmdAllInfo.Enabled = True
        cmdOk.Enabled = True
    Else
        dtpStartDate.Value = Date
        dtpEndDate.Value = Date
        m_dtBusDate = Date
        lblStartupTime = "����ʱ��:" & Format(Date, "YYYY��MM��DD��")
    End If
    Dim i As Long, nCount As Long
    With frmEnvBus.lvBus.ListItems
    For i = 1 To .Count
        If .Item(i).Selected Then
            nCount = nCount + 1
        End If
    Next i
    If nCount <= 1 Then
        cbStopBus.Visible = False
    Else
        cbStopBus.Caption = cbStopBus.Caption & nCount & "������"
    End If
    End With
    
    Call cmdAllInfo_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveFormPos Me
End Sub

Private Sub txtBusID_ButtonClick()
    Dim oBus As New CommDialog
    Dim aszTemp() As String
    oBus.Init g_oActiveUser
    aszTemp = oBus.SelectREBus(m_dtBusDate)
    Set oBus = Nothing
    If ArrayLength(aszTemp) = 0 Then Exit Sub
    txtBusID.Text = aszTemp(1, 1)
    m_szBusID = txtBusID.Text
    m_dtBusDate = dtpStartDate.Value
    FullREBusInfo m_szBusID, m_dtBusDate
End Sub



Private Sub txtBusID_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
           Case vbKeyReturn
           FullREBusInfo txtBusID.Text, dtpStartDate.Value
           Me.Caption = "����--����ͣ��[" & Format(dtpStartDate.Value, "YYYY��MM��DD��") & "]"
    End Select
End Sub

Private Sub FullREBusInfo(BusID As String, BusDate As Date)
On Error GoTo ErrHandle
    m_szBusID = BusID
    
    m_oReBus.Identify BusID, BusDate
    txtBusID.Text = m_szBusID
    lblRoute.Caption = "������·:" & m_oReBus.RouteName
    If m_oReBus.BusType = TP_ScrollBus Then
        lblStartupTime.Caption = "����ʱ��:" & Format(m_oReBus.RunDate, "YYYY��MM��DD��") & "  " & m_oReBus.ScrollBusCheckTime & "����һ��"
    Else
        lblStartupTime.Caption = "����ʱ��:" & Format(m_oReBus.RunDate, "YYYY��MM��DD��") & "  " & Format(m_oReBus.StartUpTime, "hh:mm")
    End If
    lblVehicle.Caption = "���г���:" & m_oReBus.VehicleTag
    lblSellSeat.Caption = "��������:" & (m_oReBus.TotalSeat - m_oReBus.SaleSeat - m_oReBus.ReserveSeatCount)
    If m_oReBus.AllRefundment Then
        lblAllRefundment.Caption = "ȫ����Ʊ:��"
    Else
        lblAllRefundment.Caption = "ȫ����Ʊ:��"
    End If
    Select Case m_oReBus.busStatus
           Case ST_BusStopCheck: lblStatus.Caption = "����״̬:ͣ��"
           Case ST_BusNormal: lblStatus.Caption = "����״̬:����"
           Case ST_BusStopped: lblStatus.Caption = "����״̬:ͣ��"
           Case ST_BusMergeStopped: lblStatus.Caption = "����״̬:����"
           Case ST_BusSlitpStop: lblStatus.Caption = "����״̬:���ͣ"
           
    End Select
    cmdAllInfo.Enabled = True
    cmdOk.Enabled = True
    m_szBusID = BusID
    m_dtBusDate = BusDate
Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub StopBus(szBusID As String)
    Dim szTemp As String
    Dim nDay As Integer, nCount As Integer, i As Integer
    Dim dtStop As Date
    Dim szMsgBusStatus As String
    Dim nBusStatus As Integer
On Error GoTo ErrHandle
    szTemp = Format(dtpStartDate.Value, "YYYY��MM��DD��") & "��" & Format(dtpEndDate.Value, "YYYY��MM��DD��")
    szTemp = szTemp & "����ͣ��ɹ�"
    nDay = DateDiff("d", dtpStartDate.Value, dtpEndDate.Value)
    szTemp = ""
    
    WriteProcessBar True, , nDay + 1
    For i = 0 To nDay
        dtStop = DateAdd("d", i, dtpStartDate.Value)
        m_oReBus.Identify szBusID, dtStop
        szMsgBusStatus = ""
        If m_oReBus.SaledSeatCount > 0 Then
            If MsgBox("����[" & szBusID & "]����[" & m_oReBus.SaledSeatCount & "��]��Ʊ���Ƿ�ͣ�ࣿ", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
              'szMsg = szMsg & "����[" & szbusID & "]ͣ��ʧ��!" & vbCrLf
              GoTo NextBus
            End If
        End If
        If m_oReBus.HaveLugge = True Then
            If MsgBox("����[" & szBusID & "]���а��Ƿ�ͣ��", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then Exit Sub
        End If
        If IdentifyBusStatus(m_oReBus.busStatus) = False Then GoTo NextBus
        m_oReBus.StopBus dtStop, dtStop, g_bStopAllRefundment
        WriteProcessBar , i + 1, nDay + 1, "ͣ��" & Format(dtStop, "YYYY-MM-DD")
        frmEnvBus.UpdateList m_oReBus.BusID, dtStop
        szTemp = szTemp & szBusID & "����[" & Format(dtStop, "YYYY-MM-DD") & "]ͣ��ɹ�!" & vbCrLf
NextBus:
    Next
    WriteProcessBar False, , , ""
    If szTemp <> "" Then
        MsgBox szTemp, vbInformation, Me.Caption
    End If
Exit Sub
ErrHandle:
    ShowSBInfo "[" & Format(dtStop, "YYYY-MM-DD") & "]" & err.Description
    szTemp = szTemp & vbCrLf & szBusID & "����[" & Format(dtStop, "YYYY-MM-DD") & "]" & err.Description
    Resume NextBus
End Sub
