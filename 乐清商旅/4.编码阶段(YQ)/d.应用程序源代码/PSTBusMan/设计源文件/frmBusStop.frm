VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmBusStop 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�ƻ�--����ͣ��"
   ClientHeight    =   4410
   ClientLeft      =   1485
   ClientTop       =   3885
   ClientWidth     =   7110
   HelpContextID   =   2002601
   Icon            =   "frmBusStop.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fraEnvir 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   1785
      Left            =   150
      TabIndex        =   17
      Top             =   2460
      Width           =   6705
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgBusInfo 
         Height          =   1515
         Left            =   30
         TabIndex        =   18
         Top             =   270
         Width           =   6660
         _ExtentX        =   11748
         _ExtentY        =   2672
         _Version        =   393216
         Rows            =   4
         Cols            =   6
         BackColorFixed  =   14737632
         BackColorBkg    =   14737632
         ScrollBars      =   2
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ϸ���(&Z):"
         Height          =   180
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   1080
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         X1              =   1170
         X2              =   6720
         Y1              =   90
         Y2              =   90
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         X1              =   1170
         X2              =   6750
         Y1              =   105
         Y2              =   105
      End
   End
   Begin RTComctl3.CoolButton cmdAllInfo 
      Height          =   345
      Left            =   5730
      TabIndex        =   16
      Top             =   1980
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
      MICON           =   "frmBusStop.frx":014A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   -1  'True
      VALUE           =   0   'False
   End
   Begin VB.Frame fraStop 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ͣ�෽ʽ"
      Height          =   1395
      Left            =   120
      TabIndex        =   8
      Top             =   930
      Width           =   5355
      Begin VB.OptionButton optBusStop 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ʱ���ͣ��(&S)"
         Height          =   210
         Left            =   270
         MaskColor       =   &H8000000B&
         TabIndex        =   11
         Top             =   615
         UseMaskColor    =   -1  'True
         Width           =   1830
      End
      Begin VB.OptionButton optLongStop 
         BackColor       =   &H00E0E0E0&
         Caption         =   "��ͣ(&L)"
         Height          =   195
         Left            =   270
         TabIndex        =   10
         Top             =   270
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.CheckBox chStop 
         BackColor       =   &H00E0E0E0&
         Caption         =   "����ͣ��"
         Height          =   375
         Left            =   2610
         TabIndex        =   9
         Top             =   150
         Width           =   2625
      End
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   300
         Left            =   1080
         TabIndex        =   12
         Top             =   930
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   529
         _Version        =   393216
         Format          =   60751872
         CurrentDate     =   36392
      End
      Begin MSComCtl2.DTPicker dtpEndDate 
         Height          =   300
         Left            =   3315
         TabIndex        =   13
         Top             =   930
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   529
         _Version        =   393216
         Format          =   60751872
         CurrentDate     =   36392
      End
      Begin VB.Label lblEndDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��(&E):"
         Height          =   180
         Left            =   2745
         TabIndex        =   15
         Top             =   990
         Width           =   540
      End
      Begin VB.Label lblStartDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��(&K):"
         Height          =   180
         Left            =   510
         TabIndex        =   14
         Top             =   990
         Width           =   540
      End
   End
   Begin RTComctl3.CoolButton cmdSave 
      Default         =   -1  'True
      Height          =   315
      Left            =   5730
      TabIndex        =   0
      Top             =   165
      Width           =   1170
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "ͣ��"
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
      MICON           =   "frmBusStop.frx":0166
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
      Left            =   5715
      TabIndex        =   1
      Top             =   540
      Width           =   1170
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "�ر�"
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
      MICON           =   "frmBusStop.frx":0182
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdHelp 
      Height          =   315
      Left            =   5715
      TabIndex        =   2
      Top             =   930
      Width           =   1170
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
      MICON           =   "frmBusStop.frx":019E
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
      Left            =   1080
      TabIndex        =   5
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
   Begin VB.Label lblCheck 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�������г���:6��"
      Height          =   180
      Left            =   2760
      TabIndex        =   7
      Top             =   600
      Width           =   1440
   End
   Begin VB.Label lblRouteName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������·:���˵�����"
      Height          =   180
      Left            =   255
      TabIndex        =   6
      Top             =   600
      Width           =   1710
   End
   Begin VB.Label label98 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���δ���:"
      Height          =   180
      Left            =   255
      TabIndex        =   4
      Top             =   225
      Width           =   810
   End
   Begin VB.Label lblStartupTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����ʱ��:00:01:02"
      Height          =   180
      Left            =   2760
      TabIndex        =   3
      Top             =   225
      Width           =   1530
   End
End
Attribute VB_Name = "frmBusStop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'���泵��ͣ��
Option Explicit
Option Base 1
Const cnDate = 0
Const cnStatus = 1
Const cnSellSeats = 2
Const cnVehicle = 3
Const cnLicenseTag = 4
Const cnCanSellSeats = 5
Const cnVehicleType = 6
Const cnCols = 7

Public m_szBusID As String '���δ���
Public m_szPlanID As String  '�ƻ�����
Public m_szVehicle As String


Private m_oBus As New Bus
Private m_oREBus As New REBus
Private mbShowEnvir As Boolean
Public Status As Integer '1Ϊͣ�� 0Ϊ����



Private Sub cmdAllInfo_Click()
    On Error GoTo ErrHandle
    Dim i As Integer
    If cmdAllInfo.Value = False Then
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
        m_oREBus.Init g_oActiveUser
        oVehicle.Init g_oActiveUser
        With hfgBusInfo
            .Redraw = False
            .Cols = cnCols
            .ColWidth(cnSellSeats) = 600
            .ColWidth(cnVehicle) = 800
            .ColWidth(cnCanSellSeats) = 600
            .Rows = g_nPreSell + 2
            .Row = 0
            .Col = cnDate
            .Text = "����"
            .Col = cnStatus
            .Text = "״̬"
            .Col = cnSellSeats
            .Text = "����" '"������λ��"
            .Col = cnVehicle
            .Text = "���г���"
            .Col = cnLicenseTag
            .Text = "����"
            .Col = cnCanSellSeats
            .Text = "����" '"������λ��"
            .Col = cnVehicleType
            .Text = "����"
            For i = 0 To g_nPreSell
                .Row = i + 1
                .Col = cnDate
                .Text = Format(DateAdd("d", i, Now), "YYYY-MM-DD")
                .Col = cnStatus
                .Text = "�޳�������"
                m_oREBus.Identify txtBusID.Text, DateAdd("d", i, Now)
                Select Case m_oREBus.busStatus
                Case ST_BusMergeStopped
                    .Text = "������"
                    .CellForeColor = vbBlue
                Case ST_BusNormal
                    .Text = "����"
                    .CellForeColor = vbBlack
                Case ST_BusStopCheck
                    .Text = "ͣ��"
                    .CellForeColor = vbBlack
                Case ST_BusStopped
                    .Text = "ͣ��"
                    .CellForeColor = vbRed
                End Select
                .Col = cnSellSeats
                .Text = m_oREBus.SaledSeatCount
                .Col = cnVehicle
                .Text = m_oREBus.Vehicle
                .Col = cnLicenseTag
                oVehicle.Identify m_oREBus.Vehicle
                .Text = oVehicle.LicenseTag
                .Col = cnCanSellSeats
                .Text = m_oREBus.SaleSeat
                .Col = cnVehicleType
                .Text = m_oREBus.VehicleModelName
NextBus:
            Next
            .Redraw = True
        End With
        mbShowEnvir = True
        SetNormal
    End If
    Exit Sub
ErrHandle:
    Select Case err.Number
    Case ERR_REBusNotExist
        Resume NextBus
    Case Else
        SetNormal
        ShowErrorMsg
    End Select
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    DisplayHelp Me
End Sub

Private Sub cmdSave_Click()
    If dtpEndDate.Value < dtpStartDate.Value Then
        MsgBox cmdSave.Caption & "�Ŀ�ʼ���ڱ���С�ڵ��ڽ�������", vbExclamation, Me.Caption
        Exit Sub
    End If
    
    If Me.Caption = "���γ���ͣ��" Then
        SaveBusVehicle
    Else
        SaveBus
    End If
    Unload Me
End Sub



Private Sub Form_Load()
On Error GoTo ErrHandle
    txtBusID.Text = m_szBusID
    AlignFormPos Me
    '��ʼ��������Ϣ
    RefreshBus
    '��������ͣ���Ƿ���Ч
    If Me.Caption <> "���γ���ͣ��" Then
        Dim i As Long, nCount As Long
        With frmBus.lvBus.ListItems
        For i = 1 To .Count
            If .Item(i).Selected Then
                nCount = nCount + 1
            End If
        Next i
        If nCount <= 1 Then
            chStop.Visible = False
        Else
            chStop.Caption = chStop.Caption & nCount & "������"
        End If
        End With
    End If
    
    Call cmdAllInfo_Click
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub
'ˢ�³�����Ϣ
Private Sub RefreshBus()
    m_oBus.Init g_oActiveUser
    m_oBus.Identify txtBusID.Text
    If m_oBus.BusType = TP_ScrollBus Then
        lblStartupTime.Caption = "ÿ����ʱ��:" & m_oBus.ScrollBusCheckTime & "����"
    Else
        lblStartupTime.Caption = "����ʱ��:" & Format(m_oBus.StartUpTime, "hh:mm")
    End If
    lblRouteName.Caption = "��·����:" & m_oBus.RouteName
    lblCheck.Caption = "��Ʊ��:" & m_oBus.CheckGate

    
    If Status = 1 Then
        optLongStop.Value = True
        optLongStop_Click
        dtpStartDate.Value = Date
        dtpEndDate.Value = Date
        
        optLongStop.Caption = "��ͣ"
        optBusStop.Caption = "ʱ���ͣ��"
        fraStop.Caption = "ͣ�෽ʽ"
        Caption = "�ƻ�����ͣ��"
        cmdSave.Caption = "ͣ��"
        chStop.Caption = "����ͣ��"
    Else
        optLongStop.Value = True
'        optBusStop.Visible = False
'        lblStartDate.Visible = False
'        dtpStartDate.Visible = False
'        dtpEndDate.Visible = False
'        lblEndDate.Visible = False
        optBusStop_Click
        dtpStartDate.Value = Date 'm_oBus.BeginStopDate
        dtpEndDate.Value = Date 'm_oBus.EndStopDate

        optLongStop.Caption = "����"
        optBusStop.Caption = "ʱ��θ���"
        fraStop.Caption = "���෽ʽ"
        Caption = "�ƻ����θ���"
        cmdSave.Caption = "����"
        chStop.Caption = "��������"
    End If

End Sub



Private Sub Form_Unload(Cancel As Integer)
    SaveFormPos Me
End Sub

'B2����ͣ��
Private Sub optBusStop_Click()
    dtpEndDate.Enabled = True
    dtpStartDate.Enabled = True
End Sub

'B4���θ���
Private Sub optLongStop_Click()
    dtpEndDate.Enabled = False
    dtpStartDate.Enabled = False
End Sub

Public Sub Init(Optional Bus As Bus, Optional BusID As String, Optional PlanID As String, Optional szShowType As Boolean)
    Set m_oBus = Bus
    m_szBusID = BusID
    m_szPlanID = PlanID
    
End Sub
'�ƻ�����ͣ��/����
Private Sub SaveBus()
    Dim nResult  As VbMsgBoxResult
    Dim dtREStartStop As Date, dtREEndStop As Date
    Dim i As Integer, nDate As Integer
    Dim szQuery As String, szErrString As String
    Dim liTemp As ListItem
    Dim szShowMsg As String
    Dim oReSheme As New REScheme
    Dim bEnviroment As Boolean
    Dim szSaleCountInfo As String
    Dim szString As String
    Dim szTextMsgDate As String
    Dim bISLongdate As Boolean
    Dim j As Integer
    Dim nSel As Integer
    Dim nCountBus As Integer
    Dim szbusID() As String
    Dim nSaledSeatCount As Integer

    szShowMsg = cmdSave.Caption
    nCountBus = frmBus.lvBus.ListItems.Count
    
    On Error GoTo ErrHandle
    If chStop.Value = vbChecked Then
        '����
        For j = 1 To nCountBus
            If frmBus.lvBus.ListItems.Item(j).Selected = True Then
                nSel = nSel + 1
                MakeArray szbusID, frmBus.lvBus.ListItems.Item(j).Text
            End If
        Next
        szQuery = nSel & "��������[" & Format(dtpStartDate.Value, "YYYY-MM-DD") & "]��[" & Format(dtpEndDate.Value, "YYYY-MM-DD") & "]��ʱ����" & szShowMsg
        If optLongStop.Value = True Then
            szQuery = nSel & "�����γ�" & szShowMsg & "?"
        Else
            szQuery = nSel & "������" & szShowMsg & "?"
        End If
        If MsgBox(szQuery, vbQuestion + vbYesNo, Me.Caption) = vbNo Then Exit Sub
    Else
        '������
        ReDim szbusID(1 To 1) As String
        szbusID(1) = txtBusID.Text
        szQuery = "����" & txtBusID.Text & "��[" & Format(dtpStartDate.Value, "YYYY-MM-DD") & "]��[" & Format(dtpEndDate.Value, "YYYY-MM-DD") & "]��ʱ����" & szShowMsg
        If optLongStop.Value = True Then
            If frmBusStop.Caption = "�ƻ����θ���" Then
                szQuery = "����[" & txtBusID.Text & "]����?"
            Else
                szQuery = "����[" & txtBusID.Text & "]��ͣ?"
            End If
        Else
            szQuery = "����[" & txtBusID.Text & "]" & szShowMsg & "?"
        End If
        nSel = 1
        If MsgBox(szQuery, vbQuestion + vbYesNo, Me.Caption) = vbNo Then Exit Sub
    End If
    frmBusStop.MousePointer = vbHourglass
    m_oREBus.Init g_oActiveUser
    'ͣ���࿪ʼ
    For i = 1 To nSel
        szTextMsgDate = ""
        m_oBus.Identify szbusID(i)
        If optLongStop.Value = True Then
        
            If frmBusStop.Caption = "�ƻ����θ���" Then
                '������
                m_oBus.BeginStopDate = CDate(cszEmptyDateStr)
                m_oBus.EndStopDate = CDate(cszEmptyDateStr)
                bISLongdate = True
            
            Else
                '��ͣ��
                m_oBus.BeginStopDate = CDate(cszEmptyDateStr)
                m_oBus.EndStopDate = CDate(cszForeverDateStr)
                bISLongdate = True
            End If
        Else
            'ʱ���: ͣ�� ����
            If frmBusStop.Caption = "�ƻ����θ���" Then
                '������
                m_oBus.BeginStopDate = CDate(cszEmptyDateStr)
                m_oBus.EndStopDate = CDate(cszEmptyDateStr)
            Else
                'ʱ���ͣ��
                m_oBus.BeginStopDate = CDate(dtpStartDate.Value)
                m_oBus.EndStopDate = CDate(dtpEndDate.Value)
                szTextMsgDate = "��[" & Format(dtpStartDate.Value, "YYYY-MM-DD") & "��" & Format(dtpEndDate.Value, "YYYY-MM-DD") & "]ʱ��ͣ��"
            End If
        End If
        If frmBusStop.Caption = "�ƻ����θ���" Then
            ShowSBInfo "�ƻ�����[" & szbusID(i) & "]����..."
        Else
            ShowSBInfo "�ƻ�����[" & szbusID(i) & "]ͣ��..."
        End If
        '���³���
        m_oBus.Update
        'ˢ���б��
        With frmBus.lvBus
            If DateDiff("d", Now, CDate(m_oBus.BeginStopDate)) <= 0 And DateDiff("d", Now, CDate(m_oBus.EndStopDate)) >= 0 Then
                szString = "ͣ��"
            Else
                '������ͣ�������
                szString = "����"
            End If
            Set liTemp = .FindItem(szbusID(i), , lvwPartial)
            If frmBusStop.Caption <> "�ƻ����θ���" Then
                'ͣ�����
                If Not (liTemp Is Nothing) Then
                    If m_oBus.BusType <> TP_ScrollBus Then
                        '���泵��
                        If bISLongdate = True Then
                        '��ͣ����
                            liTemp.SmallIcon = "StopBus"
                            liTemp.ListSubItems(4).ForeColor = vbRed
                            liTemp.ListSubItems(4).Text = "ͣ��"
                        Else
                            If szString = "ͣ��" Then
                                liTemp.SmallIcon = "StopBus"
                                liTemp.ListSubItems(4).ForeColor = vbRed
                            End If
                            liTemp.ListSubItems(4).Text = szString & szTextMsgDate
                        End If
                    Else
                        '��������
                        If bISLongdate = True Then
                            '��ͣ����
                            liTemp.SmallIcon = "FlowStop"
                            liTemp.ListSubItems(4).ForeColor = vbRed
                            liTemp.ListSubItems(4).Text = "ͣ��" & szTextMsgDate
                        Else
                            If szString = "ͣ��" Then
                                liTemp.SmallIcon = "FlowStop"
                                liTemp.ListSubItems(4).ForeColor = vbRed
                            End If
                            liTemp.ListSubItems(4).Text = szString & szTextMsgDate
                        End If
                    End If
                End If
            Else
                '�������
                If Not (liTemp Is Nothing) Then
                
                    If m_oBus.BusType = TP_ScrollBus Then
                        If bISLongdate = True Then
                            '���������
                            liTemp.SmallIcon = "Flow"
                            liTemp.ListSubItems(4).ForeColor = vbBlack
                            liTemp.ListSubItems(4).Text = "����"
                        Else
                            If szString = "����" Then
                                liTemp.SmallIcon = "Flow"
                                liTemp.ListSubItems(4).ForeColor = vbBlack
                            End If
                            liTemp.ListSubItems(4).Text = szString & szTextMsgDate
                        End If
                    Else
    
                        If bISLongdate = True Then
                            '���������
                            liTemp.SmallIcon = "RunBus"
                            liTemp.ListSubItems(4).ForeColor = vbBlack
                            liTemp.ListSubItems(4).Text = "����"
                        Else
                            If szString = "����" Then
                                liTemp.SmallIcon = "RunBus"
                                liTemp.ListSubItems(4).ForeColor = vbBlack
                            End If
                            liTemp.ListSubItems(4).Text = szString & szTextMsgDate
                        End If
                    End If
                End If
            End If
        End With
NotUplstTemp:
    Next i
    SetNormal
    bEnviroment = True
    '���´���������
    szQuery = "�޳���" & szShowMsg & "..."
    If Not (DateDiff("d", dtpEndDate.Value, Now) >= 1 Or DateDiff("d", DateAdd("d", g_nPreSell, Now), dtpStartDate.Value) >= 1) Or optLongStop.Value = True Then
        '����ͣ����ʱ���
        If optLongStop.Value = True Then
            '���ǳ�ͣ��ʼ��--oResheme
            oReSheme.Init g_oActiveUser
            dtREStartStop = CDate(cszEmptyDateStr)
            dtREEndStop = CDate(cszEmptyDateStr)
            szQuery = "���ڵ��ڵ�ǰʱ��ĳ���"
        Else
            '��ͨ  ͣ�� ����
            'ȷ����ʼʱ��С�ڽ���ʱ��
            If dtpStartDate.Value <= dtpEndDate.Value Then
                dtREStartStop = dtpStartDate.Value
                dtREEndStop = dtpEndDate.Value
            Else
                dtREStartStop = dtpEndDate.Value
                dtREEndStop = dtpStartDate.Value
            End If
            szQuery = "[" & Format(dtREStartStop, "YYYY-MM-DD") & "]��[" & Format(dtREEndStop, "YYYY-MM-DD") & "]"
        End If
        nResult = MsgBox("�趨��" & szShowMsg & "ʱ��Ӱ�컷�����Ƿ�ͬʱ" & szShowMsg & "��������" & vbCrLf & vbCrLf & "����" & szShowMsg & "����:" & szQuery, vbQuestion + vbYesNo + vbDefaultButton2, "����" & szShowMsg)
        szQuery = ""
        If nResult = vbYes Then
            SetBusy
            nDate = DateDiff("d", dtREStartStop, dtREEndStop)
            ShowSBInfo
            For j = 1 To nSel
                If dtREStartStop = dtREEndStop And dtREEndStop = CDate(cszEmptyDateStr) Then
                    '��������
                    ShowSBInfo szShowMsg & "�����д��ڵ�ǰʱ������г���[" & szbusID(j) & "]"
                    If szShowMsg = "����" Then
                        oReSheme.AllResumeBus szbusID(j), dtREStartStop, dtREEndStop
                        szQuery = szQuery & vbCrLf & "*������:���ڵ�ǰʱ�������[" & szbusID(j) & "]����" & szShowMsg & "�ɹ�..."
                    Else
                        '������ͣ��
                        szSaleCountInfo = oReSheme.AllStopBus(szbusID(j), dtREStartStop, dtREEndStop, g_bStopAllRefundment)
                        szQuery = szSaleCountInfo & szQuery & vbCrLf & "������:���ڵ��ڵ�ǰʱ�������[" & szbusID(j) & "]����" & szShowMsg & "�ɹ�..."
                    End If
                Else
                    'ʱ��θ���ͣ��
                    For i = 0 To nDate
                        ShowSBInfo szShowMsg & "��������[" & szbusID(j) & "]" & Format(DateAdd("d", i, dtREStartStop), "YYYY-MM-DD")
                        m_oREBus.Identify szbusID(j), DateAdd("d", i, dtREStartStop)
                        dtREEndStop = DateAdd("d", i, dtREStartStop)
                        If szShowMsg = "����" Then
                            m_oREBus.ResumeBus dtREStartStop, dtREEndStop
                            szQuery = szQuery & vbCrLf & "*��������[" & szbusID(j) & "]" & Format(dtREEndStop, "YYYY��MM��DD��") & szShowMsg & "�ɹ�..."
                        Else
                            '����Ʊ
                            nSaledSeatCount = m_oREBus.SaledSeatCount
                            If nSaledSeatCount > 0 Then
                                szQuery = szQuery & vbCrLf & "ע��:��������[" & szbusID(j) & "]��" & Format(dtREEndStop, "YYYY��MM��DD��") & "����Ʊ" & nSaledSeatCount & "��"
                            End If
                            szQuery = szQuery & vbCrLf & "*��������[" & szbusID(j) & "]" & Format(dtREEndStop, "YYYY��MM��DD��") & szShowMsg & "�ɹ�..."
                            m_oREBus.StopBus dtREStartStop, dtREEndStop, g_bStopAllRefundment
                        End If
NextBus:
                    Next i
                End If
Longdate:
            Next j
        End If
    End If
    SetNormal
    If szErrString <> "" Or szQuery <> "" Then
        MsgBox "�ƻ�" & szShowMsg & ":" & vbCrLf & "����" & szShowMsg & "�ɹ�" & vbCrLf & "����" & szShowMsg & ":" & vbCrLf & szErrString & szQuery, vbInformation, "�ƻ�"
    Else
        MsgBox "�ƻ�����" & szShowMsg & "���", vbInformation, "�ƻ�"
    End If
    ShowSBInfo
    Set oReSheme = Nothing
    Exit Sub
ErrHandle:
    If bEnviroment = True Then
        If optLongStop.Value = True Then
            szErrString = szErrString & "����[" & szbusID(j) & "]" & err.Description & vbCrLf
            Resume Longdate
        Else
            szErrString = szErrString & "����[" & szbusID(j) & "]" & Format(DateAdd("d", i, dtREStartStop), "YYYY��MM��DD�� ") & err.Description & vbCrLf
            Resume NextBus
        End If
    Else
        szErrString = szErrString & "����[" & szbusID(j) & "]" & Format(DateAdd("d", i, dtREStartStop), "YYYY��MM��DD�� ") & err.Description & vbCrLf
        Resume NotUplstTemp:
    End If
End Sub

Private Sub txtBusID_ButtonClick()
On Error GoTo ErrHandle
    Dim oShell As New CommDialog
    Dim aszTmp() As String
    oShell.Init g_oActiveUser
    aszTmp = oShell.SelectBus
    Set oShell = Nothing
    If ArrayLength(aszTmp) = 0 Then Exit Sub
    If aszTmp(1, 1) <> txtBusID.Text Then
        txtBusID.Text = aszTmp(1, 1)
        m_szPlanID = g_szExePriceTable
    
        RefreshBus
        mbShowEnvir = False
    End If
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub
Public Sub ResumeBus()
    Dim nResult As VbMsgBoxResult
    nResult = MsgBox("�Ƿ񸴰೵��[" & txtBusID.Text & "]", vbQuestion + vbYesNo + vbDefaultButton2, "�ƻ�")
    If nResult = vbNo Then Exit Sub
    m_oBus.Identify txtBusID.Text
    m_oBus.BeginStopDate = CDate(cszEmptyDateStr)
    m_oBus.EndStopDate = CDate(cszEmptyDateStr)
    m_oBus.Update
    If m_oBus.BusType = TP_ScrollBus Then
'        lvBus.SelectedItem.SmallIcon = "Flow"
'        lvBus.SelectedItem.ListSubItems(4).ForeColor = vbBlack
'        lvBus.SelectedItem.ListSubItems(4).Text = "����"
    Else
'        lvBus.SelectedItem.SmallIcon = "RunBus"
'        lvBus.SelectedItem.ListSubItems(4).ForeColor = vbBlack
'        lvBus.SelectedItem.ListSubItems(4).Text = "����"
    End If
    nResult = MsgBox("�Ƿ񸴰໷���ڳ���[" & txtBusID.Text & "]", vbQuestion + vbYesNo + vbDefaultButton2, "�ƻ�")
    If nResult = vbNo Then Exit Sub
    optLongStop.Caption = "�ƻ�����"
    optBusStop.Caption = "ʱ��θ���"
'    Label1.Caption = "�ƻ�����"
    Caption = "�������--����"
    Show vbModal
End Sub
'���γ���ͣ��
Private Function SaveBusVehicle()
 
    Dim oEnBus As New REScheme
    m_oBus.Identify txtBusID.Text

    If optLongStop.Value = True Then
        m_oBus.BusVehicleStop m_szVehicle, CDate(cszForeverDateStr), CDate(Now)
    Else
        m_oBus.BusVehicleStop m_szVehicle, dtpEndDate, dtpStartDate
    End If

    If MsgBox("�ƻ����γ���ͣ��ɹ�" & Chr(10) & "�Ƿ�Ӱ�컷����", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
     '�������γ���ͣ��
        oEnBus.StopOrResumBusVehile txtBusID.Text, m_szVehicle, False
        MsgBox "�����г���ͣ��ɹ�!", vbInformation, Me.Caption
    End If


End Function
