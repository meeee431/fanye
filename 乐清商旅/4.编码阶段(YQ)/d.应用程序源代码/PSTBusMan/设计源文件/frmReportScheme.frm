VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Object = "{D11DF87C-EFEC-4838-B7E2-15462BF136FB}#2.1#0"; "RTReportLF.ocx"
Begin VB.Form frmReportScheme 
   BackColor       =   &H80000009&
   Caption         =   "����ϵͳ����"
   ClientHeight    =   6420
   ClientLeft      =   3675
   ClientTop       =   2850
   ClientWidth     =   9030
   Icon            =   "frmReportScheme.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6420
   ScaleWidth      =   9030
   WindowState     =   2  'Maximized
   Begin VB.PictureBox ptResult 
      BackColor       =   &H8000000E&
      Height          =   5835
      Left            =   3105
      ScaleHeight     =   5775
      ScaleWidth      =   5145
      TabIndex        =   10
      Top             =   -75
      Width           =   5205
      Begin RTReportLF.RTReport RTReport 
         Height          =   2880
         Left            =   90
         TabIndex        =   8
         Top             =   1335
         Width           =   4380
         _ExtentX        =   7726
         _ExtentY        =   5080
      End
      Begin VB.PictureBox ptQ 
         BackColor       =   &H8000000E&
         BorderStyle     =   0  'None
         Height          =   1200
         Left            =   -15
         Picture         =   "frmReportScheme.frx":014A
         ScaleHeight     =   1200
         ScaleWidth      =   5100
         TabIndex        =   11
         Top             =   45
         Width           =   5100
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000E&
            BackStyle       =   0  'Transparent
            Caption         =   "XX��ѯ"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   780
            TabIndex        =   15
            Top             =   750
            Width           =   720
         End
         Begin VB.Image imgOpen 
            Height          =   240
            Left            =   0
            Picture         =   "frmReportScheme.frx":1040
            Top             =   0
            Visible         =   0   'False
            Width           =   240
         End
      End
   End
   Begin VB.PictureBox ptQuery 
      BackColor       =   &H8000000E&
      Height          =   5805
      Left            =   0
      ScaleHeight     =   5745
      ScaleWidth      =   2775
      TabIndex        =   9
      Top             =   15
      Width           =   2835
      Begin VB.Frame fraQuery 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1935
         Left            =   105
         TabIndex        =   18
         Top             =   1050
         Width           =   2535
         Begin MSComCtl2.DTPicker dtpQueryDate 
            Height          =   300
            Left            =   0
            TabIndex        =   24
            Top             =   1560
            Width           =   2505
            _ExtentX        =   4419
            _ExtentY        =   529
            _Version        =   393216
            Format          =   60424193
            CurrentDate     =   37854
         End
         Begin VB.ComboBox cboCheck 
            Height          =   300
            Left            =   0
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   960
            Width           =   2505
         End
         Begin VB.ComboBox cboSellStation 
            Height          =   300
            Left            =   0
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   360
            Width           =   2505
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            Caption         =   "ʱ��"
            Height          =   180
            Left            =   0
            TabIndex        =   23
            Top             =   1320
            Width           =   360
         End
         Begin VB.Label Label5 
            BackColor       =   &H00FFFFFF&
            Caption         =   "��Ʊ��"
            Height          =   255
            Left            =   0
            TabIndex        =   20
            Top             =   720
            Width           =   2535
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "�ϳ�վ"
            Height          =   255
            Left            =   0
            TabIndex        =   19
            Top             =   120
            Width           =   2535
         End
      End
      Begin VB.TextBox txtBusID 
         Height          =   300
         Left            =   105
         TabIndex        =   17
         Top             =   2460
         Visible         =   0   'False
         Width           =   2505
      End
      Begin RTComctl3.CoolButton flblClose 
         Height          =   225
         Left            =   2520
         TabIndex        =   13
         Top             =   15
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   397
         BTYPE           =   8
         TX              =   "r"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Marlett"
            Size            =   9
            Charset         =   2
            Weight          =   500
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
         MICON           =   "frmReportScheme.frx":118A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComCtl2.DTPicker dtpEndDate 
         Height          =   300
         Left            =   105
         TabIndex        =   5
         Top             =   1860
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   60424192
         CurrentDate     =   36523
      End
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   300
         Left            =   105
         TabIndex        =   3
         Top             =   1305
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   60424192
         CurrentDate     =   36523
      End
      Begin RTComctl3.CoolButton cmdCancel 
         Cancel          =   -1  'True
         Height          =   330
         Left            =   1395
         TabIndex        =   7
         Top             =   3120
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   582
         BTYPE           =   3
         TX              =   "�ر�(&C)"
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
         MICON           =   "frmReportScheme.frx":11A6
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
         Height          =   330
         Left            =   120
         TabIndex        =   6
         Top             =   3120
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   582
         BTYPE           =   3
         TX              =   "��ѯ"
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
         MICON           =   "frmReportScheme.frx":11C2
         PICN            =   "frmReportScheme.frx":11DE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.ComboBox cboType 
         Height          =   300
         ItemData        =   "frmReportScheme.frx":1578
         Left            =   105
         List            =   "frmReportScheme.frx":1597
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   690
         Width           =   2505
      End
      Begin RTComctl3.FlatLabel FlatLabel1 
         Height          =   285
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483644
         HorizontalAlignment=   1
         Caption         =   "��ѯ�����趨"
      End
      Begin VB.Label lblBusID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����(B):"
         Height          =   180
         Left            =   105
         TabIndex        =   16
         Top             =   2205
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000009&
         Caption         =   "��������(&E):"
         Height          =   315
         Left            =   105
         TabIndex        =   4
         Top             =   1650
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000009&
         Caption         =   "��ʼ����(&S):"
         Height          =   225
         Left            =   105
         TabIndex        =   2
         Top             =   1065
         Width           =   1170
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         Caption         =   "��ѯ����(&T):"
         Height          =   240
         Left            =   105
         TabIndex        =   0
         Top             =   420
         Width           =   1080
      End
   End
   Begin RTComctl3.Spliter spQuery 
      Height          =   1170
      Left            =   2910
      TabIndex        =   12
      Top             =   2445
      Width           =   165
      _ExtentX        =   291
      _ExtentY        =   2064
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmReportScheme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const cszPtTop = 1200

Private lMoveLeft As Long
Private m_nSeatCount As Integer
Private m_nStartSeatNo As Integer
Private F1Book As TTF160Ctl.F1Book

Private Sub cboSellStation_Click()
On Error GoTo ErrHandle
    Dim oBaseInfo As BaseInfo
    Dim aszGateInfo() As String                 '��Ʊ����Ϣ����
    Set oBaseInfo = New BaseInfo
    Dim i As Integer
    oBaseInfo.Init g_oActiveUser
    cboCheck.Clear
    aszGateInfo = oBaseInfo.GetAllCheckGate(, ResolveDisplay(cboSellStation))
    For i = 1 To ArrayLength(aszGateInfo)
        cboCheck.AddItem MakeDisplayString(aszGateInfo(i, 1), aszGateInfo(i, 2))
    Next i
    If cboCheck.ListCount > 0 Then cboCheck.ListIndex = 0
Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub cboType_Change()
'    lblTitle.Caption = "�ƻ�����(&P):"
    Select Case cboType.ListIndex
    Case 0 '0�ƻ�������Ϣ
        dtpStartDate.Enabled = False
        dtpEndDate.Enabled = False
        fraQuery.Visible = False
    Case 1 '1�ƻ����γ�������
        dtpStartDate.Enabled = False
        dtpEndDate.Enabled = False
        fraQuery.Visible = False
    Case 2 '2���������������
        dtpStartDate.Enabled = True
        dtpEndDate.Enabled = True
        fraQuery.Visible = False
    Case 5 '5�ƻ�����ͣ��ͳ��
        dtpStartDate.Enabled = True
        dtpEndDate.Enabled = True
        cmdOk.Enabled = True
        fraQuery.Visible = False
    Case 6 '6��������ͣ��ͳ��
        dtpStartDate.Enabled = True
        dtpEndDate.Enabled = True
        cmdOk.Enabled = True
        fraQuery.Visible = False
        dtpQueryDate.Visible = False
    Case 7 '7�������μӰ�ͳ��
        dtpStartDate.Enabled = True
        dtpEndDate.Enabled = True
        cmdOk.Enabled = True
        fraQuery.Visible = False
    Case 8 '8�������β�����
        dtpEndDate.Enabled = False
        dtpStartDate.Enabled = True
        fraQuery.Visible = False
    Case 9 '9������ȡ�ó�����Ϣ
        dtpStartDate.Enabled = True
        fraQuery.Visible = False
    Case 10 '��վ��
        fraQuery.Visible = False
        dtpStartDate.Enabled = False
        dtpEndDate.Enabled = False
    Case 11 '�г���¼��
        dtpStartDate.Enabled = False
        dtpEndDate.Enabled = False
        fraQuery.Visible = True
        cboCheck.Enabled = True
        FillSellStation '����ϳ�վ
        dtpQueryDate.Value = Date
        dtpQueryDate.Enabled = True
    Case 12 '��ȫ�ż�
        fraQuery.Visible = True
        dtpStartDate.Enabled = False
        dtpEndDate.Enabled = False
        FillSellStation
        cboCheck.Enabled = False
        dtpQueryDate.Value = Date
        dtpQueryDate.Enabled = True
    Case 13 '���г��ƻ�
        fraQuery.Visible = True
        dtpStartDate.Enabled = False
        dtpEndDate.Enabled = False
        FillSellStation
        cboCheck.Enabled = False
        dtpQueryDate.Value = Date
        dtpQueryDate.Enabled = True
    Case 14 '���г��ƻ�
         fraQuery.Visible = True
        dtpStartDate.Enabled = False
        dtpEndDate.Enabled = False
        FillSellStation
        cboCheck.Enabled = False
        dtpQueryDate.Value = Date
        dtpQueryDate.Enabled = False
    Case 15 '������ϱ�
        dtpStartDate.Enabled = False
        dtpEndDate.Enabled = False
        fraQuery.Visible = False
        dtpQueryDate.Value = Date
        dtpQueryDate.Enabled = True
     
    End Select
    If cboType.ListIndex = 7 Then
        lblBusID.Visible = True
        txtBusID.Visible = True
    Else
        lblBusID.Visible = False
        txtBusID.Visible = False
    End If
End Sub

Private Sub cboType_Click()
    cmdOk.Enabled = True
    cboType_Change
    lblTitle.Caption = cboType.Text
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    On Error GoTo ErrorHandle
    SetBusy
    F1Book.DeleteRange F1Book.TopRow, F1Book.LeftCol, F1Book.MaxRow, F1Book.MaxCol, F1ShiftRows
    
    Select Case cboType.ListIndex
    Case 0 '0�ƻ�������Ϣ
        PlanBusInfo
    Case 1 '1�ƻ����γ�������
        PlanBusVehicleInfo
    Case 2 '2���������������
        ReBusInfo
    Case 3 '3��·��Ϣ
        RouteInfo
    Case 4 '4������Ϣ
        VehicleInfo
    Case 5 '5�ƻ�����ͣ��ͳ��
        PlanStopInfo
    Case 6 '6��������ͣ��ͳ��
        ReBusStopInfo
    Case 7 '7�������μӰ�ͳ��
        ReBusAddInfo
    Case 8 '8�������β�����
        ReBusSiltpInfo
    Case 9 '9������ȡ�ó�����Ϣ
        PlanBusVehicleInfo
    Case 10 'վ���ѯ
        StationInfo
    Case 11  '��˾�г���¼��
        CompanyVechileInfo
    Case 12 '��˾��·Ӫ�˳�����ȫ�ż��¼��
        CompanyVechileSafeInfo
    Case 13 '�շ�����ҵ�ƻ�
        CompanyDayWorkPlan
    Case 14 '�ܷ�����ҵ�ƻ�
        CompanyWorkPlan
    Case 15 '������ϱ�
        BusInfo
    End Select
    RTReport.SetFocus
    SetNormal
    ShowSBInfo ""
'    ShowSBInfo "����" & m_lRange & "����¼", ESB_ResultCountInfo
Exit Sub
ErrorHandle:
    ShowSBInfo ""
    SetNormal
    ShowErrorMsg
End Sub


Private Sub flblClose_Click()
    ptQuery.Visible = False
    imgOpen.Visible = True
    lMoveLeft = 240
    spQuery.LayoutIt
End Sub

Private Sub Form_Activate()
    MDIScheme.SetPrintEnabled True
End Sub
Private Sub Form_Deactivate()
    MDIScheme.SetPrintEnabled False
End Sub

Private Sub Form_Load()
    spQuery.InitSpliter ptQuery, ptResult
    lMoveLeft = 0
    FillQueryType
    dtpStartDate.Value = Date
    dtpEndDate.Value = Date
    Set F1Book = RTReport.CellObject
    F1Book.ShowColHeading = True
    F1Book.ShowRowHeading = True
End Sub

Private Sub FillQueryType()
    cboType.Clear
    cboType.AddItem "�ƻ�������Ϣ"
    cboType.AddItem "�ƻ����γ�������"
    cboType.AddItem "���������������"
    cboType.AddItem "��·��Ϣ"
    cboType.AddItem "������Ϣ"
    cboType.AddItem "�ƻ�����ͣ��ͳ��"
    cboType.AddItem "��������ͣ��ͳ��"
    cboType.AddItem "�������μӰ�ͳ��" '�񻷼���
    cboType.AddItem "�������β�����"
    cboType.AddItem "������ȡ�ó�����Ϣ"
    cboType.AddItem "վ����Ϣ"
    
'*******************�������**********************
    cboType.AddItem "��˾�г���¼��"
    cboType.AddItem "��˾��·Ӫ�˳�����ȫ�ż��¼��"
    cboType.AddItem "�շ�����ҵ�ƻ�"
    cboType.AddItem "�·�����ҵ�ƻ�"
'**************************************************

    cboType.AddItem "������ϱ�" '�񻷼���
    
    cboType.ListIndex = 0
    
End Sub
Private Sub Form_Resize()
    spQuery.LayoutIt
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIScheme.SetPrintEnabled False
End Sub

Private Sub imgOpen_Click()
    ptQuery.Visible = True
    imgOpen.Visible = False
    lMoveLeft = 0
    spQuery.LayoutIt
End Sub


Private Sub ptResult_Resize()
    Dim lTemp As Long
    lTemp = IIf((ptResult.ScaleHeight - cszPtTop) <= 0, lTemp, ptResult.ScaleHeight - cszPtTop)
    RTReport.Move 0, cszPtTop, ptResult.ScaleWidth, lTemp
End Sub

Private Sub PlanBusVehicleInfo()
    Dim oPlan As New BusProject
    Dim oBaseInfo As New BaseInfo
    Dim szTemp() As TVehcileSeatType
    Dim rsTemp As Recordset
    Dim i As Integer
    Dim vData As Variant
On Error GoTo ErrorHandle
    ShowSBInfo "��ó��γ���..."
    oPlan.Init g_oActiveUser
    oBaseInfo.Init g_oActiveUser

    If cboType.ListIndex = 8 Then
        Dim oSheme As New RegularScheme
        oPlan.Identify
        Set oSheme = Nothing
        Set rsTemp = oPlan.GetBusVehicleReport(g_szExePriceTable)
    Else
        oPlan.Identify
        Set rsTemp = oPlan.GetAllBusVehicleReport
    End If
    F1Book.MaxCol = 19
    F1Book.MaxRow = rsTemp.RecordCount + 1
    WriteProcessBar , , rsTemp.RecordCount
    '����ͷ
    F1Book.TextRC(1, 1) = "����"
    F1Book.TextRC(1, 2) = "��·"
    F1Book.TextRC(1, 3) = "����ʱ��"
    F1Book.TextRC(1, 4) = "��Ʊ��"
    F1Book.TextRC(1, 5) = "��������"
    F1Book.TextRC(1, 6) = "ѭ������"
    F1Book.TextRC(1, 7) = "��ʼ���"
    F1Book.TextRC(1, 8) = "�������"
    F1Book.TextRC(1, 9) = "��������"
    F1Book.TextRC(1, 10) = "����"
    F1Book.TextRC(1, 11) = "����"
    F1Book.TextRC(1, 12) = "��λ��"
    F1Book.TextRC(1, 13) = "��ʼ����"
    F1Book.TextRC(1, 14) = "��λ����"
    F1Book.TextRC(1, 15) = "���˹�˾"
    F1Book.TextRC(1, 16) = "���ʹ�˾"
    F1Book.TextRC(1, 17) = "����"
    F1Book.TextRC(1, 18) = "ͣ�࿪ʼ����"
    F1Book.TextRC(1, 19) = "ͣ���������"
    For i = 2 To rsTemp.RecordCount + 1
        WriteProcessBar , i - 1, rsTemp.RecordCount
        F1Book.TextRC(i, 1) = Trim(rsTemp!bus_id)
        F1Book.TextRC(i, 2) = Trim(rsTemp!route_name)
        F1Book.TextRC(i, 3) = Format(rsTemp!bus_start_time, "HH:MM")
        F1Book.TextRC(i, 4) = Trim(rsTemp!check_gate_name)
'        If rsTemp!bus_type = TP_ScrollBus Then
'            f1book.TextRC i,5, "����"
'        Else
'            f1book.TextRC i,5, "�̶�"
'        End If
        F1Book.TextRC(i, 5) = Trim(rsTemp!bus_type_name)
        F1Book.TextRC(i, 6) = Trim(rsTemp!bus_run_cycle)
        F1Book.TextRC(i, 7) = Trim(rsTemp!run_start_serial)
        'run_start_serial
        F1Book.TextRC(i, 8) = Trim(rsTemp!vehicle_serial)
       F1Book.TextRC(i, 9) = Trim(rsTemp!vehicle_id)
        F1Book.TextRC(i, 10) = Trim(rsTemp!license_tag_no)
        F1Book.TextRC(i, 11) = Trim(rsTemp!vehicle_type_short_name)


        m_nSeatCount = Trim(rsTemp!seat_quantity)
        If rsTemp!sale_stand_ticket_quantity = 0 Then
            F1Book.TextRC(i, 12) = Trim(rsTemp!seat_quantity)
        Else
            F1Book.TextRC(i, 12) = Trim(rsTemp!seat_quantity) & "(" & Trim(rsTemp!sale_stand_ticket_quantity) & ")"
        End If
        m_nStartSeatNo = Val((rsTemp!start_seat_no))
        F1Book.TextRC(i, 13) = Trim(rsTemp!start_seat_no)
        szTemp = oBaseInfo.GetAllVehicleSeatTypeInfo(Trim(rsTemp!vehicle_id))
        F1Book.TextRC(i, 14) = FindSetSeatInfo(szTemp)
        F1Book.TextRC(i, 15) = Trim(rsTemp!transport_company_short_name)
        F1Book.TextRC(i, 16) = Trim(rsTemp!split_company_short_name)
        F1Book.TextRC(i, 17) = Trim(rsTemp!owner_name)
        If rsTemp!stop_start_date = CDate(cszEmptyDateStr) Then
            F1Book.TextRC(i, 18) = ""
            F1Book.TextRC(i, 19) = ""
        Else
            F1Book.TextRC(i, 18) = Format(rsTemp!stop_start_date, "YYYY-MM-DD")
            F1Book.TextRC(i, 19) = Format(rsTemp!stop_end_date, "YYYY-MM-DD")
        End If

        If rsTemp!Status = 1 Then
            vData = F1Book.TextRC(i, 18)
            If vData <> "" Then
                vData = vData & "�ҳ���ͣ"
            Else
                vData = vData & "����ͣ"
            End If
            F1Book.TextRC(i, 17) = vData
            vData = F1Book.TextRC(i, 19)
            If vData <> "" Then
                vData = vData & "�ҳ���ͣ"
            Else
                vData = vData & "����ͣ"
            End If
            vData = F1Book.TextRC(i, 19)
        End If

    rsTemp.MoveNext
    Next
    WriteProcessBar False
    
Exit Sub
ErrorHandle:
    SetNormal
    ShowErrorMsg
End Sub

Private Sub PlanBusInfo()
    Dim oPlan As New BusProject
    Dim rsTemp As Recordset
    Dim i As Integer
On Error GoTo ErrorHandle
    ShowSBInfo "��ó�����Ϣ..."
    oPlan.Init g_oActiveUser
    oPlan.Identify
    Set rsTemp = oPlan.GetAllBusReport
    F1Book.MaxCol = 9
    F1Book.MaxRow = rsTemp.RecordCount + 1
    WriteProcessBar , , rsTemp.RecordCount
    F1Book.TextRC(1, 1) = "����"
    F1Book.TextRC(1, 2) = "��·"
    F1Book.TextRC(1, 3) = "����ʱ��"
    F1Book.TextRC(1, 4) = "��Ʊ��"
    F1Book.TextRC(1, 5) = "��������"
    F1Book.TextRC(1, 6) = "��������"
    F1Book.TextRC(1, 7) = "��ʼ���"
    F1Book.TextRC(1, 8) = "ͣ�࿪ʼ����"
    F1Book.TextRC(1, 9) = "ͣ���������"
    F1Book.ColWidth(8) = 3000
    F1Book.ColWidth(9) = 3000
    For i = 2 To rsTemp.RecordCount + 1
        WriteProcessBar , i - 1, rsTemp.RecordCount
        F1Book.TextRC(i, 1) = Trim(rsTemp!bus_id)
        F1Book.TextRC(i, 2) = Trim(rsTemp!route_name)
        F1Book.TextRC(i, 3) = Format(rsTemp!bus_start_time, "HH:MM")
        F1Book.TextRC(i, 4) = Trim(rsTemp!check_gate_name)
        If rsTemp!bus_type = TP_ScrollBus Then
            F1Book.TextRC(i, 5) = "����"
        Else
            F1Book.TextRC(i, 5) = "�̶�"
        End If
        F1Book.TextRC(i, 6) = Trim(rsTemp!bus_run_cycle)
        F1Book.TextRC(i, 7) = Trim(rsTemp!run_start_serial)
        If rsTemp!stop_start_date = CDate(cszEmptyDateStr) Then
            F1Book.TextRC(i, 8) = ""
            F1Book.TextRC(i, 9) = ""
        Else
            F1Book.TextRC(i, 8) = Format(rsTemp!stop_start_date, "YYYY-MM-DD")
            F1Book.TextRC(i, 9) = Format(rsTemp!stop_end_date, "YYYY-MM-DD")
        End If
    rsTemp.MoveNext
    Next
    WriteProcessBar False
Exit Sub
ErrorHandle:
    SetNormal
    ShowErrorMsg
End Sub

'������ϱ�
Private Sub BusInfo()
On Error GoTo ErrHandle
    Dim oRScheme As New REScheme
    Dim rsTemp As Recordset
    Dim aszTmp As Variant
    ReDim aszTmp(1 To 1, 1 To 2)
    aszTmp(1, 1) = "����"
    aszTmp(1, 2) = Format(dtpQueryDate.Value, "yyyy��MM��dd��")
    Set rsTemp = oRScheme.GetBusInfo()
    ShowReport rsTemp, "������ϱ�.xls", "������ϱ�", aszTmp
Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub ReBusInfo()
    Dim oRScheme As New REScheme
    Dim rsTemp As Recordset
    Dim i As Integer
    Dim szTemp As String
On Error GoTo ErrorHandle
    ShowSBInfo "��û�������..."
    oRScheme.Init g_oActiveUser
    Set rsTemp = oRScheme.GetREBusReport(dtpStartDate.Value, dtpEndDate.Value)
    F1Book.MaxCol = 10
    F1Book.MaxRow = rsTemp.RecordCount + 1
    WriteProcessBar , , rsTemp.RecordCount
    F1Book.TextRC(1, 1) = "��������"
    F1Book.TextRC(1, 2) = "���δ���"
    F1Book.TextRC(1, 3) = "��·����"
    F1Book.TextRC(1, 4) = "����ʱ��"
    F1Book.TextRC(1, 5) = "��Ʊ��"
    F1Book.TextRC(1, 6) = "��������"
    F1Book.TextRC(1, 7) = "����״̬"
    F1Book.TextRC(1, 8) = "����"
    F1Book.TextRC(1, 9) = "���˹�˾"
    F1Book.TextRC(1, 10) = "����"
    For i = 2 To rsTemp.RecordCount + 1
        WriteProcessBar , i - 1, rsTemp.RecordCount
        F1Book.TextRC(i, 1) = Format(rsTemp!bus_date, "YYYY-MM-DD")
        F1Book.TextRC(i, 2) = Trim(rsTemp!bus_id)
        F1Book.TextRC(i, 3) = Trim(rsTemp!route_name)
        F1Book.TextRC(i, 4) = Format(rsTemp!bus_start_time, "HH:mm")
        F1Book.TextRC(i, 5) = Trim(rsTemp!check_gate_id)
        If rsTemp!bus_type = TP_ScrollBus Then
            F1Book.TextRC(i, 6) = "����"
        Else
            F1Book.TextRC(i, 6) = "�̶�"
        End If
        Select Case rsTemp!Status
        Case ST_BusChecking
            szTemp = "����"
        Case ST_BusMergeStopped
            szTemp = "����"
        Case ST_BusNormal
            szTemp = "��ͨ"
        Case ST_BusExtraChecking
            szTemp = "����"
        Case ST_BusStopCheck
            szTemp = "ͣ��"
        Case ST_BusStopped
            szTemp = "ͣ��"
        End Select
        F1Book.TextRC(i, 7) = szTemp
        F1Book.TextRC(i, 8) = Trim(rsTemp!license_tag_no)
        F1Book.TextRC(i, 9) = Trim(rsTemp!transport_company_short_name)
        F1Book.TextRC(i, 10) = Trim(rsTemp!owner_name)
        rsTemp.MoveNext
    Next
    WriteProcessBar False
Exit Sub
ErrorHandle:
    SetNormal
    ShowErrorMsg
End Sub

Private Sub RouteInfo()
    Dim oBase As New BaseInfo
    Dim aszTemp() As String
    Dim nCount As Integer, i As Integer
On Error GoTo ErrorHandle
    SetBusy
    ShowSBInfo "�����·��Ϣ..."
    oBase.Init g_oActiveUser
    aszTemp = oBase.GetRouteEx
    nCount = ArrayLength(aszTemp)
    If nCount = 0 Then Exit Sub
    F1Book.MaxCol = 6
    F1Book.MaxRow = nCount + 1
    WriteProcessBar , nCount, , True
    F1Book.TextRC(1, 1) = "��·����"
    F1Book.TextRC(1, 2) = "��·����"
    F1Book.TextRC(1, 3) = ";��վ"
    F1Book.TextRC(1, 4) = "�յ�վ"
    F1Book.TextRC(1, 5) = "״̬"
    For i = 1 To nCount
        WriteProcessBar , i, nCount, "�����·��Ϣ..."
        F1Book.TextRC(i + 1, 1) = aszTemp(i, 1)
        F1Book.TextRC(i + 1, 2) = aszTemp(i, 2)
        F1Book.TextRC(i + 1, 3) = aszTemp(i, 3)
        F1Book.TextRC(i + 1, 4) = aszTemp(i, 4)
        F1Book.TextRC(i + 1, 5) = aszTemp(i, 5)
    Next
    WriteProcessBar False
Exit Sub
ErrorHandle:
    SetNormal
    ShowErrorMsg
End Sub

Private Sub StationInfo()
    Dim aszStation() As String
    Dim szTemp As String
    Dim oBaseInfo As BaseInfo
    Dim i As Integer
    Dim nCount As Integer
    On Error GoTo ErrorHandle
    SetBusy
    Set oBaseInfo = New BaseInfo
    oBaseInfo.Init g_oActiveUser
    aszStation = oBaseInfo.GetStation()
    nCount = ArrayLength(aszStation)
    F1Book.MaxCol = 5
    F1Book.MaxRow = nCount + 1
    F1Book.TextRC(1, 1) = "վ�����"
    F1Book.TextRC(1, 2) = "վ������"
    F1Book.TextRC(1, 3) = "������"
    F1Book.TextRC(1, 4) = "�Ƿ����"
'    F1Book.TextRC(1, 5) = "������"
    F1Book.TextRC(1, 5) = "����"
    For i = 1 To nCount
        F1Book.TextRC(i + 1, 1) = aszStation(i, 1)
        F1Book.TextRC(i + 1, 2) = aszStation(i, 2)
        F1Book.TextRC(i + 1, 3) = aszStation(i, 3)
        If Val(aszStation(i, 4)) <> TP_CanSellTicket Then
            szTemp = "������"
            
        Else
            szTemp = "����"
        End If
        F1Book.TextRC(i + 1, 4) = szTemp
'        F1Book.TextRC(i + 1, 5) = aszStation(i, 5)
        F1Book.TextRC(i + 1, 5) = aszStation(i, 6)
    Next
    
    SetNormal
Exit Sub
ErrorHandle:
    SetNormal
    ShowErrorMsg

End Sub

Private Sub VehicleInfo()
    Dim oBase As New BaseInfo
    Dim aszTemp() As String
    Dim nCount As Integer, i As Integer
On Error GoTo ErrorHandle
    ShowSBInfo "��ó�����Ϣ..."
    oBase.Init g_oActiveUser
    aszTemp = oBase.GetVehicle
    nCount = ArrayLength(aszTemp)
    If nCount = 0 Then Exit Sub
    F1Book.MaxCol = 8
    F1Book.MaxRow = nCount + 1
    i = 1
    WriteProcessBar , , nCount
    F1Book.TextRC(i, 1) = "��������"
    F1Book.TextRC(i, 2) = "����"
    F1Book.TextRC(i, 3) = "��λ��"
    F1Book.TextRC(i, 4) = "���˹�˾"
    F1Book.TextRC(i, 5) = "����"
    F1Book.TextRC(i, 6) = "���ʹ���"
    F1Book.TextRC(i, 7) = "����"
    F1Book.TextRC(i, 8) = "ע��"
    For i = 1 To nCount
        WriteProcessBar , i, nCount
        F1Book.TextRC(i + 1, 1) = aszTemp(i, 1)
        F1Book.TextRC(i + 1, 2) = aszTemp(i, 2)
        F1Book.TextRC(i + 1, 3) = aszTemp(i, 3)
        F1Book.TextRC(i + 1, 4) = aszTemp(i, 4)
        F1Book.TextRC(i + 1, 5) = aszTemp(i, 5)
        F1Book.TextRC(i + 1, 6) = aszTemp(i, 7)
        F1Book.TextRC(i + 1, 7) = aszTemp(i, 8)
        F1Book.TextRC(i + 1, 8) = aszTemp(i, 9)
    Next
    WriteProcessBar False
Exit Sub
ErrorHandle:
    SetNormal
    ShowErrorMsg
End Sub




Private Sub txtPlanID_Click()
    Dim oShell As New STShell.CommDialog
    Dim aszTemp() As String
    On Error GoTo ErrorHandle
    oShell.Init g_oActiveUser
    Select Case cboType.ListIndex
    Case 7 '��������
        aszTemp = oShell.SelectREBus(dtpEndDate.Value)
        If ArrayLength(aszTemp) = 0 Then Exit Sub
    Case 8 '������ȡ�ĳ���
        aszTemp = oShell.SelectArea()
        If ArrayLength(aszTemp) = 0 Then Exit Sub
    Case Else
'        aszTemp = oShell.selectProject()
'        If ArrayLength(aszTemp) = 0 Then Exit Sub
'        txtPlanID.Text = aszTemp(1, 1) & "[" & Trim(aszTemp(1, 2)) & "]"
    End Select
    Set oShell = Nothing
    Exit Sub
ErrorHandle:
    Set oShell = Nothing
    SetNormal
    ShowErrorMsg
End Sub

Public Sub PrintReport(Optional pbShowDialog As Boolean = False)
    On Error GoTo ErrorHandle
    RTReport.PrintReport pbShowDialog
    Exit Sub
ErrorHandle:
End Sub

Public Sub PreView()
    RTReport.PrintView
End Sub

Public Sub PageSet()
    RTReport.OpenDialog EDialogType.PAGESET_TYPE
End Sub

Public Sub PrintSet()
    RTReport.OpenDialog EDialogType.PRINT_TYPE
End Sub
'�����ļ�
Public Sub ExportFile()
    Dim szFileName As String
    szFileName = RTReport.OpenDialog(EDialogType.EXPORT_FILE)
End Sub
'�����ļ�����
Public Sub ExportFileOpen()
    Dim szFileName As String
    szFileName = RTReport.OpenDialog(EDialogType.EXPORT_FILE)
    If szFileName <> "" Then
        OpenLinkedFile szFileName
    End If
End Sub

Public Sub PlanStopInfo()
    Dim oRScheme As New REScheme
    Dim rsTemp As Recordset
    Dim i As Integer
    Dim szTemp As String
On Error GoTo ErrorHandle
    oRScheme.Init g_oActiveUser
    ShowSBInfo "��ó���ͣ���¼..."
    Set rsTemp = oRScheme.GetPlanStopReport(dtpStartDate.Value, dtpEndDate.Value)
    F1Book.MaxCol = 6
    F1Book.MaxRow = rsTemp.RecordCount + 1
    WriteProcessBar , , rsTemp.RecordCount
    i = 1
    F1Book.TextRC(i, 1) = "����"
    F1Book.TextRC(i, 2) = "��·����"
    F1Book.TextRC(i, 3) = "����ʱ��"
    F1Book.TextRC(i, 4) = "Ӧ�����"
    F1Book.TextRC(i, 5) = "ʵ�ʰ��"
    F1Book.TextRC(i, 6) = "������"
    F1Book.TextRC(i, 7) = "��������"
    For i = 2 To rsTemp.RecordCount + 1
        WriteProcessBar , i - 1, rsTemp.RecordCount
        F1Book.TextRC(i, 1) = Trim(rsTemp!bus_id)
        F1Book.TextRC(i, 2) = Trim(rsTemp!route_name)
        F1Book.TextRC(i, 3) = Format(rsTemp!bus_start_time, "HH:MM")
        F1Book.TextRC(i, 4) = Trim(rsTemp!Count)
        F1Book.TextRC(i, 5) = Trim(rsTemp!Count - rsTemp!stop_count)
        F1Book.TextRC(i, 6) = Format((rsTemp!Count - rsTemp!stop_count) / rsTemp!Count, "00%")
        F1Book.TextRC(i, 7) = Trim(rsTemp!bus_run_cycle)
        rsTemp.MoveNext
    Next
    WriteProcessBar False
Exit Sub
ErrorHandle:
    SetNormal
    ShowErrorMsg
End Sub

Public Sub ReBusStopInfo()
    Dim oRScheme As New REScheme
    Dim rsTemp As Recordset
    Dim i As Integer
    Dim szTemp As String
On Error GoTo ErrorHandle
    oRScheme.Init g_oActiveUser
    ShowSBInfo "��û�������ͣ��..."
    Set rsTemp = oRScheme.GetREBusStopReport(dtpStartDate.Value, dtpEndDate.Value)
    F1Book.MaxCol = 6
    If rsTemp Is Nothing Then Exit Sub
    F1Book.MaxRow = rsTemp.RecordCount + 1
    WriteProcessBar , rsTemp.RecordCount + 1, , True
    i = 1
    F1Book.TextRC(i, 1) = "����"
    F1Book.TextRC(i, 2) = "����"
    F1Book.TextRC(i, 3) = "���˹�˾"
    F1Book.TextRC(i, 4) = "����"
    F1Book.TextRC(i, 5) = "ͣ����"
    F1Book.TextRC(i, 6) = "��·����"
    F1Book.TextRC(i, 7) = "����ʱ��"
    For i = 2 To rsTemp.RecordCount + 1
        WriteProcessBar , i - 1, rsTemp.RecordCount
        F1Book.TextRC(i, 1) = Trim(rsTemp!bus_id)
        F1Book.TextRC(i, 2) = Trim(rsTemp!license_tag_no)
        F1Book.TextRC(i, 3) = Trim(rsTemp!company_name)
        F1Book.TextRC(i, 4) = Trim(rsTemp!owner_name)
        F1Book.TextRC(i, 5) = Trim(rsTemp!stop_count)
        F1Book.TextRC(i, 6) = Trim(rsTemp!route_name)
        F1Book.TextRC(i, 7) = Format(rsTemp!bus_start_time, "HH:mm")
        rsTemp.MoveNext
    Next
'    F1Book.DoRedrawAll
    WriteProcessBar False
Exit Sub
ErrorHandle:
    SetNormal
    ShowErrorMsg
End Sub

Public Sub ReBusAddInfo()
    Dim oRScheme As New REScheme
    Dim rsTemp As Recordset
    Dim i As Integer
    Dim szTemp As String
On Error GoTo ErrorHandle
    oRScheme.Init g_oActiveUser
    ShowSBInfo "��û������μӰ�..."
    Set rsTemp = oRScheme.GetREBusAddReport(dtpStartDate.Value, dtpEndDate.Value)
    F1Book.MaxCol = 6
    If rsTemp Is Nothing Then Exit Sub
    F1Book.MaxRow = rsTemp.RecordCount + 1
    WriteProcessBar , rsTemp.RecordCount + 1, , True
    i = 1
    F1Book.TextRC(i, 1) = "����"
    F1Book.TextRC(i, 2) = "����"
    F1Book.TextRC(i, 3) = "���˹�˾"
    F1Book.TextRC(i, 4) = "����"
    F1Book.TextRC(i, 5) = "�Ӱ���"
    F1Book.TextRC(i, 6) = "��·����"
    F1Book.TextRC(i, 7) = "����ʱ��"
    For i = 2 To rsTemp.RecordCount + 1
        WriteProcessBar , i - 1, rsTemp.RecordCount
        F1Book.TextRC(i, 1) = Trim(rsTemp!bus_id)
        F1Book.TextRC(i, 2) = Trim(rsTemp!license_tag_no)
        F1Book.TextRC(i, 3) = Trim(rsTemp!company_name)
        F1Book.TextRC(i, 4) = Trim(rsTemp!owner_name)
        F1Book.TextRC(i, 5) = Trim(rsTemp!add_count)
        F1Book.TextRC(i, 6) = Trim(rsTemp!route_name)
        F1Book.TextRC(i, 7) = Format(rsTemp!bus_start_time, "HH:mm")
        rsTemp.MoveNext
    Next
'    F1Book.DoRedrawAll
    WriteProcessBar False
Exit Sub
ErrorHandle:
    SetNormal
    ShowErrorMsg
End Sub

Public Sub ReBusSiltpInfo()
Dim oRebus As New REBus
    Dim rsTemp As Recordset
    Dim i As Integer
    Dim szTemp() As String
    Dim nCountInfo As Integer
On Error GoTo ErrorHandle
    oRebus.Init g_oActiveUser
    ShowSBInfo "��û�����ֳ���..."
    
    szTemp = oRebus.GetSlitpInfo(txtBusID.Text, dtpStartDate.Value)
    nCountInfo = ArrayLength(szTemp)
    F1Book.MaxCol = 10

    F1Book.MaxRow = nCountInfo + 1
    WriteProcessBar , , nCountInfo
    i = 1
    F1Book.TextRC(i, 1) = "Ŀ�공��"
    F1Book.TextRC(i, 2) = "��ֳ���"
    F1Book.TextRC(i, 3) = "��ǰ���"
    F1Book.TextRC(i, 4) = "ԭ���"
    F1Book.TextRC(i, 5) = "Ʊ��"
    F1Book.TextRC(i, 6) = "��Ʊ��"
    F1Book.TextRC(i, 7) = "����"
    F1Book.TextRC(i, 8) = "����λ"
    F1Book.TextRC(i, 9) = "�յ�վ"
    F1Book.TextRC(i, 10) = "����ʱ��"
    For i = 2 To nCountInfo + 1
        WriteProcessBar , i - 1, nCountInfo
        F1Book.TextRC(i, 1) = szTemp(i - 1, 1)
        F1Book.TextRC(i, 2) = g_szExePriceTable
        F1Book.TextRC(i, 3) = szTemp(i - 1, 2)
        F1Book.TextRC(i, 4) = szTemp(i - 1, 3)
        F1Book.TextRC(i, 5) = szTemp(i - 1, 4)
        F1Book.TextRC(i, 6) = szTemp(i - 1, 5)
        F1Book.TextRC(i, 7) = szTemp(i - 1, 6)
        F1Book.TextRC(i, 8) = szTemp(i - 1, 7)
        F1Book.TextRC(i, 9) = szTemp(i - 1, 8)
        F1Book.TextRC(i, 10) = Format(szTemp(i - 1, 9))
    Next
'    F1Book.DoRedrawAll
    WriteProcessBar False
Exit Sub
ErrorHandle:
    SetNormal
    ShowErrorMsg
End Sub

Private Function FindSetSeatInfo(szTemp() As TVehcileSeatType) As String
    Dim nCount As Integer
    Dim i As Integer
    Dim seatInfo As String
    Dim seatInfoTemp As String
    Dim sz As String
    Dim nCountSeat  As Integer
    nCount = ArrayLength(szTemp)
    sz = ","
    If nCount = 0 Then
        seatInfo = "ȫ����ͨ"
    Else
        For i = 1 To nCount
            If i <> 1 Then
                seatInfo = seatInfo & sz
            End If
            If szTemp(i).szStartSeatNo <= szTemp(i).szEndSeatNo Then
                nCountSeat = nCountSeat - CInt(szTemp(i).szStartSeatNo) + CInt(szTemp(i).szEndSeatNo) + 1
            Else
                nCountSeat = nCountSeat + CInt(szTemp(i).szStartSeatNo) - CInt(szTemp(i).szEndSeatNo) + 1
            End If
            seatInfo = seatInfo & Trim(CStr(szTemp(i).szStartSeatNo)) & "��" & Trim(CStr(szTemp(i).szEndSeatNo)) & " " & Trim(szTemp(i).szSeatTypeName)
        Next
    End If
    If m_nSeatCount > nCountSeat Then
        seatInfo = seatInfo & sz & "������ͨ"
    End If
    FindSetSeatInfo = seatInfo
End Function

'��˾�г���¼��
Private Sub CompanyVechileInfo()
    Dim g_oREScheme As New REScheme       '��ǰ��Ʊ����
    '�õ��������г�����Ϣ
    Dim rsBusInfo As Recordset
    
    g_oREScheme.Init g_oActiveUser
    Set rsBusInfo = g_oREScheme.GetBusInfoRsReport(dtpQueryDate.Value, ResolveDisplay(cboCheck.Text), ResolveDisplay(cboSellStation))
    Dim aszTmp() As Variant
    ReDim aszTmp(1 To 3, 1 To 2)
    aszTmp(1, 1) = "��Ʊ��"
    aszTmp(1, 2) = ResolveDisplay(cboCheck)
    aszTmp(2, 1) = "����"
    aszTmp(2, 2) = Format(dtpQueryDate.Value, "yyyy��MM��dd��")
    aszTmp(3, 1) = "����"
    aszTmp(3, 2) = WeekdayName(Weekday(dtpQueryDate.Value))
    ShowReport rsBusInfo, "��˾�г���¼��.xls", "��˾�г���¼��", aszTmp
End Sub
'

'�õ�����˾��·Ӫ�˳�����ȫ�ż��¼��
Private Sub CompanyVechileSafeInfo()
On Error GoTo ErrHandle
Dim moScheme As New REScheme
Dim aszEnvBus() As Variant
Dim i As Integer
Dim rsTemp As New Recordset
Dim rsTmp As New Recordset
Dim szFindString As String


Dim aszTmp As Variant
ReDim aszTmp(1 To 1, 1 To 2)
aszTmp(1, 1) = "����"
aszTmp(1, 2) = Format(dtpQueryDate.Value, "yyyy��MM��dd��")
szFindString = " AND ebi.status= 0 "
Set rsTemp = moScheme.GetRESellStationBusReport(ResolveDisplay(cboSellStation), dtpQueryDate.Value, szFindString)
ShowReport rsTemp, "��˾��·Ӫ�˰�ȫ�ż��¼��.xls", "��˾��·Ӫ�˰�ȫ�ż��¼��", aszTmp

Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub
'��˾���г��ƻ�
Private Sub CompanyDayWorkPlan()
On Error GoTo ErrHandle
    Dim g_oREScheme As New REScheme       '��ǰ��Ʊ����
    '�õ��������г�����Ϣ
    Dim rsBusInfo As Recordset
    Dim aszTmp As Variant
    ReDim aszTmp(1 To 1, 1 To 2)
    aszTmp(1, 1) = "����"
    aszTmp(1, 2) = Format(dtpQueryDate.Value, "yyyy��MM��dd��")
    Set rsBusInfo = g_oREScheme.GetRESellStationBusReport(ResolveDisplay(cboSellStation), dtpQueryDate.Value)
    ShowReport rsBusInfo, "�շ�����ҵ�ƻ�.xls", "�շ�����ҵ�ƻ�", aszTmp
Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub
'��˾���г��ƻ�
Private Sub CompanyWorkPlan()
On Error GoTo ErrHandle
    Dim g_oREScheme As New REScheme       '��ǰ��Ʊ����
    '�õ����г�����Ϣ
    Dim rsBusInfo As Recordset

    Dim g_oTicketPriceMan As New TicketPriceMan
    Dim szFindString As String
    
    szFindString = " AND  bpl.price_table_id= '" & g_szExePriceTable & "'"
    
    
    Set rsBusInfo = g_oREScheme.GetPlanSellStationBusReport(ResolveDisplay(cboSellStation), szFindString)
    ShowReport rsBusInfo, "�·�����ҵ�ƻ�.xls", "�·�����ҵ�ƻ�"
Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub


'����ϳ�վ
Private Sub FillSellStation()
    Dim i As Integer
    cboSellStation.Clear
    For i = 1 To ArrayLength(g_atAllSellStation)
        cboSellStation.AddItem MakeDisplayString(g_atAllSellStation(i).szSellStationID, g_atAllSellStation(i).szSellStationFullName)
    Next i
    If cboSellStation.ListCount > 0 Then cboSellStation.ListIndex = 0
End Sub

Public Function ShowReport(prsData As Recordset, pszFileName As String, pszCaption As String, Optional pvaCustomData As Variant, Optional pnReportType As Integer = 0) As Long
    On Error GoTo Error_Handle
    Me.Caption = pszCaption
    WriteProcessBar True, , , "�����γɱ���..."
    RTReport.SheetTitle = ""

    RTReport.TemplateFile = App.Path & "\" & pszFileName
    RTReport.LeftLabelVisual = True
    RTReport.TopLabelVisual = True
    RTReport.ShowReport prsData, pvaCustomData
    WriteProcessBar False, , , ""
    ShowSBInfo "��" & prsData.RecordCount & "����¼", ESB_ResultCountInfo
    Exit Function
Error_Handle:
    ShowErrorMsg
End Function
