VERSION 5.00
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Begin VB.Form frmSection 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "·��"
   ClientHeight    =   2280
   ClientLeft      =   3555
   ClientTop       =   4005
   ClientWidth     =   6960
   HelpContextID   =   10000390
   Icon            =   "frmSection.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "ʹ���˴�·�ε���·"
      ForeColor       =   &H80000008&
      Height          =   1875
      Left            =   30
      TabIndex        =   21
      Top             =   2310
      Width           =   6675
      Begin VB.ListBox lstRoute 
         Appearance      =   0  'Flat
         Height          =   1470
         Left            =   105
         TabIndex        =   22
         Top             =   270
         Width           =   6465
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "ʹ���˴�·�ε���·:"
         Height          =   180
         Left            =   105
         TabIndex        =   24
         Top             =   30
         Width           =   1800
      End
   End
   Begin RTComctl3.CoolButton cmdRoute 
      Height          =   315
      Left            =   5610
      TabIndex        =   19
      Top             =   1260
      Width           =   1185
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "��·>>"
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
      MICON           =   "frmSection.frx":014A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtPathNo 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   300
      Left            =   3960
      TabIndex        =   15
      Text            =   "1"
      Top             =   1785
      Width           =   795
   End
   Begin VB.TextBox txtKm 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1200
      MaxLength       =   10
      TabIndex        =   13
      Text            =   "0"
      Top             =   1785
      Width           =   1065
   End
   Begin VB.TextBox txtSectionName 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1200
      MaxLength       =   10
      TabIndex        =   3
      Top             =   555
      Width           =   4275
   End
   Begin RTComctl3.CoolButton cmdOk 
      Default         =   -1  'True
      Height          =   315
      Left            =   5610
      TabIndex        =   16
      Top             =   135
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
      MICON           =   "frmSection.frx":0166
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
      Left            =   5610
      TabIndex        =   17
      Top             =   525
      Width           =   1185
      _ExtentX        =   0
      _ExtentY        =   0
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
      MICON           =   "frmSection.frx":0182
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
      Left            =   5610
      TabIndex        =   18
      Top             =   900
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
      MICON           =   "frmSection.frx":019E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin FText.asFlatTextBox txtStartStation 
      Height          =   300
      Left            =   1200
      TabIndex        =   9
      Top             =   1380
      Width           =   1515
      _ExtentX        =   2672
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
      Text            =   ""
      ButtonVisible   =   -1  'True
      OfficeXPColors  =   -1  'True
   End
   Begin FText.asFlatTextBox txtSectionID 
      Height          =   300
      Left            =   1200
      TabIndex        =   1
      Top             =   150
      Width           =   4275
      _ExtentX        =   7541
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
      Text            =   ""
   End
   Begin FText.asFlatTextBox txtEndStation 
      Height          =   300
      Left            =   3960
      TabIndex        =   11
      Top             =   1395
      Width           =   1515
      _ExtentX        =   2672
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
      Text            =   ""
      ButtonVisible   =   -1  'True
      OfficeXPColors  =   -1  'True
   End
   Begin FText.asFlatTextBox txtArea 
      Height          =   300
      Left            =   1200
      TabIndex        =   5
      Top             =   960
      Width           =   1515
      _ExtentX        =   2672
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
      Text            =   ""
      ButtonVisible   =   -1  'True
      OfficeXPColors  =   -1  'True
   End
   Begin FText.asFlatTextBox txtRoadLevel 
      Height          =   300
      Left            =   3960
      TabIndex        =   7
      Top             =   960
      Width           =   1515
      _ExtentX        =   2672
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
      Text            =   ""
      ButtonVisible   =   -1  'True
      OfficeXPColors  =   -1  'True
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      Height          =   180
      Left            =   2310
      TabIndex        =   23
      Top             =   1845
      Width           =   360
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "·����(&P):"
      Height          =   180
      Left            =   2880
      TabIndex        =   14
      Top             =   1845
      Width           =   900
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��·�ȼ�(&L):"
      Height          =   180
      Left            =   2880
      TabIndex        =   6
      Top             =   1020
      Width           =   1080
   End
   Begin VB.Label lblSectionId 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   1215
      TabIndex        =   20
      Top             =   150
      Width           =   3345
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�����(&K):"
      Height          =   180
      Left            =   120
      TabIndex        =   12
      Top             =   1845
      Width           =   900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���վ(&S):"
      Height          =   180
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�յ�վ(&E):"
      Height          =   180
      Left            =   2880
      TabIndex        =   10
      Top             =   1440
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�յ�վ����:"
      Height          =   180
      Left            =   120
      TabIndex        =   4
      Top             =   1020
      Width           =   990
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "·������(&N):"
      Height          =   180
      Left            =   120
      TabIndex        =   2
      Top             =   615
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "·�δ���(&I):"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   210
      Width           =   1080
   End
End
Attribute VB_Name = "frmSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************
'* Source File Name:frmSection.frm
'* Project Name:RTBusMan
'* Engineer:
'* Data Generated:2002/08/27
'* Last Revision Date:2002/09/02
'* Brief Description:
'* Relational Document:
'**********************************************************
Public m_bIsParent As Boolean '�Ƿ��Ǹ��������
Public m_bRouteArrange As Boolean '�Ƿ�����··�δ������
Public m_eStatus As EFormStatus
Public m_szSectionID As String '·�δ���,�޸�ʱ�õ�
Public m_szStartStation As String '����ʱ���Ը�����Ҫ�������վ
Public m_szEndStation As String '����ʱ���Ը�����Ҫ�����յ�վ
Public m_szSectionName As String '·������
Public m_szAreaID As String '��������

Private m_oSection As New Section  '·�ζ���


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
    '����
    On Error GoTo ErrorHandle
    Dim oBackSection As New BackSection
    Dim sgOldMileage As Single
    Select Case m_eStatus
    Case AddStatus '����
        m_oSection.AddNew
        m_oSection.SectionID = txtSectionID.Text
        m_oSection.SectionName = txtSectionName.Text
        m_oSection.AreaCode = ResolveDisplay(txtArea.Text)
        m_oSection.BeginStationCode = ResolveDisplay(txtStartStation.Text)
        m_oSection.EndStationCode = ResolveDisplay(txtEndStation.Text)
        m_oSection.Mileage = Val(txtKm.Text)
        m_oSection.RoadLevelCode = ResolveDisplay(txtRoadLevel.Text)
        m_oSection.SectionSerialNo = txtPathNo.Text
        m_oSection.Update
'        cmdPassCharge.Enabled = True
        If m_bIsParent Then
            '���ΪfrmAllSection ����
            frmAllSection.AddList txtSectionID.Text
'            txtSectionID.Enabled = True
            txtSectionID.Text = ""
            txtSectionName.Text = ""
            txtArea.Text = ""
            txtRoadLevel.Text = ""
            txtStartStation.Text = ""
            txtEndStation.Text = ""
            txtKm.Text = "0"
            txtSectionID.SetFocus
        ElseIf m_bRouteArrange Then
            '���ΪfrmArrangeSection ����
            frmArrangeSection.AppendSectionToLv txtSectionID.Text
            Unload Me
        End If
'        cmdOk.Caption = "����(&S)"
    Case ModifyStatus '�޸�
        m_oSection.Identify txtSectionID.Text
        sgOldMileage = m_oSection.Mileage
        m_oSection.SectionName = txtSectionName.Text
        m_oSection.AreaCode = ResolveDisplay(txtArea.Text)
        m_oSection.BeginStationCode = ResolveDisplay(txtStartStation.Text)
        m_oSection.EndStationCode = ResolveDisplay(txtEndStation.Text)
        m_oSection.Mileage = Val(txtKm.Text)
        m_oSection.RoadLevelCode = ResolveDisplay(txtRoadLevel.Text)
        m_oSection.SectionSerialNo = txtPathNo.Text
        m_oSection.Update
        oBackSection.Init g_oActiveUser
        oBackSection.UpdateRouteMileage txtSectionID.Text, txtKm.Text, sgOldMileage
        If m_bIsParent Then
            frmAllSection.UpdateList txtSectionID.Text
        ElseIf m_bRouteArrange Then
            frmArrangeSection.UpdateList txtSectionID.Text
        End If
        Unload Me
    Case ShowStatus '��ʾ
'        m_szSectionID = txtSectionID.Text
'        RefreshSection
'        cmdPassCharge.Enabled = True
'        cmdRoute.Enabled = True
        Unload Me
    End Select
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub cmdPassCharge_Click()
    '��ʾͨ�з�
'    frmPassCharge.m_szSectionID = txtSectionID.Text
'    frmPassCharge.Show vbModal
End Sub

Private Sub cmdRoute_Click()
    '��ʾ·������������·
    '�˴������޸�,��GetAllRoute�޸ĳɷ�����·���������ƵĽӿ�
    
    Dim aszRoute() As String
    Dim szaRouteName() As String
    Dim i As Integer
    Dim nCount As Integer
    On Error GoTo ErrorHandle
    Me.Height = Me.Height + Frame1.Height
    cmdRoute.Enabled = False
    lstRoute.Clear
    aszRoute = m_oSection.GetAllRoute
    szaRouteName = m_oSection.GetAllRouteName
    nCount = ArrayLength(aszRoute)
    For i = 1 To nCount
        lstRoute.AddItem aszRoute(i) & Space(12 - Len(aszRoute(i))) & szaRouteName(i)
    Next i
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub


Private Sub cmdHelp_Click()
    DisplayHelp Me
End Sub

Private Sub Form_Load()
    '��ʼ��
    m_oSection.Init g_oActiveUser
    On Error GoTo ErrorHandle
    Select Case m_eStatus
    Case AddStatus
        cmdOk.Caption = "����(&A)"
'        cmdPassCharge.Enabled = False
        cmdRoute.Enabled = False
        txtSectionID.Text = m_szSectionID
        txtSectionName.Text = m_szSectionName
        txtArea.Text = m_szAreaID
        txtStartStation.Text = m_szStartStation
        txtEndStation.Text = m_szEndStation
        frmSection.HelpContextID = 10000610
    Case ModifyStatus
        txtSectionID.Text = m_szSectionID
        RefreshSection
        
        txtSectionID.Enabled = False
        frmSection.HelpContextID = 10000650
    Case ShowStatus
        cmdOk.Caption = "��ѯ(&Q)"
'        cmdPassCharge.Enabled = False
        cmdRoute.Enabled = False
    End Select
    cmdOk.Enabled = False
    If m_eStatus = AddStatus Then txtSectionID.Enabled = True
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    m_bIsParent = False
    m_bRouteArrange = False
    m_szAreaID = ""
    m_szEndStation = ""
    m_szSectionID = ""
    m_szSectionName = ""
    m_szStartStation = ""
    
End Sub

Private Sub txtArea_ButtonClick()
    'ѡ�����
    Dim aszTemp() As String
    Dim oShell As New CommDialog
    On Error GoTo ErrorHandle
    oShell.Init g_oActiveUser
    aszTemp = oShell.SelectArea(False)
    If ArrayLength(aszTemp) = 0 Then Exit Sub
    txtArea.Text = MakeDisplayString(aszTemp(1, 1), aszTemp(1, 2))
Exit Sub
ErrorHandle:
    ShowErrorMsg

End Sub

Private Sub txtEndStation_ButtonClick()
    'ѡ���յ�վ
    Dim aszTemp() As String
    Dim oShell As New CommDialog
    On Error GoTo ErrorHandle
    If txtArea.Text = "" Then
        MsgBox "[�յ�վ����]����Ϊ��", vbInformation, Me.Caption
        Exit Sub
    End If
    oShell.Init g_oActiveUser
    aszTemp = oShell.SelectStation(ResolveDisplay(txtArea.Text), False)
    If ArrayLength(aszTemp) = 0 Then Exit Sub
    txtEndStation.Text = MakeDisplayString(aszTemp(1, 1), aszTemp(1, 2))
Exit Sub
ErrorHandle:
    ShowErrorMsg

End Sub

Private Sub txtEndStation_Validate(Cancel As Boolean)
    If txtEndStation.Text <> "" Then
        If txtEndStation.Text = txtStartStation.Text Then
            MsgBox "·�ε����վ���յ�վ������ͬ", vbInformation, Me.Caption
            Cancel = True
        End If
    End If
End Sub


Public Sub RefreshSection()
    If txtSectionID.Text = "" Then Exit Sub
    txtEndStation.Text = ""
    txtStartStation.Text = ""
    txtKm.Text = 0
    txtArea.Text = ""
    txtSectionName.Text = ""
    txtRoadLevel.Text = ""
    cmdRoute.Enabled = False
'    cmdPassCharge.Enabled = False
    m_oSection.Identify txtSectionID.Text
    txtEndStation.Text = MakeDisplayString(m_oSection.EndStationCode, m_oSection.EndStationName)
    txtStartStation.Text = MakeDisplayString(m_oSection.BeginStationCode, m_oSection.BeginStationName)
    txtKm.Text = m_oSection.Mileage
    txtArea.Text = MakeDisplayString(m_oSection.AreaCode, m_oSection.AreaName)
    txtSectionName.Text = m_oSection.SectionName
    txtRoadLevel.Text = MakeDisplayString(m_oSection.RoadLevelCode, m_oSection.RoadLevelName)
    txtPathNo = m_oSection.SectionSerialNo
    cmdOk.Caption = "����(&S)"
    m_eStatus = ModifyStatus
    If m_eStatus <> AddStatus Then
        cmdRoute.Enabled = True
'        cmdPassCharge.Enabled = True
    End If
End Sub

Private Sub txtKm_Change()
    IsSave
End Sub

Private Sub txtKm_GotFocus()
    txtKm.SelStart = 0
    txtKm.SelLength = 100
End Sub

Private Sub txtPathNo_Change()
    IsSave
End Sub

Private Sub txtRoadLevel_ButtonClick()
    'ѡ����·�ȼ�
    Dim aszTemp() As String
    Dim oShell As New CommDialog
    On Error GoTo ErrorHandle
    oShell.Init g_oActiveUser
    aszTemp = oShell.SelectRoadLevel(False)
    If ArrayLength(aszTemp) = 0 Then Exit Sub
    txtRoadLevel.Text = MakeDisplayString(aszTemp(1, 1), aszTemp(1, 2))
Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub txtSectionID_Change()
    IsSave
End Sub

Private Sub txtSectionName_Change()
    IsSave
End Sub

Private Sub IsSave()
    If txtArea.Text = "" Or txtSectionName.Text = "" Or txtEndStation.Text = "" Or txtSectionID.Text = "" Or txtStartStation.Text = "" Or txtRoadLevel.Text = "" Or txtPathNo.Text = "" Then 'Or Val(txtKm.Text) < 0
        cmdOk.Enabled = False
    Else
        cmdOk.Enabled = True
    End If
End Sub

Private Sub txtArea_Change()
    IsSave
End Sub

Private Sub txtEndStation_Change()
    IsSave
End Sub

Private Sub txtRoadLevel_Change()
    IsSave
End Sub

Private Sub txtStartStation_ButtonClick()
    'ѡ�����վ

    Dim aszTemp() As String
    Dim oShell As New CommDialog
    On Error GoTo ErrorHandle
    oShell.Init g_oActiveUser
    aszTemp = oShell.SelectStation(, False)
    If ArrayLength(aszTemp) = 0 Then Exit Sub
    txtStartStation.Text = MakeDisplayString(aszTemp(1, 1), aszTemp(1, 2))
Exit Sub
ErrorHandle:
    ShowErrorMsg

End Sub

Private Sub txtStartStation_Change()
    IsSave
End Sub

Private Sub txtStartStation_Validate(Cancel As Boolean)
    If txtStartStation.Text <> "" Then
        If txtEndStation.Text = txtStartStation.Text Then
            MsgBox "·�ε����վ���յ�վ������ͬ", vbInformation, Me.Caption
            Cancel = True
        End If
    End If
End Sub
