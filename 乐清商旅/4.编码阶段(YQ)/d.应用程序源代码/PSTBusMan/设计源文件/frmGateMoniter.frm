VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Begin VB.Form frmGateMoniter 
   BackColor       =   &H00E0E0E0&
   Caption         =   "��Ʊ��״̬"
   ClientHeight    =   6045
   ClientLeft      =   1695
   ClientTop       =   2625
   ClientWidth     =   9990
   HelpContextID   =   4001401
   Icon            =   "frmGateMoniter.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6045
   ScaleWidth      =   9990
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   6060
      Top             =   5265
   End
   Begin MSComctlLib.ImageList imgStatus 
      Left            =   5880
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGateMoniter.frx":014A
            Key             =   ""
            Object.Tag             =   "���ڼ�Ʊ"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGateMoniter.frx":0464
            Key             =   ""
            Object.Tag             =   "û�м�Ʊ"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGateMoniter.frx":05BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGateMoniter.frx":0718
            Key             =   ""
            Object.Tag             =   "����"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGateMoniter.frx":0872
            Key             =   ""
            Object.Tag             =   "δ��"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGateMoniter.frx":09CE
            Key             =   ""
            Object.Tag             =   "�Ѽ�"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGateMoniter.frx":0B28
            Key             =   ""
            Object.Tag             =   "ͣ��"
         EndProperty
      EndProperty
   End
   Begin RTComctl3.FlatLabel lblCheckBus 
      Height          =   255
      Left            =   2160
      TabIndex        =   23
      Top             =   75
      Width           =   7710
      _ExtentX        =   13600
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483633
      OutnerStyle     =   2
      HorizontalAlignment=   1
      Caption         =   "����"
   End
   Begin RTComctl3.FlatLabel lblCheckGate 
      Height          =   255
      Left            =   60
      TabIndex        =   22
      Top             =   75
      Width           =   2065
      _ExtentX        =   3651
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483633
      OutnerStyle     =   2
      HorizontalAlignment=   1
      Caption         =   "��Ʊ��"
   End
   Begin RTComctl3.CoolButton cmdCheckInfo 
      Height          =   345
      Left            =   6945
      TabIndex        =   3
      Top             =   5610
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "��Ʊ��Ϣ(&I)"
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
      MICON           =   "frmGateMoniter.frx":0C82
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdOK 
      Cancel          =   -1  'True
      Height          =   345
      Left            =   8505
      TabIndex        =   4
      Top             =   5610
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   609
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
      MICON           =   "frmGateMoniter.frx":0C9E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdRefresh 
      Height          =   345
      Left            =   60
      TabIndex        =   2
      Top             =   5625
      Width           =   2070
      _ExtentX        =   3651
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "ˢ��(&R)"
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
      MICON           =   "frmGateMoniter.frx":0CBA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame fraInfo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "��ǰ/��һ����Ʊ����"
      Height          =   1125
      Left            =   2160
      TabIndex        =   5
      Top             =   4335
      Width           =   7710
      Begin VB.Frame fraBusInfo 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   810
         Left            =   90
         TabIndex        =   6
         Top             =   195
         Width           =   7410
         Begin RTComctl3.FloatLabel lblOwner 
            Height          =   300
            Left            =   900
            TabIndex        =   26
            Top             =   255
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            HoverBackColor  =   -2147483633
            HorizontalAlignment=   1
            NormTextColor   =   -2147483635
            Caption         =   "�·�"
            NormUnderline   =   -1  'True
         End
         Begin RTComctl3.FloatLabel lblVehicle 
            Height          =   300
            Left            =   3090
            TabIndex        =   25
            Top             =   255
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            HoverBackColor  =   -2147483633
            HorizontalAlignment=   1
            NormTextColor   =   -2147483635
            Caption         =   "��D88888"
            NormUnderline   =   -1  'True
         End
         Begin RTComctl3.FloatLabel lblCompany 
            Height          =   300
            Left            =   5325
            TabIndex        =   24
            Top             =   255
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            HoverBackColor  =   -2147483633
            HorizontalAlignment=   1
            NormTextColor   =   -2147483635
            Caption         =   "��ͨ���"
            NormUnderline   =   -1  'True
         End
         Begin VB.Label lblStartupCheckTime 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "00:00"
            Height          =   180
            Left            =   3090
            TabIndex        =   21
            Top             =   570
            Width           =   450
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����ʱ��:"
            Height          =   180
            Left            =   2235
            TabIndex        =   20
            Top             =   570
            Width           =   810
         End
         Begin VB.Label lblEndStation 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            Height          =   180
            Left            =   900
            TabIndex        =   19
            Top             =   570
            Width           =   360
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��վ:"
            Height          =   180
            Left            =   75
            TabIndex        =   18
            Top             =   570
            Width           =   450
         End
         Begin VB.Label lblBusID 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0123"
            Height          =   180
            Left            =   900
            TabIndex        =   15
            Top             =   45
            Width           =   360
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����ʱ��:"
            Height          =   180
            Left            =   4455
            TabIndex        =   14
            Top             =   45
            Width           =   810
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��Ӫ��˾:"
            Height          =   180
            Left            =   4455
            TabIndex        =   13
            Top             =   315
            Width           =   810
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���г���:"
            Height          =   180
            Left            =   2235
            TabIndex        =   12
            Top             =   315
            Width           =   810
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����:"
            Height          =   180
            Left            =   75
            TabIndex        =   11
            Top             =   315
            Width           =   450
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��������:"
            Height          =   180
            Left            =   2235
            TabIndex        =   10
            Top             =   45
            Width           =   810
         End
         Begin VB.Label lblStartupTime 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "10:10"
            Height          =   180
            Left            =   5325
            TabIndex        =   9
            Top             =   45
            Width           =   450
         End
         Begin VB.Label lblBusType 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�̶�/��ˮ"
            Height          =   180
            Left            =   3090
            TabIndex        =   8
            Top             =   45
            Width           =   810
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���δ���:"
            Height          =   180
            Left            =   75
            TabIndex        =   7
            Top             =   45
            Width           =   810
         End
      End
      Begin VB.Frame fraNoBus 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   855
         Left            =   90
         TabIndex        =   16
         Top             =   195
         Width           =   6645
         Begin VB.Label lblNoCheckBus 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ǰ��Ʊ���޵ȴ���Ʊ����"
            Height          =   180
            Left            =   2760
            TabIndex        =   17
            Top             =   330
            Width           =   2160
         End
      End
   End
   Begin MSComctlLib.ListView lvCheckBus 
      Height          =   3885
      Left            =   2130
      TabIndex        =   1
      Top             =   360
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   6853
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "imgStatus"
      SmallIcons      =   "imgStatus"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "����"
         Object.Width           =   2064
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "����ʱ��"
         Object.Width           =   1640
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "��վ"
         Object.Width           =   1799
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "����"
         Object.Width           =   1640
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "���г���"
         Object.Width           =   2194
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "��Ӫ��˾"
         Object.Width           =   2036
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "����"
         Object.Width           =   2822
      EndProperty
   End
   Begin MSComctlLib.ListView lvCheckGate 
      Height          =   5115
      Left            =   15
      TabIndex        =   0
      Top             =   360
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   9022
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "imgStatus"
      SmallIcons      =   "imgStatus"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "��Ʊ��"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "����"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmGateMoniter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const StartCheckTicketTime = "BeginCheckTime"
Const cnMinWidth = 10110
Const cnMinHeight = 6545

Private oBase As New BaseInfo
Private oChkTk As New CheckTicket
Private m_nBusStatus As Integer

'���±������ڸ����û��ĳ����б�ѡ���Ա�����ʾѡ�г��εļ�Ʊ��Ϣ
Private m_dyBusDate As Date
Private m_szBusID As String
Private m_nBusType As Integer

'���±������ڸ��ٵ�ǰ/��һ��Ʊ���ε��й���Ϣ��������ʾ����ϸ���ݣ��糵���ȣ�
Private m_szOwnerID As String
Private m_szTransportCompanyID As String
Private m_szVehicleId As String

Private m_nOriWidth As Integer
Private m_nOriHeight As Integer
Private m_bAllowResize As Boolean

Private m_bIsShowing As Boolean
Private mabInfoGot() As Boolean        '���ڱ�ʶĳ����Ʊ���Ƿ�ȡ�����ݣ�Ϊ�˱����ظ���ȡ���ݿ�
Public m_cBusInfo As New ucCheckBusLst     '������Ϣ����
Public m_cChkingBusInfo As New ucCheckBusLst   'ÿ����Ʊ������/��һ��Ʊ���μ���


Private Sub cmdCheckInfo_Click()
    Dim oChkApp As New CommDialog
    oChkApp.Init g_oActiveUser
    If m_nBusType = EBusType.TP_RegularBus Then
        oChkApp.ShowCheckInfo m_dyBusDate, m_szBusID
    End If
    Set oChkApp = Nothing
End Sub

Private Sub cmdOk_Click()
    Unload Me
End Sub

Private Sub cmdRefresh_Click()
MousePointer = MousePointerConstants.vbHourglass
On Error GoTo ErrorPos
    Dim szCheckGateID As String
    Dim nLoop As Integer
    
    szCheckGateID = lvCheckGate.SelectedItem.Text   '����ԭѡ��ļ�Ʊ�ں�
    
    Dim nCurrPos As Integer '����ԭѡ�����
    If lvCheckGate.ListItems.Count > 0 Then
        nCurrPos = lvCheckGate.SelectedItem.Index
    Else
        nCurrPos = 0
    End If
    RefreshGateInfo     'ˢ�¼�Ʊ����Ϣ
    
    
    If lvCheckGate.ListItems.Count = 0 Then
        '���ԭ��Ʊ�����г����б��˳�
        m_cBusInfo.RemoveAll
        MsgBox "ϵͳû�м�Ʊ��", vbExclamation, "����"
        Unload Me
        Exit Sub
    End If
    
    
    If nCurrPos = 0 Or nCurrPos > lvCheckGate.ListItems.Count Then
        nCurrPos = 1
    End If
    lvCheckGate.ListItems(nCurrPos).Selected = True
    mabInfoGot(lvCheckGate.SelectedItem.Index) = False
    lvCheckGate_ItemClick lvCheckGate.SelectedItem
    
    MousePointer = MousePointerConstants.vbDefault
    Exit Sub
ErrorPos:
    MousePointer = MousePointerConstants.vbDefault
    ShowErrorMsg
End Sub


Private Sub Form_Load()
    'm_bAllowResize�������������ڵ�һ��
    '�ı䴰�ڵĸ߿�ʱform_resize��ִ�м���ؼ�λ�á���С�Ĵ���
On Error GoTo ErrorPos
    m_bAllowResize = False
    m_nOriWidth = cnMinWidth
    m_nOriHeight = cnMinHeight
    Me.Width = cnMinWidth
    Me.Height = cnMinHeight
    m_bAllowResize = True
    
    lvCheckGate.SmallIcons = imgStatus
    lvCheckBus.SmallIcons = imgStatus
'    g_oActiveUser.Login "03", "1", "elf"
    oBase.Init g_oActiveUser
    oChkTk.Init g_oActiveUser
    
    lblBusID.Caption = ""
    lblStartupTime.Caption = ""
    lblEndStation.Caption = ""
    lblBusType.Caption = ""
    lblVehicle.Caption = ""
    lblCompany.Caption = ""
    lblOwner.Caption = ""
    m_szOwnerID = ""
    m_szTransportCompanyID = ""
    m_szVehicleId = ""
    lblStartupCheckTime.Caption = ""
    cmdCheckInfo.Enabled = False
        
    Exit Sub
ErrorPos:
    ShowErrorMsg
End Sub

Private Sub Form_Resize()
    
    Dim nIncWidth, nIncHeight As Integer
    Dim bClip As Boolean '��ʾ�Ƿ����

    bClip = False
    If Not m_bAllowResize Then
        Exit Sub
    End If
    If Me.WindowState = vbMinimized Then
        '��С������ʱ���ƶ�λ��
        Exit Sub
    End If
    
'    If Not m_bHasShow Then
'        m_nOriWidth = cnMinWidth
'        m_nOriHeight = cnMinHeight
'        Me.Width = cnMinWidth
'        Me.Height = cnMinHeight
'        m_bHasShow = True
'    End If

    If Me.Width < cnMinWidth Then
        If Not Me.WindowState = vbMaximized Then
            '���ʱ�������ô�С
            Me.Width = cnMinWidth
        End If
    End If
    If Me.Height < cnMinHeight Then
        If Not Me.WindowState = vbMaximized Then
            Me.Height = cnMinHeight
        End If
    End If
    
    
    If Me.Width < cnMinWidth Then
        nIncWidth = 0
    Else
        nIncWidth = Me.Width - m_nOriWidth
        m_nOriWidth = Me.Width
    End If
    If Me.Height < cnMinHeight Then
        nIncHeight = 0
    Else
        nIncHeight = Me.Height - m_nOriHeight
        m_nOriHeight = Me.Height
    End If
    
    lvCheckBus.Width = lvCheckBus.Width + nIncWidth
    lvCheckBus.Height = lvCheckBus.Height + nIncHeight
    lvCheckGate.Height = lvCheckGate.Height + nIncHeight
    fraInfo.Width = fraInfo.Width + nIncWidth
    fraInfo.Top = fraInfo.Top + nIncHeight
    cmdRefresh.Top = cmdRefresh.Top + nIncHeight
    cmdOK.Top = cmdOK.Top + nIncHeight
    cmdOK.Left = cmdOK.Left + nIncWidth
    cmdCheckInfo.Top = cmdCheckInfo.Top + nIncHeight
    cmdCheckInfo.Left = cmdCheckInfo.Left + nIncWidth
    
    lblCheckBus.Width = lblCheckBus.Width + nIncWidth
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m_cBusInfo = Nothing
    Set m_cChkingBusInfo = Nothing
    m_bIsShowing = False
End Sub

Private Sub lblCompany_Click()
    
    If Len(lblCompany.Caption) > 0 Then
        Dim oChkApp As New CommDialog
        oChkApp.Init g_oActiveUser
        oChkApp.ShowCompanyInfo m_szTransportCompanyID
        lblCompany.NormTextColor = &H8000000D
        Set oChkApp = Nothing
    End If

End Sub

Private Sub lblOwner_Click()
    If Len(lblOwner.Caption) > 0 Then
       Dim oChkApp As New CommDialog
       oChkApp.Init g_oActiveUser
       oChkApp.ShowOwnerInfo m_szOwnerID
       lblOwner.NormTextColor = &H8000000D
       Set oChkApp = Nothing
    End If
    
End Sub

Private Sub lblVehicle_Click()
    If Len(lblVehicle.Caption) > 0 Then
        Dim oChkApp As New CommDialog
        oChkApp.Init g_oActiveUser
        oChkApp.ShowVehicleInfo m_szVehicleId
        lblVehicle.NormTextColor = &H8000000D
        Set oChkApp = Nothing
    End If
    
End Sub

Private Sub lvCheckBus_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    SortListView lvCheckBus, ColumnHeader.Index - 1
End Sub

Private Sub lvCheckGate_ItemClick(ByVal Item As MSComctlLib.ListItem) 'ѡ��Ʊ��
    MousePointer = MousePointerConstants.vbHourglass
On Error GoTo ErrorPos
    Dim szCheckGateID As String
    Dim nTmp As Integer
    
    lblCheckGate.Caption = Item.SubItems(1) '���ļ�Ʊ�ڱ�ǩ
    szCheckGateID = Item.Text
'    oChkTk.CheckGateNo = szCheckGateID

    
    '���´���õ���ǰ��Ʊ�ڵĳ�����Ϣ�б�
'    nTmp = m_cBusInfo.FindItemByGate(szCheckGateID)
    ShowSBInfo "ȡ�ü�Ʊ�ڳ����б�..."
    If Not mabInfoGot(Item.Index) Then        '�����б���Ϣ�����޸ü�Ʊ�ڵĳ������ݲ��ҵ�һ��ѡ��ü�Ʊ��
        GetBusLstInfoByGate szCheckGateID
    End If
    RefreshlvCheckBus szCheckGateID  'ˢ�³����б�
        
    
    '���´���õ����ڼ�Ʊ����һ�೵����Ϣ
'    nTmp = m_cChkingBusInfo.FindItemByGate(szCheckGateID)
    ShowSBInfo "ȡ�ü�Ʊ�����ڼ�Ʊ����һ�೵����Ϣ..."
    If Not mabInfoGot(Item.Index) Then   '���쳵����Ϣ�����޸ü�Ʊ�ڵĳ������ݲ��ҵ�һ��ѡ��ü�Ʊ��
        GetChkingBusInfo szCheckGateID
    End If
    RefreshChkingBus szCheckGateID
        
    mabInfoGot(Item.Index) = True '��ʶΪ�ü�Ʊ���Ѷ�ȡ������
    MousePointer = MousePointerConstants.vbDefault
    ShowSBInfo ""
    Exit Sub
ErrorPos:
    ShowSBInfo ""
    MousePointer = MousePointerConstants.vbDefault
    ShowErrorMsg
End Sub



Private Sub lvCheckBus_ItemClick(ByVal Item As MSComctlLib.ListItem) 'ѡ����
    Dim szBusID As String, nSerialNo As Integer
    Dim nTmp As Integer
    
    szBusID = Item.Text
    lblCheckBus.Caption = szBusID
            
'    '������ʶ����ȡ�ó��εĳ��κźͳ������(��ʶ��Ϊ ���>���κ�)
'    nTmp = InStr(1, szBusId, ">")
'    szBusId = Right(Item.Tag, Len(Item.Tag) - nTmp)
'    nSerialNo = Val(Left(Item.Tag, nTmp - 1))
        
    cmdCheckInfo.Enabled = False
    nTmp = m_cBusInfo.FindItem(szBusID)
    If nTmp > 0 Then    '���ü�Ʊ��Ϣ��ť����Ч��
        If m_cBusInfo.Item(nTmp).BusMode = EBusType.TP_RegularBus Then
            If m_cBusInfo.Item(nTmp).Status = EREBusStatus.ST_BusChecking Or m_cBusInfo.Item(nTmp).Status = EREBusStatus.ST_BusExtraChecking Or _
                m_cBusInfo.Item(nTmp).Status = EREBusStatus.ST_BusStopCheck Then
                cmdCheckInfo.Enabled = True
            End If
            m_nBusType = EBusType.TP_RegularBus
        Else
'            cmdCheckInfo.Enabled = True
            m_nBusType = EBusType.TP_ScrollBus
        End If
    End If
    m_szBusID = szBusID
    m_dyBusDate = Date
End Sub

Public Sub GetBusLstInfoByGate(szGateId As String)
'�õ�ָ����Ʊ�ڵ����г���,���복���б���Ϣ����
    Dim tTmpBusInfo As tCheckBusLstInfo
    Dim nSerialNo As Integer, m_nBusStatus As Integer
    Dim i As Integer, nCount As Integer
    m_cBusInfo.RemoveByGate szGateId
    Dim rsBus As New Recordset
    Set rsBus = oChkTk.GetBusInfoRs(Date, szGateId) 'ȡ���г�����Ϣ
    nCount = rsBus.RecordCount
    For i = 1 To nCount
        tTmpBusInfo.BusID = Trim(rsBus("bus_id"))
        tTmpBusInfo.BusMode = rsBus("bus_type")
        tTmpBusInfo.Company = rsBus("transport_company_short_name")
        tTmpBusInfo.Vehicle = rsBus("license_tag_no")
        tTmpBusInfo.StartUpTime = rsBus("bus_start_time")
  '      tTmpBusInfo.EndStationName = rsBus("end_station_name")
        tTmpBusInfo.Owner = rsBus("owner_name")
        If tTmpBusInfo.BusMode = EBusType.TP_ScrollBus Then
            If IsNull(rsBus("bus_serial_no")) Then
                m_nBusStatus = EREBusStatus.ST_BusNormal
            Else
                nSerialNo = rsBus("bus_serial_no")
                m_nBusStatus = oChkTk.GetBusStatus(Date, tTmpBusInfo.BusID, nSerialNo, szGateId)
            End If
            tTmpBusInfo.Vehicle = ""
            tTmpBusInfo.Owner = ""
        Else
            m_nBusStatus = rsBus("status")
        End If
        tTmpBusInfo.Status = m_nBusStatus
        tTmpBusInfo.CheckGate = szGateId
        m_cBusInfo.Addone tTmpBusInfo
        rsBus.MoveNext
    Next
End Sub
Public Sub RefreshlvCheckBus(szCheckGateID As String)
'��ָ����Ʊ�ںţ��ӳ��μ�Ʊ��Ϣ����(m_cBusInfo)��ȡ�����ݣ�ˢ�³����б��
    Dim tTmpBusInfo As tCheckBusLstInfo
    Dim i As Integer, j As Integer
    Dim nBusImage As Integer
    Dim lstTemp As ListItem
    Dim nCurrPos As Integer '����ԭ��ѡ���
    If lvCheckBus.ListItems.Count > 0 Then
        nCurrPos = lvCheckBus.SelectedItem.Index
    Else
        nCurrPos = 0
    End If
    
    lvCheckBus.ListItems.Clear  '����б��
    j = 1
    For i = 1 To m_cBusInfo.Count
        If m_cBusInfo.Item(i).CheckGate = szCheckGateID Then
            Dim szTmpBus As String
            nBusImage = GetImageIndexByStatus(m_cBusInfo.Item(i).Status) 'ȡ����Ӧ״̬ͼ���������
            szTmpBus = m_cBusInfo.Item(i).BusID
            Set lstTemp = lvCheckBus.ListItems.Add(j, , szTmpBus, , nBusImage)
            lstTemp.ListSubItems.Add , , Format(m_cBusInfo.Item(i).StartUpTime, "HH:mm")
            lstTemp.ListSubItems.Add , , m_cBusInfo.Item(i).EndStationName
            lstTemp.ListSubItems.Add , , oChkTk.GetBusTypeName(m_cBusInfo.Item(i).BusMode) 'IIf(m_cBusInfo.Item(i).BusMode = TP_RegularBus, "�̶�����", "��ˮ����")
            lstTemp.ListSubItems.Add , , m_cBusInfo.Item(i).Vehicle
            lstTemp.ListSubItems.Add , , m_cBusInfo.Item(i).Company
            lstTemp.ListSubItems.Add , , m_cBusInfo.Item(i).Owner
'            '��ʶ��Ϊ ���>���κ� �����ڱ�ʶ������
'            lvCheckBus.ListItems.Item(j).Tag = Trim(Val(m_cBusInfo.Item(i).BusSerial)) & ">" & m_cBusInfo.Item(i).BusID
            j = j + 1
        End If
    Next i
    
    If lvCheckBus.ListItems.Count > 0 Then  '�ָ�ԭ��ѡ���
        If nCurrPos = 0 Or nCurrPos > lvCheckBus.ListItems.Count Then
            lvCheckBus.ListItems(1).Selected = True
        Else
            lvCheckBus.ListItems(nCurrPos).Selected = True
        End If
        lvCheckBus_ItemClick lvCheckBus.SelectedItem
    End If
End Sub
Public Sub RefreshChkingBus(szCheckGateID As String)
'��ָ����Ʊ�ںţ�������/��һ��Ʊ������Ϣ����(m_cChkingBusInfo)��ȡ�����ݣ�ˢ������/��һ��Ʊ������Ϣ
    Dim nTmp As Integer
    Dim tTmpBusInfo As tCheckBusLstInfo
    nTmp = m_cChkingBusInfo.FindItemByGate(szCheckGateID)
    If nTmp > 0 Then
        fraBusInfo.Visible = True
        fraNoBus.Visible = False
    
        lblBusID.Caption = m_cChkingBusInfo.Item(nTmp).BusID
        lblStartupTime.Caption = Format(m_cChkingBusInfo.Item(nTmp).StartUpTime, "HH:mm")
        lblEndStation.Caption = m_cChkingBusInfo.Item(nTmp).EndStationName
        lblBusType.Caption = oChkTk.GetBusTypeName(m_cChkingBusInfo.Item(nTmp).BusMode)   '  IIf(m_cChkingBusInfo.Item(nTmp).BusMode = TP_RegularBus, "�̶�����", "��ˮ����")
        lblStartupCheckTime.Caption = Format(m_cChkingBusInfo.Item(nTmp).StartChkTime, cszTimeStr)
        
        '�ֱ�õ���������˾�ͳ�����ID��Name
        lblVehicle.Caption = GetContentInStr(m_cChkingBusInfo.Item(nTmp).Vehicle)
        lblCompany.Caption = GetContentInStr(m_cChkingBusInfo.Item(nTmp).Company)
        lblOwner.Caption = GetContentInStr(m_cChkingBusInfo.Item(nTmp).Owner)
        
        m_szOwnerID = GetIDinStr(m_cChkingBusInfo.Item(nTmp).Owner)
        m_szTransportCompanyID = GetIDinStr(m_cChkingBusInfo.Item(nTmp).Company)
        m_szVehicleId = GetIDinStr(m_cChkingBusInfo.Item(nTmp).Vehicle)
    Else
        fraBusInfo.Visible = False
        fraNoBus.Visible = True
    End If
End Sub
Public Sub GetChkingBusInfo(szCheckGateID As String)
'�õ�ָ����Ʊ�ڵĵ�ǰ/��һ��Ʊ����
    Dim nTmp As Integer
    Dim auCheckingBus() As TBusCheckInfo
    Dim nArrayLength As Integer
    Dim szBusID As String, nSerialNo As Integer
    Dim tTmpBusInfo As tCheckBusLstInfo
    
    m_cChkingBusInfo.RemoveByGate szCheckGateID
    szBusID = ""
    nSerialNo = 0
    oChkTk.CheckGateNo = szCheckGateID
    auCheckingBus = oChkTk.GetCheckingBus 'ȡ�������ڼ�Ʊ�ĳ���
    nArrayLength = ArrayLength(auCheckingBus) '�õ����ڼ�Ʊ�ĳ�������
    If nArrayLength > 0 Then    '�����쳵��
        szBusID = Trim(auCheckingBus(1).szBusID)
        nSerialNo = auCheckingBus(1).nBusSerialNo
        Dim oTmpVehicle As New Vehicle
        oTmpVehicle.Init g_oActiveUser
        oTmpVehicle.Identify Trim(auCheckingBus(1).szVehicleID)
                        
        '�������쳵�λ���һ��Ʊ������Ϣ����
        nTmp = m_cBusInfo.FindItem(szBusID)
        If nTmp > 0 Then
            tTmpBusInfo.BusID = szBusID
            tTmpBusInfo.BusSerial = nSerialNo
            tTmpBusInfo.BusMode = m_cBusInfo.Item(nTmp).BusMode
            tTmpBusInfo.CheckGate = szCheckGateID
            tTmpBusInfo.EndStationName = m_cBusInfo.Item(nTmp).EndStationName
            tTmpBusInfo.StartChkTime = auCheckingBus(1).dtBeginCheckTime
            tTmpBusInfo.StartUpTime = auCheckingBus(1).dtStartupTime
            tTmpBusInfo.Status = m_cBusInfo.Item(nTmp).Status
            
            tTmpBusInfo.Company = SetIDinStr(Trim(oTmpVehicle.Company)) & oTmpVehicle.CompanyName
            tTmpBusInfo.Owner = SetIDinStr(Trim(oTmpVehicle.Owner)) & oTmpVehicle.OwnerName
            tTmpBusInfo.Vehicle = SetIDinStr(Trim(oTmpVehicle.VehicleId)) & oTmpVehicle.LicenseTag
            
            nTmp = m_cChkingBusInfo.FindItem(szBusID)   '����ˢ�£�û�����
            If nTmp > 0 Then
                m_cChkingBusInfo.UpdateOne tTmpBusInfo
            Else
                m_cChkingBusInfo.Addone tTmpBusInfo
            End If
        End If
    Else
        Dim oNextCheckBus As New REBus
        Set oNextCheckBus = oChkTk.GetNextCheckBus()
        If Not oNextCheckBus Is Nothing Then
            tTmpBusInfo.BusID = Trim(oNextCheckBus.BusID)
            tTmpBusInfo.BusSerial = 0
            tTmpBusInfo.BusMode = oNextCheckBus.BusType
            tTmpBusInfo.CheckGate = szCheckGateID
            tTmpBusInfo.EndStationName = oNextCheckBus.EndStationName
            Dim oSysParam As New SystemParam
            oSysParam.Init g_oActiveUser
            tTmpBusInfo.StartUpTime = oNextCheckBus.StartUpTime
            tTmpBusInfo.StartChkTime = tTmpBusInfo.StartUpTime - oSysParam.BeginCheckTime / 60 / 24
            
            tTmpBusInfo.Company = SetIDinStr(oNextCheckBus.Company) & oNextCheckBus.CompanyName
            tTmpBusInfo.Owner = SetIDinStr(oNextCheckBus.Owner) & oNextCheckBus.OwnerName
            tTmpBusInfo.Vehicle = SetIDinStr(oNextCheckBus.Vehicle) & oNextCheckBus.VehicleTag
            m_cChkingBusInfo.Addone tTmpBusInfo
        End If
    End If

End Sub
Private Function GetImageIndexByStatus(nStatus As Integer) As Integer
'���ݳ���״̬�ŷ�����Ӧ��ͼ��������
    Select Case nStatus
        Case EREBusStatus.ST_BusNormal
            GetImageIndexByStatus = 5
        Case EREBusStatus.ST_BusStopCheck
            GetImageIndexByStatus = 6
        Case EREBusStatus.ST_BusChecking, EREBusStatus.ST_BusExtraChecking
            GetImageIndexByStatus = 4
        Case Else
            GetImageIndexByStatus = 7
    End Select
End Function

Public Property Get IsShow() As Boolean
    IsShow = m_bIsShowing
End Property

Private Function GetIDinStr(szParam As String) As String
'���������а�����ID��(<>�а������ַ�)
    Dim nTmp As Integer
    nTmp = InStr(1, szParam, ">")
    GetIDinStr = Mid(szParam, 2, nTmp - 2)
End Function
Private Function SetIDinStr(szID As String) As String
'������Ϊ�����ID��(��ID���������<>)
    SetIDinStr = "<" & szID & ">"
End Function
Private Function GetContentInStr(szParam As String)
'���������е����ݣ���ID������ַ�����
    Dim nTmp As Integer
    nTmp = InStr(1, szParam, ">")
    GetContentInStr = Right(szParam, Len(szParam) - nTmp)
End Function

Public Sub RefreshGateInfo()
'�õ����м�Ʊ����Ϣ
    Dim aszCheckGate() As String
    Dim nArrayLength As Integer
    Dim szCheckGateID As String
    
    Dim nLoop As Integer
    Dim nCheckGataStatus As Integer
    Dim nTmp As Integer
        
    aszCheckGate = oBase.GetAllCheckGate 'ȡ���м�Ʊ����Ϣ
    nArrayLength = ArrayLength(aszCheckGate)
    If lvCheckGate.ListItems.Count > 0 Then
        For nLoop = 1 To nArrayLength
            '���Ѹ��ĵľɼ�Ʊ�ڵ����г����б���Ϣɾ�����������ݵ�һ��
            If lvCheckGate.ListItems(nLoop).Text <> Trim(aszCheckGate(nLoop, 1)) Then
                mabInfoGot(nLoop) = False       '��ʶΪû�ж�ȡ����
                m_cBusInfo.RemoveByGate lvCheckGate.ListItems(nLoop).Text
                m_cChkingBusInfo.RemoveByGate lvCheckGate.ListItems(nLoop).Text
            End If
        Next nLoop
        lvCheckGate.ListItems.Clear
    End If
    For nLoop = 1 To nArrayLength
        szCheckGateID = Trim(aszCheckGate(nLoop, 1))
        oChkTk.CheckGateNo = szCheckGateID
        If oChkTk.GetCheckGateStatus = ST_CheckGateChecking Then '�жϼ�Ʊ��״̬
            nCheckGataStatus = 1
        Else
            nCheckGataStatus = 2
        End If
        lvCheckGate.ListItems.Add , , szCheckGateID, , nCheckGataStatus
        lvCheckGate.ListItems.Item(nLoop).SubItems(1) = Trim(aszCheckGate(nLoop, 2))
    Next nLoop

End Sub
Private Sub Timer1_Timer()
    Timer1.Enabled = False
        
    '��Ӽ�Ʊ����Ϣ
    ShowSBInfo "���ڶ�ȡ��Ʊ����Ϣ..."
    Me.MousePointer = vbHourglass
    Screen.MousePointer = vbHourglass
    RefreshGateInfo
    If lvCheckGate.ListItems.Count = 0 Then
        MsgBox "ϵͳû�м�Ʊ��", vbExclamation, "����"
        ShowSBInfo ""
        
        Me.MousePointer = vbDefault
        Unload Me
        
        Exit Sub
    End If
    ReDim mabInfoGot(1 To lvCheckGate.ListItems.Count)
    
    lblCheckGate.Caption = lvCheckGate.ListItems(1).SubItems(1)
    
    'ȡ��һ��Ʊ�ڵĳ�����Ϣ
    ShowSBInfo "ȡ��Ʊ�ڵĳ�����Ϣ..."
    Dim szCheckGateID As String
    szCheckGateID = lvCheckGate.ListItems(1).Text
    GetBusLstInfoByGate szCheckGateID
    GetChkingBusInfo szCheckGateID
    RefreshlvCheckBus szCheckGateID
    RefreshChkingBus szCheckGateID
    lvCheckGate.ListItems(1).Selected = True
    mabInfoGot(1) = True        '��Ʊ��1�Ѷ�������
    
    m_bIsShowing = True
    cmdRefresh_Click
    
    ShowSBInfo ""
    SetNormal
End Sub
