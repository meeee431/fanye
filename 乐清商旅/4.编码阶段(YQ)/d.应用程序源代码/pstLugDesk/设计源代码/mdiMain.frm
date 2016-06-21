VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "actbar2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.MDIForm mdiMain 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000C&
   Caption         =   "�а�����"
   ClientHeight    =   8175
   ClientLeft      =   1275
   ClientTop       =   2055
   ClientWidth     =   11100
   HelpContextID   =   7000001
   Icon            =   "mdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   WindowState     =   2  'Maximized
   Begin ActiveBar2LibraryCtl.ActiveBar2 abMenu 
      Align           =   1  'Align Top
      Height          =   8175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11100
      _LayoutVersion  =   1
      _ExtentX        =   19579
      _ExtentY        =   14420
      _DataPath       =   ""
      Bands           =   "mdiMain.frx":16AC2
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   480
         Top             =   2160
      End
      Begin MSComctlLib.ProgressBar pbLoad 
         Height          =   225
         Left            =   4920
         TabIndex        =   6
         Top             =   7020
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.PictureBox ptTitleTop 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   30
         Picture         =   "mdiMain.frx":203EE
         ScaleHeight     =   450
         ScaleWidth      =   15360
         TabIndex        =   1
         Top             =   1140
         Width           =   15360
         Begin RTComctl3.CoolButton cmdClose 
            Height          =   390
            Left            =   11670
            TabIndex        =   2
            ToolTipText     =   "����"
            Top             =   0
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   688
            BTYPE           =   12
            TX              =   ""
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
            COLTYPE         =   3
            FOCUSR          =   0   'False
            BCOL            =   12632256
            BCOLO           =   12632256
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "mdiMain.frx":25A27
            PICN            =   "mdiMain.frx":25A43
            PICH            =   "mdiMain.frx":26938
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   1
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label lblSheetNo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "����"
               Size            =   15
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   10170
            TabIndex        =   5
            Top             =   90
            Width           =   165
         End
         Begin VB.Label lblSheetNoName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ǰ�а�����:"
            Height          =   180
            Left            =   8940
            TabIndex        =   4
            Top             =   150
            Width           =   1170
         End
         Begin VB.Label fblCurrentTime 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0:00:00"
            BeginProperty Font 
               Name            =   "����"
               Size            =   15
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   7140
            TabIndex        =   3
            Top             =   60
            Width           =   1185
         End
      End
      Begin VB.Label lblInStation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ";��վ"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   540
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub abMenu_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
On Error GoTo ErrHandle
    Select Case Tool.name
        Case "mi_Accept"        '��������
            
            AcceptLuggage
        Case "mi_ReturnAccept" '�а�����
            ReturnLuggageAccept
        Case "mi_CancelAccept"  '�а�������
            CancelLuggageAccept
        Case "mi_Carry"         'ǩ���а���
            CarryLuggage
        Case "mi_ReprintSheet"  '�ش�ǩ����
            frmRePrintSheet.Show vbModal
        Case "mi_CancelSheet"   '����ǩ����
            frmCancelSheet.Show vbModal
        Case "mi_QueryAccept"   '��ѯ�а���
            frmQueryAccept.ZOrder 0
            frmQueryAccept.Show
        Case "mi_QuerySheet"    '��ѯǩ����
            frmQuerySheet.ZOrder 0
            frmQuerySheet.Show
        Case "mi_StatLuggage"   '�е�ͳ�ƽ���
            frmSumAccept.ZOrder 0
            frmSumAccept.Show
        Case "mi_SheetNo"       '�����е�������ǩ������
            RefreshNO
'            frmChgSheetNo.Show vbModal
        Case "mi_ChgPassword"
            ChangePassword
        Case "mi_SysExit"
            Unload Me
        Case "mnu_HelpIndex"
            DisplayHelp Me, Index
        Case "mnu_HelpContent"
            If Not ActiveForm Is Nothing Then
                DisplayHelp ActiveForm, content
            End If
        Case "mnu_About"
            AboutMe
            
        '������ϵͳ����
        
        Case "tbn_system_print"
            ActiveForm.PrintReport False
        Case "mnu_system_print"
            ActiveForm.PrintReport True
        Case "tbn_system_printview", "mnu_system_printview"
            ActiveForm.PreView
        Case "mnu_PageOption"
            'ҳ������
            ActiveForm.PageSet
        Case "mnu_PrintOption"
            '��ӡ����
            ActiveForm.PrintSet
        Case "tbn_system_export", "mnu_ExportFile"
            ActiveForm.ExportFile
        Case "tbn_system_exportopen", "mnu_ExportFileOpen"
            ActiveForm.ExportFileOpen
    End Select
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub AboutMe()
Dim oShell As New CommShell
    oShell.ShowAbout App.ProductName, "Luggage Desk", App.FileDescription, Me.Icon, App.Major, App.Minor, App.Revision
End Sub

Private Sub ChangePassword()

 Dim oShell As New CommDialog
    On Error GoTo ErrorHandle
    oShell.Init m_oAUser
    oShell.ShowUserInfo
    Set oShell = Nothing
    Exit Sub
ErrorHandle:
    Set oShell = Nothing
    ShowErrorMsg
End Sub
'ˢ���޸ĺ�ĵ�������ǩ������
Private Sub RefreshNO()
    frmChgSheetNo.m_bNoCancel = False
    frmChgSheetNo.Show vbModal, Me
    If frmChgSheetNo.m_bOk Then
      If lblSheetNoName.Caption = "��ǰ������:" Then
        lblSheetNo.Caption = GetTicketNo()
      ElseIf lblSheetNoName.Caption = "��ǰǩ������:" Then
        lblSheetNo.Caption = g_szCarrySheetID
      End If
    End If
End Sub
'��������
Private Sub AcceptLuggage()
    frmAccept.ZOrder 0
    frmAccept.Show
    
End Sub
'�а�����
Private Sub ReturnLuggageAccept()
    frmReturnAccept.ZOrder 0
    frmReturnAccept.Show
    
End Sub

'�а�������
Private Sub CancelLuggageAccept()
    frmCancelAccept.ZOrder 0
    frmCancelAccept.Show
    
End Sub

'ǩ���а���
Private Sub CarryLuggage()
    frmCarryLuggage.ZOrder 0
    frmCarryLuggage.Show
    
End Sub

'����ActiveBar�Ŀؼ�
Private Sub AddControlsToActBar()
    abMenu.Bands("bndTitleTop").Tools("tblTitleTop").Custom = ptTitleTop
    abMenu.Bands("statusBar").Tools("progressBar").Custom = pbLoad
End Sub

Private Sub cmdClose_Click()
On Error Resume Next
    If Not ActiveForm Is Nothing Then
        Unload ActiveForm
    End If
End Sub


Private Sub MDIForm_Load()
    AddControlsToActBar
    
    SetPrintEnabled False
    
    '��ʼ�������棬��״̬����
    frmAccept.ZOrder 0
    frmAccept.Show     'ȱʡ��������
End Sub

Private Sub Timer1_Timer()
    fblCurrentTime.Caption = Time
End Sub

Public Sub SetPrintEnabled(pbEnabled As Boolean)
    '���ò˵��Ŀ�����
    With abMenu
        .Bands("tbn_system").Tools("tbn_system_print").Enabled = pbEnabled
        .Bands("tbn_system").Tools("tbn_system_printview").Enabled = pbEnabled
        .Bands("tbn_system").Tools("tbn_system_export").Enabled = pbEnabled
        .Bands("tbn_system").Tools("tbn_system_exportopen").Enabled = pbEnabled
        .Bands("mnu_System").Tools("mnu_PageOption").Enabled = pbEnabled
        .Bands("mnu_System").Tools("mnu_PrintOption").Enabled = pbEnabled
        .Bands("mnu_System").Tools("mnu_system_print").Enabled = pbEnabled
        .Bands("mnu_System").Tools("mnu_system_printview").Enabled = pbEnabled
        .Bands("mnu_System").Tools("mnu_ExportFile").Enabled = pbEnabled
        .Bands("mnu_System").Tools("mnu_ExportFileOpen").Enabled = pbEnabled
    End With
End Sub
