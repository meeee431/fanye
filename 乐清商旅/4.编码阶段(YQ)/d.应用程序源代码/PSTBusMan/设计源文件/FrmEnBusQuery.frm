VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Begin VB.Form FrmEnBusQuery 
   BackColor       =   &H00E0E0E0&
   Caption         =   "������λϢ�Ų�ѯ"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9360
   Icon            =   "FrmEnBusQuery.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   9360
   StartUpPosition =   2  '��Ļ����
   Begin RTComctl3.CoolButton CmdExit 
      Height          =   375
      Left            =   315
      TabIndex        =   15
      Top             =   4140
      Width           =   1095
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
      MICON           =   "FrmEnBusQuery.frx":030A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdQuery 
      Default         =   -1  'True
      Height          =   420
      Left            =   315
      TabIndex        =   13
      Top             =   3645
      Width           =   1095
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "��ѯ(&Q)"
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
      MICON           =   "FrmEnBusQuery.frx":0326
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
      Caption         =   "��ѯ����"
      Height          =   3165
      Left            =   90
      TabIndex        =   1
      Top             =   90
      Width           =   1770
      Begin RTComctl3.CoolButton cmdFind 
         Height          =   300
         Left            =   10515
         TabIndex        =   4
         Top             =   15
         Width           =   1020
         _ExtentX        =   0
         _ExtentY        =   0
         BTYPE           =   2
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmEnBusQuery.frx":0342
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox txtBusID 
         Height          =   300
         Left            =   180
         MaxLength       =   5
         TabIndex        =   3
         Top             =   1845
         Width           =   1440
      End
      Begin VB.ComboBox cboStation 
         Height          =   300
         ItemData        =   "FrmEnBusQuery.frx":035E
         Left            =   180
         List            =   "FrmEnBusQuery.frx":0365
         TabIndex        =   2
         Text            =   "(ȫ��)"
         Top             =   3465
         Width           =   1440
      End
      Begin RTComctl3.TextButtonBox txtRoute 
         Height          =   300
         Left            =   180
         TabIndex        =   5
         Top             =   2610
         Width           =   1440
         _ExtentX        =   2540
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
         Text            =   "(ȫ��)"
      End
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   300
         Left            =   180
         TabIndex        =   6
         Top             =   540
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   529
         _Version        =   393216
         Format          =   65142784
         CurrentDate     =   36396
      End
      Begin MSComCtl2.DTPicker dtpEndDate 
         Height          =   300
         Left            =   180
         TabIndex        =   7
         Top             =   1170
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   529
         _Version        =   393216
         Format          =   65142784
         CurrentDate     =   36396
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��·(&R):"
         Height          =   180
         Left            =   180
         TabIndex        =   12
         Top             =   2295
         Width           =   720
      End
      Begin VB.Label lblInputBusId 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����(&B):"
         Height          =   180
         Left            =   180
         TabIndex        =   11
         Top             =   1575
         Width           =   720
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "�������ڣ�"
         Height          =   225
         Left            =   180
         TabIndex        =   10
         Top             =   900
         Width           =   960
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   ";��վ(&D):"
         Height          =   180
         Left            =   180
         TabIndex        =   9
         Top             =   3195
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ʼ����(&D):"
         Height          =   180
         Left            =   180
         TabIndex        =   8
         Top             =   270
         Width           =   1080
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5265
      Top             =   2070
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEnBusQuery.frx":0371
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEnBusQuery.frx":04CD
            Key             =   "FlowRun"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEnBusQuery.frx":0629
            Key             =   "FlowStop"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEnBusQuery.frx":0785
            Key             =   "Merge"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEnBusQuery.frx":08E1
            Key             =   "Replace"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEnBusQuery.frx":0A3D
            Key             =   "Checking"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEnBusQuery.frx":0B99
            Key             =   "Checked"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEnBusQuery.frx":0CF5
            Key             =   "ExCheck"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEnBusQuery.frx":0E51
            Key             =   "Run"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEnBusQuery.frx":0FAD
            Key             =   "SlitpBus"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvBus 
      Height          =   5790
      Left            =   1935
      TabIndex        =   0
      Top             =   360
      Width           =   7230
      _ExtentX        =   12753
      _ExtentY        =   10213
      SortKey         =   1
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "���δ���"
         Object.Width           =   1589
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "����ʱ��"
         Text            =   "����ʱ��"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "����λ��"
         Object.Width           =   1589
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "��������"
         Object.Width           =   1589
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "�ƻ�����"
         Object.Width           =   1589
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "δ�۶���"
         Object.Width           =   1589
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Ԥ������"
         Object.Width           =   1589
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "����Ԥ��"
         Object.Width           =   1588
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "������λ"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "��·��"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "�յ�վ"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "����"
         Object.Width           =   1411
      EndProperty
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "��λ��Ϣ��"
      Height          =   195
      Left            =   1980
      TabIndex        =   14
      Top             =   135
      Width           =   1230
   End
End
Attribute VB_Name = "FrmEnBusQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'**********************************************************
'* Source File Name:frmEnvBus.frm
'* Project Name:RTBusMan
'* Engineer:�ΰ
'* Data Generated:2002/08/27
'* Last Revision Date:2002/08/30
'* Brief Description:��ʱ���Ȳ�ѯ��ѯ
'* Relational Document:UI_BS_SM_23.DOC
'**********************************************************
Option Explicit
Private m_oREScheme As New REScheme
Public bIsShow As Boolean
Public Sub QueryBus(bflg As Boolean)

'Dim szaBus() As String
'Dim liTemp As ListItem
'Dim szDestStation As String, szRoute As String
'Dim nCount As Integer, eBusStatus As EREBusStatus
'Dim i As Integer, nDay As Integer, j As Integer
'Dim nResult As VbMsgBoxResult
'Dim dtDay As Date
'Dim szMsg As String
'Dim szMsgNodata As String
'
'On Error GoTo ErrorHandle
'
'nDay = DateDiff("d", dtpStartDate.Value, dtpEndDate.Value)
'szDestStation = ResolveDisplay(cboStation.Text)
'szRoute = IIf(ResolveDisplay(txtRoute.Text) = "(ȫ��)", "", ResolveDisplay(txtRoute.Text))
'szDestStation = IIf(ResolveDisplay(szDestStation) = "(ȫ��)", "", ResolveDisplay(szDestStation))
'
'If nDay > 1 And txtBusID.Text = "" And szDestStation = "" And szRoute = "" Then
'    nResult = MsgBox("�Ƿ��ѯ[" & nDay & "]���ڵĳ���", vbQuestion + vbYesNo + vbDefaultButton2, "����")
'    If nResult = vbNo Then Exit Sub
'End If
'
'Setbusy
'lvBus.ListItems.Clear
'
'For j = 0 To nDay
'    dtDay = DateAdd("d", j, dtpStartDate.Value)
'    showsbinfo "���" & Format(dtDay, "YYYY-MM-DD") & "�ĳ���"
'    szaBus = m_oREScheme.GetReBusBookAndReserveSeatInfo(dtDay, txtBusID.Text, szRoute, szDestStation, bflg)
'    nCount = ArrayLength(szaBus)
'    If nCount <> 0 Then
'        WriteProcessBar , nCount, , True
'    Else
'        szMsgNodata = szMsgNodata & "������Ϊ" & dtDay & "����,û������Ҫ������" & Chr(10)
'    End If
'
'    For i = 0 To nCount - 1
'            showsbinfo "���" & szaBus(i, 1) & "�ĳ�����Ϣ", , i
'
'            Set liTemp = lvBus.ListItems.Add(, , szaBus(i, 0), , "Run")
'            liTemp.subitems()= Format(szaBus(i, 1), "YYYY-MM-DD HH:MM")
'             liTemp.subitems()= Trim(szaBus(i, 2))
'            liTemp.subitems()= Trim(szaBus(i, 3))
'            liTemp.subitems()= szaBus(i, 4)
'            liTemp.subitems()= szaBus(i, 5)
'            liTemp.subitems()= szaBus(i, 6)
'            liTemp.subitems()= szaBus(i, 7)
'            liTemp.subitems()= szaBus(i, 12)
'            liTemp.subitems()= szaBus(i, 8)
'            liTemp.subitems()= szaBus(i, 9)
'            liTemp.subitems()= szaBus(i, 10)
'
'            eBusStatus = Val(szaBus(i, 11))
'            Select Case eBusStatus
'                    Case ST_BusStopped: liTemp.SmallIcon = "Stop"
'                    Case ST_BusChecking
'                    liTemp.SmallIcon = "Checking"
'                    Case ST_BusExtraChecking
'                    liTemp.SmallIcon = "ExCheck"
'                    Case ST_BusStopCheck
'                    liTemp.SmallIcon = "Checked"
'                    Case ST_BusReplace
'                      liTemp.SmallIcon = "Replace"
'                    Case ST_BusSlitpStop
'                      liTemp.SmallIcon = "Merge"
'            End Select
'       Next
'Next
'
'
'
'    If cboStation.Text <> "" And cboStation.Text <> "(ȫ��)" Then
'
'        If bflg = True Then
'           szMsg = "��ѯ���۵ġ�����վ��" & cboStation.Text & "�ĳ�����Ϣ"
'        Else
'          szMsg = "��ѯ����վ��" & cboStation.Text & " �ĳ�����Ϣ"
'        End If
'
'    Else
'
'         If txtRoute.Text <> "" And txtRoute.Text <> "(ȫ��)" Then
'
'             szMsg = "��ѯ������·" & txtRoute.Text & " �ĳ�����Ϣ"
'
'         Else
'
'             If txtBusID.Text <> "" Then
'
'                If Len(txtBusID.Text) < 5 Then
'                   szMsg = "��ѯ���δ���ǰ׺Ϊ" & txtBusID.Text & "�ĳ�����Ϣ"
'                End If
'
'             End If
'             szMsg = "��ѯ" & txtBusID.Text & "������Ϣ"
'
'         End If
'
'    End If
'
'    WriteProcessBar szMsg, 0, 0
'    If szMsgNodata <> "" Then
'        MsgBox szMsgNodata, vbInformation, Me.Caption
'    End If
'
'    If g_bflgImmediatelyQuery <> True Then
'        txtBusID.Text = ""
'        txtRoute.Text = "(ȫ��)"
'        cboStation.Clear
'        cboStation.Text = "(ȫ��)"
'    End If
'
'    Setnormal
'Exit Sub
'ErrorHandle:
'    Setnormal
'    ShowErrorMsg
End Sub
Private Sub cboStation_KeyDown(KeyCode As Integer, Shift As Integer)
' If KeyCode = 13 Then
'   FindBusEx cboStation, cszFrmREBusEx, g_bflgImmediatelyQuery
' End If
End Sub




Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdQuery_Click()
  QueryBus True
  
  'FindBusEx cboStation, cszFrmREBusEx, g_bflgImmediatelyQuery

End Sub

Private Sub Form_Load()
bIsShow = True
 dtpStartDate = Date
 dtpEndDate = Date
 m_oREScheme.Init g_oActiveUser
End Sub



Private Sub txtBusID_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
       Case vbKeyReturn
       cmdQuery_Click
End Select
End Sub



Private Sub txtRoute_Click()
 Dim oShell As New STShell.CommDialog
    Dim aszTemp() As String
    oShell.Init g_oActiveUser
    aszTemp = oShell.SelectRoute(False)
    Set oShell = Nothing
    If ArrayLength(aszTemp) = 0 Then Exit Sub
    txtRoute.Text = aszTemp(1, 1) & "[" & aszTemp(1, 2) & "]"

End Sub


Private Sub txtRoute_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
       Case vbKeyReturn
       cmdQuery_Click
End Select
End Sub
