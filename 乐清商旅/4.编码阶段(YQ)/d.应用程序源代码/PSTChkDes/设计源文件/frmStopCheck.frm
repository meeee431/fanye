VERSION 5.00
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Begin VB.Form frmStopCheck 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ֹͣ��Ʊ"
   ClientHeight    =   1485
   ClientLeft      =   3780
   ClientTop       =   4365
   ClientWidth     =   4365
   ControlBox      =   0   'False
   Icon            =   "frmStopCheck.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin RTComctl3.CoolButton cmdYes 
      Default         =   -1  'True
      Height          =   315
      Left            =   1050
      TabIndex        =   2
      Top             =   1050
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "��(&Y)"
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
      MICON           =   "frmStopCheck.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.CheckBox chkAutoPrint 
      BackColor       =   &H00E0E0E0&
      Caption         =   "(&A)ͣ���ֱ�Ӵ�ӡ·��"
      Height          =   345
      Left            =   1005
      TabIndex        =   1
      Top             =   510
      Visible         =   0   'False
      Width           =   2355
   End
   Begin RTComctl3.CoolButton cmdNo 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   2340
      TabIndex        =   3
      Top             =   1050
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "��(&N)"
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
      MICON           =   "frmStopCheck.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   210
      Picture         =   "frmStopCheck.frx":0044
      Top             =   225
      Width           =   480
   End
   Begin VB.Label lblMessage 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "115�����ѵ�ͣ��ʱ��,�Ƿ�ͣ��?"
      Height          =   180
      Left            =   1005
      TabIndex        =   0
      Top             =   255
      Width           =   2610
   End
End
Attribute VB_Name = "frmStopCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mbAutoPrint As Boolean      '�Ƿ�ֱ�Ӵ�ӡ·����־
Dim mnMessageStyle As Integer '��ʾ��Ϣ����
Dim mnClickButton As Integer '�û����µİ�ť
Dim mszBusID As String 'ͣ��ĳ���
Dim mbOldAutoPrint As Boolean
Public Property Let BusID(ByVal vNewValue As String)
    mszBusID = vNewValue
End Property

Public Property Get MessageStyle() As Integer
    MessageStyle = mnMessageStyle
End Property

Public Property Let MessageStyle(ByVal vNewValue As Integer)
    mnMessageStyle = vNewValue
End Property

Public Property Let AutoPrint(ByVal vNewValue As Boolean)
    mbAutoPrint = vNewValue
End Property

Public Property Get AutoPrint() As Boolean
    AutoPrint = mbAutoPrint
End Property

Public Property Get ClickButton() As Integer
    ClickButton = mnClickButton
End Property

Private Sub chkAutoPrint_Click()
    If chkAutoPrint.Value = vbChecked Then
        mbAutoPrint = True
    Else
        mbAutoPrint = False
    End If
End Sub

Private Sub cmdNo_Click()
    mnClickButton = vbNo
'    Me.Hide
    Unload Me
End Sub

Private Sub CmdYes_Click()
    mnClickButton = vbYes
'    Me.Hide
    Unload Me
End Sub
Public Sub RefreshForm()
 '���ִ���
    Select Case mnMessageStyle
        Case 0      '�Զ�ͣ�췽ʽ
            lblMessage.Caption = mszBusID & "�����ѵ�ͣ��ʱ�䣬�Ƿ�ͣ�죿"
        Case 1      '�ֶ�ͣ�췽ʽ
            lblMessage.Caption = "�Ƿ�ͣ��" & mszBusID & "���Σ�"
    End Select
    If chkAutoPrint.Value = vbChecked Then
        mbAutoPrint = True
    Else
        mbAutoPrint = False
    End If
    mbOldAutoPrint = mbAutoPrint
End Sub
Private Sub Form_Load()
    Dim oReg As New CFreeReg
    Dim szTemp As String
    
    oReg.Init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany
    szTemp = oReg.GetSetting(m_cRegSystemKey, "AutoPrint", "0")
    
    If szTemp = "1" Then
        chkAutoPrint.Value = vbChecked
    End If
    
    mnClickButton = vbYes
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If mbAutoPrint <> mbOldAutoPrint Then
        Dim oReg As New CFreeReg
        oReg.Init cszRegKeyProduct, HKEY_LOCAL_MACHINE, cszRegKeyCompany
        oReg.SaveSetting m_cRegSystemKey, "AutoPrint", _
            IIf(mbAutoPrint, "1", "0")
    End If
End Sub
