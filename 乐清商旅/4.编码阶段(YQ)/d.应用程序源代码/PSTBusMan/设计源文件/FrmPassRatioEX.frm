VERSION 5.00
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmPassRatioEx 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ͨ�з�����"
   ClientHeight    =   4455
   ClientLeft      =   3510
   ClientTop       =   3600
   ClientWidth     =   5685
   ClipControls    =   0   'False
   FillStyle       =   0  'Solid
   Icon            =   "FrmPassRatioEX.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   5685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txtPassRatio 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4155
      TabIndex        =   5
      Top             =   2895
      Width           =   1050
   End
   Begin VB.ListBox lstSeatType 
      Appearance      =   0  'Flat
      Height          =   1500
      Left            =   3210
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   1260
      Width           =   1995
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   -60
      TabIndex        =   13
      Top             =   780
      Width           =   7215
   End
   Begin VB.PictureBox ptTitle 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   795
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   7185
      TabIndex        =   11
      Top             =   0
      Width           =   7185
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������޸�ͨ�з���Ϣ:"
         Height          =   180
         Left            =   180
         TabIndex        =   12
         Top             =   300
         Width           =   1890
      End
   End
   Begin RTComctl3.CoolButton cmdDelete 
      Height          =   330
      Left            =   1710
      TabIndex        =   10
      Top             =   3975
      Width           =   1140
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "ɾ��(&D)"
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
      MICON           =   "FrmPassRatioEX.frx":014A
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
      Height          =   330
      Left            =   4260
      TabIndex        =   9
      Top             =   3975
      Width           =   1140
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
      MICON           =   "FrmPassRatioEX.frx":0166
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdSave 
      Height          =   330
      Left            =   2985
      TabIndex        =   8
      Top             =   3975
      Width           =   1140
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "����(&S)"
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
      MICON           =   "FrmPassRatioEX.frx":0182
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ListBox lstVehicleModel 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   465
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   1260
      Width           =   2505
   End
   Begin VB.TextBox txtAnnotation 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1230
      TabIndex        =   7
      Top             =   3270
      Width           =   3975
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   " "
      Enabled         =   0   'False
      Height          =   1350
      Left            =   -150
      TabIndex        =   14
      Top             =   3690
      Width           =   8745
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "��ע(&A):"
      Height          =   195
      Left            =   465
      TabIndex        =   6
      Top             =   3315
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "λ��(&S):"
      Height          =   285
      Left            =   3210
      TabIndex        =   2
      Top             =   960
      Width           =   1005
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "����(&T):"
      Height          =   240
      Left            =   480
      TabIndex        =   0
      Top             =   990
      Width           =   1050
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ͨ�з�(&P):"
      Height          =   195
      Left            =   3210
      TabIndex        =   4
      Top             =   2955
      Width           =   960
   End
   Begin VB.Menu pmnu_VehicleModel 
      Caption         =   "����"
      Visible         =   0   'False
      Begin VB.Menu pmnu_SelectAllVehicleModel 
         Caption         =   "ѡ������(&A)"
      End
      Begin VB.Menu pmnu_UnSelectAllVehicleModel 
         Caption         =   "ȡ������(&U)"
      End
   End
   Begin VB.Menu pmnu_SeatType 
      Caption         =   "λ��"
      Visible         =   0   'False
      Begin VB.Menu pmnu_SelectAllSeatType 
         Caption         =   "ѡ������(&A)"
      End
      Begin VB.Menu pmnu_UnSelectAllSeatType 
         Caption         =   "ȡ������(&U)"
      End
   End
End
Attribute VB_Name = "frmPassRatioEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public szSectionID As String

Private oBaseInfo As New BaseInfo
Private oSection As New Section
Private m_aszSeatType() As String '���е���λ����
Private m_aszVehicleModal() As String '���еĳ���
'
'Private m_anSeatType() As Integer
'Private m_anVehicleModal() As Integer

'Private m_anEmptyArray() As Integer '������,Ϊ���������


Private Sub cmdDelete_Click()
    '***********�˴�Ӧ�ø�һ�½ӿ�DeletePassCharge,Ӧ����������
'    Dim nVehicleModel As Integer
'    Dim nSeatType As Integer
    Dim i As Integer, j As Integer
'    Dim nCount As Integer
    On Error GoTo ErrorHandle
    If MsgBox("ɾ��ѡ����ķ���?", vbQuestion + vbYesNo + vbDefaultButton2, "���ʹ���") = vbNo Then Exit Sub
    SetBusy
    oSection.Identify szSectionID
    For i = 0 To lstVehicleModel.ListCount - 1
        If lstVehicleModel.Selected(i) Then
            lstVehicleModel.ListIndex = i
            For j = 0 To lstSeatType.ListCount - 1
                If lstSeatType.Selected(j) Then
                    lstSeatType.ListIndex = j
                    oSection.DeletePassCharge Trim(m_aszVehicleModal(i + 1, 1)), m_aszSeatType(j + 1, 1)
                End If
            Next
        End If
    Next
    SetNormal
    MsgBox "ѡ����ķ���ɾ���ɹ�", vbInformation, "����"
    Exit Sub
ErrorHandle:
    SetNormal
    ShowErrorMsg
End Sub

Private Sub cmdSave_Click()
    '***********�˴�Ӧ�ø�һ�½ӿ�DeletePassCharge,Ӧ����������
'    Dim nVehicleModel As Integer
'    Dim nSeatType As Integer
    Dim i, j As Integer
'    Dim nCount As Integer
    Dim tChargeInfo As TTransitChargeInfo
    On Error GoTo ErrorHandle
    If MsgBox("�޸�ѡ�����ͨ�зѷ���?", vbQuestion + vbYesNo + vbDefaultButton2, "���ʹ���") = vbNo Then Exit Sub
    SetBusy
'    nSeatType = ArrayLength(m_anSeatType)
'    nVehicleModel = ArrayLength(m_anVehicleModal)
    oSection.Identify szSectionID
'    showsbinfo "����...", (nVehicleModel - 1), , True
    For i = 0 To lstVehicleModel.ListCount - 1
        If lstVehicleModel.Selected(i) Then
            lstVehicleModel.ListIndex = i
            For j = 0 To lstSeatType.ListCount - 1
                If lstSeatType.Selected(j) Then
                    lstSeatType.ListIndex = j
                    tChargeInfo.szVehicleType = Trim(m_aszVehicleModal(i + 1, 1))
                    tChargeInfo.szPassCharge = Val(txtPassRatio.Text)
                    tChargeInfo.szSeatType = m_aszSeatType(j + 1, 1)
                    tChargeInfo.szAnnotation = txtAnnotation.Text
                    tChargeInfo.szSection = szSectionID
                    oSection.ModifyPassCharge tChargeInfo
                End If
            Next j
        End If
    Next i
    WriteProcessBar False
    SetNormal
    MsgBox "ͨ�зѷ����޸ĳɹ�", vbInformation, "ͨ�зѷ���"
    Exit Sub
ErrorHandle:
    SetNormal
    ShowErrorMsg
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    AlignFormPos Me
    On Error GoTo ErrorHandle
    oBaseInfo.Init g_oActiveUser
    oSection.Init g_oActiveUser
    FillVehicleModel
    FillSeatType
    IsSave
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub FillSeatType()
    '�����λ����
    Dim nCount As Integer
    Dim i As Integer
    m_aszSeatType = oBaseInfo.GetAllSeatType
    nCount = ArrayLength(m_aszSeatType)
    For i = 1 To nCount
        lstSeatType.AddItem MakeDisplayString(m_aszSeatType(i, 1), m_aszSeatType(i, 2))
    Next i
End Sub

Private Sub FillVehicleModel()
    '��䳵��
    Dim nCount As Integer
    Dim i As Integer
    m_aszVehicleModal = oBaseInfo.GetAllVehicleModel
    nCount = ArrayLength(m_aszVehicleModal)
    For i = 1 To nCount
        lstVehicleModel.AddItem m_aszVehicleModal(i, 2)
    Next i
    
End Sub

Private Sub IsSave()
    If txtPassRatio.Text = "" Or lstVehicleModel.SelCount = 0 Or lstSeatType.SelCount = 0 Then
        cmdSave.Enabled = False
    Else
        cmdSave.Enabled = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveFormPos Me
End Sub

Private Sub lstSeatType_ItemCheck(Item As Integer)
    IsSave
End Sub

Private Sub lstSeatType_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        PopupMenu pmnu_SeatType
    End If
End Sub

Private Sub lstVehicleModel_ItemCheck(Item As Integer)
    IsSave
    
End Sub


Private Sub lstVehicleModel_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyA And Shift = 2 Then
        '������Ctrl+A,��ȫѡ��
'        Debug.Print "Ctrl+A"
        SelAllVehicleModel
    End If
End Sub

Private Sub lstVehicleModel_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        PopupMenu pmnu_VehicleModel
    End If
End Sub

Private Sub pmnu_SelectAllSeatType_Click()
    SelAllSeatType True
    
End Sub

Private Sub pmnu_SelectAllVehicleModel_Click()
    SelAllVehicleModel True
End Sub

Private Sub pmnu_UnSelectAllSeatType_Click()
    SelAllSeatType False
End Sub

Private Sub pmnu_UnSelectAllVehicleModel_Click()
    SelAllVehicleModel False
End Sub

Private Sub txtPassRatio_Change()
    IsSave
End Sub


Private Sub SelAllVehicleModel(Optional pbSelected As Boolean = True)
    'ѡ���ȡ��ѡ�����г���
    Dim i As Integer
    For i = 1 To lstVehicleModel.ListCount
        lstVehicleModel.Selected(i - 1) = pbSelected
    Next i
End Sub

Private Sub SelAllSeatType(Optional pbSelected As Boolean = True)
    'ѡ���ȡ��ѡ�����г���
    Dim i As Integer
    For i = 1 To lstSeatType.ListCount
        lstSeatType.Selected(i - 1) = pbSelected
    Next i
End Sub

