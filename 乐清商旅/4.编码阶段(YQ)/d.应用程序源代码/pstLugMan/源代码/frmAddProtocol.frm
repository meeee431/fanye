VERSION 5.00
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmAddProtocol 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����Э��"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6300
   Icon            =   "frmAddProtocol.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   6300
   StartUpPosition =   3  '����ȱʡ
   Begin VB.OptionButton Option2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "��Ĭ��Э��"
      Height          =   315
      Left            =   4170
      TabIndex        =   17
      Top             =   2970
      Width           =   1545
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "����Ĭ��Э��"
      Height          =   345
      Left            =   2670
      TabIndex        =   16
      Top             =   2970
      Width           =   1485
   End
   Begin VB.OptionButton Option0 
      BackColor       =   &H00E0E0E0&
      Caption         =   "���Ĭ��Э��"
      Height          =   345
      Left            =   1200
      TabIndex        =   15
      Top             =   2970
      Width           =   1455
   End
   Begin RTComctl3.CoolButton cmdok 
      Height          =   330
      Left            =   3510
      TabIndex        =   3
      Top             =   3720
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   582
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
      MICON           =   "frmAddProtocol.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdSetProtocol 
      Default         =   -1  'True
      Height          =   330
      Left            =   1980
      TabIndex        =   14
      Top             =   3720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      BTYPE           =   3
      TX              =   "����Э����Ŀ"
      ENAB            =   0   'False
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
      MICON           =   "frmAddProtocol.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtAnnotation 
      Appearance      =   0  'Flat
      Height          =   510
      Left            =   1980
      TabIndex        =   2
      Top             =   2130
      Width           =   3345
   End
   Begin VB.TextBox txtProtocolName 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1995
      TabIndex        =   1
      Top             =   1635
      Width           =   3315
   End
   Begin VB.TextBox txtProtocolID 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1995
      MaxLength       =   4
      TabIndex        =   0
      Top             =   1185
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   -180
      TabIndex        =   8
      Top             =   810
      Width           =   8775
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   825
      Left            =   -120
      ScaleHeight     =   825
      ScaleWidth      =   8685
      TabIndex        =   5
      Top             =   0
      Width           =   8685
      Begin VB.Label lblContent 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������Э��������Ϣ�������ΪĬ��Э�飬����Э���Զ�����Э�顣"
         Height          =   180
         Left            =   450
         TabIndex        =   7
         Top             =   420
         Width           =   5760
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Э����Ϣ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   270
         TabIndex        =   6
         Top             =   120
         Width           =   780
      End
   End
   Begin RTComctl3.CoolButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   330
      Left            =   4800
      TabIndex        =   4
      Top             =   3720
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   582
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
      MICON           =   "frmAddProtocol.frx":0044
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   " "
      Enabled         =   0   'False
      Height          =   2880
      Left            =   -120
      TabIndex        =   9
      Top             =   3360
      Width           =   9465
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ĭ��Э��(&D):"
      Height          =   255
      Left            =   630
      TabIndex        =   13
      Top             =   2670
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Э�����(&I):"
      Height          =   285
      Left            =   630
      TabIndex        =   12
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Э������(&B):"
      Height          =   180
      Left            =   630
      TabIndex        =   11
      Top             =   1695
      Width           =   1080
   End
   Begin VB.Label Label9 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "ע��(&A):"
      Height          =   195
      Left            =   630
      TabIndex        =   10
      Top             =   2130
      Width           =   1050
   End
End
Attribute VB_Name = "frmAddProtocol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status As eFormStatus
Public iChkused As Integer
 '�������Ͷ��� A
Public mszProtocolID As String



Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
On Error GoTo ErrHandle
Dim szProtocol As Integer   'Ĭ��Э��
    If Option0.Value = True Then
        szProtocol = ELuggageProtocolDefault.LuggageProtocolDefaultGeneral
    ElseIf Option1.Value = True Then
        szProtocol = ELuggageProtocolDefault.LuggageProtocolDefaultMan
    ElseIf Option2.Value = True Then
        szProtocol = ELuggageProtocolDefault.LuggageProtocolNotDefault
    End If
    Select Case Status
        Case ST_AddObj
            '�����а�����
         
            m_oProtocol.AddNew
            m_oProtocol.ProtocolID = Trim(txtProtocolID.Text)
            m_oProtocol.ProtocolName = Trim(txtProtocolName.Text)
            m_oProtocol.Annotation = Trim(txtAnnotation.Text)
            m_oProtocol.Default = szProtocol
            m_oProtocol.Update
        Case ST_EditObj
            '�޸��а�����
            
            m_oProtocol.Identify mszProtocolID
            m_oProtocol.ProtocolID = Trim(txtProtocolID.Text)
            m_oProtocol.ProtocolName = Trim(txtProtocolName.Text)
            m_oProtocol.Annotation = Trim(txtAnnotation.Text)
            m_oProtocol.Default = szProtocol
            m_oProtocol.Update
    End Select
        
     '��ֵ���������У����ظ�������Ϣ����
    Dim aszInfo(0 To 3) As String
    aszInfo(0) = Trim(txtProtocolID.Text)
    aszInfo(1) = Trim(txtProtocolName.Text)
    aszInfo(2) = Trim(txtAnnotation.Text)

    
    
    'ˢ�»�����Ϣ����
    Dim oListItem As ListItem
    If Status = ST_EditObj Then
        frmBaseInfo.UpdateList aszInfo
        Unload Me
        Exit Sub
    End If
    If Status = ST_AddObj Then
        frmBaseInfo.AddList aszInfo
       RefreshLug
        txtProtocolID.SetFocus
        Exit Sub
    End If
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub cmdSetProtocol_Click()
    frmProtocol.m_protocolID = Trim(txtProtocolID.Text)
    frmProtocol.lbProtocol.Caption = Trim(txtProtocolID.Text)
    frmProtocol.lbProtocolName.Caption = Trim(txtProtocolName.Text)
    If Option2.Value = True Then
        frmProtocol.lbDefault.Caption = "��"
    Else
        frmProtocol.lbDefault.Caption = "��"
    End If
    frmProtocol.lbRemark.Caption = Trim(txtAnnotation.Text)
    frmProtocol.Show vbModal
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    
           Case vbKeyReturn
                SendKeys "{TAB}"
    End Select
End Sub
Private Sub Form_Load()
On Error GoTo ErrHandle
Dim rsTemp As Recordset
    '���ô���
    AlignFormPos Me
    m_oProtocol.Init m_oAUser
    Select Case Status
        Case ST_AddObj
           cmdOk.Caption = "����(&A)"
           cmdSetProtocol.Enabled = False
           RefreshLug
           Option0.Value = False
           Option1.Value = False
           Option2.Value = True
        Case ST_EditObj
           txtProtocolID.Enabled = False
           cmdSetProtocol.Enabled = True
           RefreshLug
           Set rsTemp = m_oProtocol.GetAllProtocol(Trim(txtProtocolID.Text))
           If rsTemp.RecordCount > 0 Then
               If FormatDbValue(rsTemp!default_mark) = ELuggageProtocolDefault.LuggageProtocolNotDefault Then
                    Option2.Value = True
               ElseIf FormatDbValue(rsTemp!default_mark) = ELuggageProtocolDefault.LuggageProtocolDefaultGeneral Then
                    Option0.Value = True
               ElseIf FormatDbValue(rsTemp!default_mark) = ELuggageProtocolDefault.LuggageProtocolDefaultMan Then
                    Option1.Value = True
               End If
           End If
    End Select
  cmdOk.Enabled = False

    
    Exit Sub
ErrHandle:
    Status = ST_AddObj
    ShowErrorMsg
End Sub

Public Sub RefreshLug()
    If Status = ST_AddObj Then
        Me.Caption = "����Э��"
        cmdOk.Caption = "����"
        txtProtocolID.Text = ""
        txtProtocolName.Text = ""
        txtAnnotation.Text = ""
    Else
        Me.Caption = "�޸�Э��"
        cmdOk.Caption = "�޸�"
        txtProtocolID.Text = mszProtocolID
        m_oProtocol.Identify Trim(mszProtocolID)
        txtProtocolName.Text = m_oProtocol.ProtocolName
        txtAnnotation.Text = m_oProtocol.Annotation
    End If

End Sub
Private Sub Form_Unload(Cancel As Integer)
    SaveFormPos Me
End Sub

Private Sub Option0_Click()
    IsSave
End Sub

Private Sub Option1_Click()
    IsSave
End Sub

Private Sub Option2_Click()
    IsSave
End Sub

Private Sub txtAnnotation_Change()
    IsSave
End Sub
Private Sub IsSave()
    If txtProtocolID.Text = "" Or txtProtocolName.Text = "" Then
        cmdOk.Enabled = False
    Else
        cmdOk.Enabled = True
    End If
End Sub
Private Sub txtProtocolID_Change()
    IsSave
    FormatTextBoxBySize txtProtocolID, 4
End Sub
Private Sub txtProtocolName_Change()
    IsSave
    FormatTextBoxBySize txtProtocolName, 20
End Sub

Public Sub SetAllProtocl(values As Integer)
Dim i As Integer, j As Integer
Dim nCount As Integer
Dim vbYesOrNo  As Variant
Dim aszTemp() As String
Dim aszReturn() As String
Dim aszReturn1() As String
Dim szVehicleProtocol() As TVehicleProtocol
nCount = ArrayLength(m_obase.GetVehicle)
ReDim aszReturn(1 To nCount, 1 To 10)
aszTemp = m_obase.GetVehicle
For i = 1 To nCount
   aszReturn(i, 1) = aszTemp(i, 1)
   aszReturn(i, 2) = aszTemp(i, 2)
Next i

Select Case values
       Case 0  '�����еĳ���ȡ��Ϊ��Э��
'            WriteProcessBar , i, nCount, "�õ�����[" & rsItems!chinese_name & "]"
           vbYesOrNo = MsgBox("�Ƿ����Ҫ��Ĭ��Э��" & "[" & Trim(txtProtocolName.Text) & "]" & "ȡ����?", vbQuestion + vbYesNo + vbDefaultButton2, "��Ϣ")
           If vbYesOrNo = vbYes Then
               ReDim aszReturn1(1 To nCount)
                For i = 1 To nCount
                   aszReturn1(i) = aszTemp(i, 1)
                Next i
                SetBusy
                m_oProtocol.DelVehicleProtocol aszReturn1, 1
                m_oProtocol.DelVehicleProtocol aszReturn1, 0
                SetNormal
           End If
                       
                
       Case 1  '�����еĳ���ָ��Ϊ��Э��
             
            ReDim szVehicleProtocol(1 To nCount)
          vbYesOrNo = MsgBox("�Ƿ����Ҫ��" & "[" & Trim(txtProtocolName.Text) & "]" & "���ó�ΪĬ�ϵĲ���Э����", vbQuestion + vbYesNo + vbDefaultButton2, "��Ϣ")
              If vbYesOrNo = vbYes Then
                     '�����г���ָ��Ĭ�ϵĲ���Э��
                     SetBusy
                     ShowSBInfo "���ڸ����еĳ�������,���Ժ�......" & "[" & Trim(txtProtocolName.Text) & "]", ESB_WorkingInfo
                    For i = 1 To nCount
'                     WriteProcessBar , i, nCount, "���ڸ�[" & szVehicleProtocol(i).VehicleLicense & "]" & "���ò���Э��,��ȴ���"
                         szVehicleProtocol(i).ProtocolID = Trim(txtProtocolID.Text)
                         szVehicleProtocol(i).VehicleID = aszReturn(i, 1)
                         szVehicleProtocol(i).VehicleLicense = aszReturn(i, 2)
                         szVehicleProtocol(i).AcceptType = 0
                    Next i
                   m_oProtocol.SetVehicleProtocol szVehicleProtocol
                     '���������ָ��Ĭ�ϵĲ���Э��
                    For i = 1 To nCount
'                     WriteProcessBar , i, nCount, "���ڸ�[" & szVehicleProtocol(i).VehicleLicense & "]" & "���ò���Э��,��ȴ���"
                         szVehicleProtocol(i).ProtocolID = Trim(txtProtocolID.Text)
                         szVehicleProtocol(i).VehicleID = aszReturn(i, 1)
                         szVehicleProtocol(i).VehicleLicense = aszReturn(i, 2)
                         szVehicleProtocol(i).AcceptType = 1
                    Next i
                    m_oProtocol.SetVehicleProtocol szVehicleProtocol
                    ShowSBInfo ""

                    SetNormal
                Else
                     m_oProtocol.Default = 0
                     SetNormal
              End If
   End Select
End Sub
