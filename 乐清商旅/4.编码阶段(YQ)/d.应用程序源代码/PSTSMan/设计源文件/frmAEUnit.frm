VERSION 5.00
Begin VB.Form frmAEUnit 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������λ��༭��λ����"
   ClientHeight    =   4125
   ClientLeft      =   2070
   ClientTop       =   2370
   ClientWidth     =   6315
   HelpContextID   =   50000160
   Icon            =   "frmAEUnit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "����(&H)"
      Height          =   315
      Left            =   2280
      TabIndex        =   21
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   315
      Left            =   3600
      TabIndex        =   20
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "�˳�(&E)"
      CausesValidation=   0   'False
      Height          =   315
      Left            =   4920
      TabIndex        =   19
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox txtSellCharge 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5100
      TabIndex        =   9
      Text            =   "0"
      Top             =   878
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "������λ��������"
      Height          =   1230
      Left            =   150
      TabIndex        =   17
      Top             =   1320
      Width           =   5985
      Begin VB.CommandButton cmdAgent 
         Caption         =   "��λԶ������(&R)"
         Height          =   315
         Left            =   3630
         TabIndex        =   12
         Top             =   285
         Width           =   2160
      End
      Begin PSTSMan.ucIPAddress ucIPUnit 
         Height          =   300
         Left            =   1110
         TabIndex        =   11
         Top             =   300
         Width           =   2340
         _ExtentX        =   4128
         _ExtentY        =   529
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IP��ַ(&I):"
         Height          =   180
         Left            =   105
         TabIndex        =   10
         Top             =   345
         Width           =   900
      End
      Begin VB.Label lblAnno 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   105
         TabIndex        =   18
         Top             =   690
         Width           =   5685
         WordWrap        =   -1  'True
      End
   End
   Begin VB.ComboBox cboUnitType 
      Height          =   300
      Left            =   1260
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   885
      Width           =   2355
   End
   Begin VB.TextBox txtUnitShortName 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4875
      TabIndex        =   3
      Top             =   135
      Width           =   1275
   End
   Begin VB.TextBox txtUnitAnnotation 
      Appearance      =   0  'Flat
      Height          =   540
      Left            =   150
      MultiLine       =   -1  'True
      TabIndex        =   14
      Top             =   2910
      Width           =   5985
   End
   Begin VB.TextBox txtUnitFullName 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1260
      TabIndex        =   5
      Top             =   510
      Width           =   4890
   End
   Begin VB.TextBox txtUnitID 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1260
      TabIndex        =   1
      Top             =   135
      Width           =   2355
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��Ʊ������(G):"
      Height          =   180
      Left            =   3765
      TabIndex        =   8
      Top             =   945
      Width           =   1260
   End
   Begin VB.Label lblUnitID 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   180
      Left            =   1260
      TabIndex        =   16
      Top             =   195
      Width           =   90
   End
   Begin VB.Label lblSelfUnit 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�˵�λ���Ǳ���λ"
      Height          =   180
      Left            =   1725
      TabIndex        =   15
      Top             =   945
      Width           =   1455
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��λ����(&T):"
      Height          =   180
      Left            =   150
      TabIndex        =   6
      Top             =   945
      Width           =   1080
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��λע��(&A):"
      Height          =   180
      Left            =   150
      TabIndex        =   13
      Top             =   2655
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��λȫ��(&F):"
      Height          =   180
      Left            =   150
      TabIndex        =   4
      Top             =   570
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��λ���(&S):"
      Height          =   180
      Left            =   3765
      TabIndex        =   2
      Top             =   195
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��λ����(&C):"
      Height          =   180
      Left            =   150
      TabIndex        =   0
      Top             =   195
      Width           =   1080
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   6185
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   6185
      Y1              =   3585
      Y2              =   3585
   End
End
Attribute VB_Name = "frmAEUnit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'  *******************************************************************
' *  Source File Name  : frmAEUnit                                  *
' *  Project Name: PSTSMan                                    *
' *  Engineer:                          *
' *  Date Generated: 2002/08/19                      *
' *  Last Revision Date : 2002/08/19             *
' *  Brief Description   : ��ӵ�λ��༭��λ����                   *
' *******************************************************************

Option Explicit
Public bEdit As Boolean
Dim nUnitType As EUnitType
Dim szIPs As String
Dim szUnitID As String
Dim szUnitFullName As String
Dim szUnitShortName As String
Dim szUnitAnnotation As String
Dim dbSellCharge As Double


Private Sub cboUnitType_Click()
    Select Case cboUnitType.Text
    Case "2--������Ʊ��λ"
        nUnitType = TP_UnitSC
        ucIPUnit.Enabled = True
        lblAnno.Caption = "������Ʊ��λ�뱾��λ�Ĺ�ϵΪ: ����λ�ɴ��۴˵�λ�ĳ�Ʊ,ͬʱ�˵�λ�ɴ��۱���λ��Ʊ.��IP��ַ������λ�û���¼�˵�λ�ķ�������ַ."
    Case "1--���۳�Ʊ��λ"
        nUnitType = TP_UnitClient
        ucIPUnit.Enabled = False
        lblAnno.Caption = "���۳�Ʊ��λ�뱾��λ�Ĺ�ϵΪ: ����λ�Ǵ˵�λ����Ʊ�����ṩ��λ,�˵�λ������λ������Ʊ.IP��ַ����Ч."
    Case "0--��Ʊ�����ṩ��λ"
        nUnitType = TP_UnitServer
        ucIPUnit.Enabled = True
        lblAnno.Caption = "��Ʊ�����ṩ��λ�뱾��λ�Ĺ�ϵΪ: ����λ�Ǵ˵�λ���۳�Ʊ��λ,����λ����˵�λ������Ʊ.��IP��ַ������λ�û���¼�˵�λ�ķ�������ַ."
    End Select
End Sub

Private Sub cmdAgent_Click()
    frmUnitBeUser.Show vbModal

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim oUnit As New Unit
    On Error GoTo ErrorHandle
'    oTemp.Init g_oActUser
    GetInfoFromUI
    If bEdit = True Then
        '�޸��û�
        oUnit.Init g_oActUser
        oUnit.Identify szUnitID
        oUnit.HostName = szIPs
        oUnit.UnitAnnotation = szUnitAnnotation
        oUnit.UnitFullName = szUnitFullName
        oUnit.UnitShortName = szUnitShortName
        oUnit.UnitType = nUnitType
        oUnit.SellCharge = dbSellCharge
        oUnit.Update
    Else
        '�����û�
        oUnit.Init g_oActUser
        oUnit.AddNew
        oUnit.HostName = szIPs
        oUnit.UnitAnnotation = szUnitAnnotation
        oUnit.UnitFullName = szUnitFullName
        oUnit.UnitShortName = szUnitShortName
        oUnit.UnitType = nUnitType
        oUnit.UnitID = szUnitID
        oUnit.SellCharge = dbSellCharge
        oUnit.Update
    End If
    frmStoreMenu.LoadCommonData
    frmStoreMenu.LoadUnitInfo
    Dim i As Integer
    For i = 1 To frmSMCMain.lvDetail.ListItems.Count
        If frmSMCMain.lvDetail.ListItems.Item(i).Key = "A" & oUnit.UnitID Then
            frmSMCMain.lvDetail.ListItems.Item(i).Selected = True
        Else
            frmSMCMain.lvDetail.ListItems.Item(i).Selected = False
        End If
    Next
    Set oUnit = Nothing
    Unload Me
Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub


Private Sub Command1_Click()
DisplayHelp Me
End Sub

Private Sub Form_Load()

    Me.Top = (Screen.Height - Me.ScaleHeight) / 2
    Me.Left = (Screen.Width - Me.ScaleWidth) / 2
    
    cboUnitType.AddItem "0--��Ʊ�����ṩ��λ", 0
    cboUnitType.AddItem "1--���۳�Ʊ��λ", 1
    cboUnitType.AddItem "2--������Ʊ��λ", 2
    If bEdit Then
        '�޸�
        Me.Caption = "�޸ĵ�λ����"
        cmdOk.Caption = "ȷ��(&O)"
        cmdCancel.Caption = "ȡ��(&C)"
        txtUnitID.Visible = False
        lblUnitID.Visible = True
        LoadUnitInfo
        frmAEUnit.HelpContextID = 50000210
    Else
        '����
        Me.Caption = "������λ"
        cmdOk.Caption = "����(&A)"
        cmdCancel.Caption = "�ر�(&C)"
        
        lblUnitID.Visible = False
        txtUnitID.Visible = True
        lblSelfUnit.Visible = False
        cmdAgent.Enabled = False
        cboUnitType.ListIndex = 2
        ClearTextBox Me
        frmAEUnit.HelpContextID = 50000160
    End If
End Sub


Private Sub txtUnitAnnotation_Validate(Cancel As Boolean)
    If TextLongValidate(255, txtUnitFullName.Text) Then Cancel = True
End Sub

Private Sub txtUnitFullName_Validate(Cancel As Boolean)
    If TextLongValidate(100, txtUnitFullName.Text) Then Cancel = True
    If SpacialStrValid(txtUnitFullName.Text, "[") Then Cancel = True
    If SpacialStrValid(txtUnitFullName.Text, ",") Then Cancel = True
    If SpacialStrValid(txtUnitFullName.Text, "]") Then Cancel = True

End Sub

Private Sub txtUnitID_Validate(Cancel As Boolean)
    If TextLongValidate(10, txtUnitID.Text) Then Cancel = True
    If SpacialStrValid(txtUnitID.Text, "[") Then Cancel = True
    If SpacialStrValid(txtUnitID.Text, ",") Then Cancel = True
    If SpacialStrValid(txtUnitID.Text, "]") Then Cancel = True

End Sub

Private Sub txtUnitShortName_Validate(Cancel As Boolean)

    If TextLongValidate(10, txtUnitShortName.Text) Then Cancel = True
    If SpacialStrValid(txtUnitShortName.Text, "[") Then Cancel = True
    If SpacialStrValid(txtUnitShortName.Text, ",") Then Cancel = True
    If SpacialStrValid(txtUnitShortName.Text, "]") Then Cancel = True


End Sub



Private Sub ucIPAddress1_Validate(Cancel As Boolean)
    Dim aTemp()  As String
    With ucIPUnit
        aTemp = .GetIPDistri
        If .TextNotValid(aTemp(1)) = True Then
            Cancel = True
            .SetFocus
        ElseIf .TextNotValid(aTemp(2)) = True Then
            Cancel = True
            .SetFocus
        ElseIf .TextNotValid(aTemp(3)) = True Then
            Cancel = True
            .SetFocus
        ElseIf .TextNotValid(aTemp(4)) = True Then
            Cancel = True
            .SetFocus
        Else
            Cancel = False
        End If
    
    End With
End Sub

Private Sub LoadUnitInfo()
    
    '���뵥λ��Ϣ
    Dim aszIPPart() As String
    Dim j As Integer
    Dim i As Integer
    
    
    lblUnitID.Caption = g_alvItemText(1)
    For i = 1 To ArrayLength(g_atAllUnit)
        If g_atAllUnit(i).szUnitID = g_alvItemText(1) Then
            If g_atAllUnit(i).szUnitID = g_szLocalUnit Then
                lblSelfUnit.Visible = True
                cmdAgent.Enabled = False
                cboUnitType.Enabled = False
                cboUnitType.Visible = False
                ucIPUnit.Enabled = True
                lblAnno.Caption = "����λIP��ַ���ⵥλ��¼����λ�ķ�����IP��ַ."
            Else
                lblSelfUnit.Visible = False
                cmdAgent.Enabled = True
                cboUnitType.Enabled = True
                cboUnitType.Visible = True
                Select Case g_atAllUnit(i).nUnitType
                    Case TP_UnitClient
                        cboUnitType.ListIndex = 1
                    Case TP_UnitSC
                        cboUnitType.ListIndex = 2
                    Case TP_UnitServer
                        cboUnitType.ListIndex = 0
                End Select
            End If
            
            txtUnitShortName.Text = g_atAllUnit(i).szUnitShortName
            txtUnitFullName.Text = g_atAllUnit(i).szUnitFullName
            txtUnitAnnotation.Text = g_atAllUnit(i).szAnnotation
            txtSellCharge.Text = g_atAllUnit(i).dbSellCharge
            aszIPPart = GetIPParts(g_atAllUnit(i).szIPAddress)
            '���IP��ַ
            j = ArrayLength(aszIPPart)
            If j = 4 Then
                For j = 1 To 4
                    ucIPUnit.SetIPDistri aszIPPart(j), j
                Next j
            End If
        End If
    Next i
Exit Sub
ErrorHandle:

    ShowErrorMsg
End Sub

Private Sub GetInfoFromUI()
    If bEdit = True Then
        szUnitID = lblUnitID.Caption
    Else
        szUnitID = txtUnitID.Text
    End If
    szUnitShortName = txtUnitShortName.Text
    szUnitFullName = txtUnitFullName.Text
    szUnitAnnotation = txtUnitAnnotation.Text
    If IsNumeric(txtSellCharge.Text) Then
        dbSellCharge = txtSellCharge.Text
    End If
    
    szIPs = ucIPUnit.GetIpAddress
    
End Sub
