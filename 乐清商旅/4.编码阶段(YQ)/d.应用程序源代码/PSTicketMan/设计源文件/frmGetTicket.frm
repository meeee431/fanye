VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmGetTicket 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��Ʊ�Ǽ�"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7935
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   7935
   StartUpPosition =   1  '����������
   Begin VB.CheckBox chkIsUseing 
      BackColor       =   &H00E0E0E0&
      Caption         =   "�Ƿ����ھ�����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   19
      Top             =   2400
      Width           =   3615
   End
   Begin VB.TextBox txtEndTicket 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5280
      TabIndex        =   4
      Top             =   1320
      Width           =   2130
   End
   Begin VB.TextBox txtFirstChar1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   380
      Left            =   4680
      TabIndex        =   3
      Top             =   1310
      Width           =   495
   End
   Begin VB.TextBox txtFirstChar 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   380
      Left            =   4680
      TabIndex        =   1
      Top             =   830
      Width           =   495
   End
   Begin VB.ListBox lstU 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3420
      Left            =   1200
      TabIndex        =   0
      Top             =   840
      Width           =   2295
   End
   Begin VB.TextBox txtMemo 
      Height          =   1455
      Left            =   4680
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   2880
      Width           =   2775
   End
   Begin VB.TextBox txtStartTicket 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5280
      TabIndex        =   2
      Top             =   840
      Width           =   2130
   End
   Begin MSComCtl2.DTPicker dtpGetTicketDate 
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      Top             =   1800
      Width           =   2730
      _ExtentX        =   4815
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   22872064
      CurrentDate     =   37942
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   -90
      ScaleHeight     =   705
      ScaleWidth      =   8025
      TabIndex        =   9
      Top             =   0
      Width           =   8025
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   480
         TabIndex        =   10
         Top             =   240
         Width           =   90
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   -120
      TabIndex        =   8
      Top             =   690
      Width           =   8085
   End
   Begin RTComctl3.CoolButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   6390
      TabIndex        =   17
      Top             =   4680
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "ȡ��"
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
      MICON           =   "frmGetTicket.frx":0000
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
      Height          =   315
      Left            =   5040
      TabIndex        =   7
      Top             =   4680
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "ȷ��"
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
      MICON           =   "frmGetTicket.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton CoolButton1 
      Height          =   315
      Left            =   3690
      TabIndex        =   16
      Top             =   4680
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "����"
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
      MICON           =   "frmGetTicket.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ImageList imgRoute 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGetTicket.frx":0054
            Key             =   "Station"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGetTicket.frx":01AE
            Key             =   "NoSell"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   " "
      Enabled         =   0   'False
      Height          =   2520
      Left            =   -120
      TabIndex        =   18
      Top             =   4440
      Width           =   8745
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "��    ע:"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3600
      TabIndex        =   15
      Top             =   2880
      Width           =   1080
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "����Ʊ��:"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   14
      Top             =   1373
      Width           =   1095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "��ʼƱ��:"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3600
      TabIndex        =   13
      Top             =   900
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "��Ʊ����:"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3600
      TabIndex        =   12
      Top             =   1860
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "��ƱԱ:"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   11
      Top             =   870
      Width           =   840
   End
End
Attribute VB_Name = "frmGetTicket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Status As EFormStatus
Public m_bIsParent As Boolean '�Ƿ񸸴���ֱ�ӵ���
Public m_User As String
Public m_GetTicketDate As Date
Private Sub cmdCancel_Click()
    SaveFormPos Me
    Unload Me
End Sub
Private Sub cmdOk_Click()
On Error GoTo ErrHandle
    
    If IsSave = False Then Exit Sub
    Dim rsTemp As New Recordset
    Dim rsTemp1 As New Recordset
    Dim oSTTicketMan As New TicketMan
    Dim nCount As Integer
    oSTTicketMan.Init m_oActiveUser
    With rsTemp.Fields
        .Append "user_id", adChar, 12
        .Append "getticket_date", adDate
        .Append "user_name", adVarChar, 10
        .Append "firstticket_no", adVarChar, 20
        .Append "lastticket_no", adVarChar, 20
        .Append "Memo", adVarChar, 255
        .Append "isuseing", adBigInt
    End With
    rsTemp.Open
    rsTemp.AddNew
    rsTemp!user_id = ResolveDisplay(lstU.Text)
    rsTemp!user_name = ResolveDisplayEx(lstU.Text)
    rsTemp!getticket_date = dtpGetTicketDate.Value
    rsTemp!firstticket_no = UCase(txtFirstChar) & txtStartTicket.Text
    rsTemp!lastticket_no = UCase(txtFirstChar1) & txtEndTicket.Text
    rsTemp!Memo = txtMemo.Text
    If chkIsUseing.Value = vbChecked Then
        rsTemp!isuseing = 1
    Else
        rsTemp!isuseing = 0
    End If
    rsTemp.Update
    
    Select Case Status
        Case EFormStatus.EFS_AddNew
            Set rsTemp1 = oSTTicketMan.GetTicketInfo(dtpGetTicketDate.Value, dtpGetTicketDate.Value, ResolveDisplay(lstU.Text))
            If rsTemp1.RecordCount > 0 Then MsgBox "[" & ResolveDisplayEx(lstU.Text) & "]" & "��" & Format(dtpGetTicketDate.Value, "yyyy��MM��dd��") & "�Ѿ�����Ʊ��¼��", vbOKOnly + vbExclamation, "����": Exit Sub
            oSTTicketMan.SellerGetTicket rsTemp
            frmManTicketMan.AddList ResolveDisplay(lstU.Text), dtpGetTicketDate.Value
      Case EFormStatus.EFS_Modify
             oSTTicketMan.SellerUpdateGetTicket rsTemp
             frmManTicketMan.UpdateList ResolveDisplay(lstU.Text), dtpGetTicketDate.Value
             MsgBox "�޸ĳɹ�����ע��Ʊ�ţ�", vbOKOnly + vbInformation, "��ʾ"
             Unload Me
      End Select
Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
       Case vbKeyReturn
       SendKeys "{TAB}"
End Select
End Sub
Private Sub Form_Load()
    AlignFormPos Me
    Select Case Status
        Case EFormStatus.EFS_AddNew
            FillUser
            FillFirstChar
            cmdOk.Caption = "����(&A)"
            ClearTicketMan
        Case EFormStatus.EFS_Modify
            cmdOk.Caption = "�޸�(&E)"
            lstU.AddItem m_User
            If lstU.ListCount > 0 Then lstU.Selected(0) = True
            RefreshTicketMan ResolveDisplay(m_User), m_GetTicketDate
    End Select
'    cmdOk.Enabled = False
    Exit Sub
ErrHandle:
    Status = EFS_AddNew
    ShowErrorMsg
End Sub
Private Sub FillUser()
On Error GoTo ErrHandle
  Dim nlen As Integer, i As Integer
    Dim szTemp As String
    Dim g_aszUser() As String
    Dim g_atUserInfo() As TUserInfo  '����δ��ɾ����ǵ��û�
    Dim nSelect As Integer
    Dim k As Integer
    
    Dim g_oSysMan As New SystemMan
    g_oSysMan.Init m_oActiveUser
    g_atUserInfo = g_oSysMan.GetAllUser(True, False)
    nlen = ArrayLength(g_atUserInfo)
    If nlen <= 0 Then Exit Sub
    If m_oActiveUser.SellStationID = "" Then
        For i = 1 To nlen
            lstU.AddItem MakeDisplayString(g_atUserInfo(i).UserID, g_atUserInfo(i).UserName)
            If Trim(g_atUserInfo(i).UserID) = Trim(m_oActiveUser.UserID) Then
                nSelect = i - 1
            End If
        Next i
    Else
        For i = 1 To nlen
            If m_oActiveUser.SellStationID = g_atUserInfo(i).SellStationID Then
                lstU.AddItem MakeDisplayString(g_atUserInfo(i).UserID, g_atUserInfo(i).UserName)
                k = k + 1
                If Trim(g_atUserInfo(i).UserID) = Trim(m_oActiveUser.UserID) Then
                    nSelect = k - 1
                End If
            End If
        Next i
    End If
    If lstU.ListCount > 0 Then lstU.ListIndex = nSelect
Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Function IsSave() As Boolean
    If txtStartTicket.Text = "" Then
        MsgBox "��ʼƱ�Ų���Ϊ�գ�������������ȷ����ʼƱ�ţ�", vbOKOnly + vbExclamation, "����"
        IsSave = False
        txtStartTicket.SetFocus
        Exit Function
    Else
        IsSave = True
    End If
    If txtEndTicket.Text = "" Then
        MsgBox "����Ʊ�Ų���Ϊ�գ�������������ȷ�Ľ���Ʊ�ţ�", vbOKOnly + vbExclamation, "����"
        IsSave = False
        txtEndTicket.SetFocus
        Exit Function
    Else
        IsSave = True
    End If
    
    If txtStartTicket.Text > txtEndTicket.Text Then
        MsgBox "��ʼƱ�Ų��ܴ��ڽ���Ʊ�ţ�", vbOKOnly + vbExclamation, "����"
        IsSave = False
        txtStartTicket.SetFocus
        Exit Function
    Else
        IsSave = True
    End If
End Function

'���ǰ׺
Private Sub FillFirstChar()
    If lstU.ListCount <= 0 Then Exit Sub
    GetAppSetting ResolveDisplay(lstU.Text)
    txtFirstChar.Text = m_szTicketPrefix
    txtFirstChar1.Text = m_szTicketPrefix
    

    txtFirstChar.MaxLength = m_oParam.TicketPrefixLen
    txtFirstChar1.MaxLength = m_oParam.TicketPrefixLen
    txtStartTicket.MaxLength = TicketNoNumLen()
    txtEndTicket.MaxLength = TicketNoNumLen()
End Sub

Private Sub lstU_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 40
        FillFirstChar
    Case 38
        FillFirstChar
End Select
End Sub
'���
Public Sub ClearTicketMan()
'    lstU.SetFocus
    txtStartTicket.Text = ""
    txtEndTicket.Text = ""
    txtMemo.Text = ""
    dtpGetTicketDate.Value = Date
    chkIsUseing.Value = 0
End Sub
Public Sub RefreshTicketMan(szUserID As String, dGetTicketDate As Date)
    Dim oTicketMan As New TicketMan
    Dim rsTmp As New Recordset
    Dim nTicketPrefixLen As Integer
    oTicketMan.Init m_oActiveUser
    Set rsTmp = oTicketMan.GetTicketInfo(dGetTicketDate, dGetTicketDate, szUserID)
    If rsTmp.RecordCount = 0 Then Exit Sub
    nTicketPrefixLen = m_oParam.TicketPrefixLen
    If nTicketPrefixLen = 0 Then
        txtFirstChar.Text = ""
        txtFirstChar1.Text = ""
    Else
        txtFirstChar.Text = Mid(FormatDbValue(rsTmp!firstticket_no), nTicketPrefixLen, nTicketPrefixLen)
        txtFirstChar1.Text = Mid(FormatDbValue(rsTmp!lastticket_no), nTicketPrefixLen, nTicketPrefixLen)
    End If
    txtStartTicket.Text = Mid(FormatDbValue(rsTmp!firstticket_no), nTicketPrefixLen + 1, TicketNoNumLen())
    txtEndTicket.Text = Mid(FormatDbValue(rsTmp!lastticket_no), nTicketPrefixLen + 1, TicketNoNumLen())
    dtpGetTicketDate.Value = rsTmp!getticket_date
    txtMemo.Text = rsTmp!Memo
    If rsTmp!isuseing = True Then
    chkIsUseing.Value = vbChecked
    Else
    chkIsUseing.Value = Unchecked
    End If
End Sub

