VERSION 5.00
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Begin VB.Form frmEditBusProtocol 
   BackColor       =   &H00E0E0E0&
   Caption         =   "修改车次协议"
   ClientHeight    =   2610
   ClientLeft      =   3870
   ClientTop       =   3930
   ClientWidth     =   4785
   Icon            =   "frmEditBusProtocol.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   4785
   StartUpPosition =   2  '屏幕中心
   Begin FText.asFlatTextBox txtCompany 
      Height          =   270
      Left            =   3060
      TabIndex        =   11
      Top             =   960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   476
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonHotBackColor=   -2147483633
      ButtonPressedBackColor=   -2147483627
      Text            =   ""
      ButtonBackColor =   -2147483633
      ButtonVisible   =   -1  'True
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   0
      TabIndex        =   6
      Top             =   540
      Width           =   7815
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   855
      Left            =   -30
      TabIndex        =   2
      Top             =   1920
      Width           =   5295
      Begin RTComctl3.CoolButton cmdCancel 
         Cancel          =   -1  'True
         Height          =   345
         Left            =   3570
         TabIndex        =   3
         Top             =   270
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "取消(&C)"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         MICON           =   "frmEditBusProtocol.frx":000C
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
         Height          =   345
         Left            =   2430
         TabIndex        =   4
         Top             =   270
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "确定(&E)"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         MICON           =   "frmEditBusProtocol.frx":0028
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin RTComctl3.CoolButton cmdProtocol 
         Height          =   345
         Left            =   390
         TabIndex        =   5
         Top             =   270
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "协议信息(&P)"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         MICON           =   "frmEditBusProtocol.frx":0044
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.ComboBox cboProtocol 
      Height          =   300
      Left            =   1140
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1410
      Width           =   3405
   End
   Begin VB.TextBox txtBusID 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   270
      Left            =   1140
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "车辆协议修改"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   180
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "设置协议"
      Height          =   255
      Left            =   180
      TabIndex        =   9
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车次代码"
      Height          =   180
      Left            =   150
      TabIndex        =   8
      Top             =   1005
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "公司"
      Height          =   180
      Left            =   2520
      TabIndex        =   7
      Top             =   1005
      Width           =   540
   End
End
Attribute VB_Name = "frmEditBusProtocol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public m_AllSet As Boolean '是否为批量设置
Public m_eStatus As EFormStatus
Public m_bOtherBus As Boolean '是否为计划车次中不存在的车次的新增
Public m_szBusID As String '车次代码
Public m_szCompanyID As String '该车次的公司代码


Private m_oReport As New Report
Private m_oSplit As New Split


Private Sub cboProtocol_Change()
    If cboProtocol.Text = "" Then
        cmdOk.Enabled = False
    Else
        cmdOk.Enabled = True
    End If
    
End Sub

Private Sub cboProtocol_Click()
    If cboProtocol.Text = "" Then
        cmdOk.Enabled = False
    Else
        cmdOk.Enabled = True
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
On Error GoTo here
        
        m_oSplit.SetBusProtocol txtBusID.Text, ResolveDisplay(txtCompany.Text), ResolveDisplay(cboProtocol.Text)
        m_szBusID = txtBusID.Text
        Unload Me
    
Exit Sub
here:
    ShowErrorMsg
End Sub

Private Sub cmdProtocol_Click()
    If cboProtocol.Text <> "" Then
        frmProtocolItem.m_szProtocolID = ResolveDisplay(cboProtocol.Text)
        frmProtocolItem.m_eStatus = ModifyStatus
        frmProtocolItem.Show vbModal
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo err
    Dim i As Integer
    Dim rsTmp As Recordset
    m_szBusID = ""
    
    cmdOk.Enabled = False
    m_oReport.Init g_oActiveUser
    AlignFormPos Me
    '取得所有协议
    m_oSplit.Init g_oActiveUser
    With cboProtocol
        .Clear
        .AddItem ""
    End With
    If m_AllSet = True Then '批量设置
'        lblAllEdit.Visible = True
    Else
'        lblAllEdit.Visible = False
        txtBusID.Text = Trim(frmBusProtocol.lvBusProtocol.SelectedItem.Text)
        txtCompany.Text = Trim(frmBusProtocol.lvBusProtocol.SelectedItem.SubItems(2))
    End If
    
    
    If m_bOtherBus Then '如果为新增计划车次中不存在的车次
        txtBusID.Enabled = True
        txtCompany.Enabled = True
        txtBusID.Text = ""
        txtCompany.Text = ""
    Else
        txtBusID.Enabled = False
        txtCompany.Enabled = False
        
    End If
    
    
    FillCboProtocl
    Dim aszTemp() As String
    If Not m_AllSet Then
        Set rsTmp = m_oReport.GetAllBusProtocol(txtBusID.Text, ResolveDisplay(txtCompany.Text))
        If rsTmp.RecordCount > 0 Then
            For i = 0 To cboProtocol.ListCount - 1
                If ResolveDisplay(cboProtocol.List(i)) = FormatDbValue(rsTmp!protocol_id) Then
                    cboProtocol.ListIndex = i
                    Exit For
                End If
            Next i
        End If
    End If
    If cboProtocol.ListIndex < 0 Then
        Set rsTmp = m_oReport.GetOtherBusProtocol(txtBusID.Text)
        If rsTmp.RecordCount > 0 Then
            For i = 0 To cboProtocol.ListCount - 1
                If ResolveDisplay(cboProtocol.List(i)) = FormatDbValue(rsTmp!protocol_id) Then
                    cboProtocol.ListIndex = i
                    Exit For
                End If
            Next i
        End If
    
    End If
    Exit Sub
err:
ShowErrorMsg
End Sub

Public Sub FillCboProtocl()
    On Error GoTo err
    Dim aszTemp() As String, i As Integer
    m_oReport.Init g_oActiveUser
    aszTemp = m_oReport.GetAllProtocol
    If ArrayLength(aszTemp) > 0 Then
        For i = 1 To ArrayLength(aszTemp)
            cboProtocol.AddItem MakeDisplayString(aszTemp(i, 1), aszTemp(i, 2))
        Next i
    End If
    
    
    Exit Sub
err:
ShowErrorMsg
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveFormPos Me
    Unload Me
End Sub


Private Sub txtCompany_ButtonClick()
On Error GoTo ErrHandle
    Dim oShell As New STShell.CommDialog
    Dim aszTemp() As String
    oShell.Init g_oActiveUser
    aszTemp = oShell.SelectCompany
    Set oShell = Nothing
    If ArrayLength(aszTemp) = 0 Then Exit Sub
    txtCompany.Text = MakeDisplayString(Trim(aszTemp(1, 1)), Trim(aszTemp(1, 2)))

Exit Sub
ErrHandle:
ShowErrorMsg
End Sub
