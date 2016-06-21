VERSION 5.00
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmEditVehicleProtocol 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "修改车辆协议"
   ClientHeight    =   2640
   ClientLeft      =   3855
   ClientTop       =   3915
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   4815
   StartUpPosition =   1  '所有者中心
   Begin RTComctl3.FlatLabel lblAllEdit 
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   900
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   0
      HorizontalAlignment=   1
      NormTextColor   =   16711680
      Caption         =   "请选择批量设置的协议:"
   End
   Begin VB.TextBox txtLicense 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   270
      Left            =   3390
      TabIndex        =   7
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox txtVehiclID 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   270
      Left            =   1140
      TabIndex        =   6
      Top             =   960
      Width           =   1215
   End
   Begin VB.ComboBox cboProtocol 
      Height          =   300
      Left            =   1140
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1410
      Width           =   3405
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   855
      Left            =   -30
      TabIndex        =   0
      Top             =   1920
      Width           =   5295
      Begin RTComctl3.CoolButton cmdCancel 
         Cancel          =   -1  'True
         Height          =   345
         Left            =   3570
         TabIndex        =   2
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
         MICON           =   "frmEditVehicleProtocol.frx":0000
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
         TabIndex        =   1
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
         MICON           =   "frmEditVehicleProtocol.frx":001C
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
         TabIndex        =   12
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
         MICON           =   "frmEditVehicleProtocol.frx":0038
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
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   0
      TabIndex        =   4
      Top             =   600
      Width           =   7815
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "车牌号"
      Height          =   180
      Left            =   2610
      TabIndex        =   11
      Top             =   1005
      Width           =   540
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车辆代码"
      Height          =   180
      Left            =   120
      TabIndex        =   10
      Top             =   945
      Width           =   720
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
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "车辆协议修改"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmEditVehicleProtocol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public m_AllSet As Boolean '是否为批量设置
Public m_eStatus As EFormStatus
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
    Dim i As Integer, j As Integer
    Dim aszVehicleID() As String
    Dim m_Count As Integer
    If m_AllSet = True Then
        For i = 1 To frmVehicleProtocol.lvVehicleProtocol.ListItems.Count
            If frmVehicleProtocol.lvVehicleProtocol.ListItems.Item(i).Checked = True Then
                m_Count = m_Count + 1
            End If
        Next i
        j = 1
        If m_Count > 0 Then
        ReDim aszVehicleID(1 To m_Count)
        For i = 1 To frmVehicleProtocol.lvVehicleProtocol.ListItems.Count
            If frmVehicleProtocol.lvVehicleProtocol.ListItems.Item(i).Checked = True Then
                aszVehicleID(j) = Trim(frmVehicleProtocol.lvVehicleProtocol.ListItems.Item(i).Text)
                j = j + 1
            End If
        Next i
        '把szProtocol传给接口
        
        m_oSplit.SetVehicleProtocol aszVehicleID, ResolveDisplay(cboProtocol.Text)
        For i = 1 To ArrayLength(aszVehicleID)
            frmVehicleProtocol.FillLvVehicleProtocol aszVehicleID(i)
        Next i
        End If
        Unload Me
        Exit Sub
     Else
        ReDim aszVehicleID(1 To 1)
        aszVehicleID(1) = txtVehiclID.Text
        m_oSplit.SetVehicleProtocol aszVehicleID, ResolveDisplay(cboProtocol.Text)
'        For i = 1 To ArrayLength(aszVehicleID)
'            frmVehicleProtocol.FillLvVehicleProtocol aszVehicleID(i)
'        Next i
        Unload Me
        Exit Sub
     End If
    
Exit Sub
here:
    ShowErrorMsg
End Sub

Private Sub cmdProtocol_Click()
    If cboProtocol.Text <> "" Then
'        frmProtocolItem.lblProtocol.Caption = cboProtocol.Text
        frmProtocolItem.m_szProtocolID = ResolveDisplay(cboProtocol.Text)
        frmProtocolItem.m_eStatus = ModifyStatus
        frmProtocolItem.Show vbModal
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo err
    Dim i As Integer
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
        lblAllEdit.Visible = True
    Else
        lblAllEdit.Visible = False
        txtVehiclID.Text = Trim(frmVehicleProtocol.lvVehicleProtocol.SelectedItem.Text)
        txtLicense.Text = Trim(frmVehicleProtocol.lvVehicleProtocol.SelectedItem.SubItems(1))
    End If
    FillCboProtocl
    Dim aszTemp() As String
    If Not m_AllSet Then
        aszTemp = m_oReport.GetVehicleProtocol(txtVehiclID.Text)
        If ArrayLength(aszTemp) > 0 Then
            'cboProtocol.Text = aszTemp(1, 1)
            For i = 0 To cboProtocol.ListCount - 1
                If ResolveDisplay(cboProtocol.List(i)) = aszTemp(1, 1) Then
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
    If ArrayLength(aszTemp) Then
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

