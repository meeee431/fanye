VERSION 5.00
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmEditProtocol 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "修改车辆协议"
   ClientHeight    =   3150
   ClientLeft      =   4350
   ClientTop       =   1515
   ClientWidth     =   5130
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   5130
   Begin VB.TextBox txtLicense 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3555
      TabIndex        =   8
      Top             =   1035
      Width           =   1215
   End
   Begin VB.TextBox txtAcceptType 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1410
      TabIndex        =   7
      Top             =   1515
      Width           =   1215
   End
   Begin VB.TextBox txtVehiclID 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1410
      TabIndex        =   6
      Top             =   1065
      Width           =   1215
   End
   Begin VB.ComboBox cboProtocol 
      Height          =   300
      Left            =   1410
      TabIndex        =   5
      Top             =   1965
      Width           =   3375
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   -120
      TabIndex        =   0
      Top             =   2505
      Width           =   5295
      Begin RTComctl3.CoolButton CoolButton2 
         Height          =   350
         Left            =   3720
         TabIndex        =   2
         Top             =   240
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
         MICON           =   "frmEditProtocol.frx":0000
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
         Height          =   350
         Left            =   2520
         TabIndex        =   1
         Top             =   240
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
         MICON           =   "frmEditProtocol.frx":001C
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
      Top             =   810
      Width           =   7815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "车牌号"
      Height          =   255
      Left            =   2775
      TabIndex        =   12
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "修改协议"
      Height          =   255
      Left            =   315
      TabIndex        =   11
      Top             =   1988
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "托运方式"
      Height          =   255
      Left            =   315
      TabIndex        =   10
      Top             =   1530
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "车辆代码"
      Height          =   255
      Left            =   315
      TabIndex        =   9
      Top             =   1080
      Width           =   1215
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
Attribute VB_Name = "frmEditProtocol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CoolButton1_Click()
    Dim szVehicleProtocol(1 To 1) As TVehicleProtocol
    szVehicleProtocol(1).AcceptType = GetLuggageTypeInt(txtAcceptType.Text)
    szVehicleProtocol(1).ProtocolID = ResolveDisplay(cboProtocol.Text)
    szVehicleProtocol(1).ProtocolName = ""
    szVehicleProtocol(1).VehicleID = txtVehiclID.Text
    szVehicleProtocol(1).VehicleLicense = txtLicense.Text
    m_oProtocol.SetVehicleProtocol szVehicleProtocol
    frmBaseInfo.lvObject.ListItems(frmBaseInfo.lvObject.SelectedItem.Index).ListSubItems(1).Text = cboProtocol.Text
    
    Unload Me
End Sub

Private Sub CoolButton2_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Dim atTemp() As TLugProtocol
'    Dim aaa(1 To 3) As String
'    Dim ddd() As String
'    aaa(1) = "0000001"
'    aaa(2) = "0000002"
'
'    bbb = m_oFinanceSheet.PreviewSplitCarrySheets(aaa)
AlignFormPos Me
    atTemp = m_oLugFinSvr.GetAllProtocol
    If ArrayLength(atTemp) <> 0 Then
        For i = 1 To ArrayLength(atTemp)
            cboProtocol.AddItem (MakeDisplayString(atTemp(i).ProtocolID, atTemp(i).ProtocolName))
        Next i
        cboProtocol.ListIndex = 0
    End If
End Sub

Private Sub txtProtocolID_ButtonClick()
    Dim oShell As New STShell.CommDialog
    Dim aszTemp() As String
    oShell.Init m_oAUser
    aszTemp = oShell.SelectProtocol()
    Set oShell = Nothing
    If ArrayLength(aszTemp) = 0 Then Exit Sub
    txtSplitCompany.Text = MakeDisplayString(aszTemp(1, 1), aszTemp(1, 2))
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveFormPos Me
End Sub

