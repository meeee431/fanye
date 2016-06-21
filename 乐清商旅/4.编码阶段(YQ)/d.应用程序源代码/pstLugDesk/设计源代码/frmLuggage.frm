VERSION 5.00
Object = "{695ABF14-B2D8-11D2-A5ED-DE08DCF33612}#1.2#0"; "asfcombo.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Begin VB.Form frmLuggage 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "行包"
   ClientHeight    =   3990
   ClientLeft      =   4650
   ClientTop       =   3090
   ClientWidth     =   5835
   HelpContextID   =   2002001
   Icon            =   "frmLuggage.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   5835
   Begin FCmbo.asFlatCombo cboPackType 
      Height          =   270
      Left            =   1680
      TabIndex        =   8
      Top             =   2760
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   476
      ButtonDisabledForeColor=   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonHotBackColor=   8421504
      ButtonPressedBackColor=   0
      Text            =   ""
      ButtonBackColor =   8421504
      Style           =   1
      Registered      =   -1  'True
      OfficeXPColors  =   -1  'True
   End
   Begin VB.TextBox txtBulk 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   4110
      MaxLength       =   5
      TabIndex        =   7
      Top             =   2370
      Width           =   1110
   End
   Begin VB.TextBox txtActWeight 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1680
      MaxLength       =   5
      TabIndex        =   6
      Top             =   2370
      Width           =   1110
   End
   Begin VB.TextBox txtCalWeight 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   4110
      MaxLength       =   5
      TabIndex        =   5
      Top             =   1980
      Width           =   1110
   End
   Begin VB.TextBox txtNumber 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1680
      MaxLength       =   5
      TabIndex        =   4
      Top             =   1980
      Width           =   1110
   End
   Begin VB.TextBox txtLuggageID 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1680
      MaxLength       =   5
      TabIndex        =   1
      Top             =   840
      Width           =   1110
   End
   Begin VB.PictureBox ptTitle 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   -30
      ScaleHeight     =   705
      ScaleWidth      =   6615
      TabIndex        =   15
      Top             =   0
      Width           =   6615
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   60
         Left            =   -30
         TabIndex        =   16
         Top             =   660
         Width           =   6885
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请修改或新增行包信息:"
         Height          =   180
         Left            =   120
         TabIndex        =   17
         Top             =   270
         Width           =   1890
      End
      Begin VB.Image Image1 
         Height          =   855
         Left            =   390
         Picture         =   "frmLuggage.frx":038A
         Top             =   0
         Width           =   5925
      End
   End
   Begin RTComctl3.CoolButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   4290
      TabIndex        =   10
      Top             =   3480
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "关闭(&C)"
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
      MICON           =   "frmLuggage.frx":1874
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1680
      TabIndex        =   2
      Top             =   1215
      Width           =   3540
   End
   Begin VB.TextBox txtAcceptID 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   4110
      TabIndex        =   12
      Top             =   840
      Width           =   1110
   End
   Begin FCmbo.asFlatCombo cboType 
      Height          =   270
      Left            =   1680
      TabIndex        =   3
      Top             =   1590
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   476
      ButtonDisabledForeColor=   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonHotBackColor=   8421504
      ButtonPressedBackColor=   0
      Text            =   ""
      ButtonBackColor =   8421504
      Style           =   1
      Registered      =   -1  'True
      OfficeXPColors  =   -1  'True
   End
   Begin RTComctl3.CoolButton cmdOk 
      Height          =   315
      Left            =   3060
      TabIndex        =   9
      Top             =   3480
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "保存(&S)"
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
      MICON           =   "frmLuggage.frx":1890
      PICN            =   "frmLuggage.frx":18AC
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
      Caption         =   " RTStation"
      Enabled         =   0   'False
      Height          =   780
      Left            =   -150
      TabIndex        =   18
      Top             =   3240
      Width           =   8745
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "包装形式(&D):"
      Height          =   180
      Left            =   555
      TabIndex        =   23
      Top             =   2820
      Width           =   1080
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "实重(&M):"
      Height          =   180
      Left            =   555
      TabIndex        =   22
      Top             =   2445
      Width           =   720
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "体积(&L):"
      Height          =   180
      Left            =   2985
      TabIndex        =   21
      Top             =   2445
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "件数(&M):"
      Height          =   180
      Left            =   555
      TabIndex        =   20
      Top             =   2055
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "计重(&L):"
      Height          =   180
      Left            =   3000
      TabIndex        =   19
      Top             =   2040
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "行包名称(&D):"
      Height          =   180
      Left            =   555
      TabIndex        =   13
      Top             =   1275
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "行包单号(&N):"
      Height          =   180
      Left            =   2985
      TabIndex        =   11
      Top             =   915
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "标签号(&I):"
      Height          =   180
      Left            =   555
      TabIndex        =   0
      Top             =   915
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "行包类型(&C):"
      Height          =   180
      Left            =   555
      TabIndex        =   14
      Top             =   1635
      Width           =   1080
   End
End
Attribute VB_Name = "frmLuggage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status As EFormStatus
Public LuggageItemID As String



Private Sub cboPackType_Change()
cmdOK.Enabled = True
End Sub

Private Sub cboType_Change()
cmdOK.Enabled = True
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
Dim sLugItem(1 To 1) As TLuggageItemInfo
On Error GoTo ErrHandle
    Select Case Status
        Case EFormStatus.EFS_AddNew
            
            sLugItem(1).LabelID = Trim(txtLuggageID.Text)
            sLugItem(1).LuggageID = Trim(txtAcceptID)
            sLugItem(1).LuggageName = Trim(txtName.Text)
            sLugItem(1).LuggageTypeName = ResolveDisplayEx(cboType.Text)
            sLugItem(1).LuggageType = ResolveDisplay(cboType.Text)
            sLugItem(1).Number = CInt(txtNumber.Text)
            sLugItem(1).CalWeight = CDate(txtCalWeight.Text)
            sLugItem(1).ActWeight = CDate(txtActWeight.Text)
            sLugItem(1).luggage_bulk = CInt(txtBulk.Text)
            sLugItem(1).PackType = Trim(cboPackType.Text)
            moAcceptSheet.AddLugItem sLugItem
      Case EFormStatus.EFS_Modify
            sLugItem(1).LuggageID = Trim(txtAcceptID.Text)
            sLugItem(1).LabelID = txtLuggageID
            sLugItem(1).LuggageName = txtName.Text
            sLugItem(1).LuggageTypeName = ResolveDisplayEx(cboType.Text)
            sLugItem(1).LuggageType = ResolveDisplay(cboType.Text)
            sLugItem(1).Number = CInt(txtNumber.Text)
            sLugItem(1).CalWeight = CDate(txtCalWeight.Text)
            sLugItem(1).ActWeight = CDate(txtActWeight.Text)
            sLugItem(1).luggage_bulk = CDate(txtBulk.Text)
            sLugItem(1).PackType = Trim(cboPackType.Text)
            moAcceptSheet.UpdateLugItem sLugItem
    End Select

    Unload Me
    Exit Sub
ErrHandle:
    Status = EFS_AddNew
    ShowErrorMsg

End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            SendKeys "{TAB}"
        Case vbKeyEscape
            Unload Me
    End Select
End Sub


Private Sub Form_Load()
  Dim rsTemp As Recordset
  Dim szaTemp() As String
  Dim nlen As Integer
  Dim i As Integer
    On Error GoTo ErrHandle
    '布置窗体
    AlignFormPos Me
    
    '行包类型填充
    Set rsTemp = moSysParam.GetLuggageKinds
    If rsTemp.RecordCount > 0 Then
     For i = 1 To rsTemp.RecordCount
       cboType.AddItem MakeDisplayString(Trim(rsTemp!kinds_code), Trim(rsTemp!kinds_name))
       rsTemp.MoveNext
     Next i
    End If
    cboType.ListIndex = 0
    Set rsTemp = Nothing
    '包装形式从接口读
      nlen = ArrayLength(moSysParam.GetPackageType)
     ReDim szaTemp(1 To nlen)
     szaTemp = moSysParam.GetPackageType
    If nlen > 0 Then
       For i = 1 To nlen
           cboPackType.AddItem szaTemp(i)
       Next i
    End If
    cboPackType.ListIndex = 0
    Select Case Status
        Case EFormStatus.EFS_AddNew
           cmdOK.Caption = "新增(&A)"
           txtAcceptID.Enabled = False
            RefreshLuggage
        Case EFormStatus.EFS_Modify
           txtAcceptID.Enabled = False
           cmdOK.Enabled = True
           RefreshLuggage
        Case EFormStatus.EFS_Show
           txtAcceptID.Enabled = False
           cmdOK.Enabled = False
           RefreshLuggage
    End Select
'    cmdOk.Enabled = False
    
    Exit Sub
ErrHandle:
    Status = EFS_AddNew
    ShowErrorMsg
End Sub
Private Sub RefreshLuggage()
Dim sLugItem() As TLuggageItemInfo
On Error GoTo ErrHandle
    If Status = EFS_AddNew Then
         txtAcceptID.Text = LuggageItemID
         txtLuggageID.Text = ""
         txtName.Text = ""
         txtNumber.Text = ""
         txtCalWeight.Text = ""
         txtActWeight.Text = ""
         txtBulk.Text = ""
     Else
         ReDim sLugItem(1 To 1)
         sLugItem = moAcceptSheet.GetLugItemDetail
         txtAcceptID.Text = LuggageItemID
         txtLuggageID.Text = sLugItem(1).LabelID
         txtName.Text = sLugItem(1).LuggageName
         txtNumber.Text = sLugItem(1).Number
         txtCalWeight.Text = sLugItem(1).CalWeight
         txtActWeight.Text = sLugItem(1).ActWeight
         txtBulk.Text = sLugItem(1).luggage_bulk
         cboType.Text = MakeDisplayString(sLugItem(1).LuggageType, sLugItem(1).LuggageTypeName)
         cboPackType.Text = sLugItem(1).PackType
    End If
Exit Sub
ErrHandle:
 ShowErrorMsg
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveFormPos Me
    
End Sub

Private Sub txtActWeight_Change()
cmdOK.Enabled = True
End Sub

Private Sub txtActWeight_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = vbKeyBack Or KeyAscii = 46) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtBulk_Change()
cmdOK.Enabled = True
End Sub

Private Sub txtBulk_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = vbKeyBack Or KeyAscii = 46) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtCalWeight_Change()
cmdOK.Enabled = True
End Sub

Private Sub txtCalWeight_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = vbKeyBack Or KeyAscii = 46) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtLuggageID_Change()
cmdOK.Enabled = True
End Sub

Private Sub txtName_Change()
cmdOK.Enabled = True
End Sub

Private Sub txtNumber_Change()
cmdOK.Enabled = True
End Sub

Private Sub txtNumber_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = vbKeyBack Or KeyAscii = 46) Then
        KeyAscii = 0
    End If
End Sub
