VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmSetVehicleProtocol 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "设定车次拆算公式"
   ClientHeight    =   5265
   ClientLeft      =   4800
   ClientTop       =   1635
   ClientWidth     =   7620
   HelpContextID   =   7000240
   Icon            =   "frmSetVehicleProtocol.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   7620
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboAcceptType 
      Height          =   300
      ItemData        =   "frmSetVehicleProtocol.frx":0442
      Left            =   1230
      List            =   "frmSetVehicleProtocol.frx":0444
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   4140
      Width           =   1185
   End
   Begin VB.TextBox txtLicense 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   4245
      TabIndex        =   10
      Top             =   180
      Width           =   1620
   End
   Begin VB.TextBox txtVehicle 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1305
      TabIndex        =   9
      Top             =   180
      Width           =   1590
   End
   Begin RTComctl3.CoolButton CmdOk 
      Height          =   345
      Left            =   4980
      TabIndex        =   1
      Top             =   4830
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   609
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
      MICON           =   "frmSetVehicleProtocol.frx":0446
      PICN            =   "frmSetVehicleProtocol.frx":0462
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton CmdCancel 
      Cancel          =   -1  'True
      Height          =   345
      Left            =   6240
      TabIndex        =   0
      Top             =   4830
      Width           =   1155
      _ExtentX        =   0
      _ExtentY        =   0
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
      MICON           =   "frmSetVehicleProtocol.frx":07FC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin FText.asFlatTextBox txtProtocol 
      Height          =   300
      Left            =   3420
      TabIndex        =   5
      Top             =   4140
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   529
      Enabled         =   0   'False
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
   Begin RTComctl3.CoolButton cmdFind 
      Default         =   -1  'True
      Height          =   315
      Left            =   6120
      TabIndex        =   6
      Top             =   240
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "查询(&Q)"
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
      MICON           =   "frmSetVehicleProtocol.frx":0818
      PICN            =   "frmSetVehicleProtocol.frx":0834
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin FText.asFlatTextBox txtSplitCompany 
      Height          =   300
      Left            =   4260
      TabIndex        =   7
      Top             =   600
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   529
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
      Text            =   "(全部)"
      ButtonBackColor =   -2147483633
      ButtonVisible   =   -1  'True
   End
   Begin FText.asFlatTextBox txtCompany 
      Height          =   300
      Left            =   1320
      TabIndex        =   8
      Top             =   600
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   529
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
      Text            =   "(全部)"
      ButtonBackColor =   -2147483633
      ButtonVisible   =   -1  'True
   End
   Begin MSComctlLib.ListView lvVehicle 
      Height          =   2670
      Left            =   240
      TabIndex        =   11
      Top             =   1290
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   4710
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "smallImgLists"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "车辆代码"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "车辆车牌"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "参运公司"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "拆帐公司"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "车主"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "车型"
         Object.Width           =   1764
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbSelect 
      Height          =   360
      Left            =   6570
      TabIndex        =   17
      Top             =   960
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "smallImgLists"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SelectAll"
            Object.ToolTipText     =   "全部选择"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "CancelSelect"
            Object.ToolTipText     =   "取消选择"
            ImageIndex      =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   930
      Left            =   -30
      TabIndex        =   18
      Top             =   4530
      Width           =   8745
      Begin RTComctl3.CoolButton cmdHelp 
         Height          =   315
         Left            =   270
         TabIndex        =   21
         Top             =   330
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "帮助(&H)"
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
         MICON           =   "frmSetVehicleProtocol.frx":0BCE
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
   Begin MSComctlLib.ImageList smallImgLists 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetVehicleProtocol.frx":0BEA
            Key             =   "vehicle"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetVehicleProtocol.frx":0F84
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetVehicleProtocol.frx":1561
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "托运方式(&T):"
      Height          =   180
      Left            =   180
      TabIndex        =   20
      Top             =   4200
      Width           =   1080
   End
   Begin VB.Label lblProtocolName 
      BackStyle       =   0  'Transparent
      Caption         =   "5%拆算协议"
      Height          =   180
      Left            =   5880
      TabIndex        =   4
      Top             =   4200
      Width           =   1560
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "协议名称(&N):"
      Height          =   180
      Index           =   0
      Left            =   4800
      TabIndex        =   3
      Top             =   4200
      Width           =   1080
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "协议号(&P):"
      Height          =   180
      Index           =   0
      Left            =   2460
      TabIndex        =   2
      Top             =   4200
      Width           =   900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车牌(&P):"
      Height          =   180
      Left            =   3180
      TabIndex        =   16
      Top             =   240
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "代码(&N):"
      Height          =   180
      Left            =   210
      TabIndex        =   15
      Top             =   270
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "参运公司(&Z):"
      Height          =   180
      Index           =   1
      Left            =   180
      TabIndex        =   14
      Top             =   660
      Width           =   1080
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   1290
      X2              =   6930
      Y1              =   1140
      Y2              =   1140
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "拆帐公司(&T):"
      Height          =   180
      Index           =   1
      Left            =   3150
      TabIndex        =   13
      Top             =   690
      Width           =   1080
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车辆列表(&L):"
      Height          =   180
      Left            =   180
      TabIndex        =   12
      Top             =   1035
      Width           =   1080
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   1290
      X2              =   6930
      Y1              =   1455
      Y2              =   1455
   End
End
Attribute VB_Name = "frmSetVehicleProtocol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboAcceptType_Click()
'    cmdFind_Click
    SelectItem
End Sub

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdFind_Click()
 On Error GoTo ErrHandle:
  Dim szTemp() As String
  Dim i As Integer
  Dim nlen As Integer
  Dim lvItem As ListItem
  '填充lvVehicle列表
  lvVehicle.ListItems.clear
  m_obase.Init m_oAUser
    If txtCompany.Text = "(全部)" Then
        txtCompany.Text = ""
     End If
    If txtSplitCompany.Text = "(全部)" Then
        txtSplitCompany.Text = ""
    End If
  nlen = ArrayLength(m_obase.GetVehicle(Trim(txtVehicle.Text), ResolveDisplay(txtCompany.Text), , , Trim(txtLicense.Text)))
  If nlen > 0 Then
     ReDim szTemp(1 To nlen, 1 To 10)
 
     szTemp = m_obase.GetVehicle(Trim(txtVehicle.Text), ResolveDisplay(txtCompany.Text), , , Trim(txtLicense.Text))
     For i = 1 To nlen
      Set lvItem = lvVehicle.ListItems.Add(, , szTemp(i, 1))
          lvItem.SmallIcon = "vehicle"
          lvItem.SubItems(1) = szTemp(i, 2)
          lvItem.SubItems(2) = szTemp(i, 4)
          lvItem.SubItems(3) = szTemp(i, 10)
          lvItem.SubItems(4) = szTemp(i, 5)
          lvItem.SubItems(5) = szTemp(i, 8)
          'lvitem.SubItems(5) = szTemp(i, 8) '拆算协议
     Next i
    
     
   
                     
  Else
     MsgBox "没有找到所指定的车辆信息", vbInformation, Me.Caption
     Exit Sub
  End If
  
'  cmdOk.Enabled = False
Exit Sub
ErrHandle:
 ShowErrorMsg
End Sub

Private Sub cmdHelp_Click()
    DisplayHelp Me
End Sub

Private Sub cmdOk_Click()
On Error GoTo ErrHandle
  Dim i As Integer, j As Integer
  Dim szVehicleProtocol() As TVehicleProtocol

  Dim mAnswer
  Dim nlen As Integer
  Dim mNum As Integer
  Dim k As Integer
  
  If cboAcceptType.Text = "" Then
      MsgBox "请选择托运类型!", vbInformation, Me.Caption
      Exit Sub
  End If
  k = 1
  mNum = 0
  j = 1
  If lvVehicle.ListItems.Count = 0 Then Exit Sub
  For i = 1 To lvVehicle.ListItems.Count
      If lvVehicle.ListItems.Item(i).Checked = True Then
          mNum = mNum + 1
      End If
  Next i
  
  If mNum = 0 Then Exit Sub
  ReDim szVehicleProtocol(1 To mNum)
  mAnswer = MsgBox("确认是否真的进行设置", vbInformation + vbYesNo, Me.Caption)
  If mAnswer = vbYes Then

     For i = 1 To lvVehicle.ListItems.Count
         
         If lvVehicle.ListItems.Item(i).Checked = True Then
       
            szVehicleProtocol(k).VehicleID = Trim(lvVehicle.ListItems.Item(i).Text)
            szVehicleProtocol(k).VehicleLicense = Trim(lvVehicle.ListItems.Item(i).SubItems(1))
            szVehicleProtocol(k).ProtocolID = Trim(txtProtocol.Text)
            szVehicleProtocol(k).ProtocolName = Trim(lblProtocolName.Caption)
            Select Case cboAcceptType.Text
                   Case szAcceptTypeGeneral
                     szVehicleProtocol(k).AcceptType = 0
                   Case szAcceptTypeMan
                    szVehicleProtocol(k).AcceptType = 1
            End Select
          k = k + 1
         End If
     Next i
   
     SetBusy
     m_oProtocol.SetVehicleProtocol szVehicleProtocol
     SetNormal
  End If
  Unload Me
Exit Sub
ErrHandle:
 ShowErrorMsg
End Sub

Private Sub Form_Load()
  Dim szTemp() As TLugProtocol
  Dim nlen As Integer
   AlignFormPos Me
   '填充托运方式
   With cboAcceptType
     .AddItem szAcceptTypeGeneral
     .AddItem szAcceptTypeMan
   End With
'   cboAcceptType.ListIndex = 0
   '取得默认协议信息
   m_oProtocol.Identify frmProtocol.lbProtocol.Caption
   szTemp = m_oProtocol.GetProtocol()
   nlen = ArrayLength(szTemp)
   If nlen > 0 Then
     ReDim szTemp(1 To 1)
     txtProtocol.Text = szTemp(1).ProtocolID
     lblProtocolName.Caption = szTemp(1).ProtocolName
   End If
   '显示设置的车辆列表
'    cmdFind_Click
'    SelectItem
'
'   txtProtocol.Enabled = False
'   txtVehicle.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
   SaveFormPos Me
End Sub

Private Sub lvVehicle_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
 If lvVehicle.SortOrder = lvwAscending Then
  lvVehicle.SortOrder = lvwDescending
 Else
  lvVehicle.SortOrder = lvwAscending
 End If
  lvVehicle.SortKey = ColumnHeader.Index - 1
  lvVehicle.Sorted = True
End Sub

Private Sub tbSelect_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim i As Integer
 If lvVehicle.ListItems.Count > 0 Then
    Select Case Button.Key
           Case "SelectAll"
            For i = 1 To lvVehicle.ListItems.Count
                lvVehicle.ListItems.Item(i).Checked = True
            Next i
           Case "CancelSelect"
            For i = 1 To lvVehicle.ListItems.Count
                lvVehicle.ListItems.Item(i).Checked = False
            Next i
    End Select
End If
End Sub

Private Sub txtCompany_ButtonClick()
On Error GoTo ErrHandle
    Dim oShell As New STShell.CommDialog
    Dim aszTemp() As String
    oShell.Init m_oAUser
    aszTemp = oShell.SelectCompany()
    Set oShell = Nothing
    If ArrayLength(aszTemp) = 0 Then Exit Sub
     txtCompany.Text = MakeDisplayString(aszTemp(1, 1), aszTemp(1, 2))
    
    
Exit Sub
ErrHandle:
ShowErrorMsg
End Sub

Private Sub txtProtocol_ButtonClick()
    Dim oShell As New STShell.CommDialog
    Dim aszTemp() As String
    oShell.Init m_oAUser
    aszTemp = oShell.SelectProtocol()
    Set oShell = Nothing
    If ArrayLength(aszTemp) = 0 Then Exit Sub
    txtSplitCompany.Text = MakeDisplayString(aszTemp(1, 1), aszTemp(1, 2))
End Sub

'Private Sub txtProtocol_ButtonClick()
'  On Error GoTo ErrHandle
'    Dim oShell As New STShell.CommDialog
'    Dim aszTemp() As String
'    oShell.Init m_oAUser
'    aszTemp = oShell.SelectProtocol()
'    Set oShell = Nothing
'    If ArrayLength(aszTemp) = 0 Then Exit Sub
'
'     txtSplitCompany.Text = Trim(aszTemp(1, 1))
'     lblProtocolName.Caption = Trim(aszTemp(1, 2))
'
'
'Exit Sub
'ErrHandle:
'ShowErrorMsg
'End Sub

Private Sub txtSplitCompany_ButtonClick()
On Error GoTo ErrHandle
    Dim oShell As New STShell.CommDialog
    Dim aszTemp() As String
    oShell.Init m_oAUser
    aszTemp = oShell.SelectCompany()
    Set oShell = Nothing
    If ArrayLength(aszTemp) = 0 Then Exit Sub
     txtSplitCompany.Text = MakeDisplayString(aszTemp(1, 1), aszTemp(1, 2))

    
Exit Sub
ErrHandle:
ShowErrorMsg
End Sub


'如果该协议已经存在参运车辆， 就相应的把这些参运车辆选中

Public Sub SelectItem()
    Dim i As Integer, j As Integer
      Dim szVehichleProtocol() As TVehicleProtocol
      szVehichleProtocol = m_oProtocol.GetVehicleProtocol(, (IIf((cboAcceptType.Text = szAcceptTypeGeneral), 0, 1)), frmProtocol.lbProtocol)
'       ReDim szVehichleProtocol(1 To ArrayLength(szVehichleProtocol))
       If lvVehicle.ListItems.Count > 0 Then
           For j = 1 To lvVehicle.ListItems.Count
                For i = 1 To ArrayLength(szVehichleProtocol)
                    If Trim(lvVehicle.ListItems(j).Text) = szVehichleProtocol(i).VehicleID Then lvVehicle.ListItems(j).Checked = True
                Next i
           Next j
    End If
End Sub
