VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmLugCompanySplite 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "行包营收拆算报表"
   ClientHeight    =   3690
   ClientLeft      =   3300
   ClientTop       =   3435
   ClientWidth     =   5955
   HelpContextID   =   7000410
   Icon            =   "frmLugCompanySplite.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   5955
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   2115
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   5565
      Begin MSComctlLib.ListView lvBusInfo 
         Height          =   1605
         Left            =   3390
         TabIndex        =   11
         Top             =   330
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   2831
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin FText.asFlatTextBox txtCompany 
         Height          =   300
         Left            =   1620
         TabIndex        =   10
         Top             =   330
         Width           =   1575
         _ExtentX        =   2778
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
         Text            =   ""
         ButtonBackColor =   -2147483633
         ButtonVisible   =   -1  'True
      End
      Begin VB.ComboBox cboAcceptType 
         Height          =   300
         Left            =   1620
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   780
         Width           =   1605
      End
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   315
         Left            =   1620
         TabIndex        =   5
         Top             =   1230
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy年M月dd日"
         Format          =   61669379
         CurrentDate     =   37646
      End
      Begin MSComCtl2.DTPicker dtpEndDate 
         Height          =   315
         Left            =   1620
         TabIndex        =   14
         Top             =   1650
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy年M月dd日"
         Format          =   61669379
         CurrentDate     =   37646
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "结束月份(&D):"
         Height          =   210
         Left            =   300
         TabIndex        =   13
         Top             =   1680
         Width           =   1080
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "参运公司(&S):"
         Height          =   180
         Left            =   330
         TabIndex        =   8
         Top             =   390
         Width           =   1080
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "行包类型(&T):"
         Height          =   180
         Left            =   300
         TabIndex        =   7
         Top             =   870
         Width           =   1080
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "开始月份(&D):"
         Height          =   210
         Left            =   300
         TabIndex        =   6
         Top             =   1320
         Width           =   1080
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   " "
      Height          =   960
      Left            =   -60
      TabIndex        =   0
      Top             =   2910
      Width           =   6945
      Begin RTComctl3.CoolButton cmdOk 
         Default         =   -1  'True
         Height          =   345
         Left            =   3420
         TabIndex        =   1
         Top             =   270
         Width           =   1095
         _ExtentX        =   1931
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
         MICON           =   "frmLugCompanySplite.frx":0CCA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin RTComctl3.CoolButton cmdCancel 
         Cancel          =   -1  'True
         Height          =   345
         Left            =   4650
         TabIndex        =   2
         Top             =   270
         Width           =   1095
         _ExtentX        =   1931
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
         MICON           =   "frmLugCompanySplite.frx":0CE6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin RTComctl3.CoolButton cmdHelp 
         Height          =   315
         Left            =   180
         TabIndex        =   12
         Top             =   300
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
         MICON           =   "frmLugCompanySplite.frx":0D02
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "选择行包营收拆算条件:"
      Height          =   180
      Left            =   540
      TabIndex        =   9
      Top             =   210
      Width           =   1890
   End
End
Attribute VB_Name = "frmLugCompanySplite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public m_bOk As Boolean
Public m_dtStartDate As Date
Public m_dtEndDate As Date
Public m_szAcceptType As String
Public m_SellStation As String
Public m_Company As String
Private Sub cmdCancel_Click()
 Unload Me
End Sub

Private Sub cmdHelp_Click()
    DisplayHelp Me
End Sub

Private Sub cmdOk_Click()
    Dim i As Integer
    Dim nSplitVehicleCount As Integer
    If lvBusInfo.ListItems.Count = 0 Then Exit Sub
    nSplitVehicleCount = 0
    For i = 1 To lvBusInfo.ListItems.Count
      If lvBusInfo.ListItems(i).Checked = True Then
         nSplitVehicleCount = nSplitVehicleCount + 1
      End If
    Next i
    If nSplitVehicleCount = 0 Then
       MsgBox "请选择要统计的拆帐车辆!", vbInformation, Me.Caption
       Exit Sub
    End If
    ReDim mSplitVehicleID(1 To nSplitVehicleCount)
    For i = 1 To lvBusInfo.ListItems.Count
      If lvBusInfo.ListItems(i).Checked = True Then
        mSplitVehicleID(i) = Trim(lvBusInfo.ListItems(i).Tag)
      End If
    Next i
    m_dtStartDate = dtpStartDate.Value
    m_dtEndDate = dtpEndDate.Value
    m_szAcceptType = cboAcceptType.Text
    m_Company = ResolveDisplayEx(txtCompany.Text)
    m_bOk = True
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And (Me.Controls Is txtCompany) Then
   lvBusInfo.SetFocus
End If
End Sub

Private Sub Form_Load()
  AlignFormPos Me
   txtCompany.Text = ""
   m_bOk = False
   With cboAcceptType  '行包类型
      .clear
      .AddItem ""
      .AddItem szAcceptTypeGeneral
      .AddItem szAcceptTypeMan
      .ListIndex = 0
   End With
     
   dtpStartDate.Value = Now
   dtpEndDate.Value = Now
   '填充列首
   With lvBusInfo.ColumnHeaders
      .clear
      .Add , , "    车牌号", 1700
   End With
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
 SaveFormPos Me
 Unload Me
End Sub

Private Sub lvBusInfo_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
 If lvBusInfo.SortOrder = lvwAscending Then
    lvBusInfo.SortOrder = lvwDescending
 Else
    lvBusInfo.SortOrder = lvwAscending
 End If
    lvBusInfo.SortKey = ColumnHeader.Index - 1
    lvBusInfo.Sorted = True
End Sub

Private Sub lvBusInfo_DblClick()
On Error GoTo ErrHandle
 Dim i As Integer
 If lvBusInfo.ListItems.Count = 0 Then Exit Sub
 If lvBusInfo.SelectedItem.Checked = False Then
    For i = 1 To lvBusInfo.ListItems.Count
       lvBusInfo.ListItems(i).Checked = True
    Next i
 Else
    For i = 1 To lvBusInfo.ListItems.Count
       lvBusInfo.ListItems(i).Checked = False
    Next i
 End If
 Exit Sub
ErrHandle:
ShowErrorMsg
End Sub

Private Sub txtCompany_ButtonClick()
On Error GoTo ErrHandle
    Dim oShell As New STShell.CommDialog
    Dim aszTemp() As String
    oShell.Init m_oAUser
    aszTemp = oShell.SelectCompany()
    Set oShell = Nothing
    If ArrayLength(aszTemp) = 0 Then Exit Sub
    txtCompany.Text = aszTemp(1, 1) & "[" & Trim(aszTemp(1, 2)) & "]"
    If txtCompany.Text <> "" Then
        FilllvBusInfo
    End If
Exit Sub
ErrHandle:
ShowErrorMsg
End Sub
' 填充车辆
Private Sub FilllvBusInfo()
  On Error GoTo ErrHandle
     Dim i As Integer
     Dim rsTemp As Recordset
      If txtCompany.Text <> "" Then
      Set rsTemp = m_oLugFinSvr.GetVehicleInfo(ResolveDisplay(Trim(txtCompany.Text)))
      If rsTemp.RecordCount = 0 Then Exit Sub
       lvBusInfo.ListItems.clear
       For i = 1 To rsTemp.RecordCount
          lvBusInfo.ListItems.Add , , FormatDbValue(rsTemp!license_tag_no)
          lvBusInfo.ListItems(i).Tag = FormatDbValue(rsTemp!vehicle_id)
         rsTemp.MoveNext
       Next i
      End If
  Exit Sub
ErrHandle:
  ShowErrorMsg
End Sub
Private Sub txtCompany_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
      lvBusInfo.SetFocus
   End If
End Sub

Private Sub txtCompany_LostFocus()
 If txtCompany.Text <> "" Then
  FilllvBusInfo
 End If
End Sub
