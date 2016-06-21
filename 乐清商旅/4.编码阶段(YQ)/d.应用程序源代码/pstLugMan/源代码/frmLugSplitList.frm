VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmLugSplitList 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "行包营收拆算一览表"
   ClientHeight    =   3435
   ClientLeft      =   3510
   ClientTop       =   4755
   ClientWidth     =   5850
   HelpContextID   =   7000380
   Icon            =   "frmLugSplitList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   5850
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1995
      Left            =   90
      TabIndex        =   3
      Top             =   630
      Width           =   5685
      Begin MSComctlLib.ListView lvSplitCompany 
         Height          =   1575
         Left            =   3390
         TabIndex        =   11
         Top             =   240
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   2778
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.ComboBox cboSellStation 
         Height          =   300
         ItemData        =   "frmLugSplitList.frx":0CCA
         Left            =   1590
         List            =   "frmLugSplitList.frx":0CCC
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   240
         Width           =   1605
      End
      Begin VB.ComboBox cboAcceptType 
         Height          =   300
         Left            =   1590
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   660
         Width           =   1605
      End
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   315
         Left            =   1590
         TabIndex        =   7
         Top             =   1110
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
         Left            =   1590
         TabIndex        =   14
         Top             =   1530
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
         Caption         =   "结束月份(&E):"
         Height          =   210
         Left            =   270
         TabIndex        =   13
         Top             =   1560
         Width           =   1080
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "开始月份(&D):"
         Height          =   210
         Left            =   270
         TabIndex        =   10
         Top             =   1200
         Width           =   1080
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "行包类型(&T):"
         Height          =   180
         Left            =   270
         TabIndex        =   9
         Top             =   750
         Width           =   1080
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "上车站(&S):"
         Height          =   180
         Left            =   270
         TabIndex        =   8
         Top             =   330
         Width           =   900
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   " "
      Height          =   960
      Left            =   -30
      TabIndex        =   0
      Top             =   2700
      Width           =   6945
      Begin RTComctl3.CoolButton cmdOk 
         Default         =   -1  'True
         Height          =   345
         Left            =   3480
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
         MICON           =   "frmLugSplitList.frx":0CCE
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
         Left            =   4710
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
         MICON           =   "frmLugSplitList.frx":0CEA
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
         Left            =   120
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
         MICON           =   "frmLugSplitList.frx":0D06
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
      Left            =   300
      TabIndex        =   4
      Top             =   240
      Width           =   1890
   End
End
Attribute VB_Name = "frmLugSplitList"
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


Private Sub cmdCancel_Click()
 Unload Me
End Sub

Private Sub cmdHelp_Click()
    DisplayHelp Me
End Sub

Private Sub cmdOk_Click()
    Dim i As Integer
    Dim nSplitCompanyCount As Integer
    Dim j As Integer
    If lvSplitCompany.ListItems.Count = 0 Then Exit Sub
    nSplitCompanyCount = 0
    For i = 1 To lvSplitCompany.ListItems.Count
        If lvSplitCompany.ListItems(i).Checked = True Then
            nSplitCompanyCount = nSplitCompanyCount + 1
        End If
    Next i
    If nSplitCompanyCount = 0 Then
        MsgBox "请选择要统计的拆帐公司!", vbInformation, Me.Caption
        Exit Sub
    End If
    ReDim mSplitCompanyID(1 To nSplitCompanyCount)
    j = 0
    For i = 1 To lvSplitCompany.ListItems.Count
        If lvSplitCompany.ListItems(i).Checked = True Then
            j = j + 1
            mSplitCompanyID(j) = Trim(lvSplitCompany.ListItems(i).Tag)
        End If
    Next i
    m_dtStartDate = dtpStartDate.Value
    m_dtEndDate = dtpEndDate.Value
    m_szAcceptType = cboAcceptType.Text
    m_SellStation = cboSellStation.Text
    m_bOk = True
    Unload Me
End Sub

Private Sub Form_Load()
 
  AlignFormPos Me
   m_bOk = False
   FillSellStation cboSellStation '上车站
      
   With cboAcceptType  '行包类型
      .clear
      .AddItem ""
      .AddItem szAcceptTypeGeneral
      .AddItem szAcceptTypeMan
      .ListIndex = 0
   End With
     
   '填充拆帐公司
   FillSplitCompany

    
  dtpStartDate.Value = Now
  dtpEndDate.Value = Now
End Sub

   '填充拆帐公司
Private Sub FillSplitCompany()
 On Error GoTo ErrHandle
    Dim i As Integer
    Dim rsTemp As Recordset
    
   With lvSplitCompany.ColumnHeaders
      .clear
      .Add , , "   拆帐公司", 1700
   End With
   Set rsTemp = m_oLugFinSvr.GetSplitCompany()
   lvSplitCompany.ListItems.clear
  If rsTemp.RecordCount = 0 Then Exit Sub
    For i = 1 To rsTemp.RecordCount
        lvSplitCompany.ListItems.Add , , FormatDbValue(rsTemp!split_company_name)
        lvSplitCompany.ListItems(i).Tag = Trim(rsTemp!split_company_id)
      rsTemp.MoveNext
    Next i
    
 Exit Sub
ErrHandle:
 ShowErrorMsg
End Sub
Private Sub Form_Unload(Cancel As Integer)
 SaveFormPos Me
 Unload Me
End Sub

Private Sub lvSplitCompany_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
 If lvSplitCompany.SortOrder = lvwAscending Then
    lvSplitCompany.SortOrder = lvwDescending
 Else
    lvSplitCompany.SortOrder = lvwAscending
 End If
    lvSplitCompany.SortKey = ColumnHeader.Index - 1
    lvSplitCompany.Sorted = True
End Sub

Private Sub lvSplitCompany_DblClick()
On Error GoTo ErrHandle
 Dim i As Integer
 If lvSplitCompany.SelectedItem.Checked = False Then
    For i = 1 To lvSplitCompany.ListItems.Count
       lvSplitCompany.ListItems(i).Checked = True
    Next i
 Else
    For i = 1 To lvSplitCompany.ListItems.Count
       lvSplitCompany.ListItems(i).Checked = False
    Next i
 End If
 Exit Sub
ErrHandle:
ShowErrorMsg
End Sub
