VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmRePrintLugSheet 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "行包结算统计"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   5880
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   -60
      TabIndex        =   24
      Top             =   810
      Width           =   7815
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   1215
      Left            =   -30
      TabIndex        =   0
      Top             =   4050
      Width           =   6045
      Begin RTComctl3.CoolButton cmdNext 
         Height          =   315
         Left            =   2310
         TabIndex        =   11
         Top             =   270
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "下一步(&N)"
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
         MICON           =   "frmRePrintLugSheet.frx":0000
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
         Height          =   315
         Left            =   4650
         TabIndex        =   1
         Top             =   270
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
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
         MICON           =   "frmRePrintLugSheet.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin RTComctl3.CoolButton cmdok 
         Height          =   315
         Left            =   3510
         TabIndex        =   2
         Top             =   270
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "完成(&F)"
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
         MICON           =   "frmRePrintLugSheet.frx":0038
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
   Begin VB.Frame fraWizLug2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   3165
      Left            =   30
      TabIndex        =   12
      Top             =   870
      Width           =   5895
      Begin MSComctlLib.ListView lvLugSheet 
         Height          =   1845
         Left            =   360
         TabIndex        =   13
         Top             =   660
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   3254
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "行包结算单"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "行包运费:"
         Height          =   180
         Left            =   3000
         TabIndex        =   23
         Top             =   1350
         Width           =   810
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "应拆金额:"
         Height          =   180
         Left            =   3030
         TabIndex        =   22
         Top             =   1800
         Width           =   810
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "拆算对象:"
         Height          =   180
         Left            =   3000
         TabIndex        =   21
         Top             =   840
         Width           =   810
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "拆算协议:"
         Height          =   180
         Left            =   3030
         TabIndex        =   20
         Top             =   2280
         Width           =   810
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "行包结算汇总信息:"
         Height          =   180
         Left            =   3000
         TabIndex        =   19
         Top             =   330
         Width           =   1530
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "行包结算单:"
         Height          =   180
         Left            =   360
         TabIndex        =   18
         Top             =   300
         Width           =   990
      End
      Begin VB.Label lblLubObject 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "苏州公司"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   4260
         TabIndex        =   17
         Top             =   870
         Width           =   780
      End
      Begin VB.Label lblLugTotalPrice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "4560"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   4260
         TabIndex        =   16
         Top             =   1290
         Width           =   420
      End
      Begin VB.Label lblLugNeedSplit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1215"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   4260
         TabIndex        =   15
         Top             =   1800
         Width           =   420
      End
      Begin VB.Label lblLugProtocol 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "协议比例2:8"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   4260
         TabIndex        =   14
         Top             =   2280
         Width           =   1095
      End
   End
   Begin VB.Frame fraWizLug 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   3165
      Left            =   0
      TabIndex        =   4
      Top             =   870
      Width           =   5925
      Begin VB.TextBox txtLugSheetID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   270
         TabIndex        =   7
         Text            =   "200340001"
         Top             =   600
         Width           =   1545
      End
      Begin RTComctl3.CoolButton cmdDetele 
         Height          =   345
         Left            =   2010
         TabIndex        =   5
         Top             =   1080
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "移除<<"
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
         MICON           =   "frmRePrintLugSheet.frx":0054
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin RTComctl3.CoolButton cmdAdd 
         Height          =   345
         Left            =   2010
         TabIndex        =   6
         Top             =   600
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "添加>>"
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
         MICON           =   "frmRePrintLugSheet.frx":0070
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComctlLib.ListView lvLugSheetID 
         Height          =   2175
         Left            =   3180
         TabIndex        =   8
         Top             =   630
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   3836
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "行包结算单"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "行包结算单号:"
         Height          =   180
         Left            =   270
         TabIndex        =   10
         Top             =   270
         Width           =   1170
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "所选的结算单:"
         Height          =   180
         Left            =   3240
         TabIndex        =   9
         Top             =   240
         Width           =   1170
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "行包结算汇总"
      Height          =   180
      Left            =   390
      TabIndex        =   3
      Top             =   330
      Width           =   1080
   End
End
Attribute VB_Name = "frmRePrintLugSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mObjectType As Integer

Private Sub cmdAdd_Click()
    AddLugSheet
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDetele_Click()
     On Error GoTo ErrHandle
  If lvLugSheetID.ListItems.Count = 0 Then
        cmdDetele.Enabled = False
        Exit Sub
  End If
  lvLugSheetID.ListItems.Remove (lvLugSheetID.SelectedItem.Index)
 
 Exit Sub
ErrHandle:
 ShowErrorMsg
End Sub

Private Sub cmdNext_Click()
On Error GoTo here
    If cmdNext.Caption = "下一步(&N)" Then
        cmdNext.Caption = "上一步(&P)"
        fraWizLug.Visible = False
        fraWizLug2.Visible = True
        StatLugInfo
        cmdOk.Enabled = True
    Else
        cmdNext.Caption = "下一步(&N)"
        fraWizLug.Visible = True
        fraWizLug2.Visible = False
        cmdOk.Enabled = False
    End If
    
    Exit Sub
here:
    ShowErrorMsg
End Sub

Private Sub cmdok_Click()
On Error GoTo here
    Dim i As Integer
    Dim szTemp As String
    
    
'    frmPrintFinSheet.m_dbTotalPrice = Val(lblLugTotalPrice.Caption)
'    frmPrintFinSheet.m_dbNeedSplitPrice = Val(lblLugNeedSplit.Caption)
'    frmPrintFinSheet.m_szProtocol = lblLugProtocol.Caption
    
    For i = 1 To lvLugSheet.ListItems.Count
        If i <> lvLugSheet.ListItems.Count Then
            szTemp = szTemp & lvLugSheet.ListItems(i).Text & ","
        Else
            szTemp = szTemp & lvLugSheet.ListItems(i).Text
        End If
    Next i
    frmPrintFinSheet.m_szLugSettleSheetID = szTemp
    frmPrintFinSheet.m_bRePrint = True
    
    Unload Me

    Exit Sub
here:
    ShowErrorMsg
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete And (ActiveControl Is lvLugSheetID) Then
        If lvLugSheetID.ListItems.Count = 0 Then
            cmdDetele.Enabled = False
            Exit Sub
        End If
        lvLugSheetID.ListItems.Remove (lvLugSheetID.SelectedItem.Index)
    End If
End Sub

Private Sub Form_Load()
    AlignFormPos Me
    HandleLugInfo
    
    fraWizLug.Visible = True
    fraWizLug2.Visible = False
    cmdOk.Enabled = False
    cmdNext.Caption = "下一步(&N)"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveFormPos Me
    Unload Me
    
End Sub


'处理行包结算信息
Private Sub HandleLugInfo()
    txtLugSheetID.Text = ""
    lvLugSheetID.ListItems.Clear
    cmdDetele.Enabled = False
'    txtLugSheetID.SetFocus
End Sub

'处理行包结算统计
Private Sub StatLugInfo()
On Error GoTo here
    Dim rsTemp As Recordset
    Dim m_oReport As New Report
    Dim szaLugSheetID() As String
    Dim i As Integer
    lblLubObject.Caption = ""
    lblLugTotalPrice.Caption = ""
    lblLugNeedSplit.Caption = ""
    lblLugProtocol.Caption = ""
    
    If lvLugSheetID.ListItems.Count = 0 Then Exit Sub
    lvLugSheet.ListItems.Clear
    For i = 1 To lvLugSheetID.ListItems.Count
        lvLugSheet.ListItems.Add , , lvLugSheetID.ListItems.Item(i).Text
    Next i
    If lvLugSheet.ListItems.Count = 0 Then Exit Sub
    ReDim szaLugSheetID(1 To lvLugSheet.ListItems.Count)
    For i = 1 To lvLugSheet.ListItems.Count
        szaLugSheetID(i) = Trim(lvLugSheet.ListItems.Item(i).Text)
    Next i
    m_oReport.Init g_oActiveUser
    Set rsTemp = m_oReport.PreLugFinSheet(szaLugSheetID)
    If rsTemp.RecordCount > 0 Then
        mObjectType = FormatDbValue(rsTemp!split_object_type)
        lblLubObject.Caption = FormatDbValue(rsTemp!split_object_name)
        lblLugTotalPrice.Caption = FormatDbValue(rsTemp!total_price)
        lblLugNeedSplit.Caption = FormatDbValue(rsTemp!need_split_out)
        lblLugProtocol.Caption = FormatDbValue(rsTemp!protocol_name)
    End If
    
    Exit Sub
here:
    ShowErrorMsg
End Sub

Private Sub txtLugSheetID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        AddLugSheet
    End If
End Sub

Private Sub AddLugSheet()
 On Error GoTo ErrHandle
    Dim i As Integer
    Dim m_oReport As New Report
    Dim rsTemp As Recordset
    '判断行包单是否有效,拆算对象是否为所指定的对象
    m_oReport.Init g_oActiveUser
    
    '满足条件
    If txtLugSheetID.Text <> "" Then
        Set rsTemp = m_oReport.GetLugSheet(Trim(txtLugSheetID.Text))
        If rsTemp.RecordCount = 0 Then
            MsgBox "此行包结算单无效!", vbExclamation, Me.Caption
            Exit Sub
        End If
        lvLugSheetID.ListItems.Add , , Trim(txtLugSheetID.Text)
        txtLugSheetID.Text = ""
        txtLugSheetID.SetFocus
        cmdDetele.Enabled = True
    End If
    Exit Sub
ErrHandle:
 ShowErrorMsg

End Sub
