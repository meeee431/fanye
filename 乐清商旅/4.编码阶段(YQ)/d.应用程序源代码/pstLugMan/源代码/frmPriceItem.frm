VERSION 5.00
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmPriceItem 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "行包收费项设置"
   ClientHeight    =   3060
   ClientLeft      =   3720
   ClientTop       =   3885
   ClientWidth     =   6105
   HelpContextID   =   7000260
   Icon            =   "frmPriceItem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   0
      TabIndex        =   13
      Top             =   690
      Width           =   8115
   End
   Begin VB.PictureBox ptCaption 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   7965
      TabIndex        =   11
      Top             =   0
      Width           =   7965
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请定义折算计算公式:"
         Height          =   180
         Left            =   270
         TabIndex        =   12
         Top             =   240
         Width           =   1710
      End
   End
   Begin VB.TextBox txtFormulaName 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2580
      TabIndex        =   2
      Top             =   1530
      Width           =   2505
   End
   Begin VB.CheckBox chkUsed 
      BackColor       =   &H00E0E0E0&
      Caption         =   "是否使用"
      Height          =   270
      Left            =   1125
      TabIndex        =   3
      Top             =   1965
      Width           =   1365
   End
   Begin RTComctl3.CoolButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   345
      Left            =   4620
      TabIndex        =   6
      Top             =   2580
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   609
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
      MICON           =   "frmPriceItem.frx":014A
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
      Left            =   3210
      TabIndex        =   7
      ToolTipText     =   "保存协议"
      Top             =   2580
      Width           =   1140
      _ExtentX        =   2011
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
      MICON           =   "frmPriceItem.frx":0166
      PICN            =   "frmPriceItem.frx":0182
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
      Caption         =   " "
      Height          =   930
      Left            =   -120
      TabIndex        =   8
      Top             =   2280
      Width           =   8745
      Begin RTComctl3.CoolButton cmdHelp 
         Height          =   315
         Left            =   270
         TabIndex        =   10
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
         MICON           =   "frmPriceItem.frx":051C
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
   Begin VB.Label lblAcceptType 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "快件"
      Height          =   180
      Left            =   2580
      TabIndex        =   9
      Top             =   1260
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "托运方式:"
      Height          =   180
      Left            =   1140
      TabIndex        =   5
      Top             =   1275
      Width           =   810
   End
   Begin VB.Label lblPriteItemID 
      BackStyle       =   0  'Transparent
      Caption         =   "0001"
      Height          =   225
      Left            =   2580
      TabIndex        =   4
      Top             =   945
      Width           =   615
   End
   Begin VB.Label lblExcuteTable 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "收费项代码:"
      Height          =   180
      Left            =   1140
      TabIndex        =   0
      Top             =   960
      Width           =   990
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "收费项名称(N):"
      Height          =   180
      Left            =   1140
      TabIndex        =   1
      Top             =   1605
      Width           =   1260
   End
End
Attribute VB_Name = "frmPriceItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************
'* Source File Name:frmpriceItem.frm
'* Project Name:PSTLugMan.vbp
'* Engineer:王候记
'* Date Generated:2003/01/25
'* Last Revision Date:2003/01/25
'* Brief Description:修改行包票价公式
'* Relational Document:
'**********************************************************

Option Explicit

Public m_bIsParent As Boolean '是否父窗体调用
Public m_szPriceItemId As String '公式代码
Public m_szAcceptType As Integer '托运方式



Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    DisplayHelp Me
End Sub


Private Sub cmdOk_Click()
    '保存设置
On Error GoTo ErrorHandle
    Dim tTemp As TLuggagePriceItemFormulaEx

    tTemp.PriceItem = Trim(lblPriteItemID.Caption)
    If lblAcceptType.Caption = szAcceptTypeGeneral Then
        tTemp.AcceptType = 0
    Else
        tTemp.AcceptType = 1
    End If
    tTemp.PriceItemName = txtFormulaName.Text
    If chkUsed.Value = 1 Then
        tTemp.UsedMark = 0
    Else
        tTemp.UsedMark = 1
    End If

    m_oLugParam.SetPriceItem tTemp

    Dim aszInfo(0 To 3) As String
    aszInfo(0) = Trim(lblPriteItemID.Caption)
    aszInfo(1) = Trim(txtFormulaName.Text)
    aszInfo(2) = Trim(lblAcceptType.Caption)
    If chkUsed.Value = 1 Then
       aszInfo(3) = "是"
    Else
        aszInfo(3) = "否"
    End If
    frmBaseInfo.UpdateList aszInfo
    Unload Me

  Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
On Error GoTo ErrorHandle
    FillPriceItem '填充票价项信息

    AlignFormPos Me
    Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub


Private Sub Form_Unload(Cancel As Integer)
  SaveFormPos Me
End Sub


'从数据库中取出相应的票价项信息

Public Sub FillPriceItem()
    Dim szFormulaTemp As String
    Dim rsTemp As New Recordset
    Dim i As Integer
    lblPriteItemID.Caption = m_szPriceItemId
    Set rsTemp = m_oLugParam.GetPriceItem(m_szPriceItemId, m_szAcceptType)
    If rsTemp!accept_type = 0 Then
        lblAcceptType.Caption = szAcceptTypeGeneral
    Else
        lblAcceptType.Caption = szAcceptTypeMan
    End If
    If rsTemp!use_mark = 0 Then
        chkUsed.Value = 1
    Else
        chkUsed.Value = 0
    End If
    txtFormulaName.Text = rsTemp!chinese_name


End Sub


