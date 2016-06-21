VERSION 5.00
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Begin VB.Form frmChangeSheetNo 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "更改路单号"
   ClientHeight    =   1440
   ClientLeft      =   4260
   ClientTop       =   3630
   ClientWidth     =   4680
   ControlBox      =   0   'False
   HelpContextID   =   4003201
   Icon            =   "frmChangeSheetNo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Tag             =   "Modal"
   Begin RTComctl3.CoolButton cmdChangeSheetNo 
      Default         =   -1  'True
      Height          =   315
      Left            =   3360
      TabIndex        =   2
      Top             =   180
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "确定"
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
      MICON           =   "frmChangeSheetNo.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtSheetNo 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   0
      EndProperty
      Height          =   300
      Left            =   135
      TabIndex        =   1
      Text            =   "txtSheetNo"
      Top             =   540
      Width           =   2955
   End
   Begin RTComctl3.CoolButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   3360
      TabIndex        =   3
      Top             =   570
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "取消"
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
      MICON           =   "frmChangeSheetNo.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "请检查路单打印机上的路单,保持起始路单号与打印机路单号一致"
      Height          =   390
      Left            =   135
      TabIndex        =   4
      Top             =   990
      Width           =   4365
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "起始路单号(&N):"
      Height          =   180
      Left            =   165
      TabIndex        =   0
      Top             =   240
      Width           =   1260
   End
End
Attribute VB_Name = "frmChangeSheetNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Const cszSheetNoTooLong = "路单号长度设置为" & m_cnSheetNoLen & "位"
Const cszErrTitle = "错误"
Private mbFirstLoad As Boolean      '是否是启动时调用
Private mnCheckSheetLen As Integer


Private Sub cmdCancel_Click()
    If Trim(CStr(g_tCheckInfo.CurrSheetNo)) = "" Then
        MsgBox "未初始化路单号，必须指定起始路单号!", vbExclamation + vbOKOnly
        Exit Sub
    End If
    Unload Me
End Sub


Private Sub cmdChangeSheetNo_Click()
    Dim tTmp As TCheckSheetInfo
On Error GoTo ErrHandle
'    lblInfo.Caption = "检查路单号是否已经存在..."
'    lblInfo.Refresh
    
    txtSheetNo.Text = GetCodeStr(txtSheetNo.Text, mnCheckSheetLen)
    If Len(txtSheetNo.Text) > 10 Then
        txtSheetNo.Text = Right(txtSheetNo.Text, 10)
    End If
    
'   检验路单号是否已经被使用
'    tTmp = g_oChkTicket.GetCheckSheetInfo(txtSheetNo.Text)
'    If tTmp.szCheckSheet = "" Then
    g_tCheckInfo.CurrSheetNo = txtSheetNo.Text
    lblInfo.Caption = ""
    WriteInitReg
    If Not mbFirstLoad Then
        WriteCheckGateInfo
    End If
    Unload Me
'    Else
'        lblInfo.Caption = "'"
'        Me.MousePointer = vbDefault
'        MsgBox "此路单已存在!", vbExclamation + vbOKOnly
'        txtSheetNo.SetFocus
'    End If
    

    
'    If Err.Number <> ERR_ChkTkCheckSheetIDNotExist Then
'        If Err.Number = ERR_ChkTkCheckSheetIDIsNotValid Then
'            MsgBox "此路单号不合法！", vbCritical, "错误"
'        Else
'            MsgBox "此路单已存在！", vbCritical, "错误"
'        End If
'        txtSheetNo.SetFocus
'    Else
'        g_tCheckInfo.CheckSheet = txtSheetNo.Text
'        MDIMain.lblCurrentSheetNo = g_tCheckInfo.CheckSheet
'        Unload Me
'    End If
    Exit Sub
    
ErrHandle:
    MsgBox err.Description, vbCritical, err.Number & "-" & err.Description
'    Select Case Err.Number
'        Case ERR_ChkTkCheckSheetIDNotExist, ERR_ChkTkCheckSheetIDIsNotValid
'            Resume Next
'        Case Else
'            RunErrEvent Err.Number
'    End Select
End Sub


Private Sub cmdChangeSheetNo_GotFocus()
    txtSheetNo.SelStart = 0
    txtSheetNo.SelLength = Len(txtSheetNo.Text)
End Sub

Private Sub Form_Activate()
    txtSheetNo.SetFocus
End Sub

Private Sub Form_Load()
'    flbSheet.Caption = g_tCheckInfo.CurrSheetNo
    mnCheckSheetLen = g_nCheckSheetLen
    txtSheetNo.Text = g_tCheckInfo.CurrSheetNo
'    If Trim(txtSheetNo.Text) = "" Then cmdChangeSheetNo.Enabled = False
End Sub


Private Sub txtSheetNo_Change()
    If txtSheetNo.Text <> "" Then
        cmdChangeSheetNo.Enabled = True
        If Len(txtSheetNo.Text) > mnCheckSheetLen Then
            txtSheetNo.Text = Left(txtSheetNo.Text, mnCheckSheetLen)
        End If
    Else
        cmdChangeSheetNo.Enabled = False
    End If
End Sub

Private Sub txtSheetNo_GotFocus()
    txtSheetNo.SelLength = Len(txtSheetNo.Text)
End Sub

Private Sub txtSheetNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38 '向上方向键
            UpDown1_UpClick
        Case 40 '向下方向键
            UpDown1_DownClick
    End Select
End Sub

Private Sub txtSheetNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cmdChangeSheetNo.Enabled Then
            cmdChangeSheetNo.SetFocus
        End If
    End If
    If Not (KeyAscii > 47 And KeyAscii < 58 Or KeyAscii = 8) Then '0--9 and Backspace
        KeyAscii = 0
    End If
End Sub

Private Sub UpDown1_DownClick()
    If txtSheetNo.Text = "" Then txtSheetNo.Text = "0"
    txtSheetNo.Text = NumSub(txtSheetNo.Text, 1)
End Sub

Private Sub UpDown1_UpClick()
    If txtSheetNo.Text = "" Then txtSheetNo.Text = "0"
    txtSheetNo.Text = NumAdd(txtSheetNo.Text, 1)
End Sub


Public Property Let FirstLoad(ByVal vNewValue As Boolean)
    mbFirstLoad = vNewValue
End Property
Private Sub DisplayHelp(Optional HelpType As EHelpType = content)
    Dim lActiveControl As Long
    
    Select Case HelpType
        Case content
            lActiveControl = Me.ActiveControl.HelpContextID
            If lActiveControl = 0 Then
                TopicID = Me.HelpContextID
                CallHTMLShowTopicID
            Else
                TopicID = lActiveControl
                CallHTMLShowTopicID
            End If
        Case Index
            CallHTMLHelpIndex
        Case Support
            TopicID = clSupportID
            CallHTMLShowTopicID
    End Select

End Sub

