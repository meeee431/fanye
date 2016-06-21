VERSION 5.00
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Begin VB.Form FrmSaveFormular 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "保存公式"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox TxtFormularID 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1470
      MaxLength       =   4
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
   Begin VB.TextBox TxtFormularName 
      Height          =   315
      IMEMode         =   1  'ON
      Left            =   1470
      MaxLength       =   20
      TabIndex        =   1
      Top             =   720
      Width           =   2535
   End
   Begin RTComctl3.CoolButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   345
      Left            =   2640
      TabIndex        =   2
      Top             =   1230
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
      MICON           =   "FrmSaveFormular.frx":0000
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
      Left            =   1200
      TabIndex        =   3
      ToolTipText     =   "保存协议"
      Top             =   1230
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
      MICON           =   "FrmSaveFormular.frx":001C
      PICN            =   "FrmSaveFormular.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "公式代码(&C):"
      Height          =   180
      Left            =   360
      TabIndex        =   5
      Top             =   300
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "公式名称(&N):"
      Height          =   180
      Left            =   360
      TabIndex        =   4
      Top             =   810
      Width           =   1080
   End
End
Attribute VB_Name = "FrmSaveFormular"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private m_oSplit As New Split
Public FormularContent As String
Public m_eStatus As eFormStatus

Private Sub cmdCancel_Click()
    Unload Me
    
End Sub

Private Sub cmdOk_Click()
   On Error GoTo ErrorHandle
   Dim i As Variant
 Select Case m_eStatus
        Case ST_AddObj
            m_oLugFormula.AddNew
            m_oLugFormula.FormulaID = TxtFormularID.Text
            m_oLugFormula.FormulaName = TxtFormularName.Text
            m_oLugFormula.FormulaText = FormularContent
            
            m_oLugFormula.Update
            frmFormula.txtRegFormula = m_oLugFormula.FormulaText
            frmFormula.lstFormula.AddItem MakeDisplayString(TxtFormularID.Text, TxtFormularName)
        Case ST_EditObj
            i = MsgBox("你确定要修改该计算公式吗?", vbYesNo + vbQuestion, "计算公式")
            If i = vbYes Then
                m_oLugFormula.Identify TxtFormularID.Text
                m_oLugFormula.FormulaName = TxtFormularName.Text
                m_oLugFormula.FormulaText = FormularContent
                
                m_oLugFormula.Update
                frmFormula.lstFormula.clear
                frmFormula.txtRegFormula = m_oLugFormula.FormulaText
                frmFormula.lstFormula.AddItem MakeDisplayString(TxtFormularID.Text, TxtFormularName)
            Else
                Unload Me
                Exit Sub
            End If
                   
 End Select
         '将值放入数组中，返回给基本信息窗口
    Dim aszInfo(0 To 3) As String
    aszInfo(0) = TxtFormularID.Text
    aszInfo(1) = TxtFormularName.Text
    aszInfo(2) = FormularContent

    
    '刷新基本信息窗体
    Dim oListItem As ListItem
    If m_eStatus = ST_EditObj Then
        frmBaseInfo.UpdateList aszInfo
      
    End If
    If m_eStatus = ST_AddObj Then
        frmBaseInfo.AddList aszInfo
    End If
  Unload Me
 Exit Sub
ErrorHandle:
 ShowErrorMsg
End Sub

Private Sub Form_Load()
     AlignFormPos Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
     SaveFormPos Me
End Sub

Private Sub TxtFormularID_Change()
 If Trim(TxtFormularID.Text) <> "" Then
     cmdOk.Enabled = True
 Else
     cmdOk.Enabled = False
 End If
End Sub

Private Sub TxtFormularID_KeyPress(KeyAscii As Integer)
    
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

