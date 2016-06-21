VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmAddPriceTable 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "新增线路票价表"
   ClientHeight    =   3375
   ClientLeft      =   2865
   ClientTop       =   2940
   ClientWidth     =   5895
   Icon            =   "frmAddPriceTable.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   -15
      ScaleHeight     =   735
      ScaleWidth      =   6555
      TabIndex        =   11
      Top             =   0
      Width           =   6555
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请输入票价表信息:"
         Height          =   180
         Left            =   270
         TabIndex        =   12
         Top             =   270
         Width           =   1530
      End
   End
   Begin RTComctl3.CoolButton CoolButton1 
      Height          =   315
      Left            =   570
      TabIndex        =   10
      Top             =   2955
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "帮助"
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
      MICON           =   "frmAddPriceTable.frx":014A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   60
      Left            =   -45
      TabIndex        =   9
      Top             =   720
      Width           =   6885
   End
   Begin MSComCtl2.DTPicker dtpExecute 
      Height          =   285
      Left            =   2250
      TabIndex        =   5
      Top             =   2010
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   503
      _Version        =   393216
      Format          =   123994112
      CurrentDate     =   36997
   End
   Begin RTComctl3.CoolButton cmdOk 
      Height          =   315
      Left            =   2670
      TabIndex        =   6
      Top             =   2955
      Width           =   1215
      _ExtentX        =   2143
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
      MICON           =   "frmAddPriceTable.frx":0166
      PICN            =   "frmAddPriceTable.frx":0182
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
      Left            =   4110
      TabIndex        =   7
      Top             =   2955
      Width           =   1215
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
      MICON           =   "frmAddPriceTable.frx":051C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtTableName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2250
      MaxLength       =   16
      TabIndex        =   3
      Top             =   1590
      Width           =   2865
   End
   Begin VB.TextBox txtTableID 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2250
      MaxLength       =   4
      TabIndex        =   1
      Top             =   1200
      Width           =   2865
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   " "
      Enabled         =   0   'False
      Height          =   3120
      Left            =   -105
      TabIndex        =   8
      Top             =   2745
      Width           =   8745
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "开始执行日期(&T):"
      Height          =   180
      Left            =   735
      TabIndex        =   4
      Top             =   2070
      Width           =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "票价表名(&N):"
      Height          =   180
      Left            =   735
      TabIndex        =   2
      Top             =   1650
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "票价表代码(&D):"
      Height          =   180
      Left            =   735
      TabIndex        =   0
      Top             =   1245
      Width           =   1260
   End
End
Attribute VB_Name = "frmAddPriceTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************
'* Source File Name:frmAERPTable.frm
'* Project Name:PSTBusMan.vbp
'* Engineer:陈峰
'* Date Generated:2002/09/03
'* Last Revision Date:2002/09/03
'* Brief Description:修改线路票价表
'* Relational Document:
'**********************************************************

Option Explicit

Public m_bIsParent As Boolean '是否是父窗体直接调用
Public m_szTableID As String '票价表代码
Public m_eStatus As EFormStatus
Private m_oPriceTable As New RoutePriceTable

Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub cmdOk_Click()
    Dim szExcutePriceTable() As String
    Dim liTemp As ListItem

    On Error GoTo ErrorHandle
    If dtpExecute.Value < Date Then
        MsgBox "票价表的起始执行日期不能为当前日期以前的日期！", vbOKOnly + vbInformation, Me.Caption
        Exit Sub
    End If
    SetBusy
    If m_eStatus = EFS_AddNew Then
        m_oPriceTable.AddNew
        m_oPriceTable.RoutePriceTableID = txtTableID.Text
'        m_oPriceTable.BusProject = g_szExePlanID
        m_oPriceTable.RoutePriceTableName = Trim(txtTableName.Text)
        m_oPriceTable.StartRunTime = FormatDateTime(Format(dtpExecute.Value, "YYYY-MM-DD"))
        m_oPriceTable.Update
        If m_bIsParent Then
            frmPriceTableMan.AddList txtTableID.Text
        End If
    ElseIf m_eStatus = EFS_Modify Then
        m_oPriceTable.Identify txtTableID.Text
        m_oPriceTable.RoutePriceTableName = Trim(txtTableName.Text)
        m_oPriceTable.StartRunTime = FormatDateTime(Format(dtpExecute.Value, "YYYY-MM-DD"))
        m_oPriceTable.Update
        If m_bIsParent Then
            frmPriceTableMan.UpdateList txtTableID.Text
        End If
    End If
    
    SetNormal
    Unload Me

    Exit Sub
ErrorHandle:
    SetNormal
    ShowErrorMsg
End Sub




Private Sub CoolButton1_Click()
DisplayHelp Me
End Sub

Private Sub dtpExecute_Change()
    EnableOk
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
    Dim nTemp As Integer, i As Integer
    m_oPriceTable.Init g_oActiveUser
    
    If m_eStatus = EFS_AddNew Then
        Me.Caption = "新增线路票价表"
        dtpExecute.Value = Date
'        cboProject.Enabled = False
    Else
        txtTableID.Text = m_szTableID
        Me.Caption = "线路票价表属性"
        RefreshPriceTable
        
    End If
    cmdOk.Enabled = False
End Sub


Private Sub RefreshPriceTable()
    '刷新票价表信息
    m_oPriceTable.Identify txtTableID.Text
    txtTableName.Text = m_oPriceTable.RoutePriceTableName
    dtpExecute.Value = m_oPriceTable.StartRunTime
End Sub


Private Sub EnableOk()
    If txtTableID.Text = "" Or txtTableName.Text = "" Then
        cmdOk.Enabled = False
    Else
        cmdOk.Enabled = True
    End If
End Sub


Private Sub txtTableID_Change()
    EnableOk
End Sub

Private Sub txtTableName_Change()
    EnableOk
End Sub

