VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmBusVehicleMan 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "车次车辆"
   ClientHeight    =   4065
   ClientLeft      =   1755
   ClientTop       =   4875
   ClientWidth     =   6570
   FillColor       =   &H00FFFFFF&
   HelpContextID   =   10000750
   Icon            =   "frmBusVehicleMan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin RTComctl3.CoolButton CoolButton1 
      Height          =   315
      Left            =   2910
      TabIndex        =   8
      Top             =   3600
      Width           =   1125
      _ExtentX        =   1984
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
      MICON           =   "frmBusVehicleMan.frx":014A
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
      Default         =   -1  'True
      Height          =   315
      Left            =   4140
      TabIndex        =   7
      Top             =   3600
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
      MICON           =   "frmBusVehicleMan.frx":0166
      PICN            =   "frmBusVehicleMan.frx":0182
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
      Left            =   5370
      TabIndex        =   6
      Top             =   3600
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
      MICON           =   "frmBusVehicleMan.frx":051C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.PictureBox ptTitle 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   795
      Left            =   -30
      ScaleHeight     =   795
      ScaleWidth      =   7185
      TabIndex        =   3
      Top             =   0
      Width           =   7185
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   60
         Left            =   0
         TabIndex        =   4
         Top             =   750
         Width           =   7215
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车次车辆列表(&L):"
         Height          =   180
         Left            =   120
         TabIndex        =   5
         Top             =   337
         Width           =   1440
      End
   End
   Begin VB.TextBox txtChange 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   2010
      TabIndex        =   2
      Top             =   3600
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.ComboBox cboChange 
      Appearance      =   0  'Flat
      Height          =   300
      ItemData        =   "frmBusVehicleMan.frx":0538
      Left            =   660
      List            =   "frmBusVehicleMan.frx":0542
      TabIndex        =   1
      Text            =   "(不停)"
      Top             =   3600
      Visible         =   0   'False
      Width           =   1080
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgBusVe 
      Height          =   2700
      Left            =   0
      TabIndex        =   0
      Top             =   780
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   4763
      _Version        =   393216
      Cols            =   6
      FixedCols       =   3
      BackColorFixed  =   14737632
      BackColorBkg    =   14737632
      AllowUserResizing=   3
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   6
   End
End
Attribute VB_Name = "frmBusVehicleMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_oBus As Bus '车次对象的引用 Bus
Private m_szaBusVehicle() As String '1车次车辆
Private m_nVehicleCount As Integer '车辆数
Const cszNoStop = "(不停)"
Const cszLongStop = "(长停)"

Private Sub cboChange_Click()
    cboChange_Change
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
    Dim szaReCode() As String
    Dim taBusVehicle() As TBusVehicleInfo
    Dim tBusVehicle As TBusVehicleInfo
    Dim bDate As Boolean
    Dim i As Integer
    On Error GoTo ErrHandle
    cboChange.Visible = False
    txtChange.Visible = False
    For i = 1 To hfgBusVe.Rows - 1
        bDate = True
        hfgBusVe.Row = i
        If hfgBusVe.CellForeColor = vbBlue Then
            hfgBusVe.Col = 0
            hfgBusVe.CellForeColor = vbBlack
            hfgBusVe.Col = 1
            hfgBusVe.CellForeColor = vbBlack
            hfgBusVe.Col = 2
            hfgBusVe.CellForeColor = vbBlack
            hfgBusVe.Col = 3
            hfgBusVe.CellForeColor = vbBlack
            hfgBusVe.Col = 4
            hfgBusVe.CellForeColor = vbBlack
            hfgBusVe.Col = 5
            hfgBusVe.CellForeColor = vbBlack
            With hfgBusVe
            tBusVehicle.szVehicleID = .TextArray(i * 6 + 1)
            tBusVehicle.nSerialNo = Val(.TextArray(i * 6 + 0))
            tBusVehicle.nStandTicketCount = .TextArray(i * 6 + 3)
            If .TextArray(i * 6 + 5) = cszLongStop Then '车辆是否长停
                If .TextArray(i * 6 + 4) = cszLongStop Then '车辆是否从某日期开始长停
                    tBusVehicle.dtBeginStopDate = Now
                Else
                    tBusVehicle.dtBeginStopDate = CDate(.TextArray(i * 6 + 4))
                End If
                tBusVehicle.dtEndStopDate = CDate(cszForeverDateStr)
            bDate = False
            End If
            If .TextArray(i * 6 + 4) = cszNoStop Then
                tBusVehicle.dtBeginStopDate = CDate(cszEmptyDateStr)
                tBusVehicle.dtEndStopDate = CDate(cszEmptyDateStr)
            bDate = False
            End If
            If bDate Then
                tBusVehicle.dtBeginStopDate = CDate(.TextArray(i * 6 + 4))
                tBusVehicle.dtEndStopDate = CDate(.TextArray(i * 6 + 5))
            End If
            m_oBus.ModifyRunVehicle tBusVehicle
            End With
        End If
        Next
'        taBusVehicle = m_oBus.GetAllVehicle
        frmArrangeBus.ChangeVehicle taBusVehicle
        MsgBox "车次车辆属性已保存", vbInformation, "计划"
        cmdOk.Enabled = False
        frmArrangeBus.RefreshVehicle
    Exit Sub
ErrHandle:
        If err.Number = 13 Then
        MsgBox "输入的日期不正确", vbExclamation + vbOKOnly, "日期错"
        Else
        ShowErrorMsg
        End If
End Sub

Private Sub CoolButton1_Click()
DisplayHelp Me
End Sub

Private Sub Form_Load()
Dim bDate As Boolean
Dim i As Integer, j As Integer
On Error GoTo ErrHandle
AlignFormPos Me

m_nVehicleCount = ArrayLength(m_szaBusVehicle)
hfgBusVe.Rows = m_nVehicleCount + 1
hfgBusVe.TextArray(0) = "序"
hfgBusVe.TextArray(1) = "车辆代码"
hfgBusVe.TextArray(2) = "车牌"
hfgBusVe.TextArray(3) = "站票数"
hfgBusVe.TextArray(4) = "停班开始日期"
hfgBusVe.TextArray(5) = "停班结束日期"
hfgBusVe.ColWidth(0) = 400
hfgBusVe.ColWidth(1) = 800
hfgBusVe.ColWidth(2) = 1200
hfgBusVe.ColWidth(3) = 600
hfgBusVe.ColWidth(4) = 1400
hfgBusVe.ColWidth(5) = 1400
cboChange.AddItem Format(Now, "YYYY-MM-DD")
For i = 1 To m_nVehicleCount
    bDate = True
    hfgBusVe.TextArray(i * 6 + 0) = m_szaBusVehicle(i, 2)
    hfgBusVe.TextArray(i * 6 + 1) = Trim(m_szaBusVehicle(i, 1))
    hfgBusVe.TextArray(i * 6 + 2) = Trim(m_szaBusVehicle(i, 6))
    hfgBusVe.TextArray(i * 6 + 3) = m_szaBusVehicle(i, 3)
    If Format(m_szaBusVehicle(i, 4), "YYYY-MM-DD") = cszEmptyDateStr Then
        hfgBusVe.Row = i
        hfgBusVe.TextArray(i * 6 + 4) = cszNoStop
        hfgBusVe.TextArray(i * 6 + 5) = cszNoStop
        bDate = False
    End If
    If Format(m_szaBusVehicle(i, 5), "YYYY-MM-DD") = cszForeverDateStr Then
        hfgBusVe.Row = i
        hfgBusVe.TextArray(i * 6 + 4) = Format(m_szaBusVehicle(i, 4), "YYYY-MM-DD")
        hfgBusVe.TextArray(i * 6 + 5) = cszLongStop
        bDate = False
    End If
    If bDate Then
        hfgBusVe.TextArray(i * 6 + 4) = Format(m_szaBusVehicle(i, 4), "YYYY-MM-DD")
        hfgBusVe.TextArray(i * 6 + 5) = Format(m_szaBusVehicle(i, 5), "YYYY-MM-DD")
    End If
Next
    cmdOk.Enabled = False
Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Public Sub Init(vData As Object)
On Error GoTo ErrHandle
Set m_oBus = vData
m_szaBusVehicle = m_oBus.GetAllVehicle
Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveFormPos Me
End Sub

Private Sub hfgBusVe_Click()
    Const cnMargin = 15
    If hfgBusVe.Row = 0 Then Exit Sub
    If hfgBusVe.Col = 3 Then
        cboChange.Visible = False
        txtChange.Visible = True
        txtChange.Width = hfgBusVe.CellWidth
        txtChange.Top = hfgBusVe.Top + hfgBusVe.CellTop - cnMargin
        txtChange.Height = hfgBusVe.CellHeight - 2 * cnMargin
        txtChange.Left = hfgBusVe.Left + hfgBusVe.CellLeft
        txtChange.Text = hfgBusVe.Text
        txtChange.SetFocus
        Exit Sub
    Else
        cboChange.Visible = False
        txtChange.Visible = False
    End If
    If hfgBusVe.Col = 4 Or hfgBusVe.Col = 5 Then
        If hfgBusVe.Col <> 4 Then
            If hfgBusVe.TextArray(hfgBusVe.Row * 6 + 4) = cszLongStop Or hfgBusVe.TextArray(hfgBusVe.Row * 6 + 4) = cszNoStop Then Exit Sub
        End If
        txtChange.Visible = False
        cboChange.Visible = True
        cboChange.Width = hfgBusVe.CellWidth
        cboChange.Top = hfgBusVe.Top + hfgBusVe.CellTop - cnMargin
'        cboChange.Height = hfgBusVe.CellHeight - 2 * cnMargin
        cboChange.Left = hfgBusVe.Left + hfgBusVe.CellLeft
        cboChange.Text = hfgBusVe.Text
        cboChange.SetFocus
    Else
        txtChange.Visible = False
        cboChange.Visible = False
    End If
End Sub
Private Sub hfgBusVe_Scroll()
    cboChange.Visible = False
End Sub

Private Sub cboChange_Change()
    Dim nCol As Integer
    Dim szTemp As String
    szTemp = hfgBusVe.Text
    hfgBusVe.Text = Trim(cboChange.Text)
    If Trim(szTemp) <> Trim(hfgBusVe.Text) Then
        nCol = hfgBusVe.Col
        hfgBusVe.Col = 0
        hfgBusVe.CellForeColor = cvChangeColor
        hfgBusVe.Col = 1
        hfgBusVe.CellForeColor = cvChangeColor
        hfgBusVe.Col = 2
        hfgBusVe.CellForeColor = cvChangeColor
        hfgBusVe.Col = 3
        hfgBusVe.CellForeColor = cvChangeColor
        hfgBusVe.Col = 4
        hfgBusVe.CellForeColor = cvChangeColor
        hfgBusVe.Col = 5
        hfgBusVe.CellForeColor = cvChangeColor
        hfgBusVe.Col = nCol
        If nCol = 4 Then
        hfgBusVe.Col = 5
        hfgBusVe.Text = Trim(cboChange.Text)
        hfgBusVe.Col = 0
        hfgBusVe.CellForeColor = cvChangeColor
        hfgBusVe.Col = 1
        hfgBusVe.CellForeColor = cvChangeColor
        hfgBusVe.Col = 2
        hfgBusVe.CellForeColor = cvChangeColor
        hfgBusVe.Col = 3
        hfgBusVe.CellForeColor = cvChangeColor
        hfgBusVe.Col = 4
        hfgBusVe.CellForeColor = cvChangeColor
        hfgBusVe.Col = 5
        hfgBusVe.CellForeColor = cvChangeColor
        hfgBusVe.Col = nCol
        End If
        If nCol = 4 And cboChange.Text = cszNoStop Then
        hfgBusVe.Col = 5
        hfgBusVe.Text = Trim(cboChange.Text)
        hfgBusVe.Col = 0
        hfgBusVe.CellForeColor = cvChangeColor
        hfgBusVe.Col = 1
        hfgBusVe.CellForeColor = cvChangeColor
        hfgBusVe.Col = 2
        hfgBusVe.CellForeColor = cvChangeColor
        hfgBusVe.Col = 3
        hfgBusVe.CellForeColor = cvChangeColor
        hfgBusVe.Col = 4
        hfgBusVe.CellForeColor = cvChangeColor
        hfgBusVe.Col = 5
        hfgBusVe.CellForeColor = cvChangeColor
        hfgBusVe.Col = nCol
        End If
    End If
    cmdOk.Enabled = True
End Sub

Private Sub txtChange_Change()
    Dim nCol As Integer
    Dim szTemp As String
    szTemp = hfgBusVe.Text
    hfgBusVe.Text = Trim(txtChange.Text)
    If Trim(szTemp) <> Trim(hfgBusVe.Text) Then
        nCol = hfgBusVe.Col
        hfgBusVe.Col = 0
        hfgBusVe.CellForeColor = cvChangeColor
        hfgBusVe.Col = 1
        hfgBusVe.CellForeColor = cvChangeColor
        hfgBusVe.Col = 2
        hfgBusVe.CellForeColor = cvChangeColor
        hfgBusVe.Col = 3
        hfgBusVe.CellForeColor = cvChangeColor
        hfgBusVe.Col = 4
        hfgBusVe.CellForeColor = cvChangeColor
        hfgBusVe.Col = 5
        hfgBusVe.CellForeColor = cvChangeColor
        hfgBusVe.Col = nCol
    End If
    cmdOk.Enabled = True
End Sub
