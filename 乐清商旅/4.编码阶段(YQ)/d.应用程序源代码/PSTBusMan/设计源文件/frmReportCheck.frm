VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Object = "{D11DF87C-EFEC-4838-B7E2-15462BF136FB}#1.4#0"; "RTReportlf.ocx"
Begin VB.Form frmReportCheck 
   BackColor       =   &H80000009&
   Caption         =   "车次检票报表"
   ClientHeight    =   6150
   ClientLeft      =   1545
   ClientTop       =   4275
   ClientWidth     =   9090
   Icon            =   "frmReportCheck.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6150
   ScaleWidth      =   9090
   WindowState     =   2  'Maximized
   Begin RTComctl3.Spliter spQuery 
      Height          =   1170
      Left            =   2910
      TabIndex        =   2
      Top             =   2565
      Width           =   165
      _ExtentX        =   291
      _ExtentY        =   2064
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox ptQuery 
      BackColor       =   &H8000000E&
      Height          =   5805
      Left            =   60
      ScaleHeight     =   5745
      ScaleWidth      =   2775
      TabIndex        =   1
      Top             =   60
      Width           =   2835
      Begin VB.ComboBox cboType 
         Height          =   300
         Left            =   165
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   720
         Width           =   2565
      End
      Begin RTComctl3.FloatLabel flbClose 
         Height          =   240
         Left            =   2490
         TabIndex        =   12
         Top             =   22
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   423
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Marlett"
            Size            =   9
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
         HoverBackColor  =   -2147483633
         Caption         =   "r"
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   165
         TabIndex        =   8
         Top             =   1395
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   556
         _Version        =   393216
         Format          =   60293120
         CurrentDate     =   36493
      End
      Begin RTComctl3.CoolButton Command2 
         Cancel          =   -1  'True
         Height          =   330
         Left            =   1455
         TabIndex        =   7
         Top             =   2535
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   582
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
         MICON           =   "frmReportCheck.frx":014A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin RTComctl3.CoolButton Command1 
         Default         =   -1  'True
         Height          =   330
         Left            =   150
         TabIndex        =   6
         Top             =   2535
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   582
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
         MICON           =   "frmReportCheck.frx":0166
         PICN            =   "frmReportCheck.frx":0182
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox txtBusID 
         Height          =   285
         Left            =   135
         TabIndex        =   5
         Top             =   2070
         Width           =   2505
      End
      Begin RTComctl3.FlatLabel FlatLabel1 
         Height          =   285
         Left            =   -15
         TabIndex        =   10
         Top             =   0
         Width           =   2790
         _ExtentX        =   4921
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
         VerticalAlignment=   1
         HorizontalAlignment=   1
         MarginLeft      =   0
         MarginTop       =   0
         Caption         =   "查询条件设定"
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         Caption         =   "报表类型:"
         Height          =   210
         Left            =   135
         TabIndex        =   14
         Top             =   480
         Width           =   2475
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         Caption         =   "车次代码:"
         Height          =   210
         Left            =   135
         TabIndex        =   4
         Top             =   1800
         Width           =   2475
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         Caption         =   "日期:"
         Height          =   300
         Left            =   135
         TabIndex        =   3
         Top             =   1155
         Width           =   2340
      End
   End
   Begin VB.PictureBox ptResult 
      BackColor       =   &H80000009&
      Height          =   5835
      Left            =   3105
      ScaleHeight     =   5775
      ScaleWidth      =   5145
      TabIndex        =   0
      Top             =   60
      Width           =   5205
      Begin RTReportLF.RTReport RTReport 
         Height          =   4335
         Left            =   60
         TabIndex        =   13
         Top             =   1305
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   7646
      End
      Begin VB.PictureBox ptQ 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1155
         Left            =   0
         Picture         =   "frmReportCheck.frx":051C
         ScaleHeight     =   1155
         ScaleWidth      =   5640
         TabIndex        =   9
         Top             =   0
         Width           =   5640
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "固定车次检票报表"
            BeginProperty Font 
               Name            =   "黑体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   825
            TabIndex        =   11
            Top             =   750
            Width           =   1920
         End
         Begin VB.Image imgOpen 
            Height          =   240
            Left            =   0
            Picture         =   "frmReportCheck.frx":1412
            Top             =   0
            Visible         =   0   'False
            Width           =   240
         End
      End
   End
End
Attribute VB_Name = "frmReportCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lMoveLeft As Long
Dim oReport As New CTReport
Dim mlRecordCount As Long
Const cszPtTop = 1200
Private Sub cboType_Click()
    lblTitle.Caption = cboType.Text
    If cboType.ListIndex = 0 Then
        txtBusID.BackColor = &HE0E0E0
        txtBusID.Enabled = False
    Else
        txtBusID.BackColor = &H80000005
        txtBusID.Enabled = True
    End If
End Sub

'Private WithEvents rtReport As CCellTemplate

Private Sub Command1_Click()
    Fill
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub



Private Sub flbClose_Click()
    ptQuery.Visible = False
    imgOpen.Visible = True
    lMoveLeft = 240
    spQuery.LayoutIt

End Sub

Private Sub Form_Activate()
    MDIScheme.SetPrintEnabled True
End Sub

Private Sub Form_Deactivate()
    MDIScheme.SetPrintEnabled False
End Sub


Private Sub Form_Load()
    oReport.Init g_oActiveUser
    spQuery.InitSpliter ptQuery, ptResult
    lMoveLeft = 0
    DTPicker1.Value = Date

    FillQueryType
End Sub

Private Sub imgClose_Click()
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIScheme.SetPrintEnabled False
End Sub

Private Sub imgOpen_Click()
    ptQuery.Visible = True
    imgOpen.Visible = False
    lMoveLeft = 0
'    lblTitle.Move 60 + lMoveLeft
    spQuery.LayoutIt
End Sub

Private Sub Form_Resize()
    spQuery.LayoutIt
End Sub

Private Sub lblClose_Click()
    ptQuery.Visible = False
    imgOpen.Visible = True
    lMoveLeft = 240
'    lblTitle.Move 60 + lMoveLeft
    spQuery.LayoutIt
End Sub

Private Sub lblOpen_Click()
    ptQuery.Visible = True
    imgOpen.Visible = False
    lMoveLeft = 0
'    lblTitle.Move 60 + lMoveLeft
    spQuery.LayoutIt
End Sub
Private Sub RTReport_SetProgressValue(ByVal lValue As Variant)
    WriteProcessBar , lValue, mlRecordCount
End Sub

Private Sub ptResult_Resize()
    Dim lTemp As Long
    lTemp = IIf((ptResult.ScaleHeight - cszPtTop) <= 0, lTemp, ptResult.ScaleHeight - cszPtTop)
    RTReport.Move 0, cszPtTop, ptResult.ScaleWidth, lTemp
    FlatLabel1.Width = ptQuery.ScaleWidth
    flbClose.Left = FlatLabel1.Left + FlatLabel1.Width - flbClose.Width - 30

End Sub

Public Sub Fill(Optional nDispRow As Integer = 20)
    Dim rsTemp As Recordset
    Dim vData As Variant
    Dim i, j As Integer, nCount As Integer, nStartRow, nEndRow As Integer, nEndCol As Integer
    
    Dim nLSumQuantitySale, nASumQuantitySale, nLSumQuantityChange, nASumQuantityChange, nLSumQuantityCheck, nASumQuantityCheck As Long
    Dim dLSumMoneySale, dASumMoneySale, dLSumMoneyCheck, dASumMoneyCheck  As Double
    
    On Error GoTo Error_Handle
    
    Me.Caption = cboType.Text
    ShowSBInfo ""
    WriteProcessBar True
    Select Case cboType.ListIndex
        Case 0
            RTReport.TemplateFile = App.Path & "\检票情况统计报表模板.xls"
            Set rsTemp = oReport.GetCheckTicketReport(DTPicker1.Value)
        Case 1
            RTReport.TemplateFile = App.Path & "\固定车次检票报表模板.xls"
            Set rsTemp = oReport.GetRegularCheckTicketReport(DTPicker1.Value, Trim(txtBusID))
        Case 2
            RTReport.TemplateFile = App.Path & "\流水车次检票情况报表模板.xls"
            Set rsTemp = oReport.GetScrollCheckTicketReport(DTPicker1.Value, Trim(txtBusID))
    End Select
    mlRecordCount = rsTemp.RecordCount
    RTReport.ShowReport rsTemp
    WriteProcessBar False
    ShowSBInfo ""
    ShowSBInfo "共有" & mlRecordCount & "条记录", ESB_ResultCountInfo

    
    Exit Sub
Error_Handle:
    WriteProcessBar False
    ShowSBInfo ""
    ShowErrorMsg
End Sub

Private Sub FillQueryType()
    cboType.Clear
    cboType.AddItem "检票情况统计报表"
    cboType.AddItem "固定车次检票报表"
    cboType.AddItem "流水车次检票情况报表"
    cboType.ListIndex = 0
    
End Sub

Private Sub txtBusId_Validate(Cancel As Boolean)
  If FindQuotationMarks(txtBusID, txtBusID) = False Then
       Cancel = True
  End If
End Sub
Public Sub PrintReport(Optional pbShowDialog As Boolean = False)
    On Error GoTo ErrorHandle
    RTReport.PrintReport pbShowDialog
    Exit Sub
ErrorHandle:
End Sub

Public Sub PreView()
    RTReport.PrintView
End Sub

Public Sub PageSet()
    RTReport.OpenDialog EDialogType.PAGESET_TYPE
End Sub

Public Sub PrintSet()
    RTReport.OpenDialog EDialogType.PRINT_TYPE
End Sub
'导出文件
Public Sub ExportFile()
    Dim szFileName As String
    szFileName = RTReport.OpenDialog(EDialogType.EXPORT_FILE)
End Sub
'导出文件并打开
Public Sub ExportFileOpen()
    Dim szFileName As String
    szFileName = RTReport.OpenDialog(EDialogType.EXPORT_FILE)
    If szFileName <> "" Then
        OpenLinkedFile szFileName
    End If
End Sub
