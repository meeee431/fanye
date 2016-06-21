VERSION 5.00
Object = "{4932CEF1-2CAA-11D2-A165-0060081C43D9}#2.0#0"; "actbar2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Begin VB.Form frmVehicleProtocol 
   BackColor       =   &H00E0E0E0&
   Caption         =   "车辆协议管理"
   ClientHeight    =   7440
   ClientLeft      =   2835
   ClientTop       =   2580
   ClientWidth     =   11310
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7440
   ScaleWidth      =   11310
   WindowState     =   2  'Maximized
   Begin RTComctl3.CoolButton cmdQuery 
      Default         =   -1  'True
      Height          =   375
      Left            =   9135
      TabIndex        =   11
      Top             =   390
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   661
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
      MICON           =   "frmVehicleProtocol.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ListView lvVehicleProtocol 
      Height          =   4575
      Left            =   210
      TabIndex        =   8
      Top             =   1770
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   8070
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.PictureBox ptShowInfo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1410
      Left            =   0
      ScaleHeight     =   1410
      ScaleWidth      =   11355
      TabIndex        =   9
      Top             =   -90
      Width           =   11355
      Begin VB.TextBox txtLicenseTagNo 
         Height          =   315
         Left            =   2685
         TabIndex        =   1
         Top             =   405
         Width           =   1935
      End
      Begin VB.ComboBox txtProtocol 
         Height          =   300
         Left            =   5970
         TabIndex        =   7
         Top             =   900
         Width           =   2610
      End
      Begin FText.asFlatTextBox txtVehicle 
         Height          =   315
         Left            =   5970
         TabIndex        =   3
         Top             =   405
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   556
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
         Registered      =   -1  'True
      End
      Begin FText.asFlatTextBox txtCompanyID 
         Height          =   300
         Left            =   2895
         TabIndex        =   5
         Top             =   900
         Width           =   1695
         _ExtentX        =   2990
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
         Registered      =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车牌(&L):"
         Height          =   180
         Left            =   1695
         TabIndex        =   0
         Top             =   465
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "公司代码(&C):"
         Height          =   180
         Left            =   1710
         TabIndex        =   4
         Top             =   960
         Width           =   1080
      End
      Begin VB.Label lblProtocolID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "协议号(&P):"
         Height          =   180
         Left            =   4845
         TabIndex        =   6
         Top             =   960
         Width           =   900
      End
      Begin VB.Label lblVehicleID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "车辆代码(&V):"
         Height          =   180
         Left            =   4830
         TabIndex        =   2
         Top             =   465
         Width           =   1080
      End
      Begin VB.Image Image2 
         Height          =   1275
         Left            =   60
         Picture         =   "frmVehicleProtocol.frx":001C
         Top             =   150
         Width           =   2010
      End
   End
   Begin ActiveBar2LibraryCtl.ActiveBar2 abAction 
      Height          =   5145
      Left            =   8130
      TabIndex        =   10
      Top             =   1590
      Width           =   1485
      _LayoutVersion  =   1
      _ExtentX        =   2619
      _ExtentY        =   9075
      _DataPath       =   ""
      Bands           =   "frmVehicleProtocol.frx":14EF
   End
   Begin VB.Menu pmnu_Action 
      Caption         =   "操作"
      Visible         =   0   'False
      Begin VB.Menu pmnu_edit 
         Caption         =   "属性(&E)"
      End
      Begin VB.Menu pmnu_All 
         Caption         =   "全选(&A)"
      End
      Begin VB.Menu pmnu_notall 
         Caption         =   "取消全选(&N)"
      End
      Begin VB.Menu pmnu_clear 
         Caption         =   "取消协议(&C)"
      End
      Begin VB.Menu pmnu_set 
         Caption         =   "批量设置(&S)"
      End
   End
End
Attribute VB_Name = "frmVehicleProtocol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private m_oReport As New Report
Private m_oSplit As New Split
Private m_oFormular As New Formular

'界面排列
Private Sub AlignForm()
    On Error GoTo err
    ptShowInfo.Top = 0
    ptShowInfo.Left = 0
    ptShowInfo.Width = mdiMain.ScaleWidth
      
    lvVehicleProtocol.Top = ptShowInfo.Height + 50
    lvVehicleProtocol.Left = 50
    lvVehicleProtocol.Width = mdiMain.ScaleWidth - abAction.Width - 50
    lvVehicleProtocol.Height = mdiMain.ScaleHeight - ptShowInfo.Height - 50
    
    abAction.Top = lvVehicleProtocol.Top
    abAction.Left = lvVehicleProtocol.Width + 50
    abAction.Height = lvVehicleProtocol.Height
    Exit Sub
err:
    ShowErrorMsg
End Sub

Private Sub Form_Resize()
    AlignForm
End Sub

Private Sub lvVehicleProtocol_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    SortListView lvVehicleProtocol, ColumnHeader.Index
End Sub

Private Sub lvVehicleProtocol_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu pmnu_Action
    End If
End Sub

Private Sub pmnu_Add_Click()

End Sub

Private Sub pmnu_All_Click()
    Dim i As Integer
    For i = 1 To lvVehicleProtocol.ListItems.Count
        lvVehicleProtocol.ListItems.Item(i).Checked = True
    Next i
End Sub

Private Sub pmnu_clear_Click()
ChangeProtocol
End Sub

Private Sub pmnu_edit_Click()
    frmEditVehicleProtocol.ZOrder 0
    frmEditVehicleProtocol.m_AllSet = False
    frmEditVehicleProtocol.Show vbModal
End Sub

Private Sub pmnu_notall_Click()
    Dim i As Integer
    For i = 1 To lvVehicleProtocol.ListItems.Count
        lvVehicleProtocol.ListItems.Item(i).Checked = False
    Next i
End Sub

Private Sub pmnu_set_Click()
    frmEditVehicleProtocol.ZOrder 0
    frmEditVehicleProtocol.m_AllSet = True
    frmEditVehicleProtocol.Show vbModal
End Sub

Private Sub txtVehicle_ButtonClick()
    On Error GoTo err
    Dim oShell As New STShell.CommDialog
    Dim aszTemp() As String
    oShell.Init g_oActiveUser
    aszTemp = oShell.SelectVehicle()
    Set oShell = Nothing
    If ArrayLength(aszTemp) = 0 Then Exit Sub
    txtVehicle.Text = MakeDisplayString(Trim(aszTemp(1, 1)), Trim(aszTemp(1, 2)))
    Exit Sub
err:
ShowErrorMsg
End Sub

Private Sub asFlatTextBox2_Change()

End Sub

Private Sub cmdQuery_Click()
    Query
End Sub
Private Sub Query()
    On Error GoTo err
    Dim nCount As Integer
    Dim aszTemp() As String
    Dim i As Integer
    Dim lvTemp As ListItem
    aszTemp = m_oReport.GetVehicleProtocol(ResolveDisplay(txtVehicle.Text), ResolveDisplay(txtCompanyID.Text), ResolveDisplay(txtProtocol.Text), txtLicenseTagNo.Text)
    lvVehicleProtocol.ListItems.Clear
    nCount = ArrayLength(aszTemp)
    If ArrayLength(aszTemp) <> 0 Then
        
        For i = 1 To ArrayLength(aszTemp)
            Set lvTemp = lvVehicleProtocol.ListItems.Add(, , aszTemp(i, 3))
            lvTemp.ListSubItems.Add , , aszTemp(i, 4)
            lvTemp.ListSubItems.Add , , MakeDisplayString(aszTemp(i, 5), aszTemp(i, 6))
            lvTemp.ListSubItems.Add , , MakeDisplayString(aszTemp(i, 1), aszTemp(i, 2))
            lvTemp.ListSubItems.Add , , IIf(aszTemp(i, 1) = "", "未设置协议", GetDefaultMark(aszTemp(i, 7)))
        Next i
    End If
    SetNormal
    lvVehicleProtocol.Refresh
    If lvVehicleProtocol.ListItems.Count > 0 Then lvVehicleProtocol.ListItems(1).Selected = True
    WriteProcessBar False
    ShowSBInfo "共" & nCount & "个对象", ESB_ResultCountInfo
    ShowSBInfo ""

    Exit Sub
err:
ShowErrorMsg
End Sub

Private Sub Form_Load()
    m_oReport.Init g_oActiveUser
    AlignForm
    FillHead
    FillLvVehicleProtocol
    m_oFormular.Init g_oActiveUser
    AlignHeadWidth Me.name, lvVehicleProtocol
    SortListView lvVehicleProtocol, 1
End Sub
Public Sub FillLvVehicleProtocol(Optional pszVehicleID As String = "")
    On Error GoTo err
    Dim lvTemp As ListItem
    Dim aszTemp() As String, i As Integer
    Dim nCount As Integer
    Dim j As Integer
    m_oReport.Init g_oActiveUser
    aszTemp = m_oReport.GetVehicleProtocol(pszVehicleID)
    If pszVehicleID = "" Then
        lvVehicleProtocol.ListItems.Clear
    End If
    nCount = ArrayLength(aszTemp)
    If nCount <> 0 Then
        For i = 1 To nCount
            For j = lvVehicleProtocol.ListItems.Count To 1 Step -1
                If aszTemp(i, 3) = lvVehicleProtocol.ListItems(j).Text Then
                    lvVehicleProtocol.ListItems.Remove j
                End If
            Next j
            Set lvTemp = lvVehicleProtocol.ListItems.Add(, , aszTemp(i, 3))
            lvTemp.ListSubItems.Add , , aszTemp(i, 4)
            lvTemp.ListSubItems.Add , , MakeDisplayString(aszTemp(i, 5), aszTemp(i, 6))
            lvTemp.ListSubItems.Add , , MakeDisplayString(aszTemp(i, 1), aszTemp(i, 2))
            If aszTemp(i, 7) = "" Then
                lvTemp.ListSubItems.Add , , IIf(aszTemp(i, 1) = "", "未设置协议", GetDefaultMark(0))
            Else
                lvTemp.ListSubItems.Add , , IIf(aszTemp(i, 1) = "", "未设置协议", GetDefaultMark(aszTemp(i, 7)))
            End If
            If i = nCount Then
                lvTemp.Selected = True
                lvTemp.EnsureVisible
            End If
        Next i
    End If
    SetNormal
    lvVehicleProtocol.Refresh
    If lvVehicleProtocol.ListItems.Count > 0 And pszVehicleID = "" Then
        lvVehicleProtocol.ListItems(1).Selected = True
        lvVehicleProtocol.ListItems(1).EnsureVisible
    End If
    WriteProcessBar False
    ShowSBInfo "共" & nCount & "个对象", ESB_ResultCountInfo
    ShowSBInfo ""
    Exit Sub
err:
    ShowErrorMsg
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveHeadWidth Me.name, lvVehicleProtocol
    Unload Me
End Sub

Private Sub FillHead()
    Dim aszTemp() As String
    Dim i As Integer
    With lvVehicleProtocol.ColumnHeaders
        .Clear
        .Add , , "车辆代码"
        .Add , , "车牌号"
'        .Add , , "公司代码"
        .Add , , "参运公司"
'        .Add , , "协议号"
        .Add , , "协议名称"
        .Add , , "是否默认协议"
    End With
    With lvVehicleProtocol
        .ColumnHeaders(1).Width = 1000
        .ColumnHeaders(2).Width = 1000
        .ColumnHeaders(3).Width = 1500
        .ColumnHeaders(4).Width = 2000
        .ColumnHeaders(5).Width = 1500
    End With

    aszTemp = m_oReport.GetAllProtocol
    If ArrayLength(aszTemp) <> 0 Then
        For i = 1 To ArrayLength(aszTemp)
            txtProtocol.AddItem MakeDisplayString(aszTemp(i, 1), aszTemp(i, 2))
        Next i
    End If
End Sub

Private Sub abAction_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
On Error GoTo here
    Dim i As Integer
    If lvVehicleProtocol.ListItems.Count = 0 Then Exit Sub
    Select Case Tool.Caption
        Case "属性"
            frmEditVehicleProtocol.ZOrder 0
            frmEditVehicleProtocol.m_AllSet = False
            frmEditVehicleProtocol.Show vbModal
        Case "全选"
            For i = 1 To lvVehicleProtocol.ListItems.Count
                lvVehicleProtocol.ListItems.Item(i).Checked = True
            Next i
        Case "取消全选"
            For i = 1 To lvVehicleProtocol.ListItems.Count
                lvVehicleProtocol.ListItems.Item(i).Checked = False
            Next i
        Case "取消协议"
            ChangeProtocol
        Case "批量设置"
            
            frmEditVehicleProtocol.m_AllSet = True
            frmEditVehicleProtocol.ZOrder 0
            frmEditVehicleProtocol.Show vbModal
    End Select
    
    Exit Sub
here:
ShowErrorMsg
End Sub
Private Sub ChangeProtocol()
    On Error GoTo err
    Dim i As Integer, j As Integer
    Dim aszTemp() As String
    Dim aszTemp1() As String
    Dim szProtocol As String
    Dim vbYesOrNo As VbMsgBoxResult
    vbYesOrNo = MsgBox("取消协议：是将所选车辆的协议设为默认协议，是否继续?", vbQuestion + vbYesNo + vbDefaultButton2, "协议管理")
    If vbYesOrNo = vbYes Then
        For i = 1 To lvVehicleProtocol.ListItems.Count
            If lvVehicleProtocol.ListItems.Item(i).Checked = True Then
                j = j + 1
            End If
        Next i
        If j = 0 Then
            MsgBox "你没有选择要取消协议的车辆，请在选择后再取消协议", vbInformation, "协议管理"
            Exit Sub
        End If
        ReDim aszTemp(1 To j)
        j = 0
        For i = 1 To lvVehicleProtocol.ListItems.Count
            If lvVehicleProtocol.ListItems.Item(i).Checked = True Then
                j = j + 1
                aszTemp(j) = lvVehicleProtocol.ListItems.Item(i).Text
            End If
        Next i
        j = 0
        m_oSplit.Init g_oActiveUser
            
        m_oReport.Init g_oActiveUser
    
        aszTemp1 = m_oReport.GetAllProtocol
        For i = 1 To ArrayLength(aszTemp1)
            If aszTemp1(i, 4) = 0 Then
                szProtocol = aszTemp1(i, 1)
            End If
        Next i
        m_oSplit.SetVehicleProtocol aszTemp, szProtocol
        FillLvVehicleProtocol
    End If
    Exit Sub
err:
ShowErrorMsg
End Sub
Private Sub lblProtocolName_Click()

End Sub

Private Sub lvVehicleProtocol_DblClick()
    Dim lvTemp As ListItem
    Dim aszTemp() As String
    Dim nCount As Integer
    Dim i As Integer
    If lvVehicleProtocol.SelectedItem Is Nothing Then Exit Sub
    
    frmEditVehicleProtocol.m_AllSet = False
    frmEditVehicleProtocol.Show vbModal
    
    aszTemp = m_oReport.GetVehicleProtocol(lvVehicleProtocol.SelectedItem.Text)
    nCount = ArrayLength(aszTemp)
    If ArrayLength(aszTemp) > 0 Then
        
'        For i = 1 To ArrayLength(aszTemp)
            Set lvTemp = lvVehicleProtocol.SelectedItem
            lvTemp.SubItems(1) = aszTemp(1, 4)
            lvTemp.SubItems(2) = MakeDisplayString(aszTemp(1, 5), aszTemp(1, 6))
            lvTemp.SubItems(3) = MakeDisplayString(aszTemp(1, 1), aszTemp(1, 2))
            lvTemp.SubItems(4) = IIf(aszTemp(1, 1) = "", "未设置协议", GetDefaultMark(aszTemp(1, 7)))
'        Next i
    End If
    SetNormal

    
    
End Sub

Private Sub txtCompanyID_ButtonClick()
        Dim oShell As New STShell.CommDialog
    Dim aszTemp() As String
    oShell.Init g_oActiveUser
    aszTemp = oShell.SelectCompany()
    Set oShell = Nothing
    If ArrayLength(aszTemp) = 0 Then Exit Sub
    txtCompanyID.Text = MakeDisplayString(Trim(aszTemp(1, 1)), Trim(aszTemp(1, 2)))

End Sub

