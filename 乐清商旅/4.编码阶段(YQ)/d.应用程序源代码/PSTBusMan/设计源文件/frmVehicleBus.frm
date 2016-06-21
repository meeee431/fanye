VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Begin VB.Form frmVehicleBus 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "车辆--运行车次"
   ClientHeight    =   4545
   ClientLeft      =   2550
   ClientTop       =   2490
   ClientWidth     =   7170
   HelpContextID   =   2008401
   Icon            =   "frmVehicleBus.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   7170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Timer tmStart 
      Interval        =   50
      Left            =   2265
      Top             =   2505
   End
   Begin RTComctl3.CoolButton cmdHelp 
      Height          =   315
      Left            =   5835
      TabIndex        =   11
      Top             =   4125
      Width           =   1140
      _ExtentX        =   0
      _ExtentY        =   0
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
      MICON           =   "frmVehicleBus.frx":014A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.CheckBox chkStop 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "车次车辆停班(&S)"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   135
      TabIndex        =   4
      Top             =   3300
      Width           =   2580
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2565
      Top             =   2745
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVehicleBus.frx":0166
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVehicleBus.frx":02C2
            Key             =   "Run"
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.DTPicker dtpStartDate 
      Height          =   315
      Left            =   1605
      TabIndex        =   6
      Top             =   3570
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   64946176
      CurrentDate     =   36453
   End
   Begin RTComctl3.CoolButton cmdOk 
      Default         =   -1  'True
      Height          =   315
      Left            =   3465
      TabIndex        =   9
      Top             =   4125
      Width           =   1140
      _ExtentX        =   0
      _ExtentY        =   0
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
      MICON           =   "frmVehicleBus.frx":041E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ListView lvBus 
      Height          =   1860
      Left            =   105
      TabIndex        =   3
      Top             =   1290
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   3281
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "车次代码"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "发车时间"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "运行线路"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "车次类型"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "循环周期"
         Object.Width           =   2540
      EndProperty
   End
   Begin RTComctl3.CoolButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   4650
      TabIndex        =   10
      Top             =   4125
      Width           =   1140
      _ExtentX        =   0
      _ExtentY        =   0
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
      MICON           =   "frmVehicleBus.frx":043A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker dtpEndDate 
      Height          =   315
      Left            =   5145
      TabIndex        =   8
      Top             =   3570
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   64946176
      CurrentDate     =   36453
   End
   Begin FText.asFlatTextBox txtVehicleId 
      Height          =   300
      Left            =   1215
      TabIndex        =   1
      Top             =   90
      Width           =   2280
      _ExtentX        =   4022
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
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   135
      X2              =   7050
      Y1              =   4020
      Y2              =   4020
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      X1              =   135
      X2              =   7050
      Y1              =   4005
      Y2              =   4005
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   1260
      X2              =   7065
      Y1              =   1125
      Y2              =   1125
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   1260
      X2              =   7065
      Y1              =   1110
      Y2              =   1110
   End
   Begin VB.Label lblEndDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "结束日期(&K):"
      Enabled         =   0   'False
      Height          =   180
      Left            =   3960
      TabIndex        =   7
      Top             =   3630
      Width           =   1080
   End
   Begin VB.Label lblStartDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "开始日期(&K):"
      Enabled         =   0   'False
      Height          =   180
      Left            =   420
      TabIndex        =   5
      Top             =   3630
      Width           =   1080
   End
   Begin VB.Label lblCompany 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label12"
      Height          =   180
      Left            =   5370
      TabIndex        =   21
      Top             =   510
      Width           =   630
   End
   Begin VB.Label lblOwner 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label11"
      Height          =   180
      Left            =   1680
      TabIndex        =   20
      Top             =   780
      Width           =   630
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车主:"
      Height          =   180
      Left            =   1215
      TabIndex        =   19
      Top             =   780
      Width           =   450
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "参运公司:"
      Height          =   180
      Left            =   4485
      TabIndex        =   18
      Top             =   510
      Width           =   810
   End
   Begin VB.Label lblSeatCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      Height          =   180
      Left            =   3600
      TabIndex        =   17
      Top             =   510
      Width           =   540
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label7"
      Height          =   180
      Left            =   3600
      TabIndex        =   16
      Top             =   780
      Width           =   540
   End
   Begin VB.Label lblVehicleModel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      Height          =   180
      Left            =   1680
      TabIndex        =   15
      Top             =   510
      Width           =   540
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "运行车次(&B):"
      Height          =   180
      Left            =   120
      TabIndex        =   2
      Top             =   1035
      Width           =   1080
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "状态:"
      Height          =   180
      Left            =   2940
      TabIndex        =   14
      Top             =   780
      Width           =   450
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "座位数:"
      Height          =   180
      Left            =   2940
      TabIndex        =   13
      Top             =   510
      Width           =   630
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车型:"
      Height          =   180
      Left            =   1215
      TabIndex        =   12
      Top             =   510
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车辆代码(&I):"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   165
      Width           =   1080
   End
End
Attribute VB_Name = "frmVehicleBus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************
'* Source File Name:frmVehicleBus.frm
'* Project Name:RTBusMan
'* Engineer:陈峰
'* Data Generated:2002/08/27
'* Last Revision Date:2002/09/12
'* Brief Description:车辆车次信息
'* Relational Document:
'**********************************************************

Option Explicit
Public m_bIsParent As Boolean

Private m_oVehicle As Vehicle
Private m_szVehicleId As String
Private m_oBus As New Bus
Private m_oRegularScheme As New RegularScheme



Private Sub chkStop_Click()
    If ChkStop.Value = 1 Then
        dtpStartDate.Enabled = True
        dtpEndDate.Enabled = True
        cmdOK.Enabled = True
        lblStartDate.Enabled = True
        lblEndDate.Enabled = True
    Else
        dtpStartDate.Enabled = False
        dtpEndDate.Enabled = False
        cmdOK.Enabled = False
        lblStartDate.Enabled = False
        lblEndDate.Enabled = False
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
DisplayHelp Me
End Sub

Private Sub cmdOk_Click()
    Dim bMsg As Integer
    Dim nCount As Integer
    Dim szBus_id As String
    Dim i As Integer

    bMsg = MsgBox("是否将该车辆车次停班", vbExclamation + vbYesNo + vbDefaultButton2, "车辆停班")
    If bMsg = vbYes Then


         m_oVehicle.AllBusStop g_szExePriceTable, dtpStartDate.Value, dtpEndDate.Value
         If frmVehicleBus.lvBus.ListItems.Count > 0 Then
            nCount = frmVehicleBus.lvBus.ListItems.Count

            For i = 1 To nCount
'                If frmVehicleBus.lvBus.ListItems(i).Selected = True Then
                    szBus_id = frmVehicleBus.lvBus.ListItems(i)
'                    m_oVehicle.AllBusStopEx m_szProjectID, dtpStartDate.Value, dtpEndDate.Value, szBus_id
                    If m_bIsParent = True Then
                        frmBus.UpdateList (szBus_id)
                    End If

'                End If
            Next

         End If
            MsgBox "车次车辆停班完成!", vbInformation, "车辆停班"
       cmdOK.Enabled = False
       tmStart.Enabled = True
      End If
'
End Sub

Private Sub dtpEndDate_Change()
    cmdOK.Enabled = True
End Sub

Private Sub dtpStartDate_Change()
    cmdOK.Enabled = True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
           Case vbKeyEscape
           Unload Me
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    m_bShow = False
End Sub

Private Sub lvbus_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static m_nUpColumn As Integer
    If lvBus.SelectedItem Is Nothing Then Exit Sub
    lvBus.SortKey = ColumnHeader.Index - 1
    If m_nUpColumn = ColumnHeader.Index - 1 Then
        lvBus.SortOrder = lvwDescending
        m_nUpColumn = ColumnHeader.Index
    Else
        lvBus.SortOrder = lvwAscending
        m_nUpColumn = ColumnHeader.Index - 1
    End If
    lvBus.Sorted = True
End Sub

Private Sub lvBus_DblClick()
'    Dim m_oBus As New Bus
'    If frmArrangeBus.m_bShow Then Exit Sub
'    If lvBus.SelectedItem Is Nothing Then Exit Sub
'    g_szBusID = lvBus.SelectedItem.Text
'    m_oBus.Init g_oActiveUser
'    Set frmArrangeBus.m_oBus = m_oBus
'    frmArrangeBus.Show vbModal
End Sub

Public Sub Init(vData As Vehicle)
Dim tSCheme As TSchemeArrangement
On Error GoTo ErrorHandle
    Set m_oVehicle = vData
    m_szVehicleId = m_oVehicle.VehicleId
    lblCompany.Caption = m_oVehicle.CompanyName
    lblOwner.Caption = m_oVehicle.OwnerName
    lblSeatCount.Caption = m_oVehicle.SeatCount
    If m_oVehicle.Status = ST_VehicleRun Then
        lblStatus.Caption = "正常"
    Else
        lblStatus.Caption = "停班"
    End If
    lblVehicleModel.Caption = m_oVehicle.VehicleModelName
Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub tmStart_Timer()
    Dim szaBus() As String
    Dim liTemp As ListItem
    Dim nCount As Integer, i As Integer
    On Error GoTo ErrorHandle
    lvBus.ListItems.Clear
    tmStart.Enabled = False
    txtVehicleId.Enabled = False
    SetBusy
    m_oBus.Init g_oActiveUser
    ChkStop.Caption = "车次车辆停班(&S)"
    szaBus = m_oVehicle.GetAllBus(g_szExePriceTable)
    nCount = ArrayLength(szaBus)
    txtVehicleId.Text = m_oVehicle.VehicleId
    If nCount = 0 Then
        SetNormal
        ChkStop.Enabled = False
        cmdOK.Enabled = False
        Exit Sub
    End If
    WriteProcessBar , , nCount, "获得车次..."
    For i = 1 To nCount
        WriteProcessBar , i, nCount, "获得车次[" & Trim(szaBus(i)) & "]"
        Set liTemp = lvBus.ListItems.Add(, , Trim(szaBus(i)), , "Run")
        m_oBus.Identify szaBus(i)
       '
        If DateDiff("d", CDate(m_oBus.EndStopDate), Now) <= 0 Then
                liTemp.SmallIcon = "Stop"
        Else
                liTemp.SmallIcon = "Run"
        End If
        liTemp.SubItems(1) = Format(m_oBus.StartUpTime, "HH:mm")
        liTemp.SubItems(2) = Trim(m_oBus.RouteName)
        If m_oBus.BusType = TP_ScrollBus Then
            liTemp.SubItems(3) = "滚动"
        Else
            liTemp.SubItems(3) = "固定"
        End If
        liTemp.SubItems(4) = m_oBus.RunCycle
    Next
    If lvBus.ListItems.Count = 0 Then ChkStop.Enabled = False
    SetNormal
    tmStart.Enabled = False
    dtpStartDate.Value = Date
    dtpEndDate.Value = Date
    WriteProcessBar False
Exit Sub
ErrorHandle:
    SetNormal
    ShowErrorMsg
End Sub


Private Sub txtVehicleId_ButtonClick()
    Dim oShell As New CommDialog
    oShell.Init g_oActiveUser
    Dim aszTmp() As String
    aszTmp = oShell.SelectVehicleEX(False)
    If ArrayLength(aszTmp) = 0 Then Exit Sub
    txtVehicleId.Text = MakeDisplayString(aszTmp(1, 1), aszTmp(1, 2))

End Sub


