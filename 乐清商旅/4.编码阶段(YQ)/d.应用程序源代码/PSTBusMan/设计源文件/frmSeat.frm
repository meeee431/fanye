VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Begin VB.Form frmSeat 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "环境--新增座位"
   ClientHeight    =   1455
   ClientLeft      =   2715
   ClientTop       =   4485
   ClientWidth     =   4365
   HelpContextID   =   2007201
   Icon            =   "frmSeat.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox cmbBusType 
      Height          =   300
      ItemData        =   "frmSeat.frx":014A
      Left            =   1575
      List            =   "frmSeat.frx":014C
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   990
      Width           =   1335
   End
   Begin RTComctl3.CoolButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   3135
      TabIndex        =   7
      Top             =   495
      Width           =   1125
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
      MICON           =   "frmSeat.frx":014E
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
      Left            =   3135
      TabIndex        =   6
      Top             =   120
      Width           =   1125
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "确定(&O)"
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
      MICON           =   "frmSeat.frx":016A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.UpDown UpDown2 
      Height          =   285
      Left            =   2640
      TabIndex        =   9
      Top             =   510
      Width           =   270
      _ExtentX        =   423
      _ExtentY        =   503
      _Version        =   393216
      BuddyControl    =   "txtEndSeat"
      BuddyDispid     =   196614
      OrigLeft        =   2475
      OrigTop         =   660
      OrigRight       =   2745
      OrigBottom      =   930
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   285
      Left            =   2640
      TabIndex        =   8
      Top             =   120
      Width           =   270
      _ExtentX        =   423
      _ExtentY        =   503
      _Version        =   393216
      BuddyControl    =   "txtStartSeat"
      BuddyDispid     =   196615
      OrigLeft        =   2415
      OrigTop         =   135
      OrigRight       =   2685
      OrigBottom      =   435
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtEndSeat 
      Height          =   270
      Left            =   1575
      TabIndex        =   3
      Text            =   "0"
      Top             =   510
      Width           =   1080
   End
   Begin VB.TextBox txtStartSeat 
      Height          =   270
      Left            =   1575
      TabIndex        =   1
      Text            =   "0"
      Top             =   120
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "座位类型(&T):"
      Height          =   180
      Left            =   270
      TabIndex        =   4
      Top             =   1035
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "结束座位号(&E):"
      Height          =   180
      Left            =   225
      TabIndex        =   2
      Top             =   540
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "开始座位号(&S):"
      Height          =   180
      Left            =   225
      TabIndex        =   0
      Top             =   180
      Width           =   1260
   End
End
Attribute VB_Name = "frmSeat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************
'* Source File Name:frmSeat.frm
'* Project Name:RTBusMan
'* Engineer:
'* Data Generated:2002/08/27
'* Last Revision Date:2002/09/03
'* Brief Description:新增座位
'* Relational Document:UI_BS_SM_34.DOC
'**********************************************************
Option Explicit
Public Enum eSeat
    AddSeat = 1
    DeleteSeat = 2
    ReserveSeat = 3
    UnReserveSeat = 4
    AddSeatType = 5
End Enum
Public m_oReBus As REBus
Public efrmSeat As eSeat
Public m_SeatType As Variant '由frmResSeat传入
Public m_busVehicleModal As String
Public m_szSeatNo As String

Public m_bIsParent As Boolean '是否是父窗体调用



Private Sub cmdCancel_Click()
Unload Me
End Sub
Private Sub cmdOk_Click()
    Dim i As Integer
    Dim szSeatNo As String
    Dim szModifySeatType(1 To 3) As String
    Dim oReSheme As New REScheme
    On Error GoTo ErrorHandle
    SetBusy
    oReSheme.Init g_oActiveUser
    Select Case efrmSeat
    Case 1
        ShowSBInfo "正在新增座位..."
        m_oReBus.AddSeat Val(txtEndSeat.Text) - Val(txtStartSeat.Text) + 1, txtStartSeat.Text, ResolveDisplay(cmbBusType.Text)
        ShowSBInfo "新增座位完成"
        frmEnvReserveSeat.FullSeat
        If m_bIsParent = True Then
            frmEnvBus.UpdateList m_oReBus.BusID, m_oReBus.RunDate
        End If
    Case 2
        m_oReBus.DeleteSeat Val(txtEndSeat.Text) - Val(txtStartSeat.Text), txtStartSeat.Text
    Case 3
        For i = Val(txtStartSeat.Text) To Val(txtEndSeat.Text) - Val(txtStartSeat.Text)
            szSeatNo = szSeatNo & "'" & Format(i, "00") & "',"
        Next
        szSeatNo = Left(szSeatNo, Len(szSeatNo) - 1)
        m_oReBus.ReserveSeat szSeatNo
    Case 4
        For i = Val(txtStartSeat.Text) To Val(txtEndSeat.Text) - Val(txtStartSeat.Text)
            szSeatNo = szSeatNo & "'" & Format(i, "00") & "',"
        Next
        szSeatNo = Left(szSeatNo, Len(szSeatNo) - 1)
        m_oReBus.UnReserveSeat Format(i, "00")
    Case 5
        '修改座位
        szModifySeatType(1) = ResolveDisplay(cmbBusType.Text)
        If CInt(Val(txtStartSeat.Text)) <= CInt(Val(txtEndSeat.Text)) Then
            szModifySeatType(2) = Format(Val(txtStartSeat.Text), "00")
            szModifySeatType(3) = Format(Val(txtEndSeat.Text), "00")
        Else
            szModifySeatType(2) = Format(Val(txtEndSeat.Text), "00")
            szModifySeatType(3) = Format(Val(txtStartSeat.Text), "00")
        End If
        If FindSeatNoSell(szModifySeatType) Then    '无售票
            m_oReBus.ReModifySeatType szModifySeatType
            '           oResheme.MakeRunEvironment m_oREBus.RunDate, m_oREBus.BusID, True
        Else
            MsgBox "只有可售座位或预留座位，才可修改座位类型", vbInformation, "提示"
        End If
    End Select
    SetNormal
    Unload Me
    Exit Sub
ErrorHandle:
    SetNormal
    ShowErrorMsg
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    UpDown1.Max = 100
    UpDown2.Max = 100
    If m_oReBus.BusID = "" Then
        txtEndSeat.Enabled = False
        txtStartSeat.Enabled = False
        UpDown1.Enabled = False
        UpDown2.Enabled = False
        Label1.Enabled = False
        Label2.Enabled = False
    End If
    
    
    Select Case efrmSeat
    Case 1
        AddComb
        frmSeat.Caption = "新增座位"
        txtStartSeat.Text = Format(Val(m_szSeatNo) + 1, "00")
        txtEndSeat.Text = Format(Val(m_szSeatNo) + 1, "00")
    Case 2
        frmSeat.Caption = "删除座位"
    Case 3
        frmSeat.Caption = "预留座位"
    Case 4
        frmSeat.Caption = "取消预留"
    Case 5
        frmSeat.Height = 1815
        frmSeat.Caption = "修改座位类型"
        If m_szSeatNo <> "" Then
'            m_SeatType = m_oReBus.GetReBusSeatType
            txtStartSeat.Text = m_szSeatNo
            txtEndSeat.Text = m_szSeatNo
        Else
            txtStartSeat.Text = Format(1, "00")
            txtEndSeat.Text = Format(1, "00")
        End If
        frmSeat.cmdOk.Enabled = True
        AddComb
    End Select
    'cmdOk.Enabled = False
    cmdOk.Enabled = True
End Sub

Private Sub txtEndSeat_Change()
    If Val(txtEndSeat.Text) >= Val(txtStartSeat.Text) Then
        cmdOk.Enabled = True
    Else
        cmdOk.Enabled = False
    End If
End Sub

Private Sub txtEndSeat_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
       Case vbKeyReturn
       cmdOk_Click
End Select
End Sub

Private Sub txtStartSeat_Change()
    If Val(txtEndSeat.Text) >= Val(txtStartSeat.Text) Then
        cmdOk.Enabled = True
    Else
        cmdOk.Enabled = False
    End If
End Sub
Public Sub AddComb()
   Dim i As Integer
   Dim nCount As Integer
   nCount = ArrayLength(m_SeatType)
  ' cmbBusType.AddItem cszSeatTypeIsNormal & "[普通]"
   For i = 1 To nCount
   ' If Trim(m_SeatType(i, 1)) = cszSeatTypeIsNormal Then GoTo nextHere
    cmbBusType.AddItem Trim(m_SeatType(i, 1)) & "[" & Trim(m_SeatType(i, 2)) & "]"
'nextHere:
   Next
   cmbBusType.ListIndex = 0
   cmdOk.Enabled = True
End Sub
Private Function FindSeatNoSell(szSeatNo() As String) As Boolean
 Dim i As Integer, j As Integer
 Dim nCount As Integer
  nCount = frmEnvReserveSeat.lvSeat.ListItems.Count

  For j = CInt(szSeatNo(2)) To szSeatNo(3)
    For i = 1 To nCount
        If frmEnvReserveSeat.lvSeat.ListItems(i).SubItems(2) = "已售" Then
           If Trim(frmEnvReserveSeat.lvSeat.ListItems(i).Text) = Format(j, "00") Then FindSeatNoSell = False: Exit Function
        End If
        If frmEnvReserveSeat.lvSeat.ListItems(i).SubItems(2) = "可售" Or frmEnvReserveSeat.lvSeat.ListItems(i).SubItems(2) = "预留" Then
           If Trim(frmEnvReserveSeat.lvSeat.ListItems(i).Text) = Format(j, "00") Then FindSeatNoSell = True
        End If
    Next
  Next
End Function
