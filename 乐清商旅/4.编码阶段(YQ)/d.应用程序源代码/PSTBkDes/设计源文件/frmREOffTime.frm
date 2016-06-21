VERSION 5.00
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmREBusAttr 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "环境--车次属性"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6960
   HelpContextID   =   2005201
   Icon            =   "frmREOffTime.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   Begin RTComctl3.CoolButton cmdOther 
      Height          =   315
      Left            =   5700
      TabIndex        =   28
      Top             =   809
      Width           =   1215
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "更多..."
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
      MICON           =   "frmREOffTime.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.OptionButton optNoR 
      BackColor       =   &H00E0E0E0&
      Caption         =   "不全额退票(&N)"
      Height          =   195
      Left            =   2310
      TabIndex        =   15
      Top             =   2895
      Width           =   1530
   End
   Begin VB.OptionButton optAllRe 
      BackColor       =   &H00E0E0E0&
      Caption         =   "全额退票(&R)"
      Height          =   225
      Left            =   780
      TabIndex        =   13
      Top             =   2865
      Value           =   -1  'True
      Width           =   1515
   End
   Begin RTComctl3.TextButtonBox txtCheckGate 
      Height          =   300
      Left            =   1890
      TabIndex        =   12
      Top             =   2220
      Width           =   1350
      _ExtentX        =   2381
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
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2880
      Top             =   645
   End
   Begin RTComctl3.CoolButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   5700
      TabIndex        =   3
      Top             =   457
      Width           =   1215
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "关闭"
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
      MICON           =   "frmREOffTime.frx":0028
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
      Left            =   5700
      TabIndex        =   2
      Top             =   105
      Width           =   1215
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
      MICON           =   "frmREOffTime.frx":0044
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.TextButtonBox txtBusID 
      Height          =   300
      Left            =   1890
      TabIndex        =   1
      Top             =   120
      Width           =   1350
      _ExtentX        =   2381
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
   End
   Begin MSComCtl2.DTPicker dtpNewTime 
      Height          =   300
      Left            =   1890
      TabIndex        =   22
      Top             =   1575
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   529
      _Version        =   393216
      Format          =   71106562
      CurrentDate     =   36397
   End
   Begin VB.Label lblCompany 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "参运公司:"
      Height          =   180
      Left            =   780
      TabIndex        =   30
      Top             =   3870
      Width           =   810
   End
   Begin VB.Label lblOwner 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车主:"
      Height          =   180
      Left            =   780
      TabIndex        =   29
      Top             =   3495
      Width           =   450
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "其他信息"
      Height          =   180
      Left            =   120
      TabIndex        =   27
      Top             =   3210
      Width           =   720
   End
   Begin VB.Line Line8 
      X1              =   885
      X2              =   5485
      Y1              =   3315
      Y2              =   3315
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FFFFFF&
      X1              =   885
      X2              =   5485
      Y1              =   3330
      Y2              =   3330
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FFFFFF&
      X1              =   945
      X2              =   5545
      Y1              =   2685
      Y2              =   2685
   End
   Begin VB.Line Line5 
      X1              =   945
      X2              =   5545
      Y1              =   2670
      Y2              =   2670
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   945
      X2              =   5545
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line3 
      X1              =   945
      X2              =   5545
      Y1              =   2025
      Y2              =   2025
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   945
      X2              =   5545
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line1 
      X1              =   945
      X2              =   5545
      Y1              =   1425
      Y2              =   1425
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "发车时间"
      Height          =   180
      Left            =   120
      TabIndex        =   26
      Top             =   1350
      Width           =   720
   End
   Begin VB.Label lblCurrentTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "现在时间:"
      Height          =   180
      Left            =   3570
      TabIndex        =   25
      Top             =   1635
      Width           =   810
   End
   Begin VB.Label lblNow 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   180
      Left            =   4395
      TabIndex        =   24
      Top             =   1635
      Width           =   90
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "发车时间(&T):"
      Height          =   180
      Left            =   780
      TabIndex        =   23
      Top             =   1635
      Width           =   1080
   End
   Begin VB.Label lblRoute 
      BackStyle       =   0  'Transparent
      Height          =   165
      Left            =   4425
      TabIndex        =   21
      Top             =   780
      Width           =   645
   End
   Begin VB.Label lblSellSeat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   180
      Left            =   4425
      TabIndex        =   20
      Top             =   480
      Width           =   90
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "已售座数:"
      Height          =   180
      Left            =   3600
      TabIndex        =   19
      Top             =   480
      Width           =   810
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "运行线路:"
      Height          =   180
      Left            =   3600
      TabIndex        =   18
      Top             =   765
      Width           =   810
   End
   Begin VB.Label lblAllRefundment 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "全额退票:"
      Height          =   180
      Left            =   3600
      TabIndex        =   17
      Top             =   1035
      Width           =   810
   End
   Begin VB.Label lblCheckGate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "检票口:"
      Height          =   180
      Left            =   3600
      TabIndex        =   16
      Top             =   2280
      Width           =   630
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "新检票口(&C):"
      Height          =   180
      Left            =   780
      TabIndex        =   14
      Top             =   2265
      Width           =   1080
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "全额退票"
      Height          =   180
      Left            =   120
      TabIndex        =   11
      Top             =   2580
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "检票设定"
      Height          =   180
      Left            =   120
      TabIndex        =   10
      Top             =   1935
      Width           =   720
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车次状态:"
      Height          =   180
      Left            =   780
      TabIndex        =   9
      Top             =   1035
      Width           =   810
   End
   Begin VB.Label lblOffTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   180
      Left            =   1605
      TabIndex        =   8
      Top             =   480
      Width           =   90
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "发车时间:"
      Height          =   180
      Left            =   780
      TabIndex        =   7
      Top             =   480
      Width           =   810
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "运行车辆:"
      Height          =   180
      Left            =   780
      TabIndex        =   6
      Top             =   765
      Width           =   810
   End
   Begin VB.Label lblVehicle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   180
      Left            =   1620
      TabIndex        =   5
      Top             =   765
      Width           =   90
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车次代码(&I):"
      Height          =   180
      Left            =   810
      TabIndex        =   0
      Top             =   180
      Width           =   1080
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0001"
      Height          =   180
      Left            =   3570
      TabIndex        =   4
      Top             =   180
      Width           =   360
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   120
      Picture         =   "frmREOffTime.frx":0060
      Stretch         =   -1  'True
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmREBusAttr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************
'                   TOP GROUP INC.
'* Copyright(C)1999 TOP GROUP INC.
'*
'* All rights reserved.No part of this program or publication
'* may be reproduced,transmitted,transcribed,stored in a
'* retrieval system,or translated intoany language or compute
'* language,in any form or by any means,electronic,mechanical,
'* magnetic,optical,chemical,biological,or otherwise,without
'* the prior written permission.
'*********************************************************
'
'**********************************************************
'* Source File Name:frmREOffTime.frm
'* Project Name:StationNet 2.0
'* Engineer:魏宏旭
'* Data Generated:1999/8/28
'* Last Revision Date:1999/9/2
'* Brief Description:修改发车时间
'* Relational Document:UI_BS_SM_28.DOC
'**********************************************************
Public m_szBusID As String
Public m_dtBusDate As Date
Private m_oREBus As New REBus

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCompany_Click()
    frmCorp.Status = SNModify
    frmCorp.szCompanyID = m_oREBus.Company
    frmCorp.Show vbModal
End Sub

Private Sub cmdOk_Click()
Dim szTemp As String
On Error GoTo here
    m_oREBus.Identify txtBusID.Text, m_dtBusDate
    If optNoR.Value Then
        m_oREBus.AllRefundment = False
        szTemp = "否"
    Else
        m_oREBus.AllRefundment = True
        szTemp = "是"
    End If
    m_oREBus.CheckGate = GetLString(txtCheckGate.Text)
    m_oREBus.StartupTime = dtpNewTime.Value
    m_oREBus.Update
    If frmREBus.bIsShow Then
       frmREBus.UpList txtBusID.Text, m_dtBusDate
    End If
    MsgBox "车次[" & Trim(txtBusID.Text) & "]属性修改成功", vbInformation + vbOKOnly, "环境"
    cmdOk.Enabled = False
Exit Sub
here:
    ShowErrorMsg
End Sub

Private Sub cmdOther_Click()
    Me.Height = 4605
    lblOwner.Caption = "车主:" & m_oREBus.OwnerName
    lblCompany.Caption = "参运公司:" & m_oREBus.CompanyName
    cmdOther.Enabled = False
End Sub

Private Sub cmdOwner_Click()
    frmBusOwner.Status = SNModify
    frmBusOwner.szOwnerID = m_oREBus.Owner
    frmBusOwner.Show vbModal
End Sub

Private Sub dtpNewTime_Change()
    IsSave
End Sub

Private Sub Form_Load()
m_oREBus.Init m_oActiveUser
If m_szBusID <> "" Then
    FullBus
    cmdOk.Enabled = False
Else
    m_dtBusDate = Date
End If
Me.Caption = "环境--车次属性[" & Format(m_dtBusDate, "YYYY年MM月DD日") & "]"
End Sub

Private Sub FullBus()
    Dim oCheckGate As New CheckGate
On Error GoTo here
    oCheckGate.Init m_oActiveUser
    
    m_oREBus.Identify m_szBusID, m_dtBusDate
    txtBusID.Text = m_szBusID
    lblRoute.Caption = m_oREBus.RouteName
    If m_oREBus.BusType = TP_RegularBus Then
        lblOffTime.Caption = Format(m_oREBus.StartupTime, "HH:MM:SS")
    Else
        lblOffTime.Caption = "流水车次"
        dtpNewTime.Enabled = False
    End If
    lblVehicle.Caption = m_oREBus.Vehicle
    dtpNewTime.Value = m_oREBus.StartupTime
    lblSellSeat.Caption = m_oREBus.SaledSeatCount
    oCheckGate.Identify m_oREBus.CheckGate
    lblCheckGate.Caption = "检票口:" & m_oREBus.CheckGate & "[" & oCheckGate.CheckGateName & "]"
    txtCheckGate.Text = m_oREBus.CheckGate & "[" & oCheckGate.CheckGateName & "]"
    If m_oREBus.AllRefundment Then
        lblAllRefundment.Caption = "全额退票:是"
        optAllRe.Value = True
    Else
        lblAllRefundment.Caption = "全额退票:否"
        optNoR.Value = True
    End If
    Select Case m_oREBus.BusStatus
           Case ST_BusStopCheck: lblStatus.Caption = "车次状态:停检"
           Case ST_BusNormal: lblStatus.Caption = "车次状态:正常"
           Case ST_BusStopped: lblStatus.Caption = "车次状态:停班"
           Case ST_BusMergeStopped: lblStatus.Caption = "车次状态:并班"
    End Select
    cmdOk.Enabled = False
    cmdOther.Enabled = True
Exit Sub
here:
    ShowErrorMsg
End Sub

Private Sub optAllRe_Click()
IsSave
End Sub

Private Sub optNoR_Click()
IsSave
End Sub

Private Sub Timer1_Timer()
    lblNow.Caption = Format(Now, "HH:MM:SS")
End Sub

Private Sub txtBusId_Change()
    IsSave
End Sub

Private Sub txtBusID_Click()
    Dim oBus As New CommDialog
    Dim szaTemp() As String
    oBus.Init m_oActiveUser
    szaTemp = oBus.SelectREBus(m_dtBusDate, False, False)
    Set oBus = Nothing
    If ArrayLength(szaTemp) = 0 Then Exit Sub
    txtBusID.Text = szaTemp(1, 1)
    m_szBusID = txtBusID.Text
    FullBus
End Sub

Private Sub txtBusID_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
       Case vbKeyReturn
       m_szBusID = txtBusID.Text
       FullBus
End Select
End Sub

Private Sub IsSave()
    If txtBusID.Text = "" Or txtCheckGate.Text = "" Then
    cmdOk.Enabled = False
    Else
    cmdOk.Enabled = True
    End If
End Sub

Private Sub txtCheckGate_Change()
    IsSave
End Sub

Private Sub txtCheckGate_Click()
    Dim oShell As New CommDialog
    Dim szaTemp() As String
    oShell.Init m_oActiveUser
    szaTemp = oShell.SelectCheckGate(False)
    Set oShell = Nothing
    If ArrayLength(szaTemp) = 0 Then Exit Sub
    txtCheckGate.Text = szaTemp(1, 1)
End Sub
