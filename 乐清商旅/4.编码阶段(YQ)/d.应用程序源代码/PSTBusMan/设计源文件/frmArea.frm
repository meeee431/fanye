VERSION 5.00
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmArea 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "地区"
   ClientHeight    =   4050
   ClientLeft      =   4470
   ClientTop       =   2820
   ClientWidth     =   5385
   HelpContextID   =   10000130
   Icon            =   "frmArea.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   Begin RTComctl3.CoolButton CoolButton1 
      Height          =   315
      Left            =   450
      TabIndex        =   16
      Top             =   3570
      Width           =   1065
      _ExtentX        =   1879
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
      MICON           =   "frmArea.frx":0E42
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
      Left            =   4200
      TabIndex        =   11
      Top             =   3570
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "关闭(&L)"
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
      MICON           =   "frmArea.frx":0E5E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox TxtAreaID 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1485
      TabIndex        =   1
      Top             =   900
      Width           =   3210
   End
   Begin FText.asFlatMemo txtAnnotation 
      Height          =   810
      Left            =   1485
      TabIndex        =   5
      Top             =   1770
      Width           =   3210
      _ExtentX        =   5662
      _ExtentY        =   1429
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonHotForeColor=   -2147483628
      ButtonHotBackColor=   -2147483632
   End
   Begin VB.PictureBox ptTitle 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   6615
      TabIndex        =   12
      Top             =   0
      Width           =   6615
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   60
         Left            =   0
         TabIndex        =   15
         Top             =   660
         Width           =   6885
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请修改或新增地区信息:"
         Height          =   180
         Left            =   120
         TabIndex        =   13
         Top             =   270
         Width           =   1890
      End
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1485
      TabIndex        =   3
      Top             =   1335
      Width           =   3210
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "位置"
      Height          =   600
      Left            =   1485
      TabIndex        =   6
      Top             =   2640
      Width           =   3210
      Begin VB.OptionButton OptOutCity 
         BackColor       =   &H00E0E0E0&
         Caption         =   "市外(&U)"
         Height          =   195
         Left            =   1200
         TabIndex        =   8
         Top             =   270
         Width           =   930
      End
      Begin VB.OptionButton OptInCity 
         BackColor       =   &H00E0E0E0&
         Caption         =   "市内(&I)"
         Height          =   195
         Left            =   225
         TabIndex        =   7
         Top             =   270
         Value           =   -1  'True
         Width           =   1020
      End
      Begin VB.OptionButton optOutProvince 
         BackColor       =   &H00E0E0E0&
         Caption         =   "省外(&O)"
         Height          =   195
         Left            =   2160
         TabIndex        =   9
         Top             =   270
         Width           =   1020
      End
   End
   Begin RTComctl3.CoolButton cmdOk 
      Height          =   315
      Left            =   2970
      TabIndex        =   10
      Top             =   3570
      Width           =   1065
      _ExtentX        =   1879
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
      MICON           =   "frmArea.frx":0E7A
      PICN            =   "frmArea.frx":0E96
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
      Caption         =   " RTStation"
      Enabled         =   0   'False
      Height          =   930
      Left            =   -150
      TabIndex        =   14
      Top             =   3330
      Width           =   8745
   End
   Begin VB.Label lblMemo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "备注(&R):"
      Height          =   180
      Left            =   375
      TabIndex        =   4
      Top             =   1785
      Width           =   720
   End
   Begin VB.Label lblObjectName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "名称(&N):"
      Height          =   180
      Left            =   345
      TabIndex        =   2
      Top             =   1350
      Width           =   720
   End
   Begin VB.Label lblObjectA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "地区代码(&C):"
      Height          =   180
      Left            =   345
      TabIndex        =   0
      Top             =   960
      Width           =   1170
   End
End
Attribute VB_Name = "frmArea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status As EFormStatus

Private moArea As Area '地区对象 Area
Public mszAreaID As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdOk_Click()
On Error GoTo ErrHandle
    If OptInCity.Value = True Then
      moArea.FlgProvince = EA_nInCity
    ElseIf OptOutCity.Value = True Then
      moArea.FlgProvince = EA_nOutCity
    ElseIf optOutProvince.Value = True Then
      moArea.FlgProvince = EA_nOutProvince
    End If
    
    Select Case Status
        Case EFormStatus.EFS_AddNew
            moArea.AddNew
            moArea.AreaCode = TxtAreaID.Text
            moArea.AreaName = txtName.Text
            moArea.Annotation = txtAnnotation.Text
            
            
            If OptInCity.Value = True Then
               moArea.FlgProvince = EA_nInCity
            ElseIf OptOutCity.Value = True Then
               moArea.FlgProvince = EA_nOutCity
            ElseIf optOutProvince.Value = True Then
               moArea.FlgProvince = EA_nOutProvince
            End If
            moArea.Update
      Case EFormStatus.EFS_Modify
            moArea.Identify TxtAreaID.Text
            moArea.AreaName = txtName.Text
            moArea.Annotation = txtAnnotation.Text
             
            If OptInCity.Value = True Then
                 moArea.FlgProvince = EA_nInCity
            ElseIf OptOutCity.Value = True Then
                 moArea.FlgProvince = EA_nOutCity
            ElseIf optOutProvince.Value = True Then
                moArea.FlgProvince = EA_nOutProvince
            End If
            
            moArea.Update
    End Select
        
    Dim aszInfo(1 To 4) As String
    aszInfo(1) = Trim(TxtAreaID.Text)
    aszInfo(2) = Trim(txtName.Text)
    aszInfo(3) = moArea.FlgProvince
    aszInfo(4) = Trim(txtAnnotation.Text)
    
    '刷新基本信息窗体
    Dim oListItem As ListItem
    If Status = EFormStatus.EFS_Modify Then
        frmBaseInfo.UpdateList aszInfo
        Unload Me
        Exit Sub
    End If
    If Status = EFormStatus.EFS_AddNew Then
        frmBaseInfo.AddList aszInfo
        RefreshArea
        TxtAreaID.SetFocus
        Exit Sub
    End If
    Exit Sub
ErrHandle:
    ShowErrorMsg
End Sub

Private Sub CoolButton1_Click()
    DisplayHelp Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    
           Case vbKeyReturn
                SendKeys "{TAB}"
    End Select
End Sub
Private Sub Form_Load()
On Error GoTo ErrHandle
    '布置窗体
    AlignFormPos Me
    
    Set moArea = CreateObject("STBase.Area")
    moArea.Init g_oActiveUser
    Select Case Status
        Case EFormStatus.EFS_AddNew
           cmdOk.Caption = "新增(&A)"
           RefreshArea
        Case EFormStatus.EFS_Modify
           TxtAreaID.Enabled = False
           RefreshArea
        Case EFormStatus.EFS_Show
           TxtAreaID.Enabled = False
           RefreshArea
    End Select
    cmdOk.Enabled = False
    Exit Sub
ErrHandle:
    Status = EFS_AddNew
    ShowErrorMsg
End Sub

Public Sub RefreshArea()
    If Status = EFS_AddNew Then
        TxtAreaID.Text = ""
        txtAnnotation.Text = ""
        txtName.Text = ""
    Else
        TxtAreaID.Text = mszAreaID
        moArea.Identify Trim(mszAreaID)
        txtAnnotation.Text = moArea.Annotation
        txtName = moArea.AreaName
        
         If moArea.FlgProvince = EA_nOutProvince Then
              optOutProvince.Value = True
         ElseIf moArea.FlgProvince = EA_nOutCity Then
              OptOutCity.Value = True
         ElseIf moArea.FlgProvince = EA_nInCity Then
              OptInCity.Value = True
         End If
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set moArea = Nothing
    SaveFormPos Me
End Sub

Private Sub OptInCity_Click()
    IsSave
End Sub

Private Sub OptOutCity_Click()
    IsSave
End Sub

Private Sub optoutprovince_Click()
    IsSave
End Sub
Private Sub txtAnnotation_Change()
    IsSave
End Sub

Private Sub txtAnnotation_GotFocus()
    cmdOk.Default = False
End Sub

Private Sub txtAnnotation_LostFocus()
    cmdOk.Default = True
End Sub

Private Sub TxtAreaID_Change()
    IsSave
    FormatTextBoxBySize TxtAreaID, 2
End Sub

'Private Sub TxtAreaID_Click()
'On Error GoTo ErrHandle
'    Dim aszTemp() As String
'    If Status = EFS_AddNew Then
'        MsgBox "请输入新增地区代码!", vbInformation, "地区"
'        Exit Sub
'    End If
'    Dim oShell As New stshell.commdialog
'    oShell.Init g_oActiveUser
'    aszTemp = oShell.selectArea(False)
'    Set oShell = Nothing
'    If ArrayLength(aszTemp) = 0 Then Exit Sub
'    TxtAreaID.Text = aszTemp(1, 1)
'    Exit Sub
'ErrHandle:
'    ShowErrorMsg
'End Sub

Private Sub IsSave()
    If TxtAreaID.Text = "" Or txtName.Text = "" Then
        cmdOk.Enabled = False
    Else
        cmdOk.Enabled = True
    End If
End Sub

Private Sub txtName_Change()
    IsSave
    FormatTextBoxBySize txtName, 20
End Sub

