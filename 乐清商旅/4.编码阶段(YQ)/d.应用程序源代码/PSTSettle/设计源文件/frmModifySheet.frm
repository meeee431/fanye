VERSION 5.00
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Begin VB.Form frmModifySheet 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�޸�·����������"
   ClientHeight    =   4980
   ClientLeft      =   3375
   ClientTop       =   1935
   ClientWidth     =   6630
   Icon            =   "frmModifySheet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   6630
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ע��"
      Height          =   990
      Left            =   525
      TabIndex        =   25
      Top             =   870
      Width           =   5535
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "  �޸�·��������������˾��ԭ����Ʊʱ������ĳ�����˾��Ϊ��ȷ�ģ�����·������ʱ��ʹ�ô˹��ܣ�����ʱ����ʹ�á�"
         Height          =   555
         Left            =   960
         TabIndex        =   26
         Top             =   270
         Width           =   4470
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   225
         Picture         =   "frmModifySheet.frx":000C
         Top             =   285
         Width           =   480
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "·��ժҪ��Ϣ"
      Height          =   1845
      Left            =   525
      TabIndex        =   10
      Top             =   2280
      Width           =   5535
      Begin FText.asFlatTextBox txtVehicle 
         Height          =   285
         Left            =   1050
         TabIndex        =   3
         Top             =   885
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���δ���:"
         Height          =   180
         Left            =   225
         TabIndex        =   24
         Top             =   360
         Width           =   810
      End
      Begin VB.Label lblBusID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1234"
         Height          =   180
         Left            =   1125
         TabIndex        =   23
         Top             =   360
         Width           =   360
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����:"
         Height          =   180
         Left            =   225
         TabIndex        =   22
         Top             =   630
         Width           =   450
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2005-03-25"
         Height          =   180
         Left            =   1125
         TabIndex        =   21
         Top             =   630
         Width           =   900
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������:"
         Height          =   180
         Left            =   2865
         TabIndex        =   20
         Top             =   360
         Width           =   810
      End
      Begin VB.Label lblBusSerial 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "12"
         Height          =   180
         Left            =   3705
         TabIndex        =   19
         Top             =   360
         Width           =   180
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����:"
         Height          =   180
         Left            =   225
         TabIndex        =   2
         Top             =   930
         Width           =   450
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��˾:"
         Height          =   180
         Left            =   3225
         TabIndex        =   18
         Top             =   930
         Width           =   450
      End
      Begin VB.Label lblCompany 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������˼���"
         Height          =   180
         Left            =   3705
         TabIndex        =   17
         Top             =   930
         Width           =   1080
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��·:"
         Height          =   180
         Left            =   3225
         TabIndex        =   16
         Top             =   660
         Width           =   450
      End
      Begin VB.Label lblRoute 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���˸�����"
         Height          =   180
         Left            =   3705
         TabIndex        =   15
         Top             =   660
         Width           =   900
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ƱԱ:"
         Height          =   180
         Left            =   225
         TabIndex        =   14
         Top             =   1215
         Width           =   630
      End
      Begin VB.Label lblCheckor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�·�"
         Height          =   180
         Left            =   1125
         TabIndex        =   13
         Top             =   1215
         Width           =   360
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ƶ�ʱ��:"
         Height          =   180
         Left            =   2865
         TabIndex        =   12
         Top             =   1215
         Width           =   810
      End
      Begin VB.Label lblMakeTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2005-03-25 12:30:23"
         Height          =   180
         Left            =   3705
         TabIndex        =   11
         Top             =   1215
         Width           =   1710
      End
   End
   Begin VB.TextBox txtSheetID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1830
      TabIndex        =   1
      Text            =   "0214865"
      Top             =   1950
      Width           =   1410
   End
   Begin VB.PictureBox ptTitle 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   795
      Left            =   15
      ScaleHeight     =   795
      ScaleWidth      =   7185
      TabIndex        =   7
      Top             =   0
      Width           =   7185
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Height          =   60
         Left            =   0
         TabIndex        =   8
         Top             =   750
         Width           =   7215
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������Ҫ�޸ĵ�·�����:"
         Height          =   180
         Left            =   120
         TabIndex        =   9
         Top             =   300
         Width           =   2250
      End
   End
   Begin RTComctl3.CoolButton cmdPreView 
      Height          =   345
      Left            =   525
      TabIndex        =   6
      Top             =   4500
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "·��Ԥ��(&V)"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      MICON           =   "frmModifySheet.frx":08D6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdExit 
      Cancel          =   -1  'True
      Height          =   345
      Left            =   4875
      TabIndex        =   5
      Top             =   4500
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "ȡ��(&C)"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      MICON           =   "frmModifySheet.frx":08F2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdSave 
      Height          =   345
      Left            =   3345
      TabIndex        =   4
      Top             =   4500
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "����(&S)"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      MICON           =   "frmModifySheet.frx":090E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   " "
      Enabled         =   0   'False
      Height          =   990
      Left            =   -45
      TabIndex        =   27
      Top             =   4230
      Width           =   8745
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ԭ·�����(&N):"
      Height          =   180
      Left            =   540
      TabIndex        =   0
      Top             =   1995
      Width           =   1260
   End
End
Attribute VB_Name = "frmModifySheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mtSheetInfo As TCheckSheetInfo       'ԭʼ·����Ϣ
Private m_oChkTicket As New CheckTicket

Private Sub InitForm()
    txtSheetID.Text = ""
    
    
    lblBusID.Caption = ""
    lblBusSerial.Caption = ""
    lblDate.Caption = ""
    lblCheckor.Caption = ""
    txtVehicle.Text = ""
    
    lblCompany.Caption = ""
    lblMakeTime.Caption = ""
    lblRoute.Caption = ""

End Sub




Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPreView_Click()
    '��ʾ·������
        Dim oCommDialog As New STShell.CommDialog
        Dim szCheckSheetID As String
        szCheckSheetID = txtSheetID.Text
        If szCheckSheetID <> "" Then
            oCommDialog.Init g_oActiveUser
            oCommDialog.ShowCheckSheet szCheckSheetID
        End If
        
        Set oCommDialog = Nothing
End Sub

Private Sub cmdSave_Click()
    '��������
    Dim oSplit As New Split
    On Error GoTo ErrorHandle
    SetBusy
    oSplit.Init g_oActiveUser
    oSplit.ChangeSheetVehicle txtSheetID.Text, ResolveDisplay(txtVehicle.Text)
    ShowMsg "�����Ѹ�Ϊ[" & txtVehicle.Text & "]"
    SetNormal
    Exit Sub
ErrorHandle:
    SetNormal
    ShowErrorMsg
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendTab
    End If
End Sub

Private Sub Form_Load()
    InitForm
    m_oChkTicket.Init g_oActiveUser
    
End Sub


Private Sub txtSheetID_LostFocus()

    'InitForm
    
    RefreshCheckSheet
    
End Sub

Private Sub RefreshCheckSheet()
    'ˢ��·����Ϣ(�糵�Ρ���˾��)
    
    If txtSheetID.Text <> "" Then
    
        SetBusy
        mtSheetInfo = m_oChkTicket.GetCheckSheetInfo(txtSheetID.Text)
        If mtSheetInfo.szCheckSheet = "" Then
            ShowMsg "��·�������ڣ�"
        Else
            WriteSheetInfo
        End If
        
        SetNormal
    End If

End Sub



Private Sub WriteSheetInfo()
    Dim oVehicle As New Vehicle
    Dim oRoute As New Route
    oVehicle.Init g_oActiveUser
    oVehicle.Identify mtSheetInfo.szVehicleID
    oRoute.Init g_oActiveUser
    oRoute.Identify mtSheetInfo.szRouteID
    
    lblBusID.Caption = mtSheetInfo.szBusID
    lblBusSerial.Caption = mtSheetInfo.nBusSerialNo
    lblCheckor.Caption = mtSheetInfo.szMakeSheetUser
    lblCompany.Caption = MakeDisplayString(Trim(mtSheetInfo.szCompanyID), Trim(oVehicle.CompanyName))
    lblDate.Caption = Format(mtSheetInfo.dtDate, "YYYY-MM-DD")
    txtVehicle.Text = MakeDisplayString(oVehicle.VehicleID, oVehicle.LicenseTag)
    lblMakeTime.Caption = Format(mtSheetInfo.dtMakeSheetDateTime, "YYYY-MM-DD HH:MM:SS")
    lblRoute.Caption = oRoute.RouteName
    Set oVehicle = Nothing
    Set oRoute = Nothing
End Sub

Private Sub txtVehicle_ButtonClick()

    '��ʾ����
    On Error GoTo ErrHandle
    Dim oShell As New STShell.CommDialog
    Dim aszTemp() As String
    Dim oVehicle As New Vehicle
    SetBusy
    oShell.Init g_oActiveUser

    aszTemp = oShell.SelectVehicleEX()
    
    Set oShell = Nothing
    If ArrayLength(aszTemp) = 0 Then Exit Sub
    txtVehicle.Text = MakeDisplayString(Trim(aszTemp(1, 1)), Trim(aszTemp(1, 2)))
    
    oVehicle.Init g_oActiveUser
    oVehicle.Identify mtSheetInfo.szVehicleID
    
    lblCompany.Caption = MakeDisplayString(Trim(mtSheetInfo.szCompanyID), Trim(oVehicle.CompanyName))
    
    SetNormal
    
    Exit Sub
ErrHandle:
    SetNormal
    ShowErrorMsg
    
End Sub

