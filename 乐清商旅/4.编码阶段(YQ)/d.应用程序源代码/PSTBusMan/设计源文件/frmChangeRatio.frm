VERSION 5.00
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmChangeRatio 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "修改费率"
   ClientHeight    =   4635
   ClientLeft      =   1125
   ClientTop       =   2235
   ClientWidth     =   7995
   HelpContextID   =   10000830
   Icon            =   "frmChangeRatio.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin FText.asFlatMemo txtRemark 
      Height          =   600
      Left            =   990
      TabIndex        =   13
      Top             =   3210
      Width           =   6810
      _ExtentX        =   12012
      _ExtentY        =   1058
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
   Begin RTComctl3.CoolButton cmdDelete 
      Height          =   315
      Left            =   3120
      TabIndex        =   14
      Top             =   4140
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "删除(&D)"
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
      MICON           =   "frmChangeRatio.frx":014A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtBaseRatio 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1455
      TabIndex        =   9
      Top             =   2835
      Width           =   885
   End
   Begin VB.TextBox txtRoadRatio 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   3480
      TabIndex        =   11
      Top             =   2835
      Width           =   1155
   End
   Begin VB.ListBox lstSeatType 
      Appearance      =   0  'Flat
      Height          =   2340
      ItemData        =   "frmChangeRatio.frx":0166
      Left            =   6000
      List            =   "frmChangeRatio.frx":0168
      Style           =   1  'Checkbox
      TabIndex        =   7
      Top             =   390
      Width           =   1845
   End
   Begin VB.ListBox lstRoadLevel 
      Appearance      =   0  'Flat
      Height          =   2340
      ItemData        =   "frmChangeRatio.frx":016A
      Left            =   4050
      List            =   "frmChangeRatio.frx":016C
      Style           =   1  'Checkbox
      TabIndex        =   5
      Top             =   390
      Width           =   1845
   End
   Begin VB.ListBox lstVehicleModel 
      Appearance      =   0  'Flat
      Height          =   2340
      ItemData        =   "frmChangeRatio.frx":016E
      Left            =   2100
      List            =   "frmChangeRatio.frx":0170
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   390
      Width           =   1845
   End
   Begin VB.ListBox lstArea 
      Appearance      =   0  'Flat
      Height          =   2340
      ItemData        =   "frmChangeRatio.frx":0172
      Left            =   165
      List            =   "frmChangeRatio.frx":0174
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   390
      Width           =   1845
   End
   Begin RTComctl3.CoolButton cmdSave 
      Height          =   315
      Left            =   4290
      TabIndex        =   15
      Top             =   4140
      Width           =   1095
      _ExtentX        =   1931
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
      MICON           =   "frmChangeRatio.frx":0176
      PICN            =   "frmChangeRatio.frx":0192
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
      Left            =   5460
      TabIndex        =   16
      Top             =   4140
      Width           =   1095
      _ExtentX        =   1931
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
      MICON           =   "frmChangeRatio.frx":052C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RTComctl3.CoolButton cmdHelp 
      Height          =   315
      Left            =   6630
      TabIndex        =   17
      Top             =   4140
      Width           =   1095
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
      MICON           =   "frmChangeRatio.frx":0548
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
      Caption         =   " "
      Enabled         =   0   'False
      Height          =   1110
      Left            =   -120
      TabIndex        =   19
      Top             =   3870
      Width           =   8745
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "备注(&R):"
      Height          =   180
      Left            =   195
      TabIndex        =   12
      Top             =   3240
      Width           =   720
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "座位类型(&T):"
      Height          =   180
      Left            =   6000
      TabIndex        =   6
      Top             =   180
      Width           =   1080
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "注：设定费率范围(0.0001-1.0000)"
      Height          =   180
      Left            =   5040
      TabIndex        =   18
      Top             =   2895
      Width           =   2790
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "公建金率(&R):"
      Height          =   180
      Left            =   2400
      TabIndex        =   10
      Top             =   2895
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "基本运价率(&B):"
      Height          =   180
      Left            =   150
      TabIndex        =   8
      Top             =   2895
      Width           =   1260
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "地区(&A):"
      Height          =   180
      Left            =   135
      TabIndex        =   0
      Top             =   150
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车型(&V):"
      Height          =   180
      Left            =   2070
      TabIndex        =   2
      Top             =   150
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "公路等级(&L):"
      Height          =   180
      Left            =   4050
      TabIndex        =   4
      Top             =   150
      Width           =   1080
   End
End
Attribute VB_Name = "frmChangeRatio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'**********************************************************
'* Source File Name:frmChangeRatio.frm
'* Project Name:RTBusMan
'* Engineer:
'* Data Generated:2002/08/27
'* Last Revision Date:2002/08/30
'* Brief Description:修改费率
'* Relational Document:UI_BS_SM_014.DOC
'**********************************************************
Private m_oBaseInfo As New BaseInfo
Private m_oCharge As New ChargeRatio
Private m_szaArea() As String
Private m_aszVehicleModel() As String
Private m_aszRoadlevel() As String
Private m_aszSeatType() As String

'

Private m_szaAreaTemp() As Integer
Private m_szaVehicleModelTemp() As Integer
Private m_szaRoadlevelTemp() As Integer
Private m_szaSeatTypeTemp() As Integer


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
Dim nAreaCount As Integer
Dim nVehicleModel As Integer
Dim nRoadLevel As Integer
Dim nSeatType As Integer
Dim X1, X2, X3, X4 As Integer
Dim nCount As Integer
Dim nstep As Integer
Dim tChargeRa As TChargeRatio
On Error GoTo ErrorHandle
If MsgBox("修改选择项的费率?", vbQuestion + vbYesNo + vbDefaultButton2, "费率管理") = vbNo Then Exit Sub
SetBusy
'nSeatType = ArrayLength(m_aszSeatType)
'nAreaCount = ArrayLength(m_szaArea)
'nVehicleModel = ArrayLength(m_aszVehicleModel)
'nRoadLevel = ArrayLength(m_aszRoadlevel)
nSeatType = ArrayLength(m_szaSeatTypeTemp)
nAreaCount = ArrayLength(m_szaAreaTemp)
nVehicleModel = ArrayLength(m_szaVehicleModelTemp)
nRoadLevel = ArrayLength(m_szaRoadlevelTemp)
WriteProcessBar , , (nAreaCount - 1) * (nVehicleModel - 1), "保存..."
'以下陆大师改过
For X1 = 1 To nAreaCount - 1
'     If lstArea.Selected(X1) = True Then
       If m_szaAreaTemp(X1) <> 0 Then
       lstArea.ListIndex = m_szaAreaTemp(X1) - 1
        For X2 = 1 To nVehicleModel - 1
'             If lstVehicleModel.Selected(X2) = True Then
'               lstVehicleModel.ListIndex = X2
              If m_szaVehicleModelTemp(X2) <> 0 Then
                 lstVehicleModel.ListIndex = m_szaVehicleModelTemp(X2) - 1
                For X3 = 1 To nRoadLevel - 1
'                     If lstRoadLevel.Selected(X3) = True Then
'                        lstRoadLevel.ListIndex = X3
                       If m_szaRoadlevelTemp(X3) <> 0 Then
                         lstRoadLevel.ListIndex = m_szaRoadlevelTemp(X3) - 1
                        For X4 = 1 To nSeatType - 1
'                                 If lstSeatType.Selected(X4) = True Then
'                                        lstSeatType.ListIndex = X4
                                        
                                      If m_szaSeatTypeTemp(X4) <> 0 Then
                                         lstSeatType.ListIndex = m_szaSeatTypeTemp(X4) - 1
                                        lstSeatType.ListIndex = m_szaSeatTypeTemp(X4) - 1
                                        tChargeRa.szAreaCode = Trim(m_szaArea(m_szaAreaTemp(X1), 1))
                                        tChargeRa.szVehicleModel = Trim(m_aszVehicleModel(m_szaVehicleModelTemp(X2), 1))
                                        tChargeRa.szRoadLevel = Trim(m_aszRoadlevel(m_szaRoadlevelTemp(X3), 1))
                                        tChargeRa.sgBaseCarriageRatio = Val(txtBaseRatio.Text)
                                        tChargeRa.sgRoadConstructFundRatio = Val(txtRoadRatio.Text)
                                        tChargeRa.szSeatType = m_aszSeatType(m_szaSeatTypeTemp(X4), 1)
                                        If txtRemark.Text <> "" Then
                                           tChargeRa.szAnnotation = Trim(txtRemark.Text)
                                        End If
                                       m_oCharge.DeleteChargeRatio tChargeRa
                                End If
                             
                        Next
                    End If
                 
                Next
            End If
           
        Next
    End If
   
Next
WriteProcessBar False
SetNormal
MsgBox "费率修改成功", vbInformation, "费率"
Exit Sub
ErrorHandle:
    SetNormal
    ShowErrorMsg
End Sub

Private Sub cmdHelp_Click()
DisplayHelp Me
End Sub

Private Sub cmdSave_Click()
    Dim nAreaCount As Integer
    Dim nVehicleModel As Integer
    Dim nRoadLevel As Integer
    Dim nSeatType As Integer
    Dim X1, X2, X3, X4 As Integer
    Dim nCount As Integer
    Dim nstep As Integer
    Dim tChargeRa As TChargeRatio
    On Error GoTo ErrorHandle
    If MsgBox("修改选择项的费率?", vbQuestion + vbYesNo + vbDefaultButton2, "费率管理") = vbNo Then Exit Sub
    SetBusy
    'nSeatType = ArrayLength(m_aszSeatType)
    'nAreaCount = ArrayLength(m_szaArea)
    'nVehicleModel = ArrayLength(m_aszVehicleModel)
    'nRoadLevel = ArrayLength(m_aszRoadlevel)
    nSeatType = ArrayLength(m_szaSeatTypeTemp)
    nAreaCount = ArrayLength(m_szaAreaTemp)
    nVehicleModel = ArrayLength(m_szaVehicleModelTemp)
    nRoadLevel = ArrayLength(m_szaRoadlevelTemp)
'    WriteProcessBar , , (nAreaCount - 1) * (nVehicleModel - 1) "保存..."
    '以下陆大师改过
    For X1 = 1 To nAreaCount - 1
    '     If lstArea.Selected(X1) = True Then
        If m_szaAreaTemp(X1) <> 0 Then
            lstArea.ListIndex = m_szaAreaTemp(X1) - 1
            For X2 = 1 To nVehicleModel - 1
            '             If lstVehicleModel.Selected(X2) = True Then
            '               lstVehicleModel.ListIndex = X2
                If m_szaVehicleModelTemp(X2) <> 0 Then
                    lstVehicleModel.ListIndex = m_szaVehicleModelTemp(X2) - 1
                    For X3 = 1 To nRoadLevel - 1
                    '                     If lstRoadLevel.Selected(X3) = True Then
                    '                        lstRoadLevel.ListIndex = X3
                        If m_szaRoadlevelTemp(X3) <> 0 Then
                            lstRoadLevel.ListIndex = m_szaRoadlevelTemp(X3) - 1
                            For X4 = 1 To nSeatType - 1
                            '                                 If lstSeatType.Selected(X4) = True Then
                            '                                        lstSeatType.ListIndex = X4
                            
                                If m_szaSeatTypeTemp(X4) <> 0 Then
                                    lstSeatType.ListIndex = m_szaSeatTypeTemp(X4) - 1
                                    lstSeatType.ListIndex = m_szaSeatTypeTemp(X4) - 1
                                    tChargeRa.szAreaCode = Trim(m_szaArea(m_szaAreaTemp(X1), 1))
                                    tChargeRa.szVehicleModel = Trim(m_aszVehicleModel(m_szaVehicleModelTemp(X2), 1))
                                    tChargeRa.szRoadLevel = Trim(m_aszRoadlevel(m_szaRoadlevelTemp(X3), 1))
                                    tChargeRa.sgBaseCarriageRatio = Val(txtBaseRatio.Text)
                                    tChargeRa.sgRoadConstructFundRatio = Val(txtRoadRatio.Text)
                                    tChargeRa.szSeatType = m_aszSeatType(m_szaSeatTypeTemp(X4), 1)
                                    If txtRemark.Text <> "" Then
                                        tChargeRa.szAnnotation = Trim(txtRemark.Text)
                                    End If
                                    m_oCharge.ModifyChargeRatio tChargeRa
                                End If
                            Next X4
                        End If
                    Next X3
                End If
            Next X2
        End If
    Next X1
    WriteProcessBar False
    SetNormal
    MsgBox "费率修改成功", vbInformation, "费率"
    Exit Sub
ErrorHandle:
    SetNormal
    ShowErrorMsg
End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
Dim i As Integer, nCount As Integer
On Error GoTo ErrorHandle
ReDim m_szaAreaTemp(0 To 0)
ReDim m_szaVehicleModelTemp(0 To 0)
ReDim m_szaRoadlevelTemp(0 To 0)
ReDim m_szaSeatTypeTemp(0 To 0)
m_oCharge.Init g_oActiveUser
m_oBaseInfo.Init g_oActiveUser
'填充地区
m_szaArea = m_oBaseInfo.GetAllArea
nCount = ArrayLength(m_szaArea)
For i = 1 To nCount
    lstArea.AddItem m_szaArea(i, 2)
Next

'填充车型
m_aszVehicleModel = m_oBaseInfo.GetAllVehicleModel
nCount = ArrayLength(m_aszVehicleModel)
For i = 1 To nCount
    lstVehicleModel.AddItem m_aszVehicleModel(i, 2)
Next

'填充公路等级
m_aszRoadlevel = m_oBaseInfo.GetAllRoadLevel
nCount = ArrayLength(m_aszRoadlevel)
For i = 1 To nCount
    lstRoadLevel.AddItem m_aszRoadlevel(i, 2)
Next
'填充座位类型
m_aszSeatType = m_oBaseInfo.GetAllSeatType
nCount = ArrayLength(m_aszSeatType)
For i = 1 To nCount
    lstSeatType.AddItem MakeDisplayString(m_aszSeatType(i, 1), m_aszSeatType(i, 2))
Next
Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

Private Sub IsSave()
If txtBaseRatio.Text = "" Or txtRoadRatio.Text = "" Then
    cmdSave.Enabled = False
Else
    cmdSave.Enabled = True
End If
End Sub

Private Sub lstArea_ItemCheck(Item As Integer)
   GetData m_szaAreaTemp, Item, lstArea
 End Sub

Private Sub lstArea_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
' Set g_lstfeilv = lstArea
'  If Button = vbRightButton Then
'       PopupMenu MDIScheme.pmnu_Main
'       If g_bClearfeilv = True Then
'       ReDim m_szaAreaTemp(0 To 0) As Integer
'       End If
'  End If
End Sub

Private Sub lstRoadLevel_ItemCheck(Item As Integer)
GetData m_szaRoadlevelTemp, Item, lstRoadLevel
End Sub



Private Sub lstRoadLevel_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
' Set g_lstfeilv = lstRoadLevel
'  If Button = vbRightButton Then
'       PopupMenu MDIScheme.pmnu_Main
'       If g_bClearfeilv = True Then
'       ReDim m_szaRoadlevelTemp(0 To 0) As Integer
'       End If
'  End If

End Sub

Private Sub lstSeatType_ItemCheck(Item As Integer)
GetData m_szaSeatTypeTemp, Item, lstSeatType

End Sub

Private Sub lstSeatType_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
' Set g_lstfeilv = lstSeatType
'  If Button = vbRightButton Then
'       PopupMenu MDIScheme.pmnu_Main
'        If g_bClearfeilv = True Then
'          ReDim m_szaSeatTypeTemp(0 To 0) As Integer
'        End If
'  End If
End Sub

Private Sub lstVehicleModel_ItemCheck(Item As Integer)
GetData m_szaVehicleModelTemp, Item, lstVehicleModel
End Sub

Private Sub lstVehicleModel_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
' Set g_lstfeilv = lstVehicleModel
'  If Button = vbRightButton Then
'      PopupMenu MDIScheme.pmnu_Main
'       If g_bClearfeilv = True Then
'         ReDim m_szaVehicleModelTemp(0 To 0) As Integer
'
'       End If
'  End If

End Sub

Private Sub txtBaseRatio_Change()
IsSave
End Sub
Private Sub txtRoadRatio_Change()
IsSave
End Sub

Private Function GetData(nTemp() As Integer, nItemp As Integer, ListBox As ListBox)
  Dim nCount As Integer
  Dim listTemp As ListBox
  Dim i As Integer
  Set listTemp = ListBox
  nCount = ArrayLength(nTemp)
  If nCount <> 0 Then
    For i = 0 To nCount - 1
        If nTemp(i) = nItemp + 1 Then
           nTemp(i) = 0
           Exit Function
        End If
    Next
        ReDim Preserve nTemp(0 To nCount)
        If listTemp.Selected(nItemp) = True Then
               nTemp(nCount) = nItemp + 1
        End If
  End If
End Function

