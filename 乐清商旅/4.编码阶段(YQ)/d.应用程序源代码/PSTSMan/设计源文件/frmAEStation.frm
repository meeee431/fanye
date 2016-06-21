VERSION 5.00
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Object = "{6F8DCFAB-B2C9-11D2-A5ED-DE08DCF33612}#3.2#0"; "asftext.ocx"
Begin VB.Form frmAEStation 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "新增车站或编辑车站属性"
   ClientHeight    =   3555
   ClientLeft      =   3750
   ClientTop       =   3765
   ClientWidth     =   6450
   HelpContextID   =   50000150
   Icon            =   "frmAEStation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   Begin RTComctl3.CoolButton CoolButton1 
      Height          =   315
      Left            =   2280
      TabIndex        =   16
      Top             =   3120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
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
      MICON           =   "frmAEStation.frx":0E42
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
      Left            =   3600
      TabIndex        =   12
      Top             =   3120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
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
      MICON           =   "frmAEStation.frx":0E5E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin FText.asFlatTextBox txtSiteID 
      Height          =   315
      Left            =   1320
      TabIndex        =   9
      Top             =   1250
      Width           =   2355
      _ExtentX        =   4154
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
   End
   Begin VB.TextBox txtStationID 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   2355
   End
   Begin VB.TextBox txtStationFullName 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1320
      TabIndex        =   5
      Top             =   480
      Width           =   4890
   End
   Begin VB.TextBox txtUnitAnnotation 
      Appearance      =   0  'Flat
      Height          =   930
      Left            =   210
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   1905
      Width           =   5985
   End
   Begin VB.TextBox txtStationShortName 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4905
      TabIndex        =   3
      Top             =   120
      Width           =   1275
   End
   Begin VB.ComboBox cboUnit 
      Height          =   300
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   870
      Width           =   2355
   End
   Begin RTComctl3.CoolButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   4920
      TabIndex        =   13
      Top             =   3120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
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
      MICON           =   "frmAEStation.frx":0E7A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   150
      X2              =   6215
      Y1              =   2940
      Y2              =   2940
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   150
      X2              =   6215
      Y1              =   2955
      Y2              =   2955
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车站代码(&C):"
      Height          =   180
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车站简称(&S):"
      Height          =   180
      Left            =   3795
      TabIndex        =   2
      Top             =   180
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车站全称(&F):"
      Height          =   180
      Left            =   180
      TabIndex        =   4
      Top             =   555
      Width           =   1080
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "车站注释(&A):"
      Height          =   180
      Left            =   180
      TabIndex        =   10
      Top             =   1650
      Width           =   1080
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "所属单位(&U):"
      Height          =   180
      Left            =   180
      TabIndex        =   6
      Top             =   930
      Width           =   1080
   End
   Begin VB.Label lblSelfUnit 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "此单位就是本单位"
      Height          =   180
      Left            =   1755
      TabIndex        =   15
      Top             =   930
      Width           =   1455
   End
   Begin VB.Label lblStationID 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   180
      Left            =   1320
      TabIndex        =   14
      Top             =   180
      Width           =   90
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "对应站点(G):"
      Height          =   180
      Left            =   180
      TabIndex        =   8
      Top             =   1290
      Width           =   1080
   End
End
Attribute VB_Name = "frmAEStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'  *******************************************************************
' *  Source File Name  : frmAEStation                                  *
' *  Project Name: PSTSMan                                    *
' *  Engineer:                          *
' *  Date Generated: 2002/08/19                      *
' *  Last Revision Date : 2002/08/19             *
' *  Brief Description   : 添加车站或编辑车站属性                   *
' *******************************************************************

Option Explicit
Public bEdit As Boolean
Dim szSellStationID As String
Dim szUnitID As String
Dim szSellStationName As String
Dim szSellStationFullName As String
Dim szStationID As String
Dim szAnno As String



Private Sub LoadStationInfo()
    Dim j As Integer
    Dim i As Integer
    Dim nInStrStart As Integer
    
    lblStationID.Caption = g_alvItemText(1)
    For i = 1 To ArrayLength(g_atAllSellStation)
        If g_atAllSellStation(i).szSellStationID = g_alvItemText(1) Then
            txtStationID.Text = g_atAllSellStation(i).szSellStationID
            txtStationShortName.Text = g_atAllSellStation(i).szSellStationName
            txtStationFullName.Text = g_atAllSellStation(i).szSellStationFullName
            txtSiteID.Text = Trim(g_atAllSellStation(i).szStationID) ', Trim(g_atAllSellStation(i).szStationName))
            lblStationID.Caption = g_atAllSellStation(i).szSellStationID
        End If
        For j = 0 To cboUnit.ListCount - 1
            nInStrStart = 0
            nInStrStart = InStr(1, cboUnit.List(j), g_atAllSellStation(i).szUnitID)
            If nInStrStart = 1 Then '使CboUnit显示此用户所属单位
                cboUnit.ListIndex = j
            End If
        Next j
    
    
    Next i
    
    
    
    
End Sub

Private Sub OKButton_Click()
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim oSellStation As New SellStation
    On Error GoTo ErrorHandle
    GetInfoFromUI
    If bEdit = True Then
        '修改用户
        oSellStation.Init g_oActUser
        oSellStation.Identify szSellStationID
        
        oSellStation.SellStationID = szSellStationID
        oSellStation.SellStationFullName = szSellStationFullName
        oSellStation.SellStationName = szSellStationName
        oSellStation.Anno = szAnno
        oSellStation.StationID = szStationID
        oSellStation.UnitID = szUnitID
        
        oSellStation.Update
    Else
        '新增用户
        oSellStation.Init g_oActUser
        oSellStation.AddNew
        oSellStation.SellStationFullName = szSellStationFullName
        oSellStation.SellStationID = szSellStationID
        oSellStation.UnitID = szUnitID
        oSellStation.Anno = szAnno
        oSellStation.SellStationName = szSellStationName
        oSellStation.StationID = szStationID
        oSellStation.Update
    End If
    frmStoreMenu.LoadCommonData
    frmStoreMenu.LoadStationInfo
    Set oSellStation = Nothing
    Unload Me
    Exit Sub
ErrorHandle:
    ShowErrorMsg
    
End Sub
Private Sub GetInfoFromUI()
    If bEdit = True Then
        szSellStationID = lblStationID.Caption
    Else
        szSellStationID = txtStationID.Text
    End If
    szUnitID = ResolveDisplay(cboUnit.Text)
    szAnno = txtUnitAnnotation.Text
    szSellStationFullName = txtStationFullName.Text
    szSellStationName = txtStationShortName.Text
    szStationID = txtSiteID.Text
    
    
End Sub

Private Sub CoolButton1_Click()
DisplayHelp Me
End Sub

Private Sub Form_Load()
    Dim nUnitCount As Integer
    Dim i As Integer
    AlignFormPos Me
    cboUnit.Clear
    
    nUnitCount = 0
    nUnitCount = ArrayLength(g_atAllUnit)
    If nUnitCount <> 0 Then
        For i = 1 To nUnitCount
            cboUnit.AddItem g_atAllUnit(i).szUnitID & "[" & g_atAllUnit(i).szUnitFullName & "]"
        Next
    Else
        ''''''
    End If
    If nUnitCount > 0 Then cboUnit.ListIndex = 0
    If bEdit Then
        '修改
        Me.Caption = "修改车站"
        cmdOk.Caption = "修改(&O)"
        cmdCancel.Caption = "取消(&C)"
        txtStationID.Visible = False
        lblStationID.Visible = True
        LoadStationInfo
        frmAEStation.HelpContextID = 50000200
    Else
        '新增
         Me.Caption = "新增车站"
        cmdOk.Caption = "新增(&A)"
        cmdCancel.Caption = "关闭(&C)"
        
        txtStationID.Visible = True
        lblStationID.Visible = False
        
        ClearTextBox Me
        frmAEStation.HelpContextID = 50000150
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveFormPos Me
End Sub

Private Sub txtSiteID_ButtonClick()
    Dim aszTemp() As String
    Dim oShell As New CommDialog
    On Error GoTo ErrorHandle
    
    oShell.Init g_oActUser
    aszTemp = oShell.SelectStation(, False)
    If ArrayLength(aszTemp) = 0 Then Exit Sub
    txtSiteID.Text = aszTemp(1, 1) ', aszTemp(1, 2))
Exit Sub
ErrorHandle:
    ShowErrorMsg
End Sub

