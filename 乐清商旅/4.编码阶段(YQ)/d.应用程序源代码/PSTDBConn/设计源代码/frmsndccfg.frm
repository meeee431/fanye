VERSION 5.00
Begin VB.Form frmSNDCConfig 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���ݿ���������"
   ClientHeight    =   3720
   ClientLeft      =   3045
   ClientTop       =   2430
   ClientWidth     =   6030
   Icon            =   "frmsndccfg.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   6030
   StartUpPosition =   2  '��Ļ����
   Tag             =   "10"
   Begin VB.OptionButton optUserPwd 
      BackColor       =   &H00FFFFC0&
      Caption         =   "ָ�����û�������:"
      Height          =   195
      Left            =   270
      TabIndex        =   17
      Top             =   1950
      Width           =   2175
   End
   Begin VB.OptionButton optWindows 
      BackColor       =   &H00FFFFC0&
      Caption         =   "ʹ��Windows�����֤"
      Height          =   195
      Left            =   270
      TabIndex        =   16
      Top             =   1650
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txtPassword 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1605
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   2535
      Width           =   2750
   End
   Begin VB.TextBox txtUserName 
      Height          =   300
      Left            =   1605
      TabIndex        =   5
      Top             =   2190
      Width           =   2750
   End
   Begin VB.TextBox txtTimeOut 
      Height          =   300
      Left            =   1605
      TabIndex        =   15
      Top             =   3000
      Visible         =   0   'False
      Width           =   2750
   End
   Begin VB.CommandButton cmdAdvance 
      Caption         =   " �߼�(&V)>>"
      Height          =   345
      Left            =   4650
      TabIndex        =   13
      Top             =   2490
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Frame fraSQL 
      BackColor       =   &H00FFFFC0&
      Caption         =   "���Ӳ���"
      Height          =   1035
      Left            =   150
      TabIndex        =   20
      Top             =   510
      Width           =   4335
      Begin VB.TextBox txtDatabaseName 
         Height          =   300
         Left            =   1425
         TabIndex        =   3
         Top             =   600
         Width           =   2750
      End
      Begin VB.TextBox txtServer 
         Height          =   300
         Left            =   1425
         TabIndex        =   1
         Top             =   210
         Width           =   2750
      End
      Begin VB.Label lblDataBase 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ݿ�����(&D):"
         Height          =   180
         Left            =   120
         TabIndex        =   2
         Top             =   660
         Width           =   1260
      End
      Begin VB.Label lblServer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����������(&S):"
         Height          =   180
         Left            =   120
         TabIndex        =   0
         Top             =   270
         Width           =   1260
      End
   End
   Begin VB.ComboBox cboDatabaseType 
      Height          =   300
      ItemData        =   "frmsndccfg.frx":16AC2
      Left            =   1560
      List            =   "frmsndccfg.frx":16AC4
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   120
      Width           =   2925
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��"
      Height          =   345
      Left            =   4650
      TabIndex        =   10
      Top             =   540
      Width           =   1125
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��"
      Default         =   -1  'True
      Height          =   345
      Left            =   4650
      TabIndex        =   9
      Top             =   120
      Width           =   1125
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Caption         =   "��������"
      Height          =   2865
      Left            =   7350
      TabIndex        =   18
      Top             =   0
      Width           =   4470
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   " ����(&T)"
      Height          =   345
      Left            =   4650
      MaskColor       =   &H00000000&
      TabIndex        =   8
      Tag             =   "101"
      Top             =   2070
      Width           =   1125
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����(&P):"
      Height          =   180
      Left            =   570
      TabIndex        =   6
      Top             =   2550
      Width           =   720
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�û���(&U):"
      Height          =   180
      Left            =   570
      TabIndex        =   4
      Top             =   2235
      Width           =   900
   End
   Begin VB.Label lblTimeOut 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ʱʱ��(&M):"
      Height          =   180
      Left            =   210
      TabIndex        =   14
      Top             =   3060
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���ݿ�����(&C):"
      Height          =   180
      Left            =   150
      TabIndex        =   11
      Top             =   180
      Width           =   1260
   End
   Begin VB.Image imgStep 
      BorderStyle     =   1  'Fixed Single
      Height          =   4110
      Index           =   0
      Left            =   8010
      Picture         =   "frmsndccfg.frx":16AC6
      Top             =   -360
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Image imgStep 
      BorderStyle     =   1  'Fixed Single
      Height          =   4140
      Index           =   1
      Left            =   8130
      Top             =   -150
      Width           =   2160
   End
   Begin VB.Label lblTestPrompt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���ڽ������ݿ����Ӳ���..."
      Height          =   180
      Left            =   180
      TabIndex        =   19
      Top             =   3420
      Visible         =   0   'False
      Width           =   2250
   End
End
Attribute VB_Name = "frmSNDCConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'oledb����
Const OLEDB_MSSQLSERVER = "SQLOLEDB.1"
Const OLEDB_SYBASE = "SYBASE.ASEOLEDBPROVIDER.2"
Const OLEDB_MSJET = "Microsoft.Jet.OLEDB.3.51"
Const OLEDB_ODBC = "MSDASQL.1"

Private Declare Function WinHelp Lib "user32" Alias "WinHelpA" _
    (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long

'module level vars
Dim mszSNSystemName As String    '�����õ�ϵͳ����
Dim mbHasChange As Boolean
Dim mbFinishOK      As Boolean




Private Sub optUserPwd_Click()
    LayoutParamSet False
End Sub

Private Sub optWindows_Click()
    LayoutParamSet True
End Sub

Private Sub cmdAdvance_Click()
    If lblTimeOut.Visible Then
        Expand False
        cmdAdvance.Caption = "�߼�(&V)>>"
    Else
        Expand True
        cmdAdvance.Caption = "�߼�(&V)<<"
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    
End Sub

Private Sub cmdOk_Click()
    SetDataConnectReg mszSNSystemName
    Unload Me
End Sub

Private Sub cmdTest_Click()
    '���Բ����Ƿ���ȷ
    Dim oConnect As New ADODB.Connection
    Dim szConn As String
    On Error GoTo ErrorDo
    
    '�������ΪBUSY
    Me.MousePointer = vbHourglass


    Dim szDatabaseType As String
    Dim szServer As String, szUser As String, szPassword As String, szDatabase As String, szTimeout As String
    Dim szDriverType As String
    Dim szIntegrated As String '�Ƿ񼯳��ʻ�
    Select Case cboDatabaseType.ListIndex
    Case 0
        szDatabaseType = OLEDB_MSSQLSERVER
    Case 1
        szDatabaseType = OLEDB_ODBC
        szDriverType = "Sybase System 11"
    Case 2
        szDatabaseType = OLEDB_SYBASE
    Case Else
        szDatabaseType = OLEDB_MSSQLSERVER
    End Select
    szServer = "":    szUser = "": szPassword = "": szDatabase = "": szTimeout = ""

    szServer = Trim(txtServer.Text)
    szUser = Trim(txtUserName.Text)
    szPassword = txtPassword.Text
    szDatabase = Trim(txtDatabaseName.Text)
    szTimeout = Trim(txtTimeOut.Text)
    
    Select Case szDatabaseType
        Case OLEDB_MSSQLSERVER  'SQL Server
'SQLServer��֤��ʽ
'Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=foricq;Data Source=jhxu
'NT���ɷ�ʽ
'Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=RTArchDB;Data Source=LENGEND
            '�Ƿ񼯳ɷ�ʽ
            szIntegrated = IIf(optWindows.Value = True, "SSPI", "")
            szConn = "Provider=" & szDatabaseType _
            & ";Persist Security Info=False" _
            & IIf(szIntegrated <> "", ";Integrated Security=" & szIntegrated, ";User ID=" & szUser & ";Password=" & szPassword) _
            & ";Initial Catalog=" & szDatabase _
            & ";Data Source=" & szServer _
            & IIf(szTimeout = "", "", ";Timeout=" & Val(szTimeout))
        Case OLEDB_ODBC
            'ODBC��������
            Select Case szDriverType
                Case "Sybase System 11"     'Sybase 11.xϵ��
'Sybase 11.x�����ַ���
    'Provider=MSDASQL.1;Persist Security Info=False
    ';Extended Properties="DRIVER={Sybase System 11};UID=lyq;DB=RTArchDB;SRVR=CHENF;PWD=activex"
                    szConn = "Provider=" & szDatabaseType & ";Persist Security Info=False" _
                    & ";Extended Properties=""DRIVER={" & szDriverType & "}" & ";UID=" & szUser & ";DB=" & szDatabase & ";SRVR=" & szServer & ";PWD=" & szPassword & """" _
                    & IIf(szTimeout = "", "", ";Timeout=" & Val(szTimeout))
                Case Else       '����ODBC��������
                    szConn = ""
            End Select
        Case OLEDB_SYBASE    'Sybase OLE DB���İ�
'Provider=Sybase.ASEOLEDBProvider.2;������ȫ����Ϣ=False;�û� ID=sa;����Դ=sybase11
            szConn = "Provider=" & szDatabaseType _
            & ";������ȫ����Ϣ=False" _
            & ";�û� ID=" & szUser & ";����=" & szPassword _
            & ";����Դ=" & szServer _
            & IIf(szTimeout = "", "", ";��ʱ����=" & Val(szTimeout))
        Case Else
            szConn = ""
    End Select
    
    '��ʼ����
    lblTestPrompt.Visible = True
    lblTestPrompt.Refresh
    Debug.Print szConn
    oConnect.Open szConn
    Me.MousePointer = vbDefault
    MsgBox "���ӳɹ�!", vbInformation + vbOKOnly
    On Error Resume Next
    oConnect.Close
    lblTestPrompt.Visible = False
    cmdOk.SetFocus
    Exit Sub
ErrorDo:
    Me.MousePointer = vbDefault
    MsgBox "����ʧ��!" & vbCrLf & Err.Description, vbCritical + vbOKOnly
    lblTestPrompt.Visible = False
    txtServer.SetFocus
End Sub



Private Sub Form_Load()
    Dim i As Integer
    '��ʼ�����б���
    AddDBList
    mbFinishOK = False
    GetDataConnectReg mszSNSystemName
    'Set imgStep(1).Picture = imgStep(0).Picture
'    cboDatabaseType.ListIndex = 0
    SetValid
    Expand False
End Sub


'�ж��Ƿ�������Ч
Private Function SetValid() As Boolean
    Dim bTemp As Boolean
    bTemp = (Not Trim(txtServer.Text) = "") And _
        (Not Trim(txtDatabaseName.Text) = "") And _
        (Not Trim(txtUserName.Text) = "")
    cmdTest.Enabled = bTemp
    
    SetValid = bTemp
End Function


'=========================================================
'���û�û���ṩ�㹻����Ϣ������ִ��
'����Ľ���ʱ,�˺�������ʾһ��������Ϣ
'=========================================================
Private Sub cboDatabaseType_Change()
    mbHasChange = True
    SetValid
End Sub

Private Sub cboDatabaseType_Click()
    If cboDatabaseType.ListIndex = 0 Then
        optWindows.Enabled = True
    Else
        optWindows.Value = False
        optWindows.Enabled = False
    End If
    Select Case cboDatabaseType.ListIndex
        Case 2  'Sybase OLEDB
            txtDatabaseName.Visible = False
            lblDataBase.Caption = "����Դ������Ԥ��ͨ��Sybase OLE DB��������"
            lblDataBase.ForeColor = vbRed
            lblServer.Caption = "����Դ����(&D):"
        Case Else
            txtDatabaseName.Visible = True
            lblDataBase.Caption = "���ݿ�����(&D):"
            lblDataBase.ForeColor = &H80000012
            lblServer.Caption = "����������(&S):"
    End Select
'    optUserPwd.Enabled = True
End Sub



Private Sub Label6_Click()

End Sub

Private Sub txtDatabaseName_Change()
    mbHasChange = True
    SetValid
End Sub

Private Sub txtDatabaseName_GotFocus()
    txtDatabaseName.SelStart = 0
    txtDatabaseName.SelLength = Len(txtDatabaseName.Text)
End Sub

Private Sub txtPassword_Change()
    mbHasChange = True
    SetValid
End Sub

Private Sub txtPassword_GotFocus()
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword.Text)
End Sub

Private Sub txtServer_Change()
    mbHasChange = True
    SetValid
End Sub

Private Sub txtServer_GotFocus()
    txtServer.SelStart = 0
    txtServer.SelLength = Len(txtServer.Text)
End Sub

Private Sub txtUserName_Change()
    mbHasChange = True
    SetValid
End Sub
'������ע������ݿ����Ӳ���

Private Sub GetDataConnectReg(SetConnectName As String)
    Dim oReg As New CFreeReg
    Dim szDatabaseType As String
    Dim szConnectName As String
    Dim szIntegrated  As String
    
    oReg.Init "RTStation", HKEY_LOCAL_MACHINE, "Software\RTSoft"
        
    '�������õ�ϵͳ��ȡ����
    Select Case SetConnectName
    Case "SNGeneralConnect"
        'ͨ������
        szConnectName = "DataBaseSet"
    Case Else
        szConnectName = "DataBaseSet"
    End Select
    szDatabaseType = oReg.GetSetting(szConnectName, "DBType", "sqloledb.1")
    Select Case UCase(szDatabaseType)
        Case OLEDB_MSSQLSERVER
            'MS SQL Server���ݿ�
            cboDatabaseType.ListIndex = 0
'            optWindows.Enabled = True
        Case OLEDB_ODBC
            cboDatabaseType.ListIndex = 1
        Case OLEDB_SYBASE
            cboDatabaseType.ListIndex = 2
        Case Else
            cboDatabaseType.ListIndex = 0
'            optWindows.Enabled = True
    End Select
    txtServer.Text = oReg.GetSetting(szConnectName, "DBServer", "")
    txtDatabaseName.Text = oReg.GetSetting(szConnectName, "Database", "")
'    txtPassword.Text = oReg.GetSetting(szConnectName, "Password", "")
    txtUserName.Text = oReg.GetSetting(szConnectName, "User", "")
    txtTimeOut.Text = oReg.GetSetting(szConnectName, "TimeOut", "")
    szIntegrated = oReg.GetSetting(szConnectName, "InteGrated", "")
    If UCase(szIntegrated) = "SSPI" Then
        optWindows.Value = True
        LayoutParamSet True
    Else
        optUserPwd.Value = True
        LayoutParamSet False
    End If
End Sub
'�������ݿ����Ӳ���
Private Sub SetDataConnectReg(SetConnectName As String)
    Dim oReg As New CFreeReg
    Dim szDatabaseType As String
    Dim szDBDriverType As String    'ODBC�����������
    Dim szConnectName As String
    Dim szPassword As String
    oReg.Init "RTStation", HKEY_LOCAL_MACHINE, "Software\RTSoft"
    '
    szDatabaseType = ""
    szDBDriverType = ""
    szConnectName = ""
    szPassword = ""
    Dim szTmpPassword  As String
    
    
    '�������õ�ϵͳ��ȡ����
    Select Case SetConnectName
    Case "SNGeneralConnect"
        'ͨ������
        szConnectName = "DataBaseSet"
    Case Else
        szConnectName = "DataBaseSet"
    End Select
    Select Case cboDatabaseType.ListIndex
    Case 0
        szDatabaseType = OLEDB_MSSQLSERVER
    Case 1
        szDatabaseType = OLEDB_ODBC
        szDBDriverType = "Sybase System 11"
    Case 2
        szDatabaseType = OLEDB_SYBASE
    Case Else
        szDatabaseType = OLEDB_MSSQLSERVER
    End Select
    oReg.SaveSetting szConnectName, "DBType", szDatabaseType
    oReg.SaveSetting szConnectName, "DBServer", Trim(txtServer.Text)
    oReg.SaveSetting szConnectName, "DBDriverType", szDBDriverType
    oReg.SaveSetting szConnectName, "Database", Trim(txtDatabaseName.Text)
    szTmpPassword = EncryptPassword(txtPassword.Text)
    szPassword = IIf(szTmpPassword = "", "", szTmpPassword)
    oReg.SaveSetting szConnectName, "Password", szPassword
    oReg.SaveSetting szConnectName, "User", Trim(txtUserName.Text)
    If IsNumeric(txtTimeOut.Text) Then
        oReg.SaveSetting szConnectName, "TimeOut", txtTimeOut.Text
    Else
        oReg.SaveSetting szConnectName, "TimeOut", ""
    End If
    If optWindows.Value = True Then
        oReg.SaveSetting szConnectName, "Integrated", "SSPI"
    Else
        oReg.SaveSetting szConnectName, "Integrated", ""
    End If
End Sub

Private Sub txtUserName_GotFocus()
    txtUserName.SelStart = 0
    txtUserName.SelLength = Len(txtUserName.Text)
End Sub

Private Sub Expand(pbVisible As Boolean)
    If pbVisible Then
        lblTestPrompt.Top = 3420
        Me.Height = 4095
    Else
        lblTestPrompt.Top = 2970
        Me.Height = 3660
    End If
    lblTimeOut.Visible = pbVisible
    txtTimeOut.Visible = pbVisible
End Sub



Private Sub AddDBList()
    cboDatabaseType.Clear
    cboDatabaseType.AddItem "Microsoft SQL Server 7.x,2000"
    cboDatabaseType.AddItem "Sybase Adpative Server 11.x"
    cboDatabaseType.AddItem "Sybase Adpative Server 12.x"
    
    'cboDatabaseType.AddItem "Jet3.51"
End Sub


Private Sub LayoutParamSet(pbIsWindows As Boolean)
    If pbIsWindows Then
        txtPassword.Enabled = False
        txtUserName.Enabled = False
    Else
        txtPassword.Enabled = True
        txtUserName.Enabled = True
    End If

End Sub
' *******************************************************************
' *   Brief Description: ���ܿ���                                   *
' *   Engineer: ½����                                              *
' *   Date Generated: 2002/06/21                                    *
' *   Last Revision Date:                                           *
' *******************************************************************
Public Function EncryptPassword(ByVal pszPassword As String) As String
'pszPassword ����
'ѡ��һ�ּ����㷨�Կ�����м���
    
    Dim nLen As Integer
    Dim nPwdLen As Integer
    Dim i As Integer
    Dim szResult As String
    Dim nIndex As Integer

    nPwdLen = Len(pszPassword)
    If nPwdLen = 0 Then Exit Function
    If nPwdLen > 99 Then
        nPwdLen = 99
        pszPassword = Left(pszPassword, nPwdLen)
    End If
    
    Dim szTmp As String
    Dim szTmp2 As String
    szResult = ""
    
    For i = 1 To Len(pszPassword)
        szTmp = Hex(Asc(Mid(pszPassword, i, 1)))
        If Len(szTmp) = 1 Then szTmp = "0" & szTmp
        szResult = szResult & szTmp
    Next i
    Dim nTmp As Integer
    szResult = XOREncrypt(szResult)
    nLen = Len(szResult)
    nTmp = nLen / 3
    szResult = Right(szResult, nLen - nTmp) & Left(szResult, nTmp) '���һ���
    szResult = XOREncrypt(szResult)
    szResult = Right(szResult, nLen - nTmp) & Left(szResult, nTmp) '���һ���
    
    szResult = Right(Format(nPwdLen, "00"), 1) & szResult & Left(Format(nPwdLen, "00"), 1)
    EncryptPassword = szResult
End Function


Private Function XOREncrypt(ByVal pszSource As String) As String
    Const cnPerNum = 3
    Const cnXorValue = 987
    Dim szTmp As String
    Do
        Dim nTmp As Integer
        szTmp = Left(pszSource, cnPerNum)
        If Len(szTmp) < cnPerNum Then   'С�����Χ�ں���
'            szTmp = Hex(HexToDec(szTmp) Xor (cnXorValue And 10 ^ (Len(szTmp)) - 1))
        Else
            szTmp = Hex(HexToDec(szTmp) Xor cnXorValue)
            If Len(szTmp) < cnPerNum Then szTmp = String(cnPerNum - Len(szTmp), "0") & szTmp   '��0
        End If
        XOREncrypt = XOREncrypt & szTmp
        If Len(pszSource) < cnPerNum Then
            Exit Do
        Else
            pszSource = Right(pszSource, Len(pszSource) - cnPerNum)
        End If
    Loop
End Function


Private Function HexCharToDec(pszChar As String) As Integer
    Select Case UCase(pszChar)
        Case "0" To "9"
            HexCharToDec = Val(pszChar)
        Case "A" To "F"
            HexCharToDec = 10 + Asc(pszChar) - Asc("A")
        Case Else
            HexCharToDec = 0
    End Select
End Function
Private Function HexToDec(ByVal pszHex As String) As Long
    Dim i As Integer
    For i = 1 To Len(pszHex)
        HexToDec = HexToDec + 16 ^ (Len(pszHex) - i) * HexCharToDec(Mid(pszHex, i, 1))
    Next i
End Function

