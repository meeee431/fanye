VERSION 5.00
Object = "{61C3A787-42A5-4F09-9AD8-C9DE75BAD364}#1.0#0"; "STSeatpad.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.2#0"; "RTComctl3.ocx"
Begin VB.Form frmOrderSeats 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����"
   ClientHeight    =   5385
   ClientLeft      =   3900
   ClientTop       =   3840
   ClientWidth     =   7155
   HelpContextID   =   4000100
   Icon            =   "frmOrderSeats.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   7155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin STSeatPad.SeatPad SeatPad1 
      Height          =   2955
      Left            =   120
      TabIndex        =   0
      Top             =   795
      Width           =   6900
      _ExtentX        =   12171
      _ExtentY        =   5212
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "SeatPad2"
      GridNum         =   40
      RowGrids        =   12
   End
   Begin RTComctl3.FlatLabel lblSeat 
      Height          =   435
      Left            =   120
      TabIndex        =   1
      Top             =   4155
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   767
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OutnerStyle     =   2
      HorizontalAlignment=   1
      Caption         =   ""
   End
   Begin RTComctl3.CoolButton cmdCancel 
      Height          =   315
      Left            =   5790
      TabIndex        =   2
      Top             =   4665
      Width           =   1095
      _ExtentX        =   0
      _ExtentY        =   0
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
      MICON           =   "frmOrderSeats.frx":038A
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
      Left            =   4500
      TabIndex        =   3
      Top             =   4665
      Width           =   1095
      _ExtentX        =   0
      _ExtentY        =   0
      BTYPE           =   3
      TX              =   "ȷ��(&O)"
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
      MICON           =   "frmOrderSeats.frx":03A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblTelephone 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2��:13185502533 3��:0575-8605371"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   1785
      TabIndex        =   8
      Top             =   450
      Width           =   4770
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "��ɫ��ʾԤ������ɫ��ʾ�ѳ��۵���Ʊ����ɫ��ʾ�ѳ��۵İ�Ʊ����ɫ��ʾ�ѳ��ۣ���ɫ��ʾ�ƻ�Ԥ������ƱԱ����Ȩ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   780
      Left            =   120
      TabIndex        =   7
      Top             =   4635
      Width           =   4320
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ѡ������λ��:"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   120
      TabIndex        =   6
      Top             =   3795
      Width           =   1560
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��λ�б�(&L):"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   5
      Top             =   465
      Width           =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "˵��:�밴�������������ѡ����λ,�����ո��ѡ����ȡ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6480
   End
End
Attribute VB_Name = "frmOrderSeats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_bOk As Boolean
Public m_rsSeat As Recordset
Public m_szSeat As String
Public m_rsBook As Recordset


Public m_szBookNumber As String
Public m_szSeatNumber As String
Public m_szStatus As Boolean
Dim aszTelePhone() As String '������λ�Ķ�Ӧ�ĵ绰

Public m_szSeatStatus As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Dim i As Integer
    
    m_szBookNumber = ""

'    For i = 1 To SeatPad1.GridNum
'        If SeatPad1.PadGrids.Item(i).BackColor = RGB(255, 0, 0) Then
'            If SeatPad1.PadGrids.Item(i).Value = vbChecked Then
'                m_szBookNumber = InputBox("Ԥ����:", "������Ԥ����", "")
'                Exit For
'            End If
'        End If
'    Next
    
    m_bOk = True
    m_szSeat = lblSeat.Caption
    Unload Me
End Sub

Private Sub Form_Activate()
    SeatPad1.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    Dim oPad As PadGrid
    Dim i As Integer
    Dim j As Integer
    m_bOk = False
    Init
    ReDim aszTelePhone(1 To 1)
    SeatPad1.GridNum = m_rsSeat.RecordCount
    If SeatPad1.GridNum > 0 Then
        m_rsSeat.MoveFirst
        For i = 1 To SeatPad1.GridNum
    
            Set oPad = SeatPad1.PadGrids.Item(i)
            oPad.Caption = m_rsSeat!seat_no
            oPad.Enabled = True
            
            If m_szStatus = True Then
                Select Case m_rsSeat!status
                Case ST_SeatCanSell
                    oPad.Value = vbUnchecked
                Case ST_SeatBooked
                    oPad.Value = vbUnchecked
                    oPad.BackColor = RGB(255, 0, 0)
                Case ST_SeatProjectBooked
                    oPad.Value = vbUnchecked
                    oPad.BackColor = RGB(0, 255, 0)
                Case ST_SeatReserved
                    oPad.BackColor = &H80FF&
                    oPad.Enabled = False
                Case Else
                    If m_rsSeat!ticket_type = TP_PreferentialTicket2 Then '������Ʊ��ɫ��ʾ
                        oPad.BackColor = vbBlue
                    ElseIf m_rsSeat!ticket_type = TP_HalfPrice Then '���۰�Ʊ��ɫ��ʾ
                        oPad.BackColor = RGB(251, 149, 3)
                    Else '��������Ʊ�ֻ�ɫ��ʾ
                        oPad.BackColor = vbYellow
                    End If
                    oPad.Enabled = False
                End Select
            Else
                Select Case m_rsSeat!status
                Case ST_SeatSold
                    oPad.Value = vbUnchecked
                    oPad.BackColor = vbYellow
                Case ST_SeatBooked
                    oPad.Value = vbUnchecked
                    oPad.BackColor = RGB(255, 0, 0)
                Case ST_SeatProjectBooked
                    oPad.Value = vbUnchecked
                    oPad.BackColor = RGB(0, 255, 0)
                Case Else
                    oPad.Enabled = False
                End Select
            End If
            
            m_rsSeat.MoveNext
        Next
        If m_szSeatNumber <> "" Then
          For i = 1 To GetSeatNum(m_szSeatNumber)
            For j = 1 To SeatPad1.GridNum
                Set oPad = SeatPad1.PadGrids.Item(j)
                If oPad.Caption = GetSeatNo(m_szSeatNumber, i) Then oPad.Enabled = False
            Next j
          Next i
        End If
        SeatPad1.Refresh
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
m_szSeatNumber = ""
End Sub

'
Private Sub SeatPad1_GridClick(Index As Integer)
    RefreshSeat
End Sub

Private Sub RefreshSeat()

    Dim oTemp As PadGrid
    Dim szTemp As String
    Dim szTempSeatStatus As String
    Dim aszSeatNo() As String
    Dim szRecord As String
    Dim X As Integer
    Dim y As Integer
    Dim i As Integer
    Dim n As Integer
    Dim nCount As Integer
    Dim nTelNum As Integer 'Ԥ���ĵ绰��
    Dim szTel As String '��ʱ����ĵ绰
    Dim j As Integer
    '�õ�ѡ�����λ
    lblTelephone.Caption = ""
    szTemp = ""
    szTempSeatStatus = ""
    i = 0
    For Each oTemp In SeatPad1.PadGrids
        If oTemp.Value = vbChecked Then
            
            '�õ���λ��״̬�������λΪ��ɫ������ɫ����״ֵ̬Ϊ2������Ϊ0
            If oTemp.BackColor = RGB(255, 0, 0) Or oTemp.BackColor = RGB(0, 255, 0) Then
                szTempSeatStatus = szTempSeatStatus & 2 & ","
            Else
                szTempSeatStatus = szTempSeatStatus & 0 & ","
            End If
        
            szTemp = szTemp & oTemp.Caption & ","
            
            i = i + 1
            ReDim Preserve aszSeatNo(1 To i)
            aszSeatNo(i) = oTemp.Caption
            
        End If
    Next
    
    If szTemp <> "" Then
        lblSeat.Caption = Left(szTemp, Len(szTemp) - 1)
        m_szSeatStatus = Left(szTempSeatStatus, Len(szTempSeatStatus) - 1)
        nCount = i
        ReDim Preserve aszTelePhone(1 To nCount)
        szTel = ""
        nTelNum = 0
        If m_rsBook.RecordCount > 0 Then
            '�õ�ѡ�����λ�ĵ绰����
            For i = 1 To nCount
                
                m_rsBook.MoveFirst
                For j = 1 To m_rsBook.RecordCount
                    If FormatDbValue(m_rsBook!seat_no) = aszSeatNo(i) Then
'                        If szTel <> FormatDbValue(m_rsBook!telephone) Then
                        If BArray(aszTelePhone, FormatDbValue(m_rsBook!telephone)) = False Then
                            szRecord = InputBox("��������λ" & aszSeatNo(i) & "��Ԥ���绰����!", "������绰����")
                            If szRecord <> FormatDbValue(m_rsBook!telephone) Then
                                For X = 1 To SeatPad1.GridNum
                                    If SeatPad1.PadGrids.Item(X).Caption = aszSeatNo(i) Then
                                        SeatPad1.PadGrids.Item(X).Value = vbUnchecked
                                        lblSeat.Caption = ""
                                    End If
                                Next X
                                Exit Sub
                            End If
                        End If
'                        Next
'                            nTelNum = nTelNum + 1
                            szTel = FormatDbValue(m_rsBook!telephone)
                            
                            
'                        End If
                        If BArray(aszTelePhone, FormatDbValue(m_rsBook!telephone)) = False Then
                            aszTelePhone(i) = FormatDbValue(m_rsBook!telephone)
                        End If
'                        aszTelePhone(i, 2) = FormatDbValue(m_rsBook!telephone)
                    End If
                    m_rsBook.MoveNext
                Next j
            Next i
            
            '���绰��䵽��ǩ��
            If NumArray(aszTelePhone) = 1 Then
                lblTelephone.Caption = szTel
            ElseIf NumArray(aszTelePhone) > 1 Then
                '�绰������2��,��˵���ж���˶���,��ֻ��ʾǰ�����ĵ绰����
                lblTelephone.Caption = "�����Ƕ����Ԥ����,��˶�"
            End If
        
        End If
        
        
        
        cmdOk.Enabled = True
        
    Else
        lblSeat.Caption = ""
        m_szSeatStatus = ""
        cmdOk.Enabled = False
    End If
    
    
    
    
    Set oTemp = Nothing
End Sub


Private Sub lstSeat_ItemCheck(Item As Integer)
    RefreshSeat
End Sub
'��ʾHTMLHELP,ֱ�ӿ���
Private Sub DisplayHelp(Optional HelpType As EHelpType = content)
    Dim lActiveControl As Long
    
    Select Case HelpType
        Case content
            lActiveControl = Me.ActiveControl.HelpContextID
            If lActiveControl = 0 Then
                TopicID = Me.HelpContextID
                CallHTMLShowTopicID
            Else
                TopicID = lActiveControl
                CallHTMLShowTopicID
            End If
        Case Index
            CallHTMLHelpIndex
        Case Support
            TopicID = clSupportID
            CallHTMLShowTopicID
    End Select
End Sub
'/////////////////////////////
'�õ�Ԥ����λ��
Private Function GetSeatNum(szSeatNo As String) As Integer
Dim i As Integer
Dim nSeatNo As Integer
Dim szTemp As String
szTemp = szSeatNo
nSeatNo = 0
Do While InStr(szTemp, ",") <> 0
    nSeatNo = nSeatNo + 1
    szTemp = LeftAndRight(szTemp, False, ",")
Loop
nSeatNo = nSeatNo + 1
GetSeatNum = nSeatNo
End Function

Private Sub Init()
    
    lblTelephone.Caption = ""
    If m_szStatus = True Then
        Me.Caption = "����"
        Label4.Caption = "ѡ������λ��:"
    Else
        Me.Caption = "ǿ�г���"
        Label4.Caption = "ǿ�г��۵���λ��:"
    End If
    
End Sub
