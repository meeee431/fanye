VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D11DF87C-EFEC-4838-B7E2-15462BF136FB}#2.1#0"; "RTReportLF.ocx"
Object = "{286DDD69-C676-405C-800F-55A9C4853C35}#1.1#0"; "RTComctl3.ocx"
Begin VB.Form frmSumAccept 
   BackColor       =   &H00E0E0E0&
   Caption         =   "����ͳ��"
   ClientHeight    =   5925
   ClientLeft      =   2100
   ClientTop       =   2820
   ClientWidth     =   9735
   HelpContextID   =   7000110
   Icon            =   "frmSumAccept.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5925
   ScaleWidth      =   9735
   WindowState     =   2  'Maximized
   Begin RTReportLF.RTReport RTReport1 
      Height          =   4815
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   8493
   End
   Begin VB.Frame fraCondition 
      BackColor       =   &H00E0E0E0&
      Height          =   705
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   9615
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   315
         Left            =   4350
         TabIndex        =   3
         Top             =   240
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   69992451
         UpDown          =   -1  'True
         CurrentDate     =   37062.6506944444
      End
      Begin MSComCtl2.DTPicker dtpStart 
         Height          =   315
         Left            =   1230
         TabIndex        =   1
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   69992451
         UpDown          =   -1  'True
         CurrentDate     =   37062.6506944444
      End
      Begin RTComctl3.CoolButton cmdStat 
         Default         =   -1  'True
         Height          =   345
         Left            =   6360
         TabIndex        =   4
         Top             =   210
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "ͳ��(&Q)"
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
         MICON           =   "frmSumAccept.frx":038A
         PICN            =   "frmSumAccept.frx":03A6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin RTComctl3.CoolButton cmdClose 
         Cancel          =   -1  'True
         Height          =   345
         Left            =   7770
         TabIndex        =   5
         Top             =   210
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "�ر�(&C)"
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
         MICON           =   "frmSumAccept.frx":0740
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ʼʱ��(&S):"
         Height          =   180
         Index           =   0
         Left            =   150
         TabIndex        =   0
         Top             =   300
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ʱ��(&E):"
         Height          =   180
         Left            =   3240
         TabIndex        =   2
         Top             =   300
         Width           =   1080
      End
   End
End
Attribute VB_Name = "frmSumAccept"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Const cszTemplateFile = "�а�Աÿ�ս���_��Ʊ��ϸ.xls"

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdStat_Click()
    On Error GoTo Error_Handle
    Dim lHelpContextID As Long

    Dim rsSellDetail As Recordset
    Dim rsDetailToShow As Recordset
    Dim adbOther() As Double
    Dim oCalculator As New LugDataStatSvr
    Dim i As Integer, nUserCount As Integer
    
    Dim szLastTicketID As String
    Dim szBeginTicketID As String
    Dim arsData() As Recordset, vaCostumData As Variant
    
'        Dim lFullnumber As Long, lHalfnumber As Long, lFreenumber As Long
'        Dim dbFullAmount As Double, dbHalfAmount As Double, dbFreeAmount As Double
    Dim alNumber As Long '����Ʊ�ֵ�����
    Dim adbAmount As Double  '����Ʊ�ֵĽ��
    Dim j As Integer
    Dim aszAllSeller() As String
    Dim nAllSeller As Integer
    Dim k As Integer
    'Dim l As Integer
    Dim adbPriceItem() As Double 'Ʊ������ϸ

    Dim nTicketNumberLen As Integer
    Dim nTicketPrefixLen As Integer
    nTicketNumberLen = moSysParam.LuggageIDNumberLen
    nTicketPrefixLen = moSysParam.LuggageIDPrefixLen
    
    oCalculator.Init m_oAUser
    

    
    nUserCount = 1

    
    ReDim arsData(1 To nUserCount)
    ReDim vaCostumData(1 To nUserCount, 1 To 22, 1 To 2)
'            SetProgressRange nUserCount, "�����γɼ�¼��..."
    
    For i = 1 To nUserCount
'                    For j = 1 To TP_TicketTypeCount
            alNumber = 0
            adbAmount = 0
'                    Next j
        
            Set rsSellDetail = oCalculator.AcceptEveryDaySellDetail(m_oAUser.UserID, dtpStart.Value, dtpEnd.Value)
            Set rsDetailToShow = New Recordset
            With rsDetailToShow.Fields
                .Append "ticket_id_range", adChar, 30
                '����¼�������ÿ��Ʊ�ֵ����������ֶ�
            
                .Append "number_ticket", adInteger
                .Append "amount_ticket", adCurrency
                
            End With
            
            rsDetailToShow.Open

            
            Do While Not rsSellDetail.EOF
                If rsDetailToShow.RecordCount = 0 Or Not IsTicketIDSequence(szLastTicketID, RTrim(rsSellDetail!luggage_id), nTicketNumberLen, nTicketPrefixLen) Then
                    If rsDetailToShow.RecordCount <> 0 Then
                        rsDetailToShow!ticket_id_range = szBeginTicketID & "---" & szLastTicketID
                        
                    
                        alNumber = alNumber + rsDetailToShow("number_ticket")
                        adbAmount = adbAmount + rsDetailToShow("amount_ticket")
                        
                    End If

                    szBeginTicketID = RTrim(rsSellDetail!luggage_id)
                    rsDetailToShow.AddNew
                End If
                rsDetailToShow("number_ticket") = rsDetailToShow("number_ticket") + 1
                rsDetailToShow("amount_ticket") = rsDetailToShow("amount_ticket") + rsSellDetail!price_total
                
                szLastTicketID = RTrim(rsSellDetail!luggage_id)
                
                rsSellDetail.MoveNext
            Loop
            
            If rsSellDetail.RecordCount > 0 Then
                rsSellDetail.MoveLast
                rsDetailToShow!ticket_id_range = szBeginTicketID & "---" & RTrim(rsSellDetail!luggage_id)
'                        For j = 1 To TP_TicketTypeCount
                alNumber = alNumber + rsDetailToShow("number_ticket")
                adbAmount = adbAmount + rsDetailToShow("amount_ticket")
'                        Next j

'                        rsDetailToShow.AddNew
                
'                        rsDetailToShow!ticket_id_range = "�ϼ�"
'                        For j = 1 To TP_TicketTypeCount
'                        rsDetailToShow("number_ticket") = alNumber
'                        rsDetailToShow("amount_ticket") = adbAmount
'                        Next j
'                        rsDetailToShow.Update
            End If
            vaCostumData(i, 22, 1) = "Ʊ�Ŷ�"
            
            If rsDetailToShow.RecordCount > 0 Then rsDetailToShow.MoveFirst
            For j = 1 To rsDetailToShow.RecordCount
                vaCostumData(i, 22, 2) = vaCostumData(i, 22, 2) & rsDetailToShow!ticket_id_range & "   "
                rsDetailToShow.MoveNext
            Next j
            
            Set arsData(i) = rsDetailToShow
            adbOther = oCalculator.AcceptEveryDayAnotherThing(ResolveDisplay(m_oAUser.UserID), dtpStart.Value, dtpEnd.Value)
            vaCostumData(i, 1, 1) = "�ϵ���"
            vaCostumData(i, 1, 2) = CInt(adbOther(1, 1)) & " ��  Ʊ��=" & adbOther(1, 2) & " Ԫ"
            
            vaCostumData(i, 2, 1) = "�˵���"
            vaCostumData(i, 2, 2) = CInt(adbOther(2, 1)) & " ��  Ʊ��=" & adbOther(2, 2) & " Ԫ  ������=" & adbOther(2, 3) & " Ԫ"
            
'                    vaCostumData(i, 3, 1) = "��ǩ"
'                    vaCostumData(i, 3, 2) = "����=" & CInt(adbOther(3, 1)) & " ��  Ʊ��=" & adbOther(3, 2) & " Ԫ  ������=" & adbOther(3, 3) & " Ԫ"
'
'            Dim dbAmount As Double '��������Ʊ
'            Dim lNumber As Long '������Ʊ
''                    lNumber = 0
''                    dbAmount = 0
''                    For j = 1 To TP_TicketTypeCount
''                        If j <> TP_FreeTicket Then
'            dbAmount = adbAmount
''                        End If
'            lNumber = alNumber
'                    Next j
            
            
            Dim dbAmount As Double '��������Ʊ
            Dim lNumber As Long '������Ʊ
            dbAmount = oCalculator.AcceptEveryDaySellTotal(ResolveDisplay(m_oAUser.UserID), dtpStart.Value, dtpEnd.Value)

            lNumber = alNumber
                
            vaCostumData(i, 4, 1) = "Ӧ����"
            vaCostumData(i, 4, 2) = dbAmount - adbOther(1, 2) - adbOther(2, 2) + adbOther(2, 3) & " Ԫ"
            
            vaCostumData(i, 5, 1) = "������"
            'vaCostumData(i, 5, 2) = lNumber & " ��"
            vaCostumData(i, 5, 2) = lNumber + adbOther(1, 1) + adbOther(2, 1) & " ��"
            
            vaCostumData(i, 6, 1) = "����Ʊ����"
            'vaCostumData(i, 6, 2) = lNumber - adbOther(1, 1) - adbOther(2, 1) & " ��"
            vaCostumData(i, 6, 2) = lNumber & " ��"
                 
            vaCostumData(i, 7, 1) = "�Ƶ�"
            vaCostumData(i, 7, 2) = MakeDisplayString(m_oAUser.UserID, m_oAUser.UserName)
            
            vaCostumData(i, 8, 1) = "����"
            vaCostumData(i, 8, 2) = ""
            
            vaCostumData(i, 9, 1) = "����Ա"
            vaCostumData(i, 9, 2) = m_oAUser.UserID
            
            vaCostumData(i, 10, 1) = "��������"
            vaCostumData(i, 10, 2) = Format(dtpStart.Value, "YYYY-MM-DD hh:mm:ss") & "��" & Format(dtpEnd.Value, "YYYY-MM-DD hh:mm:ss")
            
            
            adbPriceItem = oCalculator.GetAccepterPriceDetail(m_oAUser.UserID, dtpStart.Value, dtpEnd.Value)
            Dim nPrice As Integer
            
            For j = 1 To 10
                vaCostumData(i, 10 + j, 1) = "price_item_" & j
                vaCostumData(i, 10 + j, 2) = adbPriceItem(j)
            Next j
            vaCostumData(i, 21, 1) = "ticket_price_total"
            vaCostumData(i, 21, 2) = adbPriceItem(j)
            
'                    For j = 1 To arsData(i).RecordCount
'
'                    Next j
            
'                End If
'                SetProgressValue i
        
        
        
    Next
    Dim arsTemp As Variant
    Dim aszTemp As Variant
'    Dim rsTemp As Recordset
    ReDim aszTemp(1 To 1)
    ReDim arsTemp(1 To 1)
    '��Ʊ��
    aszTemp(1) = "���˷���"
    Set arsTemp(1) = g_rsPriceItem
'    m_bNeedSave = True
'    m_nReportType = pnReportType
'    Me.Caption = pszCaption
    
    WriteProcessBar True, , , "�����γɱ���..."
    
    RTReport1.CustomStringCount = aszTemp
    RTReport1.CustomString = arsTemp
    RTReport1.LeftLabelVisual = True
    RTReport1.TopLabelVisual = True
    RTReport1.TemplateFile = App.Path & "\" & cszTemplateFile
    RTReport1.ShowMultiReport arsData, vaCostumData

    
    SetNormal
    Exit Sub
Error_Handle:
    SetNormal
    ShowErrorMsg
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyF1 Then
        DisplayHelp Me
    End If
End Sub

Private Sub Form_Load()
    dtpStart.Value = Date
    dtpEnd.Value = DateAdd("s", -1, DateAdd("d", 1, dtpStart.Value))
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    fraCondition.Width = Me.ScaleWidth
    RTReport1.Width = fraCondition.Width
    RTReport1.Height = Me.ScaleHeight - fraCondition.Height - 60
End Sub



Private Function IsTicketIDSequence(pszFirstTicketID As String, pszSecondTicketID As String, nTicketNumberLen As Integer, nTicketPrefixLen As Integer) As Boolean
    Dim szTemp1 As String, szTemp2 As String
    On Error GoTo Error_Handle
    szTemp1 = UCase(Left(pszFirstTicketID, nTicketPrefixLen))
    szTemp2 = UCase(Left(pszSecondTicketID, nTicketPrefixLen))
    If szTemp1 = szTemp2 Then
        szTemp1 = Right(pszFirstTicketID, nTicketNumberLen)
        szTemp2 = Right(pszSecondTicketID, nTicketNumberLen)
        If CLng(szTemp1) + 1 = CLng(szTemp2) Then
            IsTicketIDSequence = True
        End If
    End If
    Exit Function
Error_Handle:
End Function



