Attribute VB_Name = "mdlDisplay"
Option Explicit

#If DISPLAY = 1 Then
'win98��
'Declare Function dsbdll Lib "cky95h.DLL" (ByVal Port As Integer, ByVal OutString As String) As Integer
'win2000����
Declare Function dsbdll Lib "ckyNTh.DLL" (ByVal Port As Integer, ByVal OutString As String) As Integer
#End If
Public g_lComPort As Long
Public Const cszComPort = "ComPort" 'Com�˿��ַ���
        

'���ݶ�վ��ʾ�����
'1.��ӭ����XXXX��վ ����XXXΪ������ �����У�
'2.��վ��XXXX ����ʱ��:yyyy-mm-dd��������                 ����дһ�У�
'3.������XXXX ��������
'4.Ӧ�ң�XXXX ��������
'5.ת1



Public Sub SetInit()
    '��ʼ��
    
#If DISPLAY = 1 Then
    dsbdll g_lComPort, "f"
    
#End If
    
End Sub

Public Sub SetClear(pnLine As Integer)
    '��յڼ���
#If DISPLAY = 1 Then
    dsbdll g_lComPort, "$" & pnLine
    
#End If
End Sub

Public Sub SetUser(pnUserID As String)
    '��ʾ����
#If DISPLAY = 1 Then
    SetClear 1
    dsbdll g_lComPort, "# �񻷿������� ��ӭ��#"
    SetClear 2
    dsbdll g_lComPort, "#   ����:" & pnUserID & "Ϊ������#"
#End If
End Sub

Public Sub SetStationAndTime(pszStation As String, pszTime As String, Optional pnTicketNum As Integer = 0)
    '��վ��ʱ��
#If DISPLAY = 1 Then
    SetClear 1
    
    dsbdll g_lComPort, "#��վ��" & pszStation & IIf(pnTicketNum <> 0, " " & pnTicketNum & "��#", "#")
    SetClear 2
    dsbdll g_lComPort, "#ʱ�䣺" & Format(pszTime, "MM-dd hh:mm") & "#"
#End If
End Sub


Public Sub SetReceive(pdbMoney As Double)
    'ʵ��
#If DISPLAY = 1 Then

    SetClear 2
    dsbdll g_lComPort, pdbMoney & "Y"
#End If
End Sub


Public Sub SetPay(pdbMoney As Double, pszStation As String, pszTime As String, Optional pnTicketNum As Integer = 0, Optional pdbInsurance As Double)
    '�븶��
#If DISPLAY = 1 Then
    SetClear 1
    'dsbdll g_lComPort, "#��վ��" & pszStation & IIf(pnTicketNum <> 0, IIf(Len(pszStation) < 3, " ", "") & pnTicketNum & "��", "") & " " & Format(pszTime, "mm��dd��") & "#"
    dsbdll g_lComPort, "#" & pszStation & IIf(pnTicketNum <> 0, IIf(Len(pszStation) < 3, "", "") & pnTicketNum & "��", "") & "" & Format(pszTime, "mm��dd��hh:mm") & "#"
    SetClear 2
    dsbdll g_lComPort, pdbMoney & "J"
    
    If pdbInsurance > 0 Then
        dsbdll g_lComPort, "t"
        dsbdll g_lComPort, pdbInsurance & "E"
    End If
#End If
End Sub

Public Sub SetPay2(pdbMoney As Double)
    '�븶��
#If DISPLAY = 1 Then
    SetClear 2
    dsbdll g_lComPort, pdbMoney & "J"
#End If
End Sub

Public Sub SetReturn(pdbMoney As Double)
    'Ӧ��
#If DISPLAY = 1 Then
    SetClear 2
    '����������ʾлл
    'dsbdll g_lComPort, "X"
    '��������ʾ���
    dsbdll g_lComPort, pdbMoney & "Z"
#End If
End Sub


Public Sub SetThanks()
    '��ʾлл
#If DISPLAY = 1 Then
    SetClear 2
    dsbdll g_lComPort, "X"
#End If
End Sub

Public Sub SetCal()
    '��ʾ �����뵱����壬лл
    
#If DISPLAY = 1 Then
    SetClear 2
    dsbdll g_lComPort, "D"
#End If
End Sub


Public Sub SetWait()
    '��ʾ�ȴ�
    
#If DISPLAY = 1 Then
    dsbdll g_lComPort, "W"
#End If
End Sub

Public Sub SetTicketNum()
    '��ʾҪ����
#If DISPLAY = 1 Then
    SetClear 2
    
    dsbdll g_lComPort, "c"
#End If
    
End Sub

Public Sub SetWhere()
    '��ʾҪȥ�Ķ�
#If DISPLAY = 1 Then
    SetClear 2
    dsbdll g_lComPort, "a"
#End If
End Sub

Public Sub SetCheck()
    '��ʾ�����˶�һ��
#If DISPLAY = 1 Then
    SetClear 1
    dsbdll g_lComPort, "h"
#End If
End Sub

Public Sub SetQueue()
    '��ʾ���Ŷӹ�Ʊ,лл����
#If DISPLAY = 1 Then
    SetClear 2
    dsbdll g_lComPort, "k"
#End If
End Sub


Public Sub SetProtect()
    '��ʾ��ע�Ᵽ��������Ʒ
#If DISPLAY = 1 Then
    SetClear 2
    dsbdll g_lComPort, "i"
    
#End If
    
End Sub





Public Sub SetInsurance()
    '��ʾ����
'#If DISPLAY = 1 Then
''    SetClear 1
''    dsbdll g_lComPort, "#������Ը���������֪#"
'    dsbdll g_lComPort, "l"
'    SetUser m_oAUser.UserID
'#End If
End Sub
