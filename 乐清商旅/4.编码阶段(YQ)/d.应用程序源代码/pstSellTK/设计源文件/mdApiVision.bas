Attribute VB_Name = "mdApiVision"
Option Explicit

'�������֤�Ķ������API

'ע�⣺�����ò�ѯ��ʽ�Զ��жϿ�Ƭ�Ƿ���ã�����ʱ�佨�����300ms

'��ʼ������
'������Port�����Ӵ��ڣ�COM1~COM16����USB��(1001~1016)
Public Declare Function CVR_InitComm Lib "termb.dll" (ByVal Port As Long) As Integer

'�ر�����
Public Declare Function CVR_CloseComm Lib "termb.dll" () As Integer

'��Ƶ����
Public Declare Function CVR_Authenticate Lib "termb.dll" () As Integer

'��������
'������Active
'1����������Ϣ ��������WZ.TXT?��Ƭ����XP.WLT����ƬZP.BMP(����)
'2�� ��������Ϣ ��������WZ.TXT����Ƭ����XP.WLT
'3��������סַ��Ϣ ��������סַNEWADD.TXT(�������µ�ַ�����ɿ��ļ�)
'4����������Ϣ  ����WZ.TXT(����)����ƬZP.BMP(����)
'5����оƬ����� оƬ�����IINSNDN.bin
'6����������Ϣ  ���豸Ψһ��־�ţ���������WZ.TXT(����)����ƬXP.BMP(����)�������ն����绷����
Public Declare Function CVR_Read_Content Lib "termb.dll" (ByVal Active As Long) As Integer

'����֤
Public Declare Function CVR_Ant Lib "termb.dll" (ByVal mode As Long) As Integer

'��������(�����ڴ�)
'������
'pucCHMsg�����������Ϣ�ڴ滺��ָ��/����Out
'puiCHMsgLen�����������Ϣ����/Ĭ�� 256 Byte
'pucPHMsg�������Ƭ��Ϣ�ڴ滺��ָ��/����Out
'puiPHMsgLen�������Ƭ��Ϣ����/Ĭ�� 1024 Byte
'nMode���������1�����ֱ���ΪĬ��UCS-2��ʽ����Ƭδ��ѹ��bmp�ļ�  �������2�����ֱ�����ת����GBK�������ʽ����Ƭδ��ѹ��bmp�ļ�  �������3�����ֱ���ΪĬ��UCS-2��ʽ����Ƭ�ѽ�ѹ��zp.bmp�ļ�  �������4�����ֱ�����ת����GBK�������ʽ����Ƭ�ѽ�ѹ��zp.bmp�ļ�
Public Declare Function CVR_ReadBaseMsg Lib "termb.dll" (ByVal pucCHMsg As String, ByRef puiCHMsgLen As Integer, ByVal pucPHMsg As String, ByRef puiPHMsgLen As Integer, ByRef nMode As Integer) As Integer


'˵ �������º�����������Ϊ������CVR_Read_Content ����CVR_ReadBaseMsg�������ɹ����ٷֱ�������Ϻ�����CVR_Read_Content����CVR_ReadBaseMsg�����Զ���Ӧ�ó���ǰĿ¼����BMP��Ƭ�ļ���

'�õ�������Ϣ
Public Declare Function GetPeopleName Lib "termb.dll" (ByVal lpReturnedString As String, ByRef nReturnLen As Integer) As Integer

'�õ��Ա���Ϣ
Public Declare Function GetPeopleSex Lib "termb.dll" (ByVal lpReturnedString As String, ByRef nReturnLen As Integer) As Integer

'�õ�������Ϣ
Public Declare Function GetPeopleNation Lib "termb.dll" (ByVal lpReturnedString As String, ByRef nReturnLen As Integer) As Integer

'�õ���������
Public Declare Function GetPeopleBirthday Lib "termb.dll" (ByVal lpReturnedString As String, ByRef nReturnLen As Integer) As Integer

'�õ���ַ��Ϣ
Public Declare Function GetPeopleAddress Lib "termb.dll" (ByVal lpReturnedString As String, ByRef nReturnLen As Integer) As Integer

'�õ����֤����Ϣ
Public Declare Function GetPeopleIDCode Lib "termb.dll" (ByVal lpReturnedString As String, ByRef nReturnLen As Integer) As Integer

'�õ���֤������Ϣ
Public Declare Function GetDepartment Lib "termb.dll" (ByVal lpReturnedString As String, ByRef nReturnLen As Integer) As Integer

'�õ���Ч��ʼ����
Public Declare Function GetStartDate Lib "termb.dll" (ByVal lpReturnedString As String, ByRef nReturnLen As Integer) As Integer

'�õ���Ч��ֹ����
Public Declare Function GetEndDate Lib "termb.dll" (ByVal lpReturnedString As String, ByRef nReturnLen As Integer) As Integer
