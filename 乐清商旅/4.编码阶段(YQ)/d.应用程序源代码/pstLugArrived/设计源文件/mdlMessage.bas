Attribute VB_Name = "mdlMessage"
Option Explicit

'Const SM_ID_LEN = 23
'Const SM_ORG_NO_LEN = 21
'Const SM_DST_NO_LEN = 21
'Const SM_FEE_NO_LEN = 21
'Const SM_SEND_TIME_LEN = 19
'Const SM_CONTEXT_LEN = 140
'Const BUSINESS_ID_LEN = 10
'Const SM_MAX_DST_NO_LEN = (SM_DST_NO_LEN + 1) * 50
'Const FEE_TYPE_LEN = 10
'Const FEE_SUM_LEN = 6
'Const LINK_ID_LEN = 20
'Const VER_LEN = 3
'Const CMD_WORD_LEN = 5 '�����ֵĳ���(����β��'\0')
'
Public Const IMAPI_SUCC = 0               ' �����ɹ�
Public Const IMAPI_CONN_ERR = -1          ' �������ݿ����
Public Const IMAPI_CONN_CLOSE_ERR = -2    ' ���ݿ�ر�ʧ��
Public Const IMAPI_INS_ERR = -3           ' ���ݿ�������
Public Const IMAPI_DEL_ERR = -4           ' ���ݿ�ɾ������
Public Const IMAPI_QUERY_ERR = -5         ' ���ݿ��ѯ����
Public Const IMAPI_DATA_ERR = -6          ' ���ݴ���
Public Const IMAPI_API_ERR = -7           ' API���벻����
Public Const IMAPI_DATA_TOOLONG = -8      ' ����̫��
Public Const IMAPI_INIT_ERR = -9          ' û�г�ʼ�����ʼ��ʧ��

Public Const SM_ID_LEN = 8                ' ����ID����󳤶�(0-99999999)
Public Const SM_MOBILE_LEN = 16           ' �ֻ�������󳤶�
Public Const SM_CONTEXT_LEN = 260         ' ����������󳤶�
Public Const SM_RPT_LEN = 100             ' ���Ż�ִ��������󳤶�

Public Type MOItem
    smMobile(SM_MOBILE_LEN - 1) As Byte
    smContent(SM_CONTEXT_LEN - 1) As Byte
    smID As Long
End Type

Public Type RptItem
    rptMobile(SM_MOBILE_LEN - 1) As Byte
    smID As Long
    rptId As Long
    rptDesc(SM_RPT_LEN - 1) As Byte
End Type


Public Declare Function init Lib "ImApi.dll" (ByVal ip As String, ByVal userName As String, ByVal password As String, ByVal apiCode As String) As Long
'�ͷ�
Public Declare Function release Lib "ImApi.dll" () As Long
'������Ϣ
Public Declare Function sendSM Lib "ImApi.dll" (ByVal mobile As String, ByVal content As String, ByVal smID As Long) As Long
'����Wap Push��Ϣ
Public Declare Function sendWapPushSM Lib "ImApi.dll" (ByVal mobile As String, ByVal content As String, ByVal smID As Long, ByVal mobile As String) As Long
'���ն��ţ����ز�ѯ���Ķ���������ɾ����Щ����
Public Declare Function receiveSM Lib "ImApi.dll" (ByRef MOItems As MOItem, ByVal retsize As Long) As Long
'���ջ�ִ�����ز�ѯ���Ļ�ִ������ɾ����Щ��ִ
Public Declare Function receiveRPT Lib "ImApi.dll" (ByRef RptItems As RptItem, ByVal retsize As Long) As Long
                         
