Attribute VB_Name = "mdApiPotevio"
Option Explicit

'�������֤�Ķ������API

'������Ƶ���������ͨ���ֽ���
'iPort����������ʾ�˿ںš��μ�SDT_ResetSAM��  ucByte���޷����ַ���24-255����ʾ��Ƶ���������ͨ���ֽ�����  iIfOpen���������μ�SDT_ResetSAM
'����ֵ��0 x90 �ɹ� ���� ʧ��(���庬��μ��������)
Public Declare Function SDT_SetMaxRFByte Lib "sdtapi.dll" (ByVal iPort As Integer, ByVal ucByte As String, ByVal bIfOpen As Integer) As Integer

'�򿪴���/USB��
'iPort����������ʾ�˿ںš�1-16��ʮ���ƣ�Ϊ���ڣ�1001-1016��ʮ���ƣ�ΪUSB�ڣ�USB�Ķ˿����òο�"USB�豸����ʹ���ֲ�"�� 1001��USB1 1002��USB2
'����ֵ��0 x90 �򿪶˿ڳɹ�   1 �򿪶˿�ʧ��/�˿ںŲ��Ϸ�
Public Declare Function SDT_OpenPort Lib "sdtapi.dll" (ByVal iPort As Integer) As Integer

'�رմ���/USB��
'iPort����������ʾ�˿ںš�
'����ֵ��0 x90 �رն˿ڳɹ�  0 x01 �˿ںŲ��Ϸ�
Public Declare Function SDT_ClosePort Lib "sdtapi.dll" (ByVal iPort As Integer) As Integer

'��ʼ�ҿ�
'����˵����iPort��[in] ��������ʾ�˿ںš��μ�SDT_ResetSAM�� pucIIN��[out] �޷����ַ�ָ�룬ָ�������IIN��iIfOpen��[in] �������μ�SDT_ResetSAM��
'����ֵ��0 x9f �ҿ��ɹ� 0 x80 �ҿ�ʧ��
Public Declare Function SDT_StartFindIDCard Lib "sdtapi.dll" (ByVal iPort As Integer, ByRef pucIIN As String, ByVal iIfOpen As Integer) As Integer

'ѡ��
'����˵����iPort��[in] ��������ʾ�˿ںš��μ�SDT_ResetSAM��pucSN��[out] �޷����ַ�ָ�룬ָ�������SN��iIfOpen��[in] �������μ�SDT_ResetSAM��
'����ֵ��0 x90 ѡ���ɹ�0 x81 ѡ��ʧ��
Public Declare Function SelectIDCard Lib "sdtapi.dll" (ByVal iPort As Integer, ByRef pucSN As String, ByVal iIfOpen As Integer) As Integer

'��ȡID���ڻ�����Ϣ������Ϣ
'����˵����iPort��[in] ��������ʾ�˿ںš��μ�SDT_ResetSAM��pucCHMsg��[out] �޷����ַ�ָ�룬ָ�������������Ϣ��puiCHMsgLen��[out] �޷���������ָ�룬ָ�������������Ϣ���ȡ�pucPHMsg��[out] �޷����ַ�ָ�룬ָ���������Ƭ��Ϣ��puiPHMsgLen��[out] �޷���������ָ�룬ָ���������Ƭ��Ϣ���ȡ�iIfOpen��[in] �������μ�SDT_ResetSAM��
'����ֵ��0 x90 ��������Ϣ�ɹ� ���� ��������Ϣʧ��(���庬��μ��������)
Public Declare Function SDT_ReadBaseMsg Lib "sdtapi.dll" (ByVal iPort As Integer, ByRef pucCHMsg As String, ByRef puiCHMsgLen As Integer, ByRef pucPHMsg As String, ByRef puiPHMsgLen As Integer, ByVal iIfOpen As Integer) As Integer

'��ȡID����IIN , SN��DN
'����˵����iPort��[in] ��������ʾ�˿ںš��μ�SDT_ResetSAM��pucIINSNDN��[out] �޷����ַ�ָ�룬ָ�������IIN,SN��DN,����Ϊ�̶�28�ֽڡ�iIfOpen��[in] �������μ�SDT_ResetSAM��
'����ֵ��0 x90 ��IIN, SN��DN�ɹ����� ��IIN, SN��DNʧ��(���庬��μ��������)
Public Declare Function SDT_ReadIINSNDN Lib "sdtapi.dll" (ByVal iPort As Integer, ByRef pucIINSNDN As String, ByVal iIfOpen As Integer) As Integer

