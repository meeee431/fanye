Attribute VB_Name = "mdCodeCommon"
''Public Const SUCCESS = 100 '�½����ܹ��û��ųɹ�
'''// 1. ���Ҽ���������ӵ� Rockey2 �������豸
''Declare Function RY2_Find Lib "ROCKEY2.DLL" () As Long
'''// ����:
'''// �������ֵС�� 0����ʾ���ص���һ��������룬���庬���뿴����Ĵ�����롣
'''// �������ֵΪ 0����ʾû���κ� Rockey2 �豸���ڡ�
'''// �������ֵ���� 0�����ص����ҵ��� Rockey2 �������ĸ�����
'''// ====================================================================
'''
'''// 2. ��ָ���ļ�����
''
''Declare Function RY2_Open Lib "ROCKEY2.DLL" (ByVal mode As Long, ByVal uid As Long, ByRef hid As Long) As Long
''
'''// ����:
'''// mode �Ǵ򿪼������ķ�ʽ
'''// mode = 0 ��ʾ���Ǵ򿪵�1���ҵ��� Rockey2������� UID �� HID
'''// mode > 0 ��ʾ���ǰ� UID ��ʽ�򿪣���ʱ�� mode ��ֵ��ʾ����Ҫ���������
'''//          Ʃ��: uid=12345678, mode=2 ��ʾ����ϣ���� uid ����12345678 ��
'''//          ��2�Ѽ�������
'''// mode = -1 ��ʾ���ǰ� HID ��ʽ�򿪣�Ҫ�� *hid ����Ϊ 0
''Global Const AUTO_MODE = 0
''
''Global Const HID_MODE = -1
''
'''// uid(UserID)
'''// ���� UID ��ʽ�򿪵�ʱ���������Ҫ�򿪼������� UID���� UID �����û�����
'''// RY2_GenUID ���ܻ�õ��û� ID��
'''// hid
'''// ����Ǽ�������Ӳ�� ID������1������/���ֵ������� HID ��ʽ�򿪼�������
'''// ʱ�򣬱�ʾϣ����Ӳ��ID = *hid �ļ�������
'''// ���������ַ�ʽ�򿪼��������ڳɹ��򿪼������Ժ�����ⷵ�ؼ�������Ӳ�� ID
'''// ����ֵ:
'''// ���ڵ���0        ��ʾ�ɹ������صľ��Ǵ򿪵ļ������ľ��
'''// С��0            ���ص���һ��������룬���庬���뿴����Ĵ�����벿�֡�
''
'''// ====================================================================
'''
'''// 3. �ر�ָ���ļ�����
''Declare Sub RY2_Close Lib "ROCKEY2.DLL" (ByVal handle As Long)
'''// ����:
'''// handle �豸�ľ������ RY2_Open ����ص� handle һ�¡�
'''// ����:
'''// ���ص���һ��������룬���庬���뿴����Ĵ�����벿�֡�
''
'''// ====================================================================
'''
'''// 4. �����û� ID
''Declare Function RY2_GenUID Lib "ROCKEY2.DLL" (ByVal handle As Long, ByRef uid As Long, ByVal seed As Any, ByVal isProtect As Long) As Long
'''// ����:
'''// handle �豸�ľ������ RY2_Open ����ص� handle һ�¡�
'''// uid ������������ɵ��û� ID �Ӵ˲�������
'''// seed �û����������������û� ID �����ӣ�����һ����󳤶ȿ����� 64 ���ֽڵ��ַ���
'''// isProtect ������Ϊ0
'''// ����:
'''// ���ص���һ��������룬���庬���뿴����Ĵ�����벿�֡�
'''
'''// ====================================================================
''
'''// 5. ��ȡ����������
''
''Declare Function RY2_Read Lib "ROCKEY2.DLL" (ByVal handle As Long, ByVal block_index As Integer, ByVal buffer512 As Any) As Long
'''// ����:
'''// handle �豸�ľ������ RY2_Open ����ص� handle һ�¡�
'''// block_index ��������ָ��Ҫ��ȡ������1���飬ȡֵΪ(0-4)
'''// buffer512 ������Ļ���������Ϊÿ����ĳ��ȹ̶�Ϊ 512 ���ֽڣ��������
'''// buffer �Ĵ�С������ 512 ���ֽ�
'''// ����:
'''// ���ص���һ��������룬���庬���뿴����Ĵ�����벿�֡�
''
'''// ====================================================================
'''
'''// 6. д�����������
''
''Declare Function RY2_Write Lib "ROCKEY2.DLL" (ByVal handle As Long, ByVal block_index As Integer, ByVal buffer512 As Any) As Long
''
'''// ����:
'''// handle �豸�ľ������ RY2_Open ����ص� handle һ�¡�
'''// block_index ��������ָ��Ҫд�������1���飬ȡֵΪ(0-4)
'''// buffer512 д���Ļ���������Ϊÿ����ĳ��ȹ̶�Ϊ 512 ���ֽڣ��������
'''// buffer �Ĵ�С������ 512 ���ֽ�
'''// ����:
'''// ���ص���һ��������룬���庬���뿴����Ĵ�����벿�֡�
'''
'''
'''// ������� ===========================================================
'''
'''// �ɹ���û�д���
''Global Const RY2ERR_SUCCESS = 0
'''
'''// û���ҵ�����Ҫ����豸(��������)
''Global Const RY2ERR_NO_SUCH_DEVICE = &HA0100001
'''
'''// �ڵ��ô˹���ǰ��Ҫ�ȵ��� RY2_Open ���豸(��������)
''Global Const RY2ERR_NOT_OPENED_DEVICE = &HA0100002
'''
'''// ������ UID ����(��������)
''Global Const RY2ERR_WRONG_UID = &HA0100003
'''
'''// ��д���������Ŀ���������(��������)
''Global Const RY2ERR_WRONG_INDEX = &HA0100004
'''
'''// ���� GenUID ���ܵ�ʱ�򣬸����� seed �ַ������ȳ����� 64 ���ֽ�(��������)
''Global Const RY2ERR_TOO_LONG_SEED = &HA0100005
'''
'''// ��ͼ��д�Ѿ�д������Ӳ��(��������)
''Global Const RY2ERR_WRITE_PROTECT = &HA0100006
'''
'''// ���豸��(Windows ����)
''Global Const RY2ERR_OPEN_DEVICE = &HA0100007
'''
'''// ����¼��(Windows ����)
''Global Const RY2ERR_READ_REPORT = &HA0100008
'''
'''// д��¼��(Windows ����)
''Global Const RY2ERR_WRITE_REPORT = &HA0100009
'''
'''// �ڲ�����(Windows ����)
''Global Const RY2ERR_SETUP_DI_GET_DEVICE_INTERFACE_DETAIL = &HA010000A
'''
'''// �ڲ�����(Windows ����)
''Global Const RY2ERR_GET_ATTRIBUTES = &HA010000B
'''
'''// �ڲ�����(Windows ����)
''Global Const RY2ERR_GET_PREPARSED_DATA = &HA010000C
'''
'''// �ڲ�����(Windows ����)
''Global Const RY2ERR_GETCAPS = &HA010000D
'''
'''// �ڲ�����(Windows ����)
''Global Const RY2ERR_FREE_PREPARSED_DATA = &HA010000E
'''
'''// �ڲ�����(Windows ����)
''Global Const RY2ERR_FLUSH_QUEUE = &HA010000F
'''
'''// �ڲ�����(Windows ����)
''Global Const RY2ERR_SETUP_DI_CLASS_DEVS = &HA0100010
'''
'''// �ڲ�����(Windows ����)
''Global Const RY2ERR_GET_SERIAL = &HA0100011
'''
'''// �ڲ�����(Windows ����)
''Global Const RY2ERR_GET_PRODUCT_STRING = &HA0100012
'''
'''// �ڲ�����
''Global Const RY2ERR_TOO_LONG_DEVICE_DETAIL = &HA0100013
'''
'''// δ֪���豸(Ӳ������)
''Global Const RY2ERR_UNKNOWN_DEVICE = &HA0100020
'''
'''// ������֤����(Ӳ������)
''Global Const RY2ERR_VERIFY = &HA0100021
'''
'''// δ֪����(Ӳ������)
''Global Const RY2ERR_UNKNOWN_ERROR = &HA010FFFF
'''
''
