Attribute VB_Name = "mdApiPotevio"
Option Explicit

'普天身份证阅读器相关API

'设置射频适配器最大通信字节数
'iPort：整数，表示端口号。参见SDT_ResetSAM。  ucByte：无符号字符，24-255，表示射频适配器最大通信字节数。  iIfOpen：整数，参见SDT_ResetSAM
'返回值：0 x90 成功 其他 失败(具体含义参见返回码表)
Public Declare Function SDT_SetMaxRFByte Lib "sdtapi.dll" (ByVal iPort As Integer, ByVal ucByte As String, ByVal bIfOpen As Integer) As Integer

'打开串口/USB口
'iPort：整数，表示端口号。1-16（十进制）为串口，1001-1016（十进制）为USB口，USB的端口设置参看"USB设备配置使用手册"。 1001：USB1 1002：USB2
'返回值：0 x90 打开端口成功   1 打开端口失败/端口号不合法
Public Declare Function SDT_OpenPort Lib "sdtapi.dll" (ByVal iPort As Integer) As Integer

'关闭串口/USB口
'iPort：整数，表示端口号。
'返回值：0 x90 关闭端口成功  0 x01 端口号不合法
Public Declare Function SDT_ClosePort Lib "sdtapi.dll" (ByVal iPort As Integer) As Integer

'开始找卡
'参数说明：iPort：[in] 整数，表示端口号。参见SDT_ResetSAM。 pucIIN：[out] 无符号字符指针，指向读到的IIN。iIfOpen：[in] 整数，参见SDT_ResetSAM。
'返回值：0 x9f 找卡成功 0 x80 找卡失败
Public Declare Function SDT_StartFindIDCard Lib "sdtapi.dll" (ByVal iPort As Integer, ByRef pucIIN As String, ByVal iIfOpen As Integer) As Integer

'选卡
'参数说明：iPort：[in] 整数，表示端口号。参见SDT_ResetSAM。pucSN：[out] 无符号字符指针，指向读到的SN。iIfOpen：[in] 整数，参见SDT_ResetSAM。
'返回值：0 x90 选卡成功0 x81 选卡失败
Public Declare Function SelectIDCard Lib "sdtapi.dll" (ByVal iPort As Integer, ByRef pucSN As String, ByVal iIfOpen As Integer) As Integer

'读取ID卡内基本信息区域信息
'参数说明：iPort：[in] 整数，表示端口号。参见SDT_ResetSAM。pucCHMsg：[out] 无符号字符指针，指向读到的文字信息。puiCHMsgLen：[out] 无符号整型数指针，指向读到的文字信息长度。pucPHMsg：[out] 无符号字符指针，指向读到的照片信息。puiPHMsgLen：[out] 无符号整型数指针，指向读到的照片信息长度。iIfOpen：[in] 整数，参见SDT_ResetSAM。
'返回值：0 x90 读基本信息成功 其他 读基本信息失败(具体含义参见返回码表)
Public Declare Function SDT_ReadBaseMsg Lib "sdtapi.dll" (ByVal iPort As Integer, ByRef pucCHMsg As String, ByRef puiCHMsgLen As Integer, ByRef pucPHMsg As String, ByRef puiPHMsgLen As Integer, ByVal iIfOpen As Integer) As Integer

'读取ID卡内IIN , SN和DN
'参数说明：iPort：[in] 整数，表示端口号。参见SDT_ResetSAM。pucIINSNDN：[out] 无符号字符指针，指向读到的IIN,SN和DN,长度为固定28字节。iIfOpen：[in] 整数，参见SDT_ResetSAM。
'返回值：0 x90 读IIN, SN和DN成功其他 读IIN, SN和DN失败(具体含义参见返回码表)
Public Declare Function SDT_ReadIINSNDN Lib "sdtapi.dll" (ByVal iPort As Integer, ByRef pucIINSNDN As String, ByVal iIfOpen As Integer) As Integer

