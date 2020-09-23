Attribute VB_Name = "modDeclaration"
'Public Type
Public Type UDP_CLIENT_HEADER
    Version As Long
    uin As Long
    Password As String
    SessionID As Long
    Command As Integer
    SeqNum1 As Integer
    SeqNum2 As Integer
    Parameter As String
End Type

Public Type UDP_SERVER_HEADER
    Version As Integer
    SessionID As Long
    Command As Integer
    SeqNum1 As Integer
    SeqNum2 As Integer
    uin As Long
    Parameter As String
End Type

Public Type WinsockConfig
    InternalIP As String
    ExternalIP As String
    TCPListenPort As Integer
    UDPRemoteHost As String
    UDPRemotePort As Integer
    UDPLocalPort As Integer
    ConnectionMethod As LOGIN_PROXY_STATUS
    ConnectionState As Connection_Status
    InitialLoginStatus As LOGIN_ONLINE_STATUS
End Type

Public Type USER_DETAIL_INFO
  lngUIN As Long
  lngStatus As LOGIN_ONLINE_STATUS
  strNickname As String
  strFirstname As String
  strLastName As String
  strEmail As String
  strEmail2 As String
  strEmail3 As String
  strPhone As String
  strFax As String
  strCellular As String
  strStreet As String
  strCity As String
  strState As String
  lngZip As Long
  intCountryCode As Integer
  intCountryStat As Integer
  byteAgeYear As Byte
  strHomepageURL As String
  byteBirthYear As Byte
  byteBirthMonth As Byte
  byteBirthDay As Byte
  strAboutInfo As String
  bAuthRequest As USER_AUTHORIZATION_REQUIRE
  bWebPresence As Boolean
  bPublishIP As Boolean
  bEmailHide As Boolean
End Type

Public Type CONTACT_DETAIL
  lngUIN As Long
  lngStatus As LOGIN_ONLINE_STATUS
  UserDetail As USER_DETAIL_INFO
  TCP_ExternalIP As String
  TCP_ExternalPort As Integer
  TCP_InternalIP As String
  TCP_FLAG As TCP_FLAG
  TCP_Version As Long
End Type

Public Type MESSAGE_HEADER
  lngUIN As Long
  MSG_Date As String
  MSG_Time As String
  MSG_Type As MESSAGE_TYPE
  MSG_Text As String
  URL_Address As String
  URL_Description As String
  AUTH_NickName As String
  AUTH_FirstName As String
  AUTH_LastName As String
  AUTH_Email As String
  AUTH_Reason As String
  
End Type

'Enumeration
Public Enum LOGIN_PROXY_STATUS
    LOGIN_NO_TCP = 1
    LOGIN_SNDONLY_TCP = 2
    LOGIN_SNDRCV_TCP = 4
End Enum

Public Enum TCP_FLAG
  TCP_CAPABLE = &H4
  TCP_INCAPABLE = &H6
End Enum

Public Enum USER_AUTHORIZATION_REQUIRE
    AUTHORIZATION_NOTNEEDED = 0
    AUTHORIZATION_REQUIRED = 1
End Enum

Public Enum LOGIN_ONLINE_STATUS
    STATUS_ONLINE = &H0
    STATUS_INVISIBLE = &H100
    STATUS_NA = &H5
    STATUS_OCCUPIED = &H11
    STATUS_AWAY = &H1
    STATUS_DND = &H13
    STATUS_CHAT = &H20
    STATUSF_WEBAWARE = &H10000
    STATUSF_SHOWIP = &H20000
    STATUS_OFFLINE = &HFFFF
End Enum

Public Enum RANDOM_CHAT_MODE
    RAND_General = 1
    RAND_Romance = 2
    RAND_Games = 3
    RAND_Students = 4
    RAND_20something = 6
    RAND_30something = 7
    RAND_40something = 8
    RAND_50over = 9
    RAND_ManRequestWoman = 10
    RAND_WomanRequestMan = 11
End Enum

Public Enum UDP_CLIENT_COMMAND
    UDP_CMD_ACK = &HA
    UDP_CMD_SEND_THRU_SRV = &H10E
    UDP_CMD_LOGIN = &H3E8
    UDP_CMD_CONT_LIST = &H406
    UDP_CMD_SEARCH_UIN = &H41A
    UDP_CMD_SEARCH_USER = &H424
    UDP_CMD_KEEP_ALIVE = &H42E
    UDP_CMD_KEEP_ALIVE2 = &H51E
    UDP_CMD_SEND_TEXT_CODE = &H438
    UDP_CMD_LOGIN_1 = &H44C
    UDP_CMD_INFO_REQ = &H460
    UDP_CMD_EXT_INFO_REQ = &H46A
    UDP_CMD_CHANGE_PW = &H49C
    UDP_CMD_STATUS_CHANGE = &H4D8
    UDP_CMD_LOGIN_2 = &H528
    UDP_CMD_UPDATE_INFO = &H50A
    UDP_CMD_UPDATE_AUTH = &H514
    UDP_CMD_UPDATE_EXT_INFO = &H4B0
    UDP_CMD_ADD_TO_LIST = &H53C
    UDP_CMD_REQ_ADD_LIST = &H456
    UDP_CMD_QUERY_SERVERS = &H4BA
    UDP_CMD_QUERY_ADDONS = &H4C4
    UDP_CMD_NEW_USER_1 = &H4EC
    UDP_CMD_NEW_USER_INFO = &H4A6
    UDP_CMD_ACK_MESSAGES = &H442
    UDP_CMD_MSG_TO_NEW_USER = &H456
    UDP_CMD_REG_NEW_USER = &H3FC
    UDP_CMD_VIS_LIST = &H6AE
    UDP_CMD_INVIS_LIST = &H6A4
    UDP_CMD_META_USER = &H64A
    UDP_CMD_RAND_SEARCH = &H56E
    UDP_CMD_RAND_SET = &H564
    UDP_CMD_REVERSE_TCP_CONN = &H15E
End Enum

Public Enum UDP_SERVER_REPLY
    UDP_SRV_ACK = &HA
    UDP_SRV_LOGIN_REPLY = &H5A
    UDP_SRV_USER_ONLINE = &H6E
    UDP_SRV_USER_OFFLINE = &H78
    UDP_SRV_USER_FOUND = &H8C
    UDP_SRV_OFFLINE_MESSAGE = &HDC
    UDP_SRV_END_OF_SEARCH = &HA0
    UDP_SRV_INFO_REPLY = &H118
    UDP_SRV_EXT_INFO_REPLY = &H122
    UDP_SRV_STATUS_UPDATE = &H1A4
    UDP_SRV_X1 = &H21C
    UDP_SRV_X2 = &HE6
    UDP_SRV_UPDATE = &H1E0
    UDP_SRV_UPDATE_EXT = &HC8
    UDP_SRV_NEW_UIN = &H46
    UDP_SRV_NEW_USER = &HB4
    UDP_SRV_QUERY = &H82
    UDP_SRV_SYSTEM_MESSAGE = &H1C2
    UDP_SRV_ONLINE_MESSAGE = &H104
    UDP_SRV_GO_AWAY = &HF0
    UDP_SRV_TRY_AGAIN = &HFA
    UDP_SRV_FORCE_DISCONNECT = &H28
    UDP_SRV_MULTI_PACKET = &H212
    UDP_SRV_WRONG_PASSWORD = &H64
    UDP_SRV_INVALID_UIN = &H12C
    UDP_SRV_META_USER = &H3DE
    UDP_SRV_RAND_USER = &H24E
    UDP_SRV_AUTH_UPDATE = &H1F4
End Enum

Public Enum META_COMMAND
    META_CMD_SET_INFO = 1000
    META_CMD_SET_MORE = 1020
    META_CMD_SET_ABOUT = 1030
    META_CMD_SET_SECURE = 1060
    META_CMD_SET_PASS = 1070
    META_CMD_REQ_INFO = 1200
    META_SRV_RES_INFO = 100
    META_SRV_RES_HOMEPAGE = 120
    META_SRV_RES_ABOUT = 130
    META_SRV_RES_SECURE = 160
    META_SRV_RES_PASS = 170
    META_SRV_USER_INFO = 200
    META_SRV_USER_WORK = 210
    META_SRV_USER_MORE = 220
    META_SRV_USER_ABOUT = 230
    META_SRV_USER_INTERESTS = 240
    META_SRV_USER_AFFILIATIONS = 250
    META_SRV_USER_HPCATEGORY = 270
    META_SRV_USER_FOUND = 410
End Enum

Public Enum MESSAGE_TYPE
    TYPE_MSG = &H1
    TYPE_CHAT = &H2
    TYPE_FILE = &H3
    TYPE_URL = &H4
    TYPE_AUTH_REQ = &H6
    TYPE_AUTH_DECLINE = &H7
    TYPE_AUTH_ACCEPT = &H8
    TYPE_ADDED = &HC
    TYPE_WEBPAGER = &HD
    TYPE_EXPRESS = &HE
    TYPE_CONTACT = &H13
    TYPE_MASS_MASK = &H8000
End Enum

Public Enum MESSAGE_SEND_METHOD
    ICQ_SEND_THRUSERVER = 0
    ICQ_SEND_DIRECT = 1
    ICQ_SEND_BESTWAY = 2
End Enum

Public Enum Connection_Status
    ICQ_Disconnected = 0
    ICQ_Register_New_User = 1
    ICQ_Login = 2
    ICQ_Connected = 3
End Enum
            
Public Const META_SRV_SUCCESS = 10
Public Const META_SRV_FAILURE = 50

'Constant Info
Public Const DEBUG_INFO = True
Public Const ICQ_UDP_VERSION = "0500"
Public Const ICQ_TCP_VERSION = "0600"
Public Const qIntervalTime = 10
Public Const KeepAliveIntervalTime = 110
'Public Variables
Public ICQEngine As WinsockConfig
Public Owner As UDP_CLIENT_HEADER
Public icqTimer As Integer       'Our very own built in timer
Public KeepAliveTimer As Integer 'Another timer, this one to take care of
                                 'KeepAlive command, which must be sent
                                 'every 2 minutes.

