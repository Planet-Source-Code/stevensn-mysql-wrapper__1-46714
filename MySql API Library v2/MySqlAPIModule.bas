Attribute VB_Name = "MySqlAPIModule"
' Windows API Declarations
Public Declare Function lstrlen Lib "kernel32" (ByVal lpString As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (Destination As Any, Source As Any, ByVal byteLen As Long)
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" _
    (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" _
    (ByVal lpBuffer As String, ByVal nSize As Long) As Long

'#################################################################################
'
' MySQL API Declarations
'
' For a description of each of the MySql API functions see the file MySqlAPI.txt in the same folder
' as this module
'
'#################################################################################
Public Declare Function MySqlNumRows Lib "libmysql.dll" Alias "mysql_num_rows" _
    (pResult As MYSQL_RES) As ULongLong
Public Declare Function MySqlNumFields Lib "libmysql.dll" Alias "mysql_num_fields" _
    (pResult As MYSQL_RES) As Long
Public Declare Function MySqlEOF Lib "libmysql.dll" Alias "mysql_eof" _
    (pResult As MYSQL_RES) As Byte
Public Declare Function MySqlFetchFieldDirect Lib "libmysql.dll" Alias "mysql_fetch_field_direct" _
    (pResult As MYSQL_RES, Index As Long) As Long
Public Declare Function MySqlFetchFields Lib "libmysql.dll" Alias "mysql_fetch_fields" _
    (pResult As MYSQL_RES) As Long
Public Declare Function MySqlRowTell Lib "libmysql.dll" Alias "mysql_row_tell" _
    (pResult As MYSQL_RES) As Long
Public Declare Function MySqlFieldTell Lib "libmysql.dll" Alias "mysql_field_tell" _
    (pResult As MYSQL_RES) As Long
Public Declare Function MySqlFieldCount Lib "libmysql.dll" Alias "mysql_field_count" _
    (pMySql As MYSQL) As Long
Public Declare Function MySqlAffectedRows Lib "libmysql.dll" Alias "mysql_affected_rows" _
    (pMySql As MYSQL) As ULongLong
Public Declare Function MySqlInsertId Lib "libmysql.dll" Alias "mysql_insert_id" _
    (pMySql As MYSQL) As ULongLong
Public Declare Function MySqlErrNo Lib "libmysql.dll" Alias "mysql_errno" _
    (pMySql As MYSQL) As Long
Public Declare Function MySqlError Lib "libmysql.dll" Alias "mysql_error" _
    (pMySql As MYSQL) As Long
Public Declare Function MySqlInfo Lib "libmysql.dll" Alias "mysql_info" _
    (ByRef pMySql As MYSQL) As Long
Public Declare Function MySqlThreadId Lib "libmysql.dll" Alias "mysql_thread_id" _
    (ByRef pMySql As MYSQL) As Long
Public Declare Function MySqlInit Lib "libmysql.dll" Alias "mysql_init" _
    (ByRef pMySql As MYSQL) As Long
Public Declare Function MySqlConnect Lib "libmysql.dll" Alias "mysql_connect" _
    (ByRef pMySql As MYSQL, ByVal lpHost As Long, ByVal lpUser As Long, _
    ByVal lpPassword As Long) As Long
Public Declare Function MySqlRealConnect Lib "limysql.dll" Alias "mysql_real_connect" _
    (ByRef pMySql As MYSQL, ByVal lpHost As Long, ByVal lpUser As Long, _
    ByVal lpPassword As Long, ByVal lpDb As Long, ByVal dwPort As Long, _
    ByVal lpUnixSocket As Long, dwClientFlag As Long) As Long
Public Declare Function MySqlClose Lib "libmysql.dll" Alias "mysql_close" _
    (pMySql As MYSQL) As Long
Public Declare Function MySqlSelectDb Lib "libmysql.dll" Alias "mysql_select_db" _
    (ByRef pMySql As MYSQL, ByVal lpDb As Long) As Long
Public Declare Function MySqlQuery Lib "libmysql.dll" Alias "mysql_query" _
    (ByRef pMySql As MYSQL, ByVal lpQuery As Long) As Long
Public Declare Function MySqlRealQuery Lib "libmysql.dll" Alias "mysql_real_query" _
    (ByRef pMySql As MYSQL, ByVal lpQuery As Long, ByVal dwLength As Long) As Long
Public Declare Function MySqlCreateDb Lib "libmysql.dll" Alias "mysql_create_db" _
    (ByRef pMySql As MYSQL, ByVal lpDb As Long) As Long
Public Declare Function MySqlDropDb Lib "libmysql.dll" Alias "mysql_drop_db" _
    (ByRef pMySql As MYSQL, ByVal lpDb As Long) As Long
Public Declare Function MySqlShutdown Lib "libmysql.dll" Alias "mysql_shutdown" _
    (ByRef pMySql As MYSQL) As Long
Public Declare Function MySqlDumpDebugInfo Lib "libmysql.dll" Alias "mysql_dump_debug_info" _
    (ByRef pMySql As MYSQL) As Long
Public Declare Function MySqlRefresh Lib "libmysql.dll" Alias "mysql_refresh" _
    (ByRef pMySql As MYSQL, ByVal dwRefreshOptions As Long) As Long
Public Declare Function MySqlKill Lib "libmysql.dll" Alias "mysql_kill" _
    (ByRef pMySql As MYSQL, ByVal pID As Long) As Long
Public Declare Function MySqlPing Lib "libmysql.dll" Alias "mysql_ping" _
    (ByRef pMySql As MYSQL) As Long
Public Declare Function MySqlStat Lib "libmysql.dll" Alias "mysql_stat" _
    (ByRef pMySql As MYSQL) As Long
Public Declare Function MySqlGetServerInfo Lib "libmysql.dll" Alias "mysql_get_server_info" _
    (ByRef pMySql As MYSQL) As Long
Public Declare Function MySqlGetClientInfo Lib "libmysql.dll" Alias "mysql_get_client_info" _
    () As Long
Public Declare Function MySqlGetHostInfo Lib "libmysql.dll" Alias "mysql_get_host_info" _
    (ByRef pMySql As MYSQL) As Long
Public Declare Function MySqlGetProtoInfo Lib "libmysql.dll" Alias "mysql_get_proto_info" _
    (ByRef pMySql As MYSQL) As Long
Public Declare Function MySqlListDbs Lib "libmysql.dll" Alias "mysql_list_dbs" _
    (ByRef pMySql As MYSQL, ByVal lpWildcard As Long) As Long
Public Declare Function MySqlListTables Lib "libmysql.dll" Alias "mysql_list_tables" _
    (ByRef pMySql As MYSQL, ByVal lpWildcard As Long) As Long
Public Declare Function MySqlListFields Lib "libmysql.dll" Alias "mysql_list_fields" _
    (ByRef pMySql As MYSQL, ByVal lpTable As Long, ByVal lpWildcard As Long) As Long
Public Declare Function MySqlListProcesses Lib "libmysql.dll" Alias "mysql_list_processes" _
    (ByRef pMySql As MYSQL) As Long
Public Declare Function MySqlStoreResult Lib "libmysql.dll" Alias "mysql_store_result" _
    (ByRef pMySql As MYSQL) As Long
Public Declare Function MySqlUseResult Lib "libmysql.dll" Alias "mysql_use_result" _
    (ByRef pMySql As MYSQL) As Long
Public Declare Function MySqlOptions Lib "libmysql.dll" Alias "mysql_options" _
    (ByRef pMySql As MYSQL, pMySqlOption As MYSQL_OPTION, ByVal lpArg As Long) As Long
Public Declare Function MySqlFreeResult Lib "libmysql.dll" Alias "mysql_free_result" _
    (ByRef pResult As MYSQL_RES) As Long
Public Declare Function MySqlDataSeek Lib "libmysql.dll" Alias "mysql_data_seek" _
    (ByRef pResult As MYSQL_RES, ByVal dOffset As Double) As Long
Public Declare Function MySqlRowSeek Lib "libmysql.dll" Alias "mysql_row_seek" _
    (ByRef pResult As MYSQL_RES, ByVal dwRowOffset As Long) As Long
Public Declare Function MySqlFieldSeek Lib "libmysql.dll" Alias "mysql_field_seek" _
    (ByRef pResult As MYSQL_RES, dOffset) As Long
Public Declare Function MySqlFetchRow Lib "libmysql.dll" Alias "mysql_fetch_row" _
    (ByRef pResult As MYSQL_RES) As Long
Public Declare Function MySqlFetchLengths Lib "libmysql.dll" Alias "mysql_fetch_lengths" _
    (pResult As MYSQL_RES) As Long
Public Declare Function MySqlFetchField Lib "libmysql.dll" Alias "mysql_fetch_field" _
    (pResult As MYSQL_RES) As Long
Public Declare Function MySqlEscapeString Lib "libmysql.dll" Alias "mysql_escape_string" _
    (ByRef pMySql As MYSQL, ByVal lpTo As Long, ByVal lpFrom As Long, ByVal dwLength As Long) As Long
Public Declare Function MySqlDebug Lib "libmysql.dll" Alias "mysql_debug" _
    (ByVal lpDebug As Long) As Long
Public Declare Function MySqlThreadSafe Lib "libmysql.dll" Alias "mysql_thread_safe" _
    () As Long
Public Declare Function MySqlCharacterSetName Lib "libmysql.dll" Alias "mysql_character_set_name" _
    (ByRef pMySql As MYSQL) As Long
Public Declare Function MySqlChangeUser Lib "libmysql.dll" Alias "mysql_change_user" _
    (ByRef pMySql As MYSQL, ByVal lpUser As Long) As Byte

'#################################################################################
'
' Translation from mysql_com.h
'
'#################################################################################
Public Const SIZE_OF_CHAR = 4
Public Const NAME_LEN = 64
Public Const HOSTNAME_LENGTH = 60
Public Const USERNAME_LENGTH = 16
Public Const LOCAL_HOST = "localhost"
Public Const LOCAL_HOST_NAMEDPIPE = "MySQL"
Public Const MYSQL_SERVICE_NAME = "MySql"

Public Enum enumServerCommand
    comSleep
    comQuit
    comInitDb
    comQuery
    comFieldList
    comCreateDb
    comDropDb
    comRefresh
    comShutdown
    comStatistics
    comProcessInfo
    comConnect
    comProcessKill
    comDebug
    comPing
    comTime
    comDelayedInsert
End Enum

Public Const NOT_NULL_FLAG = 1
Public Const PRI_KEY_FLAG = 2
Public Const UNIQUE_KEY_FLAG = 4
Public Const MULTIPLE_KEY_FLAG = 8
Public Const BLOB_FLAG = 16
Public Const UNSIGNED_FLAG = 32
Public Const ZEROFILL_FLAG = 64
Public Const BINARY_FLAG = 128
Public Const ENUM_FLAG = 256
Public Const AUTO_INCREMENT_FLAG = 512
Public Const TIMESTAMP_FLAG = 1024
Public Const SET_FLAG = 2048
Public Const PART_KEY_FLAG = 16384
Public Const GROUP_FLAG = 32768
Public Const UNIQUE_FLAG = 65536

Public Const REFRESH_GRANT = 1
Public Const REFRESH_LOG = 2
Public Const REFRESH_TABLES = 4
Public Const REFRESH_HOSTS = 8
Public Const REFRESH_STATUS = 16
Public Const REFRESH_THREADS = 32
Public Const REFRESH_SLAVE = 64
Public Const REFRESH_MASTER = 128
Public Const REFRESH_READ_LOCK = 256
Public Const REFRESH_FAST = 32768

Public Const CLIENT_LONG_PASSWORD = 1
Public Const CLIENT_FOUND_ROWS = 2
Public Const CLIENT_LONG_FLAG = 4
Public Const CLIENT_CONNECT_WITH_DB = 8
Public Const CLIENT_NO_SCHEMA = 16
Public Const CLIENT_COMPRESS = 32
Public Const CLIENT_ODBC = 64
Public Const CLIENT_LOCAL_FILES = 128
Public Const CLIENT_IGNORE_SPACE = 256
Public Const CLIENT_CHANGE_USER = 512
Public Const CLIENT_INTERACTIVE = 1024
Public Const CLIENT_SSL = 2048
Public Const CLIENT_IGNORE_SIGPIPE = 4096
Public Const CLIENT_TRANSACTIONS = 8196

Public Const SERVER_STATUS_IN_TRANS = 1
Public Const SERVER_STATUS_AUTOCOMMIT = 2

Public Const MYSQL_ERRMSG_SIZE = 200
Public Const NET_READ_TIMEOUT = 30
Public Const NET_WRITE_TIMEOUT = 60
Public Const NET_WAIT_TIMEOUT = 8 * 60 * 60

Public Const PACKET_ERROR = -1

Public Type NET
    vio As Long
    fd As Long
    fcntl As Long
    buff As Long
    buff_end As Long
    write_pos As Long
    read_pos As Long
    last_error(1 To MYSQL_ERRMSG_SIZE) As Byte
    last_errno As Long
    max_packet As Long
    timeout As Long
    pkt_nr As Long
    error As Byte
    return_errno As Byte
    compress As Byte
    no_send_ok As Byte
    remain_in_buf As Long
    length As Long
    buf_length As Long
    where_b As Long
    return_status As Long
    reading_or_writing As Byte
    save_char As Byte
End Type

'Public Const FIELD_TYPE_CHAR = FIELD_TYPE_TINY
'Public Const FIELD_TYPE_INTERVAL = FIELD_TYPE_ENUM

'#################################################################################
'
' Tranalation from mysql_version.h
'
'#################################################################################
Public Const PROTOCOL_VERSION = 10
Public Const MYSQL_SERVER_VERSION = "3.23.33"
Public Const MYSQL_SERVER_SUFFIX = ""
Public Const FRM_VER = 6
Public Const MYSQL_VERSION_ID = 32333
Public Const MYSQL_PORT = 3306
Public Const MYSQL_UNIX_ADDR = "/tmp/mysql.sock"

'#################################################################################
'
' Translation from mysql.h
'
'#################################################################################
Public Type ULongLong
    bytes(1 To 8) As Byte
End Type

Public Type USED_MEM
    next As Long
    left As Long
    size As Long
End Type

Public Type MEM_ROOT
    free As Long
    used As Long
    min_malloc As Long
    block_size As Long
    error_handler As Long
End Type

Public Type MYSQL_FIELD
    name As Long
    table As Long
    def As Long
    type As enumFieldTypes
    length As Long
    max_length As Long
    flags As Long
    decimals As Long
End Type

Public Type MYSQL_OPTION
    connect_timeout As Long
    client_flag As Long
    compress As Byte
    named_pipe As Byte
    port As Long
    host As Long
    init_command As Long
    user As Long
    Password As Long
    unix_socket As Long
    db As Long
    my_cnf_file As Long
    my_cnf_group As Long
    charset_dir As Long
    charset_name As Long
    use_ssl As Byte
    ssl_key As Long
    ssl_cert As Long
    ssl_ca As Long
    ssl_capath As Long
End Type

Public Enum enumMySqlOption
    MYSQL_OPT_CONNECT_TIMEOUT
    MYSQL_OPT_COMPRESS
    MYSQL_OPT_NAMED_PIPE
    MYSQL_INIT_COMMAND
    MYSQL_READ_DEFAULT_FILE
    MYSQL_READ_DEFAULT_GROUP
    MYSQL_SET_CHARSET_DIR
    MYSQL_SET_CHARSET_NAME
End Enum

Public Enum enumMySqlStatus
    MYSQL_STATUS_READY
    MYSQL_STATUS_GET_RESULT
    MYSQL_STATUS_USE_RESULT
End Enum

Public Type MYSQL
    net_a As NET
    connector_fd As Long
    host As Long
    user As Long
    Password As Long
    unix_socket As Long
    server_version As Long
    host_info As Long
    info As Long
    db As Long
    port As Long
    client_flag As Long
    server_capabilities As Long
    protocol_ver As Long
    field_count As Long
    server_status As Long
    thread_id As Long
    affected_rows As ULongLong
    insert_id As ULongLong
    extra_info As ULongLong
    packet_length As Long
    Status As enumMySqlStatus
    Fields As Long
    field_alloc As MEM_ROOT
    fix_misalignment As Long
    free_mem As Byte
    reconnect As Byte
    options As MYSQL_OPTION
    scramble_buff(1 To 9) As Byte
    charset As Long
    server_language As Long
End Type

Public Type MYSQL_DATA
    rows As ULongLong
    Fields As Long
    data As Long
    alloc As MEM_ROOT
    fix_misalignment As Long
End Type

Public Type MYSQL_ROWS
    next As Long
    data As Long
End Type

Public Type MYSQL_RES
    row_count As ULongLong
    field_count As Long
    current_field As Long
    Fields As Long
    data As Long
    data_cursor As Long
    field_alloc As MEM_ROOT
    fix_misalignment As Long
    row As Long
    current_row As Long
    lengths As Long
    handle As Long
    EOF As Byte
End Type

'#################################################################################
'
' MySQL Errors and Descriptions
'
'#################################################################################
Public Const CR_UNKNOWN_ERROR = 2000
Public Const CR_SOCKET_CREATE_ERROR = 2001
Public Const CR_CONNECTION_ERROR = 2002
Public Const CR_CONN_HOST_ERROR = 2003
Public Const CR_IPSOCK_ERROR = 2004
Public Const CR_UNKNOWN_HOST = 2005
Public Const CR_SERVER_GONE_ERROR = 2006
Public Const CR_VERSION_ERROR = 2007
Public Const CR_OUT_OF_MEMORY = 2008
Public Const CR_WRONG_HOST_INFO = 2009
Public Const CR_LOCALHOST_CONNECTION = 2010
Public Const CR_TCP_CONNECTION = 2011
Public Const CR_SERVER_HANDSHAKE_ERR = 2012
Public Const CR_SERVER_LOST = 2013
Public Const CR_COMMANDS_OUT_OF_SYNC = 2014
Public Const CR_NAMEDPIPE_CONNECTION = 2015
Public Const CR_NAMEDPIPEWAIT_ERROR = 2016
Public Const CR_NAMEDPIPEOPEN_ERROR = 2017
Public Const CR_NAMEDPIPESETSTATE_ERROR = 2018

Public Const ER_HASHCHK = 1000
Public Const ER_NISAMCHK = 1001
Public Const ER_NO = 1002
Public Const ER_YES = 1003
Public Const ER_CANT_CREATE_FILE = 1004
Public Const ER_CANT_CREATE_TABLE = 1005
Public Const ER_CANT_CREATE_DB = 1006
Public Const ER_DB_CREATE_EXISTS = 1007
Public Const ER_DB_DROP_EXISTS = 1008
Public Const ER_DB_DROP_DELETE = 1009
Public Const ER_DB_DROP_RMDIR = 1010
Public Const ER_CANT_DELETE_FILE = 1011
Public Const ER_CANT_FIND_SYSTEM_REC = 1012
Public Const ER_CANT_GET_STAT = 1013
Public Const ER_CANT_GET_WD = 1014
Public Const ER_CANT_LOCK = 1015
Public Const ER_CANT_OPEN_FILE = 1016
Public Const ER_FILE_NOT_FOUND = 1017
Public Const ER_CANT_READ_DIR = 1018
Public Const ER_CANT_SET_WD = 1019
Public Const ER_CHECKREAD = 1020
Public Const ER_DISK_FULL = 1021
Public Const ER_DUP_KEY = 1022
Public Const ER_ERROR_ON_CLOSE = 1023
Public Const ER_ERROR_ON_READ = 1024
Public Const ER_ERROR_ON_RENAME = 1025
Public Const ER_ERROR_ON_WRITE = 1026
Public Const ER_FILE_USED = 1027
Public Const ER_FILSORT_ABORT = 1028
Public Const ER_FORM_NOT_FOUND = 1029
Public Const ER_GET_ERRNO = 1030
Public Const ER_ILLEGAL_HA = 1031
Public Const ER_KEY_NOT_FOUND = 1032
Public Const ER_NOT_FORM_FILE = 1033
Public Const ER_NOT_KEYFILE = 1034
Public Const ER_OLD_KEYFILE = 1035
Public Const ER_OPEN_AS_READONLY = 1036
Public Const ER_OUTOFMEMORY = 1037
Public Const ER_OUT_OF_SORTMEMORY = 1038
Public Const ER_UNEXPECTED_EOF = 1039
Public Const ER_CON_COUNT_ERROR = 1040
Public Const ER_OUT_OF_RESOURCES = 1041
Public Const ER_BAD_HOST_ERROR = 1042
Public Const ER_HANDSHAKE_ERROR = 1043
Public Const ER_DBACCESS_DENIED_ERROR = 1044
Public Const ER_ACCESS_DENIED_ERROR = 1045
Public Const ER_NO_DB_ERROR = 1046
Public Const ER_UNKNOWN_COM_ERROR = 1047
Public Const ER_BAD_NULL_ERROR = 1048
Public Const ER_BAD_DB_ERROR = 1049
Public Const ER_TABLE_EXISTS_ERROR = 1050
Public Const ER_BAD_TABLE_ERROR = 1051
Public Const ER_NON_UNIQ_ERROR = 1052
Public Const ER_SERVER_SHUTDOWN = 1053
Public Const ER_BAD_FIELD_ERROR = 1054
Public Const ER_WRONG_FIELD_WITH_GROUP = 1055
Public Const ER_WRONG_GROUP_FIELD = 1056
Public Const ER_WRONG_SUM_SELECT = 1057
Public Const ER_WRONG_VALUE_COUNT = 1058
Public Const ER_TOO_LONG_IDENT = 1059
Public Const ER_DUP_FIELDNAME = 1060
Public Const ER_DUP_KEYNAME = 1061
Public Const ER_DUP_ENTRY = 1062
Public Const ER_WRONG_FIELD_SPEC = 1063
Public Const ER_PARSE_ERROR = 1064
Public Const ER_EMPTY_QUERY = 1065
Public Const ER_NONUNIQ_TABLE = 1066
Public Const ER_INVALID_DEFAULT = 1067
Public Const ER_MULTIPLE_PRI_KEY = 1068
Public Const ER_TOO_MANY_KEYS = 1069
Public Const ER_TOO_MANY_KEY_PARTS = 1070
Public Const ER_TOO_LONG_KEY = 1071
Public Const ER_KEY_COLUMN_DOES_NOT_EXIST = 1072
Public Const ER_BLOB_USED_AS_KEY = 1073
Public Const ER_TOO_BIG_FIELDLENGTH = 1074
Public Const ER_WRONG_AUTO_KEY = 1075
Public Const ER_READY = 1076
Public Const ER_NORMAL_SHUTDOWN = 1077
Public Const ER_GOT_SIGNAL = 1078
Public Const ER_SHUTDOWN_COMPLETE = 1079
Public Const ER_FORCING_CLOSE = 1080
Public Const ER_IPSOCK_ERROR = 1081
Public Const ER_NO_SUCH_INDEX = 1082
Public Const ER_WRONG_FIELD_TERMINATOR = 1083
Public Const ER_BLOBS_AND_NO_TERMINATED = 1084
Public Const ER_TEXTFILE_NOT_READABLE = 1085
Public Const ER_FILE_EXISTS_ERROR = 1086
Public Const ER_LOAD_INFO = 1087
Public Const ER_ALTER_INFO = 1088
Public Const ER_WRONG_SUB_KEY = 1089
Public Const ER_CANT_REMOVE_ALL_FIELDS = 1090
Public Const ER_CANT_DROP_FIELD_OR_KEY = 1091
Public Const ER_INSERT_INFO = 1092
Public Const ER_INSERT_TABLE_USED = 1093
Public Const ER_NO_SUCH_THREAD = 1094
Public Const ER_KILL_DENIED_ERROR = 1095
Public Const ER_NO_TABLES_USED = 1096
Public Const ER_TOO_BIG_SET = 1097
Public Const ER_NO_UNIQUE_LOGFILE = 1098
Public Const ER_TABLE_NOT_LOCKED_FOR_WRITE = 1099
Public Const ER_TABLE_NOT_LOCKED = 1100
Public Const ER_BLOB_CANT_HAVE_DEFAULT = 1101
Public Const ER_WRONG_DB_NAME = 1102
Public Const ER_WRONG_TABLE_NAME = 1103
Public Const ER_TOO_BIG_SELECT = 1104
Public Const ER_UNKNOWN_ERROR = 1105
Public Const ER_UNKNOWN_PROCEDURE = 1106
Public Const ER_WRONG_PARAMCOUNT_TO_PROCEDURE = 1107
Public Const ER_WRONG_PARAMETERS_TO_PROCEDURE = 1108
Public Const ER_UNKNOWN_TABLE = 1109
Public Const ER_FIELD_SPECIFIED_TWICE = 1110
Public Const ER_INVALID_GROUP_FUNC_USE = 1111
Public Const ER_UNSUPPORTED_EXTENSION = 1112
Public Const ER_TABLE_MUST_HAVE_COLUMNS = 1113
Public Const ER_RECORD_FILE_FULL = 1114
Public Const ER_UNKNOWN_CHARACTER_SET = 1115
Public Const ER_TOO_MANY_TABLES = 1116
Public Const ER_TOO_MANY_FIELDS = 1117
Public Const ER_TOO_BIG_ROWSIZE = 1118
Public Const ER_STACK_OVERRUN = 1119
Public Const ER_WRONG_OUTER_JOIN = 1120
Public Const ER_NULL_COLUMN_IN_INDEX = 1121
Public Const ER_CANT_FIND_UDF = 1122
Public Const ER_CANT_INITIALIZE_UDF = 1123
Public Const ER_UDF_NO_PATHS = 1124
Public Const ER_UDF_EXISTS = 1125
Public Const ER_CANT_OPEN_LIBRARY = 1126
Public Const ER_CANT_FIND_DL_ENTRY = 1127
Public Const ER_FUNCTION_NOT_DEFINED = 1128
Public Const ER_HOST_IS_BLOCKED = 1129
Public Const ER_HOST_NOT_PRIVILEGED = 1130
Public Const ER_PASSWORD_ANONYMOUS_USER = 1131
Public Const ER_PASSWORD_NOT_ALLOWED = 1132
Public Const ER_PASSWORD_NO_MATCH = 1133
Public Const ER_UPDATE_INFO = 1134
Public Const ER_CANT_CREATE_THREAD = 1135
Public Const ER_WRONG_VALUE_COUNT_ON_ROW = 1136
Public Const ER_CANT_REOPEN_TABLE = 1137
Public Const ER_INVALID_USE_OF_NULL = 1138
Public Const ER_REGEXP_ERROR = 1139
Public Const ER_MIX_OF_GROUP_FUNC_AND_FIELDS = 1140
Public Const ER_NONEXISTING_GRANT = 1141
Public Const ER_TABLEACCESS_DENIED_ERROR = 1142
Public Const ER_COLUMNACCESS_DENIED_ERROR = 1143
Public Const ER_ILLEGAL_GRANT_FOR_TABLE = 1144
Public Const ER_GRANT_WRONG_HOST_OR_USER = 1145
Public Const ER_NO_SUCH_TABLE = 1146
Public Const ER_NONEXISTING_TABLE_GRANT = 1147
Public Const ER_NOT_ALLOWED_COMMAND = 1148
Public Const ER_SYNTAX_ERROR = 1149
Public Const ER_DELAYED_CANT_CHANGE_LOCK = 1150
Public Const ER_TOO_MANY_DELAYED_THREADS = 1151
Public Const ER_ABORTING_CONNECTION = 1152
Public Const ER_NET_PACKET_TOO_LARGE = 1153
Public Const ER_NET_READ_ERROR_FROM_PIPE = 1154
Public Const ER_NET_FCNTL_ERROR = 1155
Public Const ER_NET_PACKETS_OUT_OF_ORDER = 1156
Public Const ER_NET_UNCOMPRESS_ERROR = 1157
Public Const ER_NET_READ_ERROR = 1158
Public Const ER_NET_READ_INTERRUPTED = 1159
Public Const ER_NET_ERROR_ON_WRITE = 1160
Public Const ER_NET_WRITE_INTERRUPTED = 1161
Public Const ER_TOO_LONG_STRING = 1162
Public Const ER_TABLE_CANT_HANDLE_BLOB = 1163
Public Const ER_TABLE_CANT_HANDLE_AUTO_INCREMENT = 1164
Public Const ER_DELAYED_INSERT_TABLE_LOCKED = 1165
Public Const ER_WRONG_COLUMN_NAME = 1166
Public Const ER_WRONG_KEY_COLUMN = 1167
Public Const ER_WRONG_MRG_TABLE = 1168
Public Const ER_DUP_UNIQUE = 1169
Public Const ER_BLOB_KEY_WITHOUT_LENGTH = 1170
Public Const ER_PRIMARY_CANT_HAVE_NULL = 1171
Public Const ER_TOO_MANY_ROWS = 1172
Public Const ER_REQUIRES_PRIMARY_KEY = 1173
Public Const ER_NO_RAID_COMPILED = 1174
Public Const ER_ERROR_MESSAGE = 175

Public Function Convert64ToLong(a As ULongLong) As Long
    Dim lngRes As Long
    
    CopyMemory lngRes, a.bytes(1), 4
    Convert64ToLong = lngRes
End Function

Public Function ConvertLongTo64(lngValue As Long) As ULongLong
    Dim a As ULongLong
    
    CopyMemory a.bytes(1), lngValue, 4
    ConvertLongTo64 = a
End Function

Public Function PointerToString(ByVal lpStr As Long) As String
On Local Error Resume Next
    Dim bytTest As Byte
    Dim bytSOut() As Byte
    Dim lngCChars As Long
    If lpStr = 0 Then Exit Function
    
    lngCChars = lstrlen(lpStr)
    ReDim bytSOut(1 To lngCChars)
    bytSOut = String$(lngCChars, " ")
    CopyMemory bytSOut(1), ByVal (lpStr), lngCChars
    
    PointerToString = StripNull(StrConv(bytSOut, vbUnicode))
End Function

Public Function StripNull(strValue As String) As String
    Dim lngLen As Long
    lngLen = InStr(strValue, vbNullChar)
    
    If lngLen > 0 Then
        StripNull = Trim(left$(strValue, lngLen - 1))
    Else
        StripNull = strValue
    End If
End Function
