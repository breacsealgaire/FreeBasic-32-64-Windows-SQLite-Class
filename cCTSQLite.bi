' ########################################################################################
' File: cCTSQLite.bi
' Contents: FreeBasic Windows SQLite class support.
' Version: 1.00
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2017 Rick Kelly
' Credits: Paul Squires, www.planetsquires.com
' Released into the public domain for private and public use without restriction
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#Pragma Once

#Include Once "windows.bi"
#Include Once "cCTSQL/cCTSQLiteSpooler.bi"

Type sqlite3        as ulong
Type sqlite3_backup as sqlite3
Type sqlite3_stmt   as sqlite3 

Enum cCTSqliteConstants

SQLITE_OK                          = &h00000000
SQLITE_TRUE                        = &h00000001
SQLITE_OPEN_READONLY               = &h00000001  
SQLITE_OPEN_READWRITE              = &h00000002  
SQLITE_OPEN_CREATE                 = &h00000004  
SQLITE_OPEN_URI                    = &h00000040  
SQLITE_OPEN_MEMORY                 = &h00000080  
SQLITE_OPEN_NOMUTEX                = &h00008000  
SQLITE_OPEN_FULLMUTEX              = &h00010000  
SQLITE_OPEN_SHAREDCACHE            = &h00020000  
SQLITE_OPEN_PRIVATECACHE           = &h00040000  
SQLITE_OPEN_WAL                    = &h00080000
SQLITE_ROW                         = 100
SQLITE_DONE                        = 101
SQLITE_INTEGER                     = 1
SQLITE_FLOAT                       = 2
SQLITE_TEXT                        = 3
SQLITE_BLOB                        = 4
SQLITE_NULL                        = 5
SQLITE_INTERNAL                    = 2
SQLITE_PERM                        = 3
SQLITE_ABORT                       = 4
SQLITE_BUSY                        = 5
SQLITE_LOCKED                      = 6
SQLITE_NOMEM                       = 7
SQLITE_READONLY                    = 8
SQLITE_INTERRUPT                   = 9
SQLITE_IOERR                       = 10
SQLITE_CORRUPT                     = 11
SQLITE_NOTFOUND                    = 12
SQLITE_FULL                        = 13
SQLITE_CANTOPEN                    = 14
SQLITE_PROTOCOL                    = 15
SQLITE_EMPTY                       = 16
SQLITE_SCHEMA                      = 17
SQLITE_TOOBIG                      = 18
SQLITE_CONSTRAINT                  = 19
SQLITE_MISMATCH                    = 20
SQLITE_MISUSE                      = 21
SQLITE_NOLFS                       = 22
SQLITE_AUTH                        = 23
SQLITE_FORMAT                      = 24
SQLITE_RANGE                       = 25
SQLITE_NOTADB                      = 26
SQLITE_NOTICE                      = 27
SQLITE_WARNING                     = 28
SQLITE_IOERR_READ                  = SQLITE_IOERR Or (1 Shl 8)
SQLITE_IOERR_SHORT_READ            = SQLITE_IOERR Or (2 Shl 8)
SQLITE_IOERR_WRITE                 = SQLITE_IOERR Or (3 Shl 8)
SQLITE_IOERR_FSYNC                 = SQLITE_IOERR Or (4 Shl 8)
SQLITE_IOERR_DIR_FSYNC             = SQLITE_IOERR Or (5 Shl 8)
SQLITE_IOERR_TRUNCATE              = SQLITE_IOERR Or (6 Shl 8)
SQLITE_IOERR_FSTAT                 = SQLITE_IOERR Or (7 Shl 8)
SQLITE_IOERR_UNLOCK                = SQLITE_IOERR Or (8 Shl 8)
SQLITE_IOERR_RDLOCK                = SQLITE_IOERR Or (9 Shl 8)
SQLITE_IOERR_DELETE                = SQLITE_IOERR Or (10 Shl 8)
SQLITE_IOERR_BLOCKED               = SQLITE_IOERR Or (11 Shl 8)
SQLITE_IOERR_NOMEM                 = SQLITE_IOERR Or (12 Shl 8)
SQLITE_IOERR_ACCESS                = SQLITE_IOERR Or (13 Shl 8)
SQLITE_IOERR_CHECKRESERVEDLOCK     = SQLITE_IOERR Or (14 Shl 8)
SQLITE_IOERR_LOCK                  = SQLITE_IOERR Or (15 Shl 8)
SQLITE_IOERR_CLOSE                 = SQLITE_IOERR Or (16 Shl 8)
SQLITE_IOERR_DIR_CLOSE             = SQLITE_IOERR Or (17 Shl 8)
SQLITE_IOERR_SHMOPEN               = SQLITE_IOERR Or (18 Shl 8)
SQLITE_IOERR_SHMSIZE               = SQLITE_IOERR Or (19 Shl 8)
SQLITE_IOERR_SHMLOCK               = SQLITE_IOERR Or (20 Shl 8)
SQLITE_IOERR_SHMMAP                = SQLITE_IOERR Or (21 Shl 8)
SQLITE_IOERR_SEEK                  = SQLITE_IOERR Or (22 Shl 8)
SQLITE_IOERR_DELETE_NOENT          = SQLITE_IOERR Or (23 Shl 8)
SQLITE_IOERR_MMAP                  = SQLITE_IOERR Or (24 Shl 8)
SQLITE_IOERR_GETTEMPPATH           = SQLITE_IOERR Or (25 Shl 8)
SQLITE_IOERR_CONVPATH              = SQLITE_IOERR Or (26 Shl 8)
SQLITE_LOCKED_SHAREDCACHE          = SQLITE_LOCKED Or (1 Shl 8)
SQLITE_BUSY_RECOVERY               = SQLITE_BUSY Or (1 Shl 8)
SQLITE_BUSY_SNAPSHOT               = SQLITE_BUSY Or (2 Shl 8)
SQLITE_CANTOPEN_NOTEMPDIR          = SQLITE_CANTOPEN Or (1 Shl 8)
SQLITE_CANTOPEN_ISDIR              = SQLITE_CANTOPEN Or (2 Shl 8)
SQLITE_CANTOPEN_FULLPATH           = SQLITE_CANTOPEN Or (3 Shl 8)
SQLITE_CANTOPEN_CONVPATH           = SQLITE_CANTOPEN Or (4 Shl 8)
SQLITE_CORRUPT_VTAB                = SQLITE_CORRUPT Or (1 Shl 8)
SQLITE_READONLY_RECOVERY           = SQLITE_READONLY Or (1 Shl 8)
SQLITE_READONLY_CANTLOCK           = SQLITE_READONLY Or (2 Shl 8)
SQLITE_READONLY_ROLLBACK           = SQLITE_READONLY Or (3 Shl 8)
SQLITE_READONLY_DBMOVED            = SQLITE_READONLY Or (4 Shl 8)
SQLITE_ABORT_ROLLBACK              = SQLITE_ABORT Or (2 Shl 8)
SQLITE_CONSTRAINT_CHECK            = SQLITE_CONSTRAINT Or (1 Shl 8)
SQLITE_CONSTRAINT_COMMITHOOK       = SQLITE_CONSTRAINT Or (2 Shl 8)
SQLITE_CONSTRAINT_FOREIGNKEY       = SQLITE_CONSTRAINT Or (3 Shl 8)
SQLITE_CONSTRAINT_FUNCTION         = SQLITE_CONSTRAINT Or (4 Shl 8)
SQLITE_CONSTRAINT_NOTNULL          = SQLITE_CONSTRAINT Or (5 Shl 8)
SQLITE_CONSTRAINT_PRIMARYKEY       = SQLITE_CONSTRAINT Or (6 Shl 8)
SQLITE_CONSTRAINT_TRIGGER          = SQLITE_CONSTRAINT Or (7 Shl 8)
SQLITE_CONSTRAINT_UNIQUE           = SQLITE_CONSTRAINT Or (8 Shl 8)
SQLITE_CONSTRAINT_VTAB             = SQLITE_CONSTRAINT Or (9 Shl 8)
SQLITE_CONSTRAINT_ROWID            = SQLITE_CONSTRAINT Or (10 Shl 8)
SQLITE_NOTICE_RECOVER_WAL          = SQLITE_NOTICE Or (1 Shl 8)
SQLITE_NOTICE_RECOVER_ROLLBACK     = SQLITE_NOTICE Or (2 Shl 8)
SQLITE_WARNING_AUTOINDEX           = SQLITE_WARNING Or (1 Shl 8)
SQLITE_AUTH_USER                   = SQLITE_AUTH Or (1 Shl 8)
SQLITE_OPEN_DELETEONCLOSE          = &h00000008
SQLITE_OPEN_EXCLUSIVE              = &h00000010
SQLITE_OPEN_AUTOPROXY              = &h00000020
SQLITE_OPEN_MAIN_DB                = &h00000100
SQLITE_OPEN_TEMP_DB                = &h00000200
SQLITE_OPEN_TRANSIENT_DB           = &h00000400
SQLITE_OPEN_MAIN_JOURNAL           = &h00000800
SQLITE_OPEN_TEMP_JOURNAL           = &h00001000
SQLITE_OPEN_SUBJOURNAL             = &h00002000
SQLITE_OPEN_MASTER_JOURNAL         = &h00004000
SQLITE_IOCAP_ATOMIC                = &h00000001
SQLITE_IOCAP_ATOMIC512             = &h00000002
SQLITE_IOCAP_ATOMIC1K              = &h00000004
SQLITE_IOCAP_ATOMIC2K              = &h00000008
SQLITE_IOCAP_ATOMIC4K              = &h00000010
SQLITE_IOCAP_ATOMIC8K              = &h00000020
SQLITE_IOCAP_ATOMIC16K             = &h00000040
SQLITE_IOCAP_ATOMIC32K             = &h00000080
SQLITE_IOCAP_ATOMIC64K             = &h00000100
SQLITE_IOCAP_SAFE_APPEND           = &h00000200
SQLITE_IOCAP_SEQUENTIAL            = &h00000400
SQLITE_IOCAP_UNDELETABLE_WHEN_OPEN = &h00000800
SQLITE_IOCAP_POWERSAFE_OVERWRITE   = &h00001000
SQLITE_IOCAP_IMMUTABLE             = &h00002000
SQLITE_LOCK_NONE                   = 0
SQLITE_LOCK_SHARED                 = 1
SQLITE_LOCK_RESERVED               = 2
SQLITE_LOCK_PENDING                = 3
SQLITE_LOCK_EXCLUSIVE              = 4
SQLITE_SYNC_NORMAL                 = &h00002
SQLITE_SYNC_FULL                   = &h00003
SQLITE_SYNC_DATAONLY               = &h00010
SQLITE_FCNTL_LOCKSTATE             = 1
SQLITE_FCNTL_GET_LOCKPROXYFILE     = 2
SQLITE_FCNTL_SET_LOCKPROXYFILE     = 3
SQLITE_FCNTL_LAST_ERRNO            = 4
SQLITE_FCNTL_SIZE_HINT             = 5
SQLITE_FCNTL_CHUNK_SIZE            = 6
SQLITE_FCNTL_FILE_POINTER          = 7
SQLITE_FCNTL_SYNC_OMITTED          = 8
SQLITE_FCNTL_WIN32_AV_RETRY        = 9
SQLITE_FCNTL_PERSIST_WAL           = 10
SQLITE_FCNTL_OVERWRITE             = 11
SQLITE_FCNTL_VFSNAME               = 12
SQLITE_FCNTL_POWERSAFE_OVERWRITE   = 13
SQLITE_FCNTL_PRAGMA                = 14
SQLITE_FCNTL_BUSYHANDLER           = 15
SQLITE_FCNTL_TEMPFILENAME          = 16
SQLITE_FCNTL_MMAP_SIZE             = 18
SQLITE_FCNTL_TRACE                 = 19
SQLITE_FCNTL_HAS_MOVED             = 20
SQLITE_FCNTL_SYNC                  = 21
SQLITE_FCNTL_COMMIT_PHASETWO       = 22
SQLITE_FCNTL_WIN32_SET_HANDLE      = 23
SQLITE_FCNTL_WAL_BLOCK             = 24
SQLITE_FCNTL_ZIPVFS                = 25
SQLITE_FCNTL_RBU                   = 26
SQLITE_GET_LOCKPROXYFILE           = SQLITE_FCNTL_GET_LOCKPROXYFILE
SQLITE_SET_LOCKPROXYFILE           = SQLITE_FCNTL_SET_LOCKPROXYFILE
SQLITE_LAST_ERRNO                  = SQLITE_FCNTL_LAST_ERRNO
SQLITE_ACCESS_EXISTS               = 0
SQLITE_ACCESS_READWRITE            = 1
SQLITE_ACCESS_READ                 = 2
SQLITE_SHM_UNLOCK                  = 1
SQLITE_SHM_LOCK                    = 2
SQLITE_SHM_SHARED                  = 4
SQLITE_SHM_EXCLUSIVE               = 8
SQLITE_SHM_NLOCK                   = 8
SQLITE_CONFIG_SINGLETHREAD         = 1
SQLITE_CONFIG_MULTITHREAD          = 2
SQLITE_CONFIG_SERIALIZED           = 3
SQLITE_CONFIG_MALLOC               = 4
SQLITE_CONFIG_GETMALLOC            = 5
SQLITE_CONFIG_SCRATCH              = 6
SQLITE_CONFIG_PAGECACHE            = 7
SQLITE_CONFIG_HEAP                 = 8
SQLITE_CONFIG_MEMSTATUS            = 9
SQLITE_CONFIG_MUTEX                = 10
SQLITE_CONFIG_GETMUTEX             = 11
SQLITE_CONFIG_LOOKASIDE            = 13
SQLITE_CONFIG_PCACHE               = 14
SQLITE_CONFIG_GETPCACHE            = 15
SQLITE_CONFIG_LOG                  = 16
SQLITE_CONFIG_URI                  = 17
SQLITE_CONFIG_PCACHE2              = 18
SQLITE_CONFIG_GETPCACHE2           = 19
SQLITE_CONFIG_COVERING_INDEX_SCAN  = 20
SQLITE_CONFIG_SQLLOG               = 21
SQLITE_CONFIG_MMAP_SIZE            = 22
SQLITE_CONFIG_WIN32_HEAPSIZE       = 23
SQLITE_CONFIG_PCACHE_HDRSZ         = 24
SQLITE_CONFIG_PMASZ                = 25
SQLITE_DBCONFIG_LOOKASIDE          = 1001
SQLITE_DBCONFIG_ENABLE_FKEY        = 1002
SQLITE_DBCONFIG_ENABLE_TRIGGER     = 1003
SQLITE_DENY                        = 1
SQLITE_IGNORE                      = 2
SQLITE_CREATE_INDEX                = 1
SQLITE_CREATE_TABLE                = 2
SQLITE_CREATE_TEMP_INDEX           = 3
SQLITE_CREATE_TEMP_TABLE           = 4
SQLITE_CREATE_TEMP_TRIGGER         = 5
SQLITE_CREATE_TEMP_VIEW            = 6
SQLITE_CREATE_TRIGGER              = 7
SQLITE_CREATE_VIEW                 = 8
SQLITE_DELETE                      = 9
SQLITE_DROP_INDEX                  = 10
SQLITE_DROP_TABLE                  = 11
SQLITE_DROP_TEMP_INDEX             = 12
SQLITE_DROP_TEMP_TABLE             = 13
SQLITE_DROP_TEMP_TRIGGER           = 14
SQLITE_DROP_TEMP_VIEW              = 15
SQLITE_DROP_TRIGGER                = 16
SQLITE_DROP_VIEW                   = 17
SQLITE_INSERT                      = 18
SQLITE_PRAGMA                      = 19
SQLITE_READ                        = 20
SQLITE_SELECT                      = 21
SQLITE_TRANSACTION                 = 22
SQLITE_UPDATE                      = 23
SQLITE_ATTACH                      = 24
SQLITE_DETACH                      = 25
SQLITE_ALTER_TABLE                 = 26
SQLITE_REINDEX                     = 27
SQLITE_ANALYZE                     = 28
SQLITE_CREATE_VTABLE               = 29
SQLITE_DROP_VTABLE                 = 30
SQLITE_FUNCTION                    = 31
SQLITE_SAVEPOINT                   = 32
SQLITE_COPY                        = 0
SQLITE_RECURSIVE                   = 33
SQLITE_LIMIT_LENGTH                = 0
SQLITE_LIMIT_SQL_LENGTH            = 1
SQLITE_LIMIT_COLUMN                = 2
SQLITE_LIMIT_EXPR_DEPTH            = 3
SQLITE_LIMIT_COMPOUND_SELECT       = 4
SQLITE_LIMIT_VDBE_OP               = 5
SQLITE_LIMIT_FUNCTION_ARG          = 6
SQLITE_LIMIT_ATTACHED              = 7
SQLITE_LIMIT_LIKE_PATTERN_LENGTH   = 8
SQLITE_LIMIT_VARIABLE_NUMBER       = 9
SQLITE_LIMIT_TRIGGER_DEPTH         = 10
SQLITE_LIMIT_WORKER_THREADS        = 11
SQLITE_UTF8                        = 1
SQLITE_UTF16LE                     = 2
SQLITE_UTF16BE                     = 3
SQLITE_UTF16                       = 4
SQLITE_ANY                         = 5
SQLITE_UTF16_ALIGNED               = 8
SQLITE_DETERMINISTIC               = &h800
SQLITE_MUTEX_FAST                  = 0
SQLITE_MUTEX_RECURSIVE             = 1
SQLITE_MUTEX_STATIC_MASTER         = 2
SQLITE_MUTEX_STATIC_MEM            = 3
SQLITE_MUTEX_STATIC_MEM2           = 4
SQLITE_MUTEX_STATIC_OPEN           = 4
SQLITE_MUTEX_STATIC_PRNG           = 5
SQLITE_MUTEX_STATIC_LRU            = 6
SQLITE_MUTEX_STATIC_LRU2           = 7
SQLITE_MUTEX_STATIC_PMEM           = 7
SQLITE_MUTEX_STATIC_APP1           = 8
SQLITE_MUTEX_STATIC_APP2           = 9
SQLITE_MUTEX_STATIC_APP3           = 10
SQLITE_MUTEX_STATIC_VFS1           = 11
SQLITE_MUTEX_STATIC_VFS2           = 12
SQLITE_MUTEX_STATIC_VFS3           = 13
SQLITE_TESTCTRL_FIRST              = 5
SQLITE_TESTCTRL_PRNG_SAVE          = 5
SQLITE_TESTCTRL_PRNG_RESTORE       = 6
SQLITE_TESTCTRL_PRNG_RESET         = 7
SQLITE_TESTCTRL_BITVEC_TEST        = 8
SQLITE_TESTCTRL_FAULT_INSTALL      = 9
SQLITE_TESTCTRL_BENIGN_MALLOC_HOOKS = 10
SQLITE_TESTCTRL_PENDING_BYTE       = 11
SQLITE_TESTCTRL_ASSERT             = 12
SQLITE_TESTCTRL_ALWAYS             = 13
SQLITE_TESTCTRL_RESERVE            = 14
SQLITE_TESTCTRL_OPTIMIZATIONS      = 15
SQLITE_TESTCTRL_ISKEYWORD          = 16
SQLITE_TESTCTRL_SCRATCHMALLOC      = 17
SQLITE_TESTCTRL_LOCALTIME_FAULT    = 18
SQLITE_TESTCTRL_EXPLAIN_STMT       = 19
SQLITE_TESTCTRL_NEVER_CORRUPT      = 20
SQLITE_TESTCTRL_VDBE_COVERAGE      = 21
SQLITE_TESTCTRL_BYTEORDER          = 22
SQLITE_TESTCTRL_ISINIT             = 23
SQLITE_TESTCTRL_SORTER_MMAP        = 24
SQLITE_TESTCTRL_IMPOSTER           = 25
SQLITE_TESTCTRL_LAST               = 25
SQLITE_STATUS_MEMORY_USED          = 0
SQLITE_STATUS_PAGECACHE_USED       = 1
SQLITE_STATUS_PAGECACHE_OVERFLOW   = 2
SQLITE_STATUS_SCRATCH_USED         = 3
SQLITE_STATUS_SCRATCH_OVERFLOW     = 4
SQLITE_STATUS_MALLOC_SIZE          = 5
SQLITE_STATUS_PARSER_STACK         = 6
SQLITE_STATUS_PAGECACHE_SIZE       = 7
SQLITE_STATUS_SCRATCH_SIZE         = 8
SQLITE_STATUS_MALLOC_COUNT         = 9
SQLITE_DBSTATUS_LOOKASIDE_USED     = 0
SQLITE_DBSTATUS_CACHE_USED         = 1
SQLITE_DBSTATUS_SCHEMA_USED        = 2
SQLITE_DBSTATUS_STMT_USED          = 3
SQLITE_DBSTATUS_LOOKASIDE_HIT      = 4
SQLITE_DBSTATUS_LOOKASIDE_MISS_SIZE = 5
SQLITE_DBSTATUS_LOOKASIDE_MISS_FULL = 6
SQLITE_DBSTATUS_CACHE_HIT          = 7
SQLITE_DBSTATUS_CACHE_MISS         = 8
SQLITE_DBSTATUS_CACHE_WRITE        = 9
SQLITE_DBSTATUS_DEFERRED_FKS       = 10
SQLITE_DBSTATUS_MAX                = 10
SQLITE_STMTSTATUS_FULLSCAN_STEP    = 1
SQLITE_STMTSTATUS_SORT             = 2
SQLITE_STMTSTATUS_AUTOINDEX        = 3
SQLITE_STMTSTATUS_VM_STEP          = 4
SQLITE_CHECKPOINT_PASSIVE          = 0
SQLITE_CHECKPOINT_FULL             = 1
SQLITE_CHECKPOINT_RESTART          = 2
SQLITE_CHECKPOINT_TRUNCATE         = 3
SQLITE_VTAB_CONSTRAINT_SUPPORT     = 1
SQLITE_ROLLBACK                    = 1
SQLITE_FAIL                        = 3
SQLITE_REPLACE                     = 5
SQLITE_SCANSTAT_NLOOP              = 0
SQLITE_SCANSTAT_NVISIT             = 1
SQLITE_SCANSTAT_EST                = 2
SQLITE_SCANSTAT_NAME               = 3
SQLITE_SCANSTAT_EXPLAIN            = 4
SQLITE_SCANSTAT_SELECTID           = 5
NOT_WITHIN                         = 0
PARTLY_WITHIN                      = 1
FULLY_WITHIN                       = 2

End Enum

Type SqlitePragma

   Name       as String
   Value      as String

End Type

' ########################################################################################
' cCTSQLite Class
' ########################################################################################

Type cCTSQLite Extends Object

Dim pSQLite          as Any Ptr
Dim lStartUp         as BOOLEAN = False
Dim iInitialize      as Long
Dim sStartUpError    as String = ""
Dim sSQLite3Version  as String = ""
Dim sRollback        as String = "ROLLBACK;"
Dim sCommit          as String = "COMMIT;"
Dim sPragmaVersion   as String = "PRAGMA USER_VERSION;"
   

' SQLite API's

' Recommendations from SQLite documentation

' For maximum portability, it is recommended that applications always invoke sqlite3_initialize()
' directly prior to using any other SQLite interface. Future releases of SQLite may require this.
' When all connections are closed sqlite3_shutdown() is needed to release all sqlite3_initialize()
' resources allocated.

' Shared cache is disabled by default. But this might change in future releases of SQLite.
' Applications that care about shared cache setting should set it explicitly. 


Dim sqlite3_libversion as Function CDecl () as ZString Ptr
Dim sqlite3_initialize as Function CDecl () as Long
Dim sqlite3_shutdown as Function CDecl () as Long
Dim sqlite3_enable_shared_cache as Function CDecl (ByVal as Long) as Long
Dim sqlite3_open_v2 as Function CDecl (ByVal pszfilename as ZString Ptr, ByRef ppDB as sqlite3, ByVal flags as Long, ByVal pszVfs as ZString Ptr) as Long
Dim sqlite3_close as Function CDecl (ByVal pDB as sqlite3) as Long
Dim sqlite3_free as Sub CDecl (ByVal as Any Ptr)
Dim sqlite3_extended_result_codes as Function CDecl (ByVal pDB as sqlite3, ByVal onoff as Long) as Long
Dim sqlite3_extended_errcode as Function CDecl (ByVal pDB as sqlite3) as Long
Dim sqlite3_errmsg as Function CDecl (ByVal pDB as sqlite3) as ZString Ptr
Dim sqlite3_errcode as Function CDecl (ByVal pDB as sqlite3) as Long
Dim sqlite3_backup_init as Function CDecl (ByVal pDest as sqlite3, ByVal zDestName as ZString Ptr, ByVal pSource as sqlite3, _
                                           ByVal zSourceName as ZString Ptr) as sqlite3_backup
Dim sqlite3_backup_step as Function CDecl (ByVal p as sqlite3_backup, ByVal iPage as Long) as Long
Dim sqlite3_backup_finish as Function CDecl (ByVal p as sqlite3_backup) as Long
Dim sqlite3_backup_remaining as Function CDecl (ByVal p as sqlite3_backup) as uLong
Dim sqlite3_backup_pagecount as Function CDecl (ByVal p as sqlite3_backup) as uLong
Dim sqlite3_prepare_v2 as Function CDecl (ByVal pDB as sqlite3, ByVal zSql as ZString Ptr, ByVal nByte as Long, ByRef ppStmt as sqlite3_stmt, _
                                          ByRef pzTail as ZString Ptr) as Long
Dim sqlite3_step as Function CDecl (ByVal pStmt as sqlite3_stmt) as Long
Dim sqlite3_finalize as Function CDecl (ByVal pStmt as sqlite3_stmt) as Long
Dim sqlite3_column_count as Function CDecl (ByVal pStmt as sqlite3_stmt) as Long
Dim sqlite3_column_text as Function CDecl (ByVal pStmt as sqlite3_stmt, _
                                           ByVal iCol as Long) as ZString Ptr
Dim sqlite3_column_blob as Function CDecl (ByVal pStmt as sqlite3_stmt, _
                                           ByVal iCol as Long) as Any Ptr
Dim sqlite3_column_name as Function CDecl (ByVal pStmt as sqlite3_stmt, _
                                           ByVal iCol as Long) as ZString Ptr
Dim sqlite3_column_bytes as Function CDecl (ByVal pStmt as sqlite3_stmt, _
                                            ByVal iCol as Long) as Long
Dim sqlite3_column_type as Function CDecl (ByVal pStmt as sqlite3_stmt, _
                                           ByVal iCol as Long) as Long
Dim sqlite3_mprintf as Function CDecl (ByVal as ZString Ptr, _
                                       ByVal as ZString Ptr) as ZString Ptr
Dim sqlite3_db_config as Function CDecl (ByVal as Long,ByVal op as Any Ptr,ByVal pAny as uLong) as Long                                     

    Private:

    Declare Function SQLExec(ByVal pDB as sqlite3, _
                             ByRef sSQL as String, _
                             ByRef iCols as Long, _
                             ByRef iRows as Long, _
                             ByRef hSpool as HANDLE, _
                             ByRef oSpooler as cCTSQLiteSpooler, _
                             ByRef iFileSize as LongInt, _
                             ByRef iErrorCode as Long, _
                             ByRef sErrorDescription as String, _
                             ByVal lResults as BOOLEAN) as BOOLEAN   
    Declare Function SQLPrepare(ByVal pDB as sqlite3, _
                                ByRef pzSQL as ZString Ptr, _
                                ByRef pStmt as sqlite3_stmt, _
                                ByRef pzTail as ZString Ptr, _
                                ByRef iErrorCode as Long, _
                                ByRef sErrorDescription as String) as BOOLEAN
    Declare Function SQLStep(ByVal pDB as sqlite3, _
                             ByVal pStmt as sqlite3_stmt, _
                             ByRef iCols as Long, _
                             ByRef iRows as Long, _
                             ByRef hSpool as HANDLE, _
                             ByRef oSpooler as cCTSQLiteSpooler, _
                             ByRef iErrorCode as Long, _
                             ByRef sErrorDescription as String, _
                             ByVal lResults as BOOLEAN) as BOOLEAN
    Declare Function BackupInit(ByRef pDest as sqlite3, _
                                ByRef sDestFileName as String, _
                                ByVal pSource as sqlite3, _
                                ByRef pBackup as sqlite3_backup, _
                                ByRef iErrorCode as Long, _
                                ByRef sErrorDescription as String) as BOOLEAN
    Declare Function BackupStep(ByVal pBackup as sqlite3_backup, _
                                ByVal iPages as Long, _
                                ByVal pSource as sqlite3, _
                                ByVal pProgress as Any Ptr) as Long
    Declare Function BackupFinish(ByVal pBackup as sqlite3_backup, _
                                  ByVal pDest as sqlite3, _
                                  ByRef iErrorCode as Long, _
                                  ByRef sErrorDescription as String) as BOOLEAN
    Declare Function ZStringPointer(ByRef sAny as String) as ZString Ptr

    Public:

    Declare Function SetPragma(ByVal pDB as sqlite3, _
                               ByRef oSpooler as cCTSQLiteSpooler, _
                               ByRef iErrorCode as Long, _
                               ByRef sErrorDescription as String, _
                               arPragma() as SqlitePragma) as BOOLEAN 
    Declare Function StartupStatus(ByRef sStartUpErrorDescription as String) as BOOLEAN
    Declare Property Version() as String
    Declare Function OpenDatabase(ByRef sDatabaseName as String, _
                                  ByRef ppDB as sqlite3, _
                                  ByRef iErrorCode as Long, _
                                  ByVal iOpenOptions as Long) as BOOLEAN
    Declare Function CloseDatabase(ByVal pDb as sqlite3, _
                                   ByRef iErrorCode as Long) as BOOLEAN
    Declare Function SQLExecQuery(ByVal pDB as sqlite3, _
                                  ByRef sSQL as String, _
                                  ByRef iCols as Long, _
                                  ByRef iRows as Long, _
                                  ByRef hSpool as HANDLE, _
                                  ByRef oSpooler as cCTSQLiteSpooler, _
                                  ByRef iFileSize as LongInt, _
                                  ByRef iErrorCode as Long, _
                                  ByRef sErrorDescription as String, _
                                  ByVal lResults as BOOLEAN) as BOOLEAN
    Declare Function SQLExecNonQuery(ByVal pDB as sqlite3, _
                                     ByRef sSQL as String, _
                                     ByRef iCols as Long, _
                                     ByRef iRows as Long, _
                                     ByRef hSpool as HANDLE, _
                                     ByRef oSpooler as cCTSQLiteSpooler, _
                                     ByRef iFileSize as LongInt, _                                     
                                     ByRef iErrorCode as Long, _
                                     ByRef sErrorDescription as String, _
                                     ByVal lResults as BOOLEAN) as BOOLEAN
    Declare Function SQLExtendedErrorDescription(ByVal pDB as sqlite3, _
                                                 ByRef iExtendedError as Long) as String
    Declare Function DatabaseBackup(ByVal pSource as sqlite3, _
                                    ByRef sDestFileName as String, _
                                    ByRef iErrorCode as Long, _
                                    ByRef sErrorDescription as String, _
                                    ByVal iPages as Long, _
                                    ByVal pProgress as Any Ptr) as BOOLEAN    
    Declare Sub SafeSQL(ByRef sSQL as String)
    Declare Function DatabaseVersion(ByVal pDB as sqlite3, _
                                     ByRef iVersion as Long, _
                                     ByRef oSpooler as cCTSQLiteSpooler, _
                                     ByRef iErrorCode as Long, _
                                     ByRef sErrorDescription as String) as BOOLEAN                                 
                                       
    Declare Constructor(ByVal pLog as Any Ptr = 0)
    Declare Destructor

End Type

Constructor cCTSQLite(ByVal pLog as Any Ptr = 0)

     This.pSQLite = DyLibLoad("sqlite3.dll")
     
     If This.pSQLite = 0 Then
     
        This.sStartupError = "Library sqlite3.dll not found."
        Exit Constructor
        
     End If
     
     This.sqlite3_db_config = DyLibSymbol(This.pSQLite,"sqlite3_db_config")
     If This.sqlite3_db_config = 0 Then
        This.sStartupError = "API sqlite3_db_config is missing."
        Exit Constructor
     End If
     
' Check if a proc ptr is provided for SQLite global error log

' The Function declaration is:

' Function SQLLogCallback CDecl(ByRef pAny As uLong,ByVal nError As Long,ByVal pMessage As ZString Ptr) as Long

' The Function should be as quick as possible and return a value of 0.

' From SQLite Documentation:

' The error logger callback should be treated like a signal handler. The application should save off or otherwise
' process the error, then return as soon as possible. No other SQLite APIs should be invoked, directly or indirectly, 
' from the error logger. SQLite is not reentrant through the error logger callback. In particular, the error logger 
' callback is invoked when a memory allocation fails, so it is generally a bad idea to try to allocate memory inside 
' the error logger. Do not even think about trying to store the error message in another SQLite database.

' The following Is a partial List of the kinds of messages that might appear in the Error logger callback.

' Any time there Is an error either compiling an SQL statement (Using sqlite3_prepare_v2() Or its siblings) Or running 
' an SQL statement (Using sqlite3_step()) that error is logged. 

' When a schema change occurs that requires a prepared statement To be reparsed And reprepared, that event is logged 
' with the error code SQLITE_SCHEMA. The reparse and reprepare is normally automatic (assuming that sqlite3_prepare_v2() 
' has been used to prepared the statements originally, which is recommended) And so these logging events are normally the 
' only way To know that reprepares are taking place.

' SQLITE_NOTICE messages are logged whenever a database has to be recovered because the previous writer crashed without 
' completing its transaction. The Error code Is SQLITE_NOTICE_RECOVER_ROLLBACK when recovering a rollback journal and 
' SQLITE_NOTICE_RECOVER_WAL when recovering a Write-ahead Log. 

' SQLITE_WARNING messages are logged when database files are renamed Or aliased in ways that can lead To database corruption.

' Out of memory (OOM) error conditions generate error logging events with the SQLITE_NOMEM Error code and a message that says 
' how many bytes of memory were requested by the failed allocation. 

' I/O errors in the OS-interface generate error logging events. The message To these events gives the line number in the source 
' code where the Error originated and the filename associated With the event when there Is a corresponding file. 

' When database corruption is detected, an SQLITE_CORRUPT error logger callback Is invoked. As With I/O errors, the error 
' message text contains the line number in the original source code where the error was first detected.

' An error logger callback is invoked On SQLITE_MISUSE errors. This Is useful in detecting application design issues when 
' return codes are not consistently checked in the application code. 

' SQLite strives to keep error logger traffic low and only send messages to the error logger when there really Is something wrong.

    If pLog <> 0 Then
    
       This.sqlite3_db_config(SQLITE_CONFIG_LOG,pLog,0)
       
    End If 
     
    This.sqlite3_initialize = DyLibSymbol(This.pSQLite,"sqlite3_initialize")
    If This.sqlite3_initialize = 0 Then
       This.sStartupError = "API sqlite3_initialize is missing."
       Exit Constructor
    Else
       iInitialize = This.sqlite3_initialize()
       If iInitialize <> SQLITE_OK Then
          This.sStartupError = Str(iInitialize) + " - Library sqlite3.dll could not be initialized."
          Exit Constructor
       End If        
    End If
        
    This.sqlite3_shutdown = DyLibSymbol(This.pSQLite,"sqlite3_shutdown")
    If This.sqlite3_shutdown = 0 Then
       This.sStartupError = "API sqlite3_shutdown is missing."
       Exit Constructor
    End If
        
    This.sqlite3_enable_shared_cache = DyLibSymbol(This.pSQLite,"sqlite3_enable_shared_cache")
    If This.sqlite3_enable_shared_cache = 0 Then
       This.sStartupError = "API sqlite3_enable_shared_cache is missing."
       Exit Constructor
    Else
       iInitialize = This.sqlite3_enable_shared_cache(SQLITE_TRUE)
       If iInitialize <> SQLITE_OK Then
          This.sStartupError = Str(iInitialize) + " - Enable shared cache failed."
          Exit Constructor
       End If        
    End If
        
    This.sqlite3_open_v2 = DyLibSymbol(This.pSQLite,"sqlite3_open_v2")
    If This.sqlite3_open_v2 = 0 Then
       This.sStartupError = "API sqlite3_open_v2 is missing."
       Exit Constructor
    End If
        
    This.sqlite3_close = DyLibSymbol(This.pSQLite,"sqlite3_close")
    If This.sqlite3_close = 0 Then
       This.sStartupError = "API sqlite3_close is missing."
       Exit Constructor
    End If
        
    This.sqlite3_free = DyLibSymbol(This.pSQLite,"sqlite3_free")
    If This.sqlite3_free = 0 Then
       This.sStartupError = "API sqlite3_free is missing."
       Exit Constructor
    End If
        
    This.sqlite3_extended_result_codes = DyLibSymbol(This.pSQLite,"sqlite3_extended_result_codes")
    If This.sqlite3_extended_result_codes = 0 Then
       This.sStartupError = "API sqlite3_extended_result_codes is missing."
       Exit Constructor
    End If
        
    This.sqlite3_extended_errcode = DyLibSymbol(This.pSQLite,"sqlite3_extended_errcode")
    If This.sqlite3_extended_errcode = 0 Then
       This.sStartupError = "API sqlite3_extended_errcode is missing."
       Exit Constructor
    End If
        
    This.sqlite3_errmsg = DyLibSymbol(This.pSQLite,"sqlite3_errmsg")
    If This.sqlite3_errmsg = 0 Then
       This.sStartupError = "API sqlite3_errmsg is missing."
       Exit Constructor
    End If
        
    This.sqlite3_backup_init = DyLibSymbol(This.pSQLite,"sqlite3_backup_init")
    If This.sqlite3_backup_init = 0 Then
       This.sStartupError = "API sqlite3_backup_init is missing."
       Exit Constructor
    End If
        
    This.sqlite3_backup_step = DyLibSymbol(This.pSQLite,"sqlite3_backup_step")
    If This.sqlite3_backup_step = 0 Then
       This.sStartupError = "API sqlite3_backup_step is missing."
       Exit Constructor
    End If        

    This.sqlite3_backup_finish = DyLibSymbol(This.pSQLite,"sqlite3_backup_finish")
    If This.sqlite3_backup_finish = 0 Then
       This.sStartupError = "API sqlite3_backup_finish is missing."
       Exit Constructor
    End If
        
    This.sqlite3_backup_remaining = DyLibSymbol(This.pSQLite,"sqlite3_backup_remaining")
    If This.sqlite3_backup_remaining = 0 Then
       This.sStartupError = "API sqlite3_backup_remaining is missing."
       Exit Constructor
    End If

    This.sqlite3_backup_pagecount = DyLibSymbol(This.pSQLite,"sqlite3_backup_pagecount")
    If This.sqlite3_backup_pagecount = 0 Then
       This.sStartupError = "API sqlite3_backup_pagecount is missing."
       Exit Constructor
    End If
        
    This.sqlite3_prepare_v2 = DyLibSymbol(This.pSQLite,"sqlite3_prepare_v2")
    If This.sqlite3_prepare_v2 = 0 Then
       This.sStartupError = "API sqlite3_prepare_v2 is missing."
       Exit Constructor
    End If
        
    This.sqlite3_step = DyLibSymbol(This.pSQLite,"sqlite3_step")
    If This.sqlite3_step = 0 Then
       This.sStartupError = "API sqlite3_step is missing."
       Exit Constructor
    End If
        
    This.sqlite3_finalize = DyLibSymbol(This.pSQLite,"sqlite3_finalize")
    If This.sqlite3_finalize = 0 Then
       This.sStartupError = "API sqlite3_finalize is missing."
       Exit Constructor
    End If
        
    This.sqlite3_column_text = DyLibSymbol(This.pSQLite,"sqlite3_column_text")
    If This.sqlite3_column_text = 0 Then
       This.sStartupError = "API sqlite3_column_text is missing."
       Exit Constructor
    End If
        
    This.sqlite3_column_blob = DyLibSymbol(This.pSQLite,"sqlite3_column_blob")
    If This.sqlite3_column_blob = 0 Then
       This.sStartupError = "API sqlite3_column_blob is missing."
       Exit Constructor
    End If
        
    This.sqlite3_column_name = DyLibSymbol(This.pSQLite,"sqlite3_column_name")
    If This.sqlite3_column_name = 0 Then
       This.sStartupError = "API sqlite3_column_name is missing."
       Exit Constructor
    End If
        
    This.sqlite3_column_type = DyLibSymbol(This.pSQLite,"sqlite3_column_type")
    If This.sqlite3_column_type = 0 Then
       This.sStartupError = "API sqlite3_column_type is missing."
       Exit Constructor
    End If
        
    This.sqlite3_column_bytes = DyLibSymbol(This.pSQLite,"sqlite3_column_bytes")
    If This.sqlite3_column_bytes = 0 Then
       This.sStartupError = "API sqlite3_column_bytes is missing."
       Exit Constructor
    End If
        
    This.sqlite3_column_count = DyLibSymbol(This.pSQLite,"sqlite3_column_count")
    If This.sqlite3_column_count = 0 Then
       This.sStartupError = "API sqlite3_column_count is missing."
       Exit Constructor
    End If
        
    This.sqlite3_mprintf = DyLibSymbol(This.pSQLite,"sqlite3_mprintf")
    If This.sqlite3_mprintf = 0 Then
       This.sStartupError = "API sqlite3_mprintf is missing."
       Exit Constructor
    End If
    
    This.sqlite3_errcode = DyLibSymbol(This.pSQLite,"sqlite3_errcode")
    If This.sqlite3_errcode = 0 Then
       This.sStartupError = "API sqlite3_errcode is missing."
       Exit Constructor
    End If                  

    This.sqlite3_libversion = DyLibSymbol(This.pSQLite,"sqlite3_libversion")
    If This.sqlite3_libversion = 0 Then
       This.sStartupError = "API sqlite3_libversion is missing."
       Exit Constructor
    Else
       This.sSQLite3Version = *Cast(ZString Ptr,This.sqlite3_libversion())
    End If
    
    lStartUp = True

End Constructor
Destructor cCTSQLite()

         If This.pSQLite <> 0 Then
         
            If This.sqlite3_initialize <> 0 Then
            
               This.sqlite3_shutdown()
            
            End If
         
            DyLibFree This.pSQLite
            
         End If

End Destructor

' =====================================================================================
' Backup Database
' =====================================================================================
Private Function cCTSQLite.DatabaseBackup (ByVal pSource as sqlite3, _
                                           ByRef sDestFileName as String, _
                                           ByRef iErrorCode as Long, _
                                           ByRef sErrorDescription as String, _
                                           ByVal iPages as Long, _
                                           ByVal pProgress as Any Ptr) as BOOLEAN

Dim pDest              as sqlite3
Dim pBackup            as sqlite3_backup
Dim lResult            as BOOLEAN = False
Dim iFinishError       as Long
Dim sFinishErrorDesc   as String
Dim iStepCode          as Long

    If BackupInit(pDest,sDestFileName,pSource,pBackup,iErrorCode,sErrorDescription) = True Then

       iStepCode = SQLITE_OK
       
       While iStepCode = SQLITE_OK
    
          iStepCode = BackupStep(pBackup,iPages,pSource,pProgress)
          
       Wend
    
       If iStepCode = SQLITE_DONE Then
       
          lResult = BackupFinish(pBackup,pDest,iErrorCode,sErrorDescription)

       Else 

          iErrorCode = iStepCode
          sErrorDescription = SQLExtendedErrorDescription(pSource,iErrorCode)
          BackupFinish(pBackup,pDest,iFinishError,sFinishErrorDesc)

       End If
    
    End If                                               
                                           
    Function = lResult                                       
                                           
End Function                                               
' =====================================================================================
' Initialize Backup Task
' =====================================================================================
Private Function cCTSQLite.BackupInit (ByRef pDest as sqlite3, _
                                       ByRef sDestFileName as String, _
                                       ByVal pSource as sqlite3, _
                                       ByRef pBackup as sqlite3_backup, _
                                       ByRef iErrorCode as Long, _
                                       ByRef sErrorDescription as String) as BOOLEAN
                                       
Dim iCloseError   as Long 

   iErrorCode = 0
   pBackup = 0
   sErrorDescription = "" 
                                       
' If Caller is multithreaded, Backup should be done in a thread

' Open Destination database

    If OpenDatabase(sDestFileName, pDest, iErrorCode, SQLITE_OPEN_READWRITE + SQLITE_OPEN_CREATE) = False Then
    
       Function = False
       Exit Function
       
    End If

' Initialize the Backup process

    pBackup = This.sqlite3_backup_init(pDest,"main",pSource,"main")
    
    If pBackup = 0 Then

       sErrorDescription = SQLExtendedErrorDescription(pDest,iErrorCode)
       
       CloseDatabase(pDest,iCloseError)
       pDest = 0 
    
       Function = False    
    
    Else
    
       Function = True
       
    End If

End Function
' =====================================================================================
' Step through Backup Task
' =====================================================================================
Private Function cCTSQLite.BackupStep (ByVal pBackup as sqlite3_backup, _
                                       ByVal iPages as Long, _
                                       ByVal pSource as sqlite3, _
                                       ByVal pProgress as Any Ptr) as Long
                                       
' Definition for progress callback

Dim Backup_Progress as Sub(ByVal pSource as sqlite3,ByVal iTotalPages as Long,ByVal iRemainingPages as Long)
Dim iStep   as Long
Dim iPageCount   as Long
Dim iRemaining   as Long

    iStep = This.sqlite3_backup_step(pBackup,iPages)
   
    If iStep = SQLITE_OK OrElse iStep = SQLITE_DONE Then
    
       iPageCount = This.sqlite3_backup_pagecount(pBackup)
       iRemaining = This.sqlite3_backup_remaining(pBackup)
    
' Check if a progress callback is provided

       If pProgress <> 0 Then

          Backup_Progress = pProgress
          Backup_Progress(pSource,iPageCount,iRemaining)
     
       End If
        
    End If
    
    Function = iStep                                       
                                       
End Function
' =====================================================================================
' Finish Backup Task
' =====================================================================================
Private Function cCTSQLite.BackupFinish (ByVal pBackup as sqlite3_backup, _
                                         ByVal pDest as sqlite3, _
                                         ByRef iErrorCode as Long, _
                                         ByRef sErrorDescription as String) as BOOLEAN

Dim lResult       as BOOLEAN = False
Dim iCloseError   as Long 

    sErrorDescription = ""
    
    iErrorCode = This.sqlite3_backup_finish(pBackup)
    
    If iErrorCode = SQLITE_OK Then
    
       lResult = True
       
    Else
    
       sErrorDescription = SQLExtendedErrorDescription(pDest,iErrorCode)   
       
    End If
    
    CloseDatabase(pDest,iCloseError)
    
    Function = lResult                                       
                                       
End Function
' =====================================================================================
' Retrieve user defined database version for connection
' =====================================================================================
Private Function cCTSQLite.DatabaseVersion (ByVal pDB as sqlite3, _
                                            ByRef iVersion as Long, _
                                            ByRef oSpooler as cCTSQLiteSpooler, _
                                            ByRef iErrorCode as Long, _
                                            ByRef sErrorDescription as String) as BOOLEAN 
                                            
Dim iCols           as Long
Dim iRows           as Long
Dim hSpool          as HANDLE
Dim iSpoolErrorCode as Long
Dim sSpoolErrorDesc as String
Dim iFileSize       as LongInt
Dim iIndex          as Long
Dim lResult         as BOOLEAN = False
Dim sVersion        as String
Dim lBLOB           as Boolean  

    iVersion = 0 

    If SQLExecQuery(pDB,sPragmaVersion,iCols,iRows,hSpool,oSpooler,iFileSize,iErrorCode,sErrorDescription,True) = True Then
    
       If oSpooler.EndSpoolStreamFile(hSpool,iErrorCode,sErrorDescription) = True Then
       
          For iIndex = 1 To 2
          
              If oSpooler.ReadSpoolerResultBlock(hSpool,lBLOB,sVersion,iErrorCode,sErrorDescription) = False Then
              
                 Exit For
                 
              End If
              
              If iIndex = 2 Then
              
                 iVersion = Val(sVersion)
                 lResult = True
                 
              End If      
       
          Next
       
       End If   
       
    End If
    
    oSpooler.CloseSpoolFile(hSpool,iSpoolErrorCode,sSpoolErrorDesc)
    
    Function = lResult                                             
                                            
End Function
' =====================================================================================
' Set Pragmas for the provided connection
' =====================================================================================
Private Function cCTSQLite.SetPragma (ByVal pDB as sqlite3, _
                                      ByRef oSpooler as cCTSQLiteSpooler, _
                                      ByRef iErrorCode as Long, _
                                      ByRef sErrorDescription as String, _
                                      arPragma() as SqlitePragma) as BOOLEAN
                               
' After a connection has been opened and before it is used
' this function supports using pragmas to change Sqlite defaults

Dim sSQL            as String = ""
Dim iRows           as Long
Dim iCols           as Long
Dim iIndex          as Long
Dim hSpool          as HANDLE
Dim iSpoolErrorCode as Long
Dim sSpoolErrorDesc as String
Dim iFileSize       as LongInt
Dim lResult         as BOOLEAN

    iErrorCode = 0
    sErrorDescription = ""
    
    If LBound(arPragma) < 0 Then
    
       Function = True
       Exit Function
       
    End If
    
    For iIndex = LBound(arPragma) To UBound(arPragma)
    
       sSQL = sSQL _
            + "PRAGMA " _
            + arPragma(iIndex).Name _
            + "=" _
            + arPragma(iIndex).Value _
            + ";"   

    Next
    
    Function = SQLExecNonQuery(pDB,sSQL,iCols,iRows,hSpool,oSpooler,iFileSize,iErrorCode,sErrorDescription,False)

End Function
' =====================================================================================
' Submit Query SQL for execution
' =====================================================================================
Private Function cCTSQLite.SQLExecQuery (ByVal pDB as sqlite3, _
                                         ByRef sSQL as String, _
                                         ByRef iCols as Long, _
                                         ByRef iRows as Long, _
                                         ByRef hSpool as HANDLE, _
                                         ByRef oSpooler as cCTSQLiteSpooler, _
                                         ByRef iFileSize as LongInt, _                                         
                                         ByRef iErrorCode as Long, _
                                         ByRef sErrorDescription as String, _
                                         ByVal lResults as BOOLEAN) as BOOLEAN

Dim iRollbackErrorCode as Long
Dim sRollbackErrorDesc as String
Dim lResult            as BOOLEAN = False
Dim iCommitRows        as Long
Dim iCommitCols        as Long

' If results are to be returned, open spool file

    If lResults = True Then
    
       If oSpooler.CreateSpoolFile(hSpool,iErrorCode,sErrorDescription) = False Then
       
          Function = False
          
          Exit Function
          
       End If
    
    End If 

    sSQL = "BEGIN TRANSACTION;" _
         + Chr(13) + Chr(10) _
         + sSQL

   If SQLExec(pDB,sSQL,iCols,iRows,hSpool,oSpooler,iFileSize,iErrorCode,sErrorDescription,lResults) = False Then
   
      SQLExec(pDB,This.sRollback,iCols,iRows,hSpool,oSpooler,iFileSize,iRollbackErrorCode,sRollbackErrorDesc,False)
      
      If lResults = True Then
      
         oSpooler.CloseSpoolFile(hSpool,iRollbackErrorCode,sRollBackErrorDesc)
         
      End If
      
   Else

      lResult = SQLExec(pDB,This.sCommit,iCommitCols,iCommitRows,hSpool,oSpooler,iFileSize,iErrorCode,sErrorDescription,False)
      
      If lResult = False Then 
      
         If lResults = True Then
      
            oSpooler.CloseSpoolFile(hSpool,iRollbackErrorCode,sRollBackErrorDesc)
            
         End If
         
      Else
   
         If lResults = True Then
      
            lResult = oSpooler.EndSpoolFile(hSpool,iFileSize,iErrorCode,sErrorDescription)
         
            If lResult = False Then
         
               oSpooler.CloseSpoolFile(hSpool,iRollbackErrorCode,sRollBackErrorDesc)
         
            End If
         
         End If
         
      End If
      
   End If
   
   Function = lResult                                   
                                        
End Function
' =====================================================================================
' Submit Query SQL for execution
' =====================================================================================
Private Function cCTSQLite.SQLExecNonQuery (ByVal pDB as sqlite3, _
                                            ByRef sSQL as String, _
                                            ByRef iCols as Long, _
                                            ByRef iRows as Long, _
                                            ByRef hSpool as HANDLE, _
                                            ByRef oSpooler as cCTSQLiteSpooler, _
                                            ByRef iFileSize as LongInt, _
                                            ByRef iErrorCode as Long, _
                                            ByRef sErrorDescription as String, _
                                            ByVal lResults as BOOLEAN) as BOOLEAN

Dim iRollbackErrorCode as Long
Dim sRollbackErrorDesc as String
Dim lResult            as BOOLEAN = False
Dim iCommitRows        as Long
Dim iCommitCols        as Long

' If results are to be returned, open spool file

    If lResults = True Then
    
       If oSpooler.CreateSpoolFile(hSpool,iErrorCode,sErrorDescription) = False Then
       
          Function = False
          
          Exit Function
          
       End If
    
    End If
                                        
   sSQL = "BEGIN IMMEDIATE TRANSACTION; " _
        + Chr(13) + Chr(10) _
        + sSQL
        
   If SQLExec(pDB,sSQL,iCols,iRows,hSpool,oSpooler,iFileSize,iErrorCode,sErrorDescription,lResults) = False Then
   
      SQLExec(pDB,This.sRollback,iCols,iRows,hSpool,oSpooler,iFileSize,iRollbackErrorCode,sRollbackErrorDesc,False)
      
      If lResults = True Then
      
         oSpooler.CloseSpoolFile(hSpool,iRollbackErrorCode,sRollBackErrorDesc)
         
      End If
      
   Else

      lResult = SQLExec(pDB,This.sCommit,iCommitCols,iCommitRows,hSpool,oSpooler,iFileSize,iErrorCode,sErrorDescription,False)
      
      If lResult = False Then 
      
         If lResults = True Then
      
            oSpooler.CloseSpoolFile(hSpool,iRollbackErrorCode,sRollBackErrorDesc)
            
         End If
         
      Else
   
         If lResults = True Then
      
            lResult = oSpooler.EndSpoolFile(hSpool,iFileSize,iErrorCode,sErrorDescription)
         
            If lResult = False Then
         
               oSpooler.CloseSpoolFile(hSpool,iRollbackErrorCode,sRollBackErrorDesc)
         
            End If
         
         End If
         
      End If
      
   End If
   
   Function = lResult                                    
                                        
End Function
' =====================================================================================
' Submit SQL for execution
' =====================================================================================
Private Function cCTSQLite.SQLExec (ByVal pDB as sqlite3, _
                                    ByRef sSQL as String, _
                                    ByRef iCols as Long, _
                                    ByRef iRows as Long, _
                                    ByRef hSpool as HANDLE, _
                                    ByRef oSpooler as cCTSQLiteSpooler, _
                                    ByRef iFileSize as LongInt, _
                                    ByRef iErrorCode as Long, _
                                    ByRef sErrorDescription as String, _
                                    ByVal lResults as BOOLEAN) as BOOLEAN

Dim iIndex          as Long
Dim iColIndex       as Long
Dim pStmt           as sqlite3_stmt = 0
Dim lResult         as BOOLEAN = True
Dim lStep           as BOOLEAN = True
Dim lPrepare        as BOOLEAN = True
Dim iStepCode       as Long
Dim pzTail          as ZString Ptr
Dim pzSQL           as ZString Ptr
Dim iSpoolErrorCode as Long
Dim sSpoolErrorDesc as String

    sErrorDescription = ""
    iErrorCode = 0
    iFileSize = 0
    iCols = 0
    iRows = 0
   
' Exec SQL presented

    pzSQL = ZStringPointer(sSQL)
   
' Loop though possible multiple SQL statements

    While lPrepare = True

       If SQLPrepare(pDB,pzSQL,pStmt,pzTail,iErrorCode,sErrorDescription) = False Then

          lPrepare = False
          lResult = False
          
       Else

          If SQLStep(pDB,pStmt,iCols,iRows,hSpool,oSpooler,iErrorCode,sErrorDescription,lResults) = False Then

             This.sqlite3_finalize(pStmt)
             lPrepare = False
             lResult = False
          
          Else

             If *Cast(ZString Ptr,pzTail) = "" Then

                This.sqlite3_finalize(pStmt)
                lPrepare = False
             
            Else

               This.sqlite3_finalize(pStmt)
            
               pzSQL = pzTail
            
            End If
            
         End If   
       
       End If         
   
    Wend

   Function = lResult
                                    
End Function
' =====================================================================================
' Prepare SQL for execution
' =====================================================================================
Private Function cCTSQLite.SQLPrepare (ByVal pDB as sqlite3, _
                                       ByRef pzSQL as ZString Ptr, _
                                       ByRef pStmt as sqlite3_stmt, _
                                       ByRef pzTail as ZString Ptr, _
                                       ByRef iErrorCode as Long, _
                                       ByRef sErrorDescription as String) as BOOLEAN

    sErrorDescription = ""
                                       
    iErrorCode = This.sqlite3_prepare_v2(pDB,pzSQL,-1,pStmt,pzTail)
    
    If iErrorCode = SQLITE_OK Then
    
       Function = True
       
    Else                                           
                                    
       sErrorDescription = SQLExtendedErrorDescription(pDB,iErrorCode)
       
       Function = False
       
    End If
                                        
End Function
' =====================================================================================
' Prepare SQL for execution
' =====================================================================================
Private Function cCTSQLite.SQLStep (ByVal pDB as sqlite3, _
                                    ByVal pStmt as sqlite3_stmt, _
                                    ByRef iCols as Long, _
                                    ByRef iRows as Long, _
                                    ByRef hSpool as HANDLE, _
                                    ByRef oSpooler as cCTSQLiteSpooler, _                                    
                                    ByRef iErrorCode as Long, _
                                    ByRef sErrorDescription as String, _
                                    ByVal lResults as BOOLEAN) as BOOLEAN

Dim lStep       as BOOLEAN = True
Dim lResult     as BOOLEAN = False
Dim iIndex      as Long
Dim iBLOBSize   as Long
Dim pzString    as ZString Ptr

    sErrorDescription = ""
    iCols = 0
    iRows = 0
    iErrorCode = 0

    While lStep = True

       iErrorCode = This.sqlite3_step(pStmt)

       Select Case iErrorCode

' Row results are available

          Case SQLITE_ROW

             If lResults = True Then
             
                If iCols = 0 Then

' Get Column Headings

                   iCols = This.sqlite3_column_count(pStmt)

                   If iCols > 0 Then

                      For iIndex = 0 To iCols - 1
                      
                          pzString = Cast(ZString Ptr,This.sqlite3_column_name(pStmt,iIndex))
                     
                          If oSpooler.WriteSpoolerResultBlock(hSpool,False,Len(*pzString),Cast(Any Ptr,pzString), _
                                                              iErrorCode,sErrorDescription) = False Then
                                                              
                             Function = False
                             Exit Function
                             
                          End If

                      Next
          
                     iRows = 1

                  End If
       
                End If
                
' Get column results if available                
             
                If iCols > 0 Then                

                   For iIndex = 0 To iCols - 1
                
                      Select Case This.sqlite3_column_type(pStmt,iIndex)
                      
' Column is NULL

                         Case SQLITE_NULL

                             If oSpooler.WriteSpoolerResultBlock(hSpool,False,0,0, _
                                                                 iErrorCode,sErrorDescription) = False Then
                                                                 
                                Function = False
                                Exit Function
                             
                             End If
                             
' Column is a BLOB
                      
                         Case SQLITE_BLOB

                            iBLOBSize = This.sqlite3_column_bytes(pStmt,iIndex)
                            
                            If iBLOBSize > 0 Then
                            
                               If oSpooler.WriteSpoolerResultBlock(hSpool,True,iBLOBSize,Cast(Any Ptr,This.sqlite3_column_blob(pStmt,iIndex)), _
                                                                   iErrorCode,sErrorDescription) = False Then
                                                                   
                                  Function = False
                                  Exit Function
                             
                               End If

                            Else

                               If oSpooler.WriteSpoolerResultBlock(hSpool,True,Len(*pzString),0,iErrorCode,sErrorDescription) = False Then
                               
                                  Function = False
                                  Exit Function
                             
                               End If                          
                            
                           End If
                      
                         Case Else
                         
' Cast column result as a string
                         
                            pzString = Cast(ZString Ptr,This.sqlite3_column_text(pStmt,iIndex))
                         
                            If oSpooler.WriteSpoolerResultBlock(hSpool,True,Len(*pzString),Cast(Any Ptr,pzString),iErrorCode,sErrorDescription) = False Then
                               
                                  Function = False
                                  Exit Function
                             
                            End If
                      
                      End Select

                   Next
                
                   iRows = iRows + 1
                
                End If
                
             End If 
             
' No more rows are available               

        Case SQLITE_DONE

             iErrorCode = 0
             lResult = True
             lStep = False

        Case Else 

           sErrorDescription = SQLExtendedErrorDescription(pDB,iErrorCode) 
           lStep = False

       End Select

    Wend

    Function = lResult

End Function                                    
' =====================================================================================
' Open Database
' =====================================================================================
Private Function cCTSQLite.OpenDatabase (ByRef sDatabaseName as String, _
                                         ByRef ppDB as sqlite3, _
                                         ByRef iErrorCode as Long, _
                                         ByVal iOpenOptions as Long) as BOOLEAN
                                         
Dim lResult                  as BOOLEAN
Dim iCloseError              as Long

    iErrorCode = This.sqlite3_open_v2(ZStringPointer(sDatabaseName),ppDB,iOpenOptions,0)

    lResult = IIf(iErrorCode = SQLITE_OK,True,False)
    
    If lResult = False Then
    
' SQLite Documentation

' A database connection handle is usually returned even if an error occurs. The only exception is that
' if SQLite is unable to allocate memory to hold the connection handle, a NULL will be written

       If ppDB <> 0 Then    
          CloseDatabase(ppDB,iCloseError)
          ppDB = 0
       End If
       
    End If
    
    Function = lResult

End Function
' =====================================================================================
' Close Database
' =====================================================================================
Private Function cCTSQLite.CloseDatabase (ByVal pDB as sqlite3, _
                                          ByRef iErrorCode as Long) as BOOLEAN

    iErrorCode = This.sqlite3_close(pDB)
    
    Function = IIf(iErrorCode = SQLITE_OK,True,False)

End Function
' =====================================================================================
' Get Server Startup Status
' =====================================================================================
Private Function cCTSQLite.StartupStatus (ByRef sStartUpErrorDescription as String) as BOOLEAN

    sStartUpErrorDescription = This.sStartUpError
    
    Function = This.lStartup

End Function
' =====================================================================================
' Get SQLite Library Version
' =====================================================================================
Private Property cCTSQLite.Version () as String
   
    Property = This.sSQLite3Version

End Property
' =====================================================================================
' Get SQLite Extended Error code and description
' =====================================================================================
Function cCTSQLite.SQLExtendedErrorDescription(ByVal pDB as sqlite3, _
                                               ByRef iExtendedError as Long) as String

Dim sErrorDescription    as String
Dim iErrorCode           as Long

   sErrorDescription = *Cast(ZString Ptr,This.sqlite3_errmsg(pDB))
   
   iErrorCode = This.sqlite3_errcode(pDB)

   iExtendedError = This.sqlite3_extended_errcode(pDB)

   iExtendedError = IIf(iExtendedError = SQLITE_OK,iErrorCode,iExtendedError)
   
   iExtendedError = IIf(iExtendedError <> iErrorCode,iExtendedError,iErrorCode)
   
   Function = sErrorDescription

End Function
' =====================================================================================
' Escape SQL for single apostrophe
' =====================================================================================
Sub cCTSQLite.SafeSQL (ByRef sSQL as String)

Dim pzSQL as ZString Ptr

   If Len(sSQL) > 0 Then

      pzSql = This.sqlite3_mprintf("%q",ZStringPointer(sSQL))
   
      sSQL = *Cast(ZString Ptr,pzSQL)

      This.sqlite3_free(pzSql)
      
   End If
     
End Sub
' =====================================================================================
' Get ZString Pointer from a STRING type
' =====================================================================================
Function cCTSQLite.ZStringPointer (ByRef sAny as String) as ZString Ptr

    If StrPtr(sAny) = 0 Then
    
       Function = @""
       
    Else
    
       Function = StrPtr(sAny)
       
    End If

End Function