' ########################################################################################
' File: cCTSQLiteSpooler.bi
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

' ########################################################################################
' cCTSQLiteSpooler Class
' ########################################################################################

Private Const SQLITESPOOLER_SIGNATURE                = &h012f38f7

Type cCTSQLiteSpooler Extends Object

Private:

Dim iSignature    as Long = SQLITESPOOLER_SIGNATURE

    Declare Function SpoolerFileName() as String
    Declare Function SpoolerErrorDescription(ByVal iErrorCode as Long) as String

    Declare Function SpoolerWrite(ByVal hFile as HANDLE, _
                                  ByVal lpBuffer as Any Ptr, _
                                  ByVal iNumberOfBytesToWrite as ulong, _
                                  ByRef iErrorCode as Long, _
                                  ByRef sErrorDescription as String) as BOOLEAN
    Declare Function SpoolerRewindFile(ByVal hSpool as HANDLE, _
                                       ByVal iPosition as Long, _
                                       ByRef iErrorCode as Long, _
                                       ByRef sErrorDescription as String) as BOOLEAN
    Declare Function SpoolerSize(ByVal hFile as HANDLE, _
                                 ByRef iFileSize as LongInt, _
                                 ByRef iErrorCode as Long, _
                                 ByRef sErrorDescription as String) as BOOLEAN

Public:

    Declare Function CreateSpoolFile(ByRef hSpool as HANDLE, _
                                     ByRef iErrorCode as Long, _
                                     ByRef sErrorDescription as String) as BOOLEAN
    Declare Function CreateSpoolStreamFile(ByRef hSpool as HANDLE, _
                                           ByRef iErrorCode as Long, _
                                           ByRef sErrorDescription as String) as BOOLEAN
    Declare Function EndSpoolFile(ByVal hSpool as HANDLE, _
                                  ByRef iFileSize as LongInt, _
                                  ByRef iErrorCode as Long, _
                                  ByRef sErrorDescription as String) as BOOLEAN
    Declare Function EndSpoolStreamFile(ByVal hSpool as HANDLE, _
                                        ByRef iErrorCode as Long, _
                                        ByRef sErrorDescription as String) as BOOLEAN
    Declare Function CloseSpoolFile(ByVal hSpool as HANDLE, _
                                    ByRef iErrorCode as Long, _
                                    ByRef sErrorDescription as String) as BOOLEAN
    Declare Function WriteSpoolerResultBlock(ByVal hSpool as HANDLE, _
                                             ByVal lBLOB as BOOLEAN, _
                                             ByVal iSQLiteColumnSize as Long, _
                                             ByVal pSQLiteColumnValue as Any Ptr, _
                                             ByRef iErrorCode as Long, _
                                             ByRef sErrorDescription as String) as BOOLEAN
    Declare Function ReadSpoolerResultBlock(ByVal hSpool as HANDLE, _
                                            ByRef lBLOB as BOOLEAN, _
                                            ByRef sResultValue as String, _
                                            ByRef iErrorCode as Long, _
                                            ByRef sErrorDescription as String) as BOOLEAN                                            
    Declare Function ReadSpoolFile(ByVal hSpool as HANDLE, _
                                   ByVal iBytesToRead as Long, _
                                   ByRef sSpoolData as String, _
                                   ByRef iErrorCode as Long, _
                                   ByRef sErrorDescription as String) as BOOLEAN

    Declare Constructor()
    Declare Destructor

End Type

Constructor cCTSQLiteSpooler()

End Constructor
Destructor cCTSQLiteSpooler()

End Destructor

' =====================================================================================
' Create Spooler file
' =====================================================================================
Private Function cCTSQLiteSpooler.CreateSpoolFile (ByRef hSpool as HANDLE, _
                                                   ByRef iErrorCode as Long, _
                                                   ByRef sErrorDescription as String) as BOOLEAN
Dim sSpoolName   as String = SpoolerFileName()

    sErrorDescription = ""                                    
 
    hSpool = CreateFile(StrPtr(sSpoolName),GENERIC_READ Or GENERIC_WRITE,0,0,CREATE_ALWAYS,FILE_ATTRIBUTE_TEMPORARY Or FILE_FLAG_DELETE_ON_CLOSE,0)
    
    iErrorCode = GetLastError()
    
    If hSpool = INVALID_HANDLE_VALUE Then
    
       sErrorDescription = SpoolerErrorDescription(iErrorCode)
 
       Function = False    
    
    Else

       Function = SpoolerWrite(hSpool,Cast(Any Ptr,VarPtr(This.iSignature)),Len(This.iSignature),iErrorCode,sErrorDescription)
       
    End If   
                                    
End Function
' =====================================================================================
' Create Spooler stream (empty) file
' =====================================================================================
Private Function cCTSQLiteSpooler.CreateSpoolStreamFile (ByRef hSpool as HANDLE, _
                                                         ByRef iErrorCode as Long, _
                                                         ByRef sErrorDescription as String) as BOOLEAN

Dim sSpoolName   as String = SpoolerFileName()

    sErrorDescription = ""                                    

    hSpool = CreateFile(StrPtr(sSpoolName),GENERIC_READ Or GENERIC_WRITE,0,0,CREATE_ALWAYS,FILE_ATTRIBUTE_TEMPORARY Or FILE_FLAG_DELETE_ON_CLOSE,0)
    
    iErrorCode = GetLastError()
    
    If hSpool = INVALID_HANDLE_VALUE Then
    
       sErrorDescription = SpoolerErrorDescription(iErrorCode)
 
       Function = False    
    
    Else
 
       Function = True
       
    End If   
                                    
End Function
' =====================================================================================
' Spooler write one result block
' =====================================================================================
Private Function cCTSQLiteSpooler.WriteSpoolerResultBlock (ByVal hSpool as HANDLE, _
                                                           ByVal lBLOB as BOOLEAN, _
                                                           ByVal iSQLiteColumnSize as Long, _
                                                           ByVal pSQLiteColumnValue as Any Ptr, _
                                                           ByRef iErrorCode as Long, _
                                                           ByRef sErrorDescription as String) as BOOLEAN

Dim lResult   as BOOLEAN
                                                      
' Write out the BLOB flag

    lResult = SpoolerWrite(hSpool,Cast(Any Ptr,VarPtr(lBLOB)),Len(lBLOB),iErrorCode,sErrorDescription)
    
    If lResult = True Then
    
       lResult = SpoolerWrite(hSpool,Cast(Any Ptr,VarPtr(iSQLiteColumnSize)),Len(iSQLiteColumnSize),iErrorCode,sErrorDescription)
       
       If lResult = True Then
       
          lResult = SpoolerWrite(hSpool,pSQLiteColumnValue,iSQLiteColumnSize,iErrorCode,sErrorDescription)
          
       End If
       
    End If
     
    Function = lResult
                                                       
End Function
' =====================================================================================
' Spooler Read Result Block and decode
' =====================================================================================
Private Function cCTSQLiteSpooler.ReadSpoolerResultBlock (ByVal hSpool as HANDLE, _
                                                          ByRef lBLOB as BOOLEAN, _
                                                          ByRef sResultValue as String, _
                                                          ByRef iErrorCode as Long, _
                                                          ByRef sErrorDescription as String) as BOOLEAN
                                                          
Dim sResultBlock    as String
Dim iResultSize     as Long
                                                          
' Read next block of 5 bytes

    If ReadSpoolFile(hSpool,5,sResultBlock,iErrorCode,sErrorDescription) = False Then
    
       Function = False
       
       Exit Function
       
    End If
    
' If less than 5 bytes, spool file is corrupt/incomplete

    If Len(sResultBlock) <> 5 Then
    
       iErrorCode = -4
       
       sErrorDescription = "Spool file is corrupt/incomplete."
       
       Function = False
       
       Exit Function
       
    End If
    
    lBLOB = *(Cast(BOOLEAN Ptr,StrPtr(sResultBlock)))
    iResultSize = *(Cast(Long Ptr,StrPtr(sResultBlock) + 1))
    
' Check if might have hit the EOF marker
    
    If iResultSize < 0 Then

       iErrorCode = -5
       
       sErrorDescription = "Attempt to read past last result block."
       
       Function = False
       
       Exit Function    
    
    End If
    
    Function = ReadSpoolFile(hSpool,iResultSize,sResultValue,iErrorCode,sErrorDescription)                                                      
                                                          
End Function                                                          
' =====================================================================================
' Spooler Read
' =====================================================================================
Private Function cCTSQLiteSpooler.ReadSpoolFile (ByVal hSpool as HANDLE, _
                                                 ByVal iBytesToRead as Long, _
                                                 ByRef sSpoolData as String, _
                                                 ByRef iErrorCode as Long, _
                                                 ByRef sErrorDescription as String) as BOOLEAN

Dim iResult    as WINBOOL
Dim iBytesRead as Long


    sErrorDescription = ""
    sSpoolData = Space(iBytesToRead)

    iResult = ReadFile(hSpool,StrPtr(sSpoolData),iBytesToRead,VarPtr(iBytesRead),0) 
    
    iErrorCode = GetLastError()
    
    If iResult = 0 Then
    
       sErrorDescription = SpoolerErrorDescription(iErrorCode)
           
       Function = False
       
    Else

       sSpoolData = Left(sSpoolData,iBytesRead)    
       Function = True
       
    End If
                                                       
End Function 
' =====================================================================================
' End Spooler file
' =====================================================================================
Private Function cCTSQLiteSpooler.EndSpoolFile (ByVal hSpool as HANDLE, _
                                                ByRef iFileSize as LongInt, _
                                                ByRef iErrorCode as Long, _
                                                ByRef sErrorDescription as String) as BOOLEAN
                                                   
' When SQLite step is done, this function is called to add the EOF block and rewind the spool file to the beginning

Dim iEOF    as LongInt = 1099511627520       ' Written to spool file as 00FFFFFFFF, length < 0 is end of file

    sErrorDescription = ""
    
    If SpoolerWrite(hSpool,Cast(Any Ptr,VarPtr(iEOF)),5,iErrorCode,sErrorDescription) = True Then
    
       FlushFileBuffers(hSpool)

       If SpoolerRewindFile(hSpool,0,iErrorCode,sErrorDescription) = True Then
       
          If SpoolerSize(hSpool,iFileSize,iErrorCode,sErrorDescription) = True Then
    
             Function = True
             
          Else
          
             CloseHandle(hSpool)
          
             Function = False
             
          End If
          
       Else
       
          CloseHandle(hSpool)
          
          Function = False
          
       End If
       
    Else
    
       CloseHandle(hSpool)
       
       Function = False
       
    End If
                                    
End Function
' =====================================================================================
' End Spooler file
' =====================================================================================
Private Function cCTSQLiteSpooler.EndSpoolStreamFile (ByVal hSpool as HANDLE, _
                                                      ByRef iErrorCode as Long, _
                                                      ByRef sErrorDescription as String) as BOOLEAN
                                                   
' Rewind the spool file to the beginning and check for a valid signature block

Dim sSpoolSignature   as String
Dim lResult           as BOOLEAN = False  

    sErrorDescription = ""
    
    FlushFileBuffers(hSpool)

    If SpoolerRewindFile(hSpool,0,iErrorCode,sErrorDescription) = True Then
       
       If ReadSpoolFile(hSpool,Len(iSignature),sSpoolSignature,iErrorCode,sErrorDescription) = True Then
       
          If This.iSignature = *(Cast(Long Ptr,StrPtr(sSpoolSignature))) Then
          
             lResult = True
             
          Else
          
             iErrorCode = -2
             sErrorDescription = "Not a recognized SQLite spool file."   
             
          End If
          
       End If
       
    End If
    
    If lResult = False Then

       CloseHandle(hSpool)
       
    End If
          
    Function = lResult
                                    
End Function
' =====================================================================================
' Close Spooler file
' =====================================================================================
Private Function cCTSQLiteSpooler.CloseSpoolFile (ByVal hSpool as HANDLE, _
                                                  ByRef iErrorCode as Long, _
                                                  ByRef sErrorDescription as String) as BOOLEAN
                                                 
Dim iResult    as WINBOOL

    sErrorDescription = ""

    iResult = CloseHandle(hSpool) 
    
    iErrorCode = GetLastError()
    
    If iResult = 0 Then
    
       sErrorDescription = SpoolerErrorDescription(iErrorCode)
           
       Function = False
       
    Else
    
       Function = True
       
    End If
    
End Function
' =====================================================================================
' Rewind Spooler file to first results block
' =====================================================================================
Private Function cCTSQLiteSpooler.SpoolerRewindFile (ByVal hSpool as HANDLE, _
                                                     ByVal iPosition as Long, _
                                                     ByRef iErrorCode as Long, _
                                                     ByRef sErrorDescription as String) as BOOLEAN

Dim iResult    as Long

    sErrorDescription = ""

    iResult = SetFilePointer(hSpool,iPosition,0,FILE_BEGIN)
    
    iErrorCode = GetLastError()
    
    If iResult = INVALID_SET_FILE_POINTER Then
    
       sErrorDescription = SpoolerErrorDescription(iErrorCode)
           
       Function = False
       
    Else
    
       Function = True
       
    End If
                                                   
End Function   
' =====================================================================================
' Get Spooler file folder and name
' =====================================================================================
Private Function cCTSQLiteSpooler.SpoolerFileName () as String

Dim sTempPath      as String * MAX_PATH
Dim sTempFileName  as String * MAX_PATH
Dim iPathLength    as Long
Dim sPrefix        as String = "cCT"

    iPathLength = GetTempPath(MAX_PATH, sTempPath)
    
    sTempPath = Left(sTempPath,iPathLength)
    
    GetTempFileName(StrPtr(sTempPath),StrPtr(sPrefix),0,StrPtr(sTempFileName))
    
    Function = sTempFileName
    
End Function
' =====================================================================================
' Write to Spooler
' =====================================================================================
Private Function cCTSQLiteSpooler.SpoolerWrite (ByVal hFile as HANDLE, _
                                                ByVal lpBuffer as Any Ptr, _
                                                ByVal iNumberOfBytesToWrite as ulong, _
                                                ByRef iErrorCode as Long, _
                                                ByRef sErrorDescription as String) as BOOLEAN

Dim iResult    as WINBOOL

    sErrorDescription = ""
    iErrorCode = 0

    if lpBuffer <> 0 then

       iResult = WriteFile(hFile,lpBuffer,iNumberOfBytesToWrite,0,0)
    
       iErrorCode = GetLastError()
    
       If iResult = 0 Then
    
          sErrorDescription = SpoolerErrorDescription(iErrorCode)
           
          Function = False
       
       Else
    
          Function = True

       end if

    else

       function = true       

    End If

End Function
' =====================================================================================
' Write to Spooler
' =====================================================================================
Private Function cCTSQLiteSpooler.SpoolerSize (ByVal hFile as HANDLE, _
                                               ByRef iFileSize as LongInt, _
                                               ByRef iErrorCode as Long, _
                                               ByRef sErrorDescription as String) as BOOLEAN

Dim iResult    as WINBOOL

    sErrorDescription = ""


    iResult = GetFileSizeEx(hFile,Cast(LARGE_INTEGER Ptr,VarPtr(iFileSize)))
    
    iErrorCode = GetLastError()
    
    If iResult = 0 Then
    
       sErrorDescription = SpoolerErrorDescription(iErrorCode)
           
       Function = False
       
    Else
    
       Function = True
       
    End If

End Function                                                 
' =====================================================================================
' Get Windows error description
' =====================================================================================
Private Function cCTSQLiteSpooler.SpoolerErrorDescription (ByVal iErrorCode as Long) as String

Dim sErrorDescription as String * 255

    FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0, iErrorCode, 0, sErrorDescription, SizeOf(sErrorDescription), ByVal 0)
    
    Function = sErrorDescription
    
End Function