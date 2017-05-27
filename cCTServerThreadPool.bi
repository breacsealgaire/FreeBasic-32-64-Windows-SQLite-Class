' ########################################################################################
' File: cCTServerThreadPool.bi
' Contents: FreeBasic Windows SQLite Client Server Thread Pool class support.
' Version: 1.00
' Compiler: FreeBasic 32 & 64-bit Windows
' Copyright (c) 2017 Rick Kelly
' Released into the public domain for private and public use without restriction
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#Pragma Once

#IFNDEF _WIN32_WINNT
const _WIN32_WINNT = &h0602
#ENDIF

#Include Once "windows.bi"


' ########################################################################################
' cCTServerThreadPool Class
' ########################################################################################

Type cCTServerThreadPool Extends Object

    Private:
    
Dim lStartUp         as BOOLEAN = False
Dim iStartUpError    as Long = 0
Dim reserved         as PVOID = NULL
Dim ucbe             as TP_CALLBACK_ENVIRON
Dim cbe              as PTP_CALLBACK_ENVIRON
Dim ptpp             as PTP_POOL
Dim ptpcg            as PTP_CLEANUP_GROUP   

    Public:


    Declare Sub CreateThreadWork(ByVal pfnwk as PTP_WORK_CALLBACK, _
                                 ByVal pv as PVOID)
    Declare Sub ShutdownThreadPool ()
    Declare Function StartupStatus(ByRef iStartUpError as Long) as BOOLEAN
                                       
    Declare Constructor
    Declare Destructor

End Type

Constructor cCTServerThreadPool()

    This.cbe = cast(PTP_CALLBACK_ENVIRON,varptr(This.ucbe))
    This.ptpp = CreateThreadpool(This.reserved)
    This.iStartUpError = GetLastError()

    If This.ptpp <> 0 Then

       TpInitializeCallbackEnviron(This.cbe)
       TpSetCallbackLongFunction(This.cbe)

       This.ptpcg = CreateThreadpoolCleanupGroup()
       This.iStartUpError = GetLastError()       

       if This.ptpcg <> 0 Then

          TpSetCallbackCleanupGroup(This.cbe,This.ptpcg,NULL)

          lStartUp = True

       End If

    End IF

End Constructor
Destructor cCTServerThreadPool()

    If This.ptpcg <> 0 Then

       CloseThreadpoolCleanupGroupMembers(ByVal This.ptpcg,0,NULL)
       CloseThreadpoolCleanupGroup(ByVal This.ptpcg)

    End If    

    TpDestroyCallbackEnviron(This.cbe)

    If This.ptpp <> 0 Then

       CloseThreadpool(This.ptpp)

    End If

End Destructor

' =====================================================================================
' Create Thread Work
' =====================================================================================
Private Sub cCTServerThreadPool.CreateThreadWork (ByVal pfnwk as PTP_WORK_CALLBACK, _
                                                  ByVal pv as PVOID)
                                                      
Dim Work          as PTP_WORK

    Work = CreateThreadpoolWork(pfnwk,pv,This.cbe)
    SubmitThreadpoolWork(Work)
    CloseThreadpoolWork(Work)                                                    
                                                      
End Sub
' =====================================================================================
' Close Thread Pool
' =====================================================================================
Private Sub cCTServerThreadPool.ShutdownThreadPool ()

' Block until all outstanding callbacks/threads have completed

    If This.ptpcg <> 0 Then

       CloseThreadpoolCleanupGroupMembers(ByVal This.ptpcg,0,NULL)
       CloseThreadpoolCleanupGroup(ByVal This.ptpcg)
       This.ptpcg = 0

    End If

End Sub
' =====================================================================================
' Get Thread Pool Startup Status
' =====================================================================================
Private Function cCTServerThreadPool.StartupStatus (ByRef iStartUpError as Long) as BOOLEAN

    iStartUpError = This.iStartUpError
    
    Function = This.lStartup

End Function