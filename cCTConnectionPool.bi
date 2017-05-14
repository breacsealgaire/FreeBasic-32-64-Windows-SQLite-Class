' ########################################################################################
' File: cCTConnectionPool.inc
' Contents: FreeBasic Windows SQLite Client Server connection pool class support.
' Version: 1.00
' Compiler: FreeBasic 32 & 64-bit Windows
' Copyright (c) 2017 Rick Kelly
' Released into the public domain for private and public use without restriction
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#Pragma Once

#Include Once "windows.bi"

Private Type POOL_CONNECTION

      DataBaseID          as Long = 0
      ConnectionID        as uLong = 0
      ConnectionInUse     as BOOLEAN = False
      ReleaseCount        as Long = 0
      RequestTime         as Double = 0 
      RequestTotalTime    as Double = 0
      SlotFree            as BOOLEAN = True 
      
End Type

' ########################################################################################
' cCTConnectionPool Class
' ########################################################################################

Type cCTConnectionPool Extends Object

    Private:
    
Dim lpCriticalSection as CRITICAL_SECTION
Dim arConnection(0 To 99) as POOL_CONNECTION    

    Public:

    Declare Function AddPoolConnection(ByVal iDatabaseID as Long, _
                                       ByVal iConnectionID as uLong) as BOOLEAN
    Declare Function ReleasePoolConnection(ByVal iDatabaseID as Long, _
                                           ByVal iConnectionID as uLong) as BOOLEAN
    Declare Function RequestPoolConnection(ByVal iDatabaseID as Long, _
                                           ByRef iConnectionID as uLong) as BOOLEAN
    Declare Function RemovePoolConnection(ByVal iDatabaseID as Long, _
                                          ByVal iConnectionID as uLong) as BOOLEAN
    Declare Sub PoolConnectionStatistics(ByRef iPoolSize as Long, _
                                         ByRef iTotalCurrentConnections as Long, _
                                         ByRef iTotalActiveConnections as Long, _
                                         ByRef nAverageRequestMilliseconds as Double)
    Declare Sub PoolConnectionDetail(arPool() as POOL_CONNECTION)
                                       
    Declare Constructor
    Declare Destructor

End Type

Constructor cCTConnectionPool()

    InitializeCriticalSection(ByVal VarPtr(lpCriticalSection))

End Constructor
Destructor cCTConnectionPool()

    DeleteCriticalSection(ByVal VarPtr(lpCriticalSection))

End Destructor

' =====================================================================================
' Add pool connection
' =====================================================================================
Private Function cCTConnectionPool.AddPoolConnection (ByVal iDatabaseID as Long, _
                                                      ByVal iConnectionID as uLong) as BOOLEAN
                                                      
Dim lReturn       as BOOLEAN
Dim iIndex        as Long 

    lReturn = False
    EnterCriticalSection(@This.lpCriticalSection)
    
    For iIndex = 0 To UBound(This.arConnection)
    
        If This.arConnection(iIndex).SlotFree = True Then
           This.arConnection(iIndex).DatabaseID = iDatabaseID
           This.arConnection(iIndex).ConnectionID = iConnectionID
           This.arConnection(iIndex).ConnectionInUse = True
           This.arConnection(iIndex).ReleaseCount = 0
           This.arConnection(iIndex).RequestTime = Timer
           This.arConnection(iIndex).RequestTotalTime = 0
           This.arConnection(iIndex).SlotFree = False
           lReturn = True
           
           Exit For
             
        End If
    
    Next 
                
    LeaveCriticalSection(@This.lpCriticalSection)

    Function = lReturn                                                    
                                                      
End Function
' =====================================================================================
' Release pool connection
' =====================================================================================
Private Function cCTConnectionPool.ReleasePoolConnection (ByVal iDatabaseID as Long, _
                                                          ByVal iConnectionID as uLong) as BOOLEAN
                                                      
Dim lReturn       as BOOLEAN
Dim iIndex        as Long
Dim nTimer        as Double

    lReturn = False
    nTimer = Timer
    EnterCriticalSection(@This.lpCriticalSection)
    
    For iIndex = 0 To UBound(This.arConnection)
    
        If cbool(This.arConnection(iIndex).DatabaseID = iDatabaseID) AndAlso _
           cbool(This.arConnection(iIndex).ConnectionID = iConnectionID) AndAlso _
           This.arConnection(iIndex).ConnectionInUse = True Then

           This.arConnection(iIndex).ConnectionInUse = False
           This.arConnection(iIndex).ReleaseCount = This.arConnection(iIndex).ReleaseCount + 1
           nTimer = IIf(nTimer >= This.arConnection(iIndex).RequestTime,nTimer,nTimer + 86400)
           This.arConnection(iIndex).RequestTotalTime = This.arConnection(iIndex).RequestTotalTime _
                                                      + nTimer - This.arConnection(iIndex).RequestTime

           lReturn = True
           
           Exit For
             
        End If
    
    Next 
                      
    LeaveCriticalSection(@This.lpCriticalSection)
    
    Function = lReturn                                                    
                                                      
End Function
' =====================================================================================
' Request pool connection
' =====================================================================================
Private Function cCTConnectionPool.RequestPoolConnection (ByVal iDatabaseID as Long, _
                                                          ByRef iConnectionID as uLong) as BOOLEAN
                                                      
Dim lReturn       as BOOLEAN
Dim iIndex        as Long 

    lReturn = False
    iConnectionID = 0
    EnterCriticalSection(@This.lpCriticalSection)
    
    For iIndex = 0 To UBound(This.arConnection)
        If cbool(This.arConnection(iIndex).DatabaseID = iDatabaseID) AndAlso _
           This.arConnection(iIndex).ConnectionInUse = False AndAlso _
           This.arConnection(iIndex).SlotFree = False Then
        
           This.arConnection(iIndex).ConnectionInUse = True
           iConnectionID = This.arConnection(iIndex).ConnectionID  
           This.arConnection(iIndex).RequestTime = Timer

           lReturn = True
           
           Exit For
             
        End If
    
    Next 
                      
    LeaveCriticalSection(@This.lpCriticalSection)
    
    Function = lReturn                                                    
                                                      
End Function
' =====================================================================================
' Remove pool connection
' =====================================================================================
Private Function cCTConnectionPool.RemovePoolConnection (ByVal iDatabaseID as Long, _
                                                         ByVal iConnectionID as uLong) as BOOLEAN
                                                      
Dim lReturn       as BOOLEAN
Dim iIndex        as Long 

    lReturn = False
    EnterCriticalSection(@This.lpCriticalSection)
    
    For iIndex = 0 To UBound(This.arConnection)
    
        If cBool(This.arConnection(iIndex).DatabaseID = iDatabaseID) AndAlso _
           cBool(This.arConnection(iIndex).ConnectionID = iConnectionID) AndAlso _
           This.arConnection(iIndex).SlotFree = False AndAlso _
           This.arConnection(iIndex).ConnectionInUse = False Then

           This.arConnection(iIndex).DatabaseID = 0
           This.arConnection(iIndex).ConnectionID = 0        
           This.arConnection(iIndex).ConnectionInUse = False
           This.arConnection(iIndex).SlotFree = True

           lReturn = True
           
           Exit For
             
        End If
    
    Next 
                      
    LeaveCriticalSection(@This.lpCriticalSection)
    
    Function = lReturn                                                    
                                                      
End Function
' =====================================================================================
' Get pool connection details
' =====================================================================================
Private Sub cCTConnectionPool.PoolConnectionDetail (arPool() as POOL_CONNECTION)

Dim iIndex        as Long
Dim iConnections  as Long

    Erase arPool
    iConnections = 0

    EnterCriticalSection(@This.lpCriticalSection)
    
' See how many connections are available to return

    For iIndex = 0 To UBound(This.arConnection)

        If This.arConnection(iIndex).SlotFree = False Then
        
           iConnections = iConnections + 1
           
        End If    
        
    Next
    
    ReDim arPool(0 To iConnections - 1)
    iConnections = 0
    
' Return known pool connections

    For iIndex = 0 To UBound(This.arConnection)
    
        If This.arConnection(iIndex).SlotFree = False Then
        
           arPool(iConnections) = This.arConnection(iIndex)
           iConnections = iConnections + 1
           
        End If    
    
    Next    
    
    LeaveCriticalSection(@This.lpCriticalSection)

End Sub
' =====================================================================================
' Get pool connection statistics
' =====================================================================================
Private Sub cCTConnectionPool.PoolConnectionStatistics (ByRef iPoolSize as Long, _
                                                        ByRef iTotalCurrentConnections as Long, _
                                                        ByRef iTotalActiveConnections as Long, _
                                                        ByRef nAverageRequestMilliseconds as Double)

Dim iIndex        as Long
Dim nRequestTime  as Double
Dim iReleaseCount as Long 

    iTotalCurrentConnections = 0
    iTotalActiveConnections = 0
    nAverageRequestMilliseconds = 0
    nRequestTime = 0
    iReleaseCount = 0
     
    EnterCriticalSection(@This.lpCriticalSection)
    
    iPoolSize = UBound(This.arConnection)
    
    For iIndex = 0 To iPoolSize
    
        If This.arConnection(iIndex).SlotFree = False Then
        
           iTotalCurrentConnections = iTotalCurrentConnections + 1
           
           If This.arConnection(iIndex).ConnectionInUse = True Then
           
              iTotalActiveConnections = iTotalActiveConnections + 1
              
           Else
           
              nRequestTime = nRequestTime + This.arConnection(iIndex).RequestTotalTime
              iReleaseCount = iReleaseCount + This.arConnection(iIndex).ReleaseCount          
              
           End If
             
        End If
    
    Next
    
    LeaveCriticalSection(@This.lpCriticalSection)
    
    nAverageRequestMilliseconds = (nRequestTime * 1000) / iReleaseCount
    iPoolSize = iPoolSize + 1
    
End Sub