' ########################################################################################
' File: cCTSQLLiteCrypto.bi
' Contents: FreeBasic Windows SQLite Client Server crypto class support.
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
#Include Once "win/bcrypt.bi"
#IncLib "bcrypt"
#Include "win/wincrypt.bi"
#IncLib "Crypt32"

' ########################################################################################
' cCTSQLLiteCrypto Class
' ########################################################################################

Type cCTSQLLiteCrypto Extends Object

Dim hRandom          as BCRYPT_ALG_HANDLE
Dim hHash            as BCRYPT_ALG_HANDLE
Dim hCrypto          as BCRYPT_ALG_HANDLE
Dim lStartup         as BOOLEAN
Dim iHashCount       as Long
Dim iCryptoIVIndex1  as Long
Dim iCryptoIVIndex2  as Long
Dim iCryptoIVIndex3  as Long
Dim iCryptoIVIndex4  as Long

Dim as Long arPrimes(0 To 255) => {2147394889, 2147394913, 2147394917, 2147394947, 2147394959, 2147394961, 2147394973, 2147394979, _
2147399759, 2147399767, 2147399777, 2147399789, 2147399803, 2147399809, 2147399819, 2147399843, _
2147395297, 2147395303, 2147395309, 2147395331, 2147395343, 2147395373, 2147395379, 2147395417, _
2147395421, 2147395423, 2147395427, 2147395487, 2147395489, 2147395499, 2147395553, 2147395589, _
2147395609, 2147395631, 2147395633, 2147395651, 2147395661, 2147395669, 2147395697, 2147395703, _
2147395709, 2147395721, 2147395729, 2147395771, 2147395777, 2147395829, 2147395841, 2147395891, _
2147399087, 2147399113, 2147399123, 2147399141, 2147399147, 2147399153, 2147399161, 2147399167, _
2147396057, 2147396063, 2147396077, 2147396129, 2147396159, 2147396179, 2147396189, 2147396203, _
2147396213, 2147396227, 2147396243, 2147396267, 2147396281, 2147396309, 2147396323, 2147396341, _
2147396351, 2147396353, 2147396399, 2147396401, 2147396411, 2147396413, 2147396441, 2147396513, _
2147397883, 2147397943, 2147397947, 2147397953, 2147397997, 2147398009, 2147398021, 2147398079, _
2147396623, 2147396659, 2147396687, 2147396711, 2147396749, 2147396759, 2147396761, 2147396807, _
2147396819, 2147396827, 2147396857, 2147396869, 2147396887, 2147396893, 2147396897, 2147396903, _
2147396921, 2147396963, 2147396987, 2147396989, 2147397011, 2147397019, 2147397029, 2147397071, _
2147397283, 2147397289, 2147397353, 2147397359, 2147397361, 2147397383, 2147397409, 2147397433, _
2147397437, 2147397443, 2147397463, 2147397479, 2147397541, 2147397557, 2147397563, 2147397569, _
2147397587, 2147397589, 2147397643, 2147397677, 2147397731, 2147397743, 2147397751, 2147397787, _
2147397809, 2147397817, 2147397821, 2147397827, 2147397853, 2147397859, 2147397877, 2147397881, _
2147397097, 2147397137, 2147397193, 2147397199, 2147397209, 2147397257, 2147397269, 2147397281, _
2147398081, 2147398091, 2147398109, 2147398157, 2147398207, 2147398219, 2147398229, 2147398261, _
2147398277, 2147398283, 2147398313, 2147398321, 2147398361, 2147398387, 2147398501, 2147398507, _
2147398529, 2147398531, 2147398549, 2147398553, 2147398559, 2147398577, 2147398597, 2147398667, _
2147398679, 2147398681, 2147398769, 2147398783, 2147398819, 2147398849, 2147398919, 2147398949, _
2147398963, 2147398969, 2147398997, 2147399017, 2147399021, 2147399053, 2147399063, 2147399069, _
2147396521, 2147396533, 2147396557, 2147396561, 2147396569, 2147396579, 2147396609, 2147396621, _
2147399197, 2147399203, 2147399227, 2147399263, 2147399329, 2147399363, 2147399381, 2147399383, _
2147399407, 2147399431, 2147399447, 2147399461, 2147399491, 2147399509, 2147399519, 2147399521, _
2147399533, 2147399561, 2147399581, 2147399587, 2147399593, 2147399603, 2147399629, 2147399663, _
2147399671, 2147399689, 2147399699, 2147399701, 2147399711, 2147399719, 2147399731, 2147399753, _
2147395907, 2147395927, 2147395937, 2147395949, 2147395961, 2147395969, 2147396039, 2147396051, _
2147395039, 2147395043, 2147395123, 2147395147, 2147395193, 2147395241, 2147395259, 2147395291, _
2147399851, 2147399909, 2147399923, 2147399927, 2147399957, 2147399959, 2147399981, 2147399983}

    Private:

    Public:

    Declare Function DecryptText(ByRef sCipherText as String, _
                                 ByRef sPlainText as String, _
                                 ByVal iEncryptIndex1 as Long, _
                                 ByVal iEncryptIndex2 as Long) as BOOLEAN
    Declare Function DecryptOneBlock(ByVal hKey as BCRYPT_KEY_HANDLE, _
                                     ByRef sCipherText as String, _
                                     ByRef sPlainText as String, _
                                     ByRef sIV as String, _
                                     ByVal lFinal as BOOLEAN) as BOOLEAN
    Declare Function EncryptText(ByRef sPlainText as String, _
                                 ByRef sCipherText as String, _
                                 ByVal iEncryptIndex1 as Long, _
                                 ByVal iEncryptIndex2 as Long) as BOOLEAN
    Declare Function EncryptOneBlock(ByVal hKey as BCRYPT_KEY_HANDLE, _
                                     ByRef sPlainText as String, _
                                     ByRef sCipherText as String, _
                                     ByRef sIV as String, _
                                     ByVal lFinal as BOOLEAN) as BOOLEAN
    Declare Sub HashString(ByRef sAny as String, _
                           ByRef sHash as String)    
    Declare Sub RandomString(ByRef sRandom as String, _
                             ByVal nLength as Long)
    Declare Function RandomIndex() as Long
    Declare Function GetPrime(ByRef iIndex as Long) as Long
    Declare Function SessionKey(ByRef iIndex1 as Long, _
                                ByRef iIndex2 as Long) as String
    Declare Function SharedSessionKey(ByVal iIndex1 as Long, _
                                      ByVal iIndex2 as Long) as String
    Declare Sub Bin2Hex(ByRef sBinary as String, _
                        ByRef sHex as String)
    Declare Sub Hex2Bin(ByRef sHex as String, _
                        ByRef sBinary as String)
    Declare Sub RestoreDefaultIV()
    Declare Property StartupStatus() as BOOLEAN
    Declare Property HashCount(ByVal iCount as Long)
    Declare Property CryptoIV1(ByVal iIndex as Long)
    Declare Property CryptoIV2(ByVal iIndex as Long)
    Declare Property CryptoIV3(ByVal iIndex as Long)
    Declare Property CryptoIV4(ByVal iIndex as Long)

    Declare Constructor
    Declare Destructor

End Type

Constructor cCTSQLLiteCrypto()

    This.lStartup = False
    This.iHashCount = 2
    RestoreDefaultIV ()
    
    If BCryptOpenAlgorithmProvider(VarPtr(This.hRandom),BCRYPT_RNG_ALGORITHM,0,0) = S_OK Then
    
       If BCryptOpenAlgorithmProvider(VarPtr(This.hHash),BCRYPT_SHA256_ALGORITHM,0,0) = S_OK Then
       
          If BCryptOpenAlgorithmProvider(VarPtr(This.hCrypto),BCRYPT_AES_ALGORITHM,0,0) = S_OK Then
          
             If BCryptSetProperty(This.hCrypto,ByVal CAST(LPCWSTR,StrPtr(BCRYPT_CHAINING_MODE)),ByVal StrPtr(BCRYPT_CHAIN_MODE_CBC),Len(BCRYPT_CHAIN_MODE_CBC),0) = S_OK Then
 
                This.lStartup = True
                
             End If
             
          End If   
      
       End If 
    
    End If

End Constructor
Destructor cCTSQLLiteCrypto()

    BCryptCloseAlgorithmProvider(This.hRandom,0)
    BCryptCloseAlgorithmProvider(This.hHash,0)
    BCryptCloseAlgorithmProvider(This.hCrypto,0)    

End Destructor

' =====================================================================================
' Decrypt a single plain text string
' =====================================================================================
Private Function cCTSQLLiteCrypto.DecryptText (ByRef sCipherText as String, _
                                               ByRef sPlainText as String, _
                                               ByVal iEncryptIndex1 as Long, _
                                               ByVal iEncryptIndex2 as Long) as BOOLEAN

Dim lStatus           as BOOLEAN
Dim hKey              as BCRYPT_KEY_HANDLE
Dim sIV               as String
Dim sKey              as String
Dim sKeyHash          as String

    lStatus = False

' Initialize IV

    sIV = SharedSessionKey (This.iCryptoIVIndex1,This.iCryptoIVIndex2) _
        + SharedSessionKey (This.iCryptoIVIndex3,This.iCryptoIVIndex4)

' Initialize key

    sKey = SharedSessionKey (iEncryptIndex1,iEncryptIndex2)
    
' SHA-2 256 bit hash of key material

    HashString(sKey,sKeyHash)

    If Len(sKeyHash) = 0 Then

       Function = False
       Exit Function

    End If

' Generate the AES key schedule and Decrypt

    If BCryptGenerateSymmetricKey(This.hCrypto,ByVal VarPtr(hKey),ByVal 0,ByVal 0,StrPtr(sKeyHash),Len(sKeyHash),0) = S_OK Then

       lStatus = DecryptOneBlock(hKey,sCipherText,sPlainText,sIV,True)

    End If

' Release key handle

    BCryptDestroyKey(hKey)
    
' Just a precaution to keep values from lingering in memory

    Clear(StrPtr(sKey),Len(sKey),0)
    Clear(StrPtr(sIV),Len(sIV),0)

    Function = lStatus

End Function
' =====================================================================================
' Decrypt one block
' =====================================================================================
Private Function cCTSQLLiteCrypto.DecryptOneBlock (ByVal hKey as BCRYPT_KEY_HANDLE, _
                                                   ByRef sCipherText as String, _
                                                   ByRef sPlainText as String, _
                                                   ByRef sIV as String, _
                                                   ByVal lFinal as BOOLEAN) as BOOLEAN

' Decrypt one block

' lFinal = TRUE, then this is last block of the decryption run
' and padding is stripped as needed

Dim iPlainTextSize    as Long
Dim iResult           as ULong
Dim lStatus           as BOOLEAN

    lStatus = False
    sPlainText = ""

' Get the Clear text size before decryption

    If BCryptDeCrypt(hKey,StrPtr(sCipherText),Len(sCipherText),ByVal 0,StrPtr(sIV),Len(sIV), _
                     ByVal 0,ByVal 0,ByVal VarPtr(iPlainTextSize),IIf(lFinal=True,BCRYPT_BLOCK_PADDING,0)) = S_OK Then

' Allocate the output plain text buffer and decrypt

       sPlainText = Space(iPlainTextSize)
       
       lStatus = BCryptDeCrypt(hKey,StrPtr(sCipherText),Len(sCipherText),ByVal 0,StrPtr(sIV),Len(sIV), _
                               StrPtr(sPlainText),iPlainTextSize,ByVal VarPtr(iResult),IIf(lFinal=True,BCRYPT_BLOCK_PADDING,0))

    End If

' If final block, iResult will have the final size

    If lFinal = True Then

       sPlainText = Left(sPlainText,iResult)

    End If

    Function = lStatus

End Function
' =====================================================================================
' Encrypt a single plain text string
' =====================================================================================
Private Function cCTSQLLiteCrypto.EncryptText (ByRef sPlainText as String, _
                                               ByRef sCipherText as String, _
                                               ByVal iEncryptIndex1 as Long, _
                                               ByVal iEncryptIndex2 as Long) as BOOLEAN

' Encrypt a single plain text string

Dim lStatus           as BOOLEAN
Dim hKey              as BCRYPT_KEY_HANDLE
Dim sIV               as String
Dim sKey              as String
Dim sKeyHash          as String

    lStatus = False

' Initialize IV

    sIV = SharedSessionKey (This.iCryptoIVIndex1,This.iCryptoIVIndex2) _
        + SharedSessionKey (This.iCryptoIVIndex3,This.iCryptoIVIndex4)

' Initialize key

    sKey = SharedSessionKey (iEncryptIndex1,iEncryptIndex2)

' SHA-2 256 bit hash of key material

    HashString(sKey,sKeyHash)

    If Len(sKeyHash) = 0 Then

       Function = False
       Exit Function

    End If

' Generate the AES key schedule and Encrypt

    If BCryptGenerateSymmetricKey(This.hCrypto,ByVal VarPtr(hKey),ByVal 0,ByVal 0,StrPtr(sKeyHash),Len(sKeyHash),0) = S_OK Then

       lStatus = EncryptOneBlock(hKey,sPlainText,sCipherText,sIV,True)

    End If
    
' Release key handle

    BCryptDestroyKey(hKey)
    
' Just a precaution to keep values from lingering in memory

    Clear(StrPtr(sKey),Len(sKey),0)
    Clear(StrPtr(sIV),Len(sIV),0)

    Function = lStatus

End Function
' =====================================================================================
' Encrypt one block
' =====================================================================================
Private Function cCTSQLLiteCrypto.EncryptOneBlock (ByVal hKey as BCRYPT_KEY_HANDLE, _
                                                   ByRef sPlainText as String, _
                                                   ByRef sCipherText as String, _
                                                   ByRef sIV as String, _
                                                   ByVal lFinal as BOOLEAN) as BOOLEAN

' lFinal = TRUE, then this is last block of the encryption run
' and padding is then added as needed

Dim nCipherTextSize   as Long
Dim iResult           as uLong
Dim lStatus           as BOOLEAN

    lStatus = False
    sCipherText = ""

' Get the Cipher text size before encryption

    If BCryptEnCrypt(hKey,StrPtr(sPlainText),Len(sPlainText),ByVal 0,StrPtr(sIV),Len(sIV), _
                     ByVal 0,ByVal 0,ByVal VarPtr(nCipherTextSize),IIf(lFinal=True,BCRYPT_BLOCK_PADDING,0)) = S_OK Then

' Allocate the output cipher text buffer and encrypt

       sCipherText = Space(nCipherTextSize)
       
       If BCryptEnCrypt(hKey,StrPtr(sPlainText),Len(sPlainText),ByVal 0,StrPtr(sIV),Len(sIV), _
                        StrPtr(sCipherText),nCipherTextSize,ByVal VarPtr(iResult),IIf(lFinal=True,BCRYPT_BLOCK_PADDING,0)) = S_OK Then
                        
          lStatus = True
         
       End If              

    End If

    Function = lStatus

End Function
' =====================================================================================
' Hash a string nCount number of times
' =====================================================================================
Private Sub cCTSQLLiteCrypto.HashString (ByRef sAny as String, _
                                         ByRef sHash as String)

' Hash a string set number of times

Dim hHashHandle   as BCRYPT_HASH_HANDLE
Dim nHashSize     as Long
Dim iLoop         as Long

    sHash = ""

    If BCryptCreateHash(This.hHash,ByVal VarPtr(hHashHandle),0,0,0,0,0) <> S_OK Then

        Exit Sub

    End If

    For iLoop = 1 To This.iHashCount

        BCryptHashData (hHashHandle,ByVal StrPtr(sAny),Len(sAny),0)

    Next

    nHashSize = 32
    sHash = Space(nHashSize)

    If BCryptFinishHash (hHashHandle,ByVal StrPtr(sHash),nHashSize,0) <> S_OK Then

        sHash = ""

    End If

    BCryptDestroyHash (hHashHandle)

End Sub
' =====================================================================================
' Build a shared session key
' =====================================================================================
Private Function cCTSQLLiteCrypto.SharedSessionKey (ByVal iIndex1 as Long, _
                                                    ByVal iIndex2 as Long) as String

Dim iSessionKey as ULongInt
Dim sSessionKey as String
Dim pFrom       as Any Ptr
Dim pTo         as Any Ptr   

    iSessionKey = arPrimes(iIndex1) * arPrimes(iIndex2) * 5539252711
    sSessionKey = Space(Len(iSessionKey))
    pTo = StrPtr(sSessionKey)
    pFrom = VarPtr(iSessionKey)
    MoveMemory(pTo,pFrom,Len(iSessionKey))

    Function = sSessionKey                                                     

End Function
' =====================================================================================
' Build a session key
' =====================================================================================
Private Function cCTSQLLiteCrypto.SessionKey (ByRef iIndex1 as Long, _
                                              ByRef iIndex2 as Long) as String

Dim iSessionKey as ULongInt
Dim sSessionKey as String
Dim pFrom       as Any Ptr
Dim pTo         as Any Ptr   

    iSessionKey = GetPrime (iIndex1) * GetPrime (iIndex2) * 5539252711
    sSessionKey = Space(Len(iSessionKey))
    pTo = StrPtr(sSessionKey)
    pFrom = VarPtr(iSessionKey)
    MoveMemory(pTo,pFrom,Len(iSessionKey))

    Function = sSessionKey

End Function
' =====================================================================================
' Convert binary string to hex representation
' =====================================================================================
Private Sub cCTSQLLiteCrypto.Bin2Hex (ByRef sBinary as String, _
                                      ByRef sHex as String)

Dim iHexLength    as Long

    sHex = ""
    iHexLength = Len(sBinary) * 2

    If Len(iHexLength) > 0 Then

        iHexLength = iHexLength + 1
        sHex = Space(iHexLength)

        CryptBinaryToString (ByVal StrPtr(sBinary), _
                             Len(sBinary), _
                             CRYPT_STRING_HEXRAW + CRYPT_STRING_NOCRLF, _
                             ByVal StrPtr(sHex), _
                             ByVal VarPtr(iHexLength))

         sHex = Left(sHex,iHexLength)

    End If

End Sub
' =====================================================================================
' Convert hex string to binary
' =====================================================================================
Private Sub cCTSQLLiteCrypto.Hex2Bin (ByRef sHex as String, _
                                      ByRef sBinary as String)

Dim iBinaryLength     as Long

    sBinary = ""

    If Len(sHex) Mod 2 = 0 Then

        iBinaryLength = Len(sHex) / 2
        sBinary = Space(iBinaryLength)

        CryptStringToBinary(ByVal StrPtr(sHex), _
                            Len(sHex), _
                            CRYPT_STRING_HEXRAW, _
                            ByVal StrPtr(sBinary), _
                            ByVal VarPtr(iBinaryLength), _
                            0, _
                            0)


    End If

End Sub
' =====================================================================================
' Retrieve a 32 bit prime number
' =====================================================================================
Private Function cCTSQLLiteCrypto.GetPrime (ByRef iIndex as Long) as Long

   iIndex = RandomIndex ()
   Function = arPrimes(iIndex)
   
End Function
' =====================================================================================
' Generate a random index
' =====================================================================================
Private Function cCTSQLLiteCrypto.RandomIndex () as Long

Dim sRandom     as String
Dim pByte       as Byte Ptr

    RandomString(sRandom,1)
    pByte = StrPtr(sRandom)
    
    Function = Abs(*pByte)
    
End Function
' =====================================================================================
' Generate a string of random values
' =====================================================================================
Private Sub cCTSQLLiteCrypto.RandomString (ByRef sRandom as String, _
                                           ByVal nLength as Long)
                                           
' Return a string of random bytes

    sRandom = ""

    If nLength > 0 Then

       sRandom = Space(nLength)

' Get a random stream of bytes

       BCryptGenRandom(This.hRandom,ByVal StrPtr(sRandom),nLength, ByVal 0)
       
    End If

End Sub
' =====================================================================================
' Restore default crypto IV indices
' =====================================================================================
Private Sub cCTSQLLiteCrypto.RestoreDefaultIV ()

    This.iCryptoIVIndex1 = 44
    This.iCryptoIVIndex2 = 72
    This.iCryptoIVIndex3 = 188
    This.iCryptoIVIndex4 = 240

End Sub
' =====================================================================================
' Crypto Startup Status
' =====================================================================================
Private Property cCTSQLLiteCrypto.StartupStatus () as BOOLEAN

    StartupStatus = This.lStartup

End Property
' =====================================================================================
' Set Hash Count
' =====================================================================================
Private Property cCTSQLLiteCrypto.HashCount (ByVal iCount as Long)

    If iCount > 0 Then
    
       This.iHashCount = iCount
       
    End If

End Property
' =====================================================================================
' Set Crypto IV Index 1
' =====================================================================================
Private Property cCTSQLLiteCrypto.CryptoIV1 (ByVal iIndex as Long)

    If iIndex >= 0 AndAlso iIndex < 256 Then
    
       This.iCryptoIVIndex1 = iIndex
       
    End If

End Property
' =====================================================================================
' Set Crypto IV Index 2
' =====================================================================================
Private Property cCTSQLLiteCrypto.CryptoIV2 (ByVal iIndex as Long)

    If iIndex >= 0 AndAlso iIndex < 256 Then
    
       This.iCryptoIVIndex2 = iIndex
       
    End If

End Property
' =====================================================================================
' Set Crypto IV Index 3
' =====================================================================================
Private Property cCTSQLLiteCrypto.CryptoIV3 (ByVal iIndex as Long)

    If iIndex >= 0 AndAlso iIndex < 256 Then
    
       This.iCryptoIVIndex3 = iIndex
       
    End If

End Property
' =====================================================================================
' Set Crypto IV Index 4
' =====================================================================================
Private Property cCTSQLLiteCrypto.CryptoIV4 (ByVal iIndex as Long)

    If iIndex >= 0 AndAlso iIndex < 256 Then
    
       This.iCryptoIVIndex4 = iIndex
       
    End If

End Property