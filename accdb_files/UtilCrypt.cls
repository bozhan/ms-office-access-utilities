VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UtilCrypt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

#If VBA7 And Win64 Then
  Private Declare PtrSafe Function CryptCreateHash Lib "advapi32.dll" _
    (ByVal hProv As LongPtr, ByVal Algid As Long, ByVal hKey As LongPtr, _
    ByVal dwFlags As Long, phHash As LongPtr) As Long
  Private Declare PtrSafe Function CryptDeriveKey Lib "advapi32.dll" _
    (ByVal hProv As LongPtr, ByVal Algid As Long, ByVal hBaseData As LongPtr, _
    ByVal dwFlags As Long, phKey As LongPtr) As Long
  Private Declare PtrSafe Function CryptDestroyHash Lib "advapi32.dll" _
    (ByVal hHash As LongPtr) As Long
  Private Declare PtrSafe Function CryptDestroyKey Lib "advapi32.dll" _
    (ByVal hKey As LongPtr) As Long
  Private Declare PtrSafe Function CryptEncrypt Lib "advapi32.dll" _
    (ByVal hKey As LongPtr, ByVal hHash As LongPtr, ByVal Final As Long, ByVal dwFlags As Long, _
    ByVal pbData As String, pdwDataLen As LongPtr, ByVal dwBufLen As Long) As Long
  Private Declare PtrSafe Function CryptDecrypt Lib "advapi32.dll" _
    (ByVal hKey As LongPtr, ByVal hHash As LongPtr, ByVal Final As Long, _
    ByVal dwFlags As Long, ByVal pbData As String, pdwDataLen As LongPtr) As Long
  Private Declare PtrSafe Function CryptExportKey Lib "advapi32.dll" _
    (ByVal hKey As LongPtr, ByVal hExpKey As LongPtr, ByVal dwBlobType As Long, _
    ByVal dwFlags As Long, ByVal pbData As String, pdwDataLen As LongPtr) As Long
  Private Declare PtrSafe Function CryptGenKey Lib "advapi32.dll" _
    (ByVal hProv As LongPtr, ByVal Algid As Long, ByVal dwFlags As Long, phKey As LongPtr) As Long
  Private Declare PtrSafe Function CryptGetProvParam Lib "advapi32.dll" _
    (ByVal hProv As LongPtr, ByVal dwParam As Long, pbData As Any, _
    pdwDataLen As LongPtr, ByVal dwFlags As Long) As Long
  Private Declare PtrSafe Function CryptGetUserKey Lib "advapi32.dll" _
    (ByVal hProv As LongPtr, ByVal dwKeySpec As Long, phUserKey As LongPtr) As Long
  Private Declare PtrSafe Function CryptHashData Lib "advapi32.dll" _
    (ByVal hHash As LongPtr, ByVal pbData As String, _
    ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
  Private Declare PtrSafe Function CryptReleaseContext Lib "advapi32.dll" _
    (ByVal hProv As LongPtr, ByVal dwFlags As Long) As Long
  Private Declare PtrSafe Function CryptSignHash Lib "advapi32.dll" Alias "CryptSignHashA" _
    (ByVal hHash As LongPtr, ByVal dwKeySpec As Long, ByVal sDescription As String, _
    ByVal dwFlags As Long, ByVal pbSignature As String, pdwSigLen As LongPtr) As Long
  Private Declare PtrSafe Function CryptVerifySignature Lib "advapi32.dll" Alias "CryptVerifySignatureA" _
    (ByVal hHash As LongPtr, ByVal pbSignature As String, ByVal dwSigLen As Long, _
    ByVal hPubKey As LongPtr, ByVal sDescription As String, ByVal dwFlags As Long) As Long
  Private Declare PtrSafe Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" _
    (phProv As LongPtr, pszContainer As String, pszProvider As String, _
    ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
  Private Declare PtrSafe Function CryptImportKey Lib "advapi32.dll" _
    (ByVal hProv As LongPtr, ByRef pbData As Any, ByVal dwDataLen As Long, _
    ByVal hPubKey As LongPtr, ByVal dwFlags As Long, ByRef phKey As LongPtr) As Long
  Private Declare PtrSafe Function CryptSetKeyParam Lib "advapi32.dll" _
    (ByVal hKey As LongPtr, ByVal dwParam As Long, ByRef pbData As Any, _
    ByVal dwFlags As Long) As Long
  Private Declare PtrSafe Function CryptGetKeyParam Lib "advapi32.dll" _
    (ByVal hKey As LongPtr, ByVal dwParam As Long, ByRef pbData As Any, _
    ByRef pdwDataLen As LongPtr, ByVal dwFlags As Long) As Long
    
  'API error function
  Private Declare PtrSafe Function GetLastError Lib "kernel32" () As Long
  
  'CryptoApi references
  Private Declare PtrSafe Function CryptBinaryToString Lib "Crypt32.dll" Alias "CryptBinaryToStringA" _
  (ByRef pbBinary As Byte, ByVal cbBinary As Long, ByVal dwFlags As Long, ByVal pszString As String, _
  ByRef pcchString As LongPtr) As Long
  
  Private Declare PtrSafe Function CryptStringToBinary Lib "Crypt32.dll" Alias "CryptStringToBinaryA" _
    (ByVal pszString As String, ByVal cchString As Long, ByVal dwFlags As Long, ByRef pbBinary As Byte, _
    ByRef pcbBinary As LongPtr, ByRef pdwSkip As Long, ByRef pdwFlags As LongPtr) As Long
#Else
  Private Declare Function CryptDeriveKey Lib "advapi32.dll" _
    (ByVal hProv As Long, ByVal Algid As Long, ByVal hBaseData As Long, _
    ByVal dwFlags As Long, phKey As Long) As Long
  Private Declare Function CryptCreateHash Lib "advapi32.dll" _
    (ByVal hProv As Long, ByVal Algid As Long, ByVal hKey As Long, _
    ByVal dwFlags As Long, phHash As Long) As Long
  Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" _
    (phProv As Long, pszContainer As String, pszProvider As String, _
    ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
  Private Declare Function CryptDestroyHash Lib "advapi32.dll" _
    (ByVal hHash As Long) As Long
  Private Declare Function CryptDestroyKey Lib "advapi32.dll" _
    (ByVal hKey As Long) As Long
  Private Declare Function CryptEncrypt Lib "advapi32.dll" _
    (ByVal hKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, _
    ByVal pbData As String, pdwDataLen As Long, ByVal dwBufLen As Long) As Long
  Private Declare Function CryptDecrypt Lib "advapi32.dll" _
    (ByVal hKey As Long, ByVal hHash As Long, ByVal Final As Long, _
    ByVal dwFlags As Long, ByVal pbData As String, pdwDataLen As Long) As Long
  Private Declare Function CryptExportKey Lib "advapi32.dll" _
    (ByVal hKey As Long, ByVal hExpKey As Long, ByVal dwBlobType As Long, _
    ByVal dwFlags As Long, ByVal pbData As String, pdwDataLen As Long) As Long
  Private Declare Function CryptGenKey Lib "advapi32.dll" _
    (ByVal hProv As Long, ByVal Algid As Long, ByVal dwFlags As Long, phKey As Long) As Long
  Private Declare Function CryptGetProvParam Lib "advapi32.dll" _
    (ByVal hProv As Long, ByVal dwParam As Long, pbData As Any, _
    pdwDataLen As Long, ByVal dwFlags As Long) As Long
  Private Declare Function CryptGetUserKey Lib "advapi32.dll" _
    (ByVal hProv As Long, ByVal dwKeySpec As Long, phUserKey As Long) As Long
  Private Declare Function CryptHashData Lib "advapi32.dll" _
    (ByVal hHash As Long, ByVal pbData As String, _
    ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
  Private Declare Function CryptReleaseContext Lib "advapi32.dll" _
    (ByVal hProv As Long, ByVal dwFlags As Long) As Long
  Private Declare Function CryptSignHash Lib "advapi32.dll" Alias "CryptSignHashA" _
    (ByVal hHash As Long, ByVal dwKeySpec As Long, ByVal sDescription As String, _
    ByVal dwFlags As Long, ByVal pbSignature As String, pdwSigLen As Long) As Long
  Private Declare Function CryptVerifySignature Lib "advapi32.dll" Alias "CryptVerifySignatureA" _
    (ByVal hHash As Long, ByVal pbSignature As String, ByVal dwSigLen As Long, _
    ByVal hPubKey As Long, ByVal sDescription As String, ByVal dwFlags As Long) As Long
  Private Declare Function CryptImportKey Lib "advapi32.dll" _
    (ByVal hProv As Long, ByRef pbData As Any, ByVal dwDataLen As Long, _
    ByVal hPubKey As Long, ByVal dwFlags As Long, ByRef phKey As Long) As Long
  Private Declare Function CryptSetKeyParam Lib "advapi32.dll" _
    (ByVal hKey As Long, ByVal dwParam As Long, ByRef pbData As Any, _
    ByVal dwFlags As Long) As Long
  Private Declare Function CryptGetKeyParam Lib "advapi32.dll" _
    (ByVal hKey As Long, ByVal dwParam As Long, ByRef pbData As Any, _
    ByRef pdwDataLen As Long, ByVal dwFlags As Long) As Long
    
  'CryptoApi references
  Private Declare Function CryptBinaryToString Lib "Crypt32.dll" Alias "CryptBinaryToStringA" _
  (ByRef pbBinary As Byte, ByVal cbBinary As Long, ByVal dwFlags As Long, ByVal pszString As String, _
  ByRef pcchString As Long) As Long
  
  Private Declare Function CryptStringToBinary Lib "Crypt32.dll" Alias "CryptStringToBinaryA" _
    (ByVal pszString As String, ByVal cchString As Long, ByVal dwFlags As Long, ByRef pbBinary As Byte, _
    ByRef pcbBinary As Long, ByRef pdwSkip As Long, ByRef pdwFlags As Long) As Long
    
  'API error function
  Private Declare Function GetLastError Lib "kernel32" () As Long
#End If




Private Const SERVICE_PROVIDER As String = "Microsoft Base Cryptographic Provider v1.0"
Private Const KEY_CONTAINER As String = "Home"
Private Const PROV_RSA_FULL As Long = 1
Private Const PP_NAME As Long = 4
Private Const PP_CONTAINER As Long = 6
Private Const CRYPT_NEWKEYSET As Long = 8
Private Const ALG_CLASS_DATA_ENCRYPT As Long = 24576
Private Const ALG_CLASS_HASH As Long = 32768
Private Const ALG_TYPE_ANY As Long = 0
Private Const ALG_TYPE_STREAM As Long = 2048
Private Const ALG_SID_RC4 As Long = 1
Private Const ALG_SID_MD5 As Long = 3
Private Const ALG_SID_MD2     As Long = 1
Private Const ALG_SID_MD4     As Long = 2
Private Const ALG_SID_SHA     As Long = 4
Private Const ALG_SID_SHA_256 As Long = 12
Private Const ALG_SID_SHA_384 As Long = 13
Private Const ALG_SID_SHA_512 As Long = 14
Private Const ALG_SID_RC5     As Long = 13
'constants from WinErr.h
Private Const NTE_NO_KEY As Long = -2146893811 '0x8009000DL
'constants
Private Const CFB_BUSY = 0
Private Const CFB_READY = 1
Private Const CFB_VALID = 2
Private Const ALG_TYPE_BLOCK         As Long = &H600&

Private Const ALG_TYPE_RSA           As Long = &H400&
Private Const ALG_SID_RSA_ANY        As Long = 0

Private Const ALG_SID_DES            As Long = 1
Private Const ALG_SID_RC2            As Long = 2
Private Const ALG_SID_3DES           As Long = 3
Private Const ALG_SID_3DES_112       As Long = 9
Private Const ALG_SID_AES_128        As Long = 14
Private Const ALG_SID_AES_192        As Long = 15
Private Const ALG_SID_AES_256        As Long = 16
'Default Encryption and signature parameters used in case non are provided in InitUser

'Private Const CALG_RC4 = (ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_STREAM Or ALG_SID_RC4)
'Private Const ENCRYPT_ALGORITHM As Long = CALG_RC4
Private Const NUMBER_ENCRYPT_PASSWORD As String = "´o¸sçPQ]"

Private Const KP_ALGID As Long = 7
Private Const KP_P As Long = 11
Private Const KP_G As Long = 12
Private Const KP_X As Long = 14
Private Const CRYPT_PREGEN As Long = &H40
Private Const AT_SIGNATURE As Long = 2
Private Const ALG_CLASS_KEY_EXCHANGE As Long = &HA000&

Private Const ALG_CLASS_SIGNATURE    As Long = &H2000&
Private Const ALG_TYPE_DH As Long = 2560
' Diffie-Hellman
Private Const ALG_SID_DH_SANDF As Long = 1
Private Const PROV_RSA_AES    As Long = 24

Private Const CALG_DH_SF = (ALG_CLASS_KEY_EXCHANGE Or ALG_TYPE_DH Or ALG_SID_DH_SANDF)
Private Const MS_ENH_DSS_DH_PROV As String = "Microsoft Enhanced DSS and Diffie-Hellman Cryptographic Provider"
Private Const MS_ENH_RSA_AES_PROV As String = "Microsoft Enhanced RSA and AES Cryptographic Provider"

Private Const DEFAULT_PROVIDER_NAME = MS_ENH_RSA_AES_PROV
Private Const DEFAULT_PROVIDER = PROV_RSA_AES

'Hash Algorithms:
Public Enum EnmHashAlgo
    CALG_MD2 = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD2)
    CALG_MD4 = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD4)
    CALG_MD5 = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD5)
    CALG_SHA = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA)
    CALG_SHA256 = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA_256)
    CALG_SHA384 = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA_384)
    CALG_SHA512 = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA_512)
End Enum

'symmetrische Verschlüsselungs Algorithmen:
Public Enum EnmCryptAlgo
    CALG_RC2 = (ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_RC2)
    CALG_RC4 = (ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_STREAM Or ALG_SID_RC4)
    CALG_DES = (ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_DES)
    CALG_3DES = (ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_3DES)
    CALG_3DES_112 = (ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_3DES_112)
    CALG_AES_128 = (ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_AES_128)
    CALG_AES_192 = (ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_AES_192)
    CALG_AES_256 = (ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_AES_256)
End Enum

#If VBA7 And Win64 Then
    Private hCryptProv As LongPtr
#Else
    Private hCryptProv As Long
#End If

'Rsa Keyexchange and Signature Algorithmen (asymmetrisch):
Public Enum EnmRsaKeyTypes
    CALG_RSA_SIGN = (ALG_CLASS_SIGNATURE Or ALG_TYPE_RSA Or ALG_SID_RSA_ANY)
    CALG_RSA_KEYX = (ALG_CLASS_KEY_EXCHANGE Or ALG_TYPE_RSA Or ALG_SID_RSA_ANY)
End Enum

Public Enum EnmRsaKeyParts
    PRIVATEKEYBLOB = &H7
    PUBLICKEYBLOB = &H6
End Enum


'merge with
'================================================================================================================
'private property buffers
Public sInBuffer As String
Public sOutBuffer As String
Public sPassword As String
Public sSignature As String
Private lStatus As Long

Private Type CRYPT_PARAMETERS
  ProviderName As String
  ProviderType As Long
  HashAlgorithm As Long
  EncryptAlgorithm As Long
End Type

Private params As CRYPT_PARAMETERS

'constants for Cryptography API functions



Public Enum CryptHashCalgIds
  MD2_id = ((ALG_CLASS_HASH Or ALG_TYPE_ANY) Or ALG_SID_MD2)
  MD4_id = ((ALG_CLASS_HASH Or ALG_TYPE_ANY) Or ALG_SID_MD4)
  MD5_id = ((ALG_CLASS_HASH Or ALG_TYPE_ANY) Or ALG_SID_MD5)
  SHA_id = ((ALG_CLASS_HASH Or ALG_TYPE_ANY) Or ALG_SID_SHA)
  SHA_256_id = ((ALG_CLASS_HASH Or ALG_TYPE_ANY) Or ALG_SID_SHA_256)
  SHA_384_id = ((ALG_CLASS_HASH Or ALG_TYPE_ANY) Or ALG_SID_SHA_384)
  SHA_512_id = ((ALG_CLASS_HASH Or ALG_TYPE_ANY) Or ALG_SID_SHA_512)
End Enum

Public Enum CryptEncryptCalgIds
  RC2_id = ((ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK) Or ALG_SID_RC2)
  RC4_id = ((ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_STREAM) Or ALG_SID_RC4)
  RC5_id = ((ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK) Or ALG_SID_RC5)
  AES_256_id = ((ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK) Or ALG_SID_AES_256)
End Enum


'================================================================================================================

Public Function EncryptionCSPConnect() As Boolean
    'Function Adapted
    'Get handle to CSP
    
    If CryptAcquireContext(hCryptProv, KEY_CONTAINER, SERVICE_PROVIDER, PROV_RSA_FULL, CRYPT_NEWKEYSET) = 0 Then
        If CryptAcquireContext(hCryptProv, KEY_CONTAINER, SERVICE_PROVIDER, PROV_RSA_FULL, 0) = 0 Then
            Debug.Print "Error during CryptAcquireContext for a new key container." & vbCrLf & _
                        "A container with this name probably already exists."
            EncryptionCSPConnect = False
            Exit Function
        End If
    End If
    
    EncryptionCSPConnect = True
End Function

Public Sub EncryptionCSPDisconnect()
    'Release provider handle.
   'Function Adapted

    If hCryptProv <> 0 Then
        CryptReleaseContext hCryptProv, 0
    End If
End Sub

Public Function EncryptData(ByVal data As String, ByVal Password As String) As String
    Dim sEncrypted As String
    Dim lEncryptionCount As Long
    Dim sTempPassword As String

    'It is possible that the normal encryption will give you a string
    'containing cr or lf characters which make it difficult to write to files
    'Do a loop changing the password and keep encrypting until the result is ok
    'To be able to decrypt we need to also store the number of loops in the result

    'Try first encryption
    lEncryptionCount = 0
    sTempPassword = Password & lEncryptionCount
    sEncrypted = EncryptDecrypt(data, sTempPassword, True)

    'Loop if this contained a bad character
    Do While (InStr(1, sEncrypted, vbCr) > 0) _
          Or (InStr(1, sEncrypted, vbLf) > 0) _
          Or (InStr(1, sEncrypted, Chr$(0)) > 0) _
          Or (InStr(1, sEncrypted, vbTab) > 0)

        'Try the next password
        lEncryptionCount = lEncryptionCount + 1
        sTempPassword = Password & lEncryptionCount
        sEncrypted = EncryptDecrypt(data, sTempPassword, True)

        'Don't go on for ever, 1 billion attempts should be plenty
        If lEncryptionCount = 99999999 Then
            err.raise vbObjectError + 999, "EncryptData", "This data cannot be successfully encrypted"
            EncryptData = ""
            Exit Function
        End If
    Loop

    'Build encrypted string, starting with number of encryption iterations
    EncryptData = EncryptNumber(lEncryptionCount) & sEncrypted
End Function

Public Function DecryptData(ByVal data As String, ByVal Password As String) As String
    Dim lEncryptionCount As Long
    Dim sDecrypted As String
    Dim sTempPassword As String

    'When encrypting we may have gone through a number of iterations
    'How many did we go through?
    lEncryptionCount = DecryptNumber(mid$(data, 1, 8))

    'start with the last password and work back
    sTempPassword = Password & lEncryptionCount
    sDecrypted = EncryptDecrypt(mid$(data, 9), sTempPassword, False)

    DecryptData = sDecrypted
End Function

Public Function GetCSPDetails() As String
    Dim lLength As Long
    Dim yContainer() As Byte
    
    If hCryptProv = 0 Then
        GetCSPDetails = "Not connected to CSP"
        Exit Function
    End If
    
    'For developer info, show what the CSP & container name is
    lLength = 1000
    ReDim yContainer(lLength)
    If CryptGetProvParam(hCryptProv, PP_NAME, yContainer(0), lLength, 0) <> 0 Then
        GetCSPDetails = "Cryptographic Service Provider name: " & ByteToStr(yContainer, lLength)
    End If
    
    lLength = 1000
    ReDim yContainer(lLength)
    If CryptGetProvParam(hCryptProv, PP_CONTAINER, yContainer(0), lLength, 0) <> 0 Then
        GetCSPDetails = GetCSPDetails & vbCrLf & "Key Container name: " & ByteToStr(yContainer, lLength)
        'ByteToStr(yContainer, lLength)
    End If
End Function

Private Function EncryptDecrypt(ByVal data As String, ByVal Password As String, ByVal encrypt As Boolean) As String
    
    #If Win64 Then
       Dim hHash As LongPtr
       Dim hKey  As LongPtr
       Dim hHashNull As LongPtr
       Dim hKeyNull As LongPtr
       
       hHashNull = 0&
       hKeyNull = 0&
    
    #Else
       Dim hHash As Long
        Dim hKey As Long
        Dim hHashNull As Long
       Dim hKeyNull As Long
    #End If
    
    
    Dim lLength As Long
    Dim sTemp As String
    Dim GetValue As Boolean
    
    If hCryptProv = 0 Then
        Debug.Print "Not connected to CSP"
        Exit Function
    End If
    
    '--------------------------------------------------------------------
    'The data will be encrypted with a session key derived from the
    'password.
    'The session key will be recreated when the data is decrypted
    'only if the password used to create the key is available.
    '--------------------------------------------------------------------

    'Create a hash object.
    GetValue = CryptCreateHash(hCryptProv, CALG_MD5, hKeyNull, 0, hHash)
    
    If GetValue = False Then
        Debug.Print "Error during CryptCreateHash!"
    End If

    'Hash the password.
    If CryptHashData(hHash, Password, Len(Password), 0) = 0 Then
        Debug.Print "Error during CryptHashData."
    End If
    
    'Derive a session key from the hash object.
    If CryptDeriveKey(hCryptProv, EnmCryptAlgo.CALG_RC4, hHash, 0, hKey) = 0 Then
        Debug.Print "Error during CryptDeriveKey!"
    End If
    
    'Do the work
    sTemp = data
    lLength = Len(data)
    If encrypt Then
        'Encrypt data.
        If CryptEncrypt(hKey, hHashNull, True, 0, sTemp, lLength, lLength) = 0 Then
            Debug.Print "Error during CryptEncrypt."
        End If
    Else
        'Encrypt data.
        GetValue = CryptDecrypt(hKey, hHashNull, True, 0, sTemp, lLength)
        If GetValue = 0 Then
            Debug.Print "Error during CryptDecrypt."
        End If
        
        'use this method to failproof the c++ functions
'      'Decrypt data
'      If Not CBool(CryptDecrypt(lHCryptKey, 0, 1, 0, sCryptBuffer, lCryptBufLen)) Then
'        err.raise ErrCode.CannotPerformRequestedOperation, _
'          Description:="Error during CryptDecrypt bytes required:" & CStr(sTemp) & ": " & CStr(GetLastError)
'      End If
    End If

    'This is what we return.
    EncryptDecrypt = mid$(sTemp, 1, lLength)
    
    'Destroy session key.
    If hKey <> 0 Then
        CryptDestroyKey hKey
    End If

    'Destroy hash object.
    If hHash <> 0 Then
        CryptDestroyHash hHash
    End If
End Function

Private Function ByteToStr(ByRef ByteArray() As Byte, ByVal lLength As Long) As String
    Dim i As Long
    For i = LBound(ByteArray) To (LBound(ByteArray) + lLength)
        ByteToStr = ByteToStr & Chr$(ByteArray(i))
    Next i
End Function

Private Function EncryptNumber(ByVal lNumber As Long) As String
    Dim i As Long
    Dim sNumber As String

    sNumber = Format$(lNumber, "00000000")

    For i = 1 To Len(CStr(sNumber))
        EncryptNumber = EncryptNumber & Chr$(Asc(mid$(NUMBER_ENCRYPT_PASSWORD, i, 1)) + val(mid$(sNumber, i, 1)))
    Next i
End Function

Private Function DecryptNumber(ByVal sNumber As String) As Long
    Dim i As Long

    For i = 1 To Len(CStr(sNumber))
        DecryptNumber = (10 * DecryptNumber) + (Asc(mid$(sNumber, i, 1)) - Asc(mid$(NUMBER_ENCRYPT_PASSWORD, i, 1)))
    Next i
End Function


'Merge With
'==================================================================================================================


'---------------------------------------------------------------------------------------
' Procedure : InitUser
' Purpose   :
'---------------------------------------------------------------------------------------
Public Function InitUser( _
  Optional ByVal ProviderName As String = DEFAULT_PROVIDER_NAME, _
  Optional ByVal ProviderType As Long = DEFAULT_PROVIDER, _
  Optional ByVal hashAlg As Long = EnmHashAlgo.CALG_SHA512, _
  Optional ByVal CryptAlg As Long = EnmCryptAlgo.CALG_AES_256 _
) As Long
#If VBA7 And Win64 Then
  Dim lHCryptprov As LongPtr, lHHash As LongPtr, lHCryptKey As LongPtr, lProviderDataLen As LongPtr
#Else
  Dim lHCryptprov As Long, lHHash As Long, lHCryptKey As Long, lProviderDataLen As Long
#End If
Dim lResult As Long, sContainer As String, sUserName As String, lPoint As Long, avProviderData(1000) As Byte
'Dim lProviderDataAddress As Long, lDataSize As Long, sBuffer As String, lMemHandle As Long, lReturn As Long
On Error GoTo InitUser_Error
  
  With params 'init params
    .ProviderName = ProviderName & vbNullChar
    .ProviderType = ProviderType
    .HashAlgorithm = hashAlg
    .EncryptAlgorithm = CryptAlg
  End With
  
  'prepare string buffers
  sContainer = vbNullChar
  
  'Attempt to acquire a handle to the default key container.
  If Not CBool(CryptAcquireContext(lHCryptprov, ByVal sContainer, ByVal params.ProviderName, params.ProviderType, 0)) Then
    'Create default key container.
    If Not CBool(CryptAcquireContext(lHCryptprov, ByVal sContainer, ByVal params.ProviderName, params.ProviderType, CRYPT_NEWKEYSET)) Then
      err.raise ErrCode.CannotPerformRequestedOperation, _
        Description:="Error creating key container! " & CStr(GetLastError)
    End If
    
    'Get name of default key container.
    lProviderDataLen = Len(avProviderData(0)) * (UBound(avProviderData) + 1)
    If Not CBool(CryptGetProvParam(lHCryptprov, PP_CONTAINER, avProviderData(0), lProviderDataLen, 0)) Then
      avProviderData(0) = 0
      err.raise ErrCode.CannotPerformRequestedOperation, _
        Description:="Error getting user name! " & CStr(GetLastError)
    End If
    
    'Get sUserName from avProviderData()
    lPoint = LBound(avProviderData)
    While lPoint <= UBound(avProviderData)
      If avProviderData(lPoint) <> 0 Then
        sUserName = sUserName & Chr$(avProviderData(lPoint))
      Else
        lPoint = UBound(avProviderData)
      End If
      lPoint = lPoint + 1
    Wend
  End If
  
  'Attempt to get handle to signature key
  If Not CBool(CryptGetUserKey(lHCryptprov, AT_SIGNATURE, lHCryptKey)) Then
    If GetLastError = NTE_NO_KEY Then
      If Not CBool(CryptGenKey(lHCryptprov, AT_SIGNATURE, 0, lHCryptKey)) Then
        err.raise ErrCode.CannotPerformRequestedOperation, _
          Description:="Error during CryptGenKey! " & CStr(GetLastError)
      Else
        lResult = CryptDestroyKey(lHCryptprov)
      End If
    Else
      err.raise ErrCode.CannotPerformRequestedOperation, _
        Description:="Error during CryptGenKey! " & CStr(GetLastError)
    End If
  End If
  
  Dim v As Variant
  v = CryptGetUserKey(lHCryptprov, AT_KEYEXCHANGE, lHCryptKey)
  'Attempt to get handle to exchange key
  If Not CBool(CryptGetUserKey(lHCryptprov, AT_KEYEXCHANGE, lHCryptKey)) Then
    If GetLastError = NTE_NO_KEY Then
      If Not CBool(CryptGenKey(lHCryptprov, AT_KEYEXCHANGE, 0, lHCryptKey)) Then
        err.raise ErrCode.CannotPerformRequestedOperation, _
          Description:="Error during CryptGenKey! " & CStr(GetLastError)
      Else
        lResult = CryptDestroyKey(lHCryptprov)
      End If
    Else
      err.raise ErrCode.CannotPerformRequestedOperation, _
         Description:="Error during CryptGetUserKey! " & CStr(GetLastError)
    End If
  End If
  
  'release handle to provider
  lResult = CryptReleaseContext(lHCryptprov, 0)
  InitUser = True
  
InitUser_Exit:
On Error Resume Next
Exit Function

InitUser_Error:
  InitUser = False
  err.raise Number:=err.Number, _
    Description:="Crypt Class cloud not execute InitUser!" & vbLf & _
    "Please check your CryptAPI installation and/or parameters in InitUser" & vbLf & _
    "Details: " & err.Description & vbLf & InfoErrMsg()
Resume InitUser_Exit
End Function

'---------------------------------------------------------------------------------------
' Procedure : Sign
' Purpose   : Creates a signature for Inbuffer and places in Signature property
'---------------------------------------------------------------------------------------
Public Sub Sign()
#If VBA7 And Win64 Then
  Dim lHCryptprov As LongPtr, lHHash As LongPtr, lSignatureLen As LongPtr
#Else
  Dim lHCryptprov As Long, lHHash As Long, lSignatureLen As Long
#End If
Dim sContainer As String, sDescription As String, lResult As Long
On Error GoTo Sign_Error

  lStatus = CFB_BUSY 'switch Status property
  sSignature = "" 'init Signature property
  sContainer = vbNullChar 'Get handle to the default provider.

  If Not CBool(CryptAcquireContext(lHCryptprov, ByVal sContainer, ByVal params.ProviderName, params.ProviderType, 0)) Then
    err.raise ErrCode.CannotPerformRequestedOperation, _
      Description:="Error during CryptAcquireContext! " & CStr(GetLastError)
  End If
  
  'Create a hash object.
  If Not CBool(CryptCreateHash(lHCryptprov, params.HashAlgorithm, 0, 0, lHHash)) Then
    err.raise ErrCode.CannotPerformRequestedOperation, _
      Description:="Error during CryptCreateHash! " & CStr(GetLastError)
  End If
  
  If Not CBool(CryptHashData(lHHash, sInBuffer, Len(sInBuffer), 0)) Then
    err.raise ErrCode.CannotPerformRequestedOperation, _
      Description:="Error during CryptHashData! " & CStr(GetLastError)
  End If
  
  'Sign hash object. Determine size of signature.
  sDescription = vbNullChar
  lResult = CryptSignHash(lHHash, AT_SIGNATURE, sDescription, 0, sSignature, lSignatureLen)
  sSignature = String(lSignatureLen, vbNullChar)
  
  'Sign hash object (with signature key).
  If Not CBool(CryptSignHash(lHHash, AT_SIGNATURE, sDescription, 0, sSignature, lSignatureLen)) Then
    err.raise ErrCode.CannotPerformRequestedOperation, _
      Description:="Error during CryptSignHash! " & CStr(GetLastError)
  End If

Sign_Exit:
On Error Resume Next
  If lHHash Then lResult = CryptDestroyHash(lHHash)
  If lHCryptprov Then lResult = CryptReleaseContext(lHCryptprov, 0)
  lStatus = CFB_READY 'switch Status property
Exit Sub

Sign_Error:
  err.raise Number:=err.Number, _
    Description:="Error in in procedure Sign of UtilSecurity:Class Module" & vbLf & _
    err.Description & vbLf & InfoErrMsg()
Resume Sign_Exit
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Validate
' Purpose   : Validates InBuffer with Signature and assigns Status with result
'---------------------------------------------------------------------------------------
Public Sub Validate()
#If VBA7 And Win64 Then
  Dim lHCryptprov As LongPtr, lHHash As LongPtr, lSignatureLen As LongPtr, lHCryptKey As LongPtr
#Else
  Dim lHCryptprov As Long, lHHash As Long, lSignatureLen As Long, lHCryptKey As Long
#End If
Dim lResult As Long, bValid As Boolean, sContainer As String, sDescription As String, sProvider As String
ReDim aByteData(0) As Byte
On Error GoTo Validate_Error

  lStatus = CFB_BUSY   'switch Status property
  bValid = True 'init internal valid flag
  
  'Get handle to the default provider.
  sContainer = vbNullChar
  If Not CBool(CryptAcquireContext(lHCryptprov, ByVal sContainer, ByVal params.ProviderName, params.ProviderType, 0)) Then
    err.raise ErrCode.CannotPerformRequestedOperation, _
      Description:="Error during CryptAcquireContext! " & CStr(GetLastError)
  End If
  
  'Create a hash object.
  If Not CBool(CryptCreateHash(lHCryptprov, params.HashAlgorithm, 0, 0, lHHash)) Then
    err.raise ErrCode.CannotPerformRequestedOperation, _
      Description:="Error during CryptCreateHash! " & CStr(GetLastError)
  End If
  
  'Add data to hash object.
  If Not CBool(CryptHashData(lHHash, sInBuffer, Len(sInBuffer), 0)) Then
    err.raise ErrCode.CannotPerformRequestedOperation, _
      Description:="Error during CryptHashData! " & CStr(GetLastError)
  End If
  
  'Determine size of signature.
  'sDescription = vbNullChar
  'lResult = CryptSignHash(lHHash, AT_SIGNATURE, sDescription, 0, 0, lSignatureLen)
  'Get handle to signature key.
  If Not CBool(CryptGetUserKey(lHCryptprov, AT_SIGNATURE, lHCryptKey)) Then
    err.raise ErrCode.CannotPerformRequestedOperation, _
      Description:="Error during CryptGetUserKey! " & CStr(GetLastError)
  End If
  
  lSignatureLen = Len(sSignature)
  
  'Verify signature.
  If Not CBool(CryptVerifySignature(lHHash, sSignature, lSignatureLen, lHCryptKey, sDescription, 0)) Then
    If GetLastError = NTE_BAD_SIGNATURE Then
      err.raise ErrCode.CannotPerformRequestedOperation, _
        Description:="Error during CryptCreateHash verifying NTE_BAD_SIGNATURE! " & CStr(NTE_BAD_SIGNATURE)
    Else
      err.raise ErrCode.CannotPerformRequestedOperation, _
        Description:="Error during CryptVerifySignature! " & CStr(GetLastError)
    End If
  End If

Validate_Exit:
On Error Resume Next
  If lHCryptKey Then lResult = CryptDestroyKey(lHCryptKey)
  If lHHash Then lResult = CryptDestroyHash(lHHash)
  If lHCryptprov Then lResult = CryptReleaseContext(lHCryptprov, 0)
  
  If bValid Then
    lStatus = CFB_VALID
  Else
    lStatus = CFB_READY
  End If
Exit Sub

Validate_Error:
  bValid = False
  err.raise Number:=err.Number, _
    Description:="Error in procedure Validate of UtilSecurity:Class Module" & vbLf & _
    err.Description & vbLf & InfoErrMsg()
Resume Validate_Exit
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Encrypt
' Purpose   : Encrypt InBuffer into OutBuffer
'---------------------------------------------------------------------------------------
Public Sub encrypt()
#If VBA7 And Win64 Then
  Dim lHCryptprov As LongPtr, lHHash As LongPtr, lSignatureLen As LongPtr
  Dim lHCryptKey As LongPtr, lHExchgKey As LongPtr, lCryptLength As LongPtr
#Else
  Dim lHCryptprov As Long, lHHash As Long, lSignatureLen As Long
  Dim lHCryptKey As Long, lHExchgKey As Long, lCryptLength As Long
#End If
Dim lResult As Long, lCryptBufLen As Long, sContainer As String, sProvider As String, sCryptBuffer As String
On Error GoTo Encrypt_Error

  lStatus = CFB_BUSY 'Get handle to the default provider
  sContainer = vbNullChar
    
  If Not CBool(CryptAcquireContext(lHCryptprov, ByVal sContainer, ByVal params.ProviderName, params.ProviderType, 0)) Then
    err.raise ErrCode.CannotPerformRequestedOperation, _
      Description:="Error during CryptAcquireContext! " & CStr(GetLastError)
  End If
 
  'Create a hash object.
  If Not CBool(CryptCreateHash(lHCryptprov, params.HashAlgorithm, 0, 0, lHHash)) Then
    err.raise ErrCode.CannotPerformRequestedOperation, _
      Description:="Error during CryptCreateHash! " & CStr(GetLastError)
  End If
 
  'Hash in the password data.
  If Not CBool(CryptHashData(lHHash, sPassword, Len(sPassword), 0)) Then
    err.raise ErrCode.CannotPerformRequestedOperation, _
      Description:="Error during CryptHashData! " & CStr(GetLastError)
  End If
 
  'Derive a session key from the hash object.
  If Not CBool(CryptDeriveKey(lHCryptprov, params.EncryptAlgorithm, lHHash, 0, lHCryptKey)) Then
    err.raise ErrCode.CannotPerformRequestedOperation, _
      Description:="Error during CryptDeriveKey! " & CStr(GetLastError)
  End If
  
  Debug.Print "key:", lHCryptKey
  'Destroy the hash object.
  CryptDestroyHash (lHHash)
  lHHash = 0
 
 'Prepare a string buffer for the CryptEncrypt function
  lCryptLength = Len(sInBuffer)
  lCryptBufLen = lCryptLength * 2
  sCryptBuffer = String(lCryptBufLen, vbNullChar)
  LSet sCryptBuffer = sInBuffer
  
  'Encrypt data
  If Not CBool(CryptEncrypt(lHCryptKey, 0, 1, 0, sCryptBuffer, lCryptLength, lCryptBufLen)) Then
    err.raise ErrCode.CannotPerformRequestedOperation, _
      Description:="Error during CryptEncrypt bytes required:" & CStr(lCryptLength) & ": " & CStr(GetLastError)
  End If
 
  sOutBuffer = mid$(sCryptBuffer, 1, lCryptLength)
  
Encrypt_Exit:
On Error Resume Next
  If (lHCryptKey) Then lResult = CryptDestroyKey(lHCryptKey)
  If lHExchgKey Then CryptDestroyKey (lHExchgKey)
  If lHHash Then CryptDestroyHash (lHHash)
  If lHCryptprov Then lResult = CryptReleaseContext(lHCryptprov, 0)
  lStatus = CFB_READY 'switch Status property
Exit Sub

Encrypt_Error:
  err.raise Number:=err.Number, _
    Description:="Error in procedure Encrypt of UtilSecurity:Class Modulee" & vbLf & _
    err.Description & vbLf & InfoErrMsg()
Resume Encrypt_Exit
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Decrypt
' Purpose   : Decrypt InBuffer into OutBuffer
'---------------------------------------------------------------------------------------
Public Sub Decrypt()
#If VBA7 And Win64 Then
  Dim lHCryptprov As LongPtr, lHHash As LongPtr, lSignatureLen As LongPtr
  Dim lHCryptKey As LongPtr, lHExchgKey As LongPtr, lCryptBufLen As LongPtr
#Else
  Dim lHCryptprov As Long, lHHash As Long, lSignatureLen As Long
  Dim lHCryptKey As Long, lHExchgKey As Long, lCryptBufLen As Long
#End If
Dim lResult As Long, sContainer As String, sCryptBuffer As String
On Error GoTo Decrypt_Error

  lStatus = CFB_BUSY 'switch Status property
  sOutBuffer = "" 'Init sOutBuffer
  sContainer = vbNullChar 'Get handle to the default provider.
  
  If Not CBool(CryptAcquireContext(lHCryptprov, ByVal sContainer, ByVal params.ProviderName, params.ProviderType, 0)) Then
    err.raise ErrCode.CannotPerformRequestedOperation, _
      Description:="Error during CryptAcquireContext! " & CStr(GetLastError)
  End If
  
  'Create a hash object.
  If Not CBool(CryptCreateHash(lHCryptprov, params.HashAlgorithm, 0, 0, lHHash)) Then
    err.raise ErrCode.CannotPerformRequestedOperation, _
      Description:="Error during CryptCreateHash! " & CStr(GetLastError)
  End If
  
  'Hash in the password data.
  If Not CBool(CryptHashData(lHHash, sPassword, Len(sPassword), 0)) Then
    err.raise ErrCode.CannotPerformRequestedOperation, _
      Description:="Error during CryptHashData! " & CStr(GetLastError)
  End If
  
  'Derive a session key from the hash object.
  If Not CBool(CryptDeriveKey(lHCryptprov, params.EncryptAlgorithm, lHHash, 0, lHCryptKey)) Then
    err.raise ErrCode.CannotPerformRequestedOperation, _
      Description:="Error during CryptDeriveKey! " & CStr(GetLastError)
  End If
  
  'Destroy the hash object.
  CryptDestroyHash (lHHash)
  lHHash = 0
  
  'Prepare sCryptBuffer for CryptDecrypt
  lCryptBufLen = Len(sInBuffer) '* 2
  sCryptBuffer = String(lCryptBufLen, vbNullChar)
  LSet sCryptBuffer = sInBuffer
  
  'Decrypt data
  If Not CBool(CryptDecrypt(lHCryptKey, 0, 1, 0, sCryptBuffer, lCryptBufLen)) Then
    err.raise ErrCode.CannotPerformRequestedOperation, _
      Description:="Error during CryptDecrypt bytes required:" & CStr(lCryptBufLen) & ": " & CStr(GetLastError)
  End If
  
  'Apply decrypted string from sCryptBuffer to private buffer for OutBuffer property
  sOutBuffer = mid$(sCryptBuffer, 1, lCryptBufLen)

Decrypt_Exit:
On Error Resume Next
  If (lHCryptKey) Then lResult = CryptDestroyKey(lHCryptKey)
  If lHExchgKey Then CryptDestroyKey (lHExchgKey)
  If lHHash Then CryptDestroyHash (lHHash)
  If lHCryptprov Then lResult = CryptReleaseContext(lHCryptprov, 0)
  lStatus = CFB_READY 'switch Status property
Exit Sub

Decrypt_Error:
MsgBox "Error " & err.Number & " (" & err.Description & ") " & _
  "in procedure Decrypt of UtilSecurity:Class Module" & vbLf & _
  InfoErrMsg(), vbExclamation, "Error occurred..."
Resume Decrypt_Exit
End Sub

'Encrypt Decrypt with .NET
'============================================================================================================
Public Function RSAEncryptDecrypt(ByVal Value As String)
Dim MS_DEF_PROV As String
MS_DEF_PROV = "Microsoft Base Cryptographic Provider v1.0"
Dim csp As Object
  Set csp = CreateObject("System.Security.Cryptography.CspParameters")
  csp.ProviderType = 1
  csp.ProviderName = MS_DEF_PROV
  csp.Flags = 2 'CspProviderFlags.UseDefaultKeyContainer
  Dim rsa As Object
  Set rsa = CreateObject("System.Security.Cryptography.RSACryptoServiceProvider")
  Dim str As String
  Dim i As Long
  str = Value
  
  Dim abData() As Byte
  ' Convert string to bytes
  abData = StrConv(str, vbFromUnicode)
  Debug.Print "UN ENCRYPTED ---------------"
  For i = 0 To UBound(abData)
      'Debug.Print Hex(abData(i)); "='" & Chr(abData(i)) & "'"
  Next
  Dim r As Variant
  r = rsa.encrypt(abData, True)
  r = StrConv(r, vbUnicode)
  Debug.Print "ENCRYPTED ---------------"
  Debug.Print r
  
  
  Debug.Print "DENCRYPTED ---------------"
  abData = StrConv(r, vbFromUnicode)
  abData = rsa.Decrypt(abData, True)
  r = StrConv(abData, vbUnicode)
  Debug.Print r
End Function

'FROM EX Crypt
Public Function RsaExportKey(KeyType As EnmRsaKeyTypes, KeyPart As EnmRsaKeyParts, _
        Optional Password As String = vbNullString, Optional Algid As EnmCryptAlgo = CALG_DES) As Byte()
    Dim btKeyData() As Byte
    Dim lKeyLen As Long
    Dim hRsaKey As Long
    Dim hCurrSessionKey As Long
    Dim lCurrAlgID As Long
    Dim hCryptProv As Long
    Dim hSessionKey As Long
    Dim lCryptAlgo As Long
    
    If KeyType = CALG_RSA_KEYX Then
        If CryptGetUserKey(hCryptProv, AT_KEYEXCHANGE, hRsaKey) = 0 Then
            err.raise err.LastDllError, , "No KeyExchange Key in the container!"
        End If
    Else
        If CryptGetUserKey(hCryptProv, AT_SIGNATURE, hRsaKey) = 0 Then
            err.raise err.LastDllError, , "No Signature Key in the container!"
        End If
    End If
    If KeyPart = PRIVATEKEYBLOB And StrPtr(Password) = 0 Then
        err.raise 513, , "No password for private key export provided!"
    End If


    'aktuellen Klassen-Sessionkey sichern:
    hCurrSessionKey = hSessionKey
    hSessionKey = 0    'damit er nicht von CreateKey zerstört wird
    lCurrAlgID = lCryptAlgo
    
    If KeyPart = PRIVATEKEYBLOB Then
        'symmetrischen Schlüssel für den Export erstellen
        btKeyData = Password
       'CreateKey btKeyData, AlgID
    Else
        'Public KeyBlob->nicht verschlüsseln
        hSessionKey = 0
    End If
    
    If CryptExportKey(hRsaKey, hSessionKey, KeyPart, 0, ByVal 0, lKeyLen) <> 0 Then
        ReDim btKeyData(lKeyLen - 1)
        If CryptExportKey(hRsaKey, hSessionKey, KeyPart, 0, btKeyData(0), lKeyLen) <> 0 Then
            'Success
            If lKeyLen <> UBound(btKeyData) + 1 Then
                ReDim Preserve btKeyData(lKeyLen - 1)
            End If
        Else
            err.raise err.LastDllError, , "CryptExportKey Error!"
        End If
    Else
        err.raise err.LastDllError, , "CryptExportKey Error!"
    End If

    'etwaigen temp key zerstören und Klassensessionkey wiederherstellen:
    If hSessionKey = 0 Then
    ElseIf CryptDestroyKey(hSessionKey) = 0 Then
        err.raise err.LastDllError, , "CryptDestroyKey Error"
    End If
    hSessionKey = hCurrSessionKey
    lCryptAlgo = lCurrAlgID
    
    RsaExportKey = btKeyData
End Function


Public Sub DH_Key_Exchange(Optional ByVal KeyLength As Integer = 1024)
Dim hCryptProv_1 As Long    ' Provider Kontext Party 1
Dim hCryptProv_2 As Long    ' Provider Kontext Party 2
Dim hDHKey_1 As Long        ' Private Key Kontext Party 1
Dim hDHKey_2 As Long        ' Private Key Kontext Party 2
Dim hDHKeyPublic As Long    ' Public Key Kontext Party 1
Dim hDHKeyPublic_2 As Long  ' Public Key Kontext Party 2
Dim hDHSessionKey_1 As Long ' SessionKey Kontext for Party 1
Dim hDHSessionKey_2 As Long ' SessionKey Kontext for Party 2
Dim hDHKeyPrivate As Long
Dim lKeyLen As Long
Dim DH_A() As Byte          ' A_1 Public Key Party 1
Dim DH_K() As Byte          ' K_1 Session Key Party 1
Dim DH_A_2() As Byte        ' A_2 Public Key Party 2
Dim DH_K_2() As Byte        ' K_2 Session Key Party 2
Dim lEncryptedDataLen As Long
Dim sData() As Byte
Dim pP As Long
Dim sContainer As String
  
    sContainer = vbNullString
        
    'MS Enhanced DSS and D.H. Provider for Party 1
    If CryptAcquireContext(hCryptProv_1, sContainer, MS_ENH_DSS_DH_PROV, PROV_DSS_DH, CRYPT_VERIFYCONTEXT) = 0 Then
        'fail
        err.raise 1, , "CryptoInitialize : CryptAcquireContext Error"
    End If

    
    KeyLength = KeyLength And Not 7
    If KeyLength < 384 Or KeyLength > 16384 Then
         err.raise 1, , "DH_CreateKeyPair : Invalid Keylength!"
    End If
    
    '  Create Keys for Party 1
    If CryptGenKey(hCryptProv_1, CALG_DH_SF, KeyLength * &H10000 Or CRYPT_EXPORTABLE, hDHKey_1) = 0 Then
        If err.LastDllError = NTE_BAD_ALGID Then
            err.raise 1, , "DH_CreateKeys : Algorithm not supported! (" & CALG_DH_SF & ")"
        Else
            err.raise 1, , "DH_CreateKeys : CryptGenKey Error"
        End If
    
    End If
    
    
    'DH-Public-Key für Party 1 exportieren A=(G^X) mod P:
    If CryptGetUserKey(hCryptProv_1, AT_KEYEXCHANGE, hDHKey_1) = 0 Then
        err.raise 1, , "DH_Export_Public_Key : No KeyExchange Key in the container!"
    Else
        Debug.Print "DH_Export_Public_Key : ExchangeKey vorhanden... " & VBA.Date & "  " & VBA.Time
        If CryptExportKey(hDHKey_1, hDHKeyPublic, PUBLICKEYBLOB, 0, 0, lKeyLen) <> 0 Then
            ReDim DH_A(lKeyLen - 1)
            If CryptExportKey(hDHKey_1, hDHKeyPublic, PUBLICKEYBLOB, 0, _
                    VarPtr(DH_A(LBound(DH_A))), lKeyLen) = 0 Then
                err.raise 1, , "DH_CreateKeys : CryptExportKey Error"
            Else
                'Success
                'Debug.Print "Public Key:", Util.Crypt.ConvToHexString(DH_A)
                'modCryptAPI.SaveFileString app.path & "\" & "DH_PublicKey.txt", util.Crypt.ConvToHexString(DH_A)
                Debug.Print "Item:", "DH_Export_Public_Key : DH-Public-Key erfolgreich exportiert..." & "  " & VBA.Date & "  " & VBA.Time
            End If
        Else
            err.raise 1, , "DH_CreateKeys : CryptExportKey Error"
        End If
    End If
    
    'MS Enhanced DSS and D.H. Provider for Party 2
    If CryptAcquireContext(hCryptProv_2, 0, MS_ENH_DSS_DH_PROV, PROV_DSS_DH, CRYPT_VERIFYCONTEXT) = 0 Then
        'fail
        err.raise 1, , "CryptoInitialize Party 2 : CryptAcquireContext Error"
    Else
        Debug.Print "MS Enhanced DSS and D.H. Crypto Provider initialisiert " & VBA.Date & "  " & VBA.Time
        '  Create Keys for Party 2
        If CryptGenKey(hCryptProv_2, CALG_DH_SF, KeyLength * &H10000 Or CRYPT_EXPORTABLE, hDHKey_2) = 0 Then
            If err.LastDllError = NTE_BAD_ALGID Then
                err.raise 1, , "DH_CreateKeys Party 2 : Algorithm not supported! (" & CALG_DH_SF & ")"
            Else
                err.raise 1, , "DH_CreateKeys Party 2 : CryptGenKey Error" & vbCr & err.LastDllError
                
            End If
        Else
            'Success
            Debug.Print "DH_CreateKeys Party 2 : DH-Keys erfolgreich erstellt..." & "  " & VBA.Date & "  " & VBA.Time
                    
                        
            ' DH-Public Key Party A importieren K=(Y^X) mod P
            If CryptGetUserKey(hCryptProv_2, AT_KEYEXCHANGE, hDHKey_2) = 0 Then
                err.raise 1, , "DH_ImportSessionKey Party 2 : No KeyExchange Key in the container!"
            End If
    
            ' etwaigen bereits vorhandenen Session-Schlüssel zerstören:
            If hDHSessionKey_2 <> 0 Then
                If CryptDestroyKey(hDHSessionKey_2) = 0 Then
                    err.raise 1, , "DH_ImportSessionKey Party 2 : CryptDestroyKey Error"
                End If
            End If
    
            ' Public Key Party 1 importieren und in SessionKey Party 2 umwandeln:
            If CryptImportKey(hCryptProv_2, VarPtr(DH_A(LBound(DH_A))), _
                    UBound(DH_A) - LBound(DH_A) + 1, hDHKey_2, 0, hDHSessionKey_2) = 0 Then
                err.raise 1, , "DH_ImportPublicKey Party 2 : CryptImportKey Error"
            Else
                ' Success
                'modCryptAPI.SaveFileString app.path & "\" & "DH_ImportPubKey_2.txt", Bytes2Long(DH_A)
                Debug.Print "DH_ImportPublicKey Party 2 : PublicKey (K) erfolgreich importiert..." & "  " & VBA.Date & "  " & VBA.Time
            
                ' Session Key umwandeln
                If CryptSetKeyParam(hDHSessionKey_2, KP_ALGID, CALG_RC4, 0) = 0 Then
                    err.raise 1, , "DH_ConvertSessionKey Party 2 : CryptSetKeyParam Error"
                Else
                    'frmChiffre.txtSessionKey.Text = util.Crypt.ConvToHexString(DH_A)
                    'modCryptAPI.SaveFileBinary App.Path & "\" & "DH_SessionKey.txt", DH_A
                    Debug.Print "DH_ConvertToSessionKey Party 2 : SessionKey (S) erfolgreich umgewandelt..." & "  " & VBA.Date & "  " & VBA.Time
                End If
            End If
            
            'DH-Public-Key für Party 2 exportieren A=(G^X) mod P:
            If CryptGetUserKey(hCryptProv_2, AT_KEYEXCHANGE, hDHKey_2) = 0 Then
                err.raise 1, , "DH_Export_Public_Key for Party 2 : No KeyExchange Key in the container!"
            Else
                Debug.Print "DH_Export_Public_Key Party 2 : ExchangeKey vorhanden... " & VBA.Date & "  " & VBA.Time
                If CryptExportKey(hDHKey_2, hDHKeyPublic_2, PUBLICKEYBLOB, 0, 0, lKeyLen) <> 0 Then
                    ReDim DH_A_2(lKeyLen - 1)
                    If CryptExportKey(hDHKey_2, hDHKeyPublic_2, PUBLICKEYBLOB, 0, _
                            VarPtr(DH_A_2(LBound(DH_A_2))), lKeyLen) = 0 Then
                        err.raise 1, , "DH_CreateKeys Party 2 : CryptExportKey Error"
                    Else
                        'Success
                        'Debug.Print "Private key:", Util.Crypt.ConvToHexString(DH_A_2)
                        'MsgBox "DH Public Key: " & Util.Crypt.ConvToHexString(DH_A_2)
                        'modCryptAPI.SaveFileString app.path & "\" & "DH_PublicKey2.txt", util.Crypt.ConvToHexString(DH_A_2)
                        Debug.Print "DH_Export_Public_Key Party 2 : DH-Public-Key erfolgreich exportiert..." & "  " & VBA.Date & "  " & VBA.Time
                        
                    End If
                Else
                    err.raise 1, , "DH_CreateKeys Party 2 : CryptExportKey Error"
                End If
            End If
        End If
    End If
    
    
    ' DH-Public Key importieren K=(Y^X) mod P
    If CryptGetUserKey(hCryptProv_1, AT_KEYEXCHANGE, hDHKey_1) = 0 Then
        err.raise 1, , "DH_ImportSessionKey : No KeyExchange Key in the container!"
    End If
    
    ' etwaigen bereits vorhandenen Session-Schlüssel zerstören:
    If hDHSessionKey_1 <> 0 Then
        If CryptDestroyKey(hDHSessionKey_1) = 0 Then
           err.raise 1, , "DH_ImportSessionKey : CryptDestroyKey Error"
        End If
    End If
    
    ' Public Key Party 2 importieren und in SessionKey Party 1 umwandeln :
    If CryptImportKey(hCryptProv_1, VarPtr(DH_A_2(LBound(DH_A_2))), _
            UBound(DH_A_2) - LBound(DH_A_2) + 1, hDHKey_1, 0, hDHSessionKey_1) = 0 Then
        err.raise 1, , "DH_ImportPublicKey : CryptImportKey Error"
    Else
        ' Success
        'frmChiffre.txtSessionKey.Text = util.Crypt.ConvToHexString(DH_A)
        'modCryptAPI.SaveFileString app.path & "\" & "DH_ImportPubKey.txt", Bytes2Long(DH_A_2)
        Debug.Print "DH_ImportPublicKey : PublicKey (K) erfolgreich importiert..." & "  " & VBA.Date & "  " & VBA.Time
            
            ' Session Key umwandeln
            If CryptSetKeyParam(hDHSessionKey_1, KP_ALGID, CALG_RC4, 0) = 0 Then
                err.raise 1, , "DH_ConvertSessionKey : CryptSetKeyParam Error"
            Else: Debug.Print "DH_ConvertToSessionKey : SessionKey (S) erfolgreich umgewandelt..." & "  " & VBA.Date & "  " & VBA.Time
            End If
    End If
    
    'Encrypt Random-Data with Party 1 SessionKey
    ReDim sData(UBound(DH_A))
    sData = DH_A
    
    lKeyLen = UBound(sData) - LBound(sData) + 1
    lEncryptedDataLen = lKeyLen
    
    'Speicherbedarf ermitteln:
    If CryptEncrypt(hDHSessionKey_1, 0, 1, 0, 0, lEncryptedDataLen, lKeyLen) <> 0 Then
        ReDim Preserve sData(LBound(sData) To LBound(sData) + lEncryptedDataLen - 1)
        If CryptEncrypt(hDHSessionKey_1, 0, 1, 0, VarPtr(sData(LBound(sData))), lKeyLen, lEncryptedDataLen) = 0 Then
            err.raise 1, , "Encrypt : CryptEncrypt Error"
        Else
            Debug.Print "Encrypt : Verschlüsselt mit " & CALG_RC4 & "  " & VBA.Date & "  " & VBA.Time
            'MsgBox "Encrypted: " & Util.Crypt.ConvToHexString(sData)
        End If
    Else
        Debug.Print "Encrypt : CryptEncrypt Error... " & VBA.Date & "  " & VBA.Time
        err.raise 1, , "Encrypt : CryptEncrypt Error"
    End If
    
    'Decrypt Data with Party 2 SessionKey
    If CryptDecrypt(hDHSessionKey_2, 0, 1, 0, VarPtr(sData(LBound(sData))), lKeyLen) = 0 Then
        err.raise 1, , "Decrypt : CryptDecrypt Error"
    Else
        Debug.Print "Decrypt : Entschlüsselt mit " & CALG_RC4 & "  " & VBA.Date & "  " & VBA.Time
        'MsgBox "Decrypted: " & Util.Crypt.ConvToHexString(sData)
    End If
    
        
    'vorhandenen Sessionkey zerstören:
    If hDHSessionKey_1 <> 0 Then
        If CryptDestroyKey(hDHSessionKey_1) = 0 Then
            err.raise 1, , ("CryptoTerminate Session Key 1 : CryptDestroyKey Error")
        Else
            Debug.Print "CryptTerminate : SessionKey zerstörrt... " & VBA.Date & "  " & VBA.Time
        End If
    End If
    If hDHSessionKey_2 <> 0 Then
        If CryptDestroyKey(hDHSessionKey_2) = 0 Then
            err.raise 1, , ("CryptoTerminate Session Key 1 : CryptDestroyKey Error")
        Else
            Debug.Print "CryptTerminate : SessionKey zerstörrt... " & VBA.Date & "  " & VBA.Time
        End If
    End If
    
    'Provider Kontext freigeben:
    If hCryptProv_1 <> 0 Then
        If CryptReleaseContext(hCryptProv_1, 0) = 0 Then
           err.raise 1, , ("CryptoTerminate Provider : CryptReleaseContext Error")
        Else
            Debug.Print "CryptTerminate : ReleaseContext zerstörrt... " & VBA.Date & "  " & VBA.Time
        End If
    End If
    If hCryptProv_2 <> 0 Then
        If CryptReleaseContext(hCryptProv_2, 0) = 0 Then
           err.raise 1, , ("CryptoTerminate Provider : CryptReleaseContext Error")
        Else
            Debug.Print "CryptTerminate : ReleaseContext zerstörrt... " & VBA.Date & "  " & VBA.Time
        End If
    End If
End Sub



'-------------------------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------------------------


'
''FROM UtilCryptKeyEx
'Public Sub Command1_Click()
'    Dim hCryptProvA As Long
'    Dim hPrivateKeyA As Long
'    Dim btPublicKeyA() As Byte
'    Dim hSessionKeyA As Long
'
'    Dim hCryptProvB As Long
'    Dim hPrivateKeyB As Long
'    Dim btPublicKeyB() As Byte
'    Dim hSessionKeyB As Long
'
'    Dim cbData As Long
'    Dim btSessionKeyData() As Byte
'    Dim btData() As Byte
'    Dim x As Long
'
'    Dim P_Blob As CRYPT_DATA_BLOB_VB
'    Dim G_Blob As CRYPT_DATA_BLOB_VB
'
'    Const KEY_LENGTH As Long = 1024
'
'    CryptAcquireContext hCryptProvA, "Mein Container B", MS_ENH_DSS_DH_PROV, PROV_DSS_DH, CRYPT_DELETEKEYSET
'    CryptAcquireContext hCryptProvB, "Mein Container A", MS_ENH_DSS_DH_PROV, PROV_DSS_DH, CRYPT_DELETEKEYSET
'
'
'    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'    'Partner A erstellt einen RSA KeyExchange Schlüssel und exportiert den öffentlichen
'    'Schlüsselteil:
'
'    '--------------------------------------------------------------------------------------
'    'Provider Kontext anfordern und Schlüssel erstellen:
'    If CryptAcquireContext(hCryptProvA, "Mein Container A", MS_ENH_DSS_DH_PROV, PROV_DSS_DH, CRYPT_NEWKEYSET) = 0 Then
'        'Fehler->Existiert der Container bereits?
'        If CryptAcquireContext(hCryptProvA, "Mein Container A", MS_ENH_DSS_DH_PROV, PROV_DSS_DH, 0) = 0 Then
'            MsgBox "Kann MS Base Provider nicht finden!", vbExclamation, "Fehler:"
'        End If
'    End If
'
'    If CryptGenKey(hCryptProvA, CALG_DH_SF, KEY_LENGTH * &H10000 Or CRYPT_EXPORTABLE, hPrivateKeyA) = 0 Then
'        MsgBox "Kann Schlüssel nicht erstellen!", vbExclamation, "Fehler:"
'    End If
'
'    ReDim P_Blob.data(KEY_LENGTH / 8 - 1)
'    P_Blob.cbData = KEY_LENGTH / 8
'    P_Blob.pbData = VarPtr(P_Blob.data(0))
'
'    ReDim G_Blob.data(KEY_LENGTH / 8 - 1)
'    G_Blob.cbData = KEY_LENGTH / 8
'    G_Blob.pbData = VarPtr(G_Blob.data(0))
'
'
'    If CryptGetKeyParam(hPrivateKeyA, KP_G, G_Blob.data(0), G_Blob.cbData, 0) = 0 Then
'        MsgBox "Kann G nicht lesen!", vbExclamation, "Fehler:"
'    ElseIf CryptGetKeyParam(hPrivateKeyA, KP_P, P_Blob.data(0), P_Blob.cbData, 0) = 0 Then
'        MsgBox "Kann P nicht setzen!", vbExclamation, "Fehler:"
'    End If
'
'    '--------------------------------------------------------------------------------------
'    'Öffentlichen Teil des KeyExchange Schlüssels exportieren:
'    If CryptExportKey(hPrivateKeyA, 0, PUBLICKEYBLOB, 0, ByVal 0, cbData) = 0 Then
'        MsgBox "Kann Schlüssel Länge nicht ermitteln!", vbExclamation, "Fehler:"
'    Else
'        'Speicher reservieren und Schlüsseldaten abholen:
'        ReDim Preserve btPublicKeyA(cbData - 1)
'        If CryptExportKey(hPrivateKeyA, 0, PUBLICKEYBLOB, 0, btPublicKeyA(0), cbData) = 0 Then
'            MsgBox "Kann Schlüssel nicht exportieren!", vbExclamation, "Fehler:"
'        End If
'    End If
'
'    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'    'Die öffentlichen Schlüsseldaten werden nun an den Partner B übermittelt, dort importiert,
'    'ein zufälliger symmetrischer Sessionkey erstellt und dieser exportiert wobei die
'    'Schlüsseldaten mit dem öffentlichen Schlüssel von Partner A verschlüsselt werden:
'
'    '--------------------------------------------------------------------------------------
'    If CryptAcquireContext(hCryptProvB, "Mein Container B", MS_ENH_DSS_DH_PROV, PROV_DSS_DH, CRYPT_NEWKEYSET) = 0 Then
'        'Fehler->Existiert der Container bereits?
'        If CryptAcquireContext(hCryptProvB, "Mein Container B", MS_ENH_DSS_DH_PROV, PROV_DSS_DH, 0) = 0 Then
'            MsgBox "Kann MS Base Provider nicht finden!", vbExclamation, "Fehler:"
'        End If
'    End If
'
'    If CryptGenKey(hCryptProvB, CALG_DH_SF, KEY_LENGTH * &H10000 Or CRYPT_EXPORTABLE Or CRYPT_PREGEN, hPrivateKeyB) = 0 Then
'        MsgBox "Kann Schlüssel nicht erstellen!", vbExclamation, "Fehler:"
'    End If
'
'    If CryptSetKeyParam(hPrivateKeyB, KP_G, G_Blob, 0) = 0 Then
'        MsgBox "Kann G nicht setzen!", vbExclamation, "Fehler:"
'    ElseIf CryptSetKeyParam(hPrivateKeyB, KP_P, P_Blob, 0) = 0 Then
'        MsgBox "Kann P nicht setzen!", vbExclamation, "Fehler:"
'    ElseIf CryptSetKeyParam(hPrivateKeyB, KP_X, ByVal 0, 0) = 0 Then
'        MsgBox "Kann X nicht generieren!", vbExclamation, "Fehler:"
'    End If
'
'
'    'Öffentlichen Schlüssel importieren:
'    If CryptImportKey(hCryptProvB, btPublicKeyA(0), UBound(btPublicKeyA) + 1, hPrivateKeyB, 0, hSessionKeyB) = 0 Then
'        MsgBox "Kann Schlüssel nicht importieren!", vbExclamation, "Fehler:"
'    End If
'
'
'
'    If CryptExportKey(hPrivateKeyB, 0, PUBLICKEYBLOB, 0, ByVal 0, cbData) = 0 Then
'        MsgBox "Kann Schlüssel Länge nicht ermitteln!", vbExclamation, "Fehler:"
'    Else
'        'Speicher reservieren und Schlüsseldaten abholen:
'        ReDim Preserve btPublicKeyB(cbData - 1)
'        If CryptExportKey(hPrivateKeyB, 0, PUBLICKEYBLOB, 0, btPublicKeyB(0), cbData) = 0 Then
'            MsgBox "Kann Schlüssel nicht exportieren!", vbExclamation, "Fehler:"
'        End If
'    End If
'
'    If CryptImportKey(hCryptProvA, btPublicKeyB(0), UBound(btPublicKeyB) + 1, hPrivateKeyA, 0, hSessionKeyA) = 0 Then
'        MsgBox "Kann Schlüssel nicht importieren!", vbExclamation, "Fehler:"
'    End If
'
'    If CryptSetKeyParam(hSessionKeyA, KP_ALGID, CALG_RC4, 0) = 0 Then
'         MsgBox "Kann Schlüssel nicht umwandeln!", vbExclamation, "Fehler:"
'    ElseIf CryptSetKeyParam(hSessionKeyB, KP_ALGID, CALG_RC4, 0) = 0 Then
'         MsgBox "Kann Schlüssel nicht umwandeln!", vbExclamation, "Fehler:"
'    End If
'
'
'
'    btData = "Hello World!"
'    'Länge der verschlüsselten Daten ermitteln:
'    cbData = UBound(btData) + 1
'    If CryptEncrypt(hSessionKeyA, 0, 1, 0, 0, cbData, 0) = 0 Then
'        MsgBox "Kann Länge nicht ermitteln!", vbExclamation, "Fehler:"
'    Else
'        'Speicher reservieren und verschlüsselte Daten abholen:
'        x = UBound(btData) + 1
'        ReDim Preserve btData(cbData - 1)
'        cbData = x
'        If CryptEncrypt(hSessionKeyA, 0, 1, 0, VarPtr(btData(0)), cbData, UBound(btData) + 1) = 0 Then
'            MsgBox "Kann Daten nicht verschlüsseln!", vbExclamation, "Fehler:"
'        Else
'            'jetzt wieder entschlüsseln:
'            If CryptDecrypt(hSessionKeyB, 0, 1, 0, VarPtr(btData(0)), cbData) = 0 Then
'                MsgBox "Kann Daten nicht entschlüsseln!", vbExclamation, "Fehler:"
'            Else
'                ReDim Preserve btData(cbData - 1)
'                MsgBox "Wieder entschlüsselt: " & CStr(btData), vbInformation, "Erfolgreich:"
'            End If
'        End If
'    End If
'
'    If hSessionKeyA <> 0 Then CryptDestroyKey hSessionKeyA
'    If hSessionKeyB <> 0 Then CryptDestroyKey hSessionKeyB
'    If hCryptProvA <> 0 Then CryptReleaseContext hCryptProvA, 0
'    If hCryptProvB <> 0 Then CryptReleaseContext hCryptProvB, 0
'    CryptAcquireContext hCryptProvA, "Mein Container B", MS_DEF_PROV, PROV_RSA_FULL, CRYPT_DELETEKEYSET
'    CryptAcquireContext hCryptProvB, "Mein Container A", MS_DEF_PROV, PROV_RSA_FULL, CRYPT_DELETEKEYSET
'End Sub


'--------------------------------------------------------------------------------------------

'Public Function RsaExportKey(KeyType As EnmRsaKeyTypes, KeyPart As EnmRsaKeyParts, _
'        Optional Password As String = vbNullString, Optional Algid As EnmCryptAlgo = CALG_DES) As Byte()
'    Dim btKeyData() As Byte
'    Dim lKeyLen As Long
'    Dim hRsaKey As Long
'    Dim hCurrSessionKey As Long
'    Dim lCurrAlgID As Long
'    Dim hCryptProv As Long
'    Dim hSessionKey As Long
'    Dim lCryptAlgo As Long
'
'    If KeyType = CALG_RSA_KEYX Then
'        If CryptGetUserKey(hCryptProv, AT_KEYEXCHANGE, hRsaKey) = 0 Then
'            err.raise err.LastDllError, , "No KeyExchange Key in the container!"
'        End If
'    Else
'        If CryptGetUserKey(hCryptProv, AT_SIGNATURE, hRsaKey) = 0 Then
'            err.raise err.LastDllError, , "No Signature Key in the container!"
'        End If
'    End If
'    If KeyPart = PRIVATEKEYBLOB And StrPtr(Password) = 0 Then
'        err.raise 513, , "No password for private key export provided!"
'    End If
'
'
'    'aktuellen Klassen-Sessionkey sichern:
'    hCurrSessionKey = hSessionKey
'    hSessionKey = 0    'damit er nicht von CreateKey zerstört wird
'    lCurrAlgID = lCryptAlgo
'
'    If KeyPart = PRIVATEKEYBLOB Then
'        'symmetrischen Schlüssel für den Export erstellen
'        btKeyData = Password
'       'CreateKey btKeyData, AlgID
'    Else
'        'Public KeyBlob->nicht verschlüsseln
'        hSessionKey = 0
'    End If
'
'    If CryptExportKey(hRsaKey, hSessionKey, KeyPart, 0, ByVal 0, lKeyLen) <> 0 Then
'        ReDim btKeyData(lKeyLen - 1)
'        If CryptExportKey(hRsaKey, hSessionKey, KeyPart, 0, btKeyData(0), lKeyLen) <> 0 Then
'            'Success
'            If lKeyLen <> UBound(btKeyData) + 1 Then
'                ReDim Preserve btKeyData(lKeyLen - 1)
'            End If
'        Else
'            err.raise err.LastDllError, , "CryptExportKey Error!"
'        End If
'    Else
'        err.raise err.LastDllError, , "CryptExportKey Error!"
'    End If
'
'    'etwaigen temp key zerstören und Klassensessionkey wiederherstellen:
'    If hSessionKey = 0 Then
'    ElseIf CryptDestroyKey(hSessionKey) = 0 Then
'        err.raise err.LastDllError, , "CryptDestroyKey Error"
'    End If
'    hSessionKey = hCurrSessionKey
'    lCryptAlgo = lCurrAlgID
'
'    RsaExportKey = btKeyData
'End Function


'--------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------

'ref: http://www.utteraccess.com/forum/Acccess-Card-Readers-t1986575.html


