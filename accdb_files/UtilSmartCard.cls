VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UtilSmartCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************************
' winscard.dll Visual Basic Function Prototypes
'**************************************************************************
'http://msdn.microsoft.com/en-us/library/aa379466(VS.85).aspx
Private Declare PtrSafe Function SCardAccessStartedEvent Lib "winscard.dll" _
  () As Long
  
'http://msdn.microsoft.com/en-us/library/aa379468(VS.85).aspx
Private Declare PtrSafe Function SCardAddReaderToGroup Lib "winscard.dll" Alias "SCardAddReaderToGroupA" _
  (ByVal hContext As Long, ByVal szReaderName As String, ByVal szGroupName As String) As Long
  
'http://msdn.microsoft.com/en-us/library/aa379469(VS.85).aspx
Private Declare PtrSafe Function SCardBeginTransaction Lib "winscard.dll" _
  (ByVal hCard As Long) As Long
  
'http://msdn.microsoft.com/en-us/library/aa379470(VS.85).aspx
Private Declare PtrSafe Function SCardCancel Lib "winscard.dll" _
  (ByVal hContext As Long) As Long
  
'http://msdn.microsoft.com/en-us/library/aa379473(VS.85).aspx
Private Declare PtrSafe Function SCardConnect Lib "winscard.dll" Alias "SCardConnectA" _
  (ByVal hContext As Long, ByVal szReader As String, ByVal dwShareMode As Long, _
  ByVal dwPreferredProtocols As Long, ByRef phCard As Long, ByRef pdwActiveProtocol As Long) As Long
  
'http://msdn.microsoft.com/en-us/library/aa379474(VS.85).aspx
Private Declare PtrSafe Function SCardControl Lib "winscard.dll" _
  (ByVal hCard As Long, ByVal dwControlCode As Long, ByRef lpInBuffer As Long, _
  ByVal nInBufferSize As Long, ByRef lpOutBuffer As Long, ByVal nOutBufferSize As Long, _
  ByRef lpBytesReturned As Long) As Long
  
'http://msdn.microsoft.com/en-us/library/aa379475(VS.85).aspx
Private Declare PtrSafe Function SCardDisconnect Lib "winscard.dll" _
  (ByVal hCard As Long, ByVal dwDisposition As Long) As Long
  
'http://msdn.microsoft.com/en-us/library/aa379477(VS.85).aspx
Private Declare PtrSafe Function SCardEndTransaction Lib "winscard.dll" _
  (ByVal hCard As Long, ByVal dwDisposition As Long) As Long
  
'http://msdn.microsoft.com/en-us/library/aa379479(VS.85).aspx
Private Declare PtrSafe Function SCardEstablishContext Lib "winscard.dll" _
  (ByVal dwScope As Long, ByVal pvReserved1 As Long, ByVal pvReserved2 As Long, _
  ByRef phContext As Long) As Long
  
'http://msdn.microsoft.com/en-us/library/aa379482(VS.85).aspx
Private Declare PtrSafe Function SCardForgetCardType Lib "winscard.dll" Alias "SCardForgetCardTypeA" _
  (ByVal hContext As Long, ByVal szCardName As String) As Long
  
'http://msdn.microsoft.com/en-us/library/aa379484(VS.85).aspx
Private Declare PtrSafe Function SCardForgetReader Lib "winscard.dll" Alias "SCardForgetReaderA" _
  (ByVal hContext As Long, ByVal szReaderName As String) As Long
  
'http://msdn.microsoft.com/en-us/library/aa379486(VS.85).aspx
Private Declare PtrSafe Function SCardForgetReaderGroup Lib "winscard.dll" Alias "SCardForgetReaderGroupA" _
  (ByVal hContext As Long, ByVal szGroupName As String) As Long
  
'http://msdn.microsoft.com/en-us/library/aa379488(VS.85).aspx
Private Declare PtrSafe Function SCardFreeMemory Lib "winscard.dll" _
  (ByVal hContext As Long, ByVal pvMem As Long) As Long
  
'http://msdn.microsoft.com/en-us/library/aa379559(VS.85).aspx
Private Declare PtrSafe Function SCardGetAttrib Lib "winscard.dll" _
  (ByVal hCard As Long, ByVal dwAttrId As Long, ByRef pbAttr As ByteArray, _
  ByRef pcbAttrLen As Long) As Long
  
'http://msdn.microsoft.com/en-us/library/aa379655(VS.85).aspx
Private Declare PtrSafe Function SCardGetCardTypeProviderName Lib "winscard.dll" Alias "SCardGetCardTypeProviderNameA" _
  (ByVal hContext As Long, ByVal szCardName As String, ByVal dwProviderId As Long, _
  ByVal szProvider As String, ByRef pcchProvider As Long) As Long
  
'http://msdn.microsoft.com/en-us/library/aa379761(VS.85).aspx
Private Declare PtrSafe Function SCardGetProviderId Lib "winscard.dll" Alias "SCardGetProviderIdA" _
  (ByVal hContext As Long, ByVal szCard As String, ByRef pguidProviderId As GUID) As Long
  
'http://msdn.microsoft.com/en-us/library/aa379773(VS.85).aspx
Private Declare PtrSafe Function SCardGetStatusChange Lib "winscard.dll" Alias "SCardGetStatusChangeA" _
  (ByVal hContext As Long, ByVal dwTimeout As Long, ByRef rgReaderStates() As SCARD_READERSTATE, _
  ByVal cReaders As Long) As Long
  
'http://msdn.microsoft.com/en-us/library/aa379784(VS.85).aspx
Private Declare PtrSafe Function SCardIntroduceCardType Lib "winscard.dll" Alias "SCardIntroduceCardTypeA" _
  (ByVal hContext As Long, ByVal szCardName As String, ByRef pguidPrimaryProvider As GUID, _
  ByRef pguidInterfaces As GUID, ByVal dwInterfaceCount As Long, ByRef pbAtr As ByteArray, _
  ByRef pbAtrMask As ByteArray, ByVal cbAtrLen As Long) As Long
  
'http://msdn.microsoft.com/en-us/library/aa379786(VS.85).aspx
Private Declare PtrSafe Function SCardIntroduceReader Lib "winscard.dll" Alias "SCardIntroduceReaderA" _
  (ByVal hContext As Long, ByVal szReaderName As String, ByVal szDeviceName As String) As Long
  
'http://msdn.microsoft.com/en-us/library/aa379787(VS.85).aspx
Private Declare PtrSafe Function SCardIntroduceReaderGroup Lib "winscard.dll" Alias "SCardIntroduceReaderGroupA" _
  (ByVal hContext As Long, ByVal szGroupName As String) As Long
  
'http://msdn.microsoft.com/en-us/library/aa379788(VS.85).aspx
Private Declare PtrSafe Function SCardIsValidContext Lib "winscard.dll" _
  (ByVal hContext As Long) As Long
  
'http://msdn.microsoft.com/en-us/library/aa379789(VS.85).aspx
Private Declare PtrSafe Function SCardListCards Lib "winscard.dll" Alias "SCardListCardsA" _
  (ByVal hContext As Long, ByRef pbAtr As ByteArray, ByVal rgguidInterfaces As Long, _
  ByVal cguidInterfaceCount As Long, ByVal mszCards As String, ByRef pcchCards As Long) As Long
  
'http://msdn.microsoft.com/en-us/library/aa379790(VS.85).aspx
Private Declare PtrSafe Function SCardListInterfaces Lib "winscard.dll" Alias "SCardListInterfacesA" _
  (ByVal hContext As Long, ByVal szCard As String, ByRef pguidInterfaces As GUID, _
  ByRef pcguidInterfaces As Long) As Long
  
'http://msdn.microsoft.com/en-us/library/aa379792(VS.85).aspx
Private Declare PtrSafe Function SCardListReaderGroups Lib "winscard.dll" Alias "SCardListReaderGroupsA" _
  (ByVal hContext As Long, ByVal mszGroups As String, ByRef pcchGroups As Long) As Long
  
'http://msdn.microsoft.com/en-us/library/aa379793(VS.85).aspx
Private Declare PtrSafe Function SCardListReaders Lib "winscard.dll" Alias "SCardListReadersA" _
  (ByVal hContext As Long, ByVal mszGroups As String, ByVal mszReaders As String, _
  ByRef pcchReaders As Long) As Long
  
'http://msdn.microsoft.com/en-us/library/aa379794(VS.85).aspx
Private Declare PtrSafe Function SCardLocateCards Lib "winscard.dll" Alias "SCardLocateCardsA" _
  (ByVal hContext As Long, ByVal mszCards As String, ByRef rgReaderStates() As SCARD_READERSTATE, _
  ByVal cReaders As Long) As Long
  
'http://msdn.microsoft.com/en-us/library/aa379796(VS.85).aspx
Private Declare PtrSafe Function SCardLocateCardsByATR Lib "winscard.dll" Alias "SCardLocateCardsByATRA" _
  (ByVal hContext As Long, ByRef rgAtrMasks() As SCARD_ATRMASK, ByVal cAtrs As Long, _
  ByRef rgReaderStates() As SCARD_READERSTATE, ByVal cReaders As Long) As Long
  
'http://msdn.microsoft.com/en-us/library/aa379797(VS.85).aspx
Private Declare PtrSafe Function SCardReconnect Lib "winscard.dll" _
  (ByVal hCard As Long, ByVal dwShareMode As Long, ByVal dwPreferredProtocols As Long, _
  ByVal dwInitialization As Long, ByRef pdwActiveProtocol As Long) As Long
  
'http://msdn.microsoft.com/en-us/library/aa379798(VS.85).aspx
Private Declare PtrSafe Function SCardReleaseContext Lib "winscard.dll" _
  (ByVal hContext As Long) As Long
  
'http://msdn.microsoft.com/en-us/library/aa379799(VS.85).aspx
Private Declare PtrSafe Sub SCardReleaseStartedEvent Lib "winscard.dll" _
  (ByVal hStartedEventHandle As Long)
  
'http://msdn.microsoft.com/en-us/library/aa379800(VS.85).aspx
Private Declare PtrSafe Function SCardRemoveReaderFromGroup Lib "winscard.dll" Alias "SCardRemoveReaderFromGroupA" _
  (ByVal hContext As Long, ByVal szReaderName As String, ByVal szGroupName As String) As Long
  
'http://msdn.microsoft.com/en-us/library/aa379801(VS.85).aspx
Private Declare PtrSafe Function SCardSetAttrib Lib "winscard.dll" _
  (ByVal hCard As Long, ByVal dwAttrId As Long, ByRef pbAttr As ByteArray, ByVal cbAttrLen As Long) As Long
  
'http://msdn.microsoft.com/en-us/library/aa379802(VS.85).aspx
Private Declare PtrSafe Function SCardSetCardTypeProviderName Lib "winscard.dll" Alias "SCardSetCardTypeProviderNameA" _
  (ByVal hContext As Long, ByVal szCardName As String, ByVal dwProviderId As Long, ByVal szProvider As String) As Long
  
'https://msdn.microsoft.com/en-us/library/windows/desktop/aa379803(v=vs.85).aspx
Private Declare PtrSafe Function SCardState Lib "winscard.dll" _
  (ByVal hCard As Long, ByRef pdwState As Long, ByRef pdwProtocol As Long, _
  ByRef pbAtr As ByteArray, ByRef pcbAtrLen As Long) As Long
  
'http://msdn.microsoft.com/en-us/library/aa379803(VS.85).aspx
Private Declare PtrSafe Function SCardStatus Lib "winscard.dll" Alias "SCardStatusA" _
  (ByVal hCard As Long, ByVal szReaderName As String, ByRef pcchReaderLen As Long, _
  ByRef pdwState As Long, ByRef pdwProtocol As Long, ByRef pbAtr As ByteArray, ByRef pcbAtrLen As Long) As Long
  
'http://msdn.microsoft.com/en-us/library/aa379804.aspx
Private Declare PtrSafe Function SCardTransmit Lib "winscard.dll" _
  (ByVal hCard As Long, ByRef pioSendPci As SCARD_IO_REQUEST, ByRef pbSendBuffer As Byte, _
  ByVal cbSendLength As Long, ByRef pioRecvPci As SCARD_IO_REQUEST, _
  ByRef pbRecvBuffer As Byte, ByRef pcbRecvLength As Long) As Long
  
'https://msdn.microsoft.com/en-us/library/windows/desktop/aa379804(v=vs.85).aspx
Private Declare PtrSafe Function SCardTransmitLong Lib "winscard.dll" Alias "SCardTransmit" _
  (ByVal hCard As Long, ByRef pioSendPci As SCARD_IO_REQUEST, ByRef pbSendBuffer As Byte, _
  ByVal cbSendLength As Long, ByVal pioRecvPci As Long, ByRef pbRecvBuffer As Byte, _
  ByRef pcbRecvLength As Long) As Long

'http://msdn.microsoft.com/en-us/library/aa379808(v=vs.85).aspx
Private Type SCARD_READERSTATE
  szReader As String
  pvUserData As Long
  dwCurrentState As Long
  dwEventState As Long
  cbAtr As Long
  rgbAtr As Byte
End Type

'http://msdn.microsoft.com/en-us/library/aa922903.aspx
Private Type SCARD_ATRMASK
  cbAtr As Long
  rgbAtr As Byte
  rgbMask As Byte
End Type

'http://msdn.microsoft.com/en-us/library/aa379807.aspx
Private Type SCARD_IO_REQUEST
  dwProtocol As Long
  dwPciLength As Long
End Type

Private Const DEFAULT_BUFFER_SIZE As Integer = 255
Private Type ByteArray
  b(0 To DEFAULT_BUFFER_SIZE) As Byte
End Type

Private Type GUID
  Data1 As Long
  Data2 As Integer
  Data3 As Integer
  Data4(0 To 7) As Byte
End Type

'used in SCardEstablishContext
Private Const SCARD_SCOPE_USER As Long = 0
Private Const SCARD_SCOPE_SYSTEM As Long = 2
Private Const SCARD_S_SUCCESS As Long = 0


'https://stackoverflow.com/questions/31592198/basic-smart-card-testing-in-windows
Public Sub Test()
  Dim ret As Long
  Dim hSCardContext As Long
  
  If Not (SCardEstablishContext(SCARD_SCOPE_USER, 0, 0, hSCardContext) = SCARD_S_SUCCESS) Then
    err.raise Number:=ErrCode.CannotPerformRequestedOperation, _
      Description:="Error in procedure SCardEstablishContext"
  End If
  
  Dim res As Long
  
  Dim sReaderList As String
  Dim RdrCount As Long
  Dim RetCode As Long
  Dim sGroup As String
  sGroup = VBA.String(0, vbNullChar)
  sReaderList = VBA.String(1024, vbNullChar)
  Dim r As String
  RdrCount = Len(sReaderList)
  RetCode = SCardListReaders(hSCardContext, sGroup, sReaderList, RdrCount)
  
  Debug.Print sReaderList
  
  
'  Select Case RetCode
'      Case -2
'          MsgBox "SCardListReaders Fail!. Connect the MSR and try again.", vbInformation
'      Case Is = SCARD_S_SUCCESS
'          Debug.Print "SCardEstablishContext OK..."
'       Case Else
'          Debug.Print "Established Reader Failed. " & Hex(RetCode)
'  End Select
'
  Dim readers As New Collection
  
  Dim i As Integer, sTemp As String, idx As Long
  sTemp = sReaderList
  For i = 1 To Util.Strings.getNumberOfOccurances(sReaderList, "0")
    idx = Util.Strings.GetIndexOf(sTemp, "0")
    If idx = 0 Then Exit For
    readers.Add Util.Strings.GetLeftOfIndex(sTemp, idx)
    sTemp = Util.Strings.GetRightOfIndex(sTemp, idx)
  Next i
  
  Dim reader As Variant
  For Each reader In readers
    Debug.Print reader
  Next reader
  
  
  'Debug.Print Util.Strings.getNumberOfOccurances(sReaderList, "0")
 ' Debug.Print InStr(1, sReaderList, "\0", vbBinaryCompare)
  
  
  
  'res = SCardListReaders(hSCardContext, 0, 0, readers)
'har *szGroups = NULL;
'long readers = 0;
'int res = SCardListReaders(hSCardContext, szGroups, NULL, &readers);
'// handle errors
'
'you get readers count first. now you can allocate memory for actual readers.
'
'szReaders = (char *) malloc(sizeof(char) * readers);
'int res = SCardListReaders(hSCardContext, szGroups, szReaders , &readers);
'
'Now you have list of readers connected.
'
'You can connect to a reader like so
'
'LPSCARDHANDLE hSCard = NULL;
'long activeProtocols = 0;
'int ret = SCardConnect(hSCardContext, myReader, SCARD_SHARE_EXCLUSIVE, SCARD_PROTOCOL_TX, &hSCard, &activeProtocols);
'// .. handle errors
  
End Sub

'C++ Code to handle unplugged card readers
'bool test()
'{
'
'    DWORD dwReaders;
'    LPSTR szReaders = NULL;
'    SCARDCONTEXT hContext;
'    bool bRunning = true;
'
'    std::vector<const char*> cards;
'
'    LONG status = SCardEstablishContext(SCARD_SCOPE_USER, NULL, NULL, &hContext);
'    if( status != SCARD_S_SUCCESS ) {
'        return false;
'    }
'    dwReaders = SCARD_AUTOALLOCATE;
'    if( SCardListReadersA(hContext, NULL, (LPSTR)&szReaders, &dwReaders) == SCARD_S_SUCCESS ) {
'        LPSTR reader = szReaders;
'        while (reader != NULL && *reader != '\0') {
'            std::cout << "Reader name: '" << reader << "'" << std::endl;
'            cards.push_back( reader );
'            reader += strlen(reader)+1;
'        }
'        LPSCARD_READERSTATEA lpState = new SCARD_READERSTATEA[cards.size()];
'        for( size_t n = 0; n < cards.size(); ++n ) {
'            memset( lpState + n, 0, sizeof(SCARD_READERSTATEA) );
'            lpState[n].szReader = cards[n];
'        }
'
'        do {
'            status = SCardGetStatusChangeA( hContext, 500, lpState, cards.size() );
'            Switch (Status)
'            {
'            Case SCARD_S_SUCCESS:
'            Case SCARD_E_TIMEOUT:
'                for( size_t n = 0; n < cards.size(); ++n ) {
'                    if( lpState[n].dwEventState & SCARD_STATE_PRESENT) {
'                        std::cout << "'" << lpState[n].szReader << "' present" << std::endl;
'                    } else {
'                        std::cout << "'" << lpState[n].szReader << "' not present" << std::endl;
'                    }
'                }
'                break;
'default:
'                std::cout << "Other result: " << status << std::endl;
'                break;
'            }
'            Sleep( 1000 );  // do not spam too bad
'        } while( bRunning );
'        // only do this after being done with the strings, or handle the names another way!
'        SCardFreeMemory( hContext, szReaders );
'    }
'    SCardReleaseContext( hContext );
'    return true;
'}

'https://www.codeproject.com/Articles/16653/A-Smart-Card-Framework-for-NET
'https://stackoverflow.com/questions/7555281/smartcard-authentication-for-net
'https://social.msdn.microsoft.com/Forums/en-US/0a4b4873-b1c2-4087-8ad0-93d75dd46e3b/smart-card-certificates?forum=wcf
