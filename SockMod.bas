Attribute VB_Name = "SockMod"
Option Explicit
Private Const WS_VERSION_REQD = &H101
Private Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
Private Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&
Private Const MIN_SOCKETS_REQD = 1
Private Const INADDR_NONE = &HFFFF
Private Const socket_Error = -1
Private Const AF_INET = 2
Private Const WSABASEERR = 10000
Private Const WSAEFAULT = (WSABASEERR + 14)
Private Const WSAEINVAL = (WSABASEERR + 22)
Private Const WSAEINPROGRESS = (WSABASEERR + 50)
Private Const WSAENETDOWN = (WSABASEERR + 50)
Private Const WSASYSNOTREADY = (WSABASEERR + 91)
Private Const WSAVERNOTSUPPORTED = (WSABASEERR + 92)
Private Const WSANOTINITIALISED = (WSABASEERR + 93)
Private Const WSAHOST_NOT_FOUND = 11001
Private Const MAX_WSADescription = 256
Private Const MAX_WSASYSStatus = 128
Private Const WSATRY_AGAIN = 11002
Private Const WSANO_RECOVERY = 11003
Private Const WSANO_DATA = 11004
Private Const PING_TIMEOUT As Long = 50
Private Const IP_SUCCESS As Long = 0

Private Const RESOURCE_CONNECTED = &H1
Private Const RESOURCE_GLOBALNET = &H2
Private Const RESOURCE_REMEMBERED = &H3

Private Const RESOURCETYPE_ANY = &H0
Private Const RESOURCETYPE_DISK = &H1
Private Const RESOURCETYPE_PRINT = &H2
Private Const RESOURCETYPE_UNKNOWN = &HFFFF

Private Const RESOURCEUSAGE_CONNECTABLE = &H1
Private Const RESOURCEUSAGE_CONTAINER = &H2
Private Const RESOURCEUSAGE_RESERVED = &H80000000

Private Const GMEM_FIXED = &H0
Private Const GMEM_ZEROINIT = &H40
Private Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)

Private Type WSADATA
    wversion As Integer
    wHighVersion As Integer
    szDescription(0 To MAX_WSADescription) As Byte
    szSystemStatus(0 To MAX_WSASYSStatus) As Byte
    wMaxSockets As Long
    wMaxUDPDG As Long
    dwVendorInfo As Long
End Type

Public Type HOSTENT
    hName     As Long
    hAliases  As Long
    hAddrType As Integer
    hLength   As Integer
    hAddrList As Long
End Type

Public Type NETRESOURCE
   dwScope As Long
   dwType As Long
   dwDisplayType As Long
   dwUsage As Long
   lpLocalName As Long
   lpRemoteName As Long
   lpComment As Long
   lpProvider As Long
End Type

Private Declare Function WNetOpenEnum Lib "mpr.dll" Alias "WNetOpenEnumA" (ByVal dwScope As Long, ByVal dwType As Long, _
  ByVal dwUsage As Long, lpNetResource As Any, lphEnum As Long) As Long
Private Declare Function WNetEnumResource Lib "mpr.dll" Alias "WNetEnumResourceA" (ByVal hEnum As Long, lpcCount As Long, _
  ByVal lpBuffer As Long, lpBufferSize As Long) As Long
Private Declare Function WNetCloseEnum Lib "mpr.dll" (ByVal hEnum As Long) As Long
Private Declare Function WSAStartup Lib "ws2_32.dll" (ByVal wVR As Long, lpWSAD As WSADATA) As Long
Private Declare Function WSACleanup Lib "ws2_32.dll" () As Long
Private Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
Private Declare Function inet_ntoa Lib "WSOCK32.DLL" (ByVal addr As Long) As Long
Private Declare Function inet_addr Lib "WSOCK32.DLL" (ByVal s As String) As Long
Private Declare Function gethostbyaddr Lib "WSOCK32.DLL" (haddr As Long, ByVal hnlen As Long, ByVal addrtype As Long) As Long
Private Declare Function gethostname Lib "WSOCK32.DLL" (ByVal hostname As String, ByVal HostLen As Long) As Long
Private Declare Function gethostbyname Lib "ws2_32.dll" (ByVal host_name As String) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function lstrcpyA Lib "kernel32" (ByVal RetVal As String, ByVal Ptr As Long) As Long
Private Declare Function lstrlenA Lib "kernel32" (ByVal Ptr As Any) As Long

Public strHostname As String
Public strIpAddress As String

Private Function SocketsInitialize() As Boolean
Dim WSAD As WSADATA
Dim lngReturnCode As Integer
Dim strLoByte As String
Dim strHiByte As String
Dim strBuffer As String
lngReturnCode = WSAStartup(WS_VERSION_REQD, WSAD)
If lngReturnCode <> 0 Then
    Err.Raise vbObjectError + 999, "SocketsInitialize", "Windows Sockets for 32 bit Windows environments is not successfully responding."
    Exit Function
End If
If LoByte(WSAD.wversion) < WS_VERSION_MAJOR Or (LoByte(WSAD.wversion) = WS_VERSION_MAJOR And HiByte(WSAD.wversion) < WS_VERSION_MINOR) Then
    strHiByte = Trim$(Str$(HiByte(WSAD.wversion)))
    strLoByte = Trim$(Str$(LoByte(WSAD.wversion)))
    strBuffer = "Windows Sockets Version " & strLoByte & "." & strHiByte
    strBuffer = strBuffer & " is not supported by Windows " & "Sockets for 32 bit Windows environments."
    Err.Raise vbObjectError + 999, "SocketsInitialize", strBuffer, vbExclamation
    Exit Function
End If
If WSAD.wMaxSockets < MIN_SOCKETS_REQD Then
    strBuffer = "This application requires a minimum of " & Trim$(Str$(MIN_SOCKETS_REQD)) & " supported sockets."
    Err.Raise vbObjectError + 999, "SocketsInitialize", strBuffer, vbExclamation
    Exit Function
End If
SocketsInitialize = True
End Function

Public Function GetComputerName() As String
Dim lngReturnCode As Long
Dim strHost As String
strHost = Space$(50)
lngReturnCode = gethostname(strHost, 50)
GetComputerName = Left$(strHost, InStr(strHost, Chr$(0)) - 1)
End Function

Public Function GetHostNameFromIP(ByVal strIpAddress As String) As String
Dim nbytes As Long
Dim ptrHosent As Long
Dim lookupIP As String
Dim lngIPAddress As Long
If SocketsInitialize() Then
    lngIPAddress = inet_addr(strIpAddress)
    DoEvents
    ptrHosent = gethostbyaddr(lngIPAddress, 4, AF_INET)
    If ptrHosent <> 0 Then
        CopyMemory ptrHosent, ByVal ptrHosent, 4
        nbytes = lstrlenA(ByVal ptrHosent)
        If nbytes > 0 Then
            lookupIP = Space$(nbytes)
            CopyMemory ByVal lookupIP, ByVal ptrHosent, nbytes
            GetHostNameFromIP = lookupIP
        End If
    Else
        GetHostNameFromIP = ""
    End If
    SocketsCleanup
End If
End Function

Public Function GetIPFromHostName(ByVal strHostname As String) As String
 Dim ptrHosent As Long
 Dim ptrName As Long
 Dim ptrAddress As Long
 Dim ptrIPAddress As Long
 Dim dwAddress As Long
 If SocketsInitialize() Then
     DoEvents
     ptrHosent = gethostbyname(strHostname & vbNullChar)
     If ptrHosent <> 0 Then
         ptrName = ptrHosent
         ptrAddress = ptrHosent + 12
         CopyMemory ptrAddress, ByVal ptrAddress, 4
         CopyMemory ptrIPAddress, ByVal ptrAddress, 4
         CopyMemory dwAddress, ByVal ptrIPAddress, 4
         GetIPFromHostName = PtrStr(inet_ntoa(dwAddress))
     End If
     SocketsCleanup
 End If
End Function

Private Sub SocketsCleanup()
Call WSACleanup
End Sub

Private Function PtrStr(ByVal lpszA As Long) As String
PtrStr = String$(lstrlenA(ByVal lpszA), 0)
Call lstrcpyA(ByVal PtrStr, ByVal lpszA)
End Function

Public Function HiByte(ByVal wParam As Integer)
HiByte = wParam \ &H100 And &HFF&
End Function

Public Function LoByte(ByVal wParam As Integer)
LoByte = wParam And &HFF&
End Function

Public Sub GetHostList(NR As NETRESOURCE)

Dim hEnum As Long, lpBuff As Long
Dim cbBuff As Long, cCount As Long
Dim P As Long, res As Long, i As Long


On Error GoTo ErrorHandler

'TODO: Enumerate Domain on network

'Setup the NETRESOURCE input structure.
cbBuff = 20000
cCount = &HFFFFFFFF
GlobalFree (cbBuff)
'Open a Net enumeration operation handle: hEnum.
DoEvents
res = WNetOpenEnum(RESOURCE_GLOBALNET, _
  RESOURCETYPE_ANY, 0, NR, hEnum)

If res = 0 Then

   'Create a buffer large enough for the results.
   '10000 bytes should be sufficient.
   lpBuff = GlobalAlloc(GPTR, cbBuff)
   'Call the enumeration function.
   DoEvents
   res = WNetEnumResource(hEnum, cCount, lpBuff, cbBuff)
   If res = 0 Then
      P = lpBuff
      'WNetEnumResource fills the buffer with an array of
      'NETRESOURCE structures. Walk through the list and print
      'each local and remote name.
      For i = 1 To cCount
         ' All we get back are the Global Network Containers --- Enumerate each of these
         CopyMemory NR, ByVal P, LenB(NR)

         'TODO: Enumerate Workstation On Domain

          'Setup the NETRESOURCE input structure.
          If Left(PtrStr(NR.lpRemoteName), 2) = "\\" Then
              Dim tHost As String, tIp As String
              'Modified this to trim "\" from computer name
              tHost = Replace(PtrStr(NR.lpRemoteName), "\", "")
              tIp = GetIPFromHostName(tHost)
              If tIp <> "" And Not (IsOnHostList(tIp)) Then
                    AddToHostList tIp, tHost
              End If
          Else
            'Do recursive enumeration to fix problem in NT environment
            GetHostList NR
          End If
         P = P + LenB(NR)
      Next i
    End If

ErrorHandler:
    On Error Resume Next
    If lpBuff <> 0 Then GlobalFree (lpBuff)
    WNetCloseEnum (hEnum) 'Close the enumeration
End If
End Sub

