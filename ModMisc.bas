Attribute VB_Name = "ModMisc"
Option Explicit

Public Const LISTENPORT = 48557
Public Const POLL_INTERVAL = 60000 * 5
Public Const TIME_OUT = 5
Public Const FileAlias = "aliases.dat"

Public Type HostList
    hostname    As String
    ipaddress   As String
    fullname    As String
    frRefer As frmChat
End Type

Public Enum sockstat
  ready = 1
  connecting = 2
  listening = 3
  moveport = 4
  connected = 5
  disconnecting = 6
End Enum

'For FindFolder--------------------
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long


                Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
                    'These constants are to be set to the ul
                    '     Flags property in the BROWSEINFO type de
                    '     pending of what result you want
                    Const BIF_RETURNONLYFSDIRS = &H1 'Allows you to browse For system folders only.
                    Const BIF_DONTGOBELOWDOMAIN = &H2 'Using this value forces the _
                    user to stay within the domain level of the _
                    Network Neighborhhood
                    Const BIF_STATUSTEXT = &H4 'Displays a statusbar on the selection dialog
                    Const BIF_RETURNFSANCESTORS = &H8 'Returns file system ancestor only
                    Const BIF_BROWSEFORCOMPUTER = &H1000 'Allows you to browse for a computer
                    Const BIF_BROWSEFORPRINTER = &H2000 'Allows you to browse the Printers folder


                Type BROWSEINFO
                    hOwner As Long
                    pidlRoot As Long
                    pszDisplayName As String
                    lpszTitle As String
                    ulFlags As Long
                    lpfn As Long
                    lParam As Long
                    iImage As Long
                    End Type
'End Find Folder-------------------

Private Declare Function StrFormatByteSize Lib "shlwapi" Alias "StrFormatByteSizeA" (ByVal dw As Long, ByVal pszBuf As String, ByRef cchBuf As Long) As String
Public Declare Function FlashWindow Lib "user32" _
  (ByVal hwnd As Long, ByVal bInvert As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Const VER_PLATFORM_WIN32_NT = 2
Private Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128
End Type

Public arrHostList() As HostList
Public arrFriendList() As HostList
Public HostsCount As Byte
Public FriendCount As Byte
Public PollingTime As Long
Public RecvFilePath As String
Public PopLoad As Byte
Public PopLevel As Byte

Public Sub DeleteFromHostList(ip_address As String)
Dim i, j As Byte
    
For i = 1 To HostsCount
    If arrHostList(i).ipaddress = ip_address Then
      For j = i To HostsCount - 1
        arrHostList(j) = arrHostList(j + 1)
      Next
      Exit For
    End If
Next
If HostsCount > 0 Then HostsCount = HostsCount - 1
End Sub

Public Function IsOnFriendList(ByVal ip_address As String) As Boolean
Dim i As Byte
IsOnFriendList = False
For i = 1 To FriendCount
  If arrFriendList(i).ipaddress = ip_address Then
    IsOnFriendList = True
    Exit For
  End If
Next
End Function

Public Function IsOnHostList(ByVal ip_address As String) As Boolean
Dim i As Byte
IsOnHostList = False
For i = 1 To HostsCount
  If arrHostList(i).ipaddress = ip_address Then
    IsOnHostList = True
    Exit For
  End If
Next
End Function

Public Sub AddToFriendList(ByVal ip_address As String)
If Not IsOnFriendList(ip_address) And IsOnHostList(ip_address) Then
    Dim i
    For i = 1 To HostsCount
        If arrHostList(i).ipaddress = ip_address Then Exit For
    Next i
    FriendCount = FriendCount + 1
    arrFriendList(FriendCount) = arrHostList(i)
    Set arrFriendList(FriendCount).frRefer = Nothing '***This is for next implementation
    frmMain.ShowList
End If
End Sub

Public Sub AddToHostList(ByVal ip_address As String, ByVal host_name As String)
If ip_address <> "" And host_name <> "" Then
    HostsCount = HostsCount + 1
    arrHostList(HostsCount).hostname = host_name
    arrHostList(HostsCount).ipaddress = ip_address
    arrHostList(HostsCount).fullname = GetFriendlyName(ip_address)
    If arrHostList(HostsCount).fullname = "" Then arrHostList(HostsCount).fullname = host_name
    Set arrHostList(HostsCount).frRefer = Nothing '***This is for next implementation
End If
End Sub
Public Sub DeleteFromFriendList(ByVal ip_address As String)
If IsOnFriendList(ip_address) Then
    Dim i, j As Byte
    For i = 1 To FriendCount
        If arrFriendList(i).ipaddress = ip_address Then
            For j = i To FriendCount - 1
                arrFriendList(j) = arrFriendList(j + 1)
            Next
            Exit For
        End If
    Next
    If FriendCount > 0 Then FriendCount = FriendCount - 1
    frmMain.ShowList
End If
End Sub

Public Function IsNetworkInstalled() As Boolean
    Const SM_NETWORK = 63
    IsNetworkInstalled = GetSystemMetrics(SM_NETWORK)
End Function

Public Function IsWindowsNTOr2000() As Boolean

    Dim lRC As Long
    Dim typOSInfo As OSVERSIONINFO
    
    typOSInfo.dwOSVersionInfoSize = Len(typOSInfo)
    lRC = GetVersionEx(typOSInfo)
    IsWindowsNTOr2000 = (typOSInfo.dwPlatformId = VER_PLATFORM_WIN32_NT) Or (typOSInfo.dwMajorVersion = 5)

End Function

Public Function GetFriendlyName(ByVal sIp As String)
Dim i, j, f
Dim tIp As String, tname As String
GetFriendlyName = ""
f = FreeFile
On Error GoTo EXIT_FILE
Open App.Path & "\" & FileAlias For Input As #f
Do While Not EOF(f)
    Input #f, tIp, tname
    If sIp = tIp Then
        GetFriendlyName = tname
        Exit Do
    End If
Loop
EXIT_FILE:
Close #f
End Function

Public Function GetIpFromList(ByVal host_name As String) As String
Dim i

For i = 1 To FriendCount
    If arrFriendList(i).hostname = host_name Then
        GetIpFromList = arrFriendList(i).ipaddress
        Exit For
    End If
Next
    
End Function

Public Function GetHostFromList(ByVal ip_address As String) As String
Dim i
Dim bFound As Boolean

bFound = False
For i = 1 To HostsCount
    If arrHostList(i).ipaddress = ip_address Then
        GetHostFromList = arrHostList(i).hostname
        bFound = False
        Exit For
    End If
Next
    
If Not bFound Then
    AddToHostList ip_address, GetHostNameFromIP(ip_address)
    AddToFriendList ip_address
    frmMain.ShowList
End If
End Function
Public Function GetFriendlyNameFromList(ByVal ip_address As String) As String
Dim i
GetFriendlyNameFromList = ""
For i = 1 To FriendCount
    If arrFriendList(i).ipaddress = ip_address Then
        GetFriendlyNameFromList = arrFriendList(i).fullname
        Exit For
    End If
Next i
End Function

Public Function isValidIp(ByVal sIp As String) As Boolean
Dim stemp, i
isValidIp = False
If InStr(sIp, ".") > 0 Then
    stemp = Split(sIp, ".")
    Debug.Print UBound(stemp)
    If UBound(stemp) = 3 Then
        For i = 0 To UBound(stemp)
            If IsNumeric(stemp(i)) Then
                isValidIp = True
            Else
                isValidIp = False
                Exit For
            End If
        Next
    End If
End If
End Function

Public Function FormatKB(ByVal Amount As Long) As String
    Dim Buffer As String
    Dim result As String
    Buffer = Space$(255)
    result = StrFormatByteSize(Amount, Buffer, Len(Buffer))
    If InStr(result, vbNullChar) > 1 Then FormatKB = Left$(result, InStr(result, vbNullChar) - 1)
End Function

Public Function GetFolder(Optional title As String, Optional hwnd) As String
Dim bi As BROWSEINFO
Dim pidl As Long
Dim Folder As String
Folder = String$(255, Chr$(0))

With bi
    If IsNumeric(hwnd) Then .hOwner = hwnd
    .ulFlags = BIF_RETURNONLYFSDIRS
    .pidlRoot = 0

    If Not IsMissing(title) Then
        .lpszTitle = title
    Else
        .lpszTitle = "Select a Folder" & Chr$(0)
    End If
End With
pidl = SHBrowseForFolder(bi)


If SHGetPathFromIDList(ByVal pidl, ByVal Folder) Then
    GetFolder = Left(Folder, InStr(Folder, Chr$(0)) - 1)
Else
    GetFolder = ""
End If
End Function
