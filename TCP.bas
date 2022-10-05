Attribute VB_Name = "mdlTCP"
' After a surprising discovery that VB6 is still running on Windows 11 I decided to upgrade
' 20 years old VB TCP/IP code to vb6WebServer.
'
' There are mainly TCP things in this module, a good source on Win32 APIs is here:
' http://www.jasinskionline.com/windowsapi/ref/i/ioctlsocket.html
'
' Bojan Jurca, 1.10.2022

Option Explicit


' ----- SETTINGS -----

' web server settings
Public Const vb6WebServerPort As Integer = 80
Public Const vb6WebServerIP As String = "127.0.0.1" ' your conputer's IP or 127.0.0.1 for local loopback for example

' define max buffer size for received data
Public Const maxRecvBufferSize = 1464 ' or check what is your optional MTU size (-28 bytes)


' ----- WIN32 API -----

Private Const WSADESCRIPTION_LEN As Integer = 256
Private Const WSADESCRIPTION_LEN_AND_1 As Integer = WSADESCRIPTION_LEN + 1
Private Const WSASYS_STATUS_LEN As Integer = 128
Private Const WSASYS_STATUS_LEN_AND_1 As Integer = WSASYS_STATUS_LEN + 1
Private Type WSADATA
    wVersion As Integer
    wHighVersion As Integer
    szDescription As String * WSADESCRIPTION_LEN_AND_1
    szSystemStatus As String * WSASYS_STATUS_LEN_AND_1
    iMaxSockets As Integer
    iMaxUdpDg As Long
    lpVendorInfo As String
End Type
Private Declare Function WSAStartup Lib "ws2_32.dll" (ByVal wVersionRequested As Integer, ByRef lpWSAData As WSADATA) As Integer
Private Declare Function WSACleanup Lib "ws2_32.dll" () As Integer

Private Const AF_INET As Integer = 2
Private Const SOCK_STREAM  As Integer = 1
Private Const SOCK_DGRAM As Integer = 2
Private Const INVALID_SOCKET As Integer = -1
Public Const SOCKET_ERROR As Integer = -1
Private Declare Function socket Lib "ws2_32.dll" (ByVal af As Integer, ByVal tpe As Integer, ByVal protocol As Integer) As Long
Private Const SD_RECEIVE As Integer = 0
Private Const SD_SEND As Integer = 1
Private Const SD_BOTH As Integer = 2
Private Declare Function shutdown Lib "ws2_32.dll" (ByVal socket As Long, ByVal how As Integer) As Integer
Public Declare Function closesocket Lib "ws2_32.dll" (ByVal socket As Long) As Integer
 
Private Const INADDR_NONE As Long = &HFFFFFFFF
Private Declare Function inet_addr Lib "ws2_32.dll" (ByVal cp As String) As Long
 
Private Type SOCKADDR_IN
    sin_family As Integer
    sin_port As Integer
    sin_addr As Long
    sin_zero As String * 8
End Type
Private Declare Function htons Lib "ws2_32.dll" (ByVal Hostshort As Integer) As Integer

Const SOL_SOCKET = 65535      ' Options for socket level.
Const IPPROTO_TCP = 6         ' Protocol constant for TCP.
' option flags per socket
Const SO_DEBUG = &H1&         ' Turn on debugging info recording
Const SO_ACCEPTCONN = &H2&    ' Socket has had listen() - READ-ONLY.
Const SO_REUSEADDR = &H4&     ' Allow local address reuse.
Const SO_KEEPALIVE = &H8&     ' Keep connections alive.
Const SO_DONTROUTE = &H10&    ' Just use interface addresses.
Const SO_BROADCAST = &H20&    ' Permit sending of broadcast msgs.
Const SO_USELOOPBACK = &H40&  ' Bypass hardware when possible.
Const SO_LINGER = &H80&       ' Linger on close if data present.
Const SO_OOBINLINE = &H100&   ' Leave received OOB data in line.
Const SO_DONTLINGER = Not SO_LINGER
Const SO_EXCLUSIVEADDRUSE = Not SO_REUSEADDR ' Disallow local address reuse.
' additional options
Const SO_SNDBUF = &H1001&     ' Send buffer size.
Const SO_RCVBUF = &H1002&     ' Receive buffer size.
Const SO_ERROR = &H1007&      ' Get error status and clear.
Const SO_TYPE = &H1008&       ' Get socket type - READ-ONLY.

Private Declare Function setsockopt Lib "wsock32.dll" (ByVal socket As Long, ByVal level As Long, ByVal optname As Long, optval As Any, ByVal optlen As Long) As Long
Private Declare Function bind Lib "ws2_32.dll" (ByVal socket As Long, ByRef name As SOCKADDR_IN, ByVal namelen As Integer) As Integer
Private Declare Function listen Lib "ws2_32.dll" (ByVal socket As Long, ByVal backlog As Integer) As Integer
' ioctlsocket constants
Const FIONBIO = &H8004667E
Const FIONREAD = &H4004667F
Const SIOCATMARK = &H40047307
Declare Function ioctlsocket Lib "wsock32.dll" (ByVal S As Long, ByVal cmd As Long, argp As Long) As Long
Private Declare Function accept Lib "ws2_32.dll" (ByVal socket As Long, ByRef addr As SOCKADDR_IN, ByRef addrlen As Long) As Long
Private Declare Function inet_ntoa Lib "ws2_32.dll" (ByVal inn As Long) As Long
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Private Declare Function ntohs Lib "ws2_32.dll" (ByVal netshort As Integer) As Integer

Public Declare Function send Lib "ws2_32.dll" (ByVal socket As Long, ByVal buffer As String, ByVal BytesToSend As Integer, ByVal flags As Integer) As Integer
Public Declare Function recv Lib "ws2_32.dll" (ByVal socket As Long, ByVal buffer As String, ByVal BytesToRecv As Integer, ByVal flags As Integer) As Integer

Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)


' ----- GLOBAL VARIABLES -----

Global requestToStopVb6WebServer As Boolean


Private Sub Main()
    
    ' ----- INITIALIZE -----
   
    Dim w As WSADATA
    Dim listeningSocket As Long, connectionSocket As Long
    Dim i As SOCKADDR_IN, r As SOCKADDR_IN ' interface sin, rmote sin
    
    ' initialize Winsock
    If WSAStartup(&H101, w) Then
        frmVb6WebServer.errorMessage = "WSASatrtup ERROR"
        Debug.Print frmVb6WebServer.errorMessage
        Exit Sub
    Else
        Debug.Print "WSAStartup OK"
    End If
        
    ' create listening socket
    i.sin_family = AF_INET
    i.sin_port = htons(vb6WebServerPort)
    i.sin_addr = inet_addr(vb6WebServerIP)
    listeningSocket = socket(AF_INET, SOCK_STREAM, 0)
    If listeningSocket = INVALID_SOCKET Then
        frmVb6WebServer.errorMessage = "(listening) socket ERROR"
        Debug.Print frmVb6WebServer.errorMessage
        GoTo lblCleanUp
    Else
        Debug.Print "(listening) socket OK"
    End If
    
    ' make address reusable - so we won't have to wait a few minutes in case server will be restarted
    Dim flag As Integer
    flag = 1
    If (setsockopt(listeningSocket, SOL_SOCKET, SO_REUSEADDR, flag, Len(flag)) = SOCKET_ERROR) Then
        frmVb6WebServer.errorMessage = "setsockoption ERROR"
        Debug.Print frmVb6WebServer.errorMessage
        ' continue anyway, it is not critical error
    Else
        Debug.Print "setsockoption OK"
    End If
     
    ' bind listening socket to IP address and port number
    If bind(listeningSocket, i, Len(i)) = SOCKET_ERROR Then
        frmVb6WebServer.errorMessage = "bind ERROR"
        Debug.Print frmVb6WebServer.errorMessage
        GoTo lblCloseListeningSocket
    Else
        Debug.Print "bind OK"
    End If
    
    ' make socket non-blocking so that aceept () won't block and the form can interact with the user meanwhile
    If ioctlsocket(listeningSocket, FIONBIO, 1) = SOCKET_ERROR Then ' instead of fcntl (ls, F_SETFL, O_NONBLOCK)
        frmVb6WebServer.errorMessage = "ioctlsocket ERROR"
        Debug.Print frmVb6WebServer.errorMessage
        GoTo lblCloseListeningSocket
    Else
        Debug.Print "ioctlsocket OK"
    End If
    
    ' start listening on listening socket
    If listen(listeningSocket, 1) = SOCKET_ERROR Then
        frmVb6WebServer.errorMessage = "listen ERROR"
        Debug.Print frmVb6WebServer.errorMessage
        GoTo lblCloseListeningSocket
    Else
        Debug.Print "listen OK"
    End If
    
    ' ----- LISTEN FOR INCOMING CONNECTIONS -----
    
    Debug.Print "vb6WebServer started on " & vb6WebServerIP & ":" & vb6WebServerPort
    frmVb6WebServer.Show
    
    Do While (True)
    
        Debug.Print "      waiting for a connection ..."
        connectionSocket = INVALID_SOCKET
        Do While connectionSocket = INVALID_SOCKET And Not requestToStopVb6WebServer
            Sleep 10 ' sleep 10 ms so we don't use processor time while waiting
            DoEvents ' give frmVb6WebServer a chace to handle its events
            Dim rLen As Long
            rLen = Len(r)
            connectionSocket = accept(listeningSocket, r, rLen)
        Loop
        If requestToStopVb6WebServer Then
            GoTo lblCloseListeningSocket
        Else
        
            Dim clientIP As String
            clientIP = String(46, Chr(0))
            lstrcpy clientIP, inet_ntoa(r.sin_addr)
            clientIP = Left$(clientIP, InStr(clientIP, Chr(0)) - 1)
            Debug.Print "      accepted connection from " & clientIP ' & " on port " & ntohs(r.sin_port) & " ..."
            
            ' in threaded environment we would start serving a new TCP connection in another thread
            ' but VB 6 is not handeling multithreading very well so we'll just continue with the same
            ' thread thus blocking new incoming connection meanwhile
            
                ' handle new TCP Connection according to HTTP protocol
                handleTcpConnection connectionSocket, clientIP
                ' after HandleConnection we can assume that connectionSocket is already closed
                
        End If
    Loop
        
        
    ' ----- CLEAN UP -----
    
lblCloseListeningSocket:
    If closesocket(listeningSocket) = SOCKET_ERROR Then
        Debug.Print "closesocket ERROR"
    Else
        Debug.Print "closesocket OK"
    End If
        
lblCleanUp:

    If WSACleanup() <> 0 Then
        Debug.Print "WSAClenaup ERROR"
    Else
        Debug.Print "WSAClenaup OK"
    End If
    
    Debug.Print "vb6WebServer stopped"
End Sub
