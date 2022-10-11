Attribute VB_Name = "mdlHTTP"
' After TCP connection is established it should be handeled according to HTTP protocol.
'
' There are mainly HTTP things in this module.
'
' Bojan Jurca, 10.10.2022

Option Explicit

Public Sub handleTcpConnection(socket As Long, clientIP As String)
    ' at this point we have a stream of incoming HTTP requests on the same TCP connection
    ' each HTTP request ends with \r\n\r\n, this is how we know that we have got the whole request

    Dim httpRequest As String
    httpRequest = "Connection: keep-alive"
    ' keep reading HTTP requests from the same TCP connection as long as the client keeps sending 'keep-alive'
    ' but let's make additional condition that each TCP connection is not opened for more than 3 seconds,
    ' this is a single threaded, blocking TCP server so while a TCP connection is alive no other
    ' TCP connection (from another client for example) can be established
    Dim startTime As Date
    startTime = Now
    Do While InStr(httpRequest, "Connection: keep-alive") > 0 And DateDiff("s", startTime, Now) <= 3
        
        ' read the whole HTTP request - many recv calls may be needed until the ned of HTTP request is reched
        Do
            httpRequest = ""
            Dim requestBuffer As String * maxRecvBufferSize ' reserve enough memory for read buffer
            Dim bytesRecv As Long
            bytesRecv = recv(socket, requestBuffer, maxRecvBufferSize, 0)
            If bytesRecv <= 0 Then ' transmission error or the client closed the connection
                ' most likely the TCP connection has been closed by the client
                ' Debug.Print "recv ERROR " & bytesRecv
                GoTo lblCloseConnectionSocket
            Else
                Debug.Print "      recived " & bytesRecv & " bytes"
                httpRequest = httpRequest & Left$(requestBuffer, bytesRecv)
            End If
        Loop While InStr(httpRequest, vbCrLf & vbCrLf) = 0 ' keep reading until end of HTTP request arrives
        ' Debug.Print httpRequest
        handleHttpRequest socket, httpRequest, clientIP
    
    Loop
    
lblCloseConnectionSocket:
    If closesocket(socket) = SOCKET_ERROR Then
        Debug.Print "closesocket ERROR"
    Else
        Debug.Print "closesocket OK"
    End If
End Sub

Private Sub handleHttpRequest(socket As Long, httpRequest As String, clientIP As String)
    ' at this point we have a single and whole HTTP request and it already can be answered with HTTP reply,
    ' for example:
    '
    '   Dim httpReply As String
    '   httpReply = "HTTP/1.0 404 Not found" & vbCrLf & "Content-Length:10" & vbCrLf & vbCrLf & "Not found."
    '   Dim bytesSent As Long
    '   bytesSent = send(socket, httpReply, Len(httpReply), 0)
    '   If bytesSent <= 0 Then
    '       Debug.Print "send ERROR " & bytesSent
    '   Else
    '       Debug.Print "      sent " & bytesSent & " bytes"
    '   End If
    '
    ' but let us parse it a somewhat to make job easyer for later use. In general HTTP request would look something like:
    '
    '               ---------------  GET / HTTP/1.1
    '              |                 Host: 10.18.1.200
    '              |                 Connection: keep-alive    field value
    '              |                 Cache-Control: max-age=0   |
    '     request  |    field name - Upgrade-Insecure-Requests: 1
    '              |                 User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.45 Safari/537.36
    '              |                 Accept: text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng;q=0.8,application/signed-exchange;v=b3;q=0.9
    '              |                 Accept-Encoding: gzip, deflate
    '              |                 Accept-Language: sl-SI,sl;q=0.9,en-GB;q=0.8,en;q=0.7
    '              |                 Cookie: refreshCounter=1
    '              |                             |          |
    '              |                        cookie name   cookie value
    '               ---------------
    '
    ' and our job is to uderstand what the client wanted and send back a reply that would look something like:
    '
    '                                                Status
    '                                                   |         cookie name, value, path and expiration
    '               --------  ---------------  HTTP/1.1 200 OK      |     |    |         |
    '              |         |                 Set-Cookie: refreshCounter=2; Path=/; Expires=Thu, 09 Dec 2021 19:07:04 GMT
    '        reply |  header |    field name - Content-Type: text/html
    '              |         |                 Content-Length: 96   |
    '              |          ---------------                      field value
    '               - content ---------------  <HTML>Web cookies<br><br>This page has been refreshed 2 times. Click refresh to see more.</HTML>
    '
    
    Debug.Print "-------------------------------------------"
    Debug.Print httpRequest
    Debug.Print "-------------------------------------------"
    
    Dim httpReply As String
    Dim httpReplyHeader As String
    Dim httpReplyContent As String
    
    On Error GoTo lblResourceError
    Dim hcp As New clsHttpCnnParams
    hcp.socket = socket
    hcp.clientIP = clientIP
    hcp.httpRequest = httpRequest
    hcp.httpReplyStatus = "200 OK" ' by default
    
    httpReplyContent = provideHttpReplyContent(httpRequest, hcp)
    Dim bytesSent As Long
    
    If httpReplyContent = "" Then ' provideHttpReplyContent did not provide the content - try to pass a file
    
        ' let's try to locate a file. If it does't exist send HTTP reply 404 to the client
        Dim fileName As String
        Dim i As Integer
        i = InStr(httpRequest, "GET ")
        If (i > 0) Then
            i = i + 5
            Dim j As Integer
            j = InStr(i, httpRequest, " ")
            If j > 0 Then
                fileName = App.path & "\html\" & Mid$(httpRequest, i, j - i)
            End If
        End If
        If fileName > "" Then
        
            ' set Content-Type if not set by provideHttpReplyContent function
            
            If InStr(1, hcp.httpReplyHeader, "Content-Type", vbTextCompare) <= 0 Then
                If Right(fileName, 4) = ".bmp" Then hcp.setHttpReplyHeaderField "Content-Type", "image/bmp"
                If Right(fileName, 4) = ".css" Then hcp.setHttpReplyHeaderField "Content-Type", "text/css"
                If Right(fileName, 4) = ".csv" Then hcp.setHttpReplyHeaderField "Content-Type", "text/csv"
                If Right(fileName, 4) = ".gif" Then hcp.setHttpReplyHeaderField "Content-Type", "image/gif"
                If Right(fileName, 4) = ".htm" Then hcp.setHttpReplyHeaderField "Content-Type", "text/html"
                If Right(fileName, 5) = ".html" Then hcp.setHttpReplyHeaderField "Content-Type", "text/html"
                If Right(fileName, 4) = ".jpg" Then hcp.setHttpReplyHeaderField "Content-Type", "image/jpeg"
                If Right(fileName, 5) = ".jpeg" Then hcp.setHttpReplyHeaderField "Content-Type", "image/jpeg"
                If Right(fileName, 3) = ".js" Then hcp.setHttpReplyHeaderField "Content-Type", "text/javascript"
                If Right(fileName, 5) = ".json" Then hcp.setHttpReplyHeaderField "Content-Type", "application/json"
                If Right(fileName, 5) = ".mpeg" Then hcp.setHttpReplyHeaderField "Content-Type", "video/mpeg"
                If Right(fileName, 4) = ".pdf" Then hcp.setHttpReplyHeaderField "Content-Type", "application/pdf"
                If Right(fileName, 4) = ".png" Then hcp.setHttpReplyHeaderField "Content-Type", "image/png"
                If Right(fileName, 5) = ".tiff" Then hcp.setHttpReplyHeaderField "Content-Type", "image/tiff"
                If Right(fileName, 4) = ".txt" Then hcp.setHttpReplyHeaderField "Content-Type", "text/plain"
                ' ... add more if needed but Contet-Type can often be omitted without problems ...
            End If
            
            On Error GoTo lblFileError
        
            Dim bytesToRead As Long
            bytesToRead = FileLen(fileName)
            If bytesToRead = 0 Then GoTo lblFileError
        
            Dim bytesRead
        
            Dim fileNum As Integer
            fileNum = FreeFile
            Open fileName For Binary Access Read As fileNum
            
            httpReplyHeader = "HTTP/1.1 " & hcp.httpReplyStatus & vbCrLf & _
                              hcp.httpReplyHeader & _
                              "Content-Length: " & bytesToRead & vbCrLf & vbCrLf
        
            Debug.Print "==========================================="
            Debug.Print httpReplyHeader
            Debug.Print fileName
                
            Dim readBuffer As String
            Dim readBufferSize As Long
            If Len(httpReplyHeader) + bytesToRead < maxRecvBufferSize Then ' calculate required buffer size
                readBufferSize = Len(httpReplyHeader) + bytesToRead
            Else
                readBufferSize = maxRecvBufferSize - Len(httpReplyHeader)
            End If
            readBuffer = String$(readBufferSize, Chr(0)) ' reserver buffer space
            
            ' read and send the first chunk of the file
            Get fileNum, , readBuffer
            bytesRead = readBufferSize
            readBuffer = httpReplyHeader & readBuffer ' add the header (only the first time)
            bytesSent = send(socket, readBuffer, Len(httpReplyHeader) + readBufferSize, 0)
            If bytesSent <= 0 Then ' transmission error or the client closed the connection
                Debug.Print "send ERROR " & bytesSent
                Close fileNum
                Exit Sub
            Else
                Debug.Print "      sent " & bytesSent & " bytes"
            End If
    
            Do While bytesRead < bytesToRead
                ' read and send chunks of file untill all the file is sent
                If bytesToRead - bytesRead > maxRecvBufferSize Then ' calculate required buffer size
                    readBufferSize = maxRecvBufferSize
                Else
                    readBufferSize = bytesToRead - bytesRead
                End If
                readBuffer = String$(readBufferSize, Chr(0)) ' reserver buffer space
                
                ' read and send the next chunk of the file
                Get fileNum, , readBuffer
                bytesRead = bytesRead + readBufferSize
                bytesSent = send(socket, readBuffer, readBufferSize, 0)
                If bytesSent <= 0 Then ' transmission error or the client closed the connection
                    Debug.Print "send ERROR " & bytesSent
                    Close fileNum
                    Exit Sub
                Else
                    Debug.Print "      sent " & bytesSent & " bytes"
                End If
            Loop
            Close fileNum
            Debug.Print "==========================================="
        End If

        If fileName = "" Then ' failed to open the file (different reasons)

lblFileError:
            
            httpReply = "HTTP/1.0 404 Not found" & vbCrLf & "Content-Length:10" & vbCrLf & vbCrLf & "Not found."
            bytesSent = send(socket, httpReply, Len(httpReply), 0)
            
            Debug.Print "==========================================="
            Debug.Print httpReply
            Debug.Print "==========================================="
                                
            bytesSent = send(socket, httpReply, Len(httpReply), 0)
            If bytesSent <= 0 Then ' transmission error or the client closed the connection
                Debug.Print "send ERROR " & bytesSent
            Else
                Debug.Print "      sent " & bytesSent & " bytes"
            End If
        End If
        
    Else ' provideHttpReplyContent provided the content
    
        ' try to guess Content-Type if not set by provideHttpReplyContent function
        
        If InStr(1, hcp.httpReplyHeader, "Content-Type", vbTextCompare) <= 0 Then
            If InStr(1, httpReplyContent, "<HTML>", vbTextCompare) Then
                hcp.setHttpReplyHeaderField "Content-Type", "text/html"
            Else
                If InStr(1, httpReplyContent, "{", vbTextCompare) Then
                    hcp.setHttpReplyHeaderField "Content-Type", "application/json"
                Else
                    hcp.setHttpReplyHeaderField "Content-Type", "text/plain"
                End If
            End If
        End If
        
        ' let's add the header
        httpReply = "HTTP/1.1 " & hcp.httpReplyStatus & vbCrLf & _
                    hcp.httpReplyHeader & _
                    "Content-Length: " & Len(httpReplyContent) & _
                    vbCrLf & vbCrLf & _
                    httpReplyContent
                    
        Debug.Print "==========================================="
        Debug.Print httpReply
        Debug.Print "==========================================="
                            
        bytesSent = send(socket, httpReply, Len(httpReply), 0)
        If bytesSent <= 0 Then ' transmission error or the client closed the connection
            Debug.Print "send ERROR " & bytesSent
        Else
            Debug.Print "      sent " & bytesSent & " bytes"
        End If
    End If
    Exit Sub
    
lblResourceError:

    httpReply = "HTTP/1.0 503 Service unavailable" & vbCrLf & "Content-Length:39" & vbCrLf & vbCrLf & "HTTP server is not available right now."
    bytesSent = send(socket, httpReply, Len(httpReply), 0)
    
    Debug.Print "==========================================="
    Debug.Print httpReply
    Debug.Print "==========================================="
                        
    bytesSent = send(socket, httpReply, Len(httpReply), 0)
    If bytesSent <= 0 Then ' transmission error or the client closed the connection
        Debug.Print "send ERROR " & bytesSent
    Else
        Debug.Print "      sent " & bytesSent & " bytes"
    End If

End Sub

