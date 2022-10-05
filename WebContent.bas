Attribute VB_Name = "mdlWebContent"
' This module is supposed to be customized for each kind of use.
'
' Modify provideHttpReplyContent function according to your needs. The function is supposed to return
' content part of HTTP reply, HTTP header will be added before sending the reply back to the
' client. If provideHttpReplyContent returns "" then vb6WebServer will try to locate a (for example.html)
' file in application directory and pass its content to the client or it will just return a
' 404 reply if not found.
'
' See examples below
'
' Bojan Jurca, 1.10.2022

Option Explicit

' httpRequest - a complete HTTP request, normally you would build a HTTP reply based on this information
' clientIP - IP address of a client browser, just in case you would need it
' socket - intended for lower level programming, if you know what this is you probably already know how to use it, if you don't you better leave it be

Public Function provideHttpReplyContent(httpRequest As String, hcp As clsHttpCnnParams) As String

    ' 1 st example - generate the whole reply content
    If httpRequestIs(httpRequest, "GET / ") Then
        provideHttpReplyContent = "<HTML><HEAD>" & vbCrLf & _
                                  "   <link rel='shortcut icon' type='image/x-icon' sizes='192x192' href='/android-192x192.png'>" & vbCrLf & _
                                  "   <link rel='icon' type='image/png' sizes='192x192' href='/android-192x192.png'>" & vbCrLf & _
                                  "   <link rel='apple-touch-icon' sizes='180x180' href='/apple-180x180.png'>" & vbCrLf & _
                                  "   <meta http-equiv='content-type' content='text/html;charset=utf-8' />" & vbCrLf & _
                                  "</HEAD><BODY>" & vbCrLf & _
                                  "   <p style='font-family:verdana; font-size:30px; color:blue'>Obsolete but not old</p>" & vbCrLf & _
                                  "</BODY></HTML>"
        Exit Function
    End If
    
    ' 2 nd example - provide your own HTTP header
    If httpRequestIs(httpRequest, "GET /text ") Then
        hcp.setHttpReplyHeaderField "Content-Type", "text/plain"
        provideHttpReplyContent = "Aufert arboribus frondes Autumnus, et idem" & vbCrLf & _
                                  "Fert secum fructus: nos faciamus idem."
        Exit Function
    End If
        
    ' 3 th example - redirect (not any more existing page) to some other page
    If httpRequestIs(httpRequest, "GET /doesNotExistAnyMore ") Then
        hcp.httpReplyStatus = "303 redirect"
        hcp.setHttpReplyHeaderField "Location", "/index.html"
        provideHttpReplyContent = "Redirected." ' whatever different from ""
        Exit Function
    End If
    
    ' 4 th example - handle cookies
    If httpRequestIs(httpRequest, "GET /counter ") Then
        Dim refreshCounter As String
        refreshCounter = hcp.httpRequestCookie("refreshCounter") ' get cookie from HTTP request
        If refreshCounter = "" Then refreshCounter = "0"
        refreshCounter = CStr(CInt(refreshCounter) + 1) ' increase refresh counter and store it as a string
        hcp.setHttpReplyCookie "refreshCounter", refreshCounter, DateAdd("s", 60, Now)  ' set 1 minute valid cookie that will be send to browser with HTTP reply
        provideHttpReplyContent = "<HTML>Web cookies<br><br>This page has been refreshed " & refreshCounter & " times last minute. Click refresh to see more.</HTML>"
        Exit Function
    End If

    ' 5 th example - let vb6WebServer return a file content
    ' if not handeled above just return "" so vb6WebServer will try to locate a file with URL name in application directory
End Function

' ----- FUNCTION THAT MAY BE USEFUL WHILE HANDLING HTTP REQUEST -----

Private Function httpRequestIs(httpRequest As String, methodAndpage As String) As Boolean
    httpRequestIs = InStr(httpRequest, methodAndpage) > 0
End Function

