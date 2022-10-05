VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHttpCnnParams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Parameter class for provideHttpReplyContent function.
'
' Bojan Jurca, 1.10.2022


Option Explicit

' ----- WIN32 API (for cooki expiration dates) -----

Private Type SYSTEMTIME
  wYear As Integer
  wMonth As Integer
  wDayOfWeek As Integer
  wDay As Integer
  wHour As Integer
  wMinute As Integer
  wSecond As Integer
  wMilliseconds As Integer
End Type
Private Type TIME_ZONE_INFORMATION
        Bias As Long
        StandardName(0 To 31) As Integer
        StandardDate As SYSTEMTIME
        StandardBias As Long
        DaylightName(0 To 31) As Integer
        DaylightDate As SYSTEMTIME
        DaylightBias As Long
End Type
Private Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
Private Declare Function SystemTimeToTzSpecificLocalTime Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION, lpUniversalTime As SYSTEMTIME, lpLocalTime As SYSTEMTIME) As Boolean
Private Declare Function TzSpecificLocalTimeToSystemTime Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION, lpLocalTime As SYSTEMTIME, lpUniversalTime As SYSTEMTIME) As Boolean


' ----- MEMBER VARIABLES -----

Private m_socket As Long
Private m_clientIP As String
Private m_httpRequest As String

Private m_httpReplyStatus As String
Private m_httpReplyHeader As String


' ----- PROPERTIES AND METHODS -----

Public Property Let socket(p_socket As Long)
    m_socket = p_socket
End Property

Public Property Get socket() As Long
    socket = m_socket
End Property

Public Property Let clientIP(p_clientIP As String)
    m_clientIP = p_clientIP
End Property

Public Property Get clientIP() As String
    clientIP = m_clientIP
End Property

Public Property Let httpRequest(p_httpRequest As String)
    m_httpRequest = p_httpRequest
End Property

' HTTP header fields are in format \r\nfieldName: fieldValue\r\n
Public Property Get httpRequestHeaderField(fieldName As String) As String
    Dim i As Integer
    i = InStr(1, m_httpRequest, vbCrLf & fieldName & ":", vbTextCompare)
    If i > 0 Then
        Dim j As Integer
        j = InStr(i + 3, m_httpRequest, vbCrLf)
        If j > 0 Then
            Dim S As String
            S = Mid(m_httpRequest, i + Len(fieldName) + 3, j - i - Len(fieldName) - 3)
            httpRequestHeaderField = Trim(S)
        End If
    End If
End Property

' cookies are passed from browser to http server in "cookie" HTTP header field
Public Property Get httpRequestCookie(cookieName As String) As String
    Dim cookieHeaderField As String
    cookieHeaderField = httpRequestHeaderField("Cookie") & ";"

    Dim i As Integer
    i = InStr(1, cookieHeaderField, cookieName, vbTextCompare)
    If i > 0 Then
        i = InStr(i + Len(cookieName), cookieHeaderField, "=")
        If i > 0 Then
            Dim j As Integer
            j = InStr(i + 1, cookieHeaderField, ";")
            If j > 0 Then
                Dim S As String
                S = Mid(cookieHeaderField, i + 1, j - i - 1)
                httpRequestCookie = Trim(S)
            End If
        End If
    End If
End Property

Public Property Get httpReplyStatus() As String
    httpReplyStatus = m_httpReplyStatus
End Property

Public Property Let httpReplyStatus(p_httpReplyStatus As String)
    m_httpReplyStatus = p_httpReplyStatus
End Property

' HTTP reply header fields without HTTP ... and status
Public Property Get httpReplyHeader() As String
    httpReplyHeader = m_httpReplyHeader
End Property

Public Sub setHttpReplyHeaderField(fieldName As String, fieldValue As String)
    m_httpReplyHeader = m_httpReplyHeader & fieldName & ": " & fieldValue & vbCrLf
End Sub

' compose Set-cookie header field like: Set-Cookie: refreshCounter=2; Path=/; Expires=Thu, 09 Dec 2021 19:07:04 GMT
'                                       Set-Cookie: refreshCounter=1; Path=/;3, 4 10 2022 19:30:31 GMT
Sub setHttpReplyCookie(cookieName As String, cookieValue As String, Optional expires As Date = #1/1/1900#, Optional path As String = "/")
    Dim e As String

    If expires <> #1/1/1900# Then
        Dim expiresGMT As Date
        expiresGMT = localTimeToGMT(expires)
        
        Dim dayOfWeek(7) As String
        dayOfWeek(1) = "Sun"
        dayOfWeek(2) = "Mon"
        dayOfWeek(3) = "Tue"
        dayOfWeek(4) = "Wed"
        dayOfWeek(5) = "Thu"
        dayOfWeek(6) = "Fri"
        dayOfWeek(7) = "Sat"

        Dim monthOfYear(12) As String
        monthOfYear(1) = "Jan"
        monthOfYear(2) = "Feb"
        monthOfYear(3) = "Mar"
        monthOfYear(4) = "Apr"
        monthOfYear(5) = "May"
        monthOfYear(6) = "Jun"
        monthOfYear(7) = "Jul"
        monthOfYear(8) = "Aug"
        monthOfYear(9) = "Sep"
        monthOfYear(10) = "Oct"
        monthOfYear(11) = "Nov"
        monthOfYear(12) = "Dec"
        
        e = " Expires=" & dayOfWeek(Weekday(expiresGMT, vbSunday)) & ", " & Day(expiresGMT) & " " & monthOfYear(Month(expiresGMT)) & " " & Year(expiresGMT) & " " & Hour(expiresGMT) & ":" & Minute(expiresGMT) & ":" & Second(expiresGMT) & " GMT"
    End If

    setHttpReplyHeaderField "Set-Cookie", cookieName & "=" & cookieValue & "; Path=" & path & ";" & e
End Sub

Private Function localTimeToGMT(localTime As Date) As Date
    Dim tzi As TIME_ZONE_INFORMATION
    Dim st As SYSTEMTIME
    Dim lt As SYSTEMTIME
    
    GetTimeZoneInformation tzi
    With lt
        .wYear = Year(localTime)
        .wMonth = Month(localTime)
        .wDay = Day(localTime)
        ' .wDayOfWeek = Weekday(localTime) - 1
        .wHour = Hour(localTime)
        .wMinute = Minute(localTime)
        .wSecond = Second(localTime)
        .wMilliseconds = 0
    End With
    TzSpecificLocalTimeToSystemTime tzi, lt, st
    With st
        localTimeToGMT = DateSerial(.wYear, .wMonth, .wDay) + TimeSerial(.wHour, .wMinute, .wSecond)
    End With
End Function

Private Function gmtToLocalTime(gmt As Date) As Date
    Dim tzi As TIME_ZONE_INFORMATION
    Dim st As SYSTEMTIME
    Dim lt As SYSTEMTIME
    
    GetTimeZoneInformation tzi
    With st
        .wYear = Year(gmt)
        .wMonth = Month(gmt)
        .wDay = Day(gmt)
        ' .wDayOfWeek = Weekday(gmt) - 1
        .wHour = Hour(gmt)
        .wMinute = Minute(gmt)
        .wSecond = Second(gmt)
        .wMilliseconds = 0
    End With
    
    SystemTimeToTzSpecificLocalTime tzi, st, lt
    With lt
        gmtToLocalTime = DateSerial(.wYear, .wMonth, .wDay) + TimeSerial(.wHour, .wMinute, .wSecond)
    End With
End Function


