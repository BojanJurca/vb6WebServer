VERSION 5.00
Begin VB.Form frmVb6WebServer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "vb6WebServer"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   7290
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmVb6WebServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public errorMessage As String

Private Sub Form_Activate()
    Print ""
    Print "   " & App.Title & " is listening on " & vb6WebServerIP & ":" & vb6WebServerPort
    If errorMessage > "" Then
        ForeColor = vbRed
        Print ""
        Print "   " & errorMessage
        ForeColor = vbBlack
    End If
    Print ""
    Print "   vb6WebServer directory: " & App.path & "\html\"
    Print ""
    Print "   Please change IP address and port number in ""Main"" subroutine and" & vbCrLf & "   modify ""provideHttpReplyContent"" function according to your needs."
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' notify Sub Main to stop vb6WebServer after unloading
    requestToStopVb6WebServer = True
End Sub

