Attribute VB_Name = "Module1"
Option Explicit

#If Win32 Then

   Public Declare Function SendMessage Lib "user32" Alias _
       "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
       ByVal wParam As Long, lParam As Long) As Long

#Else

   Public Declare Function SendMessage Lib "User" ( _
       ByVal hwnd As Integer, ByVal wMsg As Integer, _
       ByVal wParam As Integer, lParam As Any) As Long

#End If

Public Const WM_SETREDRAW = &HB
Public msOldString As String ' module level global
Public miStart As Integer    ' module level global
Public miLength As Integer   ' module level global

