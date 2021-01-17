VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   480
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   600
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Created by Rizky Khapidsyah

   Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim sComboText As String
    Dim iLoop As Integer
    Dim sTempString As String
    Dim lReturn As Long
    Dim bInList As Boolean
    Dim sItem
    
  If Not KeyCode = Asc(vbTab) And Not KeyCode = vbKeyShift And _
      Not KeyCode = vbKeyLeft And Not KeyCode = vbKeyRight And _
      Not KeyCode = vbKeyHome And Not KeyCode = vbKeyEnd Then
       
        bInList = False
        
        With Combo1
            sTempString = .Text
            If Len(sTempString) = 1 Then sComboText = sTempString
            lReturn = SendMessage(.hwnd, WM_SETREDRAW, False, 0&)
            For iLoop = 0 To (.ListCount - 1)
                sItem = .List(iLoop)
                If UCase((sTempString & Mid$(sItem, _
                 Len(sTempString) + 1))) = UCase(sItem) Then
                    .ListIndex = iLoop
                    .Text = sItem
                    msOldString = sItem
                    miStart = Len(sTempString)
                    .SelStart = miStart
                    miLength = Len(sItem) - (Len(sTempString))
                    .SelLength = miLength
                    sComboText = sComboText & Mid$(sTempString, _
                        Len(sComboText) + 1)
                    bInList = True
                    Exit For
                End If
            Next iLoop
            
            If Not bInList Then
                .Text = msOldString
                .SelStart = miStart
                .SelLength = miLength
            End If
            
            lReturn = SendMessage(.hwnd, WM_SETREDRAW, True, 0&)
        End With
    End If
   End Sub

Private Sub Form_Load()
    With Combo1
        .AddItem "Blue"
        .AddItem "Green"
        .AddItem "Yellow"
        
        .ListIndex = 0
        .Text = .List(0)
        .SelStart = 0
        .SelLength = Len(.Text)
        msOldString = .Text
        miStart = 0
        miLength = .SelLength
    End With
End Sub



