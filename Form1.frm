VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Tech Hata Kapatma"
   ClientHeight    =   1875
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   1875
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   120
      Top             =   600
   End
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   120
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
 Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Const BM_CLICK = &HF5
 Const WM_CLOSE = &H10

Private Sub Timer1_Timer()
    Dim handle1 As Long
    Dim handle2 As Long
    handle1 = FindWindow("#32770", "Notice")
    handle2 = FindWindowEx(handle1, ByVal 0&, "Button", "&Hayýr")
    If handle2 > 0 Then
   Call SendMessage(handle2, BM_CLICK, 0, 0)
    List1.AddItem Time & " >> Hata Kapandi"
    End If
End Sub
Private Sub Timer2_Timer()
    Dim handle3 As Long
    Dim handle4 As Long
    handle3 = FindWindow("#32770", "Notice")
    handle4 = FindWindowEx(handle3, ByVal 0&, "Button", "&No")
    If handle4 > 0 Then
   Call SendMessage(handle4, BM_CLICK, 0, 0)
    List1.AddItem Time & " >> Error window closed ;)"
    End If
End Sub

