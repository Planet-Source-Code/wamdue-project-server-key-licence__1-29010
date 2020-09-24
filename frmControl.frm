VERSION 5.00
Begin VB.Form frmControl 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control Panel"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1845
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   1845
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Application"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create Key"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton cmdServer 
      Caption         =   "Server"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetNextWindow Lib "user32" Alias "GetWindow" (ByVal hwnd As Long, ByVal wFlag As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long

Private Sub cmdCreate_Click()
    frmCreateKey.Show
End Sub

Private Sub cmdServer_Click()
    'check to see if any other copy of the server program is running
    'as this would be a way to use the same licence file over and over again in
    'different directories on the same machine
    Dim lhWnd As Long
    lhWnd = FindAnyWindow(Me, "Server Licence Administrator")
    If lhWnd <> 0 Then ' 0 means Server not running.
        MsgBox "Server is already running on this machine!", vbExclamation + vbOKOnly, "Error"
        Unload Me
        Exit Sub
    End If

    frmServer.Show
End Sub

Private Sub Command1_Click()
    frmApp.Show
End Sub

Private Function GetCaption(ByVal lhWnd As Long) As String
    Dim sA As String, lLen As Long
    
    lLen& = GetWindowTextLength(lhWnd&)
    sA$ = String(lLen&, 0&)
    Call GetWindowText(lhWnd&, sA$, lLen& + 1)
    GetCaption$ = sA$
End Function

Private Function FindAnyWindow(frm As Form, ByVal WinTitle As String, Optional ByVal CaseSensitive As Boolean = False) As Long
    Dim lhWnd As Long, sA As String
    lhWnd& = frm.hwnd


    Do Until lhWnd& = 0


        DoEvents
            
            sA$ = GetCaption(lhWnd&)
            If InStr(IIf(CaseSensitive = False, LCase$(sA$), sA$), IIf(CaseSensitive = False, LCase$(WinTitle$), WinTitle$)) Then FindAnyWindow& = lhWnd&: Exit Do Else FindAnyWindow& = 0
            
            lhWnd& = GetNextWindow(lhWnd&, 2)
    Loop
End Function
