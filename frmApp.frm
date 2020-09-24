VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmApp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Application"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   6630
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrCheck 
      Interval        =   30000
      Left            =   5280
      Top             =   2040
   End
   Begin RichTextLib.RichTextBox rtbOpen 
      Height          =   255
      Left            =   4920
      TabIndex        =   6
      Top             =   5160
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      _Version        =   393217
      TextRTF         =   $"frmApp.frx":0000
   End
   Begin VB.Frame Frame2 
      Caption         =   "Status"
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   3615
      Begin VB.Label lblStatus 
         Caption         =   "Disconnected..."
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      Begin MSWinsockLib.Winsock Winsock 
         Left            =   240
         Top             =   1080
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "Connect"
         Height          =   375
         Left            =   4920
         TabIndex        =   3
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtIP 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   6015
      End
      Begin VB.Label Label1 
         Caption         =   "Location Of Licensing Software"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   3615
      End
   End
End
Attribute VB_Name = "frmApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConnect_Click()

'connect to server

Dim Server As String
'get server ip
If Dir(txtIP.Text & "\SVRIP.DAT", vbNormal) = "" Then
    MsgBox "Error: Make sure the server is up and running.", vbExclamation + vbOKOnly, "Error"
    Exit Sub
Else
    rtbOpen.LoadFile txtIP.Text & "\SVRIP.DAT"
    Server = rtbOpen.Text
End If

'connect
Winsock.Close
Call Winsock.Connect(Server, 123)
        
'wait whilst connecting
Do While Winsock.State <> sckConnected: DoEvents
    'wait for socket to connect
    If Winsock.State = sckError Then Exit Sub
Loop

'if connected then send a request for a licence
If Winsock.State = sckConnected Then
    Winsock.SendData ("ConReq")
End If

End Sub

Private Sub Form_Load()
    txtIP.Text = App.path
End Sub

Private Sub Form_Terminate()
    Winsock.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Winsock.Close
End Sub

Private Sub tmrCheck_Timer()
    'checks every 30 seconds to make sure the user is still connected
    If Winsock.State <> sckConnected And cmdConnect.Enabled = False Then
        'do not allow user to do stuff
        MsgBox "You have been disconnected from the server!", vbExclamation + vbOKOnly, "Error"
        cmdConnect.Enabled = True
        lblStatus.Caption = "Disconnected"
    End If
End Sub

Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)
    Dim Data As String
    Call Winsock.GetData(Data, , bytesTotal)
    
    'check wether licence was granted or not
    Select Case Data
        Case "Granted"
            lblStatus = "Licence granted."
            cmdConnect.Enabled = False
        Case Else
            'licences all used up or no key info on server
            lblStatus = "No valid licence found."
    End Select
End Sub

