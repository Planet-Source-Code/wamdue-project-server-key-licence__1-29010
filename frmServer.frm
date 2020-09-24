VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server Licence Administrator"
   ClientHeight    =   4125
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   8535
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   8535
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close Server"
      Height          =   855
      Left            =   6840
      Picture         =   "frmServer.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2760
      Top             =   3120
   End
   Begin prjLicence.Server Server 
      Left            =   3120
      Top             =   3120
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin RichTextLib.RichTextBox rtbOpen 
      Height          =   255
      Left            =   1680
      TabIndex        =   4
      Top             =   4440
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmServer.frx":0442
   End
   Begin VB.Frame Frame2 
      Caption         =   "Server UID"
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   2535
      Begin VB.TextBox txtUID 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   360
         Width           =   2295
      End
   End
   Begin MSComDlg.CommonDialog cdMain 
      Left            =   960
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Licence Keys"
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8295
      Begin MSComctlLib.ListView lvKeys 
         Height          =   2415
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   4260
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Licence Key"
            Object.Width           =   7937
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "No. Licences"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Free"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "UID"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   8640
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Menu"
      Begin VB.Menu mnuOpenLK 
         Caption         =   "Open Licence Key"
      End
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetVolumeSerialNumber Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long

Dim iInUse As Integer
Dim s(0 To 255) As Integer 'S-Box
Dim kep(0 To 255) As Integer
'For the file actions
Dim path As String


Public Sub RC4ini(Pwd As String)
    Dim temp As Integer, a As Integer, b As Integer
    'Save Password in Byte-Array
    b = 0


    For a = 0 To 255
        b = b + 1


        If b > Len(Pwd) Then
            b = 1
        End If
        kep(a) = Asc(Mid$(Pwd, b, 1))
    Next a
    'INI S-Box


    For a = 0 To 255
        s(a) = a
    Next a
    b = 0


    For a = 0 To 255
        b = (b + s(a) + kep(a)) Mod 256
        ' Swap( S(i),S(j) )
        temp = s(a)
        s(a) = s(b)
        s(b) = temp
    Next a
End Sub

Public Function EnDeCrypt(plaintxt As Variant) As Variant
    Dim temp As Integer, a As Long, i As Integer, j As Integer, k As Integer
    Dim cipherby As Byte, cipher As Variant


    For a = 1 To Len(plaintxt)
        i = (i + 1) Mod 256
        j = (j + s(i)) Mod 256
        ' Swap( S(i),S(j) )
        temp = s(i)
        s(i) = s(j)
        s(j) = temp
        'Generate Keybyte k
        k = s((s(i) + s(j)) Mod 256)
        'Plaintextbyte xor Keybyte
        cipherby = Asc(Mid$(plaintxt, a, 1)) Xor k
        cipher = cipher & Chr(cipherby)
    Next a
    EnDeCrypt = cipher
End Function

Private Sub cmdClose_Click()
    Server.StopServer
    Unload Me
End Sub

Private Sub Form_Load()
               
    Dim Hostname As String, IPAdd As String
    'get Server IP and save it to local file
    Hostname = GetIPHostName()
    IPAdd = GetIPAddress()

    rtbOpen.Text = ""
    rtbOpen.Text = IPAdd
    rtbOpen.SaveFile App.path & "\SVRIP.DAT", rtfText
    rtbOpen.Text = ""
    
    txtUID = Trim(VolumeSerialNumber("C:\"))
    
    Call Server.StartServer(123, IPAdd)
    
    iInUse = 0
    
    OpenLocalKeys
    
End Sub

Private Sub Form_Terminate()
    On Error Resume Next
    Kill App.path & "\SVRIP.DAT"
    Call Server.StopServer
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Kill App.path & "\SVRIP.DAT"
    Call Server.StopServer
End Sub

Private Sub mnuOpenLK_Click()
    'I chose .dlk as my licence key extension
    
    On Error GoTo InvalidKey
    
    cdMain.Filter = "*.dlk|*.dlk"
    cdMain.FileName = ""
    cdMain.ShowOpen
    
    'check if a file was selected
    If cdMain.FileName <> "" Then
        'open key
        Dim MyStr As String
        Dim MyCipher As String
        Dim mlen As Integer
        rtbOpen.LoadFile cdMain.FileName, rtfText
        mlen = Left(rtbOpen.Text, 2)
        MyCipher = Mid(rtbOpen.Text, 38, mlen)
        RC4ini (txtUID.Text)
        MyStr = EnDeCrypt(MyCipher)
        
        Dim mynum As Double
        Dim myhex As Long
        Dim myhex2 As String
        
        mynum = Split(MyStr, " ")(0)
        myhex = Split(MyStr, " ")(1)
        myhex2 = Hex(myhex)
        
        a = mynum - Asc(Mid(myhex2, 2, 1))
        b = Asc(Right(myhex2, 1))
        d = Asc(Left(myhex2, 1))
        
        c = a / d
        f = Int(c / b)
        
        'first check that this is not just a renamed file
        'use the unique hex code for this
        
        For i = 1 To lvKeys.ListItems.Count
            If myhex2 = lvKeys.ListItems(i).SubItems(3) Then
                MsgBox "This key is already registered!", vbExclamation + vbOKOnly, "Error"
                Exit Sub
            End If
        Next i
        
        'f is the number of licences
        Dim LI As ListItem
        Set LI = lvKeys.ListItems.Add(, , cdMain.FileName)
        LI.SubItems(1) = Str(f)
        LI.SubItems(2) = Str(f)
        LI.SubItems(3) = myhex2
        
    End If
    
    Exit Sub
    
InvalidKey:

MsgBox "Key is invalid!", vbExclamation + vbOKOnly, "Error"
    
End Sub

Private Function VolumeSerialNumber(ByVal RootPath As String) As String

Dim VolLabel As String
Dim VolSize As Long
Dim Serial As Long
Dim MaxLen As Long
Dim Flags As Long
Dim Name As String
Dim NameSize As Long
Dim s As String
Dim ret As Boolean

ret = GetVolumeSerialNumber(RootPath, VolLabel, VolSize, Serial, MaxLen, Flags, Name, NameSize)

If ret Then
    VolumeSerialNumber = Str(Serial)
Else
    VolumeSerialNumber = "00000000"
End If

End Function

Private Sub Server_DataArrival(ByVal SckIndex As Integer, ByVal Data As String, ByVal bytesTotal As Long, ByVal RemoteIP As String, ByVal RemoteHost As String)
    'Call sOutput(FormatNumber(bytesTotal, 0, , , vbTrue) & " bytes recieved.", RemoteIP)
    
    If Data = "ConReq" Then
        For i = 1 To lvKeys.ListItems.Count
            If Val(lvKeys.ListItems(i).SubItems(2)) > 0 Then
                lvKeys.ListItems(i).SubItems(2) = Val(lvKeys.ListItems(i).SubItems(2)) - 1
                Server.SendData "Granted", SckIndex
                Exit Sub
            End If
        Next i
        Server.SendData "Invalid Command.", SckIndex
    End If
    
    If Data = "Closing" Then
        For i = 1 To lvKeys.ListItems.Count
            If Val(lvKeys.ListItems(i).SubItems(2)) < Val(lvKeys.ListItems(i).SubItems(1)) Then
                lvKeys.ListItems(i).SubItems(2) = Val(lvKeys.ListItems(i).SubItems(2)) + 1
                Exit Sub
            End If
        Next i
    End If
    
End Sub

Private Sub Timer1_Timer()
    
    'count connections, hence licences
    If Server.ConnectionCount >= iInUse Then
        iInUse = Server.ConnectionCount
    Else
        For j = 1 To iInUse - Server.ConnectionCount
            For i = 1 To lvKeys.ListItems.Count
                If Val(lvKeys.ListItems(i).SubItems(2)) < Val(lvKeys.ListItems(i).SubItems(1)) Then
                    lvKeys.ListItems(i).SubItems(2) = Val(lvKeys.ListItems(i).SubItems(2)) + 1
                    Exit For
                End If
            Next i
        Next j
    End If
    
End Sub

Private Sub OpenLocalKeys()
        Dim MyStr As String
        Dim MyCipher As String
        Dim mlen As Integer
        
        m = Dir(App.path & "\*.dlk", vbNormal)
        If m = "" Then
            Exit Sub
        End If
        
        Do
            rtbOpen.LoadFile App.path & "\" & m, rtfText
            mlen = Left(rtbOpen.Text, 2)
            MyCipher = Mid(rtbOpen.Text, 38, mlen)
            RC4ini (txtUID.Text)
            MyStr = EnDeCrypt(MyCipher)
            
            Dim mynum As Double
            Dim myhex As Long
            Dim myhex2 As String
            
            mynum = Split(MyStr, " ")(0)
            myhex = Split(MyStr, " ")(1)
            myhex2 = Hex(myhex)
            
            a = mynum - Asc(Mid(myhex2, 2, 1))
            b = Asc(Right(myhex2, 1))
            d = Asc(Left(myhex2, 1))
            
            c = a / d
            f = Int(c / b)
            
            'first check that this is not just a renamed file
            'use the unique hex code for this
            
            For i = 1 To lvKeys.ListItems.Count
                If myhex2 = lvKeys.ListItems(i).SubItems(3) Then
                    MsgBox "This key is already registered!", vbExclamation + vbOKOnly, "Error"
                    Exit Sub
                End If
            Next i
            
            'f is the number of licences
            Dim LI As ListItem
            Set LI = lvKeys.ListItems.Add(, , App.path & "\" & m)
            LI.SubItems(1) = Str(f)
            LI.SubItems(2) = Str(f)
            LI.SubItems(3) = myhex2
            m = Dir
        Loop Until m = ""
End Sub
