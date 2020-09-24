VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmCreateKey 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create Key"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3030
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   3030
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox rtbSave 
      Height          =   255
      Left            =   1920
      TabIndex        =   12
      Top             =   4800
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      _Version        =   393217
      TextRTF         =   $"frmCreateKey.frx":0000
   End
   Begin MSComDlg.CommonDialog cdMain 
      Left            =   1080
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Output"
      Height          =   1095
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   2775
      Begin VB.TextBox txtLK 
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label3 
         Caption         =   "Licence Key"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Input"
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
      Begin VB.TextBox txtLN 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Text            =   "1"
         Top             =   1200
         Width           =   2535
      End
      Begin VB.CommandButton cmdCreate 
         Caption         =   "Create Key"
         Default         =   -1  'True
         Height          =   375
         Left            =   1560
         TabIndex        =   3
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtSerialNum 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "No. Of Licences"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Client UID"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   3600
      Width           =   2775
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save Key"
         Height          =   375
         Left            =   1560
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmCreateKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim s(0 To 255) As Integer 'S-Box
Dim kep(0 To 255) As Integer
Dim i As Integer, j As Integer



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

Private Sub cmdCreate_Click()
    
    If txtSerialNum = "" Then
        MsgBox "Client UID must have valid data!", vbExclamation + vbOKOnly, "Error"
        Exit Sub
    End If
    
    RC4ini (txtSerialNum)
    Randomize
    a1 = Int(Rnd * 88888888) + 11111111
    a = Hex(a1)
    b = Asc(Right(a, 1))
    d = Asc(Left(a, 1))
    qw = Asc(Mid(a, 2, 1))
    c = ((Val(txtLN) * b) * d) + Asc(Mid(a, 2, 1))
    txtLK = EnDeCrypt(c & " " & a1)

End Sub


Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()

    cdMain.FileName = ""
    cdMain.Filter = "*.dlk|*.dlk"
    cdMain.ShowSave
    
    If cdMain.FileName <> "" Then
        Randomize
        rtbSave.Text = Len(txtLK)
        For i = 1 To 35
            a = Int(Rnd * 250) + 1
            rtbSave.Text = rtbSave.Text & Chr(a)
        Next i
        rtbSave.Text = rtbSave.Text & txtLK
        For i = 1 To 35
            a = Int(Rnd * 250) + 1
            rtbSave.Text = rtbSave.Text & Chr(a)
        Next i
        rtbSave.SaveFile cdMain.FileName, rtfText
    End If
    
End Sub
