<div align="center">

## Make a Client and Server Chat Room using Winsock


</div>

### Description

This program will allow more than a one on one direct connection chat, like previous postings show. This will show you how to make a server and client programs that you can distribue and have as many people as you want in the same chat.
 
### More Info
 
In code


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Matt Insler](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/matt-insler.md)
**Level**          |Unknown
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/matt-insler-make-a-client-and-server-chat-room-using-winsock__1-1737/archive/master.zip)

### API Declarations

In code


### Source Code

```
'Name: Client and Server Chat Room (Server)
'Author: Matt Insler
'Written: 5/7/99
'Purpose: This program will allow more than a one on one direct connection chat, like previous postings show.
' This will allow as many clients as have the client and the host name or IP to chat by using a server to
' receive the messages and send them back out to all computers in the collection. This is a good start
' for a mIrc style chat, or an AOL style chat, or any other type of chat program. By adding a listbox
' to the client and making a procedure that will send all of the names to the clients, and a procedure to
' receive and add the names, you can make a listbox showing who is in the room. Also, if you wish to make
' separate channels, or rooms, you can either run multiple versions of the server on different ports, or
' you can add more winsock controls and have them all simultaneously listening and running the server.
' If you happen to use my code as a stepping stool to a good chat program or find any ways to make this program
' better, please send it to me at racobac@aol.com. Thanks.
'Input: Nothing, but to sit back and watch people chat, or to chat with them as ServerMaster.
'Returns: Watch the chat happen, and facilitate a server for people to chat on.
'Side Effects: None that I know of. If you find any, please email me.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'*****************************************************************************************************************
'Create a new form and add three(3) text boxes, one(1) command button, one(1) list box, and add the microsoft winsock control
'Change the name of the text boxes to tMain, and tSend, tIP, name the command button cSnd, and name the list box lName
'Change the name of the winsock control to Wsck
'Change the caption of cSnd to "Send"
'Make tMain multiline = true, scrollbars = 2 - vertical, and locked = true
'Make lName Sorted = true
'Make cSnd Default = true
'Insert the following code
'Declarations:
Dim Client As New Collection
Dim Names As New Collection
Const Indicator = ":':"
Private Sub cSnd_Click()
 'Send button
 'Make string to send
 txt$ = "ServerMaster: " & tSend.Text & Chr$(13) & Chr$(10)
 'Send to clients
 Call SendOut(txt$)
 'Clear Send text box
 tSend.Text = ""
End Sub
Private Sub Form_Load()
 'Clear Main text box
 tMain.Text = ""
 'We will be using UDP for this program because it does not establish a constant connection to another computer.
 'This will allow the server to keep "listening" for messages from other addresses on a network or the internet.
 Wsck.Protocol = sckUDPProtocol
 'Set your constant port (must be the same in clients)
 Wsck.LocalPort = 2367
 'Start listening
 Wsck.Bind
 'Add the server to the name list
 'This would allow you to make a list box in the client that could receive all of the names of the people in the room.
 RmIP = Wsck.LocalIP
 RmPt = 2367
 Names.Add Key:=RmIP, Item:="ServerMaster"
 'Display your IP Address for client use, and Computer Name for network use.
 tIP.Text = RmIP & " / " & Wsck.LocalHostName
End Sub
Private Sub Form_Unload(Cancel As Integer)
 'End connection on Winsock
 Wsck.Close
 End
End Sub
Private Sub lName_DblClick()
 'Double-click an IP Address in the listbox
 'Create message with client NickName, IP Address, and Port
 txt$ = Names(lName.Text) & ", " & lName.Text & ", " & Client(lName.Text)
 MsgBox txt$, vbOKOnly, "User Information"
End Sub
Private Sub Wsck_DataArrival(ByVal bytesTotal As Long)
 'Winsock received a message
 'If an error occurs, ignore it and go on to the next command
 On Error Resume Next
 Dim DATA As String
 Dim DATA2 As String
 Dim Nam As String
 Dim MsgText As String
 'Retreive message in string format
 Wsck.GetData DATA, vbString
 'Get client's IP and Port
 RmIP = Wsck.RemoteHostIP
 RmPt = Wsck.RemotePort
 'Get first letter of message
 DATA2 = Left(DATA, 1)
 'Get the rest of the message
 DATA = Mid(DATA, 2)
 'If the message is a system command:
 If DATA2 = "s" Then
 'If a client wants to connect to the room:
 If Left(DATA, 20) = Indicator & "CoNnEcTrEqUeSt" & Indicator Then
  'Extract the client NickName from the message
  Nam = Mid(DATA, 21)
  'Add client's IP and Port to your collections
  Client.Add Key:=RmIP, Item:=RmPt
  Names.Add Key:=RmIP, Item:=Nam
  'Add client's IP to the listbox
  lName.AddItem RmIP
  Exit Sub
 'If a client wants to disconnect from the room:
 ElseIf DATA = Indicator & "CoNnEcTcAnCeL" & Indicator Then
  'Loop through listbox and find client's IP
  For X = 0 To lName.ListCount - 1
  lName.ListIndex = X
  RmEx = lName.Text
  'When found, remove IP from listbox
  If RmEx = RmIP Then lName.RemoveItem (X)
  Next
  'Remove client from your collections
  Client.Remove (RmIP)
  Names.Remove (RmIP)
  Exit Sub
 End If
 'If the message is text sent to the room:
 ElseIf DATA2 = "t" Then
 'Send text to clients
 Call SendOut(DATA)
 End If
End Sub
Private Sub Wsck_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
 'Error occured in winsock!
 MsgBox "An error occurred in winsock!"
 'Close connection
 Wsck.Close
End Sub
Sub SendOut(StringToSend As String)
 'Send a text message to all clients in collection/listbox
 'If an error occurs, ignore it and go on to the next command
 On Error Resume Next
 'Loop through all IP in listbox
 For X = 0 To lName.ListCount - 1
 'Select each IP
 lName.ListIndex = X
 'Set IP and Port to send to
 RmIP = lName.Text
 RmPt = Client(RmIP)
 Wsck.RemoteHost = RmIP
 Wsck.RemotePort = RmPt
 'Send text message
 Wsck.SendData "t" & StringToSend
 Next
 'Add the text message to your room
 tMain.Text = tMain.Text & StringToSend
 'Scroll to the bottom of the room
 tMain.SelStart = Len(tMain)
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'*****************************************************************************************************************
'Name: Client and Server Chat Room (Client)
'Author: Matt Insler
'Written: 5/7/99
'Purpose: This program will allow more than a one on one direct connection chat, like previous postings show.
' This will allow as many clients as have the client and the host name or IP to chat by using a server to
' receive the messages and send them back out to all computers in the collection. This is a good start
' for a mIrc style chat, or an AOL style chat, or any other type of chat program. By adding a listbox
' to the client and making a procedure that will send all of the names to the clients, and a procedure to
' receive and add the names, you can make a listbox showing who is in the room. Also, if you wish to make
' separate channels, or rooms, you can either run multiple versions of the server on different ports, or
' you can add more winsock controls and have them all simultaneously listening and running the server.
' If you happen to use my code as a stepping stool to a good chat program or find any ways to make this program
' better, please send it to me at racobac@aol.com. Thanks.
'Input: Host IP or Computer Name, and a NickName, along with whatever you wish to send to the room.
'Returns: What everyone who is in the room types back to you, along with your messages.
'Side Effects: None that I know of. If you find any, please email me.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'*****************************************************************************************************************
'Create a new form and add four(4) text boxes, three(3) command buttons, and add the microsoft winsock control
'Change the name of the text boxes to tHost, tName, tMain, and tSend, and name the command buttons cCon, cDis, cSnd
'Change the name of the winsock control to Wsck
'Change the caption of cCon to "Connect", cDis to "Disconnect", and cSnd to "Send"
'Make tMain multiline = true, scrollbars = 2 - vertical, and locked = true
'Make cDis and cSnd enabled = false
'
'Make cSnd Default = true
'Insert the following code
'Declarations:
Const Indicator = ":':"
Private Sub cCon_Click()
 'Connect button
 'Check if a Host Name or IP has been entered
 If Len(tHost) < 1 Then
 MsgBox ("Please make sure a Host has been entered!")
 'Put blinker in host text box
 tHost.SetFocus
 Exit Sub
 'Check if a NickName has been entered
 ElseIf Len(tName) < 1 Then
 MsgBox "You must enter a nickname first!"
 'Put blinker in NickName text box
 tName.SetFocus
 Exit Sub
 End If
 'If an error occurs, jump to Ending
 On Error GoTo Ending
 'Set the IP or Host Computer to connect to
 Wsck.RemoteHost = tHost.Text
 'Randomize a Port setting
 Wsck.LocalPort = Int((9999 * Rnd) + 1)
 'Set the Port to connect to
 Wsck.RemotePort = 2367
 'Connect!
 Wsck.Bind
 'Send system request to connect
 Wsck.SendData "s" & Indicator & "CoNnEcTrEqUeSt" & Indicator & tName.Text
 'Enable Send and Disconnect buttons, and disable Connect button and NickName text box
 cSnd.Enabled = True
 cDis.Enabled = True
 cCon.Enabled = False
 tName.Enabled = False
 'Put blinker in the Send text box
 tSend.SetFocus
 Exit Sub
Ending:
 'Error handling
 MsgBox "You are not connected to the internet or the Host is not available.", , Form1.Caption
 'Click the Disconnect button
 cDis_Click
End Sub
Private Sub cDis_Click()
 'Disconnect button
 'If an error occurs, ignore it and go on to the next command
 On Error Resume Next
 'Send system message to disconnect from server
 Wsck.SendData "s" & Indicator & "CoNnEcTcAnCeL" & Indicator
 'Close connection
 Wsck.Close
 'Enable Connect button and NickName text box, and disable Send and Disconnect buttons
 cCon.Enabled = True
 tName.Enabled = True
 cDis.Enabled = False
 cSnd.Enabled = False
 'Put blinker in NickName text box
 tName.SetFocus
End Sub
Private Sub cSnd_Click()
 'Send button
 Wsck.SendData "t" & tName.Text & ":" & vbTab & tSend.Text & Chr$(13) & Chr$(10)
 'Clear Send text box
 tSend.Text = ""
End Sub
Private Sub Form_Load()
 'We will be using UDP for this program because it does not establish a constant connection to another computer.
 'This will allow the server to keep "listening" for messages from other addresses on a network or the internet.
 Wsck.Protocol = sckUDPProtocol
 'Clear Main text box
 tMain.Text = ""
End Sub
Private Sub Form_Unload(Cancel As Integer)
 'End connection on Winsock
 Wsck.Close
 End
End Sub
Private Sub Wsck_DataArrival(ByVal bytesTotal As Long)
 'If an error occurs, ignore it and go on to the next command
 On Error Resume Next
 Dim Data As String
 Dim Data2 As String
 'Retreive message in string format
 Wsck.GetData Data, vbString
 'Get first letter of message
 Data2 = Left(Data, 1)
 'Get the rest of the message
 Data = Mid(Data, 2)
 'If the message is a system command:
 If Data2 = "s" Then
 'You can add your own system commands from the server to the client here.
 'I have made one to throw out the client if I decide to.
 'If the message is text sent to the room:
 ElseIf Data2 = "t" Then
 'Add the text message to your room
 tMain.Text = tMain.Text & Data
 'Scroll to the bottom of the room
 tMain.SelStart = Len(tMain)
 Exit Sub
 End If
End Sub
Private Sub Wsck_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
 'Error occured in winsock!
 MsgBox "An error occurred in winsock!"
 'Close connection
 Wsck.Close
End Sub
```

