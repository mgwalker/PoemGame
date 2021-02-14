VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "The Poem Game"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11850
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   11850
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAway 
      Caption         =   "Away"
      Height          =   375
      Left            =   10020
      TabIndex        =   40
      Top             =   4200
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Frame Frame1 
      Caption         =   "Vote For This Poem"
      Enabled         =   0   'False
      ForeColor       =   &H8000000D&
      Height          =   1695
      Left            =   120
      TabIndex        =   27
      Top             =   3240
      Width           =   2175
      Begin VB.CommandButton cmdSendVotes 
         Caption         =   "Vote!"
         Height          =   375
         Left            =   1200
         TabIndex        =   33
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton opt1Point 
         Caption         =   "1 Point"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   1320
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton opt2Points 
         Caption         =   "2 Points"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton opt3Points 
         Caption         =   "3 Points"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton opt4Points 
         Caption         =   "4 Points"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton opt5Points 
         Caption         =   "5 Points"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdSubmitTopics 
      Caption         =   "Submit Topics"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3240
      TabIndex        =   23
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox txtTopic2 
      Height          =   285
      Left            =   720
      TabIndex        =   22
      Top             =   2640
      Width           =   2415
   End
   Begin VB.TextBox txtTopic1 
      Height          =   285
      Left            =   720
      TabIndex        =   21
      Top             =   2400
      Width           =   2415
   End
   Begin VB.CommandButton cmdStartGame 
      BackColor       =   &H80000000&
      Caption         =   "Start Game"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1800
      TabIndex        =   18
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdSubmitPoem 
      Caption         =   "Submit Poem"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3240
      TabIndex        =   17
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox txtCompose 
      ForeColor       =   &H00000000&
      Height          =   1695
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      Top             =   600
      Width           =   4335
   End
   Begin VB.CommandButton cmdDeselect 
      Caption         =   "Deselect"
      Height          =   375
      Left            =   10980
      TabIndex        =   14
      Top             =   4200
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.CommandButton cmdKick 
      Caption         =   "Kick"
      Height          =   375
      Left            =   9120
      TabIndex        =   8
      Top             =   4200
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ListBox lstConnections 
      Height          =   3765
      ItemData        =   "frmMain.frx":08CA
      Left            =   9000
      List            =   "frmMain.frx":08CC
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   360
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   5880
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
   Begin VB.TextBox txtIP 
      Height          =   285
      Left            =   9000
      TabIndex        =   1
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox txtDialog 
      Height          =   3765
      Left            =   4560
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   360
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.TextBox txtMessage 
      Height          =   285
      Left            =   4560
      TabIndex        =   5
      Top             =   4200
      Visible         =   0   'False
      Width           =   4455
   End
   Begin MSWinsockLib.Winsock sckConnect 
      Left            =   8040
      Top             =   4680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   135
      Left            =   11640
      TabIndex        =   6
      Top             =   240
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.CommandButton cmdDisconnect 
      Caption         =   "Disconnect"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   375
      Left            =   8040
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdHost 
      Caption         =   "Host"
      Height          =   375
      Left            =   6720
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock sckConnection 
      Index           =   0
      Left            =   7680
      Top             =   4680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdEndGame 
      Caption         =   "End Game"
      Height          =   375
      Left            =   1800
      TabIndex        =   24
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer tmrPoemHold 
      Left            =   8880
      Top             =   4680
   End
   Begin VB.Timer tmrSendData 
      Interval        =   1
      Left            =   8520
      Top             =   4680
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Back"
      Height          =   375
      Left            =   10020
      TabIndex        =   41
      Top             =   4200
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label lblTopic2 
      Height          =   255
      Left            =   2520
      TabIndex        =   39
      Top             =   4800
      Width           =   2055
   End
   Begin VB.Label lblTopic1 
      Height          =   255
      Left            =   2520
      TabIndex        =   38
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Label Label10 
      Caption         =   "Topic 2:"
      Height          =   255
      Left            =   2520
      TabIndex        =   37
      Top             =   4560
      Width           =   615
   End
   Begin VB.Label Label9 
      Caption         =   "Topic 1:"
      Height          =   255
      Left            =   2520
      TabIndex        =   36
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label lblCurrPlayer 
      Height          =   255
      Left            =   2520
      TabIndex        =   35
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "Current Player:"
      Height          =   255
      Left            =   2520
      TabIndex        =   34
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label editColors 
      Caption         =   "Colors"
      Height          =   255
      Left            =   720
      TabIndex        =   26
      Top             =   0
      Width           =   495
   End
   Begin VB.Label AboutBox 
      Caption         =   "About"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000A&
      Caption         =   "Topic 2:"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label Label7 
      BackColor       =   &H8000000A&
      Caption         =   "Topic 1:"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000A&
      Caption         =   "Here's the Poem:"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      Caption         =   "Members:"
      Height          =   195
      Left            =   9480
      TabIndex        =   13
      Top             =   360
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      Caption         =   "Chat:"
      Height          =   195
      Left            =   4680
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      Caption         =   "Connect to:"
      Height          =   195
      Left            =   8160
      TabIndex        =   11
      Top             =   360
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      Caption         =   "Username:"
      Height          =   195
      Left            =   5040
      TabIndex        =   10
      Top             =   360
      Width           =   765
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Gate's Poem Game
'(c) Copyright 2001, JacketFan
'
'This game is based on the Poem Game created by
'Gate for play during Late Night JRChat.  However,
'the game was so popular among its players that a
'daytime version was needed and since JRChat is
'often crowded during the day, a new place to play
'it was needed that would not interfere with normal
'chat in JRChat.  So I made this program.  =)
'
'Revisions:
'  1.00 -- Original Public Version
'  1.01
'       Fixes several bugs which hamper or prevent
'       proper gameplay.  Makes some aesthetic changes
'       for a nicer layout and better appearance.
'  1.02
'       Fixed several new bugs.
'  1.20
'       Added the green background
'  1.21
'       Fixed some more bugs.  =P
'
'
'  2.00 --
'O       Fixed a non-fatal installation error involving msado15.dll
'X       Limited topic length to 25 characters
'X       Fixed Dead-Game situation (Player leaves/kicked during turn)
'X       Added full color customization (saves preferences to INI)
'X       Added an "About" box with current version information
'X       Added partial IPs to the userlist.
'X       Poem box permanently enabled (for scrolling)
'O       Added scoring system (From -1 to 3 points)
'X       Added /msg, .msg, and >, getting rid of Private checkbox
'X       Moved status into Chat window, got rid of status window
'
'
' 2.1 - Bug fixes
'
' 2.11 - More bug fixes
        

'This forces the declaration of all variables, preventing misspellings of variables and type mismatches.
Option Explicit

'Default port to be used when establishing connections.
Const DEFAULT_PORT = 600

'INI file to be used by the program.
Const INI_FILE = "manychat.ini"

'Sometimes parameters are sent along with the commands that are sent between computers.
'All parameters will be formatted to be exactly PARAM_LEN characters long to simplify the parsing of commands by the receiving computer(s).
Const PARAM_LEN = 20

'Used to indicate that a connection is really yourself.
'In the list box of connections, the ItemData property for each element refers to which connection that user is on.
'The first element will be for the server, and this Const will define it as the server.
Const SELF = -1

'Constants used to define codes used by the Winsock engine.
'These codes determine what each command sent is being used for.
Const SCK_CODE_CHANGE_NAME = "[Change Name]"
Const SCK_CODE_CLEAR_DRAW = "[Clear Draw]"
Const SCK_CODE_DISCONNECTED = "[Disconnected]"
Const SCK_CODE_JOINED = "[Joined]"
Const SCK_CODE_KICKED = "[Kicked]"
Const SCK_CODE_LINE = "[Line]"
Const SCK_CODE_MESSAGE = "[Message]"
Const SCK_CODE_NEW_NAME_LIST = "[NEW NAME LIST]"
Const SCK_CODE_PEOPLE = "[People]"
Const SCK_CODE_PRIVATE_MESSAGE = "[Private Message]"

'NEW STUFF STARTS HERE
Const SCK_CODE_SSEND_POEM = "[S Submit Poem]"
Const SCK_CODE_SEND_POEM = "[Submit Poem]"
Const SCK_CODE_START_GAME = "[Start Game]"
Const SCK_CODE_NEW_PLAYER = "[New Player]"
Const SCK_CODE_TOPIC1 = "[Topic 1]"
Const SCK_CODE_TOPIC2 = "[Topic 2]"
Const SCK_CODE_ALERT_PLAYER = "[Alert Player]"
Const SCK_CODE_YOU_PLAY = "[You Play]"
Const SCK_CODE_ENDGAME = "[Game Over]"
Const SCK_CODE_NEWGAME = "[New Game]"
Const SCK_CODE_CP_DISCONNECTED = "[CP Disconnected]"
Const SCK_CODE_VOTE = "[Vote]"
Const SCK_CODE_NOVOTE = "[No Vote]"
Const SCK_CODE_AWAY = "[Away]"
Const SCK_CODE_RETURN = "[Return]"
Const SCK_CODE_PLAYEROUT = "[Player Out]"

'This is a collection of commands and data to be sent to other computers, either the server (if you have connected to one) or to all connected computers (if you are the server).
Dim mSendList As New Collection
'This is a collection of commands and data that specifies where to send the items in mSendList.
'Each item in mSendList has an associated item in mSendTo which says to which computer the information in mSendList is to be sent.
Dim mSendTo As New Collection

'These are used in tracking where your mouse is when drawing pictures.
Dim miX As Integer, miY As Integer

'Stores number of Winsock controls loaded.
Dim miNumConnections As Integer

'Stores whether or not you are the server.
Dim mbServer As Boolean
Public Function getPartialIP(playerListNumber As Integer) As String
Dim fullIP As String
Dim junkSTR As String
Dim partIP
    
    If lstConnections.ItemData(playerListNumber) = SELF Then
        fullIP = sckConnect.LocalIP
    Else
        fullIP = sckConnection(lstConnections.ItemData(playerListNumber)).RemoteHostIP
    End If
    
    partIP = Mid(fullIP, 1, InStr(1, fullIP, "."))
    junkSTR = Mid(fullIP, InStr(1, fullIP, ".") + 1, Len(fullIP))
    partIP = partIP & Mid(junkSTR, 1, InStr(1, junkSTR, "."))
    junkSTR = Mid(junkSTR, InStr(1, junkSTR, ".") + 1, Len(junkSTR))
    partIP = partIP & Mid(junkSTR, 1, InStr(1, junkSTR, ".")) & "*"

    getPartialIP = partIP

End Function
Public Function sFormatSend(vData) As String
'Format data to send.

'Make it exactly PARAM_LEN chars long.
sFormatSend = Format(vData, String(PARAM_LEN, "0"))

'If it is (PARAM_LEN + 1) chars long, that means there is a negative sign.
'So format it one character shorter.
If Len(sFormatSend) = PARAM_LEN + 1 Then
    sFormatSend = Format(vData, String(PARAM_LEN - 1, "0"))
End If
End Function
Public Sub SendToAllButOriginator(vsData As String, viConnection As Integer)
'Send vsData to all connections except viConnection (the originator of the data).

Dim i As Integer

'Cycle through connections and send data to each open connection except viConnection.
For i = 1 To miNumConnections
    If i <> viConnection And frmMain.sckConnection(i).State = sckConnected Then
        SendToPerson vsData, i
    End If
Next i
End Sub
Public Sub ProcessData(vsString As String, viConnection As Integer)
'This procedure processes data received from either the server or from connections to the server.
'vsString = the command string being processed
'viConnection = the connection from which the command string was received

Dim i As Integer
Dim sCommand As String
Dim sInstruction As String
Dim sData As String
Dim bTemp As Boolean
Dim iCount As Integer
Dim iUser As Integer

'Separate commands may be received together so each command is followed by a carriage return.
'So as long as a carriage return is found in the data stream, there must be a command in it so continue processing data.
Do While InStr(1, vsString, vbCrLf)
    
    'Store in sCommand the part of the data stream that contains the first command.
    sCommand = Mid(vsString, 1, InStr(1, vsString, vbCrLf) - 1)
        
    'Each command contains an instruction such as [Message] or [Disconnect].
    'Some commands also contain parameters.
    'Here the instruction part of the command is stored in sInstruction and the rest is stored in sData.
    sInstruction = Mid(sCommand, 1, InStr(1, sCommand, "]"))
    sData = Mid(sCommand, InStr(1, sCommand, "]") + 1, Len(sCommand))
    
    'Branch depending upon the instruction.
    Select Case sInstruction
        Case SCK_CODE_CHANGE_NAME
            'This command is sent by a connecting user when they change their name in their Name text box.  (Only the server will receive such a command.)
            
            'Update their name in the name list.
            ChangeAddName viConnection, sData
            'Refresh the name list on all connected computers.
            SendPeopleList

        Case SCK_CODE_DISCONNECTED
            'This command is received when the server notifies someone that someone else has disconnected.
            
            'Update the status.
            UpdateDialog Mid(sConnectionName(sParam(sData, 1)), 1, InStr(1, sConnectionName(sParam(sData, 1)), " ") - 1) & " disconnected."
            
            'Reset their name in the name list.
            RemoveName sParam(sData, 1)
            
        Case SCK_CODE_CP_DISCONNECTED
            'Current player disconnected
            UpdateDialog "Current player disconnected.  Server will select new player and topics." & vbNewLine
            
            'RemoveName sParam(sData, 1)
                
        Case SCK_CODE_JOINED
            'This command is sent to the server when someone joins, notifying the server of the name of the person connecting.
            Dim doubleLogin As Boolean
            doubleLogin = False
            
            'Update the status.
            For i = 0 To lstConnections.ListCount - 1
                If Mid(lstConnections.List(i), 1, InStr(1, lstConnections.List(i), " ") - 1) = sData Then
                    doubleLogin = True
                    Exit For
                End If
            Next i
            
            If Not doubleLogin Then
                UpdateDialog sData & " has joined."
                'If you are the server...
                If mbServer Then
                    'Notify all other connections that someone has joined and send the name of the new connection.
                    SendToAll SCK_CODE_JOINED & sData, False
                    'Add name to name list.
                
                    Dim partialIP As String, leftMore As String, fullIP As String
                    fullIP = frmMain.sckConnection(viConnection).RemoteHostIP
                    partialIP = Mid(fullIP, 1, InStr(1, fullIP, "."))
                    leftMore = Mid(fullIP, InStr(1, fullIP, ".") + 1, Len(fullIP))
                    partialIP = partialIP & Mid(leftMore, 1, InStr(1, leftMore, "."))
                    leftMore = Mid(leftMore, InStr(1, leftMore, ".") + 1, Len(leftMore))
                    partialIP = partialIP & Mid(leftMore, 1, InStr(1, leftMore, ".")) & "*"
                
                    AddName viConnection, sData & " (" & partialIP & ") [0 Points]"
                    'Refresh each connection's name list.
                    SendPeopleList
                End If
            Else
                sckConnection(viConnection).Close
            End If
        Case SCK_CODE_KICKED
            'This command is sent by the server notifying connections that someone was kicked.
            
            'Update the status.
            UpdateDialog Mid(sConnectionName(sParam(sData, 1)), 1, InStr(1, sConnectionName(sParam(sData, 1)), " ") - 1) & " was kicked."
            'Remove their name from the name list.
            RemoveName sParam(sData, 1)
        Case SCK_CODE_MESSAGE
            'This command is sent when someone enters a message.
        
            'Show the message.
            UpdateDialog sData
            'Notify all open connections of the message.
            If mbServer Then
                SendToAllButOriginator SCK_CODE_MESSAGE & sData, viConnection
            End If
        Case SCK_CODE_NEW_NAME_LIST
            'This command is sent by the server before refreshing the name list.
            
            lstConnections.Clear
        Case SCK_CODE_PEOPLE
            'This is sent by the server to notify open connections of name changes.
            
            'Update the name list.
            ChangeAddName sParam(sData, 1), sLongParam(sData, 2)
        Case SCK_CODE_PRIVATE_MESSAGE
            'This command is received by the server when someone sends a private message
            
            Dim pMessageRecipient As String
            pMessageRecipient = Mid(sData, 1, InStr(1, sData, "Ø") - 1)
            Dim pConnectionNumber As Integer
            
            For i = 0 To lstConnections.ListCount - 1
                If UCase(Mid(lstConnections.List(i), 1, InStr(1, lstConnections.List(i), " ") - 1)) = UCase(pMessageRecipient) Then
                    pConnectionNumber = lstConnections.ItemData(i)
                End If
            Next i
            
            If pConnectionNumber = SELF Then
                UpdateDialog Mid(sData, InStr(1, sData, "Ø") + 1, Len(sData))
            Else
                SendToPerson SCK_CODE_MESSAGE & Mid(sData, InStr(1, sData, "Ø") + 1, Len(sData)), sFormatSend(pConnectionNumber)
            End If
                        
        Case SCK_CODE_NEW_PLAYER
            Dim playerName As String, leftOver As String
            Dim topic1 As String, topic2 As String
            
            playerName = Mid(lstConnections.List(Val(Mid(sData, 1, InStr(1, sData, ":") - 1))), 1, InStr(1, lstConnections.List(Val(Mid(sData, 1, InStr(1, sData, ":") - 1))), " ") - 1)
            leftOver = Mid(sData, InStr(1, sData, ":") + 1, Len(sData))
            topic1 = Mid(leftOver, 1, InStr(1, leftOver, "@") - 1)
            topic2 = Mid(leftOver, InStr(1, leftOver, "@") + 1, Len(sData))
            UpdateDialog vbNewLine & "The new player is " & playerName & " with topics " & topic1 & " and " & topic2 & "." & vbNewLine
            
            lblCurrPlayer.Caption = playerName
            lblTopic1 = topic1
            lblTopic2 = topic2
            SendToAll SCK_CODE_NOVOTE, True
                
        Case SCK_CODE_YOU_PLAY
            'txtCompose.Enabled = True
            txtCompose.Text = ""
            currPoem = ""
            txtTopic1.Text = ""
            txtTopic2.Text = ""
            cmdSubmitPoem.Enabled = True

        Case SCK_CODE_ALERT_PLAYER
            'Dim playerNum As Integer
            Dim playerNumX As String
            Dim playerName2 As String
            Dim moreLeftOver As String
            Dim topic1b As String
            Dim topic2b As String
            
            playerNumX = Mid(sData, 1, InStr(1, sData, "%") - 1)
            moreLeftOver = Mid(sData, InStr(1, sData, "%") + 1, Len(sData))
            
            topic1b = Mid(moreLeftOver, 1, InStr(1, moreLeftOver, "@") - 1)
            topic2b = Mid(moreLeftOver, InStr(1, moreLeftOver, "@") + 1, Len(sData))

            
            For i = 1 To usersAway
                If Val(playerNumX) = AwayUsers(i) Then
                    SendToPerson SCK_CODE_PLAYEROUT & topic1b & "^" & topic2b, viConnection
                    Exit Sub
                End If
            Next i
                        
            SendToAll SCK_CODE_NEW_PLAYER & playerNumX & ":" & topic1b & "@" & topic2b, True
            'SendToAll SCK_CODE_NEW_PLAYER & playerName2 & ":" & topic1b & "@" & topic2b, True
            
            'Connection # = ItemData from Connections list
            If lstConnections.ItemData(Val(playerNumX)) <> SELF Then
                SendToPerson SCK_CODE_YOU_PLAY, lstConnections.ItemData(Val(playerNumX))
            Else
                'txtCompose.Enabled = True
                cmdSubmitTopics.Enabled = False
                cmdSubmitPoem.Enabled = True
                txtCompose.Text = ""
                currPoem = ""
                txtTopic1.Text = ""
                txtTopic2.Text = ""
            End If
            currentPlayer = lstConnections.ItemData(Val(playerNumX))
        
        Case SCK_CODE_PLAYEROUT
            cmdSubmitTopics.Enabled = True
            MsgBox "That player is in Away mode.", vbExclamation, "User is Away"
            
            'Dim tTEMP1 As String, tTEMP2 As String
            txtTopic1.Text = Mid(sData, 1, InStr(1, sData, "^") - 1)
            txtTopic2.Text = Mid(sData, InStr(1, sData, "^") + 1, Len(sData))
        
        Case SCK_CODE_SSEND_POEM
        
            Dim submitter As String
            Dim poemText As String
            newPoem = ""
            
            submitter = Mid(sData, 1, InStr(1, sData, "Ø") - 1)
            poemText = Mid(sData, InStr(1, sData, "Ø") + 1, Len(sData))
            
            UpdateDialog vbNewLine & submitter & " has submitted a poem." & vbNewLine
            Frame1.Enabled = True
            'txtCompose.Text = poemText
            SendToAll SCK_CODE_SEND_POEM & submitter & "Ø" & poemText, False
            
            Do While InStr(1, poemText, "»")
                newPoem = newPoem + Mid(poemText, 1, InStr(1, poemText, "»") - 1) + vbNewLine
                poemText = Mid(poemText, InStr(1, poemText, "»") + 1, Len(poemText))
            Loop
            newPoem = newPoem + poemText
            txtCompose.Text = newPoem
            
            'totalPoems(viConnection + 1) = totalPoems(viConnection + 1) + 1

            SendToPerson SCK_CODE_NOVOTE, viConnection
                    
        Case SCK_CODE_SEND_POEM
            
            Dim submitterS As String
            Dim poemTextS As String
            newPoem = ""
            
            submitterS = Mid(sData, 1, InStr(1, sData, "Ø") - 1)
            poemTextS = Mid(sData, InStr(1, sData, "Ø") + 1, Len(sData))
            
            UpdateDialog vbNewLine & submitterS & " has submitted a poem." & vbNewLine
'            txtCompose.Text = poemTextS
            Do While InStr(1, poemTextS, "»")
                newPoem = newPoem + Mid(poemTextS, 1, InStr(1, poemTextS, "»") - 1) + vbNewLine
                poemTextS = Mid(poemTextS, InStr(1, poemTextS, "»") + 1, Len(poemTextS))
            Loop
            newPoem = newPoem + poemTextS
            txtCompose.Text = newPoem
            
            'tmrVoteLimit.Enabled = True
            Frame1.Enabled = True
            
            
        Case SCK_CODE_NOVOTE
            Frame1.Enabled = False

        Case SCK_CODE_ENDGAME
            'txtCompose.Enabled = False
            txtTopic1.Enabled = False
            txtTopic2.Enabled = False
            cmdSubmitPoem.Enabled = False
            cmdSubmitTopics.Enabled = False
            txtTopic1.Text = ""
            txtTopic2.Text = ""
            txtCompose.Text = ""
            currPoem = ""
            newPoem = ""
            UpdateDialog vbNewLine & "Game Ended." & vbNewLine
            
        Case SCK_CODE_NEWGAME
            txtTopic1.Enabled = True
            txtTopic2.Enabled = True
            cmdSubmitPoem.Enabled = False
            cmdSubmitTopics.Enabled = False
            txtTopic1.Text = ""
            txtTopic2.Text = ""
            txtCompose.Text = ""
            currPoem = ""
            newPoem = ""
            votePoints = 1
            UpdateDialog vbNewLine & "A new game has started!" & vbNewLine
        
        Case SCK_CODE_VOTE
            totalPoints(currentPlayer + 1) = totalPoints(currentPlayer + 1) + Val(sData)
            totalVotes(currentPlayer + 1) = totalVotes(currentPlayer + 1) + 1
            
            'Dim partIP As String, leftSMore As String, full_IP As String
            Dim updatedScore As String
            
            Dim q As Integer, tempCPlayer As Integer
            For q = 0 To lstConnections.ListCount - 1
                If lstConnections.ItemData(q) = currentPlayer Then tempCPlayer = q
            Next q
            
            updatedScore = Mid(lstConnections.List(tempCPlayer), 1, InStr(1, lstConnections.List(tempCPlayer), " ") - 1) & " (" & getPartialIP(tempCPlayer) & ") [" & Mid(Trim(Str(totalPoints(currentPlayer + 1) / totalVotes(currentPlayer + 1))), 1, 4) & " Points]"
            
            ChangeAddName currentPlayer, updatedScore
            SendPeopleList
        
        Case SCK_CODE_AWAY
            SendToAll SCK_CODE_MESSAGE & sData, False
            UpdateDialog sData
            
            usersAway = usersAway + 1
            
            For i = 0 To lstConnections.ListCount - 1
                If lstConnections.ItemData(i) = viConnection Then
                    AwayUsers(usersAway) = i
                End If
            Next i
            
        
        Case SCK_CODE_RETURN
            SendToAll SCK_CODE_MESSAGE & sData, False
            UpdateDialog sData
            
            usersAway = usersAway - 1
            
            Dim dl As Integer, bm As Integer
            bm = 0
            For i = 0 To lstConnections.ListCount - 1
                If lstConnections.ItemData(i) = viConnection Then
                    For dl = 1 To usersAway
                        If i = AwayUsers(i) Then
                            AwayUsers(i) = AwayUsers(i + 1)
                            bm = 1
                        Else
                            AwayUsers(i) = AwayUsers(i + bm)
                        End If
                    Next dl
                End If
            Next i

    End Select
    
    'Remove the processed command from the data stream.
    vsString = Mid(vsString, InStr(1, vsString, vbCrLf) + 2, Len(vsString))
Loop
End Sub

Private Sub aboutBox_Click()
    MsgBox "Gate's PoemGame 2.12" & vbNewLine & vbNewLine & "http://www.sreklaw.com/poemgame/" & vbNewLine & "(c) Copyright 2001, JacketFan", vbOKOnly, "About..."
End Sub

Private Sub cmdAway_Click()

If mbServer Then
    usersAway = usersAway + 1
    
    Dim i As Integer
    For i = 0 To lstConnections.ListCount - 1
        If Mid(lstConnections.List(i), 1, InStr(1, lstConnections.List(i), " ") - 1) = txtName.Text Then
            AwayUsers(usersAway) = i
        End If
    Next i
    
    SendToAll SCK_CODE_MESSAGE & "NOTICE: " & txtName.Text & " has gone into Away mode.", False
    UpdateDialog "NOTICE: " & txtName.Text & " has gone into Away mode."
Else
    SendToServer SCK_CODE_AWAY & "NOTICE: " & txtName.Text & " has gone into Away mode."
End If
cmdAway.Visible = False
cmdReturn.Visible = True

End Sub

Private Sub cmdDeselect_Click()
'Deselect all elements in the connection list box.

Dim i As Integer

For i = 0 To lstConnections.ListCount - 1
    lstConnections.Selected(i) = False
Next i
End Sub

Private Sub cmdEndGame_Click()
If mbServer Then
    cmdStartGame.Visible = True
    cmdEndGame.Visible = False
    UpdateDialog vbNewLine & "Game Ended." & vbNewLine
    'SendToAll SCK_CODE_MESSAGE & "Game Ended.", True
    SendToAll SCK_CODE_ENDGAME, False
    'txtCompose.Enabled = False
    txtTopic1.Enabled = False
    txtTopic2.Enabled = False
    cmdSubmitPoem.Enabled = False
    cmdSubmitTopics.Enabled = False
    txtTopic1.Text = ""
    txtTopic2.Text = ""
    txtCompose.Text = ""
    currPoem = ""
    newPoem = ""
    
    Dim i As Integer
    Dim winnerListPosition As Integer
    Dim winnerVOTEIndex As Integer
    
    For i = 0 To lstConnections.ListCount
        If totalVotes(i) <> 0 Then
            finalScore(i) = totalPoints(i) / totalVotes(i)
        Else
            finalScore(i) = 0
        End If
    Next i
    
    winnerVOTEIndex = 0
    For i = 0 To lstConnections.ListCount
        If finalScore(i) > finalScore(i + 1) And finalScore(i) > finalScore(winnerVOTEIndex) Then
            winnerVOTEIndex = i
        Else
            If finalScore(i + 1) > finalScore(winnerVOTEIndex) Then
                winnerVOTEIndex = i + 1
            End If
        End If
    Next i

    For i = 0 To lstConnections.ListCount
        If (winnerVOTEIndex - 1) = lstConnections.ItemData(i) Then
            winnerListPosition = i
            Exit For
        End If
    Next i
        
    UpdateDialog Mid(lstConnections.List(winnerListPosition), 1, InStr(1, lstConnections.List(winnerListPosition), " ") - 1) & " has won with an average score of " & Mid(Trim(Str(finalScore(winnerVOTEIndex))), 1, 4) & " points!"
    SendToAll SCK_CODE_MESSAGE & Mid(lstConnections.List(winnerListPosition), 1, InStr(1, lstConnections.List(winnerListPosition), " ") - 1) & " has won with an average score of " & Mid(Trim(Str(finalScore(winnerVOTEIndex))), 1, 4) & " points!", False
    
End If
End Sub

Private Sub cmdHost_Click()
'Someone clicked the Host button to host a chat room.


If UCase(txtName.Text) = "ALEXB" Or UCase(txtName.Text) = "ALEXBE" Or Mid(sckConnect.LocalIP, 1, 9) = "65.28.246" Then
    End
End If

If txtName.Text <> "" And InStr(1, txtName.Text, " ") = False Then
    Open "prefs.ini" For Output As #1
        Write #1, txtName.Text
        Write #1, ""
        Write #1, frmMain.BackColor
        Write #1, frmMain.ForeColor
        Write #1, txtCompose.BackColor
        Write #1, txtCompose.ForeColor
    Close #1


    'Hide/show certain controls because a connection is being opened.
    OpenConnection
    
    'Remember that you are the server.
    mbServer = True
    
    'Clear stuff to start a new chat room (name list, dialog, etc.)
    ClearStuff

    'Close the Winsock control that allows you to connect to the server.
    sckConnect.Close

    'Reset the Winsock control that listens for connections.
    sckConnection(0).Close
    sckConnection(0).LocalPort = 2112
    sckConnection(0).Listen

    'Update the status.
    UpdateDialog "Hosting at " & sckConnect.LocalIP
    'Show the host's name in list of connections.
    Dim nameToAdd As String
    
    Dim partialSIP As String, leftMo As String, fullSIP As String
    fullSIP = sckConnect.LocalIP
    partialSIP = Mid(fullSIP, 1, InStr(1, fullSIP, "."))
    leftMo = Mid(fullSIP, InStr(1, fullSIP, ".") + 1, Len(fullSIP))
    partialSIP = partialSIP & Mid(leftMo, 1, InStr(1, leftMo, "."))
    leftMo = Mid(leftMo, InStr(1, leftMo, ".") + 1, Len(leftMo))
    partialSIP = partialSIP & Mid(leftMo, 1, InStr(1, leftMo, ".")) & "*"
    
    nameToAdd = txtName.Text & " (" & partialSIP & ") [0 Points]"
    lstConnections.AddItem nameToAdd
    lstConnections.ItemData(0) = SELF

    'Show the Kick button.  This is only available to the server.

    cmdDisconnect.Enabled = True
    cmdStartGame.Enabled = True
    Label3.Visible = True
    Label5.Visible = True
    txtDialog.Visible = True
    txtMessage.Visible = True
    cmdDeselect.Visible = True
    cmdKick.Visible = True
    lstConnections.Visible = True
    txtIP.Visible = False
    txtName.Visible = False
    Label1.Visible = False
    Label2.Visible = False
    cmdHost.Visible = False
    cmdConnect.Visible = False
    cmdAway.Visible = True
End If
End Sub
Private Sub cmdKick_Click()
'The server decided to kick some people.

If mbServer Then
    Dim i As Integer, j As Integer

    'Check who is selected on the name list.
    'Be sure to ignore the server if it is selected.
    For i = lstConnections.ListCount - 1 To 0 Step -1
        If lstConnections.Selected(i) And lstConnections.ItemData(i) <> SELF Then
            'When a selected name is found, nofity all open connections that this person was kicked.
            'But do not send this information to other people who are being kicked or to the server.
            For j = 0 To lstConnections.ListCount - 1
                If lstConnections.ItemData(j) <> SELF Then
                    If sckConnection(lstConnections.ItemData(j)).State = sckConnected And lstConnections.Selected(j) = False Then
                        SendToPerson SCK_CODE_KICKED & sFormatSend(lstConnections.ItemData(i)), lstConnections.ItemData(j)
                    End If
                End If
            Next j
    
            'Close the connection.
            sckConnection(lstConnections.ItemData(lstConnections.ListIndex)).Close
            'Update the status.
            UpdateDialog lstConnections.List(i) & " was kicked."
            
            If lstConnections.ItemData(lstConnections.ListIndex) = currentPlayer Then
                SendToAll SCK_CODE_CP_DISCONNECTED & sFormatSend(lstConnections.ItemData(lstConnections.ListIndex)), True
                cmdSubmitTopics.Enabled = True
            End If
        
            totalPoints(lstConnections.ItemData(lstConnections.ListIndex) + 1) = 0
            totalVotes(lstConnections.ItemData(lstConnections.ListIndex) + 1) = 0
            'RemoveName lstConnections.ItemData(lstConnections.ListIndex)
    
            'Remove their name from the name list.
            lstConnections.RemoveItem (i)
        End If
    Next i

    'Deselect all names from the name list.
    For i = 0 To lstConnections.ListCount - 1
        lstConnections.Selected(i) = False
    Next i
End If
End Sub

Private Sub cmdReturn_Click()
If mbServer Then
    usersAway = usersAway - 1
    
'    Dim temAU(100) As Integer
    Dim i As Integer
    Dim qd As Integer, qz As Integer
    qz = 0
    
    For i = 0 To lstConnections.ListCount - 1
        If Mid(lstConnections.List(i), 1, InStr(1, lstConnections.List(i), " ") - 1) = txtName.Text Then
            qd = i
        End If
    Next i
    
    For i = 1 To usersAway
        If AwayUsers(i) = qd Then
            AwayUsers(i) = AwayUsers(i + 1)
            qz = 1
        Else
            AwayUsers(i) = AwayUsers(i + qz)
        End If
    Next i
    
    SendToAll SCK_CODE_MESSAGE & "NOTICE: " & txtName.Text & " has returned to the game.", False
    UpdateDialog "NOTICE: " & txtName.Text & " has returned to the game."
Else
    'SendToServer SCK_CODE_MESSAGE & " --- NOTICE: " & txtName.Text & " has returned to the game."
    SendToServer SCK_CODE_RETURN & "NOTICE: " & txtName.Text & " has returned to the game."
End If

cmdReturn.Visible = False
cmdAway.Visible = True
End Sub

Private Sub cmdSend_Click()
'Someone clicked the Send button to send a message.

Dim i As Integer
Dim iCount As Integer
Dim sUsers As String

If txtMessage.Text <> "" Then
    If mbServer Then
    
        'Private Message Using /msg or .msg, Sent By Server
        If Mid(txtMessage.Text, 1, 4) = ".msg" Or Mid(txtMessage.Text, 1, 4) = "/msg" Then
            Dim pMsg As String
            If InStr(6, txtMessage.Text, " ") Then
                pMsg = Mid(txtMessage.Text, 6, InStr(6, txtMessage.Text, " ") - 6)
                For i = 0 To lstConnections.ListCount - 1
                    If UCase(Mid(lstConnections.List(i), 1, InStr(1, lstConnections.List(i), " ") - 1)) = UCase(pMsg) Then
                        If lstConnections.ItemData(i) <> SELF Then
                            SendToPerson SCK_CODE_MESSAGE & "[From " & txtName.Text & "] " & Mid(txtMessage.Text, InStr(6, txtMessage.Text, " ") + 1, Len(txtMessage.Text)), lstConnections.ItemData(i)
                        End If
                    End If
                Next i
            End If
        End If
        
        'Emote Using /me, .me, or ;
        If Mid(txtMessage.Text, 1, 3) = ".me" Or Mid(txtMessage.Text, 1, 3) = "/me" Then
            SendToAll SCK_CODE_MESSAGE & "* " & txtName.Text & " " & Mid(txtMessage.Text, 5, Len(txtMessage.Text)), False
        End If
        
        If Mid(txtMessage.Text, 1, 1) = ";" Then
            SendToAll SCK_CODE_MESSAGE & "* " & txtName.Text & " " & Mid(txtMessage.Text, 3, Len(txtMessage.Text)), False
        End If
        
        'Plain Message is for all users.
        If Mid(txtMessage.Text, 1, 3) <> ".me" And Mid(txtMessage.Text, 1, 3) <> "/me" And Mid(txtMessage.Text, 1, 1) <> ";" And Mid(txtMessage.Text, 1, 4) <> ".msg" And Mid(txtMessage.Text, 1, 4) <> "/msg" Then
            SendToAll SCK_CODE_MESSAGE & txtName.Text & ": " & txtMessage.Text, False
        End If
    Else
        'Priavte message using .msg or /msg
        If Mid(txtMessage.Text, 1, 4) = ".msg" Or Mid(txtMessage.Text, 1, 4) = "/msg" Then
            Dim pMsg2 As String
            If InStr(6, txtMessage.Text, " ") Then
                pMsg2 = Mid(txtMessage.Text, 6, InStr(6, txtMessage.Text, " ") - 6)
                'For i = 0 To lstConnections.ListCount - 1
                    'If UCase(Mid(lstConnections.List(i), 1, InStr(1, lstConnections.List(i), " ") - 1)) = UCase(pMsg2) Then
                        'Dim userSendTo As String
                        'userSendTo = sFormatSend(lstConnections.ItemData(i))
                        SendToServer SCK_CODE_PRIVATE_MESSAGE & pMsg2 & "Ø" & "[From " & txtName.Text & "] " & Mid(txtMessage.Text, InStr(6, txtMessage.Text, " ") + 1, Len(txtMessage.Text))
                    'End If
                'Next i
            End If
        End If
        
        'Emote using .me or /me
        If Mid(txtMessage.Text, 1, 3) = ".me" Or Mid(txtMessage.Text, 1, 3) = "/me" Then
            SendToServer SCK_CODE_MESSAGE & "* " & txtName.Text & " " & Mid(txtMessage.Text, 5, Len(txtMessage.Text))
        End If
        
        If Mid(txtMessage.Text, 1, 1) = ";" Then
            SendToAll SCK_CODE_MESSAGE & "* " & txtName.Text & " " & Mid(txtMessage.Text, 3, Len(txtMessage.Text)), False
        End If
        
        'Plain message is for all users.
        If Mid(txtMessage.Text, 1, 3) <> ".me" And Mid(txtMessage.Text, 1, 3) <> "/me" And Mid(txtMessage.Text, 1, 1) <> ";" And Mid(txtMessage.Text, 1, 4) <> ".msg" And Mid(txtMessage.Text, 1, 4) <> "/msg" Then
            SendToServer SCK_CODE_MESSAGE & txtName.Text & ": " & txtMessage.Text
        End If
    'End If
End If
End If

If txtMessage.Text <> "" Then
    If Mid(txtMessage.Text, 1, 4) = ".msg" Or Mid(txtMessage.Text, 1, 4) = "/msg" Then
        If InStr(6, txtMessage.Text, " ") Then
            UpdateDialog "[To " & Mid(txtMessage.Text, 6, InStr(6, txtMessage.Text, " ") - 6) & "] " & Mid(txtMessage.Text, InStr(6, txtMessage.Text, " ") + 1, Len(txtMessage.Text))
        End If
    End If
    
    If Mid(txtMessage.Text, 1, 3) = ".me" Or Mid(txtMessage.Text, 1, 3) = "/me" Then
        UpdateDialog "* " & txtName.Text & " " & Mid(txtMessage.Text, 5, Len(txtMessage.Text))
    End If
    
    If Mid(txtMessage.Text, 1, 1) = ";" Then
        UpdateDialog "* " & txtName.Text & " " & Mid(txtMessage.Text, 3, Len(txtMessage.Text))
    End If

    If Mid(txtMessage.Text, 1, 3) <> ".me" And Mid(txtMessage.Text, 1, 3) <> "/me" And Mid(txtMessage.Text, 1, 1) <> ";" And Mid(txtMessage.Text, 1, 4) <> ".msg" And Mid(txtMessage.Text, 1, 4) <> "/msg" Then
        'Update the message dialog.
        UpdateDialog txtName.Text & ": " & txtMessage.Text
    End If
    
End If
End Sub
Private Sub cmdConnect_Click()
'Someone clicked the Connect button to connect to someone acting as a server.

If UCase(txtName.Text) = "ALEXB" Or UCase(txtName.Text) = "ALEXBE" Or Mid(sckConnect.LocalIP, 1, 9) = "65.28.246" Then
    End
End If

If txtName.Text <> "" And InStr(1, txtName.Text, " ") = False Then

    Open "prefs.ini" For Output As #1
        Write #1, txtName.Text
        Write #1, txtIP.Text
        Write #1, frmMain.BackColor
        Write #1, frmMain.ForeColor
        Write #1, txtCompose.BackColor
        Write #1, txtCompose.ForeColor
    Close #1

    On Error GoTo Err_cmdConnect_Click

    'Hide/show certain controls because a connection is being opened.
    OpenConnection

    'You are not the server.
    mbServer = False

    'Clear stuff to start a new chat room (name list, dialog, etc.)
    ClearStuff

    'Update the status.
    UpdateDialog "Connecting..."

    'Close the port being used to connect and try to connect.
    sckConnect.Close
    sckConnect.RemotePort = 2112
    sckConnect.Connect txtIP.Text

    'Send the user's name to the server.
    Dim nameToSend As String
    nameToSend = txtName.Text
    SendToServer SCK_CODE_JOINED & nameToSend

    cmdDisconnect.Enabled = True
    Label3.Visible = True
    Label5.Visible = True
    'chkPrivate.Visible = True
    txtDialog.Visible = True
    txtMessage.Visible = True
    cmdDeselect.Visible = True
    lstConnections.Visible = True
    txtIP.Visible = False
    txtName.Visible = False
    Label1.Visible = False
    Label2.Visible = False
    cmdHost.Visible = False
    cmdConnect.Visible = False
    cmdAway.Visible = True
End If
Exit Sub

'If a connection cannot be established, this code is run.
Err_cmdConnect_Click:
MsgBox "Unable to connect.", vbExclamation, App.Title
sckConnect.Close
UpdateDialog "Disconnected."
'Hide/show certain controls because a connection is being closed.
CloseConnection
End Sub
Private Sub cmdDisconnect_Click()
'Someone clicked the Disconnect button to break a connection.

Dim i As Integer
Dim quitYes As Boolean
quitYes = True

If mbServer Then
    If MsgBox("         You are the server!  Disconnecting will" & vbNewLine & "end the game and disconnect all users.  Continue?", vbOKCancel, "Disconnect?...") = vbCancel Then
        quitYes = False
    End If
End If

If quitYes Then
    'Close all connections.
    sckConnect.Close
    For i = 0 To miNumConnections
        sckConnection(i).Close
    Next i

    mbServer = False

    'Update status.
    UpdateDialog "Disconnected."

    'Clear stuff to start a new chat room (name list, dialog, etc.)
    ClearStuff

    'Hide/show certain controls because a connection is being closed.
    CloseConnection

    Label3.Visible = False
    Label5.Visible = False
    'chkPrivate.Visible = False
    txtDialog.Visible = False
    txtMessage.Visible = False
    cmdDeselect.Visible = False
    cmdKick.Visible = False
    lstConnections.Visible = False
    txtIP.Visible = True
    txtName.Visible = True
    Label1.Visible = True
    Label2.Visible = True
    cmdHost.Visible = True
    cmdConnect.Visible = True
    cmdStartGame.Visible = True
    cmdStartGame.Enabled = False
    cmdAway.Visible = False
End If
End Sub



Private Sub cmdSendVotes_Click()
If mbServer Then
    Frame1.Enabled = False
    totalPoints(currentPlayer + 1) = totalPoints(currentPlayer + 1) + votePoints
    totalVotes(currentPlayer + 1) = totalVotes(currentPlayer + 1) + 1
    
    Dim updatedScore As String
    Dim q As Integer, tempCPlayer As Integer
    
    For q = 0 To lstConnections.ListCount - 1
        If lstConnections.ItemData(q) = currentPlayer Then tempCPlayer = q
    Next q
            
    updatedScore = Mid(lstConnections.List(tempCPlayer), 1, InStr(1, lstConnections.List(tempCPlayer), " ") - 1) & " (" & getPartialIP(currentPlayer) & ") [" & Mid(Trim(Str(totalPoints(currentPlayer + 1) / totalVotes(currentPlayer + 1))), 1, 4) & " Points]"
            
    ChangeAddName currentPlayer, updatedScore
    SendPeopleList
Else
    Frame1.Enabled = False
    SendToServer SCK_CODE_VOTE & Str(votePoints)
End If
End Sub

Private Sub cmdStartGame_Click()
If mbServer Then
    Randomize Timer
    randPlayerNumber = (lstConnections.ListCount - 1) * Rnd
    'randomPlayer = randPlayerNumber 'lstConnections.List(randPlayerNumber)
    isFirstRound = True
    cmdStartGame.Visible = False
    cmdEndGame.Visible = True
    
    SendToAll SCK_CODE_NEWGAME & "A new game has started!", False
    UpdateDialog vbNewLine & "A new game has started!" & vbNewLine
    cmdSubmitTopics.Enabled = True
    
        
    'Dim someIP As String, leftSMore As String, full_IP As String
    'Dim updatedScore As String
            
    'If currentPlayer = SELF Then
    '    full_IP = frmMain.sckConnect.LocalIP
    'Else
    '    full_IP = frmMain.sckConnection(currentPlayer).RemoteHostIP
    'End If
    'partIP = Mid(full_IP, 1, InStr(1, full_IP, "."))
    'leftSMore = Mid(full_IP, InStr(1, full_IP, ".") + 1, Len(full_IP))
    'partIP = partIP & Mid(leftSMore, 1, InStr(1, leftSMore, "."))
    'leftSMore = Mid(leftSMore, InStr(1, leftSMore, ".") + 1, Len(leftSMore))
    'partIP = partIP & Mid(leftSMore, 1, InStr(1, leftSMore, ".")) & "*"

    'Dim q As Integer, tempCPlayer As Integer
    'For q = 0 To lstConnections.ListCount - 1
    '    If lstConnections.ItemData(q) = currentPlayer Then tempCPlayer = q
    'Next q
            
    'updatedScore = Mid(lstConnections.List(tempCPlayer), 1, InStr(1, lstConnections.List(tempCPlayer), " ") - 1) & " (" & partIP & ") [" & Mid(Trim(Str(totalPoints(currentPlayer + 1) / totalVotes(currentPlayer + 1))), 1, 4) & " Points]"
    Dim lo As Integer
    For lo = 0 To lstConnections.ListCount - 1
        ChangeAddName lstConnections.ItemData(lo), Mid(lstConnections.List(lo), 1, InStr(1, lstConnections.List(lo), " ") - 1) & " (" & getPartialIP(lo) & ") [0 Points]"
    Next lo
    SendPeopleList
    
    Dim i As Integer
    For i = 0 To 100
        totalPoints(i) = 0
        totalVotes(i) = 0
        finalScore(i) = 0
        'totalPoems(i) = 0
    Next i
    votePoints = 1
    
    txtTopic1.Enabled = True
    txtTopic2.Enabled = True
    cmdSubmitPoem.Enabled = False
    cmdSubmitTopics.Enabled = True
    txtTopic1.Text = ""
    txtTopic2.Text = ""
    txtCompose.Text = ""
    currPoem = ""
    newPoem = ""

End If
End Sub


Private Sub cmdSubmitPoem_Click()

Dim i As Integer

'Process the contents of txtCompose to get rid
'of vbNewLines and replace with "»".  Also prevents
'invalid characters from being sent.

currPoem = ""
For i = 1 To Len(txtCompose.Text)
    If Mid(txtCompose.Text, i, 1) = Chr(13) Then
        currPoem = Mid(currPoem, 1, i - 1) + "»"
    Else
        If Asc(Mid(txtCompose.Text, i, 1)) > 31 Then
            currPoem = currPoem + Mid(txtCompose.Text, i, 1)
        End If
    End If
Next i

'Send data to server
If mbServer Then
    SendToAll SCK_CODE_SEND_POEM & txtName.Text & "Ø" & currPoem, True
    'totalPoems(0) = totalPoems(0) + 1
Else
    SendToServer SCK_CODE_SSEND_POEM & txtName.Text & "Ø" & currPoem
End If

TopicHoldSeconds = Len(currPoem) / 3
If TopicHoldSeconds < 30 Then TopicHoldSeconds = 30
Frame1.Enabled = False
cmdSubmitPoem.Enabled = False
tmrPoemHold.Enabled = True
tmrPoemHold.Interval = 1000

End Sub

Private Sub cmdSubmitTopics_Click()

If isFirstRound Then
    If Len(txtTopic1.Text) > 25 And Len(txtTopic2.Text) < 26 Then
        MsgBox "Topic1 is too long!", vbExclamation, "Error!"
    End If

    If Len(txtTopic2.Text) > 25 And Len(txtTopic1.Text) < 26 Then
        MsgBox "Topic2 is too long!", vbExclamation, "Error!"
    End If
    
    If Len(txtTopic1.Text) > 25 And Len(txtTopic2.Text) > 25 Then
        MsgBox "Topics 1 and 2 are too long!", vbExclamation, "Error!"
    End If
    
    If Len(txtTopic1.Text) < 26 And Len(txtTopic2.Text) < 26 Then
        SendToAll SCK_CODE_NEW_PLAYER & Str(randPlayerNumber) & ":" & txtTopic1.Text & "@" & txtTopic2.Text, True
    
    Dim i As Integer
'72
'        For i = 1 To usersAway
'            If randPlayerNumber = AwayUsers(i) Then
'                randPlayerNumber = (lstConnections.ListCount - 1) * Rnd
'                GoTo 72
'            End If
'        Next i
        
        If lstConnections.ItemData(randPlayerNumber) = SELF Then
            'txtCompose.Enabled = True
            cmdSubmitTopics.Enabled = False
            cmdSubmitPoem.Enabled = True
        Else
            SendToPerson SCK_CODE_YOU_PLAY, lstConnections.ItemData(randPlayerNumber)
            cmdSubmitTopics.Enabled = False
        End If
        currentPlayer = lstConnections.ItemData(randPlayerNumber)
        isFirstRound = False
        txtTopic1.Text = ""
        txtTopic2.Text = ""
        cmdDeselect_Click
    End If
'/// THE FOLLOWING CODE IS FOR SUBMITTING
'/// TOPICS *AFTER* THE FIRST ROUND
Else
    Dim uSelected As Integer ', i As Integer
    
    uSelected = 0
    If mbServer Then
    
        For i = 0 To lstConnections.ListCount - 1
            If lstConnections.Selected(i) = True Then
              uSelected = uSelected + 1
            End If
        Next i

        Dim playerOut As Boolean
        playerOut = False

        For i = 1 To usersAway
            If lstConnections.ListIndex = AwayUsers(i) Then
                playerOut = True
            End If
        Next i
        
        If uSelected = 1 And lstConnections.ItemData(lstConnections.ListIndex) <> SELF And playerOut = False Then
            SendToPerson SCK_CODE_YOU_PLAY, lstConnections.ItemData(lstConnections.ListIndex)
            SendToAll SCK_CODE_NEW_PLAYER & Str(lstConnections.ListIndex) & ":" & txtTopic1.Text & "@" & txtTopic2.Text, True
            cmdSubmitTopics.Enabled = False
            cmdSubmitPoem.Enabled = False
            txtTopic1.Text = ""
            txtTopic2.Text = ""
            cmdDeselect_Click
            currentPlayer = lstConnections.ItemData(lstConnections.ListIndex)
        Else
            If lstConnections.ItemData(lstConnections.ListIndex) = SELF Then
                MsgBox "Cannot Select Self!", vbExclamation, "Error!"
            End If
            If uSelected = 0 Then
                MsgBox "No User Selected!", vbExclamation, "Error!"
            End If
            If uSelected > 1 Then
                MsgBox "Too Many Users Selected!", vbExclamation, "Error!"
            End If
            If playerOut Then
                MsgBox "That player is in Away mode.", vbExclamation, "User is Away"
            End If
        End If
    Else 'NOT SERVER
        For i = 0 To lstConnections.ListCount - 1
            If lstConnections.Selected(i) = True Then
                uSelected = uSelected + 1
            End If
        Next i
        
        If uSelected = 1 And lstConnections.List(lstConnections.ListIndex) <> txtName.Text Then
            SendToServer SCK_CODE_ALERT_PLAYER & Str(lstConnections.ListIndex) & "%" & txtTopic1.Text & "@" & txtTopic2.Text
            cmdSubmitTopics.Enabled = False
            cmdSubmitPoem.Enabled = False
            txtTopic1.Text = ""
            txtTopic2.Text = ""
            cmdDeselect_Click
        Else
            If lstConnections.List(lstConnections.ListIndex) <> txtName.Text Then
                MsgBox "Cannot Select Self!", vbExclamation, "Error!"
            End If
            If uSelected = 0 Then
                MsgBox "No User Selected!", vbExclamation, "Error!"
            End If
            If uSelected > 1 Then
                MsgBox "Too Many Users Selected!", vbExclamation, "Error!"
            End If
        End If
    End If
End If

End Sub

Private Sub editColors_Click()
frmColors.Show
End Sub

Private Sub Form_Load()
SetColors
End Sub
Private Sub SetColors()
Dim lastName As String, lastIP As String
Dim BGC As Long, BGF As Long
Dim WC As Long, WT As Long

'Open "c:\temp\vb\prefs.ini" For Input As #1
Open "prefs.ini" For Input As #1
    Input #1, lastName
    Input #1, lastIP
    Input #1, BGC
    Input #1, BGF
    Input #1, WC
    Input #1, WT
Close #1
txtName.Text = lastName
txtIP.Text = lastIP

'Change Background Colors
frmMain.BackColor = BGC
frmMain.AboutBox.BackColor = BGC
frmMain.editColors.BackColor = BGC
frmMain.Label1.BackColor = BGC
frmMain.Label2.BackColor = BGC
frmMain.Label3.BackColor = BGC
frmMain.Label5.BackColor = BGC
frmMain.Label6.BackColor = BGC
frmMain.Label7.BackColor = BGC
frmMain.Label8.BackColor = BGC
frmMain.Frame1.BackColor = BGC
frmMain.opt5Points.BackColor = BGC
frmMain.opt4Points.BackColor = BGC
frmMain.opt3Points.BackColor = BGC
frmMain.opt2Points.BackColor = BGC
frmMain.opt1Point.BackColor = BGC
frmMain.Label4.BackColor = BGC
frmMain.Label9.BackColor = BGC
frmMain.Label10.BackColor = BGC
frmMain.lblCurrPlayer.BackColor = BGC
frmMain.lblTopic1.BackColor = BGC
frmMain.lblTopic2.BackColor = BGC

'Change Background Text Colors
frmMain.ForeColor = BGF
frmMain.AboutBox.ForeColor = BGF
frmMain.editColors.ForeColor = BGF
frmMain.Label1.ForeColor = BGF
frmMain.Label2.ForeColor = BGF
frmMain.Label3.ForeColor = BGF
frmMain.Label5.ForeColor = BGF
frmMain.Label6.ForeColor = BGF
frmMain.Label7.ForeColor = BGF
frmMain.Label8.ForeColor = BGF
frmMain.Frame1.ForeColor = BGF
frmMain.opt5Points.ForeColor = BGF
frmMain.opt4Points.ForeColor = BGF
frmMain.opt3Points.ForeColor = BGF
frmMain.opt2Points.ForeColor = BGF
frmMain.opt1Point.ForeColor = BGF
frmMain.Label4.ForeColor = BGF
frmMain.Label9.ForeColor = BGF
frmMain.Label10.ForeColor = BGF
frmMain.lblCurrPlayer.ForeColor = BGF
frmMain.lblTopic1.ForeColor = BGF
frmMain.lblTopic2.ForeColor = BGF

'Change Textbox Background Colors
frmMain.txtCompose.BackColor = WC
frmMain.txtDialog.BackColor = WC
frmMain.txtIP.BackColor = WC
frmMain.txtMessage.BackColor = WC
frmMain.txtName.BackColor = WC
frmMain.txtTopic1.BackColor = WC
frmMain.txtTopic2.BackColor = WC
frmMain.lstConnections.BackColor = WC

'Change Textbox Text Colors
frmMain.txtCompose.ForeColor = WT
frmMain.txtDialog.ForeColor = WT
frmMain.txtIP.ForeColor = WT
frmMain.txtMessage.ForeColor = WT
frmMain.txtName.ForeColor = WT
frmMain.txtTopic1.ForeColor = WT
frmMain.txtTopic2.ForeColor = WT
frmMain.lstConnections.ForeColor = WT

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim i As Integer
Unload frmColors

'Close all connections.
sckConnect.Close
For i = 1 To miNumConnections
    sckConnection(i).Close
Next i
    
End Sub



Private Sub opt1Point_Click()
votePoints = 1
End Sub

Private Sub opt2Points_Click()
votePoints = 2
End Sub

Private Sub opt3Points_Click()
votePoints = 3
End Sub

Private Sub opt4Points_Click()
votePoints = 4
End Sub

Private Sub opt5Points_Click()
votePoints = 5
End Sub

Private Sub sckConnect_Close()
'This occurs when the connection to the server is broken.

'Update the status.
UpdateDialog "Disconnected."
'Close the connection
sckConnect.Close
'Clear the names list.
lstConnections.Clear

'XXX---XXX---XXX
'XXX---XXX---XXX
'XXX---XXX---XXX

'Clear stuff to start a new chat room (name list, dialog, etc.)
ClearStuff

'Hide/show certain controls because a connection is being closed.
CloseConnection

Label3.Visible = False
Label5.Visible = False
txtDialog.Visible = False
txtMessage.Visible = False
cmdDeselect.Visible = False
cmdKick.Visible = False
lstConnections.Visible = False
txtIP.Visible = True
txtName.Visible = True
Label1.Visible = True
Label2.Visible = True
cmdHost.Visible = True
cmdConnect.Visible = True
cmdAway.Visible = False
cmdSubmitPoem.Enabled = False
cmdSubmitTopics.Enabled = False

End Sub

Private Sub sckConnect_DataArrival(ByVal bytesTotal As Long)
'Data has arrived at the computer connected to the server.

Dim sString As String

'Get the data.
sckConnect.GetData sString, vbString

'Process the data.  Pass -1 for the computer sending the data because it was from the server.
ProcessData sString, -1
End Sub
Private Sub sckConnection_Close(Index As Integer)
'One of the connections to the server was closed.

'Close the connection.
sckConnection(Index).Close

'If someone was on that connection, notify open connections.
If sConnectionName(Index) <> "" Then
    'Update the status.
    'UpdateStatus sConnectionName(Index) & " disconnected."
    'Remove their name from the name list.

    'Have the server notify all connected computer that this person has disconnected.
    SendToAll SCK_CODE_DISCONNECTED & sFormatSend(Index), True
    
    If Index = currentPlayer Then
        SendToAll SCK_CODE_CP_DISCONNECTED & sFormatSend(Index), True
        cmdSubmitTopics.Enabled = True
    End If
    
    totalPoints(Index + 1) = 0
    totalVotes(Index + 1) = 0
    RemoveName Index

End If
End Sub
Private Sub sckConnection_ConnectionRequest(Index As Integer, ByVal requestID As Long)
'A connection was requested from the server.

Dim i As Integer
Dim iConnection As Integer

'Make sure this is control 0 in the array.  This is the only one that can accept connections.
If Index = 0 Then

    'Search for available Winsock control.
    For i = 1 To miNumConnections
        If sckConnection(i).State = sckClosed Then
            iConnection = i
            Exit For
        End If
    Next i
    
    'If none was found, create a new one.
    If iConnection = 0 Then
        'Increment number of connections.
        miNumConnections = miNumConnections + 1
        'Load a new Winsock control for this connection.
        Load sckConnection(miNumConnections)
        'Control to be used is this new control.
        iConnection = miNumConnections
    End If
    
    'Set port for this control to 0.  (Randomly assigns an available port.)
    sckConnection(iConnection).LocalPort = 0
    'Have this control accept the connection.
    sckConnection(iConnection).Accept requestID

End If
End Sub
Private Sub sckConnection_DataArrival(Index As Integer, ByVal bytesTotal As Long)
'Data has arrived at the server from an open connection.
Dim sString As String

'Get the data.
sckConnection(Index).GetData sString, vbString

'Process the data.  Pass the index of the connection from which the data came.
ProcessData sString, Index
End Sub




Private Sub tmrPoemHold_Timer()

TopicHoldSeconds = TopicHoldSeconds - 1
If TopicHoldSeconds = 0 Then
    cmdSubmitTopics.Enabled = True
    tmrPoemHold.Enabled = False
End If

End Sub

Private Sub tmrSendData_Timer()
'The is the timer that continuously checks for data to send.

'Remembers whether or not something has been sent.
'Only one piece of data can be sent at a time, otherwise the data runs togeter.
Dim bSent As Boolean

'Index variable to determine which piece of data from the queue will be sent.
Dim iSend As Long

'Remembers where the data will be sent.
Dim iConnection As Integer

'Start the index variable at 1.
iSend = 1

'Loop while nothing has been sent and while the index variable is less than the maximum.
Do While bSent = False And iSend <= mSendTo.Count
    If mSendTo.Item(iSend) = "sckConnect" And sckConnect.State = sckConnected Then
        'Check to see if it is to be sent to the server and make sure the connection is still open.
        
        'Send the data.
        sckConnect.SendData mSendList.Item(iSend) & vbCrLf
    
        'Delete the data from the queue.
        mSendTo.Remove iSend
        mSendList.Remove iSend
        
        'Something has been sent.
        bSent = True
    ElseIf Mid(mSendTo.Item(iSend), 1, 13) = "sckConnection" Then
        'Check to see if it is to be sent to one of the connections to you, the server.
    
        'Parse the string containing the name of the connection to determine which connection to send to.
        
        iConnection = Mid(mSendTo.Item(iSend), 15, Len(mSendTo.Item(iSend)) - 15)
        
        'Ensure that the connection is open.
        If sckConnection(iConnection).State = sckConnected Then
            'Send the data.
            sckConnection(iConnection).SendData mSendList.Item(iSend) & vbCrLf
        
            'Display sent data in tutorial section.
            
            'Delete the data from the queue.
            mSendTo.Remove iSend
            mSendList.Remove iSend
            
            'Something has been sent.
            bSent = True
        End If
    End If
    
    'Increment index variable.
    iSend = iSend + 1
Loop
End Sub


Private Sub txtIP_KeyPress(keyascii As Integer)
If keyascii = vbKeyReturn Then
    'If enter was pressed in the text box that inputs an IP address to connect to, simulate the pressing of the Connect button.
    cmdConnect_Click
    'Make VB think nothing was pressed on the keyboard.  This prevents it from making an annoying beep.
    keyascii = 0
End If
End Sub
Private Sub txtMessage_KeyPress(keyascii As Integer)
If keyascii = vbKeyReturn Then
    'If enter was pressed in the text box that inputs a message to send, simulate the pressing of the Send button.
    cmdSend_Click
    'Clear the text box.
    txtMessage.Text = ""
    'Make VB think nothing was pressed on the keyboard.  This prevents it from making an annoying beep.
    keyascii = 0
End If
End Sub
Private Sub txtName_KeyPress(keyascii As Integer)
If keyascii = vbKeyReturn Then
    'If enter was pressed in the text box that inputs your name, update your name on your screen and on all other computers.
    
    If mbServer Then
        'If you are the server, update your name on your screen.
        ChangeAddName SELF, txtName.Text
        'Refresh name list on all connections.
        SendPeopleList
    Else
        'Send new name to server.
        SendToServer SCK_CODE_CHANGE_NAME & txtName.Text
    End If
        
    'Make VB think nothing was pressed on the keyboard.  This prevents it from making an annoying beep.
    keyascii = 0
End If
End Sub
Public Sub SendPeopleList()
'This is a procedure to refresh each user's connection list.

Dim i As Integer, j As Integer

'Cycle through all connections.
For i = 0 To lstConnections.ListCount - 1
    
    'Do not send list to self.
    If lstConnections.ItemData(i) <> SELF Then
    
        'Send command to clear name list to user.
        SendToPerson SCK_CODE_NEW_NAME_LIST, lstConnections.ItemData(i)
    
        'Send the name for each user to each connection.
        For j = 0 To lstConnections.ListCount - 1
            SendToPerson SCK_CODE_PEOPLE & sFormatSend(lstConnections.ItemData(j)) & lstConnections.List(j), lstConnections.ItemData(i)
        Next j
        
    End If
        
Next i
End Sub
Public Sub ClearStuff()
'This procedure clears stuff out that is used during a chat room.
'It is used to reset stuff after a chat room is closed.

'Clear the data queue.
Set mSendList = Nothing
Set mSendTo = Nothing

'Hide the Kick button.
cmdKick.Visible = False

'Clear the connection list.
lstConnections.Clear

'Clear the dialog.
txtDialog.Text = ""

'Clear the drawing.
'picDraw.Cls
End Sub
Public Function sParam(vsData As String, viNum As Integer) As String
'This function pulls the (viNum)th parameter from datastream vsData, which is being processed in the ProcessData procedure.
'This parameter is exactly PARAM_LEN characters long.

sParam = Mid(vsData, PARAM_LEN * (viNum - 1) + 1, PARAM_LEN)
End Function
Public Function sLongParam(vsData As String, viNum As Integer) As String
'This function pulls the (viNum)th parameter from datastream vsData, which is being processed in the ProcessData procedure.
'This parameter can be any length and is usually at the end of a command.
'This type of parameter usually contains a name and is therefore not a fixed length.

sLongParam = Mid(vsData, PARAM_LEN * (viNum - 1) + 1, Len(vsData))
End Function

Public Function bConnected() As Boolean
'This function returns True if any connections are open.
'This is used to see if you are allowed to change port settings, host a chat room, or connect to a chat room.

Dim i As Integer

For i = 1 To miNumConnections
    If sckConnection(i).State <> sckClosed Then
        bConnected = True
        Exit Function
    End If
Next i

If sckConnect.State <> sckClosed Then
    bConnected = True
End If
End Function
Public Sub AddName(viConnection As Integer, vsName As String)
'This procedure adds a name to the name list.
'viConnection = the connection the user is on
'vsName = the name of the person

Dim i As Integer

'Add the name to the connections list.
lstConnections.AddItem vsName
'Associate that item in the name list with this connection.
For i = 0 To lstConnections.ListCount - 1
    If lstConnections.ItemData(i) = 0 Then
        lstConnections.ItemData(i) = viConnection
        Exit For
    End If
Next i
End Sub
Public Sub ChangeAddName(viConnection As Integer, vsName As String)
'This procedure changes a name in the name list, or adds it if not found.
'viConnection = the connection the user is on
'vsName = the name of the person

Dim i As Integer, j As Integer
Dim bFound As Boolean

'Search for name corresponding to that connection, remove it, and re-add it.
'This ensures that the sorted list box remains sorted.
For i = 0 To lstConnections.ListCount - 1
    If lstConnections.ItemData(i) = viConnection Then
        'Remove the name.
        lstConnections.RemoveItem i
        'Add the name.
        lstConnections.AddItem vsName
        'Find which element in the list was just added and associate correct connection with it.
        For j = 0 To lstConnections.ListCount - 1
            'New element will have ItemData of 0.
            If lstConnections.ItemData(j) = 0 Then
                lstConnections.ItemData(j) = viConnection
                Exit For
            End If
        Next j
        bFound = True
        Exit For
    End If
Next i

If Not bFound Then
    AddName viConnection, vsName
End If
End Sub
Public Sub RemoveName(viConnection As Integer)
'This procedure removes a name from the name list.
'viConnection = the connection the user is on
            
Dim i As Integer

For i = 0 To lstConnections.ListCount - 1
    If lstConnections.ItemData(i) = viConnection Then
        lstConnections.RemoveItem i
        Exit For
    End If
Next i
End Sub
Public Function sConnectionName(viConnection As Integer) As String
'This functions searches the list of connections for the name of a user.
'viConnection = the connection the user is on

Dim i As Integer

For i = 0 To lstConnections.ListCount - 1
    If lstConnections.ItemData(i) = viConnection Then
        sConnectionName = lstConnections.List(i)
        Exit For
    End If
Next i
End Function
Public Sub SendToAll(vsData As String, vbSelf As Boolean)
'Send vsData to all connections.
'vbSelf determine whether or not vsData is sent to yourself as well.

Dim i As Integer

'Cycle through connections and send data to each open connection.
For i = 1 To miNumConnections
    If frmMain.sckConnection(i).State = sckConnected Then
        SendToPerson vsData, i
    End If
Next i

'Send to self if necessary.
If vbSelf Then
    SendToSelf vsData
End If
End Sub
Public Sub SendToPerson(vsData As String, viConnection As Integer)
'Send vsData to viConnection.

mSendList.Add vsData
mSendTo.Add "sckConnection(" & viConnection & ")"
End Sub
Public Sub SendToSelf(vsData As String)
'Send vsData to yourself (the server).

'Just call ProcessData on vsData.
ProcessData vsData & vbCrLf, SELF
End Sub
Public Sub SendToServer(vsData As String)
'Send vsData to server.

mSendList.Add vsData
mSendTo.Add "sckConnect"
End Sub

'Public Sub UpdateStatus(vsStatus As String)
'Add vsStatus to the chat room status.
'txtDialog.Text = txtDialog.Text & vbCrLf & vsStatus

'Put the selection point at the end of the text box so you are seeing the most recent text.
'txtDialog.SelStart = Len(txtDialog.Text)

'If there is a blank carriage return at the beginning, delete it.
'If Mid(txtDialog.Text, 1, Len(vbCrLf)) = vbCrLf Then
'    txtDialog.Text = Mid(txtDialog.Text, Len(vbCrLf) + 1, Len(txtDialog.Text))
'End If
'End Sub

Public Sub UpdateDialog(vsDialog As String)
Dim i As Integer, numNewLines As Integer

'Add vsDialog to the chat room dialog.
txtDialog.Text = txtDialog.Text & vbCrLf & vsDialog

'If there is a blank carriage return at the beginning, delete it.
If Mid(txtDialog.Text, 1, Len(vbCrLf)) = vbCrLf Then
    txtDialog.Text = Mid(txtDialog.Text, Len(vbCrLf) + 1, Len(txtDialog.Text))
End If

'Put the selection point at the end of the text box so you are seeing the most recent text.
txtDialog.SelStart = Len(txtDialog.Text)

totalChatLines = 0
For i = 1 To Len(txtDialog.Text)
    If Mid(txtDialog.Text, i, 1) = Chr(13) Then totalChatLines = totalChatLines + 1
Next i
'totalChatLines = totalChatLines + Int(Len(vsDialog) / 55) + 1

If totalChatLines > 200 Then
    'For i = 1 To totalChatLines - 200
    '    txtDialog.Text
    'Next i
    For i = 1 To totalChatLines - 200
        txtDialog.Text = Mid(txtDialog.Text, InStr(1, txtDialog.Text, Chr(13)) + 2, Len(txtDialog.Text))
    Next i
End If

'Put the selection point at the end of the text box so you are seeing the most recent text.
txtDialog.SelStart = Len(txtDialog.Text)

End Sub

Public Sub OpenConnection()
'Hide/show certain controls because a connection is being opened.

cmdHost.Visible = False
cmdConnect.Visible = False
'cmdPorts.Visible = False
End Sub
Public Sub CloseConnection()
'Hide/show certain controls because a connection is being closed.

cmdHost.Visible = True
cmdConnect.Visible = True
End Sub

