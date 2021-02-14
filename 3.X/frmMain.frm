VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "The PoemGame"
   ClientHeight    =   4275
   ClientLeft      =   2535
   ClientTop       =   2130
   ClientWidth     =   8325
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   8325
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrTopicHold 
      Enabled         =   0   'False
      Left            =   6480
      Top             =   480
   End
   Begin VB.CommandButton cmdStartGame 
      Caption         =   "Start Game"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7080
      TabIndex        =   14
      Top             =   3360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSComctlLib.ImageList imgGameStatus 
      Left            =   4800
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0C42
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0F96
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdAway 
      Caption         =   "Go Away"
      Height          =   375
      Left            =   6000
      TabIndex        =   12
      Top             =   3120
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSComctlLib.StatusBar StatusBar2 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   4020
      Width           =   8325
      _ExtentX        =   14684
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   7303
            Text            =   "Topic 1: None"
            TextSave        =   "Topic 1: None"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   7303
            Text            =   "Topic 2: None"
            TextSave        =   "Topic 2: None"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtConnectTo 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Top             =   720
      Width           =   1815
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton cmdHost 
      Caption         =   "Host"
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   1080
      Width           =   615
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   1080
      Width           =   1095
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   3765
      Width           =   8325
      _ExtentX        =   14684
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   4860
            MinWidth        =   3881
            Text            =   "Disconnected"
            TextSave        =   "Disconnected"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3519
            Picture         =   "frmMain.frx":12EA
            Text            =   "No Game"
            TextSave        =   "No Game"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6174
            MinWidth        =   6174
            Text            =   "Current Player: None"
            TextSave        =   "Current Player: None"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdDisconnect 
      Caption         =   "Disconnect"
      Height          =   375
      Left            =   4920
      TabIndex        =   9
      Top             =   3120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Timer tmrSendData 
      Interval        =   1
      Left            =   6120
      Top             =   480
   End
   Begin MSWinsockLib.Winsock sckConnection 
      Index           =   0
      Left            =   5760
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckConnect 
      Left            =   5400
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtChat 
      Height          =   285
      Left            =   0
      TabIndex        =   5
      Top             =   3180
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.TextBox txtChatBox 
      Height          =   2700
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.ListBox lstUsers 
      Height          =   2595
      ItemData        =   "frmMain.frx":163E
      Left            =   4920
      List            =   "frmMain.frx":1640
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   480
      Width           =   3255
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Come Back"
      Height          =   375
      Left            =   6000
      TabIndex        =   13
      Top             =   3120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdEndGame 
      Caption         =   "End Game"
      Height          =   375
      Left            =   7080
      TabIndex        =   15
      Top             =   3120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Exit"
      Height          =   195
      Left            =   2760
      TabIndex        =   19
      Top             =   75
      Width           =   255
   End
   Begin VB.Line Line7 
      BorderColor     =   &H80000007&
      X1              =   3120
      X2              =   3120
      Y1              =   0
      Y2              =   360
   End
   Begin VB.Line Line6 
      X1              =   2640
      X2              =   2640
      Y1              =   0
      Y2              =   360
   End
   Begin VB.Line Line5 
      X1              =   2040
      X2              =   2040
      Y1              =   0
      Y2              =   360
   End
   Begin VB.Line Line4 
      X1              =   1080
      X2              =   1080
      Y1              =   0
      Y2              =   360
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   360
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   3120
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   3120
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label abtBox 
      AutoSize        =   -1  'True
      Caption         =   "About"
      Height          =   195
      Left            =   2160
      TabIndex        =   18
      Top             =   75
      Width           =   420
   End
   Begin VB.Label editColors 
      AutoSize        =   -1  'True
      Caption         =   "Edit Colors"
      Height          =   195
      Left            =   1200
      TabIndex        =   17
      Top             =   75
      Width           =   750
   End
   Begin VB.Label PoemHist 
      AutoSize        =   -1  'True
      Caption         =   "Poem History"
      Height          =   195
      Left            =   75
      TabIndex        =   16
      Top             =   75
      Width           =   930
   End
   Begin VB.Label Label2 
      Caption         =   "Server IP:"
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Username:"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sendTo(50) As Integer
Dim sendList(50) As String
Dim sendItems As Integer

Dim numConnections As Integer
Dim isServer As Boolean
Dim chatLines As Integer
Const SELF = -1

Dim TopicsMessage As String

'TRUE for game playing, FALSE otherwise
Dim gamePlay As Boolean
Dim currentPlayerName As String
'Connection Index of current player
Dim currentPlayerIndex As Integer
Dim currentPoem As String
Dim Topic1 As String
Dim Topic2 As String
Dim totalPoints(100) As Integer
Dim totalVotes(100) As Integer
Dim finalScore(100) As Integer
Dim PoemHistory(9) As String
Dim PoemAuthor(9) As String
Dim NumberOfPoems As Integer

Dim Ops As String

Const invalidCharacters = " (),.@~/%:|[]"

'MYSTERY MESSAGES   ;)
Const CODE_JOINED = "[Here comes another one]"
Const CODE_YOURNAME = "[You bafoon!  You don't even know your own name?]"
Const CODE_STATUS = "[And this is where we are now]"
Const CODE_DISCONNECTED = "[He died!]"
Const CODE_MESSAGE = "[Don't read this, GreenReaper.]"
Const CODE_PRIVATE_MESSAGE = "[You can't see this!]"
Const CODE_CLEAR_USERS = "[Yours is no good.]"
Const CODE_NEWLIST = "[This one is more up-to-date, dork.]"
Const CODE_NEWGAME = "[Please insert 50 cents]"
Const CODE_ENDGAME = "[Aw shucks, its over]"
Const CODE_NEWPLAYER = "[Next victim, coming down!]"
Const CODE_YOUPLAY = "[Your Turn, Bozo]"
Const CODE_NEWPOEM = "[More worthless poetry...]"
Const CODE_SCORE = "[Yours sucked.  -128 Points.]"
Const CODE_AWAY = "[I'm a wimp, master!  Please exclude me!]"
Const CODE_BACK = "[I'm not real bright.  Add me back to the game.]"
Const CODE_OPPED = "[The server deems you worthy.]"
Const CODE_DEOPPED = "[You are trash in the eyes of the server.]"
Const CODE_KICK = "[Throw da' bum out!]"

'MESSAGES TO SIMPLIFY DEBUGGING  =P
'Const CODE_JOINED = "[JOINED]"
'Const CODE_YOURNAME = "[YOUR NAME]"
'Const CODE_STATUS = "[STATUS]"
'Const CODE_DISCONNECTED = "[DISCONNECTED]"
'Const CODE_MESSAGE = "[MESSAGE]"
'Const CODE_PRIVATE_MESSAGE = "[PRIVATE MESSAGE]"
'Const CODE_CLEAR_USERS = "[CLEAR LIST]"
'Const CODE_NEWLIST = "[NEW LIST]"
'Const CODE_NEWGAME = "[NEW GAME]"
'Const CODE_ENDGAME = "[END GAME]"
'Const CODE_NEWPLAYER = "[NEW PLAYER]"
'Const CODE_YOUPLAY = "[YOU PLAY]"
'Const CODE_NEWPOEM = "[NEW POEM]"
'Const CODE_SCORE = "[SCORE]"
'Const CODE_AWAY = "[GONE AWAY]"
'Const CODE_BACK = "[COME BACK]"
'Const CODE_OPPED = "[OPPED]"
'Const CODE_DEOPPED = "[DEOPPED]"
'Const CODE_KICK = "[KICK]"
Const CODE = "[CODE]"

Public Function LoadINIData()
Dim BGC As Long
Dim BGTC As Long
Dim BBGC As Long
Dim BTC As Long
Dim uNameTemp As String
Dim lastIP As String

Open "poemgame.ini" For Input As #1
    Input #1, uNameTemp
    Input #1, lastIP
    Input #1, BGC
    Input #1, BGTC
    Input #1, BBGC
    Input #1, BTC
Close #1

txtName.Text = uNameTemp
txtConnectTo.Text = lastIP

frmColors.dlgBack.Color = BGC
frmColors.dlgBackText.Color = BGTC
frmColors.dlgBoxBack.Color = BBGC
frmColors.dlgBoxText.Color = BTC

frmColors.UpdateColors

End Function

Public Function GetNumPoems() As Integer
    GetNumPoems = NumberOfPoems
End Function

Public Function GetPoemData(index As Integer) As String

    GetPoemData = PoemAuthor(index) & "|" & PoemHistory(index)

End Function

Public Function KickUser(recipient As String)
    Dim i As Integer
    
    For i = 0 To lstUsers.ListCount - 1
        If UCase(Mid(lstUsers.List(i), 1, InStr(1, lstUsers.List(i), " ") - 1)) = UCase(recipient) Then
            If lstUsers.ItemData(i) <> SELF Then
                If InStr(1, Ops, "," & Trim(Str(lstUsers.ItemData(i))) & ",") Then
                Else
                    sckConnection(lstUsers.ItemData(i)).Close
            
                    SendToAllClients CODE_MESSAGE & Mid(lstUsers.List(i), 1, InStr(1, lstUsers.List(i), " ") - 1) & " was kicked."
                    UpdateDialog Mid(lstUsers.List(i), 1, InStr(1, lstUsers.List(i), " ") - 1) & " was kicked."
                    lstUsers.RemoveItem (i)
                    UpdateClientLists
                End If
            End If
            Exit For
        End If
    Next i
    
End Function

Public Function UpdateConnected()
    cmdStartGame.Visible = True
    txtChatBox.Visible = True
    txtChatBox.Text = ""
    txtChat.Visible = True
    txtName.Visible = False
    txtConnectTo.Visible = False
    cmdConnect.Visible = False
    cmdHost.Visible = False
    Label1.Visible = False
    Label2.Visible = False
    cmdDisconnect.Visible = True
    cmdAway.Visible = True
End Function
Public Function UpdateDisconnected()
    cmdStartGame.Visible = False
    txtChatBox.Visible = False
    txtChatBox.Text = ""
    txtChat.Visible = False
    txtName.Visible = True
    txtConnectTo.Visible = True
    cmdConnect.Visible = True
    cmdHost.Visible = True
    Label1.Visible = True
    Label2.Visible = True
    cmdDisconnect.Visible = False
    cmdAway.Visible = False
    lstUsers.Clear
    StatusBar1.Panels(2).Text = "No Game"
    StatusBar1.Panels(2).Picture = imgGameStatus.ListImages(1).Picture
    StatusBar1.Panels(3).Text = "Current Player: None"
End Function

Public Function AddUserToList(connectionNumber As Integer, name As String)
    
    lstUsers.AddItem name
    
    Dim i As Integer
    For i = 0 To lstUsers.ListCount - 1
        If lstUsers.ItemData(i) = 0 Then
            lstUsers.ItemData(i) = connectionNumber
            Exit For
        End If
    Next i

End Function
Public Function ProcessTopics()
    Dim i As Integer
    
    If isServer Then
        SendToAllClients CODE_NEWPLAYER & frmTopics.lstUsernames.List(frmTopics.lstUsernames.ListIndex) & "|" & frmTopics.txtTopic1.Text & "|" & frmTopics.txtTopic2.Text
        SendToOneClient CODE_YOUPLAY, frmTopics.lstUsernames.ItemData(frmTopics.lstUsernames.ListIndex)
        
        currentPlayerName = frmTopics.lstUsernames.List(frmTopics.lstUsernames.ListIndex)
        Topic1 = frmTopics.txtTopic1.Text
        Topic2 = frmTopics.txtTopic2.Text
        
        For i = 0 To lstUsers.ListCount - 1
            If Mid(lstUsers.List(i), 1, InStr(1, lstUsers.List(i), " ") - 1) = currentPlayerName Then
                currentPlayerIndex = lstUsers.ItemData(i)
                Exit For
            End If
        Next i
        
        UpdateDialog "New player is " & currentPlayerName & " with topics " & Topic1 & " and " & Topic2
        
    Else
        SendToServer CODE_NEWPLAYER & frmTopics.lstUsernames.ItemData(frmTopics.lstUsernames.ListIndex) & "|" & frmTopics.lstUsernames.List(frmTopics.lstUsernames.ListIndex) & "|" & frmTopics.txtTopic1.Text & "|" & frmTopics.txtTopic2.Text
    End If
    
    StatusBar1.Panels(3).Text = "Current Player: " & frmTopics.lstUsernames.List(frmTopics.lstUsernames.ListIndex)
    StatusBar2.Panels(1).Text = "Topic 1: " & frmTopics.txtTopic1.Text
    StatusBar2.Panels(2).Text = "Topic 2: " & frmTopics.txtTopic2.Text

End Function
Public Function ProcessPoem()
Dim i As Integer

    If isServer Then
        SendToAllClients CODE_NEWPOEM & txtName.Text & "|" & frmCompose.txtPoemCompose.Text
    Else
        SendToServer CODE_NEWPOEM & txtName.Text & "|" & frmCompose.txtPoemCompose.Text
    End If
    
    For i = NumberOfPoems - 1 To 0 Step -1
        PoemHistory(i + 1) = PoemHistory(i)
        PoemAuthor(i + 1) = PoemAuthor(i)
    Next i
                
    PoemAuthor(0) = txtName.Text
    PoemHistory(0) = frmCompose.txtPoemCompose.Text
                                                                
    If NumberOfPoems <> 9 Then
        NumberOfPoems = NumberOfPoems + 1
    End If
    
    currentPoem = PoemHistory(0)
    
    frmPoemBox.Show
    frmPoemBox.cmdScore.Enabled = False

End Function
Public Function getPartialIP(index As Integer) As String
Dim fullIP As String
Dim IPPart As String
Dim junk As String
Dim i
    
    If index = -1 Then
        For i = 2 To 4
            If InStr(1, Right(sckConnect.LocalIP, i), ".") Then
                getPartialIP = Mid(sckConnect.LocalIP, 1, Len(sckConnect.LocalIP) - i) & ".*"
                Exit For
            End If
        Next i
    Else
        For i = 2 To 4
            If InStr(1, Right(sckConnection(index).RemoteHostIP, i), ".") Then
                getPartialIP = Mid(sckConnection(index).RemoteHostIP, 1, Len(sckConnect.LocalIP) - i) & ".*"
                Exit For
            End If
        Next i
    End If
        
 End Function
 Public Function CalculateScore(index As Integer) As String
    If totalVotes(index + 1) <> 0 Then
        CalculateScore = Mid(Trim(Str(totalPoints(index + 1) / totalVotes(index + 1))), 1, 4)
    Else
        CalculateScore = "0"
    End If
 End Function
Public Function UpdateClientLists()
    
    Dim i As Integer, j As Integer
    Dim fullUserList As String
    
    fullUserList = Trim(Str(lstUsers.ListCount)) & ":"
    For i = 0 To lstUsers.ListCount - 1
        fullUserList = fullUserList & Trim(Str(lstUsers.ItemData(i))) & "@" & lstUsers.List(i) & "~"
    Next i
    
    For i = 0 To lstUsers.ListCount - 1
        If lstUsers.ItemData(i) <> SELF Then
            
            SendToOneClient CODE_CLEAR_USERS, lstUsers.ItemData(i)
            SendToOneClient CODE_NEWLIST & fullUserList, lstUsers.ItemData(i)
            
        End If
    Next i
    
End Function
Public Function ChangeList(name As String, index As Integer)
    Dim i As Integer, j As Integer
    Dim userInList As Boolean
    userInList = False
    
    For i = 0 To lstUsers.ListCount - 1
        If lstUsers.ItemData(i) = index Then
            lstUsers.RemoveItem i
            lstUsers.AddItem name & name
            
            For j = 0 To lstUsers.ListCount - 1
                If lstUsers.ItemData(j) = 0 Then
                    lstUsers.ItemData(j) = index
                    Exit For
                End If
            Next j
            
            userInList = True
        End If
    Next i

If Not userInList Then
    AddUserToList index, name
End If
End Function
Public Function Process(incoming As String, Optional index As Integer)

    Dim NewData As String
    Dim COMMAND As String
    Dim PARAMS As String
    Dim i As Integer
    
    Dim connectionNumber As Integer
    Dim ListData As String
    Dim tempJunk As String

    Do While InStr(1, incoming, vbCrLf)
    NewData = Mid(incoming, 1, InStr(1, incoming, vbCrLf) - 1)
    '/// FOR DEBUGGING, UNCOMMENT THE FOLLOWING LINE ///
    'UpdateDialog "RECEIVING: " & newData
    '/// FOR DEBUGGING, UNCOMMENT THE ABOVE LINE///

    COMMAND = Mid(NewData, 1, InStr(1, NewData, "]"))
    PARAMS = Mid(NewData, InStr(1, NewData, "]") + 1, Len(NewData))
        
    Select Case COMMAND
    
        Case CODE_DISCONNECTED
            UpdateDialog PARAMS & " has disconnected."
    
        Case CODE_MESSAGE
            If isServer Then
                SendToAllClients NewData
            End If
            
            UpdateDialog PARAMS
        
        Case CODE_NEWGAME
            StatusBar1.Panels(2).Text = "Game Playing"
            StatusBar1.Panels(2).Picture = imgGameStatus.ListImages(2).Picture
            UpdateDialog vbNewLine & "Game has started." & vbNewLine
        
        Case CODE_ENDGAME
            StatusBar1.Panels(2).Text = "No Game"
            StatusBar1.Panels(2).Picture = imgGameStatus.ListImages(1).Picture
            StatusBar1.Panels(3).Text = "Current Player: None"
            StatusBar2.Panels(1).Text = "Topic 1: None"
            StatusBar2.Panels(2).Text = "Topic 2: None"
            UpdateDialog vbNewLine & "Game has ended."
            Unload frmCompose
            Unload frmTopics
            frmPoemBox.cmdScore.Enabled = False
        
        Case CODE_PRIVATE_MESSAGE
            Dim senderName As String
            Dim userConnIndex As Integer
            Dim privateMessage As String
            
            privateMessage = Mid(PARAMS, InStr(1, PARAMS, "%") + 1, Len(PARAMS))
            For i = 0 To lstUsers.ListCount - 1
                If UCase(Mid(lstUsers.List(i), 1, InStr(1, lstUsers.List(i), " ") - 1)) = UCase(Mid(PARAMS, 1, InStr(1, PARAMS, "%") - 1)) Then
                    userConnIndex = lstUsers.ItemData(i)
                    Exit For
                End If
            Next i
            
            For i = 0 To lstUsers.ListCount - 1
                If lstUsers.ItemData(i) = index Then
                    senderName = Mid(lstUsers.List(i), 1, InStr(1, lstUsers.List(i), " ") - 1)
                    Exit For
                End If
            Next i
            
            If userConnIndex = SELF Then
                UpdateDialog "[From " & senderName & "] " & privateMessage
            Else
                SendToOneClient CODE_MESSAGE & "[From " & senderName & "] " & privateMessage, userConnIndex
            End If
            
        Case CODE_YOURNAME
            UpdateDialog "DUPLICATE NAME -- Your name is now " & PARAMS
            txtName.Text = PARAMS

        Case CODE_JOINED
            
            If isServer Then
                Dim duplicateName As Boolean
                
                For i = 0 To lstUsers.ListCount - 1
                    If UCase(PARAMS) = UCase(Mid(lstUsers.List(i), 1, InStr(1, lstUsers.List(i), " ") - 1)) Then
                        duplicateName = True
                        Exit For
                    End If
                Next i
                
                If Not duplicateName Then
                    SendToAllClients NewData
                    AddUserToList index, PARAMS & " (" & getPartialIP(index) & ") [" & CalculateScore(index) & " Points]"
                    UpdateDialog PARAMS & " has joined."
                Else
                    SendToOneClient CODE_YOURNAME & PARAMS & "2", index
                    SendToAllClients NewData & "2"
                    AddUserToList index, PARAMS & "2 (" & getPartialIP(index) & ") [" & CalculateScore(index) & " Points]"
                    UpdateDialog PARAMS & "2 has joined."
                End If
                UpdateClientLists
                
                If gamePlay Then
                    SendToOneClient CODE_STATUS & "1," & currentPlayerName & "," & Topic1 & "," & Topic2, index
                Else
                    SendToOneClient CODE_STATUS & "0,", index
                End If
                
            End If
            
        Case CODE_STATUS
            
            If Mid(PARAMS, 1, InStr(1, PARAMS, ",") - 1) = "1" Then
                StatusBar1.Panels(2).Text = "Game Playing"
                StatusBar1.Panels(2).Picture = imgGameStatus.ListImages(2).Picture
                tempJunk = Mid(PARAMS, InStr(1, PARAMS, ",") + 1, Len(PARAMS))
                StatusBar1.Panels(3).Text = "Current Player: " & Mid(tempJunk, 1, InStr(1, tempJunk, ",") - 1)
                tempJunk = Mid(tempJunk, InStr(1, tempJunk, ",") + 1, Len(tempJunk))
                StatusBar2.Panels(1).Text = "Topic 1: " & Mid(tempJunk, 1, InStr(1, tempJunk, ",") - 1)
                StatusBar2.Panels(2).Text = "Topic 2: " & Mid(tempJunk, InStr(1, tempJunk, ",") + 1, Len(tempJunk))
            End If
            
        Case CODE_CLEAR_USERS
            lstUsers.Clear
            
        Case CODE_NEWLIST
            
            Dim tempUsers As Integer
            
            tempUsers = Val(Mid(PARAMS, 1, InStr(1, PARAMS, ":") - 1))
            tempJunk = Mid(PARAMS, InStr(1, PARAMS, ":") + 1, Len(PARAMS))
            
            For i = 1 To tempUsers
                connectionNumber = Val(Mid(tempJunk, 1, InStr(1, tempJunk, "@") - 1))
                tempJunk = Mid(tempJunk, InStr(1, tempJunk, "@") + 1, Len(tempJunk))
                
                AddUserToList connectionNumber, Mid(tempJunk, 1, InStr(1, tempJunk, "]"))
                
                tempJunk = Mid(tempJunk, InStr(1, tempJunk, "~") + 1, Len(tempJunk))
            Next i
            
        Case CODE_NEWPLAYER
        
            tempJunk = PARAMS
            If isServer Then
                connectionNumber = Val(Mid(PARAMS, 1, InStr(1, PARAMS, "|") - 1))
                tempJunk = Mid(PARAMS, InStr(1, PARAMS, "|") + 1, Len(PARAMS))
                
                SendToAllClients CODE_NEWPLAYER & tempJunk
                If connectionNumber <> SELF Then
                    SendToOneClient CODE_YOUPLAY, connectionNumber
                Else
                    frmCompose.Show
                End If
                currentPlayerIndex = connectionNumber
            End If
                
            currentPlayerName = Mid(tempJunk, 1, InStr(1, tempJunk, "|") - 1)
                tempJunk = Mid(tempJunk, InStr(1, tempJunk, "|") + 1, Len(tempJunk))
            Topic1 = Mid(tempJunk, 1, InStr(1, tempJunk, "|") - 1)
            Topic2 = Mid(tempJunk, InStr(1, tempJunk, "|") + 1, Len(tempJunk))
                
            UpdateDialog "New player is " & currentPlayerName & " with topics " & Topic1 & " and " & Topic2
            StatusBar1.Panels(3).Text = "Current Player: " & currentPlayerName
            StatusBar2.Panels(1).Text = "Topic 1: " & Topic1
            StatusBar2.Panels(2).Text = "Topic 2: " & Topic2
            frmPoemBox.cmdScore.Enabled = False
            
        Case CODE_YOUPLAY
            frmCompose.Show
        
        Case CODE_NEWPOEM
            If isServer Then
                For i = 0 To lstUsers.ListCount - 1
                    If lstUsers.ItemData(i) <> index And lstUsers.ItemData(i) <> SELF Then
                        SendToOneClient CODE_NEWPOEM & PARAMS, lstUsers.ItemData(i)
                    End If
                Next i
            End If

                For i = NumberOfPoems - 1 To 0 Step -1
                    PoemHistory(i + 1) = PoemHistory(i)
                    PoemAuthor(i + 1) = PoemAuthor(i)
                Next i
                
                PoemAuthor(0) = Mid(PARAMS, 1, InStr(1, PARAMS, "|") - 1)
                PoemHistory(0) = Mid(PARAMS, InStr(1, PARAMS, "|") + 1, Len(PARAMS))
                
                If NumberOfPoems <> 9 Then
                    NumberOfPoems = NumberOfPoems + 1
                End If
                                                                
                currentPoem = PoemHistory(0)
                frmPoemBox.Show
                frmPoemBox.RefreshForm
                frmPoemBox.cmdScore.Enabled = True
        
        Case CODE_SCORE
        
            totalVotes(currentPlayerIndex + 1) = totalVotes(currentPlayerIndex + 1) + 1
            totalPoints(currentPlayerIndex + 1) = totalPoints(currentPlayerIndex + 1) + Val(PARAMS)
            
            For i = 0 To lstUsers.ListCount - 1
                If currentPlayerIndex = lstUsers.ItemData(i) Then
                    lstUsers.List(i) = currentPlayerName & " (" & getPartialIP(currentPlayerIndex) & ") [" & CalculateScore(currentPlayerIndex) & " Points]"
                    UpdateClientLists
                End If
            Next i
            
        Case CODE_AWAY
            SendToAllClients CODE_MESSAGE & PARAMS & " has gone AWAY."
            UpdateDialog PARAMS & " has gone AWAY."
            For i = 0 To lstUsers.ListCount - 1
                If lstUsers.ItemData(i) = index Then
                    lstUsers.List(i) = "-" & lstUsers.List(i)
                    Exit For
                End If
            Next i
            UpdateClientLists
        
        Case CODE_BACK
            SendToAllClients CODE_MESSAGE & PARAMS & " has come BACK."
            UpdateDialog PARAMS & " has come BACK."
            For i = 0 To lstUsers.ListCount - 1
                If lstUsers.ItemData(i) = index Then
                    lstUsers.List(i) = Mid(lstUsers.List(i), 2, Len(lstUsers.List(i)) - 1)
                    Exit For
                End If
            Next i
            UpdateClientLists
            
        Case CODE_KICK
        
            If InStr(1, Ops, "," & Trim(Str(index)) & ",") Then
                KickUser PARAMS
            End If

    End Select
    
    If InStr(1, incoming, vbCrLf) <> Len(incoming) Then
        incoming = Mid(incoming, InStr(1, incoming, vbCrLf) + 1, Len(incoming))
    Else
        incoming = ""
    End If
    
    Loop
    

End Function
Public Function ProcessScore(SCORE As Integer)
    Dim i As Integer
    
    If isServer Then
        totalVotes(currentPlayerIndex + 1) = totalVotes(currentPlayerIndex + 1) + 1
        totalPoints(currentPlayerIndex + 1) = totalPoints(currentPlayerIndex + 1) + SCORE
    
        For i = 0 To lstUsers.ListCount - 1
            If currentPlayerIndex = lstUsers.ItemData(i) Then
                lstUsers.List(i) = currentPlayerName & " (" & getPartialIP(currentPlayerIndex) & ") [" & CalculateScore(currentPlayerIndex) & " Points]"
                UpdateClientLists
            End If
        Next i
    Else
        SendToServer CODE_SCORE & SCORE
    End If

End Function
Public Function SendToServer(sendData As String)

    sendList(sendItems) = sendData
    sendTo(sendItems) = -1
    sendItems = sendItems + 1

End Function
Public Function SendToAllClients(sendData As String)

    Dim i As Integer
    
    'Send data to all open connections
    For i = 0 To lstUsers.ListCount - 1
        If lstUsers.ItemData(i) <> SELF Then
            If sckConnection(lstUsers.ItemData(i)).State = sckConnected Then SendToOneClient sendData, lstUsers.ItemData(i)
        End If
    Next i

End Function
Public Function SendToOneClient(sendData As String, toWho As Integer)

    'Send data to user at index toWho
    'sendList.Add sendData
    'sendTo.Add "sckConnection " & toWho
    sendList(sendItems) = sendData
    sendTo(sendItems) = toWho
    sendItems = sendItems + 1

End Function
Public Function UpdateDialog(NewData As String)

    'Whether or not to insert a NewLine
    Dim newline As String
    
    'If the chatbox is empty, do not add a NewLine
    If Len(txtChatBox.Text) = 0 Then
        newline = ""
    Else
    'Chatbox is not empty, add a NewLine
        newline = vbNewLine
    End If
    
    txtChatBox.Text = txtChatBox.Text & newline & NewData
    chatLines = chatLines + 1
    
    'Limit chatbox to 200 lines
    If chatLines > 200 Then
        txtChatBox.Text = Mid(txtChatBox.Text, InStr(1, txtChatBox.Text, Chr(13)) + 2, Len(txtChatBox.Text))
        chatLines = 200
    End If
    
    'Go to the bottom of the chatbox
    txtChatBox.SelStart = Len(txtChatBox.Text)
    
End Function

Private Sub abtBox_Click()
    MsgBox "                PoemGame 3.0" & vbNewLine & "                  By JacketFan" & vbNewLine & "               © Copyright 2001" & vbNewLine & vbNewLine & "http://www.sreklaw.com/poemgame/", vbOKOnly, "About..."
End Sub

Private Sub cmdAway_Click()
    
    If isServer Then
    Else
        SendToServer CODE_AWAY & txtName.Text
    End If
    cmdAway.Visible = False
    cmdBack.Visible = True
    
End Sub

Private Sub cmdBack_Click()
    
    If isServer Then
    Else
        SendToServer CODE_BACK & txtName.Text
    End If
    cmdAway.Visible = True
    cmdBack.Visible = Visible


End Sub

Private Sub cmdConnect_Click()
Dim invalidCharacter As Boolean
Dim i As Integer

For i = 1 To Len(txtName.Text)
    If InStr(1, invalidCharacters, Mid(txtName.Text, i, 1)) Then
        invalidCharacter = True
        Exit For
    End If
Next i

If txtName.Text <> "" And txtConnectTo.Text <> "" And Not invalidCharacter Then
    
    Open "poemgame.ini" For Output As #2
        Write #2, txtName.Text
        Write #2, txtConnectTo.Text
        Write #2, frmMain.BackColor
        Write #2, frmMain.ForeColor
        Write #2, txtChatBox.BackColor
        Write #2, txtChatBox.ForeColor
    Close #2
    
    On Error GoTo Err_cmdConnect_Click
    
    isServer = False
    sckConnect.Close
    sckConnect.RemotePort = 2112
    sckConnect.Connect txtConnectTo.Text
    
    SendToServer CODE_JOINED & txtName.Text
    UpdateDialog "Connecting..."
    
ElseIf txtName.Text = "" Then
    txtName.SetFocus
ElseIf txtConnectTo.Text = "" Then
    txtConnectTo.SetFocus
End If

Exit Sub


Err_cmdConnect_Click:
    MsgBox "Unable to connect.", vbExclamation, App.Title
    sckConnect.Close
    UpdateDialog "Disconnected."
End Sub

Private Sub cmdDisconnect_Click()

    Dim i As Integer
    If isServer Then
        sckConnect.Close
        For i = 0 To numConnections
            sckConnection(i).Close
        Next i
        StatusBar1.Panels(1).Text = "Disconnected."
    Else
        sckConnect.Close
        StatusBar1.Panels(1).Text = "Disconnected."
    End If
    UpdateDisconnected
End Sub

Private Sub cmdEndGame_Click()
    Dim i As Integer
    
    If isServer Then
        gamePlay = False
        
        StatusBar1.Panels(2).Text = "No Game"
        StatusBar1.Panels(2).Picture = imgGameStatus.ListImages(1).Picture
        StatusBar1.Panels(3).Text = "Current Player: None"
        StatusBar2.Panels(1).Text = "Topic 1: None"
        StatusBar2.Panels(2).Text = "Topic 2: None"
        
        Unload frmCompose
        Unload frmTopics
        frmPoemBox.cmdScore.Enabled = False
        
        Dim winningScore As Long
        Dim winnerIndex As Integer
        
        For i = 0 To lstUsers.ListCount - 1
            If Val(CalculateScore(lstUsers.ItemData(i))) > winningScore Then
                winningScore = Val(CalculateScore(lstUsers.ItemData(i)))
                winnerIndex = i - 1
            End If
        Next i
        
        For i = 0 To lstUsers.ListCount - 1
            If lstUsers.ItemData(i) = winnerIndex Then
                winnerIndex = i
                Exit For
            End If
        Next i
                
        For i = 0 To 100
            totalVotes(i) = 0
            totalPoints(i) = 0
        Next i
        
        For i = 0 To lstUsers.ListCount - 1
            lstUsers.List(i) = Mid(lstUsers.List(i), 1, InStr(1, lstUsers.List(i), " ") - 1) & " (" & getPartialIP(lstUsers.ItemData(i)) & ") [" & CalculateScore(lstUsers.ItemData(i)) & " Points]"
        Next i
        UpdateClientLists
        
        SendToAllClients CODE_ENDGAME
        SendToAllClients CODE_MESSAGE & Mid(lstUsers.List(winnerIndex), 1, InStr(1, lstUsers.List(winnerIndex), " ") - 1) & " has won with a score of " & Str(winningScore) & "!"
        UpdateDialog vbNewLine & "Game has ended."
        UpdateDialog Mid(lstUsers.List(winnerIndex), 1, InStr(1, lstUsers.List(winnerIndex), " ") - 1) & " has won with a score of " & Str(winningScore) & "!"

        
        cmdEndGame.Visible = False
        cmdStartGame.Visible = True
    End If
End Sub

Private Sub cmdHost_Click()
Dim invalidCharacter As Boolean
Dim i As Integer

For i = 1 To Len(txtName.Text)
    If InStr(1, invalidCharacters, Mid(txtName.Text, i, 1)) Then
        invalidCharacter = True
        Exit For
    End If
Next i

If txtName.Text <> "" And Not invalidCharacter Then
    
    Open "poemgame.ini" For Output As #2
        Write #2, txtName.Text
        Write #2, ""
        Write #2, frmMain.BackColor
        Write #2, frmMain.ForeColor
        Write #2, txtChatBox.BackColor
        Write #2, txtChatBox.ForeColor
    Close #2
    
    isServer = True
    sckConnect.Close

    sckConnection(0).Close
    sckConnection(0).LocalPort = 2112
    sckConnection(0).Listen
    
    UpdateConnected
    cmdStartGame.Enabled = True
    cmdEndGame.Enabled = True
    
    StatusBar1.Panels(1).Text = "Hosting at " & sckConnect.LocalIP
    AddUserToList -1, txtName.Text & " (" & getPartialIP(-1) & ") [0 Points]"
Else
    txtName.SetFocus
End If
End Sub

Private Sub cmdStartGame_Click()
    If isServer Then
        gamePlay = True
        currentPlayerName = ""
        currentPlayerIndex = 0
        Topic1 = ""
        Topic2 = ""
        UpdateDialog vbNewLine & "Game has started." & vbNewLine
        
        SendToAllClients CODE_NEWGAME
        
        StatusBar1.Panels(2).Text = "Game Playing"
        StatusBar1.Panels(2).Picture = imgGameStatus.ListImages(2).Picture
        
        cmdStartGame.Visible = False
        cmdEndGame.Visible = True
        frmTopics.Show
    End If
End Sub

Private Sub editColors_Click()
    frmColors.Show
End Sub



Private Sub Form_Load()
    isServer = False
    Ops = ","
   
    LoadINIData
'    Open "poemgame.ini" For Append As #1 'Output As #1
'        Write #1, frmMain.FontName
'        Write #1, frmMain.FontBold
'        Write #1, frmMain.FontItalic
'        Write #1, frmMain.FontSize
'        Write #1, frmMain.FontStrikethru
'        Write #1, frmMain.FontUnderline
'    Close #1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmTopics
    Unload frmCompose
    Unload frmPoemBox
    End
End Sub





Private Sub Label3_Click()
End
End Sub

Private Sub poemHist_Click()
    frmPoemBox.Show
    frmPoemBox.cmdScore.Enabled = False
End Sub

Private Sub sckConnect_Close()
    StatusBar1.Panels(1).Text = "Disconnected."
    UpdateDisconnected
End Sub

Private Sub sckConnect_Connect()
    txtChatBox.Text = txtChatBox.Text & "Connected."
    StatusBar1.Panels(1).Text = "Connected to " & sckConnect.RemoteHostIP 'txtConnectTo.Text
    UpdateConnected
End Sub

Private Sub sckConnect_DataArrival(ByVal bytesTotal As Long)
    Dim incomingData As String
    
    sckConnect.GetData incomingData, vbString
    
    Process incomingData
    'UpdateDialog incomingData
End Sub

Private Sub sckConnection_Close(index As Integer)
    Dim goneUser As String
    Dim i As Integer
    Dim wasCurrentPlayer As Boolean
    
    If index <> 0 Then
    For i = 0 To lstUsers.ListCount - 1
        If lstUsers.ItemData(i) = index Then
            If lstUsers.ItemData(i) = currentPlayerIndex Then
                wasCurrentPlayer = True
            End If
            goneUser = Mid(lstUsers.List(i), 1, InStr(1, lstUsers.List(i), " ") - 1)
            lstUsers.RemoveItem i
            Exit For
        End If
    Next i
    
    If goneUser <> "" Then
        UpdateDialog goneUser & " has disconnected."
        SendToAllClients CODE_DISCONNECTED & goneUser
        UpdateClientLists
    End If
    
    If wasCurrentPlayer Then
        UpdateDialog "Current player disconnected.  Server will now select new topics and the next player."
        SendToAllClients CODE_MESSAGE & "Current player disconnected.  Server will now select new topics and the next player."
        frmTopics.Show
    End If
    
    End If
    
End Sub

Private Sub sckConnection_ConnectionRequest(index As Integer, ByVal requestID As Long)

    Dim i As Integer
    Dim newConnection As Integer
    
    If index = 0 Then
        For i = 1 To numConnections
            If sckConnection(i).State = sckClosed Then
                newConnection = i
                Exit For
            End If
        Next i
    
        If newConnection = 0 Then
            numConnections = numConnections + 1
            Load sckConnection(numConnections)
            newConnection = numConnections
        End If
    
        sckConnection(newConnection).LocalPort = 0
        sckConnection(newConnection).Accept requestID
    End If

End Sub

Private Sub sckConnection_DataArrival(index As Integer, ByVal bytesTotal As Long)
    Dim incomingData As String
    
    sckConnection(index).GetData incomingData, vbString
    Process incomingData, index
'    UpdateDialog incomingData
End Sub

Private Sub tmrSendData_Timer()

    'Used in the FOR loops
    Dim i As Integer

    'Only attempt to send data if data exists
    If sendItems > 0 Then
    
        '///FOR DEBUGGING, UNCOMMENT NEXT LINE///
        'UpdateDialog "SENDING: " & sendList(0)
        '///FOR DEBUGGING, UNCOMMENT LAST LINE///
    
        'Data is for server and connection is active
        If sendTo(0) = -1 And sckConnect.State = sckConnected Then
            'Send data
            sckConnect.sendData sendList(0) & vbCrLf
        
            'Remove the data from the list and reflect
            'the change in the number of items
            For i = 0 To sendItems - 1
                sendTo(i) = sendTo(i + 1)
                sendList(i) = sendList(i + 1)
            Next i
            sendItems = sendItems - 1
            
        'Data is for client and connection is active
        Else
            If sendTo(0) <> -1 Then
            If sckConnection(sendTo(0)).State = sckConnected Then
                'Send data
                sckConnection(sendTo(0)).sendData sendList(0) & vbCrLf
        
                'Remove the data from the list and reflect
                'the change in the number of items
                For i = 0 To sendItems - 1
                    sendTo(i) = sendTo(i + 1)
                    sendList(i) = sendList(i + 1)
                Next i
                sendItems = sendItems - 1
            End If
            End If
        End If
    End If
    
End Sub

Private Sub tmrTopicHold_Timer()
    frmTopics.Show
    tmrTopicHold.Enabled = False
End Sub

Private Sub txtChat_KeyPress(KeyAscii As Integer)
Dim i As Integer
    If KeyAscii = vbKeyReturn Then
    
    If txtChat.Text <> "" Then
        
        Dim recipient As String
        Dim message As String
        Dim recipientIndex As Integer
        Dim userExists As Boolean
    
        'Message is from server to clients
        If isServer Then
            'Emote using /me or .me
            If Mid(txtChat.Text, 1, 3) = "/me" Or Mid(txtChat.Text, 1, 3) = ".me" Then
                UpdateDialog "* " & txtName.Text & " " & Mid(txtChat.Text, 5, Len(txtChat.Text))
                SendToAllClients CODE_MESSAGE & "* " & txtName.Text & " " & Mid(txtChat.Text, 5, Len(txtChat.Text))
            'Emote using ;
            ElseIf Mid(txtChat.Text, 1, 1) = ";" Then
                UpdateDialog "* " & txtName.Text & " " & Mid(txtChat.Text, 2, Len(txtChat.Text))
                SendToAllClients CODE_MESSAGE & "* " & txtName.Text & " " & Mid(txtChat.Text, 3, Len(txtChat.Text))
            'Private message using /msg or .msg
            ElseIf Mid(txtChat.Text, 1, 4) = "/msg" Or Mid(txtChat.Text, 1, 4) = ".msg" Then
                
                If InStr(7, txtChat.Text, " ") And InStr(7, txtChat.Text, " ") <> Len(txtChat.Text) Then
                    
                    recipient = Mid(txtChat.Text, 6, InStr(7, txtChat.Text, " ") - 6)
                    message = Mid(txtChat.Text, InStr(7, txtChat.Text, " ") + 1, Len(txtChat.Text))
                    
                    For i = 0 To lstUsers.ListCount - 1
                        If UCase(recipient) = UCase(Mid(lstUsers.List(i), 1, InStr(1, lstUsers.List(i), " ") - 1)) Then
                            SendToOneClient CODE_MESSAGE & "[From " & txtName.Text & "] " & message, lstUsers.ItemData(i)
                            userExists = True
                            Exit For
                        End If
                    Next i
                    If userExists Then UpdateDialog "[To " & recipient & "] " & message
                    
                End If
            'Private message using >
            ElseIf Mid(txtChat.Text, 1, 1) = ">" Then
            
                If InStr(4, txtChat.Text, " ") And InStr(4, txtChat.Text, " ") <> Len(txtChat.Text) Then
                
                    If Mid(txtChat.Text, 2, 1) = " " Then
                        recipient = Mid(txtChat.Text, 3, InStr(4, txtChat.Text, " ") - 3)
                        message = Mid(txtChat.Text, InStr(4, txtChat.Text, " ") + 1, Len(txtChat.Text))
                    Else
                        recipient = Mid(txtChat.Text, 2, InStr(3, txtChat.Text, " ") - 2)
                        message = Mid(txtChat.Text, InStr(3, txtChat.Text, " ") + 1, Len(txtChat.Text))
                    End If
                    
                    For i = 0 To lstUsers.ListCount - 1
                        If UCase(recipient) = UCase(Mid(lstUsers.List(i), 1, InStr(1, lstUsers.List(i), " ") - 1)) Then
                            SendToOneClient CODE_MESSAGE & "[From " & txtName.Text & "] " & message, lstUsers.ItemData(i)
                            userExists = True
                            Exit For
                        End If
                    Next i
                    If userExists Then UpdateDialog "[To " & recipient & "] " & message
                    
                End If
            'Server typed /quit or .quit
            ElseIf Mid(txtChat.Text, 1, 5) = "/quit" Or Mid(txtChat.Text, 1, 5) = ".quit" Then
                cmdDisconnect_Click
            'Server wants to KICK a user
            ElseIf Mid(txtChat.Text, 1, 5) = "/kick" Or Mid(txtChat.Text, 1, 5) = ".kick" Then
                
                If InStr(7, txtChat.Text, " ") Then
                    recipient = Mid(txtChat.Text, 7, InStr(7, txtChat.Text, " ") - 6)
                Else
                    recipient = Mid(txtChat.Text, 7, Len(txtChat.Text))
                End If
                
                KickUser recipient
            'Clear the chatbox
            ElseIf Mid(txtChat.Text, 1, 6) = "/clear" Or Mid(txtChat.Text, 1, 6) = ".clear" Then
                txtChatBox.Text = ""
                chatLines = 0
            'Create a non-server Operator
            ElseIf Mid(txtChat.Text, 1, 3) = "/op" Or Mid(txtChat.Text, 1, 3) = ".op" Then
                For i = 0 To lstUsers.ListCount - 1
                    If UCase(Mid(txtChat.Text, 5, Len(txtChat.Text))) = UCase(Mid(lstUsers.List(i), 1, InStr(1, lstUsers.List(i), " ") - 1)) Then
                        recipientIndex = lstUsers.ItemData(i)
                        Exit For
                    End If
                Next i
                
                If recipientIndex > 0 Then
                    If InStr(1, Ops, "," & Trim(Str(recipientIndex)) & ",") Then
                    Else
                        SendToOneClient CODE_OPPED, recipientIndex
                        SendToOneClient CODE_MESSAGE & "You have been opped.", recipientIndex
                        Ops = Ops & Trim(Str(recipientIndex)) & ","
                    End If
                End If
                
            'Remove a non-server Operator
            ElseIf Mid(txtChat.Text, 1, 5) = "/deop" Or Mid(txtChat.Text, 1, 5) = ".deop" Then
                For i = 0 To lstUsers.ListCount - 1
                    If UCase(Mid(txtChat.Text, 7, Len(txtChat.Text))) = UCase(Mid(lstUsers.List(i), 1, InStr(1, lstUsers.List(i), " ") - 1)) Then
                        recipientIndex = lstUsers.ItemData(i)
                        Exit For
                    End If
                Next i
                
                If recipientIndex > 0 And InStr(1, Ops, "," & Trim(Str(recipientIndex)) & ",") Then
                    SendToOneClient CODE_DEOPPED, recipientIndex
                    SendToOneClient CODE_MESSAGE & "You have been deopped.", recipientIndex
                    Ops = Mid(Ops, 1, InStr(1, Ops, "," & Trim(Str(recipientIndex)) & ",")) & Mid(Ops, InStr(1, Ops, "," & Trim(Str(recipientIndex))) + 2 + Len(Trim(Str(recipientIndex))), Len(Ops))
                End If
                        
            'Just a plain ol' message
            Else
                UpdateDialog txtName.Text & ": " & txtChat.Text
                SendToAllClients CODE_MESSAGE & txtName.Text & ": " & txtChat.Text
            End If
        'Message is from a client
        Else
            'Emote using /me or .me
            If Mid(txtChat.Text, 1, 3) = "/me" Or Mid(txtChat.Text, 1, 3) = ".me" Then
                SendToServer CODE_MESSAGE & "* " & txtName.Text & " " & Mid(txtChat.Text, 5, Len(txtChat.Text))
            'Emote using ;
            ElseIf Mid(txtChat.Text, 1, 1) = ";" Then
                SendToServer CODE_MESSAGE & "* " & txtName.Text & " " & Mid(txtChat.Text, 2, Len(txtChat.Text))
            'Private message using /msg or .msg
            ElseIf Mid(txtChat.Text, 1, 4) = "/msg" Or Mid(txtChat.Text, 1, 4) = ".msg" Then
                If InStr(7, txtChat.Text, " ") And InStr(7, txtChat.Text, " ") <> Len(txtChat.Text) Then
                    
                    recipient = Mid(txtChat.Text, 6, InStr(7, txtChat.Text, " ") - 6)
                    message = Mid(txtChat.Text, InStr(7, txtChat.Text, " ") + 1, Len(txtChat.Text))
                    
                    For i = 0 To lstUsers.ListCount - 1
                        If UCase(recipient) = UCase(Mid(lstUsers.List(i), 1, InStr(1, lstUsers.List(i), " ") - 1)) Then
                            userExists = True
                            Exit For
                        End If
                    Next i
                    
                    If userExists Then
                        SendToServer CODE_PRIVATE_MESSAGE & recipient & "%" & message
                        UpdateDialog "[To " & recipient & "] " & message
                    End If
                    
                End If

            'Private message using >
            ElseIf Mid(txtChat.Text, 1, 1) = ">" Then
                If InStr(4, txtChat.Text, " ") And InStr(4, txtChat.Text, " ") <> Len(txtChat.Text) Then
                
                    If Mid(txtChat.Text, 2, 1) = " " Then
                        recipient = Mid(txtChat.Text, 3, InStr(4, txtChat.Text, " ") - 3)
                        message = Mid(txtChat.Text, InStr(4, txtChat.Text, " ") + 1, Len(txtChat.Text))
                    Else
                        recipient = Mid(txtChat.Text, 2, InStr(3, txtChat.Text, " ") - 2)
                        message = Mid(txtChat.Text, InStr(3, txtChat.Text, " ") + 1, Len(txtChat.Text))
                    End If
                    
                    For i = 0 To lstUsers.ListCount - 1
                        If UCase(recipient) = UCase(Mid(lstUsers.List(i), 1, InStr(1, lstUsers.List(i), " ") - 1)) Then
                            userExists = True
                            Exit For
                        End If
                    Next i
                    
                    If userExists Then
                        SendToServer CODE_PRIVATE_MESSAGE & recipient & "%" & message
                        UpdateDialog "[To " & recipient & "] " & message
                    End If

                End If
            'Client typed /quit or .quit, so disconnect
            ElseIf Mid(txtChat.Text, 1, 5) = "/quit" Or Mid(txtChat.Text, 1, 5) = ".quit" Then
                cmdDisconnect_Click
            'Client wishes to kick (only Opped users can kick)
            ElseIf Mid(txtChat.Text, 1, 5) = "/kick" Or Mid(txtChat.Text, 1, 5) = ".kick" Then
                
                If InStr(7, txtChat.Text, " ") Then
                    recipient = Mid(txtChat.Text, 7, InStr(7, txtChat.Text, " ") - 6)
                Else
                    recipient = Mid(txtChat.Text, 7, Len(txtChat.Text))
                End If
                
                SendToServer CODE_KICK & recipient
            'Clear the chatbox
            ElseIf Mid(txtChat.Text, 1, 6) = "/clear" Or Mid(txtChat.Text, 1, 6) = ".clear" Then
                txtChatBox.Text = ""
                chatLines = 0
            'Just a plain message
            Else
                SendToServer CODE_MESSAGE & txtName.Text & ": " & txtChat.Text
            End If
        End If
    
        txtChat.Text = ""
        KeyAscii = 0
    
    End If
    End If
End Sub

Private Sub txtChatBox_KeyPress(KeyAscii As Integer)
    If Len(txtChatBox.Text) - 1 <= 0 Then
        txtChatBox.Text = ""
    Else
        txtChatBox.Text = Mid(txtChatBox.Text, 1, Len(txtChatBox.Text) - 1)
    End If
End Sub

Private Sub txtConnectTo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cmdConnect_Click
    End If
End Sub
