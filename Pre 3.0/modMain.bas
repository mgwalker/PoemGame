Attribute VB_Name = "Module1"
Option Explicit

'This is the port to be used for establishing connections.
Global glPort As Long
Global randomPlayer As String, randPlayerNumber As Integer
Global isFirstRound As Boolean
Global currPoem As String
Global newPoem As String
Global minutesPassed As Integer
Global currentPlayer As Integer

Global totalPoints(100), totalVotes(100)
Global votePoints, finalScore(100)
Global TopicHoldSeconds As Integer
Global usersAway As Integer
Global AwayUsers(100) As Integer
Global totalChatLines As Integer

'API calls used for reading and writing of preferences
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
