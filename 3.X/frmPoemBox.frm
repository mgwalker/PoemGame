VERSION 5.00
Begin VB.Form frmPoemBox 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " The PoemGame - Poem"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6585
   Icon            =   "frmPoemBox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   6585
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton opt9points 
      Caption         =   "9 Points"
      Height          =   315
      Left            =   120
      TabIndex        =   15
      Top             =   2520
      Width           =   975
   End
   Begin VB.OptionButton opt7points 
      Caption         =   "7 Points"
      Height          =   315
      Left            =   120
      TabIndex        =   14
      Top             =   3000
      Width           =   975
   End
   Begin VB.OptionButton opt6points 
      Caption         =   "6 Points"
      Height          =   315
      Left            =   120
      TabIndex        =   13
      Top             =   3240
      Width           =   975
   End
   Begin VB.OptionButton opt8points 
      Caption         =   "8 Points"
      Height          =   315
      Left            =   120
      TabIndex        =   12
      Top             =   2760
      Width           =   975
   End
   Begin VB.ListBox lstPoemHistory 
      Height          =   2010
      ItemData        =   "frmPoemBox.frx":0C42
      Left            =   5040
      List            =   "frmPoemBox.frx":0C49
      TabIndex        =   11
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton cmdSavePoem 
      Caption         =   "Save Poem"
      Height          =   375
      Left            =   3720
      TabIndex        =   8
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdScore 
      Caption         =   "Submit Score"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      Top             =   2520
      Width           =   1215
   End
   Begin VB.OptionButton opt3points 
      Caption         =   "3 Points"
      Height          =   315
      Left            =   1320
      TabIndex        =   6
      Top             =   2760
      Width           =   975
   End
   Begin VB.OptionButton opt1point 
      Caption         =   "1 Point"
      Height          =   315
      Left            =   1320
      TabIndex        =   5
      Top             =   3240
      Width           =   975
   End
   Begin VB.OptionButton opt2points 
      Caption         =   "2 Points"
      Height          =   315
      Left            =   1320
      TabIndex        =   4
      Top             =   3000
      Width           =   975
   End
   Begin VB.OptionButton opt4points 
      Caption         =   "4 Points"
      Height          =   315
      Left            =   1320
      TabIndex        =   3
      Top             =   2520
      Width           =   975
   End
   Begin VB.OptionButton opt0points 
      Caption         =   "0 Points"
      Height          =   315
      Left            =   1320
      TabIndex        =   2
      Top             =   3480
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.OptionButton opt5points 
      Caption         =   "5 Points"
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox txtPoemBox 
      Height          =   2010
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   4815
   End
   Begin VB.Label Label2 
      Caption         =   "Current Poem:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   4815
   End
   Begin VB.Label Label1 
      Caption         =   "Poem History:"
      Height          =   255
      Left            =   5040
      TabIndex        =   9
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmPoemBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim currentPoemIndex As Integer
Dim numPoems As Integer
Dim Poem(9) As String
Dim Author(9) As String

Public Function SavePoemData() As String

    SavePoemData = "By: " & Author(currentPoemIndex) & " (" & Format(Date, "m/d/yyyy") & ")" & vbNewLine & vbNewLine & Poem(currentPoemIndex)

End Function
Public Function RefreshForm()
    Dim i As Integer
    currentPoemIndex = 0
    numPoems = frmMain.GetNumPoems
    
    If numPoems = 0 Then cmdSavePoem.Enabled = False
    
    For i = 0 To numPoems
        Author(i) = Mid(frmMain.GetPoemData(i), 1, InStr(1, frmMain.GetPoemData(i), "|") - 1)
        lstPoemHistory.List(i) = Author(i)
        Poem(i) = Mid(frmMain.GetPoemData(i), InStr(1, frmMain.GetPoemData(i), "|") + 1, Len(frmMain.GetPoemData(i)))
    Next i

    txtPoemBox.Text = Poem(0)
    Label2.Caption = Label2.Caption & " By " & Author(0)

End Function

Private Sub cmdSavePoem_Click()
    frmSavePoem.Show
End Sub

Private Sub cmdScore_Click()

    If opt9points.Value = True Then
        frmMain.ProcessScore 9
    ElseIf opt8points.Value = True Then
        frmMain.ProcessScore 8
    ElseIf opt7points.Value = True Then
        frmMain.ProcessScore 7
    ElseIf opt6points.Value = True Then
        frmMain.ProcessScore 6
    ElseIf opt5points.Value = True Then
        frmMain.ProcessScore 5
    ElseIf opt4points.Value = True Then
        frmMain.ProcessScore 4
    ElseIf opt3points.Value = True Then
        frmMain.ProcessScore 3
    ElseIf opt2points.Value = True Then
        frmMain.ProcessScore 2
    ElseIf opt1point.Value = True Then
        frmMain.ProcessScore 1
    Else
        frmMain.ProcessScore 0
    End If
    
    cmdScore.Enabled = False

End Sub

Private Sub Form_Load()
    Dim i As Integer
    currentPoemIndex = 0
    numPoems = frmMain.GetNumPoems
    
    If numPoems = 0 Then cmdSavePoem.Enabled = False
    
    For i = 0 To numPoems
        Author(i) = Mid(frmMain.GetPoemData(i), 1, InStr(1, frmMain.GetPoemData(i), "|") - 1)
        lstPoemHistory.List(i) = Author(i)
        Poem(i) = Mid(frmMain.GetPoemData(i), InStr(1, frmMain.GetPoemData(i), "|") + 1, Len(frmMain.GetPoemData(i)))
    Next i

    txtPoemBox.Text = Poem(0)
    Label2.Caption = Label2.Caption & " By " & Author(0)
End Sub

Private Sub lstPoemHistory_Click()
    
    currentPoemIndex = lstPoemHistory.ListIndex
    txtPoemBox.Text = Poem(currentPoemIndex)
    Label2.Caption = "Current Poem: By " & Author(currentPoemIndex)
    cmdScore.Enabled = False

End Sub
