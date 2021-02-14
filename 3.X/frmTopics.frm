VERSION 5.00
Begin VB.Form frmTopics 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Topics..."
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4545
   ControlBox      =   0   'False
   Icon            =   "frmTopics.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   4545
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSubmitTopics 
      Caption         =   "Submit Topics"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1605
      Width           =   1455
   End
   Begin VB.TextBox txtTopic2 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   2175
   End
   Begin VB.TextBox txtTopic1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2175
   End
   Begin VB.ListBox lstUsernames 
      Height          =   1620
      ItemData        =   "frmTopics.frx":0C42
      Left            =   2400
      List            =   "frmTopics.frx":0C44
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Topic 2:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Topic 1:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Next Player:"
      Height          =   255
      Left            =   2400
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmTopics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdSubmitTopics_Click()

If Len(txtTopic1.Text) > 25 Then
    MsgBox "Topic1 is too long!", vbExclamation, "Error"
ElseIf Len(txtTopic1.Text) > 25 Then
    MsgBox "Topic1 is too long!", vbExclamation, "Error"
ElseIf lstUsernames.List(lstUsernames.ListIndex) = frmMain.txtName.Text Then
    MsgBox "Cannot select self!", vbExclamation, "Error"
ElseIf lstUsernames.SelCount = 0 Then
    MsgBox "Must select someone!", vbExclamation, "Error"
Else
    frmMain.ProcessTopics
    frmPoemBox.RefreshForm
    frmPoemBox.cmdScore.Enabled = False
    Unload frmTopics
End If
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim a As Integer

a = 0

For i = 0 To frmMain.lstUsers.ListCount - 1
    If i + a <= frmMain.lstUsers.ListCount - 1 Then
        If Mid(frmMain.lstUsers.List(i), 1, 1) <> "-" Then
            lstUsernames.List(i) = Mid(frmMain.lstUsers.List(i + a), 1, InStr(1, frmMain.lstUsers.List(i + a), " ") - 1)
            lstUsernames.ItemData(i) = frmMain.lstUsers.ItemData(i + a)
        Else
            a = a + 1
            lstUsernames.List(i) = Mid(frmMain.lstUsers.List(i + a), 1, InStr(1, frmMain.lstUsers.List(i + a), " ") - 1)
            lstUsernames.ItemData(i) = frmMain.lstUsers.ItemData(i + a)
        End If
    Else
        Exit For
    End If
Next i

End Sub

