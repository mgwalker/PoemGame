VERSION 5.00
Begin VB.Form frmSavePoem 
   Caption         =   "Form1"
   ClientHeight    =   2370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5145
   LinkTopic       =   "Form1"
   ScaleHeight     =   2370
   ScaleWidth      =   5145
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkAppend 
      Caption         =   "Append to Current File"
      Height          =   255
      Left            =   3120
      TabIndex        =   6
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   4200
      TabIndex        =   5
      Top             =   1920
      Width           =   855
   End
   Begin VB.FileListBox File1 
      Height          =   1455
      Left            =   2280
      Pattern         =   "*.txt"
      TabIndex        =   3
      Top             =   480
      Width           =   2775
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
   Begin VB.TextBox txtFilename 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   1920
      Width           =   2775
   End
   Begin VB.DirListBox Dir1 
      Height          =   1440
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Save as Filename:"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   1965
      Width           =   1335
   End
End
Attribute VB_Name = "frmSavePoem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSave_Click()

    If txtFilename.Text <> "" Then
        'APPEND to an existing file
        If chkAppend.Value = 1 Then
            If UCase(Left(txtFilename.Text, 4)) <> ".TXT" Then
                Open Dir1.Path & "\" & txtFilename.Text & ".txt" For Append As #1
                    Print #1, vbNewLine & vbNewLine & vbNewLine & frmPoemBox.SavePoemData
                Close #1
            Else
                Open Dir1.Path & "\" & txtFilename.Text For Append As #2
                    Print #2, vbNewLine & vbNewLine & vbNewLine & frmPoemBox.SavePoemData
                Close #2
            End If
        
        'Overwrite or create file
        Else
            If UCase(Left(txtFilename.Text, 4)) <> ".TXT" Then
                Open Dir1.Path & "\" & txtFilename.Text & ".txt" For Output As #3
                    Print #3, frmPoemBox.SavePoemData
                Close #3
            Else
                Open Dir1.Path & "\" & txtFilename.Text For Output As #4
                    Print #4, frmPoemBox.SavePoemData
                Close #4
            End If
        End If
    Unload frmSavePoem
    Else
        txtFilename.SetFocus
    End If

End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive & "\"
File1.Path = Drive1.Drive & "\"
End Sub

Private Sub File1_Click()
    txtFilename.Text = File1.FileName
End Sub

Private Sub Form_Load()

Drive1.Drive = "c:"
Dir1.Path = "\"
File1.Path = "\"

End Sub


Private Sub txtFilename_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        cmdSave_Click
    End If
    
End Sub
