VERSION 5.00
Begin VB.Form frmCompose 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Compose Your Poem..."
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5040
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   5040
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSubmitPoem 
      Caption         =   "Submit Poem"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txtPoemCompose 
      Height          =   1695
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "frmCompose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdSubmitPoem_Click()

    If txtPoemCompose.Text <> "" Then
        frmMain.ProcessPoem
        
        frmMain.tmrTopicHold.Interval = Int(Len(txtPoemCompose.Text) / 4)
        If frmMain.tmrTopicHold.Interval < 30 Then frmMain.tmrTopicHold.Interval = 30
        frmMain.tmrTopicHold.Interval = frmMain.tmrTopicHold.Interval * 1000
        frmMain.tmrTopicHold.Enabled = True
    
        Unload frmCompose
    End If

End Sub
