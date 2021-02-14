VERSION 5.00
Begin VB.Form frmPorts 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ports"
   ClientHeight    =   1065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   Icon            =   "frmPorts.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1065
   ScaleWidth      =   3975
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   2640
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Port for Establishing Connections:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2370
   End
End
Attribute VB_Name = "frmPorts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
Unload Me
End Sub
Private Sub cmdOK_Click()
'OK was clicked.

'Redefine the settings.
glPort = txtPort.Text

Unload Me
End Sub
Private Sub Form_Load()
'Show the current settings in the text boxes.
txtPort.Text = glPort
End Sub
