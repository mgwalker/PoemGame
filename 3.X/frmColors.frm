VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmColors 
   Caption         =   "Set Colors..."
   ClientHeight    =   1725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3690
   ControlBox      =   0   'False
   Icon            =   "frmColors.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1725
   ScaleWidth      =   3690
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dlgBoxText 
      Left            =   1560
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog dlgBoxBack 
      Left            =   1080
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog dlgBackText 
      Left            =   1560
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog dlgBack 
      Left            =   1080
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog dlgBoxTextOLD 
      Left            =   2760
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog dlgBoxBackOLD 
      Left            =   2280
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog dlgBackTextOLD 
      Left            =   2760
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog dlgBackOLD 
      Left            =   2280
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Reset"
      Height          =   375
      Left            =   2640
      TabIndex        =   10
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmdOkay 
      Caption         =   "Okay"
      Height          =   375
      Left            =   2640
      TabIndex        =   9
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   2640
      TabIndex        =   8
      Top             =   240
      Width           =   975
   End
   Begin VB.Label BGT 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   480
      Width           =   495
   End
   Begin VB.Label BBGC 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   840
      Width           =   495
   End
   Begin VB.Label BTC 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label BGC 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "Box Background Color:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Box Text Color:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Background Text Color:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Background Color:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmColors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Function UpdateColors()
    cmdApply_Click
    Unload frmColors
End Function

Private Sub BBGC_Click()
    dlgBoxBack.ShowColor
    BBGC.BackColor = dlgBoxBack.Color
End Sub

Private Sub BGC_Click()
    dlgBack.ShowColor
    BGC.BackColor = dlgBack.Color
End Sub

Private Sub BGT_Click()
    dlgBackText.ShowColor
    BGT.BackColor = dlgBackText.Color
End Sub

Private Sub BTC_Click()
    dlgBoxText.ShowColor
    BTC.BackColor = dlgBoxText.Color
End Sub

Private Sub cmdApply_Click()

'BACKGROUND COLORS
frmMain.BackColor = dlgBack.Color
frmColors.BackColor = dlgBack.Color
frmCompose.BackColor = dlgBack.Color
frmPoemBox.BackColor = dlgBack.Color
frmSavePoem.BackColor = dlgBack.Color
frmTopics.BackColor = dlgBack.Color
frmMain.Label1.BackColor = dlgBack.Color
frmMain.Label2.BackColor = dlgBack.Color
frmPoemBox.Label2.BackColor = dlgBack.Color
frmPoemBox.Label1.BackColor = dlgBack.Color
frmPoemBox.opt9points.BackColor = dlgBack.Color
frmPoemBox.opt8points.BackColor = dlgBack.Color
frmPoemBox.opt7points.BackColor = dlgBack.Color
frmPoemBox.opt6points.BackColor = dlgBack.Color
frmPoemBox.opt5points.BackColor = dlgBack.Color
frmPoemBox.opt4points.BackColor = dlgBack.Color
frmPoemBox.opt3points.BackColor = dlgBack.Color
frmPoemBox.opt2points.BackColor = dlgBack.Color
frmPoemBox.opt1point.BackColor = dlgBack.Color
frmPoemBox.opt0points.BackColor = dlgBack.Color
frmSavePoem.chkAppend.BackColor = dlgBack.Color
frmSavePoem.Label1.BackColor = dlgBack.Color
frmTopics.Label1.BackColor = dlgBack.Color
frmTopics.Label2.BackColor = dlgBack.Color
frmTopics.Label3.BackColor = dlgBack.Color
frmMain.PoemHist.BackColor = dlgBack.Color
frmMain.editColors.BackColor = dlgBack.Color
frmMain.abtBox.BackColor = dlgBack.Color
frmMain.Label3.BackColor = dlgBack.Color
Label1.BackColor = dlgBack.Color
Label2.BackColor = dlgBack.Color
Label3.BackColor = dlgBack.Color
Label4.BackColor = dlgBack.Color

'BACKGROUND TEXT COLORs
frmMain.ForeColor = dlgBackText.Color
frmColors.ForeColor = dlgBackText.Color
frmCompose.ForeColor = dlgBackText.Color
frmPoemBox.ForeColor = dlgBackText.Color
frmSavePoem.ForeColor = dlgBackText.Color
frmTopics.ForeColor = dlgBackText.Color
frmMain.Label1.ForeColor = dlgBackText.Color
frmMain.Label2.ForeColor = dlgBackText.Color
frmPoemBox.Label2.ForeColor = dlgBackText.Color
frmPoemBox.Label1.ForeColor = dlgBackText.Color
frmPoemBox.opt9points.ForeColor = dlgBackText.Color
frmPoemBox.opt8points.ForeColor = dlgBackText.Color
frmPoemBox.opt7points.ForeColor = dlgBackText.Color
frmPoemBox.opt6points.ForeColor = dlgBackText.Color
frmPoemBox.opt5points.ForeColor = dlgBackText.Color
frmPoemBox.opt4points.ForeColor = dlgBackText.Color
frmPoemBox.opt3points.ForeColor = dlgBackText.Color
frmPoemBox.opt2points.ForeColor = dlgBackText.Color
frmPoemBox.opt1point.ForeColor = dlgBackText.Color
frmPoemBox.opt0points.ForeColor = dlgBackText.Color
frmSavePoem.chkAppend.ForeColor = dlgBackText.Color
frmSavePoem.Label1.ForeColor = dlgBackText.Color
frmTopics.Label1.ForeColor = dlgBackText.Color
frmTopics.Label2.ForeColor = dlgBackText.Color
frmTopics.Label3.ForeColor = dlgBackText.Color
frmMain.PoemHist.ForeColor = dlgBackText.Color
frmMain.editColors.ForeColor = dlgBackText.Color
frmMain.abtBox.ForeColor = dlgBackText.Color
frmMain.Label3.ForeColor = dlgBackText.Color
frmMain.Line1.BorderColor = dlgBackText.Color
frmMain.Line2.BorderColor = dlgBackText.Color
frmMain.Line3.BorderColor = dlgBackText.Color
frmMain.Line4.BorderColor = dlgBackText.Color
frmMain.Line5.BorderColor = dlgBackText.Color
frmMain.Line6.BorderColor = dlgBackText.Color
frmMain.Line7.BorderColor = dlgBackText.Color
Label1.ForeColor = dlgBackText.Color
Label2.ForeColor = dlgBackText.Color
Label3.ForeColor = dlgBackText.Color
Label4.ForeColor = dlgBackText.Color

'BOX BACKGROUND COLORS
frmMain.txtName.BackColor = dlgBoxBack.Color
frmMain.txtConnectTo.BackColor = dlgBoxBack.Color
frmMain.txtChatBox.BackColor = dlgBoxBack.Color
frmMain.txtChat.BackColor = dlgBoxBack.Color
frmMain.lstUsers.BackColor = dlgBoxBack.Color
frmCompose.txtPoemCompose.BackColor = dlgBoxBack.Color
frmPoemBox.txtPoemBox.BackColor = dlgBoxBack.Color
frmPoemBox.lstPoemHistory.BackColor = dlgBoxBack.Color
frmSavePoem.Drive1.BackColor = dlgBoxBack.Color
frmSavePoem.Dir1.BackColor = dlgBoxBack.Color
frmSavePoem.File1.BackColor = dlgBoxBack.Color
frmSavePoem.txtFilename.BackColor = dlgBoxBack.Color
frmTopics.txtTopic1.BackColor = dlgBoxBack.Color
frmTopics.txtTopic2.BackColor = dlgBoxBack.Color
frmTopics.lstUsernames.BackColor = dlgBoxBack.Color

'BOX TEXT COLORS
frmMain.txtName.ForeColor = dlgBoxText.Color
frmMain.txtConnectTo.ForeColor = dlgBoxText.Color
frmMain.txtChatBox.ForeColor = dlgBoxText.Color
frmMain.txtChat.ForeColor = dlgBoxText.Color
frmMain.lstUsers.ForeColor = dlgBoxText.Color
frmCompose.txtPoemCompose.ForeColor = dlgBoxText.Color
frmPoemBox.txtPoemBox.ForeColor = dlgBoxText.Color
frmPoemBox.lstPoemHistory.ForeColor = dlgBoxText.Color
frmSavePoem.Drive1.ForeColor = dlgBoxText.Color
frmSavePoem.Dir1.ForeColor = dlgBoxText.Color
frmSavePoem.File1.ForeColor = dlgBoxText.Color
frmSavePoem.txtFilename.ForeColor = dlgBoxText.Color
frmTopics.txtTopic1.ForeColor = dlgBoxText.Color
frmTopics.txtTopic2.ForeColor = dlgBoxText.Color
frmTopics.lstUsernames.ForeColor = dlgBoxText.Color

End Sub

Private Sub cmdCancel_Click()

dlgBack.Color = dlgBackOLD.Color
dlgBackText.Color = dlgBackTextOLD.Color
dlgBoxBack.Color = dlgBoxBackOLD.Color
dlgBoxText.Color = dlgBoxTextOLD.Color

BGC.BackColor = dlgBack.Color
BGT.BackColor = dlgBackText.Color
BBGC.BackColor = dlgBoxBack.Color
BTC.BackColor = dlgBoxText.Color

'BACKGROUND COLORS
frmMain.BackColor = dlgBackOLD.Color
frmColors.BackColor = dlgBackOLD.Color
frmCompose.BackColor = dlgBackOLD.Color
frmPoemBox.BackColor = dlgBackOLD.Color
frmSavePoem.BackColor = dlgBackOLD.Color
frmTopics.BackColor = dlgBackOLD.Color
frmMain.Label1.BackColor = dlgBackOLD.Color
frmMain.Label2.BackColor = dlgBackOLD.Color
frmPoemBox.Label2.BackColor = dlgBackOLD.Color
frmPoemBox.Label1.BackColor = dlgBackOLD.Color
frmPoemBox.opt9points.BackColor = dlgBackOLD.Color
frmPoemBox.opt8points.BackColor = dlgBackOLD.Color
frmPoemBox.opt7points.BackColor = dlgBackOLD.Color
frmPoemBox.opt6points.BackColor = dlgBackOLD.Color
frmPoemBox.opt5points.BackColor = dlgBackOLD.Color
frmPoemBox.opt4points.BackColor = dlgBackOLD.Color
frmPoemBox.opt3points.BackColor = dlgBackOLD.Color
frmPoemBox.opt2points.BackColor = dlgBackOLD.Color
frmPoemBox.opt1point.BackColor = dlgBackOLD.Color
frmPoemBox.opt0points.BackColor = dlgBackOLD.Color
frmSavePoem.chkAppend.BackColor = dlgBackOLD.Color
frmSavePoem.Label1.BackColor = dlgBackOLD.Color
frmTopics.Label1.BackColor = dlgBackOLD.Color
frmTopics.Label2.BackColor = dlgBackOLD.Color
frmTopics.Label3.BackColor = dlgBackOLD.Color
frmMain.PoemHist.BackColor = dlgBackOLD.Color
frmMain.editColors.BackColor = dlgBackOLD.Color
frmMain.abtBox.BackColor = dlgBackOLD.Color
frmMain.Label3.BackColor = dlgBackOLD.Color
Label1.BackColor = dlgBackOLD.Color
Label2.BackColor = dlgBackOLD.Color
Label3.BackColor = dlgBackOLD.Color
Label4.BackColor = dlgBackOLD.Color

'BACKGROUND TEXT COLORS
frmMain.ForeColor = dlgBackTextOLD.Color
frmColors.ForeColor = dlgBackTextOLD.Color
frmCompose.ForeColor = dlgBackTextOLD.Color
frmPoemBox.ForeColor = dlgBackTextOLD.Color
frmSavePoem.ForeColor = dlgBackTextOLD.Color
frmTopics.ForeColor = dlgBackTextOLD.Color
frmMain.Label1.ForeColor = dlgBackTextOLD.Color
frmMain.Label2.ForeColor = dlgBackTextOLD.Color
frmPoemBox.Label2.ForeColor = dlgBackTextOLD.Color
frmPoemBox.Label1.ForeColor = dlgBackTextOLD.Color
frmPoemBox.opt9points.ForeColor = dlgBackTextOLD.Color
frmPoemBox.opt8points.ForeColor = dlgBackTextOLD.Color
frmPoemBox.opt7points.ForeColor = dlgBackTextOLD.Color
frmPoemBox.opt6points.ForeColor = dlgBackTextOLD.Color
frmPoemBox.opt5points.ForeColor = dlgBackTextOLD.Color
frmPoemBox.opt4points.ForeColor = dlgBackTextOLD.Color
frmPoemBox.opt3points.ForeColor = dlgBackTextOLD.Color
frmPoemBox.opt2points.ForeColor = dlgBackTextOLD.Color
frmPoemBox.opt1point.ForeColor = dlgBackTextOLD.Color
frmPoemBox.opt0points.ForeColor = dlgBackTextOLD.Color
frmSavePoem.chkAppend.ForeColor = dlgBackTextOLD.Color
frmSavePoem.Label1.ForeColor = dlgBackTextOLD.Color
frmTopics.Label1.ForeColor = dlgBackTextOLD.Color
frmTopics.Label2.ForeColor = dlgBackTextOLD.Color
frmTopics.Label3.ForeColor = dlgBackTextOLD.Color
frmMain.PoemHist.ForeColor = dlgBackTextOLD.Color
frmMain.editColors.ForeColor = dlgBackTextOLD.Color
frmMain.abtBox.ForeColor = dlgBackTextOLD.Color
frmMain.Label3.ForeColor = dlgBackTextOLD.Color
frmMain.Line1.BorderColor = dlgBackTextOLD.Color
frmMain.Line2.BorderColor = dlgBackTextOLD.Color
frmMain.Line3.BorderColor = dlgBackTextOLD.Color
frmMain.Line4.BorderColor = dlgBackTextOLD.Color
frmMain.Line5.BorderColor = dlgBackTextOLD.Color
frmMain.Line6.BorderColor = dlgBackTextOLD.Color
frmMain.Line7.BorderColor = dlgBackTextOLD.Color
Label1.ForeColor = dlgBackTextOLD.Color
Label2.ForeColor = dlgBackTextOLD.Color
Label3.ForeColor = dlgBackTextOLD.Color
Label4.ForeColor = dlgBackTextOLD.Color

'BOX BACKGROUND COLORS
frmMain.txtName.BackColor = dlgBoxBackOLD.Color
frmMain.txtConnectTo.BackColor = dlgBoxBackOLD.Color
frmMain.txtChatBox.BackColor = dlgBoxBackOLD.Color
frmMain.txtChat.BackColor = dlgBoxBackOLD.Color
frmMain.lstUsers.BackColor = dlgBoxBackOLD.Color
frmCompose.txtPoemCompose.BackColor = dlgBoxBackOLD.Color
frmPoemBox.txtPoemBox.BackColor = dlgBoxBackOLD.Color
frmPoemBox.lstPoemHistory.BackColor = dlgBoxBackOLD.Color
frmSavePoem.Drive1.BackColor = dlgBoxBackOLD.Color
frmSavePoem.Dir1.BackColor = dlgBoxBackOLD.Color
frmSavePoem.File1.BackColor = dlgBoxBackOLD.Color
frmSavePoem.txtFilename.BackColor = dlgBoxBackOLD.Color
frmTopics.txtTopic1.BackColor = dlgBoxBackOLD.Color
frmTopics.txtTopic2.BackColor = dlgBoxBackOLD.Color
frmTopics.lstUsernames.BackColor = dlgBoxBackOLD.Color

'BOX TEXT COLORS
frmMain.txtName.ForeColor = dlgBoxTextOLD.Color
frmMain.txtConnectTo.ForeColor = dlgBoxTextOLD.Color
frmMain.txtChatBox.ForeColor = dlgBoxTextOLD.Color
frmMain.txtChat.ForeColor = dlgBoxTextOLD.Color
frmMain.lstUsers.ForeColor = dlgBoxTextOLD.Color
frmCompose.txtPoemCompose.ForeColor = dlgBoxTextOLD.Color
frmPoemBox.txtPoemBox.ForeColor = dlgBoxTextOLD.Color
frmPoemBox.lstPoemHistory.ForeColor = dlgBoxTextOLD.Color
frmSavePoem.Drive1.ForeColor = dlgBoxTextOLD.Color
frmSavePoem.Dir1.ForeColor = dlgBoxTextOLD.Color
frmSavePoem.File1.ForeColor = dlgBoxTextOLD.Color
frmSavePoem.txtFilename.ForeColor = dlgBoxTextOLD.Color
frmTopics.txtTopic1.ForeColor = dlgBoxTextOLD.Color
frmTopics.txtTopic2.ForeColor = dlgBoxTextOLD.Color
frmTopics.lstUsernames.ForeColor = dlgBoxTextOLD.Color

End Sub

Private Sub cmdChangeFont_Click()
    dlgBackText.ShowFont
    Label5.Font = dlgBackText.FontName
    Label5.FontSize = dlgBackText.FontSize
    Label5.FontBold = dlgBackText.FontBold
    Label5.FontItalic = dlgBackText.FontItalic
    Label5.FontUnderline = dlgBackText.FontUnderline
    Label5.FontStrikethru = dlgBackText.FontStrikethru
    Label5.Caption = "Current Font: " & dlgBackText.FontName
    'cmdChangeFont.Left = Label5.Left + Label5.Width + 120
End Sub

Private Sub cmdOkay_Click()
    cmdApply_Click
    
    Open "poemgame.ini" For Output As #2
        Write #2, frmMain.txtName.Text
        Write #2, frmMain.txtConnectTo.Text
        Write #2, frmMain.BackColor
        Write #2, frmMain.ForeColor
        Write #2, frmMain.txtChatBox.BackColor
        Write #2, frmMain.txtChatBox.ForeColor
    Close #2
    
    Unload frmColors
End Sub

Private Sub Form_Load()
dlgBack.Color = frmMain.BackColor
dlgBackText.Color = frmMain.ForeColor
dlgBoxBack.Color = frmMain.txtChatBox.BackColor
dlgBoxText.Color = frmMain.txtChatBox.ForeColor

'Save the values in separate Common Dialog Boxes
'in case the user CANCELs and wants the old colors
'back *AFTER* pressing Apply.
dlgBackOLD.Color = frmMain.BackColor
dlgBackTextOLD.Color = frmMain.ForeColor
dlgBoxBackOLD.Color = frmMain.txtChatBox.BackColor
dlgBoxTextOLD.Color = frmMain.txtChatBox.ForeColor

BGC.BackColor = dlgBack.Color
BGT.BackColor = dlgBackText.Color
BBGC.BackColor = dlgBoxBack.Color
BTC.BackColor = dlgBoxText.Color

End Sub

