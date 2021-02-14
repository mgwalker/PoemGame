VERSION 5.00
Begin VB.Form frmColors 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Set Colors..."
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   4065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   21
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox txtWTextB 
      Height          =   285
      Left            =   2400
      TabIndex        =   20
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox txtWTextG 
      Height          =   285
      Left            =   1800
      TabIndex        =   19
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox txtWTextR 
      Height          =   285
      Left            =   1200
      TabIndex        =   18
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox txtWBackB 
      Height          =   285
      Left            =   2400
      TabIndex        =   16
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox txtWBackG 
      Height          =   285
      Left            =   1800
      TabIndex        =   15
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox txtWBackR 
      Height          =   285
      Left            =   1200
      TabIndex        =   14
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   3120
      TabIndex        =   12
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox txtBGtextB 
      Height          =   285
      Left            =   2400
      TabIndex        =   11
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox txtBGtextG 
      Height          =   285
      Left            =   1800
      TabIndex        =   10
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox txtBGtextR 
      Height          =   285
      Left            =   1200
      TabIndex        =   9
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox txtBGColorB 
      Height          =   285
      Left            =   2400
      TabIndex        =   4
      Top             =   240
      Width           =   495
   End
   Begin VB.TextBox txtBGColorG 
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Top             =   240
      Width           =   495
   End
   Begin VB.CommandButton cmdApplyColor 
      Caption         =   "Apply"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox txtBGColorR 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Text Box Text:"
      Height          =   375
      Left            =   240
      TabIndex        =   17
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Text Box Background:"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Text on Background:"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Blue"
      Height          =   255
      Left            =   2400
      TabIndex        =   7
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "Green"
      Height          =   255
      Left            =   1800
      TabIndex        =   6
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Red"
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Background:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmColors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdApplyColor_Click()
'Change Background Colors
frmMain.BackColor = RGB(Val(txtBGColorR.Text), Val(txtBGColorG.Text), Val(txtBGColorB.Text))
frmMain.AboutBox.BackColor = RGB(Val(txtBGColorR.Text), Val(txtBGColorG.Text), Val(txtBGColorB.Text))
frmMain.editColors.BackColor = RGB(Val(txtBGColorR.Text), Val(txtBGColorG.Text), Val(txtBGColorB.Text))
frmMain.Label1.BackColor = RGB(Val(txtBGColorR.Text), Val(txtBGColorG.Text), Val(txtBGColorB.Text))
frmMain.Label2.BackColor = RGB(Val(txtBGColorR.Text), Val(txtBGColorG.Text), Val(txtBGColorB.Text))
frmMain.Label3.BackColor = RGB(Val(txtBGColorR.Text), Val(txtBGColorG.Text), Val(txtBGColorB.Text))
frmMain.Label5.BackColor = RGB(Val(txtBGColorR.Text), Val(txtBGColorG.Text), Val(txtBGColorB.Text))
frmMain.Label6.BackColor = RGB(Val(txtBGColorR.Text), Val(txtBGColorG.Text), Val(txtBGColorB.Text))
frmMain.Label7.BackColor = RGB(Val(txtBGColorR.Text), Val(txtBGColorG.Text), Val(txtBGColorB.Text))
frmMain.Label8.BackColor = RGB(Val(txtBGColorR.Text), Val(txtBGColorG.Text), Val(txtBGColorB.Text))
frmMain.Frame1.BackColor = RGB(Val(txtBGColorR.Text), Val(txtBGColorG.Text), Val(txtBGColorB.Text))
frmMain.opt5Points.BackColor = RGB(Val(txtBGColorR.Text), Val(txtBGColorG.Text), Val(txtBGColorB.Text))
frmMain.opt4Points.BackColor = RGB(Val(txtBGColorR.Text), Val(txtBGColorG.Text), Val(txtBGColorB.Text))
frmMain.opt3Points.BackColor = RGB(Val(txtBGColorR.Text), Val(txtBGColorG.Text), Val(txtBGColorB.Text))
frmMain.opt2Points.BackColor = RGB(Val(txtBGColorR.Text), Val(txtBGColorG.Text), Val(txtBGColorB.Text))
frmMain.opt1Point.BackColor = RGB(Val(txtBGColorR.Text), Val(txtBGColorG.Text), Val(txtBGColorB.Text))
frmMain.Label4.BackColor = RGB(Val(txtBGColorR.Text), Val(txtBGColorG.Text), Val(txtBGColorB.Text))
frmMain.Label9.BackColor = RGB(Val(txtBGColorR.Text), Val(txtBGColorG.Text), Val(txtBGColorB.Text))
frmMain.Label10.BackColor = RGB(Val(txtBGColorR.Text), Val(txtBGColorG.Text), Val(txtBGColorB.Text))
frmMain.lblCurrPlayer.BackColor = RGB(Val(txtBGColorR.Text), Val(txtBGColorG.Text), Val(txtBGColorB.Text))
frmMain.lblTopic1.BackColor = RGB(Val(txtBGColorR.Text), Val(txtBGColorG.Text), Val(txtBGColorB.Text))
frmMain.lblTopic2.BackColor = RGB(Val(txtBGColorR.Text), Val(txtBGColorG.Text), Val(txtBGColorB.Text))

'Change Background Text Colors
frmMain.ForeColor = RGB(Val(txtBGtextR.Text), Val(txtBGtextG.Text), Val(txtBGtextB.Text))
frmMain.AboutBox.ForeColor = RGB(Val(txtBGtextR.Text), Val(txtBGtextG.Text), Val(txtBGtextB.Text))
frmMain.editColors.ForeColor = RGB(Val(txtBGtextR.Text), Val(txtBGtextG.Text), Val(txtBGtextB.Text))
frmMain.Label1.ForeColor = RGB(Val(txtBGtextR.Text), Val(txtBGtextG.Text), Val(txtBGtextB.Text))
frmMain.Label2.ForeColor = RGB(Val(txtBGtextR.Text), Val(txtBGtextG.Text), Val(txtBGtextB.Text))
frmMain.Label3.ForeColor = RGB(Val(txtBGtextR.Text), Val(txtBGtextG.Text), Val(txtBGtextB.Text))
frmMain.Label5.ForeColor = RGB(Val(txtBGtextR.Text), Val(txtBGtextG.Text), Val(txtBGtextB.Text))
frmMain.Label6.ForeColor = RGB(Val(txtBGtextR.Text), Val(txtBGtextG.Text), Val(txtBGtextB.Text))
frmMain.Label7.ForeColor = RGB(Val(txtBGtextR.Text), Val(txtBGtextG.Text), Val(txtBGtextB.Text))
frmMain.Label8.ForeColor = RGB(Val(txtBGtextR.Text), Val(txtBGtextG.Text), Val(txtBGtextB.Text))
frmMain.Frame1.ForeColor = RGB(Val(txtBGtextR.Text), Val(txtBGtextG.Text), Val(txtBGtextB.Text))
frmMain.opt5Points.ForeColor = RGB(Val(txtBGtextR.Text), Val(txtBGtextG.Text), Val(txtBGtextB.Text))
frmMain.opt4Points.ForeColor = RGB(Val(txtBGtextR.Text), Val(txtBGtextG.Text), Val(txtBGtextB.Text))
frmMain.opt3Points.ForeColor = RGB(Val(txtBGtextR.Text), Val(txtBGtextG.Text), Val(txtBGtextB.Text))
frmMain.opt2Points.ForeColor = RGB(Val(txtBGtextR.Text), Val(txtBGtextG.Text), Val(txtBGtextB.Text))
frmMain.opt1Point.ForeColor = RGB(Val(txtBGtextR.Text), Val(txtBGtextG.Text), Val(txtBGtextB.Text))
frmMain.Label4.ForeColor = RGB(Val(txtBGtextR.Text), Val(txtBGtextG.Text), Val(txtBGtextB.Text))
frmMain.Label9.ForeColor = RGB(Val(txtBGtextR.Text), Val(txtBGtextG.Text), Val(txtBGtextB.Text))
frmMain.Label10.ForeColor = RGB(Val(txtBGtextR.Text), Val(txtBGtextG.Text), Val(txtBGtextB.Text))
frmMain.lblCurrPlayer.ForeColor = RGB(Val(txtBGtextR.Text), Val(txtBGtextG.Text), Val(txtBGtextB.Text))
frmMain.lblTopic1.ForeColor = RGB(Val(txtBGtextR.Text), Val(txtBGtextG.Text), Val(txtBGtextB.Text))
frmMain.lblTopic2.ForeColor = RGB(Val(txtBGtextR.Text), Val(txtBGtextG.Text), Val(txtBGtextB.Text))

'Change Textbox Background Colors
frmMain.txtCompose.BackColor = RGB(Val(txtWBackR.Text), Val(txtWBackG.Text), Val(txtWBackB.Text))
frmMain.txtDialog.BackColor = RGB(Val(txtWBackR.Text), Val(txtWBackG.Text), Val(txtWBackB.Text))
frmMain.txtIP.BackColor = RGB(Val(txtWBackR.Text), Val(txtWBackG.Text), Val(txtWBackB.Text))
frmMain.txtMessage.BackColor = RGB(Val(txtWBackR.Text), Val(txtWBackG.Text), Val(txtWBackB.Text))
frmMain.txtName.BackColor = RGB(Val(txtWBackR.Text), Val(txtWBackG.Text), Val(txtWBackB.Text))
frmMain.txtTopic1.BackColor = RGB(Val(txtWBackR.Text), Val(txtWBackG.Text), Val(txtWBackB.Text))
frmMain.txtTopic2.BackColor = RGB(Val(txtWBackR.Text), Val(txtWBackG.Text), Val(txtWBackB.Text))
frmMain.lstConnections.BackColor = RGB(Val(txtWBackR.Text), Val(txtWBackG.Text), Val(txtWBackB.Text))

'Change Textbox Text Colors
frmMain.txtCompose.ForeColor = RGB(Val(txtWTextR.Text), Val(txtWTextG.Text), Val(txtWTextB.Text))
frmMain.txtDialog.ForeColor = RGB(Val(txtWTextR.Text), Val(txtWTextG.Text), Val(txtWTextB.Text))
frmMain.txtIP.ForeColor = RGB(Val(txtWTextR.Text), Val(txtWTextG.Text), Val(txtWTextB.Text))
frmMain.txtMessage.ForeColor = RGB(Val(txtWTextR.Text), Val(txtWTextG.Text), Val(txtWTextB.Text))
frmMain.txtName.ForeColor = RGB(Val(txtWTextR.Text), Val(txtWTextG.Text), Val(txtWTextB.Text))
frmMain.txtTopic1.ForeColor = RGB(Val(txtWTextR.Text), Val(txtWTextG.Text), Val(txtWTextB.Text))
frmMain.txtTopic2.ForeColor = RGB(Val(txtWTextR.Text), Val(txtWTextG.Text), Val(txtWTextB.Text))
frmMain.lstConnections.ForeColor = RGB(Val(txtWTextR.Text), Val(txtWTextG.Text), Val(txtWTextB.Text))

'Save New Colors to prefs.ini
Open "prefs.ini" For Output As #1
    Write #1, frmMain.txtName.Text
    Write #1, frmMain.txtIP.Text
    Write #1, RGB(Val(txtBGColorR.Text), Val(txtBGColorG.Text), Val(txtBGColorB.Text))
    Write #1, RGB(Val(txtBGtextR.Text), Val(txtBGtextG.Text), Val(txtBGtextB.Text))
    Write #1, RGB(Val(txtWBackR.Text), Val(txtWBackG.Text), Val(txtWBackB.Text))
    Write #1, RGB(Val(txtWTextR.Text), Val(txtWTextG.Text), Val(txtWTextB.Text))
Close #1
End Sub

Private Sub Command1_Click()
cmdApplyColor_Click
Unload frmColors
End Sub

Private Sub Command2_Click()
Unload frmColors
End Sub

Private Sub Form_Load()
Dim BGCblues As Integer, BGCgreens As Integer, BGCreds As Integer
Dim BGFblues As Integer, BGFgreens As Integer, BGFreds As Integer
Dim WCblues As Integer, WCgreens As Integer, WCreds As Integer
Dim WFblues As Integer, WFgreens As Integer, WFreds As Integer
Dim backRGB As Long, textRGB As Long
Dim WbackRGB As Long, WtextRGB As Long

BGCblues = BGCgreens = BGCreds = 0
BGFblues = BGFgreens = BGFreds = 0
WCblues = WCgreens = WCreds = 0
WFblues = WFgreens = WFreds = 0

'Current Background Color
backRGB = frmMain.BackColor
Do While backRGB > 65535
    BGCblues = BGCblues + 1
    backRGB = backRGB - 65536
Loop
Do While backRGB > 255
    BGCgreens = BGCgreens + 1
    backRGB = backRGB - 256
Loop
BGCreds = backRGB
txtBGColorR.Text = BGCreds
txtBGColorG.Text = BGCgreens
txtBGColorB.Text = BGCblues

'Current Background Text Color
textRGB = frmMain.ForeColor
Do While textRGB > 65535
    BGFblues = BGFblues + 1
    textRGB = textRGB - 65536
Loop
Do While textRGB > 255
    BGFgreens = BGFgreens + 1
    textRGB = textRGB - 256
Loop
BGFreds = textRGB
txtBGtextR.Text = BGFreds
txtBGtextG.Text = BGFgreens
txtBGtextB.Text = BGFblues

'Current Textbox Background Color
WbackRGB = frmMain.txtCompose.BackColor
Do While WbackRGB > 65535
    WCblues = WCblues + 1
    WbackRGB = WbackRGB - 65536
Loop
Do While WbackRGB > 255
    WCgreens = WCgreens + 1
    WbackRGB = WbackRGB - 256
Loop
WCreds = WbackRGB
txtWBackR.Text = WCreds
txtWBackG.Text = WCgreens
txtWBackB.Text = WCblues

'Current Textbox Text Color
WtextRGB = frmMain.txtCompose.ForeColor
Do While WtextRGB > 65535
    WFblues = WFblues + 1
    WtextRGB = WtextRGB - 65536
Loop
Do While WtextRGB > 255
    WFgreens = WFgreens + 1
    WtextRGB = WtextRGB - 256
Loop
WFreds = WtextRGB
txtWTextR.Text = WFreds
txtWTextG.Text = WFgreens
txtWTextB.Text = WFblues

End Sub

