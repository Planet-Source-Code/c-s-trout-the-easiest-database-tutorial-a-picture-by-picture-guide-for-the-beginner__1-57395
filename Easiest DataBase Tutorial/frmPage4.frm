VERSION 5.00
Begin VB.Form frmPage4 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "EASY DataBase"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   12915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "< Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10680
      TabIndex        =   9
      Top             =   8040
      Width           =   975
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next >"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11760
      TabIndex        =   8
      Top             =   8040
      Width           =   975
   End
   Begin VB.PictureBox Picture4 
      Height          =   3015
      Left            =   6840
      Picture         =   "frmPage4.frx":0000
      ScaleHeight     =   2955
      ScaleWidth      =   4995
      TabIndex        =   6
      Top             =   4800
      Width           =   5055
   End
   Begin VB.PictureBox Picture3 
      Height          =   2295
      Left            =   1440
      Picture         =   "frmPage4.frx":8512
      ScaleHeight     =   2235
      ScaleWidth      =   2715
      TabIndex        =   4
      Top             =   5280
      Width           =   2775
   End
   Begin VB.PictureBox Picture2 
      Height          =   2535
      Left            =   8880
      Picture         =   "frmPage4.frx":13854
      ScaleHeight     =   2475
      ScaleWidth      =   2835
      TabIndex        =   3
      Top             =   480
      Width           =   2895
   End
   Begin VB.PictureBox Picture1 
      Height          =   4335
      Left            =   1920
      Picture         =   "frmPage4.frx":176B6
      ScaleHeight     =   4275
      ScaleWidth      =   2955
      TabIndex        =   2
      Top             =   120
      Width           =   3015
   End
   Begin VB.Line Line1 
      BorderWidth     =   4
      X1              =   0
      X2              =   12840
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Label Label2 
      Caption         =   "Click ""Build the Table"" and then close the window."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5400
      TabIndex        =   7
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "This Is what you should have now."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   5
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "This window will appear. And we named the ""Table Name"", Contacts"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label7 
      Caption         =   $"frmPage4.frx":1E560
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   5400
      TabIndex        =   0
      Top             =   960
      Width           =   3495
   End
End
Attribute VB_Name = "frmPage4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdNext_Click()
Unload frmPage4
frmPage5.Show
End Sub

Private Sub Command1_Click()
Unload frmPage4
frmPage3.Show
End Sub

Private Sub Form_Load()
        Me.Width = 13005
        Me.Height = 9000
        Me.Left = 2000
        Me.Top = 2000
End Sub
