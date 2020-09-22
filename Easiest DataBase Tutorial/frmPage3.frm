VERSION 5.00
Begin VB.Form frmPage3 
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
      TabIndex        =   12
      Top             =   8040
      Width           =   975
   End
   Begin VB.PictureBox Picture5 
      Height          =   3255
      Left            =   5640
      Picture         =   "frmPage3.frx":0000
      ScaleHeight     =   3195
      ScaleWidth      =   2355
      TabIndex        =   10
      Top             =   5040
      Width           =   2415
   End
   Begin VB.PictureBox Picture4 
      Height          =   2415
      Left            =   8640
      Picture         =   "frmPage3.frx":4402
      ScaleHeight     =   2355
      ScaleWidth      =   4155
      TabIndex        =   8
      Top             =   2880
      Width           =   4215
   End
   Begin VB.PictureBox Picture3 
      Height          =   2295
      Left            =   360
      Picture         =   "frmPage3.frx":AF28
      ScaleHeight     =   2235
      ScaleWidth      =   4395
      TabIndex        =   7
      Top             =   2880
      Width           =   4455
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
      TabIndex        =   5
      Top             =   8040
      Width           =   975
   End
   Begin VB.PictureBox Picture2 
      Height          =   1815
      Left            =   3480
      Picture         =   "frmPage3.frx":108B2
      ScaleHeight     =   1755
      ScaleWidth      =   9075
      TabIndex        =   3
      Top             =   240
      Width           =   9135
   End
   Begin VB.PictureBox Picture1 
      Height          =   1815
      Left            =   480
      Picture         =   "frmPage3.frx":1C3B0
      ScaleHeight     =   1755
      ScaleWidth      =   2595
      TabIndex        =   1
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label5 
      Caption         =   "Right click on ""Properties"" and select ""New Table"""
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2640
      TabIndex        =   11
      Top             =   6600
      Width           =   2655
   End
   Begin VB.Line Line6 
      BorderWidth     =   4
      X1              =   8280
      X2              =   8280
      Y1              =   4800
      Y2              =   5760
   End
   Begin VB.Line Line5 
      BorderWidth     =   4
      X1              =   5400
      X2              =   5400
      Y1              =   4800
      Y2              =   5760
   End
   Begin VB.Line Line4 
      BorderWidth     =   4
      X1              =   5400
      X2              =   8280
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line3 
      BorderWidth     =   4
      X1              =   0
      X2              =   5400
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line Line2 
      BorderWidth     =   4
      X1              =   8280
      X2              =   12840
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Label Label4 
      Caption         =   "This window will appear."
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
      Left            =   5880
      TabIndex        =   9
      Top             =   3840
      Width           =   2655
   End
   Begin VB.Line Line1 
      BorderWidth     =   4
      X1              =   0
      X2              =   12840
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label Label3 
      Caption         =   "Name the ""New DataBase"" as ""MyContacts.mdb"" and save it to the project folder."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   5160
      Width           =   5175
   End
   Begin VB.Label Label2 
      Caption         =   "(In my case it's the ""Version 7.0 MDB...)"
      Height          =   255
      Left            =   6480
      TabIndex        =   4
      Top             =   2520
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Go to ""File"", ""New"" and select the latest ""Microsoft Access"" database."
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
      Left            =   4440
      TabIndex        =   2
      Top             =   2160
      Width           =   7335
   End
   Begin VB.Label Label7 
      Caption         =   "This is the ""Visual Data Manager"""
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
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   3495
   End
End
Attribute VB_Name = "frmPage3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdNext_Click()
Unload frmPage3
frmPage4.Show
End Sub

Private Sub Command1_Click()
Unload frmPage3
frmPage2.Show
End Sub

Private Sub Form_Load()
        Me.Width = 13005
        Me.Height = 9000
        Me.Left = 2000
        Me.Top = 2000
End Sub
