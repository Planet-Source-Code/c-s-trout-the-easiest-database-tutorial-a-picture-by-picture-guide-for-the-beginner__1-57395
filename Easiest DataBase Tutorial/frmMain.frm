VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMain 
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   5280
      Top             =   3000
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   855
      Left            =   7200
      TabIndex        =   18
      Top             =   2160
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1508
      _Version        =   393216
      Enabled         =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   285
      Left            =   10200
      TabIndex        =   13
      Text            =   "Text4"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   285
      Left            =   8880
      TabIndex        =   12
      Text            =   "Text3"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7560
      TabIndex        =   11
      Text            =   "Text2"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6240
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   1440
      Width           =   1215
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
      TabIndex        =   9
      Top             =   8040
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Height          =   5415
      Left            =   120
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   5355
      ScaleWidth      =   2595
      TabIndex        =   2
      Top             =   2880
      Width           =   2655
   End
   Begin VB.Label Label17 
      Caption         =   "Run the program and follow instructions page by page."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   26
      Top             =   600
      Width           =   5655
   End
   Begin VB.Label Label16 
      Caption         =   "4 TextBoxes "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9240
      TabIndex        =   25
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Label Label15 
      Caption         =   "4 Labels"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7440
      TabIndex        =   24
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label Label14 
      Caption         =   "Name Address Phone Email"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7440
      TabIndex        =   23
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Label Label13 
      Caption         =   "txtName txtAddress txtPhone txtEmail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9360
      TabIndex        =   22
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label12 
      Caption         =   "(This will make things easier when you link them to the DataField later on!)"
      Height          =   255
      Left            =   6360
      TabIndex        =   21
      Top             =   4560
      Width           =   5415
   End
   Begin VB.Label Label11 
      Caption         =   "Now you should name each TextBox and Label them."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   20
      Top             =   4320
      Width           =   5655
   End
   Begin VB.Label Label10 
      Caption         =   "For Simplicity, I will keep the default control names for the DataGrid and the Adodc controls"
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
      Left            =   7080
      TabIndex        =   19
      Top             =   3360
      Width           =   5295
   End
   Begin VB.Line Line8 
      X1              =   3240
      X2              =   5760
      Y1              =   2040
      Y2              =   6840
   End
   Begin VB.Line Line7 
      X1              =   5280
      X2              =   5880
      Y1              =   2400
      Y2              =   2880
   End
   Begin VB.Line Line6 
      X1              =   6000
      X2              =   7080
      Y1              =   2160
      Y2              =   2520
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   255
      Left            =   10200
      TabIndex        =   17
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   8880
      TabIndex        =   16
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   7560
      TabIndex        =   15
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   6240
      TabIndex        =   14
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Line Line5 
      X1              =   2880
      X2              =   4200
      Y1              =   7440
      Y2              =   7440
   End
   Begin VB.Line Line4 
      X1              =   3840
      X2              =   4560
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line3 
      X1              =   3720
      X2              =   4560
      Y1              =   1800
      Y2              =   2040
   End
   Begin VB.Line Line2 
      X1              =   3720
      X2              =   4560
      Y1              =   1800
      Y2              =   1560
   End
   Begin VB.Line Line1 
      X1              =   3840
      X2              =   4560
      Y1              =   1920
      Y2              =   2280
   End
   Begin VB.Label Label9 
      Caption         =   "1 DataGrid"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   8
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "1 Adodc"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   7
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "4 Labels"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   6
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "4 TextBox "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "2 - Add the needed controls."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   1680
      Width           =   3495
   End
   Begin VB.Label lblOpenNew 
      Caption         =   "1 - Open a new project."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1200
      Width           =   3015
   End
   Begin VB.Label lblInstructions 
      Caption         =   $"frmMain.frx":7C42
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4320
      TabIndex        =   1
      Top             =   6960
      Width           =   7935
   End
   Begin VB.Label lblTitle 
      Caption         =   "The EASIEST Database Tutorial EVER!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   0
      Top             =   120
      Width           =   7695
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdNext_Click()
Unload frmMain
frmPage2.Show
End Sub

Private Sub Form_Load()

        Me.Width = 13005
        Me.Height = 9000
        Me.Left = 2000
        Me.Top = 2000
End Sub
