VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmPage5 
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
      Left            =   10560
      TabIndex        =   14
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
      Left            =   11640
      TabIndex        =   13
      Top             =   8040
      Width           =   975
   End
   Begin VB.PictureBox Picture5 
      Height          =   1575
      Left            =   120
      Picture         =   "frmPage5.frx":0000
      ScaleHeight     =   1515
      ScaleWidth      =   3675
      TabIndex        =   11
      Top             =   6240
      Width           =   3735
   End
   Begin VB.PictureBox Picture4 
      Height          =   3495
      Left            =   6240
      Picture         =   "frmPage5.frx":32E2
      ScaleHeight     =   3435
      ScaleWidth      =   4395
      TabIndex        =   10
      Top             =   3600
      Width           =   4455
   End
   Begin VB.PictureBox Picture3 
      Height          =   975
      Left            =   0
      Picture         =   "frmPage5.frx":B994
      ScaleHeight     =   915
      ScaleWidth      =   4755
      TabIndex        =   7
      Top             =   4080
      Width           =   4815
   End
   Begin VB.PictureBox Picture2 
      Height          =   1095
      Left            =   6840
      Picture         =   "frmPage5.frx":E036
      ScaleHeight     =   1035
      ScaleWidth      =   3915
      TabIndex        =   5
      Top             =   1560
      Width           =   3975
   End
   Begin VB.PictureBox Picture1 
      Height          =   1215
      Left            =   0
      Picture         =   "frmPage5.frx":1078C
      ScaleHeight     =   1155
      ScaleWidth      =   5475
      TabIndex        =   2
      Top             =   1320
      Width           =   5535
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   855
      Left            =   8760
      TabIndex        =   0
      Top             =   7440
      Width           =   1575
      _ExtentX        =   2778
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
      ColumnCount     =   5
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
      BeginProperty Column02 
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
      BeginProperty Column03 
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
      BeginProperty Column04 
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
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   7200
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   $"frmPage5.frx":14736
      OLEDBString     =   $"frmPage5.frx":147DB
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Contacts"
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
   Begin VB.Line Line1 
      BorderWidth     =   4
      Index           =   6
      X1              =   10920
      X2              =   10920
      Y1              =   7200
      Y2              =   5400
   End
   Begin VB.Line Line1 
      BorderWidth     =   4
      Index           =   5
      X1              =   6000
      X2              =   6000
      Y1              =   7200
      Y2              =   5400
   End
   Begin VB.Line Line1 
      BorderWidth     =   4
      Index           =   4
      X1              =   6000
      X2              =   10920
      Y1              =   7200
      Y2              =   7200
   End
   Begin VB.Line Line1 
      BorderWidth     =   4
      Index           =   3
      X1              =   10920
      X2              =   12840
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Line Line1 
      BorderWidth     =   4
      Index           =   2
      X1              =   0
      X2              =   6000
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Line Line1 
      BorderWidth     =   4
      Index           =   1
      X1              =   0
      X2              =   12840
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line1 
      BorderWidth     =   4
      Index           =   0
      X1              =   0
      X2              =   12840
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Label Label7 
      Caption         =   $"frmPage5.frx":14880
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4080
      TabIndex        =   12
      Top             =   7440
      Width           =   4575
   End
   Begin VB.Label Label6 
      Caption         =   "When you see this, select the ""RecordSource"" tab. Now click ""Next""."
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
      Left            =   6000
      TabIndex        =   9
      Top             =   3000
      Width           =   5415
   End
   Begin VB.Label Label5 
      Caption         =   "When you see this, select the path to your database. Click ""Test Connection"". Click ""Ok"""
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   840
      TabIndex        =   8
      Top             =   3120
      Width           =   3375
   End
   Begin VB.Label Label4 
      Caption         =   "When you see this, select the ""Microsoft Jet Provider"". Now click ""Next""."
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
      Left            =   6240
      TabIndex        =   6
      Top             =   840
      Width           =   5775
   End
   Begin VB.Label Label3 
      Caption         =   "On the ""RecordSource"" tab select the following. Click ""Ok"""
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
      TabIndex        =   4
      Top             =   5640
      Width           =   3975
   End
   Begin VB.Label Label2 
      Caption         =   "When you see this, click ""Build"""
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
      Left            =   840
      TabIndex        =   3
      Top             =   960
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Right click over the ""Adodc"" and click ""ADODC Properties"""
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
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   6495
   End
End
Attribute VB_Name = "frmPage5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdNext_Click()
Unload frmPage5
frmPage6.Show
End Sub

Private Sub Command1_Click()
Unload frmPage5
frmPage4.Show
End Sub

Private Sub Form_Load()
        Me.Width = 13005
        Me.Height = 9000
        Me.Left = 2000
        Me.Top = 2000
End Sub
