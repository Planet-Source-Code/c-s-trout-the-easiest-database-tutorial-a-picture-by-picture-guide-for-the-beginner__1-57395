VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmPage7 
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
      TabIndex        =   19
      Top             =   8040
      Width           =   975
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   255
      Left            =   4440
      TabIndex        =   15
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton cmdFirstRecord 
      Caption         =   "|< First"
      Height          =   255
      Left            =   5520
      TabIndex        =   14
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton cmdLastRecord 
      Caption         =   "Last >|"
      Height          =   255
      Left            =   8040
      TabIndex        =   13
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton cmdPrevRecord 
      Caption         =   "Previous"
      Height          =   255
      Left            =   6360
      TabIndex        =   12
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton cmdNextRecord 
      Caption         =   "Next"
      Height          =   255
      Left            =   7200
      TabIndex        =   11
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   255
      Left            =   9120
      TabIndex        =   10
      Top             =   2760
      Width           =   735
   End
   Begin VB.TextBox txtEmail 
      DataField       =   "cEmail"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   8640
      TabIndex        =   4
      Top             =   2280
      Width           =   1695
   End
   Begin VB.TextBox txtTelephone 
      DataField       =   "cTelephone"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   6840
      TabIndex        =   3
      Top             =   2280
      Width           =   1695
   End
   Begin VB.TextBox txtAddress 
      DataField       =   "cAddress"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   5160
      TabIndex        =   2
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox txtName 
      DataField       =   "cName"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   3840
      TabIndex        =   1
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
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
      TabIndex        =   0
      Top             =   8040
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmPage7.frx":0000
      Height          =   2655
      Left            =   2040
      TabIndex        =   9
      Top             =   3480
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   4683
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
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
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "cName"
         Caption         =   "Name"
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
         DataField       =   "cAddress"
         Caption         =   "Address"
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
         DataField       =   "cTelephone"
         Caption         =   "Telephone"
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
         DataField       =   "cEmail"
         Caption         =   "Email"
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
            ColumnWidth     =   2009.764
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3809.764
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2009.764
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   5520
      Top             =   6240
      Visible         =   0   'False
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
      EOFAction       =   2
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"frmPage7.frx":0015
      OLEDBString     =   $"frmPage7.frx":00BA
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
   Begin VB.Label Label9 
      BackColor       =   &H80000009&
      Caption         =   "If this code has been helpful - PLEASE VOTE!                         "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   8040
      TabIndex        =   21
      Top             =   7560
      Width           =   4815
   End
   Begin VB.Line Line1 
      X1              =   11520
      X2              =   11760
      Y1              =   2280
      Y2              =   3240
   End
   Begin VB.Label Label8 
      Caption         =   "Leave extra room for the scrollbar when the DataGrid fills."
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
      Left            =   10560
      TabIndex        =   20
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label Label7 
      Caption         =   $"frmPage7.frx":015F
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   0
      TabIndex        =   18
      Top             =   6120
      Width           =   3615
   End
   Begin VB.Label Label6 
      Caption         =   "We add our code, size our controls, and have a working app!"
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
      Left            =   3240
      TabIndex        =   17
      Top             =   360
      Width           =   6375
   End
   Begin VB.Label Label5 
      Caption         =   "We set the UGLY ""Adodc"" control to ""Visible=False"" And we code the Command buttons to do the work"
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
      Left            =   0
      TabIndex        =   16
      Top             =   2280
      Width           =   3975
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Email"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8880
      TabIndex        =   8
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Telephone"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7080
      TabIndex        =   7
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5400
      TabIndex        =   6
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
   End
End
Attribute VB_Name = "frmPage7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDelete_Click()
' If you add a new record and leave it blank when you close, you will get a "RunTime Error"_
' that informs you that an "Emtpy row cannot be inserted into DataGrid"_
' This is my "shortcut" to avoid that annoyance_
' I add data into txtName field when cmdDelete is clicked
If txtName.Text = "" Then
MsgBox "Empty row cannot be inserted, Enter a field!", vbOKOnly
Exit Sub
End If
' This will delete the selected item from the database and the DataGrid
  Adodc1.Recordset.Delete
End Sub

Private Sub cmdFirstRecord_Click()
' This moves to the first record in your database that is shown in the DataGrid on top
Adodc1.Recordset.MoveFirst
End Sub

Private Sub cmdLastRecord_Click()
' This moves to the last record in your database that is shown in the DataGrid
Adodc1.Recordset.MoveLast
End Sub

Private Sub cmdNew_Click()
' This adds a new BLANK record at the end of the DataGrid_
Adodc1.Recordset.AddNew
End Sub


Private Sub cmdClose_Click()
' This unloads the form
Unload Me
End Sub

Private Sub cmdNextRecord_Click()
' This moves to the next record in your database that is shown next down in the DataGrid
Adodc1.Recordset.MoveNext
End Sub

Private Sub cmdPrevRecord_Click()
' This moves to the previous record in your database that is shown next up in the DataGrid
Adodc1.Recordset.MovePrevious
End Sub



Private Sub Command1_Click()
' This just unloaded the previous form and has nothing to do with the database code
Unload frmPage7
frmPage6.Show
End Sub

Private Sub Form_Load()
' I added this form sizing and positioning code on each form. This was only to keep each
' form uniform and has nothing to do with the database code
        Me.Width = 13005
        Me.Height = 9000
        Me.Left = 2000
        Me.Top = 2000
End Sub



Private Sub Form_Unload(Cancel As Integer)
' If you add a new record and leave it blank when you close, you will get an error msg_
' that informs you that an "Emtpy row cannot be inserted into DataGrid"_
' This is my "shortcut" to avoid that annoyance_
' I add data into a field if it is blank when the form is closed
If txtName.Text = "" Then
txtName.Text = "Blank"
End If
' This simply updates the database
Adodc1.Recordset.Update
End Sub
