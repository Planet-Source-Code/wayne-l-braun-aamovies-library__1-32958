VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMoviesLibrary 
   BackColor       =   &H8000000B&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "   Movies Library"
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11850
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C00000&
   Icon            =   "Movies Library.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   4  'Icon
   ScaleHeight     =   9120
   ScaleWidth      =   11850
   Begin VB.ListBox lstTapeNums 
      Height          =   690
      Left            =   1125
      TabIndex        =   29
      Top             =   8280
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3780
      TabIndex        =   25
      Top             =   2655
      Width           =   4770
      Begin VB.CommandButton cmdLook 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Find All"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   3555
         Picture         =   "Movies Library.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   135
         Width           =   1095
      End
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFFE1&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   945
         TabIndex        =   27
         Top             =   180
         Width           =   2565
      End
      Begin VB.Label lblSearch 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Search For"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   135
         TabIndex        =   26
         Top             =   270
         Width           =   795
      End
   End
   Begin VB.ListBox lstShowtape 
      BackColor       =   &H00FFFFE1&
      Height          =   1740
      Left            =   90
      Sorted          =   -1  'True
      TabIndex        =   24
      Top             =   315
      Width           =   9870
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Edit Movie"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   8415
      Picture         =   "Movies Library.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Click to edit the information for the movie shown."
      Top             =   7425
      Width           =   1125
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Save Movie"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   6165
      Picture         =   "Movies Library.frx":0B16
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Save the movie as entered."
      Top             =   7425
      Width           =   1125
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   3330
      Top             =   8100
      Visible         =   0   'False
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   582
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
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   5670
      Top             =   8055
   End
   Begin MSDataGridLib.DataGrid DataGrid 
      Bindings        =   "Movies Library.frx":0C60
      CausesValidation=   0   'False
      Height          =   3210
      Left            =   270
      TabIndex        =   7
      Top             =   3510
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   5662
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   12648447
      Enabled         =   -1  'True
      HeadLines       =   1
      RowHeight       =   12
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "Title"
         Caption         =   "Title"
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
         DataField       =   "Desc"
         Caption         =   "Stars & Comments"
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
         DataField       =   "TapeNum"
         Caption         =   "Tape Number"
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
         DataField       =   "Rating"
         Caption         =   "Rating"
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
         DataField       =   "RunTime"
         Caption         =   "RunTime"
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
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   3915.213
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3915.213
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   659.906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1005.165
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Delete Movie"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   7290
      Picture         =   "Movies Library.frx":0C75
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Delete the movie shown from the library."
      Top             =   7425
      Width           =   1125
   End
   Begin VB.ListBox lstTapeget 
      BackColor       =   &H00FFFFC0&
      Columns         =   4
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3015
      IntegralHeight  =   0   'False
      ItemData        =   "Movies Library.frx":0DBF
      Left            =   10035
      List            =   "Movies Library.frx":0DC6
      TabIndex        =   0
      ToolTipText     =   "List of all tapes in the library"
      Top             =   270
      Width           =   1770
   End
   Begin VB.TextBox txtDesc 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   5535
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   7065
      Width           =   3983
   End
   Begin VB.TextBox txtTitle 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   320
      Left            =   1530
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   7065
      Width           =   3990
   End
   Begin VB.TextBox txtRunTime 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   10890
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   7065
      Width           =   675
   End
   Begin VB.TextBox txtRating 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   10180
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   7065
      Width           =   675
   End
   Begin VB.TextBox txtTapeNum 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   9540
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   7065
      Width           =   637
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00C0C0FF&
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   10800
      Picture         =   "Movies Library.frx":0DD6
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Exit the program."
      Top             =   7425
      Width           =   645
   End
   Begin VB.CommandButton cmdPrintLibrary 
      BackColor       =   &H00FFC0A0&
      Caption         =   "Print Library List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   1800
      Picture         =   "Movies Library.frx":0F20
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Print a cross reference for all movies in the library"
      Top             =   2340
      Width           =   1500
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Add Movie"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   5040
      Picture         =   "Movies Library.frx":17EA
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Click to add a new movie to the library."
      Top             =   7425
      Width           =   1125
   End
   Begin VB.CommandButton cmdPrintLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0A0&
      Caption         =   "Print Label For Tape"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   90
      Picture         =   "Movies Library.frx":1934
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Print a tape box label for the tape shown above."
      Top             =   2340
      Width           =   1710
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00FFFFE1&
      Caption         =   "&Next"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   5670
      Picture         =   "Movies Library.frx":21FE
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Show the next tape in the library."
      Top             =   2070
      Width           =   975
   End
   Begin VB.CommandButton cmdLast 
      BackColor       =   &H00FFFFE1&
      Caption         =   "&Last Tape"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   6660
      Picture         =   "Movies Library.frx":2788
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Show the last tape in the library."
      Top             =   2070
      Width           =   975
   End
   Begin VB.CommandButton cmdPrevious 
      BackColor       =   &H00FFFFE1&
      Caption         =   "&Previous"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   4680
      Picture         =   "Movies Library.frx":2D12
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Show the previous tape in the library."
      Top             =   2070
      Width           =   975
   End
   Begin VB.CommandButton cmdFirst 
      BackColor       =   &H00FFFFE1&
      Caption         =   "&First Tape"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   3825
      Picture         =   "Movies Library.frx":329C
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Show the first tape in the library"
      Top             =   2070
      Width           =   855
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   0
      X2              =   11835
      Y1              =   8100
      Y2              =   8100
   End
   Begin VB.Label lblTextBoxHeaders 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   $"Movies Library.frx":3826
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3285
      TabIndex        =   23
      Top             =   6840
      Width           =   9915
   End
   Begin VB.Label lblBiggestTape 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "Biggest Tape Number in Library Is"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   585
      TabIndex        =   22
      Top             =   7785
      Width           =   2940
   End
   Begin VB.Label lblMovieCount 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "altMovies in Library"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1215
      TabIndex        =   21
      Top             =   7470
      Width           =   1665
   End
   Begin VB.Label lblTapeSelect 
      Caption         =   "Click To Display ------->  Movies On Selected Tape Number "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   8325
      TabIndex        =   20
      Top             =   2070
      Width           =   1635
   End
   Begin VB.Label lblCurrent 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Selected Movie -->"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   90
      TabIndex        =   19
      ToolTipText     =   "Click a movie in the grid above to select it."
      Top             =   7065
      Width           =   1425
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      X1              =   45
      X2              =   11880
      Y1              =   3420
      Y2              =   3420
   End
   Begin VB.Label lblHeaders 
      Caption         =   $"Movies Library.frx":38AE
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
      Left            =   45
      TabIndex        =   18
      Top             =   45
      Width           =   9945
   End
End
Attribute VB_Name = "frmMoviesLibrary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Author     : Wayne Braun
' Date       : March 5, 2002
' Version    : 1.0
' Language   : MS Visual Basic 6.0
' Purpose    : Makes a library of VHS tapes of movies
' Known Bugs : N/A
  
 '---This program uses the onyx font in cmdPrintLabel.  You can change the font
 '---in that Sub if it is not on your system.  This program is not intended to be
 '---a tutorial, but an example of ADO database using a DataGrid and SQL.  Your
 '---feedback is welcome and desired, especially code improvement comments.
 
 '---This is my first Visual Basic program and credit goes to Jerry Barnes (and
 '---planet-source-code.com for providing it) for his ADO Tutorial For Absolute
 '---Beginners.  It helped with problems I was not figuring out.
 '---Credit also to Peter Wright's book "Beginning Visual Basic 6.0".
 
 '--- MoviesLibrary is a library of video tapes each containing up to 8 movies
 '---using movies.mdb and keeping track of Title (String*36),
 '---Desc (Description) (String*36), TapeNum (String*5), Rating (String*4)
 '---and RunTime (String*4) in each record.  These size limits were set in the
 '---database creation using Visual Data Manager and put in to aid printing
 '---format concerns.
  
  '---Adodc1 uses Movies.mdb as its data source file
  '---rsVideoMovies is associated with DataGrid, rsVideoMovies, txtTitle,
  '---txtDesc, etc
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Option Explicit
Option Compare Text   '---make comparisons not case sensitive
Dim I As Integer, J As Integer
Public bolEOFReached
Public rsVideoMovies As ADODB.Recordset
Attribute rsVideoMovies.VB_VarHelpID = -1
Dim txtPrintString As String, txtPrintLabelInst As String
Private Const UPDATE_CANCELLED As Long = -2147217842
Private Const ERRORS_OCCURRED As Long = -2147217887
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Public Sub Form_Load()
   Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Movies.mdb;Mode=Read|Write|Share Deny None;Persist Security Info=False"
   '---Note that in the line above, the Data Source does not specify a path for _
      Movies.mdb.  This is done so that the current directory is assumed _
      for Movies.mdb so that the program can be installed on other computers. _
      This also means that Movies.mdb must exist in the VB folder so that the _
      program can access it during development and testing.
   Adodc1.CursorLocation = adUseClient
   '---Use a client side cursor because the data you will be accessing will be on
   '---the client machine instead of a server.
   Adodc1.CommandType = adCmdText
   Adodc1.RecordSource = "movies"
   Adodc1.CursorType = adOpenStatic
   '---The only type of cursor that you can use with
   '---a client side cursor location is adOpenStatic.
   Adodc1.LockType = adLockOptimistic
   '---This guarantees that a record that is being edited can be saved
   Adodc1.RecordSource = "Select * From Movies Order By movies.Title"
   Adodc1.Refresh
   Set rsVideoMovies = Adodc1.Recordset
   '---Source is a SQL statement indicating where to retreive the data from.
   Dim intTapeNum As Integer
   intBiggestTape = 1
   ' ---now find the biggest tape number
   MoveToRecord adRsnMoveFirst
   For I = 1 To rsVideoMovies.RecordCount
      intTapeNum = Val(txtTapeNum)
      If intBiggestTape < intTapeNum Then  '---update intBiggestTape as needed
         intBiggestTape = intTapeNum
      End If
      MoveToRecord adRsnMoveNext
   Next I
   lblBiggestTape = "Biggest Tape Number in Library Is" & Str(intBiggestTape)   '---show biggest tape number
   intCurrentTape = 1
   Call Showtape     '---display movies from tape number intCurrentTape
   MoveToRecord adRsnMoveFirst
   Call FillTapeNums   '---build the listbox of tape numbers that allow user to
                                 '---select a specific tape
   lblMovieCount = Str(rsVideoMovies.RecordCount) & " Movies in Library"
                                                                   '---refresh the lblMovieCount
   '---now reset the grid by traversing it so that it doesn't roll when you click a row
   Adodc1.RecordSource = "Select * From Movies Order By movies.Title"
   Adodc1.Refresh
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub lstTapeget_Click()   '---show movies on tape # user clicks in
                                               '---1stTapeget list box
   intCurrentTape = lstTapeget.ListIndex + 1
   Call Showtape
   MoveToRecord adRsnMoveFirst               '---move the database and
   rsVideoMovies.Bookmark = rsVideoMovies.Bookmark  '---grid back to 1st record
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub cmdPrintLabel_Click()
Dim intTextIndent As Integer, intNumIndent As Integer
   txtPrintLabelInst = InputBox$("Load printer  with a label to print a label" _
   & "for tape number" & Str$(intCurrentTape) & _
   "  Using Avery #8164 labels (3 1/3 x 4 inches), then enter the postion" & _
   " of the label to print on from 1 to 6, counting from the top as you would " & _
   "read from left to right.", "Printer Setup Verification")
   DoEvents    '---refresh screen from message box
   Select Case txtPrintLabelInst
      Case "0"
         Beep
         Exit Sub
      Case "1"         '---set up the indent and # of blank lines to fit label selected
         intTextIndent = 13
         intNumIndent = 1
         intSkipLines = 2
      Case "2"
         intTextIndent = 69
         intNumIndent = 17
         intSkipLines = 2
      Case "3"
         intTextIndent = 13
         intNumIndent = 1
         intSkipLines = 27
      Case "4"
         intTextIndent = 69
         intNumIndent = 17
         intSkipLines = 27
      Case "5"
         intTextIndent = 13
         intNumIndent = 1
         intSkipLines = 52
      Case "6"
         intTextIndent = 69
         intNumIndent = 17
         intSkipLines = 52
      Case Else
      Beep
      Exit Sub
   End Select
   For I = 1 To intSkipLines
      Printer.Print
   Next I
   Printer.FontName = "onyx"
   Printer.FontSize = 76
   Printer.Print Tab(intNumIndent); Right$(Str$(intCurrentTape), 2)
   Printer.FontName = "courier new"
   Printer.FontBold = True
   Printer.FontSize = 8.5
   J = lstShowtape.ListCount - 1
   For I = 0 To J   '---loop for j movies in tape
      txtPrintString = lstShowtape.List(I)
      Printer.Print Tab(intTextIndent); Right$(txtPrintString, 4); " "; _
      Mid$(txtPrintString, 6, 36)
      Printer.Print Tab(intTextIndent); Mid$(txtPrintString, 82, 4); " "; _
      Mid$(txtPrintString, 45, 36)
   Next I
   Printer.EndDoc
   MoveToRecord adRsnMoveFirst                                      '---move the database and
   rsVideoMovies.Bookmark = rsVideoMovies.Bookmark  '---grid back to 1st record
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub cmdPrintLibrary_Click()
Dim FontSave As String, SizeSave As String
Dim intLines As Integer, intResponse As Integer
bolEOFReached = False    '---must make sure flag false before print traverse
intResponse = MsgBox("Print the Music Library?", vbQuestion + vbOKCancel)
If intResponse = vbCancel Then
   Exit Sub
End If
FontSave = Printer.FontName
SizeSave = Printer.FontSize
Printer.FontName = "arial"
Printer.FontSize = 9
Printer.FontBold = True
MoveToRecord adRsnMoveFirst
Do
   For I = 1 To 4     '---print 4 blank lines
      Printer.Print
   Next I
   For intLines = 1 To 64   '---print 64 text lines per page
      Printer.Print Tab(4); Left$(txtTitle, 36); Tab(53); " | "; Left$(txtDesc, 36); _
      Tab(106); "| "; Left$(txtTapeNum & "   ", 5); Tab(114); " | "; _
      Left$(txtRating, 4); Tab(122); "| "; Left$(txtRunTime, 4)
      MoveToRecord adRsnMoveNext
      If bolEOFReached = True Then
         MoveToRecord adRsnMoveFirst
         Exit Do
      End If
      Next intLines
      Printer.NewPage
   Loop
   Printer.FontName = FontSave
   Printer.FontBold = False
   Printer.FontSize = SizeSave
   Printer.EndDoc
   MoveToRecord adRsnMoveFirst                                      '---move the database and
   rsVideoMovies.Bookmark = rsVideoMovies.Bookmark  '---grid back to 1st record
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub cmdFirst_Click()    '---show movies on first tape
   intCurrentTape = 1
   Call Showtape
   MoveToRecord adRsnMoveFirst                   '---move the database and
   rsVideoMovies.Bookmark = rsVideoMovies.Bookmark  '---grid back to 1st record
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub cmdPrevious_Click()     '---show movies on previous tape
   intCurrentTape = intCurrentTape - 1
   If intCurrentTape = 0 Then
      intCurrentTape = 1
      Beep
      Exit Sub
   End If
   Call Showtape
   MoveToRecord adRsnMoveFirst                                      '---move the database and
   rsVideoMovies.Bookmark = rsVideoMovies.Bookmark  '---grid back to 1st record
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub cmdNext_Click()      '---show movies on next tape
   intCurrentTape = intCurrentTape + 1
   If intCurrentTape > intBiggestTape Then
      intCurrentTape = intBiggestTape
      Beep
      Exit Sub
   End If
   Call Showtape
   MoveToRecord adRsnMoveFirst                                      '---move the database and
   rsVideoMovies.Bookmark = rsVideoMovies.Bookmark  '---grid back to 1st record
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub cmdLast_Click()     '---show movies on last tape
   intCurrentTape = intBiggestTape
   Call Showtape
   MoveToRecord adRsnMoveFirst                                      '---move the database and
   rsVideoMovies.Bookmark = rsVideoMovies.Bookmark  '---grid back to 1st record
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub DataGrid_HeadClick(ByVal ColIndex As Integer)
   '---using DataGrid.Columns(ColIndex).Datafield in place of "movies.title" _
      portion of select statement worked for all columns except desc (Stars _
      and Comments).  This caused a syntax error which could not be _
      tracked down.  Code adding select case below is a workaround.
   Select Case ColIndex
      Case 0
         Adodc1.RecordSource = "Select * From Movies Order By " & _
           "movies.Title"
         Adodc1.Refresh
      Case 1
         Adodc1.RecordSource = "Select * From Movies Order By " & _
           "movies.Desc"
         Adodc1.Refresh
      Case 2
         Adodc1.RecordSource = "Select * From Movies Order By " & _
           "movies.TapeNum"
         Adodc1.Refresh
      Case 3
         Adodc1.RecordSource = "Select * From Movies Order By " & _
           "movies.Rating"
         Adodc1.Refresh
      Case 4
         Adodc1.RecordSource = "Select * From Movies Order By " & _
           "movies.RunTime"
         Adodc1.Refresh
   End Select
   SetGridAndBoxes    '---move grid to selected record and fill movie boxes
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub cmdLook_Click()
   If rsVideoMovies.RecordCount = 0 Then      '---make sure the library not empty
      MsgBox "There are no movies in the library."
      txtSearch.Text = ""
   Else
      If txtSearch.Text = "" Then    '---check for a search string entered
         MsgBox "Enter a search string."
         txtSearch.SetFocus
         Exit Sub
      End If
      frmSearch.lstResults.Clear      '---clear any old search results
      MoveToRecord adRsnMoveFirst
      For I = 1 To rsVideoMovies.RecordCount
         '---now add all movies containing search string to the list box.
         If InStr(1, txtTitle, txtSearch) <> 0 Or InStr(1, txtDesc, txtSearch) <> 0 Then
            frmSearch.lstResults.AddItem Left$(txtTitle & Space$(36), 36) & _
            "   " & Left$(txtDesc & Space$(36), 36) & " " & _
            Left$(txtTapeNum & Space$(5), 5) & "        " & _
            Left$(txtRating & Space$(4), 4) & "    " & _
            Left$(txtRunTime & Space$(4), 4)
         End If
         MoveToRecord adRsnMoveNext
      Next I    '---loop thru movies to search
      MoveToRecord adRsnMoveFirst  '---reset to first movie @ end of search
      frmSearch.Show      '---show the results form
   End If
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub cmdAdd_Click()  '---cmdAdd is dual function depending on its
                                          '---caption set to Add Movie or Cancel
   If cmdAdd.Caption = "&Add Movie" Then
      cmdAdd.Caption = "&Cancel Add"   '---reconfigure button to allow cancel
                                                   '---the add in progress
      cmdAdd.ToolTipText = "Click to cancel the Add Movie in progress."
      MoveToRecord adRsnMoveFirst    '---now build a list of tape numbers in use
      For I = 1 To rsVideoMovies.RecordCount
         lstTapeNums.AddItem (RTrim(LTrim(txtTapeNum)))
         MoveToRecord adRsnMoveNext
      Next I    '---loop thru movies
      cmdSave.Enabled = True   '---to allow new movie to be saved
      Call DisableNavigation   '---disable everything but add functions
      Call ClearControls    '---clear the 5 text boxes so ready fo add movie
      cmdEdit.Enabled = False   '---no edit allowed during an add
      cmdDelete.Enabled = False  '---no delete allowed during an add
      txtTitle.Locked = False   '---unlock the 5 text boxes
      txtDesc.Locked = False
      txtTapeNum.Locked = False
      txtRating.Locked = False
      txtRunTime.Locked = False
      txtTitle.SetFocus   '---Go to the first field.
   ElseIf cmdAdd.Caption = "&Cancel Add" Then
      cmdAdd.Caption = "&Add Movie"    '---If a user cancels, allow another add.
      cmdAdd.ToolTipText = "Click to add a new movie to the library."
      cmdSave.Enabled = False    '---Because of cancel, no record to save.
      Call EnableNavigation     '---The user should have freedom to move now.
      cmdEdit.Enabled = True
      cmdDelete.Enabled = True
      txtTitle.Locked = True   '---lock the 5 text boxes
      txtDesc.Locked = True
      txtTapeNum.Locked = True
      txtRating.Locked = True
      txtRunTime.Locked = True
      MoveToRecord adRsnMoveFirst
      Call LoadDataInControls    '---reload the 5 text boxes
   End If
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub cmdSave_Click()
    If cmdAdd.Caption = "&Cancel Add" Then
        MoveToRecord adRsnAddNew
    End If
    Call WriteDataFromControls   '---Write the data from the text boxes to
                                                '---the appropriate fields.
    rsVideoMovies.Update     '---No data is saved until the update method
                                         '--- is executed.
    cmdSave.Enabled = False   '---Turn off the Save button
    cmdAdd.Caption = "&Add Movie"      '---Change captions back to original.
    cmdEdit.Caption = "&Edit Movie"     '---as needed
    txtTitle.Locked = True   '---lock the 5 text boxes
    txtDesc.Locked = True
    txtTapeNum.Locked = True
    txtRating.Locked = True
    txtRunTime.Locked = True
    Call EnableNavigation     '---Allow normal program function again
    cmdEdit.Enabled = True
    cmdAdd.Enabled = True
    cmdDelete.Enabled = True
    rsVideoMovies.Close   '---these 2 lines get the new movie
    rsVideoMovies.Open     '---to show in the grid
    lblMovieCount = Str(rsVideoMovies.RecordCount) & " Movies in Library"
                                                            '---refresh the lblMovieCount
   DataGrid_HeadClick (0)  '---resort the grid by title
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub cmdDelete_Click()
   If rsVideoMovies.EOF = False And rsVideoMovies.BOF = False Then
      'Check to see if there is data in the database
      On Error Resume Next    '---If there is an error, ignore it.
      If MsgBox("Delete the selected movie:" + vbCrLf + txtTitle + " " _
           + txtDesc + txtTapeNum + " " + txtRating + " " + txtRunTime, _
           vbQuestion + vbYesNo, "Delete Movie Verify") = vbYes Then
         rsVideoMovies.Delete
         rsVideoMovies.Update
         rsVideoMovies.Close   '---close,open needed to get grid and recordset
         rsVideoMovies.Open   '---records in sync
         lblMovieCount = Str(rsVideoMovies.RecordCount) & " Movies in Library"
                                                            '---refresh the lblMovieCount
         '---these next 3 lines give feedback that the movie was deleted & also
          '---purge the datagrid of the deleted row by re-sorting the grid and then
          '---resorting it back to listing by Title
         Call DataGrid_HeadClick(4)
         MsgBox "The movie was deleted."  '
         Call DataGrid_HeadClick(0)
      Else
         MoveToRecord adRsnMovePrevious
      End If
      MoveToRecord adRsnMoveNext
      If rsVideoMovies.EOF = True Then   '---if last record was deleted, go to
         MoveToRecord adRsnMoveLast    '---the record that preceeded it.
         If rsVideoMovies.BOF = True Then
            Call ClearControls  '---If the last record deleted, clear the text boxes
            MsgBox "There is no data in the recordset!"
         End If
      End If
   ElseIf rsVideoMovies.EOF = True And rsVideoMovies.BOF = True Then
        '---Warn the user that attempt is being made to delete data
        '---from a database with no records.
        MsgBox "There is no data in the recordset!"
   End If
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub cmdEdit_Click()   '---cmdEdit is dual function depending on its
                                          '---caption set to Edit Movie or Cancel Edit
   If cmdEdit.Caption = "&Edit Movie" Then
        cmdEdit.Caption = "&Cancel Edit"  '---reconfigure button to allow cancel
                                                   '---the edit in progress
        cmdEdit.ToolTipText = "Click to cancel the edits in progress."
        cmdSave.Enabled = True    '---enable save for record being edited
        Call DisableNavigation     '---disable everything but edit functions
        cmdAdd.Enabled = False   '---no add allowed during an edit
        cmdDelete.Enabled = False   '---delete allowed during an edit
        txtTitle.Locked = False   '---unlock the 5 text boxes
        txtDesc.Locked = False
        txtTapeNum.Locked = False
        txtRating.Locked = False
        txtRunTime.Locked = False
        txtTitle.SetFocus   '---Go to the first field.
   ElseIf cmdEdit.Caption = "&Cancel Edit" Then
        cmdEdit.Caption = "&Edit Movie"   '---If a user cancels, allow another edit.
        cmdEdit.ToolTipText = "Click to edit the information for the movie shown."
        cmdSave.Enabled = False    '---because of cancel, no record to save
        Call EnableNavigation      '---The user should have freedom to move now.
        cmdAdd.Enabled = True
        cmdDelete.Enabled = True
        txtTitle.Locked = True   '---lock the 5 text boxes
    txtDesc.Locked = True
    txtTapeNum.Locked = True
    txtRating.Locked = True
    txtRunTime.Locked = True
    Call LoadDataInControls    '---reload the 5 text boxes.
    End If
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub cmdExit_Click()
   End
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub SetGridAndBoxes()
   Dim ClickedTitle As String
   ClickedTitle = DataGrid.Columns("Title")  '---save the clicked title
   MoveToRecord adRsnMoveFirst  '---now move to that record
   For I = 1 To rsVideoMovies.RecordCount
      If rsVideoMovies!Title = ClickedTitle Then '---when found exit loop
         Exit For
      End If
      MoveToRecord adRsnMoveNext
   Next I
   Call LoadDataInControls    '---fill the text boxes
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub SetGrid()     '---move the grid to the current movie
   varCurrentRecord = rsVideoMovies.Bookmark
   DataGrid.Bookmark = varCurrentRecord
   Call LoadDataInControls    '---fill the text boxes
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub DataGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   Call SetGridAndBoxes     '---move grid to selected record and fill movie boxes
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub ClearControls()
   txtTitle = ""
   txtDesc = ""
   txtTapeNum = ""
   txtRating = ""
   txtRunTime = ""
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub LoadDataInControls()    '---fill 5 boxes for selected movie
   If rsVideoMovies.BOF = True Or rsVideoMovies.EOF Then
      Exit Sub
   End If
   txtTitle.Text = Left$(rsVideoMovies!Title & " ", 36)
   txtDesc.Text = Left$(rsVideoMovies!Desc & " ", 36)
   txtTapeNum.Text = Left$(rsVideoMovies!Tapenum & " ", 5)
   txtRating.Text = Left$(rsVideoMovies!Rating & " ", 4)
   txtRunTime.Text = Left$(rsVideoMovies!RunTime & " ", 4)
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub WriteDataFromControls()    '---write info in 5 text boxes to database
    rsVideoMovies("Title").Value = Left$(txtTitle.Text, 36)   '---truncating for the
    rsVideoMovies!Desc = Left$(txtDesc.Text, 36)    '---database field size
    rsVideoMovies!Tapenum = Left$(txtTapeNum.Text, 5)   '---limits.
    rsVideoMovies!Rating = Left$(txtRating.Text, 4)
    rsVideoMovies!RunTime = Left$(txtRunTime.Text, 4)
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub DisableNavigation()  '---prevent navigation when adding or editing
    cmdFirst.Enabled = False
    cmdLast.Enabled = False
    cmdNext.Enabled = False
    cmdPrevious.Enabled = False
    cmdPrintLabel.Enabled = False
    cmdPrintLibrary.Enabled = False
    lstTapeget.Enabled = False
    DataGrid.Enabled = False
    txtSearch.Locked = True
    cmdLook.Enabled = False
    cmdExit.Enabled = False
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub EnableNavigation()  '---resume normal functions after add or edit
    cmdFirst.Enabled = True
    cmdLast.Enabled = True
    cmdNext.Enabled = True
    cmdPrevious.Enabled = True
    cmdPrintLabel.Enabled = True
    cmdPrintLibrary.Enabled = True
    lstTapeget.Enabled = True
    DataGrid.Enabled = True
    txtSearch.Locked = False
    cmdLook.Enabled = True
    cmdExit.Enabled = True
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub Showtape()   '---display tape intCurrentTape in tape window
   lstShowtape.Clear
   MoveToRecord adRsnMoveFirst
   For I = 1 To rsVideoMovies.RecordCount
      If Val(txtTapeNum) = intCurrentTape Then
         lstShowtape.AddItem Left$(txtTapeNum & Space$(5), 5) & " " & _
         Left$(txtTitle & Space$(36), 36) & _
         "   " & Left$(txtDesc & Space$(36), 36) & "  " & _
         Left$(txtRating & Space$(4), 4) & " " & _
         Left$(txtRunTime & Space$(4), 4)
      End If
      MoveToRecord adRsnMoveNext
   Next I
   MoveToRecord adRsnMovePrevious
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub FillTapeNums()   '---build the list box of tape numbers user can
                                           '---click to display
   Dim strBlank As String
   lstTapeget.Clear
   For I = 1 To intBiggestTape
      If I < 10 Then strBlank = "  " Else strBlank = " "
      lstTapeget.AddItem strBlank & (Val(I))
   Next I
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub txtSearch_Change()  '---initiate to 1st record anytime txtSearch
                                                '---changes
   MoveToRecord adRsnMoveFirst
   cmdLook.Default = True
   bolEOFReached = False
   frmSearch.txtSearchString = txtSearch
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub MoveToRecord(intDirection As Integer)
'---Here in the MoveToRecord Procedure, the ADO constant (like
'---adRsnMoveFirst) is passed to intDirection which is declared here
'---to centralize all the Move commands.  This also centralizes error handling.
   On Error GoTo MoveToRecord_Err
   Select Case intDirection
      Case adRsnMoveFirst
         rsVideoMovies.MoveFirst
         Call LoadDataInControls      '---fill the text boxes
      Case adRsnMovePrevious
         rsVideoMovies.MovePrevious
         If rsVideoMovies.BOF Then
            rsVideoMovies.MoveNext
         End If
         Call LoadDataInControls      '---fill the text boxes
      Case adRsnMoveNext
         rsVideoMovies.MoveNext
         If rsVideoMovies.EOF Then
            bolEOFReached = True
            rsVideoMovies.MovePrevious
         End If
         Call LoadDataInControls      '---fill the text boxes
      Case adRsnMoveLast
         rsVideoMovies.MoveLast
         Call LoadDataInControls      '---fill the text boxes
      Case adRsnAddNew
         '---AddNew method automatically creates a new blank record
         rsVideoMovies.AddNew
   End Select
MoveToRecord_Exit:
   Exit Sub
MoveToRecord_Err:
   Select Case Err.Number
      Case UPDATE_CANCELLED, ERRORS_OCCURRED
      '---do nothing
      Case Else
         Err.Raise Err.Number, Err.Source, Err.Description
   End Select
   Resume MoveToRecord_Exit
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub cmdNext_MouseDown(Button As Integer, Shift As Integer, _
X As Single, Y As Single)
'--- the Timer, MouseDown, and Mouseup events allow Previous or Next buttons
'---to be held down to scroll the records
   Timer1.Enabled = True
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub cmdNext_MouseUp(Button As Integer, Shift As Integer, _
X As Single, Y As Single)
   Timer1.Enabled = False
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub cmdPrevious_MouseDown(Button As Integer, Shift As Integer, _
X As Single, Y As Single)
   Timer1.Enabled = True
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub cmdPrevious_MouseUp(Button As Integer, Shift As Integer, _
X As Single, Y As Single)
   Timer1.Enabled = False
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub Timer1_Timer()
'--- the Timer, MouseDown, and Mouseup events allow Previous or Next buttons
'---to be held down to scroll the records
'---The Screen object refers to the currently active form and ActiveControl is the
'---current control that has the focus.  Thus if the control is the Next button, we
'---call the Click procedure for the Next button.  If Timer is enabled, each time the
'---Timer fires (as set by Timer.interval) this Click procedure is called.
'---Thus for both Next and Previous buttons, we enable the
'---Timer on MouseDown and disable it on MouseUp

If Screen.ActiveControl.Name = "cmdNext" Then
   cmdNext_Click  '---only cmdNext and cmdPrevious enable Timer
Else
   cmdPrevious_Click
End If
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub txtTapeNum_Validate(Keepfocus As Boolean)
   If InStr(txtTapeNum, "-") = 0 Then
      Keepfocus = True
      MsgBox "Tape number must include the dash between" & _
         " the tape number and the numerical position on the tape" & _
         "   Example 23-2", vbExclamation, "Tape Number Format"
   End If
   If Mid$(txtTapeNum, 1, 1) = " " Then txtTapeNum = LTrim(txtTapeNum)
   If Mid$(txtTapeNum, 2, 1) = "-" Then
       txtTapeNum = Left$(" " & txtTapeNum, 5) '---if tape # is single digit _
                                              pad a blank on front to align with others
   End If
   For I = 0 To rsVideoMovies.RecordCount - 1
      If txtTapeNum = lstTapeNums.List(I) Then
         Keepfocus = True
         MsgBox "You entered " & txtTapeNum & " for a tape number." & vbCrLf _
                     & "That tape number is already in use."
         Exit For
      End If
   Next I
End Sub
'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub txtRunTime_Validate(Keepfocus As Boolean)
   If InStr(txtRunTime, ":") = 0 Then
       Keepfocus = True
       MsgBox "Run Time must include the colon between" & _
         " the hour and the minutes     Example  1:47", vbExclamation, _
         "Run Time Format"
   End If
End Sub


