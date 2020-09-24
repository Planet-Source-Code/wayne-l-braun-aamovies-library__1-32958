VERSION 5.00
Begin VB.Form frmSearch 
   Caption         =   "Search Results"
   ClientHeight    =   8100
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   11820
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8100
   ScaleWidth      =   11820
   Begin VB.ListBox lstResults 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4470
      Left            =   90
      MultiSelect     =   1  'Simple
      TabIndex        =   4
      Top             =   1215
      Width           =   11200
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Close Search Results"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4590
      TabIndex        =   3
      Top             =   6345
      Width           =   2760
   End
   Begin VB.TextBox txtSearchString 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3195
      TabIndex        =   1
      Top             =   180
      Width           =   6585
   End
   Begin VB.Label lblHeadings 
      Caption         =   $"frmSearch.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   135
      TabIndex        =   2
      Top             =   765
      Width           =   11205
   End
   Begin VB.Label lblSearchString 
      Caption         =   "Results of Searching For:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   495
      TabIndex        =   0
      Top             =   225
      Width           =   2670
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdDone_Click()
   frmMoviesLibrary.txtSearch = ""
   Unload Me
End Sub



