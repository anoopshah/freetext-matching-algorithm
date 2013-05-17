VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Freetext Matching Algorithm Version 15"
   ClientHeight    =   7965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10410
   LinkTopic       =   "Form1"
   ScaleHeight     =   7965
   ScaleWidth      =   10410
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox entry_freetext 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      ToolTipText     =   "Sample free text for testing"
      Top             =   7320
      Width           =   8055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "START"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8280
      TabIndex        =   6
      Top             =   6960
      Width           =   1935
   End
   Begin VB.TextBox entry_infile 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      ToolTipText     =   "Filename of text file containing free text"
      Top             =   6000
      Width           =   10215
   End
   Begin VB.TextBox entry_logfile 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      ToolTipText     =   "Filename (with full fiilepath) for the output log file."
      Top             =   4680
      Width           =   10215
   End
   Begin VB.TextBox entry_outfile 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      ToolTipText     =   "Filename (with filepath) for the output comma separated values file."
      Top             =   3360
      Width           =   10215
   End
   Begin VB.TextBox entry_medcodefile 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      ToolTipText     =   "Full filename (with filepath) for the file of medcodes corresponding to each text to be analysed."
      Top             =   2040
      Width           =   10215
   End
   Begin VB.TextBox entry_lookups 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      ToolTipText     =   "Path to folder containing lookup tables. Leave blank if the lookups are in the same folder as the program itself."
      Top             =   360
      Width           =   10215
   End
   Begin VB.Label Label8 
      Caption         =   "Free text to analyse (for testing)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   6960
      Width           =   10215
   End
   Begin VB.Label Label5 
      Caption         =   "Free text full filename (tab separated pracid, textid, text)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   5640
      Width           =   10215
   End
   Begin VB.Label Label4 
      Caption         =   "Logfile full filename"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   4320
      Width           =   10215
   End
   Begin VB.Label Label3 
      Caption         =   "Output full filename"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   3000
      Width           =   10215
   End
   Begin VB.Label Label2 
      Caption         =   "Medcode full filename (comma separated with columns pracid, textid, medcode; sorted by pracid and textid)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   10215
   End
   Begin VB.Label Label1 
      Caption         =   "Path to lookups folder"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   10215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
' Start the process using the information in the form
Dim lookupfolder As String
lookupfolder = entry_lookups.text
If lookupfolder = "" Then
    lookupfolder = App.Path
End If

fma_gold.do_analysis logfile:=entry_logfile.text, _
    lookups:=lookupfolder, _
    infile:=entry_infile.text, _
    medcodefile:=entry_medcodefile.text, _
    outfile:=entry_outfile.text, _
    freetext:=entry_freetext.text

MsgBox ("Analysis complete")
    
End Sub


