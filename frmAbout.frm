VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Quick Radio Ripper"
   ClientHeight    =   3240
   ClientLeft      =   195
   ClientTop       =   510
   ClientWidth     =   5175
   BeginProperty Font 
      Name            =   "Century Gothic"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   5175
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btDebug 
      Caption         =   "&Debug"
      Height          =   375
      Left            =   1680
      TabIndex        =   8
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton btClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "-v (starts minimized) -e (turns off EQ)"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1920
      Width           =   4455
   End
   Begin VB.Label Label5 
      Caption         =   "-a (url) -s (savefile) [-t (timelimit)] [-m (sizelimit)]"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   4455
   End
   Begin VB.Label Label4 
      Caption         =   "Command Line Operation:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   4335
   End
   Begin VB.Label Label3 
      Caption         =   "Special thanks to SHOUTCast Streaming Radio for making this project a possibility. "
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   4935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Copyright (C) 2006, Final Stand Productions"
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   4935
   End
   Begin VB.Label lbVersion 
      Caption         =   "Version 1.0.0"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   4935
   End
   Begin VB.Label Label1 
      Caption         =   "Quick Radio Ripper"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btClose_Click()
  Unload Me
End Sub

Private Sub btDebug_Click()
  Unload Me
  fDebug.Show
End Sub

Private Sub Form_Load()
  lbVersion.Caption = "Version: " & App.Major & "." & App.Minor & " [Build " & App.Revision & "]"
End Sub

