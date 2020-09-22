VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmDebug 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quick Radio Ripper - Debugging"
   ClientHeight    =   5400
   ClientLeft      =   285
   ClientTop       =   600
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
   Icon            =   "frmDebug.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   5175
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btShow 
      Caption         =   "&Show Header"
      Height          =   375
      Left            =   1800
      TabIndex        =   19
      Top             =   4920
      Width           =   1575
   End
   Begin MSComctlLib.ProgressBar Bar_Cache 
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   1920
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.TextBox txtMeta 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Top             =   4200
      Width           =   4935
   End
   Begin VB.Timer tmrUpdate 
      Interval        =   250
      Left            =   4680
      Top             =   4800
   End
   Begin MSComctlLib.ProgressBar Bar_Buffer 
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.CommandButton btHide 
      Caption         =   "&Close"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label lbResync 
      BackStyle       =   0  'Transparent
      Caption         =   "Stream is no longer in sync, attempting to recover."
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   3720
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Label lbSync 
      Caption         =   "Stream appears to be syncronized with the metadata."
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   3720
      Width           =   4935
   End
   Begin VB.Label Label7 
      Caption         =   "Stream Resync:"
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
      TabIndex        =   22
      Top             =   3240
      Width           =   4935
   End
   Begin VB.Label lbSyncStat 
      Caption         =   "0 bytes have been discarded in 0 resyncs. "
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   3480
      Width           =   4935
   End
   Begin VB.Label lbRealBytes 
      Caption         =   "0 bytes"
      Height          =   255
      Left            =   2760
      TabIndex        =   17
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label lbDCache 
      Caption         =   "0 bytes / 0 bytes"
      Height          =   255
      Left            =   1560
      TabIndex        =   14
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label lbValid 
      BackStyle       =   0  'Transparent
      Caption         =   "Metadata appears to contain valid characters."
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2880
      Width           =   4935
   End
   Begin VB.Label lbPacket_Max 
      Caption         =   "0 bytes"
      Height          =   255
      Left            =   2760
      TabIndex        =   10
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label lbBlock_Writes 
      Caption         =   "0"
      Height          =   255
      Left            =   2760
      TabIndex        =   9
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Number of Block Writes to Disk:"
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
      TabIndex        =   8
      Top             =   2280
      Width           =   4815
   End
   Begin VB.Label lbPacket_Size 
      Caption         =   "0 bytes"
      Height          =   255
      Left            =   2760
      TabIndex        =   7
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Average Incoming Packet Size:"
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
      TabIndex        =   6
      Top             =   1080
      Width           =   4815
   End
   Begin VB.Label lbBuffer_Peak 
      Caption         =   "0 bytes"
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Buffer Peak:"
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
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label lbBuffer_Current 
      Caption         =   "0 bytes"
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Buffer Size Monitor:"
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
      Width           =   5055
   End
   Begin VB.Label Label6 
      Caption         =   "Maximum Incoming Packet Size:"
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
      TabIndex        =   11
      Top             =   1320
      Width           =   4815
   End
   Begin VB.Label Label8 
      Caption         =   "Disk Cache Size:"
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
      TabIndex        =   15
      Top             =   1680
      Width           =   4815
   End
   Begin VB.Label Label9 
      Caption         =   "Actual Bytes Written To Disk:"
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
      TabIndex        =   18
      Top             =   2520
      Width           =   4815
   End
   Begin VB.Label lbInvalid 
      BackStyle       =   0  'Transparent
      Caption         =   "Metadata contains invalid characters, stream is corrupt."
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   2880
      Visible         =   0   'False
      Width           =   4935
   End
End
Attribute VB_Name = "frmDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btHide_Click()
  Me.Visible = False
End Sub

Private Sub btShow_Click()
  If btShow.Caption = "&Show Header" Then
     'Show the Header
     btShow.Caption = "&Show Metadata"
  Else
     'Show the current metadata
     btShow.Caption = "&Show Header"
  End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If Me.Tag <> "allow" Then
     Cancel = 1
     Me.Visible = False
  End If
End Sub

Private Sub tmrUpdate_Timer()
  On Error Resume Next
  
  TmpA = Len(Buffer) & " bytes"
  TmpB = Max_Buffer & " bytes"
  TmpC = Round((Len(Buffer) * 100) / Max_Buffer)
  TmpD = Average_Packet & " bytes"
  TmpE = Blocks
  TmpF = Max_Packet & " bytes"
  TmpG = Len(DCache) & " bytes / " & MaxCache & " bytes"
  TmpH = Round((Len(DCache) * 100) / MaxCache)
  TmpI = RDiscard & " bytes have been discarded in " & Resyncs & " resyncs. "
  TmpJ = BWritten & " bytes"
  
  On Error GoTo 0
  
  If TmpC > 100 Then TmpC = 100
  If TmpC < 0 Then TmpC = 0
  If TmpH > 100 Then TmpH = 100
  If TmpH < 0 Then TmpH = 0
  
  If Bar_Buffer.Value <> TmpC Then Bar_Buffer.Value = TmpC
  If lbBuffer_Current.Caption <> TmpA Then lbBuffer_Current.Caption = TmpA
  If lbBuffer_Peak.Caption <> TmpB Then lbBuffer_Peak.Caption = TmpB
  If lbPacket_Size.Caption <> TmpD Then lbPacket_Size.Caption = TmpD
  If lbPacket_Max.Caption <> TmpF Then lbPacket_Max.Caption = TmpF
  If lbBlock_Writes.Caption <> TmpE Then lbBlock_Writes.Caption = TmpE
  If lbDCache.Caption <> TmpG Then lbDCache.Caption = TmpG
  If Bar_Cache.Value <> TmpH Then Bar_Cache.Value = TmpH
  If lbRealBytes.Caption <> TmpJ Then lbRealBytes.Caption = TmpJ
  If lbSyncStat.Caption <> TmpI Then lbSyncStat.Caption = TmpI
  
  If btShow.Caption = "&Show Header" Then
    If txtMeta.Text <> Stream.iMeta Then txtMeta.Text = Stream.iMeta
  Else
    If txtMeta.Text <> Stream.iHeader Then txtMeta.Text = Stream.iHeader
  End If
  
  If IsCorrupt(Stream.iMeta) Then
     If lbInvalid.Visible = False Then lbInvalid.Visible = True: lbValid.Visible = False
  Else
     If lbValid.Visible = False Then lbValid.Visible = True: lbInvalid.Visible = False
  End If
  
  lbSync.Visible = Not Resync
  lbResync.Visible = Resync
End Sub
