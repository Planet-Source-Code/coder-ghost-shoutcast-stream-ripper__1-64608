VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quick Radio Ripper"
   ClientHeight    =   4785
   ClientLeft      =   2055
   ClientTop       =   2370
   ClientWidth     =   6255
   BeginProperty Font 
      Name            =   "Century Gothic"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   319
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   417
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrStats 
      Interval        =   100
      Left            =   4560
      Top             =   3960
   End
   Begin VB.Timer tmrLaunch 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   4560
      Top             =   4440
   End
   Begin VB.CommandButton btAbout 
      Caption         =   "&About"
      Height          =   375
      Left            =   4320
      TabIndex        =   22
      Top             =   2280
      Width           =   1815
   End
   Begin MSComctlLib.ProgressBar Bar 
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   2880
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.CommandButton btSave 
      Caption         =   "..."
      Height          =   375
      Left            =   5760
      TabIndex        =   10
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox txtSave 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   8
      Text            =   "..."
      Top             =   960
      Width           =   5535
   End
   Begin VB.CommandButton btRun 
      Caption         =   "&Start Ripping"
      Height          =   375
      Left            =   4320
      TabIndex        =   7
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Record Until:"
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   4095
      Begin VB.TextBox txtBreak 
         Height          =   360
         Index           =   1
         Left            =   2160
         TabIndex        =   6
         Text            =   "100"
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtBreak 
         Height          =   360
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Text            =   "60"
         Top             =   600
         Width           =   1695
      End
      Begin VB.OptionButton optBreak 
         Caption         =   "(x) megabytes:"
         Height          =   255
         Index           =   1
         Left            =   2160
         TabIndex        =   4
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton optBreak 
         Caption         =   "(x) minutes:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.TextBox txtStream 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Text            =   "http://64.236.34.4:80/stream/1003"
      Top             =   360
      Width           =   6015
   End
   Begin MSWinsockLib.Winsock WS 
      Left            =   5520
      Top             =   3960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   5040
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lbSong 
      BackStyle       =   0  'Transparent
      Caption         =   "[No song title recieved from ICE data yet]"
      Height          =   255
      Left            =   120
      MouseIcon       =   "frmMain.frx":058A
      TabIndex        =   25
      Top             =   4440
      Width           =   6015
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Stream Name And Current Song:"
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
      TabIndex        =   24
      Top             =   3960
      Width           =   3015
   End
   Begin VB.Label lbStream 
      BackStyle       =   0  'Transparent
      Caption         =   "[Waiting for ICE Header]"
      Height          =   255
      Left            =   120
      MouseIcon       =   "frmMain.frx":0894
      TabIndex        =   23
      Top             =   4200
      Width           =   6015
   End
   Begin VB.Label lbGenre 
      BackStyle       =   0  'Transparent
      Caption         =   "[Unknown]"
      Height          =   255
      Left            =   3720
      TabIndex        =   21
      Top             =   3360
      Width           =   2775
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Stream Genre:"
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
      Left            =   2280
      TabIndex        =   19
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label lbPackets 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Left            =   1320
      TabIndex        =   20
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label lbBitrate 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Left            =   5760
      TabIndex        =   18
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Stream Bitrate:"
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
      Left            =   4440
      TabIndex        =   17
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label lbTime 
      BackStyle       =   0  'Transparent
      Caption         =   "0:00"
      Height          =   255
      Left            =   3720
      TabIndex        =   16
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Audio Duration:"
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
      Left            =   2280
      TabIndex        =   14
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label lbStore 
      BackStyle       =   0  'Transparent
      Caption         =   "0 KB"
      Height          =   255
      Left            =   1320
      TabIndex        =   15
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Stored: "
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
      TabIndex        =   13
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "ICE Packets:"
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
      TabIndex        =   12
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Save Location:"
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
      TabIndex        =   9
      Top             =   720
      Width           =   6015
   End
   Begin VB.Label Label1 
      Caption         =   "Stream Address:"
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
      Width           =   6015
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   -120
      Top             =   3900
      Width           =   6495
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   1215
      Left            =   -120
      Top             =   2760
      Width           =   6495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Sample ICE Request
'GET /stream/1017 HTTP/1.1
'Accept: */*
'cache -Control: no -cache
'User-Agent: iTunes/6.0.1 (Windows; N)
'x-audiocast-udpport: 4171
'icy-metadata: 1
'Host: 64.236.34.67
'Connection: Close

'Sample ICE Header
'
'ICY 200 OK
'icy-notice1: <BR>This stream requires <a href="http://www.winamp.com/">Winamp</a><BR>
'icy-notice2: SHOUTcast Distributed Network Audio Server/SolarisSparc v1.9.5<BR>
'icy-name: D I G I T A L L Y - I M P O R T E D - DJ MIXES - non-stop DJ sets featuring various forms of techno & trance!
'icy-genre: Trance Techno House
'icy-url: http://www.di.fm
'icy-pub: 1
'icy-metaint: 8192
'icy-br: 96
'icy-irc: #shoutcast
'icy-icq: 0
'icy-aim: N/A

'SHOUTCast END OF HEADER Key
Const endHeader = vbCr & vbLf & vbCr & vbLf

Dim Running As Boolean

Private Sub btAbout_Click()
  frmAbout.Show 1, Me
End Sub

Private Sub btRun_Click()
  If Running Then
     'Close the stream
     EndStream
     
     'Close the file
     CloseFile
     
     Call SetMode(True)
     
     btRun.Caption = "&Start Ripping"
     Running = False
  Else
     CurrentPath = MakePath()
     
     'Overwrite Warning
     If StopOverwrite = True Then Exit Sub
     
     'Open the file
     OpenFile
     
     'Open the stream
     If StartStream = False Then Exit Sub
     
     Call SetMode(False)
     
     Running = True
     btRun.Caption = "&Stop Ripping"
  End If
End Sub

Public Sub SetMode(ByVal State As Boolean)
  txtSave.Enabled = State
  txtStream.Enabled = State
  txtBreak(0).Enabled = State
  txtBreak(1).Enabled = State
  optBreak(0).Enabled = State
  optBreak(1).Enabled = State
  btSave.Enabled = State
End Sub

Private Sub btSave_Click()
  On Error GoTo SaveNext

  CD.DialogTitle = "Stream Save Location"
  CD.Filter = "MP3 File (*.mp3)|*.mp3"
  
  CD.ShowSave
  txtSave.Text = CD.FileName
  
SaveNext:
  On Error GoTo 0
End Sub

Public Sub EndStream()
  On Error Resume Next
  
  'Flush the buffer
  'Put StreamFile, Loc(StreamFile) + 1, Buffer
  Call WriteToDisk(Buffer, True)
  Stream.Bytes = Stream.Bytes + Len(Buffer)
  Buffer = ""

  'Close the socket
  WS.Close
  
  'Stop the Playback
  Call Cleanup
  
  On Error GoTo 0
End Sub

Public Function MakePath() As String
  Path = txtSave.Text
  
  Path = Replace(Path, "%TIME%", Round(Timer))
  Path = Replace(Path, "%DATE%", Date$)
  
  MakePath = Path
End Function

Public Function StopOverwrite() As Boolean
  StopOverwrite = False
  On Error GoTo Safe
  
  Tmp = FreeFile
  
  Open CurrentPath For Input As Tmp
  Close Tmp
  
  Rtn = MsgBox("Warning! This file already exists. Overwrite it?", vbExclamation + vbYesNo, "Quick Radio Ripper")
  
  StopOverwrite = True
  If Rtn = vbYes Then StopOverwrite = False

Safe:
  On Error GoTo 0
End Function

Public Sub OpenFile()
  StreamFile = FreeFile
  
  Open CurrentPath For Output As StreamFile
End Sub

Public Sub CloseFile()
  Close StreamFile
End Sub

Private Sub Form_Load()
  Dim Cmds() As String
  
  Set fDebug = New frmDebug
  fDebug.Hide
  
  txtSave.Text = App.Path & "\rip.mp3"
  
  Running = False
  
  'Command Line Parsing
  If Trim(Command()) <> "" Then
     Cmds = Split(Command(), " ")

     '-a (url) -s (savefile) [-t (timelimit)] [-m (sizelimit)]
     For a = LBound(Cmds) To UBound(Cmds)
       Select Case Cmds(a)
         Case "-v":
           Me.WindowState = 1
         Case "-a":
           txtStream.Text = Cmds(a + 1)
         Case "-s":
           txtSave.Text = Cmds(a + 1)
         Case "-t":
           txtBreak(0).Text = Cmds(a + 1)
           optBreak(0).Value = True
         Case "-m":
           txtBreak(1).Text = Cmds(a + 1)
           optBreak(1).Value = True
       End Select
     Next a
     
     'Now start it up
     tmrLaunch.Enabled = True
  End If
End Sub

Public Sub Cleanup()
  'Nothing to cleanup yet
End Sub

Public Function StartStream() As Boolean
  StartStream = False
  
  'Parse the address
  Call ParseAddress
  
  'Start the Stream
  WS.Connect Stream.IP, Stream.Port

  'Set the 'Waiting for ICE' flag
  Stream.ICE = False
  
  'Clear old data
  Stream.iMeta = ""
  Stream.iPack = 0
  Stream.Bytes = 0
  Buffer = ""
  Resync = False

  Blocks = 0
  Max_Buffer = 0
  Average_Packet = 0
  Max_Packet = 0
  BWritten = 0
  Resyncs = 0
  RDiscard = 0
  
  'Generate the Playback System
  
  'Wait for connection
  Do While WS.State <> sckConnected
     If WS.State = sckError Then
        'Failed
        MsgBox "Failed to create connection to Radio Station.", vbCritical + vbOKOnly, "Quick Radio Ripper"
        Exit Function
     End If
     DoEvents
  Loop
  
  On Error GoTo 0
  
  'Build and send the Stream Request
  Header = "GET " & Stream.Path & " HTTP/1.0" & vbCrLf
  Header = Header & "Host: " & Stream.IP & vbCrLf
  Header = Header & "User-Agent: WinampMPEG/2.7" & vbCrLf
  Header = Header & "Accept: */*" & vbCrLf
  Header = Header & "Icy-MetaData: 1" & vbCrLf
  Header = Header & "Connection: Close" & vbCrLf & vbCrLf
  
  WS.SendData Header
  
  StartStream = True
End Function

Public Sub ParseAddress()
   Dim Data As String
   
   Data = Replace(txtStream.Text, "http://", "", , , vbTextCompare)
   
   PntA = InStr(1, Data, ":")
   PntB = InStr(PntA, Data, "/")
   If PntB < PntA Then PntB = Len(Data)
   
   Stream.IP = Mid(Data, 1, PntA - 1)
   Stream.Port = CInt(Mid(Data, PntA + 1, (PntB - 1) - PntA))
   Stream.Path = "/"
   If PntB <> Len(Data) Then Stream.Path = Mid(Data, PntB)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If Running Then
     'Stop the stream and close the file
     Call btRun_Click
  End If
  
  fDebug.Tag = "allow"
  Unload fDebug
End Sub

Private Sub lbSong_Click()
  If lbSong.Tag <> "" Then Call ShellURL(lbSong.Tag)
End Sub

Private Sub lbStream_Click()
  If lbStream.Tag <> "" Then Call ShellURL(lbStream.Tag)
End Sub

Private Sub tmrLaunch_Timer()
  tmrLaunch.Enabled = False
  Call btRun_Click
End Sub

Private Sub tmrStats_Timer()
  On Error Resume Next
 
  If Stream.Bytes < (950000) Then
     TmpA = Round(Stream.Bytes / 1024, 2) & " KB"
  Else
     TmpA = Round((Stream.Bytes / 1024) / 1024, 2) & " MB"
  End If
  
  'Recalculate our estimated 'music duration' in seconds
  Stream.Duration = (Stream.Bytes / ((Stream.Bitrate / 8) * 1000))
  
  TmpB = MakeTime(Stream.Duration)
  TmpC = Stream.Bitrate
  TmpD = Stream.iPack
  
  'Update the progress bar if needed
  If optBreak(0).Value = True Then
     'Max Minutes
     Progress = (Stream.Duration * 100) / (CInt(txtBreak(0).Text) * 60)
     If Progress > 100 Then Progress = 100
     If Progress < 0 Then Progress = 0
     Bar.Value = Progress
     
     'Do we need to stop?
     If Running = True Then
       If Progress = 100 And CInt(txtBreak(0).Text) > 0 Then Call btRun_Click
     End If
  Else
     'Max Filesize
     Progress = ((Stream.Bytes / 1024) * 100) / (CInt(txtBreak(1).Text) * 1024)
     If Progress > 100 Then Progress = 100
     If Progress < 0 Then Progress = 0
     Bar.Value = Progress
  
     'Do we need to stop?
     If Running = True Then
       If Progress = 100 And CInt(txtBreak(1).Text) > 0 Then Call btRun_Click
     End If
  End If
  
  'Update the necessary fields
  
  If lbStore.Caption <> TmpA Then lbStore.Caption = TmpA
  If lbTime.Caption <> TmpB Then lbTime.Caption = TmpB: Me.Caption = "(" & TmpB & ") Quick Radio Ripper"
  If lbBitrate.Caption <> TmpC Then lbBitrate.Caption = TmpC
  If lbPackets.Caption <> TmpD Then lbPackets.Caption = TmpD
  
  On Error GoTo 0
End Sub

Private Sub WS_Close()
  'Stop
End Sub

Private Sub WS_ConnectionRequest(ByVal requestID As Long)
  WS.Accept requestID
End Sub

Public Function MakeTime(ByVal Seconds As Long) As String
  Min = 0
  Sec = "00"
  
  Do While Seconds >= 60
    Seconds = Seconds - 60
    Min = Min + 1
  Loop
  
  Sec = CStr(Seconds)
  If Len(Sec) < 2 Then Sec = "0" & Sec
  
  MakeTime = Min & ":" & Sec
End Function

Private Sub WS_DataArrival(ByVal bytesTotal As Long)
  Dim inBuf As String, Dump As Long
  
  'Read the incoming data, attach it to the buffer
  WS.GetData inBuf
  Buffer = Buffer & inBuf
  
  'DEBUG: Update stats
  If Len(Buffer) > Max_Buffer Then Max_Buffer = Len(Buffer)
  If Average_Packet = 0 Then Average_Packet = Len(inBuf)
  Average_Packet = (Average_Packet + Len(inBuf)) / 2
  If Len(inBuf) > Max_Packet Then Max_Packet = Len(inBuf)
  
  'Are we still expecting a starting ICE header? Then make sure we get it before anything else.
  If Stream.ICE = False Then
     PntA = InStr(1, Buffer, endHeader, vbBinaryCompare)
     
     If PntA > 0 Then
        Stream.iMeta = Mid(Buffer, 1, (PntA + Len(endHeader)) - 1)
        Buffer = Mid(Buffer, PntA + Len(endHeader) + 0)
        Stream.iHeader = Stream.iMeta
        
        Call Parse_Header
        Stream.ICE = True
     End If
  
     Exit Sub
  End If
  
  'Are we in an emergency resync mode?
  If Resync = True Then
     'We need to stop dumping data to disk until we can locate a valid ICE packet.
     'Time to get our hands dirty with this experimental routine.
     
     'Try to locate the start of a valid Metadata packet
     PntA = InStr(1, Buffer, "StreamTitle") - 1
     
     If PntA <= 0 Then Exit Sub  'Nothing yet, keep waiting
     
     'Now we know where the metadata sync point SHOULD be. Calculate how much data we
     'need to dump to repair our stream.
     Dump = PntA
     Do
      Dump = Dump - (Stream.MetaRate + 1)
      'Debug.Print Asc(Mid(Buffer, Dump, 1))
     Loop Until Dump < Stream.MetaRate
     
     'Dump should tell us how much excess data we have. Use it to trim that data out of the buffer.
     Buffer = Mid(Buffer, Dump + 1)
     Resync = False
     
     Resyncs = Resyncs + 1
     RDiscard = RDiscard + Dump
  End If
  
  'Do we have enough data to extract the metadata packet?
  If Len(Buffer) > Stream.MetaRate Then
     'Read the METABYTE
     mTmp = Asc(Mid(Buffer, Stream.MetaRate + 1, 1)) * 16
     
     If mTmp = 0 Then
        'No metadata for this occurance. Flush the buffer up to this point and continue.
        'Put StreamFile, Loc(StreamFile) + 1, Mid(Buffer, 1, Stream.MetaRate)
        Call WriteToDisk(Mid(Buffer, 1, Stream.MetaRate), False)
        Buffer = Mid(Buffer, Stream.MetaRate + 2)
        Stream.Bytes = Stream.Bytes + (Stream.MetaRate)
     Else
        If mTmp + Stream.MetaRate > Len(Buffer) Then
           'Not enough data to read the whole meta-block, wait for more data
           Exit Sub
        End If
        
        'We have enough to read the block.  Extract the metadata block first.
        Meta = Mid(Buffer, Stream.MetaRate + 2, mTmp)
        
        'Is this Meta packet corrupted?
        If IsCorrupt(Meta) Then
           'Crap, time to start an emergency stream resync. Stop parsing, start buffering your ass off.
           Resync = True
           Exit Sub
        End If
        
        'Metadata appears to be intact, flush the buffer.
        Call WriteToDisk(Mid(Buffer, 1, Stream.MetaRate), False)
        Stream.Bytes = Stream.Bytes + (Stream.MetaRate)
        Buffer = Mid(Buffer, (Stream.MetaRate + 2) + (mTmp))
        
        Stream.iPack = Stream.iPack + 1
        
        'Parse the new metadata
        Stream.iMeta = Meta
        Call Parse_Header
     End If
  End If
End Sub

Public Sub Parse_Header()
  Dim Tmp() As String
  
  If InStr(1, Stream.iMeta, vbCrLf, vbBinaryCompare) <= 0 Then
     'Odd metadata
     'StreamTitle='Ying Yang Twins featuring Pitbull - Shake';StreamUrl='';
     
     Call Parse_Special
     Exit Sub
  End If
  
  Tmp = Split(Stream.iMeta, vbCrLf)
  
  If UBound(Tmp) <= 1 Then Exit Sub
  
  For i = 0 To UBound(Tmp)
    If Left(Tmp(i), Len("icy-url:")) = "icy-url:" Then Stream.URL = Trim(Mid(Tmp(i), Len("icy-url:") + 1)): Call SetHyperlinkA(Stream.URL)
    If Left(Tmp(i), Len("icy-br:")) = "icy-br:" Then Stream.Bitrate = CInt(Mid(Tmp(i), Len("icy-br:") + 1))
    If Left(Tmp(i), Len("icy-name:")) = "icy-name:" Then
       If Stream.ICE = False Then
          lbStream.Caption = Trim(Mid(Tmp(i), Len("icy-name:") + 1))
          lbStream.ToolTipText = lbStream.Caption
       Else
          lbSong.Caption = Trim(Mid(Tmp(i), Len("icy-name:") + 1))
          lbSong.ToolTipText = lbSong.Caption
       End If
    End If
    If Left(Tmp(i), Len("icy-genre:")) = "icy-genre:" Then lbGenre.Caption = Trim(Mid(Tmp(i), Len("icy-genre:") + 1))
    If Left(Tmp(i), Len("icy-metaint:")) = "icy-metaint:" Then
       Stream.MetaRate = CLng(Mid(Tmp(i), Len("icy-metaint:") + 1))
    End If
  Next i
End Sub

Private Sub WS_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
  'Stop
End Sub

Public Sub Parse_Special()
  On Error Resume Next
  
  'StreamTitle='Ying Yang Twins featuring Pitbull - Shake';StreamUrl='';
  Dim TmpA() As String, TmpB() As String
  
  TmpA = Split(Stream.iMeta, ";")
  
  For a = LBound(TmpA) To UBound(TmpA)
    TmpB = Split(TmpA(a), "=")
    
    Select Case LCase(TmpB(0)) 'Key
      Case "streamtitle":
        lbSong.Caption = Trim(Replace(TmpB(1), "'", ""))
      Case "streamurl":
        Call SetHyperlinkB(Trim(Replace(TmpB(1), "'", "")))
    End Select
  Next a
  
  On Error GoTo 0
End Sub

Public Sub SetHyperlinkA(ByVal Link As String)
  If Link <> "" Then
     lbStream.Tag = Link
     lbStream.ForeColor = RGB(0, 0, 255)
     lbStream.FontUnderline = True
     lbStream.MousePointer = 99
  Else
     lbStream.Tag = ""
     lbStream.ForeColor = vbBlack
     lbStream.FontUnderline = False
     lbStream.MousePointer = 0
  End If
End Sub

Public Sub SetHyperlinkB(ByVal Link As String)
  If Link <> "" Then
     lbSong.Tag = Link
     lbSong.ForeColor = RGB(0, 0, 255)
     lbSong.FontUnderline = True
     lbSong.MousePointer = 99
  Else
     lbSong.Tag = ""
     lbSong.ForeColor = vbBlack
     lbSong.FontUnderline = False
     lbSong.MousePointer = 0
  End If
End Sub

Public Sub WriteToDisk(ByVal Data As String, ByVal Flush As Boolean)
  'Enable Disk Caching
  DCache = DCache & Data
  
  'Do we have enough to justify a write?
  If Len(DCache) > MaxCache Or Flush = True Then
     Print #StreamFile, DCache;
     Blocks = Blocks + 1
     BWritten = BWritten + Len(DCache)
     DCache = ""
  End If
End Sub
