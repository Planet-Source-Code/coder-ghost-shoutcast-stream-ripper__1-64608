Attribute VB_Name = "modData"
Public fDebug As frmDebug

Public Stream As StreamInfo

Public CurrentPath As String

Public Type StreamInfo
  IP As String
  Port As Integer
  Path As String
  
  ICE As Boolean     'When TRUE, we have collected the ICE Header
  iHeader As String  'A copy of the ICE Header Metadata
  iMeta As String    'The ICE Metadata, ready for parsing
  iPack As Long      'How many ICE packets have we recieved?
  
  Duration As Long   'How many seconds of music we have
  Bytes As Double    'How many bytes of music we have
  
  Name As String
  URL As String
  Bitrate As Long
  
  MetaRate As Long   'How often do Metadata packets occur?
End Type

Public StreamFile As Long  'File Pointer

Public Buffer As String    'Incoming Data Buffer

Public DCache As String    'Disk Cache
Public Const MaxCache = 64000      'Write to Disk every 64K

Public Resync As Boolean           'Global RESYNC Mode. When TRUE, modifies the DataArrival activity

Public Blocks As Integer           'DEBUG: Number of block writes to disk
Public Max_Buffer As Double        'DEBUG: Peak Buffer Size
Public Average_Packet As Long      'DEBUG: Average Size of incoming packets
Public Max_Packet As Double        'DEBUG: Maximum Packet Size
Public BWritten As Double          'DEBUG: Actual Bytes written to disk
Public Resyncs As Long             'DEBUG: How many times did we have to resync?
Public RDiscard As Double          'DEBUG: How many bytes resync discarded

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Sub ShellURL(ByVal Msg As String)
    Call ShellExecute(0&, vbNullString, Msg, vbNullString, vbNullString, vbNormalFocus)
End Sub

Public Function IsCorrupt(ByVal Data As String) As Boolean
  IsCorrupt = True
  
  'Safe: 32, 39, 44 - 46, 48-57, 59, 61, 65-90, 95 97-122
  For a = 1 To Len(Data)
    Char = Asc(Mid(Data, a, 1))
    If Char <> 13 And Char <> 10 And Char <> 0 Then
      If Char > 122 Then Exit Function
      If Char < 32 Then Exit Function
    End If
  Next a
  
  IsCorrupt = False
End Function

