VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDatafile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ---------------------------------------------------
' DATAFILE CLASS (clsDatafile.cls) v1.1
' By Paul Berlin, June 2002
' berlin_paul@hotmail.com
'
' Description:
' Reads/Writes data (Strings, Bytes & Longs) to an
' binary file. Very easy & quick to use.
'
' You are free to rewrite or otherwise do as you wish
' with this code. If you make this code any better
' in any way, it would be nice if you e-mail it to me =).
' ---------------------------------------------------
' NEW IN V1.1
' * Added WriteStrFixed & ReadStrFixed.
' * Added OpenFile. You have to call OpenFile before you
'   read/write. This is MUCH faster than what I did
'   before (opened & closed the file after each
'   read/write), but made things a bit more complicated.
'   Oh well... =).
' * Fixed an bug in Get Filename.
' ---------------------------------------------------
'
' PROPERTIES:
'
' Filename
'   Get/Set the current filename to read/write to.
'
' Position
'   Get/Set the current file position. This is changed
'   automatically when you read/write, so you do not have
'   to change this if you are reading/writing from beginning
'   to end.
'
' METHODS:
'
' OpenFile
'   This has to be called after you setup the filename & before
'   you read/write anything.
'
' ReadLong As Long
'   This reads a series of 4 bytes from the current filename,
'   at the current file position, and converts and returns
'   them to an Long.
'
' ReadStr As String
'   This reads a string from the current filename, at the
'   current file position, and returns it.
'
' ReadStrFixed(vSize As Long) As String
'   This reads an fixed size string from the current file, at
'   the current file position, and returns it. vSize is the size
'   of the string to read. This is useful for file headers etc.
'
' ReadByte As Byte
'   This reads an byte from the current filename, at the
'   current file position, and returns it.
'
' WriteLong(wLong As Long)
'   This writes writes an long(wLong) as 4 bytes to the
'   current filename, at the current file position.
'
' WriteStr(wString As String)
'   This writes an string to the current filename, at the
'   current file position. The string cannot be longer
'   than 255 characters.
'
' WriteStrFixed(wString As String)
'   This writes an fixed size string to the current
'   filename, at the current file position. The string
'   can be as long as you want.
'
' WriteByte(wByte As Byte)
'   This writes an byte to the current filename, at the
'   current file position.
'
' ---------------------------------------------------
'
' THINGS TO THINK ABOUT
' There are no errorchecking or errorhandling in this
' class, so you should make sure:
' * That the filename, when reading, exists.
' * That the string when using WriteStr isn't longer
'   than 255 characters.
'
' Look in frmExample.frm for more help on how to use
' the class.
'
' ---------------------------------------------------
Option Explicit

Private strFilename   As String 'Filename
Private lngPosition   As Long   'File Position
Private Free          As Integer 'File number

Public Function ReadLong() As Long
  'This function reads an Long from the current file
  'position and returns it.
  '-------------------------------------------------
  Dim vByte(3) As Byte
  
  Get Free, lngPosition, vByte: lngPosition = lngPosition + 4
  
  ReadLong = Byte2Long(vByte())

End Function

Public Function ReadStr() As String
  'This function reads an string from the current file
  'position and returns it.
  '---------------------------------------------------
  Dim vSize As Byte   'The size of the string
  
  Get Free, lngPosition, vSize: lngPosition = lngPosition + 1
  ReadStr = Space(vSize) 'Setup the size of the string to read into
  Get Free, lngPosition, ReadStr: lngPosition = lngPosition + vSize
  
End Function

Public Function ReadStrFixed(vSize As Long) As String
  'This function reads an fixed size string from the
  'current file position and returns it. vSize is the
  'size of the string to read.
  '---------------------------------------------------
  
  ReadStrFixed = Space(vSize) 'Setup the size of the string to read into
  Get Free, lngPosition, ReadStrFixed: lngPosition = lngPosition + vSize
  
End Function

Public Function ReadByte() As Byte
  'This function reads an byte from the current file
  'position and returns it.
  '-------------------------------------------------
  
  Get Free, lngPosition, ReadByte: lngPosition = lngPosition + 1
  
End Function

Public Sub WriteStr(wString As String)
  'This sub writes an string to the current file position
  '------------------------------------------------------
  'What it writes:
  '1. The length of the string as an byte (max length 255)
  '2. The string itself (wString)
  
  Put Free, lngPosition, CByte(Len(wString)): lngPosition = lngPosition + 1
  Put Free, lngPosition, wString: lngPosition = lngPosition + Len(wString)

End Sub

Public Sub WriteStrFixed(wString As String)
  'This sub writes an fixed size string to the current
  'file position. The length can be as long as you wish.
  '------------------------------------------------------
  
  Put Free, lngPosition, wString: lngPosition = lngPosition + Len(wString)

End Sub

Public Sub WriteByte(wByte As Byte)
  'This sub writes an byte to the current file position
  '----------------------------------------------------
  'What it writes:
  '1. The byte (wByte)
  
  Put Free, lngPosition, wByte: lngPosition = lngPosition + 1

End Sub

Public Sub WriteLong(wLong As Long)
  'This sub writes an long to the current file position
  '----------------------------------------------------
  'The number is first converted to 4 bytes using Long2Byte
  'What it writes:
  '1. The Long converted to 4 bytes (as an string for easy writing)
  
  Put Free, lngPosition, Long2ByteStr(wLong)
  lngPosition = lngPosition + 4

End Sub

Public Property Get Position() As Long
  Position = lngPosition
End Property

Public Property Let Position(ByVal vData As Long)
  lngPosition = vData
  If lngPosition < 1 Then lngPosition = 1 'The position cannot be lower than 1
End Property

Public Property Let Filename(ByVal vData As String)
  strFilename = vData
End Property

Public Property Get Filename() As String
  Filename = strFilename
End Property

Private Function Long2ByteStr(ByVal val As Long) As String
  'This sub converts the an Long to 4 bytes, returned as an string
  Dim x As Integer

  ReDim byteArray(3) As Byte
  For x = 0 To 3
    byteArray(x) = val \ (2 ^ ((3 - x) * 7)) And 127
  Next x
  
  Long2ByteStr = Chr(byteArray(0)) & Chr(byteArray(1)) & Chr(byteArray(2)) & Chr(byteArray(3))

End Function

Private Function Byte2Long(ByRef byteArray() As Byte) As Long
  'This function converts 4 bytes in byteArray() to an long and retuns it.
  Dim x As Integer

  Byte2Long = 0
  For x = 0 To 3
    Byte2Long = Byte2Long + (byteArray(x) * (2 ^ ((3 - x) * 7)))
  Next x

End Function

Private Sub Class_Initialize()
  'Setup file position
  lngPosition = 1
End Sub

Public Sub OpenFile()
  'This initializes the file
  Free = FreeFile
  Open strFilename For Binary As Free
End Sub

Private Sub Class_Terminate()
  'This closes the file
  Close Free
End Sub
