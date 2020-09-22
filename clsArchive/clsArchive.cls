VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsArchive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
' Very Basic Archive System Class
'
' (c)2006 Marius Hudea
'
' Permission to use in application granted as long as
' credit is given to the original author. (leaving this text here
' is enough).
'
' This is a class designed to open and save a very basic format
' that allows you to store several files in a single large file.
'
'
' The file list is at the end of the file, it has a fixed size,
' therefore it is easy to parse. If a file needs to be added,
' the file pointer just has to be placed where the file index
' starts, write the file over the index and write again the nex
' file index at the end.
'
' The index format is very simple:
'
' <number of files> x <file data> for each file
' <number of files> x 100 bytes:
'  80 bytes : file name with spaces in front
'  10 bytes : start position in archive
'  10 bytes : file size
' <10 bytes> the number of files in the archive
'
' The archive format was designed to be compatible with lots
'of systems, withouth thinking of byte order. That's why all
'numbers are converted to string and padded with spaces to a
'fixed length. It was not designed for optimum byte usage.
' Because of this, there is only one real limitation:
'the file size for each size can not exceed 9.999.999.999 bytes
'(actually 2^32-1 bytes-the max for Long data type)
' The rest of the limitations are mostly because of not using
'the high byte of GetFileSize API function.
' This means that the max file size of the archive must be less than
'2^32-1 Bytes, the file index must start before that value
'and the number of files can be maximum that value.
' It can easily be designed to work for files larger than ~ 2 GB
'and store more than 2^32-1 files and have sizes up to terrabytes
'if you simply replace Longs with Doubles and use the high bytes of
'GetFileSize function.
' The file format does not implement a "signature" that will tell
'the "parser" it is really a valid archive, it could be easily
'implemented but I didn't need a "signature".
' For a standardized and easy to parse format that is currently
'in use, see TAR (though, the whole file has to be parsed in order
'to get the file list in that format).


Const FILE_BEGIN = 0
Const FILE_END = 2
Const CREATE_NEW = 1
Const CREATE_ALWAYS = 2
Const OPEN_EXISTING = 3
Const OPEN_ALWAYS = 4

Const FILE_SHARE_READ = &H1
Const FILE_SHARE_WRITE = &H2
Const GENERIC_READ = &H80000000
Const GENERIC_WRITE = &H40000000

Const INVALID_HANDLE_VALUE = -1

Public Enum enOpenMode
 enCREATE_NEW = 1
 enCREATE_ALWAYS = 2
 enOPEN_EXISTING = 3
 enOPEN_ALWAYS = 4
End Enum

Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function apiFileSize Lib "kernel32" Alias "GetFileSize" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long

Private Type typeFileEntry
 FileName As String
 FilePos As Long
 FileSize As Long
End Type

Private FileList() As typeFileEntry     ' file entries in storage
Private FileListCnt As Long             ' how many files are already stored
Private FileListMax As Long             ' how much can array handle
Private FileHandle As Long              ' handle for storage file
Private FilePosition As Long            ' Write Position in file

Private FileLoaded As Boolean           ' is file loaded ?
Private FileModified As Boolean         ' if true, must rewrite index

Private Sub Class_Initialize()
ReDim FileList(1 To 128)
FileListMax = 128
FileListCnt = 0
FileLoaded = False
FileModified = False
End Sub

Public Function OpenArchive(FileName As String, OpenMode As enOpenMode) As Long
Dim fslo As Long
Dim fshi As Long
Dim s As String
Dim bRead As Long
Dim i As Long

If FileLoaded Then CloseArchive     ' close previous archive if needed
FileHandle = CreateFile(FileName, GENERIC_READ Or GENERIC_WRITE, _
                        FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OpenMode, 0, 0)
If FileHandle = INVALID_HANDLE_VALUE Then
   OpenArchive = INVALID_HANDLE_VALUE
   Exit Function
End If
fslo = apiFileSize(FileHandle, fshi)            ' we're not using fshi=>files can be up to 2GB
' Archive must have at least 10 bytes, the last 10 bytes contain the number of files inside
'the archive. If smaller than 10 bytes, assume it's corrupted or not an archive and return
'as if the file can not be opened.
If fslo < 10 And fslo > 0 Then
  CloseHandle FileHandle
  OpenArchive = INVALID_HANDLE_VALUE
  Exit Function
End If
'  Archive size can be 0, if just create using CREATE_NEW, CREATE_ALWAYS or OPEN_ALWAYS
'  If Archive size is 0 bytes, we don't have to do anything anymore, except initializing some
' variables
If fslo = 0 Then
 FileListCnt = 0
 FilePosition = 0
 FileLoaded = True
 FileModified = True     ' it's not really modified but this way we force writing index
 OpenArchive = FileHandle
 Exit Function
End If
' Here, we know there are at least 10 bytes in the file, move 10 bytes backwards from the
' end of the archive and read those bytes.
SetFilePointer FileHandle, fslo - 10, 0, FILE_BEGIN
' create a "buffer"  for the data to be read and then read it.
' The value is padded at the front with spaces, so if any space in front it must be removed
s = Space(10): ReadFile FileHandle, ByVal s, 10, bRead, 0&
If bRead <> 10 Then     ' error reading from file
  CloseHandle FileHandle: OpenArchive = INVALID_HANDLE_VALUE: Exit Function
End If
s = Trim(s)
' Value must be numeric, it's simply a counter for the number of files inside the archive
FileListCnt = ConvertToLong(s)
' There may have been an error converting the string to long, if so exit,data may be corrupted
If FileListCnt = -1 Then
  FileListCnt = 0: CloseHandle FileHandle: OpenArchive = INVALID_HANDLE_VALUE: Exit Function
End If
' Returned value is 0, it means archive has no files inside
If FileListCnt = 0 Then
 FileListCnt = 0
 FilePosition = 0
 FileLoaded = True
 FileModified = True
 OpenArchive = FileHandle
 Exit Function
End If
' Archive has at least 1 file inside it, prepare array by increasing it to
' the number of entries of the archive plus some extra records, then read the content
ReDim FileList(1 To FileListCnt + 128)
' move to where the index starts
SetFilePointer FileHandle, fslo - 10 - FileListCnt * 100, 0, FILE_BEGIN
s = Space(100)
For i = 1 To FileListCnt
 ReadFile FileHandle, ByVal s, 100, bRead, 0&
 If bRead <> 100 Then     ' error reading from file
  CloseHandle FileHandle: OpenArchive = INVALID_HANDLE_VALUE: Exit Function
 End If
 With FileList(i)
  .FileName = Trim(Left(s, 80))
  .FilePos = ConvertToLong(Mid(s, 81, 10))
  .FileSize = ConvertToLong(Mid(s, 91, 10))
 End With
Next i
FilePosition = fslo - 10 - FileListCnt * 100
FileLoaded = True
FileModified = False
End Function

Private Sub WriteIndex()
Dim i As Long
Dim s As String
Dim bWrote As Long
Debug.Print "writeindex called"
SetFilePointer FileHandle, FilePosition, 0, FILE_BEGIN
If FileListCnt <> 0 Then
  For i = 1 To FileListCnt
   s = PadWithSpaces(FileList(i).FileName, 80)
   s = s & PadWithSpaces(CStr(FileList(i).FilePos))
   s = s & PadWithSpaces(CStr(FileList(i).FileSize))
   WriteFile FileHandle, ByVal s, 100, bWrote, 0&
  Next i
End If
s = PadWithSpaces(CStr(FileListCnt))
WriteFile FileHandle, ByVal s, 10, bWrote, 0&
FileModified = False
End Sub

Public Function CloseArchive() As Long
If FileLoaded Then
 If FileModified Then
    WriteIndex
 End If
 CloseHandle FileHandle
End If
FileLoaded = False          ' archive not open anymore
FileModified = False        ' naturally, not modified if not open
FileListCnt = 0             ' no files as archive is closed
FilePosition = 0            ' reset variables to default
FileHandle = 0
End Function

Public Sub AddFromString(FileName As String, Content As String)
Dim lenC As Long
Dim bWrote As Long
FileListCnt = FileListCnt + 1
If FileListCnt > FileListMax Then
 FileListMax = FileListMax + 128
 ReDim Preserve FileList(1 To FileListMax)
End If
lenC = Len(Content)
FileList(FileListCnt).FileName = FileName
FileList(FileListCnt).FilePos = FilePosition
FileList(FileListCnt).FileSize = lenC
SetFilePointer FileHandle, FilePosition, 0, FILE_BEGIN
If lenC > 0 Then
WriteFile FileHandle, ByVal Content, lenC, bWrote, 0&
End If
FilePosition = FilePosition + lenC
FileModified = True
End Sub

Public Sub AddFromFile(FileName As String, SourceFilename As String)
Dim lenC As Long
Dim bRead As Long
Dim bWrote As Long
Dim hSrc As Long
Dim flen As Long
Dim Buffer(0 To 4095) As Byte
hSrc = CreateFile(SourceFilename, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, _
                  0&, OPEN_EXISTING, 0, 0)
If hSrc = INVALID_HANDLE_VALUE Then
   Exit Sub ' oopsie, can not open file, just quit
End If
FileListCnt = FileListCnt + 1
If FileListCnt > FileListMax Then
 FileListMax = FileListMax + 128
 ReDim Preserve FileList(1 To FileListMax)
End If

SetFilePointer FileHandle, FilePosition, 0, FILE_BEGIN
bRead = -1
flen = 0
While bRead <> 0
 ReadFile hSrc, Buffer(0), 4096, bRead, 0&
 If bRead > 0 Then
  WriteFile FileHandle, Buffer(0), bRead, bWrote, 0&
  flen = flen + bRead
 End If
Wend
CloseHandle hSrc
FileList(FileListCnt).FileName = FileName
FileList(FileListCnt).FilePos = FilePosition
FileList(FileListCnt).FileSize = flen
FilePosition = FilePosition + flen
FileModified = True
End Sub

Public Function ReadToString(FileName As String) As String
Dim i As Long
Dim s As String
Dim bRead As Long
If FileListCnt = 0 Then         ' there are no files, what am i supposed to do.. return ""
ReadToString = ""
Exit Function
End If
For i = 1 To FileListCnt
 If FileList(i).FileName = FileName Then
   ' create "buffer" for data
   s = Space(FileList(i).FileSize)
   ' move inside the archive where file starts
   SetFilePointer FileHandle, FileList(i).FilePos, 0, FILE_BEGIN
   ' read it in the string variable
   ReadFile FileHandle, ByVal s, FileList(i).FileSize, bRead, 0&
   ReadToString = s
   ' read it, exit from function because no point searching for other files if already found
   Exit Function
 End If
Next i
End Function

Public Function ReadToFile(FileName As String, dstFilename As String, Optional OverwriteExisting As Boolean = False)
Dim hDst As Long
Dim i As Long
Dim j As Long
Dim chunks As Long
Dim lastchunk As Long
Dim OpenMode As Long
Dim Buffer(0 To 4095) As Byte
Dim bRead As Long
Dim bWrote As Long

OpenMode = IIf(OverwriteExisting = True, CREATE_ALWAYS, CREATE_NEW)
' find the file in the archive
If FileListCnt = 0 Then
  ReadToFile = -1: Exit Function
End If
For i = 1 To FileListCnt
 If FileList(i).FileName = FileName Then
  ' try to open the destination file
  hDst = CreateFile(dstFilename, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OpenMode, 0, 0)
  ' error creating file for output, maybe opened in create_new mode and file already exists
  If hDst = INVALID_HANDLE_VALUE Then
   ReadToFile = -1
   Exit Function
  End If
  ' we try to read 4KB chunks of data from the archive and write them to destination
  ' that's why we compute how many chunks of 4 KB exist and determine the last chunk size
  chunks = FileList(i).FileSize \ 4096
  lastchunk = FileList(i).FileSize Mod 4096
  ' move pointer to start of file in archive
  SetFilePointer FileHandle, FileList(i).FilePos, 0, FILE_BEGIN
  If chunks <> 0 Then
   For j = 1 To chunks
    ' an additional test may be needed here, to use the bread variable to count
    ' exactly how many bytes readfile was able to return.
    ReadFile FileHandle, Buffer(0), 4096, bRead, 0&
    WriteFile hDst, Buffer(0), bRead, bWrote, 0&
   Next j
  End If
  If lastchunk <> 0 Then
   ReadFile FileHandle, Buffer(0), lastchunk, bRead, 0&
   WriteFile hDst, Buffer(0), bRead, bWrote, 0&
  End If
  'close destination file
  CloseHandle hDst
  Exit Function
 End If
Next i
End Function

Private Function ConvertToLong(s As String) As Long
' Value may not be numeric, check, return -1 if not numeric
If (Not IsNumeric(s)) Then
 ConvertToLong = -1: Exit Function
End If
' Value may be numeric but not in the acceptable range (min is 0, max is 2.147.483.647)
' Using CDbl because if s had, for example, 9999999999, the function would have crashed
If CDbl(s) < 0 Or CDbl(s) > &H7FFFFFFF Then
 ConvertToLong = -1: Exit Function
End If
ConvertToLong = CLng(s)
End Function

Private Function PadWithSpaces(s As String, Optional length As Long = 10) As String
Dim t As String
t = s
If Len(t) < length Then t = Space(length - Len(t)) & t
PadWithSpaces = t
End Function

Private Sub Class_Terminate()
If FileLoaded Then CloseArchive
End Sub

Private Sub DebugMe()
Debug.Print "FileLoaded: ", FileLoaded
Debug.Print "FileModified:", FileModified
Debug.Print "FileListCnt:", FileListCnt
Debug.Print "FileListmax:", FileListMax
End Sub

Public Function FileCount() As Long
FileCount = FileListCnt
End Function

Public Function GetFileName(Index As Long) As String
 If Index < 0 Or Index > FileListCnt Then
   GetFileName = ""
   Exit Function
 End If
 GetFileName = FileList(Index).FileName
End Function

Public Function GetFileSize(Index As Long) As Long
 If Index < 0 Or Index > FileListCnt Then
   GetFileSize = 0
   Exit Function
 End If
 GetFileSize = FileList(Index).FileSize
End Function
