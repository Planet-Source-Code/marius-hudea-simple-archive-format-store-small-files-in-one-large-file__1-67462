VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4590
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List 
      Height          =   2595
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim a As clsArchive
Set a = New clsArchive
Dim ret As Long
Dim rets As String
Dim i As Long
Me.Show
List.Clear
List.AddItem "Trying to open existing archive C:\archive.bin"
ret = a.OpenArchive("C:\archive.bin", enOPEN_EXISTING)
If ret = -1 Then
 List.AddItem "=> Error opening archive.bin, maybe it does not exist."
 List.AddItem "=> Trying to create a new archive.bin..."
 ret = a.OpenArchive("C:\archive.bin", enCREATE_NEW)
 If ret = -1 Then
  List.AddItem "  => Failed. Can not open or create archive.bin"
  Exit Sub
 End If
End If
List.AddItem "=> Success. Archive.bin open and ready for use."
ret = a.FileCount()
List.AddItem "Files found in the archive: " & CStr(ret)
If ret <> 0 Then
 For i = 1 To ret
  List.AddItem "File " & CStr(i) & ": " & a.GetFileName(i) & " (" & CStr(a.GetFileSize(i)) & " Bytes)"
 Next i
End If
List.AddItem "Adding a string under the filename ""Test1.txt"""
a.AddFromString "Test1.txt", "This is a text of the archive class."
List.AddItem "Added, trying to read from archive the added file..."
rets = a.ReadToString("Test1.txt")
List.AddItem "Result: " & rets
List.AddItem "Extracting from archive and saving to C:\Test1.txt..."
a.ReadToFile "Test1.txt", "c:\Test1.txt", True
List.AddItem "File should now be saved as c:\Test1.txt"
List.AddItem "Closing archive..."
a.CloseArchive
List.AddItem "Demo done."
End Sub
