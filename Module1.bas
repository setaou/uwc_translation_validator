Attribute VB_Name = "mMisc"
Option Explicit

'-------- Liste des fichiers
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Public Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long

Public Const MAX_PATH = 260
Public Const INVALID_HANDLE_VALUE = -1
Public Const ERROR_NO_MORE_FILES = 18
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_DEVICE = &H40

Public Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type


Public Function GetDirectory(ByVal path As String) As String
    GetDirectory = Mid(path, 1, InStrRev(path, "\"))
End Function

Public Function GetFileName(ByVal path As String) As String
    GetFileName = Mid(path, InStrRev(path, "\") + 1, Len(path) - InStrRev(path, "\"))
End Function

Public Function DirectoryExists(ByVal Doss As String) As Boolean
    Dim attrib As Long
    If Right(Doss, 1) <> "\" Then Doss = Doss + "\"
    
    attrib = GetFileAttributes(Doss)
    DirectoryExists = Not (attrib = INVALID_HANDLE_VALUE)
End Function

Public Function FileExists(ByVal file As String) As Boolean
    Dim attrib As Long
    
    attrib = GetFileAttributes(file)
    FileExists = (Not (attrib = INVALID_HANDLE_VALUE)) And (Not DirectoryExists(file))
End Function




Public Function CountInStr(countwhat As String, inwhat As String) As Integer
    Dim total As Integer
    Dim pos As Integer
    pos = 0
    pos = InStr(pos + 1, inwhat, countwhat)
    While pos > 0
        total = total + 1
        pos = InStr(pos + 1, inwhat, countwhat)
    Wend
    CountInStr = total
End Function
