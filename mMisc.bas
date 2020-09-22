Attribute VB_Name = "mMisc"
'================================================
' Module:        mMisc.bas
' Author:        -
' Dependencies:  -
' Last revision: 2003.11.02
'================================================

Option Explicit

'-- API:

Private Const MAX_PATH As Long = 260

Private Type FILETIME
    dwLowDateTime  As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime   As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime  As FILETIME
    nFileSizeHigh    As Long
    nFileSizeLow     As Long
    dwReserved_      As Long
    dwReserved1      As Long
    cFileName        As String * MAX_PATH
    cAlternate       As String * 14
End Type

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function PathCompactPath Lib "shlwapi" Alias "PathCompactPathA" (ByVal hDC As Long, ByVal pszPath As String, ByVal dx As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'//

Public Function FileFound(strFileName As String) As Boolean
'-- By:
'   Eric Russell
'   Created: 1998-03-19
'   Revised: 1998-03-20

  Dim lpFindFileData As WIN32_FIND_DATA
  Dim hFindFirst     As Long

    hFindFirst = FindFirstFile(strFileName, lpFindFileData)

    If (hFindFirst > 0) Then
        FindClose hFindFirst
        FileFound = True
      Else
        FileFound = False
    End If
End Function

Public Function CompactPath(ByVal hDC As Long, ByVal FullPath As String, ByVal Width As Long) As String
'-- From:
'   KPD-Team 2000
'   URL: http://www.allapi.net/
'   E-Mail: KPDTeam@Allapi.net

  Dim ZeroPos As Long

    '-- Compact
    Call PathCompactPath(hDC, FullPath, Width)

    '-- Remove all trailing Chr$(0)'s
    ZeroPos = InStr(1, FullPath, Chr$(0))
    If (ZeroPos > 0) Then
        CompactPath = Left$(FullPath, ZeroPos - 1)
      Else
        CompactPath = FullPath
    End If
End Function

Public Sub RemoveButtonBorderEnhance(Button As CommandButton)
    Call SendMessage(Button.hWnd, &HF4&, &H0&, 0&)
End Sub

