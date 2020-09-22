Attribute VB_Name = "mBrowseFolder"
'================================================
' Module:        mBrowseFolter.bas
' Author:        Original code: Steve Anderson
' Dependencies:  -
' Last revision: 2003.11.02
'================================================

Option Explicit

'-- API:

Private Const BIF_STATUSTEXT        As Long = &H4&
Private Const BIF_RETURNONLYFSDIRS  As Long = 1
Private Const BIF_DONTGOBELOWDOMAIN As Long = 2
Private Const MAX_PATH              As Long = 260

Private Const WM_USER               As Long = &H400
Private Const BFFM_INITIALIZED      As Long = 1
Private Const BFFM_SELCHANGED       As Long = 2
Private Const BFFM_SETSTATUSTEXT    As Long = (WM_USER + 100)
Private Const BFFM_SETSELECTION     As Long = (WM_USER + 102)

Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

Private Type BrowseInfo
    hwndOwner      As Long
    pIDLRoot       As Long
    pszDisplayName As Long
    lpszTitle      As Long
    ulFlags        As Long
    lpfnCallback   As Long
    lParam         As Long
    iImage         As Long
End Type

'-- Private Variables:
Private m_CurrentFolder As String

'//

Public Function BrowseFolder(OwnerForm As Form, ByVal Tittle As String, Optional ByVal InitFolder As String) As String

  Dim lpIDList    As Long
  Dim sBuffer     As String
  Dim tBrowseInfo As BrowseInfo
  
    If (Not Len(InitFolder) > 1) Then
        If (Not Len(m_CurrentFolder) > 1) Then
            m_CurrentFolder = App.Path & vbNullChar
        End If
      Else
        m_CurrentFolder = InitFolder & vbNullChar
    End If
  
    With tBrowseInfo
        .hwndOwner = OwnerForm.hWnd
        .lpszTitle = lstrcat(Tittle, vbNullString)
        .ulFlags = BIF_RETURNONLYFSDIRS Or BIF_DONTGOBELOWDOMAIN Or BIF_STATUSTEXT
        .lpfnCallback = GetAddressOfFunction(AddressOf BrowseCallbackProc)
    End With

    lpIDList = SHBrowseForFolder(tBrowseInfo)

    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        BrowseFolder = sBuffer
      Else
        BrowseFolder = vbNullString
    End If
End Function

Private Function BrowseCallbackProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lp As Long, ByVal pData As Long) As Long

  Dim lpIDList As Long
  Dim sBuffer  As String
  
    Select Case uMsg
    
        Case BFFM_INITIALIZED
            '-- Go to initial folder
            Call SendMessage(hWnd, BFFM_SETSELECTION, 1, m_CurrentFolder)
      
        Case BFFM_SELCHANGED
            '-- Make the status text show the selected folder
            sBuffer = Space(MAX_PATH)
            lpIDList = SHGetPathFromIDList(lp, sBuffer)
            If (lpIDList) Then
                Call SendMessage(hWnd, BFFM_SETSTATUSTEXT, 0, sBuffer)
                m_CurrentFolder = sBuffer
            End If
    End Select
    
    BrowseCallbackProc = 0
End Function

Private Function GetAddressOfFunction(Address As Long) As Long
    GetAddressOfFunction = Address
End Function
