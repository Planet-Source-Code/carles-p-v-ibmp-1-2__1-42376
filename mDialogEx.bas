Attribute VB_Name = "mDialogEx"
'================================================
' Module:        mDialogEx.bas
' Author:        (*)
' Dependencies:  -
' Last revision: 2003.11.02
'================================================
'
' (*) Original code from
'     http://www.pscode.com/vb/scripts/ShowCode.asp?txtCodeId=41320&lngWId=1
'     Digital image processing, by Ark

Option Explicit
Option Compare Text

'-- API:

'-- Open & Save Dialog
Private Type OPENFILENAME
    lStructSize       As Long
    hwndOwner         As Long
    hInstance         As Long
    lpstrFilter       As String
    lpstrCustomFilter As String
    nMaxCustFilter    As Long
    nFilterIndex      As Long
    lpstrFile         As String
    nMaxFile          As Long
    lpstrFileTitle    As String
    nMaxFileTitle     As Long
    lpstrInitialDir   As String
    lpstrTitle        As String
    Flags             As Long
    nFileOffset       As Integer
    nFileExtension    As Integer
    lpstrDefExt       As String
    lCustData         As Long
    lpfnHook          As Long
    lpTemplateName    As String
End Type

'-- Hook and notification support
Private Type NMHDR
    hwndFrom As Long
    IDFrom   As Long
    Code     As Long
End Type

Private Type OFNOTIFYshort
    HDR   As NMHDR
    lpOFN As Long
End Type

Private Type LV_ITEM
    Mask       As Long
    iItem      As Long
    iSubItem   As Long
    State      As Long
    StateMask  As Long
    pszText    As String
    cchTextMax As Long
    iImage     As Long
    lParam     As Long
    iIndent    As Long
End Type

Private Type RECT2
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type
 
Private Const LVM_FIRST = &H1000
Private Const LVM_GETITEMTEXT = LVM_FIRST + 45
Private Const LVM_GETNEXTITEM = LVM_FIRST + 12
Private Const LVNI_FOCUSED = &H1
Private Const LVNI_SELECTED = &H2

Private Const ID_OPEN          As Long = &H1   ' Open or Save button
Private Const ID_CANCEL        As Long = &H2   ' Cancel Button
Private Const ID_HELP          As Long = &H40E ' Help Button
Private Const ID_READONLY      As Long = &H410 ' Read-only check box
Private Const ID_FILETYPELABEL As Long = &H441 ' FileType label
Private Const ID_FILELABEL     As Long = &H442 ' FileName label
Private Const ID_FOLDERLABEL   As Long = &H443 ' Folder label
Private Const ID_LIST          As Long = &H461 ' Parent of file list
Private Const ID_FORMAT        As Long = &H470 ' FileType combo box
Private Const ID_FOLDER        As Long = &H471 ' Folder combo box
Private Const ID_FILETEXT      As Long = &H480 ' FileName text box

Private Const OFN_HELPBUTTON      As Long = &H10
Private Const OFN_HIDEREADONLY    As Long = &H4
Private Const OFN_ENABLEHOOK      As Long = &H20
Private Const OFN_ENABLETEMPLATE  As Long = &H40
Private Const OFN_EXPLORER        As Long = &H80000
Private Const OFN_OVERWRITEPROMPT As Long = &H2
Private Const OFN_PATHMUSTEXIST   As Long = &H800
Private Const OFN_FILEMUSTEXISTS  As Long = &H1000
Private Const OFN_ENABLESIZING    As Long = &H800000

Private Const OFN_OPENFLAGS As Long = &H881024
Private Const OFN_SAVEFLAGS As Long = &H880026

Private Const WM_INITDIALOG As Long = &H110
Private Const WM_COMMAND    As Long = &H111

Private Const WM_DESTROY As Long = &H2
Private Const WM_NOTIFY  As Long = &H4E
Private Const WM_SETICON As Long = &H80

Private Const WM_USER            As Long = &H400
Private Const CDM_FIRST          As Long = (WM_USER + 100)
Private Const CDM_GETSPEC        As Long = (CDM_FIRST + &H0)
Private Const CDM_GETFILEPATH    As Long = (CDM_FIRST + &H1)
Private Const CDM_GETFOLDERPATH  As Long = (CDM_FIRST + &H2)
Private Const CDM_SETCONTROLTEXT As Long = (CDM_FIRST + &H4)
Private Const CDM_HIDECONTROL    As Long = (CDM_FIRST + &H5)
Private Const CDM_SETDEFEXT      As Long = (CDM_FIRST + &H6)
Private Const CB_GETCURSEL       As Long = &H147

Private Const CDN_FIRST        As Long = -601&
Private Const CDN_INITDONE     As Long = (CDN_FIRST)
Private Const CDN_SELCHANGE    As Long = (CDN_FIRST - &H1)
Private Const CDN_FOLDERCHANGE As Long = (CDN_FIRST - &H2)
Private Const CDN_HELP         As Long = (CDN_FIRST - &H4)
Private Const CDN_FILEOK       As Long = (CDN_FIRST - &H5)
Private Const CDN_TYPECHANGE   As Long = (CDN_FIRST - &H6)

Private Const GW_HWNDFIRST As Long = 0
Private Const GW_HWNDNEXT  As Long = 2
Private Const GW_CHILD     As Long = 5

Private Const MAX_PATH As Long = 260

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT2) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT2) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function GetDlgItem Lib "user32" (ByVal hDlg As Long, ByVal nIDDlgItem As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Private Declare Function GetOpenFileName Lib "comdlg32" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

'-- Private Variables:
Private m_hDlg       As Long
Private m_fExtraForm As Form
Private m_hOldParent As Long
Private m_bOpen      As Boolean
Private sDlgFilter   As String
Private sCurExt      As String

'//

Public Function GetFileName(Optional sPath As String, Optional sFilter As String, Optional nFltIndex As Long, Optional sTitle As String, Optional bOpen As Boolean = True, Optional fExtra As Form) As String
   
 Dim OFN  As OPENFILENAME
 Dim lRet As Long
 Dim lIdx As Long
 Dim sExt As String
 
    m_hOldParent = 0
    m_hDlg = 0
    m_bOpen = bOpen
   
    Set m_fExtraForm = Nothing
   
    For lIdx = 1 To Len(sFilter)
        If (Mid$(sFilter, lIdx, 1) = "|") Then
            Mid$(sFilter, lIdx, 1) = vbNullChar
        End If
    Next lIdx
    
    If (Len(sFilter) < MAX_PATH) Then
        sFilter = sFilter & String$(MAX_PATH - Len(sFilter), 0)
      Else
        sFilter = sFilter & Chr(0) & Chr(0)
    End If
    sDlgFilter = sFilter
    
    If (sPath <> vbNullString And (nFltIndex = 0)) Then
        nFltIndex = GetFilterIndex(sPath)
    End If
        
    With OFN
        .hwndOwner = fMain.hWnd
        .lStructSize = Len(OFN)
        .lpstrTitle = sTitle
        .lpstrFile = sPath & String(MAX_PATH - Len(sPath), 0)
        .lpstrFilter = sFilter
        .lpfnHook = lHookAddress(AddressOf DialogHookProcess)
        .hInstance = App.hInstance
        .nFilterIndex = nFltIndex
        .nMaxFile = MAX_PATH
    End With
   
    Set m_fExtraForm = fExtra
    
    If (Not m_fExtraForm Is Nothing) Then
        m_fExtraForm.fraJPEGOptions.Visible = Not m_bOpen
    End If
    
    If (m_bOpen) Then
        OFN.Flags = OFN.Flags Or OFN_OPENFLAGS
        lRet = GetOpenFileName(OFN)
      Else
        OFN.Flags = OFN.Flags Or OFN_SAVEFLAGS
        lRet = GetSaveFileName(OFN)
    End If
    
    If (lRet) Then
        GetFileName = TrimNull(OFN.lpstrFile)
        If (OFN.nFileExtension = 0) And Len(sCurExt) > 2 Then
            GetFileName = GetFileName & Mid$(sCurExt, 2)
        End If
    End If
End Function

Public Function lHookAddress(lPtr As Long) As Long
    lHookAddress = lPtr
End Function

Public Function DialogHookProcess(ByVal hDlg As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   
  Dim tNMH  As NMHDR
  Dim sPath As String
  Dim sExt  As String
  Dim nPos  As Long
  
    Select Case wMsg
    
        Case WM_NOTIFY
        
            CopyMemory tNMH, ByVal lParam, Len(tNMH)
        
            Select Case tNMH.Code
              
                Case CDN_FOLDERCHANGE
                    
                    Call SendMessage(m_hDlg, CDM_SETCONTROLTEXT, ID_FILETEXT, ByVal vbNullString)
                    With m_fExtraForm.Preview
                        Call .DIB.Destroy
                        Call .Resize
                    End With
                    m_fExtraForm.lblSize = "Size:"
                
                Case CDN_SELCHANGE
                
                    sPath = GetSelItem
                    If (sPath <> vbNullString) Then
                        Call SendMessage(m_hDlg, CDM_SETCONTROLTEXT, ID_FILETEXT, ByVal sPath)
                    End If
                    If (m_fExtraForm.chkPreview) Then
                        Call PreviewPicture
                    End If
                
                Case CDN_TYPECHANGE
              
                    If (Not m_bOpen) Then
                    
                        sPath = String(MAX_PATH, 0)
                        Call SendMessage(m_hDlg, CDM_GETSPEC, MAX_PATH, ByVal sPath)
                        sPath = TrimNull(sPath)
                        
                        If (Len(sPath) > 4) Then
                            sExt = Right$(sPath, 5)
                            nPos = InStr(1, sExt, ".")
                            If (nPos) Then
                                sPath = Left$(sPath, Len(sPath) - 6 + nPos)
                            End If
                        End If
                        
                        sCurExt = GetExtension()
                        If (Len(sCurExt) > 2) Then
                            Call SendMessage(m_hDlg, CDM_SETDEFEXT, 0, ByVal Mid$(sCurExt, 3))
                        End If
                        Call SendMessage(m_hDlg, CDM_SETCONTROLTEXT, ID_FILETEXT, ByVal sPath)
                        
                        m_fExtraForm.fraJPEGOptions.Visible = (sCurExt = "*.jpg")
                    End If
                
              Case CDN_INITDONE
                
                m_hDlg = GetParent(hDlg)
                Call ModifyDialog(True)
                
            End Select
               
        Case WM_DESTROY
        
            If (m_hOldParent) Then
                m_fExtraForm.Visible = False
                Call SetParent(m_fExtraForm.hWnd, m_hOldParent)
                fMain.FileExt = GetExtension()
            End If
   End Select
End Function
 
'========================================================================================
' Private
'========================================================================================
 
Private Sub ModifyDialog(Optional ByVal bShowExtra As Boolean)

  Dim DlgRct     As RECT2
  Dim DlgRctClnt As RECT2
  Dim ControlRct As RECT2
  
  Dim Pt   As POINTAPI
  Dim xOff As Long
  Dim yOff As Long
  Dim W    As Long
  Dim H    As Long
  
  Dim lhWnd  As Long
  Dim sClass As String
  Dim lRet   As Long
  Dim lLVTop As Long
  Const yINC As Long = 150
   
    Call GetWindowRect(m_hDlg, DlgRct)
    Call GetClientRect(m_hDlg, DlgRctClnt)
    
    '-- Resize Dialog window
    With DlgRct
        W = (.x2 + IIf(bShowExtra, m_fExtraForm.ScaleWidth, 0)) - .x1
        H = (.y2 + yINC) - .y1
    End With
    With Screen
        xOff = (.Width / .TwipsPerPixelX - W) \ 2
        yOff = (.Height / .TwipsPerPixelX - H) \ 2
    End With
    
    '-- Locate extra Form (Preview + options)
    If (bShowExtra) Then
        If (Not m_fExtraForm Is Nothing) Then
            Pt.x = DlgRctClnt.x2 * Screen.TwipsPerPixelX
            Pt.y = DlgRctClnt.y1 * Screen.TwipsPerPixelY
            Call m_fExtraForm.Move(Pt.x, Pt.y)
            m_hOldParent = SetParent(m_fExtraForm.hWnd, m_hDlg)
            m_fExtraForm.Visible = True
            m_fExtraForm.fraJPEGOptions.Visible = (m_bOpen = False And GetExtension() = "*.jpg")
        End If
    End If
    
    '-- 'Fit' controls to Dialog new size
    lhWnd = GetWindow(m_hDlg, GW_CHILD)
    
    Do: sClass = Space$(128)
        lRet = GetClassName(lhWnd, ByVal sClass, 128)
        sClass = Left$(sClass, lRet)
           
        Call GetWindowRect(lhWnd, ControlRct)
        
        With ControlRct
        
            Pt.x = .x1
            Pt.y = .y1
            ScreenToClient m_hDlg, Pt
            
            If (lLVTop = 0 And sClass = "ListBox") Then
                lLVTop = Pt.y
                Call MoveWindow(lhWnd, Pt.x, Pt.y, .x2 - .x1, .y2 - .y1 + yINC, 0)
                m_fExtraForm.fraPreview.Top = Pt.y
                lhWnd = GetWindow(lhWnd, GW_HWNDFIRST)
              Else
                If (lLVTop And Pt.y > lLVTop) Then
                    Call MoveWindow(lhWnd, Pt.x, Pt.y + yINC, .x2 - .x1, .y2 - .y1, 0)
                End If
            End If
        End With
        lhWnd = GetWindow(lhWnd, GW_HWNDNEXT)
    Loop While (lhWnd <> 0)
    
    '-- Show
    Call MoveWindow(m_hDlg, xOff, yOff, W, H, 0)
End Sub

Private Sub PreviewPicture()
   
  Dim sPath As String
  Dim sExt  As String
  
  Dim DummyPal    As cDIBPal
  Dim DummyDither As cDIBDither
   
    sPath = String(MAX_PATH, 0)
    Call SendMessage(m_hDlg, CDM_GETFILEPATH, MAX_PATH, ByVal sPath)
    sPath = TrimNull(sPath)
    
    If (sPath <> vbNullString) Then
        If (FileFound(sPath)) Then
            If (GetAttr(sPath) And vbDirectory) <> vbDirectory Then
            
                sExt = Right$(sPath, 5)
                If (InStr(sExt, ".") = 0 And Len(sCurExt) > 2) Then
                    sPath = sPath & Mid$(sCurExt, 2)
                End If
                
                m_fExtraForm.lblWait.Visible = True
                
                With m_fExtraForm.Preview
                    '-- Preview image
                    DoEvents
                    Call .DIB.CreateFromStdPicture(LoadPictureEx(sPath), DummyPal, DummyDither)
                    Call .Resize
                    '-- Show dimensions
                    m_fExtraForm.lblSize = "Size: " & .DIB.Width & "x" & .DIB.Height
                End With
                
                m_fExtraForm.lblWait.Visible = False
            End If
        End If
    End If
End Sub

Private Function TrimNull(StartStr As String) As String
  
  Dim lPos As Long
  
    lPos = InStr(StartStr, Chr$(0))
    If (lPos) Then
        TrimNull = Left$(StartStr, lPos - 1)
      Else
        TrimNull = StartStr
    End If
End Function

Private Function GetSelItem() As String
  
  Static sOldPath As String
  
  Dim LI          As LV_ITEM
  Dim lRet        As Long
  Dim hFileList   As Long
  Dim sPath       As String
  Dim sNewPath    As String
   
    sNewPath = String(MAX_PATH, 0)
    Call SendMessage(m_hDlg, CDM_GETFILEPATH, MAX_PATH, ByVal sNewPath)
    sNewPath = TrimNull(sNewPath)
    
    If (sNewPath <> sOldPath) Then
        sOldPath = sNewPath
        Exit Function
    End If
    
    hFileList = GetDlgItem(GetDlgItem(m_hDlg, ID_LIST), 1)
    
    If (hFileList <> 0) Then
        lRet = SendMessage(hFileList, LVM_GETNEXTITEM, -1, ByVal LVNI_SELECTED)
        If (lRet <> -1) Then
            LI.cchTextMax = MAX_PATH
            LI.pszText = Space$(MAX_PATH)
            lRet = SendMessage(hFileList, LVM_GETITEMTEXT, lRet, LI)
            If (lRet > 1) Then
                sPath = Left$(LI.pszText, lRet)
            End If
            GetSelItem = sPath
            sOldPath = sPath
        End If
    End If
End Function

Private Function GetExtension() As String

  Dim lIdx    As Long
  Dim nFilter As Long
  Dim nStart  As Long
  Dim hCombo  As Long
  Dim sFilter As String
  Dim sTemp   As String
   
    hCombo = GetDlgItem(m_hDlg, ID_FORMAT)
    nFilter = SendMessage(hCombo, CB_GETCURSEL, 0, ByVal 0&)
    sFilter = sDlgFilter
   
    For lIdx = 1 To nFilter * 2 + 1
        nStart = InStr(1, sFilter, Chr(0))
        If (nStart) Then
            sFilter = Mid$(sFilter, nStart + 1)
          Else
            Exit For
        End If
    Next lIdx
    
    sTemp = TrimNull(sFilter)
    If (Len(sTemp) <> 0) Then
        If (InStr(1, sTemp, ";") = 0) Then
            GetExtension = sTemp
        End If
    End If
End Function

Private Function GetFilterIndex(ByVal sPath As String) As Long

  Dim sExt   As String
  Dim nIdx   As Long
  Dim nStart As Long
  
    sExt = Right$(sPath, 4)
    
    If (Left$(sExt, 1) = ".") Then
        sExt = Mid$(sExt, 2)
    End If
    sExt = "*." & sExt & Chr(0)
    
    nStart = 1
    Do While nStart
        nStart = InStr(nStart + 1, sDlgFilter, Chr(0), vbTextCompare)
        If (Mid$(sDlgFilter, nStart + 1, Len(sExt)) = sExt) Then Exit Do
        nIdx = nIdx + 1
    Loop
    GetFilterIndex = nIdx \ 2 + 1
End Function
