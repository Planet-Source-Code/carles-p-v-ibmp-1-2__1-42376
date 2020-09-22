VERSION 5.00
Begin VB.UserControl ucInfo 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4860
   ClipControls    =   0   'False
   ScaleHeight     =   14
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   324
End
Attribute VB_Name = "ucInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================
' UserControl:   ucInfo.ctl (Filename/Size/Zoom status bar)
' Author:        Carles P.V.
' Dependencies:  -
' Last revision: 2003.11.02
'================================================

Option Explicit

'-- API:

Private Type RECT2
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

Private Const BDR_RAISEDINNER     As Long = &H4
Private Const BDR_SUNKENOUTER     As Long = &H2
Private Const BF_RECT             As Long = &HF
Private Const DFC_SCROLL          As Long = &H3
Private Const DFCS_SCROLLSIZEGRIP As Long = &H8
Private Const COLOR_BTNFACE       As Long = 15
Private Const WM_NCLBUTTONDOWN    As Long = &HA1
Private Const HTBOTTOMRIGHT       As Long = &H11

Private Declare Function SetRect Lib "user32" (lpRect As RECT2, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT2, lpSourceRect As RECT2) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT2, ByVal hBrush As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT2, ByVal wFormat As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT2, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function DrawFrameControl Lib "user32" (ByVal hDC As Long, lpRect As RECT2, ByVal un1 As Long, ByVal un2 As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
                         
'-- Private Variables:
Private m_TextFile        As String
Private m_TextInfo        As String
Private m_TextZoom        As String

Private pBarRect          As RECT2
Private pSizeGripRect     As RECT2
Private pEdgeRect(1 To 3) As RECT2
Private pTextRect(1 To 3) As RECT2

'-- Public Events:
Public Event Resize()

'========================================================================================
' UserControl
'========================================================================================

Private Sub UserControl_Show()
    Call UserControl_Resize
End Sub

Private Sub Refresh()

  Dim i As Long
    
    '-- Erase background
    Call FillRect(hDC, pBarRect, GetSysColorBrush(COLOR_BTNFACE))
    '-- Draw edges
    For i = 1 To 3
        Call DrawEdge(hDC, pEdgeRect(i), IIf(i = 1, BDR_RAISEDINNER, BDR_SUNKENOUTER), BF_RECT)
    Next i
    '-- Draw size grip
    Call DrawFrameControl(hDC, pSizeGripRect, DFC_SCROLL, DFCS_SCROLLSIZEGRIP)
    '-- Draw text
    Call DrawText(hDC, m_TextFile, Len(m_TextFile), pTextRect(1), &H0)
    Call DrawText(hDC, m_TextInfo, Len(m_TextInfo), pTextRect(2), &H0)
    Call DrawText(hDC, m_TextZoom, Len(m_TextZoom), pTextRect(3), &H0)
    
    Call UserControl.Refresh
End Sub

Private Sub UserControl_Resize()
    
  Const INFO_WIDTH As Long = 100
  Const ZOOM_WIDTH As Long = 64
    
  Dim W  As Long
  Dim H  As Long
  Dim SG As Long
  Dim i  As Long
  
    Height = 20 * Screen.TwipsPerPixelY
    W = ScaleWidth
    H = 20
    
    On Error Resume Next
    
    '-- Check parent form window state
    '-- Size Grip (Show/Hide)
    If (Parent.WindowState = vbMaximized) Then
        SG = 0
      Else
        SG = H
    End If
    '-- Set main Rect. and size grip Rect.
    Call SetRect(pBarRect, 0, 0, W, H)
    Call SetRect(pSizeGripRect, W - SG, 0, W, H)
    '-- Set text Rects. (Edge and text)
    Call SetRect(pEdgeRect(1), 0, 2, W - INFO_WIDTH - ZOOM_WIDTH - SG - 2, H)
    Call SetRect(pEdgeRect(2), W - INFO_WIDTH - ZOOM_WIDTH - SG, 2, W - ZOOM_WIDTH - SG - 2, H)
    Call SetRect(pEdgeRect(3), W - ZOOM_WIDTH - SG, 2, W - SG, H)
    For i = 1 To 3
        CopyRect pTextRect(i), pEdgeRect(i)
        With pTextRect(i)
            .x1 = .x1 + 4
            .y1 = .y1 + 2
            .x2 = .x2 - 4
        End With
    Next i
    '-- Refresh all
    Call Refresh
    '-- Resize
    RaiseEvent Resize
    
    On Error GoTo 0
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If (x > pSizeGripRect.x1) Then
        Call ReleaseCapture
        Call SendMessage(Parent.hWnd, WM_NCLBUTTONDOWN, HTBOTTOMRIGHT, ByVal 0)
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If (x > pSizeGripRect.x1) Then
        MousePointer = vbSizeNWSE
      Else
        MousePointer = vbDefault
   End If
End Sub

'========================================================================================
' Properties
'========================================================================================

Public Property Let TextFile(ByVal New_TextFile As String)
    m_TextFile = New_TextFile
    Call Refresh
End Property

Public Property Get TextFile() As String
Attribute TextFile.VB_MemberFlags = "400"
    TextFile = m_TextFile
End Property

Public Property Get TextFileWidth() As Long
    TextFileWidth = pTextRect(1).x2 - pTextRect(1).x1
End Property

Public Property Let TextInfo(ByVal New_TextInfo As String)
    m_TextInfo = New_TextInfo
    Call Refresh
End Property

Public Property Let TextZoom(ByVal New_TextZoom As String)
    m_TextZoom = New_TextZoom
    Call Refresh
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Font = Ambient.Font
End Sub
