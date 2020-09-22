VERSION 5.00
Begin VB.Form fPanView 
   BackColor       =   &H00000000&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Panoramic view"
   ClientHeight    =   2610
   ClientLeft      =   11250
   ClientTop       =   2220
   ClientWidth     =   3225
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fPanView.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   174
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   215
   ShowInTaskbar   =   0   'False
   Begin VB.Menu mnuZoomIn 
      Caption         =   " +"
   End
   Begin VB.Menu mnuZoomOut 
      Caption         =   " -"
   End
   Begin VB.Menu mnuRealSize 
      Caption         =   " 1:1"
   End
End
Attribute VB_Name = "fPanView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'================================================
' fPanView form
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

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Const RGN_DIFF As Long = 4
Private Const PS_SOLID As Long = 0

Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function FillRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT2) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT2, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT2, ByVal x As Long, ByVal y As Long) As Long

'-- Private Variables:

Private m_bfx As Long, m_bfy As Long
Private m_bfW As Long, m_bfH As Long

Private m_xD     As Long
Private m_yD     As Long
Private m_uRct   As RECT2
Private m_bInRct As Boolean

'//

Private Sub Form_Load()
    '-- Load settings
    Call mSettings.LoadPanViewSettings
End Sub

Private Sub Form_Resize()

    '-- Get best fit info
    With fMain.Canvas.DIB
        Call .GetBestFitInfo(ScaleWidth, ScaleHeight, m_bfx, m_bfy, m_bfW, m_bfH)
    End With
    '-- Refresh
    Call pvFlickerlessCls
    Call Form_Paint
End Sub

Private Sub Form_Paint()
  
  Dim cF      As Single
  Dim hPen    As Long
  Dim hOldPen As Long
  
    With fMain.Canvas
    
        If (.DIB.hDIB <> 0) Then
        
            '-- Paint DIB
            Call .DIB.Stretch(Me.hDC, m_bfx, m_bfy, m_bfW, m_bfH, 0, 0, .DIB.Width, .DIB.Height)  ' [dsHalftone] 'W2000/NT/XP
            
            '-- Get visible rectangle
            Call .GetVisibleRect(m_uRct.x1, m_uRct.y1, m_uRct.x2, m_uRct.y2)
            '-- Scale it
            If (.FitMode) Then
                With m_uRct
                    .x1 = m_bfx
                    .x2 = m_bfx + m_bfW
                    .y1 = m_bfy
                    .y2 = m_bfy + m_bfH
                End With
              Else
                cF = IIf(m_bfW > m_bfH, m_bfW / .DIB.Width, m_bfH / .DIB.Height)
                With m_uRct
                    .x1 = m_bfx + .x1 * cF
                    .x2 = m_bfx + .x2 * cF + 1: If (.x2 > m_bfx + m_bfW) Then .x2 = m_bfx + m_bfW
                    .y1 = m_bfy + .y1 * cF
                    .y2 = m_bfy + .y2 * cF + 1: If (.y2 > m_bfy + m_bfH) Then .y2 = m_bfy + m_bfH
                End With
            End If
            '-- Draw it
            hPen = CreatePen(PS_SOLID, 2, vbRed)
            hOldPen = SelectObject(hDC, hPen)
            Call Rectangle(hDC, m_uRct.x1 + 1, m_uRct.y1 + 1, m_uRct.x2, m_uRct.y2)
            Call SelectObject(hDC, hOldPen)
            Call DeleteObject(hPen)
        End If
    End With
End Sub

Public Sub Repaint()
    Call Form_Resize
End Sub

'//

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
        
  Dim WRect As RECT2
  Dim cRect As RECT2
  Dim PtCur As POINTAPI
  Dim PtOff As POINTAPI
    
    '-- In rect.?
    Call GetCursorPos(PtCur)
    Call ScreenToClient(Me.hWnd, PtCur)
    m_bInRct = (PtInRect(m_uRct, PtCur.x, PtCur.y) <> 0)
    
    If (m_bInRct) Then
    
        '-- Get client rect.
        Call GetClientRect(hWnd, WRect)
        PtOff.x = WRect.x1
        PtOff.y = WRect.y1
        Call ClientToScreen(Me.hWnd, PtOff)
        
        '-- Adjust to image rect.
        With WRect
            .x1 = .x1 + m_bfx
            .x2 = .x2 - m_bfx
            .y1 = .y1 + m_bfy
            .y2 = .y2 - m_bfy
        End With
        
        '-- Client to screen coords.
        Call OffsetRect(WRect, PtOff.x, PtOff.y)
        
        '-- Clip
        Call ClipCursor(WRect)
        
        '-- Save current position
        m_xD = x
        m_yD = y
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
  Dim scrHMax As Long, scrVMax As Long
  Dim scrHPos As Long, scrVPos As Long
  Dim Pt As POINTAPI
  Dim cF As Single
    
    '-- Get cursor coords.
    Call GetCursorPos(Pt)
    Call ScreenToClient(Me.hWnd, Pt)
    
    '-- Scroll...
    If ((Button = vbLeftButton And m_bInRct) And fMain.Canvas.DIB.hDIB <> 0) Then
        '-- Scroll main
        With fMain.Canvas
            cF = IIf(m_bfW > m_bfH, (.DIB.Width * .Zoom) / m_bfW, (.DIB.Height * .Zoom) / m_bfH)
            Call .GetScrollInfo(scrHMax, scrVMax, scrHPos, scrVPos)
            Call .SetScrollInfo(scrHPos + (Pt.x - m_xD) * cF, scrVPos + (Pt.y - m_yD) * cF)
        End With
        '-- Refresh me
        Call Form_Paint
    End If
    '-- Save current position
    m_xD = Pt.x
    m_yD = Pt.y
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call ClipCursor(ByVal 0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '-- Save settings
    Call mSettings.SavePanViewSettings
    '-- Update main menu
    fMain.mnuView(1).Checked = False
End Sub

'//

Private Sub mnuZoomIn_Click()
    Call fMain.mnuZoom_Click(0)
End Sub

Private Sub mnuZoomOut_Click()
    Call fMain.mnuZoom_Click(1)
End Sub

Private Sub mnuRealSize_Click()
    Call fMain.mnuZoom_Click(2)
End Sub

'//

Private Sub pvFlickerlessCls()
    
  Dim hRgn_1 As Long
  Dim hRgn_2 As Long
  Dim hBrush As Long
    
    '-- Create a black brush
    hBrush = CreateSolidBrush(0)
    
    '-- Create Cls region (Form client-rect. - Image rect.)
    hRgn_1 = CreateRectRgn(0, 0, ScaleWidth, ScaleHeight)
    hRgn_2 = CreateRectRgn(m_bfx, m_bfy, m_bfx + m_bfW, m_bfy + m_bfH)
    Call CombineRgn(hRgn_1, hRgn_1, hRgn_2, RGN_DIFF)
    
    '-- Fill it
    Call FillRgn(Me.hDC, hRgn_1, hBrush)
    
    '-- Clear
    Call DeleteObject(hBrush)
    Call DeleteObject(hRgn_1)
    Call DeleteObject(hRgn_2)
End Sub
