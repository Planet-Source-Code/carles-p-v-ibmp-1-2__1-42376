VERSION 5.00
Begin VB.UserControl ucCanvas 
   BackColor       =   &H8000000C&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2220
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3060
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   148
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   204
   Begin VB.PictureBox iCanvas 
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   1185
      Left            =   0
      MousePointer    =   99  'Custom
      ScaleHeight     =   79
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   77
      TabIndex        =   0
      Top             =   0
      Width           =   1155
   End
   Begin VB.Menu mnuCropTop 
      Caption         =   "Crop"
      Visible         =   0   'False
      Begin VB.Menu mnuCrop 
         Caption         =   "Crop"
         Index           =   0
      End
      Begin VB.Menu mnuCrop 
         Caption         =   "Copy"
         Index           =   1
      End
      Begin VB.Menu mnuCrop 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuCrop 
         Caption         =   "Remove frame"
         Index           =   3
      End
      Begin VB.Menu mnuCrop 
         Caption         =   "Cancel"
         Index           =   4
      End
   End
End
Attribute VB_Name = "ucCanvas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================
' UserControl:   ucCanvas.ctl
' Author:        Carles P.V.
' Dependencies:  cDIB.cls
' Last revision: 2003.11.02
'================================================

Option Explicit

'-- API:

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type RECT2
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

Private Const RGN_DIFF           As Long = 4
Private Const COLOR_APPWORKSPACE As Long = 12

Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT2) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT2, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function FillRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

'//

'-- Public enums.:
Public Enum cnvWorkModeCts
    [cnvScrollMode]
    [cnvCropMode]
    [cnvPickColorMode]
End Enum

'-- Private Variables:
Private Frame      As New cFrame
Private m_Zoom     As Long
Private m_WorkMode As cnvWorkModeCts
Private m_FitMode  As Boolean
Private m_Enabled  As Boolean
Private m_hPos     As Long
Private m_hMax     As Long
Private m_vPos     As Long
Private m_vMax     As Long
Private m_Down     As Boolean
Private m_cPt      As POINTAPI
Private m_cRct     As RECT2
Private m_lsthPos  As Single
Private m_lstvPos  As Single
Private m_lsthMax  As Single
Private m_lstvMax  As Single

'-- Public Objects:
Public WithEvents DIB As cDIB
Attribute DIB.VB_VarHelpID = -1

'-- Public events:
Public Event DIBProgress(ByVal p As Long)
Public Event DIBProgressEnd()
Public Event MouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Long, y As Long)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event Scroll()
Public Event Resize()
Public Event Crop()

'========================================================================================
' Initialization
'========================================================================================

Private Sub UserControl_Initialize()

    '-- Initialize DIB
    Set DIB = New cDIB
    
    '-- Default values
    m_Zoom = 1
    m_WorkMode = [cnvScrollMode]
    
    '-- 'Hide' preview Crop rect.
    Call SetRect(m_cRct, -1, -1, -1, -1)
End Sub

'========================================================================================
' Prigress events
'========================================================================================

Private Sub DIB_Progress(ByVal p As Long)
    RaiseEvent DIBProgress(p)
End Sub

Private Sub DIB_ProgressEnd()
    RaiseEvent DIBProgressEnd
End Sub

'========================================================================================
' Refreshing / Resizing
'========================================================================================

Private Sub iCanvas_Paint()
  
  Dim xOff As Long, yOff As Long
  Dim wDst As Long, hDst As Long
  Dim xSrc As Long, ySrc As Long
  Dim wSrc As Long, hSrc As Long
    
    If (DIB.hDIB <> 0) Then
        
        '-- Get Left and Width of source image rectangle:
        If (m_hMax And m_FitMode = 0) Then
            xOff = -m_hPos Mod m_Zoom
            wDst = (iCanvas.Width \ m_Zoom) * m_Zoom + 2 * m_Zoom
            xSrc = m_hPos \ m_Zoom
            wSrc = iCanvas.Width \ m_Zoom + 2
          Else
            xOff = 0
            wDst = iCanvas.Width
            xSrc = 0
            wSrc = DIB.Width
        End If
        '-- Get Top and Height of source image rectangle:
        If (m_vMax And m_FitMode = 0) Then
            yOff = -m_vPos Mod m_Zoom
            hDst = (iCanvas.Height \ m_Zoom) * m_Zoom + 2 * m_Zoom
            ySrc = m_vPos \ m_Zoom
            hSrc = iCanvas.Height \ m_Zoom + 2
          Else
            yOff = 0
            hDst = iCanvas.Height
            ySrc = 0
            hSrc = DIB.Height
        End If
        '-- Paint visible source rectangle:
        Call DIB.Stretch(iCanvas.hDC, xOff, yOff, wDst, hDst, xSrc, ySrc, wSrc, hSrc)
        
        '-- Paint Crop rectangle
        If (m_FitMode = False) Then
            Call Frame.PaintToDC(iCanvas.hDC, -m_hPos, -m_vPos)
          Else
            Call Frame.PaintToDC(iCanvas.hDC)
        End If
    End If
End Sub

Public Sub Repaint()
    Call iCanvas_Paint
End Sub

Public Sub Resize()
    Call UserControl_Resize
End Sub

Private Sub UserControl_Resize()
    
  Dim rW As Long, rH As Long
  Dim bfx As Long, bfy As Long
  Dim bfW As Long, bfH As Long
  
    With DIB
        
        If (.hDIB <> 0) Then
        
            If (m_FitMode = 0) Then
            
                '-- Get new Width
                If (.Width * m_Zoom > ScaleWidth) Then
                    m_hMax = .Width * m_Zoom - ScaleWidth
                    rW = ScaleWidth
                  Else
                    m_hMax = 0
                    rW = .Width * m_Zoom
                End If
                '-- Get new Height
                If (.Height * m_Zoom > ScaleHeight) Then
                    m_vMax = .Height * m_Zoom - ScaleHeight
                    rH = ScaleHeight
                  Else
                    m_vMax = 0
                    rH = .Height * m_Zoom
                End If
              
              Else
                DIB.GetBestFitInfo ScaleWidth, ScaleHeight, bfx, bfy, bfW, bfH
            End If
            
            '-- Resize
            If (m_FitMode = 0) Then
                Call MoveWindow(iCanvas.hWnd, (ScaleWidth - rW) \ 2, (ScaleHeight - rH) \ 2, rW, rH, 0)
                Frame.ScaleFactor = m_Zoom
              Else
                Call MoveWindow(iCanvas.hWnd, bfx, bfy, bfW, bfH, 0)
                Frame.ScaleFactor = IIf(bfW > bfH, bfW / DIB.Width, bfH / DIB.Height)
            End If
                                
            '== Memory position:
            '-- Horizontal position
            If (m_lsthMax) Then
                m_hPos = m_lsthPos * m_hMax / m_lsthMax
              Else
                m_hPos = m_hMax / 2
            End If
            '-- Vertical position
            If (m_lstvMax) Then
                m_vPos = m_lstvPos * m_vMax / m_lstvMax
              Else
                m_vPos = m_vMax / 2
            End If
            '-- Save values
            m_lsthPos = m_hPos: m_lstvPos = m_vPos
            m_lsthMax = m_hMax: m_lstvMax = m_vMax
            
            '-- Refresh
            Call pvCls
            Call Me.Repaint
            
            '-- Raise Resize event
            RaiseEvent Resize
          
          Else
            '-- Cls (whole)
            Call iCanvas.Move(-1, -1, 0, 0)
            Call pvCls
        End If
    End With
    
    '-- Reset pointer
    iCanvas.MouseIcon = Nothing
    '-- Change it
    If (m_WorkMode = [cnvScrollMode]) Then
        If ((m_hMax Or m_vMax) And Not m_FitMode) Then
            iCanvas.MouseIcon = LoadResPicture("CURSOR_HANDOVER", vbResCursor)
        End If
      ElseIf (m_WorkMode = [cnvPickColorMode]) Then
        iCanvas.MouseIcon = LoadResPicture("CURSOR_PICKCOLOR", vbResCursor)
    End If
End Sub

'========================================================================================
' Scrolling
'========================================================================================

Private Sub iCanvas_DblClick()
    Call iCanvas_MouseDown(vbLeftButton, 0, (m_cPt.x), CSng(m_cPt.y))
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
        
    '-- iCanvas offset
    x = x - iCanvas.Left
    y = y - iCanvas.Top
    
    '-- Set mouse capture to iCanvas
    Call SetCapture(iCanvas.hWnd)
    Call iCanvas_MouseDown(Button, Shift, x, y)
End Sub

Private Sub iCanvas_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If (DIB.hDIB <> 0) Then
    
        '-- Change pointer
        If ((m_hMax Or m_vMax) And m_WorkMode = [cnvScrollMode] And Not m_FitMode) Then
            iCanvas.MouseIcon = LoadResPicture("CURSOR_HANDCATCH", vbResCursor)
        End If
        
        '-- Inside frame...
        If (Frame.IsPointInFrame(pvDIBx(x), pvDIBy(y))) Then
            If (Button = vbRightButton) Then
                '-- Show <Crop> menu
                Call PopupMenu(mnuCropTop)
            End If
        End If
        If (Button = vbLeftButton And m_WorkMode = [cnvCropMode]) Then
            '-- Initialize frame's main region
            Call Frame.Init(0, 0, DIB.Width, DIB.Height)
            Call Me.Repaint
        End If
        
        m_Down = (Button = vbLeftButton)
        m_cPt.x = x
        m_cPt.y = y
    End If
    
    RaiseEvent MouseDown(Button, Shift, pvDIBx(x), pvDIBy(y))
End Sub

Private Sub iCanvas_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
        
    If (m_Down) Then
    
        Select Case m_WorkMode
        
            Case [cnvScrollMode]
                '-- Get displacements
                m_hPos = m_hPos + (m_cPt.x - x)
                m_vPos = m_vPos + (m_cPt.y - y)
                '-- Check margins
                If (m_hPos < 0) Then m_hPos = 0
                If (m_vPos < 0) Then m_vPos = 0
                If (m_hPos > m_hMax) Then m_hPos = m_hMax
                If (m_vPos > m_vMax) Then m_vPos = m_vMax
                '-- Save current position
                m_cPt.x = x
                m_cPt.y = y
                
                If (m_lsthPos <> m_hPos Or m_lstvPos <> m_vPos) Then
                    '-- Refresh
                    Call Me.Repaint
                    '-- Raise Scroll event
                    RaiseEvent Scroll
                End If
                m_lsthPos = m_hPos
                m_lstvPos = m_vPos
            
            Case [cnvCropMode]
                '-- Paint current frame (Invert mode)
                Call DrawFocusRect(iCanvas.hDC, m_cRct)
                Call Frame.SetFrameRect(pvDIBx(m_cPt.x), pvDIBy(m_cPt.y), pvDIBx(x), pvDIBy(y))
                Call Frame.GetFrameRect(m_cRct.x1, m_cRct.y1, m_cRct.x2, m_cRct.y2, True)
                Call pvRectClientToCanvas(m_cRct)
                Call DrawFocusRect(iCanvas.hDC, m_cRct)
        End Select
    End If
    
    RaiseEvent MouseMove(Button, Shift, pvDIBx(x), pvDIBy(y))
End Sub

Private Sub iCanvas_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    '-- Change pointer
    If ((m_hMax Or m_vMax) And WorkMode = [cnvScrollMode] And Not m_FitMode) Then
        iCanvas.MouseIcon = LoadResPicture("CURSOR_HANDOVER", vbResCursor)
    End If
    
    '-- Reset temp. rectangle
    Call DrawFocusRect(iCanvas.hDC, m_cRct)
    '-- 'Hide' it
    Call SetRect(m_cRct, -1, -1, -1, -1)
    
    If (m_Down And m_WorkMode = [cnvCropMode]) Then
        '-- Set frame rectangle (Crop to main) and paint it
        Call Frame.Crop
        If (m_FitMode = 0) Then
            Call Frame.PaintToDC(iCanvas.hDC, -m_hPos, -m_vPos)
          Else
            Call Frame.PaintToDC(iCanvas.hDC)
        End If
    End If
    m_Down = False
    
    RaiseEvent MouseUp(Button, Shift, pvDIBx(x), pvDIBy(y))
End Sub

Private Sub iCanvas_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Not m_Down) Then RaiseEvent KeyDown(KeyCode, Shift)
End Sub

'========================================================================================
' Propery-functions / Methods / Properties
'========================================================================================

Public Sub GetScrollInfo(lHMax As Long, lVMax As Long, lHPos As Long, lVPos As Long)
    '-- Pass scroll values
    lHMax = m_hMax
    lVMax = m_vMax
    lHPos = m_hPos
    lVPos = m_vPos
End Sub

Public Sub SetScrollInfo(ByVal lHPos As Long, ByVal lVPos As Long)
    '-- Check bounds
    If (lHPos < 0) Then lHPos = 0 Else If (lHPos > m_hMax) Then lHPos = m_hMax
    If (lVPos < 0) Then lVPos = 0 Else If (lVPos > m_vMax) Then lVPos = m_vMax
    '-- Pass scroll pos. values
    m_hPos = lHPos
    m_vPos = lVPos
    '-- Update last
    m_lsthPos = lHPos
    m_lstvPos = lVPos
    '-- Refresh
    Call Me.Repaint
End Sub

Public Sub GetVisibleRect(x1 As Long, y1 As Long, x2 As Long, y2 As Long)
    '-- Pass rectangle coords.
    x1 = m_hPos \ m_Zoom
    y1 = m_vPos \ m_Zoom
    x2 = x1 + iCanvas.Width \ m_Zoom
    y2 = y1 + iCanvas.Height \ m_Zoom
End Sub

Public Sub RemoveCropRectangle()
    '-- Clear frame
    Call Frame.Clear
End Sub

'//

Public Property Let Zoom(ByVal Factor As Long)
    m_Zoom = IIf(Factor < 1, 1, Factor)
End Property

Public Property Get Zoom() As Long
    Zoom = m_Zoom
End Property

Public Property Let WorkMode(ByVal Mode As cnvWorkModeCts)

    '-- Reset pointer
    iCanvas.MouseIcon = Nothing
    
    '-- Change it
    If (Mode = [cnvScrollMode]) Then
        Call Me.RemoveCropRectangle
        If ((m_hMax Or m_vMax) And Not m_FitMode) Then
            iCanvas.MouseIcon = LoadResPicture("CURSOR_HANDOVER", vbResCursor)
        End If
      ElseIf (Mode = [cnvPickColorMode]) Then
        iCanvas.MouseIcon = LoadResPicture("CURSOR_PICKCOLOR", vbResCursor)
    End If
    m_WorkMode = Mode
End Property

Public Property Get WorkMode() As cnvWorkModeCts
    WorkMode = m_WorkMode
End Property

Public Property Let FitMode(ByVal Enable As Boolean)
    m_FitMode = Enable
End Property

Public Property Get FitMode() As Boolean
    FitMode = m_FitMode
End Property

Public Property Let Enabled(ByVal Enable As Boolean)
    UserControl.Enabled = Enable
End Property

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

'========================================================================================
' Private
'========================================================================================

Private Sub pvCls()

  Dim hRgn_1 As Long
  Dim hRgn_2 As Long
  Dim hBrush As Long
    
    '-- Create brush (background)
    hBrush = GetSysColorBrush(COLOR_APPWORKSPACE)
    
    '-- Create Cls region (Control Rect. - iCanvas Rect.)
    With iCanvas
        hRgn_1 = CreateRectRgn(0, 0, ScaleWidth, ScaleHeight)
        hRgn_2 = CreateRectRgn(.Left, .Top, .Left + .Width - 1, .Top + .Height - 1)
    End With
    Call CombineRgn(hRgn_1, hRgn_1, hRgn_2, RGN_DIFF)
    
    '-- Fill it
    Call FillRgn(hDC, hRgn_1, hBrush)
    
    '-- Clear
    Call DeleteObject(hBrush)
    Call DeleteObject(hRgn_1)
    Call DeleteObject(hRgn_2)
End Sub

Private Function pvDIBx(ByVal x As Long) As Long
    pvDIBx = Int((IIf(m_FitMode, 0, m_hPos) + x) / Frame.ScaleFactor)
End Function

Private Function pvDIBy(ByVal y As Long) As Long
    pvDIBy = Int((IIf(m_FitMode, 0, m_vPos) + y) / Frame.ScaleFactor)
End Function

Private Function pvRectClientToCanvas(lpRect As RECT2) As Long
    With lpRect
        .x1 = .x1 - IIf(m_FitMode, 0, m_hPos)
        .y1 = .y1 - IIf(m_FitMode, 0, m_vPos)
        .x2 = .x2 - IIf(m_FitMode, 0, m_hPos)
        .y2 = .y2 - IIf(m_FitMode, 0, m_vPos)
    End With
End Function

'========================================================================================
' Menu
'========================================================================================

Private Sub mnuCrop_Click(Index As Integer)
    
  Dim TmpDIB As New cDIB
  
  Dim x1 As Long, y1 As Long
  Dim x2 As Long, y2 As Long
    
    Select Case Index
      
        Case 0 '-- Crop
            '-- Get frame coords.
            Call Frame.GetFrameRect(x1, y1, x2, y2)
            '-- Create temp. DIB
            Call TmpDIB.Create(x2 - x1, y2 - y1)
            Call TmpDIB.LoadBlt(DIB.hDC, x1, y1)
            '-- Resize and get from temp. DIB
            Call DIB.Create(x2 - x1, y2 - y1)
            Call DIB.LoadBlt(TmpDIB.hDC)
            '-- Clear Frame
            Call Frame.Clear
            
            '-- Resize Canvas, refresh and raise <Crop> event
            Call Me.Resize
            Call Me.Repaint
            RaiseEvent Crop
        
        Case 1 '-- Copy
            '-- Get frame coords.
            Call Frame.GetFrameRect(x1, y1, x2, y2)
            '-- Create temp. DIB
            Call TmpDIB.Create(x2 - x1, y2 - y1)
            Call TmpDIB.LoadBlt(DIB.hDC, x1, y1)
            '-- Copy to clipboard
            Call TmpDIB.CopyToClipboard
            
        Case 3 '-- Remove frame
            Call Me.RemoveCropRectangle
            Call Me.Repaint
    End Select
End Sub

