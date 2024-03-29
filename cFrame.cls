VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'================================================
' Class:         cFrame.cls (Rect. region control)
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

Private Declare Function SetRect Lib "user32" (lpRect As RECT2, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function IntersectRect Lib "user32" (lpDestRect As RECT2, lpSrc1Rect As RECT2, lpSrc2Rect As RECT2) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT2, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetRectEmpty Lib "user32" (lpRect As RECT2) As Long
Private Declare Function IsRectEmpty Lib "user32" (lpRect As RECT2) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT2, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hDC As Long, lpRect As RECT2, ByVal hBrush As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Integer) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long

'-- Private Variables:

Private m_RectMain    As RECT2  ' Main rectangle region
Private m_Rect        As RECT2  ' Frame
Private m_RectScaled  As RECT2  ' Frame scaled
Private m_ScaleFactor As Single ' Scale factor

Private m_hBrush(7)   As Long   ' Contour Frame pattern bitmaps
Private m_nBrush      As Long   ' Current brush

'========================================================================================
' Class
'========================================================================================

Private Sub Class_Initialize()

  Dim lIdx           As Long
  Dim lIdy           As Long
  Dim hBitmap        As Long
  Dim tBytes(1 To 8) As Integer
  Dim bBytes(1 To 8) As Integer
    
    '-- Brush main pattern (8x8)
    tBytes(1) = &HCC
    tBytes(2) = &H66
    tBytes(3) = &H33
    tBytes(4) = &H99
    tBytes(5) = &HCC
    tBytes(6) = &H66
    tBytes(7) = &H33
    tBytes(8) = &H99
    
    '-- Create main brush
    hBitmap = CreateBitmap(8, 8, 1, 1, tBytes(1))
    m_hBrush(0) = CreatePatternBrush(hBitmap)
    DeleteObject hBitmap

    '-- 'Scroll' pattern: Create next brushes
    For lIdx = 7 To 1 Step -1
        For lIdy = 1 To 8
            If (lIdx + lIdy < 9) Then
                bBytes(lIdy) = tBytes(lIdx + lIdy)
              Else
                bBytes(lIdy) = tBytes(lIdx + lIdy - 8)
            End If
        Next lIdy
        hBitmap = CreateBitmap(8, 8, 1, 1, bBytes(1))
        m_hBrush(lIdx) = CreatePatternBrush(hBitmap)
        Call DeleteObject(hBitmap)
    Next lIdx
    
    '-- Initialize scale factor
    m_ScaleFactor = 1
End Sub

Private Sub Class_Terminate()

  Dim lIdx As Long
  
    '-- Delete all brushes
    For lIdx = 0 To 7
        If (m_hBrush(lIdx) <> 0) Then
            Call DeleteObject(m_hBrush(lIdx))
        End If
    Next lIdx
End Sub

'========================================================================================
' Methods
'========================================================================================

Public Sub Init(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long)

    Call SetRect(m_RectMain, x1, y1, x2, y2)
    Call SetRectEmpty(m_Rect)
    Call SetRectEmpty(m_RectScaled)
End Sub

Public Sub Clear()

    Call SetRectEmpty(m_Rect)
    Call SetRectEmpty(m_RectScaled)
End Sub

Public Sub SetFrameRect(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, Optional ByVal CropMain As Boolean = 0)
    
    Call SetRect(m_Rect, x1, y1, x2, y2)
    Call pvNormalizeRect(m_Rect)
    If (CropMain) Then
        Call Crop
      Else
        Call pvScaleRect(m_ScaleFactor)
    End If
End Sub

Public Sub Crop()

    Call IntersectRect(m_Rect, m_Rect, m_RectMain)
    Call pvScaleRect(m_ScaleFactor)
End Sub

Public Sub Offset(ByVal dx As Long, ByVal dy As Long)

    Call OffsetRect(m_Rect, dx, dy)
    Call pvScaleRect(m_ScaleFactor)
End Sub

Public Sub PaintToDC(ByVal hDC As Long, Optional ByVal xOffset As Long = 0, Optional ByVal yOffset As Long = 0)
    
  Dim lOldBkColor As Long
  Dim lOldFrColor As Long
    
    If (IsRectEmpty(m_Rect) = 0) Then
    
        '-- Set back = WHITE, fore = BLACK
        lOldBkColor = SetBkColor(hDC, &HFFFFFF)
        lOldFrColor = SetTextColor(hDC, &H0)
        
        '-- Move/Paint/Restore region
        Call OffsetRect(m_RectScaled, xOffset, yOffset)
        Call FrameRect(hDC, m_RectScaled, m_hBrush(m_nBrush))
        Call OffsetRect(m_RectScaled, -xOffset, -yOffset)
        
        '-- Restore back and fore color
        Call SetBkColor(hDC, lOldBkColor)
        Call SetTextColor(hDC, lOldFrColor)
    End If
End Sub

Public Sub RotateBrush()

    m_nBrush = m_nBrush + 1
    If (m_nBrush = 8) Then
        m_nBrush = 0
    End If
End Sub

Public Function GetFrameRect(x1 As Long, y1 As Long, x2 As Long, y2 As Long, Optional ByVal Scaled As Boolean = False) As Boolean

    If (IsRectEmpty(m_Rect) = 0) Then
        If (Scaled = 0) Then
            With m_Rect
                x1 = .x1
                y1 = .y1
                x2 = .x2
                y2 = .y2
            End With
          Else
            With m_RectScaled
                x1 = .x1
                y1 = .y1
                x2 = .x2
                y2 = .y2
            End With
            GetFrameRect = True
        End If
    End If
End Function

Public Function IsFrameEmpty() As Boolean

    IsFrameEmpty = (IsRectEmpty(m_Rect) <> 0)
End Function

Public Function IsPointInFrame(ByVal x As Long, ByVal y As Long) As Boolean

    IsPointInFrame = (PtInRect(m_Rect, x, y) <> 0)
End Function

'========================================================================================
' Properties
'========================================================================================

Public Property Let ScaleFactor(ByVal Factor As Single)
    m_ScaleFactor = IIf(Factor < 0.001, 0.001, Factor)
    Call pvScaleRect(m_ScaleFactor)
End Property

Public Property Get ScaleFactor() As Single
    ScaleFactor = m_ScaleFactor
End Property

'========================================================================================
' Private
'========================================================================================

Private Sub pvNormalizeRect(pRect As RECT2)

  Dim lTmp As Long
        
    With pRect
        If (.y2 < .y1) Then
            lTmp = .y2
            .y2 = .y1
            .y1 = lTmp
        End If
        If (.x2 < .x1) Then
            lTmp = .x2
            .x2 = .x1
            .x1 = lTmp
        End If
    End With
End Sub

Private Sub pvScaleRect(ByVal Factor As Single)
    
    '-- Scale frame
    With m_RectScaled
        .x1 = m_Rect.x1 * Factor
        .y1 = m_Rect.y1 * Factor
        .x2 = m_Rect.x2 * Factor
        .y2 = m_Rect.y2 * Factor
    End With
End Sub
