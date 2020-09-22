VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDIBDither"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'================================================
' Class:         cDIBDither.cls
' Author:        Carles P.V. (*)
' Dependencies:  cDIB.cls, cDIBPal.cls
' Last revision: 2003.11.02
'================================================
'
' cDIBDither:
' 1, 4 and 8 bpp dithering [Palette indexes array storing]
'
' (*)
'
' - Color depth translations from:
'   VB Frood, http://perso.wanadoo.fr/vbfrood/deutsch/projects/bmpread.htm
'
' - Stucki 'Black and White' reduction code by Manuel Augusto Santos
'
' - Color dithering (DitherToColorPalette original code) from VB Accelerator
'   + Pre-reduction of source DIB to 4096 (16x16x16) color space (Speed up)
'   + Weighted RGB channels (16 colors case). See <cDIBPal.CreateOptimal>
'
' - ColourCount function by Robert Rayment

Option Explicit

'-- API:

Private Type RGBQUAD
    B As Byte
    G As Byte
    R As Byte
    A As Byte
End Type

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound   As Long
End Type

Private Type SAFEARRAY2D
    cDims      As Integer
    fFeatures  As Integer
    cbElements As Long
    cLocks     As Long
    pvData     As Long
    Bounds(1)  As SAFEARRAYBOUND
End Type

Private Declare Function GetNearestPaletteIndex Lib "gdi32" (ByVal hPalette As Long, ByVal crColor As Long) As Long
Private Declare Function VarPtrArray Lib "msvbvm50" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal ByteLength As Long)

'-- Public events:
Public Event Progress(ByVal p As Long)
Public Event ProgressEnd()

'-- Private Variables:
Private m_Pow2(31) As Long
Private m_PalID()  As Byte

'========================================================================================
' Class
'========================================================================================

Private Sub Class_Initialize()
  
  Dim lIdx As Long
  
    '-- Init look-up table for fast 2 ^ x
    m_Pow2(0) = 1
    For lIdx = 1 To 30
        m_Pow2(lIdx) = 2 * m_Pow2(lIdx - 1)
    Next
    m_Pow2(31) = &H80000000
End Sub

Private Sub Class_Terminate()
    Erase m_PalID()
End Sub

'========================================================================================
' Methods
'========================================================================================

Public Function InitializePalID(ByVal nWidth As Long, ByVal nHeight As Long, ByVal nbpp As Byte) As Long

    '-- Resize PalID table
    ReDim m_PalID(((nWidth * nbpp + 31) \ 32) * 4 - 1, nHeight - 1) As Byte
    
    '-- Return pointer
    InitializePalID = VarPtr(m_PalID(0, 0))
End Function

'//

Public Sub DitherToBlackAndWhite(DIB As cDIB, ByVal PalWeight As Long)
'-- Stucki

  Dim Bits() As RGBQUAD
  Dim tSA    As SAFEARRAY2D
  
  Dim yRel As Single
  
  Dim x As Long
  Dim y As Long
  Dim W As Long, W1 As Long, W2 As Long
  Dim H As Long, H1 As Long, H2 As Long
  
  Dim Spd(765) As Long
  Dim R As Long
  Dim G As Long
  Dim B As Long
  
  Dim PartErr(1 To 12, -255 To 255) As Long
  Dim Erro     As Long
  Dim VecErr() As Long
  Dim nCol     As Long
  Dim mCol     As Long
  
    If (DIB.hDIB <> 0) Then
        
        Call pvBuildSA(tSA, DIB)
        Call CopyMemory(ByVal VarPtrArray(Bits()), VarPtr(tSA), 4)
        
        W = DIB.Width - 1
        H = DIB.Height - 1
        W1 = W - 1: W2 = W - 2
        H1 = H - 1: H2 = H - 2
        
        Call InitializePalID(DIB.Width, DIB.Height, 1)
        
        For x = 0 To 765
            Spd(x) = x \ 3
        Next x
        For x = -255 To 255
            PartErr(1, x) = (8 * x) \ 42
            PartErr(2, x) = (4 * x) \ 42
            PartErr(3, x) = (2 * x) \ 42
            PartErr(4, x) = (4 * x) \ 42
            PartErr(5, x) = (8 * x) \ 42
            PartErr(6, x) = (4 * x) \ 42
            PartErr(7, x) = (2 * x) \ 42
            PartErr(8, x) = (1 * x) \ 42
            PartErr(9, x) = (2 * x) \ 42
            PartErr(10, x) = (4 * x) \ 42
            PartErr(11, x) = (2 * x) \ 42
            PartErr(12, x) = (1 * x) \ 42
        Next x
        
        ReDim VecErr(1 To 3, W) As Long
        For x = 0 To W
            VecErr(1, x) = 0
            VecErr(2, x) = 0
            VecErr(3, x) = 0
        Next x
    
        For y = 0 To H
            For x = 0 To W
                B = Bits(x, y).B
                G = Bits(x, y).G
                R = Bits(x, y).R
                B = Spd(R + G + B)
                mCol = mCol + B
                nCol = nCol + 1
            Next x
            yRel = yRel + 0.5
            RaiseEvent Progress(yRel)
        Next y

        mCol = mCol \ nCol

        For y = 0 To H
            For x = 0 To W
                
                With Bits(x, y)
                
                    B = .B
                    G = .G
                    R = .R
                    
                    B = Spd(R + G + B)
                    B = B + (VecErr(1, x) * 10) \ PalWeight
                    If (B < 0) Then B = 0
                    If (B > 255) Then B = 255
                    
                    If (B < mCol) Then
                        nCol = 0
                        m_PalID(x \ 8, y) = m_PalID(x \ 8, y) And Not m_Pow2(7 - (x Mod 8))
                      Else
                        nCol = 255
                        m_PalID(x \ 8, y) = m_PalID(x \ 8, y) Or m_Pow2(7 - (x Mod 8))
                    End If
                    
                    .B = nCol
                    .G = nCol
                    .R = nCol
                End With
                
                Erro = (B - nCol)
                If (x < W1) Then VecErr(1, x + 1) = VecErr(1, x + 1) + PartErr(1, Erro)
                If (x < W2) Then VecErr(1, x + 2) = VecErr(1, x + 2) + PartErr(2, Erro)
                If (y < H1) Then
                    If (x > 1) Then VecErr(2, x - 2) = VecErr(2, x - 2) + PartErr(3, Erro)
                    If (x > 0) Then VecErr(2, x - 1) = VecErr(2, x - 1) + PartErr(4, Erro)
                    VecErr(2, x) = VecErr(2, x) + PartErr(5, Erro)
                    If (x < W1) Then VecErr(2, x + 1) = VecErr(2, x + 1) + PartErr(6, Erro)
                    If (x < W2) Then VecErr(2, x + 2) = VecErr(2, x + 2) + PartErr(7, Erro)
                End If
                If (y < H2) Then
                    If (x > 1) Then VecErr(3, x - 2) = VecErr(3, x - 2) + PartErr(8, Erro)
                    If (x > 0) Then VecErr(3, x - 1) = VecErr(3, x - 1) + PartErr(9, Erro)
                    VecErr(3, x) = VecErr(3, x) + PartErr(10, Erro)
                    If (x < W1) Then VecErr(3, x + 1) = VecErr(3, x + 1) + PartErr(11, Erro)
                    If (x < W2) Then VecErr(3, x + 2) = VecErr(3, x + 2) + PartErr(12, Erro)
                End If
            Next x
            For x = 0 To W
                VecErr(1, x) = VecErr(2, x)
                VecErr(2, x) = VecErr(3, x)
                VecErr(3, x) = 0
            Next x
            yRel = yRel + 0.5
            RaiseEvent Progress(yRel)
        Next y
        Call CopyMemory(ByVal VarPtrArray(Bits), 0&, 4)
        RaiseEvent ProgressEnd
    End If
End Sub

Public Sub DitherToGreyPalette(DIB As cDIB, Pal As cDIBPal, Optional ByVal Diffuse As Boolean = 0)
'-- Floyd-Steinberg error diffusion

  Dim Bits() As RGBQUAD
  Dim tSA    As SAFEARRAY2D

  Dim x As Long, xIn As Long
  Dim y As Long, yIn As Long
  Dim W As Long
  Dim H As Long

  Dim bpp As Byte
  Dim L   As Byte
  
  Dim errR As Long, errG As Long, errB As Long
  Dim newR As Long, newG As Long, newB As Long
    
    If (DIB.hDIB <> 0) Then
    
        Call pvBuildSA(tSA, DIB)
        Call CopyMemory(ByVal VarPtrArray(Bits()), VarPtr(tSA), 4)

        W = DIB.Width - 1
        H = DIB.Height - 1
        
        '-- Get Palette depth and initialize indexes array
        bpp = Switch(Pal.Count = 256, 8, Pal.Count = 16, 4, Pal.Count = 2, 1)
        Call InitializePalID(DIB.Width, DIB.Height, bpp)
        
        For y = 0 To H
            For x = 0 To W
                
                With Bits(x, y)
                    '-- Get grey value (L --> Pal. index 256; L/16 --> Pal. index 16)
                    L = 0.114 * .B + 0.587 * .G + 0.299 * .R
                    '-- Set true RGB color
                    .B = L
                    .G = L
                    .R = L
                End With
               
                '-- Set 1/4/8 bpp palette indexes
                Select Case bpp
                    Case 8 '-- 256 greys
                        m_PalID(x, y) = L
                    Case 4 '-- 16 greys
                        L = L / 17
                        If (x Mod 2 = 0) Then
                            m_PalID(x \ 2, y) = (m_PalID(x \ 2, y) And &HF) Or L * 16
                          Else
                            m_PalID(x \ 2, y) = (m_PalID(x \ 2, y) And &HF0) Or L
                        End If
                        L = L * 17
                    Case 1 '-- 2 greys (B/W)
                        If (L < 128) Then
                            L = 0
                            m_PalID(x \ 8, y) = m_PalID(x \ 8, y) And Not m_Pow2(7 - (x Mod 8))
                          Else
                            L = 255
                            m_PalID(x \ 8, y) = m_PalID(x \ 8, y) Or m_Pow2(7 - (x Mod 8))
                        End If
                End Select
                
                '-- Diffuse error
                If (Diffuse) Then
                    

                    With Bits(x, y)
                        errB = CLng(.B) - L
                        errG = CLng(.G) - L
                        errR = CLng(.R) - L
                        .B = L
                        .G = L
                        .R = L
                    End With
                    
                    '-- Floyd-Steinberg error diffusion...
                    If (Abs(errB) + Abs(errG) + Abs(errR) > 3) Then
                        If (x < W) Then
                            xIn = x + 1
                            With Bits(xIn, y)
                                newB = .B + (7 * errB) \ 16
                                newG = .G + (7 * errG) \ 16
                                newR = .R + (7 * errR) \ 16
                                If (newB < 0) Then newB = 0 Else If (newB > 255) Then newB = 255
                                If (newG < 0) Then newG = 0 Else If (newG > 255) Then newG = 255
                                If (newR < 0) Then newR = 0 Else If (newR > 255) Then newR = 255
                                .B = newB
                                .G = newG
                                .R = newR
                            End With
                        End If
                        If (y < H) Then
                            yIn = y + 1
                            With Bits(x, yIn)
                                newB = .B + (5 * errB) \ 16
                                newG = .G + (5 * errG) \ 16
                                newR = .R + (5 * errR) \ 16
                                If (newB < 0) Then newB = 0 Else If (newB > 255) Then newB = 255
                                If (newG < 0) Then newG = 0 Else If (newG > 255) Then newG = 255
                                If (newR < 0) Then newR = 0 Else If (newR > 255) Then newR = 255
                                .B = newB
                                .G = newG
                                .R = newR
                            End With
                            If (x < W) Then
                                xIn = x + 1
                                With Bits(xIn, yIn)
                                    newB = .B + errB \ 16
                                    newG = .G + errG \ 16
                                    newR = .R + errR \ 16
                                    If (newB < 0) Then newB = 0 Else If (newB > 255) Then newB = 255
                                    If (newG < 0) Then newG = 0 Else If (newG > 255) Then newG = 255
                                    If (newR < 0) Then newR = 0 Else If (newR > 255) Then newR = 255
                                    .B = newB
                                    .G = newG
                                    .R = newR
                                End With
                            End If
                            If (x > 0) Then
                                xIn = x - 1
                                With Bits(xIn, yIn)
                                    newB = .B + (3 * errB) \ 16
                                    newG = .G + (3 * errG) \ 16
                                    newR = .R + (3 * errR) \ 16
                                    If (newB < 0) Then newB = 0 Else If (newB > 255) Then newB = 255
                                    If (newG < 0) Then newG = 0 Else If (newG > 255) Then newG = 255
                                    If (newR < 0) Then newR = 0 Else If (newR > 255) Then newR = 255
                                    .B = newB
                                    .G = newG
                                    .R = newR
                                End With
                            End If
                        End If
                    End If
                End If
            Next x
            RaiseEvent Progress(y)
        Next y
        Call CopyMemory(ByVal VarPtrArray(Bits), 0&, 4)
        RaiseEvent ProgressEnd
    End If
End Sub

Public Sub DitherToColorPalette(DIB As cDIB, Pal As cDIBPal, Optional ByVal Diffuse As Boolean = 0)
'-- Floyd-Steinberg error diffusion

  Dim Bits() As RGBQUAD
  Dim tSA    As SAFEARRAY2D
  
  Dim x As Long, xIn As Long
  Dim y As Long, yIn As Long
  Dim W As Long
  Dim H As Long
  
  Dim RGB4096(15, 15, 15) As Long
  Dim R As Long
  Dim G As Long
  Dim B As Long
  Dim Spd16(255) As Byte
  
  Dim bpp        As Byte
  Dim tmpEntry() As RGBQUAD
  Dim hTmpPal    As Long
  Dim nEntry     As Byte
  
  Dim errR As Long, errG As Long, errB As Long
  Dim newR As Long, newG As Long, newB As Long
    
    If (DIB.hDIB <> 0) Then
    
        Call pvBuildSA(tSA, DIB)
        Call CopyMemory(ByVal VarPtrArray(Bits()), VarPtr(tSA), 4)

        W = DIB.Width - 1
        H = DIB.Height - 1
        
        '-- Get Palette depth and initialize indexes array
        bpp = Switch(Pal.Count = 256, 8, Pal.Count = 16, 4, Pal.Count = 2, 1)
        Call InitializePalID(DIB.Width, DIB.Height, bpp)
        
        '-- Speed up for GetNearestPaletteIndex
        ReDim tmpEntry(Pal.Count - 1)
        Call CopyMemory(tmpEntry(0), ByVal Pal.PalettePtr, 4 * Pal.Count)
        hTmpPal = Pal.hPal
        
        '-- The real speed up comes here...
        '   This speeds up from 64x64 dimensions (aprox.).
        '   We only are reducing available number of colors to 4096, so we only
        '   need to call the expensive GetNearestPaletteIndex function 4096 times.
        
        '-- Build 4096-colors palette indexes table
        For R = 0 To 15
        For G = 0 To 15
        For B = 0 To 15
            RGB4096(R, G, B) = GetNearestPaletteIndex(hTmpPal, RGB(B * 17, G * 17, R * 17))
        Next B, G, R
        
        '-- Speed up translation
        For x = 0 To 255
            Spd16(x) = (x + 8) \ 17
        Next x
        
        For y = 0 To H
            For x = 0 To W
                
                '-- Get pre-calculated palette index
                nEntry = RGB4096(Spd16(Bits(x, y).R), Spd16(Bits(x, y).G), Spd16(Bits(x, y).B))
                
                '-- Set 1/4/8 bpp palette indexes
                Select Case bpp
                    Case 8 '-- 256 colors
                        m_PalID(x, y) = nEntry
                    Case 4 '-- 16 colors
                        If (x Mod 2 = 0) Then
                            m_PalID(x \ 2, y) = (m_PalID(x \ 2, y) And &HF) Or nEntry * 16
                          Else
                            m_PalID(x \ 2, y) = (m_PalID(x \ 2, y) And &HF0) Or nEntry
                        End If
                    Case 1 '-- 2 colors
                        If (nEntry = 0) Then
                            m_PalID(x \ 8, y) = m_PalID(x \ 8, y) And Not m_Pow2(7 - (x Mod 8))
                          Else
                            m_PalID(x \ 8, y) = m_PalID(x \ 8, y) Or m_Pow2(7 - (x Mod 8))
                        End If
                End Select
                
                '-- Diffuse error
                If (Diffuse) Then
                    
                    With Bits(x, y)
                        errB = CLng(.B) - tmpEntry(nEntry).B
                        errG = CLng(.G) - tmpEntry(nEntry).G
                        errR = CLng(.R) - tmpEntry(nEntry).R
                    End With
                    Bits(x, y) = tmpEntry(nEntry)
                    
                    '-- Floyd-Steinberg error diffusion...
                    If (Abs(errB) + Abs(errG) + Abs(errR) > 3) Then
                        If (x < W) Then
                            xIn = x + 1
                            With Bits(xIn, y)
                                newB = .B + (7 * errB) \ 16
                                newG = .G + (7 * errG) \ 16
                                newR = .R + (7 * errR) \ 16
                                If (newB < 0) Then newB = 0 Else If (newB > 255) Then newB = 255
                                If (newG < 0) Then newG = 0 Else If (newG > 255) Then newG = 255
                                If (newR < 0) Then newR = 0 Else If (newR > 255) Then newR = 255
                                .B = newB
                                .G = newG
                                .R = newR
                            End With
                        End If
                        If (y < H) Then
                            yIn = y + 1
                            With Bits(x, yIn)
                                newB = .B + (5 * errB) \ 16
                                newG = .G + (5 * errG) \ 16
                                newR = .R + (5 * errR) \ 16
                                If (newB < 0) Then newB = 0 Else If (newB > 255) Then newB = 255
                                If (newG < 0) Then newG = 0 Else If (newG > 255) Then newG = 255
                                If (newR < 0) Then newR = 0 Else If (newR > 255) Then newR = 255
                                .B = newB
                                .G = newG
                                .R = newR
                            End With
                            If (x < W) Then
                                xIn = x + 1
                                With Bits(xIn, yIn)
                                    newB = .B + errB \ 16
                                    newG = .G + errG \ 16
                                    newR = .R + errR \ 16
                                    If (newB < 0) Then newB = 0 Else If (newB > 255) Then newB = 255
                                    If (newG < 0) Then newG = 0 Else If (newG > 255) Then newG = 255
                                    If (newR < 0) Then newR = 0 Else If (newR > 255) Then newR = 255
                                    .B = newB
                                    .G = newG
                                    .R = newR
                                End With
                            End If
                            If (x > 0) Then
                                xIn = x - 1
                                With Bits(xIn, yIn)
                                    newB = .B + (3 * errB) \ 16
                                    newG = .G + (3 * errG) \ 16
                                    newR = .R + (3 * errR) \ 16
                                    If (newB < 0) Then newB = 0 Else If (newB > 255) Then newB = 255
                                    If (newG < 0) Then newG = 0 Else If (newG > 255) Then newG = 255
                                    If (newR < 0) Then newR = 0 Else If (newR > 255) Then newR = 255
                                    .B = newB
                                    .G = newG
                                    .R = newR
                                End With
                            End If
                        End If
                    End If
                  Else
                    Bits(x, y) = tmpEntry(nEntry)
                End If
            Next x
            RaiseEvent Progress(y)
        Next y
        Call CopyMemory(ByVal VarPtrArray(Bits), 0&, 4)
        RaiseEvent ProgressEnd
    End If
End Sub

Public Function DitherToTrueColor(DIB As cDIB)
'-- Only for raising <ProgressEnd> event

  Dim y As Long
  Dim H As Long
   
    If (DIB.hDIB <> 0) Then
        
        H = DIB.Height
        
        For y = 0 To H
            RaiseEvent Progress(y)
        Next y
        RaiseEvent ProgressEnd
    End If
End Function

Public Function CountColors(DIB As cDIB) As Long
'-- by Robert Rayment

  Dim Bits() As RGBQUAD
  Dim tSA    As SAFEARRAY2D

  Dim W As Long, H As Long
  Dim x As Long, y As Long

  Dim BRArr(0 To 255, 0 To 255)  As Integer ' Blue-Red array
  Dim GArr(0 To 255, 0 To 32767) As Byte    ' Green array

  Dim i  As Integer ' Blue
  Dim j  As Integer ' Red
  Dim kG As Byte    ' Green
  Dim k  As Long    ' Green column counter
  Dim kk As Long    ' Column number

    If (DIB.hDIB <> 0) Then

        Call pvBuildSA(tSA, DIB)
        Call CopyMemory(ByVal VarPtrArray(Bits()), VarPtr(tSA), 4)

        W = DIB.Width - 1
        H = DIB.Height - 1

        For y = 0 To H
            For x = 0 To W

                '-- Get (Blue,Red) coord.
                i = Bits(x, y).B
                j = Bits(x, y).R
                kG = Bits(x, y).G

                If (BRArr(i, j) = 0) Then
                    k = k + 1
                    If (k > 32767) Then ' Rare: give warning message
                        Exit Function   ' or increase GArr size,
                    End If              ' or Redim Preserve which will be slower.

                    '-- Flag BR with column number
                    BRArr(i, j) = k

                    '-- Flag green (0-255) in column k
                    GArr(kG, k) = 1               ' First entry into green column
                    CountColors = CountColors + 1 ' Count unique colors

                  Else
                    kk = BRArr(i, j)
                    '-- Flag green (0-255) in column kk if not flagged
                    If (GArr(kG, kk) = 0) Then
                        GArr(kG, kk) = 1
                        CountColors = CountColors + 1 ' Count unique colors
                    End If
                End If
            Next x
        Next y
        Call CopyMemory(ByVal VarPtrArray(Bits), 0&, 4)
    End If
End Function

'========================================================================================
' Properties (R.O.)
'========================================================================================

Public Property Get PalID() As Byte()
    PalID = m_PalID()
End Property

Public Property Get PalIDPtr() As Long
    PalIDPtr = VarPtr(m_PalID(0, 0))
End Property

'========================================================================================
' Private
'========================================================================================

Private Sub pvBuildSA(tSA As SAFEARRAY2D, DIB As cDIB)
    With tSA
        .cbElements = IIf(App.LogMode = 1, 1, 4)
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = DIB.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = DIB.Width
        .pvData = DIB.lpBits
    End With
End Sub

