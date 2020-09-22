VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDIBPal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'================================================
' Class:         cDIBPal.cls
' Author:        From cPalette.cls
'                Copyright © 1999 Steve McMahon
' Dependencies:  cDIB.cls
' Last revision: 2003.11.02
'================================================
'
' cDIBPal:
' 1, 4 and 8 bpp Palette storage/processing
'
' - The Octree Colour Quantisation Code (CreateOptimal) was written by
'   Brian Schimpf Copyright © 1999 Brian Schimpf
'   Notes: + Added: RGB channels weigth coeffs.
'          + Modif: Forced black entry as first entry.

Option Explicit

'-- API:

Private Type RGBQUAD
    B As Byte
    G As Byte
    R As Byte
    A As Byte
End Type

Private Type PALETTEENTRY
    peR     As Byte
    peG     As Byte
    peB     As Byte
    peFlags As Byte
End Type

Private Type LOGPALETTE002
    palVersion       As Integer
    palNumEntries    As Integer
    palPalEntry(1)   As PALETTEENTRY
End Type

Private Type LOGPALETTE016
    palVersion       As Integer
    palNumEntries    As Integer
    palPalEntry(15)  As PALETTEENTRY
End Type

Private Type LOGPALETTE256
    palVersion       As Integer
    palNumEntries    As Integer
    palPalEntry(255) As PALETTEENTRY
End Type

Private Type tNode                ' OCT-TREE node struct.
    bIsLeaf            As Boolean ' Leaf flag
    bAddedReduce       As Boolean ' Linked list flag
    vR                 As Long    ' Red Value
    vG                 As Long    ' Green Value
    vB                 As Long    ' Blue Value
    cClrs              As Long    ' Pixel count
    iChildren(1, 1, 1) As Long    ' Child pointers
    iNext              As Long    ' Next reducable node
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

Private Declare Function CreatePalette Lib "gdi32" (lpLogPalette As Any) As Long
Private Declare Function GetNearestPaletteIndex Lib "gdi32" (ByVal hPalette As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function VarPtrArray Lib "msvbvm50" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal ByteLength As Long)


'-- Private Variables:

Private m_Count    As Long
Private m_hPal     As Long
Private m_tPal()   As RGBQUAD

Private logPal002  As LOGPALETTE002
Private logPal016  As LOGPALETTE016
Private logPal256  As LOGPALETTE256

Private aNodes()   As tNode
Private cNodes     As Long
Private nDepth     As Byte
Private TopGarbage As Long
Private cClr       As Long
Private aReduce()  As Long
Private wR         As Single
Private wG         As Single
Private wB         As Single
   
'========================================================================================
' Class
'========================================================================================

Private Sub Class_Initialize()
    m_hPal = 0
    m_Count = 0
End Sub

Private Sub Class_Terminate()
    Call Me.Clear
End Sub

'========================================================================================
' Methods
'========================================================================================

Public Function InitializePalette(ByVal nEntries As Integer) As Long

    '-- Resize palette
    m_Count = nEntries
    ReDim m_tPal(nEntries - 1) As RGBQUAD
    
    '-- Return pointer
    InitializePalette = VarPtr(m_tPal(0))
End Function

Public Function PalettePtr() As Long
    PalettePtr = VarPtr(m_tPal(0))
End Function

'//

Public Sub Clear()

    '-- Clear palette array
    Erase m_tPal
    m_Count = 0
    
    '-- Clear logical palette
    If (m_hPal <> 0) Then
        Call DeleteObject(m_hPal)
        m_hPal = 0
    End If
End Sub

Public Sub ClosestColor(ByVal R As Byte, ByVal G As Byte, ByVal B As Byte, cR As Byte, cG As Byte, cb As Byte, PalEntry As Byte)
    
    PalEntry = GetNearestPaletteIndex(m_hPal, RGB(B, G, R))
    With m_tPal(PalEntry)
        cR = .R
        cG = .G
        cb = .B
    End With
End Sub

Public Function IsGreyScale() As Boolean
    
  Dim lIdx As Long
    
    '-- R = G = B
    IsGreyScale = (m_Count > 0)
    For lIdx = 0 To m_Count - 1
        With m_tPal(lIdx)
            If (.R <> .G Or .R <> .B Or .G <> .B) Then
                IsGreyScale = False
                Exit For
            End If
        End With
    Next lIdx
End Function

'//

Public Sub CreateBlackAndWhite()

    '-- Black and White:
    m_Count = 2
    ReDim m_tPal(1) As RGBQUAD
    
    With m_tPal(1)
        .R = &HFF
        .G = &HFF
        .B = &HFF
    End With
    
    Call pvBuildLogicalPalette
End Sub

Public Sub CreateEGA()

    '-- Standard EGA style 16 colour palette:
    m_Count = 16
    ReDim m_tPal(15) As RGBQUAD
    
    Call pv16ColourLow8(0)
    Call pv16ColourHigh8(8)
    
    Call pvBuildLogicalPalette
End Sub

Public Sub CreateHalftone()

  Dim lIdx As Long
  Dim R As Long, rA As Long
  Dim G As Long, gA As Long
  Dim B As Long, bA As Long

    '-- Halftone-216 (6x6x6) colour palette
    m_Count = 256
    ReDim m_tPal(255) As RGBQUAD
   
    For B = 0 To &H100 Step &H40
        If (B = &H100) Then
            bA = B - 1
          Else
            bA = B
        End If
        For G = 0 To &H100 Step &H40
            If (G = &H100) Then
                gA = G - 1
              Else
                gA = G
            End If
            For R = 0 To &H100 Step &H40
                If (R = &H100) Then
                    rA = R - 1
                  Else
                    rA = R
                End If
                With m_tPal(lIdx)
                    .R = rA
                    .G = gA
                    .B = bA
                End With
                lIdx = lIdx + 1
            Next R
        Next G
    Next B
    
    Call pvBuildLogicalPalette
End Sub

Public Sub CreateWebsafe()

  Dim lIdx As Long
  Dim R As Long, G As Long, B As Long
  Dim L As Long, i As Long

    '-- Websafe (IE 216 colour) palette
    m_Count = 256
    ReDim m_tPal(255) As RGBQUAD
    
    Call pv16ColourLow8(0)
    Call pv16ColourHigh8(248)
    lIdx = 8
    
    For B = 0 To &HFF Step &H33
        For G = 0 To &HFF Step &H33
            For R = 0 To &HFF Step &H33
            
                '-- Ignore if the output is any combination of 0 and FF
                L = R + G + B
                If (L = 0 Or L = &H2FD) Then
                    '-- ignore
                  ElseIf (L = &H1FE And (R = 0 Or G = 0 Or B = 0)) Then
                    '-- ignore
                  ElseIf (L = &HFF And ((R = 0 And G = 0) Or (R = 0 And B = 0) Or (G = 0 And B = 0))) Then
                    '-- ignore
                  Else
                    '-- add
                    With m_tPal(lIdx)
                        .R = R
                        .G = G
                        .B = B
                    End With
                    lIdx = lIdx + 1
                End If
            Next R
        Next G
    Next B
    '-- Fill the remain entries with gray shades:
    R = 8
    G = 8
    B = 8
    For lIdx = 217 To 247
        With m_tPal(lIdx)
            .R = R: R = R + 8
            .G = G: G = G + 8
            .B = B: B = B + 8
        End With
    Next lIdx
    
    Call pvBuildLogicalPalette
End Sub

Public Sub CreateGreys_016()

  Dim lIdx As Long
  
    m_Count = 16
    ReDim m_tPal(15) As RGBQUAD
    
    For lIdx = 0 To 15
        With m_tPal(lIdx)
            .R = lIdx * 17
            .G = lIdx * 17
            .B = lIdx * 17
        End With
    Next lIdx
    
    Call pvBuildLogicalPalette
End Sub

Public Sub CreateGreys_256()

  Dim lIdx As Long
  
    m_Count = 256
    ReDim m_tPal(255) As RGBQUAD
    
    For lIdx = 0 To 255
        With m_tPal(lIdx)
            .R = lIdx
            .G = lIdx
            .B = lIdx
        End With
    Next lIdx
    
    Call pvBuildLogicalPalette
End Sub

Public Sub CreateFastOptimal_016(DIB As cDIB)
'-- Not recommended

  Dim DIB04x04 As New cDIB
 
    '-- Resample DIB to 4x4
    Call DIB04x04.Create(4, 4)
    Call DIB04x04.LoadDIBBlt(DIB)
    
    '-- Bits -> Palette
    m_Count = 16
    ReDim m_tPal(15) As RGBQUAD
    Call CopyMemory(m_tPal(0), ByVal DIB04x04.lpBits, 64)
    
    Call pvBuildLogicalPalette
End Sub

Public Sub CreateFastOptimal_256(DIB As cDIB)
'-- Not recommended

  Dim DIB16x16 As New cDIB
    
    '-- Resample DIB to 16x16
    Call DIB16x16.Create(16, 16)
    Call DIB16x16.LoadDIBBlt(DIB)
    
    '-- Bits -> Palette
    m_Count = 256
    ReDim m_tPal(255) As RGBQUAD
    Call CopyMemory(m_tPal(0), ByVal DIB16x16.lpBits, 1024)
    
    Call pvBuildLogicalPalette
End Sub

'***************************************************************
' CreateOptimal: Creates an optimal palette with the
' specified number of colors using octree quantisation.
'
' Inputs:
'   - cDIB (DIB section for which palette is made)
'   - nLevels (Number of levels of color to create)
'   - PalSize (Size of palette to make)
'   - RWeight, GWeight, BWeight (Channel wheights [0,1])
'***************************************************************
Public Sub CreateOptimal( _
           DIB As cDIB, _
           ByVal PalSize As Long, _
           ByVal nLevels As Long, _
           Optional ByVal RWeight As Single = 1, _
           Optional ByVal GWeight As Single = 1, _
           Optional ByVal BWeight As Single = 1)
    
  Dim Bits() As RGBQUAD
  Dim tSA    As SAFEARRAY2D
  
  Dim x As Long, xMax As Long
  Dim y As Long, yMax As Long
  Dim swpNode As tNode
  Dim pPal    As Long
  
    '--  R-G-B weights
    wR = RWeight
    wG = GWeight
    wB = BWeight
    
    '-- Map DIB bits
    Call pvBuildSA(tSA, DIB)
    Call CopyMemory(ByVal VarPtrArray(Bits()), VarPtr(tSA), 4)
    
    '-- Allocates initial storage
    ReDim aNodes(1 To 50) As tNode
    ReDim aReduce(1 To 8) As Long
    nDepth = nLevels
    cNodes = 1
    TopGarbage = 0
    cClr = 0
    
    '-- Allow fisrt entry for black
    PalSize = PalSize - 1
    
    xMax = DIB.Width - 1
    yMax = DIB.Height - 1
    
    For y = 0 To yMax
        For x = 0 To xMax

            '-- Adds the current pixel to the color octree
            Call pvAddClr(1, 1, 0, 255, 0, 255, 0, 255, Bits(x, y).R, Bits(x, y).G, Bits(x, y).B)

            '-- Combine the levels to get down to desired palette size
            Do While (cClr > PalSize)
                If (pvCombineNodes = False) Then Exit Do
            Loop
        Next x
    Next y
    Call CopyMemory(ByVal VarPtrArray(Bits), 0&, 4)
    
    '-- Sort by frequency
    For x = 1 To UBound(aNodes) - 1
        For y = x To UBound(aNodes)
            If (aNodes(y).cClrs > aNodes(x).cClrs) Then
                swpNode = aNodes(x)
                aNodes(x) = aNodes(y)
                aNodes(y) = swpNode
            End If
        Next y
    Next x
    
    '-- Size palette
    m_Count = PalSize + 1: ReDim m_tPal(m_Count - 1)
    '-- Go through octree and extract colors
    pPal = 1
    For y = 1 To UBound(aNodes)
        If (aNodes(y).bIsLeaf) Then
            With aNodes(y)
                m_tPal(pPal).R = .vR / .cClrs
                m_tPal(pPal).G = .vG / .cClrs
                m_tPal(pPal).B = .vB / .cClrs
                pPal = pPal + 1
            End With
        End If
    Next y
    
    '-- Create logical palette
    Call pvBuildLogicalPalette
End Sub

'***************************************************************
' <Recursive>
' pvAddClr: Adds a color to the OctTree palette.
' Will call itself if not in correct level.
'
' Inputs:
'  - iBranch (Branch to look down)
'  - nLevel (Current level (depth) in tree)
'  - vMin(R, G, B) (The minimum branch value)
'  - vMax(R, G, B) (The maximum branch value)
'  - R, G, B (The Red, Green, and Blue color components)
'***************************************************************
Private Sub pvAddClr( _
        ByVal iBranch As Long, _
        ByVal nLevel As Long, _
        ByVal vMinR As Byte, ByVal vMaxR As Byte, _
        ByVal vMinG As Byte, ByVal vMaxG As Byte, _
        ByVal vMinB As Byte, ByVal vMaxB As Byte, _
        ByVal R As Byte, ByVal G As Byte, ByVal B As Byte)

  Dim iR As Byte, iG As Byte, iB As Byte
  Dim vMid As Byte, iIndex As Long
    
    '-- Find mid values for colors and decide which path to take.
    '-- Also update max and min values for later call to self.
    
    vMid = vMinR / 2 + vMaxR / 2
    If (wR * R > vMid) Then iR = 1: vMinR = vMid Else iR = 0: vMaxR = vMid

    vMid = vMinG / 2 + vMaxG / 2
    If (wG * G > vMid) Then iG = 1: vMinG = vMid Else iG = 0: vMaxG = vMid

    vMid = vMinB / 2 + vMaxB / 2
    If (wB * B > vMid) Then iB = 1: vMinB = vMid Else iB = 0: vMaxB = vMid
    
    '-- If no child here then...
    If (aNodes(iBranch).iChildren(iR, iG, iB) = 0) Then
        '-- Get a new node index
        iIndex = pvGetFreeNode
        aNodes(iBranch).iChildren(iR, iG, iB) = iIndex
        aNodes(iBranch).cClrs = aNodes(iBranch).cClrs + 1
        '-- Clear/set data
        With aNodes(iIndex)
            .bIsLeaf = (nLevel = nDepth)
            .iNext = 0
            .cClrs = 0
            .vR = 0
            .vG = 0
            .vB = 0
        End With
      Else
        '-- Has a child here
        iIndex = aNodes(iBranch).iChildren(iR, iG, iB)
    End If
 
    '-- If it is a leaf
    If (aNodes(iIndex).bIsLeaf) Then
        With aNodes(iIndex)
            If (.cClrs = 0) Then cClr = cClr + 1
            .cClrs = .cClrs + 1
            .vR = .vR + R
            .vG = .vG + G
            .vB = .vB + B
        End With
      Else
        With aNodes(iIndex)
            '-- If 2 or more colors, add to reducable aNodes list
            If (.bAddedReduce = 0) Then
                .iNext = aReduce(nLevel)
                 aReduce(nLevel) = iIndex
                .bAddedReduce = True
            End If
        End With
        '-- Search a level deeper
        Call pvAddClr(iIndex, nLevel + 1, vMinR, vMaxR, vMinG, vMaxG, vMinB, vMaxB, R, G, B)
   End If
End Sub

'***************************************************************
' pvCombineNodes: Combines octree aNodes to reduce the count of
' colors. Combines all children of a leaf into itself.
'***************************************************************
Private Function pvCombineNodes() As Boolean

  Dim i As Long, iIndex As Long
  Dim iR As Byte, iG As Byte, iB As Byte
  Dim nR As Long, nG As Long, nB As Long, nPixel As Long

    '-- Find deepest reducable level
    For i = nDepth To 1 Step -1
        If (aReduce(i) <> 0) Then Exit For
    Next i

    If (i = 0) Then Exit Function
    iIndex = aReduce(i)
    aReduce(i) = aNodes(iIndex).iNext

    For i = 0 To 7
        If (i And 1) = 1 Then iR = 1 Else iR = 0
        If (i And 2) = 2 Then iG = 1 Else iG = 0
        If (i And 4) = 4 Then iB = 1 Else iB = 0
        
        '-- If there is a child
        If (aNodes(iIndex).iChildren(iR, iG, iB) <> 0) Then
            With aNodes(aNodes(iIndex).iChildren(iR, iG, iB))
                '-- Add red, green, blue, and pixel count to running total
                nR = nR + .vR
                nG = nG + .vG
                nB = nB + .vB
                nPixel = nPixel + .cClrs
                '-- Free the node
                Call pvFreeNode(aNodes(iIndex).iChildren(iR, iG, iB))
                cClr = cClr - 1
            End With
            '-- Clear the link
            aNodes(iIndex).iChildren(iR, iG, iB) = 0
        End If
    Next i
    cClr = cClr + 1

    '-- Set the new node data
    With aNodes(iIndex)
        .cClrs = nPixel
        .bIsLeaf = True
        .vR = nR
        .vG = nG
        .vB = nB
    End With
    pvCombineNodes = True
End Function

'***************************************************************
' pvFreeNode: Puts a node on the top of the garbage list.
' Inputs:
'  - iNode
'  - Index of node to free
'***************************************************************
Private Sub pvFreeNode(ByVal iNode As Long)

    aNodes(iNode).iNext = TopGarbage
    TopGarbage = iNode
    aNodes(iNode).bIsLeaf = False ' Necessary for final loop through
    aNodes(iNode).bAddedReduce = False
    cNodes = cNodes - 1
End Sub

'***************************************************************
' pvGetFreeNode: Gets a new node index from the trash list or the
' end of the list. Clears child pointers.
' Outputs:
'  - Node index
'***************************************************************
Private Function pvGetFreeNode() As Long

  Dim i  As Long
  Dim iR As Byte
  Dim iG As Byte
  Dim iB As Byte
  
    cNodes = cNodes + 1
    If (TopGarbage = 0) Then
        If (cNodes > UBound(aNodes)) Then
            i = cNodes * 1.1
            ReDim Preserve aNodes(1 To i)
        End If
        pvGetFreeNode = cNodes
      Else
        pvGetFreeNode = TopGarbage
        TopGarbage = aNodes(TopGarbage).iNext
        For i = 0 To 7
            If (i And 1) = 1 Then iR = 1 Else iR = 0
            If (i And 2) = 2 Then iG = 1 Else iG = 0
            If (i And 4) = 4 Then iB = 1 Else iB = 0
            aNodes(pvGetFreeNode).iChildren(iR, iG, iB) = 0
        Next i
    End If
End Function

'========================================================================================
' Properties (R.O.)
'========================================================================================

Public Property Get hPal() As Long
    hPal = m_hPal
End Property

Public Property Get Count() As Long
   Count = m_Count
End Property

Public Property Get rgbR(ByVal PalEntry As Byte) As Byte
   rgbR = m_tPal(PalEntry).R
End Property

Public Property Get rgbG(ByVal PalEntry As Byte) As Byte
   rgbG = m_tPal(PalEntry).G
End Property

Public Property Get rgbB(ByVal PalEntry As Byte) As Byte
   rgbB = m_tPal(PalEntry).B
End Property

'========================================================================================
' Private
'========================================================================================

Private Sub pv16ColourLow8(ByVal lStartIndex As Long)

    With m_tPal(lStartIndex + 0)
        .R = 0: .G = 0: .B = 0
    End With
    With m_tPal(lStartIndex + 1)
        .R = &H80: .G = 0: .B = 0
    End With
    With m_tPal(lStartIndex + 2)
        .R = 0: .G = &H80: .B = 0
    End With
    With m_tPal(lStartIndex + 3)
        .R = &H80: .G = &H80: .B = 0
    End With
    With m_tPal(lStartIndex + 4)
        .R = 0: .G = 0: .B = &H80
    End With
    With m_tPal(lStartIndex + 5)
        .R = &H80: .G = 0: .B = &H80
    End With
    With m_tPal(lStartIndex + 6)
        .R = 0: .G = &H80: .B = &H80
    End With
    With m_tPal(lStartIndex + 7)
        .R = &HC0: .G = &HC0: .B = &HC0
    End With
End Sub

Private Sub pv16ColourHigh8(ByVal lStartIndex As Long)
    
    With m_tPal(lStartIndex + 0)
        .R = &H80: .G = &H80: .B = &H80
    End With
    With m_tPal(lStartIndex + 1)
        .R = &HFF: .G = 0: .B = 0
    End With
    With m_tPal(lStartIndex + 2)
        .R = 0: .G = &HFF: .B = 0
    End With
    With m_tPal(lStartIndex + 3)
        .R = &HFF: .G = &HFF: .B = 0
    End With
    With m_tPal(lStartIndex + 4)
        .R = 0: .G = 0: .B = &HFF
    End With
    With m_tPal(lStartIndex + 5)
        .R = &HFF: .G = 0: .B = &HFF
    End With
    With m_tPal(lStartIndex + 6)
        .R = 0: .G = &HFF: .B = &HFF
    End With
    With m_tPal(lStartIndex + 7)
        .R = &HFF: .G = &HFF: .B = &HFF
    End With
End Sub

Private Sub pvBuildLogicalPalette()
    
    If (m_hPal <> 0) Then
        Call DeleteObject(m_hPal)
        m_hPal = 0
    End If
    
    '-- Build logical palette
    Select Case m_Count

      Case 2   '-- 2 colors
        With logPal002
            .palNumEntries = 2
            .palVersion = &H300
             Call CopyMemory(.palPalEntry(0), m_tPal(0), 8)
        End With
        m_hPal = CreatePalette(logPal002)

      Case 16  '-- 16 colors
        With logPal016
            .palNumEntries = 16
            .palVersion = &H300
             Call CopyMemory(.palPalEntry(0), m_tPal(0), 64)
        End With
        m_hPal = CreatePalette(logPal016)

      Case Else '-- 256 colors
        With logPal256
            .palNumEntries = 256
            .palVersion = &H300
             Call CopyMemory(.palPalEntry(0), m_tPal(0), 1024)
        End With
        m_hPal = CreatePalette(logPal256)
    End Select
End Sub

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
