VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDIBSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'================================================
' Class:         cDIBSave.cls (1, 4, 8, 24 bpp)
' Author:        Carles P.V.
' Dependencies:  cDIB.cls, cDIBPal.cls, cDIBDither.cls
' Last revision: 2003.11.02
'================================================

Option Explicit

'-- API:

Private Type BITMAPFILEHEADER
    bfType      As Integer
    bfSize      As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits   As Long
End Type

Private Type BITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type

Private Const DIB_RGB_COLORS As Long = 0

Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpbi As BITMAPINFOHEADER, ByVal wUsage As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal ByteLength As Long)

'========================================================================================
' Methods
'========================================================================================

Public Sub Save_BMP(ByVal sFilename As String, DIB As cDIB, Pal As cDIBPal, Dither As cDIBDither, ByVal bpp As Byte)
  
  Dim hFile   As Integer

  Dim tBFH    As BITMAPFILEHEADER ' 14 bytes
  Dim tBIH    As BITMAPINFOHEADER ' 40 bytes
  Dim aBI()   As Byte
  Dim aData() As Byte
  
  Dim lIdx    As Long
  Dim Walg    As Long
 
    If (DIB.hDIB <> 0) Then
        
        '-- Bits data size
        Walg = ((DIB.Width * bpp + 31) \ 32) * 4
        ReDim aData(Walg - 1, DIB.Height - 1)
        
        '-- Build file header
        With tBFH
            .bfType = &H4D42
            .bfSize = Len(tBIH) + IIf(bpp <= 8, 4 * 2 ^ bpp, 0) + Walg * DIB.Height
            .bfOffBits = Len(tBFH) + Len(tBIH) + IIf(bpp <= 8, 4 * 2 ^ bpp, 0)
        End With
        '-- Build image header
        With tBIH
            .biSize = 40
            .biPlanes = 1
            .biBitCount = bpp
            .biWidth = DIB.Width
            .biHeight = DIB.Height
        End With
    
        '-- Build <BITMAPINFO> array
        ReDim aBI(1 To 40 + IIf(bpp <= 8, 4 * 2 ^ bpp, 0))
        Call CopyMemory(aBI(1), tBIH, 40)  ' Header
        If (bpp <= 8) Then                 ' Palette table/Indexes
            For lIdx = 0 To 2 ^ bpp - 1
                aBI(lIdx * 4 + 41) = Pal.rgbB(lIdx)
                aBI(lIdx * 4 + 42) = Pal.rgbG(lIdx)
                aBI(lIdx * 4 + 43) = Pal.rgbR(lIdx)
            Next lIdx
        End If

        '-- Translate to new color depth
        Select Case bpp
            Case 1, 4, 8 '-- Palette based (Get palette indexes)
                aData() = Dither.PalID()
            Case 24      '-- RGB based (Get RGB bits)
                Call GetDIBits(DIB.hDC, DIB.hDIB, 0, DIB.Height, aData(0, 0), tBIH, DIB_RGB_COLORS)
        End Select

        '-- Kill previous
        On Error Resume Next
           Kill sFilename
        On Error GoTo 0
        '-- Write file
        hFile = FreeFile
        Open sFilename For Binary Access Write As #hFile
          Put #hFile, , tBFH
          Put #hFile, , aBI()
          Put #hFile, , aData()
        Close #hFile
    End If
End Sub
