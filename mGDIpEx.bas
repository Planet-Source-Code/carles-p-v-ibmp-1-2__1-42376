Attribute VB_Name = "mGDIpEx"
'================================================
' Module:        mGDIpEx.bas
' Author:        (*)
' Dependencies:  -
' Last revision: 2004.08.02
'================================================
'
' (*) From great stuff:
'
'     + http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=37541&lngWId=1
'       by Avery
'     + Also thanks to Vlad Vissoultchev
'
' GDI+ PNG (Load/Save) and JPEG (Save) support
' High-quality scaling
'
' In case lower versions than Windows XP, needed:
'
'     Platform SDK Redistributable: GDI+ RTM
'     http://www.microsoft.com/downloads/release.asp?releaseid=32738

Option Explicit

'-- GDI+ API:

Public Enum GpImageFormat
    [ImageGIF] = 0
    [ImageJPEG] = 1
    [ImagePNG] = 2
    [ImageTIFF] = 3
End Enum

Public Enum GpStatus
    [OK] = 0
    [GenericError] = 1
    [InvalidParameter] = 2
    [OutOfMemory] = 3
    [ObjectBusy] = 4
    [InsufficientBuffer] = 5
    [NotImplemented] = 6
    [Win32Error] = 7
    [WrongState] = 8
    [Aborted] = 9
    [FileNotFound] = 10
    [ValueOverflow ] = 11
    [AccessDenied] = 12
    [UnknownImageFormat] = 13
    [FontFamilyNotFound] = 14
    [FontStyleNotFound] = 15
    [NotTrueTypeFont] = 16
    [UnsupportedGdiplusVersion] = 17
    [GdiplusNotInitialized ] = 18
    [PropertyNotFound] = 19
    [PropertyNotSupported] = 20
End Enum

Public Type GdiplusStartupInput
    GdiplusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type

'//

Private Enum GpUnit
    [UnitWorld]
    [UnitDisplay]
    [UnitPixel]
    [UnitPoint]
    [UnitInch]
    [UnitDocument]
    [UnitMillimeter]
End Enum

Private Enum QualityMode
    [QualityModeInvalid] = -1
    [QualityModeDefault] = 0
    [QualityModeLow] = 1
    [QualityModeHigh] = 2
End Enum

Private Enum PixelOffsetMode
    [PixelOffsetModeInvalid] = -1
    [PixelOffsetModeDefault]
    [PixelOffsetModeHighSpeed]
    [PixelOffsetModeHighQuality]
    [PixelOffsetModeNone]
    [PixelOffsetModeHalf]
End Enum

Private Enum InterpolationMode
    [InterpolationModeInvalid] = [QualityModeInvalid]
    [InterpolationModeDefault] = [QualityModeDefault]
    [InterpolationModeLowQuality] = [QualityModeLow]
    [InterpolationModeHighQuality] = [QualityModeHigh]
    [InterpolationModeBilinear]
    [InterpolationModeBicubic]
    [InterpolationModeNearestNeighbor]
    [InterpolationModeHighQualityBilinear]
    [InterpolationModeHighQualityBicubic]
End Enum

Private Enum EncoderParameterValueType
    [EncoderParameterValueTypeByte] = 1
    [EncoderParameterValueTypeASCII] = 2
    [EncoderParameterValueTypeShort] = 3
    [EncoderParameterValueTypeLong] = 4
    [EncoderParameterValueTypeRational] = 5
    [EncoderParameterValueTypeLongRange] = 6
    [EncoderParameterValueTypeUndefined] = 7
    [EncoderParameterValueTypeRationalRange] = 8
End Enum

Private Enum EncoderValue
    [EncoderValueColorTypeCMYK] = 0
    [EncoderValueColorTypeYCCK] = 1
    [EncoderValueCompressionLZW] = 2
    [EncoderValueCompressionCCITT3] = 3
    [EncoderValueCompressionCCITT4] = 4
    [EncoderValueCompressionRle] = 5
    [EncoderValueCompressionNone] = 6
    [EncoderValueScanMethodInterlaced]
    [EncoderValueScanMethodNonInterlaced]
    [EncoderValueVersionGif87]
    [EncoderValueVersionGif89]
    [EncoderValueRenderProgressive]
    [EncoderValueRenderNonProgressive]
    [EncoderValueTransformRotate90]
    [EncoderValueTransformRotate180]
    [EncoderValueTransformRotate270]
    [EncoderValueTransformFlipHorizontal]
    [EncoderValueTransformFlipVertical]
    [EncoderValueMultiFrame]
    [EncoderValueLastFrame]
    [EncoderValueFlush]
    [EncoderValueFrameDimensionTime]
    [EncoderValueFrameDimensionResolution]
    [EncoderValueFrameDimensionPage]
End Enum

Private Type CLSID
    Data1         As Long
    Data2         As Integer
    Data3         As Integer
    Data4(0 To 7) As Byte
End Type

Private Type ImageCodecInfo
    ClassID           As CLSID
    FormatID          As CLSID
    CodecName         As Long
    DllName           As Long
    FormatDescription As Long
    FilenameExtension As Long
    MimeType          As Long
    Flags             As Long
    Version           As Long
    SigCount          As Long
    SigSize           As Long
    SigPattern        As Long
    SigMask           As Long
End Type

'-- Encoder Parameter structure
Private Type EncoderParameter
    GUID           As CLSID
    NumberOfValues As Long
    Type           As EncoderParameterValueType
    Value          As Long
End Type

'-- Encoder Parameters structure
Private Type EncoderParameters
    Count     As Long
    Parameter As EncoderParameter
End Type

'-- Encoder parameter sets
Private Const EncoderCompression      As String = "{E09D739D-CCD4-44EE-8EBA-3FBF8BE4FC58}"
Private Const EncoderColorDepth       As String = "{66087055-AD66-4C7C-9A18-38A2310B8337}"
Private Const EncoderScanMethod       As String = "{3A4E2661-3109-4E56-8536-42C156E7DCFA}"
Private Const EncoderVersion          As String = "{24D18C76-814A-41A4-BF53-1C219CCCF797}"
Private Const EncoderRenderMethod     As String = "{6D42C53A-229A-4825-8BB7-5C99E2B9A8B8}"
Private Const EncoderQuality          As String = "{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"
Private Const EncoderTransformation   As String = "{8D0EB2D1-A58E-4EA8-AA14-108074B7B6F9}"
Private Const EncoderLuminanceTable   As String = "{EDB33BCE-0266-4A77-B904-27216099E717}"
Private Const EncoderChrominanceTable As String = "{F2E455DC-09B3-4316-8260-676ADA32481C}"
Private Const EncoderSaveFlag         As String = "{292266FC-AC40-47BF-8CFC-A85B89A655DE}"
Private Const CodecIImageBytes        As String = "{025D1823-6C7D-447B-BBDB-A3CBC3DFA2FC}"

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

Private Type RGBQUAD
    B As Byte
    G As Byte
    R As Byte
    A As Byte
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type

Private Type PICTDESC
    Size       As Long
    Type       As Long
    hBmpOrIcon As Long
    hPal       As Long
End Type

Public Declare Function GdiplusStartup Lib "gdiplus" (Token As Long, InputBuf As GdiplusStartupInput, Optional ByVal OutputBuf As Long = 0) As GpStatus
Public Declare Function GdiplusShutdown Lib "gdiplus" (ByVal Token As Long) As GpStatus

Private Declare Function GdipGetImageEncodersSize Lib "gdiplus" (numEncoders As Long, Size As Long) As GpStatus
Private Declare Function GdipGetImageEncoders Lib "gdiplus" (ByVal numEncoders As Long, ByVal Size As Long, Encoders As Any) As GpStatus
Private Declare Function GdipGetImageDecodersSize Lib "gdiplus" (numDecoders As Long, Size As Long) As GpStatus
Private Declare Function GdipGetImageDecoders Lib "gdiplus" (ByVal numDecoders As Long, ByVal Size As Long, Decoders As Any) As GpStatus

Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDC As Long, hGraphics As Long) As GpStatus
Private Declare Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" (ByVal Bitmap As Long, hBmpReturn As Long, ByVal Background As Long) As GpStatus
Private Declare Function GdipCreateBitmapFromGdiDib Lib "gdiplus" (gdiBitmapInfo As BITMAPINFO, gdiBitmapData As Any, Bitmap As Long) As GpStatus
Private Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal Filename As String, hImage As Long) As GpStatus
Private Declare Function GdipSaveImageToFile Lib "gdiplus" (ByVal hImage As Long, ByVal sFilename As String, clsidEncoder As CLSID, encoderParams As Any) As GpStatus
Private Declare Function GdipSetPixelOffsetMode Lib "gdiplus" (ByVal hGraphics As Long, ByVal OffsetMode As PixelOffsetMode) As GpStatus
Private Declare Function GdipSetInterpolationMode Lib "gdiplus" (ByVal hGraphics As Long, ByVal Interpolation As InterpolationMode) As GpStatus
Private Declare Function GdipDrawImageRectRect Lib "gdiplus" (ByVal hGraphics As Long, ByVal hImage As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal srcUnit As GpUnit, Optional ByVal imageAttributes As Long = 0, Optional ByVal Callback As Long = 0, Optional ByVal CallbackData As Long = 0) As GpStatus
Private Declare Function GdipDrawImageRectRectI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hImage As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal srcUnit As GpUnit, Optional ByVal imageAttributes As Long = 0, Optional ByVal Callback As Long = 0, Optional ByVal CallbackData As Long = 0) As GpStatus
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal hImage As Long) As GpStatus
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal hGraphics As Long) As GpStatus

Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpszProgID As Long, pCLSID As CLSID) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal psString As Any) As Long
Private Declare Function lstrlenA Lib "kernel32" (ByVal psString As Any) As Long

Private Declare Function OleCreatePictureIndirect Lib "olepro32" (lpPictDesc As PICTDESC, riid As Any, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Declare Function CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal cb As Long) As Long

'//

Public Function LoadPictureEx(ByVal sFilename As String) As StdPicture

  Dim gplRet        As Long
  Dim hImg          As Long
  Dim hBmp          As Long
  Dim uPictDesc     As PICTDESC
  Dim aGuid(0 To 3) As Long
    
    '-- Load image
    gplRet = GdipLoadImageFromFile(StrConv(sFilename, vbUnicode), hImg)
    
    '-- Create bitmap
    gplRet = GdipCreateHBITMAPFromBitmap(hImg, hBmp, vbBlack)
    
    '-- Free image
    gplRet = GdipDisposeImage(hImg)
    
    If (gplRet = [OK]) Then
    
        '-- Fill struct
        With uPictDesc
            .Size = Len(uPictDesc)
            .Type = vbPicTypeBitmap
            .hBmpOrIcon = hBmp
            .hPal = 0
        End With
        
        '-- Fill in magic IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
        aGuid(0) = &H7BF80980
        aGuid(1) = &H101ABF32
        aGuid(2) = &HAA00BB8B
        aGuid(3) = &HAB0C3000
        
        '-- Create picture from bitmap handle
        OleCreatePictureIndirect uPictDesc, aGuid(0), -1, LoadPictureEx
    End If
End Function

Public Function SaveDIB(DIB As cDIB, _
                        ByVal sFilename As String, _
                        ByVal lEncoder As GpImageFormat, _
                        Optional ByVal JPEG_Quality As Long = 90, _
                        Optional ByVal TIFF_ColorDepth As Long = 24, _
                        Optional ByVal TIFF_Compression As Long = [EncoderValueCompressionNone] _
                        ) As Boolean
    
  Dim gplRet       As Long
  Dim uInfo        As BITMAPINFO
  Dim hImg         As Long
  Dim uEncCLSID    As CLSID
  Dim uEncParams   As EncoderParameters
  Dim aEncParams() As Byte
  
    '-- Prepare struct
    With uInfo.bmiHeader
        .biSize = Len(uInfo.bmiHeader)
        .biPlanes = 1
        .biBitCount = 32
        .biWidth = DIB.Width
        .biHeight = DIB.Height
    End With
    
    '-- Create bitmap
    gplRet = GdipCreateBitmapFromGdiDib(uInfo, ByVal DIB.lpBits, hImg)
    
    '-- Get image encoder
    Select Case lEncoder
    
        Case [ImageGIF]
            '-- GIF encoder
            Call pvGetEncoderClsID("image/gif", uEncCLSID)
            ReDim aEncParams(1 To Len(uEncParams))
      
        Case [ImageJPEG]
            '-- JPEG encoder
            Call pvGetEncoderClsID("image/jpeg", uEncCLSID)
            '-- Set encoder params. (Quality)
            uEncParams.Count = 1: ReDim aEncParams(1 To Len(uEncParams))
            With uEncParams.Parameter
                .NumberOfValues = 1
                .Type = [EncoderParameterValueTypeLong]
                .GUID = pvDEFINE_GUID(EncoderQuality)
                .Value = VarPtr(JPEG_Quality)
            End With
            Call CopyMemory(aEncParams(1), uEncParams, Len(uEncParams))
        
        Case [ImagePNG]
            '-- PNG encoder
            Call pvGetEncoderClsID("image/png", uEncCLSID)
            ReDim aEncParams(1 To Len(uEncParams))
            
        Case [ImageTIFF]
            '-- TIFF encoder
            Call pvGetEncoderClsID("image/tiff", uEncCLSID)
            '-- Set encoder params. (Compression/Color depth)
            uEncParams.Count = 2: ReDim aEncParams(1 To Len(uEncParams) + Len(uEncParams.Parameter))
            With uEncParams.Parameter
                .NumberOfValues = 1
                .Type = [EncoderParameterValueTypeLong]
                .GUID = pvDEFINE_GUID(EncoderCompression)
                .Value = VarPtr(TIFF_Compression)
            End With
            Call CopyMemory(aEncParams(1), uEncParams, Len(uEncParams))
            With uEncParams.Parameter
                .NumberOfValues = 1
                .Type = [EncoderParameterValueTypeLong]
                .GUID = pvDEFINE_GUID(EncoderColorDepth)
                .Value = VarPtr(TIFF_ColorDepth)
            End With
            Call CopyMemory(aEncParams(Len(uEncParams) + 1), uEncParams.Parameter, Len(uEncParams.Parameter))
    End Select
    
    '-- Kill previous
    On Error Resume Next
       Kill sFilename
    On Error GoTo 0
    
    '-- Encode
    gplRet = GdipSaveImageToFile(hImg, StrConv(sFilename, vbUnicode), uEncCLSID, aEncParams(1))

    '-- Free image
    gplRet = GdipDisposeImage(hImg)

    '-- Success
    SaveDIB = (gplRet = [OK])
End Function

Public Function ScaleDIB(DIB As cDIB, _
                         ByVal NewWidth As Long, _
                         ByVal NewHeight As Long, _
                         Optional ByVal HighQuality As Boolean = False _
                         ) As Boolean
   
  Dim gplRet    As Long
  Dim sDIB      As New cDIB
  Dim uInfo     As BITMAPINFO
  Dim hGraphics As Long
  Dim hImg      As Long
  
  Dim OldWidth  As Long
  Dim OldHeight As Long
   
    If (DIB.hDIB <> 0) Then
   
        '-- Buffer DIB
        If (sDIB.Create(NewWidth, NewHeight)) Then
        
            '-- Get source dimensions
            OldWidth = DIB.Width
            OldHeight = DIB.Height
        
            '-- Create 'surface'
            gplRet = GdipCreateFromHDC(sDIB.hDC, hGraphics)
            
            '-- Create bitmap
            With uInfo.bmiHeader
                .biSize = Len(uInfo.bmiHeader)
                .biPlanes = 1
                .biBitCount = 32
                .biWidth = DIB.Width
                .biHeight = DIB.Height
            End With
            gplRet = GdipCreateBitmapFromGdiDib(uInfo, ByVal DIB.lpBits, hImg)
            
            '-- Scale
            If (HighQuality) Then
                gplRet = GdipSetInterpolationMode(hGraphics, [InterpolationModeHighQualityBicubic])
              Else
                gplRet = GdipSetInterpolationMode(hGraphics, [InterpolationModeNearestNeighbor])
            End If
            gplRet = GdipSetPixelOffsetMode(hGraphics, [PixelOffsetModeHighQuality])
            gplRet = GdipDrawImageRectRectI(hGraphics, hImg, 0, 0, NewWidth, NewHeight, 0, 0, OldWidth, OldHeight, [UnitPixel])
                
            '-- Clean up
            gplRet = GdipDisposeImage(hImg)
            gplRet = GdipDeleteGraphics(hGraphics)
               
            '-- Success
            If (gplRet = [OK]) Then
           
                '-- Get from Buffer
                If (DIB.Create(NewWidth, NewHeight)) Then
                    DIB.LoadBlt sDIB.hDC
                    ScaleDIB = True
                End If
            End If
        End If
    End If
End Function

'========================================================================================
' Private
'========================================================================================

Private Function pvGetEncoderClsID(strMimeType As String, ClassID As CLSID) As Long

  Dim Num      As Long
  Dim Size     As Long
  Dim lIdx     As Long
  Dim ICI()    As ImageCodecInfo
  Dim Buffer() As Byte
    
    pvGetEncoderClsID = -1 ' Failure flag
    
    '-- Get the encoder array size
    Call GdipGetImageEncodersSize(Num, Size)
    If (Size = 0) Then Exit Function ' Failed!
    
    '-- Allocate room for the arrays dynamically
    ReDim ICI(1 To Num) As ImageCodecInfo
    ReDim Buffer(1 To Size) As Byte
    
    '-- Get the array and string data
    Call GdipGetImageEncoders(Num, Size, Buffer(1))
    '-- Copy the class headers
    Call CopyMemory(ICI(1), Buffer(1), (Len(ICI(1)) * Num))
    
    '-- Loop through all the codecs
    For lIdx = 1 To Num
        '-- Must convert the pointer into a usable string
        If (StrComp(pvPtrToStrW(ICI(lIdx).MimeType), strMimeType, vbTextCompare) = 0) Then
            ClassID = ICI(lIdx).ClassID ' Save the Class ID
            pvGetEncoderClsID = lIdx      ' Return the index number for success
            Exit For
        End If
    Next lIdx
    '-- Free the memory
    Erase ICI
    Erase Buffer
End Function

Private Function pvGetDecoderClsID(strMimeType As String, ClassID As CLSID) As Long

  Dim Num      As Long
  Dim Size     As Long
  Dim lIdx     As Long
  Dim ICI()    As ImageCodecInfo
  Dim Buffer() As Byte
    
    pvGetDecoderClsID = -1 'Failure flag
    
    '-- Get the encoder array size
    Call GdipGetImageDecodersSize(Num, Size)
    If (Size = 0) Then Exit Function ' Failed!
    
    '-- Allocate room for the arrays dynamically
    ReDim ICI(1 To Num) As ImageCodecInfo
    ReDim Buffer(1 To Size) As Byte
    
    '-- Get the array and string data
    Call GdipGetImageDecoders(Num, Size, Buffer(1))
    '-- Copy the class headers
    Call CopyMemory(ICI(1), Buffer(1), (Len(ICI(1)) * Num))
    
    '-- Loop through all the codecs
    For lIdx = 1 To Num
        '-- Must convert the pointer into a usable string
        If (StrComp(pvPtrToStrW(ICI(lIdx).MimeType), strMimeType, vbTextCompare) = 0) Then
            ClassID = ICI(lIdx).ClassID ' Save the Class ID
            pvGetDecoderClsID = lIdx      ' Return the index number for success
            Exit For
        End If
    Next lIdx
    '-- Free the memory
    Erase ICI
    Erase Buffer
End Function

Private Function pvDEFINE_GUID(ByVal sGuid As String) As CLSID
'-- Courtesy of: Dana Seaman
'   Helper routine to convert a CLSID(aka GUID) string to a structure
'   Example ImageFormatBMP = {B96B3CAB-0728-11D3-9D7B-0000F81EF32E}
    Call CLSIDFromString(StrPtr(sGuid), pvDEFINE_GUID)
End Function

'-- From www.mvps.org/vbnet
'   Dereferences an ANSI or Unicode string pointer
'   and returns a normal VB BSTR

Private Function pvPtrToStrW(ByVal lpsz As Long) As String
    
  Dim sOut As String
  Dim lLen As Long

    lLen = lstrlenW(lpsz)

    If (lLen > 0) Then
        sOut = StrConv(String$(lLen, vbNullChar), vbUnicode)
        Call CopyMemory(ByVal sOut, ByVal lpsz, lLen * 2)
        pvPtrToStrW = StrConv(sOut, vbFromUnicode)
    End If
End Function

Private Function pvPtrToStrA(ByVal lpsz As Long) As String
    
  Dim sOut As String
  Dim lLen As Long

    lLen = lstrlenA(lpsz)

    If (lLen > 0) Then
        sOut = String$(lLen, vbNullChar)
        Call CopyMemory(ByVal sOut, ByVal lpsz, lLen)
        pvPtrToStrA = sOut
    End If
End Function
