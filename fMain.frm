VERSION 5.00
Begin VB.Form fMain 
   Caption         =   "iBMP"
   ClientHeight    =   6060
   ClientLeft      =   2190
   ClientTop       =   4950
   ClientWidth     =   8340
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   404
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   556
   Begin iBMP.ucToolbar Quick 
      Height          =   360
      Left            =   0
      Top             =   30
      Width           =   8340
      _ExtentX        =   14711
      _ExtentY        =   635
   End
   Begin iBMP.ucInfo Info 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      Top             =   5760
      Width           =   8340
      _ExtentX        =   14711
      _ExtentY        =   529
   End
   Begin iBMP.ucCanvas Canvas 
      Height          =   5145
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   435
      Width           =   5550
      _ExtentX        =   9790
      _ExtentY        =   9075
   End
   Begin iBMP.ucProgress Progress 
      Align           =   2  'Align Bottom
      Height          =   150
      Left            =   0
      Top             =   5610
      Width           =   8340
      _ExtentX        =   14711
      _ExtentY        =   265
      BorderStyle     =   1
   End
   Begin VB.Menu mnuFileTop 
      Caption         =   "&File"
      Begin VB.Menu mnuFile 
         Caption         =   "&Open..."
         Index           =   0
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Save"
         Index           =   1
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Save &as..."
         Index           =   2
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Print..."
         Index           =   4
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuFile 
         Caption         =   "E&xit"
         Index           =   6
      End
   End
   Begin VB.Menu mnuEditTop 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEdit 
         Caption         =   "&Undo"
         Index           =   0
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Redo"
         Index           =   1
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Copy"
         Index           =   3
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Paste"
         Index           =   4
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Re&size..."
         Index           =   6
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Orien&tation..."
         Index           =   7
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Scro&ll mode"
         Checked         =   -1  'True
         Index           =   9
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Cr&op mode"
         Index           =   10
      End
   End
   Begin VB.Menu mnuZoomTop 
      Caption         =   "&Zoom"
      Begin VB.Menu mnuZoom 
         Caption         =   "Zoom &in"
         Index           =   0
      End
      Begin VB.Menu mnuZoom 
         Caption         =   "Zoom &out"
         Index           =   1
      End
      Begin VB.Menu mnuZoom 
         Caption         =   "1 : 1"
         Index           =   2
      End
      Begin VB.Menu mnuZoom 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuZoom 
         Caption         =   "&Fit mode"
         Index           =   4
      End
   End
   Begin VB.Menu mnuColorsTop 
      Caption         =   "&Colors"
      Begin VB.Menu mnuColors 
         Caption         =   "Black and White"
         Index           =   0
      End
      Begin VB.Menu mnuColors 
         Caption         =   "16 greys"
         Index           =   1
      End
      Begin VB.Menu mnuColors 
         Caption         =   "256 greys"
         Index           =   2
      End
      Begin VB.Menu mnuColors 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuColors 
         Caption         =   "2 colors"
         Index           =   4
      End
      Begin VB.Menu mnuColors 
         Caption         =   "16 colors"
         Index           =   5
      End
      Begin VB.Menu mnuColors 
         Caption         =   "256 colors"
         Index           =   6
      End
      Begin VB.Menu mnuColors 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuColors 
         Caption         =   "True color"
         Index           =   8
      End
   End
   Begin VB.Menu mnuAdjustTop 
      Caption         =   "&Adjust"
      Begin VB.Menu mnuAdjust 
         Caption         =   "&Brightness..."
         Index           =   0
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuAdjust 
         Caption         =   "&Contrast..."
         Index           =   1
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuAdjust 
         Caption         =   "&Saturation..."
         Index           =   2
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuAdjust 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuAdjust 
         Caption         =   "Filter browser..."
         Index           =   4
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu mnuFilterTop 
      Caption         =   "Fi&lter"
      Begin VB.Menu mnuFilter 
         Caption         =   "&Color"
         Index           =   0
         Begin VB.Menu mnuColorFilter 
            Caption         =   "&Greys"
            Index           =   0
         End
         Begin VB.Menu mnuColorFilter 
            Caption         =   "&Negative"
            Index           =   1
         End
         Begin VB.Menu mnuColorFilter 
            Caption         =   "&Sepia"
            Index           =   2
         End
         Begin VB.Menu mnuColorFilter 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu mnuColorFilter 
            Caption         =   "&Colorize..."
            Index           =   4
         End
         Begin VB.Menu mnuColorFilter 
            Caption         =   "Replace &HS..."
            Index           =   5
         End
         Begin VB.Menu mnuColorFilter 
            Caption         =   "Replace &L..."
            Index           =   6
         End
         Begin VB.Menu mnuColorFilter 
            Caption         =   "Sh&ift..."
            Index           =   7
         End
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "&Definition"
         Index           =   1
         Begin VB.Menu mnuDefinition 
            Caption         =   "&Blur"
            Index           =   0
         End
         Begin VB.Menu mnuDefinition 
            Caption         =   "&Soften"
            Index           =   1
         End
         Begin VB.Menu mnuDefinition 
            Caption         =   "S&harpen"
            Index           =   2
         End
         Begin VB.Menu mnuDefinition 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu mnuDefinition 
            Caption         =   "&Diffuse"
            Index           =   4
         End
         Begin VB.Menu mnuDefinition 
            Caption         =   "&Pixelaze"
            Index           =   5
         End
         Begin VB.Menu mnuDefinition 
            Caption         =   "-"
            Index           =   6
         End
         Begin VB.Menu mnuDefinition 
            Caption         =   "Despec&kle"
            Index           =   7
         End
         Begin VB.Menu mnuDefinition 
            Caption         =   "Despeckle &more"
            Index           =   8
         End
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "&Edges"
         Index           =   2
         Begin VB.Menu mnuEdges 
            Caption         =   "&Contour"
            Index           =   0
         End
         Begin VB.Menu mnuEdges 
            Caption         =   "&Emboss"
            Index           =   1
         End
         Begin VB.Menu mnuEdges 
            Caption         =   "&Outline"
            Index           =   2
         End
         Begin VB.Menu mnuEdges 
            Caption         =   "&Relieve"
            Index           =   3
         End
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "&Special"
         Index           =   4
         Begin VB.Menu mnuSpecial 
            Caption         =   "&Noise"
            Index           =   0
         End
         Begin VB.Menu mnuSpecial 
            Caption         =   "&Scanlines"
            Index           =   1
         End
         Begin VB.Menu mnuSpecial 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu mnuSpecial 
            Caption         =   "&Dilate"
            Index           =   3
         End
         Begin VB.Menu mnuSpecial 
            Caption         =   "&Erode"
            Index           =   4
         End
         Begin VB.Menu mnuSpecial 
            Caption         =   "-"
            Index           =   5
         End
         Begin VB.Menu mnuSpecial 
            Caption         =   "&Texturize..."
            Index           =   6
         End
      End
   End
   Begin VB.Menu mnuViewTop 
      Caption         =   "&View"
      Begin VB.Menu mnuView 
         Caption         =   "&Quick bar"
         Checked         =   -1  'True
         Index           =   0
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuView 
         Caption         =   "Panoramic &view"
         Index           =   1
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuView 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuView 
         Caption         =   "&Properties..."
         Index           =   3
         Shortcut        =   {F8}
      End
   End
   Begin VB.Menu mnuHelpTop 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "&About"
         Index           =   0
      End
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'================================================
' iBMP v1.1
' Carles P.V. - 2003.01.11
'================================================
'
' - Thanks to:
'   + Allapi Network - http://www.allapi.net
'   + VB Accelerator - http://www.vbaccelerator.com
'   + VB Frood       - http://perso.wanadoo.fr/vbfrood/index.htm
'
' - Special thanks to:
'   + Ark
'   + Avery
'   + Manuel Augusto Santos
'   + Robert Rayment
'   + Vlad Vissoultchev
'
' LOG:
'
'   - 2002.12.28:
'     iBMP v1.0 finished
'
'   - 2002.12.30:
'     FIXED: fFilter initialization (Temp. DIB loading).
'     FIXED: fFilter [Close] (Incorrect main update).
'     ADDED: fTexturize Preview.
'
'   - 2002.12.31:
'     FIXED: fMain Menu/Reload command (Check Paste source / DIB exists).
'
'   - 2003.01.01:
'     FIXED: fTexturize Browse for folder (Check path access).
'
'   - 2003.01.04:
'     IMPVD: 'Despeckle' filter.
'            See: http://www.dai.ed.ac.uk/HIPR2/crimmins.htm#2 for more information.
'            Thanks, Robert.
'
'   - 2003.01.05:
'     FIXED: 'Despeckle' filter (Dark pixel processing).
'     FIXED: fFilter Color Dialog (Color wasn't initialized).
'     IMPVD: 'Replace' filter (Added Hue and Saturation tolerances).
'     IMPVD: 'Scanlines' filter (Added Black/White point params.).
'
'   - 2003.01.06:
'     FIXED: 'Despeckle' filter (Fixed again. Thanks again, Robert).
'
'   - 2003.01.07:
'     IMPVD: '16 greys' dithering.
'
'   - 2003.01.08:
'     ADDED: 'Emboss' and 'Relieve 'filters (Filter browser)
'     ADDED: Save fResize options.
'     ADDED: 'Dilate' and 'Erode' filters (Max.R.F. and Min.R.F. [4N pixels]).
'
'   - 2003.01.09:
'     ADDED: 'Replacle L' filter (Old Replace -> Replace HS).
'     ADDED: Mem. last filter (Filter browser).
'
'   - 2003.01.11: (iBMP v1.1)
'     ADDED: PNG (Load/Save) and JPEG (Save) support.
'            Thanks to Avery & Vlad Vissoultchev.
'            GDI+ needed (http://www.microsoft.com/downloads/release.asp?releaseid=32738).
'     ADDED: GDI+ Highest-quality resizing.
'     ADDED: 'Salt'n Pepper removal' filter (by Robert Rayment)
'
'   - 2003.01.12:.
'     ADDED: Dialog image preview support (by Ark) / Save last path/options
'
'   - 2003.01.13:
'     IMPVD: Increased Dialog window (original code by Ark, too).
'
'   - 2003.01.14:
'     IMPVD: Load-Flickering (Thanks to Zhu JinYong).
'     FIXED: Image validation (Load).
'
'   - 2003.01.15:
'     FIXED: Image-flickering (Resize/Zoom) really fixed (I think).
'     IMPVD: Color dialog -> API (Removed Com. Dial. ref.)
'
'   - 2003.01.22:
'     ADDED: Memory position on Zoom (Avery's suggestion).
'
'   - 2003.01.29:
'     FIXED: Clipping visible rect. in fPanView.
'     ADDED: Maximum allowed size (See fResize dialog. Current max.: 4 MPixels).
'            This value can be modified: MAX_PIXELS_SIZE constant.
'
'   - 2003.02.04:
'     IMPVD: 'Texturize' filter.
'
'   - 2003.02.05:
'     FIXED: Crop rectangle now correctly positioned (Fit mode enabled).
'
'   - 2003.02.06:
'     IMPVD: Resizable file dialog.
'     FIXED: Mouse pointer now correctly updated (Fit mode enabled).
'
'   - 2003.02.11:
'     FIXED: Pixel offset (GDI+ resizing/resampling).
'
'   - 2003.02.16:
'     IMPVD: 'Despeckler' filter simplified.
'     IMPVD: Last texture preserved (Texturize dialog).
'     IMPVD: Added 'Just Saved' flag.
'
'   - 2003.02.22:
'     ADDED: 'Outline' filter.
'
'   - 2003.02.23:
'     IMPVD: Filter browser interface:
'            · <Before> view: based on ucCanvas user control (Zoom/Scroll...).
'            · Added all filters (none parameter filters).
'     IMPVD: 'Replace HS' filter (H-S tol.: AND criteria).
'     FIXED: Load texture bitmap (Error handling).
'
'   - 2003.02.25:
'     ADDED: 'Texture 90º' rotation.
'
'   - 2003.02.26:
'     FIXED: Correct restoring main image by closing dialogs (Filter/Texturize)
'            through caption close button.
'
'   - 2003.03.01:
'     FIXED: Websafe palette generation (Entry offsets).
'     ADDED: Gif saving.
'            Default GDI+ GIF saving is std. quality (not optimal palette).
'            Check my other submission for better quality GIF saving:
'            http://www.pscode.com/vb/scripts/ShowCode.asp?txtCodeId=43599&lngWId=1
'
'   - 2003.03.13:
'     IMPVD: Removed DIB buffer copy from Dithering functions.
'
'   - 2003.03.15:
'     IMPVD: Up-Down DIB. Removed inverted storing of dithered indexes.
'     IMPVD: Optimal palette entries sorted by frequency.
'            Reduced size of source DIB to extract optimal palette,
'            increased to 150x150 max.
'
'   - 2003.03.17:
'     IMPVD: '16 Greys' dithering.
'
'   - 2003.11.03:
'     Code revised. Added owner drawn toolbar (removed extra controls dependency).
'     Added mouse wheel support (zoom +/-).
'     -> iBMP v1.2
'
'   - 2004.02.08:
'     Added TIFF Load and Save (24 bpp) support.
'
'   - 2004.05.14:
'     Fixed incorrect scrollbar background refresh (Windows XP/NT/2000)



Option Explicit
Option Compare Text

'-- Bitmap control
Public WithEvents DIBFilter As cDIBFilter   ' DIB Filter object  (24 bpp)
Attribute DIBFilter.VB_VarHelpID = -1
Public WithEvents DIBDither As cDIBDither   ' DIB Dither object  (1, 4, 8 bpp)
Attribute DIBDither.VB_VarHelpID = -1
Public DIBPal               As New cDIBPal  ' DIB Palette object (1, 4, 8 bpp)
Public DIBSave              As New cDIBSave ' Save object (BMP)  (1, 4, 8, 24 bpp)
Attribute DIBSave.VB_VarHelpID = -1
Public DIBbpp               As Byte         ' Current color depth

'-- Undo/Redo control
Private Const m_UNDO_LEVELS As Long = 25    ' Max Undo levels
Private m_AppID             As Long         ' Application ID (fMain.hwnd)
Private m_UndoPos           As Long         ' Current Undo position
Private m_UndoMax           As Long         ' Undo max. reached position
Private m_Temp              As String       ' Temporary folder

'-- Dialog
Private m_LastFilter        As Integer      ' Last used filter (Filter browser)
Private m_LastFilename      As String       ' Current file
Private m_LastPath          As String       ' Current path
Private m_Saved             As Boolean      ' Just saved
Private m_FileExt           As String       ' Current file/ext
Private m_DialogPreview     As Boolean      ' Dialog: Show preview
Private m_DialogFitMode     As Boolean      ' Dialog: Fit mode
Private m_DialogJPEGquality As Integer      ' Dialog: JPEG quality (0-100)

'-- GDI+
Private m_GDIpToken         As Long         ' Needed to close GDI+



'========================================================================================
' Main
'========================================================================================

Private Sub Form_Load()

  Dim GpInput As GdiplusStartupInput
    
    If (App.LogMode <> 1) Then
        Call MsgBox("Please, compile me to get real speed...", vbInformation)
    End If
    
    '-- Load the GDI+ Dll
    GpInput.GdiplusVersion = 1
    If (mGDIpEx.GdiplusStartup(m_GDIpToken, GpInput) <> [OK]) Then
        Call MsgBox("Error loading GDI+!", vbCritical)
        Call Unload(Me)
        Exit Sub
    End If
    
    '-- Load settings
    Call mSettings.LoadMainSettings

    '-- Initialize toolbar
    Call Quick.BuildToolbar(LoadResPicture("BITMAP_TBQUICK", vbResBitmap), &HFF00FF, 16, "NNN|NN|OO|NNN|C|NN|NNN|NNN|NN")
    Call Quick.SetTooltips("Open...|Save as...|Print...|Undo|Redo|Scroll mode|Crop mode|Zoom in|Zoom out|1:1|Fit mode|Resize...|Orientation...|Sharpen|Soften|Despeckle|Brightness...|Contrast...|Saturation...|Filter browser...|Texturize...")
    Call Quick.Refresh
    Call Quick.EnableButton(4, False)
    Call Quick.EnableButton(5, False)
    Call Quick.CheckButton(6, True)
    '-- Enable/Disable Menu/Toolbar options
    Call pvUpdateMenuAndToolbar
    '-- Initial zoom = 100%
    Info.TextZoom = "100%"
    
    '-- Initialize 'evented' objects
    Set DIBFilter = New cDIBFilter
    Set DIBDither = New cDIBDither
    
    '-- Get App. ID and <Temp> path (Undo/Redo temp. files)
    m_AppID = Me.hWnd
    m_Temp = IIf(Environ$("tmp") <> vbNullString, Environ$("tmp"), Environ$("temp"))
    
    '-- Hook wheel for zooming
    Call mHook.HookWheel
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  Dim sRet As VbMsgBoxResult
    
    If (Canvas.DIB.hDIB <> 0 And Not m_Saved And m_UndoPos > 1) Then
    
        sRet = MsgBox("Do you want to save before exit ?", vbYesNoCancel Or vbInformation)
        Select Case sRet
            Case vbYes    '-- Save
                Call mnuFile_Click(2)
                Cancel = 0
            Case vbNo     '-- Don't save
                Cancel = 0
            Case vbCancel '-- Cancel
                Cancel = 1
        End Select
    End If
    If (Cancel = 0) Then Call Unload(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)

    '-- Save settings
    Call mSettings.SaveMainSettings
    
    '-- Delete temp. files
    Call pvClearAllDIB
    
    ' Unload the GDI+ Dll
    Call mGDIpEx.GdiplusShutdown(m_GDIpToken)
       
    '-- Free forms
    Call Unload(fAbout)
    Set fAbout = Nothing
    Call Unload(fDialogEx)
    Set fDialogEx = Nothing
    Call Unload(fFilter)
    Set fFilter = Nothing
    Call Unload(fOrientation)
    Set fOrientation = Nothing
    Call Unload(fPanView)
    Set fPanView = Nothing
    Call Unload(fPrint)
    Set fPrint = Nothing
    Call Unload(fProperties)
    Set fProperties = Nothing
    Call Unload(fResize)
    Set fResize = Nothing
    Call Unload(fTexturize)
    Set fTexturize = Nothing
    
    '-- Free objects
    Set DIBFilter = Nothing
    Set DIBDither = Nothing
    Set DIBPal = Nothing
    Set DIBSave = Nothing
    
    Set fMain = Nothing
End Sub

'========================================================================================
' Processing
'========================================================================================

Public Sub Canvas_DIBProgress(ByVal p As Long)
    Progress = p
End Sub

Public Sub Canvas_DIBProgressEnd()

    '-- Progress end
    Progress = 0
    Call fPanView.Repaint
    '-- DIB processed (-> 24bpp: Size changed, orientation changed)
    DIBbpp = 24
    Call pvSetPalMode(DIBbpp)
    Info.TextInfo = Canvas.DIB.Width & "x" & Canvas.DIB.Height & "x" & DIBbpp & "bpp"
    '-- Save Undo
    Call pvSaveUndoDIB
End Sub

Public Sub DIBFilter_Progress(ByVal p As Long)
    Progress = p
End Sub

Public Sub DIBFilter_ProgressEnd()

    '-- Progress end
    Progress = 0
    Call Canvas.Repaint
    Call fPanView.Repaint
    '-- If not previewing (Filter browse box), save Undo
    If (fFilter.Previewing = False And fTexturize.Previewing = False) Then Call pvSaveUndoDIB
End Sub

Public Sub DIBDither_Progress(ByVal p As Long)
    Progress = p
End Sub

Public Sub DIBDither_ProgressEnd()

    '-- Progress end
    Progress = 0
    Call Canvas.Repaint
    Call fPanView.Repaint
    '-- Save Undo
    Call pvSaveUndoDIB
End Sub

'========================================================================================
' Resizing
'========================================================================================

Private Sub Form_Resize()
    
    '-- Resize Canvas
    On Error Resume Next
        Call Canvas.Move(0, IIf(Quick.Visible, Quick.Height + 2, 0), ScaleWidth, ScaleHeight - IIf(Quick.Visible, Quick.Height + 2, 0) - Info.Height - Progress.Height)
    On Error GoTo 0
End Sub
Private Sub Form_Paint()
    Me.Line (0, 0)-(ScaleWidth, 0), vb3DShadow
    Me.Line (0, 1)-(ScaleWidth, 1), vb3DHighlight
End Sub

Private Sub Info_Resize()
    '-- Compact path to Textfile panel width
    If (m_LastFilename <> "[Untitled]") Then
        Info.TextFile = CompactPath(Me.hDC, m_LastFilename, Info.TextFileWidth)
      Else
        Info.TextFile = "[Untitled]"
    End If
End Sub

'========================================================================================
' Quick bar / Menus
'========================================================================================

Private Sub Quick_ButtonClick(ByVal Index As Long, ByVal xLeft As Long, ByVal yTop As Long)
    
    Select Case Index
        Case 1  '-- Open...
            Call mnuFile_Click(0)
        Case 2  '-- Save as...
            Call mnuFile_Click(2)
        Case 3  '-- Print...
            Call mnuFile_Click(4)
        Case 4  '-- Undo
            Call mnuEdit_Click(0)
        Case 5  '-- Redo
            Call mnuEdit_Click(1)
        Case 6  '-- Scroll mode
            Call mnuEdit_Click(9)
        Case 7  '-- Crop mode
            Call mnuEdit_Click(10)
        Case 8  '-- Zoom in
            Call mnuZoom_Click(0)
        Case 9  '-- Zoom out
            Call mnuZoom_Click(1)
        Case 10 '-- 1:1
            Call mnuZoom_Click(2)
        Case 11 '-- Fit mode
            Call mnuZoom_Click(4)
        Case 12 '-- Resize
            Call mnuEdit_Click(6)
        Case 13 '-- Orientation
            Call mnuEdit_Click(7)
        Case 14 '-- Sharpen
            Call mnuDefinition_Click(2)
        Case 15 '-- Soften
            Call mnuDefinition_Click(1)
        Case 16 '-- Despeckle
            Call mnuDefinition_Click(7)
        Case 17 '-- Brightness...
            Call mnuAdjust_Click(0)
        Case 18 '-- Contrast...
            Call mnuAdjust_Click(1)
        Case 19 '-- Saturation
            Call mnuAdjust_Click(2)
        Case 20 '-- Filter browser...
            Call mnuAdjust_Click(4)
        Case 21 '-- Texturize...
            Call mnuSpecial_Click(6)
    End Select
End Sub

Private Sub mnuFile_Click(Index As Integer)
        
  Dim fDlg     As New fDialogEx
  Dim sRet     As String
  Dim bSuccess As Boolean
    
    Select Case Index
    
        Case 0 '-- Open...
      
            '-- Show Open Dialog
            sRet = GetFileName(m_LastPath, "Supported files|*.bmp;*.gif;*.jpg;*.png;*.tif|Bitmap files (*.bmp)|*.bmp|GIF files (*.gif)|*.gif|JPEG files (*.jpg)|*.jpg|PNG files (*.png)|*.png|TIFF files (*.tif)|*.tif", 0, "Open", True, fDlg)
            
            If (sRet <> vbNullString) Then
            
                '-- Get last path
                m_LastPath = sRet
                '-- Create DIB
                DoEvents
                Screen.MousePointer = vbHourglass
                Call pvSetDIBPicture(pvGetStdPicture(sRet, bSuccess))
                Screen.MousePointer = vbNormal
                
                If (bSuccess) Then
                    '-- Reset Undo/Redo and save first Undo
                    Call pvClearAllDIB
                    Call pvSaveUndoDIB
                    '-- Save info
                    m_LastFilename = sRet
                    Call Info_Resize
                End If
            End If
        
        Case 1 '-- Save
      
            If (m_LastFilename = "[Untitled]" Or (FileFound(pvExtToBMP(m_LastFilename)) And pvExtToBMP(m_LastFilename) <> m_LastFilename)) Then
                
                '-- Save as...
                Call Unload(fDlg)
                Set fDlg = Nothing
                Call mnuFile_Click(2)
                
              Else
                '-- Save as BMP
                DoEvents
                Call DIBSave.Save_BMP(pvExtToBMP(m_LastFilename), Canvas.DIB, DIBPal, DIBDither, DIBbpp)
                '-- Saved flag
                m_Saved = True
                '-- Save info
                m_LastFilename = pvExtToBMP(m_LastFilename)
                Call Info_Resize
            End If
        
        Case 2 '-- Save as...
      
            '-- Show Open Dialog
            sRet = GetFileName(m_LastFilename, "Bitmap (*.bmp)|*.bmp|GIF (*.gif)|*.gif|JPEG (*.jpg)|*.jpg|PNG (*.png)|*.png|TIFF (*.tif)|*.tif", 0, "Save as", False, fDlg)
            
            If (sRet <> vbNullString) Then
            
                '-- Missing ext.?
                Call pvCorrectExt(sRet)
                
                '-- Save...
                DoEvents
                Select Case m_FileExt
                    Case "*.bmp" '-- BMP
                        Call DIBSave.Save_BMP(sRet, Canvas.DIB, DIBPal, DIBDither, DIBbpp)
                    Case "*.gif" '-- GIF
                        Call mGDIpEx.SaveDIB(fMain.Canvas.DIB, sRet, [ImageGIF])
                    Case "*.jpg" '-- JPEG
                        Call mGDIpEx.SaveDIB(fMain.Canvas.DIB, sRet, [ImageJPEG], fDlg.txtJPEGQuality)
                    Case "*.png" '-- PNG
                        Call mGDIpEx.SaveDIB(fMain.Canvas.DIB, sRet, [ImagePNG])
                    Case "*.tif" '-- TIFF
                        Call mGDIpEx.SaveDIB(fMain.Canvas.DIB, sRet, [ImageTIFF])
                End Select
                '-- Saved flag
                m_Saved = True
    
                '-- Save info
                If (m_FileExt = "*.bmp") Then
                    m_LastFilename = sRet
                End If
                Call Info_Resize
            End If
      
        Case 4 '-- Print...
            If (Printers.Count) Then
                Call fPrint.Show(vbModal, Me)
              Else
                Call MsgBox("Sorry, no printers installed.", vbExclamation)
            End If
        
        Case 6 '-- Exit
            Call Unload(Me)
    End Select
    
    Call Unload(fDlg)
    Set fDlg = Nothing
End Sub

Private Sub mnuEditTop_Click()

    '-- Enable/Disable Undo/Redo commands
    mnuEdit(0).Enabled = (m_UndoPos > 1)
    mnuEdit(1).Enabled = (m_UndoPos <> m_UndoMax)

    '-- Enable/Disable Copy/Paste commands
    mnuEdit(3).Enabled = (Canvas.DIB.hDIB <> 0)
    mnuEdit(4).Enabled = (Clipboard.GetFormat(vbCFBitmap))
End Sub

Private Sub mnuEdit_Click(Index As Integer)
    
  Dim sRet As VbMsgBoxResult
    
    Select Case Index
    
        Case 0 '-- Undo
            Call Undo
          
        Case 1 '-- Redo
            Call Redo
          
        Case 3 '-- Copy
            If (Canvas.DIB.hDIB <> 0) Then
                Call Canvas.DIB.CopyToClipboard
            End If
          
        Case 4 '-- Paste
            If (Clipboard.GetFormat(vbCFBitmap)) Then
                
                '-- Something there ?
                If (Canvas.DIB.hDIB <> 0 And Not m_Saved And m_UndoPos > 1) Then
                    '-- Ask for save
                    sRet = MsgBox("Save changes before Paste ?", vbYesNoCancel Or vbInformation)
                    Select Case sRet
                        Case vbYes    '-- Save
                            Call mnuFile_Click(1)
                        Case vbNo     '-- Ignore
                        Case vbCancel '-- Exit
                            Exit Sub
                    End Select
                End If
                
                '-- Initialize DIB
                Call pvSetDIBPicture(Clipboard.GetData(vbCFBitmap))
                '-- Reset Undo/Redo and save first Undo
                Call pvClearAllDIB
                Call pvSaveUndoDIB
                '-- [Untitled] image
                m_LastFilename = "[Untitled]"
                Call Info_Resize
            End If
          
        Case 6 '-- Resize
            Call fResize.Show(vbModal, Me)
          
        Case 7 '-- Orientation
            Call fOrientation.Show(vbModal, Me)
          
        Case 9 '-- Scroll mode
            Canvas.WorkMode = [cnvScrollMode]
            Call Canvas.Repaint
            mnuEdit(9).Checked = True
            mnuEdit(10).Checked = False
            Call Quick.CheckButton(6, True)
            
        Case 10 '-- Crop mode
            Canvas.WorkMode = [cnvCropMode]
            mnuEdit(9).Checked = False
            mnuEdit(10).Checked = True
            Call Quick.CheckButton(7, True)
    End Select
End Sub

Public Sub mnuZoom_Click(Index As Integer)
    
    Select Case Index
    
        Case 0 '-- Zoom +
            Canvas.Zoom = Canvas.Zoom + IIf(Canvas.Zoom < 25, 1, 0)
          
        Case 1 '-- Zoom -
            Canvas.Zoom = Canvas.Zoom - IIf(Canvas.Zoom > 1, 1, 0)
          
        Case 2 '-- 1 : 1
            Canvas.Zoom = 1
          
        Case 4 '-- Best fit
            With mnuZoom(4)
                .Checked = Not .Checked
                Canvas.FitMode = .Checked
                Call Quick.CheckButton(11, .Checked)
            End With
    End Select
    
    Call Canvas.Resize
    Info.TextZoom = Format(Canvas.Zoom, "0%")
End Sub

Private Sub mnuColors_Click(Index As Integer)
        
  Dim sDIB As New cDIB
  Dim bfW As Long, bfH As Long
  Dim bfx As Long, bfy As Long
        
    If (Not mnuColors(Index).Checked) Then
    
        Select Case Index
        
            Case 0  '-- Black and White (Stucki)
                DIBbpp = 1
                Call DIBPal.CreateBlackAndWhite
                Call DIBDither.DitherToBlackAndWhite(Canvas.DIB, 16)
            
            Case 1  '-- 16 greys
                DIBbpp = 4
                Call DIBPal.CreateGreys_016
                Call DIBDither.DitherToGreyPalette(Canvas.DIB, DIBPal, True)
            
            Case 2  '-- 256 greys
                DIBbpp = 8
                Call DIBPal.CreateGreys_256
                Call DIBDither.DitherToGreyPalette(Canvas.DIB, DIBPal, False)
            
            Case 4  '-- 2 colors
                If (mnuColors(0).Checked = False) Then
                    DIBbpp = 1
                    Call DIBPal.CreateBlackAndWhite
                    Call DIBDither.DitherToBlackAndWhite(Canvas.DIB, 16)
                End If
            
            Case 5  '-- 16 colors
                If (mnuColors(1).Checked = False) Then
                    DIBbpp = 4
                    If (DIBPal.IsGreyScale) Then
                        Call DIBPal.CreateGreys_016
                        Call DIBDither.DitherToGreyPalette(Canvas.DIB, DIBPal, True)
                      Else
                        '-- Strecth to fit 150x150 (This will speed up all this)
                        Call Canvas.DIB.GetBestFitInfo(150, 150, bfx, bfy, bfW, bfH)
                        Call sDIB.Create(bfW, bfH)
                        Call sDIB.LoadDIBBlt(Canvas.DIB)
                        '-- Create optimal palette and dither.
                        '   I don't know why these weight coeffs. work well...
                        '   wChannel = f(Lchannel)
                        '   wR = 1/(3-0.222) = 0.360
                        '   wG = 1/(3-0.707) = 0.436
                        '   wB = 1/(3-0.071) = 0.341
                        Screen.MousePointer = vbHourglass
                        Call DIBPal.CreateOptimal(sDIB, 16, 8, 0.36, 0.436, 0.341)
                        Screen.MousePointer = vbNormal
                        Call DIBDither.DitherToColorPalette(Canvas.DIB, DIBPal, True)
                    End If
                End If
            
                Case 6  '-- 256 colors
                If (mnuColors(2).Checked = False) Then
                    DIBbpp = 8
                        If (DIBPal.IsGreyScale) Then
                        Call DIBPal.CreateGreys_256
                        Call DIBDither.DitherToGreyPalette(Canvas.DIB, DIBPal, False)
                      Else
                        '-- Strecth to fit 150x150 (This will speed up all this)
                        Call Canvas.DIB.GetBestFitInfo(150, 150, bfx, bfy, bfW, bfH)
                        Call sDIB.Create(bfW, bfH)
                        Call sDIB.LoadDIBBlt(Canvas.DIB)
                        '-- Create optimal palette and dither
                        Screen.MousePointer = vbHourglass
                        Call DIBPal.CreateOptimal(sDIB, 256, 8, 1, 1, 1)
                        Screen.MousePointer = vbNormal
                        Call DIBDither.DitherToColorPalette(Canvas.DIB, DIBPal, True)
                    End If
                End If
    
            Case 8  '-- True color (24bpp)
                DIBbpp = 24
                Call DIBPal.Clear
                Call DIBDither.DitherToTrueColor(Canvas.DIB)
        End Select
        
        '-- Refresh
        Call Canvas.Repaint
        
        '-- Select current mode
        Call pvSetPalMode(DIBbpp)
        
        '-- Update info
        Info.TextInfo = Canvas.DIB.Width & "x" & Canvas.DIB.Height & "x" & DIBbpp & "bpp"
    End If
End Sub

Private Sub mnuAdjust_Click(Index As Integer)
    
    Select Case Index
    
        Case 0 '-- Brightness...
            Call fFilter.Initialize(fltBrightness)
            Call fFilter.Show(vbModal, Me)
              
        Case 1 '-- Contrast...
            Call fFilter.Initialize(fltContrast)
            Call fFilter.Show(vbModal, Me)
              
        Case 2 '-- Saturation...
            Call fFilter.Initialize([fltSaturation])
            Call fFilter.Show(vbModal, Me)
              
        Case 4 '-- Filter browser...
            Call fFilter.Initialize(m_LastFilter)
            Call fFilter.Show(vbModal, Me)
    End Select
    
    Call Canvas.Repaint
End Sub

Private Sub mnuColorFilter_Click(Index As Integer)

    Select Case Index
    
        Case 0 '-- Greys
            Call DIBFilter.Greys(Canvas.DIB)
    
        Case 1 '-- Negative
            Call DIBFilter.Negative(Canvas.DIB)
        
        Case 2 '-- Sepia
            Call DIBFilter.Colorize(Canvas.DIB, 0.5, 0.25)
        
        Case 4 '-- Colorize...
            Call fFilter.Initialize([fltColorize])
            Call fFilter.Show(vbModal, Me)
        
        Case 5 '-- Replace HS...
            Call fFilter.Initialize([fltReplaceHS])
            Call fFilter.Show(vbModal, Me)
        
        Case 6 '-- Replace L...
            Call fFilter.Initialize([fltReplaceL])
            Call fFilter.Show(vbModal, Me)
      
        Case 7 '-- Shift...
            Call fFilter.Initialize(fltShift)
            Call fFilter.Show(vbModal, Me)
    End Select
    
    Call Canvas.Repaint
End Sub

Private Sub mnuDefinition_Click(Index As Integer)

    Select Case Index
    
        Case 0 '-- Blur
            Call DIBFilter.Blur(Canvas.DIB)
          
        Case 1 '-- Soften
            Call DIBFilter.Soften(Canvas.DIB)
          
        Case 2 '-- Sharpen
            Call DIBFilter.Sharpen(Canvas.DIB)
          
        Case 4 '-- Diffuse
            Call DIBFilter.Diffuse(Canvas.DIB)
          
        Case 5 '-- Pixelize
            Call DIBFilter.Pixelize(Canvas.DIB)
          
        Case 7 '-- Despeckle
            Call DIBFilter.Despeckle(Canvas.DIB)
          
        Case 8 '-- Despeckle more
            Call DIBFilter.DespeckleMore(Canvas.DIB)
    End Select
    
    Call Canvas.Repaint
End Sub

Private Sub mnuEdges_Click(Index As Integer)
    
    Select Case Index
    
        Case 0 '-- Contour
            Call DIBFilter.Contour(Canvas.DIB)
        
        Case 1 '-- Emboss
            Call DIBFilter.Emboss(Canvas.DIB)
        
        Case 2 '-- Outline
            Call DIBFilter.Outline(Canvas.DIB)
        
        Case 3 '-- Relieve
            Call DIBFilter.Relieve(Canvas.DIB)
    End Select
    
    Call Canvas.Repaint
End Sub

Private Sub mnuSpecial_Click(Index As Integer)

    Select Case Index
    
        Case 0 '-- Noise
            Call DIBFilter.Noise(Canvas.DIB)
        
        Case 1 '-- Scanlines
            Call DIBFilter.Scanlines(Canvas.DIB)
        
        Case 3 '-- Dilate (Max.R.F. - 4N)
            Call DIBFilter.RankFilterMaximum(Canvas.DIB)
        
        Case 4 '-- Erode (Min.R.F. - 4N)
            Call DIBFilter.RankFilterMinimum(Canvas.DIB)
        
        Case 6 '-- Texturize...
            Call fTexturize.Show(vbModal, Me)
    End Select
    
    Call Canvas.Repaint
End Sub

Private Sub mnuView_Click(Index As Integer)

    Select Case Index
    
        Case 0 '-- Quick bar
            With Quick
                .Visible = Not .Visible
                 mnuView(0).Checked = .Visible
            End With
            Call Form_Resize
            
        Case 1 '-- Panoramic view
            With mnuView(1)
                .Checked = Not .Checked
                If (.Checked) Then
                    Call fPanView.Show(, Me)
                  Else
                    Call fPanView.Hide
                End If
            End With
            
        Case 3 '-- Properties...
            Call fProperties.Show(vbModal, Me)
    End Select
End Sub

Private Sub mnuHelp_Click(Index As Integer)
    
    Select Case Index
        Case 0 '-- About
            Call fAbout.Show(vbModal, Me)
    End Select
End Sub

'========================================================================================
' Canvas key scrolling / Crop
'========================================================================================

Private Sub Canvas_Resize()
    Call fPanView.Repaint
End Sub

Private Sub Canvas_Scroll()
    Call fPanView.Repaint
End Sub

Private Sub Canvas_KeyDown(KeyCode As Integer, Shift As Integer)
    
  Dim scrHMax As Long, scrVMax As Long
  Dim scrHPos As Long, scrVPos As Long

  Dim bScroll As Boolean
  Dim lInc    As Long
    
    With Canvas
    
        Select Case KeyCode
            Case vbKeyAdd      '{NumPad +}
                Call mnuZoom_Click(0)
            Case vbKeySubtract '{NumPad -}
                Call mnuZoom_Click(1)
            Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight
                Call .GetScrollInfo(scrHMax, scrVMax, scrHPos, scrVPos)
                bScroll = True
        End Select
                    
        If (bScroll) Then
        
            lInc = 10 * Canvas.Zoom
            
            Select Case KeyCode
                Case vbKeyUp    '{Cursor Up}
                    If (scrVPos > 0) Then
                        Call .SetScrollInfo(scrHPos, scrVPos - lInc)
                      Else
                        Call .SetScrollInfo(scrHPos, 0)
                    End If
                Case vbKeyDown  '{Cursor Down}
                    If (scrVPos < scrVMax) Then
                        Call .SetScrollInfo(scrHPos, scrVPos + lInc)
                      Else
                        Call .SetScrollInfo(scrHPos, scrVMax)
                    End If
                Case vbKeyLeft  '{Cursor Left}
                    If (scrHPos > 0) Then
                        Call .SetScrollInfo(scrHPos - lInc, scrVPos)
                      Else
                        Call .SetScrollInfo(0, scrVPos)
                    End If
                Case vbKeyRight '{Cursor Right}
                    If (scrHPos < scrHMax) Then
                        Call .SetScrollInfo(scrHPos + lInc, scrVPos)
                      Else
                        Call .SetScrollInfo(scrHMax, scrVPos)
                    End If
            End Select
            Call fPanView.Repaint
        End If
        
        Call Canvas.Repaint
    End With
End Sub

Private Sub Canvas_Crop()

    '-- Change to True color mode
    DIBbpp = 24
    Call pvSetPalMode(DIBbpp)
    
    '-- Update Info and Progress
    With Canvas.DIB
        Info.TextInfo = .Width & "x" & .Height & "x" & DIBbpp & "bpp"
        Progress.Max = .Height
    End With
    
    '-- Refresh Panoramic view
    Call fPanView.Repaint
    
    '-- Save Undo
    Call pvSaveUndoDIB
End Sub

'========================================================================================
' DIB/Palette initialization
'========================================================================================

Private Function pvGetStdPicture(ByVal sFilename As String, bSuccess As Boolean) As StdPicture

    On Error Resume Next
    
    If (pvGetExt(sFilename) = "png" Or pvGetExt(sFilename) = "tif") Then
        '-- Use GDI+ loading
        Set pvGetStdPicture = mGDIpEx.LoadPictureEx(sFilename)
      Else
        '-- Use VB LoadPicture
        Set pvGetStdPicture = LoadPicture(sFilename)
    End If
    
    '-- Is there an image ?
    bSuccess = Not (pvGetStdPicture Is Nothing)
    
    If (bSuccess = False) Then
        '-- Nothing loaded
        Call MsgBox("Unexpected error loading image", vbExclamation)
    End If
    
    On Error GoTo 0
End Function
    
Private Sub pvSetDIBPicture(Image As StdPicture)
  
  Static lstW As Long
  Static lstH As Long

    If (Not Picture Is Nothing) Then

        '-- Save last DIB dimensions
        lstW = Canvas.DIB.Width
        lstH = Canvas.DIB.Height
        
        '-- Clear palette
        Call DIBPal.Clear
        
        '-- Create 32bpp DIB section from std. picture.
        '   Case source <=8bpp, palette saved in DIBPal, palette indexes in DIBDither.
        '   Return value: source color depth / 0 = Err.
        DIBbpp = Canvas.DIB.CreateFromStdPicture(Image, DIBPal, DIBDither)
        
        '-- Select current depth mode
        Call pvSetPalMode(DIBbpp)
        
        '-- Remove Crop rectangle and resize canvas
        Call Canvas.RemoveCropRectangle
        With Canvas.DIB
            If (lstW <> .Width Or lstH <> .Height) Then
                Call Canvas.Resize
              Else
                Call Canvas.Repaint
            End If
        End With
        
        '-- Refresh panoramic view
        Call fPanView.Repaint
        
        '-- Set progress bar max value
        Progress.Max = Canvas.DIB.Height
        
        '-- Show image info: Size + bpp
        With Info
            .TextInfo = Canvas.DIB.Width & "x" & Canvas.DIB.Height & "x" & DIBbpp & "bpp"
            .TextZoom = Format(Canvas.Zoom, "0%")
        End With
    End If
End Sub

Private Sub pvSetPalMode(ByVal bpp As Long)
  
  Dim lIdxNew As Long
  Dim lIdxOld As Long
    
    Select Case bpp
        Case 1  '-- 2 colors / Black and White
            lIdxNew = IIf(DIBPal.IsGreyScale, 0, 4)
        Case 4  '-- 16 colors / 16 greys
            lIdxNew = IIf(DIBPal.IsGreyScale, 1, 5)
        Case 8  '-- 256 colors / 256 greys
            lIdxNew = IIf(DIBPal.IsGreyScale, 2, 6)
        Case 24 '-- True color
            lIdxNew = 8
        Case Else
            Exit Sub
    End Select
    
    For lIdxOld = 0 To 8
        mnuColors(lIdxOld).Checked = False
    Next lIdxOld
    mnuColors(lIdxNew).Checked = True
    
    '-- Update (Enable/Disable options) Menu and Toolbar
    Call pvUpdateMenuAndToolbar
End Sub

Private Sub pvUpdateMenuAndToolbar()
    
  Dim bEnbBPP As Boolean
  Dim bEnbDIB As Boolean
  Dim lIdx    As Long
  
    bEnbBPP = (DIBbpp = 24)
    bEnbDIB = (Canvas.DIB.hDIB <> 0)
    
    On Error Resume Next
    
    '== Color depth enable/disable
    
    '-- Update Menu
    For lIdx = 0 To mnuAdjust.Count - 1
        mnuAdjust(lIdx).Enabled = bEnbBPP
    Next lIdx
    For lIdx = 0 To mnuFilter.Count - 1
        mnuFilter(lIdx).Enabled = bEnbBPP
    Next lIdx
    '-- Update Toolbar
    For lIdx = 14 To 21
        Call Quick.EnableButton(lIdx, bEnbBPP)
    Next lIdx
        
    '== DIB exists enable/disable
    
    '-- Update Menu
    For lIdx = 1 To 6
        mnuFile(lIdx).Enabled = bEnbDIB
    Next lIdx
    For lIdx = 0 To 1
        mnuEdit(lIdx).Enabled = bEnbDIB
    Next lIdx
    For lIdx = 6 To 7
        mnuEdit(lIdx).Enabled = bEnbDIB
    Next lIdx
    For lIdx = 0 To mnuColors.Count - 1
        mnuColors(lIdx).Enabled = bEnbDIB
    Next lIdx
    mnuView(3).Enabled = bEnbDIB
    '-- Update Toolbar
    With Quick
        Call .EnableButton(2, bEnbDIB)
        Call .EnableButton(3, bEnbDIB)
        Call .EnableButton(12, bEnbDIB)
        Call .EnableButton(13, bEnbDIB)
    End With
    
    On Error GoTo 0
End Sub

'========================================================================================
' Undo/Redo control
'========================================================================================

Public Sub Undo()

  Dim sPath As String

    If (m_UndoPos > 1) Then
        
        '-- Get path
        sPath = m_Temp & "\b" & m_AppID & Format(m_UndoPos - 2, "000") & ".dat"
        '-- Load Undo DIB
        Call pvSetDIBPicture(LoadPicture(sPath))
        '-- Refresh Panoramic view
        Call fPanView.Repaint
    
        If (m_UndoPos > 0) Then
            m_UndoPos = m_UndoPos - 1
        End If
    End If
    Call CheckUndoRedoState
End Sub

Public Sub Redo()

  Dim sPath As String
  
    If (m_UndoPos < m_UndoMax) Then
    
        '-- Get path
        sPath = m_Temp & "\b" & m_AppID & Format(m_UndoPos, "000") & ".dat"
        '-- Load Redo DIB
        Call pvSetDIBPicture(LoadPicture(sPath))
        '-- Refresh Panoramic view
        Call fPanView.Repaint
    
        m_UndoPos = m_UndoPos + 1
        If (m_UndoPos > m_UndoMax) Then
            m_UndoMax = m_UndoPos
        End If
    End If
    Call CheckUndoRedoState
End Sub

Public Sub CheckUndoRedoState()

    '-- Enable/disable Undo
    Call Quick.EnableButton(4, (m_UndoPos > 1 And Canvas.DIB.hDIB <> 0))
    Call fFilter.Command.EnableButton(1, Quick.IsButtonEnabled(4))
    Call fTexturize.Command.EnableButton(1, Quick.IsButtonEnabled(4))
    
    '-- Enable/disable Redo
    Call Quick.EnableButton(5, (m_UndoPos < m_UndoMax And Canvas.DIB.hDIB <> 0))
    Call fFilter.Command.EnableButton(2, Quick.IsButtonEnabled(5))
    Call fTexturize.Command.EnableButton(2, Quick.IsButtonEnabled(5))
End Sub

Private Sub pvClearAllDIB()
    
    '-- Delete all temp. files
    On Error Resume Next
       Kill m_Temp & "\b" & m_AppID & "*.dat"
    On Error GoTo 0
    
    '-- Reset 'counters'
    m_UndoPos = 0
    m_UndoMax = 0
End Sub

Private Sub pvSaveUndoDIB()

  Dim lIdx  As Long
  Dim sPath As String
    
    '-- Get path
    sPath = m_Temp & "\b" & m_AppID & Format(m_UndoPos, "000") & ".dat"
    '-- Save DIB
    With fMain
        Call .DIBSave.Save_BMP(sPath, .Canvas.DIB, .DIBPal, .DIBDither, .DIBbpp)
    End With
    '-- Saved flag
    m_Saved = False
    
    If (m_UndoMax - m_UndoPos > 0) Then
        On Error Resume Next
        For lIdx = m_UndoPos + 1 To m_UndoMax
            Kill m_Temp & "\b" & m_AppID & Format(lIdx, "000") & ".dat"
        Next lIdx
        On Error GoTo 0
    End If

    If (m_UndoPos < m_UNDO_LEVELS) Then
        m_UndoPos = m_UndoPos + 1
        m_UndoMax = m_UndoPos
      Else
        Call pvRotateUndoFiles
    End If
    Call CheckUndoRedoState
End Sub

Private Sub pvRotateUndoFiles()

  Dim bOldName As String
  Dim bNewName As String
  Dim lIdx     As Long

    On Error Resume Next
    '-- Kill first
    Kill m_Temp & "\b" & m_AppID & "000.dat"
    '-- 'Rotate' the others (Move up 1)
    For lIdx = 1 To m_UNDO_LEVELS
        bOldName = m_Temp & "\b" & m_AppID & Format(lIdx - 0, "000") & ".dat"
        bNewName = m_Temp & "\b" & m_AppID & Format(lIdx - 1, "000") & ".dat"
        Name bOldName As bNewName
    Next lIdx
    On Error GoTo 0
End Sub

Private Function pvExtToBMP(ByVal sFilename As String) As String
    pvExtToBMP = Left$(sFilename, Len(sFilename) - 3) & "bmp"
End Function

Private Function pvGetExt(ByVal sFilename As String) As String
    pvGetExt = Right$(sFilename, 3)
End Function

Private Function pvCorrectExt(sFilename As String)
    If (Right$(sFilename, 4) <> Right$(m_FileExt, 4)) Then
        sFilename = sFilename & Right$(m_FileExt, 4)
    End If
End Function

'========================================================================================
' Public properties (settings)
'========================================================================================

Public Property Let LastFilterID(ByVal FilterID As fltIDCts)
    m_LastFilter = FilterID
End Property

Public Property Get LastFilename() As String
    LastFilename = m_LastFilename
End Property

Public Property Let LastFilename(ByVal sLastFilename As String)
    m_LastFilename = sLastFilename
End Property

Public Property Get LastPath() As String
    LastPath = m_LastPath
End Property

Public Property Let LastPath(ByVal sLastPath As String)
    m_LastPath = sLastPath
End Property

Public Property Get FileExt() As String
    FileExt = m_FileExt
End Property

Public Property Let FileExt(ByVal sFileExt As String)
    m_FileExt = sFileExt
End Property

Public Property Get DialogPreview() As Boolean
    DialogPreview = m_DialogPreview
End Property

Public Property Let DialogPreview(ByVal bShow As Boolean)
    m_DialogPreview = bShow
End Property

Public Property Get DialogFitMode() As Boolean
    DialogFitMode = m_DialogFitMode
End Property

Public Property Let DialogFitMode(ByVal bEnable As Boolean)
    m_DialogFitMode = bEnable
End Property

Public Property Get DialogJPEGquality() As Integer
    DialogJPEGquality = m_DialogJPEGquality
End Property

Public Property Let DialogJPEGquality(ByVal iValue As Integer)
    m_DialogJPEGquality = iValue
End Property
