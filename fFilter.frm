VERSION 5.00
Begin VB.Form fFilter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Filter preview"
   ClientHeight    =   6315
   ClientLeft      =   1170
   ClientTop       =   1455
   ClientWidth     =   5025
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fFilter.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   421
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   335
   ShowInTaskbar   =   0   'False
   Begin iBMP.ucToolbar Command 
      Height          =   240
      Left            =   3960
      Tag             =   "Preserve"
      Top             =   4260
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   423
   End
   Begin VB.CommandButton cmdZoomOut 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Symbol"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2190
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2505
      Width           =   225
   End
   Begin VB.CommandButton cmdZoomIn 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Symbol"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1905
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2505
      Width           =   225
   End
   Begin VB.CheckBox chkPickColor 
      Caption         =   "Pick color"
      Height          =   240
      Left            =   675
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2535
      Width           =   1020
   End
   Begin VB.CheckBox chkFit 
      Caption         =   "Fit"
      Height          =   240
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2535
      Width           =   690
   End
   Begin iBMP.ucCanvas iBfrCanvas 
      Height          =   2310
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   2310
      _ExtentX        =   4075
      _ExtentY        =   4075
   End
   Begin VB.Frame fraColorSelection 
      BorderStyle     =   0  'None
      Height          =   390
      Left            =   180
      TabIndex        =   26
      Top             =   5700
      Visible         =   0   'False
      Width           =   3345
      Begin VB.CheckBox chkLock 
         Caption         =   "Lock:"
         Height          =   225
         Left            =   105
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   60
         Width           =   750
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Index           =   1
         Left            =   2985
         TabIndex        =   30
         Top             =   0
         Width           =   330
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Index           =   0
         Left            =   1920
         TabIndex        =   27
         Top             =   0
         Width           =   330
      End
      Begin VB.Label lblColorT 
         Caption         =   "To color"
         Height          =   225
         Index           =   1
         Left            =   2370
         TabIndex        =   31
         Top             =   60
         Width           =   765
      End
      Begin VB.Label lblColorT 
         Caption         =   "From color"
         Height          =   225
         Index           =   0
         Left            =   1125
         TabIndex        =   29
         Top             =   60
         Width           =   870
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   3855
      TabIndex        =   24
      TabStop         =   0   'False
      Tag             =   "Preserve"
      ToolTipText     =   "Apply last preview [and Close dialog]"
      Top             =   5340
      Width           =   1035
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      Default         =   -1  'True
      Height          =   375
      Left            =   3855
      TabIndex        =   23
      Tag             =   "Preserve"
      ToolTipText     =   "Preview on main"
      Top             =   4890
      Width           =   1035
   End
   Begin VB.CheckBox chkNoClose 
      Caption         =   "&Don't close dialog after apply"
      Height          =   270
      Left            =   570
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3885
      Width           =   2445
   End
   Begin VB.ComboBox cbFilter 
      Height          =   315
      ItemData        =   "fFilter.frx":000C
      Left            =   570
      List            =   "fFilter.frx":0064
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   3135
      Width           =   3045
   End
   Begin VB.Frame fraParams 
      Height          =   1995
      Left            =   120
      TabIndex        =   10
      Top             =   4185
      Width           =   3510
      Begin VB.Frame fraParam 
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   0
         Left            =   135
         TabIndex        =   11
         Top             =   270
         Width           =   3330
         Begin VB.HScrollBar sbParam 
            Height          =   210
            Index           =   0
            LargeChange     =   5
            Left            =   975
            TabIndex        =   13
            Top             =   0
            Width           =   1785
         End
         Begin VB.Label lblParamValue 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            Height          =   195
            Index           =   0
            Left            =   2715
            TabIndex        =   14
            Top             =   0
            Width           =   480
         End
         Begin VB.Label lblParamName 
            Caption         =   "N.A."
            Height          =   195
            Index           =   0
            Left            =   30
            TabIndex        =   12
            Top             =   0
            Width           =   915
         End
      End
      Begin VB.Frame fraParam 
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   2
         Left            =   135
         TabIndex        =   19
         Top             =   990
         Width           =   3330
         Begin VB.HScrollBar sbParam 
            Height          =   210
            Index           =   2
            LargeChange     =   5
            Left            =   975
            TabIndex        =   21
            Top             =   0
            Width           =   1785
         End
         Begin VB.Label lblParamName 
            Caption         =   "N.A."
            Height          =   195
            Index           =   2
            Left            =   30
            TabIndex        =   20
            Top             =   0
            Width           =   915
         End
         Begin VB.Label lblParamValue 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            Height          =   195
            Index           =   2
            Left            =   2715
            TabIndex        =   22
            Top             =   0
            Width           =   480
         End
      End
      Begin VB.Frame fraParam 
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   1
         Left            =   135
         TabIndex        =   15
         Top             =   630
         Width           =   3330
         Begin VB.HScrollBar sbParam 
            Height          =   210
            Index           =   1
            LargeChange     =   5
            Left            =   975
            TabIndex        =   17
            Top             =   0
            Width           =   1785
         End
         Begin VB.Label lblParamValue 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            Height          =   195
            Index           =   1
            Left            =   2715
            TabIndex        =   18
            Top             =   0
            Width           =   480
         End
         Begin VB.Label lblParamName 
            Caption         =   "N.A."
            Height          =   195
            Index           =   1
            Left            =   30
            TabIndex        =   16
            Top             =   0
            Width           =   915
         End
      End
      Begin VB.Line lnSep 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   120
         X2              =   3405
         Y1              =   1380
         Y2              =   1380
      End
      Begin VB.Line lnSep 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   120
         X2              =   3405
         Y1              =   1365
         Y2              =   1365
      End
   End
   Begin VB.PictureBox iAft 
      BackColor       =   &H8000000C&
      ClipControls    =   0   'False
      Height          =   2310
      Left            =   2595
      ScaleHeight     =   150
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   2310
   End
   Begin VB.CheckBox chkResetValues 
      Caption         =   "&Reset values after preview"
      Height          =   270
      Left            =   570
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3585
      Width           =   2445
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   3855
      TabIndex        =   25
      TabStop         =   0   'False
      Tag             =   "Preserve"
      ToolTipText     =   "[Cancel last apply and] Close dialog"
      Top             =   5805
      Width           =   1035
   End
   Begin VB.Label lblComboFilter 
      Caption         =   "Filter"
      Height          =   240
      Left            =   135
      TabIndex        =   6
      Top             =   3195
      Width           =   585
   End
End
Attribute VB_Name = "fFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'================================================
' Filter preview form
' Last revision: 2003.11.02
'================================================
' Adding a new filter:
'  1. Add FilterID constant
'  2. Add filter name in list box
'  3. Add call in 'pvApplyFilter' sub.
'  4. Define params. in 'pvInitializeFilter' sub.
'================================================

Option Explicit

'-- API:

Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

'-- Public enums.:
Public Enum fltIDCts ' Alphabetical order
    [<None>] = 0
    [fltBlur]
    [fltBrightness]
    [fltColorize]
    [fltContour]
    [fltContrast]
    [fltDespeckle]
    [fltDespeckleMore]
    [fltDiffuse]
    [fltDilate]
    [fltEmboss]
    [fltErode]
    [fltGreys]
    [fltNegative]
    [fltNoise]
    [fltOutline]
    [fltPixelize]
    [fltRelieve]
    [fltReplaceHS]
    [fltReplaceL]
    [fltRGBLevels]
    [fltSaltAndPepperRemoval]
    [fltSaturation]
    [fltScanlines]
    [fltSepia]
    [fltSharpen]
    [fltShift]
    [fltSoften]
End Enum

'-- Private objects:
Private WithEvents oDIBAftPreview As cDIBFilter
Attribute oDIBAftPreview.VB_VarHelpID = -1
Private oAftDIB                   As New cDIB

'-- Private Variables:
Private m_FilterID    As fltIDCts
Private m_Initialized As Boolean
Private m_Previewing  As Boolean

Private bfx As Long, bfy As Long
Private bfW As Long, bfH As Long

'========================================================================================
' <Before> view control
'========================================================================================

Private Sub Form_Load()

    '-- Load settings
    Call mSettings.LoadFilterSettings
    
    '-- Resize combo
    With cbFilter
        Call MoveWindow(.hWnd, .Left, .Top, .Width, ScaleHeight - .Top, 0)
    End With
    
    '-- Initialize DIB filter
    Set oDIBAftPreview = New cDIBFilter
    '-- Update Undo/Redo state
    Call fMain.CheckUndoRedoState
    
    '-- Intialize Undo/Redo toolbar
    Call Command.BuildToolbar(LoadResPicture("BITMAP_TBUNDOREDO", vbResBitmap), &HFF00FF, 16, "N|N")
    Call Command.SetTooltips("Undo|Redo")
    Call Command.Refresh
    
    '-- Removing border enhance
    Call mMisc.RemoveButtonBorderEnhance(cmdZoomIn)
    Call mMisc.RemoveButtonBorderEnhance(cmdZoomOut)
    
    '-- Hook scroll bars (correct 'incorrect' background refresh)
    Call mHook.HookFilterSBs
End Sub

Public Sub Initialize(Optional ByVal FilterID As fltIDCts = 0)
    
    '-- Create temp. DIB
    With fMain.Canvas.DIB
        Call iBfrCanvas.DIB.Create(.Width, .Height)
        Call iBfrCanvas.DIB.LoadBlt(.hDC)
        Call iBfrCanvas.Resize
    End With
    
    '-- Get <After> best fit dim. and pos.
    Call iBfrCanvas.DIB.GetBestFitInfo(iAft.ScaleWidth, iAft.ScaleHeight, bfx, bfy, bfW, bfH)
    
    '-- Create <After> DIB
    Call oAftDIB.Create(bfW, bfH)
    If (FilterID = [<None>]) Then
        Call oAftDIB.LoadDIBBlt(iBfrCanvas.DIB)
    End If
    
    '-- Initialize filter (force selection)
    cbFilter.ListIndex = FilterID
End Sub

Private Sub Form_Paint()
    Line (8, 195)-(ScaleWidth - 9, 195), vb3DShadow
    Line (8, 196)-(ScaleWidth - 9, 196), vb3DHighlight
End Sub

Private Sub iAft_Paint()
    Call oAftDIB.Paint(iAft.hDC, bfx, bfy)
End Sub

Private Sub Form_Unload(Cancel As Integer)

    '-- Restore ?
    If (cmdApply.Enabled) Then
        With fMain.Canvas
            Call .DIB.LoadBlt(iBfrCanvas.DIB.hDC)
            Call .Repaint
        End With
        Call fPanView.Repaint
    End If
    
    '-- Save settings
    Call mSettings.SaveFilterSettings
    
    '-- Destroy DIBs
    Call iBfrCanvas.DIB.Destroy
    Set oAftDIB = Nothing
End Sub

'========================================================================================
' <Before> view control
'========================================================================================

Private Sub chkFit_Click()
    iBfrCanvas.FitMode = CBool(chkFit)
    Call iBfrCanvas.Resize
    cmdZoomIn.Enabled = Not CBool(chkFit)
    cmdZoomOut.Enabled = Not CBool(chkFit)
End Sub

Private Sub chkPickColor_Click()
    If (chkPickColor) Then
        iBfrCanvas.WorkMode = [cnvPickColorMode]
      Else
        iBfrCanvas.WorkMode = [cnvScrollMode]
    End If
End Sub

Private Sub cmdZoomIn_Click()
    With iBfrCanvas
        If (.Zoom < 10) Then
            .Zoom = .Zoom + 1
        End If
        Call .Resize
    End With
End Sub

Private Sub cmdZoomOut_Click()
    With iBfrCanvas
        .Zoom = .Zoom - 1
        Call .Resize
    End With
End Sub

'========================================================================================
' Commands
'========================================================================================

Private Sub cbFilter_Click()
'-- Filter has been selected

    '-- Current filter ID
    m_FilterID = cbFilter.ListIndex
    '-- Initialize filter
    Call pvInitializeFilter
    
    '-- Enable/Disable <Preview>/<Apply> buttons
    cmdPreview.Enabled = (m_FilterID > 0)
    cmdApply.Enabled = False
    
    '-- Save to Main prop.
    fMain.LastFilterID = m_FilterID
End Sub

Private Sub sbParam_Change(Index As Integer)
'-- Param. has been changed
    
    '-- Refresh value
    With lblParamValue(Index)
        .Caption = sbParam(Index)
        Call .Refresh
    End With
    
    '-- Restore <after> preview, apply filter and refresh
    Call oAftDIB.LoadDIBBlt(iBfrCanvas.DIB)
    Call pvApplyFilter(oAftDIB)
    Call iAft_Paint
    
    '-- Enable <Preview> button / Disable <Apply> button
    cmdPreview.Enabled = (m_FilterID > 0)
    cmdApply.Enabled = False
End Sub

Private Sub sbParam_Scroll(Index As Integer)
    Call sbParam_Change(Index)
End Sub

Private Sub cmdPreview_Click()
'-- Preview on main
    
    '-- Preview filter and refresh
    m_Previewing = True
        With fMain.Canvas
            Call .DIB.LoadBlt(iBfrCanvas.DIB.hDC)
            Call pvApplyFilter(.DIB)
            Call .Repaint
        End With
    m_Previewing = False
    
    '-- Reset values
    If (chkResetValues) Then
        Call pvInitializeFilter
    End If

    '-- Disable <Preview> button / Enable <Apply> button
    cmdPreview.Enabled = False
    cmdApply.Enabled = True
    Call cmdApply.SetFocus
End Sub

Private Sub cmdApply_Click()
'-- Apply [and Close]
    
    '-- Save Undo DIB
    Call fMain.DIBFilter_ProgressEnd
    
    '-- Close dialog ?
    cmdApply.Enabled = False
    If (chkNoClose) Then
        '-- Update Preview DIBs
        Call iBfrCanvas.DIB.LoadBlt(fMain.Canvas.DIB.hDC) ' <Before>
        Call iBfrCanvas.Repaint
        Call pvInitializeFilter(0)                        ' <After>
        '-- Enable <Preview> button / Disable <Apply> button
        cmdPreview.Enabled = True
      Else
        Call Unload(Me)
    End If
End Sub

Private Sub cmdClose_Click()
'-- Cancel and Close
    Call Unload(Me)
End Sub

'========================================================================================
' Applying filter
'========================================================================================

Private Sub pvApplyFilter(DIB As cDIB)

    If (m_Initialized) Then
    
        With oDIBAftPreview
        
            Select Case m_FilterID
            
                Case [fltBlur]
                    Call .Blur(DIB, sbParam(0) + 75)
              
                Case [fltBrightness]
                    Call .ShiftRGB(DIB, sbParam(0), sbParam(0), sbParam(0))
              
                Case [fltColorize]
                    Call .Colorize(DIB, (RotateH40(2.4 * sbParam(0))) / 40, sbParam(1) / 100, sbParam(2) / 100)
              
                Case [fltContour]
                    Call .Contour(DIB)
              
                Case [fltContrast]
                    Call .Contrast(DIB, sbParam(0))
                
                Case [fltDespeckle]
                    Call .Despeckle(DIB)
                
                Case [fltDespeckleMore]
                    Call .DespeckleMore(DIB)
                
                Case [fltDiffuse]
                    Call .Diffuse(DIB, sbParam(0), sbParam(1), sbParam(2))
                
                Case [fltDilate]
                    Call .RankFilterMaximum(DIB)
              
                Case [fltEmboss]
                    Call .Emboss(DIB, sbParam(0), True)
                
                Case [fltErode]
                    Call .RankFilterMinimum(DIB)
              
                Case [fltGreys]
                    Call .Greys(DIB)
              
                Case [fltNegative]
                    Call .Negative(DIB)
              
                Case [fltNoise]
                    Call .Noise(DIB, sbParam(0), sbParam(1))
              
                Case [fltOutline]
                    Call .Outline(DIB, sbParam(0), sbParam(1))
              
                Case [fltPixelize]
                    Call .Pixelize(DIB, sbParam(0))
              
                Case [fltRelieve]
                    Call .Relieve(DIB, sbParam(0))
              
                Case [fltReplaceHS]
                    Call .ReplaceHS(DIB, lblColor(0).BackColor, lblColor(1).BackColor, sbParam(0), sbParam(1), sbParam(2))
              
                Case [fltReplaceL]
                    Call .ReplaceL(DIB, lblColor(0).BackColor, lblColor(1).BackColor, sbParam(0), sbParam(1))
              
                Case [fltRGBLevels]
                    Call .ShiftRGB(DIB, sbParam(0), sbParam(1), sbParam(2))
              
                Case [fltSaltAndPepperRemoval]
                    Call .SaltAndPepperRemoval(DIB, 255 - sbParam(0))
                
                Case [fltSaturation]
                    Call .Saturation(DIB, sbParam(0))
                
                Case [fltScanlines]
                    Call .Scanlines(DIB, sbParam(0), sbParam(1))
                
                Case [fltSepia]
                    Call .Colorize(DIB, 0.5, 0.25)
              
                Case [fltSharpen]
                    Call .Sharpen(DIB, sbParam(0) + 75)
              
                Case [fltShift]
                    Call .Shift(DIB, lblColor(0).BackColor, sbParam(0), sbParam(1))
            
                Case [fltSoften]
                    Call .Soften(DIB, sbParam(0) + 75)
            End Select
        End With
    End If
End Sub

Private Sub oDIBAftPreview_Progress(ByVal p As Long)
    If (m_Previewing) Then Call fMain.DIBFilter_Progress(p)
End Sub
Private Sub oDIBAftPreview_ProgressEnd()
    If (m_Previewing) Then Call fMain.DIBFilter_ProgressEnd
End Sub

'========================================================================================
' Initializing filter
'========================================================================================

Private Sub pvInitializeFilter(Optional ByVal InitValues As Boolean = -1)
    
    If (InitValues) Then
    
        m_Initialized = False
        
        '-- Hide param. controls
        fraParams.Visible = False
        fraParam(0).Visible = False
        fraParam(1).Visible = False
        fraParam(2).Visible = False
        fraColorSelection.Visible = False
        
        Select Case m_FilterID
        
            Case [fltBlur]
                Call pvSetParam(0, "Level", 1, 25, 1)
        
            Case [fltBrightness]
                Call pvSetParam(0, "Level", -255, 255, 0)
            
            Case [fltColorize]
                Call pvSetParam(0, "Hue", 0, 100, 0)
                Call pvSetParam(1, "Saturation", 0, 100, 50)
                Call pvSetParam(2, "Luminosity", 0, 100, 100)
          
            Case [fltContour]
          
            Case [fltContrast]
                Call pvSetParam(0, "Level", -100, 250, 0)
            
            Case [fltDiffuse]
                Call pvSetParam(0, "Horizontal factor", 0, 100, 0)
                Call pvSetParam(1, "Vertical factor", 0, 100, 0)
                Call pvSetParam(2, "Step", 1, 10, 1)
            
            Case [fltDilate]
          
            Case [fltDespeckle]
          
            Case [fltDespeckleMore]
                  
            Case [fltEmboss]
                Call pvSetParam(0, "Level", 0, 200, 0)
            
            Case [fltErode]
            
            Case [fltGreys]
            
            Case [fltNegative]
            
            Case [fltNoise]
                Call pvSetParam(0, "Amount", 0, 200, 0)
                Call pvSetParam(1, "Step", 1, 10, 1)
          
            Case [fltOutline]
                Call pvSetParam(0, "Level", -100, 100, 0)
                Call pvSetParam(1, "Shift", -50, 50, 0)
            
            Case [fltPixelize]
                Call pvSetParam(0, "Pixel size", 1, 25, 1)
          
            Case [fltRelieve]
                Call pvSetParam(0, "Level", 1, 100, 1)
            
            Case [fltReplaceHS]
                Call pvSetParam(0, "Extent", 0, 255, 0)
                Call pvSetParam(1, "H Tolerance", 0, 100, 0)
                Call pvSetParam(2, "S Tolerance", 0, 100, 0)
                Call pvColorSelection
            
            Case [fltReplaceL]
                Call pvSetParam(0, "Extent", 0, 255, 0)
                Call pvSetParam(1, "L Tolerance", 0, 100, 0)
                Call pvColorSelection
          
            Case [fltRGBLevels]
                Call pvSetParam(0, "Red", -255, 255, 0)
                Call pvSetParam(1, "Green", -255, 255, 0)
                Call pvSetParam(2, "Blue", -255, 255, 0)
          
            Case [fltSaltAndPepperRemoval]
                Call pvSetParam(0, "Threshold", 0, 255, 0)
      
            Case [fltSaturation]
                Call pvSetParam(0, "Level", -100, 250, 0)
        
            Case [fltScanlines]
                Call pvSetParam(0, "Black point", 0, 100, 0)
                Call pvSetParam(1, "White point", 0, 100, 0)
            
            Case [fltSepia]
          
            Case [fltSharpen]
                Call pvSetParam(0, "Level", 1, 25, 1)
            
            Case [fltShift]
                Call pvSetParam(0, "Amount", -255, 255, 0)
                Call pvSetParam(1, "Extent", 0, 255, 0)
                Call pvColorSelection(1)
            
            Case [fltSoften]
                Call pvSetParam(0, "Level", 1, 25, 1)
        End Select
    
        m_Initialized = True
        fraParams.Visible = True
    End If
    
    '-- Restore <after> preview, apply filter and refresh
    Call oAftDIB.LoadDIBBlt(iBfrCanvas.DIB)
    Call pvApplyFilter(oAftDIB)
    Call iAft_Paint
End Sub

Private Sub pvSetParam(ByVal Index As Integer, ByVal pInfo As String, ByVal pMin As Long, ByVal pMax As Long, ByVal pStart As Long)

  Dim tPPY As Long
  Dim lIdx As Long
    
    tPPY = Screen.TwipsPerPixelY
    
    '-- Set param. name
    lblParamName(Index) = pInfo
    '-- Set param. interval
    With sbParam(Index)
        .Min = pMin
        .Max = pMax
        .Value = pStart
    End With
    
    '-- Relocate scroll bars
    Select Case Index
        Case 0
            fraParam(0).Top = 42 * tPPY
        Case 1
            fraParam(0).Top = 29 * tPPY
            fraParam(1).Top = 51 * tPPY
        Case 2
            fraParam(0).Top = 20 * tPPY
            fraParam(1).Top = 42 * tPPY
            fraParam(2).Top = 64 * tPPY
    End Select
    
    '-- Show param. control
    For lIdx = 0 To Index
        fraParam(lIdx).Visible = True
    Next lIdx
End Sub

Private Sub pvColorSelection(Optional ByVal DisableColor As Integer = -1)

    '-- Reset all
    lblColor(0).BackColor = vbBlack
    lblColor(1).BackColor = vbBlack
    lblColor(0).Enabled = True
    lblColor(1).Enabled = True
    lblColorT(0).Enabled = True
    lblColorT(1).Enabled = True
    
    '-- Disable/Enable Color selectors
    If (DisableColor > -1) Then
        lblColor(DisableColor).Enabled = False
        lblColorT(DisableColor).Enabled = False
    End If
    fraColorSelection.Visible = True
End Sub

'========================================================================================
' Picking color
'========================================================================================

Private Sub lblColor_Click(Index As Integer)
  
  Dim lRet As Long
  
    '-- Pick from color dialog
    lRet = SelectColor(Me.hWnd, lblColor(Index).BackColor, -1)
    If (lRet <> -1) Then
        lblColor(Index).BackColor = lRet
    End If
    
    '-- Force preview
    Call sbParam_Change(0)
End Sub

Private Sub iBfrCanvas_MouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
    '-- Force pick
    Call iBfrCanvas_MouseMove(Button, Shift, x, y)
End Sub

Private Sub iBfrCanvas_MouseMove(Button As Integer, Shift As Integer, x As Long, y As Long)

  Dim lColor As Long
  Dim lIndex As Long
    
    lIndex = IIf(Button = vbLeftButton, 0, 1)
    
    If (Button And CBool(chkPickColor)) Then
        
        '-- Get pixel color
        lColor = GetPixel(iBfrCanvas.DIB.hDC, x, y)
        If (lblColor(lIndex).Enabled) Then
            lblColor(lIndex).BackColor = IIf(lColor > -1, lColor, 0)
            Call lblColor(lIndex).Refresh
        End If
        '-- Lock ?
        If (chkLock And lblColor(1).Enabled) Then
            lblColor(1).BackColor = lblColor(0).BackColor 'From->To
            Call lblColor(1).Refresh
        End If
        '-- Apply filter (Force param. change)
        Call sbParam_Change(0)
    End If
End Sub

'========================================================================================
' Undo/Redo
'========================================================================================

Private Sub Command_ButtonClick(ByVal Index As Long, ByVal xLeft As Long, ByVal yTop As Long)

  Dim bFilterDisabled As Boolean
  Dim oControl        As Control

    '-- Call main's sub
    Select Case Index
        Case 1: Call fMain.Undo
        Case 2: Call fMain.Redo
    End Select

    '-- Update previews
    With fMain.Canvas.DIB '< Before > View
        Call iBfrCanvas.DIB.Create(.Width, .Height)
        Call iBfrCanvas.DIB.LoadBlt(.hDC)
        Call iBfrCanvas.Resize
    End With
    Call iAft.Cls          '<After> view (Clear and preview)
    Call pvInitializeFilter

    '-- Disable filter (<= 8bpp images)
    bFilterDisabled = (fMain.DIBbpp <= 8)
    '-- controls...
    On Error Resume Next
    For Each oControl In Controls
        With oControl
            If (.Tag <> "Preserve") Then .Enabled = Not bFilterDisabled
        End With
    Next
    On Error GoTo 0

    '-- Enable <Preview> button / Disable <Apply> button
    cmdPreview.Enabled = (Not bFilterDisabled And m_FilterID > 0)
    cmdApply.Enabled = False
End Sub

'========================================================================================
' Properties
'========================================================================================

Public Property Get Previewing() As Boolean
    Previewing = m_Previewing
End Property
