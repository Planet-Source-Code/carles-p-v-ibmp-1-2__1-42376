VERSION 5.00
Begin VB.Form fTexturize 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Texturize"
   ClientHeight    =   5235
   ClientLeft      =   1170
   ClientTop       =   1455
   ClientWidth     =   4470
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
   Icon            =   "fTexturize.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   349
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   298
   ShowInTaskbar   =   0   'False
   Begin iBMP.ucToolbar Command 
      Height          =   360
      Left            =   3405
      Tag             =   "Preserve"
      Top             =   3300
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   635
   End
   Begin VB.CheckBox chkNoClose 
      Caption         =   "&Don't close dialog after apply"
      Height          =   270
      Left            =   120
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3345
      Width           =   2445
   End
   Begin VB.CommandButton cmdRotate90 
      Caption         =   "+90º"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3825
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Rotate texture 90º"
      Top             =   2730
      Width           =   510
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3315
      TabIndex        =   14
      TabStop         =   0   'False
      Tag             =   "Preserve"
      ToolTipText     =   "Cancel last apply and Close dialog"
      Top             =   4725
      Width           =   1035
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   3315
      TabIndex        =   12
      ToolTipText     =   "Preview on main"
      Top             =   3810
      Width           =   1035
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3315
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Apply last preview [and Close dialog]"
      Top             =   4260
      Width           =   1035
   End
   Begin VB.TextBox txtFolder 
      Height          =   285
      Left            =   135
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   360
      Width           =   3915
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "…"
      Height          =   285
      Left            =   4065
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Browse..."
      Top             =   360
      Width           =   270
   End
   Begin VB.CheckBox chkFitMode 
      Caption         =   "&Fit "
      Enabled         =   0   'False
      Height          =   255
      Left            =   2460
      TabIndex        =   5
      TabStop         =   0   'False
      Tag             =   "Preserve"
      Top             =   2745
      Width           =   705
   End
   Begin iBMP.ucCanvas Preview 
      Height          =   1875
      Left            =   2460
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   765
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   3307
   End
   Begin VB.FileListBox flTextures 
      Height          =   2235
      Left            =   120
      Pattern         =   "*.bmp;*.dib"
      TabIndex        =   3
      Top             =   750
      Width           =   2130
   End
   Begin VB.Frame fraOptions 
      Caption         =   "Options"
      Height          =   1380
      Left            =   120
      TabIndex        =   8
      Top             =   3720
      Width           =   2880
      Begin VB.CheckBox chkInvertTexture 
         Caption         =   "&Invert texture"
         Height          =   300
         Left            =   270
         TabIndex        =   15
         Top             =   855
         Width           =   1545
      End
      Begin VB.HScrollBar sbWeight 
         Height          =   225
         LargeChange     =   5
         Left            =   870
         Max             =   100
         TabIndex        =   10
         Top             =   420
         Value           =   25
         Width           =   1230
      End
      Begin VB.Label lblWeight 
         Caption         =   "Weight"
         Height          =   240
         Left            =   270
         TabIndex        =   9
         Top             =   420
         Width           =   705
      End
      Begin VB.Label lblWeightV 
         Alignment       =   1  'Right Justify
         Caption         =   "25%"
         Height          =   195
         Left            =   1980
         TabIndex        =   11
         ToolTipText     =   "Texture weight (pressure)"
         Top             =   435
         Width           =   645
      End
   End
   Begin VB.Label lblFolderTitle 
      Caption         =   "Current folder"
      Height          =   240
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   1080
   End
End
Attribute VB_Name = "fTexturize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private TmpDIB        As New cDIB
Private m_Previewing  As Boolean

Private Sub Form_Load()

    '-- Load settings
    Call mSettings.LoadTexturizeSettings

    '-- Show current folder
    txtFolder = flTextures.Path
    '-- Well...
    Call mMisc.RemoveButtonBorderEnhance(cmdBrowse)
    Call mMisc.RemoveButtonBorderEnhance(cmdRotate90)
    
    '-- Intialize Undo/Redo toolbar
    Call Command.BuildToolbar(LoadResPicture("BITMAP_TBUNDOREDO", vbResBitmap), &HFF00FF, 16, "N|N")
    Call Command.SetTooltips("Undo|Redo")
    Call Command.Refresh
    
    '-- Update Undo/Redo state
    Call fMain.CheckUndoRedoState
    
    '-- Hook scroll bar (correct 'incorrect' background refresh)
    Call mHook.HookTexturizeSB
End Sub

Private Sub Form_Activate()

    '-- Create temp. DIB
    With fMain.Canvas.DIB
        Call TmpDIB.Create(.Width, .Height)
        Call TmpDIB.LoadBlt(.hDC)
    End With
End Sub

Private Sub Form_Paint()
    Line (8, 210)-(ScaleWidth - 9, 210), vb3DShadow
    Line (8, 211)-(ScaleWidth - 9, 211), vb3DHighlight
End Sub

'//

Private Sub cmdBrowse_Click()
 
  Dim sRet As String
    
    sRet = BrowseFolder(Me, "Select source folder")
    If (Len(sRet)) Then
        On Error GoTo ErrPath
        flTextures.Path = sRet
        txtFolder = sRet
    End If
    Exit Sub
    
ErrPath:
    Call MsgBox("Unable to access path:" & vbCrLf & sRet, vbExclamation)
End Sub

Private Sub chkFitMode_Click()
    Preview.FitMode = CBool(chkFitMode)
    Call Preview.Resize
End Sub

Private Sub flTextures_PathChange()

    '-- Clear texture DIB
    Call Preview.DIB.Destroy
    Call Preview.Resize

    '-- Folder exists ?
    If (FileFound(flTextures.Path) = 0 And Len(flTextures.Path) > 3) Then
        Call MsgBox("Last folder not found. Changing current folder to application folder", vbExclamation)
        flTextures.Path = AppPath
        txtFolder = AppPath
    End If
    
    '-- Files found ?
    If (flTextures.ListCount = 0) Then
        chkFitMode.Enabled = False
        cmdRotate90.Enabled = False
        cmdPreview.Enabled = False
    End If
End Sub

Private Sub flTextures_Click()

  Dim tmpPal    As New cDIBPal
  Dim tmpDither As New cDIBDither
  Dim sFilename As String
  
    chkFitMode.Enabled = True
    cmdRotate90.Enabled = True
    cmdPreview.Enabled = True

    sFilename = flTextures.Path & IIf(Len(flTextures.Path) > 3, "\", vbNullString) & flTextures.Filename
    
    On Error GoTo ErrLoad
    Call Preview.DIB.CreateFromStdPicture(LoadPicture(sFilename), tmpPal, tmpDither)
    Call Preview.Resize
    Exit Sub
    
ErrLoad:
    Call Preview.DIB.Destroy
    Call Preview.Resize
End Sub

Private Sub sbWeight_Change()
    lblWeightV = sbWeight & "%"
    cmdPreview.Enabled = (flTextures.ListIndex <> -1)
    cmdApply.Enabled = False
End Sub

Private Sub sbWeight_Scroll()
    Call sbWeight_Change
End Sub

Private Sub chkInvertTexture_Click()
    
  Static lstMode As Integer
    
    If (lstMode <> chkInvertTexture) Then
        cmdPreview.Enabled = (flTextures.ListIndex <> -1)
        cmdApply.Enabled = False
    End If
    lstMode = chkInvertTexture
End Sub

Private Sub cmdRotate90_Click()
    
    '-- Rotate +90º
    Screen.MousePointer = vbArrowHourglass
    Call Preview.DIB.Orientation(True, False, False)
    Call Preview.Resize
    Screen.MousePointer = vbDefault
    '-- Enable Preview
    cmdPreview.Enabled = True
    cmdApply.Enabled = False
End Sub

'//

Private Sub cmdPreview_Click()
'-- Preview on main
    
    '-- Preview filter and refresh
    m_Previewing = True
        With fMain.Canvas
            Call .DIB.LoadBlt(TmpDIB.hDC)
            Call pvPreview
        End With
    m_Previewing = False
    
    '-- Disable <Preview> button / Enable <Apply> button
    cmdPreview.Enabled = False
    cmdApply.Enabled = True
    Call cmdApply.SetFocus
End Sub

Private Sub cmdApply_Click()
'-- Apply [and Close]
    
     '-- Save Undo DIB (Apply)
    Call fMain.DIBFilter_ProgressEnd
    
    '-- Close dialog ?
    cmdApply.Enabled = False
    If (chkNoClose) Then
        '-- Update temp. DIB
        Call TmpDIB.LoadBlt(fMain.Canvas.DIB.hDC)
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

'//

Private Sub pvPreview()
    
  Dim BuffDIB       As New cDIB
  Dim BuffDIBFilter As New cDIBFilter
    
    With fMain
    
        '-- Create a temp. copy (texture)
        Call BuffDIB.Create(Preview.DIB.Width, Preview.DIB.Height)
        Call BuffDIB.LoadBlt(Preview.DIB.hDC)
         
        '-- Create height-map
        If (chkInvertTexture = 0) Then
            Call BuffDIBFilter.Emboss(BuffDIB, sbWeight, True)
          Else
            Call BuffDIBFilter.Engrave(BuffDIB, sbWeight, True)
        End If
         
        '-- Texturize
        Call .DIBFilter.Texturize(.Canvas.DIB, BuffDIB)
        Call .Canvas.Repaint
    End With
End Sub

'-- Undo/Redo
Private Sub Command_ButtonClick(ByVal Index As Long, ByVal xLeft As Long, ByVal yTop As Long)
    
  Dim bFilterDisabled As Boolean
  Dim oControl        As Control
  
        '-- Call main's sub
    Select Case Index
        Case 1: Call fMain.Undo
        Case 2: Call fMain.Redo
    End Select
    
    '-- Update temp. DIB
    With fMain.Canvas.DIB
        Call TmpDIB.Create(.Width, .Height)
        Call TmpDIB.LoadBlt(.hDC)
    End With
    
    '-- Disable texturize (<= 8bpp images)
    bFilterDisabled = (fMain.DIBbpp <= 8)
    
    On Error Resume Next
    For Each oControl In Controls
        With oControl
            If (.Tag <> "Preserve") Then .Enabled = Not bFilterDisabled
        End With
    Next
    On Error GoTo 0
    
    '-- Enable <Preview> button / Disable <Apply> button
    cmdPreview.Enabled = (Not bFilterDisabled And flTextures.ListIndex <> -1)
    cmdApply.Enabled = False
End Sub

'//

Public Property Get Previewing() As Boolean
    Previewing = m_Previewing
End Property

'//

Private Sub Form_Unload(Cancel As Integer)
    
    '-- Restore ?
    If (cmdApply.Enabled) Then
        With fMain.Canvas
            Call .DIB.LoadBlt(TmpDIB.hDC)
            Call .Repaint
        End With
        Call fPanView.Repaint
    End If
    
    '-- Destroy Temp. DIBs
    Set TmpDIB = Nothing
    Call Preview.DIB.Destroy
    
    '-- Save settings
    Call mSettings.SaveTexturizeSettings
End Sub
