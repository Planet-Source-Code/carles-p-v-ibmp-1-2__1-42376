VERSION 5.00
Begin VB.Form fProperties 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Properties"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5220
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fProperties.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   211
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   348
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox iPreview 
      BackColor       =   &H8000000C&
      ClipControls    =   0   'False
      Height          =   2310
      Left            =   120
      ScaleHeight     =   150
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   2310
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4035
      TabIndex        =   13
      Top             =   2685
      Width           =   1050
   End
   Begin VB.Label lblBitmapSizeV 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3720
      TabIndex        =   12
      Top             =   2100
      Width           =   1425
   End
   Begin VB.Label lblBitmapSize 
      Caption         =   "Bitmap size:"
      Height          =   225
      Left            =   2700
      TabIndex        =   6
      Top             =   2100
      Width           =   1005
   End
   Begin VB.Label lblColorsUsed 
      Caption         =   "Colors used:"
      Height          =   225
      Left            =   2700
      TabIndex        =   5
      Top             =   1425
      Width           =   1005
   End
   Begin VB.Label lblColorsUsedV 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3720
      TabIndex        =   11
      Top             =   1425
      Width           =   1425
   End
   Begin VB.Label lblPaletteV 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3720
      TabIndex        =   10
      Top             =   1125
      Width           =   1425
   End
   Begin VB.Label lblWidthV 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3720
      TabIndex        =   7
      Top             =   225
      Width           =   1425
   End
   Begin VB.Label lblHeightV 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3720
      TabIndex        =   8
      Top             =   525
      Width           =   1425
   End
   Begin VB.Label lblColorDepthV 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3720
      TabIndex        =   9
      Top             =   825
      Width           =   1425
   End
   Begin VB.Label lblPalette 
      Caption         =   "Palette:"
      Height          =   225
      Left            =   2700
      TabIndex        =   4
      Top             =   1125
      Width           =   690
   End
   Begin VB.Label lblColorDepth 
      Caption         =   "Color depth:"
      Height          =   225
      Left            =   2700
      TabIndex        =   3
      Top             =   825
      Width           =   1005
   End
   Begin VB.Label lblHeight 
      Caption         =   "Height:"
      Height          =   225
      Left            =   2700
      TabIndex        =   2
      Top             =   525
      Width           =   690
   End
   Begin VB.Label lblWidth 
      Caption         =   "Width:"
      Height          =   225
      Left            =   2700
      TabIndex        =   1
      Top             =   225
      Width           =   690
   End
End
Attribute VB_Name = "fProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'================================================
' fProperties form
' Last revision: 2003.11.02
'================================================

Option Explicit

Private bfx As Long, bfy As Long
Private bfW As Long, bfH As Long

Private Sub Form_Load()
    
    Screen.MousePointer = vbHourglass
    
    With fMain.Canvas.DIB
        
        '== Get fit info:
        Call .GetBestFitInfo(150, 150, bfx, bfy, bfW, bfH)
    
        '== Get DIB info:
        
        '-- Width:
        lblWidthV = .Width & " pixels"
        
        '-- Height:
        lblHeightV = .Height & " pixels"
        
        '-- Color depth:
        lblColorDepthV = fMain.DIBbpp & " bpp"
        
        '-- Palette:
        Select Case fMain.DIBbpp
            Case 1
                If (fMain.DIBPal.IsGreyScale) Then
                    lblPaletteV = "Black and White"
                  Else
                    lblPaletteV = "2 colors"
                End If
            Case 4
                If (fMain.DIBPal.IsGreyScale) Then
                    lblPaletteV = "16 greys"
                  Else
                    lblPaletteV = "16 colors"
                End If
            Case 8
                If (fMain.DIBPal.IsGreyScale) Then
                    lblPaletteV = "256 greys"
                  Else
                    lblPaletteV = "256 colors"
                End If
            Case 24
                lblPaletteV = "None"
        End Select
        
        '-- Colors used:
        lblColorsUsedV = fMain.DIBDither.CountColors(fMain.Canvas.DIB)
        
        '-- Bitmap size (mem.):
        lblBitmapSizeV = Format(pvCalcSize / 1024, "#0.0 Kb")
    End With
    
    Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Paint()
    Line (0, 170)-(ScaleWidth, 170), vb3DShadow
    Line (0, 171)-(ScaleWidth, 171), vb3DHighlight
End Sub

Private Sub iPreview_Paint()
    With fMain.Canvas.DIB
        Call .Stretch(iPreview.hDC, bfx, bfy, bfW, bfH, 0, 0, .Width, .Height)
    End With
End Sub

Private Sub cmdOK_Click()
    Call Unload(Me)
End Sub

'//

Private Function pvCalcSize() As Long
    
  Dim lHeaders As Long
  Dim lPalette As Long
  Dim lData    As Long
    
    With fMain
        lHeaders = 14 + 40
        lPalette = IIf(.DIBbpp <= 8, 4 * 2 ^ .DIBbpp, 0)
        lData = (((.Canvas.DIB.Width * .DIBbpp + 31) \ 32) * 4) * .Canvas.DIB.Height
    End With
    '-- Size: File header (14) + Bitmap header (40) + [Palette (1:8, 4:64, 8:1024)] + Bits (f(bpp))
    pvCalcSize = lHeaders + lPalette + lData
End Function
