VERSION 5.00
Begin VB.Form fPreview 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Filter preview"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5085
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fPreview.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   307
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   339
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraParam 
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   0
      Left            =   255
      TabIndex        =   8
      Top             =   3315
      Width           =   3330
      Begin VB.HScrollBar sbParam 
         Height          =   210
         Index           =   0
         Left            =   1035
         TabIndex        =   9
         Top             =   0
         Width           =   1590
      End
      Begin VB.Label lblParamName 
         Caption         =   "Param. 1"
         Height          =   195
         Index           =   0
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   840
      End
      Begin VB.Label lblParamValue 
         Alignment       =   1  'Right Justify
         Caption         =   "Val. 1"
         Height          =   195
         Index           =   0
         Left            =   2715
         TabIndex        =   10
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   3900
      TabIndex        =   5
      Top             =   3600
      Width           =   1050
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   570
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   2640
      Width           =   3045
   End
   Begin VB.Frame fraParams 
      Height          =   1395
      Left            =   120
      TabIndex        =   3
      Top             =   3060
      Width           =   3495
      Begin VB.Frame fraParam 
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   2
         Left            =   135
         TabIndex        =   16
         Top             =   1005
         Width           =   3330
         Begin VB.HScrollBar sbParam 
            Height          =   210
            Index           =   2
            Left            =   1035
            TabIndex        =   17
            Top             =   0
            Width           =   1590
         End
         Begin VB.Label lblParamValue 
            Alignment       =   1  'Right Justify
            Caption         =   "Val. 1"
            Height          =   195
            Index           =   2
            Left            =   2715
            TabIndex        =   19
            Top             =   0
            Width           =   480
         End
         Begin VB.Label lblParamName 
            Caption         =   "Param. 1"
            Height          =   195
            Index           =   2
            Left            =   0
            TabIndex        =   18
            Top             =   0
            Width           =   840
         End
      End
      Begin VB.Frame fraParam 
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   1
         Left            =   135
         TabIndex        =   12
         Top             =   630
         Width           =   3330
         Begin VB.HScrollBar sbParam 
            Height          =   210
            Index           =   1
            Left            =   1035
            TabIndex        =   13
            Top             =   0
            Width           =   1590
         End
         Begin VB.Label lblParamValue 
            Alignment       =   1  'Right Justify
            Caption         =   "Val. 1"
            Height          =   195
            Index           =   1
            Left            =   2715
            TabIndex        =   15
            Top             =   0
            Width           =   480
         End
         Begin VB.Label lblParamName 
            Caption         =   "Param. 1"
            Height          =   195
            Index           =   1
            Left            =   0
            TabIndex        =   14
            Top             =   0
            Width           =   840
         End
      End
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      Default         =   -1  'True
      Height          =   375
      Left            =   3900
      TabIndex        =   4
      Top             =   2640
      Width           =   1050
   End
   Begin VB.CommandButton cmdRestore 
      Cancel          =   -1  'True
      Caption         =   "Restore"
      Height          =   375
      Left            =   3900
      TabIndex        =   6
      Top             =   4080
      Width           =   1050
   End
   Begin VB.PictureBox iDst 
      BackColor       =   &H8000000C&
      ClipControls    =   0   'False
      Height          =   2310
      Left            =   2640
      ScaleHeight     =   150
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   2310
   End
   Begin VB.PictureBox iSrc 
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
   Begin VB.Label Label1 
      Caption         =   "Filter"
      Height          =   240
      Left            =   135
      TabIndex        =   7
      Top             =   2685
      Width           =   585
   End
End
Attribute VB_Name = "fPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'//
'// Filter preview form
'//
'// To add new filter:
'// 1. Add FilterID constant
'// 2. Add call in 'ApplyFilter' sub
'// 3. Define options in 'SetupFilter' function
'//

Option Explicit

Private DIBFilterPreview As New cDIBFilter

Public Enum fltIDCts
    [fltColorize]
    
End Enum

Public FilterID As fltIDCts

Private SrcDIB As New cDIB
Private DstDIB As New cDIB

Private bfW As Long, bfH As Long
Private bfx As Long, bfy As Long

Private sbInitialized As Boolean



Public Sub Initialize(DIB As cDIB)

    '-- Get dest. best fit dim. and pos.
    DIB.GetBestFitInfo 150, 150, bfx, bfy, bfW, bfH
    '-- Clear both previews
    iSrc.Cls
    iDst.Cls
    '-- Create previews
    SrcDIB.Create bfW, bfH
    SrcDIB.LoadDIBBlt DIB
    DstDIB.Create bfW, bfH
    DstDIB.LoadBlt SrcDIB.hDIBDC
    
    iSrc_Paint
    iDst_Paint
End Sub




Private Sub iSrc_Paint()
    SrcDIB.Paint iSrc.hdc, bfx, bfy
End Sub

Private Sub iDst_Paint()
    DstDIB.Paint iDst.hdc, bfx, bfy
End Sub







Private Sub sbParam_Change(Index As Integer)

'    fltOK = 0
'    fltPreview = 0
'    fltPreviewed = 0
        DstDIB.LoadBlt SrcDIB.hDIBDC

    pvApplyFilter FilterID, DstDIB
    iDst_Paint
End Sub

Private Sub sbParam_Scroll(Index As Integer)
    sbParam_Change 0
End Sub

Private Sub pvApplyFilter(ByVal fltID As Long, DIB As cDIB)


'    If (fltOK) Then
'        Hide
'        DoEvents
'    End If

    'If (fltPreviewed = 0) Then
            Select Case fltID
                '// Color [Main]
              Case [fltColorize]
                DIBFilterPreview.Colorize DIB, RotateH40(2.4 * sbParam(0)) / 40, sbParam(1) / 100, sbParam(2) / 100
            
            End Select
    'End If
End Sub
'
'Private Sub cmdPreview_Click()
'
'    If (fltPreviewed = 0) Then
'        fltPreview = -1
'        ApplyFilter fltID
'        fltPreview = 0
'        fltPreviewed = -1
'    End If
'    cmdOK.SetFocus
'
'End Sub

Private Sub cmdOK_Click()

    fltOK = -1
    pvApplyFilter fltID
    fltOK = 0

End Sub


Private Sub SetupFilter(ByVal fltID As Long)

    sbInitialized = 0

    Select Case fltID
      
      Case [fltColorize]
        Caption = "Color [Colorize]"
        SetRange 0, "Hue", 0, 100, 1, 0
        SetRange 1, "Saturation", 0, 100, 1, 50
        SetRange 2, "Luminosity", 0, 100, 1, 100
    
    End Select

    sbInitialized = -1

End Sub

Private Sub SetRange(ByVal Index As Integer, ByVal Info As String, ByVal rMin As Long, ByVal rMax As Long, ByVal rStep As Single, ByVal Start As Long, Optional ByVal Format As String = "0", Optional ByVal ApplySimetry As Boolean = 0)

    lblParamName(Index) = Info
    sbParam(Index).Tag = rStep
    sbParam(Index).Min = rMin / rStep
    sbParam(Index).Max = rMax / rStep
    sbParam(Index) = Start / rStep

    Select Case Index
      Case 0
        fraParam(0).Top = 720
      Case 1
        fraParam(0).Top = 495
        fraParam(1).Top = 1020
      Case 2
        fraParam(0).Top = 270
        fraParam(1).Top = 750
        fraParam(2).Top = 1230
    End Select
    fraParam(Index).Visible = -1
End Sub

'Private Sub EnableColors(Optional ByVal DisableColor As Integer = -1)
'
'    lblCol(0).BackColor = iSrc.Point(xRel * bfW, yRel * bfH)
'    If (DisableColor > -1) Then
'        lblCol(DisableColor).Visible = 0
'    End If
'    frColor.Visible = -1
'
'End Sub

Private Sub lblCol_Click(Index As Integer)

    With Dlg
        .Flags = &H1
        .Color = lblCol(Index).BackColor
        .ShowColor
        lblCol(Index).BackColor = .Color
    End With
    sbParam_Change 0

End Sub

Private Sub iSrc_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    iSrc_MouseMove 1, Shift, x, y

End Sub

'Private Sub iSrc_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'
'    If (Button = 1 And fltSelectPos) Then
'        '// Get relative pos.
'        If (x < 0) Then x = 0
'        If (x > bfW - 1) Then x = bfW - 1
'        If (y < 0) Then y = 0
'        If (y > bfH - 1) Then y = bfH - 1
'        xRel = x / bfW
'        yRel = y / bfH
'        '// Get color
''        lblCol(0).BackColor = RGB(pBM.Bits(2, x, y), _
''                                  pBM.Bits(1, x, y), _
''                                  pBM.Bits(0, x, y))
''        lblCol(0).Refresh
'        '// Apply filter
'        sbParam_Change 0
'    End If
'
'End Sub



':) Ulli's VB Code Formatter V2.13.2 (16/07/02 10:53:37) 73 + 607 = 680 Lines


