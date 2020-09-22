VERSION 5.00
Begin VB.Form fOrientation 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Orientation"
   ClientHeight    =   3165
   ClientLeft      =   6135
   ClientTop       =   4980
   ClientWidth     =   4200
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
   Icon            =   "fOrientation.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   211
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cbOrientation 
      Height          =   315
      ItemData        =   "fOrientation.frx":000C
      Left            =   2610
      List            =   "fOrientation.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   375
      Width           =   930
   End
   Begin VB.CheckBox chkFlip 
      Caption         =   "Flip &vertical"
      Height          =   210
      Index           =   1
      Left            =   2610
      TabIndex        =   5
      Top             =   1605
      Width           =   1425
   End
   Begin VB.CheckBox chkFlip 
      Caption         =   "Flip &horizontal"
      Height          =   210
      Index           =   0
      Left            =   2610
      TabIndex        =   4
      Top             =   1275
      Width           =   1425
   End
   Begin VB.PictureBox iPreview 
      BackColor       =   &H8000000C&
      ClipControls    =   0   'False
      Height          =   2310
      Left            =   105
      ScaleHeight     =   150
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   2310
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3045
      TabIndex        =   7
      Top             =   2685
      Width           =   1050
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1905
      TabIndex        =   6
      Top             =   2685
      Width           =   1050
   End
   Begin VB.Label lblFlip 
      Caption         =   "Flip image"
      Height          =   225
      Left            =   2625
      TabIndex        =   3
      Top             =   945
      Width           =   885
   End
   Begin VB.Label lblOrientation 
      Caption         =   "Canvas orientation"
      Height          =   255
      Left            =   2610
      TabIndex        =   1
      Top             =   90
      Width           =   1425
   End
End
Attribute VB_Name = "fOrientation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private PrvDIB As New cDIB
Private TmpDIB As New cDIB

Private bfx As Long, bfy As Long
Private bfW As Long, bfH As Long

Private Sub Form_Load()
    
    With fMain.Canvas
        '-- Get dest. best fit dim. and pos.
        Call .DIB.GetBestFitInfo(150, 150, bfx, bfy, bfW, bfH)
        '-- Clear preview
        Call iPreview.Cls
        '-- Get preview/temp DIBs
        Call PrvDIB.Create(bfW, bfH)
        Call PrvDIB.LoadDIBBlt(.DIB)
        Call TmpDIB.Create(bfW, bfH)
        Call TmpDIB.LoadDIBBlt(.DIB)
    End With
    
    cbOrientation.ListIndex = 0
End Sub

Private Sub Form_Paint()
    Line (0, 170)-(ScaleWidth, 170), vb3DShadow
    Line (0, 171)-(ScaleWidth, 171), vb3DHighlight
End Sub

Private Sub cbOrientation_Click()
    Call pvPreview
End Sub

Private Sub chkFlip_Click(Index As Integer)
    Call pvPreview
End Sub

Private Sub pvPreview()

  Dim DIBFilter As New cDIBFilter
    
    With PrvDIB
    
        '-- Get original DIB
        Call .Create(TmpDIB.Width, TmpDIB.Height)
        Call .LoadDIBBlt(TmpDIB)
        Call .GetBestFitInfo(150, 150, bfx, bfy, bfW, bfH)
        
        '-- Change orientation
        If (cbOrientation.ListIndex = 1 Or CBool(chkFlip(0)) Or CBool(chkFlip(1))) Then
            Call .Orientation((cbOrientation.ListIndex = 1), CBool(chkFlip(0)), CBool(chkFlip(1)))
            Call .GetBestFitInfo(150, 150, bfx, bfy, bfW, bfH)
        End If
    End With
    
    '-- Refresh
    Call iPreview.Cls
    Call iPreview_Paint
End Sub

Private Sub iPreview_Paint()
    Call PrvDIB.Paint(iPreview.hDC, bfx, bfy)
End Sub

'//

Private Sub cmdOK_Click()

    Call Me.Hide
    DoEvents
        
    With fMain.Canvas
        
        '-- Change orientation
        If ((cbOrientation.ListIndex = 1) Or CBool(chkFlip(0)) Or CBool(chkFlip(1))) Then
            
            Call .DIB.Orientation((cbOrientation.ListIndex = 1), CBool(chkFlip(0)), CBool(chkFlip(1)))
            
            '-- Remove Crop rectangle and resize canvas
            Call .RemoveCropRectangle
            Call .Resize
            
            '-- Update DIB info
            With .DIB
                '-- Update progress max.
                fMain.Progress.Max = .Height
                '-- Update size info
                fMain.Info.TextInfo = .Width & "x" & .Height & "x" & fMain.DIBbpp & "bpp"
            End With
        End If
    End With
    
    Call Unload(Me)
End Sub

Private Sub cmdCancel_Click()
    Call Unload(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set PrvDIB = Nothing
    Set TmpDIB = Nothing
End Sub

