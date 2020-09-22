VERSION 5.00
Begin VB.Form fResize 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resize"
   ClientHeight    =   2805
   ClientLeft      =   6165
   ClientTop       =   5190
   ClientWidth     =   4140
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
   Icon            =   "fResize.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   187
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   276
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdRestore 
      Caption         =   "Restore"
      Height          =   375
      Left            =   2940
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   240
      Width           =   1050
   End
   Begin VB.Frame fraSize 
      Caption         =   "New size"
      Height          =   2505
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   2655
      Begin VB.ComboBox cbQuick 
         Height          =   315
         ItemData        =   "fResize.frx":000C
         Left            =   840
         List            =   "fResize.frx":002B
         TabIndex        =   8
         Text            =   "cbQuick"
         Top             =   1140
         Width           =   990
      End
      Begin VB.TextBox txtW 
         Height          =   285
         Left            =   840
         MaxLength       =   4
         TabIndex        =   2
         Top             =   390
         Width           =   990
      End
      Begin VB.TextBox txtH 
         Height          =   285
         Left            =   840
         MaxLength       =   4
         TabIndex        =   5
         Top             =   765
         Width           =   990
      End
      Begin VB.CheckBox chkAspectRatio 
         Caption         =   "&Maintain aspect ratio"
         Height          =   255
         Left            =   255
         TabIndex        =   9
         Top             =   1710
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CheckBox chkResample 
         Caption         =   "&Resample"
         Height          =   255
         Left            =   255
         TabIndex        =   10
         Top             =   2040
         Value           =   1  'Checked
         Width           =   1725
      End
      Begin VB.Label lblQuick 
         Caption         =   "Quick"
         Height          =   225
         Left            =   255
         TabIndex        =   7
         Top             =   1185
         Width           =   570
      End
      Begin VB.Label lblWidth 
         Caption         =   "Width"
         Height          =   255
         Left            =   255
         TabIndex        =   1
         Top             =   450
         Width           =   615
      End
      Begin VB.Label lblHeight 
         Caption         =   "Height"
         Height          =   255
         Left            =   255
         TabIndex        =   4
         Top             =   810
         Width           =   615
      End
      Begin VB.Label lblUnitsH 
         Caption         =   "pixels"
         Height          =   255
         Left            =   1905
         TabIndex        =   6
         Top             =   810
         Width           =   615
      End
      Begin VB.Label lblUnitsW 
         Caption         =   "pixels"
         Height          =   255
         Left            =   1905
         TabIndex        =   3
         Top             =   450
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2940
      TabIndex        =   11
      Top             =   1830
      Width           =   1050
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2940
      TabIndex        =   12
      Top             =   2280
      Width           =   1050
   End
End
Attribute VB_Name = "fResize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'================================================
' fResize form
' Last revision: 2003.11.02
'================================================

Option Explicit

Private W As Long, chgW As Boolean
Private H As Long, chgH As Boolean

Private Const MAX_PIXELS_SIZE As Long = 4000000

Private Sub Form_Load()
    
    '-- Load settings
    Call mSettings.LoadResizeSettings
End Sub

Private Sub Form_Activate()

    '-- Get DIB dimensions
    W = fMain.Canvas.DIB.Width
    H = fMain.Canvas.DIB.Height
    txtW = W
    txtH = H
    
    '-- Default 100%
    cbQuick.ListIndex = 4
    Call txtW.SetFocus
End Sub

Private Sub cbQuick_KeyPress(KeyAscii As Integer)

  Dim k As Integer

    k = KeyAscii
    If (k < 48 Or k > 57) And (k <> 8) Then
        KeyAscii = 0
    End If
    If (Val(cbQuick.Text) > 200) Then
        KeyAscii = 0
    End If
    If (Len(cbQuick.Text) > 4) Then
        KeyAscii = 0
    End If
End Sub

Private Sub cbQuick_Change()

  Dim sF As Single
  
    If (Val(cbQuick.Text) > 200) Then
        cbQuick.Text = 200
        cbQuick.SelStart = 0
        cbQuick.SelLength = 3
    End If
    
    sF = Val(cbQuick.Text) / 100
    
    txtW = Round(W * sF)
    txtH = Round(H * sF)
End Sub

Private Sub cbQuick_Click()
    
  Dim sF As Single
    
    sF = Left$(cbQuick.List(cbQuick.ListIndex), Len(cbQuick.List(cbQuick.ListIndex)) - 1)
    sF = sF / 100
    
    txtW = Round(W * sF)
    txtH = Round(H * sF)
End Sub

Private Sub txtW_KeyPress(KeyAscii As Integer)
  
  Dim k As Integer
  
    k = KeyAscii
    If (k < 48 Or k > 57) And (k <> 8) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtW_Change()

    If (Val(txtW) = 0) Then
        txtW = "1"
        txtW.SelLength = 1
    End If
    If (chkAspectRatio) Then
        If (Not chgH) Then
            chgW = True
            txtH = CInt(txtW / W * H)
            chgW = False
        End If
    End If
End Sub

Private Sub txtW_GotFocus()
    txtW.SelStart = Len(txtW)
End Sub

Private Sub txtH_KeyPress(KeyAscii As Integer)

  Dim k As Integer
  
    k = KeyAscii
    If (k < 48 Or k > 57) And (k <> 8) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtH_Change()

    If (Val(txtH) = 0) Then
        txtH = "1"
        txtH.SelLength = 1
    End If
    If (chkAspectRatio) Then
        If (Not chgW) Then
            chgH = True
            txtW = CInt(txtH / H * W)
            chgH = False
        End If
    End If
End Sub

Private Sub txtH_GotFocus()
    txtH.SelStart = Len(txtH)
End Sub

'//

Private Sub cmdRestore_Click()
    cbQuick.ListIndex = 4 '-- 1:1
    Call cbQuick_Click
End Sub

Private Sub cmdOK_Click()

    If (txtW * txtH > MAX_PIXELS_SIZE) Then
        Call MsgBox(vbCrLf & _
            "New image size exceeds maximum allowed (4 Mpixels)." & vbCrLf & vbCrLf & _
            "Please, reduce image dimensions.", vbExclamation)
        Exit Sub
    End If

    If (txtW <> W) Or (txtH <> H) Then
        
        Call Me.Hide
        DoEvents
        
        Screen.MousePointer = vbHourglass
        
        '-- Update progress max.
        fMain.Progress.Max = txtH
        '-- Resize/Resample
        Call mGDIpEx.ScaleDIB(fMain.Canvas.DIB, txtW, txtH, CBool(chkResample))
        Call fMain.Canvas_DIBProgressEnd
        '-- Remove Crop rectangle and resize canvas
        Call fMain.Canvas.RemoveCropRectangle
        Call fMain.Canvas.Resize
        '-- Update size info
        With fMain.Canvas.DIB
            fMain.Info.TextInfo = .Width & "x" & .Height & "x" & fMain.DIBbpp & "bpp"
        End With
        
        Screen.MousePointer = vbNormal
    End If
    
    Call Unload(Me)
End Sub

Private Sub cmdCancel_Click()
    Call Unload(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)

    '-- Save settings
    Call mSettings.SaveResizeSettings
End Sub
