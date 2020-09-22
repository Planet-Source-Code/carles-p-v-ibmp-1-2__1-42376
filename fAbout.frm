VERSION 5.00
Begin VB.Form fAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4440
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
   ForeColor       =   &H00C0C0C0&
   Icon            =   "fAbout.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   288
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   296
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1695
      TabIndex        =   0
      Top             =   3645
      Width           =   1050
   End
   Begin VB.Label lblCredits 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "lblCredits"
      ForeColor       =   &H80000015&
      Height          =   2220
      Left            =   345
      TabIndex        =   2
      Top             =   1125
      Width           =   3750
   End
   Begin VB.Label lblApp 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "lblApp"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1095
      TabIndex        =   1
      Top             =   300
      Width           =   2250
   End
End
Attribute VB_Name = "fAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    lblApp = "iBMP v" & App.Major & "." & App.Minor & vbCrLf & vbCrLf & "Carles P.V. - 2003"
    lblCredits = "Thanks to:" & vbCrLf & "Allapi Network" & vbCrLf & "VB Accelerator" & vbCrLf & "VB Frood" & vbCrLf & vbCrLf & "Special thanks to:" & vbCrLf & "Ark" & vbCrLf & "Avery" & vbCrLf & "Manuel Augusto Santos" & vbCrLf & "Robert Rayment" & vbCrLf & "Vlad Vissoultchev"
End Sub

Private Sub cmdOK_Click()
    Call Unload(Me)
End Sub
