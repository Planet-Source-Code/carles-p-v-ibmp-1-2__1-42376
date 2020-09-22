Attribute VB_Name = "mSettings"
'================================================
' Module:        mSettings.bas
' Last revision: 2003.11.02
'================================================

Option Explicit

Public Sub LoadMainSettings()
    With fMain
        .Width = GetINI("iBMP.ini", "Forms", "MainWidth", .Width)
        .Height = GetINI("iBMP.ini", "Forms", "MainHeight", .Height)
        .Top = GetINI("iBMP.ini", "Forms", "MainTop", (Screen.Height - .Height) \ 2)
        .Left = GetINI("iBMP.ini", "Forms", "MainLeft", (Screen.Width - .Width) \ 2)
        .WindowState = GetINI("iBMP.ini", "Forms", "MainWindowState", .WindowState)
        .LastPath = GetINI("iBMP.ini", "Folders", "MainLastPath", vbNullString)
        .DialogPreview = GetINI("iBMP.ini", "Options", "MainDialogPreview", -1)
        .DialogFitMode = GetINI("iBMP.ini", "Options", "MainDialogFitMode", -1)
        .DialogJPEGquality = GetINI("iBMP.ini", "Options", "MainDialogJPEGquality", 90)
    End With
End Sub
    
Public Sub LoadPanViewSettings()
    With fPanView
        .Width = GetINI("iBMP.ini", "Forms", "PanViewWidth", .Width)
        .Height = GetINI("iBMP.ini", "Forms", "PanViewHeight", .Height)
        .Top = GetINI("iBMP.ini", "Forms", "PanViewTop", .Top)
        .Left = GetINI("iBMP.ini", "Forms", "PanViewLeft", .Left)
    End With
End Sub
    
Public Sub LoadFilterSettings()
    With fFilter
        .Top = GetINI("iBMP.ini", "Forms", "FilterTop", .Top)
        .Left = GetINI("iBMP.ini", "Forms", "FilterLeft", .Left)
        .chkFit = GetINI("iBMP.ini", "Options", "FilterBeforeFit", 1)
        .chkPickColor = GetINI("iBMP.ini", "Options", "FilterBeforePickColor", 1)
        .chkResetValues = GetINI("iBMP.ini", "Options", "FilterResetValues", 0)
        .chkNoClose = GetINI("iBMP.ini", "Options", "FilterNoClose", 0)
    End With
End Sub

Public Sub LoadTexturizeSettings()
    With fTexturize
        .Top = GetINI("iBMP.ini", "Forms", "TexturizeTop", .Top)
        .Left = GetINI("iBMP.ini", "Forms", "TexturizeLeft", .Left)
        .sbWeight = GetINI("iBMP.ini", "Options", "TexturizeWeight", 25)
        .chkInvertTexture = GetINI("iBMP.ini", "Options", "TexturizeInvert", 0)
        .chkFitMode = GetINI("iBMP.ini", "Options", "TexturizeFitMode", 0)
        .chkNoClose = GetINI("iBMP.ini", "Options", "TexturizeNoClose", 0)
        On Error GoTo ErrPath
        .flTextures.Path = GetINI("iBMP.ini", "Folders", "TexturizeFolder", AppPath)
        .flTextures.ListIndex = GetINI("iBMP.ini", "Folders", "TexturizeFile", "<None>")
        On Error GoTo 0
    End With
    Exit Sub
ErrPath:
    fTexturize.flTextures.Path = AppPath
End Sub

Public Sub LoadResizeSettings()
    With fResize
        .chkAspectRatio = GetINI("iBMP.ini", "Options", "ResizeAspectRatio", 1)
        .chkResample = GetINI("iBMP.ini", "Options", "ResizeResample", 1)
    End With
End Sub

'========================================================================================

Public Sub SaveMainSettings()
    With fMain
        If (.WindowState = vbNormal) Then
            Call PutINI("iBMP.ini", "Forms", "MainWidth", .Width)
            Call PutINI("iBMP.ini", "Forms", "MainHeight", .Height)
            Call PutINI("iBMP.ini", "Forms", "MainTop", .Top)
            Call PutINI("iBMP.ini", "Forms", "MainLeft", .Left)
        End If
        Call PutINI("iBMP.ini", "Forms", "MainWindowState", .WindowState)
        Call PutINI("iBMP.ini", "Folders", "MainLastPath", .LastPath)
        Call PutINI("iBMP.ini", "Options", "MainDialogPreview", .DialogPreview)
        Call PutINI("iBMP.ini", "Options", "MainDialogFitMode", .DialogFitMode)
        Call PutINI("iBMP.ini", "Options", "MainDialogJPEGquality", .DialogJPEGquality)
    End With
End Sub
    
Public Sub SavePanViewSettings()
    With fPanView
        Call PutINI("iBMP.ini", "Forms", "PanViewWidth", .Width)
        Call PutINI("iBMP.ini", "Forms", "PanViewHeight", .Height)
        Call PutINI("iBMP.ini", "Forms", "PanViewTop", .Top)
        Call PutINI("iBMP.ini", "Forms", "PanViewLeft", .Left)
    End With
End Sub
    
Public Sub SaveFilterSettings()
    With fFilter
        Call PutINI("iBMP.ini", "Forms", "FilterTop", .Top)
        Call PutINI("iBMP.ini", "Forms", "FilterLeft", .Left)
        Call PutINI("iBMP.ini", "Options", "FilterBeforeFit", .chkFit)
        Call PutINI("iBMP.ini", "Options", "FilterBeforePickColor", .chkPickColor)
        Call PutINI("iBMP.ini", "Options", "FilterResetValues", .chkResetValues)
        Call PutINI("iBMP.ini", "Options", "FilterNoClose", .chkNoClose)
    End With
End Sub
    
Public Sub SaveTexturizeSettings()
    With fTexturize
        Call PutINI("iBMP.ini", "Forms", "TexturizeTop", .Top)
        Call PutINI("iBMP.ini", "Forms", "TexturizeLeft", .Left)
        Call PutINI("iBMP.ini", "Folders", "TexturizeFolder", .flTextures.Path)
        Call PutINI("iBMP.ini", "Folders", "TexturizeFile", .flTextures.ListIndex)
        Call PutINI("iBMP.ini", "Options", "TexturizeWeight", .sbWeight)
        Call PutINI("iBMP.ini", "Options", "TexturizeInvert", .chkInvertTexture)
        Call PutINI("iBMP.ini", "Options", "TexturizeFitMode", .chkFitMode)
        Call PutINI("iBMP.ini", "Options", "TexturizeNoClose", .chkNoClose)
    End With
End Sub

Public Sub SaveResizeSettings()
    With fResize
        Call PutINI("iBMP.ini", "Options", "ResizeAspectRatio", .chkAspectRatio)
        Call PutINI("iBMP.ini", "Options", "ResizeResample", .chkResample)
    End With
End Sub
