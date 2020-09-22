Attribute VB_Name = "mHook"
Option Explicit

'-- API:

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const GWL_WNDPROC   As Long = (-4)
Private Const WM_MOUSEWHEEL As Long = &H20A
Private Const WM_CTLCOLORSCROLLBAR As Long = &H137

'//

'-- Private Variables:
Private m_OldWindowProc As Long
Private m_OldSBProc(3)  As Long

Public Sub HookWheel()
    '-- New Window proc.
    m_OldWindowProc = SetWindowLong(fMain.hWnd, GWL_WNDPROC, AddressOf pvWindowProc)
End Sub

Public Sub HookFilterSBs()
    '-- New Window procs.
    m_OldSBProc(0) = SetWindowLong(fFilter.fraParam(0).hWnd, GWL_WNDPROC, AddressOf pvSBProc0)
    m_OldSBProc(1) = SetWindowLong(fFilter.fraParam(1).hWnd, GWL_WNDPROC, AddressOf pvSBProc1)
    m_OldSBProc(2) = SetWindowLong(fFilter.fraParam(2).hWnd, GWL_WNDPROC, AddressOf pvSBProc2)
End Sub

Public Sub HookTexturizeSB()
    '-- New Window proc.
    m_OldSBProc(3) = SetWindowLong(fTexturize.fraOptions.hWnd, GWL_WNDPROC, AddressOf pvSBProc3)
End Sub

Private Function pvWindowProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    Select Case wMsg
    
        Case WM_MOUSEWHEEL
            With fMain.Canvas
                If (.DIB.hDIB) Then
                    Select Case wParam
                    Case Is > 0: Call fMain.mnuZoom_Click(0)
                    Case Else:   Call fMain.mnuZoom_Click(1)
                    End Select
                End If
            End With
    End Select
    pvWindowProc = CallWindowProc(m_OldWindowProc, hWnd, wMsg, wParam, lParam)
End Function

Private Function pvSBProc0(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    If (wMsg <> WM_CTLCOLORSCROLLBAR) Then
        If (wMsg = WM_MOUSEWHEEL) Then
            With fFilter.sbParam(0)
                Select Case wParam
                Case Is > 0: If (.Value < .Max) Then .Value = .Value + 1
                Case Else:   If (.Value > .Min) Then .Value = .Value - 1
                End Select
            End With
        End If
        pvSBProc0 = CallWindowProc(m_OldSBProc(0), hWnd, wMsg, wParam, lParam)
    End If
End Function

Private Function pvSBProc1(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    If (wMsg <> WM_CTLCOLORSCROLLBAR) Then
        If (wMsg = WM_MOUSEWHEEL) Then
            With fFilter.sbParam(1)
                Select Case wParam
                Case Is > 0: If (.Value < .Max) Then .Value = .Value + 1
                Case Else:   If (.Value > .Min) Then .Value = .Value - 1
                End Select
            End With
        End If
        pvSBProc1 = CallWindowProc(m_OldSBProc(1), hWnd, wMsg, wParam, lParam)
    End If
End Function

Private Function pvSBProc2(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    If (wMsg <> WM_CTLCOLORSCROLLBAR) Then
        If (wMsg = WM_MOUSEWHEEL) Then
            With fFilter.sbParam(2)
                Select Case wParam
                Case Is > 0: If (.Value < .Max) Then .Value = .Value + 1
                Case Else:   If (.Value > .Min) Then .Value = .Value - 1
                End Select
            End With
        End If
        pvSBProc2 = CallWindowProc(m_OldSBProc(2), hWnd, wMsg, wParam, lParam)
    End If
End Function

Private Function pvSBProc3(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    If (wMsg <> WM_CTLCOLORSCROLLBAR) Then
        pvSBProc3 = CallWindowProc(m_OldSBProc(3), hWnd, wMsg, wParam, lParam)
    End If
End Function
