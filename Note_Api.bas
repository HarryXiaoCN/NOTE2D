Attribute VB_Name = "Note_Api"
Option Explicit
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Const WS_EX_LAYERED = &H80000
Public Const GWL_EXSTYLE = (-20)
Public Const LWA_ALPHA = &H2
Public Const LWA_COLORKEY = &H1
Public Const WS_EX_TRANSPARENT = &H20&
Private Const GWL_WNDPROC = -4
Private Const WM_MOUSEWHEEL = &H20A

Global lpPrevWndProcA As Long

Public bMouseFlag As Boolean '����¼������־
Private Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case uMsg
Case WM_MOUSEWHEEL '����
Dim wzDelta, wKeys As Integer
'wzDelta���ݹ��ֹ����Ŀ�������ֵС�����ʾ���������������û����򣩣�
'�������ʾ������ǰ����������ʾ������
wzDelta = HIWORD(wParam)
'wKeysָ���Ƿ���CTRL=8��SHIFT=4������(��=2����=16����=2������)���£�������
wKeys = LOWORD(wParam)
'--------------------------------------------------
If wzDelta < 0 Then '���û�����
    RollerEventHandling True
Else '����ʾ������
    RollerEventHandling False
End If
'--------------------------------------------------
Case Else
WindowProc = CallWindowProc(lpPrevWndProcA, hw, uMsg, wParam, lParam)
End Select
End Function

Private Function HIWORD(LongIn As Long) As Integer
HIWORD = (LongIn And &HFFFF0000) \ &H10000 'ȡ��32λֵ�ĸ�16λ
End Function
Private Function LOWORD(LongIn As Long) As Integer
LOWORD = LongIn And &HFFFF& 'ȡ��32λֵ�ĵ�16λ
End Function
Public Sub HookMouse(ByVal hWnd As Long)
lpPrevWndProcA = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub UnHookMouse(ByVal hWnd As Long)
SetWindowLong hWnd, GWL_WNDPROC, lpPrevWndProcA
End Sub
Public Function FormTransparent(ByRef formObj As Form, ByRef trNum As Byte)
Dim rtn As Long
rtn = GetWindowLong(formObj.hWnd, GWL_EXSTYLE)
SetWindowLong formObj.hWnd, GWL_EXSTYLE, rtn Or WS_EX_LAYERED
SetLayeredWindowAttributes formObj.hWnd, 0, trNum, LWA_ALPHA
End Function
Public Function FormStick(ByRef formObj As Form, ByRef Stick As Boolean)
    If Stick = True Then
        SetWindowPos formObj.hWnd, -1, 0, 0, 0, 0, 2 Or 1
    Else
        SetWindowPos formObj.hWnd, -2, 0, 0, 0, 0, 2 Or 1
    End If
End Function
