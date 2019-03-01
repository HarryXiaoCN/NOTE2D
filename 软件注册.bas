Attribute VB_Name = "软件注册"
Private Const REG_SZ = 1                         ' Unicode nul terminated string
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Function 注册表注册(ByVal 注册表目录名 As String, ByVal 扩展名 As String) As Long
On Error GoTo ErrHandler
If RegOpenKey(HKEY_CLASSES_ROOT, 注册表目录名, hKey) <> 0 Then
    Shell "cmd /c md " & InstallPath, vbHide
    Do While Dir(InstallPath, vbDirectory) = ""
        DoEvents
    Loop
    'FileCopy App.Path & "\" & App.EXEName & ".exe", InstallPath & App.EXEName & ".exe"
    Shell "cmd /c copy /Y " & App.Path & "\" & App.EXEName & ".exe " & InstallPath & App.EXEName & ".exe", vbHide
    FileCopy App.Path & "\Node Notes File Icon.ico ", InstallPath & "Node Notes File Icon.ico"
'    Shell "cmd /c copy /Y " & App.Path & "\Note.ico" & InstallPath & "Note.ico", vbHide

'    环境文件拷贝
    RegSetValue HKEY_CLASSES_ROOT, 扩展名, REG_SZ, 注册表目录名, 7
    RegSetValue HKEY_CLASSES_ROOT, 扩展名 & "ShellNew", REG_SZ, "", 0
    NoteFileWrite_201 InstallPath & "新建节点文件.ntx"
    Shell "cmd /c reg add HKEY_CLASSES_ROOT\.ntx\ShellNew /v FileName /t REG_SZ /d " & InstallPath & "新建节点文件.ntx /f", vbHide
    RegSetValue HKEY_CLASSES_ROOT, 注册表目录名, REG_SZ, "Node Notes File", 15
    
    RegSetValue HKEY_CLASSES_ROOT, 注册表目录名 & "\DefaultIcon", REG_SZ, InstallPath & "Node Notes File Icon.ico", 24

    RegSetValue HKEY_CLASSES_ROOT, 注册表目录名 & "\Shell", REG_SZ, "open", 4

    RegSetValue HKEY_CLASSES_ROOT, 注册表目录名 & "\Shell\open\Command", REG_SZ, InstallPath & App.EXEName & ".exe ""%1""", 22
    
    注册表注册 = 1
Else
    注册表注册 = 2
End If
Exit Function
ErrHandler:
注册表注册 = 0
End Function
Public Function 环境文件拷贝()
On Error GoTo ErrHandler
'    FileCopy App.Path & "\comdlg32.ocx", "C:\ProgramData\Note\comdlg32.ocx"
'    FileCopy App.Path & "\RICHTX32.OCX", "C:\ProgramData\Note\RICHTX32.OCX"
    Shell "cmd /c copy /Y " & App.Path & "\comdlg32.ocx " & Environ("SystemRoot") & "\System32\comdlg32.ocx", vbHide
    Shell "cmd /c copy /Y " & App.Path & "\RICHTX32.OCX " & Environ("SystemRoot") & "\System32\RICHTX32.ocx", vbHide
    Shell "regsvr32 /s c:\Windows\System32\comdlg32.ocx", vbHide
    Shell "regsvr32 /s c:\Windows\System32\RICHTX32.ocx", vbHide
    Exit Function
ErrHandler:
MsgBox "环境文件拷贝失败！如果软件无法正常运行请将安装包内的comdlg32.ocx与RICHTX32.OCX文件复制到""C:\Windows\System32""目录下！", , "警告！"
End Function
