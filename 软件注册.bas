Attribute VB_Name = "���ע��"
Private Const REG_SZ = 1                         ' Unicode nul terminated string
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Function ע���ע��(ByVal ע���Ŀ¼�� As String, ByVal ��չ�� As String) As Long
On Error GoTo ErrHandler
If RegOpenKey(HKEY_CLASSES_ROOT, ע���Ŀ¼��, hKey) <> 0 Then
    Shell "cmd /c md " & InstallPath, vbHide
    Do While Dir(InstallPath, vbDirectory) = ""
        DoEvents
    Loop
    'FileCopy App.Path & "\" & App.EXEName & ".exe", InstallPath & App.EXEName & ".exe"
    Shell "cmd /c copy /Y " & App.Path & "\" & App.EXEName & ".exe " & InstallPath & App.EXEName & ".exe", vbHide
    FileCopy App.Path & "\Node Notes File Icon.ico ", InstallPath & "Node Notes File Icon.ico"
'    Shell "cmd /c copy /Y " & App.Path & "\Note.ico" & InstallPath & "Note.ico", vbHide

'    �����ļ�����
    RegSetValue HKEY_CLASSES_ROOT, ��չ��, REG_SZ, ע���Ŀ¼��, 7
    RegSetValue HKEY_CLASSES_ROOT, ��չ�� & "ShellNew", REG_SZ, "", 0
    NoteFileWrite_201 InstallPath & "�½��ڵ��ļ�.ntx"
    Shell "cmd /c reg add HKEY_CLASSES_ROOT\.ntx\ShellNew /v FileName /t REG_SZ /d " & InstallPath & "�½��ڵ��ļ�.ntx /f", vbHide
    RegSetValue HKEY_CLASSES_ROOT, ע���Ŀ¼��, REG_SZ, "Node Notes File", 15
    
    RegSetValue HKEY_CLASSES_ROOT, ע���Ŀ¼�� & "\DefaultIcon", REG_SZ, InstallPath & "Node Notes File Icon.ico", 24

    RegSetValue HKEY_CLASSES_ROOT, ע���Ŀ¼�� & "\Shell", REG_SZ, "open", 4

    RegSetValue HKEY_CLASSES_ROOT, ע���Ŀ¼�� & "\Shell\open\Command", REG_SZ, InstallPath & App.EXEName & ".exe ""%1""", 22
    
    ע���ע�� = 1
Else
    ע���ע�� = 2
End If
Exit Function
ErrHandler:
ע���ע�� = 0
End Function
Public Function �����ļ�����()
On Error GoTo ErrHandler
'    FileCopy App.Path & "\comdlg32.ocx", "C:\ProgramData\Note\comdlg32.ocx"
'    FileCopy App.Path & "\RICHTX32.OCX", "C:\ProgramData\Note\RICHTX32.OCX"
    Shell "cmd /c copy /Y " & App.Path & "\comdlg32.ocx " & Environ("SystemRoot") & "\System32\comdlg32.ocx", vbHide
    Shell "cmd /c copy /Y " & App.Path & "\RICHTX32.OCX " & Environ("SystemRoot") & "\System32\RICHTX32.ocx", vbHide
    Shell "regsvr32 /s c:\Windows\System32\comdlg32.ocx", vbHide
    Shell "regsvr32 /s c:\Windows\System32\RICHTX32.ocx", vbHide
    Exit Function
ErrHandler:
MsgBox "�����ļ�����ʧ�ܣ��������޷����������뽫��װ���ڵ�comdlg32.ocx��RICHTX32.OCX�ļ����Ƶ�""C:\Windows\System32""Ŀ¼�£�", , "���棡"
End Function
