<div align="center">

## GetList32 \- Get 32 Bit Listbox Text with VB3


</div>

### Description

This code will get text from listbox's on 32 bit software using VB3. I've been trying to find a way to do this for a long time now. It was not possible in the past because getting listbox text using vb3 for 32 bit programs, required User32 and Kernel32 which vb3 did not allow. So I looked and looked for a way to do it and I found it. A DLL called "Call32.dll" allowed me to use 32 bit dll's. So here it is. If you like the code, please vote for me. Also, if someone can lead me on the right path for creating a .HLP file in vb3, I will create a help file on using Call32.dll.

Jeffrey C. Tatum - http://www.oaknetwork.com/vb
 
### More Info
 
You will need the DLL "Call32.dll" which can be found anywhere on the web.

Returns 32 bit listbox text using VB3. This code was written by me, Jeffrey C. Tatum, for use with getting Screen Names from AOL 32 bit (3.0 95, 4.0, 5.0, Beta 6.0)


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jeffrey C\. Tatum](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jeffrey-c-tatum.md)
**Level**          |Advanced
**User Rating**    |4.0 (12 globes from 3 users)
**Compatibility**  |VB 3\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jeffrey-c-tatum-getlist32-get-32-bit-listbox-text-with-vb3__1-8806/archive/master.zip)

### API Declarations

```
Global Const WM_USER = &H400
Global Const LB_GETITEMDATA = (WM_USER + 26)
'32-bit Call32 format:
Declare Function Declare32& Lib "call32.dll" (ByVal func$, ByVal library$, ByVal args$)
'32-bit registry functions in Call32 format:
Declare Sub CloseHandle Lib "call32.dll" Alias "call32" (ByVal hObject As Long, ByVal id As Long)
Declare Function FindWindowEx Lib "call32.dll" Alias "call32" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As Any, ByVal lpsz2 As Any, ByVal id As Long) As Long
Declare Function GetWindowThreadProcessId Lib "call32.dll" Alias "call32" (ByVal hWnd As Long, lpdwProcessId As Long, ByVal id As Long) As Long
Declare Function OpenProcess Lib "call32.dll" Alias "call32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long, ByVal id As Long) As Long
Declare Sub ReadProcessMemory Lib "call32.dll" Alias "call32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByVal lpNumberOfBytesWritten As Long, ByVal id As Long)
Declare Sub RtlMoveMemory Lib "call32.dll" Alias "call32" (destination As Any, source As Any, ByVal Length As Long, ByVal id As Long)
```


### Source Code

```
Function AOLGetList32 (tree, Index As Integer, Buffer As String)
'Tree = The listbox
'Index = Listbox Index
'Buffer = output
'Example:
'  a = GetList32(SomeList&, 0, Buffer$)
'  MsgBox Buffer$
'Buffer is the text that was taken from the 32 bit
'listbox.
On Error Resume Next
DoEvents: idGetWindowThreadProcessId = Declare32("GetWindowThreadProcessId", "user32", "ip")
DoEvents: idOpenProcess = Declare32("OpenProcess", "kernel32", "ppi")
DoEvents: idReadProcessMemory = Declare32("ReadProcessMemory", "kernel32", "iipip")
DoEvents: idRtlMoveMemory = Declare32("RtlMoveMemory", "kernel32", "ppi")
DoEvents: idCloseHandle = Declare32("CloseHandle", "kernel32", "p")
Dim AOLProcess As Long
Dim ListItemHold As Long
Dim PerSon As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
AOLThread = GetWindowThreadProcessId(tree, AOLProcess, idGetWindowThreadProcessId)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess, idOpenProcess)
If AOLProcessThread Then
PerSon$ = String$(4, 0&)
ListItemHold = SendMessage(tree, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, PerSon$, 4, ReadBytes, idReadProcessMemory)
Call RtlMoveMemory(ListPersonHold, ByVal PerSon$, 4, idRtlMoveMemory)
ListPersonHold = ListPersonHold + 6
PerSon$ = String$(17, 0&)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, PerSon$, Len(PerSon$), ReadBytes, idReadProcessMemory)
PerSon$ = Left$(PerSon$, InStr(PerSon$, Chr(0)) - 1)
Call CloseHandle(AOLProcessThread, idCloseHandle)
End If
Buffer$ = PerSon$
End Function
```

