<div align="center">

## change status message in yahoo messenger


</div>

### Description

it changes the status message in yahoo messenger. i know there are many such already present on psc, but i found them difficult to interpret (personal opinion no offence) and hence coded something myself which easiest for me and less code. any comments/suggestions/complaints are welcome!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[nagesh borate](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/nagesh-borate.md)
**Level**          |Beginner
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/nagesh-borate-change-status-message-in-yahoo-messenger__1-64529/archive/master.zip)





### Source Code

```
<pre>
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Private Const WM_COMMAND = &H111
Private Sub Form_Load()
On Error Resume Next
Set ws = CreateObject("wscript.shell")
cyid = ws.RegRead("HKEY_CURRENT_USER\Software\yahoo\pager\Yahoo! User ID")
nysm = InputBox("new yahoo status message?")
If nysm = "" Then
MsgBox "error!"
End
End If
ws.RegWrite "HKEY_CURRENT_USER\Software\yahoo\pager\profiles\" & cyid & "\custom msgs\1", nysm, "REG_SZ"
ws.RegDelete "HKEY_CURRENT_USER\Software\yahoo\pager\profiles\" & cyid & "\custom msgs\1_bin"
'if u want to show busy icon
'ws.RegWrite "HKEY_CURRENT_USER\Software\yahoo\pager\profiles\" & cyid & "\custom msgs\1_dnd", 1, "REG_DWORD"
' if u dont want then
ws.RegWrite "HKEY_CURRENT_USER\Software\yahoo\pager\profiles\" & cyid & "\custom msgs\1_dnd", 0, "REG_DWORD"
yhwnd = FindWindow("YahooBuddyMain", vbNullString)
If yhwnd = 0 Then
End
Else
SendMessageLong yhwnd, WM_COMMAND, 388, 1&
ydhwnd = FindWindow("#32770", vbNullString)
If ydhwnd <> 0 Then
SendKeys ("{enter}")
End If
End If
End
End Sub
</pre>
```

