' wsshd launcher - start wsshd-start.cmd with no console window.
'
' A Scheduled Task action of type "cmd.exe /c wsshd-start.cmd" runs in an
' interactive session (the task needs the user's ConPTY), so Windows creates a
' visible console window at every logon. The task's "Hidden" flag does not
' suppress that window; only the SW_HIDE passed to CreateProcess does. wscript
' gives us that: Run(..., 0, False) launches cmd.exe hidden and returns at once,
' while the hidden cmd.exe keeps node.exe attached for the life of the session.
Option Explicit

Dim fso, sh, here, launcher
Set fso = CreateObject("Scripting.FileSystemObject")
Set sh = CreateObject("WScript.Shell")

here = fso.GetParentFolderName(WScript.ScriptFullName)
launcher = fso.BuildPath(here, "wsshd-start.cmd")

sh.CurrentDirectory = here
' Wait for the child so the task keeps reporting Running; if node ever exits
' with a non-zero code the task's "restart on failure" applies as usual.
sh.Run "cmd.exe /d /c """ & launcher & """", 0, True
