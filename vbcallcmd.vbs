'Creates a silent cmd instance
' (c) J. Ditzel, May 2015

Dim objShell, strCmd

' Create a shell object
Set objShell = CreateObject( "WScript.Shell" )

'This is the process you wish to call silently
strCmd = "J:\sc\batchdoc2pdf.cmd"

'Runs strCmd silently
objShell.run strCmd, 0, True

'Cleans up
Set objShell = Nothing
