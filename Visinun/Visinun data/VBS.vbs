Randomize
Set WshShell = CreateObject("WScript.Shell")

' ?? Спокійний початок
WScript.Sleep 1000
MsgBox "System loading...", 64, "System"

WScript.Sleep 1500
MsgBox "Everything is fine...", 64, "System"

WScript.Sleep 1500
MsgBox "Wait...", 48, "System"

' ?? Початок хаосу
For i = 1 To 15
    WshShell.Run "notepad"
    WScript.Sleep 200
    
    WshShell.AppActivate "Notepad"
    WScript.Sleep 100
    
    WshShell.SendKeys "it is a prank! ??"
Next

' ?? “Божевілля” (більше вікон + випадковість)
For i = 1 To 15
    WshShell.Run "notepad"
    WScript.Sleep Int((400 * Rnd) + 100)
    
    WshShell.AppActivate "Notepad"
    WScript.Sleep 100
    
    WshShell.SendKeys "IT IS A PRANK!!! ??"
Next

' ??? Псевдо-страшний момент
WScript.Sleep 1000
MsgBox "Why are you still here?", 16, "???"

WScript.Sleep 1200
MsgBox "You can't stop it...", 16, "???"

WScript.Sleep 1200
MsgBox "RUN.", 16, "???"

' ?? Відкриває щось (Kinito-ефект)
WshShell.Run "https://www.youtube.com/watch?v=dQw4w9WgXcQ"

' ?? Кінець
WScript.Sleep 2000
MsgBox "Relax, it was just a prank ??", 64, "END"