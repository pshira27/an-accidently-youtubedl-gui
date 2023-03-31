Set oShell = WScript.CreateObject ("WScript.Shell")

MyValue = InputBox(Message, "Enter youtube link", Default)
DlScript = ".\lib\yt-dlp.exe -o ""C:\YoutubeDL\Videos\%(title)s.%(ext)s"" """&MyValue&""""

oShell.run DlScript