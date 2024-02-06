Set WshShell = WScript.CreateObject("WScript.Shell")

' Načte obsah .txt souboru
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set inputFile = objFSO.OpenTextFile(WScript.Arguments(0), 1)

' Čte obsah řádek po řádku
Do Until inputFile.AtEndOfStream
    line = inputFile.ReadLine
    ProcessLine line
Loop

' Zavře soubor
inputFile.Close

Sub ProcessLine(line)
    ' Rozpoznání a provedení akcí
    If InStr(line, "WAIT") > 0 Then
        ' Akce WAIT - počká 5 sekund (nebo specifikovaný čas)
        WScript.Sleep 5000  ' čas v milisekundách (zde 5 sekund)
    ElseIf InStr(line, "SEND") > 0 Then
        ' Akce SEND - simulace kláves
        keys = Mid(line, 6)  ' odstranění "SEND " ze začátku řádku
        WshShell.SendKeys keys
    ElseIf InStr(line, "SHELL") > 0 Then
        ' Akce SHELL - spustí externí příkaz
        command = Mid(line, 7)  ' odstranění "SHELL " ze začátku řádku
        WshShell.Run command
    ElseIf InStr(line, "MULT") > 0 Then
        ' Akce MULT - násobení
        factors = Split(Mid(line, 6), ",")
        result = factors(0) * factors(1)
    ElseIf InStr(line, "DIV") > 0 Then
        ' Akce DIV - dělení
        factors = Split(Mid(line, 5), ",")
        result = factors(0) / factors(1)
    ElseIf InStr(line, "ADD") > 0 Then
        ' Akce ADD - sčítání
        terms = Split(Mid(line, 5), ",")
        result = terms(0) + terms(1)
    ElseIf InStr(line, "SUB") > 0 Then
        ' Akce SUB - odečítání
        terms = Split(Mid(line, 5), ",")
        result = terms(0) - terms(1)
    End If
End Sub
