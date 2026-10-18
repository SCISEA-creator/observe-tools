Option Explicit

Dim WshShell, FSO, chemin, commande, resultat

Set WshShell = CreateObject("WScript.Shell")
Set FSO      = CreateObject("Scripting.FileSystemObject")

chemin   = FSO.GetParentFolderName(WScript.ScriptFullName)
commande = "cmd /c cd /d """ & chemin & """ && set ""PATH=%USERPROFILE%\.local\bin;%USERPROFILE%\.cargo\bin;%LOCALAPPDATA%\Programs\uv;%PATH%"" && uv run python app.py"

' Lancer sans fenetre (0) et attendre la fin de l app (True)
resultat = WshShell.Run(commande, 0, True)

' En cas d erreur, ouvrir une fenetre visible avec les instructions
If resultat <> 0 Then
    WshShell.Run "cmd /k ""cd /d """ & chemin & """ && echo. && echo ERREUR : L application n a pas pu demarrer. && echo. && echo Solutions : && echo   1. Avez-vous execute INSTALLER.bat une premiere fois ? && echo   2. Relancez INSTALLER.bat.""", 1, False
End If