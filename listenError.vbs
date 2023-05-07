Option Explicit
Dim objShell, objExec, strExtractionScript, strOutput

strExtractionScript = "script.vbs"
Set objShell = CreateObject("WScript.Shell")

' Exécutez le script d'extraction en tant que processus séparé et capturez la sortie d'erreur standard
Set objExec = objShell.Exec("cscript.exe //NoLogo " & strExtractionScript)

' Lisez la sortie d'erreur standard
strOutput = objExec.StdErr.ReadAll()

' Vérifiez s'il y a des erreurs
If Len(strOutput) > 0 Then
    ' Insérez ici le code pour gérer l'erreur (par exemple, enregistrer l'erreur dans un fichier, envoyer un e-mail, etc.)
    WScript.Echo "Erreur détectée dans le script d'extraction : " & vbCrLf & strOutput
Else
    WScript.Echo "Exécution du script d'extraction réussie, aucun code d'erreur."
End If
