Sub WriteToLog(message)
    Dim fso, logFile, logFilePath

    logFilePath = "C:\Chemin\vers\le\dossier\logfile.csv" ' Spécifiez le chemin du fichier de journalisation

    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(logFilePath) Then
        Set logFile = fso.OpenTextFile(logFilePath, 8, True) ' Ouvrir le fichier en mode Append (8)
    Else
        Set logFile = fso.CreateTextFile(logFilePath, True) ' Créer le fichier s'il n'existe pas
    End If

    logFile.WriteLine Now & "," & message ' Écrire le message avec un horodatage
    logFile.Close ' Fermer le fichier

    Set logFile = Nothing
    Set fso = Nothing
End Sub

Sub SendEmail(subject, body)
    Dim objEmail, smtpServer

    smtpServer = "smtp.example.com" ' Remplacez par l'adresse de votre serveur SMTP
    emailAddress = "your.email@example.com" ' Remplacez par votre adresse e-mail

    Set objEmail = CreateObject("CDO.Message")
    With objEmail
        .From = emailAddress
        .To = emailAddress
        .Subject = subject
        .TextBody = body
        .Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
        .Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = smtpServer
        .Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
        .Configuration.Fields.Update
        .Send
    End With
    Set objEmail = Nothing
End Sub

Sub SendWebexMessageOnError(errorMessage)
    Dim objHTTP, strURL, strJSON, objJSON

    ' API Webex URL
    strURL = "https://webexapis.com/v1/messages"

    ' Préparer le message JSON
    strJSON = "{""toPersonEmail"":""" & toPersonEmail & """, ""text"":""" & errorMessage & """}"

    ' Créer une requête HTTP pour accéder à l'API Webex
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    objHTTP.Open "POST", strURL, False
    objHTTP.setRequestHeader "Content-Type", "application/json"
    objHTTP.setRequestHeader "Authorization", "Bearer " & accessToken
    objHTTP.send strJSON

    ' Vérifier si la requête a été envoyée avec succès
    If objHTTP.Status <> 200 Then
        WScript.Echo "Erreur lors de l'envoi du message à Webex : " & objHTTP.Status & " " & objHTTP.statusText
    End If

    Set objHTTP = Nothing
End Sub

Sub CheckAndClickError()
    Dim objShell, objWMI, colItems, objItem, sapErrorTitle

    sapErrorTitle = "SAP Error" ' Remplacez "SAP Error" par le titre exact de la fenêtre d'erreur

    Set objShell = CreateObject("WScript.Shell")
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")

    Set colItems = objWMI.ExecQuery("Select * From Win32_Process WHERE Name='saplogon.exe'")

    For Each objItem In colItems
        If objShell.AppActivate(sapErrorTitle) Then
            WScript.Sleep 500
            objShell.SendKeys "{ENTER}" ' Clique sur le bouton OK de la fenêtre d'erreur
            Exit Sub
        End If
    Next

    Set colItems = Nothing
    Set objWMI = Nothing
    Set objShell = Nothing
End Sub

Sub SAPConnect()
    UserName = "NAME"
    PassWord = "PASS"
    set Wshell = CreateObject("Wscript.Shell")
    Wshell.run "saplogon.exe, //SHORTCUT=" & chr(34) "-desc" & chr(34) & "110 PRD-P01 ECC" & chr(34) & "-sid=" & chr(34) & "P01"& chr(34) & "-clt=" & chr(34) & "100" & chr(34) & "-u=" & chr(34) & UserName & chr(34) & "-l=" & chr(34) & "EN" & chr(34) & "-t=" & chr(34) & ""& chr(34)& "-cmd="& chr(34) & "" & chr(34) & "-tit="& chr(34)&"100 PRD-P01 ECC"& chr(34) & "-trc=" & chr(34)& "" & chr(34) & "-password="& chr(34) & PassWord & chr(34)
    Wscript.Sleep 3000
    set Wshell = Nothing
End Sub

sub RunSubVBS(SubPath, StrH)
    SAPConnect()
    Set oWsh = CreateObject("Shell.Application")
    oWsh.ShellEexecute SubPath
    While NbMin ()= Str2Min (StrH)    
        Wscript.Sleep 3000    
    Wend
End Sub
