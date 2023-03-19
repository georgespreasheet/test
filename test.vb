
'Test saisies

' Déclaration des variables
Dim sapConn As Object
Dim sapGuiAuto As Object
Dim sapApp As Object
Dim sapConnection As Object
Dim session As Object
Dim notificationNumber As String
Dim notificationText As String
Dim notificationTable As Object
Dim notificationRow As Object
Dim rowIndex As Long
Dim notificationIndex As Long
Dim excelApp As Object
Dim excelWorkbook As Object
Dim excelWorksheet As Object

' Ouverture de SAP
Set sapGuiAuto = GetObject("SAPGUI")
Set sapApp = sapGuiAuto.GetScriptingEngine
Set sapConnection = sapApp.Children(0)
Set session = sapConnection.Children(0)

' Connexion à SAP
If Not sapConn.Connection.Logon(0, True) = True Then
    MsgBox "Erreur de connexion à SAP"
    Exit Sub
End If

' Ouverture du fichier Excel
Set excelApp = CreateObject("Excel.Application")
Set excelWorkbook = excelApp.Workbooks.Open("C:\example\example.xlsx")
Set excelWorksheet = excelWorkbook.Worksheets("Sheet1")

' Boucle de saisie des notifications M2
For notificationIndex = 1 To 10 ' Saisir 10 notifications
    ' Ouverture de la transaction IQS21 pour créer une nouvelle notification
    session.StartTransaction "IQS21"
    
    ' Spécification du type de notification
    session.findById("wnd[0]/usr/ctxtIQMEL-QMART").Text = "M2"
    
    ' Enregistrement de la notification
    session.findById("wnd[0]/tbar[0]/btn[11]").press
    
    ' Récupération du numéro de la notification créée
    notificationNumber = session.findById("wnd[0]/usr/ctxtIQMEL-QMNUM").Text
    
    ' Spécification du texte de la notification
    notificationText = "Texte de la notification " & notificationIndex
    
    ' Ouverture de la transaction IQS26 pour ajouter du texte à la notification
    session.StartTransaction "IQS26"
    
    ' Spécification du numéro de notification et du texte
    session.findById("wnd[0]/usr/ctxtIQMEL-QMNUM").Text = notificationNumber
    session.findById("wnd[0]/usr/ctxtIQMEL-QMTXT").Text = notificationText
    
    ' Enregistrement du texte
    session.findById("wnd[0]/tbar[0]/btn[11]").press
    
    ' Fermeture de la session IQS26
    session.findById("wnd[0]/tbar[0]/btn[15]").press
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
    
    ' Ajout de la notification au fichier Excel
    Set notificationTable = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
    For rowIndex = 1 To notificationTable.RowCount
        Set notificationRow = notificationTable.GetRow(rowIndex)
        For columnIndex = 1 To notificationRow.Count
            excelWorksheet.Cells(rowIndex, columnIndex).Value = notificationRow(columnIndex)
        Next columnIndex
    Next rowIndex
Next notificationIndex

' Sauvegarde du fichier Excel
excelWorkbook.Save

' Fermeture de la session SAP
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press

' Fermeture du fichier Excel
excelWorkbook.Close

