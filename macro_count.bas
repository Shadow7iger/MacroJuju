
Sub juju()
    'd�claratin des varaibles
    Dim fileNames As Variant
    Dim i As Integer
    Dim rep As String
    Dim nomFichier As String
    Dim currSheet As String
    Dim nomSouris As String
    nomFichier = ActiveWorkbook.Name
    ChDir ActiveWorkbook.Path 'R�pertoire par d�faut de la fenetre de selections de fichiers (ActiveWorkbook.Path = chemin du fichier excel ouvert)
  
    fileNames = Application.GetOpenFilename("Excel Files,*.cmf", , , , True) 'R�cuperation des fichiers
    
    'traitement des fichiers
    If IsArray(fileNames) Then
        For i = LBound(fileNames) To UBound(fileNames) 'Boucle de traitement de tout les fichiers

            currSheet = recupNomSheet(fileNames(i)) 'recuperation nom feuille
            nomSouris = recupNomSouris(currSheet, " ") 'recuperation du nom de la souris  

            Workbooks.Open Filename:=fileNames(i) 'ouverture du fichier            
            Sheets(currSheet).Copy After:=Workbooks(nomFichier).Sheets(i) 'copie de la feuille results
            Worksheets(currSheet).Name = nomSouris 'renomage de la feuille

            fermerFichier(fileNames(i))
        Next i
    End If
     
End Sub

Function recupNomSheet(nomWorkbook)
    Dim splitA As Variant 
    splitA = Split(nomWorkbook,"\")
    splitA = Split(splitA(UBound(splitA)),".")
    recupNomSheet = splitA(0)
End Function

Function recupNomSouris(sheet,sep)
    Dim splitA As Variant
    splitA = Split(sheet, sep)
    splitA = Split(splitA(UBound(splitA)), ".")
    recupNomSouris = splitA(0)
End Function

Function fermerFichier(cheminFichier)
    Dim rep As String
    rep = Mid(cheminFichier, InStrRev(cheminFichier, "\") + 1)
    Workbooks(rep).Close False
End Function