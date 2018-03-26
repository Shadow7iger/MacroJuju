
Sub juju()
    'd�claratin des varaibles
    Dim fileNames As Variant
    Dim i As Integer
    Dim rep As String
    Dim nomFichier As String
    Dim feuille As String
    Dim splitA As Variant
    Dim nvNom As String
    nomFichier = ActiveWorkbook.Name
    rep = ActiveWorkbook.Path
    
    '......................................
    ChDir rep 'R�pertoire par d�faut de la fenetre de selections de fichiers (ActiveWorkbook.Path = chemin du fichier excel ouvert)
    '......................................

  
    fileNames = Application.GetOpenFilename("Excel Files,*.cmf", , , , True) 'R�cuperation des fichiers
    
    'traitement des fichiers
    If IsArray(fileNames) Then
        For i = LBound(fileNames) To UBound(fileNames) 'Boucle de traitement de tout les fichiers

            splitA = Split(fileNames(i),"\")
            splitA = Split(splitA(UBound(splitA)),".")
            feuille = splitA(0)

            Workbooks.Open Filename:=fileNames(i) 'ouverture du fichier
            
            Sheets(feuille).Copy After:=Workbooks(nomFichier).Sheets(i) 'copie de la feuille results

            'recuperation du nom de la souris:            
            splitA = Split(feuille, "-")
            splitA = Split(splitA(UBound(splitA)), ".")
            nvNom = splitA(0)
            
            Worksheets(feuille).Name = nvNom 'renomage de la feuille

            rep = Mid(fileNames(i), InStrRev(fileNames(i), "\") + 1) 'r�cuperation du nom du fchier
            Workbooks(rep).Close False 'fermeture du fichier
            
        Next i
    End If
     
End Sub
