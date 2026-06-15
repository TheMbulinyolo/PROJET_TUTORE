Attribute VB_Name = "Macros_Isostatique"
Option Explicit

' =============================================================================
' Interface Excel <-> moteur Python pour le calcul des poutres isostatiques.
'
' Ce module est autonome : il utilise uniquement des objets Windows disponibles
' par late binding (WScript.Shell, Scripting.FileSystemObject et ADODB.Stream).
' Aucune reference VBA ne doit etre activee manuellement.
' =============================================================================

Private Const FEUILLE_ACCUEIL As String = "ACCUEIL"
Private Const FEUILLE_STRUCTURE As String = "STRUCTURE"
Private Const FEUILLE_CHARGES As String = "CHARGES"
Private Const FEUILLE_RESULTATS As String = "RESULTATS"
Private Const FEUILLE_DIAGRAMMES As String = "DIAGRAMMES"

Private Const FICHIER_ENTREE As String = "DATA\input_structure.json"
Private Const FICHIER_SORTIE As String = "DATA\resultats.json"
Private Const SCRIPT_PYTHON As String = "PYTHON\main_excel.py"

Private Const PREMIERE_LIGNE_APPUI As Long = 20
Private Const DERNIERE_LIGNE_APPUI As Long = 29
Private Const PREMIERE_LIGNE_FORCE As Long = 21
Private Const DERNIERE_LIGNE_FORCE As Long = 28
Private Const PREMIERE_LIGNE_REPARTIE As Long = 34
Private Const DERNIERE_LIGNE_REPARTIE As Long = 41
Private Const PREMIERE_LIGNE_DIAGRAMME As Long = 24


' =============================================================================
' Navigation
' =============================================================================

Public Sub AllerAccueil()
    ActiverFeuille FEUILLE_ACCUEIL
End Sub

Public Sub AllerStructure()
    ActiverFeuille FEUILLE_STRUCTURE
End Sub

Public Sub AllerCharges()
    ActiverFeuille FEUILLE_CHARGES
End Sub

Public Sub AllerResultats()
    ActiverFeuille FEUILLE_RESULTATS
End Sub

Public Sub AllerDiagrammes()
    ActiverFeuille FEUILLE_DIAGRAMMES
End Sub

Private Sub ActiverFeuille(ByVal nomFeuille As String)
    On Error GoTo GestionErreur
    ThisWorkbook.Worksheets(nomFeuille).Activate
    Exit Sub

GestionErreur:
    AfficherErreur "Navigation", _
        "La feuille '" & nomFeuille & "' est introuvable.", Err.Number
End Sub


' =============================================================================
' Reinitialisation des zones modifiables
' =============================================================================

Public Sub ReinitialiserStructure()
    On Error GoTo GestionErreur

    With ThisWorkbook.Worksheets(FEUILLE_STRUCTURE)
        .Range("C13:C14").ClearContents
        .Range("B20:D29").ClearContents
    End With
    Exit Sub

GestionErreur:
    AfficherErreur "Reinitialisation de la structure", Err.Description, Err.Number
End Sub

Public Sub ReinitialiserCharges()
    On Error GoTo GestionErreur

    With ThisWorkbook.Worksheets(FEUILLE_CHARGES)
        ' Les colonnes de formules et les unites preformatees sont conservees.
        .Range("B21:E28").ClearContents
        .Range("I21:I28").ClearContents
        .Range("B34:E41").ClearContents
        .Range("J34:J41").ClearContents
    End With
    Exit Sub

GestionErreur:
    AfficherErreur "Reinitialisation des charges", Err.Description, Err.Number
End Sub

Public Sub ReinitialiserResultats()
    On Error GoTo GestionErreur

    With ThisWorkbook.Worksheets(FEUILLE_RESULTATS)
        .Range("D21:D27").ClearContents
        .Range("H21:I26").ClearContents
        .Range("C33:C36").ClearContents
        .Range("C43:C50").ClearContents
        .Range("H43:H50").ClearContents
    End With
    Exit Sub

GestionErreur:
    AfficherErreur "Reinitialisation des resultats", Err.Description, Err.Number
End Sub

Public Sub ReinitialiserDiagrammes()
    On Error GoTo GestionErreur

    ViderDonneesDiagrammes
    Exit Sub

GestionErreur:
    AfficherErreur "Reinitialisation des diagrammes", Err.Description, Err.Number
End Sub

Public Sub ReinitialiserTout()
    On Error GoTo GestionErreur

    If MsgBox( _
        "Effacer toutes les saisies et tous les resultats ?", _
        vbQuestion + vbYesNo, _
        "Reinitialiser le projet" _
    ) <> vbYes Then Exit Sub

    Application.ScreenUpdating = False
    ReinitialiserStructure
    ReinitialiserCharges
    ReinitialiserResultats
    ReinitialiserDiagrammes
    Application.ScreenUpdating = True
    AllerAccueil
    Exit Sub

GestionErreur:
    Application.ScreenUpdating = True
    AfficherErreur "Reinitialisation complete", Err.Description, Err.Number
End Sub


' =============================================================================
' Validation des donnees
' =============================================================================

Public Sub ValiderDonneesStructure()
    Dim messageErreur As String

    If DonneesStructureValides(messageErreur) Then
        MsgBox "Les donnees de structure et de charges sont valides.", _
            vbInformation, "Validation terminee"
    Else
        MsgBox messageErreur, vbExclamation, "Donnees invalides"
    End If
End Sub

Private Function DonneesStructureValides(ByRef messageErreur As String) As Boolean
    Dim wsStructure As Worksheet
    Dim wsCharges As Worksheet
    Dim longueur As Double
    Dim nombreAppuis As Long
    Dim nombreCharges As Long
    Dim ligne As Long
    Dim identifiants As Object
    Dim identifiant As String
    Dim typeElement As String

    Set wsStructure = ThisWorkbook.Worksheets(FEUILLE_STRUCTURE)
    Set wsCharges = ThisWorkbook.Worksheets(FEUILLE_CHARGES)
    Set identifiants = CreateObject("Scripting.Dictionary")
    identifiants.CompareMode = vbTextCompare

    If EstVide(wsStructure.Range("C13").Value) Then
        messageErreur = "La longueur de la poutre est obligatoire (STRUCTURE!C13)."
        Exit Function
    End If
    If Not EstNombreValide(wsStructure.Range("C13").Value) Then
        messageErreur = "La longueur de la poutre doit etre numerique."
        Exit Function
    End If

    longueur = CDbl(wsStructure.Range("C13").Value)
    If longueur <= 0 Then
        messageErreur = "La longueur de la poutre doit etre strictement positive."
        Exit Function
    End If

    If EstVide(wsStructure.Range("C14").Value) _
       Or Not EstNombreValide(wsStructure.Range("C14").Value) Then
        messageErreur = "Le nombre de points des diagrammes doit etre numerique."
        Exit Function
    End If
    If CDbl(wsStructure.Range("C14").Value) < 2 _
       Or CDbl(wsStructure.Range("C14").Value) > 100 _
       Or CDbl(wsStructure.Range("C14").Value) <> Fix(CDbl(wsStructure.Range("C14").Value)) Then
        messageErreur = "Le nombre de points des diagrammes doit etre un entier entre 2 et 100."
        Exit Function
    End If

    For ligne = PREMIERE_LIGNE_APPUI To DERNIERE_LIGNE_APPUI
        If LigneContientDonnees(wsStructure, ligne, "B", "D") Then
            identifiant = Trim$(CStr(wsStructure.Cells(ligne, "B").Value))
            typeElement = NormaliserType(CStr(wsStructure.Cells(ligne, "C").Value))

            If identifiant = "" Or typeElement = "" _
               Or EstVide(wsStructure.Cells(ligne, "D").Value) Then
                messageErreur = "Appui incomplet a la ligne " & ligne & _
                    " : nom, type et position sont obligatoires."
                Exit Function
            End If
            If identifiants.Exists(identifiant) Then
                messageErreur = "Identifiant d'appui duplique : " & identifiant & "."
                Exit Function
            End If
            identifiants.Add identifiant, True

            If typeElement <> "pivot" And typeElement <> "rotule" _
               And typeElement <> "rouleau" And typeElement <> "encastrement" Then
                messageErreur = "Type d'appui inconnu a la ligne " & ligne & _
                    ". Utilisez Rotule, Pivot, Rouleau ou Encastrement."
                Exit Function
            End If
            If Not EstNombreValide(wsStructure.Cells(ligne, "D").Value) Then
                messageErreur = "La position de l'appui " & identifiant & " doit etre numerique."
                Exit Function
            End If
            If CDbl(wsStructure.Cells(ligne, "D").Value) < 0 _
               Or CDbl(wsStructure.Cells(ligne, "D").Value) > longueur Then
                messageErreur = "L'appui " & identifiant & _
                    " doit etre place entre 0 et " & CStr(longueur) & " m."
                Exit Function
            End If
            nombreAppuis = nombreAppuis + 1
        End If
    Next ligne

    If nombreAppuis < 2 Then
        messageErreur = "La structure doit contenir au moins deux appuis."
        Exit Function
    End If

    Set identifiants = CreateObject("Scripting.Dictionary")
    identifiants.CompareMode = vbTextCompare

    For ligne = PREMIERE_LIGNE_FORCE To DERNIERE_LIGNE_FORCE
        If LigneContientDonnees(wsCharges, ligne, "B", "E") Then
            identifiant = Trim$(CStr(wsCharges.Cells(ligne, "B").Value))
            If identifiant = "" Or EstVide(wsCharges.Cells(ligne, "C").Value) _
               Or EstVide(wsCharges.Cells(ligne, "D").Value) _
               Or EstVide(wsCharges.Cells(ligne, "E").Value) Then
                messageErreur = "Force ponctuelle incomplete a la ligne " & ligne & _
                    " : nom, fx, fy et position sont obligatoires."
                Exit Function
            End If
            If identifiants.Exists(identifiant) Then
                messageErreur = "Identifiant de charge duplique : " & identifiant & "."
                Exit Function
            End If
            identifiants.Add identifiant, True

            If Not EstNombreValide(wsCharges.Cells(ligne, "C").Value) _
               Or Not EstNombreValide(wsCharges.Cells(ligne, "D").Value) _
               Or Not EstNombreValide(wsCharges.Cells(ligne, "E").Value) Then
                messageErreur = "fx, fy et la position de " & identifiant & _
                    " doivent etre numeriques."
                Exit Function
            End If
            If CDbl(wsCharges.Cells(ligne, "E").Value) < 0 _
               Or CDbl(wsCharges.Cells(ligne, "E").Value) > longueur Then
                messageErreur = "La force " & identifiant & _
                    " doit etre placee entre 0 et " & CStr(longueur) & " m."
                Exit Function
            End If
            nombreCharges = nombreCharges + 1
        End If
    Next ligne

    For ligne = PREMIERE_LIGNE_REPARTIE To DERNIERE_LIGNE_REPARTIE
        If LigneContientDonnees(wsCharges, ligne, "B", "E") Then
            identifiant = Trim$(CStr(wsCharges.Cells(ligne, "B").Value))
            If identifiant = "" Or EstVide(wsCharges.Cells(ligne, "C").Value) _
               Or EstVide(wsCharges.Cells(ligne, "D").Value) _
               Or EstVide(wsCharges.Cells(ligne, "E").Value) Then
                messageErreur = "Charge repartie incomplete a la ligne " & ligne & _
                    " : nom, q, debut et fin sont obligatoires."
                Exit Function
            End If
            If identifiants.Exists(identifiant) Then
                messageErreur = "Identifiant de charge duplique : " & identifiant & "."
                Exit Function
            End If
            identifiants.Add identifiant, True

            If Not EstNombreValide(wsCharges.Cells(ligne, "C").Value) _
               Or Not EstNombreValide(wsCharges.Cells(ligne, "D").Value) _
               Or Not EstNombreValide(wsCharges.Cells(ligne, "E").Value) Then
                messageErreur = "q, debut et fin de " & identifiant & _
                    " doivent etre numeriques."
                Exit Function
            End If
            If CDbl(wsCharges.Cells(ligne, "D").Value) < 0 _
               Or CDbl(wsCharges.Cells(ligne, "E").Value) > longueur _
               Or CDbl(wsCharges.Cells(ligne, "E").Value) <= _
                  CDbl(wsCharges.Cells(ligne, "D").Value) Then
                messageErreur = "La charge " & identifiant & _
                    " doit verifier 0 <= debut < fin <= " & CStr(longueur) & " m."
                Exit Function
            End If
            nombreCharges = nombreCharges + 1
        End If
    Next ligne

    If nombreCharges = 0 Then
        messageErreur = "Au moins une force ponctuelle ou une charge repartie est requise."
        Exit Function
    End If

    DonneesStructureValides = True
End Function


' =============================================================================
' Export JSON vers Python
' =============================================================================

Public Sub ExporterDonneesVersJSON()
    Dim messageErreur As String

    On Error GoTo GestionErreur
    If Not DonneesStructureValides(messageErreur) Then
        MsgBox messageErreur, vbExclamation, "Export annule"
        Exit Sub
    End If

    ExporterDonneesSansValidation
    MsgBox "Donnees exportees vers :" & vbCrLf & CheminEntree(), _
        vbInformation, "Export JSON"
    Exit Sub

GestionErreur:
    AfficherErreur "Export JSON", Err.Description, Err.Number
End Sub

Private Sub ExporterDonneesSansValidation()
    Dim wsStructure As Worksheet
    Dim wsCharges As Worksheet
    Dim json As String
    Dim appuis As String
    Dim charges As String
    Dim ligne As Long
    Dim unite As String
    Dim nombrePoints As Long
    Dim pas As Double

    Set wsStructure = ThisWorkbook.Worksheets(FEUILLE_STRUCTURE)
    Set wsCharges = ThisWorkbook.Worksheets(FEUILLE_CHARGES)

    For ligne = PREMIERE_LIGNE_APPUI To DERNIERE_LIGNE_APPUI
        If Trim$(CStr(wsStructure.Cells(ligne, "B").Value)) <> "" Then
            AjouterElementJSON appuis, _
                "    {" & _
                """id"": " & JSONTexte(wsStructure.Cells(ligne, "B").Value) & ", " & _
                """type"": " & JSONTexte(NormaliserType(wsStructure.Cells(ligne, "C").Value)) & ", " & _
                """position"": " & JSONNombre(wsStructure.Cells(ligne, "D").Value) & _
                "}"
        End If
    Next ligne

    For ligne = PREMIERE_LIGNE_FORCE To DERNIERE_LIGNE_FORCE
        If Trim$(CStr(wsCharges.Cells(ligne, "B").Value)) <> "" Then
            unite = Trim$(CStr(wsCharges.Cells(ligne, "F").Value))
            If unite = "" Then unite = "kN"

            AjouterElementJSON charges, _
                "    {" & _
                """id"": " & JSONTexte(wsCharges.Cells(ligne, "B").Value) & ", " & _
                """type"": ""ponctuelle"", " & _
                """position"": " & JSONNombre(wsCharges.Cells(ligne, "E").Value) & ", " & _
                """valeur"": " & JSONNombre(wsCharges.Cells(ligne, "D").Value) & ", " & _
                """fx"": " & JSONNombre(wsCharges.Cells(ligne, "C").Value) & ", " & _
                """fy"": " & JSONNombre(wsCharges.Cells(ligne, "D").Value) & ", " & _
                """unite"": " & JSONTexte(unite) & _
                "}"
        End If
    Next ligne

    For ligne = PREMIERE_LIGNE_REPARTIE To DERNIERE_LIGNE_REPARTIE
        If Trim$(CStr(wsCharges.Cells(ligne, "B").Value)) <> "" Then
            AjouterElementJSON charges, _
                "    {" & _
                """id"": " & JSONTexte(wsCharges.Cells(ligne, "B").Value) & ", " & _
                """type"": ""repartie_uniforme"", " & _
                """debut"": " & JSONNombre(wsCharges.Cells(ligne, "D").Value) & ", " & _
                """fin"": " & JSONNombre(wsCharges.Cells(ligne, "E").Value) & ", " & _
                """valeur"": " & JSONNombre(wsCharges.Cells(ligne, "C").Value) & ", " & _
                """unite"": ""kN/m""" & _
                "}"
        End If
    Next ligne

    nombrePoints = CLng(wsStructure.Range("C14").Value)
    pas = CDbl(wsStructure.Range("C13").Value) / (nombrePoints - 1)

    json = "{" & vbCrLf & _
        "  ""poutre"": {" & vbCrLf & _
        "    ""longueur"": " & JSONNombre(wsStructure.Range("C13").Value) & "," & vbCrLf & _
        "    ""unite"": " & JSONTexte(wsStructure.Range("D13").Value) & vbCrLf & _
        "  }," & vbCrLf & _
        "  ""appuis"": [" & vbCrLf & appuis & vbCrLf & "  ]," & vbCrLf & _
        "  ""charges"": [" & vbCrLf & charges & vbCrLf & "  ]," & vbCrLf & _
        "  ""options"": {" & vbCrLf & _
        "    ""calcul_reactions"": true," & vbCrLf & _
        "    ""calcul_efforts"": true," & vbCrLf & _
        "    ""calcul_diagrammes"": true," & vbCrLf & _
        "    ""pas_discretisation"": " & JSONNombre(pas)

    If Not EstVide(ThisWorkbook.Worksheets(FEUILLE_RESULTATS).Range("C33").Value) Then
        json = json & "," & vbCrLf & _
            "    ""position_efforts"": " & _
            JSONNombre(ThisWorkbook.Worksheets(FEUILLE_RESULTATS).Range("C33").Value)
    End If

    json = json & vbCrLf & "  }" & vbCrLf & "}" & vbCrLf

    CreerDossierSiAbsent CheminRacineProjet() & "\data"
    EcrireFichierTexte CheminEntree(), json
End Sub


' =============================================================================
' Execution du moteur Python
' =============================================================================

Public Sub LancerCalculPython()
    On Error GoTo GestionErreur

    LancerPythonSansMessage
    MsgBox "Le moteur Python a termine le calcul.", _
        vbInformation, "Calcul Python"
    Exit Sub

GestionErreur:
    AfficherErreur "Execution Python", Err.Description, Err.Number
End Sub

Private Sub LancerPythonSansMessage()
    Dim shell As Object
    Dim commande As String
    Dim codeRetour As Long
    Dim executable As String
    Dim script As String
    Dim sortie As String

    script = CheminRacineProjet() & "\" & SCRIPT_PYTHON
    sortie = CheminSortie()

    If Not FichierExiste(script) Then
        Err.Raise vbObjectError + 1001, "LancerCalculPython", _
            "Script Python introuvable : " & script
    End If
    If Not FichierExiste(CheminEntree()) Then
        Err.Raise vbObjectError + 1002, "LancerCalculPython", _
            "Fichier d'entree introuvable : " & CheminEntree()
    End If

    executable = TrouverExecutablePython()
    If FichierExiste(sortie) Then Kill sortie

    If executable = "py -3" Then
        commande = executable & " " & Q(script) & " " & _
            Q(CheminEntree()) & " " & Q(sortie)
    Else
        commande = Q(executable) & " " & Q(script) & " " & _
            Q(CheminEntree()) & " " & Q(sortie)
    End If

    Set shell = CreateObject("WScript.Shell")
    codeRetour = shell.Run(commande, 0, True)

    If Not FichierExiste(sortie) Then
        Err.Raise vbObjectError + 1003, "LancerCalculPython", _
            "Python n'a pas cree le fichier de sortie." & vbCrLf & _
            "Code retour : " & codeRetour & vbCrLf & _
            "Commande : " & commande
    End If
    If codeRetour <> 0 Then
        Err.Raise vbObjectError + 1004, "LancerCalculPython", _
            "Le moteur Python a signale une erreur (code " & codeRetour & ")." & _
            vbCrLf & JSONLireTexteCle(LireFichierTexte(sortie), "message")
    End If
End Sub


' =============================================================================
' Import des resultats JSON
' =============================================================================

Public Sub ImporterResultatsDepuisJSON()
    On Error GoTo GestionErreur

    ImporterResultatsSansMessage
    MsgBox "Les resultats Python ont ete importes dans Excel.", _
        vbInformation, "Import termine"
    Exit Sub

GestionErreur:
    AfficherErreur "Import des resultats", Err.Description, Err.Number
End Sub

Private Sub ImporterResultatsSansMessage()
    Dim json As String
    Dim statut As String
    Dim message As String
    Dim reactions As String
    Dim maxima As String
    Dim efforts As String
    Dim diagrammes As String
    Dim ws As Worksheet
    Dim paires As Object
    Dim cle As Variant

    If Not FichierExiste(CheminSortie()) Then
        Err.Raise vbObjectError + 1101, "ImporterResultatsDepuisJSON", _
            "Fichier de resultats introuvable : " & CheminSortie()
    End If

    json = LireFichierTexte(CheminSortie())
    statut = UCase$(JSONLireTexteCle(json, "statut"))
    message = JSONLireTexteCle(json, "message")
    Set ws = ThisWorkbook.Worksheets(FEUILLE_RESULTATS)

    ws.Range("C50").Value = statut
    ws.Range("H50").Value = IIf(statut = "OK", "", message)

    If statut = "ERREUR" Or statut = "ERROR" Then
        MsgBox message, vbCritical, "Erreur du moteur Python"
        Err.Raise vbObjectError + 1102, "ImporterResultatsDepuisJSON", message
    End If
    If statut <> "OK" Then
        Err.Raise vbObjectError + 1103, "ImporterResultatsDepuisJSON", _
            "Statut JSON inattendu : " & statut
    End If

    ReinitialiserResultats
    ws.Range("C50").Value = statut

    reactions = JSONLireValeurBrute(json, "reactions")
    Set paires = JSONLireObjetNombres(reactions)
    For Each cle In paires.Keys
        EcrireReaction ws, CStr(cle), CDbl(paires(cle))
    Next cle

    efforts = JSONLireValeurBrute(json, "efforts_position")
    If LCase$(Trim$(efforts)) <> "null" And efforts <> "" Then
        ws.Range("C33").Value = JSONLireNombreCle(efforts, "x")
        ws.Range("C34").Value = JSONLireNombreCle(efforts, "N")
        ws.Range("C35").Value = JSONLireNombreCle(efforts, "T")
        ws.Range("C36").Value = JSONLireNombreCle(efforts, "M")
        ws.Range("C46").Value = ws.Range("C34").Value
        ws.Range("C47").Value = ws.Range("C35").Value
        ws.Range("C48").Value = ws.Range("C36").Value
        ws.Range("C49").Value = ws.Range("C33").Value
    End If

    maxima = JSONLireValeurBrute(json, "maxima")
    If LCase$(Trim$(maxima)) <> "null" And maxima <> "" Then
        EcrireMaximum ws, 21, maxima, "Nmax", "x_Nmax"
        EcrireMaximum ws, 22, maxima, "Tmax", "x_Tmax"
        EcrireMaximum ws, 23, maxima, "Mmax", "x_Mmax"

        ws.Range("H24").Value = Abs(JSONLireNombreCle(maxima, "Nmax"))
        ws.Range("I24").Value = JSONLireNombreCle(maxima, "x_Nmax")
        ws.Range("H25").Value = Abs(JSONLireNombreCle(maxima, "Tmax"))
        ws.Range("I25").Value = JSONLireNombreCle(maxima, "x_Tmax")
        ws.Range("H26").Value = Abs(JSONLireNombreCle(maxima, "Mmax"))
        ws.Range("I26").Value = JSONLireNombreCle(maxima, "x_Mmax")

        ws.Range("H43").Value = JSONLireNombreCle(maxima, "Nmax")
        ws.Range("H44").Value = JSONLireNombreCle(maxima, "x_Nmax")
        ws.Range("H45").Value = JSONLireNombreCle(maxima, "Tmax")
        ws.Range("H46").Value = JSONLireNombreCle(maxima, "x_Tmax")
        ws.Range("H47").Value = JSONLireNombreCle(maxima, "Mmax")
        ws.Range("H48").Value = JSONLireNombreCle(maxima, "x_Mmax")
    End If

    diagrammes = JSONLireValeurBrute(json, "diagrammes")
    If LCase$(Trim$(diagrammes)) <> "null" And diagrammes <> "" Then
        ws.Range("H49").Value = JSONTableauNombres( _
            JSONLireValeurBrute(diagrammes, "x")).Count
        ActualiserDiagrammesDepuisJSON diagrammes
    End If
End Sub

Private Sub EcrireReaction(ByVal ws As Worksheet, ByVal nom As String, ByVal valeur As Double)
    Dim ligne As Long
    Dim ligneLibre As Long

    For ligne = 21 To 27
        If StrComp(Trim$(CStr(ws.Cells(ligne, "B").Value)), nom, vbTextCompare) = 0 Then
            ws.Cells(ligne, "D").Value = valeur
            EcrireReactionZoneSortie ws, nom, valeur
            Exit Sub
        End If
        If ligneLibre = 0 And Trim$(CStr(ws.Cells(ligne, "D").Value)) = "" Then
            ligneLibre = ligne
        End If
    Next ligne

    If ligneLibre = 0 Then ligneLibre = 27
    ws.Cells(ligneLibre, "B").Value = nom
    ws.Cells(ligneLibre, "D").Value = valeur
    EcrireReactionZoneSortie ws, nom, valeur
End Sub

Private Sub EcrireReactionZoneSortie(ByVal ws As Worksheet, _
                                     ByVal nom As String, _
                                     ByVal valeur As Double)
    Dim ligne As Long

    For ligne = 43 To 45
        If StrComp(Trim$(CStr(ws.Cells(ligne, "B").Value)), nom, vbTextCompare) = 0 Then
            ws.Cells(ligne, "C").Value = valeur
            Exit Sub
        End If
    Next ligne
End Sub

Private Sub EcrireMaximum(ByVal ws As Worksheet, _
                          ByVal ligne As Long, _
                          ByVal jsonMaxima As String, _
                          ByVal cleValeur As String, _
                          ByVal clePosition As String)
    ws.Cells(ligne, "H").Value = JSONLireNombreCle(jsonMaxima, cleValeur)
    ws.Cells(ligne, "I").Value = JSONLireNombreCle(jsonMaxima, clePosition)
End Sub


' =============================================================================
' Macro principale
' =============================================================================

Public Sub CalculerStructure()
    Dim messageErreur As String

    On Error GoTo GestionErreur
    Application.ScreenUpdating = False
    Application.StatusBar = "Validation des donnees..."

    If Not DonneesStructureValides(messageErreur) Then
        Application.ScreenUpdating = True
        Application.StatusBar = False
        MsgBox messageErreur, vbExclamation, "Calcul annule"
        Exit Sub
    End If

    Application.StatusBar = "Export des donnees vers JSON..."
    ExporterDonneesSansValidation

    Application.StatusBar = "Calcul en cours dans Python..."
    LancerPythonSansMessage

    Application.StatusBar = "Import des resultats..."
    ImporterResultatsSansMessage

    Application.StatusBar = False
    Application.ScreenUpdating = True
    AllerResultats
    MsgBox "Calcul termine avec succes.", vbInformation, "Structure calculee"
    Exit Sub

GestionErreur:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    AfficherErreur "Calcul de la structure", Err.Description, Err.Number
End Sub


' =============================================================================
' Diagrammes et graphiques Excel
' =============================================================================

Public Sub ActualiserDiagrammes()
    Dim json As String
    Dim diagrammes As String

    On Error GoTo GestionErreur
    If Not FichierExiste(CheminSortie()) Then
        Err.Raise vbObjectError + 1201, "ActualiserDiagrammes", _
            "Aucun resultat Python disponible. Lancez d'abord le calcul."
    End If

    json = LireFichierTexte(CheminSortie())
    diagrammes = JSONLireValeurBrute(json, "diagrammes")
    If diagrammes = "" Or LCase$(Trim$(diagrammes)) = "null" Then
        Err.Raise vbObjectError + 1202, "ActualiserDiagrammes", _
            "Le fichier de sortie ne contient pas de diagrammes."
    End If

    ActualiserDiagrammesDepuisJSON diagrammes
    AllerDiagrammes
    Exit Sub

GestionErreur:
    AfficherErreur "Actualisation des diagrammes", Err.Description, Err.Number
End Sub

Private Sub ActualiserDiagrammesDepuisJSON(ByVal jsonDiagrammes As String)
    Dim x As Collection
    Dim n As Collection
    Dim t As Collection
    Dim m As Collection
    Dim ws As Worksheet
    Dim tableau As ListObject
    Dim nombrePoints As Long
    Dim donnees() As Variant
    Dim index As Long
    Dim ligneFin As Long

    Set x = JSONTableauNombres(JSONLireValeurBrute(jsonDiagrammes, "x"))
    Set n = JSONTableauNombres(JSONLireValeurBrute(jsonDiagrammes, "N"))
    Set t = JSONTableauNombres(JSONLireValeurBrute(jsonDiagrammes, "T"))
    Set m = JSONTableauNombres(JSONLireValeurBrute(jsonDiagrammes, "M"))

    nombrePoints = x.Count
    If nombrePoints = 0 Or n.Count <> nombrePoints _
       Or t.Count <> nombrePoints Or m.Count <> nombrePoints Then
        Err.Raise vbObjectError + 1203, "ActualiserDiagrammes", _
            "Les tableaux x, N, T et M sont vides ou de tailles differentes."
    End If

    Set ws = ThisWorkbook.Worksheets(FEUILLE_DIAGRAMMES)
    ViderDonneesDiagrammes

    ReDim donnees(1 To nombrePoints, 1 To 4)
    For index = 1 To nombrePoints
        donnees(index, 1) = CDbl(x(index))
        donnees(index, 2) = CDbl(n(index))
        donnees(index, 3) = CDbl(t(index))
        donnees(index, 4) = CDbl(m(index))
    Next index

    ligneFin = PREMIERE_LIGNE_DIAGRAMME + nombrePoints - 1
    ws.Range("B" & PREMIERE_LIGNE_DIAGRAMME & ":E" & ligneFin).Value = donnees

    On Error Resume Next
    Set tableau = ws.ListObjects("TableDiagrammes")
    On Error GoTo 0
    If Not tableau Is Nothing Then
        tableau.Resize ws.Range("B23:E" & ligneFin)
    End If

    MettreAJourSeriesGraphiques ws, ligneFin
End Sub

Private Sub ViderDonneesDiagrammes()
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets(FEUILLE_DIAGRAMMES)
    ws.Range("B24:E123").ClearContents
End Sub

Private Sub MettreAJourSeriesGraphiques(ByVal ws As Worksheet, ByVal ligneFin As Long)
    Dim graphique As ChartObject
    Dim index As Long
    Dim colonnes As Variant

    colonnes = Array("C", "D", "E")
    For index = 1 To ws.ChartObjects.Count
        Set graphique = ws.ChartObjects(index)
        If graphique.Chart.SeriesCollection.Count > 0 And index <= 3 Then
            With graphique.Chart.SeriesCollection(1)
                .XValues = ws.Range("B24:B" & ligneFin)
                .Values = ws.Range(colonnes(index - 1) & "24:" & _
                    colonnes(index - 1) & ligneFin)
            End With
        End If
    Next index
End Sub


' =============================================================================
' Fichiers, dossiers et chemins Windows
' =============================================================================

Public Function FichierExiste(ByVal chemin As String) As Boolean
    FichierExiste = (Len(Dir$(chemin, vbNormal Or vbHidden Or vbSystem)) > 0)
End Function

Public Function DossierExiste(ByVal chemin As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    DossierExiste = fso.FolderExists(chemin)
End Function

Public Sub CreerDossierSiAbsent(ByVal chemin As String)
    Dim fso As Object
    Dim parent As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(chemin) Then Exit Sub

    parent = fso.GetParentFolderName(chemin)
    If parent <> "" And Not fso.FolderExists(parent) Then
        CreerDossierSiAbsent parent
    End If
    fso.CreateFolder chemin
End Sub

Public Function LireFichierTexte(ByVal chemin As String) As String
    Dim flux As Object

    If Not FichierExiste(chemin) Then
        Err.Raise vbObjectError + 1301, "LireFichierTexte", _
            "Fichier introuvable : " & chemin
    End If

    Set flux = CreateObject("ADODB.Stream")
    With flux
        .Type = 2
        .Charset = "utf-8"
        .Open
        .LoadFromFile chemin
        LireFichierTexte = .ReadText
        .Close
    End With
End Function

Public Sub EcrireFichierTexte(ByVal chemin As String, ByVal contenu As String)
    Dim flux As Object
    Dim fso As Object
    Dim dossierParent As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    dossierParent = fso.GetParentFolderName(chemin)
    If dossierParent <> "" Then CreerDossierSiAbsent dossierParent

    Set flux = CreateObject("ADODB.Stream")
    With flux
        .Type = 2
        .Charset = "utf-8"
        .Open
        .WriteText contenu
        .SaveToFile chemin, 2
        .Close
    End With
End Sub

Public Sub AfficherErreur(ByVal contexte As String, _
                          ByVal description As String, _
                          Optional ByVal numero As Long = 0)
    Dim detail As String

    detail = description
    If numero <> 0 Then detail = detail & vbCrLf & vbCrLf & _
        "Code VBA : " & CStr(numero)
    MsgBox detail, vbCritical, "Erreur - " & contexte
End Sub

Private Function CheminRacineProjet() As String
    Dim fso As Object
    Dim dossierClasseur As String
    Dim dossierParent As String

    dossierClasseur = ThisWorkbook.Path
    If dossierClasseur = "" Then
        Err.Raise vbObjectError + 1302, "CheminRacineProjet", _
            "Enregistrez le classeur avant de lancer le calcul."
    End If

    If FichierExiste(dossierClasseur & "\" & SCRIPT_PYTHON) Then
        CheminRacineProjet = dossierClasseur
        Exit Function
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    dossierParent = fso.GetParentFolderName(dossierClasseur)
    If FichierExiste(dossierParent & "\" & SCRIPT_PYTHON) Then
        CheminRacineProjet = dossierParent
        Exit Function
    End If

    Err.Raise vbObjectError + 1303, "CheminRacineProjet", _
        "Impossible de localiser la racine du projet depuis :" & vbCrLf & _
        dossierClasseur
End Function

Private Function CheminEntree() As String
    CheminEntree = CheminRacineProjet() & "\" & FICHIER_ENTREE
End Function

Private Function CheminSortie() As String
    CheminSortie = CheminRacineProjet() & "\" & FICHIER_SORTIE
End Function

Private Function TrouverExecutablePython() As String
    Dim racine As String
    Dim candidats As Variant
    Dim candidat As Variant

    racine = CheminRacineProjet()
    candidats = Array( _
        racine & "\.venv\Scripts\python.exe", _
        racine & "\venv\Scripts\python.exe", _
        racine & "\ENV\Scripts\python.exe" _
    )

    For Each candidat In candidats
        If FichierExiste(CStr(candidat)) Then
            TrouverExecutablePython = CStr(candidat)
            Exit Function
        End If
    Next candidat

    ' "py -3" est fourni par le lanceur Python officiel sous Windows.
    TrouverExecutablePython = "py -3"
End Function


' =============================================================================
' JSON : ecriture et lecture ciblee du schema du moteur
'
' Ce parseur ne cherche pas a remplacer une bibliotheque JSON generaliste.
' Il gere les objets, chaines, nombres et tableaux numeriques produits par
' CORE/interface/exporteur.py, y compris les accolades dans les chaines.
' =============================================================================

Private Function JSONTexte(ByVal valeur As Variant) As String
    Dim texte As String

    texte = CStr(valeur)
    texte = Replace(texte, "\", "\\")
    texte = Replace(texte, Chr$(34), "\" & Chr$(34))
    texte = Replace(texte, vbCrLf, "\n")
    texte = Replace(texte, vbCr, "\n")
    texte = Replace(texte, vbLf, "\n")
    texte = Replace(texte, vbTab, "\t")
    JSONTexte = """" & texte & """"
End Function

Private Function JSONNombre(ByVal valeur As Variant) As String
    Dim texte As String
    Dim separateurDecimal As String
    Dim positionExposant As Long

    If Not EstNombreValide(valeur) Then
        Err.Raise vbObjectError + 1401, "JSONNombre", _
            "Valeur numerique invalide : " & CStr(valeur)
    End If

    separateurDecimal = Application.International(xlDecimalSeparator)
    texte = Trim$(CStr(CDbl(valeur)))
    If separateurDecimal <> "." Then
        texte = Replace(texte, separateurDecimal, ".")
    End If

    positionExposant = InStr(1, texte, "E", vbTextCompare)
    If positionExposant > 0 Then
        If InStr(1, Left$(texte, positionExposant - 1), ".", vbBinaryCompare) = 0 Then
            texte = Left$(texte, positionExposant - 1) & ".0" & _
                Mid$(texte, positionExposant)
        End If
    ElseIf InStr(1, texte, ".", vbBinaryCompare) = 0 Then
        texte = texte & ".0"
    End If

    JSONNombre = texte
End Function

Private Sub AjouterElementJSON(ByRef liste As String, ByVal element As String)
    If liste <> "" Then liste = liste & "," & vbCrLf
    liste = liste & element
End Sub

Private Function JSONLireValeurBrute(ByVal json As String, ByVal cle As String) As String
    Dim positionCle As Long
    Dim positionDeuxPoints As Long
    Dim debut As Long
    Dim fin As Long
    Dim caractere As String

    positionCle = JSONTrouverCle(json, cle)
    If positionCle = 0 Then Exit Function

    positionDeuxPoints = InStr(positionCle + Len(cle) + 2, json, ":")
    If positionDeuxPoints = 0 Then Exit Function
    debut = JSONIgnorerEspaces(json, positionDeuxPoints + 1)
    If debut > Len(json) Then Exit Function

    caractere = Mid$(json, debut, 1)
    Select Case caractere
        Case "{"
            fin = JSONTrouverFermeture(json, debut, "{", "}")
        Case "["
            fin = JSONTrouverFermeture(json, debut, "[", "]")
        Case """"
            fin = JSONFinChaine(json, debut)
        Case Else
            fin = debut
            Do While fin <= Len(json)
                caractere = Mid$(json, fin, 1)
                If caractere = "," Or caractere = "}" Or caractere = "]" _
                   Or caractere = vbCr Or caractere = vbLf Then Exit Do
                fin = fin + 1
            Loop
            fin = fin - 1
    End Select

    If fin >= debut Then JSONLireValeurBrute = Trim$(Mid$(json, debut, fin - debut + 1))
End Function

Private Function JSONTrouverCle(ByVal json As String, ByVal cle As String) As Long
    Dim motif As String
    Dim position As Long
    Dim dansChaine As Boolean
    Dim echappe As Boolean
    Dim caractere As String

    motif = """" & cle & """"
    position = 1
    Do While position <= Len(json) - Len(motif) + 1
        caractere = Mid$(json, position, 1)
        If dansChaine Then
            If echappe Then
                echappe = False
            ElseIf caractere = "\" Then
                echappe = True
            ElseIf caractere = """" Then
                dansChaine = False
            End If
        ElseIf caractere = """" Then
            If Mid$(json, position, Len(motif)) = motif Then
                JSONTrouverCle = position
                Exit Function
            End If
            dansChaine = True
        End If
        position = position + 1
    Loop
End Function

Private Function JSONTrouverFermeture(ByVal json As String, _
                                      ByVal debut As Long, _
                                      ByVal ouvrant As String, _
                                      ByVal fermant As String) As Long
    Dim niveau As Long
    Dim position As Long
    Dim dansChaine As Boolean
    Dim echappe As Boolean
    Dim caractere As String

    For position = debut To Len(json)
        caractere = Mid$(json, position, 1)
        If dansChaine Then
            If echappe Then
                echappe = False
            ElseIf caractere = "\" Then
                echappe = True
            ElseIf caractere = """" Then
                dansChaine = False
            End If
        Else
            If caractere = """" Then
                dansChaine = True
            ElseIf caractere = ouvrant Then
                niveau = niveau + 1
            ElseIf caractere = fermant Then
                niveau = niveau - 1
                If niveau = 0 Then
                    JSONTrouverFermeture = position
                    Exit Function
                End If
            End If
        End If
    Next position
End Function

Private Function JSONFinChaine(ByVal json As String, ByVal debut As Long) As Long
    Dim position As Long
    Dim echappe As Boolean
    Dim caractere As String

    For position = debut + 1 To Len(json)
        caractere = Mid$(json, position, 1)
        If echappe Then
            echappe = False
        ElseIf caractere = "\" Then
            echappe = True
        ElseIf caractere = """" Then
            JSONFinChaine = position
            Exit Function
        End If
    Next position
End Function

Private Function JSONIgnorerEspaces(ByVal json As String, ByVal position As Long) As Long
    Dim caractere As String

    Do While position <= Len(json)
        caractere = Mid$(json, position, 1)
        If caractere <> " " And caractere <> vbTab _
           And caractere <> vbCr And caractere <> vbLf Then Exit Do
        position = position + 1
    Loop
    JSONIgnorerEspaces = position
End Function

Private Function JSONLireTexteCle(ByVal json As String, ByVal cle As String) As String
    Dim brut As String

    brut = JSONLireValeurBrute(json, cle)
    If Len(brut) >= 2 And Left$(brut, 1) = """" _
       And Right$(brut, 1) = """" Then
        brut = Mid$(brut, 2, Len(brut) - 2)
    End If
    JSONLireTexteCle = JSONDecoderChaine(brut)
End Function

Private Function JSONDecoderChaine(ByVal texte As String) As String
    texte = Replace(texte, "\n", vbCrLf)
    texte = Replace(texte, "\r", vbCr)
    texte = Replace(texte, "\t", vbTab)
    texte = Replace(texte, "\" & Chr$(34), Chr$(34))
    texte = Replace(texte, "\\", "\")
    JSONDecoderChaine = texte
End Function

Private Function JSONLireNombreCle(ByVal json As String, ByVal cle As String) As Double
    Dim brut As String

    brut = JSONLireValeurBrute(json, cle)
    If brut = "" Or LCase$(brut) = "null" Then
        Err.Raise vbObjectError + 1402, "JSONLireNombreCle", _
            "Cle numerique absente du JSON : " & cle
    End If
    JSONLireNombreCle = CDbl(Replace(brut, ".", _
        Application.International(xlDecimalSeparator)))
End Function

Private Function JSONTableauNombres(ByVal jsonTableau As String) As Collection
    Dim resultat As New Collection
    Dim contenu As String
    Dim elements As Variant
    Dim element As Variant
    Dim texte As String

    contenu = Trim$(jsonTableau)
    If Len(contenu) < 2 Or Left$(contenu, 1) <> "[" _
       Or Right$(contenu, 1) <> "]" Then
        Set JSONTableauNombres = resultat
        Exit Function
    End If

    contenu = Trim$(Mid$(contenu, 2, Len(contenu) - 2))
    If contenu = "" Then
        Set JSONTableauNombres = resultat
        Exit Function
    End If

    elements = Split(contenu, ",")
    For Each element In elements
        texte = Trim$(CStr(element))
        resultat.Add CDbl(Replace(texte, ".", _
            Application.International(xlDecimalSeparator)))
    Next element

    Set JSONTableauNombres = resultat
End Function

Private Function JSONLireObjetNombres(ByVal jsonObjet As String) As Object
    Dim resultat As Object
    Dim expression As Object
    Dim correspondances As Object
    Dim correspondance As Object
    Dim texteObjet As String
    Dim valeur As Double

    Set resultat = CreateObject("Scripting.Dictionary")
    resultat.CompareMode = vbTextCompare
    texteObjet = Trim$(jsonObjet)

    Set expression = CreateObject("VBScript.RegExp")
    With expression
        .Global = True
        .MultiLine = True
        .Pattern = """([^""]+)""\s*:\s*(-?[0-9]+(\.[0-9]*)?([eE][+-]?[0-9]+)?)"
    End With

    Set correspondances = expression.Execute(texteObjet)
    For Each correspondance In correspondances
        valeur = CDbl(Replace(correspondance.SubMatches(1), ".", _
            Application.International(xlDecimalSeparator)))
        resultat(correspondance.SubMatches(0)) = valeur
    Next correspondance

    Set JSONLireObjetNombres = resultat
End Function


' =============================================================================
' Utilitaires generaux
' =============================================================================

Private Function EstVide(ByVal valeur As Variant) As Boolean
    If IsError(valeur) Or IsNull(valeur) Or IsEmpty(valeur) Then
        EstVide = True
    Else
        EstVide = (Trim$(CStr(valeur)) = "")
    End If
End Function

Private Function EstNombreValide(ByVal valeur As Variant) As Boolean
    If IsError(valeur) Or IsNull(valeur) Or IsEmpty(valeur) Then Exit Function
    If VarType(valeur) = vbBoolean Then Exit Function
    If Trim$(CStr(valeur)) = "" Then Exit Function
    EstNombreValide = IsNumeric(valeur)
End Function

Private Function LigneContientDonnees(ByVal ws As Worksheet, _
                                      ByVal ligne As Long, _
                                      ByVal premiereColonne As String, _
                                      ByVal derniereColonne As String) As Boolean
    LigneContientDonnees = _
        (Application.WorksheetFunction.CountA( _
            ws.Range(premiereColonne & ligne & ":" & derniereColonne & ligne)) > 0)
End Function

Private Function NormaliserType(ByVal valeur As Variant) As String
    Dim texte As String

    texte = LCase$(Trim$(CStr(valeur)))
    texte = Replace(texte, "é", "e")
    texte = Replace(texte, "è", "e")
    texte = Replace(texte, "ê", "e")
    texte = Replace(texte, " ", "_")
    texte = Replace(texte, "-", "_")
    NormaliserType = texte
End Function

Private Function Q(ByVal texte As String) As String
    If InStr(1, texte, " ", vbBinaryCompare) > 0 _
       And Not (Left$(texte, 1) = """" And Right$(texte, 1) = """") Then
        Q = """" & texte & """"
    Else
        Q = texte
    End If
End Function
