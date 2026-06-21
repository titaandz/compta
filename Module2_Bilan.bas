Attribute VB_Name = "Module2_Bilan"
Option Explicit

' ============================================================
' ONGLETS BILAN - ACTIF
' ============================================================

Sub CreerOngletActif()
    Dim ws As Worksheet
    Dim i As Long, ligne As Long

    Set ws = CreerOnglet("Actif immobilise", COULEUR_BILAN)
    Call EcrireEnTete(ws, "ACTIF IMMOBILISE", "Comptes de classe 2 - Immobilisations")

    ligne = 4

    ' --- IMMOBILISATIONS INCORPORELLES (20) ---
    ws.Cells(ligne, 2).Value = "IMMOBILISATIONS INCORPORELLES"
    ws.Cells(ligne, 2).Font.Bold = True
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 7)).Interior.Color = RGB(220, 230, 241)
    ligne = ligne + 1

    Dim compIncorp() As String
    compIncorp = Split("200,201,203,205,206,207,208", ",")
    Dim labIncorp() As String
    labIncorp = Split("Fonds commercial,Frais etablissement,Frais R&D,Concessions brevets,Droit au bail,Fonds de commerce,Autres immob. incorporelles,Autres", ",")

    Dim totalIncorpN As Double, totalIncorpN1 As Double
    Dim j As Integer
    For j = 0 To UBound(compIncorp)
        Dim sN As Double, sN1 As Double
        sN = GetSoldeCompteN(compIncorp(j))
        sN1 = GetSoldeCompteN1(compIncorp(j))
        If sN <> 0 Or sN1 <> 0 Then
            Call EcrireLigneCompte(ws, ligne, compIncorp(j), labIncorp(j), sN, sN1)
            totalIncorpN = totalIncorpN + sN
            totalIncorpN1 = totalIncorpN1 + sN1
            ligne = ligne + 1
        End If
    Next j
    Call EcrireSousTotal(ws, ligne, "Total immobilisations incorporelles", totalIncorpN, totalIncorpN1)
    ligne = ligne + 2

    ' --- IMMOBILISATIONS CORPORELLES (21) ---
    ws.Cells(ligne, 2).Value = "IMMOBILISATIONS CORPORELLES"
    ws.Cells(ligne, 2).Font.Bold = True
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 7)).Interior.Color = RGB(220, 230, 241)
    ligne = ligne + 1

    Dim compCorp() As String
    compCorp = Split("211,212,213,214,215,218", ",")
    Dim labCorp() As String
    labCorp = Split("Terrains,Agenc. aménag. terrains,Constructions,Constructions sur sol autrui,Installations tech. mat. outillage,Autres immob. corporelles", ",")

    Dim totalCorpN As Double, totalCorpN1 As Double
    For j = 0 To UBound(compCorp)
        sN = GetSoldeCompteN(compCorp(j))
        sN1 = GetSoldeCompteN1(compCorp(j))
        If sN <> 0 Or sN1 <> 0 Then
            Call EcrireLigneCompte(ws, ligne, compCorp(j), labCorp(j), sN, sN1)
            totalCorpN = totalCorpN + sN
            totalCorpN1 = totalCorpN1 + sN1
            ligne = ligne + 1
        End If
    Next j

    ' Immobilisations en cours (23)
    sN = GetSoldeCompteN("23")
    sN1 = GetSoldeCompteN1("23")
    If sN <> 0 Or sN1 <> 0 Then
        Call EcrireLigneCompte(ws, ligne, "23", "Immobilisations en cours", sN, sN1)
        totalCorpN = totalCorpN + sN
        totalCorpN1 = totalCorpN1 + sN1
        ligne = ligne + 1
    End If

    Call EcrireSousTotal(ws, ligne, "Total immobilisations corporelles (brut)", totalCorpN, totalCorpN1)
    ligne = ligne + 1

    ' Amortissements (28)
    Dim amortN As Double, amortN1 As Double
    amortN = GetSoldeCompteN("28")
    amortN1 = GetSoldeCompteN1("28")
    If amortN <> 0 Or amortN1 <> 0 Then
        Call EcrireLigneCompte(ws, ligne, "28", "Amortissements immob. corporelles", amortN, amortN1)
        ligne = ligne + 1
    End If

    ' Depreciations (29)
    Dim deprN As Double, deprN1 As Double
    deprN = GetSoldeCompteN("29")
    deprN1 = GetSoldeCompteN1("29")
    If deprN <> 0 Or deprN1 <> 0 Then
        Call EcrireLigneCompte(ws, ligne, "29", "Depreciations immobilisations", deprN, deprN1)
        ligne = ligne + 1
    End If

    Dim corpNetN As Double, corpNetN1 As Double
    corpNetN = totalCorpN + amortN + deprN  ' amortissements sont crediteurs donc negatifs
    corpNetN1 = totalCorpN1 + amortN1 + deprN1
    Call EcrireSousTotal(ws, ligne, "Total immob. corporelles (net)", corpNetN, corpNetN1)
    ligne = ligne + 2

    ' --- IMMOBILISATIONS FINANCIERES (26-27) ---
    ws.Cells(ligne, 2).Value = "IMMOBILISATIONS FINANCIERES"
    ws.Cells(ligne, 2).Font.Bold = True
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 7)).Interior.Color = RGB(220, 230, 241)
    ligne = ligne + 1

    Dim compFin() As String
    compFin = Split("261,262,266,267,271,272,274,275,276,27", ",")
    Dim labFin() As String
    labFin = Split("Titres participation,Titres participation (autres),Autres formes participation,Creances rattachees,Titres immobilises (droit propriete),Titres immobilises (droit creance),Prets,Depots et caut. verses,Autres creances immobilisees,Immob. financieres diverses", ",")

    Dim totalFinN As Double, totalFinN1 As Double
    For j = 0 To UBound(compFin)
        sN = GetSoldeCompteN(compFin(j))
        sN1 = GetSoldeCompteN1(compFin(j))
        If sN <> 0 Or sN1 <> 0 Then
            Call EcrireLigneCompte(ws, ligne, compFin(j), labFin(j), sN, sN1)
            totalFinN = totalFinN + sN
            totalFinN1 = totalFinN1 + sN1
            ligne = ligne + 1
        End If
    Next j
    Call EcrireSousTotal(ws, ligne, "Total immobilisations financieres", totalFinN, totalFinN1)
    ligne = ligne + 2

    ' TOTAL ACTIF IMMOBILISE
    Dim totalActifImmN As Double, totalActifImmN1 As Double
    totalActifImmN = totalIncorpN + corpNetN + totalFinN
    totalActifImmN1 = totalIncorpN1 + corpNetN1 + totalFinN1
    Call EcrireTotal(ws, ligne, "TOTAL ACTIF IMMOBILISE (NET)", totalActifImmN, totalActifImmN1)

    ' Figer les volets
    ws.Range("A4").Select
    ActiveWindow.FreezePanes = True

    ' Stocker dans une cellule nommee pour recuperation
    ws.Names.Add Name="TotalActifImmN", RefersTo:=totalActifImmN
    ws.Names.Add Name="TotalActifImmN1", RefersTo:=totalActifImmN1
End Sub

' ============================================================
' STOCKS ET EN-COURS (Classe 3)
' ============================================================
Sub CreerOngletStocks()
    Dim ws As Worksheet
    Dim ligne As Long

    Set ws = CreerOnglet("Stocks", COULEUR_BILAN)
    Call EcrireEnTete(ws, "STOCKS ET EN-COURS", "Comptes de classe 3")
    ligne = 4

    Dim compStocks() As String
    compStocks = Split("31,32,33,34,35,37,38,39", ",")
    Dim labStocks() As String
    labStocks = Split("Matieres premieres et approvisionnements,Autres approvisionnements,En-cours de production (biens),En-cours de production (services),Produits intermediaires et finis,Marchandises,Stocks en voie d'acheminement,Depreciations sur stocks", ",")

    Dim totalStocksN As Double, totalStocksN1 As Double
    Dim j As Integer, sN As Double, sN1 As Double

    For j = 0 To UBound(compStocks)
        sN = GetSoldeCompteN(compStocks(j))
        sN1 = GetSoldeCompteN1(compStocks(j))
        If sN <> 0 Or sN1 <> 0 Then
            Call EcrireLigneCompte(ws, ligne, compStocks(j), labStocks(j), sN, sN1)
            If compStocks(j) <> "39" Then
                totalStocksN = totalStocksN + sN
                totalStocksN1 = totalStocksN1 + sN1
            Else
                totalStocksN = totalStocksN + sN
                totalStocksN1 = totalStocksN1 + sN1
            End If
            ligne = ligne + 1
        End If
    Next j

    ligne = ligne + 1
    Call EcrireTotal(ws, ligne, "TOTAL STOCKS (NET)", totalStocksN, totalStocksN1)

    ' Informations complementaires
    ligne = ligne + 3
    ws.Cells(ligne, 1).Value = "INFORMATIONS COMPLEMENTAIRES"
    ws.Cells(ligne, 1).Font.Bold = True
    ligne = ligne + 1
    ws.Cells(ligne, 2).Value = "Variation de stocks matieres (31+32)"
    ws.Cells(ligne, 3).Value = GetSoldeCompteN("31") + GetSoldeCompteN("32") - GetSoldeCompteN1("31") - GetSoldeCompteN1("32")
    ligne = ligne + 1
    ws.Cells(ligne, 2).Value = "Variation de stocks produits finis (35)"
    ws.Cells(ligne, 3).Value = GetSoldeCompteN("35") - GetSoldeCompteN1("35")
    ligne = ligne + 1
    ws.Cells(ligne, 2).Value = "Variation de stocks marchandises (37)"
    ws.Cells(ligne, 3).Value = GetSoldeCompteN("37") - GetSoldeCompteN1("37")

    ws.Columns("C:E").NumberFormat = "#,##0;[Red]-#,##0"
End Sub

' ============================================================
' CREANCES (Classe 4 actif)
' ============================================================
Sub CreerOngletCreances()
    Dim ws As Worksheet
    Dim ligne As Long

    Set ws = CreerOnglet("Creances", COULEUR_BILAN)
    Call EcrireEnTete(ws, "CREANCES D'EXPLOITATION ET HORS CYCLE", "Comptes de classe 4 - Actif circulant")
    ligne = 4

    ' Creances clients
    ws.Cells(ligne, 2).Value = "CREANCES CLIENTS ET COMPTES RATTACHES"
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 7)).Interior.Color = RGB(220, 230, 241)
    ws.Cells(ligne, 2).Font.Bold = True
    ligne = ligne + 1

    Dim compClients() As String
    compClients = Split("411,412,413,416,417,418,419,490", ",")
    Dim labClients() As String
    labClients = Split("Clients,Clients - effets a recevoir,Clients - eff. escomptés non échus,Clients douteux ou litigieux,Clients - autres créances,Clients - fact. a établir,Clients - avances et acomptes,Dépréciations clients", ",")

    Dim totalClientsN As Double, totalClientsN1 As Double
    Dim j As Integer, sN As Double, sN1 As Double

    For j = 0 To UBound(compClients)
        sN = GetSoldeCompteN(compClients(j))
        sN1 = GetSoldeCompteN1(compClients(j))
        If sN <> 0 Or sN1 <> 0 Then
            Call EcrireLigneCompte(ws, ligne, compClients(j), labClients(j), sN, sN1)
            totalClientsN = totalClientsN + sN
            totalClientsN1 = totalClientsN1 + sN1
            ligne = ligne + 1
        End If
    Next j
    Call EcrireSousTotal(ws, ligne, "Total creances clients (net)", totalClientsN, totalClientsN1)
    ligne = ligne + 2

    ' Autres creances
    ws.Cells(ligne, 2).Value = "AUTRES CREANCES"
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 7)).Interior.Color = RGB(220, 230, 241)
    ws.Cells(ligne, 2).Font.Bold = True
    ligne = ligne + 1

    Dim compAutres() As String
    compAutres = Split("421,425,431,437,441,444,445,446,447,448,451,455,456,458,467,486,487", ",")
    Dim labAutres() As String
    labAutres = Split("Personnel - avances,Personnel - congés payés,SS - charges à payer,Autres org. sociaux,Etat - subventions à recevoir,Etat - IS,TVA déductible,Autres impôts,Taxes sur CA,Etat - autres comptes,Groupe et assoc. - créances,Associés - cptes courants,Assoc. - opér. capital,Groupe - inter. débiteurs,Débiteurs divers,Charges constatées d'avance,Produits à recevoir", ",")

    Dim totalAutresN As Double, totalAutresN1 As Double
    For j = 0 To UBound(compAutres)
        sN = GetSoldeCompteN(compAutres(j))
        sN1 = GetSoldeCompteN1(compAutres(j))
        If sN <> 0 Or sN1 <> 0 Then
            Call EcrireLigneCompte(ws, ligne, compAutres(j), labAutres(j), sN, sN1)
            totalAutresN = totalAutresN + sN
            totalAutresN1 = totalAutresN1 + sN1
            ligne = ligne + 1
        End If
    Next j
    Call EcrireSousTotal(ws, ligne, "Total autres creances", totalAutresN, totalAutresN1)
    ligne = ligne + 2

    Dim totalCreancesN As Double, totalCreancesN1 As Double
    totalCreancesN = totalClientsN + totalAutresN
    totalCreancesN1 = totalClientsN1 + totalAutresN1
    Call EcrireTotal(ws, ligne, "TOTAL CREANCES", totalCreancesN, totalCreancesN1)

    ' DSO (Days Sales Outstanding)
    ligne = ligne + 3
    ws.Cells(ligne, 2).Value = "RATIOS CREDIT CLIENT"
    ws.Cells(ligne, 2).Font.Bold = True
    ligne = ligne + 1

    Dim caHT_N As Double, caHT_N1 As Double
    caHT_N = Abs(GetSoldeCompteN("70"))
    caHT_N1 = Abs(GetSoldeCompteN1("70"))

    ws.Cells(ligne, 2).Value = "DSO (jours) = Clients / (CA TTC / 365)"
    If caHT_N <> 0 Then
        ws.Cells(ligne, 3).Value = (Abs(totalClientsN) / (caHT_N * 1.2)) * 365
    End If
    If caHT_N1 <> 0 Then
        ws.Cells(ligne, 4).Value = (Abs(totalClientsN1) / (caHT_N1 * 1.2)) * 365
    End If
    ws.Cells(ligne, 3).NumberFormat = "0.0"
    ws.Cells(ligne, 4).NumberFormat = "0.0"
    ligne = ligne + 1
    ws.Cells(ligne, 2).Value = "Note : DSO calcule sur hypothese TVA 20%"
    ws.Cells(ligne, 2).Font.Italic = True
End Sub

' ============================================================
' TRESORERIE (Classe 5)
' ============================================================
Sub CreerOngletTresorerie()
    Dim ws As Worksheet
    Dim ligne As Long

    Set ws = CreerOnglet("Tresorerie", COULEUR_BILAN)
    Call EcrireEnTete(ws, "TRESORERIE ET EQUIVALENTS DE TRESORERIE", "Comptes de classe 5")
    ligne = 4

    Dim compTreso() As String
    compTreso = Split("50,51,512,514,515,516,517,518,519,53,54,58,59", ",")
    Dim labTreso() As String
    labTreso = Split("Valeurs mobilières de placement,Banques - compte general,Banques,Cheques postaux,Caisses,Titres à court terme,Autres organismes financiers,Intérêts courus,Concours bancaires courants,Caisse,Régies avances accréditifs,Virements internes,Dépréciations VMP", ",")

    Dim totalN As Double, totalN1 As Double
    Dim j As Integer, sN As Double, sN1 As Double

    For j = 0 To UBound(compTreso)
        sN = GetSoldeCompteN(compTreso(j))
        sN1 = GetSoldeCompteN1(compTreso(j))
        If sN <> 0 Or sN1 <> 0 Then
            Call EcrireLigneCompte(ws, ligne, compTreso(j), labTreso(j), sN, sN1)
            If compTreso(j) <> "519" Then  ' 519 = CBC (passif)
                totalN = totalN + sN
                totalN1 = totalN1 + sN1
            End If
            ligne = ligne + 1
        End If
    Next j

    ligne = ligne + 1
    Call EcrireTotal(ws, ligne, "TOTAL TRESORERIE NETTE", totalN, totalN1)

    ligne = ligne + 3
    ws.Cells(ligne, 2).Value = "DETAIL TRESORERIE"
    ws.Cells(ligne, 2).Font.Bold = True
    ligne = ligne + 1
    Dim tresoBruteN As Double, tresoBruteN1 As Double
    Dim cbcN As Double, cbcN1 As Double
    tresoBruteN = GetSoldeCompteN("51") + GetSoldeCompteN("53") + GetSoldeCompteN("54") + GetSoldeCompteN("50")
    tresoBruteN1 = GetSoldeCompteN1("51") + GetSoldeCompteN1("53") + GetSoldeCompteN1("54") + GetSoldeCompteN1("50")
    cbcN = GetSoldeCompteN("519")
    cbcN1 = GetSoldeCompteN1("519")

    ws.Cells(ligne, 2).Value = "Tresorerie active (disponibilites + VMP)"
    ws.Cells(ligne, 3).Value = tresoBruteN
    ws.Cells(ligne, 4).Value = tresoBruteN1
    ligne = ligne + 1
    ws.Cells(ligne, 2).Value = "Concours bancaires courants (519)"
    ws.Cells(ligne, 3).Value = cbcN
    ws.Cells(ligne, 4).Value = cbcN1
    ligne = ligne + 1
    ws.Cells(ligne, 2).Value = "TRESORERIE NETTE"
    ws.Cells(ligne, 3).Value = tresoBruteN + cbcN
    ws.Cells(ligne, 4).Value = tresoBruteN1 + cbcN1
    ws.Cells(ligne, 2).Font.Bold = True
    ws.Columns("C:E").NumberFormat = "#,##0;[Red]-#,##0"
End Sub

' ============================================================
' CAPITAUX PROPRES (Classe 1)
' ============================================================
Sub CreerOngletCapitauxPropres()
    Dim ws As Worksheet
    Dim ligne As Long

    Set ws = CreerOnglet("Capitaux propres", COULEUR_BILAN)
    Call EcrireEnTete(ws, "CAPITAUX PROPRES ET QUASI-FONDS PROPRES", "Comptes de classe 1")
    ligne = 4

    ' Capitaux propres stricts
    ws.Cells(ligne, 2).Value = "CAPITAUX PROPRES"
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 7)).Interior.Color = RGB(220, 230, 241)
    ws.Cells(ligne, 2).Font.Bold = True
    ligne = ligne + 1

    Dim compCP() As String
    compCP = Split("101,1011,1012,1013,104,105,106,107,108,11,12,13,14,15", ",")
    Dim labCP() As String
    labCP = Split("Capital social,Capital souscrit non appelé,Capital souscrit appelé,Capital souscrit appelé versé,Primes d'émission fusion,Ecarts de réévaluation,Réserves,Report à nouveau,Cpte exploitant,Résultat en attente d'affectation,Résultat de l'exercice,Subventions investissement,Provisions réglementées,Amort. dérogatoires", ",")

    Dim totalCPN As Double, totalCPN1 As Double
    Dim j As Integer, sN As Double, sN1 As Double

    For j = 0 To UBound(compCP)
        sN = GetSoldeCompteN(compCP(j))
        sN1 = GetSoldeCompteN1(compCP(j))
        If sN <> 0 Or sN1 <> 0 Then
            Call EcrireLigneCompte(ws, ligne, compCP(j), labCP(j), -sN, -sN1) ' Passif = negatif comptable
            totalCPN = totalCPN + sN
            totalCPN1 = totalCPN1 + sN1
            ligne = ligne + 1
        End If
    Next j
    Call EcrireSousTotal(ws, ligne, "Total capitaux propres", -totalCPN, -totalCPN1)
    ligne = ligne + 2

    ' Provisions pour risques et charges (classe 15)
    ws.Cells(ligne, 2).Value = "PROVISIONS POUR RISQUES ET CHARGES"
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 7)).Interior.Color = RGB(220, 230, 241)
    ws.Cells(ligne, 2).Font.Bold = True
    ligne = ligne + 1

    Dim compProv() As String
    compProv = Split("151,152,153,154,155,156,157,158", ",")
    Dim labProv() As String
    labProv = Split("Provisions pour litiges,Provisions pour garanties,Provisions pour pertes sur marchés,Provisions pour amendes,Provisions pour pertes change,Provisions pour pensions,Provisions pour impôts,Autres provisions pour R&C", ",")

    Dim totalProvN As Double, totalProvN1 As Double
    For j = 0 To UBound(compProv)
        sN = GetSoldeCompteN(compProv(j))
        sN1 = GetSoldeCompteN1(compProv(j))
        If sN <> 0 Or sN1 <> 0 Then
            Call EcrireLigneCompte(ws, ligne, compProv(j), labProv(j), -sN, -sN1)
            totalProvN = totalProvN + sN
            totalProvN1 = totalProvN1 + sN1
            ligne = ligne + 1
        End If
    Next j
    Call EcrireSousTotal(ws, ligne, "Total provisions pour R&C", -totalProvN, -totalProvN1)
    ligne = ligne + 2

    Call EcrireTotal(ws, ligne, "TOTAL RESSOURCES STABLES (hors dettes)", -(totalCPN + totalProvN), -(totalCPN1 + totalProvN1))

    ' Ratios capitaux propres
    ligne = ligne + 3
    ws.Cells(ligne, 2).Value = "INDICATEURS"
    ws.Cells(ligne, 2).Font.Bold = True
    ligne = ligne + 1
    Dim capN As Double, capN1 As Double
    capN = -GetSoldeCompteN("101") - GetSoldeCompteN("104") - GetSoldeCompteN("105") - GetSoldeCompteN("106") - GetSoldeCompteN("12")
    capN1 = -GetSoldeCompteN1("101") - GetSoldeCompteN1("104") - GetSoldeCompteN1("105") - GetSoldeCompteN1("106") - GetSoldeCompteN1("12")
    ws.Cells(ligne, 2).Value = "Capital + Reserves + Resultat"
    ws.Cells(ligne, 3).Value = capN
    ws.Cells(ligne, 4).Value = capN1
    ws.Cells(ligne, 3).NumberFormat = "#,##0;[Red]-#,##0"
    ws.Cells(ligne, 4).NumberFormat = "#,##0;[Red]-#,##0"
End Sub

' ============================================================
' DETTES (Classe 1 et 4 passif)
' ============================================================
Sub CreerOngletDettes()
    Dim ws As Worksheet
    Dim ligne As Long

    Set ws = CreerOnglet("Dettes", COULEUR_BILAN)
    Call EcrireEnTete(ws, "DETTES FINANCIERES ET D'EXPLOITATION", "Comptes de classe 1 et 4 - Passif")
    ligne = 4

    ' Dettes financieres long terme
    ws.Cells(ligne, 2).Value = "DETTES FINANCIERES (LONG ET MOYEN TERME)"
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 7)).Interior.Color = RGB(220, 230, 241)
    ws.Cells(ligne, 2).Font.Bold = True
    ligne = ligne + 1

    Dim compLT() As String
    compLT = Split("161,162,163,164,165,166,167,168,169", ",")
    Dim labLT() As String
    labLT = Split("Emprunts obligataires,Obligations convertibles,Autres emprunts obligataires,Emprunts auprès établissements crédit,Dépôts et cautionnements reçus,Participation salariés,Emprunts et dettes assimilées,Intérêts courus,Primes remboursement obligations", ",")

    Dim totalLTN As Double, totalLTN1 As Double
    Dim j As Integer, sN As Double, sN1 As Double

    For j = 0 To UBound(compLT)
        sN = GetSoldeCompteN(compLT(j))
        sN1 = GetSoldeCompteN1(compLT(j))
        If sN <> 0 Or sN1 <> 0 Then
            Call EcrireLigneCompte(ws, ligne, compLT(j), labLT(j), -sN, -sN1)
            totalLTN = totalLTN + sN
            totalLTN1 = totalLTN1 + sN1
            ligne = ligne + 1
        End If
    Next j
    Call EcrireSousTotal(ws, ligne, "Total dettes financières LMT", -totalLTN, -totalLTN1)
    ligne = ligne + 2

    ' Dettes fournisseurs
    ws.Cells(ligne, 2).Value = "DETTES FOURNISSEURS ET COMPTES RATTACHES"
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 7)).Interior.Color = RGB(220, 230, 241)
    ws.Cells(ligne, 2).Font.Bold = True
    ligne = ligne + 1

    Dim compFourn() As String
    compFourn = Split("401,402,403,404,405,408,409", ",")
    Dim labFourn() As String
    labFourn = Split("Fournisseurs,Fournisseurs - biens,Fournisseurs - eff. à payer,Fournisseurs d'immobilisations,Fourn. immob. - eff. à payer,Fourn. - fact. non parvenues,Fournisseurs - avances et acomptes", ",")

    Dim totalFournN As Double, totalFournN1 As Double
    For j = 0 To UBound(compFourn)
        sN = GetSoldeCompteN(compFourn(j))
        sN1 = GetSoldeCompteN1(compFourn(j))
        If sN <> 0 Or sN1 <> 0 Then
            Call EcrireLigneCompte(ws, ligne, compFourn(j), labFourn(j), -sN, -sN1)
            totalFournN = totalFournN + sN
            totalFournN1 = totalFournN1 + sN1
            ligne = ligne + 1
        End If
    Next j
    Call EcrireSousTotal(ws, ligne, "Total dettes fournisseurs", -totalFournN, -totalFournN1)
    ligne = ligne + 2

    ' Dettes fiscales et sociales
    ws.Cells(ligne, 2).Value = "DETTES FISCALES ET SOCIALES"
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 7)).Interior.Color = RGB(220, 230, 241)
    ws.Cells(ligne, 2).Font.Bold = True
    ligne = ligne + 1

    Dim compFS() As String
    compFS = Split("421,422,423,424,425,426,427,428,431,432,437,438,441,442,443,444,445,447,448", ",")
    Dim labFS() As String
    labFS = Split("Personnel - rémunérations,Œuvres sociales,Participations salariés,Part. salariés - cpt bloqués,Rémun. congés à payer,Personnel - autres créditeurs,Dépôts garanties personnel,Charges à payer personnel,SS - URSSAF,SS - mutuelles,Autres org. sociaux,Charges SS à payer,Etat - IS,TVA à décaisser,Autres impôts indirects,Etat - IS à payer,TVA collectée,Taxes parafiscales,Impôts à payer", ",")

    Dim totalFSN As Double, totalFSN1 As Double
    For j = 0 To UBound(compFS)
        sN = GetSoldeCompteN(compFS(j))
        sN1 = GetSoldeCompteN1(compFS(j))
        If sN <> 0 Or sN1 <> 0 Then
            Call EcrireLigneCompte(ws, ligne, compFS(j), labFS(j), -sN, -sN1)
            totalFSN = totalFSN + sN
            totalFSN1 = totalFSN1 + sN1
            ligne = ligne + 1
        End If
    Next j
    Call EcrireSousTotal(ws, ligne, "Total dettes fiscales et sociales", -totalFSN, -totalFSN1)
    ligne = ligne + 2

    ' Autres dettes
    Dim autresDettesN As Double, autresDettesN1 As Double
    autresDettesN = GetSoldeCompteN("455") + GetSoldeCompteN("456") + GetSoldeCompteN("457") + GetSoldeCompteN("467") + GetSoldeCompteN("486")
    autresDettesN1 = GetSoldeCompteN1("455") + GetSoldeCompteN1("456") + GetSoldeCompteN1("457") + GetSoldeCompteN1("467") + GetSoldeCompteN1("486")
    If autresDettesN <> 0 Or autresDettesN1 <> 0 Then
        Call EcrireLigneCompte(ws, ligne, "455-467", "Autres dettes (associes, divers)", -autresDettesN, -autresDettesN1)
        ligne = ligne + 1
    End If

    ' CBC
    Dim cbcN As Double, cbcN1 As Double
    cbcN = GetSoldeCompteN("519")
    cbcN1 = GetSoldeCompteN1("519")
    If cbcN <> 0 Or cbcN1 <> 0 Then
        Call EcrireLigneCompte(ws, ligne, "519", "Concours bancaires courants", -cbcN, -cbcN1)
        ligne = ligne + 1
    End If

    ligne = ligne + 1
    Dim totalDettesN As Double, totalDettesN1 As Double
    totalDettesN = totalLTN + totalFournN + totalFSN + autresDettesN + cbcN
    totalDettesN1 = totalLTN1 + totalFournN1 + totalFSN1 + autresDettesN1 + cbcN1
    Call EcrireTotal(ws, ligne, "TOTAL DETTES", -totalDettesN, -totalDettesN1)

    ' DPO (Days Payable Outstanding)
    ligne = ligne + 3
    ws.Cells(ligne, 2).Value = "RATIOS CREDIT FOURNISSEUR (DPO)"
    ws.Cells(ligne, 2).Font.Bold = True
    ligne = ligne + 1
    Dim achatsN As Double, achatsN1 As Double
    achatsN = Abs(GetSoldeCompteN("60")) + Abs(GetSoldeCompteN("61")) + Abs(GetSoldeCompteN("62"))
    achatsN1 = Abs(GetSoldeCompteN1("60")) + Abs(GetSoldeCompteN1("61")) + Abs(GetSoldeCompteN1("62"))
    ws.Cells(ligne, 2).Value = "DPO (jours) = Fournisseurs / (Achats TTC / 365)"
    If achatsN <> 0 Then ws.Cells(ligne, 3).Value = (Abs(totalFournN) / (achatsN * 1.2)) * 365
    If achatsN1 <> 0 Then ws.Cells(ligne, 4).Value = (Abs(totalFournN1) / (achatsN1 * 1.2)) * 365
    ws.Cells(ligne, 3).NumberFormat = "0.0"
    ws.Cells(ligne, 4).NumberFormat = "0.0"
End Sub

' ============================================================
' PROVISIONS POUR RISQUES ET CHARGES
' ============================================================
Sub CreerOngletProvisionsRisques()
    Dim ws As Worksheet
    Dim ligne As Long

    Set ws = CreerOnglet("Provisions", COULEUR_BILAN)
    Call EcrireEnTete(ws, "PROVISIONS POUR RISQUES ET CHARGES", "Comptes 15x - Mouvements de l'exercice")
    ligne = 4

    ' En-tetes specifiques pour le tableau de mouvement
    ws.Cells(3, 1).Value = "N° Compte"
    ws.Cells(3, 2).Value = "Nature provision"
    ws.Cells(3, 3).Value = "Debut N"
    ws.Cells(3, 4).Value = "Dotations N"
    ws.Cells(3, 5).Value = "Reprises N"
    ws.Cells(3, 6).Value = "Fin N"
    ws.Cells(3, 7).Value = "Debut N-1"
    ws.Cells(3, 8).Value = "Fin N-1"

    With ws.Range("A3:H3")
        .Font.Bold = True
        .Interior.Color = RGB(189, 215, 238)
    End With

    Dim compProv() As String
    compProv = Split("151,152,153,154,155,156,157,158", ",")
    Dim labProv() As String
    labProv = Split("Provisions pour litiges,Provisions pour garanties données,Provisions pour pertes sur marchés à terme,Provisions pour amendes et pénalités,Provisions pour pertes de change,Provisions pour pensions et obligations similaires,Provisions pour impôts,Autres provisions pour R&C", ",")

    Dim j As Integer, sN As Double, sN1 As Double
    For j = 0 To UBound(compProv)
        sN = -GetSoldeCompteN(compProv(j))
        sN1 = -GetSoldeCompteN1(compProv(j))
        If sN <> 0 Or sN1 <> 0 Then
            ws.Cells(ligne, 1).Value = compProv(j)
            ws.Cells(ligne, 2).Value = labProv(j)
            ws.Cells(ligne, 3).Value = sN1   ' Debut N = Fin N-1
            ws.Cells(ligne, 6).Value = sN    ' Fin N
            ws.Cells(ligne, 7).Value = sN1
            ws.Cells(ligne, 8).Value = sN
            ' Variation (dotation - reprise = fin - debut)
            ws.Cells(ligne, 4).Value = IIf(sN > sN1, sN - sN1, 0)
            ws.Cells(ligne, 5).Value = IIf(sN < sN1, sN1 - sN, 0)
            If ligne Mod 2 = 0 Then ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 8)).Interior.Color = RGB(242, 242, 242)
            ligne = ligne + 1
        End If
    Next j

    ' Total
    ws.Cells(ligne, 2).Value = "TOTAL PROVISIONS RISQUES & CHARGES"
    Dim tN As Double, tN1 As Double
    tN = -GetSoldeCompteN("15")
    tN1 = -GetSoldeCompteN1("15")
    ws.Cells(ligne, 3).Value = tN1
    ws.Cells(ligne, 6).Value = tN
    ws.Cells(ligne, 7).Value = tN1
    ws.Cells(ligne, 8).Value = tN
    ws.Cells(ligne, 4).Value = IIf(tN > tN1, tN - tN1, 0)
    ws.Cells(ligne, 5).Value = IIf(tN < tN1, tN1 - tN, 0)
    With ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 8))
        .Font.Bold = True
        .Interior.Color = RGB(31, 73, 125)
        .Font.Color = RGB(255, 255, 255)
    End With

    ws.Columns("C:H").NumberFormat = "#,##0;[Red]-#,##0"
    ws.Columns("C:H").ColumnWidth = 14
    ws.Columns("B").ColumnWidth = 45
End Sub
