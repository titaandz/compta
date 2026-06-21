Attribute VB_Name = "Module3_PL"
Option Explicit

' ============================================================
' ONGLETS COMPTE DE RESULTAT (P&L)
' ============================================================

Sub CreerOngletCA()
    Dim ws As Worksheet
    Dim ligne As Long

    Set ws = CreerOnglet("CA", COULEUR_PL)
    Call EcrireEnTete(ws, "CHIFFRE D'AFFAIRES ET PRODUITS ANNEXES", "Comptes 70x")
    ligne = 4

    Dim comp70() As String
    comp70 = Split("700,701,702,703,704,705,706,707,708,709", ",")
    Dim lab70() As String
    lab70 = Split("Achats revendus de marchandises,Ventes de marchandises,Ventes produits finis,Ventes produits intermédiaires,Ventes produits résiduels,Études et prestations services,Redevances et droits de licence,Locations et charges locatives,Produits des activités annexes,Rabais remises et ristournes accordés", ",")

    Dim totalCAN As Double, totalCAN1 As Double
    Dim j As Integer, sN As Double, sN1 As Double

    For j = 0 To UBound(comp70)
        sN = -GetSoldeCompteN(comp70(j))
        sN1 = -GetSoldeCompteN1(comp70(j))
        If sN <> 0 Or sN1 <> 0 Then
            Call EcrireLigneCompte(ws, ligne, comp70(j), lab70(j), sN, sN1)
            totalCAN = totalCAN + sN
            totalCAN1 = totalCAN1 + sN1
            ligne = ligne + 1
        End If
    Next j

    Call EcrireTotal(ws, ligne, "TOTAL CHIFFRE D'AFFAIRES HT", totalCAN, totalCAN1)
    ligne = ligne + 3

    ' Analyse par nature de CA
    ws.Cells(ligne, 2).Value = "ANALYSE DU CA"
    ws.Cells(ligne, 2).Font.Bold = True
    ligne = ligne + 1

    Dim venteMarchandisesN As Double, venteProduitsN As Double, prestationsN As Double
    Dim venteMarchandisesN1 As Double, venteProduitsN1 As Double, prestationsN1 As Double

    venteMarchandisesN = -GetSoldeCompteN("707")
    venteProduitsN = -GetSoldeCompteN("701") - GetSoldeCompteN("702") - GetSoldeCompteN("703") - GetSoldeCompteN("704")
    prestationsN = -GetSoldeCompteN("705") - GetSoldeCompteN("706") - GetSoldeCompteN("708")

    venteMarchandisesN1 = -GetSoldeCompteN1("707")
    venteProduitsN1 = -GetSoldeCompteN1("701") - GetSoldeCompteN1("702") - GetSoldeCompteN1("703") - GetSoldeCompteN1("704")
    prestationsN1 = -GetSoldeCompteN1("705") - GetSoldeCompteN1("706") - GetSoldeCompteN1("708")

    ws.Cells(ligne, 2).Value = "Ventes de marchandises (707)"
    ws.Cells(ligne, 3).Value = venteMarchandisesN
    ws.Cells(ligne, 4).Value = venteMarchandisesN1
    If totalCAN <> 0 Then ws.Cells(ligne, 7).Value = venteMarchandisesN / totalCAN
    ws.Cells(ligne, 7).NumberFormat = "0.0%"
    ligne = ligne + 1

    ws.Cells(ligne, 2).Value = "Ventes de produits (701-704)"
    ws.Cells(ligne, 3).Value = venteProduitsN
    ws.Cells(ligne, 4).Value = venteProduitsN1
    If totalCAN <> 0 Then ws.Cells(ligne, 7).Value = venteProduitsN / totalCAN
    ws.Cells(ligne, 7).NumberFormat = "0.0%"
    ligne = ligne + 1

    ws.Cells(ligne, 2).Value = "Prestations de services (705-706-708)"
    ws.Cells(ligne, 3).Value = prestationsN
    ws.Cells(ligne, 4).Value = prestationsN1
    If totalCAN <> 0 Then ws.Cells(ligne, 7).Value = prestationsN / totalCAN
    ws.Cells(ligne, 7).NumberFormat = "0.0%"
    ligne = ligne + 2

    ' Variation annuelle
    ws.Cells(ligne, 2).Value = "Croissance CA annuelle"
    ws.Cells(ligne, 2).Font.Bold = True
    If totalCAN1 <> 0 Then
        ws.Cells(ligne, 3).Value = (totalCAN - totalCAN1) / Abs(totalCAN1)
        ws.Cells(ligne, 3).NumberFormat = "0.0%"
        ws.Cells(ligne, 3).Font.Bold = True
    End If
End Sub

Sub CreerOngletProduitsExploitation()
    Dim ws As Worksheet
    Dim ligne As Long

    Set ws = CreerOnglet("Produits exploitation", COULEUR_PL)
    Call EcrireEnTete(ws, "PRODUITS D'EXPLOITATION (HORS CA)", "Comptes 71x à 75x")
    ligne = 4

    ' Production stockee (71)
    ws.Cells(ligne, 2).Value = "PRODUCTION STOCKEE (71)"
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 7)).Interior.Color = RGB(198, 224, 180)
    ws.Cells(ligne, 2).Font.Bold = True
    ligne = ligne + 1
    Dim ps71N As Double, ps71N1 As Double
    ps71N = -GetSoldeCompteN("71")
    ps71N1 = -GetSoldeCompteN1("71")
    If ps71N <> 0 Or ps71N1 <> 0 Then
        Call EcrireLigneCompte(ws, ligne, "71", "Variation de stock de produits", ps71N, ps71N1)
        ligne = ligne + 1
    End If

    ' Production immobilisee (72)
    ws.Cells(ligne, 2).Value = "PRODUCTION IMMOBILISEE (72)"
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 7)).Interior.Color = RGB(198, 224, 180)
    ws.Cells(ligne, 2).Font.Bold = True
    ligne = ligne + 1
    Dim pi72N As Double, pi72N1 As Double
    pi72N = -GetSoldeCompteN("72")
    pi72N1 = -GetSoldeCompteN1("72")
    If pi72N <> 0 Or pi72N1 <> 0 Then
        Call EcrireLigneCompte(ws, ligne, "72", "Production immobilisée", pi72N, pi72N1)
        ligne = ligne + 1
    End If

    ' Subventions d'exploitation (74)
    ws.Cells(ligne, 2).Value = "SUBVENTIONS D'EXPLOITATION (74)"
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 7)).Interior.Color = RGB(198, 224, 180)
    ws.Cells(ligne, 2).Font.Bold = True
    ligne = ligne + 1
    Dim sub74N As Double, sub74N1 As Double
    sub74N = -GetSoldeCompteN("74")
    sub74N1 = -GetSoldeCompteN1("74")
    If sub74N <> 0 Or sub74N1 <> 0 Then
        Call EcrireLigneCompte(ws, ligne, "74", "Subventions d'exploitation", sub74N, sub74N1)
        ligne = ligne + 1
    End If

    ' Autres produits gestion courante (75)
    ws.Cells(ligne, 2).Value = "AUTRES PRODUITS DE GESTION COURANTE (75)"
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 7)).Interior.Color = RGB(198, 224, 180)
    ws.Cells(ligne, 2).Font.Bold = True
    ligne = ligne + 1
    Dim comp75() As String
    comp75 = Split("751,752,753,754,755,758", ",")
    Dim lab75() As String
    lab75 = Split("Redevances pour concessions,Revenus immeubles non affectés,Jetons de présence,Bonis sur reprises,Quotes-parts bénéfice coentreprise,Produits gestion courante divers", ",")
    Dim total75N As Double, total75N1 As Double
    Dim j As Integer, sN As Double, sN1 As Double
    For j = 0 To UBound(comp75)
        sN = -GetSoldeCompteN(comp75(j))
        sN1 = -GetSoldeCompteN1(comp75(j))
        If sN <> 0 Or sN1 <> 0 Then
            Call EcrireLigneCompte(ws, ligne, comp75(j), lab75(j), sN, sN1)
            total75N = total75N + sN
            total75N1 = total75N1 + sN1
            ligne = ligne + 1
        End If
    Next j

    ' Reprises sur provisions d'exploitation (781)
    ws.Cells(ligne, 2).Value = "REPRISES SUR AMORTISSEMENTS ET PROVISIONS D'EXPLOITATION (781)"
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 7)).Interior.Color = RGB(198, 224, 180)
    ws.Cells(ligne, 2).Font.Bold = True
    ligne = ligne + 1
    Dim rep781N As Double, rep781N1 As Double
    rep781N = -GetSoldeCompteN("781")
    rep781N1 = -GetSoldeCompteN1("781")
    If rep781N <> 0 Or rep781N1 <> 0 Then
        Call EcrireLigneCompte(ws, ligne, "781", "Reprises sur prov. exploitation", rep781N, rep781N1)
        ligne = ligne + 1
    End If

    ' Autres produits exploitation (79)
    Dim autres79N As Double, autres79N1 As Double
    autres79N = -GetSoldeCompteN("791") - GetSoldeCompteN("796")
    autres79N1 = -GetSoldeCompteN1("791") - GetSoldeCompteN1("796")
    If autres79N <> 0 Or autres79N1 <> 0 Then
        Call EcrireLigneCompte(ws, ligne, "791/796", "Transferts de charges exploitation", autres79N, autres79N1)
        ligne = ligne + 1
    End If

    ligne = ligne + 1
    Dim totalProduitsExplN As Double, totalProduitsExplN1 As Double
    totalProduitsExplN = ps71N + pi72N + sub74N + total75N + rep781N + autres79N
    totalProduitsExplN1 = ps71N1 + pi72N1 + sub74N1 + total75N1 + rep781N1 + autres79N1
    Call EcrireTotal(ws, ligne, "TOTAL PRODUITS EXPLOITATION (HORS CA)", totalProduitsExplN, totalProduitsExplN1)

    ligne = ligne + 2
    ws.Cells(ligne, 2).Value = "PRODUCTION DE L'EXERCICE (CA + Prod.Stockee + Prod.Immob.)"
    ws.Cells(ligne, 2).Font.Bold = True
    Dim caN As Double, caN1 As Double
    caN = -GetSoldeCompteN("70") - GetSoldeCompteN("71") - GetSoldeCompteN("72")
    caN1 = -GetSoldeCompteN1("70") - GetSoldeCompteN1("71") - GetSoldeCompteN1("72")
    ws.Cells(ligne, 3).Value = caN
    ws.Cells(ligne, 4).Value = caN1
    ws.Cells(ligne, 5).Value = caN - caN1
    ws.Cells(ligne, 3).Font.Bold = True
    ws.Cells(ligne, 4).Font.Bold = True
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 7)).Interior.Color = RGB(169, 208, 142)
    ws.Columns("C:E").NumberFormat = "#,##0;[Red]-#,##0"
End Sub

Sub CreerOngletChargesExploitation()
    Dim ws As Worksheet
    Dim ligne As Long

    Set ws = CreerOnglet("Charges exploitation", COULEUR_PL)
    Call EcrireEnTete(ws, "CHARGES D'EXPLOITATION", "Comptes 60x à 65x")
    ligne = 4

    ' Achats marchandises
    ws.Cells(ligne, 2).Value = "ACHATS ET VARIATION DE STOCKS"
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 7)).Interior.Color = RGB(255, 230, 153)
    ws.Cells(ligne, 2).Font.Bold = True
    ligne = ligne + 1

    Dim comp60() As String
    comp60 = Split("601,602,603,604,605,606,607,608,609", ",")
    Dim lab60() As String
    lab60 = Split("Achats matières premières,Achats autres approvisionnements,Var. stocks MP et appros,Achats études et prestations,Achats matériels et équipements,Achats non stockés,Achats marchandises,Frais accessoires achats,Rabais remises ristournes obtenus", ",")

    Dim total60N As Double, total60N1 As Double
    Dim j As Integer, sN As Double, sN1 As Double

    For j = 0 To UBound(comp60)
        sN = GetSoldeCompteN(comp60(j))
        sN1 = GetSoldeCompteN1(comp60(j))
        If sN <> 0 Or sN1 <> 0 Then
            Call EcrireLigneCompte(ws, ligne, comp60(j), lab60(j), sN, sN1)
            total60N = total60N + sN
            total60N1 = total60N1 + sN1
            ligne = ligne + 1
        End If
    Next j
    Call EcrireSousTotal(ws, ligne, "Total achats et variation stocks", total60N, total60N1)
    ligne = ligne + 2

    ' Services exterieurs (61-62)
    ws.Cells(ligne, 2).Value = "SERVICES EXTERIEURS (61-62)"
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 7)).Interior.Color = RGB(255, 230, 153)
    ws.Cells(ligne, 2).Font.Bold = True
    ligne = ligne + 1

    Dim comp61() As String
    comp61 = Split("611,612,613,614,615,616,617,618,619,621,622,623,624,625,626,627,628", ",")
    Dim lab61() As String
    lab61 = Split("Sous-traitance,Redevances crédit-bail,Locations,Charges locatives,Entretiens et réparations,Primes d'assurances,Études et recherches,Divers (61),Rabais obtenus (61),Personnels intérimaires,Rémunérations intermédiaires,Publicité et publications,Transports biens tiers,Déplacements et réceptions,Frais postaux télécoms,Banques et établissements fin.,Divers (62)", ",")

    Dim total6162N As Double, total6162N1 As Double
    For j = 0 To UBound(comp61)
        sN = GetSoldeCompteN(comp61(j))
        sN1 = GetSoldeCompteN1(comp61(j))
        If sN <> 0 Or sN1 <> 0 Then
            Call EcrireLigneCompte(ws, ligne, comp61(j), lab61(j), sN, sN1)
            total6162N = total6162N + sN
            total6162N1 = total6162N1 + sN1
            ligne = ligne + 1
        End If
    Next j
    Call EcrireSousTotal(ws, ligne, "Total services exterieurs", total6162N, total6162N1)
    ligne = ligne + 2

    ' Impots taxes (63)
    ws.Cells(ligne, 2).Value = "IMPOTS ET TAXES (63)"
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 7)).Interior.Color = RGB(255, 230, 153)
    ws.Cells(ligne, 2).Font.Bold = True
    ligne = ligne + 1

    Dim comp63() As String
    comp63 = Split("631,632,633,635,636,637,638", ",")
    Dim lab63() As String
    lab63 = Split("Taxes sur rémunérations (apprentissage),Cotisations syndicales,Impôts personnels,Autres impôts taxes,Taxes spécifiques,Taxe foncière,Divers (63)", ",")

    Dim total63N As Double, total63N1 As Double
    For j = 0 To UBound(comp63)
        sN = GetSoldeCompteN(comp63(j))
        sN1 = GetSoldeCompteN1(comp63(j))
        If sN <> 0 Or sN1 <> 0 Then
            Call EcrireLigneCompte(ws, ligne, comp63(j), lab63(j), sN, sN1)
            total63N = total63N + sN
            total63N1 = total63N1 + sN1
            ligne = ligne + 1
        End If
    Next j
    Call EcrireSousTotal(ws, ligne, "Total impots et taxes", total63N, total63N1)
    ligne = ligne + 2

    ' Autres charges gestion courante (65)
    ws.Cells(ligne, 2).Value = "AUTRES CHARGES DE GESTION COURANTE (65)"
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 7)).Interior.Color = RGB(255, 230, 153)
    ws.Cells(ligne, 2).Font.Bold = True
    ligne = ligne + 1

    Dim comp65() As String
    comp65 = Split("651,652,653,654,655,658", ",")
    Dim lab65() As String
    lab65 = Split("Redevances concessions brevets,Charges aff. sol autrui,Jetons de présence,Pertes sur créances irrécouvrables,Quotes-parts pertes,Autres charges", ",")

    Dim total65N As Double, total65N1 As Double
    For j = 0 To UBound(comp65)
        sN = GetSoldeCompteN(comp65(j))
        sN1 = GetSoldeCompteN1(comp65(j))
        If sN <> 0 Or sN1 <> 0 Then
            Call EcrireLigneCompte(ws, ligne, comp65(j), lab65(j), sN, sN1)
            total65N = total65N + sN
            total65N1 = total65N1 + sN1
            ligne = ligne + 1
        End If
    Next j
    Call EcrireSousTotal(ws, ligne, "Total autres charges gestion courante", total65N, total65N1)
    ligne = ligne + 2

    Dim totalChargesExplN As Double, totalChargesExplN1 As Double
    totalChargesExplN = total60N + total6162N + total63N + total65N
    totalChargesExplN1 = total60N1 + total6162N1 + total63N1 + total65N1
    Call EcrireTotal(ws, ligne, "TOTAL CHARGES EXPLOITATION (HORS PERS. ET DOT.)", totalChargesExplN, totalChargesExplN1)

    ws.Columns("C:E").NumberFormat = "#,##0;[Red]-#,##0"

    ' Ratio : Achat / CA
    ligne = ligne + 3
    ws.Cells(ligne, 2).Value = "RATIOS"
    ws.Cells(ligne, 2).Font.Bold = True
    ligne = ligne + 1
    Dim caN As Double, caN1 As Double
    caN = Abs(GetSoldeCompteN("70"))
    caN1 = Abs(GetSoldeCompteN1("70"))
    ws.Cells(ligne, 2).Value = "Achats / CA HT"
    If caN <> 0 Then ws.Cells(ligne, 3).Value = total60N / caN
    If caN1 <> 0 Then ws.Cells(ligne, 4).Value = total60N1 / caN1
    ws.Cells(ligne, 3).NumberFormat = "0.0%"
    ws.Cells(ligne, 4).NumberFormat = "0.0%"
    ligne = ligne + 1
    ws.Cells(ligne, 2).Value = "Services exterieurs / CA HT"
    If caN <> 0 Then ws.Cells(ligne, 3).Value = total6162N / caN
    If caN1 <> 0 Then ws.Cells(ligne, 4).Value = total6162N1 / caN1
    ws.Cells(ligne, 3).NumberFormat = "0.0%"
    ws.Cells(ligne, 4).NumberFormat = "0.0%"
End Sub

Sub CreerOngletPersonnel()
    Dim ws As Worksheet
    Dim ligne As Long

    Set ws = CreerOnglet("Personnel", COULEUR_PL)
    Call EcrireEnTete(ws, "CHARGES DE PERSONNEL", "Comptes 64x")
    ligne = 4

    Dim comp64() As String
    comp64 = Split("641,6411,6412,6413,6414,642,6421,6422,643,6431,6432,6441,6442,6443,6444,6445,6448,645,6451,6453,6454,6458,646,647,648", ",")
    Dim lab64() As String
    lab64 = Split("Rémunérations personnel,Salaires bruts,Congés payés,Primes et gratifications,Indemnités,Rémunérations dirigeants,Gérants et associés,Administrateurs,Charges sociales,Cotisations URSSAF,Cotisations retraites,Mutuelle,Médecine du travail,Prévoyance,Œuvres sociales,ASSEDIC,Autres charges soc.,Versements organismes,Taxe apprentissage,Participation formation,Taxe transport,Autres versements,Autres charges sociales,Autres charges de personnel,Charges pers. bénévole,Divers pers.", ",")

    Dim totalN As Double, totalN1 As Double
    Dim j As Integer, sN As Double, sN1 As Double

    ' Salaires bruts
    Dim salairesBrutsN As Double, salairesBrutsN1 As Double
    salairesBrutsN = GetSoldeCompteN("641")
    salairesBrutsN1 = GetSoldeCompteN1("641")

    ' Charges sociales
    Dim chargesSocN As Double, chargesSocN1 As Double
    chargesSocN = GetSoldeCompteN("645") + GetSoldeCompteN("646") + GetSoldeCompteN("647")
    chargesSocN1 = GetSoldeCompteN1("645") + GetSoldeCompteN1("646") + GetSoldeCompteN1("647")

    For j = 0 To UBound(comp64)
        sN = GetSoldeCompteN(comp64(j))
        sN1 = GetSoldeCompteN1(comp64(j))
        If sN <> 0 Or sN1 <> 0 Then
            Call EcrireLigneCompte(ws, ligne, comp64(j), lab64(j), sN, sN1)
            totalN = totalN + sN
            totalN1 = totalN1 + sN1
            ligne = ligne + 1
        End If
    Next j

    ligne = ligne + 1
    Call EcrireTotal(ws, ligne, "TOTAL CHARGES DE PERSONNEL", totalN, totalN1)

    ' Participations et interessements
    ligne = ligne + 2
    Dim partN As Double, partN1 As Double
    partN = GetSoldeCompteN("691")
    partN1 = GetSoldeCompteN1("691")
    If partN <> 0 Or partN1 <> 0 Then
        ws.Cells(ligne, 2).Value = "Participation des salaries (691)"
        ws.Cells(ligne, 3).Value = partN
        ws.Cells(ligne, 4).Value = partN1
        ws.Cells(ligne, 5).Value = partN - partN1
        If partN1 <> 0 Then ws.Cells(ligne, 6).Value = (partN - partN1) / Abs(partN1)
        ws.Cells(ligne, 2).Font.Italic = True
        ligne = ligne + 1
    End If

    ' Ratios charges de personnel
    ligne = ligne + 2
    ws.Cells(ligne, 2).Value = "RATIOS CHARGES DE PERSONNEL"
    ws.Cells(ligne, 2).Font.Bold = True
    ligne = ligne + 1

    Dim caN As Double, caN1 As Double
    caN = Abs(GetSoldeCompteN("70"))
    caN1 = Abs(GetSoldeCompteN1("70"))

    Dim vaExpl As Double, vaExpl1 As Double
    vaExpl = caN - GetSoldeCompteN("60") - GetSoldeCompteN("61") - GetSoldeCompteN("62")
    vaExpl1 = caN1 - GetSoldeCompteN1("60") - GetSoldeCompteN1("61") - GetSoldeCompteN1("62")

    ws.Cells(ligne, 2).Value = "Salaires bruts"
    ws.Cells(ligne, 3).Value = salairesBrutsN
    ws.Cells(ligne, 4).Value = salairesBrutsN1
    ligne = ligne + 1

    ws.Cells(ligne, 2).Value = "Charges sociales"
    ws.Cells(ligne, 3).Value = chargesSocN
    ws.Cells(ligne, 4).Value = chargesSocN1
    ligne = ligne + 1

    ws.Cells(ligne, 2).Value = "Masse salariale / CA HT"
    If caN <> 0 Then ws.Cells(ligne, 3).Value = totalN / caN
    If caN1 <> 0 Then ws.Cells(ligne, 4).Value = totalN1 / caN1
    ws.Cells(ligne, 3).NumberFormat = "0.0%"
    ws.Cells(ligne, 4).NumberFormat = "0.0%"
    ligne = ligne + 1

    ws.Cells(ligne, 2).Value = "Charges sociales / Salaires bruts (taux de charges)"
    If salairesBrutsN <> 0 Then ws.Cells(ligne, 3).Value = chargesSocN / salairesBrutsN
    If salairesBrutsN1 <> 0 Then ws.Cells(ligne, 4).Value = chargesSocN1 / salairesBrutsN1
    ws.Cells(ligne, 3).NumberFormat = "0.0%"
    ws.Cells(ligne, 4).NumberFormat = "0.0%"
    ligne = ligne + 1

    ws.Cells(ligne, 2).Value = "Personnel / Valeur Ajoutee"
    If vaExpl <> 0 Then ws.Cells(ligne, 3).Value = totalN / vaExpl
    If vaExpl1 <> 0 Then ws.Cells(ligne, 4).Value = totalN1 / vaExpl1
    ws.Cells(ligne, 3).NumberFormat = "0.0%"
    ws.Cells(ligne, 4).NumberFormat = "0.0%"

    ws.Columns("C:E").NumberFormat = "#,##0;[Red]-#,##0"
    ws.Cells(ligne, 3).NumberFormat = "0.0%"
    ws.Cells(ligne, 4).NumberFormat = "0.0%"
End Sub

Sub CreerOngletDotationsAmortissements()
    Dim ws As Worksheet
    Dim ligne As Long

    Set ws = CreerOnglet("Dotations", COULEUR_PL)
    Call EcrireEnTete(ws, "DOTATIONS AUX AMORTISSEMENTS ET PROVISIONS", "Comptes 68x")
    ligne = 4

    ' Amortissements
    ws.Cells(ligne, 2).Value = "DOTATIONS AUX AMORTISSEMENTS"
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 7)).Interior.Color = RGB(255, 230, 153)
    ws.Cells(ligne, 2).Font.Bold = True
    ligne = ligne + 1

    Dim comp681() As String
    comp681 = Split("6811,6812,6813,6815,6816,6817", ",")
    Dim lab681() As String
    lab681 = Split("Dot. amort. frais établissement,Dot. amort. frais R&D,Dot. amort. immob. incorporelles,Dot. amort. immob. corporelles,Dot. dépréc. immob. financières,Dot. dépréc. immob. corpus", ",")

    Dim totalAmortN As Double, totalAmortN1 As Double
    Dim j As Integer, sN As Double, sN1 As Double

    For j = 0 To UBound(comp681)
        sN = GetSoldeCompteN(comp681(j))
        sN1 = GetSoldeCompteN1(comp681(j))
        If sN <> 0 Or sN1 <> 0 Then
            Call EcrireLigneCompte(ws, ligne, comp681(j), lab681(j), sN, sN1)
            totalAmortN = totalAmortN + sN
            totalAmortN1 = totalAmortN1 + sN1
            ligne = ligne + 1
        End If
    Next j
    Call EcrireSousTotal(ws, ligne, "Total dotations amortissements", totalAmortN, totalAmortN1)
    ligne = ligne + 2

    ' Provisions d'exploitation
    ws.Cells(ligne, 2).Value = "DOTATIONS AUX PROVISIONS D'EXPLOITATION"
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 7)).Interior.Color = RGB(255, 230, 153)
    ws.Cells(ligne, 2).Font.Bold = True
    ligne = ligne + 1

    Dim comp6817() As String
    comp6817 = Split("6817,6819", ",")
    Dim lab6817() As String
    lab6817 = Split("Dot. dépréciations actif circulant,Dot. provisions R&C exploitation", ",")

    Dim totalProvExplN As Double, totalProvExplN1 As Double
    For j = 0 To UBound(comp6817)
        sN = GetSoldeCompteN(comp6817(j))
        sN1 = GetSoldeCompteN1(comp6817(j))
        If sN <> 0 Or sN1 <> 0 Then
            Call EcrireLigneCompte(ws, ligne, comp6817(j), lab6817(j), sN, sN1)
            totalProvExplN = totalProvExplN + sN
            totalProvExplN1 = totalProvExplN1 + sN1
            ligne = ligne + 1
        End If
    Next j
    Call EcrireSousTotal(ws, ligne, "Total provisions exploitation", totalProvExplN, totalProvExplN1)
    ligne = ligne + 2

    Dim totalDotN As Double, totalDotN1 As Double
    totalDotN = totalAmortN + totalProvExplN
    totalDotN1 = totalAmortN1 + totalProvExplN1
    Call EcrireTotal(ws, ligne, "TOTAL DOTATIONS (exploitation)", totalDotN, totalDotN1)

    ' Impact sur la CAF
    ligne = ligne + 3
    ws.Cells(ligne, 2).Value = "IMPACT SUR LA CAF"
    ws.Cells(ligne, 2).Font.Bold = True
    ligne = ligne + 1
    ws.Cells(ligne, 2).Value = "Les dotations aux amort. et prov. sont des charges calculees"
    ws.Cells(ligne, 2).Font.Italic = True
    ligne = ligne + 1
    ws.Cells(ligne, 2).Value = "= Elles n'impactent pas la tresorerie et sont reintegrees dans la CAF"
    ws.Cells(ligne, 2).Font.Italic = True
    ligne = ligne + 1
    ws.Cells(ligne, 2).Value = "Dotations a reintegrer pour CAF"
    ws.Cells(ligne, 3).Value = totalDotN
    ws.Cells(ligne, 4).Value = totalDotN1
    ws.Cells(ligne, 2).Font.Bold = True
    ws.Columns("C:E").NumberFormat = "#,##0;[Red]-#,##0"
End Sub

Sub CreerOngletFinancier()
    Dim ws As Worksheet
    Dim ligne As Long

    Set ws = CreerOnglet("Financier", COULEUR_PL)
    Call EcrireEnTete(ws, "RESULTAT FINANCIER", "Comptes 66x et 76x")
    ligne = 4

    ' Produits financiers
    ws.Cells(ligne, 2).Value = "PRODUITS FINANCIERS (76x)"
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 7)).Interior.Color = RGB(198, 224, 180)
    ws.Cells(ligne, 2).Font.Bold = True
    ligne = ligne + 1

    Dim comp76() As String
    comp76 = Split("761,762,763,764,765,766,767,768,786,796", ",")
    Dim lab76() As String
    lab76 = Split("Produits participations,Produits autres immob.fin,Revenus créances immo.,Revenus VMP,Escomptes obtenus,Gains de change,Produits nets cession VMP,Autres prod. financiers,Reprises prov. financières,Transferts charges fin.", ",")

    Dim totalProdfN As Double, totalProdfN1 As Double
    Dim j As Integer, sN As Double, sN1 As Double

    For j = 0 To UBound(comp76)
        sN = -GetSoldeCompteN(comp76(j))
        sN1 = -GetSoldeCompteN1(comp76(j))
        If sN <> 0 Or sN1 <> 0 Then
            Call EcrireLigneCompte(ws, ligne, comp76(j), lab76(j), sN, sN1)
            totalProdfN = totalProdfN + sN
            totalProdfN1 = totalProdfN1 + sN1
            ligne = ligne + 1
        End If
    Next j
    Call EcrireSousTotal(ws, ligne, "Total produits financiers", totalProdfN, totalProdfN1)
    ligne = ligne + 2

    ' Charges financieres
    ws.Cells(ligne, 2).Value = "CHARGES FINANCIERES (66x)"
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 7)).Interior.Color = RGB(255, 230, 153)
    ws.Cells(ligne, 2).Font.Bold = True
    ligne = ligne + 1

    Dim comp66() As String
    comp66 = Split("661,662,663,664,665,666,667,668,686,696", ",")
    Dim lab66() As String
    lab66 = Split("Charges intérêts emprunts,Charges intérêts CBC,Pertes sur créances liées,Charges nettes/cession VMP,Escomptes accordés,Pertes de change,Charges nettes/cession IMF,Autres charges financières,Dot. prov. financières,Dot. prov. pour R&C financ.", ",")

    Dim totalChargefN As Double, totalChargefN1 As Double
    For j = 0 To UBound(comp66)
        sN = GetSoldeCompteN(comp66(j))
        sN1 = GetSoldeCompteN1(comp66(j))
        If sN <> 0 Or sN1 <> 0 Then
            Call EcrireLigneCompte(ws, ligne, comp66(j), lab66(j), sN, sN1)
            totalChargefN = totalChargefN + sN
            totalChargefN1 = totalChargefN1 + sN1
            ligne = ligne + 1
        End If
    Next j
    Call EcrireSousTotal(ws, ligne, "Total charges financieres", totalChargefN, totalChargefN1)
    ligne = ligne + 2

    Dim resFinN As Double, resFinN1 As Double
    resFinN = totalProdfN - totalChargefN
    resFinN1 = totalProdfN1 - totalChargefN1
    Call EcrireTotal(ws, ligne, "RESULTAT FINANCIER", resFinN, resFinN1)

    ' Ratios
    ligne = ligne + 3
    ws.Cells(ligne, 2).Value = "RATIOS FINANCIERS"
    ws.Cells(ligne, 2).Font.Bold = True
    ligne = ligne + 1

    Dim caN As Double, caN1 As Double
    caN = Abs(GetSoldeCompteN("70"))
    caN1 = Abs(GetSoldeCompteN1("70"))

    ws.Cells(ligne, 2).Value = "Charges financières / CA HT"
    If caN <> 0 Then ws.Cells(ligne, 3).Value = totalChargefN / caN
    If caN1 <> 0 Then ws.Cells(ligne, 4).Value = totalChargefN1 / caN1
    ws.Cells(ligne, 3).NumberFormat = "0.0%"
    ws.Cells(ligne, 4).NumberFormat = "0.0%"
    ligne = ligne + 1

    Dim excbrutexplN As Double, excbrutexplN1 As Double
    excbrutexplN = caN - GetSoldeCompteN("60") - GetSoldeCompteN("61") - GetSoldeCompteN("62") - GetSoldeCompteN("63") - GetSoldeCompteN("64") - GetSoldeCompteN("65")
    excbrutexplN1 = caN1 - GetSoldeCompteN1("60") - GetSoldeCompteN1("61") - GetSoldeCompteN1("62") - GetSoldeCompteN1("63") - GetSoldeCompteN1("64") - GetSoldeCompteN1("65")
    ws.Cells(ligne, 2).Value = "Charges financières / EBE (couverture)"
    If excbrutexplN <> 0 Then ws.Cells(ligne, 3).Value = totalChargefN / excbrutexplN
    If excbrutexplN1 <> 0 Then ws.Cells(ligne, 4).Value = totalChargefN1 / excbrutexplN1
    ws.Cells(ligne, 3).NumberFormat = "0.0%"
    ws.Cells(ligne, 4).NumberFormat = "0.0%"

    ws.Columns("C:E").NumberFormat = "#,##0;[Red]-#,##0"
    ws.Cells(ligne - 1, 3).NumberFormat = "0.0%"
    ws.Cells(ligne - 1, 4).NumberFormat = "0.0%"
    ws.Cells(ligne, 3).NumberFormat = "0.0%"
    ws.Cells(ligne, 4).NumberFormat = "0.0%"
End Sub

Sub CreerOngletExceptionnel()
    Dim ws As Worksheet
    Dim ligne As Long

    Set ws = CreerOnglet("Exceptionnel", COULEUR_PL)
    Call EcrireEnTete(ws, "RESULTAT EXCEPTIONNEL", "Comptes 67x et 77x")
    ligne = 4

    ' Produits exceptionnels
    ws.Cells(ligne, 2).Value = "PRODUITS EXCEPTIONNELS (77x)"
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 7)).Interior.Color = RGB(198, 224, 180)
    ws.Cells(ligne, 2).Font.Bold = True
    ligne = ligne + 1

    Dim comp77() As String
    comp77 = Split("771,772,775,777,778,787,797", ",")
    Dim lab77() As String
    lab77 = Split("Prod. except. sur opér. gestion,Prod. sur exercices antérieurs,Produits cession éléments actif,Quotes-parts subv. virées résultat,Autres prod. exceptionnels,Reprises prov. except.,Transferts charges except.", ",")

    Dim totalProdExcN As Double, totalProdExcN1 As Double
    Dim j As Integer, sN As Double, sN1 As Double

    For j = 0 To UBound(comp77)
        sN = -GetSoldeCompteN(comp77(j))
        sN1 = -GetSoldeCompteN1(comp77(j))
        If sN <> 0 Or sN1 <> 0 Then
            Call EcrireLigneCompte(ws, ligne, comp77(j), lab77(j), sN, sN1)
            totalProdExcN = totalProdExcN + sN
            totalProdExcN1 = totalProdExcN1 + sN1
            ligne = ligne + 1
        End If
    Next j

    ' VNCEAC (675)
    Dim vncaN As Double, vncaN1 As Double
    vncaN = GetSoldeCompteN("675")
    vncaN1 = GetSoldeCompteN1("675")

    Call EcrireSousTotal(ws, ligne, "Total produits exceptionnels", totalProdExcN, totalProdExcN1)
    ligne = ligne + 2

    ' Charges exceptionnelles
    ws.Cells(ligne, 2).Value = "CHARGES EXCEPTIONNELLES (67x)"
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 7)).Interior.Color = RGB(255, 230, 153)
    ws.Cells(ligne, 2).Font.Bold = True
    ligne = ligne + 1

    Dim comp67() As String
    comp67 = Split("671,672,675,678,687,697", ",")
    Dim lab67() As String
    lab67 = Split("Charges except. sur opér. gestion,Charges sur exercices antérieurs,VNC éléments d'actif cédés,Autres charges exceptionnelles,Dot. amort. prov. except.,Dot. prov. R&C except.", ",")

    Dim totalChargeExcN As Double, totalChargeExcN1 As Double
    For j = 0 To UBound(comp67)
        sN = GetSoldeCompteN(comp67(j))
        sN1 = GetSoldeCompteN1(comp67(j))
        If sN <> 0 Or sN1 <> 0 Then
            Call EcrireLigneCompte(ws, ligne, comp67(j), lab67(j), sN, sN1)
            totalChargeExcN = totalChargeExcN + sN
            totalChargeExcN1 = totalChargeExcN1 + sN1
            ligne = ligne + 1
        End If
    Next j
    Call EcrireSousTotal(ws, ligne, "Total charges exceptionnelles", totalChargeExcN, totalChargeExcN1)
    ligne = ligne + 2

    Dim resExcN As Double, resExcN1 As Double
    resExcN = totalProdExcN - totalChargeExcN
    resExcN1 = totalProdExcN1 - totalChargeExcN1
    Call EcrireTotal(ws, ligne, "RESULTAT EXCEPTIONNEL", resExcN, resExcN1)

    ' Plus values / Moins values cessions
    ligne = ligne + 3
    ws.Cells(ligne, 2).Value = "ANALYSE CESSIONS D'ACTIF"
    ws.Cells(ligne, 2).Font.Bold = True
    ligne = ligne + 1
    Dim pvN As Double, pvN1 As Double
    pvN = -GetSoldeCompteN("775") - GetSoldeCompteN("675")
    pvN1 = -GetSoldeCompteN1("775") - GetSoldeCompteN1("675")
    ws.Cells(ligne, 2).Value = "Plus-value / Moins-value de cession"
    ws.Cells(ligne, 3).Value = pvN
    ws.Cells(ligne, 4).Value = pvN1
    ws.Cells(ligne, 5).Value = pvN - pvN1
    ws.Cells(ligne, 3).Font.Bold = True
    ws.Columns("C:E").NumberFormat = "#,##0;[Red]-#,##0"
End Sub

Sub CreerOngletImpots()
    Dim ws As Worksheet
    Dim ligne As Long

    Set ws = CreerOnglet("Impots", COULEUR_PL)
    Call EcrireEnTete(ws, "IMPOTS SUR LES BENEFICES", "Comptes 69x")
    ligne = 4

    Dim comp69() As String
    comp69 = Split("691,695,696,697,699", ",")
    Dim lab69() As String
    lab69 = Split("Participation des salariés,IS courant,IS différé actif,IS différé passif,Produits IS (degrevement)", ",")

    Dim totalISN As Double, totalISN1 As Double
    Dim j As Integer, sN As Double, sN1 As Double

    For j = 0 To UBound(comp69)
        sN = GetSoldeCompteN(comp69(j))
        sN1 = GetSoldeCompteN1(comp69(j))
        If sN <> 0 Or sN1 <> 0 Then
            Call EcrireLigneCompte(ws, ligne, comp69(j), lab69(j), sN, sN1)
            totalISN = totalISN + sN
            totalISN1 = totalISN1 + sN1
            ligne = ligne + 1
        End If
    Next j

    ligne = ligne + 1
    Call EcrireTotal(ws, ligne, "TOTAL IS ET PARTICIPATION", totalISN, totalISN1)

    ' Taux effectif d'IS
    ligne = ligne + 3
    ws.Cells(ligne, 2).Value = "ANALYSE FISCALE"
    ws.Cells(ligne, 2).Font.Bold = True
    ligne = ligne + 1

    Dim resAvtISN As Double, resAvtISN1 As Double
    ' Resultat avant IS = total produits - total charges hors IS
    resAvtISN = -GetSoldeCompteN("70") - GetSoldeCompteN("71") - GetSoldeCompteN("72") - GetSoldeCompteN("74") - GetSoldeCompteN("75") - GetSoldeCompteN("76") - GetSoldeCompteN("77") - GetSoldeCompteN("78") - GetSoldeCompteN("79") + GetSoldeCompteN("60") + GetSoldeCompteN("61") + GetSoldeCompteN("62") + GetSoldeCompteN("63") + GetSoldeCompteN("64") + GetSoldeCompteN("65") + GetSoldeCompteN("66") + GetSoldeCompteN("67") + GetSoldeCompteN("68")
    resAvtISN1 = -GetSoldeCompteN1("70") - GetSoldeCompteN1("71") - GetSoldeCompteN1("72") - GetSoldeCompteN1("74") - GetSoldeCompteN1("75") - GetSoldeCompteN1("76") - GetSoldeCompteN1("77") - GetSoldeCompteN1("78") - GetSoldeCompteN1("79") + GetSoldeCompteN1("60") + GetSoldeCompteN1("61") + GetSoldeCompteN1("62") + GetSoldeCompteN1("63") + GetSoldeCompteN1("64") + GetSoldeCompteN1("65") + GetSoldeCompteN1("66") + GetSoldeCompteN1("67") + GetSoldeCompteN1("68")

    ws.Cells(ligne, 2).Value = "Résultat avant IS (estimé)"
    ws.Cells(ligne, 3).Value = resAvtISN
    ws.Cells(ligne, 4).Value = resAvtISN1
    ws.Cells(ligne, 3).NumberFormat = "#,##0;[Red]-#,##0"
    ws.Cells(ligne, 4).NumberFormat = "#,##0;[Red]-#,##0"
    ligne = ligne + 1

    ws.Cells(ligne, 2).Value = "IS comptabilisé"
    ws.Cells(ligne, 3).Value = GetSoldeCompteN("695")
    ws.Cells(ligne, 4).Value = GetSoldeCompteN1("695")
    ws.Cells(ligne, 3).NumberFormat = "#,##0;[Red]-#,##0"
    ws.Cells(ligne, 4).NumberFormat = "#,##0;[Red]-#,##0"
    ligne = ligne + 1

    ws.Cells(ligne, 2).Value = "Taux effectif d'IS"
    If resAvtISN <> 0 Then ws.Cells(ligne, 3).Value = GetSoldeCompteN("695") / resAvtISN
    If resAvtISN1 <> 0 Then ws.Cells(ligne, 4).Value = GetSoldeCompteN1("695") / resAvtISN1
    ws.Cells(ligne, 3).NumberFormat = "0.0%"
    ws.Cells(ligne, 4).NumberFormat = "0.0%"
    ws.Columns("C:E").NumberFormat = "#,##0;[Red]-#,##0"
    ws.Cells(ligne, 3).NumberFormat = "0.0%"
    ws.Cells(ligne, 4).NumberFormat = "0.0%"
End Sub
