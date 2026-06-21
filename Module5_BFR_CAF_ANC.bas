Attribute VB_Name = "Module5_BFR_CAF_ANC"
Option Explicit

' ============================================================
' BFR ET CAF
' ============================================================
Sub CreerOngletBFR_CAF()
    Dim ws As Worksheet
    Dim ligne As Long

    Set ws = CreerOnglet("BFR-CAF", COULEUR_RATIO)

    With ws.Range("A1:F1")
        .Merge
        .Value = "BESOIN EN FONDS DE ROULEMENT (BFR) ET CAPACITE D'AUTOFINANCEMENT (CAF)"
        .Interior.Color = RGB(197, 90, 17)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
    End With

    ws.Cells(2, 1).Value = "Indicateur"
    ws.Cells(2, 2).Value = "N"
    ws.Cells(2, 3).Value = "N-1"
    ws.Cells(2, 4).Value = "Variation"
    With ws.Range("A2:D2")
        .Font.Bold = True
        .Interior.Color = RGB(253, 233, 217)
    End With

    ligne = 3

    ' ======================== FRNG ========================
    ws.Cells(ligne, 1).Value = "I. FONDS DE ROULEMENT NET GLOBAL (FRNG)"
    With ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 4))
        .Interior.Color = RGB(197, 90, 17)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    ligne = ligne + 1

    Dim cpN As Double, cpN1 As Double
    cpN = -(GetSoldeCompteN("101") + GetSoldeCompteN("104") + GetSoldeCompteN("105") + GetSoldeCompteN("106") + GetSoldeCompteN("11") + GetSoldeCompteN("12") + GetSoldeCompteN("13") + GetSoldeCompteN("14") + GetSoldeCompteN("15"))
    cpN1 = -(GetSoldeCompteN1("101") + GetSoldeCompteN1("104") + GetSoldeCompteN1("105") + GetSoldeCompteN1("106") + GetSoldeCompteN1("11") + GetSoldeCompteN1("12") + GetSoldeCompteN1("13") + GetSoldeCompteN1("14") + GetSoldeCompteN1("15"))

    Dim dettesLTN As Double, dettesLTN1 As Double
    dettesLTN = -GetSoldeCompteN("16")
    dettesLTN1 = -GetSoldeCompteN1("16")

    Dim ressStabN As Double, ressStabN1 As Double
    ressStabN = cpN + dettesLTN
    ressStabN1 = cpN1 + dettesLTN1

    Dim actifImmN As Double, actifImmN1 As Double
    actifImmN = GetSoldeCompteN("20") + GetSoldeCompteN("21") + GetSoldeCompteN("22") + GetSoldeCompteN("23") + GetSoldeCompteN("26") + GetSoldeCompteN("27") + GetSoldeCompteN("28") + GetSoldeCompteN("29")
    actifImmN1 = GetSoldeCompteN1("20") + GetSoldeCompteN1("21") + GetSoldeCompteN1("22") + GetSoldeCompteN1("23") + GetSoldeCompteN1("26") + GetSoldeCompteN1("27") + GetSoldeCompteN1("28") + GetSoldeCompteN1("29")

    Dim frngN As Double, frngN1 As Double
    frngN = ressStabN - actifImmN
    frngN1 = ressStabN1 - actifImmN1

    Call EcrireLigneBFR(ws, ligne, "Ressources stables (CP + Dettes LMT)", ressStabN, ressStabN1, RGB(198, 224, 180)) : ligne = ligne + 1
    Call EcrireLigneBFR(ws, ligne, "(-) Actif immobilise net", actifImmN, actifImmN1, RGB(255, 199, 206)) : ligne = ligne + 1
    Call EcrireTotalBFR(ws, ligne, "= FRNG", frngN, frngN1) : ligne = ligne + 3

    ' ======================== BFR ========================
    ws.Cells(ligne, 1).Value = "II. BESOIN EN FONDS DE ROULEMENT (BFR)"
    With ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 4))
        .Interior.Color = RGB(197, 90, 17)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    ligne = ligne + 1

    Dim stocksN As Double, stocksN1 As Double
    stocksN = GetSoldeCompteN("31") + GetSoldeCompteN("32") + GetSoldeCompteN("33") + GetSoldeCompteN("34") + GetSoldeCompteN("35") + GetSoldeCompteN("37") + GetSoldeCompteN("39")
    stocksN1 = GetSoldeCompteN1("31") + GetSoldeCompteN1("32") + GetSoldeCompteN1("33") + GetSoldeCompteN1("34") + GetSoldeCompteN1("35") + GetSoldeCompteN1("37") + GetSoldeCompteN1("39")

    Dim creCliN As Double, creCliN1 As Double
    creCliN = GetSoldeCompteN("41") + GetSoldeCompteN("490")
    creCliN1 = GetSoldeCompteN1("41") + GetSoldeCompteN1("490")

    Dim autCreN As Double, autCreN1 As Double
    autCreN = GetSoldeCompteN("44") + GetSoldeCompteN("45") + GetSoldeCompteN("46") + GetSoldeCompteN("48")
    autCreN1 = GetSoldeCompteN1("44") + GetSoldeCompteN1("45") + GetSoldeCompteN1("46") + GetSoldeCompteN1("48")

    Dim detFournN As Double, detFournN1 As Double
    detFournN = -GetSoldeCompteN("40")
    detFournN1 = -GetSoldeCompteN1("40")
    If detFournN < 0 Then detFournN = 0
    If detFournN1 < 0 Then detFournN1 = 0

    Dim detFSN As Double, detFSN1 As Double
    detFSN = Abs(GetSoldeCompteN("42")) + Abs(GetSoldeCompteN("43")) + Abs(GetSoldeCompteN("44")) + Abs(GetSoldeCompteN("45"))
    detFSN1 = Abs(GetSoldeCompteN1("42")) + Abs(GetSoldeCompteN1("43")) + Abs(GetSoldeCompteN1("44")) + Abs(GetSoldeCompteN1("45"))

    Dim autDetN As Double, autDetN1 As Double
    autDetN = Abs(GetSoldeCompteN("46")) + Abs(GetSoldeCompteN("487"))
    autDetN1 = Abs(GetSoldeCompteN1("46")) + Abs(GetSoldeCompteN1("487"))

    Dim bfrExplN As Double, bfrExplN1 As Double
    bfrExplN = stocksN + creCliN - detFournN - detFSN
    bfrExplN1 = stocksN1 + creCliN1 - detFournN1 - detFSN1

    Dim bfrHExplN As Double, bfrHExplN1 As Double
    bfrHExplN = autCreN - autDetN
    bfrHExplN1 = autCreN1 - autDetN1

    Dim bfrN As Double, bfrN1 As Double
    bfrN = bfrExplN + bfrHExplN
    bfrN1 = bfrExplN1 + bfrHExplN1

    Call EcrireLigneBFR(ws, ligne, "Stocks", stocksN, stocksN1, RGB(198, 224, 180)) : ligne = ligne + 1
    Call EcrireLigneBFR(ws, ligne, "(+) Creances clients", creCliN, creCliN1, RGB(198, 224, 180)) : ligne = ligne + 1
    Call EcrireLigneBFR(ws, ligne, "(+) Autres creances d'exploitation", autCreN, autCreN1, RGB(198, 224, 180)) : ligne = ligne + 1
    Call EcrireLigneBFR(ws, ligne, "(-) Dettes fournisseurs", detFournN, detFournN1, RGB(255, 199, 206)) : ligne = ligne + 1
    Call EcrireLigneBFR(ws, ligne, "(-) Dettes fiscales et sociales", detFSN, detFSN1, RGB(255, 199, 206)) : ligne = ligne + 1
    Call EcrireLigneBFR(ws, ligne, "(-) Autres dettes d'exploitation", autDetN, autDetN1, RGB(255, 199, 206)) : ligne = ligne + 1
    Call EcrireTotalBFR(ws, ligne, "= BFR TOTAL", bfrN, bfrN1) : ligne = ligne + 1

    ws.Cells(ligne, 1).Value = "  dont BFR Exploitation"
    ws.Cells(ligne, 2).Value = bfrExplN
    ws.Cells(ligne, 3).Value = bfrExplN1
    ws.Cells(ligne, 1).Font.Italic = True
    ligne = ligne + 1
    ws.Cells(ligne, 1).Value = "  dont BFR Hors Exploitation"
    ws.Cells(ligne, 2).Value = bfrHExplN
    ws.Cells(ligne, 3).Value = bfrHExplN1
    ws.Cells(ligne, 1).Font.Italic = True
    ligne = ligne + 3

    ' ======================== TRESORERIE ========================
    ws.Cells(ligne, 1).Value = "III. TRESORERIE NETTE"
    With ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 4))
        .Interior.Color = RGB(197, 90, 17)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    ligne = ligne + 1

    Dim tresActN As Double, tresActN1 As Double
    tresActN = GetSoldeCompteN("50") + GetSoldeCompteN("51") + GetSoldeCompteN("53") + GetSoldeCompteN("54")
    tresActN1 = GetSoldeCompteN1("50") + GetSoldeCompteN1("51") + GetSoldeCompteN1("53") + GetSoldeCompteN1("54")

    Dim cbcN As Double, cbcN1 As Double
    cbcN = GetSoldeCompteN("519")
    cbcN1 = GetSoldeCompteN1("519")

    Dim tresNetteN As Double, tresNetteN1 As Double
    tresNetteN = tresActN + cbcN
    tresNetteN1 = tresActN1 + cbcN1

    Call EcrireLigneBFR(ws, ligne, "Tresorerie active (dispo + VMP)", tresActN, tresActN1, RGB(198, 224, 180)) : ligne = ligne + 1
    Call EcrireLigneBFR(ws, ligne, "(-) Concours bancaires courants (519)", cbcN, cbcN1, RGB(255, 199, 206)) : ligne = ligne + 1
    Call EcrireTotalBFR(ws, ligne, "= TRESORERIE NETTE", tresNetteN, tresNetteN1) : ligne = ligne + 3

    ' Verification FRNG = BFR + TN
    ws.Cells(ligne, 1).Value = "VERIFICATION : FRNG = BFR + Tresorerie Nette"
    ws.Cells(ligne, 1).Font.Italic = True
    ws.Cells(ligne, 1).Font.Color = RGB(128, 128, 128)
    ligne = ligne + 1
    ws.Cells(ligne, 1).Value = "Ecart (doit etre proche de zero)"
    ws.Cells(ligne, 2).Value = frngN - (bfrN + tresNetteN)
    ws.Cells(ligne, 3).Value = frngN1 - (bfrN1 + tresNetteN1)
    ws.Cells(ligne, 1).Font.Color = RGB(128, 128, 128)
    ligne = ligne + 4

    ' ======================== CAF ========================
    ws.Cells(ligne, 1).Value = "IV. CAPACITE D'AUTOFINANCEMENT (CAF) - Methode additive"
    With ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 4))
        .Interior.Color = RGB(197, 90, 17)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    ligne = ligne + 1

    Dim resNetN As Double, resNetN1 As Double
    resNetN = -GetSoldeCompteN("12")
    resNetN1 = -GetSoldeCompteN1("12")

    Dim dotTotN As Double, dotTotN1 As Double
    dotTotN = GetSoldeCompteN("681") + GetSoldeCompteN("686") + GetSoldeCompteN("687") + GetSoldeCompteN("696") + GetSoldeCompteN("697")
    dotTotN1 = GetSoldeCompteN1("681") + GetSoldeCompteN1("686") + GetSoldeCompteN1("687") + GetSoldeCompteN1("696") + GetSoldeCompteN1("697")

    Dim reprTotN As Double, reprTotN1 As Double
    reprTotN = -GetSoldeCompteN("781") - GetSoldeCompteN("786") - GetSoldeCompteN("787") - GetSoldeCompteN("791") - GetSoldeCompteN("796") - GetSoldeCompteN("797")
    reprTotN1 = -GetSoldeCompteN1("781") - GetSoldeCompteN1("786") - GetSoldeCompteN1("787") - GetSoldeCompteN1("791") - GetSoldeCompteN1("796") - GetSoldeCompteN1("797")

    Dim pvCessN As Double, pvCessN1 As Double
    pvCessN = -GetSoldeCompteN("775") - GetSoldeCompteN("675")
    pvCessN1 = -GetSoldeCompteN1("775") - GetSoldeCompteN1("675")

    Dim subvVireesN As Double, subvVireesN1 As Double
    subvVireesN = -GetSoldeCompteN("777")
    subvVireesN1 = -GetSoldeCompteN1("777")

    Dim cafN As Double, cafN1 As Double
    cafN = resNetN + dotTotN - reprTotN + pvCessN - subvVireesN
    cafN1 = resNetN1 + dotTotN1 - reprTotN1 + pvCessN1 - subvVireesN1

    Call EcrireLigneBFR(ws, ligne, "Resultat net de l'exercice", resNetN, resNetN1, RGB(198, 224, 180)) : ligne = ligne + 1
    Call EcrireLigneBFR(ws, ligne, "(+) Dotations aux amort. et provisions", dotTotN, dotTotN1, RGB(198, 224, 180)) : ligne = ligne + 1
    Call EcrireLigneBFR(ws, ligne, "(-) Reprises sur amort. et provisions", reprTotN, reprTotN1, RGB(255, 199, 206)) : ligne = ligne + 1
    Call EcrireLigneBFR(ws, ligne, "(-) Plus-values de cession nettes", pvCessN, pvCessN1, RGB(255, 199, 206)) : ligne = ligne + 1
    Call EcrireLigneBFR(ws, ligne, "(-) Subventions investissement virees", subvVireesN, subvVireesN1, RGB(255, 199, 206)) : ligne = ligne + 1
    Call EcrireTotalBFR(ws, ligne, "= CAPACITE D'AUTOFINANCEMENT (CAF)", cafN, cafN1) : ligne = ligne + 3

    ' Ratios
    ws.Cells(ligne, 1).Value = "RATIOS"
    With ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 4))
        .Interior.Color = RGB(197, 90, 17)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    ligne = ligne + 1

    Dim caN As Double, caN1 As Double
    caN = Abs(-GetSoldeCompteN("70"))
    caN1 = Abs(-GetSoldeCompteN1("70"))

    Dim endettNetN As Double, endettNetN1 As Double
    endettNetN = Abs(GetSoldeCompteN("16")) - tresActN
    endettNetN1 = Abs(GetSoldeCompteN1("16")) - tresActN1

    Dim investN As Double
    Dim immoN As Double, immoN1 As Double
    immoN = GetSoldeCompteN("20") + GetSoldeCompteN("21") + GetSoldeCompteN("22") + GetSoldeCompteN("23")
    immoN1 = GetSoldeCompteN1("20") + GetSoldeCompteN1("21") + GetSoldeCompteN1("22") + GetSoldeCompteN1("23")
    investN = immoN - immoN1

    Call EcrireRatioBFR(ws, ligne, "FRNG en jours de CA", IIf(caN <> 0, (frngN / caN) * 365, 0), IIf(caN1 <> 0, (frngN1 / caN1) * 365, 0), "j") : ligne = ligne + 1
    Call EcrireRatioBFR(ws, ligne, "BFR en jours de CA", IIf(caN <> 0, (bfrN / caN) * 365, 0), IIf(caN1 <> 0, (bfrN1 / caN1) * 365, 0), "j") : ligne = ligne + 1
    Call EcrireRatioBFR(ws, ligne, "BFR exploitation en jours de CA", IIf(caN <> 0, (bfrExplN / caN) * 365, 0), IIf(caN1 <> 0, (bfrExplN1 / caN1) * 365, 0), "j") : ligne = ligne + 1
    Call EcrireRatioBFR(ws, ligne, "Tresorerie nette en jours de CA", IIf(caN <> 0, (tresNetteN / caN) * 365, 0), IIf(caN1 <> 0, (tresNetteN1 / caN1) * 365, 0), "j") : ligne = ligne + 1
    Call EcrireRatioBFR(ws, ligne, "CAF / CA", IIf(caN <> 0, cafN / caN, 0), IIf(caN1 <> 0, cafN1 / caN1, 0), "%") : ligne = ligne + 1
    Call EcrireRatioBFR(ws, ligne, "Couverture BFR par le FRNG (FRNG/BFR)", IIf(bfrN <> 0, frngN / bfrN, 0), IIf(bfrN1 <> 0, frngN1 / bfrN1, 0), "x") : ligne = ligne + 1
    Call EcrireRatioBFR(ws, ligne, "Capacite remboursement (Endett.net/CAF)", IIf(cafN <> 0, endettNetN / cafN, 0), IIf(cafN1 <> 0, endettNetN1 / cafN1, 0), "x") : ligne = ligne + 1
    Call EcrireRatioBFR(ws, ligne, "Taux autofinancement (CAF/Investissements)", IIf(investN > 0, cafN / investN, 0), 0, "%") : ligne = ligne + 1

    ws.Columns("A").ColumnWidth = 55
    ws.Columns("B:D").ColumnWidth = 18
    ws.Range("B3:D200").NumberFormat = "#,##0;[Red]-#,##0"
End Sub

Sub EcrireLigneBFR(ws As Worksheet, ligne As Long, Libelle As String, ValN As Double, ValN1 As Double, Couleur As Long)
    ws.Cells(ligne, 1).Value = Libelle
    ws.Cells(ligne, 2).Value = ValN
    ws.Cells(ligne, 3).Value = ValN1
    ws.Cells(ligne, 4).Value = ValN - ValN1
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 4)).Interior.Color = Couleur
End Sub

Sub EcrireTotalBFR(ws As Worksheet, ligne As Long, Libelle As String, ValN As Double, ValN1 As Double)
    ws.Cells(ligne, 1).Value = Libelle
    ws.Cells(ligne, 2).Value = ValN
    ws.Cells(ligne, 3).Value = ValN1
    ws.Cells(ligne, 4).Value = ValN - ValN1
    With ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 4))
        .Font.Bold = True
        .Interior.Color = RGB(253, 233, 217)
        .Borders(xlEdgeTop).LineStyle = xlDouble
        .Borders(xlEdgeBottom).LineStyle = xlDouble
    End With
End Sub

Sub EcrireRatioBFR(ws As Worksheet, ligne As Long, Libelle As String, ValN As Double, ValN1 As Double, Unite As String)
    ws.Cells(ligne, 1).Value = Libelle
    ws.Cells(ligne, 2).Value = ValN
    ws.Cells(ligne, 3).Value = ValN1
    ws.Cells(ligne, 4).Value = ValN - ValN1
    Dim fmt As String
    Select Case Unite
        Case "j" : fmt = "0.0"
        Case "x" : fmt = "0.0"
        Case "%" : fmt = "0.0%"
        Case Else : fmt = "0.0"
    End Select
    ws.Cells(ligne, 2).NumberFormat = fmt
    ws.Cells(ligne, 3).NumberFormat = fmt
    ws.Cells(ligne, 4).NumberFormat = fmt
    If ligne Mod 2 = 0 Then ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 4)).Interior.Color = RGB(253, 233, 217)
End Sub

' ============================================================
' ANNEXE ANC
' ============================================================
Sub CreerOngletANC()
    Dim ws As Worksheet
    Dim ligne As Long
    Dim c As Integer
    Dim ci As Integer

    Set ws = CreerOnglet("Annexe ANC", COULEUR_ANC)

    With ws.Range("A1:J1")
        .Merge
        .Value = "ETATS FINANCIERS SELON LES REGLES ANC (PCG - Reglement 2014-03)"
        .Interior.Color = RGB(68, 0, 102)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
    End With
    With ws.Range("A2:J2")
        .Merge
        .Value = "Annexe aux comptes annuels - Informations selon recommandations ANC"
        .Interior.Color = RGB(112, 48, 160)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With

    ligne = 4

    ' I. Regles et methodes comptables
    ws.Cells(ligne, 1).Value = "I. REGLES ET METHODES COMPTABLES"
    With ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 10))
        .Interior.Color = RGB(68, 0, 102)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 12
    End With
    ligne = ligne + 1

    ws.Cells(ligne, 1).Value = "1. Base : Comptes etablis selon le PCG (reglement ANC 2014-03)"
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 10)).Interior.Color = RGB(237, 226, 244)
    ws.Cells(ligne, 1).WrapText = True
    ligne = ligne + 1
    ws.Cells(ligne, 1).Value = "2. Principes : Continuite exploitation / Permanence methodes / Prudence / Independence exercices"
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 10)).Interior.Color = RGB(255, 255, 255)
    ligne = ligne + 1
    ws.Cells(ligne, 1).Value = "3. Immobilisations incorporelles : Amortissement lineaire sur duree d'utilite. Fonds de commerce non amorti sauf depreciation."
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 10)).Interior.Color = RGB(237, 226, 244)
    ligne = ligne + 1
    ws.Cells(ligne, 1).Value = "4. Immobilisations corporelles : Mode lineaire ou degressif fiscal selon nature du bien."
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 10)).Interior.Color = RGB(255, 255, 255)
    ligne = ligne + 1
    ws.Cells(ligne, 1).Value = "5. Stocks : Evalues au cout de revient (PEPS ou CMP). Depreciations sur stocks obsoletes ou a rotation lente."
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 10)).Interior.Color = RGB(237, 226, 244)
    ligne = ligne + 1
    ws.Cells(ligne, 1).Value = "6. Creances : Valeur nominale. Creances douteuses : depreciation individuelle."
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 10)).Interior.Color = RGB(255, 255, 255)
    ligne = ligne + 1
    ws.Cells(ligne, 1).Value = "7. Provisions pour risques : Constituees des que l'obligation est certaine ou probable."
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 10)).Interior.Color = RGB(237, 226, 244)
    ligne = ligne + 2

    ' II. Tableau des immobilisations
    ws.Cells(ligne, 1).Value = "II. TABLEAU DES IMMOBILISATIONS"
    With ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 8))
        .Interior.Color = RGB(68, 0, 102)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 12
    End With
    ligne = ligne + 1

    ws.Cells(ligne, 1).Value = "Nature"
    ws.Cells(ligne, 2).Value = "Debut N"
    ws.Cells(ligne, 3).Value = "Acquisitions"
    ws.Cells(ligne, 4).Value = "Cessions/Retraits"
    ws.Cells(ligne, 5).Value = "Autres mvts"
    ws.Cells(ligne, 6).Value = "Fin N"
    With ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 6))
        .Font.Bold = True
        .Interior.Color = RGB(209, 184, 232)
        .HorizontalAlignment = xlCenter
    End With
    ligne = ligne + 1

    Dim typesImmo(5) As String
    Dim compImmo(5) As String
    typesImmo(0) = "Immobilisations incorporelles (20)"
    typesImmo(1) = "Terrains (211)"
    typesImmo(2) = "Constructions (213)"
    typesImmo(3) = "Installations techniques (215)"
    typesImmo(4) = "Autres immo. corporelles (218)"
    typesImmo(5) = "Immobilisations financieres (26-27)"
    compImmo(0) = "20"
    compImmo(1) = "211"
    compImmo(2) = "213"
    compImmo(3) = "215"
    compImmo(4) = "218"
    compImmo(5) = "26"

    For ci = 0 To 5
        Dim debN2 As Double, finN2 As Double
        debN2 = GetSoldeCompteN1(compImmo(ci))
        finN2 = GetSoldeCompteN(compImmo(ci))
        If debN2 <> 0 Or finN2 <> 0 Then
            ws.Cells(ligne, 1).Value = typesImmo(ci)
            ws.Cells(ligne, 2).Value = debN2
            ws.Cells(ligne, 6).Value = finN2
            ws.Cells(ligne, 5).Value = finN2 - debN2
            If ligne Mod 2 = 0 Then ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 6)).Interior.Color = RGB(242, 242, 242)
            ligne = ligne + 1
        End If
    Next ci

    Dim totImmoN As Double, totImmoN1 As Double
    totImmoN = GetSoldeCompteN("20") + GetSoldeCompteN("21") + GetSoldeCompteN("22") + GetSoldeCompteN("23") + GetSoldeCompteN("26") + GetSoldeCompteN("27")
    totImmoN1 = GetSoldeCompteN1("20") + GetSoldeCompteN1("21") + GetSoldeCompteN1("22") + GetSoldeCompteN1("23") + GetSoldeCompteN1("26") + GetSoldeCompteN1("27")
    ws.Cells(ligne, 1).Value = "TOTAL IMMOBILISATIONS BRUTES"
    ws.Cells(ligne, 2).Value = totImmoN1
    ws.Cells(ligne, 6).Value = totImmoN
    ws.Cells(ligne, 5).Value = totImmoN - totImmoN1
    With ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 6))
        .Font.Bold = True
        .Interior.Color = RGB(209, 184, 232)
    End With
    ligne = ligne + 3

    ' III. Tableau des amortissements
    ws.Cells(ligne, 1).Value = "III. TABLEAU DES AMORTISSEMENTS ET DEPRECIATIONS"
    With ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 6))
        .Interior.Color = RGB(68, 0, 102)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 12
    End With
    ligne = ligne + 1

    ws.Cells(ligne, 1).Value = "Nature"
    ws.Cells(ligne, 2).Value = "Debut N"
    ws.Cells(ligne, 3).Value = "Dotations"
    ws.Cells(ligne, 4).Value = "Reprises"
    ws.Cells(ligne, 5).Value = "Fin N"
    With ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 5))
        .Font.Bold = True
        .Interior.Color = RGB(209, 184, 232)
    End With
    ligne = ligne + 1

    Dim amortDebN As Double, amortFinN As Double
    amortDebN = -(GetSoldeCompteN1("28") + GetSoldeCompteN1("29"))
    amortFinN = -(GetSoldeCompteN("28") + GetSoldeCompteN("29"))
    Dim dotAmortN As Double
    dotAmortN = GetSoldeCompteN("681")
    ws.Cells(ligne, 1).Value = "Amortissements des immobilisations"
    ws.Cells(ligne, 2).Value = amortDebN
    ws.Cells(ligne, 3).Value = dotAmortN
    ws.Cells(ligne, 4).Value = amortFinN - amortDebN - dotAmortN
    ws.Cells(ligne, 5).Value = amortFinN
    ligne = ligne + 1

    Dim depCliDebN As Double, depCliFinN As Double
    depCliDebN = -GetSoldeCompteN1("490")
    depCliFinN = -GetSoldeCompteN("490")
    ws.Cells(ligne, 1).Value = "Depreciations clients (490)"
    ws.Cells(ligne, 2).Value = depCliDebN
    ws.Cells(ligne, 5).Value = depCliFinN
    ws.Cells(ligne, 3).Value = IIf(depCliFinN > depCliDebN, depCliFinN - depCliDebN, 0)
    ws.Cells(ligne, 4).Value = IIf(depCliFinN < depCliDebN, depCliDebN - depCliFinN, 0)
    ligne = ligne + 3

    ' IV. Provisions pour risques et charges
    ws.Cells(ligne, 1).Value = "IV. TABLEAU DES PROVISIONS POUR RISQUES ET CHARGES"
    With ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 6))
        .Interior.Color = RGB(68, 0, 102)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 12
    End With
    ligne = ligne + 1

    ws.Cells(ligne, 1).Value = "Nature"
    ws.Cells(ligne, 2).Value = "Debut N"
    ws.Cells(ligne, 3).Value = "Dotations"
    ws.Cells(ligne, 4).Value = "Reprises utilisees"
    ws.Cells(ligne, 5).Value = "Reprises non util."
    ws.Cells(ligne, 6).Value = "Fin N"
    With ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 6))
        .Font.Bold = True
        .Interior.Color = RGB(209, 184, 232)
    End With
    ligne = ligne + 1

    Dim provCodes(7) As String
    Dim provLabels(7) As String
    provCodes(0) = "151" : provLabels(0) = "Provisions pour litiges"
    provCodes(1) = "152" : provLabels(1) = "Prov. garanties donnees"
    provCodes(2) = "153" : provLabels(2) = "Prov. pertes sur marches"
    provCodes(3) = "154" : provLabels(3) = "Prov. amendes penalites"
    provCodes(4) = "155" : provLabels(4) = "Prov. pertes de change"
    provCodes(5) = "156" : provLabels(5) = "Prov. pensions retraites"
    provCodes(6) = "157" : provLabels(6) = "Prov. pour impots"
    provCodes(7) = "158" : provLabels(7) = "Autres prov. R&C"

    For ci = 0 To 7
        Dim pDebN As Double, pFinN As Double
        pDebN = -GetSoldeCompteN1(provCodes(ci))
        pFinN = -GetSoldeCompteN(provCodes(ci))
        If pDebN <> 0 Or pFinN <> 0 Then
            ws.Cells(ligne, 1).Value = provLabels(ci)
            ws.Cells(ligne, 2).Value = pDebN
            ws.Cells(ligne, 3).Value = IIf(pFinN > pDebN, pFinN - pDebN, 0)
            ws.Cells(ligne, 4).Value = IIf(pFinN < pDebN, pDebN - pFinN, 0)
            ws.Cells(ligne, 6).Value = pFinN
            If ligne Mod 2 = 0 Then ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 6)).Interior.Color = RGB(242, 242, 242)
            ligne = ligne + 1
        End If
    Next ci

    Dim totProvN As Double, totProvN1b As Double
    totProvN = -GetSoldeCompteN("15")
    totProvN1b = -GetSoldeCompteN1("15")
    ws.Cells(ligne, 1).Value = "TOTAL PROVISIONS R&C"
    ws.Cells(ligne, 2).Value = totProvN1b
    ws.Cells(ligne, 3).Value = IIf(totProvN > totProvN1b, totProvN - totProvN1b, 0)
    ws.Cells(ligne, 4).Value = IIf(totProvN < totProvN1b, totProvN1b - totProvN, 0)
    ws.Cells(ligne, 6).Value = totProvN
    With ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 6))
        .Font.Bold = True
        .Interior.Color = RGB(209, 184, 232)
    End With
    ligne = ligne + 3

    ' V. Etat des creances et dettes
    ws.Cells(ligne, 1).Value = "V. ETAT DES CREANCES ET DES DETTES"
    With ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 5))
        .Interior.Color = RGB(68, 0, 102)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 12
    End With
    ligne = ligne + 1

    ws.Cells(ligne, 1).Value = "CREANCES"
    ws.Cells(ligne, 2).Value = "Montant brut"
    ws.Cells(ligne, 3).Value = "A 1 an au plus"
    ws.Cells(ligne, 4).Value = "A plus d'1 an"
    With ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 4))
        .Font.Bold = True
        .Interior.Color = RGB(209, 184, 232)
    End With
    ligne = ligne + 1

    Dim creImmoN As Double
    creImmoN = GetSoldeCompteN("267") + GetSoldeCompteN("274") + GetSoldeCompteN("275")
    Dim creCliN2 As Double
    creCliN2 = GetSoldeCompteN("41") + GetSoldeCompteN("490")
    Dim autCreN2 As Double
    autCreN2 = GetSoldeCompteN("44") + GetSoldeCompteN("45") + GetSoldeCompteN("46")

    If creImmoN <> 0 Then
        ws.Cells(ligne, 1).Value = "Creances de l'actif immobilise (>1 an)"
        ws.Cells(ligne, 2).Value = creImmoN
        ws.Cells(ligne, 3).Value = 0
        ws.Cells(ligne, 4).Value = creImmoN
        ligne = ligne + 1
    End If
    ws.Cells(ligne, 1).Value = "Creances clients (<1 an)"
    ws.Cells(ligne, 2).Value = creCliN2
    ws.Cells(ligne, 3).Value = creCliN2
    ws.Cells(ligne, 4).Value = 0
    ligne = ligne + 1
    ws.Cells(ligne, 1).Value = "Autres creances (fiscal, social, divers)"
    ws.Cells(ligne, 2).Value = autCreN2
    ws.Cells(ligne, 3).Value = autCreN2
    ws.Cells(ligne, 4).Value = 0
    ligne = ligne + 1
    ws.Cells(ligne, 1).Value = "TOTAL CREANCES"
    ws.Cells(ligne, 2).Value = creImmoN + creCliN2 + autCreN2
    ws.Cells(ligne, 3).Value = creCliN2 + autCreN2
    ws.Cells(ligne, 4).Value = creImmoN
    With ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 4))
        .Font.Bold = True
        .Interior.Color = RGB(209, 184, 232)
    End With
    ligne = ligne + 2

    ws.Cells(ligne, 1).Value = "DETTES"
    ws.Cells(ligne, 2).Value = "Montant total"
    ws.Cells(ligne, 3).Value = "A 1 an au plus"
    ws.Cells(ligne, 4).Value = "De 1 a 5 ans"
    ws.Cells(ligne, 5).Value = "A plus de 5 ans"
    With ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 5))
        .Font.Bold = True
        .Interior.Color = RGB(209, 184, 232)
    End With
    ligne = ligne + 1

    Dim detteFinN2 As Double
    detteFinN2 = -GetSoldeCompteN("16")
    If detteFinN2 <> 0 Then
        ws.Cells(ligne, 1).Value = "Emprunts et dettes aupres etablissements de credit"
        ws.Cells(ligne, 2).Value = detteFinN2
        ws.Cells(ligne, 3).Value = detteFinN2 * 0.2
        ws.Cells(ligne, 4).Value = detteFinN2 * 0.5
        ws.Cells(ligne, 5).Value = detteFinN2 * 0.3
        ws.Cells(ligne, 6).Value = "Ventiler selon tableau d'amortissement"
        ws.Cells(ligne, 6).Font.Italic = True
        ws.Cells(ligne, 6).Font.Color = RGB(128, 128, 128)
        ligne = ligne + 1
    End If

    Dim detteFournN2 As Double
    detteFournN2 = -GetSoldeCompteN("40")
    ws.Cells(ligne, 1).Value = "Dettes fournisseurs et comptes rattaches"
    ws.Cells(ligne, 2).Value = detteFournN2
    ws.Cells(ligne, 3).Value = detteFournN2
    ligne = ligne + 1

    Dim detteFSN2 As Double
    detteFSN2 = Abs(GetSoldeCompteN("42")) + Abs(GetSoldeCompteN("43")) + Abs(GetSoldeCompteN("44"))
    ws.Cells(ligne, 1).Value = "Dettes fiscales et sociales"
    ws.Cells(ligne, 2).Value = detteFSN2
    ws.Cells(ligne, 3).Value = detteFSN2
    ligne = ligne + 1

    ws.Cells(ligne, 1).Value = "TOTAL DETTES"
    ws.Cells(ligne, 2).Value = detteFinN2 + detteFournN2 + detteFSN2
    ws.Cells(ligne, 3).Value = detteFinN2 * 0.2 + detteFournN2 + detteFSN2
    ws.Cells(ligne, 4).Value = detteFinN2 * 0.5
    ws.Cells(ligne, 5).Value = detteFinN2 * 0.3
    With ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 5))
        .Font.Bold = True
        .Interior.Color = RGB(209, 184, 232)
    End With
    ligne = ligne + 3

    ' VI. Informations personnel
    ws.Cells(ligne, 1).Value = "VI. INFORMATIONS SUR LE PERSONNEL"
    With ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 4))
        .Interior.Color = RGB(68, 0, 102)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 12
    End With
    ligne = ligne + 1

    ws.Cells(ligne, 1).Value = "Indicateur"
    ws.Cells(ligne, 2).Value = "N"
    ws.Cells(ligne, 3).Value = "N-1"
    With ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 3))
        .Font.Bold = True
        .Interior.Color = RGB(209, 184, 232)
    End With
    ligne = ligne + 1

    ws.Cells(ligne, 1).Value = "Masse salariale brute (641)"
    ws.Cells(ligne, 2).Value = GetSoldeCompteN("641")
    ws.Cells(ligne, 3).Value = GetSoldeCompteN1("641")
    ligne = ligne + 1
    ws.Cells(ligne, 1).Value = "Charges sociales employeur (645+646)"
    ws.Cells(ligne, 2).Value = GetSoldeCompteN("645") + GetSoldeCompteN("646")
    ws.Cells(ligne, 3).Value = GetSoldeCompteN1("645") + GetSoldeCompteN1("646")
    ligne = ligne + 1
    ws.Cells(ligne, 1).Value = "Participation des salaries (691)"
    ws.Cells(ligne, 2).Value = GetSoldeCompteN("691")
    ws.Cells(ligne, 3).Value = GetSoldeCompteN1("691")
    ligne = ligne + 1
    ws.Cells(ligne, 1).Value = "Effectif moyen (a renseigner manuellement)"
    ws.Cells(ligne, 2).Value = ""
    ws.Cells(ligne, 3).Value = ""
    ws.Cells(ligne, 1).Font.Italic = True
    ligne = ligne + 2

    ws.Cells(ligne, 1).Value = "Reference : Reglement ANC 2014-03 du 05/06/2014 - Plan Comptable General"
    ws.Cells(ligne, 1).Font.Bold = True
    ws.Cells(ligne, 1).Font.Color = RGB(68, 0, 102)

    ws.Columns("A").ColumnWidth = 60
    ws.Columns("B:J").ColumnWidth = 16
    ws.Range("B6:J500").NumberFormat = "#,##0;[Red]-#,##0"
End Sub
