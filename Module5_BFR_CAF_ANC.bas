Attribute VB_Name = "Module5_BFR_CAF_ANC"
Option Explicit

' ============================================================
' BFR ET CAF
' ============================================================
Sub CreerOngletBFR_CAF()
    Dim ws As Worksheet
    Dim ligne As Long

    Set ws = CreerOnglet("BFR-CAF", COULEUR_RATIO)

    With ws.Range("A1:J1")
        .Merge
        .Value = "BESOIN EN FONDS DE ROULEMENT (BFR) ET CAPACITE D'AUTOFINANCEMENT (CAF)"
        .Interior.Color = RGB(197, 90, 17)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
    End With

    ligne = 3

    ' ======================== FONDS DE ROULEMENT ========================
    ws.Cells(ligne, 1).Value = "I. FONDS DE ROULEMENT NET GLOBAL (FRNG)"
    With ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 6))
        .Interior.Color = RGB(197, 90, 17)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 12
    End With
    ligne = ligne + 1

    ' Ressources stables
    Dim cpN As Double, cpN1 As Double
    cpN = -(GetSoldeCompteN("101") + GetSoldeCompteN("104") + GetSoldeCompteN("105") + GetSoldeCompteN("106") + GetSoldeCompteN("11") + GetSoldeCompteN("12") + GetSoldeCompteN("13") + GetSoldeCompteN("14") + GetSoldeCompteN("15"))
    cpN1 = -(GetSoldeCompteN1("101") + GetSoldeCompteN1("104") + GetSoldeCompteN1("105") + GetSoldeCompteN1("106") + GetSoldeCompteN1("11") + GetSoldeCompteN1("12") + GetSoldeCompteN1("13") + GetSoldeCompteN1("14") + GetSoldeCompteN1("15"))

    Dim dettesLTN As Double, dettesLTN1 As Double
    dettesLTN = -(GetSoldeCompteN("16") - GetSoldeCompteN("519"))
    dettesLTN1 = -(GetSoldeCompteN1("16") - GetSoldeCompteN1("519"))

    Dim ressourcesStablesN As Double, ressourcesStablesN1 As Double
    ressourcesStablesN = cpN + dettesLTN
    ressourcesStablesN1 = cpN1 + dettesLTN1

    ' Emplois stables
    Dim actifImmN As Double, actifImmN1 As Double
    actifImmN = GetSoldeCompteN("20") + GetSoldeCompteN("21") + GetSoldeCompteN("22") + GetSoldeCompteN("23") + GetSoldeCompteN("26") + GetSoldeCompteN("27") + GetSoldeCompteN("28") + GetSoldeCompteN("29")
    actifImmN1 = GetSoldeCompteN1("20") + GetSoldeCompteN1("21") + GetSoldeCompteN1("22") + GetSoldeCompteN1("23") + GetSoldeCompteN1("26") + GetSoldeCompteN1("27") + GetSoldeCompteN1("28") + GetSoldeCompteN1("29")

    Dim frngN As Double, frngN1 As Double
    frngN = ressourcesStablesN - actifImmN
    frngN1 = ressourcesStablesN1 - actifImmN1

    ws.Cells(ligne, 1).Value = "Ressources stables (CP + Dettes LMT)"
    ws.Cells(ligne, 2).Value = ressourcesStablesN
    ws.Cells(ligne, 3).Value = ressourcesStablesN1
    ws.Cells(ligne, 4).Value = ressourcesStablesN - ressourcesStablesN1
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 4)).Interior.Color = RGB(198, 224, 180)
    ligne = ligne + 1

    ws.Cells(ligne, 1).Value = "(-) Actif immobilise net"
    ws.Cells(ligne, 2).Value = actifImmN
    ws.Cells(ligne, 3).Value = actifImmN1
    ws.Cells(ligne, 4).Value = actifImmN - actifImmN1
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 4)).Interior.Color = RGB(255, 199, 206)
    ligne = ligne + 1

    ws.Cells(ligne, 1).Value = "= FRNG"
    ws.Cells(ligne, 2).Value = frngN
    ws.Cells(ligne, 3).Value = frngN1
    ws.Cells(ligne, 4).Value = frngN - frngN1
    With ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 4))
        .Font.Bold = True
        .Interior.Color = RGB(253, 233, 217)
        .Borders(xlEdgeTop).LineStyle = xlDouble
        .Borders(xlEdgeBottom).LineStyle = xlDouble
    End With
    ligne = ligne + 3

    ' ======================== BFR ========================
    ws.Cells(ligne, 1).Value = "II. BESOIN EN FONDS DE ROULEMENT (BFR)"
    With ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 6))
        .Interior.Color = RGB(197, 90, 17)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 12
    End With
    ligne = ligne + 1

    ' Actif circulant hors tresorerie
    Dim stocksN As Double, stocksN1 As Double
    stocksN = GetSoldeCompteN("31") + GetSoldeCompteN("32") + GetSoldeCompteN("33") + GetSoldeCompteN("34") + GetSoldeCompteN("35") + GetSoldeCompteN("37") + GetSoldeCompteN("39")
    stocksN1 = GetSoldeCompteN1("31") + GetSoldeCompteN1("32") + GetSoldeCompteN1("33") + GetSoldeCompteN1("34") + GetSoldeCompteN1("35") + GetSoldeCompteN1("37") + GetSoldeCompteN1("39")

    Dim creancesCliN As Double, creancesCliN1 As Double
    creancesCliN = GetSoldeCompteN("41") + GetSoldeCompteN("490")
    creancesCliN1 = GetSoldeCompteN1("41") + GetSoldeCompteN1("490")

    Dim autresCreancesN As Double, autresCreancesN1 As Double
    autresCreancesN = GetSoldeCompteN("44") + GetSoldeCompteN("45") + GetSoldeCompteN("46") + GetSoldeCompteN("48")
    autresCreancesN1 = GetSoldeCompteN1("44") + GetSoldeCompteN1("45") + GetSoldeCompteN1("46") + GetSoldeCompteN1("48")

    Dim acircHorsTresN As Double, acircHorsTresN1 As Double
    acircHorsTresN = stocksN + creancesCliN + autresCreancesN
    acircHorsTresN1 = stocksN1 + creancesCliN1 + autresCreancesN1

    ' Passif circulant hors CBC
    Dim dettesFournN As Double, dettesFournN1 As Double
    dettesFournN = -(GetSoldeCompteN("40"))
    dettesFournN1 = -(GetSoldeCompteN1("40"))

    Dim dettesFiscSocN As Double, dettesFiscSocN1 As Double
    dettesFiscSocN = -(GetSoldeCompteN("42") + GetSoldeCompteN("43") + GetSoldeCompteN("44") + GetSoldeCompteN("45"))
    dettesFiscSocN1 = -(GetSoldeCompteN1("42") + GetSoldeCompteN1("43") + GetSoldeCompteN1("44") + GetSoldeCompteN1("45"))
    If dettesFiscSocN < 0 Then dettesFiscSocN = 0
    If dettesFiscSocN1 < 0 Then dettesFiscSocN1 = 0

    Dim autresDettesN As Double, autresDettesN1 As Double
    autresDettesN = -(GetSoldeCompteN("46") + GetSoldeCompteN("487"))
    autresDettesN1 = -(GetSoldeCompteN1("46") + GetSoldeCompteN1("487"))
    If autresDettesN < 0 Then autresDettesN = 0
    If autresDettesN1 < 0 Then autresDettesN1 = 0

    Dim pcircHorsCBCN As Double, pcircHorsCBCN1 As Double
    pcircHorsCBCN = dettesFournN + dettesFiscSocN + autresDettesN
    pcircHorsCBCN1 = dettesFournN1 + dettesFiscSocN1 + autresDettesN1

    Dim bfrN As Double, bfrN1 As Double
    bfrN = acircHorsTresN - pcircHorsCBCN
    bfrN1 = acircHorsTresN1 - pcircHorsCBCN1

    ' Detail BFR exploitation vs hors exploitation
    Dim bfrExplN As Double, bfrExplN1 As Double
    bfrExplN = stocksN + creancesCliN - dettesFournN - dettesFiscSocN
    bfrExplN1 = stocksN1 + creancesCliN1 - dettesFournN1 - dettesFiscSocN1

    Dim bfrHExplN As Double, bfrHExplN1 As Double
    bfrHExplN = autresCreancesN - autresDettesN
    bfrHExplN1 = autresCreancesN1 - autresDettesN1

    ws.Cells(ligne, 1).Value = "Stocks"
    ws.Cells(ligne, 2).Value = stocksN
    ws.Cells(ligne, 3).Value = stocksN1
    ws.Cells(ligne, 4).Value = stocksN - stocksN1
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 4)).Interior.Color = RGB(198, 224, 180)
    ligne = ligne + 1

    ws.Cells(ligne, 1).Value = "(+) Creances clients et rattachees"
    ws.Cells(ligne, 2).Value = creancesCliN
    ws.Cells(ligne, 3).Value = creancesCliN1
    ws.Cells(ligne, 4).Value = creancesCliN - creancesCliN1
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 4)).Interior.Color = RGB(198, 224, 180)
    ligne = ligne + 1

    ws.Cells(ligne, 1).Value = "(+) Autres creances d'exploitation"
    ws.Cells(ligne, 2).Value = autresCreancesN
    ws.Cells(ligne, 3).Value = autresCreancesN1
    ws.Cells(ligne, 4).Value = autresCreancesN - autresCreancesN1
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 4)).Interior.Color = RGB(198, 224, 180)
    ligne = ligne + 1

    ws.Cells(ligne, 1).Value = "(-) Dettes fournisseurs"
    ws.Cells(ligne, 2).Value = dettesFournN
    ws.Cells(ligne, 3).Value = dettesFournN1
    ws.Cells(ligne, 4).Value = dettesFournN - dettesFournN1
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 4)).Interior.Color = RGB(255, 199, 206)
    ligne = ligne + 1

    ws.Cells(ligne, 1).Value = "(-) Dettes fiscales et sociales"
    ws.Cells(ligne, 2).Value = dettesFiscSocN
    ws.Cells(ligne, 3).Value = dettesFiscSocN1
    ws.Cells(ligne, 4).Value = dettesFiscSocN - dettesFiscSocN1
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 4)).Interior.Color = RGB(255, 199, 206)
    ligne = ligne + 1

    ws.Cells(ligne, 1).Value = "(-) Autres dettes d'exploitation"
    ws.Cells(ligne, 2).Value = autresDettesN
    ws.Cells(ligne, 3).Value = autresDettesN1
    ws.Cells(ligne, 4).Value = autresDettesN - autresDettesN1
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 4)).Interior.Color = RGB(255, 199, 206)
    ligne = ligne + 1

    ws.Cells(ligne, 1).Value = "= BFR TOTAL"
    ws.Cells(ligne, 2).Value = bfrN
    ws.Cells(ligne, 3).Value = bfrN1
    ws.Cells(ligne, 4).Value = bfrN - bfrN1
    With ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 4))
        .Font.Bold = True
        .Interior.Color = RGB(253, 233, 217)
        .Borders(xlEdgeTop).LineStyle = xlDouble
        .Borders(xlEdgeBottom).LineStyle = xlDouble
    End With
    ligne = ligne + 2

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

    ' ======================== TRESORERIE NETTE ========================
    ws.Cells(ligne, 1).Value = "III. TRESORERIE NETTE"
    With ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 6))
        .Interior.Color = RGB(197, 90, 17)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 12
    End With
    ligne = ligne + 1

    Dim tresActN As Double, tresActN1 As Double
    Dim cbcN As Double, cbcN1 As Double
    tresActN = GetSoldeCompteN("50") + GetSoldeCompteN("51") + GetSoldeCompteN("53") + GetSoldeCompteN("54")
    tresActN1 = GetSoldeCompteN1("50") + GetSoldeCompteN1("51") + GetSoldeCompteN1("53") + GetSoldeCompteN1("54")
    cbcN = GetSoldeCompteN("519")
    cbcN1 = GetSoldeCompteN1("519")

    Dim tresNetteN As Double, tresNetteN1 As Double
    tresNetteN = tresActN + cbcN
    tresNetteN1 = tresActN1 + cbcN1

    ws.Cells(ligne, 1).Value = "Tresorerie active (disponibilites + VMP)"
    ws.Cells(ligne, 2).Value = tresActN
    ws.Cells(ligne, 3).Value = tresActN1
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 4)).Interior.Color = RGB(198, 224, 180)
    ligne = ligne + 1

    ws.Cells(ligne, 1).Value = "(-) Concours bancaires courants (519)"
    ws.Cells(ligne, 2).Value = cbcN
    ws.Cells(ligne, 3).Value = cbcN1
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 4)).Interior.Color = RGB(255, 199, 206)
    ligne = ligne + 1

    ws.Cells(ligne, 1).Value = "= TRESORERIE NETTE"
    ws.Cells(ligne, 2).Value = tresNetteN
    ws.Cells(ligne, 3).Value = tresNetteN1
    ws.Cells(ligne, 4).Value = tresNetteN - tresNetteN1
    With ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 4))
        .Font.Bold = True
        .Interior.Color = RGB(253, 233, 217)
        .Borders(xlEdgeTop).LineStyle = xlDouble
        .Borders(xlEdgeBottom).LineStyle = xlDouble
    End With
    ligne = ligne + 2

    ' Verification FRNG = BFR + TN
    ws.Cells(ligne, 1).Value = "VERIFICATION : FRNG = BFR + Tresorerie Nette"
    ws.Cells(ligne, 1).Font.Italic = True
    ws.Cells(ligne, 1).Font.Color = RGB(128, 128, 128)
    ligne = ligne + 1
    ws.Cells(ligne, 1).Value = "  FRNG"
    ws.Cells(ligne, 2).Value = frngN
    ws.Cells(ligne, 3).Value = frngN1
    ligne = ligne + 1
    ws.Cells(ligne, 1).Value = "  BFR + Tresorerie Nette"
    ws.Cells(ligne, 2).Value = bfrN + tresNetteN
    ws.Cells(ligne, 3).Value = bfrN1 + tresNetteN1
    ligne = ligne + 1
    ws.Cells(ligne, 1).Value = "  Ecart (doit etre nul)"
    ws.Cells(ligne, 2).Value = frngN - (bfrN + tresNetteN)
    ws.Cells(ligne, 3).Value = frngN1 - (bfrN1 + tresNetteN1)
    ws.Cells(ligne, 1).Font.Color = RGB(128, 128, 128)
    ligne = ligne + 4

    ' ======================== CAF ========================
    ws.Cells(ligne, 1).Value = "IV. CAPACITE D'AUTOFINANCEMENT (CAF)"
    With ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 6))
        .Interior.Color = RGB(197, 90, 17)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 12
    End With
    ligne = ligne + 1

    ' Methode additive (a partir du resultat net)
    ws.Cells(ligne, 1).Value = "METHODE ADDITIVE (a partir du résultat net)"
    ws.Cells(ligne, 1).Font.Bold = True
    ws.Cells(ligne, 1).Font.Underline = xlUnderlineStyleSingle
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
    pvCessN = -GetSoldeCompteN("775") - GetSoldeCompteN("675")  ' PV = produit cession - VNC
    pvCessN1 = -GetSoldeCompteN1("775") - GetSoldeCompteN1("675")

    Dim subvVireesN As Double, subvVireesN1 As Double
    subvVireesN = -GetSoldeCompteN("777")
    subvVireesN1 = -GetSoldeCompteN1("777")

    Dim cafN As Double, cafN1 As Double
    cafN = resNetN + dotTotN - reprTotN + pvCessN - subvVireesN
    cafN1 = resNetN1 + dotTotN1 - reprTotN1 + pvCessN1 - subvVireesN1

    Sub EcrireLigneCAF(ws As Worksheet, ligne As Long, libelle As String, valN As Double, valN1 As Double, signe As String)
        ws.Cells(ligne, 1).Value = signe & " " & libelle
        ws.Cells(ligne, 2).Value = valN
        ws.Cells(ligne, 3).Value = valN1
        ws.Cells(ligne, 4).Value = valN - valN1
        If valN > 0 Then ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 2)).Interior.Color = RGB(198, 224, 180)
        If valN < 0 Then ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 2)).Interior.Color = RGB(255, 199, 206)
    End Sub

    ws.Cells(ligne, 1).Value = "Résultat net de l'exercice"
    ws.Cells(ligne, 2).Value = resNetN
    ws.Cells(ligne, 3).Value = resNetN1
    ws.Cells(ligne, 4).Value = resNetN - resNetN1
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 4)).Interior.Color = RGB(198, 224, 180)
    ligne = ligne + 1

    ws.Cells(ligne, 1).Value = "(+) Dotations aux amortissements et provisions (charges calculées)"
    ws.Cells(ligne, 2).Value = dotTotN
    ws.Cells(ligne, 3).Value = dotTotN1
    ws.Cells(ligne, 4).Value = dotTotN - dotTotN1
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 4)).Interior.Color = RGB(198, 224, 180)
    ligne = ligne + 1

    ws.Cells(ligne, 1).Value = "(-) Reprises sur amortissements et provisions (produits calculés)"
    ws.Cells(ligne, 2).Value = reprTotN
    ws.Cells(ligne, 3).Value = reprTotN1
    ws.Cells(ligne, 4).Value = reprTotN - reprTotN1
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 4)).Interior.Color = RGB(255, 199, 206)
    ligne = ligne + 1

    ws.Cells(ligne, 1).Value = "(-) Plus-values de cession (produits exceptionnels non récurrents)"
    ws.Cells(ligne, 2).Value = pvCessN
    ws.Cells(ligne, 3).Value = pvCessN1
    ws.Cells(ligne, 4).Value = pvCessN - pvCessN1
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 4)).Interior.Color = RGB(255, 199, 206)
    ligne = ligne + 1

    ws.Cells(ligne, 1).Value = "(-) Subventions d'investissement virées au résultat"
    ws.Cells(ligne, 2).Value = subvVireesN
    ws.Cells(ligne, 3).Value = subvVireesN1
    ws.Cells(ligne, 4).Value = subvVireesN - subvVireesN1
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 4)).Interior.Color = RGB(255, 199, 206)
    ligne = ligne + 1

    ws.Cells(ligne, 1).Value = "= CAPACITE D'AUTOFINANCEMENT (CAF)"
    ws.Cells(ligne, 2).Value = cafN
    ws.Cells(ligne, 3).Value = cafN1
    ws.Cells(ligne, 4).Value = cafN - cafN1
    With ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 4))
        .Font.Bold = True
        .Font.Size = 12
        .Interior.Color = RGB(253, 233, 217)
        .Borders(xlEdgeTop).LineStyle = xlDouble
        .Borders(xlEdgeBottom).LineStyle = xlDouble
    End With
    ligne = ligne + 3

    ' Ratios BFR / CAF
    ws.Cells(ligne, 1).Value = "RATIOS D'ANALYSE"
    With ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 6))
        .Interior.Color = RGB(197, 90, 17)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    ligne = ligne + 1

    Dim caN As Double, caN1 As Double
    caN = Abs(-GetSoldeCompteN("70"))
    caN1 = Abs(-GetSoldeCompteN1("70"))

    Dim ratios() As String
    Dim vN() As Double, vN1() As Double
    ReDim ratios(7), vN(7), vN1(7)

    ratios(0) = "FRNG / CA (en jours)"
    ratios(1) = "BFR / CA (en jours)"
    ratios(2) = "BFR d'exploitation / CA (en jours)"
    ratios(3) = "Tresorerie Nette / CA (en jours)"
    ratios(4) = "CAF / Endettement net (capacite remboursement)"
    ratios(5) = "CAF / CA"
    ratios(6) = "Couverture BFR par le FRNG (FRNG / BFR)"
    ratios(7) = "Taux d'autofinancement (CAF / Investissements)"

    If caN <> 0 Then vN(0) = (frngN / caN) * 365
    If caN1 <> 0 Then vN1(0) = (frngN1 / caN1) * 365
    If caN <> 0 Then vN(1) = (bfrN / caN) * 365
    If caN1 <> 0 Then vN1(1) = (bfrN1 / caN1) * 365
    If caN <> 0 Then vN(2) = (bfrExplN / caN) * 365
    If caN1 <> 0 Then vN1(2) = (bfrExplN1 / caN1) * 365
    If caN <> 0 Then vN(3) = (tresNetteN / caN) * 365
    If caN1 <> 0 Then vN1(3) = (tresNetteN1 / caN1) * 365

    Dim endettNetN As Double, endettNetN1 As Double
    endettNetN = Abs(GetSoldeCompteN("16")) - tresActN
    endettNetN1 = Abs(GetSoldeCompteN1("16")) - tresActN1
    If cafN <> 0 Then vN(4) = endettNetN / cafN
    If cafN1 <> 0 Then vN1(4) = endettNetN1 / cafN1
    If caN <> 0 Then vN(5) = cafN / caN
    If caN1 <> 0 Then vN1(5) = cafN1 / caN1
    If bfrN <> 0 Then vN(6) = frngN / bfrN
    If bfrN1 <> 0 Then vN1(6) = frngN1 / bfrN1

    Dim investN As Double, investN1 As Double
    investN = Abs(GetSoldeCompteN("20") + GetSoldeCompteN("21") + GetSoldeCompteN("22") + GetSoldeCompteN("23")) - Abs(GetSoldeCompteN1("20") + GetSoldeCompteN1("21") + GetSoldeCompteN1("22") + GetSoldeCompteN1("23"))
    If investN > 0 And cafN <> 0 Then vN(7) = cafN / investN

    Dim j As Integer
    For j = 0 To 7
        ws.Cells(ligne, 1).Value = ratios(j)
        ws.Cells(ligne, 2).Value = vN(j)
        ws.Cells(ligne, 3).Value = vN1(j)
        ws.Cells(ligne, 4).Value = vN(j) - vN1(j)
        If j <= 3 Then
            ws.Cells(ligne, 2).NumberFormat = "0.0 ""j"""
            ws.Cells(ligne, 3).NumberFormat = "0.0 ""j"""
            ws.Cells(ligne, 4).NumberFormat = "0.0 ""j"""
        ElseIf j = 4 Then
            ws.Cells(ligne, 2).NumberFormat = "0.0x"
            ws.Cells(ligne, 3).NumberFormat = "0.0x"
        Else
            ws.Cells(ligne, 2).NumberFormat = "0.0%"
            ws.Cells(ligne, 3).NumberFormat = "0.0%"
            ws.Cells(ligne, 4).NumberFormat = "0.0%"
        End If
        If ligne Mod 2 = 0 Then ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 6)).Interior.Color = RGB(253, 233, 217)
        ligne = ligne + 1
    Next j

    ws.Columns("A").ColumnWidth = 60
    ws.Columns("B:D").ColumnWidth = 18
    ws.Range("B4:D100").NumberFormat = "#,##0;[Red]-#,##0"
    ws.Cells(3, 2).Value = "N"
    ws.Cells(3, 3).Value = "N-1"
    ws.Cells(3, 4).Value = "Variation"
    With ws.Range("A3:D3")
        .Font.Bold = True
        .Interior.Color = RGB(253, 233, 217)
    End With
End Sub

' ============================================================
' ANNEXE ANC - ETATS FINANCIERS NORMALISES
' ============================================================
Sub CreerOngletANC()
    Dim ws As Worksheet
    Dim ligne As Long

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

    ' I. Règles et méthodes comptables
    ws.Cells(ligne, 1).Value = "I. REGLES ET METHODES COMPTABLES"
    With ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 10))
        .Interior.Color = RGB(68, 0, 102)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 12
    End With
    ligne = ligne + 1

    Dim rulesMethodes() As String
    rulesMethodes = Split("Base de preparation : Les comptes ont été établis conformément aux dispositions du Plan Comptable Général (PCG) tel que modifié par le règlement ANC 2014-03.,Principes comptables : Continuité d'exploitation / Permanence des méthodes / Prudence / Indépendance des exercices / Non-compensation.,Immobilisations incorporelles : Amorties linéairement sur leur durée d'utilisation. Les fonds de commerce ne sont pas amortis sauf dépréciation.,Immobilisations corporelles : Amorties selon le mode linéaire ou dégressif fiscal selon la nature du bien.,Stocks : Évalués au coût de revient (PEPS ou CMP selon méthode retenue). Les dépréciations sont constituées sur les stocks obsolètes ou à rotation lente.,Créances : Évaluées à leur valeur nominale. Les créances douteuses font l'objet d'une dépréciation individuelle.,Instruments financiers : Les VMP sont évaluées à leur valeur de marché. Les moins-values latentes sont provisionnées.,Provisions pour risques : Constituées dès que l'obligation est certaine ou probable selon IAS 37 adapté au PCG.,Impôts différés : Non reconnus (option PCG de base). En cas d'adoption du règlement 2022-06 : voir note complémentaire.", ",")

    Dim r As Integer
    For r = 0 To UBound(rulesMethodes)
        ws.Cells(ligne, 1).Value = (r + 1) & ". " & rulesMethodes(r)
        ws.Cells(ligne, 1).WrapText = True
        ws.Rows(ligne).RowHeight = 30
        If ligne Mod 2 = 0 Then ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 10)).Interior.Color = RGB(237, 226, 244)
        ligne = ligne + 1
    Next r
    ligne = ligne + 2

    ' II. Immobilisations - Tableau des mouvements
    ws.Cells(ligne, 1).Value = "II. TABLEAU DES IMMOBILISATIONS (Art. L.123-22 C.com.)"
    With ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 10))
        .Interior.Color = RGB(68, 0, 102)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 12
    End With
    ligne = ligne + 1

    ' En-tetes tableau immobilisations
    Dim entetesImmo() As String
    entetesImmo = Split("Nature,Debut exercice,Acquisitions,Créations,Apports,Cessions/Retraits,Autres mouvements,Fin exercice", ",")
    For c = 0 To UBound(entetesImmo)
        ws.Cells(ligne, c + 1).Value = entetesImmo(c)
    Next c
    With ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 8))
        .Font.Bold = True
        .Interior.Color = RGB(209, 184, 232)
        .HorizontalAlignment = xlCenter
    End With
    ligne = ligne + 1

    ' Lignes immobilisations
    Dim typesImmo() As String
    Dim compDebutN() As String
    typesImmo = Split("Frais établissement (201),Frais R&D (203),Concessions brevets (205),Fonds commercial (207),Autres immo. incorp. (208),Terrains (211),Constructions (213),Install. tech. mat. outillage (215),Autres immo. corp. (218),Immo. en cours (23),Titres participation (261),Autres immo. financ. (27)", ",")
    compDebutN = Split("201,203,205,207,208,211,213,215,218,23,261,27", ",")

    Dim ci As Integer
    For ci = 0 To UBound(typesImmo)
        Dim debutN As Double, finN As Double
        debutN = GetSoldeCompteN1(compDebutN(ci))
        finN = GetSoldeCompteN(compDebutN(ci))
        If debutN <> 0 Or finN <> 0 Then
            ws.Cells(ligne, 1).Value = typesImmo(ci)
            ws.Cells(ligne, 2).Value = debutN
            ws.Cells(ligne, 8).Value = finN
            ws.Cells(ligne, 7).Value = finN - debutN  ' Simplification : mouvement net
            If ligne Mod 2 = 0 Then ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 8)).Interior.Color = RGB(242, 242, 242)
            ligne = ligne + 1
        End If
    Next ci

    ' Total immobilisations brutes
    Dim totImmoDebutN As Double, totImmoFinN As Double
    totImmoDebutN = GetSoldeCompteN1("20") + GetSoldeCompteN1("21") + GetSoldeCompteN1("22") + GetSoldeCompteN1("23") + GetSoldeCompteN1("26") + GetSoldeCompteN1("27")
    totImmoFinN = GetSoldeCompteN("20") + GetSoldeCompteN("21") + GetSoldeCompteN("22") + GetSoldeCompteN("23") + GetSoldeCompteN("26") + GetSoldeCompteN("27")
    ws.Cells(ligne, 1).Value = "TOTAL IMMOBILISATIONS BRUTES"
    ws.Cells(ligne, 2).Value = totImmoDebutN
    ws.Cells(ligne, 8).Value = totImmoFinN
    ws.Cells(ligne, 7).Value = totImmoFinN - totImmoDebutN
    With ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 8))
        .Font.Bold = True
        .Interior.Color = RGB(209, 184, 232)
    End With
    ligne = ligne + 3

    ' III. Tableau des amortissements
    ws.Cells(ligne, 1).Value = "III. TABLEAU DES AMORTISSEMENTS ET DEPRECIATIONS"
    With ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 10))
        .Interior.Color = RGB(68, 0, 102)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 12
    End With
    ligne = ligne + 1

    Dim entetesAmort() As String
    entetesAmort = Split("Nature,Debut exercice,Dotations exercice,Reprises exercice,Fin exercice", ",")
    For c = 0 To UBound(entetesAmort)
        ws.Cells(ligne, c + 1).Value = entetesAmort(c)
    Next c
    With ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 5))
        .Font.Bold = True
        .Interior.Color = RGB(209, 184, 232)
    End With
    ligne = ligne + 1

    Dim amortCompN1 As Double, amortCompN As Double, dotAmortN As Double
    amortCompN1 = -(GetSoldeCompteN1("28") + GetSoldeCompteN1("29"))
    amortCompN = -(GetSoldeCompteN("28") + GetSoldeCompteN("29"))
    dotAmortN = GetSoldeCompteN("681")
    ws.Cells(ligne, 1).Value = "Amortissements immobilisations corporelles et incorporelles"
    ws.Cells(ligne, 2).Value = amortCompN1
    ws.Cells(ligne, 3).Value = dotAmortN
    ws.Cells(ligne, 5).Value = amortCompN
    ws.Cells(ligne, 4).Value = amortCompN - amortCompN1 - dotAmortN  ' Reprises (négatif = correction)
    ligne = ligne + 2

    ' IV. Provisions pour risques et charges
    ws.Cells(ligne, 1).Value = "IV. TABLEAU DES PROVISIONS POUR RISQUES ET CHARGES"
    With ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 10))
        .Interior.Color = RGB(68, 0, 102)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 12
    End With
    ligne = ligne + 1

    Dim entetesProvRisques() As String
    entetesProvRisques = Split("Nature,Debut N,Dotations,Reprises utilisées,Reprises non util.,Fin N", ",")
    For c = 0 To UBound(entetesProvRisques)
        ws.Cells(ligne, c + 1).Value = entetesProvRisques(c)
    Next c
    With ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 6))
        .Font.Bold = True
        .Interior.Color = RGB(209, 184, 232)
    End With
    ligne = ligne + 1

    Dim provCodes() As String
    Dim provLabels() As String
    provCodes = Split("151,152,153,154,155,156,157,158", ",")
    provLabels = Split("Provisions pour litiges,Prov. garanties données,Prov. pertes marchés,Prov. amendes pénalités,Prov. pertes de change,Prov. pensions retraites,Prov. pour impôts,Autres prov. R&C", ",")

    For ci = 0 To UBound(provCodes)
        Dim pDebN As Double, pFinN As Double
        pDebN = -GetSoldeCompteN1(provCodes(ci))
        pFinN = -GetSoldeCompteN(provCodes(ci))
        If pDebN <> 0 Or pFinN <> 0 Then
            ws.Cells(ligne, 1).Value = provLabels(ci)
            ws.Cells(ligne, 2).Value = pDebN
            ws.Cells(ligne, 3).Value = IIf(pFinN > pDebN, pFinN - pDebN, 0)
            ws.Cells(ligne, 4) = IIf(pFinN < pDebN, pDebN - pFinN, 0)
            ws.Cells(ligne, 6).Value = pFinN
            If ligne Mod 2 = 0 Then ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 6)).Interior.Color = RGB(242, 242, 242)
            ligne = ligne + 1
        End If
    Next ci

    Dim totProvDebN As Double, totProvFinN As Double
    totProvDebN = -GetSoldeCompteN1("15")
    totProvFinN = -GetSoldeCompteN("15")
    ws.Cells(ligne, 1).Value = "TOTAL PROVISIONS R&C"
    ws.Cells(ligne, 2).Value = totProvDebN
    ws.Cells(ligne, 3).Value = IIf(totProvFinN > totProvDebN, totProvFinN - totProvDebN, 0)
    ws.Cells(ligne, 4).Value = IIf(totProvFinN < totProvDebN, totProvDebN - totProvFinN, 0)
    ws.Cells(ligne, 6).Value = totProvFinN
    With ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 6))
        .Font.Bold = True
        .Interior.Color = RGB(209, 184, 232)
    End With
    ligne = ligne + 3

    ' V. Etat des créances et dettes
    ws.Cells(ligne, 1).Value = "V. ETAT DES CREANCES ET DES DETTES (Art. R.123-198 C.com.)"
    With ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 10))
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

    ' Creances immobilisees > 1 an
    Dim creImmoN As Double
    creImmoN = GetSoldeCompteN("267") + GetSoldeCompteN("274") + GetSoldeCompteN("275")
    If creImmoN <> 0 Then
        ws.Cells(ligne, 1).Value = "Créances de l'actif immobilisé"
        ws.Cells(ligne, 2).Value = creImmoN
        ws.Cells(ligne, 3).Value = 0
        ws.Cells(ligne, 4).Value = creImmoN
        ligne = ligne + 1
    End If

    ' Creances clients
    Dim creCliN As Double
    creCliN = GetSoldeCompteN("41") + GetSoldeCompteN("490")
    ws.Cells(ligne, 1).Value = "Créances clients"
    ws.Cells(ligne, 2).Value = creCliN
    ws.Cells(ligne, 3).Value = creCliN
    ws.Cells(ligne, 4).Value = 0
    ligne = ligne + 1

    ' Autres creances
    Dim autCreN As Double
    autCreN = GetSoldeCompteN("44") + GetSoldeCompteN("45") + GetSoldeCompteN("46")
    ws.Cells(ligne, 1).Value = "Autres créances (fisc., soc., divers)"
    ws.Cells(ligne, 2).Value = autCreN
    ws.Cells(ligne, 3).Value = autCreN
    ws.Cells(ligne, 4).Value = 0
    ligne = ligne + 1

    ws.Cells(ligne, 1).Value = "TOTAL CREANCES"
    ws.Cells(ligne, 2).Value = creImmoN + creCliN + autCreN
    ws.Cells(ligne, 3).Value = creCliN + autCreN
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

    ' Dettes financieres (estimation: 1/3 CT, 2/3 MLT simplifiée)
    Dim detteFinN As Double
    detteFinN = -GetSoldeCompteN("16")
    If detteFinN <> 0 Then
        ws.Cells(ligne, 1).Value = "Emprunts et dettes auprès établissements de crédit"
        ws.Cells(ligne, 2).Value = detteFinN
        ws.Cells(ligne, 3).Value = detteFinN * 0.2  ' Hypothèse 20% CT
        ws.Cells(ligne, 4).Value = detteFinN * 0.5  ' Hypothèse 50% 1-5 ans
        ws.Cells(ligne, 5).Value = detteFinN * 0.3  ' Hypothèse 30% > 5 ans
        ws.Cells(ligne, 6).Value = "A ventiler selon tableau d'amortissement"
        ws.Cells(ligne, 6).Font.Italic = True
        ws.Cells(ligne, 6).Font.Color = RGB(128, 128, 128)
        ligne = ligne + 1
    End If

    Dim detteFournN As Double
    detteFournN = -GetSoldeCompteN("40")
    ws.Cells(ligne, 1).Value = "Dettes fournisseurs et comptes rattachés"
    ws.Cells(ligne, 2).Value = detteFournN
    ws.Cells(ligne, 3).Value = detteFournN
    ligne = ligne + 1

    Dim detteFiscSocN As Double
    detteFiscSocN = Abs(GetSoldeCompteN("42")) + Abs(GetSoldeCompteN("43")) + Abs(GetSoldeCompteN("44"))
    ws.Cells(ligne, 1).Value = "Dettes fiscales et sociales"
    ws.Cells(ligne, 2).Value = detteFiscSocN
    ws.Cells(ligne, 3).Value = detteFiscSocN
    ligne = ligne + 1

    ws.Cells(ligne, 1).Value = "TOTAL DETTES"
    ws.Cells(ligne, 2).Value = detteFinN + detteFournN + detteFiscSocN
    ws.Cells(ligne, 3).Value = detteFinN * 0.2 + detteFournN + detteFiscSocN
    ws.Cells(ligne, 4).Value = detteFinN * 0.5
    ws.Cells(ligne, 5).Value = detteFinN * 0.3
    With ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 5))
        .Font.Bold = True
        .Interior.Color = RGB(209, 184, 232)
    End With
    ligne = ligne + 3

    ' VI. Effectifs
    ws.Cells(ligne, 1).Value = "VI. INFORMATIONS SUR LE PERSONNEL (Art. L.2323-46 C.trav.)"
    With ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 10))
        .Interior.Color = RGB(68, 0, 102)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 12
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
    ws.Cells(ligne, 1).Value = "Participation des salariés (691)"
    ws.Cells(ligne, 2).Value = GetSoldeCompteN("691")
    ws.Cells(ligne, 3).Value = GetSoldeCompteN1("691")
    ligne = ligne + 1
    ws.Cells(ligne, 1).Value = "Note : L'effectif moyen est à renseigner manuellement"
    ws.Cells(ligne, 1).Font.Italic = True
    ws.Cells(ligne, 1).Font.Color = RGB(128, 128, 128)

    ' Formatage
    ws.Columns("A").ColumnWidth = 60
    ws.Columns("B:J").ColumnWidth = 16
    ws.Range("B6:J300").NumberFormat = "#,##0;[Red]-#,##0"
    ws.Rows("1:2").RowHeight = 30

    ' Note de bas de page
    ligne = ligne + 3
    ws.Cells(ligne, 1).Value = "REFERENCE REGLEMENTAIRE : Règlement ANC 2014-03 du 05/06/2014 relatif au Plan Comptable Général"
    ws.Cells(ligne, 1).Font.Bold = True
    ws.Cells(ligne, 1).Font.Italic = True
    ws.Cells(ligne, 1).Font.Color = RGB(68, 0, 102)
    ligne = ligne + 1
    ws.Cells(ligne, 1).Value = "Les états financiers ci-dessus ont été établis selon les principes et méthodes comptables définis par l'ANC."
    ws.Cells(ligne, 1).Font.Italic = True
    ws.Cells(ligne, 1).Font.Color = RGB(128, 128, 128)
End Sub
