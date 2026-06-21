Attribute VB_Name = "Module6_PPT"
Option Explicit

' ============================================================
' SYNTHESE POUR PRESENTATION PPT
' ============================================================
Sub CreerOngletPPT()
    Dim ws As Worksheet
    Dim ligne As Long

    Set ws = CreerOnglet("Synthese PPT", COULEUR_PPT)

    ' Style général : fond blanc, mise en page propre pour capture d'écran / impression
    ws.SetBackgroundPicture ""  ' Pas de fond

    ' === PAGE 1 : TITRE ===
    With ws.Range("A1:L2")
        .Merge
        .Value = "ANALYSE FINANCIERE"
        .Font.Bold = True
        .Font.Size = 28
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 73, 125)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    ws.Rows("1:2").RowHeight = 45

    With ws.Range("A3:L3")
        .Merge
        .Value = "Exercices N et N-1 | Analyse comparée | Confidentiel"
        .Font.Size = 12
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(68, 114, 196)
        .HorizontalAlignment = xlCenter
    End With
    ws.Rows(3).RowHeight = 25
    ligne = 5

    ' === KPI DASHBOARD (Ligne de KPIs) ===
    Call EcrireEntitePPT(ws, ligne, "KPI DASHBOARD - SYNTHESE EXECUTIVEE")
    ligne = ligne + 1

    ' Calcul des KPIs
    Dim caN As Double, caN1 As Double
    caN = Abs(-GetSoldeCompteN("70"))
    caN1 = Abs(-GetSoldeCompteN1("70"))

    Dim ebeN As Double, ebeN1 As Double
    Dim personnelN As Double, personnelN1 As Double
    Dim impotsTaxesN As Double, impotsTaxesN1 As Double
    Dim subvExplN As Double, subvExplN1 As Double
    personnelN = GetSoldeCompteN("64")
    personnelN1 = GetSoldeCompteN1("64")
    impotsTaxesN = GetSoldeCompteN("63")
    impotsTaxesN1 = GetSoldeCompteN1("63")
    subvExplN = -GetSoldeCompteN("74")
    subvExplN1 = -GetSoldeCompteN1("74")

    Dim vaH As Double, vaH1 As Double
    vaH = caN + (-GetSoldeCompteN("71")) + (-GetSoldeCompteN("72")) - GetSoldeCompteN("60") - GetSoldeCompteN("61") - GetSoldeCompteN("62")
    vaH1 = caN1 + (-GetSoldeCompteN1("71")) + (-GetSoldeCompteN1("72")) - GetSoldeCompteN1("60") - GetSoldeCompteN1("61") - GetSoldeCompteN1("62")

    ebeN = vaH + subvExplN - impotsTaxesN - personnelN
    ebeN1 = vaH1 + subvExplN1 - impotsTaxesN1 - personnelN1

    Dim dotN As Double, dotN1 As Double
    dotN = GetSoldeCompteN("681")
    dotN1 = GetSoldeCompteN1("681")

    Dim chargesFinN As Double, chargesFinN1 As Double
    chargesFinN = GetSoldeCompteN("66") + GetSoldeCompteN("686") + GetSoldeCompteN("696")
    chargesFinN1 = GetSoldeCompteN1("66") + GetSoldeCompteN1("686") + GetSoldeCompteN1("696")

    Dim prodFinN As Double, prodFinN1 As Double
    prodFinN = -GetSoldeCompteN("76") - GetSoldeCompteN("786") - GetSoldeCompteN("796")
    prodFinN1 = -GetSoldeCompteN1("76") - GetSoldeCompteN1("786") - GetSoldeCompteN1("796")

    Dim rexN As Double, rexN1 As Double
    rexN = ebeN - dotN - GetSoldeCompteN("65") + (-GetSoldeCompteN("75")) - (-GetSoldeCompteN("781")) - (-GetSoldeCompteN("791"))
    rexN1 = ebeN1 - dotN1 - GetSoldeCompteN1("65") + (-GetSoldeCompteN1("75")) - (-GetSoldeCompteN1("781")) - (-GetSoldeCompteN1("791"))

    Dim rcaN As Double, rcaN1 As Double
    rcaN = rexN + prodFinN - chargesFinN
    rcaN1 = rexN1 + prodFinN1 - chargesFinN1

    Dim resExcN As Double, resExcN1 As Double
    resExcN = (-GetSoldeCompteN("77") - GetSoldeCompteN("787") - GetSoldeCompteN("797")) - (GetSoldeCompteN("67") + GetSoldeCompteN("687") + GetSoldeCompteN("697"))
    resExcN1 = (-GetSoldeCompteN1("77") - GetSoldeCompteN1("787") - GetSoldeCompteN1("797")) - (GetSoldeCompteN1("67") + GetSoldeCompteN1("687") + GetSoldeCompteN1("697"))

    Dim isN As Double, isN1 As Double
    isN = GetSoldeCompteN("695") + GetSoldeCompteN("691")
    isN1 = GetSoldeCompteN1("695") + GetSoldeCompteN1("691")

    Dim resNetN As Double, resNetN1 As Double
    resNetN = rcaN + resExcN - isN
    resNetN1 = rcaN1 + resExcN1 - isN1

    ' Bilan
    Dim actifImmN As Double, actifImmN1 As Double
    actifImmN = GetSoldeCompteN("20") + GetSoldeCompteN("21") + GetSoldeCompteN("22") + GetSoldeCompteN("23") + GetSoldeCompteN("26") + GetSoldeCompteN("27") + GetSoldeCompteN("28") + GetSoldeCompteN("29")
    actifImmN1 = GetSoldeCompteN1("20") + GetSoldeCompteN1("21") + GetSoldeCompteN1("22") + GetSoldeCompteN1("23") + GetSoldeCompteN1("26") + GetSoldeCompteN1("27") + GetSoldeCompteN1("28") + GetSoldeCompteN1("29")

    Dim cpN As Double, cpN1 As Double
    cpN = -(GetSoldeCompteN("101") + GetSoldeCompteN("104") + GetSoldeCompteN("105") + GetSoldeCompteN("106") + GetSoldeCompteN("11") + GetSoldeCompteN("12") + GetSoldeCompteN("13"))
    cpN1 = -(GetSoldeCompteN1("101") + GetSoldeCompteN1("104") + GetSoldeCompteN1("105") + GetSoldeCompteN1("106") + GetSoldeCompteN1("11") + GetSoldeCompteN1("12") + GetSoldeCompteN1("13"))

    Dim dettesFinN As Double, dettesFinN1 As Double
    dettesFinN = -GetSoldeCompteN("16")
    dettesFinN1 = -GetSoldeCompteN1("16")

    Dim tresNetteN As Double, tresNetteN1 As Double
    tresNetteN = GetSoldeCompteN("50") + GetSoldeCompteN("51") + GetSoldeCompteN("53") + GetSoldeCompteN("54") + GetSoldeCompteN("519")
    tresNetteN1 = GetSoldeCompteN1("50") + GetSoldeCompteN1("51") + GetSoldeCompteN1("53") + GetSoldeCompteN1("54") + GetSoldeCompteN1("519")

    Dim bfrN As Double, bfrN1 As Double
    Dim stocksN As Double, stocksN1 As Double, creCliN As Double, creCliN1 As Double
    Dim detFournN As Double, detFournN1 As Double, detFSN As Double, detFSN1 As Double
    stocksN = GetSoldeCompteN("31") + GetSoldeCompteN("32") + GetSoldeCompteN("33") + GetSoldeCompteN("34") + GetSoldeCompteN("35") + GetSoldeCompteN("37") + GetSoldeCompteN("39")
    stocksN1 = GetSoldeCompteN1("31") + GetSoldeCompteN1("32") + GetSoldeCompteN1("33") + GetSoldeCompteN1("34") + GetSoldeCompteN1("35") + GetSoldeCompteN1("37") + GetSoldeCompteN1("39")
    creCliN = GetSoldeCompteN("41") + GetSoldeCompteN("490")
    creCliN1 = GetSoldeCompteN1("41") + GetSoldeCompteN1("490")
    detFournN = -GetSoldeCompteN("40")
    detFournN1 = -GetSoldeCompteN1("40")
    detFSN = Abs(GetSoldeCompteN("42")) + Abs(GetSoldeCompteN("43")) + Abs(GetSoldeCompteN("44"))
    detFSN1 = Abs(GetSoldeCompteN1("42")) + Abs(GetSoldeCompteN1("43")) + Abs(GetSoldeCompteN1("44"))
    bfrN = stocksN + creCliN - detFournN - detFSN
    bfrN1 = stocksN1 + creCliN1 - detFournN1 - detFSN1

    Dim cafN As Double, cafN1 As Double
    Dim dotTotN As Double, dotTotN1 As Double
    dotTotN = GetSoldeCompteN("681") + GetSoldeCompteN("686") + GetSoldeCompteN("687") + GetSoldeCompteN("696") + GetSoldeCompteN("697")
    dotTotN1 = GetSoldeCompteN1("681") + GetSoldeCompteN1("686") + GetSoldeCompteN1("687") + GetSoldeCompteN1("696") + GetSoldeCompteN1("697")
    cafN = resNetN + dotTotN - (-GetSoldeCompteN("781") - GetSoldeCompteN("786") - GetSoldeCompteN("787") - GetSoldeCompteN("791") - GetSoldeCompteN("796") - GetSoldeCompteN("797")) - (-GetSoldeCompteN("775") - GetSoldeCompteN("675")) - (-GetSoldeCompteN("777"))
    cafN1 = resNetN1 + dotTotN1 - (-GetSoldeCompteN1("781") - GetSoldeCompteN1("786") - GetSoldeCompteN1("787") - GetSoldeCompteN1("791") - GetSoldeCompteN1("796") - GetSoldeCompteN1("797")) - (-GetSoldeCompteN1("775") - GetSoldeCompteN1("675")) - (-GetSoldeCompteN1("777"))

    ' ---- Tableau KPIs ----
    ' Ligne KPI N
    ws.Cells(ligne, 1).Value = "INDICATEUR"
    ws.Cells(ligne, 3).Value = "Exercice N"
    ws.Cells(ligne, 5).Value = "Exercice N-1"
    ws.Cells(ligne, 7).Value = "Variation"
    ws.Cells(ligne, 9).Value = "Var. %"
    ws.Cells(ligne, 11).Value = "% CA"
    With ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 12))
        .Font.Bold = True
        .Interior.Color = RGB(31, 73, 125)
        .Font.Color = RGB(255, 255, 255)
        .Font.Size = 11
    End With
    ws.Rows(ligne).RowHeight = 22
    ligne = ligne + 1

    ' Liste des KPIs P&L
    Dim kpiLabels() As String
    Dim kpiValeursN() As Double, kpiValeursN1() As Double
    Dim kpiTypes() As String  ' "eur", "pct"
    Dim kpiColors() As Long

    ReDim kpiLabels(13), kpiValeursN(13), kpiValeursN1(13), kpiTypes(13), kpiColors(13)

    kpiLabels(0) = "Chiffre d'Affaires HT"
    kpiValeursN(0) = caN : kpiValeursN1(0) = caN1
    kpiTypes(0) = "eur" : kpiColors(0) = RGB(68, 114, 196)

    kpiLabels(1) = "Valeur Ajoutée (VA)"
    kpiValeursN(1) = vaH : kpiValeursN1(1) = vaH1
    kpiTypes(1) = "eur" : kpiColors(1) = RGB(70, 130, 180)

    kpiLabels(2) = "EBE (EBITDA)"
    kpiValeursN(2) = ebeN : kpiValeursN1(2) = ebeN1
    kpiTypes(2) = "eur" : kpiColors(2) = RGB(0, 128, 0)

    kpiLabels(3) = "Marge EBE / CA"
    If caN <> 0 Then kpiValeursN(3) = ebeN / caN Else kpiValeursN(3) = 0
    If caN1 <> 0 Then kpiValeursN1(3) = ebeN1 / caN1 Else kpiValeursN1(3) = 0
    kpiTypes(3) = "pct" : kpiColors(3) = RGB(0, 128, 0)

    kpiLabels(4) = "Résultat d'Exploitation (REX)"
    kpiValeursN(4) = rexN : kpiValeursN1(4) = rexN1
    kpiTypes(4) = "eur" : kpiColors(4) = RGB(0, 112, 0)

    kpiLabels(5) = "Marge REX / CA (EBIT margin)"
    If caN <> 0 Then kpiValeursN(5) = rexN / caN Else kpiValeursN(5) = 0
    If caN1 <> 0 Then kpiValeursN1(5) = rexN1 / caN1 Else kpiValeursN1(5) = 0
    kpiTypes(5) = "pct" : kpiColors(5) = RGB(0, 112, 0)

    kpiLabels(6) = "Résultat Courant Avant IS"
    kpiValeursN(6) = rcaN : kpiValeursN1(6) = rcaN1
    kpiTypes(6) = "eur" : kpiColors(6) = RGB(0, 96, 0)

    kpiLabels(7) = "Résultat Net de l'Exercice"
    kpiValeursN(7) = resNetN : kpiValeursN1(7) = resNetN1
    kpiTypes(7) = "eur" : kpiColors(7) = RGB(0, 70, 0)

    kpiLabels(8) = "Marge Nette / CA"
    If caN <> 0 Then kpiValeursN(8) = resNetN / caN Else kpiValeursN(8) = 0
    If caN1 <> 0 Then kpiValeursN1(8) = resNetN1 / caN1 Else kpiValeursN1(8) = 0
    kpiTypes(8) = "pct" : kpiColors(8) = RGB(0, 70, 0)

    kpiLabels(9) = "Capitaux Propres"
    kpiValeursN(9) = cpN : kpiValeursN1(9) = cpN1
    kpiTypes(9) = "eur" : kpiColors(9) = RGB(31, 73, 125)

    kpiLabels(10) = "Dettes Financières (LMT)"
    kpiValeursN(10) = dettesFinN : kpiValeursN1(10) = dettesFinN1
    kpiTypes(10) = "eur" : kpiColors(10) = RGB(192, 80, 77)

    kpiLabels(11) = "BFR"
    kpiValeursN(11) = bfrN : kpiValeursN1(11) = bfrN1
    kpiTypes(11) = "eur" : kpiColors(11) = RGB(197, 90, 17)

    kpiLabels(12) = "Trésorerie Nette"
    kpiValeursN(12) = tresNetteN : kpiValeursN1(12) = tresNetteN1
    kpiTypes(12) = "eur" : kpiColors(12) = RGB(0, 128, 0)

    kpiLabels(13) = "Capacité d'Autofinancement (CAF)"
    kpiValeursN(13) = cafN : kpiValeursN1(13) = cafN1
    kpiTypes(13) = "eur" : kpiColors(13) = RGB(31, 73, 125)

    Dim ki As Integer
    For ki = 0 To 13
        ' Label
        ws.Cells(ligne, 1).Value = kpiLabels(ki)
        With ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 2))
            .Merge
            .Interior.Color = kpiColors(ki)
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
            .Font.Size = 10
        End With

        ' Valeur N
        With ws.Range(ws.Cells(ligne, 3), ws.Cells(ligne, 4))
            .Merge
            .Value = kpiValeursN(ki)
            .Font.Bold = True
            .Font.Size = 11
            .HorizontalAlignment = xlRight
            If kpiTypes(ki) = "eur" Then
                .NumberFormat = "#,##0 €"
                If kpiValeursN(ki) >= 0 Then
                    .Interior.Color = RGB(235, 255, 235)
                    .Font.Color = RGB(0, 112, 0)
                Else
                    .Interior.Color = RGB(255, 230, 230)
                    .Font.Color = RGB(192, 0, 0)
                End If
            Else
                .NumberFormat = "0.0%"
                If kpiValeursN(ki) >= 0 Then
                    .Interior.Color = RGB(235, 255, 235)
                    .Font.Color = RGB(0, 112, 0)
                Else
                    .Interior.Color = RGB(255, 230, 230)
                    .Font.Color = RGB(192, 0, 0)
                End If
            End If
        End With

        ' Valeur N-1
        With ws.Range(ws.Cells(ligne, 5), ws.Cells(ligne, 6))
            .Merge
            .Value = kpiValeursN1(ki)
            .Font.Size = 10
            .HorizontalAlignment = xlRight
            .Interior.Color = RGB(240, 240, 240)
            .Font.Color = RGB(80, 80, 80)
            If kpiTypes(ki) = "eur" Then
                .NumberFormat = "#,##0 €"
            Else
                .NumberFormat = "0.0%"
            End If
        End With

        ' Variation
        Dim varAbs As Double, varPct As Double
        varAbs = kpiValeursN(ki) - kpiValeursN1(ki)
        If kpiValeursN1(ki) <> 0 Then varPct = varAbs / Abs(kpiValeursN1(ki))

        With ws.Range(ws.Cells(ligne, 7), ws.Cells(ligne, 8))
            .Merge
            .Value = varAbs
            .HorizontalAlignment = xlRight
            If kpiTypes(ki) = "eur" Then
                .NumberFormat = "+#,##0;-#,##0;0"
            Else
                .NumberFormat = "+0.0%;-0.0%;0.0%"
            End If
            If varAbs >= 0 Then
                .Font.Color = RGB(0, 112, 0)
            Else
                .Font.Color = RGB(192, 0, 0)
            End If
        End With

        With ws.Range(ws.Cells(ligne, 9), ws.Cells(ligne, 10))
            .Merge
            .Value = varPct
            .NumberFormat = "+0.0%;-0.0%;--"
            .HorizontalAlignment = xlCenter
            If varPct >= 0 Then
                .Font.Color = RGB(0, 112, 0)
                .Value = "▲ " & Format(Abs(varPct), "0.0%")
            Else
                .Font.Color = RGB(192, 0, 0)
                .Value = "▼ " & Format(Abs(varPct), "0.0%")
            End If
        End With

        ' % CA
        If kpiTypes(ki) = "eur" And caN <> 0 Then
            With ws.Range(ws.Cells(ligne, 11), ws.Cells(ligne, 12))
                .Merge
                .Value = kpiValeursN(ki) / caN
                .NumberFormat = "0.0%"
                .HorizontalAlignment = xlCenter
                .Font.Color = RGB(80, 80, 80)
            End With
        End If

        ws.Rows(ligne).RowHeight = 20
        ligne = ligne + 1
    Next ki

    ligne = ligne + 2

    ' ---- Graphique textuel : Cascade résultat ----
    Call EcrireEntitePPT(ws, ligne, "CASCADE DE FORMATION DU RESULTAT")
    ligne = ligne + 1

    ws.Cells(ligne, 1).Value = "Chiffre d'Affaires"
    ws.Cells(ligne, 3).Value = caN
    ws.Cells(ligne, 3).NumberFormat = "#,##0 €"
    ws.Cells(ligne, 3).Interior.Color = RGB(68, 114, 196)
    ws.Cells(ligne, 3).Font.Color = RGB(255, 255, 255)
    ws.Cells(ligne, 3).Font.Bold = True
    ligne = ligne + 1

    ws.Cells(ligne, 1).Value = "(-) Achats et charges externes"
    ws.Cells(ligne, 3).Value = GetSoldeCompteN("60") + GetSoldeCompteN("61") + GetSoldeCompteN("62")
    ws.Cells(ligne, 3).NumberFormat = "#,##0 €"
    ws.Cells(ligne, 3).Interior.Color = RGB(255, 199, 206)
    ligne = ligne + 1

    ws.Cells(ligne, 1).Value = "= Valeur Ajoutée"
    ws.Cells(ligne, 3).Value = vaH
    ws.Cells(ligne, 3).NumberFormat = "#,##0 €"
    ws.Cells(ligne, 3).Interior.Color = RGB(0, 128, 0)
    ws.Cells(ligne, 3).Font.Color = RGB(255, 255, 255)
    ws.Cells(ligne, 3).Font.Bold = True
    ws.Cells(ligne, 4).Value = IIf(caN <> 0, vaH / caN, 0)
    ws.Cells(ligne, 4).NumberFormat = "0.0% du CA"
    ligne = ligne + 1

    ws.Cells(ligne, 1).Value = "(-) Charges de personnel + Impôts taxes"
    ws.Cells(ligne, 3).Value = personnelN + impotsTaxesN
    ws.Cells(ligne, 3).NumberFormat = "#,##0 €"
    ws.Cells(ligne, 3).Interior.Color = RGB(255, 199, 206)
    ligne = ligne + 1

    ws.Cells(ligne, 1).Value = "= EBE (EBITDA)"
    ws.Cells(ligne, 3).Value = ebeN
    ws.Cells(ligne, 3).NumberFormat = "#,##0 €"
    ws.Cells(ligne, 3).Interior.Color = RGB(0, 112, 0)
    ws.Cells(ligne, 3).Font.Color = RGB(255, 255, 255)
    ws.Cells(ligne, 3).Font.Bold = True
    ws.Cells(ligne, 4).Value = IIf(caN <> 0, ebeN / caN, 0)
    ws.Cells(ligne, 4).NumberFormat = "0.0% du CA"
    ligne = ligne + 1

    ws.Cells(ligne, 1).Value = "(-) Dotations amortissements"
    ws.Cells(ligne, 3).Value = dotN
    ws.Cells(ligne, 3).NumberFormat = "#,##0 €"
    ws.Cells(ligne, 3).Interior.Color = RGB(255, 199, 206)
    ligne = ligne + 1

    ws.Cells(ligne, 1).Value = "= REX (EBIT)"
    ws.Cells(ligne, 3).Value = rexN
    ws.Cells(ligne, 3).NumberFormat = "#,##0 €"
    ws.Cells(ligne, 3).Interior.Color = RGB(0, 96, 0)
    ws.Cells(ligne, 3).Font.Color = RGB(255, 255, 255)
    ws.Cells(ligne, 3).Font.Bold = True
    ws.Cells(ligne, 4).Value = IIf(caN <> 0, rexN / caN, 0)
    ws.Cells(ligne, 4).NumberFormat = "0.0% du CA"
    ligne = ligne + 1

    ws.Cells(ligne, 1).Value = "(±) Résultat financier"
    ws.Cells(ligne, 3).Value = prodFinN - chargesFinN
    ws.Cells(ligne, 3).NumberFormat = "#,##0 €"
    ws.Cells(ligne, 3).Interior.Color = RGB(189, 215, 238)
    ligne = ligne + 1

    ws.Cells(ligne, 1).Value = "(±) Résultat exceptionnel"
    ws.Cells(ligne, 3).Value = resExcN
    ws.Cells(ligne, 3).NumberFormat = "#,##0 €"
    ws.Cells(ligne, 3).Interior.Color = RGB(189, 215, 238)
    ligne = ligne + 1

    ws.Cells(ligne, 1).Value = "(-) IS + Participation"
    ws.Cells(ligne, 3).Value = isN
    ws.Cells(ligne, 3).NumberFormat = "#,##0 €"
    ws.Cells(ligne, 3).Interior.Color = RGB(255, 199, 206)
    ligne = ligne + 1

    ws.Cells(ligne, 1).Value = "= RÉSULTAT NET"
    ws.Cells(ligne, 3).Value = resNetN
    ws.Cells(ligne, 3).NumberFormat = "#,##0 €"
    ws.Cells(ligne, 3).Font.Bold = True
    ws.Cells(ligne, 3).Font.Size = 12
    If resNetN >= 0 Then
        ws.Cells(ligne, 3).Interior.Color = RGB(0, 70, 0)
        ws.Cells(ligne, 3).Font.Color = RGB(255, 255, 255)
    Else
        ws.Cells(ligne, 3).Interior.Color = RGB(150, 0, 0)
        ws.Cells(ligne, 3).Font.Color = RGB(255, 255, 255)
    End If
    ws.Cells(ligne, 4).Value = IIf(caN <> 0, resNetN / caN, 0)
    ws.Cells(ligne, 4).NumberFormat = "0.0% du CA"
    ws.Cells(ligne, 4).Font.Bold = True
    ligne = ligne + 3

    ' ---- Section Bilan simplifie ----
    Call EcrireEntitePPT(ws, ligne, "BILAN SIMPLIFIE")
    ligne = ligne + 1

    ' Actif / Passif en colonnes
    ws.Cells(ligne, 1).Value = "ACTIF"
    ws.Cells(ligne, 7).Value = "PASSIF"
    ws.Cells(ligne, 1).Font.Bold = True
    ws.Cells(ligne, 7).Font.Bold = True
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 6)).Interior.Color = RGB(31, 73, 125)
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 6)).Font.Color = RGB(255, 255, 255)
    ws.Range(ws.Cells(ligne, 7), ws.Cells(ligne, 12)).Interior.Color = RGB(192, 80, 77)
    ws.Range(ws.Cells(ligne, 7), ws.Cells(ligne, 12)).Font.Color = RGB(255, 255, 255)
    ligne = ligne + 1

    ws.Cells(ligne, 1).Value = "Actif immobilisé net"
    ws.Cells(ligne, 3).Value = actifImmN
    ws.Cells(ligne, 3).NumberFormat = "#,##0 €"
    ws.Cells(ligne, 7).Value = "Capitaux propres"
    ws.Cells(ligne, 9).Value = cpN
    ws.Cells(ligne, 9).NumberFormat = "#,##0 €"
    ws.Cells(ligne, 9).Font.Color = RGB(0, 112, 0)
    ligne = ligne + 1

    ws.Cells(ligne, 1).Value = "Stocks"
    ws.Cells(ligne, 3).Value = stocksN
    ws.Cells(ligne, 3).NumberFormat = "#,##0 €"
    ws.Cells(ligne, 7).Value = "Dettes financières LMT"
    ws.Cells(ligne, 9).Value = dettesFinN
    ws.Cells(ligne, 9).NumberFormat = "#,##0 €"
    ws.Cells(ligne, 9).Font.Color = RGB(192, 80, 77)
    ligne = ligne + 1

    ws.Cells(ligne, 1).Value = "Créances clients"
    ws.Cells(ligne, 3).Value = creCliN
    ws.Cells(ligne, 3).NumberFormat = "#,##0 €"
    ws.Cells(ligne, 7).Value = "Dettes fournisseurs"
    ws.Cells(ligne, 9).Value = detFournN
    ws.Cells(ligne, 9).NumberFormat = "#,##0 €"
    ligne = ligne + 1

    ws.Cells(ligne, 1).Value = "Tresorerie active"
    Dim tresActN As Double
    tresActN = GetSoldeCompteN("50") + GetSoldeCompteN("51") + GetSoldeCompteN("53") + GetSoldeCompteN("54")
    ws.Cells(ligne, 3).Value = tresActN
    ws.Cells(ligne, 3).NumberFormat = "#,##0 €"
    ws.Cells(ligne, 7).Value = "Dettes fiscales et sociales + Autres"
    ws.Cells(ligne, 9).Value = detFSN
    ws.Cells(ligne, 9).NumberFormat = "#,##0 €"
    ligne = ligne + 2

    ' Total actif
    Dim totActN As Double, totPassN As Double
    totActN = actifImmN + stocksN + creCliN + tresActN
    totPassN = cpN + dettesFinN + detFournN + detFSN

    ws.Cells(ligne, 1).Value = "TOTAL ACTIF (simplifié)"
    ws.Cells(ligne, 3).Value = totActN
    ws.Cells(ligne, 3).NumberFormat = "#,##0 €"
    ws.Cells(ligne, 3).Font.Bold = True
    ws.Cells(ligne, 7).Value = "TOTAL PASSIF (simplifié)"
    ws.Cells(ligne, 9).Value = totPassN
    ws.Cells(ligne, 9).NumberFormat = "#,##0 €"
    ws.Cells(ligne, 9).Font.Bold = True
    ligne = ligne + 3

    ' ---- Points d'attention ----
    Call EcrireEntitePPT(ws, ligne, "POINTS D'ATTENTION ET ANALYSE")
    ligne = ligne + 1

    ' Génération automatique de commentaires
    Dim alertes() As String
    ReDim alertes(9)
    Dim nbAlertes As Integer
    nbAlertes = 0

    ' Croissance CA
    If caN1 <> 0 Then
        Dim croisCA As Double
        croisCA = (caN - caN1) / Abs(caN1)
        If croisCA > 0.1 Then
            alertes(nbAlertes) = "✓ POSITIF : Forte croissance du CA de " & Format(croisCA, "0.0%") & " sur l'exercice"
            nbAlertes = nbAlertes + 1
        ElseIf croisCA < -0.05 Then
            alertes(nbAlertes) = "⚠ ATTENTION : Baisse du CA de " & Format(Abs(croisCA), "0.0%") & " - Analyser les causes"
            nbAlertes = nbAlertes + 1
        End If
    End If

    ' EBE
    If caN <> 0 Then
        Dim margeEbe As Double
        margeEbe = ebeN / caN
        If margeEbe < 0.05 Then
            alertes(nbAlertes) = "⚠ ATTENTION : Marge EBE faible (" & Format(margeEbe, "0.0%") & ") - Surveiller la rentabilité opérationnelle"
            nbAlertes = nbAlertes + 1
        ElseIf margeEbe > 0.2 Then
            alertes(nbAlertes) = "✓ POSITIF : Excellente marge EBE (" & Format(margeEbe, "0.0%") & ")"
            nbAlertes = nbAlertes + 1
        End If
    End If

    ' Résultat net
    If resNetN < 0 Then
        alertes(nbAlertes) = "⚠ ALERTE : Résultat net négatif (" & Format(resNetN, "#,##0 €") & ") - Risque de capitaux propres négatifs"
        nbAlertes = nbAlertes + 1
    End If

    ' Trésorerie
    If tresNetteN < 0 Then
        alertes(nbAlertes) = "⚠ ALERTE : Trésorerie nette négative (" & Format(tresNetteN, "#,##0 €") & ") - Dépendance aux CBC"
        nbAlertes = nbAlertes + 1
    End If

    ' BFR
    If caN <> 0 Then
        Dim bfrJours As Double
        bfrJours = (bfrN / caN) * 365
        If bfrJours > 90 Then
            alertes(nbAlertes) = "⚠ ATTENTION : BFR élevé (" & Format(bfrJours, "0") & " jours de CA) - Optimiser le cycle d'exploitation"
            nbAlertes = nbAlertes + 1
        End If
    End If

    ' Levier financier
    If cpN <> 0 Then
        Dim levier As Double
        levier = dettesFinN / cpN
        If levier > 2 Then
            alertes(nbAlertes) = "⚠ ATTENTION : Levier financier élevé (" & Format(levier, "0.0") & "x) - Dette/CP = " & Format(levier, "0.0") & "x"
            nbAlertes = nbAlertes + 1
        End If
    End If

    ' Couverture intérêts
    If chargesFinN > 0 And ebeN <> 0 Then
        Dim couverture As Double
        couverture = ebeN / chargesFinN
        If couverture < 3 Then
            alertes(nbAlertes) = "⚠ ATTENTION : Couverture des intérêts faible (EBE/Charges fin. = " & Format(couverture, "0.0") & "x)"
            nbAlertes = nbAlertes + 1
        End If
    End If

    ' CAF vs BFR
    If cafN > 0 And bfrN > 0 Then
        If cafN < bfrN * 0.5 Then
            alertes(nbAlertes) = "⚠ ATTENTION : CAF insuffisante pour financer la croissance du BFR"
            nbAlertes = nbAlertes + 1
        End If
    End If

    Dim ai As Integer
    For ai = 0 To nbAlertes - 1
        ws.Cells(ligne, 1).Value = alertes(ai)
        If Left(alertes(ai), 1) = "✓" Then
            ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 12)).Interior.Color = RGB(235, 255, 235)
            ws.Cells(ligne, 1).Font.Color = RGB(0, 112, 0)
        Else
            ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 12)).Interior.Color = RGB(255, 245, 220)
            ws.Cells(ligne, 1).Font.Color = RGB(197, 90, 17)
        End If
        ws.Cells(ligne, 1).Font.Bold = True
        ws.Rows(ligne).RowHeight = 22
        ligne = ligne + 1
    Next ai

    ' Formatage final
    ws.Columns("A:B").ColumnWidth = 35
    ws.Columns("C:D").ColumnWidth = 16
    ws.Columns("E:F").ColumnWidth = 12
    ws.Columns("G:H").ColumnWidth = 30
    ws.Columns("I:J").ColumnWidth = 16
    ws.Columns("K:L").ColumnWidth = 12
    ws.DisplayGridlines = False
    ws.PageSetup.Orientation = xlLandscape
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.FitToPagesTall = False
    ws.PageSetup.Zoom = False
End Sub

Sub EcrireEntitePPT(ws As Worksheet, ligne As Long, Titre As String)
    With ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 12))
        .Merge
        .Value = Titre
        .Interior.Color = RGB(31, 73, 125)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 12
        .HorizontalAlignment = xlLeft
        .IndentLevel = 1
    End With
    ws.Rows(ligne).RowHeight = 25
End Sub
