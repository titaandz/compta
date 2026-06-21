Attribute VB_Name = "Module4_Recaps_SIG"
Option Explicit

' ============================================================
' HELPER P&L - doit etre au niveau module (pas imbriquee)
' ============================================================
Sub EcrireLignePL(ws As Worksheet, ligne As Long, libelle As String, note As String, valN As Double, valN1 As Double, caRef As Double)
    ws.Cells(ligne, 1).Value = libelle
    ws.Cells(ligne, 2).Value = note
    ws.Cells(ligne, 3).Value = valN
    ws.Cells(ligne, 4).Value = valN1
    ws.Cells(ligne, 5).Value = valN - valN1
    If valN1 <> 0 Then ws.Cells(ligne, 6).Value = (valN - valN1) / Abs(valN1)
    If caRef <> 0 Then ws.Cells(ligne, 7).Value = valN / caRef
    If ligne Mod 2 = 0 Then ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 7)).Interior.Color = RGB(242, 242, 242)
End Sub

' ============================================================
' RECAPITULATIF BILAN
' ============================================================
Sub CreerRecapBilan()
    Dim ws As Worksheet
    Dim ligne As Long

    Set ws = CreerOnglet("Recap Bilan", COULEUR_RECAP)

    With ws.Range("A1:I1")
        .Merge
        .Value = "BILAN COMPTABLE RECAPITULATIF - Normes PCG / ANC"
        .Interior.Color = RGB(31, 73, 125)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
    End With
    With ws.Range("A2:I2")
        .Merge
        .Value = "Exercices N et N-1 - en euros"
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With

    ws.Cells(3, 1).Value = "Poste bilan"
    ws.Cells(3, 2).Value = "Notes"
    ws.Cells(3, 3).Value = "Brut N"
    ws.Cells(3, 4).Value = "Amort/Dep N"
    ws.Cells(3, 5).Value = "Net N"
    ws.Cells(3, 6).Value = "Net N-1"
    ws.Cells(3, 7).Value = "Variation"
    ws.Cells(3, 8).Value = "Var. %"
    ws.Cells(3, 9).Value = "% Total"
    With ws.Range("A3:I3")
        .Font.Bold = True
        .Interior.Color = RGB(189, 215, 238)
        .HorizontalAlignment = xlCenter
    End With

    ligne = 4

    ' ==================== ACTIF ====================
    ws.Cells(ligne, 1).Value = "ACTIF"
    With ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 9))
        .Interior.Color = RGB(31, 73, 125)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 12
    End With
    ligne = ligne + 1

    ws.Cells(ligne, 1).Value = "ACTIF IMMOBILISE"
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 9)).Interior.Color = RGB(189, 215, 238)
    ws.Cells(ligne, 1).Font.Bold = True
    ligne = ligne + 1

    Dim iiBrutN As Double, iiAmortN As Double, iiNetN As Double, iiNetN1 As Double
    iiBrutN = GetSoldeCompteN("20")
    iiAmortN = GetSoldeCompteN("280")
    iiNetN = iiBrutN + iiAmortN
    iiNetN1 = GetSoldeCompteN1("20") + GetSoldeCompteN1("280")
    Call EcrireLigneBilan(ws, ligne, "Immobilisations incorporelles", "Note 1", iiBrutN, iiAmortN, iiNetN, iiNetN1)
    ligne = ligne + 1

    Dim icBrutN As Double, icAmortN As Double, icNetN As Double, icNetN1 As Double
    icBrutN = GetSoldeCompteN("21") + GetSoldeCompteN("22") + GetSoldeCompteN("23")
    icAmortN = GetSoldeCompteN("28")
    icNetN = icBrutN + icAmortN
    icNetN1 = GetSoldeCompteN1("21") + GetSoldeCompteN1("22") + GetSoldeCompteN1("23") + GetSoldeCompteN1("28")
    Call EcrireLigneBilan(ws, ligne, "Immobilisations corporelles", "Note 2", icBrutN, icAmortN, icNetN, icNetN1)
    ligne = ligne + 1

    Dim ifBrutN As Double, ifDepN As Double, ifNetN As Double, ifNetN1 As Double
    ifBrutN = GetSoldeCompteN("26") + GetSoldeCompteN("27")
    ifDepN = GetSoldeCompteN("296") + GetSoldeCompteN("297")
    ifNetN = ifBrutN + ifDepN
    ifNetN1 = GetSoldeCompteN1("26") + GetSoldeCompteN1("27") + GetSoldeCompteN1("296") + GetSoldeCompteN1("297")
    Call EcrireLigneBilan(ws, ligne, "Immobilisations financieres", "Note 3", ifBrutN, ifDepN, ifNetN, ifNetN1)
    ligne = ligne + 1

    Dim totalActImmN As Double, totalActImmN1 As Double
    totalActImmN = iiNetN + icNetN + ifNetN
    totalActImmN1 = iiNetN1 + icNetN1 + ifNetN1
    Call EcrireSousTotal9(ws, ligne, "TOTAL ACTIF IMMOBILISE", 0, 0, totalActImmN, totalActImmN1)
    ligne = ligne + 2

    ws.Cells(ligne, 1).Value = "ACTIF CIRCULANT"
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 9)).Interior.Color = RGB(189, 215, 238)
    ws.Cells(ligne, 1).Font.Bold = True
    ligne = ligne + 1

    Dim stBrutN As Double, stDepN As Double, stNetN As Double, stNetN1 As Double
    stBrutN = GetSoldeCompteN("31") + GetSoldeCompteN("32") + GetSoldeCompteN("33") + GetSoldeCompteN("34") + GetSoldeCompteN("35") + GetSoldeCompteN("37")
    stDepN = GetSoldeCompteN("39")
    stNetN = stBrutN + stDepN
    stNetN1 = GetSoldeCompteN1("31") + GetSoldeCompteN1("32") + GetSoldeCompteN1("33") + GetSoldeCompteN1("34") + GetSoldeCompteN1("35") + GetSoldeCompteN1("37") + GetSoldeCompteN1("39")
    Call EcrireLigneBilan(ws, ligne, "Stocks et en-cours", "Note 4", stBrutN, stDepN, stNetN, stNetN1)
    ligne = ligne + 1

    Dim cliBrutN As Double, cliDepN As Double, cliNetN As Double, cliNetN1 As Double
    cliBrutN = GetSoldeCompteN("41")
    cliDepN = GetSoldeCompteN("490")
    cliNetN = cliBrutN + cliDepN
    cliNetN1 = GetSoldeCompteN1("41") + GetSoldeCompteN1("490")
    Call EcrireLigneBilan(ws, ligne, "Creances clients", "Note 5", cliBrutN, cliDepN, cliNetN, cliNetN1)
    ligne = ligne + 1

    Dim autBrutN As Double, autNetN As Double, autNetN1 As Double
    autBrutN = GetSoldeCompteN("42") + GetSoldeCompteN("43") + GetSoldeCompteN("44") + GetSoldeCompteN("45") + GetSoldeCompteN("46") + GetSoldeCompteN("48")
    autNetN = autBrutN
    autNetN1 = GetSoldeCompteN1("42") + GetSoldeCompteN1("43") + GetSoldeCompteN1("44") + GetSoldeCompteN1("45") + GetSoldeCompteN1("46") + GetSoldeCompteN1("48")
    Call EcrireLigneBilan(ws, ligne, "Autres creances et regularisation", "Note 6", autBrutN, 0, autNetN, autNetN1)
    ligne = ligne + 1

    Dim tresBrutN As Double, tresNetN As Double, tresNetN1 As Double
    tresBrutN = GetSoldeCompteN("50") + GetSoldeCompteN("51") + GetSoldeCompteN("53") + GetSoldeCompteN("54")
    tresNetN = tresBrutN
    tresNetN1 = GetSoldeCompteN1("50") + GetSoldeCompteN1("51") + GetSoldeCompteN1("53") + GetSoldeCompteN1("54")
    Call EcrireLigneBilan(ws, ligne, "Disponibilites et VMP", "Note 7", tresBrutN, 0, tresNetN, tresNetN1)
    ligne = ligne + 1

    Dim totalActCircN As Double, totalActCircN1 As Double
    totalActCircN = stNetN + cliNetN + autNetN + tresNetN
    totalActCircN1 = stNetN1 + cliNetN1 + autNetN1 + tresNetN1
    Call EcrireSousTotal9(ws, ligne, "TOTAL ACTIF CIRCULANT", 0, 0, totalActCircN, totalActCircN1)
    ligne = ligne + 2

    Dim totalActifN As Double, totalActifN1 As Double
    totalActifN = totalActImmN + totalActCircN
    totalActifN1 = totalActImmN1 + totalActCircN1
    Call EcrireTotalBilan(ws, ligne, "TOTAL ACTIF", totalActifN, totalActifN1)
    ligne = ligne + 3

    ' ==================== PASSIF ====================
    ws.Cells(ligne, 1).Value = "PASSIF"
    With ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 9))
        .Interior.Color = RGB(31, 73, 125)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 12
    End With
    ligne = ligne + 1

    ws.Cells(ligne, 1).Value = "CAPITAUX PROPRES"
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 9)).Interior.Color = RGB(189, 215, 238)
    ws.Cells(ligne, 1).Font.Bold = True
    ligne = ligne + 1

    Dim capN As Double, capN1 As Double
    capN = -(GetSoldeCompteN("101") + GetSoldeCompteN("104") + GetSoldeCompteN("105") + GetSoldeCompteN("106") + GetSoldeCompteN("107") + GetSoldeCompteN("108") + GetSoldeCompteN("11") + GetSoldeCompteN("13") + GetSoldeCompteN("14"))
    capN1 = -(GetSoldeCompteN1("101") + GetSoldeCompteN1("104") + GetSoldeCompteN1("105") + GetSoldeCompteN1("106") + GetSoldeCompteN1("107") + GetSoldeCompteN1("108") + GetSoldeCompteN1("11") + GetSoldeCompteN1("13") + GetSoldeCompteN1("14"))

    ws.Cells(ligne, 1).Value = "Capital, reserves et report a nouveau"
    ws.Cells(ligne, 2).Value = "Note 8"
    ws.Cells(ligne, 5).Value = capN
    ws.Cells(ligne, 6).Value = capN1
    ws.Cells(ligne, 7).Value = capN - capN1
    If capN1 <> 0 Then ws.Cells(ligne, 8).Value = (capN - capN1) / Abs(capN1)
    If totalActifN <> 0 Then ws.Cells(ligne, 9).Value = capN / totalActifN
    If ligne Mod 2 = 0 Then ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 9)).Interior.Color = RGB(242, 242, 242)
    ligne = ligne + 1

    Dim resN As Double, resN1 As Double
    resN = -GetSoldeCompteN("12")
    resN1 = -GetSoldeCompteN1("12")
    ws.Cells(ligne, 1).Value = "Resultat de l'exercice"
    ws.Cells(ligne, 5).Value = resN
    ws.Cells(ligne, 6).Value = resN1
    ws.Cells(ligne, 7).Value = resN - resN1
    If resN1 <> 0 Then ws.Cells(ligne, 8).Value = (resN - resN1) / Abs(resN1)
    If totalActifN <> 0 Then ws.Cells(ligne, 9).Value = resN / totalActifN
    If ligne Mod 2 = 0 Then ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 9)).Interior.Color = RGB(242, 242, 242)
    ligne = ligne + 1

    Dim totalCPN As Double, totalCPN1 As Double
    totalCPN = capN + resN
    totalCPN1 = capN1 + resN1
    Call EcrireSousTotal9(ws, ligne, "TOTAL CAPITAUX PROPRES", 0, 0, totalCPN, totalCPN1)
    ligne = ligne + 2

    Dim prcN As Double, prcN1 As Double
    prcN = -GetSoldeCompteN("15")
    prcN1 = -GetSoldeCompteN1("15")
    ws.Cells(ligne, 1).Value = "Provisions pour risques et charges"
    ws.Cells(ligne, 2).Value = "Note 9"
    ws.Cells(ligne, 5).Value = prcN
    ws.Cells(ligne, 6).Value = prcN1
    ws.Cells(ligne, 7).Value = prcN - prcN1
    If prcN1 <> 0 Then ws.Cells(ligne, 8).Value = (prcN - prcN1) / Abs(prcN1)
    If totalActifN <> 0 Then ws.Cells(ligne, 9).Value = prcN / totalActifN
    ligne = ligne + 2

    ws.Cells(ligne, 1).Value = "DETTES"
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 9)).Interior.Color = RGB(189, 215, 238)
    ws.Cells(ligne, 1).Font.Bold = True
    ligne = ligne + 1

    Dim dfN As Double, dfN1 As Double
    dfN = -(GetSoldeCompteN("16") + GetSoldeCompteN("519"))
    dfN1 = -(GetSoldeCompteN1("16") + GetSoldeCompteN1("519"))
    ws.Cells(ligne, 1).Value = "Dettes financieres (emprunts + CBC)"
    ws.Cells(ligne, 2).Value = "Note 10"
    ws.Cells(ligne, 5).Value = dfN
    ws.Cells(ligne, 6).Value = dfN1
    ws.Cells(ligne, 7).Value = dfN - dfN1
    If dfN1 <> 0 Then ws.Cells(ligne, 8).Value = (dfN - dfN1) / Abs(dfN1)
    If totalActifN <> 0 Then ws.Cells(ligne, 9).Value = dfN / totalActifN
    If ligne Mod 2 = 0 Then ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 9)).Interior.Color = RGB(242, 242, 242)
    ligne = ligne + 1

    Dim dfournN As Double, dfournN1 As Double
    dfournN = -GetSoldeCompteN("40")
    dfournN1 = -GetSoldeCompteN1("40")
    ws.Cells(ligne, 1).Value = "Dettes fournisseurs et comptes rattaches"
    ws.Cells(ligne, 5).Value = dfournN
    ws.Cells(ligne, 6).Value = dfournN1
    ws.Cells(ligne, 7).Value = dfournN - dfournN1
    If dfournN1 <> 0 Then ws.Cells(ligne, 8).Value = (dfournN - dfournN1) / Abs(dfournN1)
    If totalActifN <> 0 Then ws.Cells(ligne, 9).Value = dfournN / totalActifN
    If ligne Mod 2 = 0 Then ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 9)).Interior.Color = RGB(242, 242, 242)
    ligne = ligne + 1

    Dim dfsN As Double, dfsN1 As Double
    dfsN = Abs(GetSoldeCompteN("42")) + Abs(GetSoldeCompteN("43")) + Abs(GetSoldeCompteN("44")) + Abs(GetSoldeCompteN("45"))
    dfsN1 = Abs(GetSoldeCompteN1("42")) + Abs(GetSoldeCompteN1("43")) + Abs(GetSoldeCompteN1("44")) + Abs(GetSoldeCompteN1("45"))
    ws.Cells(ligne, 1).Value = "Dettes fiscales et sociales"
    ws.Cells(ligne, 5).Value = dfsN
    ws.Cells(ligne, 6).Value = dfsN1
    ws.Cells(ligne, 7).Value = dfsN - dfsN1
    If dfsN1 <> 0 Then ws.Cells(ligne, 8).Value = (dfsN - dfsN1) / Abs(dfsN1)
    If totalActifN <> 0 Then ws.Cells(ligne, 9).Value = dfsN / totalActifN
    If ligne Mod 2 = 0 Then ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 9)).Interior.Color = RGB(242, 242, 242)
    ligne = ligne + 1

    Dim adN As Double, adN1 As Double
    adN = Abs(GetSoldeCompteN("46")) + Abs(GetSoldeCompteN("487"))
    adN1 = Abs(GetSoldeCompteN1("46")) + Abs(GetSoldeCompteN1("487"))
    ws.Cells(ligne, 1).Value = "Autres dettes et comptes de regularisation"
    ws.Cells(ligne, 5).Value = adN
    ws.Cells(ligne, 6).Value = adN1
    ws.Cells(ligne, 7).Value = adN - adN1
    If adN1 <> 0 Then ws.Cells(ligne, 8).Value = (adN - adN1) / Abs(adN1)
    If totalActifN <> 0 Then ws.Cells(ligne, 9).Value = adN / totalActifN
    If ligne Mod 2 = 0 Then ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 9)).Interior.Color = RGB(242, 242, 242)
    ligne = ligne + 1

    Dim totalDettesN As Double, totalDettesN1 As Double
    totalDettesN = dfN + dfournN + dfsN + adN
    totalDettesN1 = dfN1 + dfournN1 + dfsN1 + adN1
    Call EcrireSousTotal9(ws, ligne, "TOTAL DETTES", 0, 0, totalDettesN, totalDettesN1)
    ligne = ligne + 2

    Dim totalPassifN As Double, totalPassifN1 As Double
    totalPassifN = totalCPN + prcN + totalDettesN
    totalPassifN1 = totalCPN1 + prcN1 + totalDettesN1
    Call EcrireTotalBilan(ws, ligne, "TOTAL PASSIF", totalPassifN, totalPassifN1)
    ligne = ligne + 2

    ws.Cells(ligne, 1).Value = "VERIFICATION : Actif - Passif (doit etre proche de zero)"
    ws.Cells(ligne, 5).Value = totalActifN - totalPassifN
    ws.Cells(ligne, 6).Value = totalActifN1 - totalPassifN1
    ws.Cells(ligne, 5).NumberFormat = "#,##0;[Red]-#,##0"
    ws.Cells(ligne, 6).NumberFormat = "#,##0;[Red]-#,##0"
    ws.Cells(ligne, 1).Font.Italic = True
    ws.Cells(ligne, 1).Font.Color = RGB(128, 128, 128)

    ws.Columns("A").ColumnWidth = 45
    ws.Columns("B").ColumnWidth = 10
    ws.Columns("C:I").ColumnWidth = 16
    ws.Range("C4:H500").NumberFormat = "#,##0;[Red]-#,##0"
    ws.Range("H4:H500").NumberFormat = "0.0%"
    ws.Range("I4:I500").NumberFormat = "0.0%"
End Sub

Sub EcrireLigneBilan(ws As Worksheet, Ligne As Long, Libelle As String, Note As String, Brut As Double, AmortDep As Double, NetN As Double, NetN1 As Double)
    ws.Cells(Ligne, 1).Value = Libelle
    ws.Cells(Ligne, 2).Value = Note
    ws.Cells(Ligne, 3).Value = Brut
    ws.Cells(Ligne, 4).Value = AmortDep
    ws.Cells(Ligne, 5).Value = NetN
    ws.Cells(Ligne, 6).Value = NetN1
    ws.Cells(Ligne, 7).Value = NetN - NetN1
    If NetN1 <> 0 Then ws.Cells(Ligne, 8).Value = (NetN - NetN1) / Abs(NetN1)
    If Ligne Mod 2 = 0 Then ws.Range(ws.Cells(Ligne, 1), ws.Cells(Ligne, 9)).Interior.Color = RGB(242, 242, 242)
End Sub

Sub EcrireSousTotal9(ws As Worksheet, Ligne As Long, Libelle As String, Brut As Double, AmortDep As Double, NetN As Double, NetN1 As Double)
    ws.Cells(Ligne, 1).Value = Libelle
    ws.Cells(Ligne, 3).Value = Brut
    ws.Cells(Ligne, 4).Value = AmortDep
    ws.Cells(Ligne, 5).Value = NetN
    ws.Cells(Ligne, 6).Value = NetN1
    ws.Cells(Ligne, 7).Value = NetN - NetN1
    If NetN1 <> 0 Then ws.Cells(Ligne, 8).Value = (NetN - NetN1) / Abs(NetN1)
    With ws.Range(ws.Cells(Ligne, 1), ws.Cells(Ligne, 9))
        .Font.Bold = True
        .Interior.Color = RGB(189, 215, 238)
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlDouble
    End With
End Sub

Sub EcrireTotalBilan(ws As Worksheet, Ligne As Long, Libelle As String, NetN As Double, NetN1 As Double)
    ws.Cells(Ligne, 1).Value = Libelle
    ws.Cells(Ligne, 5).Value = NetN
    ws.Cells(Ligne, 6).Value = NetN1
    ws.Cells(Ligne, 7).Value = NetN - NetN1
    If NetN1 <> 0 Then ws.Cells(Ligne, 8).Value = (NetN - NetN1) / Abs(NetN1)
    With ws.Range(ws.Cells(Ligne, 1), ws.Cells(Ligne, 9))
        .Font.Bold = True
        .Font.Size = 12
        .Interior.Color = RGB(31, 73, 125)
        .Font.Color = RGB(255, 255, 255)
    End With
End Sub

' ============================================================
' RECAPITULATIF P&L
' ============================================================
Sub CreerRecapPL()
    Dim ws As Worksheet
    Dim ligne As Long

    Set ws = CreerOnglet("Recap PL", COULEUR_PL)

    With ws.Range("A1:G1")
        .Merge
        .Value = "COMPTE DE RESULTAT - Normes PCG / ANC"
        .Interior.Color = RGB(0, 112, 0)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
    End With
    With ws.Range("A2:G2")
        .Merge
        .Value = "Exercices N et N-1 - en euros"
        .Interior.Color = RGB(0, 176, 0)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With

    ws.Cells(3, 1).Value = "Rubrique"
    ws.Cells(3, 2).Value = "Notes"
    ws.Cells(3, 3).Value = "Montant N"
    ws.Cells(3, 4).Value = "Montant N-1"
    ws.Cells(3, 5).Value = "Variation"
    ws.Cells(3, 6).Value = "Var. %"
    ws.Cells(3, 7).Value = "% CA"
    With ws.Range("A3:G3")
        .Font.Bold = True
        .Interior.Color = RGB(198, 224, 180)
        .HorizontalAlignment = xlCenter
    End With

    ligne = 4

    Dim caN As Double, caN1 As Double
    caN = -GetSoldeCompteN("70")
    caN1 = -GetSoldeCompteN1("70")

    Dim prodStockN As Double, prodStockN1 As Double
    prodStockN = -GetSoldeCompteN("71")
    prodStockN1 = -GetSoldeCompteN1("71")

    Dim prodImmN As Double, prodImmN1 As Double
    prodImmN = -GetSoldeCompteN("72")
    prodImmN1 = -GetSoldeCompteN1("72")

    Dim subvExplN As Double, subvExplN1 As Double
    subvExplN = -GetSoldeCompteN("74")
    subvExplN1 = -GetSoldeCompteN1("74")

    Dim autProdExplN As Double, autProdExplN1 As Double
    autProdExplN = -GetSoldeCompteN("75") - GetSoldeCompteN("781") - GetSoldeCompteN("791")
    autProdExplN1 = -GetSoldeCompteN1("75") - GetSoldeCompteN1("781") - GetSoldeCompteN1("791")

    Dim achatsN As Double, achatsN1 As Double
    achatsN = GetSoldeCompteN("60")
    achatsN1 = GetSoldeCompteN1("60")

    Dim servExtN As Double, servExtN1 As Double
    servExtN = GetSoldeCompteN("61") + GetSoldeCompteN("62")
    servExtN1 = GetSoldeCompteN1("61") + GetSoldeCompteN1("62")

    Dim impotsTaxesN As Double, impotsTaxesN1 As Double
    impotsTaxesN = GetSoldeCompteN("63")
    impotsTaxesN1 = GetSoldeCompteN1("63")

    Dim personnelN As Double, personnelN1 As Double
    personnelN = GetSoldeCompteN("64")
    personnelN1 = GetSoldeCompteN1("64")

    Dim dotN As Double, dotN1 As Double
    dotN = GetSoldeCompteN("681") + GetSoldeCompteN("686")
    dotN1 = GetSoldeCompteN1("681") + GetSoldeCompteN1("686")

    Dim autChargesN As Double, autChargesN1 As Double
    autChargesN = GetSoldeCompteN("65")
    autChargesN1 = GetSoldeCompteN1("65")

    Dim prodFinN As Double, prodFinN1 As Double
    prodFinN = -GetSoldeCompteN("76") - GetSoldeCompteN("786") - GetSoldeCompteN("796")
    prodFinN1 = -GetSoldeCompteN1("76") - GetSoldeCompteN1("786") - GetSoldeCompteN1("796")

    Dim chargesFinN As Double, chargesFinN1 As Double
    chargesFinN = GetSoldeCompteN("66") + GetSoldeCompteN("686") + GetSoldeCompteN("696")
    chargesFinN1 = GetSoldeCompteN1("66") + GetSoldeCompteN1("686") + GetSoldeCompteN1("696")

    Dim prodExcN As Double, prodExcN1 As Double
    prodExcN = -GetSoldeCompteN("77") - GetSoldeCompteN("787") - GetSoldeCompteN("797")
    prodExcN1 = -GetSoldeCompteN1("77") - GetSoldeCompteN1("787") - GetSoldeCompteN1("797")

    Dim chargesExcN As Double, chargesExcN1 As Double
    chargesExcN = GetSoldeCompteN("67") + GetSoldeCompteN("687") + GetSoldeCompteN("697")
    chargesExcN1 = GetSoldeCompteN1("67") + GetSoldeCompteN1("687") + GetSoldeCompteN1("697")

    Dim isN As Double, isN1 As Double
    isN = GetSoldeCompteN("695") + GetSoldeCompteN("691")
    isN1 = GetSoldeCompteN1("695") + GetSoldeCompteN1("691")

    Dim rexN As Double, rexN1 As Double
    rexN = caN + prodStockN + prodImmN + subvExplN + autProdExplN - achatsN - servExtN - impotsTaxesN - personnelN - dotN - autChargesN
    rexN1 = caN1 + prodStockN1 + prodImmN1 + subvExplN1 + autProdExplN1 - achatsN1 - servExtN1 - impotsTaxesN1 - personnelN1 - dotN1 - autChargesN1

    Dim rcaN As Double, rcaN1 As Double
    rcaN = rexN + prodFinN - chargesFinN
    rcaN1 = rexN1 + prodFinN1 - chargesFinN1

    Dim resExcN As Double, resExcN1 As Double
    resExcN = prodExcN - chargesExcN
    resExcN1 = prodExcN1 - chargesExcN1

    Dim resNetN As Double, resNetN1 As Double
    resNetN = rcaN + resExcN - isN
    resNetN1 = rcaN1 + resExcN1 - isN1

    ' Produits exploitation
    ws.Cells(ligne, 1).Value = "PRODUITS D'EXPLOITATION"
    With ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 7))
        .Interior.Color = RGB(0, 112, 0)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    ligne = ligne + 1

    Call EcrireLignePL(ws, ligne, "Chiffre d'affaires HT", "Note PL1", caN, caN1, caN) : ligne = ligne + 1
    Call EcrireLignePL(ws, ligne, "Production stockee (71)", "Note PL2", prodStockN, prodStockN1, caN) : ligne = ligne + 1
    Call EcrireLignePL(ws, ligne, "Production immobilisee (72)", "Note PL3", prodImmN, prodImmN1, caN) : ligne = ligne + 1
    Call EcrireLignePL(ws, ligne, "Subventions d'exploitation (74)", "Note PL4", subvExplN, subvExplN1, caN) : ligne = ligne + 1
    Call EcrireLignePL(ws, ligne, "Autres produits exploitation (75+781+791)", "", autProdExplN, autProdExplN1, caN) : ligne = ligne + 1

    Dim totProdExplN As Double, totProdExplN1 As Double
    totProdExplN = caN + prodStockN + prodImmN + subvExplN + autProdExplN
    totProdExplN1 = caN1 + prodStockN1 + prodImmN1 + subvExplN1 + autProdExplN1
    ws.Cells(ligne, 1).Value = "TOTAL PRODUITS EXPLOITATION"
    ws.Cells(ligne, 3).Value = totProdExplN
    ws.Cells(ligne, 4).Value = totProdExplN1
    ws.Cells(ligne, 5).Value = totProdExplN - totProdExplN1
    If totProdExplN1 <> 0 Then ws.Cells(ligne, 6).Value = (totProdExplN - totProdExplN1) / Abs(totProdExplN1)
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 7)).Interior.Color = RGB(198, 224, 180)
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 7)).Font.Bold = True
    ligne = ligne + 2

    ' Charges exploitation
    ws.Cells(ligne, 1).Value = "CHARGES D'EXPLOITATION"
    With ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 7))
        .Interior.Color = RGB(192, 80, 77)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    ligne = ligne + 1

    Call EcrireLignePL(ws, ligne, "Achats et variation stocks (60)", "Note PL5", achatsN, achatsN1, caN) : ligne = ligne + 1
    Call EcrireLignePL(ws, ligne, "Services exterieurs (61+62)", "Note PL6", servExtN, servExtN1, caN) : ligne = ligne + 1
    Call EcrireLignePL(ws, ligne, "Impots et taxes (63)", "Note PL7", impotsTaxesN, impotsTaxesN1, caN) : ligne = ligne + 1
    Call EcrireLignePL(ws, ligne, "Charges de personnel (64)", "Note PL8", personnelN, personnelN1, caN) : ligne = ligne + 1
    Call EcrireLignePL(ws, ligne, "Dotations amort. et prov. (681)", "Note PL9", dotN, dotN1, caN) : ligne = ligne + 1
    Call EcrireLignePL(ws, ligne, "Autres charges gestion courante (65)", "", autChargesN, autChargesN1, caN) : ligne = ligne + 1

    Dim totChargesExplN As Double, totChargesExplN1 As Double
    totChargesExplN = achatsN + servExtN + impotsTaxesN + personnelN + dotN + autChargesN
    totChargesExplN1 = achatsN1 + servExtN1 + impotsTaxesN1 + personnelN1 + dotN1 + autChargesN1
    ws.Cells(ligne, 1).Value = "TOTAL CHARGES EXPLOITATION"
    ws.Cells(ligne, 3).Value = totChargesExplN
    ws.Cells(ligne, 4).Value = totChargesExplN1
    ws.Cells(ligne, 5).Value = totChargesExplN - totChargesExplN1
    If totChargesExplN1 <> 0 Then ws.Cells(ligne, 6).Value = (totChargesExplN - totChargesExplN1) / Abs(totChargesExplN1)
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 7)).Interior.Color = RGB(255, 199, 206)
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 7)).Font.Bold = True
    ligne = ligne + 2

    ' REX
    ws.Cells(ligne, 1).Value = "RESULTAT D'EXPLOITATION (REX)"
    ws.Cells(ligne, 3).Value = rexN
    ws.Cells(ligne, 4).Value = rexN1
    ws.Cells(ligne, 5).Value = rexN - rexN1
    If rexN1 <> 0 Then ws.Cells(ligne, 6).Value = (rexN - rexN1) / Abs(rexN1)
    If caN <> 0 Then ws.Cells(ligne, 7).Value = rexN / caN
    With ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 7))
        .Interior.Color = RGB(0, 112, 0)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 11
    End With
    ligne = ligne + 2

    Call EcrireLignePL(ws, ligne, "Produits financiers (76+786+796)", "", prodFinN, prodFinN1, caN) : ligne = ligne + 1
    Call EcrireLignePL(ws, ligne, "Charges financieres (66+686+696)", "", chargesFinN, chargesFinN1, caN) : ligne = ligne + 1

    ws.Cells(ligne, 1).Value = "RESULTAT FINANCIER"
    ws.Cells(ligne, 3).Value = prodFinN - chargesFinN
    ws.Cells(ligne, 4).Value = prodFinN1 - chargesFinN1
    ws.Cells(ligne, 5).Value = (prodFinN - chargesFinN) - (prodFinN1 - chargesFinN1)
    If caN <> 0 Then ws.Cells(ligne, 7).Value = (prodFinN - chargesFinN) / caN
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 7)).Interior.Color = RGB(189, 215, 238)
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 7)).Font.Bold = True
    ligne = ligne + 2

    ' RCA
    ws.Cells(ligne, 1).Value = "RESULTAT COURANT AVANT IS (RCA)"
    ws.Cells(ligne, 3).Value = rcaN
    ws.Cells(ligne, 4).Value = rcaN1
    ws.Cells(ligne, 5).Value = rcaN - rcaN1
    If rcaN1 <> 0 Then ws.Cells(ligne, 6).Value = (rcaN - rcaN1) / Abs(rcaN1)
    If caN <> 0 Then ws.Cells(ligne, 7).Value = rcaN / caN
    With ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 7))
        .Interior.Color = RGB(0, 112, 0)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 11
    End With
    ligne = ligne + 2

    Call EcrireLignePL(ws, ligne, "Produits exceptionnels (77+787+797)", "", prodExcN, prodExcN1, caN) : ligne = ligne + 1
    Call EcrireLignePL(ws, ligne, "Charges exceptionnelles (67+687+697)", "", chargesExcN, chargesExcN1, caN) : ligne = ligne + 1

    ws.Cells(ligne, 1).Value = "RESULTAT EXCEPTIONNEL"
    ws.Cells(ligne, 3).Value = resExcN
    ws.Cells(ligne, 4).Value = resExcN1
    ws.Cells(ligne, 5).Value = resExcN - resExcN1
    If caN <> 0 Then ws.Cells(ligne, 7).Value = resExcN / caN
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 7)).Interior.Color = RGB(189, 215, 238)
    ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 7)).Font.Bold = True
    ligne = ligne + 2

    Call EcrireLignePL(ws, ligne, "Participation + IS (69x)", "", isN, isN1, caN) : ligne = ligne + 2

    ws.Cells(ligne, 1).Value = "RESULTAT NET DE L'EXERCICE"
    ws.Cells(ligne, 3).Value = resNetN
    ws.Cells(ligne, 4).Value = resNetN1
    ws.Cells(ligne, 5).Value = resNetN - resNetN1
    If resNetN1 <> 0 Then ws.Cells(ligne, 6).Value = (resNetN - resNetN1) / Abs(resNetN1)
    If caN <> 0 Then ws.Cells(ligne, 7).Value = resNetN / caN
    With ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 7))
        .Interior.Color = RGB(0, 70, 0)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 13
    End With

    ws.Columns("A").ColumnWidth = 50
    ws.Columns("B").ColumnWidth = 12
    ws.Columns("C:G").ColumnWidth = 16
    ws.Range("C4:F500").NumberFormat = "#,##0;[Red]-#,##0"
    ws.Range("F4:F500").NumberFormat = "0.0%"
    ws.Range("G4:G500").NumberFormat = "0.0%"
End Sub

' ============================================================
' SIG
' ============================================================
Sub CreerOngletSIG()
    Dim ws As Worksheet
    Dim ligne As Long

    Set ws = CreerOnglet("SIG", COULEUR_RATIO)

    With ws.Range("A1:H1")
        .Merge
        .Value = "SOLDES INTERMEDIAIRES DE GESTION (SIG)"
        .Interior.Color = RGB(197, 90, 17)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
    End With
    With ws.Range("A2:H2")
        .Merge
        .Value = "Analyse de la formation du resultat - exercices N et N-1"
        .Interior.Color = RGB(255, 153, 0)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With

    ws.Cells(3, 1).Value = "SOLDE INTERMEDIAIRE"
    ws.Cells(3, 2).Value = "Composantes"
    ws.Cells(3, 3).Value = "Montant N"
    ws.Cells(3, 4).Value = "Montant N-1"
    ws.Cells(3, 5).Value = "Variation"
    ws.Cells(3, 6).Value = "Var. %"
    ws.Cells(3, 7).Value = "% CA N"
    ws.Cells(3, 8).Value = "% CA N-1"
    With ws.Range("A3:H3")
        .Font.Bold = True
        .Interior.Color = RGB(253, 233, 217)
    End With

    ligne = 4

    Dim caN As Double, caN1 As Double
    caN = -GetSoldeCompteN("70")
    caN1 = -GetSoldeCompteN1("70")

    Dim venteMarcN As Double, venteMarcN1 As Double
    Dim coutMarcN As Double, coutMarcN1 As Double
    venteMarcN = -GetSoldeCompteN("707")
    venteMarcN1 = -GetSoldeCompteN1("707")
    coutMarcN = GetSoldeCompteN("607") + GetSoldeCompteN("6037")
    coutMarcN1 = GetSoldeCompteN1("607") + GetSoldeCompteN1("6037")
    Dim margeComN As Double, margeComN1 As Double
    margeComN = venteMarcN - coutMarcN
    margeComN1 = venteMarcN1 - coutMarcN1

    Dim prodN As Double, prodN1 As Double
    prodN = caN + (-GetSoldeCompteN("71")) + (-GetSoldeCompteN("72"))
    prodN1 = caN1 + (-GetSoldeCompteN1("71")) + (-GetSoldeCompteN1("72"))

    Dim vaH As Double, vaH1 As Double
    vaH = prodN - GetSoldeCompteN("60") - GetSoldeCompteN("61") - GetSoldeCompteN("62")
    vaH1 = prodN1 - GetSoldeCompteN1("60") - GetSoldeCompteN1("61") - GetSoldeCompteN1("62")

    Dim subvExplN As Double, subvExplN1 As Double
    subvExplN = -GetSoldeCompteN("74")
    subvExplN1 = -GetSoldeCompteN1("74")

    Dim personnelN As Double, personnelN1 As Double
    personnelN = GetSoldeCompteN("64")
    personnelN1 = GetSoldeCompteN1("64")

    Dim impotsTaxesN As Double, impotsTaxesN1 As Double
    impotsTaxesN = GetSoldeCompteN("63")
    impotsTaxesN1 = GetSoldeCompteN1("63")

    Dim ebeN As Double, ebeN1 As Double
    ebeN = vaH + subvExplN - impotsTaxesN - personnelN
    ebeN1 = vaH1 + subvExplN1 - impotsTaxesN1 - personnelN1

    Dim dotN As Double, dotN1 As Double
    dotN = GetSoldeCompteN("681")
    dotN1 = GetSoldeCompteN1("681")

    Dim rexN As Double, rexN1 As Double
    rexN = ebeN - dotN + (-GetSoldeCompteN("75")) + (-GetSoldeCompteN("781")) + (-GetSoldeCompteN("791")) - GetSoldeCompteN("65")
    rexN1 = ebeN1 - dotN1 + (-GetSoldeCompteN1("75")) + (-GetSoldeCompteN1("781")) + (-GetSoldeCompteN1("791")) - GetSoldeCompteN1("65")

    Dim prodFinN As Double, prodFinN1 As Double
    Dim chargesFinN As Double, chargesFinN1 As Double
    prodFinN = -GetSoldeCompteN("76") - GetSoldeCompteN("786") - GetSoldeCompteN("796")
    prodFinN1 = -GetSoldeCompteN1("76") - GetSoldeCompteN1("786") - GetSoldeCompteN1("796")
    chargesFinN = GetSoldeCompteN("66") + GetSoldeCompteN("686") + GetSoldeCompteN("696")
    chargesFinN1 = GetSoldeCompteN1("66") + GetSoldeCompteN1("686") + GetSoldeCompteN1("696")

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

    Call EcrireLigneSIG(ws, ligne, "1. MARGE COMMERCIALE", "Ventes marchandises - Cout achats marchandises", margeComN, margeComN1, caN, caN1) : ligne = ligne + 2
    Call EcrireLigneSIG(ws, ligne, "2. PRODUCTION DE L'EXERCICE", "CA + Prod. stockee + Prod. immobilisee", prodN, prodN1, caN, caN1) : ligne = ligne + 2
    Call EcrireLigneSIG(ws, ligne, "3. VALEUR AJOUTEE (VA)", "Prod. exercice - Consommations intermediaires", vaH, vaH1, caN, caN1) : ligne = ligne + 2
    Call EcrireLigneSIG(ws, ligne, "4. EXCEDENT BRUT D'EXPLOITATION (EBE)", "VA + Subv. expl. - Impots taxes - Charges pers.", ebeN, ebeN1, caN, caN1) : ligne = ligne + 2
    Call EcrireLigneSIG(ws, ligne, "5. RESULTAT D'EXPLOITATION (REX)", "EBE - Dotations + Reprises + Autres prod./charges", rexN, rexN1, caN, caN1) : ligne = ligne + 2
    Call EcrireLigneSIG(ws, ligne, "6. RESULTAT COURANT AVANT IS (RCA)", "REX + Resultat financier", rcaN, rcaN1, caN, caN1) : ligne = ligne + 2
    Call EcrireLigneSIG(ws, ligne, "7. RESULTAT EXCEPTIONNEL", "Produits except. - Charges except.", resExcN, resExcN1, caN, caN1) : ligne = ligne + 2
    Call EcrireLigneSIG(ws, ligne, "8. RESULTAT NET DE L'EXERCICE", "RCA + Res. except. - Participation - IS", resNetN, resNetN1, caN, caN1) : ligne = ligne + 3

    ' Ratios
    ws.Cells(ligne, 1).Value = "RATIOS DE RENTABILITE"
    With ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 8))
        .Interior.Color = RGB(197, 90, 17)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    ligne = ligne + 1

    Dim immoBrutN As Double, immoBrutN1 As Double
    immoBrutN = GetSoldeCompteN("20") + GetSoldeCompteN("21") + GetSoldeCompteN("22") + GetSoldeCompteN("23")
    immoBrutN1 = GetSoldeCompteN1("20") + GetSoldeCompteN1("21") + GetSoldeCompteN1("22") + GetSoldeCompteN1("23")

    Dim rAvtISN As Double, rAvtISN1 As Double
    rAvtISN = rcaN + resExcN
    rAvtISN1 = rcaN1 + resExcN1

    Call EcrireRatioSIG(ws, ligne, "Taux de marge commerciale", IIf(caN <> 0, margeComN / caN, 0), IIf(caN1 <> 0, margeComN1 / caN1, 0)) : ligne = ligne + 1
    Call EcrireRatioSIG(ws, ligne, "Taux de VA / CA", IIf(caN <> 0, vaH / caN, 0), IIf(caN1 <> 0, vaH1 / caN1, 0)) : ligne = ligne + 1
    Call EcrireRatioSIG(ws, ligne, "Taux d'EBE / CA (EBITDA margin)", IIf(caN <> 0, ebeN / caN, 0), IIf(caN1 <> 0, ebeN1 / caN1, 0)) : ligne = ligne + 1
    Call EcrireRatioSIG(ws, ligne, "Taux de REX / CA (EBIT margin)", IIf(caN <> 0, rexN / caN, 0), IIf(caN1 <> 0, rexN1 / caN1, 0)) : ligne = ligne + 1
    Call EcrireRatioSIG(ws, ligne, "Taux de RCA / CA", IIf(caN <> 0, rcaN / caN, 0), IIf(caN1 <> 0, rcaN1 / caN1, 0)) : ligne = ligne + 1
    Call EcrireRatioSIG(ws, ligne, "Marge nette / CA", IIf(caN <> 0, resNetN / caN, 0), IIf(caN1 <> 0, resNetN1 / caN1, 0)) : ligne = ligne + 1
    Call EcrireRatioSIG(ws, ligne, "Personnel / VA", IIf(vaH <> 0, personnelN / vaH, 0), IIf(vaH1 <> 0, personnelN1 / vaH1, 0)) : ligne = ligne + 1
    Call EcrireRatioSIG(ws, ligne, "Dotations / Immo. brutes", IIf(immoBrutN <> 0, dotN / immoBrutN, 0), IIf(immoBrutN1 <> 0, dotN1 / immoBrutN1, 0)) : ligne = ligne + 1
    Call EcrireRatioSIG(ws, ligne, "Charges financieres / EBE", IIf(ebeN <> 0, chargesFinN / ebeN, 0), IIf(ebeN1 <> 0, chargesFinN1 / ebeN1, 0)) : ligne = ligne + 1
    Call EcrireRatioSIG(ws, ligne, "IS effectif / Resultat avant IS", IIf(rAvtISN <> 0, isN / rAvtISN, 0), IIf(rAvtISN1 <> 0, isN1 / rAvtISN1, 0)) : ligne = ligne + 1

    ws.Columns("A").ColumnWidth = 50
    ws.Columns("B").ColumnWidth = 55
    ws.Columns("C:H").ColumnWidth = 16
End Sub

Sub EcrireLigneSIG(ws As Worksheet, ligne As Long, Libelle As String, Composantes As String, ValN As Double, ValN1 As Double, CaN As Double, CaN1 As Double)
    With ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 8))
        .Interior.Color = RGB(253, 233, 217)
        .Font.Bold = True
        .Borders(xlEdgeTop).LineStyle = xlContinuous
    End With
    ws.Cells(ligne, 1).Value = Libelle
    ws.Cells(ligne, 2).Value = Composantes
    ws.Cells(ligne, 3).Value = ValN
    ws.Cells(ligne, 4).Value = ValN1
    ws.Cells(ligne, 5).Value = ValN - ValN1
    If ValN1 <> 0 Then ws.Cells(ligne, 6).Value = (ValN - ValN1) / Abs(ValN1)
    If CaN <> 0 Then ws.Cells(ligne, 7).Value = ValN / CaN
    If CaN1 <> 0 Then ws.Cells(ligne, 8).Value = ValN1 / CaN1
    ws.Cells(ligne, 3).NumberFormat = "#,##0;[Red]-#,##0"
    ws.Cells(ligne, 4).NumberFormat = "#,##0;[Red]-#,##0"
    ws.Cells(ligne, 5).NumberFormat = "#,##0;[Red]-#,##0"
    ws.Cells(ligne, 6).NumberFormat = "0.0%"
    ws.Cells(ligne, 7).NumberFormat = "0.0%"
    ws.Cells(ligne, 8).NumberFormat = "0.0%"
End Sub

Sub EcrireRatioSIG(ws As Worksheet, ligne As Long, Libelle As String, ValN As Double, ValN1 As Double)
    ws.Cells(ligne, 1).Value = Libelle
    ws.Cells(ligne, 3).Value = ValN
    ws.Cells(ligne, 4).Value = ValN1
    ws.Cells(ligne, 5).Value = ValN - ValN1
    ws.Cells(ligne, 3).NumberFormat = "0.0%"
    ws.Cells(ligne, 4).NumberFormat = "0.0%"
    ws.Cells(ligne, 5).NumberFormat = "0.0%"
    If ligne Mod 2 = 0 Then ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 8)).Interior.Color = RGB(253, 233, 217)
End Sub
