Attribute VB_Name = "INSTALLER_MACRO"
Option Explicit

' ============================================================
' MODULE D'INSTALLATION
' ============================================================

Sub CreerOngletSourceBalance()
    Dim wsN As Worksheet, wsN1 As Worksheet

    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets("Balance_N").Delete
    ThisWorkbook.Worksheets("Balance_N-1").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set wsN = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
    wsN.Name = "Balance_N"
    wsN.Tab.Color = RGB(0, 112, 192)

    Set wsN1 = ThisWorkbook.Worksheets.Add(After:=wsN)
    wsN1.Name = "Balance_N-1"
    wsN1.Tab.Color = RGB(0, 176, 240)

    Dim ws As Worksheet
    For Each ws In Array(wsN, wsN1)
        ws.Cells(1, 1).Value = "Numero compte"
        ws.Cells(1, 2).Value = "Intitule compte"
        ws.Cells(1, 3).Value = "Debit"
        ws.Cells(1, 4).Value = "Credit"
        With ws.Range("A1:D1")
            .Font.Bold = True
            .Interior.Color = RGB(31, 73, 125)
            .Font.Color = RGB(255, 255, 255)
        End With
        ws.Columns("A").ColumnWidth = 15
        ws.Columns("B").ColumnWidth = 50
        ws.Columns("C:D").ColumnWidth = 18
        ws.Columns("C:D").NumberFormat = "#,##0.00"
    Next ws

    Call InsererExempleBalance(wsN, True)
    Call InsererExempleBalance(wsN1, False)

    MsgBox "Onglets Balance_N et Balance_N-1 crees avec des donnees exemple." & vbCrLf & _
           "Remplacez les donnees exemple par votre balance reelle.", vbInformation
End Sub

Sub InsererExempleBalance(ws As Worksheet, estN As Boolean)
    Dim coeff As Double
    If estN Then coeff = 1 Else coeff = 0.92

    Dim ligne As Long
    ligne = 2

    ' Helper pour ajouter une ligne
    ' Format : compte, intitule, debit, credit

    ' --- PASSIF : Capitaux propres et dettes ---
    Call AjLigne(ws, ligne, "101000", "Capital social", 0, 500000 * coeff) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "104000", "Prime d'emission", 0, 100000 * coeff) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "106100", "Reserve legale", 0, 50000 * coeff) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "106800", "Autres reserves", 0, 200000 * coeff) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "120000", "Resultat de l'exercice", 0, 75000 * coeff) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "130000", "Subventions d'investissement", 0, 30000 * coeff) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "151000", "Provisions pour litiges", 0, 15000 * coeff) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "152000", "Provisions pour garanties", 0, 8000 * coeff) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "164000", "Emprunts aupres etab. credit", 0, 350000 * coeff) : ligne = ligne + 1

    ' --- ACTIF IMMOBILISE ---
    Call AjLigne(ws, ligne, "201000", "Frais etablissement", 5000 * coeff, 0) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "203000", "Frais de R&D", 20000 * coeff, 0) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "207000", "Fonds commercial", 150000 * coeff, 0) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "211000", "Terrains", 200000 * coeff, 0) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "213000", "Constructions", 800000 * coeff, 0) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "215000", "Installations techniques", 500000 * coeff, 0) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "218000", "Autres immob. corporelles", 100000 * coeff, 0) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "280300", "Amort. frais R&D", 0, 10000 * coeff) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "281300", "Amort. constructions", 0, 250000 * coeff) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "281500", "Amort. installations tech.", 0, 200000 * coeff) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "261000", "Titres de participation", 80000 * coeff, 0) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "274000", "Prets", 25000 * coeff, 0) : ligne = ligne + 1

    ' --- ACTIF CIRCULANT ---
    Call AjLigne(ws, ligne, "310000", "Matieres premieres", 120000 * coeff, 0) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "350000", "Produits finis", 80000 * coeff, 0) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "370000", "Marchandises", 50000 * coeff, 0) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "411000", "Clients", 850000 * coeff, 0) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "416000", "Clients douteux", 30000 * coeff, 0) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "444000", "Etat - IS", 12000 * coeff, 0) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "445660", "TVA deductible", 25000 * coeff, 0) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "490000", "Depreciations clients", 0, 15000 * coeff) : ligne = ligne + 1

    ' --- PASSIF CIRCULANT ---
    Call AjLigne(ws, ligne, "401000", "Fournisseurs", 0, 320000 * coeff) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "408000", "Fourn. - fact. non parvenues", 0, 45000 * coeff) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "421000", "Personnel - remunerations", 0, 85000 * coeff) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "431000", "SS - URSSAF", 0, 42000 * coeff) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "437000", "Autres org. sociaux", 0, 8000 * coeff) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "443000", "TVA a decaisser", 0, 65000 * coeff) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "512000", "Banque", 180000 * coeff, 0) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "530000", "Caisse", 5000 * coeff, 0) : ligne = ligne + 1

    ' --- CHARGES ---
    Call AjLigne(ws, ligne, "607000", "Achats de marchandises", 1200000 * coeff, 0) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "603700", "Variation stocks marchandises", 10000 * coeff, 0) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "601000", "Achats matieres premieres", 800000 * coeff, 0) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "603100", "Variation stocks MP", 0, 5000 * coeff) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "611000", "Sous-traitance", 150000 * coeff, 0) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "613000", "Locations", 48000 * coeff, 0) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "615000", "Entretiens et reparations", 35000 * coeff, 0) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "616000", "Primes d'assurances", 12000 * coeff, 0) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "621000", "Personnel interimaire", 80000 * coeff, 0) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "622000", "Remunerations intermediaires", 25000 * coeff, 0) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "623000", "Publicite", 30000 * coeff, 0) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "625000", "Deplacements", 15000 * coeff, 0) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "626000", "Frais postaux et telecoms", 8000 * coeff, 0) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "627000", "Services bancaires", 5000 * coeff, 0) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "631000", "Taxe apprentissage", 6000 * coeff, 0) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "635000", "Autres impots et taxes", 18000 * coeff, 0) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "641000", "Salaires bruts", 480000 * coeff, 0) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "645000", "Cotisations SS employeur", 195000 * coeff, 0) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "661000", "Charges d'interets", 18000 * coeff, 0) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "665000", "Escomptes accordes", 3000 * coeff, 0) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "681100", "Dot. amort. immo. incorp.", 2500 * coeff, 0) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "681300", "Dot. amort. constructions", 20000 * coeff, 0) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "681500", "Dot. amort. install. tech.", 50000 * coeff, 0) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "695000", "IS courant", 28000 * coeff, 0) : ligne = ligne + 1

    ' --- PRODUITS ---
    Call AjLigne(ws, ligne, "707000", "Ventes de marchandises", 0, 1800000 * coeff) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "701000", "Ventes produits finis", 0, 1200000 * coeff) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "706000", "Prestations de services", 0, 500000 * coeff) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "709000", "RRR accordes", 15000 * coeff, 0) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "740000", "Subventions d'exploitation", 0, 20000 * coeff) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "751000", "Redevances", 0, 8000 * coeff) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "761000", "Produits des participations", 0, 4000 * coeff) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "765000", "Escomptes obtenus", 0, 2000 * coeff) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "771000", "Produits except. sur oper.", 0, 5000 * coeff) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "775000", "Produits cession actif", 0, 12000 * coeff) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "675000", "VNC elements cedes", 8000 * coeff, 0) : ligne = ligne + 1
    Call AjLigne(ws, ligne, "781000", "Reprises prov. exploitation", 0, 3000 * coeff) : ligne = ligne + 1
End Sub

' Helper : ecrire une ligne de balance
Sub AjLigne(ws As Worksheet, ligne As Long, compte As String, intitule As String, debit As Double, credit As Double)
    ws.Cells(ligne, 1).Value = compte
    ws.Cells(ligne, 2).Value = intitule
    ws.Cells(ligne, 3).Value = debit
    ws.Cells(ligne, 4).Value = credit
End Sub

' ============================================================
Sub AfficherFormatAttendu()
    MsgBox "FORMAT ATTENDU POUR LES ONGLETS BALANCE :" & vbCrLf & vbCrLf & _
           "Onglets requis :" & vbCrLf & _
           "  - 'Balance_N'   : Balance exercice en cours" & vbCrLf & _
           "  - 'Balance_N-1' : Balance exercice precedent" & vbCrLf & vbCrLf & _
           "Colonnes (ligne 1 = en-tete) :" & vbCrLf & _
           "  - Colonne A : Numero de compte (ex: 411000)" & vbCrLf & _
           "  - Colonne B : Intitule du compte" & vbCrLf & _
           "  - Colonne C : Total Debit de l'exercice" & vbCrLf & _
           "  - Colonne D : Total Credit de l'exercice" & vbCrLf & vbCrLf & _
           "Notes :" & vbCrLf & _
           "  - Solde = Debit - Credit (calcule auto)" & vbCrLf & _
           "  - Balance AVANT affectation du resultat" & vbCrLf & _
           "  - Compatible : Sage, Cegid, EBP, Quadratus", _
           vbInformation, "Format Balance Comptable"
End Sub

' ============================================================
Sub DemarrerAnalyse()
    Dim reponse As Integer
    reponse = MsgBox("ANALYSE FINANCIERE COMPARATIVE" & vbCrLf & vbCrLf & _
                     "OUI     : Creer des onglets avec donnees exemple" & vbCrLf & _
                     "NON     : Lancer l'analyse sur vos balances existantes" & vbCrLf & _
                     "ANNULER : Voir le format attendu", _
                     vbYesNoCancel + vbQuestion, "Demarrage")

    Select Case reponse
        Case vbYes
            Call CreerOngletSourceBalance
        Case vbNo
            Call LancerAnalyseFinanciere
        Case vbCancel
            Call AfficherFormatAttendu
    End Select
End Sub
