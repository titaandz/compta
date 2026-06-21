Attribute VB_Name = "INSTALLER_MACRO"
Option Explicit

' ============================================================
' MODULE D'INSTALLATION - A coller en premier dans VBE
' Ce module contient également les helpers pour créer les
' onglets source (balance) et un exemple de données.
' ============================================================

' ============================================================
' CREER LES ONGLETS SOURCE AVEC LE FORMAT ATTENDU
' ============================================================
Sub CreerOngletSourceBalance()
    ' Ce sous-programme crée les onglets Balance_N et Balance_N-1
    ' avec les en-têtes attendus et un exemple de données

    Dim wsN As Worksheet, wsN1 As Worksheet

    ' Supprimer si existants
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets("Balance_N").Delete
    ThisWorkbook.Worksheets("Balance_N-1").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    ' Créer les onglets
    Set wsN = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
    wsN.Name = "Balance_N"
    wsN.Tab.Color = RGB(0, 112, 192)

    Set wsN1 = ThisWorkbook.Worksheets.Add(After:=wsN)
    wsN1.Name = "Balance_N-1"
    wsN1.Tab.Color = RGB(0, 176, 240)

    ' En-têtes (ligne 1)
    Dim enTetes() As String
    enTetes = Split("Numero compte,Intitule compte,Debit,Credit", ",")

    Dim ws As Worksheet
    For Each ws In Array(wsN, wsN1)
        Dim c As Integer
        For c = 0 To UBound(enTetes)
            ws.Cells(1, c + 1).Value = enTetes(c)
        Next c
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

    ' Données exemple pour Balance_N
    Call InsererExempleBalance(wsN, True)

    ' Données exemple pour Balance_N-1
    Call InsererExempleBalance(wsN1, False)

    MsgBox "Onglets Balance_N et Balance_N-1 créés avec des données exemple." & vbCrLf & _
           "Remplacez les données exemple par votre balance réelle.", vbInformation
End Sub

Sub InsererExempleBalance(ws As Worksheet, estN As Boolean)
    ' Insère des données exemple représentatives d'une PME
    Dim coeff As Double
    If estN Then coeff = 1 Else coeff = 0.92  ' N-1 = 92% de N (croissance de 8%)

    Dim donnees() As Variant
    ' Format : N°Compte, Intitulé, Débit, Crédit
    ' Solde débiteur = Débit > Crédit (actif, charges)
    ' Solde créditeur = Crédit > Débit (passif, produits)
    donnees = Array( _
        Array("101000", "Capital social", 0, 500000 * coeff), _
        Array("104000", "Prime d'émission", 0, 100000 * coeff), _
        Array("106100", "Reserve légale", 0, 50000 * coeff), _
        Array("106800", "Autres réserves", 0, 200000 * coeff), _
        Array("120000", "Résultat de l'exercice", 0, 75000 * coeff), _
        Array("130000", "Subventions d'investissement", 0, 30000 * coeff), _
        Array("151000", "Provisions pour litiges", 0, 15000 * coeff), _
        Array("152000", "Provisions pour garanties", 0, 8000 * coeff), _
        Array("164000", "Emprunts auprès étab. crédit", 0, 350000 * coeff), _
        Array("201000", "Frais établissement", 5000 * coeff, 0), _
        Array("203000", "Frais de R&D", 20000 * coeff, 0), _
        Array("207000", "Fonds commercial", 150000 * coeff, 0), _
        Array("211000", "Terrains", 200000 * coeff, 0), _
        Array("213000", "Constructions", 800000 * coeff, 0), _
        Array("215000", "Installations techniques", 500000 * coeff, 0), _
        Array("218000", "Autres immob. corporelles", 100000 * coeff, 0), _
        Array("280300", "Amort. frais R&D", 0, 10000 * coeff), _
        Array("280700", "Amort. fonds commercial", 0, 0), _
        Array("281300", "Amort. constructions", 0, 250000 * coeff), _
        Array("281500", "Amort. installations tech.", 0, 200000 * coeff), _
        Array("261000", "Titres de participation", 80000 * coeff, 0), _
        Array("274000", "Prêts", 25000 * coeff, 0), _
        Array("310000", "Matières premières", 120000 * coeff, 0), _
        Array("350000", "Produits finis", 80000 * coeff, 0), _
        Array("370000", "Marchandises", 50000 * coeff, 0), _
        Array("411000", "Clients", 850000 * coeff, 0), _
        Array("416000", "Clients douteux", 30000 * coeff, 0), _
        Array("444000", "Etat - IS", 12000 * coeff, 0), _
        Array("445660", "TVA déductible", 25000 * coeff, 0), _
        Array("490000", "Dépréciations clients", 0, 15000 * coeff), _
        Array("401000", "Fournisseurs", 0, 320000 * coeff), _
        Array("408000", "Fourn. - fact. non parvenues", 0, 45000 * coeff), _
        Array("421000", "Personnel - rémunérations", 0, 85000 * coeff), _
        Array("431000", "SS - URSSAF", 0, 42000 * coeff), _
        Array("437000", "Autres org. sociaux", 0, 8000 * coeff), _
        Array("443000", "TVA à décaisser", 0, 65000 * coeff), _
        Array("444000", "IS à payer", 0, 18000 * coeff), _
        Array("512000", "Banque", 180000 * coeff, 0), _
        Array("530000", "Caisse", 5000 * coeff, 0), _
        Array("607000", "Achats de marchandises", 1200000 * coeff, 0), _
        Array("603700", "Variation stocks marchandises", 10000 * coeff, 0), _
        Array("601000", "Achats matières premières", 800000 * coeff, 0), _
        Array("603100", "Variation stocks MP", 0, 5000 * coeff), _
        Array("611000", "Sous-traitance", 150000 * coeff, 0), _
        Array("613000", "Locations", 48000 * coeff, 0), _
        Array("615000", "Entretiens et réparations", 35000 * coeff, 0), _
        Array("616000", "Primes d'assurances", 12000 * coeff, 0), _
        Array("621000", "Personnel intérimaire", 80000 * coeff, 0), _
        Array("622000", "Rémunérations intermédiaires", 25000 * coeff, 0), _
        Array("623000", "Publicité", 30000 * coeff, 0), _
        Array("625000", "Déplacements", 15000 * coeff, 0), _
        Array("626000", "Frais postaux et télécoms", 8000 * coeff, 0), _
        Array("627000", "Services bancaires", 5000 * coeff, 0), _
        Array("631000", "Taxe apprentissage", 6000 * coeff, 0), _
        Array("635000", "Autres impôts et taxes", 18000 * coeff, 0), _
        Array("641000", "Salaires bruts", 480000 * coeff, 0), _
        Array("645000", "Cotisations SS employeur", 195000 * coeff, 0), _
        Array("661000", "Charges d'intérêts", 18000 * coeff, 0), _
        Array("665000", "Escomptes accordés", 3000 * coeff, 0), _
        Array("681100", "Dot. amort. immo. incorp.", 2500 * coeff, 0), _
        Array("681300", "Dot. amort. constructions", 20000 * coeff, 0), _
        Array("681500", "Dot. amort. install. tech.", 50000 * coeff, 0), _
        Array("686800", "Dot. amort. sur emprunts", 1000 * coeff, 0), _
        Array("695000", "IS courant", 28000 * coeff, 0), _
        Array("707000", "Ventes de marchandises", 0, 1800000 * coeff), _
        Array("701000", "Ventes produits finis", 0, 1200000 * coeff), _
        Array("706000", "Prestations de services", 0, 500000 * coeff), _
        Array("709000", "RRR accordés", 15000 * coeff, 0), _
        Array("740000", "Subventions d'exploitation", 0, 20000 * coeff), _
        Array("751000", "Redevances", 0, 8000 * coeff), _
        Array("761000", "Produits des participations", 0, 4000 * coeff), _
        Array("765000", "Escomptes obtenus", 0, 2000 * coeff), _
        Array("771000", "Produits except. sur opér.", 0, 5000 * coeff), _
        Array("775000", "Produits cession actif", 0, 12000 * coeff), _
        Array("675000", "VNC éléments cédés", 8000 * coeff, 0), _
        Array("781000", "Reprises prov. exploitation", 0, 3000 * coeff) _
    )

    Dim i As Integer
    For i = 0 To UBound(donnees)
        ws.Cells(i + 2, 1).Value = donnees(i)(0)
        ws.Cells(i + 2, 2).Value = donnees(i)(1)
        ws.Cells(i + 2, 3).Value = donnees(i)(2)
        ws.Cells(i + 2, 4).Value = donnees(i)(3)
    Next i
End Sub

' ============================================================
' FORMAT ATTENDU POUR VOTRE BALANCE
' ============================================================
Sub AfficherFormatAttendu()
    MsgBox "FORMAT ATTENDU POUR LES ONGLETS BALANCE :" & vbCrLf & vbCrLf & _
           "Onglets requis :" & vbCrLf & _
           "  - 'Balance_N'  : Balance de l'exercice en cours" & vbCrLf & _
           "  - 'Balance_N-1': Balance de l'exercice précédent" & vbCrLf & vbCrLf & _
           "Colonnes (ligne 1 = en-tête) :" & vbCrLf & _
           "  - Colonne A : Numéro de compte (ex: 411000)" & vbCrLf & _
           "  - Colonne B : Intitulé du compte" & vbCrLf & _
           "  - Colonne C : Total Débit de l'exercice" & vbCrLf & _
           "  - Colonne D : Total Crédit de l'exercice" & vbCrLf & vbCrLf & _
           "Notes importantes :" & vbCrLf & _
           "  - Le solde est calculé automatiquement (Débit - Crédit)" & vbCrLf & _
           "  - Les numéros de compte doivent être en texte ou nombre" & vbCrLf & _
           "  - La balance doit être AVANT affectation du résultat" & vbCrLf & _
           "  - Format PCG : comptes à 6 chiffres recommandé" & vbCrLf & vbCrLf & _
           "Compatibilité logiciels :" & vbCrLf & _
           "  Sage, Cegid, EBP, Quadratus, Loop, QuickBooks France", _
           vbInformation, "Format Balance Comptable"
End Sub

' ============================================================
' PROCEDURE COMPLETE POUR DEMARRER
' ============================================================
Sub DemarrerAnalyse()
    Dim reponse As Integer
    reponse = MsgBox("ANALYSE FINANCIERE COMPARATIVE" & vbCrLf & vbCrLf & _
                     "Que souhaitez-vous faire ?" & vbCrLf & vbCrLf & _
                     "OUI  : Créer des onglets exemple (données fictives)" & vbCrLf & _
                     "NON  : Lancer l'analyse sur vos balances existantes" & vbCrLf & _
                     "ANNULER : Voir le format attendu", _
                     vbYesNoCancel + vbQuestion, "Démarrage")

    Select Case reponse
        Case vbYes
            Call CreerOngletSourceBalance
            MsgBox "Onglets créés. Vous pouvez maintenant :" & vbCrLf & _
                   "1. Remplacer les données par votre balance réelle" & vbCrLf & _
                   "2. Relancer cette macro et choisir NON pour lancer l'analyse", vbInformation
        Case vbNo
            Call LancerAnalyseFinanciere
        Case vbCancel
            Call AfficherFormatAttendu
    End Select
End Sub
