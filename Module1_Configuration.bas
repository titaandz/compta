Attribute VB_Name = "Module1_Configuration"
Option Explicit

' ============================================================
' CONFIGURATION GLOBALE - Analyse Financiere Comparee
' Normes ANC (Plan Comptable General)
' ============================================================

' Constantes de configuration
Public Const NOM_ONGLET_N As String = "Balance_N"
Public Const NOM_ONGLET_N1 As String = "Balance_N-1"
Public Const NOM_EXERCICE_N As String = "N"
Public Const NOM_EXERCICE_N1 As String = "N-1"

' Couleurs des onglets
Public Const COULEUR_BILAN As Long = 4953346      ' Bleu fonce
Public Const COULEUR_PL As Long = 5287936          ' Vert fonce
Public Const COULEUR_RATIO As Long = 10053120      ' Orange
Public Const COULEUR_RECAP As Long = 8421376       ' Gris
Public Const COULEUR_PPT As Long = 12611584        ' Rouge fonce
Public Const COULEUR_ANC As Long = 6697728         ' Violet

' Structure d'une ligne de balance
Public Type LigneBalance
    NumCompte As String
    IntituleCompte As String
    DebitN As Double
    CreditN As Double
    SoldeN As Double
    DebitN1 As Double
    CreditN1 As Double
    SoldeN1 As Double
End Type

' Tableaux globaux
Public BalanceN() As LigneBalance
Public BalanceN1() As LigneBalance
Public NbLignesN As Long
Public NbLignesN1 As Long

' ============================================================
' POINT D'ENTREE PRINCIPAL
' ============================================================
Sub LancerAnalyseFinanciere()
    Dim msg As String

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    On Error GoTo ErreurGlobale

    ' Verification des onglets source
    If Not VerifierOngletsSource() Then
        MsgBox "Veuillez creer les onglets '" & NOM_ONGLET_N & "' et '" & NOM_ONGLET_N1 & "' avec les balances comptables.", vbExclamation
        GoTo Fin
    End If

    ' Chargement des donnees
    Call ChargerBalances

    ' Suppression des anciens onglets d'analyse
    Call SupprimerOngletsAnalyse

    ' Creation des onglets par typologie
    Call CreerOngletActif
    Call CreerOngletImmobilisations
    Call CreerOngletStocks
    Call CreerOngletCreances
    Call CreerOngletTresorerie
    Call CreerOngletCapitauxPropres
    Call CreerOngletDettes
    Call CreerOngletProvisionsRisques

    ' Onglets P&L
    Call CreerOngletCA
    Call CreerOngletProduitsExploitation
    Call CreerOngletChargesExploitation
    Call CreerOngletPersonnel
    Call CreerOngletDotationsAmortissements
    Call CreerOngletFinancier
    Call CreerOngletExceptionnel
    Call CreerOngletImpots

    ' Recapitulatifs
    Call CreerRecapBilan
    Call CreerRecapPL
    Call CreerOngletSIG
    Call CreerOngletBFR_CAF
    Call CreerOngletANC
    Call CreerOngletPPT

    MsgBox "Analyse financiere generee avec succes !" & vbCrLf & _
           NbLignesN & " comptes exercice N" & vbCrLf & _
           NbLignesN1 & " comptes exercice N-1", vbInformation

Fin:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Exit Sub

ErreurGlobale:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    MsgBox "Erreur " & Err.Number & " : " & Err.Description, vbCritical
End Sub

' ============================================================
' VERIFICATION DES ONGLETS SOURCE
' ============================================================
Function VerifierOngletsSource() As Boolean
    Dim ws As Worksheet
    Dim bN As Boolean, bN1 As Boolean

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = NOM_ONGLET_N Then bN = True
        If ws.Name = NOM_ONGLET_N1 Then bN1 = True
    Next ws

    VerifierOngletsSource = (bN And bN1)
End Function

' ============================================================
' CHARGEMENT DES BALANCES
' ============================================================
Sub ChargerBalances()
    Call ChargerUneBalance(NOM_ONGLET_N, BalanceN, NbLignesN)
    Call ChargerUneBalance(NOM_ONGLET_N1, BalanceN1, NbLignesN1)
End Sub

Sub ChargerUneBalance(NomOnglet As String, ByRef Bal() As LigneBalance, ByRef NbLignes As Long)
    Dim ws As Worksheet
    Dim i As Long, derniereLigne As Long

    Set ws = ThisWorkbook.Worksheets(NomOnglet)
    derniereLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' On suppose que la ligne 1 est l'en-tete
    NbLignes = derniereLigne - 1
    If NbLignes <= 0 Then NbLignes = 0 : Exit Sub

    ReDim Bal(1 To NbLignes)

    For i = 1 To NbLignes
        Bal(i).NumCompte = Trim(CStr(ws.Cells(i + 1, 1).Value))
        Bal(i).IntituleCompte = Trim(CStr(ws.Cells(i + 1, 2).Value))
        Bal(i).DebitN = ValeurNumerique(ws.Cells(i + 1, 3).Value)
        Bal(i).CreditN = ValeurNumerique(ws.Cells(i + 1, 4).Value)
        ' Solde = Debit - Credit (solde debiteur positif)
        Bal(i).SoldeN = Bal(i).DebitN - Bal(i).CreditN
    Next i
End Sub

' ============================================================
' FONCTIONS UTILITAIRES
' ============================================================
Function ValeurNumerique(v As Variant) As Double
    If IsNumeric(v) Then
        ValeurNumerique = CDbl(v)
    Else
        ValeurNumerique = 0
    End If
End Function

Function GetSoldeN(NumCompteDebut As String, NumCompteFin As String) As Double
    Dim i As Long
    Dim total As Double
    total = 0
    For i = 1 To NbLignesN
        If BalanceN(i).NumCompte >= NumCompteDebut And BalanceN(i).NumCompte <= NumCompteFin Then
            total = total + BalanceN(i).SoldeN
        End If
    Next i
    GetSoldeN = total
End Function

Function GetSoldeN1(NumCompteDebut As String, NumCompteFin As String) As Double
    Dim i As Long
    Dim total As Double
    total = 0
    For i = 1 To NbLignesN1
        If BalanceN1(i).NumCompte >= NumCompteDebut And BalanceN1(i).NumCompte <= NumCompteFin Then
            total = total + BalanceN1(i).SoldeN
        End If
    Next i
    GetSoldeN1 = total
End Function

Function GetSoldeCompteN(NumCompte As String) As Double
    Dim i As Long
    For i = 1 To NbLignesN
        If Left(BalanceN(i).NumCompte, Len(NumCompte)) = NumCompte Then
            GetSoldeCompteN = GetSoldeCompteN + BalanceN(i).SoldeN
        End If
    Next i
End Function

Function GetSoldeCompteN1(NumCompte As String) As Double
    Dim i As Long
    For i = 1 To NbLignesN1
        If Left(BalanceN1(i).NumCompte, Len(NumCompte)) = NumCompte Then
            GetSoldeCompteN1 = GetSoldeCompteN1 + BalanceN1(i).SoldeN
        End If
    Next i
End Function

' ============================================================
' SUPPRESSION DES ANCIENS ONGLETS D'ANALYSE
' ============================================================
Sub SupprimerOngletsAnalyse()
    Dim ws As Worksheet
    Dim nomsASupprimer() As String
    Dim noms As String

    noms = "Actif immobilise,Stocks,Creances,Tresorerie,Capitaux propres,Dettes,Provisions,CA,Produits exploitation," & _
           "Charges exploitation,Personnel,Dotations,Financier,Exceptionnel,Impots,Recap Bilan,Recap PL,SIG,BFR-CAF,Annexe ANC,Synthese PPT"

    nomsASupprimer = Split(noms, ",")

    Application.DisplayAlerts = False
    Dim ws2 As Worksheet
    For Each ws2 In ThisWorkbook.Worksheets
        Dim n As Variant
        For Each n In nomsASupprimer
            If ws2.Name = CStr(n) Then
                ws2.Delete
                Exit For
            End If
        Next n
    Next ws2
    Application.DisplayAlerts = True
End Sub

' ============================================================
' CREATION D'UN ONGLET STANDARD
' ============================================================
Function CreerOnglet(NomOnglet As String, CouleurTabulation As Long) As Worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = NomOnglet
    ws.Tab.Color = CouleurTabulation
    Set CreerOnglet = ws
End Function

' ============================================================
' EN-TETE STANDARD D'UN ONGLET
' ============================================================
Sub EcrireEnTete(ws As Worksheet, Titre As String, SousTitre As String)
    With ws
        ' Titre principal
        .Range("A1").Value = Titre
        With .Range("A1")
            .Font.Bold = True
            .Font.Size = 14
            .Font.Color = RGB(255, 255, 255)
        End With
        .Range("A1:H1").Merge
        .Range("A1:H1").Interior.Color = RGB(31, 73, 125)

        ' Sous-titre
        .Range("A2").Value = SousTitre
        .Range("A2:H2").Merge
        .Range("A2:H2").Interior.Color = RGB(68, 114, 196)
        .Range("A2").Font.Color = RGB(255, 255, 255)
        .Range("A2").Font.Italic = True

        ' En-tetes colonnes
        .Range("A3").Value = "N° Compte"
        .Range("B3").Value = "Intitule"
        .Range("C3").Value = "Solde N"
        .Range("D3").Value = "Solde N-1"
        .Range("E3").Value = "Variation"
        .Range("F3").Value = "Var. %"
        .Range("G3").Value = "Note"

        With .Range("A3:G3")
            .Font.Bold = True
            .Interior.Color = RGB(189, 215, 238)
            .HorizontalAlignment = xlCenter
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
        End With

        ' Largeurs colonnes
        .Columns("A").ColumnWidth = 12
        .Columns("B").ColumnWidth = 45
        .Columns("C").ColumnWidth = 16
        .Columns("D").ColumnWidth = 16
        .Columns("E").ColumnWidth = 16
        .Columns("F").ColumnWidth = 10
        .Columns("G").ColumnWidth = 30

        ' Format numerique
        .Columns("C:E").NumberFormat = "#,##0;[Red]-#,##0"
        .Columns("F").NumberFormat = "0.0%"
    End With
End Sub

' Ecrire une ligne de compte
Sub EcrireLigneCompte(ws As Worksheet, Ligne As Long, NumCompte As String, Intitule As String, SoldeN As Double, SoldeN1 As Double)
    ws.Cells(Ligne, 1).Value = NumCompte
    ws.Cells(Ligne, 2).Value = Intitule
    ws.Cells(Ligne, 3).Value = SoldeN
    ws.Cells(Ligne, 4).Value = SoldeN1
    ws.Cells(Ligne, 5).Value = SoldeN - SoldeN1
    If SoldeN1 <> 0 Then
        ws.Cells(Ligne, 6).Value = (SoldeN - SoldeN1) / Abs(SoldeN1)
    End If
    ' Alternance de couleur
    If Ligne Mod 2 = 0 Then
        ws.Range(ws.Cells(Ligne, 1), ws.Cells(Ligne, 7)).Interior.Color = RGB(242, 242, 242)
    End If
End Sub

' Ecrire un sous-total
Sub EcrireSousTotal(ws As Worksheet, Ligne As Long, Libelle As String, SoldeN As Double, SoldeN1 As Double)
    ws.Cells(Ligne, 2).Value = Libelle
    ws.Cells(Ligne, 3).Value = SoldeN
    ws.Cells(Ligne, 4).Value = SoldeN1
    ws.Cells(Ligne, 5).Value = SoldeN - SoldeN1
    If SoldeN1 <> 0 Then ws.Cells(Ligne, 6).Value = (SoldeN - SoldeN1) / Abs(SoldeN1)
    With ws.Range(ws.Cells(Ligne, 1), ws.Cells(Ligne, 7))
        .Font.Bold = True
        .Interior.Color = RGB(189, 215, 238)
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlDouble
    End With
End Sub

' Ecrire un total general
Sub EcrireTotal(ws As Worksheet, Ligne As Long, Libelle As String, SoldeN As Double, SoldeN1 As Double)
    ws.Cells(Ligne, 2).Value = Libelle
    ws.Cells(Ligne, 3).Value = SoldeN
    ws.Cells(Ligne, 4).Value = SoldeN1
    ws.Cells(Ligne, 5).Value = SoldeN - SoldeN1
    If SoldeN1 <> 0 Then ws.Cells(Ligne, 6).Value = (SoldeN - SoldeN1) / Abs(SoldeN1)
    With ws.Range(ws.Cells(Ligne, 1), ws.Cells(Ligne, 7))
        .Font.Bold = True
        .Font.Size = 11
        .Interior.Color = RGB(31, 73, 125)
        .Font.Color = RGB(255, 255, 255)
    End With
End Sub
