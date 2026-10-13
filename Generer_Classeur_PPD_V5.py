"""
PPD - Génération automatique d'un classeur Excel filtrable

OBJECTIF
---------
Ce script transforme un export CSV PPD en un classeur Excel au format XLSM
contenant un tableau filtrable et un outil d'export de la vue courante.

DONNÉE D'ENTRÉE
---------------
Fichier CSV PPD contenant les 16 colonnes standard du suivi documentaire
(produit automatiquement par l'outil PPD ou transmis en argument).

DONNÉE DE SORTIE
----------------
Création d'un fichier :

    <Nom_du_CSV>.xlsm

Le classeur est généré dans le même dossier que le fichier CSV source.

FONCTIONNALITÉS
---------------
- Détection automatique de l'encodage du CSV.
- Détection automatique du séparateur CSV.
- Contrôle de la présence des colonnes PPD attendues.
- Création d'un tableau Excel structuré nommé "TableauPPD".
- Conservation des identifiants au format texte.
- Conversion des dates en véritables dates Excel.
- Ajustement automatique des largeurs de colonnes.
- Mise en forme du tableau selon le standard PPD.
- Mise en évidence des documents en attente de retour TTL.
- Regroupement automatique des colonnes L à P.
- Figeage de la ligne d'en-tête.
- Création d'une feuille "Accueil" contenant un bouton d'export.


EXPORT DE LA VUE FILTRÉE
------------------------
Le bouton "Extraction CSV" présent dans l'onglet Accueil exporte
uniquement les lignes actuellement visibles dans le tableau après
application des filtres Excel.

Le fichier exporté est enregistré en :

- UTF-8 avec BOM
- séparateur virgule
- format date AAAA-MM-JJ

CONTRÔLES EFFECTUÉS
-------------------
Avant enregistrement du classeur, le script vérifie :

- la présence du tableau PPD ;
- la présence de la macro VBA ;
- l'affectation du bouton d'export ;
- l'absence de liaisons externes ;
- l'absence de connexions ou requêtes Excel.

PRÉREQUIS
---------
- Windows
- Microsoft Excel
- Python
- pywin32

Installation :

    pip install pywin32

L'option Excel "Accès approuvé au modèle d'objet du projet VBA"
doit être activée.
"""

from __future__ import annotations
import csv
import sys
from datetime import date, datetime
from pathlib import Path

try:
    import win32com.client as win32
except ImportError as exc:
    raise SystemExit("pywin32 est requis : pip install pywin32") from exc

ONGLET_ACCUEIL, ONGLET_PPD = "Accueil", "PPD"
TABLEAU, MODULE = "TableauPPD", "ModuleExtractionPPD"
MACRO, BOUTON = "ExtraireVueCouranteCSV", "BoutonExtractionCSV"
LIGNE_ENTETE = 1
XL_SRC_RANGE, XL_YES, XL_XLSM, XL_EXCEL_LINKS = 1, 1, 52, 1
XL_CENTER, XL_LEFT, XL_TOP = -4108, -4131, -4160
MSO_FORCE_DISABLE = 3

COLONNES = [
    "Num Liv.", "Responsable Titulaire", "Reference externe",
    "Indice de revision", "Titre du document", "Categorie du document",
    "Statut contractuel", "Jalon", "En attente retour TTL", "Commentaire",
    "Fournisseur", "Acceptation/Information", "Date dernier envoi",
    "N° Bordereau", "Date retour", "N° Lot",
]
COLONNES_DATE = {"Date dernier envoi", "Date retour"}
COLONNES_TEXTE = {"Num Liv.", "Reference externe", "Indice de revision", "N° Bordereau", "N° Lot"}
COLONNES_CENTREES = {"Indice de revision", "Statut contractuel", "Jalon", "En attente retour TTL", "Acceptation/Information"}
LARGEURS_FIXES = {"Indice de revision": 9, "Statut contractuel": 11.5, "Jalon": 6.5, "En attente retour TTL": 11.5, "Acceptation/Information": 12}

VBA = r'''
Option Explicit
Private Const adTypeText As Long = 2
Private Const adSaveCreateOverWrite As Long = 2

Public Sub ExtraireVueCouranteCSV()
    On Error GoTo GestionErreur
    Dim ws As Worksheet, tbl As ListObject, ligne As ListRow
    Dim contenu As String, chemin As String, nomFichier As String
    Dim i As Long, nbLignes As Long
    Set ws = ThisWorkbook.Worksheets("PPD")
    Set tbl = ws.ListObjects("TableauPPD")
    If tbl.DataBodyRange Is Nothing Then
        MsgBox "Le tableau ne contient aucune donnee.", vbInformation, "Extraction"
        Exit Sub
    End If
    For i = 1 To tbl.ListColumns.Count
        If i > 1 Then contenu = contenu & ","
        contenu = contenu & EchappementCSV(CStr(tbl.HeaderRowRange.Cells(1, i).Value2))
    Next i
    contenu = contenu & vbCrLf
    For Each ligne In tbl.ListRows
        If Not ligne.Range.EntireRow.Hidden Then
            For i = 1 To tbl.ListColumns.Count
                If i > 1 Then contenu = contenu & ","
                contenu = contenu & EchappementCSV(ValeurCSV(ligne.Range.Cells(1, i), CStr(tbl.HeaderRowRange.Cells(1, i).Value2)))
            Next i
            contenu = contenu & vbCrLf
            nbLignes = nbLignes + 1
        End If
    Next ligne
    If nbLignes = 0 Then
        MsgBox "Aucune ligne visible avec les filtres actuels.", vbInformation, "Extraction"
        Exit Sub
    End If
    nomFichier = NomSansExtension(ThisWorkbook.Name) & "_VUE_FILTREE.csv"
    chemin = ChoisirCheminCSV(nomFichier)
    If Len(chemin) = 0 Then Exit Sub
    EcrireUTF8AvecBOM chemin, contenu
    MsgBox CStr(nbLignes) & " ligne(s) exportee(s) vers :" & vbCrLf & chemin, vbInformation, "Extraction terminee"
    Exit Sub
GestionErreur:
    MsgBox "Erreur pendant l'extraction : " & Err.Number & " - " & Err.Description, vbCritical, "Extraction"
End Sub

Private Function ChoisirCheminCSV(ByVal nomFichier As String) As String
    Dim choix As Variant, dossierInitial As String
    dossierInitial = Environ$("USERPROFILE") & "\Documents\"
    choix = Application.GetSaveAsFilename(InitialFileName:=dossierInitial & nomFichier, FileFilter:="Fichiers CSV (*.csv), *.csv", Title:="Choisir l'emplacement local du fichier CSV")
    If VarType(choix) = vbBoolean Then
        ChoisirCheminCSV = ""
    Else
        ChoisirCheminCSV = CStr(choix)
        If LCase$(Right$(ChoisirCheminCSV, 4)) <> ".csv" Then ChoisirCheminCSV = ChoisirCheminCSV & ".csv"
    End If
End Function

Private Function ValeurCSV(ByVal cellule As Range, ByVal nomColonne As String) As String
    Dim valeur As Variant
    valeur = cellule.Value2
    If IsError(valeur) Or IsEmpty(valeur) Then
        ValeurCSV = ""
    ElseIf nomColonne = "Date dernier envoi" Or nomColonne = "Date retour" Then
        If IsNumeric(valeur) Then
            ValeurCSV = Format$(cellule.Value, "yyyy-mm-dd")
        ElseIf IsDate(valeur) Then
            ValeurCSV = Format$(CDate(valeur), "yyyy-mm-dd")
        Else
            ValeurCSV = CStr(valeur)
        End If
    Else
        ValeurCSV = CStr(valeur)
    End If
End Function

Private Function EchappementCSV(ByVal valeur As String) As String
    Dim resultat As String
    resultat = Replace(valeur, Chr$(34), Chr$(34) & Chr$(34))
    If InStr(resultat, ",") > 0 Or InStr(resultat, Chr$(34)) > 0 Or InStr(resultat, vbCr) > 0 Or InStr(resultat, vbLf) > 0 Then resultat = Chr$(34) & resultat & Chr$(34)
    EchappementCSV = resultat
End Function

Private Sub EcrireUTF8AvecBOM(ByVal chemin As String, ByVal contenu As String)
    Dim flux As Object
    Set flux = CreateObject("ADODB.Stream")
    With flux
        .Type = adTypeText
        .Charset = "utf-8"
        .Open
        .WriteText contenu
        .SaveToFile chemin, adSaveCreateOverWrite
        .Close
    End With
End Sub

Private Function NomSansExtension(ByVal nomComplet As String) As String
    Dim p As Long
    p = InStrRev(nomComplet, ".")
    If p > 1 Then NomSansExtension = Left$(nomComplet, p - 1) Else NomSansExtension = nomComplet
End Function
'''

def choisir_csv() -> Path:
    if len(sys.argv) > 1:
        p = Path(sys.argv[1]).expanduser()
    else:
        nom_fichier = input("\nNom ou chemin du fichier CSV à traiter : ").strip().strip('"')
        if not nom_fichier:
            raise ValueError("Aucun fichier CSV n'a été renseigné.")
        p = Path(nom_fichier).expanduser()
        if not p.is_absolute():
            p = Path(__file__).resolve().parent / p

    p = p.resolve()
    if not p.is_file():
        raise FileNotFoundError(f"Fichier CSV introuvable : {p}")
    if p.suffix.lower() != ".csv":
        raise ValueError("Le fichier source doit être au format .csv.")
    return p

def lire_csv(p: Path):
    texte = encodage = None
    for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            texte, encodage = p.read_text(encoding=enc), enc
            break
        except UnicodeDecodeError:
            pass
    if texte is None: raise UnicodeError("Encodage CSV non reconnu.")
    try: sep = csv.Sniffer().sniff(texte[:20000], delimiters=",;\t|").delimiter
    except csv.Error: sep = ","
    lecteur = csv.DictReader(texte.splitlines(), delimiter=sep)
    entetes = [str(x or "").strip() for x in (lecteur.fieldnames or [])]
    manquantes = [c for c in COLONNES if c not in entetes]
    if manquantes: raise ValueError("Colonnes absentes : " + ", ".join(manquantes))
    lignes = []
    for ligne in lecteur:
        n = {str(k or "").strip(): str(v or "").strip() for k, v in ligne.items() if k is not None}
        if any(n.values()): lignes.append(n)
    if not lignes: raise ValueError("Le CSV ne contient aucune ligne de donnees.")
    return lignes, encodage, sep

def convertir_date(v: str):
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d/%m/%y"):
        try: return datetime.strptime(v, fmt)
        except ValueError: pass
    return None

def numero_serie_excel(v: datetime) -> int:
    return (v.date() - date(1899, 12, 30)).days

def appliquer_regles(lignes) -> int:
    n = 0
    for ligne in lignes:
        if ligne.get("Statut contractuel", "").strip().upper() == "PRT" and ligne.get("En attente retour TTL", "").strip().casefold() != "non":
            ligne["En attente retour TTL"] = "Non"
            n += 1
    return n

def largeur(entete, lignes):
    maxi = 55 if entete in {"Titre du document", "Commentaire"} else 30
    return min(max(max([len(entete)] + [len(x.get(entete, "")) for x in lignes[:2000]]) + 2, 11), maxi)

def ajouter_macro(wb):
    try:
        projet = wb.VBProject
        module = projet.VBComponents.Add(1)
        module.Name = MODULE
        module.CodeModule.AddFromString(VBA)
    except Exception as exc:
        raise RuntimeError("Activez l'acces approuve au modele d'objet du projet VBA dans Excel.") from exc

def nettoyer(wb):
    for i in range(int(wb.Names.Count), 0, -1):
        try: wb.Names(i).Delete()
        except Exception: pass
    for i in range(int(wb.Connections.Count), 0, -1):
        try: wb.Connections(i).Delete()
        except Exception: pass
    try:
        for i in range(int(wb.Queries.Count), 0, -1): wb.Queries.Item(i).Delete()
    except Exception: pass
    liens = wb.LinkSources(XL_EXCEL_LINKS)
    if liens:
        for lien in list(liens): wb.BreakLink(Name=lien, Type=XL_EXCEL_LINKS)

def creer_accueil(wb):
    a = wb.Worksheets.Add(Before=wb.Worksheets(1)); a.Name = ONGLET_ACCUEIL
    a.Columns("A").ColumnWidth = 3; a.Columns("B").ColumnWidth = 28; a.Columns("C:F").ColumnWidth = 18
    a.Range("B2:F2").Merge(); a.Range("B2").Value = "Outil PPD"
    a.Range("B2").Font.Size = 20; a.Range("B2").Font.Bold = True; a.Range("B2").Font.Color = 0xFFFFFF
    a.Range("B2:F2").Interior.Color = 0xC47244; a.Range("B2:F2").HorizontalAlignment = XL_CENTER
    a.Range("B4:F4").Merge(); a.Range("B4").Value = "Appliquez les filtres dans la feuille PPD, puis revenez ici pour exporter uniquement les lignes visibles."
    a.Range("B4").WrapText = True; a.Rows(4).RowHeight = 42
    return a


def creer_classeur(csv_path: Path):
    lignes, enc, sep = lire_csv(csv_path)
    corrections = appliquer_regles(lignes)
    incoherences_prt = [
        ligne for ligne in lignes
        if ligne.get("Statut contractuel", "").strip().upper() == "PRT"
        and ligne.get("En attente retour TTL", "").strip().casefold() != "non"
    ]
    if incoherences_prt:
        raise RuntimeError("La regle PRT -> Non n'a pas ete appliquee a toutes les lignes.")
    dst = csv_path.with_name(f"{csv_path.stem}.xlsm")
    if dst.exists(): dst.unlink()
    excel = wb = None; etape = "initialisation"
    try:
        excel = win32.DispatchEx("Excel.Application"); excel.Visible = False
        excel.DisplayAlerts = False; excel.EnableEvents = False; excel.ScreenUpdating = False
        excel.AskToUpdateLinks = False; excel.AutomationSecurity = MSO_FORCE_DISABLE
        wb = excel.Workbooks.Add()
        while int(wb.Worksheets.Count) > 1: wb.Worksheets(wb.Worksheets.Count).Delete()
        ws = wb.Worksheets(1); ws.Name = ONGLET_PPD
        derniere = 1 + len(lignes)
        ws.Range(ws.Cells(1,1), ws.Cells(1,len(COLONNES))).Value2 = (tuple(COLONNES),)
        for i, ent in enumerate(COLONNES, 1):
            if ent in COLONNES_TEXTE: ws.Range(ws.Cells(2,i), ws.Cells(derniere,i)).NumberFormat = "@"
        etape = "import des donnees"
        for r, ligne in enumerate(lignes, 2):
            for c, ent in enumerate(COLONNES, 1):
                cell, txt = ws.Cells(r,c), ligne.get(ent, "")
                if ent in COLONNES_DATE:
                    d = convertir_date(txt)
                    if d:
                        cell.Value2 = numero_serie_excel(d)
                        try: cell.NumberFormatLocal = "jj/mm/aaaa"
                        except Exception: cell.NumberFormat = "dd/mm/yyyy"
                    else: cell.Value2 = txt
                else: cell.Value2 = txt
        etape = "mise en forme"
        for i, ent in enumerate(COLONNES, 1):
            ws.Columns(i).ColumnWidth = LARGEURS_FIXES.get(ent, largeur(ent, lignes))
            z = ws.Range(ws.Cells(2,i), ws.Cells(derniere,i)); z.VerticalAlignment = XL_TOP; z.WrapText = False
            if ent in COLONNES_CENTREES: z.HorizontalAlignment = XL_CENTER; z.WrapText = True
        plage = ws.Range(ws.Cells(1,1), ws.Cells(derniere,len(COLONNES)))
        table = ws.ListObjects.Add(XL_SRC_RANGE, plage, None, XL_YES); table.Name = TABLEAU; table.TableStyle = "TableStyleMedium2"
        ws.Rows(1).RowHeight = 35
        hdr = ws.Range("A1:P1"); hdr.HorizontalAlignment = XL_CENTER; hdr.VerticalAlignment = XL_CENTER; hdr.WrapText = True
        ttl = COLONNES.index("En attente retour TTL") + 1
        for r in range(2, derniere + 1):
            cell = ws.Cells(r, ttl)
            if str(cell.Value2 or "").strip().casefold() == "oui":
                cell.Interior.Color = 0xCCF2FF; cell.Font.Bold = True; cell.Font.Color = 0x00659C
        ws.Columns("L:P").Columns.Group(); ws.Columns("L:P").Hidden = True
        ws.Activate(); excel.ActiveWindow.DisplayGridlines = False; excel.ActiveWindow.SplitRow = 1; excel.ActiveWindow.FreezePanes = True
        accueil = creer_accueil(wb)
        ajouter_macro(wb); nettoyer(wb)
        wb.SaveAs(str(dst), FileFormat=XL_XLSM)
        ancre = accueil.Range("B6")
        bouton = accueil.Shapes.AddFormControl(0, float(ancre.Left), float(ancre.Top), 180.0, 38.0)
        bouton.Name = BOUTON; bouton.TextFrame.Characters().Text = "Extraction CSV"; bouton.OnAction = MACRO
        if wb.LinkSources(XL_EXCEL_LINKS) or int(wb.Connections.Count): raise RuntimeError("Une liaison externe subsiste.")
        if int(wb.VBProject.VBComponents.Item(MODULE).CodeModule.ProcStartLine(MACRO, 0)) <= 0: raise RuntimeError("Macro absente.")
        accueil.Activate(); wb.Save(); wb.Close(SaveChanges=False); wb = None
        print("=" * 72); print("GENERATION DU CLASSEUR PPD TERMINEE")
        print(f"CSV source          : {csv_path}"); print(f"Encodage detecte    : {enc}"); print(f"Separateur          : {repr(sep)}")
        print(f"Lignes importees    : {len(lignes)}"); print(f"Corrections PRT     : {corrections}"); print(f"Fichier cree        : {dst}")
        print("Fichier temporaire  : aucun"); print("=" * 72)
    except Exception as exc:
        raise RuntimeError(f"Erreur a l'etape '{etape}' : {exc}") from exc
    finally:
        if wb is not None:
            try: wb.Close(SaveChanges=False)
            except Exception: pass
        if excel is not None:
            try: excel.Quit()
            except Exception: pass

def main():
    try: creer_classeur(choisir_csv()); return 0
    except Exception as exc: print(f"ERREUR : {exc}", file=sys.stderr); return 1

if __name__ == "__main__":
    code_retour = 0

    try:
        code_retour = main()

    except Exception as exc:
        print("\nERREUR INATTENDUE :")
        print(exc)
        code_retour = 1

    finally:
        input("\nAppuyez sur Entrée pour fermer la fenêtre...")

    raise SystemExit(code_retour)
