
namespace ToolBoxTMA
{
    using OfficeOpenXml;
    using System;
    using System.Collections.Generic;
    using System.Data;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Text;

    /// <summary>
    /// Classe de gestion des fichiers
    /// </summary>
    public class FileHandler
    {
        #region Constantes

        /// <summary>
        /// Pattern d'expression régulière du format heure
        /// </summary>
        private const string timePattern = "_[0-9]{6}";
        #endregion

        #region Champs publiques

        /// <summary>
        /// Recherche V Excel
        /// </summary>
        /// <param name="searchString"></param>
        /// <param name="worksheet"></param>
        /// <returns>Retourne les range où la valeur est trouvée</returns>
        public delegate List<ExcelRange> RechercheV(string searchString, ExcelWorksheet worksheet);
        #endregion

        #region Champs privés

        /// <summary>
        /// Chemin du répertoire ou du fichier d'entré
        /// </summary>
        private string inputPath;

        /// <summary>
        /// Chemin du répertoire de sorti
        /// </summary>
        private string outputPath;

        /// <summary>
        /// Chaîne de caractères recherchée
        /// </summary>
        private string searchSequence;

        /// <summary>
        /// Extension du fichier
        /// </summary>       
        private string fileType;

        /// <summary>
        /// Séparateur de fichier
        /// </summary>
        private char separator;

        /// <summary>
        /// Nom du fichier de sortie
        /// </summary>
        private string destinationFile;

        /// <summary>
        /// Option de recherche dans une arborescence de dossiers
        /// </summary>
        private SearchOption searchOption;

        /// <summary>
        /// Liste d'onglets Excel
        /// </summary>
        private List<string> sheetNames;

        /// <summary>
        /// Table de correspondance entre des colonnes Excel
        /// </summary>
        private DataRowCollection mappedCols;
        #endregion

        #region Constructeurs

        /// <summary>
        /// Constructeur avec 9 paramètres du gestionnaire de fichiers
        /// </summary>
        /// <param name="inputPath">Chemin du répertoire ou du fichier d'entré</param>
        /// <param name="searchOption">Flag d'inclusion des sous dossiers</param>
        /// <param name="fileType">fileType des fichiers générés</param>
        /// <param name="outputPath">Répertoire de sortie</param>
        /// <param name="searchSequence">Filtre de recherche de fichier</param>
        /// <param name="destinationFile">Nom du fichier généré</param>
        /// <param name="separator">Séparateur de fichiers csv</param>
        /// <param name="sheetNames">Onglets Excel</param>
        /// <param name="mappedCols">Colonnes Excel mappées</param>
        public FileHandler(
                            string inputPath, 
                            SearchOption searchOption, 
                            string fileType, 
                            string outputPath, 
                            string searchSequence,
                            string destinationFile,
                            char? separator, 
                            List<string> sheetNames,
                            DataRowCollection mappedCols
                          )

        {
            this.inputPath = inputPath;          
            this.searchOption = searchOption;
            this.fileType = fileType;
            this.outputPath = outputPath;
            this.searchSequence = searchSequence;
            this.destinationFile = destinationFile;
            this.separator = separator.HasValue? separator.Value : default(char);
            this.sheetNames = sheetNames;
            this.mappedCols = mappedCols;
        }

        /// <summary>
        /// Constructeur avec 3 paramètres du gestionnaire de fichiers
        /// </summary>
        /// <param name="inputPath">Chemin du répertoire ou du fichier d'entré</param>
        /// <param name="searchOption">Flag d'inclusion des sous dossiers</param>
        /// <param name="fileType">fileType des fichiers générés</param>
        public FileHandler(
                            string inputPath,
                            SearchOption searchOption,
                            string fileType)
            : this(inputPath, searchOption, fileType, null, null, null, null, null, null)
        {
            
        }

        /// <summary>
        /// Constructeur avec 6 paramètres du gestionnaire de fichiers
        /// </summary>
        /// <param name="inputPath">Chemin du répertoire ou du fichier d'entré</param>
        /// <param name="searchOption">Flag d'inclusion des sous dossiers</param>
        /// <param name="fileType">fileType des fichiers générés</param>
        /// <param name="outputPath">Répertoire de sortie</param>
        /// <param name="searchSequence">Filtre de recherche de fichier</param>
        /// <param name="destinationFile">Nom du fichier généré</param>
        public FileHandler(
                            string inputPath,
                            SearchOption searchOption,
                            string fileType,
                            string outputPath,
                            string searchSequence,
                            string destinationFile)
            : this(inputPath, searchOption, fileType, outputPath, searchSequence, destinationFile, null, null, null)
        {

        }
        #endregion

        #region Propriétés

        /// <summary>
        /// Propriété FileType
        /// </summary>
        public string FileType 
        {
            get 
            {
                return this.fileType;
            }

            private set 
            {
                this.fileType = value;
            }
        }

        /// <summary>
        /// Propriété DestinationFile
        /// </summary>
        public string DestinationFile 
        {
            get
            {
                return this.destinationFile;
            }

            private set
            {
                this.destinationFile = value;
            } 
        }

        /// <summary>
        /// Proriété GetRangesFrom - Propriété en lecture seule
        /// </summary>
        public RechercheV GetRangesFrom
        {
            get 
            {
                return (str2Find, sh) => 
                                                {
                                                    Dictionary <Tuple <int, int>, Object> cellsFromSheet = this.RecupererDonneesOnglet(sh);

                                                    if (
                                                            cellsFromSheet.Any(kvp => kvp.Value != null && str2Find.Contains(kvp.Value.ToString().Trim()))
                                                       )

                                                        return
                                                                cellsFromSheet
                                                                                .Where(kvp => kvp.Value != null && str2Find.Contains(kvp.Value.ToString().Trim()))
                                                                                .Select(kvp => sh.Cells[kvp.Key.Item1, kvp.Key.Item2])
                                                                                .ToList();
                                                    else
                                                        throw new Exception("La valeur '" + str2Find + "' n'a pas été trouvé dans l'onglet " + sh);
                                                };
            }
        }

        public List<string> Onglets
        {
            get
            {
                return this.sheetNames;
            }
        }
        #endregion

        #region Méthodes publiques

        /// <summary>
        /// Récupère les fichiers de même type
        /// </summary>
        /// <returns>Renvoi la liste des fichiers du même type contenant au moins une ligne</returns>
        public List<String> RetrieveComonFiles() 
        {
            List<string> res = null;

            if (Directory.EnumerateFiles(this.inputPath, "*" + this.fileType, this.searchOption).FirstOrDefault() != null)
                res = Directory.EnumerateFiles(this.inputPath, "*" + this.fileType, this.searchOption).ToList();

            return res;
        }

        /// <summary>
        /// Renomme les fichiers plats horodatés
        /// </summary>
        /// <param name="ListeFichiers"></param>
        /// <returns>Renvoi la liste des fichiers horodatés</returns>
        public List<String> RenameTimeSpanFiles(List<string> ListeFichiers)
        {
            List<string> fichiersCrees = null;            

            // Enumération des fichiers du même type
            ListeFichiers.ForEach(
                                    fic =>
                                            {
                                                FileInfo fi = new FileInfo(fic);
                                                string nomFinal = fi.Name.Substring(0, fi.Name.Length - "_HHmmss".Length - FileType.Length);
                                                string dest = null;

                                                // Test pour savoir si d'autres fichiers auront le même nom
                                                if (ListeFichiers.Count(ffi => ffi.Contains(nomFinal)) > 1)
                                                {
                                                    if (!Directory.Exists(
                                                                            Path.Combine
                                                                            (
                                                                                fi.DirectoryName,
                                                                                fi.Name.Substring(0, fi.Name.Length - FileType.Length)
                                                                            )
                                                                          )
                                                       )
                                                    {
                                                        Directory.CreateDirectory(
                                                                                    Path.Combine
                                                                                    (
                                                                                        fi.DirectoryName,
                                                                                        fi.Name.Substring(0, fi.Name.Length - FileType.Length)
                                                                                    )
                                                                                 );
                                                    }

                                                    dest = Path.Combine(
                                                                            fi.DirectoryName,
                                                                            fi.Name.Substring(0, fi.Name.Length - FileType.Length),
                                                                            nomFinal + FileType
                                                                       );
                                                }
                                                else
                                                    dest = Path.Combine(fi.DirectoryName, nomFinal + FileType);

                                                // Enregistrement du nouveau fichier
                                                File.Copy(fic, dest, true);

                                                // Ajout du fichier renommé à la liste des fichiers renommés
                                                if (fichiersCrees == null)
                                                    fichiersCrees = new List<string>();

                                                fichiersCrees.Add(dest);
                                            }
                                 );

            return fichiersCrees;
        }

        /// <summary>
        /// Concaténation de fichiers de même type et de même En-tête (csv, dat, txt, ...)
        /// </summary>
        /// <param name="source">Liste des lignes ou des fichiers contenant les lignes</param>
        /// <param name="fromXL">Flag indiquant si les lignes proviennent d'Excel</param>
        /// <returns>Renvoie le chemin du fichier créé</returns>
        public string ConcatFiles(List<string> source, bool fromXL = false) 
        {                       

            // Initialisation du nom de fichier de sorti si les données sont issues d'Excel
            if (fromXL)
                DestinationFile = "BFC_REAL_ESTATE_" + DateTime.Now.ToString("yyyyMMdd");

            // Construction du chemin de sorti
            string dest = Path.Combine(this.outputPath, DestinationFile + FileType);

            // Suppression du fichier existant si il est situé dans le même répertoire
            if (File.Exists(dest))
                File.Delete(dest);

            if (!fromXL)
            {
                List<string> lignesFichier = null;
                int compteurFic = -1;

                // Enumération des noms de fichiers non formaté Excel
                source.ForEach(
                                fic =>
                                        {
                                            compteurFic += 1;
                                            // Récupération de toutes les lignes lues
                                            lignesFichier = File.ReadLines(fic).ToList();

                                            // Même Entête donc on passe à la ligne suivante
                                            if (compteurFic > 0 && lignesFichier[0].Trim().Contains(File.ReadAllLines(dest).ToList()[0].Trim()))
                                                lignesFichier.RemoveAt(0);

                                            // Ecriture des lignes
                                            File.AppendAllLines(dest, lignesFichier);
                                        }
                              );
            }
            else
                // Ajoût des lignes Excel lues au fichier de destination
                File.AppendAllLines(dest, source);

            return dest;
        }

        /// <summary>
        /// Constructeur de lignes issues d'Excel
        /// </summary>
        /// <returns>Renvoi l'ensemble des lignes d'un fichier Excel</returns>
        public List<string> GetLinesFromXL()
        {
            List<string> matchedDatas = null;
            string header;

            // Ouverture en lecture seule du classeur Excel 
            using (ExcelPackage pck = new ExcelPackage(File.Open(this.inputPath, FileMode.Open)))
            {
                // Contrôle d'existence des onglets
                if (!this.ChekeckedSheetName(pck.Workbook, this.sheetNames))
                    throw new Exception("Le(s) onglet(s) est/sont absent(s) du classeur");

                // Ajoût des zones nommées
                this.AddColumnsNames(pck.Workbook, out header);

                // Actions sur chaque onglet
                Onglets.ForEach(
                                    nomSh =>
                                                {
                                                    ExcelWorksheet sh = pck.Workbook.Worksheets[nomSh];

                                                    // Contrôle d'existence des zones nommées
                                                    if (sh.Names.Where(zn => zn.Name.StartsWith("RE_", StringComparison.Ordinal)) == null || sh.Names.Count(zn => zn.Name.StartsWith("RE_", StringComparison.CurrentCulture)) == 0)
                                                        throw new Exception("Aucune zones nommées dans l'onglet '" + sh.Name);

                                                    if (matchedDatas == null)
                                                    {
                                                        matchedDatas = new List<string>();
                                                        // Ajoût d'une En-tête aux lignes à récupérer
                                                        matchedDatas.Add(header);
                                                    }
                                                    // Cumul des lignes lues issues des onglets 
                                                    matchedDatas.AddRange(
                                                                             this.BuildLines(
                                                                                                this.RecupererDonneesOnglet(sh),
                                                                                                GetRangesFrom("Entity", sh).First().Start.Row + 1,
                                                                                                GetRangesFrom("TOTAL", sh).First().Start.Row,
                                                                                                sh.Names.Where(zn => zn.Name.StartsWith("RE_", StringComparison.Ordinal)).ToList()
                                                                                            )
                                                                         );
                                                }
                               );
            }                      
            return matchedDatas;
        }  
        #endregion   
   
        #region Méthodes privées

        /// <summary>
        /// Ajoute des zones nommées de feuille Excel
        /// </summary>
        /// <param name="workbook">Classeur Excel</param>
        private void AddColumnsNames(ExcelWorkbook workbook , out string header)
        {
            int cpt = 0;
            StringBuilder ch = null;
            string entete;

            // Parcours de la table de correspondance entre les colonnes d'onglets Excel 
            foreach (DataRow hd in this.mappedCols)
            {
                cpt += 1;

                // Ajoût des zones nommées sur l'onglet 'RE INVESTMENT'
                this.AddName2Sheet(
                                        hd[1].ToString().Trim(),
                                        workbook.Worksheets[this.sheetNames[0]],
                                        "RE_" + cpt.ToString()
                                  );

                // Ajoût des zones nommées sur l'onglet 'RE INVESTMENT ON USE'
                this.AddName2Sheet(
                                        hd[2].ToString().Trim(),
                                        workbook.Worksheets[this.sheetNames[1]],
                                        "RE_" + cpt.ToString()
                                  );

                // Construction d'une En-tête commune aux onglets - Si l'alias est présent alors il sert d'en-tête aux colonnes 
                if (!String.IsNullOrEmpty(hd["ALIAS_COL"].ToString().Trim()))
                    entete = hd["ALIAS_COL"].ToString().Trim();
                else
                {
                    // En-tête - Si l'alias est absent et que les en-têtes sont identiques alors le nom commun sert d'en-tête 
                    if (hd["RE_INVESTMENT"].ToString().Contains(hd["RE_OWN_USE"].ToString().Trim()))
                        entete = hd["RE_OWN_USE"].ToString().Trim();
                    else
                        // En-tête - Si l'alias est absent et que les en-têtes sont différentes alors l'en-tête est la concaténation des 5 premiers caractères de chacunes
                        entete = hd["RE_INVESTMENT"].ToString().Trim().Substring(0, 5).Trim() + "&&" + hd["RE_OWN_USE"].ToString().Trim().Substring(0, 5).Trim();
                }
                // Initialisation du buffer d'en-tête
                if (ch == null)
                {
                    // Ajoût de la première colonne d'en-tête
                    ch = new StringBuilder();                    
                    ch.Append(entete);
                }
                else
                    // Ajoût du séparateur et  de l'en-tête suivante
                    ch.Append(this.separator.ToString() + entete);
            }
            header = ch.ToString();
        }        

        /// <summary>
        /// Ajoute une zone nommée sur une feuille Excel
        /// </summary>
        /// <param name="hd2Find">Valeur recherchée dans un onglet</param>
        /// <param name="sh">Feuille Excel destinataire</param>
        /// <param name="namedZone">Nom de la zone nommée</param>
        private void AddName2Sheet(string hd2Find, ExcelWorksheet sh, string namedZone)
        {
            // Ajout de la zone de nom
            sh.Names.Add(
                            namedZone,
                            GetRangesFrom(hd2Find, sh).First()
                        );
        }        

        /// <summary>
        /// Récupération des cellules d'un onglet Excel  
        /// </summary>
        /// <param name="worksheet">Onglet Excel lu</param>
        /// <returns>Renvoi un dictionnaire de cellules d'un onglet Excel</returns>
        private Dictionary<Tuple<int, int>, Object> RecupererDonneesOnglet(ExcelWorksheet worksheet)
        {
            Dictionary<Tuple<int, int>, Object> cellsFromXLSheet = null;
            // Initialisation de la plage de recherche
            ExcelRange cells = worksheet.Cells;

            // Affectation de toutes les cellules (type Objet) dans un dictionnaire dont la clé est l'adresse de la cellule (Ligne, Colonne)
            cellsFromXLSheet = cells
                        .GroupBy(c => new { c.Start.Row, c.Start.Column })
                        .ToDictionary(
                                        rcg => new Tuple<int, int>(rcg.Key.Row, rcg.Key.Column),
                                        rcg => cells[rcg.Key.Row, rcg.Key.Column].Value
                                      );
            // Si aucune données récupérées alors une erreur est lancée
            if (cellsFromXLSheet == null || cellsFromXLSheet.Count == 0)
                throw new Exception("Aucune valeur n'a pu être récupérée de l'onglet " + worksheet.Name);

            return cellsFromXLSheet;
        }

        /// <summary>
        /// Constructeur de lignes issues d'Excel
        /// </summary>
        /// <param name="xlCells">Dictionnaire de cellules Excel</param>
        /// <param name="topLine">Ligne Excel de début de lecture</param>
        /// <param name="endLine">Ligne Excel de fin de lecture</param>
        /// <param name="nzHeaders">Liste des ranges des entêtes de colonne</param>
        /// <returns>Renvoi la liste des lignes construites</returns>
        private List<string> BuildLines(Dictionary<Tuple<int, int>, Object> xlCells, int topLine, int endLine, List<ExcelNamedRange> nzHeaders)
        {
            List<string> matchedDatas = new List<string>();
            StringBuilder rowData = null;
            ExcelWorksheet ws = nzHeaders[0].Worksheet;

            // Parcours des lignes de l'onglet Excel
            for (int numLig = topLine; numLig < endLine; numLig++)
            {
                string xlValue;
                // Parcours des colonnes Excel taguées par la zone nommée RE_X
                foreach (ExcelNamedRange namedZone in nzHeaders)
                {
                    // Récupération de la cellule située en ligne : numLig et en colonne : col RE_X
                    Object xlCel = xlCells.FirstOrDefault(kvp => kvp.Key.Item1 == numLig && kvp.Key.Item2 == namedZone.Start.Column).Value;
                    // Récupération de la valeur de la cellule
                    xlValue = xlCel == null ? String.Empty : String.Format(CultureInfo.CurrentCulture, xlCel.ToString().Trim());

                    // Si on se situe sur la colonne Building et que la valeur est nulle => on passe à la ligne suivante sans rien enregistrer
                    if (ws.Cells[namedZone.Name].Text.Trim().StartsWith("Building", StringComparison.Ordinal) && String.IsNullOrEmpty(xlValue))
                    {
                        if (rowData != null)
                            rowData = null;
                        break;
                    }

                    // Recherche du séparateur décimal ','
                    if (!string.IsNullOrEmpty(xlValue) && xlValue.Contains(','))
                    {
                        // Conversion des montants avec un séparateur décimal '.'
                        double testRes = Double.NaN;
                        if (double.TryParse(xlValue, NumberStyles.Any, CultureInfo.CurrentCulture, out testRes))
                            xlValue = testRes.ToString().Replace(',', '.');
                    }

                    // Initialisation du buffer avec la première valeur de la ligne et de la colonne
                    if (rowData == null)
                    {
                        rowData = new StringBuilder();
                        rowData.Append(xlValue);
                    }
                    else
                        // Ajoût du séparateur et de la valeur de la colonne au buffer
                        rowData.Append(this.separator.ToString() + xlValue);
                }
                if (rowData != null && rowData.ToString().Length > 0)
                {
                    // Ajout de la ligne à la collection de lignes
                    matchedDatas.Add(rowData.ToString());
                    // Reset du buffer
                    rowData = null;
                }
            }
            return matchedDatas;
        }
        
        /// <summary>
        /// Teste l'existence d'un onglet Excel
        /// </summary>
        /// <param name="wkb">Classeur</param>
        /// <param name="sheetName">Nom de l'onglet</param>
        /// <returns>Renvoi true si 'longlet existe sinon false</returns>
        private bool SheetExists(ExcelWorkbook wkb, string sheetName)
        {
            return wkb.Worksheets.Contains(wkb.Worksheets[sheetName]);
        }

        /// <summary>
        /// Teste l'existence d'une liste d'onglets
        /// </summary>
        /// <param name="excelWorkbook">Classeur</param>
        /// <param name="sheets">Liste des onglets recherchés</param>
        /// <returns>Renvoi true si toute la liste est présente sinon false</returns>
        private bool ChekeckedSheetName(ExcelWorkbook excelWorkbook, List<string> sheets)
        {
            bool res = true;
            // Parcours de la liste des onglets envoyés par Markit
            foreach (string sheetName in sheets)
            {
                // Renvoi Vrai si l'onglet existe sion Faux
                if (!this.SheetExists(excelWorkbook, sheetName))
                {
                    res = false;
                    break;
                }
            }
            return res;
        }
        #endregion

    }
}
