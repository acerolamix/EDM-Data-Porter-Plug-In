

namespace FonctionsTMA
{
    using CADIS.Plugin;
    using System;
    using System.Collections.Generic;
    using System.Data;
    using System.IO;
    using System.Linq;
    using ToolBoxTMA;

    /// <summary>
    /// Plugin de concaténation d'onglets Excel
    /// </summary>
    public class ConcatXL2CSV : DataPorterPlugin
    {
        #region Constants

        /// <summary>
        /// Liste des onglets de données
        /// </summary>
        private const string PRM_LST_SH = "Liste des onglets à traiter - Séparateur ';'";

        /// <summary>
        /// Table de Mapping des colonnes Excel
        /// </summary>
        private const string PRM_MAPPING_TABLE = "Table de référence sur le mapping BFC RE";

        /// <summary>
        /// Séparateur de colonnes
        /// </summary>
        private const string PRM_FIC_SEP = "Séparateur de colonne à utiliser";
        #endregion

        #region Properties

        /// <summary>
        /// Dictionnaire des paramètres d'entrée - Propriété en lecture seule
        /// </summary>
        public override Dictionary<string, string> InputParameters
        {
            get
            {
                // Initialisation des paramètres d'entré Markit
                Dictionary<String, String> inputs = new Dictionary<string, string>();

                inputs.Add(Concatener.PRM_REP_IN, "Mandatory - Chemin du fichier Excel");
                inputs.Add(Concatener.PRM_REP_OUT, "Mandatory - Répertoire de sortie du fichier csv");
                inputs.Add(PRM_LST_SH, "Mandatory - Liste des onglets à traiter - Séparateur ';'");
                inputs.Add(PRM_MAPPING_TABLE, "Mandatory - Table de référence sur le mapping BFC RE");

                inputs.Add(PRM_FIC_SEP, "Optional - Séparateur de colonne à utiliser - Caractère ';' si non renseigné");
                inputs.Add(Concatener.PRM_VERBOSE, "Optional - Caractère permettant d'activer les logs : 'Y' -> Active l'affichage");

                return inputs;
            }
        }

        /// <summary>
        /// Dictionnaire des paramètres de sortie - Propriété en lecture seule
        /// </summary>
        public override Dictionary<string, string> OutputParameters
        {
            get
            {
                return GetOutputParameters();
            }
        }

        /// <summary>
        /// Description du plugin Markit - Propriété en lecture seule
        /// </summary>
        public override string Description
        {
            get { return "Concaténation de deux onglets Excel au format CSV."; }
        }
        #endregion

        #region Méthodes publiques

        /// <summary>
        /// Paramètres de sortie du plugin Markit
        /// </summary>
        /// <returns>Renvoi le dictionnaire des paramètre de sortie</returns>
        public Dictionary<string, string> GetOutputParameters()
        {

            Dictionary<string, string> outputs = new Dictionary<string, string>();
            // Initialisation du dictionnaire
            outputs.Add(Concatener.OUTPUT_RETURNCODE, "ReturnCode (0 = success)");
            outputs.Add(Concatener.OUTPUT_ERRORMESSAGE, "Error Message");

            return outputs;
        }

        /// <summary>
        /// Point d'entrée de la dll/méthode sélectionnée depuis l'interface Markit
        /// </summary>
        /// <param name="inputParameters">Dictionnaire des paramètres d'entré</param>
        /// <param name="cadisVariables">Dictionnaire des variables Cadis</param>
        /// <returns></returns>
        protected override Dictionary<string, string> Run(Dictionary<string, string> inputParameters, Dictionary<string, string> cadisVariables)
        {
            Dictionary<string, string> outputParams = new Dictionary<string, string>();

            #region Validate Input params
            // Contrôle des paramètres (obligatoires et optionnels) saisis depuis l'interface Markit
            string monRepIn = GetMandatoryParameter(inputParameters, Concatener.PRM_REP_IN);
            string monRepOut = GetMandatoryParameter(inputParameters, Concatener.PRM_REP_OUT);

            List<string> listeSH = GetMandatoryParameter(inputParameters, PRM_LST_SH).Split(';').ToList();
            string mapping = GetMandatoryParameter(inputParameters, PRM_MAPPING_TABLE);

            char separator = String.IsNullOrEmpty(inputParameters[PRM_FIC_SEP].Trim()) ? ';' : inputParameters[PRM_FIC_SEP].Trim()[0];
            bool doVerbs = inputParameters[Concatener.PRM_VERBOSE].Trim() == "Y" ? true : false;
            #endregion

            #region Invocation de la fonction  

            try
            {
                // Passage de la requête SQL
                if (doVerbs)
                    LogMessage(MessageSeverity.Information, "Passage de la requête 'SELECT INDICE_COL, RE_INVESTMENT, RE_OWN_USE FROM " + mapping + "'");
                DataTable dtMapping = this.RunCadisTableQuery("SELECT INDICE_COL, RE_INVESTMENT, RE_OWN_USE, ALIAS_COL FROM " + mapping);

                // Erreur si la requête ne retourne rien
                if (dtMapping.Rows.Count < 1)
                    throw new Exception("La table de référence de mapping '" + mapping + "' est vide");

                // Initialisation de la classe de gestion des fichiers
                if (doVerbs)
                    LogMessage(MessageSeverity.Information, "Initialisation des paramètres de la fonction 'ConcatXL2CSV'");
                FileHandler gestFic = new FileHandler(monRepIn, SearchOption.TopDirectoryOnly, ".csv", monRepOut, "", "", separator, listeSH, dtMapping.Rows);

                // Récupération des lignes issues des onglets
                if (doVerbs)
                    LogMessage(MessageSeverity.Information, "Récupération des lignes issues du fichier '" + monRepIn + "'");
                List<string> lignesFichier = gestFic.GetLinesFromXL();

                // Concaténation des lignes dans un fichier csv
                if (doVerbs)
                    LogMessage(MessageSeverity.Information, "Concaténation des lignes récupérées");
                string unFic = gestFic.ConcatFiles(lignesFichier, true);

                if (doVerbs)
                    LogMessage(MessageSeverity.Information, "Le fichier '" + unFic + "' a été généré");

                // Gestion du code retour - En cas de succès
                outputParams.Add(Concatener.OUTPUT_RETURNCODE, "0");
                outputParams.Add(Concatener.OUTPUT_ERRORMESSAGE, String.Empty);
            }
            catch (Exception ex)
            {
                // Gestion du code retour - En cas d'erreur
                outputParams.Add(Concatener.OUTPUT_RETURNCODE, "1");
                outputParams.Add(Concatener.OUTPUT_ERRORMESSAGE, ex.Message);

                // Lancement d'une exception en cas d'erreur
                throw new Exception(outputParams[Concatener.OUTPUT_ERRORMESSAGE]);
            }
            return outputParams;            
            #endregion
        }
        #endregion       
    }
}
