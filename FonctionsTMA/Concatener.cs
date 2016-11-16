
namespace FonctionsTMA
{
    using CADIS.Plugin;
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Text.RegularExpressions;
    using ToolBoxTMA;

    /// <summary>
    /// Plugin de concaténation de fichiers horodatés
    /// </summary>
    public class Concatener : DataPorterPlugin
    {
        #region Constants
       

        /// <summary>
        /// Filtre de recherche à appliquer au nom du fichier horodaté
        /// </summary>
        private const String PRM_FILE_TO_FIND = "Filtre de recherche des fichiers à concaténer";

        /// <summary>
        /// Nom générique du fichier de sortie
        /// </summary>
        private const String PRM_FILE_NOM = "Nom du fichier créé";

        /// <summary>
        /// Chemin du rpertoire de recherche
        /// </summary>
        internal const String PRM_REP_IN = "Répertoire d'entré - Chemin du fichier";

        /// <summary>
        /// Chemin du répertoire de sortie
        /// </summary>
        internal const String PRM_REP_OUT = "Répertoire de sortie des fichiers concaténés";

        /// <summary>
        /// Option de recherche dans les sous dossiers
        /// </summary>
        internal const String PRM_INCLURE_SS_REP = "Flag indiquant si la fonction doit aussi traiter les sous-dossiers";

        /// <summary>
        /// Extension du fichier
        /// </summary>
        internal const String PRM_EXTENSION = "FileType des fichiers traités";

        /// <summary>
        /// Option d'affichage des logs
        /// </summary>
        internal const String PRM_VERBOSE = "Flag indiquant si les logs doivent être affichées";

        /// <summary>
        /// Code retour du plugin
        /// </summary>
        internal const String OUTPUT_RETURNCODE = "ReturnCode";

        /// <summary>
        /// Message d'erreur du plugin
        /// </summary>
        internal const String OUTPUT_ERRORMESSAGE = "ErrorMessage";

        /// <summary>
        /// Pattern du format heure
        /// </summary>
        internal const string timePattern = "_[0-9]{6}";
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

                inputs.Add(PRM_REP_IN, "Mandatory - Répertoire d'entré");
                inputs.Add(PRM_REP_OUT, "Mandatory - Répertoire de sortie des fichiers concaténés");
                inputs.Add(PRM_FILE_TO_FIND, "Mandatory - Filtre de recherche des fichiers à concaténer");
                inputs.Add(PRM_FILE_NOM, "Mandatory - Nom du fichier créé");

                inputs.Add(PRM_EXTENSION, "Optional - Type des fichiers traités. Par défaut '.csv'");
                inputs.Add(PRM_INCLURE_SS_REP, "Optional - Caractère indiquant si les sous-dossiers doivent être scrutés : 'Y' -> Inclure les sous-dossiers");
                inputs.Add(PRM_VERBOSE, "Optional - Caractère permettant d'activer les logs : 'Y' -> Active l'affichage");

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
            get { return "Concatétnation de fichiers plats (csv, dat, ...)"; }
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
            outputs.Add(OUTPUT_RETURNCODE, "ReturnCode (0 = success)");
            outputs.Add(OUTPUT_ERRORMESSAGE, "Error Message");

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
            string monRepIn = GetMandatoryParameter(inputParameters, PRM_REP_IN);
            string monRepOut = GetMandatoryParameter(inputParameters, PRM_REP_OUT);
            string filtre = GetMandatoryParameter(inputParameters, PRM_FILE_TO_FIND);
            string nomNouveauFic = GetMandatoryParameter(inputParameters, PRM_FILE_NOM);

            string extension = String.IsNullOrEmpty(inputParameters[PRM_EXTENSION].Trim()) ? ".csv" : inputParameters[PRM_EXTENSION].Trim();
            extension = extension.StartsWith(".", StringComparison.Ordinal) ? extension : "." + extension;

            SearchOption optionRecherche = inputParameters[PRM_INCLURE_SS_REP].Trim() == "Y" ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
            bool doVerbs = inputParameters[PRM_VERBOSE].Trim() == "Y" ? true : false;
            #endregion

            #region Invocation de la fonction            

            try
            {
                // Déclenchement de l'affichage de la log si le paramètre l'exige
                if (doVerbs)
                    LogMessage(MessageSeverity.Information, "Initialisation des paramètres de concaténation");
                // Initialisation de la classe de gestion des fichiers
                FileHandler gestFic = new FileHandler(monRepIn, optionRecherche, extension, monRepOut, filtre, nomNouveauFic);

                // Récupération et tri des fichiers du même type
                if (doVerbs)
                    LogMessage(MessageSeverity.Information, "Récupération des fichiers *" + extension);
                List<string> files2Process = gestFic.RetrieveComonFiles();

                // Lancement d'une exception si aucun fichier ne correspond au filtre
                if (files2Process == null || files2Process.FirstOrDefault(fi => Regex.IsMatch(fi, filtre)) == null)                    
                    throw new Exception("Aucun fichier à concaténer n'a été trouvé");

                // Initialisation des fichiers correspondants au filtre de recherche
                files2Process = files2Process.Where(fii => Regex.IsMatch(fii, filtre)).ToList();

                // Lancement d'une exception si Tous les fichiers correspondants au filtre sont vides
                if (files2Process.All(fic => File.ReadLines(fic).Count() == 0))
                {
                    if (doVerbs)
                        files2Process.ForEach(
                                                    unFic => { LogMessage(MessageSeverity.Information, new FileInfo(unFic).Name + " est vide"); }
                                             );
                    throw new Exception("Tous les fichiers correspondants au filtre sont vides");
                }                

                // Affichage des fichiers à traiter
                if (doVerbs)
                {
                    LogMessage(MessageSeverity.Information, "Liste des fichiers à concaténer contenant au moins une ligne : ");
                    files2Process
                                 .Where(foo => File.ReadLines(foo).Count() > 0)
                                 .ToList()
                                 .ForEach(
                                            fichier =>  
                                                        {
                                                            LogMessage(MessageSeverity.Information, new FileInfo(fichier).Name);
                                                        }
                                          );
                }

                // Concaténation des fichiers contenants au moins une ligne
                string res = gestFic.ConcatFiles(files2Process.Where(fic => File.ReadLines(fic).Count() > 0).ToList());

                if (doVerbs)
                    LogMessage(MessageSeverity.Information, "Concaténation du fichier '" + res + "' effectuée");
                

                // Gestion du code retour - En cas de succès
                outputParams.Add(OUTPUT_RETURNCODE, "0");
                outputParams.Add(OUTPUT_ERRORMESSAGE, String.Empty);
            }
            catch (Exception ex)
            {
                // Gestion du code retour - En cas d'erreur
                outputParams.Add(OUTPUT_RETURNCODE, "1");
                outputParams.Add(OUTPUT_ERRORMESSAGE, ex.Message);

                // Lancement d'une exception en cas d'erreur
                throw new Exception(outputParams[OUTPUT_ERRORMESSAGE]);
            }

            return outputParams;
            #endregion
        }
        #endregion        
    }
}
