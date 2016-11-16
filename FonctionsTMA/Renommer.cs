

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
    /// Plugin de renommage des fichiers horodatés
    /// </summary>
    public class Renommer : DataPorterPlugin
    {
        #region Constants

        /// <summary>
        /// Option permettant d'effacer les fichiers sources
        /// </summary>
        private const string PRM_DELETE_SRC = "Flag indiquant si les fichiers horodatés doivent être effacés";
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

                inputs.Add(Concatener.PRM_REP_IN, "Mandatory - Répertoire d'entré");
                inputs.Add(Concatener.PRM_EXTENSION, "Optional - Type des fichiers traités. CSV par défaut");
                inputs.Add(Concatener.PRM_VERBOSE, "Optional - Caractère permettant d'activer les logs : 'Y' -> Active l'affichage");

                inputs.Add(Concatener.PRM_INCLURE_SS_REP, "Optional - Caractère indiquant si les sous-dossiers doivent être scrutés : 'Y' -> Inclure les sous-dossiers");
                inputs.Add(PRM_DELETE_SRC, "Optional - Flag indiquant si les fichiers horodatés doivent être effacés : 'Y' -> Efface les sources après renommage");

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
            get { return "Renommage des fichiers horodatés."; }
        }
        #endregion

        #region Public Methods

        /// <summary>
        /// Paramètres de sortie du plugin Markit
        /// </summary>
        /// <returns>Renvoi le dictionnaire des paramètre de sortie</returns>
        public Dictionary<string, string> GetOutputParameters()
        {

            Dictionary<string, string> outputs = new Dictionary<string, string>();
            // Initialisation des paramètres de sortie Markit
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
            string extension = String.IsNullOrEmpty(inputParameters[Concatener.PRM_EXTENSION].Trim()) ? ".csv" : inputParameters[Concatener.PRM_EXTENSION].Trim();
            extension = extension.StartsWith(".", StringComparison.Ordinal) ? extension : "." + extension;           
            SearchOption optionRecherche = inputParameters[Concatener.PRM_INCLURE_SS_REP].Trim() == "Y" ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
            bool doVerbs = inputParameters[Concatener.PRM_VERBOSE].Trim() == "Y" ? true : false;
            bool effacerSRC = inputParameters[PRM_DELETE_SRC].Trim() == "Y" ? true : false;
            #endregion

            #region Invocation de la fonction  

            try
            {
                // Initialisation de la classe de gestion des fichiers
                if (doVerbs)
                    LogMessage(MessageSeverity.Information, "Initialisation de la fonction de renommage");
                FileHandler gestFic = new FileHandler(monRepIn, optionRecherche, extension);

                // Récupération des fichiers dê même type
                if (doVerbs)
                    LogMessage(MessageSeverity.Information, "Récupération des fichiers *" + extension);
                List<string> files2Process = gestFic.RetrieveComonFiles();

                // Si aucun fichier horodaté est trouvé alors on lance une exeption
                if (files2Process == null || files2Process.FirstOrDefault(fi => Regex.IsMatch(fi, Concatener.timePattern + extension + "$")) == null)
                    throw new Exception("Aucun fichier horodaté a été trouvé");

                // Initialistion des fichiers horodatés
                files2Process = files2Process.Where(fi => Regex.IsMatch(fi, Concatener.timePattern + extension + "$")).ToList();

                // Affichage des fichiers à renommer
                if (doVerbs)
                {
                    LogMessage(MessageSeverity.Information, "Liste des fichiers à renommer : ");
                    files2Process.ForEach(
                                                fi =>
                                                        {
                                                            LogMessage(MessageSeverity.Information, new FileInfo(fi).Name);
                                                        }
                                          );
                }

                List<string> oldOnes = null;

                // Sauvegarde des chemin de fichier
                if (effacerSRC)
                    oldOnes = files2Process;

                // Renommage des fichiers
                if (doVerbs)
                    LogMessage(MessageSeverity.Information, "Renommage des fichiers trouvés");
                files2Process = gestFic.RenameTimeSpanFiles(files2Process);

                // Effacement des fichiers horodatés si demandé
                if (effacerSRC)
                    oldOnes.ForEach(
                                        File.Delete
                                   );

                // Affichage de la liste des fichiers renommés
                if (doVerbs)
                {
                    LogMessage(MessageSeverity.Information, "Liste des fichiers renommés : ");
                    files2Process.ForEach(
                                                fic =>
                                                        {
                                                            LogMessage(MessageSeverity.Information, new FileInfo(fic).Name);
                                                        }
                                          );
                }

                if (doVerbs)
                    LogMessage(MessageSeverity.Information, "Renommage des fichiers terminé");

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
