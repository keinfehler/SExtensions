using SolidEdgeAssembly;
using SolidEdgeCommunity.AddIn;
using SolidEdgeFramework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace SExtensions
{
    public static class ExtendedCommands
    {
        public static string OutputDirectory =>  Properties.Settings.Default.DirectorioSalida;
        public static void AddCopyPlus(bool allOccurrences)
        {
            //MessageBox.Show("Hola Mundo!");
            ProcessSelectedOccurrences(new ProcessOccurrenceWithNewModel(ReplaceOccurrence), allOccurrences);
        }

        private static bool _allOccurrences;
        private delegate void ProcessOccurrenceWithNewModel(Occurrence occurrence, string newModelFileName);

        private static void ReplaceOccurrence(Occurrence occurrenceToReplace, string newModelFileName)
        {
            occurrenceToReplace.Replace(newModelFileName, _allOccurrences, Type.Missing);
            //if (CSAHelpers.debug == CSAHelpers.UserYearID)
            //{
            //    int num = (int)MessageBox.Show("Nueva ocurrencia: '" + occurrenceToReplace.Name + "'");
            //}
            UpdateProperties(occurrenceToReplace);
            occurrenceToReplace.PutStyleNone();
        }
        private static void UpdateProperties(Occurrence occurrence)
        {
            try
            {
                SolidEdgeFramework.PropertySets properties = (SolidEdgeFramework.PropertySets)((SolidEdgeDocument)occurrence.OccurrenceDocument).Properties;
                //properties.ChangeCustomProperty("Autores", CSAHelpers.UserDraftName);
                //properties.ChangeCustomProperty("FechaPlano", DateTime.Now.ToString("dd/MM/yyyy"));
                properties.Save();
            }
            catch (Exception ex)
            {
                throw;
            }
        }
        private static void ProcessSelectedOccurrences(ProcessOccurrenceWithNewModel processOccurrenceWithNewModel, bool allOccurrences)
        {
            _allOccurrences = allOccurrences;
            var ocurrences = GroupSelectedOccurrences();
            foreach (IGrouping<SolidEdgeDocument, Occurrence> selectedOccurrence in ocurrences)
            {
                try
                {
                    string empty1 = string.Empty;
                    SolidEdgeDocument key = selectedOccurrence.Key;

                    //17.05.2022 - RE: Out
                    //if (CSAHelpers.debug == CSAHelpers.UserYearID)
                    //{
                    //    int num1 = (int)MessageBox.Show("Procesando ocurrencia: '" + key.Name + "'", "Debug Info", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    //}
                    //bool alternativePrefix = false;
                    //switch (key.Type)
                    //{
                    //    case SolidEdgeFramework.DocumentTypeConstants.igPartDocument:
                    //    case SolidEdgeFramework.DocumentTypeConstants.igSheetMetalDocument:
                    //        alternativePrefix = key.Name.StartsWith("Z");
                    //        break;
                    //    case SolidEdgeFramework.DocumentTypeConstants.igAssemblyDocument:
                    //        alternativePrefix = key.Name.StartsWith("M");
                    //        break;
                    //}
                    //string nameInUserFolder = CSAHelpers.GetNextFileNameInUserFolder(key.Type, alternativePrefix);


                    if (!Directory.Exists(OutputDirectory))
                        Directory.CreateDirectory(OutputDirectory);
                    
                    File.Copy(key.FullName, OutputDirectory);
                    //string draftFileName1 = CSAHelpers.GetDraftFileName(key.FullName);
                    //string draftFileName2 = CSAHelpers.GetDraftFileName(nameInUserFolder);
                    //if (key.Type == SolidEdgeFramework.DocumentTypeConstants.igAssemblyDocument)
                    //{
                    //    string cfgFileName1 = CSAHelpers.GetCfgFileName(key.FullName);
                    //    string cfgFileName2 = CSAHelpers.GetCfgFileName(nameInUserFolder);
                    //    if (File.Exists(cfgFileName1))
                    //        File.Copy(cfgFileName1, cfgFileName2);
                    //}
                    //if (File.Exists(draftFileName1))
                    //{
                    //    File.Copy(draftFileName1, draftFileName2);
                    //    CSACustomCommands.RedefineLinks(draftFileName2, key.FullName, nameInUserFolder);
                    //}
                    //if (CSAHelpers.debug == CSAHelpers.UserYearID)
                    //{
                    //    int num2 = (int)MessageBox.Show("Nuevos ficheros:" + System.Environment.NewLine + "Nuevo fichero: '" + nameInUserFolder + "'" + System.Environment.NewLine + "Nuevo fichero Plano: '" + draftFileName2 + "'" + System.Environment.NewLine, "Debug Info", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    //}
            //        string empty2 = string.Empty;
            //        switch (key.Type)
            //        {
            //            case SolidEdgeFramework.DocumentTypeConstants.igPartDocument:
            //                ((PartDocument)key).GetInContextAssemblyNameForInterpartLinks(ref empty2);
            //                break;
            //            case SolidEdgeFramework.DocumentTypeConstants.igAssemblyDocument:
            //                ((AssemblyDocument)key).GetInContextAssemblyNameForInterpartLinks(ref empty2);
            //                break;
            //            case SolidEdgeFramework.DocumentTypeConstants.igSheetMetalDocument:
            //                ((SheetMetalDocument)key).GetInContextAssemblyNameForInterpartLinks(ref empty2);
            //                break;
            //        }
            //        if (!string.IsNullOrEmpty(empty2))
            //            CSACustomCommands.RedefineLinks(nameInUserFolder, CSAHelpers.GetUNCPath(empty2), CSAHelpers.GetUNCPath(((SolidEdgeDocument)SolidEdgeAddIn.Instance.Application.ActiveDocument).FullName));
            //        foreach (Occurrence occurrence in (IEnumerable<Occurrence>)selectedOccurrence)
            //            processOccurrenceWithNewModel(occurrence, nameInUserFolder);
             }
              catch (Exception ex)
                {
                //        int num = (int)MessageBox.Show(ex.Message + System.Environment.NewLine + System.Environment.NewLine + ex.StackTrace);
                }
            }
            //if (CSACustomCommands.dmApp == null)
            //    return;
            //CSACustomCommands.dmApp.Quit();
            //CSACustomCommands.dmApp = (DesignManager.Application)null;
        }

        private static IEnumerable<IGrouping<SolidEdgeDocument, Occurrence>> GroupSelectedOccurrences()
        {
            SelectSet selectSet = ((SolidEdgeDocument)SolidEdgeAddIn.Instance.Application.ActiveDocument).SelectSet;
            List<Occurrence> source = new List<Occurrence>();
            bool flag1 = false;
            bool flag2 = false;
            //if (CSAHelpers.debug == CSAHelpers.UserYearID)
            //{
            //    int num1 = (int)MessageBox.Show(string.Format("{0} occurrencias seleccionadas", (object)selectSet.Count), "Debug Info", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            //}
            int num2;
            for (int i = 1; i <= selectSet.Count; i = num2 + 1)
            {
                if (!(selectSet.Item((object)i) is Occurrence))
                {
                    flag1 = true;
                }
                else
                {
                    //if (CSACustomCommands._onlyHardwareFiles)
                    //{
                    //    string withoutExtension = System.IO.Path.GetFileNameWithoutExtension(((Occurrence)selectSet.Item((object)i)).OccurrenceFileName);
                    //    int result;
                    //    if (!withoutExtension.StartsWith("C") || withoutExtension.Length != 9 || !int.TryParse(withoutExtension.Substring(1), out result))
                    //    {
                    //        flag2 = true;
                    //        goto label_11;
                    //    }
                    //}
                    if ((uint)source.Where<Occurrence>((Func<Occurrence, bool>)(x => x.OccurrenceDocument == ((Occurrence)selectSet.Item((object)i)).OccurrenceDocument)).Count<Occurrence>() <= 0U || !_allOccurrences)
                        source.Add((Occurrence)selectSet.Item((object)i));
                }
            label_11:
                num2 = i;
            }
            if (flag1)
            {
                int num3 = (int)MessageBox.Show("Solo se procesarán las ocurrencias de primer nivel", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            if (flag2)
            {
                int num4 = (int)MessageBox.Show("Solo se procesarán los Comerciales", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            return source.GroupBy<Occurrence, SolidEdgeDocument>((Func<Occurrence, SolidEdgeDocument>)(ocu => (SolidEdgeDocument)ocu.OccurrenceDocument));
        }
    }
}
