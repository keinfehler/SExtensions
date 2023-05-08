using Helpers;
using SolidEdgeAssembly;
using SolidEdgeCommunity.AddIn;
using SolidEdgeFramework;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace DesignManager
{
    public static class ExtendedCommands
    {
        public static string OutputDirectory => @"C/:DirectorioDeSalida";
        public static void AddCopyPlus(bool allOccurrences)
        {
            ReplaceAndCopy(allOccurrences);
        }
        private static void ReplaceAndCopy(bool allOccurrences)
        {
            Helpers.DesignManagerHelpers.ReplaceAndCopy(Helpers.DesignManagerHelpers.ActiveDocument, allOccurrences);
        }
        public static void ReplaceAndCopyWithRevision(bool allOccurrences)
        {
            Helpers.DesignManagerHelpers.ReplaceAndCopy(Helpers.DesignManagerHelpers.ActiveDocument, allOccurrences, revision:true);
        }

        public static void RenameSelectedOcurrences()
        {
            var newNameForm = new FormRename();
            newNameForm.ShowDialog();
            
        }
        public static void CreateFromSelectedOcurrence()
        {
            var newNameForm = new FormRename();
            newNameForm.CreateCopy = true;
            newNameForm.ShowDialog();

        }

        internal static void OpenDirectory()
        {
            var activeDocument = SolidEdgeAddIn.Instance.Application.ActiveDocument as SolidEdgeDocument;

            var activeDocumentFullName = activeDocument.FullName;
            var activeDocumentDirectoryName = System.IO.Path.GetDirectoryName(activeDocumentFullName);

            Process.Start(activeDocumentDirectoryName);

        }

        internal static void Readonly()
        {
            var activeDocument = SolidEdgeAddIn.Instance.Application.ActiveDocument as SolidEdgeDocument;
         
            activeDocument.ReadOnly = true;


        }

        internal static void OpenDraftFiles()
        {
            DesignManagerHelpers.OpenDraftFiles();
        }
    }
}
