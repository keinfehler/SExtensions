using Helpers;
using SolidEdgeAssembly;
using SolidEdgeCommunity.AddIn;
using SolidEdgeCommunity.Extensions;
using SolidEdgeConstants;
using SolidEdgeFramework;
using SolidEdgePart;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
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

        internal static void SaveASM()
        {
 
            SolidEdgeFramework.Application seApp = (SolidEdgeFramework.Application)Marshal.GetActiveObject("SolidEdge.Application");
            SolidEdgeFramework.Application objApp = (SolidEdgeFramework.Application)Marshal.GetActiveObject("SolidEdge.Application");
            seApp.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)SolidEdgeConstants.AssemblyCommandConstants.AssemblyAssemblyToolsHideAllReferencePlanes);
            objApp.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)40080);
            objApp.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)40081);
            objApp.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)40082);
            objApp.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)40083);
            seApp.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)SolidEdgeConstants.AssemblyCommandConstants.AssemblyViewISOView);
            seApp.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)SolidEdgeConstants.AssemblyCommandConstants.AssemblyViewZoom);
            seApp.StartCommand((SolidEdgeFramework.SolidEdgeCommandConstants)SolidEdgeConstants.AssemblyCommandConstants.AssemblyFileSave);

        }

        internal static void ReplaceUsingNewCopy()
        {
            var newNameForm = new FormRename();
            newNameForm.CreateAndReplace = true;
            newNameForm.ShowDialog();
        }

        internal static void OpenDraftFiles()
        {
            DesignManagerHelpers.OpenDraftFiles();
        }

        internal static void ReplacePWD()
        {
            DesignManagerHelpers.ReplacePWD();
        }
    }
}
