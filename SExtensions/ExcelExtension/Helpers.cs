using ClosedXML.Excel;
using SolidEdgeAssembly;
using SolidEdgeCommunity.AddIn;
using SolidEdgeCommunity.Extensions; // https://github.com/SolidEdgeCommunity/SolidEdge.Community/wiki/Using-Extension-Methods
using SolidEdgeFramework;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace SExtensions
{
    public static class Helpers
    {
        public static Dictionary<string, Tuple<Occurrence, int>> InternalUniqueOcurrences { get; set; } =  new Dictionary<string, Tuple<Occurrence, int>>();
        public static void FindOccurrencesAndExport()
        {
            var start = DateTime.Now;
            InternalUniqueOcurrences.Clear();
            OutputPath = null;

            var app = SolidEdgeAddIn.Instance.Application;
            var assemblyDocument = app.GetActiveDocument<AssemblyDocument>();

            if (assemblyDocument != null)
            {
                //get occurrences and fill into a dictionary with a counter
                FillOccurrence(assemblyDocument);

                //get occurrences from dictionary and write in a excel file
                ExportOccurrences(System.IO.Path.GetFileNameWithoutExtension(assemblyDocument.FullName));

                var end = DateTime.Now;

                //MessageBox.Show("Fin" + end.Subtract(start).ToString());

                //try
                //{
                //    ProcessStartInfo startInfo = new ProcessStartInfo();
                //    startInfo.FileName = "EXCEL.EXE";
                //    startInfo.Arguments = OutputPath;
                //    Process.Start(startInfo);
                //}
                //catch (Exception ex)
                //{

                //    MessageBox.Show(ex.Message);
                //}
                Process.Start(OutputPath);
                End();

                //Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                //Microsoft.Office.Interop.Excel.Workbook wb = excel.Workbooks.Open(OutputPath);
            }
                
            else
                MessageBox.Show("Not a assembly document opened");
        }
        public static string OutputPath { get; set; }
        private static void ExportOccurrences(string documentName)
        {
            try
            {
                var wbook = new XLWorkbook();
                wbook.AddWorksheet("1");

                var ws = wbook.Worksheet("1");
                var columnNames = new string[] { "Tipo Pieza", "Cantidad", "Código", "RV", "Nº Articulo", "Descripción", "Material", "Ref Comercial", "Nombre Archivo" };

                var data = InternalUniqueOcurrences
                            .Where(o => !o.Value.Item1.FileMissing())
                            .Where(o => o.Value.Item1 is Occurrence)
                            .Select(o => new { FileName = System.IO.Path.GetFileName(o.Value.Item1.OccurrenceFileName), D = o.Value.Item1.OccurrenceDocument as SolidEdgeDocument, O = ((SolidEdgeDocument)o.Value.Item1.OccurrenceDocument).SummaryInfo as SummaryInfo, Qty = o.Value.Item2 })
                            .Select(o => new[] { o.O.Category, o.Qty.ToString(), o.O.DocumentNumber, o.O.RevisionNumber, o.O.Keywords, o.O.Title, GetMaterial(o.D), o.O.Comments.Trim(), o.FileName })
                            .ToArray();

                int col = 1;
                foreach (var item in columnNames)
                {
                    ws.Cell(1, col).Value = item;
                    int row = 2;
                    foreach (var o in data)
                    {
                        ws.Cell(row, col).Value = o[col - 1];
                        row++;
                    }
                    col++;
                }

                var outPutPath = System.IO.Path.Combine(Properties.Settings.Default.DirectorioSalida, documentName + ".xlsx");
                OutputPath = outPutPath;


                if (File.Exists(outPutPath))
                {
                    try
                    {
                        File.Delete(outPutPath);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                        End();

                    }finally
                    {
                        
                    }
                }

                wbook.SaveAs(outPutPath);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                End();

            }
            finally
            {
                
            }
        }
        static void End()
        {
            InternalUniqueOcurrences.Clear();
            OutputPath = null;
        }
       

        static void FillOccurrence(AssemblyDocument assemblyDocument)
        {
            if (assemblyDocument == null)
                return;


            foreach (Occurrence occurrence in assemblyDocument.Occurrences)
            {
                // Filter out certain occurrences.
                if (!occurrence.IncludeInBom) { continue; }
                if (occurrence.IsPatternItem) { continue; }
                if (occurrence.OccurrenceDocument == null) { continue; }

                // To make sure nothing silly happens with our dictionary key, force the file path to lowercase.
                var lowerFileName = occurrence.OccurrenceFileName.ToLower();

                // If the dictionary does not already contain the occurrence, add it.
                if (!InternalUniqueOcurrences.ContainsKey(lowerFileName))
                    InternalUniqueOcurrences.Add(lowerFileName, Tuple.Create(occurrence, 1));
                else
                    InternalUniqueOcurrences[lowerFileName] = Tuple.Create(InternalUniqueOcurrences[lowerFileName].Item1, InternalUniqueOcurrences[lowerFileName].Item2 + 1);

                if (occurrence.Subassembly)
                    FillOccurrence(occurrence.OccurrenceDocument as AssemblyDocument);
            }
        }

        static string GetMaterial(SolidEdgeDocument document)
        {
            PropertySets pps = null;

            pps = document.Properties as SolidEdgeFramework.PropertySets;

            if (pps != null)
            {
                SolidEdgeFramework.Properties mproperties = null;

                mproperties = pps.Item(7);

                if (mproperties != null)
                {
                    Property p = null;
                    p = mproperties.Item(1);


                    if (p != null)
                    {
                        var m = p.get_Value();
                        return m?.ToString();
                    }
                }
            }
            return string.Empty;
        }
    }
}
