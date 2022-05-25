using ClosedXML.Excel;
using SolidEdgeAssembly;
using SolidEdgeCommunity.AddIn;
using SolidEdgeCommunity.Extensions; // https://github.com/SolidEdgeCommunity/SolidEdge.Community/wiki/Using-Extension-Methods

using SolidEdgeFramework;
using SolidEdgePart;
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
        public static Dictionary<string, Tuple<SolidEdgeDocument, int, string>> InternalUniqueOcurrences { get; set; } = new Dictionary<string, Tuple<SolidEdgeDocument, int, string>>();
        public static void FindOccurrencesAndExport()
        {
            var start = DateTime.Now;
            InternalUniqueOcurrences.Clear();
            OutputPath = null;

            var app = SolidEdgeAddIn.Instance.Application;
            var assemblyDocument = app.GetActiveDocument<AssemblyDocument>();

            if (assemblyDocument != null)
            {

                FillOccurrence(assemblyDocument);

                ExportOccurrences(System.IO.Path.GetFileNameWithoutExtension(assemblyDocument.FullName));

                if (OutputPath == null)
                    return;

                Process.Start(OutputPath);
                End();
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

                var filteredData = InternalUniqueOcurrences.ToList();

                string[][] data = null;

                var test = filteredData.Select(o => new
                {
                    FileName = System.IO.Path.GetFileName(o.Value.Item1.FullName),
                    Material = o.Value.Item3,
                    O = o.Value.Item1.SummaryInfo as SummaryInfo,
                    Qty = o.Value.Item2

                }).ToList();

                data = test.Select(o => new[]
                                {
                                    o.O.Category,
                                    o.Qty.ToString(),
                                    o.O.DocumentNumber,
                                    o.O.RevisionNumber,
                                    o.O.Keywords,
                                    o.O.Title,
                                    o.Material,
                                    o.O.Comments.Trim(),
                                    o.FileName

                                }).ToArray();


                if (data == null)
                {
                    return;
                }

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

                    }
                    finally
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


        static void UpdateDictionary(string lowerFileName, SolidEdgeDocument solidEdgeDocument)
        {
            lowerFileName = lowerFileName.ToLower();
            if (!InternalUniqueOcurrences.ContainsKey(lowerFileName))
                InternalUniqueOcurrences.Add(lowerFileName, Tuple.Create(solidEdgeDocument, 1, GetMaterial(solidEdgeDocument)));
            else
                InternalUniqueOcurrences[lowerFileName] = Tuple.Create(InternalUniqueOcurrences[lowerFileName].Item1, InternalUniqueOcurrences[lowerFileName].Item2 + 1, InternalUniqueOcurrences[lowerFileName].Item3);
        }

        static void FillOccurrence(AssemblyDocument assemblyDocument)
        {
            if (assemblyDocument == null)
                return;


            foreach (var occ in assemblyDocument.Occurrences)
            {
                Occurrence occurrence = null;

                if (occ is Occurrence)
                {
                    occurrence = occ as Occurrence;
                }

                if (occ == null) 
                {
                    continue; 
                }


                if (occurrence.FileMissing())
                {
                    continue;
                }
                // Filter out certain occurrences.
                if (!occurrence.IncludeInBom)
                {
                    continue;
                }
                if (occurrence.IsPatternItem)
                {
                    continue;
                }
                if (occurrence.OccurrenceDocument == null)
                {
                    continue;
                }

                var doc = occurrence.OccurrenceDocument;

                if (doc is SolidEdgeDocument)
                {
                    var solidEdgeDocument = doc as SolidEdgeDocument;
                    UpdateDictionary(occurrence.OccurrenceFileName, solidEdgeDocument);
                }

                if (doc is WeldmentDocument)
                {
                    var wdoc = doc as WeldmentDocument;
                    var vmodels = wdoc.WeldmentModels as WeldmentModels;
                    if (vmodels != null)
                    {
                        var i = vmodels.Item(1);
                        if (i != null)
                        {
                            var parts = i.PartModels;

                            foreach (WeldPartModel part in parts)
                            {
                                var path = part.FileName;
                                var wd = GetDocument(path, assemblyDocument);
                                if (wd == null)
                                {
                                    continue;
                                }
                                UpdateDictionary(path, wd);
                            }


                        }
                    }
                }

                if (occurrence.Subassembly)
                    FillOccurrence(occurrence.OccurrenceDocument as AssemblyDocument);

            }
        }

        static SolidEdgeDocument GetDocument(string path, AssemblyDocument assembly)
        {
            SolidEdgeDocument d = null;
            try
            {
                var occurrence = assembly.Occurrences.OfType<Occurrence>().FirstOrDefault(o => o.OccurrenceFileName == path);
                if (occurrence != null)
                {
                    d = occurrence.OccurrenceDocument as SolidEdgeDocument;
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }

            return d;
        }

        static string GetMaterial(object obj)
        {

            SolidEdgeDocument document = null;

            if (obj is SolidEdgeDocument)
            {
                document = obj as SolidEdgeDocument;
            }

            if (document == null)
            {
                return string.Empty;
            }

            try
            {
                PropertySets pps = null;

                pps = document.Properties as PropertySets;

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
                            var m = p?.get_Value();
                            return m?.ToString();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                return string.Empty;
                //MessageBox.Show(ex.Message);
            }


            return string.Empty;
        }
    }
}
