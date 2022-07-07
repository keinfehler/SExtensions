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
using System.Threading;
using System.Windows.Forms;

namespace SExtensions
{
    

    public static class Helpers
    {
        public static Dictionary<string, Tuple<DocWrapper, int>> InternalUniqueOcurrences { get; set; } = new Dictionary<string, Tuple<DocWrapper, int>>();
        public static void FindOccurrencesAndExport()
        {
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
                var columnNames = new string[] { "Tipo Pieza", "Cantidad", "Código", "RV", "Nº Articulo", "Descripción", "Material", "Ref Comercial", "Nombre Archivo","Ruta" };

                var filteredData = InternalUniqueOcurrences.ToList();

                string[][] data = null;

                var test = filteredData.Select(o => new
                {
                    FileName = System.IO.Path.GetFileName(o.Key),
                    O = o.Value.Item1,
                    Qty = o.Value.Item2,
                    Path = o.Key

                }).ToList();

                data = test.Select(o => new[]
                                {
                                    o.O.Category,
                                    o.Qty.ToString(),
                                    o.O.DocumentNumber,
                                    o.O.RevisionNumber,
                                    o.O.Keywords,
                                    o.O.Title,
                                    o.O.Material,
                                    o.O.Comments.Trim(),
                                    o.FileName,
                                    o.Path

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


        static void UpdateDictionary(string lowerFileName, DocWrapper solidEdgeDocument)
        {
            lowerFileName = lowerFileName.ToLower();
            if (!InternalUniqueOcurrences.ContainsKey(lowerFileName))
                InternalUniqueOcurrences.Add(lowerFileName, Tuple.Create(solidEdgeDocument, 1));
            else
                InternalUniqueOcurrences[lowerFileName] = Tuple.Create(InternalUniqueOcurrences[lowerFileName].Item1, InternalUniqueOcurrences[lowerFileName].Item2 + 1);
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
                //Filter out certain occurrences.
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

                    if (solidEdgeDocument != null)
                    {
                        var summaryInfo = solidEdgeDocument.SummaryInfo as SummaryInfo;
                        var d = new DocWrapper(summaryInfo.Category, summaryInfo.DocumentNumber, summaryInfo.RevisionNumber, summaryInfo.Keywords, summaryInfo.Title, summaryInfo.Comments, GetMaterial(solidEdgeDocument));
                        UpdateDictionary(occurrence.OccurrenceFileName, d);
                    }
                  
                }

                if (doc is WeldmentDocument)
                {
                    var wdoc = doc as WeldmentDocument;
                    var vmodels = wdoc.WeldmentModels;
                    if (vmodels != null)
                    {
                        var i = vmodels.Item(1);
                        if (i != null)
                        {
                            var parts = i.PartModels;
                            

                            foreach (WeldPartModel part in parts)
                            {
                                var path = part.FileName;
                                
                                var wd = GetDocument(path);
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

        static DocWrapper GetDocument(string path)
        {
            DocWrapper d = null;
            try
            {

                SolidEdgeCommunity.Reader.SolidEdgeDocument parDocument = null;


                if (!InternalUniqueOcurrences.ContainsKey(path))
                {
                  
                    parDocument = SolidEdgeCommunity.Reader.SolidEdgeDocument.Open(path);
                    var summary = parDocument.SummaryInformation;
                    var docSummary = parDocument.DocumentSummaryInformation;

                    var material = parDocument?.MechanicalModeling?.Material;
                    var projectInformation = parDocument?.ProjectInformation;

                    var docNumber = projectInformation.DocumentNumber;
                    var revisionNumber = projectInformation.Revision;

                    d = new DocWrapper(docSummary.Category, docNumber, revisionNumber, summary.Keywords, summary.Title, summary.Comments, material);

                  
                    parDocument.Close();
                }
                else
                {
                    d = InternalUniqueOcurrences[path].Item1;
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
            catch (Exception)
            {
                return string.Empty;
                //MessageBox.Show(ex.Message);
            }


            return string.Empty;
        }

       
    }


}
