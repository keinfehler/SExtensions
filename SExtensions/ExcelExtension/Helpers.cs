using ClosedXML.Excel;
using SolidEdgeAssembly;
using SolidEdgeCommunity.AddIn;
using SolidEdgeCommunity.Extensions; // https://github.com/SolidEdgeCommunity/SolidEdge.Community/wiki/Using-Extension-Methods
using SolidEdgeFramework;
using SolidEdgePart;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

namespace SExtensions
{

    public static class ConfigurationHelper
    {
        private static System.Configuration.Configuration appConfig = ConfigurationManager.OpenExeConfiguration(Assembly.GetExecutingAssembly().Location);
        public static string GetConfigurationValue(string key)
        {
            return appConfig.AppSettings.Settings[key].Value;
        }
    }

    public static class Helpers
    {
        
        
        public static string rutaUtillaje = ConfigurationHelper.GetConfigurationValue("RutaUtillaje");
        public static string directorioSalida = ConfigurationHelper.GetConfigurationValue("DirectorioSalida");
        public static string rutaTemporal = ConfigurationHelper.GetConfigurationValue("RutaTemporal");

        public static Dictionary<string, Tuple<DocWrapper, int>> InternalUniqueOcurrences { get; set; } = new Dictionary<string, Tuple<DocWrapper, int>>();
        public static void FindOccurrencesAndExport(bool getPwdFiles,  bool rutas, bool utillaje, string rutaUtillajeOutput = null)
        {
            InternalUniqueOcurrences.Clear();
            OutputPath = null;

            var app = SolidEdgeAddIn.Instance.Application;
            var assemblyDocument = app.GetActiveDocument<AssemblyDocument>();

            if (assemblyDocument != null)
            {

                FillOccurrence(assemblyDocument, getPwdFiles, 0);

                var fileName = System.IO.Path.GetFileNameWithoutExtension(assemblyDocument.FullName);

                if (utillaje)
                {
                    rutaUtillaje = rutaUtillajeOutput;
                    //update a template .xlsm
                    ExportOccurrences(fileName, rutaUtillaje, "*RELACION UTILLAJE.xlsm", "X-HOJA A COPIAR", false, (rutas && utillaje));
                }
                else
                {
                    //to new .xlsx File
                    ExportOccurrences(fileName, directorioSalida);

                    if (OutputPath == null)
                        return;

                    Process.Start(OutputPath);
                }
                
                
              

                
                End();
            }

            else
                MessageBox.Show("Not a assembly document opened");


        }
        public static string OutputPath { get; set; }
        private static void ExportOccurrences(string documentName,  string outputPath, string documentPatern = null, string sheetName = null, bool header = true, bool rutas = true)
        {
            try
            {


                if (outputPath != null && documentPatern != null)
                {
                    var file = Directory.GetFiles(outputPath, documentPatern).FirstOrDefault();
                    if (file != null)
                    {
                        var destinationDirectory = System.IO.Path.GetDirectoryName(file);

                        var tmpFileName = System.IO.Path.Combine(rutaTemporal, System.IO.Path.GetFileNameWithoutExtension(file) /*+ "_TMP"*/ + System.IO.Path.GetExtension(file));


                        var backupFileName = System.IO.Path.Combine(rutaTemporal, System.IO.Path.GetFileNameWithoutExtension(file) + "_BACKUP_" + DateTime.Now.ToString("yyyyMMddHHmmss") + System.IO.Path.GetExtension(file));

                        try
                        {
                            if (!Directory.Exists(rutaTemporal))
                                Directory.CreateDirectory(rutaTemporal);
                            

                            if (File.Exists(tmpFileName))
                                File.Delete(tmpFileName);

                            File.Copy(file, tmpFileName);

                            
                            var wbookTmp = new XLWorkbook(tmpFileName);

                            if (wbookTmp != null)
                            {
                                IXLWorksheet workSheet = wbookTmp.Worksheets.FirstOrDefault(o => o.Name.Contains(sheetName));

                                if (workSheet != null)
                                {
                                    var newName = ConvertDocumentName(documentName);

                                    IXLWorksheet repeatedWorkSheet = null;
                                    if(wbookTmp.TryGetWorksheet(newName, out repeatedWorkSheet)) 
                                        repeatedWorkSheet.Delete();
                                    

                                    var newWorkSheet = workSheet.CopyTo(wbookTmp, newName);

                                    FillWorksheetData(newWorkSheet, false, rutas, true, rutasCheckbox:true);

                                    wbookTmp.Save();

                                    File.Replace(tmpFileName, file, backupFileName);

                                    Process.Start(file);
                                }
                            }

                           

                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }
                    }
                    return;
                }
               


                var wbook = new XLWorkbook();
                wbook.AddWorksheet("1");

                var ws = wbook.Worksheet("1");

                FillWorksheetData(ws, header, rutasCheckbox:rutas);

                var outPutPath = System.IO.Path.Combine(outputPath, documentName + ".xlsx");
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
        private static Tuple<int?, int, string> GetTuple(int? row, int column, string value)
        {
            return Tuple.Create(row, column, value);

        }
        private static void FillWorksheetData(IXLWorksheet ws, bool header, bool ruta = true, bool exportHeader = false, bool rutasCheckbox = true)
        {
            Tuple<int?, int, string>[] columnNames = null;
            var columnNamesList = new List<Tuple<int?, int, string>> 
            { 
               GetTuple(5,1, "Tipo Pieza"),
               GetTuple(5,2, "Cantidad"),
               GetTuple(5,3,"Código"),
               GetTuple(5,4, "RV"),
               GetTuple(5,5, "Nº Articulo"), 
               GetTuple(5,6, "Descripción"),
               GetTuple(5,7, "Material"),
               GetTuple(5,8, "Ref Comercial"),
               GetTuple(5,9, "Nombre Archivo" )
            };

            if (ruta)
            {
                columnNamesList.Add(GetTuple(5, rutasCheckbox ? 25 : 10, "Ruta 2D"));
                columnNamesList.Add(GetTuple(5, rutasCheckbox ? 26 : 10, "Ruta 3D"));
            }
            columnNames = columnNamesList.ToArray();

            var filteredData = InternalUniqueOcurrences.ToList();

            Tuple<int?, int, string>[][] data = null;

            var test = filteredData.Select(o => new
            {
                FileName = System.IO.Path.GetFileName(o.Key),
                O = o.Value.Item1,
                Qty = o.Value.Item2,
                Path = o.Key

            }).ToList();



            data = test.Select(o => new Tuple<int?, int, string>[]
                                {
                                    GetTuple(null, 1, o.O.Category),
                                    GetTuple(null, 2, o.Qty.ToString()),
                                    GetTuple(null, 3, o.O.DocumentNumber),
                                    GetTuple(null, 4, o.O.RevisionNumber),
                                    GetTuple(null, 5, o.O.Keywords),
                                    GetTuple(null, 6, o.O.Title),
                                    GetTuple(null, 7, o.O.Material),
                                    GetTuple(null, 8, o.O.Comments.Trim()),
                                    GetTuple(null, 9, o.FileName),
                                    GetTuple(null, rutasCheckbox ? 25 : 10, ruta ? System.IO.Path.ChangeExtension(o.Path, ".dft") : ""),
                                    GetTuple(null, rutasCheckbox ? 26 : 10, ruta ? o.Path : "")


                                }).ToArray();


            if (data == null)
            {
                return;
            }
            if (exportHeader)
            {
                ws.Cell(1, 2).Value = Head.Empresa;
                ws.Cell(2, 2).Value = Head.Maquina;
                ws.Cell(3, 2).Value = Head.Modelo;
                ws.Cell(1, 6).Value = Head.Title;
                ws.Cell(2, 6).Value = Head.NombreArchivo;
            }

            //int col = 1;
            foreach (var item in columnNames)
            {
                if (header)
                {
                    ws.Cell(item.Item1.Value, item.Item2).Value = item.Item3;
                }

                int row = 6;
                foreach (var d in data)
                {
                    foreach (var r in d)
                    {
                        ws.Cell(row, r.Item2).Value = r.Item3;
                    }
                    
                    row++;
                }
                //col++;
            }


        }
        private static string ConvertDocumentName(string documentName)
        {
            if (documentName.Length > 30)
            {
                if (documentName.StartsWith("SUBCONJUNTO"))
                {
                    documentName = documentName.Substring(12, 18);
                }
                else if (documentName.StartsWith("CONJUNTO"))
                {
                    documentName = documentName.Substring(8, 21);

                }else
                {
                    documentName = documentName.Substring(0, 30);
                }
                
            }
            return documentName;
        }

        static void End()
        {
            InternalUniqueOcurrences.Clear();
            OutputPath = null;
        }


        static void UpdateDictionary(string lowerFileName, DocWrapper solidEdgeDocument, int level)
        {
            lowerFileName = lowerFileName.ToUpper();

            //Console.WriteLine($"Nivel: {level} -- {lowerFileName}");
            if (!InternalUniqueOcurrences.ContainsKey(lowerFileName))
                InternalUniqueOcurrences.Add(lowerFileName, Tuple.Create(solidEdgeDocument, 1));
            else
                InternalUniqueOcurrences[lowerFileName] = Tuple.Create(InternalUniqueOcurrences[lowerFileName].Item1, InternalUniqueOcurrences[lowerFileName].Item2 + 1);
        }
        static string GetCustomProperty(this AssemblyDocument obj, string propertyName)
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
                var p = document.Properties;
                var propertySets = (SolidEdgeFramework.PropertySets)p;

                var customProperties = propertySets.Item(4);

                var items = customProperties.OfType<SolidEdgeFramework.Property>();

                var property = items.FirstOrDefault(o => o.Name == propertyName);

                if (property != null)
                {
                    return property.get_Value()?.ToString() ?? string.Empty;
                }
            
            }
            catch (Exception)
            {
                return string.Empty;
                //MessageBox.Show(ex.Message);
            }


            return string.Empty;
        }
        public static DocWrapper Head { get; set; }

        private static SummaryInfo GetSummaryInfoPropertyValue(this AssemblyDocument doc)
        {
            return doc.GetSummaryInfo();
        }
        
        
        public static void FillOccurrence(AssemblyDocument assemblyDocument, bool getPwdFiles, int level)
        {
            if (assemblyDocument == null)
                return;
   
            Head = null;
            Head = new DocWrapper();
            Head.Modelo = assemblyDocument.GetCustomProperty("MODELO");
            Head.Maquina = assemblyDocument.GetCustomProperty("MAQUINA");
            Head.Empresa = assemblyDocument.GetSummaryInfoPropertyValue().Company;
            Head.NombreArchivo = assemblyDocument.Name;
            Head.Title = assemblyDocument.GetSummaryInfoPropertyValue().Title;
            Head.IsHeader = true;

            foreach (var occ in assemblyDocument.Occurrences)
            {
                level++;

                Occurrence occurrence = null;
                
                if (occ is Occurrence)
                    occurrence = occ as Occurrence;
                
                if (occ == null || occurrence.FileMissing() || !occurrence.IncludeInBom /*|| occurrence.IsPatternItem*/ || occurrence.OccurrenceDocument == null) 
                    continue;

                var doc = occurrence.OccurrenceDocument;

                var name = occurrence.OccurrenceFileName;

                bool includeAsmParent = false;
                if (name.EndsWith(".asm"))
                {
                    var fileName = System.IO.Path.GetFileName(name);
                    if (fileName.StartsWith("SUBCONJUNTO SOLDADO") || fileName.StartsWith("SUBCONJUNTO MONTADO"))
                    {
                        includeAsmParent = true;
                        FillParents(occ, level);
                    }
                }
                else
                {
                    FillParents(occ, level);
                } 
                


                if (doc is WeldmentDocument)
                {

                    if (getPwdFiles)
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
                                    UpdateDictionary(path, wd, level);
                                }
                            }
                        }
                    }
                }

                //if (doc is SheetMetalDocument)
                //{
                //    var sheetMetalDocument = doc as SheetMetalDocument;
                //    if (sheetMetalDocument != null)
                //    {
                        
                //        var models = sheetMetalDocument.Models;
                //        for (int i = 1; i <= models.Count; i++)
                //        {
                //            var body = models.Item(i);

                            
                //            if (body != null)
                //            {
                               
                //                //Console.WriteLine("body '{0}'.", body.);
                //            }
                            

                //            //    var bodyDoc = body.Document as SolidEdgeFramework.SolidEdgeDocument;
                //            //    if (bodyDoc != null)
                //            //    {

                //            //        var propertySets = bodyDoc.Properties as SolidEdgeFramework.PropertySets;
                //            //        foreach (var properties in propertySets.OfType<SolidEdgeFramework.Properties>())
                //            //        {
                //            //            Console.WriteLine("PropertSet '{0}'.", properties.Name);

                //            //            foreach (var property in properties.OfType<SolidEdgeFramework.Property>())
                //            //            {
                //            //                System.Runtime.InteropServices.VarEnum nativePropertyType = System.Runtime.InteropServices.VarEnum.VT_EMPTY;
                //            //                Type runtimePropertyType = null;

                //            //                object value = null;

                //            //                nativePropertyType = (System.Runtime.InteropServices.VarEnum)property.Type;

                //            //                // Accessing Value property may throw an exception...
                //            //                try
                //            //                {
                //            //                    value = property.get_Value();
                //            //                }
                //            //                catch (System.Exception ex)
                //            //                {
                //            //                    value = ex.Message;
                //            //                }

                //            //                if (value != null)
                //            //                {
                //            //                    runtimePropertyType = value.GetType();
                //            //                }

                //            //                Console.WriteLine("\t{0} = '{1}' ({2} | {3}).", property.Name, value, nativePropertyType, runtimePropertyType);
                //            //            }

                //            //            Console.WriteLine();

                //            //    }
                //            //}
                //        }
                //    }
                //}


                if (occurrence.Subassembly)
                {
                    if (includeAsmParent)
                    {
                        if (getPwdFiles)
                        {
                            FillOccurrence(occurrence.OccurrenceDocument as AssemblyDocument, getPwdFiles, level);
                        }
                    }
                    else
                    {
                        FillOccurrence(occurrence.OccurrenceDocument as AssemblyDocument, getPwdFiles, level);
                    }
                }
                    

            }
        }

        static void FillParents(object occ, int level)
        {
            Occurrence occurrence = null;

            if (occ is Occurrence)
                occurrence = occ as Occurrence;

            var doc = occurrence.OccurrenceDocument;

            if (doc is SolidEdgeDocument)
            {
                var solidEdgeDocument = doc as SolidEdgeDocument;

                if (solidEdgeDocument != null)
                {
                    var summaryInfo = solidEdgeDocument.SummaryInfo as SummaryInfo;


                    var d = new DocWrapper(summaryInfo.Category, summaryInfo.DocumentNumber, summaryInfo.RevisionNumber, summaryInfo.Keywords, summaryInfo.Title, summaryInfo.Comments, GetMaterial(solidEdgeDocument));
                    UpdateDictionary(occurrence.OccurrenceFileName, d, level);
                }

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
