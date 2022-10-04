using SolidEdgeAssembly;
using SolidEdgeCommunity.Extensions;
using SolidEdgeFramework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Path = System.IO.Path;

namespace Helpers
{
    public static class DesignManagerHelpers
    {
        public static string GetNewName(string relatedItemfileNameWithoutExtension)
        {
            var splittedFileName = relatedItemfileNameWithoutExtension.Split('-').ToList();

            if (splittedFileName.Count > 1)
            {
                var lastItem = splittedFileName.LastOrDefault();

                splittedFileName.RemoveAt(splittedFileName.Count-1);

            }
            splittedFileName.Add("-00");

            var result = string.Empty;
            foreach (var c in splittedFileName)
            {
                result += c;
            }

            return result;
        }

        public static List<Occurrence> GetSelectedOccurrences(SolidEdgeDocument doc, bool all)
        {

            
            if (doc != null)
            {
                var selectedItems = doc.SelectSet.OfType<Occurrence>().ToList();
                return selectedItems;

            }
            return null;
        }
        public static void ReplaceAndCopy(SolidEdgeDocument document, bool allOccurrences)
        {

            var activeDocumentFullName = document.FullName;
            var activeDocumentDirectoryName = Path.GetDirectoryName(activeDocumentFullName);
            Console.WriteLine($"ACTIVEDOCUMENT:{activeDocumentFullName}");
            Console.WriteLine($"ACTIVEDOCUMENT_DIRECTORY:{activeDocumentDirectoryName}");
            Console.WriteLine($"----------------");
            Console.WriteLine($"----------------");
            Console.WriteLine($"----------------");
            Console.WriteLine($"----------------");


            foreach (var item in Helpers.DesignManagerHelpers.GetSelectedOccurrences(document, true))
            {
                Console.WriteLine($"----------------");
                Console.WriteLine($"----------------");

                var filePath = item.OccurrenceFileName;
                Console.WriteLine($"-OCCURRENCE: {filePath}");

                var fileName = Path.GetFileName(filePath);
                Console.WriteLine($"-OCCURRENCE-FILE: {fileName}");


                var directoryName = Path.GetDirectoryName(filePath);
                Console.WriteLine($"-DIRECTORY: {directoryName}");


                var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(filePath);

                var relatedItems = Directory.GetFiles(directoryName, $"{fileNameWithoutExtension}*");


                var replaceFiles = new Dictionary<string, string>();

                foreach (var relatedItem in relatedItems)
                {


                    var relatedItemDirectoryName = Path.GetDirectoryName(relatedItem);
                    var relatedItemfileName = Path.GetFileName(relatedItem);
                    var relatedItemfileNameWithoutExtension = Path.GetFileNameWithoutExtension(relatedItem);
                    var relatedItemExtension = Path.GetExtension(relatedItem);

                    var newPath = relatedItem;
                    var newName = relatedItemfileNameWithoutExtension;

                    newName = GetNewName(relatedItemfileNameWithoutExtension);

                    newPath = Path.Combine(activeDocumentDirectoryName, newName + relatedItemExtension);

                    Console.WriteLine($"--SUBITEM: {relatedItem}");
                    Console.WriteLine($"----OLDPATH: {relatedItem}");
                    Console.WriteLine($"----NEWPATH: {newPath}");

                    try
                    {
                        if (File.Exists(newPath))
                            continue;
                    }
                    catch (Exception ex)
                    {

                        throw ex;
                    }

                    try
                    {
                        File.Copy(relatedItem, newPath);
                    }
                    catch (Exception ex)
                    {

                        throw ex;
                    }

                    if (relatedItem.EndsWith(fileName))
                    {
                        replaceFiles.Add(fileName, newPath);
                    }

                }


                try
                {
                    string val = null;
                    var replacement = replaceFiles.TryGetValue(fileName, out val);
                    if (val != null)
                    {
                        item.Replace(val, allOccurrences);

                        UpdateProperties(item);

                    }

                }
                catch (Exception ex)
                {

                    throw ex;
                }
                //Console.WriteLine($"{filePath}  -  {directoryName}  -  {fileName}");
            }
            Console.ReadLine();
        }

        private static void UpdateProperties(Occurrence occurrence)
        {
            try
            {
                var doc = (SolidEdgeDocument)occurrence.OccurrenceDocument;
                SolidEdgeFramework.PropertySets properties = (SolidEdgeFramework.PropertySets)(doc.Properties);

                var summaryInfo = doc.GetSummaryInfo();
                summaryInfo.DocumentNumber = "0";
                summaryInfo.RevisionNumber = "0";
                summaryInfo.ProjectName = "0";
                doc.Status = DocumentStatus.igStatusAvailable;


                properties.ChangeCustomProperty("x", "Þ");
                //properties.ChangeCustomProperty("FechaPlano", DateTime.Now.ToString("dd/MM/yyyy"));
                properties.Save();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public static void ChangeCustomProperty(this PropertySets propertySets, string propertyName, string propertyValue)
        {
            Properties properties = null;
            Property property1 = null;
            try
            {
                properties = propertySets.Item((object)"Custom");
                property1 = properties.Item((object)propertyName);
                //if (propertyName == "FechaPlano")
                //{
                //    property1.Delete();
                //    return properties.Add((object)propertyName, (object)propertyValue);
                //}
                //Property property2 = property1;
                //object obj = (object)propertyValue;
                // ISSUE: explicit reference operation
                // ISSUE: variable of a reference type
                //object& local = @obj;
                //property2.Value = (object)local;

                property1.set_Value(propertyValue);

            }
            catch (Exception ex)
            {
                //property1 = properties.Add((object)propertyName, (object)propertyValue);
            }


        }
    }
}
