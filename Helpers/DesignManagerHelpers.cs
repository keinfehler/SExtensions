﻿using SolidEdgeAssembly;
using SolidEdgeCommunity.AddIn;
using SolidEdgeCommunity.Extensions;
using SolidEdgeFramework;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Path = System.IO.Path;

namespace Helpers
{
    public static class DesignManagerHelpers
    {
        public static int GetRevisionNumber(string relatedItemfileNameWithoutExtension)
        {
            var intValue = 0;
            var splittedFileName = relatedItemfileNameWithoutExtension.Split('-').ToList();
            if (splittedFileName.Count > 1)
            {
                var lastItem = splittedFileName.LastOrDefault();
         
                if (int.TryParse(lastItem, out intValue))
                    intValue++;
            }
            return intValue;
        }
        public static string GetNewName(string relatedItemfileNameWithoutExtension, string revNumber = "00")
        {
            var splittedFileName = relatedItemfileNameWithoutExtension.Split('-').ToList();

            if (splittedFileName.Count > 1)
            {
                var lastItem = splittedFileName.LastOrDefault();

                splittedFileName.RemoveAt(splittedFileName.Count - 1);

            }
            splittedFileName.Add($"-{revNumber}");

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
        public static SolidEdgeDocument ActiveDocument => SolidEdgeAddIn.Instance.Application.ActiveDocument as SolidEdgeDocument;
        public static void OpenDraftFiles()
        {
            if (ActiveDocument != null)
            {
                List<Occurrence> occs = new List<Occurrence>();

                var selectedItems = ActiveDocument.SelectSet.OfType<Occurrence>().ToList();
                occs.AddRange(selectedItems);

                foreach (var item in selectedItems)
                {
                    var subItems = item?.SubOccurrences?.OfType<Occurrence>();
                    if (subItems?.Any() ?? false)
                    {
                        occs.AddRange(subItems);
                    }
                }

                foreach (var selectedItem in occs)
                {

                    var draftPath = Path.ChangeExtension(selectedItem.OccurrenceFileName, ".dft");
                    //SolidEdgeAddIn.Instance.Application.ope
                    try
                    {
                        Process.Start(draftPath);
                    }
                    catch (Exception ex)
                    {

                        continue;
                    }

                }

            }

           
        }
        public static void OtroMetodo()
        {

        }

        public static void CreateCopy(string newReName)
        {
            SolidEdgeDocument document = ActiveDocument;
            
            var activeDocumentFullName = document.FullName;
            var activeDocumentDirectoryName = Path.GetDirectoryName(activeDocumentFullName);
            Console.WriteLine($"ACTIVEDOCUMENT:{activeDocumentFullName}");
            Console.WriteLine($"ACTIVEDOCUMENT_DIRECTORY:{activeDocumentDirectoryName}");
            Console.WriteLine($"----------------");
            Console.WriteLine($"----------------");
            Console.WriteLine($"----------------");
            Console.WriteLine($"----------------");

            var selectedOccs = Helpers.DesignManagerHelpers.GetSelectedOccurrences(document, true);
            if (selectedOccs.Count > 1)
            {
                return;
            }


            foreach (var item in selectedOccs)
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
                var filesToDelete = new List<string>();

                foreach (var relatedItem in relatedItems)
                {


                    var relatedItemDirectoryName = Path.GetDirectoryName(relatedItem);
                    var relatedItemfileName = Path.GetFileName(relatedItem);
                    var relatedItemfileNameWithoutExtension = Path.GetFileNameWithoutExtension(relatedItem);
                    var relatedItemExtension = Path.GetExtension(relatedItem);

                    var newPath = relatedItem;
                    var newName = relatedItemfileNameWithoutExtension;

                    filesToDelete.Add(relatedItem);

                    newName = newReName + "-00";//GetNewName(relatedItemfileNameWithoutExtension);

                    newPath = Path.Combine(relatedItemDirectoryName, newName + relatedItemExtension);

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

                        
                        double[] matrix = item.GetMatrix();
                        Occurrences occurrences = ((AssemblyDocument)SolidEdgeAddIn.Instance.Application.ActiveDocument).Occurrences;
                        
                        Occurrence occurrence = occurrences.AddWithMatrix(val, matrix);

                        UpdateProperties(occurrence);

                    }

                }
                catch (Exception ex)
                {

                    throw ex;
                }
            }
            Console.ReadLine();
        }
        public static void Rename(string newReName, bool deleteAfterRename = true)
        {
            SolidEdgeDocument document = ActiveDocument;
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
                var filesToDelete = new List<string>();

                foreach (var relatedItem in relatedItems)
                {


                    var relatedItemDirectoryName = Path.GetDirectoryName(relatedItem);
                    var relatedItemfileName = Path.GetFileName(relatedItem);
                    var relatedItemfileNameWithoutExtension = Path.GetFileNameWithoutExtension(relatedItem);
                    var relatedItemExtension = Path.GetExtension(relatedItem);

                    var newPath = relatedItem;
                    var newName = relatedItemfileNameWithoutExtension;

                    filesToDelete.Add(relatedItem);

                    newName = newReName + "-00";//GetNewName(relatedItemfileNameWithoutExtension);

                    newPath = Path.Combine(relatedItemDirectoryName, newName + relatedItemExtension);

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
                        item.Replace(val, true);

                        UpdateProperties(item);

                    }

                }
                catch (Exception ex)
                {

                    throw ex;
                }


                try
                {
                    foreach (var fileToDelete in filesToDelete)
                    {
                        File.Delete(fileToDelete);
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
        public static void ReplaceAndCopy(SolidEdgeDocument document, bool allOccurrences, bool revision = false)
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
                int revisionNumber = 0;
                var revisionNumberString = "00";
                if (revision)
                {
                    revisionNumber = GetRevisionNumber(fileNameWithoutExtension);
                    revisionNumberString = revisionNumber.ToString("00");
                }

                

                foreach (var relatedItem in relatedItems)
                {


                    var relatedItemDirectoryName = Path.GetDirectoryName(relatedItem);
                    var relatedItemfileName = Path.GetFileName(relatedItem);
                    var relatedItemfileNameWithoutExtension = Path.GetFileNameWithoutExtension(relatedItem);
                    var relatedItemExtension = Path.GetExtension(relatedItem);

                    var newPath = relatedItem;
                    var newName = relatedItemfileNameWithoutExtension;

                    newName = GetNewName(relatedItemfileNameWithoutExtension, revisionNumberString);

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

                        UpdateProperties(item, revision, revisionNumber);
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

        private static void UpdateProperties(Occurrence occurrence, bool revision = false, int revisionNumber = 0)
        {
            try
            {
                var doc = (SolidEdgeDocument)occurrence.OccurrenceDocument;
                SolidEdgeFramework.PropertySets properties = (SolidEdgeFramework.PropertySets)(doc.Properties);

                var summaryInfo = doc.GetSummaryInfo();
                if (!revision)
                {
                    summaryInfo.DocumentNumber = "Þ";
                    summaryInfo.ProjectName = "Þ";
                }


                summaryInfo.RevisionNumber = revisionNumber.ToString();

                if (!revision)
                {
                    Console.WriteLine($"---OccurrenceFileName: {occurrence.OccurrenceFileName}");
                    var occurrenceName = Path.GetFileNameWithoutExtension(occurrence.OccurrenceFileName);
                    Console.WriteLine($"---Name: {occurrenceName}");
                    var l = occurrenceName.Length - 3;
                    var title = occurrenceName.Substring(0, l);
                    Console.WriteLine($"---Title: {title}");
                    summaryInfo.Title = title;
                    doc.Status = DocumentStatus.igStatusAvailable;
                    properties.ChangeCustomProperty("X", "®");
                }
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

                

            }
            catch (Exception ex)
            {
                property1 = properties.Add((object)propertyName, "");
            }
            property1.set_Value(propertyValue);

        }
    }
}
