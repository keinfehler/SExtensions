using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;
using SolidEdgeAssembly;
using SolidEdgeCommunity.AddIn;
using SolidEdgeFramework;
using System;
using System.Runtime.CompilerServices;
using System.Windows.Forms;

namespace SExtensions
{
    public class IncludeExcludeHelper
    {
        public static string simbolo = "■■■ ";
        public static string excludeTemplateFilePath = ConfigurationHelper.GetConfigurationValue("ExcludeTemplateFilePath");

        public static void IncludeExclude(bool include)
        {
            var app = SolidEdgeAddIn.Instance.Application;
            //var assemblyDocument = app.GetActiveDocument<AssemblyDocument>();

            foreach (object activeSelect in app.ActiveSelectSet)
            {
                var o = RuntimeHelpers.GetObjectValue(RuntimeHelpers.GetObjectValue(activeSelect));
                if (include)
                    IncluirOcurrencia(o);
                else
                    ExcluirOcurrencia(o);
            }
                

        }

        public static void ExcluirOcurrencia(object obj_seleccionado)
        {
            try
            {
                Occurrence Expression = (Occurrence)null;
                string FileName = excludeTemplateFilePath;
                if (!Information.IsNothing((object)(obj_seleccionado as Occurrence)))
                {
                    Expression = (Occurrence)obj_seleccionado;
                    Expression.TopLevelDocument.ImportStyles(FileName, (object)true);
                    Expression.Style = "__NO INCLUIR";
                }
                else if (!Information.IsNothing((object)(obj_seleccionado as Reference)))
                {
                    Reference reference2 = (Reference)obj_seleccionado;
                    ((Occurrence)reference2.Parent).TopLevelDocument.ImportStyles(FileName, (object)true);
                    if (!Information.IsNothing((object)(reference2.Object as Occurrence)))
                    {
                        Expression = (Occurrence)reference2.Object;
                        Expression.TopLevelDocument.ImportStyles(FileName, (object)true);
                        reference2.Style = "__NO INCLUIR";
                        Expression.Style = "__NO INCLUIR";
                    }
                }
                if (!Information.IsNothing((object)Expression))
                {
                    Expression.IncludeInBom = false;
                    Expression.DisplayInDrawings = false;
                    Expression.IncludeInPhysicalProperties = false;
                    if (!Expression.Name.Contains(simbolo))
                    {
                        string str = Expression.OccurrenceFileName.Substring(checked(Expression.OccurrenceFileName.LastIndexOf("\\") + 1));
                        int int32 = Convert.ToInt32(Expression.Name.Substring(checked(Expression.Name.LastIndexOf(":") + 1)));
                        bool flag = false;
                        while (!flag)
                        {
                            try
                            {
                                Expression.Name = simbolo + str + ":" + Conversions.ToString(int32);
                                flag = true;
                            }
                            catch (Exception ex)
                            {
                                ProjectData.SetProjectError(ex);
                                checked { ++int32; }
                                ProjectData.ClearProjectError();
                            }
                        }
                    }
                }
                //reference1 = (Reference)null;
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
           
        }

        public static void IncluirOcurrencia(object obj_seleccionado)
        {
            try
            {
                Occurrence Expression = (Occurrence)null;
                if (!Information.IsNothing((object)(obj_seleccionado as Occurrence)))
                {
                    Expression = (Occurrence)obj_seleccionado;
                    Expression.Style = "";
                }
                else if (!Information.IsNothing((object)(obj_seleccionado as Reference)))
                {
                    Reference reference2 = (Reference)obj_seleccionado;
                    if (!Information.IsNothing((object)(reference2.Object as Occurrence)))
                    {
                        Expression = (Occurrence)reference2.Object;
                        reference2.Style = "";
                        Expression.Style = "";
                    }
                }
                if (!Information.IsNothing((object)Expression))
                {
                    Expression.IncludeInBom = true;
                    Expression.DisplayInDrawings = true;
                    Expression.IncludeInPhysicalProperties = true;
                    Expression.ResetName();
                }
                //reference1 = (Reference)null;
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
   
            
        }
    }
}
