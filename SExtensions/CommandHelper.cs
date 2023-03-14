using Microsoft.VisualBasic;
using SolidEdgeAssembly;
using SolidEdgeCommunity.AddIn;
using SolidEdgeFramework;
using System;
using System.Linq;
using System.Runtime.CompilerServices;

namespace SExtensions
{
    public static class CommandHelper
    {
        private static void SetPropertyValue(object o, string propiedad, string valor)
        {

            Occurrence occ = null;
            if (!Information.IsNothing((object)(o as Occurrence)))
            {

            }

                if (o is Occurrence)
            {
                occ = o as Occurrence;
            }else  
            if (o is Reference)
            {
                Reference reference = o as Reference;
                if (reference != null && reference.Object is Occurrence)
                {
                    occ = reference.Object as Occurrence;
                }
            }

            if (occ != null)
            {
                var document = occ.OccurrenceDocument as SolidEdgeDocument;

                if (document != null)
                {
                    var properties = document.Properties as PropertySets;
                    try
                    {
                        var p = document.Properties;
                        var propertySets = (SolidEdgeFramework.PropertySets)p;

                        var customProperties = propertySets.Item(4);

                        var items = customProperties.OfType<SolidEdgeFramework.Property>();

                        var property = items.FirstOrDefault(t => t.Name == propiedad);

                        if (property != null)
                        {
                            property.set_Value(valor);

                            //return property.get_Value()?.ToString() ?? string.Empty;
                        }
                        else
                        {
                            customProperties.Add(propiedad, valor);
                            customProperties.Save();

                        }

                    }
                    catch (Exception)
                    {
                        //return string.Empty;
                        //MessageBox.Show(ex.Message);
                    }
                }
            }
        }
        private static object GetObjectType(object item)
        {
            return item.GetType().InvokeMember("Type", System.Reflection.BindingFlags.GetProperty, null, item, null); 
        }

        /// <summary>
        /// Cambia el valor de una propiedad en las ocurrencias seleccionadas
        /// </summary>
        /// <param name="propiedad"></param>
        /// <param name="valor"></param>
        public static void SetOcurrenceProperty(this SolidEdgeFramework.Application app, string propiedad, string valor)
        {
            foreach (object item in app.ActiveSelectSet)
            {
                //string o = string.Empty;
                SolidEdgeAssembly.Occurrence occ = null;

                try
                {
                    var objectType = GetObjectType(item);
                    var t = ((SolidEdgeConstants.ObjectType)objectType);
          

                    if (t == SolidEdgeConstants.ObjectType.igReference)
                    {
                        var r = (Reference)item;
                        occ = (Occurrence)(r.Object);
                        //r.Style = "STYLE_ROJO";
                    }
                    else
                    {
                        occ = (SolidEdgeAssembly.Occurrence)item;
                    }
                    
                    //if (occ != null)
                    //{
                    //    //occ.Style = "STYLE_ROJO";
                    //    //o = occ.Name;
                    //}
                    //else
                    //{
                    //    //o = "EMPTY";
                    //}

                    // Using the .NET runtime type, get the Type property of the object.
          

                    //if (t == SolidEdgeConstants.ObjectType.seOccurrences)
                    //{
                    //    var occurrences = item as Occurrences;
                    //    foreach (var occ in occurrences)
                    //    {
                    //        var occType = GetObjectType(occ);
                    //        Console.WriteLine((SolidEdgeConstants.ObjectType)occType);
                    //    }
                    //}

                    //o = t.ToString();
                    
                }
                catch (Exception e)
                {
                    //o = e.Message;
                    

                }finally
                {
                    //Console.WriteLine(o);
                }


                //var o = RuntimeHelpers.GetObjectValue(RuntimeHelpers.GetObjectValue(activeSelect));

                if (occ != null)
                {
                    SetPropertyValue(occ, propiedad, valor);
                }
                
            }

          
        }


    }
}