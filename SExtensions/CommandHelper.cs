using SolidEdgeAssembly;
using SolidEdgeCommunity.AddIn;
using SolidEdgeFramework;
using System;
using System.Linq;
using System.Runtime.CompilerServices;

namespace SExtensions
{
    internal class CommandHelper
    {
        public static void SetPropertyValue(Occurrence occ, string propiedad, string valor)
        {
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
        /// <summary>
        /// Cambia el valor de una propiedad en las ocurrencias seleccionadas
        /// </summary>
        /// <param name="propiedad"></param>
        /// <param name="valor"></param>
        internal static void SetOcurrenceProperty(string propiedad, string valor)
        {
            var app = SolidEdgeAddIn.Instance.Application;
            foreach (SelectSet activeSelect in app.ActiveSelectSet)
            {
                var o = RuntimeHelpers.GetObjectValue(RuntimeHelpers.GetObjectValue(activeSelect));

                Occurrence occ = o as Occurrence;
                SetPropertyValue(occ, propiedad, valor);

                
               
            }
        }
        
    }
}