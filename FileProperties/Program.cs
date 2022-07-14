﻿


using SolidEdgeCommunity.Extensions; // https://github.com/SolidEdgeCommunity/SolidEdge.Community/wiki/Using-Extension-Methods
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReportFileProperties
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            SolidEdgeFramework.Application application = null;

            try
            {
                // Register with OLE to handle concurrency issues on the current thread.
                SolidEdgeCommunity.OleMessageFilter.Register();

                // Connect to or start Solid Edge.
                application = SolidEdgeCommunity.SolidEdgeUtils.Connect(true, true);

                // Get a reference to the active assembly document.
                var document = application.GetActiveDocument<SolidEdgeAssembly.AssemblyDocument>(false);

                if (document != null)
                {
                    var name = document.Name;

                    var p = document.Properties;
                    var propertySets = (SolidEdgeFramework.PropertySets)p;

                    var customProperties = propertySets.Item(4);

                    var items = customProperties.OfType<SolidEdgeFramework.Property>();
                    var modelo = items.FirstOrDefault(o => o.Name == "MODELO");


                    foreach (var pp in propertySets)
                    {

                        var properties = pp as SolidEdgeFramework.Properties;
                        if (properties != null)
                        {

                            Console.WriteLine("PropertSet '{0}'.", properties.Name);
                            

                            foreach (var property in properties.OfType<SolidEdgeFramework.Property>())
                            {
                                System.Runtime.InteropServices.VarEnum nativePropertyType = System.Runtime.InteropServices.VarEnum.VT_EMPTY;
                                Type runtimePropertyType = null;

                                object value = null;

                                nativePropertyType = (System.Runtime.InteropServices.VarEnum)property.Type;

                                // Accessing Value property may throw an exception...
                                try
                                {
                                    value = property.get_Value();
                                }
                                catch (System.Exception ex)
                                {
                                    value = ex.Message;
                                }

                                if (value != null)
                                {
                                    runtimePropertyType = value.GetType();
                                }

                                Console.WriteLine("\t{0} = '{1}' ({2} | {3}).", property.Name, value, nativePropertyType, runtimePropertyType);
                            }
                        }
                        else
                        {

                        }


                        Console.WriteLine();

                        
                    }

                    Console.ReadLine();
                }
                else
                {
                    throw new System.Exception("No active document.");
                }
            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                SolidEdgeCommunity.OleMessageFilter.Unregister();
            }
        }
    }
}